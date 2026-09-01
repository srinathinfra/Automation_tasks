"""
jira_worklogs.py
----------------
Pulls YOUR OWN worklogs out of Jira for a date range.

Works with both Jira Cloud (email + API token) and Jira Server / Data Center
(Personal Access Token). Nothing here writes to Jira - it is read-only.

How it finds your work:
  1. Runs a JQL search for issues you logged time against in the window.
  2. Pulls the full worklog list off each of those issues.
  3. Keeps only the worklogs authored by you, inside the date window.

Step 3 matters: the JQL `worklogDate` clause filters ISSUES, not individual
worklog entries, so an issue can come back holding other people's time or time
from outside your window. We filter client-side.
"""

from __future__ import annotations

import datetime as dt
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional

import requests
from requests.auth import HTTPBasicAuth


class JiraError(RuntimeError):
    """Raised for any non-recoverable problem talking to Jira."""


@dataclass
class Worklog:
    issue_key: str
    issue_summary: str
    date: dt.date
    hours: float
    comment: str
    worklog_id: str
    started_raw: str
    author: str = ""

    @property
    def seconds(self) -> int:
        return int(round(self.hours * 3600))


# --------------------------------------------------------------------------
# Atlassian Document Format helpers
# --------------------------------------------------------------------------
def adf_to_text(node: Any) -> str:
    """Flatten a Jira Cloud (API v3) rich-text comment into plain text.

    API v2 returns comments as plain strings, so this passes those straight
    through untouched.
    """
    if node is None:
        return ""
    if isinstance(node, str):
        return node
    if isinstance(node, list):
        return " ".join(p for p in (adf_to_text(n) for n in node) if p)
    if isinstance(node, dict):
        if node.get("type") == "text":
            return node.get("text", "")
        if node.get("type") == "hardBreak":
            return " "
        if node.get("type") == "mention":
            return node.get("attrs", {}).get("text", "")
        return adf_to_text(node.get("content", []))
    return str(node)


def _parse_started(started: str) -> dt.date:
    """Jira returns e.g. 2026-08-03T09:00:00.000-0500.

    We take the date portion verbatim rather than converting timezones. The
    timestamp is already expressed in the instance's timezone, which is the
    same calendar day your Jira board displays - converting to UTC would shove
    early-morning and late-evening entries onto the wrong day.
    """
    return dt.date.fromisoformat(started[:10])


# --------------------------------------------------------------------------
# Client
# --------------------------------------------------------------------------
class JiraClient:
    def __init__(
        self,
        base_url: str,
        deployment: str = "cloud",
        email: str = "",
        token: str = "",
        verify_ssl: bool = True,
        timeout: int = 60,
    ) -> None:
        self.base_url = base_url.rstrip("/")
        self.deployment = (deployment or "cloud").strip().lower()
        if self.deployment not in ("cloud", "server"):
            raise JiraError("deployment must be either 'cloud' or 'server'")
        if not self.base_url.startswith("http"):
            raise JiraError(f"base_url looks wrong: {self.base_url!r}")
        if not token:
            raise JiraError(
                "No API token supplied. Put it in config.ini under [jira] "
                "api_token, or set the JIRA_API_TOKEN environment variable."
            )

        # Cloud runs API v3, Server/DC tops out at v2.
        self.api = "3" if self.deployment == "cloud" else "2"
        self.timeout = timeout

        self.session = requests.Session()
        self.session.verify = verify_ssl
        self.session.headers.update(
            {"Accept": "application/json", "Content-Type": "application/json"}
        )
        if self.deployment == "cloud":
            if not email:
                raise JiraError("Jira Cloud needs your email address in config.ini")
            self.session.auth = HTTPBasicAuth(email, token)
        else:
            self.session.headers["Authorization"] = f"Bearer {token}"

        self._me: Optional[Dict[str, Any]] = None

    # -- plumbing ----------------------------------------------------------
    def _url(self, path: str) -> str:
        return f"{self.base_url}/rest/api/{self.api}{path}"

    def _request(self, method: str, path: str, **kwargs) -> requests.Response:
        url = self._url(path)
        try:
            resp = self.session.request(method, url, timeout=self.timeout, **kwargs)
        except requests.exceptions.SSLError as exc:
            raise JiraError(
                f"SSL error reaching {url}. If you are behind a corporate proxy, "
                f"set verify_ssl = false in config.ini.\n{exc}"
            ) from exc
        except requests.exceptions.RequestException as exc:
            raise JiraError(f"Could not reach {url}: {exc}") from exc

        if resp.status_code == 401:
            raise JiraError(
                "401 Unauthorized. Check your email/token. On Jira Cloud the "
                "token must be an API token from id.atlassian.com, not your password."
            )
        if resp.status_code == 403:
            raise JiraError(
                "403 Forbidden. Your account is authenticated but not allowed to "
                "read this. CAPTCHA lockout after failed logins can also cause this - "
                "log into Jira in a browser once and retry."
            )
        return resp

    def _get(self, path: str, params: Optional[dict] = None) -> dict:
        resp = self._request("GET", path, params=params or {})
        if not resp.ok:
            raise JiraError(f"GET {path} -> {resp.status_code}: {resp.text[:400]}")
        return resp.json()

    # -- identity ----------------------------------------------------------
    def myself(self) -> Dict[str, Any]:
        if self._me is None:
            self._me = self._get("/myself")
        return self._me

    def my_identity(self) -> str:
        """A human-readable label for whoever the token belongs to."""
        me = self.myself()
        return me.get("displayName") or me.get("name") or me.get("emailAddress") or "?"

    def _is_me(self, author: Dict[str, Any]) -> bool:
        me = self.myself()
        if not author:
            return False
        if self.deployment == "cloud":
            return author.get("accountId") == me.get("accountId")
        # Server/DC: accountId is absent, match on username or key
        return (
            (author.get("name") and author.get("name") == me.get("name"))
            or (author.get("key") and author.get("key") == me.get("key"))
        )

    # -- search ------------------------------------------------------------
    def search_issue_keys(self, jql: str) -> Dict[str, str]:
        """Return {issue_key: summary} for every issue matching the JQL."""
        if self.deployment == "cloud":
            try:
                return self._search_cloud_token_paged(jql)
            except JiraError as exc:
                # Older Cloud instances (and every Server build) only have the
                # legacy startAt-paged endpoint. Fall back rather than die.
                if "404" not in str(exc) and "410" not in str(exc):
                    raise
        return self._search_legacy_paged(jql)

    def _search_cloud_token_paged(self, jql: str) -> Dict[str, str]:
        """Current Cloud endpoint: /search/jql, cursor paging via nextPageToken."""
        out: Dict[str, str] = {}
        token: Optional[str] = None
        while True:
            body: Dict[str, Any] = {
                "jql": jql,
                "fields": ["summary"],
                "maxResults": 100,
            }
            if token:
                body["nextPageToken"] = token
            resp = self._request("POST", "/search/jql", json=body)
            if not resp.ok:
                raise JiraError(
                    f"POST /search/jql -> {resp.status_code}: {resp.text[:400]}"
                )
            data = resp.json()
            for issue in data.get("issues", []):
                out[issue["key"]] = (issue.get("fields") or {}).get("summary", "")
            token = data.get("nextPageToken")
            if not token or data.get("isLast"):
                break
        return out

    def _search_legacy_paged(self, jql: str) -> Dict[str, str]:
        """Legacy endpoint: /search, offset paging via startAt."""
        out: Dict[str, str] = {}
        start_at = 0
        while True:
            body = {
                "jql": jql,
                "fields": ["summary"],
                "startAt": start_at,
                "maxResults": 100,
            }
            resp = self._request("POST", "/search", json=body)
            if not resp.ok:
                raise JiraError(f"POST /search -> {resp.status_code}: {resp.text[:400]}")
            data = resp.json()
            issues = data.get("issues", [])
            for issue in issues:
                out[issue["key"]] = (issue.get("fields") or {}).get("summary", "")
            start_at += len(issues)
            if not issues or start_at >= data.get("total", 0):
                break
        return out

    # -- worklogs ----------------------------------------------------------
    def worklogs_for_issue(self, issue_key: str) -> List[dict]:
        out: List[dict] = []
        start_at = 0
        while True:
            data = self._get(
                f"/issue/{issue_key}/worklog",
                {"startAt": start_at, "maxResults": 1000},
            )
            batch = data.get("worklogs", [])
            out.extend(batch)
            start_at += len(batch)
            if not batch or start_at >= data.get("total", 0):
                break
        return out

    def fetch_my_worklogs(
        self, date_from: dt.date, date_to: dt.date, verbose: bool = True
    ) -> List[Worklog]:
        jql = (
            f'worklogAuthor = currentUser() '
            f'AND worklogDate >= "{date_from.isoformat()}" '
            f'AND worklogDate <= "{date_to.isoformat()}" '
            f'ORDER BY key ASC'
        )
        if verbose:
            print(f"  JQL: {jql}")

        issues = self.search_issue_keys(jql)
        if verbose:
            print(f"  {len(issues)} issue(s) carry your time in this window.")

        results: List[Worklog] = []
        for n, (key, summary) in enumerate(sorted(issues.items()), start=1):
            if verbose:
                print(f"  [{n}/{len(issues)}] reading worklogs on {key}", end="\r")
            for wl in self.worklogs_for_issue(key):
                if not self._is_me(wl.get("author", {})):
                    continue
                started = wl.get("started")
                if not started:
                    continue
                day = _parse_started(started)
                if not (date_from <= day <= date_to):
                    continue
                results.append(
                    Worklog(
                        issue_key=key,
                        issue_summary=summary,
                        date=day,
                        hours=round(wl.get("timeSpentSeconds", 0) / 3600.0, 4),
                        comment=adf_to_text(wl.get("comment")).strip(),
                        worklog_id=str(wl.get("id", "")),
                        started_raw=started,
                        author=(wl.get("author") or {}).get("displayName", ""),
                    )
                )
        if verbose:
            print(" " * 70, end="\r")
        results.sort(key=lambda w: (w.date, w.issue_key))
        return results


def hours_by_day(worklogs: List[Worklog]) -> Dict[dt.date, float]:
    totals: Dict[dt.date, float] = {}
    for wl in worklogs:
        totals[wl.date] = totals.get(wl.date, 0.0) + wl.hours
    return totals
