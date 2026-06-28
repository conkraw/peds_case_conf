"""Optional GitHub archive storage for the Pediatric Case Conference Builder."""

from __future__ import annotations

import base64
import json
import re
import uuid
from dataclasses import dataclass
from datetime import date, datetime, timezone
from typing import Any, Dict, List, Optional
from urllib.parse import quote

import requests
import streamlit as st

from case_schema import archive_summary


class GitHubArchiveError(RuntimeError):
    """Raised when GitHub archive actions fail."""


@dataclass
class GitHubResult:
    path: str
    html_url: str
    commit_sha: str


def slugify(value: str) -> str:
    value = str(value or "").strip().lower()
    value = re.sub(r"[^a-z0-9]+", "-", value)
    value = re.sub(r"-+", "-", value).strip("-")
    return value or "untitled"


def normalize_archive_id(archive_id: str | None) -> str:
    cleaned = re.sub(r"[^a-zA-Z0-9]+", "", str(archive_id or "")).lower()
    return cleaned[:24]


def generate_archive_id() -> str:
    return uuid.uuid4().hex[:12]


def _read_github_config() -> Dict[str, str]:
    try:
        raw = st.secrets.get("github", {})
    except Exception:
        raw = {}
    return {
        "token": str(raw.get("token", "")).strip(),
        "repo": str(raw.get("repo", "")).strip(),
        "branch": str(raw.get("branch", "main")).strip() or "main",
        "base_path": str(raw.get("base_path", "case-drafts")).strip().strip("/") or "case-drafts",
    }


def _github_headers(token: str) -> Dict[str, str]:
    return {
        "Authorization": f"Bearer {token}",
        "Accept": "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
    }


def github_backup_is_configured() -> bool:
    cfg = _read_github_config()
    return bool(cfg["token"] and cfg["repo"] and "/" in cfg["repo"])


def github_config_status_message() -> str:
    cfg = _read_github_config()
    missing = []
    if not cfg["token"]:
        missing.append("github.token")
    if not cfg["repo"] or "/" not in cfg["repo"]:
        missing.append("github.repo")
    if missing:
        return "Missing Streamlit secrets: " + ", ".join(missing)
    return "Case archive storage is configured."


def build_draft_filename(presenter_name: str, case_title: str, archive_id: str | None = None) -> str:
    today = date.today().isoformat()
    presenter_slug = slugify(presenter_name) or "unknown-presenter"
    case_slug = slugify(case_title) or "untitled-case"
    suffix = normalize_archive_id(archive_id)
    if suffix:
        return f"{today}_{presenter_slug}_{case_slug}_{suffix}.json"
    return f"{today}_{presenter_slug}_{case_slug}.json"


def make_draft_payload(
    deck: Dict[str, Any],
    app_version: str,
    archive_id: str | None = None,
    archive_path: str | None = None,
) -> Dict[str, Any]:
    archive_id = normalize_archive_id(archive_id) or generate_archive_id()
    summary = archive_summary(deck)
    return {
        "archive_id": archive_id,
        "archive_path": str(archive_path or "").strip().lstrip("/"),
        "presenter": summary.get("presenter", ""),
        "case": summary.get("case", ""),
        "learning_point": summary.get("learning_point", ""),
        "saved_date": date.today().isoformat(),
        "saved_at_utc": datetime.now(timezone.utc).replace(microsecond=0).isoformat(),
        "app_version": app_version,
        "deck": deck,
    }


def save_draft_to_github(
    deck: Dict[str, Any],
    app_version: str,
    archive_id: str | None = None,
    existing_path: str | None = None,
) -> GitHubResult:
    cfg = _read_github_config()
    if not github_backup_is_configured():
        raise GitHubArchiveError(github_config_status_message())

    summary = archive_summary(deck)
    archive_id = normalize_archive_id(archive_id) or generate_archive_id()
    clean_existing_path = str(existing_path or "").strip().lstrip("/")
    if clean_existing_path:
        path = clean_existing_path
    else:
        filename = build_draft_filename(summary.get("presenter", ""), summary.get("case", ""), archive_id)
        path = f"{cfg['base_path']}/{filename}"

    payload = make_draft_payload(deck, app_version=app_version, archive_id=archive_id, archive_path=path)
    json_text = json.dumps(payload, indent=2, ensure_ascii=False)
    encoded_content = base64.b64encode(json_text.encode("utf-8")).decode("utf-8")

    api_path = quote(path, safe="/")
    api_url = f"https://api.github.com/repos/{cfg['repo']}/contents/{api_path}"
    headers = _github_headers(cfg["token"])

    existing_sha: Optional[str] = None
    existing = requests.get(api_url, headers=headers, params={"ref": cfg["branch"]}, timeout=30)
    if existing.status_code == 200:
        existing_sha = existing.json().get("sha")
    elif existing.status_code not in (404,):
        raise GitHubArchiveError(f"Could not check existing archive file: {existing.status_code} {existing.text[:300]}")

    body: Dict[str, Any] = {
        "message": f"Save case conference draft: {summary.get('case', 'Untitled case')}",
        "content": encoded_content,
        "branch": cfg["branch"],
    }
    if existing_sha:
        body["sha"] = existing_sha

    response = requests.put(api_url, headers=headers, data=json.dumps(body), timeout=30)
    if response.status_code not in (200, 201):
        raise GitHubArchiveError(f"Could not save archive file: {response.status_code} {response.text[:500]}")
    result = response.json()
    content = result.get("content", {})
    commit = result.get("commit", {})
    return GitHubResult(path=path, html_url=content.get("html_url", ""), commit_sha=commit.get("sha", ""))


def _decode_content_payload(item: Dict[str, Any]) -> Dict[str, Any]:
    content = item.get("content", "")
    encoding = item.get("encoding", "")
    if encoding == "base64" and content:
        text = base64.b64decode(content).decode("utf-8")
        return json.loads(text)
    raise GitHubArchiveError("GitHub returned an unexpected file encoding.")


def load_draft_from_github(path: str) -> Dict[str, Any]:
    cfg = _read_github_config()
    if not github_backup_is_configured():
        raise GitHubArchiveError(github_config_status_message())
    clean_path = str(path or "").strip().lstrip("/")
    if not clean_path:
        raise GitHubArchiveError("No archive path was provided.")
    api_path = quote(clean_path, safe="/")
    api_url = f"https://api.github.com/repos/{cfg['repo']}/contents/{api_path}"
    response = requests.get(api_url, headers=_github_headers(cfg["token"]), params={"ref": cfg["branch"]}, timeout=30)
    if response.status_code != 200:
        raise GitHubArchiveError(f"Could not load archive file: {response.status_code} {response.text[:300]}")
    return _decode_content_payload(response.json())


def list_drafts_from_github(search_text: str = "") -> List[Dict[str, str]]:
    cfg = _read_github_config()
    if not github_backup_is_configured():
        raise GitHubArchiveError(github_config_status_message())
    list_url = f"https://api.github.com/repos/{cfg['repo']}/contents/{quote(cfg['base_path'], safe='/')}"
    response = requests.get(list_url, headers=_github_headers(cfg["token"]), params={"ref": cfg["branch"]}, timeout=30)
    if response.status_code == 404:
        return []
    if response.status_code != 200:
        raise GitHubArchiveError(f"Could not list archive files: {response.status_code} {response.text[:300]}")

    query = str(search_text or "").strip().lower()
    rows: List[Dict[str, str]] = []
    for item in response.json():
        if item.get("type") != "file" or not str(item.get("name", "")).endswith(".json"):
            continue
        path = item.get("path", "")
        try:
            payload = load_draft_from_github(path)
        except Exception:
            continue
        row = {
            "saved_date": str(payload.get("saved_date", "")),
            "presenter": str(payload.get("presenter", "")),
            "case": str(payload.get("case", "")),
            "learning_point": str(payload.get("learning_point", "")),
            "path": str(payload.get("archive_path") or path),
            "archive_id": str(payload.get("archive_id", "")),
            "saved_at_utc": str(payload.get("saved_at_utc", "")),
        }
        haystack = " ".join(row.values()).lower()
        if query and query not in haystack:
            continue
        rows.append(row)

    rows.sort(key=lambda row: row.get("saved_at_utc", ""), reverse=True)
    return rows
