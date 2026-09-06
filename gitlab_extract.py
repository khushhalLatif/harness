"""
SCRIPT 2 of 3 — GitLab extraction into the shared master Excel file.

Reads harness_gitlab_master.xlsx (as populated by Script 1). For every row,
runs TWO INDEPENDENT matching methods side by side and writes each one's
results to its own suffixed columns — they are never merged or picked
between automatically, since they can legitimately land on different repos:

  METHOD 1 — name/path matching (columns ...1)
    Matches the service to a GitLab project by searching on the Harness
    service name and the last path segment of Artifact Path (e.g.
    "ol-ms-account" out of "apm/0006279/ol-ms-account"), against every
    project under GITLAB_GROUP_PATH (fetched and cached once per run —
    NOT via GitLab's search param, which doesn't reliably match a
    project's path/slug, only its display name).

  METHOD 2 — DOCKER_IMAGE_NAME content matching (columns ...2)
    Identifies the project by CONTENT instead of name: every project's
    root .gitlab-ci.yml, and any ci-job-config.yml in a child folder, is
    scanned once for a DOCKER_IMAGE_NAME value. That scan is built into a
    reverse index (image name -> project) ONE TIME, cached to disk, so
    each row after that is a fast lookup instead of a full re-scan.
    Whichever project's CI config declares DOCKER_IMAGE_NAME matching the
    artifact name IS the match for that row.

  Once EITHER method has identified a project, finding the actual
  FORTIFY_APPLICATION_NAME/VERSION values reuses the SAME file-search
  logic (code search for the variable names, falling back to common file
  paths) — Method 2 only differs in how the PROJECT is identified, not in
  how the values are subsequently located within it. Flagging this as an
  assumption: if the FORTIFY_ values must come from the exact same file
  DOCKER_IMAGE_NAME was found in (rather than wherever they're actually
  found in that repo), this needs a small adjustment — let me know.

OVERRIDES (per method, independently): if Corresponding Gitlab
component<N> or GitLab File Path<N> is already filled in — a full URL
pasted by hand, or the result of a previous run — it's trusted directly
and that method's matching/searching is skipped for that row entirely.

Corresponding Gitlab component<N> and GitLab File Path<N> are written as
FULL, clickable GitLab URLs (project page / file blob), not bare paths.

Branch values are NOT touched by this script at all — Existing_Branch
(from Script 1) is reference-only; there is no "new" branch to write.

Run Script 1 first. Script 3 (Harness update) reads this same file next,
using Appname_final / Appversion_final (filled in by hand after comparing
Method 1 / Method 2 / Existing) as the values it actually pushes.

pip install requests pyyaml openpyxl --break-system-packages
"""

import os
import sys
import json
import time
import base64
import logging
from datetime import datetime

import yaml
import requests
from openpyxl import load_workbook

from harness_config import load_config, require_gitlab

# ===========================================================================
# CONFIG
# ===========================================================================

_config = load_config()
_gitlab = require_gitlab(_config)

GITLAB_BASE_URL = _gitlab["base_url"].rstrip("/")
GITLAB_PRIVATE_TOKEN = _gitlab["private_token"]
GITLAB_BRANCH = _gitlab.get("branch") or "main"
GITLAB_GROUP_PATH = _gitlab.get("group_path")  # narrows Method 1 search AND bounds Method 2's scan

MASTER_EXCEL_FILE = "harness_gitlab_master.xlsx"
LOG_FILE = f"script2_gitlab_extract_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
DOCKER_IMAGE_INDEX_CACHE_FILE = "docker_image_index_cache.json"

# Harness variable name -> GitLab yml key name (confirmed from the client's note)
VARIABLE_MAPPING = {
    "ssc_appname": "FORTIFY_APPLICATION_VERSION",
    "ssc_appversion": "FORTIFY_APPLICATION_NAME",
}

# Tried in order only if code search comes back empty
YML_FILE_CANDIDATES = [
    "values.yaml", "values.yml",
    "config/values.yaml", "helm/values.yaml",
    ".gitlab-ci.yml",
]

SLEEP_BETWEEN_ROWS = 0.3
SLEEP_BETWEEN_PROJECT_SCANS = 0.1
MAX_RETRIES = 3
RETRY_BACKOFF_SECONDS = 2

# ===========================================================================
# Shared column schema — MUST exactly match Script 1's ALL_COLUMNS.
# ===========================================================================

ALL_COLUMNS = [
    "Service Identifier", "Service Name", "Service Url",
    "Existing_Branch",
    "Corresponding Gitlab component1", "GitLab File Path1",
    "Corresponding Gitlab component2", "GitLab File Path2",
    "New_ssc_appname1", "New_ssc_appname2", "Existing_ssc_appname", "Appname_final",
    "Artifact Path",
    "New_ssc_appversion1", "New_ssc_appversion2", "Existing_ssc_appversion", "Appversion_final",
    "Approved for Update (Y/N)", "Variables Updated",
    "Comments", "Comments from Developer",
]

OWNED_COLUMNS = [
    "Corresponding Gitlab component1", "GitLab File Path1", "New_ssc_appname1", "New_ssc_appversion1",
    "Corresponding Gitlab component2", "GitLab File Path2", "New_ssc_appname2", "New_ssc_appversion2",
]

NEW_NAME_COLUMN = {1: "New_ssc_appname1", 2: "New_ssc_appname2"}
NEW_VERSION_COLUMN = {1: "New_ssc_appversion1", 2: "New_ssc_appversion2"}
COMPONENT_COLUMN = {1: "Corresponding Gitlab component1", 2: "Corresponding Gitlab component2"}
FILEPATH_COLUMN = {1: "GitLab File Path1", 2: "GitLab File Path2"}

# ===========================================================================
# Logging
# ===========================================================================

logger = logging.getLogger("script2")
logger.setLevel(logging.DEBUG)
fh = logging.FileHandler(LOG_FILE, encoding="utf-8")
fh.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
ch = logging.StreamHandler()
ch.setLevel(logging.INFO)
ch.setFormatter(logging.Formatter("[%(levelname)s] %(message)s"))
logger.addHandler(fh)
logger.addHandler(ch)

# ===========================================================================
# GitLab API helpers
# ===========================================================================

GITLAB_HEADERS = {"PRIVATE-TOKEN": GITLAB_PRIVATE_TOKEN}


def request_with_retry(method: str, url: str, **kwargs) -> requests.Response:
    last_exc = None
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            resp = requests.request(method, url, timeout=30, **kwargs)
            if resp.status_code >= 500:
                raise requests.HTTPError(f"{resp.status_code} server error")
            return resp
        except requests.RequestException as exc:
            last_exc = exc
            logger.warning("Request failed (attempt %s/%s) %s %s: %s",
                            attempt, MAX_RETRIES, method, url, exc)
            time.sleep(RETRY_BACKOFF_SECONDS * attempt)
    raise last_exc


def get_project_by_path(path: str) -> dict | None:
    url = f"{GITLAB_BASE_URL}/api/v4/projects/{requests.utils.quote(path, safe='')}"
    resp = request_with_retry("GET", url, headers=GITLAB_HEADERS)
    if resp.status_code == 200:
        return resp.json()
    return None


def resolve_component_value(value: str) -> dict | None:
    """
    A pre-filled Corresponding Gitlab component<N> cell may be a full URL
    (what this script itself now writes) or a bare path_with_namespace
    (what a human might type by hand). Handles either.
    """
    if not value:
        return None
    bare_path = value
    if value.startswith("http"):
        prefix = GITLAB_BASE_URL + "/"
        bare_path = value[len(prefix):] if value.startswith(prefix) else value.split("://", 1)[-1].split("/", 1)[-1]
    return get_project_by_path(bare_path.rstrip("/"))


def parse_gitlab_blob_url(value: str) -> tuple[str, str, str] | None:
    """
    Parses a GitLab file URL (what this script itself writes, or what a
    human might paste) into (project_path_with_namespace, ref, file_path).
    Returns None if `value` doesn't look like a GitLab blob URL, in which
    case it's treated as a plain relative file path instead.
    """
    if not value or "://" not in value or "/-/blob/" not in value:
        return None
    path_part = value.split("://", 1)[1].split("/", 1)[-1]
    if "/-/blob/" not in path_part:
        return None
    project_path, rest = path_part.split("/-/blob/", 1)
    if "/" not in rest:
        return None
    ref, file_path = rest.split("/", 1)
    file_path = file_path.split("?", 1)[0].split("#", 1)[0]
    if not project_path or not ref or not file_path:
        return None
    return project_path.rstrip("/"), ref, file_path


def build_project_url(project: dict) -> str:
    return f"{GITLAB_BASE_URL}/{project.get('path_with_namespace', '')}"


def build_blob_url(project: dict, ref: str, file_path: str) -> str:
    return f"{GITLAB_BASE_URL}/{project.get('path_with_namespace', '')}/-/blob/{ref}/{file_path}"


def read_gitlab_file(project_id: int, file_path: str, ref: str | None = None) -> dict | None:
    encoded_path = requests.utils.quote(file_path, safe="")
    url = f"{GITLAB_BASE_URL}/api/v4/projects/{project_id}/repository/files/{encoded_path}"
    resp = request_with_retry("GET", url, headers=GITLAB_HEADERS, params={"ref": ref or GITLAB_BRANCH})
    if resp.status_code != 200:
        return None
    content_b64 = resp.json()["content"]
    content = base64.b64decode(content_b64).decode("utf-8", errors="replace")
    try:
        return yaml.safe_load(content) or {}
    except yaml.YAMLError as exc:
        logger.warning("YAML parse error in %s: %s", file_path, exc)
        return None


def code_search_yml_file(project_id: int, search_term: str) -> str | None:
    url = f"{GITLAB_BASE_URL}/api/v4/projects/{project_id}/search"
    params = {"scope": "blobs", "search": search_term, "ref": GITLAB_BRANCH}
    resp = request_with_retry("GET", url, headers=GITLAB_HEADERS, params=params)
    if resp.status_code != 200:
        return None
    results = resp.json()
    if not results:
        return None
    yml_hits = [r for r in results if r.get("path", "").lower().endswith((".yml", ".yaml"))]
    chosen = yml_hits[0] if yml_hits else results[0]
    return chosen.get("path")


def find_and_read_yaml_file(project_id: int, target_keys: list[str]) -> tuple[str, dict] | tuple[None, None]:
    for key in target_keys:
        path = code_search_yml_file(project_id, key)
        if path:
            parsed = read_gitlab_file(project_id, path)
            if parsed:
                return path, parsed
    for path in YML_FILE_CANDIDATES:
        parsed = read_gitlab_file(project_id, path)
        if parsed:
            return path, parsed
    return None, None


def extract_vars_from_yaml(parsed_yaml: dict, var_names: list[str]) -> dict:
    """Recursive lookup — checked at every level regardless of nesting depth."""
    found = {}

    def _search(node):
        if isinstance(node, dict):
            for k, v in node.items():
                if k in var_names and k not in found:
                    found[k] = v
                else:
                    _search(v)
        elif isinstance(node, list):
            for item in node:
                _search(item)

    _search(parsed_yaml)
    return found


# ---------------------------------------------------------------------------
# Group project cache — fetched once per run, matched client-side.
# GitLab's "search" query param reliably matches a project's display NAME
# but not its PATH/slug (see GitLab issue #47909 and related reports).
# ---------------------------------------------------------------------------

_group_projects_cache: list[dict] | None = None


def fetch_all_group_projects() -> list[dict]:
    projects = []
    page = 1
    encoded = requests.utils.quote(GITLAB_GROUP_PATH, safe="")
    url = f"{GITLAB_BASE_URL}/api/v4/groups/{encoded}/projects"
    while page <= 100:
        resp = request_with_retry(
            "GET", url, headers=GITLAB_HEADERS,
            params={"include_subgroups": "true", "per_page": 100, "page": page},
        )
        resp.raise_for_status()
        batch = resp.json()
        if not batch:
            break
        projects.extend(batch)
        if len(batch) < 100:
            break
        page += 1
    logger.info("Fetched %s total project(s) under group '%s' (including subgroups)",
                len(projects), GITLAB_GROUP_PATH)
    return projects


def get_group_projects_cached() -> list[dict]:
    global _group_projects_cache
    if _group_projects_cache is None:
        _group_projects_cache = fetch_all_group_projects() if GITLAB_GROUP_PATH else []
    return _group_projects_cache


def _normalize(s: str) -> str:
    return "".join(ch for ch in (s or "").lower() if ch.isalnum())


def find_matches_in_group(term: str) -> list[dict]:
    if not term:
        return []
    term_norm = _normalize(term)
    projects = get_group_projects_cached()

    exact = [p for p in projects
             if _normalize(p.get("path", "")) == term_norm or _normalize(p.get("name", "")) == term_norm]
    if exact:
        return exact

    return [p for p in projects
            if term_norm in _normalize(p.get("path", "")) or term_norm in _normalize(p.get("name", ""))]


def search_gitlab_projects_ungrouped(term: str) -> list[dict]:
    """Fallback only used when no GITLAB_GROUP_PATH is configured at all."""
    resp = request_with_retry("GET", f"{GITLAB_BASE_URL}/api/v4/projects",
                               headers=GITLAB_HEADERS, params={"search": term, "per_page": 100})
    resp.raise_for_status()
    return resp.json()


def resolve_project_method1(service_name: str, artifact_path: str):
    """Returns (project_dict_or_None, note)."""
    terms, seen = [], set()
    for term in [service_name, (artifact_path or "").rstrip("/").split("/")[-1]]:
        if term and term not in seen:
            seen.add(term)
            terms.append(term)

    pool = {}
    if GITLAB_GROUP_PATH:
        for term in terms:
            for proj in find_matches_in_group(term):
                pool[proj["id"]] = proj
    else:
        for term in terms:
            try:
                for proj in search_gitlab_projects_ungrouped(term):
                    pool[proj["id"]] = proj
            except requests.HTTPError as exc:
                logger.warning("GitLab search failed for '%s': %s", term, exc)

    if not pool:
        return None, f"Method 1: no GitLab project found for candidates: {terms}"

    lower_terms = [t.lower() for t in terms]
    exact_matches = [p for p in pool.values()
                      if p.get("name", "").lower() in lower_terms or p.get("path", "").lower() in lower_terms]
    if len(exact_matches) == 1:
        return exact_matches[0], ""
    if len(pool) == 1:
        return next(iter(pool.values())), ""

    labels = [p.get("path_with_namespace", p.get("name")) for p in pool.values()]
    return None, (f"Method 1: ambiguous match for candidates {terms} -> found {labels}; "
                  f"set 'Corresponding Gitlab component1' manually to resolve")


# ---------------------------------------------------------------------------
# Method 2 — DOCKER_IMAGE_NAME content index (built once, cached to disk)
# ---------------------------------------------------------------------------


def list_repo_yml_paths(project_id: int, filename: str) -> list[str]:
    """Every file in the repo tree (any depth) whose basename == filename."""
    matches = []
    page = 1
    while page <= 50:
        resp = request_with_retry(
            "GET", f"{GITLAB_BASE_URL}/api/v4/projects/{project_id}/repository/tree",
            headers=GITLAB_HEADERS, params={"recursive": "true", "per_page": 100, "page": page},
        )
        if resp.status_code != 200:
            break
        batch = resp.json()
        if not batch:
            break
        for item in batch:
            if item.get("type") == "blob" and item.get("path", "").split("/")[-1] == filename:
                matches.append(item["path"])
        if len(batch) < 100:
            break
        page += 1
    return matches


def find_docker_image_name(parsed_yaml) -> str | None:
    result = {}

    def _search(node):
        if isinstance(node, dict):
            for k, v in node.items():
                if k == "DOCKER_IMAGE_NAME" and "value" not in result:
                    result["value"] = v
                else:
                    _search(v)
        elif isinstance(node, list):
            for item in node:
                _search(item)

    _search(parsed_yaml)
    return result.get("value")


def build_docker_image_index() -> dict:
    """
    One-time scan (cached to DOCKER_IMAGE_INDEX_CACHE_FILE) across every
    project under the group: reads root .gitlab-ci.yml plus any nested
    ci-job-config.yml, and indexes every DOCKER_IMAGE_NAME value found ->
    which project it came from. Delete the cache file to force a rescan
    (e.g. after new repos are added).
    """
    if os.path.exists(DOCKER_IMAGE_INDEX_CACHE_FILE):
        logger.info("Loading DOCKER_IMAGE_NAME index from cache: %s "
                    "(delete this file to force a fresh scan)", DOCKER_IMAGE_INDEX_CACHE_FILE)
        with open(DOCKER_IMAGE_INDEX_CACHE_FILE, "r", encoding="utf-8") as f:
            return json.load(f)

    if not GITLAB_GROUP_PATH:
        logger.warning("No group_path configured — Method 2 needs a bounded group to scan and will be skipped.")
        return {}

    projects = get_group_projects_cached()
    index = {}
    logger.info("Building DOCKER_IMAGE_NAME index across %s project(s) — this runs ONCE and is cached to disk.",
                len(projects))

    for i, project in enumerate(projects, start=1):
        if i % 10 == 0 or i == len(projects):
            logger.info("  [Method 2 index] scanned %s/%s projects...", i, len(projects))

        candidates = [".gitlab-ci.yml"]
        try:
            candidates += list_repo_yml_paths(project["id"], "ci-job-config.yml")
        except requests.HTTPError as exc:
            logger.warning("Could not list repo tree for %s: %s", project.get("path_with_namespace"), exc)

        for path in candidates:
            parsed = read_gitlab_file(project["id"], path)
            if not parsed:
                continue
            image_name = find_docker_image_name(parsed)
            if image_name and image_name not in index:
                index[image_name] = {
                    "project_id": project["id"],
                    "project_path": project.get("path_with_namespace", project.get("name")),
                    "default_branch": project.get("default_branch") or GITLAB_BRANCH,
                }
        time.sleep(SLEEP_BETWEEN_PROJECT_SCANS)

    with open(DOCKER_IMAGE_INDEX_CACHE_FILE, "w", encoding="utf-8") as f:
        json.dump(index, f, indent=2)
    logger.info("DOCKER_IMAGE_NAME index built: %s image name(s) found across %s project(s). Cached to %s",
                len(index), len(projects), DOCKER_IMAGE_INDEX_CACHE_FILE)
    return index


_docker_image_index: dict | None = None


def get_docker_image_index() -> dict:
    global _docker_image_index
    if _docker_image_index is None:
        _docker_image_index = build_docker_image_index()
    return _docker_image_index


def resolve_project_method2(artifact_name: str):
    """Returns (project_dict_or_None, note)."""
    if not artifact_name:
        return None, "Method 2: no artifact name to match against"
    entry = get_docker_image_index().get(artifact_name)
    if not entry:
        return None, f"Method 2: no DOCKER_IMAGE_NAME match found for '{artifact_name}'"
    project = get_project_by_path(entry["project_path"])
    if not project:
        return None, f"Method 2: indexed project '{entry['project_path']}' no longer found on GitLab"
    return project, ""


# ===========================================================================
# Master Excel file — generic read/write helpers
# ===========================================================================


def get_cell(ws, row_idx: int, column_name: str):
    col_idx = ALL_COLUMNS.index(column_name) + 1
    return ws.cell(row=row_idx, column=col_idx).value


def set_owned_cell(ws, row_idx: int, column_name: str, value):
    if column_name not in OWNED_COLUMNS:
        raise ValueError(f"Script 2 tried to write a column it doesn't own: {column_name}")
    col_idx = ALL_COLUMNS.index(column_name) + 1
    ws.cell(row=row_idx, column=col_idx, value=value)


def set_owned_link_cell(ws, row_idx: int, column_name: str, url: str):
    if column_name not in OWNED_COLUMNS:
        raise ValueError(f"Script 2 tried to write a column it doesn't own: {column_name}")
    col_idx = ALL_COLUMNS.index(column_name) + 1
    cell = ws.cell(row=row_idx, column=col_idx, value=url)
    if url:
        cell.hyperlink = url
        cell.style = "Hyperlink"


def set_script_comment(ws, row_idx: int, script_tag: str, text: str):
    col_idx = ALL_COLUMNS.index("Comments") + 1
    cell = ws.cell(row=row_idx, column=col_idx)
    existing_lines = (cell.value or "").split("\n")
    other_lines = [line for line in existing_lines if line and not line.startswith(f"[{script_tag} ")]
    if text:
        stamp = datetime.now().strftime("%Y-%m-%d %H:%M")
        other_lines.append(f"[{script_tag} {stamp}] {text}")
    cell.value = "\n".join(other_lines) if other_lines else None


# ===========================================================================
# Per-method, per-row processing
# ===========================================================================


def process_method(ws, row_idx: int, method: int, service_name: str, artifact_path: str,
                    artifact_name: str, notes: list[str]) -> str:
    """
    Runs one method (1 or 2) for one row. Writes that method's owned
    columns directly. Returns a short status string for logging/counters.
    """
    component_col = COMPONENT_COLUMN[method]
    filepath_col = FILEPATH_COLUMN[method]
    name_col = NEW_NAME_COLUMN[method]
    version_col = NEW_VERSION_COLUMN[method]
    label = f"Method {method}"

    existing_component = get_cell(ws, row_idx, component_col)
    existing_file_path = get_cell(ws, row_idx, filepath_col)

    # --- a full file URL pasted/already resolved: strongest override ---
    blob = parse_gitlab_blob_url(existing_file_path) if existing_file_path else None
    if blob:
        blob_project_path, blob_ref, blob_file_path = blob
        project = get_project_by_path(blob_project_path)
        if not project:
            notes.append(f"{label}: File Path URL points to project '{blob_project_path}', not found")
            return "no_project"
        parsed_yaml = read_gitlab_file(project["id"], blob_file_path, ref=blob_ref)
        if not parsed_yaml:
            notes.append(f"{label}: could not read '{blob_file_path}' at ref '{blob_ref}' from the pasted URL")
            return "no_yaml"
        set_owned_link_cell(ws, row_idx, component_col, build_project_url(project))
        set_owned_link_cell(ws, row_idx, filepath_col, existing_file_path)
        _apply_variables(ws, row_idx, name_col, version_col, parsed_yaml, blob_file_path, label, notes)
        return "ok"

    # --- resolve the project (trusting a pre-filled component if present) ---
    if existing_component:
        project = resolve_component_value(existing_component)
        if not project:
            notes.append(f"{label}: '{existing_component}' (already set) not found on GitLab — left as-is")
            return "no_project"
    elif method == 1:
        project, note = resolve_project_method1(service_name, artifact_path)
        if note:
            notes.append(note)
        if not project:
            return "no_project"
    else:
        project, note = resolve_project_method2(artifact_name)
        if note:
            notes.append(note)
        if not project:
            return "no_project"

    if not existing_component:
        set_owned_link_cell(ws, row_idx, component_col, build_project_url(project))

    # --- resolve the file (trusting a pre-filled plain path if present) ---
    if existing_file_path:
        parsed_yaml = read_gitlab_file(project["id"], existing_file_path)
        if not parsed_yaml:
            notes.append(f"{label}: '{existing_file_path}' (already set) could not be read — check the path")
            return "no_yaml"
        file_path = existing_file_path
    else:
        file_path, parsed_yaml = find_and_read_yaml_file(project["id"], list(VARIABLE_MAPPING.values()))
        if not file_path:
            notes.append(f"{label}: no yml file found containing the target variables in "
                         f"{project.get('path_with_namespace')}")
            return "no_yaml"
        ref = project.get("default_branch") or GITLAB_BRANCH
        set_owned_link_cell(ws, row_idx, filepath_col, build_blob_url(project, ref, file_path))

    _apply_variables(ws, row_idx, name_col, version_col, parsed_yaml, file_path, label, notes)
    return "ok"


def _apply_variables(ws, row_idx, name_col, version_col, parsed_yaml, file_path, label, notes):
    gitlab_values = extract_vars_from_yaml(parsed_yaml, list(VARIABLE_MAPPING.values()))
    for harness_var, gitlab_key in VARIABLE_MAPPING.items():
        col = name_col if harness_var == "ssc_appname" else version_col
        if gitlab_key not in gitlab_values:
            notes.append(f"{label}: '{gitlab_key}' not found in {file_path}")
            continue
        set_owned_cell(ws, row_idx, col, str(gitlab_values[gitlab_key]))


def process_row(ws, row_idx: int) -> tuple[str, str]:
    """Returns (method1_status, method2_status)."""
    notes = []
    svc_name = get_cell(ws, row_idx, "Service Name")
    if not svc_name:
        set_script_comment(ws, row_idx, "Script2", "Skipped: no Service Name yet (run Script 1 first)")
        return "skipped", "skipped"

    artifact_path = get_cell(ws, row_idx, "Artifact Path")
    artifact_name = (artifact_path or "").rstrip("/").split("/")[-1]

    status1 = process_method(ws, row_idx, 1, svc_name, artifact_path, artifact_name, notes)
    status2 = process_method(ws, row_idx, 2, svc_name, artifact_path, artifact_name, notes)

    set_script_comment(ws, row_idx, "Script2", "; ".join(notes))
    return status1, status2


# ===========================================================================
# Live per-row console output
# ===========================================================================

STATUS_LABELS = {
    "ok": "matched",
    "no_project": "NO MATCH",
    "no_yaml": "matched repo, but no yml with target variables found",
    "skipped": "skipped",
}


def log_row_outcome(ws, row_idx: int, status1: str, status2: str):
    svc_name = get_cell(ws, row_idx, "Service Name") or "(unknown)"
    if status1 == "skipped":
        logger.info("  -> %s: skipped (no Service Name — run Script 1 first)", svc_name)
        return
    c1 = get_cell(ws, row_idx, "Corresponding Gitlab component1") if status1 == "ok" else None
    c2 = get_cell(ws, row_idx, "Corresponding Gitlab component2") if status2 == "ok" else None
    logger.info("  -> %s: Method1 %s%s | Method2 %s%s", svc_name,
                STATUS_LABELS.get(status1, status1), f" -> {c1}" if c1 else "",
                STATUS_LABELS.get(status2, status2), f" -> {c2}" if c2 else "")


# ===========================================================================
# Main
# ===========================================================================


def main():
    logger.info("=== Script 2 (GitLab extract) started ===")

    if not os.path.exists(MASTER_EXCEL_FILE):
        sys.exit(f"{MASTER_EXCEL_FILE} not found — run Script 1 first.")

    wb = load_workbook(MASTER_EXCEL_FILE)
    ws = wb.active
    header = [c.value for c in ws[1]]
    if header != ALL_COLUMNS:
        sys.exit(
            f"{MASTER_EXCEL_FILE}'s header doesn't match the expected schema.\n"
            f"Expected: {ALL_COLUMNS}\nFound:    {header}"
        )

    total_rows = ws.max_row - 1
    counters1, counters2 = {}, {}

    for row_idx in range(2, ws.max_row + 1):
        logger.info("[row %s/%s]", row_idx - 1, total_rows)
        try:
            status1, status2 = process_row(ws, row_idx)
        except Exception as exc:  # noqa: BLE001 — keep going across all rows
            logger.error("Unhandled error on row %s: %s", row_idx, exc)
            set_script_comment(ws, row_idx, "Script2", f"ERROR: {exc}")
            status1 = status2 = "error"
        counters1[status1] = counters1.get(status1, 0) + 1
        counters2[status2] = counters2.get(status2, 0) + 1
        log_row_outcome(ws, row_idx, status1, status2)
        time.sleep(SLEEP_BETWEEN_ROWS)

    wb.save(MASTER_EXCEL_FILE)
    logger.info("=== Script 2 complete ===")
    logger.info("Method 1 summary: %s", counters1)
    logger.info("Method 2 summary: %s", counters2)
    print(f"\nDone.")
    print(f"Method 1: {counters1}")
    print(f"Method 2: {counters2}")
    print(f"Master file: {MASTER_EXCEL_FILE}")
    print(f"Log file:    {LOG_FILE}")


if __name__ == "__main__":
    main()
