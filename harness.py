"""
Harness Service Detail Report (read-only, no changes made to anything).

For every service in the configured org/project, pulls out:
  - Advanced > Variables          (e.g. ssc_appname, ssc_appversion)
  - Manifest(s): identifier, git fetch type, BRANCH, file/folder path(s),
    values.yaml path(s)
  - Artifact source(s): identifier, type, repository, ARTIFACT PATH

This is the discovery step before any update — run this first, eyeball the
Excel output, confirm the field names/paths look right for your services,
*then* move to the update script.

pip install requests pyyaml openpyxl --break-system-packages
"""

import json
import time
import logging
from datetime import datetime

import yaml
import requests
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

# ===========================================================================
# CONFIG
# ===========================================================================

HARNESS_BASE_URL = "https://app.harness.io/gateway"
HARNESS_API_KEY = "YOUR_HARNESS_API_KEY"
HARNESS_ACCOUNT_ID = "YOUR_ACCOUNT_ID"
HARNESS_ORG_ID = "default"
HARNESS_PROJECT_ID = "YOUR_PROJECT_ID"

# Which Advanced > Variables to pull out into their own columns
# (name -> report column label). Anything else found is still captured in
# the "All Variables" column.
VARIABLES_OF_INTEREST = {
    "ssc_appname": "ssc_appname",
    "ssc_appversion": "ssc_appversion",
}

SLEEP_BETWEEN_SERVICES = 0.15
MAX_RETRIES = 3
RETRY_BACKOFF_SECONDS = 2

RUN_ID = datetime.now().strftime("%Y%m%d_%H%M%S")
LOG_FILE = f"harness_service_report_{RUN_ID}.log"
REPORT_FILE = f"harness_service_report_{RUN_ID}.xlsx"

# ===========================================================================
# Logging
# ===========================================================================

logger = logging.getLogger("report")
logger.setLevel(logging.DEBUG)

fh = logging.FileHandler(LOG_FILE, encoding="utf-8")
fh.setLevel(logging.DEBUG)
fh.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))

ch = logging.StreamHandler()
ch.setLevel(logging.INFO)
ch.setFormatter(logging.Formatter("[%(levelname)s] %(message)s"))

logger.addHandler(fh)
logger.addHandler(ch)

# ===========================================================================
# Harness API helpers
# ===========================================================================

HEADERS = {"x-api-key": HARNESS_API_KEY, "Content-Type": "application/json"}
COMMON_PARAMS = {
    "accountIdentifier": HARNESS_ACCOUNT_ID,
    "orgIdentifier": HARNESS_ORG_ID,
    "projectIdentifier": HARNESS_PROJECT_ID,
}


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


def list_all_services() -> list[dict]:
    services = []
    page, size = 0, 100
    while True:
        params = {**COMMON_PARAMS, "page": page, "size": size}
        resp = request_with_retry("GET", f"{HARNESS_BASE_URL}/ng/api/servicesV2",
                                   headers=HEADERS, params=params)
        resp.raise_for_status()
        content = resp.json()["data"]["content"]
        if not content:
            break
        for item in content:
            svc = item["service"]
            services.append({"identifier": svc["identifier"], "name": svc["name"]})
        if len(content) < size:
            break
        page += 1
    logger.info("Fetched %s services", len(services))
    return services


def get_service_yaml(service_identifier: str) -> dict:
    url = f"{HARNESS_BASE_URL}/ng/api/servicesV2/{service_identifier}"
    resp = request_with_retry("GET", url, headers=HEADERS, params=COMMON_PARAMS)
    resp.raise_for_status()
    raw_yaml = resp.json()["data"]["service"]["yaml"]
    return yaml.safe_load(raw_yaml) or {}


# ===========================================================================
# Extraction logic
# ===========================================================================


def get_spec(yaml_dict: dict) -> dict:
    return (
        yaml_dict.get("service", {})
        .get("serviceDefinition", {})
        .get("spec", {})
    ) or {}


def extract_variables(spec: dict) -> dict:
    """Returns {var_name: value} for every Advanced > Variables entry."""
    out = {}
    for var in spec.get("variables", []) or []:
        out[var.get("name")] = var.get("value")
    return out


def extract_manifests(spec: dict) -> list[dict]:
    """One dict per manifest entry: identifier, branch, paths, etc."""
    rows = []
    for entry in spec.get("manifests", []) or []:
        m = entry.get("manifest", {}) or {}
        m_spec = m.get("spec", {}) or {}
        store = m_spec.get("store", {}) or {}
        store_spec = store.get("spec", {}) or {}

        rows.append({
            "identifier": m.get("identifier"),
            "type": m.get("type"),
            "store_type": store.get("type"),
            "connector_ref": store_spec.get("connectorRef"),
            "git_fetch_type": store_spec.get("gitFetchType"),
            "branch": store_spec.get("branch"),
            "commit_id": store_spec.get("commitId"),
            "paths": "; ".join(store_spec.get("paths", []) or []),
            "values_paths": "; ".join(m_spec.get("valuesPaths", []) or []),
        })
    return rows


# Field names Harness uses for "artifact path" across the various source
# types (Docker Registry uses imagePath, Nexus3/Artifactory use
# artifactPath, etc.) — checked in this order, first match wins.
ARTIFACT_PATH_FIELDS = ["artifactPath", "imagePath", "artifactDirectory", "repositoryUrl"]


def _extract_artifact_entry(identifier: str, a_type: str, a_spec: dict, role: str) -> dict:
    artifact_path = None
    for field in ARTIFACT_PATH_FIELDS:
        if a_spec.get(field):
            artifact_path = a_spec.get(field)
            break
    return {
        "role": role,
        "identifier": identifier,
        "type": a_type,
        "connector_ref": a_spec.get("connectorRef"),
        "repository": a_spec.get("repository"),
        "repository_format": a_spec.get("repositoryFormat"),
        "artifact_path": artifact_path,
        "tag": a_spec.get("tag"),
        "raw_spec": json.dumps(a_spec, default=str),
    }


def extract_artifacts(spec: dict) -> list[dict]:
    rows = []
    artifacts = spec.get("artifacts", {}) or {}

    primary = artifacts.get("primary", {}) or {}
    if "sources" in primary:  # newer multi-source schema
        for src in primary.get("sources", []) or []:
            rows.append(_extract_artifact_entry(
                src.get("identifier"), src.get("type"), src.get("spec", {}) or {}, "primary"
            ))
    elif primary.get("spec"):  # older single-artifact schema
        rows.append(_extract_artifact_entry(
            primary.get("identifier", "primary"), primary.get("type"),
            primary.get("spec", {}) or {}, "primary"
        ))

    for sidecar in artifacts.get("sidecars", []) or []:
        s = sidecar.get("sidecar", {}) or {}
        rows.append(_extract_artifact_entry(
            s.get("identifier"), s.get("type"), s.get("spec", {}) or {}, "sidecar"
        ))

    return rows


# ===========================================================================
# Report
# ===========================================================================

REPORT_COLUMNS = [
    "Service Identifier", "Service Name",
    "ssc_appname", "ssc_appversion", "All Variables",
    "Manifest Identifier(s)", "Branch(es)", "Git Fetch Type(s)",
    "File/Folder Path(s)", "Values.yaml Path(s)",
    "Artifact Source Identifier(s)", "Artifact Type(s)",
    "Repository", "Artifact Path(s)",
    "Notes / Error",
]

report_rows: list[list] = []


def build_row(svc_id, svc_name, variables, manifests, artifacts, notes=""):
    var_lookup = {k: variables.get(k, "") for k in VARIABLES_OF_INTEREST}
    all_vars_str = "; ".join(f"{k}={v}" for k, v in variables.items())

    manifest_ids = "; ".join(m["identifier"] or "" for m in manifests)
    branches = "; ".join((m["branch"] or m["commit_id"] or "") for m in manifests)
    fetch_types = "; ".join(m["git_fetch_type"] or "" for m in manifests)
    paths = "; ".join(m["paths"] for m in manifests if m["paths"])
    values_paths = "; ".join(m["values_paths"] for m in manifests if m["values_paths"])

    artifact_ids = "; ".join(a["identifier"] or "" for a in artifacts)
    artifact_types = "; ".join(a["type"] or "" for a in artifacts)
    repos = "; ".join(a["repository"] or "" for a in artifacts if a["repository"])
    artifact_paths = "; ".join(a["artifact_path"] or "" for a in artifacts if a["artifact_path"])

    report_rows.append([
        svc_id, svc_name,
        var_lookup.get("ssc_appname", ""), var_lookup.get("ssc_appversion", ""),
        all_vars_str,
        manifest_ids, branches, fetch_types, paths, values_paths,
        artifact_ids, artifact_types, repos, artifact_paths,
        notes,
    ])


def write_excel_report():
    wb = Workbook()
    ws = wb.active
    ws.title = "Service Details"
    ws.append(REPORT_COLUMNS)

    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    for col_idx in range(1, len(REPORT_COLUMNS) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font

    error_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    notes_col_idx = REPORT_COLUMNS.index("Notes / Error") + 1

    for row_idx, row in enumerate(report_rows, start=2):
        ws.append(row)
        if row[notes_col_idx - 1]:
            for col_idx in range(1, len(REPORT_COLUMNS) + 1):
                ws.cell(row=row_idx, column=col_idx).fill = error_fill

    for col_idx, header in enumerate(REPORT_COLUMNS, start=1):
        ws.column_dimensions[get_column_letter(col_idx)].width = max(len(header) + 2, 20)

    ws.freeze_panes = "A2"
    wb.save(REPORT_FILE)
    logger.info("Excel report written to %s", REPORT_FILE)


# ===========================================================================
# Main
# ===========================================================================


def main():
    logger.info("=== Service detail report run %s started ===", RUN_ID)
    services = list_all_services()

    for i, service in enumerate(services, start=1):
        svc_id, svc_name = service["identifier"], service["name"]
        logger.info("[%s/%s] %s (%s)", i, len(services), svc_name, svc_id)
        try:
            yaml_dict = get_service_yaml(svc_id)
            spec = get_spec(yaml_dict)
            variables = extract_variables(spec)
            manifests = extract_manifests(spec)
            artifacts = extract_artifacts(spec)

            notes = []
            if not manifests:
                notes.append("no manifests found")
            if not artifacts:
                notes.append("no artifact sources found")
            if not variables:
                notes.append("no variables found")

            build_row(svc_id, svc_name, variables, manifests, artifacts, "; ".join(notes))

        except Exception as exc:  # noqa: BLE001 — keep going across 330 services
            logger.error("Failed on service '%s': %s", svc_name, exc)
            build_row(svc_id, svc_name, {}, [], [], notes=f"ERROR: {exc}")

        time.sleep(SLEEP_BETWEEN_SERVICES)

    write_excel_report()
    logger.info("=== Run complete: %s services processed ===", len(services))
    print(f"\nDone. {len(services)} services processed.")
    print(f"Log file:     {LOG_FILE}")
    print(f"Excel report: {REPORT_FILE}")


if __name__ == "__main__":
    main()
