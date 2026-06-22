#!/usr/bin/env python3
"""Build generated JSON/JS data for the static X-LANCE site."""

from __future__ import annotations

import argparse
import json
import math
import re
import time
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from openpyxl import load_workbook


SCRIPT_DIR = Path(__file__).resolve().parent
SITE_DIR = SCRIPT_DIR.parent

DEFAULT_MEMBERS_XLSX = SCRIPT_DIR / "members/final.xlsx"
DEFAULT_PAPERS_XLSX = SCRIPT_DIR / "papers/papers.xlsx"
DEFAULT_MEMBERS_JSON = SITE_DIR / "assets/data/members.json"
DEFAULT_PAPERS_JSON = SITE_DIR / "assets/data/papers.json"
DEFAULT_MEMBERS_JS = SITE_DIR / "js/members-generated.js"
DEFAULT_PAPERS_JS = SITE_DIR / "js/papers-generated.js"
DEFAULT_MEMBERS_BASE = SITE_DIR / "js/members-data.js"

DEFAULT_IMAGE = "assets/img/members/octocat-placeholder.png"
POSTDOC_NAMES = {"缪庆亮", "彭坤杨", "刘贝易"}

SECTION_TITLES = {
    "postdoc": {"zh": "博士后", "en": "Postdocs"},
    "phd": {"zh": "博士", "en": "PhD"},
    "master": {"zh": "硕士", "en": "Master"},
    "undergraduate": {"zh": "本科生", "en": "Undergraduates"},
}

TYPE_LABELS = {
    "journal": {"zh": "期刊", "en": "Journal"},
    "conference": {"zh": "会议", "en": "Conference"},
    "book": {"zh": "书章", "en": "Book"},
    "other": {"zh": "其他", "en": "Other"},
}


def normalize_path(value: str | Path) -> Path:
    text = str(value)
    match = re.match(r"^([A-Za-z]):[\\/](.*)$", text)
    if match:
        drive = match.group(1).lower()
        rest = match.group(2).replace("\\", "/")
        return Path("/mnt") / drive / rest
    return Path(text).expanduser()


def display_source(path: Path) -> str:
    try:
        return str(path.resolve().relative_to(SCRIPT_DIR.resolve())).replace("\\", "/")
    except ValueError:
        return str(path)


def clean_cell(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and math.isnan(value):
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return re.sub(r"\s+", " ", str(value)).strip()


def clean_degree(value: str) -> str:
    return re.sub(r"[^A-Za-z]", "", value).upper()


def normalize_image(value: str) -> str:
    image = clean_cell(value)
    if not image:
        return DEFAULT_IMAGE
    image = image.replace("\\", "/")
    if image.startswith("/"):
        image = image[1:]
    return image or DEFAULT_IMAGE


def normalize_url(value: str) -> str:
    url = clean_cell(value)
    if not url:
        return ""
    if url.startswith(("http://", "https://", "mailto:")):
        return url
    return f"https://{url}"


def member_id(raw_id: str, degree: str) -> str:
    value = clean_cell(raw_id)
    if not value:
        return ""
    if re.fullmatch(r"\d+(?:\.0)?", value):
        number = int(float(value))
        if number <= 0:
            return ""
        return f"{number:03d}-{degree}" if degree else f"{number:03d}"
    if "-" in value:
        return value
    return f"{value}-{degree}" if degree else value


def row_dicts(path: Path) -> list[dict[str, str]]:
    workbook = load_workbook(path, read_only=True, data_only=True)
    sheet = workbook.active
    rows = list(sheet.iter_rows(values_only=True))
    if not rows:
        return []
    headers = [clean_cell(cell) for cell in rows[0]]
    result = []
    for row in rows[1:]:
        item = {headers[index]: clean_cell(row[index] if index < len(row) else "") for index in range(len(headers))}
        if any(item.values()):
            result.append(item)
    return result


def load_base_members(path: Path) -> dict[str, Any]:
    text = path.read_text(encoding="utf-8")
    match = re.search(r"window\.XLANCE_MEMBERS\s*=\s*(\{.*\})\s*;?\s*$", text, re.S)
    if not match:
        raise ValueError(f"Could not parse {path}")
    return json.loads(match.group(1))


def is_inactive(state: str) -> bool:
    lowered = state.lower()
    return any(marker in state for marker in ["离开", "毕业", "校友"]) or "alumni" in lowered or "left" in lowered


def section_key(name_zh: str, degree: str) -> str:
    if name_zh in POSTDOC_NAMES or "D" in degree and "POST" in degree:
        return "postdoc"
    if "P" in degree:
        return "phd"
    if "M" in degree:
        return "master"
    return "undergraduate"


def person_from_row(row: dict[str, str]) -> dict[str, Any]:
    name_zh = row.get("name", "")
    name_en = row.get("eng_name", "")
    degree = clean_degree(row.get("degree", ""))
    person = {
        "name": {"zh": name_zh, "en": name_en or name_zh},
        "id": member_id(row.get("xlanceid", ""), degree),
        "image": normalize_image(row.get("pic", "")),
        "link": normalize_url(row.get("webpage", "")),
    }
    return person


def build_members(members_xlsx: Path, base_members_js: Path) -> dict[str, Any]:
    base = load_base_members(base_members_js)
    sections = {key: [] for key in SECTION_TITLES}
    alumni = []

    for row in row_dicts(members_xlsx):
        name_zh = row.get("name", "")
        if not name_zh:
            continue
        degree = clean_degree(row.get("degree", ""))
        if "F" in degree and name_zh not in POSTDOC_NAMES:
            continue
        person = person_from_row(row)
        if is_inactive(row.get("state", "")):
            alumni.append(person)
        else:
            sections[section_key(name_zh, degree)].append(person)

    base["source"] = {
        "zh": "教师卡片保留静态站点中的特殊信息；在校生和普通校友由本地维护表自动生成。",
        "en": "Faculty cards keep the curated static-site data; current students and ordinary alumni are generated from the local maintenance spreadsheet.",
    }
    base["students"] = [
        {"key": key, "title": SECTION_TITLES[key], "members": sections[key]}
        for key in ["postdoc", "phd", "master", "undergraduate"]
    ]
    base["alumni"] = alumni
    return base


def split_files(value: str) -> list[str]:
    if not value:
        return []
    return [item.strip() for item in re.split(r"[;\n]+", value) if item.strip()]


def paper_links(value: str) -> list[dict[str, Any]]:
    try:
        raw_links = json.loads(clean_cell(value))
    except (TypeError, json.JSONDecodeError):
        return []
    if not isinstance(raw_links, list):
        return []

    links = []
    for item in raw_links:
        if not isinstance(item, dict):
            continue
        url = clean_cell(item.get("url", ""))
        if not url.startswith(("http://", "https://")) or any(link["url"] == url for link in links):
            continue
        label = item.get("label", {})
        if isinstance(label, dict):
            zh = clean_cell(label.get("zh", ""))
            en = clean_cell(label.get("en", ""))
        else:
            zh = en = clean_cell(label)
        links.append({"label": {"zh": zh or en or "链接", "en": en or zh or "Link"}, "url": url})
    return links


def build_papers(papers_xlsx: Path) -> dict[str, Any]:
    grouped: dict[str, list[dict[str, Any]]] = {}
    for row in row_dicts(papers_xlsx):
        year = row.get("year", "")
        if not year:
            continue
        paper_type = row.get("type", "conference") or "conference"
        item = {
            "type": paper_type,
            "authors": row.get("authors", ""),
            "title": row.get("title", ""),
            "venueShort": row.get("venue_short", ""),
            "citation": row.get("citation", ""),
            "files": split_files(row.get("files", "")),
            "links": paper_links(row.get("links", "")),
            "award": row.get("award", ""),
        }
        grouped.setdefault(year, []).append(item)

    years = [
        {"year": year, "count": len(items), "items": items}
        for year, items in sorted(grouped.items(), key=lambda pair: pair[0], reverse=True)
    ]
    return {
        "source": display_source(papers_xlsx),
        "typeLabels": TYPE_LABELS,
        "years": years,
    }


def write_json_and_wrapper(data: dict[str, Any], output_json: Path, output_js: Path, global_name: str) -> None:
    output_json.parent.mkdir(parents=True, exist_ok=True)
    output_js.parent.mkdir(parents=True, exist_ok=True)
    json_text = json.dumps(data, ensure_ascii=False, indent=2)
    output_json.write_text(json_text + "\n", encoding="utf-8")
    output_js.write_text(
        f"window.{global_name} = " + json_text + ";\n",
        encoding="utf-8",
    )


def build(args: argparse.Namespace) -> dict[str, Any]:
    members_xlsx = normalize_path(args.members)
    papers_xlsx = normalize_path(args.papers)
    members_json = normalize_path(args.members_json)
    papers_json = normalize_path(args.papers_json)
    members_js = normalize_path(args.members_js)
    papers_js = normalize_path(args.papers_js)
    base_members_js = normalize_path(args.base_members)

    generated_at = datetime.now(timezone.utc).isoformat()
    members = build_members(members_xlsx, base_members_js)
    members["generatedAt"] = generated_at
    members["sourceFile"] = display_source(members_xlsx)
    papers = build_papers(papers_xlsx)
    papers["generatedAt"] = generated_at
    write_json_and_wrapper(members, members_json, members_js, "XLANCE_MEMBERS")
    write_json_and_wrapper(papers, papers_json, papers_js, "XLANCE_FULL_PAPERS")
    student_count = sum(len(section["members"]) for section in members["students"])
    alumni_count = len(members["alumni"])
    paper_count = sum(year["count"] for year in papers["years"])
    print(f"wrote {members_json}")
    print(f"wrote {papers_json}")
    print(f"wrote {members_js}")
    print(f"wrote {papers_js}")
    print(f"members: students={student_count}, alumni={alumni_count}; papers={paper_count}")
    return {"members": members, "papers": papers}


def watch(args: argparse.Namespace) -> None:
    watched = [normalize_path(args.members), normalize_path(args.papers), normalize_path(args.base_members)]
    last_mtimes: dict[Path, float] = {}
    while True:
        changed = False
        for path in watched:
            mtime = path.stat().st_mtime
            if last_mtimes.get(path) != mtime:
                changed = True
                last_mtimes[path] = mtime
        if changed:
            build(args)
        time.sleep(args.interval)


def main() -> None:
    parser = argparse.ArgumentParser(description="Convert member and paper spreadsheets into JSON/JS site data.")
    parser.add_argument("--members", default=str(DEFAULT_MEMBERS_XLSX), help="Path to process_avatar/final.xlsx.")
    parser.add_argument("--papers", default=str(DEFAULT_PAPERS_XLSX), help="Path to generated papers.xlsx.")
    parser.add_argument("--members-json", default=str(DEFAULT_MEMBERS_JSON), help="Output members JSON path.")
    parser.add_argument("--papers-json", default=str(DEFAULT_PAPERS_JSON), help="Output papers JSON path.")
    parser.add_argument("--members-js", default=str(DEFAULT_MEMBERS_JS), help="Output members JS fallback path.")
    parser.add_argument("--papers-js", default=str(DEFAULT_PAPERS_JS), help="Output papers JS fallback path.")
    parser.add_argument("--base-members", default=str(DEFAULT_MEMBERS_BASE), help="Existing curated members-data.js.")
    parser.add_argument("--watch", action="store_true", help="Rebuild whenever either spreadsheet changes.")
    parser.add_argument("--interval", type=float, default=2.0, help="Polling interval for --watch.")
    args = parser.parse_args()

    if args.watch:
        watch(args)
    else:
        build(args)


if __name__ == "__main__":
    main()
