# -*- coding: utf-8 -*-
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
from pathlib import Path
"""
Okara SIS - Teacher + Student Latest Attendance - V25 FIXED
Windows one-file version.

Requirements:
    py -m pip install requests openpyxl

Run:
    py Okara_Teacher_Student_Attendance.py

The script uses the public SIS endpoints confirmed during testing:
  Teacher: /attendance/get_teachers_today_attendance_stats
  Student: /attendance/get_attendance_line_stats
"""

import re
import json
import argparse
from html import unescape
import sys
import time
from datetime import datetime
from urllib.parse import urlencode
from concurrent.futures import ThreadPoolExecutor, as_completed

try:
    import requests
    from openpyxl import Workbook
    from openpyxl.utils import get_column_letter
except ImportError:
    print("Missing package. Run:")
    print("  py -m pip install requests openpyxl")
    input("Press Enter to close...")
    raise SystemExit

BASE = "https://sis.pesrp.edu.pk"
DISTRICT = "26"
TEHSILS = {
    "23": "DEPALPUR",
    "89": "OKARA",
    "102": "RENALA KHURD",
}

# User-confirmed classifications.
SECONDARY_SCHOOL_IDS = {
    "52235","52229","52233","53040","53041","53045","53586","53595","53608",
    "52242","52245","53072","53554","53587","53614","53615","53644",
}
EXCLUDED_SCHOOL_IDS = {"55048"}
EXCLUDED_EMIS = {"39399999"}

TIMEOUT = 60
RETRIES = 3
RETRY_SLEEP = 0.75


def clean(v):
    return re.sub(r"\s+", " ", str(v or "")).strip()


def number(v):
    s = re.sub(r"[^\d-]", "", str(v or ""))
    return int(s or 0)


def parse_options(text):
    """Parse SIS <option value='ID'>NAME</option> HTML."""
    found = []
    for value, name in re.findall(
        r"<option[^>]*value\s*=\s*['\"]([^'\"]*)['\"][^>]*>\s*(.*?)\s*</option>",
        text or "", flags=re.I | re.S
    ):
        name = clean(re.sub(r"<[^>]+>", " ", name))
        value = clean(value)
        if value and name and name.lower() not in ("all tehsils", "all schools"):
            found.append((value, name))
    return found


def json_or_text(resp):
    try:
        return resp.json()
    except Exception:
        return resp.text


def request(session, path, params=None):
    """GET SIS endpoint with explicit retry/backoff and strict response checks."""
    url = BASE + path
    last = None
    for attempt in range(RETRIES + 1):
        try:
            r = session.get(
                url, params=params, timeout=TIMEOUT,
                headers={
                    "Accept": "application/json, text/javascript, */*; q=0.01",
                    "X-Requested-With": "XMLHttpRequest",
                    "Referer": BASE + "/dashboard",
                    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/151 Safari/537.36",
                },
            )
            if r.status_code == 404:
                raise RuntimeError(f"404 Not Found: {r.url}")
            r.raise_for_status()
            result = json_or_text(r)
            if result is None or result == "":
                raise RuntimeError(f"Empty response from {path}")
            return result
        except Exception as e:
            last = e
            if attempt < RETRIES:
                delay = RETRY_SLEEP * (2 ** attempt)
                print(f"      retry {attempt + 1}/{RETRIES} - {e}")
                time.sleep(delay)
    raise last


def get_csrf(session):
    r = session.get(
        BASE + "/dashboard",
        timeout=max(TIMEOUT, 90),
        headers={
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/151.0.0.0 Safari/537.36",
        },
    )
    r.raise_for_status()

    # Prefer the token embedded in the dashboard HTML.  The SIS CSRF
    # middleware validates the submitted token against its CSRF cookie.
    m = re.search(
        r'name=["\']csrf_test_name["\'][^>]*value=["\']([^"\']+)',
        r.text, re.I
    )
    if m:
        return m.group(1)

    return session.cookies.get("csrf_cookie_name", "")

def get_tehsil_markaz(session, tehsil_id, csrf):
    # CONFIRMED SIS endpoint supplied by the user.
    data = request(
        session,
        "/user/get_markazes",
        {
            "tehsil": tehsil_id,
            "selectedMarkaz": "false",
            "all": "All",
            "csrf_test_name": csrf,
        },
    )

    html = data.get("html", "") if isinstance(data, dict) else str(data)
    opts = parse_options(html)
    if opts:
        return opts

    # Some SIS responses may return markaz objects instead of <option> HTML.
    if isinstance(data, list):
        out = []
        for x in data:
            if isinstance(x, dict):
                mid = x.get("id") or x.get("markaz_id") or x.get("value")
                name = x.get("name") or x.get("markaz_name") or x.get("text")
                if mid and name:
                    out.append((str(mid), clean(name)))
        if out:
            return out

    raise RuntimeError(f"No Markaz records returned for tehsil {tehsil_id}")


def get_schools(session, markaz_id, csrf):
    data = request(
        session,
        "/user/get_schools",
        {
            "markaz": markaz_id,
            "selectedSchool": "false",
            "all": "All",
            "csrf_test_name": csrf,
        },
    )
    html = data.get("html", "") if isinstance(data, dict) else str(data)
    opts = parse_options(html)
    out = []
    for sid, name in opts:
        # School option values are usually school IDs; EMIS may be embedded
        # in data attributes in some responses, so EMIS is fetched below.
        out.append((sid, "", name))
    return out


def get_school_emis_and_name(session, school_id, markaz_id, csrf, fallback_name):
    """
    Resolve EMIS without guessing a school-detail endpoint.

    In the /user/get_schools response used by this report, the school option
    text itself contains the EMIS code, e.g.:
        39310174 - GGES AMLI MOTI

    The previous version discarded that code and therefore sent an empty
    s_id_emis_code to the student-attendance endpoint. Extracting the leading
    8-digit EMIS from the returned school name fixes that problem.
    """
    name = clean(fallback_name)

    # Normal SIS format: "39310174 - SCHOOL NAME"
    m = re.search(r"(?<!\d)(\d{8})(?!\d)", name)
    if m:
        return m.group(1), name

    # Fallback in case the option value itself is an 8-digit EMIS.
    sid = clean(school_id)
    if re.fullmatch(r"\d{8}", sid):
        return sid, name

    return "", name


def wing(markaz_name, school_name, school_id):
    x = (clean(markaz_name) + " " + clean(school_name)).upper()
    if "JOYIA" in clean(markaz_name).upper():
        return "Male Wing"
    # Determine the API wing from the SCHOOL identity first.  The designation
    # endpoint accepts Female/Male (not "Secondary Wing").  GGH/GGHSS/GGPS/GGES
    # are female-school prefixes and GMH/GMHS/GMPS/GMES are male-school prefixes.
    sx = clean(school_name).upper()
    if "FEMALE" in sx or "GIRLS" in sx or sx.startswith(("GGHS", "GGHSS", "GGPS", "GGES")):
        return "Female Wing"
    # Secondary/high-school names also carry gender in their SIS prefix:
    # GGH/GGHSS = girls, GHS/GHSS = boys.  Do this BEFORE the
    # SECONDARY_SCHOOL_IDS fallback; otherwise all high schools become
    # "Secondary Wing" and the API receives the wrong wing value.
    if sx.startswith(("GHS", "GHSS")):
        return "Male Wing"
    if "MALE" in sx or "BOYS" in sx or sx.startswith(("GMHS", "GMHSS", "GMPS", "GMES")):
        return "Male Wing"
    if "FEMALE" in x or "GIRLS" in x:
        return "Female Wing"
    if "MALE" in x or "BOYS" in x:
        return "Male Wing"
    if str(school_id) in SECONDARY_SCHOOL_IDS:
        return "Secondary Wing"
    if any(k in x for k in ("GHSS", "GHS ", "GGHS", "HIGH SCHOOL", "HIGHER SECONDARY", "HSS")):
        return "Secondary Wing"
    if "GGPS" in x or "GGES" in x:
        return "Female Wing"
    if "GMPS" in x or "GMES" in x:
        return "Male Wing"
    return "Male Wing"


def request_post(session, path, payload):
    """POST SIS AJAX endpoint with browser-like headers and CSRF token."""
    url = BASE + path
    headers = {
        "Accept": "application/json, text/javascript, */*; q=0.01",
        "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
        "X-Requested-With": "XMLHttpRequest",
        "Origin": BASE,
        "Referer": BASE + "/dashboard",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/151.0.0.0 Safari/537.36",
        "Sec-Fetch-Dest": "empty",
        "Sec-Fetch-Mode": "cors",
        "Sec-Fetch-Site": "same-origin",
    }
    csrf = payload.get("csrf_test_name") if isinstance(payload, dict) else ""
    if csrf:
        headers["X-CSRF-TOKEN"] = str(csrf)

    r = session.post(url, data=payload, headers=headers, timeout=60, allow_redirects=False)
    if r.status_code == 403:
        raise RuntimeError(
            f"403 Forbidden from {path}; CSRF/session rejected. "
            f"Response: {r.text[:500]!r}"
        )
    if r.status_code in (301, 302, 303, 307, 308):
        raise RuntimeError(f"Unexpected redirect {r.status_code} from {path}: {r.headers.get('Location','')}")
    r.raise_for_status()
    try:
        return r.json()
    except ValueError:
        text = r.text.strip()
        if not text:
            raise RuntimeError(f"Empty response from {path}")
        try:
            return json.loads(text)
        except ValueError:
            raise RuntimeError(f"Non-JSON response from {path}: {text[:500]!r}")


def get_filled_staff_from_sanctioned_posts(session, district_id, tehsil_id,
                                           markaz_id, school_id, emis_code=""):
    """Read the Filled/assigned staff count from the SIS sanctioned-posts table."""
    data = request(session, "/dashboard/sanctioned_posts_tab", {
        "district_id": str(district_id), "tehsil_id": str(tehsil_id),
        "markaz_id": str(markaz_id), "school_id": str(school_id),
        "s_id_emis_code": str(emis_code or ""),
    })
    html = data.get("data") if isinstance(data, dict) else data
    html = html or ""
    if not isinstance(html, str):
        html = str(html)
    for tr in re.findall(r"<tr\b[^>]*>(.*?)</tr>", html, flags=re.I | re.S):
        cells = [clean(unescape(re.sub(r"<[^>]+>", " ", c)))
                 for c in re.findall(r"<t[dh]\b[^>]*>(.*?)</t[dh]>", tr,
                                     flags=re.I | re.S)]
        joined = " ".join(cells).strip().lower()
        if cells and (joined.startswith("total") or "overall" in joined):
            nums = [number(c) for c in cells[1:] if re.search(r"\d", c)]
            if len(nums) >= 3:
                return max(0, nums[1])  # Total, Filled, Vacant
    raise RuntimeError("Filled staff total not found in sanctioned-posts response")


def teacher_attendance(session, tehsil, markaz, school, emis_code, filled_total):
    data = request(
        session,
        "/attendance/get_teachers_today_attendance_stats",
        {
            "district": DISTRICT,
            "tehsil": tehsil,
            "markaz": markaz,
            "school": school,
            "s_id_emis_code": emis_code,
            "ony_kpztp_districts": "false",
        },
    )

    present = number(data.get("present_count", 0))
    absent = number(data.get("absent_count", 0))

    # Working Staff from school-wise Teacher & Staff designation stats is the Teacher Total.
    total = number(filled_total)
    unmarked = max(0, total - present - absent)

    return {
        "Present": present,
        "Absent": absent,
        "Unmarked": unmarked,
        "Marked": present + absent,
        "Total": total,
    }


def get_student_attendance_table_retry(session, tehsil_id, markaz_id, date_from):
    last = None
    for attempt in range(RETRIES + 1):
        try:
            return get_student_attendance_table(session, tehsil_id, markaz_id, date_from)
        except Exception as e:
            last = e
            if attempt < RETRIES:
                delay = RETRY_SLEEP * (2 ** attempt)
                print(f"      student-table retry {attempt + 1}/{RETRIES} - {e}")
                time.sleep(delay)
    raise last


def get_student_attendance_table(session, tehsil_id, markaz_id, date_from):
    """Fetch SCHOOL-level student attendance for one Markaz.

    Important SIS behavior discovered from the user's live response:
      markaz_id=""        -> Markaz summary rows
      markaz_id=<ID>      -> Emis Code - School rows

    Therefore student attendance must be requested separately for each
    Markaz. The school rows contain the 8-digit EMIS code and can be matched
    directly to /user/get_schools.
    """
    r = session.get(
        BASE + "/dashboard/attendance_table",
        params={
            "district_id": DISTRICT,
            "tehsil_id": tehsil_id,
            "markaz_id": markaz_id,
            "date_from": date_from,
            "only_kpztp_districts": "false",
        },
        timeout=max(TIMEOUT, 90),
        headers={
            "Accept": "text/html, */*; q=0.01",
            "X-Requested-With": "XMLHttpRequest",
            "Referer": BASE + "/",
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/151.0.0.0 Safari/537.36"
            ),
        },
    )
    r.raise_for_status()

    raw = r.text
    html = raw
    try:
        payload = r.json()
        if isinstance(payload, dict) and isinstance(payload.get("data"), str):
            html = payload["data"]
        elif isinstance(payload, dict):
            html = str(payload.get("data") or payload.get("html") or raw)
    except Exception:
        pass

    records = {}
    for tr in re.findall(r"<tr\b[^>]*>(.*?)</tr>", html, flags=re.I | re.S):
        cells = re.findall(r"<t[dh]\b[^>]*>(.*?)</t[dh]>", tr, flags=re.I | re.S)
        cells = [clean(unescape(re.sub(r"<[^>]+>", " ", c))) for c in cells]
        if len(cells) < 9:
            continue

        # Expected school row:
        # # | 39310174 - GGES AMLI MOTI | Enrolled | Present | % |
        # Absent | % | Not Marked | %
        m = re.search(r"(?<!\d)(\d{8})(?!\d)", cells[1])
        if not m:
            continue

        emis = m.group(1)
        enrolled = number(cells[2].replace(",", ""))
        present = number(cells[3])
        absent = number(cells[5])

        # SIS attendance pages can occasionally return a stale enrollment
        # against the latest attendance counts. Do not permit negative
        # unmarked values or an attendance total below marked attendance.
        attendance_total = max(enrolled, present + absent)
        unmarked = max(0, attendance_total - present - absent)
        marked = present + absent

        records[emis] = {
            "School": cells[1],
            "Enrolled": attendance_total,
            "Present": present,
            "Present %": (f"{present / attendance_total * 100:.1f}%"
                          if attendance_total else ""),
            "Absent": absent,
            "Absent %": (f"{absent / attendance_total * 100:.1f}%"
                         if attendance_total else ""),
            "Unmarked": unmarked,
            "Unmarked %": (f"{unmarked / attendance_total * 100:.1f}%"
                           if attendance_total else ""),
            "Marked": marked,
            "Total": attendance_total,
            "Day": date_from,
        }

    if not records:
        try:
            meta = r.json()
            meta_text = f"status={meta.get('status')}, message={meta.get('message')!r}"
        except Exception:
            meta_text = "non-JSON response"
        raise RuntimeError(
            "No SCHOOL rows returned by /dashboard/attendance_table. "
            f"Tehsil={tehsil_id}, Markaz={markaz_id}, date_from={date_from}, "
            f"HTTP={r.status_code}, {meta_text}. "
            f"HTML start={clean(html)[:800]}"
        )
    return records


def student_attendance(student_map, emis):
    rec = student_map.get(clean(emis))
    if rec is None:
        raise RuntimeError(
            f"Student attendance row not found for EMIS={emis} "
            "in /dashboard/attendance_table"
        )
    return rec


def write_live_json(rows, errors, output_path, expected_school_keys=None, collection_failed=False):
    """Write school data plus an explicit completeness/validation block."""
    headers = [
        "Tehsil","Tehsil ID","Report Markaz","Actual Markaz ID","Wing",
        "School","School ID","EMIS",
        "Teacher Present","Teacher Absent","Teacher Unmarked","Teacher Marked","Teacher Total",
        "Working Staff",
        "Student Enrolled","Student Present","Student Present %","Student Absent","Student Absent %",
        "Student Unmarked","Student Unmarked %","Student Marked","Student Total","Student Latest Day"
    ]
    data = [dict(zip(headers, r)) for r in rows]
    actual_keys = [str(d.get("EMIS") or d.get("School ID") or "") for d in data]
    actual_set = set(k for k in actual_keys if k)
    expected_count = len(set(str(x) for x in (expected_school_keys or []) if x))
    missing_count = max(0, expected_count - len(data))
    duplicate_count = len(actual_keys) - len(actual_set)
    validation = {
        "status": "COMPLETE" if (not errors and not collection_failed and expected_count > 0 and len(data) == expected_count and duplicate_count == 0) else "INCOMPLETE",
        "expected_schools": expected_count,
        "collected_schools": len(data),
        "missing_schools": missing_count,
        "duplicate_keys": max(0, duplicate_count),
        "error_count": len(errors),
        "collection_failed": bool(collection_failed),
    }
    payload = {
        "generated_at": datetime.now().astimezone().isoformat(),
        "source": BASE,
        "district": "Okara",
        "version": "V28",
        "validation": validation,
        "rows": data,
        "errors": errors,
    }
    Path(output_path).write_text(json.dumps(payload, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--github-data", default="", help="Write live JSON for GitHub Pages")
    args = parser.parse_args()

    print("OKARA SIS - TEACHER + STUDENT LATEST ATTENDANCE V25 FIXED")
    print("=" * 58)
    if args.github_data:
        choice = "4"
        print("GitHub mode: collecting all Okara tehsils")
    else:
        print("4 = All Okara tehsils")
        print("1 = Depalpur   2 = Okara   3 = Renala Khurd   4 = All")
        choice = input("Enter choice [4]: ").strip() or "4"
    selected = {
        "1": ["23"],
        "2": ["89"],
        "3": ["102"],
        "4": ["23", "89", "102"],
    }.get(choice, ["23", "89", "102"])

    s = requests.Session()

    # Single persistent connection pool; no unnecessary duplicate adapter.
    retry = Retry(
        total=RETRIES,
        connect=RETRIES,
        read=RETRIES,
        backoff_factor=0.15,
        status_forcelist=[429, 500, 502, 503, 504],
        allowed_methods=frozenset(["GET", "POST"]),
        raise_on_status=False,
    )
    adapter = HTTPAdapter(
        max_retries=retry,
        pool_connections=32,
        pool_maxsize=32,
        pool_block=False,
    )
    s.mount("https://", adapter)
    s.mount("http://", adapter)
    csrf = get_csrf(s)
    print("SIS connection: 200")
    print("CSRF:", "found" if csrf else "not found (public GET mode)")

    rows = []
    errors = []
    expected_school_keys = set()
    collection_failed = False

    for tid in selected:
        tname = TEHSILS[tid]
        print(f"\nTEHSIL: {tname} ID: {tid}")
        try:
            markazs = get_tehsil_markaz(s, tid, csrf)
        except Exception as e:
            print("  ERROR getting markaz:", e)
            errors.append([tname, "", "", "", "", "MARKAZ", str(e)])
            continue

        print("  Markaz found:", len(markazs))

        date_from = datetime.now().strftime("%d/%m/%Y")

        for mi, (mid, mname) in enumerate(markazs, 1):
            print(f"  Markaz {mi}/{len(markazs)}: {mname}")
            # Two independent Markaz-level requests run in parallel.
            with ThreadPoolExecutor(max_workers=2) as prefetch:
                school_future = prefetch.submit(get_schools, s, mid, csrf)
                student_future = prefetch.submit(
                    get_student_attendance_table_retry, s, tid, mid, date_from
                )

                try:
                    schools = school_future.result()
                except Exception as e:
                    print("    ERROR schools:", e)
                    errors.append([tname, tid, mname, mid, "", "SCHOOLS", str(e)])
                    collection_failed = True
                    continue

                print("    Schools:", len(schools))
                expected_school_keys.update(
                    clean(item[0]) for item in schools
                    if item[0] not in EXCLUDED_SCHOOL_IDS
                )
                print("    Fetching Working Staff school-wise...")

                try:
                    student_map = student_future.result()
                    print("    Student attendance rows:", len(student_map))
                except Exception as e:
                    student_map = {}
                    print("    ERROR student attendance:", e)
                    errors.append([tname, tid, mname, mid, "", "STUDENT_TABLE", str(e)])
                    collection_failed = True

            import threading
            worker_local = threading.local()

            def get_worker_session():
                if not hasattr(worker_local, "session"):
                    worker_local.session = requests.Session()
                    worker_local.session.cookies.update(s.cookies)
                    worker_local.session.headers.update({
                        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                                      "AppleWebKit/537.36 Chrome/151 Safari/537.36"
                    })
                    worker_local.session.mount(
                        "https://",
                        HTTPAdapter(max_retries=retry, pool_connections=4, pool_maxsize=4)
                    )
                return worker_local.session

            def process_school(item):
                sid, _, original_name = item
                if sid in EXCLUDED_SCHOOL_IDS:
                    return None, None

                local = get_worker_session()

                emis, sname = get_school_emis_and_name(
                    local, sid, mid, csrf, original_name
                )
                if emis in EXCLUDED_EMIS:
                    return None, None

                w = wing(mname, sname, sid)
                report_markaz = "SECONDARY-WING" if w == "Secondary Wing" else mname
                school_filled_total = 0
                school_errors = []

                try:
                    # New SIS endpoint: use Filled/assigned staff, not sanctioned Total.
                    school_filled_total = get_filled_staff_from_sanctioned_posts(
                        local, DISTRICT, tid, mid, sid, emis
                    )
                    ta = teacher_attendance(
                        local, tid, mid, sid, emis, school_filled_total
                    )

                except Exception as e:
                    school_errors.append([tname, tid, mname, mid, sname, "TEACHER", str(e)])
                    return None, school_errors

                try:
                    sa = student_attendance(student_map, emis)
                except Exception as e:
                    school_errors.append([tname, tid, mname, mid, sname, "STUDENT", str(e)])
                    return None, school_errors

                row = [
                    tname, tid, report_markaz, mid, w,
                    sname, sid, emis,
                    ta["Present"], ta["Absent"], ta["Unmarked"], ta["Marked"], ta["Total"],
                    school_filled_total,
                    sa["Enrolled"], sa["Present"], sa["Present %"],
                    sa["Absent"], sa["Absent %"], sa["Unmarked"], sa["Unmarked %"],
                    sa["Present"] + sa["Absent"], sa["Enrolled"], sa["Day"],
                ]
                return row, school_errors

            # Run independent school API calls concurrently.
            max_workers = min(16, max(1, len(schools)))
            results = [None] * len(schools)
            with ThreadPoolExecutor(max_workers=max_workers) as executor:
                future_map = {
                    executor.submit(process_school, item): i
                    for i, item in enumerate(schools)
                }
                for future in as_completed(future_map):
                    i = future_map[future]
                    try:
                        results[i] = future.result()
                    except Exception as e:
                        sid, _, sname = schools[i]
                        results[i] = (
                            None,
                            [[tname, tid, mname, mid, sname, "SCHOOL", str(e)]]
                        )
                        collection_failed = True

            for row, school_errors in results:
                if row is not None:
                    rows.append(row)
                if school_errors:
                    errors.extend(school_errors)

    stamp = datetime.now().strftime("%Y-%m-%d_%H%M")
    out = Path.cwd() / f"Okara_Teacher_Student_Attendance_{stamp}.xlsx"

    wb = Workbook()
    ws = wb.active
    ws.title = "School Attendance"
    headers = [
        "Tehsil","Tehsil ID","Report Markaz","Actual Markaz ID","Wing",
        "School","School ID","EMIS",
        "Teacher Present","Teacher Absent","Teacher Unmarked","Teacher Marked","Teacher Total",
        "Working Staff",
        "Student Enrolled","Student Present","Student Present %","Student Absent","Student Absent %","Student Unmarked","Student Unmarked %","Student Marked","Student Total",
        "Student Latest Day",
    ]
    ws.append(headers)
    for r in rows:
        ws.append(r)
    ws.column_dimensions[get_column_letter(len(headers))].hidden = True
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for col in range(1, ws.max_column + 1):
        vals = [str(ws.cell(row=i, column=col).value or "") for i in range(1, min(ws.max_row, 100) + 1)]
        width = min(42, max(12, max((len(v) for v in vals), default=12) + 2))
        ws.column_dimensions[get_column_letter(col)].width = width

    # Summary helper.
    def summary_sheet(title, group_cols):
        sw = wb.create_sheet(title)
        idx = {h:i for i,h in enumerate(headers)}
        groups = {}
        for r in rows:
            key = tuple(r[idx[c]] for c in group_cols)
            if key not in groups:
                groups[key] = [0]*10
            g = groups[key]
            g[0] += 1
            g[1] += number(r[idx["Teacher Present"]])
            g[2] += number(r[idx["Teacher Absent"]])
            g[3] += number(r[idx["Teacher Unmarked"]])
            g[4] += number(r[idx["Teacher Total"]])
            g[5] += number(r[idx["Student Enrolled"]])
            g[6] += number(r[idx["Student Present"]])
            g[7] += number(r[idx["Student Absent"]])
            g[8] += number(r[idx["Student Unmarked"]])
            g[9] += number(r[idx["Student Marked"]])
        sh = list(group_cols) + [
            "Schools","Teacher Present","Teacher Absent","Teacher Unmarked","Teacher Total",
            "Student Enrolled","Student Present","Student Absent","Student Unmarked","Student Marked"
        ]
        sw.append(sh)
        for key, g in groups.items():
            sw.append(list(key) + g)
        sw.freeze_panes = "A2"
        sw.auto_filter.ref = sw.dimensions
        for c in range(1, sw.max_column+1):
            sw.column_dimensions[get_column_letter(c)].width = min(38, max(12, len(str(sw.cell(1,c).value))+4))
        return sw

    summary_sheet("Tehsil Summary", ["Tehsil"])

    # Markaz Summary: Working Staff is the sum of the school-wise
    # Teacher & Staff endpoint totals.
    ms = wb.create_sheet("Markaz Summary")
    ms_headers = ["Tehsil","Report Markaz","Wing","Schools",
                  "Teacher Present","Teacher Absent","Teacher Unmarked",
                  "Working Staff","Teacher Marked",
                  "Student Enrolled","Student Present","Student Absent","Student Unmarked","Student Marked"]
    ms.append(ms_headers)
    idx = {h:i for i,h in enumerate(headers)}
    mgroups = {}
    for r in rows:
        key = (r[idx["Tehsil"]], r[idx["Report Markaz"]], r[idx["Wing"]])
        if key not in mgroups:
            mgroups[key] = {"schools":0,"tp":0,"ta":0,"filled":0,
                            "se":0,"sp":0,"sa":0,"su":0,"sm":0}
        g = mgroups[key]
        g["schools"] += 1
        g["tp"] += number(r[idx["Teacher Present"]])
        g["ta"] += number(r[idx["Teacher Absent"]])
        g["filled"] += number(r[idx["Teacher Total"]])
        g["se"] += number(r[idx["Student Enrolled"]])
        g["sp"] += number(r[idx["Student Present"]])
        g["sa"] += number(r[idx["Student Absent"]])
        g["su"] += number(r[idx["Student Unmarked"]])
        g["sm"] += number(r[idx["Student Marked"]])

    for key, g in mgroups.items():
        marked = g["tp"] + g["ta"]
        unmarked = max(0, g["filled"] - marked)
        ms.append(list(key) + [g["schools"], g["tp"], g["ta"], unmarked, g["filled"], marked,
                               g["se"], g["sp"], g["sa"], g["su"], g["sm"]])
    ms.freeze_panes = "A2"
    ms.auto_filter.ref = ms.dimensions
    for c in range(1, ms.max_column + 1):
        ms.column_dimensions[get_column_letter(c)].width = min(38, max(12, len(str(ms.cell(1,c).value))+4))

    summary_sheet("Wing Summary", ["Tehsil","Wing"])

    ew = wb.create_sheet("Errors")
    ew.append(["Tehsil","Tehsil ID","Markaz","Markaz ID","School","Type","Error"])
    for e in errors:
        ew.append(e)

    wb.save(out)

    if args.github_data:
        write_live_json(rows, errors, args.github_data, expected_school_keys, collection_failed)
        print("Live JSON:", args.github_data)

    print("\nDATA VALIDATION:")
    print("  Expected schools:", len(expected_school_keys))
    print("  Collected schools:", len(rows))
    print("  Collection errors:", len(errors))
    print("  Status:", "COMPLETE" if (not errors and not collection_failed and expected_school_keys and len(rows) == len(expected_school_keys) and len({str(r[7] or r[6]) for r in rows}) == len(rows)) else "INCOMPLETE")
    print("  Staff: working-staff denominator uses SIS working/designation data and excludes exposed LPR/non-working counts.")
    print("  Students: enrollment is normalized so Present + Absent cannot exceed Total and Unmarked cannot be negative.")
    print("\nDONE")
    print("Schools:", len(rows))
    print("Errors:", len(errors))
    print("Excel:", out)
    if not args.github_data:
        input("\nPress Enter to close...")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nStopped by user.")
    except Exception as e:
        print("\nERROR:")
        print(e)
        raise
