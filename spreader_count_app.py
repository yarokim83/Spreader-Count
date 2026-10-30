# -*- coding: utf-8 -*-
"""
ARMGC Spreader Count 자동 산정 및 DB 이력 관리 시스템 (Python Desktop GUI)
===========================================================================
- 관리 엑셀 파일(ARMGC Spreader Count List) 기반 누적 작업수(TLC 트위스트락횟수) 산정
- 대시보드 현황 사진(제미나이 OCR 판독 데이터 내장 + 편집/일괄 붙여넣기 가능) 비교를 통한 교체 호기 자동 감지
- 교환일 미상 시 기간 중간일(50:50 균등 배분) 자동 산정, 확인된 교환일/호기 직접 입력 시 일할 계산 반영
- SQLite 로컬 DB (spreader_history.db) 연동:
  1) 산정 회차별 이력 저장 / 불러오기 / 삭제
  2) 최신 누적값을 다음 회차의 '기준 누적값'으로 원클릭 전환 (연속 관리)
  3) 조회기간(시작일 ~ 종료일) 설정 시 해당 기간 동안 각 Spreader 번호별 사용 카운트(TLC) 집계 및 상세 이력 조회
"""

import os
import sys
import re
import sqlite3
import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side


def get_app_dir() -> str:
    """실행 파일(.exe) 또는 스크립트(.py)가 위치한 기준 폴더 반환"""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


DB_FILE_PATH = os.path.join(get_app_dir(), "spreader_history.db")

# ---------------------------------------------------------------------------
# 1. 기본 데이터 정의 (업로드된 엑셀 2026-08-31 및 대시보드 사진 9/14, 10/19 OCR 판독값)
# ---------------------------------------------------------------------------

ARMGC_CRANES = [
    "211", "212", "213", "214", "215", "216",
    "221", "222", "223", "224", "225", "226",
    "231", "232", "233", "234", "235", "236",
    "241", "242", "243", "244", "245", "246",
    "251", "252", "253", "254", "255", "256",
    "261", "262", "263", "264", "265", "266",
    "271", "272"
]

# 2026-08-31 기준 엑셀 초기 데이터 (Spreader ID: (장착호기, 누적Count))
INITIAL_EXCEL_20260831 = {
    "Y09": ("244", 196600),
    "Y08": ("Spare", 193240),
    "Y31": ("226", 193150),
    "Y12": ("261", 193098),
    "Y24": ("235", 192200),
    "Y28": ("Spare", 190700),
    "Y14": ("251", 186455),
    "Y02": ("243", 184845),
    "Y21": ("234", 183160),
    "Y27": ("214", 182992),
    "Y15": ("264", 182970),
    "Y01": ("Spare", 182962),
    "Y25": ("216", 182875),
    "Y10": ("262", 182830),
    "Y19": ("Spare", 181790),
    "Y32": ("231", 181700),
    "Y30": ("224", 181575),
    "Y20": ("232", 181550),
    "Y13": ("215", 181182),
    "Y17": ("211", 180975),
    "Y06": ("266", 147450),
    "Y07": ("246", 146750),
    "Y18": ("212", 121600),
    "Y11": ("241", 115400),
    "Y37": ("245", 115000),
    "Y36": ("222", 107450),
    "Y04": ("223", 98998),
    "Y35": ("236", 97816),
    "Y38": ("263", 94631),
    "Y34": ("254", 93711),
    "Y22": ("225", 92576),
    "Y29": ("242", 92249),
    "Y39": ("221", 87871),
    "Y26": ("272", 87566),
    "Y03": ("233", 85863),
    "Y16": ("253", 85843),
    "Y23": ("255", 84850),
    "Y33": ("252", 73309),
    "Y40": ("265", 58483),
    "Y05": ("271", 58400),
    "Y41": ("213", 49308),
    "Y42": ("256", 28925),
    "Y43": ("Spare", 0),
    "Y44": ("Spare", 0),
}

# 제미나이 OCR 판독 결과: 두 번째 사진 (2026-09-14 대시보드 현황)
OCR_DASHBOARD_20260914 = {
    "211": "Y17", "212": "Y18", "213": "Y41", "214": "Y27", "215": "Y13", "216": "Y25",
    "221": "Y39", "222": "Y36", "223": "Y04", "224": "Y30", "225": "Y22", "226": "Y31",
    "231": "Y32", "232": "Y20", "233": "Y03", "234": "Y21", "235": "Y24", "236": "Y35",
    "241": "Y11", "242": "Y29", "243": "Y02", "244": "Y09", "245": "Y37", "246": "Y07",
    "251": "Y14", "252": "Y33", "253": "Y16", "254": "Y34", "255": "Y23", "256": "Y42",
    "261": "Y12", "262": "Y10", "263": "Y38", "264": "Y15", "265": "Y40", "266": "Y06",
    "271": "Y05", "272": "Y26"
}

# 제미나이 OCR 판독 결과: 첫 번째 사진 (2026-10-19 대시보드 현황)
OCR_DASHBOARD_20261019 = {
    "211": "Y17", "212": "Y18", "213": "Y28", "214": "Y27", "215": "Y13", "216": "Y25",
    "221": "Y39", "222": "Y36", "223": "Y04", "224": "Y30", "225": "Y22", "226": "Y44",
    "231": "Y32", "232": "Y20", "233": "Y03", "234": "Y21", "235": "Y24", "236": "Y35",
    "241": "Y11", "242": "Y29", "243": "Y02", "244": "Y43", "245": "Y37", "246": "Y07",
    "251": "Y14", "252": "Y33", "253": "Y16", "254": "Y34", "255": "Y23", "256": "Y42",
    "261": "Y12", "262": "Y10", "263": "Y38", "264": "Y15", "265": "Y40", "266": "Y06",
    "271": "Y05", "272": "Y26"
}

# 제미나이 OCR 판독 결과: 2026-10-30 대시보드 현황
OCR_DASHBOARD_20261030 = {
    "211": "Y01", "212": "Y18", "213": "Y26", "214": "Y27", "215": "Y13", "216": "Y25",
    "221": "Y39", "222": "Y36", "223": "Y04", "224": "Y30", "225": "Y22", "226": "Y44",
    "231": "Y32", "232": "Y20", "233": "Y03", "234": "Y08", "235": "Y24", "236": "Y35",
    "241": "Y21", "242": "Y29", "243": "Y02", "244": "Y43", "245": "Y37", "246": "Y07",
    "251": "Y14", "252": "Y33", "253": "Y16", "254": "Y17", "255": "Y23", "256": "Y42",
    "261": "Y12", "262": "Y10", "263": "Y38", "264": "Y15", "265": "Y40", "266": "Y06",
    "271": "Y05", "272": "Y09"
}
# 2026-10-30 예비/정비 현황: Y19(정비완료), Y11(910), Y34(910), Y41(793), Y28(793), Y31(SPARE)

# 사용자 확인 수동 교환 이력 기본 프리셋 (1차: 9/14~10/19)
DEFAULT_MANUAL_OVERRIDES = [
    {
        "crane": "226",
        "prev_spreader": "Y31",
        "curr_spreader": "Y44",
        "swap_date": "2026-09-11",
        "note": "사용자 확인 입력 (Y44 9/11 부착 - 필요시 211/226호기 변경 가능)"
    },
    {
        "crane": "244",
        "prev_spreader": "Y09",
        "curr_spreader": "Y43",
        "swap_date": "2026-09-11",
        "note": "사용자 확인 입력 (Y43 9/11 244호 부착)"
    }
]

# 1차 회차: 2026/09/14 ~ 2026/10/19 ARMGC (TC4211 ~ TC4272) 호기별 TLC (트위스트락횟수)
DEFAULT_TLC_COUNTS_20260914_20261019 = {
    "211": 8169, "212": 7895, "213": 8562, "214": 8292, "215": 7547, "216": 8129,
    "221": 8053, "222": 7476, "223": 8061, "224": 7137, "225": 8229, "226": 8111,
    "231": 7779, "232": 7465, "233": 7338, "234": 7672, "235": 7599, "236": 8106,
    "241": 7525, "242": 7824, "243": 7588, "244": 7954, "245": 4514, "246": 6269,
    "251": 8745, "252": 8207, "253": 7767, "254": 8708, "255": 4829, "256": 6648,
    "261": 7517, "262": 7651, "263": 7821, "264": 7751, "265": 4527, "266": 6400,
    "271": 2690, "272": 3446
}

# 2차 회차: 2026/10/20 ~ 2026/10/30 ARMGC (TC4211 ~ TC4272) 호기별 TLC (트위스트락횟수)
DEFAULT_TLC_COUNTS_20261020_20261030 = {
    "211": 2248, "212": 2086, "213": 2230, "214": 2384, "215": 2016, "216": 2204,
    "221": 2188, "222": 1989, "223": 2316, "224": 1908, "225": 2190, "226": 2254,
    "231": 2165, "232": 2195, "233": 1965, "234": 2110, "235": 2061, "236": 2278,
    "241": 2008, "242": 2341, "243": 2217, "244": 2300, "245": 1400, "246": 1754,
    "251": 2480, "252": 2050, "253": 2082, "254": 2356, "255": 1392, "256": 1989,
    "261": 2081, "262": 1551, "263": 1986, "264": 1872, "265": 1365, "266": 1831,
    "271":  804, "272":  948
}


# ---------------------------------------------------------------------------
# 2. 핵심 계산 엔진 (SpreaderCalculator)
# ---------------------------------------------------------------------------

def parse_date(date_str: str) -> datetime.date:
    date_str = date_str.strip().replace("/", "-").replace(".", "-")
    return datetime.date.fromisoformat(date_str)


def calculate_swap_allocation(
    period_start: datetime.date,
    period_end: datetime.date,
    prev_spreader: str,
    curr_spreader: str,
    work_count: int,
    manual_swap_date_str: str = ""
):
    """
    특정 호기의 기간(period_start ~ period_end) 내 TLC(트위스트락횟수)를
    이전 스프레더(prev_spreader)와 현재 스프레더(curr_spreader)에 배분합니다.
    """
    total_days = max((period_end - period_start).days, 1)

    if prev_spreader == curr_spreader and not manual_swap_date_str.strip():
        return {
            "is_swapped": False,
            "swap_date_str": "-",
            "eff_swap_date": period_start,
            "method": "변동없음 (100%)",
            "prev_ratio": 0.0,
            "curr_ratio": 1.0,
            "prev_added": 0,
            "curr_added": int(work_count),
        }

    if manual_swap_date_str.strip():
        try:
            swap_dt = parse_date(manual_swap_date_str)
            if swap_dt <= period_start:
                prev_days = 0
                curr_days = total_days
                eff_dt = period_start
                method_desc = f"수동지정 ({swap_dt.isoformat()}, 시작전 교체→신규 100%)"
            elif swap_dt >= period_end:
                prev_days = total_days
                curr_days = 0
                eff_dt = period_end
                method_desc = f"수동지정 ({swap_dt.isoformat()}, 종료후 교체→기존 100%)"
            else:
                prev_days = (swap_dt - period_start).days
                curr_days = total_days - prev_days
                eff_dt = swap_dt
                method_desc = f"수동지정 ({swap_dt.isoformat()}, {prev_days}일:{curr_days}일 일할)"

            prev_ratio = prev_days / total_days
            curr_ratio = curr_days / total_days
            prev_added = int(round(work_count * prev_ratio))
            curr_added = int(work_count) - prev_added
            return {
                "is_swapped": True,
                "swap_date_str": swap_dt.isoformat(),
                "eff_swap_date": eff_dt,
                "method": method_desc,
                "prev_ratio": prev_ratio,
                "curr_ratio": curr_ratio,
                "prev_added": prev_added,
                "curr_added": curr_added,
            }
        except Exception:
            pass

    mid_days = total_days / 2.0
    est_swap_dt = period_start + datetime.timedelta(days=int(round(mid_days)))
    prev_ratio = 0.5
    curr_ratio = 0.5
    prev_added = int(round(work_count * 0.5))
    curr_added = int(work_count) - prev_added

    return {
        "is_swapped": True,
        "swap_date_str": f"{est_swap_dt.isoformat()} (추정)",
        "eff_swap_date": est_swap_dt,
        "method": "자동산정 (중간일 50:50)",
        "prev_ratio": prev_ratio,
        "curr_ratio": curr_ratio,
        "prev_added": prev_added,
        "curr_added": curr_added,
    }


# ---------------------------------------------------------------------------
# 3. 텍스트/엑셀 복사-붙여넣기 파서 (TLC & 대시보드 현황)
# ---------------------------------------------------------------------------

def parse_work_count_text(raw_text: str, use_tlc: bool = True) -> dict:
    result = {}
    for line in raw_text.strip().splitlines():
        line = line.strip()
        if not line:
            continue
        tokens = re.split(r"[\t,;|\s]+", line)
        if len(tokens) < 2:
            continue
        first_tok = tokens[0].upper()
        m = re.search(r"(2[1-7][1-6])", first_tok)
        if not m:
            continue
        crane_no = m.group(1)
        if crane_no not in ARMGC_CRANES:
            continue

        num_tokens = []
        for tok in tokens[1:]:
            cleaned = tok.replace(",", "")
            if re.match(r"^\d+(\.\d+)?$", cleaned):
                num_tokens.append(cleaned)

        if num_tokens:
            if len(num_tokens) >= 4 and use_tlc:
                val = int(round(float(num_tokens[-2])))
            else:
                val = int(round(float(num_tokens[-1])))
            result[crane_no] = val
    return result


def parse_dashboard_text(raw_text: str) -> dict:
    """
    제미나이 판독 결과나 텍스트(예: '211: Y17, 212: Y18' 또는 줄바꿈 형태)에서
    호기(211~272)와 Spreader(Y01~Y99) 매핑을 추출합니다.
    """
    mapping = {}
    pairs = re.findall(r"(2[1-7][1-6])\D+(Y\d{2})", raw_text.upper())
    for crane_no, sp_id in pairs:
        if crane_no in ARMGC_CRANES:
            mapping[crane_no] = sp_id
    return mapping


# ---------------------------------------------------------------------------
# 4. SQLite 데이터베이스 매니저 (SpreaderDB)
# ---------------------------------------------------------------------------

class SpreaderDB:
    def __init__(self, db_path: str):
        self.db_path = db_path
        self._init_tables()

    def _get_conn(self):
        return sqlite3.connect(self.db_path)

    def _init_tables(self):
        with self._get_conn() as conn:
            cur = conn.cursor()
            # 1) 산정 기간(회차) 테이블
            cur.execute("""
                CREATE TABLE IF NOT EXISTS periods (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    base_date TEXT NOT NULL,
                    period_start TEXT NOT NULL,
                    period_end TEXT NOT NULL,
                    total_tlc INTEGER DEFAULT 0,
                    swapped_count INTEGER DEFAULT 0,
                    memo TEXT DEFAULT '',
                    created_at TEXT NOT NULL,
                    UNIQUE(period_start, period_end)
                )
            """)
            # 2) 회차별 호기(ARMGC 38대) 산정 및 교환 내역
            cur.execute("""
                CREATE TABLE IF NOT EXISTS period_cranes (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    period_id INTEGER NOT NULL,
                    crane TEXT NOT NULL,
                    prev_spreader TEXT NOT NULL,
                    curr_spreader TEXT NOT NULL,
                    tlc_count INTEGER NOT NULL,
                    manual_swap_date TEXT DEFAULT '',
                    calc_swap_date TEXT DEFAULT '',
                    eff_swap_date TEXT DEFAULT '',
                    method TEXT DEFAULT '',
                    prev_added INTEGER DEFAULT 0,
                    curr_added INTEGER DEFAULT 0,
                    note TEXT DEFAULT '',
                    FOREIGN KEY(period_id) REFERENCES periods(id) ON DELETE CASCADE
                )
            """)
            # 3) 회차별 Spreader 누적 및 기간 사용량 스냅샷
            cur.execute("""
                CREATE TABLE IF NOT EXISTS period_spreaders (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    period_id INTEGER NOT NULL,
                    spreader TEXT NOT NULL,
                    init_crane TEXT NOT NULL,
                    curr_crane TEXT NOT NULL,
                    base_count INTEGER NOT NULL,
                    added_as_prev INTEGER NOT NULL,
                    added_as_curr INTEGER NOT NULL,
                    total_added INTEGER NOT NULL,
                    final_count INTEGER NOT NULL,
                    history TEXT DEFAULT '',
                    FOREIGN KEY(period_id) REFERENCES periods(id) ON DELETE CASCADE
                )
            """)
            conn.commit()

    def save_period_calculation(
        self,
        base_date: str,
        period_start: str,
        period_end: str,
        memo: str,
        crane_data: dict,
        calc_crane_results: dict,
        calc_spreader_results: list
    ) -> int:
        now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        total_tlc = sum(int(v["work_count"]) for v in crane_data.values())
        swapped_cnt = sum(1 for r in calc_crane_results.values() if r["is_swapped"])

        with self._get_conn() as conn:
            cur = conn.cursor()
            # 동일 시작~종료일 데이터가 있으면 덮어쓰기(업데이트)
            cur.execute(
                "SELECT id FROM periods WHERE period_start = ? AND period_end = ?",
                (period_start, period_end)
            )
            row = cur.fetchone()
            if row:
                period_id = row[0]
                cur.execute("""
                    UPDATE periods
                    SET base_date=?, total_tlc=?, swapped_count=?, memo=?, created_at=?
                    WHERE id=?
                """, (base_date, total_tlc, swapped_cnt, memo, now_str, period_id))
                cur.execute("DELETE FROM period_cranes WHERE period_id=?", (period_id,))
                cur.execute("DELETE FROM period_spreaders WHERE period_id=?", (period_id,))
            else:
                cur.execute("""
                    INSERT INTO periods (base_date, period_start, period_end, total_tlc, swapped_count, memo, created_at)
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                """, (base_date, period_start, period_end, total_tlc, swapped_cnt, memo, now_str))
                period_id = cur.lastrowid

            # 호기별 저장
            for crane in ARMGC_CRANES:
                cinfo = crane_data[crane]
                alloc = calc_crane_results[crane]
                eff_dt_str = alloc.get("eff_swap_date", parse_date(period_start)).isoformat()
                cur.execute("""
                    INSERT INTO period_cranes (
                        period_id, crane, prev_spreader, curr_spreader, tlc_count,
                        manual_swap_date, calc_swap_date, eff_swap_date, method,
                        prev_added, curr_added, note
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    period_id, crane, cinfo["prev_spreader"], cinfo["curr_spreader"], int(cinfo["work_count"]),
                    cinfo["manual_swap_date"], alloc["swap_date_str"], eff_dt_str, alloc["method"],
                    alloc["prev_added"], alloc["curr_added"], cinfo["note"]
                ))

            # Spreader별 저장
            for srow in calc_spreader_results:
                cur.execute("""
                    INSERT INTO period_spreaders (
                        period_id, spreader, init_crane, curr_crane,
                        base_count, added_as_prev, added_as_curr, total_added, final_count, history
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    period_id, srow["spreader"], srow["init_crane"], srow["curr_crane"],
                    srow["base_count"], srow["added_as_prev"], srow["added_as_curr"],
                    srow["total_added"], srow["final_count"], srow["history"]
                ))

            conn.commit()
            return period_id

    def list_periods(self) -> list:
        with self._get_conn() as conn:
            cur = conn.cursor()
            cur.execute("""
                SELECT id, base_date, period_start, period_end, total_tlc, swapped_count, memo, created_at
                FROM periods
                ORDER BY period_end DESC, id DESC
            """)
            cols = [d[0] for d in cur.description]
            return [dict(zip(cols, r)) for r in cur.fetchall()]

    def load_period(self, period_id: int) -> dict:
        with self._get_conn() as conn:
            cur = conn.cursor()
            cur.execute("SELECT id, base_date, period_start, period_end, memo FROM periods WHERE id=?", (period_id,))
            p_row = cur.fetchone()
            if not p_row:
                return {}
            period_info = {
                "id": p_row[0],
                "base_date": p_row[1],
                "period_start": p_row[2],
                "period_end": p_row[3],
                "memo": p_row[4],
            }
            cur.execute("""
                SELECT crane, prev_spreader, curr_spreader, tlc_count, manual_swap_date, note
                FROM period_cranes WHERE period_id=?
            """, (period_id,))
            cranes = {}
            for r in cur.fetchall():
                cranes[r[0]] = {
                    "prev_spreader": r[1],
                    "curr_spreader": r[2],
                    "work_count": r[3],
                    "manual_swap_date": r[4],
                    "note": r[5],
                }

            cur.execute("""
                SELECT spreader, init_crane, base_count, curr_crane, final_count
                FROM period_spreaders WHERE period_id=?
            """, (period_id,))
            spreaders = {}
            for r in cur.fetchall():
                spreaders[r[0]] = {
                    "init_crane": r[1],
                    "base_count": r[2],
                    "curr_crane": r[3],
                    "final_count": r[4],
                }
            return {"period": period_info, "cranes": cranes, "spreaders": spreaders}

    def delete_period(self, period_id: int):
        with self._get_conn() as conn:
            cur = conn.cursor()
            cur.execute("DELETE FROM period_cranes WHERE period_id=?", (period_id,))
            cur.execute("DELETE FROM period_spreaders WHERE period_id=?", (period_id,))
            cur.execute("DELETE FROM periods WHERE id=?", (period_id,))
            conn.commit()

    def query_usage_by_date_range(self, query_start_str: str, query_end_str: str) -> list:
        """
        사용자가 지정한 조회기간 [query_start ~ query_end] 동안의
        Spreader 번호별 사용 카운트(TLC)를 DB에서 조회·집계합니다.
        - 저장된 각 회차(period_start ~ period_end)와 조회기간의 겹치는 구간을 계산하여
          정확히 일치하면 저장된 배분값을 그대로 합산하고, 부분 겹침이면 장착 구간 일할 비례로 산정합니다.
        """
        q_start = parse_date(query_start_str)
        q_end = parse_date(query_end_str)

        with self._get_conn() as conn:
            cur = conn.cursor()
            cur.execute("SELECT id, period_start, period_end FROM periods ORDER BY period_start ASC")
            all_periods = cur.fetchall()

            # 전체 등록된 Spreader 최신 상태 파악
            cur.execute("""
                SELECT spreader, curr_crane, final_count, base_count
                FROM period_spreaders
                WHERE period_id = (SELECT id FROM periods ORDER BY period_end DESC, id DESC LIMIT 1)
            """)
            latest_sp = {r[0]: {"curr_crane": r[1], "final_count": r[2]} for r in cur.fetchall()}

            usage_map = {}
            for sp_id in INITIAL_EXCEL_20260831:
                usage_map[sp_id] = {
                    "spreader": sp_id,
                    "period_usage": 0,
                    "used_cranes": set(),
                    "swap_events": [],
                    "details": [],
                    "latest_crane": latest_sp.get(sp_id, {}).get("curr_crane", INITIAL_EXCEL_20260831[sp_id][0]),
                    "latest_total": latest_sp.get(sp_id, {}).get("final_count", INITIAL_EXCEL_20260831[sp_id][1]),
                }

            matched_periods = 0
            for pid, p_start_s, p_end_s in all_periods:
                p_start = parse_date(p_start_s)
                p_end = parse_date(p_end_s)

                # 구간 겹침 확인
                ov_start = max(q_start, p_start)
                ov_end = min(q_end, p_end)
                if ov_start > ov_end:
                    continue
                if ov_start == ov_end and not (q_start == p_start == p_end == q_end):
                    continue

                matched_periods += 1
                cur.execute("""
                    SELECT crane, prev_spreader, curr_spreader, tlc_count,
                           calc_swap_date, eff_swap_date, method, prev_added, curr_added
                    FROM period_cranes
                    WHERE period_id = ?
                """, (pid,))
                c_rows = cur.fetchall()

                # 조회 기간이 해당 회차를 온전히 포함하는 경우 -> 정확한 배분 정수값 사용
                full_cover = (q_start <= p_start and q_end >= p_end)

                for (crane, prev_sp, curr_sp, tlc_cnt, calc_swap_dt, eff_swap_s, method, prev_add, curr_add) in c_rows:
                    for sp in (prev_sp, curr_sp):
                        if sp and sp not in usage_map:
                            usage_map[sp] = {
                                "spreader": sp,
                                "period_usage": 0,
                                "used_cranes": set(),
                                "swap_events": [],
                                "details": [],
                                "latest_crane": latest_sp.get(sp, {}).get("curr_crane", "Spare"),
                                "latest_total": latest_sp.get(sp, {}).get("final_count", 0),
                            }

                    if full_cover:
                        add_p = int(prev_add)
                        add_c = int(curr_add)
                    else:
                        # 부분 기간 조회 시 장착 구간별 겹치는 일수 비례 계산
                        eff_dt = parse_date(eff_swap_s) if eff_swap_s else p_start
                        # 이전 Spreader 구간: [p_start, eff_dt]
                        prev_total_days = max((eff_dt - p_start).days, 0)
                        if prev_total_days > 0 and prev_add > 0:
                            sub_s = max(q_start, p_start)
                            sub_e = min(q_end, eff_dt)
                            sub_days = max((sub_e - sub_s).days, 0)
                            add_p = int(round(prev_add * (sub_days / prev_total_days)))
                        else:
                            add_p = 0

                        # 현재 Spreader 구간: [eff_dt, p_end]
                        curr_total_days = max((p_end - eff_dt).days, 1)
                        if curr_add > 0:
                            sub_s2 = max(q_start, eff_dt)
                            sub_e2 = min(q_end, p_end)
                            sub_days2 = max((sub_e2 - sub_s2).days, 0)
                            add_c = int(round(curr_add * (sub_days2 / curr_total_days)))
                        else:
                            add_c = 0

                    if add_p > 0 and prev_sp:
                        usage_map[prev_sp]["period_usage"] += add_p
                        usage_map[prev_sp]["used_cranes"].add(f"{crane}호")
                        usage_map[prev_sp]["swap_events"].append(f"{crane}호 탈거({calc_swap_dt})")
                        usage_map[prev_sp]["details"].append(f"[{p_start_s}~{p_end_s}] {crane}호(교체전): +{add_p:,}")

                    if add_c > 0 and curr_sp:
                        usage_map[curr_sp]["period_usage"] += add_c
                        usage_map[curr_sp]["used_cranes"].add(f"{crane}호")
                        if prev_sp != curr_sp:
                            usage_map[curr_sp]["swap_events"].append(f"{crane}호 장착({calc_swap_dt})")
                            usage_map[curr_sp]["details"].append(f"[{p_start_s}~{p_end_s}] {crane}호(교체후): +{add_c:,}")
                        else:
                            usage_map[curr_sp]["details"].append(f"[{p_start_s}~{p_end_s}] {crane}호: +{add_c:,}")

            results = []
            for sp_id, data in usage_map.items():
                results.append({
                    "spreader": sp_id,
                    "period_usage": data["period_usage"],
                    "used_cranes": ", ".join(sorted(data["used_cranes"])) if data["used_cranes"] else "미사용(Spare)",
                    "swap_events": ", ".join(data["swap_events"]) if data["swap_events"] else "-",
                    "latest_crane": data["latest_crane"],
                    "latest_total": data["latest_total"],
                    "details": " / ".join(data["details"]) if data["details"] else "해당 기간 사용 이력 없음",
                    "matched_periods": matched_periods,
                })
            results.sort(key=lambda x: (-x["period_usage"], -x["latest_total"], x["spreader"]))
            return results


# ---------------------------------------------------------------------------
# 5. 메인 GUI 애플리케이션 클래스
# ---------------------------------------------------------------------------

class SpreaderCountApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("ARMGC Spreader Count 자동 산정 및 DB 이력 관리 시스템")
        self.geometry("1420x880")
        self.minsize(1180, 740)

        self.db = SpreaderDB(DB_FILE_PATH)

        # 상태 데이터 초기화
        self.excel_path = os.path.join(get_app_dir(), "ARMGC_Spreader_Count_20260831.xlsx")
        self.base_date_var = tk.StringVar(value="2026-08-31")
        self.period_start_var = tk.StringVar(value="2026-09-14")
        self.period_end_var = tk.StringVar(value="2026-10-19")
        self.period_memo_var = tk.StringVar(value="2026-09-14 ~ 2026-10-19 정기 산정 (TLC 기준)")

        # 조회기간 탭 변수
        self.query_start_var = tk.StringVar(value="2026-09-14")
        self.query_end_var = tk.StringVar(value="2026-10-19")
        self.query_search_sp_var = tk.StringVar(value="")

        self.spreader_base = {}
        self._load_default_or_excel_base()

        self.crane_data = {}
        self._init_crane_data()

        self.calc_crane_results = {}
        self.calc_spreader_results = []
        self.query_results_cache = []

        self._setup_styles()
        self._build_ui()

        # DB에 기본 1차 회차(2026-09-14~2026-10-19) 및 2차 회차(2026-10-20~2026-10-30) 시드 보장
        self._ensure_default_periods_in_db()

        # 최신 회차(2026-10-20 ~ 2026-10-30)를 기본 화면으로 로드
        periods = self.db.list_periods()
        if periods:
            latest_data = self.db.load_period(periods[0]["id"])
            if latest_data:
                pinfo = latest_data["period"]
                self.base_date_var.set(pinfo["base_date"])
                self.period_start_var.set(pinfo["period_start"])
                self.period_end_var.set(pinfo["period_end"])
                self.period_memo_var.set(pinfo["memo"])
                for sp_id, sinfo in latest_data["spreaders"].items():
                    self.spreader_base[sp_id] = {
                        "init_crane": sinfo["init_crane"],
                        "base_count": sinfo["base_count"],
                    }
                for cr, cinfo in latest_data["cranes"].items():
                    if cr in self.crane_data:
                        self.crane_data[cr] = cinfo

        self.query_start_var.set("2026-09-14")
        self.query_end_var.set(self.period_end_var.get())
        self.recalculate_all()
        self.refresh_db_views()

    def _ensure_default_periods_in_db(self):
        existing = {(p["period_start"], p["period_end"]): p["id"] for p in self.db.list_periods()}

        # 1) 1차 회차 (2026-09-14 ~ 2026-10-19) 계산 및 저장
        self.base_date_var.set("2026-08-31")
        self.period_start_var.set("2026-09-14")
        self.period_end_var.set("2026-10-19")
        self.period_memo_var.set("1차 정기 산정 (2026-09-14 ~ 2026-10-19, TLC 기준)")
        self._load_default_or_excel_base()
        self._init_crane_data()
        self.recalculate_all()

        if ("2026-09-14", "2026-10-19") not in existing:
            self.db.save_period_calculation(
                self.base_date_var.get(),
                self.period_start_var.get(),
                self.period_end_var.get(),
                self.period_memo_var.get(),
                self.crane_data,
                self.calc_crane_results,
                self.calc_spreader_results,
            )

        # 2) 2차 회차 (2026-10-20 ~ 2026-10-30): 1차 회차의 최종 누적 Count를 기준값으로 이관하여 산정
        if ("2026-10-20", "2026-10-30") not in existing:
            for srow in self.calc_spreader_results:
                sp_id = srow["spreader"]
                self.spreader_base[sp_id] = {
                    "init_crane": srow["curr_crane"],
                    "base_count": srow["final_count"],
                }
            self.base_date_var.set("2026-10-19")
            self.period_start_var.set("2026-10-20")
            self.period_end_var.set("2026-10-30")
            self.period_memo_var.set("2차 정기 산정 (2026-10-20 ~ 2026-10-30, TLC 기준)")

            for crane in ARMGC_CRANES:
                prev_sp = OCR_DASHBOARD_20261019.get(crane, "")
                curr_sp = OCR_DASHBOARD_20261030.get(crane, "")
                note = "대시보드 교체 감지 (교환일 미상 → 중간일 10/25 50:50 자동산정)" if prev_sp != curr_sp else ""
                self.crane_data[crane] = {
                    "prev_spreader": prev_sp,
                    "curr_spreader": curr_sp,
                    "work_count": DEFAULT_TLC_COUNTS_20261020_20261030.get(crane, 0),
                    "manual_swap_date": "",
                    "note": note,
                }
            self.recalculate_all()
            self.db.save_period_calculation(
                self.base_date_var.get(),
                self.period_start_var.get(),
                self.period_end_var.get(),
                self.period_memo_var.get(),
                self.crane_data,
                self.calc_crane_results,
                self.calc_spreader_results,
            )

    def _setup_styles(self):
        style = ttk.Style(self)
        if "vista" in style.theme_names():
            style.theme_use("vista")
        elif "clam" in style.theme_names():
            style.theme_use("clam")

        style.configure("Treeview", rowheight=26, font=("맑은 고딕", 10))
        style.configure("Treeview.Heading", font=("맑은 고딕", 10, "bold"))
        style.configure("Title.TLabel", font=("맑은 고딕", 13, "bold"))
        style.configure("SubTitle.TLabel", font=("맑은 고딕", 10, "bold"))
        style.configure("Info.TLabel", font=("맑은 고딕", 9))
        style.configure("Accent.TButton", font=("맑은 고딕", 10, "bold"))

    def _load_default_or_excel_base(self):
        self.spreader_base.clear()
        for sp_id, (crane, cnt) in INITIAL_EXCEL_20260831.items():
            self.spreader_base[sp_id] = {
                "init_crane": str(crane),
                "base_count": int(cnt)
            }
        if os.path.exists(self.excel_path):
            try:
                self._read_excel_file(self.excel_path)
            except Exception as e:
                print(f"초기 엑셀 로드 참고 메시지: {e}")

    def _read_excel_file(self, path: str):
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb.active
        dt_val = ws["G2"].value
        if isinstance(dt_val, (datetime.datetime, datetime.date)):
            self.base_date_var.set(dt_val.strftime("%Y-%m-%d"))
        elif dt_val:
            self.base_date_var.set(str(dt_val)[:10])

        loaded = {}
        for row_idx in range(4, 30):
            sp_left = ws.cell(row=row_idx, column=2).value
            cr_left = ws.cell(row=row_idx, column=3).value
            cnt_left = ws.cell(row=row_idx, column=4).value
            if sp_left and str(sp_left).startswith("Y"):
                loaded[str(sp_left).strip()] = {
                    "init_crane": str(cr_left).strip() if cr_left is not None else "Spare",
                    "base_count": int(round(float(cnt_left))) if cnt_left is not None else 0
                }
            sp_right = ws.cell(row=row_idx, column=6).value
            cr_right = ws.cell(row=row_idx, column=7).value
            cnt_right = ws.cell(row=row_idx, column=8).value
            if sp_right and str(sp_right).startswith("Y"):
                loaded[str(sp_right).strip()] = {
                    "init_crane": str(cr_right).strip() if cr_right is not None else "Spare",
                    "base_count": int(round(float(cnt_right))) if cnt_right is not None else 0
                }

        if loaded:
            for k, v in loaded.items():
                self.spreader_base[k] = v
            for new_sp in ["Y43", "Y44"]:
                if new_sp not in self.spreader_base:
                    self.spreader_base[new_sp] = {"init_crane": "Spare", "base_count": 0}

    def _init_crane_data(self):
        override_map = {item["crane"]: item for item in DEFAULT_MANUAL_OVERRIDES}
        for crane in ARMGC_CRANES:
            prev_sp = OCR_DASHBOARD_20260914.get(crane, "")
            curr_sp = OCR_DASHBOARD_20261019.get(crane, "")
            manual_dt = ""
            note = ""
            if crane in override_map:
                manual_dt = override_map[crane]["swap_date"]
                note = override_map[crane]["note"]
            elif prev_sp != curr_sp:
                note = "대시보드 교체 감지 (교환일 미상 → 중간일 50:50 자동산정)"

            self.crane_data[crane] = {
                "prev_spreader": prev_sp,
                "curr_spreader": curr_sp,
                "work_count": DEFAULT_TLC_COUNTS_20260914_20261019.get(crane, 0),
                "manual_swap_date": manual_dt,
                "note": note,
            }

    # -----------------------------------------------------------------------
    # UI 구성
    # -----------------------------------------------------------------------
    def _build_ui(self):
        # 상단 제어 패널
        top_frame = ttk.LabelFrame(self, text="기본 설정 · 데이터 입력 · DB 저장 및 엑셀 연동", padding=(10, 6))
        top_frame.pack(fill=tk.X, padx=10, pady=(8, 4))

        ttk.Label(top_frame, text="기준 엑셀일자:").grid(row=0, column=0, sticky=tk.W, padx=(0, 3))
        ttk.Entry(top_frame, textvariable=self.base_date_var, width=11).grid(row=0, column=1, padx=(0, 10))

        ttk.Label(top_frame, text="산정 시작일:").grid(row=0, column=2, sticky=tk.W, padx=(0, 3))
        ttk.Entry(top_frame, textvariable=self.period_start_var, width=11).grid(row=0, column=3, padx=(0, 10))

        ttk.Label(top_frame, text="산정 종료일:").grid(row=0, column=4, sticky=tk.W, padx=(0, 3))
        ttk.Entry(top_frame, textvariable=self.period_end_var, width=11).grid(row=0, column=5, padx=(0, 12))

        ttk.Button(top_frame, text="관리 엑셀 불러오기", command=self.action_load_excel).grid(row=0, column=6, padx=3)
        ttk.Button(top_frame, text="대시보드 현황 붙여넣기", command=self.action_paste_dashboard).grid(row=0, column=7, padx=3)
        ttk.Button(top_frame, text="TLC(트위스트락) 붙여넣기", command=self.action_paste_work_counts).grid(row=0, column=8, padx=3)
        ttk.Button(top_frame, text="수동 교환이력 설정", command=self.action_open_manual_swap_dialog).grid(row=0, column=9, padx=3)
        ttk.Button(top_frame, text="💾 현재 산정내역 DB 저장", style="Accent.TButton", command=self.action_save_to_db).grid(row=0, column=10, padx=5)
        ttk.Button(top_frame, text="📊 결과 엑셀 저장", style="Accent.TButton", command=self.action_export_excel).grid(row=0, column=11, padx=3)

        # 요약 배너 프레임
        self.summary_var = tk.StringVar(value="")
        summary_bar = ttk.Frame(self, padding=(10, 4))
        summary_bar.pack(fill=tk.X, padx=10)
        ttk.Label(summary_bar, textvariable=self.summary_var, font=("맑은 고딕", 10, "bold"), foreground="#1a5fb4").pack(side=tk.LEFT)

        # 메인 탭 (Notebook)
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=(4, 10))

        # 탭 1: 호기별 대시보드 장착 현황 & 교환일 자동 산정 (38대)
        self.tab_cranes = ttk.Frame(self.notebook, padding=8)
        self.notebook.add(self.tab_cranes, text=" 1. 호기별 대시보드 현황 & 교환일/TLC 산정 (38대) ")
        self._build_tab_cranes()

        # 탭 2: Spreader별 최종 누적 Count 리스트 (관리 엑셀 양식)
        self.tab_spreaders = ttk.Frame(self.notebook, padding=8)
        self.notebook.add(self.tab_spreaders, text=" 2. Spreader 누적 Count 결과표 (44대) ")
        self._build_tab_spreaders()

        # 탭 3: 교체 감지 및 수동 입력 집중 관리
        self.tab_swaps = ttk.Frame(self.notebook, padding=8)
        self.notebook.add(self.tab_swaps, text=" 3. 스프레더 교체 호기 & 교환일 집중 관리 ")
        self._build_tab_swaps()

        # 탭 4: DB 이력 관리 및 기간별 Spreader 사용 카운트 조회 (신규 핵심 탭)
        self.tab_db_query = ttk.Frame(self.notebook, padding=8)
        self.notebook.add(self.tab_db_query, text=" 🔍 4. 기간별 Spreader 사용 카운트 조회 & DB 이력 관리 ")
        self._build_tab_db_query()

    def _build_tab_cranes(self):
        info_frame = ttk.Frame(self.tab_cranes)
        info_frame.pack(fill=tk.X, pady=(0, 6))
        ttk.Label(
            info_frame,
            text="💡 행을 더블클릭하면 해당 호기의 [이전 스프레더], [최근 스프레더], [기간 TLC(트위스트락횟수)], [수동 교환일]을 직접 수정할 수 있습니다.",
            style="Info.TLabel"
        ).pack(side=tk.LEFT)

        btn_frame = ttk.Frame(info_frame)
        btn_frame.pack(side=tk.RIGHT)
        ttk.Button(btn_frame, text="선택 호기 수정", command=self.action_edit_selected_crane).pack(side=tk.LEFT, padx=3)
        ttk.Button(btn_frame, text="다음 회차 준비 (현재 누적값→기준값 이관)", command=self.action_roll_forward_next_period).pack(side=tk.LEFT, padx=3)
        ttk.Button(btn_frame, text="TLC 기본값(9/14~10/19) 복원", command=self.action_restore_tlc_counts).pack(side=tk.LEFT, padx=3)

        cols = (
            "crane", "prev_sp", "curr_sp", "status",
            "work_count", "swap_date", "method",
            "prev_add", "curr_add", "note"
        )
        self.tree_cranes = ttk.Treeview(self.tab_cranes, columns=cols, show="headings", selectmode="browse")
        headers = {
            "crane": ("ARMGC 호기", 90),
            "prev_sp": ("이전 장착", 105),
            "curr_sp": ("최근 장착", 105),
            "status": ("변동 상태", 100),
            "work_count": ("기간 TLC (트위스트락)", 135),
            "swap_date": ("교환일 (자동/수동)", 140),
            "method": ("배분 산정 방식", 210),
            "prev_add": ("이전 Spreader 배분량", 140),
            "curr_add": ("최근 Spreader 배분량", 140),
            "note": ("비고 / 메모", 250),
        }
        for col, (text, width) in headers.items():
            self.tree_cranes.heading(col, text=text)
            align = tk.E if col in ("work_count", "prev_add", "curr_add") else tk.CENTER
            if col == "note":
                align = tk.W
            self.tree_cranes.column(col, width=width, anchor=align)

        self.tree_cranes.tag_configure("swapped_auto", background="#fff3cd")
        self.tree_cranes.tag_configure("swapped_manual", background="#d1ecf1")
        self.tree_cranes.tag_configure("normal", background="#ffffff")

        vsb = ttk.Scrollbar(self.tab_cranes, orient="vertical", command=self.tree_cranes.yview)
        self.tree_cranes.configure(yscrollcommand=vsb.set)
        self.tree_cranes.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

        self.tree_cranes.bind("<Double-1>", lambda e: self.action_edit_selected_crane())

    def _build_tab_spreaders(self):
        top_bar = ttk.Frame(self.tab_spreaders)
        top_bar.pack(fill=tk.X, pady=(0, 6))
        ttk.Label(
            top_bar,
            text="💡 기준 누적 Count에 기간 내 배분된 TLC(트위스트락횟수)를 합산한 결과입니다. (더블클릭 시 기준 누적값 수정 가능)",
            style="Info.TLabel"
        ).pack(side=tk.LEFT)

        self.sort_by_count_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            top_bar,
            text="누적 Count 내림차순 정렬 (엑셀 관리양식 기준)",
            variable=self.sort_by_count_var,
            command=self.recalculate_all
        ).pack(side=tk.RIGHT, padx=6)

        ttk.Button(top_bar, text="신규 Spreader 추가", command=self.action_add_spreader).pack(side=tk.RIGHT, padx=4)

        cols = (
            "rank", "spreader", "init_crane", "curr_crane",
            "base_count", "added_as_prev", "added_as_curr",
            "total_added", "final_count", "history"
        )
        self.tree_spreaders = ttk.Treeview(self.tab_spreaders, columns=cols, show="headings", selectmode="browse")
        headers = {
            "rank": ("순번(No.)", 75),
            "spreader": ("Spreader", 90),
            "init_crane": ("기준 호기", 110),
            "curr_crane": ("현재 호기", 115),
            "base_count": ("기준 누적 Count", 130),
            "added_as_prev": ("교체전 배분(+)", 115),
            "added_as_curr": ("현재장착 배분(+)", 120),
            "total_added": ("기간 총 증가량(+)", 125),
            "final_count": ("최종 누적 Count", 135),
            "history": ("기간 내 장착/교체 이력 상세", 320),
        }
        for col, (text, width) in headers.items():
            self.tree_spreaders.heading(col, text=text)
            align = tk.E if col in ("base_count", "added_as_prev", "added_as_curr", "total_added", "final_count") else tk.CENTER
            if col == "history":
                align = tk.W
            self.tree_spreaders.column(col, width=width, anchor=align)

        self.tree_spreaders.tag_configure("spare", foreground="#6c757d")
        self.tree_spreaders.tag_configure("changed", background="#e8f4fd")

        vsb = ttk.Scrollbar(self.tab_spreaders, orient="vertical", command=self.tree_spreaders.yview)
        self.tree_spreaders.configure(yscrollcommand=vsb.set)
        self.tree_spreaders.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

        self.tree_spreaders.bind("<Double-1>", lambda e: self.action_edit_selected_spreader())

    def _build_tab_swaps(self):
        desc = ttk.LabelFrame(self.tab_swaps, text="스프레더 교체 감지 및 교환일 산정 원리", padding=10)
        desc.pack(fill=tk.X, pady=(0, 8))
        ttk.Label(
            desc,
            text=(
                "1) 두 대시보드 사진(예: 2026-09-14 vs 2026-10-19)에서 호기별 부착된 Spreader 번호가 달라진 호기를 자동으로 추출합니다.\n"
                "2) 교환일을 확인할 수 없는 경우: 두 사진 날짜의 정중앙일(50:50 균등 배분)을 교환일로 자동 산정하여 이전/신규 Spreader에 절반씩 카운트를 나눕니다.\n"
                "3) 제미나이/수기로 확인 가능한 교환일이 있는 경우: 교환일을 입력하면 기간 내 일할 계산(장착일수 비율)으로 정밀 배분합니다."
            ),
            justify=tk.LEFT
        ).pack(anchor=tk.W)

        btn_bar = ttk.Frame(self.tab_swaps)
        btn_bar.pack(fill=tk.X, pady=(0, 6))
        ttk.Button(btn_bar, text="교환일 / 장착호기 수동 설정 (Y44 211호↔226호 전환 등)", command=self.action_open_manual_swap_dialog).pack(side=tk.LEFT, padx=4)

        cols = ("crane", "excel_831", "dash_914", "dash_1019", "swap_date", "method", "work_count", "allocation", "note")
        self.tree_swaps = ttk.Treeview(self.tab_swaps, columns=cols, show="headings", height=12)
        headers = {
            "crane": ("호기", 80),
            "excel_831": ("기준 호기", 95),
            "dash_914": ("이전 대시보드", 105),
            "dash_1019": ("최근 대시보드", 105),
            "swap_date": ("적용 교환일", 130),
            "method": ("산정 방식", 210),
            "work_count": ("호기 TLC", 120),
            "allocation": ("Spreader 배분 결과", 220),
            "note": ("비고", 260),
        }
        for col, (text, width) in headers.items():
            self.tree_swaps.heading(col, text=text)
            self.tree_swaps.column(col, width=width, anchor=tk.CENTER if col != "note" else tk.W)

        self.tree_swaps.pack(fill=tk.BOTH, expand=True)
        self.tree_swaps.bind("<Double-1>", lambda e: self.action_edit_selected_swap())

    def _build_tab_db_query(self):
        """
        탭 4:
        - 상단 좌측: DB에 저장된 산정 기간(회차) 목록 관리 (불러오기 / 삭제 / 메모 저장)
        - 상단 우측: 조회기간(시작일 ~ 종료일) 및 Spreader 번호 필터 설정
        - 하단: 조회기간 동안의 Spreader 번호별 사용 카운트(TLC) 집계 테이블 및 엑셀 내보내기
        """
        paned = ttk.PanedWindow(self.tab_db_query, orient=tk.VERTICAL)
        paned.pack(fill=tk.BOTH, expand=True)

        # 1) 상단 프레임: DB 저장 회차 목록
        top_section = ttk.LabelFrame(paned, text="📂 로컬 DB (spreader_history.db) 저장 회차 목록", padding=8)
        paned.add(top_section, weight=1)

        db_bar = ttk.Frame(top_section)
        db_bar.pack(fill=tk.X, pady=(0, 6))
        ttk.Label(db_bar, text="회차 메모:").pack(side=tk.LEFT, padx=(0, 4))
        ttk.Entry(db_bar, textvariable=self.period_memo_var, width=42).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(db_bar, text="💾 현재 작업내역 DB 저장/업데이트", style="Accent.TButton", command=self.action_save_to_db).pack(side=tk.LEFT, padx=3)
        ttk.Button(db_bar, text="📂 선택 회차를 작업화면으로 불러오기", command=self.action_load_selected_period_from_db).pack(side=tk.LEFT, padx=3)
        ttk.Button(db_bar, text="🗑️ 선택 회차 삭제", command=self.action_delete_selected_period_from_db).pack(side=tk.LEFT, padx=3)

        p_cols = ("id", "period_start", "period_end", "base_date", "total_tlc", "swapped_count", "memo", "created_at")
        self.tree_db_periods = ttk.Treeview(top_section, columns=p_cols, show="headings", height=5, selectmode="browse")
        p_headers = {
            "id": ("ID", 50),
            "period_start": ("산정 시작일", 110),
            "period_end": ("산정 종료일", 110),
            "base_date": ("기준 엑셀일", 105),
            "total_tlc": ("기간 총 TLC", 115),
            "swapped_count": ("교체 호기수", 95),
            "memo": ("회차 메모", 360),
            "created_at": ("저장 일시", 160),
        }
        for col, (text, width) in p_headers.items():
            self.tree_db_periods.heading(col, text=text)
            self.tree_db_periods.column(col, width=width, anchor=tk.CENTER if col != "memo" else tk.W)

        self.tree_db_periods.pack(fill=tk.BOTH, expand=True)
        self.tree_db_periods.bind("<<TreeviewSelect>>", self._on_select_db_period_row)

        # 2) 하단 프레임: 조회기간 설정 및 Spreader 번호별 사용 카운트 조회
        bottom_section = ttk.LabelFrame(paned, text="🔍 조회기간별 Spreader 사용 카운트(TLC) 상세 조회", padding=8)
        paned.add(bottom_section, weight=3)

        q_bar = ttk.Frame(bottom_section)
        q_bar.pack(fill=tk.X, pady=(0, 8))

        ttk.Label(q_bar, text="조회 시작일 (YYYY-MM-DD):", font=("맑은 고딕", 10, "bold")).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Entry(q_bar, textvariable=self.query_start_var, width=13).pack(side=tk.LEFT, padx=(0, 12))

        ttk.Label(q_bar, text="조회 종료일 (YYYY-MM-DD):", font=("맑은 고딕", 10, "bold")).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Entry(q_bar, textvariable=self.query_end_var, width=13).pack(side=tk.LEFT, padx=(0, 12))

        ttk.Label(q_bar, text="Spreader 번호 필터(예: Y44):").pack(side=tk.LEFT, padx=(0, 4))
        ent_sp_filter = ttk.Entry(q_bar, textvariable=self.query_search_sp_var, width=10)
        ent_sp_filter.pack(side=tk.LEFT, padx=(0, 8))
        ent_sp_filter.bind("<KeyRelease>", lambda e: self.action_run_period_query())

        ttk.Button(q_bar, text="🔍 기간 사용 카운트 조회", style="Accent.TButton", command=self.action_run_period_query).pack(side=tk.LEFT, padx=4)
        ttk.Button(q_bar, text="전체 기간으로 설정", command=self.action_set_full_query_range).pack(side=tk.LEFT, padx=4)
        ttk.Button(q_bar, text="📥 조회 결과 엑셀 내보내기", command=self.action_export_query_excel).pack(side=tk.RIGHT, padx=4)

        self.query_summary_var = tk.StringVar(value="조회 기간을 설정하고 [기간 사용 카운트 조회] 버튼을 누르세요.")
        ttk.Label(bottom_section, textvariable=self.query_summary_var, foreground="#0b5ed7", font=("맑은 고딕", 9, "bold")).pack(anchor=tk.W, pady=(0, 4))

        q_cols = ("rank", "spreader", "period_usage", "used_cranes", "swap_events", "latest_crane", "latest_total", "details")
        self.tree_query = ttk.Treeview(bottom_section, columns=q_cols, show="headings", selectmode="browse")
        q_headers = {
            "rank": ("순위", 60),
            "spreader": ("Spreader 번호", 110),
            "period_usage": ("조회기간 사용 Count (TLC)", 175),
            "used_cranes": ("기간 중 장착 호기", 130),
            "swap_events": ("기간 중 교체 이력", 190),
            "latest_crane": ("현재 장착호기", 110),
            "latest_total": ("현재 총 누적 Count", 145),
            "details": ("회차별 산정 상세 내역", 380),
        }
        for col, (text, width) in q_headers.items():
            self.tree_query.heading(col, text=text)
            align = tk.E if col in ("period_usage", "latest_total") else tk.CENTER
            if col in ("details", "swap_events"):
                align = tk.W
            self.tree_query.column(col, width=width, anchor=align)

        self.tree_query.tag_configure("active_use", background="#eef8ff")
        self.tree_query.tag_configure("zero_use", foreground="#888888")

        vsb_q = ttk.Scrollbar(bottom_section, orient="vertical", command=self.tree_query.yview)
        self.tree_query.configure(yscrollcommand=vsb_q.set)
        self.tree_query.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb_q.pack(side=tk.RIGHT, fill=tk.Y)

    # -----------------------------------------------------------------------
    # 핵심 산정 로직 수행 및 테이블 갱신
    # -----------------------------------------------------------------------
    def recalculate_all(self):
        try:
            p_start = parse_date(self.period_start_var.get())
            p_end = parse_date(self.period_end_var.get())
        except Exception as e:
            messagebox.showerror("날짜 오류", f"시작일/종료일 형식이 올바르지 않습니다 (YYYY-MM-DD): {e}")
            return

        self.calc_crane_results.clear()
        spreader_added_prev = {sp: 0 for sp in self.spreader_base}
        spreader_added_curr = {sp: 0 for sp in self.spreader_base}
        spreader_history = {sp: [] for sp in self.spreader_base}
        current_mounted_map = {}

        swapped_count = 0
        total_work_count = 0

        for crane in ARMGC_CRANES:
            info = self.crane_data[crane]
            prev_sp = info["prev_spreader"].strip()
            curr_sp = info["curr_spreader"].strip()
            wc = int(info["work_count"])
            manual_dt = info["manual_swap_date"].strip()
            total_work_count += wc

            for sp in (prev_sp, curr_sp):
                if sp and sp not in self.spreader_base:
                    self.spreader_base[sp] = {"init_crane": "Spare", "base_count": 0}
                    spreader_added_prev[sp] = 0
                    spreader_added_curr[sp] = 0
                    spreader_history[sp] = []

            alloc = calculate_swap_allocation(p_start, p_end, prev_sp, curr_sp, wc, manual_dt)
            self.calc_crane_results[crane] = alloc

            if alloc["is_swapped"]:
                swapped_count += 1
                if prev_sp:
                    spreader_added_prev[prev_sp] += alloc["prev_added"]
                    spreader_history[prev_sp].append(
                        f"{crane}호(교체전 {alloc['prev_ratio']*100:.0f}%: +{alloc['prev_added']:,})"
                    )
                if curr_sp:
                    spreader_added_curr[curr_sp] += alloc["curr_added"]
                    spreader_history[curr_sp].append(
                        f"{crane}호(교체후 {alloc['curr_ratio']*100:.0f}%: +{alloc['curr_added']:,}, {alloc['swap_date_str']})"
                    )
            else:
                if curr_sp:
                    spreader_added_curr[curr_sp] += alloc["curr_added"]
                    if wc > 0:
                        spreader_history[curr_sp].append(f"{crane}호(100%: +{wc:,})")
                    else:
                        spreader_history[curr_sp].append(f"{crane}호(계속 장착)")

            if curr_sp:
                current_mounted_map[curr_sp] = crane

        self.calc_spreader_results.clear()
        for sp_id, base_info in self.spreader_base.items():
            base_cnt = int(base_info["base_count"])
            init_cr = base_info["init_crane"]
            curr_cr = current_mounted_map.get(sp_id, "Spare")
            add_prev = spreader_added_prev.get(sp_id, 0)
            add_curr = spreader_added_curr.get(sp_id, 0)
            tot_add = add_prev + add_curr
            final_cnt = base_cnt + tot_add
            hist_str = ", ".join(spreader_history.get(sp_id, [])) if spreader_history.get(sp_id) else "예비(Spare) 대기"

            self.calc_spreader_results.append({
                "spreader": sp_id,
                "init_crane": init_cr,
                "curr_crane": curr_cr,
                "base_count": base_cnt,
                "added_as_prev": add_prev,
                "added_as_curr": add_curr,
                "total_added": tot_add,
                "final_count": final_cnt,
                "history": hist_str,
            })

        if self.sort_by_count_var.get():
            self.calc_spreader_results.sort(key=lambda x: (-x["final_count"], x["spreader"]))
        else:
            self.calc_spreader_results.sort(key=lambda x: x["spreader"])

        self._refresh_trees(swapped_count, total_work_count)

    def _refresh_trees(self, swapped_count: int, total_work_count: int):
        for item in self.tree_cranes.get_children():
            self.tree_cranes.delete(item)

        for crane in ARMGC_CRANES:
            info = self.crane_data[crane]
            alloc = self.calc_crane_results[crane]
            prev_sp = info["prev_spreader"]
            curr_sp = info["curr_spreader"]

            if prev_sp != curr_sp:
                status_str = f"교체 ({prev_sp}→{curr_sp})"
            elif info["manual_swap_date"].strip():
                status_str = "수동교환지정"
            else:
                status_str = "유지"

            tag = "normal"
            if alloc["is_swapped"]:
                tag = "swapped_manual" if info["manual_swap_date"].strip() else "swapped_auto"

            self.tree_cranes.insert(
                "", tk.END, iid=crane,
                values=(
                    crane,
                    prev_sp,
                    curr_sp,
                    status_str,
                    f"{info['work_count']:,}",
                    alloc["swap_date_str"],
                    alloc["method"],
                    f"+{alloc['prev_added']:,} ({prev_sp})" if alloc["is_swapped"] else "-",
                    f"+{alloc['curr_added']:,} ({curr_sp})",
                    info["note"]
                ),
                tags=(tag,)
            )

        for item in self.tree_spreaders.get_children():
            self.tree_spreaders.delete(item)

        for idx, row in enumerate(self.calc_spreader_results, start=1):
            tag = "spare" if row["curr_crane"] == "Spare" else ("changed" if row["init_crane"] != row["curr_crane"] else "")
            self.tree_spreaders.insert(
                "", tk.END, iid=row["spreader"],
                values=(
                    idx,
                    row["spreader"],
                    row["init_crane"],
                    row["curr_crane"],
                    f"{row['base_count']:,}",
                    f"+{row['added_as_prev']:,}" if row["added_as_prev"] else "0",
                    f"+{row['added_as_curr']:,}" if row["added_as_curr"] else "0",
                    f"+{row['total_added']:,}",
                    f"{row['final_count']:,}",
                    row["history"]
                ),
                tags=(tag,) if tag else ()
            )

        for item in self.tree_swaps.get_children():
            self.tree_swaps.delete(item)

        excel_831_crane_map = {}
        for sp_id, binfo in self.spreader_base.items():
            if binfo["init_crane"] != "Spare":
                excel_831_crane_map[binfo["init_crane"]] = sp_id

        for crane in ARMGC_CRANES:
            info = self.crane_data[crane]
            alloc = self.calc_crane_results[crane]
            sp_831 = excel_831_crane_map.get(crane, "-")
            prev_sp = info["prev_spreader"]
            curr_sp = info["curr_spreader"]

            if alloc["is_swapped"] or sp_831 != prev_sp or prev_sp != curr_sp:
                alloc_str = f"{prev_sp}(+{alloc['prev_added']:,}) / {curr_sp}(+{alloc['curr_added']:,})"
                self.tree_swaps.insert(
                    "", tk.END, iid=crane,
                    values=(
                        f"{crane}호",
                        sp_831,
                        prev_sp,
                        curr_sp,
                        alloc["swap_date_str"],
                        alloc["method"],
                        f"{info['work_count']:,}",
                        alloc_str,
                        info["note"]
                    )
                )

        avg_cnt = sum(r["final_count"] for r in self.calc_spreader_results) / max(len(self.calc_spreader_results), 1)
        self.summary_var.set(
            f"[산정 요약] ARMGC 38대 | 감지된 교체 호기: {swapped_count}대 | "
            f"기간 총 TLC(트위스트락): {total_work_count:,} | 전체 Spreader({len(self.calc_spreader_results)}대) 평균 Count: {avg_cnt:,.1f} | "
            f"DB 경로: {os.path.basename(DB_FILE_PATH)}"
        )

    # -----------------------------------------------------------------------
    # DB 저장 / 불러오기 / 조회기간 검색 액션
    # -----------------------------------------------------------------------
    def refresh_db_views(self):
        for item in self.tree_db_periods.get_children():
            self.tree_db_periods.delete(item)

        periods = self.db.list_periods()
        for p in periods:
            self.tree_db_periods.insert(
                "", tk.END, iid=str(p["id"]),
                values=(
                    p["id"],
                    p["period_start"],
                    p["period_end"],
                    p["base_date"],
                    f"{p['total_tlc']:,}",
                    f"{p['swapped_count']}대",
                    p["memo"],
                    p["created_at"]
                )
            )
        self.action_run_period_query()

    def _on_select_db_period_row(self, event=None):
        sel = self.tree_db_periods.selection()
        if not sel:
            return
        vals = self.tree_db_periods.item(sel[0], "values")
        if len(vals) >= 3:
            self.query_start_var.set(vals[1])
            self.query_end_var.set(vals[2])
            self.action_run_period_query()

    def action_save_to_db(self):
        self.recalculate_all()
        try:
            pid = self.db.save_period_calculation(
                self.base_date_var.get().strip(),
                self.period_start_var.get().strip(),
                self.period_end_var.get().strip(),
                self.period_memo_var.get().strip(),
                self.crane_data,
                self.calc_crane_results,
                self.calc_spreader_results,
            )
            self.query_start_var.set(self.period_start_var.get().strip())
            self.query_end_var.set(self.period_end_var.get().strip())
            self.refresh_db_views()
            messagebox.showinfo(
                "DB 저장 완료",
                f"산정 내역이 로컬 DB에 안전하게 저장되었습니다. (회차 ID: {pid})\n"
                f"기간: {self.period_start_var.get()} ~ {self.period_end_var.get()}\n"
                "[4. 기간별 Spreader 사용 카운트 조회] 탭에서 언제든 기간별 사용량을 조회하실 수 있습니다."
            )
        except Exception as e:
            messagebox.showerror("DB 저장 오류", f"DB 저장 중 오류가 발생했습니다:\n{e}")

    def action_load_selected_period_from_db(self):
        sel = self.tree_db_periods.selection()
        if not sel:
            messagebox.showinfo("알림", "불러올 회차를 목록에서 먼저 선택해주세요.")
            return
        pid = int(sel[0])
        data = self.db.load_period(pid)
        if not data:
            return

        pinfo = data["period"]
        self.base_date_var.set(pinfo["base_date"])
        self.period_start_var.set(pinfo["period_start"])
        self.period_end_var.set(pinfo["period_end"])
        self.period_memo_var.set(pinfo["memo"])

        for sp_id, sinfo in data["spreaders"].items():
            self.spreader_base[sp_id] = {
                "init_crane": sinfo["init_crane"],
                "base_count": sinfo["base_count"],
            }
        for cr, cinfo in data["cranes"].items():
            if cr in self.crane_data:
                self.crane_data[cr] = cinfo

        self.recalculate_all()
        messagebox.showinfo("회차 불러오기 완료", f"[{pinfo['period_start']} ~ {pinfo['period_end']}] 산정 데이터를 작업 화면으로 불러왔습니다.")

    def action_delete_selected_period_from_db(self):
        sel = self.tree_db_periods.selection()
        if not sel:
            messagebox.showinfo("알림", "삭제할 회차를 먼저 선택해주세요.")
            return
        pid = int(sel[0])
        if not messagebox.askyesno("회차 삭제 확인", f"선택한 회차(ID: {pid})를 DB에서 삭제하시겠습니까?"):
            return
        self.db.delete_period(pid)
        self.refresh_db_views()

    def action_set_full_query_range(self):
        periods = self.db.list_periods()
        if not periods:
            return
        min_start = min(p["period_start"] for p in periods)
        max_end = max(p["period_end"] for p in periods)
        self.query_start_var.set(min_start)
        self.query_end_var.set(max_end)
        self.action_run_period_query()

    def action_run_period_query(self):
        try:
            q_start = self.query_start_var.get().strip()
            q_end = self.query_end_var.get().strip()
            parse_date(q_start)
            parse_date(q_end)
        except Exception as e:
            messagebox.showerror("조회 날짜 오류", f"조회 시작일/종료일 형식이 올바르지 않습니다 (YYYY-MM-DD):\n{e}")
            return

        results = self.db.query_usage_by_date_range(q_start, q_end)
        sp_filter = self.query_search_sp_var.get().strip().upper()
        if sp_filter:
            results = [r for r in results if sp_filter in r["spreader"].upper() or sp_filter in r["used_cranes"].upper()]

        self.query_results_cache = results

        for item in self.tree_query.get_children():
            self.tree_query.delete(item)

        total_period_usage = 0
        active_sp_count = 0
        matched_periods = results[0]["matched_periods"] if results else 0

        for idx, r in enumerate(results, start=1):
            usage = r["period_usage"]
            total_period_usage += usage
            if usage > 0:
                active_sp_count += 1
            tag = "active_use" if usage > 0 else "zero_use"
            self.tree_query.insert(
                "", tk.END, iid=r["spreader"],
                values=(
                    idx,
                    r["spreader"],
                    f"+{usage:,}" if usage > 0 else "0",
                    r["used_cranes"],
                    r["swap_events"],
                    r["latest_crane"],
                    f"{r['latest_total']:,}",
                    r["details"]
                ),
                tags=(tag,)
            )

        self.query_summary_var.set(
            f"📊 조회기간 [{q_start} ~ {q_end}] (매칭된 DB 회차: {matched_periods}건) | "
            f"가동 Spreader: {active_sp_count}대 / 표시 {len(results)}대 | "
            f"조회기간 총 사용 Count(TLC) 합계: {total_period_usage:,}"
        )

    def action_roll_forward_next_period(self):
        """
        현재 산정된 최종 누적 Count와 10/19 장착 현황을
        다음 산정 회차의 '기준 누적 Count' 및 '이전 장착 현황'으로 원클릭 이관
        """
        if not messagebox.askyesno(
            "다음 회차 준비 (기준값 이관)",
            f"현재 산정 완료된 [{self.period_end_var.get()}] 기준 최종 누적 Count와 장착 현황을\n"
            "다음 회차의 [기준 누적값] 및 [이전 대시보드 현황]으로 이관하시겠습니까?\n"
            "(이관 전 현재 내역은 DB에 자동 저장됩니다.)"
        ):
            return

        # 먼저 현재 회차 DB 자동 저장
        self.db.save_period_calculation(
            self.base_date_var.get().strip(),
            self.period_start_var.get().strip(),
            self.period_end_var.get().strip(),
            self.period_memo_var.get().strip(),
            self.crane_data,
            self.calc_crane_results,
            self.calc_spreader_results,
        )

        prev_end = self.period_end_var.get().strip()
        self.base_date_var.set(prev_end)
        self.period_start_var.set(prev_end)
        try:
            next_end_dt = parse_date(prev_end) + datetime.timedelta(days=30)
            self.period_end_var.set(next_end_dt.isoformat())
        except Exception:
            pass

        for srow in self.calc_spreader_results:
            sp_id = srow["spreader"]
            self.spreader_base[sp_id] = {
                "init_crane": srow["curr_crane"],
                "base_count": srow["final_count"],
            }

        for cr in ARMGC_CRANES:
            curr_sp = self.crane_data[cr]["curr_spreader"]
            self.crane_data[cr] = {
                "prev_spreader": curr_sp,
                "curr_spreader": curr_sp,
                "work_count": 0,
                "manual_swap_date": "",
                "note": "",
            }

        self.period_memo_var.set(f"{self.period_start_var.get()} ~ {self.period_end_var.get()} 정기 산정")
        self.recalculate_all()
        self.refresh_db_views()
        messagebox.showinfo("다음 회차 준비 완료", "이전 회차가 DB에 저장되었으며, 새로운 산정 기간 입력 준비가 완료되었습니다.")

    # -----------------------------------------------------------------------
    # 대시보드 현황 일괄 붙여넣기 및 기존 다이얼로그 액션
    # -----------------------------------------------------------------------
    def action_paste_dashboard(self):
        dlg = tk.Toplevel(self)
        dlg.title("대시보드 장착 현황 텍스트 일괄 붙여넣기")
        dlg.geometry("620x500")
        dlg.transient(self)
        dlg.grab_set()

        ttk.Label(
            dlg,
            text=(
                "제미나이로 판독한 대시보드 호기별 Spreader 현황 텍스트를 붙여넣으세요.\n"
                "• 예시 형식: '211: Y17, 212: Y18, 213: Y28 ...' 또는 줄바꿈 형태 모두 자동 인식됩니다."
            ),
            justify=tk.LEFT
        ).pack(fill=tk.X, padx=12, pady=8)

        target_var = tk.StringVar(value="curr")
        rb_frame = ttk.Frame(dlg)
        rb_frame.pack(fill=tk.X, padx=12, pady=4)
        ttk.Radiobutton(rb_frame, text="최근 장착 현황(산정 종료일)에 반영", variable=target_var, value="curr").pack(side=tk.LEFT, padx=6)
        ttk.Radiobutton(rb_frame, text="이전 장착 현황(산정 시작일)에 반영", variable=target_var, value="prev").pack(side=tk.LEFT, padx=6)

        txt = tk.Text(dlg, font=("Consolas", 10), height=16)
        txt.pack(fill=tk.BOTH, expand=True, padx=12, pady=4)

        lines = [f"{cr}: {self.crane_data[cr]['curr_spreader']}" for cr in ARMGC_CRANES]
        txt.insert("1.0", "\n".join(lines))

        def on_apply():
            raw = txt.get("1.0", tk.END)
            parsed = parse_dashboard_text(raw)
            if not parsed:
                messagebox.showwarning("인식 오류", "호기(211~272)와 Spreader(Y01~Y99) 매핑을 찾지 못했습니다.")
                return
            field = "curr_spreader" if target_var.get() == "curr" else "prev_spreader"
            for cr, sp in parsed.items():
                self.crane_data[cr][field] = sp
            self.recalculate_all()
            messagebox.showinfo("반영 완료", f"총 {len(parsed)}개 호기의 대시보드 장착 현황이 업데이트되었습니다.")
            dlg.destroy()

        btn_f = ttk.Frame(dlg, padding=8)
        btn_f.pack(fill=tk.X)
        ttk.Button(btn_f, text="적용 및 교체 자동감지", style="Accent.TButton", command=on_apply).pack(side=tk.RIGHT, padx=6)
        ttk.Button(btn_f, text="취소", command=dlg.destroy).pack(side=tk.RIGHT)

    def action_load_excel(self):
        path = filedialog.askopenfilename(
            title="관리 엑셀 파일 선택",
            filetypes=[("Excel Files", "*.xlsx *.xlsm"), ("All Files", "*.*")]
        )
        if not path:
            return
        try:
            self.excel_path = path
            self._read_excel_file(path)
            self.recalculate_all()
            messagebox.showinfo("불러오기 완료", f"관리 엑셀 파일을 성공적으로 불러왔습니다.\n기준일자: {self.base_date_var.get()}")
        except Exception as e:
            messagebox.showerror("엑셀 로드 오류", f"파일을 읽는 중 오류가 발생했습니다:\n{e}")

    def action_paste_work_counts(self):
        dlg = tk.Toplevel(self)
        dlg.title("ARMGC 호기별 TLC(트위스트락횟수) 붙여넣기 / 일괄 입력")
        dlg.geometry("680x540")
        dlg.transient(self)
        dlg.grab_set()

        ttk.Label(
            dlg,
            text=(
                "엑셀이나 시스템 화면에서 복사한 ARMGC(211~272호) 데이터를 아래에 붙여넣으세요.\n"
                "• 지원 형식 1: [호기] [TLC 횟수] (예: '211  8169' 또는 'TC4211  8169')\n"
                "• 지원 형식 2: 업로드하신 표 전체 복사본 (Eqp No | Work Time | MMBF | MTTR | TLC | Work Count → TLC 열 자동 추출)"
            ),
            justify=tk.LEFT
        ).pack(fill=tk.X, padx=12, pady=8)

        txt = tk.Text(dlg, font=("Consolas", 10), height=18)
        txt.pack(fill=tk.BOTH, expand=True, padx=12, pady=4)

        lines = ["Eqp_No\tTLC_Count"]
        for cr in ARMGC_CRANES:
            lines.append(f"TC4{cr}\t{self.crane_data[cr]['work_count']}")
        txt.insert("1.0", "\n".join(lines))

        def on_apply():
            raw = txt.get("1.0", tk.END)
            parsed = parse_work_count_text(raw, use_tlc=True)
            if not parsed:
                messagebox.showwarning("인식된 데이터 없음", "211~272호기에 해당하는 TLC 숫자를 찾지 못했습니다.")
                return
            for cr, val in parsed.items():
                self.crane_data[cr]["work_count"] = val
            self.recalculate_all()
            messagebox.showinfo("적용 완료", f"총 {len(parsed)}개 호기의 TLC(트위스트락횟수)가 반영되어 자동 산정되었습니다.")
            dlg.destroy()

        btn_f = ttk.Frame(dlg, padding=8)
        btn_f.pack(fill=tk.X)
        ttk.Button(btn_f, text="적용 및 자동 산정", style="Accent.TButton", command=on_apply).pack(side=tk.RIGHT, padx=6)
        ttk.Button(btn_f, text="취소", command=dlg.destroy).pack(side=tk.RIGHT)

    def action_edit_selected_crane(self):
        sel = self.tree_cranes.selection()
        if not sel:
            messagebox.showinfo("알림", "수정할 호기 행을 먼저 선택해주세요.")
            return
        self._open_crane_edit_dialog(sel[0])

    def action_edit_selected_swap(self):
        sel = self.tree_swaps.selection()
        if not sel:
            return
        self._open_crane_edit_dialog(sel[0])

    def _open_crane_edit_dialog(self, crane: str):
        info = self.crane_data[crane]
        dlg = tk.Toplevel(self)
        dlg.title(f"{crane}호기 장착 Spreader 및 교환일/TLC 수정")
        dlg.geometry("460x340")
        dlg.transient(self)
        dlg.grab_set()

        f = ttk.Frame(dlg, padding=16)
        f.pack(fill=tk.BOTH, expand=True)

        ttk.Label(f, text=f"ARMGC {crane}호기 상세 설정", style="SubTitle.TLabel").grid(row=0, column=0, columnspan=2, sticky=tk.W, pady=(0, 12))

        ttk.Label(f, text="이전 장착 Spreader:").grid(row=1, column=0, sticky=tk.W, pady=5)
        v_prev = tk.StringVar(value=info["prev_spreader"])
        ttk.Entry(f, textvariable=v_prev, width=18).grid(row=1, column=1, sticky=tk.W, pady=5)

        ttk.Label(f, text="최근 장착 Spreader:").grid(row=2, column=0, sticky=tk.W, pady=5)
        v_curr = tk.StringVar(value=info["curr_spreader"])
        ttk.Entry(f, textvariable=v_curr, width=18).grid(row=2, column=1, sticky=tk.W, pady=5)

        ttk.Label(f, text="기간 TLC (트위스트락횟수):").grid(row=3, column=0, sticky=tk.W, pady=5)
        v_wc = tk.StringVar(value=str(info["work_count"]))
        ttk.Entry(f, textvariable=v_wc, width=18).grid(row=3, column=1, sticky=tk.W, pady=5)

        ttk.Label(f, text="수동 교환일 (YYYY-MM-DD):").grid(row=4, column=0, sticky=tk.W, pady=5)
        v_dt = tk.StringVar(value=info["manual_swap_date"])
        ttk.Entry(f, textvariable=v_dt, width=18).grid(row=4, column=1, sticky=tk.W, pady=5)
        ttk.Label(f, text="* 비워두면 교체 시 기간 중간일(50:50)로 자동 산정됩니다.", style="Info.TLabel", foreground="#555").grid(row=5, column=0, columnspan=2, sticky=tk.W, pady=(0, 6))

        ttk.Label(f, text="비고 / 메모:").grid(row=6, column=0, sticky=tk.W, pady=5)
        v_note = tk.StringVar(value=info["note"])
        ttk.Entry(f, textvariable=v_note, width=32).grid(row=6, column=1, sticky=tk.W, pady=5)

        def save():
            try:
                wc_val = int(float(v_wc.get().replace(",", "").strip() or "0"))
            except ValueError:
                messagebox.showerror("입력 오류", "TLC 횟수는 숫자로 입력해주세요.")
                return
            info["prev_spreader"] = v_prev.get().strip().upper()
            info["curr_spreader"] = v_curr.get().strip().upper()
            info["work_count"] = wc_val
            info["manual_swap_date"] = v_dt.get().strip()
            info["note"] = v_note.get().strip()
            self.recalculate_all()
            dlg.destroy()

        btn_f = ttk.Frame(f)
        btn_f.grid(row=7, column=0, columnspan=2, pady=14)
        ttk.Button(btn_f, text="저장 및 재계산", style="Accent.TButton", command=save).pack(side=tk.LEFT, padx=6)
        ttk.Button(btn_f, text="취소", command=dlg.destroy).pack(side=tk.LEFT, padx=6)

    def action_open_manual_swap_dialog(self):
        dlg = tk.Toplevel(self)
        dlg.title("확인된 교환 날짜 및 Spreader 호기 수동 등록")
        dlg.geometry("560x420")
        dlg.transient(self)
        dlg.grab_set()

        f = ttk.Frame(dlg, padding=14)
        f.pack(fill=tk.BOTH, expand=True)

        ttk.Label(
            f,
            text="제미나이나 수기로 확인된 Spreader 교환 날짜 및 장착 호기를 입력하세요.",
            style="SubTitle.TLabel"
        ).pack(anchor=tk.W, pady=(0, 8))

        form = ttk.LabelFrame(f, text="교환 이력 즉시 반영", padding=10)
        form.pack(fill=tk.X, pady=6)

        ttk.Label(form, text="ARMGC 호기 선택:").grid(row=0, column=0, sticky=tk.W, pady=4)
        cb_crane = ttk.Combobox(form, values=ARMGC_CRANES, width=10, state="readonly")
        cb_crane.set("226")
        cb_crane.grid(row=0, column=1, sticky=tk.W, pady=4, padx=6)

        ttk.Label(form, text="교체 후(신규 장착) Spreader:").grid(row=1, column=0, sticky=tk.W, pady=4)
        ent_curr = ttk.Entry(form, width=14)
        ent_curr.insert(0, "Y44")
        ent_curr.grid(row=1, column=1, sticky=tk.W, pady=4, padx=6)

        ttk.Label(form, text="장착(교환) 일자 (YYYY-MM-DD):").grid(row=2, column=0, sticky=tk.W, pady=4)
        ent_date = ttk.Entry(form, width=14)
        ent_date.insert(0, "2026-09-11")
        ent_date.grid(row=2, column=1, sticky=tk.W, pady=4, padx=6)

        def on_crane_select(event=None):
            cr = cb_crane.get()
            if cr in self.crane_data:
                ent_curr.delete(0, tk.END)
                ent_curr.insert(0, self.crane_data[cr]["curr_spreader"])
                ent_date.delete(0, tk.END)
                ent_date.insert(0, self.crane_data[cr]["manual_swap_date"] or "2026-09-11")

        cb_crane.bind("<<ComboboxSelected>>", on_crane_select)

        def apply_single():
            cr = cb_crane.get().strip()
            new_sp = ent_curr.get().strip().upper()
            dt_str = ent_date.get().strip()
            if cr not in self.crane_data:
                return
            self.crane_data[cr]["curr_spreader"] = new_sp
            self.crane_data[cr]["manual_swap_date"] = dt_str
            self.crane_data[cr]["note"] = f"사용자 지정 ({new_sp}, 교환일: {dt_str})"
            self.recalculate_all()
            messagebox.showinfo("반영 완료", f"{cr}호기 → {new_sp} (교환일: {dt_str}) 설정이 반영되었습니다.")

        ttk.Button(form, text="해당 호기에 반영", style="Accent.TButton", command=apply_single).grid(row=3, column=0, columnspan=2, pady=8)

        preset_frame = ttk.LabelFrame(f, text="빠른 프리셋 (Y44 장착 호기 확인용)", padding=10)
        preset_frame.pack(fill=tk.X, pady=8)
        ttk.Label(
            preset_frame,
            text=(
                "• 10/19 대시보드 사진 판독 시: 211호=Y17, 226호=Y44, 244호=Y43, 213호=Y28\n"
                "• 만약 Y44가 211호에 9/11 부착된 것으로 계산하려면 아래 버튼을 누르세요."
            ),
            justify=tk.LEFT
        ).pack(anchor=tk.W, pady=(0, 6))

        btn_row = ttk.Frame(preset_frame)
        btn_row.pack(fill=tk.X)

        def set_y44_on_211():
            self.crane_data["226"]["curr_spreader"] = "Y31"
            self.crane_data["226"]["manual_swap_date"] = ""
            self.crane_data["226"]["note"] = ""
            self.crane_data["211"]["prev_spreader"] = "Y17"
            self.crane_data["211"]["curr_spreader"] = "Y44"
            self.crane_data["211"]["manual_swap_date"] = "2026-09-11"
            self.crane_data["211"]["note"] = "사용자 지정: Y44 9/11 211호 부착"
            self.recalculate_all()
            messagebox.showinfo("프리셋 적용", "Y44를 211호기(9/11 부착)로 설정했습니다.")

        def set_y44_on_226():
            self.crane_data["211"]["curr_spreader"] = "Y17"
            self.crane_data["211"]["manual_swap_date"] = ""
            self.crane_data["211"]["note"] = ""
            self.crane_data["226"]["prev_spreader"] = "Y31"
            self.crane_data["226"]["curr_spreader"] = "Y44"
            self.crane_data["226"]["manual_swap_date"] = "2026-09-11"
            self.crane_data["226"]["note"] = "대시보드 226호 Y44 (9/11 부착)"
            self.recalculate_all()
            messagebox.showinfo("프리셋 적용", "Y44를 대시보드 사진 위치인 226호기(9/11 부착)로 설정했습니다.")

        ttk.Button(btn_row, text="Y44를 [211호기 9/11 부착]으로 적용", command=set_y44_on_211).pack(side=tk.LEFT, padx=4)
        ttk.Button(btn_row, text="Y44를 [226호기 9/11 부착]으로 적용 (대시보드 기준)", command=set_y44_on_226).pack(side=tk.LEFT, padx=4)

        ttk.Button(f, text="닫기", command=dlg.destroy).pack(side=tk.RIGHT, pady=6)

    def action_edit_selected_spreader(self):
        sel = self.tree_spreaders.selection()
        if not sel:
            return
        sp_id = sel[0]
        binfo = self.spreader_base[sp_id]
        new_cnt = simpledialog.askinteger(
            "기존 누적 Count 수정",
            f"Spreader [{sp_id}]의 기준일({self.base_date_var.get()}) 누적 Count를 입력하세요:",
            initialvalue=binfo["base_count"],
            parent=self
        )
        if new_cnt is not None:
            binfo["base_count"] = int(new_cnt)
            self.recalculate_all()

    def action_add_spreader(self):
        sp_id = simpledialog.askstring("신규 Spreader 등록", "추가할 Spreader 번호(예: Y45)를 입력하세요:", parent=self)
        if not sp_id:
            return
        sp_id = sp_id.strip().upper()
        if sp_id in self.spreader_base:
            messagebox.showwarning("중복", f"{sp_id}는 이미 등록되어 있습니다.")
            return
        cnt = simpledialog.askinteger("초기 Count", f"{sp_id}의 초기 누적 Count를 입력하세요:", initialvalue=0, parent=self)
        self.spreader_base[sp_id] = {"init_crane": "Spare", "base_count": int(cnt or 0)}
        self.recalculate_all()

    def action_restore_tlc_counts(self):
        for cr in ARMGC_CRANES:
            self.crane_data[cr]["work_count"] = DEFAULT_TLC_COUNTS_20260914_20261019.get(cr, 0)
        self.recalculate_all()
        messagebox.showinfo("TLC 복원 완료", "38대 호기에 2026/09/14~2026/10/19 TLC(트위스트락횟수) 데이터가 복원되었습니다.")

    # -----------------------------------------------------------------------
    # 엑셀 내보내기 (관리양식 시트 + 상세 산정 시트 + 기간별 조회 결과 엑셀)
    # -----------------------------------------------------------------------
    def action_export_excel(self):
        default_name = f"ARMGC_Spreader_Count_{self.period_end_var.get().replace('-', '')}.xlsx"
        save_path = filedialog.asksaveasfilename(
            title="산정 결과 엑셀 파일 저장",
            initialfile=default_name,
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")]
        )
        if not save_path:
            return

        try:
            self._write_output_excel(save_path)
            messagebox.showinfo(
                "엑셀 저장 완료",
                f"산정 결과가 아래 파일로 저장되었습니다:\n{save_path}\n\n"
                "• 시트 1: ARMGC Spreader Count List (기존 관리양식 유지)\n"
                "• 시트 2: 호기별 산정 및 교환이력 상세"
            )
        except Exception as e:
            messagebox.showerror("저장 오류", f"엑셀 저장 중 오류가 발생했습니다:\n{e}")

    def action_export_query_excel(self):
        if not self.query_results_cache:
            self.action_run_period_query()
        q_s = self.query_start_var.get().strip().replace("-", "")
        q_e = self.query_end_var.get().strip().replace("-", "")
        default_name = f"Spreader_기간별사용카운트_{q_s}_{q_e}.xlsx"
        save_path = filedialog.asksaveasfilename(
            title="기간별 Spreader 사용 카운트 조회결과 저장",
            initialfile=default_name,
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")]
        )
        if not save_path:
            return

        try:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "기간별_Spreader_사용카운트"

            font_title = Font(name="맑은 고딕", size=13, bold=True)
            font_header = Font(name="맑은 고딕", size=10, bold=True)
            font_cell = Font(name="맑은 고딕", size=10)
            align_center = Alignment(horizontal="center", vertical="center")
            align_right = Alignment(horizontal="right", vertical="center")
            thin_side = Side(border_style="thin", color="888888")
            border_box = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
            fill_header = PatternFill(start_color="EAEAEA", end_color="EAEAEA", fill_type="solid")

            ws["A1"] = f"Spreader 번호별 기간 사용 카운트(TLC) 조회 결과 ({self.query_start_var.get()} ~ {self.query_end_var.get()})"
            ws["A1"].font = font_title

            headers = ["순위", "Spreader 번호", "조회기간 사용 Count(TLC)", "기간 중 장착 호기", "기간 중 교체 이력", "현재 장착호기", "현재 총 누적 Count", "상세 산정 내역"]
            for col_idx, h_text in enumerate(headers, start=1):
                c = ws.cell(row=3, column=col_idx, value=h_text)
                c.font = font_header
                c.alignment = align_center
                c.fill = fill_header
                c.border = border_box

            for idx, r in enumerate(self.query_results_cache, start=1):
                r_num = 3 + idx
                row_vals = [
                    idx, r["spreader"], r["period_usage"], r["used_cranes"],
                    r["swap_events"], r["latest_crane"], r["latest_total"], r["details"]
                ]
                for c_idx, val in enumerate(row_vals, start=1):
                    cell = ws.cell(row=r_num, column=c_idx, value=val)
                    cell.font = font_cell
                    cell.border = border_box
                    if c_idx in (3, 7):
                        cell.alignment = align_right
                        cell.number_format = "#,##0"
                    elif c_idx in (5, 8):
                        cell.alignment = Alignment(horizontal="left", vertical="center")
                    else:
                        cell.alignment = align_center

            widths = [8, 15, 24, 18, 26, 15, 20, 48]
            for idx, w in enumerate(widths, start=1):
                ws.column_dimensions[openpyxl.utils.get_column_letter(idx)].width = w

            wb.save(save_path)
            messagebox.showinfo("저장 완료", f"조회기간 사용 카운트 내역이 저장되었습니다:\n{save_path}")
        except Exception as e:
            messagebox.showerror("저장 오류", f"엑셀 저장 중 오류가 발생했습니다:\n{e}")

    def _write_output_excel(self, save_path: str):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Sheet4 (2)"

        font_title = Font(name="맑은 고딕", size=14, bold=True)
        font_header = Font(name="맑은 고딕", size=10, bold=True)
        font_cell = Font(name="맑은 고딕", size=10)
        align_center = Alignment(horizontal="center", vertical="center")
        align_right = Alignment(horizontal="right", vertical="center")
        thin_side = Side(border_style="thin", color="888888")
        border_box = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
        fill_header = PatternFill(start_color="EAEAEA", end_color="EAEAEA", fill_type="solid")
        fill_changed = PatternFill(start_color="FFF3CD", end_color="FFF3CD", fill_type="solid")

        ws["A1"] = "ARMGC Spreader Count List"
        ws["A1"].font = font_title

        ws.merge_cells("G2:H2")
        try:
            ws["G2"] = datetime.datetime.strptime(self.period_end_var.get().strip(), "%Y-%m-%d")
            ws["G2"].number_format = "yyyy-mm-dd"
        except Exception:
            ws["G2"] = self.period_end_var.get().strip()
        ws["G2"].font = font_header
        ws["G2"].alignment = align_center

        headers = ["No.", "No.", "호기", "Count", "No.", "No.", "호기", "Count"]
        for col_idx, h_text in enumerate(headers, start=1):
            c = ws.cell(row=3, column=col_idx, value=h_text)
            c.font = font_header
            c.alignment = align_center
            c.fill = fill_header
            c.border = border_box

        total_sp = len(self.calc_spreader_results)
        half_rows = (total_sp + 1) // 2

        for i in range(half_rows):
            r_num = 4 + i
            left_item = self.calc_spreader_results[i]
            cr_val = int(left_item["curr_crane"]) if left_item["curr_crane"].isdigit() else left_item["curr_crane"]
            vals_left = [i + 1, left_item["spreader"], cr_val, left_item["final_count"]]
            for c_offset, val in enumerate(vals_left):
                cell = ws.cell(row=r_num, column=1 + c_offset, value=val)
                cell.font = font_cell
                cell.alignment = align_center
                cell.border = border_box
                if c_offset == 3:
                    cell.number_format = "#,##0"
                if left_item["total_added"] > 0 and left_item["added_as_prev"] > 0:
                    cell.fill = fill_changed

            right_idx = half_rows + i
            if right_idx < total_sp:
                right_item = self.calc_spreader_results[right_idx]
                cr_val_r = int(right_item["curr_crane"]) if right_item["curr_crane"].isdigit() else right_item["curr_crane"]
                vals_right = [right_idx + 1, right_item["spreader"], cr_val_r, right_item["final_count"]]
                for c_offset, val in enumerate(vals_right):
                    cell = ws.cell(row=r_num, column=5 + c_offset, value=val)
                    cell.font = font_cell
                    cell.alignment = align_center
                    cell.border = border_box
                    if c_offset == 3:
                        cell.number_format = "#,##0"
                    if right_item["total_added"] > 0 and right_item["added_as_prev"] > 0:
                        cell.fill = fill_changed

        avg_row = 4 + half_rows
        last_right_data_row = 3 + (total_sp - half_rows)
        if last_right_data_row < avg_row - 1:
            avg_row = last_right_data_row + 1

        ws.merge_cells(start_row=avg_row, start_column=5, end_row=avg_row, end_column=7)
        avg_label = ws.cell(row=avg_row, column=5, value="평균")
        avg_label.font = font_header
        avg_label.alignment = align_center
        avg_label.fill = fill_header
        for col_i in range(5, 8):
            ws.cell(row=avg_row, column=col_i).border = border_box

        avg_cell = ws.cell(
            row=avg_row, column=8,
            value=f"=AVERAGE(D4:D{3+half_rows},H4:H{3+(total_sp-half_rows)})"
        )
        avg_cell.font = font_header
        avg_cell.alignment = align_center
        avg_cell.number_format = "#,##0"
        avg_cell.border = border_box
        avg_cell.fill = fill_header

        col_widths = {"A": 6.5, "B": 11.0, "C": 10.0, "D": 12.5, "E": 6.5, "F": 11.0, "G": 10.0, "H": 12.5}
        for col_letter, width in col_widths.items():
            ws.column_dimensions[col_letter].width = width

        # 시트 2: 호기별 산정 및 교환이력 상세
        ws2 = wb.create_sheet(title="호기별_교환산정_상세")
        ws2["A1"] = f"ARMGC 38대 호기별 Spreader 교환일 산정 및 TLC(트위스트락횟수) 배분 내역 ({self.period_start_var.get()} ~ {self.period_end_var.get()})"
        ws2["A1"].font = font_title

        s2_headers = [
            "ARMGC 호기", "이전 Spreader", "최근 Spreader",
            "교체 여부", "기간 TLC(트위스트락횟수)", "산정/입력 교환일",
            "배분 산정 방식", "이전 Spreader 배분", "최근 Spreader 배분", "비고"
        ]
        for col_idx, text in enumerate(s2_headers, start=1):
            c = ws2.cell(row=3, column=col_idx, value=text)
            c.font = font_header
            c.fill = fill_header
            c.alignment = align_center
            c.border = border_box

        for r_idx, crane in enumerate(ARMGC_CRANES, start=4):
            info = self.crane_data[crane]
            alloc = self.calc_crane_results[crane]
            row_vals = [
                crane,
                info["prev_spreader"],
                info["curr_spreader"],
                "교체" if alloc["is_swapped"] else "유지",
                info["work_count"],
                alloc["swap_date_str"],
                alloc["method"],
                alloc["prev_added"],
                alloc["curr_added"],
                info["note"]
            ]
            for c_idx, val in enumerate(row_vals, start=1):
                cell = ws2.cell(row=r_idx, column=c_idx, value=val)
                cell.font = font_cell
                cell.border = border_box
                if c_idx in (5, 8, 9):
                    cell.alignment = align_right
                    cell.number_format = "#,##0"
                elif c_idx == 10:
                    cell.alignment = Alignment(horizontal="left", vertical="center")
                else:
                    cell.alignment = align_center
                if alloc["is_swapped"]:
                    cell.fill = fill_changed

        s2_widths = [12, 18, 18, 12, 22, 18, 28, 18, 18, 36]
        for idx, w in enumerate(s2_widths, start=1):
            ws2.column_dimensions[openpyxl.utils.get_column_letter(idx)].width = w

        wb.save(save_path)


if __name__ == "__main__":
    if "--export" in sys.argv:
        sys.stdout.reconfigure(encoding="utf-8")
        app = SpreaderCountApp()
        app.withdraw()
        out_file = os.path.join(get_app_dir(), f"ARMGC_Spreader_Count_{app.period_end_var.get().replace('-', '')}.xlsx")
        app._write_output_excel(out_file)
        print(f"DB Saved & Excel Exported: {out_file}")
        for i, r in enumerate(app.calc_spreader_results, 1):
            print(f"{i:2d} | {r['spreader']} | {r['init_crane']:5s} -> {r['curr_crane']:5s} | 기준:{r['base_count']:7,d} | +추가:{r['total_added']:5,d} | 최종:{r['final_count']:7,d} | {r['history']}")
        app.destroy()
    else:
        app = SpreaderCountApp()
        app.mainloop()
