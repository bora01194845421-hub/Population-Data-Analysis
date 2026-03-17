"""
공통 파일 로더
주민등록인구 및 세대현황 형식(다중 연도 헤더) 자동 감지 및 전체 연도 추출
"""
from __future__ import annotations

import re
import pandas as pd
from pathlib import Path
from typing import Union


def load_file(path: Union[str, Path]) -> pd.DataFrame:
    """
    CSV/Excel 파일을 읽어 정규화된 DataFrame 반환.

    지원 형식:
    1. 표준형: 첫 행이 컬럼명
    2. 주민등록인구및세대현황 형식:
       Row 0 = 타이틀, Row 1 = 연도(다중), Row 2 = 컬럼명, Row 3+ = 데이터
       → 모든 연도 추출 후 '연도' 컬럼 추가
    3. KOSIS 연령별 형식: 행정동 + 연령 컬럼
    4. 단일 기간 스냅샷: 타이틀에 연도 포함, 단일 데이터 블록
    """
    p = Path(path)

    if p.suffix.lower() in (".xlsx", ".xls"):
        raw = pd.read_excel(p, header=None, dtype=str)
    else:
        raw = pd.read_csv(p, encoding="utf-8-sig", header=None, dtype=str)

    raw = raw.fillna("")

    # ── 헤더 행 탐색 ──
    header_row_idx = _find_header_row(raw)
    if header_row_idx is None:
        return _read_standard(p)

    # ── 헤더 바로 앞 행에서 연도 행 탐색 ──
    year_row_idx = None
    if header_row_idx > 0:
        prev = raw.iloc[header_row_idx - 1]
        if any("년" in _clean(v) for v in prev):
            year_row_idx = header_row_idx - 1

    col_names = [_clean(v) for v in raw.iloc[header_row_idx].tolist()]
    data = raw.iloc[header_row_idx + 1 :].copy().reset_index(drop=True)
    data = data[data.apply(lambda r: any(_clean(v) for v in r), axis=1)]

    if year_row_idx is not None:
        # 다중 연도 형식 → 전체 연도 추출 (각 연도에 '연도' 컬럼 추가)
        year_row_vals = [_clean(v) for v in raw.iloc[year_row_idx].tolist()]
        data = _extract_all_years(data, year_row_vals, col_names)
    else:
        data.columns = [c if c else f"col_{i}" for i, c in enumerate(col_names)]

    # ── 성별 섹션(합/남/여) 형식 → 중복 컬럼이 생기면 첫 번째(합계) 섹션만 유지 ──
    non_yr_cols = [c for c in data.columns if c != "연도"]
    if len(set(non_yr_cols)) < len(non_yr_cols):
        seen: set[str] = set()
        keep = []
        for i, c in enumerate(data.columns):
            if c == "연도" or c not in seen:
                seen.add(c)
                keep.append(i)
        data = data.iloc[:, keep]

    return data


def _clean(v) -> str:
    """openpyxl이 셀 값에 붙이는 따옴표를 제거하고 공백 정리"""
    return str(v).strip("' \t")


def _find_header_row(raw: pd.DataFrame) -> int | None:
    """행정구역/행정동 키워드가 있는 행 인덱스 반환"""
    keywords = ["행정구역", "행정동", "행정기관", "지역코드", "읍면동", "자치구", "시군구"]
    for i in range(min(6, len(raw))):
        row_str = " ".join(_clean(v) for v in raw.iloc[i].tolist())
        if any(k in row_str for k in keywords):
            return i
    return None


def _extract_all_years(
    data: pd.DataFrame, year_row: list, col_names: list
) -> pd.DataFrame:
    """연도 행을 분석해 모든 연도의 컬럼을 추출하고 '연도' 컬럼 추가"""

    year_positions: list[int] = []
    year_labels: list[str] = []
    for j, v in enumerate(year_row):
        cleaned = str(v).strip("' ")
        if "년" in cleaned and cleaned != "":
            year_positions.append(j)
            year_labels.append(cleaned)

    if not year_positions:
        data.columns = [str(c) if c != "" else f"col_{i}" for i, c in enumerate(col_names)]
        return data

    if len(year_positions) > 1:
        cols_per_year = year_positions[1] - year_positions[0]
    else:
        cols_per_year = data.shape[1] - year_positions[0]

    fixed_count = year_positions[0]
    fixed_indices = list(range(fixed_count))

    frames = []
    for yr_start, yr_label in zip(year_positions, year_labels):
        yr_end = min(yr_start + cols_per_year, data.shape[1])
        yr_indices = list(range(yr_start, yr_end))
        selected = fixed_indices + yr_indices

        frame = data.iloc[:, selected].copy()
        new_cols = []
        for idx in selected:
            name = str(col_names[idx]).strip()
            new_cols.append(name if name else f"col_{idx}")
        frame.columns = new_cols

        m = re.search(r"\d{4}", yr_label)
        frame["연도"] = int(m.group()) if m else yr_label
        frames.append(frame)

    return pd.concat(frames, ignore_index=True)


def _read_standard(p: Path) -> pd.DataFrame:
    if p.suffix.lower() in (".xlsx", ".xls"):
        return pd.read_excel(p, header=0)
    return pd.read_csv(p, encoding="utf-8-sig", header=0)
