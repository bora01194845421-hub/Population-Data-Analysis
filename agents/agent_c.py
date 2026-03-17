"""
Agent C — 수원시 생애주기 인구 분석
입력: 수원시 연령별 인구 DataFrame (행정동별 포함 가능)
출력: 생애주기별 집계, 부양비 지표, 행정동×생애주기 피벗
"""
from __future__ import annotations

import re
import pandas as pd
from pathlib import Path
from typing import Union
from agents.loader import load_file as _load_file_impl

GU_ORDER = ["장안구", "권선구", "팔달구", "영통구"]

# ── 생애주기 분류 (7단계, 노년기 65세 이상 통합) ──
LIFECYCLE_STAGES = [
    ("영유아", range(0, 7)),       # 0~6세
    ("아동", range(7, 13)),        # 7~12세
    ("청소년", range(13, 19)),     # 13~18세
    ("청년", range(19, 35)),       # 19~34세
    ("중장년", range(35, 50)),     # 35~49세
    ("장년", range(50, 65)),       # 50~64세
    ("노년기", range(65, 101)),    # 65세 이상
]

STAGE_NAMES = [s[0] for s in LIFECYCLE_STAGES]


def _load_file(path: Union[str, Path]) -> pd.DataFrame:
    return _load_file_impl(path)


def _detect_year(path: Union[str, Path]) -> int | None:
    """파일명 또는 파일 첫 5행에서 4자리 연도(20xx) 추출 — 연도 컬럼 미생성 시 fallback"""
    p = Path(path)
    # 1) 파일명에서
    m = re.search(r"(20\d{2})", p.stem)
    if m:
        return int(m.group(1))
    # 2) 파일 첫 5행 내용에서
    try:
        if p.suffix.lower() in (".xlsx", ".xls"):
            raw = pd.read_excel(p, header=None, nrows=5, dtype=str)
        else:
            raw = pd.read_csv(p, encoding="utf-8-sig", header=None, nrows=5, dtype=str)
        raw = raw.fillna("")
        for _, row in raw.iterrows():
            for v in row:
                m2 = re.search(r"(20\d{2})", str(v))
                if m2:
                    return int(m2.group(1))
    except Exception:
        pass
    return None


def _find_column(df: pd.DataFrame, *candidates: str) -> str | None:
    for c in candidates:
        if c in df.columns:
            return c
        matches = [col for col in df.columns if c in col]
        if matches:
            return matches[0]
    return None


def _to_int(val) -> int:
    try:
        return int(str(val).replace(",", ""))
    except (ValueError, TypeError):
        return 0


def _parse_gu_dong(name: str) -> tuple[str, str]:
    """'경기도 수원시 장안구 조원동' 형식에서 (구, 동) 추출"""
    parts = name.strip().split()
    dong = ""
    gu = ""
    if parts and re.search(r"동\d*$|[0-9]+동$", parts[-1]):
        dong = parts[-1]
    for p in parts:
        if p in GU_ORDER:
            gu = p
            break
    return gu, dong


def _extract_age_cols(df: pd.DataFrame) -> dict[int, str]:
    """컬럼명에서 연령(정수) → 컬럼명 매핑 반환"""
    age_map: dict[int, str] = {}
    for col in df.columns:
        m = re.search(r"(\d+)\s*세", str(col))
        if m:
            age = int(m.group(1))
            if age not in age_map:
                age_map[age] = col
    return age_map


def _lifecycle_pop(row: pd.Series, age_map: dict[int, str]) -> dict[str, int]:
    result = {}
    for name, ages in LIFECYCLE_STAGES:
        total = sum(_to_int(row.get(age_map[a], 0)) for a in ages if a in age_map)
        result[name] = total
    return result


def _compute(df: pd.DataFrame, age_map: dict[int, str]) -> dict[str, pd.DataFrame]:
    """단일 연도(또는 전체) 데이터에 대한 생애주기 분석 수행"""

    # ── combined 컬럼 여부에 따라 집계 대상 행 결정 ──
    combined_col_pre = _find_column(df, "행정기관", "행정구역명", "행정구역")
    if combined_col_pre:
        def _is_dong_row(val: str) -> bool:
            parts = val.strip().split()
            return bool(parts) and bool(re.search(r"동\d*$|[0-9]+동$", parts[-1]))
        df_agg = df[df[combined_col_pre].astype(str).apply(_is_dong_row)].copy()
    else:
        df_agg = df

    # ── 전체 생애주기 집계 ──
    lifecycle_totals = {name: 0 for name in STAGE_NAMES}
    for _, row in df_agg.iterrows():
        lc = _lifecycle_pop(row, age_map)
        for k, v in lc.items():
            lifecycle_totals[k] += v

    grand_total = sum(lifecycle_totals.values())

    total_col = _find_column(df_agg, "총인구수", "합계", "인구수", "총인구")
    if total_col:
        grand_total_raw = df_agg[total_col].apply(_to_int).sum()
        if grand_total_raw > 0:
            grand_total = grand_total_raw

    lifecycle_rows = []
    for name, ages in LIFECYCLE_STAGES:
        pop = lifecycle_totals[name]
        ratio = round(pop / grand_total * 100, 1) if grand_total else 0.0
        lifecycle_rows.append({"생애주기": name, "연령구간": f"{ages.start}~{ages.stop - 1}세", "인구": pop, "비율(%)": ratio})

    lifecycle_df = pd.DataFrame(lifecycle_rows)
    lifecycle_df.loc[len(lifecycle_df)] = {"생애주기": "합계", "연령구간": "", "인구": grand_total, "비율(%)": 100.0}

    # ── 부양비 지표 ──
    def _raw_age_sum(start: int, stop_inclusive: int) -> int:
        return sum(df_agg[age_map[a]].apply(_to_int).sum() for a in range(start, stop_inclusive + 1) if a in age_map)

    young = _raw_age_sum(0, 14)
    working = _raw_age_sum(15, 64)
    old = _raw_age_sum(65, 100)

    if working == 0:
        dep_young = dep_old = dep_total = aging_idx = float("nan")
    else:
        dep_young = round(young / working * 100, 1)
        dep_old = round(old / working * 100, 1)
        dep_total = round(dep_young + dep_old, 1)
        aging_idx = round(old / young * 100, 1) if young > 0 else float("nan")

    dependency_df = pd.DataFrame([
        {"지표": "유소년부양비", "수원시": dep_young},
        {"지표": "노년부양비", "수원시": dep_old},
        {"지표": "총부양비", "수원시": dep_total},
        {"지표": "고령화지수", "수원시": aging_idx},
    ])

    dong_col = _find_column(df, "행정동명", "읍면동명", "동명", "행정동")
    gu_col = _find_column(df, "자치구명", "구명", "자치구")
    combined_col = _find_column(df, "행정기관", "행정구역명", "행정구역")

    def _get_gu_dong(row) -> tuple[str, str]:
        if gu_col and dong_col:
            return str(row[gu_col]).strip(), str(row[dong_col]).strip()
        if combined_col:
            return _parse_gu_dong(str(row[combined_col]))
        return "", ""

    # ── 구별 생애주기 피벗 ──
    gu_pivot = pd.DataFrame()
    gu_count = pd.DataFrame()
    if gu_col or combined_col:
        gu_agg: dict[str, dict[str, int]] = {}
        for _, row in df.iterrows():
            gu, dong = _get_gu_dong(row)
            if combined_col and not gu_col:
                if not dong:
                    continue
            if not gu or gu == "nan" or gu not in GU_ORDER:
                continue
            if gu not in gu_agg:
                gu_agg[gu] = {name: 0 for name in STAGE_NAMES}
            lc = _lifecycle_pop(row, age_map)
            for k, v in lc.items():
                gu_agg[gu][k] += v
        if gu_agg:
            gu_rows = []
            gu_count_rows = []
            for gu_name in GU_ORDER:
                if gu_name not in gu_agg:
                    continue
                totals = gu_agg[gu_name]
                pop_total = sum(totals.values())
                ratio_row = {"구": gu_name}
                count_row = {"구": gu_name, "합계": pop_total}
                for name in STAGE_NAMES:
                    ratio_row[name] = round(totals[name] / pop_total * 100, 1) if pop_total else 0.0
                    count_row[name] = totals[name]
                gu_rows.append(ratio_row)
                gu_count_rows.append(count_row)
            gu_pivot = pd.DataFrame(gu_rows)
            gu_count = pd.DataFrame(gu_count_rows)

    # ── 행정동별 생애주기 피벗 ──
    dong_pivot = pd.DataFrame()
    dong_count = pd.DataFrame()
    has_dong_info = dong_col or combined_col
    if has_dong_info:
        pivot_rows = []
        count_rows = []
        for _, row in df.iterrows():
            gu, dong = _get_gu_dong(row)
            if not dong or dong == "nan":
                continue
            lc = _lifecycle_pop(row, age_map)
            pop_total = sum(lc.values())
            if pop_total == 0:
                continue
            ratio_row = {"구": gu, "행정동": dong}
            count_row = {"구": gu, "행정동": dong, "합계": pop_total}
            for name in STAGE_NAMES:
                ratio_row[name] = round(lc[name] / pop_total * 100, 1)
                count_row[name] = lc[name]
            pivot_rows.append(ratio_row)
            count_rows.append(count_row)
        if pivot_rows:
            dong_pivot = pd.DataFrame(pivot_rows)
            dong_count = pd.DataFrame(count_rows)

    return {
        "lifecycle": lifecycle_df,
        "dependency": dependency_df,
        "gu_pivot": gu_pivot,
        "gu_count": gu_count,
        "dong_pivot": dong_pivot,
        "dong_count": dong_count,
    }


def run(
    files: list[Union[str, Path]],
    agent_b_detail: pd.DataFrame | None = None,
) -> dict[str, pd.DataFrame]:
    """
    Parameters
    ----------
    files         : 수원시 연령별 인구 파일 경로 목록
    agent_b_detail: Agent B의 detail DataFrame (사용 예약)

    Returns
    -------
    dict with keys:
      "lifecycle"   : 생애주기별 인구 요약 (연도 컬럼 포함 가능)
      "dependency"  : 부양비 지표
      "gu_pivot"    : 구별 생애주기 비율 피벗
      "gu_count"    : 구별 생애주기 인원 피벗
      "dong_pivot"  : 행정동별 생애주기 비율 피벗
      "dong_count"  : 행정동별 생애주기 인원 피벗
    """
    # 파일별 로드 + 연도 컬럼 없으면 파일명/내용에서 fallback 추출
    frames = []
    for f in files:
        df_f = _load_file(f)
        if "연도" not in df_f.columns:
            yr = _detect_year(f)
            if yr:
                df_f = df_f.copy()
                df_f["연도"] = yr
        frames.append(df_f)
    df = pd.concat(frames, ignore_index=True)

    year_col = _find_column(df, "연도")
    # 연도 컬럼을 정수로 정규화 (2026.0 등 float 방지)
    if year_col:
        def _to_yr_int(x):
            try:
                return int(float(str(x)))
            except Exception:
                return x
        df[year_col] = df[year_col].apply(_to_yr_int)

    age_map = _extract_age_cols(df)

    if not age_map:
        empty = pd.DataFrame()
        msg = pd.DataFrame([{"안내": "생애주기 분석에는 1세 단위 연령별 인구 파일(KOSIS 연령별 인구)이 필요합니다."}])
        return {"lifecycle": msg, "dependency": empty, "gu_pivot": empty,
                "gu_count": empty, "dong_pivot": empty, "dong_count": empty}

    # ── 연도별 분리 처리 ──
    has_multi_year = (
        year_col is not None
        and df[year_col].notna().any()
        and df[year_col].nunique() > 1
    )

    if has_multi_year:
        years = sorted(df[year_col].dropna().unique(), key=lambda x: int(float(str(x))))
        all_lc, all_dep, all_gu_p, all_gu_c, all_dp, all_dc = [], [], [], [], [], []

        for yr in years:
            yr_df = df[df[year_col] == yr].copy()
            yr_int = int(float(str(yr)))
            result = _compute(yr_df, age_map)
            for key, lst in [
                ("lifecycle", all_lc), ("dependency", all_dep),
                ("gu_pivot", all_gu_p), ("gu_count", all_gu_c),
                ("dong_pivot", all_dp), ("dong_count", all_dc),
            ]:
                part = result[key]
                if not part.empty:
                    part = part.copy()
                    part.insert(0, "연도", yr_int)
                    lst.append(part)

        def _cat(lst: list) -> pd.DataFrame:
            valid = [x for x in lst if not x.empty]
            return pd.concat(valid, ignore_index=True) if valid else pd.DataFrame()

        return {
            "lifecycle": _cat(all_lc),
            "dependency": _cat(all_dep),
            "gu_pivot": _cat(all_gu_p),
            "gu_count": _cat(all_gu_c),
            "dong_pivot": _cat(all_dp),
            "dong_count": _cat(all_dc),
        }

    # ── 단일 연도 처리 ──
    return _compute(df, age_map)
