"""
Agent A — 전국 / 경기도 / 5개 시 인구 정리
입력: 시군구 단위 연령별 인구 DataFrame (또는 파일 경로 목록)
출력: 연도별 7개 지역 비교 DataFrame
"""
from __future__ import annotations

import re
import pandas as pd
from pathlib import Path
from typing import Union
from agents.loader import load_file as _load_file_impl

# ───────── 대상 지역 키워드 ─────────
TARGET_REGIONS = {
    "전국": "전국",
    "경기도": "경기도",
    "수원시": "수원시",
    "용인시": "용인시",
    "성남시": "성남시",
    "고양시": "고양시",
    "화성시": "화성시",
}

REGION_ORDER = ["전국", "경기도", "수원시", "용인시", "성남시", "고양시", "화성시"]


def _load_file(path: Union[str, Path]) -> pd.DataFrame:
    return _load_file_impl(path)


def _find_column(df: pd.DataFrame, *candidates: str) -> str | None:
    for c in candidates:
        if c in df.columns:
            return c
        matches = [col for col in df.columns if c in col]
        if matches:
            return matches[0]
    return None


def _to_int(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series.astype(str).str.replace(",", ""), errors="coerce").fillna(0).astype(int)


def _detect_region(row: pd.Series, region_col: str) -> str | None:
    val = str(row[region_col]).strip()
    parts = val.split()
    last = parts[-1] if parts else val
    for key in TARGET_REGIONS:
        if key == last or key in last:
            return key
    return None


def run(
    files: list[Union[str, Path]],
    prev_files: list[Union[str, Path]] | None = None,
) -> pd.DataFrame:
    frames = [_load_file(f) for f in files]
    df = pd.concat(frames, ignore_index=True)

    region_col = _find_column(df, "행정구역명", "행정기관", "행정구역(시군구)별", "행정구역", "시군구")
    total_col = _find_column(df, "총인구수", "합계", "인구수", "총인구")
    male_col = _find_column(df, "남자인구수", "남", "남자")
    female_col = _find_column(df, "여자인구수", "여", "여자")
    year_col = _find_column(df, "연도")

    if region_col is None:
        raise ValueError(f"지역 컬럼을 찾을 수 없습니다. 컬럼 목록: {list(df.columns)}")

    # ── 대상 지역 필터 & 집계 ──
    rows = []
    for _, row in df.iterrows():
        region = _detect_region(row, region_col)
        if region is None:
            continue
        year = None
        if year_col:
            try:
                year = int(row[year_col])
            except (ValueError, TypeError):
                pass
        total = _to_int(row[[total_col]]).iloc[0] if total_col else 0
        male = _to_int(row[[male_col]]).iloc[0] if male_col else 0
        female = _to_int(row[[female_col]]).iloc[0] if female_col else total - male
        rows.append({"연도": year, "지역": region, "총인구": total, "남자": male, "여자": female})

    if not rows:
        return pd.DataFrame(columns=["연도", "지역", "총인구", "남자", "여자", "성비", "전국비율(%)", "경기비율(%)", "전년대비(%)"])

    has_year = any(r["연도"] is not None for r in rows)
    group_cols = ["연도", "지역"] if has_year else ["지역"]

    result = (
        pd.DataFrame(rows)
        .groupby(group_cols, as_index=False)
        .agg(총인구=("총인구", "sum"), 남자=("남자", "sum"), 여자=("여자", "sum"))
    )

    result["성비"] = (result["남자"] / result["여자"].replace(0, float("nan")) * 100).round(1)

    # ── 연도별 비율 계산 ──
    if has_year:
        nat_by_year = result[result["지역"] == "전국"].set_index("연도")["총인구"].to_dict()
        gyeonggi_by_year = result[result["지역"] == "경기도"].set_index("연도")["총인구"].to_dict()
        result["전국비율(%)"] = result.apply(
            lambda r: round(r["총인구"] / nat_by_year[r["연도"]] * 100, 2)
            if r["연도"] in nat_by_year and nat_by_year[r["연도"]] > 0 else float("nan"), axis=1
        )
        result["경기비율(%)"] = result.apply(
            lambda r: round(r["총인구"] / gyeonggi_by_year[r["연도"]] * 100, 2)
            if r["연도"] in gyeonggi_by_year and gyeonggi_by_year[r["연도"]] > 0 else float("nan"), axis=1
        )
    else:
        nat_pop = result.loc[result["지역"] == "전국", "총인구"].sum()
        gyeonggi_pop = result.loc[result["지역"] == "경기도", "총인구"].sum()
        result["전국비율(%)"] = (result["총인구"] / nat_pop * 100).round(2) if nat_pop else float("nan")
        result["경기비율(%)"] = (result["총인구"] / gyeonggi_pop * 100).round(2) if gyeonggi_pop else float("nan")

    # ── 전년 대비 ──
    if has_year:
        result = result.sort_values(["지역", "연도"]).reset_index(drop=True)
        result["전년대비(%)"] = result.groupby("지역")["총인구"].pct_change().mul(100).round(2)
    else:
        result["전년대비(%)"] = float("nan")
        if prev_files:
            prev_result = run(prev_files)
            prev_map = dict(zip(prev_result["지역"], prev_result["총인구"]))
            def _change(row):
                prev = prev_map.get(row["지역"])
                if prev and prev > 0:
                    return round((row["총인구"] - prev) / prev * 100, 2)
                return float("nan")
            result["전년대비(%)"] = result.apply(_change, axis=1)

    # ── 순서 정렬 ──
    result["_order"] = result["지역"].map({r: i for i, r in enumerate(REGION_ORDER)})
    if has_year:
        result = result.sort_values(["연도", "_order"]).drop(columns=["_order"]).reset_index(drop=True)
    else:
        result = result.sort_values("_order").drop(columns=["_order"]).reset_index(drop=True)

    return result
