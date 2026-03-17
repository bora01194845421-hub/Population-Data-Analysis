"""
Agent B — 수원시 행정동별 인구 정리
입력: 수원시 행정동별 연령별 인구 DataFrame
출력: 연도별 행정동 전체 목록 DataFrame + 구별 소계 DataFrame
"""
from __future__ import annotations

import pandas as pd
from pathlib import Path
from typing import Union
from agents.loader import load_file as _load_file_impl

GU_ORDER = ["장안구", "권선구", "팔달구", "영통구"]


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


def _is_dong(name: str) -> bool:
    import re
    return bool(re.search(r"동\d*$|[0-9]+동$", name)) or name.endswith("동")


def _to_int(val) -> int:
    try:
        return int(str(val).replace(",", ""))
    except (ValueError, TypeError):
        return 0


def run(files: list[Union[str, Path]]) -> dict[str, pd.DataFrame]:
    """
    Returns
    -------
    dict with keys:
      "detail"      : 연도별 행정동 전체 목록 DataFrame
      "gu_summary"  : 연도별 구별 소계 DataFrame
      "top5"        : 최신 연도 인구 상위 5 행정동
      "bottom5"     : 최신 연도 인구 하위 5 행정동
    """
    frames = [_load_file(f) for f in files]
    df = pd.concat(frames, ignore_index=True)

    # ── 컬럼 탐색 ──
    gu_col = _find_column(df, "자치구명", "구명", "시군구명", "자치구")
    dong_col = _find_column(df, "행정동명", "읍면동명", "동명", "행정동")
    combined_col = _find_column(df, "행정기관", "행정구역명", "행정구역")
    total_col = _find_column(df, "총인구수", "합계", "인구수", "총인구")
    male_col = _find_column(df, "남자인구수", "남자 인구수", "남", "남자")
    female_col = _find_column(df, "여자인구수", "여자 인구수", "여", "여자")
    household_col = _find_column(df, "세대수", "가구수", "세대")
    year_col = _find_column(df, "연도")

    use_combined = (gu_col is None or dong_col is None) and combined_col is not None

    if not use_combined and gu_col is None:
        raise ValueError(f"행정동 컬럼을 찾을 수 없습니다. 컬럼 목록: {list(df.columns)}")

    # ── 집계 ──
    rows = []
    for _, row in df.iterrows():
        if use_combined:
            parts = str(row[combined_col]).strip().split()
            if not parts or not _is_dong(parts[-1]):
                continue
            dong = parts[-1]
            gu = next((p for p in parts if p in GU_ORDER), "")
            if not gu:
                continue
        else:
            gu = str(row[gu_col]).strip()
            dong = str(row[dong_col]).strip()
            if not gu or gu in ("nan", "") or not dong or dong in ("nan", ""):
                continue

        year = None
        if year_col:
            try:
                year = int(row[year_col])
            except (ValueError, TypeError):
                pass

        total = _to_int(row[total_col]) if total_col else 0
        male = _to_int(row[male_col]) if male_col else 0
        female = _to_int(row[female_col]) if female_col else total - male
        household = _to_int(row[household_col]) if household_col else 0
        rows.append(
            {"연도": year, "구": gu, "행정동": dong, "총인구": total, "남": male, "여": female, "세대수": household}
        )

    detail = pd.DataFrame(rows)
    if detail.empty:
        empty = pd.DataFrame(columns=["연도", "구", "행정동", "총인구", "남", "여", "세대수", "세대당인구"])
        return {"detail": empty, "gu_summary": empty.drop(columns=["행정동", "세대당인구"], errors="ignore"), "top5": empty, "bottom5": empty}

    # 세대당 인구
    detail["세대당인구"] = (
        detail["총인구"] / detail["세대수"].replace(0, float("nan"))
    ).round(2)

    # ── 연도별 구별 소계 ──
    has_year = detail["연도"].notna().any()
    group_cols = ["연도", "구"] if has_year else ["구"]
    gu_summary = (
        detail.groupby(group_cols, as_index=False)
        .agg(총인구=("총인구", "sum"), 남=("남", "sum"), 여=("여", "sum"), 세대수=("세대수", "sum"))
    )

    if has_year:
        for yr, grp in gu_summary.groupby("연도"):
            total_pop = grp["총인구"].sum()
            gu_summary.loc[grp.index, "구비율(%)"] = (grp["총인구"] / total_pop * 100).round(1).values
    else:
        total_pop = detail["총인구"].sum()
        gu_summary["구비율(%)"] = (gu_summary["총인구"] / total_pop * 100).round(1)

    # 구 정렬
    gu_summary["_order"] = gu_summary["구"].map({g: i for i, g in enumerate(GU_ORDER)})
    sort_cols = (["연도", "_order"] if has_year else ["_order"])
    gu_summary = gu_summary.sort_values(sort_cols).drop(columns=["_order"]).reset_index(drop=True)

    detail["_order"] = detail["구"].map({g: i for i, g in enumerate(GU_ORDER)})
    detail_sort = (["연도", "_order", "총인구"] if has_year else ["_order", "총인구"])
    detail_asc = ([True, True, False] if has_year else [True, False])
    detail = detail.sort_values(detail_sort, ascending=detail_asc).drop(columns=["_order"]).reset_index(drop=True)

    # ── 최신 연도 기준 순위 ──
    if has_year:
        latest_yr = detail["연도"].max()
        rank_base = detail[detail["연도"] == latest_yr]
    else:
        rank_base = detail

    top5 = rank_base.nlargest(5, "총인구")[["연도", "구", "행정동", "총인구"]].reset_index(drop=True)
    bottom5 = rank_base.nsmallest(5, "총인구")[["연도", "구", "행정동", "총인구"]].reset_index(drop=True)

    return {
        "detail": detail,
        "gu_summary": gu_summary,
        "top5": top5,
        "bottom5": bottom5,
    }
