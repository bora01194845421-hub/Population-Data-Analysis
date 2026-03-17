"""
Markdown 리포트 생성 모듈
"""
from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Union

import pandas as pd


def _df_to_md(df: pd.DataFrame) -> str:
    if df.empty:
        return "_데이터 없음_\n"
    cols = list(df.columns)
    header = "| " + " | ".join(str(c) for c in cols) + " |"
    sep = "| " + " | ".join("---" for _ in cols) + " |"
    rows = []
    for _, row in df.iterrows():
        rows.append("| " + " | ".join(str(row[c]) for c in cols) + " |")
    return "\n".join([header, sep] + rows) + "\n"


def save(
    output_path: Union[str, Path],
    region_df: pd.DataFrame,
    agent_b_results: dict,
    agent_c_results: dict,
    ref_date: str = "",
) -> Path:
    """
    Parameters
    ----------
    output_path     : 저장할 .md 파일 경로
    region_df       : Agent A 결과
    agent_b_results : Agent B 결과 dict
    agent_c_results : Agent C 결과 dict
    ref_date        : 기준년월 문자열 (표시용)
    """
    today = datetime.now().strftime("%Y-%m-%d")
    lines = [
        f"# 경기도 인구자료 정리 리포트",
        f"",
        f"> 기준년월: {ref_date or '최신'}  |  작성일: {today}",
        f"",
        f"---",
        f"",
        f"## 1. 지역별 인구 현황",
        f"",
        _df_to_md(region_df),
        f"",
    ]

    # ── 수원시 행정동 ──
    gu_summary = agent_b_results.get("gu_summary", pd.DataFrame())
    detail = agent_b_results.get("detail", pd.DataFrame())
    top5 = agent_b_results.get("top5", pd.DataFrame())
    bottom5 = agent_b_results.get("bottom5", pd.DataFrame())

    lines += [
        f"## 2. 수원시 행정동별 인구",
        f"",
        f"### 2-1. 구별 소계",
        f"",
        _df_to_md(gu_summary),
        f"",
        f"### 2-2. 인구 상위 5 행정동",
        f"",
        _df_to_md(top5),
        f"",
        f"### 2-3. 인구 하위 5 행정동",
        f"",
        _df_to_md(bottom5),
        f"",
    ]

    # ── 생애주기 ──
    lifecycle_df = agent_c_results.get("lifecycle", pd.DataFrame())
    dependency_df = agent_c_results.get("dependency", pd.DataFrame())

    lines += [
        f"## 3. 수원시 생애주기 분석",
        f"",
        f"### 3-1. 생애주기별 인구",
        f"",
        _df_to_md(lifecycle_df),
        f"",
        f"### 3-2. 부양비 지표",
        f"",
        _df_to_md(dependency_df),
        f"",
        f"---",
        f"",
        f"*데이터 출처: 행정안전부 주민등록인구통계 / 통계청 KOSIS*",
    ]

    out = Path(output_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text("\n".join(lines), encoding="utf-8")
    return out
