"""
Orchestrator — 파일을 분류하고 각 에이전트에 위임
사용법:
    python main.py file1.xlsx file2.csv ...
또는 Python API:
    from main import run
    results = run(["file1.xlsx"])
"""
from __future__ import annotations

import sys
from pathlib import Path
from typing import Union

import pandas as pd

from agents.agent_a import run as run_agent_a
from agents.agent_b import run as run_agent_b
from agents.agent_c import run as run_agent_c
from output.excel_writer import save as save_excel
from output.report_writer import save as save_report

OUTPUT_DIR = Path("output_results")


def _load_header(path: Union[str, Path]) -> list[str]:
    """파일의 첫 번째 행(헤더)만 읽어 컬럼 목록 반환"""
    p = Path(path)
    try:
        if p.suffix.lower() in (".xlsx", ".xls"):
            df = pd.read_excel(p, nrows=0)
        else:
            df = pd.read_csv(p, encoding="utf-8-sig", nrows=0)
        return list(df.columns)
    except Exception:
        return []


def _classify_file(path: Union[str, Path]) -> str:
    """
    파일 유형 분류:
      "dong"   → 행정동 단위 포함 (Agent A·B 모두)
      "region" → 시군구/전국 단위 (Agent A만)
      "unknown"

    주민등록인구및세대현황 파일은 데이터 내용으로 판별:
      행정동(동 레벨) 행이 포함되어 있으면 "dong"
    """
    from agents.loader import load_file
    try:
        df = load_file(path)
    except Exception:
        return "unknown"

    if df.empty:
        return "unknown"

    # 컬럼에 행정동 키워드가 있으면 확실히 dong
    header_str = " ".join(str(c) for c in df.columns)
    if any(k in header_str for k in ["행정동", "읍면동", "동명"]):
        return "dong"

    # 통합 주소 컬럼(행정기관 등)에서 행정동 레벨 행이 있는지 확인
    name_col = next(
        (c for c in df.columns if ("행정기관" in c or "행정구역" in c) and "코드" not in c),
        None,
    )
    if name_col is not None:
        import re
        has_dong = df[name_col].astype(str).str.split().apply(
            lambda parts: len(parts) > 0 and bool(re.search(r"동\d*$|[0-9]+동$", parts[-1]))
        ).any()
        if has_dong:
            return "dong"

    # 컬럼/파일명 기반 판별
    dong_name_kw = ["행정동", "동별"]
    if any(k in Path(path).stem for k in dong_name_kw):
        return "dong"

    region_keywords = ["행정구역", "시군구", "시도", "총인구", "인구수", "행정기관"]
    if any(k in header_str for k in region_keywords):
        return "region"

    return "unknown"


def run(
    files: list[Union[str, Path]],
    prev_files: list[Union[str, Path]] | None = None,
    output_dir: Union[str, Path] = OUTPUT_DIR,
    save_files: bool = True,
) -> dict:
    """
    Parameters
    ----------
    files      : 당해 연도 입력 파일 목록
    prev_files : 전년도 파일 목록 (증감 계산용)
    output_dir : 결과 저장 디렉토리
    save_files : True면 Excel/Markdown 파일 저장

    Returns
    -------
    dict {
        "region": pd.DataFrame,          # Agent A 결과
        "b": dict,                        # Agent B 결과
        "c": dict,                        # Agent C 결과
        "excel_path": Path | None,
        "report_path": Path | None,
        "missing": list[str],            # 분류 못한 파일
    }
    """
    region_files, dong_files, unknown_files = [], [], []

    for f in files:
        ftype = _classify_file(f)
        if ftype == "region":
            region_files.append(f)
        elif ftype == "dong":
            dong_files.append(f)
            region_files.append(f)  # dong 파일도 Agent A에 전달 (시 레벨 행 포함)
        else:
            unknown_files.append(str(f))

    if unknown_files:
        print(f"[Orchestrator] 분류 불가 파일 (수동 확인 필요): {unknown_files}")

    # ── Agent A ──
    region_df = pd.DataFrame()
    if region_files:
        print(f"[Agent A] 지역별 인구 처리: {[str(f) for f in region_files]}")
        region_df = run_agent_a(region_files, prev_files=prev_files)
    else:
        print("[Agent A] 처리할 지역 파일 없음")

    # ── Agent B ──
    b_results: dict = {"detail": pd.DataFrame(), "gu_summary": pd.DataFrame(), "top5": pd.DataFrame(), "bottom5": pd.DataFrame()}
    if dong_files:
        print(f"[Agent B] 행정동 인구 처리: {[str(f) for f in dong_files]}")
        b_results = run_agent_b(dong_files)

    # ── Agent C ──
    c_results: dict = {"lifecycle": pd.DataFrame(), "dependency": pd.DataFrame(), "dong_pivot": pd.DataFrame()}
    if dong_files or region_files:
        source_files = dong_files if dong_files else region_files
        print(f"[Agent C] 생애주기 분석: {[str(f) for f in source_files]}")
        c_results = run_agent_c(source_files, agent_b_detail=b_results.get("detail"))

    excel_path = None
    report_path = None

    if save_files and (not region_df.empty or not b_results["detail"].empty):
        out_dir = Path(output_dir)
        out_dir.mkdir(parents=True, exist_ok=True)
        excel_path = save_excel(out_dir / "population_summary.xlsx", region_df, b_results, c_results)
        report_path = save_report(out_dir / "population_report.md", region_df, b_results, c_results)
        print(f"[Output] Excel: {excel_path}")
        print(f"[Output] Report: {report_path}")

    return {
        "region": region_df,
        "b": b_results,
        "c": c_results,
        "excel_path": excel_path,
        "report_path": report_path,
        "missing": unknown_files,
    }


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("사용법: python main.py <파일1> [파일2] ...")
        sys.exit(1)

    input_files = [Path(p) for p in sys.argv[1:]]
    missing = [str(f) for f in input_files if not f.exists()]
    if missing:
        print(f"파일을 찾을 수 없습니다: {missing}")
        sys.exit(1)

    results = run(input_files)

    if results["region"] is not None and not results["region"].empty:
        print("\n=== 지역별 인구 현황 ===")
        print(results["region"].to_string(index=False))

    if results["c"]["lifecycle"] is not None and not results["c"]["lifecycle"].empty:
        print("\n=== 수원시 생애주기 ===")
        print(results["c"]["lifecycle"].to_string(index=False))
