"""
Streamlit 대시보드 — 경기도 인구자료 정리 서비스
실행: streamlit run dashboard.py
"""
from __future__ import annotations

import tempfile
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st

from main import run as orchestrate
from output.excel_writer import to_bytes as excel_bytes

# ───── 페이지 설정 ─────
st.set_page_config(
    page_title="인구자료 정리 에이전트",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ───── 세션 상태 ─────
if "results" not in st.session_state:
    st.session_state.results = None

def _save_uploads(files) -> list[str]:
    tmp_dir = Path(tempfile.mkdtemp())
    paths = []
    for f in files:
        dest = tmp_dir / f.name
        dest.write_bytes(f.read())
        paths.append(str(dest))
    return paths

def _styled(df: pd.DataFrame):
    """정수·대용량 숫자 컬럼에 콤마 서식, 소수 컬럼에 소수점 서식 적용"""
    fmt = {}
    for c in df.columns:
        if pd.api.types.is_integer_dtype(df[c]):
            fmt[c] = "{:,}"
        elif pd.api.types.is_float_dtype(df[c]):
            sample = df[c].dropna()
            if not sample.empty and sample.abs().mean() > 1000:
                fmt[c] = "{:,.0f}"
    return df.style.format(fmt, na_rep="-")

# ═══════════════════════════════════════════════
# 헤더
# ═══════════════════════════════════════════════
st.title("📊 인구자료 정리 에이전트")
st.caption("통계청 KOSIS · 행정안전부 주민등록 파일을 업로드하면 자동으로 분석합니다.")
st.divider()

# ═══════════════════════════════════════════════
# 업로드 & 실행 섹션 (단일 행)
# ═══════════════════════════════════════════════
up_col, btn_col = st.columns([5, 1], vertical_alignment="bottom")

with up_col:
    uploaded = st.file_uploader(
        "📂 파일 선택 (CSV / Excel, 복수 선택 가능)",
        type=["csv", "xlsx", "xls"],
        accept_multiple_files=True,
        label_visibility="visible",
    )
    st.caption("※ 분석 연도는 업로드하는 자료에 따라 달라집니다. 여러 연도 파일을 함께 올리면 연도별로 비교 분석됩니다.")

with btn_col:
    run_btn = st.button("▶ 분석 실행", type="primary", use_container_width=True)

# ═══════════════════════════════════════════════
# 분석 실행
# ═══════════════════════════════════════════════
if run_btn:
    if not uploaded:
        st.error("파일을 먼저 업로드하세요.")
    else:
        with st.spinner("데이터 분석 중..."):
            try:
                file_paths = _save_uploads(uploaded)
                results = orchestrate(file_paths, save_files=False)
                st.session_state.results = results
            except Exception as e:
                st.error(f"오류 발생: {e}")
                st.exception(e)

# ═══════════════════════════════════════════════
# 결과 없을 때 안내
# ═══════════════════════════════════════════════
if st.session_state.results is None:
    st.info("파일을 업로드하고 **▶ 분석 실행**을 누르면 결과가 여기에 표시됩니다.")
    with st.expander("지원 파일 형식 안내"):
        st.markdown("""
| 파일 | 용도 |
|------|------|
| 주민등록인구및세대현황 (전국·시도) | 전국·경기도 비교 |
| 주민등록인구및세대현황 (경기도 시군구) | 5개 시 비교 |
| 주민등록인구및세대현황 (수원시 행정동) | 행정동 상세 분석 |
| KOSIS 연령별 인구 (1세 단위) | 생애주기 분석 |
        """)
    st.stop()

# ═══════════════════════════════════════════════
# 데이터 언팩
# ═══════════════════════════════════════════════
res = st.session_state.results
region_df: pd.DataFrame = res.get("region", pd.DataFrame())
b_res: dict = res.get("b", {})
c_res: dict = res.get("c", {})
missing: list = res.get("missing", [])

if missing:
    st.warning(f"분류 불가 파일 (확인 필요): {', '.join(missing)}")

st.success("분석 완료!")
st.divider()

# 최신 연도 데이터 추출 (KPI용)
has_year = not region_df.empty and "연도" in region_df.columns
if has_year:
    latest_year = int(region_df["연도"].max())
    latest_region_df = region_df[region_df["연도"] == latest_year]
else:
    latest_year = None
    latest_region_df = region_df

# ═══════════════════════════════════════════════
# KPI 메트릭
# ═══════════════════════════════════════════════
if not region_df.empty:
    nat_row = latest_region_df[latest_region_df["지역"] == "전국"]
    gyeonggi_row = latest_region_df[latest_region_df["지역"] == "경기도"]
    suwon_row = latest_region_df[latest_region_df["지역"] == "수원시"]

    year_label = f" ({latest_year}년)" if latest_year else ""
    k1, k2, k3, k4 = st.columns(4)
    with k1:
        val = int(nat_row["총인구"].iloc[0]) if not nat_row.empty else 0
        st.metric(f"전국 총인구{year_label}", f"{val:,}명")
    with k2:
        val = int(gyeonggi_row["총인구"].iloc[0]) if not gyeonggi_row.empty else 0
        st.metric(f"경기도 총인구{year_label}", f"{val:,}명")
    with k3:
        val = int(suwon_row["총인구"].iloc[0]) if not suwon_row.empty else 0
        chg = suwon_row["전년대비(%)"].iloc[0] if (not suwon_row.empty and "전년대비(%)" in suwon_row.columns) else None
        chg_str = f"{chg:+.2f}%" if (chg is not None and not pd.isna(chg)) else None
        st.metric(f"수원시 총인구{year_label}", f"{val:,}명", delta=chg_str)
        # 비교 기준 연도 및 절대 증감 표시
        if chg is not None and not pd.isna(chg) and has_year:
            _sw_years = sorted(region_df[region_df["지역"] == "수원시"]["연도"].dropna().unique())
            _cur_idx = list(_sw_years).index(latest_year) if latest_year in _sw_years else -1
            if _cur_idx > 0:
                _prev_yr = int(_sw_years[_cur_idx - 1])
                _prev_row = region_df[(region_df["연도"] == _prev_yr) & (region_df["지역"] == "수원시")]
                if not _prev_row.empty:
                    _prev_val = int(_prev_row["총인구"].iloc[0])
                    _diff = val - _prev_val
                    st.caption(f"({_prev_yr}년 {_prev_val:,}명 대비 {_diff:+,}명)")
    with k4:
        _lc_kpi = c_res.get("lifecycle", pd.DataFrame())
        # 다년도면 최신 연도 기준
        if not _lc_kpi.empty and "연도" in _lc_kpi.columns:
            _lc_kpi_yr = int(_lc_kpi["연도"].max())
            _lc_kpi = _lc_kpi[_lc_kpi["연도"] == _lc_kpi_yr].drop(columns=["연도"])
            _lc_kpi_label = f"수원시 노년기 비율 ({_lc_kpi_yr}년)"
        else:
            _lc_kpi_label = "수원시 노년기 비율"
        if not _lc_kpi.empty and "생애주기" in _lc_kpi.columns:
            _total_row = _lc_kpi[_lc_kpi["생애주기"] == "합계"]
            _total_pop = int(_total_row["인구"].iloc[0]) if not _total_row.empty else 0
            if _total_pop > 0:
                old_row = _lc_kpi[_lc_kpi["생애주기"] == "노년기"]
                old_ratio = float(old_row["비율(%)"].iloc[0]) if not old_row.empty else 0.0
                st.metric(_lc_kpi_label, f"{old_ratio:.1f}%")
            else:
                st.metric(_lc_kpi_label, "-", help="생애주기 분석을 위한 KOSIS 연령별 인구 파일이 필요합니다.")
        else:
            st.metric("노년기 비율", "-")

    st.divider()

# ═══════════════════════════════════════════════
# 생애주기 데이터 준비 (탭 3·4 공통)
# ═══════════════════════════════════════════════
lifecycle_df   = c_res.get("lifecycle",  pd.DataFrame())
dependency_df  = c_res.get("dependency", pd.DataFrame())
gu_pivot_lc    = c_res.get("gu_pivot",   pd.DataFrame())
gu_count_lc    = c_res.get("gu_count",   pd.DataFrame())
dong_pivot_lc  = c_res.get("dong_pivot", pd.DataFrame())
dong_count_lc  = c_res.get("dong_count", pd.DataFrame())

_has_lc_year = not lifecycle_df.empty and "연도" in lifecycle_df.columns
_lc_year_label = ""
if _has_lc_year:
    _lc_years = sorted(lifecycle_df["연도"].dropna().unique().astype(int))
    _lc_col1, _lc_col2 = st.columns([1, 4])
    with _lc_col1:
        st.caption("🌀 생애주기 분석 연도")
    with _lc_col2:
        _sel_lc_year = st.radio(
            "생애주기 연도 선택", options=_lc_years,
            index=len(_lc_years) - 1, horizontal=True,
            key="lc_year_radio", label_visibility="collapsed",
        )
    _lc_year_label = f" ({_sel_lc_year}년)"

    def _lc_filter(d: pd.DataFrame) -> pd.DataFrame:
        if d.empty or "연도" not in d.columns:
            return d
        return d[d["연도"] == _sel_lc_year].drop(columns=["연도"]).reset_index(drop=True)

    lifecycle_df   = _lc_filter(lifecycle_df)
    dependency_df  = _lc_filter(dependency_df)
    gu_pivot_lc    = _lc_filter(gu_pivot_lc)
    gu_count_lc    = _lc_filter(gu_count_lc)
    dong_pivot_lc  = _lc_filter(dong_pivot_lc)
    dong_count_lc  = _lc_filter(dong_count_lc)

# ═══════════════════════════════════════════════
# 탭 레이아웃
# ═══════════════════════════════════════════════
tab1, tab2, tab3, tab4 = st.tabs(["📍 지역별 인구", "🏘️ 행정동 인구", "🌀 생애주기 분포", "🗺️ 생애주기 히트맵"])

# ── 탭 1: 지역별 인구 ──
with tab1:
    if region_df.empty:
        st.info("지역별 인구 데이터가 없습니다.")
    else:
        if has_year:
            # 연도별 추이 차트 (전국 제외)
            trend_df = region_df[region_df["지역"] != "전국"].copy()
            fig_line = px.line(
                trend_df, x="연도", y="총인구", color="지역",
                title="지역별 인구 추이 (연도별)",
                markers=True,
                labels={"총인구": "총인구 (명)", "연도": "연도"},
                color_discrete_sequence=px.colors.qualitative.Set2,
            )
            fig_line.update_layout(height=420, yaxis_tickformat=",")
            fig_line.update_traces(hovertemplate="%{y:,}<extra></extra>")
            st.plotly_chart(fig_line, use_container_width=True)

            # 연도 선택기로 막대차트 + 상세 테이블
            years = sorted(region_df["연도"].unique())
            sel_year = st.radio("연도 선택", options=years, index=len(years)-1, horizontal=True)
            sel_df = region_df[region_df["연도"] == sel_year]
            c1, c2 = st.columns([3, 2])
            with c1:
                display_df = sel_df[sel_df["지역"] != "전국"].copy()
                fig_bar = px.bar(
                    display_df, x="지역", y="총인구", color="지역",
                    title=f"{sel_year}년 지역별 총인구 비교", text_auto=True,
                    labels={"총인구": "총인구 (명)"},
                    color_discrete_sequence=px.colors.qualitative.Set2,
                )
                fig_bar.update_traces(texttemplate="%{y:,}", textposition="outside")
                fig_bar.update_layout(showlegend=False, height=400, yaxis_tickformat=",")
                st.plotly_chart(fig_bar, use_container_width=True)
            with c2:
                st.dataframe(_styled(sel_df), use_container_width=True, hide_index=True)

            st.subheader("전체 연도 데이터")
            st.dataframe(_styled(region_df), use_container_width=True, hide_index=True)
        else:
            c1, c2 = st.columns([3, 2])
            with c1:
                display_df = region_df[region_df["지역"] != "전국"].copy() if len(region_df) > 1 else region_df
                fig = px.bar(
                    display_df, x="지역", y="총인구", color="지역",
                    title="지역별 총인구 비교", text_auto=True,
                    labels={"총인구": "총인구 (명)"},
                    color_discrete_sequence=px.colors.qualitative.Set2,
                )
                fig.update_traces(texttemplate="%{y:,}", textposition="outside")
                fig.update_layout(showlegend=False, height=450, yaxis_tickformat=",")
                st.plotly_chart(fig, use_container_width=True)
            with c2:
                st.subheader("상세 데이터")
                st.dataframe(_styled(region_df), use_container_width=True, hide_index=True)

# ── 탭 2: 행정동 인구 ──
with tab2:
    detail: pd.DataFrame = b_res.get("detail", pd.DataFrame())
    gu_summary: pd.DataFrame = b_res.get("gu_summary", pd.DataFrame())
    top5: pd.DataFrame = b_res.get("top5", pd.DataFrame())

    if detail.empty:
        st.info("행정동 인구 데이터가 없습니다.")
    else:
        has_year_b = "연도" in detail.columns and detail["연도"].notna().any()

        if has_year_b and "연도" in gu_summary.columns:
            # 구별 인구 추이 (연도별 라인 차트)
            fig_gu_trend = px.line(
                gu_summary, x="연도", y="총인구", color="구",
                title="수원시 구별 인구 추이 (연도별)",
                markers=True,
                labels={"총인구": "총인구 (명)"},
                color_discrete_sequence=px.colors.qualitative.Pastel,
            )
            fig_gu_trend.update_layout(height=380, yaxis_tickformat=",")
            fig_gu_trend.update_traces(hovertemplate="%{y:,}<extra></extra>")
            st.plotly_chart(fig_gu_trend, use_container_width=True)

            latest_yr_b = int(detail["연도"].max())
            latest_gu = gu_summary[gu_summary["연도"] == latest_yr_b]
            c1, c2 = st.columns(2)
            with c1:
                fig_pie = px.pie(
                    latest_gu, names="구", values="총인구",
                    title=f"{latest_yr_b}년 수원시 구별 인구 비중",
                    color_discrete_sequence=px.colors.qualitative.Pastel,
                )
                st.plotly_chart(fig_pie, use_container_width=True)
            with c2:
                if not top5.empty:
                    st.subheader(f"인구 상위 5 행정동 ({latest_yr_b}년)")
                    fig_top5 = px.bar(
                        top5, y="행정동", x="총인구", orientation="h",
                        color="구", text_auto=True, labels={"총인구": "인구 (명)"},
                    )
                    fig_top5.update_traces(texttemplate="%{x:,}")
                    fig_top5.update_layout(height=300, yaxis={"categoryorder": "total ascending"}, xaxis_tickformat=",")
                    st.plotly_chart(fig_top5, use_container_width=True)

            years_b = sorted(detail["연도"].dropna().unique().astype(int))
            sel_year_b = st.radio("연도 선택", options=years_b, index=len(years_b)-1, horizontal=True)
            st.subheader(f"{sel_year_b}년 행정동 전체 목록")
            st.dataframe(_styled(detail[detail["연도"] == sel_year_b]), use_container_width=True, hide_index=True)
        else:
            c1, c2 = st.columns(2)
            with c1:
                if not gu_summary.empty:
                    fig = px.pie(
                        gu_summary, names="구", values="총인구",
                        title="수원시 구별 인구 비중",
                        color_discrete_sequence=px.colors.qualitative.Pastel,
                    )
                    st.plotly_chart(fig, use_container_width=True)
            with c2:
                if not top5.empty:
                    st.subheader("인구 상위 5 행정동")
                    fig = px.bar(
                        top5, y="행정동", x="총인구", orientation="h",
                        color="구", text_auto=True, labels={"총인구": "인구 (명)"},
                    )
                    fig.update_traces(texttemplate="%{x:,}")
                    fig.update_layout(height=300, yaxis={"categoryorder": "total ascending"}, xaxis_tickformat=",")
                    st.plotly_chart(fig, use_container_width=True)

            st.subheader("행정동 전체 목록")
            st.dataframe(_styled(detail), use_container_width=True, hide_index=True)

# ── 탭 3: 생애주기 분포 ──
with tab3:
    has_lifecycle = not lifecycle_df.empty and "생애주기" in lifecycle_df.columns

    # 서브탭은 항상 표시
    lc_sub1, lc_sub2, lc_sub3 = st.tabs(["🏙️ 수원시 전체", "🏛️ 구별 (4개 구)", "🏘️ 행정동별"])

    # ── 서브탭 1: 수원시 전체 ──
    with lc_sub1:
        if not has_lifecycle:
            msg = lifecycle_df.iloc[0, 0] if not lifecycle_df.empty else "생애주기 데이터가 없습니다."
            st.info(str(msg))
        else:
            lc_data = lifecycle_df[lifecycle_df["생애주기"] != "합계"].copy()
            c1, c2 = st.columns(2)
            with c1:
                fig = px.pie(
                    lc_data, names="생애주기", values="인구",
                    title="수원시 생애주기별 인구 비율", hole=0.35,
                    color_discrete_sequence=px.colors.sequential.RdBu,
                )
                st.plotly_chart(fig, use_container_width=True)
            with c2:
                fig = px.bar(
                    lc_data, x="생애주기", y="비율(%)",
                    title="수원시 생애주기별 비율 (%)", text="비율(%)",
                    color="비율(%)", color_continuous_scale="Blues",
                )
                fig.update_traces(textposition="outside")
                fig.update_layout(showlegend=False, height=400)
                st.plotly_chart(fig, use_container_width=True)

            # 생애주기 표
            st.subheader("수원시 생애주기별 인구")
            st.dataframe(_styled(lifecycle_df), use_container_width=True, hide_index=True)

            if not dependency_df.empty:
                st.subheader("부양비 지표")
                dep_cols = st.columns(len(dependency_df))
                for i, (_, row) in enumerate(dependency_df.iterrows()):
                    with dep_cols[i]:
                        val = row.iloc[1]
                        st.metric(label=row["지표"], value=f"{val:.1f}" if pd.notna(val) else "-")
                st.dataframe(_styled(dependency_df), use_container_width=True, hide_index=True)

    # ── 서브탭 2: 구별 ──
    with lc_sub2:
        if gu_pivot_lc.empty:
            st.info("구별 생애주기 데이터가 없습니다. (KOSIS 연령별 인구 파일 필요)")
        else:
            stage_cols_gu = [c for c in gu_pivot_lc.columns if c != "구"]

            # 4개 구 비교 차트
            melt_gu = gu_pivot_lc.melt(
                id_vars=["구"], value_vars=stage_cols_gu,
                var_name="생애주기", value_name="비율(%)"
            )
            fig_gu_comp = px.bar(
                melt_gu, x="생애주기", y="비율(%)", color="구",
                barmode="group", title="4개 구 생애주기별 비율 비교 (%)",
                text="비율(%)",
                color_discrete_sequence=px.colors.qualitative.Pastel,
            )
            fig_gu_comp.update_traces(textposition="outside")
            fig_gu_comp.update_layout(height=420, legend_title="구")
            st.plotly_chart(fig_gu_comp, use_container_width=True)

            # 4개 구 통합 표 (명수 + 비율)
            st.subheader("4개 구 생애주기 현황")
            if not gu_count_lc.empty:
                combined_rows = []
                for _, gr in gu_pivot_lc.iterrows():
                    gu_n = gr["구"]
                    cnt_r = gu_count_lc[gu_count_lc["구"] == gu_n].iloc[0] if gu_n in gu_count_lc["구"].values else None
                    row_data = {"구": gu_n}
                    for s in stage_cols_gu:
                        row_data[f"{s}(명)"] = int(cnt_r[s]) if cnt_r is not None else "-"
                        row_data[f"{s}(%)"] = gr[s]
                    combined_rows.append(row_data)
                st.dataframe(_styled(pd.DataFrame(combined_rows)), use_container_width=True, hide_index=True)
            else:
                st.dataframe(_styled(gu_pivot_lc), use_container_width=True, hide_index=True)

    # ── 서브탭 3: 행정동별 ──
    with lc_sub3:
        if dong_pivot_lc.empty or "행정동" not in dong_pivot_lc.columns:
            st.info("행정동별 생애주기 데이터가 없습니다. (KOSIS 연령별 인구 파일 필요)")
        else:
            stage_cols_lc = [c for c in dong_pivot_lc.columns if c not in ["구", "행정동"]]
            dong_list = dong_pivot_lc["행정동"].tolist()

            # 전체 행정동 목록 표 (행=행정동, 열=생애주기 인원+비율)
            st.subheader("행정동별 생애주기 현황")
            if not dong_count_lc.empty:
                merged_rows = []
                for _, dr in dong_pivot_lc.iterrows():
                    d_name = dr["행정동"]
                    g_name = dr.get("구", "")
                    cnt_r = dong_count_lc[dong_count_lc["행정동"] == d_name].iloc[0] if d_name in dong_count_lc["행정동"].values else None
                    row_data = {"구": g_name, "행정동": d_name, "합계(명)": int(cnt_r["합계"]) if cnt_r is not None else "-"}
                    for s in stage_cols_lc:
                        row_data[f"{s}(명)"] = int(cnt_r[s]) if cnt_r is not None else "-"
                        row_data[f"{s}(%)"] = dr[s]
                    merged_rows.append(row_data)
                st.dataframe(_styled(pd.DataFrame(merged_rows)), use_container_width=True, hide_index=True)
            else:
                st.dataframe(_styled(dong_pivot_lc), use_container_width=True, hide_index=True)

            # 전체 행정동 생애주기 비교 차트
            st.subheader("전체 행정동 생애주기 비교")
            melt_df = dong_pivot_lc.melt(
                id_vars=["구", "행정동"] if "구" in dong_pivot_lc.columns else ["행정동"],
                value_vars=stage_cols_lc,
                var_name="생애주기", value_name="비율(%)"
            )
            fig_all = px.bar(
                melt_df, x="비율(%)", y="행정동", color="생애주기",
                orientation="h", barmode="stack",
                title="행정동별 생애주기 구성 비율 (%)",
                color_discrete_sequence=px.colors.qualitative.Set3,
            )
            fig_all.update_layout(height=max(400, len(dong_list) * 22), legend_title="생애주기")
            st.plotly_chart(fig_all, use_container_width=True)

# ── 탭 4: 행정동×생애주기 히트맵 ──
with tab4:
    if dong_pivot_lc.empty or "행정동" not in dong_pivot_lc.columns:
        st.info("행정동별 생애주기 데이터가 없습니다. (KOSIS 연령별 인구 파일 필요)")
    else:
        stage_cols = [c for c in dong_pivot_lc.columns if c not in ["구", "행정동"]]
        pivot_matrix = dong_pivot_lc.set_index("행정동")[stage_cols]
        fig = px.imshow(
            pivot_matrix,
            title=f"행정동별 생애주기 비율(%) 히트맵{_lc_year_label}",
            color_continuous_scale="Blues", aspect="auto",
            labels={"color": "비율(%)"}, zmin=0,
        )
        fig.update_layout(height=max(400, len(pivot_matrix) * 20))
        st.plotly_chart(fig, use_container_width=True)
        st.dataframe(_styled(dong_pivot_lc), use_container_width=True, hide_index=True)

# ═══════════════════════════════════════════════
# 다운로드
# ═══════════════════════════════════════════════
st.divider()
st.subheader("⬇ 결과 다운로드")
dl1, dl2, _ = st.columns([1, 1, 3])

with dl1:
    if not region_df.empty or not b_res.get("detail", pd.DataFrame()).empty:
        st.download_button(
            label="Excel 다운로드",
            data=excel_bytes(region_df, b_res, c_res),
            file_name="population_summary.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

with dl2:
    if not region_df.empty:
        from output.report_writer import save as _save_report
        tmp_md = Path(tempfile.mktemp(suffix=".md"))
        _save_report(tmp_md, region_df, b_res, c_res)
        st.download_button(
            label="Markdown 리포트",
            data=tmp_md.read_text(encoding="utf-8").encode("utf-8"),
            file_name="population_report.md",
            mime="text/markdown",
            use_container_width=True,
        )
