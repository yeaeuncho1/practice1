import streamlit as st
import pandas as pd
import altair as alt

st.set_page_config(
    page_title="AI 활용 프로젝트 기획서",
    page_icon="🤖",
    layout="wide",
)

st.markdown("""
<style>
.block-container {padding-top: 2rem; padding-bottom: 3rem;}
.metric-card {
    padding: 1rem 1.2rem;
    border: 1px solid #e5e7eb;
    border-radius: 16px;
    background: #ffffff;
    box-shadow: 0 4px 14px rgba(0,0,0,0.04);
}
.section-card {
    padding: 1.2rem 1.4rem;
    border: 1px solid #e5e7eb;
    border-radius: 18px;
    background: #fafafa;
    margin-bottom: 1rem;
}
.badge {
    display: inline-block;
    padding: 0.2rem 0.6rem;
    border-radius: 999px;
    background: #eef2ff;
    color: #3730a3;
    font-size: 0.85rem;
    font-weight: 600;
}
.small-muted {
    color: #6b7280;
    font-size: 0.92rem;
}
</style>
""", unsafe_allow_html=True)

departments = [
    "MD팀",
    "스탭부서",
    "점포운영팀",
    "마케팅팀",
    "영업기획팀",
    "브랜드전략팀",
    "프로모션운영팀",
    "재고/상품관리팀",
]

kpi_df = pd.DataFrame({
    "KPI": ["데이터 취합 시간", "오류 발생률", "업데이트 처리 시간", "반복 업무 비중", "행사 준비 리드타임"],
    "현재": [7, 5, 48, 40, 2.5],
    "목표": [2, 1, 0.5, 10, 1.25],
    "단위": ["시간", "%", "시간", "%", "일"]
})

chart_df = pd.DataFrame({
    "구분": ["현재", "목표"] * 3,
    "지표": ["취합 시간"] * 2 + ["오류 발생률"] * 2 + ["반복 업무 비중"] * 2,
    "값": [7, 2, 5, 1, 40, 10]
})

roadmap_df = pd.DataFrame({
    "단계": ["1단계", "2단계", "3단계", "4단계"],
    "기간": ["1~2주", "3~4주", "5~6주", "7~8주"],
    "핵심 내용": [
        "표준 엑셀 양식 정의 및 데이터 구조 설계",
        "파일 업로드 감지 및 엑셀 처리 로직 구축",
        "특정 행사 파일럿 적용 및 오류 보완",
        "전 부서 확산 및 타 업무 영역 확대"
    ]
})

st.title("AI 활용 프로젝트 기획서")
st.caption("표준 엑셀 양식 기반 AI 자동 취합 시스템 구축안")

with st.container():
    c1, c2, c3 = st.columns(3)
    c1.metric("문제", "수기 취합 반복", "오류·누락·대기시간 발생")
    c2.metric("솔루션", "AI 자동 취합", "업로드 즉시 통합")
    c3.metric("기대효과", "취합 70% 단축", "오류 80% 감소")

st.markdown("---")

left, right = st.columns([1.15, 0.85])

with left:
    st.subheader("1. 프로젝트 개요")
    st.markdown("""
<div class="section-card">
대형 행사 운영 시 여러 부서가 엑셀 파일을 개별 전달하고, 담당자가 이를 수기로 병합하면서
오류·누락·버전 혼선·대기시간이 반복적으로 발생합니다.
<br><br>
본 프로젝트는 <b>표준 엑셀 양식</b>을 기반으로 업로드 즉시 데이터를 자동 구조화·비교·병합하는
<b>Zero-Touch AI 데이터 취합 구조</b>를 구현하는 것을 목표로 합니다.
</div>
""", unsafe_allow_html=True)

    st.subheader("2. 현황 및 문제 정의")
    st.markdown("##### 운영 규모")
    a, b, c, d = st.columns(4)
    a.metric("참여 부서", "8개")
    b.metric("참여 점포", "약 23개")
    c.metric("브랜드 수", "약 50~100개")
    d.metric("업데이트 반복", "평균 3~4회")

    st.markdown("##### 핵심 문제")
    st.markdown("""
- 다수 부서가 동일 데이터를 반복 수정·전달하면서 비효율 발생
- 수기 병합으로 취합에 **약 6~8시간** 소요
- 누락·중복·버전 혼선 포함 **오류 약 5%** 발생 가능
- 업데이트 간 대기시간이 **3~4시간 ~ 최대 1~2일**로 편차가 큼
- 전체 업무 중 단순 취합/정리 비중이 **40% 이상**
""")

    st.subheader("3. 참여 부서(텍스트 기획안 반영)")
    dept_cols = st.columns(2)
    for idx, dept in enumerate(departments):
        with dept_cols[idx % 2]:
            st.markdown(f"- {dept}")

    st.caption("참여 부서는 기획안 기준 8~10개 수준으로 반영했으며, 실제 운영 시 부서명은 조정 가능합니다.")

with right:
    st.subheader("4. KPI 요약")
    st.dataframe(kpi_df, use_container_width=True, hide_index=True)

    st.markdown("##### 주요 개선 지표 시각화")
    chart = (
        alt.Chart(chart_df)
        .mark_bar()
        .encode(
            x=alt.X("지표:N", title=None),
            y=alt.Y("값:Q", title="값"),
            column=alt.Column("구분:N", title=None),
            tooltip=["지표", "구분", "값"]
        )
        .properties(height=280)
    )
    st.altair_chart(chart, use_container_width=True)

st.markdown("---")

st.subheader("5. AI 솔루션 설계")
col1, col2, col3, col4 = st.columns(4)
col1.markdown("""
<div class="metric-card">
<span class="badge">Problem</span>
<ul>
<li>엑셀 파일 분산 관리</li>
<li>반복 수기 취합</li>
<li>부서 증가에 따른 복잡도 상승</li>
</ul>
</div>
""", unsafe_allow_html=True)
col2.markdown("""
<div class="metric-card">
<span class="badge">Data</span>
<ul>
<li>표준 엑셀 양식</li>
<li>브랜드/할인율/점포/상품 컬럼</li>
<li>부서별 제출 원본 데이터</li>
</ul>
</div>
""", unsafe_allow_html=True)
col3.markdown("""
<div class="metric-card">
<span class="badge">Insight</span>
<ul>
<li>자동 읽기 및 구조화</li>
<li>변경사항 비교 및 이력 관리</li>
<li>누락/중복 탐지</li>
</ul>
</div>
""", unsafe_allow_html=True)
col4.markdown("""
<div class="metric-card">
<span class="badge">Action</span>
<ul>
<li>통합 마스터 시트 생성</li>
<li>실시간 업데이트 반영</li>
<li>변경사항 자동 알림</li>
</ul>
</div>
""", unsafe_allow_html=True)

st.markdown("##### 자동 취합 프로세스")
p1, p2, p3 = st.columns(3)
p1.info("① 입력: 각 부서가 표준 엑셀 양식을 작성해 사내 메신저 채팅방 또는 공유 폴더에 업로드")
p2.success("② 자동 처리: 업로드 감지 → 엑셀 읽기 → 컬럼 구조화 → 기존 데이터 비교 → 변경 반영 → 중복 제거")
p3.warning("③ 결과: 최신 통합 마스터 시트 생성, 변경사항 자동 알림")

st.markdown("---")

st.subheader("6. 사람 vs AI 역할 분담")
human_ai = pd.DataFrame({
    "구분": ["사람", "AI"],
    "역할": [
        "엑셀 입력, 최종 검수, 의사결정",
        "데이터 병합, 비교, 정리, 오류 탐지"
    ]
})
st.table(human_ai)

st.markdown("---")

st.subheader("7. 확장성 설계")
ex1, ex2 = st.columns(2)
with ex1:
    st.markdown("""
**공통 시스템**
- 파일 업로드 감지
- 데이터 읽기
- 자동 병합
- 변경 추적
- 오류 탐지
""")
with ex2:
    st.markdown("""
**업무별 표준 양식**
- 행사 데이터 양식
- 프로모션 양식
- 광고 집행 양식
- 점포 운영 데이터 양식
""")

st.markdown("적용 가능 영역: 행사 기획 데이터 취합 / 프로모션·이벤트 계획 관리 / 광고 집행 현황 관리 / 점포 운영 데이터 취합 / 상품 구성·입점 관리")

st.markdown("---")

st.subheader("8. 리스크 및 대응 방안")
risk1, risk2, risk3 = st.columns(3)
risk1.markdown("""
<div class="section-card">
<b>엑셀 양식 미준수</b><br>
형식 불일치로 자동 처리 오류 발생 가능<br><br>
<b>대응</b><br>
- 입력 양식 고정
- 드롭다운/체크박스 활용
</div>
""", unsafe_allow_html=True)
risk2.markdown("""
<div class="section-card">
<b>초기 시스템 적응 문제</b><br>
기존 방식 유지 시도 가능성<br><br>
<b>대응</b><br>
- 파일럿 운영
- 가이드 및 교육 제공
</div>
""", unsafe_allow_html=True)
risk3.markdown("""
<div class="section-card">
<b>시스템 장애 리스크</b><br>
자동화 실패 시 업무 지연 우려<br><br>
<b>대응</b><br>
- 백업 시트 운영
- 수기 fallback 프로세스 유지
</div>
""", unsafe_allow_html=True)

st.markdown("---")

st.subheader("9. 8주 실행 계획")
st.dataframe(roadmap_df, use_container_width=True, hide_index=True)

st.markdown("---")

st.subheader("10. 결론")
st.markdown("""
<div class="section-card">
본 프로젝트의 핵심은 단순 자동화를 넘어 <b>업로드만 하면 자동으로 취합되는 데이터 운영 구조</b>를 구축하는 데 있습니다.
<br><br>
이를 통해 반복적인 수기 업무를 제거하고, 마케팅 조직이 전략·성과 중심 업무에 집중할 수 있는 환경을 마련할 수 있습니다.
또한 행사 데이터뿐 아니라 다양한 업무로 확장 가능한 <b>범용 데이터 취합 인프라</b>로 활용 가능합니다.
</div>
""", unsafe_allow_html=True)

st.markdown("<div class='small-muted'>실행 방법: <code>streamlit run app.py</code></div>", unsafe_allow_html=True)
