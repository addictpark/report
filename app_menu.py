import streamlit as st
import pandas as pd
import numpy as np

# ---------------------- 사이드바 메뉴 구성 ----------------------
st.sidebar.title("📋 메뉴")
menu = st.sidebar.radio("이동할 섹션을 선택하세요:", [
    "📁 파일 업로드 및 결측치 확인",
    "📊 운영 요약",
    "📈 상담 통계",
    "🧠 심리진단 통계",
    "🗂️ 상담 주제별 통계"
])

# ---------------------- 파일 업로드 ----------------------
if menu == "📁 파일 업로드 및 결측치 확인":
    st.title("보고서용 데이터 추출 프로그램")
    st.markdown("---")
    st.info(
        "업로드 전 꼭 확인하세요!\n"
        "- 상담 이력 엑셀 파일에서 아이디와 휴대전화번호 등 개인정보를 반드시 삭제한 후 업로드해야 합니다.\n"
        "- 해당 열(컬럼)은 반드시 삭제 또는 비식별(빈칸) 처리 후 업로드하세요."
    )
    st.header("파일 업로드")
    col1, col2 = st.columns(2)
    with col1:
        uploaded_counseling = st.file_uploader("상담 이력 엑셀 파일", type=["xlsx"], key="counseling")
    with col2:
        uploaded_diagnosis = st.file_uploader("진단 이력 엑셀 파일", type=["xlsx"], key="diagnosis")
    if uploaded_counseling and uploaded_diagnosis:
        st.session_state['df_counseling'] = pd.read_excel(uploaded_counseling)
        st.session_state['df_diagnosis'] = pd.read_excel(uploaded_diagnosis)
        st.success("파일 업로드가 완료되었습니다. 사이드바에서 메뉴를 선택해 주세요.")
    elif 'df_counseling' not in st.session_state or 'df_diagnosis' not in st.session_state:
        st.warning("두 파일을 모두 업로드해 주세요.")

# ---------------------- 데이터 준비 및 메뉴별 기능 ----------------------
if 'df_counseling' in st.session_state and 'df_diagnosis' in st.session_state:
    df_counseling = st.session_state['df_counseling']
    df_diagnosis = st.session_state['df_diagnosis']

    # 컬럼명 공백 제거 및 전처리
    df_counseling.columns = df_counseling.columns.str.strip()
    df_diagnosis.columns = df_diagnosis.columns.str.strip()

    # 개인정보 필드 차단
    sensitive_columns = ['신청직원이름', '휴대폰번호']
    sensitive_found = [col for col in sensitive_columns if col in df_counseling.columns]
    if sensitive_found:
        st.error(
            f"업로드 파일에 개인정보 열({', '.join(sensitive_found)})이 포함되어 있어 분석을 중단합니다.\n\n"
            "해당 열을 삭제한 후 다시 업로드해 주세요."
        )
        st.stop()

    # 날짜, 연월, 연령대, 성별 등 공통 전처리
    df_counseling['상담실시일'] = pd.to_datetime(df_counseling['상담실시일'], errors='coerce')
    df_diagnosis['진단실시일'] = pd.to_datetime(df_diagnosis['진단실시일'], errors='coerce')
    df_counseling['상담연월'] = df_counseling['상담실시일'].dt.to_period('M').astype(str)
    df_diagnosis['진단연월'] = df_diagnosis['진단실시일'].dt.to_period('M').astype(str)
    df_counseling['연령대'] = (df_counseling['신청직원나이'] // 10 * 10).astype('Int64').astype(str) + '대'
    df_counseling['성별'] = df_counseling['신청직원성별'].fillna('미상')

    # ID 클리닝 함수
    def clean_id(val):
        if pd.isnull(val):
            return None
        val = str(val).strip().replace(' ', '').replace('\u3000', '').lower()
        if val.endswith('.0'):
            val = val[:-2]
        if not val or val in {'nan', 'none', '0', '0.0'}:
            return None
        return val

    counseling_ids = df_counseling['아이디'].apply(clean_id)
    diagnosis_ids = df_diagnosis['아이디'].apply(clean_id)
    combined_ids = pd.concat([counseling_ids, diagnosis_ids])
    unique_ids = combined_ids.dropna().drop_duplicates()
    실계_인원수 = len(unique_ids)

    # 연월 목록 만들기
    valid_months_counseling = df_counseling['상담실시일'].dropna()
    valid_months_diagnosis = df_diagnosis['진단실시일'].dropna()
    all_valid_months = pd.concat([valid_months_counseling, valid_months_diagnosis])
    if len(all_valid_months) > 0:
        min_month = all_valid_months.min().to_period('M')
        max_month = all_valid_months.max().to_period('M')
        all_months = pd.period_range(min_month, max_month, freq='M').astype(str).tolist()
    else:
        all_months = []

    # ---------------------- 메뉴별 기능 구현 ----------------------
    if menu == "📊 운영 요약":
        st.header("📊 운영 요약")
        summary = pd.DataFrame({'연월': all_months})
        summary['심리상담'] = summary['연월'].apply(lambda m: df_counseling[df_counseling['상담연월'] == m]['아이디'].nunique())
        summary['심리진단'] = summary['연월'].apply(lambda m: df_diagnosis[df_diagnosis['진단연월'] == m]['아이디'].nunique())
        summary['합계'] = summary['심리상담'] + summary['심리진단']
        summary.loc[len(summary)] = ['누계', summary['심리상담'].sum(), summary['심리진단'].sum(), summary['합계'].sum()]
        summary.loc[len(summary)] = ['실계', df_counseling['아이디'].nunique(), df_diagnosis['아이디'].nunique(), 실계_인원수]
        st.subheader("서비스 이용 인원")
        st.dataframe(summary, use_container_width=True)

        summary_count = pd.DataFrame({'연월': all_months})
        summary_count['심리상담'] = summary_count['연월'].apply(lambda m: len(df_counseling[df_counseling['상담연월'] == m]))
        summary_count['심리진단'] = summary_count['연월'].apply(lambda m: len(df_diagnosis[df_diagnosis['진단연월'] == m]))
        summary_count['합계'] = summary_count['심리상담'] + summary_count['심리진단']
        summary_count.loc[len(summary_count)] = ['누계', summary_count['심리상담'].sum(), summary_count['심리진단'].sum(), summary_count['합계'].sum()]
        st.subheader("서비스 이용 횟수")
        st.dataframe(summary_count, use_container_width=True)

    elif menu == "📈 상담 통계":
        st.header("📈 상담 통계")
        # 상담유형별 인원 및 횟수
        type_people = df_counseling.groupby(['상담연월', '상담유형'])['아이디'].nunique().reset_index()
        type_people_summary = type_people.pivot(index='상담연월', columns='상담유형', values='아이디')
        type_people_summary = type_people_summary.reindex(all_months).fillna(0).astype(int)
        type_people_summary['합계'] = type_people_summary.sum(axis=1)
        type_people_summary.loc['누계'] = type_people_summary.sum()
        st.markdown("상담유형별 이용 인원")
        st.dataframe(type_people_summary)

        type_counts = df_counseling.groupby(['상담연월', '상담유형'])['사례번호'].count().reset_index()
        type_counts_summary = type_counts.pivot(index='상담연월', columns='상담유형', values='사례번호')
        type_counts_summary = type_counts_summary.reindex(all_months).fillna(0).astype(int)
        type_counts_summary['합계'] = type_counts_summary.sum(axis=1)
        type_counts_summary.loc['누계'] = type_counts_summary.sum()
        st.markdown("상담유형별 이용 횟수")
        st.dataframe(type_counts_summary)

        # 성별 인원 및 횟수
        st.markdown("---")
        st.subheader("2) 성별 이용 인원 및 횟수")
        gender_people = df_counseling.groupby(['상담연월', '성별'])['아이디'].nunique().reset_index()
        gender_people_summary = gender_people.pivot(index='상담연월', columns='성별', values='아이디')
        gender_people_summary = gender_people_summary.reindex(all_months).fillna(0).astype(int)
        gender_people_summary['합계'] = gender_people_summary.sum(axis=1)
        gender_people_summary.loc['누계'] = gender_people_summary.sum()
        st.dataframe(gender_people_summary)

        gender_counts = df_counseling.groupby(['상담연월', '성별'])['사례번호'].count().reset_index()
        gender_counts_summary = gender_counts.pivot(index='상담연월', columns='성별', values='사례번호')
        gender_counts_summary = gender_counts_summary.reindex(all_months).fillna(0).astype(int)
        gender_counts_summary['합계'] = gender_counts_summary.sum(axis=1)
        gender_counts_summary.loc['누계'] = gender_counts_summary.sum()
        st.dataframe(gender_counts_summary)

        # 연령별 인원 및 횟수
        st.markdown("---")
        st.subheader("3) 연령별 인원 및 횟수")
        age_monthly = df_counseling.groupby(['상담연월', '연령대']).agg(
            인원수=('아이디', 'nunique'),
            건수=('사례번호', 'count')
        ).reset_index()
        age_pivot_people = age_monthly.pivot(index='상담연월', columns='연령대', values='인원수')
        age_pivot_people = age_pivot_people.reindex(all_months).fillna(0).astype(int)
        age_pivot_people['합계'] = age_pivot_people.sum(axis=1)
        age_pivot_people.loc['누계'] = age_pivot_people.sum()
        st.dataframe(age_pivot_people)

        age_pivot_cases = age_monthly.pivot(index='상담연월', columns='연령대', values='건수')
        age_pivot_cases = age_pivot_cases.reindex(all_months).fillna(0).astype(int)
        age_pivot_cases['합계'] = age_pivot_cases.sum(axis=1)
        age_pivot_cases.loc['누계'] = age_pivot_cases.sum()
        st.dataframe(age_pivot_cases)

    elif menu == "🧠 심리진단 통계":
        st.header("🧠 심리진단 통계")
        diag_people = df_diagnosis.groupby(['진단연월', '진단명'])['아이디'].nunique().reset_index()
        diag_people_summary = diag_people.pivot(index='진단연월', columns='진단명', values='아이디').fillna(0).astype(int)
        diag_people_summary['합계'] = diag_people_summary.sum(axis=1)
        diag_people_summary.loc['누계'] = diag_people_summary.sum()
        st.markdown("심리진단 이용 인원")
        st.dataframe(diag_people_summary)

        diag_counts = df_diagnosis.groupby(['진단연월', '진단명'])['시행번호'].count().reset_index()
        diag_counts_summary = diag_counts.pivot(index='진단연월', columns='진단명', values='시행번호').fillna(0).astype(int)
        diag_counts_summary['합계'] = diag_counts_summary.sum(axis=1)
        diag_counts_summary.loc['누계'] = diag_counts_summary.sum()
        st.markdown("심리진단 이용 횟수")
        st.dataframe(diag_counts_summary)

    st.markdown("---")
    # 대분류 매핑
    combined_mapping_dict = {
        ('직장 내 대인관계', '상사와의 갈등'): '직장',
        ('직장 내 대인관계', '부하와의 갈등'): '직장',
        ('직장 내 대인관계', '동료와의 갈등'): '직장',
        ('직장 내 대인관계', '팀워크 및 사기 저하'): '직장',
        ('직무스트레스', '이직고민'): '직장',
        ('직무스트레스', '조직문화 부적응'): '직장',
        ('직무스트레스', '업무 스트레스'): '직장',
        ('직무스트레스', '일할 의욕의 상실'): '직장',
        ('직무스트레스', '직무재배치(TM,전근,이직)'): '직장',
        ('직무스트레스', '승진문제'): '직장',
        ('직무스트레스', '교대근무'): '직장',
        ('직무스트레스', '열악한 근무조건'): '직장',
        ('직무스트레스', '징계/감봉/직위해제'): '직장',
        ('직무스트레스', '보수에 불만'): '직장',
        ('직무스트레스', '산재 스트레스'): '직장',
        ('직무스트레스', '잔업/특근'): '직장',
        ('직무스트레스', '출퇴근 관련'): '직장',
        ('직무스트레스', '적성에 맞지 않는 업무'): '직장',
        ('직무스트레스', '성차별'): '직장',
        ('직무스트레스', '성희롱'): '직장',
        ('직무스트레스', '감정노동'): '직장',
        ('직무스트레스', '잦은 야근'): '직장',
        ('직무스트레스', '휴직후 복귀'): '직장',
        ('역량 및 경력', '리더십'): '직장',
        ('역량 및 경력', '시간관리'): '직장',
        ('역량 및 경력', '진로 및 경력계발'): '직장',
        ('역량 및 경력', '실직 or 정리해고 스트레스'): '직장',
        ('역량 및 경력', '실적압박'): '직장',
        ('역량 및 경력', '정년퇴직'): '직장',
        ('역량 및 경력', '퇴직 후 인생설계'): '직장',
        ('학교생활', '전공/진로 문제'): '학교',
        ('학교생활', '학업/학습 문제'): '학교',
        ('학교생활', '휴학 및 복학'): '학교',
        ('학교생활', '경제적 문제 (학자금)'): '학교',
        ('학교 내 대인관계', '교수와의 관계'): '학교',
        ('학교 내 대인관계', '친구와의 관계'): '학교',
        ('학교 내 대인관계', '선후배와의 관계'): '학교',
        ('부부', '부부갈등'): '가족',
        ('부부', '이혼고민'): '가족',
        ('부부', '사별'): '가족',
        ('가족', '가족갈등(부모, 형제 등)'): '가족',
        ('가족', '고부/처가 갈등'): '가족',
        ('가족', '노부모 봉양'): '가족',
        ('가족', '가정폭력'): '가족',
        ('가족', '가족사망(부모, 형제 등)'): '가족',
        ('자녀', '가족돌봄(장애보호, 요양보호)'): '가족',
        ('자녀', '자녀양육/교육'): '가족',
        ('자녀', '자녀 정서문제(우울, 불안, ADHD)'): '가족',
        ('자녀', '발달지연'): '가족',
        ('자녀', '장애자녀 돌봄'): '가족',
        ('일반', '대인관계'): '개인',
        ('일반', '삶의 의미와 목표'): '개인',
        ('일반', '성격고민'): '개인',
        ('일반', '자기계발 및 성장'): '개인',
        ('일반', '이성교제'): '개인',
        ('일반', '성생활문제'): '개인',
        ('일반', '건강 및 질병 관련 스트레스'): '개인',
        ('일반', '코로노 관련 스트레스'): '개인',
        ('일반', '종교적/영적 문제'): '개인',
        ('일반', '비만'): '개인',
        ('정신건강', '정서소진(번아웃)'): '개인',
        ('정신건강', '우울증'): '개인',
        ('정신건강', '불안'): '개인',
        ('정신건강', '공황장애'): '개인',
        ('정신건강', '불면증'): '개인',
        ('정신건강', '강박'): '개인',
        ('정신건강', '짜증/분노'): '개인',
        ('정신건강', '감정기복'): '개인',
        ('정신건강', '조울증'): '개인',
        ('정신건강', '환각/환청'): '개인',
        ('정신건강', 'PTSD'): '개인',
        ('정신건강', '자해'): '개인',
        ('정신건강', '무기력'): '개인',
        ('정신건강', '산후우울'): '개인',
        ('중독', '인터넷/게임/스마트폰'): '개인',
        ('중독', '알코올'): '개인',
        ('중독', '도박'): '개인',
        ('중독', '흡연'): '개인',
        ('자살위기', '자살생각'): '개인',
        ('자살위기', '자살계획'): '개인',
        ('자살위기', '자살시도경험'): '개인',
        ('재정 및 법률자문 등', '법률상담(변호사)'): '개인',
        ('재정 및 법률자문 등', '회계상담(회계사)'): '개인',
        ('재정 및 법률자문 등', '재무상담(재무설계사)'): '개인',
        ('재정 및 법률자문 등', '세무상담(세무사)'): '개인',
        ('재정 및 법률자문 등', '건강상담'): '개인',
    }

    def clean_str(s):
        if pd.isnull(s):
            return ""
        return str(s).strip().replace(" ", "").lower()

    combined_mapping_dict_clean = {
        (clean_str(k1), clean_str(k2)): v
        for (k1, k2), v in combined_mapping_dict.items()
    }

    def map_region(row):
        if clean_str(row['주호소1']) == '기타':
            return '기타'
        return combined_mapping_dict_clean.get(
            (clean_str(row['주호소1']), clean_str(row['하위요소1'])), None
        )

    df_counseling['영역'] = df_counseling.apply(map_region, axis=1)

    count_df = (
        df_counseling
        .groupby(['영역', '주호소1', '하위요소1'])
        .size().reset_index(name='상담건수')
        .sort_values(['영역', '주호소1', '하위요소1'])
        .reset_index(drop=True)
    )

    st.header("상담 주제별 통계")
    st.markdown("1) 개요")
    st.dataframe(count_df)

    not_mapped = df_counseling[df_counseling['영역'].isnull()]
    not_mapped = not_mapped[~(not_mapped['주호소1'].isnull() & not_mapped['하위요소1'].isnull())]
    st.write(f"매핑이 안 된 상담(미분류) 건수: {len(not_mapped)}")
    if len(not_mapped) > 0:
        st.write("미분류 건(주호소1, 하위요소1) 목록 (중복제거):")
        st.dataframe(not_mapped[['주호소1', '하위요소1']].drop_duplicates())

    area_sum_df = (
        count_df.groupby('영역')['상담건수'].sum().reset_index()
        .sort_values('상담건수', ascending=False)
    )
    area_sum_df.columns = ['영역', '영역별 상담건수 합계']
    total_row = pd.DataFrame({
        '영역': ['합계'],
        '영역별 상담건수 합계': [area_sum_df['영역별 상담건수 합계'].sum()]
    })
    area_sum_df_with_total = pd.concat([area_sum_df, total_row], ignore_index=True)
    st.markdown("2) 영역별 상담건수 합계")
    st.dataframe(area_sum_df_with_total)

    main_issue_sum_df = (
    count_df
    .groupby(['영역','주호소1'])['상담건수'].sum()
    .reset_index()
    .sort_values('상담건수', ascending=False)
    )
    main_issue_sum_df.columns = ['영역', '주호소1', '주호소1별 상담건수 합계']

    # 합계 행 추가
    total_row = pd.DataFrame([{
        '영역': '합계',
        '주호소1': '',
        '주호소1별 상담건수 합계': main_issue_sum_df['주호소1별 상담건수 합계'].sum()
    }])
    main_issue_sum_df_with_total = pd.concat([main_issue_sum_df, total_row], ignore_index=True)

    st.markdown("3) 주호소1별 상담건수 합계")
    st.dataframe(main_issue_sum_df_with_total)

    missing_count = df_counseling['주호소1'].isnull().sum()
    if missing_count > 0:
        st.warning(f"'주호소1' 열에 결측치가 {missing_count}건 있습니다. 해당 행은 분석에서 누락될 수 있습니다.")
        with st.expander("'주호소1' 결측치가 있는 상담 내역 보기"):
            st.dataframe(df_counseling[df_counseling['주호소1'].isnull()][['사례번호', '아이디', '상담실시일', '주호소1', '하위요소1']])
    else:
        st.info("'주호소1' 열에 결측치는 없습니다.")

    st.markdown("---")
    st.header("내담자별 전체 상담 이용 횟수")
    client_counts = df_counseling.groupby('아이디')['사례번호'].count().reset_index()
    client_counts.columns = ['아이디', '상담횟수']
    client_counts['아이디'] = client_counts['아이디'].apply(lambda x: str(int(float(x))) if pd.notnull(x) and str(x).replace('.', '', 1).isdigit() else str(x))
    client_counts.index = client_counts.index + 1
    client_counts.reset_index(inplace=True)
    client_counts.rename(columns={'index': 'No'}, inplace=True)

    total_row = pd.DataFrame([{
        'No': '합계',
        '아이디': f"총 {client_counts['아이디'].nunique()}명",
        '상담횟수': client_counts['상담횟수'].sum()
    }])
    client_counts_with_total = pd.concat([client_counts, total_row], ignore_index=True)
    st.dataframe(client_counts_with_total, use_container_width=True)