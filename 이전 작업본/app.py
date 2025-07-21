import streamlit as st
import pandas as pd

st.title("📊 상담 및 진단 통계 추출기")

uploaded_counseling = st.file_uploader("상담 이력 엑셀 파일 업로드", type=["xlsx"], key="counseling")
uploaded_diagnosis = st.file_uploader("진단 이력 엑셀 파일 업로드", type=["xlsx"], key="diagnosis")
uploaded_topics = st.file_uploader("상담 주제 분류 엑셀 파일 업로드", type=["xlsx"], key="topics")  # ✅ 추가

if uploaded_counseling and uploaded_diagnosis and uploaded_topics:  # ✅ 조건 수정
    df_counseling = pd.read_excel(uploaded_counseling, sheet_name='report')
    df_diagnosis = pd.read_excel(uploaded_diagnosis, sheet_name='diag')
    df_topics = pd.read_excel(uploaded_topics, sheet_name='topic')  # ✅ 주제 분류 시트명은 'topic'

    df_counseling['상담실시일'] = pd.to_datetime(df_counseling['상담실시일'], errors='coerce')
    df_diagnosis['진단실시일'] = pd.to_datetime(df_diagnosis['진단실시일'], errors='coerce')

    df_counseling['연령대'] = (df_counseling['신청직원나이'] // 10 * 10).astype('Int64').astype(str) + '대'
    df_counseling['성별'] = df_counseling['신청직원성별'].fillna('미상')
    df_counseling['상담월'] = df_counseling['상담실시일'].dt.to_period('M')
    df_diagnosis['진단월'] = df_diagnosis['진단실시일'].dt.to_period('M')

    st.header("1. 운영 요약")

    df_counseling['상담월명'] = df_counseling['상담실시일'].dt.strftime('%m월')
    df_diagnosis['진단월명'] = df_diagnosis['진단실시일'].dt.strftime('%m월')

    monthly_summary = pd.DataFrame()
    monthly_summary['월'] = sorted(set(df_counseling['상담월명'].dropna()) | set(df_diagnosis['진단월명'].dropna()))
    monthly_summary['심리상담'] = monthly_summary['월'].apply(lambda m: df_counseling[df_counseling['상담월명'] == m]['아이디'].nunique())
    monthly_summary['심리진단'] = monthly_summary['월'].apply(lambda m: df_diagnosis[df_diagnosis['진단월명'] == m]['아이디'].nunique())
    monthly_summary['합계'] = monthly_summary['심리상담'] + monthly_summary['심리진단']

    total_counseling_ids = df_counseling['아이디'].nunique()
    total_diagnosis_ids = df_diagnosis['아이디'].nunique()
    total_combined_ids = pd.concat([df_counseling[['아이디']], df_diagnosis[['아이디']]]).drop_duplicates()['아이디'].nunique()

    monthly_summary.loc[len(monthly_summary)] = ['누계', monthly_summary['심리상담'].sum(), monthly_summary['심리진단'].sum(), monthly_summary['합계'].sum()]
    monthly_summary.loc[len(monthly_summary)] = ['실계', total_counseling_ids, total_diagnosis_ids, total_combined_ids]

    st.subheader("월별 상담 및 진단 인원")
    st.dataframe(monthly_summary)

    st.subheader("횟수 기준")
    monthly_summary_count = pd.DataFrame()
    monthly_summary_count['월'] = sorted(set(df_counseling['상담월명'].dropna()) | set(df_diagnosis['진단월명'].dropna()))
    monthly_summary_count['심리상담'] = monthly_summary_count['월'].apply(lambda m: len(df_counseling[df_counseling['상담월명'] == m]))
    monthly_summary_count['심리진단'] = monthly_summary_count['월'].apply(lambda m: len(df_diagnosis[df_diagnosis['진단월명'] == m]))
    monthly_summary_count['합계'] = monthly_summary_count['심리상담'] + monthly_summary_count['심리진단']

    monthly_summary_count.loc[len(monthly_summary_count)] = ['누계', monthly_summary_count['심리상담'].sum(), monthly_summary_count['심리진단'].sum(), monthly_summary_count['합계'].sum()]

    st.dataframe(monthly_summary_count)

    st.header("2. 상담 통계")

    st.subheader("<인원> 월별 상담유형 인원")
    type_people = df_counseling.groupby(['상담월명', '상담유형'])['아이디'].nunique().reset_index()
    type_people_summary = type_people.pivot(index='상담월명', columns='상담유형', values='아이디').fillna(0).astype(int)
    type_people_summary['합계'] = type_people_summary.sum(axis=1)
    type_people_summary.loc['누계'] = type_people_summary.sum()

    real_type_people = df_counseling.groupby('상담유형')['아이디'].nunique()
    실계_행_people = real_type_people.reindex(type_people_summary.columns[:-1]).fillna(0).astype(int)
    실계_합계_people = pd.Series([real_type_people.sum()], index=['합계'])
    type_people_summary.loc['실계'] = pd.concat([실계_행_people, 실계_합계_people])
    st.dataframe(type_people_summary)

    st.subheader("<건수> 월별 상담유형 건수")
    type_counts = df_counseling.groupby(['상담월명', '상담유형'])['사례번호'].count().reset_index()
    type_counts_summary = type_counts.pivot(index='상담월명', columns='상담유형', values='사례번호').fillna(0).astype(int)
    type_counts_summary['합계'] = type_counts_summary.sum(axis=1)
    type_counts_summary.loc['누계'] = type_counts_summary.sum()
    st.dataframe(type_counts_summary)

    st.subheader("<인원> 월별 성별 상담 인원")
    gender_people = df_counseling.groupby(['상담월명', '성별'])['아이디'].nunique().reset_index()
    gender_people_summary = gender_people.pivot(index='상담월명', columns='성별', values='아이디').fillna(0).astype(int)
    gender_people_summary['합계'] = gender_people_summary.sum(axis=1)
    gender_people_summary.loc['누계'] = gender_people_summary.sum()

    real_gender_people = df_counseling.groupby('성별')['아이디'].nunique()
    실계_행_gender = real_gender_people.reindex(gender_people_summary.columns[:-1]).fillna(0).astype(int)
    실계_합계_gender = pd.Series([real_gender_people.sum()], index=['합계'])
    gender_people_summary.loc['실계'] = pd.concat([실계_행_gender, 실계_합계_gender])
    st.dataframe(gender_people_summary)

    st.subheader("<건수> 월별 성별 상담 건수")
    gender_counts = df_counseling.groupby(['상담월명', '성별'])['사례번호'].count().reset_index()
    gender_counts_summary = gender_counts.pivot(index='상담월명', columns='성별', values='사례번호').fillna(0).astype(int)
    gender_counts_summary['합계'] = gender_counts_summary.sum(axis=1)
    gender_counts_summary.loc['누계'] = gender_counts_summary.sum()
    st.dataframe(gender_counts_summary)

    st.subheader("월별 연령대별 상담 인원 및 건수")
    age_monthly = df_counseling.groupby(['상담월', '연령대']).agg(
        인원수=('아이디', 'nunique'),
        건수=('사례번호', 'count')
    ).reset_index()

    age_pivot_people = age_monthly.pivot(index='상담월', columns='연령대', values='인원수').fillna(0).astype(int)
    age_pivot_people['합계'] = age_pivot_people.sum(axis=1)

    age_pivot_cases = age_monthly.pivot(index='상담월', columns='연령대', values='건수').fillna(0).astype(int)
    age_pivot_cases['합계'] = age_pivot_cases.sum(axis=1)

    age_pivot_people.loc['누계'] = age_pivot_people.sum()
    real_count = df_counseling.groupby('연령대')['아이디'].nunique()
    실계_행 = real_count.reindex(age_pivot_people.columns[:-1]).fillna(0).astype(int)
    실계_합계 = pd.Series([real_count.sum()], index=['합계'])
    age_pivot_people.loc['실계'] = pd.concat([실계_행, 실계_합계])

    st.markdown("**연령대별 이용 인원(명)**")
    st.dataframe(age_pivot_people)

    st.markdown("**연령대별 이용 건수(회)**")
    st.dataframe(age_pivot_cases)

    st.header("3. 월별 진단 인원 및 건수")

    st.subheader("<인원> 월별 진단 인원")
    diag_people = df_diagnosis.groupby(['진단월', '진단명'])['아이디'].nunique().reset_index()
    diag_people_summary = diag_people.pivot(index='진단월', columns='진단명', values='아이디').fillna(0).astype(int)
    diag_people_summary['합계'] = diag_people_summary.sum(axis=1)
    diag_people_summary.loc['누계'] = diag_people_summary.sum()

    # 수정된 실계 계산
    # 진단명별: 한 명이 여러 번 받아도 진단명마다 한 번만 계산
    real_diag_people = df_diagnosis.drop_duplicates(subset=['아이디', '진단명']).groupby('진단명')['아이디'].count()

    # 진단명별 실계 행 추가
    실계_행_diag = real_diag_people.reindex(diag_people_summary.columns[:-1]).fillna(0).astype(int)

    # 전체 실계 합계 (중복 제거된 실제 인원 수)
    실계_합계_diag = pd.Series([df_diagnosis['아이디'].nunique()], index=['합계'])

    # 실계 행 삽입
    diag_people_summary.loc['실계'] = pd.concat([실계_행_diag, 실계_합계_diag])

    st.dataframe(diag_people_summary)

    st.subheader("<건수> 월별 진단 건수")
    diag_counts = df_diagnosis.groupby(['진단월', '진단명'])['시행번호'].count().reset_index()
    diag_counts_summary = diag_counts.pivot(index='진단월', columns='진단명', values='시행번호').fillna(0).astype(int)
    diag_counts_summary['합계'] = diag_counts_summary.sum(axis=1)
    diag_counts_summary.loc['누계'] = diag_counts_summary.sum()
    st.dataframe(diag_counts_summary)

    st.header("4. 상담 주호소문제 및 하위문제 건수")
    main_issues = (
        df_counseling[['주호소1', '주호소2', '주호소3']]
        .melt(value_name='주호소')
        .dropna()['주호소']
        .value_counts()
        .reset_index(name='건수')
        .rename(columns={'index': '문제'})
    )
    st.subheader("주호소 문제 건수")
    st.dataframe(main_issues)

    sub_issues = (
        df_counseling[['하위요소1', '하위요소2', '하위요소3']]
        .melt(value_name='하위문제')
        .dropna()['하위문제']
        .value_counts()
        .reset_index(name='건수')
        .rename(columns={'index': '문제'})
    )
    st.subheader("하위 문제 건수")
    st.dataframe(sub_issues)

    # 1열 카테고리 리스트 만들기
    category_list = df_topics.iloc[:,0].dropna().unique()
    
    # 각 카테고리별로 세아베스틸 파일의 주호소1 값이 몇 건인지 세기
    result = []
    for category in category_list:
        count = (df_counseling['주호소1'] == category).sum()
        result.append({'카테고리': category, '건수': count})
    result_df = pd.DataFrame(result)
    
    st.subheader("주호소1 카테고리별 건수")
    st.dataframe(result_df)
