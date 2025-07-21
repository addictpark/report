import streamlit as st
import pandas as pd

st.title("📊 보고서용 데이터 추출 앱")

uploaded_counseling = st.file_uploader("상담 이력 엑셀 파일 업로드", type=["xlsx"], key="counseling")
uploaded_diagnosis = st.file_uploader("진단 이력 엑셀 파일 업로드", type=["xlsx"], key="diagnosis")
uploaded_topics = st.file_uploader("상담 주제 분류 엑셀 파일 업로드", type=["xlsx"], key="topics")

if uploaded_counseling and uploaded_diagnosis and uploaded_topics:
    # 1. 파일 읽기
    df_counseling = pd.read_excel(uploaded_counseling, sheet_name=0)
    df_diagnosis = pd.read_excel(uploaded_diagnosis, sheet_name=0)
    df_topics = pd.read_excel(uploaded_topics, sheet_name=0)

    # 2. 컬럼명 확인 및 안내 (문제 생길 때 참고)
    st.write("상담 이력 파일 컬럼:", df_counseling.columns.tolist())
    st.write("상담 주제 파일 컬럼:", df_topics.columns.tolist())

    # 3. 날짜 변환
    df_counseling['상담실시일'] = pd.to_datetime(df_counseling['상담실시일'], errors='coerce')
    df_diagnosis['진단실시일'] = pd.to_datetime(df_diagnosis['진단실시일'], errors='coerce')

    # 4. 추가 전처리
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

    st.markdown("서비스 이용 인원")
    st.dataframe(monthly_summary)

    st.markdown("서비스 이용 횟수")
    monthly_summary_count = pd.DataFrame()
    monthly_summary_count['월'] = sorted(set(df_counseling['상담월명'].dropna()) | set(df_diagnosis['진단월명'].dropna()))
    monthly_summary_count['심리상담'] = monthly_summary_count['월'].apply(lambda m: len(df_counseling[df_counseling['상담월명'] == m]))
    monthly_summary_count['심리진단'] = monthly_summary_count['월'].apply(lambda m: len(df_diagnosis[df_diagnosis['진단월명'] == m]))
    monthly_summary_count['합계'] = monthly_summary_count['심리상담'] + monthly_summary_count['심리진단']

    monthly_summary_count.loc[len(monthly_summary_count)] = ['누계', monthly_summary_count['심리상담'].sum(), monthly_summary_count['심리진단'].sum(), monthly_summary_count['합계'].sum()]

    st.dataframe(monthly_summary_count)

    st.header("2. 상담 통계")

    st.subheader("1) 상담유형별 인원 및 횟수")
    st.markdown("**상담유형별 이용 인원**")
    type_people = df_counseling.groupby(['상담월명', '상담유형'])['아이디'].nunique().reset_index()
    type_people_summary = type_people.pivot(index='상담월명', columns='상담유형', values='아이디').fillna(0).astype(int)
    type_people_summary['합계'] = type_people_summary.sum(axis=1)
    type_people_summary.loc['누계'] = type_people_summary.sum()

    real_type_people = df_counseling.groupby('상담유형')['아이디'].nunique()
    실계_행_people = real_type_people.reindex(type_people_summary.columns[:-1]).fillna(0).astype(int)
    실계_합계_people = pd.Series([real_type_people.sum()], index=['합계'])
    type_people_summary.loc['실계'] = pd.concat([실계_행_people, 실계_합계_people])
    st.dataframe(type_people_summary)

    st.markdown("**상담유형별 이용 횟수**")
    type_counts = df_counseling.groupby(['상담월명', '상담유형'])['사례번호'].count().reset_index()
    type_counts_summary = type_counts.pivot(index='상담월명', columns='상담유형', values='사례번호').fillna(0).astype(int)
    type_counts_summary['합계'] = type_counts_summary.sum(axis=1)
    type_counts_summary.loc['누계'] = type_counts_summary.sum()
    st.dataframe(type_counts_summary)

    st.markdown("---")

    st.subheader("2) 성별 이용 인원 및 횟수")

    st.markdown("**성별 이용 인원**")
    gender_people = df_counseling.groupby(['상담월명', '성별'])['아이디'].nunique().reset_index()
    gender_people_summary = gender_people.pivot(index='상담월명', columns='성별', values='아이디').fillna(0).astype(int)
    gender_people_summary['합계'] = gender_people_summary.sum(axis=1)
    gender_people_summary.loc['누계'] = gender_people_summary.sum()

    real_gender_people = df_counseling.groupby('성별')['아이디'].nunique()
    실계_행_gender = real_gender_people.reindex(gender_people_summary.columns[:-1]).fillna(0).astype(int)
    실계_합계_gender = pd.Series([real_gender_people.sum()], index=['합계'])
    gender_people_summary.loc['실계'] = pd.concat([실계_행_gender, 실계_합계_gender])
    st.dataframe(gender_people_summary)

    st.markdown("**성별 이용 횟수**")
    gender_counts = df_counseling.groupby(['상담월명', '성별'])['사례번호'].count().reset_index()
    gender_counts_summary = gender_counts.pivot(index='상담월명', columns='성별', values='사례번호').fillna(0).astype(int)
    gender_counts_summary['합계'] = gender_counts_summary.sum(axis=1)
    gender_counts_summary.loc['누계'] = gender_counts_summary.sum()
    st.dataframe(gender_counts_summary)

    st.markdown("---")   

    st.subheader("3) 연령별 인원 및 횟수")
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

    st.markdown("**연령별 이용 인원**")
    st.dataframe(age_pivot_people)

    st.markdown("**연령별 이용 횟수**")
    st.dataframe(age_pivot_cases)

    st.markdown("---")

    st.header("3. 심리진단 이용 인원 및 횟수")

    st.markdown("심리진단 이용 인원")
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

    st.markdown("심리진단 이용 횟수")
    diag_counts = df_diagnosis.groupby(['진단월', '진단명'])['시행번호'].count().reset_index()
    diag_counts_summary = diag_counts.pivot(index='진단월', columns='진단명', values='시행번호').fillna(0).astype(int)
    diag_counts_summary['합계'] = diag_counts_summary.sum(axis=1)
    diag_counts_summary.loc['누계'] = diag_counts_summary.sum()
    st.dataframe(diag_counts_summary)


st.header("4. 내담자별 전체 상담 이용 횟수")

# 아이디별 상담 건수 집계
client_counts = df_counseling.groupby('아이디')['사례번호'].count().reset_index()
client_counts.columns = ['아이디', '상담횟수']

# 인덱스를 1부터 시작하는 번호로 추가
client_counts.index = client_counts.index + 1
client_counts.reset_index(inplace=True)
client_counts.rename(columns={'index': 'No'}, inplace=True)

st.dataframe(client_counts)

st.info(
    "전체 기간 동안 각 내담자(아이디)별로 상담을 이용한 총 횟수를 집계합니다."
)


st.header("5. 대분류(상위영역)별 상담 건수 집계")

# 상담이력 파일의 주호소1 값
st.write("상담이력 파일 '주호소1' 값 샘플:", df_counseling['주호소1'].unique()[:20])
# 상담주제 파일의 2열(또는 실제 주호소명 열) 값
sub_col = df_topics.columns[0]  # 또는 1, 실제 확인 필요
st.write("상담주제 파일의 주호소명(2열) 값 샘플:", df_topics[sub_col].unique()[:20])

# 1. 상담 주제 파일의 2열(하위영역) → 1열(대분류) 매핑
sub_col = df_topics.columns[0]  # 2열: 주호소명 (예: 직무스트레스, 부부...)
main_col = df_topics.columns[1] # 1열: 대분류 (예: 직장, 가족...)

sub_to_main = dict(zip(df_topics[sub_col], df_topics[main_col]))

# 2. 상담 이력의 주호소1 값을 대분류로 매핑
df_counseling['대분류'] = df_counseling['주호소1'].map(sub_to_main)

# 3. 대분류별 상담 건수 집계
main_counts = df_counseling['대분류'].value_counts().reset_index()
main_counts.columns = ['대분류', '상담건수']

st.dataframe(main_counts)

st.info(
    "상담 이력의 '주호소1' 값이 상담 주제 파일의 2열(주호소명)~1열(대분류) 구조에 따라 대분류별로 몇 건인지 집계합니다.\n"
    "즉, 각 대분류(예: 직장, 가족, 개인 등)별 상담 건수를 볼 수 있습니다."
)

# 공백, 대소문자, 특수문자 전처리
def clean_text(s):
    return str(s).strip().lower().replace(" ", "")

df_counseling['주호소1_clean'] = df_counseling['주호소1'].apply(clean_text)
df_topics['sub_clean'] = df_topics[sub_col].apply(clean_text)

# 매핑: clean 기준으로
sub_to_main = dict(zip(df_topics['sub_clean'], df_topics[main_col]))

df_counseling['대분류'] = df_counseling['주호소1_clean'].map(sub_to_main)


st.write("상담주제 파일 컬럼명:", df_topics.columns.tolist())
st.write("상담주제 파일 샘플 데이터:", df_topics.head())