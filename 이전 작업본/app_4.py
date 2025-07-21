import streamlit as st
import pandas as pd

st.title("📊 보고서용 데이터 추출 프로그램")

st.info("상담 이력 데이터에서 아이디와 휴대전화번호를 삭제하고 업로드하세요 : )")

uploaded_counseling = st.file_uploader("상담 이력 엑셀 파일 업로드", type=["xlsx"], key="counseling")
uploaded_diagnosis = st.file_uploader("진단 이력 엑셀 파일 업로드", type=["xlsx"], key="diagnosis")

if uploaded_counseling and uploaded_diagnosis:
    df_counseling = pd.read_excel(uploaded_counseling, sheet_name=0)
    df_counseling.columns = df_counseling.columns.str.strip()  # 💡추가됨: 컬럼 이름 공백 제거

    # 2️⃣ ❗ 민감정보 필드 검사
    sensitive_columns = ['신청직원이름', '휴대폰번호']
    sensitive_found = [col for col in sensitive_columns if col in df_counseling.columns]

    if sensitive_found:
        st.error(f"❌ 업로드한 상담 이력 파일에 개인정보 열({', '.join(sensitive_found)})이 포함되어 있어 분석을 중단합니다.\n\n"
                 f"⚠️ 해당 열을 삭제한 후 다시 업로드해주세요.")
        st.stop()  # 앱 실행 중단

    df_diagnosis = pd.read_excel(uploaded_diagnosis, sheet_name=0)

    def missing_summary(df, df_name="데이터"):
        summary = pd.DataFrame({
            '결측치 수': df.isnull().sum(),
            '전체 행 수': len(df),
        })
        summary['결측률 (%)'] = (summary['결측치 수'] / summary['전체 행 수'] * 100).round(1)
        summary = summary[summary['결측치 수'] > 0]
        if summary.empty:
            st.success(f"✅ '{df_name}'에는 결측치가 없습니다.")
        else:
            st.warning(f"⚠️ '{df_name}'에 결측치가 있는 열이 있습니다.")
            st.dataframe(summary)

    st.header("📌 전체 결측치 요약표")

    with st.expander("🗂 상담 이력 결측치 요약"):
        missing_summary(df_counseling, df_name="상담 이력")

    with st.expander("🗂 진단 이력 결측치 요약"):
        missing_summary(df_diagnosis, df_name="진단 이력")

    # 날짜/기초 전처리
    df_counseling['상담실시일'] = pd.to_datetime(df_counseling['상담실시일'], errors='coerce')
    df_diagnosis['진단실시일'] = pd.to_datetime(df_diagnosis['진단실시일'], errors='coerce')
    df_counseling['연령대'] = (df_counseling['신청직원나이'] // 10 * 10).astype('Int64').astype(str) + '대'
    df_counseling['성별'] = df_counseling['신청직원성별'].fillna('미상')
    df_counseling['상담월'] = df_counseling['상담실시일'].dt.to_period('M')
    df_diagnosis['진단월'] = df_diagnosis['진단실시일'].dt.to_period('M')
    df_counseling['상담월명'] = df_counseling['상담실시일'].dt.strftime('%m월')
    df_diagnosis['진단월명'] = df_diagnosis['진단실시일'].dt.strftime('%m월')

    st.header("1. 운영 요약")

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

    combined_ids = pd.concat([df_counseling[['아이디']], df_diagnosis[['아이디']]])
    unique_ids = combined_ids['아이디'].drop_duplicates()
    unique_ids = unique_ids[unique_ids.notnull() & (unique_ids != "")]
    st.write("실제 인원(중복 제거):", unique_ids.tolist())
    st.write("실제 인원 수:", len(unique_ids))

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
    age_pivot_cases.loc['합계'] = age_pivot_cases.sum()


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

    st.markdown("**심리진단 이용 인원**")
    diag_people = df_diagnosis.groupby(['진단월', '진단명'])['아이디'].nunique().reset_index()
    diag_people_summary = diag_people.pivot(index='진단월', columns='진단명', values='아이디').fillna(0).astype(int)
    diag_people_summary['합계'] = diag_people_summary.sum(axis=1)
    diag_people_summary.loc['누계'] = diag_people_summary.sum()

    real_diag_people = df_diagnosis.drop_duplicates(subset=['아이디', '진단명']).groupby('진단명')['아이디'].count()
    실계_행_diag = real_diag_people.reindex(diag_people_summary.columns[:-1]).fillna(0).astype(int)
    실계_합계_diag = pd.Series([df_diagnosis['아이디'].nunique()], index=['합계'])
    diag_people_summary.loc['실계'] = pd.concat([실계_행_diag, 실계_합계_diag])

    st.dataframe(diag_people_summary)

    st.markdown("**심리진단 이용 횟수**")
    diag_counts = df_diagnosis.groupby(['진단월', '진단명'])['시행번호'].count().reset_index()
    diag_counts_summary = diag_counts.pivot(index='진단월', columns='진단명', values='시행번호').fillna(0).astype(int)
    diag_counts_summary['합계'] = diag_counts_summary.sum(axis=1)
    diag_counts_summary.loc['누계'] = diag_counts_summary.sum()
    st.dataframe(diag_counts_summary)

   


    # =============== 영역(대분류) 자동 매핑 ================
    combined_mapping_dict = {
        # 직장
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

        # 학교
        ('학교생활', '전공/진로 문제'): '학교',
        ('학교생활', '학업/학습 문제'): '학교',
        ('학교생활', '휴학 및 복학'): '학교',
        ('학교생활', '경제적 문제 (학자금)'): '학교',
        ('학교 내 대인관계', '교수와의 관계'): '학교',
        ('학교 내 대인관계', '친구와의 관계'): '학교',
        ('학교 내 대인관계', '선후배와의 관계'): '학교',
        # 가족
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

        # 개인
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
	('재정 및 법률자문 등', '세무상담(세무사'): '개인',
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



    # =============== 상담 주제별 통계 ================

    count_df = (
        df_counseling
        .groupby(['영역', '주호소1', '하위요소1'])
        .size().reset_index(name='상담건수')
        .sort_values(['영역', '주호소1', '하위요소1'])
        .reset_index(drop=True)
    )
    st.header("4. 상담 주제별 통계")
    st.markdown("**1) 개요**")
    st.dataframe(count_df)

    # 매핑 안 된(미분류) 건 확인
    not_mapped = df_counseling[df_counseling['영역'].isnull()]

    # 주호소1/하위요소1이 모두 비어있는 행 제외
    not_mapped = not_mapped[~(not_mapped['주호소1'].isnull() & not_mapped['하위요소1'].isnull())]

    st.write(f"❗️매핑이 안 된 상담(미분류) 건수: {len(not_mapped)}")
    if len(not_mapped) > 0:
        st.write("🟡 미분류 건(주호소1, 하위요소1) 목록 (중복제거):")
        st.dataframe(not_mapped[['주호소1', '하위요소1']].drop_duplicates())


    # ====== 기존 영역별 상담건수 합계 ======
    area_sum_df = (
        count_df.groupby('영역')['상담건수'].sum().reset_index()
        .sort_values('상담건수', ascending=False)
    )
    area_sum_df.columns = ['영역', '영역별 상담건수 합계']

    # 👇 '합계' 행 추가
    total_row = pd.DataFrame({
    '영역': ['합계'],
    '영역별 상담건수 합계': [area_sum_df['영역별 상담건수 합계'].sum()]
})
    area_sum_df_with_total = pd.concat([area_sum_df, total_row], ignore_index=True)

    st.markdown("**2) 영역별 상담건수 합계**")
    st.dataframe(area_sum_df_with_total)


    # ====== 주호소1별 상담건수 합계 ======
    main_issue_sum_df = (
        count_df
        .groupby(['영역','주호소1'])['상담건수'].sum()
        .reset_index()
        .sort_values('상담건수', ascending=False)
)

    main_issue_sum_df.columns = ['영역', '주호소1', '주호소1별 상담건수 합계']
    
    st.markdown("**3) 주호소1별 상담건수 합계**")
    st.dataframe(main_issue_sum_df)

    # 💡추가됨: 주호소1 결측치 확인
    missing_count = df_counseling['주호소1'].isnull().sum()
    if missing_count > 0:
        st.warning(f"⚠️ '주호소1' 열에 결측치가 {missing_count}건 있습니다. 해당 행은 분석에서 누락될 수 있습니다.")
        with st.expander("🔍 '주호소1' 결측치가 있는 상담 내역 보기"):
            st.dataframe(df_counseling[df_counseling['주호소1'].isnull()][['사례번호', '아이디', '상담실시일', '주호소1', '하위요소1']])
    else:
        st.info("✅ '주호소1' 열에 결측치는 없습니다.")



    st.header("5. 내담자별 전체 상담 이용 횟수")

    # 아이디별 상담 건수 집계
    client_counts = df_counseling.groupby('아이디')['사례번호'].count().reset_index()
    client_counts.columns = ['아이디', '상담횟수']

    # 인덱스를 1부터 시작하는 번호로 추가
    client_counts.index = client_counts.index + 1
    client_counts.reset_index(inplace=True)
    client_counts.rename(columns={'index': 'No'}, inplace=True)

    # 💡 합계 행 추가
    total_row = pd.DataFrame([{
        'No': '합계',
        '아이디': f"총 {client_counts['아이디'].nunique()}명",
        '상담횟수': client_counts['상담횟수'].sum()
}])

    client_counts_with_total = pd.concat([client_counts, total_row], ignore_index=True)

# 출력
    st.dataframe(client_counts_with_total)



