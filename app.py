import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill, Font, Alignment
import io
import base64

# 페이지 설정 - 다크 테마
st.set_page_config(
    page_title="SNS센터 채팅분석 프로그램",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# 다크모드 + 아정당 블루 CSS
st.markdown("""
    <style>
    /* 전체 배경 다크모드 */
    .stApp {
        background-color: #0a0a0a;
        color: #ffffff;
    }
    
    /* 메인 헤더 */
    .main-header {
        background: linear-gradient(135deg, #004C99 0%, #0066CC 100%);
        padding: 2.5rem;
        border-radius: 15px;
        margin-bottom: 2rem;
        box-shadow: 0 10px 30px rgba(0, 102, 204, 0.3);
        text-align: center;
        border: 1px solid rgba(0, 102, 204, 0.5);
    }
    
    .main-header h1 {
        color: #ffffff;
        font-size: 2.5rem;
        margin-bottom: 0.5rem;
        text-shadow: 2px 2px 4px rgba(0, 0, 0, 0.3);
    }
    
    .main-header p {
        color: #b3d9ff;
        font-size: 1.1rem;
    }
    
    /* 카드 스타일 */
    .card {
        background: #1a1a1a;
        border: 1px solid #004C99;
        border-radius: 10px;
        padding: 1.5rem;
        margin-bottom: 1.5rem;
        box-shadow: 0 4px 15px rgba(0, 102, 204, 0.2);
    }
    
    .card-header {
        color: #4da6ff;
        font-size: 1.3rem;
        font-weight: bold;
        margin-bottom: 1rem;
        border-bottom: 2px solid #004C99;
        padding-bottom: 0.5rem;
    }
    
    /* 버튼 스타일 */
    .stButton > button {
        background: linear-gradient(135deg, #004C99 0%, #0066CC 100%);
        color: white;
        border: none;
        padding: 0.75rem 2rem;
        font-size: 1.1rem;
        font-weight: bold;
        border-radius: 8px;
        transition: all 0.3s ease;
        box-shadow: 0 4px 15px rgba(0, 102, 204, 0.3);
        width: 100%;
    }
    
    .stButton > button:hover {
        background: linear-gradient(135deg, #0066CC 0%, #0080ff 100%);
        box-shadow: 0 6px 20px rgba(0, 102, 204, 0.5);
        transform: translateY(-2px);
    }
    
    /* 입력 필드 스타일 */
    .stTextInput > div > div > input,
    .stTextArea > div > div > textarea,
    .stDateInput > div > div > input {
        background-color: #2a2a2a;
        color: #ffffff;
        border: 1px solid #004C99;
        border-radius: 5px;
    }
    
    .stTextInput > div > div > input:focus,
    .stTextArea > div > div > textarea:focus {
        border-color: #0066CC;
        box-shadow: 0 0 0 2px rgba(0, 102, 204, 0.2);
    }
    
    /* 파일 업로더 스타일 */
    .stFileUploader > div {
        background-color: #1a1a1a;
        border: 2px dashed #004C99;
        border-radius: 10px;
        padding: 2rem;
        transition: all 0.3s ease;
    }
    
    .stFileUploader > div:hover {
        border-color: #0066CC;
        background-color: #262626;
        box-shadow: 0 4px 15px rgba(0, 102, 204, 0.2);
    }
    
    /* 로그 박스 스타일 */
    .log-box {
        background-color: #0d0d0d;
        border: 1px solid #004C99;
        border-radius: 8px;
        padding: 1rem;
        font-family: 'Courier New', monospace;
        color: #4da6ff;
        max-height: 400px;
        overflow-y: auto;
    }
    
    /* 정보 박스 스타일 */
    .info-box {
        background: linear-gradient(135deg, #1a1a1a 0%, #262626 100%);
        border-left: 4px solid #0066CC;
        padding: 1.2rem;
        border-radius: 8px;
        margin: 1rem 0;
        box-shadow: 0 2px 10px rgba(0, 102, 204, 0.2);
    }
    
    .info-box h4 {
        color: #4da6ff;
        margin-bottom: 0.5rem;
    }
    
    .info-box ul {
        color: #b3d9ff;
    }
    
    /* Progress bar 스타일 */
    .stProgress > div > div > div > div {
        background-color: #0066CC;
    }
    
    /* Success/Error/Warning 메시지 스타일 */
    .stSuccess {
        background-color: rgba(0, 102, 204, 0.1);
        border: 1px solid #0066CC;
        color: #4da6ff;
    }
    
    .stError {
        background-color: rgba(255, 0, 0, 0.1);
        border: 1px solid #ff4444;
    }
    
    .stWarning {
        background-color: rgba(255, 193, 7, 0.1);
        border: 1px solid #ffc107;
    }
    
    /* 탭 스타일 */
    .stTabs [data-baseweb="tab-list"] {
        background-color: #1a1a1a;
        border-bottom: 2px solid #004C99;
    }
    
    .stTabs [data-baseweb="tab"] {
        color: #b3d9ff;
        background-color: transparent;
    }
    
    .stTabs [aria-selected="true"] {
        background-color: #004C99;
        color: white;
    }
    
    /* Divider 스타일 */
    hr {
        border-color: #004C99;
        opacity: 0.3;
    }
    
    /* 라벨 스타일 */
    label {
        color: #4da6ff !important;
        font-weight: 500;
    }
    
    /* 다운로드 버튼 특별 스타일 */
    .download-button {
        background: linear-gradient(135deg, #00cc66 0%, #00ff88 100%);
        animation: pulse 2s infinite;
    }
    
    @keyframes pulse {
        0% {
            box-shadow: 0 0 0 0 rgba(0, 204, 102, 0.7);
        }
        70% {
            box-shadow: 0 0 0 10px rgba(0, 204, 102, 0);
        }
        100% {
            box-shadow: 0 0 0 0 rgba(0, 204, 102, 0);
        }
    }
    
    /* 스크롤바 스타일 */
    ::-webkit-scrollbar {
        width: 10px;
        height: 10px;
    }
    
    ::-webkit-scrollbar-track {
        background: #1a1a1a;
    }
    
    ::-webkit-scrollbar-thumb {
        background: #004C99;
        border-radius: 5px;
    }
    
    ::-webkit-scrollbar-thumb:hover {
        background: #0066CC;
    }
    </style>
""", unsafe_allow_html=True)

class CollaborationAnalyzer:
    def __init__(self):
        self.initialize_session_state()

    def initialize_session_state(self):
        if 'analysis_complete' not in st.session_state:
            st.session_state.analysis_complete = False
        if 'result_file' not in st.session_state:
            st.session_state.result_file = None
        if 'log_messages' not in st.session_state:
            st.session_state.log_messages = []
        if 'processing' not in st.session_state:
            st.session_state.processing = False

    def log(self, message):
        timestamp = datetime.now().strftime('%H:%M:%S')
        log_entry = f"[{timestamp}] {message}"
        st.session_state.log_messages.append(log_entry)

    @st.cache_data(show_spinner=False)
    def load_excel_cached(_self, file_bytes, file_name):
        """엑셀 파일을 캐시하여 로드"""
        return pd.read_excel(io.BytesIO(file_bytes), sheet_name=None, engine='openpyxl')

    def load_and_process_data(self, file, start_date_str, end_date_str):
        try:
            self.log("📁 엑셀 파일 로딩 중...")
            
            # 파일을 바이트로 읽어서 캐시 활용
            file_bytes = file.read()
            file_name = file.name
            all_sheets = self.load_excel_cached(file_bytes, file_name)
            
            required_sheets = ['UserChat data', 'Message data', 'Manager data']
            sheet_data = {core_name: [] for core_name in required_sheets}
            
            for sheet_name, df in all_sheets.items():
                for core_name in required_sheets:
                    if core_name in sheet_name:
                        sheet_data[core_name].append(df)
            
            if not all(sheet_data.values()):
                st.error("❌ 필수 시트를 찾을 수 없습니다.")
                return None

            self.log("🔄 시트 통합 및 중복 제거 중...")
            user_chat_df = pd.concat(sheet_data['UserChat data'], ignore_index=True).drop_duplicates(subset=['id'])
            message_df = pd.concat(sheet_data['Message data'], ignore_index=True).drop_duplicates(subset=['chatId', 'personId', 'createdAt', 'plainText'])
            manager_df = pd.concat(sheet_data['Manager data'], ignore_index=True).drop_duplicates(subset=['id'])

            self.log("🧹 데이터 정제 및 병합 중...")
            def clean_id(series):
                return series.astype(str).str.strip().str.replace(r'\.0$', '', regex=True)

            user_chat_df['id'] = clean_id(user_chat_df['id'])
            user_chat_df['assigneeId'] = clean_id(user_chat_df['assigneeId'])
            message_df['chatId'] = clean_id(message_df['chatId'])
            message_df['personId'] = clean_id(message_df['personId'])
            manager_df['id'] = clean_id(manager_df['id'])
            
            # 날짜 필터링
            start_ts = pd.to_datetime(start_date_str)
            end_ts = pd.to_datetime(end_date_str) + pd.DateOffset(days=1)

            user_chat_df['firstOpenedAt'] = pd.to_datetime(user_chat_df['firstOpenedAt'], errors='coerce')
            user_chat_df.dropna(subset=['firstOpenedAt'], inplace=True)
            filtered_user_chat_df = user_chat_df[(user_chat_df['firstOpenedAt'] >= start_ts) & (user_chat_df['firstOpenedAt'] < end_ts)]
            self.log(f"📊 담당 상담 집계 대상: {len(filtered_user_chat_df):,}개 상담")

            message_df['createdAt'] = pd.to_datetime(message_df['createdAt'], errors='coerce')
            message_df.dropna(subset=['createdAt'], inplace=True)
            filtered_message_df = message_df[(message_df['createdAt'] >= start_ts) & (message_df['createdAt'] < end_ts)]
            self.log(f"💬 메시지 분석 대상: {len(filtered_message_df):,}개 메시지")
            
            if filtered_message_df.empty:
                st.error("선택된 기간 내에 데이터가 없습니다.")
                return None
            
            merged_df = pd.merge(filtered_message_df, user_chat_df[['id', 'assigneeId']], left_on='chatId', right_on='id', how='left').dropna(subset=['assigneeId'])
            merged_df = pd.merge(merged_df, manager_df[['id', 'name']], left_on='personId', right_on='id', how='left', suffixes=('', '_manager')).rename(columns={'name': 'authorName'}).dropna(subset=['authorName'])
            
            self.log(f"✅ 총 {len(merged_df):,}개의 유효 메시지 레코드 처리됨")
            return {'merged': merged_df, 'user_chat': filtered_user_chat_df, 'manager': manager_df}

        except Exception as e:
            st.error(f"데이터 처리 중 오류: {str(e)}")
            return None

    def auto_adjust_columns(self, writer, sheet_name, df, min_width=15):
        worksheet = writer.sheets[sheet_name]
        for idx, col in enumerate(df.columns):
            max_len = len(str(col)) + 4
            try:
                max_in_col = df[col].astype(str).map(len).max()
                if not pd.isna(max_in_col):
                    max_len = max(max_len, int(max_in_col) + 2)
            except Exception:
                pass
            adjusted_width = max(min_width, max_len)
            worksheet.column_dimensions[get_column_letter(idx + 1)].width = adjusted_width

    def style_header(self, writer, sheet_name, df):
        worksheet = writer.sheets[sheet_name]
        header_fill = PatternFill(start_color="004C99", end_color="004C99", fill_type="solid")
        header_font = Font(bold=True, color="FFFFFF")
        center_align = Alignment(horizontal='center', vertical='center')
        for col_idx in range(1, df.shape[1] + 1):
            cell = worksheet.cell(row=1, column=col_idx)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = center_align

    @st.cache_data(show_spinner=False)
    def process_analysis(_self, df_merged, user_chat_df, manager_df, managers_list, exclusion_list):
        """분석 로직을 캐시하여 처리"""
        # 기존 create_output_excel의 분석 로직 부분만 분리
        return _self._process_analysis_internal(df_merged, user_chat_df, manager_df, managers_list, exclusion_list)

    def _process_analysis_internal(self, df_merged, user_chat_df, manager_df, managers_list, exclusion_list):
        """실제 분석 로직"""
        manager_data = df_merged[df_merged['authorName'].isin(managers_list)]
        agent_data = df_merged[(~df_merged['authorName'].isin(managers_list)) & (~df_merged['authorName'].isin(exclusion_list))]
        
        all_agents = pd.DataFrame({'authorName': agent_data['authorName'].unique()})
        agent_non_assignee = agent_data[agent_data['personId'] != agent_data['assigneeId']]
        
        if not agent_non_assignee.empty:
            collaborated_chats = agent_non_assignee.groupby('authorName')['chatId'].nunique()
        else:
            collaborated_chats = pd.Series(dtype='int64', name='chatId', index=pd.Index([], name='authorName'))

        total_chats = agent_data.groupby('authorName')['chatId'].nunique()
        hir_summary = (collaborated_chats / total_chats).reset_index(name='HIR').fillna(0)
        
        total_msg_counts = agent_data.groupby('authorName').size().reset_index(name='total_messages')
        
        filter_df = pd.merge(all_agents, hir_summary, on='authorName', how='left')
        filter_df = pd.merge(filter_df, total_msg_counts, on='authorName', how='left')
        filter_df.fillna(0, inplace=True)
        
        return {
            'manager_data': manager_data,
            'agent_data': agent_data,
            'filter_df': filter_df,
            'total_msg_counts': total_msg_counts
        }

    def create_output_excel(self, processed_data, start_date_str, end_date_str, managers_list, exclusion_list):
        df_merged = processed_data['merged']
        user_chat_df = processed_data['user_chat']
        manager_df = processed_data['manager']
        
        output = io.BytesIO()
        
        if exclusion_list:
            self.log(f"🚫 제외 명단: {', '.join(exclusion_list)}")

        # 캐시된 분석 결과 사용
        analysis_result = self.process_analysis(
            df_merged, user_chat_df, manager_df, managers_list, exclusion_list
        )
        
        manager_data = analysis_result['manager_data']
        agent_data = analysis_result['agent_data']
        filter_df = analysis_result['filter_df']
        
        start_date = datetime.strptime(start_date_str, "%Y-%m-%d")
        end_date = datetime.strptime(end_date_str, "%Y-%m-%d")
        analysis_days = (end_date - start_date).days + 1
        min_msg_threshold = analysis_days * 10
        self.log(f"📅 분석 기간: {analysis_days}일, 최소 메시지 수: {min_msg_threshold}개")

        self.log(f"👥 필터링 전 상담사 수: {len(filter_df)}명")
        filtered_authors_df = filter_df[
            (filter_df['HIR'] > 0) & (filter_df['HIR'] < 1) & 
            (filter_df['total_messages'] > 10) &
            (filter_df['total_messages'] >= min_msg_threshold)
        ]
        self.log(f"👥 필터링 후 상담사 수: {len(filtered_authors_df)}명")
        
        if filtered_authors_df.empty:
            self.log("⚠️ 필터링 후 분석 대상 상담사가 없습니다.")
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                pd.DataFrame({'알림': ['필터 조건에 해당하는 상담사가 없어 데이터를 생성할 수 없습니다.']}).to_excel(writer, sheet_name='결과 없음', index=False)
            output.seek(0)
            return output

        authors_to_keep = filtered_authors_df['authorName'].unique()
        
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            def analyze_group(group_name, group_merged_df, user_chat_df_local, manager_df_local):
                self.log(f"🔍 '{group_name}' 그룹 분석 중...")
                if group_merged_df.empty and group_name == "상담사":
                    return pd.DataFrame(), pd.DataFrame()

                authors = pd.DataFrame({'authorName': pd.concat([group_merged_df['authorName'], manager_df_local.loc[manager_df_local['id'].isin(user_chat_df_local['assigneeId']), 'name']]).unique()})
                
                group_non_assignee_df = group_merged_df[group_merged_df['personId'] != group_merged_df['assigneeId']].copy()
                group_assignee_df = group_merged_df[group_merged_df['personId'] == group_merged_df['assigneeId']].copy()

                metrics_df = authors.copy()
                if not group_non_assignee_df.empty:
                    hir_s = (group_non_assignee_df.groupby('authorName')['chatId'].nunique() / group_merged_df.groupby('authorName')['chatId'].nunique()).reset_index(name='HIR').fillna(0)
                    chat_lengths = group_merged_df.groupby('chatId').size().to_dict()
                    group_non_assignee_df['chat_length'] = group_non_assignee_df['chatId'].map(chat_lengths)
                    iif_s = group_non_assignee_df.groupby('authorName')['chat_length'].sum().reset_index(name='IIF')
                    
                    core_keywords = ['월요금', '사은품', '위약금', '결합', '설치일', '설치비', '약정', '지원금', '할인', '통신사', '요금제', '인터넷', '휴대폰']
                    keyword_pattern = '|'.join(core_keywords)
                    group_non_assignee_df['cis_flag'] = group_non_assignee_df['plainText'].str.contains(keyword_pattern, na=False).astype(int)
                    cis_s = group_non_assignee_df.groupby('authorName')['cis_flag'].sum().reset_index(name='CIS')

                    group_non_assignee_df['msg_length'] = group_non_assignee_df['plainText'].str.len()
                    dls_s = group_non_assignee_df.groupby('authorName')['msg_length'].mean().reset_index(name='DLS')
                    
                    group_non_assignee_df['als_flag'] = group_non_assignee_df['plainText'].str.contains('https://form.ajd.co.kr/', na=False).astype(int)
                    als_s = group_non_assignee_df.groupby('authorName')['als_flag'].sum().reset_index(name='ALS')
                    als_s['ALS'] *= 10

                    for df in [hir_s, iif_s, cis_s, dls_s, als_s]:
                        metrics_df = pd.merge(metrics_df, df, on='authorName', how='left')
                metrics_df.fillna(0, inplace=True)

                assigned_stats = group_assignee_df.groupby('authorName').agg(
                    담당_메시지_수=('chatId', 'size'),
                    담당_글자_수=('plainText', lambda x: x.str.len().sum())
                ).reset_index()
                
                help_stats = group_non_assignee_df.groupby('authorName').agg(
                    도움_상담_수=('chatId', 'nunique'),
                    도움_메시지_수=('chatId', 'size'),
                    도움_글자_수=('plainText', lambda x: x.str.len().sum())
                ).reset_index()
                
                author_ids = manager_df_local[manager_df_local['name'].isin(authors['authorName'])]['id']
                assigned_chat_counts = user_chat_df_local[user_chat_df_local['assigneeId'].isin(author_ids)].groupby('assigneeId').agg(
                    담당_상담_수=('id', 'nunique')
                ).reset_index()
                assigned_chat_counts = pd.merge(assigned_chat_counts, manager_df_local[['id', 'name']], left_on='assigneeId', right_on='id').rename(columns={'name': 'authorName'})

                def get_base_score(df):
                    if df.empty:
                        return pd.DataFrame(columns=['authorName', 'base_score'])
                    message_counts = df['plainText'].value_counts()
                    df['score'] = df['plainText'].map(lambda x: np.log1p(len(df) / message_counts.get(x, 1)))
                    return df.groupby('authorName')['score'].sum().reset_index(name='base_score')

                base_help_score = get_base_score(group_non_assignee_df)
                base_assigned_score = get_base_score(group_assignee_df)

                summary_df = pd.merge(authors, assigned_chat_counts[['authorName', '담당_상담_수']], on='authorName', how='left')
                summary_df = pd.merge(summary_df, assigned_stats, on='authorName', how='left')
                summary_df = pd.merge(summary_df, help_stats, on='authorName', how='left')
                summary_df = pd.merge(summary_df, base_help_score.rename(columns={'base_score': 'base_help_score'}), on='authorName', how='left')
                summary_df = pd.merge(summary_df, base_assigned_score.rename(columns={'base_score': 'base_assigned_score'}), on='authorName', how='left')
                summary_df = pd.merge(summary_df, metrics_df, on='authorName', how='left')
                summary_df.fillna(0, inplace=True)

                metrics_for_correction = ['ALS']
                for col in metrics_for_correction:
                    min_val, max_val = summary_df[col].min(), summary_df[col].max()
                    if max_val > min_val:
                        summary_df[f'norm_{col}'] = (summary_df[col] - min_val) / (max_val - min_val)
                    else:
                        summary_df[f'norm_{col}'] = 0
                
                als_weight = 3
                summary_df['help_correction'] = summary_df.get('norm_ALS', 0) * als_weight
                summary_df['assigned_correction'] = 0
                
                summary_df['도움_정성_점수'] = summary_df.get('base_help_score', 0) * (1 + summary_df['help_correction'])
                summary_df['담당_정성_점수'] = summary_df.get('base_assigned_score', 0) * (1 + summary_df['assigned_correction'])

                total_help_score = summary_df['도움_정성_점수'].sum()
                total_assigned_score = summary_df['담당_정성_점수'].sum()
                if total_assigned_score > 0 and total_help_score > 0:
                    ratio_factor = (total_help_score / 3) / total_assigned_score
                    summary_df['담당_정성_점수'] *= ratio_factor
                
                summary_df['총_정성_점수'] = summary_df['도움_정성_점수'] + summary_df['담당_정성_점수']

                final_cols = ['authorName', '담당_상담_수', '담당_메시지_수', '담당_글자_수', '도움_상담_수', '도움_메시지_수', '도움_글자_수', '담당_정성_점수', '도움_정성_점수', '총_정성_점수']
                final_summary = summary_df[final_cols].rename(columns={
                    'authorName': '상담사명', '담당_상담_수': '담당 상담 수', '담당_메시지_수': '담당 메시지 수',
                    '담당_글자_수': '담당 글자 수', '도움_상담_수': '도움 상담 수', '도움_메시지_수': '도움 메시지 수',
                    '도움_글자_수': '도움 글자 수', '담당_정성_점수': '담당 정성 점수', '도움_정성_점수': '도움 정성 점수',
                    '총_정성_점수': '총 정성 점수'
                })
                
                metrics_cols = ['authorName', 'HIR', 'IIF', 'CIS', 'DLS', 'ALS']
                final_metrics = summary_df[metrics_cols].rename(columns={
                    'authorName': '상담사명', 'HIR': '도움 개입률', 'IIF': '개입 영향력 계수',
                    'CIS': '콘텐츠 정보 점수', 'DLS': '언어 깊이 점수', 'ALS': '신청서 링크 점수'
                })

                return final_summary.round(2), final_metrics.round(2)

            # 그룹별 분석 실행
            agent_summary_df, agent_metrics_df = analyze_group(
                "상담사",
                agent_data[agent_data['authorName'].isin(authors_to_keep)],
                user_chat_df[user_chat_df['assigneeId'].isin(manager_df[manager_df['name'].isin(authors_to_keep)]['id'])],
                manager_df
            )
            
            manager_summary_df, _ = analyze_group(
                "관리자",
                manager_data,
                user_chat_df[user_chat_df['assigneeId'].isin(manager_df[manager_df['name'].isin(managers_list)]['id'])],
                manager_df
            )
            
            # 엑셀 시트 생성
            agent_summary_df = agent_summary_df.sort_values(by='총 정성 점수', ascending=False)
            agent_summary_df.to_excel(writer, sheet_name='채팅분석_요약', index=False)
            self.style_header(writer, '채팅분석_요약', agent_summary_df)
            self.auto_adjust_columns(writer, '채팅분석_요약', agent_summary_df)
            self.log("✅ '채팅분석_요약' 시트 생성 완료")
            
            if not manager_summary_df.empty:
                manager_summary_df = manager_summary_df.sort_values(by='총 정성 점수', ascending=False)
                manager_summary_df.to_excel(writer, sheet_name='관리자_분석', index=False)
                self.style_header(writer, '관리자_분석', manager_summary_df)
                self.auto_adjust_columns(writer, '관리자_분석', manager_summary_df)
                self.log("✅ '관리자_분석' 시트 생성 완료")
            
            agent_metrics_df = agent_metrics_df.sort_values(by='신청서 링크 점수', ascending=False)
            agent_metrics_df.to_excel(writer, sheet_name='채팅분석_지표', index=False)
            self.style_header(writer, '채팅분석_지표', agent_metrics_df)
            self.auto_adjust_columns(writer, '채팅분석_지표', agent_metrics_df)
            self.log("✅ '채팅분석_지표' 시트 생성 완료")

            # 스코어보드 생성
            self.log("📊 '스코어보드' 시트 생성 중...")
            workbook = writer.book
            worksheet = workbook.create_sheet('스코어보드')
            
            sheets = writer.book.sheetnames
            writer.book.move_sheet('스코어보드', offset=-sheets.index('스코어보드'))
            writer.book.move_sheet('채팅분석_요약', offset=-len(writer.book.sheetnames))
            
            title_font = Font(bold=True, color="FFFFFF", size=11)
            header_font = Font(bold=True, size=10)
            center_align = Alignment(horizontal='center', vertical='center')

            top5_header_fill = PatternFill(start_color="004C99", end_color="004C99", fill_type="solid")
            bottom5_header_fill = PatternFill(start_color="CC0000", end_color="CC0000", fill_type="solid")
            top5_data_fill = PatternFill(start_color="E6F2FF", end_color="E6F2FF", fill_type="solid")
            bottom5_data_fill = PatternFill(start_color="FFE6E6", end_color="FFE6E6", fill_type="solid")
            
            metrics_to_rank = {
                '담당 상담 수': '담당 상담 수', '담당 메시지 수': '담당 메시지 수', '담당 글자 수': '담당 글자 수',
                '도움 상담 수': '도움 상담 수', '도움 메시지 수': '도움 메시지 수', '도움 글자 수': '도움 글자 수',
                '담당 정성 점수': '담당 정성 점수', '도움 정성 점수': '도움 정성 점수', '총 정성 점수': '총 정성 점수'
            }
            
            current_row = 1
            for col, title in metrics_to_rank.items():
                top_5 = agent_summary_df.nlargest(5, col)[['상담사명', col]]
                bottom_5 = agent_summary_df.nsmallest(5, col)[['상담사명', col]]

                worksheet.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=5)
                cell = worksheet.cell(row=current_row, column=1, value=f'--- {title} 순위 ---')
                cell.font = header_font
                cell.alignment = center_align
                current_row += 1
                
                headers = ['Top 5', '점수', '', 'Bottom 5', '점수']
                fills = [top5_header_fill, top5_header_fill, None, bottom5_header_fill, bottom5_header_fill]
                
                for c_idx, value in enumerate(headers):
                    if value:
                        cell = worksheet.cell(row=current_row, column=c_idx + 1, value=value)
                        cell.font = title_font
                        cell.alignment = center_align
                        cell.fill = fills[c_idx]
                current_row += 1
                
                for i in range(5):
                    row_to_write = current_row + i
                    if i < len(top_5):
                        cell_name = worksheet.cell(row=row_to_write, column=1, value=top_5.iloc[i, 0])
                        cell_score = worksheet.cell(row=row_to_write, column=2, value=top_5.iloc[i, 1])
                        cell_name.fill = top5_data_fill
                        cell_score.fill = top5_data_fill
                    
                    if i < len(bottom_5):
                        cell_name = worksheet.cell(row=row_to_write, column=4, value=bottom_5.iloc[i, 0])
                        cell_score = worksheet.cell(row=row_to_write, column=5, value=bottom_5.iloc[i, 1])
                        cell_name.fill = bottom5_data_fill
                        cell_score.fill = bottom5_data_fill
                
                current_row += 6

            worksheet.column_dimensions[get_column_letter(1)].width = 20
            worksheet.column_dimensions[get_column_letter(2)].width = 15
            worksheet.column_dimensions[get_column_letter(3)].width = 5
            worksheet.column_dimensions[get_column_letter(4)].width = 20
            worksheet.column_dimensions[get_column_letter(5)].width = 15
            self.log("✅ '스코어보드' 시트 생성 완료")

            # 상세 메시지 내역 시트 (간소화)
            non_assignee_df = df_merged[df_merged['personId'] != df_merged['assigneeId']]
            assignee_df = df_merged[df_merged['personId'] == df_merged['assigneeId']]

            details_collab = non_assignee_df[non_assignee_df['authorName'].isin(authors_to_keep)]
            if not details_collab.empty:
                details_collab = details_collab[['authorName', 'chatId', 'createdAt', 'plainText']].copy()
                details_collab.rename(columns={
                    'authorName': '상담사명', 'chatId': '채팅방 ID',
                    'createdAt': '메시지 작성일시', 'plainText': '메시지 원문'
                }, inplace=True)
                details_collab = details_collab.sort_values(by=['상담사명', '메시지 작성일시'])
                details_collab.to_excel(writer, sheet_name='상세 메시지 내역_도움', index=False)
                self.style_header(writer, '상세 메시지 내역_도움', details_collab)
                self.auto_adjust_columns(writer, '상세 메시지 내역_도움', details_collab)

            details_assignee = assignee_df[assignee_df['authorName'].isin(authors_to_keep)]
            if not details_assignee.empty:
                details_assignee = details_assignee[['authorName', 'chatId', 'createdAt', 'plainText']].copy()
                details_assignee.rename(columns={
                    'authorName': '상담사명', 'chatId': '채팅방 ID',
                    'createdAt': '메시지 작성일시', 'plainText': '메시지 원문'
                }, inplace=True)
                details_assignee = details_assignee.sort_values(by=['상담사명', '메시지 작성일시'])
                details_assignee.to_excel(writer, sheet_name='상세 메시지 내역_담당자', index=False)
                self.style_header(writer, '상세 메시지 내역_담당자', details_assignee)
                self.auto_adjust_columns(writer, '상세 메시지 내역_담당자', details_assignee)

            # 도움말 시트
            self.log("📝 '도움말' 시트 생성 중...")
            help_data = {
                '구분': ['정량 지표', '정량 지표', '정성 지표', '정성 지표', '정성 지표', '정성 점수', '정성 점수', '정성 점수'],
                '지표명': ['HIR (도움 개입률)', 'IIF (개입 영향력 계수)', 'CIS (콘텐츠 정보 점수)',
                          'DLS (언어 깊이 점수)', 'ALS (신청서 링크 점수)', '도움 정성 점수',
                          '담당 정성 점수', '총 정성 점수'],
                '정의': [
                    '한 상담사가 참여한 전체 상담 중, 협업자로 참여한 상담의 비율',
                    '단순 메시지 수를 넘어, 얼마나 길고 복잡한 대화에 개입했는지를 가중치로 평가',
                    '도움 메시지 중, 사전에 정의된 핵심 상품 키워드를 포함한 메시지의 개수',
                    '도움 메시지 1개당 평균 글자 길이',
                    '도움 메시지에서 신청서 링크를 발송한 횟수에 기반한 점수',
                    '도움 메시지의 희소성과 ALS 보정치를 반영한 질적 기여도 점수',
                    '담당 메시지의 희소성 점수를 반영하고, 도움 정성 점수와의 비율을 조정한 질적 기여도 점수',
                    '도움 정성 점수와 담당 정성 점수의 합산'
                ],
                '산식': [
                    '협업 참여 상담 건수 / 총 참여 상담 건수',
                    'Σ (협업 참여한 각 상담 건의 전체 메시지 수)',
                    'Σ (도움 메시지 내 핵심 키워드 포함 개수)',
                    '도움 메시지 총 글자 수 / 도움 메시지 총 개수',
                    'Σ (신청서 링크 발송 횟수) * 10점',
                    '기본 점수 * (1 + (정규화 ALS * 3))',
                    '기본 점수 * 비율 조정 계수',
                    '도움 정성 점수 + 담당 정성 점수'
                ]
            }
            help_df = pd.DataFrame(help_data)
            help_df.to_excel(writer, sheet_name='도움말', index=False)
            self.style_header(writer, '도움말', help_df)
            self.auto_adjust_columns(writer, '도움말', help_df)
            self.log("✅ '도움말' 시트 생성 완료")

        output.seek(0)
        self.log("🎉 결과 파일 생성 완료!")
        return output

def main():
    # 헤더
    st.markdown("""
        <div class="main-header">
            <h1>📊 SNS센터 채팅분석 프로그램 v1.9</h1>
            <p>채팅 데이터를 분석하여 상담사의 협업 성과를 평가합니다</p>
        </div>
    """, unsafe_allow_html=True)

    analyzer = CollaborationAnalyzer()

    # 메인 페이지에 탭 구성
    tab1, tab2, tab3 = st.tabs(["📤 분석 설정", "📊 분석 실행", "📥 결과 다운로드"])
    
    with tab1:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.markdown('<div class="card-header">📁 1단계: 파일 업로드</div>', unsafe_allow_html=True)
        
        uploaded_file = st.file_uploader(
            "엑셀 파일을 선택하세요",
            type=['xlsx'],
            help="UserChat data, Message data, Manager data 시트가 포함된 엑셀 파일",
            key="file_uploader"
        )
        
        if uploaded_file:
            st.success(f"✅ 파일 업로드 완료: {uploaded_file.name}")
        st.markdown('</div>', unsafe_allow_html=True)
        
        # 2개 컬럼으로 날짜와 관리자 설정
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown('<div class="card">', unsafe_allow_html=True)
            st.markdown('<div class="card-header">📅 2단계: 분석 기간 설정</div>', unsafe_allow_html=True)
            
            start_date = st.date_input(
                "시작일",
                value=datetime(2025, 7, 1),
                format="YYYY-MM-DD",
                key="start_date"
            )
            
            end_date = st.date_input(
                "종료일",
                value=datetime.now() - timedelta(days=1),
                format="YYYY-MM-DD",
                key="end_date"
            )
            
            # 기간 계산 표시
            if start_date and end_date:
                days = (end_date - start_date).days + 1
                st.info(f"📊 분석 기간: {days}일")
            st.markdown('</div>', unsafe_allow_html=True)
        
        with col2:
            st.markdown('<div class="card">', unsafe_allow_html=True)
            st.markdown('<div class="card-header">👥 3단계: 인원 설정</div>', unsafe_allow_html=True)
            
            managers = st.text_area(
                "관리자 목록 (쉼표로 구분)",
                value="이민주, 이종민, 윤도우리, 김시진, 손진우",
                height=60,
                key="managers"
            )
            
            exclusions = st.text_area(
                "제외할 이름 (쉼표로 구분, 선택사항)",
                value="채주은, 정용욱, 한승윤, 김종현",
                height=60,
                key="exclusions"
            )
            st.markdown('</div>', unsafe_allow_html=True)
    
    with tab2:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.markdown('<div class="card-header">🚀 분석 실행</div>', unsafe_allow_html=True)
        
        # 설정 요약
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("📁 파일", "✅ 업로드됨" if uploaded_file else "❌ 미업로드")
        with col2:
            if 'start_date' in locals() and 'end_date' in locals():
                days = (end_date - start_date).days + 1
                st.metric("📅 분석 기간", f"{days}일")
            else:
                st.metric("📅 분석 기간", "미설정")
        with col3:
            if 'managers' in locals():
                manager_count = len([m.strip() for m in managers.split(',') if m.strip()])
                st.metric("👥 관리자 수", f"{manager_count}명")
            else:
                st.metric("👥 관리자 수", "미설정")
        
        st.divider()
        
        # 분석 실행 버튼
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            analyze_button = st.button(
                "🎯 분석 시작",
                type="primary",
                use_container_width=True,
                disabled=not uploaded_file
            )
        
        # 진행 상황 표시
        if analyze_button:
            if not uploaded_file:
                st.error("⚠️ 파일을 먼저 업로드해주세요!")
            else:
                st.session_state.log_messages = []
                st.session_state.processing = True
                
                progress_bar = st.progress(0)
                status_text = st.empty()
                log_container = st.container()
                
                with st.spinner("분석 중입니다... 잠시만 기다려주세요 ⏳"):
                    # 진행률 업데이트
                    progress_bar.progress(10)
                    status_text.text("📝 설정 확인 중...")
                    
                    # 관리자 및 제외 명단 파싱
                    managers_list = [name.strip() for name in managers.split(',') if name.strip()]
                    exclusion_list = [name.strip() for name in exclusions.split(',') if name.strip()]
                    
                    progress_bar.progress(30)
                    status_text.text("📊 데이터 로딩 중...")
                    
                    # 데이터 처리
                    processed_data = analyzer.load_and_process_data(
                        uploaded_file,
                        start_date.strftime("%Y-%m-%d"),
                        end_date.strftime("%Y-%m-%d")
                    )
                    
                    progress_bar.progress(60)
                    status_text.text("🔍 분석 수행 중...")
                    
                    if processed_data:
                        # 결과 생성
                        progress_bar.progress(80)
                        status_text.text("📁 결과 파일 생성 중...")
                        
                        result_file = analyzer.create_output_excel(
                            processed_data,
                            start_date.strftime("%Y-%m-%d"),
                            end_date.strftime("%Y-%m-%d"),
                            managers_list,
                            exclusion_list
                        )
                        
                        progress_bar.progress(100)
                        status_text.text("✅ 분석 완료!")
                        
                        st.session_state.analysis_complete = True
                        st.session_state.result_file = result_file
                        st.session_state.processing = False
                        
                        st.success("🎉 분석이 성공적으로 완료되었습니다!")
                        st.balloons()
                
                # 로그 표시
                with log_container:
                    if st.session_state.log_messages:
                        st.markdown('<div class="log-box">', unsafe_allow_html=True)
                        for log in st.session_state.log_messages:
                            st.text(log)
                        st.markdown('</div>', unsafe_allow_html=True)
        
        elif not uploaded_file:
            st.info("📌 분석을 시작하려면 먼저 파일을 업로드해주세요.")
        
        st.markdown('</div>', unsafe_allow_html=True)
    
    with tab3:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.markdown('<div class="card-header">📥 분석 결과 다운로드</div>', unsafe_allow_html=True)
        
        if st.session_state.analysis_complete and st.session_state.result_file:
            # 결과 요약
            col1, col2 = st.columns(2)
            
            with col1:
                st.markdown("""
                    <div class="info-box">
                        <h4>📌 생성된 시트 목록</h4>
                        <ul>
                            <li>📊 스코어보드</li>
                            <li>📋 채팅분석_요약</li>
                            <li>👔 관리자_분석</li>
                            <li>📈 채팅분석_지표</li>
                            <li>💬 상세 메시지 내역 (도움/담당)</li>
                            <li>❓ 도움말</li>
                        </ul>
                    </div>
                """, unsafe_allow_html=True)
            
            with col2:
                # 다운로드 버튼
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                filename = f"SNS센터_채팅분석_결과_{timestamp}.xlsx"
                
                st.download_button(
                    label="💾 결과 파일 다운로드",
                    data=st.session_state.result_file,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    type="primary"
                )
                
                st.success("✅ 다운로드 준비 완료!")
                
                # 추가 액션
                st.divider()
                col1, col2 = st.columns(2)
                with col1:
                    if st.button("🔄 새로운 분석 시작", use_container_width=True):
                        st.session_state.clear()
                        st.rerun()
                with col2:
                    if st.button("📊 분석 설정 유지", use_container_width=True):
                        st.session_state.analysis_complete = False
                        st.session_state.result_file = None
                        st.rerun()
        else:
            st.info("📌 분석이 완료되면 여기서 결과를 다운로드할 수 있습니다.")
            st.markdown("""
                <div style="text-align: center; padding: 3rem;">
                    <h3 style="color: #4da6ff;">분석 프로세스</h3>
                    <p style="color: #b3d9ff; font-size: 1.1rem;">
                        1️⃣ 파일 업로드 → 2️⃣ 기간 설정 → 3️⃣ 인원 설정 → 4️⃣ 분석 실행 → 5️⃣ 결과 다운로드
                    </p>
                </div>
            """, unsafe_allow_html=True)
        
        st.markdown('</div>', unsafe_allow_html=True)

    # 푸터
    st.divider()
    st.markdown("""
        <div style="text-align: center; color: #4da6ff; padding: 1rem;">
            <p>© 2025 SNS센터 채팅분석 프로그램 v1.9 | Powered by Streamlit</p>
            <p style="font-size: 0.9rem; color: #666;">아정당 커뮤니케이션즈</p>
        </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
