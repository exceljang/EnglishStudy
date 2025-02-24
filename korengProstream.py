import streamlit as st
import openpyxl
import edge_tts
import asyncio
import os
from pathlib import Path
import base64
import pygame.mixer
import time

class KorEngPlayer:
    def __init__(self):
        # 다크 테마 설정
        st.set_page_config(
            page_title="English Learning",
            layout="wide",
            initial_sidebar_state="collapsed"
        )
        
        # CSS 스타일 적용
        st.markdown("""
            <style>
                .stApp {
                    background-color: #0E1117;
                    color: #FFFFFF;
                }
                .main .block-container {
                    max-width: 720px;
                    padding-top: 2rem;
                    padding-right: 1rem;
                    padding-left: 1rem;
                    padding-bottom: 2rem;
                }
                .st-emotion-cache-q8sbsg p {
                    color: #FFFFFF !important;
                }
                .st-emotion-cache-183lzff {
                    color: #FFFFFF;
                }
                .stMarkdown {
                    color: #FFFFFF;
                }
                .stSelectbox [data-baseweb="select"] {
                    max-width: 720px;
                }
                .stSlider {
                    max-width: 720px;
                }
                .stButton > button {
                    width: 150px !important;
                    display: inline-flex;
                    justify-content: center;
                    align-items: center;
                    color: #FFFFFF;
                    background-color: #1E1E1E;
                    border: 1px solid #4A4A4A;
                }
                .stButton > button:hover {
                    background-color: #2E2E2E;
                    border-color: #6A6A6A;
                }
                div[data-testid="column"] {
                    display: flex;
                    justify-content: center;
                    padding: 0 5px !important;
                }
                .title {
                    font-size: 24px !important;
                }
                .metric-label {
                    font-size: 14px !important;
                }
                .metric-value {
                    font-size: 16px !important;
                }
                /* 셀렉트박스와 체크박스 레이블 색상 */
                .stSelectbox label, .stCheckbox label {
                    color: #FFFFFF !important;
                }
                /* 체크박스 레이블 추가 스타일 */
                .st-emotion-cache-1gulkj5, .st-emotion-cache-1t9giqp {
                    color: #FFFFFF !important;
                }
                label[data-testid="stCheckbox"] {
                    color: #FFFFFF !important;
                }
            </style>
        """, unsafe_allow_html=True)
        
        # 중앙 정렬을 위한 컨테이너
        container = st.container()
        with container:
            st.markdown('<h1 class="title">English Learning</h1>', unsafe_allow_html=True)
        
        # pygame 초기화 수정 - 오디오 없이도 실행되도록
        try:
            os.environ['SDL_AUDIODRIVER'] = 'dummy'
            pygame.mixer.init()
        except Exception as e:
            st.warning("오디오 장치를 찾을 수 없습니다. 음성이 재생되지 않을 수 있습니다.")
        
        # 세션 상태 초기화
        if 'playing' not in st.session_state:
            st.session_state.playing = False
        if 'current_row' not in st.session_state:
            st.session_state.current_row = 2
        if 'repeat' not in st.session_state:
            st.session_state.repeat = False
            
        # 임시 디렉토리 설정
        self.temp_dir = Path("temp_audio")
        self.temp_dir.mkdir(exist_ok=True)
        
        # 엑셀 파일 로드 - 상대 경로 사용
        try:
            excel_path = 'korengpro.xlsx'  # 현재 작업 디렉토리 기준
            if not os.path.exists(excel_path):
                st.error(f"엑셀 파일을 찾을 수 없습니다. 현재 디렉토리: {os.getcwd()}")
                st.error("'korengpro.xlsx' 파일을 업로드해주세요.")
                uploaded_file = st.file_uploader("Excel 파일 선택", type=['xlsx'])
                if uploaded_file:
                    with open('korengpro.xlsx', 'wb') as f:
                        f.write(uploaded_file.getvalue())
                    excel_path = 'korengpro.xlsx'
                else:
                    st.stop()
            
            self.workbook = openpyxl.load_workbook(excel_path)
            self.sheet_names = self.workbook.sheetnames[:10]
        except Exception as e:
            st.error(f"엑셀 파일 로드 실패: {e}")
            st.stop()
        
        self.setup_interface()

    def cleanup_file(self, filename):
        try:
            if os.path.exists(filename):
                for _ in range(3):
                    try:
                        os.remove(filename)
                        break
                    except:
                        time.sleep(0.1)
        except Exception as e:
            st.error(f"파일 삭제 실패: {filename}")

    def play_audio(self, filename):
        try:
            if os.path.exists(filename):
                with open(filename, 'rb') as f:
                    audio_bytes = f.read()
                st.audio(audio_bytes, format='audio/mp3')
                time.sleep(0.1)  # 짧은 대기 시간 추가
        except Exception as e:
            st.error(f"음성 재생 오류: {e}")

    async def generate_speech(self, text, language):
        try:
            voice = "ko-KR-SunHiNeural" if language == "ko" else "en-US-JennyNeural"
            filename = self.temp_dir / f"temp_{language}_{hash(text)}.mp3"
            
            communicate = edge_tts.Communicate(text, voice)
            await communicate.save(str(filename))
            await asyncio.sleep(0.1)
            
            return str(filename)
        except Exception as e:
            st.error(f"음성 생성 오류: {e}")
            return None

    def setup_interface(self):
        # 이전 선택된 시트 추적
        if 'previous_sheet' not in st.session_state:
            st.session_state.previous_sheet = None
        
        # 주제 선택
        selected_sheet = st.selectbox(
            "주제를 선택하세요",
            ["선택하세요"] + self.sheet_names,
            key='sheet_selector'
        )
        
        # 시트가 변경되었을 때 현재 행 초기화
        if selected_sheet != st.session_state.previous_sheet:
            st.session_state.current_row = 2
            st.session_state.previous_sheet = selected_sheet
        
        if selected_sheet and selected_sheet != "선택하세요":
            sheet = self.workbook[selected_sheet]
            total_rows = sheet.max_row - 1
            
            # 진행률 슬라이더 (1부터 시작)
            progress = st.slider(
                "진행률",
                1, total_rows,
                max(1, st.session_state.current_row - 1),
                key='progress_slider',
                label_visibility="collapsed"
            )
            st.session_state.current_row = progress + 1
            
            # 진행률 숫자와 반복재생 체크박스를 나란히 배치
            status_cols = st.columns([1, 6])  # 비율 조정
            with status_cols[0]:
                current_number = max(1, st.session_state.current_row - 1)
                st.markdown(f'<p class="metric-value">{current_number}/{total_rows}</p>', unsafe_allow_html=True)
            with status_cols[1]:
                st.checkbox("Rept", key='repeat', help=None)
            
            # 버튼 컨트롤 - Start와 Stop 버튼만 표시
            button_cols = st.columns([1, 1, 5])  # 비율 조정
            with button_cols[0]:
                if st.button("Start", use_container_width=False):
                    st.session_state.playing = True
            with button_cols[1]:
                if st.button("Stop", use_container_width=False):
                    st.session_state.playing = False
                    pygame.mixer.music.stop()
                    pygame.mixer.quit()
                    self.cleanup()
            
            # 텍스트 표시를 위한 컨테이너 생성 (버튼 아래에 위치)
            st.write("")  # 간격 추가
            self.korean_container = st.empty()
            self.english_container = st.empty()
            
            # 재생 상태 확인 및 실행
            if st.session_state.playing:
                self.play_current_item(sheet)

    def play_current_item(self, sheet):
        if st.session_state.current_row <= sheet.max_row:
            # 현재 행의 텍스트 가져오기
            kor_text = sheet.cell(row=st.session_state.current_row, column=2).value
            eng_text = sheet.cell(row=st.session_state.current_row, column=3).value
            
            # 한글 텍스트 표시
            self.korean_container.markdown(f"<p style='color: white; font-size: 20px;'>{kor_text if kor_text else ''}</p>", unsafe_allow_html=True)
            self.english_container.empty()
            
            # 한글 음성 생성 및 재생
            if kor_text:
                asyncio.run(self.play_text(kor_text, "ko"))
            
            if not st.session_state.playing:
                return
            
            # 영어 텍스트 표시
            self.english_container.markdown(f"<p style='color: white; font-size: 20px;'>{eng_text if eng_text else ''}</p>", unsafe_allow_html=True)
            
            # 영어 음성 생성 및 재생
            if eng_text:
                asyncio.run(self.play_text(eng_text, "en"))
            
            # 진행률 업데이트를 위해 rerun 전에 현재 행 증가
            next_row = st.session_state.current_row + 1
            if next_row > sheet.max_row:
                if st.session_state.repeat:
                    next_row = 2
                else:
                    st.session_state.playing = False
                    return
            
            st.session_state.current_row = next_row
            time.sleep(0.5)
            st.rerun()

    async def play_text(self, text, language):
        try:
            filename = await self.generate_speech(text, language)
            if filename:
                self.play_audio(filename)
        except Exception as e:
            st.error(f"음성 재생 오류: {e}")

    def cleanup(self):
        try:
            for file in os.listdir(self.temp_dir):
                file_path = os.path.join(self.temp_dir, file)
                try:
                    os.remove(file_path)
                except:
                    pass
            os.rmdir(self.temp_dir)
        except:
            pass

def main():
    try:
        app = KorEngPlayer()
    except Exception as e:
        st.error(f"실행 중 오류 발생: {e}")

if __name__ == "__main__":
    main()
