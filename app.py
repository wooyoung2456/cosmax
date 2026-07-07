import base64
from pathlib import Path

import streamlit as st
import streamlit.components.v1 as components

st.set_page_config(page_title="교대 근무 인수인계 보고서", page_icon="📋", layout="wide")

# Streamlit 기본 여백/헤더를 없애서 원본 HTML 디자인이 그대로 보이게 함
st.markdown(
    """
    <style>
      .block-container { padding: 0 !important; max-width: 100% !important; }
      header[data-testid="stHeader"] { display: none; }
      iframe { border: none; }
    </style>
    """,
    unsafe_allow_html=True,
)

BASE_DIR = Path(__file__).parent


def to_data_uri(filename: str, mime: str) -> str:
    data = (BASE_DIR / filename).read_bytes()
    b64 = base64.b64encode(data).decode("ascii")
    return f"data:{mime};base64,{b64}"


html = (BASE_DIR / "인수인계보고서.html").read_text(encoding="utf-8")

# 상대경로 이미지(cosmax_logo.jpg, handover_illustration.jpg)를
# base64로 인라인 삽입해서, 배포 환경(작업 디렉터리)이 달라져도 이미지가 깨지지 않게 함
html = html.replace('src="cosmax_logo.jpg"', f'src="{to_data_uri("cosmax_logo.jpg", "image/jpeg")}"')
html = html.replace(
    'src="handover_illustration.jpg"',
    f'src="{to_data_uri("handover_illustration.jpg", "image/jpeg")}"',
)

# 이 HTML은 서버에서 AI를 직접 호출하지 않고,
# 설문 내용을 클립보드에 복사한 뒤 claude.ai를 새 창으로 열어 사용자가 직접 붙여넣는 방식이라
# ANTHROPIC_API_KEY 등 별도 API 키가 전혀 필요하지 않음
components.html(html, height=1400, scrolling=True)
