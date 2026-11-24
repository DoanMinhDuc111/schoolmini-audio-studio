import asyncio
import base64
import html
import io
import json
from datetime import datetime
from pathlib import Path

import docx
import edge_tts
from gtts import gTTS
from pypdf import PdfReader
import streamlit as st

st.set_page_config(
    page_title="schoolmini - Thư viện sách của Trường con",
    page_icon="🎧",
    layout="wide",
)

STORIES_FILE = Path("stories.json")
STATE_FILE = Path("library_state.json")
AMBIENT_DIR = Path("ambient")


def make_cover_svg_data_uri(title: str, tag: str, c1: str, c2: str) -> str:
    safe_title = title.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    safe_tag = tag.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    svg = f"""<svg xmlns="http://www.w3.org/2000/svg" width="240" height="150" viewBox="0 0 240 150">
      <defs>
        <linearGradient id="g" x1="0" y1="0" x2="1" y2="1">
          <stop offset="0%" stop-color="{c1}"/>
          <stop offset="100%" stop-color="{c2}"/>
        </linearGradient>
      </defs>
      <rect width="240" height="150" rx="10" fill="url(#g)"/>
      <circle cx="208" cy="28" r="36" fill="rgba(255,255,255,0.14)"/>
      <rect x="14" y="14" width="84" height="20" rx="10" fill="rgba(0,0,0,0.32)"/>
      <text x="56" y="28" fill="#ffffff" font-family="Arial, sans-serif" font-size="10" font-weight="bold" text-anchor="middle">{safe_tag}</text>
      <text x="14" y="76" fill="#ffffff" font-family="Arial, sans-serif" font-size="15" font-weight="bold">{safe_title[:18]}</text>
      <text x="14" y="96" fill="rgba(255,255,255,0.92)" font-family="Arial, sans-serif" font-size="13.5" font-weight="bold">{safe_title[18:36]}</text>
      <text x="14" y="132" fill="rgba(255,255,255,0.85)" font-family="Arial, sans-serif" font-size="11">schoolmini • Sách Nói</text>
    </svg>"""
    b64 = base64.b64encode(svg.encode("utf-8")).decode("ascii")
    return f"data:image/svg+xml;base64,{b64}"


DEFAULT_STORIES = {
    "Ngọn đèn dầu cũ trên gác xép": {
        "genre": "Tản văn",
        "c1": "#da251d",
        "c2": "#7f1d1d",
        "content": """Mưa đầu mùa rơi lách tách trên mái tôn cũ kỹ của ngôi nhà nhỏ ven sông. Minh ngồi bên khung cửa sổ, tay xoay nhẹ chiếc chìa khóa đồng rỉ sét mà ông nội để lại trước khi đi xa. Trong căn phòng chật hẹp chỉ có tiếng mưa và tiếng tích tắc của chiếc đồng hồ treo tường, cậu bỗng thấy lòng mình lắng lại.

Suốt ba năm lên thành phố học tập và làm việc, Minh luôn vội vã. Ngày nào cậu cũng chạy đua với những bản kế hoạch, những cuộc họp kéo dài và những tin nhắn không hồi âm. Có những lúc mệt mỏi quá, Minh quên mất lần cuối mình hít thở trọn vẹn một buổi sáng trong lành là khi nào.

Cầm chiếc chìa khóa bước lên gác xép, Minh đẩy cánh cửa gỗ kêu kẽo kẹt. Lớp bụi mỏng phủ trên những chồng sách cũ và chiếc rương sơn đỏ đã bạc màu. Khi mở chiếc rương ra, thứ đầu tiên hiện lên không phải vàng bạc hay giấy tờ quan trọng, mà là một cuốn sổ tay bìa da sờn mép cùng một ngọn đèn dầu nhỏ.

Trang đầu tiên của cuốn sổ viết nét chữ ngay ngắn: "Con trai, cuộc đời giống như đi đường dài. Đi nhanh chưa chắc đã tới sớm, quan trọng là biết lúc nào nên dừng lại để nhìn xem mình đang đi đâu."

Minh mỉm cười, bật quẹt thắp lên ngọn đèn dầu. Ánh lửa vàng ấm áp bập bùng soi sáng cả góc phòng tối. Ngoài kia, cơn mưa vẫn rơi, nhưng trong lòng Minh, những bộn bề suốt bao lâu nay dường như đã nhẹ đi rất nhiều.""",
    },
    "Chuyến tàu đêm cuối tuần": {
        "genre": "Đời thường",
        "c1": "#b91c1c",
        "c2": "#450a0a",
        "content": """Năm giờ rưỡi chiều thứ Sáu, Khoa vơ vội chiếc balo rồi lao ra khỏi văn phòng trước khi đồng hồ kịp điểm phút thứ năm. Suốt cả tuần liền, anh dán mắt vào màn hình sửa từng dòng báo cáo, nên giờ chỉ mong kịp lên chuyến tàu lúc sáu giờ tối để về quê thăm mẹ.

Vừa tới ga, dòng người xếp hàng mua vé đã dài ngoằng. Khoa thở phào khi nhớ ra mình đã đặt vé điện tử từ trưa. Anh bước nhanh qua cửa soát vé, tìm đúng toa số bốn và ngồi xuống cạnh ô cửa kính. Ngoài sân ga, những ánh đèn vàng bắt đầu bật sáng, lướt qua vội vã khi con tàu từ từ chuyển bánh.

Đối diện Khoa là một cô bé ngồi cùng mẹ, tay ôm chặt chú gấu bông cũ. Thấy tàu chạy nhanh, cô bé tò mò hỏi:
"Mẹ ơi, sao cột điện ngoài kia cứ chạy ngược lại thế ạ?"
Người mẹ mỉm cười xoa đầu con:
"Vì tàu của mình đang tiến về phía trước đấy con. Cũng giống như khi mình cố gắng đi tới, mọi thứ xung quanh sẽ đổi khác."

Khoa nghe câu trả lời ấy mà bất giác mỉm cười. Đôi khi, chỉ cần bước ra khỏi nhịp sống vội vã thường ngày và ngồi trên một chuyến tàu chậm rãi, người ta mới nhận ra mình đã đi được bao xa. Gió chiều mát lạnh lùa qua khe cửa, mang theo mùi đồng lúa quen thuộc báo hiệu quê nhà đã ở rất gần.""",
    },
    "Quán cà phê góc phố mưa": {
        "genre": "Chữa lành",
        "c1": "#dc2626",
        "c2": "#7c2d12",
        "content": """Góc phố nhỏ tên An Lạc có một quán cà phê nép dưới tán bằng lăng già. Không biển hiệu lớn, không nhạc ồn ào, chỉ có mùi cà phê rang xay thoang thoảng quyện cùng tiếng mưa rơi đều đặn trên hiên gạch.

Lan bước vào quán khi chiếc áo khoác đã thấm lạnh. Cô chọn chỗ ngồi cạnh ô cửa kính, gọi một ly cà phê sữa nóng và mở cuốn sổ ghi chép dang dở. Suốt nửa tháng qua, dự án khởi nghiệp gặp trục trặc khiến cô mất ngủ triền miên và luôn tự trách bản thân chưa đủ giỏi.

Chủ quán — một người đàn ông trung niên tóc muối tiêu — đặt xuống bàn một chiếc bánh quy nhỏ kèm lời nhắn: "Đôi khi một bước lùi không phải để thua, mà để lấy đà nhảy qua vũng nước lớn hơn."

Lan nhấp một ngụm cà phê ấm, nhìn những giọt mưa chảy dài trên kính. Cơn giông ngoài kia rồi cũng sẽ tạnh, và cô biết mình cần bắt đầu lại từng việc nhỏ nhất thay vì đứng yên lo lắng.""",
    },
}


def load_stories() -> dict:
    if STORIES_FILE.exists():
        try:
            data = json.loads(STORIES_FILE.read_text(encoding="utf-8"))
            if isinstance(data, dict) and data:
                return data
        except Exception:
            pass
    STORIES_FILE.write_text(
        json.dumps(DEFAULT_STORIES, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    return dict(DEFAULT_STORIES)


def save_stories(stories: dict) -> None:
    STORIES_FILE.write_text(
        json.dumps(stories, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def load_state() -> dict:
    if STATE_FILE.exists():
        try:
            data = json.loads(STATE_FILE.read_text(encoding="utf-8"))
            return {
                "favorites": data.get("favorites", []),
                "history": data.get("history", []),
                "bookmarks": data.get("bookmarks", {}),
                "play_counts": data.get("play_counts", {}),
                "notes": data.get("notes", {}),
            }
        except Exception:
            pass
    return {
        "favorites": [],
        "history": [],
        "bookmarks": {},
        "play_counts": {},
        "notes": {},
    }


def save_state(state: dict) -> None:
    STATE_FILE.write_text(
        json.dumps(state, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def toggle_favorite(title: str) -> None:
    state = load_state()
    if title in state["favorites"]:
        state["favorites"].remove(title)
    else:
        state["favorites"].append(title)
    save_state(state)


def record_play(title: str, mode_label: str, voice: str, speed: str) -> None:
    state = load_state()
    state["play_counts"][title] = state["play_counts"].get(title, 0) + 1
    state["history"].insert(
        0,
        {
            "Thời gian": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "Tên truyện": title,
            "Chế độ": mode_label,
            "Giọng đọc": voice,
            "Tốc độ": speed,
        },
    )
    state["history"] = state["history"][:60]
    save_state(state)


def save_bookmark(title: str, para_index: int) -> None:
    state = load_state()
    state["bookmarks"][title] = para_index
    save_state(state)


def save_note(title: str, note_text: str) -> None:
    state = load_state()
    state["notes"][title] = note_text
    save_state(state)


def extract_uploaded_text(uploaded_file) -> str:
    name = uploaded_file.name.lower()
    raw_bytes = uploaded_file.getvalue()
    if name.endswith(".txt"):
        return raw_bytes.decode("utf-8", errors="ignore")
    if name.endswith(".docx"):
        doc = docx.Document(io.BytesIO(raw_bytes))
        return "\n\n".join([p.text.strip() for p in doc.paragraphs if p.text.strip()])
    if name.endswith(".pdf"):
        reader = PdfReader(io.BytesIO(raw_bytes))
        pages = []
        for page in reader.pages:
            txt = page.extract_text() or ""
            if txt.strip():
                pages.append(txt.strip())
        return "\n\n".join(pages)
    return ""


def split_paragraphs(text: str) -> list[str]:
    return [p.strip() for p in text.split("\n\n") if p.strip()]


async def synth_edge_bytes(txt: str, v: str, r: str, p: str, vol: str) -> bytes:
    comm = edge_tts.Communicate(txt, v, rate=r, pitch=p, volume=vol)
    buf = bytearray()
    async for chunk in comm.stream():
        if chunk["type"] == "audio":
            buf.extend(chunk["data"])
    return bytes(buf)


def synth_gtts_bytes(txt: str) -> bytes:
    mp3_fp = io.BytesIO()
    tts = gTTS(text=txt, lang="vi")
    tts.write_to_fp(mp3_fp)
    return mp3_fp.getvalue()


def generate_audio_for_text(txt: str, v: str, r: str, p: str, vol: str) -> tuple[bytes, str]:
    try:
        audio_bytes = asyncio.run(synth_edge_bytes(txt, v, r, p, vol))
        if audio_bytes:
            return audio_bytes, "Edge TTS"
    except Exception:
        pass
    return synth_gtts_bytes(txt), "Google TTS (dự phòng)"


def get_ambient_tracks() -> list[Path]:
    if not AMBIENT_DIR.exists():
        return []
    tracks = []
    for ext in ("*.mp3", "*.wav", "*.ogg"):
        tracks.extend(sorted(AMBIENT_DIR.glob(ext)))
    return tracks


# --- CALLBACKS AN TOÀN CHO SESSION STATE WIDGET ---
def go_home_callback() -> None:
    st.session_state.active_tab = "📚 Tủ truyện"
    st.session_state["nav_radio"] = "📚 Tủ truyện"


def open_story_callback(title: str, content: str, para_idx: int) -> None:
    st.session_state.selected_title = title
    st.session_state.story_text = content
    st.session_state.current_para_idx = para_idx
    st.session_state.active_tab = "📖 Đọc & Nghe"
    st.session_state["nav_radio"] = "📖 Đọc & Nghe"


stories = load_stories()
state_data = load_state()

if "active_tab" not in st.session_state:
    st.session_state.active_tab = "📚 Tủ truyện"
if "nav_radio" not in st.session_state:
    st.session_state["nav_radio"] = st.session_state.active_tab
if "selected_title" not in st.session_state:
    first_title = next(iter(stories.keys()))
    st.session_state.selected_title = first_title
    st.session_state.story_text = stories[first_title]["content"]
if "reader_theme" not in st.session_state:
    st.session_state.reader_theme = "Tối"
if "current_para_idx" not in st.session_state:
    st.session_state.current_para_idx = 0

st.markdown(
    """
<style>
    @import url('https://fonts.googleapis.com/css2?family=Be+Vietnam+Pro:wght@400;500;600;700;800&family=Plus+Jakarta+Sans:wght@500;600;700;800&display=swap');

    html, body, .stApp, .stApp p, .stApp label, .stApp h1, .stApp h2, .stApp h3, .stApp span, .stApp div, .stApp button, .stApp input, .stApp textarea {
        font-family: 'Plus Jakarta Sans', 'Be Vietnam Pro', system-ui, -apple-system, BlinkMacSystemFont, sans-serif !important;
    }

    .stApp {
        background-color: #090d16;
        color: #e2e8f0;
    }
    .stApp p, .stApp label, .stApp h1, .stApp h2, .stApp h3, .stApp span, .stApp div {
        color: #e2e8f0;
    }

    /* Header mở rộng, đỏ cờ Việt Nam rực phía trên & mờ dần xuống nền tối */
    .schoolmini-hero-bg {
        position: relative;
        min-height: 195px;
        width: 100%;
        background: linear-gradient(
            180deg,
            #da251d 0%,
            #b91c1c 44%,
            rgba(127, 29, 29, 0.52) 74%,
            rgba(9, 13, 22, 0.05) 94%,
            #090d16 100%
        );
        border-radius: 24px 24px 0 0;
        padding: 24px 32px 48px 32px;
        margin-bottom: -168px;
        z-index: 1;
        pointer-events: none;
    }
    .schoolmini-center-box {
        position: relative;
        z-index: 2;
        text-align: center;
        padding-top: 30px;
        padding-bottom: 12px;
    }
    .schoolmini-center-title {
        font-size: 2.15rem !important;
        font-weight: 800 !important;
        color: #ffffff !important;
        letter-spacing: -0.025em;
        margin: 0 auto 8px auto !important;
        text-shadow: 0 3px 14px rgba(0, 0, 0, 0.45);
    }
    .schoolmini-center-sub {
        font-size: 1rem !important;
        font-weight: 500 !important;
        color: #fff5f5 !important;
        opacity: 0.96;
        margin: 0 auto !important;
        text-shadow: 0 1px 8px rgba(0, 0, 0, 0.45);
    }

    /* Nút schoolmini nằm bên trong góc trên trái header, rõ nét, không viền, phông chữ to kiểu Waka */
    .st-key-brand_home_btn {
        position: relative;
        z-index: 25 !important;
        display: inline-block !important;
        width: auto !important;
        margin-left: 24px !important;
        margin-top: 8px !important;
        margin-bottom: 105px !important;
    }
    .st-key-brand_home_btn > button,
    .st-key-brand_home_btn button[kind="secondary"] {
        background: transparent !important;
        background-color: transparent !important;
        border: none !important;
        outline: none !important;
        box-shadow: none !important;
        padding: 6px 10px !important;
        width: auto !important;
        min-height: 0 !important;
        opacity: 1 !important;
        color: #fde047 !important;
        font-size: 2.05rem !important;
        font-weight: 800 !important;
        line-height: 1.15 !important;
        letter-spacing: -0.03em !important;
        text-shadow: 0 2px 12px rgba(0, 0, 0, 0.55) !important;
        cursor: pointer !important;
        transition: color 0.2s ease, transform 0.2s ease !important;
    }
    .st-key-brand_home_btn > button:hover,
    .st-key-brand_home_btn button[kind="secondary"]:hover {
        background: transparent !important;
        background-color: transparent !important;
        border: none !important;
        box-shadow: none !important;
        color: #ffffff !important;
        transform: translateX(3px) !important;
    }
    .st-key-brand_home_btn > button p,
    .st-key-brand_home_btn > button span {
        color: inherit !important;
        font-size: inherit !important;
        font-weight: inherit !important;
    }

    .waka-badge {
        display: inline-block;
        background-color: rgba(218, 37, 29, 0.16);
        color: #fca5a5 !important;
        border: 1px solid rgba(248, 113, 113, 0.3);
        font-weight: 700;
        font-size: 0.78rem;
        padding: 4px 12px;
        border-radius: 999px;
        margin-bottom: 8px;
    }
    div[data-testid="stVerticalBlockBorderWrapper"] {
        background-color: #111827;
        border-radius: 16px;
        border: 1px solid #1e293b !important;
        padding: 12px 14px;
        box-shadow: 0 6px 18px rgba(0, 0, 0, 0.32);
    }
    textarea, input {
        background-color: #0b101b !important;
        color: #f1f5f9 !important;
        border-color: #263244 !important;
        border-radius: 10px !important;
    }
    .stButton:not(.st-key-brand_home_btn) > button:first-child {
        background-color: #c81e1e;
        color: #ffffff !important;
        font-weight: 600;
        font-size: 0.96rem;
        border: none;
        border-radius: 10px;
        padding: 0.52rem 1rem;
        width: 100%;
        transition: background-color 0.2s ease, transform 0.15s ease;
    }
    .stButton:not(.st-key-brand_home_btn) > button:first-child:hover {
        background-color: #da251d;
        color: #ffffff !important;
    }

    /* Thẻ hình chữ nhật thu gọn chỉ hiện tên -> rê chuột vào kéo ngang mở rộng */
    .waka-rect-track {
        display: flex;
        align-items: stretch;
        width: 100%;
        min-height: 54px;
        background: #141d2e;
        border: 1px solid #26354d;
        border-radius: 14px;
        overflow: hidden;
        margin-bottom: 8px;
        transition: border-color 0.25s ease, box-shadow 0.25s ease, background 0.25s ease;
    }
    .waka-rect-track:hover {
        background: #18243a;
        border-color: #da251d;
        box-shadow: 0 8px 22px rgba(218, 37, 29, 0.22);
    }
    .waka-rect-spine {
        min-width: 240px;
        max-width: 280px;
        padding: 14px 16px;
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 10px;
        background: linear-gradient(90deg, #131c2e 0%, #172238 100%);
        border-right: 1px solid #26354d;
        flex-shrink: 0;
        transition: min-width 0.35s cubic-bezier(0.22, 1, 0.36, 1), background 0.25s ease;
    }
    .waka-rect-track:hover .waka-rect-spine {
        min-width: 210px;
        background: linear-gradient(90deg, #7f1d1d 0%, #450a0a 100%);
    }
    .waka-rect-title {
        font-size: 1.04rem;
        font-weight: 700;
        color: #f8fafc !important;
        line-height: 1.35;
    }
    .waka-rect-arrow {
        font-size: 0.8rem;
        color: #94a3b8 !important;
        white-space: nowrap;
        transition: transform 0.3s ease, color 0.25s ease;
    }
    .waka-rect-track:hover .waka-rect-arrow {
        transform: translateX(4px);
        color: #fde047 !important;
    }
    .waka-rect-drawer {
        display: flex;
        align-items: center;
        gap: 14px;
        max-width: 0;
        opacity: 0;
        padding: 0;
        transform: translateX(-24px);
        overflow: hidden;
        transition:
            max-width 0.45s cubic-bezier(0.22, 1, 0.36, 1),
            opacity 0.3s ease,
            padding 0.35s ease,
            transform 0.42s cubic-bezier(0.22, 1, 0.36, 1);
    }
    .waka-rect-track:hover .waka-rect-drawer {
        max-width: 760px;
        opacity: 1;
        padding: 10px 14px;
        transform: translateX(0);
    }
    .waka-rect-thumb {
        width: 118px;
        height: 74px;
        object-fit: cover;
        border-radius: 8px;
        flex-shrink: 0;
        border: 1px solid rgba(255, 255, 255, 0.14);
    }
    .waka-rect-info {
        min-width: 240px;
    }
    .waka-rect-meta {
        font-size: 0.84rem;
        color: #fca5a5 !important;
        margin-bottom: 4px;
        font-weight: 600;
    }
    .waka-rect-desc {
        font-size: 0.88rem;
        color: #cbd5e1 !important;
        line-height: 1.45;
    }
</style>
""",
    unsafe_allow_html=True,
)

# --- HEADER MỞ RỘNG VỚI DÒNG CHỮ Ở GIỮA & NÚT SCHOOLMINI RÕ NÉT BÊN TRÁI ---
st.markdown(
    """
    <div class="schoolmini-hero-bg">
        <div class="schoolmini-center-box">
            <div class="schoolmini-center-title">Thư viện sách của Trường con</div>
            <p class="schoolmini-center-sub">Khám phá truyện đọc &amp; sách nói AI — rê chuột vào thẻ để kéo ngang mở rộng thông tin.</p>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

st.button(
    "schoolmini",
    key="brand_home_btn",
    help="Về trang chính (Tủ truyện)",
    on_click=go_home_callback,
)

tab_options = [
    "📚 Tủ truyện",
    "📖 Đọc & Nghe",
    "📁 Tải tài liệu",
    "🕒 Lịch sử & Thống kê",
]

selected_tab = st.radio(
    "Điều hướng:",
    tab_options,
    key="nav_radio",
    horizontal=True,
    label_visibility="collapsed",
)
st.session_state.active_tab = selected_tab

current_tab = st.session_state.active_tab
favorites_set = set(state_data["favorites"])
bookmarks_map = state_data["bookmarks"]
play_counts_map = state_data["play_counts"]
notes_map = state_data["notes"]

if current_tab == "📚 Tủ truyện":
    with st.container(border=True):
        st.markdown(
            '<span class="waka-badge">KHÁM PHÁ TỦ TRUYỆN (RÊ CHUỘT ĐỂ KÉO NGANG MỞ RỘNG)</span>',
            unsafe_allow_html=True,
        )

        f_col1, f_col2, f_col3 = st.columns([2, 1, 1])
        with f_col1:
            search_q = st.text_input("Tìm theo tên truyện hoặc nội dung:", placeholder="Nhập từ khóa...")
        with f_col2:
            genres = ["Tất cả"] + sorted({info.get("genre", "Khác") for info in stories.values()})
            selected_genre = st.selectbox("Thể loại:", genres)
        with f_col3:
            only_fav = st.checkbox("Chỉ hiện Yêu thích ❤️", value=False)

        filtered_items = []
        for title, info in stories.items():
            match_q = (
                not search_q.strip()
                or search_q.lower() in title.lower()
                or search_q.lower() in info.get("content", "").lower()
            )
            match_genre = selected_genre == "Tất cả" or info.get("genre") == selected_genre
            match_fav = (not only_fav) or (title in favorites_set)
            if match_q and match_genre and match_fav:
                filtered_items.append((title, info))

        if not filtered_items:
            st.info("Không tìm thấy truyện phù hợp.")
        else:
            for idx, (title, info) in enumerate(filtered_items):
                content_str = info.get("content", "")
                words = len(content_str.split())
                mins = max(1, round(words / 150))
                plays = play_counts_map.get(title, 0)
                bm_idx = bookmarks_map.get(title)
                bm_str = f" • Đọc dở đoạn {bm_idx + 1}" if bm_idx is not None else ""
                fav_icon = "❤️ " if title in favorites_set else ""

                excerpt = content_str.strip().splitlines()[0] if content_str.strip() else ""
                if len(excerpt) > 130:
                    excerpt = excerpt[:130].rstrip() + "..."

                cover_uri = make_cover_svg_data_uri(
                    title,
                    info.get("genre", "TRUYỆN").upper(),
                    info.get("c1", "#da251d"),
                    info.get("c2", "#7f1d1d"),
                )

                esc_title = html.escape(title)
                esc_genre = html.escape(info.get("genre", "Khác"))
                esc_excerpt = html.escape(excerpt)

                row_c1, row_c2, row_c3 = st.columns([4.4, 1.3, 0.9], gap="small")
                with row_c1:
                    st.markdown(
                        f"""
                        <div class="waka-rect-track">
                            <div class="waka-rect-spine">
                                <span class="waka-rect-title">{fav_icon}{esc_title}</span>
                                <span class="waka-rect-arrow">Kéo ngang ▸</span>
                            </div>
                            <div class="waka-rect-drawer">
                                <img class="waka-rect-thumb" src="{cover_uri}" alt="{esc_title}" />
                                <div class="waka-rect-info">
                                    <div class="waka-rect-meta">{esc_genre} • {words} từ • ~{mins} phút • {plays} lượt nghe{html.escape(bm_str)}</div>
                                    <div class="waka-rect-desc">{esc_excerpt}</div>
                                </div>
                            </div>
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )
                with row_c2:
                    st.button(
                        "▶️ Đọc truyện",
                        key=f"open_{idx}_{title}",
                        on_click=open_story_callback,
                        args=(title, content_str, bookmarks_map.get(title, 0)),
                    )
                with row_c3:
                    fav_label = "❤️ Thích" if title in favorites_set else "🤍 Thích"
                    if st.button(fav_label, key=f"fav_{idx}_{title}"):
                        toggle_favorite(title)
                        st.rerun()

    st.markdown("---")
    with st.container(border=True):
        st.markdown('<span class="waka-badge">➕ THÊM TRUYỆN MỚI VÀO TỦ (STORIES.JSON)</span>', unsafe_allow_html=True)
        with st.form("add_story_form", clear_on_submit=True):
            n_col1, n_col2, n_col3 = st.columns([2, 1, 1])
            with n_col1:
                new_title = st.text_input("Tên truyện mới:")
            with n_col2:
                new_genre = st.text_input("Thể loại:", value="Truyện ngắn")
            with n_col3:
                palette = st.selectbox("Tông màu bìa:", ["Đỏ cờ - Đỏ sẫm", "Đỏ tươi - Nâu trầm", "Đỏ - Xanh lam"])

            new_content = st.text_area(
                "Nội dung truyện (tách các đoạn bằng 1 dòng trống để tiện nghe từng đoạn):",
                height=150,
            )
            if st.form_submit_button("💾 Lưu truyện vào tủ"):
                if not new_title.strip() or not new_content.strip():
                    st.warning("Vui lòng nhập đầy đủ Tên truyện và Nội dung truyện.")
                else:
                    palettes = {
                        "Đỏ cờ - Đỏ sẫm": ("#da251d", "#7f1d1d"),
                        "Đỏ tươi - Nâu trầm": ("#dc2626", "#7c2d12"),
                        "Đỏ - Xanh lam": ("#b91c1c", "#0369a1"),
                    }
                    c1, c2 = palettes[palette]
                    stories[new_title.strip()] = {
                        "genre": new_genre.strip() or "Truyện ngắn",
                        "c1": c1,
                        "c2": c2,
                        "content": new_content.strip(),
                    }
                    save_stories(stories)
                    st.success(f"Đã lưu truyện '{new_title.strip()}' vào stories.json!")
                    st.rerun()

elif current_tab == "📖 Đọc & Nghe":
    left_col, right_col = st.columns([1.35, 1], gap="large")

    paragraphs = split_paragraphs(st.session_state.story_text)
    if paragraphs and st.session_state.current_para_idx >= len(paragraphs):
        st.session_state.current_para_idx = 0

    with left_col:
        with st.container(border=True):
            st.markdown('<span class="waka-badge">KHUNG ĐỌC TRUYỆN</span>', unsafe_allow_html=True)

            t_col1, t_col2 = st.columns([2, 1])
            with t_col1:
                st.subheader(st.session_state.selected_title)
            with t_col2:
                theme_choice = st.selectbox(
                    "Nền trang đọc:",
                    ["Tối", "Giấy vàng", "Trắng sáng"],
                    index=["Tối", "Giấy vàng", "Trắng sáng"].index(st.session_state.reader_theme),
                )
                st.session_state.reader_theme = theme_choice

            st.session_state.story_text = st.text_area(
                "Nội dung truyện (có thể chỉnh sửa trực tiếp):",
                value=st.session_state.story_text,
                height=175,
            )

            theme_styles = {
                "Tối": ("#0b1120", "#f1f5f9", "#26354d"),
                "Giấy vàng": ("#fbf3dc", "#3b2f1e", "#e6d5a8"),
                "Trắng sáng": ("#ffffff", "#1e293b", "#e2e8f0"),
            }
            bg_c, fg_c, bd_c = theme_styles[st.session_state.reader_theme]
            refreshed_paras = split_paragraphs(st.session_state.story_text)

            preview_html = "".join(
                f"<p style='color:{fg_c} !important; line-height:1.7; margin-bottom:10px;'>"
                f"<strong style='opacity:0.7;'>[{i+1}]</strong> {html.escape(p)}</p>"
                for i, p in enumerate(refreshed_paras)
            ) or f"<p style='color:{fg_c} !important;'>Chưa có nội dung.</p>"

            st.markdown(
                f"""
                <div style="background:{bg_c}; border:1px solid {bd_c}; border-radius:10px; padding:14px 16px; max-height:220px; overflow-y:auto;">
                    {preview_html}
                </div>
                """,
                unsafe_allow_html=True,
            )

        with st.container(border=True):
            st.markdown('<span class="waka-badge">📝 GHI CHÚ &amp; XUẤT TỆP (.TXT / .MD)</span>', unsafe_allow_html=True)
            cur_title = st.session_state.selected_title
            existing_note = notes_map.get(cur_title, "")
            note_val = st.text_area(
                "Ghi chú cá nhân cho truyện này:",
                value=existing_note,
                height=100,
                placeholder="Ghi lại ý hay, nhân vật, cảm nhận...",
            )
            if note_val != existing_note:
                save_note(cur_title, note_val)

            export_payload = (
                f"# {cur_title}\n\n"
                f"## Nội dung truyện\n\n{st.session_state.story_text}\n\n"
                f"## Ghi chú cá nhân\n\n{note_val or '(Không có ghi chú)'}\n"
            )
            e1, e2 = st.columns(2)
            with e1:
                st.download_button(
                    label="⬇️ Xuất file .TXT",
                    data=export_payload.encode("utf-8"),
                    file_name=f"{cur_title[:25].strip().replace(' ', '_')}.txt",
                    mime="text/plain",
                )
            with e2:
                st.download_button(
                    label="⬇️ Xuất file .MD",
                    data=export_payload.encode("utf-8"),
                    file_name=f"{cur_title[:25].strip().replace(' ', '_')}.md",
                    mime="text/markdown",
                )

    with right_col:
        with st.container(border=True):
            st.markdown('<span class="waka-badge">TRÌNH ĐIỀU KHIỂN SÁCH NÓI</span>', unsafe_allow_html=True)

            read_mode = st.radio(
                "Chế độ đọc:",
                ["Toàn bộ truyện", "Từng đoạn", "Ghép nhiều đoạn (MP3 liền mạch)"],
                horizontal=False,
            )

            target_text = st.session_state.story_text
            mode_label = "Toàn bộ"
            combined_audio_bytes = b""

            if read_mode == "Từng đoạn":
                if refreshed_paras:
                    idx_default = min(st.session_state.current_para_idx, len(refreshed_paras) - 1)
                    para_labels = [
                        f"Đoạn {i+1}/{len(refreshed_paras)}: {p[:32]}..." if len(p) > 32 else f"Đoạn {i+1}/{len(refreshed_paras)}: {p}"
                        for i, p in enumerate(refreshed_paras)
                    ]
                    p_idx = st.selectbox(
                        "Chọn đoạn:",
                        range(len(refreshed_paras)),
                        index=idx_default,
                        format_func=lambda i: para_labels[i],
                    )
                    st.session_state.current_para_idx = p_idx
                    target_text = refreshed_paras[p_idx]
                    mode_label = f"Đoạn {p_idx+1}/{len(refreshed_paras)}"

                    if st.button("🔖 Lưu vị trí đọc dở tại đoạn này"):
                        save_bookmark(st.session_state.selected_title, p_idx)
                        st.success(f"Đã lưu bookmark tại đoạn {p_idx + 1}!")
                else:
                    target_text = ""

            elif read_mode == "Ghép nhiều đoạn (MP3 liền mạch)":
                if refreshed_paras:
                    total_p = len(refreshed_paras)
                    g1, g2 = st.columns(2)
                    with g1:
                        start_p = st.number_input("Từ đoạn:", min_value=1, max_value=total_p, value=1)
                    with g2:
                        end_p = st.number_input("Đến đoạn:", min_value=start_p, max_value=total_p, value=total_p)
                    selected_slice = refreshed_paras[int(start_p) - 1 : int(end_p)]
                    target_text = "\n\n".join(selected_slice)
                    mode_label = f"Ghép đoạn {int(start_p)}-{int(end_p)}"
                else:
                    selected_slice = []
                    target_text = ""

            words = len(target_text.split()) if target_text.strip() else 0
            est_mins = max(1, round(words / 150)) if words > 0 else 0
            st.caption(f"Đang chọn: {words} từ • Ước tính: ~{est_mins} phút")

            voice_option = st.selectbox(
                "Giọng đọc:",
                ("Nữ miền Nam (Hoài My)", "Nam miền Bắc (Nam Minh)"),
            )
            speed_option = st.selectbox(
                "Tốc độ đọc:",
                ("Chậm hơn (-15%)", "Bình thường (0%)", "Nhanh hơn (+15%)", "Nhanh (+25%)"),
                index=1,
            )
            pitch_option = st.selectbox(
                "Cao độ giọng:",
                ("Trầm hơn (-5Hz)", "Bình thường (+0Hz)", "Cao hơn (+5Hz)"),
                index=1,
            )
            volume_option = st.selectbox(
                "Âm lượng:",
                ("Nhỏ hơn (-15%)", "Bình thường (+0%)", "Lớn hơn (+15%)"),
                index=1,
            )

            voice_code = "vi-VN-HoaiMyNeural" if "Hoài My" in voice_option else "vi-VN-NamMinhNeural"
            rate_code = speed_option.split("(")[-1].replace(")", "")
            pitch_code = pitch_option.split("(")[-1].replace(")", "")
            volume_code = volume_option.split("(")[-1].replace(")", "")

            if st.button("🔊 Tạo & Nghe sách nói"):
                if not target_text.strip():
                    st.warning("Chưa có nội dung để đọc!")
                else:
                    with st.spinner("Đang tổng hợp âm thanh..."):
                        if read_mode == "Ghép nhiều đoạn (MP3 liền mạch)" and refreshed_paras:
                            merged_buf = bytearray()
                            engine_used = "Edge TTS"
                            for seg in selected_slice:
                                b_seg, eng = generate_audio_for_text(
                                    seg, voice_code, rate_code, pitch_code, volume_code
                                )
                                merged_buf.extend(b_seg)
                                if eng != "Edge TTS":
                                    engine_used = eng
                            combined_audio_bytes = bytes(merged_buf)
                            st.success(f"Đã ghép {len(selected_slice)} đoạn thành 1 file MP3 ({engine_used}):")
                        else:
                            combined_audio_bytes, eng = generate_audio_for_text(
                                target_text, voice_code, rate_code, pitch_code, volume_code
                            )
                            st.success(f"Đã tạo âm thanh ({eng}):")

                        record_play(st.session_state.selected_title, mode_label, voice_option, rate_code)
                        st.session_state["last_audio_bytes"] = combined_audio_bytes

            if st.session_state.get("last_audio_bytes"):
                st.audio(st.session_state["last_audio_bytes"], format="audio/mp3")
                st.download_button(
                    label="⬇️ Tải file âm thanh (.mp3)",
                    data=st.session_state["last_audio_bytes"],
                    file_name="schoolmini_sach_noi.mp3",
                    mime="audio/mpeg",
                )

        with st.container(border=True):
            st.markdown('<span class="waka-badge">🎵 NHẠC NỀN THƯ GIÃN</span>', unsafe_allow_html=True)
            tracks = get_ambient_tracks()
            if not tracks:
                st.caption("Tạo thư mục `ambient/` và thả file `.mp3` / `.wav` để phát nhạc nền.")
            else:
                track_names = ["Không dùng nhạc nền"] + [t.name for t in tracks]
                chosen_track = st.selectbox("Chọn nhạc nền:", track_names)
                if chosen_track != "Không dùng nhạc nền":
                    track_path = AMBIENT_DIR / chosen_track
                    st.audio(track_path.read_bytes(), format="audio/mp3", loop=True)

elif current_tab == "📁 Tải tài liệu":
    with st.container(border=True):
        st.markdown('<span class="waka-badge">NHẬP TÀI LIỆU (.TXT / .DOCX / .PDF)</span>', unsafe_allow_html=True)
        doc_files = st.file_uploader(
            "Chọn một hoặc nhiều file truyện:",
            type=["txt", "docx", "pdf"],
            accept_multiple_files=True,
        )
        if doc_files:
            extracted_map = {uf.name: extract_uploaded_text(uf) for uf in doc_files}
            chosen_doc = st.selectbox("Chọn file:", list(extracted_map.keys()))
            preview_txt = extracted_map[chosen_doc]
            st.text_area("Xem trước nội dung:", value=preview_txt, height=220)
            if st.button("📖 Nạp & Chuyển sang Đọc"):
                st.session_state.selected_title = chosen_doc
                st.session_state.story_text = preview_txt
                st.session_state.current_para_idx = 0
                st.session_state.active_tab = "📖 Đọc & Nghe"
                st.rerun()

elif current_tab == "🕒 Lịch sử & Thống kê":
    col_fav, col_stat = st.columns([1, 1.4], gap="large")

    with col_fav:
        with st.container(border=True):
            st.markdown('<span class="waka-badge">TRUYỆN YÊU THÍCH &amp; ĐỌC DỞ</span>', unsafe_allow_html=True)
            fav_list = state_data["favorites"]
            if not fav_list:
                st.info("Chưa có truyện yêu thích nào.")
            else:
                for idx, fav_title in enumerate(fav_list):
                    c_a, c_b = st.columns([2, 1])
                    with c_a:
                        bm = bookmarks_map.get(fav_title)
                        bm_txt = f" (Đoạn {bm + 1})" if bm is not None else ""
                        st.markdown(f"❤️ **{fav_title}**{bm_txt}")
                    with c_b:
                        if st.button("Mở đọc", key=f"fav_open_{idx}_{fav_title}"):
                            if fav_title in stories:
                                st.session_state.selected_title = fav_title
                                st.session_state.story_text = stories[fav_title]["content"]
                                st.session_state.current_para_idx = bookmarks_map.get(fav_title, 0)
                                st.session_state.active_tab = "📖 Đọc & Nghe"
                                st.rerun()

    with col_stat:
        with st.container(border=True):
            st.markdown('<span class="waka-badge">THỐNG KÊ LƯỢT NGHE</span>', unsafe_allow_html=True)
            if play_counts_map:
                st.bar_chart(play_counts_map)
            else:
                st.info("Chưa có dữ liệu lượt nghe.")

    with st.container(border=True):
        st.markdown('<span class="waka-badge">LỊCH SỬ NGHE GẦN ĐÂY</span>', unsafe_allow_html=True)
        hist_list = state_data["history"]
        if not hist_list:
            st.info("Chưa có lịch sử nghe.")
        else:
            st.dataframe(hist_list, use_container_width=True, hide_index=True)
            if st.button("🗑️ Xóa lịch sử nghe"):
                new_state = load_state()
                new_state["history"] = []
                save_state(new_state)
                st.rerun()