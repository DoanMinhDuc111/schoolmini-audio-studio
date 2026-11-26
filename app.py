import asyncio
import base64
import html
import io
import json
import math
import re
import struct
import wave
from datetime import datetime
from pathlib import Path

import docx
import edge_tts
from gtts import gTTS
from pypdf import PdfReader
import streamlit as st

try:
    from pydub import AudioSegment
    PYDUB_AVAILABLE = True
except Exception:
    PYDUB_AVAILABLE = False

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
    svg = f"""<svg xmlns="http://www.w3.org/2000/svg" width="320" height="220" viewBox="0 0 320 220">
      <defs>
        <linearGradient id="g" x1="0" y1="0" x2="1" y2="1">
          <stop offset="0%" stop-color="{c1}"/>
          <stop offset="100%" stop-color="{c2}"/>
        </linearGradient>
      </defs>
      <rect width="320" height="220" rx="14" fill="url(#g)"/>
      <circle cx="275" cy="36" r="52" fill="rgba(255,255,255,0.14)"/>
      <circle cx="42" cy="195" r="44" fill="rgba(0,0,0,0.18)"/>
      <rect x="16" y="16" width="104" height="24" rx="12" fill="rgba(0,0,0,0.35)"/>
      <text x="68" y="32" fill="#ffffff" font-family="Arial, sans-serif" font-size="11" font-weight="bold" text-anchor="middle">{safe_tag}</text>
      <text x="16" y="96" fill="#ffffff" font-family="Arial, sans-serif" font-size="18" font-weight="bold">{safe_title[:20]}</text>
      <text x="16" y="120" fill="rgba(255,255,255,0.95)" font-family="Arial, sans-serif" font-size="16" font-weight="bold">{safe_title[20:40]}</text>
      <text x="16" y="192" fill="rgba(255,255,255,0.88)" font-family="Arial, sans-serif" font-size="12">schoolmini • Sách Nói</text>
    </svg>"""
    b64 = base64.b64encode(svg.encode("utf-8")).decode("ascii")
    return f"data:image/svg+xml;base64,{b64}"


DEFAULT_STORIES = {
    "Miếu ông Trường con": {
        "genre": "Hài hước",
        "c1": "#da251d",
        "c2": "#7f1d1d",
        "content": """Ở đầu con ngõ nhỏ rợp bóng cây xanh của khu phố, có một ngôi miếu nhỏ lợp ngói rêu phong mà người dân quen gọi bằng cái tên nghe vừa thân quen vừa lạ tai: Miếu ông Trường con. Chẳng ai nhớ miếu dựng từ bao giờ, chỉ biết quanh năm hương khói nghi ngút, còn nhân vật chính của những câu chuyện truyền miệng lại là một "ông" chẳng giống ai: ông Trường con — dáng người lanh lợi, tính tình hài hước, lúc nào cũng xuất hiện đúng lúc hàng quán đông vui nhất.

Chuyện về ông Trường con lan đi khắp xóm nhờ một "đặc sản" không lẫn vào đâu được: đi ăn không bao giờ trả tiền.

Buổi sáng, ở quán bún riêu đầu ngõ, khi nồi nước dùng đang bốc khói nghi ngút và khách ra vào tấp nập, ông Trường con thong thả bước tới, chọn chiếc bàn ngay góc mát rồi gọi một tô đầy đủ: thêm giò, thêm chả, thêm đậu rán. Ăn xong, ông thong dong lau miệng, khen nước dùng hôm nay đậm đà ngọt thanh, rồi cười xòa vỗ vai chủ quán:
"Hôm nay coi như gửi lộc cho miếu ông Trường con nhé, buôn may bán đắt!"
Nói rồi ông phẩy tay bước đi nhẹ tênh, để lại bà chủ quán vừa lắc đầu vừa bật cười vì chẳng bao giờ giận nổi cái giọng hóm hỉnh ấy.

Đến trưa, sang quán cơm bình dân đối diện, ông Trường con lại xuất hiện. Ông gắp đủ món từ cá kho tộ, thịt rang cháy cạnh đến canh chua rau muống, ăn ngon lành chẳng kém ai. Tới lúc tính tiền, ông lại giở bài cũ: nhờ ghi sổ "miếu Trường con", rồi quay sang trêu chọc mấy người khách ngồi bàn bên:
"Ăn cùng tôi hôm nay là có phước lớn đấy nhé, chiều về kiểu gì cũng gặp chuyện vui!"

Chiều tối, ở quán chè hay xe bánh mì nướng muối ớt cuối phố, kịch bản cũ vẫn lặp lại. Người mới dọn tới lần đầu thấy vậy thì ngơ ngác, tưởng ông Trường con quỵt tiền thật, định đứng dậy nhắc nhở. Nhưng mấy người khách quen ngồi cạnh chỉ cười bảo:
"Cứ để ông ấy đi, hôm nào quán vắng khách hoặc gặp chuyện xui xẻo, tự nhiên lại thấy đông người ghé vào, nghe đâu là nhờ ông Trường con chiếu cố đấy!"

Và rồi vào một buổi chiều mưa lất phất, khi hàng quán vãn khách, ông Trường con lại ghé vào quán nhỏ quen thuộc. Trước lúc đứng dậy, ông mỉm cười nhìn mọi người:
"Người ta đi ăn thì trả bằng tiền, còn tôi đi ăn thì trả bằng tiếng cười và sự rôm rả. Có thế quán mới vui, xóm mới ấm!"

Từ hôm ấy, mỗi lần nhắc tới cái tên Miếu ông Trường con, người ta không chỉ nhớ tới chuyện ông đi ăn chẳng bao giờ rút ví, mà còn nhớ tới niềm vui, sự gắn kết và những tràng cười giòn tan mà ông mang lại cho cả con phố nhỏ.""",
    },
    "Ngọn đèn dầu cũ trên gác xép": {
        "genre": "Tản văn",
        "c1": "#b91c1c",
        "c2": "#450a0a",
        "content": """Mưa đầu mùa rơi lách tách trên mái tôn cũ kỹ của ngôi nhà nhỏ ven sông. Minh ngồi bên khung cửa sổ, tay xoay nhẹ chiếc chìa khóa đồng rỉ sét mà ông nội để lại trước khi đi xa. Trong căn phòng chật hẹp chỉ có tiếng mưa và tiếng tích tắc của chiếc đồng hồ treo tường, cậu bỗng thấy lòng mình lắng lại.

Suốt ba năm lên thành phố học tập và làm việc, Minh luôn vội vã. Ngày nào cậu cũng chạy đua với những bản kế hoạch, những cuộc họp kéo dài và những tin nhắn không hồi âm. Có những lúc mệt mỏi quá, Minh quên mất lần cuối mình hít thở trọn vẹn một buổi sáng trong lành là khi nào.

Cầm chiếc chìa khóa bước lên gác xép, Minh đẩy cánh cửa gỗ kêu kẽo kẹt. Lớp bụi mỏng phủ trên những chồng sách cũ và chiếc rương sơn đỏ đã bạc màu. Khi mở chiếc rương ra, thứ đầu tiên hiện lên không phải vàng bạc hay giấy tờ quan trọng, mà là một cuốn sổ tay bìa da sờn mép cùng một ngọn đèn dầu nhỏ.

Trang đầu tiên của cuốn sổ viết nét chữ ngay ngắn:
"Con trai, cuộc đời giống như đi đường dài. Đi nhanh chưa chắc đã tới sớm, quan trọng là biết lúc nào nên dừng lại để nhìn xem mình đang đi đâu."

Minh mỉm cười, bật quẹt thắp lên ngọn đèn dầu. Ánh lửa vàng ấm áp bập bùng soi sáng cả góc phòng tối. Ngoài kia, cơn mưa vẫn rơi, nhưng trong lòng Minh, những bộn bề suốt bao lâu nay dường như đã nhẹ đi rất nhiều.""",
    },
    "Chuyến tàu đêm cuối tuần": {
        "genre": "Đời thường",
        "c1": "#dc2626",
        "c2": "#7c2d12",
        "content": """Năm giờ rưỡi chiều thứ Sáu, Khoa vơ vội chiếc balo rồi lao ra khỏi văn phòng trước khi đồng hồ kịp điểm phút thứ năm. Suốt cả tuần liền, anh dán mắt vào màn hình sửa từng dòng báo cáo, nên giờ chỉ mong kịp lên chuyến tàu lúc sáu giờ tối để về quê thăm mẹ.

Vừa tới ga, dòng người xếp hàng mua vé đã dài ngoằng. Khoa thở phào khi nhớ ra mình đã đặt vé điện tử từ trưa. Anh bước nhanh qua cửa soát vé, tìm đúng toa số bốn và ngồi xuống cạnh ô cửa kính. Ngoài sân ga, những ánh đèn vàng bắt đầu bật sáng, lướt qua vội vã khi con tàu từ từ chuyển bánh.

Đối diện Khoa là một cô bé ngồi cùng mẹ, tay ôm chặt chú gấu bông cũ. Thấy tàu chạy nhanh, cô bé tò mò hỏi:
"Mẹ ơi, sao cột điện ngoài kia cứ chạy ngược lại thế ạ?"
Người mẹ mỉm cười xoa đầu con:
"Vì tàu của mình đang tiến về phía trước đấy con. Cũng giống như khi mình cố gắng đi tới, mọi thứ xung quanh sẽ đổi khác."

Khoa nghe câu trả lời ấy mà bất giác mỉm cười. Đôi khi, chỉ cần bước ra khỏi nhịp sống vội vã thường ngày và ngồi trên một chuyến tàu chậm rãi, người ta mới nhận ra mình đã đi được bao xa. Gió chiều mát lạnh lùa qua khe cửa, mang theo mùi đồng lúa quen thuộc báo hiệu quê nhà đã ở rất gần.""",
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


def parse_dialogue_segments(text: str) -> list[tuple[str, str]]:
    pattern = re.compile(r'(["“][^"”\n]{2,}["”])')
    parts = pattern.split(text)
    segments = []
    for part in parts:
        clean = part.strip()
        if not clean:
            continue
        if (clean.startswith('"') and clean.endswith('"')) or (clean.startswith("“") and clean.endswith("”")):
            spoken = clean.strip('"“”').strip()
            if spoken:
                segments.append(("character", spoken))
        else:
            segments.append(("narrator", clean))
    return segments


def synthesize_builtin_bgm_wav(preset: str, duration_sec: float = 12.0, sample_rate: int = 22050) -> bytes:
    total_samples = int(duration_sec * sample_rate)
    buf = io.BytesIO()
    with wave.open(buf, "wb") as wf:
        wf.setnchannels(1)
        wf.setsampwidth(2)
        wf.setframerate(sample_rate)

        frames = bytearray()
        if "Ru ngủ" in preset:
            chords = [
                (261.63, 329.63, 392.00),
                (220.00, 261.63, 329.63),
                (174.61, 220.00, 261.63),
                (196.00, 246.94, 293.66),
            ]
        elif "Piano ấm" in preset:
            chords = [
                (293.66, 369.99, 440.00),
                (246.94, 293.66, 369.99),
                (196.00, 246.94, 293.66),
                (220.00, 277.18, 329.63),
            ]
        else:
            chords = [
                (196.00, 246.94, 392.00),
                (164.81, 246.94, 329.63),
                (174.61, 220.00, 349.23),
                (196.00, 293.66, 392.00),
            ]

        chord_len = total_samples // len(chords)
        for i in range(total_samples):
            t = i / sample_rate
            c_idx = min(i // chord_len, len(chords) - 1)
            local_t = (i % chord_len) / chord_len
            env = math.sin(math.pi * local_t) ** 0.65
            f1, f2, f3 = chords[c_idx]
            val = (
                0.45 * math.sin(2 * math.pi * f1 * t)
                + 0.35 * math.sin(2 * math.pi * f2 * t)
                + 0.25 * math.sin(2 * math.pi * f3 * t)
            ) * env * 0.22
            sample_int = max(-32767, min(32767, int(val * 32767)))
            frames.extend(struct.pack("<h", sample_int))
        wf.writeframes(bytes(frames))
    return buf.getvalue()


def mix_voice_with_bgm(voice_mp3_bytes: bytes, bgm_source: str, bgm_db_drop: int = 16) -> tuple[bytes, bool]:
    if not PYDUB_AVAILABLE or bgm_source == "Không trộn nhạc nền":
        return voice_mp3_bytes, False

    try:
        voice_seg = AudioSegment.from_file(io.BytesIO(voice_mp3_bytes), format="mp3")
        if bgm_source.startswith("File: "):
            fname = bgm_source.replace("File: ", "", 1)
            bgm_seg = AudioSegment.from_file(AMBIENT_DIR / fname)
        else:
            wav_bytes = synthesize_builtin_bgm_wav(bgm_source, duration_sec=12.0)
            bgm_seg = AudioSegment.from_file(io.BytesIO(wav_bytes), format="wav")

        bgm_seg = bgm_seg - bgm_db_drop
        target_len = len(voice_seg) + 1800
        if len(bgm_seg) < target_len:
            repeats = (target_len // len(bgm_seg)) + 1
            bgm_seg = bgm_seg * repeats
        bgm_seg = bgm_seg[:target_len].fade_in(600).fade_out(1400)

        mixed = bgm_seg.overlay(voice_seg, position=600)
        out_buf = io.BytesIO()
        mixed.export(out_buf, format="mp3", bitrate="128k")
        return out_buf.getvalue(), True
    except Exception:
        return voice_mp3_bytes, False


def build_story_companion(title: str, text: str) -> dict:
    paras = split_paragraphs(text)
    first_p = paras[0] if paras else ""
    last_p = paras[-1] if paras else ""
    dialogues = [seg for role, seg in parse_dialogue_segments(text) if role == "character"]

    summary = (
        f"{first_p[:160].rstrip()}... Và khép lại với thông điệp: {last_p[:160].rstrip()}."
        if len(paras) >= 2
        else first_p[:240]
    )

    if "Trường con" in title or "tiếng cười" in text.lower():
        lesson = "Niềm vui, sự hóm hỉnh và tinh thần gắn kết xóm giềng đôi khi còn quý giá hơn cả vật chất."
    elif "mẹ" in text.lower() or "quê" in text.lower() or "ông nội" in text.lower():
        lesson = "Giữa nhịp sống bận rộn, hãy biết chậm lại để trân trọng gia đình và những ký ức ấm áp."
    else:
        lesson = "Mỗi trải nghiệm nhỏ trong cuộc sống đều mang lại một góc nhìn tích cực nếu ta biết lắng nghe."

    questions = [
        f"1. Chi tiết nào trong truyện '{title}' khiến bạn thấy thú vị hoặc nhớ nhất?",
        f"2. Nếu gặp nhân vật chính ngoài đời thực, bạn sẽ nói gì với nhân vật ấy?",
        "3. Theo bạn, bài học ý nghĩa nhất rút ra sau khi nghe câu chuyện này là gì?",
    ]
    if dialogues:
        questions[1] = f'2. Bạn nghĩ sao về câu nói trong truyện: "{dialogues[0][:90]}"?'

    return {
        "summary": summary,
        "lesson": lesson,
        "questions": questions,
        "dialogue_count": len(dialogues),
    }


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


def go_home_callback() -> None:
    st.session_state.active_tab = "📚 Tủ truyện"
    st.session_state["nav_radio"] = "📚 Tủ truyện"
    st.session_state.preview_story_title = None


def open_story_callback(title: str, content: str, para_idx: int) -> None:
    st.session_state.selected_title = title
    st.session_state.story_text = content
    st.session_state.current_para_idx = para_idx
    st.session_state.active_tab = "📖 Đọc & Nghe"
    st.session_state["nav_radio"] = "📖 Đọc & Nghe"
    st.session_state.preview_story_title = None


def open_preview_callback(title: str) -> None:
    st.session_state.preview_story_title = title


def close_preview_callback() -> None:
    st.session_state.preview_story_title = None


def render_preview_body(title: str, stories_dict: dict, fav_set: set, bm_map: dict, plays_map: dict) -> None:
    info = stories_dict.get(title, {})
    content_str = info.get("content", "")
    words = len(content_str.split())
    mins = max(1, round(words / 150))
    plays = plays_map.get(title, 0)
    bm_idx = bm_map.get(title)
    bm_str = f"Đoạn {bm_idx + 1}" if bm_idx is not None else "Chưa đọc dở"
    genre = info.get("genre", "Khác")
    cover_uri = make_cover_svg_data_uri(
        title,
        genre.upper(),
        info.get("c1", "#da251d"),
        info.get("c2", "#7f1d1d"),
    )

    st.markdown(
        f"""
        <div class="msg-preview-card">
            <img class="msg-preview-banner" src="{cover_uri}" alt="{html.escape(title)}" />
            <div class="msg-preview-title">{html.escape(title)}</div>
            <div class="msg-preview-chips">
                <span class="msg-chip">{html.escape(genre)}</span>
                <span class="msg-chip">{words} từ • ~{mins} phút</span>
                <span class="msg-chip">{plays} lượt nghe</span>
                <span class="msg-chip">{html.escape(bm_str)}</span>
            </div>
            <div class="msg-preview-excerpt">{html.escape(content_str[:340])}{'...' if len(content_str) > 340 else ''}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    p_col1, p_col2, p_col3 = st.columns([1.4, 1, 1])
    with p_col1:
        st.button(
            "▶️ Đọc & Nghe ngay",
            key=f"dlg_open_{title}",
            on_click=open_story_callback,
            args=(title, content_str, bm_idx if bm_idx is not None else 0),
        )
    with p_col2:
        fav_txt = "❤️ Bỏ thích" if title in fav_set else "🤍 Yêu thích"
        if st.button(fav_txt, key=f"dlg_fav_{title}"):
            toggle_favorite(title)
            st.rerun()
    with p_col3:
        st.button("Đóng", key=f"dlg_close_{title}", on_click=close_preview_callback)


if hasattr(st, "dialog"):
    @st.dialog("🔍 Xem trước truyện")
    def story_preview_dialog(title: str, stories_dict: dict, fav_set: set, bm_map: dict, plays_map: dict) -> None:
        render_preview_body(title, stories_dict, fav_set, bm_map, plays_map)


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
if "reader_font_size" not in st.session_state:
    st.session_state.reader_font_size = 16
if "current_para_idx" not in st.session_state:
    st.session_state.current_para_idx = 0
if "preview_story_title" not in st.session_state:
    st.session_state.preview_story_title = None

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

    /* Header đỏ cờ Việt Nam rực phía trên & mờ dần xuống nền tối */
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
        padding: 24px 24px 48px 24px;
        margin-bottom: -168px;
        z-index: 1;
        pointer-events: none;
    }
    .schoolmini-center-box {
        position: relative;
        z-index: 2;
        text-align: center;
        padding-top: 34px;
        padding-bottom: 12px;
    }
    .schoolmini-center-title {
        font-size: 2rem !important;
        font-weight: 800 !important;
        color: #ffffff !important;
        letter-spacing: -0.025em;
        margin: 0 auto 8px auto !important;
        text-shadow: 0 3px 14px rgba(0, 0, 0, 0.45);
    }
    .schoolmini-center-sub {
        font-size: 0.96rem !important;
        font-weight: 500 !important;
        color: #fff5f5 !important;
        opacity: 0.96;
        margin: 0 auto !important;
        text-shadow: 0 1px 8px rgba(0, 0, 0, 0.45);
    }

    .st-key-brand_home_btn {
        position: relative;
        z-index: 25 !important;
        display: inline-block !important;
        width: auto !important;
        margin-left: 18px !important;
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
        font-size: 1.95rem !important;
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
        color: #ffffff !important;
        transform: translateX(2px) !important;
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
        font-size: 0.76rem;
        padding: 4px 12px;
        border-radius: 999px;
        margin-bottom: 10px;
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
        font-size: 0.88rem;
        border: none;
        border-radius: 10px;
        padding: 0.45rem 0.65rem;
        width: 100%;
        transition: background-color 0.2s ease, transform 0.15s ease;
    }
    .stButton:not(.st-key-brand_home_btn) > button:first-child:hover {
        background-color: #da251d;
        color: #ffffff !important;
    }

    /* Thẻ truyện dạng dọc gọn đẹp trên cả điện thoại & máy tính */
    .story-v-card {
        background: linear-gradient(180deg, #162033 0%, #101726 100%);
        border: 1px solid #26354d;
        border-radius: 16px;
        padding: 10px;
        margin-bottom: 8px;
        box-shadow: 0 6px 16px rgba(0, 0, 0, 0.28);
        transition: transform 0.2s ease, border-color 0.2s ease;
    }
    .story-v-card:active {
        transform: scale(0.98);
        border-color: #da251d;
    }
    .story-v-img {
        width: 100%;
        aspect-ratio: 4 / 3;
        object-fit: cover;
        border-radius: 11px;
        display: block;
        margin-bottom: 9px;
    }
    .story-v-genre {
        display: inline-block;
        font-size: 0.7rem;
        font-weight: 700;
        color: #fca5a5 !important;
        background: rgba(218, 37, 29, 0.18);
        border: 1px solid rgba(248, 113, 113, 0.25);
        padding: 2px 8px;
        border-radius: 999px;
        margin-bottom: 5px;
    }
    .story-v-title {
        font-size: 0.98rem;
        font-weight: 700;
        color: #f8fafc !important;
        line-height: 1.3;
        min-height: 2.55em;
        display: -webkit-box;
        -webkit-line-clamp: 2;
        -webkit-box-orient: vertical;
        overflow: hidden;
        margin-bottom: 4px;
    }
    .story-v-meta {
        font-size: 0.76rem;
        color: #94a3b8 !important;
    }

    /* Khung xem trước phóng to kiểu Messenger */
    .msg-preview-card {
        background: #0f172a;
        border: 1px solid #334155;
        border-radius: 18px;
        padding: 14px;
        margin-bottom: 12px;
        box-shadow: 0 18px 40px rgba(0, 0, 0, 0.65);
    }
    .msg-preview-banner {
        width: 100%;
        max-height: 190px;
        object-fit: cover;
        border-radius: 12px;
        margin-bottom: 12px;
        display: block;
    }
    .msg-preview-title {
        font-size: 1.25rem;
        font-weight: 800;
        color: #ffffff !important;
        margin-bottom: 8px;
    }
    .msg-preview-chips {
        display: flex;
        flex-wrap: wrap;
        gap: 6px;
        margin-bottom: 10px;
    }
    .msg-chip {
        font-size: 0.76rem;
        font-weight: 600;
        color: #fde047 !important;
        background: rgba(218, 37, 29, 0.22);
        border: 1px solid rgba(253, 224, 71, 0.25);
        padding: 3px 9px;
        border-radius: 999px;
    }
    .msg-preview-excerpt {
        font-size: 0.9rem;
        line-height: 1.55;
        color: #e2e8f0 !important;
        background: #162033;
        padding: 10px 12px;
        border-radius: 12px;
        border-left: 3px solid #da251d;
    }
</style>
""",
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="schoolmini-hero-bg">
        <div class="schoolmini-center-box">
            <div class="schoolmini-center-title">Thư viện sách của Trường con</div>
            <p class="schoolmini-center-sub">Sách nói Phân vai &amp; Podcast Kể chuyện AI — chạm vào thẻ để xem trước chi tiết.</p>
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
            '<span class="waka-badge">KHÁM PHÁ TỦ TRUYỆN • CHẠM XEM NHANH ĐỂ PHÓNG TO</span>',
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

        # Hiển thị popup phóng to kiểu Messenger khi bấm Xem nhanh
        preview_title = st.session_state.preview_story_title
        if preview_title and preview_title in stories:
            if hasattr(st, "dialog"):
                story_preview_dialog(preview_title, stories, favorites_set, bookmarks_map, play_counts_map)
            else:
                with st.container(border=True):
                    st.markdown('<span class="waka-badge">🔍 XEM TRƯỚC CHI TIẾT</span>', unsafe_allow_html=True)
                    render_preview_body(preview_title, stories, favorites_set, bookmarks_map, play_counts_map)

        if not filtered_items:
            st.info("Không tìm thấy truyện phù hợp.")
        else:
            # Lưới thẻ dọc: 3 thẻ mỗi hàng (tự gọn 2-3 cột responsive trên điện thoại)
            cols_per_row = 3
            for row_start in range(0, len(filtered_items), cols_per_row):
                row_items = filtered_items[row_start : row_start + cols_per_row]
                grid_cols = st.columns(cols_per_row, gap="small")
                for offset, (title, info) in enumerate(row_items):
                    idx = row_start + offset
                    content_str = info.get("content", "")
                    words = len(content_str.split())
                    mins = max(1, round(words / 150))
                    plays = play_counts_map.get(title, 0)
                    fav_icon = "❤️ " if title in favorites_set else ""
                    genre_label = info.get("genre", "Khác")

                    cover_uri = make_cover_svg_data_uri(
                        title,
                        genre_label.upper(),
                        info.get("c1", "#da251d"),
                        info.get("c2", "#7f1d1d"),
                    )

                    with grid_cols[offset]:
                        st.markdown(
                            f"""
                            <div class="story-v-card">
                                <img class="story-v-img" src="{cover_uri}" alt="{html.escape(title)}" />
                                <span class="story-v-genre">{html.escape(genre_label)}</span>
                                <div class="story-v-title">{fav_icon}{html.escape(title)}</div>
                                <div class="story-v-meta">~{mins} phút • {plays} nghe</div>
                            </div>
                            """,
                            unsafe_allow_html=True,
                        )
                        b1, b2 = st.columns(2, gap="small")
                        with b1:
                            st.button(
                                "🔍 Xem",
                                key=f"prev_{idx}_{title}",
                                help="Phóng to xem trước chi tiết",
                                on_click=open_preview_callback,
                                args=(title,),
                            )
                        with b2:
                            st.button(
                                "▶️ Đọc",
                                key=f"open_{idx}_{title}",
                                help="Mở khung Đọc & Nghe",
                                on_click=open_story_callback,
                                args=(title, content_str, bookmarks_map.get(title, 0)),
                            )

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
                'Nội dung truyện (đặt lời thoại nhân vật trong ngoặc kép "..." để bật Kể chuyện phân vai tự động):',
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
    left_col, right_col = st.columns([1.3, 1], gap="large")

    paragraphs = split_paragraphs(st.session_state.story_text)
    if paragraphs and st.session_state.current_para_idx >= len(paragraphs):
        st.session_state.current_para_idx = 0

    with left_col:
        with st.container(border=True):
            st.markdown('<span class="waka-badge">KHUNG ĐỌC TRUYỆN &amp; TÔ SÁNG ĐOẠN</span>', unsafe_allow_html=True)

            t_col1, t_col2, t_col3 = st.columns([2, 1, 0.9])
            with t_col1:
                st.subheader(st.session_state.selected_title)
            with t_col2:
                theme_choice = st.selectbox(
                    "Nền đọc:",
                    ["Tối", "Giấy vàng", "Trắng sáng"],
                    index=["Tối", "Giấy vàng", "Trắng sáng"].index(st.session_state.reader_theme),
                )
                st.session_state.reader_theme = theme_choice
            with t_col3:
                font_sz = st.selectbox(
                    "Cỡ chữ:",
                    [15, 16, 18, 20],
                    index=[15, 16, 18, 20].index(st.session_state.reader_font_size),
                )
                st.session_state.reader_font_size = font_sz

            st.session_state.story_text = st.text_area(
                "Nội dung truyện (có thể chỉnh sửa trực tiếp):",
                value=st.session_state.story_text,
                height=165,
            )

            theme_styles = {
                "Tối": ("#0b1120", "#f1f5f9", "#26354d", "rgba(218, 37, 29, 0.18)"),
                "Giấy vàng": ("#fbf3dc", "#3b2f1e", "#e6d5a8", "rgba(218, 37, 29, 0.12)"),
                "Trắng sáng": ("#ffffff", "#1e293b", "#e2e8f0", "rgba(218, 37, 29, 0.10)"),
            }
            bg_c, fg_c, bd_c, hl_c = theme_styles[st.session_state.reader_theme]
            refreshed_paras = split_paragraphs(st.session_state.story_text)

            para_blocks = []
            for i, p in enumerate(refreshed_paras):
                is_active = i == st.session_state.current_para_idx
                active_style = (
                    f"background:{hl_c}; border-left: 3px solid #da251d; padding: 6px 10px; border-radius: 6px;"
                    if is_active
                    else "padding: 2px 4px;"
                )
                para_blocks.append(
                    f"<p style='color:{fg_c} !important; font-size:{font_sz}px; line-height:1.72; margin-bottom:10px; {active_style}'>"
                    f"<strong style='opacity:0.72;'>[{i+1}]</strong> {html.escape(p)}</p>"
                )
            preview_html = "".join(para_blocks) or f"<p style='color:{fg_c} !important;'>Chưa có nội dung.</p>"

            st.markdown(
                f"""
                <div style="background:{bg_c}; border:1px solid {bd_c}; border-radius:10px; padding:14px 16px; max-height:240px; overflow-y:auto;">
                    {preview_html}
                </div>
                """,
                unsafe_allow_html=True,
            )

        with st.container(border=True):
            st.markdown('<span class="waka-badge">💡 GÓC TƯƠNG TÁC &amp; ĐỌC CÙNG CON</span>', unsafe_allow_html=True)
            companion = build_story_companion(st.session_state.selected_title, st.session_state.story_text)
            st.markdown(f"**Tóm tắt nhanh:** {companion['summary']}")
            st.markdown(f"**Thông điệp câu chuyện:** {companion['lesson']}")
            st.markdown("**Câu hỏi gợi mở sau khi nghe:**")
            for q in companion["questions"]:
                st.markdown(f"- {q}")

            if st.button("🔊 Đọc phần tóm tắt & câu hỏi tương tác", key="read_companion_btn"):
                companion_script = (
                    f"Tóm tắt truyện {st.session_state.selected_title}. {companion['summary']} "
                    f"Thông điệp câu chuyện: {companion['lesson']} "
                    f"Câu hỏi gợi mở dành cho bạn: {' '.join(companion['questions'])}"
                )
                with st.spinner("Đang tạo giọng đọc tương tác..."):
                    c_audio, _ = generate_audio_for_text(
                        companion_script, "vi-VN-HoaiMyNeural", "+0%", "+2Hz", "+0%"
                    )
                    st.session_state["companion_audio_bytes"] = c_audio

            if st.session_state.get("companion_audio_bytes"):
                st.audio(st.session_state["companion_audio_bytes"], format="audio/mp3")

        with st.container(border=True):
            st.markdown('<span class="waka-badge">📝 GHI CHÚ &amp; XUẤT TỆP (.TXT / .MD)</span>', unsafe_allow_html=True)
            cur_title = st.session_state.selected_title
            existing_note = notes_map.get(cur_title, "")
            note_val = st.text_area(
                "Ghi chú cá nhân cho truyện này:",
                value=existing_note,
                height=90,
                placeholder="Ghi lại ý hay, nhân vật, cảm nhận...",
            )
            if note_val != existing_note:
                save_note(cur_title, note_val)

            export_payload = (
                f"# {cur_title}\n\n"
                f"## Nội dung truyện\n\n{st.session_state.story_text}\n\n"
                f"## Thông điệp & Câu hỏi gợi mở\n\n- {companion['lesson']}\n"
                + "\n".join(f"- {q}" for q in companion["questions"])
                + f"\n\n## Ghi chú cá nhân\n\n{note_val or '(Không có ghi chú)'}\n"
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
            st.markdown('<span class="waka-badge">🎙️ STUDIO SÁCH NÓI PHÂN VAI &amp; MIX NHẠC</span>', unsafe_allow_html=True)

            read_mode = st.radio(
                "Phạm vi nội dung đọc:",
                ["Toàn bộ truyện", "Từng đoạn", "Ghép nhiều đoạn (MP3 liền mạch)"],
                horizontal=False,
            )

            target_text = st.session_state.story_text
            mode_label = "Toàn bộ"
            selected_slice = refreshed_paras

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

            multi_voice_mode = st.checkbox(
                '🎭 Kể chuyện phân vai tự động (Tự đổi giọng khi gặp lời thoại "...")',
                value=True,
            )

            dialogue_segments = parse_dialogue_segments(target_text)
            char_lines = sum(1 for r, _ in dialogue_segments if r == "character")
            words = len(target_text.split()) if target_text.strip() else 0
            est_mins = max(1, round(words / 150)) if words > 0 else 0
            st.caption(f"Đang chọn: {words} từ • ~{est_mins} phút • Phát hiện {char_lines} câu thoại nhân vật")

            if multi_voice_mode:
                v_col1, v_col2 = st.columns(2)
                with v_col1:
                    narrator_choice = st.selectbox(
                        "Giọng người dẫn:",
                        ("Nam miền Bắc (Nam Minh)", "Nữ miền Nam (Hoài My)"),
                        index=0,
                    )
                with v_col2:
                    character_choice = st.selectbox(
                        "Giọng nhân vật thoại:",
                        ("Nữ miền Nam (Hoài My)", "Nam miền Bắc (Nam Minh)"),
                        index=0,
                    )
                voice_option = f"Phân vai ({narrator_choice.split()[0]} + {character_choice.split()[0]})"
            else:
                voice_option = st.selectbox(
                    "Giọng đọc:",
                    ("Nữ miền Nam (Hoài My)", "Nam miền Bắc (Nam Minh)"),
                )
                narrator_choice = voice_option
                character_choice = voice_option

            s_col1, s_col2 = st.columns(2)
            with s_col1:
                speed_option = st.selectbox(
                    "Tốc độ đọc:",
                    ("Chậm hơn (-15%)", "Bình thường (0%)", "Nhanh hơn (+15%)", "Nhanh (+25%)"),
                    index=1,
                )
            with s_col2:
                pitch_option = st.selectbox(
                    "Cao độ giọng:",
                    ("Trầm hơn (-5Hz)", "Bình thường (+0Hz)", "Cao hơn (+5Hz)"),
                    index=1,
                )

            volume_option = st.selectbox(
                "Âm lượng giọng đọc:",
                ("Nhỏ hơn (-15%)", "Bình thường (+0%)", "Lớn hơn (+15%)"),
                index=1,
            )

            tracks = get_ambient_tracks()
            bgm_options = [
                "Không trộn nhạc nền",
                "✨ Nhạc nền tích hợp: Ru ngủ êm dịu",
                "✨ Nhạc nền tích hợp: Piano ấm áp",
                "✨ Nhạc nền tích hợp: Mưa đêm thư giãn",
            ] + [f"File: {t.name}" for t in tracks]

            bgm_choice = st.selectbox(
                "🎵 Trộn nhạc nền podcast vào file sách nói:",
                bgm_options,
                index=1,
            )

            rate_code = speed_option.split("(")[-1].replace(")", "")
            pitch_code = pitch_option.split("(")[-1].replace(")", "")
            volume_code = volume_option.split("(")[-1].replace(")", "")

            narrator_code = "vi-VN-HoaiMyNeural" if "Hoài My" in narrator_choice else "vi-VN-NamMinhNeural"
            character_code = "vi-VN-HoaiMyNeural" if "Hoài My" in character_choice else "vi-VN-NamMinhNeural"

            if st.button("🔊 Tạo & Nghe sách nói Studio"):
                if not target_text.strip():
                    st.warning("Chưa có nội dung để đọc!")
                else:
                    with st.spinner("Đang tổng hợp giọng đọc & hòa âm..."):
                        merged_buf = bytearray()
                        engine_used = "Edge TTS"

                        if multi_voice_mode and char_lines > 0:
                            for role, seg_txt in dialogue_segments:
                                if role == "character":
                                    v_code = character_code
                                    p_code = "+4Hz" if pitch_code == "+0Hz" else pitch_code
                                else:
                                    v_code = narrator_code
                                    p_code = pitch_code
                                b_seg, eng = generate_audio_for_text(
                                    seg_txt, v_code, rate_code, p_code, volume_code
                                )
                                merged_buf.extend(b_seg)
                                if eng != "Edge TTS":
                                    engine_used = eng
                        elif read_mode == "Ghép nhiều đoạn (MP3 liền mạch)" and refreshed_paras:
                            for seg in selected_slice:
                                b_seg, eng = generate_audio_for_text(
                                    seg, narrator_code, rate_code, pitch_code, volume_code
                                )
                                merged_buf.extend(b_seg)
                                if eng != "Edge TTS":
                                    engine_used = eng
                        else:
                            b_seg, eng = generate_audio_for_text(
                                target_text, narrator_code, rate_code, pitch_code, volume_code
                            )
                            merged_buf.extend(b_seg)
                            engine_used = eng

                        raw_voice_bytes = bytes(merged_buf)
                        final_audio_bytes, was_mixed = mix_voice_with_bgm(raw_voice_bytes, bgm_choice)

                        mix_note = " + đã trộn nhạc nền" if was_mixed else ""
                        role_note = f" • Phân vai ({char_lines} thoại)" if (multi_voice_mode and char_lines > 0) else ""
                        st.success(f"Hoàn tất sách nói ({engine_used}{role_note}{mix_note}):")

                        record_play(
                            st.session_state.selected_title,
                            f"{mode_label}{role_note}",
                            voice_option,
                            rate_code,
                        )
                        st.session_state["last_audio_bytes"] = final_audio_bytes

            if st.session_state.get("last_audio_bytes"):
                st.audio(st.session_state["last_audio_bytes"], format="audio/mp3")
                st.download_button(
                    label="⬇️ Tải file sách nói hoàn chỉnh (.mp3)",
                    data=st.session_state["last_audio_bytes"],
                    file_name="schoolmini_audio_studio.mp3",
                    mime="audio/mpeg",
                )

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