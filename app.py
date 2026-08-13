"""
Document Generation Console (Template Engine Version)
Architecture: Streamlit (UI) -> Gemini API (JSON Structured Output) -> docxtpl (Word Rendering)
"""

import io
import os
from datetime import datetime

import streamlit as st
from docxtpl import DocxTemplate
from docx import Document
from pydantic import BaseModel, Field
from google import genai
from google.genai import types

# =============================================================================
# Page configuration
# =============================================================================
st.set_page_config(layout="wide", page_title="お便り作成コンソール")

# =============================================================================
# Constants & Directory Setup
# =============================================================================
TEMPLATE_DIR = "templates"
os.makedirs(TEMPLATE_DIR, exist_ok=True)

AUDIENCE_OPTIONS = {"student": "生徒向け", "parent": "保護者向け", "grade": "学年全体向け"}
DOCUMENT_TYPE_OPTIONS = {"class_newsletter": "学級通信", "event_notice": "行事のお知らせ", "club_notice": "部活動の連絡"}
LENGTH_OPTIONS = {"short": "約400字", "long": "約1000字"}
GEMINI_MODEL = "gemini-2.5-flash"

# =============================================================================
# Default Template Generation (Auto-setup)
# =============================================================================
def ensure_default_template() -> None:
    """Creates a basic fallback docx template if none exists."""
    default_path = os.path.join(TEMPLATE_DIR, "標準テンプレート.docx")
    if not os.path.exists(default_path):
        doc = Document()
        doc.add_heading("{{ title }}", level=0)
        doc.add_paragraph("{{ greeting }}")
        doc.add_paragraph("{{ body }}")
        doc.add_heading("開催概要", level=1)
        doc.add_paragraph("日時: {{ event_date }} {{ event_time }}")
        doc.add_paragraph("場所: {{ event_place }}")
        doc.add_paragraph("持ち物: {{ event_items }}")
        doc.add_paragraph("{{ closing }}")
        doc.save(default_path)

ensure_default_template()

# =============================================================================
# Data Schema (Structured Output for Gemini)
# =============================================================================
class NewsletterData(BaseModel):
    title: str = Field(description="お便りのタイトル")
    subtitle: str = Field(description="ヘッダー用サブタイトル（例: 3年3組 学級通信 第○号）")
    date_str: str = Field(description="発行日（例: 令和8年○月○日(金)発行）")
    greeting: str = Field(description="導入の温かい挨拶文")
    body: str = Field(description="本文（600~800文字、子供の具体的な様子を含む）")
    parent_note: str = Field(description="★保護者の皆様へ★の注意書き欄の文章")
    closing: str = Field(description="結びの言葉")
    
    # 追加：時間割データ（曜日ごとの授業内容の辞書、またはリスト）
    # 例: {"月曜": ["国語", "数学", "理科", "社会", "英語", "体育"], ...}
    timetable: dict = Field(description="月曜から金曜までの1〜6限の授業予定の辞書")
    
    urls: list[str] = Field(description="参考にしたURLのリスト")

# =============================================================================
# UI Styles
# =============================================================================
def apply_enterprise_styles() -> None:
    st.markdown(
        """
        <style>
        .block-container { padding-top: 1.5rem; padding-bottom: 2rem; max-width: 100%; }
        .console-header { font-size: 1.5rem; font-weight: 600; color: #111827; margin: 0 0 0.25rem 0; }
        .console-subheader { font-size: 0.875rem; color: #6b7280; margin: 0 0 1.5rem 0; }
        .section-label { font-size: 0.7rem; font-weight: 600; letter-spacing: 0.1em; color: #374151; margin: 0 0 0.75rem 0; }
        .system-log { font-family: "Consolas", monospace; font-size: 0.8rem; background: #f9fafb; border: 1px solid #e5e7eb; padding: 0.75rem; }
        </style>
        """,
        unsafe_allow_html=True,
    )

# =============================================================================
# Logic
# =============================================================================
def get_api_key() -> str | None:
    for key_name in ("GEMINI_API_KEY", "GOOGLE_API_KEY"):
        if value := os.environ.get(key_name):
            return value.strip()
    try:
        for key_name in ("GEMINI_API_KEY", "GOOGLE_API_KEY"):
            if key_name in st.secrets:
                return str(st.secrets[key_name]).strip()
    except Exception:
        pass
    return None

def get_available_templates() -> list[str]:
    return [f for f in os.listdir(TEMPLATE_DIR) if f.endswith(".docx")]

def generate_newsletter_data(api_key: str, prompt: str, use_web_search: bool) -> NewsletterData:
    client = genai.Client(api_key=api_key)
    
    tools = [types.Tool(google_search=types.GoogleSearch())] if use_web_search else None
    
    response = client.models.generate_content(
        model=GEMINI_MODEL,
        contents=prompt,
        config=types.GenerateContentConfig(
            tools=tools,
            response_mime_type="application/json",
            response_schema=NewsletterData,
            temperature=0.7,
        ),
    )
    
    # Extract Grounding URLs
    urls = []
    if use_web_search and response.candidates:
        metadata = getattr(response.candidates[0], "grounding_metadata", None)
        if metadata and getattr(metadata, "grounding_chunks", None):
            for chunk in metadata.grounding_chunks:
                if web := getattr(chunk, "web", None):
                    if web.uri and web.uri not in urls:
                        urls.append(web.uri)
                        
    # Parse JSON into Pydantic model
    result_data = NewsletterData.model_validate_json(response.text)
    if urls:
        result_data.urls = urls
        
    return result_data

def render_docx(template_name: str, data: NewsletterData) -> bytes:
    template_path = os.path.join(TEMPLATE_DIR, template_name)
    doc = DocxTemplate(template_path)
    # Convert pydantic model to dict for docxtpl
    context = data.model_dump()
    doc.render(context)
    
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()

# =============================================================================
# Main UI
# =============================================================================
def main() -> None:
    if "doc_data" not in st.session_state:
        st.session_state.doc_data = None

    apply_enterprise_styles()
    api_key = get_api_key()

    with st.sidebar:
        st.markdown("### コントロールパネル")
        audience = st.selectbox("対象読者", options=list(AUDIENCE_OPTIONS.keys()), format_func=lambda k: AUDIENCE_OPTIONS[k])
        doc_type = st.selectbox("文書の種類", options=list(DOCUMENT_TYPE_OPTIONS.keys()), format_func=lambda k: DOCUMENT_TYPE_OPTIONS[k])
        length = st.radio("目安の分量", options=list(LENGTH_OPTIONS.keys()), format_func=lambda k: LENGTH_OPTIONS[k])
        
        st.divider()
        templates = get_available_templates()
        selected_template = st.selectbox("デザインテンプレート", options=templates)
        
        st.divider()
                use_web_search = st.checkbox("Web検索による情報補足", value=False)
                st.caption(f"APIキー状態: {'設定済み' if api_key else '未設定'}")

            # ==========================================
            # 【追加】来週の時間割（1〜6限）入力エリア
            # ==========================================
            st.markdown('<p class="section-label">来週の時間割設定（1〜6限）</p>', unsafe_allow_html=True)
            with st.expander("時間割を個別に設定する（30コマ）", expanded=False):
                st.caption("月曜〜金曜の1〜6限の教科名を入力してください。")
                days = ["月曜日", "火曜日", "水曜日", "木曜日", "金曜日"]
                periods = [f"{i}限" for i in range(1, 7)]
        
                # セッションステートで時間割データを保持
                if "timetable_input" not in st.session_state:
                    st.session_state.timetable_input = {day: ["" for _ in range(6)] for day in days}
        
                current_timetable = {}
                for day in days:
                    st.markdown(f"**{day}**")
                    cols = st.columns(6)
                    day_subjects = []
                    for i, period in enumerate(periods):
                        with cols[i]:
                            val = st.text_input(
                                f"{day}{period}",
                                value=st.session_state.timetable_input[day][i],
                                key=f"tt_{day}_{i}",
                                label_visibility="collapsed",
                                placeholder=period
                            )
                            day_subjects.append(val)
                    current_timetable[day] = day_subjects
                    st.session_state.timetable_input[day] = day_subjects　　　

    st.markdown('<p class="console-header">お便り作成コンソール</p>', unsafe_allow_html=True)
    st.markdown('<p class="console-subheader">Template Engine Architecture (docxtpl + JSON)</p>', unsafe_allow_html=True)

    if not api_key:
        st.warning("警告: APIキーが設定されていません。")

    st.markdown('<p class="section-label">入力エリア</p>', unsafe_allow_html=True)
    draft_notes = st.text_area("伝えたい内容", height=150, label_visibility="collapsed", placeholder="記載したい内容を入力してください...")

    if st.button("お便りを生成する", type="primary", disabled=not api_key):
        if not draft_notes.strip():
            st.error("エラー: 伝えたい内容を入力してください。")
        else:
            with st.status("データを生成しています...", expanded=True) as status:
                try:
                   prompt = f"""
                    あなたはベテラン教員です。以下の制約と構成を厳守して学級通信のデータを作成してください。

                    【絶対守るべき制約】
                    - 全体文字数: 本文（body）は必ず600文字〜800文字に収めること。
                    - 文体: ですます調。温かく前向きなトーン。
                    - 文構造: 読者が読みやすいよう、一文一文をできるだけ短く分割すること。
                    - 構成: 導入(greeting) → 具体：子供の具体的な様子や言葉を入れる(body) → 結び(closing)
                    - その他: 保護者への連絡事項(parent_note)も含めること。
                    - 時間割(timetable): ユーザーが指定した以下の時間割データをそのままJSONの`timetable`に含めて出力すること。
                    
                    【入力された時間割データ】
                    {st.session_state.timetable_input}

                    読者: {AUDIENCE_OPTIONS[audience]} / 種類: {DOCUMENT_TYPE_OPTIONS[doc_type]} / 分量: {LENGTH_OPTIONS[length]}
                    
                    【教員のメモ】
                    {draft_notes}
                    """
                    
                    st.markdown('<pre class="system-log">Status: Generating JSON Payload</pre>', unsafe_allow_html=True)
                    data = generate_newsletter_data(api_key, prompt, use_web_search)
                    st.session_state.doc_data = data
                    status.update(label="成功: データの生成が完了しました。", state="complete")
                    
                except Exception as error:
                    status.update(label="エラー: 生成に失敗しました。", state="error")
                    st.error(f"エラー詳細: {error}")

    st.divider()
    st.markdown('<p class="section-label">出力エリア</p>', unsafe_allow_html=True)

    if st.session_state.doc_data:
        data = st.session_state.doc_data
        
       # 画面上でのプレビュー表示
        with st.expander("生成されたデータ（プレビュー）", expanded=True):
            st.markdown(f"**タイトル:** {data.title}")
            st.markdown(f"**サブタイトル:** {data.subtitle}")
            st.markdown(f"**発行日:** {data.date_str}")
            st.markdown(f"**挨拶:** {data.greeting}")
            st.markdown(f"**本文:** {data.body}")
            st.markdown(f"**来週の予定:** {data.schedule_text}")
            st.markdown(f"**保護者向けメモ:** {data.parent_note}")
# 【追加】プレビューに時間割を表示
            st.markdown("**【来週の時間割】**")
            for day, subs in data.timetable.items():
                st.markdown(f"- **{day}**: " + ", ".join([s if s else "(なし)" for s in subs]))
                
            st.markdown(f"**結び:** {data.closing}")
        
        if data.urls:
            with st.expander("参考URL", expanded=False):
                for url in data.urls:
                    st.markdown(f"- {url}")

        if selected_template:
            try:
                docx_bytes = render_docx(selected_template, data)
                st.download_button(
                    label=f"Wordファイル（{selected_template}）をダウンロード",
                    data=docx_bytes,
                    file_name=f"document_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            except Exception as e:
                st.error(f"テンプレートへの埋め込み中にエラーが発生しました: {e}")
        else:
            st.error("テンプレートフォルダ（templates）にWordファイルがありません。")

if __name__ == "__main__":
    main()
