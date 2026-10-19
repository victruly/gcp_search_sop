#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
GCP Vertex AI Search SOP Presentation Generator
Produces a high-impact, professional widescreen (16:9) PPTX presentation
compatible with Windows PowerPoint and Google Drive / Google Slides.
"""

import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

def build_presentation(output_path="GCP_Vertex_AI_Search_SOP.pptx"):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    # --- Color Palette ---
    C_NAVY_DARK = RGBColor(15, 23, 42)      # #0F172A - Executive Navy
    C_NAVY_LIGHT = RGBColor(30, 41, 59)     # #1E293B - Slate Gray
    C_BLUE = RGBColor(26, 115, 232)         # #1A73E8 - Google Primary Blue
    C_BLUE_LIGHT = RGBColor(232, 240, 254)  # #E8F0FE - Google Soft Blue Tint
    C_BLUE_DARK = RGBColor(21, 87, 176)     # #1557B0 - Deep Blue
    C_CYAN = RGBColor(2, 132, 199)          # #0284C7 - Accent Sky
    C_BG_PAGE = RGBColor(248, 250, 252)     # #F8FAFC - Clean Light Slate
    C_CARD_BG = RGBColor(255, 255, 255)     # #FFFFFF - Pure White
    C_CARD_BORDER = RGBColor(226, 232, 240) # #E2E8F0 - Subtle Gray Border
    C_TEXT_DARK = RGBColor(15, 23, 42)      # #0F172A - Headings
    C_TEXT_BODY = RGBColor(51, 65, 85)      # #334155 - Body text
    C_TEXT_MUTED = RGBColor(100, 116, 139)  # #64748B - Secondary text
    C_GREEN = RGBColor(16, 185, 129)        # #10B981 - Success Emerald
    C_GREEN_LIGHT = RGBColor(236, 253, 245) # #ECFDF5 - Success BG
    C_GREEN_DARK = RGBColor(6, 95, 70)      # #065F46 - Success Text
    C_AMBER = RGBColor(217, 119, 6)         # #D97706 - Amber Warning
    C_AMBER_LIGHT = RGBColor(254, 243, 199) # #FEF3C7 - Amber Warning BG
    C_AMBER_DARK = RGBColor(146, 64, 14)    # #92400E - Amber Warning Text
    C_RED = RGBColor(220, 38, 38)           # #DC2626 - Alert Red
    C_CODE_BG = RGBColor(24, 32, 47)        # #18202F - Terminal Dark
    C_CODE_TEXT = RGBColor(226, 232, 240)   # #E2E8F0 - Terminal Light
    C_CODE_GREEN = RGBColor(74, 222, 128)   # #4ADE80 - Code Keyword/Comment

    FONT_FAMILY = "Microsoft JhengHei"
    FONT_CODE = "Consolas"

    def set_bg(slide, is_dark=False):
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
        bg.fill.solid()
        bg.fill.fore_color.rgb = C_NAVY_DARK if is_dark else C_BG_PAGE
        bg.line.fill.background()
        return bg

    def add_sop_header(slide, step_num, step_tag, title_text, subtitle_text):
        # Step Tag Pill
        tag_box = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(0.38), Inches(2.5), Inches(0.32)
        )
        tag_box.fill.solid()
        tag_box.fill.fore_color.rgb = C_BLUE_LIGHT
        tag_box.line.color.rgb = C_BLUE
        tag_box.line.width = Pt(1)
        tf_tag = tag_box.text_frame
        tf_tag.word_wrap = True
        tf_tag.vertical_anchor = MSO_ANCHOR.MIDDLE
        p_tag = tf_tag.paragraphs[0]
        p_tag.text = f"SOP STEP {step_num:02d} · {step_tag}"
        p_tag.font.size = Pt(10.5)
        p_tag.font.bold = True
        p_tag.font.color.rgb = C_BLUE
        p_tag.font.name = FONT_FAMILY
        p_tag.alignment = PP_ALIGN.CENTER

        # Title
        tb_title = slide.shapes.add_textbox(Inches(0.8), Inches(0.76), Inches(11.7), Inches(0.52))
        tf_title = tb_title.text_frame
        tf_title.word_wrap = True
        tf_title.margin_left = tf_title.margin_top = tf_title.margin_right = tf_title.margin_bottom = 0
        p_title = tf_title.paragraphs[0]
        p_title.text = title_text
        p_title.font.size = Pt(21)
        p_title.font.bold = True
        p_title.font.color.rgb = C_TEXT_DARK
        p_title.font.name = FONT_FAMILY

        # Subtitle / Objective
        tb_sub = slide.shapes.add_textbox(Inches(0.8), Inches(1.32), Inches(11.7), Inches(0.35))
        tf_sub = tb_sub.text_frame
        tf_sub.word_wrap = True
        tf_sub.margin_left = tf_sub.margin_top = tf_sub.margin_right = tf_sub.margin_bottom = 0
        p_sub = tf_sub.paragraphs[0]
        p_sub.text = f"🎯 本步作業目標：{subtitle_text}"
        p_sub.font.size = Pt(12)
        p_sub.font.color.rgb = C_TEXT_MUTED
        p_sub.font.name = FONT_FAMILY

    def add_card_box(slide, left, top, width, height, title, title_color=C_NAVY_DARK, bg_color=C_CARD_BG, border_color=C_CARD_BORDER):
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
        card.fill.solid()
        card.fill.fore_color.rgb = bg_color
        card.line.color.rgb = border_color
        card.line.width = Pt(1)

        if title:
            tb = slide.shapes.add_textbox(left + Inches(0.22), top + Inches(0.14), width - Inches(0.44), Inches(0.38))
            tf = tb.text_frame
            tf.word_wrap = True
            tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
            p = tf.paragraphs[0]
            p.text = title
            p.font.size = Pt(13)
            p.font.bold = True
            p.font.color.rgb = title_color
            p.font.name = FONT_FAMILY
        return card

    def add_footer(slide, current_page, total_pages=13):
        line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(6.85), Inches(11.733), Pt(1))
        line.fill.solid()
        line.fill.fore_color.rgb = C_CARD_BORDER
        line.line.fill.background()

        tb_left = slide.shapes.add_textbox(Inches(0.8), Inches(6.92), Inches(9), Inches(0.3))
        tf_l = tb_left.text_frame
        tf_l.word_wrap = True
        tf_l.margin_left = tf_l.margin_top = tf_l.margin_right = tf_l.margin_bottom = 0
        p_l = tf_l.paragraphs[0]
        p_l.text = "GCP Vertex AI Search 企業建置與整合標準作業程序 (SOP) · 內部機密技術報告"
        p_l.font.size = Pt(9.5)
        p_l.font.color.rgb = C_TEXT_MUTED
        p_l.font.name = FONT_FAMILY

        tb_right = slide.shapes.add_textbox(Inches(10.533), Inches(6.92), Inches(2.0), Inches(0.3))
        tf_r = tb_right.text_frame
        tf_r.word_wrap = True
        tf_r.margin_left = tf_r.margin_top = tf_r.margin_right = tf_r.margin_bottom = 0
        p_r = tf_r.paragraphs[0]
        p_r.text = f"{current_page} / {total_pages}"
        p_r.font.size = Pt(9.5)
        p_r.font.bold = True
        p_r.font.color.rgb = C_TEXT_MUTED
        p_r.font.name = "Calibri"
        p_r.alignment = PP_ALIGN.RIGHT

    def add_checklist_bar(slide, check_text):
        bar = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(6.12), Inches(11.733), Inches(0.56))
        bar.fill.solid()
        bar.fill.fore_color.rgb = C_GREEN_LIGHT
        bar.line.color.rgb = C_GREEN
        bar.line.width = Pt(1)

        tb = slide.shapes.add_textbox(Inches(1.0), Inches(6.20), Inches(11.3), Inches(0.42))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
        p = tf.paragraphs[0]
        p.text = f"✅ SOP 階段檢核點 (Acceptance Criteria)：{check_text}"
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = C_GREEN_DARK
        p.font.name = FONT_FAMILY

    # ==========================================
    # SLIDE 1: COVER SLIDE (Dark Executive Theme)
    # ==========================================
    s1 = prs.slides.add_slide(blank_layout)
    set_bg(s1, is_dark=True)

    # Accent decorative bar
    bar = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.5), Inches(0.18), Inches(3.8))
    bar.fill.solid()
    bar.fill.fore_color.rgb = C_BLUE
    bar.line.fill.background()

    # Tag
    tb = s1.shapes.add_textbox(Inches(1.2), Inches(1.5), Inches(8), Inches(0.4))
    tf = tb.text_frame
    p = tf.paragraphs[0]
    p.text = "ENTERPRISE SEARCH IMPLEMENTATION SOP · 企業級架構與維運指南"
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = C_CYAN
    p.font.name = FONT_FAMILY

    # Main Title
    tb = s1.shapes.add_textbox(Inches(1.2), Inches(2.0), Inches(10.5), Inches(1.5))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "Google Cloud (GCP) Vertex AI Search\n企業智慧搜尋建置與整合標準作業程序"
    p.font.size = Pt(30)
    p.font.bold = True
    p.font.color.rgb = RGBColor(255, 255, 255)
    p.font.name = FONT_FAMILY

    # Subtitle
    tb = s1.shapes.add_textbox(Inches(1.2), Inches(3.8), Inches(10.5), Inches(0.8))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "從零到一完整實作：專案建立、50 網站白名單檢索、服務帳戶私鑰授權與 Python 後端整合"
    p.font.size = Pt(15)
    p.font.color.rgb = RGBColor(203, 213, 225) # Slate 300
    p.font.name = FONT_FAMILY

    # Metadata Card (Bottom)
    meta_box = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.2), Inches(4.9), Inches(10.5), Inches(1.5))
    meta_box.fill.solid()
    meta_box.fill.fore_color.rgb = C_NAVY_LIGHT
    meta_box.line.color.rgb = RGBColor(51, 65, 85)
    meta_box.line.width = Pt(1)

    items = [
        ("呈報對象", "主管 / 專案審查主管", Inches(1.5)),
        ("報告架構", "10 步驟標準 SOP (含檢核點)", Inches(4.0)),
        ("資安等級", "機密等級 (Confidential)", Inches(6.8)),
        ("系統狀態", "已完成驗證 / 生產就緒", Inches(9.2))
    ]
    for label, val, x_pos in items:
        tb = s1.shapes.add_textbox(x_pos, Inches(5.15), Inches(2.3), Inches(1.0))
        tf = tb.text_frame
        tf.word_wrap = True
        p1 = tf.paragraphs[0]
        p1.text = label
        p1.font.size = Pt(11)
        p1.font.color.rgb = RGBColor(148, 163, 184)
        p1.font.name = FONT_FAMILY

        p2 = tf.add_paragraph()
        p2.text = val
        p2.font.size = Pt(13)
        p2.font.bold = True
        p2.font.color.rgb = RGBColor(255, 255, 255)
        p2.font.name = FONT_FAMILY

    # ==========================================
    # SLIDE 2: EXECUTIVE SUMMARY & SYSTEM ARCHITECTURE
    # ==========================================
    s2 = prs.slides.add_slide(blank_layout)
    set_bg(s2)
    add_sop_header(s2, 0, "架構綜覽", "端到端系統架構與業務價值總覽", "理解專案運作藍圖、安全性防護機制與全流程 10 大執行階段")

    # Left Column: Business & Technical Value (Width: 3.6)
    add_card_box(s2, Inches(0.8), Inches(1.75), Inches(3.6), Inches(4.2), "💡 企業建置效益與價值")
    tb = s2.shapes.add_textbox(Inches(1.0), Inches(2.3), Inches(3.2), Inches(3.4))
    tf = tb.text_frame
    tf.word_wrap = True
    val_points = [
        ("零爬蟲維護負擔", "由 Google 全球分散式爬蟲自動爬取與索引 50 個網站，免除自建 Scraper 成本。"),
        ("Google 語意搜尋與 AI", "超越傳統關鍵字比對，具備自然語言理解 (NLU) 與 Generative AI 答案摘要功能。"),
        ("企業級資安隔離", "金鑰存放於後端隔離環境，前端不接觸私鑰，符合企業最高資訊安全審查標準。"),
        ("毫秒級搜尋回應", "實測回應時間落在 200~400ms，大幅提升內部同仁與終端客戶的檢索效率。")
    ]
    for i, (head, desc) in enumerate(val_points):
        p = tf.add_paragraph() if i > 0 else tf.paragraphs[0]
        p.text = f"• {head}："
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = C_TEXT_DARK
        p.font.name = FONT_FAMILY
        
        p2 = tf.add_paragraph()
        p2.text = f"  {desc}"
        p2.font.size = Pt(10)
        p2.font.color.rgb = C_TEXT_BODY
        p2.font.name = FONT_FAMILY

    # Right Top Column: System Architecture Diagram (Width: 7.9)
    add_card_box(s2, Inches(4.6), Inches(1.75), Inches(7.933), Inches(2.2), "🏛️ 系統端到端架構流程圖")
    arch_box = s2.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(4.8), Inches(2.25), Inches(7.533), Inches(1.5))
    arch_box.fill.solid()
    arch_box.fill.fore_color.rgb = C_NAVY_DARK
    arch_box.line.color.rgb = C_NAVY_LIGHT

    tb_code = s2.shapes.add_textbox(Inches(4.9), Inches(2.32), Inches(7.333), Inches(1.35))
    tf_c = tb_code.text_frame
    tf_c.word_wrap = True
    p = tf_c.paragraphs[0]
    p.text = (
        "[50 個目標產業網站] ──(Googlebot 索引)──> [GCP Vertex AI Search (Data Store)]\n"
        "                                                      │\n"
        "[前端瀏覽器 / UI] ──(POST /api/search)──> [Python 後端 Flask API]\n"
        "                                      (使用私鑰 OAuth2 安全認證) │\n"
        "                                                      ▼\n"
        "                                      [GCP Discovery Engine API]"
    )
    p.font.size = Pt(9.5)
    p.font.color.rgb = C_CODE_GREEN
    p.font.name = FONT_CODE

    # Right Bottom: 10 Steps Roadmap (Width: 7.9)
    add_card_box(s2, Inches(4.6), Inches(4.1), Inches(7.933), Inches(1.85), "📋 SOP 10 大標準實施步驟階段")
    tb_steps = s2.shapes.add_textbox(Inches(4.8), Inches(4.55), Inches(7.533), Inches(1.3))
    tf_s = tb_steps.text_frame
    tf_s.word_wrap = True
    steps_roadmap = (
        "Phase 1 (雲端環境置備) : STEP 01 專案建立 ➔ STEP 02 啟用 API ➔ STEP 03 建立 App\n"
        "Phase 2 (檢索資料集建立): STEP 04 建立 Data Store ➔ STEP 05 批量匯入 50 個網站白名單\n"
        "Phase 3 (資安與金鑰授權): STEP 06 IAM 角色授權 ➔ STEP 07 生成並安全下載私鑰 JSON\n"
        "Phase 4 (後端串接與驗證): STEP 08 Python 環境 ➔ STEP 09 SDK 撰寫 ➔ STEP 10 前後端測試"
    )
    p = tf_s.paragraphs[0]
    p.text = steps_roadmap
    p.font.size = Pt(10.5)
    p.font.color.rgb = C_BLUE_DARK
    p.font.name = FONT_FAMILY
    p.font.bold = True

    add_checklist_bar(s2, "確認架構符合專案規範，全流程共 10 步驟均通過技術可行性評估。")
    add_footer(s2, 2)

    # ==========================================
    # SLIDE 3: STEP 1 - CREATE PROJECT & BILLING
    # ==========================================
    s3 = prs.slides.add_slide(blank_layout)
    set_bg(s3)
    add_sop_header(s3, 1, "雲端環境", "申請 GCP 帳號與建立新專案 (Project)", "建立專案命名空間、關聯計費帳戶並設定安全預算警示")

    # Left: Step-by-step UI actions (Width: 6.8)
    add_card_box(s3, Inches(0.8), Inches(1.75), Inches(6.8), Inches(4.2), "📝 介面操作步驟 (Console Operations)")
    tb = s3.shapes.add_textbox(Inches(1.0), Inches(2.25), Inches(6.4), Inches(3.5))
    tf = tb.text_frame
    tf.word_wrap = True
    s1_steps = [
        ("1. 登入 Google Cloud 控制台", "瀏覽器前往 https://console.cloud.google.com/，使用企業管理員或技術負責人帳號登入。"),
        ("2. 建立新專案 (Create Project)", "點擊頂端列的「專案選單下拉鈕」➔ 點選視窗右上角「新增專案 (New Project)」。"),
        ("3. 設定專案識別名稱", "輸入專案名稱 (如 `enterprise-search-service`)，系統會自動生成全域唯一 Project ID (如 `gen-lang-client-xxxxxxx`)。"),
        ("4. 關聯計費帳戶 (Billing Account)", "導航至「帳單 (Billing)」專區，將專案連結至公司企業合約或啟用 $300 美金免費試用帳戶。"),
        ("5. 設定預算警示 (Budget Alert)", "在帳單後台建立預算預警 (例如達 50%、80%、100% 時寄發 Email 警報通知主管)。")
    ]
    for i, (st, desc) in enumerate(s1_steps):
        p = tf.add_paragraph() if i > 0 else tf.paragraphs[0]
        p.text = f"{st}"
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = C_TEXT_DARK
        p.font.name = FONT_FAMILY

        p2 = tf.add_paragraph()
        p2.text = f"   {desc}"
        p2.font.size = Pt(10)
        p2.font.color.rgb = C_TEXT_BODY
        p2.font.name = FONT_FAMILY

    # Right Top: Key Parameters (Width: 4.7)
    add_card_box(s3, Inches(7.8), Inches(1.75), Inches(4.733), Inches(2.0), "🔑 關鍵配置參數 (Parameters)")
    tb = s3.shapes.add_textbox(Inches(8.0), Inches(2.25), Inches(4.3), Inches(1.3))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = (
        "• 專案名稱 (Project Name)：gcp-agent-search\n"
        "• 專案 ID (Project ID)：gen-lang-client-0687915907\n"
        "• 預設計費帳戶：Enterprise Billing Account\n"
        "• 資源層級：Organization / Root Folder"
    )
    p.font.size = Pt(10.5)
    p.font.color.rgb = C_TEXT_BODY
    p.font.name = FONT_FAMILY

    # Right Bottom: Warnings / Gotchas (Width: 4.7)
    add_card_box(s3, Inches(7.8), Inches(3.9), Inches(4.733), Inches(2.05), "⚠️ 主管與資安審核重點", C_AMBER, C_AMBER_LIGHT, C_AMBER)
    tb = s3.shapes.add_textbox(Inches(8.0), Inches(4.35), Inches(4.3), Inches(1.5))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = (
        "1. 計費關聯是必要前置條件：若未綁定計費帳戶，後續 Discovery Engine API 將無法被啟用。\n"
        "2. Project ID 一旦建立不可更改：所有後端環境變數、IAM 授權與 API 路徑均嚴格綁定此 ID。\n"
        "3. 成本防護措施：預算警示能保障雲端花費透明可控。"
    )
    p.font.size = Pt(10)
    p.font.color.rgb = C_AMBER_DARK
    p.font.name = FONT_FAMILY

    add_checklist_bar(s3, "取得有效且已啟用計費的專案 ID (Project ID)，能在控制台正常切換專案。")
    add_footer(s3, 3)

    # ==========================================
    # SLIDE 4: STEP 2 - ENABLE DISCOVERY ENGINE API
    # ==========================================
    s4 = prs.slides.add_slide(blank_layout)
    set_bg(s4)
    add_sop_header(s4, 2, "雲端環境", "啟用 Vertex AI Search (Agent Builder) 與 API", "啟動 Google 核心語意搜尋服務引擎 (Discovery Engine API)")

    # Left: Step-by-step UI actions (Width: 6.8)
    add_card_box(s4, Inches(0.8), Inches(1.75), Inches(6.8), Inches(4.2), "📝 介面操作步驟 (Console Operations)")
    tb = s4.shapes.add_textbox(Inches(1.0), Inches(2.25), Inches(6.4), Inches(3.5))
    tf = tb.text_frame
    tf.word_wrap = True
    s2_steps = [
        ("1. 導航至 Agent Builder 服務", "在 Google Cloud Console 頂端搜尋欄輸入「Agent Builder」或「Vertex AI Search」並按下 Enter。"),
        ("2. 首次服務啟動引導", "若是專案首次進入，頁面會出現服務條款與歡迎畫面，點選「繼續並啟用 API (Continue and activate the API)」。"),
        ("3. 等候 API 自動佈署", "GCP 系統會在背景自動配置 `discoveryengine.googleapis.com` 資源，約需等待 1 至 2 分鐘。"),
        ("4. 檢查 API 服務啟用狀態", "前往「API 和服務」➔「已啟用的 API 和服務」，確認清單中包含 Discovery Engine API 且狀態為綠色正常。")
    ]
    for i, (st, desc) in enumerate(s2_steps):
        p = tf.add_paragraph() if i > 0 else tf.paragraphs[0]
        p.text = f"{st}"
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = C_TEXT_DARK
        p.font.name = FONT_FAMILY

        p2 = tf.add_paragraph()
        p2.text = f"   {desc}"
        p2.font.size = Pt(10)
        p2.font.color.rgb = C_TEXT_BODY
        p2.font.name = FONT_FAMILY

    # Right Top: API Specifications
    add_card_box(s4, Inches(7.8), Inches(1.75), Inches(4.733), Inches(2.0), "🔑 API 服務規格與端點")
    tb = s4.shapes.add_textbox(Inches(8.0), Inches(2.25), Inches(4.3), Inches(1.3))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = (
        "• API 名稱：Discovery Engine API\n"
        "• 服務識別碼：discoveryengine.googleapis.com\n"
        "• 預設呼叫區域 (Location)：global (全球分散式)\n"
        "• 包含功能：語意檢索、AI 答案摘要、網站爬取管理"
    )
    p.font.size = Pt(10.5)
    p.font.color.rgb = C_TEXT_BODY
    p.font.name = FONT_FAMILY

    # Right Bottom: Warnings
    add_card_box(s4, Inches(7.8), Inches(3.9), Inches(4.733), Inches(2.05), "⚠️ 常見阻礙與除錯建議", C_AMBER, C_AMBER_LIGHT, C_AMBER)
    tb = s4.shapes.add_textbox(Inches(8.0), Inches(4.35), Inches(4.3), Inches(1.5))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = (
        "1. 權限不足問題：操作人員需具備專案 Owner 或 Service Usage Admin 角色，方可啟用雲端 API。\n"
        "2. 佈署延遲處理：若啟用後立即呼叫 API 出現 404/403，請靜候 2~3 分鐘待全域 IAM 與 DNS 規則生效。\n"
        "3. 區域選擇提示：企業資料若無特定主權限制，建議一律維持 global。"
    )
    p.font.size = Pt(10)
    p.font.color.rgb = C_AMBER_DARK
    p.font.name = FONT_FAMILY

    add_checklist_bar(s4, "Discovery Engine API 顯示為啟用中，Console 能順暢進入 Agent Builder 儀表板。")
    add_footer(s4, 4)

    # ==========================================
    # SLIDE 5: STEP 3 - CREATE SEARCH APP / ENGINE
    # ==========================================
    s5 = prs.slides.add_slide(blank_layout)
    set_bg(s5)
    add_sop_header(s5, 3, "搜尋實體", "建立 Search 應用程式 (App / Engine)", "配置搜尋引擎實體、自訂功能特性並生成專屬 Engine ID")

    # Left: Step-by-step UI actions
    add_card_box(s5, Inches(0.8), Inches(1.75), Inches(6.8), Inches(4.2), "📝 介面操作步驟 (Console Operations)")
    tb = s5.shapes.add_textbox(Inches(1.0), Inches(2.25), Inches(6.4), Inches(3.5))
    tf = tb.text_frame
    tf.word_wrap = True
    s3_steps = [
        ("1. 進入 Agent Builder Apps 管理區", "於 Agent Builder 儀表板，在左側主選單點選「應用程式 (Apps)」進入管理列表。"),
        ("2. 點擊「建立應用程式 (Create App)」", "點選頂端的「＋ 建立應用程式」按鈕開始設定。"),
        ("3. 選擇應用程式類型：搜尋 (Search)", "在四大類型 (Search / Chat / Recommendation / Agent) 中，選擇「搜尋 (Search)」類型。"),
        ("4. 啟用進階與 AI 生成功能", "勾選「企業版功能 (Enterprise Edition)」與「Generative Answers (AI 答案總結)」選項。"),
        ("5. 命名並產生 Engine ID", "輸入應用程式名稱 (如 `pbsearch`)，系統將生成專屬 Engine ID (如 `pbsearch_1791792965002`)。")
    ]
    for i, (st, desc) in enumerate(s3_steps):
        p = tf.add_paragraph() if i > 0 else tf.paragraphs[0]
        p.text = f"{st}"
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = C_TEXT_DARK
        p.font.name = FONT_FAMILY

        p2 = tf.add_paragraph()
        p2.text = f"   {desc}"
        p2.font.size = Pt(10)
        p2.font.color.rgb = C_TEXT_BODY
        p2.font.name = FONT_FAMILY

    # Right Top: Key Parameters
    add_card_box(s5, Inches(7.8), Inches(1.75), Inches(4.733), Inches(2.0), "🔑 應用程式配置參數")
    tb = s5.shapes.add_textbox(Inches(8.0), Inches(2.25), Inches(4.3), Inches(1.3))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = (
        "• 應用程式類型：Search (Generic Search)\n"
        "• 應用程式名稱：pbsearch\n"
        "• 引擎識別碼 (Engine ID)：pbsearch_1791792965002\n"
        "• AI 摘要功能 (Generative AI)：已啟用 (含來源標註)"
    )
    p.font.size = Pt(10.5)
    p.font.color.rgb = C_TEXT_BODY
    p.font.name = FONT_FAMILY

    # Right Bottom: Warnings
    add_card_box(s5, Inches(7.8), Inches(3.9), Inches(4.733), Inches(2.05), "⚠️ 關鍵架構決策說明")
    tb = s5.shapes.add_textbox(Inches(8.0), Inches(4.35), Inches(4.3), Inches(1.5))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = (
        "1. 引擎 (Engine) 與資料庫 (Data Store) 解耦：一個 Engine 可掛載多個 Data Store，具備高度彈性。\n"
        "2. Engine ID 為後端 API 必填路徑：後端 Python 程式碼發送 SearchRequest 時，必須帶入此 Engine ID。\n"
        "3. AI 摘要計費：啟用 Generative Summary 會帶來極佳的用戶體驗，且計費公開透明。"
    )
    p.font.size = Pt(10)
    p.font.color.rgb = C_TEXT_BODY
    p.font.name = FONT_FAMILY

    add_checklist_bar(s5, "成功建立 Search App，記錄專屬 Engine ID 供後續 Python 後端設定使用。")
    add_footer(s5, 5)

    # ==========================================
    # SLIDE 6: STEP 4 - CREATE WEBSITE DATA STORE
    # ==========================================
    s6 = prs.slides.add_slide(blank_layout)
    set_bg(s6)
    add_sop_header(s6, 4, "資料來源", "建立資料存放區 (Data Store) 並配置網站檢索", "建立資料儲存庫，核心選用「進階網站搜尋 (Advanced)」免驗證外部網域")

    # Left: Step-by-step UI actions
    add_card_box(s6, Inches(0.8), Inches(1.75), Inches(6.8), Inches(4.2), "📝 介面操作步驟 (Console Operations)")
    tb = s6.shapes.add_textbox(Inches(1.0), Inches(2.25), Inches(6.4), Inches(3.5))
    tf = tb.text_frame
    tf.word_wrap = True
    s4_steps = [
        ("1. 進入 Data Stores 建立流程", "在建立 App 的「資料」步驟或前往「資料存放區 (Data Stores)」頁面，點選「建立資料存放區」。"),
        ("2. 選擇資料來源：網站 (Websites)", "資料來源類型中選取「網站 (Websites)」以支援公開網際網路頁面檢索。"),
        ("3. ★關鍵步驟：選取進階網站搜尋 (Advanced)", "在網站搜尋模式中，務必選取「進階網站搜尋 (Advanced Website Search)」，切勿選基礎模式。"),
        ("4. 命名資料存放區並完成綁定", "輸入資料庫名稱 (如 `pbdatastore`)，系統生成 Data Store ID (如 `pbdatastore_1791793664793`)，並將其連接至 Search App。")
    ]
    for i, (st, desc) in enumerate(s4_steps):
        p = tf.add_paragraph() if i > 0 else tf.paragraphs[0]
        p.text = f"{st}"
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = C_TEXT_DARK
        p.font.name = FONT_FAMILY

        p2 = tf.add_paragraph()
        p2.text = f"   {desc}"
        p2.font.size = Pt(10)
        p2.font.color.rgb = C_TEXT_BODY
        p2.font.name = FONT_FAMILY

    # Right Top: Advanced vs Basic Comparison
    add_card_box(s6, Inches(7.8), Inches(1.75), Inches(4.733), Inches(2.1), "⭐ 核心決策：基礎 vs 進階網站搜尋", C_BLUE_DARK, C_BLUE_LIGHT, C_BLUE)
    tb = s6.shapes.add_textbox(Inches(8.0), Inches(2.2), Inches(4.3), Inches(1.5))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = (
        "• 基礎模式 (Basic)：要求 Google Search Console 驗證網域擁有權。僅能爬取公司自家持有的單一網域！\n"
        "• 進階模式 (Advanced)：免網域驗證！可自由爬取第三方、產業協會、技術論壇等 50 個外部公開網站。\n"
        "➔ 結論：要涵蓋 50 個多元網站，必須選擇 Advanced！"
    )
    p.font.size = Pt(10)
    p.font.color.rgb = C_BLUE_DARK
    p.font.name = FONT_FAMILY

    # Right Bottom: Key Parameters
    add_card_box(s6, Inches(7.8), Inches(4.0), Inches(4.733), Inches(1.95), "🔑 資料庫配置參數")
    tb = s6.shapes.add_textbox(Inches(8.0), Inches(4.45), Inches(4.3), Inches(1.4))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = (
        "• 資料來源：Website (公開網頁)\n"
        "• 檢索模式：Advanced Website Indexing\n"
        "• 資料存放區 ID (Data Store ID)：pbdatastore_1791793664793\n"
        "• 綁定狀態：已掛載於 pbsearch 引擎"
    )
    p.font.size = Pt(10.5)
    p.font.color.rgb = C_TEXT_BODY
    p.font.name = FONT_FAMILY

    add_checklist_bar(s6, "Data Store 成功建立為 Advanced Website 模式，並與 Search App 完成關聯。")
    add_footer(s6, 6)

    # ==========================================
    # SLIDE 7: STEP 5 - BATCH ADD 50 WEBSITES
    # ==========================================
    s7 = prs.slides.add_slide(blank_layout)
    set_bg(s7)
    add_sop_header(s7, 5, "資料庫建置", "批量新增 50 個目標網站至搜尋白名單", "設定 URL 規則 (Included Patterns)、配置爬蟲路徑並啟動全網索引")

    # Left: Step-by-step UI actions
    add_card_box(s7, Inches(0.8), Inches(1.75), Inches(6.8), Inches(4.2), "📝 介面操作步驟 (Console Operations)")
    tb = s7.shapes.add_textbox(Inches(1.0), Inches(2.25), Inches(6.4), Inches(3.5))
    tf = tb.text_frame
    tf.word_wrap = True
    s5_steps = [
        ("1. 進入 Data Store 網址管理介面", "點選剛建立的 Data Store (`pbdatastore`) ➔ 切換至「資料 (Data)」或「管理」頁籤。"),
        ("2. 點擊「新增比對模式 (Add URL patterns)」", "在 Included Patterns 區域點選新增，系統支援批次輸入多組網址模式。"),
        ("3. 依據正規表達式規則輸入 50 個網址", "每行輸入一個目標網域，使用萬用字元 `*` 涵蓋全站內容：\n  例：`https://www.semicontaiwan.org/*`\n  例：`https://example-tech.com/docs/*`"),
        ("4. 也可透過 Cloud Storage 批量匯入", "若網址龐大，可將 50 個網址儲存為 CSV/TXT 上傳至 Cloud Storage Bucket 一鍵批次匯入。"),
        ("5. 儲存並等待 Googlebot 索引", "點選儲存後，Google 爬蟲將依序排程檢索網頁並建立語意向量庫。")
    ]
    for i, (st, desc) in enumerate(s5_steps):
        p = tf.add_paragraph() if i > 0 else tf.paragraphs[0]
        p.text = f"{st}"
        p.font.size = Pt(10.5)
        p.font.bold = True
        p.font.color.rgb = C_TEXT_DARK
        p.font.name = FONT_FAMILY

        p2 = tf.add_paragraph()
        p2.text = f"   {desc}"
        p2.font.size = Pt(9.5)
        p2.font.color.rgb = C_TEXT_BODY
        p2.font.name = FONT_FAMILY

    # Right Top: Pattern Syntax Rules
    add_card_box(s7, Inches(7.8), Inches(1.75), Inches(4.733), Inches(2.1), "🌐 URL Pattern 比對語法規格")
    tb = s7.shapes.add_textbox(Inches(8.0), Inches(2.2), Inches(4.3), Inches(1.5))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = (
        "• 全站比對：https://domain.com/* (包含所有子路徑)\n"
        "• 精準子目錄：https://domain.com/news/* (僅爬取新聞)\n"
        "• 排除規則 (Excluded)：若特定網站包含登入頁或無效頁面，可於 Excluded Patterns 加入 `*/login/*`。"
    )
    p.font.size = Pt(10)
    p.font.color.rgb = C_TEXT_BODY
    p.font.name = FONT_FAMILY

    # Right Bottom: Crawling Timeline
    add_card_box(s7, Inches(7.8), Inches(4.0), Inches(4.733), Inches(1.95), "⏱️ 爬蟲排程與時間預估說明", C_AMBER, C_AMBER_LIGHT, C_AMBER)
    tb = s7.shapes.add_textbox(Inches(8.0), Inches(4.45), Inches(4.3), Inches(1.4))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = (
        "1. 爬蟲時間差：新增 50 個網站後，Google 需 2~24 小時完成初始爬取與索引建立。\n"
        "2. robots.txt 限制：若目標網站宣告禁止爬蟲，將無法收錄。\n"
        "3. 更新頻率：Google 會自動定期回訪更新最新發布頁面。"
    )
    p.font.size = Pt(9.5)
    p.font.color.rgb = C_AMBER_DARK
    p.font.name = FONT_FAMILY

    add_checklist_bar(s7, "50 個目標網站正規化 URL Pattern 已完整納入白名單，爬取狀態顯示運作中。")
    add_footer(s7, 7)

    # ==========================================
    # SLIDE 8: STEP 6 - CREATE SERVICE ACCOUNT & ROLES
    # ==========================================
    s8 = prs.slides.add_slide(blank_layout)
    set_bg(s8)
    add_sop_header(s8, 6, "身分與權限", "建立 IAM 服務帳戶並授權 (Service Account)", "建立程式專用連線身分，貫徹最小權限資安原則 (Least Privilege)")

    # Left: Step-by-step UI actions
    add_card_box(s8, Inches(0.8), Inches(1.75), Inches(6.8), Inches(4.2), "📝 介面操作步驟 (Console Operations)")
    tb = s8.shapes.add_textbox(Inches(1.0), Inches(2.25), Inches(6.4), Inches(3.5))
    tf = tb.text_frame
    tf.word_wrap = True
    s6_steps = [
        ("1. 導航至 IAM 服務帳戶管理區", "於主選單前往「IAM 與管理 (IAM & Admin)」➔ 點選「服務帳戶 (Service Accounts)」。"),
        ("2. 點擊「建立服務帳戶」", "點擊頂端「＋ 建立服務帳戶 (+ CREATE SERVICE ACCOUNT)」開啟配置精靈。"),
        ("3. 設定服務帳戶識別名稱", "輸入服務帳戶名稱 (如 `gcp-search-backend-sa`)，系統自動生成對應 Email 地址。"),
        ("4. 指派關鍵 IAM 權限角色 (Grant Access)", "在角色下拉選單中搜尋並選取：\n  • 核心角色：`Discovery Engine Viewer` (`roles/discoveryengine.viewer`)\n  (若後端需動態管理資料庫，可選 `Discovery Engine Admin`)"),
        ("5. 完成帳戶建立", "點選「完成 (Done)」，服務帳戶即刻具備呼叫 GCP 搜尋 API 的合法身分。")
    ]
    for i, (st, desc) in enumerate(s6_steps):
        p = tf.add_paragraph() if i > 0 else tf.paragraphs[0]
        p.text = f"{st}"
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = C_TEXT_DARK
        p.font.name = FONT_FAMILY

        p2 = tf.add_paragraph()
        p2.text = f"   {desc}"
        p2.font.size = Pt(10)
        p2.font.color.rgb = C_TEXT_BODY
        p2.font.name = FONT_FAMILY

    # Right Top: Role Details
    add_card_box(s8, Inches(7.8), Inches(1.75), Inches(4.733), Inches(2.0), "🛡️ 核心 IAM 角色權限對照表")
    tb = s8.shapes.add_textbox(Inches(8.0), Inches(2.25), Inches(4.3), Inches(1.3))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = (
        "• Discovery Engine Viewer (推薦)：\n"
        "  - 權限：唯讀檢索 (search) 與取得系統資訊。\n"
        "  - 適用：後端生產環境 API 伺服器。\n"
        "• Discovery Engine Admin：\n"
        "  - 權限：完整讀寫、新增刪除 Data Store / Engine。\n"
        "  - 適用：具備動態維運功能的管理後端。"
    )
    p.font.size = Pt(10)
    p.font.color.rgb = C_TEXT_BODY
    p.font.name = FONT_FAMILY

    # Right Bottom: Security Best Practice
    add_card_box(s8, Inches(7.8), Inches(3.9), Inches(4.733), Inches(2.05), "🔒 長官資安把關重點", C_GREEN_DARK, C_GREEN_LIGHT, C_GREEN)
    tb = s8.shapes.add_textbox(Inches(8.0), Inches(4.35), Inches(4.3), Inches(1.5))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = (
        "1. 嚴禁授予 Project Owner/Editor 角色：避免帳號遭濫用時波及全雲端專案。\n"
        "2. 角色權限精確鎖定：僅授予 Discovery Engine 專用角色，落實權限最小化原則。\n"
        "3. 稽核日誌 (Cloud Audit Logs)：所有 API 呼叫均會被完整記錄於 Cloud Logging。"
    )
    p.font.size = Pt(9.8)
    p.font.color.rgb = C_GREEN_DARK
    p.font.name = FONT_FAMILY

    add_checklist_bar(s8, "服務帳戶建立完成，並已成功繫結 Discovery Engine Viewer 角色。")
    add_footer(s8, 8)

    # ==========================================
    # SLIDE 9: STEP 7 - GENERATE & DOWNLOAD PRIVATE KEY JSON
    # ==========================================
    s9 = prs.slides.add_slide(blank_layout)
    set_bg(s9)
    add_sop_header(s9, 7, "私鑰管理", "生成並下載私鑰金鑰 (Private Key JSON)", "取得 RSA 憑證私鑰，落實本機安全存放與資安規範防護措施")

    # Left: Step-by-step UI actions
    add_card_box(s9, Inches(0.8), Inches(1.75), Inches(6.8), Inches(4.2), "📝 介面操作步驟 (Console Operations)")
    tb = s9.shapes.add_textbox(Inches(1.0), Inches(2.25), Inches(6.4), Inches(3.5))
    tf = tb.text_frame
    tf.word_wrap = True
    s7_steps = [
        ("1. 點選進入目標服務帳戶詳情", "在服務帳戶清單中，點選剛建立的帳號 Email 連結進入設定頁。"),
        ("2. 切換至「金鑰 (Keys)」頁籤", "在上方次選單列中，點選「金鑰 (Keys)」頁籤。"),
        ("3. 建立新金鑰 (Add Key)", "點擊「新增金鑰 (ADD KEY)」下拉選單 ➔ 選擇「建立新的金鑰 (Create new key)」。"),
        ("4. 選擇金鑰類型為 JSON (建議)", "在彈出的對話框中，勾選「JSON (建議)」選項 ➔ 點擊「建立 (CREATE)」。"),
        ("5. 安全下載與檔案保存", "瀏覽器將自動下載私鑰 JSON 檔案 (如 `gen-lang-client-xxxx.json`)，將其儲存於後端安全目錄。")
    ]
    for i, (st, desc) in enumerate(s7_steps):
        p = tf.add_paragraph() if i > 0 else tf.paragraphs[0]
        p.text = f"{st}"
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = C_TEXT_DARK
        p.font.name = FONT_FAMILY

        p2 = tf.add_paragraph()
        p2.text = f"   {desc}"
        p2.font.size = Pt(10)
        p2.font.color.rgb = C_TEXT_BODY
        p2.font.name = FONT_FAMILY

    # Right Top: JSON Content Structure
    add_card_box(s9, Inches(7.8), Inches(1.75), Inches(4.733), Inches(2.0), "📄 私鑰 JSON 結構說明")
    tb = s9.shapes.add_textbox(Inches(8.0), Inches(2.2), Inches(4.3), Inches(1.4))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = (
        "• type: service_account\n"
        "• project_id: gen-lang-client-0687915907\n"
        "• private_key_id: 5ba27441b36c...\n"
        "• private_key: -----BEGIN RSA PRIVATE KEY-----\n"
        "• client_email: backend-sa@...iam.gserviceaccount.com"
    )
    p.font.size = Pt(9.5)
    p.font.color.rgb = C_CODE_BG
    p.font.name = FONT_CODE

    # Right Bottom: Strict Security Rules
    add_card_box(s9, Inches(7.8), Inches(3.9), Inches(4.733), Inches(2.05), "🚨 最高資安紅線規範 (Zero Trust)", C_RED, RGBColor(254, 242, 242), C_RED)
    tb = s9.shapes.add_textbox(Inches(8.0), Inches(4.35), Inches(4.3), Inches(1.5))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = (
        "1. 嚴禁放至前端：此私鑰絕對不可放在 HTML、JS 或客戶端，否則金鑰即刻失竊！\n"
        "2. 嚴禁提交 Git：必須將 `*.json` 寫入 `.gitignore`，防止推送到 GitHub/GitLab。\n"
        "3. 生產環境建議：可改用 Google Cloud Secret Manager 或 Workload Identity 免金鑰機制。"
    )
    p.font.size = Pt(9.8)
    p.font.color.rgb = RGBColor(153, 27, 27)
    p.font.name = FONT_FAMILY

    add_checklist_bar(s9, "成功下載 JSON 私鑰，檔案已置於後端受保護目錄，並已設置 .gitignore 防護。")
    add_footer(s9, 9)

    # ==========================================
    # SLIDE 10: STEP 8 - PYTHON BACKEND ENVIRONMENT SETUP
    # ==========================================
    s10 = prs.slides.add_slide(blank_layout)
    set_bg(s10)
    add_sop_header(s10, 8, "後端環境", "Python 開發環境配置與官方 SDK 安裝", "建置獨立虛擬環境、安裝官方相依套件並配置環境變數")

    # Left: Step-by-step commands (Width: 6.8)
    add_card_box(s10, Inches(0.8), Inches(1.75), Inches(6.8), Inches(4.2), "💻 終端機指令與套件安裝 (Terminal Workflow)")
    
    # Terminal Black Box
    term = s10.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1.0), Inches(2.25), Inches(6.4), Inches(3.45))
    term.fill.solid()
    term.fill.fore_color.rgb = C_CODE_BG
    term.line.color.rgb = C_NAVY_LIGHT

    tb_t = s10.shapes.add_textbox(Inches(1.15), Inches(2.35), Inches(6.1), Inches(3.25))
    tf_t = tb_t.text_frame
    tf_t.word_wrap = True
    code_text = (
        "# 1. 建立並啟動 Python 虛擬環境 (Python 3.9+)\n"
        "python3 -m venv venv\n"
        "source venv/bin/activate  # Windows: venv\\Scripts\\activate\n\n"
        "# 2. 安裝 GCP 官方專用 Discovery Engine SDK\n"
        "pip install google-cloud-discoveryengine>=0.11.0\n"
        "pip install google-auth>=2.0.0\n\n"
        "# 3. 安裝後端 API 與輔助套件\n"
        "pip install flask flask-cors python-dotenv\n\n"
        "# 4. 產出相依性鎖定清單\n"
        "pip freeze > requirements.txt"
    )
    p = tf_t.paragraphs[0]
    p.text = code_text
    p.font.size = Pt(9.5)
    p.font.color.rgb = C_CODE_TEXT
    p.font.name = FONT_CODE

    # Right Top: .env Configuration File
    add_card_box(s10, Inches(7.8), Inches(1.75), Inches(4.733), Inches(2.2), "⚙️ 環境變數檔配置 (`.env`)")
    env_box = s10.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(8.0), Inches(2.2), Inches(4.333), Inches(1.6))
    env_box.fill.solid()
    env_box.fill.fore_color.rgb = C_NAVY_DARK
    env_box.line.color.rgb = C_NAVY_LIGHT

    tb_e = s10.shapes.add_textbox(Inches(8.1), Inches(2.25), Inches(4.1), Inches(1.5))
    tf_e = tb_e.text_frame
    tf_e.word_wrap = True
    env_content = (
        "GCP_KEY_FILE=gen-lang-client-xxxx.json\n"
        "GCP_PROJECT_ID=gen-lang-client-0687915907\n"
        "GCP_LOCATION=global\n"
        "GCP_ENGINE_ID=pbsearch_1791792965002\n"
        "GCP_DATA_STORE_ID=pbdatastore_1791793664793\n"
        "PORT=5000"
    )
    p = tf_e.paragraphs[0]
    p.text = env_content
    p.font.size = Pt(9)
    p.font.color.rgb = C_CODE_GREEN
    p.font.name = FONT_CODE

    # Right Bottom: Best Practice
    add_card_box(s10, Inches(7.8), Inches(4.1), Inches(4.733), Inches(1.85), "💡 程式碼解耦原則 (Decoupling)")
    tb = s10.shapes.add_textbox(Inches(8.0), Inches(4.55), Inches(4.3), Inches(1.3))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = (
        "• 嚴禁將 Project ID 或金鑰檔名硬編碼 (Hardcode) 於程式中。\n"
        "• 透過 `python-dotenv` 讀取 `.env`，使開發、測試與正式環境切換更加安全迅速。"
    )
    p.font.size = Pt(10)
    p.font.color.rgb = C_TEXT_BODY
    p.font.name = FONT_FAMILY

    add_checklist_bar(s10, "虛擬環境建立完畢，SDK 套件驗證正常，.env 參數載入就緒。")
    add_footer(s10, 10)

    # ==========================================
    # SLIDE 11: STEP 9 - PYTHON BACKEND CODE IMPLEMENTATION
    # ==========================================
    s11 = prs.slides.add_slide(blank_layout)
    set_bg(s11)
    add_sop_header(s11, 9, "程式實作", "Python 後端程式撰寫：憑證載入與呼叫 Search", "撰寫標準 SDK 程式邏輯，利用私鑰完成 OAuth2 驗證並發起檢索請求")

    # Left: Code Snippet (Width: 7.2)
    add_card_box(s11, Inches(0.8), Inches(1.75), Inches(7.2), Inches(4.2), "💻 後端核心程式碼 (`app.py`)")
    
    code_bg = s11.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1.0), Inches(2.25), Inches(6.8), Inches(3.45))
    code_bg.fill.solid()
    code_bg.fill.fore_color.rgb = C_CODE_BG
    code_bg.line.color.rgb = C_NAVY_LIGHT

    tb_c = s11.shapes.add_textbox(Inches(1.1), Inches(2.32), Inches(6.6), Inches(3.3))
    tf_c = tb_c.text_frame
    tf_c.word_wrap = True
    app_code = (
        "from google.oauth2 import service_account\n"
        "from google.cloud import discoveryengine_v1 as discoveryengine\n\n"
        "# 1. 載入私鑰憑證 (OAuth2 Scopes)\n"
        "credentials = service_account.Credentials.from_service_account_file(\n"
        "    'gen-lang-client-xxxx.json',\n"
        "    scopes=['https://www.googleapis.com/auth/cloud-platform']\n"
        ")\n"
        "# 2. 初始化搜尋客戶端 (Search Client)\n"
        "client = discoveryengine.SearchServiceClient(credentials=credentials)\n\n"
        "# 3. 建構搜尋請求 (指定 Serving Config 與查詢關鍵字)\n"
        "serving_path = f'projects/{PROJECT_ID}/locations/{LOCATION}/collections/default_collection/engines/{ENGINE_ID}/servingConfigs/default_search'\n"
        "request = discoveryengine.SearchRequest(\n"
        "    serving_config=serving_path,\n"
        "    query='半導體封裝測試',\n"
        "    page_size=10\n"
        ")\n"
        "# 4. 呼叫 GCP API 並取得回傳結果\n"
        "response = client.search(request=request)"
    )
    p = tf_c.paragraphs[0]
    p.text = app_code
    p.font.size = Pt(8.5)
    p.font.color.rgb = C_CODE_TEXT
    p.font.name = FONT_CODE

    # Right Top: Code Logic Breakdown
    add_card_box(s11, Inches(8.2), Inches(1.75), Inches(4.333), Inches(2.1), "🔍 核心程式架構解析")
    tb = s11.shapes.add_textbox(Inches(8.4), Inches(2.2), Inches(3.9), Inches(1.5))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = (
        "1. 憑證載入：利用 `service_account.Credentials` 自動處理短期 JWT 簽署與權限授權。\n"
        "2. Serving Config：指向預設搜尋規則路徑 (`default_search`)。\n"
        "3. 物件導向呼叫：Google SDK 自動維護連線池 (Connection Pool) 與重試機制。"
    )
    p.font.size = Pt(9.5)
    p.font.color.rgb = C_TEXT_BODY
    p.font.name = FONT_FAMILY

    # Right Bottom: Response Parsing
    add_card_box(s11, Inches(8.2), Inches(4.0), Inches(4.333), Inches(1.95), "📊 回傳資料結構 (JSON)")
    tb = s11.shapes.add_textbox(Inches(8.4), Inches(4.45), Inches(3.9), Inches(1.4))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = (
        "• title: 網頁文章標題\n"
        "• link: 原始 50 個網站來源超連結\n"
        "• snippet: 關鍵字高亮內文摘要\n"
        "• summary: Google Generative AI 智能整合摘要"
    )
    p.font.size = Pt(10)
    p.font.color.rgb = C_TEXT_BODY
    p.font.name = FONT_FAMILY

    add_checklist_bar(s11, "單元測試執行成功，Python 順暢透過私鑰調用 GCP 並取得結構化搜尋結果。")
    add_footer(s11, 11)

    # ==========================================
    # SLIDE 12: STEP 10 - FRONTEND INTEGRATION & VERIFICATION
    # ==========================================
    s12 = prs.slides.add_slide(blank_layout)
    set_bg(s12)
    add_sop_header(s12, 10, "整合驗證", "前端介面串接、搜尋驗證與端到端測試", "串聯前端 UI 與 Python API，完成端到端驗證並評估搜尋延遲與精準度")

    # Left: Two Integration Modes (Width: 6.8)
    add_card_box(s12, Inches(0.8), Inches(1.75), Inches(6.8), Inches(4.2), "🎨 前端整合支援兩種模式 (Dual Integration Modes)")
    tb = s12.shapes.add_textbox(Inches(1.0), Inches(2.25), Inches(6.4), Inches(3.5))
    tf = tb.text_frame
    tf.word_wrap = True
    s10_modes = [
        ("模式一：自訂 UI 搜尋介面 (推薦，自由度最高)", "• 架構：前端發送 AJAX POST `/api/search` 至 Python 後端。\n• 特色：UI 樣式與品牌形象 100% 自訂，卡片式呈現搜尋結果與 AI 摘要。\n• 資安：私鑰完全封裝於後端，前端只接收過濾後的搜尋資料。"),
        ("模式二：Google 官方現成小工具 (<gen-search-widget>)", "• 架構：前端直接載入 Google 官方 Web Component 小工具。\n• 特色：開箱即用，內建自動補全、分頁與精美排版。\n• 資安：後端提供 `/api/token` 動態派發短期 OAuth Token，前端不暴露私鑰。")
    ]
    for i, (st, desc) in enumerate(s10_modes):
        p = tf.add_paragraph() if i > 0 else tf.paragraphs[0]
        p.text = f"{st}"
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = C_BLUE_DARK
        p.font.name = FONT_FAMILY

        p2 = tf.add_paragraph()
        p2.text = f"{desc}\n"
        p2.font.size = Pt(9.5)
        p2.font.color.rgb = C_TEXT_BODY
        p2.font.name = FONT_FAMILY

    # Right Top: Verification Metrics
    add_card_box(s12, Inches(7.8), Inches(1.75), Inches(4.733), Inches(2.0), "📈 端到端測試驗證指標 (Benchmark)")
    tb = s12.shapes.add_textbox(Inches(8.0), Inches(2.25), Inches(4.3), Inches(1.3))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = (
        "• 搜尋延遲 (Latency)：平均 285ms (毫秒級迅速響應)\n"
        "• 50 網站白名單限制：100% 嚴格鎖定指定來源網域\n"
        "• AI 摘要生成品質：精準統整跨網站內容\n"
        "• 系統可用性 (Availability)：Google SLA 99.9%"
    )
    p.font.size = Pt(10)
    p.font.color.rgb = C_TEXT_BODY
    p.font.name = FONT_FAMILY

    # Right Bottom: User Acceptance
    add_card_box(s12, Inches(7.8), Inches(3.9), Inches(4.733), Inches(2.05), "🎯 主管驗收標準與示範情境")
    tb = s12.shapes.add_textbox(Inches(8.0), Inches(4.35), Inches(4.3), Inches(1.5))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = (
        "1. 輸入「半導體測試」等專業關鍵字，能即時呈現 50 個網站之對應文章與新聞。\n"
        "2. 點擊卡片連結能正確導向來源網站。\n"
        "3. 前後端完整解耦，系統具備高度擴展性。"
    )
    p.font.size = Pt(10)
    p.font.color.rgb = C_TEXT_BODY
    p.font.name = FONT_FAMILY

    add_checklist_bar(s12, "前後端串接完成，通過 50 網站跨站檢索測試，搜尋延遲低於 300ms。")
    add_footer(s12, 12)

    # ==========================================
    # SLIDE 13: OPERATIONS, SECURITY & COST SUMMARY
    # ==========================================
    s13 = prs.slides.add_slide(blank_layout)
    set_bg(s13)
    add_sop_header(s13, 0, "維運總結", "維運管理、資安防護與成本效益檢核表", "彙整企業長期運作方針、金鑰生命週期規範與後續技術演進規劃")

    # 3 Cards Layout Across
    card_w = Inches(3.75)
    card_h = Inches(4.2)

    # Card 1: Security & Governance
    add_card_box(s13, Inches(0.8), Inches(1.75), card_w, card_h, "🔒 資安防護與金鑰維運")
    tb1 = s13.shapes.add_textbox(Inches(1.0), Inches(2.25), card_w - Inches(0.4), Inches(3.5))
    tf1 = tb1.text_frame
    tf1.word_wrap = True
    p = tf1.paragraphs[0]
    p.text = (
        "• 金鑰定期輪替 (Key Rotation)：\n"
        "  每 90 天於 IAM 生成新金鑰並平滑替換舊金鑰，確保憑證時效安全。\n\n"
        "• Secret Manager 整合：\n"
        "  正式環境建議將 JSON 金鑰存入 Google Cloud Secret Manager，避免實體檔案儲存。\n\n"
        "• 網路層防護：\n"
        "  後端 API 可加入 IP 白名單或 Cloud Armor WAF 阻擋惡意請求。"
    )
    p.font.size = Pt(10)
    p.font.color.rgb = C_TEXT_BODY
    p.font.name = FONT_FAMILY

    # Card 2: Cost & Quota Optimization
    add_card_box(s13, Inches(4.8), Inches(1.75), card_w, card_h, "💰 成本分析與配額優化")
    tb2 = s13.shapes.add_textbox(Inches(5.0), Inches(2.25), card_w - Inches(0.4), Inches(3.5))
    tf2 = tb2.text_frame
    tf2.word_wrap = True
    p = tf2.paragraphs[0]
    p.text = (
        "• 計價機制透明：\n"
        "  Vertex AI Search 採搜尋查詢次數 (Queries) 計價，每月提供 1,000 次免費額度。\n\n"
        "• 快取機制 (Caching)：\n"
        "  後端可對熱門關鍵字建立 Redis / 記憶體快取 (TTL 1 小時)，降低 60% 重複呼叫成本。\n\n"
        "• 預算上限保護：\n"
        "  設定專案 Quotas 上限，防止非預期暴增之查詢費用。"
    )
    p.font.size = Pt(10)
    p.font.color.rgb = C_TEXT_BODY
    p.font.name = FONT_FAMILY

    # Card 3: Future Roadmap
    add_card_box(s13, Inches(8.8), Inches(1.75), card_w, card_h, "🚀 後續擴充與演進規劃")
    tb3 = s13.shapes.add_textbox(Inches(9.0), Inches(2.25), card_w - Inches(0.4), Inches(3.5))
    tf3 = tb3.text_frame
    tf3.word_wrap = True
    p = tf3.paragraphs[0]
    p.text = (
        "• 50 網站動態擴充：\n"
        "  日後增修網站名單僅需在 Data Store 後台修改 URL，程式碼零變更。\n\n"
        "• 擴充內部文件知識庫：\n"
        "  除了 50 個公開網站外，可掛載第二個 Data Store 匯入公司內部 PDF、Word 文件。\n\n"
        "• 對接企業通訊軟體：\n"
        "  本 Python 後端 API 可直接封裝為 Slack / Microsoft Teams / LINE 企業機器人。"
    )
    p.font.size = Pt(10)
    p.font.color.rgb = C_TEXT_BODY
    p.font.name = FONT_FAMILY

    add_checklist_bar(s13, "本 SOP 達成端到端完整落地，具備生產級資安、可控成本與未來高度擴充能力。")
    add_footer(s13, 13)

    # Save presentation
    prs.save(output_path)
    print(f"Presentation successfully created at: {output_path}")

if __name__ == "__main__":
    build_presentation()
