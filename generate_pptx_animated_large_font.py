# -*- coding: utf-8 -*-
"""
全面優化專案內所有 PowerPoint 簡報檔案：
1. slides.pptx & 第一場_Antigravity代理AI教學實戰_大字動畫版_GitHub發布版.pptx (28頁旗艦版)
   - 修正頁碼分母不一致問題 (原本前23頁寫 /23，後5頁寫 /28 -> 統一為 /28)
   - 修正單元跳號問題 (原本單元四直接跳第七單元 -> 梳理為單元一至單元六 + 最終總結頁)
   - 移除原本夾在中間的第 23 頁重複結尾頁，將收穫總結與思維升級整合至第 28 頁壓軸
   - 新增第 13 頁：【實測對決 03：電子學單元】BJT 共射極小訊號放大器與 r_pi 等效模型 (嵌入高解析電路圖與等效模型圖)
   - 更新第 12 頁：重新嵌入修正 \n 換行瑕疵後的 circuit3_max_power.png
   - 優化第 26、27 頁 Master Prompt 排版：由原本 8.5pt 密密麻麻小字改為雙欄結構化大字版 (11.5~13pt)，投影清晰易讀
   - 更新網址為真實上線之 https://stjedu.github.io/www/
2. yolo_webcam_slides.pptx & WebCam_AI視覺與YOLO應用_教學投影片.pptx (14頁)
   - 加入平滑淡出轉場動畫 (<p:fade/>)
   - 加入右上角頁碼 (第 XX / 14 頁) 與底部線上教材網址導覽列
   - 升級卡片頂部色條、放大提示詞與內文字級、強化視覺層次
3. esp32_mqtt_slides.pptx & ESP32_MQTT_智慧控制教學_教學投影片.pptx (11頁)
   - 加入平滑淡出轉場動畫 (<p:fade/>)
   - 加入右上角頁碼 (第 XX / 11 頁) 與底部線上教材網址導覽列
   - 升級卡片頂部色條、放大提示詞與內文字級、強化視覺層次
"""

import os
import shutil
import pptx
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls

ROOT_DIR = r"j:\我的雲端硬碟\antigravity\教師研習"
CIRCUITS_DIR = os.path.join(ROOT_DIR, "circuits")


def extract_instrument_screenshots():
    """從現有 slides.pptx 萃取第 24、25 頁的實機截圖並存入 circuits/ 資料夾，確保專案100%自給自足"""
    slides_path = os.path.join(ROOT_DIR, "slides.pptx")
    ch1_out = os.path.join(CIRCUITS_DIR, "instrument_ch1_demo.png")
    tab_out = os.path.join(CIRCUITS_DIR, "instrument_audit_tab.png")

    if os.path.exists(ch1_out) and os.path.exists(tab_out):
        return ch1_out, tab_out

    if os.path.exists(slides_path):
        prs_old = Presentation(slides_path)
        if len(prs_old.slides) >= 25:
            # Slide 24 (index 23)
            for shape in prs_old.slides[23].shapes:
                if shape.shape_type == 13:  # PICTURE
                    with open(ch1_out, "wb") as f:
                        f.write(shape.image.blob)
            # Slide 25 (index 24)
            for shape in prs_old.slides[24].shapes:
                if shape.shape_type == 13:  # PICTURE
                    with open(tab_out, "wb") as f:
                        f.write(shape.image.blob)
    return ch1_out, tab_out


def build_main_deck():
    ch1_img, tab_img = extract_instrument_screenshots()

    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    COLOR_BG_LIGHT       = RGBColor(250, 247, 242)   # #FAF7F2 溫暖米杏底色
    COLOR_PRIMARY_FOREST = RGBColor(31, 68, 53)      # #1F4435 典雅常春藤綠
    COLOR_CARAMEL        = RGBColor(194, 94, 46)     # #C25E2E 焦糖磚紅
    COLOR_AMBER          = RGBColor(224, 159, 62)    # #E09F3E 溫潤琥珀黃
    COLOR_SAGE           = RGBColor(142, 168, 157)   # #8EA89D 鼠尾草灰綠
    COLOR_CARD_BG        = RGBColor(255, 255, 255)   # #FFFFFF 純白卡片底
    COLOR_CARD_BORDER    = RGBColor(230, 224, 214)   # #E6E0D6 邊框灰
    COLOR_TEXT_MAIN      = RGBColor(43, 45, 66)      # #2B2D42 碳黑內文
    COLOR_TEXT_MUTED     = RGBColor(108, 117, 125)   # #6C757D 說明灰
    COLOR_PROMPT_BG      = RGBColor(254, 249, 239)   # #FEF9EF 羊皮紙提示詞底色
    COLOR_PROMPT_BORDER  = RGBColor(217, 138, 42)    # #D98A2A 琥珀金邊框
    COLOR_ALERT_ROSE     = RGBColor(188, 71, 73)     # #BC4749 玫瑰紅
    COLOR_EXCEL_GREEN    = RGBColor(16, 124, 65)     # #107C41 Excel 綠
    COLOR_ZEBRA_ROW      = RGBColor(248, 250, 249)   # 表格斑馬紋
    COLOR_FAIL_PINK      = RGBColor(254, 226, 226)   # 不及格粉紅
    COLOR_DARK_BOX       = RGBColor(15, 23, 42)      # 深色專業代碼卡片底
    COLOR_CODE_TEXT      = RGBColor(241, 245, 249)   # 淺白代碼字
    COLOR_PURPLE_BJT     = RGBColor(74, 37, 116)     # 電子學 BJT 紫

    FONT_FAMILY = "微軟正黑體"
    TOTAL_SLIDES = 28

    def apply_slide_transition(slide):
        trans_xml = parse_xml(f'<p:transition {nsdecls("p")} spd="med"><p:fade/></p:transition>')
        slide._element.append(trans_xml)

    def add_slide_scaffolding(slide, title_text, unit_name, slide_idx):
        apply_slide_transition(slide)

        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
        bg.fill.solid()
        bg.fill.fore_color.rgb = COLOR_BG_LIGHT
        bg.line.fill.background()

        top_bar1 = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, Inches(0.12))
        top_bar1.fill.solid()
        top_bar1.fill.fore_color.rgb = COLOR_PRIMARY_FOREST
        top_bar1.line.fill.background()

        top_bar2 = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, Inches(0.12), Inches(4.5), Inches(0.04))
        top_bar2.fill.solid()
        top_bar2.fill.fore_color.rgb = COLOR_CARAMEL
        top_bar2.line.fill.background()

        crumb_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.32), Inches(8.8), Inches(0.35))
        tf_c = crumb_box.text_frame
        tf_c.margin_left = tf_c.margin_top = tf_c.margin_right = tf_c.margin_bottom = 0
        p_c = tf_c.paragraphs[0]
        p_c.text = f"🌿 ANTIGRAVITY 研習手冊  |  {unit_name}"
        p_c.font.size = Pt(13)
        p_c.font.bold = True
        p_c.font.color.rgb = COLOR_CARAMEL
        p_c.font.name = FONT_FAMILY

        page_box = slide.shapes.add_textbox(Inches(10.0), Inches(0.32), Inches(2.5), Inches(0.35))
        tf_p = page_box.text_frame
        tf_p.margin_left = tf_p.margin_top = tf_p.margin_right = tf_p.margin_bottom = 0
        p_p = tf_p.paragraphs[0]
        p_p.alignment = PP_ALIGN.RIGHT
        p_p.text = f"第 {slide_idx:02d} / {TOTAL_SLIDES:02d} 頁"
        p_p.font.size = Pt(13)
        p_p.font.bold = True
        p_p.font.color.rgb = COLOR_TEXT_MUTED
        p_p.font.name = FONT_FAMILY

        title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.70), Inches(11.7), Inches(0.75))
        tf_t = title_box.text_frame
        tf_t.word_wrap = True
        tf_t.margin_left = tf_t.margin_top = tf_t.margin_right = tf_t.margin_bottom = 0
        p_t = tf_t.paragraphs[0]
        p_t.text = title_text
        p_t.font.size = Pt(27)
        p_t.font.bold = True
        p_t.font.color.rgb = COLOR_PRIMARY_FOREST
        p_t.font.name = FONT_FAMILY

        # 底部精緻資訊列
        footer_box = slide.shapes.add_textbox(Inches(0.8), Inches(7.12), Inches(11.7), Inches(0.3))
        tf_f = footer_box.text_frame
        tf_f.margin_left = tf_f.margin_top = tf_f.margin_right = tf_f.margin_bottom = 0
        pf = tf_f.paragraphs[0]
        pf.text = "🌐 線上教材與互動實驗室：https://stjedu.github.io/www/   |   主講：簡樹桐 老師"
        pf.font.size = Pt(10.5)
        pf.font.color.rgb = COLOR_TEXT_MUTED
        pf.font.name = FONT_FAMILY

    def add_notes(slide, notes_text):
        notes_slide = slide.notes_slide
        tf = notes_slide.notes_text_frame
        tf.text = notes_text

    def add_card(slide, left, top, width, height, bg_color=COLOR_CARD_BG, border_color=COLOR_CARD_BORDER):
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
        card.fill.solid()
        card.fill.fore_color.rgb = bg_color
        card.line.color.rgb = border_color
        card.line.width = Pt(1.5)
        return card

    def add_prompt_card(slide, left, top, width, height, prompt_title, prompt_content, font_size_pt=13.0):
        add_card(slide, left, top, width, height, bg_color=COLOR_PROMPT_BG, border_color=COLOR_PROMPT_BORDER)

        tag = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left + Inches(0.2), top + Inches(0.18), Inches(4.1), Inches(0.42))
        tag.fill.solid()
        tag.fill.fore_color.rgb = COLOR_CARAMEL
        tag.line.fill.background()
        tf_tag = tag.text_frame
        tf_tag.margin_left = tf_tag.margin_top = tf_tag.margin_right = tf_tag.margin_bottom = 0
        p_tag = tf_tag.paragraphs[0]
        p_tag.alignment = PP_ALIGN.CENTER
        p_tag.text = f"📜 {prompt_title}"
        p_tag.font.size = Pt(13)
        p_tag.font.bold = True
        p_tag.font.color.rgb = RGBColor(255, 255, 255)
        p_tag.font.name = FONT_FAMILY

        hint_box = slide.shapes.add_textbox(left + width - Inches(2.8), top + Inches(0.20), Inches(2.6), Inches(0.38))
        tf_h = hint_box.text_frame
        tf_h.margin_left = tf_h.margin_top = tf_h.margin_right = tf_h.margin_bottom = 0
        ph = tf_h.paragraphs[0]
        ph.alignment = PP_ALIGN.RIGHT
        ph.text = "💡 可直接反白複製使用"
        ph.font.size = Pt(11)
        ph.font.color.rgb = COLOR_TEXT_MUTED
        ph.font.name = FONT_FAMILY

        tb = slide.shapes.add_textbox(left + Inches(0.3), top + Inches(0.70), width - Inches(0.6), height - Inches(0.85))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0

        lines = prompt_content.strip().split('\n')
        for i, l in enumerate(lines):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.text = l
            p.font.size = Pt(font_size_pt)
            p.font.name = FONT_FAMILY
            p.font.color.rgb = COLOR_TEXT_MAIN
            p.line_spacing = 1.25

    # ==========================================
    # SLIDE 01: 封面頁
    # ==========================================
    s1 = prs.slides.add_slide(blank_layout)
    apply_slide_transition(s1)

    bg1 = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    bg1.fill.solid()
    bg1.fill.fore_color.rgb = COLOR_PRIMARY_FOREST
    bg1.line.fill.background()

    c1 = s1.shapes.add_shape(MSO_SHAPE.OVAL, Inches(9.5), Inches(-1.5), Inches(5.5), Inches(5.5))
    c1.fill.solid()
    c1.fill.fore_color.rgb = RGBColor(40, 85, 66)
    c1.line.fill.background()

    c2 = s1.shapes.add_shape(MSO_SHAPE.OVAL, Inches(-1.0), Inches(4.5), Inches(4.0), Inches(4.0))
    c2.fill.solid()
    c2.fill.fore_color.rgb = RGBColor(25, 55, 43)
    c2.line.fill.background()

    tb1 = s1.shapes.add_textbox(Inches(1.4), Inches(1.5), Inches(10.8), Inches(4.2))
    tf1 = tb1.text_frame
    tf1.word_wrap = True

    p0 = tf1.paragraphs[0]
    p0.text = "🌿 ANTIGRAVITY 代理 AI 於教學上的應用研習（旗艦完整版）"
    p0.font.size = Pt(18)
    p0.font.bold = True
    p0.font.color.rgb = COLOR_AMBER
    p0.font.name = FONT_FAMILY
    p0.space_after = Pt(14)

    p1 = tf1.add_paragraph()
    p1.text = "打通任督二脈：\n從對話到代理執行 ✕ 電學圖解 ✕ 教具開發"
    p1.font.size = Pt(42)
    p1.font.bold = True
    p1.font.color.rgb = RGBColor(250, 247, 242)
    p1.font.name = FONT_FAMILY
    p1.space_after = Pt(18)

    p2 = tf1.add_paragraph()
    p2.text = "環境巡檢建置 ✕ 基本電學與 BJT 小訊號圖解 ✕ 磨課師 MCP ✕ Excel 成績加權 ✕ 虛擬電子儀表實戰"
    p2.font.size = Pt(19)
    p2.font.color.rgb = RGBColor(218, 226, 219)
    p2.font.name = FONT_FAMILY

    tb_footer1 = s1.shapes.add_textbox(Inches(1.4), Inches(6.0), Inches(10.8), Inches(0.9))
    tf_f1 = tb_footer1.text_frame
    pf1 = tf_f1.paragraphs[0]
    pf1.text = "🎯 研習對象：中小學與技術型高中教師（零基礎友善）  |  🌐 研習網站：https://stjedu.github.io/www/"
    pf1.font.size = Pt(15)
    pf1.font.color.rgb = RGBColor(180, 195, 185)
    pf1.font.name = FONT_FAMILY

    add_notes(s1, "開場說辭：各位老師好！今天這場研習我們從零基礎環境建立出發，一路帶大家體驗基本電學與電子學 BJT 小訊號自動繪圖、MCP 自動修課、Excel 算成績，最後展示如何用一份大師級提示詞打造工業電子丙級虛擬示波器！")

    # ==========================================
    # SLIDE 02: 研習大綱學習地圖（6 大核心單元）
    # ==========================================
    s2 = prs.slides.add_slide(blank_layout)
    add_slide_scaffolding(s2, "今日研習大綱與六大單元學習地圖", "研習導覽", 2)

    agenda_items = [
        ("單元 01", "代理 AI 觀念啟蒙 (P.03~05)", "顧問 vs. 數位實習老師，跨時代的自主規劃、工具調用與實體建檔超能力", COLOR_PRIMARY_FOREST),
        ("單元 02", "環境建置與安全檢測 (P.06~09)", "AI 數位辦公桌建立、沙盒隔離機制保障、零代碼一鍵安裝 Python 與繪圖套件", COLOR_CARAMEL),
        ("單元 03", "提示語大師 ✕ 電學與電子學實測 (P.10~17)", "CRIS 架構、戴維寧等效電路、BJT 共射極 rπ 小訊號模型、MCP 磨課師自動伴讀", COLOR_AMBER),
        ("單元 04", "Excel 算成績全自動化 (P.18~20)", "段考實習加權算分、不及格粉紅底色標記、成績分布長條圖與學情分析摘要", COLOR_EXCEL_GREEN),
        ("單元 05", "成果網頁化與 GitHub 發布 (P.21~23)", "免費建立個人 www 倉庫、PAT 金鑰馬賽克資安保護、一鍵推播互動教學網站", COLOR_PURPLE_BJT),
        ("單元 06", "高階特輯：虛擬電子儀表開發 (P.24~28)", "依據三大原廠手冊打造 GDS-2102E 示波器、MFG-2110 信號源與丙級檢定模擬系統", RGBColor(2, 132, 199))
    ]
    for idx, (code, title, desc, col) in enumerate(agenda_items):
        col_idx = idx % 3
        row = idx // 3
        x = Inches(0.8 + col_idx * 4.0)
        y = Inches(1.65 + row * 2.6)
        add_card(s2, x, y, Inches(3.75), Inches(2.35))

        bar = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, Inches(3.75), Inches(0.12))
        bar.fill.solid()
        bar.fill.fore_color.rgb = col
        bar.line.fill.background()

        tb = s2.shapes.add_textbox(x + Inches(0.25), y + Inches(0.22), Inches(3.25), Inches(1.95))
        tf = tb.text_frame
        tf.word_wrap = True

        p = tf.paragraphs[0]
        p.text = code
        p.font.size = Pt(13)
        p.font.bold = True
        p.font.color.rgb = col
        p.font.name = FONT_FAMILY

        p_t = tf.add_paragraph()
        p_t.text = title
        p_t.font.size = Pt(17)
        p_t.font.bold = True
        p_t.font.color.rgb = COLOR_TEXT_MAIN
        p_t.font.name = FONT_FAMILY
        p_t.space_before = Pt(3)
        p_t.space_after = Pt(6)

        p_d = tf.add_paragraph()
        p_d.text = desc
        p_d.font.size = Pt(13.5)
        p_d.font.color.rgb = COLOR_TEXT_MUTED
        p_d.font.name = FONT_FAMILY
        p_d.line_spacing = 1.25

    # ==========================================
    # SLIDE 03: 痛點共鳴
    # ==========================================
    s3 = prs.slides.add_slide(blank_layout)
    add_slide_scaffolding(s3, "老師日常的三大行政與教學痛點", "單元一：觀念啟蒙", 3)

    pain_points = [
        ("📝 改考卷與算成績", "手動拉 Excel 函數\n段考平時加權算到眼花\n不及格篩選改色耗費數小時", COLOR_ALERT_ROSE),
        ("🌐 每年必修研習折磨", "磨課師/酷課雲必修幾十小時\n每 20 分鐘彈窗防閒置中斷\n播完不跳下一單元，被迫當點擊工人", COLOR_CARAMEL),
        ("🧠 剛入門的提示語焦慮", "面對空白對話框不知道怎麼說\n隨便問一句回答空洞、電路圖開天窗\n不知道如何下達具體執行規格", COLOR_PRIMARY_FOREST)
    ]
    for idx, (title, desc, color) in enumerate(pain_points):
        x = Inches(0.8 + idx * 4.0)
        add_card(s3, x, Inches(1.75), Inches(3.7), Inches(4.8))

        bar = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, Inches(1.75), Inches(3.7), Inches(0.16))
        bar.fill.solid()
        bar.fill.fore_color.rgb = color
        bar.line.fill.background()

        tb = s3.shapes.add_textbox(x + Inches(0.3), Inches(2.15), Inches(3.1), Inches(4.0))
        tf = tb.text_frame
        tf.word_wrap = True

        p = tf.paragraphs[0]
        p.text = title
        p.font.size = Pt(22)
        p.font.bold = True
        p.font.color.rgb = COLOR_TEXT_MAIN
        p.font.name = FONT_FAMILY
        p.space_after = Pt(20)

        p_body = tf.add_paragraph()
        p_body.text = desc
        p_body.font.size = Pt(16)
        p_body.font.color.rgb = COLOR_TEXT_MUTED
        p_body.font.name = FONT_FAMILY
        p_body.line_spacing = 1.45

    # ==========================================
    # SLIDE 04: 觀念革命
    # ==========================================
    s4 = prs.slides.add_slide(blank_layout)
    add_slide_scaffolding(s4, "從「對話問答」到「自主辦事」的跨時代躍升", "單元一：觀念啟蒙", 4)

    col_data = [
        ("傳統對話 AI（ChatGPT 等）", "就像一位【顧問】", [
            "只動嘴巴回答，不動手做事",
            "產出散落對話框，需手動複製貼上",
            "單次問答，無法連續跨檔案操作",
            "需要人盯著螢幕一步一步追問"
        ], COLOR_CARD_BG, COLOR_CARD_BORDER, COLOR_TEXT_MUTED),
        ("Antigravity 代理人 (Agent)", "就像一位【數位實習老師】", [
            "具備虛擬雙手，能直接在電腦建檔運算",
            "直接產出可開啟的 .xlsx、高解析電路圖與互動網頁",
            "支援「連續動作鏈」，自動執行多步驟並自我除錯",
            "下達目標後自主工作，完成時向您回報！"
        ], RGBColor(245, 249, 246), COLOR_PRIMARY_FOREST, COLOR_PRIMARY_FOREST)
    ]
    for idx, (head, subhead, bullets, bg_col, border_col, title_col) in enumerate(col_data):
        x = Inches(0.8 + idx * 6.0)
        add_card(s4, x, Inches(1.75), Inches(5.7), Inches(4.9), bg_color=bg_col, border_color=border_col)

        tb = s4.shapes.add_textbox(x + Inches(0.35), Inches(2.05), Inches(5.0), Inches(4.3))
        tf = tb.text_frame
        tf.word_wrap = True

        p = tf.paragraphs[0]
        p.text = head
        p.font.size = Pt(21)
        p.font.bold = True
        p.font.color.rgb = title_col
        p.font.name = FONT_FAMILY

        p_sub = tf.add_paragraph()
        p_sub.text = subhead
        p_sub.font.size = Pt(17)
        p_sub.font.bold = True
        p_sub.font.color.rgb = COLOR_CARAMEL
        p_sub.font.name = FONT_FAMILY
        p_sub.space_before = Pt(4)
        p_sub.space_after = Pt(14)

        for b in bullets:
            pb = tf.add_paragraph()
            pb.text = "• " + b
            pb.font.size = Pt(15.5)
            pb.font.color.rgb = COLOR_TEXT_MAIN
            pb.font.name = FONT_FAMILY
            pb.space_after = Pt(10)

    # ==========================================
    # SLIDE 05: 認識三大超能力
    # ==========================================
    s5 = prs.slides.add_slide(blank_layout)
    add_slide_scaffolding(s5, "認識 Antigravity 代理人的「三大超能力」", "單元一：觀念啟蒙", 5)

    powers = [
        ("🧠 自主規劃 (Plan)", "把大任務拆解為執行清單", "給定教學目標，自動啟動深度思考，拆解第一步、第二步、第三步，避免做白工。"),
        ("🛠️ 工具調用 (Tools)", "擁有讀寫檔案與運算的能力", "不只產出文字，能自主編寫並執行腳本，繪製標準 IEEE 電路圖、處理 Excel 加權與圖表。"),
        ("📦 工作區成果 (Artifacts)", "實體檔案直接存在你的硬碟", "所有成績表、電路圖與互動教具直接存在電腦資料夾，右側面板隨時即時預覽。")
    ]
    for idx, (title, sub, desc) in enumerate(powers):
        x = Inches(0.8 + idx * 4.0)
        add_card(s5, x, Inches(1.75), Inches(3.7), Inches(4.9))

        tag = s5.shapes.add_shape(MSO_SHAPE.OVAL, x + Inches(0.35), Inches(2.1), Inches(0.65), Inches(0.65))
        tag.fill.solid()
        tag.fill.fore_color.rgb = COLOR_PRIMARY_FOREST
        tag.line.fill.background()
        tf_tag = tag.text_frame
        pt = tf_tag.paragraphs[0]
        pt.alignment = PP_ALIGN.CENTER
        pt.text = str(idx + 1)
        pt.font.size = Pt(16)
        pt.font.bold = True
        pt.font.color.rgb = RGBColor(255, 255, 255)
        pt.font.name = FONT_FAMILY

        tb = s5.shapes.add_textbox(x + Inches(0.3), Inches(2.9), Inches(3.1), Inches(3.5))
        tf = tb.text_frame
        tf.word_wrap = True

        p = tf.paragraphs[0]
        p.text = title
        p.font.size = Pt(19)
        p.font.bold = True
        p.font.color.rgb = COLOR_PRIMARY_FOREST
        p.font.name = FONT_FAMILY

        p_sub = tf.add_paragraph()
        p_sub.text = sub
        p_sub.font.size = Pt(15)
        p_sub.font.bold = True
        p_sub.font.color.rgb = COLOR_CARAMEL
        p_sub.font.name = FONT_FAMILY
        p_sub.space_before = Pt(4)
        p_sub.space_after = Pt(12)

        p_d = tf.add_paragraph()
        p_d.text = desc
        p_d.font.size = Pt(14.5)
        p_d.font.color.rgb = COLOR_TEXT_MUTED
        p_d.font.name = FONT_FAMILY
        p_d.line_spacing = 1.35

    # ==========================================
    # SLIDE 06: 打造你的 AI 數位辦公桌
    # ==========================================
    s6 = prs.slides.add_slide(blank_layout)
    add_slide_scaffolding(s6, "打造你的「AI 數位辦公桌」：專屬教學工作區", "單元二：環境建置", 6)

    zones = [
        ("📂 左側：專案檔案庫", "左側工作區檔案樹", "打開 Antigravity 點選 Open Workspace，選擇或新建「2026_教學專案」資料夾。所有由 AI 建立的表格與教具都存於此。"),
        ("💬 中央：對話指揮中心", "與 AI 代理人的互動畫布", "如同你的數位助教辦公桌。在這裡下達 CRIS 提示詞，觀察 AI 的思考過程 (Thinking) 與工具執行。"),
        ("📱 右側：成果即時預覽區", "Artifacts 面板", "檔案產生後自動在此彈出預覽！試算表、電路圖與互動網頁直接呈現，免手動開外部軟體即可快速檢查確認。")
    ]
    for idx, (title, sub, sd) in enumerate(zones):
        x = Inches(0.8 + idx * 4.0)
        add_card(s6, x, Inches(1.75), Inches(3.7), Inches(4.9))

        tb = s6.shapes.add_textbox(x + Inches(0.3), Inches(2.05), Inches(3.1), Inches(4.3))
        tf = tb.text_frame
        tf.word_wrap = True

        p = tf.paragraphs[0]
        p.text = title
        p.font.size = Pt(20)
        p.font.bold = True
        p.font.color.rgb = COLOR_PRIMARY_FOREST
        p.font.name = FONT_FAMILY
        p.space_after = Pt(10)

        p2 = tf.add_paragraph()
        p2.text = sd
        p2.font.size = Pt(15)
        p2.font.color.rgb = COLOR_TEXT_MUTED
        p2.font.name = FONT_FAMILY
        p2.line_spacing = 1.35

    # ==========================================
    # SLIDE 07: 沙盒機制
    # ==========================================
    s7 = prs.slides.add_slide(blank_layout)
    add_slide_scaffolding(s7, "安心守則：什麼是「沙盒機制 (Sandbox)」？", "單元二：環境建置", 7)

    add_card(s7, Inches(0.8), Inches(1.75), Inches(11.7), Inches(4.9))
    tb_s7 = s7.shapes.add_textbox(Inches(1.2), Inches(2.05), Inches(10.9), Inches(4.3))
    tf_s7 = tb_s7.text_frame
    tf_s7.word_wrap = True

    p = tf_s7.paragraphs[0]
    p.text = "🛡️ 老師最常擔心的問題：「AI 操作我的電腦，會不會把我原本的重要公文改壞？」"
    p.font.size = Pt(20)
    p.font.bold = True
    p.font.color.rgb = COLOR_CARAMEL
    p.font.name = FONT_FAMILY
    p.space_after = Pt(16)

    points = [
        ("專屬隔離區（Sandbox 沙盒）", "就像在操場劃設安全的跑道，Antigravity 的所有自動化執行都被限制在「目前的專案資料夾」，絕對不會跑出去修改你的系統檔或桌面私人檔案。"),
        ("透明可見的執行過程", "AI 的每一步動作、修改了哪個檔案，全部在畫布上記錄得清清楚楚，隨時可以檢視與復原。"),
        ("資安與隱私第一準則", "練習時請勿使用真實學生的身分證字號、家長私人電話等敏感個資；API Token 務必妥善遮罩保存。")
    ]
    for t, d in points:
        pt = tf_s7.add_paragraph()
        pt.text = "🌿 " + t
        pt.font.size = Pt(17)
        pt.font.bold = True
        pt.font.color.rgb = COLOR_PRIMARY_FOREST
        pt.font.name = FONT_FAMILY
        pt.space_before = Pt(8)

        pd = tf_s7.add_paragraph()
        pd.text = "    " + d
        pd.font.size = Pt(15)
        pd.font.color.rgb = COLOR_TEXT_MUTED
        pd.font.name = FONT_FAMILY
        pd.space_after = Pt(8)

    # ==========================================
    # SLIDE 08: 初學者環境一鍵安裝與自動巡檢
    # ==========================================
    s8 = prs.slides.add_slide(blank_layout)
    add_slide_scaffolding(s8, "初學者必備：零代碼環境一鍵健檢與已裝自動化程式清冊", "單元二：環境建置", 8)

    add_card(s8, Inches(0.8), Inches(1.75), Inches(5.6), Inches(4.9))
    tb_i8 = s8.shapes.add_textbox(Inches(1.05), Inches(1.95), Inches(5.1), Inches(4.5))
    tf_i8 = tb_i8.text_frame
    tf_i8.word_wrap = True

    p = tf_i8.paragraphs[0]
    p.text = "🖥️ 本機實測已就緒之自動化工具清冊"
    p.font.size = Pt(18)
    p.font.bold = True
    p.font.color.rgb = COLOR_PRIMARY_FOREST
    p.font.name = FONT_FAMILY
    p.space_after = Pt(6)

    installed_groups = [
        ("⚡ 基礎環境核心", "Python 3.14 (直譯器) ✕ Node.js v24 (MCP運作) ✕ Git (版本/推播)"),
        ("📐 專業電學與電子學繪圖", "schemdraw (IEEE鋸齒電阻/BJT電晶體/受控源) ✕ matplotlib (曲線繪製) ✕ Pillow"),
        ("📊 Office 行政自動化", "python-pptx (簡報排版動畫) ✕ openpyxl / pandas (加權算分/學情統計)"),
        ("🌐 瀏覽器巡航與協定", "chrome-devtools-mcp (磨課師人機接力) ✕ KaTeX 數學公式渲染")
    ]
    for grp_title, grp_items in installed_groups:
        pt = tf_i8.add_paragraph()
        pt.text = grp_title
        pt.font.size = Pt(14.5)
        pt.font.bold = True
        pt.font.color.rgb = COLOR_CARAMEL
        pt.font.name = FONT_FAMILY

        pi = tf_i8.add_paragraph()
        pi.text = grp_items
        pi.font.size = Pt(12.5)
        pi.font.color.rgb = COLOR_TEXT_MAIN
        pi.font.name = FONT_FAMILY
        pi.space_after = Pt(5)

    p_bot = tf_i8.add_paragraph()
    p_bot.text = "💡 免敲 CMD 指令！只要下達右側咒語，代理人即在背景自動配齊上述工具。"
    p_bot.font.size = Pt(11.5)
    p_bot.font.bold = True
    p_bot.font.color.rgb = COLOR_TEXT_MUTED
    p_bot.font.name = FONT_FAMILY

    prompt8 = """請扮演我的「電腦教學環境裝備醫生」。我是一位剛接觸 AI 的零基礎老師，請幫我檢查並自動裝備好工作環境：

1. 檢查核心環境：Python 與 Node.js 是否已安裝就緒並回報版本？
2. 檢查 Git 工具：是否安裝 Git 版本控制（供後續推播 GitHub 網站使用）？
3. 檢查並自動安裝「教學必備套件庫」（若缺少請背景自動 pip install，免我開終端機）：
   - Office 自動化：openpyxl、pandas、python-pptx、python-docx
   - 專業電學與圖表：schemdraw（標準電路圖）、matplotlib（特性曲線）、pillow
4. 工作區讀寫測試：在當前目錄建立「健檢回報.txt」，並條列檢測狀態。
5. 完成後用親切繁體中文，以打勾 (✅) 方式回報所有裝備就緒清單！"""

    add_prompt_card(s8, Inches(6.6), Inches(1.75), Inches(5.9), Inches(4.9), "實戰提示詞 01：全自動環境健檢與裝備咒語", prompt8, 11.2)

    # ==========================================
    # SLIDE 09: 初試啼聲自主建檔
    # ==========================================
    s9 = prs.slides.add_slide(blank_layout)
    add_slide_scaffolding(s9, "初試啼聲：讓代理人自主在電腦建立第一份檔案", "單元二：環境建置", 9)

    add_card(s9, Inches(0.8), Inches(1.75), Inches(4.2), Inches(4.9))
    tb_i9 = s9.shapes.add_textbox(Inches(1.1), Inches(2.05), Inches(3.6), Inches(4.3))
    tf_i9 = tb_i9.text_frame
    tf_i9.word_wrap = True
    p = tf_i9.paragraphs[0]
    p.text = "🎯 實作目標"
    p.font.size = Pt(20)
    p.font.bold = True
    p.font.color.rgb = COLOR_PRIMARY_FOREST
    p.font.name = FONT_FAMILY
    p.space_after = Pt(14)

    tips9 = [
        "體驗代理 AI 的「造物」能力！",
        "觀察送出後：",
        "1. AI 思考中（Thinking）",
        "2. 工具調用（write_to_file）",
        "3. 左側檔案清單跳出「研習簽到與教學備忘.txt」",
        "4. 點開檔案，內容已真實存在硬碟中！"
    ]
    for t in tips9:
        p_t = tf_i9.add_paragraph()
        p_t.text = t
        p_t.font.size = Pt(15)
        p_t.font.color.rgb = COLOR_TEXT_MAIN
        p_t.font.name = FONT_FAMILY
        p_t.space_after = Pt(8)

    prompt9 = """你好！我是 [請填寫學校：例如 大安高工 / 彰師附工] 的 [請填寫學科：例如 電子科 / 資訊科] 老師。
請在我們的工作資料夾中，幫我建立一個名為「研習簽到與教學備忘.txt」的檔案，
內容寫入：
- 今日日期與時間
- 研習主題：Antigravity 代理 AI 教學應用實戰
- 我的教學願景：期許用 AI 幫自己每週省下 3 小時備課時間
- 給老師的一句備課激勵金句
建立完成後，請告訴我檔案已儲存。"""
    add_prompt_card(s9, Inches(5.3), Inches(1.75), Inches(7.2), Inches(4.9), "實戰提示詞範本 02：第一份自主建檔咒語", prompt9, 13.5)

    # ==========================================
    # SLIDE 10: 提示語大師修煉法（基本電學命題結構對比）
    # ==========================================
    s10 = prs.slides.add_slide(blank_layout)
    add_slide_scaffolding(s10, "為什麼別人用 AI 像神仙，我用 AI 像智障？", "單元三：提示語大師 ✕ 電學圖解", 10)

    add_card(s10, Inches(0.8), Inches(1.75), Inches(5.6), Inches(4.9))
    tb_b10 = s10.shapes.add_textbox(Inches(1.1), Inches(2.05), Inches(5.0), Inches(4.3))
    tf_b10 = tb_b10.text_frame
    tf_b10.word_wrap = True
    p = tf_b10.paragraphs[0]
    p.text = "❌ 失敗指令：隨意說一句話"
    p.font.size = Pt(20)
    p.font.bold = True
    p.font.color.rgb = COLOR_ALERT_ROSE
    p.font.name = FONT_FAMILY
    p.space_after = Pt(14)

    p_cmd1 = tf_b10.add_paragraph()
    p_cmd1.text = "「幫我出幾題基本電學題目。」"
    p_cmd1.font.size = Pt(16)
    p_cmd1.font.bold = True
    p_cmd1.font.color.rgb = COLOR_TEXT_MAIN
    p_cmd1.font.name = FONT_FAMILY
    p_cmd1.space_after = Pt(14)

    cons = [
        "AI 不知道給技術高中哪一科、幾年級學生做（高一初學？統測複習？）",
        "題目數值未經設計，常出現除不盡的繁雜小數，無法快速手算",
        "缺乏電路化簡結構，沒有串並聯拆解步驟與戴維寧等效解析",
        "老師還得花大量時間拿計算紙一題一題重新驗算電路"
    ]
    for c in cons:
        p_c = tf_b10.add_paragraph()
        p_c.text = "• " + c
        p_c.font.size = Pt(14)
        p_c.font.color.rgb = COLOR_TEXT_MUTED
        p_c.font.name = FONT_FAMILY
        p_c.space_after = Pt(8)

    add_card(s10, Inches(6.9), Inches(1.75), Inches(5.6), Inches(4.9), border_color=COLOR_PRIMARY_FOREST)
    tb_g10 = s10.shapes.add_textbox(Inches(7.2), Inches(2.05), Inches(5.0), Inches(4.3))
    tf_g10 = tb_g10.text_frame
    tf_g10.word_wrap = True
    p = tf_g10.paragraphs[0]
    p.text = "⭕ 大師級指令：結構化清晰下達"
    p.font.size = Pt(20)
    p.font.bold = True
    p.font.color.rgb = COLOR_PRIMARY_FOREST
    p.font.name = FONT_FAMILY
    p.space_after = Pt(12)

    p_cmd2 = tf_g10.add_paragraph()
    p_cmd2.text = "「你是一位技術高中電機電子群名師，正在教高一『基本電學：電阻網路分析』。請針對『直流電阻串並聯化簡與戴維寧定理 (Thevenin)』，設計 3 題統測試題等級的素養選擇題。要求：1. 數值經精心設計（電阻與電壓皆為整數，便於計算） 2. 每題包含清晰電路結構描述、4 個選項與標準答案 3. 詳解需列出完整推導步驟（含等效電阻 Rth 與等效電壓 Vth 之求法，並標註學生常錯的短路/開路盲點），並用 Python schemdraw 繪製標準閉迴路電路圖。」"
    p_cmd2.font.size = Pt(13)
    p_cmd2.font.bold = True
    p_cmd2.font.color.rgb = COLOR_TEXT_MAIN
    p_cmd2.font.name = FONT_FAMILY
    p_cmd2.space_after = Pt(12)

    pros = [
        "角色清晰：鎖定技高電機電子群高一認知程度",
        "數值友好：整數設計，避免繁雜小數干擾核心電路概念",
        "圖文並茂：自動呼叫 Python 繪製高解析 IEEE 鋸齒電阻電路圖",
        "產出即成果：一秒複製即可直接印成課堂隨堂評量卷！"
    ]
    for pr in pros:
        p_pr = tf_g10.add_paragraph()
        p_pr.text = "✔ " + pr
        p_pr.font.size = Pt(14)
        p_pr.font.color.rgb = COLOR_TEXT_MAIN
        p_pr.font.name = FONT_FAMILY
        p_pr.space_after = Pt(7)

    # ==========================================
    # SLIDE 11: 【圖文實測對決 01】隨意指令慘遭開天窗 ✕ 大師指令電路圖化簡
    # ==========================================
    s11 = prs.slides.add_slide(blank_layout)
    add_slide_scaffolding(s11, "實測對決 01：隨意指令慘遭開天窗 ✕ 大師指令電路圖化簡", "單元三：提示語大師 ✕ 電學圖解", 11)

    add_card(s11, Inches(0.8), Inches(1.75), Inches(5.4), Inches(5.15), border_color=COLOR_ALERT_ROSE)
    tb_fail = s11.shapes.add_textbox(Inches(1.05), Inches(1.95), Inches(4.9), Inches(4.8))
    tf_f = tb_fail.text_frame
    tf_f.word_wrap = True

    p = tf_f.paragraphs[0]
    p.text = "❌ 隨意指令產出（真實車禍現場）"
    p.font.size = Pt(19)
    p.font.bold = True
    p.font.color.rgb = COLOR_ALERT_ROSE
    p.font.name = FONT_FAMILY
    p.space_after = Pt(8)

    fail_items = [
        ("題 1：死背名詞", "「什麼是歐姆定律？請寫出公式。」\n➔ 國中生都知道，完全沒有技高電阻網路評量價值。"),
        ("題 2：數值極度不友善", "「100V 電壓接 33Ω 電阻，求電流？」\n➔ 答案 I = 3.0303...A（無限循環小數，學生算到懷疑人生）"),
        ("題 3：致命傷：電路開天窗！", "「試求下圖電路的戴維寧等效電路。（註：此處無圖）❌」\n➔ AI 憑空說下圖，根本沒有圖！題目直接作廢！"),
        ("老師下場：耗時重工", "無選項、無推導步驟，老師還得花半小時拿計算紙手動重出重算！")
    ]
    for tag, desc in fail_items:
        pt = tf_f.add_paragraph()
        pt.text = "• " + tag
        pt.font.size = Pt(14)
        pt.font.bold = True
        pt.font.color.rgb = COLOR_TEXT_MAIN
        pt.font.name = FONT_FAMILY

        pd = tf_f.add_paragraph()
        pd.text = "   " + desc
        pd.font.size = Pt(12.5)
        pd.font.color.rgb = COLOR_TEXT_MUTED
        pd.font.name = FONT_FAMILY
        pd.space_after = Pt(4)

    add_card(s11, Inches(6.5), Inches(1.75), Inches(6.0), Inches(5.15), border_color=COLOR_PRIMARY_FOREST)
    tb_succ1 = s11.shapes.add_textbox(Inches(6.75), Inches(1.9), Inches(5.5), Inches(1.3))
    tf_s1 = tb_succ1.text_frame
    tf_s1.word_wrap = True
    p = tf_s1.paragraphs[0]
    p.text = "⭕ 大師指令產出：題 01 電阻串並聯化簡"
    p.font.size = Pt(18)
    p.font.bold = True
    p.font.color.rgb = COLOR_PRIMARY_FOREST
    p.font.name = FONT_FAMILY

    p_body = tf_s1.add_paragraph()
    p_body.text = "【題幹】E = 24V，R1 = 6Ω 與 R2 = 12Ω 並聯後，再與 R3 = 8Ω 串聯。求 RT 與 IT？\n【選項】(A) RT=12Ω, IT=2A   (B) RT=14Ω, IT=1.7A   (C) RT=10Ω, IT=2.4A"
    p_body.font.size = Pt(12)
    p_body.font.color.rgb = COLOR_TEXT_MAIN
    p_body.font.name = FONT_FAMILY

    c1_path = os.path.join(CIRCUITS_DIR, 'circuit1_series_parallel.png')
    if os.path.exists(c1_path):
        s11.shapes.add_picture(c1_path, Inches(6.75), Inches(3.25), width=Inches(5.5))

    tb_succ1_sol = s11.shapes.add_textbox(Inches(6.75), Inches(5.8), Inches(5.5), Inches(1.05))
    tf_ss1 = tb_succ1_sol.text_frame
    tf_ss1.word_wrap = True
    p_sol = tf_ss1.paragraphs[0]
    p_sol.text = "✔ 正解 (A) | 推導：R1//R2 = 4Ω ➔ RT = 4 + 8 = 12Ω ➔ IT = 24 / 12 = 2A（整數手算超順！）"
    p_sol.font.size = Pt(11.5)
    p_sol.font.bold = True
    p_sol.font.color.rgb = COLOR_PRIMARY_FOREST
    p_sol.font.name = FONT_FAMILY

    p_sol2 = tf_ss1.add_paragraph()
    p_sol2.text = "🚨 常見盲點：學生易誤先算串聯 (6+12=18) 再並聯 8，詳解完整提示避免踩雷！"
    p_sol2.font.size = Pt(11)
    p_sol2.font.color.rgb = COLOR_CARAMEL
    p_sol2.font.name = FONT_FAMILY

    # ==========================================
    # SLIDE 12: 【圖文實測對決 02】戴維寧等效模型 ✕ 最大功率轉移曲線 (嵌入最新修正版圖片)
    # ==========================================
    s12 = prs.slides.add_slide(blank_layout)
    add_slide_scaffolding(s12, "實測對決 02：戴維寧等效模型 ✕ 最大功率轉移曲線", "單元三：提示語大師 ✕ 電學圖解", 12)

    add_card(s12, Inches(0.8), Inches(1.75), Inches(5.7), Inches(5.15))
    tb_t2 = s12.shapes.add_textbox(Inches(1.05), Inches(1.9), Inches(5.2), Inches(1.2))
    tf_t2 = tb_t2.text_frame
    tf_t2.word_wrap = True
    p = tf_t2.paragraphs[0]
    p.text = "⭕ 題 02：戴維寧等效模型與負載分析"
    p.font.size = Pt(18)
    p.font.bold = True
    p.font.color.rgb = COLOR_PRIMARY_FOREST
    p.font.name = FONT_FAMILY

    p2_b = tf_t2.add_paragraph()
    p2_b.text = "【題幹】開路電壓 Eth = 18V，等效電阻 Rth = 6Ω。外接負載 RL = 3Ω，求負載電流 IL？"
    p2_b.font.size = Pt(12)
    p2_b.font.color.rgb = COLOR_TEXT_MAIN
    p2_b.font.name = FONT_FAMILY

    c2_path = os.path.join(CIRCUITS_DIR, 'circuit2_thevenin.png')
    if os.path.exists(c2_path):
        s12.shapes.add_picture(c2_path, Inches(1.05), Inches(3.15), width=Inches(5.2))

    tb_t2_sol = s12.shapes.add_textbox(Inches(1.05), Inches(5.75), Inches(5.2), Inches(1.1))
    tf_t2s = tb_t2_sol.text_frame
    tf_t2s.word_wrap = True
    p = tf_t2s.paragraphs[0]
    p.text = "✔ 正解 (B) 2A | 單迴路串聯：IL = 18 / (6 + 3) = 2A（正負整數好算）"
    p.font.size = Pt(11.5)
    p.font.bold = True
    p.font.color.rgb = COLOR_PRIMARY_FOREST
    p.font.name = FONT_FAMILY
    p2 = tf_t2s.add_paragraph()
    p2.text = "🚨 防呆叮嚀：求 Rth 時獨立電壓源應「視為短路 (0V)」，學生常誤當開路！"
    p2.font.size = Pt(11)
    p2.font.color.rgb = COLOR_CARAMEL
    p2.font.name = FONT_FAMILY

    add_card(s12, Inches(6.8), Inches(1.75), Inches(5.7), Inches(5.15))
    tb_t3 = s12.shapes.add_textbox(Inches(7.05), Inches(1.9), Inches(5.2), Inches(1.2))
    tf_t3 = tb_t3.text_frame
    tf_t3.word_wrap = True
    p = tf_t3.paragraphs[0]
    p.text = "⭕ 題 03：戴維寧最大功率轉移特性"
    p.font.size = Pt(18)
    p.font.bold = True
    p.font.color.rgb = COLOR_PRIMARY_FOREST
    p.font.name = FONT_FAMILY

    p3_b = tf_t3.add_paragraph()
    p3_b.text = "【題幹】Eth = 20V，Rth = 5Ω。求使負載獲最大功率之 RL 阻值與最大功率 PL(max)？"
    p3_b.font.size = Pt(12)
    p3_b.font.color.rgb = COLOR_TEXT_MAIN
    p3_b.font.name = FONT_FAMILY

    c3_path = os.path.join(CIRCUITS_DIR, 'circuit3_max_power.png')
    if os.path.exists(c3_path):
        s12.shapes.add_picture(c3_path, Inches(7.05), Inches(3.15), width=Inches(5.2))

    tb_t3_sol = s12.shapes.add_textbox(Inches(7.05), Inches(5.75), Inches(5.2), Inches(1.1))
    tf_t3s = tb_t3_sol.text_frame
    tf_t3s.word_wrap = True
    p = tf_t3s.paragraphs[0]
    p.text = "✔ 正解 (A) RL=5Ω, PL=20W | 公式：PL(max) = Eth² / (4×Rth) = 400 / 20 = 20W"
    p.font.size = Pt(11.5)
    p.font.bold = True
    p.font.color.rgb = COLOR_PRIMARY_FOREST
    p.font.name = FONT_FAMILY
    p2 = tf_t3s.add_paragraph()
    p2.text = "🚨 防呆叮嚀：分母常忘乘係數 4，誤選 80W（選項 B 典型誘答陷阱）！"
    p2.font.size = Pt(11)
    p2.font.color.rgb = COLOR_CARAMEL
    p2.font.name = FONT_FAMILY

    # ==========================================
    # SLIDE 13: 【新增！實測對決 03：進階電子學】BJT 共射極小訊號放大 ✕ rπ 等效模型
    # ==========================================
    s13_bjt = prs.slides.add_slide(blank_layout)
    add_slide_scaffolding(s13_bjt, "實測對決 03：電子學 BJT 共射極放大器 ✕ rπ 小訊號等效模型", "單元三：提示語大師 ✕ 電子學專題", 13)

    # 左側卡片：BJT 實體電路圖與直流偏壓
    add_card(s13_bjt, Inches(0.8), Inches(1.75), Inches(5.7), Inches(5.15), border_color=COLOR_PURPLE_BJT)
    tb_bjt1 = s13_bjt.shapes.add_textbox(Inches(1.05), Inches(1.9), Inches(5.2), Inches(1.1))
    tf_bjt1 = tb_bjt1.text_frame
    tf_bjt1.word_wrap = True
    p = tf_bjt1.paragraphs[0]
    p.text = "🔌 步驟一：共射極 (CE) 放大電路與直流 Q 點"
    p.font.size = Pt(17.5)
    p.font.bold = True
    p.font.color.rgb = COLOR_PURPLE_BJT
    p.font.name = FONT_FAMILY

    p_bjt1_b = tf_bjt1.add_paragraph()
    p_bjt1_b.text = "【規格】VCC=12V, RB=565kΩ, RC=3kΩ, RL=3kΩ, β=100, VT=26mV，含耦合與旁路電容"
    p_bjt1_b.font.size = Pt(11.5)
    p_bjt1_b.font.color.rgb = COLOR_TEXT_MAIN
    p_bjt1_b.font.name = FONT_FAMILY

    c4_1_path = os.path.join(CIRCUITS_DIR, 'circuit4_schematic_ce_amp.png')
    if os.path.exists(c4_1_path):
        s13_bjt.shapes.add_picture(c4_1_path, Inches(1.15), Inches(2.95), height=Inches(2.65))

    tb_bjt1_sol = s13_bjt.shapes.add_textbox(Inches(1.05), Inches(5.7), Inches(5.2), Inches(1.15))
    tf_bjt1s = tb_bjt1_sol.text_frame
    tf_bjt1s.word_wrap = True
    p = tf_bjt1s.paragraphs[0]
    p.text = "✔ 直流與小訊號參數：IB = (12-0.7)/565k = 20µA ➔ IC = β·IB = 2mA"
    p.font.size = Pt(11.5)
    p.font.bold = True
    p.font.color.rgb = COLOR_PURPLE_BJT
    p.font.name = FONT_FAMILY
    p2 = tf_bjt1s.add_paragraph()
    p2.text = "⚡ 轉導 gm = IC / VT ≈ 76.92 mA/V  |  輸入電阻 rπ = β / gm = 1.3 kΩ"
    p2.font.size = Pt(11.5)
    p2.font.bold = True
    p2.font.color.rgb = COLOR_PRIMARY_FOREST
    p2.font.name = FONT_FAMILY

    # 右側卡片：r_pi 交流小訊號等效電路與增益
    add_card(s13_bjt, Inches(6.8), Inches(1.75), Inches(5.7), Inches(5.15), border_color=COLOR_PRIMARY_FOREST)
    tb_bjt2 = s13_bjt.shapes.add_textbox(Inches(7.05), Inches(1.9), Inches(5.2), Inches(1.1))
    tf_bjt2 = tb_bjt2.text_frame
    tf_bjt2.word_wrap = True
    p = tf_bjt2.paragraphs[0]
    p.text = "📐 步驟二：交流小訊號 rπ 等效電路與電壓增益"
    p.font.size = Pt(17.5)
    p.font.bold = True
    p.font.color.rgb = COLOR_PRIMARY_FOREST
    p.font.name = FONT_FAMILY

    p_bjt2_b = tf_bjt2.add_paragraph()
    p_bjt2_b.text = "【交流接地】直流電源 VCC 視為交流接地 (0V)，電容視為交流短路，射極 E 直接接地"
    p_bjt2_b.font.size = Pt(11.5)
    p_bjt2_b.font.color.rgb = COLOR_TEXT_MAIN
    p_bjt2_b.font.name = FONT_FAMILY

    c4_2_path = os.path.join(CIRCUITS_DIR, 'circuit4_r_pi_model.png')
    if os.path.exists(c4_2_path):
        s13_bjt.shapes.add_picture(c4_2_path, Inches(7.05), Inches(3.05), width=Inches(5.2))

    tb_bjt2_sol = s13_bjt.shapes.add_textbox(Inches(7.05), Inches(5.7), Inches(5.2), Inches(1.15))
    tf_bjt2s = tb_bjt2_sol.text_frame
    tf_bjt2s.word_wrap = True
    p = tf_bjt2s.paragraphs[0]
    p.text = "✔ 電壓增益：Av = vo/vs = -gm × (RC // RL) = -76.92m × (3k // 3k) ≈ -115.4 V/V"
    p.font.size = Pt(11.5)
    p.font.bold = True
    p.font.color.rgb = COLOR_PRIMARY_FOREST
    p.font.name = FONT_FAMILY
    p2 = tf_bjt2s.add_paragraph()
    p2.text = "🌐 互動模擬：stjedu.github.io/www/electronics_bjt.html（含即時 180° 反相波形與隨堂測驗）"
    p2.font.size = Pt(11)
    p2.font.bold = True
    p2.font.color.rgb = COLOR_CARAMEL
    p2.font.name = FONT_FAMILY

    # ==========================================
    # SLIDE 14: CRIS 提示語萬用框架
    # ==========================================
    s14 = prs.slides.add_slide(blank_layout)
    add_slide_scaffolding(s14, "提示語大師核心公式：CRIS 萬用架構", "單元三：提示語大師", 14)

    cris_boxes = [
        ("C - Context", "背景與角色", "設定 AI 的專業身分，以及面對的年級、學生程度與教學情境。\n例：『你是技高電機電子群名師，面對高二電子學學生...』", COLOR_PRIMARY_FOREST),
        ("R - Request", "明確目標任務", "以動詞開頭，直接告訴 AI 最終要完成什麼工作。\n例：『請設計 BJT 共射極小訊號 rπ 模型推導教材與互動計算網頁...』", COLOR_CARAMEL),
        ("I - Instructions", "條件與限制規則", "限制難易度、題型、字數、避免提及的內容。\n例：『使用 schemdraw 繪製 300DPI 電路圖、KaTeX 渲染公式、標註 180° 反相觀念...』", COLOR_AMBER),
        ("S - Style & Format", "風格與輸出格式", "指定輸出外觀，如單一 HTML 網頁、Markdown 表格或 Excel 試算表。\n例：『以溫暖校園風 HTML 呈現，含即時參數計算機與 4 題素養測驗...』", COLOR_SAGE)
    ]
    for idx, (code, title, desc, col) in enumerate(cris_boxes):
        row = idx // 2
        col_i = idx % 2
        x = Inches(0.8 + col_i * 6.0)
        y = Inches(1.75 + row * 2.5)
        add_card(s14, x, y, Inches(5.7), Inches(2.25))

        tb_c = s14.shapes.add_textbox(x + Inches(0.3), y + Inches(0.2), Inches(5.1), Inches(1.85))
        tf = tb_c.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = f"{code}（{title}）"
        p.font.size = Pt(18)
        p.font.bold = True
        p.font.color.rgb = col
        p.font.name = FONT_FAMILY
        p.space_after = Pt(8)

        p2 = tf.add_paragraph()
        p2.text = desc
        p2.font.size = Pt(14)
        p2.font.color.rgb = COLOR_TEXT_MAIN
        p2.font.name = FONT_FAMILY
        p2.line_spacing = 1.3

    # ==========================================
    # SLIDE 15: 連續動作鏈 (Chain of Action)
    # ==========================================
    s15 = prs.slides.add_slide(blank_layout)
    add_slide_scaffolding(s15, "代理人專用：下達「連續動作鏈 (Chain of Action)」", "單元三：提示語大師", 15)

    steps_chain = [
        ("步驟 1：讀取與繪圖", "「先用 Python schemdraw 繪製高解析度 BJT 共射極電路圖與 rπ 等效模型圖，存入 circuits/ 資料夾...」"),
        ("步驟 2：計算與建檔", "「接著建立 electronics_bjt.html，嵌入電路圖、KaTeX 公式推導、即時增益計算機與反相波形圖...」"),
        ("步驟 3：審查與推播", "「最後自我檢查所有公式渲染與圖片路徑無誤後，自動 Commit 並推播更新至 GitHub Pages！」")
    ]
    for idx, (st, ex) in enumerate(steps_chain):
        y = Inches(1.75 + idx * 1.62)
        add_card(s15, Inches(0.8), y, Inches(11.7), Inches(1.4))

        c_num = s15.shapes.add_shape(MSO_SHAPE.OVAL, Inches(1.1), y + Inches(0.25), Inches(0.9), Inches(0.9))
        c_num.fill.solid()
        c_num.fill.fore_color.rgb = COLOR_CARAMEL
        c_num.line.fill.background()
        tf_n = c_num.text_frame
        pn = tf_n.paragraphs[0]
        pn.alignment = PP_ALIGN.CENTER
        pn.text = str(idx + 1)
        pn.font.size = Pt(20)
        pn.font.bold = True
        pn.font.color.rgb = RGBColor(255, 255, 255)
        pn.font.name = FONT_FAMILY

        tb = s15.shapes.add_textbox(Inches(2.3), y + Inches(0.18), Inches(9.8), Inches(1.0))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = st
        p.font.size = Pt(18)
        p.font.bold = True
        p.font.color.rgb = COLOR_PRIMARY_FOREST
        p.font.name = FONT_FAMILY

        p2 = tf.add_paragraph()
        p2.text = "範例指令：" + ex
        p2.font.size = Pt(15)
        p2.font.color.rgb = COLOR_TEXT_MUTED
        p2.font.name = FONT_FAMILY
        p2.space_before = Pt(4)

    # ==========================================
    # SLIDE 16: 神級代理人：MCP 瀏覽器自動化 —— 人機接力與 1 鍵啟動
    # ==========================================
    s16 = prs.slides.add_slide(blank_layout)
    add_slide_scaffolding(s16, "神級代理人：MCP 瀏覽器自動化 —— 人機接力與 1 鍵啟動", "單元三：提示語大師 ✕ MCP", 16)

    add_card(s16, Inches(0.8), Inches(1.75), Inches(5.6), Inches(4.9))
    tb_i15 = s16.shapes.add_textbox(Inches(1.1), Inches(2.0), Inches(5.0), Inches(4.4))
    tf_i15 = tb_i15.text_frame
    tf_i15.word_wrap = True

    p = tf_i15.paragraphs[0]
    p.text = "🤝 30 秒人機接力黃金法則"
    p.font.size = Pt(20)
    p.font.bold = True
    p.font.color.rgb = COLOR_PRIMARY_FOREST
    p.font.name = FONT_FAMILY
    p.space_after = Pt(6)

    p_gold = tf_i15.add_paragraph()
    p_gold.text = "💡「資安歸老師，苦工歸 AI！」"
    p_gold.font.size = Pt(16)
    p_gold.font.bold = True
    p_gold.font.color.rgb = COLOR_CARAMEL
    p_gold.font.name = FONT_FAMILY
    p_gold.space_after = Pt(12)

    steps_mcp = [
        ("👤 老師親手做（前 30 秒免踩雷）：", [
            "1. 點兩下「啟動AI伴讀Chrome.bat」（遠端除錯埠 9222）",
            "2. 親自登入教育雲 OpenID（免個資外洩、免驗證碼）",
            "3. 點開課程第一單元，按一次播放鍵！"
        ]),
        ("🤖 AI 代理人接管（後 30 分鐘全代勞）：", [
            "• 0.5 秒自動代點「下一單元」微小按鈕",
            "• 自動按掉「禁止多重視窗」等中斷警告彈窗",
            "• 每 24 分鐘微暫停 5 秒重置防閒置計時"
        ])
    ]
    for title, items in steps_mcp:
        pt = tf_i15.add_paragraph()
        pt.text = title
        pt.font.size = Pt(15)
        pt.font.bold = True
        pt.font.color.rgb = COLOR_TEXT_MAIN
        pt.font.name = FONT_FAMILY
        pt.space_before = Pt(4)

        for it in items:
            pi = tf_i15.add_paragraph()
            pi.text = "   " + it
            pi.font.size = Pt(13)
            pi.font.color.rgb = COLOR_TEXT_MUTED
            pi.font.name = FONT_FAMILY
            pi.space_after = Pt(2)

    prompt15 = """請使用 chrome-devtools-mcp 連接我目前開啟的 Chrome 瀏覽器：
1. 找到我正在播放研習影片的磨課師（或數位學習平臺）分頁。
2. 幫我監控影片進度：當前單元播完時，自動幫我點擊「下一單元」繼續播放。
3. 若畫面彈出「禁止多重視窗」或確認警示窗，請自動點擊「確定」排除。
4. 開始執行，並在每完成一個章節時向我回報！"""
    add_prompt_card(s16, Inches(6.8), Inches(1.75), Inches(5.7), Inches(4.9), "提示詞範本 03-A：生手「一鍵通」秒殺指令", prompt15, 13.0)

    # ==========================================
    # SLIDE 17: 生手一次就成功：全自動巡航 ✕ 突發狀況急救包
    # ==========================================
    s17 = prs.slides.add_slide(blank_layout)
    add_slide_scaffolding(s17, "生手一次就成功：全自動巡航 ✕ 突發狀況急救包", "單元三：提示語大師 ✕ MCP", 17)

    prompt16 = """【角色】請扮演我的「數位研習全自動伴讀導航員」。
【情境】我已登入磨課師研習平臺（Chrome 9222 遠端除錯埠），畫面正停在課程播放頁。
【任務與執行規則】：
1. 鎖定分頁：使用 list_pages 與 select_page 鎖定包含「moocs」或「elearning」課程分頁。
2. 連播機制：以 1.0x 正常速度播放，影片播畢自動點擊「下一單元（nav_next）」。
3. 防閒置破解：每播 24 分鐘自動微暫停 5 秒後恢復，重置 25/30 分鐘防閒置計時器。
4. 彈窗攔截：呼叫 handle_dialog，一旦有警告彈窗 0 秒點確定。
5. 持續巡航，並在右側記錄完成章節數與累積研習時數！"""
    add_prompt_card(s17, Inches(0.8), Inches(1.75), Inches(6.0), Inches(4.9), "提示詞範本 03-B：大師級全自動巡航 CRIS 咒語", prompt16, 12.0)

    add_card(s17, Inches(7.1), Inches(1.75), Inches(5.4), Inches(4.9), border_color=COLOR_CARAMEL)
    tb_i16 = s17.shapes.add_textbox(Inches(7.35), Inches(2.0), Inches(4.9), Inches(4.4))
    tf_i16 = tb_i16.text_frame
    tf_i16.word_wrap = True

    p = tf_i16.paragraphs[0]
    p.text = "🛠️ 研習現場突發狀況「急救包」"
    p.font.size = Pt(20)
    p.font.bold = True
    p.font.color.rgb = COLOR_CARAMEL
    p.font.name = FONT_FAMILY
    p.space_after = Pt(4)

    p_sub = tf_i16.add_paragraph()
    p_sub.text = "遇到意外狀況？複製這幾句咒語貼上立即排除："
    p_sub.font.size = Pt(13)
    p_sub.font.color.rgb = COLOR_TEXT_MUTED
    p_sub.font.name = FONT_FAMILY
    p_sub.space_after = Pt(10)

    troubles = [
        ("⚠️ 狀況 1：抓錯分頁", "「請列出 Chrome 所有開啟分頁標題，並將焦點切換到磨課師課程那一頁。」"),
        ("⚠️ 狀況 2：彈窗卡住", "「畫面上跳出警告對話框，請呼叫 handle_dialog 幫我自動按確定關閉。」"),
        ("⚠️ 狀況 3：按鈕微小", "「請在當前網頁尋找包含『下一頁』或『下一單元』按鈕並點擊它。」"),
        ("⚠️ 狀況 4：查詢進度", "「請幫我查看當前課程的累計閱讀時數與完成進度百分比。」")
    ]
    for tag, cmd in troubles:
        pt = tf_i16.add_paragraph()
        pt.text = tag
        pt.font.size = Pt(14)
        pt.font.bold = True
        pt.font.color.rgb = COLOR_PRIMARY_FOREST
        pt.font.name = FONT_FAMILY

        pc = tf_i16.add_paragraph()
        pc.text = "   " + cmd
        pc.font.size = Pt(12)
        pc.font.color.rgb = COLOR_TEXT_MAIN
        pc.font.name = FONT_FAMILY
        pc.space_after = Pt(6)

    # ==========================================
    # SLIDE 18: Office 自動化第一彈
    # ==========================================
    s18 = prs.slides.add_slide(blank_layout)
    add_slide_scaffolding(s18, "Office 自動化第一彈：手動拉公式算成績的終結者", "單元四：Excel 自動化", 18)

    cols17 = [
        ("傳統算成績的痛苦", "• 總分、平均手動打 =SUM、=AVERAGE\n• 不及格要一欄一欄篩選、改粉紅色\n• 等第判定寫巢狀 =IF 寫到眼花\n• 全班最高分、最低分還要再算一次", COLOR_ALERT_ROSE),
        ("Antigravity 代理人解法", "• 把原始資料給它，或直接模擬整班成績\n• AI 在背景自行執行演算法運算\n• 自動標記不及格粉紅色底色\n• 幾秒鐘直接產出格式化完成的 .xlsx", COLOR_EXCEL_GREEN)
    ]
    for idx, (title, content, col) in enumerate(cols17):
        x = Inches(0.8 + idx * 6.0)
        add_card(s18, x, Inches(1.75), Inches(5.7), Inches(4.9))
        tb = s18.shapes.add_textbox(x + Inches(0.3), Inches(2.05), Inches(5.1), Inches(4.3))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = title
        p.font.size = Pt(22)
        p.font.bold = True
        p.font.color.rgb = col
        p.font.name = FONT_FAMILY
        p.space_after = Pt(18)

        for l in content.split('\n'):
            pl = tf.add_paragraph()
            pl.text = l
            pl.font.size = Pt(16)
            pl.font.color.rgb = COLOR_TEXT_MAIN
            pl.font.name = FONT_FAMILY
            pl.space_after = Pt(12)

    # ==========================================
    # SLIDE 19: 實戰演練：全自動班級成績統計與等第運算
    # ==========================================
    s19 = prs.slides.add_slide(blank_layout)
    add_slide_scaffolding(s19, "實戰演練：全自動班級成績統計與等第運算", "單元四：Excel 自動化", 19)

    prompt18 = """請扮演我的「教務與實習成績統計助理」。
任務：在工作區建立「高二電子科專業實習第一次段考成績表.xlsx」。
資料要求：
1. 模擬 15 位技高學生名單（座號 1~15、學生姓名用王小明等假名），包含三項成績：
   - 專業筆試 (占 30%)
   - 示波器實作量測 (占 50%)
   - 平時實習表現與工安規範 (占 20%)
2. 自動計算：
   - 每位學生「學期加權總分（四捨五入至小數第一位）」。
   - 根據加權總分判定「等第」（90優/80甲/70乙/60丙/未滿60丁）。
   - 實作或筆試未滿 60 分用醒目粉紅色底色標記。
3. 表格最下方計算各項「全班平均」、「最高分」、「最低分」。
完成後儲存為 .xlsx 檔案並提供預覽。"""
    add_prompt_card(s19, Inches(0.8), Inches(1.75), Inches(5.9), Inches(4.9), "實戰提示詞範本 04：Excel 自動算成績大師", prompt18, 12.0)

    add_card(s19, Inches(7.0), Inches(1.75), Inches(5.5), Inches(4.9), bg_color=RGBColor(255, 255, 255), border_color=COLOR_EXCEL_GREEN)

    excel_bar = s19.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(7.2), Inches(1.95), Inches(5.1), Inches(0.42))
    excel_bar.fill.solid()
    excel_bar.fill.fore_color.rgb = COLOR_EXCEL_GREEN
    excel_bar.line.fill.background()
    tf_eb = excel_bar.text_frame
    tf_eb.margin_left = tf_eb.margin_top = tf_eb.margin_right = tf_eb.margin_bottom = 0
    p_eb = tf_eb.paragraphs[0]
    p_eb.alignment = PP_ALIGN.CENTER
    p_eb.text = "📊 預期產出：高二電子科專業實習第一次段考成績表.xlsx"
    p_eb.font.size = Pt(12)
    p_eb.font.bold = True
    p_eb.font.color.rgb = RGBColor(255, 255, 255)
    p_eb.font.name = FONT_FAMILY

    table_shape = s19.shapes.add_table(6, 6, Inches(7.2), Inches(2.5), Inches(5.1), Inches(3.85))
    table = table_shape.table
    table.columns[0].width = Inches(0.7)
    table.columns[1].width = Inches(0.9)
    table.columns[2].width = Inches(0.8)
    table.columns[3].width = Inches(0.8)
    table.columns[4].width = Inches(0.8)
    table.columns[5].width = Inches(1.1)

    headers = ["座號", "姓名", "筆試(30%)", "實作(50%)", "平時(20%)", "加權/等第"]
    mock_data = [
        ["01", "王小明", "88", "92", "85", "89.4 (甲)"],
        ["02", "李小華", "55", "74", "60", "65.5 (丙)"],
        ["03", "張大同", "95", "96", "95", "95.5 (優)"],
        ["04", "陳美麗", "72", "48", "75", "60.6 (丙)"],
        ["班平", "全班平均", "77.5", "77.5", "78.8", "77.8"]
    ]

    for col_idx, h in enumerate(headers):
        cell = table.cell(0, col_idx)
        cell.fill.solid()
        cell.fill.fore_color.rgb = COLOR_EXCEL_GREEN
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = cell.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        p.text = h
        p.font.size = Pt(12)
        p.font.bold = True
        p.font.color.rgb = RGBColor(255, 255, 255)
        p.font.name = FONT_FAMILY

    for row_idx, r in enumerate(mock_data):
        for col_idx, val in enumerate(r):
            cell = table.cell(row_idx + 1, col_idx)
            cell.fill.solid()
            if row_idx == 4:
                cell.fill.fore_color.rgb = RGBColor(230, 240, 235)
            elif (row_idx == 1 and col_idx == 2) or (row_idx == 3 and col_idx == 3):
                cell.fill.fore_color.rgb = COLOR_FAIL_PINK
            elif row_idx % 2 == 1:
                cell.fill.fore_color.rgb = COLOR_ZEBRA_ROW
            else:
                cell.fill.fore_color.rgb = RGBColor(255, 255, 255)

            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = cell.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            p.text = val
            p.font.size = Pt(11)
            p.font.name = FONT_FAMILY
            if (row_idx == 1 and col_idx == 2) or (row_idx == 3 and col_idx == 3):
                p.font.bold = True
                p.font.color.rgb = COLOR_ALERT_ROSE
            elif row_idx == 4:
                p.font.bold = True
                p.font.color.rgb = COLOR_PRIMARY_FOREST
            else:
                p.font.color.rgb = COLOR_TEXT_MAIN

    # ==========================================
    # SLIDE 20: 進階演練：成績深入學情分析與圖表自動生成
    # ==========================================
    s20 = prs.slides.add_slide(blank_layout)
    add_slide_scaffolding(s20, "進階演練：成績深入學情分析與圖表自動生成", "單元四：Excel 自動化", 20)

    prompt19 = """請讀取剛剛建立的「高二電子科專業實習第一次段考成績表.xlsx」，進行深入學情分析：
1. 請繪製一張「實作成績分布級距長條圖」（以 10 分為一個級距），存成圖檔放入工作區。
2. 請列出分析摘要：
   - 班上實作前三名同學是誰？儀器操作亮點在哪？
   - 筆試或實作不及格人數各有幾位？
3. 請幫我寫一段 200 字左右的「技高專業實習學情檢討摘要」，適合直接複製到我的實習教學日誌或教學研究會報告中。"""
    add_prompt_card(s20, Inches(0.8), Inches(1.75), Inches(11.7), Inches(4.9), "實戰提示詞範本 05：學情診斷與視覺化圖表生成", prompt19, 14.0)

    # ==========================================
    # SLIDE 21: 成果發布 01：GitHub 帳號註冊 ✕ 建立 www 倉庫
    # ==========================================
    s21 = prs.slides.add_slide(blank_layout)
    add_slide_scaffolding(s21, "成果發布 01：免費帳號開立 ✕ 建立專屬倉庫 (www)", "單元五：成果網頁化與 GitHub 發布", 21)

    add_card(s21, Inches(0.8), Inches(1.75), Inches(5.6), Inches(4.9))
    tb_l20 = s21.shapes.add_textbox(Inches(1.1), Inches(2.05), Inches(5.0), Inches(4.3))
    tf_l20 = tb_l20.text_frame
    tf_l20.word_wrap = True
    p = tf_l20.paragraphs[0]
    p.text = "📝 步驟一：註冊個人 GitHub 帳號"
    p.font.size = Pt(20)
    p.font.bold = True
    p.font.color.rgb = COLOR_PRIMARY_FOREST
    p.font.name = FONT_FAMILY
    p.space_after = Pt(10)

    gh_steps = [
        ("1. 前往官網", "瀏覽器打開 github.com，點右上角綠色「Sign up」。"),
        ("2. 填寫三要素", "輸入信箱（收驗證碼）➔ 設定高強度密碼 ➔ 決定 Username。"),
        ("⭐ 命名即網址！", "Username 會直接變成個人教學網站網址！例如簡老師為 stjedu，日後專屬站點即為 stjedu.github.io。"),
        ("3. 驗證啟用", "完成防偽驗證，收信填入驗證碼即可啟用免費雲端空間！")
    ]
    for title, desc in gh_steps:
        pt = tf_l20.add_paragraph()
        pt.text = title
        pt.font.size = Pt(15)
        pt.font.bold = True
        pt.font.color.rgb = COLOR_CARAMEL if "命名" in title else COLOR_PRIMARY_FOREST
        pt.font.name = FONT_FAMILY
        pd = tf_l20.add_paragraph()
        pd.text = desc
        pd.font.size = Pt(13)
        pd.font.color.rgb = COLOR_TEXT_MUTED
        pd.font.name = FONT_FAMILY
        pd.space_after = Pt(6)

    add_card(s21, Inches(6.9), Inches(1.75), Inches(5.6), Inches(4.9), border_color=COLOR_CARAMEL)
    tb_r20 = s21.shapes.add_textbox(Inches(7.2), Inches(2.05), Inches(5.0), Inches(4.3))
    tf_r20 = tb_r20.text_frame
    tf_r20.word_wrap = True
    p = tf_r20.paragraphs[0]
    p.text = "📂 步驟二：建立名為 www 的儲存庫"
    p.font.size = Pt(20)
    p.font.bold = True
    p.font.color.rgb = COLOR_CARAMEL
    p.font.name = FONT_FAMILY
    p.space_after = Pt(10)

    repo_steps = [
        ("1. 點擊右上角「+」號", "在 GitHub 畫面右上角點選「New repository」。"),
        ("2. 倉庫名稱填寫 www", "Repository name 輸入 www，代表個人教學首頁根目錄。"),
        ("3. 務必勾選 Public（公開）", "勾選 Public，其餘 README、.gitignore 保持預設空白免勾。"),
        ("4. 點綠色按鈕建立", "按下「Create repository」，取得網址：github.com/stjedu/www。")
    ]
    for title, desc in repo_steps:
        pt = tf_r20.add_paragraph()
        pt.text = title
        pt.font.size = Pt(15)
        pt.font.bold = True
        pt.font.color.rgb = COLOR_PRIMARY_FOREST
        pt.font.name = FONT_FAMILY
        pd = tf_r20.add_paragraph()
        pd.text = desc
        pd.font.size = Pt(13)
        pd.font.color.rgb = COLOR_TEXT_MUTED
        pd.font.name = FONT_FAMILY
        pd.space_after = Pt(6)

    # ==========================================
    # SLIDE 22: 成果發布 02：產生安全 Token 權杖 (馬賽克保護)
    # ==========================================
    s22 = prs.slides.add_slide(blank_layout)
    add_slide_scaffolding(s22, "成果發布 02：產生個人安全金鑰 (Personal Access Token)", "單元五：成果網頁化與 GitHub 發布", 22)

    add_card(s22, Inches(0.8), Inches(1.75), Inches(5.6), Inches(4.9))
    tb_l21 = s22.shapes.add_textbox(Inches(1.1), Inches(2.05), Inches(5.0), Inches(4.3))
    tf_l21 = tb_l21.text_frame
    tf_l21.word_wrap = True
    p = tf_l21.paragraphs[0]
    p.text = "🔑 步驟三：為什麼不能用登入密碼？"
    p.font.size = Pt(19)
    p.font.bold = True
    p.font.color.rgb = COLOR_PRIMARY_FOREST
    p.font.name = FONT_FAMILY
    p.space_after = Pt(8)

    token_why = [
        ("🛡️ GitHub 資安規範升級", "為防止密碼遭竊，GitHub 禁止終端機推播使用傳統密碼，必須改用 Personal Access Token (PAT)。"),
        ("1. 前往金鑰設定頁面", "瀏覽器前往 github.com/settings/tokens，點 Generate new token (classic)。"),
        ("2. Note 填寫 antigravity", "有效期限選 30 天或自訂天數。"),
        ("3. ⭐ 關鍵勾選 repo", "請務必將第一個「repo」（倉庫完整存取）打勾！"),
        ("4. 複製金鑰代碼", "點擊最下方綠色按鈕，複製開頭為 ghp_ 的專屬權杖！")
    ]
    for title, desc in token_why:
        pt = tf_l21.add_paragraph()
        pt.text = title
        pt.font.size = Pt(14)
        pt.font.bold = True
        pt.font.color.rgb = COLOR_CARAMEL if "關鍵" in title else COLOR_PRIMARY_FOREST
        pt.font.name = FONT_FAMILY
        pd = tf_l21.add_paragraph()
        pd.text = desc
        pd.font.size = Pt(12.5)
        pd.font.color.rgb = COLOR_TEXT_MUTED
        pd.font.name = FONT_FAMILY
        pd.space_after = Pt(4)

    add_card(s22, Inches(6.9), Inches(1.75), Inches(5.6), Inches(4.9), border_color=COLOR_CARAMEL)
    tb_r21 = s22.shapes.add_textbox(Inches(7.2), Inches(2.05), Inches(5.0), Inches(4.3))
    tf_r21 = tb_r21.text_frame
    tf_r21.word_wrap = True
    p = tf_r21.paragraphs[0]
    p.text = "🔒 資安第一：金鑰遮罩防洩漏守則"
    p.font.size = Pt(19)
    p.font.bold = True
    p.font.color.rgb = COLOR_CARAMEL
    p.font.name = FONT_FAMILY
    p.space_after = Pt(10)

    security_rules = [
        ("🚨 金鑰等同於您的臨時身分證！", "任何人拿到您的 Token 都可以修改或覆寫您的雲端倉庫，切勿寫進公開腳本。"),
        ("🙈 示範或講義中必須打馬賽克：", "在簡報或公開錄影時，代碼必須做馬賽克遮罩處理："),
        ("   代碼示範：", "token = 'ghp_████████████████████████████████'"),
        ("⚡ 代理人自動化驗證機制：", "結合 Windows Git Credential Manager，代理人可安全完成推播，無需在硬碟留存任何明文金鑰！")
    ]
    for title, desc in security_rules:
        pt = tf_r21.add_paragraph()
        pt.text = title
        pt.font.size = Pt(14)
        pt.font.bold = True
        pt.font.color.rgb = COLOR_ALERT_ROSE if "🚨" in title else COLOR_PRIMARY_FOREST
        pt.font.name = FONT_FAMILY
        pd = tf_r21.add_paragraph()
        pd.text = desc
        pd.font.size = Pt(12.5)
        pd.font.color.rgb = COLOR_TEXT_MAIN
        pd.font.name = FONT_FAMILY
        pd.space_after = Pt(6)

    # ==========================================
    # SLIDE 23: 成果發布 03：代理人推播動作鏈 ✕ 網站上線 (Prompt 06)
    # ==========================================
    s23 = prs.slides.add_slide(blank_layout)
    add_slide_scaffolding(s23, "成果發布 03：代理人一鍵推播 ✕ GitHub Pages 上線 (Prompt 06)", "單元五：成果網頁化與 GitHub 發布", 23)

    add_card(s23, Inches(0.8), Inches(1.75), Inches(5.6), Inches(4.9))
    tb_l22 = s23.shapes.add_textbox(Inches(1.1), Inches(2.05), Inches(5.0), Inches(4.3))
    tf_l22 = tb_l22.text_frame
    tf_l22.word_wrap = True
    p = tf_l22.paragraphs[0]
    p.text = "📜 實戰提示詞 06：全自動推播咒語"
    p.font.size = Pt(20)
    p.font.bold = True
    p.font.color.rgb = COLOR_PRIMARY_FOREST
    p.font.name = FONT_FAMILY
    p.space_after = Pt(8)

    prompt_box22 = [
        "💡 可直接反白複製使用：",
        "我已在 GitHub 建立好公開儲存庫：https://github.com/stjedu/www，",
        "請幫我執行以下自動化連續動作鏈：",
        "1. 將今日的「基本電學與電子學 BJT 小訊號放大題庫與標準電路圖」做成個人教學入口網站 (index.html 與 electronics_bjt.html)。",
        "2. 將網頁、circuits/ 電路圖與 PPTX 投影片 Commit 並推播到 main 分支。",
        "3. 啟用 GitHub Pages 靜態伺服器！",
        "4. 回報我最終公開上線的個人教學網站網址！"
    ]
    for line in prompt_box22:
        pl = tf_l22.add_paragraph()
        pl.text = line
        if "💡" in line:
            pl.font.size = Pt(12.5)
            pl.font.bold = True
            pl.font.color.rgb = COLOR_CARAMEL
        else:
            pl.font.size = Pt(12.5)
            pl.font.color.rgb = COLOR_TEXT_MAIN
        pl.font.name = FONT_FAMILY
        pl.space_after = Pt(5)

    add_card(s23, Inches(6.9), Inches(1.75), Inches(5.6), Inches(4.9), border_color=COLOR_CARAMEL)
    tb_r22 = s23.shapes.add_textbox(Inches(7.2), Inches(2.05), Inches(5.0), Inches(4.3))
    tf_r22 = tb_r22.text_frame
    tf_r22.word_wrap = True
    p = tf_r22.paragraphs[0]
    p.text = "🎉 成果大功告成：六大教學網頁隨點隨看！"
    p.font.size = Pt(19)
    p.font.bold = True
    p.font.color.rgb = COLOR_CARAMEL
    p.font.name = FONT_FAMILY
    p.space_after = Pt(10)

    final_achieve = [
        ("🌐 研習主入口與電學題庫：", "👉 https://stjedu.github.io/www/"),
        ("📐 電子學 BJT 小訊號 rπ 專頁：", "👉 https://stjedu.github.io/www/electronics_bjt.html"),
        ("🔬 示波器與丙級檢定模擬器：", "👉 https://stjedu.github.io/www/instrument.html"),
        ("📷 YOLO 視覺與 ESP32 物聯網：", "👉 yolo_webcam_course.html  |  esp32_mqtt_course.html")
    ]
    for title, desc in final_achieve:
        pt = tf_r22.add_paragraph()
        pt.text = title
        pt.font.size = Pt(14)
        pt.font.bold = True
        pt.font.color.rgb = COLOR_PRIMARY_FOREST
        pt.font.name = FONT_FAMILY
        pd = tf_r22.add_paragraph()
        pd.text = desc
        pd.font.size = Pt(12.5)
        pd.font.color.rgb = COLOR_TEXT_MAIN
        pd.font.name = FONT_FAMILY
        pd.space_after = Pt(6)

    # ==========================================
    # SLIDE 24: 高階實戰案例展示 (實機截圖)
    # ==========================================
    s24 = prs.slides.add_slide(blank_layout)
    add_slide_scaffolding(s24, "【高階實戰案例】虛擬電子儀表 ✕ 丙級試題二模擬系統", "單元六：高階實戰特輯", 24)

    add_card(s24, Inches(0.8), Inches(1.65), Inches(5.6), Inches(5.3))
    tb_l = s24.shapes.add_textbox(Inches(1.05), Inches(1.85), Inches(5.1), Inches(4.9))
    tf_l = tb_l.text_frame
    tf_l.word_wrap = True
    p = tf_l.paragraphs[0]
    p.text = "🎯 純前端單檔教學模擬器 (instrument.html)"
    p.font.size = Pt(18)
    p.font.bold = True
    p.font.color.rgb = COLOR_PRIMARY_FOREST
    p.font.name = FONT_FAMILY
    p.space_after = Pt(10)

    points24 = [
        "⚡ 整合 GDS-2102E 示波器 ＆ MFG-2110 任意波形信號產生器",
        "🔌 一鍵連線測試：自動接線、通電、開輸出並 Autoset 展開最佳波形",
        "🔒 靜態觸發鎖定：徹底解決高頻正弦波滾動問題，穩定清晰",
        "🎬 10 步驟全自動教學導航：每 3.5s 計時自動模擬實機按鍵動作",
        "📝 台灣工業電子丙級術科試題二：12 題組電路模擬與自動評分扣分",
        "💡 零外部相依：純 SVG 手繪面板 ✕ Canvas 繪圖，完全離線可用"
    ]
    for pt_text in points24:
        pa = tf_l.add_paragraph()
        pa.text = pt_text
        pa.font.size = Pt(13.5)
        pa.font.color.rgb = COLOR_TEXT_MAIN
        pa.font.name = FONT_FAMILY
        pa.space_after = Pt(8)

    add_card(s24, Inches(6.6), Inches(1.65), Inches(5.9), Inches(5.3))
    if os.path.exists(ch1_img):
        s24.shapes.add_picture(ch1_img, Inches(6.75), Inches(1.85), width=Inches(5.6))
    cap = s24.shapes.add_textbox(Inches(6.75), Inches(5.8), Inches(5.6), Inches(0.9))
    tf_cap = cap.text_frame
    tf_cap.word_wrap = True
    p_cap = tf_cap.paragraphs[0]
    p_cap.text = "▲ 實機操作：信號產生器直連示波器 CH1，鎖定 Trig'd 靜態波形與 18px 黃色大字體 Measure 讀數"
    p_cap.font.size = Pt(11.5)
    p_cap.font.bold = True
    p_cap.font.color.rgb = COLOR_CARAMEL
    p_cap.font.name = FONT_FAMILY

    # ==========================================
    # SLIDE 25: 原廠手冊三大物理依據與自動驗收
    # ==========================================
    s25 = prs.slides.add_slide(blank_layout)
    add_slide_scaffolding(s25, "工程級教具的底層靈魂：依據三大官方資料手冊", "單元六：高階實戰特輯", 25)

    add_card(s25, Inches(0.8), Inches(1.65), Inches(5.6), Inches(5.3))
    tb_m = s25.shapes.add_textbox(Inches(1.05), Inches(1.85), Inches(5.1), Inches(4.9))
    tf_m = tb_m.text_frame
    tf_m.word_wrap = True
    p = tf_m.paragraphs[0]
    p.text = "📚 三大官方標準與物理模型"
    p.font.size = Pt(18)
    p.font.bold = True
    p.font.color.rgb = COLOR_PRIMARY_FOREST
    p.font.name = FONT_FAMILY
    p.space_after = Pt(10)

    manuals = [
        ("1. GW INSTEK GDS-2000E 使用手冊", "垂直 1-2-5 序列、時基擴展、2Vpp 1kHz 校驗方波暫態過沖公式：ΔV(t)=c×0.8×exp(-t/40µs)"),
        ("2. GW INSTEK MFG-2000 原廠手冊", "輸出阻抗 50Ω 負載效應模型：接示波器(1MΩ)端電壓放大 2 倍(設 3V 實出 6V)，High Z 則為 1 倍"),
        ("3. 工業電子丙級試題二規範 (02800-960302)", "正弦波 1kHz 3Vpp 50Ω、探棒開關 1:1、DC 耦合；12 題組 ±10% 容許度與零點偏移扣分規準"),
        ("4. 全系統自動驗收腳本 (runAutoAudit)", "內建 10 項驗收指標，單檔離線開立即可自我檢查合格！")
    ]
    for title, desc in manuals:
        pm_t = tf_m.add_paragraph()
        pm_t.text = f"🔹 {title}"
        pm_t.font.size = Pt(13.5)
        pm_t.font.bold = True
        pm_t.font.color.rgb = COLOR_CARAMEL
        pm_t.font.name = FONT_FAMILY

        pm_d = tf_m.add_paragraph()
        pm_d.text = f"   {desc}"
        pm_d.font.size = Pt(11.5)
        pm_d.font.color.rgb = COLOR_TEXT_MAIN
        pm_d.font.name = FONT_FAMILY
        pm_d.space_after = Pt(6)

    add_card(s25, Inches(6.6), Inches(1.65), Inches(5.9), Inches(5.3))
    if os.path.exists(tab_img):
        s25.shapes.add_picture(tab_img, Inches(6.75), Inches(1.85), width=Inches(5.6))
    cap2 = s25.shapes.add_textbox(Inches(6.75), Inches(5.8), Inches(5.6), Inches(0.9))
    tf_cap2 = cap2.text_frame
    tf_cap2.word_wrap = True
    p_cap2 = tf_cap2.paragraphs[0]
    p_cap2.text = "▲ 網頁內建【提示語與教學手冊】分頁：附帶一鍵複製提示語與 10 項標準即時自動稽核報告"
    p_cap2.font.size = Pt(11.5)
    p_cap2.font.bold = True
    p_cap2.font.color.rgb = COLOR_CARAMEL
    p_cap2.font.name = FONT_FAMILY

    # ==========================================
    # SLIDE 26: 完整提示語 Master Prompt (上篇) — 雙欄結構化大字版 (由 8.5pt 升級至 11.5~13pt)
    # ==========================================
    s26 = prs.slides.add_slide(blank_layout)
    add_slide_scaffolding(s26, "📌【工程級 Master Prompt (上篇)】角色定位與三大手冊物理規範", "單元六：高階實戰特輯", 26)

    # 左欄：[ROLE] + [TASK] + [WORKFLOW]
    add_card(s26, Inches(0.8), Inches(1.65), Inches(5.6), Inches(5.3), bg_color=COLOR_DARK_BOX, border_color=COLOR_AMBER)
    tb_26l = s26.shapes.add_textbox(Inches(1.05), Inches(1.85), Inches(5.1), Inches(4.9))
    tf_26l = tb_26l.text_frame
    tf_26l.word_wrap = True

    p = tf_26l.paragraphs[0]
    p.text = "🎯 [ROLE] & [TASK] 角色與系統架構"
    p.font.size = Pt(15)
    p.font.bold = True
    p.font.color.rgb = COLOR_AMBER
    p.font.name = FONT_FAMILY
    p.space_after = Pt(8)

    left_blocks_26 = [
        ("【專業角色設定】", "你是一位精通「電子量測儀器硬體模擬」與「台灣工業電子丙級術科規範」的資深全端工程師兼金牌實習指導教師。"),
        ("【核心開發任務】", "製作單一檔案、高擬真互動模擬系統 (instrument.html)：\n1. 固緯 GDS-2102E 數位示波器 ＆ MFG-2110 信號產生器\n2. 工業電子丙級試題二（12 題組電路量測與自動評分）"),
        ("【三階段工作流 WORKFLOW】", "1. 解析三份官方手冊之規格、尺寸、公式與扣分標準\n2. 模組化開發：CSS、SVG 向量面板、Canvas 波形全部內嵌於單一 HTML\n3. 跨裝置 RWD 測試 (375/768/1440px) 並通過 runAutoAudit() 自動驗收")
    ]
    for head, body in left_blocks_26:
        ph = tf_26l.add_paragraph()
        ph.text = head
        ph.font.size = Pt(13)
        ph.font.bold = True
        ph.font.color.rgb = RGBColor(125, 211, 252)
        ph.font.name = FONT_FAMILY
        pb = tf_26l.add_paragraph()
        pb.text = body
        pb.font.size = Pt(11.5)
        pb.font.color.rgb = COLOR_CODE_TEXT
        pb.font.name = FONT_FAMILY
        pb.space_after = Pt(8)
        pb.line_spacing = 1.25

    # 右欄：[MANUAL_REFERENCES]
    add_card(s26, Inches(6.7), Inches(1.65), Inches(5.8), Inches(5.3), bg_color=COLOR_DARK_BOX, border_color=COLOR_AMBER)
    tb_26r = s26.shapes.add_textbox(Inches(6.95), Inches(1.85), Inches(5.3), Inches(4.9))
    tf_26r = tb_26r.text_frame
    tf_26r.word_wrap = True

    p = tf_26r.paragraphs[0]
    p.text = "📘 [MANUAL_REFERENCES] 三大原廠手冊物理規範"
    p.font.size = Pt(15)
    p.font.bold = True
    p.font.color.rgb = COLOR_AMBER
    p.font.name = FONT_FAMILY
    p.space_after = Pt(8)

    right_blocks_26 = [
        ("1. 《GW INSTEK GDS-2000E User Manual》", "• 垂直 1mV~10V/div、時基 1ns~100s/div (1-2-5 序列)\n• 探棒校驗 2Vpp 1kHz 方波暫態過沖：ΔV(t) = c×0.8×exp(-t/40µs)\n• 800×536 螢幕，10×8 格線，上下狀態資訊列"),
        ("2. 《GW INSTEK MFG-2000 User Manual》", "• 輸出阻抗 50Ω：設 50Ω 接示波器(1MΩ)產生 2 倍電壓效應(設3V實出6V)；設 High Z 為 1 倍\n• 有效值換算：正弦波 Vpp = 2√2 Vrms；三角波 Vpp = 2√3 Vrms"),
        ("3. 《工業電子丙級術科試題二 (02800-960302)》", "• 輸入設定：正弦波 1kHz、3.000Vpp、負載 50Ω、探棒 1:1、DC 耦合\n• 評分規準：12 題組理論電壓 ±10% 容許度、探棒非 1:1 扣 10 分、零位偏離 >0.5 格扣 5 分")
    ]
    for head, body in right_blocks_26:
        ph = tf_26r.add_paragraph()
        ph.text = head
        ph.font.size = Pt(13)
        ph.font.bold = True
        ph.font.color.rgb = RGBColor(125, 211, 252)
        ph.font.name = FONT_FAMILY
        pb = tf_26r.add_paragraph()
        pb.text = body
        pb.font.size = Pt(11.5)
        pb.font.color.rgb = COLOR_CODE_TEXT
        pb.font.name = FONT_FAMILY
        pb.space_after = Pt(8)
        pb.line_spacing = 1.25

    # ==========================================
    # SLIDE 27: 完整提示語 Master Prompt (下篇) — 雙欄結構化大字版
    # ==========================================
    s27 = prs.slides.add_slide(blank_layout)
    add_slide_scaffolding(s27, "📌【工程級 Master Prompt (下篇)】硬性約束、引擎演算法與自動驗收", "單元六：高階實戰特輯", 27)

    # 左欄：[HARD_CONSTRAINTS] + [OSCILLOSCOPE_ENGINE_SPEC]
    add_card(s27, Inches(0.8), Inches(1.65), Inches(5.6), Inches(5.3), bg_color=COLOR_DARK_BOX, border_color=COLOR_AMBER)
    tb_27l = s27.shapes.add_textbox(Inches(1.05), Inches(1.85), Inches(5.1), Inches(4.9))
    tf_27l = tb_27l.text_frame
    tf_27l.word_wrap = True

    p = tf_27l.paragraphs[0]
    p.text = "⚙️ [CONSTRAINTS] & [ENGINE_SPEC] 核心引擎規範"
    p.font.size = Pt(15)
    p.font.bold = True
    p.font.color.rgb = COLOR_AMBER
    p.font.name = FONT_FAMILY
    p.space_after = Pt(8)

    left_blocks_27 = [
        ("【硬性約束 HARD_CONSTRAINTS】", "• 單一獨立檔案 instrument.html，零外部相依，完全離線可用\n• 台灣專業術語（探棒、觸發、時基、耦合、峰對峰值）\n• 頂部提供【🔌 一鍵測試：信號源 ➔ CH1】自動接線並展開波形"),
        ("【靜態觸發鎖定演算法 Stationary Trigger Lock】", "• 當狀態為 Trig'd 綠燈時，時基精確鎖定正向過零相位：\n  t_lock = arcsin((Vtrig - Vavg)/(Vpp/2)) / (2πf)\n• 波形完全靜止清晰；未觸發 (Auto) 時才水平漂移"),
        ("【Autoset 與大字體量測視窗】", "• Measure 視窗加大至 280×92px，採用 bold 18px 螢光黃顯示 Vpp 與 Freq\n• Autoset 自動將振幅限制在 3~6.5 格，時基展開 2~4 個週期")
    ]
    for head, body in left_blocks_27:
        ph = tf_27l.add_paragraph()
        ph.text = head
        ph.font.size = Pt(13)
        ph.font.bold = True
        ph.font.color.rgb = RGBColor(125, 211, 252)
        ph.font.name = FONT_FAMILY
        pb = tf_27l.add_paragraph()
        pb.text = body
        pb.font.size = Pt(11.5)
        pb.font.color.rgb = COLOR_CODE_TEXT
        pb.font.name = FONT_FAMILY
        pb.space_after = Pt(8)
        pb.line_spacing = 1.25

    # 右欄：[GUIDE_ENGINE_SPEC] + [ACCEPTANCE_CRITERIA]
    add_card(s27, Inches(6.7), Inches(1.65), Inches(5.8), Inches(5.3), bg_color=COLOR_DARK_BOX, border_color=COLOR_AMBER)
    tb_27r = s27.shapes.add_textbox(Inches(6.95), Inches(1.85), Inches(5.3), Inches(4.9))
    tf_27r = tb_27r.text_frame
    tf_27r.word_wrap = True

    p = tf_27r.paragraphs[0]
    p.text = "✅ [GUIDE_SPEC] & [ACCEPTANCE] 導航與 10 項驗收"
    p.font.size = Pt(15)
    p.font.bold = True
    p.font.color.rgb = COLOR_AMBER
    p.font.name = FONT_FAMILY
    p.space_after = Pt(8)

    right_blocks_27 = [
        ("【雙指引 10 步驟自動教學排程】", "• 示波器與信號產生器各配一組 10 步教學指引卡（每 3.5 秒自動推進）\n• 步驟推進時自動調用硬體 API（Default、接線、Probe 10X、Autoset），實機對應按鈕同步橘光呼吸高亮"),
        ("【自動化驗收腳本 window.runAutoAudit() 十項指標】",
         "01. 離線單檔無外部依賴     02. 探棒補償尖角/圓角判定\n"
         "03. 50Ω負載2倍/HiZ 1倍效應  04. Autoset 波形 3~6.5 格\n"
         "05. 觸發準位 Trig'd 穩定鎖定  06. 丙級 12 題組數據完備\n"
         "07. 觸控標靶尺寸 ≥ 40px      08. 探棒開關與讀數倍率換算\n"
         "09. 雙指引卡 10 步自動引擎    10. Measure 讀數 18px 高清晰"),
        ("💡 完整提示語可直接至網頁複製：", "👉 https://stjedu.github.io/www/instrument.html（點擊上方【提示語與教學手冊】分頁即可一鍵複製）")
    ]
    for head, body in right_blocks_27:
        ph = tf_27r.add_paragraph()
        ph.text = head
        ph.font.size = Pt(13)
        ph.font.bold = True
        ph.font.color.rgb = RGBColor(125, 211, 252)
        ph.font.name = FONT_FAMILY
        pb = tf_27r.add_paragraph()
        pb.text = body
        pb.font.size = Pt(11.5)
        pb.font.color.rgb = COLOR_CODE_TEXT
        pb.font.name = FONT_FAMILY
        pb.space_after = Pt(8)
        pb.line_spacing = 1.25

    # ==========================================
    # SLIDE 28: 研習總結與教師思維升級（整合收官頁）
    # ==========================================
    s28 = prs.slides.add_slide(blank_layout)
    add_slide_scaffolding(s28, "研習總結：今日解鎖技能 ✕ 教師的教具開發思維升級", "研習總結與未來展望", 28)

    add_card(s28, Inches(0.8), Inches(1.75), Inches(5.6), Inches(4.9))
    tb_l28 = s28.shapes.add_textbox(Inches(1.1), Inches(2.0), Inches(5.0), Inches(4.4))
    tf_l28 = tb_l28.text_frame
    tf_l28.word_wrap = True
    p = tf_l28.paragraphs[0]
    p.text = "🎉 今日已解鎖之七大實戰技能"
    p.font.size = Pt(20)
    p.font.bold = True
    p.font.color.rgb = COLOR_PRIMARY_FOREST
    p.font.name = FONT_FAMILY
    p.space_after = Pt(12)

    achieves = [
        "✅ 完成 Antigravity 工作區建置與沙盒安全檢測",
        "✅ 掌握零代碼一鍵自動安裝 Python / 繪圖 / Office 套件",
        "✅ 掌握基本電學戴維寧定理之結構化命題與標準電路圖生成",
        "✅ 掌握進階電子學 BJT 共射極小訊號 rπ 模型推導與互動網頁",
        "✅ 掌握 MCP 磨課師自動化：30 秒人機接力與突發急救包",
        "✅ 實戰 Excel 自動算技高加權成績、粉紅標記與學情分析圖表",
        "✅ 掌握 GitHub Pages 發布與工業電子丙級虛擬儀表開發心法"
    ]
    for a in achieves:
        pa = tf_l28.add_paragraph()
        pa.text = a
        pa.font.size = Pt(13)
        pa.font.bold = True
        pa.font.color.rgb = COLOR_TEXT_MAIN
        pa.font.name = FONT_FAMILY
        pa.space_after = Pt(7)

    add_card(s28, Inches(6.9), Inches(1.75), Inches(5.6), Inches(4.9), border_color=COLOR_CARAMEL)
    tb_r28 = s28.shapes.add_textbox(Inches(7.2), Inches(2.0), Inches(5.0), Inches(4.4))
    tf_r28 = tb_r28.text_frame
    tf_r28.word_wrap = True
    p = tf_r28.paragraphs[0]
    p.text = "🚀 三大關鍵思維升級 ＆ 結語金句"
    p.font.size = Pt(20)
    p.font.bold = True
    p.font.color.rgb = COLOR_CARAMEL
    p.font.name = FONT_FAMILY
    p.space_after = Pt(10)

    mindsets = [
        ("🧩 模組化規格思維", "AI 不是魔術，AI 需要精準的「規格說明書 (CRIS)」。懂得定義功能邊界，AI 就能打造工程級作品。"),
        ("🔬 公式即代碼規範", "把課本與原廠手冊公式（如 rπ 模型、50Ω 負載效應）寫進 Prompt，AI 就能做出物理級高擬真教具。"),
        ("🌟 人人皆能打造專屬實驗室", "不需等待廠商開發！只要您具備學科專業＋下對 Prompt，每門課都能擁有專屬線上互動實驗室！")
    ]
    for mt, md in mindsets:
        pa_t = tf_r28.add_paragraph()
        pa_t.text = mt
        pa_t.font.size = Pt(13.5)
        pa_t.font.bold = True
        pa_t.font.color.rgb = COLOR_PRIMARY_FOREST
        pa_t.font.name = FONT_FAMILY
        pa_d = tf_r28.add_paragraph()
        pa_d.text = md
        pa_d.font.size = Pt(12)
        pa_d.font.color.rgb = COLOR_TEXT_MAIN
        pa_d.font.name = FONT_FAMILY
        pa_d.space_after = Pt(6)

    pq = tf_r28.add_paragraph()
    pq.text = "「AI 不會取代老師，但善用代理 AI 開發教具的老師，將徹底點亮學生的學習視野！」"
    pq.font.size = Pt(13)
    pq.font.bold = True
    pq.font.color.rgb = COLOR_CARAMEL
    pq.font.name = FONT_FAMILY
    pq.space_before = Pt(6)

    out_slides = os.path.join(ROOT_DIR, "slides.pptx")
    out_named = os.path.join(ROOT_DIR, "第一場_Antigravity代理AI教學實戰_大字動畫版_GitHub發布版.pptx")
    prs.save(out_slides)
    shutil.copy2(out_slides, out_named)
    print(f"[OK] Main deck saved: {out_slides} & {out_named} ({len(prs.slides)} slides)")


def build_yolo_deck():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    C_NAVY = RGBColor(20, 35, 60)
    C_TEAL = RGBColor(0, 140, 160)
    C_ORANGE = RGBColor(235, 110, 40)
    C_BG_LIGHT = RGBColor(248, 250, 252)
    C_CARD_BG = RGBColor(255, 255, 255)
    C_BORDER = RGBColor(220, 226, 235)
    C_TEXT_DARK = RGBColor(30, 41, 59)
    C_TEXT_MUTED = RGBColor(100, 116, 139)
    C_CODE_BG = RGBColor(241, 245, 249)
    C_CODE_TEXT = RGBColor(15, 23, 42)

    TOTAL_SLIDES = 14
    blank_layout = prs.slide_layouts[6]

    def apply_transition(slide):
        trans_xml = parse_xml(f'<p:transition {nsdecls("p")} spd="med"><p:fade/></p:transition>')
        slide._element.append(trans_xml)

    def add_header(slide, title, slide_idx, category="WebCam AI 視覺與 YOLO 專題實作課程"):
        apply_transition(slide)
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
        bg.fill.solid()
        bg.fill.fore_color.rgb = C_BG_LIGHT
        bg.line.fill.background()

        header_box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(1.15))
        header_box.fill.solid()
        header_box.fill.fore_color.rgb = C_NAVY
        header_box.line.fill.background()

        accent_bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(1.15), Inches(4.5), Inches(0.05))
        accent_bar.fill.solid()
        accent_bar.fill.fore_color.rgb = C_TEAL
        accent_bar.line.fill.background()

        cat_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.12), Inches(8.5), Inches(0.3))
        tf_c = cat_box.text_frame
        tf_c.word_wrap = True
        p_c = tf_c.paragraphs[0]
        p_c.text = f"📷 {category.upper()}"
        p_c.font.size = Pt(11.5)
        p_c.font.bold = True
        p_c.font.name = "微軟正黑體"
        p_c.font.color.rgb = C_TEAL

        page_box = slide.shapes.add_textbox(Inches(10.0), Inches(0.12), Inches(2.5), Inches(0.3))
        tf_p = page_box.text_frame
        p_p = tf_p.paragraphs[0]
        p_p.alignment = PP_ALIGN.RIGHT
        p_p.text = f"第 {slide_idx:02d} / {TOTAL_SLIDES:02d} 頁"
        p_p.font.size = Pt(12)
        p_p.font.bold = True
        p_p.font.name = "微軟正黑體"
        p_p.font.color.rgb = RGBColor(180, 200, 225)

        tb = slide.shapes.add_textbox(Inches(0.8), Inches(0.40), Inches(11.7), Inches(0.65))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = title
        p.font.size = Pt(24)
        p.font.bold = True
        p.font.name = "微軟正黑體"
        p.font.color.rgb = RGBColor(255, 255, 255)

        footer_box = slide.shapes.add_textbox(Inches(0.8), Inches(7.12), Inches(11.7), Inches(0.3))
        tf_f = footer_box.text_frame
        pf = tf_f.paragraphs[0]
        pf.text = "🌐 互動課程網頁：https://stjedu.github.io/www/yolo_webcam_course.html   |   主講：簡樹桐 老師"
        pf.font.size = Pt(10.5)
        pf.font.color.rgb = C_TEXT_MUTED
        pf.font.name = "微軟正黑體"

    def add_card(slide, left, top, width, height, title, content_items, bg_color=C_CARD_BG, border_color=C_BORDER, title_color=C_NAVY, top_bar_color=C_TEAL):
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
        card.fill.solid()
        card.fill.fore_color.rgb = bg_color
        card.line.color.rgb = border_color
        card.line.width = Pt(1.5)

        bar = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, Inches(0.1))
        bar.fill.solid()
        bar.fill.fore_color.rgb = top_bar_color
        bar.line.fill.background()

        tb = slide.shapes.add_textbox(left + Inches(0.25), top + Inches(0.2), width - Inches(0.5), height - Inches(0.35))
        tf = tb.text_frame
        tf.word_wrap = True

        if title:
            p_title = tf.paragraphs[0]
            p_title.text = title
            p_title.font.size = Pt(17.5)
            p_title.font.bold = True
            p_title.font.name = "微軟正黑體"
            p_title.font.color.rgb = title_color
            p_title.space_after = Pt(10)

        for idx, item in enumerate(content_items):
            p = tf.add_paragraph() if (title or idx > 0) else tf.paragraphs[0]
            p.space_after = Pt(7)
            p.line_spacing = 1.25
            if isinstance(item, tuple):
                bold_prefix, normal_text = item
                run1 = p.add_run()
                run1.text = bold_prefix + " "
                run1.font.bold = True
                run1.font.size = Pt(13.5)
                run1.font.name = "微軟正黑體"
                run1.font.color.rgb = C_TEXT_DARK

                run2 = p.add_run()
                run2.text = normal_text
                run2.font.bold = False
                run2.font.size = Pt(13.5)
                run2.font.name = "微軟正黑體"
                run2.font.color.rgb = C_TEXT_MUTED
            else:
                run = p.add_run()
                run.text = item
                run.font.size = Pt(13.5)
                run.font.name = "微軟正黑體"
                run.font.color.rgb = C_TEXT_DARK
        return card

    def add_prompt_card(slide, left, top, width, height, prompt_title, prompt_text, tips=""):
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
        card.fill.solid()
        card.fill.fore_color.rgb = C_CODE_BG
        card.line.color.rgb = C_TEAL
        card.line.width = Pt(2)

        tb = slide.shapes.add_textbox(left + Inches(0.25), top + Inches(0.16), width - Inches(0.5), height - Inches(0.32))
        tf = tb.text_frame
        tf.word_wrap = True

        p0 = tf.paragraphs[0]
        p0.text = f"💡 實戰提示詞 (Prompt)：{prompt_title}   [可直接反白複製]"
        p0.font.size = Pt(15)
        p0.font.bold = True
        p0.font.name = "微軟正黑體"
        p0.font.color.rgb = C_ORANGE
        p0.space_after = Pt(6)

        p1 = tf.add_paragraph()
        run1 = p1.add_run()
        run1.text = f"「{prompt_text}」"
        run1.font.size = Pt(12.5)
        run1.font.name = "Consolas"
        run1.font.color.rgb = C_CODE_TEXT
        p1.space_after = Pt(6)
        p1.line_spacing = 1.2

        if tips:
            p2 = tf.add_paragraph()
            run2 = p2.add_run()
            run2.text = f"📌 教師引導技巧：{tips}"
            run2.font.size = Pt(11.5)
            run2.font.bold = True
            run2.font.name = "微軟正黑體"
            run2.font.color.rgb = C_TEAL

    # SLIDE 1: 封面
    s1 = prs.slides.add_slide(blank_layout)
    apply_transition(s1)
    bg1 = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.5))
    bg1.fill.solid()
    bg1.fill.fore_color.rgb = C_NAVY
    bg1.line.fill.background()

    c1 = s1.shapes.add_shape(MSO_SHAPE.OVAL, Inches(9.6), Inches(-1.2), Inches(5.0), Inches(5.0))
    c1.fill.solid()
    c1.fill.fore_color.rgb = RGBColor(28, 52, 88)
    c1.line.fill.background()

    tb = s1.shapes.add_textbox(Inches(1.2), Inches(1.7), Inches(11.0), Inches(4.2))
    tf = tb.text_frame
    tf.word_wrap = True

    p0 = tf.paragraphs[0]
    p0.text = "📷 AI 視覺與電腦視覺專題實作系列"
    p0.font.size = Pt(17)
    p0.font.bold = True
    p0.font.name = "微軟正黑體"
    p0.font.color.rgb = C_ORANGE
    p0.space_after = Pt(12)

    p = tf.add_paragraph()
    p.text = "WebCam AI 視覺與 YOLO 專題應用實作"
    p.font.size = Pt(40)
    p.font.bold = True
    p.font.name = "微軟正黑體"
    p.font.color.rgb = RGBColor(255, 255, 255)
    p.space_after = Pt(14)

    p2 = tf.add_paragraph()
    p2.text = "AI 輔助程式開發 ‧ Prompt 提示詞工程 ‧ 四宮格多分割監控 ‧ 跌倒與手勢控制"
    p2.font.size = Pt(20)
    p2.font.name = "微軟正黑體"
    p2.font.color.rgb = C_TEAL
    p2.space_after = Pt(28)

    p3 = tf.add_paragraph()
    p3.text = "主講：簡樹桐 老師 ｜ 適用：高中職專題製作 ‧ 教師研習 ‧ AI 應用實作\n🌐 線上互動教材：https://stjedu.github.io/www/yolo_webcam_course.html"
    p3.font.size = Pt(14.5)
    p3.font.name = "微軟正黑體"
    p3.font.color.rgb = RGBColor(200, 215, 235)

    # SLIDE 2
    s2 = prs.slides.add_slide(blank_layout)
    add_header(s2, "課程全景：從手刻程式到「AI 提示詞導向」的專題教學革新", 2)
    add_card(s2, Inches(0.8), Inches(1.45), Inches(5.6), Inches(5.45),
             "🚀 專題教學新思維 (Paradigm Shift)",
             [
                 ("告別一行行刻語法：", "學生最大門檻不是邏輯，而是卡在語法報錯與挫折感。透過 AI，專案開發速度由『半年做 1 個』提升為『半年做 10 個』。"),
                 ("提問與規格能力才是核心：", "如何把模糊的想法（例如：我要做跌倒報警），拆解成精確的『相機輸入、演算規則、輸出反饋』，是專題成敗關鍵。"),
                 ("反覆對話與微調：", "寫出不滿意的程式不要自己從頭改，叫 AI 改！只要清楚描述新需求與錯誤訊息，AI 即刻修正。"),
                 ("商品化實作體驗：", "不再只是紙上談兵，一堂課內就能完成具備商業雛型的完整監控或互動視覺系統。")
             ], border_color=C_TEAL, top_bar_color=C_TEAL)
    add_card(s2, Inches(6.9), Inches(1.45), Inches(5.6), Inches(5.45),
             "📚 四大單元實作架構 (Course Roadmap)",
             [
                 ("第 01 節課：基礎測試與串流控制", "OpenCV 環境架構、WebCam 即時串流、自訂熱鍵拍照與定時縮時攝影。"),
                 ("第 02 節課：Prompt 工程與畫面校正", "提問心法、MCP 自動自我檢核模式、1:1 等比例畫面裁切防變形與算力解析度哲學。"),
                 ("第 03 節課：多格分割視窗與縮放演算法", "四宮格 (2x2) 拼接、雙攝影機 (Dual Cam) 整合、USB 頻寬管理與點擊互換畫面。"),
                 ("第 04 節課：YOLO 骨架追蹤與智慧應用", "YOLO Pose 17 關鍵點、畫面中心主控者辨識、舉手拍照、平舉錄影與多人跌倒紅框警報。")
             ], border_color=C_NAVY, top_bar_color=C_ORANGE)

    # SLIDE 3
    s3 = prs.slides.add_slide(blank_layout)
    add_header(s3, "【第 01 節課】AI 視覺環境建置與 WebCam 基礎測試", 3)
    add_card(s3, Inches(0.8), Inches(1.45), Inches(5.6), Inches(5.45),
             "🎯 學習目標與技能樹",
             [
                 ("影像串流捕獲：", "理解 Python 與 OpenCV (`cv2.VideoCapture`) 讀取 WebCam 的基本原理。"),
                 ("事件驅動與鍵盤監聽：", "掌握 `cv2.waitKey()` 鍵盤輸入邏輯，實作即時熱鍵互動。"),
                 ("影像格式與存檔機制：", "理解幀率 (FPS) 與圖片寫入 (`cv2.imwrite`)，自動建立時間戳記檔名。"),
                 ("自主探索與進度分流：", "提供標準網頁教學範例，讓學有餘力的同學可超前自主探索進階功能。")
             ], top_bar_color=C_TEAL)
    add_card(s3, Inches(6.9), Inches(1.45), Inches(5.6), Inches(5.45),
             "⚙️ 基礎功能實作清單",
             [
                 ("1. 攝影機開啟與串流監聽：", "檢測相機索引 0 是否正常啟動，若失敗給予友善提示。"),
                 ("2. 拍照熱鍵 (s / Space)：", "按下快捷鍵即時擷取當前 Frame，自動以年月日時分秒存檔。"),
                 ("3. 結束熱鍵 (q / ESC)：", "優雅關閉串流並釋放硬體資源 (`cap.release()`)。"),
                 ("4. 自動定時拍攝 (Timelapse)：", "設定時間間隔（如每 5 秒拍一張），實作簡易縮時攝影。"),
                 ("5. 延遲除錯：", "探討解析度對串流延遲的影響，奠定流暢度優化基礎。")
             ], top_bar_color=C_ORANGE)

    # SLIDE 4
    s4 = prs.slides.add_slide(blank_layout)
    add_header(s4, "【第 01 節課】功能實作提示語 (Prompt) 與教師引導", 4)
    add_prompt_card(s4, Inches(0.8), Inches(1.45), Inches(11.7), Inches(2.65),
                    "基礎 WebCam 串流與熱鍵拍照控制",
                    "請使用 Python 和 OpenCV 編寫一個 WebCam 視窗程式。\n"
                    "功能需求：\n"
                    "1. 開啟預設攝影機（索引 0）並以 640x480 即時顯示畫面。\n"
                    "2. 在畫面上即時顯示當前 FPS 與當前時間（微軟正黑體或簡潔英文）。\n"
                    "3. 按下鍵盤 's' 鍵手動拍照，將圖片存入 'captures' 資料夾，檔名格式為 'img_YYYYMMDD_HHMMSS.jpg'，並在畫面上短暫顯示 '已拍照存檔' 綠色提示。\n"
                    "4. 按下 'q' 鍵安全釋放相機資源並關閉所有視窗。",
                    "先讓學生體驗最直接的鍵盤互動，建立成就感；若畫面延遲，引導學生觀察 FPS 數字。")
    add_prompt_card(s4, Inches(0.8), Inches(4.3), Inches(11.7), Inches(2.65),
                    "自動定時拍照（縮時攝影）與異常處理擴充",
                    "請在上述基礎上進行功能升級（請直接提供修改後的完整 Python 代碼）：\n"
                    "1. 新增自動定時拍攝功能：每隔 5 秒自動拍一張照片並存檔。\n"
                    "2. 畫面右上角顯示倒數計時秒數，提示距離下一次自動拍攝還剩幾秒。\n"
                    "3. 增加例外處理：若相機被其他軟體佔用或找不到設備，請跳出友善視窗提示，不要讓程式崩潰閃退。\n"
                    "4. 代碼請包含清晰的中文註解與變數說明。",
                    "提示學生：『程式不滿意就叫 AI 改！』讓學生觀察 AI 如何將功能累加進原程式碼。")

    # SLIDE 5
    s5 = prs.slides.add_slide(blank_layout)
    add_header(s5, "【第 02 節課】Prompt 提示詞工程與 1:1 畫面校正", 5)
    add_card(s5, Inches(0.8), Inches(1.45), Inches(5.6), Inches(5.45),
             "💡 學生不會用 AI 的真正痛點",
             [
                 ("專業詞彙不足：", "學生只會說『幫我做個相機』，AI 給出極簡範例，學生不知道如何提出具體要求。"),
                 ("缺乏情境與角色定義：", "未設定身分與技術棧，AI 容易給出無法在 Windows / Python 本地端運行的網頁版或不相容語法。"),
                 ("忽略輸入與邊界條件：", "沒有指定檔名命名法、資料夾路徑、異常狀況處理，導致程式跑一次就出錯。"),
                 ("不會追問與反饋：", "遇到錯誤就把程式碼刪除，不懂得把 Error Log 貼回給 AI 要求修復。")
             ], border_color=C_ORANGE, top_bar_color=C_ORANGE)
    add_card(s5, Inches(6.9), Inches(1.45), Inches(5.6), Inches(5.45),
             "📐 視覺核心：1:1 比例與解析度哲學",
             [
                 ("畫面變形痛點：", "很多初學者直接使用 `cv2.resize()` 強制拉滿視窗，導致人臉變寬變胖，失去真實幾何比例。"),
                 ("1:1 等比例校正：", "透過計算長寬比（Aspect Ratio），進行黑邊填充（Letterboxing）或中央裁切，確保物體幾何形狀不變形。"),
                 ("高解析度迷思：", "1080p 或 4K 畫面好看，但計算負擔呈幾何倍數上升。"),
                 ("AI 模型的算力原則：", "高解析度用於『人眼瀏覽與留存存檔』；AI 模型推論必須採用『低/中解析度 (如 640x360 或 1/4 畫面)』，才能維持高幀率！")
             ], top_bar_color=C_TEAL)

    # SLIDE 6
    s6 = prs.slides.add_slide(blank_layout)
    add_header(s6, "【第 02 節課】Prompt 工程四要素 ＆ 黃金除錯指令", 6)
    add_card(s6, Inches(0.8), Inches(1.45), Inches(5.6), Inches(2.65),
             "📋 高效 Prompt 的四層結構 (Prompt Structure)",
             [
                 ("1. 角色設定：", "明確身分（如：Python 影像處理資深工程師、高職專題導師）。"),
                 ("2. 任務與情境：", "清晰描述業務目的（如：開發具備智慧亮度警示的拍照控制系統）。"),
                 ("3. 輸入/輸出規範：", "指定解析度、熱鍵、輸出資料夾、存檔檔名格式、UI 顏色。"),
                 ("4. 限制與異常：", "相機離線保護、中文編碼處理、防變形等比例原則。")
             ], top_bar_color=C_TEAL)
    add_card(s6, Inches(6.9), Inches(1.45), Inches(5.6), Inches(2.65),
             "⚡ 影像品質自動檢核概念",
             [
                 ("亮度灰階計算：", "將 Frame 轉為灰階後計算平均像素強度（Mean Pixel Value 0~255）。"),
                 ("過暗警示 (< 50)：", "畫面上即時顯示黃色警示『環境過暗，請開燈或調整曝光』。"),
                 ("過亮警示 (> 200)：", "畫面顯示紅色警示『畫面過曝，請避開背光光源』。")
             ], top_bar_color=C_ORANGE)
    add_prompt_card(s6, Inches(0.8), Inches(4.3), Inches(11.7), Inches(2.65),
                    "簡老師強烈推薦：MCP 自我檢核黃金指令",
                    "每次開發完程式，請啟用 MCP 模式，做程式功能的自我檢核，若有問題請直接修正！",
                    "【除錯密技】教師務必傳授此句！學生遇到 Bug 時不必手足無措，讓 AI 透過工具自動讀取運行日誌、自我發現錯誤並完成修補，大幅降低課堂卡關率。")

    # SLIDE 7
    s7 = prs.slides.add_slide(blank_layout)
    add_header(s7, "【第 02 節課】功能實作提示語 (Prompt) 與教師引導", 7)
    add_prompt_card(s7, Inches(0.8), Inches(1.45), Inches(11.7), Inches(2.65),
                    "1:1 等比例畫面校正與智慧亮度檢測視窗",
                    "你是一位精通 OpenCV 演算法的 Python 工程師。請設計一個 WebCam 控制視窗：\n"
                    "1. 畫面防變形校正：不管輸入攝影機原始長寬比為何，在固定大小的視窗（如 800x800）中，自動維持 1:1 或真實長寬比，不足部分以黑色背景填充（Letterbox 演算法），嚴禁將人物畫面拉扯變形。\n"
                    "2. 亮度即時檢核：計算畫面中央區域之平均亮度（0~255）。若平均值 < 50 顯示『警告：光線過暗』；若 > 205 顯示『警告：過度曝光』；正常時顯示『光線良好』。\n"
                    "3. 支援鍵盤熱鍵操作：'s' 拍照、'c' 切換填充顏色、'q' 離開。\n"
                    "4. 請給出完整可直接執行的 Python 腳本。",
                    "引導學生站在鏡頭前拿手機手電筒照鏡頭或遮住鏡頭，親眼驗證 AI 演算法的即時檢測反應。")
    add_prompt_card(s7, Inches(0.8), Inches(4.3), Inches(11.7), Inches(2.65),
                    "推論與儲存分離：高低解析度雙流處理架構",
                    "請改寫 WebCam 程式，實作『雙解析度分流架構』：\n"
                    "1. 串流輸入維持相機最高可用畫質（如 1080p）。\n"
                    "2. 畫面顯示與後續 AI 模型輸入端，自動降採樣至 640x360（1/4 算力），確保 FPS 穩定維持在 30 幀以上。\n"
                    "3. 當使用者觸發拍照（按 's'）時，自動擷取 1080p 高畫質原始母幀進行儲存，兼顧即時流暢度與存檔清晰度。\n"
                    "4. 在視窗左上角對比顯示：【原始解析度】 vs 【顯示/運算解析度】及即時推論延遲（ms）。",
                    "這套架構是邊緣運算（如樹莓派、Jetson）與商用監控系統的核心概念，極具專題價值！")

    # SLIDE 8
    s8 = prs.slides.add_slide(blank_layout)
    add_header(s8, "【第 03 節課】多格分割視窗與畫面縮放演算法", 8)
    add_card(s8, Inches(0.8), Inches(1.45), Inches(5.6), Inches(5.45),
             "🖥️ 監控多宮格與畫面數學換算",
             [
                 ("1920x1080 螢幕佈局：", "螢幕解析度為 1920x1080，若要做 2x2 四宮格，每個宮格最佳尺寸為 960x540（寬高各除以 2）。"),
                 ("不同輸入比例縮放問題：", "當輸入訊號為 1280x720 或 4:3 鏡頭時，縮放比例為乘以 0.75 或 0.5，需透過座標補白演算法保持不失真。"),
                 ("陣列拼接 (NumPy Concatenate)：", "使用 `np.hstack()` 與 `np.vstack()` 將四個子 Frame 拼裝為完整畫面。"),
                 ("多宮格效能瓶頸：", "四個鏡頭或四個視窗同時推流，資料傳輸量暴增，常導致整體 FPS 驟降。")
             ], top_bar_color=C_TEAL)
    add_card(s8, Inches(6.9), Inches(1.45), Inches(5.6), Inches(5.45),
             "🎥 雙攝影機 (Dual WebCam) 整合實務",
             [
                 ("多鏡頭硬體管理：", "同時啟動 Camera 0 與 Camera 1。注意 USB 傳輸通道頻寬，避免因 Hub 頻寬不足導致斷流。"),
                 ("非同步緒處理：", "雙鏡頭初始化速度慢，單一線程容易造成畫面卡頓，需採用多線程 (Threading) 分別擷取。"),
                 ("焦點切換與畫面互換：", "點擊任一宮格可放大為單一全螢幕；點擊畫面 1 與畫面 2 可互相對調位置。"),
                 ("動態資源調配：", "關閉非關鍵宮格之更新，動態將系統 FPS 從 10 幀拉回至滿速 30 幀。")
             ], top_bar_color=C_ORANGE)

    # SLIDE 9
    s9 = prs.slides.add_slide(blank_layout)
    add_header(s9, "【第 03 節課】功能實作提示語 (Prompt) 與教師引導", 9)
    add_prompt_card(s9, Inches(0.8), Inches(1.45), Inches(11.7), Inches(2.65),
                    "四宮格 (2x2) 多分割監控畫面與動態切換",
                    "請使用 Python、OpenCV 與 NumPy 開發一個【四宮格監控展示系統】：\n"
                    "1. 建立一個 1920x1080 的主視窗，分割為 2x2 四個宮格（各 960x540）。\n"
                    "2. 宮格 1 顯示原色相機串流；宮格 2 顯示灰階處理畫面；宮格 3 顯示邊緣偵測 (Canny Edge)；宮格 4 顯示反相或特殊濾鏡畫面。\n"
                    "3. 支援滑鼠互動：滑鼠左鍵點擊任一宮格，該宮格瞬間切換為全螢幕單一視窗；再點一下恢復四宮格。\n"
                    "4. 按鍵盤 '1'~'4' 可獨立切換各宮格的開啟/關閉，以降低系統運算負荷並提高整體 FPS。",
                    "先用單支攝影機模擬 4 個不同影像處理通道，建立多分割與滑鼠事件概念。")
    add_prompt_card(s9, Inches(0.8), Inches(4.3), Inches(11.7), Inches(2.65),
                    "雙實體攝影機 (Dual WebCam) 即時拼接與畫面互換",
                    "請擴充上述系統為【雙實體攝影機監控系統】：\n"
                    "1. 同時偵測並開啟系統中的兩支 USB 攝影機（WebCam 0 與 WebCam 1）。\n"
                    "2. 使用多執行緒 (Threading) 同步讀取兩支鏡頭，防止因初始化或取幀等待造成視窗卡死。\n"
                    "3. 將兩支鏡頭畫面並排或放入四宮格中，確保等比例縮放不變形。\n"
                    "4. 實作『鏡頭對調功能』：按下鍵盤 'x' 鍵，鏡頭 1 與鏡頭 2 的顯示位置瞬間對調。\n"
                    "5. 提供完整的啟動檢查與防呆機制，若只接了一支相機則自動退回單鏡頭模式。",
                    "課堂可讓相鄰兩位同學互相借用相機插在同一台主機，體驗真實雙視角保全系統。")

    # SLIDE 10
    s10 = prs.slides.add_slide(blank_layout)
    add_header(s10, "【第 04 節課】YOLO 骨架追蹤與智慧應用實作", 10)
    add_card(s10, Inches(0.8), Inches(1.45), Inches(5.6), Inches(5.45),
             "🦴 YOLO Pose 骨架估計原理",
             [
                 ("17 個人體關鍵點 (Keypoints)：", "YOLOv8/YOLO11 Pose 能即時定位人體：鼻子、眼睛、耳朵、雙肩、雙肘、雙腕、雙髖、雙膝、雙踝。"),
                 ("坐標幾何分析：", "不需複雜的神經網絡重訓，直接透過關節 (X, Y) 座標的相對高低與幾何夾角，即可精準判斷動作。"),
                 ("手勢拍照判定邏輯：", "舉右手：右手腕 Y 座標 < 右肩膀 Y 座標（影像座標系 Y 軸朝下，數值越小代表位置越高）。"),
                 ("雙手平舉錄影邏輯：", "雙手腕高度接近雙肩膀（誤差在門檻值內），且雙手橫向張開超過肩膀寬度。")
             ], top_bar_color=C_TEAL)
    add_card(s10, Inches(6.9), Inches(1.45), Inches(5.6), Inches(5.45),
             "🚨 多人場景與長者跌倒警示",
             [
                 ("多人鎖定挑戰：", "鏡頭前有多人走動時，若每個人都能觸發拍照，系統會混亂崩潰。"),
                 ("主控者 (Commander) 機制：", "透過演算法選出【畫面中央最接近中心者】或【邊界框面積最大者】作為唯一具備控制權的主控者。"),
                 ("多人跌倒警示 (Fall Detection)：", "計算人體長寬比（寬度 > 高度）或頭部重心急劇驟降。"),
                 ("視覺警報回饋：", "一旦畫面中任何一人倒下，外框即刻閃爍全螢幕紅色警示，並發出 Alert 提示音與自動拍照存證。")
             ], top_bar_color=C_ORANGE)

    # SLIDE 11
    s11 = prs.slides.add_slide(blank_layout)
    add_header(s11, "【第 04 節課】功能實作提示語 (Prompt) 與教師引導", 11)
    add_prompt_card(s11, Inches(0.8), Inches(1.45), Inches(11.7), Inches(2.65),
                    "YOLO Pose 骨架手勢控制：舉手拍照與平舉錄影",
                    "請使用 Python、Ultralytics YOLO11/YOLOv8 Pose 模型與 OpenCV 開發手勢控制相機：\n"
                    "1. 載入 yolo11n-pose.pt 輕量模型，即時偵測人體 17 處關鍵點骨架。\n"
                    "2. 手勢一【舉右手拍照】：當偵測到右手腕高度高於右眼/右肩時，觸發自動拍照存檔，並加入 3 秒冷卻時間（Cooldown）防止連續誤拍。\n"
                    "3. 手勢二【雙手平舉錄影】：當偵測到雙手平舉（手腕與肩膀同高且展開）超過 1.5 秒時，切換『開始/停止錄影』狀態，畫面顯示紅色 REC 閃爍圖示。\n"
                    "4. 畫面需流暢繪製人體骨架線條與關節點，文字提示使用微軟正黑體繁體中文。",
                    "提醒學生鏡頭前退後 1.5~2 公尺，確保上半身完整入鏡，避免骨架遮蔽誤判。")
    add_prompt_card(s11, Inches(0.8), Inches(4.3), Inches(11.7), Inches(2.65),
                    "中央主控者模式 ＆ 多人跌倒全螢幕紅框警示",
                    "請在 YOLO Pose 系統中升級【多人智慧安防與主控者架構】：\n"
                    "1. 多人主控者鎖定：計算畫面中所有人的 Bounding Box 中心點，鎖定最靠近畫面正中央的人為【主控者 (VIP)】，在其頭頂標註金色『主控者』，只有此人能透過舉手觸發拍照/錄影。\n"
                    "2. 多人跌倒急救警報：同時監控畫面中所有被偵測者。若任何一人人體外框『寬度大於高度（橫躺）』且重心低於門檻，判定為【跌倒 (Fall Alert)】！\n"
                    "3. 警報回饋：整個視窗外框閃爍粗紅色警示框，播放嗶聲或顯示『🚨 偵測到人員跌倒！』，並自動截圖儲存至 'fall_events' 資料夾。",
                    "引導學生模擬不同角度的跌倒（假裝滑倒、趴下），體驗專題商品化測試流程。")

    # SLIDE 12
    s12 = prs.slides.add_slide(blank_layout)
    add_header(s12, "課堂實作評分要點 ＆ 硬體環境避坑指南", 12)
    add_card(s12, Inches(0.8), Inches(1.45), Inches(5.6), Inches(5.45),
             "🎯 學生專題評分標準 (Evaluation Matrix)",
             [
                 ("1. 功能完整性 (40%)：", "WebCam 串流穩定，快捷鍵拍照、手勢辨識與跌倒警示能精確觸發。"),
                 ("2. 程式抗崩潰度 (20%)：", "拔除鏡頭、誤按鍵盤、無人入鏡時，程式不報錯閃退，具備完整 Exception Handling。"),
                 ("3. Prompt 工程素養 (20%)：", "能展示與 AI 對話記錄，清楚描述需求演進與除錯過程，而非只抄最終程式碼。"),
                 ("4. 畫面互動體驗 (20%)：", "1:1 等比不變形、繁體中文標籤清晰、冷卻倒數與警示視覺化良好。")
             ], border_color=C_TEAL, top_bar_color=C_TEAL)
    add_card(s12, Inches(6.9), Inches(1.45), Inches(5.6), Inches(5.45),
             "⚠️ 現場環境避雷秘笈 (Troubleshooting)",
             [
                 ("攝影機絕對要架固定架：", "切忌用手拿著鏡頭晃來晃去！鏡頭劇烈震動會造成 YOLO 關節點抖動誤判與計算延遲。"),
                 ("距離與視野掌控：", "人體骨架偵測需距離鏡頭 1.5 ~ 3 公尺，避免『大頭貼』式近距離導致肢體出鏡。"),
                 ("防連拍機制 (Cooldown)：", "手勢辨識每秒推論 30 次，若未加上計時冷卻（例如拍完後鎖定 3 秒），手一舉起瞬間會拍下 50 張垃圾照片。"),
                 ("USB 頻寬與接口：", "雙鏡頭測試請插在電腦前後不同 USB Root Hub 上，避免集線器頻寬塞車。")
             ], border_color=C_ORANGE, top_bar_color=C_ORANGE)

    # SLIDE 13
    s13 = prs.slides.add_slide(blank_layout)
    add_header(s13, "教師專用：課堂 Prompt 快速取用萬用錦囊卡", 13)
    add_card(s13, Inches(0.8), Inches(1.45), Inches(11.7), Inches(5.45),
             "📌 複製即用：四階段課堂教學指令庫",
             [
                 ("階段一【起手式】：", "「我是高中專題學生，請用 Python 和 OpenCV 寫一個 WebCam 測試程式，支援按 's' 拍照、按 'q' 結束，照片自動命名並保存在 captures 資料夾，請包含中文註解。」"),
                 ("階段二【防變形與品質】：", "「請修改剛才的程式，加入 1:1 等比例畫面校正，視窗縮放時人物不可變形（以黑邊填充），並在畫面角落顯示即時畫面亮度是否過暗或過曝。」"),
                 ("階段三【多分割與效能】：", "「請將畫面改造成 2x2 四宮格監控版面，支援雙攝影機輸入與滑鼠點擊放大單格畫面，若系統過慢請提供降採樣優化代碼。」"),
                 ("階段四【AI 智慧辨識】：", "「請導入 YOLO11-Pose 模型，偵測人體 17 個骨架點。實作『舉右手拍照（冷卻 3 秒）』與『畫面中央主控者模式』；當有人橫躺跌倒時，畫面閃爍紅框發出警報。」"),
                 ("通用【自動修復黃金指令】：", "「每次開發完程式，請啟用 MCP 模式，做程式功能的自我檢核，若有問題請直接修正。」")
             ], border_color=C_NAVY, top_bar_color=C_TEAL)

    # SLIDE 14: 總結
    s14 = prs.slides.add_slide(blank_layout)
    apply_transition(s14)
    bg14 = s14.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.5))
    bg14.fill.solid()
    bg14.fill.fore_color.rgb = C_NAVY
    bg14.line.fill.background()

    tb = s14.shapes.add_textbox(Inches(1.5), Inches(1.8), Inches(10.3), Inches(4.2))
    tf = tb.text_frame
    tf.word_wrap = True

    p = tf.paragraphs[0]
    p.text = "AI 時代的專題教育：引導學生從『打字員』蛻變為『系統架構師』"
    p.font.size = Pt(30)
    p.font.bold = True
    p.font.name = "微軟正黑體"
    p.font.color.rgb = RGBColor(255, 255, 255)
    p.space_after = Pt(20)

    p2 = tf.add_paragraph()
    p2.text = "「以前專案要寫半年，現在半年可以完成十個專案！」\n" \
              "掌握 Prompt 提示詞規格化、理解電腦視覺幾何與算力本質、學會讓 AI 自我檢核，\n" \
              "每一位老師與學生都能輕鬆打造出具備商業價值的智慧視覺系統。"
    p2.font.size = Pt(17)
    p2.font.name = "微軟正黑體"
    p2.font.color.rgb = C_TEAL
    p2.space_after = Pt(30)
    p2.line_spacing = 1.35

    p3 = tf.add_paragraph()
    p3.text = "🌐 線上課程與可複製 Prompt 講義：https://stjedu.github.io/www/yolo_webcam_course.html"
    p3.font.size = Pt(15)
    p3.font.bold = True
    p3.font.name = "微軟正黑體"
    p3.font.color.rgb = RGBColor(245, 180, 0)

    out1 = os.path.join(ROOT_DIR, "WebCam_AI視覺與YOLO應用_教學投影片.pptx")
    out2 = os.path.join(ROOT_DIR, "yolo_webcam_slides.pptx")
    prs.save(out1)
    shutil.copy2(out1, out2)
    print(f"[OK] YOLO deck saved: {out1} & {out2} ({len(prs.slides)} slides)")


def build_esp32_deck():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    C_NAVY = RGBColor(15, 23, 42)
    C_BLUE = RGBColor(2, 132, 199)
    C_ORANGE = RGBColor(217, 119, 6)
    C_BG_LIGHT = RGBColor(248, 250, 252)
    C_CARD_BG = RGBColor(255, 255, 255)
    C_BORDER = RGBColor(226, 232, 240)
    C_TEXT_DARK = RGBColor(15, 23, 42)
    C_TEXT_MUTED = RGBColor(100, 116, 139)
    C_CODE_BG = RGBColor(241, 245, 249)
    C_CODE_TEXT = RGBColor(30, 41, 59)

    TOTAL_SLIDES = 11
    blank_layout = prs.slide_layouts[6]

    def apply_transition(slide):
        trans_xml = parse_xml(f'<p:transition {nsdecls("p")} spd="med"><p:fade/></p:transition>')
        slide._element.append(trans_xml)

    def add_header(slide, title, slide_idx, category="ESP32 ✕ 物聯網 MQTT ✕ 邊緣 AI 實戰"):
        apply_transition(slide)
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
        bg.fill.solid()
        bg.fill.fore_color.rgb = C_BG_LIGHT
        bg.line.fill.background()

        header_box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(1.15))
        header_box.fill.solid()
        header_box.fill.fore_color.rgb = C_NAVY
        header_box.line.fill.background()

        accent_bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(1.15), Inches(4.5), Inches(0.05))
        accent_bar.fill.solid()
        accent_bar.fill.fore_color.rgb = C_BLUE
        accent_bar.line.fill.background()

        cat_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.12), Inches(8.5), Inches(0.3))
        tf_c = cat_box.text_frame
        tf_c.word_wrap = True
        p_c = tf_c.paragraphs[0]
        p_c.text = f"📡 {category.upper()}"
        p_c.font.size = Pt(11.5)
        p_c.font.bold = True
        p_c.font.name = "微軟正黑體"
        p_c.font.color.rgb = C_BLUE

        page_box = slide.shapes.add_textbox(Inches(10.0), Inches(0.12), Inches(2.5), Inches(0.3))
        tf_p = page_box.text_frame
        p_p = tf_p.paragraphs[0]
        p_p.alignment = PP_ALIGN.RIGHT
        p_p.text = f"第 {slide_idx:02d} / {TOTAL_SLIDES:02d} 頁"
        p_p.font.size = Pt(12)
        p_p.font.bold = True
        p_p.font.name = "微軟正黑體"
        p_p.font.color.rgb = RGBColor(180, 200, 225)

        tb = slide.shapes.add_textbox(Inches(0.8), Inches(0.40), Inches(11.7), Inches(0.65))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = title
        p.font.size = Pt(24)
        p.font.bold = True
        p.font.name = "微軟正黑體"
        p.font.color.rgb = RGBColor(255, 255, 255)

        footer_box = slide.shapes.add_textbox(Inches(0.8), Inches(7.12), Inches(11.7), Inches(0.3))
        tf_f = footer_box.text_frame
        pf = tf_f.paragraphs[0]
        pf.text = "🌐 互動課程網頁：https://stjedu.github.io/www/esp32_mqtt_course.html   |   主講：簡樹桐 老師"
        pf.font.size = Pt(10.5)
        pf.font.color.rgb = C_TEXT_MUTED
        pf.font.name = "微軟正黑體"

    def add_card(slide, left, top, width, height, title, content_items, bg_color=C_CARD_BG, border_color=C_BORDER, title_color=C_NAVY, top_bar_color=C_BLUE):
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
        card.fill.solid()
        card.fill.fore_color.rgb = bg_color
        card.line.color.rgb = border_color
        card.line.width = Pt(1.5)

        bar = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, Inches(0.1))
        bar.fill.solid()
        bar.fill.fore_color.rgb = top_bar_color
        bar.line.fill.background()

        tb = slide.shapes.add_textbox(left + Inches(0.25), top + Inches(0.2), width - Inches(0.5), height - Inches(0.35))
        tf = tb.text_frame
        tf.word_wrap = True

        if title:
            p_title = tf.paragraphs[0]
            p_title.text = title
            p_title.font.size = Pt(17.5)
            p_title.font.bold = True
            p_title.font.name = "微軟正黑體"
            p_title.font.color.rgb = title_color
            p_title.space_after = Pt(10)

        for idx, item in enumerate(content_items):
            p = tf.add_paragraph() if (title or idx > 0) else tf.paragraphs[0]
            p.space_after = Pt(7)
            p.line_spacing = 1.25
            if isinstance(item, tuple):
                bold_prefix, normal_text = item
                run1 = p.add_run()
                run1.text = bold_prefix + " "
                run1.font.bold = True
                run1.font.size = Pt(13.5)
                run1.font.name = "微軟正黑體"
                run1.font.color.rgb = C_TEXT_DARK

                run2 = p.add_run()
                run2.text = normal_text
                run2.font.bold = False
                run2.font.size = Pt(13.5)
                run2.font.name = "微軟正黑體"
                run2.font.color.rgb = C_TEXT_MUTED
            else:
                run = p.add_run()
                run.text = item
                run.font.size = Pt(13.5)
                run.font.name = "微軟正黑體"
                run.font.color.rgb = C_TEXT_DARK
        return card

    def add_prompt_card(slide, left, top, width, height, prompt_title, prompt_text, tips=""):
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
        card.fill.solid()
        card.fill.fore_color.rgb = C_CODE_BG
        card.line.color.rgb = C_BLUE
        card.line.width = Pt(2)

        tb = slide.shapes.add_textbox(left + Inches(0.25), top + Inches(0.16), width - Inches(0.5), height - Inches(0.32))
        tf = tb.text_frame
        tf.word_wrap = True

        p0 = tf.paragraphs[0]
        p0.text = f"💡 實戰提示詞 (Prompt)：{prompt_title}   [可直接反白複製]"
        p0.font.size = Pt(15)
        p0.font.bold = True
        p0.font.name = "微軟正黑體"
        p0.font.color.rgb = C_ORANGE
        p0.space_after = Pt(6)

        p1 = tf.add_paragraph()
        run1 = p1.add_run()
        run1.text = f"「{prompt_text}」"
        run1.font.size = Pt(12.5)
        run1.font.name = "Consolas"
        run1.font.color.rgb = C_CODE_TEXT
        p1.space_after = Pt(6)
        p1.line_spacing = 1.2

        if tips:
            p2 = tf.add_paragraph()
            run2 = p2.add_run()
            run2.text = f"📌 教師引導技巧：{tips}"
            run2.font.size = Pt(11.5)
            run2.font.bold = True
            run2.font.name = "微軟正黑體"
            run2.font.color.rgb = C_BLUE

    # SLIDE 1: 封面
    s1 = prs.slides.add_slide(blank_layout)
    apply_transition(s1)
    bg1 = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.5))
    bg1.fill.solid()
    bg1.fill.fore_color.rgb = C_NAVY
    bg1.line.fill.background()

    c1 = s1.shapes.add_shape(MSO_SHAPE.OVAL, Inches(9.6), Inches(-1.2), Inches(5.0), Inches(5.0))
    c1.fill.solid()
    c1.fill.fore_color.rgb = RGBColor(22, 42, 74)
    c1.line.fill.background()

    tb = s1.shapes.add_textbox(Inches(1.2), Inches(1.7), Inches(11.0), Inches(4.2))
    tf = tb.text_frame
    tf.word_wrap = True

    p0 = tf.paragraphs[0]
    p0.text = "📡 智慧物聯網與邊緣 AI 實習系列"
    p0.font.size = Pt(17)
    p0.font.bold = True
    p0.font.name = "微軟正黑體"
    p0.font.color.rgb = C_ORANGE
    p0.space_after = Pt(12)

    p = tf.add_paragraph()
    p.text = "ESP32 與 MQTT 智慧控制教學實戰"
    p.font.size = Pt(40)
    p.font.bold = True
    p.font.name = "微軟正黑體"
    p.font.color.rgb = RGBColor(255, 255, 255)
    p.space_after = Pt(14)

    p2 = tf.add_paragraph()
    p2.text = "BME280 感測器探索 ‧ 四按鈕 OLED 碼表 ‧ MQTT 雲端遙控 ‧ AI 視覺辨識連動"
    p2.font.size = Pt(20)
    p2.font.name = "微軟正黑體"
    p2.font.color.rgb = C_BLUE
    p2.space_after = Pt(28)

    p3 = tf.add_paragraph()
    p3.text = "主講：簡樹桐 老師 ｜ 適用：高中職專題製作 ‧ 物聯網實習 ‧ AI 應用實作\n🌐 線上互動教材：https://stjedu.github.io/www/esp32_mqtt_course.html"
    p3.font.size = Pt(14.5)
    p3.font.name = "微軟正黑體"
    p3.font.color.rgb = RGBColor(200, 215, 235)

    # SLIDE 2
    s2 = prs.slides.add_slide(blank_layout)
    add_header(s2, "課程全景：ESP32 ✕ 物聯網 MQTT ✕ 視覺 AI 的系統整合", 2)
    add_card(s2, Inches(0.8), Inches(1.45), Inches(5.6), Inches(5.45),
             "🚀 專題學習核心思維",
             [
                 ("陌生物件不要怕：", "拿到完全沒看過的感測器（如 BME280），不需死記暫存器，直接問 AI 腳位定義、接線與讀取程式。"),
                 ("硬體插拔防呆：", "杜邦線與 Type-C 必須『水平插拔』，禁止往上撩掰，避免接頭折斷與針腳損壞。"),
                 ("Agent 自動執行設定：", "在開發環境將終端命令設定為 Allow / Always Proceed，免去每次按確認的干擾。"),
                 ("從邊緣到雲端：", "循序實作『本機感測 ➔ 本機四按鈕控制 ➔ MQTT 雲端手機遙控 ➔ AI 鏡頭視覺連動』四階段。")
             ], border_color=C_BLUE, top_bar_color=C_BLUE)
    add_card(s2, Inches(6.9), Inches(1.45), Inches(5.6), Inches(5.45),
             "📚 三大核心實作單元",
             [
                 ("第 01 節課：ESP32 與感測器電路實作", "BME280 溫濕度大氣壓感測器、I2C 協定、OLED 即時顯示與 AI 代理人安全設定。"),
                 ("第 02 節課：MQTT 協定與手機 App 智慧控制", "四按鈕碼表電路、MQTT Broker (mqttgo.io:1883)、IoT MQTT Panel 手機 App 雙向發布與訂閱。"),
                 ("第 03 節課：AI 視覺影像辨識初探與連動", "USB 攝影機環境、YOLO 物件/骨架/分割三合一模式、畫面各類別總數即時統計與截圖除錯。")
             ], border_color=C_NAVY, top_bar_color=C_ORANGE)

    # SLIDE 3
    s3 = prs.slides.add_slide(blank_layout)
    add_header(s3, "【第 01 節課】ESP32 環境建置、BME 感測器與 OLED 顯示", 3)
    add_card(s3, Inches(0.8), Inches(1.45), Inches(5.6), Inches(5.45),
             "🎯 學習目標與硬體防護",
             [
                 ("AI 代理人權限配置：", "將安全層級設為 Custom，Outside of Project 設為 Allow，命令執行設為 Always Proceed。"),
                 ("硬體插拔防損原則：", "Type-C 與模組排針必須平行施力，嚴禁向上向下折彎，避免損壞昂貴板卡。"),
                 ("BME280 感測器原理：", "整合式環境晶片，具備溫度、相對濕度與大氣壓力 (hPa) 測量。"),
                 ("I2C 匯流排並聯：", "ESP32 預設 I2C 腳位（SCL: GPIO 22, SDA: GPIO 21），可並聯多個 I2C 設備（0x76 與 0x3C）。")
             ], top_bar_color=C_BLUE)
    add_card(s3, Inches(6.9), Inches(1.45), Inches(5.6), Inches(5.45),
             "🔌 接線腳位對照表 (Pinout)",
             [
                 ("BME280 SCL ➔", "ESP32 GPIO 22（I2C 時脈線）"),
                 ("BME280 SDA ➔", "ESP32 GPIO 21（I2C 數據線）"),
                 ("OLED SCL / SDA ➔", "並聯至 GPIO 22 / 21（共用匯流排）"),
                 ("電源供應 ➔", "3.3V 與 GND 接至 ESP32 對應供電腳位"),
                 ("成果反饋 ➔", "OLED 畫面上即時輪播 Temp (23.7°C)、Humi (45.6%)、Pres (994.1 hPa)")
             ], top_bar_color=C_ORANGE)

    # SLIDE 4
    s4 = prs.slides.add_slide(blank_layout)
    add_header(s4, "【第 01 節課】陌生感測器上手 ＆ OLED 顯示 Prompt 提示語", 4)
    add_prompt_card(s4, Inches(0.8), Inches(1.45), Inches(11.7), Inches(2.65),
                    "陌生感測器快速探索與引導 Prompt",
                    "我手邊有一顆陌生的感測模組，上面標示『BME』，腳位有 VCC, GND, SCL, SDA。\n"
                    "請扮演資深物聯網硬體導師：\n"
                    "1. 告訴我這顆感測器主要能測量哪些物理量？（例如溫濕度、大氣壓力）\n"
                    "2. 我使用的是 ESP32 實驗板，請告訴我如何將它與 I2C 腳位（GPIO 21, 22）進行實體接線？\n"
                    "3. 請使用 MicroPython 撰寫一段簡單的掃描 I2C 位址（i2c.scan()）與數值讀取測試代碼。",
                    "拿到任何不知名零件時，引導學生不要停滯，第一時間向 AI 詢問定義與引腳接法。")
    add_prompt_card(s4, Inches(0.8), Inches(4.3), Inches(11.7), Inches(2.65),
                    "BME280 數值讀取並即時顯示於 OLED 螢幕",
                    "請替 ESP32 撰寫 MicroPython 程式：\n"
                    "1. 同時驅動 BME280 感測器與 0.96 吋 SSD1306 OLED 螢幕（兩者皆接在 I2C: SCL 22, SDA 21）。\n"
                    "2. 每隔 1 秒讀取一次環境數值，並在 OLED 螢幕上清晰顯示三行數據：\n"
                    "   - 第一行：『Temp: XX.X C』\n"
                    "   - 第二行：『Humi: XX.X %』\n"
                    "   - 第三行：『Pres: XXXX.X hPa』\n"
                    "3. 包含完整例外防呆機制（若讀取失敗顯示 'Sensor Error'，不可當機閃退）。",
                    "讓學生觀察 I2C 匯流排並聯的優點：兩條訊號線即可同時掛載感測器與螢幕。")

    # SLIDE 5
    s5 = prs.slides.add_slide(blank_layout)
    add_header(s5, "【第 02 節課】四按鈕碼表電路 ＆ MQTT 雲端智慧控制", 5)
    add_card(s5, Inches(0.8), Inches(1.45), Inches(5.6), Inches(5.45),
             "⏱️ 四按鈕碼表系統設計",
             [
                 ("按鈕低電位動作 (Active Low)：", "按鈕未按下為 3.3V 高電位，按下時導通接地 (GND) 為 0V 低電位，啟用內部上拉電阻 (PULL_UP)。"),
                 ("四按鈕腳位分配：", "PB1 (GPIO 14), PB2 (GPIO 27), PB3 (GPIO 26), PB4 (GPIO 25)。"),
                 ("按鈕功能定義：", "PB1: 開始/暫停 (Start/Pause)；PB2: 計圈 (Lap)；PB3: 歸零 (Reset)；PB4: 畫面模式切換。"),
                 ("OLED 即時刷新：", "以毫秒級精度更新計時器，避免阻塞式延遲 (`time.sleep`)，採用非同步或定時器中斷。")
             ], top_bar_color=C_BLUE)
    add_card(s5, Inches(6.9), Inches(1.45), Inches(5.6), Inches(5.45),
             "☁️ MQTT 雲端物聯網連線三要素",
             [
                 ("Broker 伺服器：", "使用免費公開伺服器 `mqttgo.io`（不可打錯）。"),
                 ("通訊埠 (Port)：", "`1883`（標準非加密 TCP 傳輸協定）。"),
                 ("Client ID 與 Topic：", "Client ID 自訂唯一名稱（避免撞名斷線），Topic 如 `school/esp32/control` 與 `school/esp32/status`。"),
                 ("手機 App 配對：", "安裝 IoT MQTT Panel App，建立按鈕發布開關指令，並訂閱 ESP32 傳回的計時或狀態訊息。")
             ], top_bar_color=C_ORANGE)

    # SLIDE 6
    s6 = prs.slides.add_slide(blank_layout)
    add_header(s6, "【第 02 節課】四按鈕碼表 ＆ MQTT 雲端雙向通訊 Prompt 提示語", 6)
    add_prompt_card(s6, Inches(0.8), Inches(1.45), Inches(11.7), Inches(2.65),
                    "四按鈕低電位動作之 OLED 碼表控制程式",
                    "請替 ESP32 編寫一個功能完備的 MicroPython 碼表程式：\n"
                    "1. 硬體配置：0.96 吋 OLED (I2C: 21, 22)，四顆實體按鈕接在 GPIO 14, 27, 26, 25（啟用內部 PULL_UP，低電位觸發）。\n"
                    "2. 按鈕邏輯：\n"
                    "   - PB1 (GPIO 14)：開始計時 / 暫停計時\n"
                    "   - PB2 (GPIO 27)：記錄當前圈數時間 (Lap Time)\n"
                    "   - PB3 (GPIO 26)：碼表歸零 (Reset)\n"
                    "   - PB4 (GPIO 25)：切換主畫面與歷史圈數列表\n"
                    "3. 需具備軟體防彈跳 (Debounce) 機制，確保按鍵靈敏且不重送。",
                    "讓學生體驗實體 UI 按鍵互動，並理解硬體內部上拉電阻的軟體宣告語法。")
    add_prompt_card(s6, Inches(0.8), Inches(4.3), Inches(11.7), Inches(2.65),
                    "ESP32 連線 MQTT Broker (mqttgo.io) 雲端遙控",
                    "請編寫 ESP32 MicroPython 程式，結合 Wi-Fi 與 MQTT 協定：\n"
                    "1. 連線至 Wi-Fi（SSID 與密碼），連線成功後在 OLED 顯示配發的 IP。\n"
                    "2. 連接至 MQTT 伺服器：Broker 網址為 'mqttgo.io'，Port 為 1883。\n"
                    "3. 訂閱主題 'school/stj/cmd'：若收到 'START' 啟動碼表計時，收到 'STOP' 停止，收到 'RESET' 歸零。\n"
                    "4. 發布主題 'school/stj/state'：每秒將當前碼表時間推播回 MQTT，供手機 IoT MQTT Panel 即時監看。\n"
                    "5. 請提供完整代碼與手機 App 儀表板設定指引。",
                    "展示『物聯網雙向控制』：手機在雲端按按鈕，遠端的 ESP32 馬上同步動作！")

    # SLIDE 7
    s7 = prs.slides.add_slide(blank_layout)
    add_header(s7, "【第 03 節課】AI 視覺影像辨識初探與五步驟 SOP", 7)
    add_card(s7, Inches(0.8), Inches(1.45), Inches(5.6), Inches(5.45),
             "👁️ AI 視覺辨識五步驟實作 SOP",
             [
                 ("步驟 1：硬體與環境檢查", "確認 USB 攝影機正常插入，檢查 Python、OpenCV 與 Ultralytics 套件就緒。"),
                 ("步驟 2：三種核心模式切換", "按 '1': 通用物件偵測 (80類)；按 '2': 人體姿態骨架 (Pose 17點)；按 '3': 語意實例分割 (Segmentation 去背)。"),
                 ("步驟 3：即時物件分類總數統計", "辨識後不只畫方框，更要即時統計：人數、椅子數、杯子數、筆電數並顯示於畫面。"),
                 ("步驟 4：截圖除錯 (Screenshot Debug)", "程式跑不出來不要猜，直接螢幕截圖貼給 AI，AI 立即閱讀症狀並執行修復！"),
                 ("步驟 5：專題延伸思考", "思考如何將辨識結果（如手勢求救、走道有人）透過 MQTT 送到 ESP32 發出警報。")
             ], top_bar_color=C_BLUE)
    add_card(s7, Inches(6.9), Inches(1.45), Inches(5.6), Inches(5.45),
             "💡 簡老師的 AI 開發箴言",
             [
                 ("「我只動嘴巴就好了」：", "只要你有 Idea 和清晰的邏輯，動嘴巴下達指令，AI 就會替你把程式碼生出來。"),
                 ("算力與解析度的取捨：", "電腦有多核心跑得動，但若移植到邊緣小裝置，絕對不可開高解析度，夠用即可！"),
                 ("專案不要閉門造車：", "先去看別人做過的專題，拿別人的架構過來改進為自己的方法，這才是 AI 時代的高效研發！"),
                 ("比賽與商品化思維：", "從生活情境（如視障測距、紅綠燈辨識、家中手勢求救）切入，做出具有實用價值的完整系統。")
             ], top_bar_color=C_ORANGE)

    # SLIDE 8
    s8 = prs.slides.add_slide(blank_layout)
    add_header(s8, "【第 03 節課】三合一影像辨識 ＆ 物件即時統計 Prompt 提示語", 8)
    add_prompt_card(s8, Inches(0.8), Inches(1.45), Inches(11.7), Inches(2.65),
                    "OpenCV + YOLO11 三合一模式切換與繁體中文標籤",
                    "請使用 Python、OpenCV 與 Ultralytics YOLO11 開發即時 WebCam 視覺程式：\n"
                    "1. 自動開啟相機串流，並支援鍵盤熱鍵動態切換三種模式：\n"
                    "   - 按 '1'：通用物件偵測模式（yolo11n.pt），偵測 80 種常見物品\n"
                    "   - 按 '2'：人體姿態骨架模式（yolo11n-pose.pt），繪製 17 處關節點骨架\n"
                    "   - 按 '3'：實例分割模式（yolo11n-seg.pt），繪製物體輪廓多邊形遮罩\n"
                    "2. 畫面文字標籤必須使用系統微軟正黑體顯示【繁體中文】（如：人、椅子、瓶子、筆記型電腦）。\n"
                    "3. 畫面左上角顯示當前模式名稱與即時 FPS 幀率，按 'q' 鍵退出。",
                    "先讓學生感受 3 種不同 AI 模型的視覺威力，理解『物件、骨架、分割』在電腦視覺中的差異。")
    add_prompt_card(s8, Inches(0.8), Inches(4.3), Inches(11.7), Inches(2.65),
                    "現場物件類別數量即時統計與報表看板",
                    "請在上述 YOLO 辨識程式中新增【物件分類總數即時統計面板】：\n"
                    "1. 即時遍歷畫面中所有被偵測到的物件，統計各類別的出現數量（例如：人 x 5, 椅子 x 8, 杯子 x 2, 螢幕 x 3）。\n"
                    "2. 在畫面右上方半透明資訊欄中，由多到少即時列出統計清單與畫面總物件數。\n"
                    "3. 當偵測到『人』的數量發生變化時，在終端機印出警示日誌。\n"
                    "4. 代碼需優化推論速度，避免因統計運算導致畫面掉幀。",
                    "這是人流計數、自駕辨識與智慧門禁的基礎，讓學生體會從『視覺畫方框』昇華為『數據分析』。")

    # SLIDE 9
    s9 = prs.slides.add_slide(blank_layout)
    add_header(s9, "物聯網與 AI 整合架構 ＆ 課堂避坑指南", 9)
    add_card(s9, Inches(0.8), Inches(1.45), Inches(5.6), Inches(5.45),
             "🌐 虛實整合系統架構 (IoT + AI Flow)",
             [
                 ("1. 視覺感知端：", "WebCam 透過 YOLO 辨識出特定物件（例如未戴安全帽、有人倒地或特定手勢）。"),
                 ("2. 雲端傳輸端：", "Python 腳本將事件打包為 JSON，透過 MQTT 發布至 `mqttgo.io`。"),
                 ("3. 邊緣執行端：", "ESP32 訂閱主題，接收到訊息後立即驅動蜂鳴器發出警報，或在 OLED 上顯示警示。"),
                 ("4. 手機監控端：", "教師或管理員手機透過 IoT MQTT Panel 收到推播通知與狀態回饋。")
             ], border_color=C_BLUE, top_bar_color=C_BLUE)
    add_card(s9, Inches(6.9), Inches(1.45), Inches(5.6), Inches(5.45),
             "⚠️ 現場環境避雷秘笈 (Troubleshooting)",
             [
                 ("杜邦線與排針不要往上撩：", "插拔請保持水平，垂直用力，避免排針彎折或 Type-C 接頭鬆脫。"),
                 ("MQTT Broker 拼字不要錯：", "伺服器必須是 `mqttgo.io`，不可少打點或打錯字母，埠號為 1883。"),
                 ("Client ID 衝突導致閃退：", "若兩位同學用同一個 Client ID 連線，伺服器會踢掉前一人造成頻繁斷線，請加上座號區分！"),
                 ("AI 語意除錯善用截圖：", "當程式有 Bug 但不知原因時，直接把畫面截圖給 AI，AI 讀圖診斷修復率高達 90%！")
             ], border_color=C_ORANGE, top_bar_color=C_ORANGE)

    # SLIDE 10
    s10 = prs.slides.add_slide(blank_layout)
    add_header(s10, "教師專用：ESP32 ✕ MQTT ✕ AI 課堂 Prompt 速查錦囊卡", 10)
    add_card(s10, Inches(0.8), Inches(1.45), Inches(11.7), Inches(5.45),
             "📌 複製即用：四階段實作指令庫",
             [
                 ("階段一【感測器與 OLED】：", "「我是高職學生，使用 ESP32 與 MicroPython。請教我如何接線並讀取 BME280 溫濕度大氣壓感測器，並將數值即時顯示於 0.96 吋 I2C OLED 螢幕上。」"),
                 ("階段二【四按鈕碼表實作】：", "「請使用 MicroPython 為 ESP32 撰寫一個碼表，使用 4 顆低電位按鈕（GPIO 14, 27, 26, 25，內建上拉），實現開始、暫停、計圈、歸零，並於 OLED 顯示。」"),
                 ("階段三【MQTT 雲端手機控制】：", "「請教我用 MicroPython 連上 Wi-Fi 與 MQTT 伺服器 mqttgo.io (Port 1883)，手機使用 IoT MQTT Panel 發布指令開關 LED，並將感測數值定期推播回手機。」"),
                 ("階段四【AI 視覺與多模式切換】：", "「請使用 Python OpenCV 與 YOLO11 輕量模型，做一個相機程式：按 1 切換物件偵測、按 2 切換骨架姿勢、按 3 切換去背分割，並即時統計畫面中所有物件總數量。」"),
                 ("萬能【除錯修復指令】：", "「每次開發完程式，請做程式功能的自我檢核，若有報錯或異常請直接修正代碼並提供修正說明。」")
             ], border_color=C_NAVY, top_bar_color=C_BLUE)

    # SLIDE 11: 總結
    s11 = prs.slides.add_slide(blank_layout)
    apply_transition(s11)
    bg11 = s11.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.5))
    bg11.fill.solid()
    bg11.fill.fore_color.rgb = C_NAVY
    bg11.line.fill.background()

    tb = s11.shapes.add_textbox(Inches(1.5), Inches(1.8), Inches(10.3), Inches(4.2))
    tf = tb.text_frame
    tf.word_wrap = True

    p = tf.paragraphs[0]
    p.text = "物聯網與 AI 專題：讓學生從感測器出發，串聯雲端與人工智慧"
    p.font.size = Pt(30)
    p.font.bold = True
    p.font.name = "微軟正黑體"
    p.font.color.rgb = RGBColor(255, 255, 255)
    p.space_after = Pt(20)

    p2 = tf.add_paragraph()
    p2.text = "「不需死背接線與暫存器，善用 AI 與清晰架構，\n" \
              "每位學生都能在一堂課內實現感測、雲端推播與 AI 視覺聯網的完整作品！」"
    p2.font.size = Pt(18)
    p2.font.name = "微軟正黑體"
    p2.font.color.rgb = C_BLUE
    p2.space_after = Pt(30)
    p2.line_spacing = 1.35

    p3 = tf.add_paragraph()
    p3.text = "🌐 線上課程與可複製 Prompt 講義：https://stjedu.github.io/www/esp32_mqtt_course.html"
    p3.font.size = Pt(15)
    p3.font.bold = True
    p3.font.name = "微軟正黑體"
    p3.font.color.rgb = C_ORANGE

    out1 = os.path.join(ROOT_DIR, "ESP32_MQTT_智慧控制教學_教學投影片.pptx")
    out2 = os.path.join(ROOT_DIR, "esp32_mqtt_slides.pptx")
    prs.save(out1)
    shutil.copy2(out1, out2)
    print(f"[OK] ESP32 deck saved: {out1} & {out2} ({len(prs.slides)} slides)")


if __name__ == "__main__":
    build_main_deck()
    build_yolo_deck()
    build_esp32_deck()
