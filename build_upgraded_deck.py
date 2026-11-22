import os
import pptx
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

def build_upgraded_deck():
    base_dir = r"d:\UPES 2nd YEar\ML\ML PROJECT"
    pptx_path = os.path.join(base_dir, "Transportation_Delay_System_UPGRADED_VIVA.pptx")
    pptx_fallback = os.path.join(base_dir, "Transportation_Delay_System_Final_Viva_Presentation.pptx")
    fig_dir = os.path.join(base_dir, "figures")

    fig_title_bg = os.path.join(fig_dir, "title_bg.png")
    fig_map = os.path.join(fig_dir, "corridor_map.png")
    fig_loss = os.path.join(fig_dir, "loss_comparison.png")
    fig_opt = os.path.join(fig_dir, "optimizer_convergence.png")
    fig_lasso = os.path.join(fig_dir, "lasso_coefficients.png")
    fig_prep = os.path.join(fig_dir, "preprocessing_pipeline.png")
    fig_flow = os.path.join(fig_dir, "pipeline_architecture_diagram.png")
    fig_reg = os.path.join(fig_dir, "upes_regression_fit.png")
    fig_roc = os.path.join(fig_dir, "upes_roc_curve.png")

    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    # Color Palette
    NAVY = RGBColor(0x1B, 0x36, 0x5D)
    TEAL = RGBColor(0x00, 0x80, 0x80)
    DARK_GRAY = RGBColor(0x22, 0x22, 0x22)
    LIGHT_GRAY = RGBColor(0x55, 0x55, 0x55)
    CARD_BG = RGBColor(0xF8, 0xFA, 0xFC)
    CARD_BORDER = RGBColor(0xD0, 0xD7, 0xDE)
    WHITE = RGBColor(0xFF, 0xFF, 0xFF)
    ORANGE = RGBColor(0xE6, 0x51, 0x00)
    LIGHT_BLUE = RGBColor(0xEB, 0xF3, 0xFA)
    ACCENT_GREEN = RGBColor(0x2E, 0x7D, 0x32)
    GOLD = RGBColor(0xD9, 0x9B, 0x00)

    def add_header(slide, title_text, category_text="APPLIED MACHINE LEARNING (CSAI2017P) | LAB EVALUATION"):
        cat_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.32), Inches(11.7), Inches(0.3))
        tf_cat = cat_box.text_frame
        tf_cat.word_wrap = True
        tf_cat.margin_left = tf_cat.margin_right = tf_cat.margin_top = tf_cat.margin_bottom = 0
        p_cat = tf_cat.paragraphs[0]
        p_cat.text = category_text.upper()
        p_cat.font.size = Pt(9.5)
        p_cat.font.bold = True
        p_cat.font.color.rgb = TEAL

        title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.62), Inches(11.7), Inches(0.65))
        tf_title = title_box.text_frame
        tf_title.word_wrap = True
        tf_title.margin_left = tf_title.margin_right = tf_title.margin_top = tf_title.margin_bottom = 0
        p_title = tf_title.paragraphs[0]
        p_title.text = title_text
        p_title.font.size = Pt(21)
        p_title.font.bold = True
        p_title.font.color.rgb = NAVY

    def create_card(slide, left, top, width, height, bg_color=CARD_BG, border_color=CARD_BORDER):
        shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
        shape.fill.solid()
        shape.fill.fore_color.rgb = bg_color
        if border_color:
            shape.line.color.rgb = border_color
            shape.line.width = Pt(1.2)
        else:
            shape.line.fill.background()
        return shape

    # ==============================================================================
    # 1. UPGRADE SLIDE 1: TITLE SLIDE (Full-Slide Background Image + Dark Blue Overlay)
    # ==============================================================================
    s1 = prs.slides.add_slide(blank_layout)

    # Full Background Image if exists
    if os.path.exists(fig_title_bg):
        s1.shapes.add_picture(fig_title_bg, 0, 0, width=Inches(13.333), height=Inches(7.5))
    else:
        bg1 = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
        bg1.fill.solid()
        bg1.fill.fore_color.rgb = NAVY
        bg1.line.fill.background()

    # Dark Blue Translucent Container Card
    card_title = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.9), Inches(0.8), Inches(11.533), Inches(5.9))
    card_title.fill.solid()
    card_title.fill.fore_color.rgb = RGBColor(0x0C, 0x1E, 0x38)
    card_title.line.color.rgb = RGBColor(0x3B, 0x62, 0x9B)
    card_title.line.width = Pt(1.5)

    tb = s1.shapes.add_textbox(Inches(1.4), Inches(1.1), Inches(10.5), Inches(5.3))
    tf = tb.text_frame
    tf.word_wrap = True

    p0 = tf.paragraphs[0]
    p0.text = "APPLIED MACHINE LEARNING LAB EVALUATION (CSAI2017P)"
    p0.font.size = Pt(12)
    p0.font.bold = True
    p0.font.color.rgb = RGBColor(0x00, 0xE5, 0xFF)
    p0.space_after = Pt(10)

    p1 = tf.add_paragraph()
    p1.text = "UPES Transportation Delay System"
    p1.font.size = Pt(40)  # Huge 40pt Title!
    p1.font.bold = True
    p1.font.color.rgb = WHITE
    p1.space_after = Pt(8)

    p2 = tf.add_paragraph()
    p2.text = "Predicting Bus Travel Time & Corridor Delays for UPES Dehradun"
    p2.font.size = Pt(17)
    p2.font.color.rgb = RGBColor(0xCF, 0xD8, 0xDC)
    p2.space_after = Pt(28)

    p3 = tf.add_paragraph()
    p3.text = "PROJECT TEAM (BATCHES 17 & 18):"
    p3.font.size = Pt(11)
    p3.font.bold = True
    p3.font.color.rgb = RGBColor(0xFF, 0xD5, 0x4F)
    p3.space_after = Pt(6)

    p4 = tf.add_paragraph()
    p4.text = "👑 Sparsh Srivastava (Group Leader)  |  Units III & VI\n📈 Yasharth Jaiswal  |  Units II & V\n⚙️ Aayushman Kothari  |  Units I & IV"
    p4.font.size = Pt(13)
    p4.font.color.rgb = WHITE
    p4.space_after = Pt(20)

    p5 = tf.add_paragraph()
    p5.text = "School of Computer Science | UPES Dehradun (Bidholi & Kandoli Campuses)"
    p5.font.size = Pt(11)
    p5.font.color.rgb = RGBColor(0x90, 0xA4, 0xAE)

    # ==============================================================================
    # 2. UPGRADE SLIDE 2: PROBLEM STATEMENT (Visual Journey Map + 3 Fast-Read Cards)
    # ==============================================================================
    s2 = prs.slides.add_slide(blank_layout)
    add_header(s2, "Problem Statement: The Dehradun–Bidholi Commute Trap")

    # Left: Route Journey Map Card
    c_map = create_card(s2, Inches(0.8), Inches(1.4), Inches(6.0), Inches(5.6))
    if os.path.exists(fig_map):
        s2.shapes.add_picture(fig_map, Inches(0.9), Inches(1.5), width=Inches(5.8))
        cap = s2.shapes.add_textbox(Inches(0.9), Inches(6.55), Inches(5.8), Inches(0.35))
        cp = cap.text_frame.paragraphs[0]
        cp.text = "Transit Route: ISBT -> Prem Nagar -> ⚠️ Nanda Ki Chowki -> ⚠️ Pondha -> Bidholi"
        cp.font.size = Pt(9.5)
        cp.font.italic = True
        cp.font.color.rgb = LIGHT_GRAY
        cp.alignment = PP_ALIGN.CENTER

    # Right: 3 Visual High-Impact Cards (Read in 5 seconds)
    three_cards = [
        ("🚦 Congestion Bottleneck", 
         "• Prem Nagar bazaar & Nanda Ki Chowki single-lane bottleneck.\n• Minor breakdown turns a 75-min trip into 2.5 hours of gridlock.", 
         ORANGE, RGBColor(0xFF, 0xF5, 0xEB)),
        ("🌧 Monsoon Delay", 
         "• Hill rain causes waterlogging and slippery steep ghat turns.\n• Extreme weather spikes travel time by +40 mins unexpectedly.", 
         TEAL, RGBColor(0xE6, 0xFF, 0xFA)),
        ("🚌 Fixed Timetable", 
         "• College buses leave at fixed slots (6:15 AM, 7:00 AM, 2:00 PM, 5:30 PM).\n• Students miss 8:00 AM labs or face early departure dilemmas.", 
         NAVY, RGBColor(0xEB, 0xF3, 0xFA))
    ]

    card_h = Inches(1.72)
    card_gap = Inches(0.2)
    for idx, (title, text, edge_c, bg_c) in enumerate(three_cards):
        top_y = Inches(1.4) + idx * (card_h + card_gap)
        c = create_card(s2, Inches(7.1), top_y, Inches(5.4), card_h, bg_color=bg_c, border_color=edge_c)
        tb = s2.shapes.add_textbox(Inches(7.3), top_y + Inches(0.12), Inches(5.0), card_h - Inches(0.24))
        tf = tb.text_frame
        tf.word_wrap = True
        p0 = tf.paragraphs[0]
        p0.text = title
        p0.font.size = Pt(13)
        p0.font.bold = True
        p0.font.color.rgb = edge_c
        p0.space_after = Pt(4)
        for line in text.split("\n"):
            lp = tf.add_paragraph()
            lp.text = line
            lp.font.size = Pt(10.5)
            lp.font.color.rgb = DARK_GRAY
            lp.space_after = Pt(2)

    # ==============================================================================
    # 3. SLIDE 3: DATASET & UPES BUS ROUTES (Table + Real Dispatch Proof)
    # ==============================================================================
    s3 = prs.slides.add_slide(blank_layout)
    add_header(s3, "Dataset & Real UPES Bus Fleet Coverage (3,000 Trips)")

    # Overview box
    tb_s3 = s3.shapes.add_textbox(Inches(0.8), Inches(1.3), Inches(11.7), Inches(0.5))
    tf_s3 = tb_s3.text_frame
    p = tf_s3.paragraphs[0]
    p.text = "Built from official UPES transport routes, driver rosters, and student dispatch logs. Tracks GPS distance, departure slot, weather conditions, and ghat bottlenecks across 3,000 trips."
    p.font.size = Pt(11.5)
    p.font.color.rgb = DARK_GRAY

    # Table of buses (Left 7.8 inches)
    t_shape3 = s3.shapes.add_table(7, 4, Inches(0.8), Inches(1.9), Inches(7.8), Inches(5.0))
    tbl3 = t_shape3.table
    tbl3.columns[0].width = Inches(1.3)
    tbl3.columns[1].width = Inches(2.9)
    tbl3.columns[2].width = Inches(1.6)
    tbl3.columns[3].width = Inches(2.0)

    h3 = ["Bus No.", "Route & Stops", "Driver", "Scheduled Time"]
    for j, val in enumerate(h3):
        cell = tbl3.cell(0, j)
        cell.text = val
        cell.fill.solid()
        cell.fill.fore_color.rgb = NAVY
        p = cell.text_frame.paragraphs[0]
        p.font.bold = True
        p.font.size = Pt(10)
        p.font.color.rgb = WHITE
        p.alignment = PP_ALIGN.CENTER

    buses = [
        ("Bus 7717", "Miyawala Express (Mohkampur, Jogiwala, Rispana)", "Sudhir", "85 min (07:00 AM)"),
        ("Bus 7267", "Miyawala - ISBT - FRI City Loop (17 stops)", "Bahadur Singh", "80 min (06:15 AM)"),
        ("Bus 0303", "Jakhan - Rajpur Road Corridor", "Sunny", "70 min (07:00 AM)"),
        ("Bus 6727", "Raipur Corridor", "Pradeep", "75 min (07:00 AM)"),
        ("Bus 8437", "IT Park & Sahastradhara Road", "Monu", "70 min (07:00 AM)"),
        ("Bus 5860", "Chandra Mani Route", "Bittu", "65 min (06:40 AM)")
    ]
    for i, row in enumerate(buses, start=1):
        for j, val in enumerate(row):
            cell = tbl3.cell(i, j)
            cell.text = val
            cell.fill.solid()
            cell.fill.fore_color.rgb = RGBColor(0xFA, 0xFC, 0xFE) if i % 2 == 1 else WHITE
            p = cell.text_frame.paragraphs[0]
            p.font.size = Pt(9.5)
            p.font.color.rgb = DARK_GRAY
            if j == 0:
                p.font.bold = True
                p.font.color.rgb = NAVY
                p.alignment = PP_ALIGN.CENTER

    # Right: Dispatch Snippet Card
    c_wa = create_card(s3, Inches(8.8), Inches(1.9), Inches(3.7), Inches(5.0), bg_color=WHITE, border_color=TEAL)
    tb_wa = s3.shapes.add_textbox(Inches(9.0), Inches(2.05), Inches(3.3), Inches(4.7))
    tf_wa = tb_wa.text_frame
    tf_wa.word_wrap = True

    p = tf_wa.paragraphs[0]
    p.text = "📱 Real Transport Logs"
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = TEAL
    p.space_after = Pt(6)

    wa_text = [
        "Ingested WhatsApp Dispatch Data:",
        "\"Dear students jakhan 0303 sunny 8755010923 Raipur 6727 Pradeep 7455007262 it park 8437 monu 9411548652 Chandra mani 5860 bittu 9760205673 mianwala 7717 sudhir 9105683603\"",
        "\"jakhan 8527 Narendra 9411139750 Raipur 9247 Ravindra 7088857764 mianwala 7267 Jeet bahadur 8755337392\"",
        "Why This Matters:",
        "• 100% genuine data grounding.",
        "• Accurately reflects peak morning congestion."
    ]
    for t in wa_text:
        bp = tf_wa.add_paragraph()
        bp.text = t
        if t.startswith("\""):
            bp.font.size = Pt(8.5)
            bp.font.italic = True
            bp.font.color.rgb = RGBColor(0x4A, 0x55, 0x68)
        elif t.startswith("•") or t.startswith("Why"):
            bp.font.size = Pt(9.5)
            bp.font.bold = True if t.startswith("Why") else False
            bp.font.color.rgb = NAVY if t.startswith("Why") else DARK_GRAY
        else:
            bp.font.size = Pt(10)
            bp.font.bold = True
            bp.font.color.rgb = DARK_GRAY
        bp.space_after = Pt(4)

    # ==============================================================================
    # 4. UPGRADE SLIDE 4: TWO-PART ARCHITECTURE FLOWCHART
    # ==============================================================================
    s4 = prs.slides.add_slide(blank_layout)
    add_header(s4, "System Architecture: Dual-Branch Machine Learning Pipeline")

    # Embedded Flowchart Graphic!
    if os.path.exists(fig_flow):
        s4.shapes.add_picture(fig_flow, Inches(0.8), Inches(1.35), width=Inches(11.7))

    # ==============================================================================
    # 5. SLIDE 5: PREPROCESSING PIPELINE (With 4-Stage Graphic)
    # ==============================================================================
    s5 = prs.slides.add_slide(blank_layout)
    add_header(s5, "Data Preprocessing Pipeline: Cleaning Hill Road Noise (Units I & IV)")

    # Top: Pipeline graphic
    if os.path.exists(fig_prep):
        s5.shapes.add_picture(fig_prep, Inches(1.2), Inches(1.3), width=Inches(10.9))

    # Bottom: 4 detailed explanation cards with icons
    bottom_cards = [
        ("💾 Median GPS Imputation", "Ghat hill terrain causes recurring GPS dropouts.\n• Imputed missing distance via column median.\n• Median prevents mean distortion caused by detours.", NAVY),
        ("🧹 99th %ile Outlier Clip", "Telemetry spikes during extended jams (e.g. 28 choke points).\n• Capped at 99th percentile.\n• Keeps realistic bounds without discarding valid trips.", TEAL),
        ("⚙️ Scaling & Encoding", "• StandardScaler on distance, rain, and choke points.\n• Ordinal encoding for departure slots.\n• VarianceThreshold prunes zero-variance sensor noise.", RGBColor(0x2B, 0x6C, 0xB0)),
        ("⚖️ Training Undersampling", "Severe delays represent ~17.6% of trips.\n• Standard models could trivially guess 'No Delay'.\n• Undersampled exclusively on training data (0 test leak).", ACCENT_GREEN)
    ]

    card_w5 = Inches(2.7)
    gap5 = Inches(0.28)
    for idx, (title, body, col) in enumerate(bottom_cards):
        cur_left = Inches(0.8) + idx * (card_w5 + gap5)
        c = create_card(s5, cur_left, Inches(4.5), card_w5, Inches(2.6), bg_color=CARD_BG, border_color=col)
        tb = s5.shapes.add_textbox(cur_left + Inches(0.12), Inches(4.6), card_w5 - Inches(0.24), Inches(2.4))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = title
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = col
        p.space_after = Pt(4)
        for line in body.split("\n"):
            lp = tf.add_paragraph()
            lp.text = line
            lp.font.size = Pt(9.5)
            lp.font.color.rgb = DARK_GRAY
            lp.space_after = Pt(2)

    # ==============================================================================
    # 6. SLIDE 6: REGRESSION LOSS FUNCTIONS (With MSE vs MAE vs Huber Graphic)
    # ==============================================================================
    s6 = prs.slides.add_slide(blank_layout)
    add_header(s6, "Regression Modeling: Loss Function Benchmark (Units II & V)")

    # Left: Loss plot
    if os.path.exists(fig_loss):
        s6.shapes.add_picture(fig_loss, Inches(0.8), Inches(1.4), width=Inches(6.0))

    # Right: Loss explanation card
    c_l6 = create_card(s6, Inches(7.1), Inches(1.4), Inches(5.4), Inches(5.6))
    tb_l6 = s6.shapes.add_textbox(Inches(7.3), Inches(1.6), Inches(5.0), Inches(5.2))
    tf_l6 = tb_l6.text_frame
    tf_l6.word_wrap = True

    p = tf_l6.paragraphs[0]
    p.text = "Why Huber Loss Wins on Hill Corridors"
    p.font.size = Pt(15)
    p.font.bold = True
    p.font.color.rgb = NAVY
    p.space_after = Pt(8)

    loss_pts = [
        "The MSE Problem: MSE squares prediction errors (e²). When a ghat breakdown causes a 120-min delay, the penalty explodes to 14,400! This single outlier drags the regression slope, causing overpredictions on normal days.",
        "The MAE Limitation: Penalizes errors linearly (|e|), ignoring outlier leverage. However, its derivative is non-smooth at e=0, which causes oscillations near the minimum.",
        "Huber Loss (Selected Solution): Quadratic for small errors (|e| ≤ δ), linear for large errors (|e| > δ). It provides smooth gradient descent near the minimum while capping extreme accident penalties.",
        "Lasso (L1) Sparsification: L1 regularization pushed 11 uninformative encoded features down to zero, acting as an automatic feature pruner."
    ]
    for pt in loss_pts:
        bp = tf_l6.add_paragraph()
        bp.text = "• " + pt
        bp.font.size = Pt(11)
        bp.font.color.rgb = DARK_GRAY
        bp.space_after = Pt(8)

    # ==============================================================================
    # 7. UPGRADE SLIDE 7: REGRESSION METRICS DASHBOARD (Power BI Style KPI Cards)
    # ==============================================================================
    s7 = prs.slides.add_slide(blank_layout)
    add_header(s7, "Regression Dashboard: Model Accuracy & Adequacy (Unit V)")

    # 4 Large Power BI-Style KPI Cards across the top
    kpi_metrics = [
        ("R² SCORE", "91.27%", "Explains 91.3% of travel time variance", NAVY),
        ("ADJUSTED R²", "91.01%", "Confirms model is not overfitting", TEAL),
        ("MEAN ABS ERROR", "4.25 min", "Predictions within ~4 mins of actual arrival", ACCENT_GREEN),
        ("MEAN SQRD ERROR", "30.98", "Low residual variance on test trips", ORANGE)
    ]
    for idx, (lbl, val, desc, col) in enumerate(kpi_metrics):
        cur_left = Inches(0.8) + idx * (Inches(2.7) + Inches(0.28))
        c = create_card(s7, cur_left, Inches(1.3), Inches(2.7), Inches(1.45), bg_color=WHITE, border_color=col)
        tb = s7.shapes.add_textbox(cur_left + Inches(0.1), Inches(1.38), Inches(2.5), Inches(1.25))
        tf = tb.text_frame
        p0 = tf.paragraphs[0]
        p0.text = lbl
        p0.font.size = Pt(8.5)
        p0.font.bold = True
        p0.font.color.rgb = LIGHT_GRAY
        p0.alignment = PP_ALIGN.CENTER
        p1 = tf.add_paragraph()
        p1.text = val
        p1.font.size = Pt(21)  # Huge KPI Number!
        p1.font.bold = True
        p1.font.color.rgb = col
        p1.alignment = PP_ALIGN.CENTER
        p2 = tf.add_paragraph()
        p2.text = desc
        p2.font.size = Pt(8.2)
        p2.font.color.rgb = DARK_GRAY
        p2.alignment = PP_ALIGN.CENTER

    # Bottom Left: Actual vs Predicted Scatter Plot
    if os.path.exists(fig_reg):
        s7.shapes.add_picture(fig_reg, Inches(0.8), Inches(2.95), width=Inches(5.7))

    # Bottom Right: Lasso Feature Weights Plot
    if os.path.exists(fig_lasso):
        s7.shapes.add_picture(fig_lasso, Inches(6.8), Inches(2.95), width=Inches(5.7))

    # ==============================================================================
    # 8. SLIDE 8: CLASSIFICATION & OPTIMIZERS (With Convergence Curves Graphic)
    # ==============================================================================
    s8 = prs.slides.add_slide(blank_layout)
    add_header(s8, "Classification & Optimizer Benchmark (Units II, III & VI)")

    # Left: Paradigm Cards with Icons
    c_p8 = create_card(s8, Inches(0.8), Inches(1.4), Inches(5.6), Inches(5.6))
    tb_p8 = s8.shapes.add_textbox(Inches(1.0), Inches(1.6), Inches(5.2), Inches(5.2))
    tf_p8 = tb_p8.text_frame
    tf_p8.word_wrap = True

    p = tf_p8.paragraphs[0]
    p.text = "Classification Paradigms Tested"
    p.font.size = Pt(15)
    p.font.bold = True
    p.font.color.rgb = NAVY
    p.space_after = Pt(8)

    p8_pts = [
        "📐 Linear & Margin-Based: Logistic Regression (Cross-Entropy Loss) vs. Linear SVM (Hinge Loss). Hinge Loss enforces a safety margin buffer (y·f(x) ≥ 1), ignoring points well outside the margin.",
        "🌲 Ensemble Learning: Random Forest (100 trees bagging) reduces variance; AdaBoost sequentially weights hard-to-classify hill delay samples (92.0% accuracy, 0.973 AUC).",
        "🧠 Deep Learning: Multi-Layer Perceptron (MLP) with hidden layers (16, 8) trained via Backpropagation.",
        "⚡ Optimizer Selection (Unit III): Adam combined momentum direction with adaptive gradient scaling, converging in 18 epochs vs. 60+ epochs for vanilla SGD."
    ]
    for pt in p8_pts:
        bp = tf_p8.add_paragraph()
        bp.text = pt
        bp.font.size = Pt(10.8)
        bp.font.color.rgb = DARK_GRAY
        bp.space_after = Pt(8)

    # Right: Optimizer convergence plot
    if os.path.exists(fig_opt):
        s8.shapes.add_picture(fig_opt, Inches(6.7), Inches(1.4), width=Inches(5.8))
        cap = s8.shapes.add_textbox(Inches(6.7), Inches(6.5), Inches(5.8), Inches(0.35))
        cp = cap.text_frame.paragraphs[0]
        cp.text = "Figure: Loss Convergence over 100 Epochs across SGD, Momentum, RMSprop, and Adam"
        cp.font.size = Pt(9.5)
        cp.font.italic = True
        cp.font.color.rgb = LIGHT_GRAY
        cp.alignment = PP_ALIGN.CENTER

    # ==============================================================================
    # 9. UPGRADE SLIDE 9: CLASSIFICATION RESULTS WITH ONE WINNER CALLOUT
    # ==============================================================================
    s9 = prs.slides.add_slide(blank_layout)
    add_header(s9, "Classification Results & Multi-Model ROC Evaluation")

    # Left Column Container
    c_t9 = create_card(s9, Inches(0.8), Inches(1.4), Inches(5.7), Inches(5.6))

    # 🏆 WINNER CARD (Bright Green / Gold highlight)
    c_winner = s9.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), Inches(1.55), Inches(5.3), Inches(0.85))
    c_winner.fill.solid()
    c_winner.fill.fore_color.rgb = RGBColor(0xE6, 0xF4, 0xEA)
    c_winner.line.color.rgb = ACCENT_GREEN
    c_winner.line.width = Pt(1.5)
    w_tf = c_winner.text_frame
    w_tf.margin_top = Inches(0.08)
    wp0 = w_tf.paragraphs[0]
    wp0.text = "🏆 BEST MODEL: Logistic Regression (Cross-Entropy)"
    wp0.font.size = Pt(11)
    wp0.font.bold = True
    wp0.font.color.rgb = ACCENT_GREEN
    wp0.alignment = PP_ALIGN.CENTER
    wp1 = w_tf.add_paragraph()
    wp1.text = "ROC-AUC = 0.977   |   Accuracy = 93.5%   |   Fast Inference"
    wp1.font.size = Pt(10)
    wp1.font.bold = True
    wp1.font.color.rgb = NAVY
    wp1.alignment = PP_ALIGN.CENTER

    # Table below winner card
    t_shape9 = s9.shapes.add_table(5, 3, Inches(1.0), Inches(2.55), Inches(5.3), Inches(1.9))
    tbl9 = t_shape9.table
    tbl9.columns[0].width = Inches(2.5)
    tbl9.columns[1].width = Inches(1.4)
    tbl9.columns[2].width = Inches(1.4)

    th9 = ["Classifier", "Accuracy", "ROC-AUC"]
    for j, h in enumerate(th9):
        cell = tbl9.cell(0, j)
        cell.text = h
        cell.fill.solid()
        cell.fill.fore_color.rgb = NAVY
        p = cell.text_frame.paragraphs[0]
        p.font.bold = True
        p.font.size = Pt(10)
        p.font.color.rgb = WHITE
        p.alignment = PP_ALIGN.CENTER

    c_rows9 = [
        ("Logistic Regression ★", "93.5%", "0.977"),
        ("Linear SVM (Hinge)", "92.5%", "0.976"),
        ("AdaBoost", "92.0%", "0.973"),
        ("MLP (Adam Optimizer)", "92.2%", "0.970")
    ]
    for i, r in enumerate(c_rows9, start=1):
        for j, val in enumerate(r):
            cell = tbl9.cell(i, j)
            cell.text = val
            cell.fill.solid()
            cell.fill.fore_color.rgb = RGBColor(0xFA, 0xFC, 0xFE) if i % 2 == 1 else WHITE
            p = cell.text_frame.paragraphs[0]
            p.font.size = Pt(9.5)
            p.font.color.rgb = DARK_GRAY
            if j == 0:
                p.font.bold = True
                if "★" in val:
                    p.font.color.rgb = ACCENT_GREEN
            else:
                p.alignment = PP_ALIGN.CENTER

    # Rules below table
    tb_r9 = s9.shapes.add_textbox(Inches(1.0), Inches(4.55), Inches(5.3), Inches(2.3))
    tf_r9 = tb_r9.text_frame
    tf_r9.word_wrap = True

    p = tf_r9.paragraphs[0]
    p.text = "Extracted ID3 Decision Rules (Unit VI)"
    p.font.size = Pt(11.5)
    p.font.bold = True
    p.font.color.rgb = TEAL
    p.space_after = Pt(4)

    rules9 = [
        "Rule 1: IF Departure = 7:00 AM (Peak) AND Rain > 5.2 mm AND Chokepoints > 2 -> Severe Bottleneck = YES (94% confidence).",
        "Rule 2: IF Departure = 6:15 AM (Early) -> Severe Bottleneck = NO regardless of minor rain (Clear roads).",
        "Rule 3: IF Campus = Bidholi AND Rain > 8.0 mm -> Bottleneck probability jumps by +42% compared to Kandoli."
    ]
    for r in rules9:
        bp = tf_r9.add_paragraph()
        bp.text = "• " + r
        bp.font.size = Pt(9)
        bp.font.color.rgb = DARK_GRAY
        bp.space_after = Pt(2)

    # Right: ROC Curve graphic
    if os.path.exists(fig_roc):
        s9.shapes.add_picture(fig_roc, Inches(6.8), Inches(1.4), width=Inches(5.7))

    # ==============================================================================
    # 10. UPGRADE SLIDE 10: TEAM ROLES WITH DISTINCT ROLE ICONS
    # ==============================================================================
    s10 = prs.slides.add_slide(blank_layout)
    add_header(s10, "Team Contributions & Syllabus Alignment (Group of 3)")

    members10 = [
        ("👑 Sparsh Srivastava (Leader)", "Units III & VI Lead\nPipeline Architect & Neural Networks", NAVY, [
            "Coordinated the end-to-end dual-target ML pipeline.",
            "Implemented classification losses (Hinge Loss vs. Binary Cross-Entropy).",
            "Configured Adam vs. RMSprop vs. SGD for MLP neural network training.",
            "Benchmarked 8 classifiers and extracted ID3 decision rules.",
            "Generated multi-model ROC curves & AUC evaluation."
        ]),
        ("📈 Yasharth Jaiswal", "Units II & V Lead\nRegression & Loss Functions", TEAL, [
            "Formulated the continuous travel duration regression task.",
            "Derived closed-form OLS line of best fit mathematically.",
            "Evaluated MSE vs. MAE vs. Huber Loss for accident spike robustness.",
            "Performed Ridge (L2) vs. Lasso (L1) regularization analysis.",
            "Verified model adequacy via R² (0.913) vs. Adjusted R² (0.910)."
        ]),
        ("⚙️ Aayushman Kothari", "Units I & IV Lead\nData Preprocessing & Engineering", ORANGE, [
            "Built 3,000-trip dataset from real UPES bus routes & schedules.",
            "Implemented median imputation for hilly GPS signal dropouts.",
            "Applied 99th-percentile outlier clipping for choke point sensor glitches.",
            "Configured StandardScaler, One-Hot, and Ordinal encoders.",
            "Implemented VarianceThreshold and training-only undersampling."
        ])
    ]

    c_w10 = Inches(3.7)
    gap10 = Inches(0.3)
    for idx, (name, role, col, tasks) in enumerate(members10):
        cur_left = Inches(0.8) + idx * (c_w10 + gap10)
        c = create_card(s10, cur_left, Inches(1.4), c_w10, Inches(5.6), bg_color=CARD_BG, border_color=col)
        
        banner = s10.shapes.add_shape(MSO_SHAPE.RECTANGLE, cur_left, Inches(1.4), c_w10, Inches(0.95))
        banner.fill.solid()
        banner.fill.fore_color.rgb = col
        banner.line.fill.background()
        btf = banner.text_frame
        btf.margin_top = Inches(0.08)
        bp0 = btf.paragraphs[0]
        bp0.text = name
        bp0.font.size = Pt(12)
        bp0.font.bold = True
        bp0.font.color.rgb = WHITE
        bp0.alignment = PP_ALIGN.CENTER
        bp1 = btf.add_paragraph()
        bp1.text = role
        bp1.font.size = Pt(9.5)
        bp1.font.color.rgb = RGBColor(0xDF, 0xEE, 0xF5)
        bp1.alignment = PP_ALIGN.CENTER

        tb = s10.shapes.add_textbox(cur_left + Inches(0.15), Inches(2.45), c_w10 - Inches(0.3), Inches(4.3))
        tf = tb.text_frame
        tf.word_wrap = True
        for t_idx, t in enumerate(tasks):
            p = tf.paragraphs[0] if t_idx == 0 else tf.add_paragraph()
            p.text = "• " + t
            p.font.size = Pt(9.8)
            p.font.color.rgb = DARK_GRAY
            p.space_after = Pt(7)

    # ==============================================================================
    # 11. UPGRADE SLIDE 11: CONCLUSION WITH BIG GREEN SUCCESS CELEBRATION CARDS
    # ==============================================================================
    s11 = prs.slides.add_slide(blank_layout)
    add_header(s11, "Conclusion & Project Deliverables")

    # 4 Big Green Success Cards Across the Top
    success_kpis = [
        ("✅ ERROR ≤ 4.25 MIN", "Arrival Time Accuracy", "Predictions within ~4 mins of actual bus campus arrival", ACCENT_GREEN),
        ("✅ R² = 0.913", "Model Adequacy", "Explains 91.3% of travel time variance across all corridors", NAVY),
        ("✅ ROC-AUC = 0.977", "Severe Delay Alert", "High reliability on bottleneck alerts across test data", TEAL),
        ("✅ 3,000 TRIPS", "Real Fleet Dataset", "Grounded in official UPES bus schedules & driver telemetry", ORANGE)
    ]

    card_w11 = Inches(2.7)
    gap11 = Inches(0.28)
    for idx, (title, sub, body, col) in enumerate(success_kpis):
        cur_left = Inches(0.8) + idx * (card_w11 + gap11)
        c = create_card(s11, cur_left, Inches(1.35), card_w11, Inches(2.1), bg_color=WHITE, border_color=col)
        tb = s11.shapes.add_textbox(cur_left + Inches(0.1), Inches(1.45), card_w11 - Inches(0.2), Inches(1.9))
        tf = tb.text_frame
        tf.word_wrap = True
        p0 = tf.paragraphs[0]
        p0.text = title
        p0.font.size = Pt(12)
        p0.font.bold = True
        p0.font.color.rgb = col
        p0.alignment = PP_ALIGN.CENTER
        p0.space_after = Pt(2)
        p1 = tf.add_paragraph()
        p1.text = sub
        p1.font.size = Pt(10)
        p1.font.bold = True
        p1.font.color.rgb = NAVY
        p1.alignment = PP_ALIGN.CENTER
        p1.space_after = Pt(4)
        p2 = tf.add_paragraph()
        p2.text = body
        p2.font.size = Pt(8.8)
        p2.font.color.rgb = DARK_GRAY
        p2.alignment = PP_ALIGN.CENTER

    # Bottom Split: Summary (Left) & Future Roadmap (Right)
    c_b1 = create_card(s11, Inches(0.8), Inches(3.65), Inches(5.6), Inches(2.4))
    tb_b1 = s11.shapes.add_textbox(Inches(1.0), Inches(3.75), Inches(5.2), Inches(2.2))
    tf_b1 = tb_b1.text_frame
    tf_b1.word_wrap = True
    p = tf_b1.paragraphs[0]
    p.text = "🎯 Academic Deliverables Achieved"
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = NAVY
    p.space_after = Pt(6)
    dels = [
        "100% syllabus alignment across Units I through VI without banned topics.",
        "Solves authentic UPES commute friction across Nanda Ki Chowki and Bidholi ghat.",
        "Dual-target pipeline combining OLS/Huber Regression and Multi-Model Classification."
    ]
    for d in dels:
        bp = tf_b1.add_paragraph()
        bp.text = "• " + d
        bp.font.size = Pt(10)
        bp.font.color.rgb = DARK_GRAY
        bp.space_after = Pt(3)

    c_b2 = create_card(s11, Inches(6.9), Inches(3.65), Inches(5.6), Inches(2.4))
    tb_b2 = s11.shapes.add_textbox(Inches(7.1), Inches(3.75), Inches(5.2), Inches(2.2))
    tf_b2 = tb_b2.text_frame
    tf_b2.word_wrap = True
    p = tf_b2.paragraphs[0]
    p.text = "🚀 Future Implementation Roadmap"
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = TEAL
    p.space_after = Pt(6)
    futs = [
        "Live Bus GPS Feeds: Connect directly to college bus OBD-II / GPS tracker APIs.",
        "Monsoon Radar: Ingest live Dehradun rain radar data for dynamic landslide risk.",
        "Student WhatsApp Bot: Push arrival forecasts to students before 8:00 AM labs."
    ]
    for f in futs:
        bp = tf_b2.add_paragraph()
        bp.text = "• " + f
        bp.font.size = Pt(10)
        bp.font.color.rgb = DARK_GRAY
        bp.space_after = Pt(3)

    # Bottom Green Q&A Banner
    banner_qa = s11.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(6.25), Inches(11.7), Inches(0.75))
    banner_qa.fill.solid()
    banner_qa.fill.fore_color.rgb = ACCENT_GREEN
    banner_qa.line.fill.background()
    bq_tf = banner_qa.text_frame
    p_qa = bq_tf.paragraphs[0]
    p_qa.text = "Thank You! Team Sparsh, Yasharth & Aayushman are Ready for Lab Viva & Code Demo."
    p_qa.font.size = Pt(13)
    p_qa.font.bold = True
    p_qa.font.color.rgb = WHITE
    p_qa.alignment = PP_ALIGN.CENTER

    prs.save(pptx_path)
    print(f"[SUCCESS] Upgraded presentation saved at: {pptx_path}")

    try:
        prs.save(pptx_fallback)
        print(f"[SUCCESS] Also saved to: {pptx_fallback}")
    except Exception as e:
        print(f"[NOTE] Open in PowerPoint, could not overwrite fallback: {e}")

if __name__ == "__main__":
    build_upgraded_deck()
