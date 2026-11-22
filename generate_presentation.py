import os
import pptx
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

def create_deck():
    base_dir = r"d:\UPES 2nd YEar\ML\ML PROJECT"
    pptx_path = os.path.join(base_dir, "Transportation_Delay_System_Presentation.pptx")
    fig_roc = os.path.join(base_dir, "figures", "upes_roc_curve.png")
    fig_reg = os.path.join(base_dir, "figures", "upes_regression_fit.png")

    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    # Colors
    NAVY = RGBColor(0x1B, 0x36, 0x5D)
    TEAL = RGBColor(0x00, 0x80, 0x80)
    DARK_GRAY = RGBColor(0x22, 0x22, 0x22)
    LIGHT_GRAY = RGBColor(0x55, 0x55, 0x55)
    CARD_BG = RGBColor(0xF4, 0xF7, 0xFA)
    CARD_BORDER = RGBColor(0xD0, 0xD7, 0xDE)
    WHITE = RGBColor(0xFF, 0xFF, 0xFF)
    ORANGE = RGBColor(0xE6, 0x51, 0x00)
    LIGHT_BLUE = RGBColor(0xEB, 0xF3, 0xFA)
    ACCENT_GREEN = RGBColor(0x2E, 0x7D, 0x32)

    def add_header(slide, title_text, category_text="APPLIED MACHINE LEARNING (CSAI2017P)"):
        # Top banner background line or text
        cat_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.7), Inches(0.35))
        tf_cat = cat_box.text_frame
        tf_cat.word_wrap = True
        tf_cat.margin_left = tf_cat.margin_right = tf_cat.margin_top = tf_cat.margin_bottom = 0
        p_cat = tf_cat.paragraphs[0]
        p_cat.text = category_text.upper()
        p_cat.font.size = Pt(10)
        p_cat.font.bold = True
        p_cat.font.color.rgb = TEAL

        title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.7), Inches(11.7), Inches(0.7))
        tf_title = title_box.text_frame
        tf_title.word_wrap = True
        tf_title.margin_left = tf_title.margin_right = tf_title.margin_top = tf_title.margin_bottom = 0
        p_title = tf_title.paragraphs[0]
        p_title.text = title_text
        p_title.font.size = Pt(22)
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

    # ==========================================
    # SLIDE 1: TITLE SLIDE (Dark Navy Theme)
    # ==========================================
    s1 = prs.slides.add_slide(blank_layout)
    bg1 = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
    bg1.fill.solid()
    bg1.fill.fore_color.rgb = NAVY
    bg1.line.fill.background()

    # Subtle accent card
    card_title = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), Inches(1.0), Inches(11.333), Inches(5.5))
    card_title.fill.solid()
    card_title.fill.fore_color.rgb = RGBColor(0x23, 0x43, 0x6E)
    card_title.line.color.rgb = RGBColor(0x3B, 0x62, 0x9B)
    card_title.line.width = Pt(1.5)

    tb = s1.shapes.add_textbox(Inches(1.5), Inches(1.4), Inches(10.333), Inches(4.7))
    tf = tb.text_frame
    tf.word_wrap = True

    p0 = tf.paragraphs[0]
    p0.text = "APPLIED MACHINE LEARNING LAB PROJECT (CSAI2017P)"
    p0.font.size = Pt(12)
    p0.font.bold = True
    p0.font.color.rgb = RGBColor(0x4D, 0xD0, 0xE1)
    p0.space_after = Pt(12)

    p1 = tf.add_paragraph()
    p1.text = "UPES Transportation Delay System"
    p1.font.size = Pt(32)
    p1.font.bold = True
    p1.font.color.rgb = WHITE
    p1.space_after = Pt(8)

    p2 = tf.add_paragraph()
    p2.text = "Predicting Bus Travel Time & Corridor Delays for UPES Dehradun Campuses"
    p2.font.size = Pt(16)
    p2.font.color.rgb = RGBColor(0xCF, 0xD8, 0xDC)
    p2.space_after = Pt(28)

    p3 = tf.add_paragraph()
    p3.text = "PROJECT TEAM (BATCHES 17 & 18):"
    p3.font.size = Pt(11)
    p3.font.bold = True
    p3.font.color.rgb = RGBColor(0xFF, 0xD5, 0x4F)
    p3.space_after = Pt(4)

    p4 = tf.add_paragraph()
    p4.text = "• Sparsh Srivastava (Group Leader)\n• Yasharth Jaiswal\n• Aayushman Kothari"
    p4.font.size = Pt(13)
    p4.font.color.rgb = WHITE
    p4.space_after = Pt(20)

    p5 = tf.add_paragraph()
    p5.text = "School of Computer Science | UPES Dehradun (Bidholi & Kandoli Campuses)"
    p5.font.size = Pt(11)
    p5.font.color.rgb = RGBColor(0x90, 0xA4, 0xAE)

    # ==========================================
    # SLIDE 2: PROBLEM STATEMENT
    # ==========================================
    s2 = prs.slides.add_slide(blank_layout)
    add_header(s2, "Problem Statement: The Dehradun–Bidholi Commute Trap")

    # Left Column: Context & Issues (Cards)
    c1 = create_card(s2, Inches(0.8), Inches(1.5), Inches(6.8), Inches(5.4))
    tb_c1 = s2.shapes.add_textbox(Inches(1.1), Inches(1.7), Inches(6.2), Inches(5.0))
    tf_c1 = tb_c1.text_frame
    tf_c1.word_wrap = True

    p = tf_c1.paragraphs[0]
    p.text = "The Daily Campus Commute Reality"
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = NAVY
    p.space_after = Pt(12)

    bullets = [
        "Unpredictable Travel Times: City-to-campus commutes (ISBT, Miyawala, Jakhan) average 70–85 minutes under normal conditions, but easily blow out to 2.5 hours on bad days.",
        "Severe Narrow Choke Points: The single-lane ghat stretch via Prem Nagar, Nanda Ki Chowki, Pondha, and Bidholi Gates frequently bottlenecks due to stalled trucks or monsoon showers.",
        "Rigid Bus Timetables: College buses leave at fixed slots (6:15 AM, 7:00 AM, 2:00 PM, 5:30 PM). If morning traffic hits, students miss mandatory 8:00 AM labs.",
        "Early Class Dilemma: If lectures finish early (e.g. 1:30 PM), students either wait hours on campus or pay high shared-auto fares, only to get trapped in Nanda Ki Chowki jams anyway."
    ]
    for b in bullets:
        bp = tf_c1.add_paragraph()
        bp.text = "• " + b
        bp.font.size = Pt(12.5)
        bp.font.color.rgb = DARK_GRAY
        bp.space_after = Pt(10)

    # Right Column: Visual Corridor Flow Map
    c2 = create_card(s2, Inches(7.9), Inches(1.5), Inches(4.6), Inches(5.4), bg_color=LIGHT_BLUE, border_color=TEAL)
    tb_c2 = s2.shapes.add_textbox(Inches(8.1), Inches(1.7), Inches(4.2), Inches(5.0))
    tf_c2 = tb_c2.text_frame
    tf_c2.word_wrap = True

    p = tf_c2.paragraphs[0]
    p.text = "Corridor Choke Points"
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = TEAL
    p.space_after = Pt(14)

    steps = [
        ("1. City Departure Hubs", "Miyawala, ISBT, Jakhan, Raipur, Clock Tower", CARD_BG, NAVY),
        ("2. Balliwala & Prem Nagar", "Dense morning city school and market traffic", CARD_BG, NAVY),
        ("3. ⚠️ Nanda Ki Chowki", "Critical bottle-neck: 2-way traffic on narrow road", RGBColor(0xFF, 0xEB, 0xEE), ORANGE),
        ("4. ⚠️ Pondha & Ghat Turns", "Steep turns, slow heavy vehicles, monsoon slips", RGBColor(0xFF, 0xEB, 0xEE), ORANGE),
        ("5. UPES Bidholi & Kandoli", "Campus entry gates & bus terminal unloading", CARD_BG, ACCENT_GREEN)
    ]
    y_pos = Inches(2.3)
    for title, desc, bg, col in steps:
        box = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(8.2), y_pos, Inches(4.0), Inches(0.85))
        box.fill.solid()
        box.fill.fore_color.rgb = bg
        box.line.color.rgb = col
        box.line.width = Pt(1)
        
        btf = box.text_frame
        btf.margin_left = btf.margin_right = Inches(0.15)
        btf.margin_top = Inches(0.08)
        bp0 = btf.paragraphs[0]
        bp0.text = title
        bp0.font.size = Pt(11)
        bp0.font.bold = True
        bp0.font.color.rgb = col
        bp1 = btf.add_paragraph()
        bp1.text = desc
        bp1.font.size = Pt(9.5)
        bp1.font.color.rgb = DARK_GRAY
        y_pos += Inches(0.95)

    # ==========================================
    # SLIDE 3: DATASET & UPES BUS ROUTES
    # ==========================================
    s3 = prs.slides.add_slide(blank_layout)
    add_header(s3, "Dataset & Real UPES Bus Fleet Coverage")

    # Overview text
    tb_s3 = s3.shapes.add_textbox(Inches(0.8), Inches(1.4), Inches(11.7), Inches(0.7))
    tf_s3 = tb_s3.text_frame
    tf_s3.word_wrap = True
    p = tf_s3.paragraphs[0]
    p.text = "We built a 3,000-trip operational dataset modeled on official UPES transport routes and WhatsApp dispatch notices. Features include GPS distance, departure slot, weather conditions, and ghat choke points."
    p.font.size = Pt(13)
    p.font.color.rgb = DARK_GRAY

    # Table of buses
    rows = 7
    cols = 5
    table_shape = s3.shapes.add_table(rows, cols, Inches(0.8), Inches(2.2), Inches(11.7), Inches(4.5))
    table = table_shape.table
    table.columns[0].width = Inches(1.6)
    table.columns[1].width = Inches(3.2)
    table.columns[2].width = Inches(2.0)
    table.columns[3].width = Inches(2.2)
    table.columns[4].width = Inches(2.7)

    headers = ["Bus No.", "Route & Coverage", "Driver", "Scheduled Time", "Corridor Characteristics"]
    for j, h in enumerate(headers):
        cell = table.cell(0, j)
        cell.text = h
        cell.fill.solid()
        cell.fill.fore_color.rgb = NAVY
        p = cell.text_frame.paragraphs[0]
        p.font.bold = True
        p.font.size = Pt(11)
        p.font.color.rgb = WHITE
        p.alignment = PP_ALIGN.CENTER

    bus_data = [
        ("Bus 7717", "Miyawala Express (Mohkampur, Jogiwala, Rispana)", "Sudhir", "85 mins (07:00 AM)", "Fast direct bypass route"),
        ("Bus 7267", "Miyawala - ISBT - FRI City Loop (17 stops)", "Bahadur Singh", "80 mins (06:15 AM)", "Heavy city stop pickup pattern"),
        ("Bus 0303", "Jakhan - Rajpur Road Corridor", "Sunny", "70 mins (07:00 AM)", "Urban market congestion"),
        ("Bus 6727", "Raipur Corridor", "Pradeep", "75 mins (07:00 AM)", "Ring road to Ballupur connection"),
        ("Bus 8437", "IT Park & Sahastradhara Road", "Monu", "70 mins (07:00 AM)", "Commercial morning traffic"),
        ("Bus 5860", "Chandra Mani Route", "Bittu", "65 mins (06:40 AM)", "Short corridor feeder")
    ]

    for i, row in enumerate(bus_data, start=1):
        for j, val in enumerate(row):
            cell = table.cell(i, j)
            cell.text = val
            cell.fill.solid()
            cell.fill.fore_color.rgb = RGBColor(0xFA, 0xFC, 0xFE) if i % 2 == 1 else WHITE
            p = cell.text_frame.paragraphs[0]
            p.font.size = Pt(10)
            p.font.color.rgb = DARK_GRAY
            if j == 0:
                p.font.bold = True
                p.font.color.rgb = NAVY
                p.alignment = PP_ALIGN.CENTER

    # ==========================================
    # SLIDE 4: TWO-PART ML PIPELINE
    # ==========================================
    s4 = prs.slides.add_slide(blank_layout)
    add_header(s4, "System Architecture: Dual-Target Machine Learning Pipeline")

    # Ingestion Card
    c_in = create_card(s4, Inches(0.8), Inches(1.6), Inches(11.7), Inches(1.3), bg_color=LIGHT_BLUE, border_color=TEAL)
    tb_in = s4.shapes.add_textbox(Inches(1.0), Inches(1.75), Inches(11.3), Inches(1.0))
    tf_in = tb_in.text_frame
    p = tf_in.paragraphs[0]
    p.text = "1. Raw Fleet Logs & Preprocessing (3,000 Trips)"
    p.font.size = Pt(14)
    p.font.bold = True
    p.font.color.rgb = TEAL
    p1 = tf_in.add_paragraph()
    p1.text = "Features: Bus ID, Route, Campus Destination, Time Slot, GPS Distance, Monsoon Rain (mm), Ghat Chokepoints.\nCleaned with median imputation, 99th-percentile clipping, One-Hot/Ordinal encoding, and train-only undersampling."
    p1.font.size = Pt(11)
    p1.font.color.rgb = DARK_GRAY

    # Two columns for Task 1 and Task 2
    # Task 1: Regression
    c_t1 = create_card(s4, Inches(0.8), Inches(3.2), Inches(5.6), Inches(3.7))
    tb_t1 = s4.shapes.add_textbox(Inches(1.1), Inches(3.4), Inches(5.0), Inches(3.3))
    tf_t1 = tb_t1.text_frame
    tf_t1.word_wrap = True

    p = tf_t1.paragraphs[0]
    p.text = "Task 1: Continuous Regression (Unit V)"
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = NAVY
    p.space_after = Pt(10)

    reg_pts = [
        "Objective: Predict exact total commute time in minutes.",
        "Loss Evaluation (Unit II): Compared MSE, MAE, and Huber Loss. Huber loss bounds extreme accident delays linearly while keeping quadratic smoothness near zero.",
        "Algorithms: Ordinary Least Squares (OLS), Ridge (L2), and Lasso (L1).",
        "Outcome: Arrival predictions accurate to within ~4.2 minutes of actual time."
    ]
    for pt in reg_pts:
        bp = tf_t1.add_paragraph()
        bp.text = "• " + pt
        bp.font.size = Pt(11.5)
        bp.font.color.rgb = DARK_GRAY
        bp.space_after = Pt(6)

    # Task 2: Classification
    c_t2 = create_card(s4, Inches(6.9), Inches(3.2), Inches(5.6), Inches(3.7))
    tb_t2 = s4.shapes.add_textbox(Inches(7.2), Inches(3.4), Inches(5.0), Inches(3.3))
    tf_t2 = tb_t2.text_frame
    tf_t2.word_wrap = True

    p = tf_t2.paragraphs[0]
    p.text = "Task 2: Severe Delay Classification (Unit VI)"
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = ORANGE
    p.space_after = Pt(10)

    clf_pts = [
        "Objective: Flag Severe Gridlock Alerts (> 65 minutes extra delay beyond schedule).",
        "Loss & Margin (Unit II): Evaluated Binary Cross-Entropy vs. Hinge Loss for margin separation.",
        "Optimizers (Unit III): Benchmarked Adam vs. RMSprop vs. SGD with Momentum on an MLP neural network.",
        "Outcome: AUC of 0.977 across top classifiers, enabling dependable student alerts."
    ]
    for pt in clf_pts:
        bp = tf_t2.add_paragraph()
        bp.text = "• " + pt
        bp.font.size = Pt(11.5)
        bp.font.color.rgb = DARK_GRAY
        bp.space_after = Pt(6)

    # ==========================================
    # SLIDE 5: DATA PREPROCESSING (Units I & IV)
    # ==========================================
    s5 = prs.slides.add_slide(blank_layout)
    add_header(s5, "Data Preprocessing: From Noisy Hill Logs to Clean Features")

    steps_p = [
        ("Step 1: Missing GPS Imputation", 
         "Hilly ghat terrain causes frequent GPS dropouts.\n• Missing distances imputed using column median.\n• Median avoids mean distortion caused by detours.",
         "Median Imputation"),
        ("Step 2: Outlier Clipping", 
         "Sudden telemetry spikes occur during long traffic halts.\n• Choke points capped at the 99th percentile.\n• Keeps realistic bounds without discarding genuine trips.",
         "99th Percentile Cap"),
        ("Step 3: Scaling & Encoding", 
         "Features exist on completely different scales.\n• StandardScaler applied to distance & rain.\n• Ordinal encoding for time slots; One-Hot for buses.\n• VarianceThreshold prunes dead sensor noise.",
         "Z-Score & Encoders"),
        ("Step 4: Imbalance Handling", 
         "Severe gridlocks represent only ~17.6% of trips.\n• Standard models could trivially guess 'No Traffic'.\n• Applied undersampling exclusively to the training set.\n• Preserves real-world distribution in test evaluation.",
         "Train Undersampling")
    ]

    card_w = Inches(2.7)
    gap = Inches(0.28)
    left_start = Inches(0.8)

    for idx, (title, body, tag) in enumerate(steps_p):
        cur_left = left_start + idx * (card_w + gap)
        c = create_card(s5, cur_left, Inches(1.6), card_w, Inches(5.3))
        
        # Tag badge
        badge = s5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cur_left + Inches(0.2), Inches(1.85), card_w - Inches(0.4), Inches(0.4))
        badge.fill.solid()
        badge.fill.fore_color.rgb = NAVY
        badge.line.fill.background()
        btf = badge.text_frame
        bp = btf.paragraphs[0]
        bp.text = tag
        bp.font.size = Pt(10)
        bp.font.bold = True
        bp.font.color.rgb = WHITE
        bp.alignment = PP_ALIGN.CENTER

        tb = s5.shapes.add_textbox(cur_left + Inches(0.2), Inches(2.45), card_w - Inches(0.4), Inches(4.2))
        tf = tb.text_frame
        tf.word_wrap = True
        p0 = tf.paragraphs[0]
        p0.text = title
        p0.font.size = Pt(13)
        p0.font.bold = True
        p0.font.color.rgb = TEAL
        p0.space_after = Pt(10)

        for line in body.split("\n"):
            lp = tf.add_paragraph()
            lp.text = line
            lp.font.size = Pt(10.5)
            lp.font.color.rgb = DARK_GRAY
            lp.space_after = Pt(4)

    # ==========================================
    # SLIDE 6: REGRESSION MODELS & LOSS FUNCTIONS
    # ==========================================
    s6 = prs.slides.add_slide(blank_layout)
    add_header(s6, "Regression Modeling & Loss Functions (Units II & V)")

    # Left Column: Theory & Comparison
    c_reg = create_card(s6, Inches(0.8), Inches(1.5), Inches(5.8), Inches(5.4))
    tb_reg = s6.shapes.add_textbox(Inches(1.0), Inches(1.7), Inches(5.4), Inches(5.0))
    tf_reg = tb_reg.text_frame
    tf_reg.word_wrap = True

    p = tf_reg.paragraphs[0]
    p.text = "Loss Behavior on Hilly Corridors"
    p.font.size = Pt(15)
    p.font.bold = True
    p.font.color.rgb = NAVY
    p.space_after = Pt(8)

    reg_theory = [
        "MSE Flaw: A 120-min accident delay squares to an error penalty of 14,400. This single outlier pulls the entire regression plane, leading to overestimates on normal days.",
        "MAE Robustness: Penalizes errors linearly (|e|), ignoring outlier distortion, but oscillates near zero due to a non-smooth derivative at the minimum.",
        "Huber Loss (Selected): Quadratic for small errors (|e| ≤ δ), linear for large errors (|e| > δ). Bounded outlier leverage with smooth convergence.",
        "Lasso (L1) Sparsity: L1 regularization pushed 11 uninformative encoded features down to zero, acting as an automatic feature pruner."
    ]
    for t in reg_theory:
        bp = tf_reg.add_paragraph()
        bp.text = "• " + t
        bp.font.size = Pt(11.5)
        bp.font.color.rgb = DARK_GRAY
        bp.space_after = Pt(8)

    # Right Column: Actual Plot
    if os.path.exists(fig_reg):
        s6.shapes.add_picture(fig_reg, Inches(6.9), Inches(1.5), width=Inches(5.6))
        # Caption
        cap_box = s6.shapes.add_textbox(Inches(6.9), Inches(6.8), Inches(5.6), Inches(0.4))
        cp = cap_box.text_frame.paragraphs[0]
        cp.text = "Figure 1: OLS Regression Actual vs. Predicted Commute Time (R² = 0.913)"
        cp.font.size = Pt(10)
        cp.font.italic = True
        cp.font.color.rgb = LIGHT_GRAY
        cp.alignment = PP_ALIGN.CENTER

    # ==========================================
    # SLIDE 7: REGRESSION RESULTS & MODEL ADEQUACY
    # ==========================================
    s7 = prs.slides.add_slide(blank_layout)
    add_header(s7, "Regression Results: Model Accuracy & Adequacy Checking")

    # 4 Metric Cards
    metrics = [
        ("R² Score", "0.9127", "Explains 91.3% of travel time variance across all corridors", NAVY),
        ("Adjusted R²", "0.9101", "Virtually identical to R²; proves model is not overfitting", TEAL),
        ("Mean Absolute Error", "4.25 mins", "Arrival predictions are within ~4 mins of actual time", ACCENT_GREEN),
        ("Mean Squared Error", "30.98 mins²", "Reflects low overall variance on test trips", ORANGE)
    ]

    m_w = Inches(2.7)
    m_gap = Inches(0.28)
    for idx, (label, val, desc, col) in enumerate(metrics):
        cur_left = Inches(0.8) + idx * (m_w + m_gap)
        c = create_card(s7, cur_left, Inches(1.6), m_w, Inches(2.6), bg_color=WHITE, border_color=col)
        
        tb = s7.shapes.add_textbox(cur_left + Inches(0.15), Inches(1.8), m_w - Inches(0.3), Inches(2.2))
        tf = tb.text_frame
        tf.word_wrap = True
        
        p0 = tf.paragraphs[0]
        p0.text = label.upper()
        p0.font.size = Pt(10)
        p0.font.bold = True
        p0.font.color.rgb = LIGHT_GRAY
        p0.alignment = PP_ALIGN.CENTER
        
        p1 = tf.add_paragraph()
        p1.text = val
        p1.font.size = Pt(28)
        p1.font.bold = True
        p1.font.color.rgb = col
        p1.alignment = PP_ALIGN.CENTER
        p1.space_after = Pt(8)
        
        p2 = tf.add_paragraph()
        p2.text = desc
        p2.font.size = Pt(10.5)
        p2.font.color.rgb = DARK_GRAY
        p2.alignment = PP_ALIGN.CENTER

    # Bottom Explanatory Card
    c_bot = create_card(s7, Inches(0.8), Inches(4.5), Inches(11.7), Inches(2.4), bg_color=CARD_BG, border_color=CARD_BORDER)
    tb_bot = s7.shapes.add_textbox(Inches(1.1), Inches(4.7), Inches(11.1), Inches(2.0))
    tf_bot = tb_bot.text_frame
    tf_bot.word_wrap = True

    p = tf_bot.paragraphs[0]
    p.text = "Key Takeaways from Regression Analysis"
    p.font.size = Pt(14)
    p.font.bold = True
    p.font.color.rgb = NAVY
    p.space_after = Pt(6)

    takeaways = [
        "Primary Travel Drivers: Route length and departure slot carry the largest positive coefficients (Peak 7:00 AM adds +35 mins base delay; Evening 5:30 PM adds +25 mins).",
        "Campus Differentiation: Trips to Bidholi Campus require +12 extra minutes on average compared to Kandoli Campus due to the steeper ghat section.",
        "Adequacy Checking: Adjusted R² (0.9101) closely tracking R² (0.9127) confirms that our 17 encoded predictors contribute genuine signal rather than artificial fit."
    ]
    for tk in takeaways:
        bp = tf_bot.add_paragraph()
        bp.text = "• " + tk
        bp.font.size = Pt(11)
        bp.font.color.rgb = DARK_GRAY
        bp.space_after = Pt(4)

    # ==========================================
    # SLIDE 8: CLASSIFICATION & OPTIMIZERS
    # ==========================================
    s8 = prs.slides.add_slide(blank_layout)
    add_header(s8, "Classification & Optimizers (Units II, III & VI)")

    # 3 Paradigm Cards
    paradigms = [
        ("Linear & Margin-Based", 
         "Logistic Regression & Linear SVM\n\n• Loss: Binary Cross-Entropy vs. Hinge Loss\n• Hinge Loss enforces a safety margin buffer (y·f(x) ≥ 1); ignores points well outside the margin.\n• Outcome: Fast, highly interpretable decision boundaries on corridor metrics.",
         NAVY),
        ("Ensemble Learning", 
         "Random Forest & AdaBoost\n\n• Bagging (Random Forest): 100 decision trees reduce variance and capture non-linear rain/traffic interactions.\n• Boosting (AdaBoost): Sequentially corrects misclassified hill delay instances.\n• Outcome: 92.0% accuracy with 0.973 ROC-AUC.",
         TEAL),
        ("Neural Networks & Optimizers", 
         "Multi-Layer Perceptron (MLP)\n\n• Architecture: Hidden layers (16, 8) with ReLU activations and Backpropagation.\n• Optimizers (Unit III): Benchmarked Adam vs. RMSprop vs. SGD with Momentum.\n• Adam combined momentum direction with adaptive gradient scaling, converging fastest.",
         ORANGE)
    ]

    p_w = Inches(3.7)
    p_gap = Inches(0.3)
    for idx, (title, body, col) in enumerate(paradigms):
        cur_left = Inches(0.8) + idx * (p_w + p_gap)
        c = create_card(s8, cur_left, Inches(1.5), p_w, Inches(5.4), bg_color=CARD_BG, border_color=col)
        
        tb = s8.shapes.add_textbox(cur_left + Inches(0.2), Inches(1.7), p_w - Inches(0.4), Inches(5.0))
        tf = tb.text_frame
        tf.word_wrap = True
        
        p0 = tf.paragraphs[0]
        p0.text = title
        p0.font.size = Pt(15)
        p0.font.bold = True
        p0.font.color.rgb = col
        p0.space_after = Pt(10)
        
        for line in body.split("\n"):
            lp = tf.add_paragraph()
            lp.text = line
            lp.font.size = Pt(11)
            lp.font.color.rgb = DARK_GRAY
            lp.space_after = Pt(4)

    # ==========================================
    # SLIDE 9: CLASSIFICATION RESULTS & ROC CURVES
    # ==========================================
    s9 = prs.slides.add_slide(blank_layout)
    add_header(s9, "Classification Results & Multi-Model ROC Evaluation")

    # Left: Table + Decision Rules
    c_tbl = create_card(s9, Inches(0.8), Inches(1.5), Inches(5.8), Inches(5.4))
    
    # Table inside left card
    t_shape = s9.shapes.add_table(5, 3, Inches(1.0), Inches(1.7), Inches(5.4), Inches(2.2))
    tbl = t_shape.table
    tbl.columns[0].width = Inches(2.6)
    tbl.columns[1].width = Inches(1.4)
    tbl.columns[2].width = Inches(1.4)

    th = ["Classifier", "Accuracy", "ROC-AUC"]
    for j, h in enumerate(th):
        cell = tbl.cell(0, j)
        cell.text = h
        cell.fill.solid()
        cell.fill.fore_color.rgb = NAVY
        p = cell.text_frame.paragraphs[0]
        p.font.bold = True
        p.font.size = Pt(10)
        p.font.color.rgb = WHITE
        p.alignment = PP_ALIGN.CENTER

    c_rows = [
        ("Logistic Regression", "93.5%", "0.977"),
        ("Linear SVM (Hinge)", "92.5%", "0.976"),
        ("AdaBoost", "92.0%", "0.973"),
        ("MLP (Adam Optimizer)", "92.2%", "0.970")
    ]
    for i, r in enumerate(c_rows, start=1):
        for j, val in enumerate(r):
            cell = tbl.cell(i, j)
            cell.text = val
            cell.fill.solid()
            cell.fill.fore_color.rgb = RGBColor(0xFA, 0xFC, 0xFE) if i % 2 == 1 else WHITE
            p = cell.text_frame.paragraphs[0]
            p.font.size = Pt(10)
            p.font.color.rgb = DARK_GRAY
            if j == 0:
                p.font.bold = True
            else:
                p.alignment = PP_ALIGN.CENTER

    # Decision tree rules box below table
    tb_rules = s9.shapes.add_textbox(Inches(1.0), Inches(4.1), Inches(5.4), Inches(2.6))
    tf_r = tb_rules.text_frame
    tf_r.word_wrap = True
    
    p = tf_r.paragraphs[0]
    p.text = "Extracted ID3 Decision Rules (Unit VI)"
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = TEAL
    p.space_after = Pt(6)

    rules = [
        "Rule 1: IF Departure = 7:00 AM (Peak Morning) AND Rain > 5.2 mm AND Chokepoints > 2 → Severe Bottleneck = YES (94% confidence).",
        "Rule 2: IF Departure = 6:15 AM (Early Morning) → Severe Bottleneck = NO regardless of light rain (Ghat road is clear).",
        "Rule 3: IF Campus = Bidholi AND Rain > 8.0 mm → Bottleneck probability jumps by +42% compared to Kandoli."
    ]
    for r in rules:
        bp = tf_r.add_paragraph()
        bp.text = "• " + r
        bp.font.size = Pt(10)
        bp.font.color.rgb = DARK_GRAY
        bp.space_after = Pt(4)

    # Right: ROC Curve Image
    if os.path.exists(fig_roc):
        s9.shapes.add_picture(fig_roc, Inches(6.9), Inches(1.5), width=Inches(5.6))
        cap_box = s9.shapes.add_textbox(Inches(6.9), Inches(6.8), Inches(5.6), Inches(0.4))
        cp = cap_box.text_frame.paragraphs[0]
        cp.text = "Figure 2: Multi-Model ROC & AUC Curves (All 8 Classifiers)"
        cp.font.size = Pt(10)
        cp.font.italic = True
        cp.font.color.rgb = LIGHT_GRAY
        cp.alignment = PP_ALIGN.CENTER

    # ==========================================
    # SLIDE 10: TEAM CONTRIBUTIONS & SYLLABUS
    # ==========================================
    s10 = prs.slides.add_slide(blank_layout)
    add_header(s10, "Team Contributions & Syllabus Alignment")

    members = [
        ("Sparsh Srivastava (Leader)", 
         "Units III & VI Focus", 
         NAVY,
         [
             "Designed & coordinated the dual-target ML pipeline.",
             "Implemented classification losses (Hinge Loss for Linear SVM vs. Binary Cross-Entropy).",
             "Configured Adam, RMSprop, and SGD with Momentum for MLP neural network training.",
             "Benchmarked 8 classifiers and extracted ID3 decision rules.",
             "Generated comparative ROC curves and AUC evaluation."
         ]),
        ("Yasharth Jaiswal", 
         "Units II & V Focus", 
         TEAL,
         [
             "Formulated the continuous travel duration regression task.",
             "Derived closed-form OLS line of best fit mathematically.",
             "Evaluated MSE vs. MAE vs. Huber Loss for accident spike robustness.",
             "Performed Ridge (L2) vs. Lasso (L1) regularization analysis.",
             "Verified model adequacy via R² (0.913) vs. Adjusted R² (0.910)."
         ]),
        ("Aayushman Kothari", 
         "Units I & IV Focus", 
         ORANGE,
         [
             "Synthesized 3,000-trip dataset from real UPES bus routes & schedules.",
             "Implemented median imputation for hilly GPS signal dropouts.",
             "Applied 99th-percentile outlier clipping for sensor choke point spikes.",
             "Configured StandardScaler, One-Hot, and Ordinal encoders.",
             "Implemented VarianceThreshold and training-only undersampling."
         ])
    ]

    card_w = Inches(3.7)
    gap = Inches(0.3)
    for idx, (name, role, col, tasks) in enumerate(members):
        cur_left = Inches(0.8) + idx * (card_w + gap)
        c = create_card(s10, cur_left, Inches(1.5), card_w, Inches(5.4), bg_color=CARD_BG, border_color=col)
        
        # Header banner
        banner = s10.shapes.add_shape(MSO_SHAPE.RECTANGLE, cur_left, Inches(1.5), card_w, Inches(0.9))
        banner.fill.solid()
        banner.fill.fore_color.rgb = col
        banner.line.fill.background()
        
        btf = banner.text_frame
        btf.margin_top = Inches(0.12)
        bp0 = btf.paragraphs[0]
        bp0.text = name
        bp0.font.size = Pt(13)
        bp0.font.bold = True
        bp0.font.color.rgb = WHITE
        bp0.alignment = PP_ALIGN.CENTER
        
        bp1 = btf.add_paragraph()
        bp1.text = role
        bp1.font.size = Pt(10.5)
        bp1.font.color.rgb = RGBColor(0xDF, 0xEE, 0xF5)
        bp1.alignment = PP_ALIGN.CENTER

        tb = s10.shapes.add_textbox(cur_left + Inches(0.2), Inches(2.55), card_w - Inches(0.4), Inches(4.2))
        tf = tb.text_frame
        tf.word_wrap = True
        
        for t_idx, t in enumerate(tasks):
            p = tf.paragraphs[0] if t_idx == 0 else tf.add_paragraph()
            p.text = "• " + t
            p.font.size = Pt(10.5)
            p.font.color.rgb = DARK_GRAY
            p.space_after = Pt(8)

    # ==========================================
    # SLIDE 11: CONCLUSION & FUTURE SCOPE
    # ==========================================
    s11 = prs.slides.add_slide(blank_layout)
    add_header(s11, "Conclusion & Future Enhancements")

    # Left: What We Achieved
    c_ach = create_card(s11, Inches(0.8), Inches(1.5), Inches(5.6), Inches(5.4), bg_color=LIGHT_BLUE, border_color=TEAL)
    tb_ach = s11.shapes.add_textbox(Inches(1.1), Inches(1.7), Inches(5.0), Inches(5.0))
    tf_ach = tb_ach.text_frame
    tf_ach.word_wrap = True

    p = tf_ach.paragraphs[0]
    p.text = "Project Summary & Achievements"
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = TEAL
    p.space_after = Pt(12)

    achievements = [
        "Real-World Problem: Grounded directly in UPES Dehradun commute friction across Nanda Ki Chowki and Bidholi ghat single-lane bottlenecks.",
        "Accurate Arrival Time (Unit V): OLS regression with Huber Loss estimates commute duration within ~4.25 minutes of actual arrival (R² = 0.913).",
        "Reliable Safety Alerts (Unit VI): Top classifiers (Logistic, SVM, AdaBoost, MLP) achieve 0.977 ROC-AUC, providing dependable gridlock alerts.",
        "Complete Syllabus Coverage: Successfully applied all 6 units without forbidden topics or unsupported deep learning architectures."
    ]
    for ach in achievements:
        bp = tf_ach.add_paragraph()
        bp.text = "• " + ach
        bp.font.size = Pt(11.5)
        bp.font.color.rgb = DARK_GRAY
        bp.space_after = Pt(10)

    # Right: Next Steps / Future Scope
    c_fut = create_card(s11, Inches(6.9), Inches(1.5), Inches(5.6), Inches(5.4))
    tb_fut = s11.shapes.add_textbox(Inches(7.2), Inches(1.7), Inches(5.0), Inches(5.0))
    tf_fut = tb_fut.text_frame
    tf_fut.word_wrap = True

    p = tf_fut.paragraphs[0]
    p.text = "Future Implementation Roadmap"
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = NAVY
    p.space_after = Pt(12)

    future = [
        "Live Bus GPS Telematics: Connect pipeline to live college bus OBD-II / GPS tracker feeds for continuous real-time predictions.",
        "Weather API Integration: Ingest live Dehradun rain radar data to dynamically update monsoon landslide risk scores.",
        "WhatsApp / App Dispatch Bot: Automatically push arrival estimates and gridlock warnings to students before 8:00 AM class departures.",
        "Smart Bus Rescheduling: Enable UPES transport dispatchers to dynamically reroute or stagger bus departures based on predicted bottleneck severity."
    ]
    for fut in future:
        bp = tf_fut.add_paragraph()
        bp.text = "• " + fut
        bp.font.size = Pt(11.5)
        bp.font.color.rgb = DARK_GRAY
        bp.space_after = Pt(10)

    # Thank you footer box
    ty_box = s11.shapes.add_textbox(Inches(6.9), Inches(5.8), Inches(5.6), Inches(0.8))
    tf_ty = ty_box.text_frame
    pt = tf_ty.paragraphs[0]
    pt.text = "Thank You! Ready for Lab Viva & Code Demo."
    pt.font.size = Pt(13)
    pt.font.bold = True
    pt.font.color.rgb = ACCENT_GREEN
    pt.alignment = PP_ALIGN.CENTER

    prs.save(pptx_path)
    print(f"[SUCCESS] Presentation generated at: {pptx_path}")

if __name__ == "__main__":
    create_deck()
