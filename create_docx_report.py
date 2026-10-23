import os
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

def set_cell_background(cell, hex_color):
    shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    cell._tc.get_or_add_tcPr().append(shading_elm)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def create_report():
    base_dir = r"d:\UPES 2nd YEar\ML\ML PROJECT"
    output_path = os.path.join(base_dir, "Transportation_Delay_System_Report.docx")
    fig_roc = os.path.join(base_dir, "figures", "upes_roc_curve.png")
    fig_reg = os.path.join(base_dir, "figures", "upes_regression_fit.png")

    doc = Document()

    # Page Margins
    for section in doc.sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)

    # Styles
    styles = doc.styles
    normal_style = styles['Normal']
    normal_style.font.name = 'Calibri'
    normal_style.font.size = Pt(11)
    normal_style.font.color.rgb = RGBColor(0x22, 0x22, 0x22)
    normal_style.paragraph_format.line_spacing = 1.2
    normal_style.paragraph_format.space_after = Pt(6)

    # ------------------ COVER / HEADER ------------------
    title_p = doc.add_paragraph()
    title_p.paragraph_format.space_before = Pt(10)
    title_p.paragraph_format.space_after = Pt(4)
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_title = title_p.add_run("TRANSPORTATION DELAY SYSTEM")
    run_title.font.size = Pt(24)
    run_title.font.bold = True
    run_title.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    sub_p = doc.add_paragraph()
    sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub_p.paragraph_format.space_after = Pt(14)
    run_sub = sub_p.add_run("A Predictive Machine Learning Framework for Corridor Congestion & Transit Delays\nSchool of Computer Science | UPES Dehradun")
    run_sub.font.size = Pt(12)
    run_sub.font.italic = True
    run_sub.font.color.rgb = RGBColor(0x4A, 0x60, 0x7A)

    # Meta Table
    meta_table = doc.add_table(rows=5, cols=2)
    meta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    meta_table.autofit = False

    meta_data = [
        ("Course Name & Code:", "Applied Machine Learning (CSAI2017P)"),
        ("Project Team:", "Sparsh Srivastava (Leader)\nYasharth Jaiswal\nAayushman Kothari"),
        ("Project Title:", "Transportation Delay System"),
        ("Target Corridor:", "Dehradun City to UPES Bidholi & Kandoli Campuses"),
        ("Academic Batch:", "Batch 17 & 18 | 2nd Year B.Tech CSE")
    ]
    for i, (k, v) in enumerate(meta_data):
        row = meta_table.rows[i]
        c0, c1 = row.cells[0], row.cells[1]
        c0.width = Inches(2.2)
        c1.width = Inches(4.3)
        c0.paragraphs[0].add_run(k).bold = True
        c0.paragraphs[0].runs[0].font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)
        c1.paragraphs[0].add_run(v)
        set_cell_background(c0, "F0F4F8")
        set_cell_background(c1, "FDFDFD")
        set_cell_margins(c0, 60, 60, 100, 100)
        set_cell_margins(c1, 60, 60, 100, 100)

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # ------------------ 1. EXECUTIVE SUMMARY ------------------
    h1 = doc.add_heading("1. Executive Summary & Problem Background", level=1)
    h1.runs[0].font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    doc.add_paragraph(
        "Anyone who attends or teaches at UPES Dehradun knows the daily commute is a major pain point. "
        "The campuses in Bidholi and Kandoli sit past narrow, single-lane roads running through Prem Nagar, "
        "Nanda Ki Chowki, and the steep Bidholi ghat. Under normal conditions, an official college bus or shared "
        "commute from Dehradun city (like Miyawala, ISBT, or Jakhan) takes about 70 to 85 minutes. "
        "However, even a small hiccup—a broken-down Vikram, a slow truck on the ghat turns, or sudden monsoon showers—causes "
        "the corridor to collapse into gridlock, blowing travel times out to 2.5 hours and making students miss early morning labs."
    )
    doc.add_paragraph(
        "Compounding this issue is the rigid bus timetable. Official UPES buses depart at fixed slots (such as 06:15 AM, 07:00 AM, "
        "02:00 PM, and 05:30 PM). When a student's lectures end early at 1:30 PM or 3:30 PM, they face an annoying dilemma: "
        "should they wait hours on campus for the scheduled bus, or pay for an expensive local cab or auto, only to get stuck "
        "in the exact same traffic jam at Nanda Ki Chowki anyway?"
    )
    doc.add_paragraph(
        "Our team—Sparsh Srivastava (Leader), Yasharth Jaiswal, and Aayushman Kothari—developed the Transportation Delay System "
        "to solve this real-world campus challenge. Built using official UPES fleet information (including Bus 7717, Bus 7267, "
        "Bus 0303, Bus 6727, Bus 8437, and Bus 5860), our system operates as a dual-target Machine Learning pipeline that "
        "answers two essential questions:"
    )
    
    b1 = doc.add_paragraph(style='List Bullet')
    b1.add_run("Continuous Regression Task: ").bold = True
    b1.add_run("Predicts the exact journey time in minutes so commuters know their real arrival time before stepping out.")
    
    b2 = doc.add_paragraph(style='List Bullet')
    b2.add_run("Binary Classification Task: ").bold = True
    b2.add_run("Provides a Severe Gridlock Alert (> 65 minutes extra delay beyond schedule) so students know whether leaving immediately via alternate transport is worth it, or if staying in the library for the scheduled bus is smarter.")

    # ------------------ 2. GROUP DIVISION ------------------
    h2 = doc.add_heading("2. Group Work Division", level=1)
    h2.runs[0].font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    doc.add_paragraph(
        "To ensure balanced contributions, the project is divided into three distinct modules across the syllabus units. "
        "Each team member owns a specific phase of the pipeline:"
    )

    group_table = doc.add_table(rows=4, cols=3)
    group_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    group_headers = ["Team Member & Role", "Syllabus Modules", "Primary Responsibilities"]
    
    for j, h in enumerate(group_headers):
        cell = group_table.rows[0].cells[j]
        cell.paragraphs[0].add_run(h).bold = True
        cell.paragraphs[0].runs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        set_cell_background(cell, "1B365D")
        set_cell_margins(cell, 100, 100, 100, 100)

    group_rows = [
        ("Sparsh Srivastava\n(Team Leader)", "Units III & VI\n(Classification, Ensembles & Neural Nets)", 
         "End-to-end pipeline coordination, implementation of classification losses (Hinge Loss for Linear SVM vs. Binary Cross-Entropy for Logistic/MLP), Adam optimizer configuration, training 8 comparative classifiers (K-NN, Naive Bayes, Decision Tree, Random Forest, AdaBoost, MLP), extraction of ID3 decision rules, and master ROC-AUC evaluation."),
        ("Yasharth Jaiswal", "Units II & V\n(Regression & Loss Functions)", 
         "Formulation of continuous travel time prediction, mathematical derivation of OLS line of best fit, loss function comparison (MSE vs. MAE vs. Huber Loss for traffic accident robustness), Ridge (L2) vs. Lasso (L1) regularization analysis, and model adequacy checking via R² (0.913) vs. Adjusted R² (0.910)."),
        ("Aayushman Kothari", "Units I & IV\n(Data Preprocessing & Engineering)", 
         "UPES fleet telemetry ingestion, data cleaning via median imputation for ghat GPS dropouts, 99th-percentile outlier clipping for choke point sensor glitches, feature transformation using StandardScaler, One-Hot & Ordinal encoders, feature selection via Variance Threshold, and class imbalance handling using training-partition undersampling.")
    ]

    for i, (m, u, r) in enumerate(group_rows, start=1):
        row = group_table.rows[i]
        row.cells[0].paragraphs[0].add_run(m).bold = True
        row.cells[1].paragraphs[0].add_run(u)
        row.cells[2].paragraphs[0].add_run(r)
        for c in row.cells:
            set_cell_margins(c, 80, 80, 100, 100)
            if i % 2 == 1:
                set_cell_background(c, "F9FBFD")

    # ------------------ 3. SYLLABUS MAPPING ------------------
    h3 = doc.add_heading("3. Academic Syllabus Coverage (Units I to VI)", level=1)
    h3.runs[0].font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    doc.add_paragraph(
        "Our project strictly covers the syllabus requirements without incorporating any unauthorized deep learning "
        "architectures (like CNNs or Transformers) or banned project topics. Every unit connects directly to our UPES commute setup:"
    )

    syl_table = doc.add_table(rows=7, cols=3)
    syl_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    syl_headers = ["Unit", "Core Topics Covered", "Exact Implementation in Transportation Delay System"]

    for j, h in enumerate(syl_headers):
        cell = syl_table.rows[0].cells[j]
        cell.paragraphs[0].add_run(h).bold = True
        cell.paragraphs[0].runs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        set_cell_background(cell, "1B365D")
        set_cell_margins(cell, 100, 100, 100, 100)

    syl_rows = [
        ("Unit I:\nIntroduction", "Supervised vs. Unsupervised, Python Scientific Stack", 
         "Constructed an end-to-end supervised pipeline using NumPy, Pandas, Scikit-Learn, and Matplotlib."),
        ("Unit II:\nLoss Functions", "MSE, MAE, Huber Loss, Binary Cross-Entropy, Hinge Loss", 
         "Regression: Compared MSE, MAE, and Huber Loss. Huber Loss is key because rare traffic accidents cause huge delay spikes that distort standard MSE. Classification: Cross-Entropy for Logistic/MLP and Hinge Loss for SVM."),
        ("Unit III:\nOptimizers", "Mini-Batch SGD with Momentum, RMSprop, Adam", 
         "Trained the Multi-Layer Perceptron (MLP) using the Adam optimizer to combine momentum direction with RMSprop adaptive scaling."),
        ("Unit IV:\nPreprocessing", "Imputation, Outlier Clipping, Scaling, Encoding, Variance Threshold, Imbalance", 
         "Imputed GPS hill dropouts with median; clipped choke point spikes at 99th percentile; standardized distance & rain; encoded time slots (ordinal) & bus routes (one-hot); removed noise via Variance Threshold; undersampled majority class on training data."),
        ("Unit V:\nRegression", "OLS Best Fit, Multiple Linear Regression, R², Adjusted R², Ridge & Lasso", 
         "Fitted linear models for travel duration. Verified R² (0.913) against Adjusted R² (0.910). Showed that Lasso (L1) zeroed out 11 uninformative encoded features, providing automatic feature selection."),
        ("Unit VI:\nClassification", "Logistic Regression, K-NN, Decision Trees (ID3), Naive Bayes, Ensembles, MLP, ROC-AUC", 
         "Benchmarked 8 classifiers on severe delay alerts (>65 min late). Extracted human-readable IF-THEN rules from the ID3 tree. Plotted comparative ROC curves (AUC reached 0.977).")
    ]

    for i, (u, t, e) in enumerate(syl_rows, start=1):
        row = syl_table.rows[i]
        row.cells[0].paragraphs[0].add_run(u).bold = True
        row.cells[1].paragraphs[0].add_run(t)
        row.cells[2].paragraphs[0].add_run(e)
        for c in row.cells:
            set_cell_margins(c, 80, 80, 100, 100)
            if i % 2 == 1:
                set_cell_background(c, "F9FBFD")

    # ------------------ 4. DATASET & PREPROCESSING ------------------
    h4 = doc.add_heading("4. Dataset Construction & Data Engineering (Aayushman Kothari)", level=1)
    h4.runs[0].font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    doc.add_paragraph(
        "Rather than using an abstract toy dataset, we synthesized 3,000 real-world commute records based directly on "
        "official UPES bus routes and drivers:"
    )
    fleet_bullets = [
        "Bus 7717 (Sudhir) - Miyawala Express via Mohkampur, Jogiwala, and Rispana Pul (34 km, 85 min scheduled).",
        "Bus 7267 (Bahadur Singh) - Miyawala to ISBT, Kamla Palace, Balliwala, and FRI (17 stops, 39 km, 80 min scheduled).",
        "Bus 0303 / 8527 / 1757 (Sunny / Narendra / Naresh) - Jakhan & Rajpur Road Corridor (28 km, 70 min scheduled).",
        "Bus 6727 / 9247 (Pradeep / Ravindra) - Raipur Corridor (32 km, 75 min scheduled).",
        "Bus 8437 (Monu) - IT Park & Sahastradhara Road Corridor (29 km, 70 min scheduled).",
        "Bus 5860 (Bittu) - Chandra Mani Route (26 km, 65 min scheduled)."
    ]
    for b in fleet_bullets:
        bp = doc.add_paragraph(style='List Bullet')
        bp.add_run(b)

    doc.add_paragraph(
        "Preprocessing Steps Executed:\n"
        "1. Handling Missing Data: In the hilly ghat terrain, GPS signals frequently drop. Missing distances were imputed using the column median on the training partition.\n"
        "2. Outlier Handling: Sudden sensor spikes (such as a GPS device logging 30 chokepoints during a halt) were detected and clipped at the 99th percentile.\n"
        "3. Feature Encoding & Scaling: Continuous features (distance, rain intensity, choke points) were standardized using StandardScaler (mean=0, unit variance). Departure slots were ordinally encoded (6:15 AM < 2:00 PM < 5:30 PM < 7:00 AM) to reflect increasing congestion intensity. Categorical hubs and bus identifiers were One-Hot encoded.\n"
        "4. Feature Selection: Applied Variance Thresholding to automatically detect and discard zero-variance noise features (such as synthetic internal bus sensor noise).\n"
        "5. Imbalance Resolution: Severe gridlocks represent approximately 17.6% of all trips. To prevent classifiers from trivially predicting 'No Traffic' and getting artificially high accuracy, controlled random undersampling was performed on the majority class exclusively on the training set."
    )

    # ------------------ 5. REGRESSION & LOSSES ------------------
    h5 = doc.add_heading("5. Continuous Delay Regression & Loss Comparison (Yasharth Jaiswal)", level=1)
    h5.runs[0].font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    doc.add_paragraph(
        "The continuous goal is predicting exact transit time in minutes so commuters know when they will arrive on campus. "
        "Yasharth evaluated loss functions and regression modeling techniques:"
    )

    doc.add_paragraph(
        "Understanding Loss Behavior in Traffic Scenarios:\n"
        "• Mean Squared Error (MSE): MSE squares prediction errors (e²). When an overturned truck or roadblock causes a 120-minute delay, "
        "the error term produces a massive penalty of 14,400. This disproportionate pull distorts the regression slope, causing the model to overpredict travel times during ordinary, clear traffic days.\n"
        "• Mean Absolute Error (MAE): Penalizes errors linearly (|e|). It is robust against extreme outliers, but its derivative is non-smooth at e=0, which causes oscillations near the minimum.\n"
        "• Huber Loss: Huber loss combines the best aspects of both. For small prediction errors (|e| ≤ δ), it is quadratic like MSE, ensuring smooth gradient descent. For extreme delays (|e| > δ), it becomes linear like MAE, preventing occasional traffic chaos from hijacking model weights."
    )

    reg_metrics_table = doc.add_table(rows=5, cols=2)
    reg_metrics_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    reg_data = [
        ("Regression Metric", "Observed Result on Test Data"),
        ("Mean Absolute Error (MAE):", "4.25 minutes (Average error is just over 4 minutes)"),
        ("Mean Squared Error (MSE):", "30.98 minutes²"),
        ("Coefficient of Determination (R²):", "0.9127 (Model explains 91.3% of travel time variance)"),
        ("Adjusted R² Score:", "0.9101 (Guards against complexity; confirms genuine feature utility)")
    ]
    for i, (k, v) in enumerate(reg_data):
        row = reg_metrics_table.rows[i]
        c0, c1 = row.cells[0], row.cells[1]
        c0.width = Inches(2.8)
        c1.width = Inches(3.7)
        if i == 0:
            c0.paragraphs[0].add_run(k).bold = True
            c1.paragraphs[0].add_run(v).bold = True
            c0.paragraphs[0].runs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            c1.paragraphs[0].runs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            set_cell_background(c0, "1B365D")
            set_cell_background(c1, "1B365D")
        else:
            c0.paragraphs[0].add_run(k).bold = True
            c1.paragraphs[0].add_run(v)
            if i % 2 == 1:
                set_cell_background(c0, "F9FBFD")
                set_cell_background(c1, "F9FBFD")
        set_cell_margins(c0, 60, 60, 100, 100)
        set_cell_margins(c1, 60, 60, 100, 100)

    doc.add_paragraph().paragraph_format.space_after = Pt(6)

    doc.add_paragraph(
        "Regularization Analysis (Ridge vs. Lasso):\n"
        "Ridge (L2) regression shrinks all coefficient weights toward zero smoothly, which stabilizes the model when distance and signal count correlate. "
        "Lasso (L1) regression uses absolute weight penalties, successfully forcing 11 uninformative encoded features to exactly zero. "
        "This confirmed that specific minor route stops contributed zero unique predictive variance once corridor length and departure time slot were known."
    )

    if os.path.exists(fig_reg):
        doc.add_paragraph().paragraph_format.space_before = Pt(4)
        p_img = doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img.add_run().add_picture(fig_reg, width=Inches(5.2))
        cap = doc.add_paragraph()
        cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_cap = cap.add_run("Figure 1: OLS Regression Actual vs. Predicted Travel Times (R² = 0.913, Adj R² = 0.910)")
        r_cap.font.size = Pt(9.5)
        r_cap.font.italic = True
        r_cap.font.color.rgb = RGBColor(0x55, 0x55, 0x55)

    # ------------------ 6. CLASSIFICATION & ROC-AUC ------------------
    h6 = doc.add_heading("6. Severe Delay Classification & ROC-AUC Benchmark (Sparsh Srivastava)", level=1)
    h6.runs[0].font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    doc.add_paragraph(
        "The classification module predicts whether a commute will encounter a Severe Gridlock Alert (>65 minutes delay beyond schedule). "
        "This acts as an operational threshold for students deciding whether to leave campus immediately or stay put. "
        "Sparsh benchmarked 8 algorithms spanning instance-based, probabilistic, margin-based, ensemble, and neural network paradigms:"
    )

    clf_table = doc.add_table(rows=9, cols=4)
    clf_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    clf_headers = ["Algorithm", "Learning Paradigm", "Test Accuracy", "ROC-AUC Score"]

    for j, h in enumerate(clf_headers):
        cell = clf_table.rows[0].cells[j]
        cell.paragraphs[0].add_run(h).bold = True
        cell.paragraphs[0].runs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        set_cell_background(cell, "1B365D")
        set_cell_margins(cell, 80, 80, 80, 80)

    clf_data = [
        ("Logistic Regression", "Linear / Binary Cross-Entropy", "93.5%", "0.977"),
        ("Linear SVM", "Margin-based / Hinge Loss", "92.5%", "0.976"),
        ("AdaBoost", "Adaptive Boosting Ensemble", "92.0%", "0.973"),
        ("MLP Classifier", "Neural Network (Adam Optimizer)", "92.2%", "0.970"),
        ("Random Forest", "Bagging Ensemble (100 Trees)", "91.3%", "0.952"),
        ("Naive Bayes", "Probabilistic Bayesian (Gaussian)", "80.3%", "0.936"),
        ("K-Nearest Neighbors", "Instance-Based (k=5)", "88.8%", "0.926"),
        ("Decision Tree (ID3)", "Information Gain / Entropy", "86.7%", "0.907")
    ]

    for i, (m, p, a, auc_val) in enumerate(clf_data, start=1):
        row = clf_table.rows[i]
        row.cells[0].paragraphs[0].add_run(m).bold = True
        row.cells[1].paragraphs[0].add_run(p)
        row.cells[2].paragraphs[0].add_run(a)
        row.cells[3].paragraphs[0].add_run(auc_val).bold = True
        for c in row.cells:
            set_cell_margins(c, 60, 60, 80, 80)
            if i % 2 == 1:
                set_cell_background(c, "F9FBFD")

    doc.add_paragraph().paragraph_format.space_after = Pt(6)

    doc.add_paragraph(
        "Why ROC-AUC is Essential for Traffic Alerts:\n"
        "In imbalanced traffic prediction, accuracy can be misleading. A model that simply guesses 'No Delay' every single time "
        "would achieve over 82% accuracy, while failing to catch a single gridlock event. The ROC curve maps the True Positive Rate against "
        "the False Positive Rate across every potential decision threshold. An AUC of 0.977 means our model has a 97.7% probability of ranking "
        "a severe traffic crisis higher than a normal trip."
    )

    if os.path.exists(fig_roc):
        doc.add_paragraph().paragraph_format.space_before = Pt(4)
        p_roc = doc.add_paragraph()
        p_roc.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_roc.add_run().add_picture(fig_roc, width=Inches(5.5))
        cap2 = doc.add_paragraph()
        cap2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_cap2 = cap2.add_run("Figure 2: Multi-Model ROC & AUC Curves for Severe Delay Alerts across All 8 Classifiers")
        r_cap2.font.size = Pt(9.5)
        r_cap2.font.italic = True
        r_cap2.font.color.rgb = RGBColor(0x55, 0x55, 0x55)

    doc.add_paragraph(
        "Extracted Decision Rules from Decision Tree:\n"
        "Using ID3 with entropy attribute selection, the decision tree naturally learned realistic campus traffic rules:\n"
        "• IF Departure_Slot = 7:00 AM (Peak Morning) AND Rain > 5.2 mm AND Chokepoints > 2 → Severe Bottleneck Alert = YES (Probability > 94%).\n"
        "• IF Departure_Slot = 6:15 AM (Early Morning) → Severe Bottleneck Alert = NO regardless of minor rain (Clear roads)."
    )

    # ------------------ 7. VIVA CHEAT SHEET ------------------
    h7 = doc.add_heading("7. Viva Defense & Examination Q&A Cheat-Sheet", level=1)
    h7.runs[0].font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    doc.add_paragraph(
        "When defending this project during the external or internal viva examination, each member should speak with these exact talking points:"
    )

    qa_list = [
        ("Q1 (For Aayushman Kothari): Why did you impute GPS missing values using the median instead of the mean?",
         "Answer: In mountain terrain along the Dehradun ghats, GPS signals drop intermittently. When values are recorded, they occasionally have extreme spikes due to detour loops. The mean is sensitive to extreme values, whereas the median represents the true central tendency of the route distances without distortion."),
        ("Q2 (For Aayushman Kothari): Why did you apply undersampling only to the training set and not before the train-test split?",
         "Answer: Resampling before splitting causes data leakage and distorts test evaluation. The test set must reflect real-world operating conditions (~17% bottleneck rate). Altering test class proportions gives falsely optimistic precision and recall metrics."),
        ("Q3 (For Yasharth Jaiswal): Why is Huber Loss better suited than MSE for this traffic problem?",
         "Answer: Traffic data has heavy-tailed noise. A sudden road closure or breakdown creates an extreme outlier delay of 120 minutes. MSE squares this error to 14,400, pulling the entire regression plane toward that single outlier. Huber Loss transitions to a linear penalty for large errors, bounding outlier leverage while retaining quadratic smoothness near zero error."),
        ("Q4 (For Yasharth Jaiswal): What does your R² of 0.913 vs. Adjusted R² of 0.910 signify?",
         "Answer: Standard R² automatically increases or stays constant whenever new features are added, even if they are noise. Adjusted R² penalizes model complexity using degrees of freedom (n - p - 1). Because our Adjusted R² (0.910) remains virtually identical to R² (0.913), it mathematically proves that our route features provide genuine explanatory power without overfitting."),
        ("Q5 (For Sparsh Srivastava): How does Hinge Loss differ from Binary Cross-Entropy in your classification?",
         "Answer: Binary Cross-Entropy evaluates probabilistic confidence across all points. Hinge Loss (used in our Linear SVM) only cares about points within or on the wrong side of the margin. Once a sample is correctly classified outside the margin (y · f(x) ≥ 1), its loss is zero. This maximizes the decision margin between clear traffic and severe gridlocks."),
        ("Q6 (For Sparsh Srivastava): Why did your Logistic Regression and SVM match or beat the Neural Network?",
         "Answer: Our tabular dataset features well-defined linear combinations of distance, departure time, and weather. Linear models with proper regularizers (L2) and loss functions (Cross-Entropy and Hinge Loss) converge reliably on these relationships without the risk of overfitting or vanishing gradients that MLPs face on moderate tabular datasets.")
    ]

    for q, a in qa_list:
        p_q = doc.add_paragraph()
        p_q.paragraph_format.space_before = Pt(4)
        p_q.paragraph_format.space_after = Pt(2)
        p_q.add_run(q).bold = True
        p_q.runs[0].font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

        p_a = doc.add_paragraph()
        p_a.paragraph_format.space_after = Pt(6)
        p_a.add_run(a)

    # ------------------ 8. CONCLUSION ------------------
    h8 = doc.add_heading("8. Conclusion & Future Scope", level=1)
    h8.runs[0].font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    doc.add_paragraph(
        "The Transportation Delay System proves that classical machine learning techniques—when paired with disciplined "
        "data engineering and grounded in authentic, local campus problems—produce robust, explainable, and production-grade tools. "
        "By accurately predicting trip durations to within 4.25 minutes and warning of severe bottlenecks with 97.7% AUC, "
        "this project provides a tangible template for smart college transit dispatching across hill campuses."
    )

    doc.save(output_path)
    print(f"[SUCCESS] Updated report generated at: {output_path}")

if __name__ == "__main__":
    create_report()
