"""
========================================================================================
Project: UPES Dehradun (Bidholi/Kandoli) Corridor Congestion & Delay Predictor
Course: Applied Machine Learning (CSAI2017P)
Dataset Source: upes_bidholi_commute_raw.csv (Real UPES Fleet: 7717, 7267, 0303, etc.)
========================================================================================
"""

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

# Sklearn Modules
from sklearn.model_selection import train_test_split
from sklearn.preprocessing import StandardScaler, OneHotEncoder, OrdinalEncoder
from sklearn.compose import ColumnTransformer
from sklearn.feature_selection import VarianceThreshold

# Regression (Unit V) & Losses (Unit II)
from sklearn.linear_model import LinearRegression, Ridge, Lasso, HuberRegressor
from sklearn.metrics import mean_squared_error, mean_absolute_error, r2_score

# Classification (Unit VI) & Optimizers (Unit III)
from sklearn.linear_model import LogisticRegression, SGDClassifier
from sklearn.neighbors import KNeighborsClassifier
from sklearn.tree import DecisionTreeClassifier, export_text
from sklearn.naive_bayes import GaussianNB
from sklearn.ensemble import RandomForestClassifier, AdaBoostClassifier
from sklearn.neural_network import MLPClassifier
from sklearn.metrics import roc_curve, auc

print("="*75)
print("UPES DEHRADUN OFFICIAL FLEET TRANSIT PIPELINE STARTING...")
print("="*75)

# ==============================================================================
# [MEMBER 1]: DATA INGESTION & PREPROCESSING (UNITS I & IV)
# ==============================================================================
# 1. Load Raw CSV File
csv_path = 'upes_bidholi_commute_raw.csv'
df = pd.read_csv(csv_path)
print(f"Loaded '{csv_path}' successfully. Rows: {df.shape[0]}, Columns: {df.shape[1]}")
print(f"Active Bus Fleets: {list(df['Bus_Number'].unique())}")

# 2. Data Cleaning: Impute Missing GPS Values (Median Strategy)
missing_gps = df['GPS_Distance_KM'].isnull().sum()
print(f"Found {missing_gps} missing GPS records on ghat turns. Applying Median Imputation...")
df['GPS_Distance_KM'] = df['GPS_Distance_KM'].fillna(df['GPS_Distance_KM'].median())

# 3. Data Cleaning: Detect & Clip Outliers (99th percentile)
cap_val = df['Ghat_Choke_Points'].quantile(0.99)
df['Ghat_Choke_Points'] = np.clip(df['Ghat_Choke_Points'], 0, cap_val)

# Separate Features and Targets
feature_cols = ['Bus_Number', 'Route_Corridor', 'Destination_Campus', 'Departure_Slot', 
                'Scheduled_Duration_Min', 'GPS_Distance_KM', 'Monsoon_Rain_MM', 
                'Ghat_Choke_Points', 'Bus_Sensor_Noise']

X = df[feature_cols]
y_reg = df['Actual_Travel_Time_Min']
y_clf = df['Severe_Delay_Alert']

# 4. Train-Test Split (80% Train, 20% Test)
X_train_raw, X_test_raw, y_train_clf, y_test_clf, y_train_reg, y_test_reg = train_test_split(
    X, y_clf, y_reg, test_size=0.2, random_state=42, stratify=y_clf
)

# 5. Transformation: Scaling & Encoding
ord_order = [['Slot_6_15_AM', 'Slot_2_00_PM', 'Slot_5_30_PM', 'Slot_7_00_AM']]
num_cols = ['Scheduled_Duration_Min', 'GPS_Distance_KM', 'Monsoon_Rain_MM', 
            'Ghat_Choke_Points', 'Bus_Sensor_Noise']

preprocessor = ColumnTransformer([
    ('num', StandardScaler(), num_cols),
    ('ord', OrdinalEncoder(categories=ord_order), ['Departure_Slot']),
    ('cat', OneHotEncoder(drop='first'), ['Bus_Number', 'Route_Corridor', 'Destination_Campus'])
])

X_train = preprocessor.fit_transform(X_train_raw)
X_test = preprocessor.transform(X_test_raw)

# 6. Feature Selection: Variance Threshold (Drops Bus_Sensor_Noise)
vt = VarianceThreshold(threshold=0.01)
X_train_sel = vt.fit_transform(X_train)
X_test_sel = vt.transform(X_test)

# 7. Imbalance Handling: Safe Undersampling on Training Partition
y_train_arr = np.array(y_train_clf)
pos_idx = np.where(y_train_arr == 1)[0]
neg_idx = np.where(y_train_arr == 0)[0]

sample_size = min(len(neg_idx), len(pos_idx) * 2)
sampled_neg = np.random.choice(neg_idx, size=sample_size, replace=False)
bal_idx = np.hstack([pos_idx, sampled_neg])

X_train_bal = X_train_sel[bal_idx]
y_train_clf_bal = y_train_arr[bal_idx]

print(f"\n[Member 1] Preprocessing Summary:")
print(f" - Encoded Features: {X_train.shape[1]} -> Selected Features: {X_train_sel.shape[1]}")
print(f" - Balanced Training Samples: {X_train_bal.shape[0]}")

# ==============================================================================
# [MEMBER 2]: TRAVEL TIME REGRESSION & LOSS COMPARISON (UNITS II & V)
# ==============================================================================
print("\n" + "="*75)
print("[Member 2] REGRESSION & LOSS FUNCTION BENCHMARK")
print("="*75)

# OLS Regression (MSE Loss)
ols = LinearRegression().fit(X_train_sel, y_train_reg)
y_pred_ols = ols.predict(X_test_sel)

# Ridge (L2) vs Lasso (L1)
ridge = Ridge(alpha=1.0).fit(X_train_sel, y_train_reg)
lasso = Lasso(alpha=0.2).fit(X_train_sel, y_train_reg)

# Huber Regressor (Huber Loss: Robust against accident delays on Ghat road)
huber = HuberRegressor().fit(X_train_sel, y_train_reg)

# R^2 and Adjusted R^2
n_test, p_feat = X_test_sel.shape
r2 = r2_score(y_test_reg, y_pred_ols)
adj_r2 = 1 - (1 - r2) * (n_test - 1) / (n_test - p_feat - 1)

print(f"OLS Test MSE: {mean_squared_error(y_test_reg, y_pred_ols):.2f}")
print(f"OLS Test MAE: {mean_absolute_error(y_test_reg, y_pred_ols):.2f} minutes")
print(f"OLS R^2 Score: {r2:.4f}")
print(f"OLS Adjusted R^2 Score: {adj_r2:.4f}")
print(f"Lasso Sparsified Coefficients (Set to 0): {np.sum(lasso.coef_ == 0)}")

# ==============================================================================
# [MEMBER 3]: BOTTLENECK CLASSIFICATION & ROC/AUC (UNITS III & VI)
# ==============================================================================
print("\n" + "="*75)
print("[Member 3] CLASSIFICATION BENCHMARK & ROC-AUC CURVES")
print("="*75)

models = {
    'Logistic Regression (Cross-Entropy)': LogisticRegression(),
    'Linear SVM (Hinge Loss)': SGDClassifier(loss='hinge', penalty='l2', max_iter=1000, random_state=42),
    'K-NN (k=5)': KNeighborsClassifier(n_neighbors=5),
    'Decision Tree (ID3/Entropy)': DecisionTreeClassifier(criterion='entropy', max_depth=4, random_state=42),
    'Naive Bayes': GaussianNB(),
    'Random Forest': RandomForestClassifier(n_estimators=100, random_state=42),
    'AdaBoost': AdaBoostClassifier(n_estimators=50, random_state=42),
    'MLP (Adam Optimizer)': MLPClassifier(hidden_layer_sizes=(16, 8), activation='relu', 
                                          solver='adam', max_iter=400, random_state=42)
}

# ------------------------------------------------------------------------------
# GRAPH 1: ROC & AUC CURVES
# ------------------------------------------------------------------------------
plt.figure(figsize=(10, 7))

for name, clf in models.items():
    clf.fit(X_train_bal, y_train_clf_bal)
    
    if hasattr(clf, "predict_proba"):
        scores = clf.predict_proba(X_test_sel)[:, 1]
    else:
        scores = clf.decision_function(X_test_sel)
        
    fpr, tpr, _ = roc_curve(y_test_clf, scores)
    roc_auc = auc(fpr, tpr)
    plt.plot(fpr, tpr, label=f"{name} (AUC = {roc_auc:.3f})")
    print(f"{name:35s} | Accuracy: {clf.score(X_test_sel, y_test_clf):.3f} | AUC: {roc_auc:.3f}")

plt.plot([0, 1], [0, 1], 'k--', label='Chance (0.500)')
plt.xlabel('False Positive Rate (FPR)')
plt.ylabel('True Positive Rate (TPR)')
plt.title('UPES Commute: ROC & AUC Curves for Severe Delay Alerts (> 25 min late)')
plt.legend(loc='lower right', fontsize=8)
plt.grid(True)
plt.tight_layout()

# Save & Display Graph 1
plt.savefig('upes_roc_curve.png', dpi=300)
print("\n[GRAPH 1] Saved as 'upes_roc_curve.png'!")
plt.show()

# ------------------------------------------------------------------------------
# GRAPH 2: REGRESSION FIT (Actual vs. Predicted Travel Time)
# ------------------------------------------------------------------------------
plt.figure(figsize=(8, 6))
plt.scatter(y_test_reg, y_pred_ols, alpha=0.4, color='darkorange', edgecolors='k', s=35, label='Actual Bus Runs')
plt.plot([y_test_reg.min(), y_test_reg.max()], [y_test_reg.min(), y_test_reg.max()], 'b--', lw=2.5, label='Ideal Scheduled Prediction')
plt.xlabel('Actual Commute Time (Minutes)')
plt.ylabel('Predicted Commute Time (Minutes)')
plt.title(f'Unit V: OLS Regression - Actual vs. Predicted (R² = {r2:.3f}, Adj R² = {adj_r2:.3f})')
plt.legend()
plt.grid(True)
plt.tight_layout()

# Save & Display Graph 2
plt.savefig('upes_regression_fit.png', dpi=300)
print("[GRAPH 2] Saved as 'upes_regression_fit.png'!")
plt.show()

# Extract dispatch rules from Decision Tree
dt = models['Decision Tree (ID3/Entropy)']
print("\n[Member 3] Extracted Ghat Bottleneck Decision Rules:")
print(export_text(dt)[:500] + "\n...")

print("\n" + "="*75)
print("PROJECT EXECUTION COMPLETE! ALL UNITS SUCCESSFUL WITH REAL UPES BUSES!")
print("="*75)