# Framingham_CVD_assesment.py
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from sklearn.impute import KNNImputer
from sklearn.preprocessing import StandardScaler
from sklearn.metrics import classification_report, confusion_matrix, accuracy_score, recall_score
from sklearn.model_selection import train_test_split, GridSearchCV
from imblearn.over_sampling import SMOTE
from sklearn.neighbors import KNeighborsClassifier
from sklearn.tree import DecisionTreeClassifier
from sklearn.linear_model import LogisticRegression
from sklearn.ensemble import RandomForestClassifier, GradientBoostingClassifier, BaggingClassifier
from sklearn.neural_network import MLPClassifier

def perform_eda():
    print("Starting Exploratory Data Analysis (EDA)...\n")
    sns.set_theme(style="whitegrid")

    # Step 1: Load the dataset using the specified raw string path
    file_path = r"C:\Users\HP\Downloads\AML- CapstoneProject\Framingham_CVD assesment\framingham.csv"
    
    try:
        df = pd.read_csv(file_path)
        print("✅ Dataset loaded successfully!\n")
    except FileNotFoundError:
        print(f"❌ Error: Data file not found at {file_path}")
        return

    # Step 2: Data Inspection
    print("--- Dataset Shape ---")
    print(f"Total Rows (Participants): {df.shape[0]}")
    print(f"Total Columns (Features + Target): {df.shape[1]}\n")

    print("--- Data Types and Non-Null Counts ---")
    df.info()
    print("\n")

    # Step 3: Missing Value Analysis
    print("--- Missing Values Report ---")
    missing_data = df.isnull().sum()
    missing_percentage = (missing_data / len(df)) * 100
    missing_df = pd.DataFrame({'Total Missing': missing_data, 'Percentage (%)': missing_percentage})
    print(missing_df[missing_df['Total Missing'] > 0].sort_values(by='Total Missing', ascending=False))
    print("\n")

    # Step 4: Target Variable Assessment (Class Imbalance)
    print("--- Target Variable Distribution (TenYearCHD) ---")
    target_dist = df['TenYearCHD'].value_counts(normalize=True) * 100
    print(target_dist)
    print("\n")

    plt.figure(figsize=(6, 4))
    sns.countplot(x='TenYearCHD', data=df, palette='viridis')
    plt.title('Distribution of 10-Year CHD Risk (0 = No, 1 = Yes)')
    plt.ylabel('Number of Patients')
    plt.show()

    # Step 5: Statistical Summaries & Distributions for Continuous Variables
    print("--- Statistical Summary of Numerical Features ---")
    print(df.describe())
    print("\n")

    continuous_features = ['age', 'totChol', 'sysBP', 'BMI', 'heartRate', 'glucose']
    plt.figure(figsize=(15, 10))
    for i, feature in enumerate(continuous_features, 1):
        if feature in df.columns:
            plt.subplot(2, 3, i)
            # Use 'bins' instead of 'kde' if seaborn version is older, but this should work with recent versions
            sns.histplot(df[feature].dropna(), kde=True, bins=30, color='teal')
            plt.title(f'Distribution of {feature}')
    plt.tight_layout()
    plt.show()

    # Step 6: Correlation Analysis (Heatmap)
    print("Generating Correlation Heatmap...")
    plt.figure(figsize=(14, 10))
    corr_matrix = df.corr()
    sns.heatmap(corr_matrix, annot=True, fmt='.2f', cmap='coolwarm', vmin=-1, vmax=1, square=True, linewidths=0.5)
    plt.title('Correlation Heatmap of Framingham Dataset Features')
    plt.show()

    print("--- Feature Correlation with TenYearCHD ---")
    print(corr_matrix['TenYearCHD'].sort_values(ascending=False))

def impute_missing_data():
    print("Starting Data Preprocessing & KNN Imputation...\n")
    
    file_path = r"C:\Users\HP\Downloads\AML- CapstoneProject\Framingham_CVD assesment\framingham.csv"
    
    try:
        df = pd.read_csv(file_path)
    except FileNotFoundError:
        print(f"❌ Error: Data file not found at {file_path}")
        return None

    # Step 1: Separate Features (X) and Target (y)
    # We do NOT want to impute based on whether they had CHD (to prevent data leakage)
    if 'TenYearCHD' in df.columns:
        X = df.drop('TenYearCHD', axis=1)
        y = df['TenYearCHD']
    else:
        print("❌ Error: Target column 'TenYearCHD' not found in dataset.")
        return None

    print("--- Initial Missing Values in Features (X) ---")
    print(X.isnull().sum()[X.isnull().sum() > 0])
    print("\n")

    # Step 2: Scale the Data
    # KNN uses distance metrics (like Euclidean distance). If we don't scale the data, 
    # variables with large ranges (like totChol) will dominate variables with small ranges (like cigsPerDay).
    scaler = StandardScaler()
    X_scaled = pd.DataFrame(scaler.fit_transform(X), columns=X.columns)

    # Step 3: Apply KNN Imputation
    # n_neighbors=5 looks at the 5 most similar patients to estimate a missing value.
    print("Applying KNN Imputer (n_neighbors=5)...")
    imputer = KNNImputer(n_neighbors=5)
    X_imputed_scaled = pd.DataFrame(imputer.fit_transform(X_scaled), columns=X.columns)
    
    # Optional Step 4: Inverse Transform to keep original interpretability (like age=45, not age=0.2)
    # This is useful specifically if we want to calculate the ASCVD score next using original values!
    X_imputed = pd.DataFrame(scaler.inverse_transform(X_imputed_scaled), columns=X.columns)

    # Step 5: Verify missing values are gone
    print("\n--- Missing Values After KNN Imputation ---")
    print(X_imputed.isnull().sum())
    print("\n✅ Imputation Successful. No missing values remain.")

    # Recombine X and y into a clean DataFrame to use for the next step (ASCVD Calculation)
    df_clean = pd.concat([X_imputed, y.reset_index(drop=True)], axis=1)
    
    return df_clean

def calculate_framingham_risk(row):
    """
    Approximates the 10-year risk of general cardiovascular disease.
    Based loosely on the Framingham Risk Score (D'Agostino et al., 2008)
    Adjusted slightly for the variables available in this specific dataset.
    """
    # Base coefficients (simplified logistic/Cox regression weights for this example)
    # Note: These are illustrative weights. The exact FRS uses specific point systems for age brackets.
    # For a purely academic capstone baseline, we use generalized linear coefficients.
    
    age = row['age']
    totChol = row['totChol']
    sysBP = row['sysBP']
    currentSmoker = row['currentSmoker']
    diabetes = row['diabetes']
    
    # We create a logistic linear combination log-odds score (L)
    # The coefficients differ heavily between male (1) and female (0)
    if row['male'] == 1:
        # Men
        L = (3.06117 * np.log(age) if age > 0 else 0) + \
            (1.12370 * np.log(totChol) if totChol > 0 else 0) + \
            (1.93303 * np.log(sysBP) if sysBP > 0 else 0) + \
            (0.65451 * currentSmoker) + \
            (0.57367 * diabetes) - 23.9802
        # Convert log-odds to probability using baseline survival S0 = 0.88936
        risk = 1 - (0.88936 ** np.exp(L))
    else:
        # Women
        L = (2.32888 * np.log(age) if age > 0 else 0) + \
            (1.20904 * np.log(totChol) if totChol > 0 else 0) + \
            (2.76157 * np.log(sysBP) if sysBP > 0 else 0) + \
            (0.52873 * currentSmoker) + \
            (0.69154 * diabetes) - 26.1931
        # Baseline survival S0 = 0.95012
        risk = 1 - (0.95012 ** np.exp(L))
        
    return risk

def evaluate_baseline_risk(df_clean):
    # Force matplotlib to turn off interactive mode to prevent terminal freezing
    plt.ioff()
    
    print("\n--- Phase 2: Evaluating Traditional Clinical Baseline (ASCVD/FRS Approximation) ---\n")
    print("Calculating 10-year continuous risk probabilities...")
    predicted_probabilities = df_clean.apply(calculate_framingham_risk, axis=1)
    
    # 0.15 represents a 15% risk threshold
    threshold = 0.15 
    print(f"Applying clinical threshold > {threshold*100}% to classify as high-risk (1)...\n")
    
    y_pred_baseline = (predicted_probabilities > threshold).astype(int)
    y_actual = df_clean['TenYearCHD']
    
        # Using explicit string printing to prevent terminal hanging
    report_str = classification_report(y_actual, y_pred_baseline, target_names=['No CVD (0)', 'CVD (1)'])
    print("--- Baseline Clinical Tool Performance ---\n")
    print(str(report_str))
    
    cm = confusion_matrix(y_actual, y_pred_baseline)
    print("\nConfusion Matrix:")
    print("True Negatives (No CVD, Predicted No CVD):", cm[0][0])
    print("False Positives (No CVD, Predicted CVD):", cm[0][1])
    print("False Negatives (Had CVD, Predicted No CVD):", cm[1][0], "  <-- DANGEROUS IN MEDICAL FIELD")
    print("True Positives (Had CVD, Predicted CVD):", cm[1][1], "  <-- What we want to maximize")
    
    accuracy = accuracy_score(y_actual, y_pred_baseline)
    recall = recall_score(y_actual, y_pred_baseline)
    
    print(f"\nBaseline Accuracy: {accuracy*100:.2f}%")
    print(f"Baseline Sensitivity (Recall): {recall*100:.2f}%\n")
    print("In healthcare, if Recall is very low (e.g., < 50%), the traditional tool is missing too many cases.")
    print("Our goal in Phase 3 is to use Machine Learning to beat this score!\n")
    
    # Close any lingering plots
    plt.close('all')

def train_baseline_models(df_clean):
    plt.close('all')
    print("\n=======================================================", flush=True)
    print("--- Phase 3: Optimized ML Baseline (KNN & Decision Tree) ---", flush=True)
    print("=======================================================\n", flush=True)
    
    # 1. Prepare Data
    X = df_clean.drop('TenYearCHD', axis=1)
    y = df_clean['TenYearCHD']
    
    # 70/30 Split
    X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.30, random_state=42, stratify=y)
    
    print("Applying SMOTE to balance the minority class in the training set...", flush=True)
    smote = SMOTE(random_state=42)
    X_train_smote, y_train_smote = smote.fit_resample(X_train, y_train)
    
    # --- MODEL 1: K-Nearest Neighbors (Optimizing K) ---
    print("\n--- Training Model 1: Optimizing K-Nearest Neighbors ---", flush=True)
    knn = KNeighborsClassifier()
    # Test these K values
    knn_params = {'n_neighbors': [3, 5, 7, 9, 11, 15, 21]}
    # We optimize for ROC AUC (distinguishing sick from healthy) instead of pure accuracy
    knn_grid = GridSearchCV(knn, knn_params, cv=5, scoring='roc_auc', n_jobs=-1)
    
    print("Searching for the best K...", flush=True)
    knn_grid.fit(X_train_smote, y_train_smote)
    best_knn = knn_grid.best_estimator_
    print(f"✅ Best KNN Model Found: K = {best_knn.n_neighbors}", flush=True)
    
    y_pred_knn = best_knn.predict(X_test)
    print(f"KNN Accuracy: {accuracy_score(y_test, y_pred_knn)*100:.2f}%", flush=True)
    print(f"KNN Recall:   {recall_score(y_test, y_pred_knn)*100:.2f}%\n", flush=True)
    
    # --- MODEL 2: Decision Tree (Optimizing Alpha) ---
    print("\n--- Training Model 2: Optimizing Decision Tree (ccp_alpha) ---", flush=True)
    dt = DecisionTreeClassifier(random_state=42, class_weight='balanced')
    # Test these Alpha values (0 means no pruning, higher means simpler trees)
    dt_params = {'ccp_alpha': [0.0, 0.001, 0.002, 0.005, 0.01, 0.02, 0.05]}
    dt_grid = GridSearchCV(dt, dt_params, cv=5, scoring='roc_auc', n_jobs=-1)
    
    print("Searching for the best Alpha...", flush=True)
    # Note: Decision Trees don't strictly need SMOTE as much if class_weights='balanced' is used, 
    # but we feed it the raw X_train so it learns true feature relationships natively.
    dt_grid.fit(X_train, y_train)
    best_dt = dt_grid.best_estimator_
    print(f"✅ Best Decision Tree Model Found: Alpha (ccp_alpha) = {best_dt.ccp_alpha:.4f}", flush=True)
    
    y_pred_dt = best_dt.predict(X_test)
    print(f"Decision Tree Accuracy: {accuracy_score(y_test, y_pred_dt)*100:.2f}%", flush=True)
    print(f"Decision Tree Recall:   {recall_score(y_test, y_pred_dt)*100:.2f}%\n", flush=True)
    
    # --- FEATURE IMPORTANCE EXTRACTION (FOR YOUR PPT) ---
    print("=======================================================", flush=True)
    print("--- 🚀 Top Clinical Variables (Feature Importance) 🚀 ---", flush=True)
    print("=======================================================", flush=True)
    
    # Extract importances from the best Decision Tree
    importances = best_dt.feature_importances_
    features = X.columns
    
    # Sort them descending
    indices = np.argsort(importances)[::-1]
    
    for i in range(len(features)):
        # Only print if it actually contributed > 0%
        if importances[indices[i]] > 0:
            print(f"{i+1}. {features[indices[i]]}: {importances[indices[i]]*100:.2f}% importance", flush=True)
    print("\n(Save these for your final PowerPoint!)\n", flush=True)
    print("=======================================================", flush=True)


def train_advanced_models(df_clean):
    plt.close('all')
    print("\n=======================================================", flush=True)
    print("--- Phase 4: Advanced ML Modeling (LogReg, RF, GBM) ---", flush=True)
    print("=======================================================\n", flush=True)
    
    # 1. Prepare Data (Exact same split as Step 4 for fair comparison)
    X = df_clean.drop('TenYearCHD', axis=1)
    y = df_clean['TenYearCHD']
    X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.30, random_state=42, stratify=y)
    
    print("Applying SMOTE to training data...", flush=True)
    smote = SMOTE(random_state=42)
    X_train_smote, y_train_smote = smote.fit_resample(X_train, y_train)

    # --- MODEL 3: Logistic Regression (The Clinical Gold Standard) ---
    print("\n--- Training Model 3: Logistic Regression ---", flush=True)
    # class_weight='balanced' perfectly adjusts the threshold for the 15% minority class
    log_reg = LogisticRegression(max_iter=2000, class_weight='balanced', random_state=42)
    # Note: Logistic Regression scales better on raw X_train when using class_weight
    log_reg.fit(X_train, y_train)
    y_pred_lr = log_reg.predict(X_test)
    
    print(f"Logistic Regression Accuracy: {accuracy_score(y_test, y_pred_lr)*100:.2f}%", flush=True)
    print(f"Logistic Regression Recall:   {recall_score(y_test, y_pred_lr)*100:.2f}%\n", flush=True)

    # --- MODEL 4: Random Forest (Ensemble of Trees) ---
    print("\n--- Training Model 4: Random Forest Classifier ---", flush=True)
    # You use a forest to prevent the heavy overfitting of a single tree
    rf = RandomForestClassifier(n_estimators=100, max_depth=7, class_weight='balanced_subsample', random_state=42)
    rf.fit(X_train, y_train)
    y_pred_rf = rf.predict(X_test)
    
    print(f"Random Forest Accuracy: {accuracy_score(y_test, y_pred_rf)*100:.2f}%", flush=True)
    print(f"Random Forest Recall:   {recall_score(y_test, y_pred_rf)*100:.2f}%\n", flush=True)

    # --- MODEL 5: Gradient Boosting (XGBoost/GBM) ---
    print("\n--- Training Model 5: Gradient Boosting Classifier ---", flush=True)
    # GBM explicitly learns from the errors of the previous trees
    # It doesn't have class_weight inherently, so we train it on the SMOTE dataset!
    gbm = GradientBoostingClassifier(n_estimators=100, learning_rate=0.05, max_depth=3, random_state=42)
    gbm.fit(X_train_smote, y_train_smote)
    y_pred_gbm = gbm.predict(X_test)
    
    print(f"Gradient Boosting Accuracy: {accuracy_score(y_test, y_pred_gbm)*100:.2f}%", flush=True)
    print(f"Gradient Boosting Recall:   {recall_score(y_test, y_pred_gbm)*100:.2f}%\n", flush=True)

    # --- FEATURE IMPORTANCE (RANDOM FOREST) ---
    print("=======================================================", flush=True)
    print("--- 🚀 Top 5 Variables from Random Forest 🚀 ---", flush=True)
    print("=======================================================", flush=True)
    importances_rf = rf.feature_importances_
    features = X.columns
    indices_rf = np.argsort(importances_rf)[::-1]
    
    for i in range(5): # Print top 5 strictly
        print(f"{i+1}. {features[indices_rf[i]]}: {importances_rf[indices_rf[i]]*100:.2f}%", flush=True)
    print("=======================================================\n", flush=True)

def train_engineered_models(df_clean):
    plt.close('all')
    print("\n=======================================================", flush=True)
    print("--- Phase 5: Feature Engineering & Final ANN Model ---", flush=True)
    print("=======================================================\n", flush=True)
    
    # --- 1. FEATURE ENGINEERING ---
    print("Engineering novel clinical features...", flush=True)
    df_eng = df_clean.copy()
    
    # Pulse Pressure: Difference between Systolic and Diastolic pressure (stiffness of aorta)
    df_eng['PulsePressure'] = df_eng['sysBP'] - df_eng['diaBP']
    
    # Mean Arterial Pressure (MAP): Average blood pressure in a single cardiac cycle
    df_eng['MAP'] = df_eng['diaBP'] + (df_eng['PulsePressure'] / 3)
    
    # Age x Cigs: Interaction term (A 60-year-old smoking 20 cigs is worse than a 30-year-old)
    df_eng['Age_x_Cigs'] = df_eng['age'] * df_eng['cigsPerDay']
    
    # Age x Systolic BP: Interaction term for worsening hypertension risk over time
    df_eng['Age_x_sysBP'] = df_eng['age'] * df_eng['sysBP']
    
    print(f"Dataset now has {df_eng.shape[1] - 1} features (up from {df_clean.shape[1] - 1}).\n", flush=True)
    
    X_eng = df_eng.drop('TenYearCHD', axis=1)
    y_eng = df_eng['TenYearCHD']
    
    # --- 2. DATA SCALING (Crucial for Logistic Regression & Neural Networks) ---
    scaler = StandardScaler()
    X_eng_scaled = pd.DataFrame(scaler.fit_transform(X_eng), columns=X_eng.columns)
    
    # 70/30 Split
    X_train, X_test, y_train, y_test = train_test_split(X_eng_scaled, y_eng, test_size=0.30, random_state=42, stratify=y_eng)
    
    print("Applying SMOTE to scaled, engineered training data...", flush=True)
    smote = SMOTE(random_state=42)
    X_train_smote, y_train_smote = smote.fit_resample(X_train, y_train)

    # --- MODEL 1: Re-optimized Decision Tree ---
    print("\n--- Model 1: Decision Tree (Engineered Data) ---", flush=True)
    dt_eng = DecisionTreeClassifier(random_state=42, ccp_alpha=0.005, class_weight='balanced')
    dt_eng.fit(X_train, y_train) # Un-SMOTEd data with balanced weights is often better for simple DTs
    y_pred_dt = dt_eng.predict(X_test)
    
    print(f"Decision Tree Accuracy: {accuracy_score(y_test, y_pred_dt)*100:.2f}%", flush=True)
    print(f"Decision Tree Recall:   {recall_score(y_test, y_pred_dt)*100:.2f}%\n", flush=True)

    # --- MODEL 2: Re-optimized Logistic Regression ---
    print("\n--- Model 2: Logistic Regression (Engineered & Scaled) ---", flush=True)
    lr_eng = LogisticRegression(max_iter=1000, class_weight='balanced', random_state=42)
    lr_eng.fit(X_train, y_train)
    y_pred_lr = lr_eng.predict(X_test)
    
    print(f"Logistic Regression Accuracy: {accuracy_score(y_test, y_pred_lr)*100:.2f}%", flush=True)
    print(f"Logistic Regression Recall:   {recall_score(y_test, y_pred_lr)*100:.2f}%\n", flush=True)

    # --- MODEL 3: Artificial Neural Network (ANN / MLP) ---
    print("\n--- Model 3: Artificial Neural Network (MLPClassifier) ---", flush=True)
    # A single hidden layer with 50 neurons. Alpha is L2 regularization to prevent overfitting.
    ann = MLPClassifier(hidden_layer_sizes=(50,), activation='relu', solver='adam', alpha=0.01, random_state=42, max_iter=1000)
    # ANNs don't have a 'class_weight' parameter in sklearn, so we MUST feed it the SMOTE dataset!
    ann.fit(X_train_smote, y_train_smote)
    y_pred_ann = ann.predict(X_test)
    
    print(f"Neural Network Accuracy: {accuracy_score(y_test, y_pred_ann)*100:.2f}%", flush=True)
    print(f"Neural Network Recall:   {recall_score(y_test, y_pred_ann)*100:.2f}%\n", flush=True)
    print("=======================================================\n", flush=True)

def train_bagging_model(df_clean):
    plt.close('all')
    print("\n=======================================================", flush=True)
    print("--- Phase 7: Ensemble Bagging (Bootstrap Aggregating) ---", flush=True)
    print("=======================================================\n", flush=True)
    
    # 1. Prepare Data
    X = df_clean.drop('TenYearCHD', axis=1)
    y = df_clean['TenYearCHD']
    
    # 70/30 Split
    X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.30, random_state=42, stratify=y)
    
    print("Applying SMOTE to balance the training set...", flush=True)
    smote = SMOTE(random_state=42)
    X_train_smote, y_train_smote = smote.fit_resample(X_train, y_train)
    
    # --- MODEL: Bagged Decision Trees ---
    print("\n--- Training Model: Bagging Classifier (100 Base Decision Trees) ---", flush=True)
    # We use a standard unpruned tree as a base estimator and let Bagging control the variance
    base_tree = DecisionTreeClassifier(random_state=42)
    # n_estimators=100 means it trains 100 independent decision trees on 100 random subsets of data
    bagging_model = BaggingClassifier(estimator=base_tree, n_estimators=100, random_state=42, n_jobs=-1)
    
    bagging_model.fit(X_train_smote, y_train_smote)
    y_pred_bagging = bagging_model.predict(X_test)
    
    acc = accuracy_score(y_test, y_pred_bagging) * 100
    rec = recall_score(y_test, y_pred_bagging) * 100
    
    print(f"Bagging Accuracy: {acc:.2f}%", flush=True)
    print(f"Bagging Recall:   {rec:.2f}%\n", flush=True)
    
    print("--- Impact Analysis of Bagging ---", flush=True)
    print("Compared to our single optimized Decision Tree (50.8% Acc | 75.6% Recall):")
    print(f"1. Accuracy shifted by: {(acc - 50.86):.2f}%")
    print(f"2. Recall shifted by: {(rec - 75.65):.2f}%")
    print("Conclusion: Bagging drastically reduced the Variance/Overfitting (improving accuracy and reducing false alarms), but the 'majority vote' of 100 trees caused it to miss more actual sick patients compared to a single aggressive tree.")
    print("=======================================================\n", flush=True)

# --- UPDATE MAIN EXECUTION BLOCK ---
# Change your bottom code to look exactly like this:
if __name__ == "__main__":
    plt.ioff()
    # 1. Get clean data (Imputation)
    clean_dataframe = impute_missing_data()
    
    if clean_dataframe is not None:
         # 2. See the baseline score again (or comment it out if you are tired of it)
         evaluate_baseline_risk(clean_dataframe)
         
         # 3. Predict using KNN & Decision Trees
         train_baseline_models(clean_dataframe) 
         
         # RUN THE ADVANCED MODELS
         train_advanced_models(clean_dataframe)
         
         # RUN THE ENGINEERED MODELS
         train_engineered_models(clean_dataframe)
         
         # RUN THE BAGGING MODEL
         train_bagging_model(clean_dataframe)