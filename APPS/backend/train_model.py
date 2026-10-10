"""
Train a Random Forest model on the synthetic dataset and save it as a .pkl file
so the Flask app can use it later for real predictions.
"""

import pandas as pd
from sklearn.model_selection import train_test_split
from sklearn.ensemble import RandomForestClassifier
from sklearn.metrics import accuracy_score, classification_report
import joblib

# 1. Load the dataset
df = pd.read_csv("synthetic_stress_data.csv")

# 2. Features (inputs) and label (output)
X = df[["sleep_hours", "duty_hours", "heart_rate", "mood_score"]]
y = df["stress_level"]

# 3. Split into training data (80%) and testing data (20%)
X_train, X_test, y_train, y_test = train_test_split(
    X, y, test_size=0.2, random_state=42
)

# 4. Create and train the Random Forest model
model = RandomForestClassifier(n_estimators=100, random_state=42)
model.fit(X_train, y_train)

# 5. Test the model's accuracy
y_pred = model.predict(X_test)
accuracy = accuracy_score(y_test, y_pred)

print(f"Model Accuracy: {accuracy * 100:.2f}%")
print("\nDetailed Report:")
print(classification_report(y_test, y_pred))

# 6. Save the trained model to a file
joblib.dump(model, "stress_model.pkl")
print("\nModel saved as stress_model.pkl")