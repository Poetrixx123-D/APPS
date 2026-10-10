"""
Generate synthetic dataset for training the stress prediction ML model.
This creates fake but realistic data since we don't have real personnel data.
"""

import pandas as pd
import random

random.seed(42)

def generate_row():
    sleep_hours = round(random.uniform(3, 9), 1)
    duty_hours = round(random.uniform(6, 18), 1)
    heart_rate = random.randint(60, 105)
    mood_score = random.randint(1, 5)

    # Calculate a score using similar logic to our rule-based version
    score = 0
    if sleep_hours < 5:
        score += 3
    elif sleep_hours < 6.5:
        score += 1.5

    if duty_hours > 12:
        score += 3
    elif duty_hours > 9:
        score += 1.5

    if heart_rate > 90:
        score += 2
    elif heart_rate > 80:
        score += 1

    if mood_score <= 2:
        score += 2
    elif mood_score == 3:
        score += 1

    # Add some randomness so the model has to actually learn patterns
    score += random.uniform(-1, 1)

    if score >= 6:
        stress_level = "High"
    elif score >= 3:
        stress_level = "Medium"
    else:
        stress_level = "Low"

    return {
        "sleep_hours": sleep_hours,
        "duty_hours": duty_hours,
        "heart_rate": heart_rate,
        "mood_score": mood_score,
        "stress_level": stress_level
    }


if __name__ == "__main__":
    rows = [generate_row() for _ in range(1000)]
    df = pd.DataFrame(rows)
    df.to_csv("synthetic_stress_data.csv", index=False)
    print("Dataset generated: synthetic_stress_data.csv")
    print(df["stress_level"].value_counts())