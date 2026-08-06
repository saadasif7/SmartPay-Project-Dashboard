import pandas as pd
import requests
from datetime import datetime


def load_data():
    df = pd.read_excel("data/projects.xlsx", engine="openpyxl")

    # Remove extra spaces from column names
    df.columns = df.columns.str.strip()

    return df


# ==========================================
# WEATHER FUNCTION
# ==========================================

def get_weather():
    API_KEY = "60486f002d3387d60abe0f3263bdfe0d"

    city = "Karachi"

    url = f"https://api.openweathermap.org/data/2.5/weather?q={city}&appid={API_KEY}&units=metric"

    try:
        response = requests.get(url, timeout=5).json()

        temp = round(response["main"]["temp"])
        weather = response["weather"][0]["main"]

        return city, temp, weather

    except:
        return "Karachi", "--", "Unavailable"