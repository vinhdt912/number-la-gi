import requests
import os
from dotenv import load_dotenv
load_dotenv()


SHORT_LIVED_ACCESS_TOKEN = os.getenv("SHORT_LIVED_TOKEN")
APP_SECRET = os.getenv("CLIENT_SECRET")

url = "https://graph.threads.net/access_token"

params = {
    "grant_type": "th_exchange_token",
    "client_secret": APP_SECRET,
    "access_token": SHORT_LIVED_ACCESS_TOKEN
}

response = requests.get(url, params=params)

if response.ok:
    data = response.json()

    print("Long-lived access token:")
    print(data.get("access_token"))

    print("\nToken type:")
    print(data.get("token_type"))

    print("\nExpires in:")
    print(data.get("expires_in"))

    # Convert seconds -> days
    expires_in = data.get("expires_in")

    if expires_in:
        print(f"\nExpires in approximately: {expires_in / 86400:.2f} days")
else:
    print("Error:")
    print(response.status_code)
    print(response.text)