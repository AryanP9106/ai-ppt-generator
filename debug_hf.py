import requests

API_URL = "https://api-inference.huggingface.co/models/black-forest-labs/FLUX.1-schnell"
# PASTE YOUR TOKEN BELOW
HEADERS = {"Authorization": "Bearer YOUR_TOKEN_HERE"}

def test_api():
    print("Testing connection to Hugging Face...")
    response = requests.post(API_URL, headers=HEADERS, json={"inputs": "A futuristic city"})
    
    print(f"Status Code: {response.status_code}")
    print(f"Response Content: {response.text}")

    if response.status_code == 401:
        print("❌ ERROR: Your Token is invalid or has expired.")
    elif response.status_code == 503:
        print("⏳ ERROR: The model is currently loading. Wait 2 minutes and try again.")
    elif response.status_code == 429:
        print("🚫 ERROR: You've hit the Rate Limit for the free tier.")
    elif response.status_code == 200:
        print("✅ SUCCESS: The API is working perfectly!")

if __name__ == "__main__":
    test_api()