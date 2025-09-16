import requests
import json
import os
from datetime import datetime

# Your Graph API endpoint
url = "https://graph.microsoft.com/v1.0/me/messages?$top=30&$select=subject,bodyPreview,from,toRecipients"

# Put your token here 
token = ""

headers = {
    "Authorization": f"Bearer {token}",
    "Content-Type": "application/json"
}

response = requests.get(url, headers=headers)

if response.status_code == 200:
    data = response.json()
    
    # Get today's date for filename
    today = datetime.now().strftime("%Y-%m-%d")
    
    # Path to Desktop with date in filename
    desktop_path = os.path.join(os.path.expanduser("~"), "Desktop", f"emails_{today}.json")
    
    # Save JSON on Desktop
    with open(desktop_path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=4, ensure_ascii=False)
    
    print(f"✅ Emails downloaded successfully! Saved to {desktop_path}")
else:
    print("❌ Error:", response.status_code, response.text)
