import requests
import json

url = "http://localhost:8080/api/ai/underwrite/" # Make sure port matches your server!

files = {
    'om_file': open('dummy.pdf', 'rb'),
    't12_file': open('dummy.pdf', 'rb'), 
    'rent_roll_file': open('dummy_rent_roll.xlsx', 'rb')
}

print("Sending files to AI Underwriter... this may take 10-20 seconds...")

try:
    response = requests.post(url, files=files)
    
    print(f"\nStatus Code: {response.status_code}")
    if response.status_code == 200:
        data = response.json()
        print("SUCCESS! Qwen extracted the following metrics:")
        print("-" * 60)
        # Pretty print the JSON
        print(json.dumps(data['metrics'], indent=4))
        print("-" * 60)
    else:
        print("ERROR:", response.text)
        
except Exception as e:
    print(f"Request failed: {e}")
finally:
    for f in files.values():
        f.close()