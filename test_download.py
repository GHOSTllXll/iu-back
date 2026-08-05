import requests

url = "http://localhost:8000/api/ai/underwrite/download/"

files = {
    'om_file': open('dummy.pdf', 'rb'),
    't12_file': open('dummy.pdf', 'rb'), 
    'rent_roll_file': open('dummy_rent_roll.xlsx', 'rb')
}

print("Downloading Excel file...")

try:
    response = requests.post(url, files=files, proxies={'http': None, 'https': None}, verify=False)
    
    if response.status_code == 200:
        with open('Underwriting_Model.xlsx', 'wb') as f:
            f.write(response.content)
        print("✅ SUCCESS! File saved as 'Underwriting_Model.xlsx'")
    else:
        print("❌ ERROR:", response.text)
        
except Exception as e:
    print(f" Request failed: {e}")
finally:
    for f in files.values():
        f.close()