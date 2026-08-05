import requests

# The endpoint URL
url = "http://localhost:8080/api/ai/upload/"

# Open the files in binary read mode ('rb')
# Update these filenames if your files are named differently!
files = {
    'om_file': open('dummy.pdf', 'rb'),
    't12_file': open('dummy.pdf', 'rb'), # Using the same dummy PDF twice is fine for testing
    'rent_roll_file': open('dummy_rent_roll.xlsx', 'rb')
}

print("Sending files to Django... please wait...")

try:
    # Send the POST request
    response = requests.post(url, files=files)
    
    print(f"\nStatus Code: {response.status_code}")
    if response.status_code == 200:
        print("SUCCESS! Here is the extracted Excel data preview:")
        print("-" * 50)
        print(response.json()['extracted_data']['rent_roll_preview'])
        print("-" * 50)
    else:
        print("ERROR:", response.text)
        
except Exception as e:
    print(f"Request failed: {e}")
finally:
    # Always close the files!
    for f in files.values():
        f.close()