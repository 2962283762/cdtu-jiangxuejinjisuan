import requests
url = "https://youtube.com"
response = requests.get(url)
print(response.status_code)