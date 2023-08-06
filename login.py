import requests
import json
import argparse

session = requests.Session()

url = "https://member.asistenai.com/login"
# Parse command-line arguments
parser = argparse.ArgumentParser(description="Login script")
parser.add_argument("-u", "--username", required=True, help="Your username/email")
parser.add_argument("-p", "--password", required=True, help="Your password")
args = parser.parse_args()

# Construct the payload using the provided username and password
payload = f"username={args.username}&password={args.password}&ref=%7BREF%7D&submit="
headers = {
    "authority": "member.asistenai.com",
    "accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8",
    "accept-language": "en-US,en;q=0.7",
    "cache-control": "max-age=0",
    "content-type": "application/x-www-form-urlencoded",
    "cookie": "sec_session_id=i50ladofl4oc1u19qicv590qt5; qurm=22631.c031ef47ae67abad4691b62edb30cd256df39ecb0be77b9b5c84a3c7cede11e72f69da54bbab1c3c3429b482bbd9205f103aeeab7b078a8597c140cf57204197; sec_session_id=olktf7ir6u4lfs5otgt4glj2a9",
    "origin": "https://member.asistenai.com",
    "referer": "https://member.asistenai.com/login",
    "sec-ch-ua": '"Not/A)Brand";v="99", "Brave";v="115", "Chromium";v="115"',
    "sec-ch-ua-mobile": "?0",
    "sec-ch-ua-platform": '"Windows"',
    "sec-fetch-dest": "document",
    "sec-fetch-mode": "navigate",
    "sec-fetch-site": "same-origin",
    "sec-fetch-user": "?1",
    "sec-gpc": "1",
    "upgrade-insecure-requests": "1",
    "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36",
}

response = session.request("POST", url, headers=headers, data=payload)
# save response cookies to file array of object
cookies = session.cookies.get_dict()
list_cookies = []
for key, value in cookies.items():
    list_cookies.append({"name": key, "value": value})
# save cookies to json file cookies/cookies1.json
with open("cookies/cookies1.json", "w") as outfile:
    json.dump(list_cookies, outfile)

print("perbarui cookies1.json sukses")
