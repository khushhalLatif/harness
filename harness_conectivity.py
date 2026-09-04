import requests

API_KEY = "YOUR_TOKEN"
ACCOUNT_ID = "fgDto6qoTT6ctfZS9eWbEw"
headers = {"x-api-key": API_KEY}

# Test 1: is the token valid at all, with no org/project involved?
r1 = requests.get("https://app.harness.io/gateway/ng/api/organizations",
                   headers=headers, params={"accountIdentifier": ACCOUNT_ID})
print("Orgs call:", r1.status_code, r1.text[:400])

# Test 2: does the token see the "Fiserv" org specifically?
r2 = requests.get("https://app.harness.io/gateway/ng/api/projects",
                   headers=headers, params={"accountIdentifier": ACCOUNT_ID, "orgIdentifier": "Fiserv"})
print("Projects call:", r2.status_code, r2.text[:400])
