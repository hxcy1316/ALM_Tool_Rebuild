import requests
from requests.auth import HTTPBasicAuth

url = r"http://15.83.240.100/qcbin/rest/is-authenticated"
username = "chen.si"
password = "P@ssw0rd"
cookies = dict()
headers = {}


r = requests.get(url + "?login-form-required=y", auth=HTTPBasicAuth(username, password), headers=headers)
print(r.status_code, r.text)
