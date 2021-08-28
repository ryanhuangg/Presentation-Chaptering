# https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=7f635a2b-f22c-4dc0-8a5f-0fb6f0bdd2bb&response_type=code&redirect_uri=https%3A%2F%2Flogin.microsoftonline.com%2Fcommon%2Foauth2%2Fnativeclient&response_mode=query&scope=User.Read%20offline_access%20Files.ReadWrite
import requests
import json
import os
import glob
import time
import os

# Build the POST parameters
f = open("token.txt", "r")
token = f.read()
params = {
    'grant_type': 'refresh_token',
    'client_id': '7f635a2b-f22c-4dc0-8a5f-0fb6f0bdd2bb',
    'refresh_token': str(token)
}
f.close()
response = requests.post('https://login.microsoftonline.com/common/oauth2/v2.0/token', data=params)

access_token = response.json()['access_token']
new_refresh_token = response.json()['refresh_token']
print(access_token)
print(new_refresh_token)
f = open("token.txt", "w")
f.write(new_refresh_token)
f.close()

time.sleep(5)

header = {'Authorization': 'Bearer ' + access_token}


response = requests.get('https://graph.microsoft.com/v1.0/me/drive/root:' +
                        '/Test'  + ':/children', headers=header)
d = response.json()
a = d["value"]
download_list = []
for i in a:
    download_list.append(i.get('name'))

print(download_list)

# Download the file
for j in download_list:
    response = requests.get('https://graph.microsoft.com/v1.0/me/drive/root:' +
                        '/Test' + '/' + j + ':/content', headers=header)
    with open(j, 'wb') as file:
        file.write(response.content)

os.system('python batch.py')

myfiles = glob.glob('*.txt')
print(myfiles)

for i in myfiles:
    if i == "token.txt" or i == "token_backup.txt" or i == "requirements.txt":
        myfiles.remove(i)


if "token.txt" in myfiles:
    myfiles.remove("token.txt")

print(myfiles)

for j in myfiles:
    if j != "token.txt":
        data = open(j, 'rb').read()
        response = requests.put('https://graph.microsoft.com/v1.0/me/drive/root:' +
                        '/Test' + '/' + j + ':/content', data=data, headers=header)