# https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=7f635a2b-f22c-4dc0-8a5f-0fb6f0bdd2bb&response_type=code&redirect_uri=https%3A%2F%2Flogin.microsoftonline.com%2Fcommon%2Foauth2%2Fnativeclient&response_mode=query&scope=User.Read%20offline_access%20Files.ReadWrite
import requests
import json

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
