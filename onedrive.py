# https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=7f635a2b-f22c-4dc0-8a5f-0fb6f0bdd2bb&response_type=code&redirect_uri=https%3A%2F%2Flogin.microsoftonline.com%2Fcommon%2Foauth2%2Fnativeclient&response_mode=query&scope=User.Read%20offline_access%20Files.ReadWrite
import requests

# Build the POST parameters
params = {
    'grant_type': 'refresh_token',
    'client_id': '7f635a2b-f22c-4dc0-8a5f-0fb6f0bdd2bb',
    'refresh_token': 'M.R3_BL2.-CRO8ZmgerJduIy6n*e9I52jaNhHGJxFP1np57lQH019asOZOZ2kcd0a8f1Vh*a7Bc3Nnli*8T44ER5PgVOGbEE3eqj9*u47BlHR1RI3HjAOKuIqpTSa7nyjydmXjoESvSzL3s4Nk2XMcXarV6WfXCvUTbpFR3VwDBHIMYt6KK**0qTKV5uGlbATxWRQdnrKOwji9ZD0A*H6cna9VFgBbrVwhXcV3HmtphM*eAr70y7KIZZEVgEl1BCQGtK3CEVyI2uw94Zb9QcCJej3OmXVmP6W2s9NrKrg9rFSl7cA!KwBtt9XiVy*0bUeTpgR!fi0w5g$$'
}

response = requests.post('https://login.microsoftonline.com/common/oauth2/v2.0/token', data=params)

access_token = response.json()['access_token']
new_refresh_token = response.json()['refresh_token']
print(access_token)
print(new_refresh_token)


header = {'Authorization': 'Bearer ' + access_token}

# Download the file
response = requests.get('https://graph.microsoft.com/v1.0/me/drive/root:' +
                        '/Test' + '/' + 'pgv.pptx' + ':/content', headers=header)

with open("1.pptx", 'wb') as file:
    file.write(response.content)
