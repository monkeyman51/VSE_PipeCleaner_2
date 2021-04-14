import json
import requests

personal_access_token: str = 'oa3kit3be5dlk2kfbkylhb62qn2kc4ja3363c5iogir66k5bwrwq'
teams_id: str = 'https://teams.microsoft.com/l/team/19%3ae51230b165724579a26d87baa1f6218f%40thread.skype/conversations?' \
                'groupId=f1cf49a1-033d-4d16-b879-d18c4ffd622f&tenantId=751040bd-e908-46a8-827c-ec32248b3e4d'
channel_id: str = 'https://teams.microsoft.com/l/channel/19%3ab57b6a851cbe4cc093bce2dd8b0c3191%40thread.skype/Labs?' \
                  'groupId=f1cf49a1-033d-4d16-b879-d18c4ffd622f&tenantId=751040bd-e908-46a8-827c-ec32248b3e4d'

example_id: str = "https://tasks.teams.microsoft.com/api/proxy/v1/planner/taskapi/v3.0/users('me')/details/" \
                  "all?%24deltatoken=35%257eb865dba3-5cce-484a-b49b-8e44cb608e75"

headers = {'Authorization': f'Bearer {example_id}'}

response = requests.get(channel_id, headers=headers)
print(response.text)
# print(json.dumps(json.loads(response.json()), sort_keys=True, indent=4))
# print(json.dumps(json.loads(response.json()), sort_keys=True, indent=4))