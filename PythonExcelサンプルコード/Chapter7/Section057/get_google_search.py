import json

from googleapiclient.discovery import build

api_key = 'APIキー'
search_engine_id = '検索エンジンID'
keyword = 'Python'

service = build('customsearch', 'v1', developerKey=api_key)

response = service.cse().list(
    q=keyword,
    cx=search_engine_id,
    lr='lang_ja',
    num=10,
    start=1
).execute()

with open('search.json', 'w', encoding='utf-8') as f:
    json.dump(response, f, indent=2, ensure_ascii=False)
