import http.client
from numpy import fix
import pandas as pd
import json

curr_date = "2021-05-18"    #--------------------Change Date Here

conn = http.client.HTTPSConnection("api-football-v1.p.rapidapi.com")

headers = {
    'x-rapidapi-key': "************************************",
    'x-rapidapi-host': "api-football-v1.p.rapidapi.com"
    }

conn.request("GET", "/v3/fixtures?date="+curr_date, headers=headers)

res = conn.getresponse()
fixture = res.read()


fixture = str(fixture, 'UTF-8')

js_test = json.loads(fixture)

df = pd.DataFrame()
dict = {}

for fix in js_test['response']:
    fixture_id = fix['fixture']['id']
    date = fix['fixture']['date']
    city = fix['fixture']['venue']['city']
    league = fix['league']['id']
    country = fix['league']['country']
    home_team_id = fix['teams']['home']['id']
    home_team_name = fix['teams']['home']['name']
    away_team_id = fix['teams']['away']['id']
    away_team_name = fix['teams']['away']['name']

    dict = {"fixture_id":fixture_id,"date":date,"city":city,"league":league,"country":country,"home_team_id":home_team_id,"home_team_name":home_team_name,"away_team_id":away_team_id,"away_team_name":away_team_name}

    df = df.append(dict,ignore_index=True)
   
cols = df.columns.tolist()  
df = df[['fixture_id', 'league', 'country', 'city', 'date', 'home_team_id', 'home_team_name', 'away_team_id', 'away_team_name']]  

print (df)
df.to_excel('Fixtures.xlsx', sheet_name='Sheet1')












