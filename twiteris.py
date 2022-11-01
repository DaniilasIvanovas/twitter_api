import openpyxl
import tweepy
import configparser

# Get keys
configas = configparser.ConfigParser()
configas.read('info.ini')

api_key = configas['keys']['api_key']
api_key_secret = configas['keys']['api_key_secret']

access_token = configas['keys']['access_token']
access_token_secret = configas['keys']['access_token_secret']

# Authentification

aut = tweepy.OAuthHandler(api_key, api_key_secret)
aut.set_access_token(access_token, access_token_secret)

# Query

query = 'is:retweet'

api = tweepy.API(aut)
tweets = api.search_tweets(q=query, count=10)


wb = openpyxl.load_workbook('Duomen.xlsx')  # Load workbook
ws = wb.active  # Select first sheet
row = 2
col = 1

for tweet in tweets:
    cell = ws.cell(row=row, column=col)
    cell.value = tweet.user.id_str
    row += 1

wb.save('Duomen.xlsx')

