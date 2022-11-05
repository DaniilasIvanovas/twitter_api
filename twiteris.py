import openpyxl
import tweepy
import configparser
import json
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

query = 'word -is:reply -is:retweet'
api = tweepy.API(aut)
tweets = api.search_tweets(q=query, count=100)


wb = openpyxl.load_workbook('Duomen.xlsx')  # Load workbook
ws = wb.active  # Select first sheet
row = 2
col = 'ABCDEFGHIJKLMNOPRSTUVZ'
keys = ['id_str', 'text']


def get_info_from_tweet(tweet_data):
    info = [tweet_data['id_str'], tweet_data['lang'], len(tweet_data['text']), len(tweet_data['entities']['hashtags']),
            len(tweet_data['entities']['symbols']), len(tweet_data['entities']['user_mentions']), len(tweet_data['entities']['urls'])]
    return info


for tweet in tweets:
    ws.append(get_info_from_tweet(tweet._json))


wb.save('Duomen.xlsx')
