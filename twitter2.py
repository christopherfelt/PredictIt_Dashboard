import tweepy
import json
import pandas
import credentials

consumer_key = credentials.consumer_key
consumer_secret = credentials.consumer_secret
access_token = credentials.access_token
access_token_secret = credentials.access_token_secret

auth = tweepy.OAuthHandler(consumer_key, consumer_secret)
auth.set_access_token(access_token, access_token_secret)

api = tweepy.API(auth)

count_dict = {}

for i in range(1,32):

    item = api.user_timeline(screen_name='@realdonaldtrump', count=100, page=i)
    for j in range(0,len(item)):
        created_at_raw = json.dumps(item[j]._json['created_at'])
        created_at_split = created_at_raw.split(" ")
        created_at_date = created_at_split[1]+"/"+created_at_split[2]

        if created_at_date in count_dict:
            count_dict[created_at_date] += 1
        else:
            count_dict[created_at_date] = 1


for key,value in count_dict.items():
    print(key,value)











