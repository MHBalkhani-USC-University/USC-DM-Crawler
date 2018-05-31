import scrapy

from openpyxl import Workbook

import re
from urllib.parse import urlparse



#preloads
shop_words={
    'shop':0,
    'off':0,
    'sell':0
}

shop_category_words = [
    'Fashion',
    'Motor',
    'Art',
    'Home',
    'Garden',
    'Baby',
    'Furniture',
    'Sporting',
    'Sport',
    'Toy',
    'Business',
    'Industrial',
    'Music',
    'Deal',
    'Pet',
    'Game',
    'Video',
    'Book',
    'Computer',
    'Software',
    'Hardware',
    'Movie',
    'TV',
    'Vehicle',
    'Grocery',
    'Food',
]

news_words = {
    'news':0,
    'story':0,
    'world':0
}

news_category_words = [
    'Politic',
    'Money',
    'Stock',
    'Entertainment',
    'Tech',
    'Sport',
    'Travel',
    'Style',
    'Health',
    'Finance',
    'Lifestyle',
    'Auto',
    'Art',
    'Music'
]

# Creating workbook
wb = Workbook()
ws = wb.active

columns = ['Url']

for shop_word in shop_words:
    columns.append(shop_word)

for news_word in news_words:
    columns.append(news_word)

columns.extend(['[Number]%','$[Number]','Shop Categories','News Categories','Twitter Refrences','Facebook Refrences','Instagram Refrences'])

ws.append(columns)

shop_category_words = [shop_category_word.lower() for shop_category_word in shop_category_words]
news_category_words = [news_category_word.lower() for news_category_word in news_category_words]

class ShopSpider(scrapy.Spider):
    name = 'spider'
    start_urls = [
        'https://www.cnn.com',
        'https://www.ebay.com'
    ]

    page_count = 0
    page_limit = 1200

    def parse(self, response):

        #statistics
        shop_category_words_c = 0
        news_category_words_c = 0

        currency_number_c = 0
        percentile_number_c = 0
        
        twitter_ref = 0
        facebook_ref = 0
        instagram_ref = 0

        for shop_word in shop_words:
            shop_words[shop_word] = 0

        for news_word in news_words:
            news_words[news_word] = 0

        #extracting words
        for word in response.css('span::text, strong::text, div::text, a::text').extract():

            word = word.lower()

            for shop_word in shop_words:
                if shop_word in word:
                    shop_words[shop_word]+=1

            for news_word in news_words:
                if news_word in word:
                    news_words[news_word]+=1

            for shop_category_word in shop_category_words:
                if shop_category_word in word:
                    shop_category_words_c+=1

            for news_category_word in news_category_words:
                if news_category_word in word:
                    news_category_words_c+=1
            
            if re.search(r'[$]\d+|\d+[$]',word,re.M|re.I):
                currency_number_c+=1

            if re.search(r'[%]\d+|\d+[%]',word,re.M|re.I):
                percentile_number_c+=1


        #checking refrences
        for link in response.css('a::attr(href)').extract():

            if 'facebook.com' in link:
                facebook_ref+=1
            elif 'instagram.com' in link:
                instagram_ref+=1
            elif 'twitter.com' in link:
                twitter_ref+=1

        
        #exporting
        ws_columns = [response.url]


        for shop_word in shop_words:
            ws_columns.append(shop_words[shop_word])

        for news_word in news_words:
            ws_columns.append(news_words[news_word])

        ws_columns.append(percentile_number_c)
        ws_columns.append(currency_number_c)
            
        ws_columns.append(shop_category_words_c)
        ws_columns.append(news_category_words_c)

        ws_columns.append(twitter_ref)
        ws_columns.append(facebook_ref)
        ws_columns.append(instagram_ref)

        self.log(shop_words)
        self.log(news_words)
        self.log(news_category_words)

        ws.append(ws_columns)

        for next_page in response.css('li a::attr(href)'):
            if self.page_count != self.page_limit:
                self.page_count+=1
                self.log('Page Count : %i'%self.page_count)
                yield response.follow(next_page, self.parse)
            else:
                # Save the file
                wb.save("export.xlsx")

            

        