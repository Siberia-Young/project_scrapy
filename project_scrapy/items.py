# Define here the models for your scraped items
#
# See documentation in:
# https://docs.scrapy.org/en/latest/topics/items.html

import scrapy


class Item(scrapy.Item):
    category = scrapy.Field()
    filepath = scrapy.Field()
    keyword = scrapy.Field()
    page = scrapy.Field()
    goods = scrapy.Field(list_type=list)

    end = scrapy.Field()
    keywords = scrapy.Field(list_type=list)
    relatedwords = scrapy.Field(list_type=list)
    graphicwords = scrapy.Field(list_type=list)

    
