import os
from scrapy.spider import Spider
from scrapy.http import Request
import openpyxl
from scrapy import signals
from scrapy.xlib.pydispatch import dispatcher


class FoodCrawler(Spider):

    '''
    Crawls bb-team for food values
    '''

    name = 'gipsy'
    start_urls = ['https://www.bb-team.org/hrani']

    def __init__(self, *args, **kwargs):
        super(FoodCrawler, self).__init__(*args, **kwargs)
        dispatcher.connect(self.spider_closed, signals.spider_closed)
        self.workbook = openpyxl.Workbook()
        self.worksheet = self.workbook.active
        self.worksheet.append([
            "Category",
            "Name",
            "Protein",
            "Carbs",
            "Fats",
            "Calories"
        ])
        self.wb_name = "Food_macronutrients.xlsx"
        self.workbook.save(self.wb_name)

    def parse(self, response):
        category_urls = response.xpath('//div[contains(@class, "row")][1]/div/a[1]/@href').extract()
        category_names = response.xpath('//div[contains(@class, "row")][1]/div/a/h2/text()').extract()
        for category_url, category_name in zip(category_urls, category_names):
            print('Yielding request for category: {}'.format(category_name.encode('utf-8').strip()))
            yield Request(category_url,
                          callback=self.parse_category,
                          meta={
                              'category': category_name.encode('utf-8').strip()
                          })


    def parse_category(self, response):
        category = response.meta.get('category', None)
        print(category)
        products_urls = response.xpath('//div[contains(@class, "row")][1]/div/a[1]/@href').extract()
        product_titles = response.xpath('//div[contains(@class,"row")]/div/a/div/text()').extract()[0:][::2]
        product_descriptions = response.xpath('//div[contains(@class,"row")]/div/a/div/text()').extract()[1:][::2]


        for url,title,description in zip(products_urls,
                                     product_titles,
                                     product_descriptions):
            yield Request(url,
                          callback=self.parse_product,
                          meta={
                              'title':title,
                              'description': description,
                              'category': category
                          })

    def parse_product(self, response):
        product_category = response.meta.get('category' or None)
        product_title = response.meta.get('title', None)
        product_description = response.meta.get('description', None)
        macronutrients = [
                    "calories",
                    "proteinContent",
                    "carbohydrateContent",
                    "fatContent"
        ]
        macronutrients_xpath = '//span[contains(@itemprop, "{}")]/text()'.format

        data = {}

        for nutrient in macronutrients:
            try:
                data[nutrient] = float(response.xpath(macronutrients_xpath(nutrient)).extract()[0])
            except IndexError:
                data[nutrient] = 0

        self.worksheet.append([
            product_category,
            ','.join([product_title, product_description]),
            data['proteinContent'],
            data['carbohydrateContent'],
            data['fatContent'],
            data['calories'],
        ])

    def spider_closed(self, spider):
        self.workbook.save(self.wb_name)
