# -*- coding: utf-8 -*-
from scrapy import Spider, Request
from collections import OrderedDict
from xlrd import open_workbook
import os, requests

def download(url, destfilename):
    if not os.path.exists(destfilename):

        try:
            r = requests.get(url, stream=True)
            with open(destfilename, 'wb') as f:
                for chunk in r.iter_content(chunk_size=1024):
                    if chunk:
                        f.write(chunk)
                        f.flush()
        except:
            # print(url)
            pass

def readExcel(path):
    wb = open_workbook(path)
    result = []
    for sheet in wb.sheets():
        number_of_rows = sheet.nrows
        number_of_columns = sheet.ncols
        herders = []
        for row in range(0, number_of_rows):
            values = OrderedDict()
            for col in range(number_of_columns):
                value = (sheet.cell(row,col).value)
                if row == 0:
                    herders.append(value)
                else:

                    values[herders[col]] = value
            if len(values.values()) > 0:
                result.append(values)
        break

    return result


class AngelSpider(Spider):
    name = "image_darty"
    start_urls = 'https://www.darty.com/'
    count = 0
    use_selenium = False
    models = readExcel("Input.xlsx")

    def start_requests(self):

        yield Request(self.start_urls, callback=self.parse)

    def parse(self, response):
        for i, val in enumerate(self.models):
            ern_code = val['EAN']
            url = ''
            if ern_code != '':
                try:
                    ern_code = str(int(val['EAN']))
                except:
                    ern_code = str(val['EAN'])
                url ='https://www.darty.com/nav/recherche/{}.html'.format(ern_code)
            yield Request(url, callback=self.parse1, meta={'index':i})

    def parse1(self, response):
        ordernum = response.meta['index']
        sku = self.models[ordernum]['EAN']
        try:
            sku = str(int(sku))
        except:
            sku = str(sku)
        item=OrderedDict()
        item['EAN'] = sku
        for i in range(5):
            item['Image' + str(i+1)] = ''

        index = 0

        image_urls = response.xpath('//div[contains(@class, "v6vertical_new_product_page_sizes")]/img/@data-src').extract()
        if len(image_urls) < 5:
            image_urls1 = response.xpath('//div[contains(@class, "v6horizontal_new_product_page_sizes")]/img/@data-src').extract()
            if image_urls1:
                image_urls.extend(image_urls1)

        for image_url in image_urls:
            index += 1
            if index > 5:
                break
            # baseName = image_url.strip().split('/')[-1].replace('.jpg', '')
            image_name = str(sku) + '_' + str(index) +".jpg"
            filename = "Images/" + image_name
            download(image_url.strip(), filename)
            item['Image' + str(index)] = image_name

        self.models[ordernum] = item
        yield item
    def err_parse(self, response):
        pass
        # i = response.request.meta['index']
        # item=OrderedDict()
        # item['Sku'] = self.models[i]
        # item['Image'] = ''
        # self.models[i] = item
        # yield item