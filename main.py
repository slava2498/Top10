# -*- coding: utf-8 -*-
from datetime import datetime
import sys
import os
import time
import json
import pathlib
import xlsxwriter

from bs4 import BeautifulSoup as BS

from selenium import webdriver
from selenium.webdriver.firefox.firefox_binary import FirefoxBinary
from selenium.webdriver.firefox.options import Options
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import StaleElementReferenceException
from selenium.common.exceptions import WebDriverException
from selenium.common.exceptions import SessionNotCreatedException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities

from dostoevsky.tokenization import RegexTokenizer
from dostoevsky.models import FastTextSocialNetworkModel

tokenizer = RegexTokenizer()
model = FastTextSocialNetworkModel(tokenizer=tokenizer)

search_array = ['яблоко', 'абрикос', 'киви']
systems = ['google', 'yandex']
screen_directory = 'screens'

def connect():
	global driver
	binary = FirefoxBinary(r"C:\Program Files\Mozilla Firefox\firefox.exe")
	options = Options()
	# options.add_argument("--headless")
	driver = webdriver.Firefox(options=options, firefox_binary=binary,
								executable_path=r'geckodriver.exe')
	print('start')

def collector(result, search, **kwargs):
	messages = '{} {}'.format(kwargs['h3'], kwargs['desc']).split()
	tonality = model.predict(messages, k=2)

	positive_array = []
	negative_array = []
	neutral_array = []
	for x in tonality:
		if 'positive' in x:
			positive_array.append(x['positive'])
		elif 'negative' in x:
			negative_array.append(x['negative'])
		elif 'neutral' in x:
			neutral_array.append(x['neutral'])

	if sum(positive_array) + sum(neutral_array) > sum(negative_array):
		tonality_result = 'positive'
	else:
		tonality_result = 'negative'

	result[search].append({
		'search_system': kwargs['search_system'],
		'number': kwargs['i'],
		'subject': kwargs['h3'],
		'href': kwargs['href'],
		'desc': kwargs['desc'],
		'tonality': tonality_result
	})

def parsing(url, search, result):
	global driver
	index = 1
	for page in range(2):
		driver.get(
			url.format(search) + '&start={}'.format(page) if 'google.ru' in url
			else url.format(search) + '&p={}'.format(page)
		)

		pathlib.Path(screen_directory).mkdir(parents=True, exist_ok=True)
		element = driver.find_element_by_tag_name('body')
		element.screenshot("{}/screenshot_{}.png".format(screen_directory, time.time()))
		
		if 'google.ru' in url:
			div = driver.find_element_by_xpath('//*[@id="rso"]')
			html = BS(div.get_attribute("innerHTML"), 'html.parser')

			for el in html.find_all("h3"):
				block_general = el.find_parent('div').find_parent('div')
				try:
					collector(
						result, 
						search,
						search_system='google',
						i=index,
						h3=el.text ,
						href=block_general.find('a')['href'],
						desc=block_general.find("div", class_="IsZvec").find("span").text
					)
					index += 1
					
				except Exception as error:
					continue
		else:
			div = driver.find_element_by_xpath('//*[@id="search-result"]')
			html = BS(div.get_attribute("innerHTML"), 'html.parser')

			for el in html.select(".serp-item"):
				try:
					a = el.select(".OrganicTitle-Link")[0]
					collector(
						result, 
						search,
						search_system='yandex',
						i=index,
						h3=a.select(".OrganicTitle-LinkText")[0].text,
						href=a['href'],
						desc=el.select(".extended-text__short")[0].text
					)

					index += 1
				except Exception as error:
					continue

	return result
				
			
def export_xlsx(result):
	workbook = xlsxwriter.Workbook('{}_export.xlsx'.format(time.time()))
	worksheet = workbook.add_worksheet()

	subject_cell = {
		'Поисковик': ['A', 'search_system'], 
		'Позиция': ['B', 'number'], 
		'Ссылка': ['C', 'href'], 
		'Заголовок': ['D', 'subject'], 
		'Описание': ['E', 'desc'],
		'Тон': ['F', 'tonality']
	}

	def checker_stop_search(count):
		export_result = {x: [] for x in search_array}
		for system in systems:
			for search in search_array:
				c = 1
				for x in result[search]:
					if x['search_system'] == system and c <= count:
						export_result[search].append(x)
						c += 1
		return export_result

	export_result = checker_stop_search(10)
	for key, value in subject_cell.items():
		worksheet.write('{}{}'.format(value[0], 1), key)

	row = 2
	for search in export_result.values():
		for ads in search:
			print(ads)
			print('-----------')
			for column in subject_cell.values():
				worksheet.write('{}{}'.format(column[0], row), ads[column[1]])
			row += 1

	workbook.close()

if __name__ == '__main__':
	global driver
	connect()
	
	result = {x: [] for x in search_array}
	for x in ['https://www.google.ru/search?q={}', 'https://yandex.ru/search/?text={}']:
		for y in search_array:
			result = parsing(url=x, search=y, result=result)

	export_xlsx(result)
