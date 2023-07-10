import requests
from bs4 import BeautifulSoup

r = requests.get('https://book.impress.co.jp/')
soup = BeautifulSoup(r.text, 'html.parser')
print(soup.find('h2'))
print(soup.find('h2').text)
