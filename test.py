
import requests
from openpyxl import Workbook,load_workbook


baseurl = "https://www.mcxindia.com/"
url = f"https://www.mcxindia.com/market-data/market-watch"
headers = {'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, '
                         'like Gecko) '
                         'Chrome/80.0.3987.149 Safari/537.36',
           'accept-language': 'en,gu;q=0.9,hi;q=0.8', 'accept-encoding': 'gzip, deflate, br'}
session = requests.Session()
request = session.get(baseurl, headers=headers, timeout=5)
cookies = dict(request.cookies)
response = session.get(url, headers=headers, timeout=5, cookies=cookies)
x=str(response.content)
posS = x.find("vTick")
posE=x.find("]",posS,len(x))
z=x[posS+8:posE-1].split("},{")
print(len(z))

for elm in z:
    eelm=elm.split(",")
    print(elm)