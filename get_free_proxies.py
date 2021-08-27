""" Get proxies from free proxy website """
from bs4 import BeautifulSoup
import requests

response = requests.get("https://www.us-proxy.org/")
soup = BeautifulSoup(response.text, 'lxml')
proxies = dict()
proxy_list = list()
trs = soup.select("#proxylisttable tr")
for tr in trs:
    tds = tr.select("td")
    if len(tds) > 6:
        ip = tds[0].text
        port = tds[1].text
        anonymity = tds[4].text
        ifScheme = tds[6].text
        if ifScheme == 'yes': 
            scheme = 'https'
        else: scheme = 'http'
        proxy = "%s://%s:%s"%(scheme, ip, port)
        
        proxy = {scheme:proxy}
        try:
            # check whther the proxy is alive or not
            response = requests.get("http://ipv4.icanhazip.com/", proxies = proxy, timeout=10)
        except:
            pass
        print(ip, response.text, sep=", ")
        if ip == response.text.strip():
            proxy_list.append(proxy)
            # print(proxy)

print(proxy_list)
print("size:", len(proxy_list))
..
