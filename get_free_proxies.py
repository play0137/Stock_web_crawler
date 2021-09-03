""" Get free proxies from proxy website, and write them to file """
from bs4 import BeautifulSoup
import requests
import json
import time

import global_vars

response = requests.get("https://www.us-proxy.org/")
soup = BeautifulSoup(response.text, 'lxml')
proxies = dict()
proxy_list = list()
table = soup.find_all("table", class_="table table-striped table-bordered")
trs = table.pop()
trs = trs.select("tbody tr")
print("the number of proxies:", len(trs))
for tr in trs:
    tds = tr.select("td")
    if len(tds) > 6:
        ip = tds[0].text
        port = tds[1].text
        anonymity = tds[4].text
        ifScheme = tds[6].text
        if ifScheme == 'yes': 
            scheme = 'https'
        else:
            scheme = 'http'
        proxy = "%s://%s:%s"%(scheme, ip, port)
        
        proxy = {scheme:proxy}
        try:
            # check whether the proxy is alive or not
            response = requests.session().get("http://ipv4.icanhazip.com/", proxies=proxy, timeout=5)
        except:
            pass
        # print(ip, response.text, sep=", ")
        if ip == response.text.strip():
            proxy_list.append(proxy)
            # print(proxy)

print("proxy_list:", proxy_list, sep='\n')
print("the number of valid proxies:", len(proxy_list))

with open(global_vars.DIR_PATH + "proxies.txt", 'w', encoding="UTF-8") as file_w:
    for proxy in proxy_list:
        json.dump(proxy, file_w) # writing JSON to a file
        file_w.write('\n')

