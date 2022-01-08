""" global variables are declared here """
import os
import sys
import json
import random

DIR_PATH = "C:/Users/play0/OneDrive/桌面/stock/" # modify the directory path to the one that your files are located in
def initialize_proxy():
    global proxy, proxy_list
    
    # read proxy file if it exist, or use the default proxies
    proxy_file_path = DIR_PATH + "proxies.txt"
    if os.path.isfile(proxy_file_path):
        proxy_list = list()
        with open(proxy_file_path, 'r', encoding="UTF-8") as file_r:
            for proxy in file_r:
                proxy_dict = json.loads(proxy) # JSON to dict
                proxy_list.append(proxy_dict)
    else:
        proxy_list = [{'http': 'http://34.145.247.64:3128'}, {'http': 'http://97.87.248.14:80'}, {'http': 'http://96.2.161.239:80'}, {'http': 'http://162.241.70.48:80'}, {'http': 'http://54.242.70.230:80'}, {'http': 'http://64.98.67.44:80'}, {'http': 'http://3.239.252.13:80'}, {'http': 'http://50.246.120.125:8080'}, {'http': 'http://74.205.128.200:80'}, {'http': 'http://52.43.61.135:80'}, {'http': 'http://18.219.203.105:80'}, {'http': 'http://35.164.200.0:80'}, {'http': 'http://13.64.91.93:3128'}, {'http': 'http://45.63.10.146:3128'}, {'http': 'http://52.170.137.199:3128'}, {'http': 'http://142.54.171.186:3128'}, {'http': 'http://40.132.82.159:8080'}, {'http': 'http://165.227.71.60:80'}, {'http': 'http://167.99.174.59:80'}, {'http': 'http://143.244.157.101:80'}, {'http': 'http://23.224.99.2:3128'}, {'http': 'http://13.58.114.163:3128'}, {'http': 'http://161.35.52.72:80'}, {'http': 'http://18.216.129.147:80'}, {'http': 'http://3.234.61.191:80'}, {'http': 'http://23.224.99.6:3128'}, {'http': 'http://40.121.142.45:3128'}, {'http': 'http://18.144.147.156:3128'}, {'http': 'http://52.6.12.221:3128'}, {'http': 'http://100.19.135.109:80'}, {'http': 'http://54.83.189.65:80'}, {'http': 'http://69.167.174.17:80'}, {'http': 'http://3.14.20.218:80'}, {'http': 'http://34.192.253.65:80'}, {'http': 'http://34.135.199.183:3128'}, {'http': 'http://13.57.240.98:3128'}, {'http': 'http://52.55.144.126:3128'}, {'http': 'http://54.186.218.27:3128'}, {'http': 'http://174.138.180.11:80'}, {'http': 'http://192.236.160.186:80'}, {'http': 'http://132.148.85.91:80'}, {'http': 'http://54.197.119.29:80'}, {'http': 'http://104.244.75.218:8080'}, {'http': 'http://34.145.159.5:8080'}, {'http': 'http://18.217.102.59:80'}, {'http': 'http://170.106.175.94:80'}, {'http': 'http://3.217.181.204:80'}, {'http': 'http://167.99.236.14:80'}, {'http': 'http://159.65.171.69:80'}, {'http': 'http://71.13.82.30:80'}, {'http': 'http://13.91.104.216:80'}, {'http': 'http://3.18.12.165:3128'}, {'http': 'http://40.91.94.165:3128'}, {'http': 'http://143.198.167.240:80'}, {'http': 'http://35.172.135.29:80'}, {'http': 'http://20.69.223.31:3128'}, {'http': 'http://92.204.129.161:80'}, {'http': 'http://52.149.152.236:80'}, {'http': 'http://12.69.91.226:80'}, {'http': 'http://74.208.128.22:80'}, {'http': 'http://51.81.83.125:8080'}, {'http': 'http://52.9.37.116:80'}]

    update_proxy()

# choose a proxy from proxy_list randomly, and remove it
def update_proxy():
    global proxy
    
    if proxy_list:
        proxy = random.choice(proxy_list)
        proxy_list.remove(proxy)
    else:
        sys.err.write("Please update proxies\n")
    
def main():
    pass

if __name__ == "__main__":
    main()