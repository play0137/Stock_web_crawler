""" global variables are declared here """
import random
import json

DIR_PATH = "C:/Users/play0/OneDrive/桌面/stock/" # modify the directory path to the one that your files are located in
def initialize_proxy():
    global proxy, proxy_list
    proxy_list = list()
    with open(DIR_PATH + "proxies.txt", 'r', encoding="UTF-8") as file_r:
        for proxy in file_r:
            proxy_dict = json.loads(proxy) # JSON to dict
            proxy_list.append(proxy_dict)
    update_proxy()

# choose a proxy from proxy_list randomly, and remove it
def update_proxy():
    global proxy
    proxy = random.choice(proxy_list)
    proxy_list.remove(proxy)
    
def main():
    pass

if __name__ == "__main__":
    main()