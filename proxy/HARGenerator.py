import time
import json
import pandas as pd
import sys

from browsermobproxy import Server
from selenium import webdriver
from datetime import datetime

class ProxyManager:
    __BMP = "C:/BMP/bin/browsermob-proxy.bat"

    def __init__(self):
        self.__server = Server(ProxyManager.__BMP)
        self.__client = None

    def start_server(self):
        self.__server.start()
        return self.__server

    def start_client(self):
        self.__client = self.__server.create_proxy(params={"trustAllServers": "true"})
        return self.__client

    @property
    def client(self):
        return self.__client

    def server(self):
        return self.__server


if "__main__" == __name__:
    df = pd.read_excel('pages.xlsx', sheet_name='Sheet1', header=[0])
    pages_list = df.values.tolist()
    for i in range(len(pages_list)):
        country = ""
        url = pages_list[i][0]
        now = datetime.now()
        proxy = ProxyManager()
        server = proxy.start_server()
        client = proxy.start_client()
        client.new_har(url)
        options = webdriver.ChromeOptions()
        options.add_argument("--proxy-server={}".format(client.proxy))
        driver = webdriver.Chrome('drivers/chromedriver.exe', options=options)

        driver.get("https://" + url)
        driver.maximize_window()
        time.sleep(5)
        try:
            country = str(sys.argv[1])+" "
        except:
            country = ""
        if url.find("/"):
            url = url.replace("/", " ")
        file_name = url + " " + country + now.strftime("%d%m%Y") + ".har"
        with open(file_name, 'w') as fp:
            json.dump(client.har, fp)
        server.stop()
        driver.close()
