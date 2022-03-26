from msedge.selenium_tools import Edge, EdgeOptions
from selenium.webdriver.common.by import By
from selenium import webdriver
import pandas as pd
import time
import pdb
import global_vars

class Test():
    
  def setup_method(self):
    self.driver = Edge(global_vars.DIR_PATH + "edgedriver_win64/msedgedriver.exe")
    # self.driver = webdriver.Chrome()
    self.vars = {}
  
  def teardown_method(self):
    self.driver.quit()
   
  def wait_for_window(self, timeout = 2):
    time.sleep(round(timeout / 1000))
    wh_now = self.driver.window_handles
    wh_then = self.vars["window_handles"]
    if len(wh_now) > len(wh_then):
      return set(wh_now).difference(set(wh_then)).pop()
   
  def test_(self):
    self.driver.get("https://norway.twsthr.info/StockHolders.aspx?stock=2330")
    time.sleep(2)
    self.driver.set_window_size(1095, 817)
    self.vars["window_handles"] = self.driver.window_handles
    self.driver.find_element(By.LINK_TEXT, "損益表").click()
    self.vars["win4417"] = self.wait_for_window(2000)
    self.driver.switch_to.window(self.vars["win4417"]) # 切換視窗
    html = self.driver.page_source
    table1 = pd.read_html(html)[9] # 最後的index不能固定用0，要看資料在哪裡
    pdb.set_trace()
    table1_string = table1[0].to_list()[0] # 我把資料轉成string型態，接著就是處理這個string然後寫出到檔案
    print(table1_string)

obj = Test() # 初始化物件obj
obj.setup_method()
time.sleep(2)
obj.test_()    
obj.teardown_method()    
    