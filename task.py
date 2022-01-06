import os

from RPA.Browser.Selenium import Selenium
from RPA.Excel.Files import Files
from RPA.FileSystem import FileSystem
from RPA.Tables import Tables

table = Tables()
browser = Selenium()
file = Files()
lit = []
til = []
links = []
list = []
liss = []

# browser.set_download_directory
browser.open_available_browser('https://itdashboard.gov/')
browser.click_link(locator='//*[@id="node-23"]/div/div/div/div/div/div/div/a')
file.create_workbook("output/workbook.xlsx")
lis = browser.find_elements(locator='//span[@class="h4 w200"]')
for x in lis:
    lit.append(x.text)
while '' in lit:
    lit.remove('')
test = browser.find_elements(locator='//span[@class=" h1 w900"]')
for z in test:
    til.append(z.text)
while '' in til:
    til.remove('')
data = {"Agencies": lit, "Amount": til}
file.append_rows_to_worksheet(content=data, header=True)
file.rename_worksheet(src_name="Sheet", dst_name="Agencies")
file.save_workbook()
a = browser.get_text('//*[@id="investments-table-object_wrapper"]')
var = a.split(" ")[-1]
for i in range(1, int(var)):
    UII = browser.find_element(f'//*[@id="investments-table-object"]/tbody/tr[{i}]/td[1]').text
    Bureau = browser.find_element(f'//*[@id="investments-table-object"]/tbody/tr[{i}]/td[2]').text
    Investment = browser.find_element(f'//*[@id="investments-table-object"]/tbody/tr[{i}]/td[3]').text
    Total_Spending = browser.find_element(f'//*[@id="investments-table-object"]/tbody/tr[{i}]/td[4]').text
    Type = browser.find_element(f'//*[@id="investments-table-object"]/tbody/tr[{i}]/td[5]').text
    CIO_Rating = browser.find_element(f'//*[@id="investments-table-object"]/tbody/tr[{i}]/td[6]').text
    Projects = browser.find_element(f'//*[@id="investments-table-object"]/tbody/tr[{i}]/td[7]').text
    if UII and Bureau and Investment and Total_Spending and Type and CIO_Rating and Projects:
        list.append({"UII": UII, "Bureau": Bureau, "Investment": Investment, "Total_Spending": Total_Spending,
                     "Type": Type, "CIO_Rating": CIO_Rating, "Projects": Projects})
    try:
        link = browser.find_element(
            f'//*[@id="investments-table-object"]/tbody/tr[{i}]/td[1]').find_element_by_tag_name(
            "a").get_attribute("href")
        links.append({"link": link, "UII": UII, "Investment": Investment})
    except:
        pass

print(links)

print("Done")
