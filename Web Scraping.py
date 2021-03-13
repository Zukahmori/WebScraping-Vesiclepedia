import pandas as pd
from selenium import webdriver
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.common.by import By
import os

folderPath = os.path.dirname(__file__)
option = Options()
option.headless = True
firefoxPath = f'{folderPath}/geckodriver.exe'
driver = webdriver.Firefox(executable_path=firefoxPath)

table = pd.read_excel(f'{folderPath}/Table of genes.ods')
geneName = table['GENE']

tableDictionary = geneName.to_dict()
results = []
for geneArray in tableDictionary.items():
    gene = geneArray[1]
    url = f''
    driver.get(url)

    notFoundErrorArray = driver.find_elements(By.TAG_NAME, 'font')
    if len(notFoundErrorArray) > 2:
        notFoundError = notFoundErrorArray[2].text == 'No results found'
        if notFoundError:
            results.append('Gene Not Found')
    else:
        results.append('Gene Found')

table['Found in Vesiclepedia'] = results

table.to_excel('results.ods')


