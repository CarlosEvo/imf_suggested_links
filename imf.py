from time import sleep

from openpyxl import Workbook

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import TimeoutException


def get_query():
    query_list = ['CPI', 'Consumer price index']
    return query_list


def get_html(driver, query):

    url = 'https://www.imf.org/en/search#q=' + query + '&sort=relevancy'
    driver.get(url)

    try:
        sleep(1)
        element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, 'suggestedLinks'))
        )
        results = element.text
        if not results:
            results = 'No suggested link'
        else:
            results = results[16:]
    except TimeoutException:
        return 'Timeout Error'

    # print(results)
    return results


def write_data(query_list, results_list):
    wb = Workbook()
    ws = wb.create_sheet('Search Results')
    ws.append(['query', 'suggested links'])
    for i in range(len(query_list)):
        ws.append([query_list[i], results_list[i]])

    wb.save(filename='results.xlsx')


def main():

    options = Options()
    options.headless = True
    driver = webdriver.Firefox(options=options)

    query_list = get_query()
    results_list = []

    for query in query_list:
        results_list.append(get_html(driver, query))

    write_data(query_list, results_list)

    driver.quit()


if __name__ == '__main__':
    main()
