from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
import openpyxl
from datetime import datetime

QUERIES = ["яблоко", "абрикос", "малина"]
TOP_NUM = 10
WB = openpyxl.load_workbook("report.xlsx")
DT = datetime.now().strftime('%m-%d-%Y')

CODE_RED = ["избегайте", "опасен", "опасны", "издевательство"]
CODE_GREEN = ["отличный", "легко", "полезен", "полезные", "полезный",
              "профессиональные", "профессиональное", "профессиональная",
              "натуральная", "уважение", "современной", "удобной", "лечебный", "инновации",
              "развлечения", "мощное", "современной", "благодарны", "интересные", "круто"]


def main():
    chrome_options = Options()
    chrome_options.add_argument("--headless")
    chrome_options.add_argument(f"--window-size=1920,2800")
    chrome_options.add_argument("--hide-scrollbars")
    driver = webdriver.Chrome(options=chrome_options)

    driver.get("https://www.google.com")
    for i in range(len(QUERIES)):
        collect_data_google(QUERIES[i], driver)

    driver.get("https://yandex.ru")
    for j in range(len(QUERIES)):
        collect_data_yandex(QUERIES[j], driver)

    driver.quit()
    WB.save('report.xlsx')


def collect_data_google(query, driver):
    sheet = WB[query]
    search = driver.find_element_by_name("q")
    search.clear()
    search.send_keys(query + "\n")
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.TAG_NAME, "tbody"))  # wait until the lowest table is downloaded
    )
    results_collected = 0
    page_num = 1
    while results_collected < TOP_NUM:
        driver.save_screenshot("screenshots/" + query + "-google-" + str(page_num) + "_" + DT + ".png")
        page_results = driver.find_elements_by_xpath("//div[@class='g']/div[not(ancestor::div[@class='xpdopen'])]")
        for result in page_results:
            results_collected += 1
            sheet['B' + str(1 + results_collected)] = results_collected
            header = result.find_element_by_tag_name("h3").text
            sheet['C' + str(1 + results_collected)] = header
            summary = result.find_element_by_class_name("st").text
            sheet['D' + str(1 + results_collected)] = summary
            link = result.find_element_by_tag_name("a")
            sheet['E' + str(1 + results_collected)] = link.get_attribute("href")
            rate = analyze_tone(header + " " + summary)
            sheet['F' + str(1 + results_collected)] = rate
            if results_collected == TOP_NUM:
                break
        if results_collected < TOP_NUM:
            page_num += 1
            new_page = driver.find_element_by_link_text(str(page_num))
            new_page.click()
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.TAG_NAME, "tbody"))
            )
    WB.save('report.xlsx')


def collect_data_yandex(query, driver):
    sheet = WB[query]
    search = driver.find_element_by_name("text")
    search.clear()
    search.send_keys(query + "\n")
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CLASS_NAME, "pager__items"))  # wait until the lowest table is downloaded
    )
    results_collected = 0
    page_num = 1
    while results_collected < TOP_NUM:
        driver.save_screenshot("screenshots/" + query + "-yandex-" + str(page_num) + "_" + DT + ".png")
        page_results = driver.find_elements_by_xpath("//li[contains(@class, 'serp-item')"
                                                     "and not(contains(@data-fast-wzrd, 'images'))"
                                                     "and not(contains(@data-fast-wzrd, 'collections'))"
                                                     "and not(contains(@data-fast-wzrd, 'market_constr'))"
                                                     "and not(contains(@data-fast-wzrd, 'videowiz'))"
                                                     "and not(contains(@data-fast-wzrd, 'mushroom'))]")
        for result in page_results:
            results_collected += 1
            sheet['B' + str(1 + TOP_NUM + results_collected)] = results_collected
            header = result.find_element_by_class_name("organic__url-text").text
            sheet['C' + str(1 + TOP_NUM + results_collected)] = header
            summary = result.find_element_by_tag_name("div.text-container.typo.typo_text_m.typo_line_m.organic__text").text
            sheet['D' + str(1 + TOP_NUM + results_collected)] = summary
            link = result.find_element_by_tag_name("a")
            sheet['E' + str(1 + TOP_NUM + results_collected)] = link.get_attribute("href")
            rate = analyze_tone(header + " " + summary)
            sheet['F' + str(1 + TOP_NUM + results_collected)] = rate
            if results_collected == TOP_NUM:
                break
        if results_collected < TOP_NUM:
            page_num += 1
            new_page = driver.find_element_by_link_text(str(page_num))
            new_page.click()
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CLASS_NAME, "pager__items"))
            )
    WB.save('report.xlsx')


def analyze_tone(text):
    for item in CODE_RED:
        if item in text:
            return -1
    for item in CODE_GREEN:
        if item in text:
            return 1
    return 0


if __name__ == "__main__":
    main()
