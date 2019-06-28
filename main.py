from selenium import webdriver
from bs4 import BeautifulSoup
import xlsxwriter

URL = "http://www.extra.com.br"
MAX_PAGE = 20


def main():
    # Set chrome to not load images
    chromeOptions = webdriver.ChromeOptions()
    prefs = {'profile.managed_default_content_settings.images': 2}
    chromeOptions.add_experimental_option("prefs", prefs)
    driver = webdriver.Chrome(options=chromeOptions)

    # Get url
    driver.get(URL)
    driver.implicitly_wait(100)

    # find search bar and input search terms
    search_bar = driver.find_element_by_id("ctl00_TopBar_PaginaSistemaArea1_ctl04_ctl00_txtBusca")
    search_bar.send_keys("pet shop")
    search_button = driver.find_element_by_id("ctl00_TopBar_PaginaSistemaArea1_ctl04_ctl00_btnOK")
    search_button.click()

    # get product links from 20 first pages
    links = []
    page = 1
    current_url = driver.current_url

    while page <= MAX_PAGE:
        driver.get((current_url + '&page=%s' % str(page)))
        driver.implicitly_wait(100)

        li_items = driver.find_elements_by_class_name("nm-product-item")
        for li in li_items:
            a = li.find_elements_by_class_name("nm-product-img-link")[0]
            link = a.get_attribute("href")
            links.append(link)

        page += 1

    # Open excel workbook
    workbook = xlsxwriter.Workbook('data/extra_pet_shop_products.xlsx')
    worksheet = workbook.add_worksheet()

    worksheet.write_string(0, 0, "name")
    worksheet.write_string(0, 1, "code")
    worksheet.write_string(0, 2, "price")
    worksheet.write_string(0, 3, "description")

    # Loop through links and get data
    row = 1
    for link in links:
        driver.get(link)
        driver.implicitly_wait(100)

        soup = BeautifulSoup(driver.page_source, 'html.parser')

        product_name = soup.find("h1", class_="fn name")
        product_code = soup.find(itemprop="productID")
        product_price = soup.find("i", class_="sale price")
        product_description = soup.find("div", id="descricao")

        if product_name is not None:
            worksheet.write_string(row, 0, product_name.get_text())
        else:
            worksheet.write_blank(row, 0, None)

        if product_code is not None:
            worksheet.write_string(row, 1, product_code.get_text().split()[-1][:-1])
        else:
            worksheet.write_blank(row, 1, None)

        if product_price is not None:
            worksheet.write_string(row, 2, product_price.get_text())
        else:
            worksheet.write_blank(row, 2, None)

        if product_description is not None:
            worksheet.write_string(row, 3, product_description.get_text())
        else:
            worksheet.write_blank(row, 3, None)

        row += 1

    workbook.close()
    driver.quit()


if __name__ == '__main__':
    main()
