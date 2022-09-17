import time
from datetime import datetime

from dateutil.relativedelta import relativedelta
from pydantic import ValidationError
from selenium import webdriver
from selenium.common import ElementNotInteractableException, NoSuchElementException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait

from constants import CHROMEDRIVER_PATH, SITE_URL
from setup import get_logger
from utils.custom_exception import ChromDriverException
from utils.model import Report

logger = get_logger(__file__)


def _get_chrome_options() -> Options:
    """
    Get chrome options
    :return: Chrome options
    :rtype: Options
    """
    chrome_options = Options()
    for option in ["--disable-extensions", "--ignore-certificate-errors", "--ignore-ssl-errors"]:
        chrome_options.add_argument(option)
    chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])
    return chrome_options


def _get_driver() -> webdriver.Chrome:
    """
    Get chromedriver object
    :return: Object chromedriver
    :raises ChromDriverException: Incorrect path to chrome driver
    :rtype: webdriver.Chrome
    """
    if CHROMEDRIVER_PATH.exists():
        return webdriver.Chrome(executable_path=CHROMEDRIVER_PATH, options=_get_chrome_options())
    raise ChromDriverException()


driver = _get_driver()


def search_and_click_by_element(pattern_search: str, search_element: str,
                                element: WebElement = None, flag_click: bool = True) -> WebElement | None:
    """
    Search and click on the desired element in accordance with the condition and type of search
    :param pattern_search: Search template
    :type pattern_search: str
    :param search_element: Search pattern value
    :type search_element: str
    :param element: web element to search. If not specified, the search is performed on the entire web page
    :type element: WebElement
    :param flag_click: Flag click by element or not
    :type flag_click: bool
    :return: Search web element or None
    :rtype: WebElement|None
    """
    try:
        if element:
            find_element = element.find_element(pattern_search, search_element)
        else:
            find_element = driver.find_element(pattern_search, search_element)
        if flag_click:
            find_element.click()
            time.sleep(1)
    except NoSuchElementException:
        logger.error(
            f'No search element. Check pattern search or search element. (pattern: {pattern_search}, element: {search_element})')
        find_element = None
    except ElementNotInteractableException:
        logger.error(f'Element is not interactable (pattern: {pattern_search}, element: {search_element}).',
                     exc_info=True)
        find_element = None
    return find_element


def go_to_course_page():
    """
    Go to the course page
    :return:
    :rtype:
    """
    logger.info('Starting the transition to the report data page')
    driver.get(SITE_URL)
    time.sleep(1)

    search_and_click_by_element(
        By.XPATH,
        '//button[@class="header-content-bottom__menu-button hidden-lg js-adaptive-menu-button"]'
    )
    menu = search_and_click_by_element(By.CLASS_NAME, 'header-content-bottom-menu__vertical')
    search_and_click_by_element(By.XPATH, '//a[text()="Рынки"]', menu)
    search_and_click_by_element(By.XPATH, '//a[text()="Срочный рынок"]', menu)
    search_and_click_by_element(By.XPATH, '/html/body/div[2]/div/div/div/div/div[1]/div/a[1]')
    search_and_click_by_element(By.XPATH, '//div[@class="left-menu__header"]')
    scroll_window_down(300)
    search_and_click_by_element(By.XPATH, '//*[@id="ctl00_frmLeftMenuWrap"]/div/div/div/div[2]/div[13]/a', )
    logger.info('Transition to the page with report data completed')


def select_date_report():
    """
    Setting options for receiving a report
    :return:
    :rtype:
    """
    current_date = datetime.now()
    for i in range(2, 1, -1):
        select_option_js(f'd{i}day', str(current_date.date().day))
        select_option_js(f'd{i}month', str(current_date.date().month))
        select_option_js(f'd{i}year', str(current_date.date().year))
        current_date += relativedelta(months=-1)


def scroll_window_down(shift_down: int):
    """
    Scroll window down
    :param shift_down: Shift down value
    :type shift_down: int
    :return:
    :rtype:
    """
    driver.execute_script(f'window.scrollTo(0, {shift_down})')
    time.sleep(1)


def settings_report(currency: str) -> bool:
    """
    Settings for getting a report based on currency
    :param currency: Currency value
    :type currency: str
    :return: Flag about whether there is report data or not
    :rtype: bool
    """
    select_currencies(currency)
    scroll_window_down(200)
    time.sleep(1)
    select_date_report()
    if not search_and_click_by_element(By.XPATH, '//input[@type="submit" and @value="Показать"]'):
        return False
    time.sleep(1)
    scroll_window_down(200)
    search_and_click_by_element(By.XPATH, '//a[text()="Получить данные в XML"]')
    return True


def converter_tr_in_report(tr: WebElement) -> Report | None:
    """
    Transform to Report Object Converter
    :param tr: tr WebElement
    :type tr: WebElement
    :return: Report object or None
    :rtype: Report | None
    """
    tds = []
    try:
        tds = tr.find_elements(By.TAG_NAME, 'td')
        if len(tds) < 3:
            raise ValueError('Not enough data to create a Report object')
        date_report = datetime.strptime(tds[1].text, '%Y-%m-%d %H:%M:%S')
        report = Report(date=date_report.date(), exchange_rate=float(tds[2].text), time=date_report.time())
    except ValidationError:
        logger.error('Error converting data from tr to Report.', exc_info=True)
        report = None
    except ValueError:
        logger.warning(f'Not enough data to create a Report object ({", ".join(el.text for el in tds)}).')
        report = None
    except Exception:
        logger.error('Error converting in Report object', exc_info=True)
        report = None
    return report


def handle_trs_report(trs: list[WebElement]) -> list:
    """
    Handler for extracting data from tr tags
    :param trs: List of tr elements
    :type trs: List[WebElement]
    :return: Processed report data
    :rtype: list
    """
    data_report = []
    for el in trs:
        val = converter_tr_in_report(el)
        if val:
            data_report.append(val)
    return data_report


def get_data_from_report() -> list[Report]:
    """
    Extracting data from tr tags
    :return: Report data
    :rtype: list[Report]
    """
    try:
        trs = driver.find_elements(By.TAG_NAME, 'tr')
        return handle_trs_report(trs)
    except NoSuchElementException:
        logger.error('No search element. Check pattern search or search element.', exc_info=True)
    return []


def get_report() -> dict:
    """
    Get report data
    :return: Report data
    :rtype: dict
    """
    logger.info('Starting the report data acquisition process')
    data = {
        'USD_RUB': [],
        'JPY_RUB': []
    }
    for key in data.keys():
        if settings_report(key):
            data[key] = get_data_from_report()
            driver.back()
        time.sleep(1)
    logger.info('The process of obtaining report data is completed')
    return data


def handler_open_browser() -> dict:
    """
    Handler for opening a browser and extracting data from it
    :return: Report data
    :rtype: dict
    """
    try:
        go_to_course_page()
        report_data = get_report()
        time.sleep(1)
    except Exception:
        logger.error('Error in open browser function.', exc_info=True)
        report_data = {
            'USD_RUB': [],
            'JPY_RUB': []
        }
    finally:
        driver.close()
    return report_data


def move_up_list(select_element: WebElement, currency_value: str) -> int:
    """
    Move select currencies down and search for the index how much you need to move the cursor to select
    :param select_element: Select WebElement object
    :type select_element: WebElement
    :param currency_value: Currency value option
    :type currency_value: str
    :return: Cursor shift to select the desired option
    :rtype: int
    """
    index_option = 0
    try:
        options = select_element.find_elements(By.TAG_NAME, 'option')
        for i in range(len(options)):
            select_element.send_keys(Keys.ARROW_UP)
            if options[i].get_attribute('value') == currency_value:
                index_option = i
    except Exception:
        logger.error('Error in move to the list', exc_info=True)
    return index_option


def move_down_list(element: WebElement, cnt_press_down: int):
    """
    Move select currencies down
    :param element: Select WebElement object
    :type element: WebElement
    :param cnt_press_down: How many times to shift the currency selection
    :type cnt_press_down: int
    :return:
    :rtype:
    """
    for i in range(cnt_press_down):
        element.send_keys(Keys.ARROW_DOWN)


def select_currencies(currency_value: str):
    """
    Currency selection handler
    :param currency_value: Currency value
    :type currency_value: str
    :return:
    :rtype:
    """
    try:
        wait = WebDriverWait(driver, 10)
        element = wait.until(
            EC.visibility_of_element_located((By.XPATH, '//*[@id="ctl00_PageContent_CurrencySelect"]')))
        element.click()
        time.sleep(1)
        cnt_press_down = move_up_list(element, currency_value)
        move_down_list(element, cnt_press_down)
        element.send_keys(Keys.ENTER)
        time.sleep(1)
    except ValidationError:
        logger.warning('Error select currencies', exc_info=True)


def select_option_js(select_id: str, value: str):
    """
    Emulate value selection in select tag
    :param select_id: select tag id
    :type select_id: str
    :param value: Option value
    :type value: str
    :return:
    :rtype:
    """
    driver.execute_script(f'document.getElementById("{select_id}").value = {value};')
    time.sleep(1)


if __name__ == '__main__':
    handler_open_browser()
