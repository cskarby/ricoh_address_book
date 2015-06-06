#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Using Selenium WebDriver to add/update (write) and delete(remove) users
from address book on Ricoh Aficio MP C4502 via Web Image Monitor.

http://seleniumhq.org/docs/03_webdriver.jsp
http://www.ricoh-usa.com/products/product_details.aspx?cid=8&scid=5&pid=2308


Copyright (c) 2013 NKI AS, Postboks 6674 - St. Olavs plass, NO-0129 OSLO

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to
deal in the Software without restriction, including without limitation the
rights to use, copy, modify, merge, publish, distribute, sublicense, and/or
sell copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in
all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.  IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS
IN THE SOFTWARE.
"""
__version__ = "1.0.0"

from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait


class RicohAddressBook(object):
    """ Manipulation of address book for multi function printer/scanner
    Ricoh Aficio MP C4502 via Web Image Monitor.
    """
    def __init__(self, url, username='admin', password=''):
        self.__driver = None  # init in __enter__, quit in __exit__
        self.__url = url
        self.__username = username
        self.__password = password

    def __enter__(self):
        """ initialize the browser driver, log into Web Image Monitor and
        load address book.
        """
        self.__driver = webdriver.Firefox()
        self.__driver.get(self.__url)
        WebDriverWait(self.__driver, 10).until(
            EC.title_contains("Web Image Monitor")
        )
        self.__driver.switch_to_frame("header")
        self.__driver.find_element(By.LINK_TEXT, "Login").click()
        WebDriverWait(self.__driver, 10).until(EC.title_contains("Login"))
        self.__driver.find_element(
            By.NAME, "userid_work"
        ).send_keys(self.__username)
        self.__driver.find_element(
            By.NAME, "password_work"
        ).send_keys(self.__password)
        for element in self.__driver.find_elements(By.TAG_NAME, "input"):
            if (element.get_attribute("type") == "submit"):
                element.click()
                break
        WebDriverWait(self.__driver, 10).until(
            EC.title_contains("Web Image Monitor")
        )
        self.__driver.switch_to_frame("work")
        menu_item = self.__driver.find_element(
            By.LINK_TEXT,
            "Device Management"
        )
        ActionChains(self.__driver).move_to_element(
            menu_item
        ).release(
            menu_item
        ).perform()
        self.__driver.find_element(By.LINK_TEXT, "Address Book").click()

        def find_manual_input(driver):
            """ find dom element with link text "Manual Input"
            """
            return driver.find_element(By.LINK_TEXT, "Manual Input")

        WebDriverWait(self.__driver, 10).until(find_manual_input)
        find_manual_input(self.__driver).click()
        return self

    def __exit__(self, exc_type, exc_value, traceback):
        """ logout from Web Image Monitor and quit the browser
        """
        self.__driver.switch_to_default_content()
        self.__driver.switch_to_frame("header")
        self.__driver.find_element(By.LINK_TEXT, "Logout").click()
        self.__driver.quit()

    def __wait_for_completed(self):
        """ if loadingStatus is non-empty, wait for completion
        important when address book is loading first time
        subsequent requests seems to use ajax, and then loadingStatus is empty
        """
        locator = (By.ID, "span_loadingStatus")
        find_loading_status = lambda x: x.find_element(locator[0], locator[1])
        WebDriverWait(self.__driver, 10).until(find_loading_status)
        if not find_loading_status(self.__driver).text == u'':
            WebDriverWait(self.__driver, 30).until(
                EC.text_to_be_present_in_element(locator, "Completed")
            )

    def __wait_for_addressbook(self):
        """ switch to correct frame and wait if necessary
        """
        self.__driver.switch_to_default_content()
        self.__driver.switch_to_frame("work")
        self.__wait_for_completed()

    @staticmethod
    def get_tag_label(name):
        """ Select tag label based on first character in name
        """
        char = name[0].upper()
        if char in 'ÆÅÄ':
            char = 'A'
        if char in 'ØÖ̈́':
            char = 'O'
        for tag in (
            'AB',
            'CD',
            'EF',
            'GH',
            'IJK',
            'LMN',
            'OPQ',
            'RST',
            'UVW',
            'XYZ',
        ):
            if char in tag:
                return tag
        return None

    @staticmethod
    def pad_userid(userid):
        """ Convert userid from integer to padded string
        """
        assert isinstance(userid, (int, long)) and 0 < userid <= 50000
        return "%05d" % userid

    def __select_user(self, userid):
        """ Select user in address book. return True if userid exists,
        otherwise False and no change in selection
        """
        found = False
        for option in self.__driver.find_elements(By.NAME, "entryindex"):
            if option.get_attribute("value") == userid:
                option.click()
                found = True
                break
        return found

    def write_user(self, userid, name, email):
        """ Add or update user in address book
        userid - integer i, 0 < i <= 50000
        name - unicode string with users full name
        email - unicode string with users email address
        """
        userid = self.pad_userid(userid)

        self.__wait_for_addressbook()
        userid_exists = self.__select_user(userid)
        if userid_exists:
            self.__driver.find_element(By.PARTIAL_LINK_TEXT, "Change").click()
        else:
            self.__driver.find_element(
                By.PARTIAL_LINK_TEXT,
                "Add User"
            ).click()
        self.__driver.switch_to_alert()
        for key, value in {
            'entryIndexIn': userid,
            'entryNameIn': name,
            'entryDisplayNameIn': name,
            'mailAddressIn': email,
        }.items():
            find_element = lambda x, y = key: x.find_element(By.NAME, y)
            WebDriverWait(self.__driver, 10).until(find_element)
            element = find_element(self.__driver)
            element.clear()
            element.send_keys(value)
        tag_label = self.get_tag_label(name)
        if tag_label is not None:
            tag = Select(self.__driver.find_element(By.NAME, "entryTagInfoIn"))
            tag.select_by_visible_text(tag_label)
        if not userid_exists:
            # New users should not be added automatically to frequent user
            # list.  For updates - do not touch this - allows to set this
            # manually via other interfaces
            for option in self.__driver.find_elements(
                By.NAME,
                "entryTagInfoIn"
            ):
                if option.get_attribute("value") == '2':
                    option.click()
                    break
        popup = self.__driver.find_element(By.ID, "additional")
        self.__driver.find_element(By.LINK_TEXT, "OK").click()
        popup_not_displayed = lambda x: not popup.is_displayed()
        WebDriverWait(self.__driver, 10).until(popup_not_displayed)

    def remove_user(self, userid):
        """ Delete user from address book
        userid - integer i, 0 < i <= 50000
        """
        userid = self.pad_userid(userid)

        self.__wait_for_addressbook()
        if self.__select_user(userid):
            self.__driver.find_element(By.LINK_TEXT, "Delete").click()
            find_yes = lambda x: x.find_element(By.LINK_TEXT, "Yes")
            WebDriverWait(self.__driver, 10).until(find_yes)
            popup = self.__driver.find_element(By.ID, "additional")
            find_yes(self.__driver).click()
            popup_not_displayed = lambda x: not popup.is_displayed()
            WebDriverWait(self.__driver, 10).until(popup_not_displayed)


if __name__ == '__main__':
    with RicohAddressBook('http://our-printer.example.com') as abook:
        abook.write_user(50000, "John Doe", "john.doe@example.com")
        abook.write_user(50000, "John Alexander Doe", "john.doe@example.com")
        abook.remove_user(50000)
