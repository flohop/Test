import wikipedia
import urllib.request
import requests
from bs4 import BeautifulSoup
import os
import shutil


class WikiPage():
    """"enter name of wikipedia page and shortcut of language"""
    def __init__(self, page_name, language="de"):
        self.language = language
        wikipedia.set_lang(self.language)
        self.page_name = page_name
        self.w_page = wikipedia.search(page_name)
        try:
            self.w_page = self.w_page[0]
        except IndexError:
            pass
        self.wik_page = wikipedia.page(self.w_page)

    def get_title(self):
        """returns the title of the wikipedia page """
        if "." in self.wik_page.title:
            position = self.wik_page.title.find(".")

            self.wik_page.title = self.wik_page.title[0:position]
        else:
            pass
        return self.wik_page.title

    def get_summary(self):
        """returns the summary of the wikipedia page """
        return self.wik_page.summary

    def get_section(self, section):
        """get all categories of the wikipedia page"""
        return self.wik_page.section(section)

    def get_content(self, summary):
        """returns the full content of the wikipedia section """
        content = str(self.wik_page.content).replace(summary, "")
        for letter in content:
            if letter.startswith("=="):
                print("I am a heading")

        return content

    def get_pictures(self):
        """returns the links of all the pictures of the wikipedia page """

        return self.wik_page.images

    def download_picture(self):
        """ downloads the main image of a page and save it in a specific folder"""

        """ get the url of the wiki page"""
        try:
            if " " in self.wik_page.title:
                wik_page_url = self.wik_page.title.replace(" ", "_")
            else:
                wik_page_url = self.wik_page.title
            wiki_url = "https://" + str(self.language) + ".wikipedia.org/wiki/" + str(wik_page_url)

            """find the important image"""
            html_url = requests.get(wiki_url).content
            soup = BeautifulSoup(html_url, "lxml")
            try:
                image_box = soup.find("div", {"class": "my-parser-output"})
                image = image_box.find_all("a", {"class": "image"})
                logo_url = image[0].img.get("src")
            except AttributeError:
                image_box = soup.find("div", {"class": "thumbinner"})
                image = image_box.find("img")
                logo_url = image.get("src")

            """download the image"""
            urllib.request.urlretrieve("https:" + logo_url, self.wik_page.title + "_logo.jpg")
            picture_name = str(self.wik_page.title).lower() + "_logo.jpg"

            """ move image to \logo_pictures folder """
            destination = os.getcwd() + "\logo_pictures"
            try:
                shutil.move(picture_name, destination)
            except shutil.Error:
                pass

            """delete file in old directory"""
            for file in os.listdir():
                if file.endswith(".jpg"):
                    pass
        except AttributeError:
            image_box = soup.find("table", {"class": "toccolours float-right toptextcells"})
            image = image_box.find("img", {"alt": "Logo"})
            logo_url = image.get("src")

            """download the image"""
            urllib.request.urlretrieve("https:" + logo_url, self.wik_page.title + "_logo.jpg")
            picture_name = str(self.wik_page.title).lower() + "_logo.jpg"

            """ move image to \logo_pictures folder """
            destination = os.getcwd() + "\logo_pictures"

            try:
                shutil.move(picture_name, destination)
            except shutil.Error:
                pass

        try:
            return picture_name
        except UnboundLocalError:
            pass



