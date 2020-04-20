from time import sleep
from urllib.error import HTTPError
from urllib.error import URLError
from urllib.request import urlretrieve
import platform
import excel

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import ElementClickInterceptedException


#
class Google(object):
    def __init__(self, number_of_images, product_names, delay):
        # Activating the Chrome Driver
        self.driver = webdriver.Chrome()
        self.number_of_images = number_of_images
        self.product_names = product_names
        self.product_index = 0
        self.delay = delay

        url = 'https://www.google.com/search?q={q}'.format(q = 'Allex Radu')

        # Navigating to the Google Search for particular product
        self.driver.get(url)

        sleep(delay)

        # Navigating to the Google Images tab
        self.driver.find_element_by_xpath(
            '//*[@id="hdtb-msb-vis"]/div[2]/a') \
            .click()

        sleep(delay)

        for product_name in self.product_names:
            self.download_each_product(product_name = product_name)
            self.product_index += 1
            print('product index is ', self.product_index)

        sleep(10)

        # Closing the Chrome Window
        self.driver.close()

    def download_each_product(self, product_name):
        product_name = product_name

        # Finding the search Box of Google Images
        elem = self.driver.find_element_by_name('q')

        sleep(self.delay)

        # Clearing the search box
        elem.clear()

        sleep(self.delay)

        # Typing the new search
        elem.send_keys(product_name)

        sleep(self.delay)

        # Pressing ENTER (RETURN)
        elem.send_keys(Keys.RETURN)

        download_each_image(obj = self, product_name = product_name)


def download_each_image(obj, product_name):
    thumbnails_xpath = get_the_thumbnail_xpath(obj.number_of_images)
    larger_image_xpath = '//*[@id="Sva75c"]/div/div/div[3]/div[2]/div/div[1]/div[1]/div/div[2]/a/img'

    for each_image_index in range(obj.number_of_images):

        try:
            # Navigating to the 1st thumbnail of the search and CLICKing in it
            obj.driver.find_element_by_xpath(thumbnails_xpath[each_image_index]) \
                .click()
            sleep(obj.delay)

            # Finding the larger image and assigning its tag to the img variable
            img = obj.driver.find_element_by_xpath(larger_image_xpath)
            assert " did not match any image results." not in obj.driver.page_source

            # Getting the URL out the img attribute
            src = img.get_attribute('src')

            if platform.system() == 'Windows':
                file_name_base = 'excel\\photos\\' + 'A' + '{num}'.format(num = (100 + obj.product_index))
            else:
                file_name_base = '../excel/photos/' + 'A' + '{num}'.format(num = (100 + obj.product_index))

            filename = file_name_base + '{image_index}.jpg'.format(image_index = each_image_index + 1)

            # Downloading the (bigger) image and storing it locally
            urlretrieve(src, filename)

        except NoSuchElementException:
            # no images found on this search
            print('No image found for Product: {product}'.format(product = product_name))
        except FileNotFoundError:
            # something wrong with local path
            print('File error on downloading image for product: {product} '.format(product = product_name))
        except HTTPError:
            print('HTTP error on downloading image for product: {product} '.format(product = product_name))
        except URLError:
            print('URL error on downloading image for product: {product} '.format(product = product_name))
        except ElementClickInterceptedException:
            print('The image click was interrupted on downloading product: {product} '.format(product = product_name))
        # except:
        #     # something unexpected went wrong
        #     print(
        #         'Unknown Error Image {image_index} - {q}'.format(q = query, image_index = image_index))

        else:
            # Letting the user know that the 1st image of the first product has been downloaded
            print('{q} - {filename} downloaded'.format(q = product_name, filename = filename))


def get_the_thumbnail_xpath(number_of_images):
    # Adding the x paths of the thumbnails to the list
    thumbnails_xpath = []
    for index in range(number_of_images):
        thumbnails_xpath.append(
            '//*[@id="islrg"]/div[1]/div[{index}]/a[1]/div[1]/img'.format(index = index + 1))
    return thumbnails_xpath
