import platform
from time import sleep
from urllib.error import HTTPError
from urllib.error import URLError
from urllib.request import urlretrieve

from selenium import webdriver
from selenium.common.exceptions import ElementClickInterceptedException
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.keys import Keys
from socket import error as SocketError
import excel


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

        try:
            for product_name in self.product_names:
                # Lopping over each product name
                self.download_each_product(product_name = product_name)
                self.product_index += 1
                # Updating the data in Excel
                excel.add_data_to_excel()
                print('product index is ', self.product_index)
        except TypeError:
            print('Excel file not found')

        sleep(delay)

        # Closing the Chrome Window
        self.driver.close()

    def download_each_product(self, product_name):
        product_name = product_name

        # Finding the search Box of Google Images
        try:
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
        except KeyboardInterrupt:
            print('Keyboard was interrupted')


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

            system_prefix = 'excel\\photos\\' if platform.system() == 'Windows' else 'excel/photos/'

            filename = 'A' + '{num}'.format(num = (100 + obj.product_index)) + '{image_index}.jpg'.format(
                image_index = each_image_index + 1)
            filename_with_path = system_prefix + filename

            # Downloading the (bigger) image and storing it locally
            urlretrieve(src, filename_with_path)

        except NoSuchElementException:
            add_na_to_excel(each_image_index = each_image_index, product_index = obj.product_index,
                            product_name = product_name, error_text = 'No image found for Product:')
            # no images found on this search
            print('No image found for Product: {product}'.format(product = product_name))
        except FileNotFoundError:
            add_na_to_excel(each_image_index = each_image_index, product_index = obj.product_index,
                            product_name = product_name, error_text = 'File error on downloading image for product:')
            # something wrong with local path
        except HTTPError:
            add_na_to_excel(each_image_index = each_image_index, product_index = obj.product_index,
                            product_name = product_name, error_text = 'HTTP error on downloading image for product:')
        except URLError:
            add_na_to_excel(each_image_index = each_image_index, product_index = obj.product_index,
                            product_name = product_name, error_text = 'URL error on downloading image for product:')
        except ElementClickInterceptedException:
            add_na_to_excel(each_image_index = each_image_index, product_index = obj.product_index,
                            product_name = product_name,
                            error_text = 'The image click was interrupted on downloading product:')
        except SocketError:
            add_na_to_excel(each_image_index = each_image_index, product_index = obj.product_index,
                            product_name = product_name,
                            error_text = 'TRY INCREASING THE DELAY, you got SOCKET ERROR on product:')
        # except:
        #     # something unexpected went wrong
        #     print(
        #         'Unknown Error Image {image_index} - {q}'.format(q = query, image_index = image_index))

        else:
            key = 'Image %(column_image_number)s' % dict(column_image_number = each_image_index + 1)

            # Uncomment this if you want the excel file to be clickable on a non-Windows machine
            # link_prefix = 'photos\\' if platform.system() == 'Windows' else 'photos/'

            link_prefix = 'photos\\'

            link = '=HYPERLINK("{link_prefix}{filename}","{filename}")'.format(link_prefix = link_prefix,
                                                                               filename = filename)
            excel.downloaded_images[key].insert(obj.product_index, link)

            # Letting the user know that the 1st image of the first product has been downloaded
            print('{q} - {filename} downloaded'.format(q = product_name, filename = filename))


# Writing "n/a" in all the cells that the bot couldn't download a image
def add_na_to_excel(each_image_index, product_index, product_name, error_text):
    key = 'Image %(column_image_number)s' % dict(column_image_number = each_image_index + 1)
    excel.downloaded_images[key].insert(product_index, 'n/a')
    print('{error_text} {product}'.format(error_text = error_text, product = product_name))


def get_the_thumbnail_xpath(number_of_images):
    # Adding the x paths of the thumbnails to the list
    thumbnails_xpath = []
    for index in range(number_of_images):
        thumbnails_xpath.append(
            '//*[@id="islrg"]/div[1]/div[{index}]/a[1]/div[1]/img'.format(index = index + 1))
    return thumbnails_xpath
