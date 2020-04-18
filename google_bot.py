from time import sleep
from urllib.error import HTTPError
from urllib.request import urlretrieve

import pandas as pd
from selenium import webdriver


class GoogleBot:
    def __init__(self, query, file_base, delay):
        # Activating the Chrome Driver
        self.filename = ''
        self.driver = webdriver.Chrome()
        url = 'https://www.google.com/search?q={q}'.format(q = query)

        # Navigating to the Google Search for particular product
        self.driver.get(url)

        # Navigating to the Google Images tab
        self.driver.find_element_by_xpath(
            '//*[@id="hdtb-msb-vis"]/div[2]/a') \
            .click()
        #

        sleep(delay)

        self.image_download(delay = delay, file_base = file_base, query = query,
                            thumbnail = '//*[@id="islrg"]/div[1]/div[1]/a[1]/div[1]/img',
                            image = '//*[@id="Sva75c"]/div/div/div[3]/div[2]/div/div[1]/div[1]/div/div[2]/a/img',
                            index = 1)
        self.image_download(delay = delay, file_base = file_base, query = query,
                            thumbnail = '//*[@id="islrg"]/div[1]/div[2]/a[1]/div[1]/img',
                            image = '//*[@id="Sva75c"]/div/div/div[3]/div[2]/div/div[1]/div[1]/div/div[2]/a/img',
                            index = 2)
        self.image_download(delay = delay, file_base = file_base, query = query,
                            thumbnail = '//*[@id="islrg"]/div[1]/div[3]/a[1]/div[1]/img',
                            image = '//*[@id="Sva75c"]/div/div/div[3]/div[2]/div/div[1]/div[1]/div/div[2]/a/img',
                            index = 3)
        self.image_download(delay = delay, file_base = file_base, query = query,
                            thumbnail = '//*[@id="islrg"]/div[1]/div[4]/a[1]/div[1]/img',
                            image = '//*[@id="Sva75c"]/div/div/div[3]/div[2]/div/div[1]/div[1]/div/div[2]/a/img',
                            index = 4)
        self.driver.close()

    def image_download(self, delay, file_base, query, thumbnail, image, index):
        try:
            # Navigating to the 1st thumbnail of the search and CLICKing in it
            self.driver.find_element_by_xpath(thumbnail) \
                .click()
            sleep(delay)

            # Finding the larger image and assigning its tag to the img variable
            img = self.driver.find_element_by_xpath(image)

            # Getting the URL out the img attribute
            src = img.get_attribute('src')

            self.filename = file_base + '{index}.jpg'.format(index = index)

            # Downloading the (bigger) image and storing it locally
            urlretrieve(src, self.filename)
        except FileNotFoundError as err:
            print(err)  # something wrong with local path
        except HTTPError as err:
            print(err)  # something wrong with url
        except:
            print(
                'Unknown Error Image {index} - {q}'.format(q = query, index = index))  # something unexpected went wrong
        else:
            # Letting the user know that the 1st image of the first product has been downloaded
            print('{q} - {filename} downloaded'.format(q = query, filename = self.filename))


print('========================================================================')
print('|           Google_Bot 1.0.0.5 by Allex Radu [www.ATFR.net]             |')
print('|     Get the latest version at https://github.com/allexradu/gBot       |')
print('========================================================================')
print('| Instructions: Save your Excel Workbook as "a.xls" and place it in     |')
print('| the same folder as this file, make sure the file in not opened.       |')
print('========================================================================')
print('|      WARNING!!! WRITE THIS DOWN! To stop the bot press CTRL + C       |')
print('========================================================================')

while True:
    try:
        seconds_delay = float(input('Number of seconds delay from one image to another: ' +
                                    '\n (the slower the computer / connection the higher the number)' +
                                    '\n [Minimum 1 sec recommended] seconds: '))
    except:
        print('Invalid Input!!! Try again!')
    else:
        break

# Reading first column of a local excel file
try:
    df = pd.read_excel('a.xls', sheet_name = 0)
    print('Excel Read Complete!')

    product_names = df['Name'].tolist()

    for i in range(len(product_names)):
        print(product_names[i])

    # Placing all the product names in a list

    for i in range(len(product_names)):
        file_name_base = 'A' + '{num}'.format(num = (100 + i))
        product_name = product_names[i]
        GoogleBot(product_name, file_name_base, seconds_delay)
except:
    print('Excel File NOT READ. Name your file "a.xls" with the first column "Name"')
    print('and place it in the same directory and the bot file.')
