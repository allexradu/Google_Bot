import gBot
import excel


def main():
    number_of_images = 0
    product_names = excel.read_excel_first_column()
    delay = 0

    print('========================================================================')
    print('|           Google_Bot 1.1.0.0 by Allex Radu [www.ATFR.net]             |')
    print('|     Get the latest version at https://github.com/allexradu/gBot       |')
    print('========================================================================')
    print('| Instructions: Save your Excel Workbook as "a.xls" and place it in     |')
    print('| the same folder as this file, make sure the file in not opened.       |')
    print('========================================================================')
    print('|      WARNING!!! WRITE THIS DOWN! To stop the bot press CTRL + C       |')
    print('========================================================================')

    while True:
        try:
            no_of_images = int(input('Number of images per product [1-9]: '))
        except:
            print('Invalid Input!!! Try again!')
        else:
            if not (0 < no_of_images <= 9):
                print('Number not in range, try again!!')
                continue
            else:
                number_of_images = no_of_images
                break

    while True:
        try:
            seconds_delay = float(input('Number of seconds delay from one image to another: ' +
                                        '\n (the slower the computer / connection the higher the number)' +
                                        '\n [Minimum 1 sec recommended] seconds: '))
        except:
            print('Invalid Input!!! Try again!')
        else:
            delay = seconds_delay
            break

    gBot.Google(number_of_images = number_of_images, product_names = product_names, delay = delay)


if __name__ == '__main__':
    main()
