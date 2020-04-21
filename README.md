# Google Bot for eCommerce
by Allex Radu

This bot allows you to quickly download product images from your online shop using Google Image Search.

All you need is a Excel Spreadsheet

---------------------------------------

Download requirements using pip install:

setuptools

pandas

selenium

xlrd

openpyxl

---------------------------------------
Install Selenium Chrome Driver

Place it on a directory c:\driver and update your path variable to include this driver.

To update your path variable go to Start > Search "Control Panel" > Search "Path" > Edit environment variables for your account > Path > And copy the path your placed your driver like:

C:\driver

(no backslash at the end)

Download the driver from the link bellow and make sure it matches your Chrome Version (Check your Google Chrome version by going to Three Dots (top right corner) > Help > About Google Chrome )

https://chromedriver.chromium.org/downloads 

 ---------------------------------------
 
Step 1. Put your products name in column 1 and name it "Name"

Step 2. Save your workbook as "a.xlsx" and place it excel folder

Step 3. Execute this software using "python google_bot.py"

Step 4. Add the delay between images in seconds using a number that can, optionally, include a decimal point [1 recommended]. 

Note. A lower delay than 1 second my result in lower quality images.

Step 5. Pick how many photos / search do you want the bot do download starting with 1 with a maximum of 9.

Note: The bot with download images with the file name stating with A1001.jpg (for the first image of the first product), A1002.jpg (for the second image), A1003.jpg (for the third), A1004.jpg (for the forth) ... A1009.jpg (for the ninth).

Step 6. Wait for the robot to finish processing your product names, the robot will write in your excel file the link to image files so you can review them.

IMPORTANT! All the image cells contain LINKS if you want to import your file please press Select columns, and paste special as Values.

Remember: some images might not be in the JPEG format even though they have the suffix .jpg, we recommend converting all images with a batch converter just in case. 



 
