# What is SheetPoint?
Script to automate filling Sharepoint forms with information from Google Sheets. 
## How does it work?
1. Getting access to google sheet from where it gonna take information for filling forms.
2. Running web driver and logging into Sharepoint.
3. Starting filling forms with information from sheet one by one. Single form is filled with information from single row of sheet.
4. After form filled it checks if element of next page appears. If it doesn`t appear it leaves error message and starts next row after page refresh.
5. Repeat from step 3.
## Can I use it somehow and what would i need to start?
Yes, of course, but it needs good understanding of how it runs - it deals with your personal accounts and data and this information is needed to be inserted into code or from console.

You gonna need web driver. I use Chrome (https://sites.google.com/a/chromium.org/chromedriver/downloads). But you can try another which supports Selenium.

Credentials to access Google Sheets. (http://gspread.readthedocs.io/en/latest/oauth2.html). If experiencing any trouble with it read gspread documentation.

Basic knowledge of gspread and Selenium (for every form you need to build special module). Also it helps to understan code.

Dependicals specified in "requirements.txt"
## WHYYYYYYY???
Project started when occured need to fill many different sharepoint forms with big amount of information i had in sheets. And I am too lazy to perform it manually. So i wrote this. If u have simmilar situation these files can help you or at least give some ideas.
