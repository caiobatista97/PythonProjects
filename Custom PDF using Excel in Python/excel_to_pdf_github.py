# Focus Group Data Analytics
# Caio Batista
# Last updated: 6/15/22

# Start by installing the fpdf package on the terminal:
# python -m pip install fpdf

from fpdf import FPDF
from datetime import date
import pandas as pd
from PIL import Image

# The spreadsheet with all of the tabs
# Note: make sure Excel file is saved as "Excel Workbook (*.xlsx)"
doc_name = 'FC TEMPLATE FINAL.xlsx'
path = '/Users/caiobatista/Documents/Programming/VSCode/Vela/Python Projects/'
analysis_sheet = pd.ExcelFile(path + doc_name)

# Number of questions = number of tabs
num_questions = len(analysis_sheet.sheet_names)

# Content will be the content of the PDF file, it will have num_questions * 2 keys
# Each key will be a question or "QX Themes"
# Each value will either be a list of answers or a list of themes
content = {}

# This will loop through each tab (each question)
for i in range(num_questions):
    i += 1

    # Referencing each question, then reading the tab
    tab_name = 'Q' + str(i)
    this_sheet = pd.read_excel(path + doc_name, sheet_name = tab_name)   

    columns = this_sheet.columns
    question = columns[3]

    # This will get rid of any empty cells in the excel file
    answers = this_sheet[question]
    answers = [x for x in answers if type(x) == str]
    
    content[question] = answers

    theme_cols = columns[6:]
    all_themes = []

    for theme_col in theme_cols:
        this_theme_list = this_sheet[theme_col]
        this_theme_list = [x for x in this_theme_list if type(x) == str]
    
        # This will make sure the themes aren't repeated
        for theme in this_theme_list:
            if theme not in all_themes and len(theme) > 0:
                all_themes.append(theme)
    
    content[tab_name + ' Themes'] = all_themes
            
# Initializing/creating the PDF document, page and font
pdf = FPDF(orientation = "P", unit = "mm", format = "A4")
pdf.add_page()

# # creating a new image file with light blue color with A4 size dimensions using PIL
# img = Image.new('RGB', (210,297), "#e7ecef")
# img.save('blue_colored.png')

# # adding image to pdf page that e created using fpdf
# pdf.image('blue_colored.png', x = 0, y = 0, w = 210, h = 297, type = '', link = '')

font = 'Courier'

pdf.set_font(font, style = '', size = 12)
pdf.set_text_color(0,0,0)

# A4 paper has 210x297mm dimensions
w = 190
h = 5

for key, values in content.items():
    # Determining if the key is a question or a theme
    t = 'Themes' in key
    q = not t

    # The color of the box will depend on whether it's a question or a theme
    if q:
        pdf.set_fill_color(240, 93, 94)
    else:
        pdf.set_fill_color(222, 184, 65)

    # Text box with the question or QX Themes
    pdf.set_font(font, style = 'B', size = 12)
    pdf.multi_cell(w, h, txt = str(key), border = 0, align = 'L', fill = True)  
    
    # Text box with the answers or themes
    pdf.set_font(font, style = '', size = 10)
    pdf.set_fill_color(255, 255, 255)
    for value in values:
        pdf.multi_cell(w, h, txt = '- ' + str(value), border = 0, align = 'L', fill = True)  
    
    # If it's a theme, then put a line to separate from the next question
    if t:
        pdf.multi_cell(w, h, txt = '', border = 'B', align = 'L', fill = False) 
    else:
        pdf.multi_cell(w, h, txt = '', border = 0, align = 'L', fill = False) 


# Outputting the PDF file
today_date = date.today().strftime("%m_%d_%y")
pdf.output(path + '[REPORT]' + doc_name[:-5] + '.pdf')

# https://medium.com/@theprasadpatil/how-to-create-a-pdf-report-from-excel-using-python-b882c725fcf6#id_token=eyJhbGciOiJSUzI1NiIsImtpZCI6IjU4MGFkYjBjMzJhMTc1ZDk1MGExYzE5MDFjMTgyZmMxNzM0MWRkYzQiLCJ0eXAiOiJKV1QifQ.eyJpc3MiOiJodHRwczovL2FjY291bnRzLmdvb2dsZS5jb20iLCJuYmYiOjE2NTUzMTI4MjAsImF1ZCI6IjIxNjI5NjAzNTgzNC1rMWs2cWUwNjBzMnRwMmEyamFtNGxqZGNtczAwc3R0Zy5hcHBzLmdvb2dsZXVzZXJjb250ZW50LmNvbSIsInN1YiI6IjExNTYwMjM5MDEwNzA4MDU0Mjc2NyIsImVtYWlsIjoiY2Fpb2JhdGlzdGE5N0BnbWFpbC5jb20iLCJlbWFpbF92ZXJpZmllZCI6dHJ1ZSwiYXpwIjoiMjE2Mjk2MDM1ODM0LWsxazZxZTA2MHMydHAyYTJqYW00bGpkY21zMDBzdHRnLmFwcHMuZ29vZ2xldXNlcmNvbnRlbnQuY29tIiwibmFtZSI6IkNhaW8gQmF0aXN0YSIsInBpY3R1cmUiOiJodHRwczovL2xoMy5nb29nbGV1c2VyY29udGVudC5jb20vYS0vQU9oMTRHaFFTVGxFajViYzJ5aEc3c3BjUU5IYlpyOHJScThrZGJiZjlGWkJhXzQ9czk2LWMiLCJnaXZlbl9uYW1lIjoiQ2FpbyIsImZhbWlseV9uYW1lIjoiQmF0aXN0YSIsImlhdCI6MTY1NTMxMzEyMCwiZXhwIjoxNjU1MzE2NzIwLCJqdGkiOiI5NDRkYzVmMGYzZDU0Y2EzNGNmOGZhNTRhZDE4NmVkZTZlOGMxNjA4In0.DooIToaNikoR_SUW9GJ4dz5dkmx8HQsrxl-YrJyMWw7IGJZRXk9iBO2E88L0XRbpZuKRs8ZKviTpgVJTVVaopuUyYyglyv4pPKtIYaKQLTb3oQ_N8Ei_J7nPRuTwackRyped5UxdlmtGKveQlgLCgviVJU_OoBhbkdk8P0AnDYRXdafHEA3Aj-2hNqIeNJ3zSgfCw4wmhm8wbo5C4OcdVW25eYvFC6eXSeV9p-L2rDXNCgEjUhAz3OSzHdSvI5MAVNBrdohFrqgcBgSr48HsPW92oNe1v6ME8KgWrPRxlysWq38YuXqgAV3XibJNNbp5ISbEh2bBv46OSI7olNVBGg
# https://github.com/PBPatil/Python-Automation-Projects/blob/main/Automate%20PDF%20using%20Python/Excel_to_PDF_report.py
# https://betterprogramming.pub/how-to-create-a-pdf-in-python-71fac9f7bcd6
