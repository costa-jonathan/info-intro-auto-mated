import numpy as np
import pandas as pd
import os
import time

from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver import Chrome
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By

import bs4 as bs

# Hard coded stuff:
# 1. Moodle user credentials
moodle_user = ""  # todo: put in your moodle username in the string
moodle_pw = ""  # todo: put your moodle password in the string
# 2. Grading page (has to be updated every week!
page = input('Enter link of grading page, for example: ' 
          '\"www.moodle.tum.de/mod/assign/view.php?id=9999999&action=grading\". '
          'Must be from the tab \"Alle anzeigen\".\nEnter here: ')
# 3. Excel files directory:
excel_folder = ''  # todo: replace the string with the file path for the excel files
# 4. columns:
columns = ['Name', 'MatrNr', 'Comment', 'Task 1 (2P)', 'Task 2 (2P)', 'Task 3 (2P)']
# (5. week of correction)
week = int(input('Enter week number: '))
# emails
emails = {
        'Jonathan': 'jonathan@realmail.com',
        'Santa': 'santa@northpole.np'
}   # todo: fill the dictonary with 'Name': 'email-adress' pairs. The names must be the same as the excel sheets are named!
# standard comment text: todo: Change this text if you want to change the standard text of the comment under the students submission
standard_comment = '\n\n This Homework was corrected by {}. If you need an explanation for the grade please reach out to {} however please consider first comparing your solution to the solutions sheet.'

stats = [0,0]

def fix_matr_nr(nr):  # just to get a clear number; if not done: can lead to errors because of excel formatting
    try:
        return str(int(nr))
    except ValueError:
        return np.NaN  # str(nr)


def fill_grade(row):
    input_field_grade = ''
    try:
        time.sleep(2)
        input_field_grade = webDriver.find_element(By.XPATH, "//input[@name='grade']")
        stats[0] += 1
    except NoSuchElementException:
        time.sleep(6)
        input_field_grade = webDriver.find_element(By.XPATH, "//input[@name='grade']")
        stats[1] += 1
    input_field_grade.clear()
    input_field_grade.send_keys(row['Sum'])

    input_field_comment = webDriver.find_element(By.XPATH,
                                               "//div[@id='id_assignfeedbackcomments_editoreditable']")

    input_field_comment.clear()
    if str(row['Comment']) != "nan" and str(row['Comment']) != "NaN":
        input_field_comment.send_keys(str(row['Comment']) + standard_comment.format(person, emails[person]))
    else:
        input_field_comment.send_keys(standard_comment.format(person, emails[person]))

    # silent mode: # todo: remove the # of the line under this one to enable 'silent mode' which will not notify the students of their mark being entered
    # webDriver.find_element(By.XPATH, "//input[@name='sendstudentnotifications' and @type='checkbox']").click()
    time.sleep(0.5)
    webDriver.find_element(By.XPATH, "//button[@name='savechanges']").click()
    time.sleep(1.5)


grading_sheets = {}
grading_errors = pd.DataFrame(columns=columns)
# grading_sheets = pd.DataFrame(columns=columns)
for file in os.listdir(excel_folder):
    if file.endswith('.xlsx') and not file.startswith('~$'):  # and file != 'Homework_Isabella.xlsx'
        print(file)
        grading_temp = pd.read_excel(pd.ExcelFile(excel_folder + file), f'HA{week:02d}')[columns]
        grading_temp.dropna(inplace=True, subset=['Task 1 (2P)']) #
        grading_temp['MatrNr'] = grading_temp['MatrNr'].apply(lambda x: fix_matr_nr(x))
        try:
            grading_temp['Sum'] = grading_temp['Task 1 (2P)'].astype('int') + grading_temp['Task 2 (2P)'].astype('int') + \
                                  grading_temp['Task 3 (2P)'].astype('int')
        except pd.errors.IntCastingNaNError:
            try:
                grading_temp['Sum'] = grading_temp['Task 1 (2P)'].astype('int') + grading_temp['Task 2 (2P)'].astype(
                    'int')
            except pd.errors.IntCastingNaNError:
                grading_temp['Sum'] = grading_temp['Task 1 (2P)'].astype('int')

        grading_sheets[file.split('_')[1][:-5]] = grading_temp


webDriver = Chrome(service=Service(ChromeDriverManager().install()))
webDriver.maximize_window()
webDriver.get(page)

# login
webDriver.find_element(By.XPATH, "//a[text()='TUM LOGIN']").click()
webDriver.find_element(By.XPATH, "//input[@id='username']").send_keys(moodle_user)
webDriver.find_element(By.XPATH, "//input[@id='password']").send_keys(moodle_pw)
webDriver.find_element(By.XPATH, "//input[@id='password']").send_keys(Keys.ENTER)
time.sleep(4)
html = webDriver.page_source
soup = bs.BeautifulSoup(html, features="lxml")
table = soup.find_all('table')
links = []
for tr in table[0].findAll("tr"):
    trs = tr.findAll("td")
    for each in trs:
        try:
            link = each.find('a')
            if link.get("class") == ['btn', 'btn-primary']:
                links.append(link['href'])
        except:
            pass

# get moodle name, mtrnr & links; the comment line is to check if it was already graded
moodle_table_read = pd.read_html(str(table))[0]

moodle_table = pd.DataFrame()
moodle_table[['name', 'MatrNr', 'comment']] = moodle_table_read[[r'VornameSortiert nach Vorname Aufsteigend / NachnameSortiert nach Nachname Aufsteigend', r'MatrikelnummerSortiert nach Matrikelnummer Aufsteigend', r'Feedback als Kommentar']]
moodle_table['MatrNr'] = moodle_table['MatrNr'].astype(int, errors='ignore').astype(str)
moodle_table = moodle_table[moodle_table['MatrNr'] != 'nan']
moodle_table['MatrNr'] = moodle_table['MatrNr'].str.slice(0, -2)
moodle_table['links'] = links
moodle_table_already_graded = moodle_table[moodle_table.comment.notnull()]

already_done = moodle_table_already_graded['name'].to_list()
already_done += moodle_table_already_graded['MatrNr'].to_list()

# override mode: redo also already done grades by uncommenting the line below & comment the line below that one out
# already_done = []
moodle_table = moodle_table[moodle_table.comment.isnull()]

name_lookup = moodle_table[['name', 'links']].set_index('name').to_dict()['links']
number_lookup = moodle_table[['MatrNr', 'links']].set_index('MatrNr').to_dict()['links']

for person in grading_sheets.keys():
    print(f'Grading for {person}')
    grading = grading_sheets[person]
    grading_len = len(grading)
    if grading.empty:
        continue
    else:
        for line_number, (index, row) in enumerate(grading.iterrows()):
            print("Currently on row: {}; Currently iterated {}% of rows".format(
                line_number, round(100 * (line_number + 1) / grading_len, 1)))
            try: # check name
                if row['Name'] in already_done:
                    grading.drop(line_number, inplace=True)
                    continue
                webDriver.get(name_lookup[row['Name']])
                fill_grade(row)
            except (KeyError, NoSuchElementException) as e:
                try:
                    if row['MatrNr'] in already_done:
                        grading.drop(line_number, inplace=True)
                        continue
                    webDriver.get(number_lookup[row['MatrNr']])
                    fill_grade(row)
                except (KeyError, NoSuchElementException) as e:
                    continue

            grading.drop(line_number, inplace=True)

# a printout in the console of the failed lines
for person in grading_sheets.keys():
    grading = grading_sheets[person]
    print(grading)

# save said failed lines in a text file
for person in grading_sheets.keys():
    grading = grading_sheets[person]
    if not grading.empty:
        with open(f'{excel_folder}failed_inputs/{week}.txt', "a") as file_object:
            file_object.write(person + "\n")
        grading.to_csv(f'{excel_folder}failed_inputs/{week}.txt', index=None, sep='\t', mode='a')


print(stats)