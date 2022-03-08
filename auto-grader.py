import numpy as np
from selenium.webdriver import Chrome
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import Select
from webdriver_manager.chrome import ChromeDriverManager
import time
from selenium.webdriver.common.by import By
import pandas as pd
from selenium.webdriver.common.keys import Keys
import pyautogui
import os


# todo: if help neede contact Jonathan Costa @ guithub: https://github.com/costa-jonathan
# Hard coded stuff:
# 1. Moodle user credentials
moodle_user = ""  # todo: put in your moodle username in the string
moodle_pw = ""  # todo: put your moodle password in the string
# 2. Grading page (has to be updated every week!
page = input('Enter link of grading page, for example: '
             '\"www.moodle.tum.de/mod/assign/view.php?id=9999999&action=grading\". '
             'Must be from the tab \"Alle anzeigen\".\nEnter here: ')
# 3. Excel files directory:
excel_folder = "C:\\Users\\Name\\Dropbox\\EIDI WS 21_22\\Homework\\" # todo: replace the string with the file path for the excel files
# 4. columns:
columns = ['Name', 'MatrNr', 'Comment', 'Sum']
# (5. week of correction)
week = int(input('Enter week number: '))
# emails
emails = {
        'Jonathan': 'jonathan@realmail.com',
        'Santa': 'santa@northpole.np'
          } # todo: fill the dictonary with 'Name': 'email-adress' pairs. The names must be the same as the excel sheets are named!
# standard comment text: todo: Change this text if you want to change the standard text of the comment under the students submission
standard_comment = '\n\n This Homework was corrected by {}. If you need an explanation for the grade please reach out to {} however please consider first comparing your solution to the solutions sheet.'


def fix_matr_nr(nr):  # just to get a clear number; if not done: can lead to errors because of excel formatting
    try:
        return str(int(nr))
    except ValueError:
        return np.NaN  # str(nr)


grading_sheets = {}
grading_errors = pd.DataFrame(columns=columns)
# grading_sheets = pd.DataFrame(columns=columns)
for file in os.listdir(excel_folder):
    if file.endswith('.xlsx') and not file.startswith('~$'):  # and file != 'Homework_Isabella.xlsx'
        print(file)
        grading_temp = pd.read_excel(pd.ExcelFile(excel_folder + file), f'HA{week:02d}')[columns]
        grading_temp['MatrNr'] = grading_temp['MatrNr'].apply(lambda x: fix_matr_nr(x))
        grading_sheets[file.split('_')[1][:-5]] = grading_temp

webDriver = Chrome(service=Service(ChromeDriverManager().install()))
webDriver.maximize_window()
webDriver.get(page)

pyautogui.keyDown('ctrl')
pyautogui.press('t')
pyautogui.press('tab')
pyautogui.keyUp('ctrl')

# login
webDriver.find_element(By.XPATH, "//a[text()='TUM LOGIN']").click()
webDriver.find_element(By.XPATH, "//input[@id='username']").send_keys(moodle_user)
webDriver.find_element(By.XPATH, "//input[@id='password']").send_keys(moodle_pw)
webDriver.find_element(By.XPATH, "//input[@id='password']").send_keys(Keys.ENTER)

"//tr[@id='mod_assign_grading-3387058_r0']/td[@id='mod_assign_grading-3387058_r0_c3']"

select = Select(webDriver.find_element(By.ID, 'id_filter'))
select.select_by_value('submitted')
time.sleep(10)
select = Select(webDriver.find_element(By.ID, 'id_perpage'))
select.select_by_value('-1')


mtr_nrs = webDriver.find_elements(By.XPATH, "//tr[contains(@class, 'unselectedrow')]")
online_list_name = {str(k.find_element(By.XPATH, ".//td[contains(@class, 'c2')]").text): k for k in
                    mtr_nrs}
online_list_matr_nr = {str(k.find_element(By.XPATH, ".//td[contains(@class, 'c3')]").text): k for k in
                       mtr_nrs}

main_window = webDriver.current_window_handle


# this is the first loop over all grades in the Excel sheets and comparing the NAMEs
for person in grading_sheets.keys():
    print(f'Grading for {person}')
    grading = grading_sheets[person]
    grading_len = len(grading)
    for line_number, (index, row) in enumerate(grading.iterrows()):
        print("Currently on row: {}; Currently iterated {}% of rows".format(
            line_number, round(100 * (line_number + 1) / grading_len, 1)))
        try:
            web_row_element = online_list_name[(str(row['Name'].strip()))]
            link = web_row_element.find_element(By.XPATH, ".//td[contains(@class, 'c6')]/a").get_attribute(
                'href')

            webDriver.switch_to.window(webDriver.window_handles[1])
            webDriver.get(link)

            input_field_grade = ''
            try:
                time.sleep(2)
                input_field_grade = webDriver.find_element(By.XPATH, "//input[@name='grade']")
            except:
                time.sleep(3)
                input_field_grade = webDriver.find_element(By.XPATH, "//input[@name='grade']")
            input_field_grade.clear()
            input_field_grade.send_keys(row['Sum'])

            try:
                time.sleep(1)
                input_field_grade = webDriver.find_element(By.XPATH,
                                                           "//div[@id='id_assignfeedbackcomments_editoreditable']")
            except:
                time.sleep(1)
                input_field_grade = webDriver.find_element(By.XPATH,
                                                           "//div[@id='id_assignfeedbackcomments_editoreditable']")

            input_field_grade.clear()
            if str(row['Comment']) != "nan" and str(row['Comment']) != "NaN":


                input_field_grade.send_keys(str(row['Comment']) + standard_comment.format(person, emails[person]))
            else:
                input_field_grade.send_keys(standard_comment.format(person, emails[person]))

            # silent mode: # todo: remove the # of the line under this one to enable 'silent mode' which will not notify the students of their mark being entered
            # webDriver.find_element(By.XPATH, "//input[@name='sendstudentnotifications' and @type='checkbox']").click()

            webDriver.find_element(By.XPATH, "//button[@name='savechanges']").click()
            time.sleep(1.5)
            webDriver.switch_to.window(webDriver.window_handles[0])
            grading.drop(line_number, inplace=True)

        except (KeyError, AttributeError) as error:
            continue


# this second loop (just copied because I am too lazy to think of an elegant solution) does the same but checks the
# MATRNR of the students
for person in grading_sheets.keys():
    print(f'Grading for {person}')
    grading = grading_sheets[person]
    for line_number, (index, row) in enumerate(grading.iterrows()):
        try:
            web_row_element = online_list_matr_nr[("0" + row['MatrNr'])]
            link = web_row_element.find_element(By.XPATH, ".//td[contains(@class, 'c6')]/a").get_attribute(
                'href')

            webDriver.switch_to.window(webDriver.window_handles[1])
            webDriver.get(link)

            input_field_grade = ''
            try:
                time.sleep(2)
                input_field_grade = webDriver.find_element(By.XPATH, "//input[@name='grade']")
            except:
                time.sleep(3)
                input_field_grade = webDriver.find_element(By.XPATH, "//input[@name='grade']")
            input_field_grade.clear()
            input_field_grade.send_keys(row['Sum'])

            try:
                time.sleep(1)
                input_field_grade = webDriver.find_element(By.XPATH,
                                                           "//div[@id='id_assignfeedbackcomments_editoreditable']")
            except:
                time.sleep(1)
                input_field_grade = webDriver.find_element(By.XPATH,
                                                           "//div[@id='id_assignfeedbackcomments_editoreditable']")

            input_field_grade.clear()
            if str(row['Comment']) != "nan" and str(row['Comment']) != "NaN":

                input_field_grade.send_keys(str(row['Comment']) + standard_comment.format(person, emails[person]))
            else:
                input_field_grade.send_keys(standard_comment.format(person, emails[person]))

            # silent mode (because second loop might do some double time):
            webDriver.find_element(By.XPATH, "//input[@name='sendstudentnotifications' and @type='checkbox']").click()

            webDriver.find_element(By.XPATH, "//button[@name='savechanges']").click()
            time.sleep(1.5)
            webDriver.switch_to.window(webDriver.window_handles[0])
            grading.drop(line_number, inplace=True)

        except (KeyError, AttributeError, TypeError) as error:
            print(row)
            continue


# a printout in the console of the failed lines
for person in grading_sheets.keys():
    grading = grading_sheets[person]
    print(grading)


# save said failed lines in a text file
for person in grading_sheets.keys():
    grading = grading_sheets[person]
    with open(f'{excel_folder}failed_inputs\\{week}.txt', "a") as file_object:
        file_object.write(person + "\n")
    grading.to_csv(f'{excel_folder}failed_inputs\\{week}.txt', index=None, sep='\t', mode='a')
