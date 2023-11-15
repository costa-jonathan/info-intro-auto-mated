import os  # for finding the files and opening in default app (Windows)
import subprocess  # for macOS and Linux
import re  # for matching the matriculation number
import platform  # checking for operating systems
import shutil  # move everything
from pathlib import Path  # for making the _todel folder if it doesn't exist


def open_file(file_path): # if using this you don't need to distinguish between / and \ for the different file systems
    if platform.system() == 'Darwin':  # macOS
        if file_path.endswith('.pdf'):
            subprocess.call(['open', file_path.replace('\\', '/')])
            # webbrowser.open_new(file_path)
        else:
            subprocess.call(('open', file_path.replace('\\', '/')))
    elif platform.system() == 'Windows':  # Windows
        os.startfile(file_path.replace('/', '\\'))
    else:  # linux
        subprocess.call(('xdg-open', file_path.replace('\\', '/')))


homework_path = input('Input the path to the homework folder: ')  # getting the homework path
person = ''
temp_path = ''

if platform.system() == 'Windows':  # Windows
    temp_path = "\\".join(homework_path.split('\\')[:-2])  # getting the root path of the homework folder structure as in the Dropbox
    person = homework_path.split('\\')[-1]  # extracting the correcting person from the homework path
    Path(homework_path + '\\_todel').mkdir(parents=True, exist_ok=True)  # making the _todel folder in the homework folder to move the corrected homeworks to
else:  # linux & macOS
    temp_path = "/".join(homework_path.split('/')[:-2])
    person = homework_path.split('/')[-1]
    Path(homework_path + '/_todel').mkdir(parents=True, exist_ok=True)

print(f'Hello {person}. I am your not-personal homework assistant, glad to see you. Let\'s get started.')


for path, directories, files in os.walk(homework_path):  # getting all the files
    for i in range(len(files)):
        # iterating over the files
        if '_todel' not in path and (files[i].endswith('pdf') or files[i].endswith('zip') or files[i].endswith('7z')):
            # this block is for extracting name and matr_nr from the filenames and folder structure
            # name will always work (because it comes automatically from moodle)
            # but MatrNr might fail dependent on the naming of the file of the student
            matr_nr = ""
            if platform.system() == 'Windows':  # Windows
                name = path.split('\\')[-1].split('_')[0]
            else:  # linux & macOS
                name = path.split('/')[-1].split('_')[0]
            try:
                matr_nr = re.findall(r"\d{8}", files[i])[0]  # regex für MatrNr
            except IndexError:
                matr_nr = '0'
                print("Matriculation Number not found for "
                      + path + '/' + files[i] +
                      "; using name instead")
            print("Rating for: " + matr_nr + ' - ' + name)

            # opening the homework file in the default app:
            filepath = path + '\\' + files[i]

            open_file(filepath)


            # initiating a new line for storing the data for this student and the iteration variable i
            line = str(name) + "\t" + str(matr_nr) + "\t"
            i = 1
            # the logic is only designed with while because it enables backtracking. Otherwise i don't kow how to do it.
            # I already forgot how it works exactly so good luck understanding it
            while i < 5:
                rating = -1
                while i < 4:
                    rating = input('Points (0-2) for Task {}: '.format(i))
                    try:
                        rating = int(rating)
                        if 3 > rating > -1:
                            i += 1
                            break
                    except ValueError:
                        if rating == 'b' or rating == '0' or rating == 'n':
                            i -= 1
                            line = line[:len(line) - 2]
                            continue
                        else:
                            continue
                line += str(rating) + "\t"
                if i == 4:
                    comment = input('Enter a voluntary comment (or redo the whole grading with n, b or 0): ')
                    if comment == 'b' or comment == '0' or comment == 'n':
                        line = str(name) + "\t" + str(matr_nr) + "\t"
                        i = 1
                    else:
                        line += comment + '\n'

                        # writing the line in the temp text file which effectively saves the progress
                        # & moving the corrected homework into the _todel folder
                        if platform.system() == 'Windows':  # Windows
                            with open(f'{temp_path}\\Homework_{person}_temp.txt', 'a',
                                      encoding="utf-8") as md:
                                md.write(line)
                            i += 1
                            shutil.move(path, homework_path + "\\_todel")
                        else:  # linux & macOS
                            with open(f'{temp_path}/Homework_{person}_temp.txt', 'a',
                                      encoding="utf-8") as md:
                                md.write(line)
                            i += 1
                            shutil.move(path, homework_path + "/_todel")

input('Please transfer the data now to the Excel file (press enter to open the files)')
# opening the text and Excel file in the default app:
try:
    open_file(f'{temp_path}\\Homework_{person}_temp.txt')
    open_file(f'{temp_path}\\Homework_{person}.xlsx')
except FileExistsError and FileNotFoundError:
    'I cannot find the excel file - maybe you don\'t have it downloaded - in that case please open it in the browser'


input('Press press enter to finish your Homework correction')

print('You are done now - great job!')
