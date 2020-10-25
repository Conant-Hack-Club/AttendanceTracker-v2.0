import xlrd
from xlutils.copy import copy
from datetime import *
from pyfiglet import Figlet

'''
Attendance Spreadsheet
'''

#open file
path = ('attendance.xlsx')

#reading
rb = xlrd.open_workbook(path)
r_sheet = rb.sheet_by_index(0) #sheet for reading

wb = copy(rb)
w_sheet = wb.get_sheet(0) #sheet for writing

def updateData():
    global rb, r_sheet, wb, w_sheet

    rb = xlrd.open_workbook(path)
    r_sheet = rb.sheet_by_index(0)

    wb = copy(rb)
    w_sheet = wb.get_sheet(0)


'''
Commands Available
'''

today = date.today()
f = Figlet(font='slant')

def new_meeting():
    global rb, r_sheet, wb, w_sheet

    date = today.strftime("%m/%d/%y")

    w_sheet.write(0, r_sheet.ncols, date)

    wb.save(path)
    updateData()

    print(f.renderText('New Hacker Meeting Created'))


def add(firstname, lastname, email):

    w_sheet.write(r_sheet.nrows, 0, firstname.capitalize())
    w_sheet.write(r_sheet.nrows, 1, lastname.capitalize())
    w_sheet.write(r_sheet.nrows, 2, email)
    w_sheet.write(r_sheet.nrows, 3, 0)

    wb.save(path)
    updateData()

    print(f.renderText(firstname + ' has been added'))

def here(name):

    found = False

    for i in range(r_sheet.nrows):
        if r_sheet.cell_value(i, 0).lower() == name.lower():
            w_sheet.write(i, r_sheet.ncols-1, "X")
            found = True
            break

    if not found:
        print(f.renderText(name + " is not a real person"))
        return

    wb.save(path)
    updateData()

    print(f.renderText(name + ' is here!'))

def points(name, points):

    new_points = 0

    found = False

    for i in range(r_sheet.nrows):
        if r_sheet.cell_value(i, 0).lower() == name.lower():
            new_points = int(r_sheet.cell_value(i, 3)) + points
            w_sheet.write(i, 3, new_points)
            found = True
            break

    if not found:
        print(f.renderText(name + " is not a real person"))
        return

    wb.save(path)
    updateData()

    p = 'points'

    if(new_points == 1):
        p = 'point'

    print(f.renderText(name + ' now has ' + str(new_points) + " " + p))


def leaderboard():
    names = []
    points = []
    att = []

    for i in range(1, r_sheet.nrows):
        names.append(r_sheet.cell_value(i, 0))
        points.append(int(r_sheet.cell_value(i, 3)))
        a = 0
        for j in range(4, r_sheet.ncols):
            if r_sheet.cell_value(i, j) == "X":
                a += 1
        att.append(a)

    leader_1 =  sorted(zip(att, names), reverse=True)
    leader_2 =  sorted(zip(points, names), reverse=True)

    print(leader_1)

    i = 1
    print(f.renderText('attendance leaderboard'))
    for a in leader_1:
        print(str(i) + ". " + a[1] + " - " + str(a[0]))
        i += 1
    print("\n")

    i = 1

    print(f.renderText('points leaderboard'))
    for a in leader_2:
        print(str(i) + ". " + a[1] + " - " + str(a[0]))
        i += 1
    print("\n\n\n")


def current_meeting():

    current = r_sheet.cell_value(0,r_sheet.ncols - 1)

    if not isinstance(current, str): #date is saved as a float, which sometimes happens with spreadsheets
        dt = datetime.fromordinal(datetime(1900, 1, 1).toordinal() + int(current) - 2)
        current = dt.strftime("%m/%d/%y")

    print("\n\nThe current meeting is " + current + "\n\n")


'''
Main Loop
'''

print('-------------Commands-------------')
print('new meeting - creates new meeting')
print('current meeting - date of current meeting')
print('add <firstname> <lastname> <email> - adds person')
print('give <firstname> <number of points> - gives person points')
print('here <firstname> - mark person on attendance')
print('remove <firstname> <number of points> - remove person points')
print("leaderboard - see who is on top")
print('exit - I want to get out!')

print('----------------------------------')
print(f.renderText('WELCOME TO THE DATAFRAME FELLOW HACKER'))

while(True):
    a = input('what should I hack? ')

    commands = a.split(' ')

    if commands[0] == "new":
        new_meeting()
    elif commands[0] == "add":
        add(commands[1], commands[2], commands[3])
    elif commands[0] == "give":
        points(commands[1], int(commands[2]))
    elif commands[0] == "here":
        here(commands[1])
    elif commands[0] == "remove":
        points(commands[1], int(commands[2])*-1)
    elif commands[0] == "current":
        current_meeting()
    elif commands[0] == "leaderboard":
        leaderboard()
    elif commands[0] == "exit":
        break
    else:
        print(f.renderText('bad command'))


# new meeting - creates new meeting
# add <firstname> <lastname> <email> - adds person
# give <firstname> <number> - give them points
# here <firstname> - here that meeting
# remove <firstname> <number> - remove points
#current meeting