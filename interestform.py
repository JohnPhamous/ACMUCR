import openpyxl, getpass, os, time

global counter
global first_name

counter = 2
first_name = None

# Spreadsheet set up
signin = openpyxl.Workbook()
signin_sheet = signin.get_active_sheet()
signin_sheet.title = "ACM Interest Form 8-4-16"
signin_sheet["A1"] = "First Name"
signin_sheet["B1"] = "Last Name"
signin_sheet["C1"] = "Email"
signin_sheet["D1"] = "Major"

def printHeader():
    os.system('clear')
    logo = "         ___       ______ .___  ___. \n        /   \     /      ||   \/   | \n       /  ^  \   |  ,----'|  \  /  | \n      /  /_\  \  |  |     |  |\/|  | \n     /  _____  \ |  `----.|  |  |  | \n    /__/     \__\ \______||__|  |__| \n UC Riverside Summer Orientation 8/4/2016"
    print("\n")
    print("\n".join('{:^170}'.format(s) for s in logo.split("\n")))
    print(("/------------------------------------------------------------------------------------------------\\").center(525))
    print(("|  Hello! Thanks for showing interest in the Association of Computing Machinery at UC Riverside!  |").center(525))
    print(("\\------------------------------------------------------------------------------------------------/").center(525))

def userInput():
    global counter
    global first_name

    first_name = input("(1/4). What's your first name? \n")
    if first_name != "exit":
        last_name = input("\n(2/4). What's your last name? \n")
        email = input("\n(3/4). What's your email? \n")
        major = input("\n(4/4). What's your major? \n")
        signin_sheet['A' + str(counter)] = str(first_name).upper()
        signin_sheet['B' + str(counter)] = str(last_name).upper()
        signin_sheet['C' + str(counter)] = str(email).upper()
        signin_sheet['D' + str(counter)] = str(major).upper()
        counter += 1

while first_name != "exit":
    printHeader()
    userInput()

signin.save("Orientation-8-4-16.xlsx")
print("Sheet saved")
