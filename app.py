import sqlite3, datetime, os, openpyxl, pyfiglet, msvcrt
from pyfiglet import Figlet
import keyboard, pyperclip as pc
import data

os.system('mode con: cols=200 lines=100')
keyboard.press('f11')
try:
    f = Figlet('slant')
    print("\n\tWelcome to")
    print(f.renderText("SF6 PRO"))
    print(f"\t\t\t\t\tby Yahia ElGamily\n")
except:
    pass
# Check if database exists
if not os.path.isfile("Links_DB.db"):
    print("No Database found\nChoose SF6 File to Create Database: \n")
    data.update_db()

msg = """
--------------------------------------------------------------------
Enter A Choice:
[1] Search links by site ID
[2] Get Affected Sites
[3] Update Database
[4] Set Preferences
[0] Quit Application
"""
while True:
    print(msg)
    option = msvcrt.getwch()
    if option == '1':
        site = input("Enter Site ID: ").strip().title()
        if os.path.isfile("Links_DB.db"):
            data.search_links(site)
        else:
            print("Database Doesn't Exist")

    elif option == '2':
        site = input("Enter Site ID: ").strip().title()
        sites = set(data.affected_sites(site))
        print(f"\nSite {site} has {len(sites)} affected sites:-\n")
        print(", ".join(sites))
        pc.copy(", ".join(sites))
        pa_sites = open('PA_Sites.csv','w')
        pa_sites.writelines(f"{id}, Yes\n" for id in sites)
        pa_sites.close()
        print("\nList copied to clipboard and File PA_Sites.csv Created successfully")
    elif option == '3':
        data.update_db()
    elif option == '4':
        pass
    elif option == '0':
        break
    else:
        print("Wrong Choice")
