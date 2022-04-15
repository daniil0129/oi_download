# from rich.progress import track
# import win32com.client as client
# from rich.console import Console
# import os
#
# outlook = client.Dispatch("Outlook.Application")
# name_space = outlook.GetNameSpace("MAPI")
# console = Console(width=120)
# style = "bold grey0 on white"
# style1 = "blink turquoise2 on black"
# style2 = " green on black"
#
# lis = []
# for i in name_space.Folders:
#     lis.append(i.FolderPath)
#     for j in i.Folders:
#         lis.append(j.FolderPath)
#
# def search_mailitems(scope, sql_query):
#     result = []
#
#     for name in name_space.Folders:
#         for fol_i in name.Folders:
#             if (fol_i.FolderPath == scope):
#                 for mi in fol_i.Items.Restrict("@SQL=" + sql_query):
#                     result.append(mi)
#     return result
#
#
# def download_attach():
#
#     console.print("Привет, это утилита для выгрузки вложений из почты Outlook!", style=style, justify="center")
#     console.input("[blink dodger_blue3](press Enter to continue)[/blink dodger_blue3]")
#     for i in lis:
#         console.print("\t" + i, style=style2)
#     scope = console.input("\nСкопируй ссылку на папку поиска из списка выше: ")
#     sql_query = console.input("\nСкопируй sql запрос из Outlook: ")
#
#     mii = search_mailitems(scope, sql_query)
#     os.mkdir(os.getcwd() + "\\" + "attachments")
#
#
#     for i in track(mii, description="Идет загрузка вложений..."):
#         for j in i.Attachments:
#             j.SaveAsFile(os.getcwd() + "\\" + "attachments" + "\\" + str(j.Parent.CreationTime.strftime("%d_%m_%y %H-%M-%S")) + j.FileName)
#
#
# def main():
#     download_attach()
#
# if __name__ == '__main__':
#     main()

#"urn:schemas:httpmail:hasattachment" = 1
# 02.03.2022 17:38:10# : Date
    # scope = r'\\daniil.s.h.e.p.k.o.v@gmail.com\Входящие'  # "C:\Users\DANIIL\Desktop\qwqwqwqw"
    # sql_query = """("http://schemas.microsoft.com/mapi/proptag/0x0065001f" CI_STARTSWITH 'Щепков' OR "http://schemas.microsoft.com/mapi/proptag/0x0042001f" CI_STARTSWITH 'Щепков')"""

from rich.progress import track
import win32com.client as client
from rich.console import Console
import os

outlook = client.Dispatch("Outlook.Application")
name_space = outlook.GetNameSpace("MAPI")
console = Console(width=120)
style = "bold grey0 on white"
style1 = "blink turquoise2 on black"
style2 = " green on black"

lis = []
for i in name_space.Folders:
    lis.append(i.FolderPath)
    for j in i.Folders:
        lis.append(j.FolderPath)

def search_mailitems(scope, sql_query):
    result = []

    for name in name_space.Folders:
        for fol_i in name.Folders:
            if (fol_i.FolderPath == scope):
                for mi in fol_i.Items.Restrict("@SQL=" + sql_query): #https://docs.microsoft.com/ru-ru/office/vba/api/outlook.items.restrict
                    if mi.Class == 43:
                        result.append(mi)
    return result


def download_attach():

    console.print("Привет, это утилита для выгрузки вложений из почты Outlook!", style=style, justify="center")
    console.input("[blink dodger_blue3](press Enter to continue)[/blink dodger_blue3]")
    for i in lis:
        console.print("\t" + f"[{lis.index(i)}] - " + i, style=style2)
    scope = lis[int(console.input("\nВведи индекс области поиска из списка выше: "))]
    sql_query = console.input("\nСкопируй sql запрос из Outlook: ")

    mii = search_mailitems(scope, sql_query)
    os.mkdir(os.getcwd() + "\\" + "attachments")

    for i in track(mii, description="Идет загрузка вложений..."):
        for j in i.Attachments:         #https://docs.microsoft.com/ru-ru/office/vba/api/outlook.mailitem#properties
            j.SaveAsFile(os.getcwd() + "\\" + "attachments" + "\\" + str(j.Parent.CreationTime.strftime("%d %B %Y (%H'%M'%S) - ")) + j.FileName)

    console.input("[blink dodger_blue3](press Enter to exit)[/blink dodger_blue3]")
def main():
    download_attach()

if __name__ == '__main__':
    main()

#"urn:schemas:httpmail:hasattachment" = 1
# 02.03.2022 17:38:10# : Date
# scope = r'\\daniil.s.h.e.p.k.o.v@gmail.com\Входящие'

