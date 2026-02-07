from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.worksheet.table import Table, TableStyleInfo
import os

def setup_console():
    os.system("title Darbo ataskaitos įrankis")
    os.system("color 8F")

def setup():
    year = int(input(" Įveskite datą:\n   Metus:"))
    month = input("   Mėnesį (žodžiais): ")
    name = input("   Serviso inžinieriaus v. pavardė (pvz.: Antano Antanausko): ")
    return year, month, name

def create_excel(year, month, name, records):
    wb = Workbook()
    ws = wb.active
    ws.title = "Ataskaita"

    ws["A2"] = "Data:"
    ws["B2"] = f"{year}, {month}"

    ws["A2"].font = Font(bold=True)

    ws.merge_cells("A4:E4")
    ws["A4"] = f"Serviso inžinieriaus {name} neaktuotų valandų ataskaita"
    ws["A4"].font = Font(size=14, bold=True)
    ws["A4"].alignment = Alignment(horizontal="center")

    headers = [
        "Diena",
        "Objektas",
        "Pradirbtos valandos",
        "Kelionės laikas",
        "Darbo tipas (S - šefmontažas, G - garantinis remontas, SP - serviso projektas, V - vizitas, 0 - nei vienas)"
    ]

    ws.append([])
    ws.append(headers)

    for row in records:
        ws.append([
            row["Diena"],
            row["Objektas"],
            row["Pradirbtos valandos"],
            row["Kelionės laikas"],
            row["Darbo tipas"]
        ])

    start_row = 6
    end_row = start_row + len(records)
    table = Table(displayName="Ataskaita", ref=f"A6:E{end_row}")

    style = TableStyleInfo(
        name="TableStyleMedium9",
        showRowStripes=True
    )
    table.tableStyleInfo = style
    ws.add_table(table)

    ws.column_dimensions["A"].width = 8
    ws.column_dimensions["B"].width = 30
    ws.column_dimensions["C"].width = 20
    ws.column_dimensions["D"].width = 15
    ws.column_dimensions["E"].width = 10

    filename = f"Neaktuotos_valandos_{year}_{month}.xlsx"
    wb.save(filename)

    print(f"Sukurta Excel ataskaita: {filename}")

def data_list():
    all_data = []
    while True:
        data = {}
        data["Diena"] = (input("Diena: ")).strip()
        data["Objektas"] = input("Objekto pavadinimas: ").strip()
        data["Pradirbtos valandos"] = (input("Pradirbtos valandos: ")).strip()
        data["Kelionės laikas"] = (input("Kelionės laikas: ")).strip()
        while True:
            work_type = input("S - šefmontažas, G - garantinis remontas, SP - serviso projektas, V - vizitas, 0 - nei vienas: ").strip().upper()
            if work_type in {"S", "G", "SP", "V", "0"}:
                data["Darbo tipas"] = work_type
                break
            print("Netinkamas simbolis..(S, G, SP, V, 0)")
        all_data.append(data)
        more = input("Sekantis įrašas? (y/n): ").strip().lower()
        if more == "n":
            break

    return all_data
setup_console()
year, month, name = setup()
records = data_list()
create_excel(year, month, name, records)

input("Spausti ENTER, kad uždaryti...")