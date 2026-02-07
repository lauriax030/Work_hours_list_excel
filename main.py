from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.worksheet.table import Table, TableStyleInfo
import os
import calendar
#import objektai

MONTHS = ["Sausis", "Vasaris", "Kovas", "Balandis", "Gegužė", "Birželis", "Liepa", "Rugpjūtis", "Rugsėjis", "Spalis", "Lapkritis", "Gruodis",]

def input_year():
    print(" Įveskite datą:\n")
    while True:
        y = input("   Metus:").strip()
        if y.isdigit() and len(y) == 4:
            return int(y)
        print("Netinkami metai! Įveskite keturženklį skaičių, pvz., 2026.")

def input_month():
    while True:
        m = input("   Mėnesį: ").strip().lower()
        if m.isdigit():
            i = int(m)
            if 1 <= i <= 12:
                return MONTHS[i - 1]
        if len(m) >= 4:
            for month in MONTHS:
                if month.lower().startswith(m[:4]):
                    return month
        print("Netinkamas mėnuo. Pvz: 8, 08, rugpjutis, Rugpjūtis")

def input_day(year, month_name):
    month_num = None
    for i, name in enumerate(MONTHS, start=1):
        if name.lower() == month_name.lower():
            month_num = i
            break
    if month_num is None:
        month_num = 1
    _, max_day = calendar.monthrange(year, month_num)
    while True:
        d = input(f"Diena (1-{max_day}): ").strip()
        if d.isdigit():
            day = int(d)
            if 1 <= day <= max_day:
                return day
        print(f"Netinkama diena! Įveskite skaičių nuo 1 iki {max_day}.")

def input_hours(prompt):
    while True:
        h = input(f"{prompt}: ").strip().replace(',', '.')
        try:
            val = float(h)
            if val >= 0:
                return val
            else:
                print("Negalima įvesti neigiamų valandų!")
        except ValueError:
            print("Netinkamas formatas! Įveskite skaičių, pvz.: 3.5")

def setup_console():
    os.system("title Darbo ataskaitos įrankis")
    os.system("color 8F")

def setup():
    year = input_year()
    month = input_month()
    name = input("   Serviso inžinieriaus v. pavardė (pvz.: Antano Antanausko): ")
    return year, month, name

def summarize_records(records):
    summary = {}
    for row in records:
        key = (row["Objektas"], row["Darbo tipas"])
        if key not in summary:
            summary[key] = {"Pradirbtos valandos": 0, "Kelionės laikas": 0, "Viršvalandžiai": 0}
        summary[key]["Pradirbtos valandos"] += row["Pradirbtos valandos"]
        summary[key]["Kelionės laikas"] += row["Kelionės laikas"]
        summary[key]["Viršvalandžiai"] += row.get("Viršvalandžiai")
    return summary


def append_summary_table(ws, summary):
    ws.append([])
    ws.append([])
    ws.append([])
    ws.append([
        "Darbo tipas", "Objektas", "Pradirbtos valandos",
        "Kelionės laikas", "AX S.C.", "Viršvalandžiai"
    ])
    
    start_row = ws.max_row
    for (obj, work_type), totals in summary.items():
        ws.append([
            work_type,
            obj,
            totals["Pradirbtos valandos"],
            totals["Kelionės laikas"],
            "",
            totals["Viršvalandžiai"]
        ])
    end_row = ws.max_row

    table = Table(displayName="Summary", ref=f"A{start_row}:F{end_row}")
    style = TableStyleInfo(
        name="TableStyleMedium2",
        showRowStripes=True
    )
    table.tableStyleInfo = style
    ws.add_table(table)

    for row in ws.iter_rows(min_row=start_row+1, min_col=3, max_col=4, max_row=end_row):
        for cell in row:
            cell.number_format = '0.0'
    for row in ws.iter_rows(min_row=start_row+1, min_col=6, max_col=6, max_row=end_row):
        for cell in row:
            cell.number_format = '0.0'

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
        "Viršvalandžiai",
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
            row["Viršvalandžiai"],
            row["Darbo tipas"]
        ])

    start_row = 6
    end_row = start_row + len(records)
    table = Table(displayName="Ataskaita", ref=f"A6:F{end_row}")

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
    ws.column_dimensions["E"].width = 15
    ws.column_dimensions["F"].width = 30

    summary = summarize_records(records)
    append_summary_table(ws, summary)


    filename = f"Neaktuotos_valandos_{year}_{month}.xlsx"
    wb.save(filename)

    print(f"Sukurta Excel ataskaita: {filename}")

def data_list(year, month):
    all_data = []
    while True:
        data = {}
        data["Diena"] = input_day(year, month)
        data["Objektas"] = input("Objekto pavadinimas: ").strip()
        data["Pradirbtos valandos"] = input_hours("Pradirbtos valandos")
        data["Kelionės laikas"] = input_hours("Kelionės laikas")
        data["Viršvalandžiai"] = input_hours("Viršvalandžiai")
        
        while True:
            work_type = input(
                "S - šefmontažas, G - garantinis remontas, SP - serviso projektas, V - vizitas, 0 - nei vienas: "
            ).strip().upper()
            if work_type in {"S", "G", "SP", "V", "0"}:
                data["Darbo tipas"] = work_type
                break
            print("Netinkamas simbolis..(S, G, SP, V, 0)")
        
        all_data.append(data)
        more = input("Sekantis įrašas? (y/n): ").strip().lower()
        if more == "n":
            break

    return all_data

def main():
    setup_console()
    year, month, name = setup()
    records = data_list(year, month)
    create_excel(year, month, name, records)
    input("Spausti ENTER, kad uždaryti...")

if __name__ == "__main__":
    main()