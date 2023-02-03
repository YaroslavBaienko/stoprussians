import xlsxwriter
import requests
from pathlib import Path
from datetime import date, timedelta
import string

start = date(2022, 2, 24)
end = date.today()
current_dir = Path.cwd()
headers = {
    'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_10_1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/39.0.2171.95 Safari/537.36'}

workbook = xlsxwriter.Workbook('russians_dies.xlsx')
worksheet = workbook.add_worksheet()

upper_letters = string.ascii_uppercase

field_names = ["date", "units", "tanks", "armoured_fighting_vehicles", "artillery_systems", "mlrs",
               "aa_warfare_systems",
               "planes", "helicopters", "vehicles_fuel_tanks", "warships_cutters", "cruise_missiles",
               "uav_systems", "special_military_equip", "atgm_srbm_systems"]

for col, name in zip(upper_letters, field_names):
    worksheet.write(col + "1", name)

days = 1
datas = start
column = 1
while datas < end:
    datas = start + timedelta(days=days)
    days += 1
    print(datas)
    url = f"https://russianwarship.rip/api/v1/statistics/{datas}"

    try:
        response = requests.get(url=url, headers=headers)
        response.raise_for_status()
        result = response.json()
        # string_json = json.dumps(result)
        # with open(f"{current_dir}/data/war.json", 'w') as file:
        #     file.write(string_json)
        #     file.close()

    except requests.exceptions.HTTPError as error:
        print(f"HTTP error occurred: {error}")
    except requests.exceptions.RequestException as error:
        print(f"Request error occurred: {error}")

    stats = result["data"]["stats"]
    units = stats['personnel_units']
    tanks = stats["tanks"]
    armoured_fighting_vehicles = stats["armoured_fighting_vehicles"]
    artillery_systems = stats["artillery_systems"]
    mlrs = stats["mlrs"]
    aa_warfare_systems = stats["aa_warfare_systems"]
    planes = stats["planes"]
    helicopters = stats["helicopters"]
    vehicles_fuel_tanks = stats["vehicles_fuel_tanks"]
    warships_cutters = stats["warships_cutters"]
    cruise_missiles = stats["cruise_missiles"]
    uav_systems = stats["uav_systems"]
    special_military_equip = stats["special_military_equip"]
    atgm_srbm_systems = stats["atgm_srbm_systems"]

    content = [datas, units, tanks, armoured_fighting_vehicles, artillery_systems, mlrs, aa_warfare_systems, planes,
               helicopters,
               vehicles_fuel_tanks, warships_cutters, cruise_missiles, uav_systems, special_military_equip,
               atgm_srbm_systems]
    column += 1

    for row, item in zip(upper_letters, content):
        # write operation perform
        worksheet.write((row + str(column)), item)

        # incrementing the value of row by one
        # with each iterations.

workbook.close()
