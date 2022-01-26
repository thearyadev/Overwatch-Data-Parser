import gspread
from utils import Console
import datetime
import time
import numpy as np
import psutil
from utils import MAP_LIST
from gspread.exceptions import APIError


class SheetManager:
    def __init__(self, credentials, sheet_name):
        self.spreadsheet = self.authorize(credentials, sheet_name)
        self.data_sheet = self.spreadsheet.worksheet("Logs")
        self.stat_sheet = self.spreadsheet.worksheet("Generated Statistics")
        self.data: list = list()
        self.console = Console()

    def loop(self, repeat_time=120):
        self.console.info_log("Script is active.")
        while True:
            try:
                last_update = datetime.datetime.now()
                self.refresh_backend(last_refresh=last_update, time_till="updating")
                self.reload_data()
                self.write()
                for second in range(repeat_time):
                    if second % 10 == 0:
                        self.refresh_backend(last_refresh=last_update, time_till=repeat_time - second)
                    time.sleep(1)
            except APIError:
                time.sleep(60)
                self.stat_sheet.update("G23:H23", "RATE LIMITED")
                time.sleep(60)

    @staticmethod
    def authorize(credentials, sheet_name):
        return gspread.service_account(credentials).open(sheet_name)

    def reload_data(self):
        self.data = [[cell for cell in row if cell] for row in self.data_sheet.get_all_values()[1:]]
        for row in self.data:
            row[0] = datetime.datetime.strptime(row[0], "%m/%d/%Y")
        self.data = [row for row in self.data if len(row) > 3]

    def write(self):
        all_data = self.count(datetime.timedelta(days=999999))
        four_weeks_data = self.count(datetime.timedelta(days=28))
        seven_days_data = self.count(datetime.timedelta(days=7))
        maps = self.map_count()
        for key, val in all_data.items():
            if key == "Sunday":
                self.stat_sheet.update("C4:C5", [[v] for v in val])
            elif key == "Monday":
                self.stat_sheet.update("D4:D5", [[v] for v in val])
            elif key == "Tuesday":
                self.stat_sheet.update("E4:E5", [[v] for v in val])
            elif key == "Wednesday":
                self.stat_sheet.update("F4:F5", [[v] for v in val])
            elif key == "Thursday":
                self.stat_sheet.update("G4:G5", [[v] for v in val])
            elif key == "Friday":
                self.stat_sheet.update("H4:H5", [[v] for v in val])
            elif key == "Saturday":
                self.stat_sheet.update("I4:I5", [[v] for v in val])
        time.sleep(1)
        for key, val in four_weeks_data.items():
            if key == "Sunday":
                self.stat_sheet.update("C10:C11", [[v] for v in val])
            elif key == "Monday":
                self.stat_sheet.update("D10:D11", [[v] for v in val])
            elif key == "Tuesday":
                self.stat_sheet.update("E10:E11", [[v] for v in val])
            elif key == "Wednesday":
                self.stat_sheet.update("F10:F11", [[v] for v in val])
            elif key == "Thursday":
                self.stat_sheet.update("G10:G11", [[v] for v in val])
            elif key == "Friday":
                self.stat_sheet.update("H10:H11", [[v] for v in val])
            elif key == "Saturday":
                self.stat_sheet.update("I10:I11", [[v] for v in val])
        time.sleep(1)
        for key, val in seven_days_data.items():
            if key == "Sunday":
                self.stat_sheet.update("C16:C17", [[v] for v in val])
            elif key == "Monday":
                self.stat_sheet.update("D16:D17", [[v] for v in val])
            elif key == "Tuesday":
                self.stat_sheet.update("E16:E17", [[v] for v in val])
            elif key == "Wednesday":
                self.stat_sheet.update("F16:F17", [[v] for v in val])
            elif key == "Thursday":
                self.stat_sheet.update("G16:G17", [[v] for v in val])
            elif key == "Friday":
                self.stat_sheet.update("H16:H17", [[v] for v in val])
            elif key == "Saturday":
                self.stat_sheet.update("I16:I17", [[v] for v in val])
        time.sleep(1)
        for key, val in maps['Assault'].items():
            if key == "Anubis":
                self.stat_sheet.update("L4:N4", [val])
            elif key == "Hanamura":
                self.stat_sheet.update("L5:N5", [val])
            elif key == "Volskaya":
                self.stat_sheet.update("L6:N6", [val])
        time.sleep(1)
        for key, val in maps["KOTH"].items():
            if key == "Busan":
                self.stat_sheet.update("L9:N9", [val])
            elif key == "Ilios":
                self.stat_sheet.update("L10:N10", [val])
            elif key == "Lijiang":
                self.stat_sheet.update("L11:N11", [val])
            elif key == "Nepal":
                self.stat_sheet.update("L12:N12", [val])
            elif key == "Oasis":
                self.stat_sheet.update("L13:N13", [val])
        time.sleep(1)
        for key, val in maps['Hybrid'].items():
            if key == "Blizzard World":
                self.stat_sheet.update("L16:N16", [val])
            elif key == "Eichenwalde":
                self.stat_sheet.update("L17:N17", [val])
            elif key == "Hollywood":
                self.stat_sheet.update("L18:N18", [val])
            elif key == "Kings Row":
                self.stat_sheet.update("L19:N19", [val])
            elif key == "Numbani":
                self.stat_sheet.update("L20:N20", [val])
        time.sleep(1)
        for key, val in maps["Escort"].items():
            if key == "Dorado":
                self.stat_sheet.update("L23:N23", [val])
            elif key == "Gibraltar":
                self.stat_sheet.update("L24:N24", [val])
            elif key == "Havana":
                self.stat_sheet.update("L25:N25", [val])
            elif key == "Junkertown":
                self.stat_sheet.update("L26:N26", [val])
            elif key == "Rialto":
                self.stat_sheet.update("L27:N27", [val])
            elif key == "Route 66":
                self.stat_sheet.update("L28:N28", [val])
        time.sleep(1)

        self.stat_sheet.update("R4:T22", self.all_maps_ordered())
        self.stat_sheet.update("C22:C29", [[v] for v in self.misc().values()])

    def count(self, query: datetime.timedelta):
        count = {
            "Sunday": [0, 0],
            "Monday": [0, 0],
            "Tuesday": [0, 0],
            "Wednesday": [0, 0],
            "Thursday": [0, 0],
            "Friday": [0, 0],
            "Saturday": [0, 0],
        }
        for entry in self.data:
            if (datetime.datetime.now() - entry[0]) < query:
                if entry[0].weekday() == 0:
                    count['Monday'][1] += 1
                    if entry[-1] == "W":
                        count['Monday'][0] += 1
                elif entry[0].weekday() == 1:
                    count['Tuesday'][1] += 1
                    if entry[-1] == "W":
                        count['Tuesday'][0] += 1
                elif entry[0].weekday() == 2:
                    count['Wednesday'][1] += 1
                    if entry[-1] == "W":
                        count['Wednesday'][0] += 1
                elif entry[0].weekday() == 3:
                    count['Thursday'][1] += 1
                    if entry[-1] == "W":
                        count['Thursday'][0] += 1
                elif entry[0].weekday() == 4:
                    count['Friday'][1] += 1
                    if entry[-1] == "W":
                        count['Friday'][0] += 1
                elif entry[0].weekday() == 5:
                    count['Saturday'][1] += 1
                    if entry[-1] == "W":
                        count['Saturday'][0] += 1
                elif entry[0].weekday() == 6:
                    count['Sunday'][1] += 1
                    if entry[-1] == "W":
                        count['Sunday'][0] += 1
        return count

    def map_count(self):
        types = {
            "Assault": dict({"Anubis": list([0, 0, 0]), "Hanamura": list([0, 0, 0]), "Volskaya": list([0, 0, 0])}),
            "KOTH": dict(
                {"Busan": list([0, 0, 0]), "Ilios": list([0, 0, 0]), "Lijiang": list([0, 0, 0]),
                 "Nepal": list([0, 0, 0]), "Oasis": list([0, 0, 0])}),
            "Hybrid": dict(
                {"Blizzard World": list([0, 0, 0]), "Eichenwalde": list([0, 0, 0]), "Hollywood": list([0, 0, 0]),
                 "Kings Row": list([0, 0, 0]),
                 "Numbani": list([0, 0, 0])}),
            "Escort": dict({"Dorado": list([0, 0, 0]), "Gibraltar": list([0, 0, 0]), "Havana": list([0, 0, 0]),
                            "Junkertown": list([0, 0, 0]),
                            "Rialto": list([0, 0, 0]), "Route 66": list([0, 0, 0])})
        }
        for entry in self.data:
            # [win, loss, played]
            try:
                types[entry[3]][entry[2]][2] += 1  # inc played
                if entry[4] == "W":
                    types[entry[3]][entry[2]][0] += 1  # inc wins
                elif entry[4] == "L":
                    types[entry[3]][entry[2]][1] += 1  # inc losses
            except KeyError:
                pass
        return types

    def all_maps_ordered(self):
        parsed = dict()

        for entry in self.data:
            entry_map = entry[2]
            if entry_map not in parsed.keys():
                parsed[entry_map] = [0, 0]  # wins played

        for entry in self.data:
            parsed[entry[2]][1] += 1
            if entry[4] == "W":
                parsed[entry[2]][0] += 1
        zipped = zip(
            parsed.keys(),
            [(ls[0] / ls[1]) for ls in parsed.values()],
            [(oc[1] / len(self.data)) for oc in parsed.values()]
        )
        zipped = list(zipped)
        for M in MAP_LIST:
            if M not in parsed.keys():
                zipped.append(
                    (M, 0, 0,)
                )

        return sorted(zipped, key=lambda x: x[1], reverse=True)

    def misc(self):
        datas = dict({
            "delta_positive": float,
            "delta_negative": float,
            "total": int,
            "games_per_day": float,
            "least_played_day": str,
            "most_played_day": str,
            "highest_sr": int,
            "average_sr": float
        })
        datas["delta_positive"] = \
            sum([i for i in np.diff(np.array([int(row[1]) for row in self.data])) if i > 0]) / \
            len([i for i in np.diff(np.array([int(row[1]) for row in self.data])) if i > 0])
        datas["delta_negative"] = \
            sum([i for i in np.diff(np.array([int(row[1]) for row in self.data])) if i < 0]) / \
            len([i for i in np.diff(np.array([int(row[1]) for row in self.data])) if i < 0])

        datas['total'] = len(self.data)

        d = dict()
        for entry in self.data:
            try:
                d[entry[0].strftime("%m/%d/%Y")] += 1
            except KeyError:
                d[entry[0].strftime("%m/%d/%Y")] = 1
        datas['games_per_day'] = sum(d.values()) / len(d.values())
        rng = self.stat_sheet.range("C5:I5")
        days = [cell.value for cell in rng[0:7]]
        gp = [int(cell.value) for cell in rng[-7:]]
        zipped = sorted(list(zip(days, gp)), key=lambda x: x[1])
        datas["least_played_day"] = zipped[0][0]
        datas['most_played_day'] = zipped[-1][0]
        datas['highest_sr'] = max([int(row[1]) for row in self.data])
        datas["average_sr"] = sum([int(row[1]) for row in self.data]) / len([int(row[1]) for row in self.data])

        return datas

    def refresh_backend(self, last_refresh: datetime.datetime, time_till):
        if time_till == "updating":
            datas = {
                "time_till": last_refresh.strftime("%m/%d at %I:%S %p"),
                "last_refresh": "Currently Refreshing...",
                "battery_percentage": "Error",
                "plugged": "Error"
            }
        else:
            datas = {
                "time_till": last_refresh.strftime("%m/%d at %I:%S %p"),
                "last_refresh": "Approximately " + str(time_till) + " seconds",
                "battery_percentage": "Error",
                "plugged": "Error"
            }

        if battery := psutil.sensors_battery():
            datas['battery_percentage'] = battery.percent
            datas['plugged'] = battery.power_plugged

        self.stat_sheet.update("G22:H25", [[v] for v in list(datas.values())])


if __name__ == "__main__":
    manager = SheetManager("creds.json", "Sens: Season 32")
    manager.loop(180)
