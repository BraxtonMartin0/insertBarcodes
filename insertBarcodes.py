import sys
from datetime import timedelta

from openpyxl import load_workbook


def insertBarcodes(file, gridsFile):
    wb = load_workbook(file)
    data = wb['Sheet1']

    allClasses = []

    count = data.max_row
    x = 2
    while x < count - 1:
        info_cell = "A" + str(x)
        allInfo = data[info_cell].value
        info = allInfo.split("-")

        activityNumber = info[0]
        if len(info) > 2:
            swimClassName = "PVL"
            level = "PVL"
        else:
            swimClassNameInfo = info[1].split("–")
            swimClassName = swimClassNameInfo[0]
            level = swimClassNameInfo[1]


        if "|" in level:
            temp = level.split("|")
            level = temp[1]

        if "Private Lesson" in level:
            level = "PVL"

        if "low ratio" in level:
            lowRatio = True
            levelTemp = level.split("(")
            level = levelTemp[0]
        else:
            lowRatio = False

        print(level)
        day_cell = "D" + str(x)
        day = data[day_cell].value

        time_cell = "E" + str(x)
        timeTemp = data[time_cell].value.split(":")

        if "PM" in timeTemp[1]:
            PM = True
        else:
            PM = False

        hour = timeTemp[0]
        minute = timeTemp[1][:2]

        # print(str(x - 1) + ":" + info[0] + " : " + swimClassName + " : " + level + " : " + day + " : " + str(
        # hour) + ":" + str(minute) + " PM: " + str(PM) + ": Low Ratio: " + str(lowRatio))

        newClass = SwimClass(activityNumber, swimClassName, level, day, hour, minute, PM, lowRatio)

        allClasses.append(newClass)

        x += 1

    print(len(allClasses))

    wb.save(file)

    for swimmingClass in allClasses:
        gridsWb = load_workbook(gridsFile)
        if swimmingClass.PM and not int(swimmingClass.hour) == 12:
            classTime = str(int(swimmingClass.hour) + 12) + ":" + str(swimmingClass.minute) + ":" + "00"
        else:
            classTime = str(swimmingClass.hour) + ":" + str(swimmingClass.minute)
        if swimmingClass.day == "M":
            if int(swimmingClass.hour) > 3 and swimmingClass.PM and not (int(swimmingClass.hour) == 12):
                grid = gridsWb["Monday PM"]
                daytime = False
                gridSearch(grid, classTime, swimmingClass.level, swimmingClass.activityNumber, swimmingClass.lowRatio,
                           daytime)

            else:
                grid = gridsWb["Daytime"]
                daytime = True
        else:
            print("not monday")
        gridsWb.save(gridsFile)


def gridSearch(sheet, time, level, activityNumber, lowRatio, daytime):
    letters = ["C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W",
               "X", "Y", "Z"]

    timeFound = False
    classNameFound = False
    x = 0
    classNameRow = 7
    activityNumberRow = 8
    end = False
    print("Looking for " + level + " - " + activityNumber + " " + time + " " + str(lowRatio))
    while not timeFound or end:
        gridTimeCell = sheet[letters[x] + "6"]
        if time == str(gridTimeCell.value):
            print("found time")
            while not classNameFound:
                classNameValue = sheet[letters[x] + str(classNameRow)]
                if classNameValue.value is not None:
                    if not "Lifeguarding" in str(classNameValue.value):
                        activityNumberValue = sheet[letters[x] + str(activityNumberRow)]
                        if level.replace(" ","")[:5] in str(classNameValue.value) and activityNumberValue.value is None:
                            sheet[letters[x] + str(activityNumberRow)].value = activityNumber
                            classNameFound = True
                            print("found class, entering barcode")
                            print(level + " - " + activityNumber)

                classNameRow += 2
                activityNumberRow += 2
                if classNameRow > 200 or activityNumberRow > 200:
                    print("not found class")
                    classNameFound = True
            timeFound = True

        if letters[x] == "Z" and not timeFound:
            print("Not Found time")
            end = True
        x += 1

    # print(activityNumber + " - " + level)


class SwimClass:
    def __init__(self, activityNumber, swimClassName, level, day, hour, minute, PM, lowRatio):
        self.activityNumber = activityNumber
        self.swimClassName = swimClassName
        self.level = level
        self.day = day
        self.hour = hour
        self.minute = minute
        self.lowRatio = lowRatio
        self.PM = PM


if __name__ == '__main__':
    # file = sys.argv[1]
    insertBarcodes('active_report.xlsx', '2024 WBSC Grids - Winter.xlsx')
