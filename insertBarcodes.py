import sys
from datetime import timedelta

from openpyxl import load_workbook

def insertBarcodes(file, gridsFile):
    wb = load_workbook(file)
    data = wb['Sheet1']

    allClasses = []

    count = data.max_row
    x=2
    while x < count - 1:
        info_cell = "A" + str(x)
        allInfo = data[info_cell].value
        info = allInfo.split("-")

        activityNumber = info[0]
        swimClassNameInfo = info[1].split("–")
        swimClassName = swimClassNameInfo[0]
        level = swimClassNameInfo[1].replace(" ","")

        if "|" in level:
            temp = level.split("|")
            level = temp[1]

        if "lowratio" in level:
            lowRatio = True
            levelTemp = level.split("(")
            level = levelTemp[0]
        else:
            lowRatio = False

        day_cell = "D" + str(x)
        day = data[day_cell].value

        time_cell = "E" + str(x)
        timeTemp = data[time_cell].value.split(":")

        if "PM" in  timeTemp[1]:
            PM = True
        else:
            PM = False

        hour = timeTemp[0]
        minute = timeTemp[1][:2]

        print(str(x-1) + ":" + info[0] + " : " + swimClassName + " : " + level + " : " + day + " : " + str(hour) + ":" + str(minute) + ": Low Ratio: " + str(lowRatio) )

        newClass = SwimClass(activityNumber,swimClassName, level, day, hour, minute, PM, lowRatio)

        allClasses.append(newClass)

        x += 1

    print(len(allClasses))

    wb.save(file)

    gridsWb = load_workbook(gridsFile)

    for swimmingClass in allClasses:
        if swimmingClass.day == "M":
            if swimmingClass.hour > 3 and swimmingClass.PM:
                grid = gridsWb["Monday PM"]


class SwimClass:
    def __init__(self,activityNumber, swimClassName, level, day, hour, minute, PM, lowRatio):
        self.activityNumber = activityNumber
        self.swimClassName = swimClassName
        self.level = level
        self.day = day
        self.hour = hour
        self.minute = minute
        self.lowRatio = lowRatio
        self.PM = PM

if __name__ == '__main__':
    #file = sys.argv[1]
    insertBarcodes('active_report.xlsx','2024 WBSC Grids - Winter.xlsx')