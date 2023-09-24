import xlsxwriter
import calendar
from shiftAssigner import *

def xlsxWriter(people, month, year, shifts):
    workbook = xlsxwriter.Workbook(calendar.month_name[month] + " " + str(year) + '.xlsx')
    worksheet = workbook.add_worksheet()
    mainCalendar = calendar.monthcalendar(year, month)
    
    # formats
    format1 = workbook.add_format({'bg_color': '#fde4ce', 'align': 'center'})
    format2 = workbook.add_format({'bg_color': '#72a7db', 'align': 'center'})
    format3 = workbook.add_format({'bg_color': '#d9f1f3', 'align': 'center'})
    format4 = workbook.add_format({'align': 'center'})
    
    # month title
    worksheet.merge_range('A1:T1', calendar.month_name[month] + " " + str(year), format1)

    # weekday title
    worksheet.merge_range('A2:D2', 'Monday', format2)
    worksheet.merge_range('E2:H2', 'Tuesday', format2)
    worksheet.merge_range('I2:L2', 'Wednesday', format2)
    worksheet.merge_range('M2:P2', 'Thursday', format2)
    worksheet.merge_range('Q2:T2', 'Friday', format2)
    
    mainCalendar[0]
    # date title
    for week in range(0,len(mainCalendar)):
        for day in range(0,5):
            if mainCalendar[week][day] != 0: # If the day is in the month
                startCol = chr(65 + day*4)
                endCol = chr(68 + day*4)
                row = 3 + 7 * week
                formatString = f"{startCol}{row}:{endCol}{row}"
                worksheet.merge_range(formatString, mainCalendar[week][day], format3)

                
                # fill people
                peopleCol = chr(65 + day * 4)
                peopleRow = row
                for shiftNum in range(0,3):
                    for i in range(0,2):
                        peopleRow += 1
                        currShift = Shift(mainCalendar[week][day], day, shiftNum, "CHECK")
                        for shift in shifts:
                            if (shift.equal(currShift)):
                                formatString = f"{peopleCol}{peopleRow}"
                                worksheet.write(formatString, shift.name, format4)
                                shifts.remove(shift)                  
                                break
                                

                # fill in start time, end time, num of hours
                tempRow = row
                for j in range (0, 3):
                    startTimeCol = chr(66 + day*4)
                    endTimeCol = chr(67 + day*4)
                    startTime = str(9 + j * 3).zfill(2) + ":30"
                    endTime = str(12 + j * 3).zfill(2) + ":30"
                    for k in range(0, 2):
                        tempRow += 1
                        worksheet.write(startTimeCol + str(tempRow), startTime, format4)
                        worksheet.write(endTimeCol + str(tempRow), endTime, format4)
                        worksheet.write(endCol + str(tempRow), f'={endTimeCol}{tempRow}-{startTimeCol}{tempRow}', format4)

                
        # fill in hours per person per week
        hoursCol = 65 + week * 2
        hoursRow = 47
        formatStringHours = f"{chr(hoursCol)}{hoursRow}:{chr(hoursCol + 1)}{hoursRow}"
        worksheet.merge_range(formatStringHours, "Week " + str(week + 1), format2)

        hoursRow += 1
        worksheet.write(chr(hoursCol) + str(hoursRow), "Name", format3)
        worksheet.write(chr(hoursCol + 1) + str(hoursRow), "Total Hours Per Week", format3)

        for person in people:
            hoursRow += 1
            worksheet.write(chr(hoursCol) + str(hoursRow), person.name, format4)
            worksheet.write(chr(hoursCol + 1) + str(hoursRow), person.week[week] * 3 , format4)
                    
                    
    # Close the workbook to save it
    workbook.close()