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
    
    hoursCol = 65
    
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
                personHoursPerWeek = {}
                for shiftNum in range(0,3):
                    for i in range(0,2):
                        peopleRow += 1
                        currShift = Shift(mainCalendar[week][day], day, shiftNum, "CHECK")
                        for shift in shifts:
                            if (shift.equal(currShift)):
                                formatString = f"{peopleCol}{peopleRow}"
                                worksheet.write(formatString, shift.name, format4)
                                shifts.remove(shift)
                                
                                #update personHoursPerWeek
                                if shift.name in personHoursPerWeek:
                                    personHoursPerWeek[shift.name] += 3
                                else:
                                    personHoursPerWeek[shift.name] = 3
                                    
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
                        worksheet.write(endCol + str(tempRow), f'=HOUR({endTimeCol}{tempRow}-{startTimeCol}{tempRow}) + MINUTE({endTimeCol}{tempRow}-{startTimeCol}{tempRow})/60', format4)

                
                # fill in hours per person per week
                hoursRow = 47
                start_cell = f"{chr(hoursCol)}{hoursRow}"
                end_cell = f"{chr(hoursCol) + 1}{str(hoursRow)}"
                worksheet.merge_range(start_cell + ':' + end_cell, 'Week ' + str(week), format2)

                hoursRow += 1
                worksheet.write(chr(hoursCol) + str(hoursRow), "Name", format3)
                worksheet.write(chr(hoursCol + 1) + str(hoursRow), "Total Hours Per Week", format3)
                hoursRow += 1

                for name, hours in personHoursPerWeek.items():
                    worksheet.write(chr(hoursCol) + str(hoursRow), name, format4)
                    worksheet.write(chr(hoursCol + 1) + str(hoursRow), hours, format4)
                    hoursRow += 1
                    
                hoursCol += 3
                    
    # Close the workbook to save it
    workbook.close()