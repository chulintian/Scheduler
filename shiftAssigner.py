import heapq as hq
import calendar

def shiftAssigner(people, year, month):
    mainCalendar = calendar.monthcalendar(year, month)
    shiftList = []

    # 3 shifts per day, 5 days a week
    # Assign shifts to people in the heap
    for week in range(0, len(mainCalendar)): 
        for day in range(0,5):
            if mainCalendar[week][day] != 0: # If the day is in the month
                for shifts in range(0,3):
                    for i in [0,1]:
                        shift = Shift(mainCalendar[week][day], day, shifts, "NIL")
                        notAssigned = True
                        putBack = []
                        while (notAssigned):
                            if (len(people) == 0):
                                break
                            person = hq.heappop(people)
                            shift.name = person.name
                            putBack.append(person)
                            if (person.addShift(shift, week, month)):
                                notAssigned = False
                                shiftList.append(shift)

                        for person in putBack:
                            hq.heappush(people, person)
    return shiftList

class Person:
    def __init__(self, name, shifts, unavailable):
        self.shifts = []
        self.desiredShifts = shifts
        self.daysNotAvailable = unavailable
        self.jobs = 0
        self.week = [0,0,0,0,0,0]
        self.name = name

    def numJobs(self):
        return self.jobs

    def addShift(self, shift, weekNum, month):
        additionalShiftMonths = [5, 6, 7, 12]
        if (shift.getDayAndShift() in self.desiredShifts and shift.getDate() not in self.daysNotAvailable 
            and len([x for x in self.shifts if x.equal(shift)]) == 0 and (month in additionalShiftMonths and self.week[weekNum] < 12) or (self.week[weekNum] < 5)):
            self.shifts.append(shift)
            self.jobs += 1
            self.week[weekNum] += 1
            return True
        else:
            return False

    def removeShift(self, i):
        self.shifts.remove(i)
        self.jobs -= 1
    

    def __lt__(self, nxt):
        return self.jobs < nxt.jobs
    
class Shift:
    def __init__(self, date, day, shiftNum, name):
        self.date = date
        self.day = day
        self.shiftNum = shiftNum
        self.name = name

    def getDayAndShift(self):
        return (self.day, self.shiftNum)
    
    def getDate(self):
        return self.date

    def equal(self, shift):
        if (shift.date == self.date and shift.shiftNum == self.shiftNum):
            return True
        else:
            return False
    