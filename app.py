# import module
from xlsxReader import *
from shiftAssigner import *
from xlsxWriter import *
import heapq as hq
import argparse


def main(xlsxFilename, monthYear):
    # get month and year from param
    monthYear = monthYear.split("/")
    month = int(monthYear[0])
    year = int(monthYear[1])
    
    people = xlsxReader(xlsxFilename)

    shiftList = shiftAssigner(people, year, month)

    xlsxWriter(people, month, year, shiftList)
	
    
if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Scheduling Tool")
    parser.add_argument("xlsxFile", help="Path to the Excel file")
    parser.add_argument("monthYear", help="Month and year (e.g., '9/2023')")

    args = parser.parse_args()
    main(args.xlsxFile, args.monthYear)
    