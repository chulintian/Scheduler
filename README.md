# Automatic Scheduler

## Overview

This project is an Automatic Scheduler that takes input data from an Excel file retrieved from a Microsoft Form, assigns people to slots based on their availability, and generates an output Excel file with the scheduled assignments.

## Features

- Collects input data from a Microsoft Form.
- Retrieves data from the Excel file generated by the form.
- Processes and cleans the data for scheduling.
- Implements a scheduling algorithm to assign people to available slots.
- Generates an output Excel file with scheduled assignments.

## Prerequisites

Before running the Automatic Scheduler, ensure you have the following prerequisites:

- [Python](https://www.python.org/downloads/)

## Setup

1. Clone this repository to your local machine:
2. Install the necessary dependencies:

- For Python, use pip:

  ```
  pip install -r requirements.txt
  ```

3. Set up the Microsoft Form:

- Create a Microsoft Form to collect input data.
- Export responses to an Excel file.

## Usage

1. Open a command prompt or terminal on your computer.
2. Navigate to the directory where the `app.py` script is located using the `cd` command.
3. Run the script
   - *Replace "C:\path\to\your\file.xlsx" with the actual path to the Excel file you want to process.* 
   - *Replace "month/year" with the actual month/year (e.g. 9/2023) you want to process.* 
  ```
  python app.py C:\path\to\your\file.xlsx month/year
  ```

