# Auto Data Merge

Auto Data Merge (Multisheet Analyzer) is a Python script designed to streamline and automate the process of collecting, consolidating, and reporting student attendance data for weekly events at UTFPR by aggregating the data, and generating a consolidated Excel file containing student information and attendance statistics. The primary goal of this project is to eliminate the manual effort required to track student attendance, compile the data, and generate comprehensive reports for the UTFPR administration.

## Table of Contents

- [Introduction](#introduction)
- [Problem Statement](#problem-statement)
- [Solution](#solution)
- [Requirements](#requirements)
- [Installation](#installation)
- [Usage](#usage)
- [Author](#author)

## Introduction

This script streamlines the manual process of collecting student attendance information by providing an automated solution. It scans through subdirectories of a specified main directory, extracts student data from Excel files, and generates a consolidated Excel file with detailed student information and attendance counts.

## Problem Statement

Traditionally, tracking student attendance for weekly events at UTFPR required manual data collection, where DACOMP members would note down student names, RA (student ID) numbers, emails, and courses every Friday. This process was not only time-consuming but also prone to errors and inconsistencies.

Additionally, manually compiling this attendance data from various sheets and creating detailed reports consumed significant time and effort. The DACOMP team then needed to combine RA numbers, student names, email addresses, courses, and the number of attendance instances, producing a comprehensive report for the UTFPR administration. This manual process often led to delays in providing accurate attendance statistics and consumed valuable resources.

## Solution

The Auto Data Merge addresses these challenges by automating the entire process, from student data collection to generating consolidated attendance reports. It leverages Python and the pandas library to perform the following tasks:

1. **Data Collection**: The script traverses through subdirectories within the specified main directory, targeting weekly event data folders. It extracts student information, including names, RA numbers (student IDs), emails, and courses, from Excel files.

2. **Data Aggregation**: The collected student data is aggregated based on their RA numbers. The script combines student names, email addresses, courses, and calculates the number of attendance instances for each student.

3. **Consolidated Report**: The script compiles the aggregated data into a consolidated Excel file. This file provides a comprehensive overview of student attendance, including RA numbers, names, emails, courses, and the number of times they attended the weekly events.

## Requirements

- Python 3.6+
- pandas library (`pip install pandas`)
- openpyxl library (`pip install openpyxl`)

## Installation

1. Clone or download this repository.
2. Make sure you have Python 3.6 or higher installed.
3. Install the required libraries by running the following commands:

```bash
pip install pandas openpyxl
```

## Usage

1. Run the script by executing `python AutoDataMerge.py`.
2. Provide the main directory path containing the weekly event data folders.
3. The script will traverse through the subdirectories, extract student information and attendance data from Excel files, and generate a consolidated Excel file with student details and attendance counts, saving it in the 'TotalHours' folder.

## Note

In the context of UTFPR, RA (Registro Acadêmico) refers to the unique identification number assigned to each student. This project specifically addresses the needs of UTFPR and its weekly event attendance tracking process.

## Author

Matheus Mondaini Alegre de Miranda