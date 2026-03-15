# Amazon WBR Automation

Python automation tool for preparing weekly business reporting outputs from multiple data sources.

## Overview
This project automates a recurring reporting workflow that was previously prepared manually from several files.  
The script collects data, applies the required transformations, and generates output ready to be used in management reporting.

## Business problem
Weekly reporting required manual work across multiple Excel sources.  
The process was repetitive, time-consuming, and dependent on manual calculations such as averages and aggregations.

## Solution
I built a lightweight ETL-style workflow in Python that:
- reads data from multiple source files,
- transforms and aggregates the data,
- prepares final output in a format ready for reporting.

## Tech stack
- Python
- pandas
- Excel-based input/output

## Workflow
1. Load source files
2. Clean and filter the data
3. Apply required calculations (for example averages or sums)
4. Generate output for weekly reporting

## Result
The tool reduced manual work in weekly report preparation and made the process faster and more repeatable.

## Notes
This repository contains an anonymized version of the project.  
Any business-sensitive names, files, and internal identifiers were removed or generalized.
