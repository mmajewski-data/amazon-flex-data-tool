# Amazon FLEX Data Tool

Python data-processing tool for cleaning large product files and generating task-specific workbooks and upload-ready outputs.

## Overview
This project streamlines the workflow of preparing product attribute correction files from raw source exports.

## Business problem
Raw source files contained many unnecessary columns and were difficult to use efficiently in daily operations.  
Preparing task-specific files and final upload files manually was time-consuming and error-prone.

## Solution
I built a Python-based workflow tool that:
- cleans raw product exports,
- keeps only the relevant attributes,
- creates task-specific worksheets,
- supports a simple READY-based workflow,
- generates upload-ready output files.

## Tech stack
- Python
- pandas
- openpyxl
- Excel-based workflow

## Workflow
1. Load raw source file
2. Mark completed records with READY = yes
3. Generate final files ready for upload

## Result
The tool improved structure, reduced manual file preparation, and made the overall workflow faster and easier to manage.

## Notes
This repository is a generalized version of the original internal tool.  
Sensitive business-specific details were anonymized.
