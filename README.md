# Amazon FLEX Data Tool

Python data-processing tool for preparing structured product attribute datasets from raw source exports.

## Overview

This project streamlines the preparation of product attribute correction files from raw data exports.

The tool cleans input files, filters relevant attributes, and generates structured worksheets used for attribute improvement workflows.

---

## Business Problem

Raw product export files contained a large number of columns and attributes that were not relevant for daily work.

Preparing task-specific datasets and final upload files manually required significant time and introduced the risk of formatting errors.

---

## Solution

A Python-based data processing tool that prepares structured datasets and automates the creation of upload-ready files.

The script simplifies the workflow by keeping only relevant attributes and generating task-specific worksheets.

---

## Tech Stack

- Python
- pandas
- openpyxl
- Excel-based workflows

---

## Usage

1. Export raw product data from the source system.
2. Place the raw file in the input directory.
3. Run the script.
4. The script generates processed worksheets used for attribute correction tasks.
5. After completing updates, run the generation script to produce upload-ready files.

---

## Processing Logic

Internally the tool performs the following steps:

1. Load the raw product dataset.
2. Remove unnecessary columns and attributes.
3. Keep required and supporting fields used during attribute correction.
4. Create task-specific worksheets.
5. Generate structured output files ready for system upload.

---

## Result

The tool reduced the complexity of working with large product datasets and helped standardize the workflow used by the team.

---

## Notes

This repository contains a generalized version of the internal tool.  
Business-specific identifiers and confidential data have been removed.
