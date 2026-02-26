# Event Kit Manager

A desktop application built in Python for managing kit distribution in sports events.

## Overview

Event Kit Manager is an offline system designed to control participant kit delivery using Excel spreadsheets as the data source.

The system focuses on:

- Fast search and delivery confirmation
- Real-time statistics
- Safe Excel cell updates
- Automatic backups
- Data integrity during live events

## Features

- Import Excel spreadsheets
- Search participants by name
- Confirm delivery with timestamp
- Real-time delivery statistics
- Automatic backup creation
- Safe row-level updates using EXCEL_ROW mapping

## Technologies

- Python 3.10+
- Tkinter (GUI)
- Pandas (Data handling)
- Openpyxl (Excel manipulation)

## Installation

Install dependencies:

```bash
pip install pandas openpyxl
