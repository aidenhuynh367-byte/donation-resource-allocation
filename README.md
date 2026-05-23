# Donation Resource Allocation

A Python tool that allocates school-supply donations between partner organisations using rule-based logic, weighted prioritisation, and Excel data processing.

## Overview

This project was created for my Tools for School drive, a school-supply donation initiative supporting children in need in Bali. The drive collected notebooks, children’s books, pencils, stationery, toys, and many other learning materials for partner organisations.

Manual allocation became inefficient because donations varied by category, quantity, and usefulness and there were many to each organisation. This program uses Python to make the allocation process more systematic, transparent, and fair.

## Features

- Reads donation data from an Excel spreadsheet
- Calculates total donated items by category
- Applies rule-based allocation logic
- Uses weighted prioritisation for essential school supplies
- Handles organisation-specific needs and constraints
- Generates allocation and total donation reports

## Algorithm Logic

The program uses rule-based logic because different donations had different constraints. Some items could be split between organisations, while others were assigned directly based on partner needs.

Essential school supplies such as notebooks, pencils, and erasers are prioritised using quantity thresholds and weighted logic. Other items, such as children’s books or art supplies, can be allocated based on the needs of specific partner organisations.

## Technologies Used

- Python
- pandas
- openpyxl
- Excel

## Files

- `allocation_algorithm.py` — runs the main allocation logic
- `total_amount_report.py` — calculates total donated items by category
- `sample_input.xlsx` — anonymised sample spreadsheet used as input data

## How to Run

Install the required packages:

```bash
pip install pandas openpyxl
```

Run the main allocation algorithm:

```bash
python allocation_algorithm.py
```

Run the total donation report:

```bash
python total_amount_report.py
```

## Algorithms

### 1. Main Allocation Algorithm

The main allocation algorithm reads donation records from the Excel spreadsheet and allocates school-supply items between partner organisations using rule-based logic.

It considers:

- item category
- total quantity donated
- essential supply needs
- organisation-specific priorities
- manual notes in the spreadsheet
- items that should not be split randomly

This algorithm is designed for real donation constraints rather than a generic mathematical model. For example, some supplies are prioritised for one organisation because they are more useful for that partner’s needs.

### 2. Total Donation Report Algorithm

The total donation report script reads the same Excel spreadsheet and calculates the total amount of donated items by category.

It is used to create a clearer summary of the overall donation drive, making it easier to understand what was collected before making allocation decisions.

## Privacy Note

The uploaded spreadsheet is a sample/anonymised version of the original donation data. Private donor, student, and school information has been removed.
