# 492-GROUP2
Final project source code and outputs
# CNC Overtime Minimization Heuristic

## Overview
This repository contains the source code and outputs of our graduation project. The project proposes a heuristic optimization approach for CNC production scheduling with the objective of minimizing total overtime while satisfying all production constraints.

## Objective
The optimization model minimizes the total overtime of CNC machines while considering:

- Machine availability
- Operation precedence (Op10 → Op20)
- Setup times
- Machine group eligibility
- Quantity conservation
- Maximum allowable tardiness

## Project Structure

- `main.py` – Main optimization algorithm
- Input Excel files
- Output Excel reports
- Gantt chart (HTML)
- Constraint verification report

## Requirements

- Python 3.x
- pandas
- numpy
- openpyxl
- plotly

Install the required packages using:

```bash
pip install pandas numpy openpyxl plotly
```

## Outputs

The algorithm generates:

- Optimized production schedule
- Machine utilization summary
- Constraint verification report
- Interactive Gantt chart
- Performance statistics

## Authors

EMRE CEM ARAYICI
DILA DONMEZ
ENES GUNEY
ILAYDA KESKIN

Department of Industrial Engineering
