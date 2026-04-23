# Weyland-Yutani Mining Dashboard

A two-part data engineering and analytics project built around a configurable <b>Google Sheets data generator</b> and a separate <b>Flask web dashboard</b> for statistical analysis, anomaly detection, visualization, and PDF reporting.

This project simulates realistic daily extraction data for multiple Weyland-Yutani Corporation mines and provides a polished analytical interface for exploring the generated time series.

## Project Overview

The project consists of two connected parts:

### 1. Google Sheets Data Generator

A publicly available Google Spreadsheet that generates realistic mine output data using <b>cell formulas only</b>.

The spreadsheet allows users to configure:

- number of mines and mine names
- date range
- distribution type (```Normal``` or ```Uniform```)
- distribution parameters
- smoothing / correlation
- day-of-week effects
- overall trend
- events representing spikes or drops
- automatically updating chart

The spreadsheet is structured into four sheets:

- <b>Settings</b> — control panel for generation parameters
- <b>Engine</b> — internal calculation layer with helper logic and intermediate computations
- <b>Data</b> — final clean output in long format: ```Date | Mine | Output```
- <b>Chart</b> — pivot-based visualization with an overall line chart

Google Sheets generator: https://docs.google.com/spreadsheets/d/1mWVv5q5YkU6wH3VcPybJfO-PCkLCHNrJQJW3VW_kor4/edit?usp=sharing

### 2. Web Analytics Dashboard

A separate Flask-based web application that reads the generated data source and performs statistical analysis.

The dashboard supports:

- descriptive statistics per mine and for total output
- anomaly detection using multiple methods
- interactive chart selection
- trendline fitting
- professional PDF report generation

Live application: https://remarkable-reprieve-production.up.railway.app/

## Features

### Data Generation in Google Sheets

- formula-based generator with no Apps Script
- realistic time series instead of pure white noise
- dynamic mine configuration
- configurable time horizon
- support for Normal and Uniform distributions
- dynamic distribution parameter block
- smoothing / correlation between neighboring days
- day-of-week seasonality factors
- configurable daily trend
- at least four event definitions with:
  - start date
  - duration
  - magnitude
  - probability
- clean output table for downstream analysis
- pivot table and summary line chart

### Web Dashboard

- reads data from the spreadsheet data source
- descriptive statistics for each mine and total output:
  - mean daily output
  - standard deviation
  - median
  - interquartile range
- anomaly detection with adjustable parameters:
  - IQR rule
  - z-score
  - distance from moving average
  - Grubbs' test
- multiple chart types:
  - line
  - bar
  - stacked
- trendline selection:
  - polynomial degree 1
  - polynomial degree 2
  - polynomial degree 3
  - polynomial degree 4
- highlighted anomalies on charts
- detailed PDF report export
- separate sections for spikes and drops
- UI and PDF layout designed to look professional

## Spreadsheet Architecture

### 1. Settings

This sheet acts as the control panel for the generator.

Users can adjust:

- start date
- number of days
- mine list
- distribution type
- distribution parameters
- day-of-week coefficients
- trend
- events

### 2. Engine

This is the internal calculation layer.

It contains:

- generated calendar structure
- intermediate calculations
- helper coefficients
- event logic
- smoothing logic
- derived values used to build the final output

This sheet is intended to power the generator rather than serve as the final user-facing dataset.

### 3. Data

This is the final clean dataset used by the dashboard.

Format:

```
Date | Mine | Output
```

### 4. Chart

This sheet contains pivot-based chart preparation and a combined line chart showing mine output over time.

## Dashboard Functionality

The Flask application retrieves the generated data and provides a full analytics workflow.

### Statistical Summary

For each mine and for the total output, the dashboard computes:

- mean daily output
- standard deviation
- median
- interquartile range

### Anomaly Detection

Users can select and configure one or more anomaly detection methods:

- IQR rule
- z-score
- distance from moving average
- Grubbs' test

### Visualization

The dashboard supports multiple chart types:

- line chart
- bar chart
- stacked chart

It also supports user-selected polynomial trendlines:

- degree 1
- degree 2
- degree 3
- degree 4

Anomalies are visually highlighted on the charts.

### PDF Reporting

The application can generate a detailed PDF report containing:

- selected statistics
- charts
- anomaly summaries
- separate sections for spikes and drops

The report is designed to be readable and presentation-ready.

## Tech Stack

### Backend

- Python
- Flask
- Pandas
- NumPy
- SciPy
- Matplotlib

### Reporting

- WeasyPrint

### Frontend

- HTML
- Flask templates

### Data Source

- Google Sheets

## Repository Structure

The structure of this repository:
```
project/
│
├── app.py
├── templates/
│   └── dashboard.html
├── README.md
└── requirements.txt
```

## How It Works

### Step 1 — Data Generation

The Google Spreadsheet generates realistic mine extraction time series using formulas and configurable parameters.

### Step 2 — Data Retrieval

The Flask dashboard accesses the generated Data sheet as the analysis source.

### Step 3 — Analysis

The application computes descriptive statistics and anomaly detection results.

### Step 4 — Visualization

The user selects chart types and trendline degree, and the dashboard renders updated charts.

### Step 5 — Reporting

The dashboard generates a PDF report with statistics, charts, and anomaly explanations.

## Key Design Decisions

### Long-format data source

The final dataset is stored in long format:

```
Date | Mine | Output
```

This makes the data source flexible when the number of mines changes and keeps the analytical layer easier to maintain.

### Separate calculation layer

The spreadsheet uses a dedicated Engine sheet so the final Data layer remains clean and dashboard-friendly.

### Realistic synthetic data

The generator was designed to avoid naive white-noise output by combining:

- baseline mine capacity
- distribution-driven randomness
- smoothing / correlation
- day-of-week effects
- trend
- event-based anomalies

## Deployment

### Google Sheets Generator

Public spreadsheet link:

https://docs.google.com/spreadsheets/d/1mWVv5q5YkU6wH3VcPybJfO-PCkLCHNrJQJW3VW_kor4/edit?usp=sharing

### Flask Dashboard

Live deployment on Railway:

https://remarkable-reprieve-production.up.railway.app/

## Possible Future Improvements

- support for affected mine count per event instead of probability
- more advanced event shapes such as bell curves
- downloadable CSV/Excel exports
- dashboard filters by mine and time range
- additional anomaly detection techniques
- authentication and protected report access
- richer PDF styling and branding

## Why This Project Matters

This project demonstrates skills relevant to data engineering and analytics:

- building configurable synthetic data sources
- structuring spreadsheet-based data generation systems
- designing clean analytical data layers
- implementing statistical analysis in Python
- detecting anomalies in time series
- creating interactive dashboards
- generating polished automated reports
- integrating spreadsheet data sources with web applications
