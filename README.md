# Stats & Charts — Descriptive Statistics and Data Visualization PWA

A Progressive Web App for exploratory data analysis: load a dataset, get full descriptive statistics, normality testing, outlier detection, correlation matrix, and interactive charts — all in the browser, with no backend, no installation, no data leaving your device.

---

## Why This Exists

Quick EDA on small-to-medium datasets shouldn't require Python, R, or a cloud tool. This app gives an analyst or engineer immediate statistical feedback on any dataset — from a CSV upload or manual entry — directly on their phone or laptop.

---

## Features

### Data Input
- **Manual grid entry** — editable spreadsheet with add/remove rows and columns
- **CSV / TSV upload** — drag and drop or file picker
- **Formula columns** — computed columns via `fx` button
- **Multiple datasets** — tabbed interface, each with independent state
- **Undo / Redo** — full edit history (Ctrl+Z / Ctrl+Shift+Z)
- **Persistent storage** — data survives page refresh via localStorage

### Statistical Analysis
| Output | Detail |
|---|---|
| Descriptive stats | Count, Mean, Median, Mode, Std Dev, Variance |
| Distribution shape | Skewness, Kurtosis (excess) |
| Percentiles | Q1, Q2, Q3, IQR, Min, Max |
| Normality test | D'Agostino–Pearson (K² statistic) → Normal / Non-normal label |
| Outlier detection | IQR fence method — outliers highlighted directly in the grid |
| Correlation matrix | Pearson r for all numeric column pairs |
| Group statistics | Descriptive stats broken down by categorical variable |
| Effect size | Cohen's d and related metrics |
| Summary insights | Auto-generated plain-language interpretation of each column |

### Charts (15 types)
Bar · Line · Scatter · Histogram · Pie/Doughnut · Box Plot · Area · Radar · Polar Area · Bubble · Stem-and-Leaf · Cumulative Frequency · **Pareto** · **KDE (Kernel Density Estimation)**

Auto mode selects the most appropriate chart type based on the data structure.

### Export
- **CSV** — cleaned dataset
- **PNG** — any chart as image

---

## Technical Stack

- **PWA** — installable on iPhone/Android/desktop, works fully offline
- **Pure HTML / CSS / JavaScript** — no frameworks, no build step, no dependencies
- **Chart.js** — rendering engine for all chart types
- **Service Worker** — full offline caching
- **localStorage** — client-side persistence across sessions
- **Dark mode** — system-aware toggle

---

## Relevant Use Cases

- EDA on manufacturing process data before building a formal SPC model
- Quick distribution check and outlier flagging on a sample dataset
- Correlation screening between process variables
- Normality verification before applying parametric statistical tests

---

*Part of [Davide Incaini's portfolio](https://github.com/davideincaini)*
