# 🍵 Kombucha Brewery — Dynamic Production Planning AI Agent

[![Python](https://img.shields.io/badge/Python-3.10+-3776AB?style=flat-square&logo=python&logoColor=white)](https://python.org)
[![pandas](https://img.shields.io/badge/pandas-2.x-150458?style=flat-square&logo=pandas)](https://pandas.pydata.org)
[![OpenPyXL](https://img.shields.io/badge/openpyxl-3.x-217346?style=flat-square&logo=microsoft-excel)](https://openpyxl.readthedocs.io)
[![Live Demo](https://img.shields.io/badge/Live%20Demo-GitHub%20Pages-0A66C2?style=flat-square&logo=github)](https://YOUR-USERNAME.github.io/kombucha-production-planner)
[![IE Concepts](https://img.shields.io/badge/IE-MRP%20%7C%20Capacity%20%7C%20Scheduling-orange?style=flat-square)](#industrial-engineering-concepts)

> **Industrial Engineering Portfolio Project** — A Python-based AI agent that automates production planning for a kombucha brewery. Reads demand from Excel, plans fermentation batches, allocates tanks, schedules bottling, detects bottlenecks, and writes the full production plan back to Excel — automatically.

**[🚀 Try the Live Web App →](https://YOUR-USERNAME.github.io/kombucha-production-planner)**

---

## 📋 Table of Contents

- [Demo](#-demo)
- [What It Does](#-what-it-does)
- [System Architecture](#-system-architecture)
- [Quick Start](#-quick-start)
- [Excel Workbook Structure](#-excel-workbook-structure)
- [Industrial Engineering Concepts](#-industrial-engineering-concepts)
- [Project Structure](#-project-structure)
- [Future Improvements](#-future-improvements)

---

## 🎥 Demo

| Dashboard | Bottleneck Report | Ingredient Alerts |
|-----------|-------------------|-------------------|
| KPI cards + full schedule | 16 actionable insights | 5 shortage alerts |

**Live agent output (March 2026 planning cycle):**
```
📋 Production batches planned  : 14
✅ On schedule                 : 6  (43%)
⚠️  At risk of being late       : 8  (57%)
🔔 Alerts generated            : 16
🔴 Ingredient shortages        : 5 materials
```

---

## ✅ What It Does

The agent executes a **9-step production planning pipeline**:

| Step | Function | Output |
|------|----------|--------|
| 1 | `load_data()` | Loads 7 Excel sheets into DataFrames |
| 2 | `forecast_demand()` | Aggregates demand by flavor & priority |
| 3 | `convert_to_batches()` | Splits demand into feasible tank batches |
| 4 | `check_materials()` | BOM explosion — checks inventory shortages |
| 5–6 | `assign_tanks_and_schedule()` | Greedy EDD tank assignment + fermentation dates |
| 7 | `schedule_bottling()` | Assigns bottling slots, detects capacity overloads |
| 8 | `detect_bottlenecks()` | Consolidates all alerts and insights |
| 9 | `write_to_excel()` | Writes schedule + 3 new sheets back to Excel |

**Example insights generated:**
```
🔵 Tank T-03 will become available on 2026-03-18.
🔴 SHORTAGE: 'Blueberry Puree' needs 16 kg but only 6 in stock.
⚠️  Ginger Lemon batch (due Mar 27) will bottle Apr 5 — AT RISK.
▶  Start brewing Mango Turmeric tomorrow to meet demand.
```

---

## 🏗 System Architecture

```
Excel Workbook (Input)
    ├── Orders           → Customer demand + due dates
    ├── Inventory        → Raw material stock levels
    ├── Recipe_BOM       → Bill of Materials per flavor
    ├── Fermentation_Tanks → Tank capacity + availability
    └── Bottling_Line    → Line speed + operating hours
            │
            ▼
    kombucha_agent.py
    ┌─────────────────────────────────────────────────────┐
    │  load → forecast → batch → BOM check → tank assign  │
    │       → fermentation schedule → bottling schedule    │
    │       → bottleneck detection → write to Excel        │
    └─────────────────────────────────────────────────────┘
            │
            ▼
Excel Workbook (Output)
    ├── Dashboard           (KPI cards + full schedule)
    ├── Production_Plan     (14 batches with dates)
    ├── Bottleneck_Report   (16 insights, color-coded)
    └── Ingredient_Alerts   (shortage quantities)
```

---

## 🚀 Quick Start

### 1. Clone the repo
```bash
git clone https://github.com/YOUR-USERNAME/kombucha-production-planner.git
cd kombucha-production-planner
```

### 2. Install dependencies
```bash
pip install pandas openpyxl
```

### 3. Build the Excel workbook (first time only)
```bash
python src/build_workbook.py
```

### 4. Run the production planning agent
```bash
python src/kombucha_agent.py
```

### 5. Open `data/kombucha_brewery.xlsx` → navigate to the **Dashboard** tab

---

## 📊 Excel Workbook Structure

The workbook (`data/kombucha_brewery.xlsx`) is both the **data interface** and **output layer**:

| Sheet | Type | Description |
|-------|------|-------------|
| `Orders` | INPUT | 10 sample orders across 5 flavors |
| `Inventory` | INPUT | 16 ingredients + packaging items |
| `Recipe_BOM` | INPUT | Bill of Materials for all 5 flavors |
| `Fermentation_Tanks` | INPUT | 8 tanks (200L, 150L, 100L) |
| `Bottling_Line` | INPUT | 3 lines with speed + hours |
| `Production_Plan` | **OUTPUT** | Full schedule with status flags |
| `Actual_Production` | LOG | Manual yield tracking |
| `Dashboard` | **OUTPUT** | KPI summary + schedule view |
| `Bottleneck_Report` | **OUTPUT** | All insights, color-coded by severity |
| `Ingredient_Alerts` | **OUTPUT** | Shortage table |

---

## 🔬 Industrial Engineering Concepts

This project demonstrates core IE and operations management principles:

| Concept | Application |
|---------|-------------|
| **MRP (Material Requirements Planning)** | BOM explosion × batch count vs. current inventory |
| **Capacity Planning** | Tank utilization scheduling + bottling line constraint |
| **Scheduling Theory** | EDD (Earliest Due Date) greedy policy with priority sort |
| **Theory of Constraints** | Bottleneck detection across materials, tanks, bottling |
| **Inventory Management** | Safety stock alerts + reorder point tracking |
| **Discrete-Event Logic** | State transitions: Available → Fermenting → Bottling → Done |
| **Quality Metrics** | Yield % tracking (produced vs. waste) in Actual_Production |

---

## 📁 Project Structure

```
kombucha-production-planner/
├── README.md
├── requirements.txt
├── .gitignore
├── src/
│   ├── kombucha_agent.py      # Main AI planning agent (9-step pipeline)
│   └── build_workbook.py      # One-time Excel workbook builder
├── data/
│   └── kombucha_brewery.xlsx  # Pre-loaded sample workbook
└── docs/
    └── index.html             # Live web app (GitHub Pages)
```

---

## 🔮 Future Improvements

- [ ] **LP Optimizer** — Replace greedy tank assignment with PuLP linear programming (target: 15–25% makespan reduction)
- [ ] **ARIMA Forecasting** — Rolling demand forecast to reduce safety stock by 10–15%
- [ ] **Arena DES Integration** — Stress-test scheduling under ±20% demand variance
- [ ] **ML Yield Prediction** — Train on Actual_Production to predict batch waste rates
- [ ] **Supplier API** — Real-time lead time and price data via REST API
- [ ] **Power BI Dashboard** — Executive reporting layer on top of Excel output

---

## 📄 License

MIT License — free to use, fork, and build on.

---

*Built for the IE Production Planning & Simulation portfolio | Bradley University*
