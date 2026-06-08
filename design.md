# System Design

## Overview

This system converts structured Excel data into a PowerPoint presentation using a mapping configuration.

---

## Core Components

### 1. Excel Reader

Reads workbook using openpyxl.

Responsibilities:

* Load sheets
* Extract ranges
* Return raw data

---

### 2. Validator

Ensures:

* Required sheets exist
* Required ranges exist
* Data types are correct

---

### 3. Transformer

Converts raw Excel → structured JSON model

Example:

```json
{
  "on_time_delivery": {
    "categories": ["App1", "App2"],
    "track1": [10, 20],
    "track2": [15, 25]
  }
}
```

---

### 4. Config (YAML)

Defines mapping between:

* Excel → Slides

Example:

```yaml
slides:
  - id: delivery
    type: bar_chart
    source:
      sheet: Master-Apps (DoNotChange)
      categories: B5:B10
      series:
        - name: Track 1
          values: D5:D10
```

---

### 5. PPT Generator

Uses PptxGenJS to:

* Create slides
* Add charts
* Add text
* Add tables

---

## Data Flow

```
Excel → Reader → Validator → Transformer → JSON → PPT Generator → PPT
```

---

## Design Principles

* Config-driven
* Modular
* Extensible
* Fail-fast validation
* No direct PPT editing

---

## Slide Strategy

### Fully Automated

* Charts
* KPIs
* Tables

### Semi-Automated

* Highlights text

### Manual (rare)

* Custom commentary slides
