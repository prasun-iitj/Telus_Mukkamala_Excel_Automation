# Requirements

## Functional Requirements

### 1. Excel Input

* Accept Excel workbook (.xlsx)
* Must contain:

  * Master-Apps (DoNotChange)
  * Master-Platforms (DoNotChange)

---

### 2. Data Extraction

* Read:

  * cell values
  * column ranges
  * row ranges
* Support merged cells
* Ignore formulas (use computed values)

---

### 3. Slide Generation

Support following slide types:

* title_slide
* bar_chart
* line_chart (future)
* kpi
* table
* text_block

---

### 4. Chart Support

* Multiple series
* Category labels
* Numeric validation
* Consistent styling

---

### 5. Text Handling

* Multi-line text
* Bullet points
* Overflow handling (truncate/wrap)

---

### 6. Validation

* Required sheets must exist
* Required ranges must exist
* Missing data should raise warnings/errors

---

### 7. Output

* Generate `.pptx`
* Maintain consistent layout
* Naming convention:

  ```
  <Project>_<Month>_<Year>.pptx
  ```

---

## Non-Functional Requirements

* Fast execution (<10 seconds)
* Modular design
* Config-driven (no hardcoding)
* Easy to extend for new slides

---

## Constraints

* Excel structure must remain consistent
* Slide layout is fixed template-based
* Charts recreated (not copied)
