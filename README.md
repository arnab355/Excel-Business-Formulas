# Excel Business Formula Integration

This repository contains my completed Excel project focused on applying core business formulas and functions to solve practical data challenges. The project covers conditional logic, dynamic array filtering, lookup functions, statistical calculations, and text manipulation.

---

## What's Inside

The workbook (`module-2-business-formula-practice.xlsx`) features three organized worksheets:

1. **`Sales Data`**
   * Preprocessed 15 transaction records across various sales regions (`North`, `South`, `East`, `West`) and products (`Laptop`, `Mobile`, `Tablet`).
   * **Calculated Fields:** Computed total sales (`Quantity × Unit Price`).
   * **Logic & Conditions:** Applied `IF`, `AND`, and `OR` functions to tag sale status (`High Sale` vs `Low Sale`) and salesperson performance (`Excellent`, `Good`, `Average`).
   * **Text Formatting:** Combined rep details into a custom `Sales Info` field using `TEXTJOIN`.
   * **Summary Stats:** Included dynamic summary formulas using `SUMIF`, `SUMIFS`, `COUNTIF`, `COUNTIFS`, `AVERAGEIF`, and `AVERAGEIFS`.
   * **Lookups:** Implemented `VLOOKUP` and `INDEX + MATCH` to retrieve product prices dynamically.

2. **`High Value Sales`**
   * Utilized the dynamic array `FILTER` function to automatically display transactions with total sales exceeding **$2,000.00**.

3. **`Product Price List`**
   * Serves as the central reference table containing standard pricing for all product categories.

---

## Formulas & Functions Used

* **Lookup & Reference:** `VLOOKUP`, `INDEX`, `MATCH`
* **Dynamic Arrays:** `FILTER`
* **Logical Functions:** `IF`, `AND`, `OR`
* **Conditional Math & Stats:** `SUMIF`, `SUMIFS`, `COUNTIF`, `COUNTIFS`, `AVERAGEIF`, `AVERAGEIFS`
* **Text Functions:** `TEXTJOIN`
