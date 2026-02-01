# TNB Smart Meter Reporting Automation (LKS Pipeline)

## 🚀 Overview

A high-performance **ETL (Extract, Transform, Load) Pipeline** built to automate the processing of *Laporan Kerja Selesai* (LKS) for Tenaga Nasional Berhad (TNB) Smart Meter installations.

This tool replaces a manual, error-prone workflow (copy-pasting rows & taking screenshots) with a single-click Python solution. It reduces weekly reporting time by **90%** (from 20 hours to <2 hours) and ensures 100% data compliance.

## ⚡ Key Features

- ** automated Data Ingestion**: Parses legacy `.xls`, `.xlsx`, and raw dump files using `pandas`.
- **Duplicate Detection**: Identifies and flags duplicate Service Orders (SO) to prevent billing errors.
- **Smart Image Injection**: Programmatically parses raw image URLs and injects dynamic Excel `=IMAGE()` formulas directly into cells, eliminating manual screenshotting.
- **Data Standardization**: Automatically cleanses headers, normalizes date formats, and validates "3MS" compliance rules.
- **GUI Dashboard**: Built with `CustomTkinter` for a user-friendly, drag-and-drop interface for non-technical staff.

## 🛠️ Tech Stack

- **Language:** Python 3.10+
- **Core Processing:** Pandas, NumPy
- **Excel Engine:** OpenPyXL, Xlwings
- **GUI:** CustomTkinter
- **QC/Validation:** Custom Logic Engine

## 📊 Workflow

1.  **Extract**: Ingest raw "Claim Sheets" and "Legacy Reports" from field technicians.
2.  **Transform**:
    -   Clean column headers to standard TNB format.
    -   Filter out "TRAS" (Terminated) orders.
    -   Merge duplicates.
3.  **Load**: Generate a strict format `.xlsm` submission file with embedded image evidence.

## 📈 Impact

-   **Speed:** Reduced daily processing time from **4 hours** to **30 seconds**.
-   **Quality:** Eliminated human copy-paste errors and "Missing Photo" rejections.
-   **Scalability:** Capable of processing 5,000+ rows per batch.

## 🔒 Confidentiality Note

*This repository contains the source code for the automation logic. All proprietary TNB data, config keys, and customer information have been removed for public demonstration.*

---
**Author:** Muhammad Syahmi Nuaim
**Role:** Data Automation Developer
