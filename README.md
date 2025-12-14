# Guardium-Automation-Monitor: CM Processor

## 🛡️ Overview

This Python script was developed to **automate and centralize the monitoring and health check process** for IBM Security Guardium environments.

It processes raw logs and reports from Central Management (CM) and Collectors, transforming spreadsheet data into a clear, structured executive report (**Word and Excel formats**). This dramatically simplifies and accelerates the identification of critical operational issues.

**Motivation:** Consolidate the analysis of multiple operational logs (Agent Status, Aggregation Processes, Collection Quality, etc.) into a **single, automated, and fast** workflow.

---

## ✨ Key Monitoring Features

The script focuses on the following critical health indicators:

* **Agent Status (STAP)**

  * Counts **Active vs. Inactive** agents.

* **Inactive Agent Detail**

  * Generates a **detailed table** in the Word report listing all **Inactive STAPs**, including:

    * Host
    * Version (Revision)

* **Aggregation Failures**

  * Identifies failures in critical processes such as:

    * Purge
    * Export
    * Archive
  * Filters out successful executions.
  * Displays failures in the following format:

    ```
    [Collector Name] - [Failed Process (Status)] - [Most Recent Failure Date]
    ```

---

## ⚙️ Setup and Dependencies

### Requirements

* Python **3.6+**

### Python Dependencies

Install the required libraries using `pip`:

```bash
pip install pandas openpyxl python-docx
```

---

## 📁 Working Directory Structure

The script automatically creates and manages a working directory named `CM/` at the project root.

> ⚠️ On startup, the script **cleans this directory** to guarantee a fresh execution.

### Directory Layout

```text
.
└── CM/
    ├── Central Management/          # INPUT: Central Management report
    ├── STAP status/                 # INPUT: STAP status logs
    ├── Processos de agregação/      
    │   └── [Collector Hostname]/    # INPUT: Aggregation process logs
    ├── Qualidade da coleta/         # INPUT: Collection quality logs
    └── output/                      # OUTPUT: Word and Excel reports
```

---

## 🚀 How to Use (Step-by-Step)

The script guides the user through the entire process via **interactive terminal prompts**.

---

### ▶️ Step 1 — Run the Script

Navigate to the root directory of the project:

```bash
python CM-Processor/cm_processor.py
```

At startup, the script:

* Cleans old files
* Validates and recreates the `CM/` directory structure

---

### 🧭 Step 2 — Central Management Spreadsheet

**Prompt:**

```text
➡ Place the Central Management spreadsheet in the 'Central Management' folder and press ENTER...
```

**User Action:**

* Place the Central Management spreadsheet inside:

  ```
  CM/Central Management/
  ```
* The spreadsheet must contain at least:

  * `Unit name`
  * `Unit type`

**Script Behavior:**

* Automatically detects all **Collectors**
* Creates dedicated subfolders for each collector, for example:

```text
CM/Processos de agregação/collector.prd.01/
```

---

### 📊 Step 3 — Insert Detailed Logs

**Prompt:**

```text
➡ Now place the files into the subfolders (STAP status, Aggregation/Collector X) and press ENTER to start processing...
```

**User Action:**

* Place **STAP Status** spreadsheets in:

  ```
  CM/STAP status/
  ```

  Required columns:

  * Host
  * Status
  * Revision

* Place **Aggregation Process logs** in the corresponding collector folder:

  ```
  CM/Processos de agregação/[Collector Hostname]/
  ```

  Required columns:

  * Activity Type
  * Status
  * Date

---

## 📦 Final Output

After processing, the script generates the following files in:

```text
CM/output/
```

### 📄 Relatorio_Executivo.docx

* Executive summary
* Total Active vs. Inactive STAPs
* Detailed table of **Inactive Agents**
* Aggregation process failures

### 📊 CM_report.xlsx

* Full STAP inventory
* Raw aggregation failure records
* Structured data for auditing and troubleshooting

---

## ✅ Benefits

* Eliminates manual log analysis
* Prevents counting errors (per-host vs per-agent)
* Improves incident response time
* Provides a clear executive-ready report

---

🛡️ **Guardium Automation Monitor — turning raw logs into actionable intelligence.**
