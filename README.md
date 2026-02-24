# 📦 Sales & Stock Operation (SO) Automation System

A Google Sheets–based Sales & Stock Operation (SO) automation system powered by Google Apps Script.
This system manages product distribution, promotional stock, returns, expiration tracking, and automated monthly archiving with dynamic analytics.

---

## 📌 Overview

This application was built to digitize and automate manual stock operation workflows commonly used in distribution-based businesses (such as pharmacy, retail, or product sales environments).

The system integrates:

* 📊 Real-time operational tracking
* 🔁 Automated stock movement logic
* 📆 Monthly archival with dynamic period grouping
* 📈 Automatic recap & analytics generation
* ☑️ Conditional checkbox-driven workflow automation

All processes run directly inside Google Sheets using custom Apps Script logic.

---

## 🚀 Core Features

### 1️⃣ Distribution Management

* Tracks distributed products per outlet
* Auto-handles:

  * Sold status (`Sudah` / `Belum`)
  * Partial sales confirmation
  * Automatic stock deduction
* Smart checkbox activation based on logic conditions
* Timestamp auto-update when status changes

---

### 2️⃣ Promo Stock Handling

* Separate promo program tracking
* Outlet-based distribution
* Independent tracking from regular distribution
* Auto integration into history & archive

---

### 3️⃣ Expiry (EXP) Monitoring

* Detects expired products
* Automatically moves expired stock to EXP sheet
* Real-time expired product count in dashboard

---

### 4️⃣ Return (Retur) System

* Tracks returned products
* Completion status via checkbox
* Real-time count of:

  * Returned (Completed)
  * Pending (Not completed)

---

### 5️⃣ History Dashboard (Real-Time Analytics)

Automatically updates:

* ✔️ Retur Sudah
* ⏳ Retur Belum
* 📦 Total EXP

Live recap updates whenever sheet data changes.

---

### 6️⃣ Monthly Archive System (Dynamic Period Labeling)

Each month:

* Automatically generates a new:

  ```
  PERIODE: Month Year
  ```
* Data is grouped under its respective period
* Dynamic recap per period:

  * Total Distribution Qty
  * Total Promo Qty
  * Recap statistics from History

The system calculates totals only within the active period block, stopping at the next period label.

---

## 🧠 Smart Automation Logic

This project heavily utilizes:

* `onEdit(e)` event triggers
* Conditional UI behavior (checkbox injection)
* Range-based period detection
* Dynamic formula generation
* Auto label regeneration
* Controlled data flow between sheets

The logic ensures:

* No duplicate movements
* No inconsistent stock subtraction
* Fully automated recap integrity

---

## 🏗 System Architecture

Sheets Structure:

```
Input_Data
Distribusi
Promo
Retur
EXP
History
Arsip
```

Flow Overview:

```
Input_Data
   ↓
Distribusi / Promo
   ↓
History (Real-time recap)
   ↓
Arsip (Monthly grouping & archive)
```

---

## ⚙️ Technologies Used

* Google Sheets
* Google Apps Script (JavaScript-based)
* Spreadsheet event triggers
* Dynamic formula injection (LET, MATCH, INDIRECT)
* Custom automation logic

---

## 🎯 Problem Solved

Traditional stock operations often suffer from:

* Manual recap errors
* Duplicate stock deduction
* Lost historical tracking
* No per-period grouping
* No real-time monitoring

This system solves those issues by:

* Automating movement logic
* Locking stock consistency rules
* Structuring historical data by period
* Providing instant recap visibility

---

## 💡 Key Highlights

* Fully automated period-based archiving
* Smart conditional UI logic
* Zero manual recap calculation
* Clean separation between operational sheets and archival records
* Scalable structure for additional analytics

---

## 📊 Project Type

Internal Operations Automation Tool
Portfolio Demonstration Project
Business Process Optimization System

---

> [!IMPORTANT]
> **Disclaimer:** > Kode yang ditampilkan dalam repositori ini merupakan **sampel (logic demo)** untuk keperluan portofolio. Implementasi ini mencerminkan arsitektur dasar dari sistem yang lebih kompleks dan bertujuan untuk menunjukkan keahlian dalam logika pemrograman dan otomatisasi alur kerja.

---
*Developed by **Muhammad Farhan Putra Pratama, S.H.*** *Sebuah solusi efisiensi administrasi kesehatan berbasis teknologi otomasi.*


<img width="978" height="890" alt="image" src="https://github.com/user-attachments/assets/2abf3aa0-5fa1-4f1b-bf57-d63fbe6c4faf" />
<img width="1146" height="863" alt="image" src="https://github.com/user-attachments/assets/6e331dd0-aabb-4666-b037-174d5918392e" />
<img width="1209" height="862" alt="image" src="https://github.com/user-attachments/assets/ad7c3cb7-afab-45d7-84b9-302ca0872f9f" />
<img width="1382" height="862" alt="image" src="https://github.com/user-attachments/assets/2a7d87e7-73c6-41ab-95ca-82f3223c738e" />
<img width="904" height="865" alt="image" src="https://github.com/user-attachments/assets/e7e71c94-4385-4620-bd88-93f9ad05a679" />
<img width="1004" height="865" alt="image" src="https://github.com/user-attachments/assets/e5ee6cf7-d5fa-4350-8c01-3c1e00f0371a" />
<img width="1920" height="884" alt="image" src="https://github.com/user-attachments/assets/76721603-2bb2-413c-939d-aa31c4abd4e9" />
