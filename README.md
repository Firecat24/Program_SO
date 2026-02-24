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
*Developed by **Muhammad Farhan Putra Pratama, S.H.*** *Sebuah solusi efisiensi administrasi kesehatan berbasis teknologi otomasi.*
