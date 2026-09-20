# Franchise Project Lifecycle Tracker (PMP® Aligned)

A comprehensive, cloud-based **Project Management & Lifecycle Automation System** built using **Google Apps Script** and **Google Sheets**. Designed strictly in accordance with **PMI PMBOK® Guide Standards**, this solution enables Project Managers (PMs) and Operations Heads to oversee multi-site franchise deployment, enforce strict governance, track Work Breakdown Structure (WBS) deliverables, and execute seamless project closures.

---

## 🌟 Key Features & PMP Alignment

### 1. Phase Gate Governance (Lifecycle Control)
- Structured transition through 6 distinct project lifecycle phases:
  1. **Initiation** (Pre-Start)
  2. **Planning & Procurement** (Material Receiving)
  3. **Execution & Installation** (Material Assembly)
  4. **Quality Control & Testing** (Soda Filling - Quarantine)
  5. **Operational Readiness** (FG Filling - Quarantine)
  6. **Operations & Handover** (FG Completed / Live)

### 2. Owner Governance & Direct Override (No Pop-up Delay)
- Eliminates multi-step prompt dialogs for rapid, cell-level status updates.
- Enables direct control over Work Package (WBS) statuses (`Done`, `In Progress`, `Issue`, `Closed`, `Not Started`).

### 3. Dual Audit Trail Logging
- **Field Operations Feed (`Submissions`):** Append-only log capturing field resource submissions via Google Forms.
- **PM Governance Log (`Owner_Audit_Log`):** Automated trigger (`onEdit`) recording every direct cell change made by the Project Manager (timestamp, cell location, old value, new value).

### 4. Technical Staff Registry & Email Allocation
- Dedicated **`Team_Master`** registry for technical resources (Staff ID, Name, Email, Phone, Role).
- Dynamic dropdown allocation in the **`Projects`** charter.
- One-click manual email notification trigger with selection checkboxes to notify assigned resources instantly.

### 5. Administrative Closure & Archiving
- Automated transfer of 100% completed or retired franchise records to dedicated archive sheets (`Archived_Projects` and `Archived_Submissions`).
- Maintains optimal performance and zero clutter in active operational dashboards.

---

## 📂 System Architecture & Sheets Layout

| Sheet / Tab Name | PMP Equivalent Term | Functional Description |
| :--- | :--- | :--- |
| **`Projects`** | Project Charter & Registry | Master registry of active franchise sites, timelines, assigned resources, and lifecycle phases. |
| **`Master`** | WBS Master Dictionary | Dictionary defining all 21 Work Package tasks, stage mappings, and valid status dropdown options. |
| **`Team_Master`** | Human Resource Registry | Staff directory containing contact details and skills for dynamic allocation. |
| **`Submissions`** | Work Performance Data | Live stream of field technician progress updates submitted via Google Forms. |
| **`Queries`** | Issue Log & Risk Register | Field-raised technical concerns, procurement risks, and change requests. |
| **`Owner_Audit_Log`** | Change & Audit Log | System-generated log capturing all direct administrative overrides and edits. |
| **`Dashboard`** | Portfolio Performance Rollup | Executive overview featuring stage pipeline visualizers and completion KPIs. |
| **`Franchise View`** | Stakeholder Progress Report | Single-franchise printable/exportable status report designed for individual franchisees. |
| **`Archived_Projects`** | Historical Closure Registry | Permanent repository for closed and archived project charters. |

---

## ⚙️ Installation & Setup Guide

1. **Open Google Sheet:** Create a new Google Sheet (or import the workbook) and rename it to `Franchise_Project_Lifecycle_Tracker`.
2. **Access Script Editor:** Navigate to **Extensions > Apps Script**.
3. **Deploy Code:** Replace all content in `Code.gs` with the provided script in this repository and save (`Ctrl + S`).
4. **Initialize System:** Refresh your Google Sheet. Click the new custom menu:
   $$\text{👑 Project Governance (PMP)} \rightarrow \text{1. Initialize Architecture \& Master Sheets}$$
5. **Populate Resource Master:** Open the newly generated **`Team_Master`** sheet and add your technical staff names, email IDs, and phone numbers.
6. **Assign & Notify:** Select staff members from the dropdowns in the **`Projects`** sheet, check the `Send Email?` checkbox, and click:
   $$\text{👑 Project Governance (PMP)} \rightarrow \text{2. Send Staff Allocation Emails}$$

---

## 🔐 Security & Access Model

- **Field Technicians / External Resources:** Access restricted to Google Form links only. No direct access to the Google Sheet.
- **Project Manager / Sponsor:** Exclusive Edit Access to the central control spreadsheet with automated audit logging enabled.
- **Franchisees / External Stakeholders:** Read-only export/PDF via the `Franchise View` tab.

---

## 👤 Author & Governance

- **Project Lead / Operations Head:** Sivakumar Krishnan (`mksiva0699@gmail.com`)
- **Methodology Framework:** PMI PMBOK® Guide Standards
