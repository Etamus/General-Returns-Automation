# Reverse Logistics Automation (VBA)

**Automated Excel macros for managing reverse logistics processes, including transport, remittance, and order adjustments.**

---

## Overview

This repository hosts the fully revised and optimized version of the **Planilha Reversa** project. All macros for reverse logistics have been unified into a single environment, providing better organization, improved performance, and a wide range of automated functionalities.

---

## Features

The current version includes:

- **General Alterations**: OI (Inverse Orders), Remittance, and TR adjustments  
- **Pre-Calculation Execution**  
- **Transport and Remittance Creation**  
- **Cancel, Reactivate, and Reset Inverse Orders**  
- **NFD Modification**  
- **ZREC Cancellation**  
- **RFQ Adjustments**  
- **Occurrence Recording**  
- **Action/Providence Logging**  
- **Weight Lookup**  
- **Access Key and MLOG Search**  
- **Address Lookup**

---

## Access Control

To prevent simultaneous usage, the workbook has a **temporary lock system**. Access is granted via a code and lasts **1 hour**, after which the workbook locks automatically.  

**General access code:** `qb7p7Z001UQTwL`

**Network location:** `LOGISTICA\Atendimento ao Pedido\02. DEVOLUCAO\Reversa`

---

## User Access Levels

| Role                | Access Code            | Duration |
|--------------------|----------------------|---------|
| Monitoring/General | qb7p7Z001UQTwL       | 1 hour  |
| CDP                | wxFd4L99wx6O         | 1 hour  |
| Administrative     | By request only      | —       |

---

## Installation & Usage

1. Open the `.xlsm` file in Excel with macros enabled.  
2. Enter your access code when prompted.  
3. Navigate through the available macros via the **Macros menu** or assigned buttons.  
4. Follow prompts for each process (OI, Transport, RFQ, etc.).  

> ⚠️ Ensure macros are enabled and you have proper SAP connection if using automated SAP functionalities.

---

## File Structure

- `src/` – Exported VBA modules, classes, and forms (`.bas`, `.cls`, `.frm`)  
- `build/` – Macro-enabled Excel workbook (`.xlsm`) ready for use  
- `README.md` – Project overview and instructions  
- `.gitattributes` – For proper language detection and file handling

---

## Notes

- Designed for SAP-integrated reverse logistics processes.  
- Consolidates multiple macro scripts into a single, optimized workflow.  
- Provides secure access management to prevent conflicts.  

---


