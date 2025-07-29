# Replay Institute - Google Sheets Screening Automation

This repository contains the Apps Script code used to automate a screening process for **Replay Baseball Institute**. The system is built on Google Sheets and designed to help staff efficiently evaluate new members of their program based on key athletic metrics.

## 📌 Project Overview

Replay Baseball Institute conducts athletic screenings for potential new members. These screenings measure:

- Mobility
- Strength
- Power
- Rotation
- Arm strength and health

This project streamlines the data entry and reporting process by adding custom functionality to Google Sheets via a custom dropdown menu. The result is a more efficient, automated experience for staff during evaluations.

## ⚙️ Key Features

A custom dropdown menu provides staff with **three automated actions**:

### 1. **Add to Database**
- After a screening is completed, staff select this option.
- They're prompted to enter the athlete's **first and last name**.
- All screening data is then automatically transferred to a master database under the athlete’s name.
- ✅ Eliminates the need for manual data entry into the database.

### 2. **Clear Screening Sheet**
- Used immediately after a player's data has been stored.
- This function clears all fields in the screening sheet, resetting it for the next athlete.
- ✅ Saves time by avoiding manual clearing.

### 3. **Generate Slide for Player**
- Staff enter a player's first and last name to retrieve stored screening results.
- The system:
  - Locates the player's data in the database
  - Creates a **copy of a template slide** in Google Slides
  - Automatically fills in the copied slide with the player's screening results
- ✅ Provides a professional, visual report for communication or internal use.

## 🧰 Tech Stack

- **Google Apps Script**
- **Google Sheets**
- **Google Slides (via Apps Script)**
- **HTML/CSS (for custom UI prompts)**

## 🛠️ Development Environment

- **Google Apps Script Editor** (via Google Sheets)
- Git + GitHub used only for archiving code using [`clasp`](https://github.com/google/clasp)

## 🔒 Important Notes

- To keep this repository clear and focused, some supporting files—such as the HTML interface files and `.clasp.json` configuration—have been intentionally excluded.
- The main purpose is to highlight the core Apps Script (macro) code that drives the project’s functionality.
- This repo serves as an archival and demonstration resource. For live usage or further development, clone the repo and link it to your own Apps Script project using `clasp`.



