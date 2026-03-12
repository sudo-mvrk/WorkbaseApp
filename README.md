# Freelance Billing & Order Automation 💼

A modular desktop application built with Python and PyQt6, designed to streamline order management, dynamic pricing calculations, and financial reporting for a freelance design business.

## 🎯 The Problem It Solves
Manual calculation of design orders and invoice generation is prone to human error and time-consuming. This application automates the entire workflow: from calculating the cost of individual layout designs to exporting ready-to-send financial reports in Excel, JSON, and XML formats.

## 🏗️ Architecture & Engineering Focus
Originally a monolithic script, this project has been heavily refactored into a maintainable, decoupled architecture following **Separation of Concerns (SoC)** principles:

* **`core.py`**: Pure business logic and mathematical calculations, isolated from the UI.
* **`io_handlers.py`**: Data persistence layer handling JSON, XML, and Excel (OpenPyXL) operations without direct dependency on GUI components.
* **`ui_components.py`**: Reusable custom PyQt6 widgets (e.g., Drag & Drop tables, custom delegates).
* **`main_window.py`**: The orchestration layer that connects the UI with the underlying business logic.

## ✨ Key Features
* **Dynamic Cost Calculation**: Real-time total updates based on item price and quantity.
* **Drag-and-Drop Support**: Instantly load previous layouts or price lists by dragging JSON/XML files into the application.
* **Multi-Format Export**: Generate professional `.xlsx` spreadsheets, structured `.json` data, or plain text reports.
* **Clean UI/UX**: Intuitive, clutter-free interface optimized for daily business use.

## 🛠️ Tech Stack
* **Language:** Python 3
* **GUI Framework:** PyQt6
* **Data Handling:** OpenPyXL (Excel), built-in JSON/XML parsers
* **Dependency Management:** Poetry

## 🚀 How to Run

1. Clone the repository:
   ```bash
   git clone [https://github.com/sudo-mvrk/WorkbaseApp](https://github.com/sudo-mvrk/WorkbaseApp)
   cd WorkbaseApp
   ```
2.Install dependencies using Poetry:
  ```bash
  poetry install
  ```
3. Run the application:
  ```bash
  poetry run python main.py
```
