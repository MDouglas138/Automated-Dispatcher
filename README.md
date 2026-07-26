# Automated Inventory & Restock Dispatcher



## Description

I built this application to automate the analysis and distribution of dealership inventory data. It combines a graphical interface for scheduling and managing email lists with a background processor that reads Excel files, calculates inventory metrics, and automatically emails a PDF report to designated teams.

## Tech Stack

* **Language:** Python


* **GUI Framework:** Tkinter with ttkbootstrap


* **Database Management:** SQLite3


* **Data Processing & Reporting:** Pandas and FPDF


* **Task Integration:** Windows Task Scheduler via PowerShell



## Key Features

* **Dynamic Task Scheduling:** Uses a custom GUI to configure tasks that trigger via Windows Task Scheduler on once, daily, weekly, or monthly intervals.


* **Contact Management:** Includes an integrated system to add, group, and validate email addresses for report distribution.


* **Data Analysis:** Processes Excel workbooks using Pandas to calculate advanced metrics like pacing status, months of supply, and restock targets.


* **Automated PDF Reports:** Generates a structured "Dealership Intelligence Report" highlighting critical shortages, actionable dealer trades, unmatched overstock, and dead stock.


* **Secure Email Dispatch:** Automatically attaches the generated PDF and emails it to scheduled recipients using an SMTP server and a secure `.env` file for credentials.


## How to Run

1. Ensure Python, Pandas, FPDF, and ttkbootstrap are installed on your local machine.
2. Create a `.env` file in the root directory containing your secure `EMAIL_PASSWORD`.
3. Clone this repository to your local environment.
4. Execute the main script to launch the application interface:


python dispatcher.py

```
