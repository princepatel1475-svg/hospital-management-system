# Hospital Management System (Oracle Edition)

A GUI-based hospital management system built with **Python (Tkinter)** and **Oracle Database 21c XE**,
with live Excel synchronization and auto-calculated summary reporting.

## Screenshots

| Doctors | Billing |
|---|---|
| ![Doctors tab](screenshots/doctors.png) | ![Billing tab](screenshots/billing.png) |

| Pharmacy / Medical Store | Reports & Live Summary |
|---|---|
| ![Pharmacy tab](screenshots/pharmacy_medical_store.png) | ![Reports tab](screenshots/reports_summary.png) |

## Features
- Normalized relational schema across 7 tables: `patients`, `staff`, `doctors`, `appointments`,
  `bills`, `medicines`, `medicine_sales` (foreign keys, CHECK constraints, a generated/computed
  column for bill totals, sequences+triggers for IDs, and a trigger that auto-decrements
  pharmacy stock on every sale)
- Full CRUD GUI across 6 tabs: Patients, Staff, Doctors, Appointments, Billing, Pharmacy/Medical Store
- Every add/delete automatically re-syncs a linked Excel workbook (`hospital_data_export.xlsx`) — 7 sheets total
- One-click manual export/download button
- Auto-calculated summary statistics sheet, including total billing revenue, pending bills,
  and total pharmacy/medical-store revenue

---

## 1. Set your Oracle password (already done)

You reset it earlier:
```sql
ALTER USER system IDENTIFIED BY Prince1475;
```
So your credentials are `system` / `Prince1475`, service `XE`.

## 2. Load the schema

From PowerShell / Command Prompt:
```
sqlplus system/Prince1475@XE @schema_oracle.sql
```
(Run this from inside the `hospital_management_system` folder, or give the full path to `schema_oracle.sql`.)

This creates the 3 core tables, sequences, ID-generating triggers, indexes, and sample data.

Then add the extra modules (Doctors, Billing, Pharmacy/Medical Store):
```
sqlplus system/Prince1475@XE @schema_additions_oracle.sql
```

If you'd rather do it interactively:
```
sqlplus system/Prince1475@XE
SQL> @schema_oracle.sql
SQL> @schema_additions_oracle.sql
```

## 3. Install Python dependencies

```
pip install -r requirements_oracle.txt
```

`oracledb` runs in "thin" mode by default — **no need to install Oracle Instant Client separately**,
since you already have the full Oracle XE database installed.

## 4. Configure the connection (already set)

`app_oracle.py` already has:
```python
DB_CONFIG = {
    "user": "system",
    "password": "Prince1475",
    "dsn": "localhost:1521/XE"
}
```
Change these if you later create a dedicated non-`system` user, or your listener port isn't 1521.

## 5. Run the app

```
python app_oracle.py
```

Four tabs: **Patients**, **Staff**, **Appointments**, **Reports/Excel Export**.
Every add/delete auto-refreshes `hospital_data_export.xlsx` in the project folder.
Use the "Export/Refresh Excel Now" button any time for a manual re-sync.

---

## Project Structure
```
hospital_management_system/
├── schema_oracle.sql            # Core schema: patients, staff, appointments
├── schema_additions_oracle.sql   # Extra modules: doctors, billing, pharmacy/medical store
├── cleanup_oracle.sql             # Drops all objects (for a clean re-run)
├── app_oracle.py                   # Main GUI application (Oracle version)
├── requirements_oracle.txt          # Python dependencies
└── README_oracle.md
```

## CV Bullet Points (ready to use — Oracle version)
- Designed and implemented a relational database-backed hospital management system in Oracle DB, with normalized schemas spanning patients, staff, doctors, appointments, billing, and pharmacy/medical-store inventory.
- Built a multi-module GUI interface for real-time data entry and management, enforcing data integrity constraints (foreign keys, CHECK constraints, computed columns) across all operations.
- Implemented automated business logic at the database level, including a trigger that auto-decrements pharmacy stock on every medicine sale.
- Developed live Excel synchronization — any change made through the GUI automatically updates a linked, multi-sheet Excel workbook, with a one-click download option for offline access and reporting.
- Added auto-calculated summary statistics (total patients, daily appointments, billing revenue, pharmacy revenue, pending payments) within the Excel export, enabling quick at-a-glance reporting without manual formula setup.

**Updated Key Skills:** Oracle DB, SQL, PL/SQL, Schema Design, Python, Excel Automation, GUI Development, Database Design, Relational Data Modeling
