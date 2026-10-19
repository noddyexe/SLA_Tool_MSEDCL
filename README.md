# SLA_TOOL_up

**A desktop tool for extracting smart-electricity-meter data and producing SLA compliance reports — without writing SQL by hand.**

---

## What this project does (plain English)

Electricity distribution companies install **smart meters** at consumers' homes. These meters send data back to a central server every day — billing readings, load profiles, alarm events, etc. The company has **SLA (Service Level Agreement)** targets: *"95 % of meters must send billing data within 72 hours"*, *"at least 98 % of daily load profiles must arrive by 24 hours"*, and so on.

Producing a monthly SLA report used to mean:
- running the same SQL queries every month
- pasting the results into Excel by hand
- cross-checking which meters were on holiday, faulty, or removed from service

**This tool automates all of that.** A user logs in, picks a date range, clicks **Run**, and gets a ready-to-share Excel workbook with per-HES summaries, per-SDO summaries, and raw data sheets — including the exclusions (disconnected meters, meter-to-meter swaps, power-off events) that make the SLA numbers *fair*.

---

## What is inside

The application has three modules accessible from a single GUI login screen:

| Module | What it does |
|---|---|
| **SAT Extraction** | Pulls meter-level data from **SLA**, **MDM** and **HES** databases — billing, BLP (block load profile), DLP (daily load profile). Works in **Manual** mode (Excel list of meters) or **Auto** mode (auto-populates the master meter list, no Excel needed). Supports batching, VPN-drop retry, and resume-from-last-batch. |
| **ALARM Report** | Queries the alarm table for the selected alarm type (Last Gasp, First Breath, Over Voltage, Voltage unbalance) and produces a count report per meter. Resumes from last completed batch if interrupted. |
| **SLA Extraction** | The full end-to-end SLA workflow: fetches SATC master, BLP / DLP / billing raw data, Power-Off events, PDC, SM-to-SM, non-operational and RC-DC data — then runs the final summary that produces `SLA_HES_Wise_Final_<MMYY>_<date>.xlsx`. |

All modules are wrapped in a **Tkinter** GUI with live log output, **Stop** buttons, VPN auto-retry, and **Resume** support for long-running fetches.

---

## Technical overview

- **Language:** Python 3.10+
- **GUI:** Tkinter / ttk
- **Databases:** PostgreSQL (four separate DBs: SLA, MDM, HES, RC_DC)
- **DB driver:** `psycopg2`
- **Data processing:** `pandas`
- **Excel I/O:** `openpyxl`, `xlsxwriter`
- **Packaging:** PyInstaller → single-file `.exe`

```
app.py              → login + main menu, routes to the three modules
alarm.py            → Alarm Report module
sat.py              → SAT Extraction module (Manual / Auto)
SLA_SATC/sla.py     → SLA Extraction module (raw fetch + final summary)
config_loader.py    → loads credentials from config.ini
config.example.ini  → template — copy to config.ini and fill in
build_exe.py        → PyInstaller build script
run.bat             → one-click install + build
requirements.txt    → pip dependencies
```

---

## Setup

### First-time setup

1. Install **Python 3.10+** and make sure `python` is on your PATH.
2. Install dependencies:
   ```
   pip install -r requirements.txt
   ```
3. **Copy the config template and fill in your credentials:**
   ```
   copy config.example.ini config.ini
   ```
   Open `config.ini` and enter:
   - the IP, port, database name, user, password for the **MDM / SLA / HES / RC_DC** servers
   - the application login credentials and expiry date

4. Run the app:
   ```
   python app.py
   ```

`config.ini` is listed in `.gitignore` — **it is never committed** to this repository. All database credentials stay on your local machine.

### Building the executable

```
run.bat
```

The output is at `dist\DataProcessingTool.exe`. Place your real `config.ini` **next to the exe** before running it.

---

## Configuration reference

`config.ini` has five sections — one per database, plus login credentials.

```ini
[MDM]
ip = <host>
port = 6412
uid = <user>
pwd = <password>
database = meterdata

[SLA]
ip = <host>
port = 6412
uid = <user>
pwd = <password>
database = unified_sla_db

[HES]
ip = <host>
port = 6412
uid = <user>
pwd = <password>
database = mcl_hes

[RC_DC]
ip = <host>
port = 6412
uid = <user>
pwd = <password>
database = masterdata

[LOGIN]
mcl = <password>
admin = <password>
user = <password>
f = <password>
expiration_date = DDMMYYYY
```

The `expiration_date` blocks the application after that date — used for internal access control.

---

## Outputs produced

| Module | Filename | Content |
|---|---|---|
| SAT Extraction | `BLP_RAW_METER_<module>_<MMYY>_<date>.xlsx`<br>`DLP_RAW_METER_<module>_<MMYY>_<date>.xlsx`<br>`BILLING_RAW_METER_<module>_<MMYY>_<date>.xlsx` | One row per meter (or per day) with the requested metric |
| ALARM Report | `Final_<alarm>_alarm_report_<MMYY>_<date>.xlsx` | One row per meter with alarm count |
| SLA Extraction | `SLA_HES_Wise_Final_<MMYY>_<date>.xlsx` | 6 sheets: HES_Wise_Summary, SDO_Wise_Summary, RC_DC, SLA_BCITS_Raw, SLA_Cyan_Raw, SLA_Crystal_Raw |

Every output is date-stamped, so past runs are never overwritten.

---

## Who this is for

- **Utility operations teams** — running their monthly SLA report
- **MIS / reporting staff** — extracting raw AMI data for ad-hoc analysis
- **Field engineers** — checking which meters have missed their SLA window

It is *not* a customer-facing app. It requires **VPN access** to the utility's internal databases.

---

## Notes

- This tool has only been tested against the **MSEDCL Nagpur** region database schema — column names and table names are specific to that deployment.
- The `ami_loc.xlsx` file maps site-codes to Zone / Circle / Division / Subdivision for the Nagpur region and is bundled with the tool.
- VPN drops during a long fetch are handled automatically: the tool retries the current batch, and if the user presses **Stop** mid-run, progress is saved so it can be resumed.

---

## License

Internal use only — not licensed for redistribution.