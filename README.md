# Healthcare MIS Dashboard (Excel)

An MIS reporting project for a medical billing (RCM) team. It takes a messy claims dataset, cleans it, builds MIS reports (KPIs, aging, payer, denials, agent productivity) and presents them on a one-page Excel dashboard.

> **Note:** All data is synthetic (dummy) and created for practice. No real patient or company data is used. Report date: **15-Oct-2026**.

dashboard.png

## Scenario

I act as the MIS Executive for a team of 10 AR agents (3 teams). The manager needs to know: how much is outstanding and how old is it, which payers hold it, why claims are denied, and who is behind target.

## Tools and Excel Features

- Pivot Tables and Pivot Charts
- XLOOKUP, SUMIF, AVERAGEIF, COUNTIF, IFS, IF, SUMPRODUCT, UNIQUE
- TRIM, PROPER, Find and Replace (match case), Remove Duplicates
- Conditional formatting and chart formatting

## Dataset (workbook sheets)

| Sheet | Contents |
|---|---|
| Claims_Raw | 1,010 claims: Claim ID, Patient ID, Facility, Payer, DOS, Submission Date, Billed, Paid, Status, Denial Reason, Assigned Agent |
| Agents | 10 agents, team lead, daily target |
| Daily_Productivity | Daily accounts worked, hours and quality % (1 Sep to 15 Oct 2026) |
| Claims_Clean | Cleaned copy with added columns (raw sheet kept untouched) |
| MIS_Report | KPI summary |
| Productivity_Report | Agent productivity vs target |
| Pivots | Pivot tables and chart data |
| MIS_Dashboard | One-page dashboard |

## Step 1: Data check and cleaning

Before cleaning, every column was checked for: type, blanks, duplicates, spelling and spaces, and wrong values (row/column counts, `COUNTBLANK`, `LEN`, filters, `MIN`/`MAX`).

| Issue found | Fix |
|---|---|
| 10 duplicate claims (same Claim ID) | Remove Duplicates on Claim ID (1,010 to 1,000 rows) |
| 15 blank Payer values | Replaced with "Unknown" (checked with `LEN(TRIM())=0`) |
| Extra spaces and mixed capitals in Payer, Status, Agent | `TRIM` and `PROPER`; "UnitedHealthcare" spelling and ALL CAPS payer names fixed with Find and Replace (match case) |
| Status written in different ways | `PROPER(TRIM())` gives Paid, Denied, Pending |

Blank Paid Amount (Pending claims) and blank Denial Reason (Paid and Pending claims) are expected, so they were left as they are.

## Step 2: Calculated columns

```
Age_Days           = ReportDate - Submission Date
Aging_Bucket       = IFS(Age<=30,"0-30", Age<=60,"31-60", Age<=90,"61-90", TRUE,"90+")
Outstanding_Amount = IF(Status="Pending", Billed - Paid, 0)
Denied_Amount      = IF(Status="Denied", Billed, 0)
Team_Lead          = XLOOKUP(Agent, Agents[Agent], Agents[Team Lead])
```

**Design decision:** only Pending claims are treated as outstanding. Denied claims are tracked separately because they need an appeal.

## Step 3: Reports

- **MIS_Report:** Total Claims 1,000, Total Billed 77.77 lakh, Total Collected 33.74 lakh, Total Outstanding 25.35 lakh, Denied Claims 114, Denial Rate 11.4%, Collected Rate 43.4%, Outstanding 90+ 12.87 lakh (50.8%).
- **Productivity_Report:** per agent: days worked, total accounts (`SUMIF`), target total, productivity % and average quality (`AVERAGEIF`). Team productivity = total accounts / total target = 95.8%.
- **Pivots and charts:** outstanding by aging bucket, by payer, denied amount by reason, productivity % by agent.

## Key Insights

1. **50.8% of outstanding (12.87 lakh of 25.35 lakh) is more than 90 days old.**
2. **Medicare holds the largest share of outstanding** (5.90 lakh, about 23%).
3. **Coding Error and Patient Not Eligible make up about 45% of denied amount,** and both can be prevented before submission.
4. **Team productivity is 95.8% of target.** Amit V is lowest at 91.3% while quality is similar across agents (92.8 to 94.6), so the gap is speed, not accuracy.

## Recommended Actions

- Work the 90+ day bucket first.
- Add pre-submission checks for coding errors and eligibility.
- Review Amit V's workload and process with the team lead.

## A mistake I caught (lesson learned)

After cleaning, I validated the Denial Reason column against the source using an `XLOOKUP` on Claim ID. 201 rows did not match because values had shifted when pasted while a filter was on. I rebuilt the column with a formula (not a paste), so every Claim ID gets its own reason. Lesson: validate cleaned data against the source, and avoid pasting into filtered ranges.

## Files

- `Healthcare_MIS_Dashboard.xlsx`: complete workbook
- `images/`: dashboard screenshot

## Author

Nitish Arya
