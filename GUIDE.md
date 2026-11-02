# Build guide: student dropout early warning (Excel)

**Business question:** can a university spot the students who will drop out using only what it knows after their first semester?

**Result:** one workbook, `student_dropout.xlsx`, with seven sheets (Read me, Data, Courses, Analysis, Early warning, Dashboard, Chart data), a PDF of the dashboard and screenshots. Time: three or four sessions.

Works in **Excel 2016 or later**: every formula uses INDEX/MATCH and nested IF rather than XLOOKUP, IFS or FILTER. Type the `=` of each formula yourself: a pasted leading space makes Excel show the formula as text.

**Data:** Realinho, V., Vieira Martins, M., Machado, J. and Baptista, L. (2021). *Predict Students' Dropout and Academic Success*. UCI Machine Learning Repository, CC BY 4.0, DOI 10.24432/C5MC89. 4,424 students of the Polytechnic Institute of Portalegre, Portugal; one row per student, outcome at the end of the normal course length: Dropout, Enrolled (still studying) or Graduate.

---

## 1. Set up (5 minutes)

1. Unzip `student-dropout-excel` into your `Project Portfolio` folder. The raw file is already in `data/raw/data.csv`, and the course names are in `lookups/course_names.csv`.
2. Open Excel, create a blank workbook and save it straight away in the project folder as `student_dropout.xlsx`.
3. **Do not double-click `data.csv` to open it.** It uses semicolons between columns, so Excel would squash everything into column A. Power Query reads it properly.

## 2. Import the data with Power Query (20 minutes)

1. **Data > Get Data > From File > From Text/CSV**, choose `data/raw/data.csv`.
2. In the preview, set **File Origin: 65001: Unicode (UTF-8)** and **Delimiter: Semicolon**. Click **Transform Data** (not Load).
3. In the Power Query editor, rename the query (right panel, *Name*) to `Students`.
4. **Keep only the columns you need**: **Home > Choose Columns**, tick these 14 and untick the rest:
   `Course`, `Daytime/evening attendance`, `Debtor`, `Tuition fees up to date`, `Gender`, `Scholarship holder`, `Age at enrollment`, `Admission grade`, `Curricular units 1st sem (enrolled)`, `Curricular units 1st sem (approved)`, `Curricular units 1st sem (grade)`, `Curricular units 2nd sem (enrolled)`, `Curricular units 2nd sem (approved)`, `Target`.
5. **Fix the attendance header.** Its name contains a hidden tab character at the end (a real quirk of this file). Double-click the header and retype it as `Daytime/evening attendance`.
6. Rename `Target` to `Outcome`.
7. **Add a student ID**: **Add Column > Index Column > From 1**. Rename it `StudentID` and drag it to the first position.
8. **Check the types** (icon left of each header): `Admission grade` and `Curricular units 1st sem (grade)` = Decimal Number; `Outcome` = Text; everything else = Whole Number.
9. **Home > Close & Load**. Power Query creates a sheet with a table named `Students`. Rename the sheet `Data`.

**Check:** the status bar (or the table) shows **4,424 rows** and 15 columns.

## 3. Course names lookup (10 minutes)

1. **Data > Get Data > From Text/CSV**, choose `lookups/course_names.csv`, delimiter **Comma**, click **Load**. Rename the new sheet `Courses` and make sure the table is named `Courses` (**Table Design > Table Name**).

## 4. Add calculated columns to the Students table (30 minutes)

Click the first empty column to the right of the `Students` table, type the header, then the formula in the first row. Excel fills it down the whole table. Columns added this way stay when the query is refreshed.

| New column | Formula |
| --- | --- |
| `Course name` | `=INDEX(Courses[Course name],MATCH([@Course],Courses[Code],0))` |
| `Dropped` | `=IF([@Outcome]="Dropout",1,0)` |
| `Sem 1 pass rate` | `=IF([@[Curricular units 1st sem (enrolled)]]=0,"",[@[Curricular units 1st sem (approved)]]/[@[Curricular units 1st sem (enrolled)]])` |
| `Sem 1 result` | `=IF([@[Curricular units 1st sem (enrolled)]]=0,"No courses",IF([@[Curricular units 1st sem (approved)]]=0,"None passed",IF([@[Sem 1 pass rate]]<0.5,"Under half",IF([@[Sem 1 pass rate]]<1,"Half or more","All passed"))))` |
| `Age band` | `=IF([@[Age at enrollment]]<=20,"17 to 20",IF([@[Age at enrollment]]<=24,"21 to 24",IF([@[Age at enrollment]]<=34,"25 to 34","35+")))` |
| `Fees up to date` | `=IF([@[Tuition fees up to date]]=1,"Yes","No")` |
| `Scholarship` | `=IF([@[Scholarship holder]]=1,"Yes","No")` |
| `Gender label` | `=IF([@Gender]=1,"Male","Female")` |
| `Risk flag` | `=IF(OR([@[Tuition fees up to date]]=0,AND([@[Curricular units 1st sem (enrolled)]]>0,[@[Curricular units 1st sem (approved)]]<0.5*[@[Curricular units 1st sem (enrolled)]])),"At risk","Not flagged")` |

Format `Sem 1 pass rate` as Percentage, 0 decimals.

In Excel 365 or 2021, XLOOKUP and IFS give the same results.

**What the risk flag means:** a student is flagged after semester one if their tuition fees are not up to date, **or** they passed fewer than half of the first-semester courses they enrolled in. Both are things a university knows by the end of the first semester.

## 5. Analysis sheet: PivotTables (45 minutes)

Create a sheet named `Analysis`. For each table below: click inside the `Students` table, **Insert > PivotTable > Existing Worksheet**, place it on `Analysis`, then:

- **Rows:** the field named below
- **Values:** `StudentID` set to **Count** (rename "Students"), and `Dropped` set to **Average** (rename "Dropout rate", format as Percentage, 1 decimal)

Average of a 1/0 column is a rate: this is the key trick of the sheet.

| PivotTable | Rows |
| --- | --- |
| 1 | `Outcome` (Count only) |
| 2 | `Sem 1 result` |
| 3 | `Fees up to date` |
| 4 | `Scholarship` |
| 5 | `Age band` |
| 6 | `Course name`, sorted by Dropout rate, largest first |
| 7 | `Risk flag` |

Give each PivotTable a title in the cell above it ("Dropout rate by first-semester result", and so on).

## 6. Early warning sheet (30 minutes)

Create a sheet named `Early warning`.

**Summary**: type `Measure` in A1 and `Value` in B1, then the labels below in A2 to A9 and the formulas in B2 to B9:

| Label | Formula |
| --- | --- |
| Students | `=ROWS(Students[StudentID])` |
| Dropped out | `=SUM(Students[Dropped])` |
| Dropout rate | `=AVERAGE(Students[Dropped])` |
| Flagged after semester one | `=COUNTIFS(Students[Risk flag],"At risk")` |
| Share of students flagged | `=B5/B2` |
| Dropout rate among flagged | `=AVERAGEIFS(Students[Dropped],Students[Risk flag],"At risk")` |
| Dropout rate among the rest | `=AVERAGEIFS(Students[Dropped],Students[Risk flag],"Not flagged")` |
| Share of all dropouts caught | `=COUNTIFS(Students[Risk flag],"At risk",Students[Outcome],"Dropout")/B3` |

Format rates as Percentage, 1 decimal.

**Student lookup** (a few rows below):

1. In A12 type `Student ID` and in B12 type `25`. Add a check on B12: **Data > Data Validation > Whole number between 1 and 4424**.
2. In A13 to A17 type `Course`, `Fees up to date`, `Sem 1 result`, `Risk flag`, `Outcome`, and in B13 to B17 use for example
   `=INDEX(Students[Course name],MATCH($B$12,Students[StudentID],0))` (change the last column for each row).
3. Conditional formatting on the `Risk flag` cell: red fill if it equals "At risk".

## 7. Dashboard sheet (1 to 2 hours)

Create a sheet named `Dashboard`. Turn off gridlines (**View > untick Gridlines**).

1. **Title** in A1: "Spotting dropouts after the first semester". Subtitle: "4,424 students, Polytechnic Institute of Portalegre. Source: UCI Machine Learning Repository, CC BY 4.0."
2. **Four KPI cards** (merged cells with a border, big number linked to the Early warning sheet): Dropout rate, Flagged after semester one, Dropout rate among flagged, Share of dropouts caught.
3. **Three charts**: add a sheet `Chart data`. For each chart, build a new PivotTable there from `Students` (Rows: `Sem 1 result`, `Fees up to date` or `Age band`; Values: `Dropped` as Average, named `Dropout rate`), then **Insert > PivotChart** (Clustered Bar for the first two, Clustered Column for age). Separate pivots keep the Analysis tables untouched. Sort the first two Smallest to Largest so the biggest bar is on top. Delete the legend, axis and gridlines, hide the field buttons, add data labels, and give each a title that states the finding, for example "Passed no first-semester course: 92% dropped out". Cut and paste each chart onto `Dashboard`.
4. **Slicers**: click a PivotChart, **PivotChart Analyze > Insert Slicer**, tick `Course name` and `Gender label`. Then **Slicer > Report Connections** and tick the three Chart data pivots, so one slicer filters every chart. The KPI cards come from the Early warning formulas and always show all students.
5. One accent colour for the bars that matter, grey for the rest.
6. **Read me sheet**: question, data source and licence, what each sheet contains, the risk rule, and the limitations below.

## 8. Check your numbers

| Where | Check | Expected |
| --- | --- | --- |
| Data | Rows | 4,424 |
| Analysis 1 | Dropout / Enrolled / Graduate | 1,421 / 794 / 2,209 |
| Early warning | Dropout rate | 32.1% |
| Analysis 2 | Dropout rate: None passed / Under half / Half or more / All passed / No courses | 91.6% / 72.9% / 28.4% / 8.7% / 42.8% |
| Analysis 2 | Students: None passed / Under half / Half or more / All passed / No courses | 538 / 310 / 1,668 / 1,728 / 180 |
| Analysis 3 | Fees up to date: No / Yes | 86.6% (528 students) / 24.7% (3,896) |
| Analysis 4 | Scholarship: No / Yes | 38.7% / 12.2% |
| Analysis 5 | 17 to 20 / 21 to 24 / 25 to 34 / 35+ | 21.2% / 33.7% / 56.5% / 53.7% |
| Analysis 6 | Lowest / highest course (12+ students) | Nursing 15.4% / Biofuel Production Technologies 66.7% (only 12 students); Equinculture 55.3%, Informatics Engineering 54.1% |
| Early warning | Flagged after semester one | 1,115 (25.2% of students) |
| Early warning | Dropout rate among flagged / among the rest | 82.4% / 15.2% |
| Early warning | Share of all dropouts caught | 64.7% (919 of 1,421) |

If a number differs, check the formula for that column first: it is almost always a reference to the wrong column.

## 9. Export and publish

1. Save the workbook.
2. **Dashboard to PDF**: on the Dashboard sheet, **File > Export > Create PDF**, Options: *Active sheet*. Save as `student_dropout_dashboard.pdf`.
3. **Screenshots** (Windows key + Shift + S) of the Dashboard and Early warning sheets, saved in `screenshots` as `dashboard.png` and `early_warning.png`.
4. Send me the screenshots and I will fill in `README.md` with your results.
5. On GitHub, create a public repository `student-dropout-excel` (README off). Description: *Excel early-warning model for student dropout: Power Query, INDEX/MATCH, PivotTables and a dashboard on 4,424 real students.* Upload everything except `~$` temporary files. Topics: `excel`, `power-query`, `pivot-tables`, `dashboard`, `data-analysis`, `education`.

## Points to think about (for your findings and interviews)

- Why measure the rule by "share of dropouts caught" and "dropout rate among flagged" rather than overall accuracy? (Same idea as the churn project.)
- The flag catches 65% of dropouts. Who does it miss, and what would you add to catch more? (Look at age and scholarship.)
- 37 flagged students still graduated. What would it cost to give them extra support anyway?
- Can you conclude that unpaid fees *cause* dropout? What else could explain the link?

## Limitations (for the Read me sheet and README)

- One institution in Portugal; patterns may differ elsewhere.
- "Enrolled" students had not finished at the end of the normal course length; some of them may still drop out, so the dropout rate is a minimum.
- The data shows associations, not causes: unpaid fees may be a sign of money problems that also lead to dropout.
- The rule uses the first semester only; a model could do better, but this one fits on a sticky note and anyone can apply it.
