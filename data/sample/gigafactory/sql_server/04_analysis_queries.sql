-- Analysis queries for Project Saguaro (fictional gigafactory project).
-- Each query answers one question a project manager or discipline lead would ask.
-- Run them one at a time in SSMS after the database and views are created.

USE SaguaroPIIM;
GO

-- Q1. Is the RFI backlog growing? Raised vs closed per month, with average response time.
WITH raised AS (
    SELECT DATEFROMPARTS(YEAR(Date_Raised), MONTH(Date_Raised), 1) AS Month_Start,
           COUNT(*) AS Raised,
           AVG(CAST(DATEDIFF(day, Date_Raised, Date_Responded) AS decimal(9, 2))) AS Avg_Days_To_Respond
    FROM eng.rfi
    WHERE Status <> 'Void'
    GROUP BY DATEFROMPARTS(YEAR(Date_Raised), MONTH(Date_Raised), 1)
),
closed AS (
    SELECT DATEFROMPARTS(YEAR(Date_Closed), MONTH(Date_Closed), 1) AS Month_Start, COUNT(*) AS Closed
    FROM eng.rfi
    WHERE Date_Closed IS NOT NULL
    GROUP BY DATEFROMPARTS(YEAR(Date_Closed), MONTH(Date_Closed), 1)
)
SELECT r.Month_Start,
       r.Raised,
       COALESCE(c.Closed, 0) AS Closed,
       SUM(r.Raised - COALESCE(c.Closed, 0)) OVER (ORDER BY r.Month_Start) AS Net_Open_Running,
       r.Avg_Days_To_Respond
FROM raised AS r
LEFT JOIN closed AS c ON c.Month_Start = r.Month_Start
ORDER BY r.Month_Start;

-- Q2. Which subcontractors have the most overdue RFIs, and how old are they?
SELECT TOP 10
       Raised_By,
       SUM(Is_Overdue) AS Overdue_RFIs,
       SUM(Is_Open) AS Open_RFIs,
       AVG(CASE WHEN Is_Open = 1 THEN CAST(Age_Days AS decimal(9, 2)) END) AS Avg_Age_Open
FROM rpt.vw_rfi_status
GROUP BY Raised_By
ORDER BY Overdue_RFIs DESC;

-- Q3. Which disciplines are furthest behind on IFC drawings?
SELECT Discipline, IFC_Docs, Planned_To_Date, Issued, Behind_Plan,
       Pct_Planned, Pct_Issued, Pct_Issued - Pct_Planned AS Variance_Pts
FROM rpt.vw_ifc_progress_by_discipline
ORDER BY Variance_Pts;

-- Q4. Engineering rework: documents re-issued for construction the most times.
SELECT TOP 20
       Doc_No, Discipline_Code, Doc_Type,
       COUNT(*) AS IFC_Revisions_Issued
FROM eng.document_revision
WHERE Purpose_Code = 'IFC' AND Actual_Issue_Date IS NOT NULL
GROUP BY Doc_No, Discipline_Code, Doc_Type
HAVING COUNT(*) > 2
ORDER BY IFC_Revisions_Issued DESC, Doc_No;

-- Q5. First-pass approval rate by document originator (Rev A returned Code 1 or 2).
SELECT c.Company_Name AS Originator,
       COUNT(*) AS First_Submissions_Returned,
       CAST(100.0 * SUM(CASE WHEN dr.Review_Code IN ('1', '2') THEN 1 ELSE 0 END) / COUNT(*) AS decimal(5, 1)) AS First_Pass_Pct
FROM eng.document_revision AS dr
JOIN ref.company AS c ON c.Company_ID = dr.Originator_ID
WHERE dr.Revision = 'A' AND dr.Review_Code IS NOT NULL
GROUP BY c.Company_Name
HAVING COUNT(*) >= 20
ORDER BY First_Pass_Pct;

-- Q6. RFI density: RFIs per 100 IFC documents by building. High density points to design quality issues.
WITH rfis AS (
    SELECT a.Building, COUNT(*) AS RFIs
    FROM eng.rfi AS r JOIN ref.area AS a ON a.Area_Code = r.Area_Code
    WHERE r.Status <> 'Void'
    GROUP BY a.Building
),
docs AS (
    SELECT a.Building, COUNT(*) AS IFC_Docs
    FROM eng.document_revision AS dr JOIN ref.area AS a ON a.Area_Code = dr.Area_Code
    WHERE dr.Revision = '0'
    GROUP BY a.Building
)
SELECT d.Building, d.IFC_Docs, COALESCE(r.RFIs, 0) AS RFIs,
       CAST(100.0 * COALESCE(r.RFIs, 0) / d.IFC_Docs AS decimal(6, 1)) AS RFIs_Per_100_Docs
FROM docs AS d
LEFT JOIN rfis AS r ON r.Building = d.Building
ORDER BY RFIs_Per_100_Docs DESC;

-- Q7. Cumulative IFC curve: planned vs actual by month (window functions).
WITH months AS (
    SELECT DATEFROMPARTS(YEAR(Planned_Issue_Date), MONTH(Planned_Issue_Date), 1) AS Month_Start,
           COUNT(*) AS Planned, 0 AS Actual
    FROM eng.document_revision WHERE Revision = '0'
    GROUP BY DATEFROMPARTS(YEAR(Planned_Issue_Date), MONTH(Planned_Issue_Date), 1)
    UNION ALL
    SELECT DATEFROMPARTS(YEAR(Actual_Issue_Date), MONTH(Actual_Issue_Date), 1), 0, COUNT(*)
    FROM eng.document_revision WHERE Revision = '0' AND Actual_Issue_Date IS NOT NULL
    GROUP BY DATEFROMPARTS(YEAR(Actual_Issue_Date), MONTH(Actual_Issue_Date), 1)
),
m AS (
    SELECT Month_Start, SUM(Planned) AS Planned, SUM(Actual) AS Actual
    FROM months GROUP BY Month_Start
)
SELECT Month_Start,
       SUM(Planned) OVER (ORDER BY Month_Start) AS Planned_Cum,
       SUM(Actual) OVER (ORDER BY Month_Start) AS Actual_Cum
FROM m
ORDER BY Month_Start;

-- Q8. Turnover risk: priority systems with subsystems past planned MC that are not complete.
SELECT System_Code, System_Name, Priority_Group, Subsystems,
       MC_Planned_To_Date, MC_Achieved, MC_Planned_To_Date - MC_Achieved AS MC_Behind,
       Next_Planned_MC
FROM rpt.vw_turnover_by_system
WHERE MC_Planned_To_Date > MC_Achieved
ORDER BY Priority_Group, MC_Behind DESC;

-- Q9. Submittal bottlenecks: subcontractors with low approval rates and slow turnaround.
SELECT Company_Name, Submittals, Returned, Overdue_With_Reviewer, Pct_Approved, Pct_Resubmit, Avg_Turnaround_Days
FROM rpt.vw_submittal_performance
ORDER BY Pct_Resubmit DESC;
