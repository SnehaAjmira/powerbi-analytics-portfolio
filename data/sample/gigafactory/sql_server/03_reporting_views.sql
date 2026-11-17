-- Reporting views for Project Saguaro (fictional gigafactory project).
-- Run after 01_create_database.sql and 02_load_data.sql.
-- Power BI can read these views directly, or you can use them for ad-hoc analysis in SSMS.

USE SaguaroPIIM;
GO

-- RFI status with aging, overdue and on-time flags, as of the project data date
CREATE OR ALTER VIEW rpt.vw_rfi_status AS
SELECT
    r.RFI_No,
    r.Discipline_Code,
    d.Discipline,
    r.Area_Code,
    a.Building,
    r.Raised_By_ID,
    c.Company_Name AS Raised_By,
    r.Priority,
    r.Subject_Category,
    r.Status,
    r.Date_Raised,
    r.Date_Required,
    r.Date_Responded,
    r.Date_Closed,
    p.Data_Date,
    DATEDIFF(day, r.Date_Raised, COALESCE(r.Date_Closed, p.Data_Date)) AS Age_Days,
    CASE
        WHEN DATEDIFF(day, r.Date_Raised, COALESCE(r.Date_Closed, p.Data_Date)) <= 7 THEN '0-7 days'
        WHEN DATEDIFF(day, r.Date_Raised, COALESCE(r.Date_Closed, p.Data_Date)) <= 14 THEN '8-14 days'
        WHEN DATEDIFF(day, r.Date_Raised, COALESCE(r.Date_Closed, p.Data_Date)) <= 30 THEN '15-30 days'
        ELSE 'Over 30 days'
    END AS Age_Bucket,
    CASE WHEN r.Status <> 'Void' AND r.Date_Closed IS NULL THEN 1 ELSE 0 END AS Is_Open,
    CASE WHEN r.Status <> 'Void' AND r.Date_Responded IS NULL AND r.Date_Required < p.Data_Date THEN 1 ELSE 0 END AS Is_Overdue,
    CASE
        WHEN r.Date_Responded IS NULL THEN NULL
        WHEN r.Date_Responded <= r.Date_Required THEN 1
        ELSE 0
    END AS Responded_On_Time,
    DATEDIFF(day, r.Date_Raised, r.Date_Responded) AS Days_To_Respond
FROM eng.rfi AS r
JOIN ref.discipline AS d ON d.Discipline_Code = r.Discipline_Code
JOIN ref.area AS a ON a.Area_Code = r.Area_Code
JOIN ref.company AS c ON c.Company_ID = r.Raised_By_ID
CROSS JOIN ref.project_info AS p;
GO

-- IFC (issued for construction) progress by discipline, planned vs actual to the data date
CREATE OR ALTER VIEW rpt.vw_ifc_progress_by_discipline AS
SELECT
    d.Discipline_Code,
    d.Discipline,
    COUNT(*) AS IFC_Docs,
    SUM(CASE WHEN dr.Planned_Issue_Date <= p.Data_Date THEN 1 ELSE 0 END) AS Planned_To_Date,
    SUM(CASE WHEN dr.Actual_Issue_Date IS NOT NULL THEN 1 ELSE 0 END) AS Issued,
    SUM(CASE WHEN dr.Planned_Issue_Date <= p.Data_Date AND dr.Actual_Issue_Date IS NULL THEN 1 ELSE 0 END) AS Behind_Plan,
    CAST(100.0 * SUM(CASE WHEN dr.Planned_Issue_Date <= p.Data_Date THEN 1 ELSE 0 END) / COUNT(*) AS decimal(5, 1)) AS Pct_Planned,
    CAST(100.0 * SUM(CASE WHEN dr.Actual_Issue_Date IS NOT NULL THEN 1 ELSE 0 END) / COUNT(*) AS decimal(5, 1)) AS Pct_Issued
FROM eng.document_revision AS dr
JOIN ref.discipline AS d ON d.Discipline_Code = dr.Discipline_Code
CROSS JOIN ref.project_info AS p
WHERE dr.Revision = '0'
GROUP BY d.Discipline_Code, d.Discipline;
GO

-- Document reviews sent out and not yet returned
CREATE OR ALTER VIEW rpt.vw_review_backlog AS
SELECT
    dr.Doc_No,
    dr.Revision,
    dr.Purpose_Code,
    dr.Discipline_Code,
    dr.Doc_Type,
    dr.Originator_ID,
    c.Company_Name AS Originator,
    dr.Actual_Issue_Date,
    dr.Review_Due_Date,
    DATEDIFF(day, dr.Review_Due_Date, p.Data_Date) AS Days_Past_Due
FROM eng.document_revision AS dr
JOIN ref.company AS c ON c.Company_ID = dr.Originator_ID
CROSS JOIN ref.project_info AS p
WHERE dr.Review_Due_Date IS NOT NULL
  AND dr.Review_Return_Date IS NULL;
GO

-- Submittal performance by subcontractor
CREATE OR ALTER VIEW rpt.vw_submittal_performance AS
SELECT
    s.Company_ID,
    c.Company_Name,
    COUNT(*) AS Submittals,
    SUM(CASE WHEN s.Status = 'Returned' THEN 1 ELSE 0 END) AS Returned,
    SUM(CASE WHEN s.Status = 'Under Review' AND s.Date_Due < p.Data_Date THEN 1 ELSE 0 END) AS Overdue_With_Reviewer,
    CAST(100.0 * SUM(CASE WHEN s.Action_Code IN ('A - Approved', 'B - Approved as Noted') THEN 1 ELSE 0 END)
         / NULLIF(SUM(CASE WHEN s.Status = 'Returned' THEN 1 ELSE 0 END), 0) AS decimal(5, 1)) AS Pct_Approved,
    CAST(100.0 * SUM(CASE WHEN s.Action_Code IN ('C - Revise & Resubmit', 'D - Rejected') THEN 1 ELSE 0 END)
         / NULLIF(SUM(CASE WHEN s.Status = 'Returned' THEN 1 ELSE 0 END), 0) AS decimal(5, 1)) AS Pct_Resubmit,
    AVG(CAST(DATEDIFF(day, s.Date_Submitted, s.Date_Returned) AS decimal(9, 2))) AS Avg_Turnaround_Days
FROM eng.submittal AS s
JOIN ref.company AS c ON c.Company_ID = s.Company_ID
CROSS JOIN ref.project_info AS p
GROUP BY s.Company_ID, c.Company_Name;
GO

-- Subsystem mechanical completion (MC) and turnover status by system
CREATE OR ALTER VIEW rpt.vw_turnover_by_system AS
SELECT
    fs.System_Code,
    fs.System_Name,
    fs.Priority_Group,
    COUNT(*) AS Subsystems,
    SUM(CASE WHEN ss.Planned_MC_Date <= p.Data_Date THEN 1 ELSE 0 END) AS MC_Planned_To_Date,
    SUM(CASE WHEN ss.Actual_MC_Date IS NOT NULL THEN 1 ELSE 0 END) AS MC_Achieved,
    SUM(CASE WHEN ss.Actual_Turnover_Date IS NOT NULL THEN 1 ELSE 0 END) AS Turned_Over,
    MIN(CASE WHEN ss.Actual_MC_Date IS NULL THEN ss.Planned_MC_Date END) AS Next_Planned_MC
FROM ref.subsystem AS ss
JOIN ref.facility_system AS fs ON fs.System_Code = ss.System_Code
CROSS JOIN ref.project_info AS p
GROUP BY fs.System_Code, fs.System_Name, fs.Priority_Group;
GO
