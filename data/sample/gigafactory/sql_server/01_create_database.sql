-- Engineering database for Project Saguaro (fictional). SQL Server 2017 or later.
-- Run in SSMS or Azure Data Studio against your SQL Server / SQL Server Express instance.

IF DB_ID('SaguaroPIIM') IS NULL CREATE DATABASE SaguaroPIIM;
GO
USE SaguaroPIIM;
GO
IF SCHEMA_ID('ref') IS NULL EXEC('CREATE SCHEMA ref');
IF SCHEMA_ID('eng') IS NULL EXEC('CREATE SCHEMA eng');
IF SCHEMA_ID('rpt') IS NULL EXEC('CREATE SCHEMA rpt');
GO

IF OBJECT_ID('eng.submittal') IS NOT NULL DROP TABLE eng.submittal;
IF OBJECT_ID('eng.rfi') IS NOT NULL DROP TABLE eng.rfi;
IF OBJECT_ID('eng.document_revision') IS NOT NULL DROP TABLE eng.document_revision;
IF OBJECT_ID('ref.project_info') IS NOT NULL DROP TABLE ref.project_info;
IF OBJECT_ID('ref.subsystem') IS NOT NULL DROP TABLE ref.subsystem;
IF OBJECT_ID('ref.facility_system') IS NOT NULL DROP TABLE ref.facility_system;
IF OBJECT_ID('ref.company_erp_xref') IS NOT NULL DROP TABLE ref.company_erp_xref;
IF OBJECT_ID('ref.company') IS NOT NULL DROP TABLE ref.company;
IF OBJECT_ID('ref.area') IS NOT NULL DROP TABLE ref.area;
IF OBJECT_ID('ref.discipline') IS NOT NULL DROP TABLE ref.discipline;
GO

CREATE TABLE ref.discipline (
    [Discipline_Code] nvarchar(10) NOT NULL,
    [Discipline] nvarchar(60) NULL,
    [Discipline_Group] nvarchar(20) NULL,
    CONSTRAINT PK_ref_discipline PRIMARY KEY ([Discipline_Code])
);
GO

CREATE TABLE ref.area (
    [Area_Code] nvarchar(20) NOT NULL,
    [Building_Code] nvarchar(10) NULL,
    [Building] nvarchar(50) NULL,
    [Level] nvarchar(30) NULL,
    [Zone] nvarchar(10) NULL,
    [Area_Sort] int NULL,
    CONSTRAINT PK_ref_area PRIMARY KEY ([Area_Code])
);
GO

CREATE TABLE ref.company (
    [Company_ID] nvarchar(10) NOT NULL,
    [Company_Name] nvarchar(60) NULL,
    [Company_Type] nvarchar(30) NULL,
    [Primary_Discipline] nvarchar(10) NULL,
    CONSTRAINT PK_ref_company PRIMARY KEY ([Company_ID])
);
GO

CREATE TABLE ref.company_erp_xref (
    [Company_ID] nvarchar(10) NOT NULL,
    [ERP_Vendor_No] nvarchar(20) NULL,
    CONSTRAINT PK_ref_company_erp_xref PRIMARY KEY ([Company_ID]),
    CONSTRAINT FK_ref_company_erp_xref_Company_ID FOREIGN KEY ([Company_ID]) REFERENCES ref.company(Company_ID)
);
GO

CREATE TABLE ref.facility_system (
    [System_Code] nvarchar(10) NOT NULL,
    [System_Name] nvarchar(60) NULL,
    [Discipline_Code] nvarchar(10) NULL,
    [Building_Code] nvarchar(10) NULL,
    [Priority_Group] int NULL,
    CONSTRAINT PK_ref_facility_system PRIMARY KEY ([System_Code])
);
GO

CREATE TABLE ref.subsystem (
    [Subsystem_Code] nvarchar(20) NOT NULL,
    [System_Code] nvarchar(10) NULL,
    [Subsystem_Name] nvarchar(90) NULL,
    [Area_Code] nvarchar(20) NULL,
    [Planned_MC_Date] date NULL,
    [Planned_Turnover_Date] date NULL,
    [Actual_MC_Date] date NULL,
    [Actual_Turnover_Date] date NULL,
    CONSTRAINT PK_ref_subsystem PRIMARY KEY ([Subsystem_Code]),
    CONSTRAINT FK_ref_subsystem_System_Code FOREIGN KEY ([System_Code]) REFERENCES ref.facility_system(System_Code),
    CONSTRAINT FK_ref_subsystem_Area_Code FOREIGN KEY ([Area_Code]) REFERENCES ref.area(Area_Code)
);
GO

CREATE TABLE ref.project_info (
    [Data_Date] date NULL,
    [Project] nvarchar(60) NULL,
    [Owner] nvarchar(40) NULL
);
GO

CREATE TABLE eng.document_revision (
    [Doc_Rev_ID] int NOT NULL,
    [Doc_No] nvarchar(40) NULL,
    [Revision] nvarchar(10) NULL,
    [Purpose_Code] nvarchar(20) NULL,
    [Discipline_Code] nvarchar(10) NULL,
    [Doc_Type] nvarchar(30) NULL,
    [Area_Code] nvarchar(20) NULL,
    [Originator_ID] nvarchar(10) NULL,
    [Is_Vendor_Doc] nvarchar(10) NULL,
    [Planned_Issue_Date] date NULL,
    [Actual_Issue_Date] date NULL,
    [Review_Due_Date] date NULL,
    [Review_Return_Date] date NULL,
    [Review_Code] nvarchar(10) NULL,
    [Is_Latest_Planned] nvarchar(10) NULL,
    CONSTRAINT PK_eng_document_revision PRIMARY KEY ([Doc_Rev_ID]),
    CONSTRAINT FK_eng_document_revision_Discipline_Code FOREIGN KEY ([Discipline_Code]) REFERENCES ref.discipline(Discipline_Code),
    CONSTRAINT FK_eng_document_revision_Area_Code FOREIGN KEY ([Area_Code]) REFERENCES ref.area(Area_Code),
    CONSTRAINT FK_eng_document_revision_Originator_ID FOREIGN KEY ([Originator_ID]) REFERENCES ref.company(Company_ID)
);
GO

CREATE TABLE eng.rfi (
    [RFI_No] nvarchar(20) NOT NULL,
    [Discipline_Code] nvarchar(10) NULL,
    [Area_Code] nvarchar(20) NULL,
    [Raised_By_ID] nvarchar(10) NULL,
    [Subject_Category] nvarchar(50) NULL,
    [Priority] nvarchar(20) NULL,
    [Date_Raised] date NULL,
    [Date_Required] date NULL,
    [Date_Responded] date NULL,
    [Date_Closed] date NULL,
    [Status] nvarchar(20) NULL,
    [Cost_Impact] nvarchar(10) NULL,
    [Schedule_Impact_Days] int NULL,
    [Linked_Doc_No] nvarchar(40) NULL,
    CONSTRAINT PK_eng_rfi PRIMARY KEY ([RFI_No]),
    CONSTRAINT FK_eng_rfi_Discipline_Code FOREIGN KEY ([Discipline_Code]) REFERENCES ref.discipline(Discipline_Code),
    CONSTRAINT FK_eng_rfi_Area_Code FOREIGN KEY ([Area_Code]) REFERENCES ref.area(Area_Code),
    CONSTRAINT FK_eng_rfi_Raised_By_ID FOREIGN KEY ([Raised_By_ID]) REFERENCES ref.company(Company_ID)
);
GO

CREATE TABLE eng.submittal (
    [Submittal_No] nvarchar(20) NOT NULL,
    [Spec_Section] nvarchar(80) NULL,
    [Discipline_Code] nvarchar(10) NULL,
    [Company_ID] nvarchar(10) NULL,
    [Submittal_Type] nvarchar(30) NULL,
    [Revision] int NULL,
    [Date_Submitted] date NULL,
    [Date_Due] date NULL,
    [Date_Returned] date NULL,
    [Action_Code] nvarchar(40) NULL,
    [Status] nvarchar(20) NULL,
    [Critical_Path] nvarchar(10) NULL,
    CONSTRAINT PK_eng_submittal PRIMARY KEY ([Submittal_No]),
    CONSTRAINT FK_eng_submittal_Discipline_Code FOREIGN KEY ([Discipline_Code]) REFERENCES ref.discipline(Discipline_Code),
    CONSTRAINT FK_eng_submittal_Company_ID FOREIGN KEY ([Company_ID]) REFERENCES ref.company(Company_ID)
);
GO
