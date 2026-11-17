-- Loads the CSV files into SaguaroPIIM.
-- 1. Copy the sql_server\csv folder to a path the SQL Server service can read, e.g. C:\PowerBI\gigafactory\sql_server\csv\
-- 2. Set @path below, then run the whole script.

USE SaguaroPIIM;
GO
SET NOCOUNT ON;
DECLARE @path nvarchar(400) = N'C:\PowerBI\gigafactory\sql_server\csv\';
DECLARE @sql nvarchar(max);
DECLARE @n int;

SET @sql = N'BULK INSERT ref.discipline FROM ''' + @path + N'ref_discipline.csv'' WITH (FORMAT = ''CSV'', FIRSTROW = 2, CODEPAGE = ''65001'', TABLOCK);';
EXEC sp_executesql @sql;
SELECT @n = COUNT(*) FROM ref.discipline;
PRINT 'ref.discipline: ' + CAST(@n AS varchar(20)) + ' rows';

SET @sql = N'BULK INSERT ref.area FROM ''' + @path + N'ref_area.csv'' WITH (FORMAT = ''CSV'', FIRSTROW = 2, CODEPAGE = ''65001'', TABLOCK);';
EXEC sp_executesql @sql;
SELECT @n = COUNT(*) FROM ref.area;
PRINT 'ref.area: ' + CAST(@n AS varchar(20)) + ' rows';

SET @sql = N'BULK INSERT ref.company FROM ''' + @path + N'ref_company.csv'' WITH (FORMAT = ''CSV'', FIRSTROW = 2, CODEPAGE = ''65001'', TABLOCK);';
EXEC sp_executesql @sql;
SELECT @n = COUNT(*) FROM ref.company;
PRINT 'ref.company: ' + CAST(@n AS varchar(20)) + ' rows';

SET @sql = N'BULK INSERT ref.company_erp_xref FROM ''' + @path + N'ref_company_erp_xref.csv'' WITH (FORMAT = ''CSV'', FIRSTROW = 2, CODEPAGE = ''65001'', TABLOCK);';
EXEC sp_executesql @sql;
SELECT @n = COUNT(*) FROM ref.company_erp_xref;
PRINT 'ref.company_erp_xref: ' + CAST(@n AS varchar(20)) + ' rows';

SET @sql = N'BULK INSERT ref.facility_system FROM ''' + @path + N'ref_facility_system.csv'' WITH (FORMAT = ''CSV'', FIRSTROW = 2, CODEPAGE = ''65001'', TABLOCK);';
EXEC sp_executesql @sql;
SELECT @n = COUNT(*) FROM ref.facility_system;
PRINT 'ref.facility_system: ' + CAST(@n AS varchar(20)) + ' rows';

SET @sql = N'BULK INSERT ref.subsystem FROM ''' + @path + N'ref_subsystem.csv'' WITH (FORMAT = ''CSV'', FIRSTROW = 2, CODEPAGE = ''65001'', TABLOCK);';
EXEC sp_executesql @sql;
SELECT @n = COUNT(*) FROM ref.subsystem;
PRINT 'ref.subsystem: ' + CAST(@n AS varchar(20)) + ' rows';

SET @sql = N'BULK INSERT ref.project_info FROM ''' + @path + N'ref_project_info.csv'' WITH (FORMAT = ''CSV'', FIRSTROW = 2, CODEPAGE = ''65001'', TABLOCK);';
EXEC sp_executesql @sql;
SELECT @n = COUNT(*) FROM ref.project_info;
PRINT 'ref.project_info: ' + CAST(@n AS varchar(20)) + ' rows';

SET @sql = N'BULK INSERT eng.document_revision FROM ''' + @path + N'eng_document_revision.csv'' WITH (FORMAT = ''CSV'', FIRSTROW = 2, CODEPAGE = ''65001'', TABLOCK);';
EXEC sp_executesql @sql;
SELECT @n = COUNT(*) FROM eng.document_revision;
PRINT 'eng.document_revision: ' + CAST(@n AS varchar(20)) + ' rows';

SET @sql = N'BULK INSERT eng.rfi FROM ''' + @path + N'eng_rfi.csv'' WITH (FORMAT = ''CSV'', FIRSTROW = 2, CODEPAGE = ''65001'', TABLOCK);';
EXEC sp_executesql @sql;
SELECT @n = COUNT(*) FROM eng.rfi;
PRINT 'eng.rfi: ' + CAST(@n AS varchar(20)) + ' rows';

SET @sql = N'BULK INSERT eng.submittal FROM ''' + @path + N'eng_submittal.csv'' WITH (FORMAT = ''CSV'', FIRSTROW = 2, CODEPAGE = ''65001'', TABLOCK);';
EXEC sp_executesql @sql;
SELECT @n = COUNT(*) FROM eng.submittal;
PRINT 'eng.submittal: ' + CAST(@n AS varchar(20)) + ' rows';
