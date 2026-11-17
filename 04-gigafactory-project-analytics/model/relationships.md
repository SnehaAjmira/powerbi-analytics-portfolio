# Data model: Project Saguaro (fictional battery cell gigafactory)

Star schema built from six source systems (SQL Server, ERP, P6, completions system, HSE API, SharePoint). Shared keys: Discipline_Code, Area_Code, Company_ID, Subsystem_Code, Tag_No. ERP vendor numbers are translated to Company_ID in Power Query through `ref.company_erp_xref`. All relationships are many-to-one, single direction, from the fact (or child dimension) to the dimension.

## Dimensions

| Table | Key | Notes |
|---|---|---|
| dim_date | Date | DAX calculated table. Mark as date table. |
| dim_discipline | Discipline_Code | 10 engineering / construction disciplines |
| dim_area | Area_Code | Building > level > zone (49 areas) |
| dim_company | Company_ID | General contractor, owner, design partner, subcontractors, suppliers |
| dim_system | System_Code | 30 facility systems (dry room dehumidification, NMP recovery, PCW, MV power, FMCS and so on) |
| dim_subsystem | Subsystem_Code | ~540 subsystems with planned and actual MC / turnover dates |
| dim_cwp | CWP_ID | Construction work packages with budget hours (from P6) |
| project_info | (none) | One row: data date, project name. No relationships. |

## Relationships

| From (many) | To (one) | Active |
|---|---|---|
| fact_document_revisions[Actual_Issue_Date] | dim_date[Date] | Yes |
| fact_document_revisions[Discipline_Code] | dim_discipline[Discipline_Code] | Yes |
| fact_document_revisions[Area_Code] | dim_area[Area_Code] | Yes |
| fact_document_revisions[Originator_ID] | dim_company[Company_ID] | Yes |
| fact_rfi[Date_Raised] | dim_date[Date] | Yes |
| fact_rfi[Discipline_Code] | dim_discipline[Discipline_Code] | Yes |
| fact_rfi[Area_Code] | dim_area[Area_Code] | Yes |
| fact_rfi[Raised_By_ID] | dim_company[Company_ID] | Yes |
| fact_submittals[Date_Submitted] | dim_date[Date] | Yes |
| fact_submittals[Discipline_Code] | dim_discipline[Discipline_Code] | Yes |
| fact_submittals[Company_ID] | dim_company[Company_ID] | Yes |
| fact_procurement[ROS_Date] | dim_date[Date] | Yes |
| fact_procurement[Discipline_Code] | dim_discipline[Discipline_Code] | Yes |
| fact_procurement[Area_Code] | dim_area[Area_Code] | Yes |
| fact_procurement[Supplier_ID] | dim_company[Company_ID] | Yes |
| fact_progress_weekly[Week_Ending] | dim_date[Date] | Yes |
| fact_progress_weekly[CWP_ID] | dim_cwp[CWP_ID] | Yes |
| dim_cwp[Discipline_Code] | dim_discipline[Discipline_Code] | Yes |
| dim_cwp[Area_Code] | dim_area[Area_Code] | Yes |
| dim_cwp[Company_ID] | dim_company[Company_ID] | Yes |
| dim_subsystem[System_Code] | dim_system[System_Code] | Yes |
| dim_subsystem[Area_Code] | dim_area[Area_Code] | Yes |
| fact_tag_register[Subsystem_Code] | dim_subsystem[Subsystem_Code] | Yes |
| fact_tag_register[Discipline_Code] | dim_discipline[Discipline_Code] | Yes |
| fact_tag_register[Install_Date] | dim_date[Date] | Yes |
| fact_itr[Completed_Date] | dim_date[Date] | Yes |
| fact_itr[Subsystem_Code] | dim_subsystem[Subsystem_Code] | Yes |
| fact_itr[Discipline_Code] | dim_discipline[Discipline_Code] | Yes |
| fact_itr[Completed_By_ID] | dim_company[Company_ID] | Yes |
| fact_punch[Date_Raised] | dim_date[Date] | Yes |
| fact_punch[Date_Cleared] | dim_date[Date] | **No** (used by USERELATIONSHIP in [Punch Cleared]) |
| fact_punch[Subsystem_Code] | dim_subsystem[Subsystem_Code] | Yes |
| fact_punch[Discipline_Code] | dim_discipline[Discipline_Code] | Yes |
| fact_punch[Responsible_ID] | dim_company[Company_ID] | Yes |
| fact_safety_weekly[Week_Ending] | dim_date[Date] | Yes |
| fact_safety_weekly[Company_ID] | dim_company[Company_ID] | Yes |
| fact_ncr[Date_Opened] | dim_date[Date] | Yes |
| fact_ncr[Discipline_Code] | dim_discipline[Discipline_Code] | Yes |
| fact_ncr[Area_Code] | dim_area[Area_Code] | Yes |
| fact_ncr[Company_ID] | dim_company[Company_ID] | Yes |

## Design decisions (worth explaining in an interview)

- **No relationship between dim_date and dim_subsystem.** Subsystems carry four dates (planned/actual MC, planned/actual turnover). Linking any of them to dim_date would create a second filter path into fact_itr and fact_punch. The MC and turnover S-curve measures read the axis date with `MAX(dim_date[Date])` and filter the subsystem date columns directly.
- **No relationship between dim_system and dim_discipline.** dim_discipline already filters fact_itr and fact_punch directly. A second path through dim_system > dim_subsystem would make the model ambiguous. System discipline stays as an attribute column.
- **fact_tag_register and fact_punch get area through dim_subsystem**, not their own Area_Code, for the same reason.
- **Role-playing dates are handled in measures.** Each fact has one active date. Planned dates are filtered with `REMOVEFILTERS(dim_date)` plus a column predicate, which avoids a large number of inactive relationships.
- **Actual curves stop at the data date** so S-curves don't flatline into the future.
