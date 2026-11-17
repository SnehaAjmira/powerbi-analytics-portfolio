# Re-pointing the three lab reports to the synthetic data

The synthetic files in `data/lab/` use the **same file names, sheet names, Excel table names and column names** as the original sources. No DAX or visual changes are needed for the data to load. Do this on a **copy** of each .pbix, never on the work file.

## 1. Swap the data source (about 5 minutes per report)

1. Copy all files from `data/lab/` into one folder on your PC, for example `C:\PowerBI\lab\`.
2. Open the copied .pbix in Power BI Desktop. If it asks you to sign in to the work account to refresh, cancel.
3. **Home > Transform data > Data source settings**.
4. Select each listed file, click **Change Source**, and browse to the file with the same name in `C:\PowerBI\lab\`:

| Original file | Used by |
|---|---|
| Raw DATA.xlsx | Pareto, Client, Test Performance (Departments) |
| Tests Completed.xlsx | Test Performance |
| Submission Release.xlsx | Test Performance |
| Quality Data.xlsx | Test Performance |
| OosByClient.xlsx | Test Performance |
| OosByTechnician.xlsx | Test Performance |
| Raw Data Canada.xlsx | Client |
| Product Test.xlsx | Client |
| Consulting Activities.xlsx | Client |

5. **Close**, then **Home > Refresh**. The refresh replaces every row of the old data inside the file.
6. Check the row counts in Table view. `Tests Complete` should have about 371,000 rows; `AllData` / `Fact Table (Services)` about 860,000.

Optional, cleaner for GitHub: replace the hard-coded path with a parameter. Create a Text parameter `DataFolder` (e.g. `C:\PowerBI\lab\`) and use the queries in each project's `powerquery/queries.m`, which already read `DataFolder & "file name.xlsx"`.

## 2. Remove what still points to the real company

Refreshing replaces the data, but some things are saved in the report layout itself. These were found by scanning the report definitions:

**All three reports**
- The page background image (`Backdrop.png`) is the company logo banner. On every page: **Format page > Canvas background > Image**, browse to `assets/backdrop-neutral.png`, Image fit = Fit.
- Retitle pages and text boxes so they say "Contoso Analytical Labs" or nothing at all.

**Client report**
- Slicers on *Tests Tooltip, Client Details, Client KPI's, Client Segmentation, Cohorting for Plates and Swabs, QoQ* were saved with a test-facility selection. Click each slicer and choose **Clear selections**, then pick `Contoso - US Central` if you want it pre-filtered.
- The decomposition tree on *Cohorting for Plates and Swabs* has a real client name saved as an expanded node. Collapse it and expand again on the new data.
- Icons (`client.png`, people, revenue, lab, batch, swab kit) are generic, but confirm you have the right to reuse any you downloaded. If one came from a stock site with a watermark, replace it.

**Test Performance report**
- The decomposition tree on *Test Activity* (and its mobile layout) has two real analyst usernames saved as expanded nodes. Collapse and re-expand it on the new data.
- The *NCR* page: check that no text box mentions real investigations.

**Pareto report**
- Only the backdrop image needs replacing.

## 3. Final check before publishing

1. **File > Options > Current file > Data load**: leave as is. Then in **Transform data** confirm every query preview shows Contoso data only.
2. Use the search in Table view for the old company name, an old client name and an old analyst name. Nothing should come back.
3. **File > Save as > Power BI project (.pbip)** into the matching project folder of this repo. Commit the `.pbip`, the `*.Report` folder and the `*.SemanticModel` folder. The `.gitignore` already excludes the local cache.
4. Take screenshots of each page into `screenshots/` for the README.
5. Keep the original work .pbix files out of the repo folder entirely.
