"""Export DAX measures, calculated columns/tables and Power Query code from a .pbix
into readable text files for GitHub. Local file paths are replaced with a
DataFolder parameter, and auto date tables are skipped."""
import os, re, sys
from pbixray import PBIXRay

pbix, out = sys.argv[1], sys.argv[2]
os.makedirs(os.path.join(out, 'dax'), exist_ok=True)
os.makedirs(os.path.join(out, 'powerquery'), exist_ok=True)
p = PBIXRay(pbix)
skip = lambda t: t.startswith('LocalDateTable') or t.startswith('DateTableTemplate')

lines = ['// Measures exported from the report model.',
         '// Paste into DAX query view and click "Update model with changes" to recreate them.', '', 'DEFINE']
for _, r in p.dax_measures.iterrows():
    expr = str(r.Expression).strip().replace('\r\n', '\n')
    expr = '\n'.join('        ' + l for l in expr.split('\n'))
    lines.append(f"    MEASURE '{r.TableName}'[{r.Name}] =\n{expr}\n")
lines.append('EVALUATE { 1 }')
open(os.path.join(out, 'dax', 'measures.dax'), 'w').write('\n'.join(lines))

cl = ['// Calculated tables and columns (Modeling > New table / New column)', '']
for _, r in p.dax_tables.iterrows():
    if skip(r.TableName): continue
    cl.append(f"// ---- calculated table\n{r.TableName} =\n{str(r.Expression).strip()}\n")
for _, r in p.dax_columns.iterrows():
    if skip(r.TableName): continue
    cl.append(f"// ---- calculated column\n{r.TableName}[{r.ColumnName}] =\n{str(r.Expression).strip()}\n")
open(os.path.join(out, 'dax', 'calculated_tables_and_columns.dax'), 'w').write('\n'.join(cl).replace('\r\n', '\n'))

pq = ['// Power Query (M). Create a text parameter named DataFolder that points to the data folder',
      '// (ending with a backslash), e.g. C:\\PowerBI\\lab\\', '']
for _, r in p.power_query.iterrows():
    m = str(r.Expression).replace('\r\n', '\n')
    m = re.sub(r'File\.Contents\("[^"]*\\([^"\\]+)"\)', r'File.Contents(DataFolder & "\1")', m)
    pq.append(f'// ===== {r.TableName}\n{m.strip()}\n')
open(os.path.join(out, 'powerquery', 'queries.m'), 'w').write('\n'.join(pq))

rel = p.relationships
with open(os.path.join(out, 'dax', 'relationships.txt'), 'w') as f:
    for _, r in rel.iterrows():
        if skip(r.ToTableName) or skip(r.FromTableName): continue
        f.write(f"{r.FromTableName}[{r.FromColumnName}] -> {r.ToTableName}[{r.ToColumnName}]  {r.Cardinality}  "
                f"{'active' if r.IsActive else 'inactive'}  cross-filter: {r.CrossFilteringBehavior}\n")
print(out, len(p.dax_measures), 'measures')
