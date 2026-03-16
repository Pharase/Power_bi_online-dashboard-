import pandas as pd
from tkinter import Tk
import glob
import datetime
from datetime import datetime
from tkinter.filedialog import askopenfilename

def select_excel_file(name):
    root = Tk()
    root.title(f"Select specific file : {name}")
    root.withdraw()
    print(f"Select the Excel file: {name}")
    file_path = askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    if not file_path:
        print("No file selected.")
        exit()
    return file_path

def get_latest_file(directory, file_pattern):
    files = glob.glob(os.path.join(directory, file_pattern))
    if not files:
        return None
    return max(files, key=os.path.getmtime)  # Get the most recently modified file
    
REPORT_DIRS = {
        "Summary Report": (r"data\report\Phone_report", "Phone_report_*bi.xlsx")
    }

# Main excel for BI dashboard
main_report = select_excel_file("Main file")
main_df = pd.read_excel(main_report,sheet_name='Sheet1')
print(main_report)

# Phone report _ bi
# Get the latest files
ATTACHMENT_PATHS = [get_latest_file(dir, pattern) for dir, pattern in REPORT_DIRS.values()]
ATTACHMENT_PATHS = [file for file in ATTACHMENT_PATHS if file]  # Remove None values

# Attach Files
if ATTACHMENT_PATHS:
    for file_path in ATTACHMENT_PATHS:
        with open(file_path, "rb") as file:
            file_data = file.read()
            file_name = os.path.basename(file_path)
        print(f"Check with latest phone report: {file_name}")
else:
    print("No valid Phone report files found.")
    
new_df = pd.read_excel(file_path,sheet_name='Sheet1')
print(new_report)

# Step 2: Rename new_df columns
new_df.rename(columns={
    new_df.columns[20]: 'map_debtor',
    new_df.columns[26]: 'map_oa',
}, inplace=True)

# Step 3: Add missing columns if not exist
final_cols = [
    'map_debtor', 'pamoareport_resultcode_id', 'map_result',
    'pamoareport_reportdate', 'project_portname', 'pamoareport_ppdate',
    'pamoareport_pddate', 'pamoareport_followupdate', 'created_at',
    'AQ_date', 'start_date', 'end_date', 'map_oa', 'fix_date'
]

for col in final_cols:
    if col not in new_df.columns:
        new_df[col] = ""

# Step 4: Reorder
new_df = new_df[final_cols]
new_df['year'] = new_df['fix_date'].dt.year
new_df['month'] = new_df['fix_date'].dt.month
new_df['day'] = new_df['fix_date'].dt.day
new_df['unique'] = (
    new_df['map_debtor'].astype(str) + "-" +
    new_df['year'].astype(str) + "-" +
    new_df['month'].astype(str) + "-" +
    new_df['day'].astype(str)
)

# Step 5: Concat & deduplicate
combined_df = pd.concat([main_df, new_df], ignore_index=True)

# Optional: Drop duplicates based on key columns (adjust as needed)
combined_df.drop_duplicates(
    subset=final_cols,
    keep='last',
    inplace=True
)

# Step 6: Save back to Excel
combined_df.reset_index(drop=True, inplace=True)
combined_df.to_excel(main_report, index=False, sheet_name="Sheet1")

# Step 7: Create sorted unique values
summary_info = pd.DataFrame()
summary_info['A'] = sorted(combined_df['unique'].dropna().unique())

# Step 8: Split 'A' into components
summary_info[['B', 'D', 'E', 'F']] = summary_info['A'].str.split('-', expand=True)

# Step 9: Merge to get 'map_oa' (Column C)
# Drop duplicates to avoid reindexing error
map_oa_lookup = combined_df.drop_duplicates(subset='unique')[['unique', 'map_oa']]
summary_info = summary_info.merge(map_oa_lookup, how='left', left_on='A', right_on='unique')
summary_info.rename(columns={'map_oa': 'C'}, inplace=True)
summary_info.drop(columns=['unique'], inplace=True)

# Step 10: Create Column G
summary_info['G'] = summary_info['B'] + '-' + summary_info['D'] + '-' + summary_info['E']

# Step 11: Count occurrences in G to create H
summary_info['H'] = summary_info['G'].map(summary_info['G'].value_counts())

summary_info = summary_info[['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H']]

summary_info.columns = [
    'map_debtor',            # A
    'pamcode',               # B
    'OA',                    # C
    'year',                  # D
    'month',                 # E
    'day',                   # F
    'uniqe_text',            # G
    'working_day_per_month'  # H
]

# Summary report
summary_pay = select_excel_file("Payment transaction file")
summary_pay_df = pd.read_excel(summary_pay, sheet_name='transaction_record', dtype=str)

# Step 8: Create Sheet2
transaction_df = pd.DataFrame()
transaction_df['A'] = summary_pay_df['pam_code']
transaction_df['B'] = summary_pay_df['code']
transaction_df['C'] = summary_pay_df['TR_Date']
transaction_df['D'] = pd.to_datetime(summary_pay_df['Pay_Date'], errors='coerce')

# Extract year, month, day safely
transaction_df['E'] = transaction_df['D'].dt.year.astype('Int64').astype(str)
transaction_df['F'] = transaction_df['D'].dt.month.astype('Int64').astype(str)
transaction_df['G'] = transaction_df['D'].dt.day.astype('Int64').astype(str)

# Construct 'H' safely and format numbers as integers (no .0)
transaction_df['H'] = (
    transaction_df['A'].astype(str) + '-' +
    transaction_df['E'].fillna(0).astype(str) + '-' +
    transaction_df['F'].fillna(0).astype(str)
)

transaction_df.columns = [
    'pam_code',           
    'code',               
    'TR_date',               
    'Pay_Date',               
    'year',               
    'month',                 
    'day',          
    'uniqe_text' 
]

# Simulate Excel logic: paid / not paid
summary_info['payment'] = summary_info['uniqe_text'].apply(
    lambda x: 'paid' if x in transaction_df['uniqe_text'].values else 'not paid'
)

# Code column only if paid
code_lookup = transaction_df[['uniqe_text', 'code']].drop_duplicates().set_index('uniqe_text')['code']
summary_info['Code'] = summary_info.apply(
    lambda row: code_lookup.get(row['uniqe_text']) if row['payment'] == 'paid' else '-', axis=1
)

# Save output
with pd.ExcelWriter(main_report, engine='xlsxwriter') as writer:
    combined_df.to_excel(writer, sheet_name="Sheet1", index=False)
    summary_info.to_excel(writer, sheet_name="Sheet2", index=False)
    transaction_df.to_excel(writer, sheet_name="transaction_record", index=False)

print(f"Updated and saved back to: {main_report}")
