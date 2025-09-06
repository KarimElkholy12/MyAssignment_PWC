import pandas as pd
import re
import os

# Check if file exists
if not os.path.exists('Employees_Cleaned.xlsx'):
    print("Error: Employees_Cleaned.xlsx not found in current directory")
    exit(1)

# Read the Excel file
try:
    df = pd.read_excel('Employees_Cleaned.xlsx')
except Exception as e:
    print(f"Error reading Excel file: {e}")
    exit(1)

# Check if required columns exist
required_columns = ['Company', 'Email']
missing_columns = [col for col in required_columns if col not in df.columns]
if missing_columns:
    print(f"Error: Missing required columns: {missing_columns}")
    exit(1)

# Check if dataframe is empty
if df.empty:
    print("Error: The Excel file contains no data")
    exit(1)

# 1. Count unique companies
unique_companies = df['Company'].nunique()

# 2. Find top 5 most common email domains
def extract_domain(email):
    if pd.isna(email):
        return None
    match = re.search(r'@([^\s]+)', str(email))
    return match.group(1) if match else None

df['Email_Domain'] = df['Email'].apply(extract_domain)
domain_counts = df['Email_Domain'].value_counts()
top_5_domains = domain_counts.head(5)

# 3. Count employees per company
# Handle NaN/None company names
df_clean_company = df.dropna(subset=['Company'])
if df_clean_company.empty:
    print("Warning: No valid company names found")
    employees_per_company = pd.Series()
else:
    employees_per_company = df_clean_company['Company'].value_counts().sort_index()

# Create summary report
summary_data = []

# Add unique companies count
summary_data.append({
    'Metric': 'Unique Companies',
    'Value': unique_companies,
    'Details': ''
})

# Add top email domains (up to 5)
if len(top_5_domains) > 0:
    for i, (domain, count) in enumerate(top_5_domains.items(), 1):
        summary_data.append({
            'Metric': f'Top Email Domain #{i}',
            'Value': domain,
            'Details': f'{count} employees'
        })
    
    # If less than 5 domains exist, fill remaining slots
    if len(top_5_domains) < 5:
        for i in range(len(top_5_domains) + 1, 6):
            summary_data.append({
                'Metric': f'Top Email Domain #{i}',
                'Value': f'only {len(top_5_domains)} email domains found',
                'Details': ''
            })
else:
    # No email domains found
    for i in range(1, 6):
        summary_data.append({
            'Metric': f'Top Email Domain #{i}',
            'Value': 'No email domains found',
            'Details': ''
        })

# Add employee counts per company
if len(employees_per_company) > 0:
    for company, count in employees_per_company.items():
        summary_data.append({
            'Metric': 'Employees per Company',
            'Value': company,
            'Details': f'{count} employees'
        })
else:
    summary_data.append({
        'Metric': 'Employees per Company',
        'Value': 'No companies found',
        'Details': '0 employees'
    })

# Create DataFrame and save to CSV
summary_df = pd.DataFrame(summary_data)
summary_df.to_csv('Summary_Report.csv', index=False)

print(f"Analysis Complete!")
print(f"Number of unique companies: {unique_companies}")
if len(top_5_domains) > 0:
    print(f"\nTop {min(5, len(top_5_domains))} email domains:")
    for domain, count in top_5_domains.items():
        print(f"  {domain}: {count} employees")
    if len(top_5_domains) < 5:
        print(f"  (Note: Only {len(top_5_domains)} unique domains found in dataset)")
else:
    print("\nNo email domains found in dataset")

print(f"\nEmployees per company:")
if len(employees_per_company) > 0:
    for company, count in employees_per_company.items():
        print(f"  {company}: {count} employees")
else:
    print("  No companies found with employees")
print(f"\nResults saved to Summary_Report.csv")