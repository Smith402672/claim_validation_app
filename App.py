import pandas as pd

# --- Helper: Clean Serial Number ---
def clean_serial(s):
    return str(s).strip().upper().replace('\xa0', '').replace('\t', '')

# --- Load input files ---
partner_df = pd.read_excel("Partner_Claim_File.xlsx")
promo_df = pd.read_excel("Promotion_Policy.xlsx")
sales_df = pd.read_excel("Sales_Master.xlsx")
billing_df = pd.read_excel("Billing_Price.xlsx")
claimed_df = pd.read_excel("Previously_Claimed.xlsx", usecols=['Serial Number', 'Month'])
install_df = pd.read_excel("Installation.xlsx")

# --- Normalize Installation columns ---
install_df.columns = install_df.columns.str.strip().str.lower().str.replace(' ', '_')
if 'serial_number' not in install_df.columns or 'installation_date' not in install_df.columns:
    raise ValueError("Installation file must contain 'serial_number' and 'installation_date' columns.")
install_df['serial_number'] = install_df['serial_number'].apply(clean_serial)
install_df['installation_date'] = pd.to_datetime(install_df['installation_date'], errors='coerce', dayfirst=True)

# --- Normalize Sales Master ---
sales_df.columns = ['Serial Number', 'Invoice Number', 'Invoice Date', 'Unused1', 'Unused2', 'Customer Name', 'Model', 'Unused3']
sales_df['Serial Number'] = sales_df['Serial Number'].apply(clean_serial)
sales_df['Invoice Date'] = pd.to_datetime(sales_df['Invoice Date'], errors='coerce', dayfirst=True).dt.date

# --- Normalize Partner Claim File ---
partner_df['Serial Number'] = partner_df['Serial Number'].apply(clean_serial)

# --- Merge Sales Master into Partner Claim File ---
partner_df = partner_df.merge(
    sales_df[['Serial Number', 'Customer Name', 'Invoice Number', 'Invoice Date', 'Model']],
    on='Serial Number',
    how='left'
)

# --- Normalize Promotion Policy and merge Promo NLC ---
promo_df['Model No'] = promo_df['Model No'].astype(str).str.strip().str.upper()
partner_df['Model No'] = partner_df.get('Model No', partner_df['Model']).astype(str).str.strip().str.upper()
promo_lookup = promo_df[['Model No', 'Promo NLC']].drop_duplicates()
partner_df = partner_df.merge(promo_lookup, on='Model No', how='left')

# --- Normalize Billing Price file ---
billing_df.columns = billing_df.columns.str.strip()
billing_df['Customer Name'] = billing_df['Customer Name'].astype(str).str.strip()
billing_df['Invoice Number'] = billing_df['Invoice Number'].astype(str).str.strip()
billing_df['Model'] = billing_df['Model'].astype(str).str.strip()
billing_df['Billing Price'] = pd.to_numeric(billing_df['Billing Price'], errors='coerce')

# --- Normalize keys in partner_df ---
partner_df['Customer Name'] = partner_df['Customer Name'].astype(str).str.strip()
partner_df['Invoice Number'] = partner_df['Invoice Number'].astype(str).str.strip()
partner_df['Model'] = partner_df['Model'].astype(str).str.strip()

# --- SUMIFS-style Billing Price lookup (without Invoice Date) ---
def sumifs_billing_price(customer, invoice_no, model):
    if pd.isnull(customer) or pd.isnull(invoice_no) or pd.isnull(model):
        return None
    match_rows = billing_df[
        (billing_df['Customer Name'] == customer) &
        (billing_df['Invoice Number'] == invoice_no) &
        (billing_df['Model'] == model)
    ]
    return match_rows['Billing Price'].sum()

partner_df['Billing Price'] = partner_df.apply(
    lambda row: sumifs_billing_price(row['Customer Name'], row['Invoice Number'], row['Model']),
    axis=1
)

# --- Initial Support Calculation ---
partner_df['Support'] = partner_df['Billing Price'] - partner_df['Promo NLC']

# --- Merge Previously Claimed Month Info ---
claimed_df['Serial Number'] = claimed_df['Serial Number'].apply(clean_serial)
partner_df = partner_df.merge(claimed_df, on='Serial Number', how='left')

# --- Merge Installation Date ---
partner_df = partner_df.merge(
    install_df.rename(columns={'serial_number': 'Serial Number'}),
    on='Serial Number',
    how='left'
)

# --- Extract Claimed Month from Invoice Date ---
partner_df['Claimed Month'] = pd.to_datetime(partner_df['Invoice Date'], errors='coerce').dt.to_period('M')
partner_df['Install Month'] = partner_df['installation_date'].dt.to_period('M')

# --- Apply Remark Logic ---
def generate_remark(row):
    if pd.notnull(row['Month']):
        return f"Already claimed in {row['Month']}"
    elif pd.notnull(row['Support']) and row['Support'] < 0:
        return "NLC is greater than billing price"
    elif pd.notnull(row['Install Month']) and row['Install Month'] < row['Claimed Month']:
        return f"Installation done in {row['Install Month'].strftime('%b-%Y')}"
    else:
        return "Eligible"

partner_df['Remark'] = partner_df.apply(generate_remark, axis=1)

# --- Override Support if Already Claimed or Negative or Installation mismatch ---
partner_df.loc[partner_df['Remark'].str.startswith("Already claimed"), 'Support'] = 0
partner_df.loc[partner_df['Remark'] == "NLC is greater than billing price", 'Support'] = 0
partner_df.loc[partner_df['Remark'].str.startswith("Installation done in"), 'Support'] = 0

# --- Export Final Output ---
partner_df.to_excel("Validated_Claims_Output.xlsx", index=False)
print("✅ Claim validation completed using SUMIFS logic without Invoice Date. Output saved to Validated_Claims_Output.xlsx")