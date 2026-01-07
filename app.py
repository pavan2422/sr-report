import streamlit as st
import pandas as pd
import io
from openpyxl import load_workbook
from openpyxl.styles import Font
from openpyxl.formatting.rule import FormulaRule, DataBarRule
from openpyxl.utils import get_column_letter

# --- PAGE CONFIG ---
st.set_page_config(page_title="Payment SR Analysis", layout="wide")

st.title("SR Report Generator - Product Ops")

# 1. HYPERLINK
st.markdown("""
<a href="https://metabase.cashfree.com/question/23625-sr-report-data?merchantid=&start_date=&end_date=" target="_blank" style="font-size:18px; font-weight:bold;">
CLICK HERE TO GET THE DATA FROM METABASE
</a>
""", unsafe_allow_html=True)

st.markdown("---")

# --- PSP MAPPING DICTIONARY ---
PSP_MAP = {
    'abcdicici': 'ABCD (Aditya Birla Capital)', 'abfspay': 'Bajaj Finserv / Markets', 'airtel': 'Airtel Thanks App',
    'allbank': 'Allahabad Bank (Now Indian Bank)', 'amazonpay': 'Amazon Pay', 'andb': 'Andhra Bank (Now Union Bank)',
    'apl': 'Amazon Pay', 'ary': 'Aryavart Bank', 'aubank': 'AU Small Finance Bank', 'axisb': 'Axis Mobile / CRED',
    'axisbank': 'Axis Mobile App', 'axl': 'PhonePe', 'bandhan': 'Bandhan Bank', 'barodampay': 'bob World',
    'barodapay': 'bob World', 'bhim': 'BHIM App', 'boi': 'BHIM Aadhaar / Mobile', 'bombob': 'BHIM BOM',
    'bpaywallet': 'BharatPe', 'bpunity': 'BharatPe', 'cbin': 'Central Bank Mobile', 'centralbank': 'Central Bank Mobile',
    'citi': 'Citibank India (Now Axis)', 'citigold': 'Citibank India (Now Axis)', 'cnrb': 'Candy / Canara ai1',
    'corp': 'Corporation Bank (Now Union Bank)', 'cred': 'CRED', 'csbpay': 'CSB Bank', 'cub': 'CUB mBank Plus',
    'db': 'Deutsche Bank', 'dbs': 'DBS Digibank', 'dhani': 'Dhani App', 'digikhata': 'DigiKhata',
    'dlb': 'Dhanlaxmi Bank', 'dnsbank': 'DNS Bank', 'draxisbank': 'DR', 'drbob': 'DR', 'drcanb': 'DR',
    'drfederal': 'DR', 'drhdfcbank': 'DR', 'dricici': 'DR', 'dridbi': 'DR', 'dridfc': 'DR', 'drindus': 'DR',
    'drkotak': 'DR', 'drpnb': 'DR', 'drsbi': 'DR', 'drubi': 'DR', 'druco': 'DR', 'dryesb': 'DR',
    'ebixcash': 'EbixCash', 'equitas': 'Equitas Mobile', 'equitasbank': 'Equitas Mobile', 'esaf': 'ESAF Mobile',
    'fam': 'FamPay', 'fbl': 'FedMobile / CoinTab', 'federal': 'FedMobile', 'fifederal': 'Fi Money',
    'fincarebank': 'Fincare Mobile', 'finobank': 'FinoPay', 'fkaxis': 'Flipkart', 'freecharge': 'Freecharge',
    'freoicici': 'Freo', 'goaxb': 'Kiwi', 'gwaxis': 'Genwise', 'hdfc': 'HDFC Bank Mobile',
    'hdfcbank': 'HDFC Bank Mobile', 'hsbc': 'HSBC India', 'hsbcbank': 'HSBC India', 'ibl': 'PhonePe',
    'icici': 'iMobile / Pockets', 'idbi': 'IDBI Go Mobile+', 'idfcbank': 'IDFC First Mobile',
    'idfcfirst': 'IDFC First Mobile', 'ikwik': 'MobiKwik', 'imobile': 'iMobile Pay', 'indianbank': 'IndOASIS',
    'indianbk': 'IndOASIS', 'indie': 'INDIE', 'indus': 'IndusMobile', 'indusind': 'IndusMobile',
    'inhdfc': 'Tata Neu', 'iob': 'IOB Mobile', 'janabank': 'Jana Mobile', 'jarunity': 'Jar App',
    'jio': 'MyJio / JioPay', 'jkb': 'J&K Bank', 'jsb': 'Janaseva Bank', 'jupiter': 'Jupiter Money',
    'jupiteraxis': 'Jupiter Money', 'kaypay': 'Kotak (Legacy)', 'kbaxis': 'KreditBee', 'kbl': 'KBL Mobile',
    'kotak': 'Kotak Mobile', 'kotak811': 'Kotak 811', 'kphdfc': 'Kredit.Pe', 'kvb': 'KVB Dlite',
    'liv': 'LivQuik', 'lxaxis': 'LiquiLoans', 'mahb': 'MahaMobile', 'maxaxis': 'Max Life', 'mbk': 'MobiKwik',
    'mbkns': 'MobiKwik', 'mboi': 'Bank of India Mobile', 'mvhdfc': 'Money View', 'naviaxis': 'Navi App',
    'niyoicici': 'Niyo', 'nsdl': 'NSDL Jiffy', 'nye': 'Niyo', 'nyes': 'Niyo', 'obopay': 'Obopay',
    'okaxis': 'Google Pay', 'okhdfcbank': 'Google Pay', 'okicici': 'Google Pay', 'oksbi': 'Google Pay',
    'omni': 'OmniCard', 'oneyes': 'OneCard', 'paytm': 'Paytm', 'paytmwallet': 'Paytm', 'payu': 'PayU',
    'payworld': 'Payworld', 'payzapp': 'PayZapp', 'phonepe': 'PhonePe', 'pinelabs': 'Pine Labs',
    'pingpay': 'Samsung Pay Mini', 'pnb': 'PNB One', 'pnyes': 'PennyDrop', 'pockets': 'Pockets',
    'postbank': 'IPPB Mobile', 'psb': 'PSB UnIC', 'psbank': 'PSB UnIC', 'ptaxis': 'Paytm', 'pthdfc': 'Paytm',
    'ptsbi': 'Paytm', 'ptyes': 'Paytm', 'pz': 'PayZapp', 'pzh': 'PayZapp', 'pzw': 'PayZapp', 'rapl': 'Amazon Pay',
    'razorpay': 'Razorpay', 'rbl': 'RBL MoBank', 'rmrbl': 'Resilient', 'sbi': 'BHIM SBI Pay / Yono',
    'sbmbank': 'SBM Bank', 'scb': 'SC Mobile', 'seyes': 'SalarySe', 'shriramhdfcbank': 'Shriram Finance',
    'sib': 'Mirror+', 'slc': 'Slice', 'slice': 'Slice', 'sliceaxis': 'Slice', 'slicepay': 'Slice',
    'spicepay': 'Spice Money', 'superyes': 'Super.Money', 'suryoday': 'Suryoday Mobile', 'tapicici': 'Tata Neu',
    'tbl': 'Thane Bharat Bank', 'timecosmos': 'TimePay', 'tjsb': 'TJSB Mobile', 'tmb': 'TMB Digilobby',
    'topay': 'ToPay', 'trans': 'Cheq / Transcorp', 'trio': 'Trio', 'ubi': 'Union Bank', 'uboi': 'Union Bank',
    'uco': 'UCO mBanking', 'ujjivan': 'Ujjivan Mobile', 'unionbank': 'Union Bank', 'unionbankofindia': 'Union Bank',
    'unitypay': 'Unity Bank', 'upi': 'BHIM', 'utkarshbank': 'Utkarsh Mobile', 'waaxis': 'WhatsApp Pay',
    'wahdfcbank': 'WhatsApp Pay', 'waicici': 'WhatsApp Pay', 'wasbi': 'WhatsApp Pay', 'yapl': 'Amazon Pay',
    'ybl': 'PhonePe', 'yes': 'Yes Bank', 'yesbank': 'Iris by Yes Bank', 'yescred': 'CRED',
    'yescurie': 'CRED (Curie)', 'yesfam': 'FamPay', 'yesg': 'Groww Pay', 'yesgo': 'Yes Bank',
    'yespay': 'Yes Pay Next', 'yespop': 'POP', 'yestp': 'Third Party App', 'zoicici': 'Zomato', 'ztrbl': 'Zeta'
}

TIER_1_BANKS = ['AXIS BANK', 'HDFC BANK', 'ICICI BANK', 'KOTAK MAHINDRA BANK', 'STATE BANK OF INDIA', 'YES BANK LTD']

# --- HELPER FUNCTIONS ---
def clean_columns(df):
    df.columns = df.columns.str.strip().str.lower().str.replace(" ", "_")
    return df

def compute_sr(data, group_cols, merchant_col, tx_id_col, amount_col, status_col, num_merchants):
    total_volume = len(data)
    group_cols = [c for c in group_cols if c is not None]
    
    cols = [merchant_col] + group_cols + [
        "SR (%)", "SR without User Drops (%)", "Volume", "Success",
        "Unsuccessful Count", "UserDrops", "Total_Value", "GMV", "% of Volume (Global)"
    ]
    if num_merchants > 1:
        cols.append("% of Volume (Per Merchant)")
    
    if total_volume == 0:
        return pd.DataFrame(columns=cols)

    grouped = data.groupby([merchant_col] + group_cols, dropna=False).agg(
        Volume=(tx_id_col, 'count'),
        Success=('is_success', 'sum'),
        UserDrops=('is_userdrop', 'sum'),
        Total_Value=(amount_col, 'sum')
    ).reset_index()

    grouped['Unsuccessful Count'] = grouped['Volume'] - grouped['Success']
    
    gmv_data = data[data['is_success'] == 1].groupby([merchant_col] + group_cols, dropna=False)[amount_col].sum().reset_index(name='GMV')
    grouped = grouped.merge(gmv_data, on=[merchant_col] + group_cols, how='left').fillna({'GMV': 0})

    grouped['SR (%)'] = (grouped['Success'] / grouped['Volume'].replace(0, 1) * 100).round(2)

    df_no_drops = data[data[status_col] != 'USER_DROPPED']
    if not df_no_drops.empty:
        sr_no_drops_grouped = df_no_drops.groupby([merchant_col] + group_cols, dropna=False)
        sr_no_drops_calc = (sr_no_drops_grouped['is_success'].sum() / sr_no_drops_grouped[tx_id_col].count().replace(0, 1) * 100).round(2).reset_index(name='SR without User Drops (%)')
        grouped = grouped.merge(sr_no_drops_calc, on=[merchant_col] + group_cols, how='left')
    else:
        grouped['SR without User Drops (%)'] = 0.0

    grouped['% of Volume (Global)'] = (grouped['Volume'] / total_volume * 100).round(2)

    if num_merchants > 1:
        merchant_totals = data.groupby(merchant_col)[tx_id_col].count().reset_index(name='Merchant_Total')
        grouped = grouped.merge(merchant_totals, on=merchant_col, how='left')
        grouped['% of Volume (Per Merchant)'] = (grouped['Volume'] / grouped['Merchant_Total'].replace(0, 1) * 100).round(2)

    grouped.fillna({'SR without User Drops (%)': 0.0}, inplace=True)
    grouped = grouped[[c for c in cols if c in grouped.columns]]
    return grouped.sort_values(by='Volume', ascending=False)

def compute_mom_change(df, group_cols, merchant_col):
    if df.empty: return df
    group_cols = [c for c in group_cols if c is not None]
    
    df = df.sort_values([merchant_col] + group_cols)
    key_cols = [c for c in group_cols if c != 'Month' and c != 'Week' and c != 'Hour']
    
    result = df.copy()
    for metric in ["Volume", "SR (%)"]:
        if metric in result.columns:
            result[f"{metric} Δ"] = result.groupby([merchant_col] + key_cols)[metric].diff().fillna(0)
            result[f"{metric} % Change"] = (result.groupby([merchant_col] + key_cols)[metric].pct_change().fillna(0) * 100).round(2)
    return result

def apply_formatting(buffer):
    buffer.seek(0)
    wb = load_workbook(buffer)
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        if ws.max_row <= 1: continue
        headers = [cell.value for cell in ws[1]]
        vol_letter = get_column_letter(headers.index("Volume") + 1) if "Volume" in headers else None
        
        for col_idx, header in enumerate(headers, 1):
            col_letter = get_column_letter(col_idx)
            if header in ["SR (%)", "SR without User Drops (%)"] and vol_letter:
                    ws.conditional_formatting.add(f"{col_letter}2:{col_letter}{ws.max_row}", FormulaRule(formula=[f"AND({vol_letter}2>=10,{col_letter}2<50)"], font=Font(color="FF0000")))
                    ws.conditional_formatting.add(f"{col_letter}2:{col_letter}{ws.max_row}", FormulaRule(formula=[f"AND({vol_letter}2>=10,{col_letter}2>90)"], font=Font(color="008000")))
            if "Volume" in str(header) and "%" in str(header) and "Change" not in str(header):
                    ws.conditional_formatting.add(f"{col_letter}2:{col_letter}{ws.max_row}", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=100, color="638EC6"))
            if header == "Unsuccessful Count" and vol_letter:
                    ws.conditional_formatting.add(f"{col_letter}2:{col_letter}{ws.max_row}", FormulaRule(formula=[f"AND({vol_letter}2>=10,{col_letter}2>=20,{col_letter}2>0.5*{vol_letter}2)"], font=Font(color="FF4500")))

        for col in ws.columns:
            max_length = 0
            col_letter = col[0].column_letter
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length: max_length = len(str(cell.value))
                except: pass
            ws.column_dimensions[col_letter].width = min(max_length + 2, 50)
    
    final_output = io.BytesIO()
    wb.save(final_output)
    final_output.seek(0)
    return final_output

# --- MAIN APP ---

uploaded_file = st.file_uploader("Upload CSV File", type=['csv'])

if uploaded_file is not None:
    try:
        df = pd.read_csv(uploaded_file)
        df = clean_columns(df)
        
        # --- CONSTANTS ---
        DATE_COLUMN = 'txtime'
        TX_STATUS_COLUMN = 'txstatus'
        TX_ID_COLUMN = 'transactionid'
        AMOUNT_COLUMN = 'amount'
        MERCHANT_COLUMN = 'merchantid'

        if not all(col in df.columns for col in [DATE_COLUMN, TX_STATUS_COLUMN, TX_ID_COLUMN, MERCHANT_COLUMN]):
            st.error(f"❌ Missing required columns.")
            st.stop()

        # --- PRE-PROCESSING ---
        if AMOUNT_COLUMN in df.columns:
            df[AMOUNT_COLUMN] = df[AMOUNT_COLUMN].astype(str).str.replace(r'[^\d.]', '', regex=True)
            df[AMOUNT_COLUMN] = pd.to_numeric(df[AMOUNT_COLUMN], errors='coerce').fillna(0)
        else: df[AMOUNT_COLUMN] = 0

        df[DATE_COLUMN] = pd.to_datetime(df[DATE_COLUMN], errors='coerce')
        df.dropna(subset=[DATE_COLUMN], inplace=True)
        
        df[TX_STATUS_COLUMN] = df[TX_STATUS_COLUMN].astype(str).str.upper().str.strip()
        for col in ['paymentmode', 'cardtype', 'pg', 'bankname', 'cardcountry']:
            if col in df.columns: df[col] = df[col].astype(str).str.upper().str.strip()

        df['is_success'] = (df[TX_STATUS_COLUMN] == 'SUCCESS').astype(int)
        df['is_userdrop'] = (df[TX_STATUS_COLUMN] == 'USER_DROPPED').astype(int)
        df['Day'] = df[DATE_COLUMN].dt.date
        df['Week'] = df[DATE_COLUMN].dt.to_period('W').astype(str) 
        df['Month'] = df[DATE_COLUMN].dt.to_period('M').astype(str)
        df['Hour'] = df[DATE_COLUMN].dt.floor('H').astype(str) 

        bins = [0, 500, 5000, 25000, 100000, float('inf')]
        labels = ['₹1 – ₹500', '₹501 – ₹5,000', '₹5,001 – ₹25,000', '₹25,001 – ₹1,00,000', '₹1,00,001 and above']
        df['amount_category'] = pd.cut(df[AMOUNT_COLUMN], bins=bins, labels=labels, right=True)

        if 'cardnumber' in df.columns:
             df['upi_handle'] = df['cardnumber'].astype(str).apply(lambda x: x.split('@')[1] if '@' in x else None).str.lower()
             df['psp_app'] = df['upi_handle'].map(PSP_MAP).fillna(df['upi_handle'])

        if 'bankname' in df.columns:
            df['bank_tier'] = df['bankname'].apply(lambda x: 'Tier 1 Bank' if x in TIER_1_BANKS else 'Tier 2 Bank')

        # --- SIDEBAR FILTERS ---
        st.sidebar.header("Data Settings")
        
        share_with_merchant = st.sidebar.checkbox("Share with Merchant (Remove PG Data)", value=False)
        if share_with_merchant:
            if 'pg' in df.columns:
                df = df.drop(columns=['pg'])
                st.sidebar.warning("PG Data removed for external sharing.")

        # --- HOURLY OPTION ---
        st.sidebar.markdown("---")
        st.sidebar.subheader("🕒 Hourly Analysis")
        run_hourly = st.sidebar.checkbox("Get Hourly SR Report")
        hourly_df = None
        if run_hourly:
            min_d, max_d = df['Day'].min(), df['Day'].max()
            st.sidebar.markdown("**Select Date for Hourly Analysis:**")
            hourly_range = st.sidebar.date_input("Date Range (Hourly)", [min_d, max_d], min_value=min_d, max_value=max_d)
            if len(hourly_range) == 2:
                 hourly_df = df[(df['Day'] >= hourly_range[0]) & (df['Day'] <= hourly_range[1])].copy()
            else:
                 st.sidebar.warning("Pick start and end date.")

        st.sidebar.markdown("---")
        st.sidebar.subheader("Filter Data (General)")
        use_date_filter = st.sidebar.checkbox("Filter by Date Range", value=False)
        if use_date_filter:
            min_date = df['Day'].min()
            max_date = df['Day'].max()
            date_range = st.sidebar.date_input("Select Date Range", [min_date, max_date], min_value=min_date, max_value=max_date)
            if len(date_range) == 2:
                df = df[(df['Day'] >= date_range[0]) & (df['Day'] <= date_range[1])]

        all_merchants = sorted(df[MERCHANT_COLUMN].astype(str).unique().tolist())
        selected_merchants = st.sidebar.multiselect("Select Merchants", all_merchants, default=all_merchants)
        if selected_merchants: 
            df = df[df[MERCHANT_COLUMN].astype(str).isin(selected_merchants)]
            if hourly_df is not None:
                hourly_df = hourly_df[hourly_df[MERCHANT_COLUMN].astype(str).isin(selected_merchants)]

        if 'paymentmode' in df.columns:
            all_modes = sorted(df['paymentmode'].unique().tolist())
            selected_modes = st.sidebar.multiselect("Select Payment Modes", all_modes, default=all_modes)
            if selected_modes:
                df = df[df['paymentmode'].isin(selected_modes)]
                if hourly_df is not None: hourly_df = hourly_df[hourly_df['paymentmode'].isin(selected_modes)]

                if any(mode in selected_modes for mode in ['CREDIT_CARD', 'DEBIT_CARD', 'CARD']):
                    st.sidebar.markdown("---")
                    st.sidebar.subheader("💳 Card Filters")
                    if 'cardcountry' in df.columns:
                        df['card_category'] = df['cardcountry'].apply(lambda x: 'DOMESTIC' if x == 'IN' else 'IPG')
                        if hourly_df is not None: hourly_df['card_category'] = hourly_df['cardcountry'].apply(lambda x: 'DOMESTIC' if x == 'IN' else 'IPG')
                        
                        all_cats = sorted(df['card_category'].unique().tolist())
                        selected_cats = st.sidebar.multiselect("Select Card Category (Geo)", all_cats, default=all_cats)
                        if selected_cats: 
                            df = df[df['card_category'].isin(selected_cats)]
                            if hourly_df is not None: hourly_df = hourly_df[hourly_df['card_category'].isin(selected_cats)]
                    
                    if 'cardtype' in df.columns:
                        all_card_types = sorted(df['cardtype'].unique().tolist())
                        selected_card_types = st.sidebar.multiselect("Select Card Network", all_card_types, default=all_card_types)
                        if selected_card_types: 
                            df = df[df['cardtype'].isin(selected_card_types)]
                            if hourly_df is not None: hourly_df = hourly_df[hourly_df['cardtype'].isin(selected_card_types)]

        st.info(f"Ready to analyze **{len(df)}** transactions.")
        
        # --- PROCESSING LOGIC ---
        if st.button("🚀 Run Analysis"):
            with st.spinner('Generating Reports...'):
                
                num_merchants = df[MERCHANT_COLUMN].nunique()
                
                # REPORT CONFIGURATIONS
                report_configs = {
                    'Overview': {'time_col': None, 'sheets': {}},
                    'Daily': {'time_col': 'Day', 'sheets': {}},
                    'Weekly': {'time_col': 'Week', 'sheets': {}},
                    'Monthly': {'time_col': 'Month', 'sheets': {}}
                }

                if run_hourly and hourly_df is not None and not hourly_df.empty:
                    report_configs['Hourly'] = {'time_col': 'Hour', 'sheets': {}, 'data': hourly_df}

                base_breakdowns = {
                    'Paymode': ['paymentmode'],
                    'PG': ['pg'],
                    'Bank': ['bankname'],
                    'Paymode+PG': ['paymentmode', 'pg'],
                    'Paymode+Bank': ['paymentmode', 'bankname']
                }

                # --- MAIN LOOP ---
                generated_buffers = {}
                
                for report_type, config in report_configs.items():
                    current_df = config.get('data', df)
                    time_col = config['time_col']
                    time_group = [time_col] if time_col else []

                    # 1. Base SR
                    if report_type == 'Overview': config['sheets']['SR Overall'] = ([], current_df)
                    else: config['sheets'][f'SR {report_type}'] = (time_group, current_df)
                    
                    # 2. Base Breakdowns
                    for name, group_cols in base_breakdowns.items():
                        if not all(col in current_df.columns for col in group_cols): continue
                        dataset_to_use = current_df if name != 'Card Network' else current_df[current_df['paymentmode'].isin(['CREDIT_CARD', 'DEBIT_CARD'])]
                        config['sheets'][f'SR by {name}'] = (group_cols + time_group, dataset_to_use)

                    # 3. Custom Views
                    config['sheets'][f'SR by Paymode'] = (time_group + ['paymentmode'], current_df)
                    
                    if 'cardtype' in current_df.columns:
                        card_df = current_df[current_df['paymentmode'].isin(['CREDIT_CARD', 'DEBIT_CARD'])]
                        config['sheets'][f'SR by Card Type'] = (time_group + ['paymentmode', 'cardtype'], card_df)
                    
                    if 'upi_handle' in current_df.columns:
                         handle_df = current_df[current_df['upi_handle'].notna()]
                         if not handle_df.empty:
                             config['sheets'][f'SR by UPI Handle'] = (time_group + ['paymentmode', 'upi_handle'], handle_df)
                    
                    if 'psp_app' in current_df.columns:
                         psp_df = current_df[current_df['psp_app'].notna()]
                         if not psp_df.empty:
                             config['sheets'][f'SR by PSP App'] = (time_group + ['paymentmode', 'psp_app'], psp_df)
                    
                    if 'bank_tier' in current_df.columns:
                         # FILTER: NET_BANKING ONLY
                         bank_tier_df = current_df[current_df['paymentmode'] == 'NET_BANKING']
                         if not bank_tier_df.empty:
                            config['sheets'][f'SR by Bank Tier'] = (time_group + ['paymentmode', 'bank_tier'], bank_tier_df)

                    if 'card_category' in current_df.columns:
                        card_cat_df = current_df[current_df['paymentmode'].isin(['CREDIT_CARD', 'DEBIT_CARD'])]
                        if not card_cat_df.empty:
                            config['sheets'][f'SR by Card Geo'] = (time_group + ['paymentmode', 'card_category'], card_cat_df)
                    
                    # 4. EXCEL GENERATION
                    output_buffer = io.BytesIO()
                    with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
                        has_data = False
                        
                        for sheet_name, (cols, dataset) in config['sheets'].items():
                            result = compute_sr(dataset, cols, MERCHANT_COLUMN, TX_ID_COLUMN, AMOUNT_COLUMN, TX_STATUS_COLUMN, num_merchants)
                            
                            if config['time_col'] and config['time_col'] in result.columns:
                                result = result.sort_values(by=config['time_col'], ascending=True)
                            
                            if report_type in ['Monthly', 'Weekly'] and not result.empty:
                                result = compute_mom_change(result, cols, MERCHANT_COLUMN)

                            safe_sheet_name = sheet_name[:31]
                            if not result.empty:
                                result.to_excel(writer, sheet_name=safe_sheet_name, index=False)
                                has_data = True
                        
                        # --- FAILURE ANALYSIS (DETAILED BY TIME) ---
                        failure_data = current_df[current_df[TX_STATUS_COLUMN] != "SUCCESS"]
                        if not failure_data.empty:
                            # Add Time Column to grouping if it exists (Daily/Weekly/Monthly)
                            fail_group_cols = [MERCHANT_COLUMN]
                            if time_col: fail_group_cols.append(time_col)
                            fail_group_cols.extend(['paymentmode', 'txmsg'])
                            
                            fail_summary = failure_data.groupby(fail_group_cols, dropna=False)[TX_ID_COLUMN].count().reset_index(name='Volume')
                            
                            # Sort by time first if applicable, else by Volume
                            sort_cols = [time_col] if time_col else ['Volume']
                            ascending_sort = [True] if time_col else [False]
                            fail_summary = fail_summary.sort_values(by=sort_cols, ascending=ascending_sort)
                            
                            fail_summary.to_excel(writer, sheet_name='Failures Analysis', index=False)
                            has_data = True

                        if not has_data:
                            pd.DataFrame([{"Info": "No data available"}]).to_excel(writer, sheet_name="No Data", index=False)
                    
                    generated_buffers[report_type] = apply_formatting(output_buffer)

                st.session_state['daily_report'] = generated_buffers['Daily']
                st.session_state['weekly_report'] = generated_buffers['Weekly']
                st.session_state['monthly_report'] = generated_buffers['Monthly']
                st.session_state['overview_report'] = generated_buffers['Overview']
                if 'Hourly' in generated_buffers:
                    st.session_state['hourly_report'] = generated_buffers['Hourly']
                
                st.success("✅ Analysis Complete! Download your reports below.")

        # --- DOWNLOAD SECTION ---
        if 'daily_report' in st.session_state:
            st.markdown("### 📥 Download Reports")
            c1, c2 = st.columns(2)
            with c1:
                st.download_button(label="🌍 Download Overview Report (Entire Data)", data=st.session_state['overview_report'], file_name="Overview_SR_Report.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
            
            if 'hourly_report' in st.session_state:
                with c2:
                    st.download_button(label="🕒 Download Hourly Report", data=st.session_state['hourly_report'], file_name="Hourly_SR_Report.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

            c3, c4, c5 = st.columns(3)
            with c3:
                st.download_button(label="📅 Daily Report", data=st.session_state['daily_report'], file_name="Daily_SR_Report.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            with c4:
                st.download_button(label="📆 Weekly Report", data=st.session_state['weekly_report'], file_name="Weekly_SR_Report.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            with c5:
                st.download_button(label="🗓️ Monthly Report", data=st.session_state['monthly_report'], file_name="Monthly_SR_Report.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    except Exception as e:
        st.error(f"Error: {e}")
        st.exception(e)
else:
    st.info("👆 Upload CSV to start.")