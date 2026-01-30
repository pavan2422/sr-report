import streamlit as st
import pandas as pd
import io
import gc
import plotly.graph_objects as go
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

# --- CONSTANTS & MAPPINGS ---
TIER_1_BANKS = ['AXIS BANK', 'HDFC BANK', 'ICICI BANK', 'KOTAK MAHINDRA BANK', 'STATE BANK OF INDIA', 'YES BANK LTD']
CARD_MODES = ['CREDIT_CARD', 'DEBIT_CARD', 'CARD', 'PREPAID_CARD', 'CREDIT_CARD_EMI']

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

@st.cache_data
def preprocess_data(df):
    cols_to_upper = ['paymentmode', 'txstatus', 'bankname', 'cardtype', 'cardcountry', 'pg', 'txmsg']
    for col in cols_to_upper:
        if col in df.columns:
            df[col] = df[col].astype(str).str.upper().str.strip().replace('NAN', None)

    df['txtime'] = pd.to_datetime(df['txtime'], errors='coerce', format='mixed')
    df.dropna(subset=['txtime'], inplace=True)
    df['Day'] = df['txtime'].dt.date
    df['Month'] = df['txtime'].dt.to_period('M').astype(str)
    df['Week'] = df['txtime'].dt.to_period('W').astype(str)
    df['Hour'] = df['txtime'].dt.floor('h').astype(str)

    # Format Date for Display (e.g. 01-Jan)
    df['Display_Date'] = df['txtime'].dt.strftime('%d-%b')

    if 'amount' in df.columns:
        df['amount'] = pd.to_numeric(df['amount'].astype(str).str.replace(r'[^\d.]', '', regex=True), errors='coerce').fillna(0)
    else: df['amount'] = 0
    
    # --- AMOUNT CATEGORY LOGIC ---
    def get_amt_cat(amt):
        if amt <= 1000: return '0-1k'
        elif amt <= 10000: return '1k-10k'
        elif amt <= 50000: return '10k-50k'
        elif amt <= 100000: return '50k-1L'
        elif amt <= 200000: return '1L-2L'
        else: return '>2L'
    
    df['amount_category'] = df['amount'].apply(get_amt_cat)
    
    df['is_success'] = (df['txstatus'] == 'SUCCESS').astype(int)
    df['is_userdrop'] = (df['txstatus'] == 'USER_DROPPED').astype(int)
    df['is_failed'] = (df['txstatus'] == 'FAILED').astype(int)
    
    # Logic for Incomplete/Pending
    df['is_incomplete'] = df['txstatus'].isin(['PENDING', 'INCOMPLETE', 'FLAGGED', 'CANCELLED']).astype(int)

    if 'bankname' in df.columns:
        df['bank_tier'] = df['bankname'].apply(lambda x: 'Tier 1 Bank' if x in TIER_1_BANKS else 'Tier 2 Bank')
    
    if 'cardnumber' in df.columns:
         df['upi_handle'] = df['cardnumber'].astype(str).apply(lambda x: x.split('@')[1] if '@' in x else None).str.lower()
         df['psp_app'] = df['upi_handle'].map(PSP_MAP).fillna(df['upi_handle'])

    if 'cardcountry' in df.columns:
        df['card_category'] = df['cardcountry'].apply(lambda x: 'DOMESTIC' if x == 'IN' else 'IPG')

    def categorize(row):
        mode = row.get('paymentmode')
        if mode == 'UPI':
            if pd.notna(row.get('bankname')) and row.get('bankname') != 'NAN': return 'UPI_INTENT'
            return 'UPI_COLLECT'
        elif mode in CARD_MODES:
            geo = 'DOMESTIC' if row.get('cardcountry') == 'IN' else 'INTERNATIONAL'
            return f"{mode}_{geo}"
        elif mode == 'NET_BANKING':
            return 'NB_TIER_1' if row.get('bankname') in TIER_1_BANKS else 'NB_TIER_2'
        return mode 
    df['sub_category'] = df.apply(categorize, axis=1)

    return df

# --- RCA HELPER ---
def get_failure_spike(curr_df, prev_df, group_cols, mode_group):
    curr_err = curr_df[curr_df['txstatus']!='SUCCESS'].groupby(group_cols).size().reset_index(name='curr_count')
    prev_err = prev_df[prev_df['txstatus']!='SUCCESS'].groupby(group_cols).size().reset_index(name='prev_count')
    
    merged = pd.merge(curr_err, prev_err, on=group_cols, how='outer').fillna(0)
    merged['spike'] = merged['curr_count'] - merged['prev_count']
    
    # Calculate contribution
    c_fails_total = len(curr_df[curr_df['txstatus']!='SUCCESS'])
    p_fails_total = len(prev_df[prev_df['txstatus']!='SUCCESS'])
    total_fail_increase = c_fails_total - p_fails_total
    
    if total_fail_increase > 0:
        merged['contribution'] = (merged['spike'] / total_fail_increase) * 100
    else:
        merged['contribution'] = 0.0
    
    # Context
    context_list = []
    drill_col = None
    if mode_group == 'UPI': drill_col = 'sub_category'
    elif mode_group == 'CARDS': drill_col = 'cardtype'
    elif mode_group == 'NET_BANKING': drill_col = 'bankname'
    
    if drill_col:
        for index, row in merged.iterrows():
            reason = row['txmsg']
            specific_fails = curr_df[(curr_df['txstatus']!='SUCCESS') & (curr_df['txmsg'] == reason)]
            if not specific_fails.empty:
                top_affected = specific_fails[drill_col].value_counts().head(3)
                context_str = " | ".join([f"{k}: {v}" for k, v in top_affected.items()])
                context_list.append(context_str)
            else: context_list.append("N/A")
    else: context_list = ["N/A"] * len(merged)

    merged['Context'] = context_list
    return merged.sort_values('spike', ascending=False)

def compare_periods(curr_df, prev_df, group_cols):
    if isinstance(group_cols, str): group_cols = [group_cols]
    stats_c = curr_df.groupby(group_cols)['is_success'].agg(['count','sum']).reset_index().rename(columns={'count':'Vol_curr', 'sum':'Succ_curr'})
    stats_p = prev_df.groupby(group_cols)['is_success'].agg(['count','sum']).reset_index().rename(columns={'count':'Vol_prev', 'sum':'Succ_prev'})
    merged = pd.merge(stats_c, stats_p, on=group_cols, how='outer').fillna(0)
    merged['SR_curr'] = (merged['Succ_curr'] / merged['Vol_curr'].replace(0,1)) * 100
    merged['SR_prev'] = (merged['Succ_prev'] / merged['Vol_prev'].replace(0,1)) * 100
    merged['SR_Delta'] = merged['SR_curr'] - merged['SR_prev']
    merged['Vol_Delta'] = merged['Vol_curr'] - merged['Vol_prev']
    return merged.sort_values('SR_Delta')

def color_delta(val):
    if val < 0: color = '#ff4b4b' # Red
    elif val > 0: color = '#09ab3b' # Green
    else: color = 'white'
    return f'color: {color}'

# --- SHORT SUMMARY GENERATOR ---
def generate_day_summary(curr_df, prev_df):
    if curr_df.empty: return "No data."
    
    c_vol = len(curr_df)
    p_vol = len(prev_df)
    c_sr = (curr_df['is_success'].sum()/c_vol*100) if c_vol > 0 else 0
    p_sr = (prev_df['is_success'].sum()/p_vol*100) if p_vol > 0 else 0
    sr_diff = c_sr - p_sr
    
    if sr_diff >= -0.5:
        return "✅ Performance is stable."
    
    c_ud = len(curr_df[curr_df['is_userdrop']==1])
    p_ud = len(prev_df[prev_df['is_userdrop']==1])
    ud_inc = c_ud - p_ud
    
    c_fail = len(curr_df[curr_df['is_failed']==1])
    p_fail = len(prev_df[prev_df['is_failed']==1])
    fail_inc = c_fail - p_fail
    
    reason = "User Drops" if ud_inc > fail_inc else "Technical Failures"
    mode_counts = curr_df[curr_df['txstatus']!='SUCCESS']['paymentmode'].value_counts()
    top_mode = mode_counts.index[0] if not mode_counts.empty else "Unknown"
    
    return f"📉 SR dropped by {abs(sr_diff):.1f}% due to increased **{reason}** in **{top_mode}**."

# --- DRILL DOWN RENDERER ---
def render_metric_tab(curr_df, prev_df, metric_col, title):
    c_count = curr_df[metric_col].sum()
    p_count = prev_df[metric_col].sum()
    diff = c_count - p_count
    
    c_vol_total = len(curr_df)
    c_pct = (c_count / c_vol_total * 100) if c_vol_total > 0 else 0
    
    p_vol_total = len(prev_df)
    p_pct = (p_count / p_vol_total * 100) if p_vol_total > 0 else 0
    pct_diff = c_pct - p_pct
    
    status = "📈" if pct_diff > 0 else "📉"
    st.markdown(f"**{title}: {c_pct:.1f}% ({pct_diff:+.1f}%) {status}**")
    
    if diff <= 0 and pct_diff <= 0:
        st.caption(f"No increase in {title}. (Count: {p_count} -> {c_count})")
        return

    st.markdown("---")
    
    modes = ['UPI', 'CARDS', 'NET_BANKING']
    
    for mode in modes:
        if mode == 'UPI':
            m_curr = curr_df[curr_df['paymentmode']=='UPI']
            m_prev = prev_df[prev_df['paymentmode']=='UPI']
            drill_col = 'upi_handle'
        elif mode == 'CARDS':
            m_curr = curr_df[curr_df['paymentmode'].isin(CARD_MODES)]
            m_prev = prev_df[prev_df['paymentmode'].isin(CARD_MODES)]
            drill_col = 'cardtype'
        else:
            m_curr = curr_df[curr_df['paymentmode']=='NET_BANKING']
            m_prev = prev_df[prev_df['paymentmode']=='NET_BANKING']
            drill_col = 'bankname'
            
        if m_curr.empty: continue

        m_c_count = m_curr[metric_col].sum()
        m_p_count = m_prev[metric_col].sum()
        m_diff = m_c_count - m_p_count
        
        m_c_vol = len(m_curr)
        m_c_rate = (m_c_count / m_c_vol * 100) if m_c_vol > 0 else 0
        
        m_p_vol = len(m_prev)
        m_p_rate = (m_p_count / m_p_vol * 100) if m_p_vol > 0 else 0
        
        if m_diff > 0:
            c1, c2 = st.columns([1.5, 2])
            with c1:
                st.markdown(f"**{mode}**")
                # Show Previous -> Current Count
                st.write(f"Count: {int(m_p_count)} → {int(m_c_count)} (+{int(m_diff)})")
                # Show Previous -> Current Rate (with Delta)
                st.caption(f"Rate: {m_p_rate:.1f}% → {m_c_rate:.1f}% ({m_c_rate-m_p_rate:+.1f}%)")
            
            with c2:
                if drill_col in m_curr.columns:
                    grp_c = m_curr[m_curr[metric_col]==1].groupby(drill_col).size().reset_index(name='Curr')
                    grp_p = m_prev[m_prev[metric_col]==1].groupby(drill_col).size().reset_index(name='Prev')
                    merged = pd.merge(grp_c, grp_p, on=drill_col, how='outer').fillna(0)
                    merged['Inc'] = merged['Curr'] - merged['Prev']
                    merged['% Impact'] = (merged['Curr'] / m_c_vol * 100).round(2)
                    
                    top = merged.sort_values('Inc', ascending=False).head(3)
                    
                    if not top.empty and top.iloc[0]['Inc'] > 0:
                        st.dataframe(
                            top.style.format({'Curr':'{:.0f}', 'Prev':'{:.0f}', 'Inc':'{:.0f}', '% Impact':'{:.2f}%'}), 
                            hide_index=True, 
                            width="content"
                        )

# --- MAIN APP ---
uploaded_file = st.file_uploader("Upload CSV File", type=['csv'])

if uploaded_file:
    with st.spinner("Processing & Cleaning Data..."):
        # ⚡️ REMOVED pyarrow engine to avoid crashes
        raw_df = pd.read_csv(uploaded_file, low_memory=False) 
        raw_df = clean_columns(raw_df)
        df = preprocess_data(raw_df)
        del raw_df
        gc.collect()

    tab_report, tab_rca, tab_summary = st.tabs(["📊 Report Generator", "🔍 Insights & RCA", "📝 Smart Summary"])

    # ==========================================
    # TAB 1: REPORT GENERATOR
    # ==========================================
    with tab_report:
        st.header("Generate Standard Excel Reports")
        col_f1, col_f2 = st.columns(2)
        with col_f1: share_with_merchant = st.checkbox("Share with Merchant (Remove PG Data)", value=False)
        with col_f2: run_hourly = st.checkbox("Get Hourly SR Report")

        df_display = df.copy()
        if share_with_merchant:
            if 'pg' in df_display.columns:
                df_display = df_display.drop(columns=['pg'])
                st.warning("PG Data removed from reports.")

        hourly_df = None
        if run_hourly:
            min_d, max_d = df_display['Day'].min(), df_display['Day'].max()
            hourly_range = st.date_input("Select Date for Hourly Report", [min_d, max_d], min_value=min_d, max_value=max_d)
            if len(hourly_range) == 2:
                hourly_df = df_display[(df_display['Day'] >= hourly_range[0]) & (df_display['Day'] <= hourly_range[1])].copy()

        with st.expander("Filter Data (Date, Merchant, Mode)", expanded=True):
            use_date_filter = st.checkbox("Filter by Date Range", value=False)
            if use_date_filter:
                d_r = st.date_input("Select Date Range", [df_display['Day'].min(), df_display['Day'].max()])
                if len(d_r) == 2: df_display = df_display[(df_display['Day'] >= d_r[0]) & (df_display['Day'] <= d_r[1])]

            merchants = sorted([str(x) for x in df_display['merchantid'].unique() if pd.notna(x)])
            sel_merch = st.multiselect("Select Merchants", merchants, default=[]) 
            if sel_merch:
                df_display = df_display[df_display['merchantid'].astype(str).isin(sel_merch)]
                if hourly_df is not None: hourly_df = hourly_df[hourly_df['merchantid'].astype(str).isin(sel_merch)]

            if 'paymentmode' in df_display.columns:
                modes = sorted([str(x) for x in df_display['paymentmode'].unique() if pd.notna(x)])
                sel_modes = st.multiselect("Select Payment Modes", modes, default=[]) 
                if sel_modes:
                    df_display = df_display[df_display['paymentmode'].isin(sel_modes)]
                    if hourly_df is not None: hourly_df = hourly_df[hourly_df['paymentmode'].isin(sel_modes)]

                    if any(mode in sel_modes for mode in CARD_MODES):
                        st.markdown("---")
                        st.subheader("💳 Card Filters")
                        if 'cardcountry' in df_display.columns:
                            all_cats = sorted([str(x) for x in df_display['card_category'].unique() if pd.notna(x)])
                            selected_cats = st.multiselect("Select Card Category (Geo)", all_cats, default=[])
                            if selected_cats:
                                df_display = df_display[(df_display['card_category'].isin(selected_cats)) | (~df_display['paymentmode'].isin(CARD_MODES))]
                                if hourly_df is not None: hourly_df = hourly_df[(hourly_df['card_category'].isin(selected_cats)) | (~hourly_df['paymentmode'].isin(CARD_MODES))]
                        
                        if 'cardtype' in df_display.columns:
                            all_card_types = sorted([str(x) for x in df_display['cardtype'].unique() if pd.notna(x)])
                            selected_card_types = st.multiselect("Select Card Network", all_card_types, default=[])
                            if selected_card_types: 
                                df_display = df_display[(df_display['cardtype'].isin(selected_card_types)) | (~df_display['paymentmode'].isin(CARD_MODES))]
                                if hourly_df is not None: hourly_df = hourly_df[(hourly_df['cardtype'].isin(selected_card_types)) | (~hourly_df['paymentmode'].isin(CARD_MODES))]

        st.info(f"Ready to analyze **{len(df_display)}** transactions.")

        if st.button("🚀 Run Report Analysis", key="btn_rep"):
            with st.spinner('Generating Excel Files...'):
                num_merchants = df_display['merchantid'].nunique()
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
                    'Paymode+Bank': ['paymentmode', 'bankname'],
                    'Paymode+PG+Bank': ['paymentmode', 'pg', 'bankname']
                }

                generated_buffers = {}
                
                for report_type, config in report_configs.items():
                    current_df = config.get('data', df_display)
                    time_col = config['time_col']
                    time_group = [time_col] if time_col else []

                    if report_type == 'Overview': config['sheets']['SR Overall'] = ([], current_df)
                    else: config['sheets'][f'SR {report_type}'] = (time_group, current_df)
                    
                    for name, group_cols in base_breakdowns.items():
                        if not all(col in current_df.columns for col in group_cols): continue
                        config['sheets'][f'SR by {name}'] = (group_cols + time_group, current_df)

                    config['sheets'][f'SR by Paymode'] = (time_group + ['paymentmode'], current_df)
                    
                    if 'cardtype' in current_df.columns:
                        card_df = current_df[current_df['paymentmode'].isin(CARD_MODES)]
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
                         bank_tier_df = current_df[current_df['paymentmode'] == 'NET_BANKING']
                         if not bank_tier_df.empty:
                            config['sheets'][f'SR by Bank Tier'] = (time_group + ['paymentmode', 'bank_tier'], bank_tier_df)

                    if 'card_category' in current_df.columns:
                        card_cat_df = current_df[current_df['paymentmode'].isin(CARD_MODES)]
                        if not card_cat_df.empty:
                            config['sheets'][f'SR by Card Geo'] = (time_group + ['paymentmode', 'card_category'], card_cat_df)
                            
                    # --- NEW: SR BY AMOUNT CATEGORY ---
                    if 'amount_category' in current_df.columns:
                        config['sheets']['SR by Amount Category'] = (time_group + ['paymentmode', 'amount_category'], current_df)

                    output_buffer = io.BytesIO()
                    with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
                        has_data = False
                        for sheet_name, (cols, dataset) in config['sheets'].items():
                            result = compute_sr(dataset, cols, 'merchantid', 'transactionid', 'amount', 'txstatus', num_merchants)
                            if config['time_col'] and config['time_col'] in result.columns:
                                result = result.sort_values(by=config['time_col'], ascending=True)
                            if report_type in ['Monthly', 'Weekly'] and not result.empty:
                                result = compute_mom_change(result, cols, 'merchantid')
                            if not result.empty:
                                result.to_excel(writer, sheet_name=sheet_name[:31], index=False)
                                has_data = True
                        
                        fail_data = current_df[current_df['txstatus'] != "SUCCESS"]
                        if not fail_data.empty:
                            fail_cols = ['merchantid']
                            if time_col: fail_cols.append(time_col)
                            fail_cols.extend(['paymentmode', 'txmsg'])
                            fail_summary = fail_data.groupby(fail_cols, dropna=False)['transactionid'].count().reset_index(name='Volume')
                            sort_c = [time_col] if time_col else ['Volume']
                            asc_s = [True] if time_col else [False]
                            fail_summary.sort_values(by=sort_c, ascending=asc_s).to_excel(writer, sheet_name='Failures Analysis', index=False)
                            
                            if 'bankname' in fail_data.columns:
                                bank_fail_cols = ['merchantid', 'paymentmode', 'bankname', 'txmsg']
                                if time_col: bank_fail_cols.insert(1, time_col)
                                bank_fail_summary = fail_data.groupby(bank_fail_cols, dropna=False)['transactionid'].count().reset_index(name='Volume')
                                bank_fail_summary.sort_values(by='Volume', ascending=False).to_excel(writer, sheet_name='Failures (Paymode+Bank)', index=False)
                            
                            # --- NEW: FAILURES (PAYMODE+PG+REASON) ---
                            if 'pg' in fail_data.columns:
                                pg_fail_cols = ['merchantid', 'paymentmode', 'pg', 'txmsg']
                                if time_col: pg_fail_cols.insert(1, time_col)
                                pg_fail_summary = fail_data.groupby(pg_fail_cols, dropna=False)['transactionid'].count().reset_index(name='Volume')
                                pg_fail_summary.sort_values(by='Volume', ascending=False).to_excel(writer, sheet_name='Failures (Paymode+PG)', index=False)

                            has_data = True

                        if not has_data: pd.DataFrame([{"Info": "No data"}]).to_excel(writer, sheet_name="No Data", index=False)
                    
                    generated_buffers[report_type] = apply_formatting(output_buffer)

                st.session_state['daily_report'] = generated_buffers['Daily']
                st.session_state['weekly_report'] = generated_buffers['Weekly']
                st.session_state['monthly_report'] = generated_buffers['Monthly']
                st.session_state['overview_report'] = generated_buffers['Overview']
                if 'Hourly' in generated_buffers: st.session_state['hourly_report'] = generated_buffers['Hourly']
                st.success("✅ Analysis Complete! Download reports below.")

        if 'daily_report' in st.session_state:
            st.markdown("### 📥 Download Reports")
            c1, c2 = st.columns(2)
            with c1: st.download_button("🌍 Overview Report", st.session_state['overview_report'], "Overview_SR.xlsx", use_container_width=True)
            if 'hourly_report' in st.session_state:
                with c2: st.download_button("🕒 Hourly Report", st.session_state['hourly_report'], "Hourly_SR.xlsx", use_container_width=True)
            c3, c4, c5 = st.columns(3)
            with c3: st.download_button("📅 Daily Report", st.session_state['daily_report'], "Daily_SR.xlsx")
            with c4: st.download_button("📆 Weekly Report", st.session_state['weekly_report'], "Weekly_SR.xlsx")
            with c5: st.download_button("🗓️ Monthly Report", st.session_state['monthly_report'], "Monthly_SR.xlsx")

    with tab_rca:
        st.header("📉 Automated Root Cause Analysis")
        col_rca_1, col_rca_2 = st.columns(2)
        with col_rca_1:
            days_options = [2, 3, 5, 7, 10, 15, 30]
            selected_days = st.slider("Select Comparison Period (Last X Days)", min_value=1, max_value=30, value=7)
        with col_rca_2:
            rca_merchants = sorted([str(x) for x in df['merchantid'].unique() if pd.notna(x)])
            selected_rca_merch = st.multiselect("Filter Merchant ID (RCA Only)", rca_merchants, default=[])

        rca_df = df.copy()
        if selected_rca_merch:
            rca_df = rca_df[rca_df['merchantid'].astype(str).isin(selected_rca_merch)]

        max_date = rca_df['Day'].max()
        curr_start = max_date - pd.Timedelta(days=selected_days - 1)
        prev_start = curr_start - pd.Timedelta(days=selected_days)
        prev_end = curr_start - pd.Timedelta(days=1)
        st.caption(f"**Current:** {curr_start} to {max_date} | **Previous:** {prev_start} to {prev_end}")
        
        curr_df = rca_df[(rca_df['Day'] >= curr_start) & (rca_df['Day'] <= max_date)]
        prev_df = rca_df[(rca_df['Day'] >= prev_start) & (rca_df['Day'] <= prev_end)]
        
        if curr_df.empty or prev_df.empty:
            st.error("Not enough data for comparison in the selected range/merchant.")
        else:
            c_vol = len(curr_df)
            p_vol = len(prev_df)
            vol_delta = c_vol - p_vol
            c_sr = (curr_df['is_success'].sum()/c_vol*100) if c_vol > 0 else 0
            p_sr = (prev_df['is_success'].sum()/p_vol*100) if p_vol > 0 else 0
            sr_delta = c_sr - p_sr
            c_gmv = curr_df[curr_df['txstatus']=='SUCCESS']['amount'].sum()
            p_gmv = prev_df[prev_df['txstatus']=='SUCCESS']['amount'].sum()
            gmv_delta = c_gmv - p_gmv

            m1, m2, m3, m4, m5, m6 = st.columns(6)
            m1.metric("Curr SR", f"{c_sr:.2f}%")
            m2.metric("Prev SR", f"{p_sr:.2f}%", f"{sr_delta:.2f}%")
            m3.metric("Curr Vol", f"{c_vol}")
            m4.metric("Prev Vol", f"{p_vol}", f"{vol_delta}")
            m5.metric("Curr GMV", f"₹{c_gmv:,.0f}")
            m6.metric("Prev GMV", f"₹{p_gmv:,.0f}", f"{gmv_delta:,.0f}")
            
            st.markdown("---")
            modes = ['UPI', 'CARDS', 'NET_BANKING']
            for mode_group in modes:
                with st.expander(f"Analysis: {mode_group}", expanded=(mode_group=='UPI')):
                    if mode_group == 'UPI':
                        m_curr = curr_df[curr_df['paymentmode']=='UPI']
                        m_prev = prev_df[prev_df['paymentmode']=='UPI']
                    elif mode_group == 'CARDS':
                        m_curr = curr_df[curr_df['paymentmode'].isin(CARD_MODES)]
                        m_prev = prev_df[prev_df['paymentmode'].isin(CARD_MODES)]
                    else: 
                        m_curr = curr_df[curr_df['paymentmode']=='NET_BANKING']
                        m_prev = prev_df[prev_df['paymentmode']=='NET_BANKING']

                    if m_curr.empty:
                        st.info("No data for this mode.")
                        continue

                    stats_merged = compare_periods(m_curr, m_prev, 'sub_category')
                    worst_sub = stats_merged.sort_values('SR_Delta').head(1)
                    
                    c1, c2 = st.columns([1, 1])
                    with c1:
                        st.write("**Sub-Category Performance**")
                        st.dataframe(
                            stats_merged[['sub_category', 'SR_curr', 'SR_prev', 'SR_Delta', 'Vol_curr', 'Vol_prev', 'Vol_Delta']]
                            .style.format({
                                'SR_curr': '{:.2f}%', 'SR_prev': '{:.2f}%', 'SR_Delta': '{:.2f}%',
                                'Vol_curr': '{:.0f}', 'Vol_prev': '{:.0f}', 'Vol_Delta': '{:.0f}'
                            }).map(color_delta, subset=['SR_Delta', 'Vol_Delta']),
                            width="content", hide_index=True
                        )
                        if not worst_sub.empty and worst_sub['SR_Delta'].values[0] < -1:
                            st.error(f"🚨 **Issue Detected in {worst_sub['sub_category'].values[0]}** (Dropped {worst_sub['SR_Delta'].values[0]:.2f}%)")
                        else: st.success("✅ No major sub-category drop.")

                        if mode_group == 'UPI' and 'upi_handle' in m_curr.columns:
                            st.markdown("##### 🔍 UPI Handle Breakdown")
                            handle_stats = compare_periods(m_curr, m_prev, 'upi_handle')
                            handle_stats = handle_stats.sort_values('Vol_Delta', ascending=True)
                            st.dataframe(
                                handle_stats[['upi_handle', 'SR_curr', 'SR_prev', 'SR_Delta', 'Vol_curr', 'Vol_prev', 'Vol_Delta']]
                                .style.format({
                                    'SR_curr': '{:.2f}%', 'SR_prev': '{:.2f}%', 'SR_Delta': '{:.2f}%', 
                                    'Vol_curr': '{:.0f}', 'Vol_prev': '{:.0f}', 'Vol_Delta': '{:.0f}'
                                }).map(color_delta, subset=['SR_Delta', 'Vol_Delta']),
                                width="content", hide_index=True
                            )

                        if mode_group == 'CARDS' and 'cardtype' in m_curr.columns:
                            st.markdown("##### 🔍 Card Type Breakdown")
                            card_stats = compare_periods(m_curr, m_prev, ['paymentmode', 'cardtype'])
                            st.dataframe(
                                card_stats[['paymentmode', 'cardtype', 'SR_curr', 'SR_prev', 'SR_Delta', 'Vol_curr', 'Vol_prev', 'Vol_Delta']]
                                .style.format({
                                    'SR_curr': '{:.2f}%', 'SR_prev': '{:.2f}%', 'SR_Delta': '{:.2f}%', 
                                    'Vol_curr': '{:.0f}', 'Vol_prev': '{:.0f}', 'Vol_Delta': '{:.0f}'
                                }).map(color_delta, subset=['SR_Delta', 'Vol_Delta']),
                                width="content", hide_index=True
                            )

                    with c2:
                        trend = m_curr.groupby('Day').agg(SR=('is_success', 'mean'), Vol=('transactionid', 'count')).reset_index()
                        trend['SR'] = trend['SR'] * 100
                        fig = go.Figure()
                        fig.add_trace(go.Bar(x=trend['Day'], y=trend['Vol'], name='Volume', hovertemplate='<b>Vol:</b> %{y:.0f}<extra></extra>', marker_color='rgba(135, 206, 250, 0.6)'))
                        fig.add_trace(go.Scatter(x=trend['Day'], y=trend['SR'], name='SR %', yaxis='y2', hovertemplate='<b>SR:</b> %{y:.2f}%<extra></extra>', line=dict(color='red', width=3)))
                        fig.update_layout(title=f'{mode_group} Trend', yaxis=dict(title='Volume'), yaxis2=dict(title='SR %', overlaying='y', side='right', range=[0, 100]), hovermode="x unified")
                        st.plotly_chart(fig,width='stretch',config={"displayModeBar": False}
)


                    st.subheader("Why did it drop?")
                    fc1, fc2 = st.columns(2)
                    with fc1:
                        st.markdown("**1. PG Performance**")
                        if 'pg' in m_curr.columns:
                            pg_merged = compare_periods(m_curr, m_prev, 'pg')
                            pg_merged = pg_merged.sort_values('SR_Delta', ascending=True)
                            st.write("📉 **PGs Performance**")
                            st.dataframe(
                                pg_merged[['pg', 'SR_curr', 'SR_prev', 'SR_Delta', 'Vol_curr', 'Vol_prev', 'Vol_Delta']]
                                .style.format({
                                    'SR_curr': '{:.2f}%', 'SR_prev': '{:.2f}%', 'SR_Delta': '{:.2f}%', 
                                    'Vol_curr': '{:.0f}', 'Vol_prev': '{:.0f}', 'Vol_Delta': '{:.0f}'
                                }).map(color_delta, subset=['SR_Delta', 'Vol_Delta']),
                                width="content", hide_index=True
                            )

                    with fc2:
                        st.markdown("**2. Error Analysis (Volume Spikes)**")
                        spikes = get_failure_spike(m_curr, m_prev, ['txmsg'], mode_group)
                        if not spikes.empty:
                            disp_spikes = spikes[['txmsg', 'curr_count', 'prev_count', 'spike', 'contribution', 'Context']].copy()
                            disp_spikes.columns = ['Error Message', 'Curr Vol', 'Prev Vol', 'Vol Spike', 'Contrib %', 'Failure Context']
                            st.dataframe(
                                disp_spikes.style.format({
                                    'Curr Vol': '{:.0f}', 
                                    'Prev Vol': '{:.0f}', 
                                    'Vol Spike': '{:.0f}',
                                    'Contrib %': '{:.2f}%'
                                }).background_gradient(subset=['Vol Spike'], cmap='Reds'),
                                width="content",
                                hide_index=True
                            )
                        else: st.info("No specific error spike detected.")

    # ==========================================
    # TAB 3: SMART SUMMARY (METRIC FOCUSED)
    # ==========================================
    with tab_summary:
        st.header("📝 Smart Day-over-Day Analysis")
        
        col_s1, col_s2 = st.columns(2)
        with col_s1:
            start_d = st.date_input("Start Date", value=df['Day'].min())
        with col_s2:
            end_d = st.date_input("End Date", value=df['Day'].max())
            
        sum_merch = st.multiselect("Filter Merchant (Optional)", sorted([str(x) for x in df['merchantid'].unique() if pd.notna(x)]), default=[])
        
        sum_df = df[(df['Day'] >= start_d) & (df['Day'] <= end_d)].copy()
        if sum_merch:
            sum_df = sum_df[sum_df['merchantid'].astype(str).isin(sum_merch)]
            
        if sum_df.empty:
            st.error("No data found for selected range.")
        else:
            dates = sorted(sum_df['Day'].unique())
            if len(dates) < 2:
                st.warning("Need at least 2 days of data.")
            else:
                st.subheader(f"📊 Daily Deep Dives ({len(dates)-1} comparisons)")
                
                for i in range(1, len(dates)):
                    curr_date = dates[i]
                    prev_date = dates[i-1]
                    
                    d_curr = sum_df[sum_df['Day'] == curr_date]
                    d_prev = sum_df[sum_df['Day'] == prev_date]
                    
                    # Header Stats
                    c_sr = (d_curr['is_success'].sum()/len(d_curr)*100) if len(d_curr)>0 else 0
                    p_sr = (d_prev['is_success'].sum()/len(d_prev)*100) if len(d_prev)>0 else 0
                    sr_diff = c_sr - p_sr
                    
                    c_vol = len(d_curr)
                    p_vol = len(d_prev)
                    vol_diff = c_vol - p_vol
                    
                    icon = "📉" if sr_diff < -1 else "✅"
                    
                    # Generate Summary String
                    summary_text = generate_day_summary(d_curr, d_prev)
                    
                    # Expanded logic
                    with st.expander(f"{icon} **{curr_date.strftime('%d-%b')}** (vs {prev_date.strftime('%d-%b')}) | SR: {c_sr:.1f}% ({sr_diff:+.1f}%) | Vol: {p_vol} → {c_vol} ({vol_diff:+})"):
                        
                        st.info(summary_text)
                        
                        t_sr, t_ud, t_fail, t_inc = st.tabs(["📉 Success Rate", "🚧 User Dropped", "💥 Failed", "⏳ Incomplete"])
                        
                        # 1. SR Tab
                        with t_sr:
                            # Show Modes Table
                            stats_mode = compare_periods(d_curr, d_prev, 'paymentmode').sort_values('SR_Delta')
                            st.dataframe(stats_mode[['paymentmode', 'SR_curr', 'SR_prev', 'SR_Delta', 'Vol_curr']].style.format({
                                'SR_curr': '{:.1f}%', 'SR_prev': '{:.1f}%', 'SR_Delta': '{:+.1f}%', 'Vol_curr': '{:.0f}'
                            }).map(color_delta, subset=['SR_Delta']), width="content", hide_index=True)

                        # 2. User Dropped Tab
                        with t_ud:
                            render_metric_tab(d_curr, d_prev, 'is_userdrop', "User Drops")

                        # 3. Failed Tab
                        with t_fail:
                            render_metric_tab(d_curr, d_prev, 'is_failed', "Technical Failures")

                        # 4. Incomplete Tab
                        with t_inc:
                            render_metric_tab(d_curr, d_prev, 'is_incomplete', "Incomplete Txns")

else:
    st.info("👆 Upload CSV to start.")
