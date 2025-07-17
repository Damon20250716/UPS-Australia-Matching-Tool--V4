
import streamlit as st
import pandas as pd
import re
from rapidfuzz import process, fuzz

# ---------------------- Normalization Utilities ---------------------- #
def normalize_name(name):
    if pd.isna(name):
        return ""
    name = name.upper()
    name = re.sub(r'[.,&/\\-]', ' ', name)
    name = re.sub(r'\s+', ' ', name).strip()

    # Replace known abbreviations
    replacements = {
        "LABS": "LABORATORIES",
        "AUST": "AUSTRALIA",
        "P L": "PTY LTD",
        "P/L": "PTY LTD",
        "P.L.": "PTY LTD",
        "PTY. LTD": "PTY LTD",
        "LTD": "",
        "LIMITED": "",
        "CO": "COMPANY",
    }
    for k, v in replacements.items():
        name = name.replace(k, v)
    name = re.sub(r'PTY LTD|AUSTRALIA|LIMITED|CORPORATION|INCORPORATED|INC|LTD|LLC', '', name)
    return name.strip()

def is_personal_name(name):
    return len(name.split()) <= 2 and not any(x in name.upper() for x in ['PTY', 'LTD', 'INC', 'CO', 'CORP'])

# ---------------------- Matching Logic ---------------------- #
def match_recipient_to_account(rec_name, account_df, threshold):
    rec_norm = normalize_name(rec_name)
    if is_personal_name(rec_name):
        return 'Cash', 0, [], 'Likely personal name'

    acc_df = account_df.copy()
    acc_df['Normalized Name'] = acc_df['Customer Name'].apply(normalize_name)

    # First round: high threshold fuzzy match
    best_match = process.extractOne(
        rec_norm,
        acc_df['Normalized Name'],
        scorer=fuzz.token_set_ratio,
        score_cutoff=threshold
    )
    if best_match:
        idx = acc_df[acc_df['Normalized Name'] == best_match[0]].index[0]
        return acc_df.at[idx, 'Account Number'], best_match[1], [acc_df.at[idx, 'Customer Name']], 'High confidence'

    # Second round: try partial matching with relaxed threshold
    fallback_match = process.extractOne(
        rec_norm,
        acc_df['Normalized Name'],
        scorer=fuzz.partial_ratio,
        score_cutoff=threshold - 10
    )
    if fallback_match:
        idx = acc_df[acc_df['Normalized Name'] == fallback_match[0]].index[0]
        return acc_df.at[idx, 'Account Number'], fallback_match[1], [acc_df.at[idx, 'Customer Name']], 'Partial fallback match'

    return 'Cash', 0, [], 'No confident match'

# ---------------------- Streamlit App ---------------------- #
st.set_page_config(page_title="UPS AU Matching Tool", layout="wide")
st.title("🇦🇺 UPS AU Recipient Name Matching Tool")

uploaded_ship = st.file_uploader("Upload Shipment File (Excel)", type=["xlsx"])
uploaded_acc = st.file_uploader("Upload Account File (Excel)", type=["xlsx"])
threshold = st.slider("Similarity Threshold", 70, 100, 85)

if uploaded_ship and uploaded_acc:
    ship_df = pd.read_excel(uploaded_ship)
    acc_df = pd.read_excel(uploaded_acc)

    if not all(col in ship_df.columns for col in ['Tracking Number', 'Recipient Company Name']):
        st.error("Shipment file must contain 'Tracking Number' and 'Recipient Company Name' columns.")
    elif not all(col in acc_df.columns for col in ['Customer Name', 'Account Number']):
        st.error("Account file must contain 'Customer Name' and 'Account Number' columns.")
    else:
        results = []
        for _, row in ship_df.iterrows():
            acct, score, suggestions, comment = match_recipient_to_account(
                row['Recipient Company Name'], acc_df, threshold
            )
            results.append({
                "Tracking Number": row['Tracking Number'],
                "Recipient Company Name": row['Recipient Company Name'],
                "Assigned Account": acct,
                "Match Score": score,
                "Top Suggestion": suggestions[0] if suggestions else "",
                "Comment": comment
            })
        result_df = pd.DataFrame(results)
        st.dataframe(result_df.style.apply(
            lambda row: ['background-color: red' if row.Match_Score < threshold and row.Assigned_Account == 'Cash' else
                         'background-color: yellow' if row.Match_Score < threshold else
                         'background-color: lightgreen'] * len(row), axis=1))

        # Download option
        def convert_df(df):
            return df.to_excel(index=False, engine='openpyxl')
        st.download_button("Download Result as Excel", convert_df(result_df), "matching_result.xlsx")
