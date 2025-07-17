import streamlit as st
import pandas as pd
import difflib
import re
import io

st.set_page_config(page_title="UPS Australia Matching Tool", layout="wide")

# ---------- Normalization Helper ----------
def normalize_name(name):
    if pd.isna(name):
        return ''
    name = str(name).upper()
    name = re.sub(r'[^A-Z0-9 ]', '', name)  # Remove punctuation
    name = re.sub(r'\b(AUSTRALIA|AUST|PTY|P/L|LTD|LIMITED|CORPORATION|INC|PTE|CO|THE|AND|&)\b', '', name)
    name = re.sub(r'\s+', ' ', name).strip()
    return name

# ---------- Matching Function with improved personal name heuristic ----------
def match_recipient_to_account(recipient_name, acc_df, threshold):
    normalized_recipient = normalize_name(recipient_name)

    acc_df['Normalized Name'] = acc_df['Customer Name'].apply(normalize_name)

    # Compute similarity scores
    acc_df['Similarity'] = acc_df['Normalized Name'].apply(
        lambda x: difflib.SequenceMatcher(None, normalized_recipient, x).ratio()
    )

    sorted_matches = acc_df.sort_values(by='Similarity', ascending=False)
    top_matches = sorted_matches.head(3)

    high_conf_matches = top_matches[top_matches['Similarity'] >= threshold]

    if len(high_conf_matches) == 1:
        best_match = high_conf_matches.iloc[0]
        return best_match['Account Number'], best_match['Similarity'], list(top_matches['Customer Name']), "✅ One strong match"
    
    elif len(high_conf_matches) > 1:
        # Check if first two words match
        rec_words = normalized_recipient.split()[:2]
        for _, row in high_conf_matches.iterrows():
            acc_words = row['Normalized Name'].split()[:2]
            if rec_words == acc_words:
                return row['Account Number'], row['Similarity'], list(top_matches['Customer Name']), "✅ First-two-word match"

        best_match = high_conf_matches.iloc[0]
        return best_match['Account Number'], best_match['Similarity'], list(top_matches['Customer Name']), "⚠ Multiple close matches"

    # Improved personal name heuristic:
    company_indicators = ['PTY', 'LTD', 'P/L', 'LABS', 'LABORATORIES', 'CORPORATION', 'INC', 'CO']
    if (len(normalized_recipient.split()) <= 1 and 
        not any(indicator in recipient_name.upper() for indicator in company_indicators)):
        return "Cash", 0, [], "👤 Treated as personal name"

    return "Cash", 0, list(top_matches['Customer Name']), "❌ No good match"

# ---------- UI ----------
st.title("🇦🇺 UPS AU Recipient Name Matching Tool")

uploaded_shipment = st.file_uploader("📦 Upload Shipment File (Excel)", type=["xls", "xlsx"])
uploaded_accounts = st.file_uploader("📒 Upload Account File (Excel)", type=["xls", "xlsx"])

threshold = st.slider("🔧 Similarity Threshold", 0.5, 1.0, 0.8, 0.01)

if uploaded_shipment and uploaded_accounts:
    try:
        ship_df = pd.read_excel(uploaded_shipment)
        acc_df = pd.read_excel(uploaded_accounts)

        if 'Recipient Company Name' not in ship_df.columns or 'Customer Name' not in acc_df.columns:
            st.error("❌ Required columns missing. Make sure files have 'Recipient Company Name' and 'Customer Name'.")
        else:
            results = []
            for _, row in ship_df.iterrows():
                acct, score, suggestions, comment = match_recipient_to_account(
                    row['Recipient Company Name'], acc_df, threshold
                )
                results.append({
                    'Tracking Number': row.get('Tracking Number', ''),
                    'Recipient Company Name': row['Recipient Company Name'],
                    'Matched Account': acct,
                    'Similarity Score': round(score, 3),
                    'Suggestions': ', '.join(str(s) for s in suggestions if pd.notna(s)),
                    'Match Notes': comment
                })

            result_df = pd.DataFrame(results)
            st.success("✅ Matching complete. Preview below:")
            st.dataframe(result_df, use_container_width=True)

            # Download helper
            def convert_df(df):
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df.to_excel(writer, index=False)
                return output.getvalue()

            st.download_button(
                label="📥 Download Matching Result",
                data=convert_df(result_df),
                file_name="matching_result.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    except Exception as e:
        st.error(f"❌ Error processing file: {e}")
else:
    st.info("👆 Please upload both shipment and account files.")
