import streamlit as st
import pandas as pd
import io
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

st.title("🧾 EHR vs Card Reconciliation Tool")

ehr_file = st.file_uploader("Upload EHR CSV file", type=['csv'])
card_file = st.file_uploader("Upload Card Excel file", type=['xlsx'])

if ehr_file and card_file:
    if st.button("🔍 Run Reconciliation"):
        ehr_df = pd.read_csv(ehr_file)
        card_df = pd.read_excel(card_file)

        ehr_df['FormattedDate'] = pd.to_datetime(ehr_df['Payment Date'], errors='coerce').dt.strftime('%m/%d/%Y')
        card_df['FormattedDate'] = pd.to_datetime(card_df['Date'], errors='coerce').dt.strftime('%m/%d/%Y')

        ehr_df['Last4'] = ehr_df['Ref #'].astype(str).str.strip().str.extract(r'(\d{4})$')[0].fillna('MISSING')
        card_df['Last4'] = card_df['Card No'].astype(str).str.extract(r'(\d{4})$')[0].fillna('MISSING')
        ehr_df['Last4'] = ehr_df['Last4'].astype(str).str.zfill(4)
        card_df['Last4'] = card_df['Last4'].astype(str).str.zfill(4)

        ehr_df['Amount'] = pd.to_numeric(ehr_df['Amount'].astype(str).str.replace(r'[\$,]', '', regex=True), errors='coerce')
        card_df['Tran Amt'] = pd.to_numeric(card_df['Tran Amt'].astype(str).str.replace(r'[\$,]', '', regex=True), errors='coerce')

        ehr_with_card = ehr_df[ehr_df['Last4'] != 'MISSING']
        ehr_missing_card = ehr_df[ehr_df['Last4'] == 'MISSING']

        ehr_grouped = ehr_with_card.groupby(['FormattedDate', 'Last4'], as_index=False).agg({
            'Ref #': 'first', 'Payer Type': 'first', 'Amount': list
        }).rename(columns={'Ref #': 'Ref #', 'Payer Type': 'Payer', 'Amount': 'EHR Trxs'})
        ehr_grouped['EHR Tot Amt'] = ehr_grouped['EHR Trxs'].apply(sum)
        ehr_grouped = ehr_grouped.rename(columns={'FormattedDate': 'Date', 'Last4': 'Card Last 4'})

        ehr_missing_grouped = ehr_missing_card.copy()
        ehr_missing_grouped = ehr_missing_grouped.rename(columns={
            'FormattedDate': 'Date',
            'Last4': 'Card Last 4',
            'Ref #': 'Ref #',
            'Payer Type': 'Payer'
        })
        ehr_missing_grouped['EHR Trxs'] = ehr_missing_grouped[['Amount']].values.tolist()
        ehr_missing_grouped['EHR Tot Amt'] = ehr_missing_grouped['Amount']

        # 🔧 Fix duplicate column names
        ehr_missing_grouped = ehr_missing_grouped.loc[:, ~ehr_missing_grouped.columns.duplicated()]

        ehr_missing_grouped = ehr_missing_grouped[['Date', 'Card Last 4', 'Ref #', 'Payer', 'EHR Trxs', 'EHR Tot Amt']]
        ehr_combined = pd.concat([ehr_grouped, ehr_missing_grouped], ignore_index=True)

        card_grouped = card_df.groupby(['FormattedDate', 'Last4']).agg({'Tran Amt': list}).reset_index()
        card_grouped['Card Tot Amt'] = card_grouped['Tran Amt'].apply(sum)
        card_grouped = card_grouped.rename(columns={'FormattedDate': 'Date', 'Last4': 'Card Last 4', 'Tran Amt': 'Card Trxs'})

        merged_df = pd.merge(ehr_combined, card_grouped, how='outer', on=['Date', 'Card Last 4'])
        merged_df['EHR Tot Amt'] = merged_df['EHR Tot Amt'].fillna(0)
        merged_df['Card Tot Amt'] = merged_df['Card Tot Amt'].fillna(0)

        merged_df['Trx Match Status'] = merged_df.apply(
            lambda row: 'MATCH' if abs(row['EHR Tot Amt'] - row['Card Tot Amt']) < 0.01 else 'NON MATCH',
            axis=1
        )

        final_df = merged_df.sort_values(by='Date').reset_index(drop=True)

        grouped = []
        for date, group in final_df.groupby('Date', sort=True):
            grouped.append(group)
            total_row = {
                'Date': date,
                'Card Last 4': '',
                'Ref #': '',
                'Payer': '',
                'EHR Tot Amt': group['EHR Tot Amt'].sum(),
                'Card Tot Amt': group['Card Tot Amt'].sum(),
                'Trx Match Status': 'TOTAL',
                'EHR Trxs': '',
                'Card Trxs': ''
            }
            grouped.append(pd.DataFrame([total_row]))

        final_df_with_totals = pd.concat(grouped, ignore_index=True)

        final_df_with_totals['Daily Match Status'] = final_df_with_totals.apply(
            lambda row: 'MATCH' if row['Trx Match Status'] == 'TOTAL' and abs(row['EHR Tot Amt'] - row['Card Tot Amt']) < 0.01
            else 'NON MATCH' if row['Trx Match Status'] == 'TOTAL' else '',
            axis=1
        )

        card_lookup = card_df[['FormattedDate', 'Tran Amt']].dropna()

        def check_possible_match(row):
            if row['Trx Match Status'] != 'NON MATCH':
                return ''
            if row['Card Last 4'] == 'MISSING':
                matches = card_lookup[
                    (card_lookup['FormattedDate'] == row['Date']) &
                    (abs(card_lookup['Tran Amt'] - row['EHR Tot Amt']) < 0.01)
                ]
                return 'YES' if not matches.empty else ''
            return ''

        final_df_with_totals['Possible Match'] = final_df_with_totals.apply(check_possible_match, axis=1)

        # Save to Excel in memory
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            final_df_with_totals.to_excel(writer, sheet_name='Reconciliation', index=False)
            possible_matches = final_df_with_totals[final_df_with_totals['Possible Match'] == 'YES']
            if not possible_matches.empty:
                possible_matches.to_excel(writer, sheet_name='Possible Matches', index=False)
        output.seek(0)

        # Download button
        st.success("✅ Reconciliation complete!")
        st.download_button("⬇ Download Excel File", output, file_name="EHR_vs_Card_Reconciliation.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
