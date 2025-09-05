import pandas as pd

cost_centres = ["SAFROPS", "SAFROP2", "SAFROP3", "SAFROP4", "SAFROP5", "SAFROP6", "SAFROP7", "SAHSSHU", "SASSADM"]

def xlsx_gen(df):
    """Generates an Excel file with sheets corresponding to cost centres."""
    with pd.ExcelWriter('output.xlsx', engine='xlsxwriter') as writer:

        orange_format = writer.book.add_format({'bg_color': '#FFC000', 'bold': True, 'border': 1})
        bold_format = writer.book.add_format({'bold': True})

        for cc in cost_centres:
            ids = set()
            current_df = df[df['Cost Centre'] == cc]
            current_df.to_excel(writer, sheet_name=cc, startrow=0, index=False)
            working_sheet = writer.sheets[cc]
            
            for col_num, col_name in enumerate(current_df.columns):
                working_sheet.write(0, col_num, col_name, orange_format)

            cc_amount = current_df['Amount'].sum()
            working_sheet.write(len(current_df) + 2, 3, f"Total: {cc_amount}")

            for col_num, col_name in enumerate(current_df.columns):
                column_len = max(current_df[col_name].astype(str).map(len).max(), len(col_name) + 2)
                working_sheet.set_column(col_num, col_num, column_len)

            for id in current_df['Name']:
                ids.add(id)
            
            offset = 6
            working_sheet.write(len(current_df) + (offset - 1), 2, "Name", orange_format)
            working_sheet.write(len(current_df) + (offset - 1), 3, "Total", orange_format)
            for id in ids:
                bold = False
                if len(current_df[current_df['Name'] == id]) > 1:
                    bold = True
                id_total = current_df[current_df['Name'] == id]['Amount'].sum()
                working_sheet.write(len(current_df) + offset, 2, id)
                working_sheet.write(len(current_df) + offset, 3, id_total) if not bold else working_sheet.write(len(current_df) + offset, 3, id_total, bold_format)
                offset += 1
        