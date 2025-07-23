import gradio as gr
import pandas as pd
from openpyxl import Workbook
from io import BytesIO

def calculate_commission_file(file,file_name):
    # 判斷副檔名
    filename = file.name
    if filename.endswith('.xlsx'):
        df = pd.read_excel(file.name)
    elif filename.endswith('.txt'):
        df = pd.read_csv(file.name, sep=None, engine='python')
    else:
        raise ValueError("請提供 .xlsx 或 .txt 檔案")

    # 加入年-月欄位
    df['銷售日期'] = pd.to_datetime(df['銷售日期'], errors='coerce')
    year_month = df['銷售日期'].dt.to_period('M')

    if df['銷售日期'].isna().any():
        raise ValueError("請檢查日期格式")

    df['銷售日期'] = df['銷售日期'].dt.date
    df = df.sort_values(by=['銷售日期']).reset_index(drop=True)

    df['銷售累計百分比'] = 0.0
    df['抽成率'] = 0
    df['抽成額'] = 0

    month_cumsum = df.groupby(['櫃位編號', year_month])['銷售淨額'].cumsum()
    monthly_totals = df.groupby(['櫃位編號', year_month])['銷售淨額'].transform('sum')

    for i, row in df.iterrows():
        discount_rate = row['折扣率']
        cumulative = month_cumsum.iloc[i]
        monthly_total = monthly_totals.iloc[i]

        if discount_rate < 90:
            commission_rate = 22 if cumulative / monthly_total * 100 < 40 else 25
        else:
            commission_rate = 25

        df.at[i, '抽成率'] = commission_rate
        df.at[i, '抽成額'] = round(row['銷售淨額'] * commission_rate / 100)

    df = df.sort_values(by=['折扣率', '銷售日期']).reset_index(drop=True)

    running_total = 0
    for i, row in df.iterrows():
        running_total += row['銷售淨額']
        monthly_total = monthly_totals.iloc[i]
        pct = running_total / monthly_total * 100
        df.loc[i, '銷售累計百分比'] = round(pct, 2)

    grouped = []
    for discount_rate, group in df.groupby('折扣率', sort=False):
        group = group.copy()
        discount_total = group['銷售淨額'].sum()
        commission_total = group['抽成額'].sum()
        summary_row = pd.Series({
            '櫃位編號': '',
            '銷售日期': '',
            '折扣率': '小計',
            '銷售淨額': discount_total,
            '銷售累計百分比': '',
            '抽成率': '',
            '抽成額': commission_total
        })
        group = pd.concat([group, summary_row.to_frame().T], ignore_index=True)
        grouped.append(group)

    df = pd.concat(grouped, ignore_index=True)

    # 寫入 Excel
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "抽成計算"
    ws.append(list(df.columns))
    for _, row in df.iterrows():
        ws.append(list(row))

    for col in ws.columns:
        max_length = max(len(str(cell.value)) if cell.value else 0 for cell in col)
        ws.column_dimensions[col[0].column_letter].width = max_length + 2

    wb.save(output)
    output.seek(0)

    # 寫入暫存檔，改名給下載用
    temp_path = f"{file_name}.xlsx"
    df.to_excel(temp_path, index=False)
    return temp_path


# Gradio UI
description_html = """
上傳銷售資料 Excel（.xlsx），會根據表格中的資料自動計算抽成，然後讓你下載處理結果。<br><br>

<strong>📄 表格欄位格式範例如下：</strong><br>

<table border="1" style="border-collapse: collapse;">
    <tr>
        <th>櫃位編號</th>
        <th>銷售日期</th>
        <th>折扣率</th>
        <th>銷售淨額</th>
    </tr>
    <tr><td>320408</td><td>2025/7/1</td><td>50</td><td>990</td></tr>
    <tr><td>320408</td><td>2025/7/2</td><td>60</td><td>11000</td></tr>
    <tr><td>320408</td><td>2025/7/2</td><td>70</td><td>21010</td></tr>
    <tr><td>320408</td><td>2025/7/3</td><td>50</td><td>31020</td></tr>
    <tr><td>320408</td><td>2025/7/3</td><td>80</td><td>41030</td></tr>
    <tr><td>320408</td><td>2025/7/4</td><td>50</td><td>51040</td></tr>
</table>
"""

demo = gr.Interface(
    fn=calculate_commission_file,
    inputs=[
        gr.File(file_types=[".xlsx"]),
        gr.Textbox(label="輸出檔案名稱", value="處理結果")
    ],
    outputs=gr.File(label="下載抽成結果"),
    title="抽成計算工具",
    description=description_html
)

if __name__ == "__main__":
    demo.launch(share=True)
