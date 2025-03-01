import openpyxl
import pandas as pd
import plotly.graph_objects as go
import tkinter as tk
from tkinter import filedialog


def analyze_file():
    filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if not filepath:
        return  # Pokud není vybrán soubor, ukonči funkci

    try:
        # Přečti soubor Excel
        df_excel = pd.read_excel(filepath, header=36, engine="openpyxl")

        # Vypiš názvy sloupců pro diagnostiku
        print("Názvy sloupců:", df_excel.columns)

        # Odstraň poslední řádek
        df_dropped = df_excel.drop(index=df_excel.index[-1])
        
        # Oprava nesprávného kódování sloupce s datem
        df_dropped['#Datum zaúčtování transakce'] = pd.to_datetime(
            df_dropped['#Datum zaúčtování transakce'], format='%d-%m-%Y', errors='coerce'
        )

        # Měsíce pro analýzu
        months = ['2025-01', '2025-02', '2025-03', '2025-04', '2025-05', '2025-06',
                  '2025-07', '2025-08', '2025-09', '2025-10', '2025-11', '2025-12']
        sums_per_month = []

        for month in months:
            # Filtrace dat pro daný měsíc
            df_filtered = df_dropped[df_dropped['#Datum zaúčtování transakce'].dt.strftime('%Y-%m') == month]
            df_filtered.loc[:, '#Částka transakce'] = pd.to_numeric(df_filtered['#Částka transakce'], errors='coerce')
            total_sum = df_filtered['#Částka transakce'].sum()
            sums_per_month.append(total_sum)

        # Vytvoření grafu
        fig = go.Figure()
        fig.add_trace(go.Bar(
            x=months,
            y=sums_per_month,
            name='Součet transakcí',
            marker_color='indianred'
        ))
        fig.update_layout(
            title="Součet transakcí za měsíce v roce 2025 (Mbank)",
            barmode='group',
            xaxis_tickangle=-45
        )
        fig.show()
    except Exception as e:
        print(f"Chyba při zpracování souboru: {e}")


# Tkinter setup
root = tk.Tk()
root.title("Analyze")

tk.Label(root, text="Month transactions 2025 (Mbank)").pack(pady=5)
tk.Label(root, text="Click on button for select file .xlsx:").pack(pady=5)
tk.Button(root, text="Select and analyze", command=analyze_file).pack(pady=10)

root.mainloop()
