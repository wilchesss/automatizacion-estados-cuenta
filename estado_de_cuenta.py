"""
Automatización de Estados de Cuenta
=====================================
Genera reportes individuales por cliente en formato Excel
a partir del reporte CUSTOM.xlsx exportado desde NetSuite.

Autor: Rodolfo Del Castillo Wilches
"""

import pandas as pd
from datetime import datetime

# ── 1. CARGAR Y LIMPIAR DATOS ────────────────────────────────────────────────

data = pd.read_excel("CUSTOM.xlsx")
data = data.iloc[5:]
data.columns = data.iloc[0]
data = data[1:].reset_index(drop=True)

# Filtrar filas de totales y limpiar fechas
data = data[~data['Factura #'].str.contains('Total', na=False)]
data['Fecha de vencimiento'] = pd.to_datetime(data['Fecha de vencimiento'], dayfirst=True)
data['Fecha Factura']        = pd.to_datetime(data['Fecha Factura'], dayfirst=True)

# Rellenar nombre de cliente hacia abajo (forward fill)
data["Cliente:Trabajo"] = data["Cliente:Trabajo"].replace(r'^\s*$', pd.NA, regex=True)
data["Cliente:Trabajo"] = data["Cliente:Trabajo"].ffill()

# Eliminar columnas innecesarias
cols_to_drop = ["C.E: UUID CFDI", "Ubicación: Nombre", "Tipo de transacción", "Balance"]
data = data.drop(columns=[c for c in cols_to_drop if c in data.columns])

# Asignar USD a facturas que tienen valor en columna Nota
if 'Nota' in data.columns:
    data.loc[data['Nota'].notna(), 'Moneda'] = 'USD'

# Renombrar columnas a inglés
data = data.rename(columns={
    'Fecha Factura':       'Inv Date',
    'Fecha de vencimiento':'Due Date',
    'Factura #':           'Invoice',
    'Nota':                'File',
    'Moneda':              'Currency',
    'Monto Factura':       'Amount',
    'Vencimiento':         'Days Overdue'
})

# Filtrar filas de totales que puedan quedar en Cliente:Trabajo
data = data[~data["Cliente:Trabajo"].str.contains("Total", na=False)]

# Asegurar tipos numéricos
data['Amount']      = pd.to_numeric(data['Amount'],      errors='coerce').fillna(0)
data['Days Overdue']= pd.to_numeric(data['Days Overdue'],errors='coerce')


# ── 2. GENERAR UN EXCEL POR CLIENTE ─────────────────────────────────────────

fecha_hoy = datetime.now().strftime("%d.%m.%Y")
HEADERS   = ['Inv Date', 'Due Date', 'Invoice', 'File', 'Currency', 'Amount', 'Days Overdue']

for cliente, grupo in data.groupby("Cliente:Trabajo"):

    cliente_clean = "".join(c for c in str(cliente) if c.isalnum() or c in (' ', '_')).strip()
    file_name     = f"Reporte_{cliente_clean}_{fecha_hoy}.xlsx"

    writer    = pd.ExcelWriter(file_name, engine='xlsxwriter',
                               engine_kwargs={'options': {'nan_inf_to_errors': True}})
    workbook  = writer.book
    worksheet = workbook.add_worksheet("Estado de Cuenta")

    # ── Formatos ──────────────────────────────────────────────────────────────
    header_fmt = workbook.add_format({
        'bg_color': '#92D050', 'bold': True, 'border': 1,
        'align': 'center', 'valign': 'vcenter'
    })
    header_fmt_low = workbook.add_format({
        'bg_color': '#C4C4C4', 'bold': True, 'border': 1,
        'align': 'center', 'valign': 'vcenter'
    })
    cell_fmt     = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter'})
    date_fmt     = workbook.add_format({'border': 1, 'num_format': 'dd/mm/yyyy',
                                        'align': 'center', 'valign': 'vcenter'})
    currency_fmt = workbook.add_format({'border': 1, 'num_format': '$#,##0.00',
                                        'align': 'center', 'valign': 'vcenter'})
    overdue_fmt  = workbook.add_format({'border': 1, 'font_color': 'red',
                                        'align': 'center', 'valign': 'vcenter'})
    total_label_fmt = workbook.add_format({'bold': True, 'border': 1,
                                           'align': 'center', 'valign': 'vcenter'})
    total_val_fmt   = workbook.add_format({'bold': True, 'border': 1,
                                           'num_format': '$#,##0.00',
                                           'align': 'center', 'valign': 'vcenter'})

    # ── Encabezado del cliente ────────────────────────────────────────────────
    worksheet.merge_range(0, 0, 0, 6, cliente, header_fmt)

    # ── Encabezados de columnas ───────────────────────────────────────────────
    for col_num, header in enumerate(HEADERS):
        worksheet.write(1, col_num, header, header_fmt_low)

    # ── Datos ─────────────────────────────────────────────────────────────────
    row_idx = 2
    for _, row in grupo.iterrows():
        if pd.isna(row['Inv Date']):
            continue

        worksheet.write(row_idx, 0, row['Inv Date'],    date_fmt)
        worksheet.write(row_idx, 1, row['Due Date'],    date_fmt)
        worksheet.write(row_idx, 2, row['Invoice'],     cell_fmt)
        worksheet.write(row_idx, 3, row['File'],        cell_fmt)
        worksheet.write(row_idx, 4, row['Currency'],    cell_fmt)
        worksheet.write(row_idx, 5, row['Amount'],      currency_fmt)

        val_overdue = row['Days Overdue']
        if pd.isna(val_overdue):
            worksheet.write(row_idx, 6, '', cell_fmt)
        else:
            fmt = overdue_fmt if val_overdue >= 0 else cell_fmt
            worksheet.write(row_idx, 6, val_overdue, fmt)

        row_idx += 1

    # ── Totales ───────────────────────────────────────────────────────────────
    total_portfolio = grupo['Amount'].sum()
    total_overdue   = grupo.loc[grupo['Days Overdue'] >= 0, 'Amount'].sum()

    start_total_row = row_idx + 2
    worksheet.write(start_total_row,     4, "Total Overdue", total_label_fmt)
    worksheet.write(start_total_row,     5, total_overdue,   total_val_fmt)
    worksheet.write(start_total_row + 1, 4, "Total",         total_label_fmt)
    worksheet.write(start_total_row + 1, 5, total_portfolio, total_val_fmt)

    # ── Auto-ajuste de columnas ───────────────────────────────────────────────
    for i, col in enumerate(HEADERS):
        max_len = max(
            grupo[col].astype(str).map(len).max() if not grupo[col].empty else 0,
            len(col)
        ) + 4
        worksheet.set_column(i, i, max_len)

    worksheet.set_column(0, 1, 14)
    worksheet.set_column(4, 5, 16)

    writer.close()
    print(f"✅ Generado: {file_name}")

print("\n🎉 Todos los estados de cuenta han sido generados.")
