# 📊 Automatización de Estados de Cuenta — SBLogistics

Herramienta desarrollada para automatizar la generación de **Estados de Cuenta por cliente** a partir de reportes exportados de NetSuite (CargoWise), eliminando el proceso manual de formateo en Excel.

## 🚀 Impacto

| Antes | Después |
|-------|---------|
| ~2 horas de trabajo manual por ciclo | ~10 minutos (descarga + ejecución) |
| Verificación manual línea por línea | Validación automática de moneda, fechas y vencimientos |
| Un archivo genérico para todos | Reporte individual por cliente con formato profesional |

## ¿Qué hace?

- Lee el reporte `CUSTOM.xlsx` exportado desde NetSuite
- Limpia y estructura los datos automáticamente (fechas, monedas, clientes)
- Detecta facturas en **USD** vs **MXN** automáticamente por campo de nota
- Genera un archivo `.xlsx` por cliente con:
  - Encabezado con nombre del cliente
  - Tabla de facturas con fechas, montos y días de vencimiento
  - **Días vencidos en rojo** para facturas >= 0 días overdue
  - Totales: **Total Overdue** y **Total Portfolio**
  - Columnas auto-ajustadas al contenido
- Nombra cada archivo automáticamente: `Reporte_Cliente_DD.MM.YYYY.xlsx`

## 🛠️ Tecnologías

- Python 3
- `pandas` — limpieza y agrupación de datos
- `xlsxwriter` — generación de archivos Excel con formato
- Google Colab (entorno de ejecución)

## 📁 Estructura

```
automatizacion-estados-cuenta/
│
├── estado_de_cuenta.py     # Script principal
├── CUSTOM.xlsx             # (No incluido) Reporte exportado de NetSuite
└── README.md
```

## ▶️ Cómo usar

1. Exporta el reporte de cuentas por cobrar desde NetSuite como `CUSTOM.xlsx`
2. Sube el archivo a Google Colab (o colócalo en la misma carpeta que el script)
3. Instala la dependencia si es necesario:
   ```bash
   pip install xlsxwriter
   ```
4. Ejecuta el script:
   ```bash
   python estado_de_cuenta.py
   ```
5. Se generará un archivo Excel por cada cliente en la carpeta actual

## 📌 Notas

- El archivo de entrada debe llamarse `CUSTOM.xlsx` y ser el reporte estándar exportado de NetSuite
- Las facturas con valor en campo `Nota` son asignadas automáticamente como `USD`
- Los días de vencimiento negativos indican facturas aún no vencidas (se muestran en negro)
- Desarrollado y probado en Google Colab con Python 3.10+
