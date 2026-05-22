import pyodbc
import pandas as pd
from datetime import datetime

# ================================================
# CONEXIÓN A SQL SERVER
# ================================================
conn = pyodbc.connect(
    'DRIVER={SQL Server};'
    'SERVER=localhost;'
    'DATABASE=retail_analytics;'
    'Trusted_Connection=yes;'
)

# ================================================
# CONSULTAS
# ================================================

# 1. Resumen mensual
df_mensual = pd.read_sql("""
    SELECT nombre_mes, cantidad_ventas, ingresos, costos, ganancia, margen_pct
    FROM vw_resumen_mensual
    ORDER BY mes
""", conn)

# 2. Ranking de productos
df_productos = pd.read_sql("""
    SELECT p.nombre, cat.nombre AS categoria,
           SUM(vd.cantidad) AS unidades_vendidas,
           SUM(vd.subtotal_venta) AS ingresos,
           SUM(vd.subtotal_venta - vd.subtotal_costo) AS ganancia
    FROM venta_detalle vd
    JOIN productos p    ON vd.id_producto = p.id_producto
    JOIN categorias cat ON p.id_categoria = cat.id_categoria
    GROUP BY p.nombre, cat.nombre
    ORDER BY ganancia DESC
""", conn)

# 3. Ventas por cliente
df_clientes = pd.read_sql("""
    SELECT c.razon_social, c.pais,
           COUNT(v.id_venta) AS cantidad_ventas,
           SUM(v.total) AS total_facturado
    FROM clientes c
    JOIN ventas v ON c.id_cliente = v.id_cliente
    GROUP BY c.razon_social, c.pais
    ORDER BY total_facturado DESC
""", conn)

conn.close()

# ================================================
# EXPORTAR A EXCEL
# ================================================
fecha_hoy = datetime.now().strftime("%Y%m%d")
archivo = f"reporte_ventas_{fecha_hoy}.xlsx"

with pd.ExcelWriter(archivo, engine='openpyxl') as writer:
    df_mensual.to_excel(writer,   sheet_name='Resumen Mensual',  index=False)
    df_productos.to_excel(writer, sheet_name='Ranking Productos', index=False)
    df_clientes.to_excel(writer,  sheet_name='Ventas x Cliente',  index=False)

print(f"Reporte generado: {archivo}")
