# Visor de negocios cautivos (Streamlit)

## Cómo ejecutarlo
1. Instala Python 3.10 o superior.
2. En esta carpeta, abre una terminal y ejecuta:
   ```
   pip install -r requirements.txt
   streamlit run app.py
   ```
3. Se abre el navegador en http://localhost:8501.

La app lee `BASE_DE_DATOS_PROYECTO_CAUTIVOS_migrada.xlsx` de esta misma carpeta.
También puedes subir otra versión del libro desde la barra lateral.

## Páginas
- **Resumen**: indicadores, completitud, conciliación, cotizaciones por cliente y margen base.
- **Cotizaciones**: filtros por cliente, proveedor, origen, destino, unidad, servicio y estado.
  Al seleccionar una fila se abre la ficha: servicios, cargos y condiciones, notas completas, proveedores, pendientes y trazabilidad a DB RAW.
- **Comparar alternativas**: alternativas del mismo cliente con la misma ruta y unidad, lado a lado (sin mezclar servicios).
- **Proveedores**: asignaciones y cotizaciones por proveedor.
- **Pendientes**: lo que requiere decisión, agrupado por tema.
- **Calidad de datos**: controles de cobertura e integridad y el diccionario de datos.

## Notas
- Los totales, la completitud y la conciliación se recalculan en la app a partir del detalle, con las mismas reglas que el Excel.
- Los importes se muestran sin símbolo de moneda porque el origen casi nunca la indica.
- La app solo lee el libro: las decisiones sobre pendientes se registran en el Excel (hoja PENDIENTES_REVISION, columna decision_tomada).
