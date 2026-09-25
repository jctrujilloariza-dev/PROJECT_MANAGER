# App local — Centro de Control de Proyectos (Streamlit)

## Que es esto
Aplicacion local (no requiere servidor ni internet) para:
- Subir tu Excel/CSV actualizado.
- Ver Dashboard con avance y semaforo calculado automaticamente.
- Ver la Ficha de cada Proyecto y marcar el checklist con clics (no texto libre).
- Crear un proyecto nuevo y generar automaticamente sus 39 acciones estandar.
- Exportar de vuelta a un .xlsx con las 4 hojas ya calculadas, para integrarlo en tu archivo maestro.

## Instalacion (una sola vez)

1. Instala Python 3.11 o superior desde https://www.python.org/downloads/ (marca "Add to PATH" durante la instalacion).
2. Abre una terminal dentro de esta carpeta.
3. Ejecuta:
   ```
   pip install -r requirements.txt
   ```

## Uso (cada vez que quieras trabajar)

1. Abre una terminal dentro de esta carpeta.
2. Ejecuta:
   ```
   streamlit run app.py
   ```
3. Se abrira automaticamente en tu navegador en `http://localhost:8501`.
4. Cierra la terminal cuando termines: la app se apaga y no queda nada corriendo en segundo plano.

## Flujo de trabajo recomendado

1. Actualiza tu Excel de Facturacion / Control de Proyectos como siempre.
2. Abre la app (`streamlit run app.py`).
3. Pestana "Cargar datos" -> sube tu Excel o los CSV.
4. Revisa el Dashboard y marca el checklist en la Ficha de Proyecto.
5. Pestana "Exportar a Excel" -> descarga el archivo actualizado.
6. Copia esas hojas a tu Centro_Control_Proyectos.xlsx si quieres mantener ambos sincronizados.

## Siguiente version (cuando confirmes que esta funciona)

- Guardar automaticamente en un archivo local (sin tener que exportar cada vez).
- Grafico de avance por categoria (Oferta, Compras, Ejecucion...).
- Filtros por PM y por cliente en el Dashboard.
- Checklist especifico por proyecto (ademas del estandar de 39 acciones).
