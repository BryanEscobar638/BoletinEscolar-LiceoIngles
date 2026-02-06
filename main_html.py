import pandas as pd
import os
import asyncio # Necesario para Playwright
from jinja2 import Environment, FileSystemLoader
from playwright.async_api import async_playwright # Reemplaza a pisa
import time  # Para medir el tiempo

from mapeos import TITULO, ESTUDIANTE, MATERIAS_MAPEO

# Cambiamos la función a 'async' para que Playwright funcione correctamente
async def generar_boletin(id_estudiante, ruta_excel, trimestre_a_imprimir):
    # --- 1. CARGA DE DATOS --- (Se mantiene igual)
    df_notas = pd.read_excel(ruta_excel, sheet_name='Tablero_notas_Oficial')
    df_comentarios = pd.read_excel(ruta_excel, sheet_name='Ls_Comments_Oficial')
    
    for df in [df_notas, df_comentarios]:
        df.columns = df.columns.str.strip()
        df['CodigoEstudiante'] = df['CodigoEstudiante'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()

    id_busqueda = str(id_estudiante).strip()
    info_est_coments = df_comentarios[df_comentarios['CodigoEstudiante'] == id_busqueda]
    info_est_notas = df_notas[df_notas['CodigoEstudiante'] == id_busqueda]

    if info_est_notas.empty:
        print(f"❌ No se encontraron datos para el ID: {id_busqueda}")
        return

    def limpiar_fijo(tag):
        return tag.replace("{{", "").replace("}}", "").replace(" ", "").strip()

    contexto = {}
    
    # --- 2. LLENADO DE CONTEXTO --- (Se mantiene igual)
    primer_reg = info_est_notas.iloc[0]
    contexto[limpiar_fijo(ESTUDIANTE["NOMBRE"])] = str(primer_reg.get('StudentName', '')).strip()
    contexto[limpiar_fijo(ESTUDIANTE["GRADO"])] = str(primer_reg.get('HR', '')).strip()
    contexto[limpiar_fijo(ESTUDIANTE["PROFE"])] = str(primer_reg.get('HR_Teacher', '')).strip()
    contexto[limpiar_fijo(ESTUDIANTE["ID"])] = id_busqueda
    contexto[limpiar_fijo(TITULO["FINALREPORT"])] = f"FINAL REPORT TRIMESTER {trimestre_a_imprimir}"

    tipos_comentario = ["Strength", "Growth", "Goal", "Work Habits", "Participation", "Working in groups", "Behavior and school values"]

    for materia_nombre, items in MATERIAS_MAPEO.items():
        datos_m_notas = info_est_notas[info_est_notas['Subject'].str.strip() == materia_nombre]
        datos_m_coments = info_est_coments[info_est_coments['Subject'].str.strip() == materia_nombre]

        for nombre_item, etiqueta_base in items.items():
            if nombre_item == "Teacher":
                tag = limpiar_fijo(etiqueta_base)
                contexto[tag] = str(datos_m_notas.iloc[0].get('S_Teacher', '')) if not datos_m_notas.empty else ""
            elif nombre_item in tipos_comentario:
                tag_html = limpiar_fijo(etiqueta_base.replace("{t}", ""))
                if not datos_m_coments.empty:
                    col_excel = f"{nombre_item}_T{trimestre_a_imprimir}"
                    val = datos_m_coments.iloc[0].get(col_excel, '')
                    contexto[tag_html] = " ".join(str(val).split()) if pd.notna(val) else ""
                else:
                    contexto[tag_html] = ""
            else:
                for t in range(1, 4):
                    tag_nota = etiqueta_base.replace("{{", "").replace("}}", "").replace(" ", "").format(t=t)
                    if t <= trimestre_a_imprimir:
                        if not datos_m_notas.empty:
                            fila_dominio = datos_m_notas[datos_m_notas['Domain'].str.strip() == nombre_item.strip()]
                            if not fila_dominio.empty:
                                col_trimestre = f'Trimester{t}'
                                valor = fila_dominio.iloc[0].get(col_trimestre, '')
                                contexto[tag_nota] = str(valor).strip() if pd.notna(valor) else ""
                            else:
                                contexto[tag_nota] = ""
                        else:
                            contexto[tag_nota] = ""
                    else:
                        contexto[tag_nota] = ""

    # --- 3. RENDERIZADO HTML ---
    try:
        env = Environment(loader=FileSystemLoader("plantillas_html"))
        grado_val = contexto.get(limpiar_fijo(ESTUDIANTE["GRADO"]), "1")
        primer_caracter = str(grado_val)[0]
        nombre_plantilla = ""
        if primer_caracter in ["1", "2"]:
            nombre_plantilla = "Grades1&2template.html"
        elif primer_caracter in ["3", "4", "5"]:
            # Este actúa como el caso por defecto para 3, 4 y 5
            nombre_plantilla = "Grades3,4&5template.html"
        elif primer_caracter in ["6"]:
            nombre_plantilla = "Grades6template.html"
        template = env.get_template(nombre_plantilla)
        html_renderizado = template.render(contexto)
    except Exception as e:
        print(f"❌ Error al cargar la plantilla HTML: {e}")
        return

    # --- 4. GENERACIÓN DE PDF CON PLAYWRIGHT ---
    output_dir = "reportes"
    if not os.path.exists(output_dir): os.makedirs(output_dir)
    path_pdf = os.path.join(output_dir, f"Boletin_{id_busqueda}.pdf")

    try:
        async with async_playwright() as p:
            # Lanzamos el navegador (Chromium)
            browser = await p.chromium.launch()
            page = await browser.new_page()
            
            # Cargamos el contenido HTML
            await page.set_content(html_renderizado)
            
            # Generamos el PDF
            # 'print_background' es clave para que salgan los colores de tus tablas
            await page.pdf(
                path=path_pdf,
                format="Letter",
                print_background=True,
                margin={"top": "1cm", "bottom": "1cm", "left": "1cm", "right": "1cm"}
            )
            
            await browser.close()
        print(f"✅ PDF generado con éxito con Playwright: {path_pdf}")
            
    except Exception as e:
        print(f"❌ Error crítico al generar PDF con Playwright: {e}")

async def procesar_grado_1(ruta_excel, trimestre):
    # --- Inicio del cronómetro global ---
    tiempo_inicio_total = time.time()

    # 1. Carga y limpieza inicial
    print("📊 Cargando base de datos...")
    df_notas = pd.read_excel(ruta_excel, sheet_name='Tablero_notas_Oficial')
    df_coments = pd.read_excel(ruta_excel, sheet_name='Ls_Comments_Oficial')
    
    for df in [df_notas, df_coments]:
        df.columns = df.columns.str.strip()
        df['CodigoEstudiante'] = df['CodigoEstudiante'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
    
    # 2. Filtrado de IDs únicos de grado 1
    estudiantes_grado_1 = df_notas[df_notas['HR'].astype(str).str.startswith('1')].copy()
    lista_ids = estudiantes_grado_1['CodigoEstudiante'].unique().tolist()
    
    total_a_generar = len(lista_ids)
    print(f"🚀 Iniciando generación de {total_a_generar} boletines PDF con Playwright...")

    # 3. BUCLE DE GENERACIÓN
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        
        for i, id_est in enumerate(lista_ids):
            # --- Inicio cronómetro por estudiante ---
            tiempo_inicio_est = time.time()
            
            try:
                # Llamada a la función principal
                await generar_boletin(id_est, ruta_excel, trimestre)
                
                # --- Cálculo de tiempo por estudiante ---
                tiempo_est = time.time() - tiempo_inicio_est
                faltantes = total_a_generar - (i + 1)
                
                print(f"✅ [{i+1}/{total_a_generar}] ID {id_est} generado en {tiempo_est:.2f} segundos.")
                
                if faltantes % 5 == 0 and faltantes > 0:
                    print(f"⏳ Faltan {faltantes} archivos por procesar...")
                    
            except Exception as e:
                print(f"❌ Error con el estudiante {id_est}: {e}")

        await browser.close()

    # --- Fin del cronómetro global ---
    tiempo_total = time.time() - tiempo_inicio_total
    minutos = int(tiempo_total // 60)
    segundos = int(tiempo_total % 60)

    print(f"\n{'='*40}")
    print("✅ ¡PROCESO TOTALMENTE TERMINADO!")
    print(f"⏱️ Tiempo total de ejecución: {minutos} min {segundos} seg")
    print(f"📂 Revisa la carpeta 'reportes'.")
    print(f"{'='*40}")

# === Ejecución (Cambia un poco por ser asíncrono) ===
if __name__ == "__main__":
    asyncio.run(generar_boletin("23112", "baseprueba.xlsx", 1))
    # asyncio.run(procesar_grado_1("baseprueba.xlsx", 1))