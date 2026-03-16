import pandas as pd
import os
import asyncio
from jinja2 import Environment, FileSystemLoader
from playwright.async_api import async_playwright
import time

from mapeos import TITULO, ESTUDIANTE, MATERIAS_MAPEO, MATERIAS_MAPEO_HS

# --- UTILIDADES DE LIMPIEZA ---
def limpiar_tag(tag):
    return str(tag).replace("{{", "").replace("}}", "").replace("{", "").replace("}", "").replace(" ", "").strip()

# --- 1. LÓGICA DE PREPARACIÓN DE CONTEXTO (MS/ELEMENTARY) ---
def preparar_contexto_ms(id_est, df_notas_total, df_coments_total, trimestre):
    info_notas = df_notas_total[df_notas_total['CodigoEstudiante'] == id_est]
    info_coments = df_coments_total[df_coments_total['CodigoEstudiante'] == id_est]
    
    if info_notas.empty: return None

    contexto = {}
    reg = info_notas.iloc[0]
    contexto[limpiar_tag(ESTUDIANTE["NOMBRE"])] = str(reg.get('StudentName', '')).strip()
    contexto[limpiar_tag(ESTUDIANTE["GRADO"])] = str(reg.get('HR', '')).strip()
    contexto[limpiar_tag(ESTUDIANTE["PROFE"])] = str(reg.get('HR_Teacher', '')).strip()
    contexto[limpiar_tag(ESTUDIANTE["ID"])] = id_est
    contexto[limpiar_tag(TITULO["FINALREPORT"])] = f"FINAL REPORT TRIMESTER {trimestre}"

    tipos_comentario = ["Strength", "Growth", "Goal", "Work Habits", "Participation", "Working in groups", "Behavior and school values"]

    for materia_nombre, items in MATERIAS_MAPEO.items():
        nombres_busqueda = list(materia_nombre) if isinstance(materia_nombre, (list, tuple)) else [materia_nombre]
        datos_m_notas = info_notas[info_notas['Subject'].str.strip().isin(nombres_busqueda)]
        datos_m_coments = info_coments[info_coments['Subject'].str.strip().isin(nombres_busqueda)]

        for nombre_item, etiqueta_base in items.items():
            if nombre_item == "Teacher":
                tag = limpiar_tag(etiqueta_base)
                contexto[tag] = str(datos_m_notas.iloc[0].get('S_Teacher', '')) if not datos_m_notas.empty else ""
            elif nombre_item in tipos_comentario:
                tag_html = limpiar_tag(etiqueta_base.replace("{t}", ""))
                if not datos_m_coments.empty:
                    val = datos_m_coments.iloc[0].get(f"{nombre_item}_T{trimestre}", '')
                    contexto[tag_html] = " ".join(str(val).split()) if pd.notna(val) else ""
                else:
                    contexto[tag_html] = ""
            else:
                for t in range(1, 4):
                    tag_nota = limpiar_tag(etiqueta_base.format(t=t))
                    if t <= trimestre and not datos_m_notas.empty:
                        fila_dom = datos_m_notas[datos_m_notas['Domain'].str.strip() == nombre_item.strip()]
                        val = fila_dom.iloc[0].get(f'Trimester{t}', '') if not fila_dom.empty else ""
                        # Línea a reemplazar para que el N/A SI SALGA en Elementary/MS
                        # Reemplaza la línea de contexto[tag_nota] por esta:
                        contexto[tag_nota] = str(val).strip() if str(val).lower() != 'nan' else ""
                    else:
                        contexto[tag_nota] = ""
    return contexto

# --- 2. LÓGICA DE PREPARACIÓN DE CONTEXTO (HIGH SCHOOL) ---
def preparar_contexto_hs(id_est, df_notas_total, df_coments_total, trimestre):
    info_notas = df_notas_total[df_notas_total['CodigoEstudiante'] == id_est]
    info_coments = df_coments_total[df_coments_total['CodigoEstudiante'] == id_est]
    
    if info_notas.empty: return None

    contexto = {}
    reg = info_notas.iloc[-1]
    contexto[limpiar_tag(ESTUDIANTE["NOMBRE"])] = str(reg.get('StudentName', ''))
    contexto[limpiar_tag(ESTUDIANTE["GRADO"])] = str(reg.get('HR', '9')).strip()
    contexto[limpiar_tag(ESTUDIANTE["PROFE"])] = str(reg.get('HR_Teacher', '')).strip()
    contexto[limpiar_tag(ESTUDIANTE["ID"])] = id_est
    contexto["Final_report"] = f"FINAL REPORT {trimestre}"

    columnas_ls = {"Work Habits": "HS_Work Habits_T{t}", "Participation": "Hs_Participation_T{t}", 
                   "Working in groups": "Hs_Working in groups_T{t}", "Behavior and school values": "Hs_Behavior and school values_T{t}", 
                   "comments": "Hs_Comment_T{t}"}

    for materia_key, items in MATERIAS_MAPEO_HS.items():
        if materia_key == "FINALREPORT": continue
        match_key = str(materia_key).lower().replace('.', '').strip()
        f_m = info_notas[info_notas['Subj_Match'] == match_key]
        f_c = info_coments[info_coments['Subj_Match'] == match_key]

        if f_m.empty: continue
        datos_m = f_m.iloc[0]

        for nombre_item, etiqueta_base in items.items():
            tag = limpiar_tag(etiqueta_base)
            if nombre_item == "nota":
                for t in range(1, 4):
                    tag_t = limpiar_tag(etiqueta_base.format(t=t))
                    val = datos_m.get(f"Trimester{t}", "") if t <= trimestre else ""
                    # Línea a reemplazar para que el N/A SI SALGA en notas de HS
                    contexto[tag_t] = str(val).replace('.0', '').strip() if pd.notna(val) and str(val).strip() != "" else ""
            elif nombre_item == "Teacher":
                contexto[tag] = str(datos_m.get('S_Teacher', '')).strip()
            elif nombre_item in columnas_ls:
                val_ls = ""
                if not f_c.empty:
                    col_target = columnas_ls[nombre_item].format(t=trimestre).lower()
                    for c_real in f_c.columns:
                        if c_real.lower().strip() == col_target:
                            val_raw = f_c.iloc[0].get(c_real, "")
                            # Línea a reemplazar para que el N/A SI SALGA en comentarios/LS de HS
                            val_ls = str(val_raw).strip() if pd.notna(val_raw) and str(val_raw).strip() != "" else ""
                            break
                contexto[tag] = val_ls
    return contexto

# --- 3. PROCESO PRINCIPAL OPTIMIZADO ---
async def procesar_boletines_completos(rutaPyMS, rutaHS, trimestre):
    inicio_total = time.time()
    print("📊 [1/3] Cargando bases de datos en memoria...")
    
    # Carga única de Excels
    df_pyms_notas = pd.read_excel(
        rutaPyMS,
        sheet_name='Tablero_notas_Oficial',
        keep_default_na=False
    )
    df_pyms_coments = pd.read_excel(rutaPyMS, sheet_name='LS_Comments')
    df_hs_notas = pd.read_excel(rutaHS, sheet_name='Destination_oficial')
    df_hs_coments = pd.read_excel(rutaHS, sheet_name='Ls_Comments_Oficial_HS')

    # Limpieza masiva vectorizada
    for df in [df_pyms_notas, df_pyms_coments, df_hs_notas, df_hs_coments]:
        df.columns = df.columns.str.strip()
        df['CodigoEstudiante'] = df['CodigoEstudiante'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
        if 'Subject' in df.columns:
            df['Subj_Match'] = df['Subject'].astype(str).str.lower().str.replace('.', '', regex=False).str.strip()

    # Preparar lista de tareas
    tareas = []
    for grado in range(1, 9):
        ids = df_pyms_notas[df_pyms_notas['HR'].astype(str).str.startswith(str(grado))]['CodigoEstudiante'].unique()
        for id_est in ids: tareas.append({'id': id_est, 'tipo': 'PyMS', 'grado': grado})
    
    for grado in range(9, 13):
        ids = df_hs_notas[df_hs_notas['HR'].astype(str).str.startswith(str(grado))]['CodigoEstudiante'].unique()
        for id_est in ids: tareas.append({'id': id_est, 'tipo': 'HS', 'grado': grado})

    total = len(tareas)
    print(f"🚀 [2/3] Se detectaron {total} estudiantes. Iniciando Playwright...")
    
    if not os.path.exists("reportes"): os.makedirs("reportes")

    # Iniciar Playwright una sola vez
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        page = await browser.new_page()
        env = Environment(loader=FileSystemLoader("plantillas_html"))

        print("📄 [3/3] Generando PDFs...")
        for i, tarea in enumerate(tareas):
            t_inicio_est = time.time()
            id_est = tarea['id']
            
            # Obtener contexto y plantilla según el tipo
            if tarea['tipo'] == 'PyMS':
                ctx = preparar_contexto_ms(id_est, df_pyms_notas, df_pyms_coments, trimestre)
                grado_val = str(ctx.get(limpiar_tag(ESTUDIANTE["GRADO"]), "1"))[0]
                if grado_val in ["1", "2"]: plant = "Grades1&2template.html"
                elif grado_val in ["3", "4", "5"]: plant = "Grades3,4&5template.html"
                elif grado_val in ["6"]: plant = "Grades6template.html"
                elif grado_val in ["7"]: plant = "Grades7template.html"
                else: plant = "Grades8template.html"
            else:
                ctx = preparar_contexto_hs(id_est, df_hs_notas, df_hs_coments, trimestre)
                
                # 1. Obtenemos el valor de la columna de grado usando la llave mapeada
                llave_grado = limpiar_tag(ESTUDIANTE["GRADO"])
                raw_hr = str(ctx.get(llave_grado, "9"))
                
                # 2. Extraemos solo los números (para manejar "10", "10A", "11th", etc.)
                import re
                match = re.search(r'(\d+)', raw_hr)
                hr_numero = match.group(1) if match else "9"
                
                # 3. Asignación de plantilla basada en el número extraído
                if hr_numero in ["9", "10", "11", "12"]:
                    plant = f"Grades{hr_numero}template.html"
                else:
                    plant = "Grades9template.html" # Fallback por si acaso

            if ctx:
                html_render = env.get_template(plant).render(ctx)
                await page.set_content(html_render)
                
                # Definimos el nombre del archivo con el trimestre
                # Ejemplo: Boletin_HS_28211_TRIMESTER_1.pdf
                nombre_pdf = f"Boletin_{tarea['tipo']}_{id_est}_TRIMESTER_{trimestre}.pdf"
                ruta_salida = os.path.join("reportes", nombre_pdf)

                await page.pdf(
                    path=ruta_salida,
                    format="Letter",
                    print_background=True,
                    margin={"top": "1cm", "bottom": "1cm", "left": "1cm", "right": "1cm"}
                )
                
                # Imprimimos el progreso incluyendo el trimestre en el mensaje de consola
                print(f"✅ [{i+1}/{total}] ID {id_est} (T{trimestre}) - {time.time()-t_inicio_est:.2f}s")

        await browser.close()

    final = time.time() - inicio_total
    print(f"\n{'='*30}\n✨ PROCESO FINALIZADO\n⏱️ Tiempo total: {final//60:.0f}m {final%60:.0f}s\n{'='*30}")

if __name__ == "__main__":
    asyncio.run(procesar_boletines_completos("Data Domains Elementary.xlsx", "Destination HS.xlsx", 2))