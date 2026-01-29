import pandas as pd
from docxtpl import DocxTemplate
import os
from docx2pdf import convert
from mapeos import TITULO, ESTUDIANTE, MATERIAS_MAPEO

def generar_boletin(id_estudiante, ruta_excel, trimestre_a_imprimir):
    # --- 1. CARGA ---
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

    # Limpiador simple para etiquetas fijas
    def limpiar_fijo(tag):
        return tag.replace("{{", "").replace("}}", "").replace(" ", "").strip()

    # Limpiador para etiquetas con formato {t}
    def limpiar_dinamico(tag, t):
        # Primero aplicamos el formato y LUEGO limpiamos llaves y espacios
        tag_formateada = tag.format(t=t)
        return tag_formateada.replace("{{", "").replace("}}", "").replace(" ", "").strip()

    contexto = {}
    
    # --- 2. DATOS BÁSICOS ---
    primer_reg = info_est_notas.iloc[0]
    contexto[limpiar_fijo(ESTUDIANTE["NOMBRE"])] = str(primer_reg.get('StudentName', '')).strip()
    contexto[limpiar_fijo(ESTUDIANTE["GRADO"])] = str(primer_reg.get('HR', '')).strip()
    contexto[limpiar_fijo(ESTUDIANTE["PROFE"])] = str(primer_reg.get('HR_Teacher', '')).strip()
    contexto[limpiar_fijo(ESTUDIANTE["ID"])] = id_busqueda
    contexto[limpiar_fijo(TITULO["FINALREPORT"])] = f"FINAL REPORT TRIMESTER {trimestre_a_imprimir}"

    tipos_comentario = ["Strength", "Growth", "Goal", "Work Habits", "Participation", "Working in groups", "Behavior and school values"]

    # --- 3. PROCESAMIENTO ---
    for materia_nombre, items in MATERIAS_MAPEO.items():
        datos_m_notas = info_est_notas[info_est_notas['Subject'].str.strip() == materia_nombre]
        datos_m_coments = info_est_coments[info_est_coments['Subject'].str.strip() == materia_nombre]

        for nombre_item, etiqueta_base in items.items():
            
            # CASO A: PROFESOR
            if nombre_item == "Teacher":
                tag = limpiar_fijo(etiqueta_base)
                contexto[tag] = str(datos_m_notas.iloc[0].get('S_Teacher', '')) if not datos_m_notas.empty else ""

            # CASO B: COMENTARIOS (Etiqueta fija sin T)
            elif nombre_item in tipos_comentario:
                # Quitamos {t} porque el usuario quiere etiquetas fijas para comentarios
                tag_word = limpiar_fijo(etiqueta_base.replace("{t}", ""))
                if not datos_m_coments.empty:
                    col_excel = f"{nombre_item}_T{trimestre_a_imprimir}"
                    val = datos_m_coments.iloc[0].get(col_excel, '')
                    contexto[tag_word] = " ".join(str(val).split()) if pd.notna(val) else ""
                else:
                    contexto[tag_word] = ""

            # CASO C: NOTAS (Acumulativas con T1, T2, T3)
            # CASO C: NOTAS (Acumulativas con T1, T2, T3)
            else:
                for t in range(1, 4):
                    # CORRECCIÓN: Generamos la etiqueta limpia
                    # etiqueta_base es p.ej. "{{R_Lit_T{t}}}"
                    tag_nota = etiqueta_base.replace("{{", "").replace("}}", "").replace(" ", "").format(t=t)
                    
                    if t <= trimestre_a_imprimir:
                        if not datos_m_notas.empty:
                            # Buscamos el dominio comparando sin espacios y en minúsculas para asegurar
                            # nombre_item es la clave del diccionario (ej: "Literature & Information")
                            fila_dominio = datos_m_notas[datos_m_notas['Domain'].str.strip() == nombre_item.strip()]
                            
                            if not fila_dominio.empty:
                                col_trimestre = f'Trimester{t}'
                                valor = fila_dominio.iloc[0].get(col_trimestre, '')
                                contexto[tag_nota] = str(valor).strip() if pd.notna(valor) else ""
                                # print(f"✅ NOTA: {materia_nombre} - {nombre_item} -> {tag_nota} = {valor}") # Debug opcional
                            else:
                                contexto[tag_nota] = ""
                        else:
                            contexto[tag_nota] = ""
                    else:
                        contexto[tag_nota] = ""

    # --- 4. RENDERIZADO ---
    contexto_final = {k: v for k, v in contexto.items()}
    
    grado_val = contexto_final.get(limpiar_fijo(ESTUDIANTE["GRADO"]), "1")
    primer_caracter_grado = str(grado_val)[0]
    if primer_caracter_grado in ["1", "2"]:
        plantilla = "Grades1&2template.docx"
    elif primer_caracter_grado in ["3", "4", "5"]:
        plantilla = "Grades3,4&5template.docx"
    else:
        # Caso por defecto si el grado no coincide con los anteriores
        plantilla = "Grades3,4&5template.docx" 
        print(f"⚠️ Grado '{grado_val}' no reconocido, usando plantilla por defecto.")
    
    # doc = DocxTemplate(os.path.join("PLANTILLAS", plantilla))
    # doc.render(contexto_final, autoescape=True)
    
    # nombre_base = f"Boletin_{id_busqueda}"
    # doc.save(f"{nombre_base}.docx")
    # print(f"✅ Generado correctamente: {nombre_base}.docx")

    # 4. Generación y Conversión Final
    # Definimos el nombre base y las rutas completas usando la carpeta 'reportes'
    output_dir = "reportes"
    base_name = f"Boletin_{id_busqueda}"
    path_word = os.path.join(output_dir, f"{base_name}.docx")
    path_pdf = os.path.join(output_dir, f"{base_name}.pdf")

    # Cargamos la plantilla desde la carpeta PLANTILLAS
    doc = DocxTemplate(os.path.join("PLANTILLAS", plantilla))
    
    # Renderizamos los datos
    doc.render(contexto, autoescape=True)
    
    # Guardamos el archivo Word temporalmente en la carpeta 'reportes'
    doc.save(path_word)
    
    # Convertimos el Word recién guardado a PDF en la misma carpeta
    print(f"⏳ Convirtiendo {base_name} a PDF...")
    convert(path_word, path_pdf)
    
    # Borramos el Word para que no estorbe y solo quede el PDF
    if os.path.exists(path_pdf):
        os.remove(path_word)
        print(f"✅ Proceso completo: {path_pdf}")

def procesar_grado_1(ruta_excel, trimestre):
    # 1. Cargamos el Excel de notas
    df_notas = pd.read_excel(ruta_excel, sheet_name='Tablero_notas_Oficial')
    
    # Limpiamos nombres de columnas y el ID
    df_notas.columns = df_notas.columns.str.strip()
    df_notas['CodigoEstudiante'] = df_notas['CodigoEstudiante'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
    
    # 2. Filtramos solo los estudiantes de grado 1
    # Usamos .str.startswith('1') por si el grado es '1A', '1B', etc.
    estudiantes_grado_1 = df_notas[df_notas['HR'].astype(str).str.startswith('1')].copy()
    
    # Obtenemos la lista de IDs únicos (sin repetir)
    lista_ids = estudiantes_grado_1['CodigoEstudiante'].unique().tolist()
    
    total_a_generar = len(lista_ids)
    print(f"🚀 Iniciando generación para {total_a_generar} estudiantes de Grado 1...")

    # 3. Bucle para generar cada boletín
    for i, id_est in enumerate(lista_ids):
        # Llamamos a tu función original
        try:
            generar_boletin(id_est, ruta_excel, trimestre)
            
            # Cálculo de cuántos faltan
            faltantes = total_a_generar - (i + 1)
            if faltantes > 0:
                print(f"⏳ Faltan {faltantes} documentos por generar...")
            else:
                print("✅ ¡Todos los documentos han sido generados!")
                
        except Exception as e:
            print(f"❌ Error con el estudiante {id_est}: {e}")

# --- EJECUCIÓN ---
procesar_grado_1("baseprueba.xlsx", 1)

# === Ejecuciones ===
# EJECUTAR UN CODIGO, Y TRIMESTRE ESPECIFICO.

# generar_boletin("24771", "baseprueba.xlsx", 1)
# generar_boletin("20622", "Data Domains Elementary.xlsx", 1)