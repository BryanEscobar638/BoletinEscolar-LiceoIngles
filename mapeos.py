TITULO = {
    "FINALREPORT": "<<[FINALREPORT]>>"
    # "FINALREPORT": "FINAL REPORT TRIMESTER 1"
}

ESTUDIANTE = {
    "NOMBRE": "<<[Student_Es].[StudentName]>>",
    "GRADO":  "<<[Student_Es].[HR]>>",
    "PROFE":  "<<[Student_Es].[HR_Teacher]>>",
    "ID":     "<<[Student_Es].[CodigoEstudiante]>>"
}

READING = {
    # ===== NOMBRE PROFE =====
    "NOMBREPROFE" : '<<ANY(SELECT(Tablero_notas_Oficial[S_Teacher], AND([CodigoEstudiante] = [_THISROW].[Student_Es].[CodigoEstudiante], [Subject] = "Reading")))>>',
    # ===== Literature & Information =====
    # TRIMESTRE 1
    "Literature&Information_T1": '<<ANY(SELECT(Tablero_notas_Oficial[Trimester1], AND([CodigoEstudiante] = [_THISROW].[Student_Es].[CodigoEstudiante], [Subject] = "Reading", [Domain] = "Literature & Information")))>>',
    # TRIMESTRE 2
    "Literature&Information_T2": '<<ANY(SELECT(Tablero_notas_Oficial[Trimester2], AND([CodigoEstudiante] = [_THISROW].[Student_Es].[CodigoEstudiante], [Subject] = "Reading", [Domain] = "Literature & Information")))>>',
    # TRIMESTRE 3
    "Literature&Information_T3": '<<ANY(SELECT(Tablero_notas_Oficial[Trimester3], AND([CodigoEstudiante] = [_THISROW].[Student_Es].[CodigoEstudiante], [Subject] = "Reading", [Domain] = "Literature & Information")))>>'
}