import os
import csv

# ---- Campos esperados ----
ALUMNOS_FIELDS = ['No_Control', 'nombre', 'area', 'profesor']

DOCUMENTOS_FIELDS = [
    'No_Control',
    'Apellido_Paterno','Apellido_Materno','Nombre','Sexo','Telefono','Correo','Domicilio',
    'Carrera','Periodo','Semestre','Creditos','Dependencia','Domicilio_Dependencia',
    'Titular_Dependencia','Cargo_Responsable','Responsable_Proyecto','Nombre_Programa',
    'Modalidad_Externa','Modalidad_Interna','Fecha_Inicio','Fecha_Terminacion','Actividades',
    'TP_Edu_Adultos','TP_Desarrollo','TP_Deportivo','TP_Cultural','TP_Civico',
    'TP_Sustentable','TP_Salud','TP_Medio_Amb','TP_Otros',
    'Dia_Solicitud','Mes_Solicitud','Anio_Solicitud',
    # Fechas de bimestres
] + [f'{part}_{n}' for n in (1,2,3) for part in ('Dia1','Mes1','Anio1','Dia2','Mes2','Anio2')] + [
    'Nombre_supervisor','Puesto_supervisor'
] + [f'Actividad_{i}' for i in range(1,9)] + [
    'x1','x2','x3'
] + [f'resp{n}_{i}' for n in range(3,18) for i in range(5)] + [
    f'estu{n}_{i}' for n in range(3,18) for i in range(5)
] + [
    # Final
    'Municipio','Estado'
] + [f'Actividad{i}' for i in range(1,9)] + [f'Logro{i}' for i in range(1,9)] + [
    f'Aprendizaje{i}' for i in range(1,9)
] + [f'Beneficio{i}' for i in range(1,9)] + [
    'Nombre_Responsable','Cargo_Responsable','Num_Control',
    'Observaciones_Encargado','Observaciones_Estudiante'
]

def cargar_csv(path, expected_fields=None):
    rows = []
    if os.path.isfile(path):
        with open(path, newline='', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            for raw in reader:
                norm = {k.strip(): (v.strip() if v else "") for k,v in raw.items()}
                if expected_fields:
                    for fld in expected_fields:
                        norm.setdefault(fld,'')
                rows.append(norm)
    return rows

def guardar_csv(path, rows, fieldnames):
    try:
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(path,'w',newline='',encoding='utf-8') as f:
            w = csv.DictWriter(f, fieldnames=fieldnames)
            w.writeheader(); w.writerows(rows)
    except Exception as e:
        print(f"Error guardando CSV: {e}")
