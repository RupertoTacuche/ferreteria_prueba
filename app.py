import streamlit as st
from faker import Faker
import pandas as pd
import random
import numpy as np
from io import BytesIO

# Inicializar Faker con localización de México
fake = Faker('es_MX')

# --- Listas de datos realistas para el nicho de Ferretería ---
PRODUCTOS_DETALLE = [
    {"sku": "FER-001", "desc": "Pija punta de broca 4x1/2''", "cat": "Fijación", "precio": 0.85},
    {"sku": "FER-002", "desc": "Disco de corte para metal 4 1/2''", "cat": "Abrasivos", "precio": 25.00},
    {"sku": "FER-003", "desc": "Bulto de cemento Gris 50kg", "cat": "Materiales", "precio": 260.00},
    {"sku": "FER-004", "desc": "Martillo de uña curva 16 oz", "cat": "Herramientas Manuales", "precio": 145.00},
    {"sku": "FER-005", "desc": "Cinta teflón 1/2'' x 260''", "cat": "Plomería", "precio": 12.50},
    {"sku": "FER-006", "desc": "Tubo PVC sanitario 4'' x 6m", "cat": "Plomería", "precio": 380.00},
    {"sku": "FER-007", "desc": "Contacto dúplex armado blanco", "cat": "Eléctrico", "precio": 35.00},
    {"sku": "FER-008", "desc": "Cable cal. 12 THW negro 100m", "cat": "Eléctrico", "precio": 1250.00},
    {"sku": "FER-009", "desc": "Flexómetro Gripper 5 metros", "cat": "Herramientas Manuales", "precio": 85.00},
    {"sku": "FER-010", "desc": "Lija de agua grano 180", "cat": "Abrasivos", "precio": 9.00}
]

METODOS_PAGO = ["Efectivo", "Tarjeta de Crédito", "Transferencia SPEI", "EFECTIVO", "tarjeta", "  EFECTIVO  "]

# --- Funciones generadoras con inyección de "suciedad" ---
def get_dirty_vendedor():
    vendedor = fake.name()
    if random.random() < 0.20:
        return f"   {vendedor}  "  # Espacios fantasmas
    if random.random() < 0.15:
        return vendedor.upper()     # Inconsistencia de mayúsculas
    return vendedor

def get_dirty_cliente():
    if random.random() < 0.25:
        return "Mostrador"
    cliente = random.choice(["Constructora Alfa", "Ing. Rodríguez", "Plomería 'El Güero'", "Sistemas de Riego", "Arq. Silva", "Materiales San Juan"])
    if random.random() < 0.20:
        return f"  {cliente} "
    return cliente

def get_dirty_fecha():
    # Faker genera la fecha, nosotros variamos el formato str/datetime aleatoriamente
    dt = fake.date_between(start_date='-5m', end_date='today')
    if random.random() < 0.15:
        return dt.strftime("%d/%m/%Y")
    if random.random() < 0.10:
        return dt.strftime("%Y-%m-%d %H:%M:%S")
    return dt

# Mapeo de campos disponibles para Streamlit
available_fields = {
    'ID Venta': lambda: f"V-{random.randint(1000, 9999)}",
    'Fecha': get_dirty_fecha,
    'Código SKU': lambda: random.choice(PRODUCTOS_DETALLE)["sku"],
    'Descripción Producto': lambda: random.choice(PRODUCTOS_DETALLE)["desc"],
    'Categoría': lambda: random.choice(PRODUCTOS_DETALLE)["cat"],
    'Cantidad': lambda: f" {random.randint(1, 100)} " if random.random() < 0.15 else random.randint(1, 100),
    'Precio Unitario': lambda: random.choice(PRODUCTOS_DETALLE)["precio"],
    'Vendedor': get_dirty_vendedor,
    'Cliente': get_dirty_cliente,
    'Método de Pago': lambda: random.choice(METODOS_PAGO)
}

def generate_fake_data(fields, num_rows):
    # Generamos la estructura base celda por celda usando las funciones
    data = {field: [func() for _ in range(num_rows)] for field, func in fields.items()}
    df = pd.DataFrame(data)
    
    # Inyección secundaria de Suciedad: Celdas vacías (NaN) aleatorias en columnas clave
    for col in df.columns:
        if col in ['Vendedor', 'Cliente', 'Método de Pago', 'Fecha']:
            # 8% de probabilidad de dejar la celda vacía como un error de captura real
            df.loc[df.sample(frac=0.08).index, col] = np.nan
            
    return df

# --- Interfaz de Usuario de Streamlit ---
st.title('Generador de Datasets de prueba - Giro Ferretería 🛠️')
st.write('Configura tu dataset de prueba para demostrar el poder de tu herramienta de limpieza de datos.')

# Selección de campos con un default coherente para ferreterías
selected_fields = st.multiselect(
    'Selecciona los campos que incluirá el archivo:',
    options=list(available_fields.keys()),
    default=list(available_fields.keys())
)

num_rows = st.number_input(
    'Cantidad de registros de venta a generar:', 
    min_value=5, 
    max_value=1000, 
    value=50
)

if st.button('Generar Datos de Prueba'):
    if not selected_fields:
        st.warning('Por favor, selecciona al menos un campo.')
    else:
        # Filtrar funciones seleccionadas
        selected_funcs = {field: available_fields[field] for field in selected_fields}
        
        # Generar DataFrame
        df = generate_fake_data(selected_funcs, num_rows)
        
        # Mostrar vista previa en la app
        st.success(f'¡Se generaron {num_rows} registros con éxito!')
        st.write("### Vista previa del archivo sucio (Nota los espacios y formatos mixtos):")
        st.dataframe(df) # st.dataframe maneja mejor los NaN visualmente que st.write
        
        # Preparar buffer de descarga de Excel
        output = BytesIO()
        
        # Corrección del doble to_excel que tenías en tu código original
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            writer.book.use_constant_memory = True
            df.to_excel(writer, index=False, sheet_name="Registro_Ventas")
            
        output.seek(0)
        
        # Botón de descarga
        st.download_button(
            label='📥 Descargar Excel Sucio de Ferretería',
            data=output,
            file_name='ventas_ferreteria_sucias.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )