import sys
import cv2
from pyzbar import pyzbar
import numpy as np
import pandas as pd

# Obtener la dirección del archivo seleccionado como argumento
selected_excel = sys.argv[1]

# Inicializar la cámara
cap = cv2.VideoCapture(0)

# Crear un conjunto para almacenar los códigos QR únicos detectados
seen_qr_codes = set()

# Crea un DataFrame vacío para almacenar todos los datos
all_data = pd.DataFrame()

# Leer el código QR
while True:
    _, frame = cap.read()
    decoded_objs = pyzbar.decode(frame)
    for obj in decoded_objs:
        data = obj.data.decode('utf-8')
        if data not in seen_qr_codes:
            print(f'Datos del código QR: {data}')

            # Leer los datos de la página web
            try:
                data_from_web = pd.read_html(data)
            except ValueError:
                print(f'Error al intentar leer los datos de la URL: {data}')
                continue

            # Concatenar cada conjunto de datos en all_data
            for df in data_from_web:
                df.reset_index(drop=True, inplace=True)
                if 'NIT DEL EMISOR' in df.columns:
                    df['NIT DEL EMISOR'] = df['NIT DEL EMISOR'].fillna(0).astype(int).replace(0, np.nan)
                if 0 in df.columns:
                    df = df.drop(columns=[0])
                if 1 in df.columns:
                    df = df.drop(columns=[1])
                # Descartar filas completamente vacías
                df = df.dropna(how='all')
                all_data = pd.concat([all_data, df], ignore_index=True)

            # Agregar este código QR a nuestro conjunto de códigos QR vistos
            seen_qr_codes.add(data)

        # Dibujar un rectángulo rojo alrededor del código QR
        cv2.rectangle(frame, (obj.rect.left, obj.rect.top), (obj.rect.left + obj.rect.width, obj.rect.top + obj.rect.height), (0, 0, 255), 2)

        # Mostrar el texto "QR reconocido" sobre el código QR
        cv2.putText(frame, "QR reconocido", (obj.rect.left, obj.rect.top - 10), cv2.FONT_HERSHEY_SIMPLEX, 0.5, (0, 0, 255), 2)

    cv2.imshow('frame', frame)
    
    # Después de cada bucle, escribe all_data en el archivo Excel seleccionado
    writer = pd.ExcelWriter(selected_excel, engine='openpyxl')
    all_data.to_excel(writer, sheet_name='All Data', index=False)
    writer.close()

    if cv2.waitKey(1) & 0xFF == ord('q'):
        break

cap.release()
cv2.destroyAllWindows()
