import pandas as pd
import matplotlib.pyplot as plt
from openpyxl import Workbook
from openpyxl.drawing.image import Image
from openpyxl.chart import BarChart, Reference

def generar_reporte():
    # 1. Leer datos desde CSV
    df_produccion = pd.read_csv('../data/datos_produccion.csv')

    # 2. Calcular producción diaria
    produccion_diaria = df_produccion.groupby('fecha')['cantidad_producida'].sum()

    # 3. Crear y guardar gráfico
    plt.figure(figsize=(10, 6))
    plt.plot(produccion_diaria.index, produccion_diaria.values, marker='o')
    plt.title('Producción Diaria')
    plt.xlabel('Fecha')
    plt.ylabel('Cantidad Producida')
    plt.grid(True)
    plt.savefig('produccion_diaria.png')
    plt.close()  # Cerramos la figura para no saturar la memoria

    # 4. Crear un archivo Excel y escribir los datos
    wb = Workbook()
    ws = wb.active
    ws.title = "Reporte Producción"

    # 5. Escribir encabezados
    ws.append(["Fecha", "Cantidad Producida"])

    # 6. Volcar los datos al Excel
    for fecha, cantidad in produccion_diaria.items():
        ws.append([fecha, cantidad])

    # 7. Insertar el gráfico como imagen (si se desea)
    img = Image('produccion_diaria.png')
    # Vamos a poner la imagen en la celda E1
    ws.add_image(img, "E1")

    # 8. Guardar el archivo Excel
    wb.save("Reporte_Produccion.xlsx")
    print("Reporte_Produccion.xlsx generado exitosamente.")

if __name__ == "__main__":
    generar_reporte()
