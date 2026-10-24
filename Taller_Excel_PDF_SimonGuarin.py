import openpyxl
from tabulate import tabulate
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

ARCHIVO_EXCEL = "ventas_tecnologia_v2.xlsx"
ARCHIVO_PDF = "reporte_ejecutivo_ventas.pdf"

encabezados = ["ID", "Producto", "Categoría", "Precio ($)", "Ventas (Unid)"]
alineacion = ("center", "left", "center", "right", "right")


# 1. Cargar el archivo Excel y leer los datos
def cargar_datos():
    wb = openpyxl.load_workbook(ARCHIVO_EXCEL, data_only=True)
    hoja = wb.active
    print(f"Hoja cargada exitosamente: {hoja.title}")

    datos = []
    for fila in hoja.iter_rows(min_row=2, values_only=True):
        # se omiten filas sin ID (filas vacias o de totales)
        if fila[0] is None:
            continue
        datos.append(list(fila))

    print(f"Total de registros leídos en memoria: {len(datos)}")
    return datos


# Reto 1: filtrar productos con precio mayor a $100
def filtrar_por_precio(datos, precio_minimo=100):
    filtrados = []
    for r in datos:
        if r[3] > precio_minimo:
            filtrados.append(r)
    return filtrados


# Reto 2: total de unidades vendidas y precio promedio
def calcular_metricas(datos):
    total_unidades = 0
    suma_precios = 0
    for r in datos:
        total_unidades = total_unidades + r[4]
        suma_precios = suma_precios + r[3]

    if len(datos) > 0:
        precio_promedio = suma_precios / len(datos)
    else:
        precio_promedio = 0
    return total_unidades, precio_promedio


# 2. Mostrar una tabla en consola con tabulate
def mostrar_tabla(datos, titulo):
    print(f"\n=== {titulo} ({len(datos)} PRODUCTOS) ===")
    print(tabulate(datos, headers=encabezados, tablefmt="fancy_grid", colalign=alineacion))


def mostrar_metricas(datos):
    total_unidades, precio_promedio = calcular_metricas(datos)
    resumen = [
        ["Total de unidades vendidas", total_unidades],
        ["Precio promedio", f"${precio_promedio:,.2f}"],
    ]
    print("\n=== MÉTRICAS DE RESUMEN ===")
    print(tabulate(resumen, tablefmt="fancy_grid", colalign=("left", "right")))


# Reto 3: buscar un producto por su ID
def buscar_por_id(datos):
    entrada = input("\nIngrese el ID del producto: ").strip()
    if not entrada.isdigit():
        print("[ADVERTENCIA] El ID debe ser un número entero.")
        return

    id_buscado = int(entrada)
    for r in datos:
        if r[0] == id_buscado:
            print(tabulate([r], headers=encabezados, tablefmt="fancy_grid", colalign=alineacion))
            return

    print(f"[ADVERTENCIA] No existe ningún producto con el ID {id_buscado}.")


# 3. Construir la tabla del PDF con una fila final de resumen (Reto 2)
def crear_tabla_pdf(datos, estilo_cuerpo, estilo_encabezado):
    encabezados_pdf = [Paragraph(f"<b>{h}</b>", estilo_encabezado) for h in encabezados]
    filas_pdf = [
        [
            Paragraph(str(r[0]), ParagraphStyle('C', parent=estilo_cuerpo, alignment=1)),
            Paragraph(str(r[1]), ParagraphStyle('L', parent=estilo_cuerpo, alignment=0)),
            Paragraph(str(r[2]), ParagraphStyle('C', parent=estilo_cuerpo, alignment=1)),
            Paragraph(f"${r[3]:,.2f}", ParagraphStyle('R', parent=estilo_cuerpo, alignment=2)),
            Paragraph(str(r[4]), ParagraphStyle('R', parent=estilo_cuerpo, alignment=2))
        ]
        for r in datos
    ]

    total_unidades, precio_promedio = calcular_metricas(datos)
    estilo_resumen = ParagraphStyle('Resumen', parent=estilo_cuerpo, fontName='Helvetica-Bold')
    fila_resumen = [
        Paragraph("", estilo_resumen),
        Paragraph("TOTAL / PROMEDIO", estilo_resumen),
        Paragraph("", estilo_resumen),
        Paragraph(f"${precio_promedio:,.2f}", ParagraphStyle('RR', parent=estilo_resumen, alignment=2)),
        Paragraph(str(total_unidades), ParagraphStyle('RR', parent=estilo_resumen, alignment=2))
    ]

    anchos_cols = [40, 160, 100, 80, 80]
    tabla_pdf = Table([encabezados_pdf] + filas_pdf + [fila_resumen], colWidths=anchos_cols, repeatRows=1)
    tabla_pdf.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1A365D")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E0")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -2), [colors.white, colors.HexColor("#F7FAFC")]),
        ("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#E2E8F0")),
        ("LINEABOVE", (0, -1), (-1, -1), 1, colors.HexColor("#1A365D")),
    ]))
    return tabla_pdf


# 4. Generar el informe ejecutivo en PDF
def exportar_pdf(datos):
    doc = SimpleDocTemplate(ARCHIVO_PDF, pagesize=letter)
    estilos = getSampleStyleSheet()

    estilo_titulo = ParagraphStyle(
        'EstiloTitulo', parent=estilos['Title'],
        fontName='Helvetica-Bold', fontSize=18,
        textColor=colors.HexColor("#1A365D"), spaceAfter=10
    )
    estilo_subtitulo = ParagraphStyle(
        'EstiloSubtitulo', parent=estilos['Heading2'],
        fontName='Helvetica-Bold', fontSize=13,
        textColor=colors.HexColor("#1A365D")
    )
    estilo_cuerpo = ParagraphStyle('EstiloCuerpo', parent=estilos['Normal'], fontSize=9)
    estilo_encabezado = ParagraphStyle(
        'EstiloEncabezado', parent=estilos['Normal'],
        fontName='Helvetica-Bold', fontSize=10,
        textColor=colors.whitesmoke, alignment=1
    )

    filtrados = filtrar_por_precio(datos)

    story = [
        Paragraph("Informe Ejecutivo de Ventas Tecno", estilo_titulo),
        Paragraph(f"<b>Total de Productos Procesados:</b> {len(datos)}", estilo_cuerpo),
        Spacer(1, 12),
        Paragraph("Catálogo completo", estilo_subtitulo),
        crear_tabla_pdf(datos, estilo_cuerpo, estilo_encabezado),
        PageBreak(),
        Paragraph("Productos con precio superior a $100", estilo_subtitulo),
        Paragraph(f"<b>Productos que cumplen el filtro:</b> {len(filtrados)}", estilo_cuerpo),
        Spacer(1, 8),
        crear_tabla_pdf(filtrados, estilo_cuerpo, estilo_encabezado),
    ]

    doc.build(story)
    print(f"\n[EXITO] Archivo '{ARCHIVO_PDF}' generado correctamente.")


# Reto 4: menu interactivo
def mostrar_menu():
    print("\n========== MENÚ TECNOVENTAS ==========")
    print("1) Ver catálogo completo")
    print("2) Filtrar por precio (> $100)")
    print("3) Buscar por ID")
    print("4) Exportar informe PDF")
    print("5) Salir del programa")


def main():
    datos = cargar_datos()

    while True:
        mostrar_menu()
        opcion = input("Seleccione una opción: ").strip()

        if opcion == "1":
            mostrar_tabla(datos, "REPORTE DE VENTAS EN CONSOLA")
            mostrar_metricas(datos)
        elif opcion == "2":
            filtrados = filtrar_por_precio(datos)
            mostrar_tabla(filtrados, "PRODUCTOS CON PRECIO MAYOR A $100")
            mostrar_metricas(filtrados)
        elif opcion == "3":
            buscar_por_id(datos)
        elif opcion == "4":
            exportar_pdf(datos)
        elif opcion == "5":
            print("Saliendo del programa...")
            break
        else:
            print("[ADVERTENCIA] Opción no válida, intente de nuevo.")


if __name__ == "__main__":
    main()
