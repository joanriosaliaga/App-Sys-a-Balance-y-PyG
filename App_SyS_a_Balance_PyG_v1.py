import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import xlsxwriter

def clasificar_cuenta(cuenta, valor):
    """
    Clasifica una cuenta contable según las reglas del código VBA original
    """
    cuenta = str(cuenta).strip()
    prefijo2 = cuenta[:2] if len(cuenta) >= 2 else cuenta
    prefijo3 = cuenta[:3] if len(cuenta) >= 3 else prefijo2
    prefijo4 = cuenta[:4] if len(cuenta) >= 4 else prefijo3
    
    grupo = ""
    
    # 1) Clasificación específica de 4 dígitos
    clasificacion_4_digitos = {
        # Activo/Pasivo corriente - Inversiones/Deudas con empresas del grupo
        ("5523", "5524", "5525"): lambda v: "Activo corriente - Inversiones en empresas del grupo y asociadas a corto plazo" if v > 0 else "Pasivo corriente - Deudas con empresas del grupo y asociadas a corto plazo",
        
        # Pasivo corriente - Acreedores comerciales
        ("4750", "4751", "4752", "4758"): "Pasivo corriente - Acreedores comerciales y otras cuentas a pagar",
        
        # Activo no corriente - Inversiones empresas del grupo a largo plazo
        ("2403", "2404", "2493", "2494", "2423", "2424", "2413", "2414", "2943", "2944", "2953", "2954"): "Activo no corriente - Inversiones en empresas del grupo y asociadas a largo plazo",
        
        # Activo no corriente - Inversiones financieras a largo plazo
        ("2405", "2495", "2425", "2415", "2955", "2945"): "Activo no corriente - Inversiones financieras a largo plazo",
        
        # Activo corriente - Inversiones empresas del grupo a corto plazo
        ("5303", "5304", "5393", "5394", "5323", "5324", "5343", "5344", "5313", "5314", "5333", "5334", "5353", "5354", "5953", "5954", "5943", "5944"): "Activo corriente - Inversiones en empresas del grupo y asociadas a corto plazo",
        
        # Activo corriente - Inversiones financieras a corto plazo
        ("5305", "5395", "5325", "5345", "5315", "5335", "5355", "5945", "5955", "5590", "5593"): "Activo corriente - Inversiones financieras a corto plazo",
        
        # Pasivo corriente - Deudas a corto plazo
        ("5105", "5125", "5115", "5135", "5145", "5595", "5598", "5565", "5566", "1034", "1044", "5530", "5532"): "Pasivo corriente - Deudas a corto plazo",
        
        # Pasivo corriente - Deudas con empresas del grupo a corto plazo
        ("5103", "5104", "5113", "5114", "5123", "5124", "5133", "5134", "5143", "5144", "5563", "5564"): "Pasivo corriente - Deudas con empresas del grupo y asociadas a corto plazo",
        
        # Activo corriente - Deudores comerciales
        ("5531", "5533", "5580", "4708", "4709", "4933", "4934", "4935", "4700"): "Activo corriente - Deudores comerciales y otras cuentas a cobrar",
        
        # Pasivo no corriente - Deudas a largo plazo
        ("1605", "1625", "1615", "1635"): "Pasivo no corriente - Deudas a largo plazo",
        
        # Pasivo no corriente - Deudas con empresas del grupo a largo plazo
        ("1603", "1604", "1613", "1614", "1623", "1624", "1633", "1634"): "Pasivo no corriente - Deudas con empresas del grupo y asociadas a largo plazo",
        
        # PyG - Variación de existencias
        ("6930", "7930"): "Variación de existencias de productos terminados y en curso de fabricación",
        
        # PyG - Aprovisionamientos
        ("6931", "6932", "6933", "7931", "7932", "7933", "6061", "6062", "6081", "6082", "6091", "6092"): "Aprovisionamientos",
        
        # PyG - Excesos de provisiones
        ("7951", "7952", "7955", "7956"): "Excesos de provisiones",
        
        # PyG - Ingresos financieros
        ("7600", "7601", "7602", "7603", "7610", "7611", "7612", "7613", "7620", "7621"): "Ingresos financieros",
        
        # PyG - Gastos financieros
        ("6610", "6611", "6612", "6613", "6615", "6616", "6617", "6618", "6620", "6621", "6622", "6623", "6624", "6640", "6641", "6642", "6643", "6651", "6652", "6653", "6654", "6655", "6656", "6657"): "Gastos financieros",
        
        # PyG - Variación valor razonable
        ("6630", "6631", "6632", "6633", "7630", "7631", "7632", "7633"): "Variación de valor razonable en instrumentos financieros",
        
        # PyG - Impuestos sobre beneficios
        ("6300", "6301"): "Impuestos sobre beneficios",
        
        # Balance - Capital
        ("1030", "1040"): "Capital",
        
        # Balance - Otras reservas
        ("1140", "1142", "1143", "1144"): "Otras reservas",
        
        # Balance - Reserva legal
        ("1141",): "Reserva Legal y estatutarias",
        
        # Balance - Ajustes por cambio de valor
        ("1340",): "Ajustes por cambio de valor",
        
        # PyG - Gastos de personal
        ("6457", "7957"): "Gastos de personal",
        
        # PyG - Otros gastos de explotación
        ("7954",): "Otros gastos de explotación"
    }
    
    for cuentas, clasificacion in clasificacion_4_digitos.items():
        if prefijo4 in cuentas:
            if callable(clasificacion):
                grupo = clasificacion(valor)
            else:
                grupo = clasificacion
            break
    
    # 2) Clasificación de 3 dígitos si no se encontró en 4 dígitos
    if not grupo:
        clasificacion_3_digitos = {
            # Inversiones/Deudas con empresas del grupo a corto plazo
            ("551", "552", "554"): lambda v: "Activo corriente - Inversiones en empresas del grupo y asociadas a corto plazo" if v > 0 else "Pasivo corriente - Deudas con empresas del grupo y asociadas a corto plazo",
            
            # Pasivo corriente - Acreedores comerciales
            ("465", "466", "476", "477", "438"): "Pasivo corriente - Acreedores comerciales y otras cuentas a pagar",
            
            # Activo no corriente - Activos por impuesto diferido
            ("474",): "Activo no corriente - Activos por impuesto diferido",
            
            # Activo corriente - Periodificaciones a corto plazo
            ("480", "567"): "Activo corriente - Periodificaciones a corto plazo",
            
            # Activo corriente - Activos no corrientes mantenidos para la venta
            ("580", "581", "582", "583", "584", "599"): "Activo corriente - Activos no corrientes mantenidos para la venta",
            
            # Activo no corriente - Inversiones empresas del grupo a largo plazo
            ("255", "293", "294"): "Activo no corriente - Inversiones en empresas del grupo y asociadas a largo plazo",
            
            # Activo no corriente - Inversiones financieras a largo plazo
            ("298", "297", "258", "251", "250", "259", "252", "253", "254"): "Activo no corriente - Inversiones financieras a largo plazo",
            
            # Activo corriente - Inversiones financieras a corto plazo
            ("540", "541", "542", "543", "545", "546", "547", "548", "549", "565", "566", "598", "597"): "Activo corriente - Inversiones financieras a corto plazo",
            
            # Inmovilizado
            ("280", "290"): "Activo no corriente - Inmovilizado Intangible",
            ("281", "291"): "Activo no corriente - Inmovilizado Material",
            ("282", "292"): "Activo no corriente - Inversiones Inmobiliarias",
            
            # Patrimonio Neto
            ("100", "101", "102"): "Capital",
            ("110",): "Prima de emisión",
            ("112",): "Reserva Legal y estatutarias",
            ("113", "115", "119"): "Otras reservas",
            ("108", "109"): "Acciones y participaciones en patrimonio propias",
            ("120",): "Remanente",
            ("121",): "Resultados de ejercicios anteriores",
            ("118",): "Otras aportaciones de socios",
            ("129",): "Resultado del ejercicio",
            ("557",): "Dividendo a cuenta",
            ("111",): "Otros instrumentos de patrimonio neto",
            ("133", "137"): "Ajustes por cambio de valor",
            ("130", "131", "132"): "Subvenciones, donaciones y legados recibidos",
            
            # Pasivo no corriente
            ("180", "185", "189"): "Pasivo no corriente - Deudas a largo plazo",
            ("479",): "Pasivo no corriente - Pasivos por impuesto diferido",
            ("181",): "Pasivo no corriente - Periodificaciones a largo plazo",
            
            # Pasivo corriente
            ("400", "401", "403", "404", "405", "406", "475"): "Pasivo corriente - Acreedores comerciales y otras cuentas a pagar",
            ("585", "586", "587", "588", "589"): "Pasivo corriente - Pasivos vinculados con activos no corrientes mantenidos para la venta",
            ("499", "529"): "Pasivo corriente - Provisiones a corto plazo",
            ("485", "568"): "Pasivo corriente - Periodificaciones a corto plazo",
            ("593",): "Activo corriente - Inversiones en empresas del grupo y asociadas a corto plazo",
            ("560", "561", "555", "500", "501", "505", "506", "520", "527", "524", "509", "190", "192", "525", "526", "521", "522", "523", "569", "194", "528"): "Pasivo corriente - Deudas a corto plazo",
            
            # Activo corriente
            ("407",): "Activo corriente - Existencias",
            ("430", "431", "432", "433", "434", "435", "436", "437", "460", "470", "471", "472", "490", "493", "544"): "Activo corriente - Deudores comerciales y otras cuentas a cobrar",
            
            # PyG - Ingresos
            ("700", "701", "702", "703", "704", "705", "706", "708", "709"): "Importe neto de la cifra de negocios",
            ("740", "747", "750"): "Otros ingresos de explotación",
            ("760", "767", "769", "778"): "Ingresos financieros",
            
            # PyG - Gastos
            ("600", "601", "602", "607", "606", "608", "609", "610", "611", "612"): "Aprovisionamientos",
            ("640", "641", "642", "643", "644", "645", "649", "795"): "Gastos de personal",
            ("660", "661", "662", "664", "665", "669"): "Gastos financieros",
            ("680", "681", "682"): "Amortización del inmovilizado",
            ("670", "671", "672", "690", "691", "692", "790", "770", "771", "772", "791", "792"): "Deterioro y resultado por enajenaciones del inmovilizado",
            ("666", "667", "673", "675", "678", "696", "697", "698", "699", "766", "773", "775", "796", "797", "798", "799"): "Deterioro y resultado por enajenaciones de instrumentos financieros",
            ("668", "768"): "Diferencias de cambio",
            ("630", "633", "638"): "Impuestos sobre beneficios"
        }
        
        # Tratamiento especial para algunos prefijos de 3 dígitos
        grupos_especiales_3 = ["62", "631", "634", "650", "651", "659", "695", "636", "639", "694", "794"]
        if prefijo3 in grupos_especiales_3:
            grupo = "Otros gastos de explotación"
        else:
            for cuentas, clasificacion in clasificacion_3_digitos.items():
                if prefijo3 in cuentas:
                    if callable(clasificacion):
                        grupo = clasificacion(valor)
                    else:
                        grupo = clasificacion
                    break
    
    # 3) Clasificación genérica de 2 dígitos
    if not grupo:
        clasificacion_2_digitos = {
            "20": "Activo no corriente - Inmovilizado Intangible",
            "21": "Activo no corriente - Inmovilizado Material",
            "22": "Activo no corriente - Inversiones Inmobiliarias",
            "23": "Activo no corriente - Inmovilizado Material",
            "24": "Activo no corriente - Inversiones financieras a largo plazo",
            "25": "Activo no corriente - Inversiones financieras a largo plazo",
            "26": "Activo no corriente - Inversiones financieras a largo plazo",
            "28": "Activo - Otros",
            "29": "Activo - Otros",
            "46": "Activo - Otros",
            "47": "Pasivo corriente - Acreedores comerciales y otras cuentas a pagar",
            "57": "Activo corriente - Tesorería",
            "10": "Patrimonio Neto",
            "11": "Patrimonio Neto",
            "12": "Patrimonio Neto",
            "13": "Patrimonio Neto",
            "14": "Pasivo no corriente - Provisiones a largo plazo",
            "19": "Pasivo corriente - Deudas a corto plazo",
            "70": "Importe neto de la cifra de negocios",
            "71": "Variación de existencias de productos terminados y en curso de fabricación",
            "73": "Trabajos realizados por la empresa para su activo",
            "74": "Otros ingresos de explotación",
            "75": "Otros ingresos de explotación",
            "76": "Ingresos financieros",
            "60": "Aprovisionamientos",
            "61": "Aprovisionamientos",
            "62": "Otros gastos de explotación",
            "63": "Otros gastos de explotación",
            "64": "Gastos de personal",
            "65": "Otros gastos de explotación",
            "66": "Gastos financieros",
            "68": "Amortización del inmovilizado"
        }
        
        # Rangos especiales
        if prefijo2 in ["30", "31", "32", "33", "34", "35", "36", "37", "38", "39"]:
            grupo = "Activo corriente - Existencias"
        elif prefijo2 in ["43", "44", "45", "49"]:
            grupo = "Activo corriente - Deudores comerciales y otras cuentas a cobrar"
        elif prefijo2 in ["40", "41", "42"]:
            grupo = "Pasivo corriente - Acreedores comerciales y otras cuentas a pagar"
        elif prefijo2 in ["50", "51", "52"]:
            grupo = "Pasivo corriente - Deudas a corto plazo"
        elif prefijo2 in ["15", "16", "17", "18"]:
            grupo = "Pasivo no corriente - Deudas a largo plazo"
        elif prefijo2 in ["67", "69"]:
            grupo = "Otros gastos"
        elif prefijo2 in ["77", "78", "79"]:
            grupo = "Otros ingresos de explotación"
        else:
            grupo = clasificacion_2_digitos.get(prefijo2, "Otros")
    
    return grupo

def ajustar_signo(valor, grupo, prefijo2):
    """
    Ajusta el signo del valor según las reglas contables
    """
    adjusted_valor = valor
    
    # Grupos que deben cambiar de signo (están en el haber)
    grupos_cambio_signo = [
        "Patrimonio Neto", "Capital", "Prima de emisión", "Reserva Legal y estatutarias", 
        "Otras reservas", "Acciones y participaciones en patrimonio propias", "Remanente", 
        "Resultados de ejercicios anteriores", "Otras aportaciones de socios", 
        "Resultado del ejercicio", "Dividendo a cuenta", "Otros instrumentos de patrimonio neto", 
        "Ajustes por cambio de valor", "Subvenciones, donaciones y legados recibidos"
    ]
    
    # Todos los grupos de pasivo
    if "Pasivo" in grupo or grupo in grupos_cambio_signo:
        adjusted_valor = -valor
    
    # PyG: todas las cuentas de ingresos y gastos cambian de signo
    if prefijo2 >= "60" and prefijo2 <= "79":
        adjusted_valor = -valor
    
    return adjusted_valor

def procesar_sys(df):
    """
    Procesa el DataFrame de Sumas y Saldos y lo clasifica en Balance y PyG
    """
    # Buscar la columna de Saldo Definitivo
    saldo_col = None
    for col in df.columns:
        if 'saldo' in str(col).lower() and 'definitivo' in str(col).lower():
            saldo_col = col
            break
    
    if saldo_col is None:
        # Buscar alternativas
        for col in df.columns:
            if 'saldo' in str(col).lower():
                saldo_col = col
                break
    
    if saldo_col is None:
        st.error("No se encontró la columna 'Saldo Definitivo'. Verifica que el archivo tenga la estructura correcta.")
        return None, None
    
    # Limpiar datos
    df_clean = df.dropna(subset=[df.columns[0], saldo_col])  # Eliminar filas sin cuenta o saldo
    df_clean = df_clean[df_clean[saldo_col] != 0]  # Eliminar saldos cero
    
    datos_balance = {}
    datos_pyg = {}
    
    for _, row in df_clean.iterrows():
        cuenta = str(row[df.columns[0]]).strip()
        descripcion = str(row[df.columns[1]]).strip() if len(df.columns) > 1 else ""
        try:
            valor = float(row[saldo_col])
        except:
            continue
        
        if valor == 0:
            continue
        
        grupo = clasificar_cuenta(cuenta, valor)
        prefijo2 = cuenta[:2] if len(cuenta) >= 2 else cuenta
        
        # Ajustar signo
        adjusted_valor = ajustar_signo(valor, grupo, prefijo2)
        
        # Clasificar en Balance o PyG
        if prefijo2 >= "60" and prefijo2 <= "79":
            # PyG
            if grupo not in datos_pyg:
                datos_pyg[grupo] = []
            datos_pyg[grupo].append([cuenta, descripcion, adjusted_valor])
        else:
            # Balance
            if grupo not in datos_balance:
                datos_balance[grupo] = []
            datos_balance[grupo].append([cuenta, descripcion, adjusted_valor])
    
    return datos_balance, datos_pyg

def crear_excel_balance_pyg(datos_balance, datos_pyg):
    """
    Crea un archivo Excel con las hojas Balance, PyG y Resumen
    """
    output = BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    
    # Formatos
    bold_format = workbook.add_format({'bold': True})
    italic_format = workbook.add_format({'italic': True})
    header_format = workbook.add_format({'bold': True, 'bg_color': '#DCDCDC'})
    number_format = workbook.add_format({'num_format': '#,##0.00'})
    bold_number_format = workbook.add_format({'bold': True, 'num_format': '#,##0.00'})
    
    # HOJA BALANCE
    ws_balance = workbook.add_worksheet('Balance')
    
    orden_balance = [
        "Activo no corriente - Inmovilizado Intangible",
        "Activo no corriente - Inmovilizado Material", 
        "Activo no corriente - Inversiones Inmobiliarias",
        "Activo no corriente - Inversiones en empresas del grupo y asociadas a largo plazo",
        "Activo no corriente - Inversiones financieras a largo plazo",
        "Activo no corriente - Activos por impuesto diferido",
        "Activo corriente - Activos no corrientes mantenidos para la venta",
        "Activo corriente - Existencias",
        "Activo corriente - Deudores comerciales y otras cuentas a cobrar",
        "Activo corriente - Inversiones en empresas del grupo y asociadas a corto plazo",
        "Activo corriente - Inversiones financieras a corto plazo",
        "Activo corriente - Periodificaciones a corto plazo",
        "Activo corriente - Tesorería",
        "Patrimonio Neto", "Capital", "Prima de emisión",
        "Reserva Legal y estatutarias", "Otras reservas",
        "Acciones y participaciones en patrimonio propias", "Remanente",
        "Resultados de ejercicios anteriores", "Otras aportaciones de socios",
        "Resultado del ejercicio", "Dividendo a cuenta",
        "Otros instrumentos de patrimonio neto", "Ajustes por cambio de valor",
        "Subvenciones, donaciones y legados recibidos",
        "Pasivo no corriente - Provisiones a largo plazo",
        "Pasivo no corriente - Deudas a largo plazo",
        "Pasivo no corriente - Deudas con empresas del grupo y asociadas a largo plazo",
        "Pasivo no corriente - Pasivos por impuesto diferido",
        "Pasivo no corriente - Periodificaciones a largo plazo",
        "Pasivo corriente - Pasivos vinculados con activos no corrientes mantenidos para la venta",
        "Pasivo corriente - Provisiones a corto plazo",
        "Pasivo corriente - Deudas a corto plazo",
        "Pasivo corriente - Deudas con empresas del grupo y asociadas a corto plazo",
        "Pasivo corriente - Acreedores comerciales y otras cuentas a pagar",
        "Pasivo corriente - Periodificaciones a corto plazo",
        "Otros"
    ]
    
    fila = 0
    subtotales_activo = []
    subtotales_pasivo_patrimonio = []
    
    for grupo in orden_balance:
        if grupo in datos_balance:
            # Escribir encabezado del grupo
            ws_balance.write(fila, 0, grupo, header_format)
            ws_balance.write(fila, 1, "", header_format)
            ws_balance.write(fila, 2, "", header_format)
            fila += 1
            
            inicio_detalle = fila
            
            # Escribir detalles
            for cuenta, descripcion, valor in datos_balance[grupo]:
                ws_balance.write(fila, 0, cuenta)
                ws_balance.write(fila, 1, descripcion)
                ws_balance.write(fila, 2, valor, number_format)
                fila += 1
            
            # Escribir subtotal
            ws_balance.write(fila, 0, f"Subtotal {grupo}", italic_format)
            ws_balance.write_formula(fila, 2, f"=SUM(C{inicio_detalle+1}:C{fila})", bold_number_format)
            
            # Guardar referencia del subtotal
            if "Activo" in grupo:
                subtotales_activo.append(f"C{fila+1}")
            else:
                subtotales_pasivo_patrimonio.append(f"C{fila+1}")
            
            fila += 2
    
    # Totales del Balance
    ws_balance.write(fila, 0, "TOTAL ACTIVO", bold_format)
    if subtotales_activo:
        ws_balance.write_formula(fila, 2, f"={'+'.join(subtotales_activo)}", bold_number_format)
    fila_total_activo = fila
    fila += 2
    
    ws_balance.write(fila, 0, "TOTAL PATRIMONIO NETO Y PASIVO", bold_format)
    if subtotales_pasivo_patrimonio:
        ws_balance.write_formula(fila, 2, f"={'+'.join(subtotales_pasivo_patrimonio)}", bold_number_format)
    fila_total_pasivo = fila
    fila += 2
    
    ws_balance.write(fila, 0, "Diferencia", italic_format)
    ws_balance.write_formula(fila, 2, f"=C{fila_total_activo+1}-C{fila_total_pasivo+1}", number_format)
    
    # HOJA PyG
    ws_pyg = workbook.add_worksheet('PyG')
    
    # Encabezados
    ws_pyg.write(0, 0, "Cuenta", bold_format)
    ws_pyg.write(0, 1, "Descripción", bold_format)
    ws_pyg.write(0, 2, "Importe", bold_format)
    
    orden_pyg = [
        "Importe neto de la cifra de negocios",
        "Variación de existencias de productos terminados y en curso de fabricación",
        "Trabajos realizados por la empresa para su activo",
        "Aprovisionamientos",
        "Otros ingresos de explotación",
        "Gastos de personal",
        "Otros gastos de explotación",
        "Amortización del inmovilizado",
        "Excesos de provisiones",
        "Deterioro y resultado por enajenaciones del inmovilizado",
        "Ingresos financieros",
        "Gastos financieros",
        "Variación de valor razonable en instrumentos financieros",
        "Diferencias de cambio",
        "Deterioro y resultado por enajenaciones de instrumentos financieros",
        "Impuestos sobre beneficios",
        "Otros"
    ]
    
    fila = 1
    subtotales_explotacion = []
    subtotales_financiero = []
    subtotales_impuestos = []
    
    for grupo in orden_pyg:
        if grupo in datos_pyg:
            # Escribir encabezado del grupo
            ws_pyg.write(fila, 0, grupo, header_format)
            ws_pyg.write(fila, 1, "", header_format)
            ws_pyg.write(fila, 2, "", header_format)
            fila += 1
            
            inicio_detalle = fila
            
            # Escribir detalles
            for cuenta, descripcion, valor in datos_pyg[grupo]:
                ws_pyg.write(fila, 0, cuenta)
                ws_pyg.write(fila, 1, descripcion)
                ws_pyg.write(fila, 2, valor, number_format)
                fila += 1
            
            # Escribir subtotal
            ws_pyg.write(fila, 0, f"Subtotal {grupo}", italic_format)
            ws_pyg.write_formula(fila, 2, f"=SUM(C{inicio_detalle+1}:C{fila})", bold_number_format)
            
            # Clasificar subtotales
            if grupo in ["Importe neto de la cifra de negocios", "Variación de existencias de productos terminados y en curso de fabricación",
                        "Trabajos realizados por la empresa para su activo", "Otros ingresos de explotación", "Excesos de provisiones"]:
                subtotales_explotacion.append(f"C{fila+1}")
            elif grupo in ["Aprovisionamientos", "Gastos de personal", "Otros gastos de explotación", 
                          "Amortización del inmovilizado", "Deterioro y resultado por enajenaciones del inmovilizado"]:
                subtotales_explotacion.append(f"C{fila+1}")
            elif grupo in ["Ingresos financieros"]:
                subtotales_financiero.append(f"C{fila+1}")
            elif grupo in ["Gastos financieros", "Variación de valor razonable en instrumentos financieros",
                          "Diferencias de cambio", "Deterioro y resultado por enajenaciones de instrumentos financieros"]:
                subtotales_financiero.append(f"C{fila+1}")
            elif grupo in ["Impuestos sobre beneficios"]:
                subtotales_impuestos.append(f"C{fila+1}")
            
            fila += 2
    
    # Resultado de Explotación
    ws_pyg.write(fila, 0, "RESULTADO DE EXPLOTACIÓN", bold_format)
    if subtotales_explotacion:
        ws_pyg.write_formula(fila, 2, f"={'+'.join(subtotales_explotacion)}", bold_number_format)
    fila_explotacion = fila
    fila += 1
    
    # Resultado Financiero
    ws_pyg.write(fila, 0, "RESULTADO FINANCIERO", bold_format)
    if subtotales_financiero:
        ws_pyg.write_formula(fila, 2, f"={'+'.join(subtotales_financiero)}", bold_number_format)
    fila_financiero = fila
    fila += 2
    
    # Resultado antes de impuestos
    ws_pyg.write(fila, 0, "RESULTADO DEL EJERCICIO ANTES DE IMPUESTOS", bold_format)
    ws_pyg.write_formula(fila, 2, f"=C{fila_explotacion+1}+C{fila_financiero+1}", bold_number_format)
    fila_antes_impuestos = fila
    fila += 2
    
    # Resultado después de impuestos
    ws_pyg.write(fila, 0, "RESULTADO DEL EJERCICIO DESPUÉS DE IMPUESTOS", bold_format)
    if subtotales_impuestos:
        ws_pyg.write_formula(fila, 2, f"=C{fila_antes_impuestos+1}+{'+'.join(subtotales_impuestos)}", bold_number_format)
    else:
        ws_pyg.write_formula(fila, 2, f"=C{fila_antes_impuestos+1}", bold_number_format)
    
    # HOJA RESUMEN
    ws_resumen = workbook.add_worksheet('Resumen')
    
    # Título Balance
    ws_resumen.write(0, 0, "BALANCE", bold_format)
    ws_resumen.write(1, 0, "Grupo", bold_format)
    ws_resumen.write(1, 1, "Importe", bold_format)
    
    fila = 2
    total_activo = 0
    total_pasivo_patrimonio = 0
    
    # Agrupaciones para el resumen
    grupos_resumen = {
        "Activo no corriente": [
            "Activo no corriente - Inmovilizado Intangible",
            "Activo no corriente - Inmovilizado Material",
            "Activo no corriente - Inversiones Inmobiliarias",
            "Activo no corriente - Inversiones en empresas del grupo y asociadas a largo plazo",
            "Activo no corriente - Inversiones financieras a largo plazo",
            "Activo no corriente - Activos por impuesto diferido"
        ],
        "Activo corriente": [
            "Activo corriente - Existencias",
            "Activo corriente - Deudores comerciales y otras cuentas a cobrar",
            "Activo corriente - Inversiones en empresas del grupo y asociadas a corto plazo",
            "Activo corriente - Inversiones financieras a corto plazo",
            "Activo corriente - Periodificaciones a corto plazo",
            "Activo corriente - Tesorería",
            "Activo corriente - Activos no corrientes mantenidos para la venta"
        ],
        "Patrimonio Neto": [
            "Patrimonio Neto", "Capital", "Prima de emisión",
            "Reserva Legal y estatutarias", "Otras reservas",
            "Acciones y participaciones en patrimonio propias",
            "Remanente", "Resultados de ejercicios anteriores",
            "Otras aportaciones de socios", "Resultado del ejercicio",
            "Dividendo a cuenta", "Otros instrumentos de patrimonio neto",
            "Ajustes por cambio de valor", "Subvenciones, donaciones y legados recibidos"
        ],
        "Pasivo no corriente": [
            "Pasivo no corriente - Provisiones a largo plazo",
            "Pasivo no corriente - Deudas a largo plazo",
            "Pasivo no corriente - Deudas con empresas del grupo y asociadas a largo plazo",
            "Pasivo no corriente - Pasivos por impuesto diferido",
            "Pasivo no corriente - Periodificaciones a largo plazo"
        ],
        "Pasivo corriente": [
            "Pasivo corriente - Pasivos vinculados con activos no corrientes mantenidos para la venta",
            "Pasivo corriente - Provisiones a corto plazo",
            "Pasivo corriente - Deudas a corto plazo",
            "Pasivo corriente - Deudas con empresas del grupo y asociadas a corto plazo",
            "Pasivo corriente - Acreedores comerciales y otras cuentas a pagar",
            "Pasivo corriente - Periodificaciones a corto plazo"
        ]
    }
    
    for bloque, subgrupos in grupos_resumen.items():
        subtotal_bloque = 0
        
        # Calcular subtotal del bloque
        for subgrupo in subgrupos:
            if subgrupo in datos_balance:
                for cuenta, descripcion, valor in datos_balance[subgrupo]:
                    subtotal_bloque += valor
        
        if subtotal_bloque != 0:
            # Escribir línea del bloque
            ws_resumen.write(fila, 0, bloque, bold_format)
            ws_resumen.write(fila, 1, subtotal_bloque, bold_number_format)
            fila += 1
            
            # Escribir subgrupos con datos
            for subgrupo in subgrupos:
                if subgrupo in datos_balance:
                    subtotal_grupo = sum(valor for cuenta, descripcion, valor in datos_balance[subgrupo])
                    if subtotal_grupo != 0:
                        nombre_corto = subgrupo.replace(f"{bloque} - ", "")
                        ws_resumen.write(fila, 0, nombre_corto)
                        ws_resumen.write(fila, 1, subtotal_grupo, number_format)
                        fila += 1
            
            fila += 1  # Línea en blanco
            
            if "Activo" in bloque:
                total_activo += subtotal_bloque
            else:
                total_pasivo_patrimonio += subtotal_bloque
    
    # Totales del Balance en Resumen
    ws_resumen.write(fila, 0, "TOTAL ACTIVO", bold_format)
    ws_resumen.write(fila, 1, total_activo, bold_number_format)
    fila += 2
    
    ws_resumen.write(fila, 0, "TOTAL PATRIMONIO NETO Y PASIVO", bold_format)
    ws_resumen.write(fila, 1, total_pasivo_patrimonio, bold_number_format)
    fila += 2
    
    ws_resumen.write(fila, 0, "Diferencia", bold_format)
    ws_resumen.write(fila, 1, total_activo - total_pasivo_patrimonio, number_format)
    fila += 4
    
    # PyG en Resumen
    ws_resumen.write(fila, 0, "CUENTA DE PÉRDIDAS Y GANANCIAS", bold_format)
    fila += 1
    ws_resumen.write(fila, 0, "Concepto", bold_format)
    ws_resumen.write(fila, 1, "Importe", bold_format)
    fila += 1
    
    # Calcular totales PyG para resumen
    totales_pyg = {}
    for grupo, items in datos_pyg.items():
        totales_pyg[grupo] = sum(valor for cuenta, descripcion, valor in items)
    
    # Calcular resultados intermedios
    resultado_explotacion = (
        totales_pyg.get("Importe neto de la cifra de negocios", 0) +
        totales_pyg.get("Variación de existencias de productos terminados y en curso de fabricación", 0) +
        totales_pyg.get("Trabajos realizados por la empresa para su activo", 0) +
        totales_pyg.get("Otros ingresos de explotación", 0) +
        totales_pyg.get("Aprovisionamientos", 0) +
        totales_pyg.get("Gastos de personal", 0) +
        totales_pyg.get("Otros gastos de explotación", 0) +
        totales_pyg.get("Amortización del inmovilizado", 0) +
        totales_pyg.get("Excesos de provisiones", 0) +
        totales_pyg.get("Deterioro y resultado por enajenaciones del inmovilizado", 0)
    )
    
    resultado_financiero = (
        totales_pyg.get("Ingresos financieros", 0) +
        totales_pyg.get("Gastos financieros", 0) +
        totales_pyg.get("Variación de valor razonable en instrumentos financieros", 0) +
        totales_pyg.get("Diferencias de cambio", 0) +
        totales_pyg.get("Deterioro y resultado por enajenaciones de instrumentos financieros", 0)
    )
    
    resultado_antes_impuestos = resultado_explotacion + resultado_financiero
    resultado_final = resultado_antes_impuestos + totales_pyg.get("Impuestos sobre beneficios", 0)
    
    # Mostrar PyG en orden
    orden_resumen_pyg = [
        ("Importe neto de la cifra de negocios", False),
        ("Otros ingresos de explotación", False),
        ("Trabajos realizados por la empresa para su activo", False),
        ("Variación de existencias de productos terminados y en curso de fabricación", False),
        ("Aprovisionamientos", False),
        ("Gastos de personal", False),
        ("Otros gastos de explotación", False),
        ("Amortización del inmovilizado", False),
        ("Excesos de provisiones", False),
        ("Deterioro y resultado por enajenaciones del inmovilizado", False),
        ("RESULTADO DE EXPLOTACIÓN", True),
        ("Ingresos financieros", False),
        ("Gastos financieros", False),
        ("Variación de valor razonable en instrumentos financieros", False),
        ("Diferencias de cambio", False),
        ("Deterioro y resultado por enajenaciones de instrumentos financieros", False),
        ("RESULTADO FINANCIERO", True),
        ("RESULTADO ANTES DE IMPUESTOS", True),
        ("Impuestos sobre beneficios", False),
        ("RESULTADO DEL EJERCICIO", True)
    ]
    
    for concepto, es_calculo in orden_resumen_pyg:
        if es_calculo:
            # Es un resultado calculado
            if concepto == "RESULTADO DE EXPLOTACIÓN":
                valor = resultado_explotacion
            elif concepto == "RESULTADO FINANCIERO":
                valor = resultado_financiero
            elif concepto == "RESULTADO ANTES DE IMPUESTOS":
                valor = resultado_antes_impuestos
            elif concepto == "RESULTADO DEL EJERCICIO":
                valor = resultado_final
            
            ws_resumen.write(fila, 0, concepto, bold_format)
            ws_resumen.write(fila, 1, valor, bold_number_format)
            if concepto == "RESULTADO FINANCIERO":
                fila += 2
            else:
                fila += 1
        else:
            # Es un dato base
            if concepto in totales_pyg and totales_pyg[concepto] != 0:
                ws_resumen.write(fila, 0, concepto)
                ws_resumen.write(fila, 1, totales_pyg[concepto], number_format)
                fila += 1
    
    # Ajustar ancho de columnas
    ws_balance.set_column('A:A', 50)
    ws_balance.set_column('B:B', 30)
    ws_balance.set_column('C:C', 15)
    
    ws_pyg.set_column('A:A', 50)
    ws_pyg.set_column('B:B', 30)
    ws_pyg.set_column('C:C', 15)
    
    ws_resumen.set_column('A:A', 50)
    ws_resumen.set_column('B:B', 15)
    
    workbook.close()
    output.seek(0)
    
    return output

def main():
    st.set_page_config(
        page_title="Conversor de Sumas y Saldos a Balance y PyG",
        page_icon="📊",
        layout="wide"
    )
    
    st.title("📊 Conversor de Sumas y Saldos a Balance y PyG")
    st.markdown("---")
    
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.markdown("""
        ### Instrucciones:
        1. **Sube tu archivo** de Sumas y Saldos (Excel o CSV)
        2. **Verifica** que contenga una columna llamada "Saldo Definitivo"
        3. **Descarga** el archivo Excel generado con las hojas:
           - **Balance**: Balance de situación clasificado
           - **PyG**: Cuenta de Pérdidas y Ganancias
           - **Resumen**: Versión resumida de ambos estados
        """)
    
    with col2:
        st.info("""
        **Formatos soportados:**
        - Excel (.xlsx, .xls)
        - CSV (.csv)
        
        **Estructura requerida:**
        - Columna 1: Cuenta
        - Columna 2: Descripción
        - Columna con "Saldo Definitivo"
        """)
    
    # Upload file
    uploaded_file = st.file_uploader(
        "Selecciona tu archivo de Sumas y Saldos",
        type=['xlsx', 'xls', 'csv'],
        help="Asegúrate de que el archivo tenga la estructura correcta con las columnas de cuenta, descripción y saldo definitivo"
    )
    
    if uploaded_file is not None:
        try:
            # Leer archivo
            if uploaded_file.name.endswith('.csv'):
                df = pd.read_csv(uploaded_file, encoding='utf-8')
            else:
                df = pd.read_excel(uploaded_file)
            
            st.success(f"✅ Archivo cargado correctamente: {uploaded_file.name}")
            
            # Mostrar información del archivo
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Filas", len(df))
            with col2:
                st.metric("Columnas", len(df.columns))
            with col3:
                saldo_cols = [col for col in df.columns if 'saldo' in str(col).lower()]
                st.metric("Columnas de Saldo", len(saldo_cols))
            
            # Mostrar preview
            st.markdown("### Vista previa del archivo:")
            st.dataframe(df.head(10), use_container_width=True)
            
            # Mostrar columnas disponibles
            st.markdown("### Columnas detectadas:")
            cols_info = []
            for i, col in enumerate(df.columns):
                cols_info.append(f"**Columna {i+1}:** {col}")
            st.markdown(" | ".join(cols_info))
            
            # Botón para procesar
            if st.button("🔄 Procesar y Generar Balance y PyG", type="primary", use_container_width=True):
                with st.spinner("Procesando archivo..."):
                    # Procesar datos
                    datos_balance, datos_pyg = procesar_sys(df)
                    
                    if datos_balance is not None and datos_pyg is not None:
                        # Mostrar estadísticas
                        col1, col2 = st.columns(2)
                        with col1:
                            st.success(f"✅ Balance: {len(datos_balance)} grupos procesados")
                            balance_cuentas = sum(len(items) for items in datos_balance.values())
                            st.info(f"📊 Total cuentas en Balance: {balance_cuentas}")
                        
                        with col2:
                            st.success(f"✅ PyG: {len(datos_pyg)} grupos procesados")
                            pyg_cuentas = sum(len(items) for items in datos_pyg.values())
                            st.info(f"📊 Total cuentas en PyG: {pyg_cuentas}")
                        
                        # Generar Excel
                        excel_file = crear_excel_balance_pyg(datos_balance, datos_pyg)
                        
                        # Botón de descarga
                        st.download_button(
                            label="📥 Descargar Balance y PyG (Excel)",
                            data=excel_file,
                            file_name=f"Balance_PyG_{uploaded_file.name.split('.')[0]}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
                        
                        st.success("🎉 ¡Archivo procesado correctamente! Puedes descargar el resultado.")
                        
                        # Mostrar algunos grupos detectados
                        if datos_balance:
                            st.markdown("### Grupos detectados en Balance:")
                            balance_grupos = list(datos_balance.keys())[:10]  # Mostrar solo los primeros 10
                            for grupo in balance_grupos:
                                st.write(f"- {grupo} ({len(datos_balance[grupo])} cuentas)")
                            if len(datos_balance) > 10:
                                st.write(f"... y {len(datos_balance) - 10} grupos más")
                        
                        if datos_pyg:
                            st.markdown("### Grupos detectados en PyG:")
                            pyg_grupos = list(datos_pyg.keys())[:10]  # Mostrar solo los primeros 10
                            for grupo in pyg_grupos:
                                st.write(f"- {grupo} ({len(datos_pyg[grupo])} cuentas)")
                            if len(datos_pyg) > 10:
                                st.write(f"... y {len(datos_pyg) - 10} grupos más")
                    
        except Exception as e:
            st.error(f"❌ Error al procesar el archivo: {str(e)}")
            st.markdown("""
            **Posibles soluciones:**
            - Verifica que el archivo tenga la estructura correcta
            - Asegúrate de que existe una columna con 'Saldo Definitivo' o similar
            - Revisa que las cuentas estén en la primera columna
            - Intenta con un archivo de ejemplo más pequeño
            """)
    
    # Información adicional
    st.markdown("---")
    st.markdown("""
    ### ℹ️ Información adicional
    
    Esta aplicación replica la funcionalidad del código VBA original para convertir un archivo de Sumas y Saldos 
    en Balance de Situación y Cuenta de Pérdidas y Ganancias, siguiendo el Plan General Contable Español.
    
    **Características:**
    - ✅ Clasificación automática según Plan General Contable
    - ✅ Generación de Balance de Situación estructurado
    - ✅ Generación de Cuenta de Pérdidas y Ganancias
    - ✅ Hoja resumen con totales y subtotales
    - ✅ Ajuste automático de signos contables
    - ✅ Formato Excel profesional
    
    **Desarrollado con:** Python + Streamlit + Pandas + XlsxWriter
    """)

if __name__ == "__main__":
    main()