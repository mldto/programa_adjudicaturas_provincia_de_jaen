import customtkinter as ctk
from tkinter import filedialog, messagebox
import pdfplumber
import re
import os
import glob
import pandas as pd
import threading

# -------------------------------------------------------------------------
# INYECCIÓN DE ALTA RESOLUCIÓN NATIVA (Antes de cargar la interfaz)
# -------------------------------------------------------------------------
try:
    import ctypes
    # Decimos a Windows que esta app maneja su propia resolución por monitor (HD)
    ctypes.windll.shcore.SetProcessDpiAwareness(2)
except Exception:
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass

# Configuración estética global de CustomTkinter
ctk.set_appearance_mode("Light")  
ctk.set_default_color_theme("green")  

class AnalizadorJaenApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Configuración de la Ventana
        self.title("Auditoría de Adjudicaciones - Provincia de Jaén")
        self.geometry("680x480")
        self.resizable(False, False)

        # Forzar el escalado de CustomTkinter al 100% de nitidez según el sistema
        try:
            ctk.deactivate_automatic_dpi_awareness()
        except Exception:
            pass

        # Cargar icono si existe
        ruta_icono = os.path.join(os.path.dirname(__file__), "logo.ico")
        if os.path.exists(ruta_icono):
            try:
                self.iconbitmap(ruta_icono)
            except Exception:
                pass

        # Variables de control
        self.carpeta_seleccionada = ctk.StringVar(value="No se ha seleccionado ninguna carpeta")

        # --- DISEÑO DE LA INTERFAZ ---
        # Banner Superior - Verde Corporativo
        self.banner = ctk.CTkFrame(self, height=80, fg_color="#006241", corner_radius=0)
        self.banner.pack(fill="x", side="top")
        
        self.lbl_titulo = ctk.CTkLabel(
            self.banner, 
            text="CONTROL ESTRATÉGICO DE PLAZAS - JAÉN", 
            font=ctk.CTkFont(family="Segoe UI", size=18, weight="bold"),
            text_color="white"
        )
        self.lbl_titulo.place(relx=0.5, rely=0.5, anchor="center")

        # Contenedor Principal (Cuerpo)
        self.cuerpo = ctk.CTkFrame(self, fg_color="white", corner_radius=0)
        self.cuerpo.pack(fill="both", expand=True)

        # Instrucciones de uso
        instrucciones = (
            "Este software analiza los anexos de adjudicaciones de todas las Consejerías.\n"
            "Detectará automáticamente a las personas con plaza asignada en Jaén y "
            "consolidará en una única fila si aparecen adjudicadas en otras provincias (Riesgo de Fuga)."
        )
        self.lbl_instrucciones = ctk.CTkLabel(
            self.cuerpo, text=instrucciones, text_color="#333333",
            font=ctk.CTkFont(family="Segoe UI", size=12), justify="center"
        )
        self.lbl_instrucciones.pack(pady=20)

        # Sección: Selección de Carpeta
        self.frame_carpeta = ctk.CTkFrame(self.cuerpo, fg_color="#F4F4F4", height=70, corner_radius=8)
        self.frame_carpeta.pack(fill="x", padx=30, pady=10)
        self.frame_carpeta.pack_propagate(False)

        self.btn_carpeta = ctk.CTkButton(
            self.frame_carpeta, text="Seleccionar Carpeta", 
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            fg_color="#006241", hover_color="#004d33",
            command=self.buscar_carpeta
        )
        self.btn_carpeta.pack(side="left", padx=15, pady=20)

        self.lbl_ruta = ctk.CTkLabel(
            self.frame_carpeta, textvariable=self.carpeta_seleccionada, 
            text_color="#666666", font=ctk.CTkFont(family="Segoe UI", size=11), anchor="w"
        )
        self.lbl_ruta.pack(side="left", fill="x", expand=True, padx=10)

        # Barra de progreso y Estado
        self.lbl_estado = ctk.CTkLabel(
            self.cuerpo, text="Estado: Esperando selección de carpeta...", 
            text_color="#333333", font=ctk.CTkFont(family="Segoe UI", weight="bold")
        )
        self.lbl_estado.pack(pady=(20, 5))

        self.progreso = ctk.CTkProgressBar(self.cuerpo, width=400, progress_color="#006241")
        self.progreso.pack(pady=5)
        self.progreso.set(0)

        # Botón de Acción Principal (Procesar)
        self.btn_procesar = ctk.CTkButton(
            self.cuerpo, text="🎯 GENERAR INFORME DE CONTROL", 
            font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"),
            fg_color="#28A745", hover_color="#218838", height=45,
            state="disabled", command=self.iniciar_procesamiento_hilo
        )
        self.btn_procesar.pack(pady=30)

    # --- LÓGICA INTERNA DE DATOS ---
    def buscar_carpeta(self):
        ruta = filedialog.askdirectory()
        if ruta:
            self.carpeta_seleccionada.set(ruta)
            pdfs = glob.glob(os.path.join(ruta, "*.pdf"))
            if pdfs:
                self.lbl_estado.configure(text=f"✅ Detectados {len(pdfs)} archivos PDF listos para analizar.")
                self.btn_procesar.configure(state="normal")
            else:
                self.lbl_estado.configure(text="❌ Error: No se encontraron archivos PDF en la carpeta.")
                self.btn_procesar.configure(state="disabled")

    def iniciar_procesamiento_hilo(self):
        self.btn_carpeta.configure(state="disabled")
        self.btn_procesar.configure(state="disabled")
        hilo = threading.Thread(target=self.ejecutar_analisis_motor)
        hilo.start()

    def limpiar_espacios(self, texto):
        if not texto: return ""
        return re.sub(r'\s+', ' ', texto).strip()

    def aislar_nombre_real(self, texto_sucio):
        patrones_corte = [
            r"\bNG\b", r"\bAGENTE\b", r"\bAUXILIAR\b", r"\bEXPERIENCIA\b", 
            r"\bA\.T\.-", r"\bA\d\.\d{4}", r"\bC\d\.\d{4}", r"\bSC\b", r"\bDP\b"
        ]
        nombre_limpio = texto_sucio
        for patron in patrones_corte:
            nombre_limpio = re.split(patron, nombre_limpio, flags=re.IGNORECASE)[0]
        return self.limpiar_espacios(re.sub(r'\d+', '', nombre_limpio))

    def extraer_puntuacion_exacta(self, linea_actual, lineas_vecinas):
        todo_el_bloque = " ".join(lineas_vecinas)
        total_match = re.search(r"Total:\s*([\d,.]+)", todo_el_bloque, re.IGNORECASE)
        if total_match:
            return float(total_match.group(1).replace(',', '.'))
        
        decimales = re.findall(r"[\d,]+", linea_actual)
        for num in reversed(decimales):
            if ',' in num and len(num.split(',')[1]) >= 2:
                try:
                    return float(num.replace(',', '.'))
                except ValueError:
                    continue
        return 0.0

    def ejecutar_analisis_motor(self):
        carpeta = self.carpeta_seleccionada.get()
        archivos_pdf = glob.glob(os.path.join(carpeta, "*.pdf"))
        todos_los_concursantes = []
        total_archivos = len(archivos_pdf)
        
        for idx, ruta in enumerate(archivos_pdf, 1):
            nombre_archivo = os.path.basename(ruta)
            self.lbl_estado.configure(text=f"Analizando ({idx}/{total_archivos}): {nombre_archivo[:30]}...")
            self.progreso.set(idx / total_archivos)
            
            try:
                with pdfplumber.open(ruta) as pdf:
                    for num_pagina, pagina in enumerate(pdf.pages, 1):
                        texto_pagina = pagina.extract_text()
                        if not texto_pagina: continue
                        lineas = texto_pagina.split("\n")
                        
                        for i, linea in enumerate(lineas):
                            linea_actual = self.limpiar_espacios(linea)
                            if " / " in linea_actual or "PROVINCIAL" in linea_actual.upper() or "D.T." in linea_actual.upper():
                                linea_anterior = self.limpiar_espacios(lineas[i-1]) if i > 0 else ""
                                dni_match = re.search(r"(\*\*\*\d{4}\*\*|\*\*\*\d{4}\*)", linea_anterior)
                                
                                if dni_match:
                                    dni = dni_match.group(1)
                                    es_jaen = any(prov in linea_actual.upper() for prov in ["JAÉN", "JAEN"])
                                    texto_nombre_sucio = linea_anterior.replace(dni, "")
                                    nombre_completo = self.aislar_nombre_real(texto_nombre_sucio)
                                    
                                    destino_match = re.search(r"(.*?)(?:[A-ZÁÉÍÓÚÑ]+\s*/\s*[A-ZÁÉÍÓÚÑ]+)", linea_actual, re.IGNORECASE)
                                    centro_destino = self.limpiar_espacios(destino_match.group(1)) if destino_match else "Ver Anexo"
                                    if len(centro_destino) < 5: centro_destino = linea_actual[:35]
                                    
                                    vecinas = lineas[max(0, i-2):min(len(lineas), i+3)]
                                    puntuacion = self.extraer_puntuacion_exacta(linea_actual, vecinas)

                                    todos_los_concursantes.append({
                                        "DNI": dni,
                                        "Nombre Completo": nombre_completo,
                                        "EsJaen": es_jaen,
                                        "Centro Jaén": centro_destino if es_jaen else "",
                                        "Puntuación Jaén": puntuacion if es_jaen else 0.0,
                                        "PDF Origen Jaén": nombre_archivo if es_jaen else "",
                                        "PDF Completo": nombre_archivo
                                    })
            except Exception as e:
                print(f"Error: {e}")

        if not todos_los_concursantes:
            self.finalizar_interfaz("❌ Error", "No se extrajeron datos.")
            return

        df_bruto = pd.DataFrame(todos_los_concursantes)
        dnis_jaen = df_bruto[df_bruto["EsJaen"] == True]["DNI"].unique()

        if len(dnis_jaen) == 0:
            self.finalizar_interfaz("⚠️ Aviso", "No hay adjudicatarios en Jaén.")
            return

        registros_consolidados = []
        for dni in dnis_jaen:
            movimientos_persona = df_bruto[df_bruto["DNI"] == dni]
            fila_jaen = movimientos_persona[movimientos_persona["EsJaen"] == True].iloc[0]
            
            otros_pdfs = movimientos_persona[movimientos_persona["PDF Completo"] != fila_jaen["PDF Origen Jaén"]]["PDF Completo"].unique()
            alerta_duplicados = ", ".join(otros_pdfs) if len(otros_pdfs) > 0 else "Ninguno (Plaza Segura)"
            
            registros_consolidados.append({
                "DNI": dni,
                "Nombre Completo": fila_jaen["Nombre Completo"],
                "Centro Adjudicado (Jaén)": fila_jaen["Centro Jaén"],
                "Puntuación": fila_jaen["Puntuación Jaén"],
                "PDF Convocatoria Jaén": fila_jaen["PDF Origen Jaén"],
                "Riesgo (Otras Adjudicaciones)": alerta_duplicados
            })

        df_final = pd.DataFrame(registros_consolidados).sort_values(by="Puntuación", ascending=False)
        archivo_excel = os.path.join(carpeta, "Informe_Final_Jaen_Limpio.xlsx")
        
        with pd.ExcelWriter(archivo_excel, engine='xlsxwriter') as writer:
            df_final.to_excel(writer, sheet_name='Control de Plazas', index=False)
            workbook  = writer.book
            worksheet = writer.sheets['Control de Plazas']
            
            formato_cabecera = workbook.add_format({
                'bold': True, 'text_wrap': True, 'valign': 'vcenter', 'align': 'center',
                'fg_color': '#006241', 'font_color': 'white', 'border': 1
            })
            formato_puntos = workbook.add_format({'num_format': '0.0000', 'align': 'right'})
            formato_alerta = workbook.add_format({'fg_color': '#FCE4D6', 'font_color': '#C00000'})
            formato_seguro = workbook.add_format({'fg_color': '#E2EFDA', 'font_color': '#375623'})
            
            worksheet.set_row(0, 28, formato_cabecera)
            
            for col_num, col_nombre in enumerate(df_final.columns):
                max_len = max(df_final[col_nombre].astype(str).map(len).max(), len(col_nombre)) + 4
                if col_nombre == "Puntuación":
                    worksheet.set_column(col_num, col_num, max_len, formato_puntos)
                else:
                    worksheet.set_column(col_num, col_num, min(max_len, 45))
                    
            for fila_idx in range(1, len(df_final) + 1):
                valor_riesgo = df_final.iloc[fila_idx-1]["Riesgo (Otras Adjudicaciones)"]
                if "Ninguno" in valor_riesgo:
                    worksheet.write(fila_idx, 5, valor_riesgo, formato_seguro)
                else:
                    worksheet.write(fila_idx, 5, valor_riesgo, formato_alerta)

        self.finalizar_interfaz("🎉 Éxito", f"Informe generado.\nGuardado como:\n'Informe_Final_Jaen_Limpio.xlsx'")

    def finalizar_interfaz(self, titulo, mensaje):
        self.btn_carpeta.configure(state="normal")
        self.btn_procesar.configure(state="normal")
        self.lbl_estado.configure(text="Estado: Proceso finalizado.")
        messagebox.showinfo(titulo, mensaje)

if __name__ == "__main__":
    app = AnalizadorJaenApp()
    app.mainloop()