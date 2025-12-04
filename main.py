import customtkinter as ctk
import pandas as pd
from tkinter import messagebox
import os
from datetime import datetime

# --- BIBLIOTECAS ---
from reportlab.lib.pagesizes import A4, landscape
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.utils import simpleSplit 

ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

class SistemaEtiquetas(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Sistema NutriBem + (V10.0 - Final)")
        self.geometry("1150x800") 
        self.resizable(False, False)

        # --- ÍCONE DO SISTEMA ---
        # Se existir o favicon.ico na pasta, ele usa.
        if os.path.exists("favicon.ico"):
            try:
                self.iconbitmap("favicon.ico")
            except:
                pass # Se der erro no ícone, o programa abre igual

        self.df_completo = None  
        self.df_pacientes = None 
        self.paciente_selecionado = None
        self.fila_impressao = [] 

        self.grid_columnconfigure(0, weight=1) 
        self.grid_columnconfigure(1, weight=1) 

        self.criar_interface_esquerda()
        self.criar_interface_direita()
        
        self.carregar_dados(feedback=False)

    def limpar_leito(self, valor):
        if pd.isna(valor) or valor == "":
            return ""
        try:
            return str(int(float(valor)))
        except:
            return str(valor)

    def criar_interface_esquerda(self):
        self.frame_esq = ctk.CTkFrame(self, width=400, corner_radius=10)
        self.frame_esq.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")

        ctk.CTkLabel(self.frame_esq, text="📋 Pacientes Internados", font=("Arial", 18, "bold")).pack(pady=(15, 5))

        self.btn_reload = ctk.CTkButton(self.frame_esq, text="🔄 Atualizar Lista (Ler Excel)", 
                                        fg_color="#555555", hover_color="#333333", 
                                        height=30, command=lambda: self.carregar_dados(feedback=True))
        self.btn_reload.pack(pady=5)

        self.entrada_busca = ctk.CTkEntry(self.frame_esq, placeholder_text="Buscar...", width=300)
        self.entrada_busca.pack(pady=10)
        self.entrada_busca.bind("<KeyRelease>", self.filtrar_lista)

        self.scroll_pacientes = ctk.CTkScrollableFrame(self.frame_esq, width=350, height=500)
        self.scroll_pacientes.pack(pady=5, padx=10, fill="both", expand=True)

        self.lbl_status_lista = ctk.CTkLabel(self.frame_esq, text="Carregando...", text_color="gray")
        self.lbl_status_lista.pack(pady=5)

    def criar_interface_direita(self):
        self.frame_dir = ctk.CTkFrame(self, width=550, corner_radius=10)
        self.frame_dir.grid(row=0, column=1, padx=20, pady=20, sticky="nsew")

        ctk.CTkLabel(self.frame_dir, text="⚙️ Central de Controle", font=("Arial", 18, "bold")).pack(pady=10)

        # --- ETIQUETAS ---
        self.frame_etiquetas = ctk.CTkFrame(self.frame_dir, fg_color="#2b2b2b")
        self.frame_etiquetas.pack(pady=5, padx=10, fill="x")
        
        ctk.CTkLabel(self.frame_etiquetas, text="ÁREA DE ETIQUETAS", font=("Arial", 12, "bold"), text_color="silver").pack(pady=5)
        self.lbl_detalhe_info = ctk.CTkLabel(self.frame_etiquetas, text="Selecione um paciente.", justify="center")
        self.lbl_detalhe_info.pack(pady=2)

        self.btn_add_fila = ctk.CTkButton(self.frame_etiquetas, text="⬇️ Adicionar à Fila", command=self.adicionar_selecionado_fila, state="disabled")
        self.btn_add_fila.pack(pady=5)
        self.btn_add_todos = ctk.CTkButton(self.frame_etiquetas, text="⬇️⬇️ ADICIONAR TODOS ⬇️⬇️", fg_color="#1f6aa5", command=self.adicionar_todos_fila)
        self.btn_add_todos.pack(pady=5)

        self.scroll_fila = ctk.CTkScrollableFrame(self.frame_etiquetas, height=120, fg_color="#3a3a3a")
        self.scroll_fila.pack(pady=5, padx=10, fill="x")
        self.lbl_contador_fila = ctk.CTkLabel(self.frame_etiquetas, text="0 etiquetas na fila", text_color="yellow")
        self.lbl_contador_fila.pack(pady=0)

        self.frame_botoes_print = ctk.CTkFrame(self.frame_etiquetas, fg_color="transparent")
        self.frame_botoes_print.pack(pady=10)
        ctk.CTkButton(self.frame_botoes_print, text="Limpar", fg_color="red", width=80, command=self.limpar_fila).pack(side="left", padx=5)
        ctk.CTkButton(self.frame_botoes_print, text="🖨️ IMPRIMIR ETIQUETAS", fg_color="green", width=200, height=40, command=self.gerar_pdf_etiquetas).pack(side="left", padx=5)

        # --- RELATÓRIOS ---
        ctk.CTkFrame(self.frame_dir, height=2, fg_color="gray").pack(fill="x", pady=15, padx=20) 
        ctk.CTkLabel(self.frame_dir, text="📑 Relatórios Gerenciais", font=("Arial", 16, "bold")).pack(pady=5)
        
        self.btn_relatorio_ativos = ctk.CTkButton(self.frame_dir, text="📄 RELATÓRIO SIMPLES (Só Ocupados)", 
                                                 fg_color="#D35400", hover_color="#A04000", height=40, width=400,
                                                 command=self.gerar_relatorio_ativos)
        self.btn_relatorio_ativos.pack(pady=5)

        self.btn_relatorio_full = ctk.CTkButton(self.frame_dir, text="📊 MAPA GERAL (Mesclado)", 
                                                 fg_color="#5B2C6F", hover_color="#4A235A", height=50, width=400,
                                                 command=self.gerar_relatorio_completo_com_vazios)
        self.btn_relatorio_full.pack(pady=10)

    def carregar_dados(self, feedback=False):
        try:
            self.limpar_fila()
            df_raw = pd.read_excel("pacientes.xlsx")
            
            df_raw['ENFERMARIA'] = df_raw['ENFERMARIA'].ffill()
            df_raw['LEITO'] = df_raw['LEITO'].apply(self.limpar_leito)
            
            self.df_completo = df_raw.copy()
            self.df_pacientes = df_raw.dropna(subset=['NOME DO PACIENTE']).copy()
            self.df_pacientes['NOME DO PACIENTE'] = self.df_pacientes['NOME DO PACIENTE'].str.strip()
            
            self.povoar_lista_pacientes(self.df_pacientes)
            
            if feedback:
                qtd = len(self.df_pacientes)
                messagebox.showinfo("Atualizado", f"Dados recarregados!\n{qtd} pacientes encontrados.")
            
        except PermissionError:
            messagebox.showerror("Erro", "O Excel está aberto! Feche e tente novamente.")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro: {e}")

    def povoar_lista_pacientes(self, df):
        for widget in self.scroll_pacientes.winfo_children(): widget.destroy()
        if df.empty:
            self.lbl_status_lista.configure(text="Nenhum paciente.")
            return
        for index, row in df.iterrows():
            btn = ctk.CTkButton(self.scroll_pacientes, text=f"{row['LEITO']} - {row['NOME DO PACIENTE']}", 
                                fg_color="transparent", border_width=1, border_color="gray",
                                text_color=("black", "white"), anchor="w",
                                command=lambda r=row: self.selecionar_paciente(r))
            btn.pack(pady=2, fill="x")
        self.lbl_status_lista.configure(text=f"{len(df)} pacientes ativos.")

    def filtrar_lista(self, event=None):
        if self.df_pacientes is None: return
        termo = self.entrada_busca.get().lower()
        mask = (self.df_pacientes['NOME DO PACIENTE'].str.lower().str.contains(termo, na=False)) | \
               (self.df_pacientes['LEITO'].str.contains(termo, na=False))
        self.povoar_lista_pacientes(self.df_pacientes[mask])

    def selecionar_paciente(self, row):
        self.paciente_selecionado = row
        dieta = row['DIETA'] if pd.notna(row['DIETA']) else "---"
        self.lbl_detalhe_info.configure(text=f"{row['NOME DO PACIENTE']}\nLeito: {row['LEITO']} | Dieta: {dieta}")
        self.btn_add_fila.configure(state="normal")

    def adicionar_selecionado_fila(self):
        if self.paciente_selecionado is None: return
        self.fila_impressao.append(self.paciente_selecionado)
        self.atualizar_fila_visual()

    def adicionar_todos_fila(self):
        if self.df_pacientes is None or self.df_pacientes.empty: return
        if messagebox.askyesno("Confirmar", f"Adicionar {len(self.df_pacientes)} etiquetas?"):
            for index, row in self.df_pacientes.iterrows():
                self.fila_impressao.append(row)
            self.atualizar_fila_visual()

    def atualizar_fila_visual(self):
        for widget in self.scroll_fila.winfo_children(): widget.destroy()
        for p in self.fila_impressao:
            ctk.CTkLabel(self.scroll_fila, text=f"✅ {p['LEITO']} - {p['NOME DO PACIENTE']}", anchor="w").pack(fill="x")
        self.lbl_contador_fila.configure(text=f"{len(self.fila_impressao)} etiquetas na fila")

    def limpar_fila(self):
        self.fila_impressao = []
        self.atualizar_fila_visual()

    # --- GERADOR ETIQUETAS ---
    def gerar_pdf_etiquetas(self):
        if not self.fila_impressao:
            messagebox.showwarning("Vazio", "Fila vazia!")
            return
        arquivo_pdf = "etiquetas_imprimir.pdf"
        c = canvas.Canvas(arquivo_pdf, pagesize=A4)
        
        largura_etiqueta, altura_etiqueta = 95*mm, 52*mm
        gap_vertical = 3*mm
        colunas, linhas_por_pag = 2, 5
        
        for i, p in enumerate(self.fila_impressao):
            if i > 0 and i % (colunas * linhas_por_pag) == 0: c.showPage()
            pos_pag = i % (colunas * linhas_por_pag)
            x = 10*mm + ((pos_pag % colunas) * (largura_etiqueta + 5*mm))
            y = A4[1] - 10*mm - (((pos_pag // colunas) + 1) * (altura_etiqueta + gap_vertical))
            self.desenhar_etiqueta_individual(c, x, y, largura_etiqueta, altura_etiqueta, p)
        c.save()
        try: os.startfile(arquivo_pdf)
        except: messagebox.showinfo("Sucesso", "Gerado!")

    def desenhar_etiqueta_individual(self, c, x, y, w, h, p):
        c.setStrokeColorRGB(0, 0, 0); c.rect(x, y, w, h)
        c.setFont("Helvetica-Bold", 9); c.drawCentredString(x + w/2, y + h - 8*mm, "SILVA E TEIXEIRA")
        c.setFont("Helvetica", 7); c.drawCentredString(x + w/2, y + h - 12*mm, "IDENTIFICAÇÃO DE DIETAS")
        
        obs = str(p['OBSERVAÇÕES']) if pd.notna(p['OBSERVAÇÕES']) else ''
        dieta = str(p['DIETA']) if pd.notna(p['DIETA']) else ''
        nome = p['NOME DO PACIENTE']
        
        def desenhar_campo_quebrado(canvas_obj, texto_label, texto_valor, pos_x, pos_y, max_width):
            canvas_obj.setFont("Helvetica-Bold", 8)
            texto_completo = f"{texto_label} {texto_valor}"
            linhas = simpleSplit(texto_completo, "Helvetica-Bold", 8, max_width)
            for linha in linhas:
                canvas_obj.drawString(pos_x, pos_y, linha)
                pos_y -= 4 * mm 
            return pos_y - 2*mm 

        margem_esq = x + 3*mm
        cursor_y = y + h - 20*mm
        largura_texto = w - 6*mm 

        cursor_y = desenhar_campo_quebrado(c, "PACIENTE:", nome, margem_esq, cursor_y, largura_texto)
        cursor_y = desenhar_campo_quebrado(c, "ENF:", f"{p['ENFERMARIA']} - LEITO: {p['LEITO']}", margem_esq, cursor_y, largura_texto)
        cursor_y = desenhar_campo_quebrado(c, "DIETA:", dieta, margem_esq, cursor_y, largura_texto)
        cursor_y = desenhar_campo_quebrado(c, "OBS:", obs, margem_esq, cursor_y, largura_texto)
        c.drawString(margem_esq, y + 2*mm, f"DATA: {datetime.now().strftime('%d/%m/%Y')}")

    # --- RELATÓRIOS ---
    def gerar_relatorio_ativos(self):
        self.gerar_tabela_pdf(self.df_pacientes, "relatorio_ativos.pdf", "PACIENTES OCUPADOS", mesclar=False)

    def gerar_relatorio_completo_com_vazios(self):
        self.gerar_tabela_pdf(self.df_completo, "MAPA_GERAL.pdf", "MAPA GERAL (AUDITORIA)", mesclar=True)

    def gerar_tabela_pdf(self, df_alvo, nome_arquivo, subtitulo, mesclar=False):
        if df_alvo is None or df_alvo.empty:
            messagebox.showwarning("Erro", "Sem dados.")
            return

        doc = SimpleDocTemplate(nome_arquivo, pagesize=landscape(A4), rightMargin=20, leftMargin=20, topMargin=20, bottomMargin=20)
        elements = []
        styles = getSampleStyleSheet()
        
        if os.path.exists("logo.png"):
            logo = Image("logo.png", width=35*mm, height=35*mm)
            logo.hAlign = 'CENTER' 
            elements.append(logo)
            elements.append(Spacer(1, 10))

        estilo_sub = ParagraphStyle('SubTitle', parent=styles['Normal'], alignment=1, fontSize=10)
        elements.append(Paragraph(f"<b>{subtitulo}</b> - Emitido em: {datetime.now().strftime('%d/%m/%Y %H:%M')}", estilo_sub))
        elements.append(Spacer(1, 15))

        estilo_celula = ParagraphStyle('CellStyle', parent=styles['Normal'], fontSize=9, leading=11)
        
        # --- ORDEM DAS COLUNAS ALTERADA AQUI ---
        # 1. Enfermaria, 2. Leito, 3. Nome, 4. Dieta, 5. Obs
        data = [['ENFERMARIA', 'LEITO', 'NOME DO PACIENTE', 'DIETA', 'OBSERVAÇÕES']]
        
        for index, row in df_alvo.iterrows():
            nome = str(row['NOME DO PACIENTE']) if pd.notna(row['NOME DO PACIENTE']) else ""
            enf = str(row['ENFERMARIA']) if pd.notna(row['ENFERMARIA']) else ""
            leito = str(row['LEITO']) 
            dieta = str(row['DIETA']) if pd.notna(row['DIETA']) else ""
            obs = str(row['OBSERVAÇÕES']) if pd.notna(row['OBSERVAÇÕES']) else ""

            data.append([
                Paragraph(enf, estilo_celula),   # Agora é o primeiro
                leito,                           # Agora é o segundo
                Paragraph(nome, estilo_celula),
                Paragraph(dieta, estilo_celula),
                Paragraph(obs, estilo_celula)
            ])

        # --- AJUSTE DE LARGURA PARA A NOVA ORDEM ---
        # Enf(110), Leito(50), Nome(250), Dieta(160), Obs(200)
        col_widths = [110, 50, 250, 160, 200]
        t = Table(data, colWidths=col_widths, repeatRows=1)

        comandos_estilo = [
            ('BACKGROUND', (0, 0), (-1, 0), colors.Color(0.2, 0.6, 0.3)), 
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
            ('TOPPADDING', (0, 0), (-1, -1), 6),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.whitesmoke, colors.white]) 
        ]

        if mesclar:
            grupo_anterior = None
            inicio_grupo = 1 
            df_reset = df_alvo.reset_index(drop=True)
            for i in range(len(df_reset)):
                linha_atual_tabela = i + 1 
                enf_atual = df_reset.iloc[i]['ENFERMARIA']
                if enf_atual != grupo_anterior:
                    if grupo_anterior is not None:
                        fim_grupo = linha_atual_tabela - 1
                        # ATENÇÃO: Índice da coluna mudou para 0 (Enfermaria é a primeira)
                        comandos_estilo.append(('SPAN', (0, inicio_grupo), (0, fim_grupo)))
                        comandos_estilo.append(('VALIGN', (0, inicio_grupo), (0, fim_grupo), 'MIDDLE'))
                    grupo_anterior = enf_atual
                    inicio_grupo = linha_atual_tabela
            
            # Fecha o último grupo
            comandos_estilo.append(('SPAN', (0, inicio_grupo), (0, len(df_reset))))
            comandos_estilo.append(('VALIGN', (0, inicio_grupo), (0, len(df_reset)), 'MIDDLE'))

        t.setStyle(TableStyle(comandos_estilo))
        elements.append(t)
        
        elements.append(Spacer(1, 40))
        estilo_assinatura = ParagraphStyle('Assinatura', parent=styles['Normal'], alignment=TA_CENTER)
        elements.append(Paragraph("_"*60, estilo_assinatura))
        elements.append(Paragraph("<b>NUTRICIONISTA RESPONSÁVEL</b>", estilo_assinatura))

        try:
            doc.build(elements)
            os.startfile(nome_arquivo)
        except Exception as e:
            messagebox.showerror("Erro PDF", f"Erro: {e}")

if __name__ == "__main__":
    app = SistemaEtiquetas()
    app.mainloop()