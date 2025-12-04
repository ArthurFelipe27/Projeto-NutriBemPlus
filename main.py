import customtkinter as ctk
import pandas as pd
from tkinter import messagebox
import os
from datetime import datetime

# --- BIBLIOTECAS DE PDF E TABELAS ---
from reportlab.lib.pagesizes import A4, landscape
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.utils import simpleSplit # Importante para quebrar texto na etiqueta

# Configuração visual
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

class SistemaEtiquetas(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Sistema Hospitalar - Nutrição (V5.0 - Refinado)")
        self.geometry("1150x800") 
        self.resizable(False, False)

        # Dados
        self.df_completo = None  
        self.df_pacientes = None 
        self.paciente_selecionado = None
        self.fila_impressao = [] 

        # Layout (2 Colunas)
        self.grid_columnconfigure(0, weight=1) 
        self.grid_columnconfigure(1, weight=1) 

        self.criar_interface_esquerda()
        self.criar_interface_direita()
        
        self.carregar_dados()

    def criar_interface_esquerda(self):
        self.frame_esq = ctk.CTkFrame(self, width=400, corner_radius=10)
        self.frame_esq.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")

        ctk.CTkLabel(self.frame_esq, text="📋 Pacientes Ativos (Etiquetas)", font=("Arial", 18, "bold")).pack(pady=10)

        self.entrada_busca = ctk.CTkEntry(self.frame_esq, placeholder_text="Filtrar nome, leito ou dieta...", width=300)
        self.entrada_busca.pack(pady=5)
        self.entrada_busca.bind("<KeyRelease>", self.filtrar_lista)

        self.scroll_pacientes = ctk.CTkScrollableFrame(self.frame_esq, width=350, height=550)
        self.scroll_pacientes.pack(pady=10, padx=10, fill="both", expand=True)

        self.lbl_status_lista = ctk.CTkLabel(self.frame_esq, text="Carregando...", text_color="gray")
        self.lbl_status_lista.pack(pady=5)

    def criar_interface_direita(self):
        self.frame_dir = ctk.CTkFrame(self, width=550, corner_radius=10)
        self.frame_dir.grid(row=0, column=1, padx=20, pady=20, sticky="nsew")

        ctk.CTkLabel(self.frame_dir, text="⚙️ Painel de Controle", font=("Arial", 18, "bold")).pack(pady=10)

        # --- SEÇÃO 1: ETIQUETAS ---
        self.frame_etiquetas = ctk.CTkFrame(self.frame_dir, fg_color="#2b2b2b")
        self.frame_etiquetas.pack(pady=5, padx=10, fill="x")
        
        ctk.CTkLabel(self.frame_etiquetas, text="ÁREA DE ETIQUETAS", font=("Arial", 12, "bold"), text_color="silver").pack(pady=5)
        
        self.lbl_detalhe_info = ctk.CTkLabel(self.frame_etiquetas, text="Selecione um paciente na lista ao lado.", justify="center")
        self.lbl_detalhe_info.pack(pady=2)

        self.btn_add_fila = ctk.CTkButton(self.frame_etiquetas, text="⬇️ Adicionar à Fila", command=self.adicionar_selecionado_fila, state="disabled")
        self.btn_add_fila.pack(pady=5)

        self.btn_add_todos = ctk.CTkButton(self.frame_etiquetas, text="⬇️⬇️ ADICIONAR TODOS ATIVOS ⬇️⬇️", fg_color="#1f6aa5", command=self.adicionar_todos_fila)
        self.btn_add_todos.pack(pady=5)

        # Fila Visual
        self.scroll_fila = ctk.CTkScrollableFrame(self.frame_etiquetas, height=120, fg_color="#3a3a3a")
        self.scroll_fila.pack(pady=5, padx=10, fill="x")
        self.lbl_contador_fila = ctk.CTkLabel(self.frame_etiquetas, text="0 etiquetas na fila", text_color="yellow")
        self.lbl_contador_fila.pack(pady=0)

        self.frame_botoes_print = ctk.CTkFrame(self.frame_etiquetas, fg_color="transparent")
        self.frame_botoes_print.pack(pady=10)
        ctk.CTkButton(self.frame_botoes_print, text="Limpar", fg_color="red", width=80, command=self.limpar_fila).pack(side="left", padx=5)
        ctk.CTkButton(self.frame_botoes_print, text="🖨️ IMPRIMIR ETIQUETAS", fg_color="green", width=200, height=40, command=self.gerar_pdf_etiquetas).pack(side="left", padx=5)

        # --- SEÇÃO 2: RELATÓRIOS ---
        ctk.CTkFrame(self.frame_dir, height=2, fg_color="gray").pack(fill="x", pady=15, padx=20) 
        
        ctk.CTkLabel(self.frame_dir, text="📑 Relatórios Gerenciais (Tabelas A4)", font=("Arial", 16, "bold")).pack(pady=5)
        
        self.btn_relatorio_ativos = ctk.CTkButton(self.frame_dir, text="📄 RELATÓRIO DE PACIENTES (Só Ocupados)", 
                                                 fg_color="#D35400", hover_color="#A04000", height=40, width=400,
                                                 command=self.gerar_relatorio_ativos)
        self.btn_relatorio_ativos.pack(pady=5)

        self.btn_relatorio_full = ctk.CTkButton(self.frame_dir, text="📊 IMPRIMIR PLANILHA COMPLETA (Audit)", 
                                                 fg_color="#5B2C6F", hover_color="#4A235A", height=50, width=400,
                                                 command=self.gerar_relatorio_completo_com_vazios)
        self.btn_relatorio_full.pack(pady=10)

    def carregar_dados(self):
        try:
            df_raw = pd.read_excel("pacientes.xlsx")
            df_raw['ENFERMARIA'] = df_raw['ENFERMARIA'].ffill()
            df_raw['LEITO'] = df_raw['LEITO'].astype(str)
            
            self.df_completo = df_raw.copy()
            
            self.df_pacientes = df_raw.dropna(subset=['NOME DO PACIENTE']).copy()
            self.df_pacientes['NOME DO PACIENTE'] = self.df_pacientes['NOME DO PACIENTE'].str.strip()
            
            self.povoar_lista_pacientes(self.df_pacientes)
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao ler Excel: {e}")

    def povoar_lista_pacientes(self, df):
        for widget in self.scroll_pacientes.winfo_children(): widget.destroy()

        if df.empty:
            self.lbl_status_lista.configure(text="Nenhum paciente ativo.")
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
               (self.df_pacientes['LEITO'].str.contains(termo, na=False)) | \
               (self.df_pacientes['DIETA'].astype(str).str.lower().str.contains(termo, na=False))
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
        if messagebox.askyesno("Confirmar", f"Adicionar todas as {len(self.df_pacientes)} etiquetas?"):
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

    # --- GERADOR DE ETIQUETAS (COM QUEBRA DE LINHA) ---
    def gerar_pdf_etiquetas(self):
        if not self.fila_impressao:
            messagebox.showwarning("Vazio", "Fila vazia!")
            return

        arquivo_pdf = "etiquetas_imprimir.pdf"
        c = canvas.Canvas(arquivo_pdf, pagesize=A4)
        
        largura_etiqueta = 95 * mm 
        altura_etiqueta = 52 * mm  
        gap_vertical = 3 * mm      
        colunas, linhas_por_pag = 2, 5
        
        for i, p in enumerate(self.fila_impressao):
            if i > 0 and i % (colunas * linhas_por_pag) == 0: c.showPage()
            
            pos_pag = i % (colunas * linhas_por_pag)
            x = 10*mm + ((pos_pag % colunas) * (largura_etiqueta + 5*mm))
            y = A4[1] - 10*mm - (((pos_pag // colunas) + 1) * (altura_etiqueta + gap_vertical))
            
            self.desenhar_etiqueta_individual(c, x, y, largura_etiqueta, altura_etiqueta, p)
            
        c.save()
        try: os.startfile(arquivo_pdf)
        except: messagebox.showinfo("Sucesso", "Etiquetas geradas!")

    def desenhar_etiqueta_individual(self, c, x, y, w, h, p):
        # Borda
        c.setStrokeColorRGB(0, 0, 0)
        c.rect(x, y, w, h)
        
        # Cabeçalho Fixo
        c.setFont("Helvetica-Bold", 9)
        c.drawCentredString(x + w/2, y + h - 8*mm, "SILVA E TEIXEIRA")
        c.setFont("Helvetica", 7)
        c.drawCentredString(x + w/2, y + h - 12*mm, "IDENTIFICAÇÃO DE DIETAS")

        # Dados Variáveis
        obs = str(p['OBSERVAÇÕES']) if pd.notna(p['OBSERVAÇÕES']) else ''
        dieta = str(p['DIETA']) if pd.notna(p['DIETA']) else ''
        nome = p['NOME DO PACIENTE']
        
        # Função Auxiliar para desenhar texto com quebra (WRAP)
        # Retorna a posição Y onde parou de escrever
        def desenhar_campo_quebrado(canvas_obj, texto_label, texto_valor, pos_x, pos_y, max_width):
            canvas_obj.setFont("Helvetica-Bold", 8)
            
            # Monta o texto completo
            texto_completo = f"{texto_label} {texto_valor}"
            
            # Quebra o texto em linhas que cabem na largura
            linhas = simpleSplit(texto_completo, "Helvetica-Bold", 8, max_width)
            
            for linha in linhas:
                canvas_obj.drawString(pos_x, pos_y, linha)
                pos_y -= 4 * mm # Desce 4mm por linha
            
            return pos_y - 2*mm # Retorna nova posição Y com um pequeno espaço extra

        # Configurações iniciais de desenho
        margem_esq = x + 3*mm
        cursor_y = y + h - 20*mm
        largura_texto = w - 6*mm # Largura disponível para texto

        # Desenhando linha por linha com quebra automática
        cursor_y = desenhar_campo_quebrado(c, "PACIENTE:", nome, margem_esq, cursor_y, largura_texto)
        cursor_y = desenhar_campo_quebrado(c, "ENF:", f"{p['ENFERMARIA']} - LEITO: {p['LEITO']}", margem_esq, cursor_y, largura_texto)
        cursor_y = desenhar_campo_quebrado(c, "DIETA:", dieta, margem_esq, cursor_y, largura_texto)
        cursor_y = desenhar_campo_quebrado(c, "OBS:", obs, margem_esq, cursor_y, largura_texto)
        
        # Data (Sempre cabe em uma linha, desenha direto no fim)
        c.drawString(margem_esq, y + 2*mm, f"DATA: {datetime.now().strftime('%d/%m/%Y')}")

    # --- RELATÓRIOS ---
    def gerar_relatorio_ativos(self):
        self.gerar_tabela_pdf(self.df_pacientes, "relatorio_ativos.pdf", "RELATÓRIO DE PACIENTES INTERNADOS")

    def gerar_relatorio_completo_com_vazios(self):
        self.gerar_tabela_pdf(self.df_completo, "relatorio_completo.pdf", "MAPA GERAL DE LEITOS E DIETAS")

    def gerar_tabela_pdf(self, df_alvo, nome_arquivo, titulo):
        if df_alvo is None or df_alvo.empty:
            messagebox.showwarning("Erro", "Sem dados.")
            return

        doc = SimpleDocTemplate(nome_arquivo, pagesize=landscape(A4), rightMargin=20, leftMargin=20, topMargin=30, bottomMargin=20)
        elements = []
        
        # Estilos de Texto
        styles = getSampleStyleSheet()
        estilo_celula = ParagraphStyle('CellStyle', parent=styles['Normal'], fontSize=8, leading=10) # Texto menor na tabela

        elements.append(Paragraph(f"SILVA E TEIXEIRA - {titulo}", styles['Title']))
        elements.append(Paragraph(f"Emitido em: {datetime.now().strftime('%d/%m/%Y %H:%M')}", styles['Normal']))
        elements.append(Spacer(1, 15))

        # Cabeçalho da Tabela
        data = [['LEITO', 'NOME DO PACIENTE', 'ENFERMARIA', 'DIETA', 'OBSERVAÇÕES']]
        
        for index, row in df_alvo.iterrows():
            nome = str(row['NOME DO PACIENTE']) if pd.notna(row['NOME DO PACIENTE']) else ""
            enf = str(row['ENFERMARIA']) if pd.notna(row['ENFERMARIA']) else ""
            leito = str(row['LEITO']) if pd.notna(row['LEITO']) else ""
            dieta = str(row['DIETA']) if pd.notna(row['DIETA']) else ""
            obs = str(row['OBSERVAÇÕES']) if pd.notna(row['OBSERVAÇÕES']) else ""

            # TRUQUE: Usar Paragraph dentro da célula para quebrar linha automaticamente
            p_nome = Paragraph(nome, estilo_celula)
            p_enf = Paragraph(enf, estilo_celula)
            p_dieta = Paragraph(dieta, estilo_celula)
            p_obs = Paragraph(obs, estilo_celula)
            
            data.append([leito, p_nome, p_enf, p_dieta, p_obs])

        # Larguras
        col_widths = [50, 240, 130, 150, 200]
        t = Table(data, colWidths=col_widths, repeatRows=1)

        # CORES ALTERNADAS (ZEBRA)
        estilo_tabela = TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.darkblue), # Cabeçalho Azul Escuro
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),     # Texto Branco
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'), # Alinhar texto ao topo
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
            
            # MÁGICA DA COR ALTERNADA:
            # Pinta do registro 1 (ignora cabeçalho 0) até o fim (-1)
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.whitesmoke, colors.lightgrey]) 
        ])
        
        t.setStyle(estilo_tabela)
        elements.append(t)

        # Assinatura
        elements.append(Spacer(1, 40))
        elements.append(Paragraph("_"*60, styles['Normal']))
        elements.append(Paragraph("<b>NUTRICIONISTA RESPONSÁVEL</b>", styles['Normal']))

        try:
            doc.build(elements)
            os.startfile(nome_arquivo)
        except Exception as e:
            messagebox.showerror("Erro PDF", f"Erro: {e}")

if __name__ == "__main__":
    app = SistemaEtiquetas()
    app.mainloop()