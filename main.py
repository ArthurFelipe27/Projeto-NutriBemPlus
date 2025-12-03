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
from reportlab.lib.styles import getSampleStyleSheet

# Configuração visual
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

class SistemaEtiquetas(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Sistema Hospitalar - Nutrição (V3.0 - Completo)")
        self.geometry("1100x750") 
        self.resizable(False, False)

        # Dados
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

        ctk.CTkLabel(self.frame_esq, text="📋 Lista de Pacientes", font=("Arial", 18, "bold")).pack(pady=10)

        self.entrada_busca = ctk.CTkEntry(self.frame_esq, placeholder_text="Filtrar nome ou leito...", width=300)
        self.entrada_busca.pack(pady=5)
        self.entrada_busca.bind("<KeyRelease>", self.filtrar_lista)

        self.scroll_pacientes = ctk.CTkScrollableFrame(self.frame_esq, width=350, height=480)
        self.scroll_pacientes.pack(pady=10, padx=10, fill="both", expand=True)

        self.lbl_status_lista = ctk.CTkLabel(self.frame_esq, text="Carregando...", text_color="gray")
        self.lbl_status_lista.pack(pady=5)

    def criar_interface_direita(self):
        self.frame_dir = ctk.CTkFrame(self, width=500, corner_radius=10)
        self.frame_dir.grid(row=0, column=1, padx=20, pady=20, sticky="nsew")

        ctk.CTkLabel(self.frame_dir, text="⚙️ Painel de Etiquetas", font=("Arial", 18, "bold")).pack(pady=10)

        # Detalhes
        self.frame_detalhes = ctk.CTkFrame(self.frame_dir, fg_color="#2b2b2b")
        self.frame_detalhes.pack(pady=5, padx=10, fill="x")
        
        self.lbl_detalhe_titulo = ctk.CTkLabel(self.frame_detalhes, text="Selecione um paciente", font=("Arial", 14, "bold"))
        self.lbl_detalhe_titulo.pack(pady=5)
        self.lbl_detalhe_info = ctk.CTkLabel(self.frame_detalhes, text="Clique na lista ao lado.", justify="left")
        self.lbl_detalhe_info.pack(pady=5)

        # Botoes Fila
        self.frame_add_btns = ctk.CTkFrame(self.frame_dir, fg_color="transparent")
        self.frame_add_btns.pack(pady=5)

        self.btn_add_fila = ctk.CTkButton(self.frame_add_btns, text="⬇️ Adicionar Este à Fila", command=self.adicionar_selecionado_fila, state="disabled")
        self.btn_add_fila.pack(pady=2, fill="x")

        self.btn_add_todos = ctk.CTkButton(self.frame_add_btns, text="⬇️⬇️ ADICIONAR TODOS DA LISTA ⬇️⬇️", fg_color="#1f6aa5", command=self.adicionar_todos_fila)
        self.btn_add_todos.pack(pady=5, fill="x")

        # Visual Fila
        self.scroll_fila = ctk.CTkScrollableFrame(self.frame_dir, height=150, fg_color="#3a3a3a")
        self.scroll_fila.pack(pady=5, padx=10, fill="x")
        self.lbl_contador_fila = ctk.CTkLabel(self.frame_dir, text="0 etiquetas na fila", text_color="yellow")
        self.lbl_contador_fila.pack(pady=0)

        # Ações Etiquetas
        self.frame_botoes = ctk.CTkFrame(self.frame_dir, fg_color="transparent")
        self.frame_botoes.pack(pady=10)
        self.btn_limpar = ctk.CTkButton(self.frame_botoes, text="Limpar", fg_color="red", width=80, command=self.limpar_fila)
        self.btn_limpar.pack(side="left", padx=5)
        self.btn_gerar_pdf = ctk.CTkButton(self.frame_botoes, text="🖨️ IMPRIMIR ETIQUETAS", fg_color="green", width=200, height=40, command=self.gerar_pdf_multiplo)
        self.btn_gerar_pdf.pack(side="left", padx=5)

        # --- NOVA ÁREA: RELATÓRIOS ---
        ctk.CTkFrame(self.frame_dir, height=2, fg_color="gray").pack(fill="x", pady=20, padx=20) # Linha divisória
        
        ctk.CTkLabel(self.frame_dir, text="📑 Relatórios Administrativos", font=("Arial", 16, "bold")).pack(pady=5)
        
        self.btn_relatorio_geral = ctk.CTkButton(self.frame_dir, text="📄 GERAR MAPA GERAL DE DIETAS (A4)", 
                                                 fg_color="#D35400", hover_color="#A04000", height=50, width=350,
                                                 command=self.gerar_relatorio_completo)
        self.btn_relatorio_geral.pack(pady=10)


    def carregar_dados(self):
        try:
            self.df_pacientes = pd.read_excel("pacientes.xlsx")
            self.df_pacientes['ENFERMARIA'] = self.df_pacientes['ENFERMARIA'].ffill()
            self.df_pacientes['LEITO'] = self.df_pacientes['LEITO'].astype(str)
            self.df_pacientes = self.df_pacientes.dropna(subset=['NOME DO PACIENTE'])
            self.df_pacientes['NOME DO PACIENTE'] = self.df_pacientes['NOME DO PACIENTE'].str.strip()
            
            self.povoar_lista_pacientes(self.df_pacientes)
        except Exception as e:
            messagebox.showerror("Erro", f"Erro: {e}")

    def povoar_lista_pacientes(self, df):
        for widget in self.scroll_pacientes.winfo_children():
            widget.destroy()

        if df.empty:
            self.lbl_status_lista.configure(text="Nenhum paciente.")
            return

        for index, row in df.iterrows():
            btn = ctk.CTkButton(self.scroll_pacientes, text=f"{row['LEITO']} - {row['NOME DO PACIENTE']}", 
                                fg_color="transparent", border_width=1, border_color="gray",
                                text_color=("black", "white"), anchor="w",
                                command=lambda r=row: self.selecionar_paciente(r))
            btn.pack(pady=2, fill="x")

        self.lbl_status_lista.configure(text=f"{len(df)} pacientes listados.")

    def filtrar_lista(self, event=None):
        if self.df_pacientes is None: return
        termo = self.entrada_busca.get().lower()
        mask = (self.df_pacientes['NOME DO PACIENTE'].str.lower().str.contains(termo, na=False)) | \
               (self.df_pacientes['LEITO'].str.contains(termo, na=False))
        self.povoar_lista_pacientes(self.df_pacientes[mask])

    def selecionar_paciente(self, row):
        self.paciente_selecionado = row
        self.lbl_detalhe_titulo.configure(text=row['NOME DO PACIENTE'])
        info = (f"Enfermaria: {row['ENFERMARIA']}\nLeito: {row['LEITO']}\nDieta: {row['DIETA']}")
        self.lbl_detalhe_info.configure(text=info)
        self.btn_add_fila.configure(state="normal")

    def adicionar_selecionado_fila(self):
        if self.paciente_selecionado is None: return
        self.adicionar_item_visual_fila(self.paciente_selecionado)

    def adicionar_todos_fila(self):
        if self.df_pacientes is None or self.df_pacientes.empty: return
        qtd = len(self.df_pacientes)
        if messagebox.askyesno("Confirmar", f"Adicionar TODOS os {qtd} pacientes?"):
            for index, row in self.df_pacientes.iterrows():
                self.adicionar_item_visual_fila(row)

    def adicionar_item_visual_fila(self, row):
        self.fila_impressao.append(row)
        frame_item = ctk.CTkFrame(self.scroll_fila, fg_color="transparent")
        frame_item.pack(fill="x", pady=1)
        ctk.CTkLabel(frame_item, text=f"✅ {row['LEITO']} - {row['NOME DO PACIENTE']}", anchor="w").pack(side="left")
        self.atualizar_contador()

    def atualizar_contador(self):
        self.lbl_contador_fila.configure(text=f"{len(self.fila_impressao)} etiquetas na fila")

    def limpar_fila(self):
        self.fila_impressao = []
        for widget in self.scroll_fila.winfo_children():
            widget.destroy()
        self.atualizar_contador()

    # --- FUNÇÃO DAS ETIQUETAS (V2.2 - Mantida igual) ---
    def gerar_pdf_multiplo(self):
        if not self.fila_impressao:
            messagebox.showwarning("Vazio", "Fila vazia!")
            return

        arquivo_pdf = "etiquetas_final.pdf"
        c = canvas.Canvas(arquivo_pdf, pagesize=A4)
        
        largura_etiqueta = 95 * mm 
        altura_etiqueta = 52 * mm  
        gap_vertical = 3 * mm      
        
        margem_x = 10 * mm
        margem_y_top = 10 * mm
        
        colunas = 2
        linhas_por_pag = 5 
        
        for i, p in enumerate(self.fila_impressao):
            pos_pag = i % (colunas * linhas_por_pag)
            if i > 0 and pos_pag == 0:
                c.showPage()
            
            col_atual = pos_pag % colunas
            row_atual = pos_pag // colunas
            
            x = margem_x + (col_atual * (largura_etiqueta + 5*mm))
            passo_vertical = altura_etiqueta + gap_vertical
            y = A4[1] - margem_y_top - ((row_atual + 1) * passo_vertical)
            
            self.desenhar_etiqueta(c, x, y, largura_etiqueta, altura_etiqueta, p)
            
        c.save()
        try: os.startfile(arquivo_pdf)
        except: messagebox.showinfo("Sucesso", "Etiquetas geradas!")

    def desenhar_etiqueta(self, c, x, y, w, h, p):
        c.setStrokeColorRGB(0, 0, 0)
        c.rect(x, y, w, h)
        c.setFont("Helvetica-Bold", 9)
        c.drawCentredString(x + w/2, y + h - 8*mm, "SILVA E TEIXEIRA")
        c.setFont("Helvetica", 7)
        c.drawCentredString(x + w/2, y + h - 12*mm, "IDENTIFICAÇÃO DE DIETAS")
        
        margem_texto = x + 3*mm
        topo_texto = y + h - 18*mm 
        espaco = 6*mm
        c.setFont("Helvetica-Bold", 8)
        
        obs = str(p['OBSERVAÇÕES']) if pd.notna(p['OBSERVAÇÕES']) else ''
        dieta = str(p['DIETA']) if pd.notna(p['DIETA']) else ''
        
        itens = [
            f"PACIENTE: {p['NOME DO PACIENTE'][:30]}",
            f"ENF: {p['ENFERMARIA']} - LEITO: {p['LEITO']}",
            f"DIETA: {dieta}",
            f"OBS: {obs[:40]}",
            f"DATA: {datetime.now().strftime('%d/%m/%Y')}"
        ]
        for i, item in enumerate(itens):
            c.drawString(margem_texto, topo_texto - (i*espaco), item)

    # --- NOVA FUNÇÃO: RELATÓRIO GERAL (TABELA) ---
    def gerar_relatorio_completo(self):
        if self.df_pacientes is None or self.df_pacientes.empty:
            messagebox.showwarning("Erro", "Não há dados para gerar relatório.")
            return

        nome_arquivo = "mapa_dietas.pdf"
        
        # Cria o documento em Paisagem (Landscape)
        doc = SimpleDocTemplate(nome_arquivo, pagesize=landscape(A4), rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=18)
        elements = []

        # 1. Título e Data
        styles = getSampleStyleSheet()
        elements.append(Paragraph("MAPA GERAL DE DIETAS - SILVA E TEIXEIRA", styles['Title']))
        elements.append(Paragraph(f"Data de Emissão: {datetime.now().strftime('%d/%m/%Y às %H:%M')}", styles['Normal']))
        elements.append(Spacer(1, 20)) # Espaço

        # 2. Preparar Dados para a Tabela
        # Cabeçalho da Tabela
        data = [['LEITO', 'NOME DO PACIENTE', 'ENFERMARIA', 'DIETA', 'OBSERVAÇÕES']]
        
        # Preenche com dados do DataFrame
        for index, row in self.df_pacientes.iterrows():
            obs = str(row['OBSERVAÇÕES']) if pd.notna(row['OBSERVAÇÕES']) else ''
            dieta = str(row['DIETA']) if pd.notna(row['DIETA']) else ''
            
            # Adiciona linha
            data.append([
                row['LEITO'], 
                row['NOME DO PACIENTE'][:35], # Corta nomes gigantes
                row['ENFERMARIA'],
                dieta,
                obs[:50] # Corta obs gigantes
            ])

        # 3. Configurar Tabela (Largura das colunas em pontos)
        # Total A4 Landscape ~840 pontos.
        # Leito(60), Nome(250), Enf(120), Dieta(150), Obs(200) = 780 (Margem segura)
        t = Table(data, colWidths=[60, 250, 120, 150, 200])

        # 4. Estilo da Tabela (Cores, Bordas)
        t.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey), # Fundo Cinza no Cabeçalho
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke), # Texto Branco no Cabeçalho
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'), # Fonte Negrito Cabeçalho
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige), # Fundo Bege nas linhas
            ('GRID', (0, 0), (-1, -1), 1, colors.black) # Grades Pretas
        ]))
        
        elements.append(t)
        
        # 5. Assinatura no Final
        elements.append(Spacer(1, 50)) # Espaço grande antes da assinatura
        
        texto_assinatura = "_______________________________________________<br/><b>NUTRICIONISTA RESPONSÁVEL</b>"
        elements.append(Paragraph(texto_assinatura, styles['Normal']))

        # Gera o PDF
        try:
            doc.build(elements)
            os.startfile(nome_arquivo)
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível gerar o PDF.\nErro: {e}")

if __name__ == "__main__":
    app = SistemaEtiquetas()
    app.mainloop()