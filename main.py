import customtkinter as ctk
import pandas as pd
from tkinter import messagebox
import os
from datetime import datetime
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm

# Configuração visual
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

class SistemaEtiquetas(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Gerador de Etiquetas - Hospital (V2.2 - Ajuste Margem)")
        self.geometry("1100x700") 
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

        self.scroll_pacientes = ctk.CTkScrollableFrame(self.frame_esq, width=350, height=500)
        self.scroll_pacientes.pack(pady=10, padx=10, fill="both", expand=True)

        self.lbl_status_lista = ctk.CTkLabel(self.frame_esq, text="Carregando...", text_color="gray")
        self.lbl_status_lista.pack(pady=5)

    def criar_interface_direita(self):
        self.frame_dir = ctk.CTkFrame(self, width=500, corner_radius=10)
        self.frame_dir.grid(row=0, column=1, padx=20, pady=20, sticky="nsew")

        ctk.CTkLabel(self.frame_dir, text="⚙️ Painel de Ação", font=("Arial", 18, "bold")).pack(pady=10)

        # Detalhes
        self.frame_detalhes = ctk.CTkFrame(self.frame_dir, fg_color="#2b2b2b")
        self.frame_detalhes.pack(pady=10, padx=10, fill="x")
        
        self.lbl_detalhe_titulo = ctk.CTkLabel(self.frame_detalhes, text="Selecione um paciente", font=("Arial", 14, "bold"))
        self.lbl_detalhe_titulo.pack(pady=5)
        self.lbl_detalhe_info = ctk.CTkLabel(self.frame_detalhes, text="Clique na lista ao lado.", justify="left")
        self.lbl_detalhe_info.pack(pady=5)

        # --- BOTOES DE ADICIONAR ---
        self.frame_add_btns = ctk.CTkFrame(self.frame_dir, fg_color="transparent")
        self.frame_add_btns.pack(pady=5)

        self.btn_add_fila = ctk.CTkButton(self.frame_add_btns, text="⬇️ Adicionar Este à Fila", command=self.adicionar_selecionado_fila, state="disabled")
        self.btn_add_fila.pack(pady=2, fill="x")

        self.btn_add_todos = ctk.CTkButton(self.frame_add_btns, text="⬇️⬇️ ADICIONAR TODOS DA LISTA ⬇️⬇️", fg_color="#1f6aa5", command=self.adicionar_todos_fila)
        self.btn_add_todos.pack(pady=5, fill="x")
        # ---------------------------

        ctk.CTkLabel(self.frame_dir, text="Fila para Imprimir:", font=("Arial", 14)).pack(pady=(20, 5))
        self.scroll_fila = ctk.CTkScrollableFrame(self.frame_dir, height=200, fg_color="#3a3a3a")
        self.scroll_fila.pack(pady=5, padx=10, fill="x")

        self.lbl_contador_fila = ctk.CTkLabel(self.frame_dir, text="0 etiquetas na fila", text_color="yellow")
        self.lbl_contador_fila.pack(pady=0)

        # Botões Finais
        self.frame_botoes = ctk.CTkFrame(self.frame_dir, fg_color="transparent")
        self.frame_botoes.pack(pady=20)

        self.btn_limpar = ctk.CTkButton(self.frame_botoes, text="Limpar Fila", fg_color="red", hover_color="darkred", width=100, command=self.limpar_fila)
        self.btn_limpar.pack(side="left", padx=5)

        self.btn_gerar_pdf = ctk.CTkButton(self.frame_botoes, text="🖨️ GERAR PDF", fg_color="green", hover_color="darkgreen", width=200, height=40, command=self.gerar_pdf_multiplo)
        self.btn_gerar_pdf.pack(side="left", padx=5)

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
        resposta = messagebox.askyesno("Confirmar Lote", f"Deseja adicionar TODOS os {qtd} pacientes à fila de impressão?")
        
        if resposta:
            for index, row in self.df_pacientes.iterrows():
                self.adicionar_item_visual_fila(row)
            messagebox.showinfo("Sucesso", f"{qtd} etiquetas adicionadas!")

    def adicionar_item_visual_fila(self, row):
        self.fila_impressao.append(row)
        
        frame_item = ctk.CTkFrame(self.scroll_fila, fg_color="transparent")
        frame_item.pack(fill="x", pady=1)
        
        lbl = ctk.CTkLabel(frame_item, text=f"✅ {row['LEITO']} - {row['NOME DO PACIENTE']}", anchor="w", font=("Arial", 12))
        lbl.pack(side="left")
        
        self.atualizar_contador()

    def atualizar_contador(self):
        qtd = len(self.fila_impressao)
        self.lbl_contador_fila.configure(text=f"{qtd} etiquetas na fila")

    def limpar_fila(self):
        self.fila_impressao = []
        for widget in self.scroll_fila.winfo_children():
            widget.destroy()
        self.atualizar_contador()

    # --- AQUI ESTÁ A CORREÇÃO DE TAMANHO ---
    def gerar_pdf_multiplo(self):
        if not self.fila_impressao:
            messagebox.showwarning("Vazio", "Fila vazia!")
            return

        arquivo_pdf = "etiquetas_lote.pdf"
        c = canvas.Canvas(arquivo_pdf, pagesize=A4)
        
        largura_pag, altura_pag = A4
        
        # AJUSTE DE MEDIDAS (Reduzi a altura e o gap)
        largura_etiqueta = 95 * mm 
        altura_etiqueta = 52 * mm  # Era 55mm
        gap_vertical = 3 * mm      # Era 5mm
        
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
            
            # Cálculo corrigido para caber 5 fileiras sem estourar
            # Passo vertical = Altura da etiqueta + Espacinho entre elas
            passo_vertical = altura_etiqueta + gap_vertical
            y = altura_pag - margem_y_top - ((row_atual + 1) * passo_vertical)
            
            self.desenhar_etiqueta(c, x, y, largura_etiqueta, altura_etiqueta, p)
            
        c.save()
        try:
            os.startfile(arquivo_pdf)
        except:
            messagebox.showinfo("Sucesso", "PDF Gerado!")

    def desenhar_etiqueta(self, c, x, y, w, h, p):
        c.setStrokeColorRGB(0, 0, 0)
        c.rect(x, y, w, h)
        
        c.setFont("Helvetica-Bold", 9)
        c.drawCentredString(x + w/2, y + h - 8*mm, "SILVA E TEIXEIRA")
        c.setFont("Helvetica", 7)
        c.drawCentredString(x + w/2, y + h - 12*mm, "IDENTIFICAÇÃO DE DIETAS")

        margem_texto = x + 3*mm
        # Ajustei o topo do texto para acompanhar a nova altura
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

if __name__ == "__main__":
    app = SistemaEtiquetas()
    app.mainloop()