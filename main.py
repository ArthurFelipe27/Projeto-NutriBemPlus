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

        self.title("Gerador de Etiquetas - Hospital (Versão Multi-Impressão)")
        self.geometry("1100x650") # Janela maior para caber tudo
        self.resizable(False, False)

        # Dados
        self.df_pacientes = None
        self.paciente_selecionado = None
        self.fila_impressao = [] # Lista de pacientes para imprimir

        # Layout Principal (2 Colunas)
        self.grid_columnconfigure(0, weight=1) # Esquerda (Lista)
        self.grid_columnconfigure(1, weight=1) # Direita (Controles)

        self.criar_interface_esquerda()
        self.criar_interface_direita()
        
        # Carrega dados ao iniciar
        self.carregar_dados()

    def criar_interface_esquerda(self):
        # Frame da Esquerda
        self.frame_esq = ctk.CTkFrame(self, width=400, corner_radius=10)
        self.frame_esq.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")

        ctk.CTkLabel(self.frame_esq, text="📋 Lista de Pacientes Internados", font=("Arial", 18, "bold")).pack(pady=10)

        # Busca Rápida
        self.entrada_busca = ctk.CTkEntry(self.frame_esq, placeholder_text="Filtrar por nome ou leito...", width=300)
        self.entrada_busca.pack(pady=5)
        self.entrada_busca.bind("<KeyRelease>", self.filtrar_lista) # Filtra enquanto digita

        # Lista Rolável (Scroll)
        self.scroll_pacientes = ctk.CTkScrollableFrame(self.frame_esq, width=350, height=450)
        self.scroll_pacientes.pack(pady=10, padx=10, fill="both", expand=True)

        self.lbl_status_lista = ctk.CTkLabel(self.frame_esq, text="Carregando...", text_color="gray")
        self.lbl_status_lista.pack(pady=5)

    def criar_interface_direita(self):
        # Frame da Direita
        self.frame_dir = ctk.CTkFrame(self, width=500, corner_radius=10)
        self.frame_dir.grid(row=0, column=1, padx=20, pady=20, sticky="nsew")

        ctk.CTkLabel(self.frame_dir, text="⚙️ Painel de Impressão", font=("Arial", 18, "bold")).pack(pady=10)

        # Detalhes do Selecionado
        self.frame_detalhes = ctk.CTkFrame(self.frame_dir, fg_color="#2b2b2b")
        self.frame_detalhes.pack(pady=10, padx=10, fill="x")
        
        self.lbl_detalhe_titulo = ctk.CTkLabel(self.frame_detalhes, text="Nenhum paciente selecionado", font=("Arial", 14, "bold"))
        self.lbl_detalhe_titulo.pack(pady=5)
        self.lbl_detalhe_info = ctk.CTkLabel(self.frame_detalhes, text="Clique em um nome na lista ao lado.", justify="left")
        self.lbl_detalhe_info.pack(pady=5)

        # Botão Adicionar à Fila
        self.btn_add_fila = ctk.CTkButton(self.frame_dir, text="⬇️ Adicionar à Fila de Impressão", command=self.adicionar_a_fila, state="disabled")
        self.btn_add_fila.pack(pady=5)

        # A Fila Visual
        ctk.CTkLabel(self.frame_dir, text="Fila para Imprimir (Folha A4):", font=("Arial", 14)).pack(pady=(20, 5))
        self.scroll_fila = ctk.CTkScrollableFrame(self.frame_dir, height=150, fg_color="#3a3a3a")
        self.scroll_fila.pack(pady=5, padx=10, fill="x")

        # Botões de Ação Final
        self.frame_botoes = ctk.CTkFrame(self.frame_dir, fg_color="transparent")
        self.frame_botoes.pack(pady=20)

        self.btn_limpar = ctk.CTkButton(self.frame_botoes, text="Limpar Fila", fg_color="red", hover_color="darkred", command=self.limpar_fila)
        self.btn_limpar.pack(side="left", padx=5)

        self.btn_gerar_pdf = ctk.CTkButton(self.frame_botoes, text="🖨️ GERAR PDF COM TODOS", fg_color="green", hover_color="darkgreen", height=40, command=self.gerar_pdf_multiplo)
        self.btn_gerar_pdf.pack(side="left", padx=5)

    def carregar_dados(self):
        try:
            self.df_pacientes = pd.read_excel("pacientes.xlsx")
            # Tratamento de dados (Mágica do ffill e limpeza)
            self.df_pacientes['ENFERMARIA'] = self.df_pacientes['ENFERMARIA'].ffill()
            self.df_pacientes['LEITO'] = self.df_pacientes['LEITO'].astype(str)
            self.df_pacientes = self.df_pacientes.dropna(subset=['NOME DO PACIENTE'])
            self.df_pacientes['NOME DO PACIENTE'] = self.df_pacientes['NOME DO PACIENTE'].str.strip()
            
            # Atualiza a lista visual
            self.povoar_lista_pacientes(self.df_pacientes)

        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao ler Excel: {e}")

    def povoar_lista_pacientes(self, df):
        # Limpa a lista atual
        for widget in self.scroll_pacientes.winfo_children():
            widget.destroy()

        if df.empty:
            self.lbl_status_lista.configure(text="Nenhum paciente encontrado.")
            return

        # Cria um botão para cada paciente
        for index, row in df.iterrows():
            nome = row['NOME DO PACIENTE']
            leito = row['LEITO']
            
            # Botão estilo "Card"
            texto_btn = f"{leito} - {nome}"
            btn = ctk.CTkButton(self.scroll_pacientes, text=texto_btn, 
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
        
        df_filtrado = self.df_pacientes[mask]
        self.povoar_lista_pacientes(df_filtrado)

    def selecionar_paciente(self, row):
        self.paciente_selecionado = row
        self.lbl_detalhe_titulo.configure(text=row['NOME DO PACIENTE'])
        
        info = (f"Enfermaria: {row['ENFERMARIA']}\n"
                f"Leito: {row['LEITO']}\n"
                f"Dieta: {row['DIETA']}\n"
                f"Obs: {row['OBSERVAÇÕES'] if pd.notna(row['OBSERVAÇÕES']) else '-'}")
        
        self.lbl_detalhe_info.configure(text=info)
        self.btn_add_fila.configure(state="normal")

    def adicionar_a_fila(self):
        if self.paciente_selecionado is None: return
        
        # Adiciona na lista da memória
        self.fila_impressao.append(self.paciente_selecionado)
        
        # Adiciona visualmente na lista da direita
        nome = self.paciente_selecionado['NOME DO PACIENTE']
        lbl = ctk.CTkLabel(self.scroll_fila, text=f"✅ {nome}", anchor="w")
        lbl.pack(fill="x", padx=5)

    def limpar_fila(self):
        self.fila_impressao = []
        for widget in self.scroll_fila.winfo_children():
            widget.destroy()

    def gerar_pdf_multiplo(self):
        if not self.fila_impressao:
            messagebox.showwarning("Vazio", "Adicione pacientes à fila antes de gerar!")
            return

        arquivo_pdf = "etiquetas_multiplas.pdf"
        c = canvas.Canvas(arquivo_pdf, pagesize=A4)
        
        # --- Lógica de Grid (2 colunas) ---
        largura_pag, altura_pag = A4
        largura_etiqueta = 95 * mm 
        altura_etiqueta = 55 * mm # Ajustado para caber umas 10 por folha
        margem_x = 10 * mm
        margem_y_top = 10 * mm
        
        colunas = 2
        linhas_por_pag = 5 # 2 colunas x 5 linhas = 10 etiquetas por página
        
        for i, p in enumerate(self.fila_impressao):
            # Calcula posição na matriz
            pos_pag = i % (colunas * linhas_por_pag) # 0 a 9
            
            # Se encheu a página, cria nova
            if i > 0 and pos_pag == 0:
                c.showPage()
            
            col_atual = pos_pag % colunas
            row_atual = pos_pag // colunas
            
            x = margem_x + (col_atual * (largura_etiqueta + 5*mm))
            # O Y no PDF começa de baixo para cima. Por isso subtraímos da altura total
            y = altura_pag - margem_y_top - ((row_atual + 1) * (altura_etiqueta + 5*mm))
            
            # Desenha a Etiqueta nesta posição X, Y
            self.desenhar_etiqueta(c, x, y, largura_etiqueta, altura_etiqueta, p)
            
        c.save()
        
        try:
            os.startfile(arquivo_pdf)
            self.limpar_fila() # Limpa após imprimir
        except:
            messagebox.showinfo("Sucesso", "PDF Gerado: etiquetas_multiplas.pdf")

    def desenhar_etiqueta(self, c, x, y, w, h, p):
        # Borda
        c.setStrokeColorRGB(0, 0, 0)
        c.rect(x, y, w, h)
        
        # Cabeçalho
        c.setFont("Helvetica-Bold", 9)
        c.drawCentredString(x + w/2, y + h - 8*mm, "SILVA E TEIXEIRA")
        c.setFont("Helvetica", 7)
        c.drawCentredString(x + w/2, y + h - 12*mm, "IDENTIFICAÇÃO DE DIETAS")

        # Conteúdo
        margem_texto = x + 3*mm
        topo_texto = y + h - 20*mm
        espaco = 6*mm
        
        c.setFont("Helvetica-Bold", 8)
        
        itens = [
            f"PACIENTE: {p['NOME DO PACIENTE'][:30]}", # Corta nome muito longo
            f"ENF: {p['ENFERMARIA']} - LEITO: {p['LEITO']}",
            f"DIETA: {p['DIETA'] if pd.notna(p['DIETA']) else ''}",
            f"OBS: {str(p['OBSERVAÇÕES'])[:35] if pd.notna(p['OBSERVAÇÕES']) else ''}",
            f"DATA: {datetime.now().strftime('%d/%m/%Y')}"
        ]
        
        for i, item in enumerate(itens):
            c.drawString(margem_texto, topo_texto - (i*espaco), item)

if __name__ == "__main__":
    app = SistemaEtiquetas()
    app.mainloop()