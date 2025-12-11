import webview
import pandas as pd
import os
import traceback
from datetime import datetime

# --- BIBLIOTECAS DE PDF ---
from reportlab.lib.pagesizes import A4, landscape
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.lib.colors import HexColor 
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.utils import simpleSplit 

# Variáveis Globais
df_pacientes_enf = None
df_completo_enf = None
df_pacientes_uti = None
df_completo_uti = None
df_pacientes_upa = None
df_completo_upa = None

class Api:
    
    def log_erro(self, msg):
        try:
            with open("log_erros.txt", "a", encoding="utf-8") as f:
                f.write(f"{datetime.now()}: {msg}\n")
        except: pass

    def criar_excel_padrao(self, nome_arquivo):
        try:
            with pd.ExcelWriter(nome_arquivo, engine='openpyxl') as writer:
                pd.DataFrame(columns=['ENFERMARIA', 'LEITO', 'NOME DO PACIENTE', 'DIETA', 'OBSERVAÇÕES']).to_excel(writer, sheet_name="Enfermaria", index=False)
                pd.DataFrame(columns=['LEITO', 'NOME DO PACIENTE', 'DIETA', 'OBSERVAÇÕES']).to_excel(writer, sheet_name="UTI", index=False)
                pd.DataFrame(columns=['LEITO', 'NOME DO PACIENTE', 'DIETA', 'OBSERVAÇÕES']).to_excel(writer, sheet_name="UPA", index=False)
        except Exception as e: self.log_erro(f"Erro criar excel: {e}")

    def carregar_dados_excel(self):
        global df_pacientes_enf, df_completo_enf, df_pacientes_uti, df_completo_uti, df_pacientes_upa, df_completo_upa
        
        arquivo = "pacientes.xlsx"
        if not os.path.exists(arquivo): self.criar_excel_padrao(arquivo)

        try:
            # --- FUNÇÃO HÍBRIDA PARA LEITOS (NÚMERO OU TEXTO) ---
            def limpar_leito(val):
                if pd.isna(val) or str(val).strip() == "": return ""
                try: 
                    # Tenta converter para inteiro se for número (ex: 401.0 -> 401)
                    return str(int(float(val)))
                except: 
                    # Se falhar (ex: "ESTAB 01"), retorna o texto limpo
                    return str(val).strip().upper()

            def normalizar_colunas(df):
                df.columns = [str(col).upper().strip() for col in df.columns]
                return df

            # --- 1. ENFERMARIA ---
            try: df_enf = pd.read_excel(arquivo, sheet_name="Enfermaria")
            except: 
                try: df_enf = pd.read_excel(arquivo, sheet_name=0)
                except: df_enf = pd.DataFrame(columns=['ENFERMARIA', 'LEITO', 'NOME DO PACIENTE', 'DIETA', 'OBSERVAÇÕES'])

            df_enf = normalizar_colunas(df_enf)
            for col in ['ENFERMARIA', 'LEITO', 'NOME DO PACIENTE', 'DIETA', 'OBSERVAÇÕES']:
                if col not in df_enf.columns: df_enf[col] = ""

            if 'ENFERMARIA' in df_enf.columns: df_enf['ENFERMARIA'] = df_enf['ENFERMARIA'].ffill()
            df_enf['LEITO'] = df_enf['LEITO'].apply(limpar_leito)
            
            df_completo_enf = df_enf.copy()
            df_pacientes_enf = df_enf.dropna(subset=['NOME DO PACIENTE']).copy()
            df_pacientes_enf = df_pacientes_enf[df_pacientes_enf['NOME DO PACIENTE'].astype(str).str.strip() != '']

            # --- 2. UTI ---
            try: df_uti = pd.read_excel(arquivo, sheet_name="UTI")
            except: df_uti = pd.DataFrame(columns=['LEITO', 'NOME DO PACIENTE', 'DIETA', 'OBSERVAÇÕES'])

            df_uti = normalizar_colunas(df_uti)
            for col in ['LEITO', 'NOME DO PACIENTE', 'DIETA', 'OBSERVAÇÕES']:
                if col not in df_uti.columns: df_uti[col] = ""

            df_uti['LEITO'] = df_uti['LEITO'].apply(limpar_leito)
            df_completo_uti = df_uti.copy()
            df_pacientes_uti = df_uti.dropna(subset=['NOME DO PACIENTE']).copy()
            df_pacientes_uti = df_pacientes_uti[df_pacientes_uti['NOME DO PACIENTE'].astype(str).str.strip() != '']
            
            df_pacientes_uti['ENFERMARIA'] = "UTI HRMSS"
            df_completo_uti['ENFERMARIA'] = "UTI HRMSS" 

            # --- 3. UPA (NOVO!) ---
            try: df_upa = pd.read_excel(arquivo, sheet_name="UPA")
            except: df_upa = pd.DataFrame(columns=['LEITO', 'NOME DO PACIENTE', 'DIETA', 'OBSERVAÇÕES'])

            df_upa = normalizar_colunas(df_upa)
            for col in ['LEITO', 'NOME DO PACIENTE', 'DIETA', 'OBSERVAÇÕES']:
                if col not in df_upa.columns: df_upa[col] = ""

            # Aplica a limpeza que aceita "ESTAB 01"
            df_upa['LEITO'] = df_upa['LEITO'].apply(limpar_leito)
            
            df_completo_upa = df_upa.copy()
            df_pacientes_upa = df_upa.dropna(subset=['NOME DO PACIENTE']).copy()
            df_pacientes_upa = df_pacientes_upa[df_pacientes_upa['NOME DO PACIENTE'].astype(str).str.strip() != '']
            
            df_pacientes_upa['ENFERMARIA'] = "UPA"
            df_completo_upa['ENFERMARIA'] = "UPA"

            return {
                "sucesso": True, 
                "dados_enf": df_pacientes_enf.fillna('').to_dict(orient='records'),
                "dados_uti": df_pacientes_uti.fillna('').to_dict(orient='records'),
                "dados_upa": df_pacientes_upa.fillna('').to_dict(orient='records'),
                "editor_enf": df_completo_enf.fillna('').to_dict(orient='records'),
                "editor_uti": df_completo_uti.fillna('').to_dict(orient='records'),
                "editor_upa": df_completo_upa.fillna('').to_dict(orient='records')
            }
        except PermissionError: return {"sucesso": False, "erro": "O Excel está aberto! Feche e tente novamente."}
        except Exception as e:
            self.log_erro(f"Erro leitura: {traceback.format_exc()}")
            return {"sucesso": False, "erro": f"Erro leitura: {str(e)}"}

    def salvar_dados_excel(self, dados_enf, dados_uti, dados_upa):
        try:
            with pd.ExcelWriter("pacientes.xlsx", engine='openpyxl') as writer:
                # Enf
                df_enf = pd.DataFrame(dados_enf)
                cols_enf = ['ENFERMARIA', 'LEITO', 'NOME DO PACIENTE', 'DIETA', 'OBSERVAÇÕES']
                for c in cols_enf: 
                    if c not in df_enf.columns: df_enf[c] = ""
                df_enf[cols_enf].to_excel(writer, sheet_name="Enfermaria", index=False)
                
                # UTI
                df_uti = pd.DataFrame(dados_uti)
                cols_rest = ['LEITO', 'NOME DO PACIENTE', 'DIETA', 'OBSERVAÇÕES']
                for c in cols_rest: 
                    if c not in df_uti.columns: df_uti[c] = ""
                df_uti[cols_rest].to_excel(writer, sheet_name="UTI", index=False)

                # UPA
                df_upa = pd.DataFrame(dados_upa)
                for c in cols_rest: 
                    if c not in df_upa.columns: df_upa[c] = ""
                df_upa[cols_rest].to_excel(writer, sheet_name="UPA", index=False)

            self.carregar_dados_excel()
            return {"sucesso": True, "msg": "Planilha salva com sucesso!"}
        except PermissionError: return {"sucesso": False, "msg": "Erro: O Excel está aberto. Feche-o."}
        except Exception as e: return {"sucesso": False, "msg": f"Erro ao salvar: {str(e)}"}

    def pedir_caminho_salvar(self, nome_sugerido):
        try:
            caminho = webview.windows[0].create_file_dialog(
                webview.SAVE_DIALOG, directory='', save_filename=nome_sugerido, file_types=('Arquivos PDF (*.pdf)',)
            )
            if not caminho: return None
            if isinstance(caminho, (tuple, list)):
                if len(caminho) > 0: caminho = caminho[0]
                else: return None
            
            caminho = str(caminho)
            if not caminho.lower().endswith('.pdf'): caminho += '.pdf'
            return caminho
        except Exception as e:
            self.log_erro(f"Erro no diálogo salvar: {e}")
            return None

    def imprimir_etiquetas(self, lista_pacientes):
        if not lista_pacientes: return "Fila vazia!"
        caminho = self.pedir_caminho_salvar("etiquetas.pdf")
        if not caminho: return "Cancelado."

        try:
            c = canvas.Canvas(caminho, pagesize=A4)
            largura, altura = 95*mm, 52*mm
            gap = 3*mm
            colunas, linhas = 2, 5
            
            for i, p in enumerate(lista_pacientes):
                if i > 0 and i % (colunas * linhas) == 0: c.showPage()
                pos = i % (colunas * linhas)
                x = 10*mm + ((pos % colunas) * (largura + 5*mm))
                y = A4[1] - 10*mm - (((pos // colunas) + 1) * (altura + gap))
                desenhar_etiqueta_individual(c, x, y, largura, altura, p)
            
            c.save()
            os.startfile(caminho)
            return "PDF salvo e aberto com sucesso!"
        except PermissionError: return "ERRO: Feche o PDF antes de salvar!"
        except Exception as e: 
            self.log_erro(f"Erro imprimir etiquetas: {traceback.format_exc()}")
            return f"Erro: {e}"

    # --- RELATÓRIOS ---
    def gerar_relatorio_enf(self, tipo):
        df = df_pacientes_enf if tipo == 'simples' else df_completo_enf
        if df is None: return "Carregue os dados."
        nome = "relatorio_enf.pdf" if tipo == 'simples' else "mapa_enf.pdf"
        titulo = "PACIENTES OCUPADOS" if tipo == 'simples' else "MAPA GERAL (AUDITORIA)"
        caminho = self.pedir_caminho_salvar(nome)
        if not caminho: return "Cancelado."
        try:
            gerar_tabela_padrao(df, caminho, titulo, mesclar=(tipo=='geral'))
            return "Relatório Salvo!"
        except PermissionError: return "ERRO: Feche o PDF!"
        except Exception as e: return f"Erro: {e}"

    def gerar_relatorio_uti(self, tipo):
        df = df_pacientes_uti if tipo == 'simples' else df_completo_uti
        if df is None or df.empty: return "Sem dados."
        nome = "relatorio_uti.pdf" if tipo == 'simples' else "mapa_uti.pdf"
        titulo = "NUTRIÇÃO - CORRIDA DE LEITO - UTI HRMSS" 
        caminho = self.pedir_caminho_salvar(nome)
        if not caminho: return "Cancelado."
        try:
            gerar_tabela_especifica(df, caminho, titulo)
            return "Relatório Salvo!"
        except PermissionError: return "ERRO: Feche o PDF!"
        except Exception as e: return f"Erro: {e}"

    def gerar_relatorio_upa(self, tipo):
        df = df_pacientes_upa if tipo == 'simples' else df_completo_upa
        if df is None or df.empty: return "Sem dados na UPA."
        nome = "relatorio_upa.pdf" if tipo == 'simples' else "mapa_upa.pdf"
        titulo = "NUTRIÇÃO - CORRIDA DE LEITO - UPA / SALA VERMELHA / AMARELA" 
        caminho = self.pedir_caminho_salvar(nome)
        if not caminho: return "Cancelado."
        try:
            gerar_tabela_especifica(df, caminho, titulo)
            return "Relatório Salvo!"
        except PermissionError: return "ERRO: Feche o PDF!"
        except Exception as e: return f"Erro: {e}"

# --- DESENHO ETIQUETA (V18) ---
def desenhar_etiqueta_individual(c, x, y, w, h, p):
    TAMANHO_FONTE = 9
    c.setStrokeColorRGB(0, 0, 0); c.setLineWidth(1); c.rect(x, y, w, h)
    cor_header = HexColor('#355a31'); c.setFillColor(cor_header); c.setStrokeColor(cor_header)
    c.roundRect(x + 1*mm, y + h - 15*mm - 1*mm, w - 2*mm, 15*mm, 3*mm, fill=1, stroke=0)
    if os.path.exists("logo.png"): c.drawImage("logo.png", x + 3*mm, y + h - 13*mm, width=12*mm, height=12*mm, mask='auto', preserveAspectRatio=True)
    c.setFillColor(colors.white); c.setFont("Helvetica-Bold", 11)
    c.drawCentredString(x + w/2 + 5*mm, y + h - 6*mm, "SILVA E TEIXEIRA")
    c.setFont("Helvetica-Bold", 7)
    c.drawCentredString(x + w/2 + 5*mm, y + h - 9.5*mm, "IDENTIFICAÇÃO DE DIETAS")
    c.drawCentredString(x + w/2 + 5*mm, y + h - 12.5*mm, "PARA PACIENTES")
    c.setFillColor(colors.black)
    margem_esq = x + 3*mm
    nome = str(p.get('NOME DO PACIENTE', '')); enf = str(p.get('ENFERMARIA', '')); leito = str(p.get('LEITO', ''))
    dieta = str(p.get('DIETA', '')); obs = str(p.get('OBSERVAÇÕES', ''))
    if obs == "nan": obs = ""; 
    if dieta == "nan": dieta = ""
    if len(obs) > 100: obs = obs[:97] + "..."
    
    cursor_y = y + h - 20*mm 
    c.setFont("Helvetica-Bold", TAMANHO_FONTE); c.drawString(margem_esq, cursor_y, "PACIENTE:")
    c.setFont("Helvetica", TAMANHO_FONTE); c.drawString(margem_esq + 19*mm, cursor_y, nome[:40]) 
    cursor_y -= 5*mm 
    c.setFont("Helvetica-Bold", TAMANHO_FONTE); c.drawString(margem_esq, cursor_y, "ENF:")
    c.setFont("Helvetica", TAMANHO_FONTE); c.drawString(margem_esq + 9*mm, cursor_y, enf[:15])
    c.setFont("Helvetica-Bold", TAMANHO_FONTE); c.drawString(margem_esq + 38*mm, cursor_y, "LEITO:")
    c.setFont("Helvetica", TAMANHO_FONTE); c.drawString(margem_esq + 50*mm, cursor_y, leito)
    c.setFont("Helvetica-Bold", TAMANHO_FONTE); c.drawString(margem_esq + 63*mm, cursor_y, "DATA:")
    c.setFont("Helvetica", TAMANHO_FONTE); c.drawString(margem_esq + 74*mm, cursor_y, datetime.now().strftime('%d/%m/%Y'))
    cursor_y -= 5*mm 
    c.setFont("Helvetica-Bold", TAMANHO_FONTE); c.drawString(margem_esq, cursor_y, "TIPO DE DIETA:")
    cursor_y -= 4*mm 
    c.setFont("Helvetica", TAMANHO_FONTE)
    for linha in simpleSplit(dieta, "Helvetica", TAMANHO_FONTE, w - 6*mm): c.drawString(margem_esq, cursor_y, linha); cursor_y -= 4*mm 
    cursor_y -= 1*mm 
    c.setFont("Helvetica-Bold", TAMANHO_FONTE); c.drawString(margem_esq, cursor_y, "OBSERVAÇÃO:")
    cursor_y -= 4*mm 
    c.setFont("Helvetica", TAMANHO_FONTE)
    for linha in simpleSplit(obs, "Helvetica", TAMANHO_FONTE, w - 6*mm): 
        if cursor_y < y + 2*mm: break 
        c.drawString(margem_esq, cursor_y, linha); cursor_y -= 4*mm

def gerar_tabela_padrao(df, nome, tit, mesclar=False):
    doc = SimpleDocTemplate(nome, pagesize=landscape(A4), rightMargin=20, leftMargin=20, topMargin=20, bottomMargin=20)
    elements = []
    styles = getSampleStyleSheet()
    if os.path.exists("logo.png"): logo = Image("logo.png", width=35*mm, height=35*mm); logo.hAlign = 'CENTER'; elements.append(logo); elements.append(Spacer(1, 10))
    estilo_sub = ParagraphStyle('SubTitle', parent=styles['Normal'], alignment=TA_CENTER, fontSize=10)
    elements.append(Paragraph(f"<b>{tit}</b> - Emitido em: {datetime.now().strftime('%d/%m/%Y %H:%M')}", estilo_sub))
    elements.append(Spacer(1, 15))
    estilo_celula = ParagraphStyle('CellStyle', parent=styles['Normal'], fontSize=9, leading=11)
    data = [['ENFERMARIA', 'LEITO', 'NOME DO PACIENTE', 'DIETA', 'OBSERVAÇÕES']]
    for index, row in df.iterrows():
        nome = str(row['NOME DO PACIENTE']) if pd.notna(row['NOME DO PACIENTE']) else ""
        enf = str(row['ENFERMARIA']) if pd.notna(row['ENFERMARIA']) else ""
        leito = str(row['LEITO']); dieta = str(row['DIETA']) if pd.notna(row['DIETA']) else ""; obs = str(row['OBSERVAÇÕES']) if pd.notna(row['OBSERVAÇÕES']) else ""
        data.append([Paragraph(enf, estilo_celula), leito, Paragraph(nome, estilo_celula), Paragraph(dieta, estilo_celula), Paragraph(obs, estilo_celula)])
    t = Table(data, colWidths=[110, 50, 250, 160, 200], repeatRows=1)
    estilo = [('BACKGROUND',(0,0),(-1,0),colors.Color(0.2,0.6,0.3)), ('TEXTCOLOR',(0,0),(-1,0),colors.white), ('GRID',(0,0),(-1,-1),0.5,colors.grey), ('ROWBACKGROUNDS',(0,1),(-1,-1),[colors.whitesmoke,colors.white])]
    if mesclar:
        g=None; ini=1; dfr=df.reset_index(drop=True)
        for i in range(len(dfr)):
            atual=dfr.iloc[i]['ENFERMARIA']
            if atual!=g:
                if g: estilo.append(('SPAN',(0,ini),(0,i))); estilo.append(('VALIGN',(0,ini),(0,i),'MIDDLE'))
                g=atual; ini=i+1
        estilo.append(('SPAN',(0,ini),(0,len(dfr)))); estilo.append(('VALIGN',(0,ini),(0,len(dfr)),'MIDDLE'))
    t.setStyle(TableStyle(estilo)); elements.append(t); elements.append(Spacer(1,40))
    elements.append(Paragraph("_"*60, ParagraphStyle('A', parent=s['Normal'], alignment=TA_CENTER))); elements.append(Paragraph("<b>NUTRICIONISTA RESPONSÁVEL</b>", ParagraphStyle('A', parent=s['Normal'], alignment=TA_CENTER)))
    doc.build(elements)
    if os.path.exists(nome): os.startfile(nome)

# --- GERADOR TABELA ESPECÍFICA (UTI E UPA) ---
def gerar_tabela_especifica(df, nome, tit):
    doc = SimpleDocTemplate(nome, pagesize=landscape(A4), rightMargin=20, leftMargin=20, topMargin=20, bottomMargin=20)
    elements = []
    styles = getSampleStyleSheet()
    if os.path.exists("logo.png"): logo = Image("logo.png", width=35*mm, height=35*mm); logo.hAlign = 'CENTER'; elements.append(logo); elements.append(Spacer(1, 5))
    elements.append(Paragraph(f"DATA: {datetime.now().strftime('%d/%m/%Y')}", ParagraphStyle('DT', parent=styles['Normal'], alignment=TA_CENTER, fontSize=12)))
    elements.append(Paragraph(tit, ParagraphStyle('T', parent=styles['Title'], alignment=TA_CENTER, textColor=colors.darkblue)))
    data = [['LEITO', 'NOME DO PACIENTE', 'DIETA', 'OBSERVAÇÕES']]
    style = ParagraphStyle('C', parent=styles['Normal'], fontSize=10)
    for i, r in df.iterrows(): data.append([str(r['LEITO']), Paragraph(str(r['NOME DO PACIENTE']), style), Paragraph(str(r['DIETA']), style), Paragraph(str(r['OBSERVAÇÕES']), style)])
    t = Table(data, colWidths=[60, 280, 200, 230], repeatRows=1)
    t.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.darkblue), ('TEXTCOLOR',(0,0),(-1,0),colors.white), ('GRID',(0,0),(-1,-1),0.5,colors.grey), ('ROWBACKGROUNDS',(0,1),(-1,-1),[colors.whitesmoke,colors.white])]))
    elements.append(t); elements.append(Spacer(1, 30))
    rodape = [[Paragraph("<b>Nº DE FUNCIONÁRIOS DIA:</b> _______________", styles['Normal']), Paragraph("<b>Nº DE FUNCIONÁRIOS NOITE:</b> _______________", styles['Normal'])]]
    elements.append(Table(rodape, colWidths=[350,350])); elements.append(Spacer(1, 30))
    elements.append(Paragraph("_"*60, ParagraphStyle('A', parent=styles['Normal'], alignment=TA_CENTER))); elements.append(Paragraph("<b>NUTRICIONISTA RESPONSÁVEL</b>", ParagraphStyle('A', parent=styles['Normal'], alignment=TA_CENTER)))
    doc.build(elements)
    if os.path.exists(nome): os.startfile(nome)

if __name__ == '__main__':
    api = Api()
    webview.create_window('Sistema NutriBem +', 'web/index.html', js_api=api, width=1200, height=800, resizable=True)
    webview.start()