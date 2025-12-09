import webview
import pandas as pd
import os
from datetime import datetime

# --- BIBLIOTECAS DE PDF ---
from reportlab.lib.pagesizes import A4, landscape
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.utils import simpleSplit 

# Variáveis Globais
df_pacientes = None
df_completo = None

class Api:
    
    def carregar_dados_excel(self):
        """Lê o Excel e retorna duas listas: uma filtrada (painel) e uma completa (editor)"""
        global df_pacientes, df_completo
        try:
            if not os.path.exists("pacientes.xlsx"):
                return {"sucesso": False, "erro": "Arquivo 'pacientes.xlsx' não encontrado."}

            df_raw = pd.read_excel("pacientes.xlsx")
            
            # Tratamento visual
            df_raw['ENFERMARIA'] = df_raw['ENFERMARIA'].ffill()
            
            def limpar_leito(val):
                if pd.isna(val) or val == "": return ""
                try: return str(int(float(val)))
                except: return str(val)

            df_raw['LEITO'] = df_raw['LEITO'].apply(limpar_leito)
            
            # 1. Lista Completa (Inclui linhas vazias - Para o Editor)
            df_completo = df_raw.copy()
            lista_editor = df_completo.fillna('').to_dict(orient='records')

            # 2. Lista Filtrada (Só com nomes - Para o Dashboard/Etiquetas)
            df_pacientes = df_raw.dropna(subset=['NOME DO PACIENTE']).copy()
            df_pacientes['NOME DO PACIENTE'] = df_pacientes['NOME DO PACIENTE'].str.strip()
            lista_painel = df_pacientes.fillna('').to_dict(orient='records')
            
            return {
                "sucesso": True, 
                "dados": lista_painel,       # Para a lista lateral e busca
                "dados_editor": lista_editor # Para a tabela de edição
            }
        
        except PermissionError:
            return {"sucesso": False, "erro": "O Excel está aberto! Feche e tente novamente."}
        except Exception as e:
            return {"sucesso": False, "erro": str(e)}

    def salvar_dados_excel(self, novos_dados):
        """Recebe os dados da tabela do HTML e salva no arquivo .xlsx"""
        try:
            df_novo = pd.DataFrame(novos_dados)
            
            # Garante a ordem e existência das colunas
            colunas_ordem = ['ENFERMARIA', 'LEITO', 'NOME DO PACIENTE', 'DIETA', 'OBSERVAÇÕES']
            for col in colunas_ordem:
                if col not in df_novo.columns:
                    df_novo[col] = ""
            
            df_final = df_novo[colunas_ordem]
            
            # Salva sobrescrevendo o arquivo
            df_final.to_excel("pacientes.xlsx", index=False)
            
            # Recarrega as variáveis globais automaticamente
            self.carregar_dados_excel()
            
            return {"sucesso": True, "msg": "Planilha salva com sucesso!"}
            
        except PermissionError:
            return {"sucesso": False, "msg": "Erro: O Excel está aberto. Feche o arquivo e tente novamente."}
        except Exception as e:
            return {"sucesso": False, "msg": f"Erro ao salvar: {str(e)}"}

    def pedir_caminho_salvar(self, nome_sugerido):
        """Abre janela nativa do Windows para salvar arquivo"""
        # --- AQUI ESTAVA O ERRO, AGORA ESTÁ CORRIGIDO ---
        caminho = webview.windows[0].create_file_dialog(
            webview.SAVE_DIALOG, 
            directory='', 
            save_filename=nome_sugerido,
            file_types=('Arquivos PDF (*.pdf)',)
        )
        return caminho 

    def imprimir_etiquetas(self, lista_pacientes):
        if not lista_pacientes: return "Fila vazia!"

        caminho_arquivo = self.pedir_caminho_salvar("etiquetas.pdf")
        if not caminho_arquivo: return "Cancelado."
        
        # Ajuste para garantir que seja string e tenha .pdf
        if isinstance(caminho_arquivo, (tuple, list)): caminho_arquivo = caminho_arquivo[0]
        if not caminho_arquivo.endswith('.pdf'): caminho_arquivo += '.pdf'

        try:
            c = canvas.Canvas(caminho_arquivo, pagesize=A4)
            largura_etiqueta, altura_etiqueta = 95*mm, 52*mm
            gap_vertical = 3*mm
            colunas, linhas_por_pag = 2, 5
            
            for i, p in enumerate(lista_pacientes):
                if i > 0 and i % (colunas * linhas_por_pag) == 0: c.showPage()
                pos_pag = i % (colunas * linhas_por_pag)
                x = 10*mm + ((pos_pag % colunas) * (largura_etiqueta + 5*mm))
                y = A4[1] - 10*mm - (((pos_pag // colunas) + 1) * (altura_etiqueta + gap_vertical))
                desenhar_etiqueta_individual(c, x, y, largura_etiqueta, altura_etiqueta, p)
                
            c.save()
            os.startfile(caminho_arquivo)
            return "PDF salvo e aberto!"
        except Exception as e:
            return f"Erro: {e}"

    def gerar_relatorio_simples(self):
        if df_pacientes is None: return "Erro: Carregue a lista primeiro."
        
        caminho_arquivo = self.pedir_caminho_salvar("relatorio_ocupados.pdf")
        if not caminho_arquivo: return "Cancelado."
        if isinstance(caminho_arquivo, (tuple, list)): caminho_arquivo = caminho_arquivo[0]
        if not caminho_arquivo.endswith('.pdf'): caminho_arquivo += '.pdf'

        try:
            gerar_tabela_pdf(df_pacientes, caminho_arquivo, "PACIENTES OCUPADOS", mesclar=False)
            return "Relatório salvo!"
        except Exception as e:
            return f"Erro: {e}"

    def gerar_mapa_geral(self):
        if df_completo is None: return "Erro: Carregue a lista primeiro."
        
        caminho_arquivo = self.pedir_caminho_salvar("mapa_auditoria.pdf")
        if not caminho_arquivo: return "Cancelado."
        if isinstance(caminho_arquivo, (tuple, list)): caminho_arquivo = caminho_arquivo[0]
        if not caminho_arquivo.endswith('.pdf'): caminho_arquivo += '.pdf'

        try:
            gerar_tabela_pdf(df_completo, caminho_arquivo, "MAPA GERAL (AUDITORIA)", mesclar=True)
            return "Mapa Geral salvo!"
        except Exception as e:
            return f"Erro: {e}"

# --- FUNÇÕES AUXILIARES DE PDF ---

def desenhar_etiqueta_individual(c, x, y, w, h, p):
    c.setStrokeColorRGB(0, 0, 0); c.rect(x, y, w, h)
    c.setFont("Helvetica-Bold", 9); c.drawCentredString(x + w/2, y + h - 8*mm, "SILVA E TEIXEIRA")
    c.setFont("Helvetica", 7); c.drawCentredString(x + w/2, y + h - 12*mm, "IDENTIFICAÇÃO DE DIETAS")
    
    obs = str(p.get('OBSERVAÇÕES', ''))
    dieta = str(p.get('DIETA', ''))
    nome = p.get('NOME DO PACIENTE', '')
    enf = p.get('ENFERMARIA', '')
    leito = str(p.get('LEITO', ''))
    
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
    cursor_y = desenhar_campo_quebrado(c, "ENF:", f"{enf} - LEITO: {leito}", margem_esq, cursor_y, largura_texto)
    cursor_y = desenhar_campo_quebrado(c, "DIETA:", dieta, margem_esq, cursor_y, largura_texto)
    cursor_y = desenhar_campo_quebrado(c, "OBS:", obs, margem_esq, cursor_y, largura_texto)
    c.drawString(margem_esq, y + 2*mm, f"DATA: {datetime.now().strftime('%d/%m/%Y')}")

def gerar_tabela_pdf(df_alvo, nome_arquivo, subtitulo, mesclar=False):
    doc = SimpleDocTemplate(nome_arquivo, pagesize=landscape(A4), rightMargin=20, leftMargin=20, topMargin=20, bottomMargin=20)
    elements = []
    styles = getSampleStyleSheet()
    
    if os.path.exists("logo.png"):
        logo = Image("logo.png", width=15*mm, height=15*mm)
        logo.hAlign = 'CENTER'; elements.append(logo); elements.append(Spacer(1, 10))

    estilo_sub = ParagraphStyle('SubTitle', parent=styles['Normal'], alignment=1, fontSize=10)
    elements.append(Paragraph(f"<b>{subtitulo}</b> - Emitido em: {datetime.now().strftime('%d/%m/%Y %H:%M')}", estilo_sub))
    elements.append(Spacer(1, 15))

    estilo_celula = ParagraphStyle('CellStyle', parent=styles['Normal'], fontSize=9, leading=11)
    
    # Ordem: ENFERMARIA, LEITO, NOME, DIETA, OBS
    data = [['ENFERMARIA', 'LEITO', 'NOME DO PACIENTE', 'DIETA', 'OBSERVAÇÕES']]
    
    for index, row in df_alvo.iterrows():
        nome = str(row['NOME DO PACIENTE']) if pd.notna(row['NOME DO PACIENTE']) else ""
        enf = str(row['ENFERMARIA']) if pd.notna(row['ENFERMARIA']) else ""
        leito = str(row['LEITO']) 
        dieta = str(row['DIETA']) if pd.notna(row['DIETA']) else ""
        obs = str(row['OBSERVAÇÕES']) if pd.notna(row['OBSERVAÇÕES']) else ""

        data.append([
            Paragraph(enf, estilo_celula), leito, Paragraph(nome, estilo_celula),
            Paragraph(dieta, estilo_celula), Paragraph(obs, estilo_celula)
        ])

    col_widths = [110, 50, 250, 160, 200]
    t = Table(data, colWidths=col_widths, repeatRows=1)

    comandos_estilo = [
        ('BACKGROUND', (0, 0), (-1, 0), colors.Color(0.2, 0.6, 0.3)), 
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.whitesmoke, colors.white]) 
    ]

    if mesclar:
        grupo_anterior = None; inicio_grupo = 1 
        df_reset = df_alvo.reset_index(drop=True)
        for i in range(len(df_reset)):
            linha_atual_tabela = i + 1 
            enf_atual = df_reset.iloc[i]['ENFERMARIA']
            if enf_atual != grupo_anterior:
                if grupo_anterior is not None:
                    fim_grupo = linha_atual_tabela - 1
                    comandos_estilo.append(('SPAN', (0, inicio_grupo), (0, fim_grupo)))
                    comandos_estilo.append(('VALIGN', (0, inicio_grupo), (0, fim_grupo), 'MIDDLE'))
                grupo_anterior = enf_atual; inicio_grupo = linha_atual_tabela
        comandos_estilo.append(('SPAN', (0, inicio_grupo), (0, len(df_reset))))
        comandos_estilo.append(('VALIGN', (0, inicio_grupo), (0, len(df_reset)), 'MIDDLE'))

    t.setStyle(TableStyle(comandos_estilo))
    elements.append(t); elements.append(Spacer(1, 40))
    estilo_assinatura = ParagraphStyle('Assinatura', parent=styles['Normal'], alignment=TA_CENTER)
    elements.append(Paragraph("_"*60, estilo_assinatura))
    elements.append(Paragraph("<b>NUTRICIONISTA RESPONSÁVEL</b>", estilo_assinatura))
    doc.build(elements)
    
    # Se o nome do arquivo existe, abre
    if os.path.exists(nome_arquivo): os.startfile(nome_arquivo)

# --- INICIALIZAÇÃO ---
if __name__ == '__main__':
    api = Api()
    webview.create_window(
        'Sistema NutriBem +', 
        'web/index.html', 
        js_api=api,
        width=1200, 
        height=800,
        resizable=True
    )
    webview.start()