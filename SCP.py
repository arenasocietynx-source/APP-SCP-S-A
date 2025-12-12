import streamlit as st
import pandas as pd
from fpdf import FPDF
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os
from datetime import datetime
from streamlit_gsheets import GSheetsConnection 

# --- CONFIGURAÇÃO INICIAL ---
st.set_page_config(page_title="Cadastro GCS (Cloud)", layout="wide")

# --- CONEXÃO COM GOOGLE SHEETS ---
conn = st.connection("gsheets", type=GSheetsConnection)

# --- LEITURA DAS SENHAS (SECRETS) ---
# Como o diagnóstico confirmou que elas existem, lemos diretamente.
if "EMAIL_CONTA" in st.secrets:
    EMAIL_CONTA = st.secrets["EMAIL_CONTA"]
    EMAIL_SENHA = st.secrets["EMAIL_SENHA"]
else:
    st.error("ERRO CRÍTICO: As senhas de e-mail não foram encontradas no Secrets.")
    st.stop()

# --- FUNÇÃO 1: PROTOCOLO SEQUENCIAL ---
def gerar_novo_protocolo():
    try:
        df_controle = conn.read(worksheet="Controle", usecols=[0], ttl=0)
        if df_controle.empty:
            ultimo_id = 0
        else:
            valor = df_controle.iloc[0, 0]
            ultimo_id = int(str(valor).replace(',', '').replace('.', ''))
    except Exception:
        ultimo_id = 0
            
    novo_id = ultimo_id + 1
    
    # Atualiza a planilha
    df_novo_numero = pd.DataFrame({'ULTIMO_ID': [novo_id]})
    conn.update(worksheet="Controle", data=df_novo_numero)
    
    return novo_id

# --- FUNÇÃO 2: CLASSE PDF ---
class PDF(FPDF):
    def header(self):
        if os.path.exists('logo_esq.png'):
            self.image('logo_esq.png', 10, 8, 30) 
        if os.path.exists('logo_dir.png'):
            self.image('logo_dir.png', 257, 8, 30) 

        self.set_font('Arial', 'B', 15)
        self.cell(0, 10, 'SOLICITAÇÃO DE CADASTRO (LOTE)', 0, 1, 'C')
        self.ln(20)

    def footer(self):
        self.set_y(-15) 
        self.set_font('Arial', 'I', 8)
        texto1 = f"Gerado e Criado por GCS - Gestão da Cadeia de Suprimentos {chr(174)} 2025"
        self.cell(0, 5, texto1.encode('latin-1', 'replace').decode('latin-1'), 0, 0, 'C')
        
        self.set_y(-10) 
        self.set_font('Arial', 'BI', 8)
        texto2 = "Simplificando Suprimentos, Impulsionando Negócios"
        self.cell(0, 5, texto2.encode('latin-1', 'replace').decode('latin-1'), 0, 0, 'C')

# --- FUNÇÃO 3: GERAR ARQUIVO PDF ---
def gerar_arquivo_pdf(protocolo, cabecalho_dados, df_itens):
    pdf = PDF(orientation='L') 
    pdf.add_page()
    pdf.set_font("Arial", size=10)
    
    # Cabeçalho
    pdf.set_fill_color(220, 220, 220)
    pdf.set_font("Arial", 'B', 12)
    pdf.cell(0, 10, txt=f"Protocolo: #{protocolo}  |  Data: {cabecalho_dados['Data']}", ln=1, fill=True)
    
    pdf.set_font("Arial", size=11)
    pdf.cell(90, 8, txt=f"Solicitante: {cabecalho_dados['Solicitante']}", border=0)
    pdf.cell(90, 8, txt=f"Departamento: {cabecalho_dados['Departamento']}", border=0)
    pdf.cell(0, 8, txt=f"Tipo: {cabecalho_dados['Tipo']}", ln=1, border=0)
    pdf.ln(5)

    # Tabela
    pdf.set_font("Arial", 'B', 9) 
    pdf.set_fill_color(200, 220, 255) 
    
    w_desc, w_pn, w_fab, w_und = 65, 35, 35, 12
    w_app, w_eq, w_min, w_max = 50, 35, 15, 15
    
    pdf.cell(w_desc, 8, "Descrição", 1, 0, 'C', True)
    pdf.cell(w_pn, 8, "PN/Ref", 1, 0, 'C', True)
    pdf.cell(w_fab, 8, "Fabricante", 1, 0, 'C', True)
    pdf.cell(w_und, 8, "UN", 1, 0, 'C', True)
    pdf.cell(w_app, 8, "Aplicação", 1, 0, 'C', True)
    pdf.cell(w_eq, 8, "Equipamento", 1, 0, 'C', True)
    pdf.cell(w_min, 8, "Min", 1, 0, 'C', True)
    pdf.cell(w_max, 8, "Max", 1, 1, 'C', True)

    pdf.set_font("Arial", size=8) 
    def safe(txt): return str(txt).encode('latin-1', 'replace').decode('latin-1')

    for index, row in df_itens.iterrows():
        fab_txt = row['Fabricante'] if row['Fabricante'] else ""
        pdf.cell(w_desc, 7, safe(row['Descrição'])[:40], 1)
        pdf.cell(w_pn, 7, safe(row['PN/Referência'])[:18], 1)
        pdf.cell(w_fab, 7, safe(fab_txt)[:18], 1)
        pdf.cell(w_und, 7, safe(row['Unidade']), 1, 0, 'C')
        pdf.cell(w_app, 7, safe(row['Aplicação'])[:25], 1)
        pdf.cell(w_eq, 7, safe(row['Equipamento'])[:18], 1)
        pdf.cell(w_min, 7, str(row['Estoque Mín']), 1, 0, 'C')
        pdf.cell(w_max, 7, str(row['Estoque Máx']), 1, 1, 'C')

    nome_arquivo = f"Solicitacao_{protocolo}.pdf"
    pdf.output(nome_arquivo)
    return nome_arquivo

# --- FUNÇÃO 4: ENVIO DE E-MAIL ---
def enviar_email_com_anexo(destinatario, assunto, corpo, arquivo):
    msg = MIMEMultipart()
    msg['From'] = EMAIL_CONTA
    msg['To'] = destinatario
    msg['Subject'] = assunto
    msg.attach(MIMEText(corpo, 'plain'))

    try:
        with open(arquivo, "rb") as anexo:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(anexo.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f"attachment; filename= {arquivo}")
            msg.attach(part)

        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(EMAIL_CONTA, EMAIL_SENHA)
        server.sendmail(EMAIL_CONTA, destinatario, msg.as_string())
        server.quit()
        return True
    except Exception as e:
        st.error(f"Erro no envio de e-mail: {e}")
        return False

# --- INTERFACE VISUAL ---
col_logo, col_titulo = st.columns([1, 6])
with col_logo:
    if os.path.exists("logo_esq.png"):
        st.image("logo_esq.png", width=120)
    else:
        st.write("🏭") 

with col_titulo:
    st.title("Solicitação de Cadastro - Sousa & Andrade")

st.markdown("---")

with st.container():
    col1, col2, col3 = st.columns([2, 2, 1])
    with col1:
        solicitante = st.text_input("Nome do Solicitante")
    with col2:
        departamento = st.text_input("Departamento / Área")
    with col3:
        tipo_solicitacao = st.radio("Finalidade:", ("Aplicação Direta", "Estoque"))

st.markdown("### Lista de Itens (Máximo 20)")
st.info("Preencha a tabela abaixo.")

df_template = pd.DataFrame(columns=[
    "Descrição", "PN/Referência", "Fabricante", "Unidade", "Aplicação", 
    "Equipamento", "Estoque Mín", "Estoque Máx"
])

config_colunas = {
    "Descrição": st.column_config.TextColumn("Descrição *", width="medium", required=True),
    "PN/Referência": st.column_config.TextColumn("PN/Ref *", width="small", required=True),
    "Fabricante": st.column_config.TextColumn("Fabricante", width="small"),
    "Unidade": st.column_config.SelectboxColumn("UN *", options=["UN", "PC", "CX", "KG", "MT", "LT", "CJ", "PAR"], width="small", required=True),
    "Aplicação": st.column_config.TextColumn("Aplicação *", width="medium", required=True),
    "Equipamento": st.column_config.TextColumn("Equipamento", width="medium"),
    "Estoque Mín": st.column_config.NumberColumn("Min", min_value=0, step=1, default=0),
    "Estoque Máx": st.column_config.NumberColumn("Max", min_value=0, step=1, default=0),
}

itens_preenchidos = st.data_editor(df_template, column_config=config_colunas, num_rows="dynamic", use_container_width=True, hide_index=True)

st.markdown("---")

# --- DEFINE QUEM RECEBE O E-MAIL ---
# Mudei aqui para enviar para o seu próprio e-mail para você testar se chega
EMAIL_DESTINO = "gcsconsultoriaeservicos@gmail.com" 

if st.button("Validar e Enviar Solicitação", type="primary"):
    erros = []
    
    # Limpeza de dados
    def limpar_dados_tabela(x):
        return x[0] if isinstance(x, list) and len(x) > 0 else x

    for col in itens_preenchidos.columns:
        itens_preenchidos[col] = itens_preenchidos[col].apply(limpar_dados_tabela)

    cols_texto = ['Descrição', 'PN/Referência', 'Fabricante', 'Unidade', 'Aplicação', 'Equipamento']
    for col in cols_texto:
        itens_preenchidos[col] = itens_preenchidos[col].apply(lambda x: str(x) if x is not None and str(x) != 'nan' else "")

    if not solicitante or not departamento: erros.append("Preencha Solicitante e Departamento.")
    if itens_preenchidos.empty: erros.append("Tabela vazia.")
    if len(itens_preenchidos) > 20: erros.append("Limite de 20 itens excedido.")
    
    duplicados = itens_preenchidos[itens_preenchidos.duplicated(subset=['Descrição', 'PN/Referência'], keep=False)]
    if not duplicados.empty:
        st.error("Itens duplicados encontrados:")
        st.dataframe(duplicados)
        erros.append("Corrija as duplicidades.")

    if "Estoque" in tipo_solicitacao:
        itens_preenchidos['Estoque Mín'] = pd.to_numeric(itens_preenchidos['Estoque Mín'], errors='coerce').fillna(0)
        itens_preenchidos['Estoque Máx'] = pd.to_numeric(itens_preenchidos['Estoque Máx'], errors='coerce').fillna(0)
        check = itens_preenchidos[(itens_preenchidos['Estoque Mín'] <= 0) | (itens_preenchidos['Estoque Máx'] <= 0)]
        if not check.empty: erros.append("Para Estoque, Min e Max devem ser > 0.")

    if (itens_preenchidos['Descrição'] == '').any(): erros.append("Descrição obrigatória.")

    if erros:
        for e in erros: st.warning(e)
    else:
        # EXECUÇÃO DO PROCESSO
        with st.spinner("Processando... (Conectando Google Sheets + Gerando PDF + Enviando E-mail)"):
            
            # 1. Gerar Protocolo
            protocolo_numero = gerar_novo_protocolo()
            protocolo_formatado = str(protocolo_numero).zfill(4)
            data_hora = datetime.now().strftime('%d/%m/%Y %H:%M')
            
            cabecalho = {
                "Solicitante": solicitante, "Departamento": departamento,
                "Tipo": tipo_solicitacao, "Data": data_hora
            }

            # 2. Gerar PDF
            arquivo_pdf = gerar_arquivo_pdf(protocolo_formatado, cabecalho, itens_preenchidos)
            
            # 3. Enviar E-mail
            sucesso_email = enviar_email_com_anexo(
                EMAIL_DESTINO, 
                f"Solicitação #{protocolo_formatado} - {departamento}", 
                f"Segue anexo com {len(itens_preenchidos)} itens.", 
                arquivo_pdf
            )

            if sucesso_email:
                st.success(f"✅ Solicitação #{protocolo_formatado} enviada com sucesso!")
                st.balloons()
                
                # 4. Salvar Histórico na Nuvem
                itens_preenchidos['Protocolo'] = protocolo_formatado
                itens_preenchidos['Data'] = data_hora
                itens_preenchidos['Solicitante'] = solicitante
                itens_preenchidos['Departamento'] = departamento
                itens_preenchidos['Tipo'] = tipo_solicitacao
                
                colunas_ordem = ['Protocolo', 'Data', 'Solicitante', 'Departamento', 'Tipo', 'Descrição', 'PN/Referência', 'Fabricante', 'Unidade', 'Aplicação', 'Equipamento', 'Estoque Mín', 'Estoque Máx']
                df_novo = itens_preenchidos[colunas_ordem]

                try:
                    df_antigo = conn.read(worksheet="Dados", ttl=0)
                    df_final = pd.concat([df_antigo, df_novo], ignore_index=True)
                    conn.update(worksheet="Dados", data=df_final)
                except:
                    conn.update(worksheet="Dados", data=df_novo)

