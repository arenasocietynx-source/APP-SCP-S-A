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
# Conecta com o Google Sheets usando os segredos configurados
conn = st.connection("gsheets", type=GSheetsConnection)

# --- CONFIGURAÇÕES GERAIS ---
ARQUIVO_ID = "id_controle.txt"
ARQUIVO_DADOS = "historico_solicitacoes.csv"
EMAIL_ALMOXARIFADO = "gcsconsultoriaeservicos@gmail.com" 

# --- FUNÇÃO 1: PROTOCOLO SEQUENCIAL ---
def gerar_novo_protocolo():
    # Lê a aba 'Controle'
    df_controle = conn.read(worksheet="Controle", usecols=[0], header=None)
    
    # Se estiver vazio ou der erro, começa do 0
    if df_controle.empty:
        ultimo_id = 0
    else:
        try:
            ultimo_id = int(df_controle.iloc[0, 0])
        except:
            ultimo_id = 0
            
    novo_id = ultimo_id + 1
    
    # Atualiza o Google Sheets com o novo número
    # Criamos um DataFrame simples com o novo número
    df_novo = pd.DataFrame([novo_id])
    conn.update(worksheet="Controle", data=df_novo, header=False)
    
    return novo_id
    
# --- FUNÇÃO 2: CLASSE PDF (LAYOUT PAISAGEM) ---
class PDF(FPDF):
    def header(self):
        # Imagens (Ajustadas para Paisagem A4 -> Largura aprox 297mm)
        if os.path.exists('logo_esq.png'):
            self.image('logo_esq.png', 10, 8, 30) 
        if os.path.exists('logo_dir.png'):
            self.image('logo_dir.png', 257, 8, 30) 

        self.set_font('Arial', 'B', 15)
        self.cell(0, 10, 'SOLICITAÇÃO DE CADASTRO DE PRODUTOS', 0, 1, 'C')
        self.ln(20)

    def footer(self):
        # --- LINHA 1 (Posição Y = -15mm do fim) ---
        self.set_y(-15) 
        self.set_font('Arial', 'I', 8)
        
        # Cria a variável texto1
        texto1 = f"Gerado e Criado por GCS - Gestão da Cadeia de Suprimentos {chr(174)} 2025"
        
        # Imprime texto1 (Note que mudei 'texto' para 'texto1' aqui dentro)
        self.cell(0, 5, texto1.encode('latin-1', 'replace').decode('latin-1'), 0, 0, 'C')
        
        # --- LINHA 2 (Posição Y = -10mm do fim) ---
        self.set_y(-10) 
        self.set_font('Arial', 'BI', 8)
        
        # Cria a variável texto2
        texto2 = "Simplificando Suprimentos, Impulsionando Negócios"
        
        # Imprime texto2
        self.cell(0, 5, texto2.encode('latin-1', 'replace').decode('latin-1'), 0, 0, 'C')

    
# --- FUNÇÃO 3: GERAR O PDF COM TABELA DE ITENS ---
def gerar_arquivo_pdf(protocolo, cabecalho_dados, df_itens):
    # Orientation 'L' = Landscape (Paisagem)
    pdf = PDF(orientation='L') 
    pdf.add_page()
    pdf.set_font("Arial", size=10)
    
    # 1. Dados do Solicitante
    pdf.set_fill_color(220, 220, 220)
    pdf.set_font("Arial", 'B', 12)
    pdf.cell(0, 10, txt=f"Protocolo: #{protocolo}  |  Data: {cabecalho_dados['Data']}", ln=1, fill=True)
    
    pdf.set_font("Arial", size=11)
    pdf.cell(90, 8, txt=f"Solicitante: {cabecalho_dados['Solicitante']}", border=0)
    pdf.cell(90, 8, txt=f"Departamento: {cabecalho_dados['Departamento']}", border=0)
    pdf.cell(0, 8, txt=f"Tipo: {cabecalho_dados['Tipo']}", ln=1, border=0)
    pdf.ln(5)

    # 2. Tabela de Itens
    pdf.set_font("Arial", 'B', 9) 
    pdf.set_fill_color(200, 220, 255) 
    
    # Larguras
    w_desc = 65
    w_pn = 35
    w_fab = 35
    w_und = 12
    w_app = 50
    w_eq = 35
    w_min = 15
    w_max = 15
    
    # Cabeçalho da Tabela
    pdf.cell(w_desc, 8, "Descrição", 1, 0, 'C', True)
    pdf.cell(w_pn, 8, "PN/Ref", 1, 0, 'C', True)
    pdf.cell(w_fab, 8, "Fabricante", 1, 0, 'C', True)
    pdf.cell(w_und, 8, "UN", 1, 0, 'C', True)
    pdf.cell(w_app, 8, "Aplicação", 1, 0, 'C', True)
    pdf.cell(w_eq, 8, "Equipamento", 1, 0, 'C', True)
    pdf.cell(w_min, 8, "Min", 1, 0, 'C', True)
    pdf.cell(w_max, 8, "Max", 1, 1, 'C', True)

    # Dados da Tabela
    pdf.set_font("Arial", size=8) 
    
    def safe(txt): return str(txt).encode('latin-1', 'replace').decode('latin-1')

    for index, row in df_itens.iterrows():
        # Garante que None vire string vazia
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
    # Lê cofre seguro do Streamlit
    remetente = st.secrets["arenasocietynx@gmail.com"]  # <--- PREENCHER
    senha_app = st.secrets["kaic paey yqdt ckoz"]     # <--- PREENCHER
    
    smtp_server = "smtp.gmail.com"
    smtp_port = 587

    msg = MIMEMultipart()
    msg['From'] = remetente
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

        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(remetente, senha_app)
        server.sendmail(remetente, destinatario, msg.as_string())
        server.quit()
        return True
    except Exception as e:
        st.error(f"Erro no envio: {e}")
        return False

# --- INTERFACE VISUAL (STREAMLIT) ---
st.set_page_config(page_title="Cadastro de Produtos", layout="wide")

# --- CABEÇALHO COM LOGO ---
col_logo, col_titulo = st.columns([1, 6])
with col_logo:
    # Use 'logo_esq.png' ou o nome da sua imagem de logo da interface
    if os.path.exists("logo_esq.png"):
        st.image("logo_esq.png", width=120)
    else:
        st.write("🏭") 

with col_titulo:
    st.title("Solicitação de Cadastro - Sousa & Andrade")
# --------------------------

st.markdown("---")

# 1. CAMPOS DE INPUT (O ERRO ESTAVA NA FALTA DISSO AQUI)
with st.container():
    col1, col2, col3 = st.columns([2, 2, 1])
    with col1:
        solicitante = st.text_input("Nome do Solicitante")
    with col2:
        departamento = st.text_input("Departamento / Área")
    with col3:
        tipo_solicitacao = st.radio("Finalidade:", ("Aplicação Direta", "Estoque"))

st.markdown("### Lista de Itens (Máximo 20)")
st.info("Preencha a tabela abaixo. Para adicionar linhas, clique no '+'.")

# 2. CONFIGURAÇÃO DA TABELA
df_template = pd.DataFrame(columns=[
    "Descrição", "PN/Referência", "Fabricante", "Unidade", "Aplicação", 
    "Equipamento", "Estoque Mín", "Estoque Máx"
])

config_colunas = {
    "Descrição": st.column_config.TextColumn("Descrição do Produto *", width="medium", required=True),
    "PN/Referência": st.column_config.TextColumn("PN ou Ref *", width="small", required=True),
    "Fabricante": st.column_config.TextColumn("Fabricante", width="small"),
    "Unidade": st.column_config.SelectboxColumn("UN *", options=["UN", "PC", "CX", "KG", "MT", "LT", "CJ", "PAR"], width="small", required=True),
    "Aplicação": st.column_config.TextColumn("Aplicação *", width="medium", required=True),
    "Equipamento": st.column_config.TextColumn("Equipamento", width="medium"),
    "Estoque Mín": st.column_config.NumberColumn("Est. Mín", min_value=0, step=1, default=0),
    "Estoque Máx": st.column_config.NumberColumn("Est. Máx", min_value=0, step=1, default=0),
}

itens_preenchidos = st.data_editor(
    df_template,
    column_config=config_colunas,
    num_rows="dynamic",
    use_container_width=True,
    hide_index=True
)

st.markdown("---")
botao_enviar = st.button("Validar e Enviar Solicitação", type="primary")

if botao_enviar:
    erros = []
    
    # --- LIMPEZA DE DADOS (COLCHETES) ---
    def limpar_dados_tabela(x):
        if isinstance(x, list):
            return x[0] if len(x) > 0 else None
        return x

    for col in itens_preenchidos.columns:
        itens_preenchidos[col] = itens_preenchidos[col].apply(limpar_dados_tabela)

    cols_texto = ['Descrição', 'PN/Referência', 'Fabricante', 'Unidade', 'Aplicação', 'Equipamento']
    for col in cols_texto:
        itens_preenchidos[col] = itens_preenchidos[col].apply(lambda x: str(x) if x is not None and str(x) != 'nan' else "")
    # ------------------------------------

    # Validações
    if not solicitante or not departamento:
        erros.append("Preencha o Solicitante e o Departamento no topo.")
    
    if itens_preenchidos.empty:
        erros.append("A tabela de itens está vazia.")

    if len(itens_preenchidos) > 20:
        erros.append(f"O limite é de 20 itens. Você inseriu {len(itens_preenchidos)}.")

    # Duplicidade
    duplicados = itens_preenchidos[itens_preenchidos.duplicated(subset=['Descrição', 'PN/Referência'], keep=False)]
    if not duplicados.empty:
        st.error("⚠️ Item em duplicidade encontrado! Verifique os itens abaixo:")
        st.dataframe(duplicados)
        erros.append("Remova ou corrija os itens duplicados acima.")

    # Estoque
    if "Estoque" in tipo_solicitacao:
        itens_preenchidos['Estoque Mín'] = pd.to_numeric(itens_preenchidos['Estoque Mín'], errors='coerce').fillna(0)
        itens_preenchidos['Estoque Máx'] = pd.to_numeric(itens_preenchidos['Estoque Máx'], errors='coerce').fillna(0)
        
        check_estoque = itens_preenchidos[(itens_preenchidos['Estoque Mín'] <= 0) | (itens_preenchidos['Estoque Máx'] <= 0)]
        if not check_estoque.empty:
            erros.append("Para solicitação de ESTOQUE, preencha Mínimo e Máximo.")

    # Campos Obrigatórios
    if (itens_preenchidos['Descrição'] == '').any():
        erros.append("A coluna 'Descrição' é obrigatória.")

    if erros:
        for e in erros:
            st.warning(e)
    else:
        # --- SUCESSO ---
        protocolo_numero = gerar_novo_protocolo()
        protocolo_formatado = str(protocolo_numero).zfill(4)
        data_hora = datetime.now().strftime('%d/%m/%Y %H:%M')
        
        cabecalho = {
            "Solicitante": solicitante,
            "Departamento": departamento,
            "Tipo": tipo_solicitacao,
            "Data": data_hora
        }

        arquivo_pdf = gerar_arquivo_pdf(protocolo_formatado, cabecalho, itens_preenchidos)
        
        # MODO SIMULAÇÃO (Troque quando for usar real)
        # st.info(f"Simulação: PDF gerado ({arquivo_pdf}) na pasta.")
        # sucesso_email = True 
        
        # MODO REAL
        with st.spinner('Enviando e-mail...'):
              sucesso_email = enviar_email_com_anexo(
                 EMAIL_ALMOXARIFADO, 
                 f"Solicitação Lote #{protocolo_formatado} - {departamento}", 
                 f"Prezados, segue anexo solicitação com {len(itens_preenchidos)} itens.", 
                 arquivo_pdf
             )

        if sucesso_email:
                st.success(f"✅ Solicitação #{protocolo_formatado} enviada com sucesso!")
                st.balloons()
                
                # --- NOVO CÓDIGO DE SALVAMENTO NO GOOGLE SHEETS ---
                
                # 1. Prepara os dados novos
                itens_preenchidos['Protocolo'] = protocolo_formatado
                itens_preenchidos['Data'] = data_hora
                itens_preenchidos['Solicitante'] = solicitante
                itens_preenchidos['Departamento'] = departamento
                itens_preenchidos['Tipo'] = tipo_solicitacao
                
                colunas_ordem = [
                    'Protocolo', 'Data', 'Solicitante', 'Departamento', 'Tipo', 
                    'Descrição', 'PN/Referência', 'Fabricante', 'Unidade', 'Aplicação', 
                    'Equipamento', 'Estoque Mín', 'Estoque Máx'
                ]
                df_novo_registro = itens_preenchidos[colunas_ordem]

                # 2. Lê o histórico antigo da nuvem para não perder nada
                try:
                    df_antigo = conn.read(worksheet="Dados")
                    # Junta o antigo com o novo
                    df_final = pd.concat([df_antigo, df_novo_registro], ignore_index=True)
                except:
                    # Se for a primeira vez e a planilha estiver vazia
                    df_final = df_novo_registro

                # 3. Envia tudo de volta para a nuvem
                conn.update(worksheet="Dados", data=df_final)
