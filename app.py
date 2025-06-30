import streamlit as st
import os
import json
from openai import OpenAI
from datetime import date, datetime
from docx import Document
from babel.dates import format_date

# ---------------------------
# LOGIN + LOG DE ACESSO
# ---------------------------
import gspread
from oauth2client.service_account import ServiceAccountCredentials

USUARIOS = {
    "herbert": "12345",
    "katia": "12345"
}

def log_acesso_google_sheets(usuario):
    try:
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        #creds = ServiceAccountCredentials.from_json_keyfile_name("credenciais.json", scope)
        creds = ServiceAccountCredentials.from_json_keyfile_name("/etc/secrets/credenciais.json", scope)
        client = gspread.authorize(creds)
        sheet = client.open("Log Acessos").worksheet("Acessos")
        agora = datetime.now().strftime("%d/%m/%Y %H:%M")
        sheet.append_row([usuario, agora])
    except Exception as e:
        st.warning(f"Erro ao registrar log: {e}")

def login():
    st.title("🔒 Login necessário")
    usuario = st.text_input("Usuário")
    senha = st.text_input("Senha", type="password")
    if st.button("Entrar"):
        if usuario in USUARIOS and USUARIOS[usuario] == senha:
            st.session_state["logado"] = True
            st.session_state["usuario_logado"] = usuario
            log_acesso_google_sheets(usuario)
        else:
            st.error("Usuário ou senha incorretos.")

if "logado" not in st.session_state:
    st.session_state["logado"] = False

if not st.session_state["logado"]:
    login()
    st.stop()

# ---------------------------
# FUNÇÕES DO APLICATIVO
# ---------------------------

def gera_oficio(demanda, client):
    response = client.chat.completions.create(
        model="gpt-4o", #"gpt-4-turbo-preview",
        response_format={"type": "json_object"},
        messages=[
            {"role": "system", "content": """
            Você é um escritor de cartas oficiais muito confiável.
            Você receberá uma demanda de um cidadão e será responsável por criar
            cartas às autoridades públicas solicitando intervenções para resolver problemas comunitários
            em formato json.
            
            A saída deve ter os elementos:
            1. Um termo com no máximo 3 palavras que identifica o assunto da demanda. a chave desse elemento deve ser "assunto".
            2. Um resumo do tema da demanda. a chave desse elemento deve ser "resumo".
            3. O texto da carta. a chave desse elemento deve ser "texto".
            
            Por favor, preste atenção:
            - O texto deve começar com a seguinte expressão: Cumprimentando-o cordialmente, encaminho a V. Ex.ª...
            - O texto deve ser escrito em primeira pessoa do singular, é uma carta de um vereador para o secretário municipal.
            - O texto deve ter um parágrafo apresentando a demanda,
               um segundo parágrafo explicando os detalhes e
               um terceiro parágrafo concluindo o pedido.
            - Os parágrafos devem estar separados por quebra de linha dupla e é obrigatório que haja 3 parágrafos no resultado.
            - O texto deve terminar com a expressão: ... Por oportuno, agradeço a atenção despendida e renovo meus votos de estima e consideração.
            - O texto deve estar em português.            
            - Não se esqueça que o elemento texto deve estar organizado em 3 parágrafos. Caso não esteja, reorganize para enviar a resposta.
            """},
            {"role": "user", "content": f"Demand: '{demanda}'"}
        ],
    )

    response_json = json.loads(response.choices[0].message.content)
    
    # Divide o texto em parágrafos com base em quebras de linha duplas
    texto = response_json["texto"]
    paragrafos = texto.split("\n\n")

    # Verifica se existem 3 parágrafos, caso contrário divide manualmente
    if len(paragrafos) != 3:
        # Se o texto não tiver 3 parágrafos, forçamos a divisão
        texto_unico = texto.replace("\n\n", " ")  # Junta tudo em um único bloco
        partes = len(texto_unico) // 3
        paragrafos = [
            texto_unico[:partes],
            texto_unico[partes:partes * 2],
            texto_unico[partes * 2:]
        ]
    
    response_json["texto"] = "\n\n".join(paragrafos)
    
    return response_json

# Função para substituir marcadores de posição no arquivo DOCX
def preencher_docx(num_oficio, ano_oficio, assunto, dt_envio, parag1, parag2, parag3):
    doc = Document("/etc/secrets/layout_oficio.docx")
    for p in doc.paragraphs:
        if "{{Num/Ano}}" in p.text:
            p.text = p.text.replace("{{Num/Ano}}", f"{num_oficio}-{ano_oficio}")
        if "{{Assunto}}" in p.text:
            p.text = p.text.replace("{{Assunto}}", assunto)
        if "{{DT. Envio}}" in p.text:
            p.text = p.text.replace("{{DT. Envio}}", dt_envio)
        if "{{Parag. 1}}" in p.text:
            p.text = p.text.replace("{{Parag. 1}}", parag1)
        if "{{Parag. 2}}" in p.text:
            p.text = p.text.replace("{{Parag. 2}}", parag2)
        if "{{Parag. 3}}" in p.text:
            p.text = p.text.replace("{{Parag. 3}}", parag3)
    
    # Gera o nome do arquivo sem caracteres especiais
    nome_arquivo = f"Vereador Professor Juliano Lopes_N° {num_oficio}-{ano_oficio}_{assunto.replace(' ', '-').replace('/', '_')}.docx"
    doc.save(nome_arquivo)
    return nome_arquivo

# Função para converter a data para o formato por extenso
def formatar_data_por_extenso(data):
    return format_date(data, format="long", locale="pt_BR")

# ---------------------------
# CHAVE API OPENAI
# ---------------------------

api_key = os.getenv("OPENAI_API_KEY")
client = OpenAI(api_key=api_key)

# ---------------------------
# INTERFACE STREAMLIT
# ---------------------------

# Título do app
st.title("Gerador de Ofícios")

# Campos para inserir a demanda, número, ano e data de envio
demanda = st.text_area("Insira a demanda:", height=200)
num_oficio = st.text_input("Número do Ofício:")
ano_oficio = st.text_input("Ano do Ofício:", value=str(date.today().year))
dt_envio = st.date_input("Data de Envio:", value=date.today())

# Inicializa o estado da sessão para armazenar os resultados gerados
if "oficio_data" not in st.session_state:
    st.session_state.oficio_data = {}

# Botão para gerar o ofício
if st.button("Gerar Ofício"):
    # Verificação se os campos necessários estão preenchidos
    if not demanda or not num_oficio or not ano_oficio or not dt_envio:
        st.error("Por favor, preencha todos os campos.")
    else:
        # Chamada da função para gerar o ofício
        resposta = gera_oficio(demanda, client)
        
        # Processamento da resposta
        assunto = resposta["assunto"]
        resumo = resposta["resumo"]
        texto = resposta["texto"]
        paragrafos = texto.split("\n\n")
        
        # Verificação para garantir que temos exatamente 3 parágrafos
        if len(paragrafos) == 3:
            st.session_state.oficio_data = {
                "assunto": assunto,
                "resumo": resumo,
                "parag1": paragrafos[0],
                "parag2": paragrafos[1],
                "parag3": paragrafos[2],
                "num_oficio": num_oficio,
                "ano_oficio": ano_oficio,
                "dt_envio": dt_envio
            }
        else:
            st.error("Erro ao gerar o texto do ofício. Tente novamente.")

# Exibe os campos editáveis e o botão para salvar se houver dados gerados
if st.session_state.oficio_data:
    st.subheader("Assunto:")
    assunto_edit = st.text_input("Assunto", value=st.session_state.oficio_data["assunto"])
    st.subheader("Resumo:")
    resumo_edit = st.text_area("Resumo", value=st.session_state.oficio_data["resumo"], height=100)
    st.subheader("Texto do Ofício:")
    parag1_edit = st.text_area("Parágrafo 1", value=st.session_state.oficio_data["parag1"], height=100)
    parag2_edit = st.text_area("Parágrafo 2", value=st.session_state.oficio_data["parag2"], height=100)
    parag3_edit = st.text_area("Parágrafo 3", value=st.session_state.oficio_data["parag3"], height=100)

    # Botão para salvar o arquivo DOCX editado
    if st.button("Salvar Ofício Editado"):
        # Preenchimento do DOCX
        nome_arquivo = preencher_docx(
            st.session_state.oficio_data["num_oficio"],
            st.session_state.oficio_data["ano_oficio"],
            assunto_edit,
            formatar_data_por_extenso(st.session_state.oficio_data["dt_envio"]),
            parag1_edit,
            parag2_edit,
            parag3_edit
        )
        
        st.success("Ofício editado salvo com sucesso!")
        st.download_button(
            label="Baixar Ofício",
            data=open(nome_arquivo, "rb"),
            file_name=nome_arquivo,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
