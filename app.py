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
    "juliano": "senha123",
    "admin": "oficio456"
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
        model="gpt-4o",
        response_format={"type": "json_object"},
        messages=[
            {"role": "system", "content": """
            Você é um escritor de cartas oficiais muito confiável...
            (instruções truncadas aqui para brevidade, mantenha completas no seu arquivo)
            """},
            {"role": "user", "content": f"Demand: '{demanda}'"}
        ],
    )
    response_json = json.loads(response.choices[0].message.content)
    texto = response_json["texto"]
    paragrafos = texto.split("\n\n")
    if len(paragrafos) != 3:
        texto_unico = texto.replace("\n\n", " ")
        partes = len(texto_unico) // 3
        paragrafos = [
            texto_unico[:partes],
            texto_unico[partes:partes * 2],
            texto_unico[partes * 2:]
        ]
    response_json["texto"] = "\n\n".join(paragrafos)
    return response_json

def preencher_docx(num_oficio, ano_oficio, assunto, dt_envio, parag1, parag2, parag3):
    doc = Document("layout_oficio.docx")
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

    nome_arquivo = f"Vereador Professor Juliano Lopes_N° {num_oficio}-{ano_oficio}_{assunto.replace(' ', '-').replace('/', '_')}.docx"
    doc.save(nome_arquivo)
    return nome_arquivo

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

st.title("📄 Gerador de Ofícios")

demanda = st.text_area("Insira a demanda:", height=200)
num_oficio = st.text_input("Número do Ofício:")
ano_oficio = st.text_input("Ano do Ofício:", value=str(date.today().year))
dt_envio = st.date_input("Data de Envio:", value=date.today())

if "oficio_data" not in st.session_state:
    st.session_state.oficio_data = {}

if st.button("Gerar Ofício"):
    if not demanda or not num_oficio or not ano_oficio or not dt_envio:
        st.error("Por favor, preencha todos os campos.")
    else:
        resposta = gera_oficio(demanda, client)
        assunto = resposta["assunto"]
        resumo = resposta["resumo"]
        texto = resposta["texto"]
        paragrafos = texto.split("\n\n")

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

if st.session_state.oficio_data:
    st.subheader("Assunto:")
    assunto_edit = st.text_input("Assunto", value=st.session_state.oficio_data["assunto"])
    st.subheader("Resumo:")
    resumo_edit = st.text_area("Resumo", value=st.session_state.oficio_data["resumo"], height=100)
    st.subheader("Texto do Ofício:")
    parag1_edit = st.text_area("Parágrafo 1", value=st.session_state.oficio_data["parag1"], height=100)
    parag2_edit = st.text_area("Parágrafo 2", value=st.session_state.oficio_data["parag2"], height=100)
    parag3_edit = st.text_area("Parágrafo 3", value=st.session_state.oficio_data["parag3"], height=100)

    if st.button("Salvar Ofício Editado"):
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
            label="📥 Baixar Ofício",
            data=open(nome_arquivo, "rb"),
            file_name=nome_arquivo,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

