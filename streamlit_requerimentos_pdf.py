import io
import os
import re
import zipfile
from datetime import datetime

import openpyxl
import streamlit as st
from pypdf import PdfReader, PdfWriter
from pypdf.generic import BooleanObject, NameObject, TextStringObject

# ==========================================
# CONFIG
# ==========================================
TEMPLATE_PDF_PATH = "modelo.pdf"

st.set_page_config(page_title="Requerimentos", page_icon="📄", layout="wide")
st.title("📄 Gerador de Requerimentos")

# ==========================================
# CAMPOS REAIS DO PDF
# ==========================================
PDF_FIELDS = {
    "cabecalho": {
        "cartorio": "Text1",
    },
    "nascimento": {
        "check": "1",
        "nome": "Nome_1",
        "termo": "TextField",
        "fls": "TextField_1",
        "livro": "TextField_2",
    },
    "casamento": {
        "check": "TextField_3",
        "nome1": "Text3",
        "nome2": "Nome 2",
        "termo": "TextField_4",
        "fls": "TextField_5",
        "livro": "TextField_6",
    },
    "obito": {
        "check": "TextField_7",
        "nome": "Nome_2",
        "termo": "Termo nº",
        "fls": "Fls",
        "livro": "Livro",
    },
    "especificacoes": {
        "digitada": "TextField_16",
        "fotocopia": "TextField_17",
        "duas": "TextField_18",
        "firma_nao": "TextField_20",
        "haia_nao": "Sim(Se positivo o serviço será cobrado)",
    },
    "rodape": {
        "local": "Local",
        "data": "Data",
    },
}

# ==========================================
# UTIL
# ==========================================
def limpar_texto(valor) -> str:
    if valor is None:
        return ""
    return str(valor).strip()


def formatar_valor_excel(valor) -> str:
    if valor is None:
        return ""

    if isinstance(valor, datetime):
        return valor.strftime("%d/%m/%Y")

    if isinstance(valor, float):
        if valor.is_integer():
            return str(int(valor))
        return str(valor).replace(".", ",")

    return str(valor).strip()


def sanitizar_nome_arquivo(valor: str) -> str:
    valor = limpar_texto(valor)
    valor = re.sub(r'[\\/:*?"<>|]+', "", valor)
    valor = re.sub(r"\s+", "_", valor)
    return valor[:120] if valor else "SemNome"


def normalizar_tipo(valor: str) -> str:
    v = limpar_texto(valor).lower()
    if "casamento" in v:
        return "casamento"
    if "óbito" in v or "obito" in v:
        return "obito"
    if "nascimento" in v:
        return "nascimento"
    raise ValueError(f"Tipo de certidão inválido: {valor}")


def normalizar_formato(valor: str) -> str:
    v = limpar_texto(valor).lower()
    if "digitada" in v:
        return "digitada"
    if "fotocópia" in v or "fotocopia" in v:
        return "fotocopia"
    if "duas" in v:
        return "duas"
    raise ValueError(f"Formato inválido na coluna H: {valor}")


# ==========================================
# LEITURA EXCEL
# ==========================================
def carregar_excel(arquivo_excel) -> list:
    wb = openpyxl.load_workbook(arquivo_excel, data_only=True)
    ws = wb.active

    registros = []

    for i in range(2, ws.max_row + 1):
        nome = ws[f"A{i}"].value
        cartorio = ws[f"B{i}"].value
        tipo = ws[f"C{i}"].value
        conjuge = ws[f"D{i}"].value
        termo = ws[f"E{i}"].value
        fls = ws[f"F{i}"].value
        livro = ws[f"G{i}"].value
        formato = ws[f"H{i}"].value
        local = ws[f"I{i}"].value
        data = ws[f"J{i}"].value

        if not any([nome, cartorio, tipo]):
            continue

        registros.append(
            {
                "linha": i,
                "nome": formatar_valor_excel(nome),
                "cartorio": formatar_valor_excel(cartorio),
                "tipo": formatar_valor_excel(tipo),
                "conjuge": formatar_valor_excel(conjuge),
                "termo": formatar_valor_excel(termo),
                "fls": formatar_valor_excel(fls),
                "livro": formatar_valor_excel(livro),
                "formato": formatar_valor_excel(formato),
                "local": formatar_valor_excel(local) or "Rio de Janeiro",
                "data": formatar_valor_excel(data) or datetime.today().strftime("%d/%m/%Y"),
            }
        )

    return registros


# ==========================================
# APARÊNCIA DOS CAMPOS
# ==========================================
def configurar_aparencia_campos(writer: PdfWriter):
    root = writer._root_object

    if "/AcroForm" in root:
        acroform = root["/AcroForm"]
        acroform.update(
            {
                NameObject("/NeedAppearances"): BooleanObject(True),
                NameObject("/DA"): TextStringObject("/Helv 0 Tf 0 g"),
            }
        )


# ==========================================
# MONTAR CAMPOS
# ==========================================
def montar_campos_pdf(registro: dict) -> dict:
    tipo = normalizar_tipo(registro["tipo"])
    formato = normalizar_formato(registro["formato"])

    campos = {
        PDF_FIELDS["cabecalho"]["cartorio"]: registro["cartorio"],
        PDF_FIELDS["rodape"]["local"]: registro["local"],
        PDF_FIELDS["rodape"]["data"]: registro["data"],
        PDF_FIELDS["especificacoes"]["firma_nao"]: "X",
        PDF_FIELDS["especificacoes"]["haia_nao"]: "X",
    }

    if tipo == "nascimento":
        campos.update({
            "1": "X",
            "Nome_1": registro["nome"],
            "TextField": registro["termo"],
            "TextField_1": registro["fls"],
            "TextField_2": registro["livro"],
        })

    elif tipo == "casamento":
        campos.update({
            "TextField_3": "X",
            "Text3": registro["nome"],
            "Nome 2": registro["conjuge"],
            "TextField_4": registro["termo"],
            "TextField_5": registro["fls"],
            "TextField_6": registro["livro"],
        })

    elif tipo == "obito":
        campos.update({
            "TextField_7": "X",
            "Nome_2": registro["nome"],
            "Termo nº": registro["termo"],
            "Fls": registro["fls"],
            "Livro": registro["livro"],
        })

    if formato == "digitada":
        campos["TextField_16"] = "X"
    elif formato == "fotocopia":
        campos["TextField_17"] = "X"
    else:
        campos["TextField_18"] = "X"

    return campos


# ==========================================
# GERAR PDF NORMAL
# ==========================================
def gerar_pdf_preenchido(template_bytes, registro):
    reader = PdfReader(io.BytesIO(template_bytes))
    writer = PdfWriter()

    for page in reader.pages:
        writer.add_page(page)

    configurar_aparencia_campos(writer)

    campos = montar_campos_pdf(registro)

    writer.update_page_form_field_values(writer.pages[0], campos)

    buffer = io.BytesIO()
    writer.write(buffer)
    return buffer.getvalue()


# ==========================================
# "IMPRIMIR" PDF (CORREÇÃO DO GOV)
# ==========================================
def flatten_pdf(pdf_bytes):
    reader = PdfReader(io.BytesIO(pdf_bytes))
    writer = PdfWriter()

    for page in reader.pages:
        writer.add_page(page)

    # remove formulário → vira PDF "impresso"
    if "/AcroForm" in writer._root_object:
        del writer._root_object["/AcroForm"]

    buffer = io.BytesIO()
    writer.write(buffer)
    return buffer.getvalue()


def nome_saida(registro):
    tipo = normalizar_tipo(registro["tipo"]).capitalize()
    return f"Requerimento_{tipo}_{sanitizar_nome_arquivo(registro['nome'])}.pdf"


# ==========================================
# UI
# ==========================================
file = st.file_uploader("Excel", type=["xlsx"])

if not os.path.exists(TEMPLATE_PDF_PATH):
    st.error("modelo.pdf não encontrado")

elif file:
    template = open(TEMPLATE_PDF_PATH, "rb").read()
    dados = carregar_excel(file)

    zip_buffer = io.BytesIO()

    with zipfile.ZipFile(zip_buffer, "w") as z:
        for reg in dados:
            pdf = gerar_pdf_preenchido(template, reg)
            pdf_final = flatten_pdf(pdf)  # 🔥 AQUI ESTÁ A MÁGICA
            z.writestr(nome_saida(reg), pdf_final)

    zip_buffer.seek(0)

    st.download_button(
        "📦 Baixar PDFs",
        data=zip_buffer,
        file_name="requerimentos.zip"
    )
