from __future__ import annotations

from datetime import date, datetime
from pathlib import Path
import re
import unicodedata
from typing import Any

import pandas as pd
import streamlit as st


APP_DIR = Path(__file__).resolve().parent
DATA_DIR = APP_DIR / "data"
DATA_FILE = DATA_DIR / "entregas_produtores.xlsx"
ORIGINAL_FILE = Path(r"C:\Users\lucassilverio\Downloads\CONTROLE DE MERCADORIAS.xlsx")
SHEET_NAME = "Lancamentos"
GOOGLE_SHEET_GID_PADRAO = 448939446

COLS = ["data", "produtor", "fruta", "quantidade", "destino", "origem", "observacao"]
FRUTAS_PADRAO = ["GOIABA", "BANANA"]


st.set_page_config(
    page_title="Lancamento de Entregas",
    layout="wide",
    initial_sidebar_state="expanded",
)


def normalizar_texto(valor: object) -> str:
    texto = "" if pd.isna(valor) else str(valor).strip()
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(char for char in texto if not unicodedata.combining(char))
    return re.sub(r"\s+", " ", texto).strip().upper()


def normalizar_colunas(df: pd.DataFrame) -> pd.DataFrame:
    aliases = {
        "DATA": "data",
        "PRODUTOR": "produtor",
        "FRUTA": "fruta",
        "QUANTIDADE": "quantidade",
        "QTD": "quantidade",
        "KG": "quantidade",
        "DESTINO": "destino",
        "DE ONDE E": "destino",
        "ONDE": "destino",
        "ORIGEM": "origem",
        "OBSERVACAO": "observacao",
        "OBS": "observacao",
    }
    renomear = {}
    for coluna in df.columns:
        chave = normalizar_texto(coluna)
        if chave in aliases:
            renomear[coluna] = aliases[chave]
    return df.rename(columns=renomear)


def parece_data(valor: object) -> bool:
    if pd.isna(valor):
        return False
    if isinstance(valor, (datetime, date, pd.Timestamp)):
        return True
    return bool(re.match(r"^\d{1,2}/\d{1,2}(/\d{2,4})?$", str(valor).strip()))


def converter_data(valor: object, ano_padrao: int = 2026) -> pd.Timestamp | Any:
    if pd.isna(valor):
        return pd.NaT
    if isinstance(valor, (datetime, date, pd.Timestamp)):
        return pd.to_datetime(valor).normalize()

    texto = str(valor).strip()
    if re.match(r"^\d{1,2}/\d{1,2}$", texto):
        texto = f"{texto}/{ano_padrao}"

    data = pd.to_datetime(texto, dayfirst=True, errors="coerce")
    if pd.isna(data):
        data = pd.to_datetime(texto, errors="coerce")
    return data.normalize() if not pd.isna(data) else pd.NaT


def converter_quantidade(valor: object) -> float | None:
    if pd.isna(valor):
        return None
    texto = str(valor).upper().replace("KG", "").replace(".", "").replace(",", ".").strip()
    numero = pd.to_numeric(texto, errors="coerce")
    if pd.isna(numero) or float(numero) <= 0:
        return None
    return float(numero)


def mes_para_numero(nome_aba: str) -> int | None:
    meses = {
        "JANEIRO": 1,
        "FEVEREIRO": 2,
        "MARCO": 3,
        "MARÇO": 3,
        "ABRIL": 4,
        "MAIO": 5,
        "JUNHO": 6,
        "JULHO": 7,
        "AGOSTO": 8,
        "SETEMBRO": 9,
        "OUTUBRO": 10,
        "NOVEMBRO": 11,
        "DEZEMBRO": 12,
    }
    aba = normalizar_texto(nome_aba)
    for nome, numero in meses.items():
        if normalizar_texto(nome) in aba:
            return numero
    return None


def corrigir_mes_da_aba(data_movimento: pd.Timestamp, nome_aba: str) -> pd.Timestamp:
    mes = mes_para_numero(nome_aba)
    if pd.isna(data_movimento) or mes is None:
        return data_movimento
    return pd.Timestamp(year=data_movimento.year, month=mes, day=min(data_movimento.day, 28 if mes == 2 else 31))


def importar_planilha_original(caminho: Path) -> pd.DataFrame:
    if not caminho.exists():
        return pd.DataFrame(columns=COLS)

    registros: list[dict[str, object]] = []
    excel = pd.ExcelFile(caminho)

    for aba in excel.sheet_names:
        grade = pd.read_excel(caminho, sheet_name=aba, header=None)
        linhas, colunas = grade.shape

        for linha in range(linhas):
            valores = [normalizar_texto(grade.iat[linha, col]) for col in range(colunas)]
            colunas_data = [idx for idx, valor in enumerate(valores) if valor == "DATA"]

            for col_data in colunas_data:
                produtor = "PRODUTOR NAO INFORMADO"
                if linha > 0 and normalizar_texto(grade.iat[linha - 1, col_data]):
                    produtor = normalizar_texto(grade.iat[linha - 1, col_data])

                mapa_frutas: dict[int, str] = {}
                coluna_destino = None
                for col in range(col_data + 1, min(colunas, col_data + 5)):
                    cabecalho = normalizar_texto(grade.iat[linha, col])
                    if cabecalho in FRUTAS_PADRAO:
                        mapa_frutas[col] = cabecalho
                    if "ONDE" in cabecalho or "DESTINO" in cabecalho:
                        coluna_destino = col

                if not mapa_frutas or coluna_destino is None:
                    continue

                for row in range(linha + 1, linhas):
                    data_bruta = grade.iat[row, col_data]
                    if not parece_data(data_bruta):
                        continue

                    data_movimento = converter_data(data_bruta)
                    data_movimento = corrigir_mes_da_aba(data_movimento, aba)
                    destino = normalizar_texto(grade.iat[row, coluna_destino])
                    if not destino:
                        destino = "DESTINO NAO INFORMADO"

                    for col_fruta, fruta in mapa_frutas.items():
                        quantidade = converter_quantidade(grade.iat[row, col_fruta])
                        if quantidade is None:
                            continue
                        registros.append(
                            {
                                "data": data_movimento.date() if not pd.isna(data_movimento) else pd.NaT,
                                "produtor": produtor,
                                "fruta": fruta,
                                "quantidade": quantidade,
                                "destino": destino,
                                "origem": f"Importado - {aba}",
                                "observacao": "",
                            }
                        )

    return pd.DataFrame(registros, columns=COLS)


def preparar_base(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=COLS)

    base = normalizar_colunas(df.copy())
    for coluna in COLS:
        if coluna not in base.columns:
            base[coluna] = ""

    base = base[COLS]
    base["data"] = pd.to_datetime(base["data"], errors="coerce").dt.date
    base["produtor"] = base["produtor"].map(normalizar_texto)
    base["fruta"] = base["fruta"].map(normalizar_texto)
    base["destino"] = base["destino"].map(normalizar_texto)
    base["origem"] = base["origem"].fillna("").astype(str)
    base["observacao"] = base["observacao"].fillna("").astype(str)
    base["quantidade"] = pd.to_numeric(base["quantidade"], errors="coerce").fillna(0)
    base = base[base["quantidade"] > 0]
    return base.sort_values(["data", "produtor", "destino", "fruta"], na_position="last").reset_index(drop=True)


def google_sheets_configurado() -> bool:
    try:
        return bool(st.secrets.get("google_sheet_id") and st.secrets.get("gcp_service_account"))
    except Exception:
        return False


def worksheet_tem_base_valida(worksheet) -> bool:
    valores = worksheet.get_all_values()
    if not valores:
        return False
    return encontrar_linha_cabecalho(valores) is not None


def encontrar_linha_cabecalho(valores: list[list[str]]) -> int | None:
    obrigatorias = {"DATA", "PRODUTOR", "FRUTA", "QUANTIDADE", "DESTINO"}
    for indice, linha in enumerate(valores[:20]):
        cabecalhos = {normalizar_texto(valor) for valor in linha}
        if obrigatorias.issubset(cabecalhos):
            return indice
    return None


def dataframe_do_worksheet(worksheet) -> pd.DataFrame:
    valores = worksheet.get_all_values()
    if not valores:
        return pd.DataFrame(columns=COLS)

    linha_cabecalho = encontrar_linha_cabecalho(valores)
    if linha_cabecalho is None:
        return pd.DataFrame(columns=COLS)

    cabecalhos = valores[linha_cabecalho]
    linhas = []
    for linha in valores[linha_cabecalho + 1 :]:
        linha_ajustada = (linha + [""] * len(cabecalhos))[: len(cabecalhos)]
        linhas.append(linha_ajustada)
    return pd.DataFrame(linhas, columns=cabecalhos)


def obter_worksheet_por_gid(planilha, gid: int):
    for worksheet in planilha.worksheets():
        if getattr(worksheet, "id", None) == gid:
            return worksheet
    return None


def abrir_worksheet_google():
    import gspread
    from google.oauth2.service_account import Credentials

    escopos = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    credenciais = Credentials.from_service_account_info(
        dict(st.secrets["gcp_service_account"]),
        scopes=escopos,
    )
    cliente = gspread.authorize(credenciais)
    planilha = cliente.open_by_key(st.secrets["google_sheet_id"])

    gid_configurado = int(st.secrets.get("google_sheet_gid", GOOGLE_SHEET_GID_PADRAO))
    worksheet_por_gid = obter_worksheet_por_gid(planilha, gid_configurado)
    if worksheet_por_gid is not None:
        if not worksheet_tem_base_valida(worksheet_por_gid):
            valores = worksheet_por_gid.get_all_values()
            if not valores:
                worksheet_por_gid.update([COLS])
        return worksheet_por_gid

    try:
        worksheet = planilha.worksheet(SHEET_NAME)
    except gspread.WorksheetNotFound:
        for candidata in planilha.worksheets():
            if worksheet_tem_base_valida(candidata):
                return candidata

        worksheet = planilha.add_worksheet(title=SHEET_NAME, rows=1000, cols=len(COLS))
        worksheet.update([COLS])

    valores = worksheet.get_all_values()
    if not valores:
        worksheet.update([COLS])
    elif not worksheet_tem_base_valida(worksheet):
        worksheet.update("A1:G1", [COLS])
    return worksheet


def carregar_base_google_sheets() -> pd.DataFrame:
    worksheet = abrir_worksheet_google()
    return preparar_base(dataframe_do_worksheet(worksheet))


def salvar_base_google_sheets(df: pd.DataFrame) -> None:
    worksheet = abrir_worksheet_google()
    base = preparar_base(df)
    dados = base.copy()
    dados["data"] = pd.to_datetime(dados["data"], errors="coerce").dt.strftime("%Y-%m-%d")
    dados = dados.fillna("")
    worksheet.clear()
    worksheet.update([COLS] + dados.astype(str).values.tolist())


@st.cache_data(show_spinner=False)
def carregar_base() -> pd.DataFrame:
    if google_sheets_configurado():
        return carregar_base_google_sheets()

    DATA_DIR.mkdir(exist_ok=True)
    if DATA_FILE.exists():
        return preparar_base(pd.read_excel(DATA_FILE))

    base = importar_planilha_original(ORIGINAL_FILE)
    base = preparar_base(base)
    salvar_base(base)
    return base


def salvar_base(df: pd.DataFrame) -> None:
    if google_sheets_configurado():
        salvar_base_google_sheets(df)
        return

    DATA_DIR.mkdir(exist_ok=True)
    with pd.ExcelWriter(DATA_FILE, engine="openpyxl") as writer:
        preparar_base(df).to_excel(writer, index=False, sheet_name=SHEET_NAME)


def opcoes(df: pd.DataFrame, coluna: str, extras: list[str] | None = None) -> list[str]:
    valores = [] if df.empty else sorted(v for v in df[coluna].dropna().astype(str).unique() if v.strip())
    for item in extras or []:
        item = normalizar_texto(item)
        if item and item not in valores:
            valores.append(item)
    return sorted(valores)


def filtro_multiselect(rotulo: str, valores: list[str]) -> list[str]:
    if not valores:
        return []
    return st.sidebar.multiselect(rotulo, valores, default=valores)


def meses_disponiveis(df: pd.DataFrame) -> list[str]:
    datas = pd.to_datetime(df["data"], errors="coerce").dropna()
    if datas.empty:
        return []
    meses = datas.dt.to_period("M").drop_duplicates().sort_values()
    return [periodo.strftime("%m/%Y") for periodo in meses]


def aplicar_filtros(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df

    st.sidebar.header("Filtros")
    meses = meses_disponiveis(df)
    mes_selecionado = st.sidebar.selectbox("Mes", ["Todos os meses"] + meses)

    min_data = min(df["data"])
    max_data = max(df["data"])
    intervalo = st.sidebar.date_input("Periodo", value=(min_data, max_data), min_value=min_data, max_value=max_data)

    produtores = filtro_multiselect("Produtores", opcoes(df, "produtor"))
    frutas = filtro_multiselect("Frutas", opcoes(df, "fruta"))
    destinos = filtro_multiselect("Destinos", opcoes(df, "destino"))

    filtrado = df.copy()
    if mes_selecionado != "Todos os meses":
        mes_periodo = pd.Period(mes_selecionado, freq="M")
        datas = pd.to_datetime(filtrado["data"], errors="coerce")
        filtrado = filtrado[datas.dt.to_period("M") == mes_periodo]

    if isinstance(intervalo, tuple) and len(intervalo) == 2:
        inicio, fim = intervalo
        filtrado = filtrado[(filtrado["data"] >= inicio) & (filtrado["data"] <= fim)]
    if produtores:
        filtrado = filtrado[filtrado["produtor"].isin(produtores)]
    if frutas:
        filtrado = filtrado[filtrado["fruta"].isin(frutas)]
    if destinos:
        filtrado = filtrado[filtrado["destino"].isin(destinos)]
    return filtrado


def resumo_tabela(df: pd.DataFrame, dimensao: str) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=[dimensao, "quantidade", "entregas"])
    return (
        df.groupby(dimensao, as_index=False)
        .agg(quantidade=("quantidade", "sum"), entregas=("quantidade", "count"))
        .sort_values("quantidade", ascending=False)
    )


def renderizar_lancamento(base: pd.DataFrame) -> None:
    st.subheader("Novo lancamento")

    produtores = opcoes(base, "produtor")
    frutas = opcoes(base, "fruta", FRUTAS_PADRAO)
    destinos = opcoes(base, "destino")

    with st.form("form_lancamento", clear_on_submit=True):
        col1, col2, col3, col4 = st.columns([1, 1.2, 1, 1.2])
        data_lancamento = col1.date_input("Data", value=date.today())
        produtor_escolhido = col2.selectbox("Produtor", ["NOVO PRODUTOR"] + produtores)
        fruta_escolhida = col3.selectbox("Fruta", ["NOVA FRUTA"] + frutas)
        destino_escolhido = col4.selectbox("Destino", ["NOVO DESTINO"] + destinos)

        col5, col6, col7 = st.columns([1.2, 1.2, 2])
        produtor_novo = col5.text_input("Nome do produtor", disabled=produtor_escolhido != "NOVO PRODUTOR")
        fruta_nova = col6.text_input("Nome da fruta", disabled=fruta_escolhida != "NOVA FRUTA")
        destino_novo = col7.text_input("Nome do destino", disabled=destino_escolhido != "NOVO DESTINO")

        col8, col9 = st.columns([1, 3])
        quantidade = col8.number_input("Quantidade (kg)", min_value=0.0, step=1.0, format="%.2f")
        observacao = col9.text_input("Observacao")
        salvar = st.form_submit_button("Salvar lancamento", type="primary")

    if not salvar:
        return

    produtor = normalizar_texto(produtor_novo if produtor_escolhido == "NOVO PRODUTOR" else produtor_escolhido)
    fruta = normalizar_texto(fruta_nova if fruta_escolhida == "NOVA FRUTA" else fruta_escolhida)
    destino = normalizar_texto(destino_novo if destino_escolhido == "NOVO DESTINO" else destino_escolhido)

    erros = []
    if not produtor:
        erros.append("Informe o produtor.")
    if not fruta:
        erros.append("Informe a fruta.")
    if not destino:
        erros.append("Informe o destino.")
    if quantidade <= 0:
        erros.append("Informe uma quantidade maior que zero.")

    if erros:
        for erro in erros:
            st.error(erro)
        return

    novo = pd.DataFrame(
        [
            {
                "data": data_lancamento,
                "produtor": produtor,
                "fruta": fruta,
                "quantidade": float(quantidade),
                "destino": destino,
                "origem": "Lancamento manual",
                "observacao": observacao,
            }
        ],
        columns=COLS,
    )
    salvar_base(pd.concat([base, novo], ignore_index=True))
    st.cache_data.clear()
    st.success("Lancamento salvo com sucesso.")
    st.rerun()


def renderizar_importacao(base: pd.DataFrame) -> None:
    with st.expander("Importar novamente a planilha original ou outro arquivo"):
        arquivo = st.file_uploader("Arquivo .xlsx no mesmo formato", type=["xlsx"])
        col1, col2 = st.columns([1, 4])
        substituir = col1.checkbox("Substituir base atual")
        if st.button("Importar arquivo", disabled=arquivo is None):
            DATA_DIR.mkdir(exist_ok=True)
            caminho_temp = DATA_DIR / "_importacao_temp.xlsx"
            try:
                caminho_temp.write_bytes(arquivo.getvalue())
                importado = preparar_base(importar_planilha_original(caminho_temp))
            except Exception as exc:
                st.error(f"Nao foi possivel ler o arquivo enviado: {exc}")
                return
            finally:
                caminho_temp.unlink(missing_ok=True)

            if importado.empty:
                st.warning(
                    "Nenhum lancamento foi encontrado. Confira se o arquivo tem os blocos com DATA, GOIABA, BANANA e DE ONDE E."
                )
                return

            try:
                if substituir:
                    nova_base = importado
                else:
                    nova_base = pd.concat([base, importado], ignore_index=True)
                nova_base = nova_base.drop_duplicates(subset=COLS)
                salvar_base(nova_base)
            except Exception as exc:
                st.error(f"Os dados foram lidos, mas nao foi possivel gravar na base compartilhada: {exc}")
                return

            st.cache_data.clear()
            destino_base = "Google Sheets" if google_sheets_configurado() else DATA_FILE.name
            st.success(f"{len(importado)} linhas importadas e salvas em {destino_base}.")
            st.rerun()


def renderizar_painel(df: pd.DataFrame) -> None:
    st.subheader("Painel")

    total_kg = df["quantidade"].sum() if not df.empty else 0
    total_entregas = len(df)
    total_produtores = df["produtor"].nunique() if not df.empty else 0
    total_destinos = df["destino"].nunique() if not df.empty else 0

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Quantidade total", f"{total_kg:,.2f} kg".replace(",", "X").replace(".", ",").replace("X", "."))
    col2.metric("Entregas", total_entregas)
    col3.metric("Produtores", total_produtores)
    col4.metric("Destinos", total_destinos)

    if df.empty:
        st.info("Nenhum lancamento encontrado para os filtros atuais.")
        return

    col5, col6 = st.columns(2)
    with col5:
        st.markdown("**Quantidade por produtor**")
        por_produtor = resumo_tabela(df, "produtor")
        st.bar_chart(por_produtor.set_index("produtor")["quantidade"])
        st.dataframe(por_produtor, width="stretch", hide_index=True)

    with col6:
        st.markdown("**Quantidade por destino**")
        por_destino = resumo_tabela(df, "destino")
        st.bar_chart(por_destino.set_index("destino")["quantidade"])
        st.dataframe(por_destino, width="stretch", hide_index=True)

    col7, col8 = st.columns(2)
    with col7:
        st.markdown("**Quantidade por fruta**")
        por_fruta = resumo_tabela(df, "fruta")
        st.bar_chart(por_fruta.set_index("fruta")["quantidade"])
    with col8:
        st.markdown("**Evolucao diaria**")
        diario = df.groupby("data", as_index=False).agg(quantidade=("quantidade", "sum"))
        st.line_chart(diario.set_index("data")["quantidade"])


def renderizar_historico(df: pd.DataFrame) -> None:
    st.subheader("Historico de lancamentos")
    st.dataframe(df.sort_values("data", ascending=False), width="stretch", hide_index=True)
    st.download_button(
        "Baixar historico filtrado em CSV",
        data=df.to_csv(index=False).encode("utf-8-sig"),
        file_name="historico_entregas_filtrado.csv",
        mime="text/csv",
    )


def main() -> None:
    st.title("Controle de Entregas de Produtores Rurais")
    st.caption("Lance produtor, fruta, quantidade e destino. O historico fica salvo em Excel.")

    base = carregar_base()

    with st.sidebar:
        st.header("Base")
        if google_sheets_configurado():
            st.success("Base compartilhada: Google Sheets")
        else:
            st.info(f"Base local: `{DATA_FILE.name}`")
        st.write(f"Registros salvos: **{len(base)}**")

    filtrado = aplicar_filtros(base)

    aba_lancamento, aba_painel, aba_historico = st.tabs(["Lancamento", "Painel", "Historico"])
    with aba_lancamento:
        renderizar_lancamento(base)
        renderizar_importacao(base)
    with aba_painel:
        renderizar_painel(filtrado)
    with aba_historico:
        renderizar_historico(filtrado)


if __name__ == "__main__":
    main()
