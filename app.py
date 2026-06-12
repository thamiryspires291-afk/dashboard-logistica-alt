from flask import Flask, render_template, request, redirect
import pandas as pd
import plotly.express as px
from openpyxl import load_workbook

app = Flask(__name__)
ARQUIVO = "logistica.xlsx"

COLUNAS = [
    "Data", "Placa_do_veiculo", "Em_Rota", "Comandada",
    "Coletada", "Total", "Ajudante_ALT", "Ajudante_SK", "Faltas"
]

def preparar_dataframe(df):
    df.columns = df.columns.astype(str).str.strip()
    for coluna in COLUNAS:
        if coluna not in df.columns:
            df[coluna] = ""
    df["Data"] = pd.to_datetime(df["Data"], errors="coerce")
    return df

def carregar_lancamento():
    df = pd.read_excel(ARQUIVO, sheet_name="LANCAMENTO")
    return preparar_dataframe(df)

def carregar_base():
    df = pd.read_excel(ARQUIVO, sheet_name="BASE_DADOS")
    return preparar_dataframe(df)

@app.route("/")
def dashboard():
    df = carregar_lancamento()

    data = request.args.get("data")

    if data:
        data = pd.to_datetime(data)
        df = df[df["Data"].dt.date == data.date()]

    df["Em_Rota"] = df["Em_Rota"].fillna("").astype(str).str.strip().str.upper()
    df_rota = df[df["Em_Rota"].isin(["SIM", "S", "TRUE", "1"])]

    veiculos_rota = int(
        df_rota["Placa_do_veiculo"]
        .fillna("")
        .astype(str)
        .str.strip()
        .replace("", pd.NA)
        .dropna()
        .nunique()
    )

    total_comandada = pd.to_numeric(df_rota["Comandada"], errors="coerce").fillna(0).sum()
    total_coletada = pd.to_numeric(df_rota["Coletada"], errors="coerce").fillna(0).sum()
    total_geral = pd.to_numeric(df_rota["Total"], errors="coerce").fillna(0).sum()

    ajudante_alt = df["Ajudante_ALT"].fillna("").astype(str).str.strip().str.upper().eq("SIM").sum()
    ajudante_sk = df["Ajudante_SK"].fillna("").astype(str).str.strip().str.upper().eq("SIM").sum()

    faltas = pd.to_numeric(df["Faltas"], errors="coerce").fillna(0).sum()

    grafico = px.bar(
        df_rota,
        x="Placa_do_veiculo",
        y="Total",
        title="Total por Veículo",
        color="Total",
        color_continuous_scale="Greens"
    )

    grafico.update_layout(
        height=450,
        xaxis_title="Placa do Veículo",
        yaxis_title="Total",
        xaxis_tickangle=-45,
        margin=dict(l=40, r=20, t=60, b=120),
        plot_bgcolor="white",
        paper_bgcolor="white"
    )
total_sacflow = adm["Sacflow"].sum()

total_atendimento = adm["Atendimento"].sum()
    return render_template(
        total_sacflow=total_sacflow,

total_atendimento=total_atendimento,
        "index.html",
        veiculos_rota=veiculos_rota,
        total_comandada=total_comandada,
        total_coletada=total_coletada,
        total_geral=total_geral,
        ajudante_alt=ajudante_alt,
        ajudante_sk=ajudante_sk,
        faltas=faltas,
        data_filtro=data.strftime("%Y-%m-%d") if data else "",
        grafico1=grafico.to_html(full_html=False)
    )

@app.route("/dados")
def dados():
    df = carregar_base()

    data = request.args.get("data")

    if data:
        data = pd.to_datetime(data)
        df = df[df["Data"].dt.date == data.date()]

    tabela = df.fillna("").to_dict(orient="records")

    return render_template(
        "dados.html",
        tabela=tabela,
        data_filtro=data.strftime("%Y-%m-%d") if data else ""
    )

@app.route("/performance")
def performance():
    data = request.args.get("data")

    adm = pd.read_excel(ARQUIVO, sheet_name="ADM_ATENDIMENTO")
    mobile = pd.read_excel(ARQUIVO, sheet_name="BAIXA_MOBILE")
    piloto = pd.read_excel(ARQUIVO, sheet_name="BAIXA_PILOTO")

    for df_temp in [adm, mobile, piloto]:
        df_temp.columns = df_temp.columns.astype(str).str.strip()
        df_temp["Data"] = pd.to_datetime(df_temp["Data"], errors="coerce")

    if data:
        data_ref = pd.to_datetime(data)
        adm = adm[adm["Data"].dt.date == data_ref.date()]
        mobile = mobile[mobile["Data"].dt.date == data_ref.date()]
        piloto = piloto[piloto["Data"].dt.date == data_ref.date()]

    media_adm = pd.to_numeric(adm["Percentual_Atendimento"], errors="coerce").fillna(0).mean()

    total_ligacoes = pd.to_numeric(adm["Ligacoes"], errors="coerce").fillna(0).sum()
    total_atendidas = pd.to_numeric(adm["Atendidas"], errors="coerce").fillna(0).sum()
    total_nao_atendidas = pd.to_numeric(adm["Nao_Atendidas"], errors="coerce").fillna(0).sum()
    total_mobile = pd.to_numeric(mobile["Total"], errors="coerce").fillna(0).sum()
    total_piloto = pd.to_numeric(piloto["Total"], errors="coerce").fillna(0).sum()

    ranking_adm = []

    for _, linha in adm.iterrows():
        nota = pd.to_numeric(linha.get("Percentual_Atendimento", 0), errors="coerce")

        if pd.isna(nota):
            nota = 0

        estrelas = int(round(nota))

        ranking_adm.append({
            "nome": linha.get("Atendente", ""),
            "atendidas": linha.get("Atendidas", 0),
            "nota": round(float(nota), 1),
            "estrelas": estrelas,
            "barra": round((float(nota) / 5) * 100, 1)
        })
    ranking_atendentes = (
        adm.sort_values("Percentual_Atendimento", ascending=False)
        [["Atendente", "Percentual_Atendimento"]]
        .head(10)
        .to_dict("records")
    )

    ranking_placas = (
        mobile.groupby("Placa_do_veiculo")["Total"]
        .sum()
        .reset_index()
        .sort_values("Total", ascending=False)
        .head(10)
        .to_dict("records")
    )
    grafico_mobile = px.bar(
        mobile,
        x="Placa_do_veiculo",
        y="Total",
        title="Baixa Mobile por Veículo",
        color_discrete_sequence=["#006b32"]
    )

    grafico_piloto = px.bar(
        piloto,
        x="Atendente",
        y="Total",
        title="Baixa Piloto por Atendente",
        color_discrete_sequence=["#006b32"]
    )

    for grafico_item in [grafico_mobile, grafico_piloto]:
        grafico_item.update_layout(
            height=350,
            plot_bgcolor="white",
            paper_bgcolor="white",
            font_color="#222",
            margin=dict(l=40, r=20, t=60, b=80)
        )

    return render_template(
        "performance.html",
        media_adm=round(media_adm, 1),
        total_ligacoes=total_ligacoes,
        total_atendidas=total_atendidas,
        total_nao_atendidas=total_nao_atendidas,
        total_mobile=total_mobile,
        total_piloto=total_piloto,
        ranking_adm=ranking_adm,
        ranking_atendentes=ranking_atendentes,
        ranking_placas=ranking_placas,
        adm=adm.fillna("").to_dict(orient="records"),
        mobile=mobile.fillna("").to_dict(orient="records"),
        piloto=piloto.fillna("").to_dict(orient="records"),
        grafico_mobile=grafico_mobile.to_html(full_html=False),
        grafico_piloto=grafico_piloto.to_html(full_html=False),
        data_filtro=data if data else ""
    )

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
