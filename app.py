import streamlit as st
import pandas as pd
import plotly.express as px
import unicodedata
import altair as alt

def normalizar_colunas(df):
    novas_colunas = []

    for col in df.columns:
        col = str(col)

        # remove acentos
        col = unicodedata.normalize("NFKD", col).encode("ASCII", "ignore").decode("ASCII")

        # remove espaços extras
        col = col.strip()

        novas_colunas.append(col)

    df.columns = novas_colunas
    return df

# -------------------------------------------
# CONFIGURAÇÃO DA PÁGINA
# -------------------------------------------
st.set_page_config(page_title="Dashboard Evelog", layout="wide")
st.sidebar.image("logo.svg", width=180)
st.title("Dashboard Evelog")

# -------------------------------------------
# SIDEBAR
# -------------------------------------------
st.sidebar.header("Importar Planilhas")
uploaded_files = st.sidebar.file_uploader(
    "Selecione uma ou mais planilhas Excel",
    type=["xlsx", "xls"],
    accept_multiple_files=True
)

@st.cache_data
def exportar_excel(df):
    from io import BytesIO

    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Base_Unificada")

    return buffer.getvalue()

# -------------------------------------------
# FUNÇÃO OTIMIZADA (CACHE)
# -------------------------------------------
@st.cache_data(show_spinner="📥 Processando planilha...")
def carregar_planilha(file):

    # Leitura da planilha sem cabeçalho
    df_raw = pd.read_excel(file, header=None)

    # Detecta linha de cabeçalho
    linha_cabecalho = df_raw.notna().sum(axis=1).idxmax()

    # Cria dataframe com cabeçalho correto
    df = pd.read_excel(file, skiprows=linha_cabecalho)

    df = normalizar_colunas(df)

    # -----------------------------------
    # LOCALIZA COLUNAS PELO NOME
    # -----------------------------------
    col_emissao = df.columns.get_loc("Dt Emissao")
    col_cliente = df.columns.get_loc("Cliente")

    # -----------------------------------
    # PEGA OS VALORES NA PLANILHA ORIGINAL
    # -----------------------------------
    try:
        data_ref = pd.to_datetime(df_raw.iloc[2, col_emissao], dayfirst=True)
        mes_ref = data_ref.strftime("%m/%Y")
    except:
        mes_ref = "Mês desconhecido"

    try:
        cliente = str(df_raw.iloc[2, col_cliente]).strip()
    except:
        cliente = "Cliente desconhecido"

    # -----------------------------------
    # ADICIONA NA BASE
    # -----------------------------------
    df["Arquivo_Origem"] = mes_ref
    df["Cliente"] = cliente

    from datetime import datetime

    # -----------------------------------
    # CALCULO PRAZO
    # -----------------------------------

    status_excluidos = [
        "DEVOLVIDO",
        "EM DEVOLUCAO",
        "ENCERRADO",
        "LISTA DEVOLUCAO",
        "SINISTRO",
        "TRAVADO",
        "UNITIZADO"
    ]

    descricoes_excluidas = [
        "AVARIA / DANO PARCIAL",
        "AVARIA / DANO TOTAL",
        "ERRO DO EMISSOR",
        "EXTRAVIO PARCIAL",
        "EXTRAVIO TOTAL",
        "FALTA DE ACAREAÇÃO / TOTAL",
        "FALTA DE COMPROVANTE DE ENTREGA / TOTAL",
        "INDENIZACAO RECUSADA",
        "INDICIO VIOLACAO",
        "LIBERADO",
        "PERDA POR PRAZO SLA / TOTAL"
    ]

    col_previsao = next((c for c in df.columns if "previsao" in c.lower()), None)
    col_evento = next((c for c in df.columns if "evento" in c.lower()), None)
    col_status = next((c for c in df.columns if "status" in c.lower()), None)
    col_descricao = next((c for c in df.columns if "descricao" in c.lower()), None)

    if col_previsao and col_status:

        previsao = pd.to_datetime(df[col_previsao], dayfirst=True, errors="coerce").dt.normalize()
        status = df[col_status].astype(str).str.strip().str.upper()

        descricao = None
        if col_descricao:
            descricao = df[col_descricao].astype(str).str.strip().str.upper()

        hoje = pd.Timestamp.today().normalize()

        df["Prazo"] = ""

        # -----------------------------------
        # MASCARA DE EXCLUSÃO
        # -----------------------------------

        mask_excluir = status.isin(status_excluidos)

        if descricao is not None:
            mask_excluir = mask_excluir | (
                (status == "CUSTODIA") &
                (descricao.isin(descricoes_excluidas))
            )

        # -----------------------------------
        # ENTREGUES
        # -----------------------------------

        if col_evento:

            evento = pd.to_datetime(df[col_evento], dayfirst=True, errors="coerce").dt.normalize()

            mask_entregue = (status == "ENTREGUE") & (~mask_excluir)

            df.loc[mask_entregue & (evento <= previsao), "Prazo"] = "NO PRAZO"
            df.loc[mask_entregue & (evento > previsao), "Prazo"] = "FORA DO PRAZO"

        # -----------------------------------
        # PEDIDOS EM ABERTO
        # -----------------------------------

        mask_abertos = (
            (status != "ENTREGUE") &
            (~mask_excluir) &
            (previsao.notna())
        )

        diff = (previsao - hoje).dt.days

        df.loc[mask_abertos & (diff >= 0), "Prazo"] = "FALTAM " + diff.astype(str).str.zfill(3) + " DIAS"
        df.loc[mask_abertos & (diff < 0), "Prazo"] = "ATRASADO " + diff.abs().astype(str).str.zfill(3) + " DIAS"

        df["Prazo"] = df["Prazo"].replace("", None)

        # -----------------------------------
        # PADRONIZAÇÃO DE OCORRÊNCIAS
        # -----------------------------------

        col_descricao = next((c for c in df.columns if "descricao" in c.lower()), None)

        if col_descricao:

            descricao = df[col_descricao].fillna("").astype(str).str.strip().str.upper()

            mapa_ocorrencias = {

                # Destinatário ausente
                "AUSENTE": "DEST. AUSENTE",
                "AUSENTE 2": "DEST. AUSENTE",
                "AUSENTE 3": "DEST. AUSENTE",
                "FECHADO": "DEST. AUSENTE",
                "FECHADO 2": "DEST. AUSENTE",
                "TEMPO DE ESPERA EXCEDIDO NO DESTINATARIO": "DEST. AUSENTE",
                "AUSENTE EM FERIAS": "DEST. AUSENTE",

                # Pedido avariado
                "AVARIA / DANO PARCIAL": "PEDIDO AVARIADO",
                "AVARIA / DANO TOTAL": "PEDIDO AVARIADO",
                "INDICIO VIOLACAO": "PEDIDO AVARIADO",

                # Problema de endereço
                "DESTINATARIO DESCONHECIDO": "PROB. ENDEREÇO",
                "ENDERECO INSUFICIENTE": "PROB. ENDEREÇO",
                "ENDERECO NAO LOCALIZADO": "PROB. ENDEREÇO",
                "MUDOU-SE": "PROB. ENDEREÇO",
                "NUMERO NAO LOCALIZADO": "PROB. ENDEREÇO",

                # Agência
                "DESTINATARIO SOLICITOU RETIRAR NA UNIDADE": "AG. RETIRADA AGÊNCIA",

                # Last Mile
                "ATRASO TRANSPORTE": "LAST MILE",
                "FALHA ENTREGA": "LAST MILE",
                "SOLICITACAO ENTREGA FUTURA": "LAST MILE",

                # Fiscal
                "EM ANÁLISE NO POSTO FISCAL": "PROB. FISCAL",
                "RETENCAO FISCAL DE DOCUMENTO E/OU MERCADORIA": "PROB. FISCAL",
                "SAIDA FISCALIZACAO": "PROB. FISCAL",

                # Zona rural
                "ENDERECO EM ZONA RURAL": "ZONA RURAL",

                # Middle Mile
                "BUSCA": "MIDDLE MILE",
                "NAO ENTROU NA UNIDADE": "MIDDLE MILE",

                # Feriado
                "FECHADO EM VESPERA OU APOS FERIADO": "FERIADO",

                # Layout
                #"TERMO DE IRREGULARIDADE": "ERRO LAYOUT ETIQUETA",

                # Rodovia
                "TRAFEGO INTERROMPIDO": "RODOVIA INTERDITADA",

                # Clima
                "TEMPORAL": "PROB. CLIMÁTICO",

                # Triagem
                "ERRO DE TRIAGEM / SEPARACAO": "ERRO DE TRIAGEM",

                # Acareação
                "FALTA DE ACAREAÇÃO / TOTAL": "ACAREAÇÃO",
                "FALTA DE COMPROVANTE DE ENTREGA / TOTAL": "ACAREAÇÃO",

                # Área de risco
                "RESTRICAO DE ACESSO / MOVIMENTACAO": "ÁREA DE RISCO",

                # Sinistro
                "SINISTRO / ACIDENTE TRANSPORTE": "SINISTRO",
                "FURTO / ROUBO": "SINISTRO",

                # Tentativa de furto
                "PARADO B.O POLICIAL": "TENTATIVA DE FURTO",

                # Pedido recusado
                "RECUSADO": "PEDIDO RECUSADO",
                "RECUSADO - DIVERGENCIA DE PEDIDO": "PEDIDO RECUSADO",
                "RECUSADO - NAO PAGA FRETE": "PEDIDO RECUSADO",

                # Extravio
                "EXTRAVIO PARCIAL": "EXTRAVIO",
                "EXTRAVIO TOTAL": "EXTRAVIO",
                "PERDA POR PRAZO SLA / TOTAL": "EXTRAVIO",

                # Devolucao
                "DEVOLUCAO POR INSTRUCAO MATRIZ": "DEVOLUCAO",
                "DEVOLUCAO POR INSTRUCAO REMETENTE": "DEVOLUCAO",
                "DEVOLUCAO RECUSADA": "DEVOLUCAO",

                "": "SEM OCORRENCIA",

            }

            df["Ocorrencias"] = descricao.map(mapa_ocorrencias).fillna(descricao)

    return df

# -------------------------------------------
# CARREGAMENTO DAS PLANILHAS
# -------------------------------------------
base_unificada = None

if uploaded_files:

    dfs = []

    with st.spinner("🔄 Unificando bases..."):
        for file in uploaded_files:
            df = carregar_planilha(file)
            dfs.append(df)

            mes_ref = df["Arquivo_Origem"].iloc[0]
            cliente = df["Cliente"].iloc[0]

            st.sidebar.write(
                f"📄 **{file.name}** carregado → {mes_ref} | 🏷️ {cliente}"
            )

    base_unificada = pd.concat(dfs, ignore_index=True)
    st.sidebar.success(f"{len(uploaded_files)} planilha(s) carregada(s)")

st.sidebar.markdown("---")

mapa_regiao = {
    # Norte
    "AC": "Norte", "AP": "Norte", "AM": "Norte", "PA": "Norte",
    "RO": "Norte", "RR": "Norte", "TO": "Norte",

    # Nordeste
    "AL": "Nordeste", "BA": "Nordeste", "CE": "Nordeste", "MA": "Nordeste",
    "PB": "Nordeste", "PE": "Nordeste", "PI": "Nordeste",
    "RN": "Nordeste", "SE": "Nordeste",

    # Centro-Oeste
    "DF": "Centro-Oeste", "GO": "Centro-Oeste",
    "MT": "Centro-Oeste", "MS": "Centro-Oeste",

    # Sudeste
    "ES": "Sudeste", "MG": "Sudeste", "RJ": "Sudeste", "SP": "Sudeste",

    # Sul
    "PR": "Sul", "RS": "Sul", "SC": "Sul"
}

with st.sidebar:

    st.markdown("### 📦 Base unificada")

    if base_unificada is not None and not base_unificada.empty:
        st.metric(
            label="Total de pedidos",
            value=len(base_unificada)
        )
    else:
        st.metric(
            label="Total de pedidos",
            value="—"
        )

if base_unificada is not None and not base_unificada.empty:
    excel_bytes = exportar_excel(base_unificada)

    st.sidebar.download_button(
        label="⬇️ Exportar base (.xlsx)",
        data=excel_bytes,
        file_name="base_unificada.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.sidebar.download_button(
        label="⬇️ Exportar base (.xlsx)",
        data=b"",
        disabled=True
    )

st.sidebar.markdown("---")

# -------------------------------------------
# FILTRO GLOBAL - DATA DE EMISSÃO
# -------------------------------------------

if base_unificada is not None and not base_unificada.empty:

    base_unificada["Dt Emissao"] = base_unificada["Dt Emissao"].astype(str).str.strip()

    base_unificada["Dt Emissao"] = pd.to_datetime(
        base_unificada["Dt Emissao"],
        dayfirst=True,   # 👈 ESSENCIAL pro padrão BR
        errors="coerce"
    )

    df_datas_validas = base_unificada.dropna(subset=["Dt Emissao"])

    min_emissao = df_datas_validas["Dt Emissao"].min().date()
    max_emissao = df_datas_validas["Dt Emissao"].max().date()

    data_emissao = st.sidebar.date_input(
        "📅 Período de emissão",
        value=(min_emissao, max_emissao),
        min_value=min_emissao,
        max_value=max_emissao,
        key="filtro_global_emissao"
    )

    if isinstance(data_emissao, tuple) and len(data_emissao) == 2:
        data_ini, data_fim = data_emissao

        base_unificada = base_unificada[
            (base_unificada["Dt Emissao"].notna()) &
            (base_unificada["Dt Emissao"].dt.date >= data_ini) &
            (base_unificada["Dt Emissao"].dt.date <= data_fim)
        ]
    
    qtd_pedidos = len(base_unificada)

    st.sidebar.metric(
        label="Total de pedidos filtrados",
        value=f"{qtd_pedidos:,}".replace(",", ".")
    )

    st.sidebar.markdown("---")

# -------------------------------------------
# DASHBOARD
# -------------------------------------------
if base_unificada is not None:

    df = base_unificada.copy()
    
    tipo_pedido = st.segmented_control(
        "Tipo de pedidos",
        ["Pedidos Em aberto (Vencidos, A vencer)", "Pedidos Entregues (Performance OTD)", "Pedidos não concluídos (devoluções, sinistros, etc.)"]
    )

    if tipo_pedido == "Pedidos Em aberto (Vencidos, A vencer)":

        # -----------------------------------
        # BASE INICIAL
        # -----------------------------------

        df_abertos = df[
            (~df["Prazo"].isin(["NO PRAZO", "FORA DO PRAZO"])) &
            (df["Prazo"].notna()) &
            (df["Prazo"].astype(str).str.strip() != "")
        ].copy()

        if df_abertos.empty:

            st.info("Não há pedidos em aberto na base.")

        else:

            st.subheader("Pedidos em aberto")

            with st.sidebar:

                st.markdown("### 📦 Base de pedidos em aberto")

                if df_abertos is not None and not df_abertos.empty:
                    st.metric(
                        label="Total de pedidos",
                        value=len(df_abertos)
                    )
                else:
                    st.metric(
                        label="Total de pedidos",
                        value="—"
                    )

            if df_abertos is not None and not df_abertos.empty:
                excel_bytes = exportar_excel(df_abertos)

                st.sidebar.download_button(
                    label="⬇️ Exportar base (.xlsx)",
                    data=excel_bytes,
                    file_name="df_abertos.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.sidebar.download_button(
                    label="⬇️ Exportar base (.xlsx)",
                    data=b"",
                    disabled=True
                )

            # -----------------------------------
            # TOPO (RADIO + MÉTRICA)
            # -----------------------------------

            col1, col2 = st.columns([3,1])

            with col1:
                tipo = st.radio(
                    "Visualização",
                    ["Vencidos", "A vencer"],
                    horizontal=True
                )

            with col2:
                if tipo == "Vencidos":
                    total = df_abertos["Prazo"].str.contains("ATRASADO", na=False).sum()
                    st.metric("Total de pedidos vencidos", total)
                else:
                    total = df_abertos["Prazo"].str.contains("FALTAM", na=False).sum()
                    st.metric("Total de pedidos a vencer", total)

            # -----------------------------------
            # BASE ORIGINAL
            # -----------------------------------

            if tipo == "Vencidos":
                df_base_original = df_abertos[
                    df_abertos["Prazo"].str.contains("ATRASADO", na=False)
                ].copy()
            else:
                df_base_original = df_abertos[
                    df_abertos["Prazo"].str.contains("FALTAM", na=False)
                ].copy()

            df_base_original["Dias"] = (
                df_base_original["Prazo"]
                .str.extract(r'(\d+)', expand=False)
                .astype(float)
            )

            # -----------------------------------
            # NOVO CAMPO - DIAS SEM MOVIMENTAÇÃO
            # -----------------------------------

            hoje = pd.Timestamp.now().normalize()

            # 🔥 PARSE CORRETO (2 formatos)
            def parse_data_br(col):
                return pd.to_datetime(col, format="%d/%m/%Y %H:%M:%S", errors="coerce") \
                    .fillna(pd.to_datetime(col, format="%d/%m/%Y", errors="coerce"))

            df_base_original["Dt Evento"] = (
                df_base_original["Dt Evento"]
                .astype(str)
                .str.strip()
            )

            df_base_original["Dt Evento"] = parse_data_br(df_base_original["Dt Evento"])

            # cálculo correto
            df_base_original["Dias_sem_mov"] = (
                (hoje - df_base_original["Dt Evento"].dt.normalize())
                .dt.days
            )

            # 🔥 evita negativos (caso ainda exista dado estranho)
            df_base_original["Dias_sem_mov"] = df_base_original["Dias_sem_mov"].clip(lower=0)

            # -----------------------------------
            # SESSION STATE (keys fixas)
            # -----------------------------------

            if "dias" not in st.session_state:
                st.session_state.dias = []

            if "status" not in st.session_state:
                st.session_state.status = []

            if "ocorrencias" not in st.session_state:
                st.session_state.ocorrencias = []

            if "dias_sem_mov" not in st.session_state:
                st.session_state.dias_sem_mov = []

            # -----------------------------------
            # FUNÇÃO FILTRO
            # -----------------------------------

            def aplicar_filtros(df, dias, status, ocorrencias, dias_sem_mov):
                if dias:
                    df = df[df["Dias"].isin(dias)]
                if status:
                    df = df[df["Status"].isin(status)]
                if ocorrencias:
                    df = df[df["Ocorrencias"].isin(ocorrencias)]
                if dias_sem_mov:
                    df = df[df["Dias_sem_mov"].isin(dias_sem_mov)]
                return df

            # -----------------------------------
            # FILTROS
            # -----------------------------------

            col1, col2, col3, col4 = st.columns([3,3,3,3])

            # ---------------- DIAS ----------------
            with col1:

                df_temp = aplicar_filtros(
                    df_base_original,
                    [],
                    st.session_state.status,
                    st.session_state.ocorrencias,
                    st.session_state.dias_sem_mov
                )

                dias_valores = sorted(df_temp["Dias"].dropna().unique())

                # -------------------------------
                # MAPA VALOR -> LABEL
                # -------------------------------
                mapa_valor_label = {}

                for d in dias_valores:
                    d_int = int(d)

                    if tipo == "Vencidos":
                        label = f"{d_int} dia{'s' if d_int > 1 else ''} vencido"
                    else:
                        label = f"{d_int} dia{'s' if d_int > 1 else ''} para vencer"

                    mapa_valor_label[d] = label

                mapa_label_valor = {v: k for k, v in mapa_valor_label.items()}

                opcoes_labels = list(mapa_valor_label.values())

                # -------------------------------
                # LIMPAR INVÁLIDOS (BASE NUMÉRICA)
                # -------------------------------
                if "dias" not in st.session_state:
                    st.session_state.dias = []

                st.session_state.dias = [
                    d for d in st.session_state.dias if d in dias_valores
                ]

                # -------------------------------
                # CONVERTER DEFAULT (número -> label)
                # -------------------------------
                default_labels = [
                    mapa_valor_label[d]
                    for d in st.session_state.dias
                    if d in mapa_valor_label
                ]

                # -------------------------------
                # MULTISELECT (COM KEY!)
                # -------------------------------
                selecionados_labels = st.multiselect(
                    "Dias",
                    opcoes_labels,
                    default=default_labels,
                    key="dias_widget"
                )

                # -------------------------------
                # ATUALIZA ESTADO REAL (NÚMERO)
                # -------------------------------
                st.session_state.dias = [
                    mapa_label_valor[label]
                    for label in selecionados_labels
                ]

            # ---------------- STATUS ----------------
            with col2:

                df_temp = aplicar_filtros(
                    df_base_original,
                    st.session_state.dias,
                    [],
                    st.session_state.ocorrencias,
                    st.session_state.dias_sem_mov
                )

                status_opcoes = sorted(
                    df_temp["Status"]
                    .dropna()
                    .astype(str)
                    .str.strip()
                    .unique()
                )

                st.session_state.status = [
                    s for s in st.session_state.status if s in status_opcoes
                ]

                st.multiselect(
                    "Status",
                    status_opcoes,
                    key="status"
                )

            # ---------------- OCORRÊNCIAS ----------------
            with col3:

                df_temp = aplicar_filtros(
                    df_base_original,
                    st.session_state.dias,
                    st.session_state.status,
                    [],
                    st.session_state.dias_sem_mov 
                )

                ocorrencias_opcoes = sorted(
                    df_temp["Ocorrencias"]
                    .dropna()
                    .astype(str)
                    .str.strip()
                    .loc[lambda x: x != ""]
                    .unique()
                )

                st.session_state.ocorrencias = [
                    o for o in st.session_state.ocorrencias if o in ocorrencias_opcoes
                ]

                st.multiselect(
                    "Ocorrências",
                    ocorrencias_opcoes,
                    key="ocorrencias"
                )
            
            with col4:

                df_temp = aplicar_filtros(
                    df_base_original,
                    st.session_state.dias,
                    st.session_state.status,
                    st.session_state.ocorrencias,
                    []
                )

                dias_sem_mov_opcoes = sorted(
                    df_temp["Dias_sem_mov"]
                    .dropna()
                    .astype(int)
                    .unique()
                )

                st.session_state.dias_sem_mov = [
                    d for d in st.session_state.dias_sem_mov if d in dias_sem_mov_opcoes
                ]

                st.multiselect(
                    "Dias sem movimentação",
                    dias_sem_mov_opcoes,
                    key="dias_sem_mov"
                )

            # ---------------- LIMPAR ----------------
            def limpar_filtros():
                st.session_state.dias = []
                st.session_state.status = []
                st.session_state.ocorrencias = []
                st.session_state.dias_sem_mov = []
            st.button("Limpar filtros", on_click=limpar_filtros)

            # -----------------------------------
            # BASE FINAL
            # -----------------------------------

            df_base = aplicar_filtros(
                df_base_original,
                st.session_state.dias,
                st.session_state.status,
                st.session_state.ocorrencias,
                st.session_state.dias_sem_mov
            )

            # -----------------------------------
            # GRÁFICOS
            # -----------------------------------

            grafico = (
                df_base.groupby("Dias")
                .size()
                .reset_index(name="Quantidade")
                .sort_values("Dias")
            )

            if not grafico.empty:

                # DIAS
                if tipo == "Vencidos":

                    tituloDias = "Distribuição por dias (Pedidos vencidos)"
                    tituloStatus = "Status (Pedidos vencidos)"
                    tituloOcorrencias = "Ocorrencias (Pedidos vencidos)"
                    tituloDiasSemMov = "Dias sem movimentação (Pedidos vencidos)"
                    tituloPedidosFiltrados = "Pedidos filtrados (Pedidos vencidos)"

                    eixo_x_titulo = "Dias vencidos"
                else:
                    tituloDias = "Distribuição por dias (Pedidos a vencer)"
                    tituloStatus = "Status (Pedidos a vencer)"
                    tituloOcorrencias = "Ocorrencias (Pedidos a vencer)"
                    tituloDiasSemMov = "Dias sem movimentação (Pedidos a vencer)"
                    tituloPedidosFiltrados = "Pedidos filtrados (Pedidos a vencer)"

                    eixo_x_titulo = "Dias a vencer"

                st.subheader(tituloDias)

                base_chart = alt.Chart(grafico).encode(
                    x=alt.X(
                        "Dias:O",
                        title=eixo_x_titulo,
                        axis=alt.Axis(labelAngle=0)
                    ),
                    y=alt.Y(
                        "Quantidade:Q",
                        axis=alt.Axis(tickMinStep=1)
                    )
                )

                chart = base_chart.mark_bar() + base_chart.mark_text(
                    dy=-10,
                    color="white"
                ).encode(text="Quantidade:Q")

                st.altair_chart(chart, use_container_width=True)

            # STATUS
            status_df = (
                df_base.groupby("Status")
                .size()
                .reset_index(name="Quantidade")
                .sort_values("Quantidade", ascending=False)
            )

            if not status_df.empty:

                st.subheader(tituloStatus)

                base_st = alt.Chart(status_df).encode(
                    y=alt.Y("Status:N", sort="-x"),
                    x=alt.X(
                        "Quantidade:Q",
                        axis=alt.Axis(tickMinStep=1)
                    )
                )

                chart_st = base_st.mark_bar() + base_st.mark_text(
                    align="left",
                    dx=5,
                    color="white"
                ).encode(text="Quantidade:Q")

                st.altair_chart(chart_st, use_container_width=True)

            # OCORRÊNCIAS
            df_oc = df_base[
                df_base["Ocorrencias"].notna() &
                (df_base["Ocorrencias"].str.strip() != "")
            ]

            ocorrencias = (
                df_oc.groupby("Ocorrencias")
                .size()
                .reset_index(name="Quantidade")
                .sort_values("Quantidade", ascending=False)
            )

            if not ocorrencias.empty:

                st.subheader(tituloOcorrencias)

                base_oc = alt.Chart(ocorrencias).encode(
                    y=alt.Y("Ocorrencias:N", sort="-x"),
                    x=alt.X(
                        "Quantidade:Q",
                        axis=alt.Axis(tickMinStep=1)
                    )
                )

                chart_oc = base_oc.mark_bar() + base_oc.mark_text(
                    align="left",
                    dx=5,
                    color="white"
                ).encode(text="Quantidade:Q")

                st.altair_chart(chart_oc, use_container_width=True)

            # -----------------------------------
            # GRÁFICO - DIAS SEM MOVIMENTAÇÃO
            # -----------------------------------

            grafico_sm = (
                df_base.groupby("Dias_sem_mov")
                .size()
                .reset_index(name="Quantidade")
                .sort_values("Dias_sem_mov")
            )

            if not grafico_sm.empty:

                st.subheader(tituloDiasSemMov)

                base_sm = alt.Chart(grafico_sm).encode(
                    x=alt.X(
                        "Dias_sem_mov:O",
                        title="Dias sem movimentação",
                        axis=alt.Axis(labelAngle=0)
                    ),
                    y=alt.Y(
                        "Quantidade:Q",
                        title="Quantidade de pedidos",
                        axis=alt.Axis(tickMinStep=1)
                    )
                )

                chart_sm = base_sm.mark_bar() + base_sm.mark_text(
                    dy=-10,
                    color="white"
                ).encode(
                    text="Quantidade:Q"
                )

                st.altair_chart(chart_sm, use_container_width=True)

            # -----------------------------------
            # TABELA FINAL
            # -----------------------------------

            st.subheader(tituloPedidosFiltrados)

            st.write(f"Total: {len(df_base)}")

            st.dataframe(
                df_base,
                use_container_width=True,
                hide_index=True
            )

            import io

            # -----------------------------------
            # EXPORTAR PARA EXCEL
            # -----------------------------------

            def gerar_excel(df):
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                    df.to_excel(writer, index=False, sheet_name="Pedidos")
                return output.getvalue()

            excel_data = gerar_excel(df_base)

            st.download_button(
                label="📥 Exportar para Excel",
                data=excel_data,
                file_name="pedidos_filtrados.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    elif tipo_pedido == "Pedidos Entregues (Performance OTD)":

        from datetime import timedelta

        df_perf = base_unificada.copy()

        # -----------------------------------
        # SOMENTE ENTREGUES
        # -----------------------------------

        df_perf = df_perf[
            df_perf["Prazo"].isin(["NO PRAZO", "FORA DO PRAZO"])
        ].copy()

        if df_perf.empty:

            st.info("Não há pedidos entregues na base.")

        else:

            st.subheader("Performance OTD")

            # -----------------------------------
            # TRATAMENTO DE DATA (🔥 CORREÇÃO PRINCIPAL)
            # -----------------------------------

            df_perf["Dt Evento"] = df_perf["Dt Evento"].astype(str).str.strip()

            df_perf["Dt Evento"] = pd.to_datetime(
                df_perf["Dt Evento"],
                format="%d/%m/%Y %H:%M:%S",
                errors="coerce"
            )

            # -----------------------------------
            # RANGE REAL
            # -----------------------------------

            min_data = df_perf["Dt Evento"].min().date()
            max_data = df_perf["Dt Evento"].max().date()

            # -----------------------------------
            # FILTROS
            # -----------------------------------

            col1, col2, col3 = st.columns([1,1,1])

            with col1:
                data_evento = st.date_input(
                    "Período de entrega",
                    value=(min_data, max_data),
                    min_value=min_data,
                    max_value=max_data
                )

            with col2:
                tipo_ordem = st.radio(
                    "Ordenar por",
                    ["Quantidade", "Percentual"],
                    horizontal=True
                )

            with col3:
                tipo_visao = st.radio(
                    "Visualização",
                    ["UF", "Região"],
                    horizontal=True
                )

            # -----------------------------------
            # FILTRO DE DATA (🔥 CORRETO)
            # -----------------------------------

            if isinstance(data_evento, tuple) and len(data_evento) == 2:
                data_ini, data_fim = data_evento

                # inclui o dia inteiro final
                data_fim = data_fim + timedelta(days=1)

                df_perf = df_perf[
                    (df_perf["Dt Evento"].notna()) &
                    (df_perf["Dt Evento"] >= pd.to_datetime(data_ini)) &
                    (df_perf["Dt Evento"] < pd.to_datetime(data_fim))
                ]

            # -----------------------------------
            # AGRUPAMENTO
            # -----------------------------------

            if tipo_visao == "Região":
                df_perf["Grupo"] = df_perf["UF"].map(mapa_regiao)
            else:
                df_perf["Grupo"] = df_perf["UF"]

            # -----------------------------------
            # AGRUPAR
            # -----------------------------------

            df_grouped = (
                df_perf.groupby(["Grupo", "Prazo"])
                .size()
                .unstack(fill_value=0)
                .reset_index()
            )

            # garantir colunas
            for col in ["NO PRAZO", "FORA DO PRAZO"]:
                if col not in df_grouped.columns:
                    df_grouped[col] = 0

            # -----------------------------------
            # TOTAL
            # -----------------------------------

            df_grouped["Total"] = df_grouped["NO PRAZO"] + df_grouped["FORA DO PRAZO"]

            # -----------------------------------
            # %
            # -----------------------------------

            df_grouped["Percentual No Prazo"] = (df_grouped["NO PRAZO"] / df_grouped["Total"]) * 100
            df_grouped["Percentual Atrasado"] = (df_grouped["FORA DO PRAZO"] / df_grouped["Total"]) * 100

            # -----------------------------------
            # ORDENAÇÃO
            # -----------------------------------

            if tipo_ordem == "Percentual":
                df_grouped = df_grouped.sort_values(
                    ["Percentual No Prazo", "Total"],
                    ascending=[True, True]
                )
            else:
                df_grouped = df_grouped.sort_values(
                    "Total",
                    ascending=True
                )

            # -----------------------------------
            # TOTAL GERAL (NOVA BARRA)
            # -----------------------------------

            total_geral = pd.DataFrame({
                "Grupo": ["Total"],
                "NO PRAZO": [df_grouped["NO PRAZO"].sum()],
                "FORA DO PRAZO": [df_grouped["FORA DO PRAZO"].sum()]
            })

            total_geral["Total"] = total_geral["NO PRAZO"] + total_geral["FORA DO PRAZO"]

            total_geral["Percentual No Prazo"] = (
                total_geral["NO PRAZO"] / total_geral["Total"]
            ) * 100

            total_geral["Percentual Atrasado"] = (
                total_geral["FORA DO PRAZO"] / total_geral["Total"]
            ) * 100

            # juntar com o original
            df_grouped = pd.concat([df_grouped, total_geral], ignore_index=True)

            # -----------------------------------
            # GRÁFICO
            # -----------------------------------

            fig = px.bar(
                df_grouped,
                x="Grupo",
                y=["Percentual No Prazo", "Percentual Atrasado"],
                labels={"value": "Percentual (%)", "variable": ""},
                color_discrete_map={
                    "Percentual No Prazo": "#2ca02c",
                    "Percentual Atrasado": "#d62728"
                }
            )

            # -----------------------------------
            # HOVER COM QUANTIDADE
            # -----------------------------------

            fig.data[0].customdata = df_grouped["NO PRAZO"]
            fig.data[1].customdata = df_grouped["FORA DO PRAZO"]

            fig.data[0].hovertemplate = (
                "Grupo: %{x}<br>"
                "No prazo<br>"
                "Percentual: %{y:.1f}%<br>"
                "Qtd: %{customdata}<extra></extra>"
            )

            fig.data[1].hovertemplate = (
                "Grupo: %{x}<br>"
                "Fora do prazo<br>"
                "Percentual: %{y:.1f}%<br>"
                "Qtd: %{customdata}<extra></extra>"
            )

            # -----------------------------------
            # TEXTO DENTRO DAS BARRAS
            # -----------------------------------

            fig.update_traces(
                texttemplate="%{y:.0f}%",
                textposition="inside",
                textangle=0   # 🔥 força horizontal
            )

            # -----------------------------------
            # TOTAL EM CIMA DA BARRA
            # -----------------------------------

            fig.add_scatter(
                x=df_grouped["Grupo"],
                y=[100] * len(df_grouped),
                mode="text",
                text=df_grouped["Total"],
                textposition="top center",
                showlegend=False,
                hoverinfo="skip"
            )

            # -----------------------------------
            # LAYOUT
            # -----------------------------------

            fig.update_layout(
                barmode="stack",
                yaxis=dict(range=[0, 110], ticksuffix="%"),  # 🔥 sobe o teto
                legend_title_text=""
            )

            df_grouped["ordem"] = df_grouped["Grupo"].apply(
                lambda x: 1 if x == "Total" else 0
            )

            df_grouped = df_grouped.sort_values(
                ["ordem", "Percentual No Prazo" if tipo_ordem == "Percentual" else "Total"],
                ascending=True
            ).drop(columns="ordem")

            # -----------------------------------
            # STREAMLIT
            # -----------------------------------

            st.plotly_chart(fig, use_container_width=True)

            # -----------------------------------
            # BASE PARA OTD ORIGINAL
            # -----------------------------------

            total_no_prazo = (df_perf["Prazo"] == "NO PRAZO").sum()
            total_atrasado = (df_perf["Prazo"] == "FORA DO PRAZO").sum()

            df_otd = pd.DataFrame({
                "Status": ["No Prazo", "Fora do Prazo"],
                "Quantidade": [total_no_prazo, total_atrasado]
            })

            # -----------------------------------
            # BASE PARA OTD JUSTIFICADO
            # -----------------------------------

            df_calc = df_perf.copy()

            # padronizar ocorrência
            df_calc["Ocorrencias"] = (
                df_calc["Ocorrencias"]
                .fillna("")
                .astype(str)
                .str.strip()
                .str.upper()
            )

            # regra de justificativa
            cond_justificado = (
                (df_calc["Prazo"] == "FORA DO PRAZO") &
                (
                    df_calc["Ocorrencias"].str.contains(
                        "DEST\\. AUSENTE|PROB\\. ENDEREÇO",
                        regex=True
                    )
                )
            )

            # recalcular "no prazo ajustado"
            df_calc["No Prazo Ajustado"] = (
                (df_calc["Prazo"] == "NO PRAZO") |
                (cond_justificado)
            )

            total_no_prazo_just = df_calc["No Prazo Ajustado"].sum()
            total_atrasado_just = len(df_calc) - total_no_prazo_just

            df_otd_just = pd.DataFrame({
                "Status": ["No Prazo", "Fora do Prazo"],
                "Quantidade": [total_no_prazo_just, total_atrasado_just]
            })

            # -----------------------------------
            # GRÁFICOS
            # -----------------------------------

            col1, col2 = st.columns(2)

            with col1:
                fig1 = px.pie(
                    df_otd,
                    names="Status",
                    values="Quantidade",
                    title="OTD Original",
                    hole=0.4
                )
                st.plotly_chart(fig1, use_container_width=True)

            with col2:
                fig2 = px.pie(
                    df_otd_just,
                    names="Status",
                    values="Quantidade",
                    title="OTD Justificado",
                    hole=0.4
                )
                st.plotly_chart(fig2, use_container_width=True)

            st.subheader("Evolução do OTD")

            df_linha = base_unificada.copy()

            # -----------------------------------
            # SOMENTE ENTREGUES
            # -----------------------------------

            df_linha = df_linha[
                df_linha["Prazo"].isin(["NO PRAZO", "FORA DO PRAZO"])
            ].copy()

            # -----------------------------------
            # TRATAMENTO DATA
            # -----------------------------------

            df_linha["Dt Evento"] = df_linha["Dt Evento"].astype(str).str.strip()

            df_linha["Dt Evento"] = pd.to_datetime(
                df_linha["Dt Evento"],
                format="%d/%m/%Y %H:%M:%S",
                errors="coerce"
            )

            df_linha = df_linha[df_linha["Dt Evento"].notna()]

            # -----------------------------------
            # FILTRO PERÍODO (usa o mesmo do gráfico anterior se quiser)
            # -----------------------------------

            if isinstance(data_evento, tuple) and len(data_evento) == 2:
                data_ini, data_fim = data_evento
                data_fim = data_fim + pd.Timedelta(days=1)

                df_linha = df_linha[
                    (df_linha["Dt Evento"] >= pd.to_datetime(data_ini)) &
                    (df_linha["Dt Evento"] < pd.to_datetime(data_fim))
                ]

            # -----------------------------------
            # CONTROLE DE VISÃO
            # -----------------------------------

            tipo_periodo = st.radio(
                "Período",
                ["Diário", "Semanal", "Mensal"],
                horizontal=True
            )

            # -----------------------------------
            # AGRUPAMENTO POR PERÍODO
            # -----------------------------------

            if tipo_periodo == "Diário":
                df_linha["Periodo"] = df_linha["Dt Evento"].dt.date

            elif tipo_periodo == "Semanal":
                df_linha["Periodo"] = df_linha["Dt Evento"].dt.to_period("W").apply(lambda r: r.start_time)

            else:  # Mensal
                df_linha["Periodo"] = df_linha["Dt Evento"].dt.to_period("M").apply(lambda r: r.start_time)

            # -----------------------------------
            # AGRUPAR
            # -----------------------------------

            df_grouped_linha = (
                df_linha.groupby(["Periodo", "Prazo"])
                .size()
                .unstack(fill_value=0)
                .reset_index()
            )

            # garantir colunas
            for col in ["NO PRAZO", "FORA DO PRAZO"]:
                if col not in df_grouped_linha.columns:
                    df_grouped_linha[col] = 0

            # -----------------------------------
            # CALCULAR OTD
            # -----------------------------------

            df_grouped_linha["Total"] = (
                df_grouped_linha["NO PRAZO"] + df_grouped_linha["FORA DO PRAZO"]
            )

            df_grouped_linha["OTD"] = (
                df_grouped_linha["NO PRAZO"] / df_grouped_linha["Total"]
            ) * 100

            # ordenar por data
            df_grouped_linha = df_grouped_linha.sort_values("Periodo")

            # -----------------------------------
            # GRÁFICO
            # -----------------------------------

            fig_linha = px.line(
                df_grouped_linha,
                x="Periodo",
                y="OTD",
                markers=True
            )

            # hover bonito
            fig_linha.update_traces(
                hovertemplate=
                "Período: %{x}<br>" +
                "OTD: %{y:.1f}%<br>" +
                "Pedidos: %{customdata}<extra></extra>",
                customdata=df_grouped_linha["Total"]
            )

            # layout
            fig_linha.update_layout(
                yaxis=dict(ticksuffix="%"),
                xaxis_title="",
                yaxis_title="OTD (%)"
            )

            # -----------------------------------
            # STREAMLIT
            # -----------------------------------

            st.plotly_chart(fig_linha, use_container_width=True)

            df_atraso = base_unificada.copy()

            # -----------------------------------
            # SOMENTE ENTREGUES
            # -----------------------------------

            df_atraso = df_atraso[df_atraso["Status"] == "ENTREGUE"].copy()

            # -----------------------------------
            # TRATAMENTO DATA
            # -----------------------------------

            df_atraso["Dt Evento"] = pd.to_datetime(
                df_atraso["Dt Evento"].astype(str).str.strip(),
                format="%d/%m/%Y %H:%M:%S",
                errors="coerce"
            )

            df_atraso["Previsao"] = pd.to_datetime(
                df_atraso["Previsao"].astype(str).str.strip(),
                format="%d/%m/%Y",
                errors="coerce"
            )

            df_atraso = df_atraso.dropna(subset=["Dt Evento", "Previsao"])

            # -----------------------------------
            # FILTRO DE DATA (🔥 MESMO DO DASH)
            # -----------------------------------

            if isinstance(data_evento, tuple) and len(data_evento) == 2:
                ini, fim = data_evento
                fim = fim + pd.Timedelta(days=1)

                df_atraso = df_atraso[
                    (df_atraso["Dt Evento"] >= pd.to_datetime(ini)) &
                    (df_atraso["Dt Evento"] < pd.to_datetime(fim))
                ]

            # -----------------------------------
            # CALCULAR DIAS DE ATRASO
            # -----------------------------------

            df_atraso["Dias Atraso"] = (
                df_atraso["Dt Evento"] - df_atraso["Previsao"]
            ).dt.days

            # -----------------------------------
            # SOMENTE ATRASOS
            # -----------------------------------

            df_atraso = df_atraso[df_atraso["Dias Atraso"] > 0]

            # -----------------------------------
            # AGRUPAR
            # -----------------------------------

            df_dist = (
                df_atraso.groupby("Dias Atraso")
                .size()
                .reset_index(name="Pedidos")
            )

            # -----------------------------------
            # TOTAL
            # -----------------------------------

            total_atrasados = df_dist["Pedidos"].sum()

            st.subheader("Distribuição de atrasos (dias)")

            st.markdown(f"**Total de pedidos em atraso: {total_atrasados:,}**")

            # -----------------------------------
            # ORDENAR POR DIAS (CORRETO)
            # -----------------------------------

            df_dist = df_dist.sort_values("Dias Atraso")

            # criar versão string só pra exibir
            df_dist["Dias_str"] = df_dist["Dias Atraso"].astype(int).astype(str)

            ordem = df_dist["Dias_str"].tolist()

            base = alt.Chart(df_dist)

            # -----------------------------------
            # BARRAS
            # -----------------------------------

            bars = base.mark_bar(
                color="#6baed6"
            ).encode(
                x=alt.X(
                    "Dias_str:N",
                    sort=ordem,
                    title="Dias de atraso",
                    axis=alt.Axis(labelAngle=0)  # 🔥 aqui
                ),
                y=alt.Y(
                    "Pedidos:Q",
                    title="Quantidade"
                )
            )

            text = base.mark_text(
                dy=-8,
                color="white"
            ).encode(
                x=alt.X(
                    "Dias_str:N",
                    sort=ordem,
                    axis=alt.Axis(labelAngle=0)  # 🔥 aqui também
                ),
                y="Pedidos:Q",
                text="Pedidos:Q"
            )

            # -----------------------------------
            # FINAL
            # -----------------------------------

            chart = (bars + text).properties(
                height=400
            )

            st.altair_chart(chart, use_container_width=True)

            st.subheader("Ocorrências - Pedidos em Atraso")

            # -----------------------------------
            # FILTRAR SOMENTE ATRASADOS COM OCORRÊNCIA
            # -----------------------------------

            df_atraso = df_perf[
                (df_perf["Prazo"] == "FORA DO PRAZO") &
                (df_perf["Ocorrencias"].notna()) &
                (df_perf["Ocorrencias"] != "")
            ].copy()

            # -----------------------------------
            # AGRUPAR OCORRÊNCIAS
            # -----------------------------------

            df_ocorrencias = (
                df_atraso.groupby("Ocorrencias")
                .size()
                .reset_index(name="Quantidade")
            )

            # -----------------------------------
            # TOTAL
            # -----------------------------------

            total_ocorrencias = df_ocorrencias["Quantidade"].sum()

            # -----------------------------------
            # PERCENTUAL
            # -----------------------------------

            df_ocorrencias["Percentual (%)"] = (
                df_ocorrencias["Quantidade"] / total_ocorrencias
            ) * 100

            # -----------------------------------
            # ORDENAR (maior para menor)
            # -----------------------------------

            df_ocorrencias = df_ocorrencias.sort_values(
                "Quantidade",
                ascending=False
            )

            # -----------------------------------
            # LINHA TOTAL
            # -----------------------------------

            linha_total = pd.DataFrame({
                "Ocorrencias": ["Total"],
                "Quantidade": [total_ocorrencias],
                "Percentual (%)": [100.0]
            })

            df_ocorrencias = pd.concat(
                [df_ocorrencias, linha_total],
                ignore_index=True
            )

            # -----------------------------------
            # FORMATAÇÃO
            # -----------------------------------

            df_ocorrencias["Percentual (%)"] = df_ocorrencias["Percentual (%)"].round(2)

            # -----------------------------------
            # EXIBIR
            # -----------------------------------

            st.dataframe(
                df_ocorrencias, 
                use_container_width=True,
                hide_index=True
            )

            # -----------------------------------
            # FILTRAR ATRASADOS
            # -----------------------------------

            df_atraso = df_perf[
                (df_perf["Prazo"] == "FORA DO PRAZO") &
                (df_perf["Destino"].notna())
            ].copy()

            if df_atraso.empty:

                st.info("Não há unidades ofensoras.")
                
            else:

                st.subheader("Unidades ofensores")

                # padronizar ocorrência
                df_atraso["Ocorrencias"] = (
                    df_atraso["Ocorrencias"]
                    .fillna("")
                    .astype(str)
                    .str.strip()
                )

                df_atraso.loc[
                    (df_atraso["Ocorrencias"] == "SEM OCORRENCIA") |
                    (df_atraso["Ocorrencias"].str.upper() == "NAN"),
                    "Ocorrencias"
                ] = "Atraso sem ocorrência"

                # -----------------------------------
                # AGRUPAR DESTINO + OCORRÊNCIA
                # -----------------------------------

                df_temp = (
                    df_atraso
                    .groupby(["Destino", "Ocorrencias"])
                    .size()
                    .reset_index(name="Qtd")
                )

                # -----------------------------------
                # MONTAR TEXTO DAS OCORRÊNCIAS
                # -----------------------------------

                def montar_ocorrencias(grupo):
                    grupo = grupo.sort_values("Qtd", ascending=False)
                    return ", ".join(
                        [f"{row['Ocorrencias']} ({row['Qtd']})" for _, row in grupo.iterrows()]
                    )

                df_ocorrencias = (
                    df_temp
                    .groupby("Destino")
                    .apply(montar_ocorrencias)
                    .reset_index(name="Ocorrências")
                )

                # -----------------------------------
                # QUANTIDADE TOTAL POR UNIDADE
                # -----------------------------------

                df_qtd = (
                    df_atraso
                    .groupby("Destino")
                    .size()
                    .reset_index(name="Quantidade")
                )

                # juntar
                df_final = df_ocorrencias.merge(df_qtd, on="Destino")

                # -----------------------------------
                # PERCENTUAL
                # -----------------------------------

                total_geral = df_final["Quantidade"].sum()

                df_final["Percentual (%)"] = (
                    df_final["Quantidade"] / total_geral
                ) * 100

                # -----------------------------------
                # ORDENAR (maior ofensores primeiro)
                # -----------------------------------

                df_final = df_final.sort_values(
                    "Quantidade",
                    ascending=False
                )

                # -----------------------------------
                # LINHA TOTAL
                # -----------------------------------

                linha_total = pd.DataFrame({
                    "Destino": ["Total"],
                    "Ocorrências": [""],
                    "Quantidade": [total_geral],
                    "Percentual (%)": [100.0]
                })

                df_final = pd.concat([df_final, linha_total], ignore_index=True)

                # -----------------------------------
                # FORMATAÇÃO
                # -----------------------------------

                df_final["Percentual (%)"] = df_final["Percentual (%)"].round(2)

                # renomear colunas igual imagem
                df_final = df_final.rename(columns={
                    "Destino": "Unidade"
                })

                # -----------------------------------
                # EXIBIR
                # -----------------------------------

                st.dataframe(
                    df_final, 
                    use_container_width=True,
                    hide_index=True
                )

                import io

                # -----------------------------------
                # BASE PARA EXPORT
                # -----------------------------------

                df_export = df_atraso.copy()

                # selecionar colunas relevantes (ajuste se quiser)
                colunas_export = [
                    "Codigo",
                    "Destino",
                    "Dt Evento",
                    "Previsao",
                    "Dias Atraso",
                    "Ocorrencias"
                ]

                colunas_existentes = [c for c in colunas_export if c in df_export.columns]

                df_export = df_export[colunas_existentes]

                # -----------------------------------
                # GERAR EXCEL EM MEMÓRIA
                # -----------------------------------

                buffer = io.BytesIO()

                with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
                    df_export.to_excel(writer, index=False, sheet_name="Atrasos")

                # -----------------------------------
                # BOTÃO DOWNLOAD
                # -----------------------------------

                st.download_button(
                    label="📥 Exportar base de atrasos",
                    data=buffer.getvalue(),
                    file_name="base_atrasos.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            st.subheader("Performance OTD – Visão Detalhada")

            df_tabela = base_unificada.copy()

            # -----------------------------------
            # SOMENTE ENTREGUES
            # -----------------------------------

            df_tabela = df_tabela[
                df_tabela["Prazo"].isin(["NO PRAZO", "FORA DO PRAZO"])
            ].copy()

            # -----------------------------------
            # TRATAMENTO DATA
            # -----------------------------------

            df_tabela["Dt Evento"] = df_tabela["Dt Evento"].astype(str).str.strip()

            df_tabela["Dt Evento"] = pd.to_datetime(
                df_tabela["Dt Evento"],
                format="%d/%m/%Y %H:%M:%S",
                errors="coerce"
            )

            # -----------------------------------
            # FILTRO DE DATA
            # -----------------------------------

            if isinstance(data_evento, tuple) and len(data_evento) == 2:
                data_ini, data_fim = data_evento
                data_fim = data_fim + pd.Timedelta(days=1)

                df_tabela = df_tabela[
                    (df_tabela["Dt Evento"] >= pd.to_datetime(data_ini)) &
                    (df_tabela["Dt Evento"] < pd.to_datetime(data_fim))
                ]

            # -----------------------------------
            # CONTROLES (CHECKBOX + INPUT)
            # -----------------------------------

            col4, col5, col6 = st.columns([1, 1, 1])

            with col4:
                usar_justificados = st.checkbox("Atrasos justificados", value=False)

            with col5:
                usar_baixa_indevida = st.checkbox("Baixas indevidas (1 dia de atraso e sem ocorrência)", value=False)

            with col6:
                dias_extra = st.number_input("Dias extras", min_value=0, max_value=10, value=0)

            # -----------------------------------
            # PREPARAÇÃO
            # -----------------------------------

            df_tabela["Ocorrencias"] = df_tabela["Ocorrencias"].astype(str).str.strip()

            df_tabela["Previsao"] = pd.to_datetime(
                df_tabela["Previsao"].astype(str).str.strip(),
                format="%d/%m/%Y",
                errors="coerce"
            )

            # aplicar dias extras
            df_tabela["Previsao Ajustada"] = df_tabela["Previsao"] + pd.to_timedelta(dias_extra, unit="D")

            # calcular atraso
            df_tabela["Dias Atraso"] = (
                df_tabela["Dt Evento"].dt.normalize() - df_tabela["Previsao Ajustada"].dt.normalize()
            ).dt.days

            # novo prazo
            df_tabela["Prazo Ajustado"] = df_tabela["Dias Atraso"].apply(
                lambda x: "NO PRAZO" if x <= 0 else "FORA DO PRAZO"
            )

            # -----------------------------------
            # REGRAS
            # -----------------------------------

            ocorrencias_validas = ["DEST. AUSENTE", "PROB. ENDEREÇO"]

            cond_justificado = (
                (df_tabela["Prazo Ajustado"] == "FORA DO PRAZO") &
                (df_tabela["Ocorrencias"].apply(lambda x: any(oc in x for oc in ocorrencias_validas)))
            )

            cond_baixa_indevida = (
                (df_tabela["Prazo Ajustado"] == "FORA DO PRAZO") &
                (df_tabela["Dias Atraso"] == 1) &
                (df_tabela["Ocorrencias"] == "")
            )

            # aplicar flags
            df_tabela["Justificado"] = False

            if usar_justificados:
                df_tabela.loc[cond_justificado, "Justificado"] = True

            if usar_baixa_indevida:
                df_tabela.loc[cond_baixa_indevida, "Justificado"] = True

            # no prazo ajustado
            df_tabela["No Prazo Ajustado"] = (
                (df_tabela["Prazo Ajustado"] == "NO PRAZO") |
                (df_tabela["Justificado"])
            ).astype(int)

            if tipo_visao == "Região":
                df_tabela["Grupo"] = df_tabela["UF"].map(mapa_regiao)
            else:
                df_tabela["Grupo"] = df_tabela["UF"]

            # -----------------------------------
            # AGRUPAR (USANDO PRAZO AJUSTADO)
            # -----------------------------------

            df_resumo = (
                df_tabela.groupby(["Grupo", "Prazo Ajustado"])
                .size()
                .unstack(fill_value=0)
                .reset_index()
            )

            for col in ["NO PRAZO", "FORA DO PRAZO"]:
                if col not in df_resumo.columns:
                    df_resumo[col] = 0

            # -----------------------------------
            # MÉTRICAS
            # -----------------------------------

            df_resumo["Total geral"] = df_resumo["NO PRAZO"] + df_resumo["FORA DO PRAZO"]

            df_resumo["Share"] = (df_resumo["Total geral"] / df_resumo["Total geral"].sum()) * 100

            df_resumo["OTD"] = (df_resumo["NO PRAZO"] / df_resumo["Total geral"]) * 100

            # -----------------------------------
            # OTD JUSTIFICADO
            # -----------------------------------

            df_just = (
                df_tabela.groupby("Grupo")["No Prazo Ajustado"]
                .sum()
                .reset_index()
            )

            df_resumo = df_resumo.merge(df_just, on="Grupo", how="left")

            df_resumo["OTD Justificado"] = (
                df_resumo["No Prazo Ajustado"] / df_resumo["Total geral"]
            ) * 100

            # -----------------------------------
            # ORDENAÇÃO
            # -----------------------------------

            if tipo_ordem == "Percentual":
                df_resumo = df_resumo.sort_values(
                    ["OTD", "Total geral"],
                    ascending=[False, False]
                )
            else:
                df_resumo = df_resumo.sort_values(
                    "Total geral",
                    ascending=False
                )

            # -----------------------------------
            # TOTAL GERAL
            # -----------------------------------

            total = pd.DataFrame({
                "Grupo": ["Total geral"],
                "NO PRAZO": [df_resumo["NO PRAZO"].sum()],
                "FORA DO PRAZO": [df_resumo["FORA DO PRAZO"].sum()],
                "No Prazo Ajustado": [df_resumo["No Prazo Ajustado"].sum()]
            })

            total["Total geral"] = total["NO PRAZO"] + total["FORA DO PRAZO"]
            total["Share"] = 100.0
            total["OTD"] = (total["NO PRAZO"] / total["Total geral"]) * 100
            total["OTD Justificado"] = (total["No Prazo Ajustado"] / total["Total geral"]) * 100

            df_resumo = pd.concat([df_resumo, total], ignore_index=True)

            # -----------------------------------
            # FORMATAR
            # -----------------------------------

            df_resumo["Share"] = df_resumo["Share"].map("{:.2f}%".format)
            df_resumo["OTD"] = df_resumo["OTD"].map("{:.2f}%".format)
            df_resumo["OTD Justificado"] = df_resumo["OTD Justificado"].map("{:.2f}%".format)

            df_resumo = df_resumo.rename(columns={
                "Grupo": "UF" if tipo_visao == "UF" else "Região",
                "NO PRAZO": "No prazo",
                "FORA DO PRAZO": "Fora do prazo"
            })

            primeira_coluna = "UF" if tipo_visao == "UF" else "Região"

            df_resumo = df_resumo[
                [
                    primeira_coluna,
                    "No prazo",
                    "Fora do prazo",
                    "Total geral",
                    "Share",
                    "OTD",
                    "OTD Justificado"
                ]
            ]

            # -----------------------------------
            # EXIBIR
            # -----------------------------------

            st.dataframe(
                df_resumo, 
                use_container_width=True,
                hide_index=True
            )

            with st.expander("🖼️ Imagens complementares"):
                        imagens = st.file_uploader(
                            "Importar imagens",
                            type=["png", "jpg", "jpeg"],
                            accept_multiple_files=True,
                            key="imgs_apresentacao"
                        )

                        if imagens:
                            for i, img in enumerate(imagens, start=1):
                                st.image(
                                    img,
                                    caption=f"Imagem {i}",
                                    use_container_width=True
                                )

    elif tipo_pedido == "Pedidos não concluídos (devoluções, sinistros, etc.)":

        df_encerrados = df[
            (df["Prazo"].isna()) |
            (df["Prazo"].astype(str).str.strip() == "")
        ].copy()

        if df_encerrados.empty:

            st.info("Não a pedidos não entregues.")

        else:

            with st.sidebar:

                st.markdown("### 📦 Base de pedidos não entregues")

                if df_encerrados is not None and not df_encerrados.empty:
                    st.metric(
                        label="Total de pedidos",
                        value=len(df_encerrados)
                    )
                else:
                    st.metric(
                        label="Total de pedidos",
                        value="—"
                    )

            if df_encerrados is not None and not df_encerrados.empty:
                excel_bytes = exportar_excel(df_encerrados)

                st.sidebar.download_button(
                    label="⬇️ Exportar base (.xlsx)",
                    data=excel_bytes,
                    file_name="df_encerrados.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.sidebar.download_button(
                    label="⬇️ Exportar base (.xlsx)",
                    data=b"",
                    disabled=True
                )

            status_counts = df_encerrados["Status"].value_counts().reset_index()
            status_counts.columns = ["Status", "Quantidade"]

            st.subheader("Pedidos não entregues (Devoluções, Sinistros, entre outros)")

            st.write(f"Total de pedidos: {len(df_encerrados):,}".replace(",", "."))

            max_val = status_counts["Quantidade"].max()

            base_st = alt.Chart(status_counts).encode(
                y=alt.Y(
                    "Status:N",
                    sort="-x",
                    title="Status"  # 👈 nome das categorias (opcional)
                ),
                x=alt.X(
                    "Quantidade:Q",
                    title="Quantidade de pedidos",  # 👈 título do eixo (fica embaixo)
                    scale=alt.Scale(domain=[0, max_val * 1.15])
                )
            )

            bars = base_st.mark_bar()

            text = base_st.mark_text(
                align="left",
                dx=8,
                color="white"
            ).encode(
                text=alt.Text("Quantidade:Q", format=",.0f")
            )

            chart = bars + text

            st.altair_chart(chart, use_container_width=True)

else:
    st.info("Importe planilhas no menu lateral para visualizar o dashboard.")
    