import streamlit as st
import pandas as pd
import plotly.express as px

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

    # Leitura única
    df_raw = pd.read_excel(file, header=None)

    # Detecta linha de cabeçalho
    linha_cabecalho = df_raw.notna().sum(axis=1).idxmax()

    # Mês de referência
    try:
        data_ref = pd.to_datetime(df_raw.iloc[2, 15], dayfirst=True)
        mes_ref = data_ref.strftime("%m/%Y")
    except:
        mes_ref = "Mês desconhecido"

    # Nome do cliente (D3)
    try:
        cliente = str(df_raw.iloc[2, 3]).strip()
    except:
        cliente = "Cliente desconhecido"

    # Base final
    df = pd.read_excel(file, skiprows=linha_cabecalho)
    df["Arquivo_Origem"] = mes_ref
    df["Cliente"] = cliente

    # Datas
    df["Previsao"] = pd.to_datetime(df.iloc[:, 17], dayfirst=True, errors="coerce")
    df["Dt Entrega"] = pd.to_datetime(df.iloc[:, 28], dayfirst=True, errors="coerce")

    previsao_d = df["Previsao"].dt.normalize()
    entrega_d = df["Dt Entrega"].dt.normalize()
    hoje = pd.Timestamp.today().normalize()

    # -----------------------------------
    # SLA (VETORIZADO)
    # -----------------------------------
    df["SLA"] = ""

    mask_fechado = entrega_d.notna() & previsao_d.notna()
    df.loc[mask_fechado & (entrega_d <= previsao_d), "SLA"] = "No prazo"
    df.loc[mask_fechado & (entrega_d > previsao_d), "SLA"] = "Atrasado"

    # -----------------------------------
    # PRAZO (VETORIZADO)
    # -----------------------------------
    df["Prazo"] = ""

    mask_aberto = entrega_d.isna() & previsao_d.notna()
    dias = (previsao_d - hoje).dt.days

    df.loc[mask_aberto & (dias < 0), "Prazo"] = (
        "Vencido há " + dias.abs().astype(str).str.zfill(2) + " dia(s)"
    )

    df.loc[mask_aberto & (dias == 0), "Prazo"] = "Vence hoje"

    df.loc[mask_aberto & (dias > 0), "Prazo"] = (
        "Faltam " + dias.astype(str).str.zfill(2) + " dia(s)"
    )

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
            mes_ref = df["Arquivo_Origem"].iloc[0]
            cliente = df["Cliente"].iloc[0]

            st.sidebar.write(
                f"📄 **{file.name}** carregado → {mes_ref} | 🏷️ {cliente}"
            )

    base_unificada = pd.concat(dfs, ignore_index=True)
    st.sidebar.success(f"{len(uploaded_files)} planilha(s) carregada(s)")

st.sidebar.markdown("---")

mapa_ocorrencias = {
    # Destinatário ausente
    "AUSENTE": "Dest. Ausente",
    "AUSENTE 2": "Dest. Ausente",
    "AUSENTE 3": "Dest. Ausente",
    "FECHADO": "Dest. Ausente",
    "FECHADO 2": "Dest. Ausente",
    "TEMPO DE ESPERA EXCEDIDO NO DESTINATARIO": "Dest. Ausente",

    # Pedido avariado
    "AVARIA / DANO PARCIAL": "Pedido Avariado",
    "AVARIA / DANO TOTAL": "Pedido Avariado",

    # Problema de endereço
    "DESTINATARIO DESCONHECIDO": "Prob. Endereço",
    "ENDERECO INSUFICIENTE": "Prob. Endereço",
    "ENDERECO NAO LOCALIZADO": "Prob. Endereço",
    "MUDOU-SE": "Prob. Endereço",
    "NUMERO NAO LOCALIZADO": "Prob. Endereço",

    # Agência
    "DESTINATARIO SOLICITOU RETIRAR NA UNIDADE": "Ag. Retirada Agência",

    # Last Mile
    "ATRASO TRANSPORTE": "Last Mile",
    "FALHA ENTREGA": "Last Mile",
    "SOLICITACAO ENTREGA FUTURA": "Last Mile",

    # Fiscal
    "EM ANÁLISE NO POSTO FISCAL": "Prob. Fiscal",
    "SAIDA FISCALIZACAO": "Prob. Fiscal",

    # Zona rural
    "ENDERECO EM ZONA RURAL": "Zona Rural",

    # Middle Mile
    "BUSCA": "Middle Mile",
    "NAO ENTROU NA UNIDADE": "Middle Mile",

    # Feriado
    "FECHADO EM VESPERA OU APOS FERIADO": "Feriado",

    # Layout / etiqueta
    "TERMO DE IRREGULARIDADE": "Erro Layout Etiqueta",

    # Rodovia
    "TRAFEGO INTERROMPIDO": "Rodovia Interditada",

    # Clima
    "TEMPORAL": "Prob. Climático",

    # Triagem
    "ERRO DE TRIAGEM / SEPARACAO": "Erro de triagem",

    # Acareação
    "FALTA DE ACAREAÇÃO / TOTAL": "Acareação",

    # Área de risco
    "RESTRICAO DE ACESSO / MOVIMENTACAO": "Área de Risco",

    # Sinistro
    "SINISTRO / ACIDENTE TRANSPORTE": "Sinistro",

    # Pedido recusado
    "RECUSADO": "Pedido Recusado",
    "RECUSADO - DIVERGENCIA DE PEDIDO": "Pedido Recusado",
    "RECUSADO - NAO PAGA FRETE": "Pedido Recusado",
}

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

if base_unificada is not None and not base_unificada.empty:
    # cria a coluna vazia
    base_unificada["Ocorrencias"] = ""

    # máscara de pedidos atrasados
    mask_atrasado = base_unificada["SLA"].str.upper().eq("ATRASADO")

    # descrição tratada (sem transformar vazio em NaN definitivo)
    descricao_limpa = (
        base_unificada["Descricao"]
        .astype(str)
        .str.strip()
    )

    # aplica somente nos atrasados
    base_unificada.loc[mask_atrasado, "Ocorrencias"] = (
        descricao_limpa[mask_atrasado]
        .str.upper()
        .map(mapa_ocorrencias)
        # se não estiver no mapa → volta a descrição original
        .fillna(descricao_limpa[mask_atrasado])
        # se ainda assim for "nan" → string vazia
        .replace("nan", "")
    )

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
# DASHBOARD
# -------------------------------------------
if base_unificada is not None:

    if {"UF", "SLA", "Dt Entrega"}.issubset(base_unificada.columns):

        hoje = pd.Timestamp.today().normalize()

        status_excluir = [
            "DEVOLVIDO",
            "EM DEVOLUCAO",
            "LISTA DEVOLUCAO",
            "TRAVADO",
            "SINISTRO",
            "UNITIZADO"
        ]

        descricao_custodia_excluir = [
            "AVARIA / DANO PARCIAL",
            "AVARIA / DANO TOTAL",
            "EXTRAVIO PARCIAL",
            "EXTRAVIO TOTAL",
            "PERDA POR PRAZO SLA / TOTAL",
            "ERRO DO EMISSOR",
            "INDENIZACAO RECUSADA",
            "INDICIO VIOLACAO"
        ]

        mask_custodia_excluir = (
            base_unificada["Status"].str.upper().str.strip().eq("CUSTODIA") &
            base_unificada["Descricao"].str.upper().str.strip().isin(descricao_custodia_excluir)
        )

        df_abertos = base_unificada[
            base_unificada["Dt Entrega"].isna() &
            base_unificada["Previsao"].notna() &
            ~base_unificada["Status"].str.upper().str.strip().isin(status_excluir) &
            ~mask_custodia_excluir
        ].copy()

        with st.sidebar:

            st.markdown("### 📦 Base em abertos")

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
                file_name="em_aberto.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.sidebar.download_button(
                label="⬇️ Exportar base (.xlsx)",
                data=b"",
                disabled=True
            )

        df_abertos["Dias_Atraso"] = (hoje - df_abertos["Previsao"].dt.normalize()).dt.days

        # Somente atrasados
        df_atrasados = df_abertos[df_abertos["Dias_Atraso"] > 0]

        contagem_atrasos = (
            df_atrasados
            .groupby("Dias_Atraso")
            .size()
            .reset_index(name="Qtde")
            .sort_values("Dias_Atraso")
        )

        # 🔴 CASO 1: não há atrasos
        if contagem_atrasos.empty:

            st.info("Não há pedidos em atraso no momento.")

        else:
            LIMITE_DIAS = 60

            contagem_grafico = contagem_atrasos[
                contagem_atrasos["Dias_Atraso"] <= LIMITE_DIAS
            ]

            contagem_excedente = contagem_atrasos[
                contagem_atrasos["Dias_Atraso"] > LIMITE_DIAS
            ]

            # 🔹 GRÁFICO (somente se houver até 60 dias)
            if not contagem_grafico.empty:

                st.subheader("⏳ Pedidos em aberto — Dias em atraso (até 60 dias)")
                st.metric(
                    label="Total de pedidos atrasados",
                    value=len(df_atrasados)
                )

                fig_atraso = px.bar(
                    contagem_grafico,
                    x="Dias_Atraso",
                    y="Qtde",
                    text="Qtde",
                    labels={
                        "Dias_Atraso": "Dias em atraso",
                        "Qtde": "Quantidade de pedidos"
                    }
                )

                fig_atraso.update_traces(
                    textposition="outside",
                    customdata=(
                        contagem_grafico["Qtde"]
                        / contagem_grafico["Qtde"].sum() * 100
                    ).round(1),
                    hovertemplate=(
                        "Dias em atraso: %{x}<br>"
                        "Pedidos: %{y}<br>"
                        "Percentual: %{customdata}%"
                        "<extra></extra>"
                    )
                )

                fig_atraso.update_layout(
                    xaxis_tickmode="linear",
                    bargap=0.2
                )

                st.plotly_chart(fig_atraso, use_container_width=True)

            # 🔹 TABELA (> 60 dias)
            if not contagem_excedente.empty:

                st.subheader("📋 Pedidos com atraso superior a 60 dias")

                st.dataframe(
                    contagem_excedente.rename(
                        columns={
                            "Dias_Atraso": "Dias em atraso",
                            "Qtde": "Quantidade de pedidos"
                        }
                    ),
                    use_container_width=True,
                    hide_index=True
                )
            
            # ==============================
            # 🔽 SELEÇÃO DE FAIXAS DE ATRASO
            # ==============================

            if not df_atrasados.empty:

                KEY_ATRASOS = "atrasos_sel"

                # 🔽 Opções da multiselect = TODOS os dias em atraso
                opcoes_atraso = (
                    contagem_atrasos["Dias_Atraso"]
                    .dropna()
                    .astype(int)
                    .sort_values()
                    .tolist()
                )

                # Inicializa seleção com TODOS
                if "faixas_atraso" not in st.session_state:
                    st.session_state["faixas_atraso"] = opcoes_atraso

                if KEY_ATRASOS not in st.session_state:
                    st.session_state[KEY_ATRASOS] = []

                col_a, col_b, _ = st.columns([0.2, 0.25, 1.5])

                with col_a:
                    if st.button("✅ Marcar todos"):
                        st.session_state[KEY_ATRASOS] = opcoes_atraso.copy()

                with col_b:
                    if st.button("❌ Limpar seleção"):
                        st.session_state[KEY_ATRASOS] = []

                atrasos_sel = st.multiselect(
                    "🔍 Ver pedidos por dias de atraso",
                    options=opcoes_atraso,
                    key=KEY_ATRASOS,
                    format_func=lambda x: f"{x} dias de atraso"
                )

                if atrasos_sel:

                    df_detalhe_atraso = df_atrasados[
                        df_atrasados["Dias_Atraso"].isin(atrasos_sel)
                    ].copy()

                    st.subheader(
                        f"📄 Pedidos com atraso selecionado "
                        f"({len(df_detalhe_atraso)} pedidos)"
                    )

                    st.dataframe(
                        df_detalhe_atraso,
                        use_container_width=True,
                        hide_index=True
                    )

                    excel_bytes = exportar_excel(df_detalhe_atraso)

                    st.download_button(
                        "⬇️ Exportar pedidos selecionados",
                        excel_bytes,
                        file_name="pedidos_atrasos_selecionados.xlsx"
                    )

                else:
                    st.info("Selecione uma ou mais faixas de atraso.")

        st.markdown("---")

        # -------------------------------
        # PEDIDOS A VENCER
        # -------------------------------
        df_a_vencer = df_abertos[df_abertos["Dias_Atraso"] <= 0].copy()

        # Dias para vencer (positivo)
        df_a_vencer["Dias_Para_Vencer"] = df_a_vencer["Dias_Atraso"].abs()

        contagem_vencer = (
            df_a_vencer
            .groupby("Dias_Para_Vencer")
            .size()
            .reset_index(name="Qtde")
            .sort_values("Dias_Para_Vencer")
        )

        total_avencer = contagem_vencer["Qtde"].sum()

        contagem_vencer["Percentual"] = (
            contagem_vencer["Qtde"] / total_avencer * 100
        )

        if not contagem_vencer.empty:

            st.subheader("⏱️ Pedidos em aberto — Dias a vencer")
            st.metric(
                    label="Total de pedidos a vencer",
                    value=len(df_a_vencer)
                )

            fig_vencer = px.bar(
                contagem_vencer,
                x="Dias_Para_Vencer",
                y="Qtde",
                labels={
                    "Dias_Para_Vencer": "Dias para vencer",
                    "Qtde": "Quantidade de pedidos"
                },
                text="Qtde"
            )

            fig_vencer.update_traces(
                customdata=contagem_vencer[["Percentual"]],
                hovertemplate=(
                    "Dias para vencer: %{x}<br>"
                    "Pedidos: %{y}<br>"
                    "Percentual: %{customdata[0]:.1f}%"
                    "<extra></extra>"
                )
            )

            fig_vencer.update_traces(textposition="outside")
            fig_vencer.update_layout(
                xaxis_tickmode="linear",
                bargap=0.2
            )

            st.plotly_chart(fig_vencer, use_container_width=True)

        else:
            st.info("Não há pedidos a vencer no momento.")

        # ==============================
        # 🔽 SELEÇÃO DE FAIXAS — PEDIDOS A VENCER
        # ==============================

        if not contagem_vencer.empty:

            KEY_VENCER = "vencer_sel"

            # 🔽 Opções da multiselect = TODOS os dias a vencer
            opcoes_vencer = (
                contagem_vencer["Dias_Para_Vencer"]
                .dropna()
                .astype(int)
                .sort_values()
                .tolist()
            )

            # Inicializa seleção com TODOS
            if "faixas_avencer" not in st.session_state:
                st.session_state["faixas_avencer"] = opcoes_vencer.copy()

            col_a, col_b, _ = st.columns([0.2, 0.25, 1.5])

            with col_a:
                if st.button("✅ Marcar todos", key="btn_vencer_all"):
                    st.session_state[KEY_VENCER] = opcoes_vencer.copy()

            with col_b:
                if st.button("❌ Limpar seleção", key="btn_vencer_clear"):
                    st.session_state[KEY_VENCER] = []

            vencer_sel = st.multiselect(
                "🔍 Ver pedidos por dias para vencer",
                options=opcoes_vencer,
                key=KEY_VENCER,
                format_func=lambda x: f"{x} dias para vencer"
            )

            if vencer_sel:

                df_detalhe_vencer = df_a_vencer[
                    df_a_vencer["Dias_Para_Vencer"].isin(vencer_sel)
                ].copy()

                st.subheader(
                    f"📄 Pedidos a vencer selecionados "
                    f"({len(df_detalhe_vencer)} pedidos)"
                )

                st.dataframe(
                    df_detalhe_vencer,
                    use_container_width=True,
                    hide_index=True
                )

                excel_bytes = exportar_excel(df_detalhe_vencer)

                st.download_button(
                    "⬇️ Exportar pedidos a vencer selecionados",
                    excel_bytes,
                    file_name="pedidos_a_vencer_selecionados.xlsx",
                    key="download_vencer"
                )

            else:
                st.info("Selecione uma ou mais faixas de dias para vencer.")

        df_entregues = base_unificada[base_unificada["Dt Entrega"].notna()].copy()

        if not df_entregues.empty:

            st.markdown("---")

            st.subheader("📊 Performance OTD por UF")
            st.metric(
                label="Total de pedidos",
                value=len(df_entregues)
            )
            
            c1, c2, _ = st.columns([0.25, 0.25, 0.5], gap="small")

            with c1:
                min_d = df_entregues["Dt Entrega"].min().date()
                max_d = df_entregues["Dt Entrega"].max().date()
                intervalo = st.date_input(
                    "Filtro: Dt Entrega",
                    value=(min_d, max_d),
                    min_value=min_d,
                    max_value=max_d
                )

            with c2:
                criterio = st.selectbox(
                    "Ordenar por",
                    ["Percentual (100% → 0%)", "Número de pedidos (maior → menor)"]
                )

            if isinstance(intervalo, tuple) and len(intervalo) == 2:
                ini, fim = intervalo

                if ini and fim:
                    df_entregues = df_entregues[
                        (df_entregues["Dt Entrega"] >= pd.to_datetime(ini)) &
                        (df_entregues["Dt Entrega"] <= pd.to_datetime(fim) + pd.Timedelta(days=1))
                    ]

            # AGRUPAMENTO
            contagem = df_entregues.groupby(["UF", "SLA"]).size().unstack(fill_value=0)
            contagem["Total"] = contagem.sum(axis=1)

            contagem["Percentual No Prazo"] = contagem.get("No prazo", 0) / contagem["Total"] * 100
            contagem["Percentual Atrasado"] = contagem.get("Atrasado", 0) / contagem["Total"] * 100

            total = pd.DataFrame({
                "No prazo": [contagem.get("No prazo", pd.Series(dtype=int)).sum()],
                "Atrasado": [contagem.get("Atrasado", pd.Series(dtype=int)).sum()],
                "Total": [contagem["Total"].sum()]
            }, index=["TOTAL"])

            total["Percentual No Prazo"] = total["No prazo"] / total["Total"] * 100
            total["Percentual Atrasado"] = total["Atrasado"] / total["Total"] * 100

            contagem = pd.concat([contagem, total])
            df_grouped = contagem.reset_index().rename(columns={"index": "UF"})

            df_total = df_grouped[df_grouped["UF"] == "TOTAL"]
            df_ufs = df_grouped[df_grouped["UF"] != "TOTAL"]

            if criterio == "Percentual (100% → 0%)":
                df_ufs = df_ufs.sort_values("Percentual No Prazo", ascending=False)
            else:
                df_ufs = df_ufs.sort_values("Total", ascending=False)

            df_grouped = pd.concat([df_ufs, df_total], ignore_index=True)

            # GRÁFICO
            fig = px.bar(
                df_grouped,
                x="UF",
                y=["Percentual No Prazo", "Percentual Atrasado"],
                labels={"value": "Percentual (%)", "variable": "SLA"},
                color_discrete_map={
                    "Percentual No Prazo": "#2ca02c",
                    "Percentual Atrasado": "#d62728"
                }
            )

            # Quantidade correta no hover
            fig.data[0].customdata = df_grouped["No prazo"]
            fig.data[1].customdata = df_grouped["Atrasado"]

            fig.data[0].hovertemplate = (
                "UF: %{x}<br>"
                "No prazo<br>"
                "Percentual: %{y:.1f}%<br>"
                "Qtd: %{customdata}<extra></extra>"
            )

            fig.data[1].hovertemplate = (
                "UF: %{x}<br>"
                "Atrasos<br>"
                "Percentual: %{y:.1f}%<br>"
                "Qtd: %{customdata}<extra></extra>"
            )

            # Total no topo
            fig.add_scatter(
                x=df_grouped["UF"],
                y=[100] * len(df_grouped),
                mode="text",
                text=df_grouped["Total"],
                textposition="top center",
                showlegend=False,
                hoverinfo="skip"
            )

            # Percentual dentro da barra
            fig.update_traces(
                texttemplate="%{y:.0f}%",
                textposition="inside",
                selector=dict(type="bar")
            )

            fig.update_yaxes(range=[0, 100], ticksuffix="%")
            fig.update_layout(barmode="stack", legend_title_text="")


            st.plotly_chart(fig, use_container_width=True)

            # ==============================
            # 📊 PERFORMANCE POR REGIÃO
            # ==============================

            df_regiao = df_entregues.copy()

            df_regiao["Regiao"] = (
                df_regiao["UF"]
                .map(mapa_regiao)
                .fillna("Outros")
            )

            if not df_regiao.empty:

                st.subheader("🌎 Performance OTD por Região")

                # AGRUPAMENTO
                contagem_regiao = (
                    df_regiao
                    .groupby(["Regiao", "SLA"])
                    .size()
                    .unstack(fill_value=0)
                )

                contagem_regiao["Total"] = contagem_regiao.sum(axis=1)

                contagem_regiao["Percentual No Prazo"] = (
                    contagem_regiao.get("No prazo", 0) / contagem_regiao["Total"] * 100
                )

                contagem_regiao["Percentual Atrasado"] = (
                    contagem_regiao.get("Atrasado", 0) / contagem_regiao["Total"] * 100
                )

                # TOTAL GERAL
                linha_total_regiao = pd.DataFrame({
                    "No prazo": [contagem_regiao.get("No prazo", pd.Series(dtype=int)).sum()],
                    "Atrasado": [contagem_regiao.get("Atrasado", pd.Series(dtype=int)).sum()],
                    "Total": [contagem_regiao["Total"].sum()]
                }, index=["TOTAL"])

                linha_total_regiao["Percentual No Prazo"] = (
                    linha_total_regiao["No prazo"] / linha_total_regiao["Total"] * 100
                )

                linha_total_regiao["Percentual Atrasado"] = (
                    linha_total_regiao["Atrasado"] / linha_total_regiao["Total"] * 100
                )

                contagem_regiao = pd.concat([contagem_regiao, linha_total_regiao])

                df_regiao_grouped = (
                    contagem_regiao
                    .reset_index()
                    .rename(columns={"index": "Regiao"})
                )

                # ORDENAÇÃO (mesma lógica do UF)
                if criterio == "Percentual (100% → 0%)":
                    df_body = df_regiao_grouped[df_regiao_grouped["Regiao"] != "TOTAL"] \
                        .sort_values("Percentual No Prazo", ascending=False)
                else:
                    df_body = df_regiao_grouped[df_regiao_grouped["Regiao"] != "TOTAL"] \
                        .sort_values("Total", ascending=False)

                df_total = df_regiao_grouped[df_regiao_grouped["Regiao"] == "TOTAL"]

                df_regiao_grouped = pd.concat([df_body, df_total], ignore_index=True)

                # GRÁFICO
                fig_regiao = px.bar(
                    df_regiao_grouped,
                    x="Regiao",
                    y=["Percentual No Prazo", "Percentual Atrasado"],
                    labels={"value": "Percentual (%)", "variable": "SLA"},
                    color_discrete_map={
                        "Percentual No Prazo": "#2ca02c",
                        "Percentual Atrasado": "#d62728"
                    }
                )

                # Hover com quantidade
                fig_regiao.data[0].customdata = df_regiao_grouped["No prazo"]
                fig_regiao.data[1].customdata = df_regiao_grouped["Atrasado"]

                fig_regiao.data[0].hovertemplate = (
                    "Região: %{x}<br>"
                    "No prazo<br>"
                    "Percentual: %{y:.1f}%<br>"
                    "Qtd: %{customdata}<extra></extra>"
                )

                fig_regiao.data[1].hovertemplate = (
                    "Região: %{x}<br>"
                    "Atrasado<br>"
                    "Percentual: %{y:.1f}%<br>"
                    "Qtd: %{customdata}<extra></extra>"
                )

                # Total no topo
                fig_regiao.add_scatter(
                    x=df_regiao_grouped["Regiao"],
                    y=[100] * len(df_regiao_grouped),
                    mode="text",
                    text=df_regiao_grouped["Total"],
                    textposition="top center",
                    showlegend=False,
                    hoverinfo="skip"
                )

                fig_regiao.update_traces(
                    texttemplate="%{y:.0f}%",
                    textposition="inside",
                    selector=dict(type="bar")
                )

                fig_regiao.update_yaxes(range=[0, 100], ticksuffix="%")
                fig_regiao.update_layout(
                    barmode="stack",
                    legend_title_text=""
                )

                st.plotly_chart(fig_regiao, use_container_width=True, key="perf_regiao")
            
            else:
                st.info("Não há dados suficientes para a performance regional.")
        
            df_perf_dia = base_unificada[
                base_unificada["Dt Entrega"].notna() &
                base_unificada["Previsao"].notna()
            ].copy()

            df_perf_dia["Data_Entrega"] = df_perf_dia["Dt Entrega"].dt.date
            df_perf_dia["No_Prazo"] = (
                df_perf_dia["Dt Entrega"].dt.date
                <= df_perf_dia["Previsao"].dt.date
            )

            if isinstance(intervalo, tuple) and len(intervalo) == 2:
                ini, fim = intervalo
                df_perf_dia = df_perf_dia[
                    (df_perf_dia["Data_Entrega"] >= ini) &
                    (df_perf_dia["Data_Entrega"] <= fim)
                ]

                
            performance_diaria = (
                df_perf_dia
                .groupby("Data_Entrega")
                .agg(
                    Total=("No_Prazo", "size"),
                    No_Prazo=("No_Prazo", "sum")
                )
                .reset_index()
            )

            performance_diaria["Performance"] = (
                performance_diaria["No_Prazo"] / performance_diaria["Total"] * 100
            )

            if not performance_diaria.empty:

                st.subheader("📈 Performance OTD diária")

                fig_perf_dia = px.line(
                    performance_diaria,
                    x="Data_Entrega",
                    y="Performance",
                    markers=True,
                    labels={
                        "Data_Entrega": "Data",
                        "Performance": "% no prazo"
                    }
                )

                fig_perf_dia.update_yaxes(range=[0, 100])

                st.plotly_chart(fig_perf_dia, use_container_width=True)

            else:
                st.info("Não há dados suficientes para a performance diária.")

            # ==============================
            # 📈 PERFORMANCE SEMANAL
            # ==============================

            df_perf_semana = base_unificada[
                base_unificada["Dt Entrega"].notna() &
                base_unificada["Previsao"].notna()
            ].copy()

            # Datas normalizadas
            df_perf_semana["Data_Entrega"] = df_perf_semana["Dt Entrega"].dt.normalize()

            # No prazo
            df_perf_semana["No_Prazo"] = (
                df_perf_semana["Dt Entrega"].dt.date
                <= df_perf_semana["Previsao"].dt.date
            )

            # Semana (segunda-feira como início)
            df_perf_semana["Semana"] = (
                df_perf_semana["Data_Entrega"]
                - pd.to_timedelta(df_perf_semana["Data_Entrega"].dt.weekday, unit="D")
            )

            # 🔹 Aplica o MESMO filtro de data, se existir
            if isinstance(intervalo, tuple) and len(intervalo) == 2:
                ini, fim = intervalo

                ini_dt = pd.to_datetime(ini)
                fim_dt = pd.to_datetime(fim) + pd.Timedelta(days=1)

                df_perf_semana = df_perf_semana[
                    (df_perf_semana["Data_Entrega"] >= ini_dt) &
                    (df_perf_semana["Data_Entrega"] < fim_dt)
                ]

            performance_semanal = (
                df_perf_semana
                .groupby("Semana")
                .agg(
                    Total=("No_Prazo", "size"),
                    No_Prazo=("No_Prazo", "sum")
                )
                .reset_index()
            )

            performance_semanal["Performance"] = (
                performance_semanal["No_Prazo"] / performance_semanal["Total"] * 100
            )

            if not performance_semanal.empty:

                st.subheader("📈 Performance OTD semanal")

                fig_perf_semana = px.line(
                    performance_semanal,
                    x="Semana",
                    y="Performance",
                    markers=True,
                    labels={
                        "Semana": "Semana",
                        "Performance": "% no prazo"
                    }
                )

                fig_perf_semana.update_yaxes(range=[0, 100])

                fig_perf_semana.update_traces(
                    hovertemplate=(
                        "Semana: %{x|%d/%m/%Y}<br>"
                        "Performance: %{y:.1f}%<extra></extra>"
                    )
                )

                st.plotly_chart(
                    fig_perf_semana,
                    use_container_width=True,
                    key="perf_semanal"
                )

            else:
                st.info("Não há dados suficientes para a performance semanal.")

            # ==============================
            # 📊 PERFORMANCE MENSAL
            # ==============================

            df_perf_mes = base_unificada[
                base_unificada["Dt Entrega"].notna() &
                base_unificada["Previsao"].notna()
            ].copy()

            # Datas normalizadas
            df_perf_mes["Data_Entrega"] = df_perf_mes["Dt Entrega"].dt.normalize()

            # No prazo
            df_perf_mes["No_Prazo"] = (
                df_perf_mes["Dt Entrega"].dt.date
                <= df_perf_mes["Previsao"].dt.date
            )

            # Mês (início do mês)
            df_perf_mes["Mes"] = df_perf_mes["Data_Entrega"].dt.to_period("M").dt.to_timestamp()

            # 🔹 Aplica filtro de data (se houver)
            if isinstance(intervalo, tuple) and len(intervalo) == 2:
                ini, fim = intervalo

                ini_dt = pd.to_datetime(ini)
                fim_dt = pd.to_datetime(fim) + pd.Timedelta(days=1)

                df_perf_mes = df_perf_mes[
                    (df_perf_mes["Data_Entrega"] >= ini_dt) &
                    (df_perf_mes["Data_Entrega"] < fim_dt)
                ]

            performance_mensal = (
                df_perf_mes
                .groupby("Mes")
                .agg(
                    Total=("No_Prazo", "size"),
                    No_Prazo=("No_Prazo", "sum")
                )
                .reset_index()
            )

            performance_mensal["Performance"] = (
                performance_mensal["No_Prazo"] / performance_mensal["Total"] * 100
            )

            if not performance_mensal.empty:

                st.subheader("📊 Performance OTD mensal")

                fig_perf_mes = px.bar(
                    performance_mensal,
                    x="Mes",
                    y="Performance",
                    text=performance_mensal["Performance"].round(1),
                    labels={
                        "Mes": "Mês",
                        "Performance": "% no prazo"
                    }
                )

                fig_perf_mes.update_yaxes(range=[0, 100])

                fig_perf_mes.update_traces(
                    texttemplate="%{text}%",
                    textposition="outside",
                    hovertemplate=(
                        "Mês: %{x|%b/%Y}<br>"
                        "Performance: %{y:.1f}%<extra></extra>"
                    )
                )

                fig_perf_mes.update_layout(
                    xaxis=dict(
                        tickmode="linear",
                        dtick="M1",          # 1 mês
                        tickformat="%b/%Y"  # Jan/2024
                    ),
                    bargap=0.3
                )

                st.plotly_chart(
                    fig_perf_mes,
                    use_container_width=True,
                    key="perf_mensal"
                )

            else:
                st.info("Não há dados suficientes para a performance mensal.")


            # Ocorrências que justificam atraso
            JUSTIFICADAS = ["DEST. AUSENTE", "PROB. ENDEREÇO"]

            # =========================
            # PERFORMANCE ORIGINAL
            # =========================
            df_perf_original = df_entregues.copy()

            resumo_original = (
                df_perf_original["SLA"]
                .value_counts()
                .rename_axis("Status")
                .reset_index(name="Quantidade")
            )

            # =========================
            # PERFORMANCE JUSTIFICADA
            # =========================
            df_perf_just = df_entregues.copy()

            # Default: tudo atrasado
            df_perf_just["Status_Justificado"] = "Atrasado"

            # No prazo original
            df_perf_just.loc[
                df_perf_just["SLA"].str.upper() == "NO PRAZO",
                "Status_Justificado"
            ] = "No prazo"

            # Atrasos justificados passam a contar como NO PRAZO
            df_perf_just.loc[
                (df_perf_just["SLA"].str.upper() == "ATRASADO") &
                (df_perf_just["Ocorrencias"].str.upper().isin(JUSTIFICADAS)),
                "Status_Justificado"
            ] = "No prazo"

            resumo_justificado = (
                df_perf_just["Status_Justificado"]
                .value_counts()
                .rename_axis("Status")
                .reset_index(name="Quantidade")
            )

            # =========================
            # LAYOUT LADO A LADO
            # =========================

            if not resumo_justificado.empty:
            
                col1, col2 = st.columns(2)

                with col1:
                    st.markdown("### 📊 Performance OTD Original")
                    fig_orig = px.pie(
                        resumo_original,
                        names="Status",
                        values="Quantidade",
                        hole=0.4
                    )
                    fig_orig.update_traces(textinfo="percent+label")
                    st.plotly_chart(
                        fig_orig,
                        use_container_width=True,
                        key="pizza_performance_original"
                    )


                with col2:
                    st.markdown("### ✅ Performance OTD Justificada")
                    fig_just = px.pie(
                        resumo_justificado,
                        names="Status",
                        values="Quantidade",
                        hole=0.4
                    )
                    fig_just.update_traces(textinfo="percent+label")
                    st.plotly_chart(
                        fig_just,
                        use_container_width=True,
                        key="pizza_performance_justificada"
                    )
            else:
                st.info("Não há dados suficientes para a performance justificada.")

            if not df_entregues.empty:

                st.subheader("📋 Resumo Performance OTD por UF")

                df_resumo = df_entregues.copy()

                # Flags
                df_resumo["No_Prazo"] = df_resumo["SLA"].str.upper() == "NO PRAZO"

                df_resumo["Atrasado"] = df_resumo["SLA"].str.upper() == "ATRASADO"

                df_resumo["Justificado"] = (
                    (df_resumo["SLA"].str.upper() == "ATRASADO") &
                    (df_resumo["Ocorrencias"].str.upper().isin(JUSTIFICADAS))
                )

                # Agrupamento por UF
                tabela_otd = (
                    df_resumo
                    .groupby("UF")
                    .agg(
                        No_Prazo=("No_Prazo", "sum"),
                        Atrasado=("Atrasado", "sum"),
                        Total=("SLA", "size"),
                        Justificados=("Justificado", "sum")
                    )
                    .reset_index()
                )

                # OTDs
                tabela_otd["OTD Original (%)"] = (
                    tabela_otd["No_Prazo"] / tabela_otd["Total"] * 100
                ).round(1)

                tabela_otd["OTD Justificado (%)"] = (
                    (tabela_otd["No_Prazo"] + tabela_otd["Justificados"])
                    / tabela_otd["Total"] * 100
                ).round(1)

                # 🔹 Linha TOTAL
                linha_total = pd.DataFrame({
                    "UF": ["TOTAL"],
                    "No_Prazo": [tabela_otd["No_Prazo"].sum()],
                    "Atrasado": [tabela_otd["Atrasado"].sum()],
                    "Total": [tabela_otd["Total"].sum()],
                    "Justificados": [tabela_otd["Justificados"].sum()]
                })

                linha_total["OTD Original (%)"] = (
                    linha_total["No_Prazo"] / linha_total["Total"] * 100
                ).round(1)

                linha_total["OTD Justificado (%)"] = (
                    (linha_total["No_Prazo"] + linha_total["Justificados"])
                    / linha_total["Total"] * 100
                ).round(1)

                tabela_otd = pd.concat([tabela_otd, linha_total], ignore_index=True)

                # Ordenar por total de pedidos (maior → menor), mantendo TOTAL no final
                tabela_otd_ufs = tabela_otd[tabela_otd["UF"] != "TOTAL"] \
                    .sort_values("Total", ascending=False)

                tabela_otd_total = tabela_otd[tabela_otd["UF"] == "TOTAL"]

                tabela_otd = pd.concat(
                    [tabela_otd_ufs, tabela_otd_total],
                    ignore_index=True
                )

                # Exibição
                st.dataframe(
                    tabela_otd[
                        [
                            "UF",
                            "No_Prazo",
                            "Atrasado",
                            "Total",
                            "OTD Original (%)",
                            "OTD Justificado (%)"
                        ]
                    ],
                    use_container_width=True,
                    hide_index=True
                )
            else:
                st.info("Não há dados suficientes para exibir o resumo.")

            df_ocorrencias = df_entregues[
                (df_entregues["SLA"].str.upper() == "ATRASADO") &
                df_entregues["Ocorrencias"].notna() &
                (df_entregues["Ocorrencias"].str.strip() != "")
            ].copy()

            if df_ocorrencias.empty:
                st.info("Não há ocorrências registradas para pedidos em atraso.")

            else:
                st.subheader("📋 Ocorrências — Pedidos em atraso")

                tabela_ocorrencias = (
                    df_ocorrencias
                    .groupby("Ocorrencias")
                    .size()
                    .reset_index(name="Quantidade")
                    .sort_values("Quantidade", ascending=False)
                )

                total_qtd = tabela_ocorrencias["Quantidade"].sum()

                tabela_ocorrencias["Percentual"] = (
                    tabela_ocorrencias["Quantidade"] / total_qtd * 100
                ).round(1)

                # 🔹 linha TOTAL
                linha_total = pd.DataFrame({
                    "Ocorrencias": ["TOTAL"],
                    "Quantidade": [total_qtd],
                    "Percentual": [100.0]
                })

                tabela_ocorrencias = pd.concat(
                    [tabela_ocorrencias, linha_total],
                    ignore_index=True
                )

                st.dataframe(
                    tabela_ocorrencias,
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

            df_base_filtro_data = base_unificada.copy()

            if isinstance(intervalo, tuple) and len(intervalo) == 2:
                ini, fim = intervalo

                df_base_filtro_data = df_base_filtro_data[
                    (df_base_filtro_data["Dt Entrega"].notna()) &
                    (df_base_filtro_data["Dt Entrega"] >= pd.to_datetime(ini)) &
                    (df_base_filtro_data["Dt Entrega"] <= pd.to_datetime(fim) + pd.Timedelta(days=1))
                ]

            # ==========================
            # BASE DE ATRASOS (FILTROS)
            # ==========================
            df_atrasos_unidade = df_base_filtro_data[
                (df_base_filtro_data["SLA"].str.upper() == "ATRASADO") &
                (~df_base_filtro_data["Ocorrencias"].isin([
                    "Dest. Ausente",
                    "Prob. Endereço",
                    "Middle Mile"
                ]))
            ].copy()


            if df_atrasos_unidade.empty:
                st.info("Não há atrasos não justificados para exibir.")

            else:

                st.subheader("🏭 Unidades ofensoras")

                # ==========================
                # TRATA OCORRÊNCIAS VAZIAS
                # ==========================
                df_atrasos_unidade["Ocorrencias"] = (
                    df_atrasos_unidade["Ocorrencias"]
                    .fillna("Sem ocorrência")
                    .replace("", "Sem ocorrência")
                )

                # ==========================
                # OCORRÊNCIAS POR UNIDADE
                # ==========================
                ocorrencias_unidade = (
                    df_atrasos_unidade
                    .groupby(["Destino", "Ocorrencias"])
                    .size()
                    .reset_index(name="Qtd")
                )

                ocorrencias_formatadas = (
                    ocorrencias_unidade
                    .sort_values(["Destino", "Qtd"], ascending=[True, False])
                    .groupby("Destino")
                    .apply(
                        lambda x: ", ".join(
                            f"{row.Ocorrencias} ({row.Qtd})"
                            for _, row in x.iterrows()
                        )
                    )
                    .reset_index(name="Ocorrências")
                )

                # ==========================
                # QUANTIDADE POR UNIDADE
                # ==========================
                qtd_por_unidade = (
                    df_atrasos_unidade
                    .groupby("Destino")
                    .size()
                    .reset_index(name="Quantidade")
                )

                total_atrasos = qtd_por_unidade["Quantidade"].sum()

                qtd_por_unidade["Percentual (%)"] = (
                    qtd_por_unidade["Quantidade"] / total_atrasos * 100
                ).round(2)

                # ==========================
                # JUNTA TUDO
                # ==========================
                tabela_final = (
                    qtd_por_unidade
                    .merge(ocorrencias_formatadas, on="Destino", how="left")
                    .rename(columns={"Destino": "Unidade"})
                    .sort_values("Quantidade", ascending=False)
                )

                # ==========================
                # LINHA TOTAL FIXA
                # ==========================
                linha_total = pd.DataFrame([{
                    "Unidade": "TOTAL",
                    "Ocorrências": "—",
                    "Quantidade": total_atrasos,
                    "Percentual (%)": 100.0
                }])

                tabela_final = pd.concat(
                    [tabela_final, linha_total],
                    ignore_index=True
                )

                # ==========================
                # ORDEM FINAL DAS COLUNAS
                # ==========================
                tabela_final = tabela_final[
                    ["Unidade", "Ocorrências", "Quantidade", "Percentual (%)"]
                ]

                # ==========================
                # EXIBE
                # ==========================
                st.dataframe(
                    tabela_final,
                    use_container_width=True,
                    hide_index=True
                )

        else:
            st.info("Não há dados de performance para o período selecionado.")

else:
    st.info("Importe planilhas no menu lateral para visualizar o dashboard.")
