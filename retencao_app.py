import streamlit as st
import pandas as pd
from datetime import date, datetime
import json
import os

# --- VALORES DE CONFIGURAÇÃO PADRÃO ---
DEFAULT_USUARIOS_OFICIAIS = ['DEJESF5', 'EDUARM11', 'LEMESAM', 'MARTIE90', 'CHRISA13', 'SILVAJ49', 'AFONSS1', 'LARAQA', 'ALVESM30', 'VITORJ11']
DEFAULT_USUARIOS_BACKUP = ['HENRIM12', 'BARBOC20', 'ROBERE16', 'CAROLA12', 'FERNAM40']
DEFAULT_USUARIOS_STAFF = ['OLIVEA34', 'SSILVA']

DEFAULT_RETENTION_BANDS = [
    (0.00, 0.55, 27.59),
    (0.5501, 0.59, 31.58),
    (0.5901, 0.65, 36.12),
    (0.6501, 1.00, 42.00)
]
# --- FIM DOS VALORES DE CONFIGURAÇÃO PADRÃO ---

# Caminhos relativos para os arquivos dentro do repositório
CONFIG_FILE = "config.json"
EXCEL_FILE_PATH = "Retenção - Macro.xlsx"

MOTIVOS_A_DESCONSIDERAR_PADRAO = ["FALECIMENTO DO TITULAR", "AQUISIÇÃO DE BBLEND"]


# Function to load configuration (adapted for Streamlit)
def load_config():
    try:
        with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
            config_data = json.load(f)
            usuarios_oficiais = config_data.get('usuarios_oficiais', DEFAULT_USUARIOS_OFICIAIS)
            usuarios_backup = config_data.get('usuarios_backup', DEFAULT_USUARIOS_BACKUP)
            usuarios_staff = config_data.get('usuarios_staff', DEFAULT_USUARIOS_STAFF)
            loaded_bands = config_data.get('retention_bands', DEFAULT_RETENTION_BANDS)
            retention_bands = [(float(b[0]), float(b[1]), float(b[2])) for b in loaded_bands]
            return usuarios_oficiais, usuarios_backup, usuarios_staff, retention_bands
    except FileNotFoundError:
        return list(DEFAULT_USUARIOS_OFICIAIS), list(DEFAULT_USUARIOS_BACKUP), \
               list(DEFAULT_USUARIOS_STAFF), [list(band) for band in DEFAULT_RETENTION_BANDS]
    except json.JSONDecodeError:
        st.warning("Arquivo de configuração corrompido ou inválido. Usando valores padrão.")
        return list(DEFAULT_USUARIOS_OFICIAIS), list(DEFAULT_USUARIOS_BACKUP), \
               list(DEFAULT_USUARIOS_STAFF), [list(band) for band in DEFAULT_RETENTION_BANDS]

# Function to save configuration (adapted for Streamlit)
def save_config(usuarios_oficiais, usuarios_backup, usuarios_staff, retention_bands):
    serializable_retention_bands = [list(band) for band in retention_bands]
    config_data = {
        'usuarios_oficiais': usuarios_oficiais,
        'usuarios_backup': usuarios_backup,
        'usuarios_staff': usuarios_staff,
        'retention_bands': serializable_retention_bands
    }
    try:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config_data, f, indent=4)
        st.success("Configurações salvas com sucesso!")
    except Exception as e:
        st.error(f"Não foi possível salvar as configurações: {e}")

# Data processing functions (minimal changes needed)
def process_data(df, usuarios_oficiais, usuarios_backup, usuarios_staff):
    df.columns = df.columns.str.strip()
    df = df.rename(columns={
        df.columns[1]: "Login",
        df.columns[7]: "Status",
        df.columns[8]: "DataCriacao",
        df.columns[16]: "Categoria2Motivo",
        df.columns[4]: "TipoRetido",
        df.columns[11]: "Franquia"
    })

    df["Login"] = df["Login"].astype(str).str.strip().str.upper()
    df["DataCriacao"] = pd.to_datetime(df["DataCriacao"], errors='coerce').dt.date
    df = df.dropna(subset=["DataCriacao"])

    if "Categoria2Motivo" not in df.columns:
        df["Categoria2Motivo"] = ""
    else:
        df["Categoria2Motivo"] = df["Categoria2Motivo"].astype(str).str.strip().str.upper()

    if "TipoRetido" not in df.columns:
        df["TipoRetido"] = ""
    else:
        df["TipoRetido"] = df["TipoRetido"].astype(str).str.strip()

    if "Franquia" not in df.columns:
        df["Franquia"] = ""
    else:
        df["Franquia"] = df["Franquia"].astype(str).str.strip().str.upper()

    def classificar_grupo(login):
        if login in usuarios_oficiais:
            return "Retenção"
        elif login in usuarios_backup:
            return "Backup"
        elif login in usuarios_staff:
            return "Supervisão"
        else:
            return "Demais Operações"

    df["Operacao"] = df["Login"].apply(classificar_grupo)
    df.rename(columns={"Operacao": "Operação"}, inplace=True)
    return df

def _get_value_for_conversion_rate(conversion_rate, retention_bands):
    conversion_rate_decimal = conversion_rate / 100.0
    retention_band_names = [
        "0,00% a 55,00%",
        "55,01% a 59,00%",
        "59,01% a 65,00%",
        "65,01% a Acima"
    ]
    for i, (lower, upper, value) in enumerate(retention_bands):
        if lower <= conversion_rate_decimal <= (upper + 1e-9):
            return value, retention_band_names[i]
    return 0.00, "N/A"

def calcular_resumo_retencao(df_filtrado, retention_bands):
    all_dates = sorted(df_filtrado["DataCriacao"].unique())
    date_headers = [d.strftime("%d-%b").lower() for d in all_dates]

    col_names = ["Métrica"] + date_headers + ["Consolidado"]

    agrupado_por_data_status = df_filtrado.groupby(["DataCriacao", "Status"]).size().unstack(fill_value=0).reset_index()

    retido_diario = {d: 0 for d in all_dates}
    nao_retido_diario = {d: 0 for d in all_dates}

    for _, row in agrupado_por_data_status.iterrows():
        data = row["DataCriacao"]
        retido_diario[data] = row.get("Retido", 0)
        nao_retido_diario[data] = row.get("Não Retido", 0)

    total_retido_geral_abs = sum(retido_diario.values())
    total_nao_retido_geral_abs = sum(nao_retido_diario.values())

    intencoes_cancelamento_diario_calculado = {
        d: retido_diario.get(d, 0) + nao_retido_diario.get(d, 0)
        for d in all_dates
    }
    total_intencoes_cancelamento_calculado_geral = sum(intencoes_cancelamento_diario_calculado.values())

    conversao_ecohouse_diaria_values = ["Conversão Ecohouse"]
    for d in all_dates:
        retido_count = retido_diario.get(d, 0)
        nao_retido_count = nao_retido_diario.get(d, 0)
        denominador_conversao_dia = retido_count + nao_retido_count
        if denominador_conversao_dia > 0:
            percent_conversao_dia = (retido_count / denominador_conversao_dia) * 100
            conversao_ecohouse_diaria_values.append(f"{percent_conversao_dia:.2f}%")
        else:
            conversao_ecohouse_diaria_values.append("0.00%")

    denominador_conversao_geral = total_retido_geral_abs + total_nao_retido_geral_abs
    if denominador_conversao_geral > 0:
        percent_conversao_geral_sum = (total_retido_geral_abs / denominador_conversao_geral) * 100
        conversao_ecohouse_diaria_values.append(f"{percent_conversao_geral_sum:.2f}%")
    else:
        conversao_ecohouse_diaria_values.append("0.00%")

    nao_retido_a_desconsiderar_diario = {d: 0 for d in all_dates}
    if "Categoria2Motivo" in df_filtrado.columns:
        df_nao_retido_excluir = df_filtrado[(df_filtrado["Status"] == "Não Retido") &
                                             (df_filtrado["Categoria2Motivo"].isin(MOTIVOS_A_DESCONSIDERAR_PADRAO))]
        nao_retido_excluido_por_data = df_nao_retido_excluir.groupby("DataCriacao").size()
        for d in all_dates:
            nao_retido_a_desconsiderar_diario[d] = nao_retido_excluido_por_data.get(d, 0)

    conversao_faturamento_values = ["Conversão Faturamento"]
    faturamento_percent_diario = {}
    for d in all_dates:
        retido_count = retido_diario.get(d, 0)
        nao_retido_total_dia = nao_retido_diario.get(d, 0)
        nao_retido_excluido_dia = nao_retido_a_desconsiderar_diario.get(d, 0)

        nao_retido_ajustado_dia = max(0, nao_retido_total_dia - nao_retido_excluido_dia)

        denominador_faturamento_dia = retido_count + nao_retido_ajustado_dia
        if denominador_faturamento_dia > 0:
            percent_faturamento_dia = (retido_count / denominador_faturamento_dia) * 100
            conversao_faturamento_values.append(f"{percent_faturamento_dia:.2f}%")
            faturamento_percent_diario[d] = percent_faturamento_dia
        else:
            conversao_faturamento_values.append("0.00%")
            faturamento_percent_diario[d] = 0.00

    total_nao_retido_geral_ajustado = max(0, total_nao_retido_geral_abs - sum(nao_retido_a_desconsiderar_diario.values()))
    denominador_faturamento_geral = total_retido_geral_abs + total_nao_retido_geral_ajustado
    if denominador_faturamento_geral > 0:
        percent_faturamento_geral_sum = (total_retido_geral_abs / denominador_faturamento_geral) * 100
        consolidado_faturamento_percent = percent_faturamento_geral_sum
    else:
        consolidado_faturamento_percent = 0.00

    consolidated_value_per_intent, consolidated_band_name = _get_value_for_conversion_rate(consolidado_faturamento_percent, retention_bands)
    intencoes_cancelamento_ajustado_diario = {
        d: intencoes_cancelamento_diario_calculado.get(d, 0) - nao_retido_a_desconsiderar_diario.get(d, 0)
        for d in all_dates
    }
    total_intencoes_cancelamento_ajustado_geral = max(0, sum(intencoes_cancelamento_ajustado_diario.values()))
    final_total_payment_consolidado = total_intencoes_cancelamento_ajustado_geral * consolidated_value_per_intent

    retido_row_values = ["Retido"] + [retido_diario.get(d, 0) for d in all_dates] + [total_retido_geral_abs]
    nao_retido_row_values = ["Não Retido"] + [nao_retido_diario.get(d, 0) for d in all_dates] + [total_nao_retido_geral_abs]
    intencoes_cancelamento_row_values = ["Intenções de Cancelamento"] + [intencoes_cancelamento_diario_calculado.get(d, 0) for d in all_dates] + [total_intencoes_cancelamento_calculado_geral]

    data_for_export = []
    data_for_export.append(conversao_ecohouse_diaria_values)
    data_for_export.append(conversao_faturamento_values)
    data_for_export.append(retido_row_values)
    data_for_export.append(nao_retido_row_values)
    data_for_export.append(intencoes_cancelamento_row_values)

    df_resumo = pd.DataFrame(data_for_export, columns=col_names)

    # Retornar também os valores consolidados brutos para os indicadores de performance
    return df_resumo, final_total_payment_consolidado, consolidated_band_name, \
           total_retido_geral_abs, total_nao_retido_geral_abs, total_intencoes_cancelamento_calculado_geral, \
           f"{consolidado_faturamento_percent:.2f}%", f"{percent_conversao_geral_sum:.2f}%"

def calcular_detalhe_por_status(df_filtrado, status_filter):
    df_filtered = df_filtrado[df_filtrado["Status"] == status_filter]
    all_dates_detalhe = sorted(df_filtrado["DataCriacao"].unique())

    agrupado = df_filtered.groupby(["Operação", "Login", "DataCriacao"]).size().unstack(fill_value=0)

    if isinstance(agrupado.index, pd.MultiIndex):
        detalhe_para_exibir = agrupado.reset_index()
    else:
        detalhe_para_exibir = agrupado.copy()

    for dt in all_dates_detalhe:
        if dt not in detalhe_para_exibir.columns:
            detalhe_para_exibir[dt] = 0

    detalhe_para_exibir.columns = [col if not isinstance(col, datetime) else col.date() for col in detalhe_para_exibir.columns]
    cols_ordered = ["Operação", "Login"] + sorted([col for col in all_dates_detalhe])
    detalhe_para_exibir = detalhe_para_exibir[cols_ordered]

    detalhe_para_exibir["Consolidado"] = detalhe_para_exibir[[col for col in all_dates_detalhe]].sum(axis=1)

    soma_por_dia = df_filtered.groupby("DataCriacao").size()
    soma_detalhe_row_values = ["", "Consolidado Dia"]
    for dt in all_dates_detalhe:
        soma_detalhe_row_values.append(soma_por_dia.get(dt, 0))
    soma_detalhe_row_values.append(sum(soma_por_dia.values))
    
    soma_df_row = pd.DataFrame([soma_detalhe_row_values], columns=["Operação", "Login"] + all_dates_detalhe + ["Consolidado"])
    df_final = pd.concat([detalhe_para_exibir, soma_df_row], ignore_index=True)
    
    # Format date columns to string
    df_final = df_final.rename(columns={col: col.strftime("%d-%b").lower() for col in all_dates_detalhe})

    return df_final

def calcular_conversao_por_usuario(df_filtrado):
    agrupado_usuario_dia_status = df_filtrado.groupby(["Operação", "Login", "DataCriacao", "Status"]).size().unstack(fill_value=0)
    unique_operacoes_logins = df_filtrado[["Operação", "Login"]].drop_duplicates().sort_values(by=["Operação", "Login"]).values
    all_dates_conversao = sorted(df_filtrado["DataCriacao"].unique())
    date_headers_conversao = [d.strftime("%d-%b").lower() for d in all_dates_conversao]
    colunas_conversao = ["Operação", "Login"] + date_headers_conversao + ["Consolidado"]
    data_for_display_export = []

    for operacao, login in unique_operacoes_logins:
        user_row_values = [operacao, login]
        user_data_for_conversion = agrupado_usuario_dia_status.loc[(operacao, login), :] if (operacao, login) in agrupado_usuario_dia_status.index else pd.DataFrame(columns=['Retido', 'Não Retido'])

        total_retido_user = user_data_for_conversion.get("Retido", pd.Series([0])).sum()
        total_nao_retido_user = user_data_for_conversion.get("Não Retido", pd.Series([0])).sum()

        for d in all_dates_conversao:
            retido_day = user_data_for_conversion.loc[d, "Retido"] if d in user_data_for_conversion.index and "Retido" in user_data_for_conversion.columns else 0
            nao_retido_day = user_data_for_conversion.loc[d, "Não Retido"] if d in user_data_for_conversion.index and "Não Retido" in user_data_for_conversion.columns else 0

            denominador_day = retido_day + nao_retido_day
            if denominador_day > 0:
                percent_day = (retido_day / denominador_day) * 100
                user_row_values.append(f"{percent_day:.2f}%")
            else:
                user_row_values.append("-")

        denominador_consolidado_user = total_retido_user + total_nao_retido_user
        if denominador_consolidado_user > 0:
            consolidado_percent_user = (total_retido_user / denominador_consolidado_user) * 100
            user_row_values.append(f"{consolidado_percent_user:.2f}%")
        else:
            user_row_values.append("-")

        data_for_display_export.append(user_row_values)
    
    df_conversao = pd.DataFrame(data_for_display_export, columns=colunas_conversao)
    return df_conversao

def calcular_motivos_cancelamento(df_filtrado):
    colunas_display = ["Motivo", "Quantidade", "Percentual"]
    if "Categoria2Motivo" not in df_filtrado.columns:
        return pd.DataFrame([["Coluna 'Categoria 2' não encontrada.", "-", "-"]], columns=colunas_display)

    df_nao_retido = df_filtrado[df_filtrado["Status"] == "Não Retido"]
    if df_nao_retido.empty:
        return pd.DataFrame([["Nenhum 'Não Retido' encontrado.", "-", "-"]], columns=colunas_display)

    motivos_contagem = df_nao_retido["Categoria2Motivo"].value_counts().reset_index()
    motivos_contagem.columns = ["Motivo", "Quantidade"]
    total_nao_retido_motivos = motivos_contagem["Quantidade"].sum()
    motivos_contagem["Percentual"] = (motivos_contagem["Quantidade"] / total_nao_retido_motivos) * 100
    motivos_contagem["Percentual"] = motivos_contagem["Percentual"].map("{:.2f}%".format)
    
    # Add total row
    total_row = pd.DataFrame([["Total", total_nao_retido_motivos, "100.00%"]], columns=colunas_display)
    df_motivos = pd.concat([motivos_contagem, total_row], ignore_index=True)
    return df_motivos

def calcular_tipos_retido(df_filtrado):
    colunas_display = ["Tipo de Retido", "Quantidade", "Percentual"]
    if "TipoRetido" not in df_filtrado.columns:
        return pd.DataFrame([["Coluna 'Tipo de Retido' não encontrada.", "-", "-"]], columns=colunas_display)

    df_retido_filtrado_tipo = df_filtrado[(df_filtrado["Status"] == "Retido") &
                                          (df_filtrado["TipoRetido"].astype(str).str.startswith("Retido"))]
    if df_retido_filtrado_tipo.empty:
        return pd.DataFrame([["Nenhum 'Retido' com tipo específico encontrado.", "-", "-"]], columns=colunas_display)

    tipos_retido_contagem = df_retido_filtrado_tipo["TipoRetido"].value_counts().reset_index()
    tipos_retido_contagem.columns = ["Tipo de Retido", "Quantidade"]
    total_retido_tipo = tipos_retido_contagem["Quantidade"].sum()
    tipos_retido_contagem["Percentual"] = (tipos_retido_contagem["Quantidade"] / total_retido_tipo) * 100
    tipos_retido_contagem["Percentual"] = tipos_retido_contagem["Percentual"].map("{:.2f}%".format)

    # Add total row
    total_row = pd.DataFrame([["Total", total_retido_tipo, "100.00%"]], columns=colunas_display)
    df_tipos = pd.concat([tipos_retido_contagem, total_row], ignore_index=True)
    return df_tipos

def calcular_franquias_nao_retido(df_filtrado):
    colunas_display = ["Franquia", "Quantidade", "Percentual"]
    if "Franquia" not in df_filtrado.columns:
        return pd.DataFrame([["Coluna 'Franquias' não encontrada.", "-", "-"]], columns=colunas_display)

    df_nao_retido_franquia = df_filtrado[df_filtrado["Status"] == "Não Retido"]
    if df_nao_retido_franquia.empty:
        return pd.DataFrame([["Nenhum 'Não Retido' por franquia encontrado.", "-", "-"]], columns=colunas_display)

    franquias_contagem = df_nao_retido_franquia["Franquia"].value_counts().reset_index()
    franquias_contagem.columns = ["Franquia", "Quantidade"]
    total_franquias = franquias_contagem["Quantidade"].sum()
    franquias_contagem["Percentual"] = (franquias_contagem["Quantidade"] / total_franquias) * 100
    franquias_contagem["Percentual"] = franquias_contagem["Percentual"].map("{:.2f}%".format)

    # Add total row
    total_row = pd.DataFrame([["Total", total_franquias, "100.00%"]], columns=colunas_display)
    df_franquias = pd.concat([franquias_contagem, total_row], ignore_index=True)
    return df_franquias


# Streamlit App
def main():
    st.set_page_config(layout="wide", page_title="Acompanhamento Retenção")

    # Load initial configuration
    if 'usuarios_oficiais' not in st.session_state:
        st.session_state.usuarios_oficiais, \
        st.session_state.usuarios_backup, \
        st.session_state.usuarios_staff, \
        st.session_state.retention_bands = load_config()

    st.title("📊 Acompanhamento de Retenção")

    # Sidebar for configuration
    with st.sidebar:
        st.header("Configurações")
        
        st.subheader("Configurar Grupos de Usuários")
        with st.expander("Oficiais"):
            oficiais_input = st.text_area("Usuários Oficiais (um por linha)", "\n".join(st.session_state.usuarios_oficiais))
        with st.expander("Backup"):
            backup_input = st.text_area("Usuários Backup (um por linha)", "\n".join(st.session_state.usuarios_backup))
        with st.expander("Staff"):
            staff_input = st.text_area("Usuários Staff (um por linha)", "\n".join(st.session_state.usuarios_staff))

        if st.button("Salvar Configurações de Usuários"):
            st.session_state.usuarios_oficiais = [x.strip().upper() for x in oficiais_input.splitlines() if x.strip()]
            st.session_state.usuarios_backup = [x.strip().upper() for x in backup_input.splitlines() if x.strip()]
            st.session_state.usuarios_staff = [x.strip().upper() for x in staff_input.splitlines() if x.strip()]
            save_config(st.session_state.usuarios_oficiais, st.session_state.usuarios_backup, st.session_state.usuarios_staff, st.session_state.retention_bands)

        st.subheader("Configurar Faixas de Conversão")
        new_retention_bands = []
        for i, (lower, upper, value) in enumerate(st.session_state.retention_bands):
            band_name = ""
            if i == 0: band_name = "0,00% a 55,00%"
            elif i == 1: band_name = "55,01% a 59,00%"
            elif i == 2: band_name = "59,01% a 65,00%"
            elif i == 3: band_name = "65,01% a Acima"

            st.write(f"**Faixa {i+1}: {lower*100:.2f}% a {upper*100:.2f}% (Nome: {band_name})**")
            col1, col2 = st.columns(2)
            with col1:
                new_value = st.number_input(f"Valor R$ para Faixa {i+1}", value=float(value), format="%.2f", key=f"band_value_{i}")
            new_retention_bands.append((lower, upper, new_value))

        if st.button("Salvar Faixas de Conversão"):
            st.session_state.retention_bands = new_retention_bands
            save_config(st.session_state.usuarios_oficiais, st.session_state.usuarios_backup, st.session_state.usuarios_staff, st.session_state.retention_bands)

    # Carregar o arquivo Excel diretamente do caminho fixo
    try:
        if not os.path.exists(EXCEL_FILE_PATH):
            st.error(f"Erro: O arquivo não foi encontrado no caminho especificado: `{EXCEL_FILE_PATH}`")
            st.info("Por favor, verifique se o arquivo 'Retenção - Macro.xlsx' está na mesma pasta do script no repositório.")
            st.session_state.df_original = None
        else:
            df_original = pd.read_excel(EXCEL_FILE_PATH)
            st.session_state.df_original = process_data(df_original.copy(), st.session_state.usuarios_oficiais, st.session_state.usuarios_backup, st.session_state.usuarios_staff)
            
            # Obter e exibir a data de última atualização do arquivo
            last_modified_timestamp = os.path.getmtime(EXCEL_FILE_PATH)
            last_modified_datetime = datetime.fromtimestamp(last_modified_timestamp)
            st.markdown(f"**Última atualização:** {last_modified_datetime.strftime('%d/%m/%Y às %Hh%M')}")

    except Exception as e:
        st.error(f"Erro ao carregar ou processar o arquivo: {e}")
        st.session_state.df_original = None

    if 'df_original' in st.session_state and st.session_state.df_original is not None:
        st.subheader("Filtrar Grupos de Usuários para Análise")
        col_oficiais, col_backup, col_staff, col_outros = st.columns(4)
        with col_oficiais:
            chk_oficiais = st.checkbox("Oficiais", value=True, key="chk_oficiais")
        with col_backup:
            chk_backup = st.checkbox("Backup", value=True, key="chk_backup")
        with col_staff:
            chk_staff = st.checkbox("Staff", value=True, key="chk_staff")
        with col_outros:
            chk_outros = st.checkbox("Demais Operações", value=True, key="chk_outros")

        grupos_selecionados = []
        if chk_oficiais: grupos_selecionados.append("Retenção")
        if chk_backup: grupos_selecionados.append("Backup")
        if chk_staff: grupos_selecionados.append("Supervisão")
        if chk_outros: grupos_selecionados.append("Demais Operações")

        df_filtrado = st.session_state.df_original[st.session_state.df_original["Operação"].isin(grupos_selecionados)].copy()

        if df_filtrado.empty:
            st.warning("Nenhum dado encontrado para os filtros selecionados. Ajuste os filtros de usuários ou verifique o arquivo.")
            return

        # --- Seção de Indicadores de Performance ---
        st.markdown("---")
        st.header("Indicadores de Performance")

        # Recalcular resumo para pegar os valores consolidados
        df_resumo, valor_fatura, faixa_faturamento, \
        total_retido_geral_abs, total_nao_retido_geral_abs, total_intencoes_cancelamento_calculado_geral, \
        percent_faturamento_geral_sum_str, percent_conversao_geral_sum_str = calcular_resumo_retencao(df_filtrado, st.session_state.retention_bands)

        col_kpi1, col_kpi2, col_kpi3, col_kpi4, col_kpi5 = st.columns(5)
        with col_kpi1:
            st.metric(label="Retidos", value=total_retido_geral_abs)
        with col_kpi2:
            st.metric(label="Não Retidos", value=total_nao_retido_geral_abs)
        with col_kpi3:
            st.metric(label="Intenções de Cancelamento", value=total_intencoes_cancelamento_calculado_geral)
        with col_kpi4:
            st.metric(label="Conversão Faturamento", value=percent_faturamento_geral_sum_str)
        with col_kpi5:
            st.metric(label="Conversão Ecohouse", value=percent_conversao_geral_sum_str)

        st.markdown("---")
        # --- Fim da Seção de Indicadores de Performance ---

        st.subheader("Resumo de Retenção")
        col_metric1, col_metric2 = st.columns(2)
        with col_metric1:
            st.metric("Valor Fatura Estimado", f"R$ {valor_fatura:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")) # Format to Brazilian currency
        with col_metric2:
            st.metric("Faixa de Faturamento", faixa_faturamento)
        st.dataframe(df_resumo, hide_index=True)

        st.markdown("---")

        # Create tabs for different analytical views
        tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
            "Não Retidos",
            "Retidos",
            "Conversão por Usuário",
            "Motivos de Cancelamento",
            "Tipos de Retido",
            "Franquias (Não Retido)"
        ])

        with tab1:
            st.subheader("Detalhes de Não Retidos")
            df_nao_retido = calcular_detalhe_por_status(df_filtrado, "Não Retido")
            st.dataframe(df_nao_retido, hide_index=True)

        with tab2:
            st.subheader("Detalhes de Retidos")
            df_retido = calcular_detalhe_por_status(df_filtrado, "Retido")
            st.dataframe(df_retido, hide_index=True)

        with tab3:
            st.subheader("Conversão por Usuário")
            df_conversao_usuario = calcular_conversao_por_usuario(df_filtrado)
            st.dataframe(df_conversao_usuario, hide_index=True)

        with tab4:
            st.subheader("Motivos de Cancelamento (Não Retidos)")
            df_motivos = calcular_motivos_cancelamento(df_filtrado)
            st.dataframe(df_motivos, hide_index=True)

        with tab5:
            st.subheader("Tipos de Retido")
            df_tipos_retido = calcular_tipos_retido(df_filtrado)
            st.dataframe(df_tipos_retido, hide_index=True)

        with tab6:
            st.subheader("Franquias (Não Retido)")
            df_franquias_nao_retido = calcular_franquias_nao_retido(df_filtrado)
            st.dataframe(df_franquias_nao_retido, hide_index=True)

        st.markdown("---")
        st.write("Criado por: Pedro Fregulhe")

        # Export functionality
        st.subheader("Exportar Análise Completa")
        if st.button("Gerar Arquivo de Exportação (.xlsx)"):
            # Recalculate all DFs with current filters for export
            df_resumo_export, _, _, _, _, _, _, _ = calcular_resumo_retencao(df_filtrado, st.session_state.retention_bands)
            df_nao_retido_export = calcular_detalhe_por_status(st.session_state.df_original[st.session_state.df_original["Operação"].isin(grupos_selecionados)], "Não Retido")
            df_retido_export = calcular_detalhe_por_status(st.session_state.df_original[st.session_state.df_original["Operação"].isin(grupos_selecionados)], "Retido")
            df_conversao_usuario_export = calcular_conversao_por_usuario(df_filtrado)
            df_motivos_export = calcular_motivos_cancelamento(df_filtrado)
            df_tipos_retido_export = calcular_tipos_retido(df_filtrado)
            df_franquias_nao_retido_export = calcular_franquias_nao_retido(df_filtrado)

            output = pd.ExcelWriter("analise_retencao.xlsx", engine='xlsxwriter')

            if df_resumo_export is not None:
                df_resumo_export.to_excel(output, sheet_name='Resumo de Retenção', index=False)
            if df_nao_retido_export is not None:
                df_nao_retido_export.to_excel(output, sheet_name='Nao Retidos', index=False)
            if df_retido_export is not None:
                df_retido_export.to_excel(output, sheet_name='Retidos', index=False)
            if df_conversao_usuario_export is not None:
                df_conversao_usuario_export.to_excel(output, sheet_name='Conversao', index=False)
            if df_motivos_export is not None:
                # Remove 'Total' row before exporting if it exists
                if 'Motivo' in df_motivos_export.columns and 'Total' in df_motivos_export['Motivo'].values:
                    df_export_motivos = df_motivos_export[df_motivos_export['Motivo'] != 'Total'].copy()
                else:
                    df_export_motivos = df_motivos_export.copy()
                df_export_motivos.to_excel(output, sheet_name='Motivos Cancelamento', index=False)
            if df_tipos_retido_export is not None:
                # Remove 'Total' row before exporting if it exists
                if 'Tipo de Retido' in df_tipos_retido_export.columns and 'Total' in df_tipos_retido_export['Tipo de Retido'].values:
                    df_export_tipos = df_tipos_retido_export[df_tipos_retido_export['Tipo de Retido'] != 'Total'].copy()
                else:
                    df_export_tipos = df_tipos_retido_export.copy()
                df_export_tipos.to_excel(output, sheet_name='Tipos Retido', index=False)
            if df_franquias_nao_retido_export is not None:
                # Remove 'Total' row before exporting if it exists
                if 'Franquia' in df_franquias_nao_retido_export.columns and 'Total' in df_franquias_nao_retido_export['Franquia'].values:
                    df_export_franquias = df_franquias_nao_retido_export[df_franquias_nao_retido_export['Franquia'] != 'Total'].copy()
                else:
                    df_export_franquias = df_franquias_nao_retido_export.copy()
                df_export_franquias.to_excel(output, sheet_name='Franquias Nao Retido', index=False)

            output.close()

            with open("analise_retencao.xlsx", "rb") as file:
                btn = st.download_button(
                    label="Download Excel da Análise",
                    data=file,
                    file_name="analise_retencao.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            st.success("Arquivo 'analise_retencao.xlsx' gerado e pronto para download!")
    else:
        st.info("O aplicativo está pronto. Verifique o caminho do arquivo Excel se houver erros de carregamento.")


if __name__ == "__main__":
    main()
