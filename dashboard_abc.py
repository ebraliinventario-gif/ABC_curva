import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from pathlib import Path
from io import BytesIO

# Configuração da página
st.set_page_config(page_title="Análise Curva ABC", layout="wide")

# Sidebar com instruções
with st.sidebar:
    st.markdown("## 📋 Instruções de Uso")
    st.markdown("""
    **Bem-vindo ao Dashboard ABC!** 
    
    Este aplicativo ajuda você a analisar a curva ABC dos seus produtos.
    
    ### Como usar:
    1. **Upload do Arquivo**: Faça upload de um arquivo Excel (.xlsx) com os dados dos produtos.
    
    2. **Colunas Necessárias**:
       - `descricao`: Nome do produto
       - `KG`: Quantidade em quilogramas
       - `% individual`: Percentual individual
       - `Tipo Item`: Categoria do produto
       - `% acumulado`: Percentual acumulado
    
    3. **Filtros**:
       - Selecione o percentual de faturamento (60%-100%)
       - Escolha o tipo de item ou 'Todos'
       - Selecione o tipo de análise
    
    4. **Visualizações**:
       - Curva ABC (Pareto)
       - Distribuição por tipo
       - Tabela de ranking
       - Distribuição das classes ABC
    
    ### Dicas:
    - Use os filtros para focar em categorias específicas
    - A curva ABC classifica produtos em A (80%), B (15%), C (5%)
    - Produtos A são os mais importantes
    """)

# CSS customizado
st.markdown("""
    <style>
    .main {
        background: linear-gradient(135deg, #0f0f1e 0%, #1a1a2e 100%);
    }

    div.stDownloadButton > button {
        width: 100%;
        background: linear-gradient(135deg, #06d6a0 0%, #1f77b4 100%);
        color: #0f0f1e;
        border: 1px solid rgba(255, 255, 255, 0.15);
        border-radius: 12px;
        padding: 0.7rem 1rem;
        font-weight: 700;
        letter-spacing: 0.2px;
        box-shadow: 0 8px 18px rgba(0, 0, 0, 0.35);
        transition: transform 120ms ease, box-shadow 120ms ease, filter 120ms ease;
    }

    div.stDownloadButton > button:hover {
        filter: brightness(1.05);
        transform: translateY(-1px);
        box-shadow: 0 12px 24px rgba(0, 0, 0, 0.45);
        border-color: rgba(255, 255, 255, 0.22);
    }

    div.stDownloadButton > button:active {
        transform: translateY(0px);
        box-shadow: 0 8px 18px rgba(0, 0, 0, 0.35);
    }

    div.stDownloadButton > button:focus,
    div.stDownloadButton > button:focus-visible {
        outline: none !important;
        box-shadow: 0 0 0 3px rgba(6, 214, 160, 0.25), 0 12px 24px rgba(0, 0, 0, 0.45);
    }

    div.stDownloadButton > button:disabled {
        background: rgba(255, 255, 255, 0.08) !important;
        color: rgba(255, 255, 255, 0.55) !important;
        border-color: rgba(255, 255, 255, 0.10) !important;
        box-shadow: none !important;
        transform: none !important;
        cursor: not-allowed;
    }

    div.stDownloadButton {
        margin-top: 0.25rem;
        margin-bottom: 0.25rem;
    }
    
    [data-testid="stMetric"] {
        background-color: rgba(31, 119, 180, 0.1);
        padding: 20px;
        border-radius: 12px;
        border-left: 4px solid #1f77b4;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.2);
    }
    
    h1, h2, h3 {
        color: #ffffff;
        font-weight: 700;
        letter-spacing: 0.5px;
    }
    
    .stDataFrame {
        background-color: #1c1f26;
    }
    
    .sidebar .sidebar-content {
        background: linear-gradient(180deg, #1a1a2e 0%, #16213e 100%);
    }
    
    .stSidebar {
        background: linear-gradient(180deg, #1a1a2e 0%, #16213e 100%);
        border-right: 2px solid #0f3460;
    }
    
    .stMarkdown {
        color: #e0e0e0;
    }
    </style>
""", unsafe_allow_html=True)

# Definir valor padrão para analysis_type
analysis_type = "Análise ABC"

# Header padrão
col1, col2 = st.columns([1, 4])
with col1:
    st.markdown("## 📊 ABC")
with col2:
    st.markdown(f"## SEGMENTAÇÃO DE PRODUTOS - {analysis_type}")

st.markdown("---")

data_source = st.radio("Fonte dos dados", options=["Planilhas fixas", "Upload"], horizontal=True, index=0)

_base_dir = Path(__file__).resolve().parent
_fixed_files = {
    "ABC PLAN.xlsx": _base_dir / "ABC PLAN.xlsx",
    "Curva ABC (QTD).xlsx": _base_dir / "Curva ABC (QTD).xlsx",
}

excel_source = None
file_name = None

if data_source == "Planilhas fixas":
    fixed_choice = st.selectbox("Selecione a planilha", options=list(_fixed_files.keys()))
    fixed_path = _fixed_files[fixed_choice]
    if not fixed_path.exists():
        st.error(f"❌ Planilha fixa não encontrada: {fixed_path}")
        st.stop()
    excel_source = fixed_path
    file_name = fixed_path.name.lower()
else:
    uploaded_file = st.file_uploader("📤 Faça upload do arquivo Excel", type=['xlsx'])
    if uploaded_file is not None:
        excel_source = uploaded_file
        file_name = uploaded_file.name.lower()

if excel_source is not None:
    # Detectar o tipo de arquivo baseado no nome
    
    if "abc plan" in file_name:
        selected_sheet = "Planilha1"  # Aba padrão do Excel
        analysis_type = "Volume (KG)"
        col_quantidade = "KG"
    elif "curva abc" in file_name and "qtd" in file_name:
        selected_sheet = "Planilha1"  # Aba padrão do Excel
        analysis_type = "Quantidade (QTD)"
        col_quantidade = "Total"
    else:
        st.error("❌ Nome do arquivo não reconhecido. Use 'ABC PLAN.xlsx' ou 'Curva ABC (QTD).xlsx'")
        st.stop()

    try:
        df = pd.read_excel(excel_source, sheet_name=selected_sheet)
    except Exception as e:
        st.error(f"❌ Erro ao ler o arquivo Excel: {str(e)}")
        st.stop()

    df.columns = df.columns.astype(str).str.strip()
    
    # Atualizar header com o tipo de análise detectado
    col1, col2 = st.columns([1, 4])
    with col1:
        st.markdown("## 📊 ABC")
    with col2:
        st.markdown(f"## SEGMENTAÇÃO DE PRODUTOS - {analysis_type}")

    st.markdown("---")
    
    # Nomes das colunas esperadas
    col_descricao = "descricao"
    col_individual = "% individual"
    col_tipo = "Tipo Item"
    col_acumulado = "% acumulado"

    def _find_col(df_: pd.DataFrame, expected: str) -> str | None:
        if expected in df_.columns:
            return expected
        expected_lower = str(expected).strip().lower()
        for c in df_.columns:
            if str(c).strip().lower() == expected_lower:
                return c
        return None

    col_descricao_found = _find_col(df, col_descricao)
    col_quantidade_found = _find_col(df, col_quantidade)
    col_individual_found = _find_col(df, col_individual)
    col_tipo_found = _find_col(df, col_tipo)
    col_acumulado_found = _find_col(df, col_acumulado)
    
    # Validação de colunas baseada no tipo de arquivo
    if "abc plan" in file_name:
        required_cols = [col_descricao, col_quantidade, col_individual, col_tipo, col_acumulado]
    elif "curva abc" in file_name and "qtd" in file_name:
        required_cols = [col_descricao, col_quantidade]
        # Para QTD, se não tiver as colunas calculadas, vamos criá-las
        if col_individual_found is None:
            df[col_individual] = None
            col_individual_found = col_individual
        if col_tipo_found is None:
            df[col_tipo] = "Produto"  # Valor padrão
            col_tipo_found = col_tipo
        if col_acumulado_found is None:
            df[col_acumulado] = None
            col_acumulado_found = col_acumulado
    else:
        required_cols = [col_descricao, col_quantidade, col_individual, col_tipo, col_acumulado]
    
    found_cols = {
        col_descricao: col_descricao_found,
        col_quantidade: col_quantidade_found,
        col_individual: col_individual_found,
        col_tipo: col_tipo_found,
        col_acumulado: col_acumulado_found,
    }
    missing_cols = [col for col in required_cols if found_cols[col] is None]
    
    if missing_cols:
        st.error(f"❌ Colunas não encontradas: {missing_cols}")
        st.write("Colunas encontradas no arquivo:")
        st.write(list(df.columns))
        st.stop()

    col_descricao = col_descricao_found
    col_quantidade = col_quantidade_found
    col_individual = col_individual_found
    col_tipo = col_tipo_found
    col_acumulado = col_acumulado_found

    def _to_number_ptbr(series: pd.Series) -> pd.Series:
        s0 = series.copy()
        if pd.api.types.is_numeric_dtype(s0):
            return pd.to_numeric(s0, errors='coerce')

        s = s0.astype(str).str.strip()
        s = s.replace({"": pd.NA, "nan": pd.NA, "None": pd.NA, "NaN": pd.NA})
        s = s.str.replace("%", "", regex=False)
        s = s.str.replace(" ", "", regex=False)
        s = s.str.replace("\u00a0", "", regex=False)
        s = s.str.replace(r"[^0-9,\.\-]", "", regex=True)

        has_comma = s.str.contains(",", na=False)
        has_dot = s.str.contains(r"\.", na=False)

        # Caso pt-BR típico: 1.234,56 -> 1234.56
        mask_pt = has_comma & has_dot
        s.loc[mask_pt] = s.loc[mask_pt].str.replace(".", "", regex=False).str.replace(",", ".", regex=False)

        # Caso com vírgula apenas: 1234,56 -> 1234.56
        mask_comma_only = has_comma & (~has_dot)
        s.loc[mask_comma_only] = s.loc[mask_comma_only].str.replace(",", ".", regex=False)

        # Caso com ponto apenas: 1234.56 (já OK). Se tiver separador de milhar com vírgula: 1,234.56 -> 1234.56
        mask_dot_only = has_dot & (~has_comma)
        s.loc[mask_dot_only] = s.loc[mask_dot_only].str.replace(",", "", regex=False)

        # Caso sem separador: 1234
        return pd.to_numeric(s, errors='coerce')

    df_initial_count = len(df)
    load_messages = []
    nan_before = {
        col_quantidade: df[col_quantidade].isna().sum(),
        col_individual: df[col_individual].isna().sum(),
        col_acumulado: df[col_acumulado].isna().sum(),
    }

    df[col_individual] = _to_number_ptbr(df[col_individual])
    df[col_acumulado] = _to_number_ptbr(df[col_acumulado])
    df[col_quantidade] = _to_number_ptbr(df[col_quantidade])

    nan_after = {
        col_quantidade: df[col_quantidade].isna().sum(),
        col_individual: df[col_individual].isna().sum(),
        col_acumulado: df[col_acumulado].isna().sum(),
    }

    coerced = {k: nan_after[k] - nan_before[k] for k in nan_after}
    load_messages.append(
        f"🔎 Diagnóstico de conversão (novos NaN após parsing): "
        f"{col_quantidade}={coerced[col_quantidade]}, % individual={coerced[col_individual]}, % acumulado={coerced[col_acumulado]}"
    )

    missing_acum_ratio = float(df[col_acumulado].isna().mean())
    missing_ind_ratio = float(df[col_individual].isna().mean())
    if missing_acum_ratio > 0.1 or missing_ind_ratio > 0.1:
        df_calc = df.dropna(subset=[col_quantidade]).copy()
        total_quantidade_calc = float(df_calc[col_quantidade].sum())
        if total_quantidade_calc > 0:
            df_calc = df_calc.sort_values(by=col_quantidade, ascending=False)
            df_calc[col_individual] = (df_calc[col_quantidade] / total_quantidade_calc) * 100
            df_calc[col_acumulado] = df_calc[col_individual].cumsum()
            df = df_calc
            df_initial_count = len(df)
            load_messages.append(
                f"🧮 Recalculei '% individual' e '% acumulado' a partir de '{col_quantidade}' "
                "(a planilha estava com muitos valores ausentes nessas colunas)."
            )
    
    # REMOVER LINHAS COM DADOS INCOMPLETOS
    df = df.dropna(subset=[col_quantidade, col_individual, col_acumulado])
    df_after_count = len(df)

    load_messages.append(f"Dados carregados: {df_initial_count} linhas → {df_after_count} linhas (removidas {df_initial_count - df_after_count} incompletas)")
    
    # Se %acumulado está em formato decimal (0-1), converter para porcentagem (0-100)
    if df[col_acumulado].max() < 2:
        df[col_acumulado] = df[col_acumulado] * 100
    if df[col_individual].max() < 2:
        df[col_individual] = df[col_individual] * 100

    df_for_class = df.dropna(subset=[col_quantidade]).sort_values(by=col_quantidade, ascending=False).reset_index()
    total_for_class = float(df_for_class[col_quantidade].sum())
    if total_for_class > 0:
        df_for_class['_pct_ind_calc'] = (df_for_class[col_quantidade] / total_for_class) * 100
        df_for_class['_pct_acum_calc'] = df_for_class['_pct_ind_calc'].cumsum()

        idx_a = df_for_class.loc[df_for_class['_pct_acum_calc'] >= 80].index.min()
        idx_b = df_for_class.loc[df_for_class['_pct_acum_calc'] >= 95].index.min()

        if pd.isna(idx_a):
            idx_a = df_for_class.index.max()
        if pd.isna(idx_b):
            idx_b = df_for_class.index.max()

        def _class_from_pos(pos: int) -> str:
            if pos <= idx_a:
                return 'A'
            if pos <= idx_b:
                return 'B'
            return 'C'

        df_for_class['Classificação ABC'] = df_for_class.index.to_series().apply(_class_from_pos)
        class_map = df_for_class.set_index('index')['Classificação ABC']
        df['Classificação ABC'] = df.index.map(class_map)
    else:
        df['Classificação ABC'] = pd.NA
    
    # ===== FILTROS SUPERIORES =====
    col_filter1, col_filter2, col_filter3 = st.columns(3)
    
    with col_filter1:
        st.markdown("### Selecione o percentual de faturamento")
        threshold_options = {
            '60%': 60,
            '70%': 70,
            '80%': 80,
            '90%': 90,
            '100%': 100
        }
        selected_threshold = st.radio("", options=list(threshold_options.keys()), horizontal=True)
        threshold_value = threshold_options[selected_threshold]
    
    with col_filter2:
        if "qtd" in file_name:
            st.markdown("### Tipo de Item")
            st.info("📝 Para análise por quantidade, todos os itens são categorizados como 'Produto'")
            selected_tipo = 'Todos'  # Força 'Todos' para QTD
        else:
            st.markdown("### Selecione o tipo de item")
            tipos = sorted([str(t) for t in df[col_tipo].dropna().unique() if str(t).strip() != ''])
            selected_tipo = st.selectbox("", ['Todos'] + list(tipos), label_visibility="collapsed")
    
    with col_filter3:
        if "qtd" in file_name:
            st.markdown("### Análise selecionada")
            analysis = st.selectbox("", ["Quantidade"], label_visibility="collapsed")
        else:
            st.markdown("### Análise selecionada")
            analysis = st.selectbox("", ["Faturamento", "Volume", "Margem"], label_visibility="collapsed")
    
    # Filtrar dados
    if selected_tipo != 'Todos':
        df_filtered = df[df[col_tipo] == selected_tipo].copy()
    else:
        df_filtered = df.copy()
    
    # Calcular produtos na classe A até o threshold - USANDO DADOS NÃO FILTRADOS
    df_sorted_all = df.sort_values(by=col_acumulado)
    _eps = 1e-9
    if threshold_value >= 100:
        produtos_ate_threshold = int(df[col_quantidade].notna().sum())
        total_quantidade_threshold = df[col_quantidade].sum()
    else:
        produtos_ate_threshold = len(df_sorted_all[df_sorted_all[col_acumulado] <= (threshold_value + _eps)])
        total_quantidade_threshold = df_sorted_all[df_sorted_all[col_acumulado] <= (threshold_value + _eps)][col_quantidade].sum()
    total_quantidade_all = df[col_quantidade].sum()
    
    st.markdown("---")
    
    # ===== KPIs PRINCIPAIS =====
    metric_col1, metric_col2, metric_col3, metric_col4 = st.columns(4)
    
    with metric_col1:
        st.metric(
            label="PRODUTOS",
            value=produtos_ate_threshold,
            delta=f"{(produtos_ate_threshold/len(df)*100):.0f}% do total"
        )
    
    with metric_col2:
        if "qtd" in file_name:
            st.metric(
                label="REPRESENTAM",
                value=f"{selected_threshold}",
                delta="da quantidade total"
            )
        else:
            st.metric(
                label="REPRESENTAM",
                value=f"{selected_threshold}",
                delta="do faturamento"
            )
    
    with metric_col3:
        st.metric(
            label=f"TOTAL {col_quantidade}",
            value=f"{total_quantidade_threshold:,.0f}",
            delta=f"de {total_quantidade_all:,.0f} {col_quantidade} totais"
        )
    
    with metric_col4:
        st.metric(
            label="CLASSES ABC",
            value=f"{len(df[df['Classificação ABC']=='A'])} / {len(df[df['Classificação ABC']=='B'])} / {len(df[df['Classificação ABC']=='C'])}",
            delta="A / B / C (Total)"
        )
    
    st.markdown("---")
    
    # ===== GRÁFICOS PRINCIPAIS =====
    if "qtd" in file_name:
        charts_ctrl1, charts_ctrl2 = st.columns([2, 3])
        with charts_ctrl1:
            qtd_classes = st.multiselect(
                "Filtrar por Classes ABC:",
                ["A", "B", "C"],
                default=["A", "B", "C"],
                key="qtd_classes_filter",
            )
        with charts_ctrl2:
            st.markdown(" ")
    else:
        pareto_col1, pareto_col2 = st.columns([2, 3])
        with pareto_col1:
            pareto_view = st.selectbox(
                "",
                ["Completo", "Top N por Classe", "Top N (Geral)"],
                label_visibility="collapsed",
            )
        with pareto_col2:
            if pareto_view == "Top N por Classe":
                pareto_classes = st.multiselect(
                    "",
                    ["A", "B", "C"],
                    default=["A"],
                    label_visibility="collapsed",
                )
                pareto_top_n = st.slider(
                    "",
                    min_value=5,
                    max_value=50,
                    value=10,
                    step=5,
                    label_visibility="collapsed",
                )
            elif pareto_view == "Top N (Geral)":
                pareto_classes = []
                pareto_top_n = st.slider(
                    "",
                    min_value=5,
                    max_value=100,
                    value=20,
                    step=5,
                    label_visibility="collapsed",
                )
            else:
                pareto_classes = []
                pareto_top_n = 0
    col_graph1, col_graph2 = st.columns(2)
    
    with col_graph1:
        if "qtd" in file_name:
            st.markdown("### 📊 TOTAIS POR PRODUTO")
            
            # Filtrar dados baseado nas classes selecionadas
            if qtd_classes:
                df_qtd_filtered = df_filtered[df_filtered['Classificação ABC'].isin(qtd_classes)]
            else:
                df_qtd_filtered = df_filtered.iloc[0:0]  # DataFrame vazio se nenhuma classe selecionada
            
            # Gráfico de barras dos top produtos por quantidade
            produto_totals = df_qtd_filtered.groupby(col_descricao)[col_quantidade].sum().sort_values(ascending=False).head(20)
            
            fig_produto = go.Figure(data=[
                go.Bar(
                    y=produto_totals.index,
                    x=produto_totals.values,
                    orientation='h',
                    marker=dict(
                        color=produto_totals.values,
                        colorscale=[[0, '#073b4c'], [0.5, '#118ab2'], [1, '#06d6a0']],
                        line=dict(color='rgba(255,255,255,0.3)', width=1)
                    ),
                    text=[f'{v:,.0f}' for v in produto_totals.values],
                    textposition='outside',
                    textfont=dict(size=10, color='#ffffff'),
                    hovertemplate='<b>%{y}</b><br>Total: %{x:,.0f}<extra></extra>'
                )
            ])
            
            fig_produto.update_layout(
                title=dict(text='Top 20 Produtos por Quantidade Total', font=dict(size=16, color='#ffffff', family='Arial Black')),
                xaxis_title=dict(text='Quantidade Total', font=dict(size=12, color='#cccccc')),
                yaxis_title='',
                plot_bgcolor='rgba(20, 20, 40, 0.5)',
                paper_bgcolor='rgba(15, 15, 30, 0.9)',
                font=dict(color='#ffffff', size=11, family='Arial'),
                height=550,
                showlegend=False,
                xaxis=dict(
                    showgrid=True, 
                    gridwidth=1, 
                    gridcolor='rgba(100,100,100,0.2)',
                    tickfont=dict(color='#cccccc', size=11)
                ),
                yaxis=dict(
                    tickfont=dict(color='#cccccc', size=11),
                    autorange='reversed'  # Para mostrar o maior no topo
                ),
                margin=dict(l=200, r=80, t=80, b=80)
            )
            
            st.plotly_chart(fig_produto, use_container_width=True)
        else:
            st.markdown("### 📊 CURVA ABC")
            
            df_plot_base = df_filtered.copy()
            if pareto_view == "Top N por Classe":
                if pareto_classes:
                    parts = []
                    for cls in pareto_classes:
                        part = df_plot_base[df_plot_base['Classificação ABC'] == cls]
                        part = part.dropna(subset=[col_quantidade]).nlargest(pareto_top_n, col_quantidade)
                        parts.append(part)
                    if parts:
                        df_plot_base = pd.concat(parts, axis=0).drop_duplicates()
                    else:
                        df_plot_base = df_plot_base.iloc[0:0]
                else:
                    df_plot_base = df_plot_base.iloc[0:0]
            elif pareto_view == "Top N (Geral)":
                df_plot_base = df_plot_base.dropna(subset=[col_quantidade]).nlargest(pareto_top_n, col_quantidade)

            # Gráfico de Pareto
            df_plot = df_plot_base.sort_values(by=col_acumulado).reset_index(drop=True)
            
            fig_pareto = go.Figure()
            
            # Barras
            fig_pareto.add_trace(go.Bar(
                x=df_plot[col_descricao],
                y=df_plot[col_individual],
                name='% Individual',
                marker=dict(
                    color=df_plot['Classificação ABC'].map({'A': '#06d6a0', 'B': '#118ab2', 'C': '#ef476f'}),
                    line=dict(color='rgba(255,255,255,0.3)', width=1)
                ),
                text=[f"{v:.1f}%" for v in df_plot[col_individual]],
                textposition='outside',
                textfont=dict(size=10, color='#ffffff'),
                hovertemplate='<b>%{x}</b><br>% Individual: %{y:.1f}%<extra></extra>'
            ))
            
            # Linha de % acumulado
            fig_pareto.add_trace(go.Scatter(
                x=df_plot[col_descricao],
                y=df_plot[col_acumulado],
                name='% Acumulado',
                yaxis='y2',
                line=dict(color='#ef476f', width=4),
                mode='lines+markers',
                marker=dict(size=8, color='#ef476f', symbol='circle', line=dict(color='white', width=2)),
                hovertemplate='<b>%{x}</b><br>% Acumulado: %{y:.1f}%<extra></extra>',
                fill='tozeroy',
                fillcolor='rgba(239, 71, 111, 0.1)'
            ))
            
            # Layout
            title_text = '📊 Análise de Pareto - Curva ABC'
            
            fig_pareto.update_layout(
                title=dict(text=title_text, font=dict(size=18, color='#ffffff', family='Arial Black')),
                xaxis_title=dict(text='Produtos', font=dict(size=12, color='#cccccc')),
                yaxis=dict(
                    title=dict(text='% Individual', font=dict(color='#118ab2', size=12)),
                    tickfont=dict(color='#cccccc', size=11),
                    showgrid=True,
                    gridwidth=1,
                    gridcolor='rgba(100,100,100,0.2)',
                    zeroline=False
                ),
                yaxis2=dict(
                    title=dict(text='% Acumulado', font=dict(color='#ef476f', size=12)),
                    tickfont=dict(color='#cccccc', size=11),
                    overlaying='y',
                    side='right',
                    range=[0, 110],
                    zeroline=False
                ),
                plot_bgcolor='rgba(20, 20, 40, 0.5)',
                paper_bgcolor='rgba(15, 15, 30, 0.9)',
                font=dict(color='#ffffff', size=11, family='Arial'),
                height=550,
                hovermode='x unified',
                showlegend=True,
                legend=dict(
                    x=0.01, 
                    y=0.99, 
                    bgcolor='rgba(0,0,0,0.5)',
                    bordercolor='rgba(255,255,255,0.2)',
                    borderwidth=1,
                    font=dict(color='#ffffff', size=11)
                ),
                margin=dict(l=80, r=80, t=80, b=80)
            )
            
            st.plotly_chart(fig_pareto, use_container_width=True)
    
    with col_graph2:
        if "qtd" in file_name:
            st.markdown("### 📈 DISTRIBUIÇÃO POR QUANTIDADE")
        else:
            st.markdown("### 📈 DISTRIBUIÇÃO POR TIPO DE ITEM")
        
        # Gráfico de barras por tipo
        tipo_summary = df_filtered.groupby(col_tipo)[col_quantidade].sum().sort_values(ascending=True)
        
        fig_bar = go.Figure(data=[
            go.Bar(
                y=tipo_summary.index,
                x=tipo_summary.values,
                orientation='h',
                marker=dict(
                    color=tipo_summary.values,
                    colorscale=[[0, '#073b4c'], [0.5, '#118ab2'], [1, '#06d6a0']],
                    line=dict(color='rgba(255,255,255,0.3)', width=1)
                ),
                text=[f'{v:,.0f}' for v in tipo_summary.values],
                textposition='outside',
                textfont=dict(size=11, color='#ffffff'),
                hovertemplate=f'<b>%{{y}}</b><br>{col_quantidade}: %{{x:,.0f}}<extra></extra>'
            )
        ])
        
        fig_bar.update_layout(
            title=dict(text=f'{analysis_type} por Tipo', font=dict(size=14, color='#ffffff', family='Arial Black')),
            xaxis_title=dict(text=col_quantidade, font=dict(size=11, color='#cccccc')),
            yaxis_title='',
            plot_bgcolor='rgba(20, 20, 40, 0.5)',
            paper_bgcolor='rgba(15, 15, 30, 0.9)',
            font=dict(color='#ffffff', size=11, family='Arial'),
            height=550,
            showlegend=False,
            xaxis=dict(
                showgrid=True, 
                gridwidth=1, 
                gridcolor='rgba(100,100,100,0.2)',
                tickfont=dict(color='#cccccc', size=11)
            ),
            yaxis=dict(
                tickfont=dict(color='#cccccc', size=11)
            ),
            margin=dict(l=150, r=80, t=80, b=80)
        )
        
        st.plotly_chart(fig_bar, use_container_width=True)
    
    st.markdown("---")
    
    # ===== TABELA DE DETALHAMENTO =====
    if "qtd" in file_name:
        st.markdown("### 📋 DETALHAMENTO - RANKING POR QUANTIDADE")
    else:
        st.markdown("### 📋 DETALHAMENTO PARETO - RANKING")
    
    # Preparar tabela
    df_table_base = df_filtered.copy()
    if 'pareto_view' in locals() and pareto_view != "Completo":
        df_table_base = df_plot_base.copy()
    df_table = df_table_base.sort_values(by=col_acumulado).reset_index(drop=True)
    df_table['Rank'] = range(1, len(df_table) + 1)
    df_table['Rank LV'] = df_table.groupby('Classificação ABC').cumcount() + 1
    
    # Selecionar colunas para exibir
    cols_display = ['Rank', col_descricao, 'Classificação ABC', col_quantidade, col_individual, col_acumulado]
    df_display = df_table[cols_display].copy()
    df_display.columns = ['Rank', 'Produto', 'Classe', col_quantidade, '% Individual', '% Acumulado']
    
    # Formatar valores
    df_display[col_quantidade] = df_display[col_quantidade].apply(lambda x: f'{x:,.0f}')
    df_display['% Individual'] = df_display['% Individual'].apply(lambda x: f'{x:.2f}%')
    df_display['% Acumulado'] = df_display['% Acumulado'].apply(lambda x: f'{x:.2f}%')
    
    st.dataframe(
        df_display,
        use_container_width=True,
        hide_index=True,
        column_config={
            'Rank': st.column_config.NumberColumn(width='small'),
            'Produto': st.column_config.TextColumn(width='large'),
            'Classe': st.column_config.TextColumn(width='small'),
            col_quantidade: st.column_config.TextColumn(width='medium'),
            '% Individual': st.column_config.TextColumn(width='medium'),
            '% Acumulado': st.column_config.TextColumn(width='medium'),
        }
    )
    
    st.markdown("---")
    
    # ===== DISTRIBUIÇÃO ABC =====
    col_dist1, col_dist2 = st.columns(2)
    
    with col_dist1:
        if "qtd" in file_name.lower():
            st.markdown("### 🎯 DISTRIBUIÇÃO CLASSES ABC POR QUANTIDADE")
        else:
            st.markdown("### 🎯 DISTRIBUIÇÃO CLASSES ABC")
        abc_counts = df_filtered['Classificação ABC'].value_counts().sort_index()
        
        # Garantir que todas as classes apareçam
        for classe in ['A', 'B', 'C']:
            if classe not in abc_counts.index:
                abc_counts[classe] = 0
        abc_counts = abc_counts.sort_index()
        
        colors_map = {'A': '#06d6a0', 'B': '#118ab2', 'C': '#ef476f'}
        colors = [colors_map.get(idx, '#666666') for idx in abc_counts.index]
        
        fig_pie = go.Figure(data=[go.Pie(
            labels=['Classe ' + label for label in abc_counts.index],
            values=abc_counts.values,
            hole=0.55,
            marker=dict(colors=colors, line=dict(color='rgba(255,255,255,0.25)', width=2)),
            textinfo='percent',
            textposition='inside',
            insidetextorientation='horizontal',
            textfont=dict(size=16, color='#ffffff', family='Arial Black'),
            hovertemplate='<b>%{label}</b><br>Quantidade: %{value}<br>Percentual: %{percent}<extra></extra>'
        )])

        fig_pie.update_traces(
            sort=False,
            pull=[0.02] * len(abc_counts.index),
        )
        
        fig_pie.update_layout(
            title=dict(text='', font=dict(size=14, color='#ffffff')),
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
            font=dict(color='#ffffff', size=12, family='Arial'),
            height=520,
            showlegend=True,
            margin=dict(l=20, r=20, t=20, b=70),
            legend=dict(
                orientation='h',
                yanchor='bottom',
                y=-0.12,
                xanchor='center',
                x=0.5,
                bgcolor='rgba(0,0,0,0)',
                bordercolor='rgba(255,255,255,0.2)',
                font=dict(color='#ffffff', size=11)
            )
        )
        
        st.plotly_chart(fig_pie, use_container_width=True)
    
    with col_dist2:
        st.markdown("### 📊 RESUMO ANALÍTICO")
        
        # Garantir que todas as classes apareçam no resumo
        classe_a_count = len(df_filtered[df_filtered['Classificação ABC'] == 'A'])
        classe_b_count = len(df_filtered[df_filtered['Classificação ABC'] == 'B'])
        classe_c_count = len(df_filtered[df_filtered['Classificação ABC'] == 'C'])
        
        # Card de estatísticas com melhor formatação
        stats_data = {
            '📦 Total de Produtos': str(len(df_filtered)),
            f'⚖️ Total {col_quantidade}': f"{df_filtered[col_quantidade].sum():,.0f}",
            '🟢 Classe A': str(classe_a_count),
            '🔵 Classe B': str(classe_b_count),
            '🔴 Classe C': str(classe_c_count),
            f'📈 {col_quantidade} Médio': f"{df_filtered[col_quantidade].mean():,.0f}",
            f'⬆️ {col_quantidade} Máximo': f"{df_filtered[col_quantidade].max():,.0f}",
            f'⬇️ {col_quantidade} Mínimo': f"{df_filtered[col_quantidade].min():,.0f}",
        }
        
        for label, value in stats_data.items():
            st.markdown(f"""
            <div style="
                background-color: rgba(31, 119, 180, 0.1);
                padding: 12px 16px;
                border-radius: 8px;
                border-left: 3px solid #06d6a0;
                margin-bottom: 8px;
                font-size: 14px;
            ">
                <span style="color: #cccccc;">{label}:</span> 
                <span style="color: #06d6a0; font-weight: bold; font-size: 16px;">{value}</span>
            </div>
            """, unsafe_allow_html=True)

    st.markdown("---")
    tab_downloads, tab_descricao = st.tabs(["⬇️ Downloads", "📝 Descrição do carregamento"])
    with tab_downloads:
        def _df_to_csv_bytes(df_: pd.DataFrame) -> bytes:
            return df_.to_csv(index=False, sep=";", decimal=",").encode("utf-8-sig")

        base_filename = Path(file_name).stem.replace(" ", "_")

        col_csv_1, col_csv_2 = st.columns(2)
        with col_csv_1:
            st.download_button(
                label="Baixar base tratada (CSV)",
                data=_df_to_csv_bytes(df),
                file_name=f"{base_filename}_base_tratada.csv",
                mime="text/csv",
            )

        with col_csv_2:
            st.download_button(
                label="Baixar base filtrada (CSV)",
                data=_df_to_csv_bytes(df_filtered),
                file_name=f"{base_filename}_base_filtrada.csv",
                mime="text/csv",
            )

        st.markdown("### Planilhas modelo (Excel)")
        col_dl_1, col_dl_2 = st.columns(2)

        def _read_binary(path: Path) -> bytes:
            return path.read_bytes()

        with col_dl_1:
            model_abc_path = _fixed_files.get("ABC PLAN.xlsx")
            if isinstance(model_abc_path, Path) and model_abc_path.exists():
                st.download_button(
                    label="Baixar ABC PLAN.xlsx",
                    data=_read_binary(model_abc_path),
                    file_name="ABC PLAN.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            else:
                st.error("❌ Planilha modelo não encontrada: ABC PLAN.xlsx")

        with col_dl_2:
            model_qtd_path = _fixed_files.get("Curva ABC (QTD).xlsx")
            if isinstance(model_qtd_path, Path) and model_qtd_path.exists():
                st.download_button(
                    label="Baixar Curva ABC (QTD).xlsx",
                    data=_read_binary(model_qtd_path),
                    file_name="Curva ABC (QTD).xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            else:
                st.error("❌ Planilha modelo não encontrada: Curva ABC (QTD).xlsx")

    with tab_descricao:
        for msg in load_messages:
            st.info(msg)
else:
    st.markdown("""
    <div style="
        text-align: center;
        padding: 50px;
        background: linear-gradient(135deg, rgba(6, 214, 160, 0.1) 0%, rgba(17, 138, 178, 0.1) 100%);
        border-radius: 20px;
        border: 2px solid rgba(6, 214, 160, 0.3);
        margin: 20px 0;
    ">
        <h1 style="color: #06d6a0; font-size: 3em; margin-bottom: 20px;">👋 Bem-vindo ao Dashboard ABC!</h1>
        <p style="color: #e0e0e0; font-size: 1.2em; margin-bottom: 30px;">
            Análise inteligente da curva ABC para otimização de estoques e vendas.
        </p>
        <div style="
            background: rgba(31, 119, 180, 0.1);
            padding: 20px;
            border-radius: 10px;
            margin: 20px 0;
            text-align: left;
            max-width: 600px;
            margin-left: auto;
            margin-right: auto;
        ">
            <h3 style="color: #ffffff; margin-top: 0;">📤 Faça upload do seu arquivo Excel</h3>
            <p style="color: #cccccc; margin-bottom: 10px;">O arquivo deve conter as seguintes colunas:</p>
            <ul style="color: #cccccc;">
                <li><code>descricao</code> - Descrição do produto</li>
                <li><code>KG</code> - Quantidade em quilogramas</li>
                <li><code>% individual</code> - Percentual individual</li>
                <li><code>Tipo Item</code> - Tipo/categoria do item</li>
                <li><code>% acumulado</code> - Percentual acumulado</li>
            </ul>
        </div>
        <p style="color: #b0b0b0; font-style: italic;">
            Após o upload, explore gráficos interativos e insights sobre seus produtos mais importantes!
        </p>
    </div>
    """, unsafe_allow_html=True)
