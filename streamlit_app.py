import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO
from datetime import datetime
import numpy as np
import re
from template_manager import (
    setup_template, get_template_base, add_new_data_to_template,
    prepare_data_for_template, export_with_template, merge_with_history,
    TEMPLATE_BASE_PATH, LAST_EXPORT_PATH
)
from pathlib import Path

# Função para converter tempo em texto para horas decimais
def parse_time_to_decimal(time_str):
    """
    Converte strings de tempo como '2 h 30 m 15 s' para horas decimais
    Também converte milissegundos para horas decimais
    """
    if pd.isna(time_str):
        return 0
    
    time_str = str(time_str).strip()
    
    # Se for número (milissegundos from "Time Tracked" column)
    try:
        ms = float(time_str)
        if ms > 100000:  # Parece ser milissegundos
            return ms / 3600000  # Converter para horas
    except:
        pass
    
    # Parse texto como "2 h 30 m 15 s"
    hours = 0
    minutes = 0
    seconds = 0
    
    # Buscar horas
    h_match = re.search(r'(\d+\.?\d*)\s*h', time_str)
    if h_match:
        hours = float(h_match.group(1))
    
    # Buscar minutos
    m_match = re.search(r'(\d+\.?\d*)\s*m', time_str)
    if m_match:
        minutes = float(m_match.group(1))
    
    # Buscar segundos
    s_match = re.search(r'(\d+\.?\d*)\s*s', time_str)
    if s_match:
        seconds = float(s_match.group(1))
    
    # Converter tudo para horas
    total_hours = hours + (minutes / 60) + (seconds / 3600)
    return total_hours

# Configuração da página
st.set_page_config(
    page_title="Gerador de Relatório de Horas",
    page_icon="⏰",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.title("⏰ Gerador de Relatório de Horas")
st.markdown("---")

# Sidebar para upload
with st.sidebar:
    st.header("📤 Carregar dados")
    uploaded_file = st.file_uploader(
        "Selecione um arquivo (Excel ou CSV)",
        type=["xlsx", "csv", "xls"],
        help="Suporta: Formato simples (Data, Horas, Projeto) ou exports do ClickUp/Time Tracking"
    )
    
    st.markdown("---")
    st.markdown("""
    ### 📋 Formatos suportados:
    
    **Formato Simples:**
    - Data, Horas, Projeto
    
    **ClickUp/Time Tracking:**
    - "Time Tracked Text" (13 m, 2 h, etc)
    - "Task Name" (Projeto/Task)
    - "Start Text" (Data do registro)
    """)

# Processamento dos dados
if uploaded_file is not None:
    try:
        # Ler o arquivo
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        else:
            # Tentar ler Excel com diferentes engines
            try:
                df = pd.read_excel(uploaded_file, engine='openpyxl')
            except ImportError:
                try:
                    df = pd.read_excel(uploaded_file, engine='xlrd')
                except ImportError:
                    st.error("❌ Não foi possível ler o arquivo Excel. Instale openpyxl ou xlrd.")
                    st.stop()
            except Exception as e:
                st.error(f"❌ Erro ao ler arquivo Excel: {str(e)}")
                st.stop()
        
        # Se o arquivo é um template (contém as abas específicas), salvá-lo automaticamente
        if uploaded_file.name.endswith('.xlsx'):
            try:
                xl_file = pd.ExcelFile(uploaded_file)
                if 'Time Reporting Export' in xl_file.sheet_names:
                    # É um template! Salvar para uso futuro
                    template_path = Path("templates") / "uploaded_template.xlsx"
                    Path("templates").mkdir(exist_ok=True)
                    uploaded_file.seek(0)
                    template_path.write_bytes(uploaded_file.getvalue())
                    st.info("✅ Template salvo para uso futuro na exportação!")
                    setup_template(template_path)
            except:
                pass
        
        # Normalizar nomes das colunas
        df.columns = df.columns.str.strip()
        
        # Detectar o formato do arquivo
        is_time_tracking = "Time Tracked Text" in df.columns or "Task Name" in df.columns
        
        if is_time_tracking:
            # Formato ClickUp/Time Tracking
            st.info("✅ Formato detectado: ClickUp/Time Tracking")
            
            # Verificar colunas obrigatórias
            if "Time Tracked Text" not in df.columns or "Task Name" not in df.columns or "Start Text" not in df.columns:
                st.error("❌ Arquivo deve conter: 'Time Tracked Text', 'Task Name' e 'Start Text'")
            else:
                # Criar dataframe processado
                df_processed = pd.DataFrame()
                
                # Extrair data do "Start Text" (formato: "02/02/2026, 7:47:45 AM -03")
                df_processed['data'] = pd.to_datetime(
                    df["Start Text"].str.split(',').str[0],
                    format='%d/%m/%Y',
                    errors='coerce'
                )
                
                # Converter tempo para decimal
                df_processed['horas'] = df["Time Tracked Text"].apply(parse_time_to_decimal)
                
                # Task como projeto
                df_processed['projeto'] = df["Task Name"].fillna('Sem tarefa')
                
                df_valid = df_processed.dropna(subset=['data', 'horas'])
                
                if len(df_valid) == 0:
                    st.error("❌ Nenhuma linha com dados válidos encontrada")
                else:
                    # Remover linhas com 0 horas
                    df_valid = df_valid[df_valid['horas'] > 0]
                    
                    if len(df_valid) == 0:
                        st.error("❌ Nenhum registro com horas > 0 encontrado")
                    else:
                        st.success(f"✅ {len(df_valid)} registros processados com sucesso!")
                        
                        # Ordenar por data
                        df_valid = df_valid.sort_values('data')
                        
                        # Tabs para diferentes visualizações
                        tab1, tab2, tab3, tab4 = st.tabs(["📊 Resumo", "📈 Gráficos", "📋 Detalhes", "💾 Exportar"])
                        
                        with tab1:
                            col1, col2, col3, col4 = st.columns(4)
                            
                            total_horas = df_valid['horas'].sum()
                            total_dias = df_valid['data'].dt.date.nunique()
                            media_horas = df_valid['horas'].mean()
                            num_projetos = df_valid['projeto'].nunique()
                            
                            with col1:
                                st.metric("Total de Horas", f"{total_horas:.2f}h")
                            with col2:
                                st.metric("Total de Dias", f"{total_dias}")
                            with col3:
                                st.metric("Média por Dia", f"{media_horas:.2f}h")
                            with col4:
                                st.metric("Número de Tarefas", f"{num_projetos}")
                            
                            st.markdown("---")
                            
                            # Horas por tarefa/projeto
                            st.subheader("⏱️ Horas por Tarefa")
                            horas_por_projeto = df_valid.groupby('projeto')['horas'].sum().sort_values(ascending=False)
                            
                            col1, col2 = st.columns([1, 2])
                            with col1:
                                st.dataframe(
                                    horas_por_projeto.reset_index().rename(columns={'horas': 'Total de Horas', 'projeto': 'Tarefa'}),
                                    use_container_width=True,
                                    hide_index=True
                                )
                            
                            with col2:
                                fig = px.pie(
                                    values=horas_por_projeto.values,
                                    names=horas_por_projeto.index,
                                    title="Distribuição de Horas por Tarefa",
                                    hole=0.3
                                )
                                st.plotly_chart(fig, use_container_width=True)
                        
                        with tab2:
                            col1, col2 = st.columns(2)
                            
                            with col1:
                                # Gráfico de horas por dia
                                horas_por_dia = df_valid.groupby(df_valid['data'].dt.date)['horas'].sum()
                                fig1 = px.bar(
                                    x=horas_por_dia.index,
                                    y=horas_por_dia.values,
                                    title="Horas Trabalhadas por Dia",
                                    labels={'x': 'Data', 'y': 'Horas'},
                                    color=horas_por_dia.values,
                                    color_continuous_scale='Viridis'
                                )
                                fig1.update_layout(showlegend=False)
                                st.plotly_chart(fig1, use_container_width=True)
                            
                            with col2:
                                # Gráfico de horas acumuladas
                                df_sorted = df_valid.sort_values('data')
                                df_sorted['horas_acumuladas'] = df_sorted['horas'].cumsum()
                                
                                fig2 = go.Figure()
                                fig2.add_trace(go.Scatter(
                                    x=df_sorted['data'],
                                    y=df_sorted['horas_acumuladas'],
                                    mode='lines+markers',
                                    name='Horas Acumuladas',
                                    line=dict(color='#03A9F4', width=3),
                                    fill='tozeroy'
                                ))
                                fig2.update_layout(
                                    title="Horas Acumuladas",
                                    xaxis_title="Data",
                                    yaxis_title="Total de Horas",
                                    hovermode='x unified'
                                )
                                st.plotly_chart(fig2, use_container_width=True)
                        
                        with tab3:
                            st.subheader("Detalhes dos Registros de Tempo")
                            
                            # Filtros
                            col1, col2 = st.columns(2)
                            with col1:
                                projetos_filtro = st.multiselect(
                                    "Filtrar por Tarefa",
                                    options=sorted(df_valid['projeto'].unique()),
                                    default=sorted(df_valid['projeto'].unique())
                                )
                            
                            with col2:
                                data_range = st.date_input(
                                    "Intervalo de Datas",
                                    value=(df_valid['data'].min().date(), df_valid['data'].max().date()),
                                    max_value=df_valid['data'].max().date(),
                                    format="DD/MM/YYYY"
                                )
                            
                            # Aplicar filtros
                            df_filtrado = df_valid[
                                (df_valid['projeto'].isin(projetos_filtro)) &
                                (df_valid['data'].dt.date >= data_range[0]) &
                                (df_valid['data'].dt.date <= data_range[1])
                            ].copy()
                            
                            df_filtrado['data'] = df_filtrado['data'].dt.strftime('%d/%m/%Y')
                            df_filtrado['horas'] = df_filtrado['horas'].round(2)
                            df_filtrado = df_filtrado[['data', 'projeto', 'horas']].reset_index(drop=True)
                            df_filtrado.columns = ['Data', 'Tarefa', 'Horas']
                            
                            st.dataframe(df_filtrado, use_container_width=True, hide_index=True)
                            
                            st.metric("Total (filtrado)", f"{df_filtrado['Horas'].sum():.2f}h")
                        
                        with tab4:
                            st.subheader("💾 Exportar Relatório")
                            
                            # Verificar se existe template salvo
                            template_exists = TEMPLATE_BASE_PATH.exists() or get_template_base() is not None
                            
                            col1, col2 = st.columns(2)
                            
                            with col1:
                                if st.button("📊 Exportar Simples (Excel)", use_container_width=True):
                                    # Criar arquivo Excel com múltiplas abas
                                    excel_buffer = BytesIO()
                                    with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                                        # Aba 1: Resumo
                                        resumo_data = {
                                            'Métrica': ['Total de Horas', 'Total de Dias', 'Média por Dia', 'Número de Tarefas'],
                                            'Valor': [f"{total_horas:.2f}h", f"{total_dias}", f"{media_horas:.2f}h", f"{num_projetos}"]
                                        }
                                        pd.DataFrame(resumo_data).to_excel(writer, sheet_name='Resumo', index=False)
                                        
                                        # Aba 2: Horas por Tarefa
                                        horas_export = horas_por_projeto.reset_index()
                                        horas_export.columns = ['Tarefa', 'Total de Horas']
                                        horas_export['Total de Horas'] = horas_export['Total de Horas'].round(2)
                                        horas_export.to_excel(writer, sheet_name='Por Tarefa', index=False)
                                        
                                        # Aba 3: Dados Detalhados
                                        df_export = df_valid.copy()
                                        df_export['data'] = df_export['data'].dt.strftime('%d/%m/%Y')
                                        df_export['horas'] = df_export['horas'].round(2)
                                        df_export = df_export[['data', 'projeto', 'horas']]
                                        df_export.columns = ['Data', 'Tarefa', 'Horas']
                                        df_export.to_excel(writer, sheet_name='Detalhes', index=False)
                                    
                                    excel_buffer.seek(0)
                                    
                                    st.download_button(
                                        label="📥 Baixar Excel",
                                        data=excel_buffer,
                                        file_name=f"relatorio_horas_{datetime.now().strftime('%d%m%Y_%H%M%S')}.xlsx",
                                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                        use_container_width=True
                                    )
                                    
                                    st.success("✅ Arquivo gerado com sucesso!")
                            
                            with col2:
                                if template_exists:
                                    if st.button("📋 Exportar com Template", use_container_width=True, type="primary"):
                                        try:
                                            st.info("⏳ Processando export com template...")
                                            
                                            # Preparar dados no formato do template
                                            df_prep = df_valid.copy()
                                            df_prep.rename(columns={'projeto': 'Task Name'}, inplace=True)
                                            
                                            # Combinar com histórico se existir
                                            df_merged = merge_with_history(df_prep, use_last_export=True)
                                            
                                            # Adicionar dados ao template
                                            df_combined, wb = add_new_data_to_template(
                                                df_merged,
                                                template_path=get_template_base()
                                            )
                                            
                                            # Exportar
                                            output_file = BytesIO()
                                            wb.save(output_file)
                                            output_file.seek(0)
                                            
                                            # Salvar última exportação como referência
                                            last_export_path = Path("templates") / "last_export.xlsx"
                                            with open(last_export_path, 'wb') as f:
                                                f.write(output_file.getvalue())
                                            
                                            output_file.seek(0)
                                            
                                            st.download_button(
                                                label="📥 Baixar Arquivo Completo",
                                                data=output_file,
                                                file_name=f"Relatorio_Completo_{datetime.now().strftime('%d%m%Y_%H%M%S')}.xlsx",
                                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                                use_container_width=True
                                            )
                                            
                                            st.success("✅ Arquivo exportado com template! Histórico salvo para próximas vezes.")
                                            st.info(f"📊 Dados adicionados: {len(df_merged)} registros")
                                        
                                        except Exception as e:
                                            st.error(f"❌ Erro ao exportar: {str(e)}")
                                            st.info("💡 Dica: Verifique se o arquivo do template está no formato correto")
                                else:
                                    st.button("📋 Exportar com Template (não disponível)", disabled=True, use_container_width=True)
                                    st.info("⚠️ Nenhum template salvo. Faça upload de um arquivo de template primeiro!")
                            
                            st.markdown("---")
                            
                            if template_exists:
                                st.success(f"✅ **Template salvo**: {TEMPLATE_BASE_PATH.name}")
                            else:
                                st.warning("❌ Nenhum template configurado")
        
        else:
            # Formato simples original
            st.info("✅ Formato detectado: Formato Simples")
            
            # Normalizar nomes das colunas
            df.columns = df.columns.str.strip().str.lower()
            
            # Validar colunas obrigatórias
            required_cols = ['data', 'horas', 'projeto']
            missing_cols = [col for col in required_cols if col not in df.columns]
            
            if missing_cols:
                st.error(f"❌ Colunas obrigatórias ausentes: {', '.join(missing_cols)}")
                st.info("O arquivo deve conter as colunas: Data, Horas, Projeto")
            else:
                # Limpeza e conversão de dados
                df['data'] = pd.to_datetime(df['data'], format='%d/%m/%Y', errors='coerce')
                df['horas'] = pd.to_numeric(df['horas'], errors='coerce')
                df['projeto'] = df['projeto'].fillna('Sem projeto')
                
                # Remover linhas com dados inválidos
                df_valid = df.dropna(subset=['data', 'horas'])
                
                if len(df_valid) == 0:
                    st.error("❌ Nenhuma linha com dados válidos encontrada")
                else:
                    if len(df_valid) < len(df):
                        st.warning(f"⚠️ {len(df) - len(df_valid)} linha(s) com dados inválidos foram removidas")
                    
                    # Ordenar por data
                    df_valid = df_valid.sort_values('data')
                    
                    # Tabs para diferentes visualizações
                    tab1, tab2, tab3, tab4 = st.tabs(["📊 Resumo", "📈 Gráficos", "📋 Detalhes", "💾 Exportar"])
                    
                    with tab1:
                        col1, col2, col3, col4 = st.columns(4)
                        
                        total_horas = df_valid['horas'].sum()
                        total_dias = df_valid['data'].dt.date.nunique()
                        media_horas = df_valid['horas'].mean()
                        num_projetos = df_valid['projeto'].nunique()
                        
                        with col1:
                            st.metric("Total de Horas", f"{total_horas:.1f}h")
                        with col2:
                            st.metric("Total de Dias", f"{total_dias}")
                        with col3:
                            st.metric("Média por Dia", f"{media_horas:.1f}h")
                        with col4:
                            st.metric("Número de Projetos", f"{num_projetos}")
                        
                        st.markdown("---")
                        
                        # Horas por projeto
                        st.subheader("Horas por Projeto")
                        horas_por_projeto = df_valid.groupby('projeto')['horas'].sum().sort_values(ascending=False)
                        
                        col1, col2 = st.columns([1, 2])
                        with col1:
                            st.dataframe(
                                horas_por_projeto.reset_index(name='Total de Horas'),
                                use_container_width=True,
                                hide_index=True
                            )
                        
                        with col2:
                            fig = px.pie(
                                values=horas_por_projeto.values,
                                names=horas_por_projeto.index,
                                title="Distribuição de Horas por Projeto",
                                hole=0.3
                            )
                            st.plotly_chart(fig, use_container_width=True)
                    
                    with tab2:
                        col1, col2 = st.columns(2)
                        
                        with col1:
                            # Gráfico de horas por dia
                            horas_por_dia = df_valid.groupby(df_valid['data'].dt.date)['horas'].sum()
                            fig1 = px.bar(
                                x=horas_por_dia.index,
                                y=horas_por_dia.values,
                                title="Horas Trabalhadas por Dia",
                                labels={'x': 'Data', 'y': 'Horas'},
                                color=horas_por_dia.values,
                                color_continuous_scale='Viridis'
                            )
                            fig1.update_layout(showlegend=False)
                            st.plotly_chart(fig1, use_container_width=True)
                        
                        with col2:
                            # Gráfico de horas acumuladas
                            df_sorted = df_valid.sort_values('data')
                            df_sorted['horas_acumuladas'] = df_sorted['horas'].cumsum()
                            
                            fig2 = go.Figure()
                            fig2.add_trace(go.Scatter(
                                x=df_sorted['data'],
                                y=df_sorted['horas_acumuladas'],
                                mode='lines+markers',
                                name='Horas Acumuladas',
                                line=dict(color='#03A9F4', width=3),
                                fill='tozeroy'
                            ))
                            fig2.update_layout(
                                title="Horas Acumuladas",
                                xaxis_title="Data",
                                yaxis_title="Total de Horas",
                                hovermode='x unified'
                            )
                            st.plotly_chart(fig2, use_container_width=True)
                    
                    with tab3:
                        st.subheader("Detalhes das Horas Registradas")
                        
                        # Filtros
                        col1, col2 = st.columns(2)
                        with col1:
                            projetos_filtro = st.multiselect(
                                "Filtrar por Projeto",
                                options=sorted(df_valid['projeto'].unique()),
                                default=sorted(df_valid['projeto'].unique())
                            )
                        
                        with col2:
                            data_range = st.date_input(
                                "Intervalo de Datas",
                                value=(df_valid['data'].min().date(), df_valid['data'].max().date()),
                                max_value=df_valid['data'].max().date(),
                                format="DD/MM/YYYY"
                            )
                        
                        # Aplicar filtros
                        df_filtrado = df_valid[
                            (df_valid['projeto'].isin(projetos_filtro)) &
                            (df_valid['data'].dt.date >= data_range[0]) &
                            (df_valid['data'].dt.date <= data_range[1])
                        ].copy()
                        
                        df_filtrado['data'] = df_filtrado['data'].dt.strftime('%d/%m/%Y')
                        df_filtrado = df_filtrado[['data', 'projeto', 'horas']].reset_index(drop=True)
                        
                        st.dataframe(df_filtrado, use_container_width=True, hide_index=True)
                        
                        st.metric("Total (filtrado)", f"{df_filtrado['horas'].sum():.1f}h")
                    
                    with tab4:
                        st.subheader("💾 Exportar Relatório")
                        
                        if st.button("Gerar arquivo Excel", use_container_width=True, type="primary"):
                            # Criar arquivo Excel com múltiplas abas
                            excel_buffer = BytesIO()
                            with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                                # Aba 1: Resumo
                                resumo_data = {
                                    'Métrica': ['Total de Horas', 'Total de Dias', 'Média por Dia', 'Número de Projetos'],
                                    'Valor': [f"{total_horas:.1f}h", f"{total_dias}", f"{media_horas:.1f}h", f"{num_projetos}"]
                                }
                                pd.DataFrame(resumo_data).to_excel(writer, sheet_name='Resumo', index=False)
                                
                                # Aba 2: Horas por Projeto
                                horas_por_projeto.reset_index(name='Total de Horas').to_excel(
                                    writer, sheet_name='Por Projeto', index=False
                                )
                                
                                # Aba 3: Dados Detalhados
                                df_export = df_valid.copy()
                                df_export['data'] = df_export['data'].dt.strftime('%d/%m/%Y')
                                df_export = df_export[['data', 'projeto', 'horas']]
                                df_export.to_excel(writer, sheet_name='Detalhes', index=False)
                            
                            excel_buffer.seek(0)
                            
                            st.download_button(
                                label="📥 Baixar Excel",
                                data=excel_buffer,
                                file_name=f"relatorio_horas_{datetime.now().strftime('%d%m%Y_%H%M%S')}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                use_container_width=True
                            )
                            
                            st.success("✅ Arquivo gerado com sucesso!")
    
    except Exception as e:
        st.error(f"❌ Erro ao processar arquivo: {str(e)}")
        st.info("Verifique se o arquivo está no formato correto")

else:
    st.info(
        """
        👈 **Comece carregando um arquivo!**
        
        **Formatos suportados:**
        
        1️⃣ **Formato Simples:**
           - Data (DD/MM/YYYY)
           - Horas (número decimal)
           - Projeto (texto)
        
        2️⃣ **ClickUp/Time Tracking:**
           - Time Tracked Text (13 m, 2 h 30 m, etc)
           - Task Name (nome da tarefa)
           - Start Text (data e hora)
        
        A aplicação detecta automaticamente o formato! ✨
        """
    )

# Rodapé
st.markdown("---")
st.markdown(
    """
    <div style='text-align: center; color: #888; font-size: 12px;'>
        Desenvolvido por MathZerstrer | 2025 - All rights reserved
    </div>
    """,
    unsafe_allow_html=True
)
