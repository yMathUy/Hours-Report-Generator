# ⏰ Gerador de Relatório de Horas

Uma aplicação Streamlit para gerar relatórios automáticos de horas trabalhadas com análise e exportação em Excel.

[![Streamlit App](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://hours-report-generator.streamlit.app/)

## 🎯 Funcionalidades

✅ **Upload de arquivos** - Suporta Excel e CSV, incluindo exports do ClickUp
✅ **Análise automática** - Calcula total de horas, média diária e distribuição por projeto
✅ **Visualizações** - Gráficos interativos com Plotly
✅ **Exportação** - Gera relatório em Excel simples ou mantém estrutura completa do template
✅ **Template Management** - Salva templates e mantém histórico de dados
✅ **Filtros** - Filtre dados por projeto e período

## 📋 Formatos suportados

### ClickUp Export
- Coluna "Time Tracked Text" (ex: '2 h 30 m')
- Coluna "Start Text" (data/hora)
- Coluna "Task Name" (nome da task)

### Formato customizado
- **Data**: DD/MM/YYYY
- **Horas**: 8.5 (decimal)
- **Task**: Nome da task

## 🚀 Como usar

### Localmente
```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
```

### Streamlit Cloud
1. Fork este repositório
2. Acesse [share.streamlit.io](https://share.streamlit.io)
3. Conecte seu GitHub e selecione este repositório
4. Clique em "Deploy"

## 📊 Abas da Aplicação

### Resumo
- Total de horas registradas
- Total de dias trabalhados
- Média de horas por dia
- Número de tarefas diferentes
- Gráfico de distribuição por tarefa

### Gráficos
- Horas trabalhadas por dia (gráfico de barras)
- Horas acumuladas ao longo do tempo (gráfico de linhas)

### Detalhes
- Lista completa de registros
- Filtros por tarefa e período de datas

### Exportar
- **Exportar Simples**: Excel com 3 abas (Resumo, Por Tarefa, Detalhes)
- **Exportar com Template**: Mantém estrutura completa do arquivo original

## 🛠️ Tecnologias

- **Streamlit** - Interface web interativa
- **Pandas** - Processamento de dados
- **Plotly** - Gráficos interativos
- **OpenPyXL** - Suporte a Excel
- **Python-dateutil** - Processamento de datas

## 📝 Dependências

- `streamlit` - Framework web
- `pandas` - Manipulação de dados
- `openpyxl` - Suporte a Excel
- `plotly` - Visualizações interativas
- `python-dateutil` - Processamento de datas

## 🐛 Troubleshooting

**Erro: "Colunas obrigatórias ausentes"**
- Verifique se seu arquivo tem as colunas corretas
- Para ClickUp: "Time Tracked Text", "Task Name", "Start Text"

**Erro: "Template não encontrado"**
- Faça upload de um arquivo template primeiro
- O template será salvo automaticamente

## 📄 Licença

Desenvolvido por MathZerstrer | 2025 - All rights reserved

## 📄 Licença

MIT License - veja o arquivo LICENSE para detalhes
