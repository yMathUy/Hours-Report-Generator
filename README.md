# ⏰ Gerador de Relatório de Horas

Uma aplicação Streamlit para gerar relatórios automáticos de horas trabalhadas com análise e exportação em Excel.

## 🎯 Funcionalidades

✅ **Upload de arquivos** - Suporta Excel e CSV  
✅ **Análise automática** - Calcula total de horas, média diária e distribuição por projeto  
✅ **Visualizações** - Gráficos interativos com Plotly  
✅ **Exportação** - Gera relatório em Excel com múltiplas abas  
✅ **Filtros** - Filtre dados por projeto e período  

## 📋 Formato de Entrada

O arquivo deve conter as seguintes colunas:

| Coluna | Descrição | Exemplo |
|--------|-----------|---------|
| **Data** | Data do trabalho | 01/01/2024 |
| **Horas** | Horas trabalhadas | 8.5 |
| **Projeto** | Nome do projeto/categoria | Projeto A |

### Exemplo de arquivo (CSV ou Excel)
```
Data,Horas,Projeto
01/01/2024,8,Projeto A
02/01/2024,7.5,Projeto B
03/01/2024,8,Projeto A
04/01/2024,6,Projeto C
```

## 🚀 Como usar

### 1. Instalar dependências
```bash
pip install -r requirements.txt
```

### 2. Executar a aplicação
```bash
streamlit run streamlit_app.py
```

### 3. Usar a aplicação
1. Abra a aplicação em seu navegador (geralmente em `http://localhost:8501`)
2. Carregue seu arquivo Excel ou CSV na barra lateral
3. Visualize os dados nas diferentes abas:
   - **📊 Resumo**: Métricas principais e distribuição por projeto
   - **📈 Gráficos**: Visualizações interativas
   - **📋 Detalhes**: Lista completa de registros com filtros
   - **💾 Exportar**: Baixe o relatório em Excel

## 📊 Abas da Aplicação

### Resumo
- Total de horas registradas
- Total de dias trabalhados
- Média de horas por dia
- Número de projetos diferentes
- Gráfico de distribuição por projeto

### Gráficos
- Horas trabalhadas por dia (gráfico de barras)
- Horas acumuladas ao longo do tempo (gráfico de linhas)

### Detalhes
- Lista completa de registros
- Filtros por projeto e período de datas

### Exportar
- Gera arquivo Excel com:
  - Aba "Resumo": Métricas principais
  - Aba "Por Projeto": Análise por projeto
  - Aba "Detalhes": Registros completos

## 🛠️ Tecnologias

- **Streamlit** - Interface web interativa
- **Pandas** - Processamento de dados
- **Plotly** - Gráficos interativos
- **OpenPyXL** - Exportação para Excel

## 📝 Dependências

- `streamlit` - Framework web
- `pandas` - Manipulação de dados
- `openpyxl` - Suporte a Excel
- `plotly` - Visualizações interativas
- `python-dateutil` - Processamento de datas

## 🐛 Troubleshooting

**Erro: "Colunas obrigatórias ausentes"**
- Verifique se seu arquivo tem as colunas: Data, Horas e Projeto
- Os nomes das colunas são case-insensitive (MAIÚSCULAS ou minúsculas)

**Erro: "Nenhuma linha com dados válidos encontrada"**
- Verifique se o formato das datas está correto (DD/MM/YYYY)
- Verifique se os valores de horas são números válidos

## 📄 Licença

MIT License - veja o arquivo LICENSE para detalhes
