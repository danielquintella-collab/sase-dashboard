#!/usr/bin/env python3
"""
Script para Automação do Dashboard SASE
Processa dados do Google Drive e gera HTML atualizado
Rodar automaticamente via GitHub Actions todo dia às 23h
"""

import os
import json
import pandas as pd
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from io import BytesIO
from datetime import datetime

def baixar_excel_google_drive():
    """Baixa arquivo Excel do Google Drive usando Service Account"""
    
    # Carregar credenciais
    creds_json = os.environ.get('GOOGLE_CREDENTIALS')
    if not creds_json:
        raise ValueError("GOOGLE_CREDENTIALS não configurado nos secrets do GitHub")
    
    creds_dict = json.loads(creds_json)
    credentials = Credentials.from_service_account_info(creds_dict)
    
    # Conectar ao Google Drive
    drive_service = build('drive', 'v3', credentials=credentials)
    
    # Buscar arquivo
    folder_id = os.environ.get('GOOGLE_DRIVE_FOLDER_ID')
    if not folder_id:
        raise ValueError("GOOGLE_DRIVE_FOLDER_ID não configurado nos secrets do GitHub")
    
    query = f"'{folder_id}' in parents and name='CONTROLE-SASE-CAXIAS.xlsx' and trashed=false"
    results = drive_service.files().list(q=query, spaces='drive', pageSize=1).execute()
    items = results.get('files', [])
    
    if not items:
        raise FileNotFoundError("Arquivo CONTROLE-SASE-CAXIAS.xlsx não encontrado no Google Drive")
    
    file_id = items[0]['id']
    print(f"✅ Arquivo encontrado: {items[0]['name']}")
    
    # Baixar arquivo
    request = drive_service.files().get_media(fileId=file_id)
    file_content = BytesIO()
    downloader = MediaIoBaseDownload(file_content, request)
    
    done = False
    while not done:
        status, done = downloader.next_chunk()
    
    file_content.seek(0)
    print("✅ Arquivo baixado com sucesso")
    return file_content

def processar_dados(file_content):
    """Processa dados Excel e gera agregações"""
    
    df = pd.read_excel(file_content, sheet_name='Página1', header=0)
    
    # Renomear colunas
    df.columns = ['Nome_Paciente', 'Data_Realizacao', 'Convenio', 'Exame_Realizado', 
                  'Valor_Exame', 'Taxa_Cartao', 'Percentual_SASE', 'Percentual_Medico', 'RDF']
    
    # Remover cabeçalho duplicado
    df = df[df['Nome_Paciente'] != 'Nome do Paciente'].copy()
    
    # Converter tipos
    df['Data_Realizacao'] = pd.to_datetime(df['Data_Realizacao'], errors='coerce')
    df['Valor_Exame'] = pd.to_numeric(df['Valor_Exame'], errors='coerce')
    df['Taxa_Cartao'] = pd.to_numeric(df['Taxa_Cartao'], errors='coerce')
    df['Percentual_SASE'] = pd.to_numeric(df['Percentual_SASE'], errors='coerce')
    df['Percentual_Medico'] = pd.to_numeric(df['Percentual_Medico'], errors='coerce')
    df['RDF'] = pd.to_numeric(df['RDF'], errors='coerce')
    
    print(f"✅ Dados carregados: {len(df)} registros")
    
    # Categorizar modalidade - AUTOMÁTICO
    def categorizar_modalidade(exame):
        if pd.isna(exame):
            return 'Outros'
        exame_upper = str(exame).upper()
        
        # Mapeamento que reconhece novos tipos automaticamente
        modalidades = {
            'MAMOGRAFIA': 'Mamografia',
            'DENSITOMETRIA': 'Densitometria',
            'ULTRASSONOGRAFIA': 'Ultrassonografia',
            'ULTRASSOM': 'Ultrassonografia',
            'RAIO X': 'Raio X',
            'RAIO-X': 'Raio X',
            'RX': 'Raio X',
            'DOPPLER': 'Doppler',
            'DOPPLER VASCULAR': 'Doppler',
            'ECODOPPLER': 'Doppler',
        }
        
        for chave, valor in modalidades.items():
            if chave in exame_upper:
                return valor
        return 'Outros'
    
    df['Modalidade'] = df['Exame_Realizado'].apply(categorizar_modalidade)
    df['Ano_Mes'] = df['Data_Realizacao'].dt.to_period('M').astype(str)
    
    # Agregação por Mês, Convênio e Modalidade
    exames_mes_convenio_modalidade = df.groupby(['Ano_Mes', 'Convenio', 'Modalidade']).agg({
        'Nome_Paciente': 'count',
        'Valor_Exame': 'sum',
        'RDF': 'sum'
    }).reset_index()
    exames_mes_convenio_modalidade.columns = ['Mes', 'Convenio', 'Modalidade', 'Qtd', 'Receita_Bruta', 'Receita_Liquida']
    exames_mes_convenio_modalidade = exames_mes_convenio_modalidade[exames_mes_convenio_modalidade['Mes'].notna()]
    
    # Resumo Mensal
    resumo_mensal = df.groupby('Ano_Mes').agg({
        'Nome_Paciente': 'count',
        'Valor_Exame': 'sum',
        'RDF': 'sum'
    }).reset_index()
    resumo_mensal.columns = ['Mes', 'Qtd_Exames', 'Receita_Bruta', 'Receita_Liquida']
    resumo_mensal['Percentual_Lucro'] = (resumo_mensal['Receita_Liquida'] / resumo_mensal['Receita_Bruta'] * 100).round(2)
    resumo_mensal = resumo_mensal[resumo_mensal['Mes'].notna()]
    
    # Distribuição por Convênio
    dist_convenio = df.groupby('Convenio').agg({
        'Nome_Paciente': 'count',
        'Valor_Exame': 'sum',
        'RDF': 'sum'
    }).reset_index()
    dist_convenio.columns = ['Convenio', 'Qtd', 'Receita_Bruta', 'Receita_Liquida']
    dist_convenio['Percentual'] = (dist_convenio['Qtd'] / dist_convenio['Qtd'].sum() * 100).round(2)
    dist_convenio = dist_convenio.sort_values('Qtd', ascending=False)
    
    # Distribuição por Modalidade
    dist_modalidade = df.groupby('Modalidade').agg({
        'Nome_Paciente': 'count',
        'Valor_Exame': 'sum',
        'RDF': 'sum'
    }).reset_index()
    dist_modalidade.columns = ['Modalidade', 'Qtd', 'Receita_Bruta', 'Receita_Liquida']
    dist_modalidade['Percentual'] = (dist_modalidade['Qtd'] / dist_modalidade['Qtd'].sum() * 100).round(2)
    dist_modalidade = dist_modalidade.sort_values('Qtd', ascending=False)
    
    # Distribuição Modalidade por Mês
    dist_modalidade_mes = df.groupby(['Ano_Mes', 'Modalidade']).agg({
        'Nome_Paciente': 'count'
    }).reset_index()
    dist_modalidade_mes.columns = ['Mes', 'Modalidade', 'Qtd']
    dist_modalidade_mes = dist_modalidade_mes[dist_modalidade_mes['Mes'].notna()]
    
    # Gerar JSON
    data_json = {
        'exames_mes_convenio_modalidade': exames_mes_convenio_modalidade.to_dict(orient='records'),
        'resumo_mensal': resumo_mensal.to_dict(orient='records'),
        'dist_convenio': dist_convenio.to_dict(orient='records'),
        'dist_modalidade': dist_modalidade.to_dict(orient='records'),
        'dist_modalidade_mes': dist_modalidade_mes.to_dict(orient='records'),
        'ultima_atualizacao': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    }
    
    print("✅ Dados processados com sucesso")
    print(f"   - Total de exames: {resumo_mensal['Qtd_Exames'].sum() if len(resumo_mensal) > 0 else 0}")
    print(f"   - Modalidades encontradas: {', '.join(dist_modalidade['Modalidade'].tolist())}")
    
    return data_json

def gerar_html(data_json):
    """Gera HTML com dados embutidos"""
    
    # Converter JSON para string JavaScript
    data_js = json.dumps(data_json, ensure_ascii=False)
    
    html_content = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Dashboard SASE Caxias - BI</title>
    <script src="https://cdn.jsdelivr.net/npm/chart.js@3.9.1/dist/chart.min.js"></script>
    <style>
        * {{
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }}
        
        body {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background-color: #f5f7fa;
            color: #333;
        }}
        
        .header {{
            background: linear-gradient(135deg, #1e3a5f 0%, #2c5282 100%);
            color: white;
            padding: 25px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
            margin-bottom: 20px;
        }}
        
        .header h1 {{
            font-size: 28px;
            margin-bottom: 5px;
        }}
        
        .header p {{
            font-size: 14px;
            opacity: 0.9;
        }}
        
        .container {{
            display: flex;
            max-width: 1600px;
            margin: 0 auto;
            gap: 20px;
            padding: 20px;
        }}
        
        .sidebar {{
            width: 250px;
            background: white;
            padding: 20px;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.05);
            height: fit-content;
        }}
        
        .sidebar h3 {{
            font-size: 16px;
            margin-bottom: 15px;
            color: #1e3a5f;
            border-bottom: 2px solid #ff6b35;
            padding-bottom: 10px;
        }}
        
        .filter-group {{
            margin-bottom: 20px;
        }}
        
        .filter-group label {{
            display: block;
            font-size: 13px;
            font-weight: 600;
            margin-bottom: 8px;
            color: #555;
        }}
        
        .filter-group select {{
            width: 100%;
            padding: 8px;
            border: 1px solid #ddd;
            border-radius: 4px;
            font-size: 13px;
        }}
        
        .filter-group input[type="checkbox"] {{
            margin-right: 8px;
        }}
        
        .checkbox-item {{
            display: flex;
            align-items: center;
            margin-bottom: 8px;
            font-size: 13px;
        }}
        
        .btn {{
            width: 100%;
            padding: 10px;
            background: #ff6b35;
            color: white;
            border: none;
            border-radius: 4px;
            cursor: pointer;
            font-weight: 600;
            margin-top: 10px;
            transition: background 0.3s;
        }}
        
        .btn:hover {{
            background: #e55a24;
        }}
        
        .main-content {{
            flex: 1;
        }}
        
        .kpi-container {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 15px;
            margin-bottom: 30px;
        }}
        
        .kpi-card {{
            background: white;
            padding: 20px;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.05);
            border-left: 4px solid #ff6b35;
        }}
        
        .kpi-value {{
            font-size: 28px;
            font-weight: bold;
            color: #1e3a5f;
            margin: 10px 0;
        }}
        
        .kpi-label {{
            font-size: 12px;
            color: #999;
            text-transform: uppercase;
            letter-spacing: 1px;
        }}
        
        .charts-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(500px, 1fr));
            gap: 20px;
            margin-bottom: 30px;
        }}
        
        .chart-container {{
            background: white;
            padding: 20px;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.05);
        }}
        
        .chart-title {{
            font-size: 16px;
            font-weight: 600;
            margin-bottom: 15px;
            color: #1e3a5f;
        }}
        
        .chart-wrapper {{
            position: relative;
            height: 300px;
        }}
        
        .modalidade-cards {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
            gap: 15px;
            margin-bottom: 30px;
        }}
        
        .modalidade-card {{
            background: white;
            padding: 20px;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.05);
            border-top: 3px solid #ff6b35;
        }}
        
        .modalidade-card h3 {{
            font-size: 16px;
            margin-bottom: 15px;
            color: #1e3a5f;
        }}
        
        .modalidade-stat {{
            display: flex;
            justify-content: space-between;
            margin-bottom: 10px;
            font-size: 13px;
        }}
        
        .modalidade-stat span:first-child {{
            color: #999;
        }}
        
        .modalidade-stat span:last-child {{
            font-weight: 600;
            color: #1e3a5f;
        }}
        
        .table-container {{
            background: white;
            padding: 20px;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.05);
            overflow-x: auto;
        }}
        
        .table-title {{
            font-size: 16px;
            font-weight: 600;
            margin-bottom: 15px;
            color: #1e3a5f;
        }}
        
        table {{
            width: 100%;
            border-collapse: collapse;
            font-size: 13px;
        }}
        
        th {{
            background: #f0f0f0;
            padding: 12px;
            text-align: left;
            font-weight: 600;
            color: #333;
            border-bottom: 2px solid #ddd;
        }}
        
        td {{
            padding: 10px 12px;
            border-bottom: 1px solid #eee;
        }}
        
        tr:hover {{
            background: #f9f9f9;
        }}
        
        .footer {{
            background: white;
            padding: 20px;
            border-radius: 8px;
            margin-top: 30px;
            text-align: center;
            color: #999;
            font-size: 12px;
        }}
        
        @media (max-width: 768px) {{
            .container {{
                flex-direction: column;
            }}
            
            .sidebar {{
                width: 100%;
            }}
            
            .charts-grid {{
                grid-template-columns: 1fr;
            }}
        }}
    </style>
</head>
<body>
    <div class="header">
        <h1>📊 Dashboard SASE Caxias</h1>
        <p>Análise Completa de Exames Médicos - Última atualização: {data_json['ultima_atualizacao']}</p>
    </div>
    
    <div class="container">
        <div class="sidebar">
            <h3>🔍 Filtros</h3>
            
            <div class="filter-group">
                <label>Convênio</label>
                <select id="filterConvenio" onchange="aplicarFiltros()">
                    <option value="">Todos</option>
                </select>
            </div>
            
            <div class="filter-group">
                <label>Modalidade</label>
                <div id="filterModalidade"></div>
            </div>
            
            <div class="filter-group">
                <label>Período</label>
                <select id="filterMes" onchange="aplicarFiltros()">
                    <option value="">Todos</option>
                </select>
            </div>
            
            <button class="btn" onclick="limparFiltros()">🔄 Limpar Filtros</button>
        </div>
        
        <div class="main-content">
            <!-- KPIs -->
            <div class="kpi-container">
                <div class="kpi-card">
                    <div class="kpi-label">Total de Exames</div>
                    <div class="kpi-value" id="kpiExames">0</div>
                </div>
                <div class="kpi-card">
                    <div class="kpi-label">Receita Bruta</div>
                    <div class="kpi-value" id="kpiReceitaBruta">R$ 0</div>
                </div>
                <div class="kpi-card">
                    <div class="kpi-label">Receita Líquida</div>
                    <div class="kpi-value" id="kpiReceitaLiquida">R$ 0</div>
                </div>
                <div class="kpi-card">
                    <div class="kpi-label">% de Lucro</div>
                    <div class="kpi-value" id="kpiLucro">0%</div>
                </div>
            </div>
            
            <!-- Gráficos -->
            <div class="charts-grid">
                <div class="chart-container">
                    <div class="chart-title">📈 Exames por Mês</div>
                    <div class="chart-wrapper">
                        <canvas id="chartExamesMes"></canvas>
                    </div>
                </div>
                <div class="chart-container">
                    <div class="chart-title">💰 Distribuição por Convênio</div>
                    <div class="chart-wrapper">
                        <canvas id="chartConvenio"></canvas>
                    </div>
                </div>
            </div>
            
            <div class="charts-grid">
                <div class="chart-container">
                    <div class="chart-title">🏥 Distribuição por Modalidade</div>
                    <div class="chart-wrapper">
                        <canvas id="chartModalidade"></canvas>
                    </div>
                </div>
                <div class="chart-container">
                    <div class="chart-title">📊 Quantidade por Modalidade</div>
                    <div class="chart-wrapper">
                        <canvas id="chartModalidadeQtd"></canvas>
                    </div>
                </div>
            </div>
            
            <!-- Cards por Modalidade -->
            <div class="modalidade-cards" id="modalidadeCards"></div>
            
            <!-- Tabela Detalhada -->
            <div class="table-container">
                <div class="table-title">📋 Detalhamento por Mês, Convênio e Modalidade</div>
                <table>
                    <thead>
                        <tr>
                            <th>Mês</th>
                            <th>Convênio</th>
                            <th>Modalidade</th>
                            <th>Quantidade</th>
                            <th>Receita Bruta</th>
                            <th>Receita Líquida</th>
                        </tr>
                    </thead>
                    <tbody id="tableBody">
                    </tbody>
                </table>
            </div>
        </div>
    </div>
    
    <div class="container">
        <div class="footer">
            <p>✅ Dashboard atualizado automaticamente | Dados processados de forma segura</p>
        </div>
    </div>
    
    <script>
        // Dados embutidos
        const allData = {data_json};
        let filteredData = JSON.parse(JSON.stringify(allData));
        
        // Inicializar
        document.addEventListener('DOMContentLoaded', function() {{
            preencherFiltros();
            atualizarDashboard();
        }});
        
        function preencherFiltros() {{
            // Convênios
            const convenios = [...new Set(allData.exames_mes_convenio_modalidade.map(x => x.Convenio))];
            const selectConvenio = document.getElementById('filterConvenio');
            convenios.forEach(c => {{
                const option = document.createElement('option');
                option.value = c;
                option.textContent = c;
                selectConvenio.appendChild(option);
            }});
            
            // Modalidades
            const modalidades = [...new Set(allData.dist_modalidade.map(x => x.Modalidade))];
            const filterModalidade = document.getElementById('filterModalidade');
            modalidades.forEach(m => {{
                const div = document.createElement('div');
                div.className = 'checkbox-item';
                const checkbox = document.createElement('input');
                checkbox.type = 'checkbox';
                checkbox.value = m;
                checkbox.onchange = aplicarFiltros;
                const label = document.createElement('label');
                label.style.marginBottom = '0';
                label.appendChild(checkbox);
                label.appendChild(document.createTextNode(m));
                div.appendChild(label);
                filterModalidade.appendChild(div);
            }});
            
            // Períodos
            const periodos = [...new Set(allData.resumo_mensal.map(x => x.Mes))];
            const selectMes = document.getElementById('filterMes');
            periodos.forEach(p => {{
                if (p) {{
                    const option = document.createElement('option');
                    option.value = p;
                    option.textContent = p;
                    selectMes.appendChild(option);
                }}
            }});
        }}
        
        function aplicarFiltros() {{
            const convenio = document.getElementById('filterConvenio').value;
            const mes = document.getElementById('filterMes').value;
            const modalidades = [];
            document.querySelectorAll('#filterModalidade input[type="checkbox"]:checked').forEach(cb => {{
                modalidades.push(cb.value);
            }});
            
            filteredData = JSON.parse(JSON.stringify(allData));
            
            if (convenio) {{
                filteredData.exames_mes_convenio_modalidade = filteredData.exames_mes_convenio_modalidade.filter(x => x.Convenio === convenio);
            }}
            
            if (mes) {{
                filteredData.exames_mes_convenio_modalidade = filteredData.exames_mes_convenio_modalidade.filter(x => x.Mes === mes);
            }}
            
            if (modalidades.length > 0) {{
                filteredData.exames_mes_convenio_modalidade = filteredData.exames_mes_convenio_modalidade.filter(x => modalidades.includes(x.Modalidade));
            }}
            
            atualizarDashboard();
        }}
        
        function limparFiltros() {{
            document.getElementById('filterConvenio').value = '';
            document.getElementById('filterMes').value = '';
            document.querySelectorAll('#filterModalidade input[type="checkbox"]').forEach(cb => cb.checked = false);
            filteredData = JSON.parse(JSON.stringify(allData));
            atualizarDashboard();
        }}
        
        function atualizarDashboard() {{
            atualizarKPIs();
            atualizarGraficos();
            atualizarTabela();
            atualizarCards();
        }}
        
        function atualizarKPIs() {{
            const dados = filteredData.exames_mes_convenio_modalidade;
            const qtd = dados.reduce((sum, x) => sum + x.Qtd, 0);
            const bruta = dados.reduce((sum, x) => sum + x.Receita_Bruta, 0);
            const liquida = dados.reduce((sum, x) => sum + x.Receita_Liquida, 0);
            const lucro = bruta > 0 ? ((liquida / bruta) * 100).toFixed(2) : 0;
            
            document.getElementById('kpiExames').textContent = qtd.toLocaleString('pt-BR');
            document.getElementById('kpiReceitaBruta').textContent = 'R$ ' + bruta.toLocaleString('pt-BR', {{maximumFractionDigits: 2}});
            document.getElementById('kpiReceitaLiquida').textContent = 'R$ ' + liquida.toLocaleString('pt-BR', {{maximumFractionDigits: 2}});
            document.getElementById('kpiLucro').textContent = lucro + '%';
        }}
        
        function atualizarGraficos() {{
            // Gráfico de exames por mês
            const examesPosMes = {{}};
            filteredData.exames_mes_convenio_modalidade.forEach(x => {{
                examesPosMes[x.Mes] = (examesPosMes[x.Mes] || 0) + x.Qtd;
            }});
            
            const ctxMes = document.getElementById('chartExamesMes').getContext('2d');
            if (window.chartMes) window.chartMes.destroy();
            window.chartMes = new Chart(ctxMes, {{
                type: 'line',
                data: {{
                    labels: Object.keys(examesPosMes),
                    datasets: [{{
                        label: 'Exames',
                        data: Object.values(examesPosMes),
                        borderColor: '#ff6b35',
                        backgroundColor: 'rgba(255, 107, 53, 0.1)',
                        tension: 0.4,
                        fill: true
                    }}]
                }},
                options: {{responsive: true, maintainAspectRatio: false, plugins: {{legend: {{display: false}}}}}}
            }});
            
            // Gráfico de convênio
            const ctxConvenio = document.getElementById('chartConvenio').getContext('2d');
            if (window.chartConvenio) window.chartConvenio.destroy();
            const convenioDados = allData.dist_convenio.filter(x => x.Convenio !== 'CORTESIA');
            window.chartConvenio = new Chart(ctxConvenio, {{
                type: 'bar',
                data: {{
                    labels: convenioDados.map(x => x.Convenio),
                    datasets: [{{
                        label: 'Quantidade',
                        data: convenioDados.map(x => x.Qtd),
                        backgroundColor: '#1e3a5f'
                    }}]
                }},
                options: {{indexAxis: 'y', responsive: true, maintainAspectRatio: false, plugins: {{legend: {{display: false}}}}}}
            }});
            
            // Gráfico de modalidade (pizza)
            const ctxModalidade = document.getElementById('chartModalidade').getContext('2d');
            if (window.chartModalidade) window.chartModalidade.destroy();
            const modalidadeDados = allData.dist_modalidade.filter(x => x.Modalidade !== 'Outros');
            window.chartModalidade = new Chart(ctxModalidade, {{
                type: 'doughnut',
                data: {{
                    labels: modalidadeDados.map(x => x.Modalidade),
                    datasets: [{{
                        data: modalidadeDados.map(x => x.Qtd),
                        backgroundColor: ['#1e3a5f', '#ff6b35', '#2c5282', '#e55a24']
                    }}]
                }},
                options: {{responsive: true, maintainAspectRatio: false, plugins: {{legend: {{position: 'bottom'}}}}}}
            }});
            
            // Gráfico de modalidade (barras)
            const ctxModalidadeQtd = document.getElementById('chartModalidadeQtd').getContext('2d');
            if (window.chartModalidadeQtd) window.chartModalidadeQtd.destroy();
            window.chartModalidadeQtd = new Chart(ctxModalidadeQtd, {{
                type: 'bar',
                data: {{
                    labels: modalidadeDados.map(x => x.Modalidade),
                    datasets: [{{
                        label: 'Quantidade',
                        data: modalidadeDados.map(x => x.Qtd),
                        backgroundColor: '#ff6b35'
                    }}]
                }},
                options: {{responsive: true, maintainAspectRatio: false, plugins: {{legend: {{display: false}}}}}}
            }});
        }}
        
        function atualizarTabela() {{
            const tbody = document.getElementById('tableBody');
            tbody.innerHTML = '';
            filteredData.exames_mes_convenio_modalidade.forEach(row => {{
                const tr = document.createElement('tr');
                tr.innerHTML = `
                    <td>${{row.Mes}}</td>
                    <td>${{row.Convenio}}</td>
                    <td>${{row.Modalidade}}</td>
                    <td>${{row.Qtd}}</td>
                    <td>R$ ${{row.Receita_Bruta.toLocaleString('pt-BR', {{maximumFractionDigits: 2}})}}</td>
                    <td>R$ ${{row.Receita_Liquida.toLocaleString('pt-BR', {{maximumFractionDigits: 2}})}}</td>
                `;
                tbody.appendChild(tr);
            }});
        }}
        
        function atualizarCards() {{
            const container = document.getElementById('modalidadeCards');
            container.innerHTML = '';
            allData.dist_modalidade.filter(x => x.Modalidade !== 'Outros').forEach(mod => {{
                const card = document.createElement('div');
                card.className = 'modalidade-card';
                card.innerHTML = `
                    <h3>${{mod.Modalidade}}</h3>
                    <div class="modalidade-stat">
                        <span>Quantidade:</span>
                        <span>${{mod.Qtd}}</span>
                    </div>
                    <div class="modalidade-stat">
                        <span>Receita Bruta:</span>
                        <span>R$ ${{mod.Receita_Bruta.toLocaleString('pt-BR', {{maximumFractionDigits: 2}})}}</span>
                    </div>
                    <div class="modalidade-stat">
                        <span>Receita Líquida:</span>
                        <span>R$ ${{mod.Receita_Liquida.toLocaleString('pt-BR', {{maximumFractionDigits: 2}})}}</span>
                    </div>
                `;
                container.appendChild(card);
            }});
        }}
    </script>
</body>
</html>
"""
    
    # Salvar HTML
    with open('dashboard.html', 'w', encoding='utf-8') as f:
        f.write(html_content)
    
    print("✅ HTML gerado com sucesso: dashboard.html")

def main():
    """Executar pipeline completo"""
    print("="*60)
    print("🚀 Iniciando atualização automática do Dashboard SASE")
    print("="*60)
    
    try:
        # Passo 1: Baixar arquivo
        print("\n📥 Etapa 1: Baixando arquivo do Google Drive...")
        file_content = baixar_excel_google_drive()
        
        # Passo 2: Processar dados
        print("\n⚙️  Etapa 2: Processando dados...")
        data_json = processar_dados(file_content)
        
        # Passo 3: Gerar HTML
        print("\n🎨 Etapa 3: Gerando HTML...")
        gerar_html(data_json)
        
        print("\n" + "="*60)
        print("✅ SUCESSO! Dashboard atualizado com sucesso!")
        print("="*60)
        print("\n📊 Resumo:")
        print(f"   - Arquivo salvo: dashboard.html")
        print(f"   - Última atualização: {data_json['ultima_atualizacao']}")
        print(f"   - Próxima atualização: Amanhã às 23h (Rio)")
        print("\n")
        
    except Exception as e:
        print(f"\n❌ ERRO: {str(e)}")
        exit(1)

if __name__ == '__main__':
    main()
