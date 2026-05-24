"""
VALIDADOR FINANCEIRO CAMPAX — Flask Backend para Railway
Deploy em: https://railway.app
"""

from flask import Flask, render_template, request, jsonify, send_file
from werkzeug.utils import secure_filename
import pandas as pd
import json
import os
from datetime import datetime
from pathlib import Path
import io

app = Flask(__name__, static_folder='static', template_folder='templates')
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max

# Database simulado (JSON) — depois pode migrar para PostgreSQL
HISTORICO_FILE = 'historico.json'
UPLOAD_FOLDER = '/tmp/uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

CATEGORIAS = {
    '1': 'Despesas Fixas',
    '37': 'Funcionários',
    '38': 'Funcionários Extra',
    '39': 'Fretes/Transportes',
    '41': 'Matéria Prima',
    '43': 'Taxas Bancárias',
    '44': 'Impostos',
    '80': 'Diversos',
    'REC': 'Receita'
}

# ═══════════════════════════════════════════════════════════════════════════════

def carregar_historico():
    """Carrega histórico de validações"""
    if os.path.exists(HISTORICO_FILE):
        with open(HISTORICO_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {'validacoes': []}

def salvar_historico(dados):
    """Salva histórico de validações"""
    with open(HISTORICO_FILE, 'w', encoding='utf-8') as f:
        json.dump(dados, f, ensure_ascii=False, indent=2)

def extrair_mes(nomeArquivo):
    """Extrai nome do mês do arquivo"""
    meses = ['janeiro', 'fevereiro', 'março', 'abril', 'maio', 'junho',
             'julho', 'agosto', 'setembro', 'outubro', 'novembro', 'dezembro']
    arquivo_lower = nomeArquivo.lower()
    for mes in meses:
        if mes in arquivo_lower:
            return mes.capitalize()
    return 'Desconhecido'

def processar_excel(arquivo):
    """Processa arquivo Excel e retorna dados estruturados"""
    try:
        df = pd.read_excel(arquivo, sheet_name=0, header=None)
    except Exception as e:
        return None, f"Erro ao ler Excel: {str(e)}"
    
    receitas = []
    despesas = []
    
    # Processa a partir da linha 5 (após cabeçalhos)
    for idx, row in df.iterrows():
        if idx < 5:
            continue
        
        data = row[0]
        codigo = row[1]
        descricao = row[2]
        entrada = float(row[3]) if pd.notna(row[3]) else 0
        saida = float(row[4]) if pd.notna(row[4]) else 0
        
        if pd.isna(data) or (entrada == 0 and saida == 0):
            continue
        
        # Trata receitas
        if codigo == 'REC' and entrada > 0:
            receitas.append({
                'data': str(data)[:10] if hasattr(data, '__str__') else str(data),
                'descricao': str(descricao),
                'valor': round(entrada, 2)
            })
        
        # Trata despesas
        elif codigo != 'REC' and saida > 0:
            despesas.append({
                'data': str(data)[:10] if hasattr(data, '__str__') else str(data),
                'codigo': str(codigo),
                'descricao': str(descricao),
                'valor': round(saida, 2),
                'categoria': CATEGORIAS.get(str(codigo), 'Não categorizado')
            })
    
    # Calcula totais
    total_rec = sum(r['valor'] for r in receitas)
    total_desp = sum(d['valor'] for d in despesas)
    resultado = total_rec - total_desp
    margem = (resultado / total_rec * 100) if total_rec > 0 else 0
    
    return {
        'receitas': receitas,
        'despesas': despesas,
        'totais': {
            'receita': round(total_rec, 2),
            'despesa': round(total_desp, 2),
            'resultado': round(resultado, 2),
            'margem': round(margem, 2)
        },
        'contadores': {
            'receitas_count': len(receitas),
            'despesas_count': len(despesas)
        }
    }, None

# ═══════════════════════════════════════════════════════════════════════════════

@app.route('/')
def index():
    """Página principal"""
    return render_template('index.html')

@app.route('/api/processar', methods=['POST'])
def api_processar():
    """Endpoint para processar arquivo Excel"""
    
    if 'file' not in request.files:
        return jsonify({'erro': 'Arquivo não enviado'}), 400
    
    arquivo = request.files['file']
    
    if arquivo.filename == '':
        return jsonify({'erro': 'Arquivo vazio'}), 400
    
    if not arquivo.filename.endswith(('.xlsx', '.xls')):
        return jsonify({'erro': 'Apenas .xlsx e .xls são aceitos'}), 400
    
    # Processa arquivo
    dados, erro = processar_excel(arquivo)
    
    if erro:
        return jsonify({'erro': erro}), 400
    
    # Extrai mês
    mes = extrair_mes(arquivo.filename)
    
    # Cria registro
    registro = {
        'mes': mes,
        'arquivo': secure_filename(arquivo.filename),
        'timestamp': datetime.now().isoformat(),
        'dados': dados
    }
    
    # Salva no histórico
    historico = carregar_historico()
    
    # Remove entrada anterior do mesmo mês (atualiza)
    historico['validacoes'] = [
        v for v in historico['validacoes'] 
        if v['mes'].lower() != mes.lower()
    ]
    
    historico['validacoes'].append(registro)
    salvar_historico(historico)
    
    return jsonify({
        'sucesso': True,
        'mes': mes,
        'dados': dados
    })

@app.route('/api/historico', methods=['GET'])
def api_historico():
    """Retorna histórico de todas as validações"""
    historico = carregar_historico()
    
    # Calcula resumo anual
    total_receita = sum(v['dados']['totais']['receita'] for v in historico['validacoes'])
    total_despesa = sum(v['dados']['totais']['despesa'] for v in historico['validacoes'])
    resultado = total_receita - total_despesa
    margem = (resultado / total_receita * 100) if total_receita > 0 else 0
    
    # Top 3 meses
    top_meses = sorted(
        historico['validacoes'],
        key=lambda x: x['dados']['totais']['resultado'],
        reverse=True
    )[:3]
    
    return jsonify({
        'validacoes': historico['validacoes'],
        'resumo': {
            'receita_total': round(total_receita, 2),
            'despesa_total': round(total_despesa, 2),
            'resultado_total': round(resultado, 2),
            'margem_media': round(margem, 2),
            'quantidade_meses': len(historico['validacoes']),
            'top_3_meses': [
                {
                    'mes': m['mes'],
                    'resultado': m['dados']['totais']['resultado'],
                    'margem': m['dados']['totais']['margem']
                }
                for m in top_meses
            ]
        }
    })

@app.route('/api/mes/<mes_nome>', methods=['GET'])
def api_mes_detalhes(mes_nome):
    """Retorna detalhes de um mês específico"""
    historico = carregar_historico()
    
    for validacao in historico['validacoes']:
        if validacao['mes'].lower() == mes_nome.lower():
            return jsonify(validacao['dados'])
    
    return jsonify({'erro': 'Mês não encontrado'}), 404

@app.route('/api/exportar', methods=['GET'])
def api_exportar():
    """Exporta relatório em JSON"""
    historico = carregar_historico()
    
    # Cria arquivo
    buffer = io.BytesIO()
    buffer.write(json.dumps(historico, ensure_ascii=False, indent=2).encode('utf-8'))
    buffer.seek(0)
    
    return send_file(
        buffer,
        mimetype='application/json',
        as_attachment=True,
        download_name=f'relatorio_campax_{datetime.now().strftime("%Y%m%d")}.json'
    )

@app.route('/api/saude', methods=['GET'])
def api_saude():
    """Health check"""
    historico = carregar_historico()
    return jsonify({
        'status': 'ok',
        'meses_processados': len(historico['validacoes']),
        'timestamp': datetime.now().isoformat()
    })

# ═══════════════════════════════════════════════════════════════════════════════

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
