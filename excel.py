import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

def aplicar_estilo(ws, headers):
    ws.append(headers)
    fundo = PatternFill(start_color="2563EB", end_color="2563EB", fill_type="solid")
    fonte = Font(color="FFFFFF", bold=True)
    for cell in ws[1]:
        cell.font = fonte
        cell.fill = fundo
        cell.alignment = Alignment(horizontal="center")
        ws.column_dimensions[cell.column_letter].width = 25

def gerar_planilha(empresa, pendencias, socios, certidoes, pgfn):
    wb = openpyxl.Workbook()
    
    # Aba 1: Resumo Empresa
    ws_empresa = wb.active
    ws_empresa.title = "Empresa"
    aplicar_estilo(ws_empresa, ["CNPJ", "Razão Social", "Situação", "Porte", "Status PGFN"])
    ws_empresa.append([empresa.get('cnpj', ''), empresa.get('razao_social', ''), empresa.get('situacao', ''), empresa.get('porte', ''), pgfn.get('status', '')])
    
    # Aba 2: Pendências
    ws_pend = wb.create_sheet(title="Pendências")
    aplicar_estilo(ws_pend, ["Órgão", "Tipo", "Período", "Status"])
    for p in pendencias:
        ws_pend.append([p.get('orgao'), p.get('tipo'), p.get('periodo'), p.get('status')])
        
    # Aba 3: Sócios
    ws_socios = wb.create_sheet(title="Sócios")
    aplicar_estilo(ws_socios, ["Documento", "Nome", "Qualificação"])
    for s in socios:
        ws_socios.append([s.get('documento'), s.get('nome'), s.get('qualificacao')])
        
    # Aba 4: Certidões
    ws_cert = wb.create_sheet(title="Certidões")
    aplicar_estilo(ws_cert, ["Tipo", "Código", "Emissão", "Validade"])
    for c in certidoes:
        ws_cert.append([c.get('tipo'), c.get('codigo'), c.get('emissao'), c.get('validade')])
    
    nome_arquivo = f"Relatorio_{empresa.get('cnpj', 'erro').replace('/', '').replace('.', '').replace('-', '')}.xlsx"
    wb.save(nome_arquivo)
    return nome_arquivo