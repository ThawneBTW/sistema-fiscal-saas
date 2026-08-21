import openpyxl
from openpyxl.styles import Font, PatternFill

def gerar_planilha(empresa, pendencias):
    wb = openpyxl.Workbook()
    
    # Aba Empresa
    ws_empresa = wb.active
    ws_empresa.title = "Resumo Empresa"
    ws_empresa.append(["CNPJ", "Razão Social", "Situação"])
    ws_empresa.append([empresa.get('cnpj', ''), empresa.get('razao_social', ''), empresa.get('situacao', '')])
    
    # Aba Pendências
    ws_pend = wb.create_sheet(title="Pendências")
    headers = ["Órgão", "Tipo", "Período", "Status"]
    ws_pend.append(headers)
    
    fundo_azul = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    for celula in ws_pend[1]:
        celula.font = Font(color="FFFFFF", bold=True)
        celula.fill = fundo_azul

    for p in pendencias:
        ws_pend.append([p.get('orgao'), p.get('tipo'), p.get('periodo'), p.get('status')])
    
    nome_arquivo = f"Relatorio_{empresa.get('cnpj', 'erro').replace('/', '').replace('.', '').replace('-', '')}.xlsx"
    wb.save(nome_arquivo)
    return nome_arquivo