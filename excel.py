import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.table import Table, TableStyleInfo

def ajustar_largura(ws):
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                if len(str(cell.value)) > max_length: max_length = len(str(cell.value))
            except: pass
        adjusted_width = (max_length + 2) * 1.2
        ws.column_dimensions[column].width = adjusted_width if adjusted_width > 15 else 15

def criar_tabela(ws, range_ref, table_name, style="TableStyleMedium2"):
    tab = Table(displayName=table_name, ref=range_ref)
    style = TableStyleInfo(name=style, showFirstColumn=False, showLastColumn=False, showRowStripes=True, showColumnStripes=False)
    tab.tableStyleInfo = style
    ws.add_table(tab)

def gerar_planilha(dados):
    wb = openpyxl.Workbook()
    
    # Aba 1: Empresa & Simples Nacional
    ws_empresa = wb.active
    ws_empresa.title = "Dados Cadastrais"
    ws_empresa.append(["CNPJ", "Razão Social", "Situação", "Porte"])
    ws_empresa.append([dados['empresa'].get('cnpj'), dados['empresa'].get('razao_social'), dados['empresa'].get('situacao'), dados['empresa'].get('porte')])
    criar_tabela(ws_empresa, "A1:D2", "TabelaEmpresa", "TableStyleDark1")
    
    ws_empresa.append([])
    ws_empresa.append(["Evento", "Data Inclusão", "Data Exclusão"])
    row_count = 4
    for s in dados.get('simples_nacional', []):
        ws_empresa.append([s['evento'], s['inclusao'], s['exclusao']])
        row_count += 1
    if row_count > 4:
        criar_tabela(ws_empresa, f"A4:C{row_count-1}", "TabelaSimples", "TableStyleMedium2")
    ajustar_largura(ws_empresa)

    # Aba 2: Sócios
    ws_socios = wb.create_sheet(title="Quadro Societário")
    ws_socios.append(["CPF/CNPJ", "Nome do Sócio", "Qualificação"])
    row_count = 1
    for s in dados.get('socios', []):
        ws_socios.append([s['documento'], s['nome'], s['qualificacao']])
        row_count += 1
    if row_count > 1: criar_tabela(ws_socios, f"A1:C{row_count}", "TabelaSocios")
    ajustar_largura(ws_socios)

    # Aba 3: Pendências RFB Detalhadas
    ws_pend = wb.create_sheet(title="Pendências RFB")
    ws_pend.append(["Órgão", "Tipo / Tributo", "Período / Info", "Valor Original", "Multa", "Juros", "Saldo Consolidado", "Status Atual"])
    row_count = 1
    for p in dados.get('pendencias', []):
        ws_pend.append([p['orgao'], p['tipo'], p['periodo'], p['vl_original'], p['multa'], p['juros'], p['vl_consolidado'], p['status']])
        row_count += 1
    if row_count > 1: criar_tabela(ws_pend, f"A1:H{row_count}", "TabelaPendencias", "TableStyleMedium9")
    ajustar_largura(ws_pend)

    # Aba 4: PGFN (Dívida Ativa)
    ws_pgfn = wb.create_sheet(title="Dívida Ativa PGFN")
    ws_pgfn.append(["Inscrição PGFN", "Origem / Tributo", "Situação"])
    row_count = 1
    if not dados.get('pgfn'):
        ws_pgfn.append(["-", "Nenhuma pendência PGFN detectada", "-"])
        row_count += 1
    else:
        for pgfn in dados.get('pgfn', []):
            ws_pgfn.append([pgfn['inscricao'], pgfn['tributo'], pgfn['status']])
            row_count += 1
    criar_tabela(ws_pgfn, f"A1:C{row_count}", "TabelaPGFN", "TableStyleMedium3")
    ajustar_largura(ws_pgfn)

    # Aba 5: Certidões
    ws_cert = wb.create_sheet(title="Certidões")
    ws_cert.append(["Tipo", "Código Controle", "Emissão", "Validade"])
    row_count = 1
    for c in dados.get('certidoes', []):
        ws_cert.append([c['tipo'], c['codigo'], c['emissao'], c['validade']])
        row_count += 1
    if row_count > 1: criar_tabela(ws_cert, f"A1:D{row_count}", "TabelaCertidoes")
    ajustar_largura(ws_cert)

    nome_arquivo = f"SitFiscal_{dados['empresa'].get('cnpj', '000').replace('/', '').replace('.', '').replace('-', '')}.xlsx"
    wb.save(nome_arquivo)
    return nome_arquivo