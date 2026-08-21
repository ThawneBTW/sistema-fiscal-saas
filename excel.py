import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

def converter_para_numero(valor_str):
    if not valor_str or valor_str in ("-", "Consultar", "Ver Detalhes", ""): return 0.0
    try:
        return float(valor_str.replace('.', '').replace(',', '.'))
    except:
        return 0.0

def estilizar_cabecalho(ws, headers, linha=3):
    ws.append(headers)
    fundo = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    fonte = Font(color="FFFFFF", bold=True)
    
    for i, cell in enumerate(ws[linha], 1):
        cell.font = fonte
        cell.fill = fundo
        cell.alignment = Alignment(horizontal="center", vertical="center")
    
    ws.freeze_panes = f"A{linha+1}"
    ws.auto_filter.ref = f"A{linha}:{openpyxl.utils.get_column_letter(len(headers))}{ws.max_row}"

def ajustar_largura(ws):
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                if len(str(cell.value)) > max_length: max_length = len(str(cell.value))
            except: pass
        adjusted_width = (max_length + 2) * 1.1
        ws.column_dimensions[column].width = adjusted_width if adjusted_width > 15 else 15

def gerar_planilha(dados):
    wb = openpyxl.Workbook()
    FORMATO_MOEDA = 'R$ #,##0.00'
    
    # ---------------------------------------------------------
    # ABA 1: TODOS OS DÉBITOS (UNIFICADA E ANALÍTICA)
    # ---------------------------------------------------------
    ws_debitos = wb.active
    ws_debitos.title = "Todos os Débitos"
    
    # Identificação da Empresa no Topo
    cnpj = dados['empresa'].get('cnpj', 'N/A')
    razao = dados['empresa'].get('razao_social', 'N/A')
    ws_debitos["A1"] = f"CNPJ: {cnpj} - {razao}"
    ws_debitos["A1"].font = Font(size=14, bold=True, color="0F172A")
    ws_debitos.merge_cells("A1:H1")
    ws_debitos.append([]) # Linha em branco
    
    # Cabeçalho da Tabela
    cabecalhos = ["Órgão", "Natureza / Tributo / Inscrição", "Período Apuração", "Valor Original", "Multa", "Juros", "Saldo Consolidado", "Status Atual"]
    estilizar_cabecalho(ws_debitos, cabecalhos, linha=3)
    
    row_idx = 4
    total_vencido = 0.0
    total_analisar = 0.0
    
    # Adicionando Débitos da Receita Federal e Parcelamentos
    for p in dados.get('pendencias', []):
        v_orig = converter_para_numero(p['vl_original'])
        v_multa = converter_para_numero(p['multa'])
        v_juros = converter_para_numero(p['juros'])
        v_cons = converter_para_numero(p['vl_consolidado'])
        
        status = str(p['status']).upper()
        # Regra de classificação de risco
        is_devedor = any(term in status for term in ["DEVEDOR", "ATRASO", "ATRASADA"])
        
        if is_devedor: total_vencido += v_cons
        else: total_analisar += v_cons
        
        ws_debitos.append([p['orgao'], p['tipo'], p['periodo'], v_orig, v_multa, v_juros, v_cons, p['status']])
        
        # Formatação Visual
        for col in ['D', 'E', 'F', 'G']: ws_debitos[f"{col}{row_idx}"].number_format = FORMATO_MOEDA
        
        status_cell = ws_debitos[f"H{row_idx}"]
        if is_devedor:
            status_cell.font = Font(color="991B1B", bold=True)
            status_cell.fill = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
        elif any(term in status for term in ["A ANALISAR", "A VENCER", "SUSPENSA", "PARCELAMENTO"]):
            status_cell.font = Font(color="0F766E", bold=True)
            status_cell.fill = PatternFill(start_color="CCFBF1", end_color="CCFBF1", fill_type="solid")
            
        row_idx += 1

    # Adicionando Dívida Ativa (PGFN) na MESMA ABA
    for pgfn in dados.get('pgfn', []):
        ws_debitos.append(["PGFN", f"Dívida Ativa: {pgfn['tributo']}", pgfn['inscricao'], "-", "-", "-", "Consultar", pgfn['status']])
        status_cell = ws_debitos[f"H{row_idx}"]
        status_cell.font = Font(color="991B1B", bold=True)
        status_cell.fill = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
        row_idx += 1

    # --- QUADRO DE RESUMO FINANCEIRO NO FINAL DA TABELA ---
    ws_debitos.append([])
    
    row_idx += 1
    ws_debitos[f"F{row_idx}"] = "TOTAIS COM DÉBITO / VENCIDO:"
    ws_debitos[f"F{row_idx}"].font = Font(bold=True, color="991B1B")
    ws_debitos[f"G{row_idx}"] = total_vencido
    ws_debitos[f"G{row_idx}"].number_format = FORMATO_MOEDA
    ws_debitos[f"G{row_idx}"].font = Font(bold=True)
    
    row_idx += 1
    ws_debitos[f"F{row_idx}"] = "TOTAIS A ANALISAR / A VENCER:"
    ws_debitos[f"F{row_idx}"].font = Font(bold=True, color="0F766E")
    ws_debitos[f"G{row_idx}"] = total_analisar
    ws_debitos[f"G{row_idx}"].number_format = FORMATO_MOEDA
    ws_debitos[f"G{row_idx}"].font = Font(bold=True)
    
    row_idx += 1
    ws_debitos[f"F{row_idx}"] = "DÍVIDA ESTIMADA TOTAL:"
    ws_debitos[f"F{row_idx}"].font = Font(bold=True, size=12)
    ws_debitos[f"G{row_idx}"] = total_vencido + total_analisar
    ws_debitos[f"G{row_idx}"].number_format = FORMATO_MOEDA
    ws_debitos[f"G{row_idx}"].font = Font(bold=True, size=12)

    ajustar_largura(ws_debitos)

    # ---------------------------------------------------------
    # ABA 2: DADOS CADASTRAIS (Socios, Simples, Certidões)
    # ---------------------------------------------------------
    ws_cad = wb.create_sheet(title="Ficha Cadastral")
    
    ws_cad.append(["QUADRO SOCIETÁRIO"])
    estilizar_cabecalho(ws_cad, ["CPF/CNPJ", "Nome do Sócio", "Qualificação"], linha=2)
    for s in dados.get('socios', []): ws_cad.append([s['documento'], s['nome'], s['qualificacao']])
    ws_cad.append([])
    
    ws_cad.append(["HISTÓRICO SIMPLES NACIONAL"])
    linha_simples = ws_cad.max_row + 1
    estilizar_cabecalho(ws_cad, ["Evento", "Data Inclusão", "Data Exclusão"], linha=linha_simples)
    if not dados.get('simples_nacional'): ws_cad.append(["-", "Sem histórico detectado", "-"])
    else:
        for s in dados.get('simples_nacional', []): ws_cad.append([s['evento'], s['inclusao'], s['exclusao']])
    ws_cad.append([])
    
    ws_cad.append(["CERTIDÕES EMITIDAS"])
    linha_cert = ws_cad.max_row + 1
    estilizar_cabecalho(ws_cad, ["Tipo de Certidão", "Código de Controle", "Emissão", "Validade"], linha=linha_cert)
    if not dados.get('certidoes'): ws_cad.append(["-", "Nenhuma certidão localizada", "-", "-"])
    else:
        for c in dados.get('certidoes', []): ws_cad.append([c['tipo'], c['codigo'], c['emissao'], c['validade']])
        
    ajustar_largura(ws_cad)

    cnpj_limpo = dados['empresa'].get('cnpj', '000').replace('/', '').replace('.', '').replace('-', '')
    nome_arquivo = f"SitFiscal_{cnpj_limpo}.xlsx"
    wb.save(nome_arquivo)
    return nome_arquivo