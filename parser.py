import pdfplumber
import re

class RelatorioParser:
    def __init__(self, caminho_pdf):
        self.caminho_pdf = caminho_pdf
        self.texto = ""
        self._ler_pdf()

    def _ler_pdf(self):
        with pdfplumber.open(self.caminho_pdf) as pdf:
            for pagina in pdf.pages:
                texto_pagina = pagina.extract_text(x_tolerance=2, y_tolerance=2)
                if texto_pagina:
                    self.texto += texto_pagina + "\n"

    def extrair_dados(self):
        return {
            "empresa": self._extrair_empresa(),
            "simples_nacional": self._extrair_simples(),
            "socios": self._extrair_socios(),
            "pendencias": self._extrair_pendencias(),
            "pgfn": self._extrair_pgfn(),
            "certidoes": self._extrair_certidoes()
        }

    def _extrair_empresa(self):
        dados = {}
        match_cnpj = re.search(r'CNPJ:\s*(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})', self.texto)
        if match_cnpj: dados['cnpj'] = match_cnpj.group(1)

        match_razao = re.search(r'CNPJ:\s*\d{2}\.\d{3}\.\d{3}\s+(.+?)(?=\n)', self.texto)
        if match_razao: dados['razao_social'] = match_razao.group(1).strip()

        sit_match = re.search(r'Situação:\s*([A-Z\s]+?)(?=\n|Natureza)', self.texto)
        if sit_match: dados['situacao'] = sit_match.group(1).strip()
        
        porte = re.search(r'Porte da Empresa:\s*(.+)', self.texto)
        if porte: dados['porte'] = porte.group(1).strip()
        
        return dados

    def _extrair_simples(self):
        historico = []
        bloco_simples = re.search(r'Opção pelo Simples Nacional\n(.*?)(?=Opção pelo SIMEI|Sócios e Administradores)', self.texto, re.DOTALL)
        if bloco_simples:
            for linha in bloco_simples.group(1).split('\n'):
                datas = re.findall(r'\d{2}/\d{2}/\d{4}', linha)
                if len(datas) == 2:
                    historico.append({"evento": "Inclusão/Exclusão", "inclusao": datas[0], "exclusao": datas[1]})
                elif len(datas) == 1:
                    historico.append({"evento": "Ativo", "inclusao": datas[0], "exclusao": "Atual"})
        return historico

    def _extrair_socios(self):
        socios = []
        matches = re.finditer(r'(\d{3}\.\d{3}\.\d{3}-\d{2})\n([A-ZÀ-Ú\s]+)\n', self.texto)
        for m in matches:
            nome = m.group(2).strip()
            if "CERTIDÃO" not in nome and "CÓDIGO" not in nome:
                socios.append({"documento": m.group(1), "nome": nome, "qualificacao": "SÓCIO-ADMINISTRADOR"})
        return socios

    def _extrair_pendencias(self):
        pendencias = []
        secao = None

        for linha in self.texto.split('\n'):
            linha = linha.strip()
            if not linha: continue

            if "Omissão de DCTFWeb" in linha: secao = "DCTFWEB"; continue
            if "Débito (SIEF)" in linha or "Débito com Exigibilidade Suspensa" in linha: secao = "SIEF"; continue
            if "Pendência Processo Fiscal" in linha: secao = "PROCESSO"; continue
            if "Parcelamento com Exigibilidade Suspensa" in linha or "Pendência Parcelamento" in linha: secao = "PARCELAMENTO"; continue
            if "Diagnóstico Fiscal na Procuradoria" in linha or "Final do Relatório" in linha: secao = "FIM"

            if secao == "DCTFWEB":
                match = re.search(r'^(\d{4})\s+([A-Z\s]+)', linha)
                if match:
                    meses = match.group(2).strip().split()
                    pendencias.append({
                        "orgao": "RFB", "tipo": "Omissão de DCTFWeb", 
                        "periodo": f"{match.group(1)} - {', '.join(meses)}", 
                        "vl_original": "-", "multa": "-", "juros": "-", "vl_consolidado": "-", 
                        "status": "Pendente"
                    })
                    secao = None 
                    
            elif secao == "SIEF":
                periodo_match = re.search(r'(\d{2}/\d{4})', linha)
                status_match = re.search(r'(DEVEDOR|A ANALISAR|EXIGIBILIDADE SUSPENSA)', linha)
                if periodo_match and status_match:
                    tributo = linha.split(periodo_match.group(1))[0].replace('|', '').strip()
                    
                    # Extração detalhada de todos os valores na linha
                    valores = re.findall(r'\d{1,3}(?:\.\d{3})*,\d{2}', linha)
                    vl_orig = valores[0] if len(valores) > 0 else "-"
                    multa = valores[2] if len(valores) >= 4 else "-"
                    juros = valores[3] if len(valores) >= 4 else "-"
                    vl_cons = valores[-1] if valores else "-"

                    pendencias.append({
                        "orgao": "RFB", "tipo": f"Débito SIEF: {tributo}", 
                        "periodo": periodo_match.group(1), 
                        "vl_original": vl_orig, "multa": multa, "juros": juros, "vl_consolidado": vl_cons, 
                        "status": status_match.group(1)
                    })

            elif secao == "PROCESSO":
                match = re.search(r'([\d\.\-\/]+)\s+(DEVEDOR|EM ANALISE)', linha)
                if match:
                    pendencias.append({
                        "orgao": "RFB", "tipo": f"Processo Fiscal: {match.group(1)}", 
                        "periodo": "-", "vl_original": "-", "multa": "-", "juros": "-", "vl_consolidado": "-", 
                        "status": match.group(2)
                    })

            elif secao == "PARCELAMENTO":
                if "Parcelamento:" in linha or "Conta" in linha:
                    num = linha.replace("Parcelamento:", "").replace("Conta", "").strip()
                    pendencias.append({
                        "orgao": "RFB", "tipo": "Parcelamento RFB", 
                        "periodo": num, "vl_original": "-", "multa": "-", "juros": "-", "vl_consolidado": "-", 
                        "status": "Em Parcelamento"
                    })
                elif "Valor em Atraso:" in linha:
                    val_atraso = re.search(r'Valor em Atraso:\s*([\d\.,]+)', linha)
                    if val_atraso and pendencias and pendencias[-1]["tipo"] == "Parcelamento RFB":
                        pendencias[-1]["vl_consolidado"] = val_atraso.group(1)
                        pendencias[-1]["status"] = "Atrasado"

        return pendencias

    def _extrair_pgfn(self):
        pgfn_lista = []
        secao_pgfn = False
        for linha in self.texto.split('\n'):
            if "Diagnóstico Fiscal na Procuradoria" in linha: secao_pgfn = True; continue
            if "Final do Relatório" in linha: break
            if secao_pgfn:
                match_insc = re.search(r'(\d{2}\.\d\.\d{2}\.\d{6}-\d{2})\s+([A-Z0-9\-\s]+?)(?=\d{2}/)', linha)
                if match_insc:
                    pgfn_lista.append({"inscricao": match_insc.group(1), "tributo": match_insc.group(2).strip(), "status": "Inscrito em Dívida Ativa"})
        return pgfn_lista

    def _extrair_certidoes(self):
        certidoes = []
        if "Certidão Emitida" in self.texto:
            cert = re.search(r'Certidão (.*?):\s*([A-Z0-9\.]+)\nEmissão:\s*([\d/]+)\nData de Validade:\s*([\d/]+)', self.texto)
            if cert: certidoes.append({"tipo": cert.group(1).strip(), "codigo": cert.group(2).strip(), "emissao": cert.group(3).strip(), "validade": cert.group(4).strip()})
        return certidoes