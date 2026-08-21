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
                texto_pagina = pagina.extract_text()
                if texto_pagina:
                    self.texto += texto_pagina + "\n"

    def extrair_dados(self):
        return {
            "empresa": self._extrair_empresa(),
            "socios": self._extrair_socios(),
            "pendencias": self._extrair_pendencias(),
            "certidoes": self._extrair_certidoes(),
            "pgfn": self._extrair_pgfn()
        }

    def _extrair_empresa(self):
        dados = {}
        # Captura o CNPJ raiz (apenas números iniciais) e a Razão Social que vem logo em seguida
        match_base = re.search(r'CNPJ:\s*(\d{2}\.\d{3}\.\d{3})\s+(.+)', self.texto)
        if match_base:
            dados['razao_social'] = match_base.group(2).strip()
        
        # Captura o CNPJ Completo com barra e traço
        match_completo = re.search(r'CNPJ:\s*(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})', self.texto)
        if match_completo:
            dados['cnpj'] = match_completo.group(1)

        # Captura Situação isolando quebras de linha
        sit_match = re.search(r'Situação:\s*([A-Z\s]+?)(?=\n|Natureza|CEP)', self.texto)
        if sit_match: 
            dados['situacao'] = sit_match.group(1).strip()
        
        porte = re.search(r'Porte da Empresa:\s*(.+)', self.texto)
        if porte: 
            dados['porte'] = porte.group(1).strip()
        
        return dados

    def _extrair_socios(self):
        socios = []
        # Tenta capturar o padrão onde o nome desce para a linha seguinte devido à quebra de coluna
        matches = re.finditer(r'(\d{3}\.\d{3}\.\d{3}-\d{2})\n([A-ZÀ-Ú\s]+)\n', self.texto)
        for m in matches:
            nome_encontrado = m.group(2).strip()
            # Filtro de segurança para não capturar títulos soltos do PDF
            if "CERTIDÃO" not in nome_encontrado and "CÓDIGO" not in nome_encontrado:
                socios.append({
                    "documento": m.group(1),
                    "nome": nome_encontrado,
                    "qualificacao": "SÓCIO/ADMINISTRADOR"
                })
        return socios

    def _extrair_pendencias(self):
        pendencias = []
        linhas = self.texto.split('\n')
        secao_atual = None

        for linha in linhas:
            linha = linha.strip()
            if not linha:
                continue

            # 1. Identificador de Seções (Máquina de Estados)
            if "Omissão de DCTFWeb" in linha:
                secao_atual = "DCTFWEB"
                continue
            elif "Débito (SIEF)" in linha:
                secao_atual = "SIEF_ATIVO"
                continue
            elif "Débito com Exigibilidade Suspensa" in linha:
                secao_atual = "SIEF_SUSPENSO"
                continue
            elif "Pendência Processo Fiscal" in linha:
                secao_atual = "PROCESSO"
                continue
            elif "Pendência Parcelamento" in linha:
                secao_atual = "PARCELAMENTO"
                continue
            elif "Diagnóstico Fiscal na Procuradoria" in linha or "Final do Relatório" in linha:
                secao_atual = "FIM"
                break

            # 2. Lógica de Extração baseada na Seção Atual
            if secao_atual == "DCTFWEB":
                match = re.search(r'(\d{4})\s+([A-Z]{3})', linha)
                if match:
                    pendencias.append({
                        "orgao": "Receita Federal",
                        "tipo": "Omissão de DCTFWeb",
                        "periodo": f"{match.group(2)}/{match.group(1)}",
                        "status": "Pendente"
                    })
                    secao_atual = None 

            elif secao_atual in ["SIEF_ATIVO", "SIEF_SUSPENSO"]:
                # Uma linha de débito SIEF sempre possui um período MM/YYYY e um status no fim
                periodo_match = re.search(r'(\d{2}/\d{4})', linha)
                status_match = re.search(r'(DEVEDOR|A ANALISAR|EXIGIBILIDADE SUSPENSA)', linha)
                
                if periodo_match and status_match:
                    # Limpa a formatação de tabela (pipes) para isolar o tipo do tributo
                    texto_limpo = linha.replace('|', '').strip()
                    tipo_tributo = texto_limpo.split(periodo_match.group(1))[0].strip()
                    
                    pendencias.append({
                        "orgao": "Receita Federal",
                        "tipo": f"Débito SIEF: {tipo_tributo}",
                        "periodo": periodo_match.group(1),
                        "status": status_match.group(1)
                    })

            elif secao_atual == "PROCESSO":
                # Captura processos no formato NNNNN.NNN.NNN/NNNN-NN e o respectivo status
                match = re.search(r'([\d\.\-\/]+)\s+(DEVEDOR|EM ANALISE)', linha)
                if match:
                    pendencias.append({
                        "orgao": "Receita Federal",
                        "tipo": f"Processo Fiscal: {match.group(1)}",
                        "periodo": "-",
                        "status": match.group(2)
                    })

            elif secao_atual == "PARCELAMENTO":
                # Captura dados da cobrança de parcelamento
                if "Parcelamento:" in linha:
                    num_parc = linha.split("Parcelamento:")[1].strip()
                    pendencias.append({
                        "orgao": "Receita Federal",
                        "tipo": f"Parcelamento (Atraso): {num_parc}",
                        "periodo": "-",
                        "status": "Em Atraso"
                    })
                    secao_atual = None
                    
        return pendencias

    def _extrair_certidoes(self):
        certidoes = []
        if "Certidão Emitida" in self.texto:
            cert = re.search(r'Certidão (.*?):\s*([A-Z0-9\.]+)\nEmissão:\s*([\d/]+)\nData de Validade:\s*([\d/]+)', self.texto)
            if cert:
                certidoes.append({
                    "tipo": f"Certidão {cert.group(1).strip()}",
                    "codigo": cert.group(2).strip(),
                    "emissao": cert.group(3).strip(),
                    "validade": cert.group(4).strip()
                })
        return certidoes

    def _extrair_pgfn(self):
        if "Não foram detectadas pendências" in self.texto:
            return {"status": "Sem pendências detectadas", "possui_pendencias": False}
        return {"status": "Pendências detectadas", "possui_pendencias": True}