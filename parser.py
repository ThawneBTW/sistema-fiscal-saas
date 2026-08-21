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
                self.texto += pagina.extract_text() + "\n"

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
        cnpj_busca = re.search(r'CNPJ:\s*([\d\.\-\/]+)\s+(.+?LTDA|.+?S\.A\.|.+?MEI)', self.texto)
        if cnpj_busca:
            dados['cnpj'] = cnpj_busca.group(1).strip()
            dados['razao_social'] = cnpj_busca.group(2).strip()
            
        sit_busca = re.search(r'Situação:\s*([A-Z]+)', self.texto)
        if sit_busca: dados['situacao'] = sit_busca.group(1).strip()
        
        porte = re.search(r'Porte da Empresa:\s*(.+)', self.texto)
        if porte: dados['porte'] = porte.group(1).strip()
        
        return dados

    def _extrair_socios(self):
        socios = []
        if "Sócios e Administradores" in self.texto:
            # Captura o padrão de CPF/CNPJ, Nome e Qualificação
            matches = re.finditer(r'([\d\.\-]+)\n\s*\|\s*Nome\n(.+?)\n\s*\|\s*Qualificação\n(.+?)\n', self.texto)
            for match in matches:
                socios.append({
                    "documento": match.group(1).strip(),
                    "nome": match.group(2).strip(),
                    "qualificacao": match.group(3).strip()
                })
        return socios

    def _extrair_pendencias(self):
        pendencias = []
        if "Omissão de DCTFWeb" in self.texto:
            periodo = re.search(r'\(Período de Apuração\)\n(\d{4})\n([A-Z]{3})', self.texto)
            if periodo:
                pendencias.append({
                    "orgao": "Receita Federal",
                    "tipo": "Omissão de DCTFWeb",
                    "periodo": f"{periodo.group(2)}/{periodo.group(1)}",
                    "status": "Pendente"
                })
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
            return {"status": "Sem pendências detectadas na PGFN", "possui_pendencias": False}
        return {"status": "Pendências detectadas na PGFN", "possui_pendencias": True}