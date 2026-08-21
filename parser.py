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
            "pendencias": self._extrair_pendencias()
        }

    def _extrair_empresa(self):
        dados = {}
        cnpj_busca = re.search(r'CNPJ:\s*([\d\.\-\/]+)\s+(.+?LTDA|.+?S\.A\.|.+?MEI)', self.texto)
        if cnpj_busca:
            dados['cnpj'] = cnpj_busca.group(1).strip()
            dados['razao_social'] = cnpj_busca.group(2).strip()
            
        sit_busca = re.search(r'Situação:\s*([A-Z]+)', self.texto)
        if sit_busca:
            dados['situacao'] = sit_busca.group(1).strip()
            
        return dados

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