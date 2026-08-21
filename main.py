from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse, HTMLResponse
import os
import shutil
from supabase import create_client, Client
from parser import RelatorioParser
from excel import gerar_planilha

app = FastAPI()

# 1. Configuração para servir o visual (Frontend)
app.mount("/static", StaticFiles(directory="static"), name="static")

# 2. Conexão com o Banco de Dados (Supabase)
URL_SUPABASE = os.environ.get("SUPABASE_URL", "")
CHAVE_SUPABASE = os.environ.get("SUPABASE_KEY", "")

supabase = None
if URL_SUPABASE and CHAVE_SUPABASE:
    supabase = create_client(URL_SUPABASE, CHAVE_SUPABASE)

# 3. Rota principal: Carrega a interface bonita
@app.get("/", response_class=HTMLResponse)
async def ler_index():
    try:
        with open("static/index.html", "r", encoding="utf-8") as f:
            return f.read()
    except FileNotFoundError:
        return "Erro: O arquivo static/index.html não foi encontrado. Verifique as pastas."

# 4. Rota de Upload: A mágica acontece aqui
@app.post("/upload")
async def processar_pdf(file: UploadFile = File(...)):
    if not file.filename.endswith(".pdf"):
        raise HTTPException(status_code=400, detail="Apenas arquivos PDF são aceitos.")

    caminho_temp = f"temp_{file.filename}"
    
    with open(caminho_temp, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)
        
    try:
        # A. Faz a leitura do PDF
        leitor = RelatorioParser(caminho_temp)
        dados = leitor.extrair_dados()
        
        empresa = dados.get('empresa', {})
        pendencias = dados.get('pendencias', [])
        socios = dados.get('socios', [])
        certidoes = dados.get('certidoes', [])
        pgfn = dados.get('pgfn', {})

        # B. Salva no Banco de Dados
        if supabase and 'cnpj' in empresa:
            # Salva ou atualiza a empresa
            supabase.table("empresas").upsert({
                "cnpj": empresa['cnpj'],
                "razao_social": empresa.get('razao_social', ''),
                "situacao": empresa.get('situacao', '')
            }).execute()

            # Salva as pendências INDEPENDENTES
            for p in pendencias:
                supabase.table("pendencias").insert({
                    "cnpj_empresa": empresa['cnpj'],
                    "orgao": p.get('orgao', ''),
                    "tipo": p.get('tipo', ''),
                    "periodo": p.get('periodo', ''),
                    "status": p.get('status', 'Pendente')
                }).execute()

        # C. Gera a planilha Excel Profissional com todas as abas
        arquivo_excel = gerar_planilha(empresa, pendencias, socios, certidoes, pgfn)
        
        # D. Devolve tudo mastigado pro Javascript mostrar na tela
        return {
            "mensagem": "Processado com sucesso",
            "dados": dados,
            "excel_url": f"/download/{arquivo_excel}"
        }
        
    except Exception as e:
        print(f"Erro no processamento: {e}")
        raise HTTPException(status_code=500, detail="Erro interno ao processar o relatório.")
        
    finally:
        # Sempre apaga o PDF temporário para não lotar o servidor
        if os.path.exists(caminho_temp):
            os.remove(caminho_temp)

# 5. Rota de Download: Entrega o arquivo Excel gerado
@app.get("/download/{nome_arquivo}")
async def baixar_excel(nome_arquivo: str):
    if os.path.exists(nome_arquivo):
        return FileResponse(
            path=nome_arquivo, 
            filename=nome_arquivo, 
            media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    raise HTTPException(status_code=404, detail="Arquivo Excel não encontrado.")