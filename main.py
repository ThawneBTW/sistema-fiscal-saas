from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse, HTMLResponse
import os
import shutil
import traceback
from supabase import create_client, Client
from parser import RelatorioParser
from excel import gerar_planilha

app = FastAPI()

app.mount("/static", StaticFiles(directory="static"), name="static")

URL_SUPABASE = os.environ.get("SUPABASE_URL", "")
CHAVE_SUPABASE = os.environ.get("SUPABASE_KEY", "")

supabase = None
if URL_SUPABASE and CHAVE_SUPABASE:
    supabase = create_client(URL_SUPABASE, CHAVE_SUPABASE)

@app.get("/", response_class=HTMLResponse)
async def ler_index():
    try:
        with open("static/index.html", "r", encoding="utf-8") as f:
            return f.read()
    except FileNotFoundError:
        return "Erro: Arquivo index.html não encontrado."

@app.post("/upload")
async def processar_pdf(file: UploadFile = File(...)):
    if not file.filename.endswith(".pdf"):
        raise HTTPException(status_code=400, detail="Apenas arquivos PDF são aceitos.")

    caminho_temp = f"temp_{file.filename}"
    
    with open(caminho_temp, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)
        
    try:
        # Extrai todas as 5 partes do documento
        leitor = RelatorioParser(caminho_temp)
        dados = leitor.extrair_dados()
        
        empresa = dados.get('empresa', {})
        pendencias = dados.get('pendencias', [])
        socios = dados.get('socios', [])
        certidoes = dados.get('certidoes', [])
        pgfn = dados.get('pgfn', {})

        # Salva as informações nas tabelas que você acabou de criar!
        if supabase and 'cnpj' in empresa:
            try:
                supabase.table("empresas").upsert({
                    "cnpj": empresa['cnpj'],
                    "razao_social": empresa.get('razao_social', ''),
                    "situacao": empresa.get('situacao', '')
                }).execute()

                for p in pendencias:
                    supabase.table("pendencias").insert({
                        "cnpj_empresa": empresa['cnpj'],
                        "orgao": p.get('orgao', ''),
                        "tipo": p.get('tipo', ''),
                        "periodo": p.get('periodo', ''),
                        "status": p.get('status', 'Pendente')
                    }).execute()
            except Exception as erro_banco:
                print(f"Aviso do Banco: {erro_banco}")

        # Gera o Excel passando TODAS as 5 informações (Isso resolve o bug!)
        arquivo_excel = gerar_planilha(empresa, pendencias, socios, certidoes, pgfn)
        
        return {
            "mensagem": "Processado com sucesso",
            "dados": dados,
            "excel_url": f"/download/{arquivo_excel}"
        }
        
    except Exception as e:
        # Se der erro agora, o Toast vermelho vai mostrar a mensagem EXATA!
        erro_detalhado = traceback.format_exc()
        print(erro_detalhado)
        raise HTTPException(status_code=500, detail=str(e))
        
    finally:
        if os.path.exists(caminho_temp):
            os.remove(caminho_temp)

@app.get("/download/{nome_arquivo}")
async def baixar_excel(nome_arquivo: str):
    if os.path.exists(nome_arquivo):
        return FileResponse(
            path=nome_arquivo, 
            filename=nome_arquivo, 
            media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    raise HTTPException(status_code=404, detail="Arquivo Excel não encontrado.")