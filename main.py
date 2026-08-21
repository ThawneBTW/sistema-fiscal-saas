from fastapi import FastAPI, UploadFile, File
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
supabase = create_client(URL_SUPABASE, CHAVE_SUPABASE) if URL_SUPABASE and CHAVE_SUPABASE else None

@app.get("/", response_class=HTMLResponse)
async def ler_index():
    if os.path.exists("static/index.html"):
        with open("static/index.html", "r", encoding="utf-8") as f:
            return f.read()
    return "Erro: Interface não encontrada."

@app.post("/upload")
async def processar_pdf(file: UploadFile = File(...)):
    if not file.filename.endswith(".pdf"):
        return {"error": "Formato inválido. Envie apenas PDF."}

    caminho = f"temp_{file.filename}"
    with open(caminho, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)
        
    try:
        leitor = RelatorioParser(caminho)
        dados = leitor.extrair_dados()
        empresa = dados.get('empresa', {})
        pendencias = dados.get('pendencias', [])

        if supabase and empresa.get('cnpj'):
            try:
                supabase.table("empresas").upsert({"cnpj": empresa['cnpj'], "razao_social": empresa.get('razao_social', ''), "situacao": empresa.get('situacao', '')}).execute()
                for p in pendencias:
                    supabase.table("pendencias").insert({
                        "cnpj_empresa": empresa['cnpj'],
                        "orgao": p.get('orgao', ''),
                        "tipo": p.get('tipo', ''),
                        "periodo": p.get('periodo', ''),
                        "status": p.get('status', 'Pendente')
                    }).execute()
            except Exception as bd_err:
                print(f"Aviso BD: {bd_err}")

        arquivo_excel = gerar_planilha(dados)
        
        return {
            "sucesso": True,
            "dados": dados,
            "excel_url": f"/download/{arquivo_excel}"
        }
        
    except Exception as e:
        traceback.print_exc()
        return {"error": f"Falha na extração: {str(e)}"}
        
    finally:
        if os.path.exists(caminho):
            os.remove(caminho)

@app.get("/download/{arquivo}")
async def baixar(arquivo: str):
    if os.path.exists(arquivo):
        return FileResponse(path=arquivo, filename=arquivo, media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    return {"error": "Arquivo não encontrado."}