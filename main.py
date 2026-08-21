from fastapi import FastAPI, UploadFile, File
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse, HTMLResponse
import os
import shutil
from parser import RelatorioParser
from excel import gerar_planilha

app = FastAPI()

# Configuração para servir o HTML (Frontend)
app.mount("/static", StaticFiles(directory="static"), name="static")

@app.get("/", response_class=HTMLResponse)
async def ler_index():
    with open("static/index.html", "r", encoding="utf-8") as f:
        return f.read()

@app.post("/upload")
async def processar_pdf(file: UploadFile = File(...)):
    caminho_temp = f"temp_{file.filename}"
    
    with open(caminho_temp, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)
        
    try:
        # Processa o PDF
        leitor = RelatorioParser(caminho_temp)
        dados = leitor.extrair_dados()
        
        # Gera o Excel
        arquivo_excel = gerar_planilha(dados['empresa'], dados['pendencias'])
        
        # Aqui, no futuro, inserimos no Supabase.
        # Por enquanto, retornamos os dados processados para a tela.
        return {
            "mensagem": "Processado com sucesso",
            "dados": dados,
            "excel_url": f"/download/{arquivo_excel}"
        }
    finally:
        if os.path.exists(caminho_temp):
            os.remove(caminho_temp)

@app.get("/download/{nome_arquivo}")
async def baixar_excel(nome_arquivo: str):
    return FileResponse(path=nome_arquivo, filename=nome_arquivo, media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')