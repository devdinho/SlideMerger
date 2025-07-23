import shutil
import subprocess
import os
from fastapi import FastAPI, File, UploadFile, Response
from tempfile import NamedTemporaryFile

app = FastAPI()

@app.post("/normaliza")
async def normaliza(file: UploadFile = File(...)):
    # Salva o arquivo enviado provisoriamente
    suffix = ".pptx"
    with NamedTemporaryFile(delete=False, suffix=suffix) as tmp_in:
        shutil.copyfileobj(file.file, tmp_in)
        tmp_in_path = tmp_in.name

    # Define o caminho de saída (LibreOffice cria na mesma pasta)
    pasta_saida = os.path.dirname(tmp_in_path)

    # Chamar o LibreOffice para converter/salvar de novo
    process = subprocess.run([
        "soffice", "--headless", "--convert-to", "pptx", "--outdir", pasta_saida, tmp_in_path
    ], capture_output=True)

    # O nome do arquivo convertido é igual ao de entrada (mas pode sobrescrever)
    pptx_out = os.path.join(
        pasta_saida, os.path.basename(tmp_in_path).replace(".pptx", ".pptx")
    )

    # Remove arquivo de entrada
    os.remove(tmp_in_path)
    # Lê o arquivo convertido/final
    with open(pptx_out, "rb") as f:
        pptx_data = f.read()
    os.remove(pptx_out)

    # Retorna o arquivo PPTX normalizado
    return Response(content=pptx_data, media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation")
