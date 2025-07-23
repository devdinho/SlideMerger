import shutil, subprocess, os, glob, time
from tempfile import NamedTemporaryFile
from fastapi import FastAPI, File, UploadFile, Response

app = FastAPI()

@app.post("/normaliza")
async def normaliza(file: UploadFile = File(...)):
    with NamedTemporaryFile(delete=False, suffix=".pptx") as tmp_in:
        shutil.copyfileobj(file.file, tmp_in)
        tmp_in_path = tmp_in.name

    pasta_saida_odp = "/tmp/pptx_to_odp"
    pasta_saida_pptx = "/tmp/odp_to_pptx"
    os.makedirs(pasta_saida_odp, exist_ok=True)
    os.makedirs(pasta_saida_pptx, exist_ok=True)

    # Converter PPTX -> ODP
    process = subprocess.run([
        "soffice", "--headless",
        "--convert-to", "odp",
        "--outdir", pasta_saida_odp,
        tmp_in_path
    ], capture_output=True)

    if process.returncode != 0:
        raise Exception(f"Falha PPTX->ODP:\nSTDOUT: {process.stdout.decode()}\nSTDERR: {process.stderr.decode()}")

    # Buscar arquivo ODP criado
    odp_files = [f for f in os.listdir(pasta_saida_odp) if f.endswith(".odp")]
    if not odp_files:
        raise Exception("Falha: LibreOffice não criou arquivo .odp de saída.")
    odp_path = os.path.join(pasta_saida_odp, odp_files[0])

    # Preparar caminho para saída PPTX
    pptx_saida_path = os.path.join(pasta_saida_pptx, os.path.splitext(odp_files[0])[0] + ".pptx")
    # Se existir, remove arquivo antes
    if os.path.exists(pptx_saida_path):
        os.remove(pptx_saida_path)

    # Converter ODP -> PPTX
    process2 = subprocess.run([
        "soffice", "--headless",
        "--convert-to", "pptx",
        "--outdir", pasta_saida_pptx,
        odp_path
    ], capture_output=True)

    if process2.returncode != 0:
        raise Exception(f"Falha ODP->PPTX:\nSTDOUT: {process2.stdout.decode()}\nSTDERR: {process2.stderr.decode()}")

    # Esperar arquivo PPTX ser criado
    timeout = 10
    for _ in range(timeout):
        if os.path.exists(pptx_saida_path):
            break
        time.sleep(0.5)
    else:
        raise Exception("Falha: LibreOffice não criou arquivo .pptx final de saída.")

    # Ler resultado e limpar temporários
    with open(pptx_saida_path, "rb") as f:
        pptx_data = f.read()

    os.remove(tmp_in_path)
    os.remove(odp_path)
    os.remove(pptx_saida_path)

    return Response(
        content=pptx_data,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )
