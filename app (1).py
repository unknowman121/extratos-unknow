
import pdfplumber
import re
import pandas as pd
from tqdm import tqdm
from IPython.display import display
import unicodedata
import datetime
import os

def remover_acentos(texto):
    return ''.join(c for c in unicodedata.normalize('NFD', texto)
                   if unicodedata.category(c) != 'Mn')

def validar_data(data):
    for fmt in ("%d/%m/%Y", "%d/%m"):
        try:
            datetime.datetime.strptime(data, fmt)
            return True
        except ValueError:
            continue
    return False

palavras_chave_fixas = [
    "TAR BLOQUETO ITAU", "TAR SISPAG TIT OUTRO BCO", "TAR TED SISPAG",
    "TAR C/C SISPAG", "TAR PIX", "COB COMPE", "TAR DEV CHQ 000372",
    "TAR DEV CH 000363", "EST TAR CH VAL SUP000363", "TAR DEV CHQ 000366",
    "TAR EXT EL", "TARIFA MANUTENCAO CONTA A", "MENSALIDADE CESTA SERVICO",
    "COB INTERN", "COB BX 063", "DEB PIX CH", "TARIFA TRANSF RECURSO AG"
]

palavras_parteadas = [
    ["TAR", "CH", "VALOR", "SUP", "000266"],
    ["TAR", "CH", "VALOR", "SUP", "000327"]
]

data_filtro = input("📅 Digite a data para filtrar (ex: 12/07 ou 12/07/2025): ").strip()
while not validar_data(data_filtro):
    print("❌ Formato de data inválido.")
    data_filtro = input("📅 Digite novamente a data (ex: 12/07 ou 12/07/2025): ").strip()

print("📂 Envie os PDFs com extratos bancários:")
uploaded = os.listdir()
arquivos = [f for f in uploaded if f.lower().endswith(".pdf")]

resultados_finais = []

for arquivo in arquivos:
    print(f"📁 Processando: {arquivo}")
    total = 0.0
    registros = []

    try:
        with pdfplumber.open(arquivo) as pdf:
            total_paginas = len(pdf.pages)
            for i in tqdm(range(total_paginas), desc="🔍 Analisando páginas"):
                texto_pagina = pdf.pages[i].extract_text()
                if not texto_pagina:
                    continue
                linhas = texto_pagina.split('\n')
                for linha in linhas:
                    linha_limpa = remover_acentos(' '.join(linha.upper().split()))
                    if data_filtro not in linha_limpa:
                        continue

                    encontrou_chave = any(chave in linha_limpa for chave in palavras_chave_fixas)
                    if not encontrou_chave:
                        for grupo in palavras_parteadas:
                            if all(p in linha_limpa for p in grupo):
                                encontrou_chave = True
                                break

                    if encontrou_chave:
                        match = re.search(r'(\d{1,3}(?:\.\d{3})*,\d{2})D\b', linha_limpa)
                        if match:
                            valor = float(match.group(1).replace('.', '').replace(',', '.'))
                            total += valor
                            registros.append({
                                "Página": i + 1,
                                "Linha": linha,
                                "Valor (R$)": valor
                            })
    except Exception as e:
        print(f"⚠️ Erro ao processar {arquivo}: {e}")
        continue

    df_registros = pd.DataFrame(registros).sort_values(by="Valor (R$)", ascending=False)

    if df_registros.empty:
        print("✅ Nenhuma tarifa correspondente foi encontrada nesse arquivo.")
    else:
        print(df_registros)
        df_registros.to_excel(f"{arquivo}_detalhado.xlsx", index=False)

    resultados_finais.append({
        "Arquivo": arquivo,
        "Total Tarifas Encontradas (R$)": round(total, 2)
    })

print("\n📋 Resumo Final:")
df_resultado = pd.DataFrame(resultados_finais)
print(df_resultado)
df_resultado.to_excel("resumo_final.xlsx", index=False)
