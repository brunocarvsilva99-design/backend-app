import xlwings as xw
import os
from tkinter import Tk, filedialog
import re

# === Selecionar pasta com arquivos Excel ===
Tk().withdraw()
pasta = filedialog.askdirectory(title="Selecione a pasta onde estão os arquivos Excel")

if not pasta:
    print("⚠️ Nenhuma pasta selecionada. Encerrando...")
    exit()

# Arquivo de destino fixo
destino_path = os.path.join(pasta, "CVC_Banco_de_Dados.xlsx")
if not os.path.exists(destino_path):
    print(f"❌ Arquivo de destino não encontrado: {destino_path}")
    exit()

# Abrir Excel
app = xw.App(visible=True)  # mudar para False para rodar em background
wb_destino = xw.Book(destino_path)

# Selecionar ou criar aba de destino
if "Plan1" not in [s.name for s in wb_destino.sheets]:
    aba_destino = wb_destino.sheets.add("Plan1")
else:
    aba_destino = wb_destino.sheets["Plan1"]

# Procurar arquivos de origem
arquivos_origem = [f for f in os.listdir(pasta) if re.match(r"VMDA_Seção_P\d+\.xlsx", f)]

linha_destino = 2  # linha inicial do banco de dados

# Função para ler intervalo de células e garantir lista simples de valores
def ler_intervalo(aba, intervalo):
    valores = aba.range(intervalo).value
    if not valores:
        return [0]
    if isinstance(valores, list):
        # Se for matriz, achata para lista simples
        valores = [item for sublist in valores for item in (sublist if isinstance(sublist, list) else [sublist])]
    else:
        valores = [valores]
    return [v or 0 for v in valores]

for arquivo in sorted(arquivos_origem):
    origem_path = os.path.join(pasta, arquivo)
    try:
        wb_origem = xw.Book(origem_path)
    except Exception as e:
        print(f"❌ Não foi possível abrir {arquivo}: {e}")
        continue

    if "VMDA" not in [s.name for s in wb_origem.sheets]:
        print(f"❌ Aba 'VMDA' não encontrada em {arquivo}. Pulando...")
        wb_origem.close()
        continue

    aba_origem = wb_origem.sheets["VMDA"]
    try:
        # V = D32
        aba_destino.range(f"V{linha_destino}").value = aba_origem.range("D32").value or 0
        # W = E32 + F32
        aba_destino.range(f"W{linha_destino}").value = sum(ler_intervalo(aba_origem, "E32:F32"))
        # X = soma G32:K32
        aba_destino.range(f"X{linha_destino}").value = sum(ler_intervalo(aba_origem, "G32:K32"))
        # Y = L32
        aba_destino.range(f"Y{linha_destino}").value = aba_origem.range("L32").value or 0
        # Z = M32 + N32
        aba_destino.range(f"Z{linha_destino}").value = sum(ler_intervalo(aba_origem, "M32:N32"))
        # AA = soma O32:T32
        aba_destino.range(f"AA{linha_destino}").value = sum(ler_intervalo(aba_origem, "O32:T32"))
        # AB = soma U32:AA32
        aba_destino.range(f"AB{linha_destino}").value = sum(ler_intervalo(aba_origem, "U32:AA32"))
        # AC = soma AB32:AE32
        aba_destino.range(f"AC{linha_destino}").value = sum(ler_intervalo(aba_origem, "AB32:AE32"))
        # AD = AF32 + AG32
        aba_destino.range(f"AD{linha_destino}").value = sum(ler_intervalo(aba_origem, "AF32:AG32"))
        # AF = soma AH32:AJ32
        aba_destino.range(f"AF{linha_destino}").value = sum(ler_intervalo(aba_origem, "AH32:AJ32"))
        # AG = soma D32:AJ32
        aba_destino.range(f"AG{linha_destino}").value = sum(ler_intervalo(aba_origem, "D32:AJ32"))
        print(f"✅ Dados de {arquivo} copiados para linha {linha_destino}")
    except Exception as e:
        print(f"❌ Erro ao processar {arquivo}: {e}")
    finally:
        wb_origem.close()
    linha_destino += 1

# Salvar e fechar o arquivo de destino
try:
    wb_destino.save(destino_path)
    print("✅ Arquivo de destino salvo com sucesso!")
except Exception as e:
    print(f"⚠️ Erro ao salvar arquivo de destino: {e}")
finally:
    wb_destino.close()
    app.quit()
