from openpyxl import load_workbook

def criar_aba(nome_bairro, arquivo):
    if nome_bairro and nome_bairro not in arquivo.sheetnames:
        nova_aba = arquivo.create_sheet(nome_bairro) # Cria e já guarda a referência
        nova_aba["A1"].value = "data de nascimento"
        nova_aba["B1"].value = "pessoa"
        nova_aba["C1"].value = "Bairro"
    return arquivo[nome_bairro] # Retorna o objeto da aba (existente ou nova)

def transferir_informacoes_aba(aba_origem, aba_destino, linha_origem):
    # Usamos aba_destino (o objeto) para achar a última linha dela
    linha_destino = aba_destino.max_row + 1 
    for coluna in range(1, 4):
        valor = aba_origem.cell(row=linha_origem, column=coluna).value
        aba_destino.cell(row=linha_destino, column=coluna).value = valor

# --- Execução principal ---
arquivo = load_workbook("planilha_ribeirao_preto.xlsx")
aba_base = arquivo['Dados']
ultima_linha = aba_base.max_row

for i in range(2, ultima_linha + 1):
    nome_do_bairro = aba_base.cell(row=i, column=3).value
    
    if not nome_do_bairro:
        continue 
    
    # 1. Garantimos que a aba existe e pegamos o objeto dela
    aba_especifica = criar_aba(nome_do_bairro, arquivo)

    # 2. Passamos o objeto da aba destino (aba_especifica) em vez de apenas o nome
    transferir_informacoes_aba(aba_base, aba_especifica, i)

arquivo.save("Bairrosrp.xlsx")
print("Processo concluído: Abas criadas e dados transferidos!")
