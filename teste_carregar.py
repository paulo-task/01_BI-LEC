# Leitura manual do .pass
credenciais = {}
with open(".pass", "r") as f:
    for linha in f:
        linha = linha.strip()
        if "=" in linha:
            chave, valor = linha.split("=", 1)
            credenciais[chave] = valor

print("Credenciais lidas:")
print(f"CPFL_USER: {credenciais.get('CPFL_USER')}")
print(f"CPFL_PASS: {credenciais.get('CPFL_PASS')}")