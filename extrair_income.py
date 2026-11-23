"""Extrai do income_bruto.json (contas a receber) só as parcelas de alguns títulos, para enviar um arquivo pequeno.
Uso:  python extrair_income.py 27682 27819 27558
Não chama a API.
"""
import sys, json
dados = json.load(open("income_bruto.json", encoding="utf-8"))["data"]
ids = {int(a) for a in sys.argv[1:]}
sel = [p for p in dados if p.get("billId") in ids]
json.dump(sel, open("income_amostra.json", "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print(f"{len(sel)} parcelas gravadas em income_amostra.json")
