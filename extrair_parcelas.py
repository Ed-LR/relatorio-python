"""Extrai do arquivo bruto do /outcome só as parcelas de alguns títulos (para me enviar um arquivo pequeno).
Uso:  python extrair_parcelas.py 744638 744508 744416
"""
import sys, json
dados = json.load(open("bruto_outcome_2025-01-01_2030-12-31.json", encoding="utf-8"))["data"]
ids = {int(a) for a in sys.argv[1:]}
sel = [p for p in dados if p.get("billId") in ids]
json.dump(sel, open("parcelas_amostra.json", "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print(f"{len(sel)} parcelas gravadas em parcelas_amostra.json")
