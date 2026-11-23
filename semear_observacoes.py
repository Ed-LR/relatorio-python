"""Semeia o cache de observações com a coluna 'Observação' dos seus relatórios manuais do Sienge (Realizado.csv e Arealizar.csv).
Conferi: nos títulos em comum, a observação do relatório é idêntica ao campo 'notes' da API. Assim o histórico entra sem usar a cota de 100 chamadas/dia.
- títulos a pagar  -> observacoes_cache.json     (não sobrescreve o que a API já trouxe)
- títulos a receber (origem CR) -> observacoes_cr_cache.json   (a API não tem permissão; vem só dos relatórios)
Uso:  python semear_observacoes.py [pasta_dos_relatorios]     (padrão: DASHBOARD_DESTINO, depois a pasta atual)
Não chama a API.
"""
import os, sys, json, pandas as pd

pasta = sys.argv[1] if len(sys.argv) > 1 else os.environ.get("DASHBOARD_DESTINO") or "."
quadros = []
for nome in ("Realizado.csv", "Arealizar.csv"):
    caminho = os.path.join(pasta, nome)
    if not os.path.exists(caminho): print(f"Não encontrei {caminho}"); continue
    quadros.append(pd.read_csv(caminho, sep=";", encoding="cp1252", dtype=str))
if not quadros: sys.exit("Nenhum relatório encontrado.")
D = pd.concat(quadros, ignore_index=True)
D["bill"] = D["Lançamento"].str.extract(r"^(\d+)/")[0]
D = D[D.bill.notna() & D["Observação"].notna()].copy()
D["Observação"] = D["Observação"].str.strip()
D = D[D["Observação"] != ""]

def semear(df, arquivo):
    cache = json.load(open(arquivo, encoding="utf-8")) if os.path.exists(arquivo) else {}
    novos = {b: t for b, t in df.groupby("bill")["Observação"].first().items() if b not in cache}
    cache.update(novos)
    tmp = arquivo + ".tmp"; json.dump(cache, open(tmp, "w", encoding="utf-8"), ensure_ascii=False); os.replace(tmp, arquivo)
    print(f"{arquivo}: +{len(novos)} títulos (total {len(cache)})")

semear(D[D.Origem != "CR"], "observacoes_cache.json")
semear(D[D.Origem == "CR"], "observacoes_cr_cache.json")
