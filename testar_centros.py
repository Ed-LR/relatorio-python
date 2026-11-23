"""Consulta o cadastro de centros de custo para ver quais campos a API devolve (situação, projeto...).
Uso:  python testar_centros.py 42068
Variáveis de ambiente: SIENGE_SUBDOMINIO, SIENGE_USUARIO, SIENGE_SENHA
"""
import os, sys, json, requests

auth = (os.environ["SIENGE_USUARIO"], os.environ["SIENGE_SENHA"])
base = f"https://api.sienge.com.br/{os.environ['SIENGE_SUBDOMINIO']}/public/api/v1/cost-centers"
h = {"accept": "application/json"}
saida = {}
r = requests.get(base, params={"limit": 200, "offset": 0}, auth=auth, headers=h, timeout=120)
print("Lista:", r.status_code)
if r.status_code == 200:
    j = r.json(); res = j.get("results", [])
    saida["metadata"] = j.get("resultSetMetadata"); saida["campos_da_lista"] = sorted({k for x in res for k in x}); saida["primeiros_3"] = res[:3]
    print("Centros nesta página:", len(res))
else:
    saida["lista_erro"] = r.text[:300]
for cc in sys.argv[1:]:
    r = requests.get(f"{base}/{cc}", auth=auth, headers=h, timeout=120)
    print(f"Centro {cc}:", r.status_code)
    saida[f"centro_{cc}"] = r.json() if r.status_code == 200 else r.text[:300]
json.dump(saida, open("centros_amostra.json", "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("Gravado em centros_amostra.json")
