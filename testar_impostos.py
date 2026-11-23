"""Consulta os impostos de alguns títulos a pagar:  GET /bills/{billId}/taxes
Uso:  python testar_impostos.py 744638 744508 744416
Variáveis de ambiente: SIENGE_SUBDOMINIO, SIENGE_USUARIO, SIENGE_SENHA   (a senha nunca vai para o arquivo)
"""
import os, sys, json, requests

auth = (os.environ["SIENGE_USUARIO"], os.environ["SIENGE_SENHA"])
sub = os.environ["SIENGE_SUBDOMINIO"]
saida = []
for bill in sys.argv[1:]:
    url = f"https://api.sienge.com.br/{sub}/public/api/v1/bills/{bill}/taxes"
    r = requests.get(url, params={"limit": 100, "offset": 0}, auth=auth, headers={"accept": "application/json"}, timeout=120)
    print(f"Título {bill}: status {r.status_code}")
    try: corpo = r.json()
    except Exception: corpo = r.text[:500]
    if r.status_code != 200: print("   resposta:", str(corpo)[:300])
    saida.append({"billId": bill, "status": r.status_code, "resposta": corpo})
json.dump(saida, open("impostos_amostra.json", "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("Gravado em impostos_amostra.json")
