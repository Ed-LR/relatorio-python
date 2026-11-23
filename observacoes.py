"""Observação dos títulos a pagar: GET /bills/{billId} -> campo 'notes' (confirmado com testar_observacao.py).
Uma consulta por título, guardada em observacoes_cache.json (igual ao cache de impostos). Só títulos ainda não guardados são consultados,
do mais novo para o mais antigo, até o limite de chamadas da execução; o que sobrar entra nas execuções seguintes.
Atenção: contas a RECEBER (origem CR) NÃO usam este endereço (os números de título se parecem, mas são outra sequência);
o endereço de recebíveis respondeu 403 (sem permissão) e fica de fora até ser liberado no Sienge.
"""
import os, json, requests

CACHE = "observacoes_cache.json"

def carregar():
    return json.load(open(CACHE, encoding="utf-8")) if os.path.exists(CACHE) else {}

def salvar(cache):
    tmp = CACHE + ".tmp"
    json.dump(cache, open(tmp, "w", encoding="utf-8"), ensure_ascii=False)
    os.replace(tmp, CACHE)

def consultar(bill_ids, max_chamadas=1500, offline=False):
    """Completa o cache com os títulos que faltam. Devolve (cache, resumo)."""
    cache = carregar()
    faltam = sorted((b for b in {str(int(b)) for b in bill_ids} if b not in cache), key=int, reverse=True)
    resumo = {"titulos": len(set(map(str, bill_ids))), "em_cache": 0, "consultadas": 0, "pendentes": len(faltam), "parou": ""}
    if offline or not faltam:
        resumo["em_cache"] = resumo["titulos"] - len(faltam); return cache, resumo
    url = f"https://api.sienge.com.br/{os.environ['SIENGE_SUBDOMINIO']}/public/api/v1/bills/" + "{}"
    s = requests.Session(); s.auth = (os.environ["SIENGE_USUARIO"], os.environ["SIENGE_SENHA"]); s.headers["accept"] = "application/json"
    seguidos = 0
    for n, b in enumerate(faltam[:max_chamadas], 1):
        try:
            r = s.get(url.format(b), timeout=120)
        except requests.RequestException as e:
            seguidos += 1
            if seguidos >= 5: resumo["parou"] = f"falhas de conexão seguidas ({e})"; break
            continue
        if r.status_code == 429: resumo["parou"] = "limite diário (429)"; break
        if r.status_code in (401, 403): resumo["parou"] = f"sem permissão ({r.status_code}) em /bills"; break
        if r.status_code == 200: cache[b] = (r.json().get("notes") or "").strip(); seguidos = 0; resumo["consultadas"] += 1
        elif r.status_code == 404: cache[b] = ""; seguidos = 0; resumo["consultadas"] += 1
        else:
            seguidos += 1
            if seguidos >= 5: resumo["parou"] = f"status {r.status_code} repetido"; break
        if n % 100 == 0: salvar(cache)
    salvar(cache)
    resumo["pendentes"] = len(faltam) - resumo["consultadas"]
    resumo["em_cache"] = resumo["titulos"] - resumo["pendentes"]
    return cache, resumo
