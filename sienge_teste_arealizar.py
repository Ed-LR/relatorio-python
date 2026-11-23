"""Teste do A REALIZAR (contas a pagar) via API do Sienge: /outcome  + conferência com o Arealizar.csv.
Uso:  python sienge_teste_arealizar.py 2025-01-01 2030-12-31 [reusar]
Variáveis de ambiente: SIENGE_SUBDOMINIO, SIENGE_USUARIO, SIENGE_SENHA, SIENGE_INDEXADOR (0 = REAL)
As retenções (INSS, ISS...) vêm de GET /bills/{billId}/taxes, uma consulta por título, guardada em impostos_cache.json.
Precisa estar na mesma pasta de sienge_teste_realizado.py, Arealizar.csv, Baseapropriação.xlsx e relatorio_centro_de_custo.xlsx
"""
import os, sys, json, datetime, requests, pandas as pd
from sienge_teste_realizado import centros_ativos, classificar

BASE = "https://api.sienge.com.br/{sub}/public/api/bulk-data/v1/outcome"
TAXES = "https://api.sienge.com.br/{sub}/public/api/v1/bills/{bill}/taxes"
CONTA_IMPOSTO = {"INSS": "2010124", "ISS": "2010125", "PCC": "2010122", "IR": "2010123"}   # confirmados pelos valores e pelos nomes das contas; outro imposto vira pendência (nada é adivinhado)
CACHE = "impostos_cache.json"

def buscar(ini, fim, centros):
    auth = (os.environ["SIENGE_USUARIO"], os.environ["SIENGE_SENHA"])
    url = BASE.format(sub=os.environ["SIENGE_SUBDOMINIO"])
    base = {"startDate": ini, "endDate": fim, "selectionType": "D",            # D = data de vencimento da parcela
            "correctionIndexerId": int(os.environ["SIENGE_INDEXADOR"]), "correctionDate": datetime.date.today().isoformat(),
            "withBankMovements": "false", "withAuthorizations": "false"}
    for lista in (centros, ",".join(map(str, centros))):
        r = requests.get(url, params=dict(base, costCentersId=lista), auth=auth, headers={"accept": "application/json"}, timeout=900)
        print("Status:", r.status_code, "| bytes:", len(r.content))
        if r.status_code == 404: return []
        if r.status_code == 429: sys.exit("Limite diário atingido (429). Tente amanhã.")
        if r.status_code == 400:
            print("Resposta da API:", r.text[:400]); continue
        r.raise_for_status()
        with open(f"bruto_outcome_{ini}_{fim}.json", "w", encoding="utf-8") as f: f.write(r.text)
        return r.json().get("data", [])
    sys.exit("A API recusou a consulta (400). Confira o SIENGE_INDEXADOR e a mensagem acima.")

def buscar_impostos(bill_ids):
    cache = json.load(open(CACHE, encoding="utf-8")) if os.path.exists(CACHE) else {}
    faltam = [b for b in dict.fromkeys(bill_ids) if str(b) not in cache]
    print(f"{len(set(bill_ids))} títulos com retenção | já guardados: {len(set(bill_ids)) - len(faltam)} | a consultar: {len(faltam)}")
    if faltam:
        auth = (os.environ["SIENGE_USUARIO"], os.environ["SIENGE_SENHA"]); sub = os.environ["SIENGE_SUBDOMINIO"]
        for i, b in enumerate(faltam, 1):
            r = requests.get(TAXES.format(sub=sub, bill=b), params={"limit": 100, "offset": 0}, auth=auth, headers={"accept": "application/json"}, timeout=120)
            if r.status_code == 429: print("Limite diário (429) atingido; o que já foi consultado fica guardado."); break
            if r.status_code != 200: sys.exit(f"Título {b}: status {r.status_code} {r.text[:200]}")
            cache[str(b)] = r.json().get("results", [])
            if i % 20 == 0: json.dump(cache, open(CACHE, "w", encoding="utf-8"), ensure_ascii=False)
        json.dump(cache, open(CACHE, "w", encoding="utf-8"), ensure_ascii=False)
    return cache

def linhas_retencao(p, doc, lanc, cats, impostos):
    """Abre o taxAmount da parcela por imposto (valores de /taxes) e rateia pelos mesmos centros da parcela."""
    taxas = impostos.get(str(p["billId"])) or []
    total_api = sum(t["amount"] for t in taxas)
    if not taxas or total_api == 0: return []
    por_centro = {}
    for fc in cats: por_centro[fc["costCenterId"]] = por_centro.get(fc["costCenterId"], 0) + fc["financialCategoryRate"]
    out = []
    for t in taxas:
        parte = p["taxAmount"] * t["amount"] / total_api            # parcela única: igual ao valor do imposto
        for cc, taxa in por_centro.items():
            v = round(parte * taxa / 100, 2)
            if v: out.append(dict(data=p["dueDate"], doc=doc, lanc=lanc, centro=str(cc), conta=CONTA_IMPOSTO.get(t["taxId"], "IMPOSTO_" + str(t["taxId"])),
                                  valor=v, valor_original=v, origem=p.get("originId") or "", previsao="", credor="retenção " + str(t["taxId"])))
    return out

def linhas(parcelas, ativos, impostos=None):
    ativos = set(ativos); out = []
    for p in parcelas:
        doc = f"{(p.get('documentIdentificationId') or '').strip()}.{(p.get('documentNumber') or '').strip()}"
        lanc = f"{p['billId']}/{p['installmentId']}"
        cats_ativas = [fc for fc in (p.get("paymentsCategories") or []) if fc["costCenterId"] in ativos]
        if impostos is not None and p.get("taxAmount") and p.get("balanceAmount"):
            out += linhas_retencao(p, doc, lanc, cats_ativas, impostos)
        for fc in cats_ativas:
            sinal = 1 if fc.get("financialCategoryReducer") == "S" else -1      # saída negativa; conta redutora inverte
            out.append(dict(data=p["dueDate"], doc=doc, lanc=lanc, centro=str(fc["costCenterId"]), conta=str(fc["financialCategoryId"]).strip(),
                            valor=sinal * round(p["balanceAmount"] * fc["financialCategoryRate"] / 100, 2),
                            valor_original=sinal * round(p["originalAmount"] * fc["financialCategoryRate"] / 100, 2),
                            origem=p.get("originId") or "", previsao=p.get("forecastDocument") or "", credor=(p.get("creditorName") or "").strip()))
    d = pd.DataFrame(out)
    return d[(d.valor != 0) | (d.valor_original != 0)] if len(d) else d

def csv_referencia(ini, fim, ativos):
    D = pd.read_csv("Arealizar.csv", sep=";", encoding="cp1252", dtype=str)
    D = D[D["Código centro de custo"].isin(map(str, ativos)) & (D.Origem != "CR")].copy()      # CR = a receber (virá do /income)
    D["data"] = pd.to_datetime(D.Data, format="%d/%m/%Y").dt.strftime("%Y-%m-%d"); D = D[(D.data >= ini) & (D.data <= fim)]
    num = lambda s: s.str.replace(".", "", regex=False).str.replace(",", ".", regex=False).astype(float)
    ap = pd.read_excel("Baseapropriação.xlsx", dtype=str); ap.columns = ["cod", "cn", "tc", "cat"]
    ap["k"] = ap.cod.str.replace(r"\D", "", regex=True); nat = ap.drop_duplicates("k").set_index("k").cat
    D["conta"] = D.Conta.str.extract(r"^(\d+)")[0]
    n = D.conta.map(nat); n[D.conta == "2120101"] = "Despesa"; n[D.conta == "1020103"] = "Receita"
    v = num(D.Valor); D["valor"] = v.where(n != "Despesa", -v)
    D["doc"] = D.Documento.str.strip(); D["lanc"] = D["Lançamento"].fillna(""); D["centro"] = D["Código centro de custo"]
    return D[D.valor != 0][["data", "doc", "lanc", "centro", "conta", "valor"]]

def comparar(api, ref, col):
    k = ["data", "doc", "lanc", "centro", "conta"]
    a = api.groupby(k)[col].sum().round(2); b = ref.groupby(k).valor.sum().round(2)
    j = pd.concat([a.rename("api"), b.rename("csv")], axis=1).fillna(0); j["dif"] = (j.api - j.csv).round(2)
    ruins = j[j.dif.abs() > 0.02]
    print(f"\n=== Comparação usando '{col}' ===\nAPI: {int((api[col] != 0).sum())} linhas, total {api[col].sum():,.2f} | CSV: {len(ref)} linhas, total {ref.valor.sum():,.2f}")
    print(f"Chaves com diferença acima de 2 centavos: {len(ruins)} de {len(j)}")
    return j, ruins

if __name__ == "__main__":
    ini, fim = (sys.argv[1], sys.argv[2]) if len(sys.argv) > 2 else ("2025-01-01", "2030-12-31")
    ativos = centros_ativos(); print(f"{len(ativos)} centros em andamento")
    if len(sys.argv) > 3 and sys.argv[3] == "reusar":
        parc = json.load(open(f"bruto_outcome_{ini}_{fim}.json", encoding="utf-8"))["data"]; print("Reaproveitando o arquivo bruto salvo (nenhuma chamada à API).")
    else:
        parc = buscar(ini, fim, ativos)
    com_imposto = [p["billId"] for p in parc if p.get("taxAmount") and p.get("balanceAmount") and any(c["costCenterId"] in set(ativos) for c in (p.get("paymentsCategories") or []))]
    impostos = buscar_impostos(com_imposto)
    api = classificar(linhas(parc, ativos, impostos).assign(hist=""))
    api.to_csv(f"api_outcome_{ini}_{fim}.csv", sep=";", decimal=",", index=False, encoding="utf-8-sig")
    ref = csv_referencia(ini, fim, ativos)
    res = {c: comparar(api, ref, c) for c in ("valor", "valor_original")}
    melhor = min(res, key=lambda c: len(res[c][1])); j, ruins = res[melhor]
    print(f"\n>>> Melhor aderência ao CSV: '{melhor}' ({len(ruins)} diferenças).  valor = saldo em aberto; valor_original = valor original da parcela")
    if len(ruins):
        ruins.reset_index().to_csv("diferencas_arealizar.csv", sep=";", decimal=",", index=False, encoding="utf-8-sig")
        print("\nPor sentido (api - csv):"); print(pd.DataFrame({"so_api": [(ruins.csv == 0).sum()], "so_csv": [(ruins.api == 0).sum()], "valor_diferente": [((ruins.api != 0) & (ruins.csv != 0)).sum()]}).to_string(index=False))
        print("\nMaiores diferenças:"); print(ruins.sort_values("dif", key=abs, ascending=False).head(15).to_string())
        print("\nLista completa gravada em diferencas_arealizar.csv")
