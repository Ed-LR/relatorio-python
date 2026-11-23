"""Teste do A REALIZAR - contas a receber (/income) + conferência com as linhas CR do Arealizar.csv.
Uso:  python sienge_teste_income.py 2025-01-01 2030-12-31 [reusar]     (sem 'reusar' faz 1 chamada BULK)
Variáveis: SIENGE_SUBDOMINIO, SIENGE_USUARIO, SIENGE_SENHA, SIENGE_INDEXADOR (0 = REAL)
Retenções de INSS/ISS (contas 1080105 e 1080106): o /income só traz o total (taxAmount); a abertura segue a regra validada nos seus relatórios.
"""
import os, sys, json, requests, pandas as pd
from sienge_teste_realizado import centros_ativos, classificar

BRUTO = "income_bruto.json"
ALIQUOTA_INSS = 0.035                                   # INSS retido nas notas de serviço: 3,5% do total do título (25 de 26 notas; a 26ª é título de 2 parcelas)
CONTA_RET = {"INSS": "1080105", "ISS": "1080106"}

def baixar(ini, fim, ativos):
    auth = (os.environ["SIENGE_USUARIO"], os.environ["SIENGE_SENHA"])
    url = f"https://api.sienge.com.br/{os.environ['SIENGE_SUBDOMINIO']}/public/api/bulk-data/v1/income"
    base = {"startDate": ini, "endDate": fim, "selectionType": "D", "correctionIndexerId": os.environ.get("SIENGE_INDEXADOR", "0")}
    ids = sorted(ativos)
    for lista in (ids, ",".join(map(str, ids))):
        r = requests.get(url, params=dict(base, costCentersId=lista), auth=auth, headers={"accept": "application/json"}, timeout=900)
        print("Status:", r.status_code, "| bytes:", len(r.content))
        if r.status_code == 429: sys.exit("Limite diário atingido (429).")
        if r.status_code == 400: print("Resposta da API:", r.text[:300]); continue
        if r.status_code == 404: return {"data": []}
        r.raise_for_status()
        open(BRUTO, "wb").write(r.content); return r.json()
    sys.exit("A API recusou a consulta (400).")

def linhas(parcelas, ativos):
    ativos = set(ativos); out = []; avisos = []
    total_titulo, primeira = {}, {}
    for p in parcelas:                                      # total do título e primeira parcela
        b = p["billId"]; total_titulo[b] = total_titulo.get(b, 0) + (p.get("originalAmount") or 0)
        primeira[b] = min(primeira.get(b, 10**9), p["installmentId"])
    for p in parcelas:
        doc = f"{(p.get('documentIdentificationId') or '').strip()}.{(p.get('documentNumber') or '').strip()}"
        lanc = f"{p['billId']}/{p['installmentId']}"
        cats = [c for c in (p.get("receiptsCategories") or []) if c["costCenterId"] in ativos]
        if not cats: continue
        for c in cats:
            f = c["financialCategoryRate"] / 100
            out.append(dict(data=p["dueDate"], doc=doc, lanc=lanc, centro=str(c["costCenterId"]), conta=str(c["financialCategoryId"]).strip(),
                            valor=round(p["balanceAmount"] * f, 2), origem=p.get("originId") or "", previsao=p.get("documentForecast") or "", cliente=(p.get("clientName") or "").strip()))
        # retenção: uma vez por título (1ª parcela), só se a parcela está inteira em aberto e o imposto ainda não foi retido em recebimento
        tax = p.get("taxAmount") or 0
        ja_retido = any((r.get("taxAmount") or 0) > 0 for r in (p.get("receipts") or []))
        if tax > 0 and p["installmentId"] == primeira[p["billId"]] and not ja_retido and abs(p["balanceAmount"] - p["originalAmount"]) < 0.01:
            inss = round(total_titulo[p["billId"]] * ALIQUOTA_INSS, 2)
            if inss > tax + 0.02:                           # não cabe: não adivinha
                avisos.append((lanc, doc, tax)); partes = {"IMPOSTO_INDEFINIDO": tax}
            else:
                partes = {"INSS": inss, "ISS": round(tax - inss, 2)}
            por_centro = {}
            for c in cats: por_centro[c["costCenterId"]] = por_centro.get(c["costCenterId"], 0) + c["financialCategoryRate"]
            for nome, v in partes.items():
                for cc, taxa in por_centro.items():
                    val = round(v * taxa / 100, 2)
                    if val: out.append(dict(data=p["dueDate"], doc=doc, lanc=lanc, centro=str(cc), conta=CONTA_RET.get(nome, nome), valor=-val,
                                            origem=p.get("originId") or "", previsao="", cliente="retenção " + nome))
    d = pd.DataFrame(out)
    return (d[d.valor != 0] if len(d) else d), avisos

def csv_referencia(ini, fim, ativos):
    D = pd.read_csv("Arealizar.csv", sep=";", encoding="cp1252", dtype=str)
    D = D[D["Código centro de custo"].isin(map(str, ativos)) & (D.Origem == "CR")].copy()
    D["data"] = pd.to_datetime(D.Data, format="%d/%m/%Y").dt.strftime("%Y-%m-%d"); D = D[(D.data >= ini) & (D.data <= fim)]
    num = lambda s: s.str.replace(".", "", regex=False).str.replace(",", ".", regex=False).astype(float)
    ap = pd.read_excel("Baseapropriação.xlsx", dtype=str); ap.columns = ["cod", "cn", "tc", "cat"]
    ap["k"] = ap.cod.str.replace(r"\D", "", regex=True); nat = ap.drop_duplicates("k").set_index("k").cat
    D["conta"] = D.Conta.str.extract(r"^(\d+)")[0]; n = D.conta.map(nat); v = num(D.Valor)
    D["valor"] = v.where(n != "Despesa", -v)
    D["doc"] = D.Documento.str.strip(); D["lanc"] = D["Lançamento"].fillna(""); D["centro"] = D["Código centro de custo"]
    return D[D.valor != 0][["data", "doc", "lanc", "centro", "conta", "valor"]]

def comparar(api, ref):
    k = ["data", "doc", "lanc", "centro", "conta"]
    a = api.groupby(k).valor.sum().round(2); b = ref.groupby(k).valor.sum().round(2)
    j = pd.concat([a.rename("api"), b.rename("csv")], axis=1).fillna(0); j["dif"] = (j.api - j.csv).round(2); r = j[j.dif.abs() > 0.02]
    print(f"\nAPI: {len(api)} linhas, total {api.valor.sum():,.2f} | CSV: {len(ref)} linhas, total {ref.valor.sum():,.2f}")
    print(f"Chaves com diferença acima de 2 centavos: {len(r)} de {len(j)}")
    print(f"  só API: {(r.csv == 0).sum()} ({r[r.csv == 0].dif.sum():,.2f}) | só CSV: {(r.api == 0).sum()} ({r[r.api == 0].dif.sum():,.2f}) | valor diferente: {((r.api != 0) & (r.csv != 0)).sum()} ({r[(r.api != 0) & (r.csv != 0)].dif.sum():,.2f})")
    return j, r

if __name__ == "__main__":
    ini, fim = (sys.argv[1], sys.argv[2]) if len(sys.argv) > 2 else ("2025-01-01", "2030-12-31")
    ativos = centros_ativos(); print(f"{len(ativos)} centros em andamento")
    if len(sys.argv) > 3 and sys.argv[3] == "reusar":
        dados = json.load(open(BRUTO, encoding="utf-8")); print("Reaproveitando o arquivo bruto salvo (nenhuma chamada à API).")
    else:
        dados = baixar(ini, fim, ativos)
    api, avisos = linhas(dados["data"], ativos)
    if avisos: print(f"\nATENÇÃO: {len(avisos)} títulos com retenção que não cabem na regra do INSS (não classificados):", avisos[:5])
    api = classificar(api.assign(hist=""))
    api.to_csv("api_income_linhas.csv", sep=";", decimal=",", index=False, encoding="utf-8-sig")
    j, r = comparar(api, csv_referencia(ini, fim, ativos))
    if len(r):
        r.reset_index().to_csv("diferencas_income.csv", sep=";", decimal=",", index=False, encoding="utf-8-sig")
        print("\nMaiores diferenças:"); print(r.sort_values("dif", key=abs, ascending=False).head(12).to_string())
