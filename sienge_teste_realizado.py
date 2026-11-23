"""Teste do Realizado via API do Sienge (bank-movement) + conferência com o Realizado.csv.
Uso:  python sienge_teste_realizado.py 2026-01-01 2026-01-31 [reusar]
Variáveis de ambiente: SIENGE_SUBDOMINIO, SIENGE_USUARIO, SIENGE_SENHA  (nunca grave a senha no arquivo).
Arquivos na mesma pasta: relatorio_centro_de_custo.xlsx, Baseapropriação.xlsx, Realizado.csv
"""
import os, sys, json, requests, pandas as pd

BASE = "https://api.sienge.com.br/{sub}/public/api/bulk-data/v1/bank-movement"

def centros_ativos(path="relatorio_centro_de_custo.xlsx"):
    r = pd.read_excel(path, header=6, dtype=str).dropna(subset=["Projeto"])
    r = r[(r["Situação da obra/centro de custo"] == "Em andamento") & (r.Projeto == "OBRA EM ANDAMENTO")]
    return sorted(int(c) for c in r["Obra/Centro de custo"].str.extract(r"^(\d+)")[0])

def buscar(inicio, fim, centros):
    auth = (os.environ["SIENGE_USUARIO"], os.environ["SIENGE_SENHA"])
    url = BASE.format(sub=os.environ["SIENGE_SUBDOMINIO"])
    for params in ({"startDate": inicio, "endDate": fim, "costCentersId": centros},
                   {"startDate": inicio, "endDate": fim, "costCentersId": ",".join(map(str, centros))}):
        r = requests.get(url, params=params, auth=auth, headers={"accept": "application/json"}, timeout=600)
        print("Status:", r.status_code, "| bytes:", len(r.content))
        if r.status_code == 404: return []
        if r.status_code == 429: sys.exit("Limite diário atingido (429). Tente amanhã.")
        if r.status_code == 400: continue          # tenta o formato alternativo da lista
        r.raise_for_status()
        with open(f"bruto_{inicio}_{fim}.json", "w", encoding="utf-8") as f: f.write(r.text)
        return r.json().get("data", [])
    sys.exit("A API recusou o formato da lista de centros de custo (400).")

def linhas(movs, ativos):
    ativos = set(ativos); out = []
    for m in movs:
        doc = f"{(m.get('documentIdentificationId') or '').strip()}.{(m.get('documentIdentificationNumber') or '').strip()}"
        lanc = f"{m['billId']}/{m['installmentId']}" if m.get("billId") else ""
        sinal = 1 if m.get("bankMovementOperationType") == "E" else -1     # entrada +, saída -
        for fc in m.get("financialCategories") or []:
            if fc["costCenterId"] not in ativos: continue                  # a API devolve o rateio inteiro
            v = round(m["bankMovementAmount"] * fc["financialCategoryRate"] / 100, 2)
            out.append(dict(data=m["bankMovementDate"], doc=doc, lanc=lanc, centro=str(fc["costCenterId"]),
                            conta=str(fc["financialCategoryId"]).strip(), valor=sinal * v,
                            credor=(m.get("creditorName") or m.get("clientName") or "").strip(),
                            hist=m.get("bankMovementHistoricName") or ""))
    d = pd.DataFrame(out)
    return d[(d.valor != 0) & ~d.doc.str.contains("PRV")] if len(d) else d      # mesmas regras do painel

CONFIRMADAS = {"1020103": ("Receita", "")}   # mapeamentos já confirmados (Rendimento de Poupança = Receita)
REGRA_2120101 = (("prv.equipe", "EQUIPE"), ("prv.custo", "CUSTO DIRETO"))   # aprovado: pelo nome do documento; o que fugir disso fica em pendência

def classificar(api, caminho="Baseapropriação.xlsx"):
    """Tipo de conta e Receita/Despesa vêm da SUA Baseapropriação. Nada é classificado em silêncio."""
    ap = pd.read_excel(caminho, dtype=str); ap.columns = ["cod", "cn", "tc", "cat"]
    ap["k"] = ap.cod.str.replace(r"\D", "", regex=True); ap = ap.drop_duplicates("k").set_index("k")
    api = api.copy()
    api["categoria"] = api.conta.map(ap.cat); api["tipo_conta"] = api.conta.map(ap.tc)
    for k, (c, t) in CONFIRMADAS.items():
        api.loc[api.conta == k, ["categoria", "tipo_conta"]] = [c, t]
    m = api.conta == "2120101"                                   # conta 2120101: Equipe x Custo Direto pelo documento
    api.loc[m, ["categoria", "tipo_conta"]] = [None, None]
    d = api.doc.str.strip().str.lower()
    for prefixo, tipo in REGRA_2120101:
        api.loc[m & d.str.startswith(prefixo), ["categoria", "tipo_conta"]] = ["Despesa", tipo]
    pend = api[api.categoria.isna()]
    if len(pend):
        r = pend.groupby("conta").valor.agg(linhas="count", valor_liquido="sum").round(2)
        r["motivo"] = ["2120101 fora da regra PRV.equipe / PRV.custo" if c == "2120101" else "conta nova, fora da Baseapropriação" for c in r.index]
        r.to_csv("pendencias_contas.csv", sep=";", decimal=",", encoding="utf-8-sig")
        print(f"\nPENDÊNCIAS: {len(pend)} linhas, valor líquido {pend.valor.sum():,.2f} -> pendencias_contas.csv")
        print(r.to_string())
    else:
        print("\nSem pendências de classificação.")
    return api

def csv_referencia(inicio, fim, ativos):
    D = pd.read_csv("Realizado.csv", sep=";", encoding="cp1252", dtype=str)
    D = D[D["Código centro de custo"].isin(map(str, ativos)) & ~D.Documento.str.contains("PRV", na=False)].copy()
    D["data"] = pd.to_datetime(D.Data, format="%d/%m/%Y").dt.strftime("%Y-%m-%d")
    D = D[(D.data >= inicio) & (D.data <= fim)]
    num = lambda s: s.str.replace(".", "", regex=False).str.replace(",", ".", regex=False).astype(float)
    ap = pd.read_excel("Baseapropriação.xlsx", dtype=str); ap.columns = ["cod", "cn", "tc", "cat"]
    ap["k"] = ap.cod.str.replace(r"\D", "", regex=True); nat = ap.drop_duplicates("k").set_index("k").cat
    D["conta"] = D.Conta.str.extract(r"^(\d+)")[0]
    n = D.conta.map(nat); n[D.conta == "2120101"] = "Despesa"; n[D.conta == "1020103"] = "Receita"
    v = num(D.Valor); D["valor"] = v.where(n != "Despesa", -v)                 # convenção do painel
    D["lanc"] = D["Lançamento"].fillna(""); D["doc"] = D.Documento.str.strip(); D["centro"] = D["Código centro de custo"]
    return D[D.valor != 0][["data", "doc", "lanc", "centro", "conta", "valor"]]

def comparar(api, ref):
    k = ["data", "doc", "lanc", "centro", "conta"]
    a = api.groupby(k).valor.sum().round(2); b = ref.groupby(k).valor.sum().round(2)
    j = pd.concat([a.rename("api"), b.rename("csv")], axis=1).fillna(0); j["dif"] = (j.api - j.csv).round(2)
    print(f"\nAPI: {len(api)} linhas, total {api.valor.sum():,.2f} | CSV: {len(ref)} linhas, total {ref.valor.sum():,.2f}")
    ruins = j[j.dif.abs() > 0.02]
    print(f"Chaves com diferença acima de 2 centavos: {len(ruins)} de {len(j)}")
    if len(ruins):
        pd_ = ruins.reset_index().groupby("data").agg(chaves=("dif", "size"), dif_total=("dif", "sum")).round(2)
        print("\nDiferenças agrupadas por data (api - csv):"); print(pd_.to_string())
        ruins.reset_index().to_csv("diferencas_api_vs_csv.csv", sep=";", decimal=",", index=False, encoding="utf-8-sig")
        ultima = j.index.get_level_values("data").max()
        outras = ruins[ruins.index.get_level_values("data") != ultima]
        if len(outras):
            print(f"\nDiferenças em datas anteriores a {ultima} ({len(outras)} chaves):"); print(outras.head(40).to_string())
        print("\nMaiores diferenças:"); print(ruins.sort_values("dif", key=abs, ascending=False).head(15).to_string())
        print("\nLista completa gravada em diferencas_api_vs_csv.csv")
    return j

if __name__ == "__main__":
    ini, fim = (sys.argv[1], sys.argv[2]) if len(sys.argv) > 2 else ("2026-01-01", "2026-01-31")
    ativos = centros_ativos()
    print(f"{len(ativos)} centros em andamento")
    if len(sys.argv) > 3 and sys.argv[3] == "reusar":      # usa a resposta já baixada, sem nova chamada à API
        movs = json.load(open(f"bruto_{ini}_{fim}.json", encoding="utf-8"))["data"]; print("Reaproveitando o arquivo bruto salvo (nenhuma chamada à API).")
    else:
        movs = buscar(ini, fim, ativos)
    api = classificar(linhas(movs, ativos))
    api.to_csv(f"api_realizado_{ini}_{fim}.csv", sep=";", decimal=",", index=False, encoding="utf-8-sig")
    comparar(api, csv_referencia(ini, fim, ativos))
