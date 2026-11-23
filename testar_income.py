"""Teste do A Realizar - contas a receber (/income).

Uso:
    python testar_income.py 2025-01-01 2030-12-31          (1 chamada BULK)
    python testar_income.py 2025-01-01 2030-12-31 reusar   (usa o arquivo ja baixado, sem chamar a API)

Gera:
    income_bruto.json   resposta bruta da API
    income_linhas.csv   uma linha por parcela x centro de custo x conta
"""
import os
import sys
import json

import requests
import pandas as pd

from sienge_teste_realizado import centros_ativos

BRUTO = "income_bruto.json"
SAIDA = "income_linhas.csv"


def baixar(ini, fim, centros):
    sub = os.environ["SIENGE_SUBDOMINIO"]
    auth = (os.environ["SIENGE_USUARIO"], os.environ["SIENGE_SENHA"])
    url = f"https://api.sienge.com.br/{sub}/public/api/bulk-data/v1/income"
    base = {
        "startDate": ini,
        "endDate": fim,
        "selectionType": "D",  # D = data de vencimento da parcela
        "correctionIndexerId": os.environ.get("SIENGE_INDEXADOR", "0"),
    }
    ids = sorted(centros)
    formatos = [ids, ",".join(str(i) for i in ids)]
    r = None
    for i, f in enumerate(formatos):
        p = dict(base)
        p["costCentersId"] = f
        r = requests.get(url, params=p, auth=auth,
                         headers={"accept": "application/json"}, timeout=600)
        print("Status:", r.status_code, "| bytes:", len(r.content))
        if r.status_code == 400 and i == 0:
            print("Resposta 400:", r.text[:300])
            print("Tentando a lista de centros separada por virgula...")
            continue
        break
    if r.status_code == 404:
        print("404: nenhuma parcela encontrada.")
        return {"data": []}
    if r.status_code != 200:
        print(r.text[:500])
        sys.exit(1)
    with open(BRUTO, "wb") as fh:
        fh.write(r.content)
    return r.json()


def montar(dados, ativos):
    regs = dados.get("data", []) if isinstance(dados, dict) else dados
    linhas = []
    taxa_fora = 0
    com_imposto = 0
    for p in regs:
        doc = "{}.{}".format((p.get("documentIdentificationId") or "").strip(),
                             (p.get("documentNumber") or "").strip())
        lanc = "{}/{}".format(p.get("billId"), p.get("installmentId"))
        cats = p.get("receiptsCategories") or []
        soma = sum((c.get("financialCategoryRate") or 0) for c in cats)
        if cats and abs(soma - 100) > 0.01:
            taxa_fora += 1
        achou = False
        for c in cats:
            cc = c.get("costCenterId")
            if cc not in ativos:
                continue
            achou = True
            taxa = (c.get("financialCategoryRate") or 0) / 100.0
            linhas.append({
                "data": p.get("dueDate"),
                "doc": doc,
                "lanc": lanc,
                "centro": cc,
                "conta": (c.get("financialCategoryId") or "").strip(),
                "conta_nome": c.get("financialCategoryName"),
                "origem": p.get("originId"),
                "cliente": p.get("clientName"),
                "previsao": p.get("documentForecast"),
                "valor": round((p.get("balanceAmount") or 0) * taxa, 2),
                "valor_corrigido": round((p.get("correctedBalanceAmount") or 0) * taxa, 2),
                "valor_original": round((p.get("originalAmount") or 0) * taxa, 2),
                "imposto_parcela": p.get("taxAmount") or 0,
            })
        if achou and (p.get("taxAmount") or 0) > 0:
            com_imposto += 1
    return pd.DataFrame(linhas), len(regs), taxa_fora, com_imposto


def main():
    if len(sys.argv) < 3:
        print(__doc__)
        sys.exit(1)
    ini, fim = sys.argv[1], sys.argv[2]
    reusar = len(sys.argv) > 3 and sys.argv[3] == "reusar"

    ativos = set(int(str(x).strip()) for x in centros_ativos())
    print(len(ativos), "centros em andamento")

    if reusar:
        print("Reaproveitando o arquivo bruto salvo (nenhuma chamada a API).")
        with open(BRUTO, "r", encoding="utf-8") as fh:
            dados = json.load(fh)
    else:
        dados = baixar(ini, fim, ativos)

    df, n_parcelas, taxa_fora, com_imposto = montar(dados, ativos)
    print("Parcelas devolvidas pela API:", n_parcelas)
    if df.empty:
        print("Nenhuma linha nos centros em andamento.")
        return

    zeros = int((df["valor"] == 0).sum())
    df = df[df["valor"] != 0].copy()
    df.to_csv(SAIDA, index=False, encoding="utf-8-sig")

    print("\nLinhas (parcela x centro x conta):", len(df), "| zeradas descartadas:", zeros)
    print("Total saldo em aberto:      {:,.2f}".format(df["valor"].sum()))
    print("Total saldo corrigido:      {:,.2f}".format(df["valor_corrigido"].sum()))
    print("Total valor original:       {:,.2f}".format(df["valor_original"].sum()))
    print("Parcelas com imposto retido:", com_imposto)
    print("Parcelas cujo rateio nao soma 100%:", taxa_fora)
    print("\nPor origem:")
    print(df.groupby("origem")["valor"].agg(["count", "sum"]).round(2))
    print("\nPor conta (10 maiores):")
    print(df.groupby(["conta", "conta_nome"])["valor"].agg(["count", "sum"])
          .sort_values("sum", key=abs, ascending=False).head(10).round(2))
    print("\nLinhas gravadas em", SAIDA)


if __name__ == "__main__":
    main()
