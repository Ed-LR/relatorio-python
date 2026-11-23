"""Gera o painel "Impacto das Obras no Caixa da CYMZ" a partir dos CSVs da API (saida/realizado.csv e saida/arealizar.csv).

O HTML modelo (modelo_dashboard.html = o dashboard_resultado_obra_geral.html que você já usa) é mantido como está: o script troca só os dados
embutidos (const DB = {...} e o bloco detail-data). Layout, filtros e gráficos continuam exatamente os do modelo.

Saída (pasta DASHBOARD_DESTINO, padrão saida/paineis; aponte para uma pasta sincronizada com o OneDrive):
  Dashboard_Impacto_das_obra_geral.html
  Dashboard_Impacto_das_obra_<Coordenador>.html   (um por coordenador, só com as obras dele)

Regras (as do painel atual):
  - só centros da lista centros_ativos.csv (Em andamento / OBRA EM ANDAMENTO); o Realizado exclui documentos PRV
  - Realizado usa a data do pagamento/recebimento; A Realizar usa o vencimento, e o que venceu antes da data de processamento vai para 31/01/2027
    (data original preservada em originalDate/originalDueDate)
  - valor do detalhe = valor do relatório do Sienge (despesa positiva, retenção negativa); o painel aplica o sinal pelo impactSign
  - contas sem classificação na Baseapropriação NÃO entram: são contadas em summary.unclassifiedFacts e listadas em saida/pendencias_contas.csv

Uso:  python gerar_dashboard.py [--modelo modelo_dashboard.html] [--destino PASTA] [--data-processamento AAAA-MM-DD] [--sem-individuais]
"""
import os, sys, re, json, gzip, base64, argparse, datetime, unicodedata
import pandas as pd

PASTA = os.path.dirname(os.path.abspath(__file__))
os.chdir(PASTA)

SAIDA = "saida"
VENCIDOS_PARA = "2027-01-31"
ANOS = [2026, 2027, 2028, 2029, 2030]          # períodos gravados em cada conta (o painel mostra os anos com movimento)
ANO_BASE = ANOS[0]                             # tudo antes de 01/01/2026 vira "saldo anterior" (prior)
PROJETO, SITUACAO = "OBRA EM ANDAMENTO", "Em andamento"
STATUS_ORDEM = {"Realizado": 0, "A realizar": 1}
CAMPOS_DETALHE = ["status", "centerCode", "centerName", "coordinator", "engineer", "foreman", "administrative", "project", "statusCenter",
                  "accountCode", "accountName", "category", "nature", "amount", "titleAmount", "date", "paymentDate", "dueDate",
                  "originalDueDate", "originalDate", "supplier", "complement", "observation", "title", "document", "workLabel", "businessArea"]


# ---------------------------------------------------------------- leitura das bases
def so_digitos(x): return re.sub(r"\D", "", str(x))

def ler_centros():
    """centros_ativos.csv (centro;nome). Sem arquivo, cai no relatorio_centro_de_custo.xlsx (mesmo filtro de antes)."""
    if os.path.exists("centros_ativos.csv"):
        c = pd.read_csv("centros_ativos.csv", sep=None, engine="python", dtype=str, encoding="utf-8-sig")
        col = "centro" if "centro" in c.columns else c.columns[0]
        nome = c["nome"] if "nome" in c.columns else pd.Series([""] * len(c))
        return {so_digitos(a): (str(b) if pd.notna(b) else "").strip() for a, b in zip(c[col], nome) if so_digitos(a)}
    r = pd.read_excel("relatorio_centro_de_custo.xlsx", header=6, dtype=str).dropna(subset=["Projeto"])
    r = r[(r["Situação da obra/centro de custo"] == SITUACAO) & (r.Projeto == PROJETO)]
    ex = r["Obra/Centro de custo"].str.extract(r"^(\d+)\s*-?\s*(.*)$")
    return {a: b.strip() for a, b in zip(ex[0], ex[1])}

def ler_equipe():
    e = pd.read_excel("BaseEquipe.xlsx", dtype=str)
    e.columns = [str(c).strip() for c in e.columns]
    col = lambda *nomes: next(c for c in e.columns if c.upper() in nomes)
    k, coord, eng, enc, adm = col("Nº CENTRO DE CUSTO", "N° CENTRO DE CUSTO"), col("COORDENADOR"), col("ENGENHEIRO"), col("ENCARREGADO"), col("ADM")
    e["k"] = e[k].map(so_digitos)
    e = e[e.k != ""].drop_duplicates("k", keep="last").set_index("k")
    f = lambda c: e[c].fillna("").astype(str).str.strip().to_dict()
    return f(coord), f(eng), f(enc), f(adm)

def ler_contas():
    ap = pd.read_excel("Baseapropriação.xlsx", dtype=str); ap.columns = ["cod", "cn", "tc", "cat"]
    ap["k"] = ap.cod.map(so_digitos); ap = ap[ap.k != ""].drop_duplicates("k").set_index("k")
    return ap.cod.str.strip().to_dict(), ap.cn.fillna("").str.strip().to_dict()

def ler_csv(nome):
    caminho = os.path.join(SAIDA, nome)
    if not os.path.exists(caminho): sys.exit(f"Não encontrei {caminho}. Rode antes o atualizar_sienge.py.")
    return pd.read_csv(caminho, sep=";", decimal=",", encoding="utf-8-sig", dtype={"centro": str, "conta": str, "lanc": str, "doc": str})


# ---------------------------------------------------------------- detalhes (uma linha por lançamento x centro x conta)
def montar_detalhes(real, areal, hoje, centros, equipe, contas):
    coord, eng, enc, adm = equipe; cod_conta, nome_conta = contas
    ativos = set(centros)
    txt = lambda s: s.fillna("").astype(str).str.strip()

    r = real[real.centro.isin(ativos) & (real.valor != 0) & ~real.doc.fillna("").str.contains("PRV")].copy()
    r["status"] = "Realizado"; r["date"] = r["originalDate"] = r["paymentDate"] = r.data; r["dueDate"] = r["originalDueDate"] = ""
    cr, hi = txt(r.credor), txt(r["hist"] if "hist" in r.columns else pd.Series("", index=r.index))
    r["supplier"] = cr; r["complement"] = [f"{h} - {c}" if h else c for h, c in zip(hi, cr)]

    a = areal[areal.centro.isin(ativos) & (areal.valor != 0)].copy()
    a["status"] = "A realizar"; a["originalDate"] = a.data
    vencido = a.data < hoje
    a["date"] = a.data.where(~vencido, VENCIDOS_PARA); a["dueDate"] = a["date"]; a["originalDueDate"] = a.data.where(vencido, ""); a["paymentDate"] = ""
    c = txt(a.credor); imposto = c.str.startswith("Imposto a reter"); receber = a.origem.fillna("") == "CR"
    a["supplier"] = c
    a["complement"] = [x if i else (f"A receber de {x}" if rc else f"A pagar para {x}") for x, i, rc in zip(c, imposto, receber)]

    d = pd.concat([r, a], ignore_index=True)
    d["observation"] = txt(d.observacao) if "observacao" in d.columns else ""
    nat = d.categoria
    sem_class = d[nat.isna()]
    d = d[nat.notna()].copy()
    d["nature"] = d.categoria
    d["amount"] = (d.valor * d.nature.map({"Despesa": -1}).fillna(1)).round(2)           # volta ao sinal do relatório do Sienge
    d["titleAmount"] = (d.valor_titulo if "valor_titulo" in d.columns else pd.Series(float("nan"), index=d.index))
    d["titleAmount"] = d.titleAmount.where(d.titleAmount.notna(), d.amount.abs()).round(2)
    d["centerCode"] = d.centro
    d["centerName"] = [f"{k} - {centros[k]}" if centros[k] else k for k in d.centro]
    d["workLabel"] = d.centerName
    d["coordinator"], d["engineer"], d["foreman"], d["administrative"] = (d.centro.map(m).fillna("") for m in (coord, eng, enc, adm))
    d["project"], d["statusCenter"], d["businessArea"] = PROJETO, SITUACAO, ""
    d["accountCode"] = [cod_conta.get(k) or (f"{k[0]}.{k[1:3]}.{k[3:5]}.{k[5:]}" if k.isdigit() and len(k) == 7 else k) for k in d.conta]
    d["accountName"] = [nome_conta.get(k) or k for k in d.conta]
    d["category"] = txt(d.tipo_conta)
    d["title"] = txt(d.lanc); d["document"] = txt(d.doc)
    d["_ord"] = d.status.map(STATUS_ORDEM)
    d = d.sort_values(["_ord", "centerCode", "accountCode", "date", "title"], kind="stable").reset_index(drop=True)
    nao_mapeadas = sorted({k for k in pd.concat([real.conta, areal.conta]).dropna() if k not in cod_conta})
    return d, sem_class, nao_mapeadas


# ---------------------------------------------------------------- fatos (soma por status x obra x conta)
def montar_fatos(d, centros_modelo):
    fatos = []
    for (st, cc, ac), g in d.groupby(["status", "centerCode", "accountCode"], sort=False):
        f0 = g.iloc[0]; sinal = -1 if f0.nature == "Despesa" else 1
        prior = round(float(g.loc[g.date < f"{ANO_BASE}-01-01", "amount"].sum()), 2) + 0.0
        periodos = {}
        for ano in ANOS:
            ga = g[g.date.str[:4] == str(ano)]
            por_mes = ga.groupby(ga.date.str[5:7].astype(int)).amount.sum()
            periodos[str(ano)] = [round(float(por_mes.get(m, 0.0)), 2) + 0.0 for m in range(1, 13)]
        fatos.append({"status": st, "centerCode": cc, "centerName": f0.centerName, "workCode": cc, "coordinator": f0.coordinator,
                      "engineer": f0.engineer, "foreman": f0.foreman, "administrative": f0.administrative,
                      "accountCode": ac, "accountName": f0.accountName, "category": f0.category, "nature": f0.nature,
                      "prior": prior, "months": periodos[str(ANO_BASE)], "periods": periodos, "type": "R", "impactSign": sinal,
                      "project": PROJETO, "statusCenter": SITUACAO, "workLabel": f0.workLabel, "businessArea": "", "detailCount": int(len(g))})
    return fatos


def anos_com_movimento(d):
    anos = [int(x[:4]) for x in d.date if x[:4].isdigit() and int(x[:4]) >= ANO_BASE]
    return list(range(ANO_BASE, max(anos + [ANO_BASE]) + 1))


# ---------------------------------------------------------------- injeção no modelo
# Trecho do modelo que escreve "Atualização dd/mm/aaaa"; trocado na geração para mostrar também a hora (o modelo em si não é alterado).
JS_DATA_ANTIGO = "new Date(DB.updated+'T12:00:00').toLocaleDateString('pt-BR')"
JS_DATA_NOVO = ("(DB.updatedAt?(d=>d.toLocaleDateString('pt-BR')+' às '+d.toLocaleTimeString('pt-BR',{hour:'2-digit',minute:'2-digit'}))(new Date(DB.updatedAt))"
                ":new Date(DB.updated+'T12:00:00').toLocaleDateString('pt-BR'))")

def carregar_modelo(caminho):
    html = open(caminho, encoding="utf-8").read()
    i = html.index("const DB=") + len("const DB=")
    db, n = json.JSONDecoder().raw_decode(html[i:])
    if len(re.findall(r'id="detail-data"', html)) != 1: sys.exit("O modelo precisa ter exatamente um bloco detail-data.")
    return html, i, n, db

def gerar_html(html, i, n, db_modelo, fatos, detalhes, hoje, resumo, anos, atualizado_em=""):
    rotulo = (datetime.datetime.fromisoformat(atualizado_em).strftime("%d/%m/%Y às %H:%M") if atualizado_em
              else datetime.date.fromisoformat(hoje).strftime("%d/%m/%Y"))
    db = dict(db_modelo); db["updated"] = hoje; db["updatedAt"] = atualizado_em
    db["updatedLabel"] = rotulo                                   # texto pronto do cabeçalho (modelo novo)
    db["years"] = anos; db["summary"] = resumo; db["facts"] = fatos
    db["sources"] = {"realized": "API Sienge /bank-movement", "toRealize": "API Sienge /outcome e /income", "accounts": "Baseapropriação.xlsx",
                     "team": "BaseEquipe.xlsx", "centers": "centros_ativos.csv"}
    novo_db = json.dumps(db, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
    html = html[:i] + novo_db + html[i + n:]
    if "DB.updatedLabel" in html: pass                                                       # modelo novo já lê o texto pronto (updatedLabel)
    elif JS_DATA_ANTIGO in html: html = html.replace(JS_DATA_ANTIGO, JS_DATA_NOVO, 1)         # modelo antigo: troca o trecho para mostrar data + hora
    else: print("  Aviso: não achei o trecho da data no modelo; o painel mostrará só a data.")
    blob = base64.b64encode(gzip.compress(json.dumps(detalhes, ensure_ascii=False, separators=(",", ":")).encode("utf-8"), mtime=0)).decode("ascii")
    return re.sub(r'(<script[^>]*id="detail-data"[^>]*>)(.*?)(</script>)', lambda m: m.group(1) + blob + m.group(3), html, count=1, flags=re.S)

def gravar_html(texto, destino, nome):
    os.makedirs(destino, exist_ok=True)
    caminho = os.path.join(destino, nome); tmp = caminho + ".tmp"
    with open(tmp, "w", encoding="utf-8", newline="") as f: f.write(texto)
    os.replace(tmp, caminho)
    print(f"  {caminho}  ({len(texto) / 1e6:.1f} MB)")

def nome_arquivo(coordenador):
    s = unicodedata.normalize("NFKD", coordenador).encode("ascii", "ignore").decode()
    return "Dashboard_Impacto_das_obra_" + re.sub(r"[^A-Za-z0-9]+", "_", s).strip("_") + ".html"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--modelo", default="modelo_dashboard.html")
    ap.add_argument("--destino", default=os.environ.get("DASHBOARD_DESTINO", os.path.join(SAIDA, "paineis")))
    ap.add_argument("--data-processamento", default=datetime.date.today().isoformat())
    ap.add_argument("--sem-individuais", action="store_true")
    a = ap.parse_args()
    hoje = a.data_processamento
    atualizado_em = datetime.datetime.now().strftime("%Y-%m-%dT%H:%M") if hoje == datetime.date.today().isoformat() else ""   # hora local do PC

    if not os.path.exists(a.modelo): sys.exit(f"Modelo não encontrado: {a.modelo} (copie o seu dashboard_resultado_obra_geral.html para esse nome).")
    centros = ler_centros(); equipe = ler_equipe(); contas = ler_contas()
    real, areal = ler_csv("realizado.csv"), ler_csv("arealizar.csv")
    d, sem_class, nao_mapeadas = montar_detalhes(real, areal, hoje, centros, equipe, contas)
    fatos = montar_fatos(d, centros)
    n_sem = int(sem_class.groupby(["status", "centro", "conta"]).ngroups) if len(sem_class) else 0
    if n_sem: print(f"ATENÇÃO: {n_sem} combinações obra x conta sem classificação ficaram FORA do painel (veja saida/pendencias_contas.csv).")

    html, i, n, db_modelo = carregar_modelo(a.modelo)
    anos = anos_com_movimento(d)
    resumo = {"rows": int(len(d)), "centers": len(centros), "realizedRows": int((d.status == "Realizado").sum()),
              "toRealizeRows": int((d.status == "A realizar").sum()), "unmappedAccountCodes": len(nao_mapeadas), "unclassifiedFacts": n_sem}
    det_registros = d[CAMPOS_DETALHE].to_dict("records")
    print(f"Painel de {hoje}: {resumo['rows']} lançamentos ({resumo['realizedRows']} realizados, {resumo['toRealizeRows']} a realizar), "
          f"{len(fatos)} contas, {len(centros)} obras, anos {anos[0]}-{anos[-1]}")
    gravar_html(gerar_html(html, i, n, db_modelo, fatos, det_registros, hoje, resumo, anos, atualizado_em), a.destino, "Dashboard_Impacto_das_obra_geral.html")

    if not a.sem_individuais:
        for c in sorted({f["coordinator"] for f in fatos if f["coordinator"]}):
            fc = [f for f in fatos if f["coordinator"] == c]
            dc = [x for x in det_registros if x["coordinator"] == c]
            rc = {"rows": len(dc), "centers": len({f["centerCode"] for f in fc}), "realizedRows": sum(x["status"] == "Realizado" for x in dc),
                  "toRealizeRows": sum(x["status"] == "A realizar" for x in dc), "unmappedAccountCodes": 0, "unclassifiedFacts": 0}
            gravar_html(gerar_html(html, i, n, db_modelo, fc, dc, hoje, rc, anos, atualizado_em), a.destino, nome_arquivo(c))
    return 0

if __name__ == "__main__":
    sys.exit(main())
