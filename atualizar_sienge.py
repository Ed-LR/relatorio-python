"""Atualização diária dos dados do Sienge (Realizado + A Realizar) para o painel de impactos das obras no caixa.

Roda os três relatórios em sequência, usando as funções já validadas nos scripts de teste:
  - Realizado           -> /bank-movement          (sienge_teste_realizado.py)
  - A Realizar (pagar)  -> /outcome + /taxes       (sienge_teste_arealizar.py)
  - A Realizar (receber)-> /income                 (sienge_teste_income.py)

Saída (pasta saida/):
  realizado.csv            histórico do Realizado (atualização incremental, ver --janela-dias)
  arealizar.csv            pagar + receber juntos (coluna 'valor' = saldo em aberto)
  pendencias_contas.csv    contas fora da Baseapropriação / regras (só existe se houver pendência)
  centros_novos.csv        centros que apareceram no Sienge e não estão na sua lista (apague as linhas depois de tratar)
  status.json              resultado da última execução (para conferir pelo painel ou por e-mail)

Uso:
  python atualizar_sienge.py                    execução normal (a do agendador)
  python atualizar_sienge.py --completo         refaz todo o histórico do Realizado
  python atualizar_sienge.py --reusar           usa os JSONs brutos já baixados (nenhuma chamada à API)
  python atualizar_sienge.py --criar-lista-centros   cria centros_ativos.csv a partir do relatorio_centro_de_custo.xlsx (uma vez)
  python atualizar_sienge.py --max-obs 3000     limite de consultas de observação nesta execução (padrão 1500)

Variáveis de ambiente: SIENGE_SUBDOMINIO, SIENGE_USUARIO, SIENGE_SENHA, SIENGE_INDEXADOR (0 = REAL)
Arquivos na mesma pasta: sienge_teste_*.py, observacoes.py, gerar_dashboard.py, modelo_dashboard.html (o seu dashboard_resultado_obra_geral.html),
Baseapropriação.xlsx, BaseEquipe.xlsx e a lista de centros (centros_ativos.csv OU relatorio_centro_de_custo.xlsx).
Ao final, se os dados foram gravados sem erro, chama gerar_dashboard.py, que grava os painéis em DASHBOARD_DESTINO (pasta do OneDrive).
"""
import os, sys, json, glob, shutil, argparse, datetime, traceback

PASTA = os.path.dirname(os.path.abspath(__file__))
os.chdir(PASTA)                      # o Agendador de Tarefas não roda na pasta do script
sys.path.insert(0, PASTA)

import pandas as pd

# ---------------------------------------------------------------- configuração
SAIDA = "saida"
HISTORICO_INI = "2024-01-01"         # início do Realizado (o Realizado.csv começa em jan/2024)
AREALIZAR_INI = "2024-01-01"         # há parcelas em aberto de 2024; o teste validado usou 2025-01-01
AREALIZAR_FIM = "2030-12-31"
JANELA_DIAS = 90                     # no dia a dia, o Realizado rebusca só os últimos N dias
MANTER_LOGS = 30                     # dias de log guardados
CACHE_CR = "observacoes_cr_cache.json"   # observações das contas a receber (semeadas por semear_observacoes.py)
MAX_OBS = 1500                       # consultas de observação (GET /bills/{id}) por execução; o histórico completa em alguns dias
QUEDA_MAXIMA = 0.5                   # recusa substituir um arquivo se o novo tiver menos de 50% das linhas do anterior
GERAR_DASHBOARD = [sys.executable, "gerar_dashboard.py"]   # rodado só se os dados foram gravados sem erro; None desliga

TOTAIS = {"CP": {}, "CR": {}}              # 'Valor total título' = soma das parcelas do título (preenchido nas etapas de A Realizar)
COLUNAS_AREALIZAR = ["data", "doc", "lanc", "centro", "conta", "valor", "valor_original", "origem", "previsao", "credor", "observacao", "valor_titulo", "categoria", "tipo_conta"]


# ---------------------------------------------------------------- log (tela + arquivo)
class Tee:
    def __init__(self, *fs): self.fs = fs
    def write(self, s):
        for f in self.fs:
            try: f.write(s); f.flush()
            except Exception: pass
    def flush(self): pass

def iniciar_log():
    os.makedirs("logs", exist_ok=True)
    arq = open(os.path.join("logs", f"atualizar_{datetime.date.today().isoformat()}.log"), "a", encoding="utf-8")
    sys.stdout = sys.stderr = Tee(sys.__stdout__, arq)
    print(f"\n{'=' * 70}\n{datetime.datetime.now():%Y-%m-%d %H:%M:%S}  início da atualização")
    for f in sorted(glob.glob("logs/atualizar_*.log"))[:-MANTER_LOGS]:
        try: os.remove(f)
        except OSError: pass


# ---------------------------------------------------------------- utilitários
def centros():
    """Lista de centros de custo em andamento. centros_ativos.csv (coluna 'centro') tem prioridade sobre o xlsx."""
    if os.path.exists("centros_ativos.csv"):
        c = pd.read_csv("centros_ativos.csv", sep=None, engine="python", dtype=str, encoding="utf-8-sig")
        col = "centro" if "centro" in c.columns else c.columns[0]
        lista = sorted(int(x) for x in c[col].dropna().str.extract(r"(\d+)")[0].dropna())
        print(f"{len(lista)} centros em andamento (centros_ativos.csv)")
        return lista
    from sienge_teste_realizado import centros_ativos
    lista = centros_ativos()
    print(f"{len(lista)} centros em andamento (relatorio_centro_de_custo.xlsx)")
    return lista

def gravar(df, nome, colunas=None):
    """Grava em saida/ só se o arquivo novo for plausível; mantém uma cópia .anterior. Escrita atômica."""
    os.makedirs(SAIDA, exist_ok=True)
    destino = os.path.join(SAIDA, nome)
    if colunas: df = df.reindex(columns=colunas)
    if os.path.exists(destino):
        antes = sum(1 for _ in open(destino, encoding="utf-8-sig")) - 1
        if antes > 20 and len(df) < antes * QUEDA_MAXIMA:
            raise RuntimeError(f"{nome}: o novo tem {len(df)} linhas contra {antes} do anterior; não substituí (conferir manualmente).")
        shutil.copy2(destino, destino + ".anterior")
    tmp = destino + ".tmp"
    df.to_csv(tmp, sep=";", decimal=",", index=False, encoding="utf-8-sig")
    os.replace(tmp, destino)
    print(f"  gravado {destino}: {len(df)} linhas")

def so_pendentes(df):
    return df[df.categoria.isna()] if len(df) else df


# ---------------------------------------------------------------- relatórios
def realizado(ativos, completo, reusar, ini_arg, fim_arg):
    import sienge_teste_realizado as R
    hoje = datetime.date.today()
    fim = fim_arg or hoje.isoformat()
    hist_path = os.path.join(SAIDA, "realizado.csv")
    tem_hist = os.path.exists(hist_path) and not completo
    ini = ini_arg or ((hoje - datetime.timedelta(days=JANELA_DIAS)).isoformat() if tem_hist else HISTORICO_INI)
    print(f"\n[Realizado] {'incremental' if tem_hist else 'completo'}: {ini} a {fim}")
    if reusar:
        movs = json.load(open(f"bruto_{ini}_{fim}.json", encoding="utf-8"))["data"]; print("  reaproveitando arquivo bruto (sem API)")
    else:
        movs = R.buscar(ini, fim, ativos)
    base = R.linhas(movs, ativos)
    if len(base):                                                    # origem do título (CP/CR): decide qual endereço de observação vale
        base["origem"] = base.lanc.map({f"{m['billId']}/{m['installmentId']}": m.get("bankMovementOriginId") for m in movs if m.get("billId")})
    novo = R.classificar(base)
    if tem_hist:
        antigo = pd.read_csv(hist_path, sep=";", decimal=",", encoding="utf-8-sig", dtype={"centro": str, "conta": str})
        novo = pd.concat([antigo[(antigo.data < ini) | (antigo.data > fim)], novo], ignore_index=True)     # troca só a janela [ini, fim]
    return novo.sort_values(["data", "doc", "lanc"], kind="stable").reset_index(drop=True)

def arealizar_pagar(ativos, reusar):
    import sienge_teste_arealizar as A
    ini, fim = A_INI, A_FIM
    print(f"\n[A Realizar - pagar] {ini} a {fim}")
    if reusar:
        parc = json.load(open(f"bruto_outcome_{ini}_{fim}.json", encoding="utf-8"))["data"]; print("  reaproveitando arquivo bruto (sem API)")
    else:
        parc = A.buscar(ini, fim, ativos)
    cs = set(ativos)
    com_imposto = [p["billId"] for p in parc if p.get("taxAmount") and p.get("balanceAmount")
                   and any(c["costCenterId"] in cs for c in (p.get("paymentsCategories") or []))]
    impostos = A.buscar_impostos(com_imposto)                      # só consulta títulos novos (cache impostos_cache.json)
    df = A.classificar(A.linhas(parc, ativos, impostos).assign(hist=""))
    cred = {p["billId"]: (p.get("creditorName") or "").strip() for p in parc}
    for p in parc: TOTAIS["CP"][p["billId"]] = TOTAIS["CP"].get(p["billId"], 0) + (p.get("originalAmount") or 0)
    if len(df):
        bill = df.lanc.str.split("/").str[0].astype(int)
        df["credor"] = [f"Imposto a reter {c[9:]} {cred.get(b, '')}".strip() if c.startswith("retenção ") else c for c, b in zip(df.credor, bill)]
        df["valor_titulo"] = bill.map(TOTAIS["CP"]).round(2)
    return df

def arealizar_receber(ativos, reusar):
    import sienge_teste_income as I
    print(f"\n[A Realizar - receber] {A_INI} a {A_FIM}")
    if reusar:
        dados = json.load(open(I.BRUTO, encoding="utf-8")); print("  reaproveitando arquivo bruto (sem API)")
    else:
        dados = I.baixar(A_INI, A_FIM, ativos)
    df, avisos = I.linhas(dados["data"], ativos)
    if avisos: print(f"  ATENÇÃO: {len(avisos)} títulos com retenção fora da regra do INSS (ficam como pendência): {avisos[:5]}")
    df = I.classificar(df.assign(hist="")) if hasattr(I, "classificar") else df
    df = df.rename(columns={"cliente": "credor"}); df["valor_original"] = df["valor"]
    cli = {p["billId"]: (p.get("clientName") or "").strip() for p in dados["data"]}
    for p in dados["data"]: TOTAIS["CR"][p["billId"]] = TOTAIS["CR"].get(p["billId"], 0) + (p.get("originalAmount") or 0)
    if len(df):
        bill = df.lanc.str.split("/").str[0].astype(int)
        df["credor"] = [f"Imposto a reter {c[9:]} {cli.get(b, '')}".strip() if c.startswith("retenção ") else c for c, b in zip(df.credor, bill)]
        df["valor_titulo"] = bill.map(TOTAIS["CR"]).round(2)
    return df



def completar_valor_titulo(df):
    """Realizado: 'Valor total título' = soma das parcelas do título (CP vem do /outcome, CR do /income). Linhas antigas já gravadas ficam como estão."""
    if "origem" not in df.columns: return df
    bill = pd.to_numeric(df.lanc.astype(str).str.split("/").str[0], errors="coerce")
    novo = pd.Series([TOTAIS.get(o, {}).get(b) if pd.notna(b) else None for o, b in zip(df.origem, bill)], index=df.index, dtype=float)
    antigo = df["valor_titulo"] if "valor_titulo" in df.columns else pd.Series(float("nan"), index=df.index)
    df["valor_titulo"] = novo.round(2).where(novo.notna(), antigo)
    return df

def etapa_observacoes(resultados, offline, max_obs):
    """Preenche 'observacao' (campo notes de GET /bills/{id}) nos títulos a pagar. Receber (CR) fica de fora (endpoint sem permissão)."""
    import observacoes as OBS
    bill = lambda df: df.lanc.astype(str).str.split("/").str[0]
    mascaras = {}
    if "realizado" in resultados:
        r = resultados["realizado"]
        mascaras["realizado"] = (r.origem.notna() & (r.origem != "CR")) if "origem" in r.columns else pd.Series(False, index=r.index)
    if "a_pagar" in resultados:
        p = resultados["a_pagar"]; mascaras["a_pagar"] = p.valor != 0           # só o que está em aberto aparece no painel
    ids = set()
    for nome, m in mascaras.items(): ids |= set(bill(resultados[nome][m]))
    cache, resumo = OBS.consultar(ids, max_obs, offline)
    for nome, m in mascaras.items():
        df = resultados[nome]
        antigo = df["observacao"] if "observacao" in df.columns else pd.Series([None] * len(df), index=df.index, dtype=object)
        df["observacao"] = bill(df).map(cache).where(m & bill(df).isin(cache.keys()), antigo)
    if os.path.exists(CACHE_CR):                                      # contas a receber: observação semeada dos seus relatórios (a API não tem permissão)
        cr = json.load(open(CACHE_CR, encoding="utf-8"))
        for nome in ("realizado", "a_receber"):
            if nome in resultados and "origem" in resultados[nome].columns:
                df = resultados[nome]; novo = bill(df).map(cr)
                antigo = df["observacao"] if "observacao" in df.columns else pd.Series([None] * len(df), index=df.index, dtype=object)
                df["observacao"] = novo.where((df.origem == "CR") & novo.notna(), antigo)
    return resumo

def avisar_centros_novos(ativos):
    """Compara o /cost-centers com os centros já vistos; o que for novo e não estiver na sua lista vai para saida/centros_novos.csv."""
    import requests
    auth = (os.environ["SIENGE_USUARIO"], os.environ["SIENGE_SENHA"])
    url = f"https://api.sienge.com.br/{os.environ['SIENGE_SUBDOMINIO']}/public/api/v1/cost-centers"
    todos, off = [], 0
    while True:
        r = requests.get(url, params={"limit": 200, "offset": off}, auth=auth, headers={"accept": "application/json"}, timeout=120)
        if r.status_code != 200: raise RuntimeError(f"/cost-centers respondeu {r.status_code}")
        j = r.json(); res = j.get("results", []); todos += res; off += 200
        if not res or off >= (j.get("resultSetMetadata") or {}).get("count", 0): break
    atuais = {str(c["id"]): c.get("name", "") for c in todos}
    arq = "centros_conhecidos.json"
    conhecidos = set(json.load(open(arq, encoding="utf-8"))) if os.path.exists(arq) else None   # 1ª vez: só registra, não avisa
    novos = [] if conhecidos is None else [(i, n) for i, n in atuais.items() if i not in conhecidos and int(i) not in set(ativos)]
    if novos:
        os.makedirs(SAIDA, exist_ok=True); dest = os.path.join(SAIDA, "centros_novos.csv"); novo_arq = not os.path.exists(dest)
        with open(dest, "a", encoding="utf-8-sig", newline="") as f:
            if novo_arq: f.write("centro;nome;visto_em\r\n")
            for i, n in sorted(novos): f.write(f"{i};{n};{datetime.date.today().isoformat()}\r\n")
        print(f"\nCENTROS NOVOS no Sienge fora da sua lista: {len(novos)} -> {dest}")
    json.dump(sorted(atuais), open(arq, "w", encoding="utf-8"))
    return len(novos)

def criar_lista_centros():
    """Gera centros_ativos.csv a partir do relatorio_centro_de_custo.xlsx (mesmo filtro de hoje). Depois disso o xlsx deixa de ser lido."""
    r = pd.read_excel("relatorio_centro_de_custo.xlsx", header=6, dtype=str).dropna(subset=["Projeto"])
    r = r[(r["Situação da obra/centro de custo"] == "Em andamento") & (r.Projeto == "OBRA EM ANDAMENTO")]
    ex = r["Obra/Centro de custo"].str.extract(r"^(\d+)\s*-?\s*(.*)$")
    out = pd.DataFrame({"centro": ex[0], "nome": ex[1]}).sort_values("centro")
    out.to_csv("centros_ativos.csv", sep=";", index=False, encoding="utf-8-sig")
    print(f"centros_ativos.csv criado com {len(out)} centros. A partir de agora é esta lista que vale: inclua/retire obras nela.")

# ---------------------------------------------------------------- principal
PLANILHAS_MANUAIS = [("Baseapropriação.xlsx", "Baseapropriação.xlsx"), ("BaseEquipe.xlsx", "BaseEquipe.xlsx"),
                     ("relatorio centro de custo.xlsx", "relatorio_centro_de_custo.xlsx")]     # (nome no OneDrive, nome nesta pasta)

def sincronizar_planilhas():
    """As planilhas que você mantém à mão ficam na pasta do OneDrive (DASHBOARD_DESTINO). Antes de rodar, traz a versão atual para esta pasta.
    Sem isso o painel usava cópias velhas: tipo de conta, coordenador e nome da conta saíam diferentes do painel manual."""
    destino = os.environ.get("DASHBOARD_DESTINO")
    if not destino or not os.path.isdir(destino):
        print("Aviso: DASHBOARD_DESTINO não definido; usando as planilhas desta pasta."); return
    for origem, local in PLANILHAS_MANUAIS:
        o = os.path.join(destino, origem)
        if not os.path.exists(o):
            print(f"  Aviso: {origem} não está na pasta do OneDrive; mantendo a cópia desta pasta."); continue
        if not os.path.exists(local) or open(o, "rb").read() != open(local, "rb").read():
            shutil.copy2(o, local); print(f"  planilha atualizada a partir do OneDrive: {origem}")
        else:
            print(f"  planilha igual à do OneDrive: {origem}")

def main():
    global A_INI, A_FIM
    ap = argparse.ArgumentParser()
    ap.add_argument("--completo", action="store_true", help="refaz todo o histórico do Realizado")
    ap.add_argument("--reusar", action="store_true", help="usa os JSONs brutos salvos (teste, sem API)")
    ap.add_argument("--criar-lista-centros", action="store_true", help="cria centros_ativos.csv a partir do xlsx (uma vez)")
    ap.add_argument("--max-obs", type=int, default=MAX_OBS, help="limite de consultas de observação nesta execução")
    ap.add_argument("--realizado-ini"); ap.add_argument("--realizado-fim")
    ap.add_argument("--arealizar-ini", default=AREALIZAR_INI); ap.add_argument("--arealizar-fim", default=AREALIZAR_FIM)
    a = ap.parse_args()
    A_INI, A_FIM = a.arealizar_ini, a.arealizar_fim

    if a.criar_lista_centros:
        criar_lista_centros(); return 0
    iniciar_log()
    faltam = [v for v in ("SIENGE_SUBDOMINIO", "SIENGE_USUARIO", "SIENGE_SENHA", "SIENGE_INDEXADOR") if v not in os.environ]
    if faltam and not a.reusar:
        print("Variáveis de ambiente ausentes:", ", ".join(faltam)); return 2

    print("Planilhas manuais:"); sincronizar_planilhas()
    status = {"inicio": datetime.datetime.now().isoformat(timespec="seconds"), "relatorios": {}}
    resultados, erros = {}, []
    try:
        ativos = centros()
    except (Exception, SystemExit) as e:
        print("Falha ao obter a lista de centros:", e); traceback.print_exc(); return 1

    etapas = [("realizado", lambda: realizado(ativos, a.completo, a.reusar, a.realizado_ini, a.realizado_fim)),
              ("a_pagar", lambda: arealizar_pagar(ativos, a.reusar)),
              ("a_receber", lambda: arealizar_receber(ativos, a.reusar))]
    for nome, fn in etapas:                                         # uma etapa com erro não derruba as outras
        try:
            resultados[nome] = fn(); status["relatorios"][nome] = {"ok": True, "linhas": len(resultados[nome])}
        except (Exception, SystemExit) as e:                        # SystemExit: os scripts de teste usam sys.exit no 429/400
            erros.append(nome); status["relatorios"][nome] = {"ok": False, "erro": str(e)[:300]}
            print(f"\n*** ERRO em '{nome}': {e}"); traceback.print_exc()

    if "realizado" in resultados: completar_valor_titulo(resultados["realizado"])
    try:
        status["observacoes"] = etapa_observacoes(resultados, a.reusar, a.max_obs)
        print(f"\nObservações dos títulos a pagar: {status['observacoes']}")
    except (Exception, SystemExit) as e:                            # não derruba o resto: sem observação o painel ainda é gerado
        status["observacoes"] = {"erro": str(e)[:300]}; print(f"\nAviso: observações não atualizadas ({e})"); traceback.print_exc()
    if not a.reusar:
        try: status["centros_novos"] = avisar_centros_novos(ativos)
        except (Exception, SystemExit) as e: print(f"\nAviso: não consegui comparar com /cost-centers ({e})")

    gravados = []
    try:
        if "realizado" in resultados:
            gravar(resultados["realizado"], "realizado.csv"); gravados.append("realizado")
        if "a_pagar" in resultados and "a_receber" in resultados:   # o A Realizar só é gravado completo (pagar + receber)
            ar = pd.concat([resultados["a_pagar"], resultados["a_receber"]], ignore_index=True)
            gravar(ar, "arealizar.csv", COLUNAS_AREALIZAR); gravados.append("arealizar")
        elif "a_pagar" in resultados or "a_receber" in resultados:
            print("\nA Realizar NÃO atualizado: uma das metades falhou; mantido o arquivo anterior.")
    except Exception as e:
        erros.append("gravacao"); status["gravacao_erro"] = str(e); print(f"\n*** ERRO ao gravar: {e}")

    pend = pd.concat([so_pendentes(d) for d in resultados.values() if len(d)], ignore_index=True) if resultados else pd.DataFrame()
    pf = os.path.join(SAIDA, "pendencias_contas.csv")
    if len(pend):
        r = pend.groupby("conta").valor.agg(linhas="count", valor_liquido="sum").round(2)
        os.makedirs(SAIDA, exist_ok=True); r.to_csv(pf, sep=";", decimal=",", encoding="utf-8-sig")
        print(f"\nPENDÊNCIAS de classificação: {len(pend)} linhas em {len(r)} contas -> {pf}")
    elif os.path.exists(pf):
        os.remove(pf)
    status["pendencias_linhas"] = int(len(pend))

    if GERAR_DASHBOARD and gravados and not erros:
        import subprocess
        print("\nGerando o painel...")
        rc = subprocess.call(GERAR_DASHBOARD)
        status["dashboard_rc"] = rc
        if rc: erros.append("dashboard")

    status["fim"] = datetime.datetime.now().isoformat(timespec="seconds"); status["gravados"] = gravados; status["erros"] = erros
    os.makedirs(SAIDA, exist_ok=True)
    json.dump(status, open(os.path.join(SAIDA, "status.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print(f"\n{status['fim']}  fim | gravados: {gravados or 'nada'} | erros: {erros or 'nenhum'}")
    return 1 if erros else 0

if __name__ == "__main__":
    sys.exit(main())
