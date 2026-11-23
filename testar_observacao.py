"""Sondagem: onde a API do Sienge entrega a OBSERVAÇÃO do título (coluna 'Observação' dos relatórios Realizado / A Realizar).
Os retornos bulk (bank-movement, outcome, income) não trazem esse texto; este script testa os endpoints de consulta de título.

Uso (pegue títulos da coluna Lançamento do relatório, o número antes da barra):
  python testar_observacao.py --pagar 738817 740362 744877 --receber 27639
Variáveis de ambiente: SIENGE_SUBDOMINIO, SIENGE_USUARIO, SIENGE_SENHA   (a senha nunca vai para o arquivo)
Faz poucas chamadas (uma por endereço testado e por título). Grava observacao_amostra.json: me envie esse arquivo.
"""
import os, sys, re, json, argparse, requests

ap = argparse.ArgumentParser()
ap.add_argument("--pagar", nargs="*", default=[], help="billId de títulos a pagar")
ap.add_argument("--receber", nargs="*", default=[], help="billId de títulos a receber")
a = ap.parse_args()
if not (a.pagar or a.receber): sys.exit("Informe ao menos um título: --pagar 738817  e/ou  --receber 27639")

auth = (os.environ["SIENGE_USUARIO"], os.environ["SIENGE_SENHA"])
base = f"https://api.sienge.com.br/{os.environ['SIENGE_SUBDOMINIO']}/public/api/v1"
H = {"accept": "application/json"}

# endereços candidatos (nem todos existem em todas as contas; o status 404 também é informação)
CAND = {
    "pagar":   ["/bills/{id}", "/bills/{id}/installments", "/bills/{id}/installments/1"],
    "receber": ["/accounts-receivable/receivable-bills/{id}", "/receivable-bills/{id}", "/income/{id}"],
}
CHAVE = re.compile(r"observ|note|remark|compl|descr|histor|memo|comment", re.I)

def achar(obj, caminho=""):
    """Percorre o JSON e devolve (caminho, valor) dos campos cujo nome parece ser observação/notas."""
    achados = []
    if isinstance(obj, dict):
        for k, v in obj.items():
            c = f"{caminho}.{k}" if caminho else k
            if CHAVE.search(k) and isinstance(v, (str, int, float)) and str(v).strip(): achados.append((c, str(v)[:120]))
            achados += achar(v, c)
    elif isinstance(obj, list):
        for i, v in enumerate(obj[:5]): achados += achar(v, f"{caminho}[{i}]")
    return achados

saida = []
for tipo, ids in (("pagar", a.pagar), ("receber", a.receber)):
    for bid in ids:
        for modelo in CAND[tipo]:
            url = base + modelo.format(id=bid)
            try:
                r = requests.get(url, auth=auth, headers=H, timeout=120)
            except Exception as e:
                print(f"{tipo} {bid} {modelo}: erro de conexão {e}"); continue
            if r.status_code == 429: print("Limite diário (429) atingido."); json.dump(saida, open("observacao_amostra.json", "w", encoding="utf-8"), ensure_ascii=False, indent=1); sys.exit(1)
            try: corpo = r.json()
            except Exception: corpo = r.text[:500]
            campos = achar(corpo) if r.status_code == 200 else []
            print(f"{tipo} {bid} {modelo}: status {r.status_code}" + (f" | campos de texto: {campos[:6]}" if campos else ""))
            texto = json.dumps(corpo, ensure_ascii=False)
            saida.append({"tipo": tipo, "billId": bid, "endereco": modelo, "status": r.status_code,
                          "campos_de_texto": campos, "resposta": corpo if len(texto) < 20000 else texto[:20000] + "...(cortado)"})

json.dump(saida, open("observacao_amostra.json", "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("\nGravado em observacao_amostra.json")
