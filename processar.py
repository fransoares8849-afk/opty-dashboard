"""
processar.py
Lê o Excel exportado do iris.opty.com.br e atualiza o index.html do dashboard SLA Opty.
"""

import pandas as pd
import json
import sys
import os
from datetime import datetime, timedelta
from collections import Counter

# ── Caminhos ────────────────────────────────────────────────────────────────
EXCEL_PATH  = sys.argv[1] if len(sys.argv) > 1 else "relatorio-agendamentos.xlsx"
HTML_TMPL   = sys.argv[2] if len(sys.argv) > 2 else "template.html"
HTML_OUT    = sys.argv[3] if len(sys.argv) > 3 else "index.html"

# ── Feriados nacionais (atualizar anualmente) ────────────────────────────────
HOLIDAYS_BR = {
    datetime(2026, 1, 1), datetime(2026, 2, 16), datetime(2026, 2, 17),
    datetime(2026, 4, 3), datetime(2026, 4, 21), datetime(2026, 5, 1),
    datetime(2026, 6, 4), datetime(2026, 9, 7),  datetime(2026, 10, 12),
    datetime(2026, 11, 2), datetime(2026, 11, 15), datetime(2026, 12, 25),
}

# ── Mapa Unidade → Regional ──────────────────────────────────────────────────
REG_MAP = {
    # BA
    "Dr Vis Garibaldi": "BA",
    "Hapclinicas": "BA",
    "Dr. Vis - Rio Vermelho": "BA",
    "Instituto de Olhos Villas": "BA",
    # DF
    "HOB | Unidade Hélio Prates": "DF",
    "HOB | L2 Sul": "DF",
    "HOB | Taguatinga Sul": "DF",
    "Visão - Filial Asa Norte": "DF",
    "Visão - Taguatinga": "DF",
    "Dr. Vis | Unidade Sobradinho": "DF",
    "COOA | Unidade Asa Norte": "DF",
    "Visão - Filial Gama": "DF",
    "Visão - Filial Samambaia": "DF",
    "HOG - Hospital de Olhos do Gama": "DF",
    # RJ
    "Dr Vis - São Cristovão": "RJ",
    "Eye center Nova Iguaçu": "RJ",
    "Eye Center Méier - Casarão": "RJ",
    "Eye Center Downtown": "RJ",
    "RIO MEDICINA OCULAR": "RJ",
    "CLÍNICA CECOF": "RJ",
    "CLÍNICA FOCUS": "RJ",
    "CLÍNICA ABRAÃO": "RJ",
    "HOBrasil - Meier": "RJ",
    "CLÍNICA CEO": "RJ",
    "CLÍNICA COA": "RJ",
    "CLÍNICA INODUC": "RJ",
    "CLÍNICA FONCHAM": "RJ",
    "CLÍNICA EDUARDO AZEVEDO": "RJ",
    "CLÍNICA CLINOP": "RJ",
    "CLÍNICA DR. OLHO": "RJ",
    "COSC - São Cristóvão": "RJ",
    "CLÍNICA SAUDE RIO": "RJ",
}

EM_ATENDIMENTO_STATUS = {
    "Em_atendimento", "Pendente_agenda", "Pendente_contato", "Pendente_documentação"
}

# ── Funções auxiliares ───────────────────────────────────────────────────────
def add_business_days(start, n):
    d = start
    added = 0
    while added < n:
        d += timedelta(days=1)
        if d.weekday() < 5 and d.replace(hour=0, minute=0, second=0, microsecond=0) not in HOLIDAYS_BR:
            added += 1
    return d

def business_days_between(start, end):
    if pd.isna(start) or pd.isna(end):
        return None
    count = 0
    d = start
    while d < end:
        d += timedelta(days=1)
        if d.weekday() < 5 and d.replace(hour=0, minute=0, second=0, microsecond=0) not in HOLIDAYS_BR:
            count += 1
    return count

# ── Leitura e limpeza ────────────────────────────────────────────────────────
print(f"📂 Lendo {EXCEL_PATH}...")
df = pd.read_excel(EXCEL_PATH)
df["Data Solicitação"]      = pd.to_datetime(df["Data Solicitação"],      dayfirst=True, errors="coerce")
df["Data Alteração Status"] = pd.to_datetime(df["Data Alteração Status"], dayfirst=True, errors="coerce")
df["CPF_norm"]  = df["CPF"].astype(str).str.replace(r"\D", "", regex=True).str.zfill(11)
df["Tel_norm"]  = df["Telefone"].astype(str).str.replace(r"\D", "", regex=True)
df["Nome_norm"] = df["Paciente"].astype(str).str.strip().str.upper()
df["dup_key"]   = df["CPF_norm"] + "|" + df["Tel_norm"] + "|" + df["Nome_norm"]

TOTAL_BRUTO = len(df)
print(f"   Total bruto: {TOTAL_BRUTO}")

# ── Deduplicação ─────────────────────────────────────────────────────────────
print("🔄 Deduplicando (CPF+Tel+Nome ≤ 20 dias)...")

def dedup(group):
    group = group.sort_values("Data Solicitação", ascending=False)
    keep = []
    kept_dates = []
    for idx, row in group.iterrows():
        d = row["Data Solicitação"]
        is_dup = any(
            pd.notna(d) and pd.notna(kd) and abs((d - kd).days) <= 20
            for kd in kept_dates
        )
        if not is_dup:
            keep.append(idx)
            kept_dates.append(d)
    return group.loc[keep]

df_dedup = df.groupby("dup_key", group_keys=False).apply(dedup).copy()
TOTAL_DUPS   = TOTAL_BRUTO - len(df_dedup)
TOTAL_UNICOS = len(df_dedup)
print(f"   Duplicatas removidas: {TOTAL_DUPS} | Únicos: {TOTAL_UNICOS}")

# ── Verificar unidades não mapeadas ──────────────────────────────────────────
unmapped = set(df_dedup["Unidade Preferência"].dropna().unique()) - set(REG_MAP.keys())
if unmapped:
    print(f"\n⚠️  UNIDADES NÃO MAPEADAS — adicione ao REG_MAP antes de continuar:")
    for u in sorted(unmapped):
        print(f"   \"{u}\": \"??\"")
    sys.exit(1)

# ── Processar linhas ─────────────────────────────────────────────────────────
print("⚙️  Processando registros...")
today = datetime.now()
rows  = []

for _, row in df_dedup.iterrows():
    unidade  = str(row["Unidade Preferência"]) if pd.notna(row["Unidade Preferência"]) else ""
    conv     = str(row["Convênio"]) if pd.notna(row["Convênio"]) else ""
    conv_s   = conv.replace(" Rede Exclusiva","").replace(" CAP DRVIS","").replace(" BARRIS CAP","").strip()
    status   = str(row["Status"])
    tipo     = str(row["Tipo"]) if pd.notna(row["Tipo"]) else ""
    reg      = REG_MAP.get(unidade, "N/A")
    data_ab  = row["Data Solicitação"]
    data_alt = row["Data Alteração Status"]
    prazo    = add_business_days(data_ab, 3) if pd.notna(data_ab) else None

    is_tratado       = status.startswith("Finalizado_") or status == "Cancelado"
    is_em_atendimento = status in EM_ATENDIMENTO_STATUS

    if is_tratado:
        dias_trat = business_days_between(data_ab, data_alt) if pd.notna(data_alt) else None
        classif   = "Tratado dentro do prazo" if (dias_trat is not None and dias_trat <= 3) else "Tratado fora do prazo"
        dias_hoje = int((today - data_ab).days) if pd.notna(data_ab) else None
    elif is_em_atendimento:
        dias_hoje = int((today - data_ab).days) if pd.notna(data_ab) else None
        classif   = "Em atendimento - dentro do prazo" if (prazo and today <= prazo) else "Em atendimento - fora do prazo"
        dias_trat = None
    else:
        dias_hoje = int((today - data_ab).days) if pd.notna(data_ab) else None
        classif   = "Em aberto - dentro do prazo" if (prazo and today <= prazo) else "Em aberto - fora do prazo"
        dias_trat = None

    rows.append({
        "reg": reg, "conv": conv_s, "unidade": unidade, "tipo": tipo,
        "status": status, "classif": classif, "isDup": False,
        "cpf": str(row["CPF_norm"]), "diasTrat": dias_trat, "diasHoje": dias_hoje,
        "dataAb":  data_ab.isoformat()  if pd.notna(data_ab)  else None,
        "dataAlt": data_alt.isoformat() if pd.notna(data_alt) else None,
        "prazo":   prazo.isoformat()    if prazo               else None,
    })

# ── Sumários ─────────────────────────────────────────────────────────────────
dups_rows = []
for _, row in df[~df.index.isin(df_dedup.index)].iterrows():
    unidade = str(row["Unidade Preferência"]) if pd.notna(row["Unidade Preferência"]) else ""
    conv    = str(row["Convênio"]) if pd.notna(row["Convênio"]) else ""
    conv_s  = conv.replace(" Rede Exclusiva","").replace(" CAP DRVIS","").replace(" BARRIS CAP","").strip()
    dups_rows.append({"reg": REG_MAP.get(unidade,"N/A"), "conv": conv_s})

dups_by_reg  = dict(Counter(r["reg"]  for r in dups_rows))
dups_by_conv = dict(Counter(r["conv"] for r in dups_rows))

bruto_by_reg = Counter(r["reg"]  for r in rows)
for r in dups_rows: bruto_by_reg[r["reg"]] += 1

bruto_by_conv = Counter(r["conv"] for r in rows)
for r in dups_rows: bruto_by_conv[r["conv"]] += 1

pd_data = {
    "rows": rows,
    "updatedAt":    today.strftime("%Y-%m-%dT%H:%M:%S"),
    "totalBruto":   TOTAL_BRUTO,
    "totalDups":    TOTAL_DUPS,
    "totalUnicos":  TOTAL_UNICOS,
    "dupsByReg":    dups_by_reg,
    "brutoByReg":   dict(bruto_by_reg),
    "unicosByReg":  dict(Counter(r["reg"]  for r in rows)),
    "dupsByConv":   dups_by_conv,
    "brutoByConv":  dict(bruto_by_conv),
    "unicosByConv": dict(Counter(r["conv"] for r in rows)),
}

# ── Injetar no HTML ──────────────────────────────────────────────────────────
print(f"💉 Injetando dados no {HTML_TMPL}...")
with open(HTML_TMPL, "r", encoding="utf-8") as f:
    content = f.read()

marker = "var pd = "
idx    = content.find(marker)
if idx == -1:
    print("❌ Marcador 'var pd = ' não encontrado no HTML template!")
    sys.exit(1)

start  = idx + len(marker)
depth  = 0; in_str = False; escape = False
for i, c in enumerate(content[start:start + 2_000_000]):
    if escape:    escape = False; continue
    if c == "\\" and in_str: escape = True; continue
    if c == '"' and not escape: in_str = not in_str
    if not in_str:
        if c == "{": depth += 1
        elif c == "}":
            depth -= 1
            if depth == 0: end_pos = i + 1; break

new_json    = json.dumps(pd_data, ensure_ascii=False, separators=(",", ":"))
new_content = content[:start] + new_json + content[start + end_pos:]

with open(HTML_OUT, "w", encoding="utf-8") as f:
    f.write(new_content)

print(f"\n✅ Dashboard gerado: {HTML_OUT}")
print(f"   Bruto: {TOTAL_BRUTO} | Dups: {TOTAL_DUPS} | Únicos: {TOTAL_UNICOS}")
print(f"   Classifs: {dict(Counter(r['classif'] for r in rows))}")
