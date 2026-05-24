#!/usr/bin/env python3
"""
Dashboard BOVA11 — Opções + GEX (v3 · validado com dados reais)
Gera output/index.html para GitHub Pages.

Validações feitas com xlsx do opcoes.net.br (Mai/2026):
  - Strikes em escala 100x (17300 = R$173,00) — detectado automaticamente
  - Vol. Impl. (%) vem como inteiro ×10 (169 = 16,9%) — dividido por 10
  - Gamma vem como inteiro ×10000 (169 = 0,0169) — dividido por 10000 ✅
  - OI = Tit. + Lanç. — ambas são posições em contratos, lados opostos ✅
  - GEX Flip = strike onde GEX cumulativo cruza de negativo→positivo
    (robusto: usa abs().idxmin() como fallback se não cruzar)
  - Fórmula GEX = Gamma_d × OI × Spot² (canônica) ✅
"""

import sys, os, glob, json
import pandas as pd
import numpy as np
from datetime import datetime

GAMMA_SCALE = 10_000.0

def find_latest_xlsx():
    files = glob.glob("data/*.xlsx")
    if not files:
        sys.exit("❌ Nenhum arquivo .xlsx encontrado em data/")
    return max(files, key=os.path.getmtime)

def load_data(path):
    df = pd.read_excel(path, header=1)
    df.columns = [str(c).strip().replace("\xa0", "") for c in df.columns]
    if "Vol. Financeiro" in df.columns:
        df["Vol. Financeiro"] = (
            df["Vol. Financeiro"].astype(str)
            .str.replace(".", "", regex=False)
            .str.replace(",", ".", regex=False)
            .str.replace("nan", "0")
            .astype(float)
        )
    numeric_cols = ["Strike", "Último", "Vol. Impl. (%)", "Delta", "Gamma",
                    "Theta ($)", "Theta (%)", "Vega", "Núm. de Neg.",
                    "Coberto", "Travado", "Descob.", "Tit.", "Lanç."]
    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    if "Tipo" not in df.columns:
        for col in df.columns:
            if "tipo" in col.lower():
                df["Tipo"] = df[col]
                break
    return df

def get_spot_and_scale(df):
    """
    Busca spot real do yfinance (ex: R$ 177,00).
    Detecta fator de escala comparando com mediana dos strikes do xlsx.
    """
    strike_median = float(df["Strike"].median())

    spot_real = None
    data_ref  = "N/D"
    try:
        import yfinance as yf
        hist = yf.Ticker("BOVA11.SA").history(period="5d")
        if not hist.empty:
            spot_real = round(float(hist["Close"].iloc[-1]), 2)
            data_ref  = hist.index[-1].strftime("%d/%m/%Y")
            print(f"✅ yfinance: BOVA11 = R$ {spot_real} ({data_ref})")
    except Exception as e:
        print(f"⚠️  yfinance indisponível ({e})")

    if spot_real is not None and spot_real > 0:
        raw_ratio = strike_median / spot_real
        if   raw_ratio < 3:    scale = 1
        elif raw_ratio < 30:   scale = 10
        elif raw_ratio < 300:  scale = 100
        else:                  scale = 1000
        print(f"📐 Strike mediana={strike_median:.0f} | Spot={spot_real} | Escala detectada: {scale}x")
    else:
        spot_real = round(strike_median, 2)
        scale = 1
        print(f"⚠️  Usando mediana dos strikes como spot: {spot_real}")

    spot_calc = round(spot_real * scale, 2)
    return spot_real, spot_calc, scale, data_ref

def calc_gex(df, spot_calc):
    """
    GEX canônico: Gamma_d × OI × Spot²

    OI = Tit. + Lanç.
      Tit.  = contratos em aberto do lado comprador
      Lanç. = contratos em aberto do lado vendedor (inclui posições descobertas)
      São lados opostos e de magnitudes diferentes — a soma representa
      o total de exposição a gamma no mercado.

    Gamma vem como inteiro ×10000 no xlsx (ex: 169 = 0,0169).
    GAMMA_SCALE = 10000 normaliza para o valor real.

    GEX Flip: strike onde o GEX cumulativo cruza de negativo→positivo.
    Fallback: strike onde |GEX_cum| é mínimo (abs().idxmin()).
    """
    df = df.copy()
    df["OI"]      = df["Tit."] + df["Lanç."]
    df["Gamma_d"] = df["Gamma"] / GAMMA_SCALE
    df["GEX"]     = df.apply(
        lambda r: (1 if r["Tipo"] == "CALL" else -1)
                  * r["Gamma_d"] * r["OI"] * (spot_calc ** 2),
        axis=1,
    )
    gex_s = (df.groupby("Strike")["GEX"].sum()
               .reset_index().sort_values("Strike")
               .rename(columns={"GEX": "GEX_net"}))
    gex_s["GEX_cum"] = gex_s["GEX_net"].cumsum()

    # Flip: cruzamento negativo → positivo
    flip_strike = None
    for i in range(1, len(gex_s)):
        if gex_s["GEX_cum"].iloc[i-1] < 0 and gex_s["GEX_cum"].iloc[i] >= 0:
            flip_strike = float(gex_s["Strike"].iloc[i])
            break

    # Fallback: ponto mais próximo de zero
    if flip_strike is None:
        flip_idx    = gex_s["GEX_cum"].abs().idxmin()
        flip_strike = float(gex_s.loc[flip_idx, "Strike"])
        print("⚠️  GEX_cum não cruza zero — usando ponto mais próximo como flip")

    return gex_s, flip_strike

def calc_max_pain(df):
    """
    Max Pain = strike que minimiza o valor total pago pelos vendedores.
    OI = Tit. + Lanç. (consistente com calc_gex).
    """
    strikes = sorted(df["Strike"].unique())
    calls = df[df["Tipo"] == "CALL"].copy()
    puts  = df[df["Tipo"] == "PUT"].copy()
    calls["OI"] = calls["Tit."] + calls["Lanç."]
    puts["OI"]  = puts["Tit."]  + puts["Lanç."]
    pain_vals = {}
    for s in strikes:
        pain_vals[s] = (
            ((s - calls["Strike"]).clip(lower=0) * calls["OI"]).sum() +
            ((puts["Strike"] - s).clip(lower=0) * puts["OI"]).sum()
        )
    return float(min(pain_vals, key=pain_vals.get))

def calc_kpis(df):
    calls = df[df["Tipo"] == "CALL"]
    puts  = df[df["Tipo"] == "PUT"]
    def ss(s): return float(s.replace([np.inf, -np.inf], np.nan).fillna(0).sum())
    vc  = ss(calls["Vol. Financeiro"]); vp = ss(puts["Vol. Financeiro"])
    nc  = ss(calls["Núm. de Neg."]);   np_ = ss(puts["Núm. de Neg."])
    # Vol. Impl. (%) vem como inteiro ×10 no xlsx (169 = 16,9%)
    iv_c = float(calls["Vol. Impl. (%)"].replace([np.inf,-np.inf],np.nan).mean() or 0) / 10
    iv_p = float(puts["Vol. Impl. (%)"].replace([np.inf,-np.inf],np.nan).mean() or 0) / 10
    return {
        "iv_call": round(iv_c,1), "iv_put": round(iv_p,1),
        "skew": round(iv_p - iv_c,1),
        "pc_ratio": round(vp/vc,3) if vc>0 else 0,
        "vol_call": round(vc,0), "vol_put": round(vp,0),
        "neg_call": int(nc), "neg_put": int(np_),
        "n_calls": int(len(calls)), "n_puts": int(len(puts)),
    }

def fmt_brl(v):
    return f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def fmt_strike_display(v, scale):
    return fmt_brl(v / scale)

def gerar_html(spot_real, spot_calc, scale, data_ref, kpis, max_pain_raw, flip_raw, gex_s):
    skew_color = "#ef4444" if kpis["skew"] > 0 else "#22c55e"
    pc_color   = "#ef4444" if kpis["pc_ratio"] > 1 else "#22c55e"

    spot_fmt  = fmt_brl(spot_real)
    pain_fmt  = fmt_strike_display(max_pain_raw, scale)
    flip_fmt  = fmt_strike_display(flip_raw, scale) if flip_raw else "N/D"

    pain_diff_pct = round((spot_calc - max_pain_raw) / max_pain_raw * 100, 2)
    pain_dir  = "Acima ↑" if spot_calc >= max_pain_raw else "Abaixo ↓"
    pain_col  = "#22c55e" if spot_calc >= max_pain_raw else "#ef4444"

    flip_level_html = ""
    if flip_raw:
        flip_diff_pct = round((spot_calc - flip_raw) / flip_raw * 100, 2)
        flip_dir = "Acima ↑" if spot_calc >= flip_raw else "Abaixo ↓"
        flip_col = "#22c55e" if spot_calc >= flip_raw else "#ef4444"
        flip_level_html = f"""
    <div class="sep"></div>
    <div class="li"><span class="lbl">GEX Flip</span><span class="val" style="color:var(--accent)">{flip_fmt}</span></div>
    <div class="sep"></div>
    <div class="li"><span class="lbl">Spot vs GEX Flip</span><span class="val" style="color:{flip_col}">{flip_dir} ({abs(flip_diff_pct):.2f}%)</span></div>"""

    top20    = gex_s.loc[gex_s["GEX_net"].abs().nlargest(20).index].sort_values("GEX_net")
    t_labels = [fmt_strike_display(v, scale) for v in top20["Strike"].tolist()]
    t_values = [round(v,2) for v in top20["GEX_net"].tolist()]
    t_colors = ["'#22c55e'" if v>=0 else "'#ef4444'" for v in t_values]

    near = gex_s[(gex_s["Strike"] >= spot_calc*0.92) & (gex_s["Strike"] <= spot_calc*1.08)]
    if near.empty:
        near = gex_s[(gex_s["Strike"] >= spot_calc*0.85) & (gex_s["Strike"] <= spot_calc*1.15)]
    near     = near.loc[near["GEX_net"].abs().nlargest(25).index].sort_values("GEX_net")
    n_labels = [fmt_strike_display(v, scale) for v in near["Strike"].tolist()]
    n_values = [round(v,2) for v in near["GEX_net"].tolist()]
    n_colors = ["'#22c55e'" if v>=0 else "'#ef4444'" for v in n_values]

    h1  = max(320, len(t_labels)*22)
    h2  = max(320, len(n_labels)*22)
    now = datetime.now().strftime("%d/%m/%Y %H:%M")

    html = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>BOVA11 Dashboard</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<link href="https://fonts.googleapis.com/css2?family=Space+Mono:wght@400;700&family=DM+Sans:wght@400;600;700&display=swap" rel="stylesheet">
<style>
:root{{--bg:#0a0e1a;--surface:#111827;--border:#1f2937;--text:#e5e7eb;--muted:#6b7280;--green:#22c55e;--red:#ef4444;--blue:#3b82f6;--yellow:#f59e0b;--accent:#818cf8;}}
*{{box-sizing:border-box;margin:0;padding:0;}}
body{{background:var(--bg);color:var(--text);font-family:'DM Sans',sans-serif;padding:20px 18px 60px;}}
.header{{display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:10px;margin-bottom:22px;padding-bottom:16px;border-bottom:1px solid var(--border);}}
.header h1{{font-family:'Space Mono',monospace;font-size:clamp(1.2rem,3.5vw,1.8rem);color:#fff;}}
.header h1 span{{color:var(--accent);}}
.header p{{font-size:.73rem;color:var(--muted);margin-top:4px;font-family:'Space Mono',monospace;}}
.spot-badge{{background:var(--surface);border:1px solid var(--border);border-radius:10px;padding:10px 18px;text-align:right;}}
.spot-badge .lbl{{font-size:.63rem;color:var(--muted);text-transform:uppercase;letter-spacing:.1em;}}
.spot-badge .val{{font-family:'Space Mono',monospace;font-size:1.4rem;font-weight:700;color:var(--green);}}
.kpi-grid{{display:grid;grid-template-columns:repeat(auto-fill,minmax(138px,1fr));gap:10px;margin-bottom:16px;}}
.kpi{{background:var(--surface);border:1px solid var(--border);border-radius:10px;padding:12px 14px;}}
.k-lbl{{font-size:.62rem;text-transform:uppercase;letter-spacing:.1em;color:var(--muted);margin-bottom:6px;}}
.k-val{{font-family:'Space Mono',monospace;font-size:1.2rem;font-weight:700;}}
.k-sub{{font-size:.67rem;color:var(--muted);margin-top:3px;}}
.levels{{background:var(--surface);border:1px solid var(--border);border-radius:10px;padding:14px 18px;margin-bottom:16px;}}
.levels h3{{font-size:.66rem;text-transform:uppercase;letter-spacing:.1em;color:var(--muted);margin-bottom:10px;}}
.levels-row{{display:flex;flex-wrap:wrap;gap:14px;align-items:center;}}
.li{{display:flex;flex-direction:column;gap:3px;}}
.lbl{{font-size:.62rem;color:var(--muted);text-transform:uppercase;letter-spacing:.08em;}}
.val{{font-family:'Space Mono',monospace;font-size:.93rem;font-weight:700;}}
.sep{{width:1px;height:34px;background:var(--border);}}
.charts-grid{{display:grid;grid-template-columns:1fr 1fr;gap:14px;}}
@media(max-width:620px){{.charts-grid{{grid-template-columns:1fr;}}}}
.chart-card{{background:var(--surface);border:1px solid var(--border);border-radius:10px;padding:16px;}}
.chart-card h3{{font-size:.66rem;text-transform:uppercase;letter-spacing:.1em;color:var(--muted);margin-bottom:10px;}}
.chart-card h3 span{{font-family:'Space Mono',monospace;font-size:.8rem;color:var(--text);text-transform:none;letter-spacing:0;margin-left:5px;}}
.chart-wrap{{position:relative;}}
.footer{{text-align:center;font-size:.66rem;color:var(--muted);margin-top:26px;font-family:'Space Mono',monospace;}}
</style>
</head>
<body>
<div class="header">
  <div>
    <h1>BOVA11 <span>Dashboard</span></h1>
    <p>GEX · Max Pain · Skew · Put/Call &nbsp;|&nbsp; {data_ref} &nbsp;|&nbsp; {now}</p>
  </div>
  <div class="spot-badge">
    <div class="lbl">BOVA11 Fechamento</div>
    <div class="val">{spot_fmt}</div>
  </div>
</div>

<div class="kpi-grid">
  <div class="kpi"><div class="k-lbl">IV Call</div><div class="k-val" style="color:var(--green)">{kpis['iv_call']}%</div><div class="k-sub">{kpis['n_calls']} contratos</div></div>
  <div class="kpi"><div class="k-lbl">IV Put</div><div class="k-val" style="color:var(--red)">{kpis['iv_put']}%</div><div class="k-sub">{kpis['n_puts']} contratos</div></div>
  <div class="kpi"><div class="k-lbl">Skew Put−Call</div><div class="k-val" style="color:{skew_color}">{kpis['skew']:+.1f}%</div><div class="k-sub">{'Bearish' if kpis['skew']>0 else 'Bullish'} bias</div></div>
  <div class="kpi"><div class="k-lbl">Put/Call Ratio</div><div class="k-val" style="color:{pc_color}">{kpis['pc_ratio']:.3f}</div><div class="k-sub">vol. financeiro</div></div>
  <div class="kpi"><div class="k-lbl">Vol. Calls</div><div class="k-val" style="color:var(--green)">R$ {kpis['vol_call']/1e6:.1f}M</div><div class="k-sub">{kpis['neg_call']:,} negócios</div></div>
  <div class="kpi"><div class="k-lbl">Vol. Puts</div><div class="k-val" style="color:var(--red)">R$ {kpis['vol_put']/1e6:.1f}M</div><div class="k-sub">{kpis['neg_put']:,} negócios</div></div>
</div>

<div class="levels">
  <h3>Níveis-Chave</h3>
  <div class="levels-row">
    <div class="li"><span class="lbl">Fechamento</span><span class="val" style="color:var(--blue)">{spot_fmt}</span></div>
    <div class="sep"></div>
    <div class="li"><span class="lbl">Max Pain</span><span class="val" style="color:var(--yellow)">{pain_fmt}</span></div>
    <div class="sep"></div>
    <div class="li"><span class="lbl">Spot vs Max Pain</span><span class="val" style="color:{pain_col}">{pain_dir} ({abs(pain_diff_pct):.2f}%)</span></div>
    {flip_level_html}
  </div>
</div>

<div class="charts-grid">
  <div class="chart-card">
    <h3>GEX<span>Top 20 por magnitude</span></h3>
    <div class="chart-wrap" style="height:{h1}px"><canvas id="cTop"></canvas></div>
  </div>
  <div class="chart-card">
    <h3>GEX<span>Zona ±8% do Spot</span></h3>
    <div class="chart-wrap" style="height:{h2}px"><canvas id="cNear"></canvas></div>
  </div>
</div>

<div class="footer">BOVA11 Dashboard · dados opcoes.net.br · GEX = Γ × OI × Spot² · não é recomendação de investimento</div>

<script>
const opts = {{
  indexAxis:'y', responsive:true, maintainAspectRatio:false,
  plugins:{{ legend:{{display:false}}, tooltip:{{ callbacks:{{ label: c=>' GEX: '+c.raw.toLocaleString('pt-BR',{{maximumFractionDigits:0}}) }} }} }},
  scales:{{
    x:{{ ticks:{{color:'#6b7280',font:{{size:10}}}}, grid:{{color:'#1f2937'}}, border:{{color:'#374151'}} }},
    y:{{ ticks:{{color:'#9ca3af',font:{{size:10,family:"'Space Mono',monospace"}}}}, grid:{{display:false}}, border:{{color:'#374151'}} }}
  }}
}};
new Chart(document.getElementById('cTop'),{{
  type:'bar',
  data:{{labels:{json.dumps(t_labels)},datasets:[{{data:{json.dumps(t_values)},backgroundColor:[{','.join(t_colors)}],borderRadius:4,borderSkipped:false,barThickness:16}}]}},
  options:JSON.parse(JSON.stringify(opts))
}});
new Chart(document.getElementById('cNear'),{{
  type:'bar',
  data:{{labels:{json.dumps(n_labels)},datasets:[{{data:{json.dumps(n_values)},backgroundColor:[{','.join(n_colors)}],borderRadius:4,borderSkipped:false,barThickness:16}}]}},
  options:JSON.parse(JSON.stringify(opts))
}});
</script>
</body>
</html>"""
    return html

def main():
    path = find_latest_xlsx()
    print(f"📂 Lendo: {path}")
    df = load_data(path)
    spot_real, spot_calc, scale, data_ref = get_spot_and_scale(df)
    kpis     = calc_kpis(df)
    max_pain = calc_max_pain(df)
    gex_s, flip_strike = calc_gex(df, spot_calc)

    print(f"📊 {len(df)} opções | CALLs:{kpis['n_calls']} PUTs:{kpis['n_puts']}")
    print(f"📈 IV:{kpis['iv_call']}%/{kpis['iv_put']}% Skew:{kpis['skew']}% P/C:{kpis['pc_ratio']}")
    print(f"🎯 MaxPain: {max_pain:,.0f} (R$ {max_pain/scale:,.2f})")
    print(f"🔄 GEXFlip: {flip_strike:,.0f} (R$ {flip_strike/scale:,.2f})")

    os.makedirs("output", exist_ok=True)
    html = gerar_html(spot_real, spot_calc, scale, data_ref, kpis, max_pain, flip_strike, gex_s)
    with open("output/index.html", "w", encoding="utf-8") as f:
        f.write(html)
    print("✅ output/index.html gerado!")

if __name__ == "__main__":
    main()
