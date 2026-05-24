#!/usr/bin/env python3
"""
Dashboard BOVA11 — Opções + GEX (v2 · Nível Tesouraria)
Gera output/index.html para GitHub Pages.
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

def get_spot_price(df):
    """
    Tenta yfinance. Se o preço retornado estiver em escala muito diferente
    dos strikes do xlsx, usa a mediana dos strikes como referência de spot.
    """
    strike_median = float(df["Strike"].median())

    yf_price = None
    yf_date  = None
    try:
        import yfinance as yf
        hist = yf.Ticker("BOVA11.SA").history(period="5d")
        if not hist.empty:
            yf_price = round(float(hist["Close"].iloc[-1]), 2)
            yf_date  = hist.index[-1].strftime("%d/%m/%Y")
            print(f"✅ yfinance: BOVA11 = R$ {yf_price} ({yf_date})")
    except Exception as e:
        print(f"⚠️  yfinance indisponível ({e})")

    if yf_price is not None:
        # Se o preço do yfinance for compatível com os strikes (razão < 10x), usa direto
        ratio = strike_median / yf_price
        if 0.1 <= ratio <= 10:
            return yf_price, yf_date
        else:
            # Escalas incompatíveis: usa mediana dos strikes como spot
            print(f"⚠️  Escala incompatível (yfinance={yf_price}, strike mediana={strike_median:.0f}). Usando mediana dos strikes.")
            return round(strike_median, 2), yf_date or "N/D"

    # Fallback via Dist%
    dist = pd.to_numeric(df.get("Dist. (%) do Strike", pd.Series()), errors="coerce")
    if dist.notna().any():
        spot = (df["Strike"] / (1 + dist / 10_000)).median()
    else:
        spot = strike_median
    date = df["Data/Hora"].iloc[0] if "Data/Hora" in df.columns else "N/D"
    print(f"✅ Spot (fallback): R$ {spot:.2f}")
    return round(float(spot), 2), str(date)

def calc_gex(df, spot):
    df = df.copy()
    df["OI"]      = df["Tit."] + df["Lanç."]
    df["Gamma_d"] = df["Gamma"] / GAMMA_SCALE
    df["GEX"]     = df.apply(
        lambda r: (1 if r["Tipo"] == "CALL" else -1) * r["Gamma_d"] * r["OI"] * spot, axis=1)
    gex_s = (df.groupby("Strike")["GEX"].sum()
               .reset_index().sort_values("Strike")
               .rename(columns={"GEX": "GEX_net"}))
    gex_s["GEX_cum"] = gex_s["GEX_net"].cumsum()
    flip_strike = None
    for i in range(1, len(gex_s)):
        if gex_s["GEX_cum"].iloc[i-1] < 0 and gex_s["GEX_cum"].iloc[i] >= 0:
            flip_strike = float(gex_s["Strike"].iloc[i])
            break
    return gex_s, flip_strike

def calc_max_pain(df):
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
    iv_c = float(calls["Vol. Impl. (%)"].replace([np.inf,-np.inf],np.nan).mean() or 0)
    iv_p = float(puts["Vol. Impl. (%)"].replace([np.inf,-np.inf],np.nan).mean() or 0)
    return {
        "iv_call": round(iv_c,1), "iv_put": round(iv_p,1),
        "skew": round(iv_p - iv_c,1),
        "pc_ratio": round(vp/vc,3) if vc>0 else 0,
        "vol_call": round(vc,0), "vol_put": round(vp,0),
        "neg_call": int(nc), "neg_put": int(np_),
        "n_calls": int(len(calls)), "n_puts": int(len(puts)),
    }

def fmt_strike(v):
    """Formata strike no mesmo padrão dos dados (sem prefixo R$ para evitar confusão de escala)."""
    return f"{v:,.0f}".replace(",", ".")

def gerar_html(spot, data_ref, kpis, max_pain, flip_strike, gex_s):
    skew_color = "#ef4444" if kpis["skew"] > 0 else "#22c55e"
    pc_color   = "#ef4444" if kpis["pc_ratio"] > 1 else "#22c55e"

    spot_fmt     = fmt_strike(spot)
    pain_fmt     = fmt_strike(max_pain)
    flip_fmt     = fmt_strike(flip_strike) if flip_strike else "N/D"

    # Diferença percentual (evita problema de escala)
    pain_diff_pct = round((spot - max_pain) / max_pain * 100, 2)
    pain_dir  = "Acima ↑" if spot >= max_pain else "Abaixo ↓"
    pain_col  = "#22c55e" if spot >= max_pain else "#ef4444"

    flip_level_html = ""
    if flip_strike:
        flip_diff_pct = round((spot - flip_strike) / flip_strike * 100, 2)
        flip_dir  = "Acima ↑" if spot >= flip_strike else "Abaixo ↓"
        flip_col  = "#22c55e" if spot >= flip_strike else "#ef4444"
        flip_level_html = f"""
    <div class="sep"></div>
    <div class="li"><span class="lbl">GEX Flip</span><span class="val" style="color:var(--accent)">{flip_fmt}</span></div>
    <div class="sep"></div>
    <div class="li"><span class="lbl">Spot vs GEX Flip</span><span class="val" style="color:{flip_col}">{flip_dir} ({abs(flip_diff_pct):.2f}%)</span></div>"""

    # ── Gráfico 1: Top 20 por |GEX| ────────────────────────────────────────
    top20 = gex_s.loc[gex_s["GEX_net"].abs().nlargest(20).index].sort_values("GEX_net")
    t_labels = [fmt_strike(v) for v in top20["Strike"].tolist()]
    t_values = [round(v,2) for v in top20["GEX_net"].tolist()]
    t_colors = ["'#22c55e'" if v>=0 else "'#ef4444'" for v in t_values]

    # ── Gráfico 2: ±8% da MEDIANA dos strikes (zona relevante) ─────────────
    ref = float(gex_s["Strike"].median())
    near = gex_s[(gex_s["Strike"] >= ref*0.92) & (gex_s["Strike"] <= ref*1.08)]
    # se ficar vazio, abre para ±15%
    if near.empty:
        near = gex_s[(gex_s["Strike"] >= ref*0.85) & (gex_s["Strike"] <= ref*1.15)]
    near = near.loc[near["GEX_net"].abs().nlargest(25).index].sort_values("GEX_net")
    n_labels = [fmt_strike(v) for v in near["Strike"].tolist()]
    n_values = [round(v,2) for v in near["GEX_net"].tolist()]
    n_colors = ["'#22c55e'" if v>=0 else "'#ef4444'" for v in n_values]

    h1 = max(320, len(t_labels)*22)
    h2 = max(320, len(n_labels)*22)
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
    <div class="lbl">Spot Ref.</div>
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
    <div class="li"><span class="lbl">Spot Ref.</span><span class="val" style="color:var(--blue)">{spot_fmt}</span></div>
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
    <h3>GEX<span>Zona central (±8% da mediana)</span></h3>
    <div class="chart-wrap" style="height:{h2}px"><canvas id="cNear"></canvas></div>
  </div>
</div>

<div class="footer">BOVA11 Dashboard · dados opcoes.net.br · GEX = Γ × OI × Spot · não é recomendação de investimento</div>

<script>
const opts = {{
  indexAxis:'y', responsive:true, maintainAspectRatio:false,
  plugins:{{ legend:{{display:false}}, tooltip:{{ callbacks:{{ label: c=>' '+c.raw.toLocaleString('pt-BR',{{maximumFractionDigits:0}}) }} }} }},
  scales:{{
    x:{{ ticks:{{color:'#6b7280',font:{{size:10}}}}, grid:{{color:'#1f2937'}}, border:{{color:'#374151'}} }},
    y:{{ ticks:{{color:'#9ca3af',font:{{size:11,family:"'Space Mono',monospace"}}}}, grid:{{display:false}}, border:{{color:'#374151'}} }}
  }}
}};
new Chart(document.getElementById('cTop'),{{
  type:'bar',
  data:{{
    labels:{json.dumps(t_labels)},
    datasets:[{{data:{json.dumps(t_values)},backgroundColor:[{','.join(t_colors)}],borderRadius:4,borderSkipped:false,barThickness:16}}]
  }},
  options:JSON.parse(JSON.stringify(opts))
}});
new Chart(document.getElementById('cNear'),{{
  type:'bar',
  data:{{
    labels:{json.dumps(n_labels)},
    datasets:[{{data:{json.dumps(n_values)},backgroundColor:[{','.join(n_colors)}],borderRadius:4,borderSkipped:false,barThickness:16}}]
  }},
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
    spot, data_ref = get_spot_price(df)
    kpis      = calc_kpis(df)
    max_pain  = calc_max_pain(df)
    gex_s, flip_strike = calc_gex(df, spot)
    print(f"📊 {len(df)} opções | CALLs:{kpis['n_calls']} PUTs:{kpis['n_puts']}")
    print(f"📈 IV CALL:{kpis['iv_call']}% IV PUT:{kpis['iv_put']}% Skew:{kpis['skew']}%")
    print(f"⚖️  P/C:{kpis['pc_ratio']} | 🎯 MaxPain:{max_pain:,.0f} | 🔄 Flip:{flip_strike or 'N/D'}")
    os.makedirs("output", exist_ok=True)
    html = gerar_html(spot, data_ref, kpis, max_pain, flip_strike, gex_s)
    with open("output/index.html", "w", encoding="utf-8") as f:
        f.write(html)
    print("✅ output/index.html gerado!")

if __name__ == "__main__":
    main()
