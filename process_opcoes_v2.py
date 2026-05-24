#!/usr/bin/env python3
"""
Dashboard BOVA11 — Opções + GEX (v2 · Nível Tesouraria)
Lê o Excel exportado do opcoes.net.br, calcula GEX com normalização profissional,
Max Pain, Skew, Put/Call ratio — e puxa fechamento via yfinance.
Gera output/index.html para GitHub Pages.
"""

import sys, os, glob, json, math
import pandas as pd
import numpy as np
from datetime import datetime

GAMMA_SCALE = 10_000.0
DELTA_SCALE = 10_000.0

# ── 1. Localiza xlsx mais recente ─────────────────────────────────────────────
def find_latest_xlsx():
    files = glob.glob("data/*.xlsx")
    if not files:
        sys.exit("❌ Nenhum arquivo .xlsx encontrado em data/")
    return max(files, key=os.path.getmtime)

# ── 2. Spot price ─────────────────────────────────────────────────────────────
def get_spot_price(df_fallback):
    try:
        import yfinance as yf
        hist = yf.Ticker("BOVA11.SA").history(period="5d")
        if not hist.empty:
            price = round(float(hist["Close"].iloc[-1]), 2)
            date  = hist.index[-1].strftime("%d/%m/%Y")
            print(f"✅ yfinance: BOVA11 = R$ {price} ({date})")
            return price, date
    except Exception as e:
        print(f"⚠️  yfinance indisponível ({e})")

    dist = pd.to_numeric(df_fallback.get("Dist. (%) do Strike", pd.Series()), errors="coerce")
    spot = (df_fallback["Strike"] / (1 + dist / 10_000)).median()
    date = df_fallback["Data/Hora"].iloc[0] if "Data/Hora" in df_fallback.columns else "N/D"
    print(f"✅ Spot (fallback mediana): R$ {spot:.2f}")
    return round(float(spot), 2), str(date)

# ── 3. Carrega e normaliza o Excel ────────────────────────────────────────────
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

    # Garante coluna Tipo
    if "Tipo" not in df.columns:
        for col in df.columns:
            if "tipo" in col.lower() or "call" in col.lower():
                df["Tipo"] = df[col]
                break

    return df

# ── 4. GEX ────────────────────────────────────────────────────────────────────
def calc_gex(df, spot):
    df = df.copy()
    df["OI"]      = df["Tit."] + df["Lanç."]
    df["Gamma_d"] = df["Gamma"] / GAMMA_SCALE
    df["GEX"]     = df.apply(
        lambda r: (1 if r["Tipo"] == "CALL" else -1) * r["Gamma_d"] * r["OI"] * spot,
        axis=1
    )
    gex_s = (
        df.groupby("Strike")["GEX"].sum()
          .reset_index()
          .sort_values("Strike")
          .rename(columns={"GEX": "GEX_net"})
    )
    gex_s["GEX_cum"] = gex_s["GEX_net"].cumsum()

    flip_strike = None
    cum_arr = gex_s["GEX_cum"].values
    for i in range(1, len(cum_arr)):
        if cum_arr[i - 1] < 0 and cum_arr[i] >= 0:
            flip_strike = float(gex_s["Strike"].iloc[i])
            break

    return gex_s, flip_strike

# ── 5. Max Pain ───────────────────────────────────────────────────────────────
def calc_max_pain(df):
    strikes = sorted(df["Strike"].unique())
    calls = df[df["Tipo"] == "CALL"].copy()
    puts  = df[df["Tipo"] == "PUT"].copy()
    calls["OI"] = calls["Tit."] + calls["Lanç."]
    puts["OI"]  = puts["Tit."]  + puts["Lanç."]
    pain_vals = {}
    for s in strikes:
        c_pain = ((s - calls["Strike"]).clip(lower=0) * calls["OI"]).sum()
        p_pain = ((puts["Strike"] - s).clip(lower=0) * puts["OI"]).sum()
        pain_vals[s] = c_pain + p_pain
    return float(min(pain_vals, key=pain_vals.get))

# ── 6. KPIs ───────────────────────────────────────────────────────────────────
def calc_kpis(df):
    calls = df[df["Tipo"] == "CALL"]
    puts  = df[df["Tipo"] == "PUT"]

    def safe_sum(series):
        return float(series.replace([np.inf, -np.inf], np.nan).fillna(0).sum())

    vc  = safe_sum(calls["Vol. Financeiro"])
    vp  = safe_sum(puts["Vol. Financeiro"])
    nc  = safe_sum(calls["Núm. de Neg."])
    np_ = safe_sum(puts["Núm. de Neg."])
    iv_c = float(calls["Vol. Impl. (%)"].replace([np.inf, -np.inf], np.nan).mean() or 0)
    iv_p = float(puts["Vol. Impl. (%)"].replace([np.inf, -np.inf], np.nan).mean() or 0)

    return {
        "iv_call":  round(iv_c, 1),
        "iv_put":   round(iv_p, 1),
        "skew":     round(iv_p - iv_c, 1),
        "pc_ratio": round(vp / vc, 3) if vc > 0 else 0,
        "vol_call": round(vc, 0), "vol_put": round(vp, 0),
        "neg_call": int(nc), "neg_put": int(np_),
        "n_calls":  int(len(calls)), "n_puts": int(len(puts)),
    }

# ── 7. Gera HTML ──────────────────────────────────────────────────────────────
def gerar_html(spot, data_ref, kpis, max_pain, flip_strike, gex_s):
    flip_str   = f"R$ {flip_strike:,.0f}".replace(",", ".") if flip_strike else "N/D"
    pain_str   = f"R$ {max_pain:,.0f}".replace(",", ".")
    spot_str   = f"R$ {spot:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    skew_color = "#ef4444" if kpis["skew"] > 0 else "#22c55e"
    pc_color   = "#ef4444" if kpis["pc_ratio"] > 1 else "#22c55e"

    # Dados para o gráfico GEX
    strikes_list = gex_s["Strike"].tolist()
    gex_list     = gex_s["GEX_net"].round(2).tolist()
    colors_list  = ["'#22c55e'" if v >= 0 else "'#ef4444'" for v in gex_list]

    # Top 15 strikes por |GEX|
    top15 = gex_s.reindex(gex_s["GEX_net"].abs().nlargest(15).index).sort_values("Strike")
    t_strikes = top15["Strike"].tolist()
    t_gex     = top15["GEX_net"].round(2).tolist()
    t_colors  = ["'#22c55e'" if v >= 0 else "'#ef4444'" for v in t_gex]

    now = datetime.now().strftime("%d/%m/%Y %H:%M")

    html = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>BOVA11 Dashboard · GEX & Opções</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Space+Mono:wght@400;700&family=DM+Sans:wght@300;400;600;700&display=swap" rel="stylesheet">
<style>
  :root {{
    --bg:        #0a0e1a;
    --surface:   #111827;
    --border:    #1f2937;
    --text:      #e5e7eb;
    --muted:     #6b7280;
    --green:     #22c55e;
    --red:       #ef4444;
    --blue:      #3b82f6;
    --yellow:    #f59e0b;
    --accent:    #818cf8;
  }}
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{
    background: var(--bg);
    color: var(--text);
    font-family: 'DM Sans', sans-serif;
    min-height: 100vh;
    padding: 24px 20px 60px;
  }}
  /* ── Header ── */
  .header {{
    display: flex;
    align-items: center;
    justify-content: space-between;
    flex-wrap: wrap;
    gap: 12px;
    margin-bottom: 32px;
    padding-bottom: 20px;
    border-bottom: 1px solid var(--border);
  }}
  .header-left h1 {{
    font-family: 'Space Mono', monospace;
    font-size: clamp(1.4rem, 4vw, 2rem);
    font-weight: 700;
    letter-spacing: -0.02em;
    color: #fff;
  }}
  .header-left h1 span {{ color: var(--accent); }}
  .header-left p {{
    font-size: .8rem;
    color: var(--muted);
    margin-top: 4px;
    font-family: 'Space Mono', monospace;
  }}
  .spot-badge {{
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: 12px;
    padding: 12px 20px;
    text-align: right;
  }}
  .spot-badge .label {{ font-size: .7rem; color: var(--muted); text-transform: uppercase; letter-spacing: .08em; }}
  .spot-badge .value {{
    font-family: 'Space Mono', monospace;
    font-size: 1.5rem;
    font-weight: 700;
    color: var(--green);
  }}
  /* ── KPI grid ── */
  .kpi-grid {{
    display: grid;
    grid-template-columns: repeat(auto-fill, minmax(160px, 1fr));
    gap: 14px;
    margin-bottom: 28px;
  }}
  .kpi-card {{
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: 12px;
    padding: 16px;
    transition: border-color .2s;
  }}
  .kpi-card:hover {{ border-color: var(--accent); }}
  .kpi-card .kpi-label {{
    font-size: .68rem;
    text-transform: uppercase;
    letter-spacing: .1em;
    color: var(--muted);
    margin-bottom: 8px;
  }}
  .kpi-card .kpi-value {{
    font-family: 'Space Mono', monospace;
    font-size: 1.3rem;
    font-weight: 700;
  }}
  .kpi-card .kpi-sub {{
    font-size: .72rem;
    color: var(--muted);
    margin-top: 4px;
  }}
  /* ── Charts ── */
  .charts-grid {{
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 20px;
    margin-bottom: 28px;
  }}
  @media (max-width: 768px) {{ .charts-grid {{ grid-template-columns: 1fr; }} }}
  .chart-card {{
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: 14px;
    padding: 20px;
  }}
  .chart-card h3 {{
    font-size: .8rem;
    text-transform: uppercase;
    letter-spacing: .1em;
    color: var(--muted);
    margin-bottom: 16px;
  }}
  .chart-card h3 span {{
    font-family: 'Space Mono', monospace;
    font-size: .9rem;
    color: var(--text);
    text-transform: none;
    letter-spacing: 0;
    margin-left: 8px;
  }}
  /* ── Footer ── */
  .footer {{
    text-align: center;
    font-size: .72rem;
    color: var(--muted);
    margin-top: 40px;
    font-family: 'Space Mono', monospace;
  }}
  /* ── Levels bar ── */
  .levels-bar {{
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: 14px;
    padding: 20px;
    margin-bottom: 28px;
  }}
  .levels-bar h3 {{
    font-size: .8rem;
    text-transform: uppercase;
    letter-spacing: .1em;
    color: var(--muted);
    margin-bottom: 16px;
  }}
  .levels-row {{
    display: flex;
    flex-wrap: wrap;
    gap: 20px;
    align-items: center;
  }}
  .level-item {{
    display: flex;
    flex-direction: column;
    gap: 4px;
  }}
  .level-item .lbl {{ font-size: .68rem; color: var(--muted); text-transform: uppercase; letter-spacing: .08em; }}
  .level-item .val {{
    font-family: 'Space Mono', monospace;
    font-size: 1.05rem;
    font-weight: 700;
  }}
  .sep {{ width: 1px; height: 40px; background: var(--border); }}
</style>
</head>
<body>

<div class="header">
  <div class="header-left">
    <h1>BOVA11 <span>Dashboard</span></h1>
    <p>Opções · GEX · Max Pain · Skew &nbsp;|&nbsp; Ref: {data_ref} &nbsp;|&nbsp; Gerado: {now}</p>
  </div>
  <div class="spot-badge">
    <div class="label">BOVA11 Spot</div>
    <div class="value">{spot_str}</div>
  </div>
</div>

<!-- KPIs -->
<div class="kpi-grid">
  <div class="kpi-card">
    <div class="kpi-label">IV Call (média)</div>
    <div class="kpi-value" style="color:var(--green)">{kpis['iv_call']}%</div>
    <div class="kpi-sub">{kpis['n_calls']} contratos</div>
  </div>
  <div class="kpi-card">
    <div class="kpi-label">IV Put (média)</div>
    <div class="kpi-value" style="color:var(--red)">{kpis['iv_put']}%</div>
    <div class="kpi-sub">{kpis['n_puts']} contratos</div>
  </div>
  <div class="kpi-card">
    <div class="kpi-label">Skew IV (Put − Call)</div>
    <div class="kpi-value" style="color:{skew_color}">{kpis['skew']:+.1f}%</div>
    <div class="kpi-sub">{'Bearish' if kpis['skew'] > 0 else 'Bullish'} bias</div>
  </div>
  <div class="kpi-card">
    <div class="kpi-label">Put/Call Ratio</div>
    <div class="kpi-value" style="color:{pc_color}">{kpis['pc_ratio']:.3f}</div>
    <div class="kpi-sub">Vol. financeiro</div>
  </div>
  <div class="kpi-card">
    <div class="kpi-label">Vol. Calls</div>
    <div class="kpi-value" style="color:var(--green)">R$ {kpis['vol_call']/1e6:.1f}M</div>
    <div class="kpi-sub">{kpis['neg_call']:,} negócios</div>
  </div>
  <div class="kpi-card">
    <div class="kpi-label">Vol. Puts</div>
    <div class="kpi-value" style="color:var(--red)">R$ {kpis['vol_put']/1e6:.1f}M</div>
    <div class="kpi-sub">{kpis['neg_put']:,} negócios</div>
  </div>
</div>

<!-- Levels -->
<div class="levels-bar">
  <h3>Níveis-Chave</h3>
  <div class="levels-row">
    <div class="level-item">
      <span class="lbl">Spot</span>
      <span class="val" style="color:var(--blue)">{spot_str}</span>
    </div>
    <div class="sep"></div>
    <div class="level-item">
      <span class="lbl">Max Pain</span>
      <span class="val" style="color:var(--yellow)">{pain_str}</span>
    </div>
    <div class="sep"></div>
    <div class="level-item">
      <span class="lbl">GEX Flip</span>
      <span class="val" style="color:var(--accent)">{flip_str}</span>
    </div>
    <div class="sep"></div>
    <div class="level-item">
      <span class="lbl">Posição vs Max Pain</span>
      <span class="val" style="color:{'var(--green)' if spot >= max_pain else 'var(--red)'}">
        {'Acima ↑' if spot >= max_pain else 'Abaixo ↓'} ({abs(spot - max_pain):.1f} pts)
      </span>
    </div>
    {'<div class="sep"></div><div class="level-item"><span class="lbl">Posição vs GEX Flip</span><span class="val" style="color:' + ("'var(--green)'" if flip_strike and spot >= flip_strike else "'var(--red)'") + '">' + ('Acima ↑' if flip_strike and spot >= flip_strike else 'Abaixo ↓') + (' (' + f'{abs(spot - flip_strike):.1f} pts)' if flip_strike else '') + '</span></div>' if flip_strike else ''}
  </div>
</div>

<!-- Charts -->
<div class="charts-grid">
  <div class="chart-card">
    <h3>GEX por Strike <span>— todos os strikes</span></h3>
    <canvas id="gexAll" height="280"></canvas>
  </div>
  <div class="chart-card">
    <h3>GEX por Strike <span>— top 15 por magnitude</span></h3>
    <canvas id="gexTop" height="280"></canvas>
  </div>
</div>

<div class="footer">
  BOVA11 Dashboard · dados opcoes.net.br · GEX = Γ × OI × Spot · não é recomendação de investimento
</div>

<script>
const SPOT = {spot};
const MAX_PAIN = {max_pain};
const FLIP = {flip_strike if flip_strike else 'null'};

const optGex = {{
  responsive: true,
  maintainAspectRatio: false,
  plugins: {{
    legend: {{ display: false }},
    tooltip: {{
      callbacks: {{
        label: ctx => ' GEX: ' + ctx.raw.toLocaleString('pt-BR', {{maximumFractionDigits:0}})
      }}
    }}
  }},
  scales: {{
    x: {{
      ticks: {{ color:'#6b7280', font:{{size:10}}, maxRotation:45 }},
      grid:  {{ color:'#1f2937' }}
    }},
    y: {{
      ticks: {{ color:'#6b7280', font:{{size:10}} }},
      grid:  {{ color:'#1f2937' }}
    }}
  }}
}};

// Chart 1 — todos os strikes
new Chart(document.getElementById('gexAll'), {{
  type: 'bar',
  data: {{
    labels: {json.dumps(strikes_list)},
    datasets: [{{
      data: {json.dumps(gex_list)},
      backgroundColor: [{', '.join(colors_list)}],
      borderRadius: 2,
      borderSkipped: false,
    }}]
  }},
  options: optGex
}});

// Chart 2 — top 15
new Chart(document.getElementById('gexTop'), {{
  type: 'bar',
  data: {{
    labels: {json.dumps(t_strikes)},
    datasets: [{{
      data: {json.dumps(t_gex)},
      backgroundColor: [{', '.join(t_colors)}],
      borderRadius: 3,
      borderSkipped: false,
    }}]
  }},
  options: {{
    ...optGex,
    plugins: {{
      ...optGex.plugins,
      annotation: {{}}
    }}
  }}
}});
</script>
</body>
</html>"""
    return html

# ── 8. Main ───────────────────────────────────────────────────────────────────
def main():
    path = find_latest_xlsx()
    print(f"📂 Lendo: {path}")
    df = load_data(path)
    spot, data_ref = get_spot_price(df)
    kpis      = calc_kpis(df)
    max_pain  = calc_max_pain(df)
    gex_s, flip_strike = calc_gex(df, spot)

    print(f"📊 {len(df)} opções | CALLs: {kpis['n_calls']} | PUTs: {kpis['n_puts']}")
    print(f"📈 IV CALL: {kpis['iv_call']}% | IV PUT: {kpis['iv_put']}% | Skew: {kpis['skew']}%")
    print(f"⚖️  Put/Call: {kpis['pc_ratio']}")
    print(f"🎯 Max Pain: {max_pain:,}")
    flip_str = f"{flip_strike:,}" if flip_strike else "N/D"
    print(f"🔄 GEX Flip: {flip_str}")

    os.makedirs("output", exist_ok=True)
    html = gerar_html(spot, data_ref, kpis, max_pain, flip_strike, gex_s)
    out_path = os.path.join("output", "index.html")
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"✅ Dashboard gerado: {out_path}")

if __name__ == "__main__":
    main()
