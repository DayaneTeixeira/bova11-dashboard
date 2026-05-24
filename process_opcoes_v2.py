#!/usr/bin/env python3
"""
Dashboard BOVA11 — Opções + GEX (v2 · Nível Tesouraria)
Lê o Excel exportado do opcoes.net.br, calcula GEX com normalização profissional,
Max Pain, Skew, Put/Call ratio — e puxa fechamento via yfinance.

Fórmulas:
  GEX = Γ_decimal × OI × Spot
  onde: Γ_decimal = Gamma_raw / 10.000  (dado do opcoes.net.br em bps)
        OI = Tit. + Lanç. (titulares + lançadores)
  GEX Flip = primeiro cruzamento do GEX cumulativo em zero (baixo → alto)
  Max Pain  = argmin_S Σ_K OI_K × max(S−K, 0)  [calls]  +  Σ_K OI_K × max(K−S, 0)  [puts]
"""

import sys, os, glob, json, math
import pandas as pd
import numpy as np
from datetime import datetime

GAMMA_SCALE = 10_000.0   # opcoes.net.br entrega Gamma × 10.000
DELTA_SCALE = 10_000.0   # idem para Delta

# ── 1. Localiza xlsx mais recente ────────────────────────────────────────────
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

    # Fallback via Dist%: Spot = Strike / (1 + Dist/10000)
    dist = pd.to_numeric(df_fallback["Dist. (%) do Strike"], errors="coerce")
    spot = (df_fallback["Strike"] / (1 + dist / 10_000)).median()
    date = df_fallback["Data/Hora"].iloc[0] if "Data/Hora" in df_fallback.columns else "N/D"
    print(f"✅ Spot (fallback mediana): R$ {spot:.2f}")
    return round(spot, 2), date

# ── 3. Carrega e normaliza o Excel ───────────────────────────────────────────
def load_data(path):
    df = pd.read_excel(path, header=1)
    df.columns = [str(c).strip().replace("\xa0", "") for c in df.columns]

    # Volume financeiro (formato pt-BR)
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

    return df

# ── 4. GEX profissional ───────────────────────────────────────────────────────
def calc_gex(df, spot):
    """
    GEX = Γ_dec × OI × Spot
    Γ_dec = Gamma_raw / GAMMA_SCALE   (bps → decimal)
    OI    = Tit. + Lanç.
    CALL → +GEX (dealer long gamma → compra alta, vende queda = estabiliza)
    PUT  → -GEX (dealer short gamma → compra queda, vende alta = amplifica)
    """
    df = df.copy()
    df["OI"]      = df["Tit."] + df["Lanç."]
    df["Gamma_d"] = df["Gamma"] / GAMMA_SCALE
    df["Delta_d"] = df["Delta"] / DELTA_SCALE
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

    # GEX Flip: primeira transição negativo → positivo
    flip_strike = None
    cum_arr = gex_s["GEX_cum"].values
    for i in range(1, len(cum_arr)):
        if cum_arr[i - 1] < 0 and cum_arr[i] >= 0:
            flip_strike = int(gex_s["Strike"].iloc[i])
            break

    # Top 10 por |GEX_net|
    top10 = gex_s.nlargest(10, "GEX_net", keep="all").copy()
    top10_abs = gex_s.iloc[gex_s["GEX_net"].abs().nlargest(10).index]

    return gex_s, flip_strike, top10_abs

# ── 5. Max Pain ───────────────────────────────────────────────────────────────
def calc_max_pain(df):
    """
    Max Pain = strike S que minimiza o valor total de exercício a pagar pelos lançadores.
    Para cada S candidato:
      pain = Σ_K_call OI_call(K) × max(S−K, 0)  +  Σ_K_put OI_put(K) × max(K−S, 0)
    """
    strikes = sorted(df["Strike"].unique())
    calls = df[df["Tipo"] == "CALL"].copy()
    puts  = df[df["Tipo"] == "PUT"].copy()
    calls["OI"] = calls["Tit."] + calls["Lanç."]
    puts["OI"]  = puts["Tit."] + puts["Lanç."]

    pain_vals = {}
    for s in strikes:
        c_pain = ((s - calls["Strike"]).clip(lower=0) * calls["OI"]).sum()
        p_pain = ((puts["Strike"] - s).clip(lower=0) * puts["OI"]).sum()
        pain_vals[s] = c_pain + p_pain

    return min(pain_vals, key=pain_vals.get)

# ── 6. KPIs gerais ────────────────────────────────────────────────────────────
def calc_kpis(df):
    calls = df[df["Tipo"] == "CALL"]
    puts  = df[df["Tipo"] == "PUT"]

    def safe_sum(series):
        return float(series.replace([np.inf, -np.inf], np.nan).fillna(0).sum())

    vc = safe_sum(calls["Vol. Financeiro"])
    vp = safe_sum(puts["Vol. Financeiro"])
    nc = safe_sum(calls["Núm. de Neg."])
    np_ = safe_sum(puts["Núm. de Neg."])
    iv_c = float(calls["Vol. Impl. (%)"].mean())
    iv_p = float(puts["Vol. Impl. (%)"].mean())

    return {
        "iv_call":  round(iv_c, 1),
        "iv_put":   round(iv_p, 1),
        "skew":     round(iv_p - iv_c, 1),
        "pc_ratio": round(vp / vc, 3) if vc > 0 else 0,
        "vol_call": vc, "vol_put": vp,
        "neg_call": int(nc), "neg_put": int(np_),
        "n_calls": len(calls), "n_puts": len(puts),
    }

# ── 7. Main ───────────────────────────────────────────────────────────────────
def main():
    path = find_latest_xlsx()
    print(f"📂 Lendo: {path}")
    df = load_data(path)
    spot, data_ref = get_spot_price(df)
    kpis = calc_kpis(df)
    max_pain = calc_max_pain(df)
    gex_s, flip_strike, top10 = calc_gex(df, spot)

    print(f"📊 {len(df)} opções | CALLs: {kpis['n_calls']} | PUTs: {kpis['n_puts']}")
    print(f"📈 IV CALL: {kpis['iv_call']}% | IV PUT: {kpis['iv_put']}% | Skew: {kpis['skew']}%")
    print(f"⚖️  Put/Call: {kpis['pc_ratio']}")
    print(f"🎯 Max Pain: {max_pain:,}")
    flip_str = f"{flip_strike:,}" if flip_strike else "N/D"
    print(f"🔄 GEX Flip: {flip_str}")
    print(f"📌 Top 10 GEX: {top10['Strike'].sort_values().tolist()}")

    # (geração HTML omitida — usar template separado)
    print("✅ Processo concluído. Integrar com template HTML v2.")

if __name__ == "__main__":
    main()
