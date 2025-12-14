#!/usr/bin/env python3
"""
Ski Resort Decision Engine: Corviglia vs Corvatsch
Sends daily email at 7am CET with today's pick + 5-day forecast.
"""
from __future__ import annotations

import datetime as dt
import os
import smtplib
import statistics
import sys
from dataclasses import dataclass
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from typing import Any, Optional
from zoneinfo import ZoneInfo

import requests

# =============================================================================
# CONFIGURATION
# =============================================================================
SMTP_HOST = "smtp.gmail.com"
SMTP_PORT = 587
# (Optional) load local .env for convenience (kept out of git by your .gitignore)
try:
    from dotenv import load_dotenv
    load_dotenv()
except Exception:
    pass

SENDER_EMAIL = os.environ.get("SMTP_SENDER_EMAIL")
SENDER_PASSWORD = os.environ.get("SMTP_PASSWORD")
RECIPIENT_EMAIL = os.environ.get("SMTP_RECIPIENT_EMAIL")


TZ = ZoneInfo("Europe/Zurich")
FORECAST_DAYS = 5  # Today + next 4 days
START_HOUR = 9
END_HOUR = 16  # Extended slightly for afternoon skiing


# =============================================================================
# RESORT DEFINITIONS
# =============================================================================
# =============================================================================
# RESORT DEFINITIONS
# =============================================================================
@dataclass(frozen=True)
class Resort:
    name: str
    short: str
    emoji: str
    region: str  # "Engadin" or "Davos"
    lat: float
    lon: float
    elevation_m: int
    aspect: str  # "south" or "north" (approx; tweak if picks feel off)
    wind_exposure: float  # 1.0 = normal, >1 = more exposed


# --- Engadin / St. Moritz area ---
CORVIGLIA = Resort(
    name="Corviglia (St. Moritz)",
    short="Corviglia",
    emoji="⛷️",
    region="Engadin",
    lat=46.5079,
    lon=9.8192,
    elevation_m=2486,
    aspect="south",
    wind_exposure=0.85,
)

CORVATSCH = Resort(
    name="Corvatsch 3303",
    short="Corvatsch",
    emoji="🏔️",
    region="Engadin",
    lat=46.4179,
    lon=9.8212,
    elevation_m=3303,
    aspect="north",
    wind_exposure=1.25,
)

DIAVOLEZZA = Resort(
    name="Diavolezza",
    short="Diavolezza",
    emoji="🧊",
    region="Engadin",
    lat=46.4073,
    lon=9.9593,
    elevation_m=2978,
    aspect="north",
    wind_exposure=1.30,
)

ZUOZ = Resort(
    name="Zuoz (Pizzet/Albanas)",
    short="Zuoz",
    emoji="🌞",
    region="Engadin",
    lat=46.6029,
    lon=9.9600,
    elevation_m=2465,
    aspect="south",
    wind_exposure=0.90,
)

# --- Davos / Klosters areas ---
PARSENN = Resort(
    name="Parsenn (Davos/Klosters)",
    short="Parsenn",
    emoji="🚠",
    region="Davos",
    lat=46.8400,
    lon=9.8100,
    elevation_m=2817,
    aspect="north",
    wind_exposure=1.10,
)

JAKOBSHORN = Resort(
    name="Jakobshorn (Davos Platz)",
    short="Jakobshorn",
    emoji="🏂",
    region="Davos",
    lat=46.7724,
    lon=9.8494,
    elevation_m=2590,
    aspect="south",
    wind_exposure=1.05,
)

RINERHORN = Resort(
    name="Rinerhorn",
    short="Rinerhorn",
    emoji="👨‍👩‍👧‍👦",
    region="Davos",
    lat=46.7394,
    lon=9.8141,
    elevation_m=2528,
    aspect="south",
    wind_exposure=1.00,
)

PISCHA = Resort(
    name="Pischa",
    short="Pischa",
    emoji="🌄",
    region="Davos",
    lat=46.8096,
    lon=9.9192,
    elevation_m=2483,
    aspect="south",
    wind_exposure=1.15,
)

MADRISA = Resort(
    name="Madrisa (Klosters)",
    short="Madrisa",
    emoji="🌲",
    region="Davos",
    lat=46.9253,
    lon=9.8699,
    elevation_m=2600,
    aspect="north",
    wind_exposure=1.05,
)

SCHATZALP = Resort(
    name="Schatzalp (Strela)",
    short="Schatzalp",
    emoji="🐢",
    region="Davos",
    lat=46.7971,
    lon=9.8215,
    elevation_m=1861,
    aspect="south",
    wind_exposure=0.80,
)

ALL_RESORTS: list[Resort] = [
    CORVIGLIA, CORVATSCH, DIAVOLEZZA, ZUOZ,
    PARSENN, JAKOBSHORN, RINERHORN, PISCHA, MADRISA, SCHATZALP,
]
RESORT_BY_SHORT = {r.short: r for r in ALL_RESORTS}



# =============================================================================
# OPEN-METEO API
# =============================================================================
API_URL = "https://api.open-meteo.com/v1/forecast"

HOURLY_VARS = [
    "temperature_2m",
    "apparent_temperature",
    "precipitation",
    "snowfall",
    "snow_depth",
    "cloud_cover",
    "cloud_cover_low",
    "cloud_cover_mid",
    "visibility",
    "wind_speed_10m",
    "wind_gusts_10m",
    "freezing_level_height",
    "sunshine_duration",
    "weather_code",
]


def fetch_forecast(resort: Resort, forecast_days: int = 7) -> dict[str, Any]:
    """Fetch multi-day forecast from Open-Meteo."""
    params = {
        "latitude": resort.lat,
        "longitude": resort.lon,
        "hourly": ",".join(HOURLY_VARS),
        "timezone": "Europe/Zurich",
        "forecast_days": forecast_days,
        "wind_speed_unit": "kmh",
        "precipitation_unit": "mm",
    }
    r = requests.get(API_URL, params=params, timeout=30)
    r.raise_for_status()
    return r.json()


# =============================================================================
# SCORING ENGINE (improved)
# =============================================================================
def clamp(x: float, lo: float, hi: float) -> float:
    return max(lo, min(hi, x))


def get_val(row: dict, key: str, default: float = 0.0) -> float:
    """Safely get numeric value from row."""
    v = row.get(key)
    return float(v) if v is not None else default


def hour_score(resort: Resort, row: dict[str, float]) -> tuple[float, list[str]]:
    """
    Score a single hour. Returns (score, list_of_concerns).
    Score 0-100: higher = better skiing conditions.
    """
    score = 100.0
    concerns: list[str] = []

    temp = get_val(row, "temperature_2m")
    apparent = get_val(row, "apparent_temperature", temp)
    precip = get_val(row, "precipitation")
    snowfall = get_val(row, "snowfall")
    cloud = get_val(row, "cloud_cover")
    cloud_low = get_val(row, "cloud_cover_low")
    vis = get_val(row, "visibility", 20000)
    wind = get_val(row, "wind_speed_10m")
    gust = get_val(row, "wind_gusts_10m")
    frz = get_val(row, "freezing_level_height")
    weather_code = int(get_val(row, "weather_code"))

    # Adjust wind for exposure
    eff_gust = gust * resort.wind_exposure
    eff_wind = wind * resort.wind_exposure

    # --- 1) WIND (biggest factor for lift ops + comfort) ---
    if eff_gust > 60:
        score -= 50
        concerns.append(f"Severe gusts ({gust:.0f} km/h)")
    elif eff_gust > 45:
        score -= clamp((eff_gust - 45) * 2.5, 0, 35)
        concerns.append(f"High gusts ({gust:.0f} km/h)")
    elif eff_gust > 30:
        score -= clamp((eff_gust - 30) * 0.8, 0, 12)

    score -= clamp((eff_wind - 20) * 0.3, 0, 10)

    # --- 2) VISIBILITY & FOG ---
    if vis < 500:
        score -= 40
        concerns.append("Very poor visibility (<500m)")
    elif vis < 1500:
        score -= clamp((1500 - vis) / 40, 0, 25)
        concerns.append(f"Low visibility ({vis:.0f}m)")
    elif vis < 3000:
        score -= clamp((3000 - vis) / 150, 0, 10)

    # Low cloud = flat light / whiteout risk
    if cloud_low > 80:
        score -= 15
        concerns.append("Heavy low cloud (flat light)")
    elif cloud_low > 60:
        score -= clamp((cloud_low - 60) / 4, 0, 8)

    # --- 3) PRECIPITATION ---
    if precip > 3.0:
        score -= 30
        concerns.append(f"Heavy precip ({precip:.1f}mm/h)")
    elif precip > 1.0:
        score -= clamp((precip - 1.0) * 12, 0, 24)
        concerns.append(f"Moderate precip ({precip:.1f}mm/h)")
    elif precip > 0.3:
        score -= clamp((precip - 0.3) * 6, 0, 8)

    # Fresh snow bonus (light snow, not windy)
    if snowfall > 0 and eff_gust < 40 and precip < 1.5:
        score += clamp(snowfall * 2, 0, 8)

    # --- 4) TEMPERATURE ---
    # Ideal: -8 to -3°C. Too warm = slush, too cold = miserable
    if temp > 0:
        score -= clamp(temp * 8, 0, 30)
        concerns.append(f"Above freezing ({temp:.1f}°C)")
    elif temp > -2:
        score -= clamp((temp + 2) * 4, 0, 10)
    elif temp < -18:
        score -= clamp((-18 - temp) * 2, 0, 20)
        concerns.append(f"Very cold ({temp:.1f}°C)")
    elif temp < -12:
        score -= clamp((-12 - temp) * 1.2, 0, 10)

    # Apparent temp (windchill) factor
    if apparent < -20:
        score -= clamp((-20 - apparent) * 0.8, 0, 12)

    # --- 5) FREEZING LEVEL ---
    # If freezing level well above resort, expect wet/slushy snow
    if frz > resort.elevation_m + 200:
        penalty = clamp((frz - resort.elevation_m - 200) * 0.015, 0, 20)
        score -= penalty
        if penalty > 10:
            concerns.append(f"High freezing level ({frz:.0f}m)")

    # --- 6) ASPECT-SPECIFIC ADJUSTMENTS ---
    if resort.aspect == "south":
        # South-facing: softens fast in sun when warm
        if temp > -3 and cloud < 50:
            score -= clamp((temp + 3) * 2.5, 0, 10)
    else:
        # North-facing: colder in shade, but snow lasts
        # Bonus for holding snow when it's warm elsewhere
        if temp > -2 and cloud < 60:
            score += clamp((temp + 5) * 0.8, 0, 6)

    # --- 7) WEATHER CODE PENALTIES ---
    # 95-99: thunderstorms
    if weather_code >= 95:
        score -= 40
        concerns.append("Thunderstorm risk")
    # 71-77: snow showers (usually fine, slight penalty for heavy)
    elif weather_code in (75, 77):
        score -= 5
    # 80-82: rain showers
    elif weather_code >= 80:
        score -= 15
        concerns.append("Rain showers")
    # 61-67: rain/freezing rain
    elif 61 <= weather_code <= 67:
        score -= 20
        concerns.append("Rain/freezing rain")

    return clamp(score, 0, 100), concerns


def extract_window_rows(
    hourly: dict[str, list], date_iso: str, start_h: int, end_h: int
) -> list[dict[str, Any]]:
    """Extract hourly data for a specific date and hour window."""
    times = hourly["time"]
    rows = []
    for i, t in enumerate(times):
        if not t.startswith(date_iso):
            continue
        hour = int(t[11:13])
        if start_h <= hour < end_h:
            row = {"time": t, "hour": hour}
            for k, v in hourly.items():
                if k != "time" and i < len(v):
                    row[k] = v[i]
            rows.append(row)
    return rows


def score_day(
    resort: Resort, hourly: dict[str, list], date_iso: str
) -> dict[str, Any]:
    """Score a single day for a resort."""
    rows = extract_window_rows(hourly, date_iso, START_HOUR, END_HOUR)
    if not rows:
        return {"score": 0, "concerns": ["No forecast data"], "summary": {}}

    scores = []
    all_concerns: list[str] = []
    for r in rows:
        s, c = hour_score(resort, r)
        scores.append(s)
        all_concerns.extend(c)

    # Aggregate summary stats
    def avg(key: str) -> float:
        vals = [get_val(r, key) for r in rows if get_val(r, key, None) is not None]
        return statistics.fmean(vals) if vals else 0.0

    def max_val(key: str) -> float:
        vals = [get_val(r, key) for r in rows]
        return max(vals) if vals else 0.0

    # Calculate basic summary stats
    temp_avg = avg("temperature_2m")
    temp_min = min(get_val(r, "temperature_2m") for r in rows)
    temp_max = max(get_val(r, "temperature_2m") for r in rows)
    gust_max = max_val("wind_gusts_10m")
    gust_eff_max = gust_max * resort.wind_exposure
    snowfall_total = sum(get_val(r, "snowfall") for r in rows)
    cloud_low_avg = avg("cloud_cover_low")
    vis_min = min(get_val(r, "visibility", 20000) for r in rows)
    sun_seconds = sum(get_val(r, "sunshine_duration") for r in rows)
    snow_depth_avg = avg("snow_depth")

    # Freeze-thaw risk: temps swing from below freezing to above
    freeze_thaw_risk = temp_max > 0 and temp_min < -2

    # Snow quality hints
    snow_quality_hint = None
    if snowfall_total > 2:
        if temp_avg < -6 and gust_eff_max < 40:
            snow_quality_hint = "Dry powder-ish"
        elif gust_eff_max > 45:
            snow_quality_hint = "Wind slab risk"

    # Flat light risk
    flat_light_risk = cloud_low_avg > 70 or vis_min < 1500

    summary = {
        "temp_avg": round(temp_avg, 1),
        "temp_min": round(temp_min, 1),
        "temp_max": round(temp_max, 1),
        "gust_max": round(gust_max, 0),
        "gust_eff_max": round(gust_eff_max, 0),
        "wind_avg": round(avg("wind_speed_10m"), 0),
        "precip_total": round(sum(get_val(r, "precipitation") for r in rows), 1),
        "snowfall_total": round(snowfall_total, 1),
        "cloud_avg": round(avg("cloud_cover"), 0),
        "cloud_low_avg": round(cloud_low_avg, 0),
        "vis_min": round(vis_min, 0),
        "sun_seconds": round(sun_seconds, 0),
        "snow_depth_avg": round(snow_depth_avg, 1),
        "freeze_thaw_risk": freeze_thaw_risk,
        "snow_quality_hint": snow_quality_hint,
        "flat_light_risk": flat_light_risk,
    }

    # Deduplicate concerns
    unique_concerns = list(dict.fromkeys(all_concerns))[:4]

    return {
        "score": round(statistics.fmean(scores), 1),
        "concerns": unique_concerns,
        "summary": summary,
    }


# =============================================================================
# DECISION ENGINE
# =============================================================================
# =============================================================================
# DECISION ENGINE
# =============================================================================
def rank_resorts(day_results: dict[str, dict[str, Any]]) -> list[tuple[str, dict[str, Any]]]:
    return sorted(day_results.items(), key=lambda kv: kv[1].get("score", 0), reverse=True)


def best_in_region(day_results: dict[str, dict[str, Any]], region: str) -> Optional[str]:
    candidates = [
        (short, res) for short, res in day_results.items()
        if RESORT_BY_SHORT[short].region == region
    ]
    if not candidates:
        return None
    return max(candidates, key=lambda kv: kv[1].get("score", 0))[0]


def lift_disruption_risk(gust_eff_max: float) -> str:
    """Map effective gust to lift disruption risk category."""
    if gust_eff_max >= 60:
        return "Very High"
    elif gust_eff_max >= 50:
        return "High"
    elif gust_eff_max >= 35:
        return "Moderate"
    else:
        return "Low"


def calculate_confidence(day_results: dict[str, dict[str, Any]], ranking: list[str]) -> tuple[str, float, float]:
    """
    Calculate confidence level based on score spreads.
    Returns: (confidence_label, spread_1_2, spread_1_5)
    """
    if len(ranking) < 2:
        return "High", 0.0, 0.0

    scores = [day_results[s]["score"] for s in ranking]
    spread_1_2 = scores[0] - scores[1]

    if len(ranking) >= 5:
        spread_1_5 = scores[0] - scores[4]
    else:
        spread_1_5 = 0.0

    # Determine confidence level
    if spread_1_2 >= 10:
        confidence = "High"
    elif spread_1_2 >= 5:
        confidence = "Medium"
    else:
        confidence = "Low"

    return confidence, spread_1_2, spread_1_5


def decide_day_multi(day_results: dict[str, dict[str, Any]]) -> tuple[str, str, str, list[str], str, float, float]:
    """
    Decide the best resort for the day.
    Returns: (pick_short, pick_emoji, reason, ranking, confidence, spread_1_2, spread_1_5)
    """
    ranking = rank_resorts(day_results)
    if not ranking:
        return "N/A", "❓", "No forecast data", [], "High", 0.0, 0.0

    pick_short = ranking[0][0]
    pick_emoji = RESORT_BY_SHORT[pick_short].emoji
    ranking_shorts = [s for s, _ in ranking]

    # Calculate confidence
    confidence, spread_1_2, spread_1_5 = calculate_confidence(day_results, ranking_shorts)

    if len(ranking) >= 2:
        second_short = ranking[1][0]
        diff = ranking[0][1]["score"] - ranking[1][1]["score"]
        if diff >= 8:
            reason = f"{pick_short} leads by {diff:.0f} pts over {second_short}"
        else:
            reason = f"Close call: {pick_short} +{diff:.0f} pts vs {second_short}"
    else:
        reason = "Best available score"

    return pick_short, pick_emoji, reason, ranking_shorts, confidence, spread_1_2, spread_1_5


def generate_forecast(days: int = FORECAST_DAYS) -> dict[str, Any]:
    today = dt.datetime.now(TZ).date()

    # Fetch forecasts once per resort
    forecasts = {r.short: fetch_forecast(r, forecast_days=days + 1) for r in ALL_RESORTS}

    out_days: list[dict[str, Any]] = []
    for i in range(days):
        d = today + dt.timedelta(days=i)
        date_iso = d.isoformat()
        weekday = d.strftime("%a")

        day_results: dict[str, dict[str, Any]] = {}
        for r in ALL_RESORTS:
            day_results[r.short] = score_day(r, forecasts[r.short]["hourly"], date_iso)

        pick, emoji, reason, ranking, confidence, spread_1_2, spread_1_5 = decide_day_multi(day_results)
        pick_engadin = best_in_region(day_results, "Engadin")
        pick_davos = best_in_region(day_results, "Davos")

        out_days.append({
            "date": date_iso,
            "weekday": weekday,
            "results": day_results,      # {short -> score_day(...) dict}
            "ranking": ranking,          # [short, short, ...] best -> worst
            "pick": pick,
            "emoji": emoji,
            "reason": reason,
            "pick_engadin": pick_engadin,
            "pick_davos": pick_davos,
            "confidence": confidence,
            "spread_1_2": spread_1_2,
            "spread_1_5": spread_1_5,
        })

    return {"generated_at": dt.datetime.now(TZ).isoformat(), "days": out_days}



# =============================================================================
# EMAIL FORMATTING
# =============================================================================
def summarize_conditions(day_results: dict[str, dict[str, Any]]) -> str:
    summaries = [r.get("summary", {}) for r in day_results.values()]

    temp_max = max((s.get("temp_avg", 0) for s in summaries), default=0)
    temp_min = min((s.get("temp_avg", 0) for s in summaries), default=0)
    precip_max = max((s.get("precip_total", 0) for s in summaries), default=0)
    snow_max = max((s.get("snowfall_total", 0) for s in summaries), default=0)
    gust_eff_max = max((s.get("gust_eff_max", 0) for s in summaries), default=0)
    sun_hours_max = max((s.get("sun_seconds", 0) / 3600 for s in summaries), default=0)
    flat_light = any(s.get("flat_light_risk", False) for s in summaries)
    freeze_thaw = any(s.get("freeze_thaw_risk", False) for s in summaries)

    # Get most common snow quality hint
    snow_hints = [s.get("snow_quality_hint") for s in summaries if s.get("snow_quality_hint")]
    snow_quality = snow_hints[0] if snow_hints else None

    parts: list[str] = []

    # Temperature
    if temp_max > 0:
        parts.append("☀️ Warm")
    elif temp_min < -12:
        parts.append("🥶 Cold")

    # Snow & precipitation
    if snow_max > 2 or precip_max > 5:
        parts.append("🌨️ Snowy")
        if snow_quality:
            parts.append(f"({snow_quality})")
    elif snow_max > 0.5 or precip_max > 1:
        parts.append("❄️ Light snow")

    # Wind & lift risk
    if gust_eff_max > 50:
        parts.append("💨 Windy")

    # Sunshine/visibility
    if flat_light:
        parts.append("☁️ Flat light")
    elif sun_hours_max > 5:
        parts.append(f"☀️ {sun_hours_max:.1f}h sun")

    # Freeze-thaw warning
    if freeze_thaw:
        parts.append("⚠️ Freeze-thaw")

    return " ".join(parts) if parts else "✓ Good"


def format_html_email(forecast: dict[str, Any]) -> str:
    days = forecast["days"]
    t = days[0]

    pick_region = RESORT_BY_SHORT[t["pick"]].region
    pick_class = "pick-engadin" if pick_region == "Engadin" else "pick-davos"

    # Concerns: pick + runner-up
    concerns: list[str] = []
    for short in t["ranking"][:2]:
        concerns.extend(t["results"][short].get("concerns", []))
    concerns = list(dict.fromkeys(concerns))[:5]

    # Show ALL resorts (today)
    top_cards = t["ranking"]


    html = """
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8" />
<style>
body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
       background: #f5f5f5; margin: 0; padding: 20px; }
.container { max-width: 680px; margin: 0 auto; background: white; border-radius: 12px;
             box-shadow: 0 2px 8px rgba(0,0,0,0.1); overflow: hidden; }
.header { background: linear-gradient(135deg, #1e3a5f 0%, #2d5a87 100%);
          color: white; padding: 24px; text-align: center; }
.header h1 { margin: 0 0 8px 0; font-size: 28px; }
.header .date { opacity: 0.9; font-size: 14px; }

.today { padding: 24px; border-bottom: 1px solid #eee; }
.pick-box { background: linear-gradient(135deg, #4CAF50 0%, #45a049 100%);
            color: white; padding: 18px; border-radius: 10px; text-align: center; margin-bottom: 16px; }
.pick-box .emoji { font-size: 44px; margin-bottom: 6px; }
.pick-box .resort { font-size: 24px; font-weight: 650; }
.pick-box .reason { font-size: 13px; opacity: 0.95; margin-top: 8px; line-height: 1.35; }

.scores { display: flex; flex-wrap: wrap; gap: 12px; margin-top: 14px; }
.score-card { flex: 1 1 45%; background: #f8f9fa; border-radius: 8px; padding: 12px; text-align: center; }
.score-card .name { font-size: 12px; color: #666; margin-bottom: 4px; }
.score-card .value { font-size: 28px; font-weight: 800; }
.score-card.winner .value { color: #4CAF50; }
.metric { font-size: 11px; color: #888; }

.concerns { background: #fff3cd; border-left: 4px solid #ffc107; padding: 12px;
            border-radius: 0 8px 8px 0; margin-top: 14px; }
.concerns h4 { margin: 0 0 8px 0; font-size: 13px; color: #856404; }
.concerns ul { margin: 0; padding-left: 20px; font-size: 13px; color: #856404; }

.forecast { padding: 20px; }
.forecast h3 { margin: 0 0 16px 0; font-size: 16px; color: #333; }
.forecast-table { width: 100%; border-collapse: collapse; font-size: 13px; }
.forecast-table th { text-align: left; padding: 10px 8px; border-bottom: 2px solid #ddd;
                     color: #666; font-weight: 600; }
.forecast-table td { padding: 10px 8px; border-bottom: 1px solid #eee; vertical-align: top; }
.forecast-table tr:hover { background: #f8f9fa; }
.pick-cell { font-weight: 650; }
.pick-engadin { color: #1e88e5; }
.pick-davos { color: #8e24aa; }

.footer { background: #f8f9fa; padding: 16px; text-align: center; font-size: 12px; color: #666; }
</style>
</head>
<body>
<div class="container">
"""
    html += f"""
<div class="header">
    <h1>🎿 Graubünden Ski Report</h1>
    <div class="date">{dt.datetime.now(TZ).strftime('%A, %B %d, %Y')}</div>
</div>
"""

    eng = t.get("pick_engadin")
    dav = t.get("pick_davos")
    eng_txt = f"{RESORT_BY_SHORT[eng].emoji} {eng}" if eng else "-"
    dav_txt = f"{RESORT_BY_SHORT[dav].emoji} {dav}" if dav else "-"

    confidence = t.get("confidence", "N/A")
    confidence_emoji = "🎯" if confidence == "High" else "🎲" if confidence == "Low" else "⚖️"

    html += f"""
<div class="today">
  <div class="pick-box">
    <div class="emoji">{t['emoji']}</div>
    <div class="resort">Today: {t['pick']}</div>
    <div class="reason">
      {t['reason']}<br/>
      <span class="metric">Confidence: {confidence_emoji} {confidence} • Best Engadin: {eng_txt} • Best Davos: {dav_txt}</span>
    </div>
  </div>

  <div class="scores">
"""
    for short in top_cards:
        rdef = RESORT_BY_SHORT[short]
        res = t["results"][short]
        summary = res.get("summary", {})
        score = res.get("score", 0)
        temp_avg = summary.get("temp_avg", "-")
        temp_min = summary.get("temp_min", "-")
        temp_max = summary.get("temp_max", "-")
        gust_eff = summary.get("gust_eff_max", "-")
        snow_depth = summary.get("snow_depth_avg", 0)

        # Lift disruption risk
        lift_risk = lift_disruption_risk(gust_eff if isinstance(gust_eff, (int, float)) else 0)
        lift_emoji = "🟢" if lift_risk == "Low" else "🟡" if lift_risk == "Moderate" else "🟠" if lift_risk == "High" else "🔴"

        # Build metric line
        metric_parts = [f"{temp_avg}°C ({temp_min} to {temp_max})"]
        metric_parts.append(f"💨 {gust_eff} km/h {lift_emoji}")
        if snow_depth > 0:
            metric_parts.append(f"Base {snow_depth:.0f}cm")

        html += f"""
    <div class="score-card {'winner' if short == t['pick'] else ''}">
      <div class="name">{rdef.emoji} {short} <span class="metric">({rdef.region})</span></div>
      <div class="value">{score:.0f}</div>
      <div class="metric">{' | '.join(metric_parts)}</div>
    </div>
"""
    html += "  </div>"

    if concerns:
        html += """
  <div class="concerns">
    <h4>⚠️ Watch Out For</h4>
    <ul>
"""
        for c in concerns:
            html += f"      <li>{c}</li>\n"
        html += """
    </ul>
  </div>
"""

    html += "</div>"

    html += """
<div class="forecast">
  <h3>📅 5-Day Outlook</h3>
  <table class="forecast-table">
    <tr>
      <th>Day</th>
      <th>Pick</th>
      <th>All scores</th>

      <th>Conditions</th>
    </tr>
"""
    for day in days:
        pick = day["pick"]
        pick_region = RESORT_BY_SHORT[pick].region
        pick_class = "pick-engadin" if pick_region == "Engadin" else "pick-davos"

        scores_html = "<br/>".join(
            f"{RESORT_BY_SHORT[s].emoji} {s} {day['results'][s]['score']:.0f}"
            for s in day["ranking"]
        )
        cond = summarize_conditions(day["results"])

        html += f"""
    <tr>
      <td><strong>{day['weekday']}</strong><br><span class="metric">{day['date'][5:]}</span></td>
      <td class="pick-cell {pick_class}">{RESORT_BY_SHORT[pick].emoji} {pick}</td>
      <td style="font-size:12px; line-height:1.35;">{scores_html}</td>
      <td>{cond}</td>
    </tr>
"""
    html += f"""
  </table>
</div>

<div class="footer">
  Generated at {forecast['generated_at'][:16].replace('T', ' ')} CET<br>
  Scores: 0-100 (higher = better) | Data: Open-Meteo
</div>

</div>
</body>
</html>
"""
    return html


def format_plain_email(forecast: dict[str, Any]) -> str:
    days = forecast["days"]
    t = days[0]
    eng = t.get("pick_engadin")
    dav = t.get("pick_davos")
    confidence = t.get("confidence", "N/A")

    lines = [
        "🎿 GRAUBÜNDEN SKI REPORT",
        f"   {dt.datetime.now(TZ).strftime('%A, %B %d, %Y')}",
        "",
        "=" * 48,
        f"TODAY'S PICK: {t['pick'].upper()} {t['emoji']}  ({RESORT_BY_SHORT[t['pick']].region})",
        f"Reason: {t['reason']}",
        f"Confidence: {confidence}",
        f"Best Engadin: {eng or '-'}   |   Best Davos: {dav or '-'}",
        "",
        "Top scores today:",
    ]

    for short in t["ranking"]:
        res = t["results"][short]
        summary = res.get("summary", {})
        gust_eff = summary.get("gust_eff_max", 0)
        lift_risk = lift_disruption_risk(gust_eff)
        lines.append(f"  {RESORT_BY_SHORT[short].emoji} {short:<10} {res['score']:>5.1f}  (Wind: {gust_eff:.0f} km/h, Lift: {lift_risk})")

    # concerns
    concerns: list[str] = []
    for short in t["ranking"][:2]:
        concerns.extend(t["results"][short].get("concerns", []))
    concerns = list(dict.fromkeys(concerns))[:5]
    if concerns:
        lines.append("")
        lines.append("Watch out for:")
        for c in concerns:
            lines.append(f"  • {c}")

    lines.extend(["", "=" * 48, "5-DAY OUTLOOK", ""])
    for d in days:
        cond = summarize_conditions(d["results"])
        all_txt = ", ".join(
            f"{s} {d['results'][s]['score']:.0f}"
            for s in d["ranking"]
        )
        lines.append(
            f"{d['weekday']} {d['date'][5:]:>5} | {d['pick']:<10} | {all_txt} | {cond}"
        )


    lines.extend(["", "-" * 48, f"Generated: {forecast['generated_at'][:16]}"])
    return "\n".join(lines)



# =============================================================================
# EMAIL SENDING
# =============================================================================
def send_email(forecast: dict[str, Any]) -> bool:
    """Send the forecast email."""
    today = forecast["days"][0]
    subject = f"🎿 {today['pick']} Today | St. Moritz {dt.datetime.now(TZ).strftime('%b %d')}"

    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"] = SENDER_EMAIL
    msg["To"] = RECIPIENT_EMAIL

    plain = format_plain_email(forecast)
    html = format_html_email(forecast)

    msg.attach(MIMEText(plain, "plain"))
    msg.attach(MIMEText(html, "html"))

    missing = [k for k, v in {
        "SMTP_SENDER_EMAIL": SENDER_EMAIL,
        "SMTP_PASSWORD": SENDER_PASSWORD,
        "SMTP_RECIPIENT_EMAIL": RECIPIENT_EMAIL,
    }.items() if not v]

    if missing:
        print(f"✗ Missing env vars: {', '.join(missing)}; email not sent.")
        return False

    try:
        with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as server:
            server.starttls()
            server.login(SENDER_EMAIL, SENDER_PASSWORD)
            server.sendmail(SENDER_EMAIL, RECIPIENT_EMAIL, msg.as_string())
        print(f"✓ Email sent to {RECIPIENT_EMAIL}")
        return True
    except Exception as e:
        print(f"✗ Email failed: {e}")
        return False


# =============================================================================
# MAIN
# =============================================================================
def main():
    """Generate forecast and send email."""
    import argparse
    import io

    # Fix Windows console encoding issues
    if sys.platform == "win32":
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

    parser = argparse.ArgumentParser(description="St. Moritz Ski Decision Engine")
    parser.add_argument("--dry-run", action="store_true", help="Generate report without sending email")
    parser.add_argument("--html-out", type=str, help="Save HTML to file (for preview)")
    args = parser.parse_args()

    print(f"Generating forecast at {dt.datetime.now(TZ).isoformat()}...")

    try:
        forecast = generate_forecast(FORECAST_DAYS)
    except Exception as e:
        print(f"✗ Forecast generation failed: {e}")
        sys.exit(1)

    # Print summary (multi-resort)
    today = forecast["days"][0]
    pick = today["pick"]
    pick_region = RESORT_BY_SHORT[pick].region
    confidence = today.get("confidence", "N/A")
    print(f"\nToday's pick: {pick} ({pick_region}) — {today['reason']}")
    print(f"Confidence: {confidence}")

    eng = today.get("pick_engadin")
    dav = today.get("pick_davos")
    print(f"Best Engadin: {eng or '-'} | Best Davos: {dav or '-'}")

    print("\nTop today:")
    for short in today["ranking"][:6]:
        res = today["results"][short]
        rdef = RESORT_BY_SHORT[short]
        summary = res.get("summary", {})
        gust_eff = summary.get("gust_eff_max", 0)
        lift_risk = lift_disruption_risk(gust_eff)
        print(f"  {rdef.emoji} {short:<10} {res['score']:>5.1f} ({rdef.region}) | Wind: {gust_eff:.0f} km/h ({lift_risk} lift risk)")

    # Print 5-day outlook (pick + top3)
    print("\n5-Day Outlook:")
    for d in forecast["days"]:
        top3 = d["ranking"][:3]
        top3_str = ", ".join(f"{s}:{d['results'][s]['score']:.0f}" for s in top3)
        print(
            f"  {d['weekday']} {d['date'][5:]:>5} | Pick {d['pick']:<10} | "
            f"Eng {d.get('pick_engadin') or '-':<10} | Dav {d.get('pick_davos') or '-':<10} | "
            f"Top3 {top3_str}"
        )

    # Save HTML preview if requested
    if args.html_out:
        html = format_html_email(forecast)
        with open(args.html_out, "w", encoding="utf-8") as f:
            f.write(html)

        print(f"\n✓ HTML saved to {args.html_out}")
    
    # Send email unless dry-run
    if args.dry_run:
        print("\n[Dry run - email not sent]")
    else:
        if not send_email(forecast):
            sys.exit(1)


if __name__ == "__main__":
    main()
