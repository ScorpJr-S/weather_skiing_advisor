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
@dataclass(frozen=True)
class Resort:
    name: str
    short: str
    lat: float
    lon: float
    elevation_m: int
    aspect: str  # "south" or "north"
    wind_exposure: float  # 1.0 = normal, >1 = more exposed


CORVIGLIA = Resort(
    name="Corviglia (St. Moritz)",
    short="Corviglia",
    lat=46.5079,
    lon=9.8192,
    elevation_m=2486,
    aspect="south",
    wind_exposure=0.85,  # More sheltered
)

CORVATSCH = Resort(
    name="Corvatsch 3303",
    short="Corvatsch",
    lat=46.4179,
    lon=9.8212,
    elevation_m=3303,
    aspect="north",
    wind_exposure=1.25,  # Notoriously windy at top
)


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

    summary = {
        "temp_avg": round(avg("temperature_2m"), 1),
        "temp_min": round(min(get_val(r, "temperature_2m") for r in rows), 1),
        "gust_max": round(max_val("wind_gusts_10m"), 0),
        "wind_avg": round(avg("wind_speed_10m"), 0),
        "precip_total": round(sum(get_val(r, "precipitation") for r in rows), 1),
        "snowfall_total": round(sum(get_val(r, "snowfall") for r in rows), 1),
        "cloud_avg": round(avg("cloud_cover"), 0),
        "vis_min": round(min(get_val(r, "visibility", 20000) for r in rows), 0),
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
def decide_day(
    corviglia_result: dict, corvatsch_result: dict
) -> tuple[str, str, str]:
    """
    Returns (winner_short, winner_emoji, reason).
    """
    c1 = corviglia_result["score"]
    c2 = corvatsch_result["score"]
    diff = c2 - c1

    # Strong preference threshold
    if diff > 8:
        return "Corvatsch", "🏔️", f"Corvatsch scores {diff:.0f} pts higher"
    elif diff < -8:
        return "Corviglia", "⛷️", f"Corviglia scores {-diff:.0f} pts higher"

    # Close call - use tiebreakers
    s1, s2 = corviglia_result["summary"], corvatsch_result["summary"]

    # If warm, prefer north-facing (Corvatsch)
    if s1["temp_avg"] > -2:
        return "Corvatsch", "🏔️", "Close scores; warmer day favors north-facing slopes"

    # If Corvatsch much windier, prefer Corviglia
    if s2["gust_max"] > s1["gust_max"] + 15:
        return "Corviglia", "⛷️", "Close scores; Corvatsch looks windier"

    # If heavy precip, prefer lower/easier access
    if s1["precip_total"] > 5:
        return "Corviglia", "⛷️", "Close scores; stormy day, easier access wins"

    # Default to Corviglia (easier, more reliable)
    if diff >= 0:
        return "Corvatsch", "🏔️", "Marginal advantage to Corvatsch"
    return "Corviglia", "⛷️", "Marginal advantage to Corviglia"


def generate_forecast(days: int = FORECAST_DAYS) -> dict[str, Any]:
    """Generate full forecast comparison."""
    today = dt.datetime.now(TZ).date()

    # Fetch forecasts once per resort
    fc_corv = fetch_forecast(CORVIGLIA, forecast_days=days + 1)
    fc_cort = fetch_forecast(CORVATSCH, forecast_days=days + 1)

    results = []
    for i in range(days):
        d = today + dt.timedelta(days=i)
        date_iso = d.isoformat()
        weekday = d.strftime("%a")

        r1 = score_day(CORVIGLIA, fc_corv["hourly"], date_iso)
        r2 = score_day(CORVATSCH, fc_cort["hourly"], date_iso)
        winner, emoji, reason = decide_day(r1, r2)

        results.append({
            "date": date_iso,
            "weekday": weekday,
            "corviglia": r1,
            "corvatsch": r2,
            "pick": winner,
            "emoji": emoji,
            "reason": reason,
        })

    return {
        "generated_at": dt.datetime.now(TZ).isoformat(),
        "days": results,
    }


# =============================================================================
# EMAIL FORMATTING
# =============================================================================
def format_html_email(forecast: dict[str, Any]) -> str:
    """Generate HTML email body."""
    days = forecast["days"]
    today = days[0]

    # Styles
    html = """
<!DOCTYPE html>
<html>
<head>
<style>
body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; 
       background: #f5f5f5; margin: 0; padding: 20px; }
.container { max-width: 600px; margin: 0 auto; background: white; border-radius: 12px; 
             box-shadow: 0 2px 8px rgba(0,0,0,0.1); overflow: hidden; }
.header { background: linear-gradient(135deg, #1e3a5f 0%, #2d5a87 100%); 
          color: white; padding: 24px; text-align: center; }
.header h1 { margin: 0 0 8px 0; font-size: 28px; }
.header .date { opacity: 0.9; font-size: 14px; }
.today { padding: 24px; border-bottom: 1px solid #eee; }
.pick-box { background: linear-gradient(135deg, #4CAF50 0%, #45a049 100%); 
            color: white; padding: 20px; border-radius: 10px; text-align: center; margin-bottom: 16px; }
.pick-box .emoji { font-size: 48px; margin-bottom: 8px; }
.pick-box .resort { font-size: 24px; font-weight: 600; }
.pick-box .reason { font-size: 13px; opacity: 0.95; margin-top: 8px; }
.scores { display: flex; gap: 12px; margin-bottom: 16px; }
.score-card { flex: 1; background: #f8f9fa; border-radius: 8px; padding: 12px; text-align: center; }
.score-card .name { font-size: 12px; color: #666; margin-bottom: 4px; }
.score-card .value { font-size: 28px; font-weight: 700; }
.score-card.winner .value { color: #4CAF50; }
.concerns { background: #fff3cd; border-left: 4px solid #ffc107; padding: 12px; 
            border-radius: 0 8px 8px 0; margin-top: 12px; }
.concerns h4 { margin: 0 0 8px 0; font-size: 13px; color: #856404; }
.concerns ul { margin: 0; padding-left: 20px; font-size: 13px; color: #856404; }
.forecast { padding: 20px; }
.forecast h3 { margin: 0 0 16px 0; font-size: 16px; color: #333; }
.forecast-table { width: 100%; border-collapse: collapse; font-size: 13px; }
.forecast-table th { text-align: left; padding: 10px 8px; border-bottom: 2px solid #ddd; 
                     color: #666; font-weight: 600; }
.forecast-table td { padding: 10px 8px; border-bottom: 1px solid #eee; }
.forecast-table tr:hover { background: #f8f9fa; }
.pick-cell { font-weight: 600; }
.pick-corviglia { color: #2196F3; }
.pick-corvatsch { color: #9C27B0; }
.footer { background: #f8f9fa; padding: 16px; text-align: center; font-size: 12px; color: #666; }
.metric { font-size: 11px; color: #888; }
</style>
</head>
<body>
<div class="container">
"""

    # Header
    html += f"""
<div class="header">
    <h1>🎿 St. Moritz Ski Report</h1>
    <div class="date">{dt.datetime.now(TZ).strftime('%A, %B %d, %Y')}</div>
</div>
"""

    # Today's pick
    t = today
    winner_is_corv = t["pick"] == "Corviglia"
    html += f"""
<div class="today">
    <div class="pick-box">
        <div class="emoji">{t['emoji']}</div>
        <div class="resort">Today: {t['pick']}</div>
        <div class="reason">{t['reason']}</div>
    </div>
    
    <div class="scores">
        <div class="score-card {'winner' if winner_is_corv else ''}">
            <div class="name">Corviglia</div>
            <div class="value">{t['corviglia']['score']:.0f}</div>
            <div class="metric">{t['corviglia']['summary'].get('temp_avg', 'N/A')}°C | 
                 💨 {t['corviglia']['summary'].get('gust_max', 'N/A')} km/h</div>
        </div>
        <div class="score-card {'winner' if not winner_is_corv else ''}">
            <div class="name">Corvatsch</div>
            <div class="value">{t['corvatsch']['score']:.0f}</div>
            <div class="metric">{t['corvatsch']['summary'].get('temp_avg', 'N/A')}°C | 
                 💨 {t['corvatsch']['summary'].get('gust_max', 'N/A')} km/h</div>
        </div>
    </div>
"""

    # Concerns
    all_concerns = t["corviglia"]["concerns"] + t["corvatsch"]["concerns"]
    if all_concerns:
        html += """
    <div class="concerns">
        <h4>⚠️ Watch Out For</h4>
        <ul>
"""
        for c in list(dict.fromkeys(all_concerns))[:4]:
            html += f"            <li>{c}</li>\n"
        html += """
        </ul>
    </div>
"""

    html += "</div>"

    # 5-day forecast table
    html += """
<div class="forecast">
    <h3>📅 5-Day Outlook</h3>
    <table class="forecast-table">
        <tr>
            <th>Day</th>
            <th>Pick</th>
            <th>Corviglia</th>
            <th>Corvatsch</th>
            <th>Conditions</th>
        </tr>
"""

    for day in days:
        pick_class = "pick-corviglia" if day["pick"] == "Corviglia" else "pick-corvatsch"
        c1, c2 = day["corviglia"], day["corvatsch"]

        # Condition summary
        cond_parts = []
        temp = c1["summary"].get("temp_avg", 0)
        if temp > 0:
            cond_parts.append("☀️ Warm")
        elif temp < -10:
            cond_parts.append("🥶 Cold")

        precip = c1["summary"].get("precip_total", 0) + c2["summary"].get("precip_total", 0)
        if precip > 5:
            cond_parts.append("🌨️ Snowy")
        elif precip > 1:
            cond_parts.append("❄️ Light snow")

        gust = max(c1["summary"].get("gust_max", 0), c2["summary"].get("gust_max", 0))
        if gust > 50:
            cond_parts.append("💨 Windy")

        cond = " ".join(cond_parts) if cond_parts else "✓ Good"

        html += f"""
        <tr>
            <td><strong>{day['weekday']}</strong><br><span class="metric">{day['date'][5:]}</span></td>
            <td class="pick-cell {pick_class}">{day['emoji']} {day['pick']}</td>
            <td>{c1['score']:.0f}<br><span class="metric">{c1['summary'].get('temp_avg', '-')}°C</span></td>
            <td>{c2['score']:.0f}<br><span class="metric">{c2['summary'].get('temp_avg', '-')}°C</span></td>
            <td>{cond}</td>
        </tr>
"""

    html += """
    </table>
</div>
"""

    # Footer
    html += f"""
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
    """Generate plain text fallback."""
    days = forecast["days"]
    today = days[0]

    lines = [
        "🎿 ST. MORITZ SKI REPORT",
        f"   {dt.datetime.now(TZ).strftime('%A, %B %d, %Y')}",
        "",
        "=" * 40,
        f"TODAY'S PICK: {today['pick'].upper()} {today['emoji']}",
        f"Reason: {today['reason']}",
        "",
        f"Scores: Corviglia {today['corviglia']['score']:.0f} | Corvatsch {today['corvatsch']['score']:.0f}",
        "",
    ]

    concerns = today["corviglia"]["concerns"] + today["corvatsch"]["concerns"]
    if concerns:
        lines.append("Watch out for:")
        for c in list(dict.fromkeys(concerns))[:4]:
            lines.append(f"  • {c}")
        lines.append("")

    lines.extend(["=" * 40, "5-DAY OUTLOOK", ""])

    for d in days:
        lines.append(
            f"{d['weekday']} {d['date'][5:]:>5} | {d['pick']:>9} | "
            f"Corv:{d['corviglia']['score']:>3.0f} Cort:{d['corvatsch']['score']:>3.0f}"
        )

    lines.extend([
        "",
        "-" * 40,
        f"Generated: {forecast['generated_at'][:16]}",
    ])

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

    # Print summary
    today = forecast["days"][0]
    print(f"\nToday's pick: {today['pick']} ({today['reason']})")
    print(f"Scores: Corviglia={today['corviglia']['score']:.1f}, Corvatsch={today['corvatsch']['score']:.1f}")
    
    # Print 5-day outlook
    print("\n5-Day Outlook:")
    for d in forecast["days"]:
        print(f"  {d['weekday']} {d['date'][5:]:>5} | {d['pick']:>9} | "
              f"Corviglia:{d['corviglia']['score']:>5.1f}  Corvatsch:{d['corvatsch']['score']:>5.1f}")
    
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
