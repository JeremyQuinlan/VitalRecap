"""
The Victory Lane — Market Intelligence Digest
----------------------------------------------
Polls Yahoo Mail via IMAP for newsletters, summarizes them with Claude,
saves styled HTML digests, and commits to GitHub Pages.

Folder: C:\\Tools\\TheVictoryLane\\
Setup:
  pip install anthropic requests
  Then configure the CONFIG block below for local use.

TO RUN LOCALLY:
  python vital_knowledge_digest.py
"""

import imaplib
import email
import json
import os
import re
import subprocess
import requests
import anthropic
from datetime import datetime, timedelta, timezone
from email.header import decode_header
from email.utils import parsedate_to_datetime
from pathlib import Path
from zoneinfo import ZoneInfo

EASTERN = ZoneInfo("America/New_York")

# ─────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────
CONFIG = {
    "yahoo_email":        os.environ.get("YAHOO_EMAIL",        "YOUR_EMAIL@yahoo.com"),
    "yahoo_app_password": os.environ.get("YAHOO_APP_PASSWORD", "YOUR_APP_PASSWORD"),
    "anthropic_api_key":  os.environ.get("ANTHROPIC_API_KEY",  "YOUR_ANTHROPIC_API_KEY"),
    "sender_filter":      "vitalknowledge",
    "lookback_hours":     168,
    "html_output":        r"C:\Tools\VitalRecap\digest.html",
    "state_file":         r"C:\Tools\VitalRecap\processed_ids.json",
    "edge_exe":           r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    "tts_rate":           1.0,
    "github_actions":     os.environ.get("GITHUB_ACTIONS", "false").lower() == "true",
}
# ─────────────────────────────────────────────

SITE_NAME = "The Victory Lane"

SUMMARIZE_PROMPT = """You are summarizing a financial newsletter for a trading team called Victory Lane.
Your job is to produce a concise but substantive digest that captures the author's voice — direct, confident,
uses parenthetical asides, cites sources inline (NYT, WSJ, FT, Bloomberg, etc.), and isn't shy about
giving a market view. Write in flowing prose for the narrative sections, not just bullet fragments.

Structure your response in exactly this order using these exact section headers:

## MARKETS
Report SPX and Nasdaq as PERCENTAGE CHANGE only (e.g. -0.26%, +0.54%).
Do NOT use basis points or point values for SPX and Nasdaq.
Also include Dow, R2K, Brent, Gold, Silver, BTC, DXY if present — these can use their native units.

## RATES & FED
Treasury move and Fed expectations. Keep to 2-3 sentences.
When referencing basis point moves on yields, write "basis points" in full, never "bp".

## MARKET OUTLOOK
Editorial view on where markets are headed. 3-5 sentences, confident tone.

## GEOPOLITICAL
Key geopolitical developments. Prose bullets with source attribution in parentheses.
No repeated information — consolidate if the newsletter says the same thing twice.

## COMPANY NEWS
Individual company items with ticker in bold. Include EPS/revenue figures where reported.
One paragraph per company. Cite sources. No duplicates.

## MACRO & FED DATES
Upcoming macro, Fed, and consumer data dates. Format: DATE — description.
Include context on what to watch for where relevant.

## EARNINGS THIS WEEK
Pre and post market, grouped by day.

## EARNINGS NEXT 2 WEEKS
Pre and post market, grouped by day.

Newsletter content:
{body}

Important rules:
- Write in a direct, confident voice throughout
- Do NOT mention "Vital Knowledge", "Vital", "Dawn", or the newsletter's name anywhere in your output
- Consolidate any repeated news items — mention each story only once
- Include specific numbers and source attributions
- Always write "basis points" in full, never abbreviate as "bp"
- Keep total length under 4000 characters for the TTS readout
- Do not include scheduling notes, subscription info, or technical notices"""


def clean_subject(raw_subject):
    """Rewrite email subject to Victory Lane branding, stripping VK references."""
    s = re.sub(r"^Vital Knowledge:\s*", "", raw_subject).strip()

    # Extract date portion if present
    date_match = re.search(r"(Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday)\s+\w+\s+\d+,?\s*\d{4}", s)
    date_str = date_match.group(0) if date_match else ""

    s_lower = s.lower()

    if "dawn" in s_lower or "morning" in s_lower:
        return f"Morning Intelligentsia · {date_str}".strip(" ·")
    elif "mid-day" in s_lower or "midday" in s_lower:
        return f"Mid-Day Update · {date_str}".strip(" ·")
    elif "recap" in s_lower or "close" in s_lower:
        return f"Market Recap · {date_str}".strip(" ·")
    else:
        # Strip any remaining "vital" or "knowledge" from other subjects
        s = re.sub(r"\b(vital|knowledge|dawn)\b", "", s, flags=re.IGNORECASE)
        s = re.sub(r"\s+", " ", s).strip(" -·")
        return s


def get_category_tag(subject):
    s = subject.lower()
    if "morning intelligentsia" in s or "morning" in s:
        return "MORNING"
    elif "mid-day" in s or "midday" in s:
        return "MID-DAY"
    elif "recap" in s or "close" in s:
        return "RECAP"
    elif "intraday" in s:
        return "INTRADAY"
    else:
        return "UPDATE"


def get_tag_color(tag):
    colors = {
        "MORNING":  ("#1a3a2a", "#4caf82"),
        "MID-DAY":  ("#1a2a3a", "#4c8faf"),
        "RECAP":    ("#2a1a1a", "#af4c4c"),
        "INTRADAY": ("#2a2a1a", "#af9f4c"),
        "UPDATE":   ("#2a1a3a", "#8f4caf"),
    }
    return colors.get(tag, ("#1a1a1a", "#888888"))


def load_processed_ids(state_file):
    if os.path.exists(state_file):
        with open(state_file) as f:
            return set(json.load(f))
    return set()


def save_processed_ids(state_file, ids):
    os.makedirs(os.path.dirname(state_file), exist_ok=True)
    with open(state_file, "w") as f:
        json.dump(list(ids), f)


def load_processed_ids_github():
    path = "processed_ids.json"
    if os.path.exists(path):
        with open(path, encoding="utf-8") as f:
            return set(json.load(f))
    return set()


def save_processed_ids_github(ids):
    with open("processed_ids.json", "w", encoding="utf-8") as f:
        json.dump(list(ids), f)


def decode_str(s):
    if s is None:
        return ""
    parts = decode_header(s)
    result = []
    for part, enc in parts:
        if isinstance(part, bytes):
            result.append(part.decode(enc or "utf-8", errors="replace"))
        else:
            result.append(part)
    return " ".join(result)


def get_email_body(msg):
    body = ""
    if msg.is_multipart():
        for part in msg.walk():
            ct = part.get_content_type()
            cd = str(part.get("Content-Disposition", ""))
            if ct == "text/plain" and "attachment" not in cd:
                body = part.get_payload(decode=True).decode("utf-8", errors="replace")
                break
            elif ct == "text/html" and "attachment" not in cd and not body:
                raw_html = part.get_payload(decode=True).decode("utf-8", errors="replace")
                body = re.sub(r"<[^>]+>", " ", raw_html)
                body = re.sub(r"\s+", " ", body).strip()
    else:
        body = msg.get_payload(decode=True).decode("utf-8", errors="replace")
    return body[:14000]


def get_email_date(msg):
    date_str = msg.get("Date", "")
    try:
        dt = parsedate_to_datetime(date_str)
        dt_eastern = dt.astimezone(EASTERN)
        return dt_eastern.strftime("%b %d, %Y · %I:%M %p ET")
    except Exception:
        return datetime.now(EASTERN).strftime("%b %d, %Y · %I:%M %p ET")


def get_email_sent_utc(msg):
    date_str = msg.get("Date", "")
    try:
        dt = parsedate_to_datetime(date_str)
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=timezone.utc)
        return dt.astimezone(timezone.utc)
    except Exception:
        return datetime.now(timezone.utc)


def fetch_new_emails(config):
    mail = imaplib.IMAP4_SSL("imap.mail.yahoo.com", 993)
    mail.login(config["yahoo_email"], config["yahoo_app_password"])
    mail.select("inbox")

    since_date = (datetime.utcnow() - timedelta(hours=config["lookback_hours"] + 24)).strftime("%d-%b-%Y")
    status, data = mail.search(None, f'(SINCE "{since_date}")')

    uids = data[0].split()
    results = []

    if config["github_actions"]:
        processed = load_processed_ids_github()
    else:
        processed = load_processed_ids(config["state_file"])

    cutoff = datetime.now(timezone.utc) - timedelta(hours=config["lookback_hours"])

    for uid in uids:
        uid_str = uid.decode()
        if uid_str in processed:
            print(f"  Skipping already processed UID {uid_str}")
            continue

        status, msg_data = mail.fetch(uid, "(RFC822)")
        raw = msg_data[0][1]
        msg = email.message_from_bytes(raw)

        from_addr = decode_str(msg.get("From", ""))
        if config["sender_filter"].lower() not in from_addr.lower():
            continue

        sent_utc = get_email_sent_utc(msg)
        if sent_utc < cutoff:
            print(f"  Skipping old email (sent {sent_utc.strftime('%Y-%m-%d %H:%M UTC')})")
            continue

        raw_subject = decode_str(msg.get("Subject", "(No Subject)"))
        subject = clean_subject(raw_subject)
        body = get_email_body(msg)
        email_date = get_email_date(msg)

        if body.strip():
            results.append((uid_str, subject, body, email_date, sent_utc))
            print(f"  Found: {subject} ({email_date})")

    mail.logout()
    results.sort(key=lambda x: x[4])
    return results, processed


def summarize_with_claude(body, api_key):
    client = anthropic.Anthropic(api_key=api_key)
    message = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=1800,
        messages=[{"role": "user", "content": SUMMARIZE_PROMPT.format(body=body)}]
    )
    return message.content[0].text


def get_preview(digest_text):
    lines = [l.strip() for l in digest_text.split("\n") if l.strip() and not l.startswith("#")]
    if lines:
        preview = lines[0][:140]
        if len(lines[0]) > 140:
            preview += "..."
        return preview
    return ""


def markdown_to_html_body(text):
    lines = text.split("\n")
    html_parts = []
    for line in lines:
        line = line.strip()
        if not line:
            continue
        if line.startswith("## "):
            html_parts.append(f'<h2>{line[3:]}</h2>')
        elif line.startswith("**") and line.endswith("**"):
            html_parts.append(f'<p class="ticker-line">{line[2:-2]}</p>')
        else:
            line = re.sub(r"\*\*(.+?)\*\*", r"<strong>\1</strong>", line)
            html_parts.append(f'<p>{line}</p>')
    return "\n".join(html_parts)


def prepare_tts_text(html_body):
    """Convert HTML digest to clean TTS text.
    Skips sections with fewer than 120 chars of content (placeholder sections).
    Also strips any remaining VK brand references.
    """
    sections = re.split(r"<h2>[^<]*</h2>", html_body)
    headers = re.findall(r"<h2>([^<]*)</h2>", html_body)

    tts_parts = []

    for i, section_html in enumerate(sections):
        section_text = re.sub(r"<[^>]+>", " ", section_html)
        section_text = re.sub(r"\s+", " ", section_text).strip()

        if i == 0:
            if section_text:
                tts_parts.append(section_text)
            continue

        header = headers[i - 1] if i - 1 < len(headers) else ""

        if len(section_text) < 120:
            continue

        if header:
            tts_parts.append(header + ".")
        tts_parts.append(section_text)

    text = " ".join(tts_parts)

    # Strip any remaining brand references
    text = re.sub(r"\b(Vital Knowledge|Vital Dawn|Vital)\b", "", text, flags=re.IGNORECASE)

    # Abbreviation expansions
    text = re.sub(r"\bbp\b", "basis points", text)
    text = re.sub(r"\bBP\b", "basis points", text)
    text = re.sub(r"\bpct\b", "percent", text, flags=re.IGNORECASE)
    text = re.sub(r"\bYoY\b", "year over year", text, flags=re.IGNORECASE)
    text = re.sub(r"\bQoQ\b", "quarter over quarter", text, flags=re.IGNORECASE)
    text = re.sub(r"\bY/Y\b", "year over year", text)
    text = re.sub(r"\bQ/Q\b", "quarter over quarter", text)
    text = re.sub(r"\bEPS\b", "earnings per share", text)
    text = re.sub(r"\bET\b", "Eastern time", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def build_html(digest_text, subject, email_date, tts_rate):
    body_html = markdown_to_html_body(digest_text)
    tts_text = prepare_tts_text(body_html)
    tts_text_escaped = (tts_text
        .replace("\\", "\\\\")
        .replace('"', '\\"')
        .replace("\n", " ")
        .replace("`", "'"))
    tag = get_category_tag(subject)
    bg, fg = get_tag_color(tag)
    chars_per_10sec = int(125 * tts_rate)

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>{subject} · {SITE_NAME}</title>
<style>
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{
    font-family: Georgia, 'Times New Roman', serif;
    background: #0a0a0a;
    color: #e8e4d9;
    max-width: 820px;
    margin: 0 auto;
    padding: 40px 32px 100px;
    line-height: 1.75;
    font-size: 16px;
  }}
  .top-bar {{
    position: sticky;
    top: 0;
    background: #0a0a0a;
    z-index: 100;
    display: flex;
    align-items: center;
    justify-content: space-between;
    padding: 12px 0;
    margin-bottom: 24px;
    border-bottom: 1px solid #1a1a1a;
  }}
  .back-link {{
    font-size: 12px;
    color: #555;
    text-decoration: none;
    font-family: 'Courier New', monospace;
    letter-spacing: 0.05em;
  }}
  .back-link:hover {{ color: #c9b97a; }}
  .site-name {{
    font-size: 12px;
    color: #333;
    font-family: 'Courier New', monospace;
    letter-spacing: 0.1em;
  }}
  header {{ margin-bottom: 28px; }}
  .tag {{
    display: inline-block;
    font-size: 11px;
    font-family: 'Courier New', monospace;
    font-weight: bold;
    letter-spacing: 0.1em;
    padding: 3px 8px;
    border-radius: 3px;
    margin-bottom: 10px;
    background: {bg};
    color: {fg};
    border: 1px solid {fg}44;
  }}
  header h1 {{
    font-size: 26px;
    font-weight: normal;
    letter-spacing: 0.02em;
    color: #c9b97a;
    line-height: 1.3;
    margin-bottom: 8px;
  }}
  header .meta {{
    font-size: 13px;
    color: #555;
    font-family: 'Courier New', monospace;
  }}
  .divider {{
    border: none;
    border-top: 1px solid #1a1a1a;
    margin: 20px 0 28px;
  }}
  h2 {{
    font-size: 10px;
    font-weight: normal;
    letter-spacing: 0.15em;
    text-transform: uppercase;
    color: #666;
    margin: 32px 0 10px;
    padding-bottom: 6px;
    border-bottom: 1px solid #1a1a1a;
  }}
  p {{ margin-bottom: 12px; color: #ccc8bc; }}
  strong {{ color: #c9b97a; font-weight: bold; }}
  .ticker-line {{ color: #c9b97a; font-weight: bold; }}
  #tts-bar {{
    position: fixed;
    bottom: 0; left: 0; right: 0;
    background: #111;
    border-top: 1px solid #222;
    padding: 8px 16px;
    display: flex;
    align-items: center;
    gap: 8px;
    font-family: 'Courier New', monospace;
    font-size: 11px;
    flex-wrap: wrap;
  }}
  #tts-bar button {{
    background: #1a1a1a;
    border: 1px solid #333;
    color: #e8e4d9;
    padding: 5px 10px;
    cursor: pointer;
    font-size: 11px;
    border-radius: 3px;
    font-family: 'Courier New', monospace;
    transition: background 0.15s;
    white-space: nowrap;
  }}
  #tts-bar button:hover {{ background: #2a2a2a; }}
  #tts-bar button.active {{ background: #3d3820; border-color: #c9b97a; color: #c9b97a; }}
  #tts-bar button.skip {{ color: #777; }}
  #tts-bar button.skip:hover {{ color: #e8e4d9; }}
  .speed-group {{
    display: flex;
    align-items: center;
    gap: 6px;
    margin-left: 4px;
  }}
  .speed-group label {{
    color: #555;
    font-size: 11px;
    white-space: nowrap;
  }}
  #speed-slider {{
    width: 80px;
    accent-color: #c9b97a;
  }}
  #speed-val {{
    color: #c9b97a;
    font-size: 11px;
    min-width: 28px;
  }}
  #tts-status {{ color: #555; flex: 1; font-size: 11px; min-width: 120px; }}
  .notify-btn {{
    margin-left: auto;
    background: #1a1a2a !important;
    border-color: #334 !important;
    color: #668 !important;
  }}
  .notify-btn.enabled {{ color: #4c8faf !important; border-color: #4c8faf44 !important; background: #1a2a3a !important; }}
</style>
</head>
<body>

<div class="top-bar">
  <a class="back-link" href="index.html">&#8592; All Dispatches</a>
  <span class="site-name">THE VICTORY LANE</span>
</div>

<header>
  <div class="tag">{tag}</div>
  <h1>{subject}</h1>
  <div class="meta">{email_date}</div>
</header>
<hr class="divider">

{body_html}

<div id="tts-bar">
  <button id="btn-play" onclick="togglePlay()">&#9654; Play</button>
  <button class="skip" onclick="skipBack()">&#8592; 10s</button>
  <button class="skip" onclick="skipForward()">10s &#8594;</button>
  <button onclick="stopReading()">&#9632; Stop</button>
  <button onclick="replayReading()">&#8635; Replay</button>
  <div class="speed-group">
    <label>Speed</label>
    <input type="range" id="speed-slider" min="0.5" max="2.0" step="0.1" value="{tts_rate}" oninput="updateSpeed(this.value)">
    <span id="speed-val">{tts_rate}x</span>
  </div>
  <span id="tts-status">Ready &mdash; auto-starting...</span>
  <button class="notify-btn" id="notify-btn" onclick="requestNotifications()">Notify me</button>
</div>

<script>
  const digestText = "{tts_text_escaped}";
  const CHARS_PER_10S = {chars_per_10sec};
  let rate = {tts_rate};
  let charIndex = 0;
  let isPaused = false;
  let utterance = null;

  function getVoice() {{
    const voices = window.speechSynthesis.getVoices();
    return (
      voices.find(v => v.name.includes("Guy Online"))  ||  // Edge: Microsoft Guy Online
      voices.find(v => v.name.includes("Guy"))         ||  // Edge fallback
      voices.find(v => v.name === "Daniel")            ||  // iOS/macOS British male
      voices.find(v => v.name === "Aaron")             ||  // iOS male
      voices.find(v => v.name === "Samantha")          ||  // iOS default female
      voices.find(v => v.name.includes("Male") && v.lang.startsWith("en")) ||
      voices.find(v => v.lang.startsWith("en-US"))     ||
      voices.find(v => v.lang.startsWith("en"))        ||
      voices[0]
    );
  }}

  function updateSpeed(val) {{
    rate = parseFloat(val);
    document.getElementById("speed-val").textContent = rate.toFixed(1) + "x";
    // If currently playing, restart from current position at new speed
    if (window.speechSynthesis.speaking && !isPaused) {{
      speakFrom(charIndex);
    }}
  }}

  function setStatus(msg) {{ document.getElementById("tts-status").textContent = msg; }}

  function setPlayBtn(playing) {{
    const btn = document.getElementById("btn-play");
    if (playing) {{ btn.innerHTML = "&#9646;&#9646; Pause"; btn.classList.add("active"); }}
    else {{ btn.innerHTML = "&#9654; Play"; btn.classList.remove("active"); }}
  }}

  function speakFrom(startChar) {{
    window.speechSynthesis.cancel();
    startChar = Math.max(0, Math.min(startChar, digestText.length - 1));
    const text = startChar > 0 ? digestText.slice(startChar) : digestText;
    utterance = new SpeechSynthesisUtterance(text);
    utterance.rate = rate;
    utterance.pitch = 1.0;
    utterance.volume = 1.0;
    utterance.onboundary = (e) => {{ if (e.name === "word") charIndex = startChar + e.charIndex; }};
    utterance.onend = () => {{
      if (!isPaused) {{ charIndex = 0; isPaused = false; setPlayBtn(false); setStatus("Done \u2014 press Replay to listen again"); }}
    }};
    const doSpeak = () => {{
      const voice = getVoice();
      if (voice) utterance.voice = voice;
      setStatus("Voice: " + (voice ? voice.name : "default"));
      setPlayBtn(true);
      window.speechSynthesis.speak(utterance);
    }};
    const voices = window.speechSynthesis.getVoices();
    if (voices.length === 0) {{ window.speechSynthesis.onvoiceschanged = doSpeak; }} else {{ doSpeak(); }}
  }}

  function togglePlay() {{
    if (isPaused) {{ isPaused = false; speakFrom(charIndex); setStatus("Resumed..."); }}
    else if (window.speechSynthesis.speaking) {{ isPaused = true; window.speechSynthesis.cancel(); setPlayBtn(false); setStatus("Paused \u2014 press Play to resume"); }}
    else {{ charIndex = 0; isPaused = false; speakFrom(0); }}
  }}

  function skipForward() {{
    const wasPlaying = window.speechSynthesis.speaking && !isPaused;
    window.speechSynthesis.cancel();
    charIndex = Math.min(charIndex + CHARS_PER_10S, digestText.length - 1);
    if (wasPlaying) {{ isPaused = false; speakFrom(charIndex); setStatus("Skipped forward 10s..."); }}
    else {{ setStatus("Skipped forward \u2014 press Play to resume"); }}
  }}

  function skipBack() {{
    const wasPlaying = window.speechSynthesis.speaking && !isPaused;
    window.speechSynthesis.cancel();
    charIndex = Math.max(0, charIndex - CHARS_PER_10S);
    if (wasPlaying) {{ isPaused = false; speakFrom(charIndex); setStatus("Skipped back 10s..."); }}
    else {{ setStatus("Skipped back \u2014 press Play to resume"); }}
  }}

  function stopReading() {{
    isPaused = false; charIndex = 0;
    window.speechSynthesis.cancel();
    setPlayBtn(false);
    setStatus("Stopped \u2014 press Replay to start over");
  }}

  function replayReading() {{
    isPaused = false; charIndex = 0;
    window.speechSynthesis.cancel();
    setStatus("Restarting...");
    setTimeout(() => speakFrom(0), 300);
  }}

  function updateNotifyBtn() {{
    const btn = document.getElementById("notify-btn");
    if (Notification.permission === "granted") {{
      btn.textContent = "Notifications on";
      btn.classList.add("enabled");
    }} else if (Notification.permission === "denied") {{
      btn.textContent = "Notifications blocked";
    }} else {{
      btn.textContent = "Notify me";
      btn.classList.remove("enabled");
    }}
  }}

  function requestNotifications() {{
    if (!("Notification" in window)) {{ setStatus("Notifications not supported"); return; }}
    if (Notification.permission === "granted") {{ setStatus("Notifications already enabled!"); return; }}
    Notification.requestPermission().then(permission => {{
      updateNotifyBtn();
      if (permission === "granted") {{
        setStatus("Notifications enabled");
        new Notification("{SITE_NAME}", {{ body: "You'll be notified when new dispatches arrive." }});
      }}
    }});
  }}

  function checkForNewDigest() {{
    fetch("digests.json?t=" + Date.now())
      .then(r => r.json())
      .then(data => {{
        if (data && data.length > 0) {{
          const latest = data[0];
          const lastSeen = localStorage.getItem("lastSeenDigest");
          if (lastSeen !== latest.filename) {{
            localStorage.setItem("lastSeenDigest", latest.filename);
            if (Notification.permission === "granted" && lastSeen !== null) {{
              const n = new Notification("{SITE_NAME}", {{
                body: latest.subject + " \u00b7 " + latest.email_date,
              }});
              n.onclick = () => {{ window.open(latest.filename); }};
            }}
          }}
        }}
      }})
      .catch(() => {{}});
  }}

  window.addEventListener("load", () => {{
    updateNotifyBtn();
    localStorage.setItem("lastSeenDigest", window.location.pathname.split("/").pop());
    checkForNewDigest();
    setInterval(checkForNewDigest, 5 * 60 * 1000);
    setTimeout(() => speakFrom(0), 1800);
  }});
  window.addEventListener("pagehide", () => {{ window.speechSynthesis.cancel(); }});
  window.addEventListener("visibilitychange", () => {{ if (document.hidden) {{ window.speechSynthesis.cancel(); setPlayBtn(false); }} }});
</script>

</body>
</html>"""


def build_index_html(digests):
    cards = ""
    for i, entry in enumerate(digests):
        tag = get_category_tag(entry["subject"])
        bg, fg = get_tag_color(tag)
        latest = f' <span style="font-size:10px; background:#3d3820; color:#c9b97a; border:1px solid #c9b97a44; padding:2px 6px; border-radius:3px; font-family:\'Courier New\',monospace; vertical-align:middle; margin-left:6px;">LATEST</span>' if i == 0 else ""
        preview = entry.get("preview", "")

        cards += f"""
    <a class="card" href="{entry['filename']}">
      <div class="card-tag" style="background:{bg}; color:{fg}; border-color:{fg}44;">{tag}</div>
      <div class="card-meta">{entry['email_date']}</div>
      <div class="card-title">{entry['subject']}{latest}</div>
      {"<div class='card-preview'>" + preview + "</div>" if preview else ""}
    </a>"""

    updated = datetime.now(EASTERN).strftime("%b %d, %Y %I:%M %p ET")

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>{SITE_NAME}</title>
<style>
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{
    font-family: Georgia, 'Times New Roman', serif;
    background: #0a0a0a;
    color: #e8e4d9;
    max-width: 860px;
    margin: 0 auto;
    padding: 40px 32px 60px;
    font-size: 16px;
  }}
  .site-header {{
    position: sticky;
    top: 0;
    background: #0a0a0a;
    z-index: 100;
    display: flex;
    align-items: center;
    justify-content: space-between;
    border-bottom: 1px solid #1a1a1a;
    padding: 16px 0 16px;
    margin-bottom: 32px;
  }}
  .site-title-group {{
    display: flex;
    flex-direction: column;
    gap: 2px;
  }}
  .site-title {{
    font-size: 22px;
    font-weight: normal;
    letter-spacing: 0.02em;
    color: #f0ece0;
    font-family: Georgia, 'Times New Roman', serif;
  }}
  .site-subtitle {{
    font-size: 11px;
    letter-spacing: 0.08em;
    color: #f0ece0;
    font-family: 'Courier New', monospace;
  }}
  .site-updated {{
    font-size: 11px;
    color: #444;
    font-family: 'Courier New', monospace;
  }}
  .card {{
    display: block;
    text-decoration: none;
    color: inherit;
    border-bottom: 1px solid #141414;
    padding: 20px 0;
    transition: padding-left 0.15s;
  }}
  .card:hover {{ padding-left: 8px; }}
  .card:hover .card-title {{ color: #c9b97a; }}
  .card:first-of-type {{ border-top: 1px solid #141414; margin-top: 8px; }}
  .card-tag {{
    display: inline-block;
    font-size: 10px;
    font-family: 'Courier New', monospace;
    font-weight: bold;
    letter-spacing: 0.1em;
    padding: 2px 7px;
    border-radius: 3px;
    border: 1px solid;
    margin-bottom: 6px;
  }}
  .card-meta {{
    font-size: 12px;
    color: #444;
    font-family: 'Courier New', monospace;
    margin-bottom: 6px;
  }}
  .card-title {{
    font-size: 18px;
    color: #d8d4c8;
    line-height: 1.4;
    margin-bottom: 6px;
    transition: color 0.15s;
  }}
  .card-preview {{
    font-size: 14px;
    color: #555;
    line-height: 1.5;
  }}
</style>
</head>
<body>

<div class="site-header">
  <div class="site-title-group">
    <div class="site-title">THE VICTORY LANE</div>
    <div class="site-subtitle">MARKET INTELLIGENTSIA</div>
  </div>
  <div class="site-updated">Updated {updated}</div>
</div>

{cards}

</body>
</html>"""


def save_html_local(html, output_path):
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"  ✓ HTML digest saved: {output_path}")


def save_html_github(html, subject):
    os.makedirs("docs", exist_ok=True)
    slug = re.sub(r"[^a-z0-9]+", "-", subject.lower()).strip("-")
    slug = slug[:80]
    filename = f"{slug}.html"
    filepath = f"docs/{filename}"
    with open(filepath, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"  ✓ HTML digest saved: {filepath}")
    return filename


def update_index(new_entries):
    meta_path = "docs/digests.json"
    existing = []

    if os.path.exists(meta_path):
        with open(meta_path, encoding="utf-8") as f:
            existing = json.load(f)

    existing_filenames = {e["filename"] for e in existing}
    for filename, subject, email_date, preview, sent_ts in new_entries:
        if filename not in existing_filenames:
            existing.append({
                "filename": filename,
                "subject": subject,
                "email_date": email_date,
                "preview": preview,
                "sent_ts": sent_ts
            })

    existing.sort(key=lambda x: x.get("sent_ts", x["filename"]), reverse=True)

    with open(meta_path, "w", encoding="utf-8") as f:
        json.dump(existing, f, indent=2, ensure_ascii=False)

    index_html = build_index_html(existing)
    with open("docs/index.html", "w", encoding="utf-8") as f:
        f.write(index_html)

    print(f"  ✓ Index updated — {len(existing)} dispatch(es) listed")


def launch_edge(html_path, edge_exe):
    if not os.path.exists(edge_exe):
        edge_exe = r"C:\Program Files\Microsoft\Edge\Application\msedge.exe"
    if os.path.exists(edge_exe):
        subprocess.Popen([edge_exe, html_path])
        print(f"  ✓ Launched Edge")
    else:
        os.startfile(html_path)


def run():
    print(f"[{datetime.now(EASTERN).strftime('%H:%M:%S ET')}] The Victory Lane starting...")
    config = CONFIG

    emails, processed_ids = fetch_new_emails(config)

    if not emails:
        print("  No new emails found.")
        return

    print(f"  Found {len(emails)} unprocessed email(s) — processing in chronological order...")

    new_ids = set()
    new_index_entries = []

    for uid, subject, body, email_date, sent_utc in emails:
        print(f"\n  ── Processing: {subject} ({email_date}) ──")
        try:
            print("  Summarizing with Claude...")
            digest = summarize_with_claude(body, config["anthropic_api_key"])
            preview = get_preview(digest)
            html = build_html(digest, subject, email_date, config["tts_rate"])

            if config["github_actions"]:
                filename = save_html_github(html, subject)
                new_index_entries.append((filename, subject, email_date, preview, sent_utc.isoformat()))
            else:
                save_html_local(html, config["html_output"])
                launch_edge(config["html_output"], config["edge_exe"])

            new_ids.add(uid)

        except Exception as e:
            print(f"  ✗ Error processing '{subject}': {e}")
            raise

    if config["github_actions"]:
        update_index(new_index_entries)
        save_processed_ids_github(processed_ids | new_ids)
    else:
        save_processed_ids(config["state_file"], processed_ids | new_ids)

    print(f"\n  Done. Processed {len(new_ids)} dispatch(es).")


if __name__ == "__main__":
    run()
