"""
Vital Knowledge Newsletter → Discord Digest + Edge/Guy Readout
---------------------------------------------------------------
Polls Yahoo Mail via IMAP for newsletters from Vital Knowledge,
summarizes them with Claude in Adam's voice, saves a styled HTML
digest, launches Edge to read it aloud via Guy, and posts to Discord
with the HTML file attached.

Processes ALL unprocessed emails from the last N hours in one run.

Folder: C:\\Tools\\VitalRecaps\\
Setup:
  pip install anthropic requests
  Then configure the CONFIG block below.

TO RUN:
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
# CONFIG — fill these in before running
# ─────────────────────────────────────────────
CONFIG = {
    # Yahoo Mail credentials
    "yahoo_email":        "jeremyquinlan@yahoo.com",
    "yahoo_app_password": "mnvpshecwpfckeeh",

    # Sender filter — partial match on From address
    "sender_filter":      "vitalknowledge",

    # Anthropic API key
    "anthropic_api_key":  "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",

    # Discord webhook URL
    "discord_webhook_url": "YOUR_DISCORD_WEBHOOK_URL",

    # How far back to look for emails (hours)
    "lookback_hours": 6,

    # Output paths
    "html_output":  r"C:\Tools\VitalRecaps\digest.html",
    "state_file":   r"C:\Tools\VitalRecaps\processed_ids.json",

    # Edge browser executable
    "edge_exe": r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",

    # TTS reading speed (0.1 slow – 2.0 fast, 1.0 = normal)
    "tts_rate": 1.0,
}
# ─────────────────────────────────────────────


SUMMARIZE_PROMPT = """You are summarizing a Vital Knowledge financial newsletter written by Adam Crisafulli.
Your job is to produce a concise but substantive digest that captures Adam's voice — direct, confident,
uses parenthetical asides, cites sources inline (NYT, WSJ, FT, Bloomberg, etc.), and isn't shy about
giving a market view. Write in flowing prose for the narrative sections, not just bullet fragments.

Structure your response in exactly this order using these exact section headers:

## MARKETS
Report SPX and Nasdaq as PERCENTAGE CHANGE only (e.g. -0.26%, +0.54%).
Do NOT use basis points or point values for SPX and Nasdaq.
Also include Dow, R2K, Brent, Gold, Silver, BTC, DXY if present — these can use their native units.

## RATES & FED
Treasury move and Fed expectations. Keep to 2-3 sentences in Adam's voice.
When referencing basis point moves on yields, write "basis points" in full, never "bp".

## MARKET OUTLOOK
Adam's editorial view on where markets are headed. 3-5 sentences, confident tone.

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
- Write in Adam Crisafulli's voice throughout
- Consolidate any repeated news items — mention each story only once
- Include specific numbers and source attributions
- Always write "basis points" in full, never abbreviate as "bp"
- Keep total length under 4000 characters for the TTS readout
- Do not include scheduling notes, subscription info, or technical notices"""


def load_processed_ids(state_file):
    if os.path.exists(state_file):
        with open(state_file) as f:
            return set(json.load(f))
    return set()


def save_processed_ids(state_file, ids):
    os.makedirs(os.path.dirname(state_file), exist_ok=True)
    with open(state_file, "w") as f:
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
    """Extract sent date from email headers and convert to Eastern time."""
    date_str = msg.get("Date", "")
    try:
        dt = parsedate_to_datetime(date_str)
        # Convert to Eastern time regardless of what tz the header is in
        dt_eastern = dt.astimezone(EASTERN)
        return dt_eastern.strftime("%A %B %d, %Y · %I:%M %p ET")
    except Exception:
        return datetime.now(EASTERN).strftime("%A %B %d, %Y · %I:%M %p ET")


def get_email_sent_utc(msg):
    """Return the sent time as UTC-aware datetime for age filtering."""
    date_str = msg.get("Date", "")
    try:
        dt = parsedate_to_datetime(date_str)
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=timezone.utc)
        return dt.astimezone(timezone.utc)
    except Exception:
        return datetime.now(timezone.utc)


def fetch_new_emails(config):
    """Fetch all unprocessed Vital Knowledge emails within lookback window."""
    mail = imaplib.IMAP4_SSL("imap.mail.yahoo.com", 993)
    mail.login(config["yahoo_email"], config["yahoo_app_password"])
    mail.select("inbox")

    # IMAP SINCE is date-only so we cast a wider net and filter precisely below
    since_date = (datetime.utcnow() - timedelta(hours=config["lookback_hours"] + 24)).strftime("%d-%b-%Y")
    status, data = mail.search(None, f'(SINCE "{since_date}")')

    uids = data[0].split()
    results = []
    processed = load_processed_ids(config["state_file"])
    cutoff = datetime.now(timezone.utc) - timedelta(hours=config["lookback_hours"])

    for uid in uids:
        uid_str = uid.decode()

        # Skip already processed
        if uid_str in processed:
            print(f"  Skipping already processed UID {uid_str}")
            continue

        status, msg_data = mail.fetch(uid, "(RFC822)")
        raw = msg_data[0][1]
        msg = email.message_from_bytes(raw)

        # Filter by sender
        from_addr = decode_str(msg.get("From", ""))
        if config["sender_filter"].lower() not in from_addr.lower():
            continue

        # Filter by actual sent time — must be within lookback window
        sent_utc = get_email_sent_utc(msg)
        if sent_utc < cutoff:
            print(f"  Skipping old email (sent {sent_utc.strftime('%Y-%m-%d %H:%M UTC')})")
            continue

        subject = decode_str(msg.get("Subject", "(No Subject)"))
        clean_subject = re.sub(r"^Vital Knowledge:\s*", "", subject).strip()
        body = get_email_body(msg)
        email_date = get_email_date(msg)

        if body.strip():
            results.append((uid_str, clean_subject, body, email_date, sent_utc))
            print(f"  Found: {clean_subject} ({email_date})")

    mail.logout()

    # Sort oldest to newest so they play in chronological order
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


def markdown_to_html_body(text):
    """Convert simple markdown digest to styled HTML paragraphs."""
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
    """Strip HTML and expand abbreviations for clean TTS reading."""
    text = re.sub(r"<[^>]+>", " ", html_body)
    text = re.sub(r"\s+", " ", text).strip()
    text = re.sub(r"\bbp\b", "basis points", text)
    text = re.sub(r"\bBP\b", "basis points", text)
    text = re.sub(r"\bpct\b", "percent", text, flags=re.IGNORECASE)
    text = re.sub(r"\bYoY\b", "year over year", text, flags=re.IGNORECASE)
    text = re.sub(r"\bQoQ\b", "quarter over quarter", text, flags=re.IGNORECASE)
    text = re.sub(r"\bY/Y\b", "year over year", text)
    text = re.sub(r"\bQ/Q\b", "quarter over quarter", text)
    text = re.sub(r"\bEPS\b", "earnings per share", text)
    text = re.sub(r"\bET\b", "Eastern time", text)
    return text


def save_html_digest(digest_text, subject, email_date, output_path, tts_rate):
    """Save styled HTML file with embedded Guy auto-readout."""
    body_html = markdown_to_html_body(digest_text)
    tts_text = prepare_tts_text(body_html)
    tts_text_escaped = (tts_text
        .replace("\\", "\\\\")
        .replace('"', '\\"')
        .replace("\n", " ")
        .replace("`", "'"))

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>{subject}</title>
<style>
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{
    font-family: Georgia, 'Times New Roman', serif;
    background: #0f0f0f;
    color: #e8e4d9;
    max-width: 820px;
    margin: 0 auto;
    padding: 40px 32px 80px;
    line-height: 1.75;
    font-size: 16px;
  }}
  header {{
    border-bottom: 1px solid #333;
    padding-bottom: 16px;
    margin-bottom: 28px;
  }}
  header h1 {{
    font-size: 22px;
    font-weight: normal;
    letter-spacing: 0.04em;
    color: #f0ece0;
  }}
  header .meta {{
    font-size: 13px;
    color: #666;
    margin-top: 6px;
    font-family: 'Courier New', monospace;
  }}
  header .meta .label {{
    color: #444;
    margin-right: 4px;
  }}
  h2 {{
    font-size: 11px;
    font-weight: normal;
    letter-spacing: 0.12em;
    text-transform: uppercase;
    color: #888;
    margin: 32px 0 10px;
    padding-bottom: 6px;
    border-bottom: 1px solid #222;
  }}
  p {{
    margin-bottom: 12px;
    color: #d8d4c8;
  }}
  strong {{
    color: #f0ece0;
    font-weight: bold;
  }}
  .ticker-line {{
    color: #c9b97a;
    font-weight: bold;
  }}
  #tts-bar {{
    position: fixed;
    bottom: 0; left: 0; right: 0;
    background: #1a1a1a;
    border-top: 1px solid #333;
    padding: 12px 24px;
    display: flex;
    align-items: center;
    gap: 16px;
    font-family: 'Courier New', monospace;
    font-size: 13px;
  }}
  #tts-bar button {{
    background: #2a2a2a;
    border: 1px solid #444;
    color: #e8e4d9;
    padding: 6px 16px;
    cursor: pointer;
    font-size: 13px;
    border-radius: 3px;
    font-family: 'Courier New', monospace;
    transition: background 0.15s;
  }}
  #tts-bar button:hover {{ background: #3a3a3a; }}
  #tts-bar button.active {{ background: #3d3820; border-color: #c9b97a; color: #c9b97a; }}
  #tts-status {{ color: #888; flex: 1; }}
</style>
</head>
<body>

<header>
  <h1>{subject}</h1>
  <div class="meta">
    <span class="label">Received:</span>{email_date}
  </div>
</header>

{body_html}

<div id="tts-bar">
  <button id="btn-play" onclick="togglePlay()">&#9654; Play</button>
  <button id="btn-stop" onclick="stopReading()">&#9632; Stop</button>
  <button id="btn-replay" onclick="replayReading()">&#8635; Replay</button>
  <span id="tts-status">Ready &mdash; auto-starting in a moment...</span>
</div>

<script>
  const digestText = "{tts_text_escaped}";
  const rate = {tts_rate};

  let charIndex = 0;
  let isPaused = false;
  let utterance = null;

  function getVoice() {{
    const voices = window.speechSynthesis.getVoices();
    return (
      voices.find(v => v.name.includes("Guy")) ||
      voices.find(v => v.name.includes("Male") && v.lang.startsWith("en")) ||
      voices.find(v => v.lang.startsWith("en-US")) ||
      voices[0]
    );
  }}

  function setStatus(msg) {{
    document.getElementById("tts-status").textContent = msg;
  }}

  function setPlayBtn(playing) {{
    const btn = document.getElementById("btn-play");
    if (playing) {{
      btn.innerHTML = "&#9646;&#9646; Pause";
      btn.classList.add("active");
    }} else {{
      btn.innerHTML = "&#9654; Play";
      btn.classList.remove("active");
    }}
  }}

  function speakFrom(startChar) {{
    window.speechSynthesis.cancel();
    const text = startChar > 0 ? digestText.slice(startChar) : digestText;
    utterance = new SpeechSynthesisUtterance(text);
    utterance.rate = rate;
    utterance.pitch = 1.0;
    utterance.volume = 1.0;

    utterance.onboundary = (e) => {{
      if (e.name === "word") {{
        charIndex = startChar + e.charIndex;
      }}
    }};

    utterance.onend = () => {{
      if (!isPaused) {{
        charIndex = 0;
        isPaused = false;
        setPlayBtn(false);
        setStatus("Done \u2014 press Replay to listen again");
      }}
    }};

    const doSpeak = () => {{
      const voice = getVoice();
      if (voice) utterance.voice = voice;
      setStatus("Reading with: " + (voice ? voice.name : "default voice"));
      setPlayBtn(true);
      window.speechSynthesis.speak(utterance);
    }};

    const voices = window.speechSynthesis.getVoices();
    if (voices.length === 0) {{
      window.speechSynthesis.onvoiceschanged = doSpeak;
    }} else {{
      doSpeak();
    }}
  }}

  function togglePlay() {{
    if (isPaused) {{
      isPaused = false;
      speakFrom(charIndex);
      setStatus("Resumed...");
    }} else if (window.speechSynthesis.speaking) {{
      isPaused = true;
      window.speechSynthesis.cancel();
      setPlayBtn(false);
      setStatus("Paused \u2014 press Play to resume");
    }} else {{
      charIndex = 0;
      isPaused = false;
      speakFrom(0);
    }}
  }}

  function stopReading() {{
    isPaused = false;
    charIndex = 0;
    window.speechSynthesis.cancel();
    setPlayBtn(false);
    setStatus("Stopped \u2014 press Replay to start over");
  }}

  function replayReading() {{
    isPaused = false;
    charIndex = 0;
    window.speechSynthesis.cancel();
    setStatus("Restarting...");
    setTimeout(() => speakFrom(0), 300);
  }}

  window.addEventListener("load", () => {{
    setTimeout(() => speakFrom(0), 1800);
  }});
</script>

</body>
</html>"""

    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"  ✓ HTML digest saved: {output_path}")


def launch_edge(html_path, edge_exe):
    """Open the digest HTML in Edge."""
    if not os.path.exists(edge_exe):
        edge_exe = r"C:\Program Files\Microsoft\Edge\Application\msedge.exe"
    if os.path.exists(edge_exe):
        subprocess.Popen([edge_exe, html_path])
        print(f"  ✓ Launched Edge for Guy readout")
    else:
        os.startfile(html_path)
        print(f"  ✓ Opened digest in default browser")


def post_to_discord(webhook_url, subject, email_date, html_path):
    """Post short message to Discord with HTML file attached."""
    content = (
        f"📬 **{subject}**\n"
        f"🕐 {email_date}\n"
        f"🔊 Full audio digest attached — open in Edge"
    )
    with open(html_path, "rb") as f:
        resp = requests.post(
            webhook_url,
            data={"content": content},
            files={"file": ("digest.html", f, "text/html")}
        )
    resp.raise_for_status()
    print(f"  ✓ Posted to Discord with HTML attachment")


def run():
    print(f"[{datetime.now(EASTERN).strftime('%H:%M:%S ET')}] Vital Knowledge Digest starting...")
    config = CONFIG

    emails, processed_ids = fetch_new_emails(config)

    if not emails:
        print("  No new Vital Knowledge emails found.")
        return

    print(f"  Found {len(emails)} unprocessed email(s) — processing in chronological order...")

    new_ids = set()
    for uid, subject, body, email_date, sent_utc in emails:
        print(f"\n  ── Processing: {subject} ({email_date}) ──")
        try:
            print("  Summarizing with Claude...")
            digest = summarize_with_claude(body, config["anthropic_api_key"])

            print("  Saving HTML digest...")
            save_html_digest(digest, subject, email_date, config["html_output"], config["tts_rate"])

            print("  Launching Edge with Guy readout...")
            launch_edge(config["html_output"], config["edge_exe"])

            print("  Posting to Discord...")
            post_to_discord(
                config["discord_webhook_url"],
                subject,
                email_date,
                config["html_output"]
            )

            new_ids.add(uid)

        except Exception as e:
            print(f"  ✗ Error processing '{subject}': {e}")
            raise

    save_processed_ids(config["state_file"], processed_ids | new_ids)
    print(f"\n  Done. Processed {len(new_ids)} email(s).")


if __name__ == "__main__":
    run()
