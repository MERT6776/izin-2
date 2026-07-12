from flask import Flask, Response, jsonify, request
import base64
import glob
import hmac
import math
import os
import re
import time
import unicodedata
from datetime import datetime
from threading import Lock

import pandas as pd

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 1024 * 1024

WHATSAPP_NUMBER = "905459157444"
DEFAULT_EXCEL_FILE = "BURHAN BİLİKTÜ İZİN.xlsx"

# Basit giriş denemesi koruması: 10 dakika içinde 6 hatalı denemeden sonra 10 dakika bekletir.
FAILED_LOGINS = {}
FAILED_LOCK = Lock()
MAX_FAILED_ATTEMPTS = 6
ATTEMPT_WINDOW_SECONDS = 10 * 60
BLOCK_SECONDS = 10 * 60

ICON_192_B64 = "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"
ICON_512_B64 = "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"


def normalize_text(value):
    text = str(value or "").strip().upper()
    replacements = str.maketrans({
        "Ç": "C", "Ğ": "G", "İ": "I", "I": "I",
        "Ö": "O", "Ş": "S", "Ü": "U",
    })
    text = text.translate(replacements)
    text = unicodedata.normalize("NFKD", text)
    return re.sub(r"[^A-Z0-9]", "", text)


def clean_scalar(value):
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass

    if isinstance(value, float) and math.isfinite(value) and value.is_integer():
        return str(int(value))

    text = str(value).strip()
    if re.fullmatch(r"-?\d+\.0", text):
        return text[:-2]
    return text


def as_number(value):
    if value is None:
        return 0.0
    try:
        if pd.isna(value):
            return 0.0
    except Exception:
        pass

    text = str(value).strip().replace(" ", "")
    if not text:
        return 0.0

    # 1.234,5 ve 29,5 gibi Türkçe sayı biçimlerini destekler.
    if "," in text and "." in text:
        text = text.replace(".", "").replace(",", ".")
    else:
        text = text.replace(",", ".")

    try:
        number = float(text)
        return number if math.isfinite(number) else 0.0
    except (TypeError, ValueError):
        return 0.0


def same_identifier(left, right):
    left_clean = clean_scalar(left)
    right_clean = clean_scalar(right)
    if hmac.compare_digest(left_clean, right_clean):
        return True

    # Excel kullanıcı adını sayı olarak kaydetmişse "00123" / "123" eşleşmesine izin verir.
    if left_clean.isdigit() and right_clean.isdigit():
        return int(left_clean) == int(right_clean)
    return False


def get_excel_path():
    configured = os.getenv("EXCEL_FILE", DEFAULT_EXCEL_FILE).strip()
    if configured and os.path.exists(configured):
        return configured

    candidates = [
        item for item in glob.glob("*.xlsx")
        if not os.path.basename(item).startswith("~$")
    ]
    if not candidates:
        raise FileNotFoundError(
            f"Excel bulunamadı. '{DEFAULT_EXCEL_FILE}' dosyasını app.py ile aynı klasöre koyun."
        )

    preferred = [
        item for item in candidates
        if "BURHAN" in normalize_text(os.path.basename(item))
        and "IZIN" in normalize_text(os.path.basename(item))
    ]
    return preferred[0] if preferred else candidates[0]


def find_column(columns, *aliases):
    lookup = {normalize_text(column): column for column in columns}
    for alias in aliases:
        normalized = normalize_text(alias)
        if normalized in lookup:
            return lookup[normalized]
    return None


def format_file_update_time(path):
    timestamp = os.path.getmtime(path)
    return datetime.fromtimestamp(timestamp).strftime("%d.%m.%Y %H:%M")


def initials_from_name(name):
    parts = [part for part in str(name).split() if part]
    if not parts:
        return "P"
    return "".join(part[0] for part in parts[:2]).upper()


def get_user_data(username, password):
    excel_path = get_excel_path()
    dataframe = pd.read_excel(excel_path, sheet_name=0, dtype=object)
    dataframe.columns = [str(column).strip() for column in dataframe.columns]

    username_col = find_column(dataframe.columns, "KULLANICI ADI", "KULLANICIADI", "USERNAME")
    password_col = find_column(dataframe.columns, "ŞİFRE", "SIFRE", "PASSWORD")
    name_col = find_column(dataframe.columns, "ADI SOYADI", "AD SOYAD", "ADISOYADI", "PERSONEL")
    role_col = find_column(dataframe.columns, "GÖREVİ", "GOREVI", "GÖREV", "UNVAN")
    sunday_col = find_column(dataframe.columns, "PAZAR İZİNLERİ", "PAZAR IZINLERI", "PAZAR İZNİ")
    holiday_col = find_column(dataframe.columns, "RESMİ TATİL", "RESMI TATIL", "RESMİ TATİL İZNİ")
    remaining_col = find_column(
        dataframe.columns,
        "KALAN İZİN HAKKI",
        "KALAN IZIN HAKKI",
        "KALAN İZİN",
        "KALANIZINHAKKI",
    )
    updated_col = find_column(
        dataframe.columns,
        "GÜNCELLEME TARİHİ",
        "GUNCELLEME TARIHI",
        "SON GÜNCELLEME",
    )

    missing = []
    if username_col is None:
        missing.append("KULLANICI ADI")
    if password_col is None:
        missing.append("ŞİFRE")
    if name_col is None:
        missing.append("ADI SOYADI")
    if remaining_col is None:
        missing.append("KALAN İZİN HAKKI")
    if missing:
        raise ValueError("Excel'de gerekli sütunlar bulunamadı: " + ", ".join(missing))

    matching_row = None
    for _, row in dataframe.iterrows():
        if same_identifier(row.get(username_col), username):
            matching_row = row
            break

    if matching_row is None:
        return None

    stored_password = clean_scalar(matching_row.get(password_col))
    entered_password = clean_scalar(password)

    password_matches = hmac.compare_digest(stored_password, entered_password)
    if not password_matches and stored_password.isdigit() and entered_password.isdigit():
        password_matches = int(stored_password) == int(entered_password)

    if not password_matches:
        return None

    name = clean_scalar(matching_row.get(name_col)) or "Personel"
    role = clean_scalar(matching_row.get(role_col)) if role_col else ""
    update_value = clean_scalar(matching_row.get(updated_col)) if updated_col else ""

    return {
        "name": name,
        "role": role or "Personel",
        "username": clean_scalar(matching_row.get(username_col)),
        "sunday_leave": as_number(matching_row.get(sunday_col)) if sunday_col else 0.0,
        "official_holiday": as_number(matching_row.get(holiday_col)) if holiday_col else 0.0,
        "remaining_leave": as_number(matching_row.get(remaining_col)),
        "updated_at": update_value or format_file_update_time(excel_path),
        "initials": initials_from_name(name),
    }


def client_ip():
    forwarded = request.headers.get("X-Forwarded-For", "")
    if forwarded:
        return forwarded.split(",")[0].strip()
    return request.remote_addr or "unknown"


def login_is_blocked(ip_address):
    now = time.time()
    with FAILED_LOCK:
        record = FAILED_LOGINS.get(ip_address)
        if not record:
            return False, 0

        failures = [stamp for stamp in record["failures"] if now - stamp <= ATTEMPT_WINDOW_SECONDS]
        blocked_until = record.get("blocked_until", 0)

        if blocked_until > now:
            return True, int(math.ceil(blocked_until - now))

        if len(failures) >= MAX_FAILED_ATTEMPTS:
            blocked_until = now + BLOCK_SECONDS
            FAILED_LOGINS[ip_address] = {
                "failures": failures,
                "blocked_until": blocked_until,
            }
            return True, BLOCK_SECONDS

        if failures:
            FAILED_LOGINS[ip_address] = {
                "failures": failures,
                "blocked_until": 0,
            }
        else:
            FAILED_LOGINS.pop(ip_address, None)

    return False, 0


def record_failed_login(ip_address):
    now = time.time()
    with FAILED_LOCK:
        record = FAILED_LOGINS.get(ip_address, {"failures": [], "blocked_until": 0})
        failures = [stamp for stamp in record["failures"] if now - stamp <= ATTEMPT_WINDOW_SECONDS]
        failures.append(now)
        FAILED_LOGINS[ip_address] = {
            "failures": failures,
            "blocked_until": record.get("blocked_until", 0),
        }


def clear_failed_logins(ip_address):
    with FAILED_LOCK:
        FAILED_LOGINS.pop(ip_address, None)


HTML_SAYFASI = r'''<!DOCTYPE html>
<html lang="tr">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
    <meta name="theme-color" content="#071b30">
    <meta name="description" content="Personel izin hakları görüntüleme ve izin talep sistemi">
    <meta name="apple-mobile-web-app-capable" content="yes">
    <meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">
    <meta name="apple-mobile-web-app-title" content="İzin Portalı">
    <link rel="manifest" href="/manifest.webmanifest">
    <link rel="apple-touch-icon" href="/icon-192.png">
    <link rel="icon" href="/icon-192.png">
    <title>Personel İzin Portalı</title>

    <!-- Gerçek QR üretimi için. Yüklenemezse sistem kendi QR-benzeri yedek görselini çizer. -->
    <script src="https://cdn.jsdelivr.net/npm/qrious@4.0.2/dist/qrious.min.js" defer></script>

    <style>
        :root {
            --bg-1: #061426;
            --bg-2: #0b3152;
            --panel: rgba(9, 35, 59, 0.76);
            --panel-strong: rgba(6, 26, 46, 0.94);
            --text: #f4fbff;
            --muted: #a8c4d8;
            --primary: #55d9ff;
            --primary-2: #1489c9;
            --success: #2bd881;
            --warning: #ffca57;
            --danger: #ff6678;
            --line: rgba(143, 222, 255, 0.22);
            --shadow: 0 28px 70px rgba(0, 0, 0, 0.34);
            --radius: 24px;
        }

        body.light {
            --bg-1: #dff4ff;
            --bg-2: #eff9ff;
            --panel: rgba(255, 255, 255, 0.78);
            --panel-strong: rgba(255, 255, 255, 0.96);
            --text: #10283d;
            --muted: #527089;
            --primary: #087fc4;
            --primary-2: #23b5e8;
            --line: rgba(14, 116, 174, 0.16);
            --shadow: 0 28px 70px rgba(39, 96, 131, 0.18);
        }

        * { box-sizing: border-box; }
        html { min-height: 100%; background: var(--bg-1); }
        body {
            margin: 0;
            min-height: 100vh;
            min-height: 100dvh;
            color: var(--text);
            font-family: Inter, "Segoe UI", system-ui, -apple-system, BlinkMacSystemFont, sans-serif;
            background:
                radial-gradient(circle at 12% 12%, rgba(63, 196, 255, 0.18), transparent 34%),
                radial-gradient(circle at 88% 80%, rgba(52, 224, 174, 0.11), transparent 28%),
                linear-gradient(145deg, var(--bg-1), var(--bg-2));
            overflow-x: hidden;
            transition: background .35s ease, color .35s ease;
        }

        body::before,
        body::after {
            content: "";
            position: fixed;
            width: 360px;
            height: 360px;
            border-radius: 50%;
            filter: blur(70px);
            opacity: .14;
            pointer-events: none;
            animation: floatBlob 12s ease-in-out infinite;
        }
        body::before { left: -120px; top: -100px; background: #34c9ff; }
        body::after { right: -130px; bottom: -120px; background: #28d99b; animation-delay: -6s; }
        @keyframes floatBlob {
            0%, 100% { transform: translate3d(0, 0, 0) scale(1); }
            50% { transform: translate3d(35px, 24px, 0) scale(1.12); }
        }

        button, input, select, textarea { font: inherit; }
        button, a { -webkit-tap-highlight-color: transparent; }
        button { color: inherit; }
        [hidden] { display: none !important; }

        .topbar {
            position: fixed;
            z-index: 60;
            inset: max(12px, env(safe-area-inset-top)) 14px auto 14px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            pointer-events: none;
        }
        .brand-mini, .top-actions {
            pointer-events: auto;
            display: flex;
            align-items: center;
            gap: 8px;
        }
        .brand-mini {
            padding: 9px 13px;
            border: 1px solid var(--line);
            border-radius: 16px;
            background: var(--panel);
            backdrop-filter: blur(18px);
            box-shadow: 0 12px 30px rgba(0,0,0,.14);
            font-weight: 800;
            letter-spacing: .03em;
        }
        .brand-mark {
            width: 28px;
            height: 28px;
            border-radius: 9px;
            display: grid;
            place-items: center;
            background: linear-gradient(145deg, var(--primary), var(--primary-2));
            color: #042039;
            font-weight: 1000;
        }

        .icon-btn, .install-btn {
            border: 1px solid var(--line);
            background: var(--panel);
            backdrop-filter: blur(18px);
            border-radius: 14px;
            min-height: 42px;
            padding: 0 12px;
            cursor: pointer;
            box-shadow: 0 12px 30px rgba(0,0,0,.12);
            transition: transform .2s ease, border-color .2s ease, background .2s ease;
        }
        .icon-btn:hover, .install-btn:hover { transform: translateY(-2px); border-color: var(--primary); }
        .install-btn { display: none; font-weight: 800; gap: 7px; align-items: center; }
        .install-btn.show { display: inline-flex; }

        .page-shell {
            min-height: 100vh;
            min-height: 100dvh;
            display: grid;
            place-items: center;
            padding: 86px 18px 28px;
            position: relative;
            z-index: 2;
        }

        .glass {
            border: 1px solid var(--line);
            background: var(--panel);
            backdrop-filter: blur(22px);
            -webkit-backdrop-filter: blur(22px);
            box-shadow: var(--shadow);
            border-radius: var(--radius);
        }

        .login-card {
            width: min(440px, 100%);
            padding: clamp(24px, 6vw, 40px);
            position: relative;
            overflow: hidden;
            animation: enterCard .75s cubic-bezier(.2,.8,.2,1) both;
        }
        .login-card::before {
            content: "";
            position: absolute;
            inset: -2px;
            background: linear-gradient(125deg, transparent 15%, rgba(82, 219, 255, .18), transparent 58%);
            transform: translateX(-100%);
            animation: sweep 5.5s ease-in-out infinite;
            pointer-events: none;
        }
        @keyframes enterCard {
            from { opacity: 0; transform: translateY(22px) scale(.97); }
            to { opacity: 1; transform: translateY(0) scale(1); }
        }
        @keyframes sweep {
            0%, 28% { transform: translateX(-100%); }
            65%, 100% { transform: translateX(100%); }
        }

        .login-logo {
            width: 78px;
            height: 78px;
            margin: 0 auto 18px;
            border-radius: 24px;
            display: grid;
            place-items: center;
            background: linear-gradient(145deg, rgba(76,216,255,.22), rgba(18,103,164,.42));
            border: 1px solid rgba(111,225,255,.4);
            box-shadow: inset 0 0 30px rgba(79,218,255,.12), 0 14px 34px rgba(0,0,0,.18);
        }
        .elevator-icon {
            width: 42px; height: 48px;
            border: 2px solid var(--primary);
            border-radius: 7px;
            position: relative;
            overflow: hidden;
        }
        .elevator-icon::before, .elevator-icon::after {
            content: "";
            position: absolute;
            top: 0; bottom: 0; width: 50%;
            background: rgba(82, 217, 255, .14);
        }
        .elevator-icon::before { left: 0; border-right: 1px solid var(--primary); }
        .elevator-icon::after { right: 0; }
        .login-card h1 {
            text-align: center;
            margin: 0;
            font-size: clamp(1.55rem, 5vw, 2.05rem);
            letter-spacing: -.03em;
        }
        .login-subtitle {
            text-align: center;
            color: var(--muted);
            margin: 9px 0 28px;
            line-height: 1.5;
        }

        .field { margin-bottom: 16px; }
        .field label {
            display: block;
            margin-bottom: 8px;
            color: var(--muted);
            font-size: .9rem;
            font-weight: 700;
        }
        .input-wrap { position: relative; }
        .input-wrap input, .modal input, .modal select, .modal textarea {
            width: 100%;
            color: var(--text);
            border: 1px solid var(--line);
            background: rgba(255,255,255,.055);
            border-radius: 15px;
            padding: 15px 46px 15px 15px;
            outline: none;
            transition: border-color .2s ease, box-shadow .2s ease, background .2s ease;
        }
        body.light .input-wrap input,
        body.light .modal input,
        body.light .modal select,
        body.light .modal textarea { background: rgba(8,70,110,.04); }
        input::placeholder, textarea::placeholder { color: color-mix(in srgb, var(--muted) 75%, transparent); }
        .input-wrap input:focus, .modal input:focus, .modal select:focus, .modal textarea:focus {
            border-color: var(--primary);
            box-shadow: 0 0 0 4px rgba(75, 207, 255, .13);
            background: rgba(255,255,255,.08);
        }
        .field-icon {
            position: absolute;
            right: 14px;
            top: 50%;
            transform: translateY(-50%);
            color: var(--muted);
            pointer-events: none;
        }
        .password-toggle {
            pointer-events: auto;
            border: 0;
            background: transparent;
            padding: 4px;
            cursor: pointer;
        }

        .primary-btn, .action-btn {
            width: 100%;
            border: 0;
            border-radius: 15px;
            padding: 15px 18px;
            font-weight: 900;
            cursor: pointer;
            display: inline-flex;
            justify-content: center;
            align-items: center;
            gap: 9px;
            text-decoration: none;
            transition: transform .2s ease, box-shadow .2s ease, filter .2s ease;
        }
        .primary-btn {
            background: linear-gradient(135deg, var(--primary), var(--primary-2));
            color: #032039;
            box-shadow: 0 14px 30px rgba(34, 172, 229, .28);
        }
        .primary-btn:hover, .action-btn:hover { transform: translateY(-2px); filter: brightness(1.05); }
        .primary-btn:active, .action-btn:active { transform: translateY(1px) scale(.99); }
        .primary-btn:disabled { opacity: .65; cursor: wait; }

        .spinner {
            width: 19px;
            height: 19px;
            border-radius: 50%;
            border: 2px solid rgba(3,32,57,.25);
            border-top-color: #032039;
            animation: spin .8s linear infinite;
            display: none;
        }
        .loading .spinner { display: block; }
        @keyframes spin { to { transform: rotate(360deg); } }

        .login-links {
            display: flex;
            justify-content: center;
            margin-top: 17px;
        }
        .text-btn {
            border: 0;
            background: transparent;
            color: var(--primary);
            cursor: pointer;
            font-weight: 800;
            padding: 7px;
        }

        /* ASANSÖR: tamamen birinci şahıs bakış */
        .elevator-scene {
            position: fixed;
            z-index: 100;
            inset: 0;
            overflow: hidden;
            background: #030b13;
            color: #eaf9ff;
            perspective: 1100px;
            --ride-shake: 0px;
        }
        .elevator-scene::after {
            content: "";
            position: absolute;
            inset: 0;
            pointer-events: none;
            background: radial-gradient(circle at center, transparent 42%, rgba(0,0,0,.52) 100%);
            z-index: 15;
        }
        .cabin {
            position: absolute;
            inset: 0;
            transform-origin: center center;
            animation: cabinIdle 1.8s ease-in-out infinite;
            background:
                linear-gradient(90deg, rgba(255,255,255,.04), transparent 10% 90%, rgba(255,255,255,.04)),
                linear-gradient(#0a1c2c, #07131f);
        }
        .riding .cabin { animation: rideVibration .52s ease-in-out infinite; }
        @keyframes cabinIdle {
            0%,100% { transform: translateY(0); }
            50% { transform: translateY(1px); }
        }
        @keyframes rideVibration {
            0%,100% { transform: translate3d(0,0,0); }
            25% { transform: translate3d(.5px,-.6px,0); }
            50% { transform: translate3d(-.4px,.5px,0); }
            75% { transform: translate3d(.3px,-.2px,0); }
        }
        .cabin-ceiling {
            position: absolute;
            inset: 0 0 auto;
            height: 22%;
            background: linear-gradient(#101f2c, #09131d);
            clip-path: polygon(0 0, 100% 0, 82% 100%, 18% 100%);
            box-shadow: inset 0 -20px 30px rgba(0,0,0,.4);
        }
        .ceiling-light {
            position: absolute;
            left: 50%; top: 5%;
            transform: translateX(-50%);
            width: min(55vw, 500px); height: 18px;
            border-radius: 50%;
            background: #d7f7ff;
            box-shadow: 0 0 25px #7ddfff, 0 0 70px rgba(75,209,255,.42);
        }
        .side-wall {
            position: absolute;
            top: 12%; bottom: 0; width: 20%;
            background: linear-gradient(90deg, #091827, #132a3b);
            border: 1px solid rgba(141,217,255,.12);
        }
        .side-wall.left { left: 0; clip-path: polygon(0 0,100% 12%,100% 100%,0 100%); }
        .side-wall.right { right: 0; transform: scaleX(-1); clip-path: polygon(0 0,100% 12%,100% 100%,0 100%); }
        .metal-strip {
            position: absolute;
            top: 17%; bottom: 0; width: 6px;
            background: linear-gradient(90deg,#2d5269,#b8e7fa,#25455a);
            opacity: .42;
        }
        .metal-strip.one { left: 18%; }
        .metal-strip.two { right: 18%; }

        .floor-display {
            position: absolute;
            z-index: 25;
            top: max(7%, env(safe-area-inset-top));
            left: 50%;
            transform: translateX(-50%);
            min-width: 160px;
            padding: 10px 18px 8px;
            text-align: center;
            border-radius: 12px;
            border: 1px solid rgba(92,220,255,.55);
            background: #02090f;
            box-shadow: inset 0 0 18px rgba(71,212,255,.1), 0 0 22px rgba(45,194,240,.18);
        }
        .floor-value {
            display: block;
            font: 900 clamp(1.6rem, 6vw, 2.5rem)/1 ui-monospace, SFMono-Regular, Menlo, monospace;
            color: #6de8ff;
            text-shadow: 0 0 14px rgba(76,224,255,.65);
            letter-spacing: .08em;
        }
        .floor-label {
            display: block;
            margin-top: 5px;
            font-size: .62rem;
            letter-spacing: .2em;
            color: #8ba8ba;
        }
        .half-level {
            position: absolute;
            z-index: 26;
            top: calc(max(7%, env(safe-area-inset-top)) + 82px);
            left: 50%;
            transform: translate(-50%, -8px);
            opacity: 0;
            border: 1px solid rgba(255,211,99,.5);
            background: rgba(34,26,8,.86);
            color: #ffdb79;
            border-radius: 99px;
            padding: 7px 13px;
            font-size: .72rem;
            font-weight: 900;
            letter-spacing: .08em;
            transition: opacity .5s ease, transform .5s ease;
        }
        .arrived .half-level.show { opacity: 1; transform: translate(-50%, 0); }

        .door-frame {
            position: absolute;
            z-index: 7;
            left: 19%;
            right: 19%;
            top: 18%;
            bottom: 4%;
            border: clamp(8px,1.5vw,18px) solid #355064;
            border-bottom-width: clamp(10px,2vw,24px);
            background: #02070c;
            box-shadow: inset 0 0 0 2px rgba(206,244,255,.15), 0 18px 45px rgba(0,0,0,.5);
            overflow: hidden;
        }
        .hallway {
            position: absolute;
            inset: 0;
            overflow: hidden;
            background:
                linear-gradient(rgba(203,236,247,.9),rgba(169,213,230,.9)) top/100% 18% no-repeat,
                linear-gradient(90deg,#bed5de 0 13%,#edf8fb 13% 87%,#bed5de 87% 100%);
            transform: scale(.94);
            transition: transform 2s cubic-bezier(.2,.7,.2,1);
        }
        .hallway::before {
            content: "";
            position: absolute;
            left: -10%; right: -10%; bottom: -18%;
            height: 48%;
            background:
                repeating-linear-gradient(90deg, rgba(27,75,94,.13) 0 1px, transparent 1px 82px),
                linear-gradient(#95b7c4,#dcebf0);
            transform: perspective(500px) rotateX(63deg);
            transform-origin: top;
        }
        .hallway::after {
            content: "PERSONEL VE ÇALIŞMA İLİŞKİLERİ";
            position: absolute;
            top: 25%;
            left: 50%;
            transform: translateX(-50%);
            white-space: nowrap;
            color: #183f54;
            font-weight: 900;
            letter-spacing: .16em;
            font-size: clamp(.68rem, 2vw, 1rem);
        }
        .walking .hallway { transform: scale(1.06); }

        .elevator-doors {
            position: absolute;
            inset: 0;
            z-index: 12;
            pointer-events: none;
        }
        .door {
            position: absolute;
            top: 0; bottom: 0; width: 50.2%;
            background:
                repeating-linear-gradient(90deg, rgba(255,255,255,.035) 0 2px, transparent 2px 12px),
                linear-gradient(90deg,#253b4a,#668194 45%,#233a4a);
            transition: transform 2.35s cubic-bezier(.72,.02,.18,1);
            box-shadow: inset 0 0 25px rgba(0,0,0,.34);
        }
        .door.left { left: 0; border-right: 1px solid #0a1117; }
        .door.right { right: 0; border-left: 1px solid rgba(218,244,255,.18); transform: scaleX(-1); }
        .doors-open .door.left { transform: translateX(-100%); }
        .doors-open .door.right { transform: scaleX(-1) translateX(-100%); }

        .speed-lines {
            position: absolute;
            z-index: 5;
            left: 20%; right: 20%; top: 18%; bottom: 4%;
            opacity: 0;
            pointer-events: none;
            background: repeating-linear-gradient(180deg, transparent 0 26px, rgba(84,210,255,.12) 27px 29px);
            animation: speedLines .55s linear infinite;
        }
        .riding .speed-lines { opacity: 1; }
        @keyframes speedLines { to { background-position-y: 58px; } }

        .first-person-hands {
            position: absolute;
            z-index: 30;
            left: 0; right: 0; bottom: -18px;
            height: 32%;
            pointer-events: none;
        }
        .hand {
            position: absolute;
            bottom: -12%;
            width: clamp(72px, 15vw, 150px);
            height: clamp(170px, 30vw, 300px);
            border-radius: 50% 50% 32% 32%;
            background: linear-gradient(90deg,#ba7d53,#efbd8e 48%,#c98b5e);
            box-shadow: inset 10px 0 20px rgba(88,42,20,.22);
            opacity: .88;
        }
        .hand.left { left: 5%; transform: rotate(10deg); }
        .hand.right { right: 5%; transform: rotate(-10deg); }
        .receiving .hand.right {
            animation: takeNote 1.8s cubic-bezier(.3,.7,.2,1) forwards;
        }
        @keyframes takeNote {
            0% { transform: rotate(-10deg) translate(0,0); }
            55% { transform: rotate(-4deg) translate(-14vw,-24vh); }
            100% { transform: rotate(-8deg) translate(-6vw,-7vh); }
        }

        .secretary {
            position: absolute;
            z-index: 4;
            left: 50%;
            bottom: 12%;
            width: clamp(95px, 16vw, 150px);
            height: clamp(220px, 43vw, 390px);
            transform: translate(-50%, 26%) scale(.38);
            opacity: 0;
            transition: transform 3.2s cubic-bezier(.18,.72,.16,1), opacity .8s ease;
            filter: drop-shadow(0 14px 15px rgba(27,62,75,.25));
        }
        .secretary.approach { transform: translate(-50%, 0) scale(1); opacity: 1; }
        .secretary-head {
            position: absolute;
            left: 50%; top: 0;
            width: 38%; aspect-ratio: .82;
            transform: translateX(-50%);
            border-radius: 48% 48% 44% 44%;
            background: linear-gradient(90deg,#d39a6d,#f3c49b);
            border: 6px solid #3d241f;
            border-bottom-width: 2px;
        }
        .secretary-hair {
            position: absolute;
            left: 24%; top: -2%;
            width: 52%; height: 27%;
            border-radius: 50% 50% 35% 35%;
            background: #2c1b18;
        }
        .secretary-body {
            position: absolute;
            left: 19%; right: 19%; top: 23%; bottom: 7%;
            clip-path: polygon(22% 0,78% 0,100% 100%,0 100%);
            border-radius: 20px 20px 8px 8px;
            background:
                linear-gradient(90deg, transparent 47%, rgba(255,255,255,.42) 48% 52%, transparent 53%),
                linear-gradient(#172b45,#0c1c30);
        }
        .secretary-collar {
            position: absolute;
            top: 23%; left: 37%; width: 26%; height: 16%;
            background: white;
            clip-path: polygon(0 0,50% 75%,100% 0,85% 100%,15% 100%);
        }
        .secretary-arm {
            position: absolute;
            top: 35%;
            width: 16%; height: 45%;
            border-radius: 30px;
            background: #172b45;
            transform-origin: top;
            transition: transform 1.4s ease;
        }
        .secretary-arm.left { left: 12%; transform: rotate(8deg); }
        .secretary-arm.right { right: 12%; transform: rotate(-8deg); }
        .secretary.offer .secretary-arm.right { transform: rotate(42deg) translate(-28%, -4%); }
        .paper-mini {
            position: absolute;
            z-index: 5;
            right: -9%;
            top: 60%;
            width: 43%; height: 25%;
            background: #fffdf2;
            border: 1px solid #c4a96a;
            transform: rotate(-8deg);
            opacity: 0;
            transition: opacity .6s ease .55s;
        }
        .secretary.offer .paper-mini { opacity: 1; }

        .note-overlay {
            position: fixed;
            z-index: 160;
            inset: 0;
            display: grid;
            place-items: center;
            padding: 22px;
            background: rgba(2,9,15,.18);
            opacity: 0;
            pointer-events: none;
            transition: opacity .8s ease;
        }
        .note-overlay.show { opacity: 1; }
        .note-paper {
            width: min(560px, 94vw);
            min-height: 320px;
            transform: translateY(28vh) rotate(-7deg) scale(.72);
            transition: transform 1.3s cubic-bezier(.2,.82,.2,1);
            border-radius: 8px;
            color: #173044;
            background:
                linear-gradient(rgba(32,93,122,.08) 1px, transparent 1px) 0 64px/100% 34px,
                #fffdf1;
            box-shadow: 0 35px 90px rgba(0,0,0,.5);
            padding: clamp(28px,7vw,56px);
            position: relative;
            overflow: hidden;
        }
        .note-overlay.show .note-paper { transform: translateY(0) rotate(-1.2deg) scale(1); }
        .note-paper::before {
            content: "";
            position: absolute;
            inset: 14px;
            border: 2px solid rgba(29,87,112,.18);
            pointer-events: none;
        }
        .note-logo {
            display: flex;
            align-items: center;
            gap: 10px;
            font-weight: 1000;
            letter-spacing: .08em;
            color: #245874;
        }
        .note-logo-mark {
            width: 36px; height: 36px;
            display: grid; place-items: center;
            border-radius: 10px;
            color: #fff;
            background: linear-gradient(145deg,#168bc7,#0b456b);
        }
        .note-title {
            margin: 42px 0 10px;
            font-family: Georgia, "Times New Roman", serif;
            font-size: clamp(1.35rem,4.8vw,2rem);
            color: #19394e;
        }
        .note-balance {
            font-size: clamp(2.15rem,9vw,4.4rem);
            line-height: 1;
            font-weight: 1000;
            color: #087eb7;
            margin: 20px 0 10px;
            letter-spacing: -.045em;
        }
        .note-caption { color: #4f6c7c; font-weight: 700; }
        .note-signature {
            margin-top: 34px;
            text-align: right;
            font-family: "Segoe Script", cursive;
            color: #315b70;
        }
        .skip-animation {
            position: fixed;
            z-index: 190;
            right: max(14px, env(safe-area-inset-right));
            bottom: max(14px, env(safe-area-inset-bottom));
            border: 1px solid rgba(159,226,255,.28);
            background: rgba(3,16,27,.72);
            color: #dff7ff;
            border-radius: 14px;
            padding: 11px 15px;
            cursor: pointer;
            backdrop-filter: blur(12px);
            font-weight: 800;
        }
        .elevator-status {
            position: fixed;
            z-index: 190;
            left: max(14px, env(safe-area-inset-left));
            bottom: max(14px, env(safe-area-inset-bottom));
            color: #b5d7e7;
            font-size: .78rem;
            letter-spacing: .08em;
        }

        /* DASHBOARD */
        .dashboard {
            width: min(1120px, 100%);
            margin: 0 auto;
            animation: dashboardIn .65s cubic-bezier(.2,.8,.2,1) both;
        }
        @keyframes dashboardIn {
            from { opacity: 0; transform: translateY(18px); }
            to { opacity: 1; transform: translateY(0); }
        }
        .dashboard-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            gap: 16px;
            margin-bottom: 18px;
        }
        .person {
            display: flex;
            align-items: center;
            gap: 14px;
            min-width: 0;
        }
        .avatar {
            width: 58px; height: 58px;
            flex: 0 0 58px;
            display: grid; place-items: center;
            border-radius: 18px;
            background: linear-gradient(145deg,var(--primary),var(--primary-2));
            color: #032039;
            font-size: 1.15rem;
            font-weight: 1000;
            box-shadow: 0 12px 26px rgba(28,164,220,.25);
        }
        .person h2 {
            margin: 0;
            font-size: clamp(1.25rem,4vw,1.85rem);
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
        }
        .person p { margin: 4px 0 0; color: var(--muted); }
        .header-buttons { display: flex; gap: 8px; }
        .small-btn {
            border: 1px solid var(--line);
            background: var(--panel);
            border-radius: 13px;
            padding: 11px 13px;
            cursor: pointer;
            font-weight: 800;
        }

        .dashboard-grid {
            display: grid;
            grid-template-columns: minmax(0,1.35fr) minmax(280px,.65fr);
            gap: 18px;
        }
        .main-column, .side-column { display: grid; gap: 18px; align-content: start; }
        .hero-card { padding: clamp(22px,5vw,36px); position: relative; overflow: hidden; }
        .hero-card::after {
            content: "";
            position: absolute;
            width: 260px; height: 260px;
            right: -90px; top: -90px;
            border-radius: 50%;
            background: radial-gradient(circle, rgba(85,217,255,.24), transparent 67%);
            pointer-events: none;
        }
        .hero-content {
            display: grid;
            grid-template-columns: 1fr auto;
            gap: 22px;
            align-items: center;
            position: relative;
            z-index: 2;
        }
        .eyebrow {
            color: var(--primary);
            text-transform: uppercase;
            letter-spacing: .13em;
            font-size: .76rem;
            font-weight: 1000;
        }
        .big-balance {
            margin: 10px 0 5px;
            font-size: clamp(3.2rem,12vw,6.3rem);
            line-height: .92;
            letter-spacing: -.075em;
            font-weight: 1000;
        }
        .big-balance small {
            font-size: .2em;
            letter-spacing: .02em;
            color: var(--muted);
            margin-left: 8px;
        }
        .update-info { color: var(--muted); font-size: .85rem; margin-top: 13px; }
        .progress-ring {
            --progress: 75deg;
            width: clamp(112px,20vw,160px);
            aspect-ratio: 1;
            border-radius: 50%;
            display: grid;
            place-items: center;
            background: conic-gradient(var(--primary) var(--progress), rgba(127,202,230,.13) 0);
            position: relative;
            box-shadow: 0 0 36px rgba(62,205,255,.12);
        }
        .progress-ring::before {
            content: "";
            position: absolute;
            inset: 12px;
            border-radius: 50%;
            background: var(--panel-strong);
            border: 1px solid var(--line);
        }
        .progress-ring span {
            position: relative;
            z-index: 2;
            text-align: center;
            color: var(--muted);
            font-size: .75rem;
            font-weight: 800;
        }
        .progress-ring b {
            display: block;
            color: var(--text);
            font-size: 1.15rem;
            margin-bottom: 2px;
        }

        .mini-cards {
            display: grid;
            grid-template-columns: repeat(2,minmax(0,1fr));
            gap: 14px;
        }
        .mini-card { padding: 20px; }
        .mini-card .value {
            margin-top: 9px;
            font-size: 2rem;
            font-weight: 1000;
        }
        .mini-card .label { color: var(--muted); font-size: .86rem; font-weight: 700; }

        .actions-card { padding: 20px; }
        .actions-card h3, .info-card h3, .id-card-wrap h3 {
            margin: 0 0 14px;
            font-size: 1rem;
        }
        .action-list { display: grid; gap: 10px; }
        .action-btn { justify-content: flex-start; }
        .action-btn.whatsapp { background: linear-gradient(135deg,#2bd881,#0d9d64); color: #04291b; }
        .action-btn.leave { background: linear-gradient(135deg,#ffd36d,#e99c25); color: #3b2601; }
        .action-btn.secondary {
            background: rgba(255,255,255,.06);
            border: 1px solid var(--line);
            color: var(--text);
        }

        .info-card { padding: 20px; }
        .info-row {
            display: flex;
            justify-content: space-between;
            gap: 12px;
            padding: 12px 0;
            border-bottom: 1px solid var(--line);
        }
        .info-row:last-child { border-bottom: 0; padding-bottom: 0; }
        .info-row span { color: var(--muted); }
        .info-row strong { text-align: right; }

        .holiday-radar {
            position: relative;
            min-height: 162px;
            display: grid;
            place-items: center;
            overflow: hidden;
            text-align: center;
        }
        .radar {
            position: absolute;
            width: 190px; height: 190px;
            border-radius: 50%;
            border: 1px solid rgba(84,218,255,.18);
            background:
                linear-gradient(90deg,transparent 49.5%,rgba(84,218,255,.12) 50%,transparent 50.5%),
                linear-gradient(transparent 49.5%,rgba(84,218,255,.12) 50%,transparent 50.5%),
                radial-gradient(circle,transparent 0 24%,rgba(84,218,255,.08) 25% 25.8%,transparent 26% 49%,rgba(84,218,255,.08) 50% 50.8%,transparent 51%);
        }
        .radar::after {
            content: "";
            position: absolute;
            inset: 50% 50% auto auto;
            width: 50%; height: 2px;
            transform-origin: left;
            background: linear-gradient(90deg,var(--primary),transparent);
            animation: radarSpin 4s linear infinite;
        }
        @keyframes radarSpin { to { transform: rotate(360deg); } }
        .holiday-content { position: relative; z-index: 2; }
        .holiday-days { font-size: 2.5rem; font-weight: 1000; color: var(--primary); }
        .holiday-name { font-weight: 900; }
        .holiday-date { color: var(--muted); font-size: .82rem; margin-top: 4px; }

        .id-card-wrap { padding: 20px; }
        .digital-id {
            min-height: 218px;
            padding: 20px;
            border-radius: 20px;
            color: white;
            background:
                radial-gradient(circle at 85% 18%,rgba(83,229,255,.32),transparent 30%),
                linear-gradient(145deg,#08203a,#0a5279);
            border: 1px solid rgba(130,224,255,.33);
            position: relative;
            overflow: hidden;
            box-shadow: 0 18px 38px rgba(0,0,0,.22);
        }
        .digital-id::after {
            content: "";
            position: absolute;
            width: 160px; height: 160px;
            border: 24px solid rgba(255,255,255,.035);
            border-radius: 50%;
            right: -55px; bottom: -78px;
        }
        .id-top {
            display: flex;
            justify-content: space-between;
            align-items: flex-start;
            gap: 14px;
        }
        .id-company { font-size: .72rem; letter-spacing: .12em; font-weight: 1000; color: #a9eaff; }
        .id-chip {
            width: 42px; height: 32px;
            border-radius: 7px;
            background:
                linear-gradient(90deg,transparent 47%,rgba(77,60,5,.3) 48% 52%,transparent 53%),
                linear-gradient(#f3d77b,#bb9027);
            border: 1px solid rgba(255,239,169,.8);
        }
        .id-main {
            display: grid;
            grid-template-columns: 1fr 92px;
            align-items: end;
            gap: 14px;
            margin-top: 28px;
            position: relative;
            z-index: 2;
        }
        .id-name { font-size: 1.25rem; font-weight: 1000; }
        .id-role { color: #b9dded; font-size: .78rem; margin: 5px 0 20px; }
        .id-number { color: #86dfff; font: 800 .78rem ui-monospace, monospace; }
        .qr-box {
            width: 92px; height: 92px;
            padding: 7px;
            border-radius: 10px;
            background: white;
            display: grid;
            place-items: center;
        }
        #qrCanvas { width: 78px; height: 78px; }

        /* MODALS + TOAST */
        .modal-backdrop {
            position: fixed;
            z-index: 220;
            inset: 0;
            display: grid;
            place-items: center;
            padding: 20px;
            background: rgba(1,10,18,.68);
            backdrop-filter: blur(9px);
            opacity: 0;
            pointer-events: none;
            transition: opacity .25s ease;
        }
        .modal-backdrop.open { opacity: 1; pointer-events: auto; }
        .modal {
            width: min(520px,100%);
            max-height: min(760px,90dvh);
            overflow: auto;
            padding: 24px;
            transform: translateY(14px) scale(.98);
            transition: transform .25s ease;
        }
        .modal-backdrop.open .modal { transform: translateY(0) scale(1); }
        .modal-head {
            display: flex;
            justify-content: space-between;
            align-items: center;
            gap: 12px;
            margin-bottom: 18px;
        }
        .modal h3 { margin: 0; }
        .close-btn {
            width: 38px; height: 38px;
            display: grid; place-items: center;
            border-radius: 12px;
            border: 1px solid var(--line);
            background: rgba(255,255,255,.05);
            cursor: pointer;
        }
        .modal textarea { min-height: 100px; resize: vertical; padding-right: 15px; }
        .modal select { padding-right: 15px; }
        .form-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; }
        .modal .field { margin-bottom: 13px; }

        #toast {
            position: fixed;
            z-index: 400;
            top: max(72px, calc(env(safe-area-inset-top) + 58px));
            right: 16px;
            width: min(360px, calc(100vw - 32px));
            padding: 14px 16px;
            border-radius: 15px;
            background: var(--panel-strong);
            border: 1px solid var(--line);
            box-shadow: 0 18px 45px rgba(0,0,0,.3);
            transform: translateX(calc(100% + 30px));
            opacity: 0;
            transition: transform .4s cubic-bezier(.2,.8,.2,1), opacity .3s ease;
            font-weight: 800;
        }
        #toast.show { transform: translateX(0); opacity: 1; }
        #toast.error { border-color: rgba(255,102,120,.55); }
        #toast.success { border-color: rgba(43,216,129,.55); }

        @media (max-width: 820px) {
            .dashboard-grid { grid-template-columns: 1fr; }
            .dashboard-header { align-items: flex-start; }
            .header-buttons { flex-direction: column; }
            .hero-content { grid-template-columns: 1fr; }
            .progress-ring { width: 126px; }
            .hero-content .progress-ring { position: absolute; right: 0; top: 0; opacity: .5; transform: scale(.72); transform-origin: top right; }
            .door-frame { left: 6%; right: 6%; }
            .side-wall { width: 8%; }
            .metal-strip.one { left: 6%; }
            .metal-strip.two { right: 6%; }
        }
        @media (max-width: 520px) {
            .brand-mini span:last-child { display: none; }
            .install-btn span { display: none; }
            .topbar { inset-left: 10px; inset-right: 10px; }
            .page-shell { padding-left: 12px; padding-right: 12px; }
            .login-card { border-radius: 22px; }
            .dashboard-header { display: grid; grid-template-columns: 1fr auto; }
            .person { min-width: 0; }
            .avatar { width: 50px; height: 50px; flex-basis: 50px; }
            .person p { font-size: .78rem; }
            .mini-cards { grid-template-columns: 1fr 1fr; }
            .form-grid { grid-template-columns: 1fr; }
            .id-main { grid-template-columns: 1fr 84px; }
            .qr-box { width: 84px; height: 84px; }
            #qrCanvas { width: 70px; height: 70px; }
            .elevator-status { display: none; }
        }
        @media (prefers-reduced-motion: reduce) {
            *, *::before, *::after {
                animation-duration: .01ms !important;
                animation-iteration-count: 1 !important;
                scroll-behavior: auto !important;
                transition-duration: .01ms !important;
            }
        }
    </style>
</head>
<body>
    <div class="topbar">
        <div class="brand-mini">
            <span class="brand-mark">İ</span>
            <span data-i18n="brand">İzin Portalı</span>
        </div>
        <div class="top-actions">
            <button class="install-btn" id="installBtn" type="button" aria-label="Uygulamayı yükle">
                <span>＋</span><span data-i18n="install">Uygulamayı Yükle</span>
            </button>
            <button class="icon-btn" id="languageBtn" type="button" aria-label="Dil değiştir">TR</button>
            <button class="icon-btn" id="themeBtn" type="button" aria-label="Tema değiştir">☾</button>
        </div>
    </div>

    <main class="page-shell">
        <section class="login-card glass" id="loginPanel">
            <div class="login-logo"><div class="elevator-icon"></div></div>
            <h1 data-i18n="loginTitle">Personel İzin Sistemi</h1>
            <p class="login-subtitle" data-i18n="loginSubtitle">Kişisel izin bilgilerinize güvenli şekilde ulaşın.</p>

            <form id="loginForm" novalidate>
                <div class="field">
                    <label for="username" data-i18n="username">Kullanıcı Adı</label>
                    <div class="input-wrap">
                        <input id="username" name="username" autocomplete="username" inputmode="text"
                               data-i18n-placeholder="usernamePlaceholder" placeholder="Kullanıcı adınızı girin">
                        <span class="field-icon">●</span>
                    </div>
                </div>

                <div class="field">
                    <label for="password" data-i18n="password">Şifre</label>
                    <div class="input-wrap">
                        <input id="password" name="password" type="password" autocomplete="current-password"
                               data-i18n-placeholder="passwordPlaceholder" placeholder="Şifrenizi girin">
                        <button class="field-icon password-toggle" id="passwordToggle" type="button" aria-label="Şifreyi göster">◉</button>
                    </div>
                </div>

                <button class="primary-btn" id="loginBtn" type="submit">
                    <span id="loginBtnText" data-i18n="login">Giriş Yap</span>
                    <span class="spinner"></span>
                </button>
            </form>

            <div class="login-links">
                <button class="text-btn" id="forgotBtn" type="button" data-i18n="forgot">Şifremi Unuttum</button>
            </div>
        </section>

        <section class="dashboard" id="dashboard" hidden>
            <div class="dashboard-header">
                <div class="person">
                    <div class="avatar" id="avatar">BB</div>
                    <div>
                        <h2 id="greeting">Hoş geldiniz</h2>
                        <p><span id="personRole">Personel</span> · <span id="personUsername"></span></p>
                    </div>
                </div>
                <div class="header-buttons">
                    <button class="small-btn" id="replayBtn" type="button" data-i18n="replay">Asansörü Tekrar İzle</button>
                    <button class="small-btn" id="logoutBtn" type="button" data-i18n="logout">Çıkış</button>
                </div>
            </div>

            <div class="dashboard-grid">
                <div class="main-column">
                    <article class="hero-card glass">
                        <div class="hero-content">
                            <div>
                                <div class="eyebrow" data-i18n="remainingLeave">Kalan İzin Hakkınız</div>
                                <div class="big-balance"><span id="remainingLeave">0</span><small data-i18n="day">GÜN</small></div>
                                <div class="update-info"><span data-i18n="updated">Son güncelleme:</span> <strong id="updatedAt">-</strong></div>
                            </div>
                            <div class="progress-ring" id="progressRing">
                                <span><b id="ringValue">0</b><span data-i18n="leaveLevel">İzin seviyesi</span></span>
                            </div>
                        </div>
                    </article>

                    <div class="mini-cards">
                        <article class="mini-card glass">
                            <div class="label" data-i18n="sundayLeave">Pazar İzinleri</div>
                            <div class="value"><span id="sundayLeave">0</span> <small data-i18n="dayLower">gün</small></div>
                        </article>
                        <article class="mini-card glass">
                            <div class="label" data-i18n="officialHoliday">Resmî Tatil</div>
                            <div class="value"><span id="officialHoliday">0</span> <small data-i18n="dayLower">gün</small></div>
                        </article>
                    </div>

                    <article class="actions-card glass">
                        <h3 data-i18n="quickActions">Hızlı İşlemler</h3>
                        <div class="action-list">
                            <a class="action-btn whatsapp" id="objectionBtn" href="#" target="_blank" rel="noopener">
                                <span>◉</span><span data-i18n="objectLeave">İzin Gününe İtiraz Et</span>
                            </a>
                            <button class="action-btn leave" id="leaveRequestBtn" type="button">
                                <span>▣</span><span data-i18n="requestLeave">İzin Talebi Oluştur</span>
                            </button>
                            <button class="action-btn secondary" id="installActionBtn" type="button">
                                <span>＋</span><span data-i18n="addHome">Ana Ekrana Uygulama Olarak Ekle</span>
                            </button>
                        </div>
                    </article>
                </div>

                <aside class="side-column">
                    <article class="info-card glass holiday-radar">
                        <div class="radar"></div>
                        <div class="holiday-content">
                            <div class="eyebrow" data-i18n="nextHoliday">Yaklaşan Resmî Tatil</div>
                            <div class="holiday-days" id="holidayDays">-</div>
                            <div class="holiday-name" id="holidayName">-</div>
                            <div class="holiday-date" id="holidayDate">-</div>
                        </div>
                    </article>

                    <article class="info-card glass">
                        <h3 data-i18n="security">Güvenlik Bilgisi</h3>
                        <div class="info-row">
                            <span data-i18n="lastLogin">Son girişiniz</span>
                            <strong id="lastLogin">İlk giriş</strong>
                        </div>
                        <div class="info-row">
                            <span data-i18n="session">Oturum</span>
                            <strong data-i18n="active">Aktif</strong>
                        </div>
                    </article>

                    <article class="id-card-wrap glass">
                        <h3 data-i18n="digitalId">Dijital Personel Kimliği</h3>
                        <div class="digital-id">
                            <div class="id-top">
                                <div class="id-company">PERSONEL PORTALI</div>
                                <div class="id-chip"></div>
                            </div>
                            <div class="id-main">
                                <div>
                                    <div class="id-name" id="idName">Personel</div>
                                    <div class="id-role" id="idRole">Görev</div>
                                    <div class="id-number">ID: <span id="idNumber">-</span></div>
                                </div>
                                <div class="qr-box"><canvas id="qrCanvas" width="156" height="156"></canvas></div>
                            </div>
                        </div>
                    </article>
                </aside>
            </div>
        </section>
    </main>

    <!-- Birinci şahıs asansör sahnesi -->
    <section class="elevator-scene" id="elevatorScene" hidden aria-label="İzin asansörü animasyonu">
        <div class="cabin">
            <div class="cabin-ceiling"></div>
            <div class="ceiling-light"></div>
            <div class="side-wall left"></div>
            <div class="side-wall right"></div>
            <div class="metal-strip one"></div>
            <div class="metal-strip two"></div>

            <div class="floor-display">
                <span class="floor-value" id="floorValue">0</span>
                <span class="floor-label" data-i18n="leaveLevelUpper">İZİN SEVİYESİ</span>
            </div>
            <div class="half-level" id="halfLevel" data-i18n="halfDay">ARA SEVİYE · ½ GÜN</div>

            <div class="door-frame">
                <div class="hallway">
                    <div class="secretary" id="secretary">
                        <div class="secretary-hair"></div>
                        <div class="secretary-head"></div>
                        <div class="secretary-collar"></div>
                        <div class="secretary-body"></div>
                        <div class="secretary-arm left"></div>
                        <div class="secretary-arm right"></div>
                        <div class="paper-mini"></div>
                    </div>
                </div>
                <div class="speed-lines"></div>
                <div class="elevator-doors">
                    <div class="door left"></div>
                    <div class="door right"></div>
                </div>
            </div>

            <div class="first-person-hands">
                <div class="hand left"></div>
                <div class="hand right"></div>
            </div>
        </div>

        <div class="note-overlay" id="noteOverlay">
            <div class="note-paper">
                <div class="note-logo"><span class="note-logo-mark">İ</span> PERSONEL İZİN BİLDİRİMİ</div>
                <div class="note-title" id="noteTitle">Sn. Burhan Biliktü</div>
                <div class="note-caption" data-i18n="noteCaption">Kalan İzin Hakkınız</div>
                <div class="note-balance"><span id="noteBalance">29,5</span> <small data-i18n="dayLower">Gün</small></div>
                <div class="note-signature" data-i18n="hrUnit">Personel ve Çalışma İlişkileri</div>
            </div>
        </div>

        <div class="elevator-status" id="elevatorStatus">YUKARI ÇIKIYOR · YAVAŞ SEYİR</div>
        <button class="skip-animation" id="skipAnimationBtn" type="button" data-i18n="skip">Animasyonu Geç</button>
    </section>

    <!-- Şifre unutma -->
    <div class="modal-backdrop" id="forgotModal" aria-hidden="true">
        <div class="modal glass">
            <div class="modal-head">
                <h3 data-i18n="forgotTitle">Şifre Talebi</h3>
                <button class="close-btn" type="button" data-close="forgotModal">×</button>
            </div>
            <div class="field">
                <label for="forgotIdentity" data-i18n="nameOrUsername">Ad Soyad veya Kullanıcı Adı</label>
                <input id="forgotIdentity" data-i18n-placeholder="identityPlaceholder" placeholder="Bilginizi yazın">
            </div>
            <button class="primary-btn" id="forgotWhatsappBtn" type="button" data-i18n="sendWhatsapp">WhatsApp’tan Gönder</button>
        </div>
    </div>

    <!-- İzin talebi -->
    <div class="modal-backdrop" id="leaveModal" aria-hidden="true">
        <div class="modal glass">
            <div class="modal-head">
                <h3 data-i18n="leaveRequestTitle">İzin Talebi Oluştur</h3>
                <button class="close-btn" type="button" data-close="leaveModal">×</button>
            </div>
            <div class="field">
                <label for="leaveType" data-i18n="leaveType">İzin Türü</label>
                <select id="leaveType">
                    <option value="Yıllık İzin" data-i18n="annualLeave">Yıllık İzin</option>
                    <option value="Mazeret İzni" data-i18n="excuseLeave">Mazeret İzni</option>
                    <option value="Ücretsiz İzin" data-i18n="unpaidLeave">Ücretsiz İzin</option>
                </select>
            </div>
            <div class="form-grid">
                <div class="field">
                    <label for="leaveStart" data-i18n="startDate">Başlangıç Tarihi</label>
                    <input id="leaveStart" type="date">
                </div>
                <div class="field">
                    <label for="leaveEnd" data-i18n="endDate">Bitiş Tarihi</label>
                    <input id="leaveEnd" type="date">
                </div>
            </div>
            <div class="field">
                <label for="leaveDescription" data-i18n="description">Açıklama</label>
                <textarea id="leaveDescription" data-i18n-placeholder="descriptionPlaceholder" placeholder="Talebinizle ilgili kısa açıklama"></textarea>
            </div>
            <button class="primary-btn" id="sendLeaveBtn" type="button" data-i18n="sendWhatsapp">WhatsApp’tan Gönder</button>
        </div>
    </div>

    <div id="toast" role="status" aria-live="polite"></div>

    <script>
        const WHATSAPP_NUMBER = "905459157444";
        let currentUser = null;
        let deferredInstallPrompt = null;
        let currentLanguage = localStorage.getItem("izin-language") || "tr";
        let animationRunId = 0;

        const translations = {
            tr: {
                brand: "İzin Portalı", install: "Uygulamayı Yükle",
                loginTitle: "Personel İzin Sistemi",
                loginSubtitle: "Kişisel izin bilgilerinize güvenli şekilde ulaşın.",
                username: "Kullanıcı Adı", password: "Şifre",
                usernamePlaceholder: "Kullanıcı adınızı girin",
                passwordPlaceholder: "Şifrenizi girin",
                login: "Giriş Yap", forgot: "Şifremi Unuttum",
                replay: "Asansörü Tekrar İzle", logout: "Çıkış",
                remainingLeave: "Kalan İzin Hakkınız", day: "GÜN", dayLower: "gün",
                updated: "Son güncelleme:", leaveLevel: "İzin seviyesi",
                sundayLeave: "Pazar İzinleri", officialHoliday: "Resmî Tatil",
                quickActions: "Hızlı İşlemler", objectLeave: "İzin Gününe İtiraz Et",
                requestLeave: "İzin Talebi Oluştur", addHome: "Ana Ekrana Uygulama Olarak Ekle",
                nextHoliday: "Yaklaşan Resmî Tatil", security: "Güvenlik Bilgisi",
                lastLogin: "Son girişiniz", session: "Oturum", active: "Aktif",
                digitalId: "Dijital Personel Kimliği", leaveLevelUpper: "İZİN SEVİYESİ",
                halfDay: "ARA SEVİYE · ½ GÜN", noteCaption: "Kalan İzin Hakkınız",
                hrUnit: "Personel ve Çalışma İlişkileri", skip: "Animasyonu Geç",
                forgotTitle: "Şifre Talebi", nameOrUsername: "Ad Soyad veya Kullanıcı Adı",
                identityPlaceholder: "Bilginizi yazın", sendWhatsapp: "WhatsApp’tan Gönder",
                leaveRequestTitle: "İzin Talebi Oluştur", leaveType: "İzin Türü",
                annualLeave: "Yıllık İzin", excuseLeave: "Mazeret İzni",
                unpaidLeave: "Ücretsiz İzin", startDate: "Başlangıç Tarihi",
                endDate: "Bitiş Tarihi", description: "Açıklama",
                descriptionPlaceholder: "Talebinizle ilgili kısa açıklama",
                firstLogin: "İlk giriş", greetingMorning: "Günaydın", greetingDay: "Hoş geldiniz",
                greetingEvening: "İyi akşamlar", elevatorRising: "YUKARI ÇIKIYOR · YAVAŞ SEYİR",
                elevatorArrived: "29,5 SEVİYESİNE ULAŞILDI", doorsOpening: "KAPILAR AÇILIYOR",
                secretaryComing: "PERSONEL BİLDİRİMİ HAZIRLANIYOR",
                noteDelivery: "BİLDİRİM TESLİM EDİLİYOR",
                fillFields: "Lütfen kullanıcı adı ve şifreyi girin.",
                loginError: "Kullanıcı adı veya şifre hatalı.",
                serverError: "Sunucuya bağlanılamadı.",
                blocked: "Çok fazla hatalı deneme. Lütfen daha sonra tekrar deneyin.",
                installReady: "Uygulama ana ekrana eklenmeye hazır.",
                iosInstall: "iPhone’da Paylaş simgesine dokunup “Ana Ekrana Ekle” seçeneğini kullanın.",
                installUnavailable: "Tarayıcı menüsünden “Ana ekrana ekle” seçeneğini kullanabilirsiniz.",
                fillIdentity: "Lütfen adınızı veya kullanıcı adınızı yazın.",
                fillDates: "Lütfen başlangıç ve bitiş tarihlerini seçin.",
                invalidDates: "Bitiş tarihi başlangıç tarihinden önce olamaz.",
                sentToWhatsapp: "WhatsApp açılıyor.",
                notePrefix: "Sn.",
                objectionMessage: name => `Merhaba, adım ${name}. İzin sisteminde görünen gün sayısının hatalı olduğunu düşünüyorum. Kalan izin hakkım: ${formatNumber(currentUser.remaining_leave)} gün. Son güncelleme: ${currentUser.updated_at}. İtiraz etmek istiyorum.`,
                forgotMessage: identity => `Merhaba, Personel İzin Portalı şifremi unuttum. Ad Soyad / Kullanıcı Adı: ${identity}. Şifre konusunda destek rica ederim.`,
                leaveMessage: data => `Merhaba, izin talebimi iletmek istiyorum.\n\nAd Soyad: ${currentUser.name}\nKullanıcı Adı: ${currentUser.username}\nİzin Türü: ${data.type}\nBaşlangıç: ${data.start}\nBitiş: ${data.end}\nToplam Takvim Günü: ${data.days}\nMevcut Kalan İzin: ${formatNumber(currentUser.remaining_leave)} gün\nAçıklama: ${data.description || "-"}\n\nOnaya sunarım.`,
                noteTitle: name => `Sn. ${name}`,
                holidayIn: days => `${days} gün`,
            },
            en: {
                brand: "Leave Portal", install: "Install App",
                loginTitle: "Employee Leave System",
                loginSubtitle: "Securely access your personal leave information.",
                username: "Username", password: "Password",
                usernamePlaceholder: "Enter your username",
                passwordPlaceholder: "Enter your password",
                login: "Sign In", forgot: "Forgot Password",
                replay: "Replay Elevator", logout: "Sign Out",
                remainingLeave: "Remaining Leave Balance", day: "DAYS", dayLower: "days",
                updated: "Last update:", leaveLevel: "Leave level",
                sundayLeave: "Sunday Leave", officialHoliday: "Public Holiday",
                quickActions: "Quick Actions", objectLeave: "Object to Leave Balance",
                requestLeave: "Create Leave Request", addHome: "Add App to Home Screen",
                nextHoliday: "Next Public Holiday", security: "Security Information",
                lastLogin: "Your last sign-in", session: "Session", active: "Active",
                digitalId: "Digital Employee ID", leaveLevelUpper: "LEAVE LEVEL",
                halfDay: "HALF LEVEL · ½ DAY", noteCaption: "Your Remaining Leave Balance",
                hrUnit: "Personnel and Labour Relations", skip: "Skip Animation",
                forgotTitle: "Password Request", nameOrUsername: "Full Name or Username",
                identityPlaceholder: "Enter your information", sendWhatsapp: "Send via WhatsApp",
                leaveRequestTitle: "Create Leave Request", leaveType: "Leave Type",
                annualLeave: "Annual Leave", excuseLeave: "Excuse Leave",
                unpaidLeave: "Unpaid Leave", startDate: "Start Date",
                endDate: "End Date", description: "Description",
                descriptionPlaceholder: "Briefly explain your request",
                firstLogin: "First sign-in", greetingMorning: "Good morning", greetingDay: "Welcome",
                greetingEvening: "Good evening", elevatorRising: "GOING UP · SLOW RIDE",
                elevatorArrived: "LEAVE LEVEL REACHED", doorsOpening: "DOORS OPENING",
                secretaryComing: "PREPARING PERSONNEL NOTICE",
                noteDelivery: "DELIVERING NOTICE",
                fillFields: "Please enter your username and password.",
                loginError: "Incorrect username or password.",
                serverError: "Could not connect to the server.",
                blocked: "Too many failed attempts. Please try again later.",
                installReady: "The app is ready to be installed.",
                iosInstall: "On iPhone, tap Share and choose “Add to Home Screen”.",
                installUnavailable: "Use your browser menu and choose “Add to Home screen”.",
                fillIdentity: "Please enter your name or username.",
                fillDates: "Please select start and end dates.",
                invalidDates: "End date cannot be before start date.",
                sentToWhatsapp: "Opening WhatsApp.",
                notePrefix: "Dear",
                objectionMessage: name => `Hello, my name is ${name}. I believe the leave balance shown in the system is incorrect. Remaining leave: ${formatNumber(currentUser.remaining_leave)} days. Last update: ${currentUser.updated_at}. I would like to submit an objection.`,
                forgotMessage: identity => `Hello, I forgot my Employee Leave Portal password. Full Name / Username: ${identity}. I kindly request support.`,
                leaveMessage: data => `Hello, I would like to submit a leave request.\n\nName: ${currentUser.name}\nUsername: ${currentUser.username}\nLeave Type: ${data.type}\nStart: ${data.start}\nEnd: ${data.end}\nCalendar Days: ${data.days}\nCurrent Leave Balance: ${formatNumber(currentUser.remaining_leave)} days\nDescription: ${data.description || "-"}\n\nSubmitted for approval.`,
                noteTitle: name => `Dear ${name}`,
                holidayIn: days => `${days} days`,
            }
        };

        function t(key) {
            return translations[currentLanguage][key] ?? translations.tr[key] ?? key;
        }

        function applyLanguage() {
            document.documentElement.lang = currentLanguage;
            document.getElementById("languageBtn").textContent = currentLanguage.toUpperCase();

            document.querySelectorAll("[data-i18n]").forEach(element => {
                const value = t(element.dataset.i18n);
                if (typeof value === "string") element.textContent = value;
            });
            document.querySelectorAll("[data-i18n-placeholder]").forEach(element => {
                element.placeholder = t(element.dataset.i18nPlaceholder);
            });

            if (currentUser) {
                renderUser(currentUser, false);
                document.getElementById("noteTitle").textContent = t("noteTitle")(currentUser.name);
            }
            calculateNextHoliday();
        }

        function autoTheme() {
            const saved = localStorage.getItem("izin-theme");
            const hour = new Date().getHours();
            const light = saved ? saved === "light" : !(hour >= 19 || hour < 7);
            document.body.classList.toggle("light", light);
            document.getElementById("themeBtn").textContent = light ? "☀" : "☾";
            document.querySelector('meta[name="theme-color"]').setAttribute("content", light ? "#e7f7ff" : "#071b30");
        }

        function toggleTheme() {
            const light = !document.body.classList.contains("light");
            document.body.classList.toggle("light", light);
            localStorage.setItem("izin-theme", light ? "light" : "dark");
            autoTheme();
        }

        function formatNumber(value) {
            const number = Number(value || 0);
            return new Intl.NumberFormat(currentLanguage === "tr" ? "tr-TR" : "en-US", {
                maximumFractionDigits: 2
            }).format(number);
        }

        function formatDateTime(isoValue) {
            if (!isoValue) return t("firstLogin");
            const date = new Date(isoValue);
            if (Number.isNaN(date.getTime())) return t("firstLogin");
            return new Intl.DateTimeFormat(currentLanguage === "tr" ? "tr-TR" : "en-GB", {
                dateStyle: "medium", timeStyle: "short"
            }).format(date);
        }

        function showToast(message, type = "") {
            const toast = document.getElementById("toast");
            toast.textContent = message;
            toast.className = type;
            requestAnimationFrame(() => toast.classList.add("show"));
            clearTimeout(showToast.timer);
            showToast.timer = setTimeout(() => toast.classList.remove("show"), 3500);
        }

        function openModal(id) {
            const modal = document.getElementById(id);
            modal.classList.add("open");
            modal.setAttribute("aria-hidden", "false");
        }

        function closeModal(id) {
            const modal = document.getElementById(id);
            modal.classList.remove("open");
            modal.setAttribute("aria-hidden", "true");
        }

        function openWhatsApp(message) {
            const url = `https://wa.me/${WHATSAPP_NUMBER}?text=${encodeURIComponent(message)}`;
            window.open(url, "_blank", "noopener");
            showToast(t("sentToWhatsapp"), "success");
        }

        async function login(event) {
            event.preventDefault();
            const username = document.getElementById("username").value.trim();
            const password = document.getElementById("password").value.trim();
            if (!username || !password) {
                showToast(t("fillFields"), "error");
                return;
            }

            const button = document.getElementById("loginBtn");
            button.disabled = true;
            button.classList.add("loading");

            try {
                const response = await fetch("/login", {
                    method: "POST",
                    headers: {"Content-Type": "application/json"},
                    body: JSON.stringify({username, password})
                });
                const result = await response.json();

                if (!response.ok || result.status !== "success") {
                    if (response.status === 429) showToast(result.message || t("blocked"), "error");
                    else showToast(result.message || t("loginError"), "error");
                    return;
                }

                currentUser = result.data;
                const loginKey = `izin-last-login-${currentUser.username}`;
                currentUser.previous_login = localStorage.getItem(loginKey);
                localStorage.setItem(loginKey, new Date().toISOString());

                document.getElementById("password").value = "";
                await playElevatorAnimation(currentUser);
            } catch (error) {
                console.error(error);
                showToast(t("serverError"), "error");
            } finally {
                button.disabled = false;
                button.classList.remove("loading");
            }
        }

        function delay(ms, runId) {
            return new Promise(resolve => {
                const timer = setTimeout(() => resolve(true), ms);
                const checker = setInterval(() => {
                    if (runId !== animationRunId) {
                        clearTimeout(timer);
                        clearInterval(checker);
                        resolve(false);
                    }
                }, 80);
                setTimeout(() => clearInterval(checker), ms + 120);
            });
        }

        function animateFloor(target, duration, runId) {
            return new Promise(resolve => {
                const start = performance.now();
                const floor = document.getElementById("floorValue");

                function frame(now) {
                    if (runId !== animationRunId) {
                        resolve(false);
                        return;
                    }

                    const progress = Math.min((now - start) / duration, 1);
                    // Yavaş başlangıç, ortada düzenli hareket, hedefe yaklaşırken yavaşlama.
                    const eased = progress < .5
                        ? 2 * progress * progress
                        : 1 - Math.pow(-2 * progress + 2, 2) / 2;

                    let value = target * eased;
                    value = Math.floor(value * 2) / 2;
                    floor.textContent = formatNumber(value);

                    if (progress < 1) {
                        requestAnimationFrame(frame);
                    } else {
                        floor.textContent = formatNumber(target);
                        resolve(true);
                    }
                }
                requestAnimationFrame(frame);
            });
        }

        function resetElevatorScene() {
            const scene = document.getElementById("elevatorScene");
            scene.className = "elevator-scene";
            document.getElementById("secretary").className = "secretary";
            document.getElementById("noteOverlay").className = "note-overlay";
            document.getElementById("halfLevel").className = "half-level";
            document.getElementById("floorValue").textContent = "0";
        }

        async function playElevatorAnimation(user) {
            const runId = ++animationRunId;
            resetElevatorScene();

            const scene = document.getElementById("elevatorScene");
            const loginPanel = document.getElementById("loginPanel");
            const dashboard = document.getElementById("dashboard");
            const secretary = document.getElementById("secretary");
            const noteOverlay = document.getElementById("noteOverlay");
            const halfLevel = document.getElementById("halfLevel");
            const status = document.getElementById("elevatorStatus");

            loginPanel.hidden = true;
            dashboard.hidden = true;
            scene.hidden = false;

            document.getElementById("noteTitle").textContent = t("noteTitle")(user.name);
            document.getElementById("noteBalance").textContent = formatNumber(user.remaining_leave);

            const target = Math.max(0, Number(user.remaining_leave || 0));
            const hasHalf = Math.abs(target * 2 - Math.round(target * 2)) < .001 &&
                            Math.abs(target - Math.round(target)) > .001;
            if (hasHalf) halfLevel.classList.add("show");

            status.textContent = t("elevatorRising");
            scene.classList.add("riding");

            // Kullanıcının istediği gibi hızlı bitmez: ana yükseliş yaklaşık 14 saniye sürer.
            const completed = await animateFloor(target, 14000, runId);
            if (!completed || runId !== animationRunId) return;

            scene.classList.remove("riding");
            scene.classList.add("arrived");
            status.textContent = currentLanguage === "tr"
                ? `${formatNumber(target)} SEVİYESİNE ULAŞILDI`
                : `LEVEL ${formatNumber(target)} REACHED`;

            if (!(await delay(900, runId))) return;
            scene.classList.add("doors-open");
            status.textContent = t("doorsOpening");

            if (!(await delay(2500, runId))) return;
            scene.classList.add("walking");
            status.textContent = t("secretaryComing");

            if (!(await delay(1600, runId))) return;
            secretary.classList.add("approach");

            if (!(await delay(3300, runId))) return;
            secretary.classList.add("offer");
            scene.classList.add("receiving");
            status.textContent = t("noteDelivery");

            if (!(await delay(1700, runId))) return;
            noteOverlay.classList.add("show");

            // Notun okunabilmesi için ekranda yeterince uzun kalır.
            if (!(await delay(5200, runId))) return;
            finishAnimation(runId);
        }

        function finishAnimation(runId = null) {
            if (runId !== null && runId !== animationRunId) return;
            animationRunId++;
            document.getElementById("elevatorScene").hidden = true;
            renderUser(currentUser, true);
            document.getElementById("dashboard").hidden = false;
            window.scrollTo({top: 0, behavior: "smooth"});
        }

        function getGreeting(name) {
            const hour = new Date().getHours();
            const firstName = String(name || "").split(" ")[0];
            if (hour < 12) return `${t("greetingMorning")}, ${firstName}`;
            if (hour >= 18) return `${t("greetingEvening")}, ${firstName}`;
            return `${t("greetingDay")}, ${firstName}`;
        }

        function renderUser(user, animate = true) {
            if (!user) return;
            document.getElementById("avatar").textContent = user.initials || "P";
            document.getElementById("greeting").textContent = getGreeting(user.name);
            document.getElementById("personRole").textContent = user.role || "Personel";
            document.getElementById("personUsername").textContent = user.username || "";
            document.getElementById("remainingLeave").textContent = formatNumber(user.remaining_leave);
            document.getElementById("ringValue").textContent = formatNumber(user.remaining_leave);
            document.getElementById("sundayLeave").textContent = formatNumber(user.sunday_leave);
            document.getElementById("officialHoliday").textContent = formatNumber(user.official_holiday);
            document.getElementById("updatedAt").textContent = user.updated_at || "-";
            document.getElementById("lastLogin").textContent = formatDateTime(user.previous_login);
            document.getElementById("idName").textContent = user.name;
            document.getElementById("idRole").textContent = user.role;
            document.getElementById("idNumber").textContent = user.username;

            const degrees = Math.max(14, Math.min(360, (Number(user.remaining_leave || 0) / 30) * 360));
            document.getElementById("progressRing").style.setProperty("--progress", `${degrees}deg`);

            const objectionText = t("objectionMessage")(user.name);
            document.getElementById("objectionBtn").href =
                `https://wa.me/${WHATSAPP_NUMBER}?text=${encodeURIComponent(objectionText)}`;

            drawQr({
                portal: "Personel İzin Portalı",
                id: user.username,
                name: user.name,
                role: user.role
            });

            calculateNextHoliday();
        }

        function drawQr(data) {
            const canvas = document.getElementById("qrCanvas");
            const value = JSON.stringify(data);

            if (window.QRious) {
                new QRious({
                    element: canvas,
                    value,
                    size: 156,
                    level: "M",
                    background: "white",
                    foreground: "#07263c"
                });
                return;
            }

            // CDN erişilemezse deterministik, QR-benzeri yedek görsel.
            const ctx = canvas.getContext("2d");
            const size = canvas.width;
            const cells = 21;
            const cell = size / cells;
            ctx.fillStyle = "white";
            ctx.fillRect(0, 0, size, size);

            let seed = 0;
            for (const char of value) seed = (seed * 31 + char.charCodeAt(0)) >>> 0;
            function random() {
                seed = (seed * 1664525 + 1013904223) >>> 0;
                return seed / 4294967296;
            }
            ctx.fillStyle = "#07263c";
            for (let y = 0; y < cells; y++) {
                for (let x = 0; x < cells; x++) {
                    if (random() > .53) ctx.fillRect(x * cell, y * cell, Math.ceil(cell), Math.ceil(cell));
                }
            }
            [[1,1],[13,1],[1,13]].forEach(([x,y]) => {
                ctx.fillRect(x*cell,y*cell,7*cell,7*cell);
                ctx.fillStyle = "white";
                ctx.fillRect((x+1)*cell,(y+1)*cell,5*cell,5*cell);
                ctx.fillStyle = "#07263c";
                ctx.fillRect((x+2)*cell,(y+2)*cell,3*cell,3*cell);
            });
        }

        function calculateNextHoliday() {
            const now = new Date();
            now.setHours(0,0,0,0);
            const year = now.getFullYear();

            const fixedHolidays = [
                [1,1,"Yılbaşı","New Year’s Day"],
                [4,23,"Ulusal Egemenlik ve Çocuk Bayramı","National Sovereignty and Children’s Day"],
                [5,1,"Emek ve Dayanışma Günü","Labour and Solidarity Day"],
                [5,19,"Atatürk’ü Anma, Gençlik ve Spor Bayramı","Commemoration of Atatürk, Youth and Sports Day"],
                [7,15,"Demokrasi ve Millî Birlik Günü","Democracy and National Unity Day"],
                [8,30,"Zafer Bayramı","Victory Day"],
                [10,29,"Cumhuriyet Bayramı","Republic Day"]
            ];

            let candidates = [];
            [year, year + 1].forEach(y => {
                fixedHolidays.forEach(([month, day, trName, enName]) => {
                    candidates.push({date: new Date(y, month - 1, day), trName, enName});
                });
            });
            candidates = candidates.filter(item => item.date >= now).sort((a,b) => a.date - b.date);
            const next = candidates[0];
            if (!next) return;

            const diff = Math.ceil((next.date - now) / 86400000);
            document.getElementById("holidayDays").textContent = t("holidayIn")(diff);
            document.getElementById("holidayName").textContent =
                currentLanguage === "tr" ? next.trName : next.enName;
            document.getElementById("holidayDate").textContent =
                new Intl.DateTimeFormat(currentLanguage === "tr" ? "tr-TR" : "en-GB", {
                    day: "2-digit", month: "long", year: "numeric"
                }).format(next.date);
        }

        async function installPwa() {
            if (deferredInstallPrompt) {
                deferredInstallPrompt.prompt();
                await deferredInstallPrompt.userChoice;
                deferredInstallPrompt = null;
                document.getElementById("installBtn").classList.remove("show");
                return;
            }
            const isIos = /iphone|ipad|ipod/i.test(navigator.userAgent);
            showToast(isIos ? t("iosInstall") : t("installUnavailable"));
        }

        function setupPwa() {
            if ("serviceWorker" in navigator) {
                window.addEventListener("load", () => {
                    navigator.serviceWorker.register("/service-worker.js").catch(console.error);
                });
            }

            window.addEventListener("beforeinstallprompt", event => {
                event.preventDefault();
                deferredInstallPrompt = event;
                document.getElementById("installBtn").classList.add("show");
                showToast(t("installReady"), "success");
            });

            window.addEventListener("appinstalled", () => {
                deferredInstallPrompt = null;
                document.getElementById("installBtn").classList.remove("show");
            });
        }

        document.getElementById("loginForm").addEventListener("submit", login);
        document.getElementById("themeBtn").addEventListener("click", toggleTheme);
        document.getElementById("languageBtn").addEventListener("click", () => {
            currentLanguage = currentLanguage === "tr" ? "en" : "tr";
            localStorage.setItem("izin-language", currentLanguage);
            applyLanguage();
        });
        document.getElementById("passwordToggle").addEventListener("click", () => {
            const password = document.getElementById("password");
            password.type = password.type === "password" ? "text" : "password";
        });
        document.getElementById("forgotBtn").addEventListener("click", () => openModal("forgotModal"));
        document.getElementById("leaveRequestBtn").addEventListener("click", () => openModal("leaveModal"));
        document.getElementById("installBtn").addEventListener("click", installPwa);
        document.getElementById("installActionBtn").addEventListener("click", installPwa);
        document.getElementById("skipAnimationBtn").addEventListener("click", () => finishAnimation());
        document.getElementById("replayBtn").addEventListener("click", () => playElevatorAnimation(currentUser));
        document.getElementById("logoutBtn").addEventListener("click", () => {
            animationRunId++;
            currentUser = null;
            document.getElementById("dashboard").hidden = true;
            document.getElementById("loginPanel").hidden = false;
            document.getElementById("username").focus();
        });

        document.querySelectorAll("[data-close]").forEach(button => {
            button.addEventListener("click", () => closeModal(button.dataset.close));
        });
        document.querySelectorAll(".modal-backdrop").forEach(backdrop => {
            backdrop.addEventListener("click", event => {
                if (event.target === backdrop) closeModal(backdrop.id);
            });
        });

        document.getElementById("forgotWhatsappBtn").addEventListener("click", () => {
            const identity = document.getElementById("forgotIdentity").value.trim();
            if (!identity) {
                showToast(t("fillIdentity"), "error");
                return;
            }
            openWhatsApp(t("forgotMessage")(identity));
            closeModal("forgotModal");
        });

        document.getElementById("sendLeaveBtn").addEventListener("click", () => {
            if (!currentUser) return;
            const start = document.getElementById("leaveStart").value;
            const end = document.getElementById("leaveEnd").value;
            if (!start || !end) {
                showToast(t("fillDates"), "error");
                return;
            }

            const startDate = new Date(`${start}T00:00:00`);
            const endDate = new Date(`${end}T00:00:00`);
            if (endDate < startDate) {
                showToast(t("invalidDates"), "error");
                return;
            }

            const days = Math.floor((endDate - startDate) / 86400000) + 1;
            const data = {
                type: document.getElementById("leaveType").value,
                start: new Intl.DateTimeFormat(currentLanguage === "tr" ? "tr-TR" : "en-GB").format(startDate),
                end: new Intl.DateTimeFormat(currentLanguage === "tr" ? "tr-TR" : "en-GB").format(endDate),
                days,
                description: document.getElementById("leaveDescription").value.trim()
            };
            openWhatsApp(t("leaveMessage")(data));
            closeModal("leaveModal");
        });

        window.addEventListener("keydown", event => {
            if (event.key === "Escape") {
                document.querySelectorAll(".modal-backdrop.open").forEach(item => closeModal(item.id));
            }
        });

        autoTheme();
        applyLanguage();
        setupPwa();
    </script>
</body>
</html>'''


SERVICE_WORKER = r'''const CACHE_NAME = "izin-portali-v3";
const APP_SHELL = ["/", "/manifest.webmanifest", "/icon-192.png", "/icon-512.png"];

self.addEventListener("install", event => {
    event.waitUntil(
        caches.open(CACHE_NAME)
            .then(cache => cache.addAll(APP_SHELL))
            .then(() => self.skipWaiting())
    );
});

self.addEventListener("activate", event => {
    event.waitUntil(
        caches.keys()
            .then(keys => Promise.all(keys.filter(key => key !== CACHE_NAME).map(key => caches.delete(key))))
            .then(() => self.clients.claim())
    );
});

self.addEventListener("fetch", event => {
    const request = event.request;
    if (request.method !== "GET") return;

    const url = new URL(request.url);
    if (url.origin !== self.location.origin) return;
    if (url.pathname === "/login") return;

    if (request.mode === "navigate") {
        event.respondWith(
            fetch(request)
                .then(response => {
                    const copy = response.clone();
                    caches.open(CACHE_NAME).then(cache => cache.put("/", copy));
                    return response;
                })
                .catch(() => caches.match("/"))
        );
        return;
    }

    event.respondWith(
        caches.match(request).then(cached => {
            if (cached) return cached;
            return fetch(request).then(response => {
                if (response.ok) {
                    const copy = response.clone();
                    caches.open(CACHE_NAME).then(cache => cache.put(request, copy));
                }
                return response;
            });
        })
    );
});'''


@app.after_request
def add_security_headers(response):
    response.headers["X-Content-Type-Options"] = "nosniff"
    response.headers["X-Frame-Options"] = "DENY"
    response.headers["Referrer-Policy"] = "strict-origin-when-cross-origin"
    response.headers["Permissions-Policy"] = "camera=(), microphone=(), geolocation=()"
    return response


@app.route("/")
def index():
    response = Response(HTML_SAYFASI, mimetype="text/html")
    response.headers["Cache-Control"] = "no-cache"
    return response


@app.route("/login", methods=["POST"])
def login():
    ip_address = client_ip()
    blocked, seconds = login_is_blocked(ip_address)
    if blocked:
        minutes = max(1, math.ceil(seconds / 60))
        return jsonify({
            "status": "error",
            "message": f"Çok fazla hatalı deneme. Yaklaşık {minutes} dakika sonra tekrar deneyin."
        }), 429

    payload = request.get_json(silent=True) or {}
    username = clean_scalar(payload.get("username"))
    password = clean_scalar(payload.get("password"))

    if not username or not password:
        return jsonify({
            "status": "error",
            "message": "Kullanıcı adı ve şifre zorunludur."
        }), 400

    try:
        user_data = get_user_data(username, password)
    except (FileNotFoundError, ValueError) as error:
        app.logger.error("Yapılandırma hatası: %s", error)
        return jsonify({
            "status": "error",
            "message": "Sistem yapılandırması kontrol edilmelidir."
        }), 500
    except Exception:
        app.logger.exception("Excel okuma veya giriş işlemi sırasında hata")
        return jsonify({
            "status": "error",
            "message": "Giriş işlemi sırasında beklenmeyen bir hata oluştu."
        }), 500

    if user_data is None:
        record_failed_login(ip_address)
        return jsonify({
            "status": "error",
            "message": "Kullanıcı adı veya şifre hatalı."
        }), 401

    clear_failed_logins(ip_address)
    response = jsonify({"status": "success", "data": user_data})
    response.headers["Cache-Control"] = "no-store"
    return response


@app.route("/manifest.webmanifest")
def manifest():
    return jsonify({
        "name": "Personel İzin Portalı",
        "short_name": "İzin Portalı",
        "description": "Personel izin bilgileri, itiraz ve izin talep sistemi",
        "lang": "tr",
        "start_url": "/",
        "scope": "/",
        "display": "standalone",
        "orientation": "portrait-primary",
        "background_color": "#061426",
        "theme_color": "#071b30",
        "icons": [
            {"src": "/icon-192.png", "sizes": "192x192", "type": "image/png", "purpose": "any maskable"},
            {"src": "/icon-512.png", "sizes": "512x512", "type": "image/png", "purpose": "any maskable"}
        ],
        "shortcuts": [
            {
                "name": "İzin Portalını Aç",
                "short_name": "Portal",
                "url": "/"
            }
        ]
    })


@app.route("/service-worker.js")
def service_worker():
    response = Response(SERVICE_WORKER, mimetype="application/javascript")
    response.headers["Cache-Control"] = "no-cache"
    response.headers["Service-Worker-Allowed"] = "/"
    return response


@app.route("/icon-192.png")
def icon_192():
    return Response(base64.b64decode(ICON_192_B64), mimetype="image/png")


@app.route("/icon-512.png")
def icon_512():
    return Response(base64.b64decode(ICON_512_B64), mimetype="image/png")


@app.route("/health")
def health():
    return jsonify({"status": "ok"})


if __name__ == "__main__":
    port = int(os.getenv("PORT", "5000"))
    app.run(host="0.0.0.0", port=port, debug=False)
