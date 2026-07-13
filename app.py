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

import struct
import zlib


def _make_png_icon(size, bg=(21, 81, 146), fg=(238, 251, 255)):
    """Bağımlılık gerektirmeyen, geçerli bir PNG ikon üretir (marka renkli, 'İ' harfli)."""
    # Basit bir tuval: arka plan + ortada dikey bir 'çubuk' (İ hissi veren sade işaret).
    px = [[bg[0], bg[1], bg[2]] for _ in range(size * size)]

    bar_w = max(2, size // 9)
    bar_h = int(size * 0.46)
    cx = size // 2
    top = int(size * 0.30)
    dot_r = max(2, size // 14)
    dot_cy = int(size * 0.20)

    def put(x, y, c):
        if 0 <= x < size and 0 <= y < size:
            px[y * size + x] = [c[0], c[1], c[2]]

    # Gövde çubuğu
    for y in range(top, top + bar_h):
        for x in range(cx - bar_w // 2, cx - bar_w // 2 + bar_w):
            put(x, y, fg)
    # Üst/alt serifler
    serif_w = bar_w * 3
    for x in range(cx - serif_w // 2, cx + serif_w // 2):
        for t in range(max(2, bar_w // 2)):
            put(x, top + t, fg)
            put(x, top + bar_h - 1 - t, fg)
    # Noktalı 'İ' için üst nokta
    for y in range(dot_cy - dot_r, dot_cy + dot_r):
        for x in range(cx - dot_r, cx + dot_r):
            if (x - cx) ** 2 + (y - dot_cy) ** 2 <= dot_r * dot_r:
                put(x, y, fg)

    # Ham RGB -> PNG (her satırın başına filtre byte'ı 0)
    raw = bytearray()
    for y in range(size):
        raw.append(0)
        for x in range(size):
            raw.extend(px[y * size + x])

    def chunk(tag, data):
        c = tag + data
        return struct.pack(">I", len(data)) + c + struct.pack(">I", zlib.crc32(c) & 0xFFFFFFFF)

    sig = b"\x89PNG\r\n\x1a\n"
    ihdr = struct.pack(">IIBBBBB", size, size, 8, 2, 0, 0, 0)  # 8-bit, truecolor
    idat = zlib.compress(bytes(raw), 9)
    return sig + chunk(b"IHDR", ihdr) + chunk(b"IDAT", idat) + chunk(b"IEND", b"")


ICON_192_BYTES = _make_png_icon(192)
ICON_512_BYTES = _make_png_icon(512)


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
    <script src="https://cdn.jsdelivr.net/npm/qrious@4.0.2/dist/qrious.min.js" defer></script>
    <style>
        :root{
            --bg-1:#061426; --bg-2:#0b3152;
            --panel:rgba(11,38,62,.94); --panel-strong:rgba(6,26,46,.94);
            --text:#f4fbff; --muted:#a8c4d8;
            --primary:#55d9ff; --primary-2:#1489c9;
            --brand:#155192; --brand-2:#1e6fbf;
            --success:#2bd881; --warning:#ffca57; --danger:#ff6678;
            --line:rgba(143,222,255,.22);
            --shadow:0 28px 70px rgba(0,0,0,.34); --radius:24px;
        }
        body.light{
            --bg-1:#dff4ff; --bg-2:#eff9ff;
            --panel:rgba(255,255,255,.96); --panel-strong:rgba(255,255,255,.96);
            --text:#10283d; --muted:#527089;
            --primary:#087fc4; --primary-2:#23b5e8;
            --line:rgba(14,116,174,.16); --shadow:0 28px 70px rgba(39,96,131,.18);
        }
        *{box-sizing:border-box;}
        html{min-height:100%; background:#061426; overflow-x:hidden;}
        html.light{background:#eff9ff;}
        body{
            margin:0; min-height:100vh; min-height:100dvh; width:100%; max-width:100%; overflow-x:hidden; color:var(--text);
            font-family:Inter,"Segoe UI",system-ui,-apple-system,BlinkMacSystemFont,sans-serif;
            background:
                radial-gradient(circle at 12% 12%, rgba(63,196,255,.18), transparent 34%),
                radial-gradient(circle at 88% 80%, rgba(52,224,174,.11), transparent 28%),
                linear-gradient(145deg, var(--bg-1), var(--bg-2));
            overflow-x:hidden; transition:background .35s ease, color .35s ease;
        }
        button,input,select,textarea{font:inherit;}
        button,a{-webkit-tap-highlight-color:transparent;}
        button{color:inherit;}
        [hidden]{display:none !important;}

        .topbar{
            position:fixed; z-index:60; inset:max(12px,env(safe-area-inset-top)) 14px auto 14px;
            display:flex; justify-content:space-between; align-items:center; pointer-events:none;
        }
        .topbar::before{
            content:""; position:fixed; left:0; right:0; top:0; z-index:-1; pointer-events:none;
            height:calc(env(safe-area-inset-top) + 74px);
            background:linear-gradient(180deg, var(--bg-1) 40%, transparent);
        }
        .brand-mini,.top-actions{pointer-events:auto; display:flex; align-items:center; gap:8px;}
        .brand-mini{
            padding:9px 13px; border:1px solid var(--line); border-radius:16px;
            background:var(--panel);
            box-shadow:0 12px 30px rgba(0,0,0,.14); font-weight:800; letter-spacing:.03em;
        }
        .brand-mark{
            width:28px; height:28px; border-radius:9px; display:grid; place-items:center;
            background:linear-gradient(145deg,var(--primary),var(--primary-2));
            color:#042039; font-weight:1000;
        }
        .icon-btn,.install-btn{
            border:1px solid var(--line); background:var(--panel);
            border-radius:14px; min-height:42px; padding:0 12px; cursor:pointer;
            box-shadow:0 12px 30px rgba(0,0,0,.12);
            transition:transform .2s ease, border-color .2s ease;
        }
        .icon-btn:hover,.install-btn:hover{transform:translateY(-2px); border-color:var(--primary);}
        .install-btn{display:none; font-weight:800; gap:7px; align-items:center;}
        .install-btn.show{display:inline-flex;}

        .page-shell{
            min-height:100vh; min-height:100dvh; display:grid; place-items:center;
            padding:86px 18px 28px; position:relative; z-index:2;
        }
        .glass{
            border:1px solid var(--line); background:var(--panel);
            
            box-shadow:var(--shadow); border-radius:var(--radius);
        }

        /* ---------- LOGIN ---------- */
        .login-card{
            width:min(440px,100%); padding:clamp(24px,6vw,40px); position:relative; overflow:hidden;
            animation:enterCard .6s cubic-bezier(.2,.8,.2,1) both;
        }
        @keyframes enterCard{from{opacity:0; transform:translateY(22px) scale(.97);}to{opacity:1; transform:none;}}
        /* B) hatalı şifrede sallanma */
        .login-card.shake{animation:cardShake .5s cubic-bezier(.36,.07,.19,.97);}
        @keyframes cardShake{
            10%,90%{transform:translateX(-2px);} 20%,80%{transform:translateX(4px);}
            30%,50%,70%{transform:translateX(-9px);} 40%,60%{transform:translateX(9px);}
        }
        /* 76) başarıda yeşil onay dalgası */
        .login-card .success-wave{
            position:absolute; inset:0; border-radius:var(--radius); pointer-events:none; opacity:0;
            background:radial-gradient(circle at 50% 62%, rgba(43,216,129,.55), rgba(43,216,129,0) 60%);
        }
        .login-card.success .success-wave{animation:successWave .9s ease forwards;}
        @keyframes successWave{0%{opacity:0; transform:scale(.6);}35%{opacity:1;}100%{opacity:0; transform:scale(1.5);}}
        .login-logo{
            width:78px; height:78px; margin:0 auto 18px; border-radius:24px; display:grid; place-items:center;
            background:linear-gradient(145deg,rgba(76,216,255,.22),rgba(18,103,164,.42));
            border:1px solid rgba(111,225,255,.4);
            box-shadow:inset 0 0 30px rgba(79,218,255,.12), 0 14px 34px rgba(0,0,0,.18);
            animation:logoBreath 3.4s ease-in-out infinite;
        }
        @keyframes logoBreath{0%,100%{box-shadow:inset 0 0 30px rgba(79,218,255,.12),0 14px 34px rgba(0,0,0,.18);}50%{box-shadow:inset 0 0 42px rgba(79,218,255,.32),0 14px 40px rgba(0,0,0,.2);}}
        .elevator-icon{width:42px; height:48px; border:2px solid var(--primary); border-radius:7px; position:relative; overflow:hidden;}
        .elevator-icon::before,.elevator-icon::after{content:""; position:absolute; top:0; bottom:0; width:50%; background:rgba(82,217,255,.14);}
        .elevator-icon::before{left:0; border-right:1px solid var(--primary);}
        .elevator-icon::after{right:0;}
        .login-card h1{text-align:center; margin:0; font-size:clamp(1.55rem,5vw,2.05rem); letter-spacing:-.03em;}
        .login-subtitle{text-align:center; color:var(--muted); margin:9px 0 28px; line-height:1.5;}
        .field{margin-bottom:16px;}
        .field label{display:block; margin-bottom:8px; color:var(--muted); font-size:.9rem; font-weight:700;}
        .input-wrap{position:relative; border-radius:15px; transition:transform .2s ease, box-shadow .2s ease;}
        /* 106) odak animasyonu */
        .input-wrap:focus-within{transform:translateY(-1px); box-shadow:0 0 0 4px rgba(75,207,255,.16);}
        .input-wrap input,.modal input,.modal select,.modal textarea{
            width:100%; color:var(--text); border:1px solid var(--line);
            background:rgba(255,255,255,.055); border-radius:15px; padding:15px 46px 15px 15px; outline:none;
            transition:border-color .2s ease, box-shadow .2s ease, background .2s ease;
        }
        body.light .input-wrap input,body.light .modal input,body.light .modal select,body.light .modal textarea{background:rgba(8,70,110,.04);}
        input::placeholder,textarea::placeholder{color:color-mix(in srgb, var(--muted) 75%, transparent);}
        .input-wrap input:focus,.modal input:focus,.modal select:focus,.modal textarea:focus{border-color:var(--primary); background:rgba(255,255,255,.08);}
        .field-icon{position:absolute; right:14px; top:50%; transform:translateY(-50%); color:var(--muted); pointer-events:none;}
        .password-toggle{pointer-events:auto; border:0; background:transparent; padding:4px; cursor:pointer;}
        /* 95) caps lock uyarısı */
        .caps-hint{display:none; margin-top:8px; color:var(--warning); font-size:.78rem; font-weight:800;}
        .caps-hint.show{display:block; animation:enterCard .25s ease both;}
        .primary-btn,.action-btn{
            width:100%; border:0; border-radius:15px; padding:15px 18px; font-weight:900; cursor:pointer;
            display:inline-flex; justify-content:center; align-items:center; gap:9px; text-decoration:none;
            position:relative; overflow:hidden; transition:transform .2s ease, filter .2s ease;
        }
        .primary-btn{background:linear-gradient(135deg,var(--primary),var(--primary-2)); color:#032039; box-shadow:0 14px 30px rgba(34,172,229,.28);}
        .primary-btn:hover,.action-btn:hover{transform:translateY(-2px); filter:brightness(1.05);}
        .primary-btn:active,.action-btn:active{transform:translateY(1px) scale(.99);}
        .primary-btn:disabled{opacity:.65; cursor:wait;}
        /* 3) buton ripple */
        .ripple{position:absolute; border-radius:50%; transform:translate(-50%,-50%) scale(0); background:rgba(255,255,255,.5); pointer-events:none; animation:rippleGo .6s ease-out forwards;}
        @keyframes rippleGo{to{transform:translate(-50%,-50%) scale(1); opacity:0;}}
        .spinner{width:19px; height:19px; border-radius:50%; border:2px solid rgba(3,32,57,.25); border-top-color:#032039; animation:spin .8s linear infinite; display:none;}
        .loading .spinner{display:block;}
        @keyframes spin{to{transform:rotate(360deg);}}
        .remember-row{display:flex; align-items:center; gap:9px; margin:2px 0 16px; color:var(--muted); font-size:.86rem; font-weight:700;}
        .remember-row input{width:18px; height:18px; accent-color:var(--primary);}
        .login-links{display:flex; justify-content:center; margin-top:17px;}
        .text-btn{border:0; background:transparent; color:var(--primary); cursor:pointer; font-weight:800; padding:7px;}
        .sys-status{display:flex; align-items:center; justify-content:center; gap:7px; margin-top:14px; color:var(--muted); font-size:.76rem; font-weight:700;}
        .sys-dot{width:8px; height:8px; border-radius:50%; background:var(--success); box-shadow:0 0 10px var(--success); animation:sysPulse 2s ease-in-out infinite;}
        @keyframes sysPulse{0%,100%{opacity:.5;}50%{opacity:1;}}

        /* A) giriş kapıları (login üstüne kapanır, sim açılınca aralanır) — ilk render jank'ini de gizler */
        .entry-doors{position:fixed; inset:0; z-index:150; pointer-events:none; display:none;}
        .entry-doors.active{display:block;}
        .entry-doors .ed{
            position:absolute; top:0; bottom:0; width:52%;
            background:
                repeating-linear-gradient(90deg, rgba(255,255,255,.05) 0 2px, transparent 2px 16px),
                linear-gradient(90deg,#243441,#66788a 48%,#28394a);
            box-shadow:inset 0 0 60px rgba(0,0,0,.5);
            transform:translateX(-100%); transition:transform 1s cubic-bezier(.7,.02,.18,1); will-change:transform;
        }
        .entry-doors .ed.l{left:0;} .entry-doors .ed.r{right:0; transform:translateX(100%); background:linear-gradient(90deg,#28394a,#66788a 52%,#243441);}
        .entry-doors.closing .ed.l{transform:translateX(0);} .entry-doors.closing .ed.r{transform:translateX(0);}

        /* ============================================================
           KESİNTİSİZ ASANSÖR SİMÜLASYONU
           - Alttaki katmanlar daima opak; üstteki söner -> siyah ekran yok
           - Kamera asla sağa-sola dönmez
           - GPU katmanı (translateZ/ will-change) -> render takılması yok
           ============================================================ */
        .sim{position:fixed; inset:0; z-index:100; overflow:hidden; background:#0c141c; color:#eef7ff; font-family:Inter,"Segoe UI",system-ui,sans-serif;}
        .sim[hidden]{display:none !important;}
        .sim *{box-sizing:border-box;}
        .sim-world{position:absolute; inset:0; transform-origin:50% 55%; transition:transform 1.6s cubic-bezier(.22,.78,.18,1); will-change:transform; transform:translateZ(0);}

        .stage{position:absolute; inset:0; transition:opacity 1.1s ease, transform 1.6s cubic-bezier(.22,.78,.18,1); will-change:opacity,transform; transform:translateZ(0); backface-visibility:hidden;}
        .stage.office{z-index:1; opacity:1;}
        .stage.cabin{z-index:2; opacity:1;}
        .stage.lobby{z-index:3; opacity:1;}

        /* ---- ORTAK KAPI ---- */
        .doorway{position:absolute; overflow:hidden; background:transparent;}
        .door{
            position:absolute; top:0; bottom:0; width:50.5%;
            background:
                repeating-linear-gradient(90deg, rgba(255,255,255,.05) 0 2px, transparent 2px 15px),
                linear-gradient(90deg,#2b3c47,#7c8f9a 48%,#2e3f4a);
            box-shadow:inset 0 0 30px rgba(0,0,0,.42);
            transition:transform 2.1s cubic-bezier(.7,.02,.18,1); will-change:transform;
        }
        .door.l{left:0; border-right:1px solid #0b1217;}
        .door.r{right:0; border-left:1px solid rgba(255,255,255,.16); transform:scaleX(-1);}
        /* 41/çelik: kenar highlight */
        .door::after{content:""; position:absolute; top:0; bottom:0; right:0; width:3px; background:linear-gradient(rgba(255,255,255,.5),rgba(255,255,255,.05));}

        /* ---- LOBİ ---- */
        .stage.lobby .lobby-wall{position:absolute; inset:0; background:linear-gradient(180deg,#e7eef2 0 16%,#cfd9df 16% 74%,#7c8892 74% 100%);}
        .stage.lobby .lobby-wall::after{content:""; position:absolute; left:-10%; right:-10%; bottom:-18%; height:46%; background:repeating-linear-gradient(90deg, rgba(24,49,63,.16) 0 1px, transparent 1px 12%),linear-gradient(#8a97a0,#d3dde1); transform:perspective(620px) rotateX(62deg); transform-origin:top;}
        .lobby-frame{position:absolute; left:50%; top:9%; bottom:9%; width:min(72vw,560px); transform:translateX(-50%); border:clamp(10px,1.6vw,20px) solid #56646f; border-bottom-width:26px; background:#0a1016; box-shadow:0 22px 55px rgba(0,0,0,.35), inset 0 0 0 2px rgba(255,255,255,.12);}
        .lobby-head{position:absolute; left:50%; top:-46px; transform:translateX(-50%); min-width:110px; padding:8px 16px; border-radius:8px; text-align:center; color:#6fe6ff; background:#03080d; border:1px solid #486370; font:900 1.15rem/1 ui-monospace,monospace; letter-spacing:.18em; text-shadow:0 0 12px rgba(91,226,255,.7);}
        .lobby-doorway{inset:0;}
        .sim.lobby-open .stage.lobby .door.l{transform:translateX(-102%);}
        .sim.lobby-open .stage.lobby .door.r{transform:scaleX(-1) translateX(-102%);}
        .sim.stepped .stage.lobby{opacity:0; transform:scale(1.35);}

        /* ---- KABİN ---- */
        .stage.cabin{background:linear-gradient(90deg,#0a1620 0 20%,#1a2b37 20% 80%,#0a1620 80%),#0d1a24;}
        .cabin-ceil{position:absolute; left:-6%; right:-6%; top:-10%; height:30%; background:linear-gradient(#1a2a34,#0a141c); transform:perspective(680px) rotateX(-52deg); box-shadow:0 30px 60px rgba(0,0,0,.5);}
        /* 52) tavan gömme spotları + ışık havuzu */
        .spot{position:absolute; top:6%; width:clamp(26px,6vw,54px); height:8px; border-radius:50%; background:#eafcff; box-shadow:0 0 22px #bdf1ff,0 0 60px rgba(120,225,255,.4);}
        .spot.s1{left:26%;} .spot.s2{left:50%; transform:translateX(-50%);} .spot.s3{right:26%;}
        .spot::after{content:""; position:absolute; left:50%; top:100%; width:220%; height:26vh; transform:translateX(-50%); background:radial-gradient(ellipse at top, rgba(190,240,255,.12), transparent 60%); pointer-events:none;}
        .cabin-light-wrap{position:absolute; inset:0; transition:opacity .25s; }
        /* 12) kalkışta ışık titremesi (bir kez) */
        .sim.flicker .stage.cabin{animation:cabinFlicker .7s ease 1;}
        @keyframes cabinFlicker{0%,100%{filter:none;}20%{filter:brightness(.6);}30%{filter:brightness(1.15);}45%{filter:brightness(.75);}60%{filter:brightness(1);}}
        .cabin-floor{position:absolute; left:-10%; right:-10%; bottom:-14%; height:34%; background:repeating-linear-gradient(90deg, rgba(140,200,225,.10) 0 1px, transparent 1px 9%),linear-gradient(#16303f,#0a151d); transform:perspective(560px) rotateX(60deg); transform-origin:top;}
        .stage.cabin .wall{position:absolute; top:12%; bottom:0; width:20%; background:linear-gradient(90deg,rgba(255,255,255,.04),transparent),repeating-linear-gradient(90deg,rgba(180,210,225,.05) 0 1px,transparent 1px 6px),linear-gradient(#12242f,#0a151d); border:1px solid rgba(150,215,240,.10);}
        .stage.cabin .wall.left{left:0; clip-path:polygon(0 0,100% 9%,100% 100%,0 100%);}
        .stage.cabin .wall.right{right:0; transform:scaleX(-1); clip-path:polygon(0 0,100% 9%,100% 100%,0 100%);}

        /* C) MAX 8 KİŞİ · 630 KG tabelası + 89) acil buton/interkom + 111) havalandırma */
        .cabin-plate{position:absolute; z-index:16; left:3%; top:20%; width:min(20vw,120px); padding:7px 9px; border-radius:7px; background:linear-gradient(#20323d,#0d1a22); border:1px solid rgba(160,210,230,.3); color:#bcd4de; font-size:clamp(.5rem,1.7vw,.66rem); font-weight:900; letter-spacing:.06em; line-height:1.35; text-align:center; box-shadow:inset 0 0 8px rgba(0,0,0,.4);}

        /* Kat göstergesi (kapı üstünde) — 39 LED font, 11 ok, 30 flash, ding */
        .floor-ind{position:absolute; z-index:20; left:50%; top:max(3.5%,env(safe-area-inset-top)); transform:translateX(-50%); min-width:clamp(140px,40vw,190px); padding:8px 16px 7px; text-align:center; border-radius:11px; color:#ffb14a; background:#0a0602; border:1px solid #5a4326; box-shadow:inset 0 0 16px rgba(255,150,40,.14),0 0 22px rgba(255,150,40,.12);}
        .floor-ind .fi-top{display:flex; align-items:center; justify-content:center; gap:8px;}
        .fi-arrow{color:#ffcf6a; font-size:1rem; opacity:.35; transition:opacity .2s;}
        .sim.riding .fi-arrow{animation:arrowPulse .9s ease-in-out infinite;}
        @keyframes arrowPulse{0%,100%{opacity:.25;}50%{opacity:1; text-shadow:0 0 12px #ffcf6a;}}
        .floor-ind b{font-family:"DS-Digital",ui-monospace,"Courier New",monospace; font-size:clamp(1.5rem,6.5vw,2.4rem); line-height:1; letter-spacing:.14em; font-weight:800; text-shadow:0 0 10px rgba(255,170,60,.75),0 0 2px rgba(255,120,20,.9);}
        .floor-ind small{display:block; margin-top:3px; color:#b98b52; font-size:.6rem; letter-spacing:.24em; font-weight:900;}
        .floor-ind.flash{animation:indFlash .5s ease;}
        @keyframes indFlash{0%,100%{background:#0a0602;}50%{background:#3a2a08; box-shadow:0 0 34px rgba(255,180,70,.6);}}
        .floor-ind.ding{animation:dingPulse 1s ease;}
        @keyframes dingPulse{0%{box-shadow:inset 0 0 16px rgba(255,150,40,.14),0 0 22px rgba(255,150,40,.12);}30%{box-shadow:inset 0 0 22px rgba(255,190,90,.4),0 0 52px rgba(255,180,70,.85); transform:translateX(-50%) scale(1.05);}100%{transform:translateX(-50%) scale(1);}}


        /* Kabin kapısı + cam (61) */
        .cabin-doorway{left:24%; right:24%; top:16%; bottom:2%; border:clamp(7px,1.4vw,16px) solid #3b5361; border-bottom-width:20px; border-radius:2px; box-shadow:inset 0 0 0 2px rgba(255,255,255,.1),0 18px 45px rgba(0,0,0,.5);}
        .door-glass{position:absolute; top:8%; bottom:26%; width:22%; border-radius:4px; background:linear-gradient(180deg,rgba(120,180,210,.18),rgba(30,60,80,.28)); border:1px solid rgba(160,210,235,.3); overflow:hidden;}
        .door .door-glass{left:50%; transform:translateX(-50%);}
        .door-glass::after{content:""; position:absolute; left:0; right:0; height:200%; background:repeating-linear-gradient(180deg, transparent 0 30px, rgba(10,20,28,.55) 30px 46px);}
        .sim.riding .door-glass::after{animation:passFloors .5s linear infinite;}
        .sim.riding.fast .door-glass::after{animation-duration:.24s;}
        @keyframes passFloors{from{transform:translateY(-46px);}to{transform:translateY(0);}}
        /* 23) kapı açılırken ışık sızması */
        .door-leak{position:absolute; z-index:19; left:50%; top:16%; bottom:2%; width:2px; transform:translateX(-50%); background:linear-gradient(180deg,transparent,rgba(255,247,220,.9),transparent); opacity:0; filter:blur(1px);}
        .sim.arrived .door-leak{animation:leak 1.1s ease forwards;}
        @keyframes leak{0%{opacity:0; width:2px;}30%{opacity:1; width:5px;}100%{opacity:0; width:26vw;}}

        .sim.arrived .stage.cabin .cabin-doorway .door.l{transform:translateX(-102%);}
        .sim.arrived .stage.cabin .cabin-doorway .door.r{transform:scaleX(-1) translateX(-102%);}
        .sim.exiting .stage.cabin{opacity:0; transform:scale(1.28);}

        /* 47) gerçek zamanlı dijital saat */

        /* Departman paneli — içeride, geniş, yazılar tam */
        .dept-panel{position:absolute; z-index:22; right:12%; top:19%; width:min(40vw,270px); padding:12px; border-radius:15px; background:linear-gradient(150deg,#43576a,#17262f 55%,#0c1820); border:1px solid rgba(200,235,250,.26); box-shadow:0 20px 40px rgba(0,0,0,.34), inset 0 0 20px rgba(255,255,255,.04); transition:opacity .8s ease; transform:translateZ(0); will-change:opacity;}
        .dept-title{margin-bottom:9px; color:#a7bfcb; text-align:center; font-size:.62rem; font-weight:1000; letter-spacing:.18em;}
        .dept-btn{position:relative; overflow:hidden; width:100%; min-height:46px; margin:6px 0; padding:9px 11px; border-radius:10px; display:flex; align-items:center; gap:9px; text-align:left; color:#dce9ee; background:linear-gradient(#1b2b36,#0c161d); border:1px solid #5c7280; box-shadow:inset 0 0 10px rgba(0,0,0,.42); font-size:clamp(.72rem,2.4vw,.86rem); line-height:1.2; font-weight:850; transition:transform .12s, box-shadow .2s, background .2s;}
        .dept-num{flex:0 0 22px; width:22px; height:22px; border-radius:50%; display:grid; place-items:center; background:#263840; border:1px solid #7a8f98; color:#cfe0e8; font-weight:1000; font-size:.72rem;}
        /* 3) basılınca çöküş + ışık halkası; picking veya kalıcı .lit */
        .dept-btn.pressed{transform:scale(.96);}
        .sim.picking .dept-btn.target,.dept-btn.target.lit{color:#06131a; background:linear-gradient(#b8f5ff,#4ed7f4); border-color:#d7fbff; box-shadow:0 0 24px rgba(71,220,251,.7), inset 0 0 12px rgba(255,255,255,.7);}
        .sim.picking .dept-btn.target .dept-num,.dept-btn.target.lit .dept-num{background:#fff; border-color:#fff; color:#0a5279; box-shadow:0 0 12px #fff;}
        .dept-btn .halo{position:absolute; left:18px; top:50%; width:8px; height:8px; border-radius:50%; transform:translate(-50%,-50%) scale(0); background:rgba(120,240,255,.6); pointer-events:none;}
        .dept-btn.pressed .halo{animation:haloGo .6s ease-out;}
        @keyframes haloGo{to{transform:translate(-50%,-50%) scale(9); opacity:0;}}

        /* ---- OFİS ---- */
        .stage.office{background:linear-gradient(180deg,#eaf4f8 0 18%,#cfe1e9 18% 74%,#9db2bd 74% 100%); transform:scale(.9);}
        .sim.exiting .stage.office,.sim.in-office .stage.office{transform:scale(1);}
        .sim.reading .stage.office{transform:scale(1.06);}
        .office-ceil{position:absolute; left:-8%; right:-8%; top:-12%; height:32%; background:linear-gradient(#f4fafc,#d3dfe4); transform:perspective(680px) rotateX(-52deg);}
        .office-floor{position:absolute; left:-14%; right:-14%; bottom:-20%; height:52%; background:repeating-linear-gradient(90deg, rgba(27,75,94,.12) 0 1px, transparent 1px 8%),linear-gradient(#9cb9c5,#e2edf1); transform:perspective(560px) rotateX(63deg); transform-origin:top;}
        .office-wall{position:absolute; top:20%; bottom:16%; width:22%; border:2px solid rgba(51,108,133,.3); background:linear-gradient(135deg,rgba(255,255,255,.4),rgba(111,187,218,.1));}
        .office-wall.left{left:2%;} .office-wall.right{right:2%;}
        /* 15) 67. kata özel ışıklı tabela */
        .office-sign{position:absolute; z-index:5; left:50%; top:9.5%; transform:translateX(-50%); white-space:nowrap; text-align:center; padding:9px 20px; border-radius:12px; color:#eaf7ff; background:linear-gradient(#1c4a66,#0f3550); border:1px solid rgba(160,220,245,.55); box-shadow:0 10px 26px rgba(9,40,58,.35), inset 0 1px 0 rgba(255,255,255,.22); font-weight:1000; letter-spacing:.1em;}
        .office-sign .os-main{font-size:clamp(.85rem,3vw,1.45rem);}
        .office-sign .os-floor{display:block; margin-top:2px; font-size:clamp(.6rem,2vw,.85rem); color:#7fe6ff; text-shadow:0 0 12px rgba(90,220,255,.7);}

        /* 5 kişilik ekip — sade, orantılı figürler (kol/bacak yok) */
        .crew-row{position:absolute; z-index:5; left:0; right:0; bottom:15%; display:flex; justify-content:center; align-items:flex-end; gap:clamp(8px,3vw,26px);}
        .crew{position:relative; width:clamp(44px,12vw,78px); height:var(--ch,150px); opacity:0; transform:translateY(20px); transition:opacity .5s ease, transform .7s cubic-bezier(.2,.8,.2,1), filter .5s ease; will-change:transform,opacity;}
        .crew:nth-child(1){--ch:clamp(120px,24vh,168px);}
        .crew:nth-child(2){--ch:clamp(168px,34vh,240px);}   /* arkada duran — baya uzun boylu */
        .crew:nth-child(3){--ch:clamp(128px,25vh,178px);}   /* Semih — normal boy */
        .crew:nth-child(4){--ch:clamp(124px,24vh,172px);}
        .crew:nth-child(5){--ch:clamp(130px,25vh,182px);}
        .sim.in-office .crew{opacity:1; transform:none;}
        .sim.in-office .crew:nth-child(1){transition-delay:.05s;}
        .sim.in-office .crew:nth-child(2){transition-delay:.15s;}
        .sim.in-office .crew:nth-child(3){transition-delay:.25s;}
        .sim.in-office .crew:nth-child(4){transition-delay:.35s;}
        .sim.in-office .crew:nth-child(5){transition-delay:.45s;}
        .crew .c-shadow{position:absolute; left:50%; bottom:-7px; transform:translateX(-50%); width:82%; height:9px; border-radius:50%; background:radial-gradient(closest-side, rgba(0,0,0,.32), transparent 72%);}
        .crew .c-head{position:absolute; left:50%; top:0; transform:translateX(-50%); width:42%; aspect-ratio:1; border-radius:50%; background:linear-gradient(90deg,#c68a5e,var(--skin,#e6b183) 55%,#cf9366);}
        .crew .c-hair{position:absolute; left:50%; top:-2%; transform:translateX(-50%); width:47%; height:26%; border-radius:50% 50% 42% 42%; background:var(--hair,#241a14);}
        .crew .c-body{position:absolute; left:50%; bottom:0; transform:translateX(-50%); width:100%; height:64%; border-radius:46% 46% 18% 18% / 62% 62% 14% 14%; background:var(--suit,#2b5170); box-shadow:inset 0 8px 14px rgba(255,255,255,.1), inset 0 -8px 14px rgba(0,0,0,.15);}
        /* Semih öne çıkar ve sahnede biraz durur (hafif salınım) */
        .sim.crew-step .crew{filter:brightness(.82);}
        .sim.crew-step .crew.semih{filter:none; z-index:9; animation:semihIn 4.8s ease forwards;}
        @keyframes semihIn{
            0%{transform:translateY(0) scale(1);}
            16%{transform:translateY(42%) scale(1.5);}
            40%{transform:translateY(39%) scale(1.5);}
            62%{transform:translateY(43%) scale(1.5);}
            84%{transform:translateY(39.5%) scale(1.5);}
            100%{transform:translateY(42%) scale(1.5);}
        }
        /* isim rozeti */
        .name-badge{position:absolute; z-index:10; left:50%; bottom:7%; transform:translate(-50%,14px); padding:9px 20px; border-radius:14px; text-align:center; background:rgba(6,26,46,.94); border:1px solid var(--primary); box-shadow:0 12px 30px rgba(0,0,0,.4); opacity:0; pointer-events:none; transition:opacity .5s ease, transform .5s ease;}
        .name-badge .nb-name{display:block; color:#eaf7ff; font-weight:1000; font-size:clamp(1rem,4vw,1.35rem); letter-spacing:.01em;}
        .name-badge small{display:block; margin-top:2px; color:#8fbfe0; font-size:clamp(.6rem,2.2vw,.78rem); font-weight:700;}
        .sim.crew-step .name-badge{opacity:1; transform:translate(-50%,0);}


        /* BELGE — klasörden çıkar, açılır (54), önümüze gelir; QR(8), mühür(26/45) */
        .doc{position:fixed; z-index:125; left:50%; bottom:50%; width:min(420px,90vw); max-height:88vh; overflow:auto; padding:22px 22px 18px; border-radius:12px; background:linear-gradient(180deg,#fffdf6,#f4ecd7); color:#173044; border:1px solid #d8c692; box-shadow:0 26px 64px rgba(0,0,0,.5); transform:translate(-50%,54%) scale(.96); transform-origin:center; opacity:0; pointer-events:none; transition:transform .85s cubic-bezier(.2,.83,.2,1), opacity .7s ease; will-change:transform,opacity;}
        .doc::before{content:""; position:absolute; inset:8px; border:1.5px solid rgba(31,87,112,.14); border-radius:8px; pointer-events:none;}
        .sim.reading .doc{opacity:1; transform:translate(-50%,50%) scale(1);}
        .read-dim{position:fixed; inset:0; z-index:110; background:rgba(4,12,20,.74); opacity:0; pointer-events:none; transition:opacity 1.1s ease;}
        .sim.reading .read-dim{opacity:1;}
        .doc-inner{position:relative; z-index:2; display:flex; flex-direction:column; gap:9px;}
        .doc-band{display:flex; align-items:center; gap:10px; padding-bottom:11px; border-bottom:2px solid rgba(31,87,112,.16); color:#245874; font-weight:900; letter-spacing:.03em; font-size:.8rem;}
        .doc-logo-mark{width:34px; height:34px; flex:0 0 34px; display:grid; place-items:center; border-radius:8px; color:#fff; background:linear-gradient(145deg,#168bc7,#0b456b); font-weight:1000; font-size:1.1rem;}
        .doc-ref{color:#6b7f8c; font-size:.72rem; font-weight:800; letter-spacing:.02em;}
        .doc-title{font-family:Georgia,"Times New Roman",serif; font-size:1.45rem; line-height:1.15; color:#19394e; font-weight:800;}
        .doc-caption{color:#4f6c7c; font-weight:700; font-size:.88rem; margin-top:-3px;}
        .doc-balance{font-size:2.9rem; line-height:1; color:#087eb7; font-weight:1000; letter-spacing:-.03em; margin:1px 0 4px;}
        .doc-balance small{font-size:.32em; color:#4f6c7c; font-weight:800; letter-spacing:0;}
        .doc-mid{display:flex; align-items:flex-end; justify-content:space-between; gap:14px; padding-top:11px; border-top:1px solid rgba(31,87,112,.14);}
        .doc-sign{color:#315b70; min-width:0;}
        .doc-sign .sg-role{display:block; font-size:.72rem; font-weight:700; color:#4f6c7c; line-height:1.3;}
        .doc-sign .sg-name{display:block; margin-top:3px; font-weight:900; font-size:1.02rem; color:#1b3a4c;}
        .sg-line{height:2px; margin:6px 0; width:70%; background:#315b70;}
        .doc-stamp{display:inline-block; margin-top:4px; padding:5px 12px; border-radius:6px; border:2.5px solid #1f7a5a; color:#1f7a5a; font-weight:1000; font-size:.66rem; letter-spacing:.12em; transform:rotate(-4deg);}
        .sim.reading .doc-stamp{animation:stampIn .4s cubic-bezier(.2,1.3,.4,1) .35s both;}
        @keyframes stampIn{0%{opacity:0; transform:rotate(-4deg) scale(1.5);}100%{opacity:1; transform:rotate(-4deg) scale(1);}}
        .doc-qr{width:70px; height:70px; flex:0 0 70px; background:#fff; border:1px solid #cbb98f; border-radius:6px; padding:4px;}
        .doc-qr canvas{width:100%; height:100%; display:block;}
        .doc-foot{text-align:center; color:#7c8f9a; font-size:.6rem; font-weight:700; margin-top:3px;}

        /* Kamera hareketleri — sadece yukarı-aşağı, asla sağa-sola */
        .sim.walking .sim-world{animation:walkBob .72s ease-in-out infinite;}
        @keyframes walkBob{0%,100%{transform:translateY(0);}50%{transform:translateY(.9%);}}
        .sim.riding .sim-world{animation:cabinShake .5s ease-in-out infinite;}
        @keyframes cabinShake{0%,100%{transform:translateY(0);}25%{transform:translateY(-.22%);}75%{transform:translateY(.18%);}}
        /* 42) kalkış/duruşta bir kez yaylanma */
        .sim.bump .sim-world{animation:bump .6s ease;}
        @keyframes bump{0%{transform:translateY(0);}30%{transform:translateY(1.4%);}60%{transform:translateY(-.6%);}100%{transform:translateY(0);}}
        .sim.reading .sim-world{animation:none;}

        /* Birinci şahıs el */
        .fp-hand{position:absolute; z-index:40; left:0; right:0; bottom:-4%; height:44%; pointer-events:none;}
        .fp-arm{position:absolute; bottom:-30%; right:6%; width:clamp(120px,24vw,240px); height:clamp(260px,44vw,460px); transform-origin:bottom; transform:translateY(120%); opacity:0; transition:transform 1.1s cubic-bezier(.2,.8,.2,1), opacity .5s ease; will-change:transform;}
        .sim.picking .fp-arm{transform:translateY(28%) rotate(-6deg); opacity:1;}
        .sim.reading .fp-arm{opacity:0;}
        .fp-sleeve{position:absolute; left:20%; right:20%; bottom:0; height:66%; border-radius:38% 38% 16% 16%; background:linear-gradient(90deg,#0a1a2a,#1d3e5d 48%,#0b2034);}
        .fp-palm{position:absolute; left:24%; top:2%; width:52%; height:36%; border-radius:45% 45% 40% 40%; background:linear-gradient(90deg,#ad704b,#edba8d 50%,#c8865c);}
        .fp-arm i{position:absolute; top:-3%; width:13%; height:30%; border-radius:45% 45% 35% 35%; background:linear-gradient(90deg,#b67852,#edba8d 52%,#c88961);}
        .fp-arm i:nth-child(3){left:24%; transform:rotate(-8deg); height:26%;}
        .fp-arm i:nth-child(4){left:36%; height:31%;}
        .fp-arm i:nth-child(5){left:49%; height:29%;}
        .fp-arm i:nth-child(6){left:62%; transform:rotate(8deg); height:24%;}

        /* D) 67. kat penceresi (kapı açılmadan görünür) */

        .sim-status{position:absolute; z-index:130; left:50%; bottom:max(16px,env(safe-area-inset-bottom)); transform:translateX(-50%); max-width:calc(100vw - 30px); padding:9px 16px; border-radius:999px; color:#d9f5ff; background:rgba(3,16,26,.92); border:1px solid rgba(141,222,255,.24); font-size:clamp(.66rem,2.2vw,.8rem); font-weight:900; letter-spacing:.08em; text-align:center;}

        /* 36) kapanışta logo parlaması */
        .logo-flash{position:fixed; inset:0; z-index:160; display:none; place-items:center; background:radial-gradient(circle at 50% 50%, #0a2440, #04101d); }
        .logo-flash.show{display:grid; animation:logoFlash 1.1s ease forwards;}
        @keyframes logoFlash{0%{opacity:0;}25%{opacity:1;}75%{opacity:1;}100%{opacity:0;}}
        .logo-flash .lf-mark{width:96px; height:96px; border-radius:26px; display:grid; place-items:center; font-size:2.6rem; font-weight:1000; color:#eaf7ff; background:linear-gradient(145deg,var(--brand),var(--brand-2)); box-shadow:0 0 60px rgba(85,217,255,.6); animation:lfPop .9s cubic-bezier(.2,1.3,.4,1);}
        @keyframes lfPop{0%{transform:scale(.4); opacity:0;}50%{transform:scale(1.08); opacity:1;}100%{transform:scale(1);}}

        /* ---------- DASHBOARD ---------- */
        .dashboard{width:min(1120px,100%); margin:0 auto; animation:dashboardIn .5s cubic-bezier(.2,.8,.2,1) both;}
        @keyframes dashboardIn{from{opacity:0; transform:translateY(18px);}to{opacity:1; transform:none;}}
        /* header düzeltmesi: butonlar bilgilerin önüne binmez */
        .dashboard-header{display:flex; flex-wrap:wrap; align-items:center; gap:12px; margin-bottom:18px;}
        .person{display:flex; align-items:center; gap:12px; min-width:0; flex:1 1 100%;}
        .avatar{width:58px; height:58px; flex:0 0 58px; display:grid; place-items:center; border-radius:18px; background:linear-gradient(145deg,var(--primary),var(--primary-2)); color:#032039; font-size:1.15rem; font-weight:1000; box-shadow:0 12px 26px rgba(28,164,220,.25);}
        .person > div{min-width:0;}
        .person h2{margin:0; font-size:clamp(1.2rem,4vw,1.8rem); white-space:nowrap; overflow:hidden; text-overflow:ellipsis;}
        .person p{margin:4px 0 0; color:var(--muted); white-space:nowrap; overflow:hidden; text-overflow:ellipsis;}
        .header-buttons{display:flex; gap:10px; width:100%;}
        .small-btn{flex:1 1 0; min-width:0; border:1px solid var(--line); background:var(--panel); border-radius:13px; padding:12px 14px; cursor:pointer; font-weight:800; text-align:center; white-space:normal; transition:transform .2s, border-color .2s;}
        .small-btn:hover{transform:translateY(-2px); border-color:var(--primary);}
        .dashboard-grid{display:grid; grid-template-columns:minmax(0,1.35fr) minmax(280px,.65fr); gap:18px;}
        .main-column,.side-column{display:grid; gap:18px; align-content:start; min-width:0;}
        /* 49) kart canlılığı */
        .hero-card,.mini-card,.actions-card,.info-card,.id-card-wrap{transition:transform .25s ease, box-shadow .25s ease;}
        .mini-card:hover,.info-card:hover,.id-card-wrap:hover{transform:translateY(-4px); box-shadow:0 34px 80px rgba(0,0,0,.4);}
        .mini-card:active,.info-card:active{transform:translateY(-1px) scale(.995);}
        .hero-card{padding:clamp(22px,5vw,36px); position:relative; overflow:hidden;}
        .hero-card::after{content:""; position:absolute; width:260px; height:260px; right:-90px; top:-90px; border-radius:50%; background:radial-gradient(circle,rgba(85,217,255,.24),transparent 67%); pointer-events:none;}
        .hero-content{position:relative; z-index:2;}
        .eyebrow{color:var(--primary); text-transform:uppercase; letter-spacing:.13em; font-size:.76rem; font-weight:1000;}
        .big-balance{margin:10px 0 5px; font-size:clamp(3rem,12vw,6.3rem); line-height:.92; letter-spacing:-.075em; font-weight:1000;}
        .big-balance small{font-size:.2em; letter-spacing:.02em; color:var(--muted); margin-left:8px;}
        .update-info{color:var(--muted); font-size:.85rem; margin-top:13px;}
        .progress-ring{--progress:75deg; width:clamp(108px,18vw,150px); aspect-ratio:1; border-radius:50%; display:grid; place-items:center; background:conic-gradient(var(--primary) var(--progress), rgba(127,202,230,.13) 0); position:relative; box-shadow:0 0 36px rgba(62,205,255,.12); transition:--progress 1s ease;}
        .progress-ring::before{content:""; position:absolute; inset:12px; border-radius:50%; background:var(--panel-strong); border:1px solid var(--line);}
        .progress-ring span{position:relative; z-index:2; text-align:center; color:var(--muted); font-size:.75rem; font-weight:800;}
        .progress-ring b{display:block; color:var(--text); font-size:1.15rem; margin-bottom:2px;}
        .mini-cards{display:grid; grid-template-columns:repeat(2,minmax(0,1fr)); gap:14px;}
        .mini-card{padding:20px;}
        .mini-card .value{margin-top:9px; font-size:2rem; font-weight:1000;}
        .mini-card .label{color:var(--muted); font-size:.86rem; font-weight:700;}
        .actions-card{padding:20px;}
        .actions-card h3,.info-card h3,.id-card-wrap h3{margin:0 0 14px; font-size:1rem;}
        .action-list{display:grid; gap:10px;}
        .action-btn{justify-content:flex-start;}
        .action-btn.whatsapp{background:linear-gradient(135deg,#2bd881,#0d9d64); color:#04291b;}
        .action-btn.leave{background:linear-gradient(135deg,#ffd36d,#e99c25); color:#3b2601;}
        .action-btn.secondary{background:rgba(255,255,255,.06); border:1px solid var(--line); color:var(--text);}
        .info-card{padding:20px;}
        .info-row{display:flex; justify-content:space-between; gap:12px; padding:12px 0; border-bottom:1px solid var(--line);}
        .info-row:last-child{border-bottom:0; padding-bottom:0;}
        .info-row span{color:var(--muted);} .info-row strong{text-align:right;}
        .holiday-radar{position:relative; min-height:162px; display:grid; place-items:center; overflow:hidden; text-align:center;}
        .radar{position:absolute; width:190px; height:190px; border-radius:50%; border:1px solid rgba(84,218,255,.18); background:linear-gradient(90deg,transparent 49.5%,rgba(84,218,255,.12) 50%,transparent 50.5%),linear-gradient(transparent 49.5%,rgba(84,218,255,.12) 50%,transparent 50.5%),radial-gradient(circle,transparent 0 24%,rgba(84,218,255,.08) 25% 25.8%,transparent 26% 49%,rgba(84,218,255,.08) 50% 50.8%,transparent 51%);}
        .radar::after{content:""; position:absolute; inset:50% 50% auto auto; width:50%; height:2px; transform-origin:left; background:linear-gradient(90deg,var(--primary),transparent); animation:radarSpin 4s linear infinite;}
        @keyframes radarSpin{to{transform:rotate(360deg);}}
        .holiday-content{position:relative; z-index:2;}
        .holiday-days{font-size:2.5rem; font-weight:1000; color:var(--primary);}
        .holiday-name{font-weight:900;}
        .holiday-date{color:var(--muted); font-size:.82rem; margin-top:4px;}
        .id-card-wrap{padding:20px;}
        .digital-id{min-height:218px; padding:20px; border-radius:20px; color:white; background:radial-gradient(circle at 85% 18%,rgba(83,229,255,.32),transparent 30%),linear-gradient(145deg,#08203a,#0a5279); border:1px solid rgba(130,224,255,.33); position:relative; overflow:hidden; box-shadow:0 18px 38px rgba(0,0,0,.22);}
        .digital-id::after{content:""; position:absolute; width:160px; height:160px; border:24px solid rgba(255,255,255,.035); border-radius:50%; right:-55px; bottom:-78px;}
        .id-top{display:flex; justify-content:space-between; align-items:flex-start; gap:14px;}
        .id-company{font-size:.72rem; letter-spacing:.12em; font-weight:1000; color:#a9eaff;}
        .id-chip{width:42px; height:32px; border-radius:7px; background:linear-gradient(90deg,transparent 47%,rgba(77,60,5,.3) 48% 52%,transparent 53%),linear-gradient(#f3d77b,#bb9027); border:1px solid rgba(255,239,169,.8);}
        .id-main{display:grid; grid-template-columns:1fr 92px; align-items:end; gap:14px; margin-top:28px; position:relative; z-index:2;}
        .id-name{font-size:1.25rem; font-weight:1000; word-break:break-word;}
        .id-role{color:#b9dded; font-size:.78rem; margin:5px 0 20px;}
        .id-number{color:#86dfff; font:800 .78rem ui-monospace,monospace;}
        .qr-box{width:92px; height:92px; padding:7px; border-radius:10px; background:white; display:grid; place-items:center;}
        #qrCanvas{width:78px; height:78px;}
        .modal-backdrop{position:fixed; z-index:220; inset:0; display:grid; place-items:center; padding:20px; background:rgba(1,10,18,.82); opacity:0; pointer-events:none; transition:opacity .25s ease;}
        .modal-backdrop.open{opacity:1; pointer-events:auto;}
        .modal{width:min(520px,100%); max-height:min(760px,90dvh); overflow:auto; padding:24px; transform:translateY(14px) scale(.98); transition:transform .25s ease;}
        .modal-backdrop.open .modal{transform:none;}
        .modal-head{display:flex; justify-content:space-between; align-items:center; gap:12px; margin-bottom:18px;}
        .modal h3{margin:0;}
        .close-btn{width:38px; height:38px; display:grid; place-items:center; border-radius:12px; border:1px solid var(--line); background:rgba(255,255,255,.05); cursor:pointer;}
        .modal textarea{min-height:100px; resize:vertical; padding-right:15px;}
        .modal select{padding-right:15px;}
        .form-grid{display:grid; grid-template-columns:1fr 1fr; gap:12px;}
        .modal .field{margin-bottom:13px;}
        #toast{position:fixed; z-index:400; top:max(72px,calc(env(safe-area-inset-top) + 58px)); right:16px; width:min(360px,calc(100vw - 32px)); padding:14px 16px; border-radius:15px; background:var(--panel-strong); border:1px solid var(--line); box-shadow:0 18px 45px rgba(0,0,0,.3); transform:translateX(calc(100% + 30px)); opacity:0; transition:transform .4s cubic-bezier(.2,.8,.2,1), opacity .3s ease; font-weight:800;}
        #toast.show{transform:none; opacity:1;}
        #toast.error{border-color:rgba(255,102,120,.55);}
        #toast.success{border-color:rgba(43,216,129,.55);}

        /* ---------- RESPONSIVE (5 kırılım) ---------- */
        @media (max-width:1200px){
            .dept-panel{right:8%; width:min(38vw,250px);}
        }
        @media (max-width:900px){
            .dashboard-grid{grid-template-columns:1fr;}
            .dept-panel{right:6%; top:17%; width:min(46vw,240px);}
            .cabin-doorway{left:20%; right:20%;}
            .cabin-plate{width:min(24vw,110px);}
        }
        @media (max-width:640px){
            .brand-mini span:last-child{display:none;}
            .install-btn span{display:none;}
            .page-shell{padding-left:12px; padding-right:12px;}
            .form-grid{grid-template-columns:1fr;}
            .id-main{grid-template-columns:1fr 84px;}
            .qr-box{width:84px; height:84px;} #qrCanvas{width:70px; height:70px;}
            .dept-panel{right:4%; top:15%; width:min(52vw,220px); padding:9px;}
            .dept-btn{min-height:42px; margin:5px 0;}
            .cabin-doorway{left:15%; right:15%;}
            .cabin-plate{left:2%; top:16%; font-size:.5rem;}
            .office-sign .os-main{font-size:.82rem;}
        }
        @media (max-width:400px){
            .login-card{padding:22px 18px;}
            .big-balance{font-size:clamp(2.6rem,20vw,3.4rem);}
            .dept-panel{width:min(58vw,200px);}
            .floor-ind{min-width:120px; padding:6px 10px;}
            .doc{width:min(92vw,360px); padding:18px;}
        }
        @media (prefers-reduced-motion:reduce){
            *,*::before,*::after{animation-duration:.01ms !important; animation-iteration-count:1 !important; transition-duration:.24s !important;}
        }
    </style>
</head>
<body>
    <div class="topbar">
        <div class="brand-mini"><span class="brand-mark">İ</span><span data-i18n="brand">İzin Portalı</span></div>
        <div class="top-actions">
            <button class="install-btn" id="installBtn" type="button" aria-label="Uygulamayı yükle"><span>＋</span><span data-i18n="install">Uygulamayı Yükle</span></button>
            <button class="icon-btn" id="languageBtn" type="button" aria-label="Dil değiştir">TR</button>
            <button class="icon-btn" id="themeBtn" type="button" aria-label="Tema değiştir">☾</button>
        </div>
    </div>

    <main class="page-shell">
        <section class="login-card glass" id="loginPanel">
            <div class="success-wave"></div>
            <div class="login-logo"><div class="elevator-icon"></div></div>
            <h1 data-i18n="loginTitle">Personel İzin Sistemi</h1>
            <p class="login-subtitle" data-i18n="loginSubtitle">Kişisel izin bilgilerinize güvenli şekilde ulaşın.</p>
            <form id="loginForm" novalidate>
                <div class="field">
                    <label for="username" data-i18n="username">Kullanıcı Adı</label>
                    <div class="input-wrap">
                        <input id="username" name="username" autocomplete="username" inputmode="text" data-i18n-placeholder="usernamePlaceholder" placeholder="Kullanıcı adınızı girin">
                        <span class="field-icon">●</span>
                    </div>
                </div>
                <div class="field">
                    <label for="password" data-i18n="password">Şifre</label>
                    <div class="input-wrap">
                        <input id="password" name="password" type="password" autocomplete="current-password" data-i18n-placeholder="passwordPlaceholder" placeholder="Şifrenizi girin">
                        <button class="field-icon password-toggle" id="passwordToggle" type="button" aria-label="Şifreyi göster">◉</button>
                    </div>
                    <div class="caps-hint" id="capsHint" data-i18n="capsLock">⚠ Caps Lock açık</div>
                </div>
                <label class="remember-row"><input type="checkbox" id="rememberMe"><span data-i18n="rememberMe">Bu cihazda kullanıcı adımı hatırla</span></label>
                <button class="primary-btn" id="loginBtn" type="submit">
                    <span id="loginBtnText" data-i18n="login">Giriş Yap</span>
                    <span class="spinner"></span>
                </button>
            </form>
            <div class="login-links"><button class="text-btn" id="forgotBtn" type="button" data-i18n="forgot">Şifremi Unuttum</button></div>
            <div class="sys-status"><span class="sys-dot"></span><span data-i18n="systemActive">Sistem aktif</span></div>
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
                    <button class="small-btn" id="replayBtn" type="button" data-i18n="replay">Tekrar İzle</button>
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
                        </div>
                    </article>
                    <div class="mini-cards">
                        <article class="mini-card glass"><div class="label" data-i18n="sundayLeave">Pazar İzinleri</div><div class="value"><span id="sundayLeave">0</span> <small data-i18n="dayLower">gün</small></div></article>
                        <article class="mini-card glass"><div class="label" data-i18n="officialHoliday">Resmî Tatil</div><div class="value"><span id="officialHoliday">0</span> <small data-i18n="dayLower">gün</small></div></article>
                    </div>
                    <article class="actions-card glass">
                        <h3 data-i18n="quickActions">Hızlı İşlemler</h3>
                        <div class="action-list">
                            <a class="action-btn whatsapp" id="objectionBtn" href="#" target="_blank" rel="noopener"><span>◉</span><span data-i18n="objectLeave">İzin Gününe İtiraz Et</span></a>
                            <button class="action-btn leave" id="leaveRequestBtn" type="button"><span>▣</span><span data-i18n="requestLeave">İzin Talebi Oluştur</span></button>
                            <button class="action-btn secondary" id="installActionBtn" type="button"><span>＋</span><span data-i18n="addHome">Ana Ekrana Uygulama Olarak Ekle</span></button>
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
                        <div class="info-row"><span data-i18n="lastLogin">Son girişiniz</span><strong id="lastLogin">İlk giriş</strong></div>
                        <div class="info-row"><span data-i18n="session">Oturum</span><strong data-i18n="active">Aktif</strong></div>
                    </article>
                    <article class="id-card-wrap glass">
                        <h3 data-i18n="digitalId">Dijital Personel Kimliği</h3>
                        <div class="digital-id">
                            <div class="id-top"><div class="id-company">PERSONEL PORTALI</div><div class="id-chip"></div></div>
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

    <!-- A) giriş kapıları -->
    <div class="entry-doors" id="entryDoors"><div class="ed l"></div><div class="ed r"></div></div>

    <!-- 36) kapanış logosu -->
    <div class="logo-flash" id="logoFlash"><div class="lf-mark">İ</div></div>

    <!-- ============ KESİNTİSİZ ASANSÖR SİMÜLASYONU ============ -->
    <section class="sim" id="sim" hidden aria-label="Asansör animasyonu">
        <div class="read-dim"></div>
        <div class="sim-world" id="simWorld">

            <!-- OFİS (sade: aydınlık koridor + kat tabelası) -->
            <div class="stage office">
                <div class="office-ceil"></div>
                <div class="office-floor"></div>
                <div class="office-wall left"></div>
                <div class="office-wall right"></div>
                <div class="office-sign"><span class="os-main">PERSONEL VE ÇALIŞMA İLİŞKİLERİ</span><span class="os-floor" data-i18n="floor67">67. KAT</span></div>

                <!-- 5 kişilik departman ekibi; Semih Bayat biraz uzun boylu, öne çıkar -->
                <div class="crew-row">
                    <div class="crew" style="--suit:#2f5a7d; --hair:#241a14"><span class="c-shadow"></span><span class="c-hair"></span><span class="c-head"></span><span class="c-body"></span></div>
                    <div class="crew" style="--suit:#3a4a58; --hair:#1c1410"><span class="c-shadow"></span><span class="c-hair"></span><span class="c-head"></span><span class="c-body"></span></div>
                    <div class="crew semih" style="--suit:#1f6f8b; --hair:#2b1d15"><span class="c-shadow"></span><span class="c-hair"></span><span class="c-head"></span><span class="c-body"></span></div>
                    <div class="crew" style="--suit:#4a3f5c; --hair:#20160f"><span class="c-shadow"></span><span class="c-hair"></span><span class="c-head"></span><span class="c-body"></span></div>
                    <div class="crew" style="--suit:#2e5148; --hair:#241a14"><span class="c-shadow"></span><span class="c-hair"></span><span class="c-head"></span><span class="c-body"></span></div>
                </div>
                <div class="name-badge"><span class="nb-name">Semih Bayat</span><small>Personel ve Çalışma İlişkileri</small></div>
            </div>

            <!-- KABİN (sade) -->
            <div class="stage cabin">
                <div class="cabin-ceil"></div>
                <div class="spot s1"></div><div class="spot s2"></div><div class="spot s3"></div>
                <div class="cabin-floor"></div>
                <div class="wall left"></div>
                <div class="wall right"></div>

                <div class="cabin-plate" data-i18n="cabinPlate">MAX 8 KİŞİ<br>630 KG</div>

                <div class="floor-ind"><div class="fi-top"><span class="fi-arrow">▲</span><b id="floorValue">0</b></div><small data-i18n="leaveLevelUpper">KAT</small></div>

                <div class="doorway cabin-doorway">
                    <div class="door l"><div class="door-glass"></div></div>
                    <div class="door r"><div class="door-glass"></div></div>
                </div>
                <div class="door-leak"></div>

                <div class="dept-panel">
                    <div class="dept-title" data-i18n="deptTitle">DEPARTMAN</div>
                    <button class="dept-btn" type="button"><span class="halo"></span><span class="dept-num">1</span><span>İdari İşler</span></button>
                    <button class="dept-btn" type="button"><span class="halo"></span><span class="dept-num">2</span><span>İSG</span></button>
                    <button class="dept-btn" type="button"><span class="halo"></span><span class="dept-num">3</span><span>Yönetim</span></button>
                    <button class="dept-btn" type="button"><span class="halo"></span><span class="dept-num">4</span><span>Kalite</span></button>
                    <button class="dept-btn target" type="button"><span class="halo"></span><span class="dept-num">5</span><span>Personel ve Çalışma İlişkileri</span></button>
                </div>
            </div>

            <!-- LOBİ -->
            <div class="stage lobby">
                <div class="lobby-wall"></div>
                <div class="lobby-frame">
                    <div class="lobby-head" id="lobbyDisplay">G</div>
                    <div class="doorway lobby-doorway">
                        <div class="door l"></div>
                        <div class="door r"></div>
                    </div>
                </div>
            </div>
        </div>

        <!-- Belge: direkt onaylı, üst katmanda net -->
        <div class="doc" id="doc">
                            <div class="doc-inner">
                                <div class="doc-band"><span class="doc-logo-mark">İ</span><span>PERSONEL İZİN BİLDİRİMİ</span></div>
                                <div class="doc-ref" id="docRef">Ref: — · —</div>
                                <div class="doc-title" id="docTitle">Sn. Burhan Biliktü</div>
                                <div class="doc-caption" data-i18n="noteCaption">Kalan İzin Hakkınız</div>
                                <div class="doc-balance"><span id="docBalance">0</span> <small data-i18n="dayLower">gün</small></div>
                                <div class="doc-mid">
                                    <div class="doc-sign">
                                        <span class="sg-role" data-i18n="hrChief">Personel ve Çalışma İlişkileri Şefi</span>
                                        <span class="sg-name">İlker Sezgin</span>
                                        <span class="sg-line"></span>
                                        <span class="doc-stamp" data-i18n="digitalApproval">DİJİTAL ONAY</span>
                                    </div>
                                    <div class="doc-qr"><canvas id="docQr" width="120" height="120"></canvas></div>
                                </div>
                                <div class="doc-foot" data-i18n="docFoot">Bu belge dijital sistem üzerinden oluşturulmuştur.</div>
                            </div>
                        </div>

        <div class="fp-hand" id="fpHand">
            <div class="fp-arm"><div class="fp-sleeve"></div><div class="fp-palm"></div><i></i><i></i><i></i><i></i></div>
        </div>

        <div class="sim-status" id="elevatorStatus">ASANSÖR</div>
    </section>

    <div class="modal-backdrop" id="forgotModal" aria-hidden="true">
        <div class="modal glass">
            <div class="modal-head"><h3 data-i18n="forgotTitle">Şifre Talebi</h3><button class="close-btn" type="button" data-close="forgotModal">×</button></div>
            <div class="field"><label for="forgotIdentity" data-i18n="nameOrUsername">Ad Soyad veya Kullanıcı Adı</label><input id="forgotIdentity" data-i18n-placeholder="identityPlaceholder" placeholder="Bilginizi yazın"></div>
            <button class="primary-btn" id="forgotWhatsappBtn" type="button" data-i18n="sendWhatsapp">WhatsApp'tan Gönder</button>
        </div>
    </div>

    <div class="modal-backdrop" id="leaveModal" aria-hidden="true">
        <div class="modal glass">
            <div class="modal-head"><h3 data-i18n="leaveRequestTitle">İzin Talebi Oluştur</h3><button class="close-btn" type="button" data-close="leaveModal">×</button></div>
            <div class="field">
                <label for="leaveType" data-i18n="leaveType">İzin Türü</label>
                <select id="leaveType">
                    <option value="Yıllık İzin" data-i18n="annualLeave">Yıllık İzin</option>
                    <option value="Mazeret İzni" data-i18n="excuseLeave">Mazeret İzni</option>
                    <option value="Ücretsiz İzin" data-i18n="unpaidLeave">Ücretsiz İzin</option>
                </select>
            </div>
            <div class="form-grid">
                <div class="field"><label for="leaveStart" data-i18n="startDate">Başlangıç Tarihi</label><input id="leaveStart" type="date"></div>
                <div class="field"><label for="leaveEnd" data-i18n="endDate">Bitiş Tarihi</label><input id="leaveEnd" type="date"></div>
            </div>
            <div class="field"><label for="leaveDescription" data-i18n="description">Açıklama</label><textarea id="leaveDescription" data-i18n-placeholder="descriptionPlaceholder" placeholder="Talebinizle ilgili kısa açıklama"></textarea></div>
            <button class="primary-btn" id="sendLeaveBtn" type="button" data-i18n="sendWhatsapp">WhatsApp'tan Gönder</button>
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
                brand:"İzin Portalı", install:"Uygulamayı Yükle",
                loginTitle:"Personel İzin Sistemi", loginSubtitle:"Kişisel izin bilgilerinize güvenli şekilde ulaşın.",
                username:"Kullanıcı Adı", password:"Şifre",
                usernamePlaceholder:"Kullanıcı adınızı girin", passwordPlaceholder:"Şifrenizi girin",
                login:"Giriş Yap", forgot:"Şifremi Unuttum",
                rememberMe:"Bu cihazda kullanıcı adımı hatırla", systemActive:"Sistem aktif", capsLock:"⚠ Caps Lock açık",
                replay:"Tekrar İzle", logout:"Çıkış",
                remainingLeave:"Kalan İzin Hakkınız", day:"GÜN", dayLower:"gün",
                updated:"Son güncelleme:", leaveLevel:"Kat",
                sundayLeave:"Pazar İzinleri", officialHoliday:"Resmî Tatil",
                quickActions:"Hızlı İşlemler", objectLeave:"İzin Gününe İtiraz Et",
                requestLeave:"İzin Talebi Oluştur", addHome:"Ana Ekrana Uygulama Olarak Ekle",
                nextHoliday:"Yaklaşan Resmî Tatil", security:"Güvenlik Bilgisi",
                lastLogin:"Son girişiniz", session:"Oturum", active:"Aktif",
                digitalId:"Dijital Personel Kimliği", leaveLevelUpper:"KAT",
                floor67:"67. KAT", cabinPlate:"MAX 8 KİŞİ<br>630 KG",
                noteCaption:"Kalan İzin Hakkınız",
                hrChief:"Personel ve Çalışma İlişkileri Şefi", digitalApproval:"DİJİTAL ONAY",
                docFoot:"Bu belge dijital sistem üzerinden oluşturulmuştur.", deptTitle:"DEPARTMAN",
                forgotTitle:"Şifre Talebi", nameOrUsername:"Ad Soyad veya Kullanıcı Adı",
                identityPlaceholder:"Bilginizi yazın", sendWhatsapp:"WhatsApp'tan Gönder",
                leaveRequestTitle:"İzin Talebi Oluştur", leaveType:"İzin Türü",
                annualLeave:"Yıllık İzin", excuseLeave:"Mazeret İzni", unpaidLeave:"Ücretsiz İzin",
                startDate:"Başlangıç Tarihi", endDate:"Bitiş Tarihi", description:"Açıklama",
                descriptionPlaceholder:"Talebinizle ilgili kısa açıklama",
                firstLogin:"İlk giriş", greetingMorning:"Günaydın", greetingDay:"Hoş geldiniz", greetingEvening:"İyi akşamlar",
                stCalling:"ASANSÖR ÇAĞRILIYOR", stDoorsOpen:"KAPILAR AÇILIYOR", stEntering:"ASANSÖRE BİNİLİYOR",
                stPicking:"PERSONEL VE ÇALIŞMA İLİŞKİLERİ SEÇİLİYOR", stRising:"YUKARI ÇIKIYOR",
                stArrived:"67. KATA ULAŞILDI", stOffice:"KATA ÇIKILIYOR", stCrew:"PERSONEL VE ÇALIŞMA İLİŞKİLERİ EKİBİ", stSemih:"SEMİH BAYAT KARŞILIYOR", stSecretary:"SEKRETER YAKLAŞIYOR", stDelivery:"BELGE TESLİM EDİLİYOR",
                fillFields:"Lütfen kullanıcı adı ve şifreyi girin.", loginError:"Kullanıcı adı veya şifre hatalı.",
                serverError:"Sunucuya bağlanılamadı.", blocked:"Çok fazla hatalı deneme. Lütfen daha sonra tekrar deneyin.",
                installReady:"Uygulama ana ekrana eklenmeye hazır.",
                iosInstall:"iPhone'da Paylaş simgesine dokunup \"Ana Ekrana Ekle\" seçeneğini kullanın.",
                installUnavailable:"Tarayıcı menüsünden \"Ana ekrana ekle\" seçeneğini kullanabilirsiniz.",
                fillIdentity:"Lütfen adınızı veya kullanıcı adınızı yazın.",
                fillDates:"Lütfen başlangıç ve bitiş tarihlerini seçin.", invalidDates:"Bitiş tarihi başlangıç tarihinden önce olamaz.",
                sentToWhatsapp:"WhatsApp açılıyor.",
                objectionMessage:name=>`Merhaba, adım ${name}. İzin sisteminde görünen gün sayısının hatalı olduğunu düşünüyorum. Kalan izin hakkım: ${formatNumber(currentUser.remaining_leave)} gün. Son güncelleme: ${currentUser.updated_at}. İtiraz etmek istiyorum.`,
                forgotMessage:identity=>`Merhaba, Personel İzin Portalı şifremi unuttum. Ad Soyad / Kullanıcı Adı: ${identity}. Şifre konusunda destek rica ederim.`,
                leaveMessage:data=>`Merhaba, izin talebimi iletmek istiyorum.\n\nAd Soyad: ${currentUser.name}\nKullanıcı Adı: ${currentUser.username}\nİzin Türü: ${data.type}\nBaşlangıç: ${data.start}\nBitiş: ${data.end}\nToplam Takvim Günü: ${data.days}\nMevcut Kalan İzin: ${formatNumber(currentUser.remaining_leave)} gün\nAçıklama: ${data.description || "-"}\n\nOnaya sunarım.`,
                noteTitle:name=>`Sn. ${name}`,
                holidayIn:days=>`${days} gün`,
            },
            en: {
                brand:"Leave Portal", install:"Install App",
                loginTitle:"Employee Leave System", loginSubtitle:"Securely access your personal leave information.",
                username:"Username", password:"Password",
                usernamePlaceholder:"Enter your username", passwordPlaceholder:"Enter your password",
                login:"Sign In", forgot:"Forgot Password",
                rememberMe:"Remember my username on this device", systemActive:"System active", capsLock:"⚠ Caps Lock is on",
                replay:"Replay", logout:"Sign Out",
                remainingLeave:"Remaining Leave Balance", day:"DAYS", dayLower:"days",
                updated:"Last update:", leaveLevel:"Floor",
                sundayLeave:"Sunday Leave", officialHoliday:"Public Holiday",
                quickActions:"Quick Actions", objectLeave:"Object to Leave Balance",
                requestLeave:"Create Leave Request", addHome:"Add App to Home Screen",
                nextHoliday:"Next Public Holiday", security:"Security Information",
                lastLogin:"Your last sign-in", session:"Session", active:"Active",
                digitalId:"Digital Employee ID", leaveLevelUpper:"FLOOR",
                floor67:"FLOOR 67", cabinPlate:"MAX 8 PERSONS<br>630 KG",
                noteCaption:"Your Remaining Leave Balance",
                hrChief:"Chief of Personnel and Labour Relations", digitalApproval:"DIGITAL APPROVAL",
                docFoot:"This document was generated by the digital system.", deptTitle:"DEPARTMENT",
                forgotTitle:"Password Request", nameOrUsername:"Full Name or Username",
                identityPlaceholder:"Enter your information", sendWhatsapp:"Send via WhatsApp",
                leaveRequestTitle:"Create Leave Request", leaveType:"Leave Type",
                annualLeave:"Annual Leave", excuseLeave:"Excuse Leave", unpaidLeave:"Unpaid Leave",
                startDate:"Start Date", endDate:"End Date", description:"Description",
                descriptionPlaceholder:"Briefly explain your request",
                firstLogin:"First sign-in", greetingMorning:"Good morning", greetingDay:"Welcome", greetingEvening:"Good evening",
                stCalling:"CALLING ELEVATOR", stDoorsOpen:"DOORS OPENING", stEntering:"ENTERING ELEVATOR",
                stPicking:"SELECTING PERSONNEL AND LABOUR RELATIONS", stRising:"GOING UP",
                stArrived:"REACHED FLOOR 67", stOffice:"ENTERING THE FLOOR", stCrew:"PERSONNEL & LABOUR RELATIONS TEAM", stSemih:"SEMİH BAYAT WELCOMES YOU", stSecretary:"SECRETARY IS APPROACHING", stDelivery:"DELIVERING DOCUMENT",
                fillFields:"Please enter your username and password.", loginError:"Incorrect username or password.",
                serverError:"Could not connect to the server.", blocked:"Too many failed attempts. Please try again later.",
                installReady:"The app is ready to be installed.",
                iosInstall:"On iPhone, tap Share and choose \"Add to Home Screen\".",
                installUnavailable:"Use your browser menu and choose \"Add to Home screen\".",
                fillIdentity:"Please enter your name or username.",
                fillDates:"Please select start and end dates.", invalidDates:"End date cannot be before start date.",
                sentToWhatsapp:"Opening WhatsApp.",
                objectionMessage:name=>`Hello, my name is ${name}. I believe the leave balance shown in the system is incorrect. Remaining leave: ${formatNumber(currentUser.remaining_leave)} days. Last update: ${currentUser.updated_at}. I would like to submit an objection.`,
                forgotMessage:identity=>`Hello, I forgot my Employee Leave Portal password. Full Name / Username: ${identity}. I kindly request support.`,
                leaveMessage:data=>`Hello, I would like to submit a leave request.\n\nName: ${currentUser.name}\nUsername: ${currentUser.username}\nLeave Type: ${data.type}\nStart: ${data.start}\nEnd: ${data.end}\nCalendar Days: ${data.days}\nCurrent Leave Balance: ${formatNumber(currentUser.remaining_leave)} days\nDescription: ${data.description || "-"}\n\nSubmitted for approval.`,
                noteTitle:name=>`Dear ${name}`,
                holidayIn:days=>`${days} days`,
            }
        };

        function t(key){ return translations[currentLanguage][key] ?? translations.tr[key] ?? key; }

        function applyLanguage(){
            document.documentElement.lang = currentLanguage;
            document.getElementById("languageBtn").textContent = currentLanguage.toUpperCase();
            document.querySelectorAll("[data-i18n]").forEach(el=>{
                const v = t(el.dataset.i18n);
                if(typeof v !== "string") return;
                if(el.dataset.i18n === "cabinPlate") el.innerHTML = v; else el.textContent = v;
            });
            document.querySelectorAll("[data-i18n-placeholder]").forEach(el=>{ el.placeholder = t(el.dataset.i18nPlaceholder); });
            if(currentUser){
                renderUser(currentUser, false);
                document.getElementById("docTitle").textContent = t("noteTitle")(currentUser.name);
            }
            calculateNextHoliday();
        }

        function autoTheme(){
            const saved = localStorage.getItem("izin-theme");
            const hour = new Date().getHours();
            const light = saved ? saved === "light" : !(hour >= 19 || hour < 7);
            document.body.classList.toggle("light", light);
            document.documentElement.classList.toggle("light", light);
            document.getElementById("themeBtn").textContent = light ? "☀" : "☾";
            document.querySelector('meta[name="theme-color"]').setAttribute("content", light ? "#e7f7ff" : "#071b30");
        }
        function toggleTheme(){
            const light = !document.body.classList.contains("light");
            localStorage.setItem("izin-theme", light ? "light" : "dark");
            autoTheme();
        }

        function formatNumber(value){
            return new Intl.NumberFormat(currentLanguage === "tr" ? "tr-TR" : "en-US", {maximumFractionDigits:2}).format(Number(value || 0));
        }
        function formatDateTime(isoValue){
            if(!isoValue) return t("firstLogin");
            const date = new Date(isoValue);
            if(Number.isNaN(date.getTime())) return t("firstLogin");
            return new Intl.DateTimeFormat(currentLanguage === "tr" ? "tr-TR" : "en-GB", {dateStyle:"medium", timeStyle:"short"}).format(date);
        }

        function showToast(message, type=""){
            const toast = document.getElementById("toast");
            toast.textContent = message; toast.className = type;
            requestAnimationFrame(()=>toast.classList.add("show"));
            clearTimeout(showToast.timer);
            showToast.timer = setTimeout(()=>toast.classList.remove("show"), 3500);
        }
        function openModal(id){ const m=document.getElementById(id); m.classList.add("open"); m.setAttribute("aria-hidden","false"); }
        function closeModal(id){ const m=document.getElementById(id); m.classList.remove("open"); m.setAttribute("aria-hidden","true"); }
        function openWhatsApp(message){
            window.open(`https://wa.me/${WHATSAPP_NUMBER}?text=${encodeURIComponent(message)}`, "_blank", "noopener");
            showToast(t("sentToWhatsapp"), "success");
        }
        function buzz(pattern){ try{ if(navigator.vibrate) navigator.vibrate(pattern); }catch(e){} }

        /* ---- Ses (21): motor uğultusu + ding. Giriş jestiyle açılır, kısık, sessizde çalmaz ---- */
        const Audio = (function(){
            let ctx=null, hum=null, humGain=null, master=null;
            function ensure(){
                if(ctx) return true;
                try{
                    const AC = window.AudioContext || window.webkitAudioContext;
                    if(!AC) return false;
                    ctx = new AC();
                    master = ctx.createGain(); master.gain.value = 0.5; master.connect(ctx.destination);
                }catch(e){ return false; }
                return !!ctx;
            }
            function resume(){ if(ctx && ctx.state === "suspended"){ ctx.resume().catch(()=>{}); } }
            return {
                init(){ ensure(); resume(); },
                startHum(){
                    if(!ensure()) return; resume();
                    try{
                        this.stopHum();
                        hum = ctx.createOscillator(); hum.type = "sawtooth"; hum.frequency.value = 58;
                        const lp = ctx.createBiquadFilter(); lp.type="lowpass"; lp.frequency.value=220;
                        humGain = ctx.createGain(); humGain.gain.value = 0.0001;
                        hum.connect(lp); lp.connect(humGain); humGain.connect(master);
                        hum.start();
                        humGain.gain.exponentialRampToValueAtTime(0.06, ctx.currentTime + 1.2);
                    }catch(e){}
                },
                fastHum(on){ if(hum){ try{ hum.frequency.exponentialRampToValueAtTime(on?92:58, ctx.currentTime+0.6);}catch(e){} } },
                stopHum(){ try{ if(humGain){ humGain.gain.exponentialRampToValueAtTime(0.0001, ctx.currentTime+0.5);} if(hum){ const h=hum; setTimeout(()=>{try{h.stop();}catch(e){}},600); } hum=null; }catch(e){} },
                ding(){
                    if(!ensure()) return; resume();
                    try{
                        [880,660].forEach((f,i)=>{
                            const o=ctx.createOscillator(); o.type="sine"; o.frequency.value=f;
                            const g=ctx.createGain(); g.gain.value=0.0001;
                            o.connect(g); g.connect(master);
                            const s=ctx.currentTime + i*0.18;
                            g.gain.exponentialRampToValueAtTime(0.12, s+0.02);
                            g.gain.exponentialRampToValueAtTime(0.0001, s+0.45);
                            o.start(s); o.stop(s+0.5);
                        });
                    }catch(e){}
                }
            };
        })();

        /* ---- ripple (3) ---- */
        function attachRipple(el){
            el.addEventListener("pointerdown", e=>{
                const r = el.getBoundingClientRect();
                const s = document.createElement("span");
                s.className = "ripple";
                const size = Math.max(r.width, r.height) * 1.4;
                s.style.width = s.style.height = size + "px";
                s.style.left = (e.clientX - r.left) + "px";
                s.style.top = (e.clientY - r.top) + "px";
                el.appendChild(s);
                setTimeout(()=>s.remove(), 620);
            });
        }

        /* ---- sayı sayması (2/57) ---- */
        function countUp(el, target, duration=1100){
            const start = performance.now();
            const from = 0;
            function frame(now){
                const p = Math.min((now-start)/duration, 1);
                const eased = 1 - Math.pow(1-p, 3);
                el.textContent = formatNumber(from + (target-from)*eased);
                if(p<1) requestAnimationFrame(frame); else el.textContent = formatNumber(target);
            }
            requestAnimationFrame(frame);
        }

        async function login(event){
            event.preventDefault();
            Audio.init();
            const username = document.getElementById("username").value.trim();
            const password = document.getElementById("password").value.trim();
            const card = document.getElementById("loginPanel");
            if(!username || !password){ showToast(t("fillFields"), "error"); card.classList.remove("shake"); void card.offsetWidth; card.classList.add("shake"); return; }
            const button = document.getElementById("loginBtn");
            button.disabled = true; button.classList.add("loading");
            try{
                const response = await fetch("/login", {method:"POST", headers:{"Content-Type":"application/json"}, body:JSON.stringify({username, password})});
                const result = await response.json();
                if(!response.ok || result.status !== "success"){
                    card.classList.remove("shake"); void card.offsetWidth; card.classList.add("shake");
                    buzz([60,40,60]);
                    if(response.status === 429) showToast(result.message || t("blocked"), "error");
                    else showToast(result.message || t("loginError"), "error");
                    return;
                }
                // 1) kullanıcı adını hatırla
                if(document.getElementById("rememberMe").checked) localStorage.setItem("izin-remember-user", username);
                else localStorage.removeItem("izin-remember-user");
                currentUser = result.data;
                const loginKey = `izin-last-login-${currentUser.username}`;
                currentUser.previous_login = localStorage.getItem(loginKey);
                localStorage.setItem(loginKey, new Date().toISOString());
                document.getElementById("password").value = "";
                // 76) yeşil onay dalgası, sonra kapılar
                card.classList.add("success");
                await new Promise(r=>setTimeout(r, 650));
                await playElevatorAnimation(currentUser);
            }catch(error){
                console.error(error); showToast(t("serverError"), "error");
            }finally{
                button.disabled = false; button.classList.remove("loading");
                document.getElementById("loginPanel").classList.remove("success");
            }
        }

        /* ---- Animasyon yardımcıları ---- */
        function delay(ms, runId){
            return new Promise(resolve=>{
                const timer = setTimeout(()=>resolve(true), ms);
                const checker = setInterval(()=>{
                    if(runId !== animationRunId){ clearTimeout(timer); clearInterval(checker); resolve(false); }
                }, 80);
                setTimeout(()=>clearInterval(checker), ms + 120);
            });
        }
        function nextFrames(n=2){ return new Promise(res=>{ let i=0; (function f(){ i++; if(i>=n) res(); else requestAnimationFrame(f); })(); }); }
        function setStatus(key){ document.getElementById("elevatorStatus").textContent = t(key); }
        function simClass(add, remove){
            const sim = document.getElementById("sim");
            (remove||[]).forEach(c=>sim.classList.remove(c));
            (add||[]).forEach(c=>sim.classList.add(c));
        }
        function pulse(cls, ms=600){ const sim=document.getElementById("sim"); sim.classList.add(cls); setTimeout(()=>sim.classList.remove(cls), ms); }

        function animateFloor(target, duration, runId){
            return new Promise(resolve=>{
                const start = performance.now();
                const floor = document.getElementById("floorValue");
                const ind = document.querySelector(".floor-ind");
                const sim = document.getElementById("sim");
                let lastTen = 0;
                function ease(p){
                    if(p < .82){ const q = p / .82; return .9 * q * q * q; }
                    const q = (p - .82) / .18; return .9 + .1 * (1 - Math.pow(1 - q, 2));
                }
                function frame(now){
                    if(runId !== animationRunId){ sim.classList.remove("fast"); resolve(false); return; }
                    const progress = Math.min((now - start) / duration, 1);
                    sim.classList.toggle("fast", progress > .3 && progress < .88);
                    const val = Math.floor(target * ease(progress));
                    floor.textContent = formatNumber(val);
                    // her 10 katta hafif parlama
                    const ten = Math.floor(val/10);
                    if(ten !== lastTen && val > 0){ lastTen = ten; ind.classList.remove("flash"); void ind.offsetWidth; ind.classList.add("flash"); }
                    if(progress < 1) requestAnimationFrame(frame);
                    else { sim.classList.remove("fast"); floor.textContent = formatNumber(target); resolve(true); }
                }
                requestAnimationFrame(frame);
            });
        }

        function resetSim(){
            const sim = document.getElementById("sim");
            sim.className = "sim";
            document.getElementById("floorValue").textContent = "0";
            document.getElementById("lobbyDisplay").textContent = "G";
            document.querySelector(".dept-btn.target").classList.remove("lit","pressed");
            document.querySelector(".floor-ind").classList.remove("ding","flash");
            document.getElementById("docBalance").textContent = "0";
        }


        async function playElevatorAnimation(user){
            const runId = ++animationRunId;
            resetSim();
            const sim = document.getElementById("sim");
            const loginPanel = document.getElementById("loginPanel");
            const dashboard = document.getElementById("dashboard");
            const entry = document.getElementById("entryDoors");

            // Belge hazırlığı — direkt onaylı
            document.getElementById("docTitle").textContent = t("noteTitle")(user.name);
            document.getElementById("docBalance").textContent = formatNumber(user.remaining_leave);
            const now = new Date();
            const stamp = `${now.getFullYear()}${String(now.getMonth()+1).padStart(2,"0")}${String(now.getDate()).padStart(2,"0")}-${String(now.getHours()).padStart(2,"0")}${String(now.getMinutes()).padStart(2,"0")}`;
            const uid = String(user.username).replace(/\D/g,"").padStart(4,"0").slice(-4) || "0000";
            const ref = `RET-${stamp}-${uid}`;
            document.getElementById("docRef").textContent = `Ref: ${ref}`;
            drawDocQr({ref, id:user.username, name:user.name, bal:user.remaining_leave});

            // Giriş kapıları kapanır -> arkada sim hazırlanır (ilk render gizlenir) -> aralanır
            entry.classList.add("active");
            await nextFrames(2);
            entry.classList.add("closing");
            await delay(950, runId); if(runId!==animationRunId){ entry.classList.remove("active","closing"); return; }
            loginPanel.hidden = true; dashboard.hidden = true;
            sim.hidden = false;
            await nextFrames(3);
            entry.classList.remove("closing");
            await delay(900, runId);
            entry.classList.remove("active");

            const target = 67;

            setStatus("stCalling");
            if(!(await delay(650, runId))) return;
            simClass(["lobby-open"]); setStatus("stDoorsOpen");
            if(!(await delay(1800, runId))) return;

            simClass(["stepped"]); setStatus("stEntering");
            if(!(await delay(1400, runId))) return;

            // Departman seçimi: el gelir, buton basılır ve seçili kalır
            simClass(["picking"]); setStatus("stPicking"); buzz(25);
            if(!(await delay(1200, runId))) return;
            const target5 = document.querySelector(".dept-btn.target");
            target5.classList.add("pressed");
            if(!(await delay(600, runId))) return;
            target5.classList.add("lit");
            simClass([], ["picking"]);
            if(!(await delay(500, runId))) return;

            // Yolculuk: kısa, akıcı (motor sesi + gösterge)
            simClass(["riding","bump"]); setStatus("stRising");
            Audio.startHum();
            setTimeout(()=>{ if(runId===animationRunId) Audio.fastHum(true); }, 2500);
            setTimeout(()=>{ if(runId===animationRunId) sim.classList.remove("bump"); }, 700);
            const floorDone = await animateFloor(target, 8000, runId);
            if(!floorDone || runId !== animationRunId){ Audio.stopHum(); return; }
            Audio.fastHum(false); Audio.stopHum();
            simClass(["bump"], ["riding"]);
            setTimeout(()=>sim.classList.remove("bump"), 600);
            document.getElementById("floorValue").textContent = formatNumber(target);
            document.querySelector(".floor-ind").classList.add("ding");
            Audio.ding(); buzz([40, 60, 40]);
            setStatus("stArrived");
            if(!(await delay(1700, runId))) return;   // 67 sayısı net görünsün

            // Varış: kapılar açılır -> aydınlık kat görünür, ekip belirir
            simClass(["arrived"]); setStatus("stDoorsOpen");
            if(!(await delay(1700, runId))) return;
            simClass(["exiting","in-office"]); setStatus("stCrew");
            if(!(await delay(2000, runId))) return;   // 5 kişilik ekip belirsin

            // Semih Bayat öne çıkar, adı görünür — sahne biraz uzun
            simClass(["crew-step"]); setStatus("stSemih"); buzz(20);
            if(!(await delay(4800, runId))) return;

            // Onaylı belge yumuşakça merkeze gelir
            simClass(["reading"]); setStatus("stDelivery");
            buzz(30);
            if(!(await delay(3200, runId))) return;

            finishAnimation(runId);
        }

        function finishAnimation(runId=null){
            if(runId !== null && runId !== animationRunId) return;
            animationRunId++;
            Audio.stopHum();
           
            // 36) kapanış logosu
            const lf = document.getElementById("logoFlash");
            lf.classList.add("show");
            setTimeout(()=>{
                document.getElementById("sim").hidden = true;
                resetSim();
                renderUser(currentUser, true);
                document.getElementById("dashboard").hidden = false;
                window.scrollTo({top:0, behavior:"smooth"});
            }, 550);
            setTimeout(()=>lf.classList.remove("show"), 1150);
        }

        function getGreeting(name){
            const hour = new Date().getHours();
            const firstName = String(name || "").split(" ")[0];
            if(hour < 12) return `${t("greetingMorning")}, ${firstName}`;
            if(hour >= 18) return `${t("greetingEvening")}, ${firstName}`;
            return `${t("greetingDay")}, ${firstName}`;
        }

        function renderUser(user, animate=true){
            if(!user) return;
            document.getElementById("avatar").textContent = user.initials || "P";
            document.getElementById("greeting").textContent = getGreeting(user.name);
            document.getElementById("personRole").textContent = user.role || "Personel";
            document.getElementById("personUsername").textContent = user.username || "";
            const bal = Number(user.remaining_leave || 0);
            const remEl = document.getElementById("remainingLeave");
            if(animate){ countUp(remEl, bal, 1200); }
            else { remEl.textContent = formatNumber(bal); }
            document.getElementById("sundayLeave").textContent = formatNumber(user.sunday_leave);
            document.getElementById("officialHoliday").textContent = formatNumber(user.official_holiday);
            document.getElementById("updatedAt").textContent = user.updated_at || "-";
            document.getElementById("lastLogin").textContent = formatDateTime(user.previous_login);
            document.getElementById("idName").textContent = user.name;
            document.getElementById("idRole").textContent = user.role;
            document.getElementById("idNumber").textContent = user.username;
            document.getElementById("objectionBtn").href = `https://wa.me/${WHATSAPP_NUMBER}?text=${encodeURIComponent(t("objectionMessage")(user.name))}`;
            drawQr(document.getElementById("qrCanvas"), {portal:"Personel İzin Portalı", id:user.username, name:user.name, role:user.role}, 156);
            calculateNextHoliday();
        }

        function drawQr(canvas, data, size){
            const value = JSON.stringify(data);
            if(window.QRious){
                new QRious({element:canvas, value, size, level:"M", background:"white", foreground:"#07263c"});
                return;
            }
            const ctx = canvas.getContext("2d");
            const cells = 21, cell = size / cells; canvas.width = size; canvas.height = size;
            ctx.fillStyle = "white"; ctx.fillRect(0,0,size,size);
            let seed = 0; for(const ch of value) seed = (seed*31 + ch.charCodeAt(0)) >>> 0;
            function rnd(){ seed = (seed*1664525 + 1013904223) >>> 0; return seed/4294967296; }
            ctx.fillStyle = "#07263c";
            for(let y=0;y<cells;y++) for(let x=0;x<cells;x++) if(rnd() > .53) ctx.fillRect(x*cell,y*cell,Math.ceil(cell),Math.ceil(cell));
            [[1,1],[13,1],[1,13]].forEach(([x,y])=>{
                ctx.fillStyle="#07263c"; ctx.fillRect(x*cell,y*cell,7*cell,7*cell);
                ctx.fillStyle="white"; ctx.fillRect((x+1)*cell,(y+1)*cell,5*cell,5*cell);
                ctx.fillStyle="#07263c"; ctx.fillRect((x+2)*cell,(y+2)*cell,3*cell,3*cell);
            });
        }
        function drawDocQr(data){ drawQr(document.getElementById("docQr"), data, 120); }

        function calculateNextHoliday(){
            const now = new Date(); now.setHours(0,0,0,0);
            const year = now.getFullYear();
            const fixed = [
                [1,1,"Yılbaşı","New Year's Day"],
                [4,23,"Ulusal Egemenlik ve Çocuk Bayramı","National Sovereignty and Children's Day"],
                [5,1,"Emek ve Dayanışma Günü","Labour and Solidarity Day"],
                [5,19,"Atatürk'ü Anma, Gençlik ve Spor Bayramı","Commemoration of Atatürk, Youth and Sports Day"],
                [7,15,"Demokrasi ve Millî Birlik Günü","Democracy and National Unity Day"],
                [8,30,"Zafer Bayramı","Victory Day"],
                [10,29,"Cumhuriyet Bayramı","Republic Day"]
            ];
            let candidates = [];
            [year, year+1].forEach(y=>{ fixed.forEach(([m,d,tr,en])=>{ candidates.push({date:new Date(y,m-1,d), tr, en}); }); });
            candidates = candidates.filter(i=>i.date >= now).sort((a,b)=>a.date-b.date);
            const next = candidates[0]; if(!next) return;
            const diff = Math.ceil((next.date - now) / 86400000);
            document.getElementById("holidayDays").textContent = t("holidayIn")(diff);
            document.getElementById("holidayName").textContent = currentLanguage === "tr" ? next.tr : next.en;
            document.getElementById("holidayDate").textContent = new Intl.DateTimeFormat(currentLanguage === "tr" ? "tr-TR" : "en-GB", {day:"2-digit", month:"long", year:"numeric"}).format(next.date);
        }

        async function installPwa(){
            if(deferredInstallPrompt){
                deferredInstallPrompt.prompt();
                await deferredInstallPrompt.userChoice;
                deferredInstallPrompt = null;
                document.getElementById("installBtn").classList.remove("show");
                return;
            }
            const isIos = /iphone|ipad|ipod/i.test(navigator.userAgent);
            showToast(isIos ? t("iosInstall") : t("installUnavailable"));
        }
        function setupPwa(){
            if("serviceWorker" in navigator){
                window.addEventListener("load", ()=>{ navigator.serviceWorker.register("/service-worker.js").catch(console.error); });
            }
            window.addEventListener("beforeinstallprompt", event=>{
                event.preventDefault(); deferredInstallPrompt = event;
                document.getElementById("installBtn").classList.add("show");
                showToast(t("installReady"), "success");
            });
            window.addEventListener("appinstalled", ()=>{ deferredInstallPrompt = null; document.getElementById("installBtn").classList.remove("show"); });
        }

        /* ---- olay bağlantıları ---- */
        document.getElementById("loginForm").addEventListener("submit", login);
        document.getElementById("themeBtn").addEventListener("click", toggleTheme);
        document.getElementById("languageBtn").addEventListener("click", ()=>{
            currentLanguage = currentLanguage === "tr" ? "en" : "tr";
            localStorage.setItem("izin-language", currentLanguage);
            applyLanguage();
        });
        document.getElementById("passwordToggle").addEventListener("click", ()=>{
            const p = document.getElementById("password");
            p.type = p.type === "password" ? "text" : "password";
        });
        // 95) caps lock uyarısı
        function capsCheck(e){ const on = e.getModifierState && e.getModifierState("CapsLock"); document.getElementById("capsHint").classList.toggle("show", !!on); }
        document.getElementById("password").addEventListener("keyup", capsCheck);
        document.getElementById("password").addEventListener("keydown", capsCheck);
        document.getElementById("password").addEventListener("blur", ()=>document.getElementById("capsHint").classList.remove("show"));
        document.getElementById("forgotBtn").addEventListener("click", ()=>openModal("forgotModal"));
        document.getElementById("leaveRequestBtn").addEventListener("click", ()=>openModal("leaveModal"));
        document.getElementById("installBtn").addEventListener("click", installPwa);
        document.getElementById("installActionBtn").addEventListener("click", installPwa);
        document.getElementById("replayBtn").addEventListener("click", ()=>{ Audio.init(); playElevatorAnimation(currentUser); });
        document.getElementById("logoutBtn").addEventListener("click", ()=>{
            animationRunId++; Audio.stopHum(); currentUser = null;
            document.getElementById("dashboard").hidden = true;
            document.getElementById("loginPanel").hidden = false;
            const ru = localStorage.getItem("izin-remember-user");
            if(ru){ document.getElementById("username").value = ru; document.getElementById("password").focus(); }
            else document.getElementById("username").focus();
        });
        document.querySelectorAll("[data-close]").forEach(b=>b.addEventListener("click", ()=>closeModal(b.dataset.close)));
        document.querySelectorAll(".modal-backdrop").forEach(bg=>bg.addEventListener("click", e=>{ if(e.target === bg) closeModal(bg.id); }));
        document.getElementById("forgotWhatsappBtn").addEventListener("click", ()=>{
            const identity = document.getElementById("forgotIdentity").value.trim();
            if(!identity){ showToast(t("fillIdentity"), "error"); return; }
            openWhatsApp(t("forgotMessage")(identity)); closeModal("forgotModal");
        });
        document.getElementById("sendLeaveBtn").addEventListener("click", ()=>{
            if(!currentUser) return;
            const start = document.getElementById("leaveStart").value;
            const end = document.getElementById("leaveEnd").value;
            if(!start || !end){ showToast(t("fillDates"), "error"); return; }
            const s = new Date(`${start}T00:00:00`), e = new Date(`${end}T00:00:00`);
            if(e < s){ showToast(t("invalidDates"), "error"); return; }
            const days = Math.floor((e - s) / 86400000) + 1;
            const data = {
                type:document.getElementById("leaveType").value,
                start:new Intl.DateTimeFormat(currentLanguage === "tr" ? "tr-TR" : "en-GB").format(s),
                end:new Intl.DateTimeFormat(currentLanguage === "tr" ? "tr-TR" : "en-GB").format(e),
                days, description:document.getElementById("leaveDescription").value.trim()
            };
            openWhatsApp(t("leaveMessage")(data)); closeModal("leaveModal");
        });
        window.addEventListener("keydown", e=>{ if(e.key === "Escape") document.querySelectorAll(".modal-backdrop.open").forEach(m=>closeModal(m.id)); });

        // ripple: tüm butonlar
        document.querySelectorAll(".primary-btn,.action-btn,.dept-btn,.small-btn").forEach(attachRipple);

        // init
        (function init(){
            autoTheme();
            applyLanguage();
            setupPwa();
            const ru = localStorage.getItem("izin-remember-user");
            if(ru){ document.getElementById("username").value = ru; document.getElementById("rememberMe").checked = true; }
        })();
    </script>
</body>
</html>'''

SERVICE_WORKER = """
const CACHE_NAME = "izin-portali-v3";
const CORE_ASSETS = ["/", "/manifest.webmanifest", "/icon-192.png", "/icon-512.png"];

self.addEventListener("install", (event) => {
    event.waitUntil(caches.open(CACHE_NAME).then((cache) => cache.addAll(CORE_ASSETS)).then(() => self.skipWaiting()));
});

self.addEventListener("activate", (event) => {
    event.waitUntil(
        caches.keys().then((keys) => Promise.all(keys.filter((key) => key !== CACHE_NAME).map((key) => caches.delete(key)))).then(() => self.clients.claim())
    );
});

self.addEventListener("fetch", (event) => {
    const request = event.request;
    if (request.method !== "GET") return;
    const url = new URL(request.url);
    if (url.pathname === "/login") return;

    if (request.mode === "navigate") {
        event.respondWith(fetch(request).catch(() => caches.match("/")));
        return;
    }

    event.respondWith(
        caches.match(request).then((cached) => {
            if (cached) return cached;
            return fetch(request).then((response) => {
                if (response && response.status === 200 && url.origin === self.location.origin) {
                    const clone = response.clone();
                    caches.open(CACHE_NAME).then((cache) => cache.put(request, clone));
                }
                return response;
            }).catch(() => cached);
        })
    );
});
"""


@app.after_request
def set_security_headers(response):
    response.headers["X-Content-Type-Options"] = "nosniff"
    response.headers["X-Frame-Options"] = "SAMEORIGIN"
    response.headers["Referrer-Policy"] = "strict-origin-when-cross-origin"
    response.headers["Permissions-Policy"] = "geolocation=(), microphone=(), camera=()"
    return response


@app.route("/")
def index():
    return Response(HTML_SAYFASI, mimetype="text/html; charset=utf-8")


@app.route("/login", methods=["POST"])
def login():
    ip_address = client_ip()
    blocked, retry_after = login_is_blocked(ip_address)
    if blocked:
        minutes = max(1, int(math.ceil(retry_after / 60)))
        response = jsonify({
            "status": "error",
            "message": f"Çok fazla hatalı deneme. Lütfen {minutes} dakika sonra tekrar deneyin.",
        })
        response.status_code = 429
        response.headers["Retry-After"] = str(retry_after)
        return response

    payload = request.get_json(silent=True) or {}
    username = str(payload.get("username", "")).strip()
    password = str(payload.get("password", "")).strip()

    if not username or not password:
        return jsonify({"status": "error", "message": "Kullanıcı adı ve şifre gereklidir."}), 400

    try:
        user_data = get_user_data(username, password)
    except FileNotFoundError as error:
        return jsonify({"status": "error", "message": str(error)}), 500
    except ValueError as error:
        return jsonify({"status": "error", "message": str(error)}), 500
    except Exception:
        return jsonify({"status": "error", "message": "Beklenmeyen bir hata oluştu. Lütfen tekrar deneyin."}), 500

    if user_data is None:
        record_failed_login(ip_address)
        return jsonify({"status": "error", "message": "Kullanıcı adı veya şifre hatalı."}), 401

    clear_failed_logins(ip_address)
    return jsonify({"status": "success", "data": user_data})


@app.route("/manifest.webmanifest")
def manifest():
    data = {
        "name": "Personel İzin Portalı",
        "short_name": "İzin Portalı",
        "description": "Personel izin hakları görüntüleme ve izin talep sistemi",
        "start_url": "/",
        "scope": "/",
        "display": "standalone",
        "orientation": "portrait",
        "background_color": "#061426",
        "theme_color": "#071b30",
        "lang": "tr",
        "icons": [
            {"src": "/icon-192.png", "sizes": "192x192", "type": "image/png", "purpose": "any maskable"},
            {"src": "/icon-512.png", "sizes": "512x512", "type": "image/png", "purpose": "any maskable"},
        ],
    }
    return jsonify(data)


@app.route("/service-worker.js")
def service_worker():
    response = Response(SERVICE_WORKER, mimetype="application/javascript")
    response.headers["Cache-Control"] = "no-cache"
    return response


@app.route("/icon-192.png")
def icon_192():
    return Response(ICON_192_BYTES, mimetype="image/png")


@app.route("/icon-512.png")
def icon_512():
    return Response(ICON_512_BYTES, mimetype="image/png")


@app.route("/health")
def health():
    return jsonify({"status": "ok"})


if __name__ == "__main__":
    port = int(os.getenv("PORT", "5000"))
    app.run(host="0.0.0.0", port=port, debug=False)
