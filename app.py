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
            --panel:rgba(9,35,59,.76); --panel-strong:rgba(6,26,46,.94);
            --text:#f4fbff; --muted:#a8c4d8;
            --primary:#55d9ff; --primary-2:#1489c9;
            --success:#2bd881; --warning:#ffca57; --danger:#ff6678;
            --line:rgba(143,222,255,.22);
            --shadow:0 28px 70px rgba(0,0,0,.34); --radius:24px;
        }
        body.light{
            --bg-1:#dff4ff; --bg-2:#eff9ff;
            --panel:rgba(255,255,255,.78); --panel-strong:rgba(255,255,255,.96);
            --text:#10283d; --muted:#527089;
            --primary:#087fc4; --primary-2:#23b5e8;
            --line:rgba(14,116,174,.16); --shadow:0 28px 70px rgba(39,96,131,.18);
        }
        *{box-sizing:border-box;}
        html{min-height:100%; background:var(--bg-1);}
        body{
            margin:0; min-height:100vh; min-height:100dvh; color:var(--text);
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
        .brand-mini,.top-actions{pointer-events:auto; display:flex; align-items:center; gap:8px;}
        .brand-mini{
            padding:9px 13px; border:1px solid var(--line); border-radius:16px;
            background:var(--panel); backdrop-filter:blur(18px);
            box-shadow:0 12px 30px rgba(0,0,0,.14); font-weight:800; letter-spacing:.03em;
        }
        .brand-mark{
            width:28px; height:28px; border-radius:9px; display:grid; place-items:center;
            background:linear-gradient(145deg,var(--primary),var(--primary-2));
            color:#042039; font-weight:1000;
        }
        .icon-btn,.install-btn{
            border:1px solid var(--line); background:var(--panel); backdrop-filter:blur(18px);
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
            backdrop-filter:blur(22px); -webkit-backdrop-filter:blur(22px);
            box-shadow:var(--shadow); border-radius:var(--radius);
        }

        /* ---------- LOGIN ---------- */
        .login-card{
            width:min(440px,100%); padding:clamp(24px,6vw,40px); position:relative; overflow:hidden;
            animation:enterCard .6s cubic-bezier(.2,.8,.2,1) both;
        }
        @keyframes enterCard{from{opacity:0; transform:translateY(22px) scale(.97);}to{opacity:1; transform:none;}}
        .login-logo{
            width:78px; height:78px; margin:0 auto 18px; border-radius:24px; display:grid; place-items:center;
            background:linear-gradient(145deg,rgba(76,216,255,.22),rgba(18,103,164,.42));
            border:1px solid rgba(111,225,255,.4);
            box-shadow:inset 0 0 30px rgba(79,218,255,.12), 0 14px 34px rgba(0,0,0,.18);
        }
        .elevator-icon{width:42px; height:48px; border:2px solid var(--primary); border-radius:7px; position:relative; overflow:hidden;}
        .elevator-icon::before,.elevator-icon::after{content:""; position:absolute; top:0; bottom:0; width:50%; background:rgba(82,217,255,.14);}
        .elevator-icon::before{left:0; border-right:1px solid var(--primary);}
        .elevator-icon::after{right:0;}
        .login-card h1{text-align:center; margin:0; font-size:clamp(1.55rem,5vw,2.05rem); letter-spacing:-.03em;}
        .login-subtitle{text-align:center; color:var(--muted); margin:9px 0 28px; line-height:1.5;}
        .field{margin-bottom:16px;}
        .field label{display:block; margin-bottom:8px; color:var(--muted); font-size:.9rem; font-weight:700;}
        .input-wrap{position:relative;}
        .input-wrap input,.modal input,.modal select,.modal textarea{
            width:100%; color:var(--text); border:1px solid var(--line);
            background:rgba(255,255,255,.055); border-radius:15px; padding:15px 46px 15px 15px; outline:none;
            transition:border-color .2s ease, box-shadow .2s ease, background .2s ease;
        }
        body.light .input-wrap input,body.light .modal input,body.light .modal select,body.light .modal textarea{background:rgba(8,70,110,.04);}
        input::placeholder,textarea::placeholder{color:color-mix(in srgb, var(--muted) 75%, transparent);}
        .input-wrap input:focus,.modal input:focus,.modal select:focus,.modal textarea:focus{
            border-color:var(--primary); box-shadow:0 0 0 4px rgba(75,207,255,.13); background:rgba(255,255,255,.08);
        }
        .field-icon{position:absolute; right:14px; top:50%; transform:translateY(-50%); color:var(--muted); pointer-events:none;}
        .password-toggle{pointer-events:auto; border:0; background:transparent; padding:4px; cursor:pointer;}
        .primary-btn,.action-btn{
            width:100%; border:0; border-radius:15px; padding:15px 18px; font-weight:900; cursor:pointer;
            display:inline-flex; justify-content:center; align-items:center; gap:9px; text-decoration:none;
            transition:transform .2s ease, filter .2s ease;
        }
        .primary-btn{background:linear-gradient(135deg,var(--primary),var(--primary-2)); color:#032039; box-shadow:0 14px 30px rgba(34,172,229,.28);}
        .primary-btn:hover,.action-btn:hover{transform:translateY(-2px); filter:brightness(1.05);}
        .primary-btn:active,.action-btn:active{transform:translateY(1px) scale(.99);}
        .primary-btn:disabled{opacity:.65; cursor:wait;}
        .spinner{width:19px; height:19px; border-radius:50%; border:2px solid rgba(3,32,57,.25); border-top-color:#032039; animation:spin .8s linear infinite; display:none;}
        .loading .spinner{display:block;}
        @keyframes spin{to{transform:rotate(360deg);}}
        .login-links{display:flex; justify-content:center; margin-top:17px;}
        .text-btn{border:0; background:transparent; color:var(--primary); cursor:pointer; font-weight:800; padding:7px;}

        /* ============================================================
           KESİNTİSİZ ASANSÖR SİMÜLASYONU
           Kural: alttaki katmanlar HER ZAMAN opak. Üstteki katmanlar
           sadece SÖNEREK alttakini gösterir. Böylece hiçbir anda
           siyah ekran oluşmaz. Kamera asla sağa-sola dönmez.
           ============================================================ */
        .sim{
            position:fixed; inset:0; z-index:100; overflow:hidden;
            background:#0c141c; color:#eef7ff;
            font-family:Inter,"Segoe UI",system-ui,sans-serif;
        }
        .sim[hidden]{display:none !important;}
        .sim *{box-sizing:border-box;}
        .sim-world{position:absolute; inset:0; transform-origin:50% 55%; transition:transform 1.6s cubic-bezier(.22,.78,.18,1); will-change:transform;}

        /* Katmanlar (alt -> üst): office < cabin < lobby */
        .stage{position:absolute; inset:0; transition:opacity 1.1s ease, transform 1.6s cubic-bezier(.22,.78,.18,1);}
        .stage.office{z-index:1; opacity:1;}           /* daima görünür */
        .stage.cabin{z-index:2; opacity:1;}
        .stage.lobby{z-index:3; opacity:1;}

        /* ---- ORTAK KAPI ---- */
        .doorway{position:absolute; overflow:hidden; background:transparent;}
        .door{
            position:absolute; top:0; bottom:0; width:50.4%;
            background:
                repeating-linear-gradient(90deg, rgba(255,255,255,.05) 0 2px, transparent 2px 15px),
                linear-gradient(90deg,#2b3c47,#7c8f9a 48%,#2e3f4a);
            box-shadow:inset 0 0 30px rgba(0,0,0,.42);
            transition:transform 2.1s cubic-bezier(.7,.02,.18,1);
        }
        .door.l{left:0; border-right:1px solid #0b1217;}
        .door.r{right:0; border-left:1px solid rgba(255,255,255,.16); transform:scaleX(-1);}

        /* ---- LOBİ ---- */
        .stage.lobby .lobby-wall{
            position:absolute; inset:0;
            background:
                linear-gradient(180deg,#e7eef2 0 16%,#cfd9df 16% 74%,#7c8892 74% 100%);
        }
        .stage.lobby .lobby-wall::after{
            content:""; position:absolute; left:-10%; right:-10%; bottom:-18%; height:46%;
            background:
                repeating-linear-gradient(90deg, rgba(24,49,63,.16) 0 1px, transparent 1px 12%),
                linear-gradient(#8a97a0,#d3dde1);
            transform:perspective(620px) rotateX(62deg); transform-origin:top;
        }
        .lobby-frame{
            position:absolute; left:50%; top:9%; bottom:9%; width:min(72vw,560px);
            transform:translateX(-50%);
            border:clamp(10px,1.6vw,20px) solid #56646f; border-bottom-width:26px;
            background:#0a1016;
            box-shadow:0 22px 55px rgba(0,0,0,.35), inset 0 0 0 2px rgba(255,255,255,.12);
        }
        .lobby-head{
            position:absolute; left:50%; top:-46px; transform:translateX(-50%);
            min-width:110px; padding:8px 16px; border-radius:8px; text-align:center;
            color:#6fe6ff; background:#03080d; border:1px solid #486370;
            font:900 1.15rem/1 ui-monospace,monospace; letter-spacing:.18em;
            text-shadow:0 0 12px rgba(91,226,255,.7);
        }
        .lobby-doorway{inset:0;}
        /* kapı açılınca alttaki KABİN görünür (transparan doorway) */
        .sim.lobby-open .stage.lobby .door.l{transform:translateX(-102%);}
        .sim.lobby-open .stage.lobby .door.r{transform:scaleX(-1) translateX(-102%);}
        /* içeri girince lobi söner, altındaki kabin (opak) ortaya çıkar -> siyahlık yok */
        .sim.stepped .stage.lobby{opacity:0; transform:scale(1.35);}

        /* ---- KABİN ---- */
        .stage.cabin{
            background:
                linear-gradient(90deg,#0a1620 0 20%,#1a2b37 20% 80%,#0a1620 80%),
                #0d1a24;
        }
        .cabin-ceil{position:absolute; left:-6%; right:-6%; top:-10%; height:30%; background:linear-gradient(#1a2a34,#0a141c); transform:perspective(680px) rotateX(-52deg); box-shadow:0 30px 60px rgba(0,0,0,.5);}
        .cabin-light{position:absolute; top:5%; left:50%; transform:translateX(-50%); width:min(46vw,460px); height:16px; border-radius:50%; background:#dcf8ff; box-shadow:0 0 26px #8ee9ff,0 0 80px rgba(84,211,255,.34); transition:opacity .6s;}
        .cabin-floor{
            position:absolute; left:-10%; right:-10%; bottom:-14%; height:34%;
            background:
                repeating-linear-gradient(90deg, rgba(140,200,225,.10) 0 1px, transparent 1px 9%),
                linear-gradient(#16303f,#0a151d);
            transform:perspective(560px) rotateX(60deg); transform-origin:top;
        }
        .stage.cabin .wall{position:absolute; top:12%; bottom:0; width:20%; background:linear-gradient(90deg,rgba(255,255,255,.04),transparent),linear-gradient(#12242f,#0a151d); border:1px solid rgba(150,215,240,.10);}
        .stage.cabin .wall.left{left:0; clip-path:polygon(0 0,100% 9%,100% 100%,0 100%);}
        .stage.cabin .wall.right{right:0; transform:scaleX(-1); clip-path:polygon(0 0,100% 9%,100% 100%,0 100%);}

        /* Kat göstergesi (kapı üstünde, kabin içinde) */
        .floor-ind{
            position:absolute; z-index:20; left:50%; top:max(4%,env(safe-area-inset-top)); transform:translateX(-50%);
            min-width:150px; padding:9px 18px 7px; text-align:center; border-radius:11px;
            color:#69e8ff; background:#02080d; border:1px solid #4e7384;
            box-shadow:inset 0 0 16px rgba(89,221,255,.12),0 0 22px rgba(59,207,247,.16);
        }
        .floor-ind b{display:block; font:1000 clamp(1.5rem,6vw,2.3rem)/1 ui-monospace,monospace; letter-spacing:.1em; text-shadow:0 0 14px rgba(80,225,255,.75);}
        .floor-ind small{display:block; margin-top:4px; color:#94adba; font-size:.6rem; letter-spacing:.24em; font-weight:900;}
        .half-tag{
            position:absolute; z-index:21; left:50%; top:calc(max(4%,env(safe-area-inset-top)) + 78px); transform:translate(-50%,-8px);
            opacity:0; padding:6px 12px; border-radius:999px; color:#ffe18c;
            background:rgba(45,33,7,.92); border:1px solid rgba(255,220,116,.5);
            font-size:.66rem; font-weight:1000; letter-spacing:.1em; transition:.5s ease;
        }
        .half-tag.show{opacity:1; transform:translate(-50%,0);}

        /* Kabin kapısı — ortada */
        .cabin-doorway{left:24%; right:24%; top:16%; bottom:2%;
            border:clamp(7px,1.4vw,16px) solid #3b5361; border-bottom-width:20px; border-radius:2px;
            box-shadow:inset 0 0 0 2px rgba(255,255,255,.1),0 18px 45px rgba(0,0,0,.5);
        }
        .sim.arrived .stage.cabin .cabin-doorway .door.l{transform:translateX(-102%);}
        .sim.arrived .stage.cabin .cabin-doorway .door.r{transform:scaleX(-1) translateX(-102%);}
        /* ofise çık: kabin söner + ofis büyür -> siyahlık yok çünkü ofis altta opak */
        .sim.exiting .stage.cabin{opacity:0; transform:scale(1.28);}

        /* Departman paneli — İÇERİDE, sağda, GENİŞ, yazılar tam görünür */
        .dept-panel{
            position:absolute; z-index:22; right:2.5%; top:20%; width:min(38vw,260px);
            padding:12px; border-radius:15px;
            background:linear-gradient(150deg,#43576a,#17262f 55%,#0c1820);
            border:1px solid rgba(200,235,250,.26);
            box-shadow:0 20px 40px rgba(0,0,0,.34), inset 0 0 20px rgba(255,255,255,.04);
            transition:opacity .8s ease;
        }
        .dept-title{margin-bottom:9px; color:#a7bfcb; text-align:center; font-size:.62rem; font-weight:1000; letter-spacing:.18em;}
        .dept-btn{
            width:100%; min-height:46px; margin:6px 0; padding:9px 11px; border-radius:10px;
            display:flex; align-items:center; gap:9px; text-align:left; color:#dce9ee;
            background:linear-gradient(#1b2b36,#0c161d); border:1px solid #5c7280;
            box-shadow:inset 0 0 10px rgba(0,0,0,.42);
            font-size:clamp(.72rem,2.4vw,.86rem); line-height:1.2; font-weight:850;
            transition:transform .2s, box-shadow .2s, background .2s;
        }
        .dept-num{
            flex:0 0 22px; width:22px; height:22px; border-radius:50%; display:grid; place-items:center;
            background:#263840; border:1px solid #7a8f98; color:#cfe0e8; font-weight:1000; font-size:.72rem;
        }
        .sim.picking .dept-btn.target,.dept-btn.target.lit{
            color:#06131a; background:linear-gradient(#b8f5ff,#4ed7f4); border-color:#d7fbff;
            box-shadow:0 0 24px rgba(71,220,251,.7), inset 0 0 12px rgba(255,255,255,.7);
            transform:translateY(1px);
        }
        .sim.picking .dept-btn.target .dept-num,.dept-btn.target.lit .dept-num{background:#fff; border-color:#fff; color:#0a5279; box-shadow:0 0 12px #fff;}

        /* Kabindeki yolcular — yolculuk boyunca kaybolmaz */
        .pax{
            position:absolute; z-index:15; bottom:3%;
            width:clamp(74px,13vw,120px); height:clamp(210px,34vw,360px);
            transform-origin:bottom; animation:breathe 3.4s ease-in-out infinite;
            transition:transform 2.2s cubic-bezier(.4,.6,.2,1), opacity 1.4s ease;
        }
        .pax.p1{left:5%; transform:scale(.9);}
        .pax.p2{left:16%; transform:scale(.8); animation-delay:-1.1s;}
        .pax.p3{right:6%; transform:scale(.86); animation-delay:-1.9s;}
        @keyframes breathe{0%,100%{margin-bottom:0;}50%{margin-bottom:2px;}}
        /* ofise geçince yolcular yürüyerek çıkar (aniden yok olmaz) */
        .sim.exiting .pax.p1{transform:scale(1.3) translate(-40%,40%); opacity:0;}
        .sim.exiting .pax.p2{transform:scale(1.4) translate(-30%,50%); opacity:0;}
        .sim.exiting .pax.p3{transform:scale(1.3) translate(40%,45%); opacity:0;}

        .fig-head{position:absolute; left:50%; top:0; width:34%; aspect-ratio:.82; transform:translateX(-50%); border-radius:46% 46% 42% 42%; background:linear-gradient(90deg,#a96d46,#e2aa7e 52%,#b97951); box-shadow:inset 6px 0 11px rgba(76,34,17,.16);}
        .fig-hair{position:absolute; left:31%; top:-1%; width:38%; height:12%; border-radius:50% 50% 24% 24%; background:#241b19;}
        .fig-torso{position:absolute; left:22%; right:22%; top:20%; bottom:24%; border-radius:16px 16px 8px 8px; background:linear-gradient(90deg,#10263a,#274864 50%,#0e2234);}
        .pax.p2 .fig-torso{background:linear-gradient(90deg,#343b46,#69717d 50%,#2d343e);}
        .pax.p3 .fig-torso{background:linear-gradient(90deg,#3c2b42,#6f4d76 50%,#322337);}
        .fig-arm{position:absolute; top:24%; width:13%; height:44%; border-radius:20px; background:#17334b;}
        .fig-arm.left{left:14%; transform:rotate(6deg);} .fig-arm.right{right:14%; transform:rotate(-6deg);}
        .pax.p2 .fig-arm{background:#454e5a;} .pax.p3 .fig-arm{background:#55385a;}
        .fig-leg{position:absolute; bottom:0; width:19%; height:32%; border-radius:8px 8px 14px 14px; background:#111a23;}
        .fig-leg.left{left:28%;} .fig-leg.right{right:28%;}
        .fig-prop{position:absolute; z-index:4; left:26%; top:42%; width:48%; height:24%; border-radius:5px; background:linear-gradient(#697985,#26343d); border:2px solid #899aa4; box-shadow:0 8px 15px rgba(0,0,0,.3);}
        .pax.p2 .fig-prop{left:22%; width:54%; background:#0c151b; border:3px solid #738791;}
        .pax.p3 .fig-prop{right:2%; left:auto; width:36%; height:26%; background:linear-gradient(#7a4e2f,#3e2819); border:2px solid #9a6d48;}

        /* ---- OFİS (kabin kapısı açılınca görünür, sonra kamera içine girer) ---- */
        .stage.office{
            background:
                linear-gradient(180deg,#eaf4f8 0 18%,#cfe1e9 18% 74%,#9db2bd 74% 100%);
            transform:scale(.9);
        }
        .sim.exiting .stage.office,.sim.in-office .stage.office{transform:scale(1);}
        .sim.reading .stage.office{transform:scale(1.06);}
        .office-ceil{position:absolute; left:-8%; right:-8%; top:-12%; height:32%; background:linear-gradient(#f4fafc,#d3dfe4); transform:perspective(680px) rotateX(-52deg);}
        .office-floor{position:absolute; left:-14%; right:-14%; bottom:-20%; height:52%;
            background:
                repeating-linear-gradient(90deg, rgba(27,75,94,.12) 0 1px, transparent 1px 8%),
                linear-gradient(#9cb9c5,#e2edf1);
            transform:perspective(560px) rotateX(63deg); transform-origin:top;}
        .office-wall{position:absolute; top:20%; bottom:16%; width:22%; border:2px solid rgba(51,108,133,.3); background:linear-gradient(135deg,rgba(255,255,255,.4),rgba(111,187,218,.1));}
        .office-wall.left{left:2%;} .office-wall.right{right:2%;}
        .office-sign{
            position:absolute; left:50%; top:10.5%; transform:translateX(-50%); white-space:nowrap;
            padding:10px 22px; border-radius:12px;
            color:#eaf7ff; background:linear-gradient(#1c4a66,#0f3550);
            border:1px solid rgba(160,220,245,.5); box-shadow:0 10px 26px rgba(9,40,58,.35), inset 0 1px 0 rgba(255,255,255,.22);
            font-size:clamp(.95rem,3.4vw,1.7rem); font-weight:1000; letter-spacing:.12em;
        }
        .desk{position:absolute; bottom:20%; width:22%; height:11%; border-radius:5px 5px 0 0; background:linear-gradient(#b37d50,#6e472d); box-shadow:0 10px 16px rgba(0,0,0,.18);}
        .desk.d1{left:9%;} .desk.d2{right:9%;}
        .desk::after{content:""; position:absolute; left:30%; bottom:100%; width:40%; height:88%; border-radius:4px; background:#16232b; border:3px solid #6b7e88;}
        .worker{position:absolute; z-index:3; bottom:17%; width:clamp(34px,5vw,54px);}
        .worker .wk-head{display:block; width:60%; margin:0 auto; aspect-ratio:1; border-radius:50%; background:#c98f66;}
        .worker .wk-body{display:block; width:100%; height:clamp(70px,10vw,110px); border-radius:12px 12px 4px 4px; margin-top:4px; background:linear-gradient(#2a4a63,#16293a);}
        .worker.wk1{left:14%; animation:sway 5s ease-in-out infinite;}
        .worker.wk2{right:14%; animation:sway 6s ease-in-out infinite -2s;}
        .worker.wk3{left:44%; transform:scale(.8); animation:sway 7s ease-in-out infinite -3s;}
        @keyframes sway{0%,100%{transform:translateX(0);}50%{transform:translateX(4px);}}

        /* SEKRETER — uzaktan yaklaşır, belgeyi uzatır */
        .secretary{
            position:absolute; z-index:8; left:50%; bottom:15%;
            width:clamp(96px,14vw,150px); height:clamp(250px,42vw,400px);
            transform:translate(-50%,26%) scale(.32); opacity:0;
            transition:transform 3s cubic-bezier(.18,.74,.16,1), opacity .8s ease;
        }
        .sim.sec-in .secretary{transform:translate(-50%,0) scale(1); opacity:1;}
        .secretary .fig-head{top:1%; background:linear-gradient(90deg,#bd8058,#efbc8f 52%,#c78b63);}
        .secretary .fig-hair{left:26%; top:-1%; width:48%; height:18%; border-radius:50% 50% 28% 28%; background:#2b1c1a;}
        .secretary .fig-torso{left:19%; right:19%; top:22%; bottom:24%; background:linear-gradient(90deg,transparent 47%,rgba(255,255,255,.42) 48% 52%,transparent 53%),linear-gradient(#193450,#0c1e31);}
        .secretary .fig-arm{background:#16314b; transition:transform 1.3s ease;}
        .secretary .fig-arm.right{transform-origin:top;}
        .sim.sec-offer .secretary .fig-arm.right{transform:rotate(52deg) translate(-30%,-6%);}
        .secretary .fig-leg{background:#101a22;}

        /* BELGE — sekreterin elinde başlar, sonra okumak için kadraja gelir. Işınlanmaz. */
        .doc{
            position:absolute; z-index:10; right:-6%; top:50%;
            width:min(46vw,320px); aspect-ratio:1/1.32;
            padding:6% 6% 5%; border-radius:4px;
            background:linear-gradient(180deg,#fffdf3,#f6efdc); color:#173044;
            border:1px solid #d8c692;
            box-shadow:0 14px 34px rgba(0,0,0,.4);
            transform:rotate(-8deg) scale(.5); transform-origin:right center; opacity:0;
            transition:opacity .5s ease .3s;
            overflow:hidden;
        }
        .sim.sec-offer .doc{opacity:1;}
        /* okuma anı: belge büyür ve önümüze gelir */
        .sim.reading .doc{
            position:fixed; left:50%; top:47%; right:auto; z-index:120;
            width:min(560px,90vw); aspect-ratio:1/1.3;
            transform:translate(-50%,-50%) rotate(-1deg) scale(1); opacity:1;
            box-shadow:0 42px 100px rgba(0,0,0,.55);
            transition:all 1.2s cubic-bezier(.2,.83,.2,1);
        }
        .doc::before{content:""; position:absolute; inset:5%; border:2px solid rgba(31,87,112,.16);}
        .doc-inner{position:relative; z-index:2; height:100%; display:flex; flex-direction:column;}
        .doc-logo{display:flex; align-items:center; gap:8px; color:#245874; font-weight:1000; letter-spacing:.06em; font-size:clamp(.55rem,2vw,.8rem);}
        .doc-logo-mark{width:clamp(24px,6vw,34px); aspect-ratio:1; display:grid; place-items:center; border-radius:8px; color:#fff; background:linear-gradient(145deg,#168bc7,#0b456b); font-weight:1000;}
        .doc-ref{margin-top:4px; color:#6b7f8c; font-size:clamp(.5rem,1.7vw,.66rem); font-weight:700;}
        .doc-title{margin-top:6%; font-family:Georgia,"Times New Roman",serif; font-size:clamp(1rem,4.4vw,1.7rem); color:#19394e; font-weight:800;}
        .doc-caption{color:#4f6c7c; font-weight:750; font-size:clamp(.6rem,2.2vw,.85rem);}
        .doc-balance{margin:3% 0; font-size:clamp(1.8rem,10vw,3.6rem); line-height:1; color:#087eb7; font-weight:1000; letter-spacing:-.04em;}
        .doc-balance small{font-size:.3em; color:#4f6c7c; font-weight:800;}
        .doc-sign{margin-top:auto; text-align:right; color:#315b70;}
        .doc-sign .sg-role{display:block; font-size:clamp(.55rem,2vw,.78rem); font-weight:750; color:#4f6c7c;}
        .doc-sign .sg-name{display:block; margin-top:2px; font-family:Inter,"Segoe UI",sans-serif; font-weight:900; font-size:clamp(.8rem,3vw,1.05rem); color:#1b3a4c; letter-spacing:.01em;}
        .doc-stamp{
            display:inline-block; margin-top:6px; padding:5px 12px; border-radius:6px;
            border:2px solid #1f7a5a; color:#1f7a5a; font-weight:1000; font-size:clamp(.5rem,1.8vw,.68rem);
            letter-spacing:.14em; transform:rotate(-6deg); background:rgba(31,122,90,.06);
        }
        .doc-foot{margin-top:5%; text-align:center; color:#7c8f9a; font-size:clamp(.44rem,1.5vw,.58rem); font-weight:700;}

        /* Kamera doğal hareket: sadece yukarı-aşağı yürüyüş; ASLA sağa-sola */
        .sim.walking .sim-world{animation:walkBob .72s ease-in-out infinite;}
        @keyframes walkBob{0%,100%{transform:translateY(0);}50%{transform:translateY(.9%);}}
        .sim.riding .sim-world{animation:cabinShake .5s ease-in-out infinite;}
        @keyframes cabinShake{0%,100%{transform:translateY(0);}25%{transform:translateY(-.25%);}75%{transform:translateY(.2%);}}
        .sim.reading .sim-world{animation:none; transform:translateY(3%);}

        /* Yolculukta geçen kat ışıkları (hız hissi) */
        .ride-fx{
            position:absolute; inset:0; z-index:18; opacity:0; pointer-events:none;
            background:linear-gradient(180deg,transparent 0 42%,rgba(150,228,255,.09) 50%,transparent 58%);
            background-size:100% 240%;
        }
        .sim.riding .ride-fx{opacity:1; animation:floorSweep 1s linear infinite;}
        .sim.riding.fast .ride-fx{animation-duration:.42s;}
        @keyframes floorSweep{from{background-position:0 -140%;}to{background-position:0 140%;}}
        .sim.riding.fast .sim-world{animation-duration:.34s;}

        /* Varışta gösterge 'ding' parlar */
        .floor-ind.ding{animation:dingPulse 1s ease;}
        @keyframes dingPulse{
            0%{box-shadow:inset 0 0 16px rgba(89,221,255,.12),0 0 22px rgba(59,207,247,.16);}
            30%{box-shadow:inset 0 0 22px rgba(140,236,255,.35),0 0 52px rgba(89,225,255,.85); transform:translateX(-50%) scale(1.05);}
            100%{transform:translateX(-50%) scale(1);}
        }

        /* Birinci şahıs el */
        .fp-hand{position:absolute; z-index:40; left:0; right:0; bottom:-4%; height:44%; pointer-events:none;}
        .fp-arm{
            position:absolute; bottom:-30%; right:6%;
            width:clamp(120px,24vw,240px); height:clamp(260px,44vw,460px);
            transform-origin:bottom; transform:translateY(120%); opacity:0;
            transition:transform 1.1s cubic-bezier(.2,.8,.2,1), opacity .5s ease;
        }
        .sim.picking .fp-arm{transform:translateY(28%) rotate(-6deg); opacity:1;}
        /* okuma anında el belgeyi ALTINDAN tutar; imzayı asla kapatmaz */
        .sim.reading .fp-arm{transform:translateY(58%) rotate(-3deg); opacity:1;}
        .fp-sleeve{position:absolute; left:20%; right:20%; bottom:0; height:66%; border-radius:38% 38% 16% 16%; background:linear-gradient(90deg,#0a1a2a,#1d3e5d 48%,#0b2034);}
        .fp-palm{position:absolute; left:24%; top:2%; width:52%; height:36%; border-radius:45% 45% 40% 40%; background:linear-gradient(90deg,#ad704b,#edba8d 50%,#c8865c);}
        .fp-arm i{position:absolute; top:-3%; width:13%; height:30%; border-radius:45% 45% 35% 35%; background:linear-gradient(90deg,#b67852,#edba8d 52%,#c88961);}
        .fp-arm i:nth-child(3){left:24%; transform:rotate(-8deg); height:26%;}
        .fp-arm i:nth-child(4){left:36%; height:31%;}
        .fp-arm i:nth-child(5){left:49%; height:29%;}
        .fp-arm i:nth-child(6){left:62%; transform:rotate(8deg); height:24%;}

        .sim-status{
            position:absolute; z-index:50; left:50%; bottom:max(16px,env(safe-area-inset-bottom)); transform:translateX(-50%);
            max-width:calc(100vw - 30px); padding:9px 16px; border-radius:999px;
            color:#d9f5ff; background:rgba(3,16,26,.78); border:1px solid rgba(141,222,255,.24);
            backdrop-filter:blur(12px); font-size:clamp(.66rem,2.2vw,.8rem);
            font-weight:900; letter-spacing:.08em; text-align:center;
        }

        /* ---------- DASHBOARD ---------- */
        .dashboard{width:min(1120px,100%); margin:0 auto; animation:dashboardIn .5s cubic-bezier(.2,.8,.2,1) both;}
        @keyframes dashboardIn{from{opacity:0; transform:translateY(18px);}to{opacity:1; transform:none;}}
        .dashboard-header{display:flex; justify-content:space-between; align-items:center; gap:16px; margin-bottom:18px;}
        .person{display:flex; align-items:center; gap:14px; min-width:0;}
        .avatar{width:58px; height:58px; flex:0 0 58px; display:grid; place-items:center; border-radius:18px; background:linear-gradient(145deg,var(--primary),var(--primary-2)); color:#032039; font-size:1.15rem; font-weight:1000; box-shadow:0 12px 26px rgba(28,164,220,.25);}
        .person h2{margin:0; font-size:clamp(1.25rem,4vw,1.85rem); white-space:nowrap; overflow:hidden; text-overflow:ellipsis;}
        .person p{margin:4px 0 0; color:var(--muted);}
        .header-buttons{display:flex; gap:8px;}
        .small-btn{border:1px solid var(--line); background:var(--panel); border-radius:13px; padding:11px 13px; cursor:pointer; font-weight:800;}
        .dashboard-grid{display:grid; grid-template-columns:minmax(0,1.35fr) minmax(280px,.65fr); gap:18px;}
        .main-column,.side-column{display:grid; gap:18px; align-content:start;}
        .hero-card{padding:clamp(22px,5vw,36px); position:relative; overflow:hidden;}
        .hero-card::after{content:""; position:absolute; width:260px; height:260px; right:-90px; top:-90px; border-radius:50%; background:radial-gradient(circle,rgba(85,217,255,.24),transparent 67%); pointer-events:none;}
        .hero-content{display:grid; grid-template-columns:1fr auto; gap:22px; align-items:center; position:relative; z-index:2;}
        .eyebrow{color:var(--primary); text-transform:uppercase; letter-spacing:.13em; font-size:.76rem; font-weight:1000;}
        .big-balance{margin:10px 0 5px; font-size:clamp(3.2rem,12vw,6.3rem); line-height:.92; letter-spacing:-.075em; font-weight:1000;}
        .big-balance small{font-size:.2em; letter-spacing:.02em; color:var(--muted); margin-left:8px;}
        .update-info{color:var(--muted); font-size:.85rem; margin-top:13px;}
        .progress-ring{--progress:75deg; width:clamp(112px,20vw,160px); aspect-ratio:1; border-radius:50%; display:grid; place-items:center; background:conic-gradient(var(--primary) var(--progress), rgba(127,202,230,.13) 0); position:relative; box-shadow:0 0 36px rgba(62,205,255,.12);}
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
        .id-name{font-size:1.25rem; font-weight:1000;}
        .id-role{color:#b9dded; font-size:.78rem; margin:5px 0 20px;}
        .id-number{color:#86dfff; font:800 .78rem ui-monospace,monospace;}
        .qr-box{width:92px; height:92px; padding:7px; border-radius:10px; background:white; display:grid; place-items:center;}
        #qrCanvas{width:78px; height:78px;}
        .modal-backdrop{position:fixed; z-index:220; inset:0; display:grid; place-items:center; padding:20px; background:rgba(1,10,18,.68); backdrop-filter:blur(9px); opacity:0; pointer-events:none; transition:opacity .25s ease;}
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

        @media (max-width:820px){
            .dashboard-grid{grid-template-columns:1fr;}
            .dashboard-header{align-items:flex-start;}
            .header-buttons{flex-direction:column;}
            .hero-content{grid-template-columns:1fr;}
            .hero-content .progress-ring{position:absolute; right:0; top:0; opacity:.5; transform:scale(.72); transform-origin:top right;}
            .dept-panel{width:min(46vw,220px);}
            .cabin-doorway{left:18%; right:18%;}
        }
        @media (max-width:520px){
            .brand-mini span:last-child{display:none;}
            .install-btn span{display:none;}
            .page-shell{padding-left:12px; padding-right:12px;}
            .dashboard-header{display:grid; grid-template-columns:1fr auto;}
            .avatar{width:50px; height:50px; flex-basis:50px;}
            .person p{font-size:.78rem;}
            .form-grid{grid-template-columns:1fr;}
            .id-main{grid-template-columns:1fr 84px;}
            .qr-box{width:84px; height:84px;} #qrCanvas{width:70px; height:70px;}
            .office-sign{font-size:.82rem; padding:8px 14px; letter-spacing:.08em;}
        }
        @media (prefers-reduced-motion:reduce){
            *,*::before,*::after{animation-duration:.01ms !important; animation-iteration-count:1 !important; transition-duration:.2s !important;}
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
                </div>
                <button class="primary-btn" id="loginBtn" type="submit">
                    <span id="loginBtnText" data-i18n="login">Giriş Yap</span>
                    <span class="spinner"></span>
                </button>
            </form>
            <div class="login-links"><button class="text-btn" id="forgotBtn" type="button" data-i18n="forgot">Şifremi Unuttum</button></div>
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
                            <div class="progress-ring" id="progressRing"><span><b id="ringValue">0</b><span data-i18n="leaveLevel">Kat</span></span></div>
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

    <!-- ============ KESİNTİSİZ ASANSÖR SİMÜLASYONU ============ -->
    <section class="sim" id="sim" hidden aria-label="Asansör animasyonu">
        <div class="sim-world" id="simWorld">

            <!-- OFİS (en altta, daima opak) -->
            <div class="stage office">
                <div class="office-ceil"></div>
                <div class="office-floor"></div>
                <div class="office-wall left"></div>
                <div class="office-wall right"></div>
                <div class="office-sign">PERSONEL VE ÇALIŞMA İLİŞKİLERİ</div>
                <div class="desk d1"></div><div class="desk d2"></div>
                <div class="worker wk1"><span class="wk-head"></span><span class="wk-body"></span></div>
                <div class="worker wk2"><span class="wk-head"></span><span class="wk-body"></span></div>
                <div class="worker wk3"><span class="wk-head"></span><span class="wk-body"></span></div>

                <div class="secretary" id="secretary">
                    <div class="fig-hair"></div><div class="fig-head"></div>
                    <div class="fig-torso"></div>
                    <div class="fig-arm left"></div><div class="fig-arm right"></div>
                    <div class="fig-leg left"></div><div class="fig-leg right"></div>
                    <div class="doc" id="doc">
                        <div class="doc-inner">
                            <div class="doc-logo"><span class="doc-logo-mark">İ</span><span>PERSONEL İZİN BİLDİRİMİ</span></div>
                            <div class="doc-ref" id="docRef">Ref: — · —</div>
                            <div class="doc-title" id="docTitle">Sn. Burhan Biliktü</div>
                            <div class="doc-caption" data-i18n="noteCaption">Kalan İzin Hakkınız</div>
                            <div class="doc-balance"><span id="docBalance">29,5</span> <small data-i18n="dayLower">gün</small></div>
                            <div class="doc-sign">
                                <span class="sg-role" data-i18n="hrChief">Personel ve Çalışma İlişkileri Şefi</span>
                                <span class="sg-name">İlker Sezgin</span>
                                <span class="doc-stamp" data-i18n="digitalApproval">DİJİTAL ONAY</span>
                            </div>
                            <div class="doc-foot" data-i18n="docFoot">Bu belge dijital sistem üzerinden oluşturulmuştur.</div>
                        </div>
                    </div>
                </div>
            </div>

            <!-- KABİN (ofisin üstünde, daima opak; okuma anında söner) -->
            <div class="stage cabin">
                <div class="cabin-ceil"></div>
                <div class="cabin-light"></div>
                <div class="cabin-floor"></div>
                <div class="wall left"></div>
                <div class="wall right"></div>

                <div class="pax p1"><div class="fig-hair"></div><div class="fig-head"></div><div class="fig-torso"></div><div class="fig-arm left"></div><div class="fig-arm right"></div><div class="fig-leg left"></div><div class="fig-leg right"></div><div class="fig-prop"></div></div>
                <div class="pax p2"><div class="fig-hair"></div><div class="fig-head"></div><div class="fig-torso"></div><div class="fig-arm left"></div><div class="fig-arm right"></div><div class="fig-leg left"></div><div class="fig-leg right"></div><div class="fig-prop"></div></div>
                <div class="pax p3"><div class="fig-hair"></div><div class="fig-head"></div><div class="fig-torso"></div><div class="fig-arm left"></div><div class="fig-arm right"></div><div class="fig-leg left"></div><div class="fig-leg right"></div><div class="fig-prop"></div></div>

                <div class="floor-ind"><b id="floorValue">0</b><small data-i18n="leaveLevelUpper">KAT</small></div>
                <div class="half-tag" id="halfLevel" data-i18n="halfDay">ARA KAT · ½ GÜN</div>

                <div class="doorway cabin-doorway">
                    <div class="door l"></div>
                    <div class="door r"></div>
                </div>
                <div class="ride-fx"></div>

                <div class="dept-panel">
                    <div class="dept-title" data-i18n="deptTitle">DEPARTMAN</div>
                    <button class="dept-btn" type="button"><span class="dept-num">1</span><span>İdari İşler</span></button>
                    <button class="dept-btn" type="button"><span class="dept-num">2</span><span>İSG</span></button>
                    <button class="dept-btn" type="button"><span class="dept-num">3</span><span>Yönetim</span></button>
                    <button class="dept-btn" type="button"><span class="dept-num">4</span><span>Kalite</span></button>
                    <button class="dept-btn target" type="button"><span class="dept-num">5</span><span>Personel ve Çalışma İlişkileri</span></button>
                </div>
            </div>

            <!-- LOBİ (en üstte, girişte söner) -->
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

        <div class="fp-hand" id="fpHand">
            <div class="fp-arm"><div class="fp-sleeve"></div><div class="fp-palm"></div><i></i><i></i><i></i><i></i></div>
        </div>

        <div class="sim-status" id="elevatorStatus">ASANSÖR</div>
    </section>

    <!-- Şifre unutma -->
    <div class="modal-backdrop" id="forgotModal" aria-hidden="true">
        <div class="modal glass">
            <div class="modal-head"><h3 data-i18n="forgotTitle">Şifre Talebi</h3><button class="close-btn" type="button" data-close="forgotModal">×</button></div>
            <div class="field"><label for="forgotIdentity" data-i18n="nameOrUsername">Ad Soyad veya Kullanıcı Adı</label><input id="forgotIdentity" data-i18n-placeholder="identityPlaceholder" placeholder="Bilginizi yazın"></div>
            <button class="primary-btn" id="forgotWhatsappBtn" type="button" data-i18n="sendWhatsapp">WhatsApp'tan Gönder</button>
        </div>
    </div>

    <!-- İzin talebi -->
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
                replay:"Asansörü Tekrar İzle", logout:"Çıkış",
                remainingLeave:"Kalan İzin Hakkınız", day:"GÜN", dayLower:"gün",
                updated:"Son güncelleme:", leaveLevel:"Kat",
                sundayLeave:"Pazar İzinleri", officialHoliday:"Resmî Tatil",
                quickActions:"Hızlı İşlemler", objectLeave:"İzin Gününe İtiraz Et",
                requestLeave:"İzin Talebi Oluştur", addHome:"Ana Ekrana Uygulama Olarak Ekle",
                nextHoliday:"Yaklaşan Resmî Tatil", security:"Güvenlik Bilgisi",
                lastLogin:"Son girişiniz", session:"Oturum", active:"Aktif",
                digitalId:"Dijital Personel Kimliği", leaveLevelUpper:"KAT",
                halfDay:"ARA KAT · ½ GÜN", noteCaption:"Kalan İzin Hakkınız",
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
                stArrived:"KATA ULAŞILDI", stOffice:"KATA ÇIKILIYOR", stSecretary:"SEKRETER YAKLAŞIYOR", stDelivery:"BELGE TESLİM EDİLİYOR",
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
                replay:"Replay Elevator", logout:"Sign Out",
                remainingLeave:"Remaining Leave Balance", day:"DAYS", dayLower:"days",
                updated:"Last update:", leaveLevel:"Floor",
                sundayLeave:"Sunday Leave", officialHoliday:"Public Holiday",
                quickActions:"Quick Actions", objectLeave:"Object to Leave Balance",
                requestLeave:"Create Leave Request", addHome:"Add App to Home Screen",
                nextHoliday:"Next Public Holiday", security:"Security Information",
                lastLogin:"Your last sign-in", session:"Session", active:"Active",
                digitalId:"Digital Employee ID", leaveLevelUpper:"FLOOR",
                halfDay:"HALF FLOOR · ½ DAY", noteCaption:"Your Remaining Leave Balance",
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
                stArrived:"FLOOR REACHED", stOffice:"ENTERING THE FLOOR", stSecretary:"SECRETARY IS APPROACHING", stDelivery:"DELIVERING DOCUMENT",
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
                if(typeof v === "string") el.textContent = v;
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

        async function login(event){
            event.preventDefault();
            const username = document.getElementById("username").value.trim();
            const password = document.getElementById("password").value.trim();
            if(!username || !password){ showToast(t("fillFields"), "error"); return; }
            const button = document.getElementById("loginBtn");
            button.disabled = true; button.classList.add("loading");
            try{
                const response = await fetch("/login", {method:"POST", headers:{"Content-Type":"application/json"}, body:JSON.stringify({username, password})});
                const result = await response.json();
                if(!response.ok || result.status !== "success"){
                    if(response.status === 429) showToast(result.message || t("blocked"), "error");
                    else showToast(result.message || t("loginError"), "error");
                    return;
                }
                currentUser = result.data;
                const loginKey = `izin-last-login-${currentUser.username}`;
                currentUser.previous_login = localStorage.getItem(loginKey);
                localStorage.setItem(loginKey, new Date().toISOString());
                document.getElementById("password").value = "";
                await playElevatorAnimation(currentUser);
            }catch(error){
                console.error(error); showToast(t("serverError"), "error");
            }finally{
                button.disabled = false; button.classList.remove("loading");
            }
        }

        /* ---- Animasyon yardımcıları (iptal edilebilir) ---- */
        function delay(ms, runId){
            return new Promise(resolve=>{
                const timer = setTimeout(()=>resolve(true), ms);
                const checker = setInterval(()=>{
                    if(runId !== animationRunId){ clearTimeout(timer); clearInterval(checker); resolve(false); }
                }, 80);
                setTimeout(()=>clearInterval(checker), ms + 120);
            });
        }
        function animateFloor(target, duration, runId){
            return new Promise(resolve=>{
                const start = performance.now();
                const floor = document.getElementById("floorValue");
                const sim = document.getElementById("sim");
                // Gerçek asansör ivmesi: yavaş kalkış -> hızlanma -> sonda kısa fren
                function ease(p){
                    if(p < .82){ const q = p / .82; return .9 * q * q * q; }
                    const q = (p - .82) / .18; return .9 + .1 * (1 - Math.pow(1 - q, 2));
                }
                function frame(now){
                    if(runId !== animationRunId){ sim.classList.remove("fast"); resolve(false); return; }
                    const progress = Math.min((now - start) / duration, 1);
                    sim.classList.toggle("fast", progress > .3 && progress < .88);
                    floor.textContent = formatNumber(Math.floor(target * ease(progress)));
                    if(progress < 1) requestAnimationFrame(frame);
                    else { sim.classList.remove("fast"); floor.textContent = formatNumber(target); resolve(true); }
                }
                requestAnimationFrame(frame);
            });
        }
        function buzz(pattern){ try{ if(navigator.vibrate) navigator.vibrate(pattern); }catch(e){} }
        function setStatus(key){ document.getElementById("elevatorStatus").textContent = t(key); }
        function simClass(add, remove){
            const sim = document.getElementById("sim");
            (remove||[]).forEach(c=>sim.classList.remove(c));
            (add||[]).forEach(c=>sim.classList.add(c));
        }
        function resetSim(){
            const sim = document.getElementById("sim");
            sim.className = "sim";
            document.getElementById("floorValue").textContent = "0";
            document.getElementById("lobbyDisplay").textContent = "G";
            document.getElementById("halfLevel").classList.remove("show");
            document.querySelector(".dept-btn.target").classList.remove("lit");
            document.querySelector(".floor-ind").classList.remove("ding");
        }

        async function playElevatorAnimation(user){
            const runId = ++animationRunId;
            resetSim();
            const sim = document.getElementById("sim");
            const loginPanel = document.getElementById("loginPanel");
            const dashboard = document.getElementById("dashboard");

            // Belge içeriğini hazırla
            document.getElementById("docTitle").textContent = t("noteTitle")(user.name);
            document.getElementById("docBalance").textContent = formatNumber(user.remaining_leave);
            const now = new Date();
            const ref = `RET-${now.getFullYear()}-${String(user.username).replace(/\D/g,"").padStart(4,"0").slice(-4)}`;
            document.getElementById("docRef").textContent =
                `Ref: ${ref} · ${new Intl.DateTimeFormat(currentLanguage==="tr"?"tr-TR":"en-GB",{day:"2-digit",month:"2-digit",year:"numeric"}).format(now)}`;

            loginPanel.hidden = true; dashboard.hidden = true; sim.hidden = false;

            // Asansör hedefi: 67. kat (Personel ve Çalışma İlişkileri)
            const target = 67;

            // 1) Lobi: asansör çağrılır, kapı açılır
            setStatus("stCalling");
            if(!(await delay(1000, runId))) return;
            simClass(["lobby-open"]); setStatus("stDoorsOpen");
            if(!(await delay(2400, runId))) return;

            // 2) İçeri girilir (lobi söner, kabin ortaya çıkar — siyahlık yok)
            simClass(["stepped"]); setStatus("stEntering");
            if(!(await delay(1900, runId))) return;

            // 3) Departman paneline basılır (el gelir, Personel butonu yanar ve SEÇİLİ KALIR)
            simClass(["picking"]); setStatus("stPicking"); buzz(25);
            if(!(await delay(2400, runId))) return;
            document.querySelector(".dept-btn.target").classList.add("lit");
            simClass([], ["picking"]);
            if(!(await delay(700, runId))) return;

            // 4) Gerçek yolculuk: yavaş kalkış -> hızlanma -> fren; kat 67'ye çıkar
            simClass(["riding"]); setStatus("stRising");
            const floorDone = await animateFloor(target, 13000, runId);
            if(!floorDone || runId !== animationRunId) return;
            simClass([], ["riding"]);
            document.getElementById("floorValue").textContent = formatNumber(target);
            document.querySelector(".floor-ind").classList.add("ding");
            buzz([40, 60, 40]);
            setStatus("stArrived");
            if(!(await delay(1200, runId))) return;

            // 5) Varış: kabin kapıları açılır -> ofis görünür (alttaki opak katman)
            simClass(["arrived"]); setStatus("stDoorsOpen");
            if(!(await delay(2000, runId))) return;

            // 6) Ofise çıkış: kamera içeri girer, yolcular yürüyerek çıkar
            simClass(["exiting","walking","in-office"]); setStatus("stOffice");
            if(!(await delay(2800, runId))) return;

            // 7) Sekreter uzaktan yürüyerek yaklaşır
            simClass(["sec-in"]); setStatus("stSecretary");
            if(!(await delay(3000, runId))) return;

            // 8) Belgeyi uzatır, el belgeyi alır, kamera okumak için eğilir
            simClass(["sec-offer"]); setStatus("stDelivery");
            if(!(await delay(1400, runId))) return;
            simClass(["reading"], ["walking"]);
            if(!(await delay(7800, runId))) return;

            finishAnimation(runId);
        }

        function finishAnimation(runId=null){
            if(runId !== null && runId !== animationRunId) return;
            animationRunId++;
            document.getElementById("sim").hidden = true;
            resetSim();
            renderUser(currentUser, true);
            document.getElementById("dashboard").hidden = false;
            window.scrollTo({top:0, behavior:"smooth"});
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
            document.getElementById("objectionBtn").href = `https://wa.me/${WHATSAPP_NUMBER}?text=${encodeURIComponent(t("objectionMessage")(user.name))}`;
            drawQr({portal:"Personel İzin Portalı", id:user.username, name:user.name, role:user.role});
            calculateNextHoliday();
        }

        function drawQr(data){
            const canvas = document.getElementById("qrCanvas");
            const value = JSON.stringify(data);
            if(window.QRious){
                new QRious({element:canvas, value, size:156, level:"M", background:"white", foreground:"#07263c"});
                return;
            }
            const ctx = canvas.getContext("2d");
            const size = canvas.width, cells = 21, cell = size / cells;
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
        document.getElementById("forgotBtn").addEventListener("click", ()=>openModal("forgotModal"));
        document.getElementById("leaveRequestBtn").addEventListener("click", ()=>openModal("leaveModal"));
        document.getElementById("installBtn").addEventListener("click", installPwa);
        document.getElementById("installActionBtn").addEventListener("click", installPwa);
        document.getElementById("replayBtn").addEventListener("click", ()=>playElevatorAnimation(currentUser));
        document.getElementById("logoutBtn").addEventListener("click", ()=>{
            animationRunId++; currentUser = null;
            document.getElementById("dashboard").hidden = true;
            document.getElementById("loginPanel").hidden = false;
            document.getElementById("username").focus();
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

        autoTheme();
        applyLanguage();
        setupPwa();
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
