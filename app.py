from flask import Flask, request, jsonify, render_template_string
import pandas as pd

app = Flask(__name__)

# BÜTÜN GÖRSELLİK VE HTML KODU BURAYA GÖMÜLDÜ (TEK DOSYA MANTIĞI)
HTML_SAYFASI = """
<!DOCTYPE html>
<html lang="tr">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Personel İzin Sistemi</title>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <style>
        body {
            margin: 0; padding: 0; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%);
            min-height: 100vh; display: flex; justify-content: center; align-items: center;
            color: #fff; overflow: hidden;
        }
        .glass-panel {
            background: rgba(255, 255, 255, 0.1); backdrop-filter: blur(15px); -webkit-backdrop-filter: blur(15px);
            border-radius: 20px; border: 1px solid rgba(255, 255, 255, 0.2); padding: 40px; 
            width: 90%; max-width: 400px; box-shadow: 0 25px 45px rgba(0,0,0,0.2); transition: all 0.5s ease-in-out; position: relative;
        }
        h2 { text-align: center; margin-bottom: 30px; font-weight: 300; letter-spacing: 2px; }
        .input-group { margin-bottom: 20px; }
        .input-group input {
            width: 100%; padding: 15px; border: none; border-radius: 10px; background: rgba(255, 255, 255, 0.2); 
            color: #fff; font-size: 16px; outline: none; box-sizing: border-box; transition: 0.3s;
        }
        .input-group input::placeholder { color: rgba(255,255,255,0.7); }
        .input-group input:focus { background: rgba(255, 255, 255, 0.3); border-left: 4px solid #00f2fe; }
        button {
            width: 100%; padding: 15px; border: none; border-radius: 10px; background: linear-gradient(45deg, #4facfe 0%, #00f2fe 100%);
            color: white; font-size: 16px; font-weight: bold; cursor: pointer; transition: 0.3s; display: flex; justify-content: center; align-items: center;
        }
        button:hover { transform: translateY(-2px); box-shadow: 0 10px 20px rgba(0, 242, 254, 0.4); }
        .spinner {
            border: 3px solid rgba(255,255,255,0.3); border-top: 3px solid white; border-radius: 50%; width: 20px; height: 20px; 
            animation: spin 1s linear infinite; display: none; margin-left: 10px;
        }
        @keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }
        #toast {
            visibility: hidden; min-width: 250px; background-color: #ff4757; color: #fff; text-align: center; border-radius: 8px; 
            padding: 16px; position: fixed; z-index: 1; right: 30px; top: 30px; font-size: 15px; box-shadow: 0 10px 20px rgba(0,0,0,0.2);
            transform: translateX(100%); transition: all 0.5s cubic-bezier(0.68, -0.55, 0.265, 1.55);
        }
        #toast.show { visibility: visible; transform: translateX(0); }
        #dashboard { display: none; max-width: 600px; opacity: 0; }
        .profile-header { text-align: center; margin-bottom: 30px; }
        .badge { background: rgba(0, 242, 254, 0.2); color: #00f2fe; padding: 5px 15px; border-radius: 20px; font-size: 0.85rem; font-weight: bold; display: inline-block; margin-top: 10px; border: 1px solid #00f2fe; }
        .kalan-izin-container { text-align: center; background: rgba(0,0,0,0.2); padding: 20px; border-radius: 15px; margin-bottom: 30px; }
        .kalan-izin-title { font-size: 1.2rem; color: rgba(255,255,255,0.8); }
        .kalan-izin-sayi { font-size: 4rem; font-weight: bold; color: #4facfe; text-shadow: 0 0 20px rgba(79, 172, 254, 0.5); }
        .chart-container { width: 100%; max-height: 250px; display: flex; justify-content: center; }
    </style>
</head>
<body>
    <div id="toast">Hatalı bilgi girdiniz!</div>

    <div class="glass-panel" id="login-panel">
        <h2>SİSTEM GİRİŞİ</h2>
        <div class="input-group"><input type="text" id="username" placeholder="Kullanıcı Adı" autocomplete="off"></div>
        <div class="input-group"><input type="password" id="password" placeholder="Şifre"></div>
        <button onclick="girisYap()" id="login-btn">Giriş Yap <div class="spinner" id="spinner"></div></button>
    </div>

    <div class="glass-panel" id="dashboard">
        <div class="profile-header">
            <h2 id="isimSoyisim" style="margin-bottom: 5px;">İsim Soyisim</h2>
            <div class="badge" id="gorevBadge">Görev</div>
        </div>
        <div class="kalan-izin-container">
            <div class="kalan-izin-title">KALAN İZİN HAKKINIZ</div>
            <div class="kalan-izin-sayi" id="kalanIzinGosterge">0</div>
        </div>
        <div class="chart-container"><canvas id="izinGrafik"></canvas></div>
    </div>

    <script>
        function girisYap() {
            const user = document.getElementById('username').value;
            const pass = document.getElementById('password').value;
            const btn = document.getElementById('login-btn');
            const spinner = document.getElementById('spinner');

            if(!user || !pass) { showToast("Lütfen tüm alanları doldurun!"); return; }

            btn.style.opacity = "0.8"; spinner.style.display = "block";

            fetch('/login', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ username: user, password: pass })
            }).then(r => r.json()).then(data => {
                spinner.style.display = "none"; btn.style.opacity = "1";
                if (data.status === 'success') { ekranGecisi(data.data); } else { showToast(data.message); }
            }).catch(e => {
                spinner.style.display = "none"; btn.style.opacity = "1"; showToast("Sunucu ile bağlantı kurulamadı.");
            });
        }

        function showToast(mesaj) {
            const toast = document.getElementById("toast"); toast.innerText = mesaj; toast.classList.add("show");
            setTimeout(() => { toast.classList.remove("show"); }, 3000);
        }

        function ekranGecisi(userData) {
            const loginPanel = document.getElementById('login-panel'); const dashboard = document.getElementById('dashboard');
            loginPanel.style.transform = "scale(0.9)"; loginPanel.style.opacity = "0";
            
            setTimeout(() => {
                loginPanel.style.display = "none";
                document.getElementById('isimSoyisim').innerText = userData['ADI SOYADI'];
                document.getElementById('gorevBadge').innerText = userData['GÖREVİ'];
                dashboard.style.display = "block";
                
                setTimeout(() => {
                    dashboard.style.opacity = "1"; dashboard.style.transform = "scale(1)";
                    sayacAnimasyonu(userData['KALAN İZİN HAKKI']); grafikCiz(userData['PAZAR İZİNLERİ'], userData['RESMİ TATİL']);
                }, 50);
            }, 500);
        }

        function sayacAnimasyonu(hedefRakam) {
            const gosterge = document.getElementById('kalanIzinGosterge'); let baslangic = 0; const adim = hedefRakam / (1500 / 20); 
            const interval = setInterval(() => {
                baslangic += adim;
                if (baslangic >= hedefRakam) { gosterge.innerText = hedefRakam; clearInterval(interval); } 
                else { gosterge.innerText = baslangic.toFixed(1).replace('.0', ''); }
            }, 20);
        }

        function grafikCiz(pazar, resmi) {
            const ctx = document.getElementById('izinGrafik').getContext('2d');
            new Chart(ctx, {
                type: 'doughnut',
                data: {
                    labels: ['Pazar İzinleri', 'Resmi Tatil'],
                    datasets: [{ data: [pazar, resmi], backgroundColor: ['#00f2fe', '#4facfe'], borderWidth: 0, hoverOffset: 10 }]
                },
                options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { position: 'bottom', labels: { color: 'white' } } }, cutout: '70%' }
            });
        }
    </script>
</body>
</html>
"""

def get_user_data(username, password):
    try:
        df = pd.read_excel("BURHAN BİLİKTÜ İZİN.xlsx", sheet_name="Sayfa1")
        df.columns = df.columns.str.strip() # Sütun isimlerindeki boşlukları siler
        
        user_row = df[(df['KULLANICI ADI'].astype(str) == str(username))]
        if not user_row.empty:
            excel_pw = str(user_row.iloc[0]['ŞİFRE'])
            if excel_pw.endswith('.0'):
                excel_pw = excel_pw[:-2]
                
            if excel_pw == str(password):
                return user_row.iloc[0].fillna(0).to_dict()
    except Exception as e:
        print("Excel okuma hatası:", e)
    return None

@app.route('/')
def index():
    # Artık render_template değil, yukarıdaki HTML_SAYFASI değişkenini ekrana basıyoruz
    return render_template_string(HTML_SAYFASI)

@app.route('/login', methods=['POST'])
def login():
    data = request.get_json()
    username = data.get('username')
    password = data.get('password')
    
    user_data = get_user_data(username, password)
    
    if user_data:
        return jsonify({"status": "success", "data": user_data})
    else:
        return jsonify({"status": "error", "message": "Kullanıcı adı veya şifre hatalı!"})

if __name__ == '__main__':
    app.run(debug=True)