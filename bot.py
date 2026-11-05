import discord
from discord import app_commands
from discord.ext import commands
import json
import os
import threading
from datetime import datetime
from pathlib import Path
from openpyxl import Workbook
from flask import Flask, request, jsonify, render_template_string

TOKEN = os.getenv("TOKEN") or os.getenv("DISCORD_TOKEN")
GUILD_ID = int(os.getenv("GUILD_ID", "1567715818366832650"))
LOG_CHANNEL_ID = int(os.getenv("LOG_CHANNEL_ID", "1567760248654143609"))
ADMIN_KEY = os.getenv("ADMIN_KEY", "zone123") # clave para borrar desde la web

intents = discord.Intents.default()
intents.message_content = True
intents.members = True
intents.presences = True


bot = commands.Bot(command_prefix="!", intents=intents)
nicks_file = Path("nicks.json")

def cargar():
    if not nicks_file.exists():
        return []
    try:
        with open(nicks_file, "r", encoding="utf-8") as f:
            data = json.load(f)
            # Compatibilidad: convierte formato viejo a nuevo
            out = []
            for r in data:
                if "registrado_por" in r:
                    out.append({"nick": r.get("nick",""), "user": str(r.get("user_id","")), "name": r.get("registrado_por",""), "fecha": r.get("fecha","")})
                else:
                    out.append(r)
            return out
    except:
        return []

def guardar(registros):
    with open(nicks_file, "w", encoding="utf-8") as f:
        json.dump(registros, f, indent=2, ensure_ascii=False)

async def mandar_log(texto, color=discord.Color.greyple()):
    if LOG_CHANNEL_ID == 0:
        return
    canal = bot.get_channel(LOG_CHANNEL_ID)
    if canal:
        embed = discord.Embed(description=texto, color=color)
        await canal.send(embed=embed)

# ========== BOT COMMANDS (igual que antes) ==========
@bot.event
async def on_ready():
    print(f"Bot conectado como {bot.user}")
    try:
        guild = discord.Object(id=GUILD_ID)
        bot.tree.copy_global_to(guild=guild)
        synced = await bot.tree.sync(guild=guild)
        print(f"Sincronizados {len(synced)} comandos")
    except Exception as e:
        print(f"Error sync: {e}")

@bot.tree.command(name="register", description="Registra tu nick")
@app_commands.describe(nick="Tu nick en el juego")
async def register(interaction: discord.Interaction, nick: str):
    nick = nick.strip()
    data = cargar()
    for r in data:
        if r.get("nick","").lower() == nick.lower():
            await interaction.response.send_message(f"❌ El nick **{nick}** ya está ocupado.", ephemeral=True)
            return
    data.append({"nick": nick, "user": str(interaction.user.id), "name": str(interaction.user), "fecha": datetime.now().strftime("%d/%m/%Y %H:%M")})
    guardar(data)
    await interaction.response.send_message(f"✅ Registraste el nick **{nick}**", ephemeral=True)
    await mandar_log(f"✅ **{interaction.user.mention}** registró el nick **{nick}**", discord.Color.green())

@bot.tree.command(name="unregister", description="Borra uno de tus nicks")
@app_commands.describe(nick="Nick que quieres borrar")
async def unregister(interaction: discord.Interaction, nick: str):
    data = cargar()
    for i, r in enumerate(data):
        if r.get("nick","").lower() == nick.strip().lower():
            if r.get("user") != str(interaction.user.id) and not interaction.user.guild_permissions.administrator:
                await interaction.response.send_message("❌ Ese nick no es tuyo.", ephemeral=True)
                return
            del data[i]
            guardar(data)
            await interaction.response.send_message(f"🗑️ Borraste el nick **{nick}**", ephemeral=True)
            await mandar_log(f"🗑️ **{interaction.user.mention}** borró el nick **{nick}**", discord.Color.red())
            return
    await interaction.response.send_message(f"❌ No existe el nick **{nick}**.", ephemeral=True)

@bot.tree.command(name="mios", description="Ver tus nicks registrados")
async def mios(interaction: discord.Interaction):
    data = cargar()
    mios = [r["nick"] for r in data if r.get("user") == str(interaction.user.id)]
    if not mios:
        await interaction.response.send_message("No tienes nicks registrados.", ephemeral=True)
        return
    await interaction.response.send_message("Tus nicks:\n" + "\n".join([f"- **{n}**" for n in mios]), ephemeral=True)

@bot.tree.command(name="buscar", description="Buscar si un nick está ocupado")
@app_commands.describe(nick="Nick a buscar")
async def buscar(interaction: discord.Interaction, nick: str):
    data = cargar()
    for r in data:
        if r.get("nick","").lower() == nick.strip().lower():
            await interaction.response.send_message(f"❌ **{r['nick']}** está ocupado por **{r.get('name','')}**", ephemeral=True)
            return
    await interaction.response.send_message(f"✅ **{nick}** está libre!", ephemeral=True)

@bot.tree.command(name="todos", description="Ver todos los nicks registrados")
async def todos(interaction: discord.Interaction):
    data = cargar()
    if not data:
        await interaction.response.send_message("Aún no hay nicks.", ephemeral=True)
        return
    texto = "\n".join([f"- **{r['nick']}** | {r.get('name','')}" for r in data])
    await interaction.response.send_message(f"**Nicks registrados ({len(data)}):**\n{texto}", ephemeral=True)

@bot.tree.command(name="ayuda", description="Ver todos los comandos")
async def ayuda(interaction: discord.Interaction):
    texto = """**📋 Comandos de Zone:**
/register <nick> - Registrar tu nick
/unregister <nick> - Borrar tu nick
/mios - Ver tus nicks
/buscar <nick> - Ver si un nick está libre
/todos - Ver todos los nicks
/ayuda - Ver esta ayuda
/export - Descargar backup en Excel (solo admins)
"""
    await interaction.response.send_message(texto, ephemeral=True)

@bot.tree.command(name="export", description="Descargar backup de nicks en Excel (solo admins)")
async def export(interaction: discord.Interaction):
    if not interaction.user.guild_permissions.administrator:
        await interaction.response.send_message("❌ Solo admins pueden usar esto.", ephemeral=True)
        return
    data = cargar()
    wb = Workbook()
    ws = wb.active
    ws.title = "Nicks Zone"
    ws.append(["Nick", "Usuario", "User ID", "Fecha Registro"])
    for r in data:
        ws.append([r.get("nick",""), r.get("name",""), r.get("user",""), r.get("fecha","")])
    ws.column_dimensions['A'].width = 25
    ws.column_dimensions['B'].width = 25
    ws.column_dimensions['C'].width = 22
    ws.column_dimensions['D'].width = 20
    archivo = "nicks_backup.xlsx"
    wb.save(archivo)
    await interaction.response.send_message("📥 Aquí está tu backup en Excel:", file=discord.File(archivo), ephemeral=True)
    await mandar_log(f"📥 **{interaction.user.mention}** descargó el backup en Excel ({len(data)} nicks)", discord.Color.blue())

# ========== DASHBOARD WEB - ESTO ES LO QUE MANTIENE PRENDIDO EL BOT ==========
app = Flask(__name__)

HTML = """
<!DOCTYPE html><html lang="es"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>ZONE - Nicks Registrados</title>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;600;800;900&display=swap" rel="stylesheet">
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{background:#0a0a0f;color:#fff;font-family:'Inter',Arial,sans-serif;min-height:100vh}
.topbar{background:linear-gradient(90deg,#7c3aed,#a855f7,#ec4899);padding:2px}
header{max-width:1000px;margin:0 auto;padding:28px 20px 10px;display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:12px}
.logo{font-size:32px;font-weight:900;letter-spacing:-1px}
.logo span{background:linear-gradient(90deg,#a78bfa,#f472b6);-webkit-background-clip:text;-webkit-text-fill-color:transparent}
.badge{background:#22c55e;color:#000;font-weight:800;font-size:12px;padding:6px 12px;border-radius:999px}
.sub{max-width:1000px;margin:0 auto;padding:0 20px 18px;color:#9ca3af;font-size:14px}
.wrap{max-width:1000px;margin:0 auto;padding:0 20px 40px}
.stats{display:grid;grid-template-columns:1fr 1fr 1fr;gap:14px;margin:18px 0}
@media(max-width:600px){.stats{grid-template-columns:1fr}}
.card{background:#15151d;border:1px solid #26262f;border-radius:16px;padding:18px}
.card b{font-size:32px;display:block;line-height:1}
.card span{color:#9ca3af;font-size:13px}
.card small{font-size:28px}
.search{width:100%;padding:14px 16px;border-radius:14px;border:1px solid #2e2e3a;background:#15151d;color:#fff;font-size:15px;outline:none;margin:6px 0 18px}
.search:focus{border-color:#a855f7;box-shadow:0 0 0 3px rgba(168,85,247,.25)}
.user{background:#15151d;border:1px solid #26262f;border-radius:16px;padding:16px;margin:14px 0}
.user-head{display:flex;justify-content:space-between;align-items:center;margin-bottom:12px;flex-wrap:wrap;gap:8px}
.user-head h3{font-size:17px}
.pill{background:#2e1065;color:#c4b5fd;padding:4px 10px;border-radius:999px;font-size:12px;font-weight:700}
.nick{display:flex;justify-content:space-between;align-items:center;padding:12px 10px;border-top:1px solid #26262f;gap:10px}
.nick b{font-size:15px}
.nick small{color:#888;font-size:12px}
.btn{background:#26262f;color:#fca5a5;border:none;padding:7px 12px;border-radius:999px;cursor:pointer;font-size:12px;font-weight:700}
.btn:hover{background:#ef4444;color:#fff}
.empty{text-align:center;color:#888;padding:40px}
footer{text-align:center;color:#555;font-size:12px;padding:20px}
</style></head><body>
<div class="topbar"></div>
<header><div class="logo">🎮 ZONE<span>BOT</span></div><div class="badge">● ONLINE</div></header>
<p class="sub">Nicks registrados en vivo desde Discord • Actualiza solo cada 10 segundos</p>
<div class="wrap">
<div class="stats">
<div class="card"><b id="totalUsers">0</b><span>👥 Usuarios registrados</span></div>
<div class="card"><b id="totalNicks">0</b><span>🏷️ Nicks totales</span></div>
<div class="card"><b>🟢 <small>Online</small></b><span>Estado del Bot</span></div>
</div>
<input id="search" class="search" placeholder="🔍 Buscar nick o usuario..." oninput="filtrar()">
<div id="lista"></div>
<footer>ZONEBOT • Hecho para la comunidad de Zone</footer>
</div>
<script>
let datos=[];
async function load(){ let r=await fetch('/api/nicks'); datos=await r.json();
document.getElementById('totalNicks').innerText=datos.length;
let users={}; datos.forEach(d=>{users[d.name||d.user]=users[d.name||d.user]||[]; users[d.name||d.user].push(d)});
document.getElementById('totalUsers').innerText=Object.keys(users).length;
render(users); }
function render(users){ let q=document.getElementById('search').value.toLowerCase(); let html='';
for(let u in users){ let nicks=users[u].filter(n=> (n.nick+n.name).toLowerCase().includes(q)); if(!nicks.length) continue;
html+=`<div class="user"><div class="user-head"><h3>👤 ${u}</h3><span class="pill">${nicks.length} nicks</span></div>`+nicks.map(n=>`<div class="nick"><span><b>${n.nick}</b> <small>• ${n.fecha||'sin fecha'}</small></span><button class="btn" onclick="borrar('${n.nick}')">Eliminar</button></div>`).join('')+`</div>`; }
document.getElementById('lista').innerHTML=html||'<div class="empty">😴 Aún no hay nicks<br><small>Registra uno en Discord con /register</small></div>'; }
function filtrar(){ let users={}; datos.forEach(d=>{users[d.name||d.user]=users[d.name||d.user]||[]; users[d.name||d.user].push(d)}); render(users); }
async function borrar(nick){ let key=prompt('Clave de admin para borrar:'); if(!key) return;
let r=await fetch('/api/delete',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({nick,key})});
let j=await r.json(); alert(j.msg); load(); }
load(); setInterval(load,10000);
</script></body></html>
"""

@app.route("/")
def home():
    return render_template_string(HTML)

@app.route("/api/nicks")
def api_nicks():
    return jsonify(cargar())

@app.route("/api/delete", methods=["POST"])
def api_delete():
    d = request.json or {}
    if d.get("key") != ADMIN_KEY:
        return jsonify({"msg": "❌ Clave incorrecta"}), 403
    data = cargar()
    for i, r in enumerate(data):
        if r.get("nick","").lower() == d.get("nick","").lower():
            del data[i]
            guardar(data)
            return jsonify({"msg": f"✅ Borrado {d.get('nick')}"})
    return jsonify({"msg": "❌ No existe"}), 404

def run_web():
    port = int(os.getenv("PORT", "10000"))
    app.run(host="0.0.0.0", port=port)

if __name__ == "__main__":
    threading.Thread(target=run_web, daemon=True).start()
    bot.run(TOKEN)
