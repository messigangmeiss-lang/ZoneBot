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

TOKEN = os.getenv("TOKEN")
GUILD_ID = int(os.getenv("GUILD_ID", "1567715818366832650"))
LOG_CHANNEL_ID = int(os.getenv("LOG_CHANNEL_ID", "1567760248654143609"))
ADMIN_KEY = os.getenv("ADMIN_KEY", "zone123") # clave para borrar desde la web

intents = discord.Intents.default()
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
<!DOCTYPE html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>ZoneBot Dashboard</title>
<style>
body{background:#0f0f12;color:#fff;font-family:Arial;padding:20px}
h1{font-size:28px} .cards{display:flex;gap:12px;flex-wrap:wrap;margin:20px 0}
.card{background:#1a1a20;border:1px solid #2a2a33;border-radius:12px;padding:16px;min-width:160px;flex:1}
.card b{font-size:28px;display:block} .card span{color:#aaa;font-size:13px}
input{width:100%;padding:12px;border-radius:10px;border:1px solid #333;background:#1a1a20;color:#fff;margin:10px 0}
.user{background:#1a1a20;border:1px solid #2a2a33;border-radius:12px;padding:14px;margin:10px 0}
.user h3{margin:0 0 8px 0} .nick{display:flex;justify-content:space-between;padding:8px 0;border-top:1px solid #2a2a33}
button{background:#ff3b30;color:#fff;border:none;padding:6px 12px;border-radius:8px;cursor:pointer}
</style></head><body>
<h1>🎮 ZoneBot Dashboard</h1><p style="color:#aaa">Nicks registrados en vivo desde Discord</p>
<div class="cards">
<div class="card"><b id="totalUsers">0</b><span>Usuarios registrados</span></div>
<div class="card"><b id="totalNicks">0</b><span>Nicks totales</span></div>
<div class="card"><b id="status">🟢 Online</b><span>Estado del Bot</span></div>
</div>
<input id="search" placeholder="🔍 Buscar nick o usuario..." oninput="filtrar()">
<div id="lista"></div>
<script>
let datos=[];
async function load(){ let r=await fetch('/api/nicks'); datos=await r.json();
document.getElementById('totalNicks').innerText=datos.length;
let users={}; datos.forEach(d=>{users[d.name||d.user]=users[d.name||d.user]||[]; users[d.name||d.user].push(d)});
document.getElementById('totalUsers').innerText=Object.keys(users).length;
render(users); }
function render(users){ let q=document.getElementById('search').value.toLowerCase(); let html='';
for(let u in users){ let nicks=users[u].filter(n=> (n.nick+n.name).toLowerCase().includes(q)); if(!nicks.length) continue;
html+=`<div class="user"><h3>👤 ${u} <span style="color:#aaa;font-size:13px">(${nicks.length} nicks)</span></h3>`+nicks.map(n=>`<div class="nick"><span><b>${n.nick}</b> <small style="color:#888"> - ${n.fecha||''}</small></span><button onclick="borrar('${n.nick}')">Eliminar</button></div>`).join('')+`</div>`; }
document.getElementById('lista').innerHTML=html||'<p>No hay nicks</p>'; }
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