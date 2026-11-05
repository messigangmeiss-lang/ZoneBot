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
body{background:#0b0505;color:#fff;font-family:'Inter',Arial,sans-serif;min-height:100vh}
@keyframes fadeUp{from{opacity:0;transform:translateY(15px)}to{opacity:1;transform:translateY(0)}}
@keyframes gradientMove{0%{background-position:0% 50%}50%{background-position:100% 50%}100%{background-position:0% 50%}}
@keyframes pulseDot{0%,100%{transform:scale(1);box-shadow:0 0 0 0 rgba(34,197,94,.5)}50%{transform:scale(1.15);box-shadow:0 0 0 6px rgba(34,197,94,0)}}
.topbar{height:4px;background:linear-gradient(90deg,#ef4444,#f97316,#dc2626,#ef4444);background-size:300% 100%;animation:gradientMove 4s ease infinite}
header{max-width:1100px;margin:0 auto;padding:26px 20px 8px;display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:10px;animation:fadeUp .4s ease}
.logo{font-size:32px;font-weight:900;letter-spacing:-1px}
.logo span{background:linear-gradient(90deg,#ef4444,#f97316,#fca5a5);background-size:200% 100%;-webkit-background-clip:text;-webkit-text-fill-color:transparent;animation:gradientMove 3s ease infinite}
.badge{display:flex;align-items:center;gap:7px;background:#14532d;color:#bbf7d0;font-weight:800;font-size:12px;padding:7px 13px;border-radius:999px}
.dot{width:9px;height:9px;background:#22c55e;border-radius:50%;animation:pulseDot 1.8s infinite}
.sub{max-width:1100px;margin:0 auto;padding:0 20px 16px;color:#a8a29e;font-size:14px}
.wrap{max-width:1100px;margin:0 auto;padding:0 20px 40px}
.stats{display:grid;grid-template-columns:1fr 1fr 1fr;gap:14px;margin:14px 0 18px}
@media(max-width:700px){.stats{grid-template-columns:1fr}.main{grid-template-columns:1fr !important}}
.card{background:linear-gradient(145deg,#1c0a0a,#160707);border:1px solid #3f1d1d;border-radius:16px;padding:18px;animation:fadeUp .5s ease;transition:.25s}
.card:hover{transform:translateY(-3px);border-color:#ef4444}
.card b{font-size:30px;display:block}
.card span{color:#a8a29e;font-size:13px}
.search{width:100%;padding:13px 15px;border-radius:12px;border:1px solid #3f1d1d;background:#1c0a0a;color:#fff;font-size:14px;outline:none;margin-bottom:16px}
.search:focus{border-color:#ef4444;box-shadow:0 0 0 3px rgba(239,68,68,.25)}
.main{display:grid;grid-template-columns:300px 1fr;gap:16px;align-items:start}
.left{background:#151010;border:1px solid #2a1a1a;border-radius:16px;padding:14px;animation:fadeUp .6s ease}
.left h3{font-size:15px;margin-bottom:12px}
.userItem{padding:12px 12px;border-radius:12px;cursor:pointer;border:1px solid transparent;transition:.2s;margin-bottom:8px}
.userItem:hover{background:#1f1212}
.userItem.active{background:#2a0f0f;border-color:#ef4444}
.userItem b{display:block;font-size:15px}
.userItem small{color:#a8a29e;font-size:12px}
.userItem.active small{color:#fca5a5}
.right{background:#151010;border:1px solid #2a1a1a;border-radius:16px;padding:18px;min-height:340px;animation:fadeUp .6s ease}
.rightHead{display:flex;justify-content:space-between;align-items:flex-start;gap:10px;flex-wrap:wrap;margin-bottom:6px}
.rightHead h2{font-size:20px}
.rightHead p{color:#a8a29e;font-size:13px;margin-top:4px}
.btnUpdate{background:#262020;color:#fff;border:1px solid #3a3a3a;padding:9px 14px;border-radius:999px;cursor:pointer;font-size:13px;font-weight:700;transition:.2s}
.btnUpdate:hover{background:#ef4444;border-color:#ef4444;transform:scale(1.03)}
.nick{display:flex;justify-content:space-between;align-items:center;padding:13px 8px;border-top:1px solid #262020;gap:10px}
.nick:hover{background:#1f1212;border-radius:10px}
.nick b{font-size:15px}
.nick small{color:#888;font-size:12px}
.actions{display:flex;gap:6px}
.btnCopy{background:#262020;color:#fff;border:1px solid #3a3a3a;padding:7px 11px;border-radius:999px;cursor:pointer;font-size:12px;font-weight:700}
.btnCopy:hover{background:#fff;color:#000}
.btnDel{background:#2a1212;color:#fca5a5;border:none;padding:7px 12px;border-radius:999px;cursor:pointer;font-size:12px;font-weight:700}
.btnDel:hover{background:#ef4444;color:#fff}
.empty{text-align:center;color:#888;padding:50px 10px}
footer{text-align:center;color:#555;font-size:12px;padding:20px}
</style></head><body>
<div class="topbar"></div>
<header><div class="logo">ZONE<span>BOT</span></div><div class="badge"><span class="dot"></span> ONLINE</div></header>
<p class="sub">Nicks registrados en vivo desde Discord • Actualiza solo cada 10 segundos</p>
<div class="wrap">
<div class="stats">
<div class="card"><b id="totalUsers">0</b><span>Usuarios registrados</span></div>
<div class="card"><b id="totalNicks">0</b><span>Nicks totales</span></div>
<div class="card"><b style="color:#22c55e">Online</b><span>Estado del Bot</span></div>
</div>
<input id="search" class="search" placeholder="Buscar nick o usuario..." oninput="filtrar()">
<div class="main">
<div class="left"><h3>Usuarios</h3><div id="userList"></div></div>
<div class="right" id="detail"></div>
</div>
<footer>ZONEBOT • Hecho para la comunidad de Zone</footer>
</div>
<script>
let datos=[]; let selected=null; let filterQ='';
async function load(){ let r=await fetch('/api/nicks'); datos=await r.json();
document.getElementById('totalNicks').innerText=datos.length;
let users=getUsers(); document.getElementById('totalUsers').innerText=Object.keys(users).length;
if(!selected || !users[selected]){ let keys=Object.keys(users); selected=keys[0]||null; }
render(); }
function getUsers(){ let users={}; datos.forEach(d=>{ let k=d.name||d.user; users[k]=users[k]||[]; users[k].push(d)}); return users; }
function filtrar(){ filterQ=document.getElementById('search').value.toLowerCase(); render(); }
function render(){
let users=getUsers();
let filtered={}; for(let u in users){ let nicks=users[u].filter(n=>(n.nick+' '+n.name).toLowerCase().includes(filterQ)); if(nicks.length) filtered[u]=nicks; }
let html='';
for(let u in filtered){ html+=`<div class="userItem ${u===selected?'active':''}" onclick="selectUser('${u.replace(/'/g,"")}')"><b>${u}</b><small>${filtered[u].length} nicks</small></div>`; }
document.getElementById('userList').innerHTML=html||'<p style="color:#888;font-size:13px">No hay usuarios</p>';
if(!selected || !filtered[selected]){
  let keys=Object.keys(filtered); selected=keys[0]||null;
  if(!selected){ document.getElementById('detail').innerHTML='<div class="empty">Aun no hay nicks<br><small>Registra uno en Discord con /register</small></div>'; return; }
}
let nicks=filtered[selected]||[];
let totalHtml=`<div class="rightHead"><div><h2>${selected}</h2><p>${nicks.length} nicks • Click en Actualizar para recargar</p></div><button class="btnUpdate" onclick="actualizar()">↻ Actualizar</button></div>` + nicks.map(n=>`<div class="nick"><span><b>${n.nick}</b> ${n.fecha?`<small>• ${n.fecha}</small>`:''}</span><div class="actions"><button class="btnCopy" onclick="copiar('${n.nick}',this)">Copiar</button><button class="btnDel" onclick="borrar('${n.nick}')">Eliminar</button></div></div>`).join('');
document.getElementById('detail').innerHTML=totalHtml;
}
function selectUser(u){ selected=u; render(); }
async function actualizar(){ let btn=document.querySelector('.btnUpdate'); if(btn) btn.innerText='Actualizando...'; await load(); }
function copiar(nick,btn){ navigator.clipboard.writeText(nick); let o=btn.innerText; btn.innerText='Copiado!'; setTimeout(()=>btn.innerText=o,1000); }
async function borrar(nick){ let key=prompt('Clave de admin para borrar:'); if(!key) return; let r=await fetch('/api/delete',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({nick,key})}); let j=await r.json(); alert(j.msg); load(); }
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
