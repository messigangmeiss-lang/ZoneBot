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
GUILD_ID = int(os.getenv("GUILD_ID", "1564463273615560706"))
LOG_CHANNEL_ID = int(os.getenv("LOG_CHANNEL_ID", "1567935320803250227"))
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
        # Ahora vuelve a subir los buenos
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
            
    mis_nicks = [r.get("nick","") for r in data if r.get("user") == str(interaction.user.id)]
    embed = discord.Embed(title="✅ Nuevo registro", color=0x2ecc71)
    embed.add_field(name="🧾 Nick:", value=nick, inline=False)
    embed.add_field(name="👤 Usuario:", value=f"@{interaction.user.name}", inline=False)
    embed.add_field(name="📅 Fecha:", value=datetime.now().strftime("%d/%m/%Y %H:%M"), inline=False)
    embed.add_field(name=f"📋 Tus nicks registrados ({len(mis_nicks)}):", value=", ".join(mis_nicks), inline=False)
    embed.set_thumbnail(url=interaction.user.display_avatar.url)
    embed.set_footer(text="ZoneBot • Registro de nicks")
    await interaction.response.send_message(embed=embed, ephemeral=False)
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
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>ZoneBot • Panel de Control</title>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
    <style>
        :root {
            --bg-base: #09090b;
            --bg-card: #121216;
            --border: #27272a;
            --text-main: #f4f4f5;
            --text-muted: #a1a1aa;
            --accent: #3b82f6;
            --danger: #ef4444;
            --success: #22c55e;
        }
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body {
            background-color: var(--bg-base);
            color: var(--text-main);
            font-family: 'Inter', sans-serif;
            padding: 30px 20px;
            max-width: 1200px;
            margin: 0 auto;
        }
        header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            border-bottom: 1px solid var(--border);
            padding-bottom: 20px;
            margin-bottom: 25px;
        }
        .logo-area h1 { font-size: 24px; font-weight: 700; letter-spacing: -0.5px; }
        .logo-area p { color: var(--text-muted); font-size: 13px; margin-top: 4px; }
        .badge-online {
            background: rgba(34, 197, 94, 0.1);
            color: var(--success);
            border: 1px solid rgba(34, 197, 94, 0.2);
            padding: 6px 12px;
            border-radius: 20px;
            font-size: 12px;
            font-weight: 600;
            display: flex;
            align-items: center;
            gap: 6px;
        }
        .badge-online::before {
            content: '';
            width: 8px;
            height: 8px;
            background: var(--success);
            border-radius: 50%;
            display: inline-block;
        }
        .cards {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
            gap: 16px;
            margin-bottom: 25px;
        }
        .card {
            background-color: var(--bg-card);
            border: 1px solid var(--border);
            border-radius: 12px;
            padding: 20px;
            transition: border-color 0.2s;
        }
        .card:hover { border-color: #3f3f46; }
        .card b { font-size: 28px; display: block; font-weight: 700; margin-bottom: 4px; }
        .card span { color: var(--text-muted); font-size: 13px; font-weight: 500; }
        .search-box {
            width: 100%;
            padding: 14px 16px;
            border-radius: 10px;
            border: 1px solid var(--border);
            background-color: var(--bg-card);
            color: var(--text-main);
            font-size: 14px;
            margin-bottom: 25px;
            outline: none;
            transition: border-color 0.2s;
        }
        .search-box:focus { border-color: var(--accent); }
        .grid-container {
            display: grid;
            grid-template-columns: repeat(auto-fill, minmax(320px, 1fr));
            gap: 16px;
        }
        .user-card {
            background-color: var(--bg-card);
            border: 1px solid var(--border);
            border-radius: 12px;
            padding: 18px;
        }
        .user-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 12px;
            padding-bottom: 10px;
            border-bottom: 1px solid var(--border);
        }
        .user-header h3 { font-size: 15px; font-weight: 600; display: flex; align-items: center; gap: 8px; }
        .user-count { background: #27272a; color: var(--text-muted); padding: 2px 8px; border-radius: 6px; font-size: 11px; }
        .nick-row {
            display: flex;
            justify-content: space-between;
            align-items: center;
            padding: 8px 0;
            border-bottom: 1px dashed rgba(255,255,255,0.05);
            font-size: 13px;
        }
        .nick-row:last-child { border-bottom: none; }
        .nick-info b { font-weight: 600; color: #fff; }
        .nick-info small { display: block; color: var(--text-muted); font-size: 11px; margin-top: 2px; }
        button.btn-delete {
            background: rgba(239, 68, 68, 0.1);
            color: var(--danger);
            border: 1px solid rgba(239, 68, 68, 0.2);
            padding: 5px 10px;
            border-radius: 6px;
            font-size: 12px;
            font-weight: 500;
            cursor: pointer;
            transition: background 0.2s;
        }
        button.btn-delete:hover { background: var(--danger); color: #fff; }
        .empty-state { text-align: center; color: var(--text-muted); padding: 40px; font-size: 14px; grid-column: 1 / -1; }
        footer { text-align: center; color: var(--text-muted); font-size: 12px; margin-top: 40px; border-top: 1px solid var(--border); padding-top: 20px; }
    </style>
</head>
<body>
    <header>
        <div class="logo-area">
            <h1>⚡ ZONEBO PANEL</h1>
            <p>Control de registros en vivo • Actualización automática</p>
        </div>
        <div class="badge-online">Sistema Activo</div>
    </header>

    <div class="cards">
        <div class="card">
            <b id="totalNicks">0</b>
            <span>Nicks Totales</span>
        </div>
        <div class="card">
            <b id="totalUsers">0</b>
            <span>Colaboradores Activos</span>
        </div>
        <div class="card">
            <b id="totalPay" style="color: var(--success);">$0</b>
            <span>Total a Pagar ($4 / c-u)</span>
        </div>
    </div>

    <input type="text" id="search" class="search-box" placeholder="🔍 Buscar por nick o nombre de usuario..." oninput="filtrar()">

    <div id="lista" class="grid-container"></div>

    <footer>
        ZONEBO • Panel interno de gestión de cuentas
    </footer>

<script>
let datos = [];

async function load() {
    try {
        let r = await fetch('/api/nicks');
        datos = await r.json();
        
        document.getElementById('totalNicks').innerText = datos.length;
        document.getElementById('totalPay').innerText = '$' + (datos.length * 4);
        
        let users = {};
        datos.forEach(d => {
            let uName = d.name || d.user;
            users[uName] = users[uName] || [];
            users[uName].push(d);
        });
        
        document.getElementById('totalUsers').innerText = Object.keys(users).length;
        render(users);
    } catch (e) {
        console.error("Error cargando datos", e);
    }
}

function render(users) {
    let q = document.getElementById('search').value.toLowerCase();
    let html = '';
    
    for (let u in users) {
        let nicks = users[u].filter(n => (n.nick + u).toLowerCase().includes(q));
        if (!nicks.length) continue;
        
        html += `
            <div class="user-card">
                <div class="user-header">
                    <h3>👤 ${u}</h3>
                    <span class="user-count">${nicks.length} registradas</span>
                </div>
        `;
        
        nicks.forEach(n => {
            html += `
                <div class="nick-row">
                    <div class="nick-info">
                        <b>${n.nick}</b>
                        <small>${n.fecha || 'Sin fecha'}</small>
                    </div>
                    <button class="btn-delete" onclick="borrar('${n.nick}')">Eliminar</button>
                </div>
            `;
        });
        
        html += `</div>`;
    }
    
    document.getElementById('lista').innerHTML = html || '<div class="empty-state">No se encontraron registros coincidentes.</div>';
}

function filtrar() {
    let users = {};
    datos.forEach(d => {
        let uName = d.name || d.user;
        users[uName] = users[uName] || [];
        users[uName].push(d);
    });
    render(users);
}

async function borrar(nick) {
    let key = prompt('Introduce la clave de administrador para borrar este nick:');
    if (!key) return;
    
    try {
        let r = await fetch('/api/delete', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ nick, key })
        });
        let j = await r.json();
        alert(j.msg);
        load();
    } catch (e) {
        alert('❌ Error al conectar con el servidor.');
    }
}

load();
setInterval(load, 10000);
</script>
</body>
</html>
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
@bot.tree.command(name="mynicks", description="Ver tus nicks registrados")
async def mynicks(interaction: discord.Interaction):
    data = cargar()
    mis_nicks = [r for r in data if r.get("user") == str(interaction.user.id)]
    if not mis_nicks:
        embed = discord.Embed(title="📋 Tus nicks", description="No tienes ningún nick registrado todavía.", color=0xe74c3c)
        embed.set_thumbnail(url=interaction.user.display_avatar.url)
        await interaction.response.send_message(embed=embed, ephemeral=True)
        return
    lista = "\n".join([f"**{i+1}.** {r.get('nick','')} — {r.get('fecha','')}" for i, r in enumerate(mis_nicks)])
    embed = discord.Embed(title="📋 Tus nicks registrados", description=lista, color=0x3498db)
    embed.add_field(name="📊 Total:", value=f"{len(mis_nicks)} nicks", inline=False)
    embed.set_thumbnail(url=interaction.user.display_avatar.url)
    embed.set_footer(text="ZoneBot • Registro de nicks")
    await interaction.response.send_message(embed=embed, ephemeral=True)


if __name__ == "__main__":
    threading.Thread(target=run_web, daemon=True).start()
    bot.run(TOKEN)
