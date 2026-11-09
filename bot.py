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
ADMIN_KEY = os.getenv("ADMIN_KEY", "zone123") # clave para borrar y editar desde la web

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

# ========== BOT COMMANDS ==========
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
            
    fecha_actual = datetime.now().strftime("%d/%m/%Y %H:%M")
    data.append({"nick": nick, "user": str(interaction.user.id), "name": str(interaction.user), "fecha": fecha_actual})
    guardar(data)
    
    # Contar nicks del usuario para el embed
    nicks_usuario = [r["nick"] for r in data if r.get("user") == str(interaction.user.id)]
    
    embed = discord.Embed(
        title="✅ Nuevo registro",
        color=discord.Color.green(),
        timestamp=datetime.now()
    )
    embed.add_field(name="📄 Nick:", value=nick, inline=False)
    embed.add_field(name="👤 Usuario:", value=interaction.user.mention, inline=False)
    embed.add_field(name="📅 Fecha:", value=fecha_actual, inline=False)
    embed.add_field(name=f"📋 Tus nicks registrados ({len(nicks_usuario)}):", value=", ".join(nicks_usuario), inline=False)
    
    if interaction.user.avatar:
        embed.set_thumbnail(url=interaction.user.avatar.url)
        
    embed.set_footer(text="ZoneBot • Registro de nicks")
    
    await interaction.response.send_message(embed=embed, ephemeral=True)
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
    mios_list = [r["nick"] for r in data if r.get("user") == str(interaction.user.id)]
    if not mios_list:
        await interaction.response.send_message("No tienes nicks registrados.", ephemeral=True)
        return
    await interaction.response.send_message("Tus nicks:\n" + "\n".join([f"- **{n}**" for n in mios_list]), ephemeral=True)

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

# ========== DASHBOARD WEB CON MODAL DE EDITAR Y DEGRADADOS ==========
app = Flask(__name__)

HTML = """
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>ZONEBOT • Panel de Control</title>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
    <style>
        :root {
            --bg-base: #09090b;
            --bg-card: linear-gradient(145deg, #18181b, #0f0f12);
            --border: #27272a;
            --text-main: #f4f4f5;
            --text-muted: #a1a1aa;
            --accent: #3b82f6;
            --danger: #ef4444;
            --success: #22c55e;
            --warning: #f59e0b;
        }
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body {
            background: radial-gradient(circle at top, #180d1c 0%, var(--bg-base) 60%);
            color: var(--text-main);
            font-family: 'Inter', sans-serif;
            padding: 30px 20px;
            max-width: 1200px;
            margin: 0 auto;
            min-height: 100vh;
        }
        header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            border-bottom: 1px solid var(--border);
            padding-bottom: 20px;
            margin-bottom: 25px;
        }
        .logo-area h1 { 
            font-size: 26px; 
            font-weight: 700; 
            letter-spacing: -0.5px;
            background: linear-gradient(90deg, #fff, #a1a1aa);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
        }
        .logo-area p { color: var(--text-muted); font-size: 13px; margin-top: 4px; }
        .badge-online {
            background: rgba(34, 197, 94, 0.1);
            color: var(--success);
            border: 1px solid rgba(34, 197, 94, 0.2);
            padding: 6px 14px;
            border-radius: 20px;
            font-size: 12px;
            font-weight: 600;
            display: flex;
            align-items: center;
            gap: 6px;
            box-shadow: 0 0 15px rgba(34, 197, 94, 0.1);
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
            background: var(--bg-card);
            border: 1px solid var(--border);
            border-radius: 14px;
            padding: 20px;
            box-shadow: 0 10px 30px rgba(0,0,0,0.3);
            transition: transform 0.2s, border-color 0.2s;
        }
        .card:hover { transform: translateY(-2px); border-color: #3f3f46; }
        .card b { font-size: 28px; display: block; font-weight: 700; margin-bottom: 4px; }
        .card span { color: var(--text-muted); font-size: 13px; font-weight: 500; }
        .search-box {
            width: 100%;
            padding: 14px 16px;
            border-radius: 12px;
            border: 1px solid var(--border);
            background: linear-gradient(145deg, #121216, #09090b);
            color: var(--text-main);
            font-size: 14px;
            margin-bottom: 25px;
            outline: none;
            transition: border-color 0.2s, box-shadow 0.2s;
        }
        .search-box:focus { border-color: var(--accent); box-shadow: 0 0 10px rgba(59, 130, 246, 0.2); }
        .grid-container {
            display: grid;
            grid-template-columns: repeat(auto-fill, minmax(320px, 1fr));
            gap: 16px;
        }
        .user-card {
            background: var(--bg-card);
            border: 1px solid var(--border);
            border-radius: 14px;
            padding: 18px;
            box-shadow: 0 8px 24px rgba(0,0,0,0.2);
        }
        .user-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 12px;
            padding-bottom:
