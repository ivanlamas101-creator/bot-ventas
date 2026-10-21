from datetime import datetime
import io
import os
import re
import unicodedata
import pandas as pd
from PIL import Image, ImageDraw, ImageFont
from telegram import InlineKeyboardButton, InlineKeyboardMarkup, Update
from telegram.ext import (
    ApplicationBuilder,
    CallbackQueryHandler,
    CommandHandler,
    ContextTypes,
    MessageHandler,
    filters,
)

# --- CONFIGURACIÓN ---
TOKEN = "8820960067:AAHYOLbJ595UTCQA3NVYZdI1rJoROsjX1ik"
EXCEL_FILE = "datos.xlsx"


# Función para cargar productos
def cargar_productos():
  if os.path.exists(EXCEL_FILE):
    try:
      return pd.read_excel(
          EXCEL_FILE, sheet_name="Productos", engine="openpyxl"
      )
    except Exception:
      pass

  df = pd.DataFrame({
      "Producto": ["pipa", "mogul", "coca cola", "quilmes"],
      "Precio": [500, 1200, 800, 1500],
      "Stock": [50, 100, 30, 40],
  })
  with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
    df.to_excel(writer, sheet_name="Productos", index=False)
    pd.DataFrame(
        columns=["Fecha_Hora", "Producto", "Cantidad", "Total"]
    ).to_excel(writer, sheet_name="Ventas", index=False)
  return df


# Normalizar texto (elimina tildes, pasa a minúsculas)
def normalizar(texto):
  nfkd = unicodedata.normalize("NFKD", texto)
  sin_tildes = "".join([c for c in nfkd if not unicodedata.combining(c)])
  return sin_tildes.lower().strip()


# Buscar el mejor producto con razonamiento de similitud y plurales
def buscar_mejor_producto(nombre_ingresado, df_productos):
  nombre_ingresado = normalizar(nombre_ingresado)
  mejor_prod = None
  mejor_precio = 0
  max_score = 0

  for _, row in df_productos.iterrows():
    prod_real = str(row["Producto"])
    prod_norm = normalizar(prod_real)

    if prod_norm in nombre_ingresado or nombre_ingresado in prod_norm:
      score = len(prod_norm)
      if score > max_score:
        max_score = score
        mejor_prod = row["Producto"]
        mejor_precio = float(row["Precio"])
    else:
      prod_singular = (
          prod_norm[:-1] if prod_norm.endswith("s") else prod_norm
      )
      ing_singular = (
          nombre_ingresado[:-1]
          if nombre_ingresado.endswith("s")
          else nombre_ingresado
      )

      if prod_singular in ing_singular or ing_singular in prod_singular:
        score = len(prod_singular)
        if score > max_score:
          max_score = score
          mejor_prod = row["Producto"]
          mejor_precio = float(row["Precio"])

  return mejor_prod, mejor_precio


# Guardar venta en el historial Y descontar stock en el Excel
def guardar_en_historial(items_venta):
  fecha_hora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

  df_productos = pd.read_excel(
      EXCEL_FILE, sheet_name="Productos", engine="openpyxl"
  )
  try:
    df_ventas = pd.read_excel(EXCEL_FILE, sheet_name="Ventas", engine="openpyxl")
  except Exception:
    df_ventas = pd.DataFrame(
        columns=["Fecha_Hora", "Producto", "Cantidad", "Total"]
    )

  nuevas_ventas = []
  for item in items_venta:
    prod_nombre = item["producto"]
    cant_vendida = item["cantidad"]
    subtotal = item["subtotal"]

    nuevas_ventas.append({
        "Fecha_Hora": fecha_hora,
        "Producto": prod_nombre,
        "Cantidad": cant_vendida,
        "Total": subtotal,
    })

    if "Stock" in df_productos.columns:
      mask = df_productos["Producto"].str.lower() == prod_nombre.lower()
      if mask.any():
        stock_actual = int(df_productos.loc[mask, "Stock"].values[0])
        nuevo_stock = max(0, stock_actual - cant_vendida)
        df_productos.loc[mask, "Stock"] = nuevo_stock

  df_nueva_fila = pd.DataFrame(nuevas_ventas)
  df_ventas_final = pd.concat([df_ventas, df_nueva_fila], ignore_index=True)

  with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
    df_productos.to_excel(writer, sheet_name="Productos", index=False)
    df_ventas_final.to_excel(writer, sheet_name="Ventas", index=False)


# --- FUNCIÓN PARA GENERAR LA IMAGEN TICKET DE CIERRE ---
def generar_imagen_ticket(resumen_prod, df_productos, total_items, total_dinero):
  fecha_actual = datetime.now().strftime("%d/%m/%Y %H:%M")

  ancho = 500
  altura_base = 250
  altura_por_item = 30
  num_items = len(resumen_prod) + len(df_productos)
  alto = altura_base + (num_items * altura_por_item)

  img = Image.new("RGB", (ancho, alto), color=(255, 255, 255))
  d = ImageDraw.Draw(img)

  try:
    font_titulo = ImageFont.truetype("arialbd.ttf", 20)
    font_texto = ImageFont.truetype("arial.ttf", 15)
    font_negrita = ImageFont.truetype("arialbd.ttf", 15)
  except Exception:
    font_titulo = ImageFont.load_default()
    font_texto = ImageFont.load_default()
    font_negrita = ImageFont.load_default()

  d.text((25, 20), "NEGOCIO CLUB - CIERRE DE CAJA", fill=(0, 0, 0), font=font_titulo)
  d.text((25, 50), f"Fecha: {fecha_actual}", fill=(100, 100, 100), font=font_texto)
  d.line([(25, 80), (475, 80)], fill=(200, 200, 200), width=2)

  y = 100
  d.text((25, y), "📦 VENTAS REALIZADAS:", fill=(0, 0, 0), font=font_negrita)
  y += 25

  for _, row in resumen_prod.iterrows():
    nombre = str(row["Producto"]).capitalize()
    cant = int(row["Cantidad"])
    tot = float(row["Total"])
    texto_item = f"• {nombre}: {cant} un."
    texto_precio = f"${tot:,.0f}"
    d.text((35, y), texto_item, fill=(50, 50, 50), font=font_texto)
    d.text((380, y), texto_precio, fill=(50, 50, 50), font=font_texto)
    y += 22

  y += 10
  d.line([(25, y), (475, y)], fill=(200, 200, 200), width=1)
  y += 15

  d.text((25, y), "🏪 STOCK RESTANTE:", fill=(0, 0, 0), font=font_negrita)
  y += 25

  for _, row in df_productos.iterrows():
    nombre = str(row["Producto"]).capitalize()
    stock = int(row["Stock"]) if "Stock" in row else 0
    d.text((35, y), f"• {nombre}:", fill=(50, 50, 50), font=font_texto)
    d.text(
        (380, y), f"{stock} disp.", fill=(50, 50, 50), font=font_texto
    )
    y += 22

  y += 10
  d.line([(25, y), (475, y)], fill=(0, 0, 0), width=2)
  y += 20

  d.text((25, y), f"Total Unidades: {total_items}", fill=(0, 0, 0), font=font_negrita)
  y += 25
  d.text((25, y), f"TOTAL RECOLECTADO: ${total_dinero:,.0f}", fill=(0, 120, 0), font=font_titulo)

  bio = io.BytesIO()
  bio.name = "cierre_caja.png"
  img.save(bio, "PNG")
  bio.seek(0)
  return bio


# Comando /start
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
  mensaje = (
      "👋 **¡Bienvenido! Éxito el día de hoy - Comencemos cuando quieras.**\n\n"
      "Solo escribe el producto y te diré todo lo que necesites."
  )

  teclado = [
      [InlineKeyboardButton("🔍 Ver Precios", callback_data="menu_precio")],
      [InlineKeyboardButton("📊 Cierre del Día", callback_data="menu_cierre")],
  ]

  if update.message:
    await update.message.reply_text(
        mensaje, parse_mode="Markdown", reply_markup=InlineKeyboardMarkup(teclado)
    )
  elif update.callback_query:
    query = update.callback_query
    await query.answer()
    await query.edit_message_text(
        mensaje, parse_mode="Markdown", reply_markup=InlineKeyboardMarkup(teclado)
    )


# Manejador de botones
async def boton_menu(update: Update, context: ContextTypes.DEFAULT_TYPE):
  query = update.callback_query
  await query.answer()

  if query.data == "menu_precio":
    df = cargar_productos()
    lista = "🏷️ **Lista de Productos, Precios y Stock:**\n\n"
    for index, row in df.iterrows():
      stock_disp = row["Stock"] if "Stock" in row else "?"
      lista += f"• {str(row['Producto']).capitalize()}: ${row['Precio']} (Stock: {stock_disp})\n"
    lista += "\n*(Escríbeme tu venta en el chat cuando quieras)*"

    teclado = [[InlineKeyboardButton("🔙 Volver", callback_data="volver")]]
    await query.edit_message_text(
        lista, parse_mode="Markdown", reply_markup=InlineKeyboardMarkup(teclado)
    )

  elif query.data == "menu_cierre":
    try:
      df_ventas = pd.read_excel(
          EXCEL_FILE, sheet_name="Ventas", engine="openpyxl"
      )
      df_productos = pd.read_excel(
          EXCEL_FILE, sheet_name="Productos", engine="openpyxl"
      )

      if df_ventas.empty:
        resumen = "📊 **Cierre de Caja del Día**\n\n❌ Aún no se han registrado ventas hoy."
        teclado = [[InlineKeyboardButton("🔙 Volver", callback_data="volver")]]
        await query.edit_message_text(
            resumen, parse_mode="Markdown", reply_markup=InlineKeyboardMarkup(teclado)
        )
      else:
        resumen_prod = (
            df_ventas.groupby("Producto")
            .agg({"Cantidad": "sum", "Total": "sum"})
            .reset_index()
        )

        total_dinero = df_ventas["Total"].sum()
        total_items = df_ventas["Cantidad"].sum()

        imagen_ticket = generar_imagen_ticket(
            resumen_prod, df_productos, total_items, total_dinero
        )

        teclado_ticket = [
            [
                InlineKeyboardButton(
                    "🚀 Iniciar el Día",
                    callback_data="cerrar_caja_definitivo",
                )
            ],
            [InlineKeyboardButton("🔙 Menú Principal", callback_data="volver")],
        ]

        # Enviar la foto con el botón de Iniciar el Día
        await query.message.reply_photo(
            photo=imagen_ticket,
            caption=(
                "📄 **Comprobante de Cierre Generado.**\n"
                "El stock ya fue descontado en el sistema. Presiona **'🚀 Iniciar el Día'** abajo para archivar las ventas, dejar la caja en $0 y arrancar la nueva jornada con el stock actualizado."
            ),
            parse_mode="Markdown",
            reply_markup=InlineKeyboardMarkup(teclado_ticket),
        )

        # Borrar el mensaje de texto anterior del menú para limpiar la pantalla
        try:
          await query.message.delete()
        except Exception:
          pass

    except Exception as e:
      await query.edit_message_text(
          f"❌ Error al generar el cierre: {e}",
          reply_markup=InlineKeyboardMarkup(
              [[InlineKeyboardButton("🔙 Volver", callback_data="volver")]]
          ),
      )

  elif query.data == "cerrar_caja_definitivo":
    try:
      fecha_historial = datetime.now().strftime("%Y%m%d_%H%M%S")
      df_ventas = pd.read_excel(
          EXCEL_FILE, sheet_name="Ventas", engine="openpyxl"
      )
      df_productos = pd.read_excel(
          EXCEL_FILE, sheet_name="Productos", engine="openpyxl"
      )

      with pd.ExcelWriter(
          EXCEL_FILE, mode="a", if_sheet_exists="replace", engine="openpyxl"
      ) as writer:
        df_productos.to_excel(writer, sheet_name="Productos", index=False)
        pd.DataFrame(
            columns=["Fecha_Hora", "Producto", "Cantidad", "Total"]
        ).to_excel(writer, sheet_name="Ventas", index=False)
        df_ventas.to_excel(
            writer, sheet_name=f"Cierre_{fecha_historial}", index=False
        )

      teclado = [
          [InlineKeyboardButton("🔙 Menú Principal", callback_data="volver")]
      ]
      await query.message.reply_text(
          "🚀 **¡Nuevo día iniciado con éxito!**\n\n"
          "• Las ventas anteriores se archivaron.\n"
          "• La caja comenzó en $0.\n"
          "• **El stock actual ya quedó actualizado en el Excel** para que comiences a vender con base en él.",
          parse_mode="Markdown",
          reply_markup=InlineKeyboardMarkup(teclado),
      )
    except Exception as e:
      await query.message.reply_text(
          f"⚠️ **Error al iniciar el día:** Asegúrate de tener el archivo Excel cerrado. ({e})"
      )

  elif query.data == "volver":
    context.user_data.pop("pendiente_venta", None)
    await start(update, context)

  elif query.data == "confirmar_venta":
    items_pendientes = context.user_data.get("pendiente_venta")
    if items_pendientes:
      try:
        guardar_en_historial(items_pendientes)
        context.user_data.pop("pendiente_venta", None)

        teclado = [
            [InlineKeyboardButton("🔙 Menú Principal", callback_data="volver")]
        ]
        await query.edit_message_text(
            "✅ **¡Venta registrada, stock descontado y guardado con"
            " éxito!** 📊",
            parse_mode="Markdown",
            reply_markup=InlineKeyboardMarkup(teclado),
        )
      except PermissionError:
        await query.edit_message_text(
            "⚠️ **Error:** Cierra el archivo `datos.xlsx` en tu computadora."
        )
    else:
      await query.edit_message_text(
          "❌ No hay ninguna venta pendiente para guardar."
      )

  elif query.data == "cancelar_venta":
    context.user_data.pop("pendiente_venta", None)
    teclado = [[InlineKeyboardButton("🔙 Menú Principal", callback_data="volver")]]
    await query.edit_message_text(
        "❌ **Venta cancelada.** No se guardó nada en el historial.",
        parse_mode="Markdown",
        reply_markup=InlineKeyboardMarkup(teclado),
    )


# Interpretador de texto libre inteligente
async def interpretar_venta(
    update: Update, context: ContextTypes.DEFAULT_TYPE
):
  texto_raw = update.message.text
  if texto_raw.startswith("/") or texto_raw.lower().strip() in [
      "hola",
      "buenas",
      "ola",
  ]:
    return

  texto = normalizar(texto_raw)

  palabras_relleno = [
      "vendi",
      "vendí",
      "vende",
      "quiero",
      "cobrar",
      "el",
      "la",
      "los",
      "las",
      "un",
      "una",
      "de",
      "y",
      "con",
  ]
  for palabra in palabras_relleno:
    texto = re.sub(r"\b" + palabra + r"\b", "", texto)

  df_productos = cargar_productos()
  segmentos = re.split(r"[,.]", texto)

  items_venta = []
  total_general = 0

  for seg in segmentos:
    seg = seg.strip()
    if not seg:
      continue

    matches = list(re.finditer(r"(\d+)", seg))

    if not matches:
      nombre_buscado = seg.strip()
      if nombre_buscado:
        prod, precio = buscar_mejor_producto(nombre_buscado, df_productos)
        if prod:
          subtotal = precio * 1
          total_general += subtotal
          items_venta.append({
              "producto": prod,
              "cantidad": 1,
              "subtotal": subtotal,
              "precio_unitario": precio,
          })
    else:
      for i, match in enumerate(matches):
        cantidad = int(match.group(1))
        inicio_num = match.start()
        fin_num = match.end()

        fin_texto = (
            matches[i + 1].start() if i < len(matches) - 1 else len(seg)
        )
        inicio_texto = matches[i - 1].end() if i > 0 else 0

        texto_antes = seg[inicio_texto:inicio_num].strip()
        texto_despues = seg[fin_num:fin_texto].strip()

        nombre_buscado = texto_despues if texto_despues else texto_antes
        if not nombre_buscado:
          nombre_buscado = texto_antes

        if nombre_buscado:
          prod, precio = buscar_mejor_producto(nombre_buscado, df_productos)
          if prod:
            subtotal = precio * cantidad
            total_general += subtotal
            items_venta.append({
                "producto": prod,
                "cantidad": cantidad,
                "subtotal": subtotal,
                "precio_unitario": precio,
            })

  if items_venta:
    context.user_data["pendiente_venta"] = items_venta

    detalle_texto = ""
    for item in items_venta:
      detalle_texto += f"• {item['cantidad']}x {str(item['producto']).capitalize()} (${item['precio_unitario']} c/u) = **${item['subtotal']}**\n"

    mensaje = (
        f"🛒 **Resumen de la venta detectada:**\n\n"
        f"{detalle_texto}\n"
        f"💵 **Total a cobrar: ${total_general}**\n\n"
        f"¿Deseas registrar esta venta y descontar stock?"
    )

    teclado = [
        [
            InlineKeyboardButton(
                "✅ Sí, Registrar Venta", callback_data="confirmar_venta"
            )
        ],
        [InlineKeyboardButton("❌ Cancelar", callback_data="cancelar_venta")],
    ]

    await update.message.reply_text(
        mensaje, parse_mode="Markdown", reply_markup=InlineKeyboardMarkup(teclado)
    )
  else:
    await update.message.reply_text(
        "🤔 No logré identificar los productos. Escribe algo como:\n👉 *'3"
        " pipa, 6 mogul'*",
        parse_mode="Markdown",
    )


def main():
  app = (
      ApplicationBuilder()
      .token(TOKEN)
      .connect_timeout(30.0)
      .read_timeout(30.0)
      .write_timeout(30.0)
      .build()
  )

  app.add_handler(CommandHandler("start", start))
  app.add_handler(CallbackQueryHandler(boton_menu))
  app.add_handler(
      MessageHandler(filters.TEXT & ~filters.COMMAND, interpretar_venta)
  )

  print("🤖 ¡Bot corregido y listo para iniciar el día desde el comprobante!")
  app.run_polling()


if __name__ == "__main__":
  main()