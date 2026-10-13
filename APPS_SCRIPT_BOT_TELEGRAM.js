/**
 * 🌙 BOT DE TELEGRAM + GOOGLE SHEETS PARA "LUNI PIJAMAS" 🌙
 * 
 * Este script se coloca en el editor de Google Apps Script vinculado a una hoja de Google Sheets.
 * Permite registrar pedidos, consultar precios de combos y ver métricas desde Telegram.
 * 
 * -------------------------------------------------------------
 * ESTRUCTURA DE LA HOJA DE GOOGLE SHEETS (Crear estas 3 pestañas):
 * -------------------------------------------------------------
 * 1. Pestaña: "Ventas"
 *    Columnas (Fila 1): 
 *    [Fecha, Combo/Producto, Valor Venta, Medio de Pago, Cliente/Instagram, Estado]
 * 
 * 2. Pestaña: "Combos"
 *    Columnas (Fila 1): 
 *    [Nombre Combo, Incluye, Costo Aprox, Precio Sugerido]
 * 
 * 3. Pestaña: "Config"
 *    (Opcional para almacenar variables o metas)
 */

// ⚙️ CONFIGURACIÓN: Reemplaza con tus valores
const TELEGRAM_TOKEN = "PEGA_AQUI_TU_TELEGRAM_BOT_TOKEN";
const TELEGRAM_API = "https://api.telegram.org/bot" + TELEGRAM_TOKEN;

/**
 * Función Webhook que recibe las peticiones de Telegram
 */
function doPost(e) {
  try {
    const contents = JSON.parse(e.postData.contents);
    
    if (contents.message && contents.message.text) {
      const chatId = contents.message.chat.id;
      const text = contents.message.text.trim();
      const userName = contents.message.from.first_name || "Lune Lover";

      procesarComando(chatId, text, userName);
    }
  } catch (error) {
    Logger.log("Error en doPost: " + error.toString());
  }
}

/**
 * Enrutador de comandos de Telegram
 */
function procesarComando(chatId, text, userName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Comando de Bienvenida / Ayuda
  if (text.startsWith("/start") || text.startsWith("/ayuda")) {
    const mensaje = 
      "🌙 *¡Hola, " + userName + "! Bienvenido al Bot Operativo de Lune* 💜\n\n" +
      "Aquí puedes registrar pedidos y consultar datos al instante:\n\n" +
      "📌 *Comandos disponibles:*\n" +
      "👉 `/venta [Combo] | [Valor] | [MedioPago] | [Cliente]`\n" +
      "   _Ejemplo:_ `/venta Combo Dulces Sueños | 72000 | Nequi | @laura_m`\n\n" +
      "👉 `/combos` - Consulta los combos y precios sugeridos.\n" +
      "👉 `/resumen` - Ver total de ventas y ganancias del día.\n";
    enviarMensajeTelegram(chatId, mensaje);
    return;
  }

  // 2. Comando para Registrar Venta
  if (text.startsWith("/venta")) {
    const datos = text.replace("/venta", "").trim().split("|");
    
    if (datos.length < 3) {
      enviarMensajeTelegram(chatId, "⚠️ *Formato incorrecto.*\nUsa: `/venta Combo | Valor | MedioPago | @cliente`");
      return;
    }

    const combo = datos[0].trim();
    const valor = parseFloat(datos[1].replace(/[^0-9]/g, '')) || 0;
    const medioPago = datos[2].trim();
    const cliente = datos[3] ? datos[3].trim() : "Directo";
    const fecha = Utilities.formatDate(new Date(), "America/Bogota", "yyyy-MM-dd HH:mm");

    const hojaVentas = ss.getSheetByName("Ventas");
    if (!hojaVentas) {
      enviarMensajeTelegram(chatId, "❌ Error: No se encontró la pestaña llamada 'Ventas' en el Sheet.");
      return;
    }

    // Insertar fila
    hojaVentas.appendRow([fecha, combo, valor, medioPago, cliente, "Confirmado"]);

    const respuesta = 
      "✅ *¡Venta registrada con éxito en Lune!* 🌙✨\n\n" +
      "📦 *Producto:* " + combo + "\n" +
      "💰 *Valor:* $" + valor.toLocaleString("es-CO") + " COP\n" +
      "💳 *Medio:* " + medioPago + "\n" +
      "👤 *Cliente:* " + cliente + "\n" +
      "📅 *Fecha:* " + fecha;
    
    enviarMensajeTelegram(chatId, respuesta);
    return;
  }

  // 3. Comando para Consultar Combos
  if (text.startsWith("/combos")) {
    const hojaCombos = ss.getSheetByName("Combos");
    if (!hojaCombos) {
      // Mensaje por defecto si no ha llenado la hoja aún
      const defaultMsg = 
        "🌙 *Combos Oficiales Lune:*\n\n" +
        "✨ *Combo Dulces Sueños (Pijama Corta)*\n" +
        "• Incluye: Pijama corta + Antifaz/Scrunchie + Medias + Empaque\n" +
        "• Precio sugerido: *$70.000 COP*\n\n" +
        "🌙 *Combo Sleep & Glow (Pijama Larga)*\n" +
        "• Incluye: Pijama pantalón + Balaca spa + Skincare + Cajita\n" +
        "• Precio sugerido: *$82.000 COP*\n\n" +
        "👑 *Combo Reina de Noche (Satín)*\n" +
        "• Incluye: Pijama/Bata satín + Cosmetiquera + Antifaz + Bálsamo\n" +
        "• Precio sugerido: *$110.000 COP*";
      enviarMensajeTelegram(chatId, defaultMsg);
      return;
    }

    const filas = hojaCombos.getDataRange().getValues();
    let respuesta = "🌙 *Catálogo de Combos Lune:*\n\n";
    for (let i = 1; i < filas.length; i++) {
      respuesta += "✨ *" + filas[i][0] + "*\n• Incluye: " + filas[i][1] + "\n• PVP: *$" + Number(filas[i][3]).toLocaleString("es-CO") + " COP*\n\n";
    }
    enviarMensajeTelegram(chatId, respuesta);
    return;
  }

  // 4. Comando de Resumen del Día
  if (text.startsWith("/resumen")) {
    const hojaVentas = ss.getSheetByName("Ventas");
    if (!hojaVentas) {
      enviarMensajeTelegram(chatId, "❌ No se encontró la pestaña 'Ventas'.");
      return;
    }

    const filas = hojaVentas.getDataRange().getValues();
    const hoyStr = Utilities.formatDate(new Date(), "America/Bogota", "yyyy-MM-dd");
    let totalVendidoHoy = 0;
    let cantidadPedidosHoy = 0;

    for (let i = 1; i < filas.length; i++) {
      const fechaFila = String(filas[i][0]);
      if (fechaFila.includes(hoyStr)) {
        totalVendidoHoy += Number(filas[i][2]) || 0;
        cantidadPedidosHoy++;
      }
    }

    const respuesta = 
      "📊 *Resumen de Ventas Hoy (" + hoyStr + "):*\n\n" +
      "🛍️ *Total Pedidos:* " + cantidadPedidosHoy + "\n" +
      "💵 *Total Facturado:* $" + totalVendidoHoy.toLocaleString("es-CO") + " COP\n\n" +
      "¡Sigan soñando en grande con Lune! 🌙💛💜";

    enviarMensajeTelegram(chatId, respuesta);
    return;
  }

  // Mensaje por defecto si no reconoce el comando
  enviarMensajeTelegram(chatId, "Escribe `/ayuda` para ver los comandos disponibles para Lune 🌙");
}

/**
 * Función auxiliar para enviar mensajes a Telegram vía HTTP POST
 */
function enviarMensajeTelegram(chatId, textoMarkdown) {
  const url = TELEGRAM_API + "/sendMessage";
  const payload = {
    chat_id: chatId,
    text: textoMarkdown,
    parse_mode: "Markdown"
  };

  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload)
  };

  UrlFetchApp.fetch(url, options);
}

/**
 * Ejecuta esta función UNA SOLA VEZ para conectar tu Webhook de Telegram
 * tras publicar el Apps Script como Web App (Implementar > Nueva Implementación > Aplicación Web).
 */
function registrarWebhook() {
  const WEB_APP_URL = "PEGA_AQUI_LA_URL_DE_IMPLEMENTACION_DE_APPS_SCRIPT";
  const url = TELEGRAM_API + "/setWebhook?url=" + encodeURIComponent(WEB_APP_URL);
  const response = UrlFetchApp.fetch(url);
  Logger.log(response.getContentText());
}
