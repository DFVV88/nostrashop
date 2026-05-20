/*
  NostraShop Perú - Backend gratuito con Google Apps Script
  ---------------------------------------------------------
  Qué hace este script:
  1. Recibe pedidos desde la web por POST.
  2. Crea automáticamente la hoja 'Pedidos' si no existe.
  3. Guarda el pedido en Google Sheets.
  4. Envía correo de confirmación al cliente.
  5. Envía correo de aviso al dueño de la tienda.

  Uso:
  - Crea una hoja de cálculo en Google Sheets llamada 'Pedidos NostraShop'.
  - Abre Extensiones > Apps Script.
  - Pega este código.
  - Cambia OWNER_EMAIL si deseas recibir avisos en otro correo.
  - Implementa como aplicación web con acceso: Cualquier persona.
  - Copia la URL /exec y pégala en CONFIG.orderWebhookUrl del index.html.
*/

const OWNER_EMAIL = 'fernandodaniel8888@gmail.com';
const SHEET_NAME = 'Pedidos';

function doGet() {
  return jsonResponse({
    ok: true,
    message: 'Backend NostraShop activo',
    now: new Date().toISOString()
  });
}

function doPost(e) {
  try {
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);

    if (!e || !e.postData || !e.postData.contents) {
      throw new Error('No se recibieron datos del pedido.');
    }

    const order = JSON.parse(e.postData.contents);
    validateOrder(order);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = getOrCreateOrdersSheet_(ss);
    const normalized = normalizeOrder_(order);

    sheet.appendRow([
      normalized.createdAt,
      normalized.orderId,
      normalized.name,
      normalized.phone,
      normalized.email,
      normalized.city,
      normalized.address,
      normalized.payment,
      normalized.itemsText,
      normalized.subtotal,
      normalized.shipping,
      normalized.total,
      'Nuevo',
      'Pedido recibido desde la web'
    ]);

    sendCustomerEmail_(normalized);
    sendOwnerEmail_(normalized);

    lock.releaseLock();

    return jsonResponse({
      ok: true,
      message: 'Pedido registrado correctamente',
      orderId: normalized.orderId
    });
  } catch (error) {
    console.error(error);
    return jsonResponse({
      ok: false,
      message: error.message || 'Error al registrar pedido'
    });
  }
}

function validateOrder(order) {
  if (!order) throw new Error('Pedido vacío.');
  if (!order.customer) throw new Error('Faltan datos del cliente.');
  if (!order.items || !Array.isArray(order.items) || order.items.length === 0) {
    throw new Error('El pedido no tiene productos.');
  }
  if (!order.totals) throw new Error('Faltan totales del pedido.');
}

function getOrCreateOrdersSheet_(ss) {
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      'Fecha',
      'ID Pedido',
      'Nombre',
      'Celular',
      'Correo',
      'Ciudad',
      'Dirección',
      'Método de pago',
      'Productos',
      'Subtotal',
      'Envío',
      'Total',
      'Estado',
      'Observaciones'
    ]);
    sheet.getRange(1, 1, 1, 14).setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, 14);
  }

  return sheet;
}

function normalizeOrder_(order) {
  const customer = order.customer || {};
  const totals = order.totals || {};
  const items = order.items || [];
  const itemsText = items.map(item => {
    const name = item.name || 'Producto';
    const qty = Number(item.qty || 1);
    const price = Number(item.price || 0);
    return `${name} x${qty} - S/ ${(price * qty).toFixed(2)}`;
  }).join(' | ');

  return {
    createdAt: order.createdAt ? new Date(order.createdAt) : new Date(),
    orderId: order.orderId || `NS-${Date.now()}`,
    name: customer.name || '',
    phone: customer.phone || '',
    email: customer.email || '',
    city: customer.city || '',
    address: customer.address || '',
    payment: customer.payment || '',
    itemsText,
    subtotal: Number(totals.subtotal || 0),
    shipping: Number(totals.shipping || 0),
    total: Number(totals.total || 0)
  };
}

function sendCustomerEmail_(order) {
  if (!order.email) return;

  const subject = `Pedido recibido ${order.orderId} - NostraShop Perú`;
  const body = `Hola ${order.name || 'cliente'},\n\n` +
    `Hemos recibido tu pedido en NostraShop Perú.\n\n` +
    `ID de pedido: ${order.orderId}\n` +
    `Productos: ${order.itemsText}\n` +
    `Total: S/ ${order.total.toFixed(2)}\n` +
    `Método de pago: ${order.payment}\n\n` +
    `Tus datos de entrega:\n` +
    `Ciudad: ${order.city}\n` +
    `Dirección: ${order.address}\n` +
    `Celular: ${order.phone}\n\n` +
    `Siguiente paso: validaremos el pago y coordinaremos la preparación del pedido.\n\n` +
    `Gracias por comprar en NostraShop Perú.`;

  MailApp.sendEmail(order.email, subject, body);
}

function sendOwnerEmail_(order) {
  if (!OWNER_EMAIL) return;

  const subject = `Nueva venta web: ${order.orderId}`;
  const body = `Nuevo pedido recibido desde la web.\n\n` +
    `ID: ${order.orderId}\n` +
    `Cliente: ${order.name}\n` +
    `Celular: ${order.phone}\n` +
    `Correo: ${order.email}\n` +
    `Ciudad: ${order.city}\n` +
    `Dirección: ${order.address}\n` +
    `Pago: ${order.payment}\n` +
    `Productos: ${order.itemsText}\n` +
    `Total: S/ ${order.total.toFixed(2)}\n\n` +
    `Estado inicial: Nuevo`;

  MailApp.sendEmail(OWNER_EMAIL, subject, body);
}

function jsonResponse(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
