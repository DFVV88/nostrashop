/*
  NostraShop Perú - Backend gratuito con Google Apps Script
  ---------------------------------------------------------
  Qué hace este script:
  1. Recibe pedidos desde la web por POST.
  2. Guarda el resumen del pedido en la hoja 'Pedidos'.
  3. Guarda cada producto en la hoja 'DetalleProductos'.
  4. Agrega costos referenciales, ganancia y margen para contabilidad básica.
  5. Envía correo de confirmación al cliente.
  6. Envía correo de aviso al dueño de la tienda.
*/

const OWNER_EMAIL = 'fernandodaniel8888@gmail.com';
const SHEET_NAME = 'Pedidos';
const DETAIL_SHEET_NAME = 'DetalleProductos';

// Costos referenciales de prueba. Luego los cambiaremos por costos reales del proveedor.
const PRODUCT_COSTS = {
  'organizador-cocina-001': { cost: 35, supplier: 'Proveedor referencial A' },
  'limpiador-electrico-002': { cost: 45, supplier: 'Proveedor referencial B' },
  'luz-sensor-003': { cost: 22, supplier: 'Proveedor referencial C' }
};

const PRODUCT_COSTS_BY_NAME = {
  'Organizador plegable multiuso': { cost: 35, supplier: 'Proveedor referencial A' },
  'Cepillo limpiador eléctrico': { cost: 45, supplier: 'Proveedor referencial B' },
  'Luz LED con sensor': { cost: 22, supplier: 'Proveedor referencial C' }
};

function doGet() {
  return jsonResponse({
    ok: true,
    message: 'Backend NostraShop activo',
    now: new Date().toISOString()
  });
}

function doPost(e) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(10000);

    if (!e || !e.postData || !e.postData.contents) {
      throw new Error('No se recibieron datos del pedido.');
    }

    const order = JSON.parse(e.postData.contents);
    validateOrder(order);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ordersSheet = getOrCreateOrdersSheet_(ss);
    const detailsSheet = getOrCreateDetailsSheet_(ss);
    const normalized = normalizeOrder_(order);

    ordersSheet.appendRow([
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
      normalized.totalCost,
      normalized.totalProfit,
      normalized.marginPercent,
      'Nuevo',
      'Pedido recibido desde la web'
    ]);

    appendProductDetails_(detailsSheet, normalized);

    formatOrdersLastRow_(ordersSheet);
    formatDetailsSheet_(detailsSheet);

    sendCustomerEmail_(normalized);
    sendOwnerEmail_(normalized);

    return jsonResponse({
      ok: true,
      message: 'Pedido registrado correctamente',
      orderId: normalized.orderId,
      productsCount: normalized.productsCount,
      totalProfit: normalized.totalProfit
    });
  } catch (error) {
    console.error(error);
    return jsonResponse({
      ok: false,
      message: error.message || 'Error al registrar pedido'
    });
  } finally {
    try {
      lock.releaseLock();
    } catch (err) {}
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
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);

  const headers = [
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
    'Total venta',
    'Costo total',
    'Ganancia total',
    'Margen %',
    'Estado',
    'Observaciones'
  ];

  ensureHeaders_(sheet, headers);
  return sheet;
}

function getOrCreateDetailsSheet_(ss) {
  let sheet = ss.getSheetByName(DETAIL_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(DETAIL_SHEET_NAME);

  const headers = [
    'Fecha',
    'ID Pedido',
    'Cliente',
    'Celular',
    'Correo',
    'Producto',
    'Cantidad',
    'Precio unitario',
    'Total venta producto',
    'Costo unitario ref.',
    'Costo total ref.',
    'Ganancia unitaria ref.',
    'Ganancia total ref.',
    'Margen producto %',
    'Proveedor ref.',
    'Método de pago',
    'Estado'
  ];

  ensureHeaders_(sheet, headers);
  return sheet;
}

function ensureHeaders_(sheet, headers) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
  } else {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);
}

function appendProductDetails_(sheet, order) {
  const rows = order.items.map(item => {
    const qty = Number(item.qty || 1);
    const price = Number(item.price || 0);
    const productInfo = getProductCostInfo_(item);
    const unitCost = productInfo.cost;
    const totalSale = qty * price;
    const totalCost = qty * unitCost;
    const unitProfit = price - unitCost;
    const totalProfit = totalSale - totalCost;
    const marginPercent = totalSale > 0 ? totalProfit / totalSale : 0;

    return [
      order.createdAt,
      order.orderId,
      order.name,
      order.phone,
      order.email,
      item.name || 'Producto',
      qty,
      price,
      totalSale,
      unitCost,
      totalCost,
      unitProfit,
      totalProfit,
      marginPercent,
      productInfo.supplier,
      order.payment,
      'Nuevo'
    ];
  });

  if (rows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  }
}

function getProductCostInfo_(item) {
  const id = item.id || '';
  const name = item.name || '';
  return PRODUCT_COSTS[id] || PRODUCT_COSTS_BY_NAME[name] || { cost: 0, supplier: 'Sin proveedor asignado' };
}

function formatOrdersLastRow_(sheet) {
  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow, 1, 1, 17).setVerticalAlignment('middle');
  sheet.getRange(lastRow, 9).setWrap(true);
  sheet.getRange(lastRow, 10, 1, 5).setNumberFormat('S/ #,##0.00');
  sheet.getRange(lastRow, 15).setNumberFormat('0.00%');
  sheet.setColumnWidth(9, 420);
  sheet.setColumnWidth(7, 230);
  sheet.autoResizeColumns(1, 8);
  sheet.autoResizeColumns(10, 8);
}

function formatDetailsSheet_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 8, lastRow - 1, 6).setNumberFormat('S/ #,##0.00');
    sheet.getRange(2, 14, lastRow - 1, 1).setNumberFormat('0.00%');
  }

  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 220);
  sheet.setColumnWidth(5, 230);
  sheet.setColumnWidth(6, 260);
  sheet.setColumnWidth(15, 220);
  sheet.autoResizeColumns(1, 17);
}

function normalizeOrder_(order) {
  const customer = order.customer || {};
  const totals = order.totals || {};
  const items = order.items || [];

  let totalCost = 0;
  let totalProfit = 0;

  const itemsText = items.map((item, index) => {
    const name = item.name || 'Producto';
    const qty = Number(item.qty || 1);
    const price = Number(item.price || 0);
    const productInfo = getProductCostInfo_(item);
    const lineTotal = price * qty;
    const lineCost = productInfo.cost * qty;
    const lineProfit = lineTotal - lineCost;

    totalCost += lineCost;
    totalProfit += lineProfit;

    return `${index + 1}. ${name} x${qty} - Venta: S/ ${lineTotal.toFixed(2)} - Costo ref.: S/ ${lineCost.toFixed(2)} - Ganancia ref.: S/ ${lineProfit.toFixed(2)}`;
  }).join('\n');

  const total = Number(totals.total || 0);
  const marginPercent = total > 0 ? totalProfit / total : 0;

  return {
    createdAt: order.createdAt ? new Date(order.createdAt) : new Date(),
    orderId: order.orderId || `NS-${Date.now()}`,
    name: customer.name || '',
    phone: customer.phone || '',
    email: customer.email || '',
    city: customer.city || '',
    address: customer.address || '',
    payment: customer.payment || '',
    items,
    itemsText,
    productsCount: items.length,
    subtotal: Number(totals.subtotal || 0),
    shipping: Number(totals.shipping || 0),
    total,
    totalCost,
    totalProfit,
    marginPercent
  };
}

function sendCustomerEmail_(order) {
  if (!order.email) return;

  const subject = `Pedido recibido ${order.orderId} - NostraShop Perú`;
  const customerItemsText = order.items.map((item, index) => {
    const qty = Number(item.qty || 1);
    const price = Number(item.price || 0);
    return `${index + 1}. ${item.name || 'Producto'} x${qty} - S/ ${(price * qty).toFixed(2)}`;
  }).join('\n');

  const body = `Hola ${order.name || 'cliente'},\n\n` +
    `Hemos recibido tu pedido en NostraShop Perú.\n\n` +
    `ID de pedido: ${order.orderId}\n` +
    `Productos:\n${customerItemsText}\n` +
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
    `Productos:\n${order.itemsText}\n\n` +
    `Total venta: S/ ${order.total.toFixed(2)}\n` +
    `Costo total ref.: S/ ${order.totalCost.toFixed(2)}\n` +
    `Ganancia total ref.: S/ ${order.totalProfit.toFixed(2)}\n` +
    `Margen ref.: ${(order.marginPercent * 100).toFixed(2)}%\n\n` +
    `Estado inicial: Nuevo`;

  MailApp.sendEmail(OWNER_EMAIL, subject, body);
}

function jsonResponse(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
