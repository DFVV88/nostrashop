/*
  NostraShop Perú - Backend gratuito con Google Apps Script
  ---------------------------------------------------------
  Qué hace este script:
  1. Recibe pedidos desde la web por POST.
  2. Guarda el resumen del pedido en la hoja 'Pedidos'.
  3. Guarda cada producto en la hoja 'DetalleProductos'.
  4. Agrega costos referenciales, ganancia y margen para contabilidad básica.
  5. Envía correo de confirmación profesional al cliente.
  6. Envía correo de aviso al dueño de la tienda.
*/

const OWNER_EMAIL = 'fernandodaniel8888@gmail.com';
const STORE_NAME = 'NostraShop Perú';
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
    'Fecha', 'ID Pedido', 'Nombre', 'Celular', 'Correo', 'Ciudad', 'Dirección',
    'Método de pago', 'Productos', 'Subtotal', 'Envío', 'Total venta', 'Costo total',
    'Ganancia total', 'Margen %', 'Estado', 'Observaciones'
  ];

  ensureHeaders_(sheet, headers);
  return sheet;
}

function getOrCreateDetailsSheet_(ss) {
  let sheet = ss.getSheetByName(DETAIL_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(DETAIL_SHEET_NAME);

  const headers = [
    'Fecha', 'ID Pedido', 'Cliente', 'Celular', 'Correo', 'Producto', 'Cantidad',
    'Precio unitario', 'Total venta producto', 'Costo unitario ref.', 'Costo total ref.',
    'Ganancia unitaria ref.', 'Ganancia total ref.', 'Margen producto %', 'Proveedor ref.',
    'Método de pago', 'Estado'
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
      order.createdAt, order.orderId, order.name, order.phone, order.email,
      item.name || 'Producto', qty, price, totalSale, unitCost, totalCost,
      unitProfit, totalProfit, marginPercent, productInfo.supplier, order.payment, 'Nuevo'
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

  const subject = `Confirmación de pedido ${order.orderId} - ${STORE_NAME}`;
  const customerItemsText = order.items.map((item, index) => {
    const qty = Number(item.qty || 1);
    const price = Number(item.price || 0);
    return `${index + 1}. ${item.name || 'Producto'} x${qty} - S/ ${(price * qty).toFixed(2)}`;
  }).join('\n');

  const plainBody = `Hola ${order.name || 'cliente'},\n\n` +
    `Gracias por comprar en ${STORE_NAME}. Hemos recibido tu pedido.\n\n` +
    `ID de pedido: ${order.orderId}\n` +
    `Productos:\n${customerItemsText}\n` +
    `Total: S/ ${order.total.toFixed(2)}\n` +
    `Método de pago: ${order.payment}\n\n` +
    `Datos de entrega:\nCiudad: ${order.city}\nDirección: ${order.address}\nCelular: ${order.phone}\n\n` +
    `Siguiente paso: validaremos el pago y coordinaremos la preparación del pedido.\n\n` +
    `${STORE_NAME}`;

  const htmlBody = buildCustomerHtmlEmail_(order);

  GmailApp.sendEmail(order.email, subject, plainBody, {
    name: STORE_NAME,
    htmlBody: htmlBody,
    replyTo: OWNER_EMAIL
  });
}

function buildCustomerHtmlEmail_(order) {
  const rows = order.items.map(item => {
    const qty = Number(item.qty || 1);
    const price = Number(item.price || 0);
    const lineTotal = price * qty;
    return `
      <tr>
        <td style="padding:12px;border-bottom:1px solid #e5e7eb;color:#111827;">${escapeHtml_(item.name || 'Producto')}</td>
        <td style="padding:12px;border-bottom:1px solid #e5e7eb;text-align:center;color:#111827;">${qty}</td>
        <td style="padding:12px;border-bottom:1px solid #e5e7eb;text-align:right;color:#111827;">S/ ${lineTotal.toFixed(2)}</td>
      </tr>`;
  }).join('');

  return `
  <div style="margin:0;padding:0;background:#f3f4f6;font-family:Arial,Helvetica,sans-serif;color:#111827;">
    <div style="max-width:620px;margin:0 auto;padding:24px;">
      <div style="background:linear-gradient(135deg,#020617,#0f172a);border-radius:22px 22px 0 0;padding:28px;text-align:center;color:#ffffff;">
        <div style="font-size:30px;font-weight:800;letter-spacing:-1px;">${STORE_NAME}</div>
        <div style="margin-top:8px;color:#a5f3fc;font-size:14px;">Confirmación automática de pedido</div>
      </div>

      <div style="background:#ffffff;border-radius:0 0 22px 22px;padding:26px;border:1px solid #e5e7eb;border-top:0;">
        <h2 style="margin:0 0 10px;font-size:22px;color:#111827;">¡Pedido recibido correctamente!</h2>
        <p style="margin:0 0 18px;line-height:1.6;color:#374151;">Hola <strong>${escapeHtml_(order.name || 'cliente')}</strong>, gracias por tu compra. Hemos registrado tu pedido y pronto continuaremos con la validación del pago y la coordinación de entrega.</p>

        <div style="background:#ecfeff;border:1px solid #bae6fd;border-radius:16px;padding:16px;margin:18px 0;">
          <div style="font-size:13px;color:#0369a1;font-weight:700;">ID DE PEDIDO</div>
          <div style="font-size:20px;color:#0f172a;font-weight:800;margin-top:4px;">${escapeHtml_(order.orderId)}</div>
        </div>

        <table style="width:100%;border-collapse:collapse;margin:18px 0;border:1px solid #e5e7eb;border-radius:14px;overflow:hidden;">
          <thead>
            <tr style="background:#f9fafb;">
              <th style="padding:12px;text-align:left;color:#374151;font-size:13px;">Producto</th>
              <th style="padding:12px;text-align:center;color:#374151;font-size:13px;">Cant.</th>
              <th style="padding:12px;text-align:right;color:#374151;font-size:13px;">Importe</th>
            </tr>
          </thead>
          <tbody>${rows}</tbody>
        </table>

        <div style="text-align:right;margin-top:12px;">
          <div style="color:#6b7280;font-size:14px;">Total del pedido</div>
          <div style="font-size:26px;font-weight:900;color:#111827;">S/ ${order.total.toFixed(2)}</div>
        </div>

        <div style="margin-top:22px;padding:16px;border-radius:16px;background:#f9fafb;border:1px solid #e5e7eb;">
          <div style="font-weight:800;margin-bottom:8px;color:#111827;">Datos de entrega</div>
          <div style="line-height:1.7;color:#374151;font-size:14px;">
            <strong>Ciudad:</strong> ${escapeHtml_(order.city)}<br>
            <strong>Dirección:</strong> ${escapeHtml_(order.address)}<br>
            <strong>Celular:</strong> ${escapeHtml_(order.phone)}<br>
            <strong>Método de pago:</strong> ${escapeHtml_(order.payment)}
          </div>
        </div>

        <p style="margin:22px 0 0;line-height:1.6;color:#6b7280;font-size:13px;">Este correo fue generado automáticamente por ${STORE_NAME}. Si necesitas corregir algún dato, responde a este mismo correo.</p>
      </div>
    </div>
  </div>`;
}

function escapeHtml_(value) {
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
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

  GmailApp.sendEmail(OWNER_EMAIL, subject, body, { name: STORE_NAME });
}

function jsonResponse(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
