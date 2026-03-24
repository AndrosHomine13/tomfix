const STORAGE_KEY = "tomfix-orders";

const orderForm = document.getElementById("orderForm");
const generatedOrderSection = document.getElementById("generatedOrderSection");
const generatedOrder = document.getElementById("generatedOrder");
const ordersTableBody = document.getElementById("ordersTableBody");
const printOrderBtn = document.getElementById("printOrderBtn");
const exportExcelBtn = document.getElementById("exportExcelBtn");
const clearOrdersBtn = document.getElementById("clearOrdersBtn");
const COMPANY_INFO = {
  name: "TomFix",
  slogan: "Servicio técnico especializado en celulares",
  whatsapp: "302 5980351",
  email: "tomfixbga@gmail.com",
};

const orders = loadOrders();
renderOrdersTable();

ordersTableBody.addEventListener("change", (event) => {
  const target = event.target;
  if (!(target instanceof HTMLSelectElement)) return;
  if (!target.matches(".status-select")) return;

  const index = Number.parseInt(target.dataset.orderIndex || "-1", 10);
  if (!Number.isInteger(index) || !orders[index]) return;

  orders[index].status = target.value;
  persistOrders();
  renderOrdersTable();
});

orderForm.addEventListener("submit", (event) => {
  event.preventDefault();
  const order = buildOrder(new FormData(orderForm));
  orders.unshift(order);
  persistOrders();
  renderOrdersTable();
  renderGeneratedOrder(order);
  orderForm.reset();
});

printOrderBtn.addEventListener("click", () => window.print());

clearOrdersBtn.addEventListener("click", () => {
  const shouldClear = window.confirm(
    "¿Seguro que deseas borrar todo el historial de órdenes?"
  );
  if (!shouldClear) return;

  orders.length = 0;
  persistOrders();
  renderOrdersTable();
});

exportExcelBtn.addEventListener("click", () => {
  if (!orders.length) {
    window.alert("No hay órdenes para exportar.");
    return;
  }

  if (!window.XLSX) {
    downloadCsvFallback();
    return;
  }

  const excelRows = orders.map((order) => ({
    Orden: order.orderNumber,
    Fecha: order.createdAt,
    Cliente: order.customerName,
    Telefono: order.customerPhone,
    Equipo: order.deviceModel,
    IMEI: order.deviceImei,
    Estado: order.status,
    Falla: order.reportedIssue,
    Diagnostico: order.diagnosis,
    Costo: order.estimatedCost,
    Anticipo: order.advancePayment,
    Observaciones: order.notes,
  }));

  const workbook = XLSX.utils.book_new();
  const worksheet = XLSX.utils.json_to_sheet(excelRows);
  XLSX.utils.book_append_sheet(workbook, worksheet, "OrdenesTomFix");
  XLSX.writeFile(workbook, `ordenes-tomfix-${todayToken()}.xlsx`);
});

function buildOrder(formData) {
  const now = new Date();
  const values = Object.fromEntries(formData.entries());

  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, "0");
  const day = String(now.getDate()).padStart(2, "0");
  const todayKey = `${year}${month}${day}`;

  let orderData = JSON.parse(localStorage.getItem("orderCounter")) || {};

  if (!orderData.date || orderData.date !== todayKey) {
    orderData = {
      date: todayKey,
      counter: 1,
    };
  } else {
    orderData.counter++;
  }

  localStorage.setItem("orderCounter", JSON.stringify(orderData));

  const consecutive = String(orderData.counter).padStart(2, "0");

  return {
    orderNumber: `TF-${todayKey}-${consecutive}`,
    createdAt: now.toLocaleString("es-CO"),

    customerName: values.customerName?.trim() || "",
    customerPhone: values.customerPhone?.trim() || "",
    deviceModel: values.deviceModel?.trim() || "",
    deviceImei: values.deviceImei?.trim() || "",

    status: "Recibido",
    reportedIssue: values.reportedIssue?.trim() || "",
    diagnosis: values.diagnosis?.trim() || "",

    estimatedCost: numberOrZero(values.estimatedCost),
    advancePayment: numberOrZero(values.advancePayment),
    notes: values.notes?.trim() || "",
  };
}

function renderGeneratedOrder(order) {
  generatedOrderSection.hidden = false;
  generatedOrder.innerHTML = `
    <div class="order-ticket">
      <div class="ticket-top">
        <div class="ticket-brand">
          <img src="logo.png" alt="Logo de TomFix" />
          <div>
            <h3>${COMPANY_INFO.name}</h3>
            <p>${COMPANY_INFO.slogan}</p>
          </div>
        </div>
        <div class="ticket-meta">
          <span><strong>Orden:</strong> ${order.orderNumber}</span>
          <span><strong>Fecha:</strong> ${order.createdAt}</span>
          <span class="status-chip">${order.status}</span>
        </div>
      </div>

      <div class="ticket-grid">
        <section>
          <h4>Datos del cliente</h4>
          <dl>
            <dt>Cliente:</dt><dd>${order.customerName}</dd>
            <dt>Teléfono:</dt><dd>${order.customerPhone}</dd>
          </dl>
        </section>
        <section>
          <h4>Datos del equipo</h4>
          <dl>
            <dt>Modelo:</dt><dd>${order.deviceModel}</dd>
            <dt>IMEI / Serie:</dt><dd>${order.deviceImei || "N/A"}</dd>
          </dl>
        </section>
      </div>

      <section class="ticket-details">
        <h4>Detalle del servicio</h4>
        <dl>
          <dt>Falla reportada:</dt><dd>${order.reportedIssue}</dd>
          <dt>Diagnóstico:</dt><dd>${order.diagnosis || "Pendiente"}</dd>
          <dt>Costo:</dt><dd>$${order.estimatedCost.toLocaleString("es-CO")}</dd>
          <dt>Anticipo:</dt><dd>$${order.advancePayment.toLocaleString("es-CO")}</dd>
          <dt>Observaciones:</dt><dd>${order.notes || "Ninguna"}</dd>
        </dl>
      </section>

      <footer class="ticket-footer">
        <p><strong>WhatsApp ${COMPANY_INFO.whatsapp} · ${COMPANY_INFO.email}</p>
        <p>Gracias por confiar en ${COMPANY_INFO.name}. Conserva esta orden para la entrega de tu equipo.</p>
      </footer>
    </div>
  `;
}

function renderOrdersTable() {
  if (!orders.length) {
    ordersTableBody.innerHTML =
      '<tr><td colspan="9">Aún no hay órdenes registradas.</td></tr>';
    return;
  }

  ordersTableBody.innerHTML = orders
    .map(
      (order, index) => `
      <tr>
        <td>${order.orderNumber}</td>
        <td>${order.customerName}</td>
        <td>${order.customerPhone}</td>
        <td>${order.deviceModel}</td>
        <td>${order.reportedIssue}</td>
        <td>${order.status}</td>
        <td>$${order.estimatedCost.toFixed(2)}</td>
        <td>${order.createdAt}</td>
        <td>
          <select class="status-select" data-order-index="${index}" aria-label="Actualizar estado">
            ${statusOptions(order.status)}
          </select>
        </td>
      </tr>
    `
    )
    .join("");
}

function persistOrders() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(orders));
}

function loadOrders() {
  const raw = localStorage.getItem(STORAGE_KEY);
  if (!raw) return [];

  try {
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

function numberOrZero(value) {
  const number = Number.parseFloat(value);
  return Number.isFinite(number) ? number : 0;
}

function todayToken() {
  const now = new Date();
  const month = String(now.getMonth() + 1).padStart(2, "0");
  const day = String(now.getDate()).padStart(2, "0");
  return `${now.getFullYear()}-${month}-${day}`;
}

function downloadCsvFallback() {
  const headers = [
    "Orden",
    "Fecha",
    "Cliente",
    "Telefono",
    "Equipo",
    "IMEI",
    "Estado",
    "Falla",
    "Diagnostico",
    "Costo",
    "Anticipo",
    "Observaciones",
  ];

  const rows = orders.map((order) => [
    order.orderNumber,
    order.createdAt,
    order.customerName,
    order.customerPhone,
    order.deviceModel,
    order.deviceImei,
    order.status,
    order.reportedIssue,
    order.diagnosis,
    order.estimatedCost,
    order.advancePayment,
    order.notes,
  ]);

  const csvContent = [headers, ...rows]
    .map((row) => row.map((item) => `"${String(item).replaceAll('"', '""')}"`).join(","))
    .join("\n");

  const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `ordenes-tomfix-${todayToken()}.csv`;
  link.click();
  URL.revokeObjectURL(url);
}

function statusOptions(selectedStatus) {
  const statuses = [
    "Recibido",
    "En reparación",
    "Listo para entrega",
    "Entregado",
  ];

  return statuses
    .map(
      (status) =>
        `<option value="${status}" ${status === selectedStatus ? "selected" : ""}>${status}</option>`
    )
    .join("");
}
