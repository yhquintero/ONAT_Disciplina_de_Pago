import { useEffect, useMemo, useRef, useState } from 'react';
import { version as appVersion } from '../package.json';
import {
  AlertTriangle,
  ArrowRight,
  BarChart3,
  CalendarCheck,
  CalendarDays,
  Check,
  CheckCircle2,
  ChevronDown,
  ChevronRight,
  CircleDollarSign,
  Clock3,
  Download,
  FileDown,
  FileOutput,
  FileSpreadsheet,
  FileText,
  Info,
  LayoutDashboard,
  LockKeyhole,
  Printer,
  RotateCcw,
  Search,
  Settings2,
  SlidersHorizontal,
  Sparkles,
  Upload,
  Users,
  Wallet,
  X,
} from 'lucide-react';
import { Bar, Doughnut, Line } from 'react-chartjs-2';
import {
  ArcElement,
  BarElement,
  CategoryScale,
  Chart as ChartJS,
  Filler,
  Legend,
  LinearScale,
  LineElement,
  PointElement,
  Tooltip,
} from 'chart.js';
import {
  bestPeriod,
  buildDashboardModel,
  buildPivot,
  formatDate,
  getPaymentsForPeriod,
  monthLabel,
  uniqueTaxOptions,
} from './analytics.js';
import { normalizeText, parseWorkbook } from './data.js';
import { formatCompactMoney, formatMoney } from './formatting.js';
import { createFilename, createPdfBlob, createXlsxBlob } from './exporters.js';

ChartJS.register(CategoryScale, LinearScale, PointElement, LineElement, BarElement, ArcElement, Tooltip, Legend, Filler);

const integerFormat = new Intl.NumberFormat('es-ES', { maximumFractionDigits: 0 });

const money = formatMoney;
const compactMoney = formatCompactMoney;
const countLabel = (value) => integerFormat.format(Number(value) || 0);

const navItems = [
  { id: 'overview', label: 'Resumen', icon: LayoutDashboard },
  { id: 'control', label: 'Panel de control', icon: SlidersHorizontal },
  { id: 'charts', label: 'Gráficos dinámicos', icon: BarChart3 },
  { id: 'clients', label: 'Contribuyentes', icon: Users },
];

const APP_VERSION = appVersion;
const DEMO_DATA_ENABLED = import.meta.env.DEV;
const DEMO_CLIENTS_URL = '/.xlsx/Cliente%20-%20obligaci%C3%B3n%20(vector%20fiscal).xlsx';
const DEMO_PAYMENTS_URL = '/.xlsx/Pago1%20(persona%20jur%C3%ADdica).xlsx';

const dimensionLabels = {
  dpa: 'Dirección provincial (DPA)',
  tributo: 'Tributo',
  formaorganizativa: 'Forma organizativa',
  organismodireccion: 'Organismo de dirección',
};

function safeStorageRead(key, fallback) {
  try {
    const value = window.localStorage.getItem(key);
    return value == null ? fallback : JSON.parse(value);
  } catch {
    return fallback;
  }
}

function safeStorageWrite(key, value) {
  try {
    window.localStorage.setItem(key, JSON.stringify(value));
  } catch {
    // Preferences are optional; the panel still works if browser storage is disabled.
  }
}

function clientDateToday() {
  const now = new Date();
  return new Date(Date.UTC(now.getFullYear(), now.getMonth(), now.getDate()));
}

function statusLabel(status) {
  return {
    paid: 'En fecha',
    late: 'Fuera de fecha',
    unpaid: 'Vencido sin pago',
    pending: 'Pendiente',
  }[status] || status;
}

function statusTone(status) {
  return {
    paid: 'success',
    late: 'warning',
    unpaid: 'danger',
    pending: 'neutral',
  }[status] || 'neutral';
}

function formatHeaderName(value) {
  const key = normalizeText(value).replace(/\s+/g, '');
  return dimensionLabels[key] || value;
}

function fileSize(bytes) {
  if (!bytes) return '';
  return bytes > 1024 * 1024
    ? `${(bytes / (1024 * 1024)).toFixed(1)} MB`
    : `${Math.max(1, Math.round(bytes / 1024))} KB`;
}

function chartCurrencyOptions(horizontal = false) {
  return {
    responsive: true,
    maintainAspectRatio: false,
    interaction: { intersect: false, mode: 'index' },
    plugins: {
      legend: { display: false },
      tooltip: {
        backgroundColor: '#14283a',
        padding: 11,
        titleFont: { family: 'Inter, sans-serif', size: 12, weight: '600' },
        bodyFont: { family: 'Inter, sans-serif', size: 12 },
        callbacks: {
          label: (context) => ` ${money(context.parsed.y ?? context.parsed.x ?? context.raw)}`,
        },
      },
    },
    scales: {
      x: {
        grid: { display: false },
        border: { display: false },
        ticks: { color: '#8492a4', maxRotation: 0, autoSkip: true, font: { size: 11 } },
      },
      y: {
        beginAtZero: true,
        grid: { color: 'rgba(130, 149, 168, .15)', drawTicks: false },
        border: { display: false, dash: [3, 5] },
        ticks: {
          color: '#8492a4',
          padding: 10,
          maxTicksLimit: 5,
          callback: (value) => compactMoney(value),
          font: { size: 11 },
        },
      },
      ...(horizontal ? {
        x: {
          beginAtZero: true,
          grid: { color: 'rgba(130, 149, 168, .15)', drawTicks: false },
          border: { display: false, dash: [3, 5] },
          ticks: { color: '#8492a4', padding: 8, callback: (value) => compactMoney(value), font: { size: 11 } },
        },
        y: {
          grid: { display: false },
          border: { display: false },
          ticks: { color: '#718096', autoSkip: false, font: { size: 11 } },
        },
      } : {}),
    },
  };
}

function chartCountOptions(horizontal = false) {
  const options = chartCurrencyOptions(horizontal);
  options.plugins.tooltip.callbacks.label = (context) => ` ${countLabel(context.parsed.y ?? context.parsed.x ?? context.raw)}`;
  options.scales.y.ticks.callback = (value) => countLabel(value);
  if (horizontal) options.scales.x.ticks.callback = (value) => countLabel(value);
  return options;
}

function Toast({ toast, onClose }) {
  useEffect(() => {
    if (!toast) return undefined;
    const timer = window.setTimeout(onClose, 4300);
    return () => window.clearTimeout(timer);
  }, [toast, onClose]);

  if (!toast) return null;
  const Icon = toast.type === 'error' ? AlertTriangle : CheckCircle2;
  return (
    <div className={`toast toast-${toast.type || 'success'}`} role="status">
      <Icon size={18} />
      <span>{toast.message}</span>
      <button type="button" aria-label="Cerrar aviso" onClick={onClose}><X size={16} /></button>
    </div>
  );
}

function Sidebar({ activeView, onNavigate, model, loading }) {
  return (
    <aside className="sidebar">
      <div className="brand-lockup">
        <div className="brand-symbol"><LandmarkMark /></div>
        <div className="brand-copy">
          <span className="brand-name">ONAT</span>
          <span className="brand-caption">CONTROL DE PAGO</span>
        </div>
      </div>

      <div className="sidebar-section-label">ESPACIO DE TRABAJO</div>
      <nav className="sidebar-nav" aria-label="Navegación principal">
        {navItems.map(({ id, label, icon: Icon }) => (
          <button
            type="button"
            className={`nav-item ${activeView === id ? 'active' : ''}`}
            key={id}
            onClick={() => onNavigate(id)}
            aria-current={activeView === id ? 'page' : undefined}
            aria-label={label}
            title={label}
          >
            <Icon size={19} strokeWidth={1.8} />
            <span>{label}</span>
            {id === 'clients' && !loading && <span className="nav-count">{countLabel(model.expectedCount)}</span>}
          </button>
        ))}
      </nav>

      <div className="sidebar-spacer" />
      <div className="privacy-card">
        <div className="privacy-icon"><LockKeyhole size={15} /></div>
        <div className="privacy-copy">
          <strong>Datos en tu equipo</strong>
          <span>Los Excel se procesan aquí</span>
        </div>
        <span className="privacy-dot" />
      </div>
      <div className="sidebar-footer">
        <span className="sidebar-version">Versión {APP_VERSION}</span>
      </div>
    </aside>
  );
}

function LandmarkMark() {
  return (
    <svg viewBox="0 0 36 36" aria-hidden="true" className="brand-mark">
      <path d="M5.5 14.5 18 7l12.5 7.5" />
      <path d="M8 15h20M10 27h16M7 30h22" />
      <path d="M11 16v9M16 16v9M20 16v9M25 16v9" />
    </svg>
  );
}

function TopBar({ onImport, onExport, exportDisabled, showExport, showImport = true, currentTitle }) {
  return (
    <div className="topbar">
      <div className="topbar-breadcrumb">
        <span className="breadcrumb-root">Panel local</span>
        <ChevronRight size={14} />
        <span className="breadcrumb-current">{currentTitle}</span>
      </div>
      <div className="topbar-actions">
        <div className="private-status"><span className="status-led" /> Procesamiento local</div>
        {showExport && <button type="button" className="secondary-button top-export" onClick={onExport} disabled={exportDisabled}><Download size={15} /><span>Exportar</span></button>}
        {showImport && <button type="button" className="primary-button top-import" onClick={onImport}>
          <Upload size={16} />
          <span>Importar Excel</span>
        </button>}
      </div>
    </div>
  );
}

function PageHeading({ title, subtitle, period, setPeriod, periods, taxFilter, setTaxFilter, taxOptions, onSettings, loading }) {
  return (
    <div className="page-heading-row">
      <div className="page-heading-copy">
        <div className="eyebrow"><span className="eyebrow-line" /> SEGUIMIENTO TRIBUTARIO</div>
        <h1>{title}</h1>
        <p>{subtitle}</p>
      </div>
      <div className="heading-controls">
        <label className="select-control" htmlFor="period-filter">
          <CalendarDays size={17} />
          <span className="select-control-label">Período</span>
          <select id="period-filter" value={period} onChange={(event) => setPeriod(event.target.value)} disabled={!periods.length || loading}>
            {!periods.length && <option value="">Sin períodos</option>}
            {[...periods].sort((a, b) => b.localeCompare(a)).map((item) => (
              <option key={item} value={item}>{monthLabel(item)}</option>
            ))}
          </select>
          <ChevronDown size={15} className="select-chevron" />
        </label>
        <label className="select-control tax-select" htmlFor="tax-filter">
          <span className="select-control-label">Tributo</span>
          <select id="tax-filter" value={taxFilter} onChange={(event) => setTaxFilter(event.target.value)} disabled={loading}>
            <option value="">Todos</option>
            {taxOptions.map((tax) => <option key={tax.key} value={tax.key}>{tax.label}</option>)}
          </select>
          <ChevronDown size={15} className="select-chevron" />
        </label>
        <button type="button" className="rule-button" onClick={onSettings} title="Editar plazo de pago">
          <SlidersHorizontal size={16} />
          <span>Regla</span>
        </button>
      </div>
    </div>
  );
}

function DeadlineStrip({ model, period, workingDays, holidays, loading }) {
  if (loading) {
    return <div className="deadline-strip deadline-loading"><span className="loading-dot" /> Preparando los archivos de ejemplo…</div>;
  }
  if (!period) {
    return <div className="deadline-strip deadline-empty"><Info size={17} /> Carga un Excel de pagos para calcular el plazo del período.</div>;
  }

  const date = model.deadline;
  return (
    <div className="deadline-strip">
      <div className="deadline-icon"><CalendarCheck size={19} /></div>
      <div className="deadline-copy">
        <span className="deadline-kicker">PLAZO DE PAGO · {monthLabel(period).toLocaleUpperCase('es-CU')}</span>
        <strong>{date ? `Hasta el ${formatDate(date)}` : 'No se pudo calcular el vencimiento'}</strong>
      </div>
      <span className="deadline-divider" />
      <div className="deadline-rule-text">
        <span>Primeros <b>{workingDays} días hábiles</b></span>
        <small>{holidays.length ? `Se excluyen fines de semana y ${holidays.length} feriado${holidays.length === 1 ? '' : 's'} configurado${holidays.length === 1 ? '' : 's'}` : 'Se excluyen sábados y domingos'}</small>
      </div>
      <div className={`deadline-state ${model.overdue ? 'is-overdue' : ''}`}>
        {model.overdue ? <Clock3 size={14} /> : <CheckCircle2 size={14} />}
        {model.overdue ? 'Período vencido' : 'Plazo vigente'}
      </div>
    </div>
  );
}

function PeriodBasisNotice({ payments }) {
  const datedRows = payments.filter((payment) => payment.dateBasis === 'payment').length;
  if (!datedRows) return null;
  const mixed = datedRows < payments.length;
  return (
    <Notice icon={Info} tone="info" title={mixed ? 'Algunas filas no tienen columna Período' : 'El Excel no incluye la columna Período'}>
      {mixed
        ? `En ${countLabel(datedRows)} pagos sin Período se usa el mes de Fecha de pago para el análisis.`
        : 'Se usa el mes de Fecha de pago como período de obligación. Si el pago puede registrarse en un mes distinto al que corresponde, importa también la columna Período para clasificar la puntualidad con mayor precisión.'}
    </Notice>
  );
}

function Notice({ icon: Icon = Info, tone = 'info', title, children, onClose }) {
  return (
    <div className={`notice notice-${tone}`}>
      <div className="notice-icon"><Icon size={17} /></div>
      <div className="notice-body">
        {title && <strong>{title}</strong>}
        <span>{children}</span>
      </div>
      {onClose && <button type="button" className="notice-close" onClick={onClose} aria-label="Cerrar aviso"><X size={15} /></button>}
    </div>
  );
}

function StatCard({ icon: Icon, label, value, description, tone = 'blue', suffix, loading }) {
  return (
    <div className={`stat-card stat-${tone}`}>
      <div className="stat-topline">
        <div className="stat-icon"><Icon size={18} strokeWidth={1.9} /></div>
        <span className="stat-label">{label}</span>
        {suffix && <span className="stat-suffix">{suffix}</span>}
      </div>
      <div className={`stat-value ${loading ? 'stat-skeleton' : ''}`}>{loading ? '—' : value}</div>
      <div className="stat-description">{description}</div>
      <span className="stat-accent" />
    </div>
  );
}

function StatsGrid({ model, loading }) {
  const noPayment = model.unpaidCount + model.pendingCount;
  return (
    <div className="stats-grid">
      <StatCard
        icon={Wallet}
        label="Importe cobrado"
        value={money(model.totalAmount)}
        description={`${countLabel(model.transactions)} pagos registrados en el período`}
        tone="blue"
        loading={loading}
      />
      <StatCard
        icon={Users}
        label="Contribuyentes con pago"
        value={countLabel(model.paidCount)}
        description={`de ${countLabel(model.expectedCount)} en el Patrón`}
        tone="teal"
        loading={loading}
      />
      <StatCard
        icon={Clock3}
        label="Fuera de fecha"
        value={countLabel(model.lateCount)}
        description={`${money(model.lateAmount)} recibidos después del plazo`}
        tone="amber"
        loading={loading}
      />
      <StatCard
        icon={CircleDollarSign}
        label="Sin pago registrado"
        value={countLabel(noPayment)}
        description={`${countLabel(model.unpaidCount)} vencidos · ${countLabel(model.pendingCount)} pendientes`}
        tone="rose"
        loading={loading}
      />
    </div>
  );
}

function CardHeader({ eyebrow, title, subtitle, action }) {
  return (
    <div className="card-header">
      <div>
        {eyebrow && <div className="card-eyebrow">{eyebrow}</div>}
        <h2>{title}</h2>
        {subtitle && <p>{subtitle}</p>}
      </div>
      {action && <div className="card-header-action">{action}</div>}
    </div>
  );
}

function EmptyChart({ message = 'No hay pagos para mostrar en esta selección.' }) {
  return (
    <div className="empty-chart">
      <div className="empty-chart-icon"><BarChart3 size={21} /></div>
      <span>{message}</span>
    </div>
  );
}

function TrendChart({ model, period }) {
  const entries = model.trend.slice(-8);
  if (!entries.length || entries.every((item) => item.amount === 0)) return <EmptyChart />;
  const currentIndex = entries.findIndex((item) => item.key === period);
  return (
    <Line
      data={{
        labels: entries.map((item) => item.label),
        datasets: [{
          data: entries.map((item) => item.amount),
          borderColor: '#2876ce',
          backgroundColor: 'rgba(40, 118, 206, .11)',
          pointBackgroundColor: entries.map((item) => item.key === period ? '#1a8e87' : '#ffffff'),
          pointBorderColor: entries.map((item) => item.key === period ? '#1a8e87' : '#2876ce'),
          pointBorderWidth: 2,
          pointRadius: entries.map((item) => item.key === period ? 5 : 3),
          pointHoverRadius: 6,
          borderWidth: 2.4,
          tension: 0.34,
          fill: true,
        }],
      }}
      options={{
        ...chartCurrencyOptions(),
        plugins: {
          ...chartCurrencyOptions().plugins,
          tooltip: {
            ...chartCurrencyOptions().plugins.tooltip,
            callbacks: {
              title: (items) => entries[items[0]?.dataIndex]?.key ? monthLabel(entries[items[0].dataIndex].key) : '',
              label: (context) => ` ${money(context.parsed.y)}`,
            },
          },
        },
        scales: {
          ...chartCurrencyOptions().scales,
          x: { ...chartCurrencyOptions().scales.x, ticks: { ...chartCurrencyOptions().scales.x.ticks, autoSkip: false } },
        },
      }}
    />
  );
}

function StatusDonut({ model, valueType = 'amount' }) {
  const data = valueType === 'clients'
    ? [model.statusCounts.paid, model.statusCounts.late, model.statusCounts.unpaid + model.statusCounts.pending]
    : [model.onTimeAmount, model.lateAmount];
  if (!data.some((value) => value > 0)) return <EmptyChart message="No se registraron cobros en este período." />;
  const labels = valueType === 'clients'
    ? ['En fecha', 'Fuera de fecha', 'Sin pago']
    : ['En fecha', 'Fuera de fecha'];
  return (
    <div className="donut-layout">
      <div className="donut-canvas">
        <Doughnut
          data={{
            labels,
            datasets: [{
              data,
              backgroundColor: valueType === 'clients' ? ['#35aa94', '#f3ad57', '#e8edf3'] : ['#35aa94', '#f3ad57'],
              borderWidth: 0,
              hoverOffset: 5,
              borderRadius: 4,
              spacing: 2,
            }],
          }}
          options={{
            responsive: true,
            maintainAspectRatio: false,
            cutout: '74%',
            plugins: {
              legend: { display: false },
              tooltip: {
                backgroundColor: '#14283a',
                padding: 11,
                callbacks: {
                  label: (context) => ` ${valueType === 'clients' ? countLabel(context.raw) : money(context.raw)}`,
                },
              },
            },
          }}
        />
        <div className="donut-center">
          <strong>{valueType === 'clients' ? countLabel(model.expectedCount) : compactMoney(model.totalAmount)}</strong>
          <span>{valueType === 'clients' ? 'en Patrón' : 'cobrado'}</span>
        </div>
      </div>
      <div className="donut-legend">
        {labels.map((label, index) => (
          <div className="donut-legend-row" key={label}>
            <span className={`legend-dot legend-${index}`} />
            <span className="legend-name">{label}</span>
            <strong>{valueType === 'clients' ? countLabel(data[index]) : money(data[index])}</strong>
          </div>
        ))}
      </div>
    </div>
  );
}

function TopClientsChart({ model }) {
  const entries = model.topClients.slice(0, 6).reverse();
  if (!entries.length || entries.every((item) => item.amount === 0)) return <EmptyChart message="Todavía no hay importes para comparar." />;
  const options = chartCurrencyOptions(true);
  options.plugins.tooltip.callbacks.title = (items) => entries[items[0]?.dataIndex]?.name || '';
  options.scales.y.ticks.callback = function callback(value) {
    const label = this.getLabelForValue(value);
    return label.length > 22 ? `${label.slice(0, 20)}…` : label;
  };
  return (
    <Bar
      data={{
        labels: entries.map((item) => item.name),
        datasets: [{
          data: entries.map((item) => item.amount),
          backgroundColor: entries.map((item, index) => index === entries.length - 1 ? '#2982d4' : 'rgba(41, 130, 212, .62)'),
          hoverBackgroundColor: '#1d6eb8',
          borderRadius: 5,
          barThickness: 13,
        }],
      }}
      options={{ ...options, plugins: { ...options.plugins, tooltip: { ...options.plugins.tooltip, callbacks: { ...options.plugins.tooltip.callbacks, label: (context) => ` ${money(context.parsed.x)}` } } } }}
    />
  );
}

function PaymentMonthChart({ model }) {
  const entries = model.byPaymentMonth;
  if (!entries.length || entries.every((item) => item.amount === 0)) return <EmptyChart message="No hay fechas de cobro para este período." />;
  const options = chartCurrencyOptions();
  options.plugins.tooltip.callbacks.title = (items) => entries[items[0]?.dataIndex]?.key ? monthLabel(entries[items[0].dataIndex].key) : '';
  return (
    <Bar
      data={{
        labels: entries.map((item) => item.label),
        datasets: [{
          data: entries.map((item) => item.amount),
          backgroundColor: entries.map((item) => {
            const [year, month] = item.key.split('-').map(Number);
            const paymentDate = new Date(Date.UTC(year, month - 1, 1));
            const activeMonth = model.startDate?.getUTCMonth() === paymentDate.getUTCMonth() && model.startDate?.getUTCFullYear() === paymentDate.getUTCFullYear();
            return activeMonth ? '#1a9b8d' : '#84c7bf';
          }),
          borderRadius: 5,
          maxBarThickness: 36,
        }],
      }}
      options={options}
    />
  );
}

function PreviewTable({ model, onViewAll }) {
  const rows = model.rosterRows.slice(0, 6);
  return (
    <section className="surface-card preview-table-card">
      <CardHeader
        eyebrow="SEGUIMIENTO DEL PATRÓN"
        title="Últimos estados"
        subtitle={model.expectedCount ? `${countLabel(model.expectedCount)} contribuyentes en el período` : 'Sin registros de Patrón'}
        action={<button type="button" className="text-button" onClick={onViewAll}>Ver todos <ArrowRight size={15} /></button>}
      />
      {rows.length ? (
        <div className="table-scroll preview-scroll">
          <table className="data-table preview-table">
            <thead>
              <tr><th>Contribuyente</th><th>Estado</th><th className="align-right">Importe ($)</th></tr>
            </thead>
            <tbody>
              {rows.map((row) => (
                <tr key={row.key}>
                  <td>
                    <div className="contributor-cell">
                      <span className={`client-avatar avatar-${statusTone(row.status)}`}>{row.contributor.charAt(0).toUpperCase()}</span>
                      <span className="client-name-wrap"><strong title={row.contributor}>{row.contributor}</strong><small>{row.tax || 'Contribuyente'}</small></span>
                    </div>
                  </td>
                  <td><StatusBadge status={row.status} /></td>
                  <td className="align-right amount-cell">{row.paymentCount ? money(row.amount) : '—'}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      ) : <EmptyTable onReview={onViewAll} />}
    </section>
  );
}

function EmptyTable({ onReview }) {
  return (
    <div className="empty-table-state">
      <div className="empty-chart-icon"><Users size={20} /></div>
      <strong>No hay contribuyentes para este filtro</strong>
      <span>Revisa el tributo seleccionado o importa un Patrón de clientes.</span>
      <button type="button" className="text-button" onClick={onReview}>Revisar datos <ArrowRight size={15} /></button>
    </div>
  );
}

function StatusBadge({ status }) {
  const Icon = status === 'paid' ? Check : status === 'late' ? Clock3 : status === 'unpaid' ? AlertTriangle : Clock3;
  return (
    <span className={`status-badge badge-${statusTone(status)}`}>
      <Icon size={13} strokeWidth={2.2} />
      {statusLabel(status)}
    </span>
  );
}

function GetStartedCard({ data, onImport }) {
  return (
    <section className="surface-card welcome-card">
      <div className="welcome-symbol"><FileSpreadsheet size={23} /></div>
      <div className="welcome-kicker">CONFIGURACIÓN INICIAL</div>
      <h2>Importa tus Excel para comenzar</h2>
      <p>Sube el Patrón completo y el registro de pagos. El panel los cruzará en este navegador para mostrar cobros, atrasos y clientes sin pago registrado.</p>
      <div className="welcome-steps">
        <div className={`welcome-step ${data.clientMeta ? 'step-ready' : ''}`}>
          <span className="welcome-step-number">{data.clientMeta ? <Check size={14} /> : '01'}</span>
          <span className="welcome-step-copy"><strong>Patrón de clientes</strong><small>{data.clientMeta ? `${countLabel(data.clientMeta.count)} contribuyentes cargados` : 'Columna requerida: Contribuyente'}</small></span>
          <span className="welcome-step-status">{data.clientMeta ? 'Listo' : 'Pendiente'}</span>
        </div>
        <div className={`welcome-step ${data.paymentMeta ? 'step-ready' : ''}`}>
          <span className="welcome-step-number">{data.paymentMeta ? <Check size={14} /> : '02'}</span>
          <span className="welcome-step-copy"><strong>Registro de pagos</strong><small>{data.paymentMeta ? `${countLabel(data.paymentMeta.count)} pagos cargados` : 'Contribuyente, Importe ($) y Fecha de pago'}</small></span>
          <span className="welcome-step-status">{data.paymentMeta ? 'Listo' : 'Pendiente'}</span>
        </div>
      </div>
      <button type="button" className="primary-button welcome-import" onClick={onImport}><Upload size={16} /> Importar archivos Excel</button>
      <div className="welcome-privacy"><LockKeyhole size={14} /> Los datos se procesan localmente y no se envían a un servidor.</div>
    </section>
  );
}

function ControlSourceCard({ role, title, description, meta, fields = [] }) {
  const isClients = role === 'clients';
  const Icon = isClients ? Users : FileSpreadsheet;
  return (
    <article className={`control-source-card ${meta ? 'source-ready' : ''}`}>
      <div className="control-source-top">
        <span className={`control-source-icon ${isClients ? 'source-icon-violet' : 'source-icon-green'}`}><Icon size={17} /></span>
        <div className="control-source-copy">
          <strong>{title}</strong>
          <small>{meta ? (meta.sample ? 'Libro de muestra · desarrollo' : 'Procesado en este navegador') : description}</small>
        </div>
        <span className={`control-source-status ${meta ? 'status-ready' : 'status-pending'}`}>{meta ? 'Cargado' : 'Pendiente'}</span>
      </div>
      {meta ? (
        <>
          <div className="control-source-file">
            <FileText size={14} />
            <span title={meta.fileName}>{meta.fileName}</span>
          </div>
          <div className="control-source-stats">
            <span><strong>{countLabel(meta.count)}</strong> registros</span>
            <span><strong>{countLabel(fields.length)}</strong> columnas</span>
            {meta.size > 0 && <span>{fileSize(meta.size)}</span>}
          </div>
          {fields.length > 0 && (
            <div className="control-source-fields" aria-label={`Columnas detectadas para ${title}`}>
              {fields.slice(0, 5).map((field, index) => <span key={`${field.key}-${index}`}>{field.label}</span>)}
              {fields.length > 5 && <span className="field-overflow">+{countLabel(fields.length - 5)}</span>}
            </div>
          )}
          {meta.skippedRows > 0 && <div className="control-source-warning"><AlertTriangle size={13} /> {countLabel(meta.skippedRows)} filas incompletas omitidas</div>}
        </>
      ) : (
        <div className="control-source-empty">
          <span>{description}. Usa “Importar Excel” en la barra superior para cargarlo.</span>
        </div>
      )}
    </article>
  );
}

function ControlPanelView({ model, period, workingDays, holidays, data }) {
  const periodBasedPayments = data.payments.filter((payment) => payment.dateBasis === 'period').length;
  const paymentDateFallbacks = data.payments.length - periodBasedPayments;
  const skippedRows = (data.clientMeta?.skippedRows || 0) + (data.paymentMeta?.skippedRows || 0);
  const fieldCount = data.clientFields.length + data.paymentFields.length;
  const qualityChecks = [
    {
      key: 'clients',
      label: 'Patrón de clientes',
      detail: data.clients.length ? `${countLabel(data.clients.length)} contribuyentes disponibles` : 'Importa el listado completo para identificar quién no pagó',
      status: data.clients.length ? 'good' : 'pending',
    },
    {
      key: 'payments',
      label: 'Registro de pagos',
      detail: data.payments.length ? `${countLabel(data.payments.length)} movimientos leídos` : 'Importa un Excel con importe y fecha de pago',
      status: data.payments.length ? 'good' : 'pending',
    },
    {
      key: 'period',
      label: 'Período de obligación',
      detail: !data.payments.length
        ? 'Pendiente de revisar'
        : paymentDateFallbacks
          ? `${countLabel(paymentDateFallbacks)} pagos usan Fecha de pago como período`
          : `${countLabel(periodBasedPayments)} pagos tienen Período explícito`,
      status: !data.payments.length ? 'pending' : paymentDateFallbacks ? 'warning' : 'good',
    },
    {
      key: 'columns',
      label: 'Columnas detectadas',
      detail: fieldCount ? `${countLabel(fieldCount)} campos originales conservados` : 'Las columnas aparecerán al importar los libros',
      status: fieldCount ? 'good' : 'pending',
    },
    {
      key: 'rows',
      label: 'Filas incompletas',
      detail: skippedRows ? `${countLabel(skippedRows)} filas omitidas durante la lectura` : data.payments.length ? 'No se omitieron filas incompletas' : 'Se comprobará al importar',
      status: skippedRows ? 'warning' : data.payments.length ? 'good' : 'pending',
    },
  ];
  const readyChecks = qualityChecks.filter((check) => check.status === 'good').length;
  const completion = Math.round((readyChecks / qualityChecks.length) * 100);
  const reviewCount = model.lateCount + model.unpaidCount + model.pendingCount;
  const punctuality = model.paidCount ? Math.round(model.punctuality) : null;

  return (
    <div className="control-view">
      <section className="control-hero">
        <div className="control-hero-copy">
          <div className="control-hero-kicker"><span className="control-live-dot" /> CENTRO DE CONTROL · SOLO EN ESTE NAVEGADOR</div>
          <h2>Todo tu seguimiento,<br />en un solo lugar.</h2>
          <p>Administra los libros, revisa la calidad de los datos y ajusta la regla de puntualidad antes de consultar los resultados.</p>
        </div>
        <div className="control-score-panel">
          <span className="control-score-caption">PUNTUALIDAD DEL PERÍODO</span>
          <div
            className="control-score-ring"
            style={{ background: punctuality == null ? 'conic-gradient(rgba(255,255,255,.16) 0 100%)' : `conic-gradient(#72cbbb 0 ${punctuality}%, rgba(255,255,255,.16) ${punctuality}% 100%)` }}
          >
            <div className="control-score-inner"><strong>{punctuality == null ? '—' : `${punctuality}%`}</strong><small>en fecha</small></div>
          </div>
          <p>{model.paidCount ? `${countLabel(model.statusCounts.paid)} puntuales de ${countLabel(model.paidCount)} con pago` : 'Sin pagos registrados en este período'}</p>
        </div>
      </section>

      <div className="control-main-grid">
        <section className="surface-card control-card control-sources-panel">
          <div className="control-card-heading">
            <div><span className="card-eyebrow">LIBROS DE EXCEL</span><h3>Fuentes de información</h3><p>Archivos procesados localmente; nunca se envían a un servidor.</p></div>
            <span className="control-field-count"><FileSpreadsheet size={14} /> {countLabel(fieldCount)} campos</span>
          </div>
          <div className="control-source-stack">
            <ControlSourceCard
              role="clients"
              title="Patrón de clientes"
              description="Listado de contribuyentes esperados"
              meta={data.clientMeta}
              fields={data.clientFields}
            />
            <ControlSourceCard
              role="payments"
              title="Registro de pagos"
              description="Contribuyente, importe, fecha y atributos"
              meta={data.paymentMeta}
              fields={data.paymentFields}
            />
          </div>
        </section>

        <section className="surface-card control-card control-quality-panel">
          <div className="control-card-heading">
            <div><span className="card-eyebrow">REVISIÓN PREVIA</span><h3>Estado de preparación</h3><p>Una guía rápida antes de usar el análisis.</p></div>
            <span className={`control-completion-badge ${completion === 100 ? 'is-complete' : ''}`}>{completion === 100 ? 'Listo' : `${readyChecks}/${qualityChecks.length}`}</span>
          </div>
          <div className="control-progress-track" role="progressbar" aria-label="Controles de preparación completados" aria-valuemin={0} aria-valuemax={100} aria-valuenow={completion}><span style={{ width: `${completion}%` }} /></div>
          <div className="control-quality-list">
            {qualityChecks.map((check) => {
              const Icon = check.status === 'good' ? CheckCircle2 : check.status === 'warning' ? AlertTriangle : Clock3;
              return (
                <div className={`control-quality-row quality-${check.status}`} key={check.key}>
                  <span className="control-quality-icon"><Icon size={15} /></span>
                  <span className="control-quality-copy"><strong>{check.label}</strong><small>{check.detail}</small></span>
                  <span className="control-quality-state">{check.status === 'good' ? 'OK' : check.status === 'warning' ? 'Revisar' : 'Pendiente'}</span>
                </div>
              );
            })}
          </div>
          {paymentDateFallbacks > 0 && (
            <div className="control-quality-note"><Info size={14} /> Los períodos inferidos desde Fecha de pago pueden diferir del mes real de obligación.</div>
          )}
        </section>
      </div>

      <div className="control-bottom-grid">
        <section className="surface-card control-card control-focus-panel">
          <div className="control-card-heading">
            <div><span className="card-eyebrow">ATENCIÓN DEL PERÍODO</span><h3>{model.overdue ? 'Revisa los casos vencidos' : 'Casos a revisar'}</h3></div>
            <span className={`control-attention-badge ${reviewCount ? 'has-attention' : ''}`}>{countLabel(reviewCount)} casos</span>
          </div>
          <p className="control-focus-description">{!data.payments.length
            ? 'Importa los pagos para ver atrasos, pendientes y contribuyentes sin registro.'
            : reviewCount
              ? 'Consulta los estados de cada contribuyente para priorizar el seguimiento del período.'
              : 'No se detectan atrasos ni pagos pendientes en la selección actual.'}</p>
          <div className="control-focus-breakdown">
            <span><i className="focus-dot focus-late" /> Fuera de fecha <strong>{countLabel(model.lateCount)}</strong></span>
            <span><i className="focus-dot focus-unpaid" /> Vencidos sin pago <strong>{countLabel(model.unpaidCount)}</strong></span>
            <span><i className="focus-dot focus-pending" /> Pendientes <strong>{countLabel(model.pendingCount)}</strong></span>
          </div>
        </section>

        <section className="surface-card control-card control-rule-panel">
          <div className="control-card-heading">
            <div><span className="card-eyebrow">REGLA ACTIVA</span><h3>Plazo de puntualidad</h3></div>
            <span className={`control-rule-status ${model.overdue ? 'is-overdue' : ''}`}>{period ? (model.overdue ? 'Vencido' : 'Vigente') : 'Sin período'}</span>
          </div>
          <div className="control-rule-summary"><strong>{workingDays}</strong><span>días hábiles<br />desde el inicio del mes</span></div>
          <p>{model.deadline ? `Vencimiento calculado: ${formatDate(model.deadline)}.` : 'Selecciona un período para calcular su vencimiento.'} {holidays.length ? `${countLabel(holidays.length)} feriados personalizados.` : 'No hay feriados personalizados.'}</p>
        </section>
      </div>
    </div>
  );
}

function SummaryView({ model, period, workingDays, onViewAll, onDismissQuarterly }) {
  return (
    <>
      {model.missingRoster && (
        <Notice icon={Info} tone="info" title="Falta el Patrón de clientes">
          El Excel de pagos solo muestra a quienes pagaron. Usa “Importar Excel” para cargar un Patrón completo e identificar quién no tiene pago registrado.
        </Notice>
      )}
      {model.quarterlyWarning && (
        <Notice icon={AlertTriangle} tone="warning" title="El Patrón indica periodicidad trimestral" onClose={onDismissQuarterly}>
          Los estados se calculan con la regla mensual de {workingDays} días hábiles que configuraste. Comprueba que el período seleccionado corresponda a una obligación mensual antes de usar estos resultados como notificación oficial.
        </Notice>
      )}
      <div className="content-grid chart-grid-summary">
        <section className="surface-card chart-card main-trend-card">
          <CardHeader
            eyebrow="EVOLUCIÓN DE COBROS"
            title="Importe por período"
            subtitle="Suma de Importe ($) según el mes de obligación"
            action={<span className="chart-live-pill"><span /> Dinámico</span>}
          />
          <div className="chart-area trend-area"><TrendChart model={model} period={period} /></div>
          <div className="chart-footnote"><span className="current-period-dot" /> Período seleccionado</div>
        </section>
        <section className="surface-card chart-card donut-card">
          <CardHeader
            eyebrow="PUNTUALIDAD"
            title="En fecha vs. tarde"
            subtitle="Distribución del importe recaudado"
          />
          <div className="chart-area donut-area"><StatusDonut model={model} /></div>
        </section>
      </div>

      <div className="content-grid bottom-grid-summary">
        <PreviewTable model={model} onViewAll={onViewAll} />
        <section className="surface-card chart-card top-card">
          <CardHeader
            eyebrow="MAYORES COBROS"
            title="Por contribuyente"
            subtitle={period ? `Importe total · ${monthLabel(period)}` : 'Importe total'}
            action={<span className="mini-top-badge">TOP 6</span>}
          />
          <div className="chart-area top-chart-area"><TopClientsChart model={model} /></div>
        </section>
      </div>
    </>
  );
}

function PivotExplorer({
  pivot,
  options,
  chartRef,
  dimension,
  setDimension,
  metric,
  setMetric,
  scope,
  setScope,
  rosterOnly,
  setRosterOnly,
  chartType,
  setChartType,
  limit,
  setLimit,
  period,
}) {
  const isCount = metric !== 'amount';
  const isCategory = !['period', 'paymentMonth'].includes(dimension);
  const horizontal = chartType === 'bar' && isCategory;
  const hasData = pivot.values.some((value) => value > 0);
  const palette = ['#2e7bc7', '#249789', '#e9a24b', '#6276c4', '#89b7dd', '#58b8ad', '#e48666', '#8a91bd', '#5d9c88', '#bd8a48', '#6d9dc9', '#a7b8c9'];
  const valueLabel = (value) => isCount ? countLabel(value) : money(value);
  const barOptions = isCount ? chartCountOptions(horizontal) : chartCurrencyOptions(horizontal);
  barOptions.plugins.tooltip.callbacks.title = (items) => pivot.labels[items[0]?.dataIndex] || '';
  barOptions.plugins.tooltip.callbacks.label = (context) => ` ${valueLabel(context.parsed[horizontal ? 'x' : 'y'] ?? context.raw)}`;
  const lineOptions = isCount ? chartCountOptions(false) : chartCurrencyOptions(false);
  lineOptions.plugins.tooltip.callbacks.title = (items) => pivot.labels[items[0]?.dataIndex] || '';
  lineOptions.plugins.tooltip.callbacks.label = (context) => ` ${valueLabel(context.parsed.y ?? context.raw)}`;
  const donutOptions = {
    responsive: true,
    maintainAspectRatio: false,
    cutout: '63%',
    plugins: {
      legend: {
        display: true,
        position: 'bottom',
        labels: { color: '#718394', usePointStyle: true, pointStyle: 'circle', boxWidth: 7, boxHeight: 7, padding: 12, font: { size: 9 } },
      },
      tooltip: {
        backgroundColor: '#14283a',
        padding: 11,
        callbacks: {
          title: (items) => pivot.labels[items[0]?.dataIndex] || '',
          label: (context) => ` ${valueLabel(context.raw)}`,
        },
      },
    },
  };

  return (
    <section className="surface-card pivot-card">
      <div className="pivot-card-head">
        <div className="pivot-heading">
          <span className="pivot-icon"><SlidersHorizontal size={18} /></span>
          <div>
            <div className="card-eyebrow">EXPLORADOR DE DATOS</div>
            <h2>Gráfico dinámico</h2>
            <p>Elige agrupación, métrica y estilo para leer mejor los cobros.</p>
          </div>
        </div>
        <span className="pivot-count">{countLabel(pivot.transactionCount)} registros</span>
      </div>
      <div className="pivot-controls">
        <label className="pivot-select-wrap">
          <span>Agrupar por</span>
          <select value={dimension} onChange={(event) => setDimension(event.target.value)}>
            {options.map((option) => <option key={option.key} value={option.key}>{option.label}</option>)}
          </select>
          <ChevronDown size={14} />
        </label>
        <label className="pivot-select-wrap">
          <span>Mostrar</span>
          <select value={metric} onChange={(event) => setMetric(event.target.value)}>
            <option value="amount">Suma de Importe ($)</option>
            <option value="payments">Número de pagos</option>
            <option value="contributors">Contribuyentes únicos</option>
          </select>
          <ChevronDown size={14} />
        </label>
        <label className="pivot-select-wrap">
          <span>Tipo de gráfico</span>
          <select value={chartType} onChange={(event) => setChartType(event.target.value)}>
            <option value="bar">Barras</option>
            <option value="line">Líneas</option>
            <option value="area">Área</option>
            <option value="doughnut">Dona</option>
          </select>
          <ChevronDown size={14} />
        </label>
        <label className="pivot-select-wrap">
          <span>Grupos visibles</span>
          <select value={limit} onChange={(event) => setLimit(Number(event.target.value))}>
            <option value={5}>Top 5</option>
            <option value={10}>Top 10</option>
            <option value={20}>Top 20</option>
            <option value={0}>Todos</option>
          </select>
          <ChevronDown size={14} />
        </label>
        <label className="pivot-select-wrap">
          <span>Período</span>
          <select value={scope} onChange={(event) => setScope(event.target.value)}>
            <option value="all">Todos los períodos</option>
            <option value="selected">Solo {period ? monthLabel(period) : 'período seleccionado'}</option>
          </select>
          <ChevronDown size={14} />
        </label>
        <label className="pivot-toggle">
          <input type="checkbox" checked={rosterOnly} onChange={(event) => setRosterOnly(event.target.checked)} />
          <span className="toggle-ui" />
          <span>Solo Patrón</span>
        </label>
      </div>
      <div className={`pivot-chart-area ${horizontal ? 'pivot-horizontal' : ''} ${chartType === 'doughnut' ? 'pivot-doughnut' : ''}`}>
        {hasData ? (
          chartType === 'bar' ? (
            <Bar
              ref={chartRef}
              data={{
                labels: pivot.labels,
                datasets: [{
                  data: pivot.values,
                  backgroundColor: pivot.values.map((_, index) => palette[index % palette.length]),
                  borderRadius: 5,
                  maxBarThickness: horizontal ? 19 : 34,
                  barPercentage: 0.72,
                }],
              }}
              options={{ ...barOptions, indexAxis: horizontal ? 'y' : 'x', plugins: { ...barOptions.plugins, legend: { display: false } } }}
            />
          ) : chartType === 'doughnut' ? (
            <Doughnut
              ref={chartRef}
              data={{
                labels: pivot.labels,
                datasets: [{ data: pivot.values, backgroundColor: pivot.values.map((_, index) => palette[index % palette.length]), borderWidth: 0, hoverOffset: 5, spacing: 2 }],
              }}
              options={donutOptions}
            />
          ) : (
            <Line
              ref={chartRef}
              data={{
                labels: pivot.labels,
                datasets: [{
                  data: pivot.values,
                  borderColor: '#2876ce',
                  backgroundColor: chartType === 'area' ? 'rgba(40, 118, 206, .14)' : 'rgba(40, 118, 206, .03)',
                  pointBackgroundColor: '#ffffff',
                  pointBorderColor: '#2876ce',
                  pointBorderWidth: 2,
                  pointRadius: 3.5,
                  borderWidth: 2.4,
                  tension: 0.34,
                  fill: chartType === 'area',
                }],
              }}
              options={{ ...lineOptions, plugins: { ...lineOptions.plugins, legend: { display: false } } }}
            />
          )
        ) : <EmptyChart message="No hay datos para los filtros seleccionados." />}
      </div>
      <div className="pivot-card-foot">
        <span><Info size={14} /> La dimensión y el período se recalculan al cambiar los filtros.</span>
        <span>{limit > 0 && pivot.totalGroupCount > pivot.rows.length ? `Se muestran los ${limit} grupos principales de ${countLabel(pivot.totalGroupCount)}.` : 'Se muestran todos los grupos disponibles.'}</span>
      </div>
    </section>
  );
}

function ChartsView({ model, period, pivot, pivotOptions, pivotConfig, setPivotConfig, chartRef, onExport, onExportPayments }) {
  const { dimension, metric, scope, rosterOnly, chartType, limit } = pivotConfig;
  const change = (key, value) => setPivotConfig((previous) => ({ ...previous, [key]: value }));
  return (
    <div className="charts-view">
      <div className="view-intro-row">
        <div>
          <div className="section-kicker"><Sparkles size={15} /> ANÁLISIS INTERACTIVO</div>
          <h2>Gráficos dinámicos</h2>
          <p>Explora la recaudación por tributo, período, fecha de pago y las columnas disponibles en tu Excel.</p>
        </div>
        <div className="view-intro-meta"><span className="mini-live-dot" /> Actualizado con los filtros del panel</div>
      </div>
      <PivotExplorer
        pivot={pivot}
        options={pivotOptions}
        chartRef={chartRef}
        dimension={dimension}
        setDimension={(value) => change('dimension', value)}
        metric={metric}
        setMetric={(value) => change('metric', value)}
        scope={scope}
        setScope={(value) => change('scope', value)}
        rosterOnly={rosterOnly}
        setRosterOnly={(value) => change('rosterOnly', value)}
        chartType={chartType}
        setChartType={(value) => change('chartType', value)}
        limit={limit}
        setLimit={(value) => change('limit', value)}
        period={period}
      />
      <div className="charts-section-heading">
        <div><span className="card-eyebrow">RESUMEN DEL PERÍODO</span><h3>{period ? monthLabel(period) : 'Sin período seleccionado'}</h3></div>
        <div className="chart-export-actions">
          <button type="button" className="secondary-button export-button" onClick={onExport}><Download size={15} /> Exportar gráfico</button>
          <button type="button" className="secondary-button export-button" onClick={onExportPayments}><FileSpreadsheet size={15} /> Exportar pagos</button>
        </div>
      </div>
      <div className="content-grid chart-grid-summary charts-secondary-grid">
        <section className="surface-card chart-card">
          <CardHeader eyebrow="MES DE COBRO" title="Cuándo entró el dinero" subtitle="Agrupado por Fecha de pago real" />
          <div className="chart-area payment-month-area"><PaymentMonthChart model={model} /></div>
        </section>
        <section className="surface-card chart-card">
          <CardHeader eyebrow="ESTADO DE LOS CLIENTES" title="Puntualidad del Patrón" subtitle="Contribuyentes, no importe" />
          <div className="chart-area donut-area"><StatusDonut model={model} valueType="clients" /></div>
        </section>
      </div>
      <section className="surface-card chart-card top-chart-full">
        <CardHeader eyebrow="RANKING DEL PERÍODO" title="Contribuyentes con mayor importe" subtitle="Suma de los cobros registrados para el período seleccionado" />
        <div className="chart-area wide-top-area"><TopClientsChart model={model} /></div>
      </section>
    </div>
  );
}

function ClientsView({ model, period, search, setSearch, statusTab, setStatusTab, rows, onExport }) {
  const tabs = [
    { key: 'all', label: 'Todos', count: model.expectedCount },
    { key: 'late', label: 'Fuera de fecha', count: model.lateCount },
    { key: 'unpaid', label: 'Sin pago', count: model.unpaidCount + model.pendingCount },
    { key: 'paid', label: 'En fecha', count: model.statusCounts.paid },
  ];
  return (
    <div className="clients-view">
      <div className="view-intro-row clients-intro">
        <div>
          <div className="section-kicker"><Users size={15} /> PATRÓN DE CONTRIBUYENTES</div>
          <h2>Seguimiento de pagos</h2>
          <p>{period ? `Estado de pago de cada cliente para ${monthLabel(period)}.` : 'Selecciona un período para ver el estado de pago.'}</p>
        </div>
      </div>
      {model.missingRoster && (
        <Notice icon={Info} tone="info" title="Falta un Patrón completo">
          Sin el listado de clientes esperados no es posible identificar a quienes no pagaron. Usa “Importar Excel” para cargar el Patrón y completar este seguimiento.
        </Notice>
      )}
      <section className="surface-card clients-table-card">
        <div className="clients-table-toolbar">
          <div className="table-tabs" role="tablist" aria-label="Filtrar por estado">
            {tabs.map((tab) => (
              <button
                type="button"
                role="tab"
                aria-selected={statusTab === tab.key}
                className={`table-tab ${statusTab === tab.key ? 'selected' : ''}`}
                key={tab.key}
                onClick={() => setStatusTab(tab.key)}
              >
                {tab.label}<span>{countLabel(tab.count)}</span>
              </button>
            ))}
          </div>
          <div className="table-actions">
            <label className="search-field">
              <Search size={16} />
              <input value={search} onChange={(event) => setSearch(event.target.value)} placeholder="Buscar contribuyente…" />
              {search && <button type="button" aria-label="Limpiar búsqueda" onClick={() => setSearch('')}><X size={14} /></button>}
            </label>
            <button type="button" className="secondary-button export-button table-export" onClick={onExport}><Download size={15} /><span>Exportar</span></button>
          </div>
        </div>
        <div className="table-result-line">
          <span><strong>{countLabel(rows.length)}</strong> resultados{period ? ` · ${monthLabel(period)}` : ''}</span>
          {model.deadline && <span className="table-deadline"><CalendarDays size={14} /> Vencimiento: {formatDate(model.deadline)}</span>}
        </div>
        {rows.length ? (
          <div className="table-scroll large-table-scroll">
            <table className="data-table clients-table">
              <thead>
                <tr>
                  <th>Contribuyente</th>
                  <th>Estado</th>
                  <th className="align-right">Importe ($)</th>
                  <th>Último pago</th>
                  <th className="align-right">Atraso</th>
                </tr>
              </thead>
              <tbody>
                {rows.map((row) => (
                  <tr key={row.key}>
                    <td>
                      <div className="contributor-cell">
                        <span className={`client-avatar avatar-${statusTone(row.status)}`}>{row.contributor.charAt(0).toUpperCase()}</span>
                        <span className="client-name-wrap"><strong title={row.contributor}>{row.contributor}</strong><small>{row.tax || 'Sin tributo asociado'}</small></span>
                      </div>
                    </td>
                    <td><StatusBadge status={row.status} /></td>
                    <td className="align-right amount-cell">{row.paymentCount ? money(row.amount) : '—'}</td>
                    <td className="date-cell">{row.lastPayment ? formatDate(row.lastPayment) : '—'}{row.paymentCount > 1 && <small>{row.paymentCount} pagos</small>}</td>
                    <td className={`align-right late-days-cell ${row.lateDays > 0 ? 'late-days-active' : ''}`}>
                      {row.lateDays > 0 ? `${countLabel(row.lateDays)} d` : '—'}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        ) : (
          <div className="table-empty-result">
            <div className="empty-chart-icon"><Search size={19} /></div>
            <strong>No hay resultados con esos filtros</strong>
            <span>Prueba otra búsqueda o cambia el estado seleccionado.</span>
          </div>
        )}
        <div className="table-footnote"><Info size={14} /> “Sin pago” significa que no aparece una transacción asociada en el Excel cargado; no calcula deuda ni saldo pendiente.</div>
      </section>
    </div>
  );
}

function ImportCard({ role, meta, busy, onSelect }) {
  const inputRef = useRef(null);
  const isClients = role === 'clients';
  const handleDrop = (event) => {
    event.preventDefault();
    const file = event.dataTransfer.files?.[0];
    if (file) onSelect(file, role);
  };
  return (
    <div className={`import-card ${meta ? 'has-file' : ''}`} onDragOver={(event) => event.preventDefault()} onDrop={handleDrop}>
      <input
        ref={inputRef}
        className="visually-hidden"
        type="file"
        accept=".xls,.xlsx,.xlsm,.xlsb"
        onChange={(event) => {
          const file = event.target.files?.[0];
          if (file) onSelect(file, role);
          event.target.value = '';
        }}
      />
      <div className={`import-card-icon ${isClients ? 'icon-violet' : 'icon-green'}`}>
        {isClients ? <Users size={20} /> : <FileSpreadsheet size={20} />}
      </div>
      <div className="import-card-copy">
        <div className="import-card-title-row">
          <h3>{isClients ? 'Patrón de clientes' : 'Registro de pagos'}</h3>
          {meta && <span className="file-ready-chip"><Check size={12} /> Cargado</span>}
        </div>
        <p>{isClients ? 'Listado con todos los contribuyentes que deben ser revisados.' : 'Transacciones con Contribuyente, Importe ($) y Fecha de pago.'}</p>
        {meta ? (
          <>
            <div className="import-file-meta">
              <FileText size={15} />
              <span title={meta.fileName}>{meta.fileName}</span>
              <small>{countLabel(meta.count)} filas · {fileSize(meta.size)}</small>
            </div>
            {meta.fields?.length > 0 && (
              <div className="import-detected-fields" aria-label="Columnas extraídas del archivo">
                <strong>Datos detectados</strong>
                <div className="import-detected-field-list">
                  {meta.fields.map((field, index) => <span key={`${field}-${index}`}>{field}</span>)}
                </div>
              </div>
            )}
          </>
        ) : (
          <div className="import-card-hint">Arrastra un archivo aquí o selecciónalo</div>
        )}
      </div>
      <button type="button" className={`import-card-button ${meta ? 'replace-button' : ''}`} onClick={() => inputRef.current?.click()} disabled={busy}>
        {busy ? <><span className="button-spinner" /> Leyendo</> : <><Upload size={15} /> {meta ? 'Reemplazar' : 'Seleccionar'}</>}
      </button>
    </div>
  );
}

function ImportModal({ open, onClose, onImport, data, busyRole, lastImport }) {
  useEffect(() => {
    if (!open) return undefined;
    const onKeyDown = (event) => event.key === 'Escape' && onClose();
    window.addEventListener('keydown', onKeyDown);
    return () => window.removeEventListener('keydown', onKeyDown);
  }, [open, onClose]);
  if (!open) return null;

  return (
    <div className="modal-backdrop" role="presentation" onMouseDown={(event) => event.target === event.currentTarget && onClose()}>
      <section className="modal import-modal" role="dialog" aria-modal="true" aria-labelledby="import-title">
        <div className="modal-header">
          <div className="modal-title-icon"><FileSpreadsheet size={20} /></div>
          <div><div className="modal-eyebrow">ACTUALIZAR INFORMACIÓN</div><h2 id="import-title">Importar desde Excel</h2></div>
          <button type="button" className="icon-button modal-close" aria-label="Cerrar" onClick={onClose}><X size={18} /></button>
        </div>
        <p className="modal-intro">Separa el Patrón de clientes del registro de pagos para detectar contribuyentes sin pago y analizar los cobros.</p>
        <div className="import-cards">
          <ImportCard role="clients" meta={data.clientMeta} busy={busyRole === 'clients'} onSelect={onImport} />
          <ImportCard role="payments" meta={data.paymentMeta} busy={busyRole === 'payments'} onSelect={onImport} />
        </div>
        {lastImport && <div className="import-result"><CheckCircle2 size={16} /><span>{lastImport}</span></div>}
        <div className="import-guidance">
          <div className="guidance-icon"><Info size={16} /></div>
          <p><strong>Archivos compatibles: .xls, .xlsx, .xlsm y .xlsb.</strong> Se detectan y conservan las columnas originales para exportación; los campos adicionales de pagos también se pueden usar en gráficos. El cruce prioriza el código de Contribuyente (antes del guion); si falta, usa el nombre. Sin “Período”, se toma el mes de “Fecha de pago”.</p>
        </div>
        <div className="modal-privacy"><LockKeyhole size={14} /> Los libros se leen en este navegador y no se envían a un servidor.</div>
        <div className="modal-footer"><span>Las preferencias de vencimiento sí se guardan en este dispositivo.</span><button type="button" className="secondary-button" onClick={onClose}>Listo</button></div>
      </section>
    </div>
  );
}

const paperOptions = [
  { key: 'a4', label: 'A4', description: '210 × 297 mm' },
  { key: 'letter', label: 'Carta', description: '8.5 × 11 in' },
  { key: 'a3', label: 'A3', description: '297 × 420 mm' },
  { key: 'legal', label: 'Legal', description: '8.5 × 14 in' },
  { key: 'a5', label: 'A5', description: '148 × 210 mm' },
];

function ExportModal({ open, onClose, onExport, payload, busy }) {
  const [format, setFormat] = useState('xlsx');
  const [paperSize, setPaperSize] = useState('a4');
  const [orientation, setOrientation] = useState('landscape');
  useEffect(() => {
    if (open) {
      setFormat('xlsx');
      setPaperSize('a4');
      setOrientation('landscape');
    }
  }, [open]);
  useEffect(() => {
    if (!open) return undefined;
    const onKeyDown = (event) => event.key === 'Escape' && onClose();
    window.addEventListener('keydown', onKeyDown);
    return () => window.removeEventListener('keydown', onKeyDown);
  }, [open, onClose]);
  if (!open || !payload) return null;

  const formats = [
    { key: 'xlsx', title: 'Excel (.xlsx)', detail: 'Editable y preparado para impresión', icon: FileSpreadsheet },
    { key: 'pdf', title: 'PDF', detail: payload.chartImage ? 'Incluye el gráfico y su tabla de datos' : 'Documento paginado para compartir o imprimir', icon: FileText },
    { key: 'csv', title: 'CSV', detail: 'Datos simples para otros sistemas', icon: FileOutput },
  ].filter((item) => !payload.formats || payload.formats.includes(item.key));
  const extension = format === 'xlsx' ? '.xlsx' : format === 'pdf' ? '.pdf' : '.csv';
  return (
    <div className="modal-backdrop" role="presentation" onMouseDown={(event) => event.target === event.currentTarget && onClose()}>
      <section className="modal export-modal" role="dialog" aria-modal="true" aria-labelledby="export-title">
        <div className="modal-header">
          <div className="modal-title-icon export-modal-icon"><FileDown size={20} /></div>
          <div><div className="modal-eyebrow">DESCARGAR REPORTE</div><h2 id="export-title">Exportar datos</h2></div>
          <button type="button" className="icon-button modal-close" aria-label="Cerrar" onClick={onClose}><X size={18} /></button>
        </div>
        <p className="modal-intro">Se exportarán <strong>{countLabel(payload.rows.length)} filas</strong> de “{payload.title}”{payload.subtitle ? ` · ${payload.subtitle}` : ''}.</p>
        {payload.note && <div className="export-format-note"><Info size={14} /> {payload.note}</div>}
        <div className="export-format-grid" role="radiogroup" aria-label="Formato de exportación">
          {formats.map(({ key, title: formatTitle, detail, icon: Icon }) => (
            <button
              type="button"
              role="radio"
              aria-checked={format === key}
              className={`export-format-card ${format === key ? 'format-selected' : ''}`}
              key={key}
              onClick={() => setFormat(key)}
            >
              <span className="format-card-icon"><Icon size={19} /></span>
              <span className="format-card-copy"><strong>{formatTitle}</strong><small>{detail}</small></span>
              <span className="format-radio" />
            </button>
          ))}
        </div>
        {format !== 'csv' && (
          <div className="export-page-settings">
            <div className="export-settings-heading"><Printer size={16} /><strong>Configuración de página</strong><span>{!payload.formats || payload.formats.includes('pdf') ? 'Se aplica al Excel y al PDF' : 'Se aplica al Excel'}</span></div>
            <div className="export-settings-grid">
              <label className="export-control">
                <span>Tamaño de papel</span>
                <select value={paperSize} onChange={(event) => setPaperSize(event.target.value)}>
                  {paperOptions.map((paper) => <option key={paper.key} value={paper.key}>{paper.label} · {paper.description}</option>)}
                </select>
                <ChevronDown size={14} />
              </label>
              <div className="export-control orientation-control">
                <span>Orientación</span>
                <div className="orientation-options" role="group" aria-label="Orientación de la hoja">
                  <button type="button" aria-pressed={orientation === 'landscape'} className={orientation === 'landscape' ? 'orientation-selected' : ''} onClick={() => setOrientation('landscape')}><RotateCcw size={14} /> Horizontal</button>
                  <button type="button" aria-pressed={orientation === 'portrait'} className={orientation === 'portrait' ? 'orientation-selected' : ''} onClick={() => setOrientation('portrait')}><RotateCcw className="portrait-icon" size={14} /> Vertical</button>
                </div>
              </div>
            </div>
            <div className="export-format-note"><Info size={14} /> La primera columna irá en negrita; los anchos se ajustarán al contenido y la impresión se ajustará al ancho de la página.</div>
          </div>
        )}
        <div className="modal-footer export-modal-footer">
          <span>{format === 'csv' ? 'CSV conserva los valores sin formato de página.' : 'La configuración elegida se aplicará a este archivo.'}</span>
          <div>
            <button type="button" className="secondary-button" onClick={onClose} disabled={busy}>Cancelar</button>
            <button type="button" className="primary-button" onClick={() => onExport({ format, paperSize, orientation })} disabled={busy}>
              {busy ? <><span className="button-spinner" /> Preparando…</> : <><Download size={15} /> Descargar {extension}</>}
            </button>
          </div>
        </div>
      </section>
    </div>
  );
}

function SettingsModal({ open, onClose, workingDays, holidays, onSave }) {
  const [draftDays, setDraftDays] = useState(workingDays);
  const [draftHolidays, setDraftHolidays] = useState(holidays);
  const [holidayInput, setHolidayInput] = useState('');
  useEffect(() => {
    if (open) {
      setDraftDays(workingDays);
      setDraftHolidays(holidays);
      setHolidayInput('');
    }
  }, [open, workingDays, holidays]);
  useEffect(() => {
    if (!open) return undefined;
    const onKeyDown = (event) => event.key === 'Escape' && onClose();
    window.addEventListener('keydown', onKeyDown);
    return () => window.removeEventListener('keydown', onKeyDown);
  }, [open, onClose]);
  if (!open) return null;

  const addHoliday = () => {
    if (!holidayInput || draftHolidays.includes(holidayInput)) return;
    setDraftHolidays((previous) => [...previous, holidayInput].sort());
    setHolidayInput('');
  };
  const save = () => {
    onSave(Math.max(1, Math.min(23, Number(draftDays) || 10)), draftHolidays);
    onClose();
  };

  return (
    <div className="modal-backdrop" role="presentation" onMouseDown={(event) => event.target === event.currentTarget && onClose()}>
      <section className="modal settings-modal" role="dialog" aria-modal="true" aria-labelledby="settings-title">
        <div className="modal-header">
          <div className="modal-title-icon settings-modal-icon"><Settings2 size={20} /></div>
          <div><div className="modal-eyebrow">REGLA DE PUNTUALIDAD</div><h2 id="settings-title">Configurar vencimiento</h2></div>
          <button type="button" className="icon-button modal-close" aria-label="Cerrar" onClick={onClose}><X size={18} /></button>
        </div>
        <p className="modal-intro">El panel marca un pago como “fuera de fecha” cuando su Fecha de pago supera el vencimiento calculado para el período.</p>
        <div className="settings-field-block">
          <label htmlFor="workday-count">Día hábil límite</label>
          <div className="workday-input-wrap">
            <input id="workday-count" type="number" min="1" max="23" value={draftDays} onChange={(event) => setDraftDays(event.target.value)} />
            <span>primeros días hábiles del mes</span>
          </div>
          <small>Se excluyen sábados y domingos. Un pago realizado el mismo día límite cuenta como puntual.</small>
        </div>
        <div className="settings-field-block holiday-field-block">
          <div className="holiday-heading"><label>Feriados que no cuentan como hábiles</label><span>{draftHolidays.length} agregados</span></div>
          <div className="holiday-add-row">
            <input type="date" value={holidayInput} onChange={(event) => setHolidayInput(event.target.value)} aria-label="Fecha de feriado" />
            <button type="button" className="secondary-button" onClick={addHoliday} disabled={!holidayInput}><span>+</span> Añadir fecha</button>
          </div>
          {draftHolidays.length ? (
            <div className="holiday-list">
              {draftHolidays.map((date) => (
                <span className="holiday-chip" key={date}>{formatDate(new Date(`${date}T00:00:00Z`))}<button type="button" aria-label={`Quitar feriado ${date}`} onClick={() => setDraftHolidays((previous) => previous.filter((item) => item !== date))}><X size={13} /></button></span>
              ))}
            </div>
          ) : <p className="holiday-empty">No hay feriados personalizados. No se precarga un calendario oficial.</p>}
        </div>
        <div className="settings-note"><Info size={15} /> Revisa y agrega los feriados aplicables a tu territorio; el sistema no descarga calendarios externos.</div>
        <div className="modal-footer settings-footer"><span>La regla queda guardada en este navegador.</span><div><button type="button" className="secondary-button" onClick={onClose}>Cancelar</button><button type="button" className="primary-button" onClick={save}>Guardar regla</button></div></div>
      </section>
    </div>
  );
}

function App() {
  const [data, setData] = useState({ clients: [], payments: [], dimensions: [], clientFields: [], paymentFields: [], clientMeta: null, paymentMeta: null });
  const [loadingSample, setLoadingSample] = useState(true);
  const [sampleError, setSampleError] = useState('');
  const [period, setPeriod] = useState('');
  const [taxFilter, setTaxFilter] = useState('');
  const [activeView, setActiveView] = useState('overview');
  const [workingDays, setWorkingDays] = useState(() => safeStorageRead('onat-working-days', 10));
  const [holidays, setHolidays] = useState(() => safeStorageRead('onat-holidays', []));
  const [importOpen, setImportOpen] = useState(false);
  const [settingsOpen, setSettingsOpen] = useState(false);
  const [exportOpen, setExportOpen] = useState(false);
  const [exportPayload, setExportPayload] = useState(null);
  const [exportBusy, setExportBusy] = useState(false);
  const [busyRole, setBusyRole] = useState('');
  const [lastImport, setLastImport] = useState('');
  const [toast, setToast] = useState(null);
  const [search, setSearch] = useState('');
  const [statusTab, setStatusTab] = useState('all');
  const [quarterlyNoticeDismissed, setQuarterlyNoticeDismissed] = useState(false);
  const [pivotConfig, setPivotConfig] = useState({ dimension: 'period', metric: 'amount', scope: 'all', rosterOnly: true, chartType: 'bar', limit: 10 });
  const pivotChartRef = useRef(null);

  const periods = useMemo(() => Array.from(new Set(data.payments.map((payment) => payment.periodKey).filter(Boolean))).sort(), [data.payments]);
  const taxOptions = useMemo(() => uniqueTaxOptions(data.clients, data.payments), [data.clients, data.payments]);

  useEffect(() => {
    let cancelled = false;
    if (!DEMO_DATA_ENABLED) {
      // A production/HTTPS build starts blank so the repository's example
      // spreadsheets are never copied into the public Pages artifact.
      setLoadingSample(false);
      return () => { cancelled = true; };
    }

    async function loadSamples() {
      setLoadingSample(true);
      setSampleError('');
      try {
        const [clientsResponse, paymentsResponse] = await Promise.all([fetch(DEMO_CLIENTS_URL), fetch(DEMO_PAYMENTS_URL)]);
        if (!clientsResponse.ok || !paymentsResponse.ok) throw new Error('No se pudieron abrir los Excel de muestra.');
        const [clientsBuffer, paymentsBuffer] = await Promise.all([clientsResponse.arrayBuffer(), paymentsResponse.arrayBuffer()]);
        const [clientsResult, paymentsResult] = await Promise.all([
          parseWorkbook(clientsBuffer, 'clients', 'Cliente - obligación (vector fiscal).xlsx'),
          parseWorkbook(paymentsBuffer, 'payments', 'Pago1 (persona jurídica).xlsx'),
        ]);
        if (cancelled) return;
        setData({
          clients: clientsResult.clients,
          payments: paymentsResult.payments,
          dimensions: paymentsResult.dimensions,
          clientFields: clientsResult.fields,
          paymentFields: paymentsResult.fields,
          clientMeta: { fileName: clientsResult.fileName, count: clientsResult.clients.length, size: clientsBuffer.byteLength, sample: true, fields: clientsResult.fields.map((field) => field.label) },
          paymentMeta: { fileName: paymentsResult.fileName, count: paymentsResult.payments.length, size: paymentsBuffer.byteLength, sample: true, skippedRows: paymentsResult.skippedRows, fields: paymentsResult.fields.map((field) => field.label) },
        });
      } catch (error) {
        if (!cancelled) setSampleError(error.message || 'No se pudieron cargar los Excel de muestra.');
      } finally {
        if (!cancelled) setLoadingSample(false);
      }
    }
    loadSamples();
    return () => { cancelled = true; };
  }, []);

  useEffect(() => {
    if (!periods.length) {
      setPeriod('');
      return;
    }
    setPeriod((current) => periods.includes(current) ? current : bestPeriod(data.payments));
  }, [periods, data.payments]);

  useEffect(() => {
    const masterTaxes = Array.from(new Set(data.clients.map((client) => client.taxKey).filter(Boolean)));
    if (masterTaxes.length === 1) setTaxFilter(masterTaxes[0]);
    else if (taxFilter && !taxOptions.some((option) => option.key === taxFilter)) setTaxFilter('');
  }, [data.clients, data.clientMeta?.fileName, taxOptions]); // set a helpful default only when a new source is loaded

  useEffect(() => safeStorageWrite('onat-working-days', workingDays), [workingDays]);
  useEffect(() => safeStorageWrite('onat-holidays', holidays), [holidays]);

  const model = useMemo(() => buildDashboardModel({
    clients: data.clients,
    payments: data.payments,
    period,
    taxFilter,
    workingDays,
    holidays,
    today: clientDateToday(),
  }), [data.clients, data.payments, period, taxFilter, workingDays, holidays]);

  const pivotOptions = useMemo(() => {
    const options = [
      { key: 'period', label: 'Período de obligación' },
      { key: 'paymentMonth', label: 'Mes de Fecha de pago' },
      { key: 'contributor', label: 'Contribuyente' },
      { key: 'tax', label: 'Tributo' },
    ];
    const known = new Set([...options.map((item) => item.key), 'tributo']);
    data.dimensions.forEach((dimension) => {
      if (known.has(dimension.key)) return;
      options.push({ key: dimension.key, label: formatHeaderName(dimension.label) });
      known.add(dimension.key);
    });
    return options;
  }, [data.dimensions]);

  useEffect(() => {
    if (!pivotOptions.some((option) => option.key === pivotConfig.dimension)) {
      setPivotConfig((previous) => ({ ...previous, dimension: 'period' }));
    }
  }, [pivotOptions, pivotConfig.dimension]);

  const pivot = useMemo(() => buildPivot({
    payments: data.payments,
    clients: data.clients,
    period,
    taxFilter,
    dimension: pivotConfig.dimension,
    metric: pivotConfig.metric,
    scope: pivotConfig.scope,
    rosterOnly: pivotConfig.rosterOnly,
    limit: pivotConfig.limit,
  }), [data.payments, data.clients, period, taxFilter, pivotConfig]);

  const visibleClientRows = useMemo(() => {
    const query = normalizeText(search);
    return model.rosterRows.filter((row) => {
      const statusMatch = statusTab === 'all'
        || (statusTab === 'late' && row.status === 'late')
        || (statusTab === 'unpaid' && ['unpaid', 'pending'].includes(row.status))
        || (statusTab === 'paid' && row.status === 'paid');
      const queryMatch = !query || normalizeText(`${row.contributor} ${row.tax} ${row.key}`).includes(query);
      return statusMatch && queryMatch;
    });
  }, [model.rosterRows, search, statusTab]);

  const currentTitle = activeView === 'control' ? 'Panel de control' : activeView === 'clients' ? 'Contribuyentes' : activeView === 'charts' ? 'Gráficos dinámicos' : 'Resumen';

  const handleImport = async (file, role) => {
    setBusyRole(role);
    setLastImport('');
    try {
      const buffer = await file.arrayBuffer();
      const parsed = await parseWorkbook(buffer, role, file.name);
      if (role === 'clients') {
        setData((previous) => ({
          ...previous,
          clients: parsed.clients,
          clientFields: parsed.fields,
          clientMeta: { fileName: file.name, count: parsed.clients.length, size: file.size, skippedRows: parsed.skippedRows, fields: parsed.fields.map((field) => field.label) },
        }));
      } else {
        setData((previous) => ({
          ...previous,
          payments: parsed.payments,
          dimensions: parsed.dimensions,
          paymentFields: parsed.fields,
          paymentMeta: { fileName: file.name, count: parsed.payments.length, size: file.size, skippedRows: parsed.skippedRows, fields: parsed.fields.map((field) => field.label) },
        }));
      }
      const roleName = role === 'clients' ? 'Patrón' : 'registro de pagos';
      const skipped = parsed.skippedRows ? ` Se omitieron ${countLabel(parsed.skippedRows)} filas incompletas.` : '';
      const dateNote = role === 'payments' && parsed.payments.every((payment) => payment.dateBasis === 'payment')
        ? ' El período se tomará de Fecha de pago.'
        : '';
      const fieldsNote = parsed.fields?.length ? ` Se detectaron ${countLabel(parsed.fields.length)} columnas y se conservaron para consulta y exportación.` : '';
      setLastImport(`${countLabel(role === 'clients' ? parsed.clients.length : parsed.payments.length)} filas cargadas en el ${roleName}.${skipped}${dateNote}${fieldsNote}`);
      setToast({ type: 'success', message: `${role === 'clients' ? 'Patrón' : 'Pagos'} importados: ${countLabel(role === 'clients' ? parsed.clients.length : parsed.payments.length)} filas.` });
    } catch (error) {
      setToast({ type: 'error', message: error.message || 'No se pudo leer el archivo.' });
    } finally {
      setBusyRole('');
    }
  };

  const handleSaveRule = (days, nextHolidays) => {
    setWorkingDays(days);
    setHolidays(nextHolidays);
    setToast({ type: 'success', message: 'Regla de vencimiento actualizada.' });
  };

  const makeClientExport = (rows, title = 'Seguimiento de pagos') => {
    const standardHeaders = ['Contribuyente', 'Tributo', 'Estado', 'Importe cobrado ($)', 'Cantidad de pagos', 'Primer pago', 'Último pago', 'Días de atraso'];
    const extraFields = data.clientFields.filter((field) => !['contribuyente', 'tributo'].includes(field.key));
    const provenanceHeaders = ['Fila de origen (Patrón)', 'Hoja de origen (Patrón)'];
    const headings = [...standardHeaders, ...provenanceHeaders, ...extraFields.map((field) => field.label)];
    const taxLabel = taxFilter ? (taxOptions.find((option) => option.key === taxFilter)?.label || taxFilter) : 'Todos los tributos';
    const subtitle = `Período: ${period ? monthLabel(period) : 'todos'} · Tributo: ${taxLabel} · Regla: ${workingDays} días hábiles`;
    const exportRows = rows.map((row) => [
      row.contributor,
      row.tax || '—',
      statusLabel(row.status),
      Number(row.amount) || 0,
      Number(row.paymentCount) || 0,
      row.firstPayment || null,
      row.lastPayment || null,
      Number(row.lateDays) || 0,
      row.sourceRow ?? '',
      row.sourceSheet || '',
      ...extraFields.map((field) => row.sourceData?.[field.key] ?? ''),
    ]);
    return { title, fileBase: 'seguimiento', period, subtitle, headers: headings, rows: exportRows };
  };

  const openClientExport = (rows = model.rosterRows, title = 'Seguimiento de pagos') => {
    if (!rows.length) {
      setToast({ type: 'error', message: 'No hay filas para exportar con estos filtros.' });
      return;
    }
    setExportPayload(makeClientExport(rows, title));
    setExportOpen(true);
  };

  const openPaymentExport = () => {
    const selectedPeriod = pivotConfig.scope === 'selected' && period ? period : '';
    const periodsToExport = selectedPeriod
      ? [selectedPeriod]
      : Array.from(new Set(data.payments.map((payment) => payment.periodKey).filter(Boolean))).sort();
    const paymentRows = periodsToExport.flatMap((periodKey) => getPaymentsForPeriod({
      payments: data.payments,
      clients: data.clients,
      period: periodKey,
      taxFilter,
      rosterOnly: pivotConfig.rosterOnly,
    }));
    if (!paymentRows.length) {
      setToast({ type: 'error', message: 'No hay movimientos para exportar con estos filtros.' });
      return;
    }

    const sourceFields = data.paymentFields;
    const headers = [
      ...sourceFields.map((field) => field.label),
      'Período analizado',
      'Base del período',
      'Fila de origen',
      'Hoja de origen',
    ];
    const rows = paymentRows.map((payment) => [
      ...sourceFields.map((field) => payment.sourceData?.[field.key] ?? ''),
      payment.periodKey,
      payment.dateBasis === 'period' ? 'Columna Período' : 'Fecha de pago',
      payment.sourceRow,
      payment.sheetName,
    ]);
    const taxLabel = taxFilter ? (taxOptions.find((option) => option.key === taxFilter)?.label || taxFilter) : 'Todos los tributos';
    setExportPayload({
      title: 'Movimientos de pago',
      fileBase: 'movimientos-de-pago',
      period: selectedPeriod,
      subtitle: `${selectedPeriod ? monthLabel(selectedPeriod) : 'Todos los períodos'} · ${taxLabel} · ${countLabel(paymentRows.length)} movimientos`,
      headers,
      rows,
      formats: ['xlsx', 'csv'],
      note: 'Se incluyen las columnas originales del Excel, el período usado para el análisis y la fila/hoja de origen. Para este volumen se ofrecen XLSX y CSV.',
    });
    setExportOpen(true);
  };

  const openPivotExport = () => {
    const metricLabel = pivotConfig.metric === 'amount' ? 'Suma de Importe ($)' : pivotConfig.metric === 'payments' ? 'Número de pagos' : 'Contribuyentes únicos';
    const dimensionLabel = pivotOptions.find((option) => option.key === pivotConfig.dimension)?.label || pivotConfig.dimension;
    const rows = pivot.rows.map((row) => [row.label, row.value, metricLabel]);
    if (!rows.length) {
      setToast({ type: 'error', message: 'No hay grupos para exportar.' });
      return;
    }
    setExportPayload({
      title: 'Gráfico dinámico',
      fileBase: `grafico-${pivotConfig.dimension}`,
      period,
      subtitle: `${dimensionLabel} · ${metricLabel} · ${pivotConfig.scope === 'all' ? 'todos los períodos' : monthLabel(period)}`,
      chartImage: pivotChartRef.current?.toBase64Image?.() || null,
      headers: ['Agrupación', pivotConfig.metric === 'amount' ? 'Importe ($)' : 'Valor', 'Métrica'],
      rows,
    });
    setExportOpen(true);
  };

  const handleExport = async ({ format, paperSize, orientation }) => {
    if (!exportPayload || exportBusy) return;
    setExportBusy(true);
    try {
      const fileName = createFilename(exportPayload.fileBase || exportPayload.title, exportPayload.period || '', format);
      if (format === 'csv') {
        downloadCSV(exportPayload.headers, exportPayload.rows, fileName);
      } else {
        const options = { ...exportPayload, paperSize, orientation };
        const blob = format === 'xlsx' ? await createXlsxBlob(options) : await createPdfBlob(options);
        downloadBlob(blob, fileName);
      }
      setExportOpen(false);
      setToast({ type: 'success', message: `Reporte ${format.toUpperCase()} descargado: ${fileName}` });
    } catch (error) {
      setToast({ type: 'error', message: error.message || 'No se pudo preparar la exportación.' });
    } finally {
      setExportBusy(false);
    }
  };

  const handleNavigate = (view) => {
    setActiveView(view);
    setSearch('');
    setStatusTab('all');
  };

  const title = activeView === 'control' ? 'Panel de control' : activeView === 'charts' ? 'Análisis de cobros' : activeView === 'clients' ? 'Seguimiento de contribuyentes' : 'Disciplina de pago';
  const subtitle = activeView === 'control'
    ? 'Administra tus fuentes, valida los datos y coordina el seguimiento del período.'
    : activeView === 'charts'
      ? 'Cruza las columnas de tu Excel y explora la recaudación con gráficos interactivos.'
      : activeView === 'clients'
        ? 'Consulta quién pagó, quién se atrasó y quién no tiene pago registrado.'
        : 'Una vista clara de los cobros, la puntualidad y los contribuyentes pendientes.';

  return (
    <div className="app-shell">
      <Sidebar activeView={activeView} onNavigate={handleNavigate} model={model} loading={loadingSample} />
      <main className="main-shell">
        <TopBar
          onImport={() => setImportOpen(true)}
          onExport={() => openClientExport(model.rosterRows, 'Seguimiento de pagos')}
          exportDisabled={loadingSample || !model.rosterRows.length}
          showExport={(activeView === 'overview' || activeView === 'control') && model.rosterRows.length > 0}
          showImport={activeView === 'control' || data.payments.length > 0}
          currentTitle={currentTitle}
        />
        <div className="main-content">
          <PageHeading
            title={title}
            subtitle={subtitle}
            period={period}
            setPeriod={setPeriod}
            periods={periods}
            taxFilter={taxFilter}
            setTaxFilter={setTaxFilter}
            taxOptions={taxOptions}
            onSettings={() => setSettingsOpen(true)}
            loading={loadingSample}
          />
          {activeView === 'overview' && (loadingSample || data.payments.length > 0) && (
            <DeadlineStrip model={model} period={period} workingDays={workingDays} holidays={holidays} loading={loadingSample} />
          )}
          {!loadingSample && activeView !== 'control' && <PeriodBasisNotice payments={data.payments} />}
          {activeView !== 'overview' && model.quarterlyWarning && !quarterlyNoticeDismissed && (
            <Notice icon={AlertTriangle} tone="warning" title="El Patrón indica periodicidad trimestral" onClose={() => setQuarterlyNoticeDismissed(true)}>
              Los estados se calculan con la regla mensual de {workingDays} días hábiles. Comprueba que el período seleccionado corresponda a una obligación mensual.
            </Notice>
          )}

          {sampleError && (
            <Notice icon={AlertTriangle} tone="warning" title="No se cargó la muestra">
              {sampleError} {activeView === 'control' ? 'Usa “Importar Excel” en la barra superior para cargar tus archivos.' : 'Usa el botón de la tarjeta de inicio para cargar tus propios archivos.'}
            </Notice>
          )}

          {!loadingSample && !data.payments.length && activeView !== 'control' ? (
            <GetStartedCard data={data} onImport={() => setImportOpen(true)} />
          ) : (
            <>
              {activeView === 'overview' && <StatsGrid model={model} loading={loadingSample} />}

              {activeView === 'control' && (
                <ControlPanelView
                  model={model}
                  period={period}
                  workingDays={workingDays}
                  holidays={holidays}
                  data={data}
                />
              )}

              {activeView === 'overview' && (
                <SummaryView
                  model={{ ...model, quarterlyWarning: model.quarterlyWarning && !quarterlyNoticeDismissed }}
                  period={period}
                  workingDays={workingDays}
                  onViewAll={() => handleNavigate('clients')}
                  onDismissQuarterly={() => setQuarterlyNoticeDismissed(true)}
                />
              )}

              {activeView === 'charts' && (
                <ChartsView
                  model={model}
                  period={period}
                  pivot={pivot}
                  pivotOptions={pivotOptions}
                  pivotConfig={pivotConfig}
                  setPivotConfig={setPivotConfig}
                  chartRef={pivotChartRef}
                  onExport={openPivotExport}
                  onExportPayments={openPaymentExport}
                />
              )}

              {activeView === 'clients' && (
                <ClientsView
                  model={model}
                  period={period}
                  search={search}
                  setSearch={setSearch}
                  statusTab={statusTab}
                  setStatusTab={setStatusTab}
                  rows={visibleClientRows}
                  onExport={() => openClientExport(visibleClientRows, 'Seguimiento de contribuyentes')}
                />
              )}
            </>
          )}

          <footer className="page-footer">
            <div className="page-footer-local">
              <span className="footer-privacy"><span className="footer-lock"><LockKeyhole size={13} /></span> Datos procesados localmente · No se transmite información fiscal</span>
              <span className="footer-version">Versión {APP_VERSION}</span>
            </div>
            <span className="page-footer-brand">ONAT · Panel de disciplina de pago</span>
          </footer>
        </div>
      </main>
      <ImportModal open={importOpen} onClose={() => setImportOpen(false)} onImport={handleImport} data={data} busyRole={busyRole} lastImport={lastImport} />
      <ExportModal open={exportOpen} onClose={() => setExportOpen(false)} onExport={handleExport} payload={exportPayload} busy={exportBusy} />
      <SettingsModal open={settingsOpen} onClose={() => setSettingsOpen(false)} workingDays={workingDays} holidays={holidays} onSave={handleSaveRule} />
      <Toast toast={toast} onClose={() => setToast(null)} />
    </div>
  );
}

function downloadCSV(headings, rows, fileName) {
  const escapeCell = (value) => {
    const text = value instanceof Date ? formatDate(value) : String(value ?? '');
    return `"${text.replace(/"/g, '""')}"`;
  };
  const content = [headings, ...rows].map((row) => row.map(escapeCell).join(';')).join('\r\n');
  const blob = new Blob([`\ufeff${content}`], { type: 'text/csv;charset=utf-8' });
  downloadBlob(blob, fileName);
}

function downloadBlob(blob, fileName) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = fileName;
  link.click();
  window.setTimeout(() => URL.revokeObjectURL(url), 1000);
}

export default App;
