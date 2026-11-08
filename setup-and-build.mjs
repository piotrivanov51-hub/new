/**
 * ПОЛНЫЙ САМОДОСТАТОЧНЫЙ АВТОСБОРЩИК LINELLA ORDER INTELLIGENCE PRO
 * Содержит внутри себя все конфигурационные файлы и весь исходный код src/*
 */
import fs from 'fs';
import path from 'path';

const files = {
  'package.json': JSON.stringify({
    name: 'linella-order-intelligence-pro',
    private: true,
    version: '1.0.0',
    type: 'module',
    scripts: {
      dev: 'vite',
      build: 'vite build',
      preview: 'vite preview'
    },
    dependencies: {
      '@capacitor/android': '^6.0.0',
      '@capacitor/core': '^6.0.0',
      dexie: '^4.0.8',
      'lucide-react': '^0.454.0',
      react: '^18.3.1',
      'react-dom': '^18.3.1',
      xlsx: '^0.18.5'
    },
    devDependencies: {
      '@capacitor/cli': '^6.0.0',
      '@types/react': '^18.3.11',
      '@types/react-dom': '^18.3.1',
      '@vitejs/plugin-react': '^4.3.3',
      autoprefixer: '^10.4.20',
      postcss: '^8.4.47',
      tailwindcss: '^3.4.14',
      typescript: '^5.6.2',
      vite: '^5.4.10'
    }
  }, null, 2),

  'tsconfig.json': JSON.stringify({
    compilerOptions: {
      target: 'ES2020',
      useDefineForClassFields: true,
      lib: ['ES2020', 'DOM', 'DOM.Iterable'],
      module: 'ESNext',
      skipLibCheck: true,
      moduleResolution: 'bundler',
      allowImportingTsExtensions: true,
      resolveJsonModule: true,
      isolatedModules: true,
      noEmit: true,
      jsx: 'react-jsx',
      strict: false
    },
    include: ['src']
  }, null, 2),

  'vite.config.ts': `import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

export default defineConfig({
  plugins: [react()],
  build: {
    outDir: 'dist',
    sourcemap: false
  }
});
`,

  'postcss.config.js': `export default {
  plugins: {
    tailwindcss: {},
    autoprefixer: {},
  },
};
`,

  'tailwind.config.js': `export default {
  content: ['./index.html', './src/**/*.{js,ts,jsx,tsx}'],
  theme: { extend: {} },
  plugins: [],
};
`,

  'capacitor.config.ts': `const config = {
  appId: 'md.linella.orderintelligence.pro',
  appName: 'Linella Order PRO',
  webDir: 'dist',
  server: { androidScheme: 'https' }
};
export default config;
`,

  'index.html': `<!DOCTYPE html>
<html lang="ru">
  <head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no" />
    <title>Linella Order Intelligence PRO</title>
  </head>
  <body class="bg-slate-950 text-slate-100 antialiased">
    <div id="root"></div>
    <script type="module" src="/src/main.tsx"></script>
  </body>
</html>
`,

  'src/index.css': `@tailwind base;
@tailwind components;
@tailwind utilities;
`,

  'src/main.tsx': `import React from 'react';
import ReactDOM from 'react-dom/client';
import App from './App';
import './index.css';

ReactDOM.createRoot(document.getElementById('root')!).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>
);
`,

  'src/App.tsx': `import React, { useEffect, useMemo, useState } from 'react';
import * as XLSX from 'xlsx';
import Dexie, { Table } from 'dexie';
import {
  AlertTriangle, BarChart3, CheckCircle2, Database, Download,
  FileSpreadsheet, Filter, Moon, PackageCheck, Search,
  ShieldAlert, Sun, TrendingUp, Upload, UserCheck
} from 'lucide-react';

// --- БАЗА ДАННЫХ INDEXEDDB ---
export interface ProductSKU {
  id: string;
  barcode: string;
  skuCode: string;
  name: string;
  category: string;
  supplier: string;
  unit: string;
  packSize: number;
  purchasePrice: number;
  shelfLifeDays: number | null;
  currentStock: number | null;
}

export interface OrderRecord {
  id?: number;
  productId: string;
  barcode: string;
  productName: string;
  orderDate: string;
  qty: number;
  unit: string;
  price: number;
  supplier: string;
  category: string;
}

class LinellaDB extends Dexie {
  products!: Table<ProductSKU, string>;
  orders!: Table<OrderRecord, number>;
  constructor() {
    super('LinellaOrderProDB');
    this.version(1).stores({
      products: 'id, barcode, skuCode, name, category, supplier',
      orders: '++id, productId, barcode, orderDate, supplier, category'
    });
  }
}
const db = new LinellaDB();

// --- УТИЛИТЫ ИМПОРТА И ОЧИСТКИ ---
function sanitizeBarcode(val: unknown): string {
  if (val === null || val === undefined) return '';
  if (typeof val === 'number') return Math.round(val).toLocaleString('fullwide', { useGrouping: false });
  const str = String(val).trim();
  if (/^[0-9]+(\\.[0-9]+)?[eE]\\+[0-9]+$/.test(str)) {
    const n = Number(str);
    if (!Number.isNaN(n)) return Math.round(n).toLocaleString('fullwide', { useGrouping: false });
  }
  return str.replace(/\\.0+$/, '').replace(/[^a-zA-Z0-9]/g, '');
}

function parseExcelDate(val: unknown): string {
  if (!val) return new Date().toISOString().slice(0, 10);
  if (typeof val === 'number') {
    const d = new Date(Math.floor(val - 25569) * 86400 * 1000);
    if (!Number.isNaN(d.getTime())) return d.toISOString().slice(0, 10);
  }
  const str = String(val).trim();
  const dmy = str.match(/^(\\d{1,2})[./-](\\d{1,2})[./-](\\d{4})/);
  if (dmy) return \`\${dmy[3]}-\${dmy[2].padStart(2, '0')}-\${dmy[1].padStart(2, '0')}\`;
  return str.slice(0, 10);
}

function roundToPack(rawQty: number, pack: number): number {
  if (rawQty <= 0) return 0;
  const p = Math.max(1, Math.round(pack || 1));
  return Math.ceil(rawQty / p) * p;
}

export default function App() {
  const [darkMode, setDarkMode] = useState(true);
  const [activeTab, setActiveTab] = useState<'DASHBOARD' | 'SMART_ORDER' | 'HISTORY' | 'PATTERNS' | 'ABC' | 'IMPORT'>('DASHBOARD');
  const [products, setProducts] = useState<ProductSKU[]>([]);
  const [orders, setOrders] = useState<OrderRecord[]>([]);
  const [search, setSearch] = useState('');
  const [supplierFilter, setSupplierFilter] = useState('ALL');
  const [manualQty, setManualQty] = useState<Record<string, number>>({});
  const [excluded, setExcluded] = useState<Record<string, boolean>>({});
  const [statusMsg, setStatusMsg] = useState<string | null>(null);

  const loadDB = async () => {
    const [p, o] = await Promise.all([db.products.toArray(), db.orders.toArray()]);
    setProducts(p);
    setOrders(o);
  };

  useEffect(() => { loadDB(); }, []);

  // Демо-данные для проверки всех модулей
  const loadSyntheticDemo = async () => {
    const demoProducts: ProductSKU[] = [
      { id: 'BC_4840012001015', barcode: '4840012001015', skuCode: 'SYN-01', name: '[ТЕСТ] Молоко 2.5% 1л', category: 'Молочные продукты', supplier: 'JLC (Тест)', unit: 'шт', packSize: 12, purchasePrice: 18.5, shelfLifeDays: 7, currentStock: 10 },
      { id: 'BC_4840024005519', barcode: '4840024005519', skuCode: 'SYN-02', name: '[ТЕСТ] Хлеб Тостовый 500г', category: 'Хлебобулочные', supplier: 'Франзелуца (Тест)', unit: 'шт', packSize: 10, purchasePrice: 11.2, shelfLifeDays: 5, currentStock: 75 },
      { id: 'BC_4840088003121', barcode: '4840088003121', skuCode: 'SYN-03', name: '[ТЕСТ] Вода Негазированная 1.5л', category: 'Напитки', supplier: 'Rusnac (Тест)', unit: 'бут', packSize: 6, purchasePrice: 9.8, shelfLifeDays: 365, currentStock: 8 },
      { id: 'BC_5449000000996', barcode: '5449000000996', skuCode: 'SYN-04', name: '[ТЕСТ] Напиток Классический 2л', category: 'Напитки', supplier: 'Beverage (Тест)', unit: 'бут', packSize: 6, purchasePrice: 24.0, shelfLifeDays: 180, currentStock: null }
    ];
    const demoOrders: OrderRecord[] = [];
    const baseDate = new Date('2026-11-08T12:00:00Z');
    demoProducts.forEach((p, idx) => {
      for (let d = 30; d >= 0; d -= 3) {
        const dt = new Date(baseDate.getTime() - d * 86400 * 1000).toISOString().slice(0, 10);
        let q = idx === 0 ? 36 : idx === 1 ? 30 : idx === 2 ? 42 : 24;
        if (idx === 2 && d === 6) q = 160; // Аномальный скачок
        demoOrders.push({ productId: p.id, barcode: p.barcode, productName: p.name, orderDate: dt, qty: q, unit: p.unit, price: p.purchasePrice, supplier: p.supplier, category: p.category });
      }
    });
    await db.transaction('rw', [db.products, db.orders], async () => {
      await db.products.clear();
      await db.orders.clear();
      await db.products.bulkPut(demoProducts);
      await db.orders.bulkAdd(demoOrders);
    });
    await loadDB();
    setStatusMsg('Загружен синтетический тестовый набор данных.');
  };

  // Импорт реального Excel / CSV
  const handleExcelUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (!files || files.length === 0) return;
    let importedRows = 0;
    const prodMap = new Map<string, ProductSKU>();
    products.forEach((p) => prodMap.set(p.id, p));
    const newOrders: OrderRecord[] = [];

    for (let f = 0; f < files.length; f++) {
      const buf = await files[f].arrayBuffer();
      const wb = XLSX.read(buf, { type: 'array' });
      for (const sName of wb.SheetNames) {
        const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(wb.Sheets[sName], { defval: '' });
        for (const r of rows) {
          const keys = Object.keys(r);
          const findCol = (syns: string[]) => keys.find((k) => syns.some((s) => k.toLowerCase().includes(s)));
          const nameKey = findCol(['наимен', 'товар', 'номенкл', 'назван', 'denumire', 'produs', 'name']);
          const bcKey = findCol(['штрих', 'шк', 'ean', 'bare', 'barcode']);
          const qtyKey = findCol(['колич', 'кол-во', 'заказ', 'cantitate', 'qty']);
          const dateKey = findCol(['дата', 'период', 'data', 'date']);
          const supKey = findCol(['поставщ', 'контраг', 'furnizor', 'supplier']);
          const catKey = findCol(['катег', 'групп', 'отдел', 'categorie', 'category']);
          const priceKey = findCol(['цена', 'себест', 'pret', 'price']);
          const stockKey = findCol(['остат', 'запас', 'stoc', 'stock']);
          const packKey = findCol(['кратн', 'упак', 'multiplu', 'pack']);

          const name = nameKey ? String(r[nameKey]).trim() : '';
          const barcode = bcKey ? sanitizeBarcode(r[bcKey]) : '';
          if (!name && !barcode) continue;
          if (/^(итого|всего|total)/i.test(name)) continue;

          const id = barcode.length >= 8 ? \`BC_\${barcode}\` : \`NAME_\${name.toLowerCase()}\`;
          const qty = qtyKey ? parseFloat(String(r[qtyKey]).replace(',', '.')) || 0 : 0;
          const price = priceKey ? parseFloat(String(r[priceKey]).replace(',', '.')) || 0 : 0;
          const pack = packKey ? Math.max(1, parseInt(String(r[packKey]), 10) || 1) : 1;
          const stock = stockKey && r[stockKey] !== '' ? parseFloat(String(r[stockKey]).replace(',', '.')) : null;
          const supplier = supKey ? String(r[supKey]).trim() || 'Основной поставщик' : 'Основной поставщик';
          const category = catKey ? String(r[catKey]).trim() || 'Общая' : 'Общая';

          const prev = prodMap.get(id);
          prodMap.set(id, {
            id,
            barcode: barcode || prev?.barcode || '',
            skuCode: prev?.skuCode || id,
            name: name || prev?.name || id,
            category: category !== 'Общая' ? category : prev?.category || 'Общая',
            supplier: supplier !== 'Основной поставщик' ? supplier : prev?.supplier || 'Основной поставщик',
            unit: 'шт',
            packSize: pack > 1 ? pack : prev?.packSize || 1,
            purchasePrice: price > 0 ? price : prev?.purchasePrice || 0,
            shelfLifeDays: prev?.shelfLifeDays ?? null,
            currentStock: stock !== null ? stock : prev?.currentStock ?? null
          });

          if (qty > 0) {
            newOrders.push({
              productId: id,
              barcode: barcode || prev?.barcode || '',
              productName: name || prev?.name || id,
              orderDate: dateKey ? parseExcelDate(r[dateKey]) : new Date().toISOString().slice(0, 10),
              qty,
              unit: 'шт',
              price: price > 0 ? price : prev?.purchasePrice || 0,
              supplier,
              category
            });
            importedRows++;
          }
        }
      }
    }

    await db.transaction('rw', [db.products, db.orders], async () => {
      await db.products.bulkPut(Array.from(prodMap.values()));
      if (newOrders.length > 0) await db.orders.bulkAdd(newOrders);
    });
    await loadDB();
    setStatusMsg(\`Импорт успешно завершен! Добавлено записей заказов: \${importedRows}, товаров в базе: \${prodMap.size}\`);
  };

  // Полная аналитика и расчет Smart Order Assistant
  const analytics = useMemo(() => {
    const byProd = new Map<string, OrderRecord[]>();
    orders.forEach((o) => {
      const arr = byProd.get(o.productId) || [];
      arr.push(o);
      byProd.set(o.productId, arr);
    });

    let totalValue = 0;
    const items = products.map((p) => {
      const list = (byProd.get(p.id) || []).sort((a, b) => a.orderDate.localeCompare(b.orderDate));
      const count = list.length;
      const qtys = list.map((x) => x.qty);
      const sum = qtys.reduce((a, b) => a + b, 0);
      const avg = count > 0 ? sum / count : 0;
      const sortedQ = [...qtys].sort((a, b) => a - b);
      const median = count > 0 ? sortedQ[Math.floor(count / 2)] : 0;
      const max = count > 0 ? sortedQ[count - 1] : 0;
      const variance = count > 1 ? qtys.reduce((a, q) => a + Math.pow(q - avg, 2), 0) / (count - 1) : 0;
      const stdDev = Math.sqrt(variance);
      const cv = avg > 0 ? stdDev / avg : 0;

      // Интервал между заказами
      const dates = Array.from(new Set(list.map((x) => x.orderDate))).sort();
      let interval = 3;
      if (dates.length >= 2) {
        const d1 = new Date(dates[0]).getTime();
        const d2 = new Date(dates[dates.length - 1]).getTime();
        interval = Math.max(1, Math.round((d2 - d1) / (86400000 * (dates.length - 1))));
      }

      // Прогноз спроса (WMA по последним 3 точкам)
      const recent = qtys.slice(-3);
      const wmaOrder = recent.length === 3 ? recent[2] * 0.5 + recent[1] * 0.3 + recent[0] * 0.2 : avg;
      const dailyDemand = interval > 0 ? wmaOrder / interval : 0;
      const forecast3d = Number((dailyDemand * 3).toFixed(1));
      const forecast7d = Number((dailyDemand * 7).toFixed(1));

      // Smart Order Formula
      const coverageDemand = dailyDemand * 4; // 4 дня покрытия
      const safetyStock = coverageDemand * (cv <= 0.25 ? 0.2 : 0.35);
      const targetStock = coverageDemand + safetyStock;
      const hasStock = p.currentStock !== null && p.currentStock !== undefined;
      const rawRec = hasStock ? Math.max(0, targetStock - (p.currentStock || 0)) : coverageDemand;
      const finalRec = roundToPack(rawRec, p.packSize);

      const effectiveQty = excluded[p.id] ? 0 : manualQty[p.id] !== undefined ? manualQty[p.id] : finalRec;
      const estCost = Number((effectiveQty * (p.purchasePrice || 0)).toFixed(2));
      const itemVal = sum * (p.purchasePrice || 1);
      totalValue += itemVal;

      const isSpike = max >= median * 2.4 && median > 0;
      const perishWarn = p.shelfLifeDays && p.shelfLifeDays <= 7 && (p.currentStock || 0) + finalRec > dailyDemand * p.shelfLifeDays
        ? \`Риск списания (срок годности \${p.shelfLifeDays} дн.)\`
        : null;

      return {
        product: p,
        ordersCount: count,
        avgOrder: Number(avg.toFixed(1)),
        medianOrder: Number(median.toFixed(1)),
        maxOrder: max,
        lastDate: dates[dates.length - 1] || '—',
        interval,
        cv: Number(cv.toFixed(2)),
        forecast3d,
        forecast7d,
        finalRec,
        effectiveQty,
        estCost,
        itemVal,
        isSpike,
        perishWarn,
        reason: hasStock
          ? \`Целевой запас: \${targetStock.toFixed(0)}, остаток: \${p.currentStock}, кратность уп.: \${p.packSize}\`
          : \`Рекомендация на основе истории заказов (ср. \${avg.toFixed(1)}). Требуется подтверждение остатков и спроса.\`
      };
    });

    // ABC / XYZ классификация
    items.sort((a, b) => b.itemVal - a.itemVal);
    let cum = 0;
    const enriched = items.map((it) => {
      const share = totalValue > 0 ? (it.itemVal / totalValue) * 100 : 0;
      cum += share;
      const abc = cum <= 80 ? 'A' : cum <= 95 ? 'B' : 'C';
      const xyz = it.cv <= 0.25 ? 'X' : it.cv <= 0.55 ? 'Y' : 'Z';
      return { ...it, abcXyz: \`\${abc}\${xyz}\`, share: Number(share.toFixed(1)) };
    });

    return enriched;
  }, [products, orders, manualQty, excluded]);

  const suppliers = useMemo(() => Array.from(new Set(products.map((p) => p.supplier))).sort(), [products]);

  const filtered = useMemo(() => {
    const q = search.toLowerCase();
    return analytics.filter((a) => {
      const matchSup = supplierFilter === 'ALL' || a.product.supplier === supplierFilter;
      const matchQ = !q || a.product.name.toLowerCase().includes(q) || a.product.barcode.includes(q);
      return matchSup && matchQ;
    });
  }, [analytics, search, supplierFilter]);

  const exportOrderXLSX = () => {
    const rows = filtered
      .filter((x) => x.effectiveQty > 0)
      .map((x, idx) => ({
        '№': idx + 1,
        'Штрихкод': x.product.barcode,
        'Наименование': x.product.name,
        'Поставщик': x.product.supplier,
        'Остаток': x.product.currentStock ?? 'Не указан',
        'Кратность уп.': x.product.packSize,
        'К заказу': x.effectiveQty,
        'Цена (MDL)': x.product.purchasePrice,
        'Сумма (MDL)': x.estCost,
        'Обоснование': x.reason
      }));
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Заказ');
    XLSX.writeFile(wb, \`Linella_Order_\${new Date().toISOString().slice(0, 10)}.xlsx\`);
  };

  return (
    <div className={darkMode ? 'min-h-screen bg-slate-950 text-slate-100' : 'min-h-screen bg-slate-50 text-slate-900'}>
      <header className="border-b border-slate-800 bg-slate-900 px-4 py-3 sticky top-0 z-20">
        <div className="max-w-7xl mx-auto flex flex-wrap items-center justify-between gap-3">
          <div className="flex items-center gap-2.5">
            <div className="w-9 h-9 rounded-lg bg-red-600 flex items-center justify-center font-bold text-white">L</div>
            <div>
              <h1 className="text-base font-bold text-white">LINELLA ORDER INTELLIGENCE PRO</h1>
              <p className="text-[11px] text-slate-400">Автономная система анализа заказов и прогнозирования спроса</p>
            </div>
          </div>
          <div className="flex items-center gap-2">
            <button onClick={loadSyntheticDemo} className="text-xs px-3 py-1.5 rounded-lg bg-amber-500/20 text-amber-300 border border-amber-500/40">
              Загрузить демо-данные (Тест)
            </button>
            <button onClick={() => setDarkMode(!darkMode)} className="p-2 rounded-lg border border-slate-700 text-slate-300">
              {darkMode ? <Sun className="w-4 h-4" /> : <Moon className="w-4 h-4" />}
            </button>
          </div>
        </div>
        <nav className="max-w-7xl mx-auto mt-3 flex flex-wrap gap-1.5">
          {[
            ['DASHBOARD', 'Dashboard', BarChart3],
            ['SMART_ORDER', 'Smart Order (Заказ)', PackageCheck],
            ['HISTORY', 'История заказов', TrendingUp],
            ['PATTERNS', 'Анализ пред. менеджера', UserCheck],
            ['ABC', 'Аномалии и ABC/XYZ', ShieldAlert],
            ['IMPORT', 'Импорт Excel', FileSpreadsheet]
          ].map(([id, label, Icon]: any) => (
            <button
              key={id}
              onClick={() => setActiveTab(id)}
              className={\`flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-medium \${
                activeTab === id ? 'bg-red-600 text-white' : 'bg-slate-800 text-slate-300'
              }\`}
            >
              <Icon className="w-3.5 h-3.5" /> {label}
            </button>
          ))}
        </nav>
      </header>

      <main className="max-w-7xl mx-auto p-4 space-y-4">
        {statusMsg && (
          <div className="p-3 rounded-lg bg-emerald-950/60 border border-emerald-700 text-xs text-emerald-300">
            {statusMsg}
          </div>
        )}

        {/* Поиск и фильтр */}
        <div className="flex flex-wrap gap-3 p-3 rounded-xl bg-slate-900/60 border border-slate-800">
          <div className="relative flex-1 min-w-[200px]">
            <Search className="w-4 h-4 text-slate-400 absolute left-3 top-2" />
            <input
              value={search}
              onChange={(e) => setSearch(e.target.value)}
              placeholder="Поиск по товару или штрихкоду..."
              className="w-full pl-9 pr-3 py-1 rounded-lg bg-slate-950 border border-slate-700 text-xs"
            />
          </div>
          <select
            value={supplierFilter}
            onChange={(e) => setSupplierFilter(e.target.value)}
            className="px-3 py-1 rounded-lg bg-slate-950 border border-slate-700 text-xs"
          >
            <option value="ALL">Все поставщики ({suppliers.length})</option>
            {suppliers.map((s) => <option key={s} value={s}>{s}</option>)}
          </select>
        </div>

        {activeTab === 'DASHBOARD' && (
          <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
            <div className="p-4 rounded-xl bg-slate-900 border border-slate-800">
              <div className="text-xs text-slate-400">Всего товаров (SKU)</div>
              <div className="text-2xl font-bold mt-1">{products.length}</div>
            </div>
            <div className="p-4 rounded-xl bg-slate-900 border border-slate-800">
              <div className="text-xs text-slate-400">Записей заказов в БД</div>
              <div className="text-2xl font-bold mt-1 text-blue-400">{orders.length}</div>
            </div>
            <div className="p-4 rounded-xl bg-slate-900 border border-slate-800">
              <div className="text-xs text-slate-400">Рекомендовано к заказу</div>
              <div className="text-2xl font-bold mt-1 text-emerald-400">
                {analytics.filter((x) => x.effectiveQty > 0).length}
              </div>
            </div>
            <div className="p-4 rounded-xl bg-slate-900 border border-slate-800">
              <div className="text-xs text-slate-400">Сумма заказа (MDL)</div>
              <div className="text-2xl font-bold mt-1 text-amber-400">
                {analytics.reduce((a, b) => a + b.estCost, 0).toFixed(0)} MDL
              </div>
            </div>
          </div>
        )}

        {activeTab === 'SMART_ORDER' && (
          <div className="space-y-3">
            <div className="flex justify-between items-center">
              <h2 className="text-sm font-bold">Рекомендуемый заказ (Smart Order Assistant)</h2>
              <button onClick={exportOrderXLSX} className="px-3 py-1.5 rounded-lg bg-emerald-600 text-white text-xs font-semibold flex items-center gap-1.5">
                <Download className="w-4 h-4" /> Экспорт заказа в Excel
              </button>
            </div>
            <div className="overflow-x-auto rounded-xl border border-slate-800 bg-slate-900/40">
              <table className="w-full text-left text-xs">
                <thead className="bg-slate-900 text-slate-400 border-b border-slate-800">
                  <tr>
                    <th className="p-2.5">Искл.</th>
                    <th className="p-2.5">Товар / ШК</th>
                    <th className="p-2.5">Остаток</th>
                    <th className="p-2.5">Прогноз (3/7д)</th>
                    <th className="p-2.5">Упак.</th>
                    <th className="p-2.5">Рек. заказ</th>
                    <th className="p-2.5">К заказу</th>
                    <th className="p-2.5">Причина</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-800">
                  {filtered.map((item) => (
                    <tr key={item.product.id}>
                      <td className="p-2.5">
                        <input
                          type="checkbox"
                          checked={Boolean(excluded[item.product.id])}
                          onChange={(e) => setExcluded({ ...excluded, [item.product.id]: e.target.checked })}
                        />
                      </td>
                      <td className="p-2.5">
                        <div className="font-semibold">{item.product.name}</div>
                        <div className="text-slate-400">{item.product.barcode} • {item.product.supplier}</div>
                      </td>
                      <td className="p-2.5">{item.product.currentStock ?? 'Неизв.'}</td>
                      <td className="p-2.5">{item.forecast3d} / {item.forecast7d}</td>
                      <td className="p-2.5">{item.product.packSize}</td>
                      <td className="p-2.5 font-bold text-emerald-400">{item.finalRec}</td>
                      <td className="p-2.5">
                        <input
                          type="number"
                          value={item.effectiveQty}
                          onChange={(e) => setManualQty({ ...manualQty, [item.product.id]: Number(e.target.value) })}
                          className="w-16 px-2 py-1 rounded bg-slate-950 border border-slate-700"
                        />
                      </td>
                      <td className="p-2.5 text-slate-300">
                        {item.reason}
                        {item.perishWarn && <div className="text-amber-400 font-semibold">{item.perishWarn}</div>}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {activeTab === 'HISTORY' && (
          <div className="overflow-x-auto rounded-xl border border-slate-800 bg-slate-900/40">
            <table className="w-full text-left text-xs">
              <thead className="bg-slate-900 text-slate-400 border-b border-slate-800">
                <tr>
                  <th className="p-2.5">Товар</th>
                  <th className="p-2.5">Штрихкод</th>
                  <th className="p-2.5">Заказов</th>
                  <th className="p-2.5">Среднее количество</th>
                  <th className="p-2.5">Максимум</th>
                  <th className="p-2.5">Последний заказ</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-800">
                {filtered.map((item) => (
                  <tr key={item.product.id}>
                    <td className="p-2.5 font-medium">{item.product.name}</td>
                    <td className="p-2.5 font-mono">{item.product.barcode}</td>
                    <td className="p-2.5">{item.ordersCount}</td>
                    <td className="p-2.5">{item.avgOrder} (мед. {item.medianOrder})</td>
                    <td className="p-2.5">{item.maxOrder}</td>
                    <td className="p-2.5">{item.lastDate}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}

        {activeTab === 'PATTERNS' && (
          <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
            {filtered.map((item) => (
              <div key={item.product.id} className="p-3.5 rounded-xl bg-slate-900 border border-slate-800 text-xs space-y-1">
                <div className="font-bold text-sm">{item.product.name}</div>
                <div className="text-slate-400">Интервал заказов: ~{item.interval} дн. • Вариация CV: {(item.cv * 100).toFixed(0)}%</div>
                <div className="text-emerald-400">
                  Наблюдение: {item.cv <= 0.25 ? 'Стабильный регулярный заказ' : 'Переменный объем партии'} (Всего заказов: {item.ordersCount})
                </div>
              </div>
            ))}
          </div>
        )}

        {activeTab === 'ABC' && (
          <div className="overflow-x-auto rounded-xl border border-slate-800 bg-slate-900/40">
            <table className="w-full text-left text-xs">
              <thead className="bg-slate-900 text-slate-400 border-b border-slate-800">
                <tr>
                  <th className="p-2.5">Товар</th>
                  <th className="p-2.5">Класс ABC/XYZ</th>
                  <th className="p-2.5">Доля оборота</th>
                  <th className="p-2.5">Аномалии</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-800">
                {filtered.map((item) => (
                  <tr key={item.product.id}>
                    <td className="p-2.5 font-medium">{item.product.name}</td>
                    <td className="p-2.5 font-bold text-indigo-400">{item.abcXyz}</td>
                    <td className="p-2.5">{item.share}%</td>
                    <td className="p-2.5">
                      {item.isSpike ? (
                        <span className="text-red-400 font-semibold">Обнаружен всплеск: макс. заказ {item.maxOrder} при медиане {item.medianOrder}</span>
                      ) : (
                        <span className="text-slate-400">Без резких аномалий</span>
                      )}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}

        {activeTab === 'IMPORT' && (
          <div className="p-5 rounded-xl bg-slate-900 border border-slate-800 space-y-3">
            <h2 className="text-sm font-bold flex items-center gap-2">
              <Upload className="w-4 h-4 text-red-400" /> Импорт исторических Excel-файлов (.xlsx, .xls, .csv)
            </h2>
            <p className="text-xs text-slate-400">
              Выберите один или несколько файлов заказов предыдущего менеджера. Система автоматически распознает столбцы (Штрихкод, Наименование, Количество, Дата, Поставщик, Остаток).
            </p>
            <input
              type="file"
              accept=".xlsx,.xls,.csv"
              multiple
              onChange={handleExcelUpload}
              className="text-xs file:mr-3 file:py-2 file:px-4 file:rounded-lg file:border-0 file:bg-red-600 file:text-white cursor-pointer"
            />
          </div>
        )}
      </main>
    </div>
  );
}
`
};

for (const [rel, content] of Object.entries(files)) {
  const full = path.join(process.cwd(), rel);
  fs.mkdirSync(path.dirname(full), { recursive: true });
  fs.writeFileSync(full, content, 'utf8');
}
console.log('✅ Все исходники и конфигурация успешно сгенерированы!');
`
