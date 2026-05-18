import { useState, useEffect } from 'react';
import { motion, AnimatePresence } from 'framer-motion';
import Papa from 'papaparse';

interface Product {
  asin: string;
  sku: string;
  title: string;
  roi: number;
  stock_amz: number;
  velocity: number;
  days_left: number;
  amazon_rec: number;
  final_rec: number;
  order_limit: number;
  supp_stock: number;
  reserved: number;
  sent_to_fba: number;
  effective_stock: number;
  provider: string;
  status: 'critical' | 'warning' | 'ok';
  sales_7?: number;
  sales_365?: number;
  sales_60?: number;
  is_back_in_stock?: boolean;
  is_slow_moving?: boolean;
}

interface FBMRecommendation {
  sku: string;
  asin: string;
  title?: string;
  sales_7?: number;
  sales_365: number;
  type: string;
}

interface Summary {
  total_skus: number;
  critical_count: number;
  warning_count: number;
  out_of_supplier_stock: number;
  last_update: string;
  providers: {
    [key: string]: number;
  };
}

interface Data {
  summary: Summary;
  products: Product[];
  fbm_recommendations?: FBMRecommendation[];
}

const PROVIDERS = ['All', 'Signes', 'Minerales', 'Dcasa', 'Trediser'];

const getEffectiveStock = (product: Product) => (
  product.effective_stock ?? (product.stock_amz + product.sent_to_fba + product.reserved)
);

type ExportRow = Record<string, string | number>;

function loadSetFromLocalStorage(key: string) {
  const saved = localStorage.getItem(key);
  if (!saved) return new Set<string>();

  try {
    return new Set<string>(JSON.parse(saved));
  } catch (e) {
    console.error(`Error loading ${key}`, e);
    return new Set<string>();
  }
}

function App() {
  const [data, setData] = useState<Data | null>(null);
  const [activeProvider, setActiveProvider] = useState<string>('All');
  const [searchTerm, setSearchTerm] = useState('');
  const [isSyncing, setIsSyncing] = useState(false);
  const [filterOpen, setFilterOpen] = useState(false);
  const [activeTab, setActiveTab] = useState<'reordering' | 'dashboard' | 'inventory' | 'recommendations' | 'slow_moving' | 'settings'>('reordering');
  const [inventoryFilter, setInventoryFilter] = useState<'none' | 'out_of_stock' | 'back_in_stock' | 'critical'>('none');

  // Sorting state
  const [sortConfig, setSortConfig] = useState<{ column: string; direction: 'asc' | 'desc' } | null>(null);

  // Column filters state
  const [columnFilters, setColumnFilters] = useState<{ [key: string]: string }>({});
  
  // Excluded SKUs (No restock)
  const [excludedSkus, setExcludedSkus] = useState<Set<string>>(() => loadSetFromLocalStorage('excluded_skus'));
  
  // Discarded recommendations
  const [discardedRecommendations, setDiscardedRecommendations] = useState<Set<string>>(() => loadSetFromLocalStorage('discarded_recommendations'));

  const toggleExclude = (sku: string) => {
    setExcludedSkus(prev => {
      const next = new Set(prev);
      if (next.has(sku)) next.delete(sku);
      else next.add(sku);
      localStorage.setItem('excluded_skus', JSON.stringify(Array.from(next)));
      return next;
    });
  };

  const toggleDiscardRecommendation = (sku: string) => {
    setDiscardedRecommendations(prev => {
      const next = new Set(prev);
      if (next.has(sku)) next.delete(sku);
      else next.add(sku);
      localStorage.setItem('discarded_recommendations', JSON.stringify(Array.from(next)));
      return next;
    });
  };

  const handleSort = (column: string) => {
    setSortConfig(prev => {
      if (prev?.column === column) {
        if (prev.direction === 'asc') return { column, direction: 'desc' };
        return null;
      }
      return { column, direction: 'asc' };
    });
  };

  const handleColumnFilter = (column: string, value: string) => {
    setColumnFilters(prev => ({
      ...prev,
      [column]: value
    }));
    setCurrentPage(1); // Reset pagination on filter change
  };

  const [currentPage, setCurrentPage] = useState(1);
  const pageSize = 50;

  useEffect(() => {
    setCurrentPage(1); // Reset on tab change
  }, [activeTab, inventoryFilter, activeProvider]);

  function fetchData() {
    fetch(`/data.json?v=${Date.now()}`, { cache: 'no-store' })
      .then(res => res.json())
      .then(d => {
        setData(d);
      })
      .catch(err => {
        console.error("Failed to load data:", err);
      });
  }

  useEffect(() => {
    fetchData();
  }, []);

  const handleSync = () => {
    setIsSyncing(true);
    setTimeout(() => {
      fetchData();
      setIsSyncing(false);
    }, 2000);
  };

  const exportSlowMoving = () => {
    if (!filteredSlowMoving.length) return;
    
    const exportData: ExportRow[] = filteredSlowMoving.map(p => ({
      SKU: p.sku,
      ASIN: p.asin,
      TITULO: p.title,
      'VENTAS 7D': p.sales_7 || 0,
      'VENTAS 365D': p.sales_365 || 0,
      'VENTAS 60D': p.sales_60 || 0,
      'STOCK FBA+ENV': getEffectiveStock(p),
      'STOCK FBA': p.stock_amz,
      'ROI %': p.roi,
      PROVEEDOR: p.provider
    }));

    const csv = Papa.unparse(exportData, { delimiter: ';' });
    const blob = new Blob(["\ufeff" + csv], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    const url = URL.createObjectURL(blob);
    
    link.setAttribute('href', url);
    link.setAttribute('download', `Slow_Moving_${new Date().toISOString().split('T')[0]}.csv`);
    link.style.visibility = 'hidden';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  const exportCurrentTable = () => {
    if (!filteredProducts.length) return;
    
    let exportData: ExportRow[];
    if (activeTab === 'recommendations') {
      exportData = data?.fbm_recommendations?.map(p => ({
        SKU: p.sku,
        TITULO: p.title || '',
        'VENTAS 7D (FBM)': p.sales_7 || 0,
        'VENTAS 365D (FBM)': p.sales_365,
        RECOMENDACION: 'Mover a FBA'
      })) || [];
    } else {
      exportData = filteredProducts.map(p => ({
        SKU: p.sku,
        ASIN: p.asin,
        TITULO: p.title,
        'VENTAS 7D': p.sales_7 || 0,
        'VENTAS 60D': p.sales_60 || 0,
        'VENTAS 365D': p.sales_365 || 0,
        'STOCK FBA+ENV': getEffectiveStock(p),
        'STOCK FBA': p.stock_amz,
        'TRANSITO': p.sent_to_fba,
        'RESERVADO': p.reserved,
        'VELOCIDAD': p.velocity.toFixed(2),
        'DIAS COBERTURA': p.days_left,
        'STOCK PROV': p.supp_stock,
        'A PEDIR': excludedSkus.has(p.sku) ? 0 : p.final_rec,
        PROVEEDOR: p.provider,
        EXCLUIDO: excludedSkus.has(p.sku) ? 'SI' : 'NO'
      }));
    }

    const csv = Papa.unparse(exportData, { delimiter: ';' });
    const blob = new Blob(["\ufeff" + csv], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    const url = URL.createObjectURL(blob);
    
    link.setAttribute('href', url);
    link.setAttribute('download', `FBA_Replenishment_${activeTab}_${new Date().toISOString().split('T')[0]}.csv`);
    link.style.visibility = 'hidden';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  const filteredProducts = (() => {
    if (!data) return [];
    
    // 1. Initial Filtering
    const filtered = data.products.filter(p => {
      const matchesProvider = activeProvider === 'All' || p.provider === activeProvider;
      const matchesSearch =
        p.sku.toLowerCase().includes(searchTerm.toLowerCase()) ||
        p.title.toLowerCase().includes(searchTerm.toLowerCase()) ||
        p.asin.toLowerCase().includes(searchTerm.toLowerCase());
      
      // Apply the dashboard metric filter
      let matchesMetric = true;
      if (inventoryFilter === 'out_of_stock') matchesMetric = p.supp_stock === 0;
      else if (inventoryFilter === 'back_in_stock') matchesMetric = !!p.is_back_in_stock;
      else if (inventoryFilter === 'critical') matchesMetric = p.days_left < 7;
      
      const isExcluded = excludedSkus.has(p.sku);
      
      if (activeTab === 'reordering') {
        if (isExcluded) return false;
        const effectiveStock = getEffectiveStock(p);
        const hasSupplierStock = p.supp_stock > 0;
        const needsReorder = p.final_rec > 0;
        // Show products that need reordering OR have effective stock
        if (!needsReorder && effectiveStock === 0) return false;
        if (needsReorder && !hasSupplierStock) return false;
        return matchesProvider && matchesSearch && matchesMetric;
      }
      
      return matchesProvider && matchesSearch && matchesMetric;
    });

    // 2. Column Filters
    let columnFiltered = filtered;
    if (Object.keys(columnFilters).length > 0) {
      columnFiltered = filtered.filter(p => {
        for (const [col, value] of Object.entries(columnFilters)) {
          if (!value) continue;
          const lowerValue = value.toLowerCase();
          let cellValue: string | number;
          switch (col) {
            case 'sku': cellValue = p.sku; break;
            case 'title': cellValue = p.title; break;
            case 'asin': cellValue = p.asin; break;
            case 'sales_365': cellValue = p.sales_365 || 0; break;
            case 'stock_amz': cellValue = getEffectiveStock(p); break;
            case 'velocity': cellValue = p.velocity; break;
            case 'days_left': cellValue = p.days_left; break;
            case 'supp_stock': cellValue = p.supp_stock; break;
            case 'final_rec': cellValue = p.final_rec; break;
            case 'roi': cellValue = p.roi; break;
            default: cellValue = '';
          }
          if (typeof cellValue === 'number') {
            if (!cellValue.toString().includes(lowerValue)) return false;
          } else {
            if (!cellValue.toLowerCase().includes(lowerValue)) return false;
          }
        }
        return true;
      });
    }

    // 3. Sorting (if user sorted)
    if (sortConfig) {
      return columnFiltered.sort((a, b) => {
        let aVal: string | number;
        let bVal: string | number;
        switch (sortConfig.column) {
          case 'sku': aVal = a.sku; bVal = b.sku; break;
          case 'title': aVal = a.title; bVal = b.title; break;
          case 'asin': aVal = a.asin; bVal = b.asin; break;
          case 'sales_365': aVal = a.sales_365 || 0; bVal = b.sales_365 || 0; break;
          case 'stock_amz': aVal = getEffectiveStock(a); bVal = getEffectiveStock(b); break;
          case 'velocity': aVal = a.velocity; bVal = b.velocity; break;
          case 'days_left': aVal = a.days_left; bVal = b.days_left; break;
          case 'supp_stock': aVal = a.supp_stock; bVal = b.supp_stock; break;
          case 'final_rec': aVal = a.final_rec; bVal = b.final_rec; break;
          case 'roi': aVal = a.roi; bVal = b.roi; break;
          default: return 0;
        }
        if (typeof aVal === 'number' && typeof bVal === 'number') {
          return sortConfig.direction === 'asc' ? aVal - bVal : bVal - aVal;
        }
        const aStr = String(aVal).toLowerCase();
        const bStr = String(bVal).toLowerCase();
        return sortConfig.direction === 'asc' ? aStr.localeCompare(bStr) : bStr.localeCompare(aStr);
      });
    }

    // 4. Default Priority Sorting (only if no user sort)
    return columnFiltered.sort((a, b) => {
      const aNeeds = a.final_rec > 0;
      const bNeeds = b.final_rec > 0;
      
      if (aNeeds && !bNeeds) return -1;
      if (!aNeeds && bNeeds) return 1;
      
      if (aNeeds && bNeeds) {
        const aZero = getEffectiveStock(a) === 0;
        const bZero = getEffectiveStock(b) === 0;
        
        if (aZero && !bZero) return -1;
        if (!aZero && bZero) return 1;
        
        if (aZero && bZero) {
          return (b.sales_365 || 0) - (a.sales_365 || 0);
        }
        
        return a.days_left - b.days_left;
      }
      
      return 0;
    });
  })();

  const filteredSlowMoving = (() => {
    if (!data) return [];
    
    let filtered = data.products.filter(p => 
      p.is_slow_moving || (p.stock_amz > 0 && (p.sales_60 === 0 || p.days_left > 90))
    );
    
    // Apply column filters
    if (Object.keys(columnFilters).length > 0) {
      filtered = filtered.filter(p => {
        for (const [col, value] of Object.entries(columnFilters)) {
          if (!value) continue;
          const lowerValue = value.toLowerCase();
          let cellValue: string | number;
          switch (col) {
            case 'sku': cellValue = p.sku; break;
            case 'title': cellValue = p.title; break;
            case 'asin': cellValue = p.asin; break;
            case 'sales_365': cellValue = p.sales_365 || 0; break;
            case 'stock_amz': cellValue = p.stock_amz; break;
            case 'roi': cellValue = p.roi; break;
            default: cellValue = '';
          }
          if (typeof cellValue === 'number') {
            if (!cellValue.toString().includes(lowerValue)) return false;
          } else {
            if (!cellValue.toLowerCase().includes(lowerValue)) return false;
          }
        }
        return true;
      });
    }

    // Apply sorting
    if (sortConfig) {
      return filtered.sort((a, b) => {
        let aVal: string | number;
        let bVal: string | number;
        switch (sortConfig.column) {
          case 'sku': aVal = a.sku; bVal = b.sku; break;
          case 'title': aVal = a.title; bVal = b.title; break;
          case 'asin': aVal = a.asin; bVal = b.asin; break;
          case 'sales_365': aVal = a.sales_365 || 0; bVal = b.sales_365 || 0; break;
          case 'stock_amz': aVal = a.stock_amz; bVal = b.stock_amz; break;
          case 'roi': aVal = a.roi; bVal = b.roi; break;
          default: return 0;
        }
        if (typeof aVal === 'number' && typeof bVal === 'number') {
          return sortConfig.direction === 'asc' ? aVal - bVal : bVal - aVal;
        }
        const aStr = String(aVal).toLowerCase();
        const bStr = String(bVal).toLowerCase();
        return sortConfig.direction === 'asc' ? aStr.localeCompare(bStr) : bStr.localeCompare(aStr);
      });
    }

    return filtered;
  })();

  if (!data) {
    return (
      <div className="flex items-center justify-center min-h-screen">
        <div className="flex flex-col items-center gap-4">
          <div className="w-12 h-12 border-4 border-indigo-500 border-t-transparent rounded-full animate-spin"></div>
          <p className="text-slate-400 font-medium animate-pulse text-sm">Iniciando FBA Pulse...</p>
        </div>
      </div>
    );
  }

  if (!data) {
    return (
      <div className="flex items-center justify-center min-h-screen">
        <p className="text-slate-500 p-10">Error loading data. Check console.</p>
      </div>
    );
  }

  return (
    <div className="text-on-background pb-32 font-plus">
      {/* TopAppBar */}
      <header className="fixed top-0 left-0 w-full z-50 flex justify-between items-center px-6 h-16 bg-white/80 backdrop-blur-[20px] border-b border-indigo-500/10 shadow-[0_4px_20px_rgba(79,70,229,0.05)]">
        <div className="flex items-center gap-3">
          <span className="material-symbols-outlined text-indigo-600" style={{ fontVariationSettings: "'FILL' 1" }}>insights</span>
          <h1 className="text-xl font-extrabold tracking-tighter text-indigo-600">FBA Pulse</h1>
        </div>
        <div className="flex items-center gap-4">
          <nav className="hidden md:flex gap-1 mr-4">
            {(['dashboard', 'inventory', 'reordering', 'recommendations', 'slow_moving', 'settings'] as const).map(tab => (
              <button
                key={tab}
                onClick={() => { setActiveTab(tab); setInventoryFilter('none'); }}
                className={`px-3 py-1.5 rounded-lg text-sm font-medium capitalize transition-colors ${
                  activeTab === tab
                    ? 'text-indigo-700 bg-indigo-50'
                    : 'text-slate-500 hover:bg-indigo-50/60 hover:text-slate-700'
                }`}
              >
                {tab === 'dashboard' ? 'Panel' : 
                 tab === 'inventory' ? 'Inventario' : 
                 tab === 'reordering' ? 'Reposición' : 
                 tab === 'recommendations' ? 'Recomendaciones' : 
                 tab === 'slow_moving' ? 'Baja Rotación' :
                 'Ajustes'}
              </button>
            ))}
          </nav>
          <div className="w-9 h-9 rounded-full bg-indigo-100 flex items-center justify-center border-2 border-indigo-200 overflow-hidden">
            <span className="material-symbols-outlined text-indigo-500 text-[20px]">person</span>
          </div>
        </div>
      </header>

      <main className="max-w-[1440px] mx-auto px-6 pt-24">
        {/* Hero */}
        <section className="mb-xl">
          <div className="flex flex-col md:flex-row md:items-end justify-between gap-gutter">
            <div>
              <p className="text-label-caps text-indigo-400 mb-2 tracking-widest">REABASTECIMIENTO</p>
              <h2 className="text-h1 text-on-surface mb-2">Gestión de Inventario</h2>
              <p className="text-body-lg text-on-surface-variant max-w-2xl">
                Cruza la demanda real de Amazon con la disponibilidad de tus proveedores en tiempo real.
              </p>
            </div>
            <div className="flex gap-3">
              {(activeTab === 'reordering' || activeTab === 'inventory' || activeTab === 'recommendations') && (
                <button 
                  onClick={exportCurrentTable}
                  className="bg-emerald-600 text-white px-5 py-3 rounded-xl font-bold flex items-center gap-2 hover:bg-emerald-700 transition-all shadow-lg shadow-emerald-200 active:scale-95 text-sm"
                >
                  <span className="material-symbols-outlined text-[20px]">download</span>
                  Exportar Excel
                </button>
              )}
              {/* Filter dropdown */}
              <div className="relative">
                <button
                  onClick={() => setFilterOpen(v => !v)}
                  className="bg-surface-container-highest text-on-surface-variant px-5 py-3 rounded-xl font-semibold flex items-center gap-2 hover:bg-surface-container-high transition-colors active:scale-95 text-sm"
                >
                  <span className="material-symbols-outlined text-[20px]">filter_list</span>
                  {activeProvider === 'All' ? 'Proveedor' : activeProvider}
                  <span className="material-symbols-outlined text-[16px]">expand_more</span>
                </button>
                <AnimatePresence>
                  {filterOpen && (
                    <motion.div
                      initial={{ opacity: 0, y: -8, scale: 0.95 }}
                      animate={{ opacity: 1, y: 0, scale: 1 }}
                      exit={{ opacity: 0, y: -8, scale: 0.95 }}
                      transition={{ duration: 0.15 }}
                      className="absolute top-full right-0 mt-2 bg-white rounded-2xl shadow-xl border border-indigo-100 p-2 min-w-[160px] z-20"
                    >
                      {PROVIDERS.map(p => (
                        <button
                          key={p}
                          onClick={() => { setActiveProvider(p); setFilterOpen(false); }}
                          className={`w-full text-left px-4 py-2 rounded-xl text-sm font-medium transition-colors flex items-center gap-2 ${
                            activeProvider === p
                              ? 'text-indigo-600 bg-indigo-50'
                              : 'text-slate-600 hover:bg-slate-50'
                          }`}
                        >
                          {activeProvider === p && (
                            <span className="material-symbols-outlined text-[16px] text-indigo-500">check</span>
                          )}
                          {p === 'All' ? 'Todos' : p}
                        </button>
                      ))}
                    </motion.div>
                  )}
                </AnimatePresence>
              </div>

              <button
                onClick={handleSync}
                disabled={isSyncing}
                className="bg-primary-container text-on-primary px-5 py-3 rounded-xl font-semibold flex items-center gap-2 shadow-lg shadow-indigo-500/20 active:scale-95 transition-all duration-200 disabled:opacity-60 text-sm"
              >
                <span className={`material-symbols-outlined text-[20px] ${isSyncing ? 'animate-spin' : ''}`}>sync</span>
                {isSyncing ? 'Sincronizando...' : 'Sincronizar'}
              </button>
            </div>
          </div>
        </section>

        {/* Summary Metrics */}
        <div className="grid grid-cols-1 sm:grid-cols-5 gap-gutter mb-xl">
          <MetricCard
            label="PRODUCTOS EN RIESGO"
            value={data.summary.critical_count}
            subtext="Stock menor a 7 días"
            icon="warning"
            accent="error"
            onClick={() => { setActiveTab('reordering'); setInventoryFilter('critical'); }}
          />
          <MetricCard
            label="SIN STOCK PROVEEDOR"
            value={data.summary.out_of_supplier_stock}
            subtext="Sin disponibilidad para reposición"
            icon="inventory_2"
            accent="secondary"
            onClick={() => { setActiveTab('inventory'); setInventoryFilter('out_of_stock'); }}
          />
          <MetricCard
            label="BACK IN STOCK (FBA)"
            value={data.products.filter(p => p.is_back_in_stock).length}
            subtext="Vuelven tras rotura de stock"
            icon="rebase"
            accent="tertiary"
            onClick={() => { setActiveTab('inventory'); setInventoryFilter('back_in_stock'); }}
          />
          <MetricCard
            label="RECOMENDACIONES FBA"
            value={data.fbm_recommendations?.length || 0}
            subtext="Productos FBM con éxito"
            icon="auto_awesome"
            accent="success"
            onClick={() => { setActiveTab('recommendations'); setInventoryFilter('none'); }}
          />
          <MetricCard
            label="SLOW MOVING (FBA)"
            value={data.products.filter(p => p.is_slow_moving).length}
            subtext="Baja rotación / Liquidación"
            icon="trending_down"
            accent="warning"
            onClick={() => { setActiveTab('slow_moving'); setInventoryFilter('none'); }}
          />
          <MetricCard
            label="EXCLUIDOS"
            value={excludedSkus.size}
            subtext="Marcados como no reponer"
            icon="block"
            accent="secondary"
            onClick={() => { setActiveTab('inventory'); setSearchTerm(''); /* Trigger special filter if needed */ }}
          />
        </div>

        {/* Search */}
        {(activeTab === 'reordering' || activeTab === 'inventory') && (
          <div className="mb-8">
            <div className="relative max-w-lg">
              <span className="material-symbols-outlined absolute left-4 top-1/2 -translate-y-1/2 text-slate-400 text-[20px]">search</span>
              <input
                type="text"
                placeholder="Buscar por SKU, ASIN o título..."
                className="w-full bg-white border border-indigo-100 rounded-2xl py-3.5 pl-11 pr-4 text-sm text-on-surface placeholder:text-slate-400 focus:outline-none focus:ring-2 focus:ring-indigo-300/40 shadow-sm transition-all"
                value={searchTerm}
                onChange={e => setSearchTerm(e.target.value)}
              />
              {searchTerm && (
                <button
                  onClick={() => setSearchTerm('')}
                  className="absolute right-4 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600"
                >
                  <span className="material-symbols-outlined text-[18px]">close</span>
                </button>
              )}
            </div>
            {inventoryFilter !== 'none' && (
              <div className="flex items-center gap-2 mt-4">
                <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Filtro activo:</span>
                <button 
                  onClick={() => setInventoryFilter('none')}
                  className="flex items-center gap-1.5 bg-indigo-50 text-indigo-600 px-3 py-1 rounded-full text-xs font-bold border border-indigo-100 hover:bg-indigo-100 transition-colors"
                >
                  {inventoryFilter === 'critical' ? 'Riesgo Crítico' : 
                   inventoryFilter === 'out_of_stock' ? 'Sin Stock Proveedor' : 'Back In Stock'}
                  <span className="material-symbols-outlined text-[14px]">close</span>
                </button>
              </div>
            )}
          </div>
        )}

        {/* Table / Recommendations */}
        <div className="space-y-sm">
          {(activeTab === 'reordering' || activeTab === 'inventory') ? (
            <>
              <div className="flex items-center justify-between px-2 mb-gutter">
                <h3 className="text-h2 text-on-surface">
                  {activeTab === 'reordering' ? 'Análisis de Reabastecimiento' : 'Inventario FBA'}
                </h3>
                <div className="flex items-center gap-2">
                  <Legend color="bg-error" label="Crítico <7d" />
                  <Legend color="bg-orange-400" label="Alerta <15d" />
                  <Legend color="bg-emerald-500" label="OK" />
                  <span className="text-xs text-on-surface-variant font-medium bg-surface-container px-3 py-1 rounded-full ml-2">
                    {filteredProducts.length} SKUs
                  </span>
                </div>
              </div>

              <div className="glass-card rounded-[28px] shadow-[0_12px_40px_rgba(79,70,229,0.06)] overflow-hidden">
                <div className="overflow-x-auto styled-scroll">
                  <table className="w-full text-left border-collapse">
                    <thead>
                      <tr className="bg-surface-container-low border-b border-outline-variant">
                        <TableHeaderSortable column="sku" label="PRODUCTO" sortConfig={sortConfig} onSort={handleSort} columnFilters={columnFilters} onFilter={handleColumnFilter} />
                        <TableHeaderSortable column="sales_60" label="VENTAS" sublabel="7d/60d/365d" align="center" sortConfig={sortConfig} onSort={handleSort} columnFilters={columnFilters} onFilter={handleColumnFilter} />
                        <TableHeaderSortable column="stock_amz" label="STOCK" sublabel="FBA+Env" align="center" sortConfig={sortConfig} onSort={handleSort} columnFilters={columnFilters} onFilter={handleColumnFilter} />
                        <TableHeaderSortable column="velocity" label="COBER" sublabel="Días/Vel" align="center" sortConfig={sortConfig} onSort={handleSort} columnFilters={columnFilters} onFilter={handleColumnFilter} />
                        <TableHeaderSortable column="supp_stock" label="PROV" sublabel="Stock" align="center" sortConfig={sortConfig} onSort={handleSort} columnFilters={columnFilters} onFilter={handleColumnFilter} />
                        <TableHeaderSortable column="final_rec" label="PEDIR" sublabel="Uds" align="center" sortConfig={sortConfig} onSort={handleSort} columnFilters={columnFilters} onFilter={handleColumnFilter} />
                        <th className="px-1 py-2 text-label-caps text-on-surface-variant text-right whitespace-nowrap">ACC</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-outline-variant/30">
                      <AnimatePresence mode="wait">
                        {filteredProducts.slice((currentPage - 1) * pageSize, currentPage * pageSize).map((p, idx) => (
                          <TableRow 
                            key={p.sku} 
                            product={p} 
                            index={idx} 
                            isExcluded={excludedSkus.has(p.sku)}
                            onToggleExclude={() => toggleExclude(p.sku)}
                          />
                        ))}
                      </AnimatePresence>
                      {filteredProducts.length === 0 && (
                        <tr>
                          <td colSpan={7} className="px-gutter py-16 text-center text-on-surface-variant text-sm">
                            <span className="material-symbols-outlined text-4xl block mb-2 text-slate-300">search_off</span>
                            No se encontraron resultados
                          </td>
                        </tr>
                      )}
                    </tbody>
                  </table>
                </div>
                <div className="bg-surface-container-low px-gutter py-3 border-t border-outline-variant flex items-center justify-between">
                  <div className="flex flex-col">
                    <span className="text-xs text-on-surface-variant font-medium">
                      Mostrando {Math.min((currentPage - 1) * pageSize + 1, filteredProducts.length)} - {Math.min(currentPage * pageSize, filteredProducts.length)} de {filteredProducts.length} SKUs
                    </span>
                    <span className="text-[10px] text-on-surface-variant/50">Página {currentPage} de {Math.ceil(filteredProducts.length / pageSize)}</span>
                  </div>
                  <div className="flex gap-2">
                    <button 
                      onClick={() => setCurrentPage(prev => Math.max(1, prev - 1))}
                      disabled={currentPage === 1}
                      className="w-10 h-10 flex items-center justify-center bg-white border border-outline-variant rounded-xl transition-all hover:bg-indigo-50 disabled:opacity-30 disabled:cursor-not-allowed shadow-sm active:scale-90"
                    >
                      <span className="material-symbols-outlined text-[20px]">chevron_left</span>
                    </button>
                    <button 
                      onClick={() => setCurrentPage(prev => Math.min(Math.ceil(filteredProducts.length / pageSize), prev + 1))}
                      disabled={currentPage >= Math.ceil(filteredProducts.length / pageSize)}
                      className="w-10 h-10 flex items-center justify-center bg-white border border-outline-variant rounded-xl transition-all hover:bg-indigo-50 disabled:opacity-30 disabled:cursor-not-allowed shadow-sm active:scale-90"
                    >
                      <span className="material-symbols-outlined text-[20px]">chevron_right</span>
                    </button>
                  </div>
                </div>
              </div>
            </>
          ) : activeTab === 'recommendations' ? (
            <div className="bg-white/60 backdrop-blur-[12px] rounded-[32px] border border-indigo-500/10 shadow-[0_8px_40px_rgba(79,70,229,0.06)] overflow-hidden">
              <div className="overflow-x-auto">
                <table className="w-full text-left border-collapse">
                  <thead>
                    <tr className="bg-surface-container-low border-b border-outline-variant">
                      <th className="px-6 py-4 text-label-caps text-on-surface-variant">SKU PRODUCTO</th>
                      <th className="px-6 py-4 text-label-caps text-on-surface-variant text-center">VENTAS 365D (FBM)</th>
                      <th className="px-6 py-4 text-label-caps text-on-surface-variant text-center">ESTADO</th>
                      <th className="px-6 py-4 text-label-caps text-on-surface-variant text-right">ACCIÓN</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-outline-variant/30">
                    {data.fbm_recommendations
                      ?.filter(rec => !discardedRecommendations.has(rec.sku))
                      ?.map((rec, i) => (
                      <motion.tr
                        key={rec.sku}
                        initial={{ opacity: 0, y: 10 }}
                        animate={{ opacity: 1, y: 0 }}
                        transition={{ delay: i * 0.03 }}
                        className="hover:bg-indigo-50/30 transition-colors"
                      >
                        <td className="px-6 py-4">
                          <div className="flex items-center gap-3">
                            <button
                              onClick={() => toggleDiscardRecommendation(rec.sku)}
                              className="text-slate-300 hover:text-red-500 transition-colors p-1"
                              title="Desestimar recomendación"
                            >
                              <span className="material-symbols-outlined text-[20px]">block</span>
                            </button>
                            <div className="flex flex-col">
                              <div className="flex items-center gap-2 mb-1">
                                <span className="w-1.5 h-1.5 rounded-full bg-emerald-400" />
                                <span className="font-bold text-slate-800">{rec.sku}</span>
                              </div>
                              <span className="text-[11px] text-slate-500 leading-tight whitespace-normal max-w-xl">{rec.title}</span>
                            </div>
                          </div>
                        </td>
                        <td className="px-6 py-4 text-center">
                          <span className="font-black text-indigo-600 text-lg">{rec.sales_365}</span>
                        </td>
                        <td className="px-6 py-4 text-center">
                          <span className="bg-emerald-100 text-emerald-700 text-[10px] font-black px-2.5 py-1 rounded-full uppercase tracking-wider">
                            Éxito en FBM
                          </span>
                        </td>
                        <td className="px-6 py-4 text-right">
                          <button className="bg-indigo-600 text-white px-5 py-2 rounded-xl text-xs font-bold hover:bg-indigo-700 transition-all shadow-md shadow-indigo-100 active:scale-95">
                            Pasar a FBA
                          </button>
                        </td>
                      </motion.tr>
                    ))}
                    {(data.fbm_recommendations?.length || 0) === 0 && (
                      <tr>
                        <td colSpan={4} className="px-6 py-20 text-center text-slate-400">
                          No hay recomendaciones nuevas en este momento.
                        </td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>
            </div>
          ) : activeTab === 'slow_moving' ? (
            <div className="bg-white/60 backdrop-blur-[12px] rounded-[32px] border border-indigo-500/10 shadow-[0_8px_40px_rgba(79,70,229,0.06)] overflow-hidden">
              <div className="p-6 border-b border-indigo-100 flex items-center justify-between bg-white/40">
                <div>
                  <h3 className="text-xl font-bold text-slate-800">Slow Moving (Baja Rotación)</h3>
                  <p className="text-xs text-slate-500 mt-1">Productos con stock pero sin ventas significativas recientes</p>
                </div>
                <button 
                  onClick={exportSlowMoving}
                  className="flex items-center gap-2 bg-emerald-600 text-white px-5 py-2.5 rounded-2xl text-xs font-bold hover:bg-emerald-700 transition-all shadow-lg shadow-emerald-200 active:scale-95"
                >
                  <span className="material-symbols-outlined text-[18px]">download</span>
                  Exportar a Excel
                </button>
              </div>
              <div className="overflow-x-auto">
                <table className="w-full border-collapse">
                  <thead>
                    <tr className="border-b border-outline-variant bg-surface-container-low/50">
                      <TableHeaderSortable column="sku" label="PRODUCTO" sortConfig={sortConfig} onSort={handleSort} columnFilters={columnFilters} onFilter={handleColumnFilter} />
                      <TableHeaderSortable column="sales_60" label="VENTAS" sublabel="7d/60d/365d" align="center" sortConfig={sortConfig} onSort={handleSort} columnFilters={columnFilters} onFilter={handleColumnFilter} />
                      <TableHeaderSortable column="stock_amz" label="STOCK FBA" align="center" sortConfig={sortConfig} onSort={handleSort} columnFilters={columnFilters} onFilter={handleColumnFilter} />
                      <TableHeaderSortable column="roi" label="ROI" align="center" sortConfig={sortConfig} onSort={handleSort} columnFilters={columnFilters} onFilter={handleColumnFilter} />
                      <th className="px-gutter py-4 text-center text-label-caps text-on-surface-variant font-black">ACCIÓN</th>
                    </tr>
                  </thead>
                  <tbody>
                    {filteredSlowMoving.map((p) => (
                      <tr key={p.sku} className="border-b border-outline-variant hover:bg-indigo-50/50 transition-colors">
                        <td className="px-gutter py-4">
                          <div>
                            <p className="font-bold text-sm text-on-surface">{p.sku}</p>
                            <p className="text-[10px] text-on-surface-variant">{p.title}</p>
                          </div>
                        </td>
                        <td className="px-gutter py-4 text-center">
                          <SalesCell
                            sales7={p.sales_7}
                            sales60={p.sales_60}
                            sales365={p.sales_365}
                          />
                        </td>
                        <td className="px-gutter py-4 text-center font-bold">
                          <StockCell
                            stock={p.stock_amz}
                            sentToFba={p.sent_to_fba}
                            reserved={p.reserved}
                          />
                        </td>
                        <td className="px-gutter py-4 text-center text-sm">{p.roi}%</td>
                        <td className="px-gutter py-4 text-center">
                          <button className="bg-orange-100 text-orange-700 text-[10px] font-bold px-3 py-1.5 rounded-full hover:bg-orange-200 transition-colors uppercase">
                            Promocionar
                          </button>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          ) : activeTab === 'dashboard' ? (
            <div className="py-20 text-center">
              <div className="max-w-md mx-auto">
                <span className="material-symbols-outlined text-6xl text-indigo-200 mb-4">analytics</span>
                <h3 className="text-xl font-bold text-slate-800 mb-2">Bienvenido a FBA Pulse</h3>
                <p className="text-slate-500 mb-8">Haz clic en cualquier tarjeta de arriba para filtrar los productos o utiliza las pestañas de navegación.</p>
                <button 
                  onClick={() => setActiveTab('reordering')}
                  className="bg-indigo-600 text-white px-8 py-3 rounded-2xl font-bold hover:bg-indigo-700 transition-all shadow-lg shadow-indigo-100"
                >
                  Ver Reabastecimiento
                </button>
              </div>
            </div>
          ) : (
            <div className="py-20 text-center text-slate-400">
              Próximamente: {activeTab.charAt(0).toUpperCase() + activeTab.slice(1)}
            </div>
          )}
        </div>
      </main>

      {/* Bottom Nav */}
      <nav className="fixed bottom-0 left-0 w-full z-50 flex justify-around items-center px-4 pb-6 pt-3 bg-white/80 backdrop-blur-[20px] rounded-t-[32px] border-t border-indigo-500/10 shadow-[0_-8px_32px_rgba(79,70,229,0.08)]">
        <NavItem icon="dashboard" label="Panel" active={activeTab === 'dashboard'} onClick={() => setActiveTab('dashboard')} />
        <NavItem icon="inventory_2" label="Inventario" active={activeTab === 'inventory'} onClick={() => setActiveTab('inventory')} />
        <NavItem icon="published_with_changes" label="Reposición" active={activeTab === 'reordering'} onClick={() => setActiveTab('reordering')} />
        <NavItem icon="auto_awesome" label="Recomendaciones" active={activeTab === 'recommendations'} onClick={() => setActiveTab('recommendations')} />
        <NavItem icon="trending_down" label="Baja Rotación" active={activeTab === 'slow_moving'} onClick={() => setActiveTab('slow_moving')} />
        <NavItem icon="settings" label="Settings" active={activeTab === 'settings'} onClick={() => setActiveTab('settings')} />
      </nav>
    </div>
  );
}

// ─── Metric Card ──────────────────────────────────────────────────────────────

const accentMap = {
  error: {
    text: 'text-error',
    icon: 'text-error',
    bg: 'bg-error-container/20',
    border: 'border-l-error',
  },
  secondary: {
    text: 'text-secondary',
    icon: 'text-secondary',
    bg: 'bg-secondary-container/20',
    border: 'border-l-secondary-container',
  },
  tertiary: {
    text: 'text-tertiary',
    icon: 'text-tertiary',
    bg: 'bg-tertiary-container/20',
    border: 'border-l-tertiary-container',
  },
  success: {
    text: 'text-emerald-600',
    icon: 'text-emerald-500',
    bg: 'bg-emerald-50',
    border: 'border-l-emerald-500',
  },
  warning: {
    text: 'text-orange-600',
    icon: 'text-orange-500',
    bg: 'bg-orange-50',
    border: 'border-l-orange-500',
  },
};

function MetricCard({
  label,
  value,
  subtext,
  icon,
  accent,
  onClick,
}: {
  label: string;
  value: string | number;
  subtext: string;
  icon: string;
  accent: keyof typeof accentMap;
  onClick?: () => void;
}) {
  const a = accentMap[accent];
  return (
    <motion.button
      whileHover={{ scale: 1.02, translateY: -4 }}
      whileTap={{ scale: 0.98 }}
      onClick={onClick}
      initial={{ opacity: 0, y: 16 }}
      animate={{ opacity: 1, y: 0 }}
      transition={{ duration: 0.4 }}
      className={`glass-card p-md rounded-[28px] shadow-[0_8px_32px_rgba(79,70,229,0.05)] flex items-start justify-between border-l-4 text-left w-full ${a.border} ${onClick ? 'cursor-pointer' : ''}`}
    >
      <div>
        <p className="text-label-caps text-on-surface-variant mb-2">{label}</p>
        <p className={`text-h1 ${a.text}`}>{value}</p>
        <p className="text-body-md text-on-surface-variant mt-1">{subtext}</p>
      </div>
      <div className={`p-3 rounded-2xl ${a.bg}`}>
        <span className={`material-symbols-outlined ${a.icon}`} style={{ fontVariationSettings: "'FILL' 1" }}>
          {icon}
        </span>
      </div>
    </motion.button>
  );
}

// ─── Legend ───────────────────────────────────────────────────────────────────

function Legend({ color, label }: { color: string; label: string }) {
  return (
    <span className="hidden sm:inline-flex items-center gap-1.5 text-[11px] text-on-surface-variant font-medium">
      <span className={`w-2 h-2 rounded-full ${color}`} />
      {label}
    </span>
  );
}

function StockCell({
  stock,
  sentToFba,
  reserved,
  compact = false,
}: {
  stock: number;
  sentToFba: number;
  reserved: number;
  compact?: boolean;
}) {
  const transit = sentToFba;
  return (
    <div className="flex flex-col items-center">
      <span
        className={
          compact
            ? 'text-[12px] font-bold tabular-nums text-on-surface leading-none'
            : 'text-sm font-bold tabular-nums text-on-surface'
        }
      >
        {stock}
      </span>
      {(transit > 0 || reserved > 0) && (
        <div className="mt-0.5 flex items-center justify-center gap-0.5 flex-wrap">
          {transit > 0 && (
            <span
              className={
                compact
                  ? 'text-[8px] font-bold text-indigo-600 bg-indigo-50 px-0.5 rounded border border-indigo-100/50'
                  : 'text-[10px] font-bold text-indigo-600 bg-indigo-50 px-1.5 py-0.5 rounded-full border border-indigo-100/50'
              }
            >
              +{transit}
            </span>
          )}
          {reserved > 0 && (
            <span
              className={
                compact
                  ? 'text-[8px] font-bold text-orange-600 bg-orange-50 px-0.5 rounded border border-orange-100/50'
                  : 'text-[10px] font-bold text-orange-600 bg-orange-50 px-1.5 py-0.5 rounded-full border border-orange-100/50'
              }
            >
              R{reserved}
            </span>
          )}
        </div>
      )}
    </div>
  );
}

function SalesCell({ sales7, sales60, sales365, compact = false }: {
  sales7?: number;
  sales60?: number;
  sales365?: number;
  compact?: boolean;
}) {
  return (
    <div className="flex flex-col items-center gap-0.5">
      <span className={`font-black tabular-nums leading-none ${compact ? 'text-[12px] text-indigo-600' : 'text-sm text-indigo-600'}`}>
        7d {sales7 || 0}
      </span>
      <span className={`tabular-nums leading-none ${compact ? 'text-[10px] text-slate-700' : 'text-xs text-slate-700'}`}>
        60d {sales60 || 0}
      </span>
      <span className={`tabular-nums leading-none ${compact ? 'text-[9px] text-on-surface-variant/60' : 'text-[10px] text-on-surface-variant/60'}`}>
        365d {sales365 || 0}
      </span>
    </div>
  );
}

// ─── Table Row ────────────────────────────────────────────────────────────────

function TableRow({ 
  product, 
  index, 
  isExcluded, 
  onToggleExclude 
}: { 
  product: Product; 
  index: number;
  isExcluded: boolean;
  onToggleExclude: () => void;
}) {
  const isCritical = product.days_left < 7;
  const isWarning = product.days_left >= 7 && product.days_left < 15;
  const outOfStock = product.final_rec > 0 && product.supp_stock === 0;

  return (
    <motion.tr
      initial={{ opacity: 0, y: 8 }}
      animate={{ opacity: 1, y: 0 }}
      transition={{ delay: Math.min(index * 0.025, 0.35), duration: 0.25 }}
      className={`transition-colors border-b border-outline-variant/30 ${
        isExcluded 
          ? 'bg-slate-50/50 grayscale-[0.5] opacity-60' 
          : isCritical
          ? 'bg-red-50/40 hover:bg-red-50/70'
          : isWarning
          ? 'bg-orange-50/30 hover:bg-orange-50/60'
          : 'hover:bg-indigo-500/[0.03]'
      }`}
    >
      {/* SKU */}
      <td className="px-1.5 py-2">
        <div className="flex items-start gap-1.5">
          <button 
            onClick={(e) => { e.stopPropagation(); onToggleExclude(); }}
            className={`flex-shrink-0 transition-all p-1 rounded-lg mt-0.5 ${
              isExcluded 
                ? 'text-indigo-600 bg-indigo-50 border border-indigo-100' 
                : 'text-slate-200 hover:text-error hover:bg-red-50'
            }`}
            title={isExcluded ? "Permitir reabastecimiento" : "No reabastecer"}
          >
            <span className="material-symbols-outlined text-[15px]">
              {isExcluded ? 'settings_backup_restore' : 'block'}
            </span>
          </button>

          <span
            className={`w-0.5 h-5 rounded-full flex-shrink-0 mt-1.5 ${
              isExcluded ? 'bg-slate-300' : isCritical ? 'bg-error' : isWarning ? 'bg-orange-400' : 'bg-emerald-400'
            }`}
          />
          <div className="min-w-0">
            <div className="flex items-center gap-1">
              <p className={`font-bold text-[12px] ${isExcluded ? 'text-slate-400 line-through' : 'text-on-surface'}`}>
                {product.sku}
              </p>
            </div>
            <span className="text-[10px] text-on-surface-variant block leading-tight opacity-70 whitespace-normal">{product.title}</span>
          </div>
        </div>
      </td>

      {/* Ventas Combinadas */}
      <td className="px-1 py-2 text-center">
        <div className="flex flex-col gap-0.5">
          <span className="text-[12px] font-black tabular-nums text-indigo-600 leading-none">
            7d {product.sales_7 || 0}
          </span>
          <span className="text-[10px] tabular-nums text-slate-700 leading-none">
            60d {product.sales_60 || 0}
          </span>
          <span className="text-[9px] tabular-nums text-on-surface-variant/50">
            365d {product.sales_365 || 0}
          </span>
        </div>
      </td>

      {/* Stock AMZ Compacto */}
      <td className="px-1 py-2 text-center">
        <StockCell
          stock={product.stock_amz}
          sentToFba={product.sent_to_fba}
          reserved={product.reserved}
          compact
        />
      </td>

      {/* Cobertura (Velocidad + Días) */}
      <td className="px-1 py-2 text-center">
        <div className="flex flex-col items-center gap-0.5">
          <span className={`text-[11px] font-bold tabular-nums px-1.5 py-0 rounded-full ${
            isCritical ? 'bg-red-100 text-error' : isWarning ? 'bg-orange-100 text-orange-700' : 'bg-emerald-50 text-emerald-700'
          }`}>
            {product.days_left >= 999 ? '∞' : `${product.days_left}d`}
          </span>
          <span className="text-[8px] tabular-nums text-on-surface-variant/40">
            {product.velocity.toFixed(1)}
          </span>
        </div>
      </td>

      {/* Stock proveedor Compacto */}
      <td className="px-1 py-2 text-center">
        <span className="text-[11px] tabular-nums text-on-surface font-medium leading-none">
          {product.supp_stock > 999 ? '99+' : product.supp_stock}
        </span>
      </td>

      {/* Cantidad a pedir */}
      <td className="px-1 py-2 text-center">
        <span className={`text-[12px] font-bold tabular-nums ${
          isExcluded ? 'text-slate-200' : outOfStock ? 'text-slate-300 line-through' : product.final_rec > 0 ? 'text-indigo-600' : 'text-on-surface-variant/20'
        }`}>
          {isExcluded ? '—' : (product.final_rec > 0 ? product.final_rec : '—')}
        </span>
      </td>

      {/* Acción Compacta */}
      <td className="px-1 py-2 text-right">
        <div className="flex items-center justify-end">
          {!isExcluded && product.final_rec > 0 && !outOfStock && (
            <button className="bg-primary text-white w-6 h-6 rounded flex items-center justify-center hover:bg-indigo-700 transition-all active:scale-90">
              <span className="material-symbols-outlined text-[16px]">shopping_cart</span>
            </button>
          )}
        </div>
      </td>
    </motion.tr>
  );
}

// ─── Nav Item ─────────────────────────────────────────────────────────────────

function NavItem({
  icon,
  label,
  active = false,
  onClick,
}: {
  icon: string;
  label: string;
  active?: boolean;
  onClick?: () => void;
}) {
  return (
    <button
      onClick={onClick}
      className={`flex flex-col items-center justify-center px-4 py-1 transition-all duration-200 rounded-2xl ${
        active
          ? 'text-indigo-600 bg-indigo-50/70'
          : 'text-slate-400 hover:text-indigo-400'
      }`}
    >
      <span
        className="material-symbols-outlined"
        style={{ fontVariationSettings: active ? "'FILL' 1" : "'FILL' 0" }}
      >
        {icon}
      </span>
      <span className="text-[11px] font-medium mt-0.5">{label}</span>
    </button>
  );
}

// ─── Sortable Table Header ─────────────────────────────────────────────────────

function TableHeaderSortable({
  column,
  label,
  sublabel,
  align = 'left',
  sortable = true,
  filterable = true,
  sortConfig,
  onSort,
  columnFilters,
  onFilter,
}: {
  column: string;
  label: string;
  sublabel?: string;
  align?: 'left' | 'center' | 'right';
  sortable?: boolean;
  filterable?: boolean;
  sortConfig: { column: string; direction: 'asc' | 'desc' } | null;
  onSort: (column: string) => void;
  columnFilters: { [key: string]: string };
  onFilter: (column: string, value: string) => void;
}) {
  const isActive = sortConfig?.column === column;
  const direction = isActive ? sortConfig.direction : null;

  return (
    <th className={`px-gutter py-3 ${align === 'center' ? 'text-center' : align === 'right' ? 'text-right' : 'text-left'}`}>
      <div className="flex flex-col gap-1">
        <button
          onClick={() => sortable && onSort(column)}
          className={`flex items-center gap-1 text-label-caps text-on-surface-variant whitespace-nowrap hover:text-indigo-600 transition-colors ${sortable ? 'cursor-pointer' : 'cursor-default'}`}
        >
          {label}
          {sublabel && <span className="text-[9px] font-normal text-on-surface-variant/60 normal-case tracking-normal">{sublabel}</span>}
          {sortable && (
            <span className="material-symbols-outlined text-[14px] ml-0.5">
              {isActive ? (direction === 'asc' ? 'arrow_upward' : direction === 'desc' ? 'arrow_downward' : 'unfold_more') : 'unfold_more'}
            </span>
          )}
        </button>
        {filterable && (
          <input
            type="text"
            placeholder="🔍"
            className="w-full bg-white/50 border border-indigo-100 rounded-lg py-1 px-2 text-[10px] text-on-surface placeholder:text-slate-300 focus:outline-none focus:ring-1 focus:ring-indigo-300/40"
            value={columnFilters[column] || ''}
            onChange={(e) => onFilter(column, e.target.value)}
            onClick={(e) => e.stopPropagation()}
          />
        )}
      </div>
    </th>
  );
}

export default App;
