
import * as XLSX from 'xlsx';
import React, { useState, useMemo, useEffect, useRef } from 'react';
import {
  Monitor,
  Flame,
  Wine,
  Receipt,
  Grid,
  Package,
  BookOpen,
  DollarSign,
  BarChart3,
  ClipboardList,
  Trash2,
  Users,
  Settings,
  Search,
  Plus,
  CheckCircle2,
  AlertTriangle,
  Clock,
  Printer,
  ChevronRight,
  Send,
  X,
  CreditCard,
  Banknote,
  LogOut,
  ArrowUpRight,
  Calendar,
  Lock,
  KeyRound,
  ShieldCheck,
  Percent,
  TrendingUp,
  TrendingDown,
  PieChart,
  RefreshCw,
  Eye,
  Sliders,
  Check,
  Layers,
  FileSpreadsheet,
  Coffee,
  Coins,
  Usb,
  Volume2,
  Zap,
  ShoppingBag,
  Upload,
  Menu,
  Edit3,
  LayoutGrid,
  Download,
  Building,
  HardDrive,
  Mail,
  Briefcase,
  FileText
} from 'lucide-react';
import { syncToCloud, subscribeToCloud } from './firebase'; // <-- ADD THIS LINE
import * as pdfjsLib from 'pdfjs-dist';

// Configure the worker for client-side PDF parsing
pdfjsLib.GlobalWorkerOptions.workerSrc = `https://cdnjs.cloudflare.com/ajax/libs/pdf.js/${pdfjsLib.version}/pdf.worker.min.mjs`;

const ROLE_PERMISSIONS = {
  Administrator: ['pos', 'kds', 'bar', 'billing', 'tables', 'stock', 'recipes', 'shifts', 'reports', 'menu_admin', 'accounting','vendor_bills','payroll', 'staff', 'settings'],
  Manager: ['pos', 'kds', 'bar', 'billing', 'tables', 'stock', 'recipes', 'shifts', 'reports', 'menu_admin','payroll', 'settings'],
  Cashier: ['pos', 'billing', 'tables', 'shifts', 'reports'],
  'Kitchen Chef': ['kds', 'recipes', 'stock'],
  Bartender: ['bar', 'recipes', 'stock'],
  Accountant: ['stock', 'reports','payroll', 'accounting'],
  'Floor Server': ['pos', 'tables', 'billing']
};

const INITIAL_STAFF = [
  { id: 'usr_admin', name: 'System Administrator', role: 'Administrator', pin: '2022', avatar: 'SA', email: 'admin@linolicove.com' }
];

const INITIAL_RAW_INVENTORY = [
  { id: 'ing_rice', name: 'Basmati Rice', category: 'Dry Goods', stock: 24500, unit: 'g', cost: 0.25, threshold: 5000 },
  { id: 'ing_seafood_mix', name: 'Prawns & Calamari Mix', category: 'Seafood', stock: 7200, unit: 'g', cost: 1.80, threshold: 1500 },
  { id: 'ing_eggs', name: 'Farm Fresh Eggs', category: 'Dairy & Eggs', stock: 118, unit: 'pcs', cost: 35.00, threshold: 30 },
  { id: 'ing_espresso_beans', name: 'Roasted Arabica Beans', category: 'Beverages', stock: 4320, unit: 'g', cost: 4.50, threshold: 1000 },
  { id: 'ing_milk', name: 'Fresh Whole Milk', category: 'Dairy & Eggs', stock: 11800, unit: 'ml', cost: 0.30, threshold: 2500 },
  { id: 'ing_beef_patty', name: 'Prime Angus Beef Patty', category: 'Meat', stock: 32, unit: 'pcs', cost: 450.00, threshold: 10 },
  { id: 'ing_burger_bun', name: 'Brioche Bun', category: 'Bakery', stock: 38, unit: 'pcs', cost: 65.00, threshold: 12 },
  { id: 'ing_cheddar', name: 'Aged Cheddar Cheese', category: 'Dairy & Eggs', stock: 2100, unit: 'g', cost: 1.20, threshold: 400 },
  { id: 'ing_calamari', name: 'Fresh Reef Calamari', category: 'Seafood', stock: 6500, unit: 'g', cost: 1.95, threshold: 1200 },
  { id: 'ing_rum', name: 'White Rum', category: 'Bar Supplies', stock: 4500, unit: 'ml', cost: 3.20, threshold: 1000 },
  { id: 'ing_lime', name: 'Fresh Lime Juice', category: 'Produce', stock: 3200, unit: 'ml', cost: 0.80, threshold: 500 },
  { id: 'ing_mint', name: 'Garden Fresh Mint', category: 'Produce', stock: 850, unit: 'g', cost: 1.50, threshold: 200 },
  { id: 'ing_soda', name: 'Sparkling Soda Water', category: 'Beverages', stock: 9500, unit: 'ml', cost: 0.15, threshold: 2000 },
  { id: 'ing_lion_lager', name: 'Lion Lager 625ml', category: 'Bar Supplies', stock: 54, unit: 'pcs', cost: 650.00, threshold: 15 }
];

const INITIAL_MENU_ITEMS = [
  {
    id: 'dish_seafood_rice',
    name: 'SEAFOOD FRIED RICE',
    department: 'Kitchen',
    category: 'Rice & Noodles',
    price: 2250.00,
    prepTime: '15m',
    imageUrl: 'https://images.unsplash.com/photo-1512058564366-18510be2db19?auto=format&fit=crop&w=400&q=80',
    description: 'Wok-tossed basmati rice with tiger prawns, fresh calamari, egg & scallions.',
    recipe: [
      { ingredientId: 'ing_rice', amount: 250 },
      { ingredientId: 'ing_seafood_mix', amount: 150 },
      { ingredientId: 'ing_eggs', amount: 1 }
    ]
  },
  {
    id: 'dish_classic_burger',
    name: 'Angus Truffle Burger',
    department: 'Kitchen',
    category: 'Mains & Grills',
    price: 2450.00,
    prepTime: '12m',
    imageUrl: 'https://images.unsplash.com/photo-1568901346375-23c9450c58cd?auto=format&fit=crop&w=400&q=80',
    description: 'Flame-grilled prime beef patty, aged cheddar, caramelized onion on brioche.',
    recipe: [
      { ingredientId: 'ing_beef_patty', amount: 1 },
      { ingredientId: 'ing_burger_bun', amount: 1 },
      { ingredientId: 'ing_cheddar', amount: 30 }
    ]
  },
  {
    id: 'dish_hot_butter_calamari',
    name: 'Hot Butter Calamari',
    department: 'Kitchen',
    category: 'Starters',
    price: 1850.00,
    prepTime: '10m',
    imageUrl: 'https://images.unsplash.com/photo-1599488615731-7e5c2823ff28?auto=format&fit=crop&w=400&q=80',
    description: 'Crispy seasoned calamari tossed with fresh chili butter and leeks.',
    recipe: [
      { ingredientId: 'ing_calamari', amount: 200 }
    ]
  },
  {
    id: 'drink_mojito',
    name: 'Classic Coastal Mojito',
    department: 'Bar',
    category: 'Cocktails',
    price: 1450.00,
    prepTime: '4m',
    imageUrl: 'https://images.unsplash.com/photo-1551024709-8f23befc6f87?auto=format&fit=crop&w=400&q=80',
    description: 'White rum, fresh garden mint, lime wedges, crushed ice & sparkling soda.',
    recipe: [
      { ingredientId: 'ing_rum', amount: 60 },
      { ingredientId: 'ing_lime', amount: 30 },
      { ingredientId: 'ing_mint', amount: 15 },
      { ingredientId: 'ing_soda', amount: 120 }
    ]
  },
  {
    id: 'drink_lion_beer',
    name: 'Lion Lager 625ml',
    department: 'Bar',
    category: 'Beer & Wine',
    price: 950.00,
    prepTime: '1m',
    imageUrl: 'https://images.unsplash.com/photo-1608270586620-248524c67de9?auto=format&fit=crop&w=400&q=80',
    description: 'Crisp chilled local Sri Lankan lager bottle.',
    recipe: [
      { ingredientId: 'ing_lion_lager', amount: 1 }
    ]
  },
  {
    id: 'drink_cappuccino',
    name: 'Café Cappuccino',
    department: 'Bar',
    category: 'Hot Coffee',
    price: 850.00,
    prepTime: '5m',
    imageUrl: 'https://images.unsplash.com/photo-1534778101976-62847782c213?auto=format&fit=crop&w=400&q=80',
    description: 'Silky microfoam over fresh espresso double shot.',
    recipe: [
      { ingredientId: 'ing_espresso_beans', amount: 18 },
      { ingredientId: 'ing_milk', amount: 180 }
    ]
  },
  {
    id: 'drink_espresso',
    name: 'Double Espresso',
    department: 'Bar',
    category: 'Hot Coffee',
    price: 600.00,
    prepTime: '3m',
    imageUrl: 'https://images.unsplash.com/photo-1510591509098-f4fdc6d0ff04?auto=format&fit=crop&w=400&q=80',
    description: 'Double shot of single-origin dark roasted arabica coffee.',
    recipe: [
      { ingredientId: 'ing_espresso_beans', amount: 18 }
    ]
  }
];

const INITIAL_FLOOR_TABLES = [
  { id: 'T-01', name: 'Table 1', zone: 'Indoor Main Hall', capacity: 2, status: 'VACANT', currentOrderRef: null },
  { id: 'T-02', name: 'Table 2', zone: 'Indoor Main Hall', capacity: 4, status: 'VACANT', currentOrderRef: null },
  { id: 'T-03', name: 'Table 3', zone: 'Indoor Main Hall', capacity: 4, status: 'VACANT', currentOrderRef: null },
  { id: 'T-04', name: 'Table 4', zone: 'Deck Lounge', capacity: 6, status: 'VACANT', currentOrderRef: null },
  { id: 'T-05', name: 'Table 5', zone: 'Deck Lounge', capacity: 4, status: 'VACANT', currentOrderRef: null },
  { id: 'BAR-01', name: 'Bar Seat 01', zone: 'Cocktail Counter', capacity: 1, status: 'VACANT', currentOrderRef: null },
  { id: 'BAR-02', name: 'Bar Seat 02', zone: 'Cocktail Counter', capacity: 1, status: 'VACANT', currentOrderRef: null },
  { id: 'VIP-01', name: 'VIP Cabana 1', zone: 'Private Ocean View', capacity: 8, status: 'VACANT', currentOrderRef: null }
];

function usePersistentState(key, initialValue) {
  const [state, setState] = useState(() => {
    try {
      const stored = localStorage.getItem(key);
      return stored !== null ? JSON.parse(stored) : initialValue;
    } catch (e) {
      console.warn(`LocalStorage read error for ${key}:`, e);
      return initialValue;
    }
  });

  useEffect(() => {
    try {
      localStorage.setItem(key, JSON.stringify(state));
    } catch (e) {
      console.warn(`LocalStorage write error for ${key}:`, e);
    }
  }, [key, state]);

  return [state, setState];
}

const getLocalDateStr = (d = new Date()) => {
  const year = d.getFullYear();
  const month = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
};

// Clean text and extract numeric value and optional unit
const parseQtyAndUnit = (rawStr, defaultUnit = 'g') => {
  if (typeof rawStr === 'number') return { qty: rawStr, unit: defaultUnit };
  if (!rawStr) return { qty: 0, unit: defaultUnit };

  const str = String(rawStr).trim();
  const match = str.match(/([0-9]+(?:\.[0-9]+)?)\s*([a-zA-Z]+)?/);
  if (!match) return { qty: parseFloat(str) || 0, unit: defaultUnit };

  const qty = parseFloat(match[1]) || 0;
  let unit = (match[2] || defaultUnit).toLowerCase();
  
  // Normalize units
  if (['gram', 'grams', 'gm', 'g'].includes(unit)) unit = 'g';
  else if (['kilogram', 'kilograms', 'kg', 'kgs'].includes(unit)) unit = 'kg';
  else if (['milliliter', 'milliliters', 'ml'].includes(unit)) unit = 'ml';
  else if (['liter', 'liters', 'l', 'ltr'].includes(unit)) unit = 'l';
  else if (['piece', 'pieces', 'pcs', 'pc', 'nos'].includes(unit)) unit = 'pcs';

  return { qty, unit };
};

const cleanCostValue = (rawCost) => {
  if (typeof rawCost === 'number') return rawCost;
  if (!rawCost) return 0;
  const cleaned = String(rawCost).replace(/[^0-9.]/g, '');
  return parseFloat(cleaned) || 0;
};

// 1. EXCEL (.xlsx, .xls, .csv) INVENTORY PARSER
const extractInventoryFromExcel = async (file) => {
  const data = await file.arrayBuffer();
  const workbook = XLSX.read(data, { type: 'array' });
  const firstSheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[firstSheetName];
  const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });

  if (rows.length < 2) return [];

  // Identify column headers
  let headerIndex = -1;
  let colMap = { name: 0, category: 1, stock: 2, threshold: 3, cost: 4, unit: -1 };

  for (let i = 0; i < Math.min(10, rows.length); i++) {
    const row = rows[i].map(c => String(c).toLowerCase().trim());
    const nIdx = row.findIndex(c => c.includes('ingredient') || c.includes('item') || c.includes('name') || c.includes('material'));
    if (nIdx !== -1) {
      headerIndex = i;
      colMap.name = nIdx;
      colMap.category = row.findIndex(c => c.includes('cat'));
      colMap.stock = row.findIndex(c => c.includes('stock') || c.includes('remain') || c.includes('qty'));
      colMap.threshold = row.findIndex(c => c.includes('thresh') || c.includes('reorder') || c.includes('min') || c.includes('alert'));
      colMap.cost = row.findIndex(c => c.includes('cost') || c.includes('price') || c.includes('rate'));
      colMap.unit = row.findIndex(c => c.includes('unit'));
      break;
    }
  }

  const startIndex = headerIndex !== -1 ? headerIndex + 1 : 1;
  const parsedItems = [];

  for (let i = startIndex; i < rows.length; i++) {
    const r = rows[i];
    const name = String(r[colMap.name !== -1 ? colMap.name : 0] || '').trim();
    if (!name || name.toLowerCase().includes('total') || name.toLowerCase().includes('raw ingredient')) continue;

    const category = colMap.category !== -1 && r[colMap.category] ? String(r[colMap.category]).trim() : 'Dry Goods';
    const explicitUnit = colMap.unit !== -1 && r[colMap.unit] ? String(r[colMap.unit]).trim() : '';

    const rawStock = r[colMap.stock !== -1 ? colMap.stock : 2];
    const rawThresh = r[colMap.threshold !== -1 ? colMap.threshold : 3];
    const rawCost = r[colMap.cost !== -1 ? colMap.cost : 4];

    const stockParsed = parseQtyAndUnit(rawStock, explicitUnit || 'g');
    const threshParsed = parseQtyAndUnit(rawThresh, stockParsed.unit);
    const unitCost = cleanCostValue(rawCost);

    const safeId = `ing_${name.toLowerCase().replace(/[^a-z0-9]/g, '_').slice(0, 15)}_${Date.now().toString().slice(-4)}_${i}`;

    parsedItems.push({
      id: safeId,
      name,
      category: category || 'Dry Goods',
      stock: stockParsed.qty,
      unit: explicitUnit || stockParsed.unit || 'g',
      cost: unitCost,
      threshold: threshParsed.qty || 10
    });
  }

  return parsedItems;
};

// 2. PDF INVENTORY PARSER (Matches the Visual Table Structure)
const extractInventoryFromPDF = async (file) => {
  const arrayBuffer = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
  const tokens = [];

  for (let p = 1; p <= pdf.numPages; p++) {
    const page = await pdf.getPage(p);
    const content = await page.getTextContent();
    for (let i = 0; i < content.items.length; i++) {
      const str = (content.items[i].str || '').trim();
      if (!str || str === '|' || str.startsWith('Page ')) continue;
      tokens.push(str);
    }
  }

  const items = [];
  const knownCategories = ['Dry Goods', 'Dairy & Eggs', 'Meat', 'Poultry', 'Seafood', 'Beverages', 'Bar Supplies', 'Bakery', 'Produce'];

  for (let i = 0; i < tokens.length; i++) {
    const token = tokens[i];
    const catMatch = knownCategories.find(c => c.toLowerCase() === token.toLowerCase());

    if (catMatch && i > 0) {
      const name = tokens[i - 1].replace(/^[|•\-\s]+|[|•\-\s]+$/g, '').trim();
      
      // Look forward for Stock, Threshold and Cost
      let stock = 0;
      let unit = 'g';
      let threshold = 10;
      let cost = 0;

      const fwd1 = tokens[i + 1] || '';
      const fwd2 = tokens[i + 2] || '';
      const fwd3 = tokens[i + 3] || '';
      const fwd4 = tokens[i + 4] || '';

      const sParsed = parseQtyAndUnit(fwd1);
      if (sParsed.qty > 0) {
        stock = sParsed.qty;
        unit = sParsed.unit;
      }

      const tParsed = parseQtyAndUnit(fwd2, unit);
      if (tParsed.qty > 0) {
        threshold = tParsed.qty;
      }

      cost = cleanCostValue(fwd3) || cleanCostValue(fwd4);

      if (name.length >= 2 && !name.toLowerCase().includes('raw ingredient')) {
        const safeId = `ing_${name.toLowerCase().replace(/[^a-z0-9]/g, '_').slice(0, 15)}_${Date.now().toString().slice(-4)}_${items.length}`;
        items.push({
          id: safeId,
          name,
          category: catMatch,
          stock,
          unit,
          cost,
          threshold
        });
      }
    }
  }

  return items;
};

const extractMenuFromPDF = async (file) => {
  const arrayBuffer = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
  const rawTokens = [];

  // 1. Collect all non-empty text tokens across all pages
  for (let p = 1; p <= pdf.numPages; p++) {
    const page = await pdf.getPage(p);
    const content = await page.getTextContent();
    for (let i = 0; i < content.items.length; i++) {
      const str = (content.items[i].str || '').trim();
      // Skip empty fragments and headers/footers
      if (!str || str === '|' || str.startsWith('Page ') || str.includes('Linoli Cove POS Menu Import File')) {
        continue;
      }
      rawTokens.push(str);
    }
  }

  const parsedItems = [];
  const priceRegex = /^(?:Rs\.?|LKR|\$)?\s*([0-9]{3,5}(?:\.[0-9]{2})?)$/i;

  // 2. Scan through tokens finding prices and associating surrounding metadata
  for (let i = 0; i < rawTokens.length; i++) {
    const token = rawTokens[i];
    const match = token.match(priceRegex);

    if (match) {
      const priceVal = parseFloat(match[1]);
      
      // Look back for candidate item name and category
      let name = '';
      let category = 'Main Menu';
      let department = 'Kitchen';

      // Look back 1 and 2 steps
      const prev1 = rawTokens[i - 1] || '';
      const prev2 = rawTokens[i - 2] || '';
      // Look forward 1 step
      const next1 = rawTokens[i + 1] || '';

      // Skip table header row labels
      if (token.toLowerCase().includes('price') || prev1.toLowerCase().includes('item name')) {
        continue;
      }

      // Check whether prev1 is category or item name
      // If prev2 exists and isn't a known structural word, prev2 is likely Name and prev1 is Category
      if (prev2 && !prev2.toLowerCase().includes('item name') && !prev2.toLowerCase().includes('price')) {
        name = prev2;
        category = prev1;
      } else {
        name = prev1;
      }

      // Check if next token designates the department (Kitchen/Bar)
      if (next1 && (next1.toLowerCase() === 'kitchen' || next1.toLowerCase() === 'bar')) {
        department = next1.charAt(0).toUpperCase() + next1.slice(1).toLowerCase();
      }

      // Clean cleanup formatting
      name = name.replace(/^[|•\-\s]+|[|•\-\s]+$/g, '').trim();
      category = category.replace(/^[|•\-\s]+|[|•\-\s]+$/g, '').trim();

      if (name.length >= 2 && !isNaN(priceVal) && priceVal > 0) {
        parsedItems.push({
          id: `dish_${Date.now()}_${Math.floor(Math.random() * 100000)}_${parsedItems.length}`,
          name: name,
          department: department,
          category: category || 'Main Menu',
          price: priceVal,
          prepTime: '15m',
          imageUrl: '',
          description: `Imported: ${category}`
        });
      }
    }
  }

  return parsedItems;
};

export default function App() {
  // Authentication & Navigation
  const [isAuthenticated, setIsAuthenticated] = useState(false);
  const [loginPinInput, setLoginPinInput] = useState('');
  const [loginError, setLoginError] = useState('');
  const [activeTab, setActiveTab] = useState('pos');
  const [reportSubTab, setReportSubTab] = useState('Daily Overview');
  const [sidebarOpen, setSidebarOpen] = useState(false);
  const [posViewMode, setPosViewMode] = useState('grid'); // 'grid' | 'compact' | 'list'
  const [adminMenuCategory, setAdminMenuCategory] = useState('All');

  // Generic Excel exporter using the already installed XLSX package
const exportReportToExcel = (reportTitle, dataRows, filenamePrefix = 'Report') => {
  if (!Array.isArray(dataRows) || dataRows.length === 0) {
    alert('No data available to export for the selected date range.');
    return;
  }
  try {
    const worksheet = XLSX.utils.json_to_sheet(dataRows);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Report Data');
    XLSX.writeFile(workbook, `${filenamePrefix}_${getLocalDateStr()}.xlsx`);
  } catch (err) {
    console.error('Excel Export Error:', err);
    alert('Failed to generate Excel file: ' + err.message);
  }
};
  
  // Persistent collections
  const [staffList, setStaffList] = usePersistentState('linoli_staff_list', INITIAL_STAFF);
  const [currentUser, setCurrentUser] = useState(INITIAL_STAFF[0]);
  const [inventory, setInventory] = usePersistentState('linoli_inventory', INITIAL_RAW_INVENTORY);
  const [menuItems, setMenuItems] = usePersistentState('linoli_menu_items', INITIAL_MENU_ITEMS);
  const [floorTables, setFloorTables] = usePersistentState('linoli_floor_tables', INITIAL_FLOOR_TABLES);
  const [activeOrders, setActiveOrders] = usePersistentState('linoli_active_orders', []);
  const [transactions, setTransactions] = usePersistentState('linoli_transactions', []);
  const [auditLogs, setAuditLogs] = usePersistentState('linoli_audit_logs', []);
  const [cancelledTickets, setCancelledTickets] = usePersistentState('linoli_cancelled_tickets', []);
  // Persistent sequential counters for sequential numbering
  const [seqCounters, setSeqCounters] = usePersistentState('linoli_seq_counters', {
    order: 1,
    invoice: 1,
    cashOut: 1
  });
  const [vendorBills, setVendorBills] = usePersistentState('linoli_vendor_bills', []);
  const [addVendorBillModalOpen, setAddVendorBillModalOpen] = useState(false);
  const [vendorBillForm, setVendorBillForm] = useState({
    invoiceNumber: '',
    vendorName: '',
    category: 'Food & Beverage Supply',
    billDate: getLocalDateStr(),
    dueDate: getLocalDateStr(),
    amount: '',
    paymentMethod: 'BANK_TRANSFER', // 'BANK_TRANSFER' | 'CHEQUE' | 'CREDIT_CARD' | 'ONLINE_PAYMENT'
    paymentStatus: 'PAID', // 'PAID' | 'UNPAID' | 'PARTIAL'
    notes: ''
  });
  const prevVendorBillsRef = useRef('');

  const [accountingPeriod, setAccountingPeriod] = useState('ALL'); // 'ALL' | 'TODAY' | 'THIS_MONTH' | 'LAST_MONTH'

  // Persistent Attendance & Time Clock
  const [attendanceLogs, setAttendanceLogs] = usePersistentState('linoli_attendance_logs', []);
  const [payrollRecords, setPayrollRecords] = usePersistentState('linoli_payroll_records', []);
  const [payrollSubTab, setPayrollSubTab] = useState('attendance'); // 'attendance' | 'payslips' | 'epf_etf' | 'profiles'
  const [editingPayrollId, setEditingPayrollId] = useState(null);

  // Persistent collection for standalone advance disbursements
  const [salaryAdvances, setSalaryAdvances] = usePersistentState('linoli_salary_advances', []);
  const [issueAdvanceModalOpen, setIssueAdvanceModalOpen] = useState(false);
  const [advanceForm, setAdvanceForm] = useState({
    staffId: '',
    period: `${new Date().getFullYear()}-${String(new Date().getMonth() + 1).padStart(2, '0')}`,
    amount: '',
    paymentMethod: 'CASH', // 'CASH' | 'BANK_TRANSFER'
    reason: 'Personal Advance',
    notes: ''
  });
  const prevAdvancesRef = useRef('');
  
  // 1. Declare modal open state
  const [processPayModalOpen, setProcessPayModalOpen] = useState(false);

  const [payrollInputForm, setPayrollInputForm] = useState({
    staffId: '',
    period: `${new Date().getFullYear()}-${String(new Date().getMonth() + 1).padStart(2, '0')}`,
    epfEtfEnabled: true,
    basicSalary: 35000,
    budgetaryAllowance: 2500,
    otherAllowances: 0,
    serviceChargeBonus: 0,
    incentiveBonus: 0,
    overtimeHours: 0,
    overtimeRate: 250,
    salaryAdvance: 0,
    otherDeductions: 0,
    standardWorkingDays: 26,
    workedDays: 26,
    paidLeaves: 0,
    unpaidLeaves: 0,
    holidaysCount: 4,
    shortShiftsCount: 0,
    notes: ''
  });

  // 3. Auto-fill payroll inputs AFTER payrollInputForm is defined
  useEffect(() => {
    if (!payrollInputForm.staffId || !payrollInputForm.period || editingPayrollId) return;

    // 1. Calculate unrecovered advances issued for this staff & month
    const totalAdvancesIssued = (salaryAdvances || [])
      .filter(adv => adv.staffId === payrollInputForm.staffId && adv.period === payrollInputForm.period && adv.status !== 'REJECTED')
      .reduce((sum, adv) => sum + (Number(adv.amount) || 0), 0);

    // 2. Fetch biometric attendance days
    const relevantLogs = (attendanceLogs || []).filter(log => {
      if (!log || log.staffId !== payrollInputForm.staffId) return false;
      const lDate = log.date || (log.timestamp ? extractDateStr(log.timestamp) : '');
      return lDate.startsWith(payrollInputForm.period);
    });

    const uniqueDates = new Set();
    let totalOtHours = 0;
    let shortShifts = 0;

    relevantLogs.forEach(l => {
      if (l.date) uniqueDates.add(l.date);
      const hrs = Number(l.totalHours) || 0;
      if (hrs > 0 && hrs < 5) shortShifts += 1;
      if (hrs > 8) totalOtHours += (hrs - 8);
    });

    const autoWorkedDays = uniqueDates.size > 0 ? uniqueDates.size : 26;

    setPayrollInputForm(prev => ({
      ...prev,
      salaryAdvance: totalAdvancesIssued, // <-- Pre-fills sum of advances issued
      workedDays: autoWorkedDays,
      shortShiftsCount: shortShifts,
      overtimeHours: Number(totalOtHours.toFixed(1))
    }));
  }, [payrollInputForm.staffId, payrollInputForm.period, attendanceLogs, salaryAdvances, editingPayrollId]);

  // Sequential Number Helpers
  const getNextOrderNumber = () => {
    let nextNum = seqCounters.order || 1;
    // Fallback: If existing active orders or transactions exist, ensure no collisions
    const maxActive = activeOrders.reduce((max, o) => {
      const match = (o.orderId || '').match(/ORD-(\d+)/);
      return match ? Math.max(max, parseInt(match[1], 10)) : max;
    }, 0);
    const maxTrans = transactions.reduce((max, t) => {
      const match = (t.orderRef || '').match(/ORD-(\d+)/);
      return match ? Math.max(max, parseInt(match[1], 10)) : max;
    }, 0);
    nextNum = Math.max(nextNum, maxActive + 1, maxTrans + 1);

    setSeqCounters(prev => ({ ...prev, order: nextNum + 1 }));
    return `ORD-${String(nextNum).padStart(5, '0')}`;
  };

  const getNextInvoiceNumber = () => {
    let nextNum = seqCounters.invoice || 1;
    // Fallback: Ensure no collision with recorded invoices
    const maxInv = transactions.reduce((max, t) => {
      const match = (t.invoiceNo || '').match(/INV-(\d+)/);
      return match ? Math.max(max, parseInt(match[1], 10)) : max;
    }, 0);
    nextNum = Math.max(nextNum, maxInv + 1);

    setSeqCounters(prev => ({ ...prev, invoice: nextNum + 1 }));
    return `INV-${String(nextNum).padStart(5, '0')}`;
  };

  const getNextCashOutNumber = () => {
    let nextNum = seqCounters.cashOut || 1;
    const allPayouts = Array.isArray(expenses) ? expenses : [];
    const maxCo = allPayouts.reduce((max, c) => {
      const match = (c.id || '').match(/CO-(\d+)/);
      return match ? Math.max(max, parseInt(match[1], 10)) : max;
    }, 0);
    nextNum = Math.max(nextNum, maxCo + 1);

    setSeqCounters(prev => ({ ...prev, cashOut: nextNum + 1 }));
    return `CO-${String(nextNum).padStart(5, '0')}`;
  };

  // Tracks unsettled cash discrepancy carried across shifts
  const [unsettledVariance, setUnsettledVariance] = usePersistentState('linoli_unsettled_variance', 0);
// Persistent stock movement difference & intake ledger
  const [stockLogs, setStockLogs] = usePersistentState('linoli_stock_logs', []);
  // System Settings
  const [settings, setSettings] = usePersistentState('linoli_system_settings', {
    restaurantName: 'Linoli Cove Midigama',
    tagline: 'RESTAURANT & BAR',
    legalName: 'Linoli Cove Leisure (Pvt) Ltd',
    businessRegNo: 'PV-00289144',
    taxId: 'TIN-109284719',
    terminalId: 'LINOLI-MAIN-01',
    phone: '+94 74 036 6741',
    email: 'info@linolicove.me',
    website: 'www.linolicove.me',
    address: '380 A Matara Road, Midigama, 81700',
    currency: 'Rs.',
    serviceChargeRate: 10,
    taxRate: 8,
    receiptRollWidth: '80mm',
    receiptFontSize: '11px',
    receiptFontFamily: 'monospace',
    receiptMargin: '2mm',
    autoPrintOrder: true,
    autoPrintBill: true,
    autoDrawerKick: 'ENABLED',
    drawerKickTrigger: 'CASH_ONLY',
    drawerPinout: 'PIN_2',
    chimeAudio: true,
    receiptHeader: 'Linoli Cove Beach Resort & Dining\nBeach Road, Midigama\nTel: +94 74 036 6741',
    receiptFooter: 'Thank you for your visit!\nPlease come again.'
  });

  // Shifts state
  const [currentShift, setCurrentShift] = usePersistentState('linoli_current_shift', {
    shiftId: `SHIFT-${getLocalDateStr().replace(/-/g, '')}-01`,
    openedDate: getLocalDateStr(),
    openedAt: '09:00 AM',
    openedBy: 'System Administrator',
    startingFloat: 0.00,
    status: 'OPEN',
    payouts: []
  });

  // Dedicated persistent running float that preserves cash across shift rollovers
  const [runningFloat, setRunningFloat] = usePersistentState('linoli_running_float', 10000.00);

  const [shiftHistory, setShiftHistory] = usePersistentState('linoli_shift_history', []);
  const [denominations, setDenominations] = usePersistentState('linoli_denominations', {
    5000: 0, 1000: 0, 500: 0, 100: 0, 50: 0, 20: 0
  });
  const [payoutForm, setPayoutForm] = useState({ amount: '', reason: '' });

  const [cashOutForm, setCashOutForm] = useState({
    amount: '',
    category: 'Supplier / Vendor',
    reason: '',
    recipient: ''
  });
  const [cashOutApprovalModal, setCashOutApprovalModal] = useState({
    open: false,
    item: null,
    managerPin: '',
    error: ''
  });

  // POS State
  const [orderMode, setOrderMode] = useState('DINING');
  const [selectedTable, setSelectedTable] = useState(INITIAL_FLOOR_TABLES[0]);
  const [takeawayInfo, setTakeawayInfo] = useState({ name: 'Walk-in Guest', phone: '', token: 'TK-101' });
  const [guestCount, setGuestCount] = useState(2);
  const [cart, setCart] = useState([]);
  const [selectedCategory, setSelectedCategory] = useState('All');
  const [menuSearchQuery, setMenuSearchQuery] = useState('');
  const [serviceChargeActive, setServiceChargeActive] = useState(true);
  const [taxActive, setTaxActive] = useState(false);
  const [discountPercent, setDiscountPercent] = useState(0);
  const [mobileCartDrawerOpen, setMobileCartDrawerOpen] = useState(false);

  // Date Filter State for Reports
  const [reportStartDate, setReportStartDate] = useState(getLocalDateStr());
  const [reportEndDate, setReportEndDate] = useState(getLocalDateStr());

  // Hardware State
  const [pairedUsbDevice, setPairedUsbDevice] = useState(null);
  const [usbStatusMessage, setUsbStatusMessage] = useState('');
  const [settingsNotice, setSettingsNotice] = useState(null);

  const [emailSettings, setEmailSettings] = usePersistentState('linoli_email_settings', {
    enabled: true,
    recipient: 'linolicove@gmail.com',
    scheduledTime: '23:30',
    webhookUrl: '',
    emailjsServiceId: '',
    emailjsTemplateId: '',
    emailjsPublicKey: '',
    lastSentDate: ''
  });
  const [isSendingEmail, setIsSendingEmail] = useState(false);

  // Background Auto-Print State (No blocking modal)
  const [activePrintSlip, setActivePrintSlip] = useState(null);
  const [printNotice, setPrintNotice] = useState(null);
  const [checkoutModalOpen, setCheckoutModalOpen] = useState(false);
  const [settlingOrder, setSettlingOrder] = useState(null);
  const [paymentMethod, setPaymentMethod] = useState('CASH');
  const [cashTendered, setCashTendered] = useState('');
  const [pinModalOpen, setPinModalOpen] = useState(false);
  const [pinInput, setPinInput] = useState('');
  const [targetStaffForSwitch, setTargetStaffForSwitch] = useState(null);
  const [pinError, setPinError] = useState('');
  const [addStaffModalOpen, setAddStaffModalOpen] = useState(false);
  const [editingStaffId, setEditingStaffId] = useState(null);
  const [newStaffForm, setNewStaffForm] = useState({
    name: '',
    role: 'Cashier',
    pin: '',
    email: '',
    basicSalary: 35000,
    budgetaryAllowance: 2500,
    otherAllowances: 0,
    fixedBonus: 0,
    overtimeRate: 250,
    epfEtfEnabled: true
  });
  const [staffFormError, setStaffFormError] = useState('');
  const [addTableModalOpen, setAddTableModalOpen] = useState(false);
  const [newTableForm, setNewTableForm] = useState({ name: '', zone: 'Indoor Main Hall', capacity: 4 });
  const [allocationModalOpen, setAllocationModalOpen] = useState(false);
  const [voidModalOpen, setVoidModalOpen] = useState(false);
  const [voidPayload, setVoidPayload] = useState({ item: null, reason: '' });
  

  // Bill Editing Modal (Billing queue)
  const [editBillModalOpen, setEditBillModalOpen] = useState(false);
  const [editingBill, setEditingBill] = useState(null);

  // Inventory & Recipe Modals
  const [addInventoryModalOpen, setAddInventoryModalOpen] = useState(false);
  const [newInventoryForm, setNewInventoryForm] = useState({ name: '', category: 'Dry Goods', stock: '', unit: 'g', cost: '', threshold: '' });
  const [receiveStockModalOpen, setReceiveStockModalOpen] = useState(false);
  const [editingInventoryItem, setEditingInventoryItem] = useState(null);
  const [editInventoryModalOpen, setEditInventoryModalOpen] = useState(false);
  const [receiveStockForm, setReceiveStockForm] = useState({ ingredientId: '', quantity: '', supplier: '', invoiceRef: '', newCost: '' });
  const [addItemModalOpen, setAddItemModalOpen] = useState(false);
  const [editingMenuItem, setEditingMenuItem] = useState(null); // <-- ADD THIS
  const [isEditModalOpen, setIsEditModalOpen] = useState(false);  // <-- ADD THIS
  const [newDishForm, setNewDishForm] = useState({
    name: '',
    department: 'Kitchen',
    category: 'Rice & Noodles',
    customCategory: '',
    price: '',
    prepTime: '10m',
    description: '',
    imageUrl: '',
    recipeIngredients: []
  });
  const [recipeConfigModalOpen, setRecipeConfigModalOpen] = useState(false);
  const [editingDishForRecipe, setEditingDishForRecipe] = useState(null);
  const [currentRecipeIngredients, setCurrentRecipeIngredients] = useState([]);
  const [tempIngredientSelect, setTempIngredientSelect] = useState({ ingredientId: '', amount: '' });
  const [editDishBomInput, setEditDishBomInput] = useState({ ingredientId: '', amount: '' });
  const handleAddIngredientToRecipe = (e) => {
    if (e) {
      e.preventDefault();
      e.stopPropagation();
    }

    const ingId = tempIngredientSelect.ingredientId || (inventory[0]?.id || '');
    const amountVal = parseFloat(tempIngredientSelect.amount);

    if (!ingId) {
      alert('Please select an ingredient.');
      return;
    }

    if (isNaN(amountVal) || amountVal <= 0) {
      alert('Please enter a valid amount greater than 0.');
      return;
    }

    setCurrentRecipeIngredients((prev) => {
      const existingIndex = prev.findIndex((r) => r.ingredientId === ingId);
      if (existingIndex >= 0) {
        return prev.map((r, i) =>
          i === existingIndex
            ? { ...r, amount: Number((r.amount + amountVal).toFixed(2)) }
            : r
        );
      }
      return [...prev, { ingredientId: ingId, amount: amountVal }];
    });

    setTempIngredientSelect((prev) => ({ ...prev, amount: '' }));
  };

  const handleRemoveIngredientFromRecipe = (indexToRemove) => {
    setCurrentRecipeIngredients((prev) =>
      prev.filter((_, idx) => idx !== indexToRemove)
    );
  };

  const handleSaveRecipeConfiguration = (e) => {
    if (e) e.preventDefault();
    if (!editingDishForRecipe) return;

    setMenuItems((prev) =>
      prev.map((dish) =>
        dish.id === editingDishForRecipe.id
          ? { ...dish, recipe: currentRecipeIngredients }
          : dish
      )
    );

    recordAuditLog(
      'RECIPE_UPDATED',
      editingDishForRecipe.id,
      `Updated BOM recipe for ${editingDishForRecipe.name} (${currentRecipeIngredients.length} linked raw materials)`
    );

    setRecipeConfigModalOpen(false);
    setEditingDishForRecipe(null);
    setCurrentRecipeIngredients([]);
    setTempIngredientSelect({ ingredientId: '', amount: '' });
  };
  
 // Refs to prevent premature uploads and re-render loops
  const isCloudSynced = useRef(false);
  const prevOrdersRef = useRef('');
  const prevTransRef = useRef('');
  const prevTablesRef = useRef('');
  const prevAuditsRef = useRef('');
  const prevMenuRef = useRef('');
  const prevInventoryRef = useRef(''); // <-- ADD THIS
  const prevExpensesRef = useRef('');  // <-- ADD THIS
  const prevStaffRef = useRef('');
  const prevShiftRef = useRef('');
  const prevAttendanceRef = useRef('');
  const prevPayrollRef = useRef('');
  const prevDenomRef = useRef('');
  const [expenses, setExpenses] = useState(() => {
  try {
    const local = localStorage.getItem('linoli_expenses');
    return local ? JSON.parse(local) : [];
  } catch (e) {
    return [];
  }
  });

  // ============================================================
  // 1. REAL-TIME CLOUD LISTENERS (Download from Firebase)
  // ============================================================
  useEffect(() => {
    // 1. Receive incoming active orders
    const unsubOrders = subscribeToCloud('active_orders', (remoteOrders) => {
      isCloudSynced.current = true;
      if (Array.isArray(remoteOrders)) {
        const serialized = JSON.stringify(remoteOrders);
        if (prevOrdersRef.current === serialized) return;
        prevOrdersRef.current = serialized;
        setActiveOrders(remoteOrders);
        localStorage.setItem('linoli_active_orders', serialized);
      }
    });

    // 2. Receive incoming settled transactions
    const unsubTrans = subscribeToCloud('transactions', (remoteTrans) => {
      isCloudSynced.current = true;
      if (Array.isArray(remoteTrans)) {
        const serialized = JSON.stringify(remoteTrans);
        if (prevTransRef.current === serialized) return;
        prevTransRef.current = serialized;
        setTransactions(remoteTrans);
        localStorage.setItem('linoli_transactions', serialized);
      }
    });

    // 3. Receive incoming table occupancy / floor status
    const unsubTables = subscribeToCloud('floor_tables', (remoteTables) => {
      isCloudSynced.current = true;
      if (Array.isArray(remoteTables)) {
        const serialized = JSON.stringify(remoteTables);
        if (prevTablesRef.current === serialized) return;
        prevTablesRef.current = serialized;
        setFloorTables(remoteTables);
        localStorage.setItem('linoli_floor_tables', serialized);
      }
    });

    // 4. Receive incoming audit trail activity logs
    const unsubAudits = subscribeToCloud('audit_logs', (remoteAudits) => {
      isCloudSynced.current = true;
      if (Array.isArray(remoteAudits)) {
        const serialized = JSON.stringify(remoteAudits);
        if (prevAuditsRef.current === serialized) return;
        prevAuditsRef.current = serialized;
        setAuditLogs(remoteAudits);
        localStorage.setItem('linoli_audit_logs', serialized);
      }
    });

    // 5. Receive incoming menu items
    const unsubMenu = subscribeToCloud('menu_items', (remoteMenu) => {
      isCloudSynced.current = true;
      if (Array.isArray(remoteMenu)) {
        const serialized = JSON.stringify(remoteMenu);
        if (prevMenuRef.current === serialized) return;
        prevMenuRef.current = serialized;
        setMenuItems(remoteMenu);
        localStorage.setItem('linoli_menu_items', serialized);
      }
    });
    // 6. Receive raw inventory adjustments
  const unsubInventory = subscribeToCloud('inventory', (remoteInv) => {
    isCloudSynced.current = true;
    if (Array.isArray(remoteInv)) {
      const serialized = JSON.stringify(remoteInv);
      if (prevInventoryRef.current === serialized) return;
      prevInventoryRef.current = serialized;
      setInventory(remoteInv);
      localStorage.setItem('linoli_inventory', serialized);
    }
  });

  // 7. Receive recorded cash expenses
  const unsubExpenses = subscribeToCloud('expenses', (remoteExp) => {
    isCloudSynced.current = true;
    if (Array.isArray(remoteExp)) {
      const serialized = JSON.stringify(remoteExp);
      if (prevExpensesRef.current === serialized) return;
      prevExpensesRef.current = serialized;
      setExpenses(remoteExp);
      localStorage.setItem('linoli_expenses', serialized);
    }
  }); 
  // 8. Receive incoming cashier shift updates
const unsubShift = subscribeToCloud('current_shift', (remoteShift) => {
  isCloudSynced.current = true;
  if (remoteShift && typeof remoteShift === 'object') {
    const serialized = JSON.stringify(remoteShift);
    if (prevShiftRef.current === serialized) return;
    prevShiftRef.current = serialized;
    setCurrentShift(remoteShift);
    localStorage.setItem('linoli_current_shift', serialized);
  }
});
  // 8. Receive incoming staff list
    const unsubStaff = subscribeToCloud('staff_list', (remoteStaff) => {
      isCloudSynced.current = true;
      if (Array.isArray(remoteStaff) && remoteStaff.length > 0) {
        const serialized = JSON.stringify(remoteStaff);
        if (prevStaffRef.current === serialized) return;
        prevStaffRef.current = serialized;
        setStaffList(remoteStaff);
        localStorage.setItem('linoli_staff_list', serialized);
      }
    });
    // 9. Receive physical denomination counts from cloud
    const unsubDenominations = subscribeToCloud('denominations', (remoteDenom) => {
      isCloudSynced.current = true;
      if (remoteDenom && typeof remoteDenom === 'object') {
        const serialized = JSON.stringify(remoteDenom);
        if (prevDenomRef.current === serialized) return;
        prevDenomRef.current = serialized;
        setDenominations(remoteDenom);
        localStorage.setItem('linoli_denominations', serialized);
      }
    });

    // 10. Receive attendance logs from cloud in real time
    const unsubAttendance = subscribeToCloud('attendance_logs', (remoteAtt) => {
      isCloudSynced.current = true;
      if (Array.isArray(remoteAtt)) {
        const serialized = JSON.stringify(remoteAtt);
        if (prevAttendanceRef.current === serialized) return;
        prevAttendanceRef.current = serialized;
        setAttendanceLogs(remoteAtt);
        localStorage.setItem('linoli_attendance_logs', serialized);
      }
    });

    // 11. Receive processed payroll records from cloud in real time
    const unsubPayroll = subscribeToCloud('payroll_records', (remotePay) => {
      isCloudSynced.current = true;
      if (Array.isArray(remotePay)) {
        const serialized = JSON.stringify(remotePay);
        if (prevPayrollRef.current === serialized) return;
        prevPayrollRef.current = serialized;
        setPayrollRecords(remotePay);
        localStorage.setItem('linoli_payroll_records', serialized);
      }
    });
    
    // 12. Receive accounting settings from cloud
    const unsubAccounting = subscribeToCloud('accounting_settings', (remoteSettings) => {
      isCloudSynced.current = true;
      if (remoteSettings && remoteSettings.period) {
        setAccountingPeriod(remoteSettings.period);
      }
    });
    // Receive external vendor bills & invoices from cloud
    const unsubVendorBills = subscribeToCloud('vendor_bills', (remoteBills) => {
      isCloudSynced.current = true;
      if (Array.isArray(remoteBills)) {
        const serialized = JSON.stringify(remoteBills);
        if (prevVendorBillsRef.current === serialized) return;
        prevVendorBillsRef.current = serialized;
        setVendorBills(remoteBills);
        localStorage.setItem('linoli_vendor_bills', serialized);
      }
    });
    // Inbound listener for salary advances
    const unsubAdvances = subscribeToCloud('salary_advances', (remoteAdv) => {
      isCloudSynced.current = true;
      if (Array.isArray(remoteAdv)) {
        const serialized = JSON.stringify(remoteAdv);
        if (prevAdvancesRef.current === serialized) return;
        prevAdvancesRef.current = serialized;
        setSalaryAdvances(remoteAdv);
        localStorage.setItem('linoli_salary_advances', serialized);
      }
    });
    
    return () => {
      if (typeof unsubOrders === 'function') unsubOrders();
      if (typeof unsubTrans === 'function') unsubTrans();
      if (typeof unsubTables === 'function') unsubTables();
      if (typeof unsubAudits === 'function') unsubAudits();
      if (typeof unsubMenu === 'function') unsubMenu();
      if (typeof unsubInventory === 'function') unsubInventory();
      if (typeof unsubExpenses === 'function') unsubExpenses();
      if (typeof unsubShift === 'function') unsubShift();
      if (typeof unsubStaff === 'function') unsubStaff();
      if (typeof unsubDenominations === 'function') unsubDenominations();
      if (typeof unsubAttendance === 'function') unsubAttendance();
      if (typeof unsubPayroll === 'function') unsubPayroll();
      if (typeof unsubAccounting === 'function') unsubAccounting();
      if (typeof unsubVendorBills === 'function') unsubVendorBills();
      if (typeof unsubAdvances === 'function') unsubAdvances();
    };
  }, []);

  // ============================================================
  // 2. BROADCAST LOCAL CHANGES UP TO FIREBASE (Guarded)
  // ============================================================
  useEffect(() => {
    if (!isCloudSynced.current) return;
    if (activeOrders !== undefined) {
      const current = JSON.stringify(activeOrders);
      if (current !== prevOrdersRef.current) {
        prevOrdersRef.current = current;
        syncToCloud('active_orders', activeOrders);
      }
    }
  }, [activeOrders]);

  useEffect(() => {
    if (!isCloudSynced.current) return;
    if (transactions !== undefined) {
      const current = JSON.stringify(transactions);
      if (current !== prevTransRef.current) {
        prevTransRef.current = current;
        syncToCloud('transactions', transactions);
      }
    }
  }, [transactions]);

  useEffect(() => {
    if (!isCloudSynced.current) return;
    if (floorTables !== undefined) {
      const current = JSON.stringify(floorTables);
      if (current !== prevTablesRef.current) {
        prevTablesRef.current = current;
        syncToCloud('floor_tables', floorTables);
      }
    }
  }, [floorTables]);

  useEffect(() => {
    if (!isCloudSynced.current) return;
    if (auditLogs !== undefined) {
      const current = JSON.stringify(auditLogs);
      if (current !== prevAuditsRef.current) {
        prevAuditsRef.current = current;
        syncToCloud('audit_logs', auditLogs);
      }
    }
  }, [auditLogs]);

  useEffect(() => {
    if (!isCloudSynced.current) return;
    if (menuItems !== undefined && menuItems.length > 0) {
      const current = JSON.stringify(menuItems);
      if (current !== prevMenuRef.current) {
        prevMenuRef.current = current;
        syncToCloud('menu_items', menuItems);
      }
    }
  }, [menuItems]);
  useEffect(() => {
    if (!isCloudSynced.current) return;
    if (inventory !== undefined && inventory.length > 0) {
      const current = JSON.stringify(inventory);
      if (current !== prevInventoryRef.current) {
        prevInventoryRef.current = current;
        syncToCloud('inventory', inventory);
      }
    }
  }, [inventory]);

  useEffect(() => {
    if (!isCloudSynced.current) return;
    if (expenses !== undefined) {
      const current = JSON.stringify(expenses);
      if (current !== prevExpensesRef.current) {
        prevExpensesRef.current = current;
        syncToCloud('expenses', expenses);
      }
    }
  }, [expenses]);
  
  useEffect(() => {
  if (!isCloudSynced.current) return;
  if (currentShift !== undefined) {
    const current = JSON.stringify(currentShift);
    if (current !== prevShiftRef.current) {
      prevShiftRef.current = current;
      syncToCloud('current_shift', currentShift);
    }
  }
  }, [currentShift]);

  useEffect(() => {
    if (!isCloudSynced.current) return;
    if (staffList !== undefined && staffList.length > 0) {
      const current = JSON.stringify(staffList);
      if (current !== prevStaffRef.current) {
        prevStaffRef.current = current;
        syncToCloud('staff_list', staffList);
      }
    }
  }, [staffList]);
  useEffect(() => {
    if (!isCloudSynced.current) return;
    if (denominations && typeof denominations === 'object') {
      const current = JSON.stringify(denominations);
      if (current !== prevDenomRef.current) {
        prevDenomRef.current = current;
        syncToCloud('denominations', denominations);
      }
    }
    }, [denominations]);

   // Live Bidirectional Sync: Attendance Logs (pos_state/attendance_logs)
  useEffect(() => {
    if (!isCloudSynced.current) return;
    if (attendanceLogs !== undefined) {
      const current = JSON.stringify(attendanceLogs);
      if (current !== prevAttendanceRef.current) {
        prevAttendanceRef.current = current;
        syncToCloud('attendance_logs', attendanceLogs);
      }
    }
  }, [attendanceLogs]);

  // Live Bidirectional Sync: Processed Payroll Records (pos_state/payroll_records)
  useEffect(() => {
    if (!isCloudSynced.current) return;
    if (payrollRecords !== undefined) {
      const current = JSON.stringify(payrollRecords);
      if (current !== prevPayrollRef.current) {
        prevPayrollRef.current = current;
        syncToCloud('payroll_records', payrollRecords);
      }
    }
  }, [payrollRecords]);

  // Broadcast accounting view filters/snapshots to cloud
  useEffect(() => {
    if (!isCloudSynced.current) return;
    if (accountingPeriod !== undefined) {
      syncToCloud('accounting_settings', { period: accountingPeriod });
    }
  }, [accountingPeriod]);

  useEffect(() => {
    if (!isCloudSynced.current) return;
    if (vendorBills !== undefined) {
      const current = JSON.stringify(vendorBills);
      if (current !== prevVendorBillsRef.current) {
        prevVendorBillsRef.current = current;
        syncToCloud('vendor_bills', vendorBills);
      }
    }
  }, [vendorBills]);

  // Outbound broadcaster for salary advances
  useEffect(() => {
    if (!isCloudSynced.current) return;
    if (salaryAdvances !== undefined) {
      const current = JSON.stringify(salaryAdvances);
      if (current !== prevAdvancesRef.current) {
        prevAdvancesRef.current = current;
        syncToCloud('salary_advances', salaryAdvances);
      }
    }
  }, [salaryAdvances]);

  const handleCreateStaff = (e) => {
    e.preventDefault();
    setStaffFormError('');
    if (!newStaffForm.name.trim()) {
      setStaffFormError('Staff name is required.');
      return;
    }
    const cleanPin = newStaffForm.pin.trim();
    if (!cleanPin || cleanPin.length !== 4 || !/^\d{4}$/.test(cleanPin)) {
      setStaffFormError('PIN must be exactly 4 numeric digits.');
      return;
    }
    if (staffList.some(s => s.pin === cleanPin && s.id !== editingStaffId)) {
      setStaffFormError('This 4-digit PIN is already in use by another employee.');
      return;
    }

    const nameParts = newStaffForm.name.trim().split(' ');
    const initials = nameParts.length > 1
      ? `${nameParts[0][0]}${nameParts[nameParts.length - 1][0]}`.toUpperCase()
      : newStaffForm.name.trim().substring(0, 2).toUpperCase();

    const staffPayload = {
      name: newStaffForm.name.trim(),
      role: newStaffForm.role,
      pin: cleanPin,
      avatar: initials || 'ST',
      email: newStaffForm.email.trim() || `${newStaffForm.name.trim().toLowerCase().replace(/\s+/g, '')}@linolicove.me`,
      basicSalary: Number(newStaffForm.basicSalary) || 0,
      budgetaryAllowance: Number(newStaffForm.budgetaryAllowance) || 0,
      otherAllowances: Number(newStaffForm.otherAllowances) || 0,
      fixedBonus: Number(newStaffForm.fixedBonus) || 0,
      overtimeRate: Number(newStaffForm.overtimeRate) || 0,
      epfEtfEnabled: newStaffForm.epfEtfEnabled !== false
    };

    if (editingStaffId) {
      setStaffList(prev => prev.map(s => s.id === editingStaffId ? { ...s, ...staffPayload, id: editingStaffId } : s));
      recordAuditLog('STAFF_UPDATED', editingStaffId, `Updated salary and credentials for ${staffPayload.name} (${staffPayload.role})`);
    } else {
      const newStaff = {
        id: `usr_${Date.now().toString().slice(-6)}`,
        ...staffPayload
      };
      setStaffList(prev => [...prev, newStaff]);
      recordAuditLog('STAFF_CREATED', newStaff.id, `Created staff member ${newStaff.name} with Base Salary ${settings.currency} ${staffPayload.basicSalary}`);
    }

    setAddStaffModalOpen(false);
    setEditingStaffId(null);
    setNewStaffForm({
      name: '',
      role: 'Cashier',
      pin: '',
      email: '',
      basicSalary: 35000,
      budgetaryAllowance: 2500,
      otherAllowances: 0,
      fixedBonus: 0,
      overtimeRate: 250,
      epfEtfEnabled: true
    });
    setStaffFormError('');
  };

  const handleCreateInventoryItem = (e) => {
    if (e && e.preventDefault) e.preventDefault();

    const cleanName = (newInventoryForm.name || '').trim();
    const parsedCost = parseFloat(newInventoryForm.cost);
    const parsedStock = parseFloat(newInventoryForm.stock) || 0;
    const parsedThreshold = parseFloat(newInventoryForm.threshold) || 10;

    if (!cleanName) {
      alert('Please enter a name for the raw material.');
      return;
    }

    if (isNaN(parsedCost) || parsedCost <= 0) {
      alert('Please enter a valid unit cost greater than 0.');
      return;
    }

    const newItem = {
      id: `ing_${Date.now().toString().slice(-6)}`,
      name: cleanName,
      category: newInventoryForm.category || 'Dry Goods',
      stock: parsedStock,
      unit: newInventoryForm.unit || 'g',
      cost: parsedCost,
      threshold: parsedThreshold
    };

    setInventory(prev => [...prev, newItem]);

    recordAuditLog(
      'INVENTORY_ITEM_CREATED',
      newItem.id,
      `Added raw material ${newItem.name} (${newItem.stock} ${newItem.unit} @ ${settings.currency} ${newItem.cost.toFixed(2)}/${newItem.unit})`
    );

    recordStockMovement(
      'INITIAL_ENTRY',
      newItem.id,
      newItem.name,
      newItem.stock,
      0,
      newItem.stock,
      newItem.unit,
      newItem.cost,
      'Initial stock registered',
      'NEW_MATERIAL'
    );

    setAddInventoryModalOpen(false);
    setNewInventoryForm({ name: '', category: 'Dry Goods', stock: '', unit: 'g', cost: '', threshold: '10' });
  };

  const handleReceiveStock = (e) => {
    if (e && e.preventDefault) e.preventDefault();

    if (!receiveStockForm.ingredientId || !receiveStockForm.quantity) {
      alert('Please select an ingredient and enter a quantity.');
      return;
    }

    const qtyToAdd = parseFloat(receiveStockForm.quantity);
    if (isNaN(qtyToAdd) || qtyToAdd <= 0) {
      alert('Please enter a valid positive quantity.');
      return;
    }

    const targetItem = inventoryMap[receiveStockForm.ingredientId];
    const oldStock = targetItem ? targetItem.stock : 0;
    const newStock = Number((oldStock + qtyToAdd).toFixed(2));
    const parsedNewCost = parseFloat(receiveStockForm.newCost);
    const hasValidNewCost = !isNaN(parsedNewCost) && parsedNewCost > 0;

    setInventory(prev => prev.map(item => {
      if (item.id === receiveStockForm.ingredientId) {
        return {
          ...item,
          stock: newStock,
          cost: hasValidNewCost ? parsedNewCost : item.cost
        };
      }
      return item;
    }));

    recordAuditLog(
      'STOCK_RECEIVED',
      receiveStockForm.ingredientId,
      `Received ${qtyToAdd} ${targetItem?.unit || 'units'} of ${targetItem?.name || receiveStockForm.ingredientId}. Supplier: ${receiveStockForm.supplier || 'N/A'}. Invoice: ${receiveStockForm.invoiceRef || 'N/A'}`
    );

    recordStockMovement(
      'INTAKE',
      receiveStockForm.ingredientId,
      targetItem?.name || receiveStockForm.ingredientId,
      qtyToAdd,
      oldStock,
      newStock,
      targetItem?.unit || 'units',
      hasValidNewCost ? parsedNewCost : (targetItem?.cost || 0),
      `Intake from ${receiveStockForm.supplier || 'Vendor'} (Inv: ${receiveStockForm.invoiceRef || 'N/A'})`,
      receiveStockForm.invoiceRef || 'GRN'
    );

    setReceiveStockModalOpen(false);
    setReceiveStockForm({ ingredientId: '', quantity: '', supplier: '', invoiceRef: '', newCost: '' });
  };

  const handleCreateTable = (e) => {
    e.preventDefault();
    if (!newTableForm.name.trim()) return;

    const count = floorTables.length + 1;
    const tableId = `T-${String(count).padStart(2, '0')}`;
    const newTable = {
      id: tableId,
      name: newTableForm.name.trim(),
      zone: newTableForm.zone || 'Indoor Main Hall',
      capacity: parseInt(newTableForm.capacity) || 4,
      status: 'VACANT',
      currentOrderRef: null
    };

    setFloorTables(prev => [...prev, newTable]);
    recordAuditLog(
      'TABLE_CREATED',
      newTable.id,
      `Created table "${newTable.name}" in ${newTable.zone} with capacity of ${newTable.capacity} seats`
    );

    setAddTableModalOpen(false);
    setNewTableForm({ name: '', zone: 'Indoor Main Hall', capacity: 4 });
  };

  // Backup & Restore Handlers
  const handleExportBackup = () => {
    const backupData = {
      app: 'Linoli Cove POS & ERP',
      version: '2.5.0',
      exportedAt: new Date().toISOString(),
      exportedBy: currentUser.name,
      settings,
      staffList,
      inventory,
      menuItems,
      floorTables,
      activeOrders,
      transactions,
      vendorBills, // <-- ADD THIS
      auditLogs,
      currentShift,
      shiftHistory
    };
    const jsonStr = JSON.stringify(backupData, null, 2);
    const blob = new Blob([jsonStr], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `linoli_cove_backup_${getLocalDateStr()}.json`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    recordAuditLog('BACKUP_EXPORTED', 'DATABASE', `Full system backup file downloaded by ${currentUser.name}`);
    setSettingsNotice({
      title: 'Backup Downloaded Successfully',
      detail: 'All menus, staff, inventory, invoices & audit records exported to JSON.'
    });
    setTimeout(() => setSettingsNotice(null), 4000);
  };

  const handleImportBackup = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (event) => {
      try {
        const parsed = JSON.parse(event.target.result);
        if (!parsed || (!parsed.settings && !parsed.menuItems)) {
          setSettingsNotice({
            title: 'Invalid Backup File',
            detail: 'The selected file is not a valid Linoli Cove POS database backup.'
          });
          setTimeout(() => setSettingsNotice(null), 4500);
          return;
        }
        if (parsed.settings) setSettings(parsed.settings);
        if (parsed.staffList) setStaffList(parsed.staffList);
        if (parsed.inventory) setInventory(parsed.inventory);
        if (parsed.menuItems) setMenuItems(parsed.menuItems);
        if (parsed.floorTables) setFloorTables(parsed.floorTables);
        if (parsed.activeOrders) setActiveOrders(parsed.activeOrders);
        if (parsed.transactions) setTransactions(parsed.transactions);
        if (parsed.auditLogs) setAuditLogs(parsed.auditLogs);
        if (parsed.cancelledTickets) setCancelledTickets(parsed.cancelledTickets);
        if (parsed.currentShift) setCurrentShift(parsed.currentShift);
        if (parsed.shiftHistory) setShiftHistory(parsed.shiftHistory);
        if (parsed.vendorBills) setVendorBills(parsed.vendorBills);

        recordAuditLog('BACKUP_RESTORED', file.name, `System restored from backup file by ${currentUser.name}`);
        setSettingsNotice({
          title: 'Database Restored Successfully',
          detail: `System data loaded from ${file.name}. All records synchronized.`
        });
        setTimeout(() => setSettingsNotice(null), 4500);
      } catch (err) {
        setSettingsNotice({
          title: 'Import Failed',
          detail: `Could not parse backup JSON file: ${err.message}`
        });
        setTimeout(() => setSettingsNotice(null), 4500);
      }
    };
    reader.readAsText(file);
    e.target.value = '';
  };

  useEffect(() => {
    if (navigator.usb) {
      navigator.usb.getDevices().then(devices => {
        if (devices.length > 0) {
          setPairedUsbDevice(devices[0]);
          setUsbStatusMessage(`Auto-connected to ${devices[0].productName || 'USB Thermal Printer'}`);
        }
      }).catch(err => console.warn('WebUSB auto-detect notice:', err));
    }
  }, []);

  // Direct push-to-print trigger (Dispatches immediately without opening any modal dialog)
  const triggerAutoPrint = (slipConfig, noticeText = 'Printing thermal receipt...') => {
    setActivePrintSlip(slipConfig);
    setPrintNotice({
      title: slipConfig.type === 'KOT_BOT_DISPATCH'
        ? 'Order Sent • Auto-Printing KOT & BOT'
        : slipConfig.type === 'FINAL_BILL'
        ? 'Bill Settled • Auto-Printing Tax Invoice'
        : slipConfig.type === 'TEMP_BILL'
        ? 'Auto-Printing Proforma Temp Bill'
        : slipConfig.type === 'CASH_OUT_VOUCHER'
        ? 'Auto-Printing Cash Out Voucher'
        : 'Auto-Printing...',
      detail: noticeText
    });
    setTimeout(() => setPrintNotice(null), 2500);
  };

  // Instant print invocation as soon as slip data is staged
  useEffect(() => {
    if (activePrintSlip) {
      document.body.classList.add('printing-thermal');
      const timer = setTimeout(() => {
        try {
          window.print();
        } catch (e) {
          console.warn('Auto print spooler notice:', e);
        } finally {
          // Reset after print dialog closes
          setTimeout(() => {
            document.body.classList.remove('printing-thermal');
            setActivePrintSlip(null);
          }, 500);
        }
      }, 80);
      return () => {
        clearTimeout(timer);
        document.body.classList.remove('printing-thermal');
      };
    } else {
      document.body.classList.remove('printing-thermal');
    }
  }, [activePrintSlip]);
  const recordAuditLog = (action, targetRef, details) => {
    const newLog = {
      id: `LOG-${Date.now()}-${Math.floor(100 + Math.random() * 900)}`,
      timestamp: new Date().toLocaleString(),
      action,
      targetRef,
      staff: currentUser.name,
      role: currentUser.role,
      details
    };
    setAuditLogs(prev => [newLog, ...prev]);
  };

  // Record Stock Movement Difference Ledger Entry
  const recordStockMovement = (type, ingredientId, ingredientName, diffQty, oldStock, newStock, unit, unitCost, reason, reference = '') => {
    const numericDiff = Number(diffQty) || 0;
    const numericCost = Number(unitCost) || 0;

    const newStockEntry = {
      id: `STK-${Date.now()}-${Math.floor(100 + Math.random() * 900)}`,
      timestamp: new Date().toLocaleString(),
      date: getLocalDateStr(),
      type, // 'INTAKE' | 'SALE_DEPLETION' | 'MANUAL_EDIT' | 'INITIAL_ENTRY'
      ingredientId,
      ingredientName,
      diffQty: numericDiff,
      oldStock: Number(oldStock) || 0,
      newStock: Number(newStock) || 0,
      unit: unit || 'units',
      unitCost: numericCost,
      totalCostImpact: Number((Math.abs(numericDiff) * numericCost).toFixed(2)),
      reason: reason || 'Stock level updated',
      reference: reference || 'N/A',
      staff: currentUser?.name || 'System'
    };

    setStockLogs(prev => [newStockEntry, ...(Array.isArray(prev) ? prev : [])]);
  };

  const inventoryMap = useMemo(() => {
    const map = {};
    (inventory || []).forEach(item => { if (item) map[item.id] = item; });
    return map;
  }, [inventory]);

  // Total Inventory Valuation & Cost Tracking Metrics
  const inventoryValuation = useMemo(() => {
    const list = Array.isArray(inventory) ? inventory : [];
    let totalStockValue = 0;
    let lowStockCount = 0;
    const categoryTotals = {};

    list.forEach(item => {
      if (!item) return;
      const stock = Number(item.stock) || 0;
      const cost = Number(item.cost) || 0;
      const threshold = Number(item.threshold) || 0;
      const itemVal = stock * cost;

      totalStockValue += itemVal;
      if (stock <= threshold) lowStockCount += 1;

      const cat = item.category || 'General';
      categoryTotals[cat] = (categoryTotals[cat] || 0) + itemVal;
    });

    // Detect highest value category investment
    let topCat = 'None';
    let maxCatVal = 0;
    Object.entries(categoryTotals).forEach(([cat, val]) => {
      if (val > maxCatVal) {
        maxCatVal = val;
        topCat = cat;
      }
    });

    return {
      totalStockValue,
      totalItems: list.length,
      lowStockCount,
      topCat,
      maxCatVal
    };
  }, [inventory]);

  const calculateDishAvailability = (recipe) => {
    if (!recipe || !Array.isArray(recipe) || recipe.length === 0) {
      return { cogs: 0, portions: 999, isSoldOut: false };
    }
    let cogs = 0;
    let minPortions = Infinity;

    recipe.forEach(r => {
      const ing = inventoryMap[r.ingredientId];
      if (ing) {
        cogs += (ing.cost * r.amount);
        const available = r.amount > 0 ? Math.floor(ing.stock / r.amount) : 0;
        if (available < minPortions) minPortions = available;
      } else {
        minPortions = 0;
      }
    });

    if (minPortions === Infinity) minPortions = 0;
    return { cogs, portions: minPortions, isSoldOut: minPortions <= 0 };
  };

  const generateEodReportData = (targetDate = getLocalDateStr()) => {
    const dayTransactions = transactions.filter(t => extractDateStr(t.date) === targetDate);
    const dayVoids = cancelledTickets.filter(v => extractDateStr(v.timestamp) === targetDate);
    const dayLogs = auditLogs.filter(l => extractDateStr(l.timestamp) === targetDate);
    const dayCashOuts = (currentShift.payouts || []).filter(p => (p.date === targetDate || !p.date));

    let gross = 0;
    let net = 0;
    let service = 0;
    let tax = 0;
    let discount = 0;
    const paymentBreakdown = {};

    dayTransactions.forEach(t => {
      gross += t.total;
      net += t.subtotal;
      service += t.serviceCharge;
      tax += t.tax;
      discount += (t.discount || 0);
      paymentBreakdown[t.paymentMethod] = (paymentBreakdown[t.paymentMethod] || 0) + t.total;
    });

    return {
      restaurant: settings.restaurantName,
      terminal: settings.terminalId,
      date: targetDate,
      generatedAt: new Date().toLocaleString(),
      summary: {
        totalSettledBills: dayTransactions.length,
        grossRevenue: gross,
        netSubtotal: net,
        serviceCharge: service,
        tax: tax,
        discounts: discount,
        paymentMethods: paymentBreakdown
      },
      cashierShift: {
        shiftId: currentShift.shiftId,
        openedBy: currentShift.openedBy,
        openedAt: currentShift.openedAt,
        startingFloat: currentShift.startingFloat,
        cashPayouts: dayCashOuts
      },
      invoices: dayTransactions,
      voidedTickets: dayVoids,
      activityAuditLogs: dayLogs
    };
  };

  const sendDailyEodEmail = async (isManual = false) => {
    const todayStr = getLocalDateStr();
    setIsSendingEmail(true);

    const report = generateEodReportData(todayStr);
    const recipient = emailSettings.recipient || 'linolicove@gmail.com';

    try {
      let sentSuccessfully = false;

      // Transport 1: Custom Webhook Endpoint (e.g. Zapier, Make, n8n, custom server)
      if (emailSettings.webhookUrl) {
        const res = await fetch(emailSettings.webhookUrl, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            recipient,
            subject: `[EOD REPORT] ${settings.restaurantName} - ${todayStr}`,
            report
          })
        });
        if (res.ok) sentSuccessfully = true;
      }

      // Transport 2: EmailJS API
      if (!sentSuccessfully && emailSettings.emailjsServiceId && emailSettings.emailjsTemplateId && emailSettings.emailjsPublicKey) {
        const emailjsRes = await fetch('https://api.emailjs.com/api/v1.0/email/send', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            service_id: emailSettings.emailjsServiceId,
            template_id: emailSettings.emailjsTemplateId,
            user_id: emailSettings.emailjsPublicKey,
            template_params: {
              to_email: recipient,
              subject: `[Daily Summary] ${settings.restaurantName} - ${todayStr}`,
              summary: `Gross: ${settings.currency} ${report.summary.grossRevenue.toFixed(2)} | Net: ${settings.currency} ${report.summary.netSubtotal.toFixed(2)} | Bills: ${report.summary.totalSettledBills}`,
              report_json: JSON.stringify(report, null, 2)
            }
          })
        });
        if (emailjsRes.ok) sentSuccessfully = true;
      }

      // Transport 3: Client-side mailto & JSON download fallback
      if (!sentSuccessfully) {
        const subject = encodeURIComponent(`[Daily Activity Report] ${settings.restaurantName} - ${todayStr}`);
        const bodyText = encodeURIComponent(
          `Linoli Cove Daily Business & Activity Report\n` +
          `Date: ${todayStr}\n` +
          `Terminal: ${settings.terminalId}\n\n` +
          `--- FINANCIAL SUMMARY ---\n` +
          `Gross Revenue: ${settings.currency} ${report.summary.grossRevenue.toFixed(2)}\n` +
          `Net Sales: ${settings.currency} ${report.summary.netSubtotal.toFixed(2)}\n` +
          `Service Charge: ${settings.currency} ${report.summary.serviceCharge.toFixed(2)}\n` +
          `Taxes: ${settings.currency} ${report.summary.tax.toFixed(2)}\n` +
          `Discounts: ${settings.currency} ${report.summary.discounts.toFixed(2)}\n` +
          `Total Settled Invoices: ${report.summary.totalSettledBills}\n\n` +
          `--- CASH DRAWER ---\n` +
          `Opening Float: ${settings.currency} ${report.cashierShift.startingFloat.toFixed(2)}\n` +
          `Cash Out Disbursements: ${report.cashierShift.cashPayouts.length} records\n\n` +
          `--- AUDIT & VOIDS ---\n` +
          `Voided Tickets: ${report.voidedTickets.length}\n` +
          `Activity Audit Logs: ${report.activityAuditLogs.length} events\n\n` +
          `(Full complete JSON data bundle has also been exported for archive)`
        );

        if (isManual) {
          window.open(`mailto:${recipient}?subject=${subject}&body=${bodyText}`, '_blank');
        }

        // Also download complete JSON archive
        const blob = new Blob([JSON.stringify(report, null, 2)], { type: 'application/json' });
        const downloadUrl = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = downloadUrl;
        a.download = `EOD_Complete_Report_${todayStr}.json`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(downloadUrl);

        sentSuccessfully = true;
      }

      // Update state and record audit log
      setEmailSettings(prev => ({ ...prev, lastSentDate: todayStr }));
      recordAuditLog('DAILY_EOD_EMAIL_DISPATCHED', recipient, `Dispatched complete day activity & financial report to ${recipient} (Date: ${todayStr})`);

      setSettingsNotice({
        title: 'Daily Report Dispatched',
        detail: `Complete operational report sent to ${recipient} for ${todayStr}.`
      });
      setTimeout(() => setSettingsNotice(null), 5000);
    } catch (err) {
      setSettingsNotice({
        title: 'Email Dispatch Notice',
        detail: `Could not send automatically: ${err.message}`
      });
      setTimeout(() => setSettingsNotice(null), 5000);
    } finally {
      setIsSendingEmail(false);
    }
  };

  useEffect(() => {
    if (!emailSettings.enabled) return;

    const intervalId = setInterval(() => {
      const now = new Date();
      const hours = String(now.getHours()).padStart(2, '0');
      const minutes = String(now.getMinutes()).padStart(2, '0');
      const currentTimeStr = `${hours}:${minutes}`;
      const todayStr = getLocalDateStr(now);

      const targetTime = emailSettings.scheduledTime || '23:30';

      // Fire when time matches 11:30 PM (23:30) and has not yet been sent today
      if (currentTimeStr === targetTime && emailSettings.lastSentDate !== todayStr && !isSendingEmail) {
        sendDailyEodEmail(false);
      }
    }, 30000); // Check every 30 seconds

    return () => clearInterval(intervalId);
  }, [emailSettings, isSendingEmail]);

  // Cart financials
  const cartSubtotal = cart.reduce((acc, item) => acc + item.price * item.qty, 0);
  const cartDiscountAmount = (cartSubtotal * discountPercent) / 100;
  const taxableBasis = Math.max(0, cartSubtotal - cartDiscountAmount);
  const cartServiceCharge = serviceChargeActive ? (taxableBasis * settings.serviceChargeRate) / 100 : 0;
  const cartTax = taxActive ? (taxableBasis * settings.taxRate) / 100 : 0;
  const cartGrandTotal = taxableBasis + cartServiceCharge + cartTax;

  const calculateOrderFinancials = (order) => {
    if (!order || !order.items) return { subtotal: 0, discount: 0, service: 0, tax: 0, total: 0 };
    const subtotal = order.items.reduce((sum, item) => sum + item.price * item.qty, 0);
    const discount = (subtotal * (order.discountPercent || 0)) / 100;
    const basis = Math.max(0, subtotal - discount);
    const service = order.serviceChargeActive ? (basis * settings.serviceChargeRate) / 100 : 0;
    const tax = order.taxActive ? (basis * settings.taxRate) / 100 : 0;
    const total = basis + service + tax;
    return { subtotal, discount, service, tax, total };
  };

  const playCashRegisterChime = () => {
    try {
      const AudioCtx = window.AudioContext || window.webkitAudioContext;
      if (!AudioCtx) return;
      const ctx = new AudioCtx();
      const osc = ctx.createOscillator();
      const gain = ctx.createGain();
      osc.type = 'sine';
      osc.frequency.setValueAtTime(1320, ctx.currentTime);
      osc.frequency.exponentialRampToValueAtTime(880, ctx.currentTime + 0.3);
      gain.gain.setValueAtTime(0.5, ctx.currentTime);
      gain.gain.exponentialRampToValueAtTime(0.01, ctx.currentTime + 0.3);
      osc.connect(gain);
      gain.connect(ctx.destination);
      osc.start();
      osc.stop(ctx.currentTime + 0.35);
    } catch (e) {
      console.warn('Audio notice:', e);
    }
  };

  const extractDateStr = (dateVal) => {
    if (!dateVal) return '';
    if (typeof dateVal !== 'string') return '';
    if (dateVal.includes('T')) return dateVal.split('T')[0];
    const firstToken = dateVal.split(' ')[0].replace(/,/g, '');
    if (/^\d{4}-\d{2}-\d{2}$/.test(firstToken)) return firstToken;
    const parts = firstToken.split('/');
    if (parts.length === 3 && parts[2].length === 4) {
      return `${parts[2]}-${String(parts[0]).padStart(2, '0')}-${String(parts[1]).padStart(2, '0')}`;
    }
    return firstToken;
  };

  const filteredTransactions = useMemo(() => {
    if (!reportStartDate && !reportEndDate) return transactions;
    return transactions.filter(t => {
      const tDate = extractDateStr(t.date);
      if (!tDate) return true;
      if (reportStartDate && tDate < reportStartDate) return false;
      if (reportEndDate && tDate > reportEndDate) return false;
      return true;
    });
  }, [transactions, reportStartDate, reportEndDate]);

  const salesMetrics = useMemo(() => {
    let grossRevenue = 0;
    let itemSubtotal = 0;
    let serviceCharge = 0;
    let taxes = 0;
    let discounts = 0;
    let kitchenRevenue = 0;
    let kitchenItemsCount = 0;
    let barRevenue = 0;
    let barItemsCount = 0;
    const paymentMethods = {};
    const itemSalesMap = {};
    const ingredientUsageMap = {};

    filteredTransactions.forEach(t => {
      grossRevenue += t.total;
      itemSubtotal += t.subtotal;
      serviceCharge += t.serviceCharge;
      taxes += t.tax;
      discounts += (t.discount || 0);

      paymentMethods[t.paymentMethod] = (paymentMethods[t.paymentMethod] || { count: 0, total: 0 });
      paymentMethods[t.paymentMethod].count += 1;
      paymentMethods[t.paymentMethod].total += t.total;

      (t.items || []).forEach(item => {
        const isKitchen = item.department === 'Kitchen';
        if (isKitchen) {
          kitchenRevenue += (item.price * item.qty);
          kitchenItemsCount += item.qty;
        } else {
          barRevenue += (item.price * item.qty);
          barItemsCount += item.qty;
        }

        if (!itemSalesMap[item.name]) {
          itemSalesMap[item.name] = {
            id: item.id,
            name: item.name,
            department: item.department || (isKitchen ? 'Kitchen' : 'Bar'),
            category: item.category || 'General',
            unitPrice: item.price || 0,
            sold: 0,
            revenue: 0,
            cogs: item.cogs || 0
          };
        }
        itemSalesMap[item.name].sold += item.qty;
        itemSalesMap[item.name].revenue += (item.price * item.qty);

        // Track raw ingredients consumed
        const menuItemRef = menuItems.find(m => m.name === item.name || m.id === item.id);
        if (menuItemRef?.recipe && Array.isArray(menuItemRef.recipe)) {
          menuItemRef.recipe.forEach(r => {
            const consumed = r.amount * item.qty;
            if (!ingredientUsageMap[r.ingredientId]) {
              ingredientUsageMap[r.ingredientId] = {
                ingredientId: r.ingredientId,
                totalConsumed: 0
              };
            }
            ingredientUsageMap[r.ingredientId].totalConsumed += consumed;
          });
        }
      });
    });

    const topItems = Object.values(itemSalesMap).sort((a, b) => b.sold - a.sold);

    return {
      grossRevenue,
      itemSubtotal,
      serviceCharge,
      taxes,
      discounts,
      kitchenRevenue,
      kitchenItemsCount,
      barRevenue,
      barItemsCount,
      paymentMethods,
      topItems,
      ingredientUsageMap,
      paidBillsCount: filteredTransactions.length
    };
  }, [filteredTransactions, menuItems]);

  const categoriesList = useMemo(() => {
    const cats = new Set(['All']);
    menuItems.forEach(m => cats.add(m.category));
    return Array.from(cats);
  }, [menuItems]);

  const hasAccess = (tabKey) => {
    const allowed = ROLE_PERMISSIONS[currentUser.role] || [];
    return allowed.includes(tabKey);
  };

  // Centralized Cart Action Mutators with Audit Logging
  const handleAddToCart = (dish) => {
    setCart(prev => {
      const existing = prev.find(i => i.id === dish.id);
      if (existing) {
        const nextQty = existing.qty + 1;
        recordAuditLog(
          'CART_ITEM_INCREMENTED',
          dish.id,
          `Increased "${dish.name}" quantity to ${nextQty} on active ticket`
        );
        return prev.map(i => i.id === dish.id ? { ...i, qty: nextQty } : i);
      }
      const cartItemId = `cart_${Date.now()}_${Math.random().toString(36).substring(2, 5)}`;
      recordAuditLog(
        'CART_ITEM_ADDED',
        dish.id,
        `Added "${dish.name}" (${settings.currency} ${dish.price.toFixed(2)}) to active ticket`
      );
      return [...prev, { ...dish, cartItemId, qty: 1, notes: '' }];
    });
  };

  const handleIncrementCartItem = (item) => {
    setCart(prev =>
      prev.map(i => {
        if (i.cartItemId === item.cartItemId) {
          const nextQty = i.qty + 1;
          recordAuditLog(
            'CART_QTY_INCREMENT',
            item.id,
            `Incremented "${item.name}" qty: ${i.qty} -> ${nextQty} (Line Total: ${settings.currency} ${(item.price * nextQty).toFixed(2)})`
          );
          return { ...i, qty: nextQty };
        }
        return i;
      })
    );
  };

  const handleDecrementCartItem = (item) => {
    if (item.qty <= 1) {
      handleRemoveCartItem(item, 'DECREMENT_ZERO');
      return;
    }
    setCart(prev =>
      prev.map(i => {
        if (i.cartItemId === item.cartItemId) {
          const nextQty = i.qty - 1;
          recordAuditLog(
            'CART_QTY_DECREMENT',
            item.id,
            `Decremented "${item.name}" qty: ${i.qty} -> ${nextQty} (Line Total: ${settings.currency} ${(item.price * nextQty).toFixed(2)})`
          );
          return { ...i, qty: nextQty };
        }
        return i;
      })
    );
  };

  const handleRemoveCartItem = (item, reason = 'MANUAL_REMOVE') => {
    setCart(prev => prev.filter(i => i.cartItemId !== item.cartItemId));
    recordAuditLog(
      'CART_ITEM_REMOVED',
      item.id,
      `Removed "${item.name}" (Qty: ${item.qty}, Amount: ${settings.currency} ${(item.price * item.qty).toFixed(2)}) from ticket [${reason}]`
    );
  };

  const handleSendOrder = () => {
    if (cart.length === 0) return;

    const nowTime = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });

    // =========================================================================
    // CASE A: Updating an existing active bill redirected from Billing Queue
    // =========================================================================
    if (settlingOrder && activeOrders.some(o => o.orderId === settlingOrder.orderId)) {
      const originalOrder = activeOrders.find(o => o.orderId === settlingOrder.orderId);
      const originalItems = originalOrder?.items || [];
      const updatedItems = cart;
      const changes = [];
      const newlyAddedOrIncremented = [];

      // 1. Detect item additions and increased quantities
      updatedItems.forEach(item => {
        const prev = originalItems.find(i => (i.cartItemId || i.id) === (item.cartItemId || item.id));
        if (!prev) {
          changes.push(`ADDED "${item.name}" (Qty: ${item.qty})`);
          newlyAddedOrIncremented.push({ ...item, qty: item.qty });
        } else if (item.qty > prev.qty) {
          const diff = item.qty - prev.qty;
          changes.push(`INCREASED "${item.name}" (+${diff}, now ${item.qty})`);
          newlyAddedOrIncremented.push({ ...item, qty: diff });
        } else if (item.qty < prev.qty) {
          const diff = prev.qty - item.qty;
          changes.push(`DECREASED "${item.name}" (-${diff}, now ${item.qty})`);
        }
      });

      // 2. Detect removed items
      originalItems.forEach(item => {
        const stillExists = updatedItems.some(i => (i.cartItemId || i.id) === (item.cartItemId || item.id));
        if (!stillExists) {
          changes.push(`REMOVED "${item.name}" (was Qty: ${item.qty})`);
        }
      });

      const changeSummary = changes.length > 0 ? changes.join(' | ') : 'No line changes';

      // 3. Update the existing active order with the updated cart items & financial settings
      const updatedOrder = {
        ...originalOrder,
        items: [...cart],
        serviceChargeActive,
        taxActive,
        discountPercent,
        lastUpdatedAt: nowTime
      };

      setActiveOrders(prev => prev.map(o => o.orderId === originalOrder.orderId ? updatedOrder : o));

      // 4. Auto-print ONLY the newly added items / increments to Kitchen & Bar
      if (newlyAddedOrIncremented.length > 0 && settings.autoPrintOrder !== false) {
        const kitchenItems = newlyAddedOrIncremented.filter(i => i.department === 'Kitchen');
        const barItems = newlyAddedOrIncremented.filter(i => i.department === 'Bar');

        if (kitchenItems.length > 0 || barItems.length > 0) {
          triggerAutoPrint({
            type: 'KOT_BOT_DISPATCH',
            data: {
              order: {
                ...updatedOrder,
                sentAt: nowTime,
                isAddon: true
              },
              kitchenItems,
              barItems
            }
          }, `${updatedOrder.tableName} • Add-on KOT/BOT Auto-Printed`);
        }
      }

      // 5. Audit Log
      recordAuditLog(
        'ORDER_UPDATED_FROM_POS',
        originalOrder.orderId,
        `Updated ${originalOrder.tableName} from POS: ${changeSummary}`
      );

      // Clean up editing state & return to Billing Queue
      setSettlingOrder(null);
      setCart([]);
      setActiveTab('billing');
      return;
    }

    // =========================================================================
    // CASE B: Standard New Order Creation
    // =========================================================================
    const newOrderId = getNextOrderNumber();

    const orderPayload = {
      orderId: newOrderId,
      mode: orderMode,
      tableId: orderMode === 'DINING' ? selectedTable.id : null,
      tableName: orderMode === 'DINING' ? selectedTable.name : takeawayInfo.token,
      zone: orderMode === 'DINING' ? selectedTable.zone : 'Takeaway Express',
      guestCount: orderMode === 'DINING' ? guestCount : 1,
      customerName: orderMode === 'TAKEAWAY' ? takeawayInfo.name : undefined,
      server: currentUser.name,
      sentAt: nowTime,
      status: 'PREPARING',
      serviceChargeActive,
      taxActive,
      discountPercent,
      items: [...cart]
    };

    setActiveOrders(prev => [orderPayload, ...prev]);

    if (orderMode === 'DINING') {
      setFloorTables(prev => prev.map(t => t.id === selectedTable.id ? { ...t, status: 'OCCUPIED', currentOrderRef: newOrderId } : t));
    }

    // Comprehensive snapshot logging: items, quantities, modifiers, and financial total
    const itemSummary = cart.map(i => `${i.qty}x ${i.name}${i.notes ? ` (${i.notes})` : ''}`).join(', ');
    const totalPortions = cart.reduce((acc, i) => acc + i.qty, 0);

    recordAuditLog(
      'ORDER_DISPATCHED',
      newOrderId,
      `Dispatched order ${newOrderId} (${orderPayload.tableName}) with ${totalPortions} portions across ${cart.length} line items. Total: ${settings.currency} ${cartGrandTotal.toFixed(2)}. Items: [${itemSummary}]`
    );

    const kitchenItems = cart.filter(i => i.department === 'Kitchen');
    const barItems = cart.filter(i => i.department === 'Bar');

    // Immediately auto-push KOT & BOT to thermal printer without modal popup
    triggerAutoPrint({
      type: 'KOT_BOT_DISPATCH',
      data: {
        order: orderPayload,
        kitchenItems,
        barItems
      }
    }, `${orderPayload.tableName} • 2 Slips (KOT & BOT)`);

    setCart([]);
  };
  
  
  const handleCompleteSettlement = () => {
    const targetOrder = settlingOrder || {
      orderId: getNextOrderNumber(),
      mode: orderMode,
      tableName: orderMode === 'DINING' ? selectedTable.name : takeawayInfo.token,
      tableId: orderMode === 'DINING' ? selectedTable.id : null,
      items: cart,
      serviceChargeActive,
      taxActive,
      discountPercent
    };

    if (!targetOrder.items || targetOrder.items.length === 0) return;

    const { subtotal, discount, service, tax, total } = calculateOrderFinancials(targetOrder);

    // Inventory BOM deduction calculation
    const deductions = {};
    let orderRawCost = 0;

    targetOrder.items.forEach(cartItem => {
      const dish = menuItems.find(m => m.id === cartItem.id) || cartItem;
      if (dish.recipe && Array.isArray(dish.recipe)) {
        dish.recipe.forEach(r => {
          const needed = r.amount * cartItem.qty;
          deductions[r.ingredientId] = (deductions[r.ingredientId] || 0) + needed;
          const ing = inventoryMap[r.ingredientId];
          if (ing) orderRawCost += (ing.cost * needed);
        });
      }
    });

    // Apply inventory deductions
    setInventory(prev => prev.map(item => {
      if (deductions[item.id]) {
        return {
          ...item,
          stock: Math.max(0, Number((item.stock - deductions[item.id]).toFixed(2)))
        };
      }
      return item;
    }));

    const newInvoice = {
      invoiceNo: getNextInvoiceNumber(),
      orderRef: targetOrder.orderId,
      shiftId: currentShift.shiftId,
      date: `${getLocalDateStr()} ${new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}`,
      table: targetOrder.tableName,
      mode: targetOrder.mode,
      cashier: currentUser.name,
      items: targetOrder.items.map(i => ({
        id: i.id,
        name: i.name,
        department: i.department,
        category: i.category || 'General',
        qty: i.qty,
        price: i.price
      })),
      subtotal,
      serviceCharge: service,
      tax,
      discount,
      total,
      paymentMethod,
      cogs: orderRawCost,
      cashTendered: paymentMethod === 'CASH' ? (parseFloat(cashTendered) || total) : undefined,
      changeDue: paymentMethod === 'CASH' ? Math.max(0, (parseFloat(cashTendered) || total) - total) : 0
    };

    // Log each recipe ingredient deduction in the Stock Movement Ledger
    Object.entries(deductions).forEach(([ingId, needed]) => {
      const ing = inventoryMap[ingId];
      if (ing) {
        const curStock = Number(ing.stock) || 0;
        const afterStock = Math.max(0, Number((curStock - needed).toFixed(2)));
        recordStockMovement(
          'SALE_DEPLETION',
          ingId,
          ing.name,
          -needed,
          curStock,
          afterStock,
          ing.unit,
          ing.cost,
          `BOM deduction for settled Bill #${newInvoice.invoiceNo} (${newInvoice.table})`,
          newInvoice.invoiceNo
        );
      }
    });

    setTransactions(prev => [newInvoice, ...prev]);
    setActiveOrders(prev => prev.filter(o => o.orderId !== targetOrder.orderId));

    if (targetOrder.tableId) {
      setFloorTables(prev => prev.map(t => t.id === targetOrder.tableId ? { ...t, status: 'VACANT', currentOrderRef: null } : t));
    }

    recordAuditLog('BILL_SETTLED', newInvoice.invoiceNo, `Settled ${newInvoice.invoiceNo} (${newInvoice.table}) for ${settings.currency} ${newInvoice.total.toFixed(2)} via ${paymentMethod}.`);

    if (settings.autoDrawerKick === 'ENABLED' && (settings.drawerKickTrigger === 'ALL' || (settings.drawerKickTrigger === 'CASH_ONLY' && paymentMethod === 'CASH'))) {
      if (settings.chimeAudio) playCashRegisterChime();
    }

    // Immediately auto-push Final Tax Invoice to thermal printer without modal popup
    triggerAutoPrint({
      type: 'FINAL_BILL',
      data: newInvoice
    }, `${newInvoice.table} • ${settings.currency} ${newInvoice.total.toFixed(2)}`);

    setSettlingOrder(null);
    setCheckoutModalOpen(false);
    setCashTendered('');
    setCart([]);
  };
  const handlePinSubmit = (pinVal) => {
    const pin = pinVal || loginPinInput;
    setLoginError('');
    const found = staffList.find(s => s.pin === pin);
    if (found) {
      setCurrentUser(found);
      setIsAuthenticated(true);
      setLoginPinInput('');
      const allowed = ROLE_PERMISSIONS[found.role] || [];
      setActiveTab(allowed.includes('pos') ? 'pos' : (allowed[0] || 'pos'));
      recordAuditLog('STAFF_LOGIN', found.id, `User ${found.name} logged in with ${found.role} role.`);
    } else {
      setLoginError('Invalid security PIN. Default Admin: 1234, Cashier: 1111');
    }
  };

  const calculateSriLankanPayroll = ({
    epfEtfEnabled = true,
    basicSalary = 0,
    budgetaryAllowance = 0,
    otherAllowances = 0,
    serviceChargeBonus = 0,
    incentiveBonus = 0,
    overtimeHours = 0,
    overtimeRate = 0,
    salaryAdvance = 0,
    otherDeductions = 0,
    // Attendance & Leave Metrics
    standardWorkingDays = 26,
    workedDays = 26,
    paidLeaves = 0,
    unpaidLeaves = 0,
    holidaysCount = 0,
    shortShiftsCount = 0
  }) => {
    const basic = Number(basicSalary) || 0;
    const bra = Number(budgetaryAllowance) || 0;
    const allowances = Number(otherAllowances) || 0;
    const pool = Number(serviceChargeBonus) || 0;
    const bonus = Number(incentiveBonus) || 0;
    const ot = (Number(overtimeHours) || 0) * (Number(overtimeRate) || 0);
    const advance = Number(salaryAdvance) || 0;
    const otherDed = Number(otherDeductions) || 0;

    const stdDays = Math.max(1, Number(standardWorkingDays) || 26);
    const unpdLeaves = Number(unpaidLeaves) || 0;

    // Prorated daily rate for unpaid leaves deduction
    const perDayBasicRate = Number((basic / stdDays).toFixed(2));
    const unpaidLeaveDeduction = Number((perDayBasicRate * unpdLeaves).toFixed(2));

    // EPF Base Earnings (Basic + BRA + Allowances)
    const epfLiableEarnings = Number((basic + bra + allowances).toFixed(2));

    // EPF/ETF calculations
    const epfEmployee = epfEtfEnabled ? Number((epfLiableEarnings * 0.08).toFixed(2)) : 0;
    const epfEmployer = epfEtfEnabled ? Number((epfLiableEarnings * 0.12).toFixed(2)) : 0;
    const etfEmployer = epfEtfEnabled ? Number((epfLiableEarnings * 0.03).toFixed(2)) : 0;
    const totalEpfFund = Number((epfEmployee + epfEmployer).toFixed(2));

    // Gross, Deductions, and Net
    const grossEarnings = Number((basic + bra + allowances + pool + bonus + ot).toFixed(2));
    const totalDeductions = Number((epfEmployee + advance + unpaidLeaveDeduction + otherDed).toFixed(2));
    const netSalary = Number((grossEarnings - totalDeductions).toFixed(2));
    const costToCompany = Number((grossEarnings + epfEmployer + etfEmployer).toFixed(2));

    return {
      epfEtfEnabled,
      basic,
      bra,
      allowances,
      epfLiableEarnings,
      epfEmployee,
      epfEmployer,
      etfEmployer,
      totalEpfFund,
      serviceChargeBonus: pool,
      incentiveBonus: bonus,
      overtimePay: ot,
      salaryAdvance: advance,
      unpaidLeaveDeduction,
      perDayBasicRate,
      standardWorkingDays: stdDays,
      workedDays: Number(workedDays) || 0,
      paidLeaves: Number(paidLeaves) || 0,
      unpaidLeaves: unpdLeaves,
      holidaysCount: Number(holidaysCount) || 0,
      shortShiftsCount: Number(shortShiftsCount) || 0,
      grossEarnings,
      totalDeductions,
      netSalary,
      costToCompany
    };
  };

  // Quick helper to check if a staff member is currently clocked in
  const getStaffClockStatus = (staffId) => {
    const today = getLocalDateStr();
    const userTodayLogs = (attendanceLogs || []).filter(a => a.staffId === staffId && a.date === today);
    if (userTodayLogs.length === 0) return { clockedIn: false, log: null };
    const latest = userTodayLogs[userTodayLogs.length - 1];
    return { clockedIn: !latest.clockOutTime, log: latest };
  };

  if (!isAuthenticated) {
    return (
      <div className="flex min-h-screen w-full items-center justify-center bg-[#070b14] p-4 font-sans select-none antialiased">
        <div className="w-full max-w-[390px] rounded-[32px] border border-[#1b253b] bg-[#0c1424]/95 p-8 shadow-2xl shadow-black/80 backdrop-blur-md">
          <div className="flex flex-col items-center text-center">
            <div className="flex h-16 w-16 items-center justify-center rounded-2xl bg-gradient-to-tr from-[#ff4500] to-[#ff6a00] text-2xl font-black text-white shadow-lg shadow-orange-600/40">
              LC
            </div>
            <h1 className="mt-3.5 text-2xl font-black tracking-tight text-white">{settings.restaurantName}</h1>
            <p className="mt-0.5 text-[10px] font-extrabold tracking-[0.2em] text-[#ff5500] uppercase">
              RESTAURANT &amp; BAR ERP
            </p>
            <div className="mt-3 inline-flex items-center gap-1.5 rounded-full bg-emerald-500/10 px-3 py-1 text-[11px] font-semibold text-emerald-400 border border-emerald-500/20">
              <span className="h-2 w-2 rounded-full bg-emerald-400 animate-pulse" />
              Terminal Ready • Enter PIN
            </div>
          </div>

          <div className="mt-6 flex flex-col items-center">
            <div className="flex h-12 w-full items-center justify-center rounded-2xl border border-zinc-800 bg-[#070b14] px-4">
              <div className="flex items-center gap-3">
                {[0, 1, 2, 3].map(idx => (
                  <span
                    key={idx}
                    className={`h-3.5 w-3.5 rounded-full transition-all ${
                      loginPinInput.length > idx ? 'bg-orange-500 scale-110' : 'bg-zinc-700'
                    }`}
                  />
                ))}
              </div>
            </div>
          </div>

          {loginError && (
            <p className="mt-2 text-center text-xs font-bold text-rose-500">{loginError}</p>
          )}

          <div className="mt-4 grid grid-cols-3 gap-2">
            {[1, 2, 3, 4, 5, 6, 7, 8, 9].map(num => (
              <button
                key={num}
                type="button"
                onClick={() => {
                  if (loginPinInput.length < 4) {
                    const next = loginPinInput + num;
                    setLoginPinInput(next);
                    if (next.length === 4) handlePinSubmit(next);
                  }
                }}
                className="flex h-12 items-center justify-center rounded-2xl border border-zinc-800 bg-[#10192b] text-base font-bold text-white hover:bg-zinc-800 active:scale-95 transition-all"
              >
                {num}
              </button>
            ))}
            <button
              type="button"
              onClick={() => setLoginPinInput('')}
              className="flex h-12 items-center justify-center rounded-2xl border border-zinc-800 bg-[#10192b] text-xs font-bold text-zinc-400 hover:bg-zinc-800"
            >
              Clear
            </button>
            <button
              type="button"
              onClick={() => {
                if (loginPinInput.length < 4) {
                  const next = loginPinInput + '0';
                  setLoginPinInput(next);
                  if (next.length === 4) handlePinSubmit(next);
                }
              }}
              className="flex h-12 items-center justify-center rounded-2xl border border-zinc-800 bg-[#10192b] text-base font-bold text-white hover:bg-zinc-800 active:scale-95 transition-all"
            >
              0
            </button>
            <button
              type="button"
              onClick={() => setLoginPinInput(prev => prev.slice(0, -1))}
              className="flex h-12 items-center justify-center rounded-2xl border border-zinc-800 bg-[#10192b] text-xs font-bold text-zinc-400 hover:bg-zinc-800"
            >
              Del
            </button>
          </div>

          <div className="mt-5 border-t border-zinc-800/80 pt-3 text-center">
            <p className="text-[10px] text-zinc-500">
              Authorized Terminal • Enter Assigned 4-Digit Security PIN
            </p>
          </div>
        </div>
      </div>
    );
  }

  {/* Global CSS fix for select dropdowns & isolated print modes */}
  return (
    <div className="flex h-screen w-full bg-[#0b0f19] text-zinc-100 font-sans select-none overflow-hidden antialiased">
      {/* Global CSS fix for select dropdowns & dual thermal / standard PDF print styles */}
      <style>{`
        select, option {
          color: #0f172a !important;
          background-color: #ffffff !important;
        }
        select:focus, option:focus, option:checked {
          color: #ff5500 !important;
          background-color: #fff7ed !important;
        }

        @media print {
          /* 1. Global Print Cleanups */
          *, *::before, *::after {
            box-sizing: border-box !important;
            -webkit-print-color-adjust: exact !important;
            print-color-adjust: exact !important;
          }

          /* Hide application UI elements */
          aside,
          header,
          button,
          input,
          .fixed,
          nav,
          [title="Touch to Open Menu"] {
            display: none !important;
          }

          body {
            background-color: #ffffff !important;
            color: #000000 !important;
            margin: 0 !important;
            padding: 0 !important;
          }

          /* 2. THERMAL PRINT MODE (When activePrintSlip exists / #thermal-print-area has content) */
          body.printing-thermal * {
            visibility: hidden !important;
          }

          body.printing-thermal main {
            display: none !important;
          }

          body.printing-thermal #thermal-print-area,
          body.printing-thermal #thermal-print-area * {
            visibility: visible !important;
            display: block !important;
          }

          body.printing-thermal #thermal-print-area {
            position: absolute !important;
            left: 0 !important;
            top: 0 !important;
            width: ${settings.receiptRollWidth === '58mm' ? '58mm' : '80mm'} !important;
            max-width: ${settings.receiptRollWidth === '58mm' ? '58mm' : '80mm'} !important;
            margin: 0 !important;
            padding: ${settings.receiptMargin || '2mm'} !important;
            background: #ffffff !important;
            color: #000000 !important;
            font-size: ${settings.receiptFontSize || '11px'} !important;
            font-family: ${settings.receiptFontFamily || 'monospace'} !important;
            line-height: 1.25 !important;
          }

          body.printing-thermal #thermal-print-area .flex {
            display: flex !important;
          }

          body.printing-thermal #thermal-print-area .border-dashed,
          body.printing-thermal #thermal-print-area .border-t,
          body.printing-thermal #thermal-print-area .border-b {
            border-color: #000000 !important;
          }

          body.printing-thermal #thermal-print-area div {
            break-inside: avoid;
          }

          /* 3. REPORT / DOCUMENT MODE (When no thermal slip is active) */
          body:not(.printing-thermal) main {
            display: block !important;
            width: 100% !important;
            overflow: visible !important;
            padding: 0 !important;
            margin: 0 !important;
          }

          body:not(.printing-thermal) table {
            width: 100% !important;
            border-collapse: collapse !important;
            color: #000000 !important;
          }

          body:not(.printing-thermal) th,
          body:not(.printing-thermal) td {
            color: #000000 !important;
            border-bottom: 1px solid #cbd5e1 !important;
            padding: 6px 8px !important;
          }

          body:not(.printing-thermal) thead {
            display: table-header-group !important;
          }

          body:not(.printing-thermal) tr {
            break-inside: avoid !important;
          }
        }
      `}</style>

      {/* Floating edge tab to open menu on touch */}
      <button
        onClick={() => setSidebarOpen(true)}
        className="fixed left-0 top-1/2 -translate-y-1/2 z-40 bg-[#ff5500] hover:bg-orange-600 text-white px-2 py-4 rounded-r-2xl shadow-2xl flex flex-col items-center gap-1.5 transition-transform hover:scale-105"
        title="Touch to Open Menu"
      >
        <Menu className="h-4 w-4" />
        <span className="text-[9px] font-black uppercase tracking-widest [writing-mode:vertical-lr]">MENU</span>
      </button>

      {/* Backdrop overlay for slide menu */}
      {sidebarOpen && (
        <div
          onClick={() => setSidebarOpen(false)}
          className="fixed inset-0 bg-black/60 backdrop-blur-xs z-40 transition-opacity"
        />
      )}

      {/* SLIDE-OUT DRAWER SIDEBAR */}
      <aside
        className={`fixed inset-y-0 left-0 w-64 bg-[#060813] border-r border-zinc-800/80 flex flex-col justify-between shrink-0 z-50 transition-transform duration-300 ease-in-out shadow-2xl ${
          sidebarOpen ? 'translate-x-0' : '-translate-x-full'
        }`}
      >
        <div className="overflow-y-auto flex-1">
          {/* Brand header */}
          <div className="p-5 pb-4 flex items-center justify-between border-b border-zinc-900">
            <div className="flex items-center gap-3">
              <div className="h-10 w-10 rounded-xl bg-[#ff5500] text-white font-black text-xl flex items-center justify-center shadow-lg shadow-orange-600/30">
                LC
              </div>
              <div>
                <h1 className="text-sm font-extrabold tracking-tight text-white leading-none">
                  {settings.restaurantName}
                </h1>
                <p className="text-[9px] font-bold tracking-widest text-[#ff5500] uppercase mt-1">
                  {settings.tagline}
                </p>
              </div>
            </div>
            <button
              onClick={() => setSidebarOpen(false)}
              className="p-1 text-zinc-400 hover:text-white rounded-lg hover:bg-zinc-800"
            >
              <X className="h-5 w-5" />
            </button>
          </div>

          {/* Navigation Links */}
          <nav className="p-3 space-y-1">
            {[
              { id: 'pos', name: 'POS Terminal', icon: Monitor, badge: cart.reduce((a, b) => a + b.qty, 0) },
              { id: 'kds', name: 'Kitchen Display', icon: Flame },
              { id: 'bar', name: 'Bar Display', icon: Wine },
              { id: 'billing', name: 'Billing & Settlement', icon: Receipt, badge: activeOrders.length },
              { id: 'tables', name: 'Table Management', icon: Grid },
              { id: 'stock', name: 'Stock & Inventory', icon: Package, alert: inventory.some(i => i.stock <= i.threshold) },
              { id: 'recipes', name: 'Recipes & Portions', icon: BookOpen },
              { id: 'shifts', name: 'Cashier Shifts', icon: DollarSign },
              { id: 'reports', name: 'Sales Reports', icon: BarChart3 }
            ].map(item => {
              const allowed = hasAccess(item.id);
              return (
                <button
                  key={item.id}
                  disabled={!allowed}
                  onClick={() => {
                    setActiveTab(item.id);
                    setSidebarOpen(false);
                  }}
                  className={`w-full flex items-center justify-between px-3.5 py-2.5 rounded-xl text-xs font-semibold tracking-wide transition-all ${
                    !allowed
                      ? 'opacity-30 cursor-not-allowed text-zinc-600'
                      : activeTab === item.id
                      ? 'bg-[#ff5500] text-white font-bold shadow-lg shadow-orange-600/20'
                      : 'text-zinc-400 hover:bg-zinc-900 hover:text-white'
                  }`}
                >
                  <div className="flex items-center gap-3">
                    <item.icon className="h-4 w-4" />
                    <span>{item.name}</span>
                  </div>
                  {item.badge > 0 && (
                    <span className={`px-2 py-0.5 text-[10px] rounded-full font-bold ${
                      activeTab === item.id ? 'bg-white text-zinc-900' : 'bg-orange-500/20 text-[#ff5500]'
                    }`}>
                      {item.badge}
                    </span>
                  )}
                  {item.alert && (
                    <span className="h-2 w-2 rounded-full bg-rose-500 animate-pulse" />
                  )}
                </button>
              );
            })}

            <div className="pt-4 pb-1.5 px-3">
              <span className="text-[10px] font-bold text-zinc-500 uppercase tracking-wider">
                ADMINISTRATION
              </span>
            </div>

            {[
              { id: 'menu_admin', name: 'Menu Management', icon: ClipboardList, badgeText: '+Add' },
              { id: 'vendor_bills', name: 'Vendor Bills & Invoices', icon: FileText }, // <-- ADD THIS
              { id: 'accounting', name: 'Accounting & P&L', icon: PieChart }, // <-- ADD THIS
              { id: 'staff', name: 'Staff Management', icon: Users },
              { id: 'payroll', name: 'Employment & Payroll', icon: Briefcase },
              { id: 'settings', name: 'System Settings', icon: Settings }
            ].map(item => {
              const allowed = hasAccess(item.id);
              return (
                <button
                  key={item.id}
                  disabled={!allowed}
                  onClick={() => {
                    setActiveTab(item.id);
                    setSidebarOpen(false);
                  }}
                  className={`w-full flex items-center justify-between px-3.5 py-2.5 rounded-xl text-xs font-semibold tracking-wide transition-all ${
                    !allowed
                      ? 'opacity-30 cursor-not-allowed text-zinc-600'
                      : activeTab === item.id
                      ? 'bg-[#ff5500] text-white font-bold shadow-lg shadow-orange-600/20'
                      : 'text-zinc-400 hover:bg-zinc-900 hover:text-white'
                  }`}
                >
                  <div className="flex items-center gap-3">
                    <item.icon className="h-4 w-4" />
                    <span>{item.name}</span>
                  </div>
                  {item.badgeText && allowed && (
                    <span className="text-[10px] bg-orange-500/20 text-[#ff5500] px-1.5 py-0.2 rounded font-bold">
                      {item.badgeText}
                    </span>
                  )}
                </button>
              );
            })}
          </nav>
        </div>

        {/* User profile footer */}
        <div className="p-3 border-t border-zinc-900 flex items-center justify-between shrink-0">
          <div className="flex items-center gap-2.5">
            <div className="h-9 w-9 rounded-full bg-zinc-800 text-orange-400 font-black text-xs flex items-center justify-center border border-zinc-700">
              {currentUser.avatar}
            </div>
            <div>
              <p className="text-xs font-bold text-white leading-tight">{currentUser.name}</p>
              <div className="flex items-center gap-1.5">
                <span className="text-[10px] text-zinc-400">{currentUser.role}</span>
                <span className="h-1.5 w-1.5 rounded-full bg-emerald-500" />
              </div>
            </div>
          </div>
          <button
            onClick={() => setIsAuthenticated(false)}
            title="Lock Terminal"
            className="p-1.5 rounded-xl text-zinc-400 hover:text-rose-400 hover:bg-zinc-800"
          >
            <LogOut className="h-4 w-4" />
          </button>
        </div>
      </aside>

      {/* MAIN CONTAINER */}
      <main className="flex-1 flex flex-col overflow-hidden bg-slate-50 text-slate-900">
        
       {/* Top Header Bar */}
        <header className="h-14 px-6 bg-white border-b border-slate-200 flex items-center justify-between shrink-0 shadow-xs z-10">
          <div className="flex items-center gap-4">
            {/* ENLARGED HIGH-VISIBILITY MENU BUTTON */}
            <button
              type="button"
              onClick={() => setSidebarOpen(true)}
              className="px-4 py-2 rounded-xl bg-slate-900 hover:bg-black text-white font-black text-xs uppercase tracking-wider flex items-center gap-2.5 transition-all shadow-md border border-slate-700 hover:border-slate-500 cursor-pointer active:scale-95"
            >
              <Menu className="h-5 w-5 text-[#ff5500] stroke-[2.5]" />
              <span className="font-extrabold text-xs tracking-wider text-white">Menu</span>
            </button>

            <span className="text-xs font-bold text-slate-500 uppercase tracking-wider font-mono">
              TERMINAL: {settings.terminalId}
            </span>
            <div className="flex items-center gap-1.5 bg-emerald-50 text-emerald-700 px-2.5 py-0.5 rounded-full text-xs font-semibold">
              <span className="h-2 w-2 rounded-full bg-emerald-500 animate-ping" />
              Online
            </div>
          </div>

          <div className="flex items-center gap-3">
            {hasAccess('menu_admin') && (
              <button
                onClick={() => setAddItemModalOpen(true)}
                className="px-3.5 py-1.5 bg-[#ff5500] hover:bg-orange-600 text-white rounded-xl text-xs font-bold flex items-center gap-1.5 shadow-sm shadow-orange-600/20 transition-all"
              >
                <Plus className="h-3.5 w-3.5" />
                <span>Add Item</span>
              </button>
            )}

            <div className="px-3 py-1 bg-slate-100 border border-slate-200 rounded-xl text-xs font-mono font-bold text-slate-700">
              Currency: {settings.currency}
            </div>

            <button
              onClick={() => setIsAuthenticated(false)}
              className="px-3 py-1.5 bg-slate-900 hover:bg-slate-800 text-white rounded-xl text-xs font-bold flex items-center gap-1.5"
            >
              <Lock className="h-3.5 w-3.5" />
              <span>Lock</span>
            </button>
          </div>
        </header>

{/* VIEW 1: RESPONSIVE POS TERMINAL (MOBILE, TABLET & DESKTOP) */}
        {activeTab === 'pos' && (
          <div className="flex-1 flex flex-col lg:flex-row overflow-hidden relative min-h-0">
            {/* Catalog Grid Area */}
            <div className="flex-1 flex flex-col p-3 sm:p-5 overflow-hidden min-h-0">
              {/* TOP DOCK: CATEGORIES, DENSITY, SEARCH */}
              <div className="bg-slate-900 border border-slate-800 rounded-2xl p-2 sm:p-2.5 mb-3 sm:mb-4 shrink-0 shadow-md flex flex-col sm:flex-row items-stretch sm:items-center gap-2.5">
                {/* Horizontal Scrollable Category Bar */}
                <div 
                  className="flex-1 min-w-0 overflow-x-auto pb-1 pt-0.5 scrollbar-thin"
                  style={{
                    scrollbarWidth: 'thin',
                    scrollbarColor: '#475569 #1e293b'
                  }}
                >
                  <div className="flex items-center gap-1.5 sm:gap-2">
                    {categoriesList.map((cat) => {
                      const count = cat === 'All' 
                        ? menuItems.length 
                        : menuItems.filter(m => m.category === cat).length;
                      const isSelected = selectedCategory === cat;

                      return (
                        <button
                          key={cat}
                          type="button"
                          onClick={() => setSelectedCategory(cat)}
                          className={`px-3 py-1.5 sm:px-4 sm:py-2 rounded-xl font-black text-[11px] sm:text-xs uppercase tracking-wider shrink-0 transition-all flex items-center gap-1.5 sm:gap-2 cursor-pointer shadow-xs ${
                            isSelected
                              ? 'bg-[#ff5500] text-white shadow-orange-500/40 ring-2 ring-[#ff5500]/50'
                              : 'bg-slate-800 hover:bg-slate-700 text-slate-100 border border-slate-700 hover:border-slate-600'
                          }`}
                        >
                          <span>{cat}</span>
                          <span
                            className={`px-1.5 py-0.2 rounded-full text-[9px] sm:text-[10px] font-black ${
                              isSelected
                                ? 'bg-black/30 text-white'
                                : 'bg-slate-900/80 text-slate-300 border border-slate-700/60'
                            }`}
                          >
                            {count}
                          </span>
                        </button>
                      );
                    })}
                  </div>
                </div>

                {/* Right Controls: View Switcher & Search Bar */}
                <div className="flex items-center gap-2 shrink-0">
                  {/* Density switcher (hidden on very small phones) */}
                  <div className="hidden sm:flex items-center bg-slate-800 border border-slate-700 rounded-xl p-0.5 shadow-inner">
                    <button
                      type="button"
                      onClick={() => setPosViewMode('grid')}
                      className={`p-1.5 rounded-lg transition-colors cursor-pointer ${
                        posViewMode === 'grid' ? 'bg-[#ff5500] text-white shadow-xs' : 'text-slate-400 hover:text-white'
                      }`}
                      title="Visual Grid"
                    >
                      <LayoutGrid className="h-3.5 w-3.5" />
                    </button>
                    <button
                      type="button"
                      onClick={() => setPosViewMode('compact')}
                      className={`p-1.5 rounded-lg transition-colors cursor-pointer ${
                        posViewMode === 'compact' ? 'bg-[#ff5500] text-white shadow-xs' : 'text-slate-400 hover:text-white'
                      }`}
                      title="Compact Tiles"
                    >
                      <Grid className="h-3.5 w-3.5" />
                    </button>
                    <button
                      type="button"
                      onClick={() => setPosViewMode('list')}
                      className={`p-1.5 rounded-lg transition-colors cursor-pointer ${
                        posViewMode === 'list' ? 'bg-[#ff5500] text-white shadow-xs' : 'text-slate-400 hover:text-white'
                      }`}
                      title="Compact List"
                    >
                      <Layers className="h-3.5 w-3.5" />
                    </button>
                  </div>

                  {/* Search Input */}
                  <div className="relative flex-1 sm:w-48">
                    <Search className="h-3.5 w-3.5 absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
                    <input
                      type="text"
                      value={menuSearchQuery}
                      onChange={e => setMenuSearchQuery(e.target.value)}
                      placeholder="Search items..."
                      className="w-full pl-8 pr-3 py-1.5 bg-slate-800 border border-slate-700 rounded-xl text-xs text-slate-100 placeholder-slate-400 focus:outline-none focus:border-[#ff5500]"
                    />
                  </div>
                </div>
              </div>

              {/* Responsive Catalog Items */}
              <div className="flex-1 overflow-y-auto pr-1 min-h-0 pb-20 lg:pb-0">
                {(() => {
                  const filteredDishes = menuItems.filter(item => {
                    const matchCat = selectedCategory === 'All' || item.category === selectedCategory;
                    const matchQuery = item.name.toLowerCase().includes(menuSearchQuery.toLowerCase());
                    return matchCat && matchQuery;
                  });

                  if (filteredDishes.length === 0) {
                    return (
                      <div className="h-64 flex flex-col items-center justify-center text-slate-400 text-center">
                        <Coffee className="h-10 w-10 mb-2 stroke-[1]" />
                        <p className="text-xs font-bold text-slate-600">No matching menu items</p>
                        <p className="text-[11px] text-slate-400 mt-0.5">Try searching a different item or category.</p>
                      </div>
                    );
                  }

                  if (posViewMode === 'compact') {
                    return (
                      <div className="bg-slate-100/70 p-2 sm:p-3 rounded-2xl border border-slate-200 shadow-inner">
                        <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-4 lg:grid-cols-4 xl:grid-cols-6 gap-2 sm:gap-2.5">
                          {filteredDishes.map(dish => {
                            const { portions } = calculateDishAvailability(dish.recipe);
                            const isInCart = cart.some(i => i.id === dish.id);

                            return (
                              <button
                                key={dish.id}
                                type="button"
                                onClick={() => handleAddToCart(dish)}
                                className={`p-2.5 sm:p-3 rounded-xl border-2 text-left flex flex-col justify-between transition-all bg-white cursor-pointer active:scale-95 shadow-xs ${
                                  isInCart
                                    ? 'border-[#ff5500] ring-2 ring-[#ff5500]/25 shadow-orange-500/10'
                                    : 'border-slate-300 hover:border-slate-400'
                                }`}
                              >
                                <div>
                                  <span className={`text-[8px] sm:text-[9px] font-black px-1.5 py-0.2 rounded uppercase ${
                                    dish.department === 'Bar' ? 'bg-indigo-50 text-indigo-700' : 'bg-rose-50 text-rose-700'
                                  }`}>
                                    {dish.department}
                                  </span>
                                  <p className="font-black text-xs text-slate-950 mt-1 line-clamp-1">{dish.name}</p>
                                </div>
                                <div className="mt-2 pt-1.5 border-t border-slate-100 flex justify-between items-center text-[10px]">
                                  <span className="font-mono font-black text-[#ff5500]">
                                    {settings.currency} {dish.price.toFixed(0)}
                                  </span>
                                  <span className={`font-bold ${portions <= 0 ? 'text-amber-600' : 'text-emerald-700'}`}>
                                    {portions <= 0 ? '0 left' : `${portions} left`}
                                  </span>
                                </div>
                              </button>
                            );
                          })}
                        </div>
                      </div>
                    );
                  }

                  if (posViewMode === 'list') {
                    return (
                      <div className="bg-white rounded-2xl border border-slate-200 overflow-hidden divide-y divide-slate-100">
                        {filteredDishes.map(dish => {
                          const { portions } = calculateDishAvailability(dish.recipe);
                          return (
                            <div
                              key={dish.id}
                              onClick={() => handleAddToCart(dish)}
                              className="p-3 flex items-center justify-between hover:bg-slate-50 cursor-pointer transition-colors"
                            >
                              <div className="flex items-center gap-2.5 sm:gap-3">
                                <span className={`text-[9px] font-black px-1.5 py-0.5 rounded ${
                                  dish.department === 'Bar' ? 'bg-indigo-100 text-indigo-700' : 'bg-rose-100 text-rose-700'
                                }`}>
                                  {dish.department}
                                </span>
                                <div>
                                  <p className="font-bold text-xs text-slate-900">{dish.name}</p>
                                  <span className="text-[10px] text-slate-400">{dish.category} • {dish.prepTime}</span>
                                </div>
                              </div>
                              <div className="flex items-center gap-3 sm:gap-4">
                                <span className={`text-xs font-mono hidden sm:inline ${portions <= 0 ? 'text-amber-600 font-bold' : 'text-slate-500'}`}>
                                  {portions <= 0 ? '0 ready' : `${portions} ready`}
                                </span>
                                <span className="font-mono font-bold text-xs sm:text-sm text-[#ff5500]">
                                  {settings.currency} {dish.price.toFixed(2)}
                                </span>
                                <button type="button" className="px-2 py-1 bg-slate-100 text-slate-700 rounded-lg text-xs font-bold">+ Add</button>
                              </div>
                            </div>
                          );
                        })}
                      </div>
                    );
                  }

                  // Default Grid Layout with multi-screen scaling
                  return (
                    <div className="bg-slate-100/80 p-2.5 sm:p-3.5 rounded-3xl border border-slate-200/90 shadow-inner">
                      <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-4 lg:grid-cols-3 xl:grid-cols-4 2xl:grid-cols-6 gap-2.5 sm:gap-3.5">
                        {filteredDishes.map(dish => {
                          const { cogs, portions } = calculateDishAvailability(dish.recipe);
                          const isInCart = cart.some(i => i.id === dish.id);

                          return (
                            <div
                              key={dish.id}
                              onClick={() => handleAddToCart(dish)}
                              className={`bg-white rounded-2xl border-2 cursor-pointer active:scale-[0.98] overflow-hidden flex flex-col justify-between transition-all duration-150 shadow-xs hover:shadow-md ${
                                isInCart
                                  ? 'border-[#ff5500] ring-2 ring-[#ff5500]/25 shadow-orange-500/15'
                                  : 'border-slate-300 hover:border-slate-400'
                              }`}
                            >
                              {dish.imageUrl ? (
                                <div className="relative h-20 sm:h-24 md:h-28 w-full bg-slate-100 overflow-hidden shrink-0 border-b border-slate-200">
                                  <img
                                    src={dish.imageUrl}
                                    alt={dish.name}
                                    className="w-full h-full object-cover"
                                    onError={(e) => { e.currentTarget.style.display = 'none'; }}
                                  />
                                  <span className={`absolute top-1.5 left-1.5 text-[8px] sm:text-[9px] font-black uppercase px-1.5 py-0.2 rounded text-white ${
                                    dish.department === 'Bar' ? 'bg-indigo-600' : 'bg-rose-600'
                                  }`}>
                                    {dish.department}
                                  </span>
                                  <span className="absolute bottom-1.5 right-1.5 px-2 py-0.5 rounded-md bg-black/80 backdrop-blur-xs text-white font-mono font-black text-[10px] sm:text-xs">
                                    {settings.currency} {dish.price.toFixed(0)}
                                  </span>
                                </div>
                              ) : (
                                <div className="p-2.5 pb-0 flex justify-between items-start">
                                  <span className={`text-[8px] sm:text-[9px] font-black uppercase px-1.5 py-0.2 rounded ${
                                    dish.department === 'Bar' ? 'bg-indigo-50 text-indigo-700' : 'bg-rose-50 text-rose-700'
                                  }`}>
                                    {dish.department}
                                  </span>
                                  <span className="text-xs font-mono font-black text-[#ff5500]">
                                    {settings.currency} {dish.price.toFixed(0)}
                                  </span>
                                </div>
                              )}

                              <div className="p-2.5 sm:p-3 flex-1 flex flex-col justify-center">
                                <h4 className="font-black text-xs sm:text-sm text-slate-950 leading-snug line-clamp-2">
                                  {dish.name}
                                </h4>
                                <p className="text-[10px] text-slate-500 font-bold mt-0.5 line-clamp-1 uppercase">
                                  {dish.category || 'General Menu'}
                                </p>
                              </div>

                              <div className="p-2 sm:p-2.5 border-t border-slate-200/90 flex items-center justify-between text-[10px] bg-slate-50/70">
                                <span className={`font-bold ${portions <= 0 ? 'text-amber-600' : 'text-emerald-700'}`}>
                                  {portions <= 0 ? '0 ready' : `${portions} left`}
                                </span>
                                <span className="text-slate-400 font-mono hidden sm:inline">
                                  BOM: {cogs.toFixed(0)}
                                </span>
                              </div>
                            </div>
                          );
                        })}
                      </div>
                    </div>
                  );
                })()}
              </div>
            </div>

            {/* FLOATING ACTION BAR FOR MOBILE/TABLET (Bottom Dock) */}
            <div className="lg:hidden fixed bottom-0 inset-x-0 bg-slate-900 border-t border-slate-800 p-3 z-30 flex items-center justify-between shadow-2xl backdrop-blur-md">
              <div className="flex items-center gap-3">
                <div className="relative">
                  <button
                    type="button"
                    onClick={() => setMobileCartDrawerOpen(true)}
                    className="h-11 w-11 rounded-xl bg-orange-600 text-white flex items-center justify-center shadow-lg shadow-orange-600/40 cursor-pointer active:scale-95"
                  >
                    <ShoppingBag className="h-5 w-5" />
                  </button>
                  {cart.length > 0 && (
                    <span className="absolute -top-1.5 -right-1.5 h-5 w-5 rounded-full bg-emerald-500 text-white text-[10px] font-black flex items-center justify-center border-2 border-slate-900">
                      {cart.reduce((sum, item) => sum + item.qty, 0)}
                    </span>
                  )}
                </div>
                <div>
                  <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider block">
                    {orderMode === 'DINING' ? selectedTable.name : 'Takeaway'}
                  </span>
                  <span className="text-sm font-black font-mono text-white">
                    {settings.currency} {cartGrandTotal.toFixed(2)}
                  </span>
                </div>
              </div>

              <div className="flex items-center gap-2">
                <button
                  type="button"
                  disabled={cart.length === 0}
                  onClick={() => setMobileCartDrawerOpen(true)}
                  className="px-3 py-2 bg-slate-800 hover:bg-slate-700 text-white text-xs font-bold rounded-xl transition-all cursor-pointer disabled:opacity-30"
                >
                  View Bill
                </button>
                <button
                  type="button"
                  disabled={cart.length === 0}
                  onClick={handleSendOrder}
                  className="px-4 py-2 bg-[#ff5500] hover:bg-orange-600 text-white text-xs font-black uppercase tracking-wider rounded-xl shadow-lg shadow-orange-600/30 transition-all cursor-pointer disabled:opacity-30 flex items-center gap-1.5"
                >
                  <Send className="h-3.5 w-3.5" />
                  <span>Send</span>
                </button>
              </div>
            </div>

            {/* MOBILE CART BACKDROP OVERLAY */}
            {mobileCartDrawerOpen && (
              <div
                onClick={() => setMobileCartDrawerOpen(false)}
                className="lg:hidden fixed inset-0 bg-black/70 backdrop-blur-xs z-40 transition-opacity"
              />
            )}

            {/* RIGHT TICKET BAR (Side pane on Desktop / Slide-up Bottom Sheet on Mobile & Tablet) */}
            <div
              className={`fixed lg:static inset-x-0 bottom-0 lg:inset-auto z-50 lg:z-auto w-full lg:w-96 max-h-[85vh] lg:max-h-full bg-white border-t lg:border-t-0 lg:border-l border-slate-200 flex flex-col justify-between shrink-0 shadow-2xl lg:shadow-none rounded-t-3xl lg:rounded-none transition-transform duration-300 ease-in-out ${
                mobileCartDrawerOpen ? 'translate-y-0' : 'translate-y-full lg:translate-y-0'
              }`}
            >
              {/* Mobile Drawer Handle Header */}
              <div className="p-3 border-b border-slate-100 flex lg:hidden items-center justify-between">
                <div className="h-1.5 w-12 bg-slate-300 rounded-full mx-auto" />
                <button
                  type="button"
                  onClick={() => setMobileCartDrawerOpen(false)}
                  className="p-1 text-slate-400 hover:text-slate-900 absolute right-3"
                >
                  <X className="h-5 w-5" />
                </button>
              </div>

              {/* Order Assignment Controls */}
              <div className="p-3.5 border-b border-slate-200 space-y-2.5 shrink-0 bg-white">
                {/* Modifying Existing Bill Alert */}
                {settlingOrder && (
                  <div className="p-2.5 bg-amber-500/10 border-2 border-amber-500/40 rounded-xl flex items-center justify-between text-xs">
                    <div>
                      <span className="font-black text-amber-900 block leading-tight">
                        Modifying Bill: {settlingOrder.tableName}
                      </span>
                      <span className="text-[10px] text-amber-700 font-mono font-bold">
                        Ref #{settlingOrder.orderId}
                      </span>
                    </div>
                    <button
                      type="button"
                      onClick={() => {
                        setSettlingOrder(null);
                        setCart([]);
                      }}
                      className="px-2 py-1 bg-amber-500/20 hover:bg-amber-500/30 text-amber-900 text-[10px] font-black rounded-lg transition-colors cursor-pointer"
                    >
                      Cancel Edit
                    </button>
                  </div>
                )}

                {/* Dine-In vs Takeaway Selector */}
                <div>
                  <span className="text-[10px] font-black tracking-wider uppercase text-slate-400 block mb-1.5">
                    Order Assignment
                  </span>
                  <div className="grid grid-cols-2 gap-2">
                    <button
                      type="button"
                      onClick={() => setOrderMode('DINING')}
                      className={`flex items-center justify-between px-2.5 py-2 rounded-xl border text-xs font-bold transition-all cursor-pointer ${
                        orderMode === 'DINING'
                          ? 'bg-orange-50/80 border-[#ff5500] text-slate-900 shadow-xs'
                          : 'bg-slate-50 border-slate-200 text-slate-500 hover:bg-slate-100'
                      }`}
                    >
                      <div className="flex items-center gap-1.5">
                        <Grid className="h-3.5 w-3.5 text-[#ff5500]" />
                        <span>Dine-In</span>
                      </div>
                      <div className={`h-4 w-4 rounded-full flex items-center justify-center border ${
                        orderMode === 'DINING' ? 'bg-[#ff5500] border-[#ff5500] text-white' : 'border-slate-300 bg-white'
                      }`}>
                        {orderMode === 'DINING' && <Check className="h-2.5 w-2.5 stroke-[3]" />}
                      </div>
                    </button>

                    <button
                      type="button"
                      onClick={() => setOrderMode('TAKEAWAY')}
                      className={`flex items-center justify-between px-2.5 py-2 rounded-xl border text-xs font-bold transition-all cursor-pointer ${
                        orderMode === 'TAKEAWAY'
                          ? 'bg-orange-50/80 border-[#ff5500] text-slate-900 shadow-xs'
                          : 'bg-slate-50 border-slate-200 text-slate-500 hover:bg-slate-100'
                      }`}
                    >
                      <div className="flex items-center gap-1.5">
                        <ShoppingBag className="h-3.5 w-3.5 text-[#ff5500]" />
                        <span>Takeaway</span>
                      </div>
                      <div className={`h-4 w-4 rounded-full flex items-center justify-center border ${
                        orderMode === 'TAKEAWAY' ? 'bg-[#ff5500] border-[#ff5500] text-white' : 'border-slate-300 bg-white'
                      }`}>
                        {orderMode === 'TAKEAWAY' && <Check className="h-2.5 w-2.5 stroke-[3]" />}
                      </div>
                    </button>
                  </div>
                </div>

                {/* Table Picker or Takeaway Information */}
                {orderMode === 'DINING' ? (
                  <div className="p-2.5 bg-slate-50 border border-slate-200 rounded-xl space-y-2">
                    <div className="flex items-center justify-between text-xs">
                      <span className="font-bold text-slate-700">Assigned Table:</span>
                      <button
                        type="button"
                        onClick={() => setAllocationModalOpen(true)}
                        className="text-[10px] font-bold text-[#ff5500] hover:underline flex items-center gap-1 cursor-pointer"
                      >
                        <LayoutGrid className="h-3 w-3" /> View Floor Map
                      </button>
                    </div>

                    <div className="flex items-center gap-2">
                      <select
                        value={selectedTable.id}
                        onChange={(e) => {
                          const tbl = floorTables.find(t => t.id === e.target.value);
                          if (tbl) setSelectedTable(tbl);
                        }}
                        className="flex-1 px-2.5 py-1.5 bg-white border border-slate-200 rounded-lg text-xs font-bold text-slate-900 focus:outline-none focus:border-[#ff5500]"
                      >
                        {floorTables.map(t => (
                          <option key={t.id} value={t.id}>
                            {t.name} ({t.zone}) - {t.capacity}p [{t.status}]
                          </option>
                        ))}
                      </select>

                      <div className="flex items-center bg-white border border-slate-200 rounded-lg px-1.5 py-0.5" title="Guest Count">
                        <button
                          type="button"
                          onClick={() => setGuestCount(Math.max(1, guestCount - 1))}
                          className="text-xs font-bold px-1 text-slate-500 hover:text-slate-900 cursor-pointer"
                        >
                          -
                        </button>
                        <span className="text-xs font-bold px-1.5 font-mono text-slate-800">{guestCount}p</span>
                        <button
                          type="button"
                          onClick={() => setGuestCount(guestCount + 1)}
                          className="text-xs font-bold px-1 text-slate-500 hover:text-slate-900 cursor-pointer"
                        >
                          +
                        </button>
                      </div>
                    </div>
                  </div>
                ) : (
                  <div className="p-2.5 bg-slate-50 border border-slate-200 rounded-xl space-y-2">
                    <div className="flex items-center justify-between text-xs">
                      <span className="font-bold text-slate-700">Takeaway Details:</span>
                      <span className="font-mono font-bold text-xs bg-orange-100 text-[#ff5500] px-2 py-0.5 rounded-md">
                        Token: {takeawayInfo.token}
                      </span>
                    </div>
                    <div className="grid grid-cols-2 gap-2">
                      <input
                        type="text"
                        value={takeawayInfo.name}
                        onChange={(e) => setTakeawayInfo(prev => ({ ...prev, name: e.target.value }))}
                        placeholder="Guest Name"
                        className="px-2.5 py-1.5 bg-white border border-slate-200 rounded-lg text-xs text-slate-900 focus:outline-none focus:border-[#ff5500]"
                      />
                      <input
                        type="text"
                        value={takeawayInfo.phone}
                        onChange={(e) => setTakeawayInfo(prev => ({ ...prev, phone: e.target.value }))}
                        placeholder="Phone (Optional)"
                        className="px-2.5 py-1.5 bg-white border border-slate-200 rounded-lg text-xs font-mono text-slate-900 focus:outline-none focus:border-[#ff5500]"
                      />
                    </div>
                  </div>
                )}

                {/* Surcharges & Rates Toggle */}
                <div className="grid grid-cols-2 gap-2 pt-1">
                  <button
                    type="button"
                    onClick={() => setServiceChargeActive(!serviceChargeActive)}
                    className={`py-1 px-2 rounded-lg border text-xs font-bold transition-all cursor-pointer ${
                      serviceChargeActive
                        ? 'bg-emerald-50 border-emerald-300 text-emerald-700'
                        : 'bg-slate-50 border-slate-200 text-slate-400'
                    }`}
                  >
                    Service ({settings.serviceChargeRate}%): {serviceChargeActive ? 'ON' : 'OFF'}
                  </button>

                  <button
                    type="button"
                    onClick={() => setTaxActive(!taxActive)}
                    className={`py-1 px-2 rounded-lg border text-xs font-bold transition-all cursor-pointer ${
                      taxActive
                        ? 'bg-indigo-50 border-indigo-300 text-indigo-700'
                        : 'bg-slate-50 border-slate-200 text-slate-400'
                    }`}
                  >
                    Tax ({settings.taxRate}%): {taxActive ? 'ON' : 'OFF'}
                  </button>
                </div>

                {/* Quick Discounts */}
                <div>
                  <span className="text-[10px] font-bold text-slate-400 uppercase">Discount:</span>
                  <div className="flex gap-1 mt-1">
                    {[0, 5, 10, 15, 20].map(pct => (
                      <button
                        key={pct}
                        type="button"
                        onClick={() => setDiscountPercent(pct)}
                        className={`flex-1 py-0.5 rounded text-[10px] font-bold border transition-colors cursor-pointer ${
                          discountPercent === pct
                            ? 'bg-slate-900 text-white border-slate-900'
                            : 'bg-slate-50 text-slate-600 border-slate-200 hover:bg-slate-100'
                        }`}
                      >
                        {pct === 0 ? 'None' : `${pct}%`}
                      </button>
                    ))}
                  </div>
                </div>
              </div>

              {/* Cart Items List */}
              <div className="flex-1 overflow-y-auto p-3.5 space-y-2 min-h-0">
                {cart.length === 0 ? (
                  <div className="h-full flex flex-col items-center justify-center text-slate-400 text-center p-6">
                    <Monitor className="h-10 w-10 mb-2 stroke-[1]" />
                    <p className="text-xs font-bold text-slate-600">Ticket is empty</p>
                    <p className="text-[11px] text-slate-400 mt-1">Tap items to build order.</p>
                  </div>
                ) : (
                  cart.map(item => (
                    <div key={item.cartItemId} className="p-2.5 sm:p-3 bg-slate-50 rounded-xl border border-slate-200">
                      <div className="flex items-start justify-between">
                        <div>
                          <p className="text-xs font-bold text-slate-900">{item.name}</p>
                          <p className="text-xs font-mono font-bold text-[#ff5500] mt-0.5">
                            {settings.currency} {(item.price * item.qty).toFixed(2)}
                          </p>
                        </div>
                        <div className="flex items-center gap-1.5">
                          <button
                            type="button"
                            onClick={() => {
                              setCart(prev => prev.map(i => i.cartItemId === item.cartItemId ? { ...i, qty: Math.max(1, i.qty - 1) } : i));
                            }}
                            className="h-6 w-6 rounded-md bg-white border border-slate-200 text-slate-600 flex items-center justify-center text-xs font-bold cursor-pointer"
                          >
                            -
                          </button>
                          <span className="text-xs font-bold w-5 text-center font-mono">{item.qty}</span>
                          <button
                            type="button"
                            onClick={() => {
                              setCart(prev => prev.map(i => i.cartItemId === item.cartItemId ? { ...i, qty: i.qty + 1 } : i));
                            }}
                            className="h-6 w-6 rounded-md bg-white border border-slate-200 text-slate-600 flex items-center justify-center text-xs font-bold cursor-pointer"
                          >
                            +
                          </button>
                          <button
                            type="button"
                            onClick={() => {
                              setCart(prev => prev.filter(i => i.cartItemId !== item.cartItemId));
                            }}
                            className="h-6 w-6 rounded-md bg-white border border-slate-200 text-slate-600 flex items-center justify-center text-xs font-bold hover:bg-slate-100 cursor-pointer"
                          >
                            <Trash2 className="h-3.5 w-3.5" />
                          </button>
                        </div>
                      </div>

                      <input
                        type="text"
                        value={item.notes || ''}
                        onChange={e => {
                          const val = e.target.value;
                          setCart(prev => prev.map(i => i.cartItemId === item.cartItemId ? { ...i, notes: val } : i));
                        }}
                        placeholder="Add kitchen/bar modifier note..."
                        className="w-full mt-2 text-[11px] px-2 py-1 bg-white border border-slate-200 rounded text-slate-700 placeholder-slate-400 focus:outline-none focus:border-[#ff5500]"
                      />
                    </div>
                  ))
                )}
              </div>

              {/* Ticket Financial Totals & Actions */}
              <div className="p-3.5 sm:p-4 border-t border-slate-200 bg-slate-50 space-y-2.5 shrink-0">
                <div className="space-y-1 text-xs text-slate-600">
                  <div className="flex justify-between">
                    <span>Subtotal</span>
                    <span className="font-mono font-bold text-slate-900">{settings.currency} {cartSubtotal.toFixed(2)}</span>
                  </div>
                  {discountPercent > 0 && (
                    <div className="flex justify-between text-rose-600 font-medium">
                      <span>Discount ({discountPercent}%)</span>
                      <span className="font-mono">-{settings.currency} {cartDiscountAmount.toFixed(2)}</span>
                    </div>
                  )}
                  {serviceChargeActive && (
                    <div className="flex justify-between text-emerald-700 font-medium">
                      <span>Service Charge ({settings.serviceChargeRate}%)</span>
                      <span className="font-mono">+{settings.currency} {cartServiceCharge.toFixed(2)}</span>
                    </div>
                  )}
                  {taxActive && (
                    <div className="flex justify-between text-indigo-700 font-medium">
                      <span>Taxes ({settings.taxRate}%):</span>
                      <span className="font-mono">+{settings.currency} {cartTax.toFixed(2)}</span>
                    </div>
                  )}
                  <div className="flex justify-between text-sm font-black text-slate-900 pt-1.5 border-t border-slate-200">
                    <span>Grand Total</span>
                    <span className="font-mono text-base text-[#ff5500]">{settings.currency} {cartGrandTotal.toFixed(2)}</span>
                  </div>
                </div>

                <button
                  type="button"
                  disabled={cart.length === 0}
                  onClick={() => {
                    handleSendOrder();
                    setMobileCartDrawerOpen(false);
                  }}
                  className="w-full py-2.5 sm:py-3 bg-[#ff5500] hover:bg-orange-600 text-white font-extrabold rounded-xl text-xs uppercase tracking-wider flex items-center justify-center gap-2 shadow-sm shadow-orange-600/30 disabled:opacity-40 cursor-pointer"
                >
                  <Send className="h-4 w-4" />
                  <span>{settlingOrder ? 'Save & Send Add-On (KOT / BOT)' : 'Send Order (Prints KOT / BOT)'}</span>
                </button>

                <button
                  type="button"
                  disabled={cart.length === 0}
                  onClick={() => {
                    setPaymentMethod('CASH');
                    setCheckoutModalOpen(true);
                    setMobileCartDrawerOpen(false);
                  }}
                  className="w-full py-2 bg-slate-900 hover:bg-slate-800 text-white font-bold rounded-xl text-xs flex items-center justify-center gap-2 disabled:opacity-40 cursor-pointer"
                >
                  <Receipt className="h-3.5 w-3.5 text-emerald-400" />
                  <span>Direct Settle &amp; Pay ({settings.currency} {cartGrandTotal.toFixed(2)})</span>
                </button>
              </div>
            </div>
          </div>
        )}
        
        {/* VIEW 2: BILLING & SETTLEMENT QUEUE */}
        {activeTab === 'billing' && (
          <div className="flex-1 overflow-y-auto p-4 sm:p-6 space-y-6">
            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2">
              <div>
                <h2 className="text-xl font-black text-slate-900">Billing &amp; Settlement Queue</h2>
                <p className="text-xs text-slate-500 mt-0.5">
                  Manage active tables, edit line items, print temporary bills, and settle final accounts.
                </p>
              </div>
              <span className="self-start sm:self-auto px-3 py-1.5 bg-white border border-slate-200 rounded-xl text-xs font-mono font-bold text-slate-700">
                {activeOrders.length} Open Bills
              </span>
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-2 xl:grid-cols-3 gap-3.5 sm:gap-4">
              {activeOrders.length === 0 ? (
                <div className="col-span-full h-64 bg-white rounded-2xl border border-dashed border-slate-300 flex flex-col items-center justify-center text-slate-400">
                  <Receipt className="h-10 w-10 mb-2 stroke-[1]" />
                  <p className="text-sm font-bold text-slate-700">No active tables pending billing</p>
                  <p className="text-xs mt-1">Send an order from the POS Terminal to populate this list.</p>
                </div>
              ) : (
                activeOrders.map(order => {
                  const fin = calculateOrderFinancials(order);
                  return (
                    <div key={order.orderId} className="bg-white rounded-2xl border border-slate-200 p-4 sm:p-5 shadow-xs flex flex-col justify-between">
                      <div>
                        <div className="flex items-start justify-between mb-2">
                          <div>
                            <span className="text-[10px] font-bold px-2 py-0.5 rounded-full bg-orange-100 text-[#ff5500]">
                              {order.mode}
                            </span>
                            <h3 className="text-base font-extrabold text-slate-900 mt-1">{order.tableName}</h3>
                            <p className="text-xs text-slate-500">Waitstaff: {order.server} • {order.sentAt}</p>
                          </div>
                          <span className="text-lg font-black font-mono text-[#ff5500]">
                            {settings.currency} {fin.total.toFixed(2)}
                          </span>
                        </div>

                        <div className="bg-slate-50 rounded-xl p-3 my-3 space-y-1.5 text-xs max-h-40 overflow-y-auto border border-slate-100">
                          {order.items.map((item, idx) => (
                            <div key={idx} className="flex justify-between">
                              <span className="font-bold text-slate-800">{item.qty}x {item.name}</span>
                              <span className="font-mono text-slate-500">{settings.currency} {(item.price * item.qty).toFixed(2)}</span>
                            </div>
                          ))}
                        </div>
                      </div>

                      <div className="pt-2 border-t border-slate-100 space-y-2">
                        <div className="grid grid-cols-3 gap-1.5 sm:gap-2 text-[11px] sm:text-xs">
                          {/* Edit Items in POS */}
                          <button
                            type="button"
                            onClick={() => {
                              setCart(order.items ? JSON.parse(JSON.stringify(order.items)) : []);
                              setOrderMode(order.mode || 'DINING');
                              if (order.tableId) {
                                const tbl = floorTables.find(t => t.id === order.tableId);
                                if (tbl) setSelectedTable(tbl);
                              } else {
                                setTakeawayInfo(prev => ({
                                  ...prev,
                                  name: order.customerName || 'Walk-in Guest',
                                  token: order.tableName || 'TK-101'
                                }));
                              }
                              setServiceChargeActive(order.serviceChargeActive !== false);
                              setTaxActive(Boolean(order.taxActive));
                              setDiscountPercent(order.discountPercent || 0);
                              setSettlingOrder(order);
                              setActiveTab('pos');
                            }}
                            className="py-2.5 sm:py-2 px-1 bg-slate-100 hover:bg-slate-200 rounded-xl text-slate-700 font-bold flex items-center justify-center gap-1 cursor-pointer transition-colors"
                          >
                            <Edit3 className="h-3 w-3 text-[#ff5500] shrink-0" />
                            <span className="truncate">Edit in POS</span>
                          </button>

                          {/* Temp Bill */}
                          <button
                            type="button"
                            onClick={() => {
                              triggerAutoPrint({
                                type: 'TEMP_BILL',
                                data: {
                                  table: order.tableName,
                                  server: order.server,
                                  items: order.items,
                                  subtotal: fin.subtotal,
                                  discount: fin.discount,
                                  service: fin.service,
                                  tax: fin.tax,
                                  total: fin.total
                                }
                              }, `Proforma Bill for ${order.tableName}`);
                            }}
                            className="py-2.5 sm:py-2 px-1 bg-indigo-50 hover:bg-indigo-100 rounded-xl text-indigo-700 font-bold flex items-center justify-center gap-1 cursor-pointer transition-colors"
                          >
                            <Printer className="h-3 w-3 shrink-0" />
                            <span className="truncate">Temp Bill</span>
                          </button>

                          {/* Settle */}
                          <button
                            type="button"
                            onClick={() => {
                              setSettlingOrder(order);
                              setPaymentMethod('CASH');
                              setCheckoutModalOpen(true);
                            }}
                            className="py-2.5 sm:py-2 px-1 bg-emerald-600 hover:bg-emerald-700 text-white rounded-xl font-bold flex items-center justify-center gap-1 shadow-xs cursor-pointer transition-colors"
                          >
                            <DollarSign className="h-3 w-3 shrink-0" />
                            <span className="truncate">Settle</span>
                          </button>
                        </div>

                        {/* Admin Delete */}
                        {currentUser.role === 'Administrator' && (
                          <button
                            type="button"
                            onClick={() => {
                              setActiveOrders(prev => prev.filter(o => o.orderId !== order.orderId));
                              if (order.tableId) {
                                setFloorTables(prev => prev.map(t => t.id === order.tableId ? { ...t, status: 'VACANT', currentOrderRef: null } : t));
                              }
                              recordAuditLog('ADMIN_DELETE_ACTIVE_BILL', order.orderId, `Admin deleted open bill ${order.orderId} (${order.tableName})`);
                            }}
                            className="w-full py-1.5 bg-rose-50 hover:bg-rose-100 text-rose-700 rounded-xl font-bold text-[11px] flex items-center justify-center gap-1 transition-colors border border-rose-200 cursor-pointer"
                          >
                            <Trash2 className="h-3 w-3" /> Delete Active Bill (Admin)
                          </button>
                        )}
                      </div>
                    </div>
                  );
                })
              )}
            </div>
          </div>
        )}

        {/* VIEW 3: SALES & REVENUE REPORTS WITH DATE FILTERS */}
        {activeTab === 'reports' && (
          <div className="flex-1 overflow-y-auto p-6 space-y-6">
            <div className="flex items-center gap-2 sm:gap-6 border-b border-slate-200 pb-2.5 text-xs font-bold overflow-x-auto scrollbar-thin">
              {['Daily Overview', 'All Items Sales', 'Sales Detail', 'Cash Out Report', 'KOT Report', 'BOT Report', 'Sales Summary', 'Food vs Beverage', 'Stock Usage', 'Stock Movement Ledger', 'Audit Trail'].map(sub => (
                <button
                  key={sub}
                  type="button"
                  onClick={() => setReportSubTab(sub)}
                  className={`transition-colors relative pb-1 whitespace-nowrap cursor-pointer ${
                    reportSubTab === sub
                      ? 'text-[#ff5500] after:absolute after:bottom-0 after:left-0 after:w-full after:h-0.5 after:bg-[#ff5500]'
                      : 'text-slate-500 hover:text-slate-900'
                  }`}
                >
                  {sub}
                </button>
              ))}
            </div>

            {/* Date Filters & Export Action Bar */}
            <div className="bg-white p-3 sm:p-4 rounded-2xl border border-slate-200 flex flex-col xl:flex-row items-stretch xl:items-center justify-between gap-3 shadow-xs">
              <div className="flex flex-wrap items-center gap-2">
                <span className="text-xs font-bold text-slate-700 flex items-center gap-1">
                  <Calendar className="h-3.5 w-3.5 text-[#ff5500]" /> Date Filter:
                </span>
                <input
                  type="date"
                  value={reportStartDate}
                  onChange={e => setReportStartDate(e.target.value)}
                  className="px-2.5 py-1 bg-slate-50 border border-slate-200 rounded-xl text-xs font-mono text-slate-800"
                />
                <span className="text-slate-400 text-xs">to</span>
                <input
                  type="date"
                  value={reportEndDate}
                  onChange={e => setReportEndDate(e.target.value)}
                  className="px-2.5 py-1 bg-slate-50 border border-slate-200 rounded-xl text-xs font-mono text-slate-800"
                />
              </div>

              {/* Presets & Export Actions */}
              <div className="flex flex-wrap items-center gap-2">
                <div className="flex items-center gap-1.5 overflow-x-auto">
                  {[
                    { label: 'Today', start: getLocalDateStr(), end: getLocalDateStr() },
                    { label: 'Yesterday', start: getLocalDateStr(new Date(Date.now() - 86400000)), end: getLocalDateStr(new Date(Date.now() - 86400000)) },
                    { label: 'Last 7 Days', start: getLocalDateStr(new Date(Date.now() - 7 * 86400000)), end: getLocalDateStr() },
                    { label: 'All Time', start: '', end: '' }
                  ].map(preset => (
                    <button
                      key={preset.label}
                      type="button"
                      onClick={() => {
                        setReportStartDate(preset.start);
                        setReportEndDate(preset.end);
                      }}
                      className="px-2.5 py-1 bg-slate-100 hover:bg-slate-200 text-slate-700 text-[11px] font-bold rounded-lg whitespace-nowrap cursor-pointer transition-colors"
                    >
                      {preset.label}
                    </button>
                  ))}
                </div>

                <div className="h-5 w-px bg-slate-200 hidden sm:block" />

                {/* 1. EXCEL EXPORT BUTTON */}
                <button
                  type="button"
                  onClick={() => {
                    let exportRows = [];
                    const prefix = `${reportSubTab.toLowerCase().replace(/\s+/g, '_')}`;

                    if (reportSubTab === 'Daily Overview' || reportSubTab === 'All Items Sales') {
                      exportRows = salesMetrics.topItems.map(item => ({
                        'Menu Item': item.name,
                        'Area': item.department,
                        'Category': item.category,
                        'Portions Sold': item.sold,
                        'Unit Price': item.unitPrice,
                        'Total Revenue': item.revenue
                      }));
                    } else if (reportSubTab === 'Sales Detail') {
                      exportRows = filteredTransactions.map(t => ({
                        'Invoice #': t.invoiceNo,
                        'Date': t.date,
                        'Table': t.table,
                        'Cashier': t.cashier,
                        'Method': t.paymentMethod,
                        'Subtotal': t.subtotal,
                        'Service Charge': t.serviceCharge,
                        'Tax': t.tax,
                        'Discount': t.discount || 0,
                        'Grand Total': t.total
                      }));
                    } else if (reportSubTab === 'Cash Out Report') {
                      const allRecords = Array.isArray(expenses) ? expenses : [];
                      exportRows = allRecords
                        .filter(exp => {
                          const expDate = exp.date || extractDateStr(exp.createdAt);
                          if (!expDate) return true;
                          if (reportStartDate && expDate < reportStartDate) return false;
                          if (reportEndDate && expDate > reportEndDate) return false;
                          return true;
                        })
                        .map(e => ({
                          'Voucher Ref': e.id,
                          'Date': e.date || getLocalDateStr(),
                          'Shift ID': e.shiftId || 'N/A',
                          'Category': e.category,
                          'Recipient': e.recipient || 'N/A',
                          'Description': e.reason,
                          'Amount': parseFloat(e.amount) || 0,
                          'Status': e.status || 'APPROVED',
                          'Approved By': e.approvedBy || 'N/A'
                        }));
                    } else if (reportSubTab === 'KOT Report') {
                      exportRows = salesMetrics.topItems
                        .filter(i => i.department === 'Kitchen')
                        .map(i => ({
                          'Kitchen Dish': i.name,
                          'Category': i.category,
                          'Portions Prepared': i.sold,
                          'Total Revenue': i.revenue
                        }));
                    } else if (reportSubTab === 'BOT Report') {
                      exportRows = salesMetrics.topItems
                        .filter(i => i.department === 'Bar')
                        .map(i => ({
                          'Beverage': i.name,
                          'Category': i.category,
                          'Glasses / Units': i.sold,
                          'Total Revenue': i.revenue
                        }));
                    } else if (reportSubTab === 'Stock Usage') {
                      exportRows = inventory.map(ing => {
                        const used = salesMetrics.ingredientUsageMap[ing.id]?.totalConsumed || 0;
                        return {
                          'Raw Material': ing.name,
                          'Category': ing.category,
                          'Depleted Qty': used,
                          'Unit': ing.unit,
                          'Remaining Stock': ing.stock,
                          'Unit Cost': ing.cost,
                          'Total Depletion Cost': Number((used * ing.cost).toFixed(2))
                        };
                      });
                    } else if (reportSubTab === 'Stock Movement Ledger') {
                      exportRows = (stockLogs || [])
                        .filter(log => {
                          const lDate = log.date || extractDateStr(log.timestamp);
                          if (!lDate) return true;
                          if (reportStartDate && lDate < reportStartDate) return false;
                          if (reportEndDate && lDate > reportEndDate) return false;
                          return true;
                        })
                        .map(l => ({
                          'Date & Time': l.timestamp,
                          'Material': l.ingredientName,
                          'Action Type': l.type,
                          'Previous Stock': l.oldStock,
                          'Change': l.diffQty,
                          'New Level': l.newStock,
                          'Unit': l.unit,
                          'Cost Impact': l.totalCostImpact,
                          'Reason': l.reason,
                          'Staff': l.staff
                        }));
                    } else if (reportSubTab === 'Audit Trail') {
                      exportRows = auditLogs.map(l => ({
                        'Time': l.timestamp,
                        'Action': l.action,
                        'Authorized User': l.staff,
                        'Role': l.role,
                        'Target Ref': l.targetRef,
                        'Details': l.details
                      }));
                    } else {
                      // Fallback: Sales Summary export
                      exportRows = [
                        { 'Metric': 'Total Settled Invoices', 'Value': salesMetrics.paidBillsCount },
                        { 'Metric': 'Gross Revenue', 'Value': salesMetrics.grossRevenue },
                        { 'Metric': 'Net Food & Beverage Subtotal', 'Value': salesMetrics.itemSubtotal },
                        { 'Metric': 'Service Charge Pool', 'Value': salesMetrics.serviceCharge },
                        { 'Metric': 'Statutory Taxes / VAT', 'Value': salesMetrics.taxes },
                        { 'Metric': 'Discounts Given', 'Value': salesMetrics.discounts }
                      ];
                    }

                    exportReportToExcel(reportSubTab, exportRows, prefix);
                  }}
                  className="px-3 py-1 bg-emerald-600 hover:bg-emerald-700 text-white font-bold rounded-lg text-xs flex items-center gap-1.5 shadow-xs cursor-pointer transition-colors active:scale-95"
                  title="Export this report to Excel (.xlsx)"
                >
                  <Download className="h-3.5 w-3.5" />
                  <span>Excel</span>
                </button>

                {/* 2. PDF / PRINT EXPORT BUTTON */}
                <button
                  type="button"
                  onClick={() => {
                    window.print();
                  }}
                  className="px-3 py-1 bg-slate-900 hover:bg-slate-800 text-white font-bold rounded-lg text-xs flex items-center gap-1.5 shadow-xs cursor-pointer transition-colors active:scale-95"
                  title="Print or Save this report as PDF"
                >
                  <Printer className="h-3.5 w-3.5 text-orange-400" />
                  <span>PDF / Print</span>
                </button>
              </div>
            </div>

            {/* Daily Overview */}
            {reportSubTab === 'Daily Overview' && (
              <>
                <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
                  <div className="bg-white rounded-2xl border border-slate-200 p-5 shadow-xs">
                    <p className="text-[10px] font-black uppercase tracking-wider text-slate-400">GROSS REVENUE</p>
                    <p className="text-2xl font-black text-slate-900 mt-2 font-mono">
                      {settings.currency} {salesMetrics.grossRevenue.toLocaleString('en-US', { minimumFractionDigits: 2 })}
                    </p>
                    <p className="text-xs text-slate-500 mt-1 font-medium">{salesMetrics.paidBillsCount} Paid Bills</p>
                  </div>

                  <div className="bg-white rounded-2xl border border-slate-200 p-5 shadow-xs">
                    <p className="text-[10px] font-black uppercase tracking-wider text-slate-400">ITEM SUBTOTAL</p>
                    <p className="text-2xl font-black text-slate-900 mt-2 font-mono">
                      {settings.currency} {salesMetrics.itemSubtotal.toLocaleString('en-US', { minimumFractionDigits: 2 })}
                    </p>
                    <p className="text-xs text-slate-500 mt-1 font-medium">Food &amp; Beverage Sales</p>
                  </div>

                  <div className="bg-white rounded-2xl border border-slate-200 p-5 shadow-xs">
                    <p className="text-[10px] font-black uppercase tracking-wider text-slate-400">SERVICE CHARGE</p>
                    <p className="text-2xl font-black text-emerald-600 mt-2 font-mono">
                      {settings.currency} {salesMetrics.serviceCharge.toLocaleString('en-US', { minimumFractionDigits: 2 })}
                    </p>
                    <p className="text-xs text-slate-500 mt-1 font-medium">Collected for staff pool</p>
                  </div>

                  <div className="bg-white rounded-2xl border border-slate-200 p-5 shadow-xs">
                    <p className="text-[10px] font-black uppercase tracking-wider text-slate-400">TAXES / DISCOUNTS</p>
                    <p className="text-2xl font-black text-slate-900 mt-2 font-mono">
                      {settings.currency} {salesMetrics.taxes.toLocaleString('en-US', { minimumFractionDigits: 2 })}
                    </p>
                    <p className="text-xs text-slate-500 mt-1 font-medium">Discounts: {settings.currency} {salesMetrics.discounts.toFixed(2)}</p>
                  </div>
                </div>

                <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
                  <div className="bg-white rounded-2xl border border-slate-200 p-5 shadow-xs">
                    <h3 className="text-xs font-extrabold uppercase tracking-wider text-slate-900 mb-4">
                      Payment Methods Breakdown
                    </h3>
                    <div className="space-y-3">
                      {Object.keys(salesMetrics.paymentMethods).length === 0 ? (
                        <p className="text-xs text-slate-400 italic">No payments collected in selected period.</p>
                      ) : (
                        Object.entries(salesMetrics.paymentMethods).map(([method, data]) => (
                          <div key={method} className="flex items-center justify-between p-3 rounded-xl bg-slate-50 border border-slate-100">
                            <span className="text-xs font-bold text-slate-900">{method} ({data.count} bills)</span>
                            <span className="text-sm font-black font-mono text-slate-900">
                              {settings.currency} {data.total.toLocaleString('en-US', { minimumFractionDigits: 2 })}
                            </span>
                          </div>
                        ))
                      )}
                    </div>
                  </div>

                  <div className="bg-white rounded-2xl border border-slate-200 p-5 shadow-xs">
                    <h3 className="text-xs font-extrabold uppercase tracking-wider text-slate-900 mb-4">
                      Preparation Area Sales
                    </h3>
                    <div className="space-y-3">
                      <div className="flex items-center justify-between p-3 rounded-xl bg-slate-50 border border-slate-100">
                        <span className="text-xs font-bold text-slate-900">Kitchen ({salesMetrics.kitchenItemsCount} items)</span>
                        <span className="text-sm font-black font-mono text-slate-900">
                          {settings.currency} {salesMetrics.kitchenRevenue.toLocaleString('en-US', { minimumFractionDigits: 2 })}
                        </span>
                      </div>
                      <div className="flex items-center justify-between p-3 rounded-xl bg-slate-50 border border-slate-100">
                        <span className="text-xs font-bold text-slate-900">Bar ({salesMetrics.barItemsCount} drinks)</span>
                        <span className="text-sm font-black font-mono text-slate-900">
                          {settings.currency} {salesMetrics.barRevenue.toLocaleString('en-US', { minimumFractionDigits: 2 })}
                        </span>
                      </div>
                    </div>
                  </div>
                </div>

                {/* All Items & Top Selling Menu Items Table */}
                <div className="bg-white rounded-2xl border border-slate-200 p-5 shadow-xs">
                  <div className="flex items-center justify-between mb-4">
                    <h3 className="text-xs font-extrabold uppercase tracking-wider text-slate-900">
                      Top &amp; All Sold Menu Items
                    </h3>
                    <span className="text-xs font-mono text-slate-500">{salesMetrics.topItems.length} Products Sold</span>
                  </div>
                  <div className="overflow-x-auto">
                    <table className="w-full text-left text-xs">
                      <thead className="text-[10px] font-black uppercase text-slate-400 border-b border-slate-200">
                        <tr>
                          <th className="py-2.5">Menu Item</th>
                          <th className="py-2.5">Area</th>
                          <th className="py-2.5">Category</th>
                          <th className="py-2.5 text-center">Portions Sold</th>
                          <th className="py-2.5 text-right">Price</th>
                          <th className="py-2.5 text-right">Total Revenue</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-100">
                        {salesMetrics.topItems.length === 0 ? (
                          <tr>
                            <td colSpan={6} className="py-6 text-center text-slate-400 italic">No menu items sold in this period.</td>
                          </tr>
                        ) : (
                          salesMetrics.topItems.map((item, idx) => (
                            <tr key={idx} className="hover:bg-slate-50">
                              <td className="py-3 font-bold text-slate-900">{item.name}</td>
                              <td className="py-3">
                                <span className={`px-2 py-0.5 rounded text-[10px] font-bold ${
                                  item.department === 'Kitchen' ? 'bg-rose-100 text-rose-700' : 'bg-indigo-100 text-indigo-700'
                                }`}>
                                  {item.department}
                                </span>
                              </td>
                              <td className="py-3 text-slate-500">{item.category}</td>
                              <td className="py-3 text-center font-mono font-bold text-slate-800">{item.sold}</td>
                              <td className="py-3 text-right font-mono text-slate-600">{settings.currency} {(item.unitPrice || 0).toFixed(2)}</td>
                              <td className="py-3 text-right font-mono font-black text-slate-900">
                                {settings.currency} {item.revenue.toLocaleString('en-US', { minimumFractionDigits: 2 })}
                              </td>
                            </tr>
                          ))
                        )}
                      </tbody>
                    </table>
                  </div>
                </div>
              </>
            )}

            {/* Subtab: All Items Sales */}
            {reportSubTab === 'All Items Sales' && (
              <div className="bg-white rounded-2xl border border-slate-200 p-5 shadow-xs space-y-4">
                <div className="flex items-center justify-between">
                  <div>
                    <h3 className="text-sm font-black text-slate-900 uppercase">Itemized Menu Sales Report</h3>
                    <p className="text-xs text-slate-500">Every menu item ordered within the selected date filter range</p>
                  </div>
                  <span className="text-xs font-mono font-bold px-3 py-1 bg-slate-100 rounded-xl text-slate-700">
                    {salesMetrics.topItems.reduce((acc, i) => acc + i.sold, 0)} Total Portions
                  </span>
                </div>

                <div className="overflow-x-auto">
                  <table className="w-full text-left text-xs">
                    <thead className="text-[10px] font-black uppercase text-slate-400 border-b border-slate-200">
                      <tr>
                        <th className="py-2.5">Menu Item</th>
                        <th className="py-2.5">Department</th>
                        <th className="py-2.5">Category</th>
                        <th className="py-2.5 text-center">Portions Sold</th>
                        <th className="py-2.5 text-right">Selling Price</th>
                        <th className="py-2.5 text-right">Total Revenue</th>
                        <th className="py-2.5 text-right">% of Item Sales</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100">
                      {salesMetrics.topItems.length === 0 ? (
                        <tr>
                          <td colSpan={7} className="py-8 text-center text-slate-400 italic">No sales recorded for any items in this period.</td>
                        </tr>
                      ) : (
                        salesMetrics.topItems.map((item, idx) => {
                          const pct = salesMetrics.itemSubtotal > 0
                            ? ((item.revenue / salesMetrics.itemSubtotal) * 100).toFixed(1)
                            : '0.0';
                          return (
                            <tr key={idx} className="hover:bg-slate-50">
                              <td className="py-3 font-bold text-slate-900">{item.name}</td>
                              <td className="py-3">
                                <span className={`px-2 py-0.5 rounded text-[10px] font-bold ${
                                  item.department === 'Kitchen' ? 'bg-rose-100 text-rose-700' : 'bg-indigo-100 text-indigo-700'
                                }`}>
                                  {item.department}
                                </span>
                              </td>
                              <td className="py-3 text-slate-500">{item.category}</td>
                              <td className="py-3 text-center font-mono font-bold text-slate-800">{item.sold}</td>
                              <td className="py-3 text-right font-mono text-slate-600">{settings.currency} {(item.unitPrice || 0).toFixed(2)}</td>
                              <td className="py-3 text-right font-mono font-black text-slate-900">
                                {settings.currency} {item.revenue.toLocaleString('en-US', { minimumFractionDigits: 2 })}
                              </td>
                              <td className="py-3 text-right font-mono text-slate-500">{pct}%</td>
                            </tr>
                          );
                        })
                      )}
                    </tbody>
                  </table>
                </div>
              </div>
            )}
            
            {/* Subtab: Cash Out & Expense Disbursement Report */}
            {reportSubTab === 'Cash Out Report' && (
              <div className="bg-white rounded-2xl border border-slate-200 p-5 shadow-xs space-y-4">
                <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 border-b border-slate-100 pb-3">
                  <div>
                    <h3 className="text-sm font-black text-slate-900 uppercase">Cash Out &amp; Expense Disbursements Report</h3>
                    <p className="text-xs text-slate-500">All drawer cash disbursements, categorized expenses, and approval audit records</p>
                  </div>

                  {(() => {
                    const allRecords = Array.isArray(expenses) ? expenses : [];
                    const filtered = allRecords.filter(exp => {
                      const expDate = exp.date || extractDateStr(exp.createdAt);
                      if (!expDate) return true;
                      if (reportStartDate && expDate < reportStartDate) return false;
                      if (reportEndDate && expDate > reportEndDate) return false;
                      return true;
                    });
                    const totalDisbursed = filtered
                      .filter(p => p.status === 'APPROVED' || !p.status)
                      .reduce((sum, p) => sum + (parseFloat(p.amount) || 0), 0);

                    return (
                      <div className="flex items-center gap-2">
                        <span className="px-3 py-1 bg-rose-50 text-rose-700 font-mono font-bold text-xs rounded-xl border border-rose-200">
                          Total Approved: {settings.currency} {totalDisbursed.toFixed(2)}
                        </span>
                        <span className="px-3 py-1 bg-slate-100 font-mono font-bold text-xs rounded-xl text-slate-700">
                          {filtered.length} Requests
                        </span>
                      </div>
                    );
                  })()}
                </div>

                <div className="overflow-x-auto">
                  <table className="w-full text-left text-xs">
                    <thead className="bg-slate-50 text-[10px] font-black uppercase text-slate-400 border-b border-slate-200">
                      <tr>
                        <th className="py-2.5 px-3">Voucher Ref</th>
                        <th className="py-2.5 px-3">Date &amp; Time</th>
                        <th className="py-2.5 px-3">Shift ID</th>
                        <th className="py-2.5 px-3">Category</th>
                        <th className="py-2.5 px-3">Recipient / Paid To</th>
                        <th className="py-2.5 px-3">Description / Reason</th>
                        <th className="py-2.5 px-3 text-right">Amount</th>
                        <th className="py-2.5 px-3 text-center">Status</th>
                        <th className="py-2.5 px-3">Approved By</th>
                        <th className="py-2.5 px-3 text-right">Action</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100">
                      {(() => {
                        const allRecords = Array.isArray(expenses) ? expenses : [];
                        const filtered = allRecords.filter(exp => {
                          const expDate = exp.date || extractDateStr(exp.createdAt);
                          if (!expDate) return true;
                          if (reportStartDate && expDate < reportStartDate) return false;
                          if (reportEndDate && expDate > reportEndDate) return false;
                          return true;
                        });

                        if (filtered.length === 0) {
                          return (
                            <tr>
                              <td colSpan={10} className="py-8 text-center text-slate-400 italic">
                                No cash-out expense disbursements recorded in this date range.
                              </td>
                            </tr>
                          );
                        }

                        return filtered.map(item => {
                          const isApproved = item.status === 'APPROVED' || !item.status;
                          const isPending = item.status === 'PENDING';

                          return (
                            <tr key={item.id} className="hover:bg-slate-50">
                              <td className="py-3 px-3 font-mono font-bold text-slate-800">{item.id}</td>
                              <td className="py-3 px-3 text-slate-500 whitespace-nowrap">{item.date || getLocalDateStr()} {item.createdAt || item.time}</td>
                              <td className="py-3 px-3 font-mono text-[11px] text-slate-600">{item.shiftId || 'N/A'}</td>
                              <td className="py-3 px-3 font-medium text-slate-800">{item.category}</td>
                              <td className="py-3 px-3 text-slate-700 font-semibold">{item.recipient || 'General Expense'}</td>
                              <td className="py-3 px-3 text-slate-600 italic max-w-xs truncate" title={item.reason}>{item.reason}</td>
                              <td className="py-3 px-3 text-right font-mono font-black text-slate-900">
                                {settings.currency} {(parseFloat(item.amount) || 0).toFixed(2)}
                              </td>
                              <td className="py-3 px-3 text-center">
                                <span className={`px-2 py-0.5 rounded text-[10px] font-black uppercase ${
                                  isApproved
                                    ? 'bg-emerald-100 text-emerald-800'
                                    : isPending
                                    ? 'bg-amber-100 text-amber-800'
                                    : 'bg-rose-100 text-rose-800'
                                }`}>
                                  {item.status || 'APPROVED'}
                                </span>
                              </td>
                              <td className="py-3 px-3 text-slate-700">
                                {item.approvedBy ? `${item.approvedBy} (${item.approvedAt || 'Verified'})` : <span className="text-slate-400">Pending</span>}
                              </td>
                              <td className="py-3 px-3 text-right">
                                  <div className="flex items-center justify-end gap-1.5">
                                    {/* Reprint Voucher */}
                                    <button
                                      type="button"
                                      onClick={() => {
                                        triggerAutoPrint({
                                          type: 'CASH_OUT_VOUCHER',
                                          data: item
                                        }, `Reprint Cash Out Ref ${item.id}`);
                                      }}
                                      className="p-1.5 text-slate-400 hover:text-slate-900 rounded-lg hover:bg-slate-100 transition-colors cursor-pointer"
                                      title="Reprint Cash Out Voucher"
                                    >
                                      <Printer className="h-3.5 w-3.5" />
                                    </button>

                                    {/* Admin Delete Action */}
                                    {currentUser.role === 'Administrator' && (
                                      <button
                                        type="button"
                                        onClick={() => {
                                          if (window.confirm(`Delete cash out voucher ${item.id} (${settings.currency} ${item.amount})?`)) {
                                            // 1. Remove from expenses list
                                            setExpenses(prev => (Array.isArray(prev) ? prev.filter(e => e.id !== item.id) : []));

                                            // 2. Remove from active shift payouts if present
                                            setCurrentShift(prev => ({
                                              ...prev,
                                              payouts: Array.isArray(prev?.payouts) 
                                                ? prev.payouts.filter(p => p.id !== item.id) 
                                                : []
                                            }));

                                            // 3. Record Audit Log
                                            recordAuditLog(
                                              'CASH_OUT_DELETED',
                                              item.id,
                                              `Admin ${currentUser.name} deleted cash out voucher ${item.id} for ${settings.currency} ${item.amount}`
                                            );
                                          }
                                        }}
                                        className="p-1.5 text-slate-400 hover:text-rose-600 rounded-lg hover:bg-rose-50 transition-colors cursor-pointer"
                                        title="Delete Cash Out Record"
                                      >
                                        <Trash2 className="h-3.5 w-3.5" />
                                      </button>
                                    )}
                                  </div>
                                </td>
                            </tr>
                          );
                        });
                      })()}
                    </tbody>
                  </table>
                </div>
              </div>
            )}

            {/* Subtab: KOT Report */}
            {reportSubTab === 'KOT Report' && (
              <div className="bg-white rounded-2xl border border-slate-200 p-5 shadow-xs space-y-4">
                <div className="flex items-center justify-between">
                  <div>
                    <h3 className="text-sm font-black text-slate-900 uppercase">Kitchen Order Tickets (KOT) Production Report</h3>
                    <p className="text-xs text-slate-500">Breakdown of all kitchen items prepped in the filtered period</p>
                  </div>
                  <span className="px-3 py-1 bg-rose-100 text-rose-800 text-xs font-bold rounded-xl">
                    {salesMetrics.kitchenItemsCount} Kitchen Items • {settings.currency} {salesMetrics.kitchenRevenue.toFixed(2)}
                  </span>
                </div>
                <div className="overflow-x-auto">
                  <table className="w-full text-left text-xs">
                    <thead className="text-[10px] font-black uppercase text-slate-400 border-b border-slate-200">
                      <tr>
                        <th className="py-2.5">Kitchen Dish</th>
                        <th className="py-2.5">Category</th>
                        <th className="py-2.5 text-center">Portions Prepared</th>
                        <th className="py-2.5 text-right">Revenue</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100">
                      {salesMetrics.topItems.filter(i => i.department === 'Kitchen').length === 0 ? (
                        <tr>
                          <td colSpan={4} className="py-6 text-center text-slate-400 italic">No kitchen orders recorded in this date range.</td>
                        </tr>
                      ) : (
                        salesMetrics.topItems.filter(i => i.department === 'Kitchen').map((item, idx) => (
                          <tr key={idx} className="hover:bg-slate-50">
                            <td className="py-3 font-bold text-slate-900">{item.name}</td>
                            <td className="py-3 text-slate-500">{item.category}</td>
                            <td className="py-3 text-center font-mono font-bold text-slate-800">{item.sold}</td>
                            <td className="py-3 text-right font-mono font-bold text-slate-900">{settings.currency} {item.revenue.toFixed(2)}</td>
                          </tr>
                        ))
                      )}
                    </tbody>
                  </table>
                </div>
              </div>
            )}

            {/* Subtab: BOT Report */}
            {reportSubTab === 'BOT Report' && (
              <div className="bg-white rounded-2xl border border-slate-200 p-5 shadow-xs space-y-4">
                <div className="flex items-center justify-between">
                  <div>
                    <h3 className="text-sm font-black text-slate-900 uppercase">Bar Order Tickets (BOT) Dispense Report</h3>
                    <p className="text-xs text-slate-500">Breakdown of all bar beverages, cocktails &amp; coffees served</p>
                  </div>
                  <span className="px-3 py-1 bg-indigo-100 text-indigo-800 text-xs font-bold rounded-xl">
                    {salesMetrics.barItemsCount} Drinks Served • {settings.currency} {salesMetrics.barRevenue.toFixed(2)}
                  </span>
                </div>
                <div className="overflow-x-auto">
                  <table className="w-full text-left text-xs">
                    <thead className="text-[10px] font-black uppercase text-slate-400 border-b border-slate-200">
                      <tr>
                        <th className="py-2.5">Beverage / Drink</th>
                        <th className="py-2.5">Category</th>
                        <th className="py-2.5 text-center">Glasses / Units</th>
                        <th className="py-2.5 text-right">Revenue</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100">
                      {salesMetrics.topItems.filter(i => i.department === 'Bar').length === 0 ? (
                        <tr>
                          <td colSpan={4} className="py-6 text-center text-slate-400 italic">No bar beverage orders recorded in this date range.</td>
                        </tr>
                      ) : (
                        salesMetrics.topItems.filter(i => i.department === 'Bar').map((item, idx) => (
                          <tr key={idx} className="hover:bg-slate-50">
                            <td className="py-3 font-bold text-slate-900">{item.name}</td>
                            <td className="py-3 text-slate-500">{item.category}</td>
                            <td className="py-3 text-center font-mono font-bold text-slate-800">{item.sold}</td>
                            <td className="py-3 text-right font-mono font-bold text-slate-900">{settings.currency} {item.revenue.toFixed(2)}</td>
                          </tr>
                        ))
                      )}
                    </tbody>
                  </table>
                </div>
              </div>
            )}

            {/* Subtab: Sales Summary */}
            {reportSubTab === 'Sales Summary' && (
              <div className="bg-white rounded-2xl border border-slate-200 p-6 shadow-xs space-y-4 max-w-2xl">
                <h3 className="text-base font-black text-slate-900 uppercase">Executive Financial Summary</h3>
                <div className="space-y-2.5 text-xs divide-y divide-slate-100">
                  <div className="flex justify-between py-2">
                    <span className="text-slate-600">Total Settled Invoices</span>
                    <span className="font-mono font-bold text-slate-900">{salesMetrics.paidBillsCount}</span>
                  </div>
                  <div className="flex justify-between py-2">
                    <span className="text-slate-600">Average Order Value (AOV)</span>
                    <span className="font-mono font-bold text-slate-900">
                      {settings.currency} {salesMetrics.paidBillsCount > 0 ? (salesMetrics.grossRevenue / salesMetrics.paidBillsCount).toFixed(2) : '0.00'}
                    </span>
                  </div>
                  <div className="flex justify-between py-2">
                    <span className="text-slate-600">Net Food &amp; Beverage Subtotal</span>
                    <span className="font-mono font-bold text-slate-900">{settings.currency} {salesMetrics.itemSubtotal.toFixed(2)}</span>
                  </div>
                  <div className="flex justify-between py-2">
                    <span className="text-slate-600">Service Charge Pool ({settings.serviceChargeRate}%)</span>
                    <span className="font-mono font-bold text-emerald-600">+{settings.currency} {salesMetrics.serviceCharge.toFixed(2)}</span>
                  </div>
                  <div className="flex justify-between py-2">
                    <span className="text-slate-600">Statutory Taxes / VAT ({settings.taxRate}%)</span>
                    <span className="font-mono font-bold text-slate-900">+{settings.currency} {salesMetrics.taxes.toFixed(2)}</span>
                  </div>
                  <div className="flex justify-between py-2">
                    <span className="text-slate-600">Total Discounts Deducted</span>
                    <span className="font-mono font-bold text-rose-600">-{settings.currency} {salesMetrics.discounts.toFixed(2)}</span>
                  </div>
                  <div className="flex justify-between py-3 text-sm font-black text-slate-900 border-t-2 border-slate-900">
                    <span>Total Gross Revenue</span>
                    <span className="font-mono text-base text-[#ff5500]">{settings.currency} {salesMetrics.grossRevenue.toFixed(2)}</span>
                  </div>
                </div>
              </div>
            )}

            {/* Subtab: Food vs Beverage */}
            {reportSubTab === 'Food vs Beverage' && (
              <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                <div className="bg-white rounded-2xl border border-slate-200 p-6 shadow-xs flex flex-col justify-between">
                  <div>
                    <span className="text-xs font-bold text-rose-600 uppercase tracking-wider">Food / Kitchen (KOT)</span>
                    <h3 className="text-3xl font-black font-mono text-slate-900 mt-2">
                      {settings.currency} {salesMetrics.kitchenRevenue.toFixed(2)}
                    </h3>
                    <p className="text-xs text-slate-500 mt-1">
                      {((salesMetrics.kitchenRevenue / (salesMetrics.itemSubtotal || 1)) * 100).toFixed(1)}% of total menu sales
                    </p>
                  </div>
                  <div className="mt-6 pt-4 border-t border-slate-100">
                    <span className="text-xs font-bold text-slate-700">Total Kitchen Portions: {salesMetrics.kitchenItemsCount}</span>
                  </div>
                </div>

                <div className="bg-white rounded-2xl border border-slate-200 p-6 shadow-xs flex flex-col justify-between">
                  <div>
                    <span className="text-xs font-bold text-indigo-600 uppercase tracking-wider">Beverages / Bar (BOT)</span>
                    <h3 className="text-3xl font-black font-mono text-slate-900 mt-2">
                      {settings.currency} {salesMetrics.barRevenue.toFixed(2)}
                    </h3>
                    <p className="text-xs text-slate-500 mt-1">
                      {((salesMetrics.barRevenue / (salesMetrics.itemSubtotal || 1)) * 100).toFixed(1)}% of total menu sales
                    </p>
                  </div>
                  <div className="mt-6 pt-4 border-t border-slate-100">
                    <span className="text-xs font-bold text-slate-700">Total Bar Drinks: {salesMetrics.barItemsCount}</span>
                  </div>
                </div>
              </div>
            )}

            {/* Subtab: Stock Usage */}
            {reportSubTab === 'Stock Usage' && (
              <div className="bg-white rounded-2xl border border-slate-200 p-5 shadow-xs space-y-4">
                <div className="flex items-center justify-between">
                  <div>
                    <h3 className="text-sm font-black text-slate-900 uppercase">Raw Ingredient Depletion &amp; Usage</h3>
                    <p className="text-xs text-slate-500">Calculated through dish recipes for all sold orders in this period</p>
                  </div>
                </div>
                <div className="overflow-x-auto">
                  <table className="w-full text-left text-xs">
                    <thead className="text-[10px] font-black uppercase text-slate-400 border-b border-slate-200">
                      <tr>
                        <th className="py-2.5">Raw Material</th>
                        <th className="py-2.5">Category</th>
                        <th className="py-2.5 text-center">Depleted (Used)</th>
                        <th className="py-2.5 text-center">Remaining In Stock</th>
                        <th className="py-2.5 text-right">Unit Cost</th>
                        <th className="py-2.5 text-right">Total Depletion Cost</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100">
                      {inventory.map(ing => {
                        const used = salesMetrics.ingredientUsageMap[ing.id]?.totalConsumed || 0;
                        const cost = used * ing.cost;
                        return (
                          <tr key={ing.id} className="hover:bg-slate-50">
                            <td className="py-3 font-bold text-slate-900">{ing.name}</td>
                            <td className="py-3 text-slate-500">{ing.category}</td>
                            <td className="py-3 text-center font-mono font-bold text-slate-800">
                              {used > 0 ? `${used.toFixed(1)} ${ing.unit}` : '-'}
                            </td>
                            <td className="py-3 text-center font-mono text-slate-600">{ing.stock} {ing.unit}</td>
                            <td className="py-3 text-right font-mono text-slate-500">{settings.currency} {ing.cost.toFixed(2)}</td>
                            <td className="py-3 text-right font-mono font-black text-slate-900">
                              {cost > 0 ? `${settings.currency} ${cost.toFixed(2)}` : '-'}
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              </div>
            )}

            {/* Subtab: Stock Movement Ledger */}
            {reportSubTab === 'Stock Movement Ledger' && (
              <div className="bg-white rounded-2xl border border-slate-200 p-5 shadow-xs space-y-4">
                <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                  <div>
                    <h3 className="text-sm font-black text-slate-900 uppercase">Stock Intake &amp; Variance Change Ledger</h3>
                    <p className="text-xs text-slate-500">History of every intake (+), sales deduction (-), or manual edit</p>
                  </div>
                  <span className="text-xs font-mono font-bold px-3 py-1 bg-slate-100 rounded-xl text-slate-700">
                    {stockLogs.length} Records
                  </span>
                </div>

                <div className="overflow-x-auto">
                  <table className="w-full text-left text-xs">
                    <thead className="bg-slate-50 text-[10px] font-black uppercase text-slate-400 border-b border-slate-200">
                      <tr>
                        <th className="py-2.5 px-3">Date &amp; Time</th>
                        <th className="py-2.5 px-3">Material</th>
                        <th className="py-2.5 px-3">Action Type</th>
                        <th className="py-2.5 px-3 text-right">Previous</th>
                        <th className="py-2.5 px-3 text-center">Change (Diff)</th>
                        <th className="py-2.5 px-3 text-right">New Level</th>
                        <th className="py-2.5 px-3 text-right">Cost Impact</th>
                        <th className="py-2.5 px-3">Reason / Ref</th>
                        <th className="py-2.5 px-3">Staff</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100">
                      {(() => {
                        const filteredLogs = (stockLogs || []).filter(log => {
                          const lDate = log.date || extractDateStr(log.timestamp);
                          if (!lDate) return true;
                          if (reportStartDate && lDate < reportStartDate) return false;
                          if (reportEndDate && lDate > reportEndDate) return false;
                          return true;
                        });

                        if (filteredLogs.length === 0) {
                          return (
                            <tr>
                              <td colSpan={9} className="py-8 text-center text-slate-400 italic">
                                No stock movements or intake events recorded in this period.
                              </td>
                            </tr>
                          );
                        }

                        return filteredLogs.map(log => {
                          const isPositive = log.diffQty > 0;
                          return (
                            <tr key={log.id} className="hover:bg-slate-50/70">
                              <td className="py-3 px-3 font-mono text-slate-500 whitespace-nowrap">{log.timestamp}</td>
                              <td className="py-3 px-3 font-bold text-slate-900">{log.ingredientName}</td>
                              <td className="py-3 px-3">
                                <span className={`px-2 py-0.5 rounded text-[10px] font-bold ${
                                  log.type === 'INTAKE'
                                    ? 'bg-emerald-100 text-emerald-800'
                                    : log.type === 'SALE_DEPLETION'
                                    ? 'bg-blue-100 text-blue-800'
                                    : log.type === 'MANUAL_EDIT'
                                    ? 'bg-amber-100 text-amber-800'
                                    : 'bg-slate-100 text-slate-800'
                                }`}>
                                  {log.type}
                                </span>
                              </td>
                              <td className="py-3 px-3 text-right font-mono text-slate-600">
                                {log.oldStock} {log.unit}
                              </td>
                              <td className="py-3 px-3 text-center font-mono font-black">
                                <span className={`px-2 py-0.5 rounded ${
                                  isPositive ? 'bg-emerald-50 text-emerald-700' : 'bg-rose-50 text-rose-700'
                                }`}>
                                  {isPositive ? `+${log.diffQty}` : log.diffQty} {log.unit}
                                </span>
                              </td>
                              <td className="py-3 px-3 text-right font-mono font-bold text-slate-900">
                                {log.newStock} {log.unit}
                              </td>
                              <td className="py-3 px-3 text-right font-mono font-bold text-slate-700">
                                {settings.currency} {log.totalCostImpact.toFixed(2)}
                              </td>
                              <td className="py-3 px-3 text-slate-600 truncate max-w-[200px]" title={log.reason}>
                                {log.reason}
                              </td>
                              <td className="py-3 px-3 text-slate-700 font-medium">{log.staff}</td>
                            </tr>
                          );
                        });
                      })()}
                    </tbody>
                  </table>
                </div>
              </div>
            )}

            {/* Subtab: Sales Detail */}
            {reportSubTab === 'Sales Detail' && (
              <div className="bg-white rounded-2xl border border-slate-200 p-5 shadow-xs">
                <div className="flex items-center justify-between mb-4">
                  <h3 className="text-sm font-black text-slate-900 uppercase">Paid Invoices Ledger</h3>
                  <span className="text-xs font-mono text-slate-500">{filteredTransactions.length} Filtered Records</span>
                </div>
                <div className="overflow-x-auto">
                  <table className="w-full text-left text-xs">
                    <thead className="text-[10px] font-black uppercase text-slate-400 border-b border-slate-200">
                      <tr>
                        <th className="py-2.5">Invoice #</th>
                        <th className="py-2.5">Date &amp; Time</th>
                        <th className="py-2.5">Table</th>
                        <th className="py-2.5">Cashier</th>
                        <th className="py-2.5">Method</th>
                        <th className="py-2.5 text-right">Subtotal</th>
                        <th className="py-2.5 text-right">Grand Total</th>
                        {currentUser.role === 'Administrator' && (
                          <th className="py-2.5 text-right">Admin Action</th>
                        )}
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100">
                      {filteredTransactions.length === 0 ? (
                        <tr>
                          <td colSpan={currentUser.role === 'Administrator' ? 8 : 7} className="py-4 text-center text-slate-400 italic">No transactions recorded.</td>
                        </tr>
                      ) : (
                        filteredTransactions.map(t => (
                          <tr key={t.invoiceNo} className="hover:bg-slate-50">
                            <td className="py-3 font-mono font-bold text-slate-800">{t.invoiceNo}</td>
                            <td className="py-3 text-slate-500">{t.date}</td>
                            <td className="py-3 font-semibold text-slate-900">{t.table}</td>
                            <td className="py-3 text-slate-600">{t.cashier}</td>
                            <td className="py-3">
                              <span className="px-2 py-0.5 bg-slate-100 rounded text-[10px] font-bold text-slate-700">
                                {t.paymentMethod}
                              </span>
                            </td>
                            <td className="py-3 text-right font-mono">{settings.currency} {t.subtotal.toFixed(2)}</td>
                            <td className="py-3 text-right font-mono font-black text-slate-900">{settings.currency} {t.total.toFixed(2)}</td>
                            {currentUser.role === 'Administrator' && (
                              <td className="py-3 text-right">
                                <button
                                  type="button"
                                  onClick={() => {
                                    setTransactions(prev => prev.filter(inv => inv.invoiceNo !== t.invoiceNo));
                                    recordAuditLog('ADMIN_DELETE_TRANSACTION', t.invoiceNo, `Admin deleted invoice ${t.invoiceNo} for ${settings.currency} ${t.total.toFixed(2)}`);
                                  }}
                                  className="text-slate-400 hover:text-rose-600 p-1 cursor-pointer"
                                  title="Admin Delete Transaction"
                                >
                                  <Trash2 className="h-3.5 w-3.5" />
                                </button>
                              </td>
                            )}
                          </tr>
                        ))
                      )}
                    </tbody>
                  </table>
                </div>
              </div>
            )}

            {/* Subtab: Audit Trail */}
            {reportSubTab === 'Audit Trail' && (
              <div className="bg-white rounded-2xl border border-slate-200 p-5 shadow-xs">
                <div className="flex items-center justify-between mb-3">
                  <h3 className="text-sm font-black text-slate-900 uppercase">System Activity &amp; Change Audit Log</h3>
                  {currentUser.role === 'Administrator' && auditLogs.length > 0 && (
                    <button
                      type="button"
                      onClick={() => {
                        setAuditLogs([]);
                        recordAuditLog('ADMIN_CLEAR_AUDIT_LOGS', 'ALL', 'Purged all audit log entries');
                      }}
                      className="px-2.5 py-1 bg-rose-50 hover:bg-rose-100 text-rose-600 border border-rose-200 rounded-lg text-xs font-bold flex items-center gap-1 transition-colors cursor-pointer"
                    >
                      <Trash2 className="h-3 w-3" /> Clear Audit Logs (Admin)
                    </button>
                  )}
                </div>
                <div className="overflow-x-auto">
                  <table className="w-full text-left text-xs">
                    <thead className="text-[10px] font-black uppercase text-slate-400 border-b border-slate-200">
                      <tr>
                        <th className="py-2.5">Time</th>
                        <th className="py-2.5">Action</th>
                        <th className="py-2.5">Authorized User</th>
                        <th className="py-2.5">Role</th>
                        <th className="py-2.5">Target / Reference</th>
                        <th className="py-2.5">Details</th>
                        {currentUser.role === 'Administrator' && (
                          <th className="py-2.5 text-right">Admin Action</th>
                        )}
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100">
                      {auditLogs.length === 0 ? (
                        <tr>
                          <td colSpan={currentUser.role === 'Administrator' ? 7 : 6} className="py-4 text-center text-slate-400 italic">No activity logs recorded yet.</td>
                        </tr>
                      ) : (
                        auditLogs.map(log => (
                          <tr key={log.id} className="hover:bg-slate-50">
                            <td className="py-2.5 font-mono text-slate-500">{log.timestamp}</td>
                            <td className="py-2.5 font-bold text-slate-900 font-mono">{log.action}</td>
                            <td className="py-2.5 font-semibold text-slate-800">{log.staff}</td>
                            <td className="py-2.5 text-[10px] text-slate-500">{log.role}</td>
                            <td className="py-2.5 font-mono text-slate-600">{log.targetRef}</td>
                            <td className="py-2.5 text-slate-600">{log.details}</td>
                            {currentUser.role === 'Administrator' && (
                              <td className="py-2.5 text-right">
                                <button
                                  type="button"
                                  onClick={() => {
                                    setAuditLogs(prev => prev.filter(l => l.id !== log.id));
                                  }}
                                  className="text-slate-400 hover:text-rose-600 p-1 cursor-pointer"
                                  title="Admin Delete Log Entry"
                                >
                                  <Trash2 className="h-3.5 w-3.5" />
                                </button>
                              </td>
                            )}
                          </tr>
                        ))
                      )}
                    </tbody>
                  </table>
                </div>
              </div>
            )}
          </div>
        )}

        {/* VIEW 4: RECIPES & PORTIONS */}
        {activeTab === 'recipes' && (
          <div className="flex-1 overflow-y-auto p-6 space-y-6">
            <div className="flex items-center justify-between">
              <div>
                <h2 className="text-xl font-black text-slate-900">Recipes &amp; BOM Portion Costing</h2>
                <p className="text-xs text-slate-500 mt-0.5">
                  Cost of Goods Sold (COGS), profit margins, and remaining portions linked to raw inventory.
                </p>
              </div>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
              {menuItems.map(dish => {
                const { cogs, portions, isSoldOut } = calculateDishAvailability(dish.recipe);
                const margin = dish.price > 0 ? (((dish.price - cogs) / dish.price) * 100).toFixed(1) : 0;
                return (
                  <div key={dish.id} className="bg-white rounded-2xl border border-slate-200 p-5 shadow-xs space-y-4">
                    <div className="flex items-start justify-between">
                      <div>
                        <span className="text-[10px] font-bold text-slate-400 uppercase">{dish.department} • {dish.category}</span>
                        <h3 className="text-base font-black text-slate-900 mt-0.5">{dish.name}</h3>
                        <p className="text-xs text-slate-500 mt-0.5">{dish.description}</p>
                      </div>
                      <div className="text-right">
                        <span className="text-base font-black font-mono text-[#ff5500]">
                          {settings.currency} {dish.price.toFixed(2)}
                        </span>
                        <p className="text-[10px] font-bold text-emerald-600">{margin}% Gross Margin</p>
                      </div>
                    </div>

                    <div className="p-3 bg-slate-50 rounded-xl border border-slate-100 space-y-2">
                      <div className="flex justify-between items-center">
                        <p className="text-[10px] font-black uppercase text-slate-400">Bill of Materials (BOM)</p>
                        <button
                          onClick={() => {
                            setEditingDishForRecipe(dish);
                            setCurrentRecipeIngredients(dish.recipe ? [...dish.recipe] : []);
                            setTempIngredientSelect({ ingredientId: inventory[0]?.id || '', amount: '' });
                            setRecipeConfigModalOpen(true);
                          }}
                          className="text-[11px] font-bold text-[#ff5500] hover:underline flex items-center gap-1 cursor-pointer"
                        >
                          <Sliders className="h-3 w-3" /> Configure Recipe
                        </button>
                      </div>
                      {dish.recipe && dish.recipe.length > 0 ? (
                        dish.recipe.map((r, i) => {
                          const ing = inventoryMap[r.ingredientId];
                          const lineCost = ing ? (ing.cost * r.amount) : 0;
                          return (
                            <div key={i} className="flex justify-between text-xs">
                              <span className="text-slate-700">{ing ? ing.name : r.ingredientId} ({r.amount} {ing ? ing.unit : ''})</span>
                              <span className="font-mono text-slate-500">{settings.currency} {lineCost.toFixed(2)}</span>
                            </div>
                          );
                        })
                      ) : (
                        <p className="text-xs text-slate-400 italic">No raw inventory linked to this item.</p>
                      )}
                    </div>

                    <div className="flex justify-between items-center text-xs pt-2 border-t border-slate-100">
                      <span className="text-slate-600 font-medium">
                        Live Portions: <strong className={isSoldOut ? 'text-rose-600 font-mono' : 'text-slate-900 font-mono'}>{portions}</strong>
                      </span>
                      <span className="font-mono font-bold text-slate-700">
                        Total Raw Cost: {settings.currency} {cogs.toFixed(2)}
                      </span>
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        )}

        {/* VIEW 5: STOCK & RAW INVENTORY */}
        {activeTab === 'stock' && (
          <div className="flex-1 overflow-y-auto p-6 space-y-6">
            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4">
              <div>
                <h2 className="text-xl font-black text-slate-900">Stock &amp; Raw Inventory Valuation</h2>
                <p className="text-xs text-slate-500 mt-0.5">
                  Track total raw material investment, asset value, unit costs, and real-time recipe depletion.
                </p>
              </div>

              <div className="flex flex-wrap items-center gap-2 shrink-0">
                {/* 1. EXCEL EXPORT BUTTON */}
                <button
                  type="button"
                  onClick={() => {
                    const list = Array.isArray(inventory) ? inventory : [];
                    if (list.length === 0) {
                      alert('No raw inventory records available to export.');
                      return;
                    }

                    const exportRows = list.map(item => {
                      const stockVal = Number(item.stock) || 0;
                      const costVal = Number(item.cost) || 0;
                      const threshVal = Number(item.threshold) || 0;
                      const totalVal = Number((stockVal * costVal).toFixed(2));

                      return {
                        'Ingredient ID': item.id,
                        'Raw Ingredient': item.name,
                        'Category': item.category || 'General',
                        'Remaining Stock': stockVal,
                        'Unit': item.unit || 'g',
                        'Reorder Threshold': threshVal,
                        'Unit Cost': costVal,
                        'Total Valuation': totalVal,
                        'Stock Status': stockVal <= threshVal ? 'LOW STOCK' : 'OPTIMAL'
                      };
                    });

                    exportReportToExcel('Stock_Inventory_Valuation', exportRows, 'Stock_Inventory_Valuation');
                  }}
                  className="px-3.5 py-2 bg-emerald-600 hover:bg-emerald-700 text-white font-bold rounded-xl text-xs flex items-center gap-1.5 shadow-xs cursor-pointer transition-all active:scale-95"
                  title="Export raw inventory list and valuation to Excel (.xlsx)"
                >
                  <Download className="h-3.5 w-3.5" />
                  <span>Export Excel</span>
                </button>

                {/* 2. IMPORT EXCEL / PDF INVENTORY FILE */}
                <label className="px-3.5 py-2 bg-slate-800 hover:bg-slate-700 text-white font-bold rounded-xl text-xs flex items-center gap-1.5 shadow-xs shrink-0 cursor-pointer transition-colors active:scale-95">
                  <Upload className="h-3.5 w-3.5 text-[#ff5500]" />
                  <span>Import Excel / PDF</span>
                  <input
                    type="file"
                    accept=".xlsx, .xls, .csv, application/pdf"
                    className="hidden"
                    onChange={async (e) => {
                      const file = e.target.files?.[0];
                      if (!file) return;

                      try {
                        let imported = [];
                        const fileName = file.name.toLowerCase();

                        if (fileName.endsWith('.pdf')) {
                          if (typeof extractInventoryFromPDF === 'function') {
                            imported = await extractInventoryFromPDF(file);
                          } else {
                            alert('PDF inventory extractor helper is missing.');
                            return;
                          }
                        } else if (fileName.endsWith('.xlsx') || fileName.endsWith('.xls') || fileName.endsWith('.csv')) {
                          if (typeof extractInventoryFromExcel === 'function') {
                            imported = await extractInventoryFromExcel(file);
                          } else {
                            alert('Excel inventory extractor helper is missing.');
                            return;
                          }
                        }

                        if (!imported || imported.length === 0) {
                          alert(`Could not extract raw material records from "${file.name}". Please ensure the file has columns: Raw Ingredient, Category, Remaining Stock, Reorder Threshold, and Unit Cost.`);
                          return;
                        }

                        setInventory(prev => {
                          const existingList = Array.isArray(prev) ? [...prev] : [];
                          imported.forEach(newIng => {
                            const matchIdx = existingList.findIndex(
                              ex => ex.name.toLowerCase().trim() === newIng.name.toLowerCase().trim()
                            );
                            if (matchIdx >= 0) {
                              existingList[matchIdx] = {
                                ...existingList[matchIdx],
                                stock: Number((existingList[matchIdx].stock + newIng.stock).toFixed(2)),
                                cost: newIng.cost > 0 ? newIng.cost : existingList[matchIdx].cost,
                                threshold: newIng.threshold || existingList[matchIdx].threshold
                              };
                            } else {
                              existingList.push(newIng);
                            }
                          });
                          return existingList;
                        });

                        recordAuditLog(
                          'INVENTORY_IMPORTED',
                          file.name,
                          `Imported/updated ${imported.length} raw material records from ${file.name}`
                        );

                        alert(`Successfully imported ${imported.length} raw inventory materials from ${file.name}! Inventory valuation has updated.`);
                      } catch (err) {
                        console.error('Inventory import failed:', err);
                        alert(`Failed to import file: ${err.message}`);
                      } finally {
                        e.target.value = '';
                      }
                    }}
                  />
                </label>

                {/* 3. RECEIVE STOCK */}
                <button
                  type="button"
                  onClick={() => {
                    const firstId = Array.isArray(inventory) && inventory.length > 0 ? inventory[0].id : '';
                    setReceiveStockForm({
                      ingredientId: firstId,
                      quantity: '',
                      supplier: '',
                      invoiceRef: '',
                      newCost: ''
                    });
                    setReceiveStockModalOpen(true);
                  }}
                  className="px-3.5 py-2 bg-slate-900 hover:bg-slate-800 text-white font-bold rounded-xl text-xs flex items-center gap-1.5 shadow-xs cursor-pointer active:scale-95 transition-all"
                >
                  <Package className="h-3.5 w-3.5 text-emerald-400" />
                  <span>Receive Stock</span>
                </button>

                {/* 4. ADD MATERIAL */}
                <button
                  type="button"
                  onClick={() => {
                    setNewInventoryForm({
                      name: '',
                      category: 'Dry Goods',
                      stock: '',
                      unit: 'g',
                      cost: '',
                      threshold: '10'
                    });
                    setAddInventoryModalOpen(true);
                  }}
                  className="px-3.5 py-2 bg-[#ff5500] hover:bg-orange-600 text-white font-bold rounded-xl text-xs flex items-center gap-1.5 shadow-xs cursor-pointer active:scale-95 transition-all"
                >
                  <Plus className="h-3.5 w-3.5" />
                  <span>Add Material</span>
                </button>
              </div>
            </div>

            {/* INVENTORY VALUATION KPI CARDS */}
            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
              <div className="bg-white rounded-2xl border border-slate-200 p-5 shadow-xs">
                <div className="flex items-center justify-between">
                  <p className="text-[10px] font-black uppercase tracking-wider text-slate-400">Total Stock Value</p>
                  <span className="p-1.5 rounded-lg bg-orange-50 text-[#ff5500]">
                    <DollarSign className="h-4 w-4" />
                  </span>
                </div>
                <p className="text-2xl font-black text-slate-900 mt-2 font-mono">
                  {settings.currency} {inventoryValuation.totalStockValue.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                </p>
                <p className="text-[11px] text-slate-500 mt-1 font-medium">Total capital tied in current inventory</p>
              </div>

              <div className="bg-white rounded-2xl border border-slate-200 p-5 shadow-xs">
                <div className="flex items-center justify-between">
                  <p className="text-[10px] font-black uppercase tracking-wider text-slate-400">Tracked Ingredients</p>
                  <span className="p-1.5 rounded-lg bg-slate-100 text-slate-700">
                    <Package className="h-4 w-4" />
                  </span>
                </div>
                <p className="text-2xl font-black text-slate-900 mt-2 font-mono">
                  {inventoryValuation.totalItems} Items
                </p>
                <p className="text-[11px] text-slate-500 mt-1 font-medium">Active Bill of Materials stock records</p>
              </div>

              <div className="bg-white rounded-2xl border border-slate-200 p-5 shadow-xs">
                <div className="flex items-center justify-between">
                  <p className="text-[10px] font-black uppercase tracking-wider text-slate-400">Reorder Alerts</p>
                  <span className={`p-1.5 rounded-lg ${inventoryValuation.lowStockCount > 0 ? 'bg-rose-50 text-rose-600' : 'bg-emerald-50 text-emerald-600'}`}>
                    <AlertTriangle className="h-4 w-4" />
                  </span>
                </div>
                <p className={`text-2xl font-black mt-2 font-mono ${inventoryValuation.lowStockCount > 0 ? 'text-rose-600' : 'text-emerald-600'}`}>
                  {inventoryValuation.lowStockCount} Items Low
                </p>
                <p className="text-[11px] text-slate-500 mt-1 font-medium">Below configured threshold limit</p>
              </div>

              <div className="bg-white rounded-2xl border border-slate-200 p-5 shadow-xs">
                <div className="flex items-center justify-between">
                  <p className="text-[10px] font-black uppercase tracking-wider text-slate-400">Top Category Asset</p>
                  <span className="p-1.5 rounded-lg bg-indigo-50 text-indigo-600">
                    <BarChart3 className="h-4 w-4" />
                  </span>
                </div>
                <p className="text-lg font-black text-slate-900 mt-2 truncate">
                  {inventoryValuation.topCat}
                </p>
                <p className="text-[11px] font-mono font-bold text-indigo-600 mt-1">
                  {settings.currency} {inventoryValuation.maxCatVal.toLocaleString('en-US', { minimumFractionDigits: 2 })}
                </p>
              </div>
            </div>

            {/* INVENTORY TABLE WITH UNIT COST & TOTAL ASSET VALUE */}
            <div className="bg-white rounded-2xl border border-slate-200 overflow-hidden shadow-xs">
              <table className="w-full text-left text-xs">
                <thead className="bg-slate-50 text-[10px] font-black uppercase text-slate-400 border-b border-slate-200">
                  <tr>
                    <th className="py-3 px-4">Raw Ingredient</th>
                    <th className="py-3 px-4">Category</th>
                    <th className="py-3 px-4">Remaining Stock</th>
                    <th className="py-3 px-4">Reorder Threshold</th>
                    <th className="py-3 px-4 text-right">Unit Cost</th>
                    <th className="py-3 px-4 text-right">Total Valuation</th>
                    <th className="py-3 px-4 text-right">Actions</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {(!Array.isArray(inventory) || inventory.length === 0) ? (
                    <tr>
                      <td colSpan={7} className="py-8 text-center text-slate-400 italic">
                        No raw inventory items registered yet. Click &ldquo;Add Material&rdquo; above to record ingredients.
                      </td>
                    </tr>
                  ) : (
                    inventory.map(ing => {
                      if (!ing) return null;
                      const stockNum = Number(ing.stock) || 0;
                      const threshNum = Number(ing.threshold) || 0;
                      const costNum = Number(ing.cost) || 0;
                      const totalAssetVal = stockNum * costNum;
                      const isLow = stockNum <= threshNum;

                      return (
                        <tr key={ing.id} className="hover:bg-slate-50/70">
                          <td className="py-3 px-4">
                            <p className="font-extrabold text-slate-900">{ing.name}</p>
                            <span className="text-[10px] text-slate-400 font-mono">{ing.id}</span>
                          </td>
                          <td className="py-3 px-4">
                            <span className="px-2 py-0.5 bg-slate-100 text-slate-700 rounded text-[10px] font-bold">
                              {ing.category || 'General'}
                            </span>
                          </td>
                          <td className="py-3 px-4 font-mono font-bold text-slate-800">
                            {stockNum} {ing.unit}
                            {isLow && (
                              <span className="ml-2 px-1.5 py-0.5 bg-rose-100 text-rose-700 text-[10px] rounded font-bold">
                                Low
                              </span>
                            )}
                          </td>
                          <td className="py-3 px-4 font-mono text-slate-400">{threshNum} {ing.unit}</td>
                          <td className="py-3 px-4 text-right font-mono text-slate-600">
                            {settings.currency} {costNum.toFixed(2)} / {ing.unit}
                          </td>
                          <td className="py-3 px-4 text-right font-mono font-black text-slate-900">
                            {settings.currency} {totalAssetVal.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                          </td>
                          <td className="py-3 px-4 text-right">
                            <div className="flex items-center justify-end gap-1.5">
                              {/* Quick Restock / Intake */}
                              <button
                                type="button"
                                onClick={() => {
                                  setReceiveStockForm({
                                    ingredientId: ing.id,
                                    quantity: ing.unit === 'g' || ing.unit === 'ml' ? '1000' : '10',
                                    supplier: 'Local Market',
                                    invoiceRef: `REC-${Math.floor(100 + Math.random() * 900)}`,
                                    newCost: costNum.toString()
                                  });
                                  setReceiveStockModalOpen(true);
                                }}
                                className="px-2.5 py-1 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-lg text-xs font-bold transition-colors cursor-pointer"
                                title="Intake / Receive Stock"
                              >
                                + Intake
                              </button>

                              {/* Edit Material */}
                              <button
                                type="button"
                                onClick={() => {
                                  setEditingInventoryItem({ ...ing });
                                  setEditInventoryModalOpen(true);
                                }}
                                className="p-1.5 text-slate-400 hover:text-slate-800 hover:bg-slate-100 rounded-lg transition-colors cursor-pointer"
                                title="Edit Material"
                              >
                                <Edit3 className="h-4 w-4" />
                              </button>

                              {/* Remove Material */}
                              <button
                                type="button"
                                onClick={() => {
                                  if (window.confirm(`Are you sure you want to remove "${ing.name}" from inventory?`)) {
                                    setInventory(prev => (prev || []).filter(item => item.id !== ing.id));
                                    if (typeof recordAuditLog === 'function') {
                                      recordAuditLog(
                                        'INVENTORY_ITEM_DELETED',
                                        ing.id,
                                        `Deleted raw material "${ing.name}" (${stockNum} ${ing.unit} @ ${settings.currency} ${costNum.toFixed(2)}/${ing.unit})`
                                      );
                                    }
                                  }
                                }}
                                className="p-1.5 text-slate-400 hover:text-rose-600 hover:bg-rose-50 rounded-lg transition-colors cursor-pointer"
                                title="Remove Material"
                              >
                                <Trash2 className="h-4 w-4" />
                              </button>
                            </div>
                          </td>
                        </tr>
                      );
                    })
                  )}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {/* VIEW 6: TABLE MANAGEMENT */}
        {activeTab === 'tables' && (
          <div className="flex-1 overflow-y-auto p-6 space-y-6">
            <div className="flex items-center justify-between">
              <div>
                <h2 className="text-xl font-black text-slate-900">Table &amp; Floor Management</h2>
                <p className="text-xs text-slate-500 mt-0.5">Floor layout, real-time occupancy status, and table launching.</p>
              </div>
              <button
                type="button"
                onClick={() => setAddTableModalOpen(true)}
                className="px-3.5 py-2 bg-[#ff5500] hover:bg-orange-600 text-white font-bold rounded-xl text-xs flex items-center gap-1.5 shadow-xs cursor-pointer"
              >
                <Plus className="h-3.5 w-3.5" /> Add Table
              </button>
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 lg:grid-cols-4 gap-4">
              {floorTables.map(tbl => {
                const isOccupied = tbl.status === 'OCCUPIED';
                return (
                  <div key={tbl.id} className={`bg-white rounded-2xl border p-5 shadow-xs flex flex-col justify-between ${
                    isOccupied ? 'border-orange-300 ring-2 ring-orange-500/10' : 'border-slate-200'
                  }`}>
                    <div>
                      <div className="flex justify-between items-start">
                        <div>
                          <span className="text-[10px] font-mono font-bold text-slate-400">{tbl.id}</span>
                          <h3 className="text-base font-black text-slate-900">{tbl.name}</h3>
                          <p className="text-xs text-slate-500">{tbl.zone}</p>
                        </div>
                        <span className={`px-2 py-0.5 rounded-full text-[10px] font-bold ${
                          isOccupied ? 'bg-orange-100 text-[#ff5500]' : 'bg-emerald-100 text-emerald-700'
                        }`}>
                          {tbl.status}
                        </span>
                      </div>
                      <p className="text-xs font-mono text-slate-600 mt-3">Capacity: {tbl.capacity} Seats</p>
                    </div>

                    <div className="mt-4 pt-3 border-t border-slate-100 flex gap-2">
                      <button
                        type="button"
                        onClick={() => {
                          setSelectedTable(tbl);
                          setOrderMode('DINING');
                          setActiveTab('pos');
                        }}
                        className="flex-1 py-1.5 bg-slate-900 hover:bg-slate-800 text-white rounded-lg text-xs font-bold cursor-pointer"
                      >
                        {isOccupied ? 'Open Order' : 'Seat Table'}
                      </button>
                      {currentUser.role === 'Administrator' && !isOccupied && floorTables.length > 1 && (
                        <button
                          type="button"
                          onClick={() => {
                            setFloorTables(prev => prev.filter(t => t.id !== tbl.id));
                            recordAuditLog('TABLE_DELETED', tbl.id, `Admin removed table ${tbl.name} (${tbl.id})`);
                          }}
                          className="p-1.5 text-slate-400 hover:text-rose-600 rounded-lg hover:bg-slate-100 transition-colors cursor-pointer"
                          title="Admin: Delete Table"
                        >
                          <Trash2 className="h-4 w-4" />
                        </button>
                      )}
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        )}

        {/* VIEW 7: CASHIER SHIFTS & DRAWER */}
        {activeTab === 'shifts' && (
          <div className="flex-1 overflow-y-auto p-6 space-y-6">
            {(() => {
              // 1. Calculate Approved Cash-Out Payouts
              const allPayouts = currentShift.payouts || [];
              const approvedPayouts = allPayouts.filter(p => p.status === 'APPROVED' || !p.status);
              const pendingPayouts = allPayouts.filter(p => p.status === 'PENDING');
              const totalApprovedCashOut = approvedPayouts.reduce((acc, p) => acc + (parseFloat(p.amount) || 0), 0);

              // 2. Calculate Cash Sales for current shift
              let shiftCashSales = 0;
              let shiftCardSales = 0;
              let shiftOtherSales = 0;
              let shiftTotalBills = 0;

              if (Array.isArray(transactions)) {
                for (let i = 0; i < transactions.length; i++) {
                  const t = transactions[i];
                  if (!t) continue;
                  const method = String(t.paymentMethod || '').trim().toUpperCase();
                  const matchesShift = t.shiftId 
                    ? t.shiftId === currentShift.shiftId 
                    : (extractDateStr(t.date) === currentShift.openedDate && !t.shiftId);

                  if (matchesShift) {
                    shiftTotalBills += 1;
                    const amt = Number(t.total) || 0;
                    if (method === 'CASH') shiftCashSales += amt;
                    else if (method === 'CARD') shiftCardSales += amt;
                    else shiftOtherSales += amt;
                  }
                }
              }

              // 3. Physical Cash Notes Count
              let countedCash = 0;
              let hasCounted = false;
              if (denominations && typeof denominations === 'object') {
                Object.entries(denominations).forEach(([denom, count]) => {
                  const n = Number(count) || 0;
                  if (n > 0) hasCounted = true;
                  countedCash += Number(denom) * n;
                });
              }

              // 4. Expected Drawer Cash: Float + Sales - CashOut
              const openingFloat = Number(currentShift.startingFloat) || 10000.00;
              const expectedCash = Number((openingFloat + shiftCashSales - totalApprovedCashOut).toFixed(2));

              // 5. Drawer Variance = Counted Cash - Expected Cash
              const currentShiftVariance = Number((countedCash - expectedCash).toFixed(2));
              const carriedShortage = Number(unsettledVariance) || 0;
              const netTotalVariance = Number((currentShiftVariance + carriedShortage).toFixed(2));

              return (
                <>
                  {/* Shift Top Header */}
                  <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4">
                    <div>
                      <h2 className="text-xl font-black text-slate-900">Cashier Shift &amp; Drawer Balancing</h2>
                      <p className="text-xs text-slate-500 mt-0.5">
                        Active Shift: <strong className="font-mono text-slate-800">{currentShift.shiftId}</strong> • Opened by {currentShift.openedBy} at {currentShift.openedAt}
                      </p>
                    </div>

                    <div className="flex items-center gap-2 shrink-0">
                      {/* Settle Drawer Button */}
                      <button
                        type="button"
                        onClick={() => {
                          const currentFloatVal = Number(currentShift.startingFloat) || 10000.00;
                          const inputVal = prompt(
                            `Enter standard base float to retain in drawer for next shift:`,
                            currentFloatVal.toString()
                          );
                          if (inputVal !== null) {
                            const parsed = parseFloat(inputVal);
                            const targetFloat = !isNaN(parsed) && parsed >= 0 ? parsed : currentFloatVal;

                            setCurrentShift(prev => ({
                              ...prev,
                              startingFloat: targetFloat
                            }));
                            setUnsettledVariance(0);

                            recordAuditLog(
                              'DRAWER_SETTLED_MANUAL',
                              currentShift.shiftId,
                              `Shift ${currentShift.shiftId} settled. Base float set to ${settings.currency} ${targetFloat.toFixed(2)}. Unsettled variance cleared.`
                            );
                            alert('Drawer settled and discrepancy balance cleared.');
                          }
                        }}
                        className="px-3.5 py-2 bg-emerald-600 hover:bg-emerald-700 text-white font-bold rounded-xl text-xs flex items-center gap-1.5 shadow-xs cursor-pointer transition-colors"
                      >
                        <CheckCircle2 className="h-4 w-4" />
                        <span>Settle &amp; Bank Drawer</span>
                      </button>

                      {/* Close Shift & Print Z-Report Button */}
                      <button
                        type="button"
                        onClick={() => {
                          const effectiveCountedCash = hasCounted ? countedCash : expectedCash;
                          const effectiveOpeningFloat = Number(currentShift.startingFloat) || 10000.00;

                          const closedShift = {
                            ...currentShift,
                            closedDate: getLocalDateStr(),
                            closedAt: new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }),
                            closedBy: currentUser.name,
                            status: 'CLOSED',
                            metrics: {
                              startingFloat: effectiveOpeningFloat,
                              cashSales: shiftCashSales,
                              cardSales: shiftCardSales,
                              otherSales: shiftOtherSales,
                              grossSales: shiftCashSales + shiftCardSales + shiftOtherSales,
                              totalBills: shiftTotalBills,
                              cashOutTotal: totalApprovedCashOut,
                              approvedPayouts,
                              expectedCash,
                              countedCash: effectiveCountedCash,
                              variance: currentShiftVariance,
                              carriedDiscrepancy: netTotalVariance,
                              denominations: { ...denominations }
                            }
                          };

                          // 1. Save to shift archive
                          setShiftHistory(prev => [closedShift, ...prev]);

                          // 2. Print Z-Report
                          triggerAutoPrint({
                            type: 'Z_REPORT',
                            data: closedShift
                          }, `Shift ${closedShift.shiftId} Closed`);

                          // 3. Tag invoices to this closed shift
                          setTransactions(prev => prev.map(t => {
                            const matchesThisShift = t.shiftId === currentShift.shiftId || (!t.shiftId && extractDateStr(t.date) === currentShift.openedDate);
                            return matchesThisShift ? { ...t, shiftId: currentShift.shiftId } : t;
                          }));

                          // 4. Open new shift using the physical note count as the new base float
                          const nextDate = getLocalDateStr();
                          const shiftSequence = Date.now().toString().slice(-4);
                          const newShiftId = `SHIFT-${nextDate.replace(/-/g, '')}-${shiftSequence}`;

                          setCurrentShift({
                            shiftId: newShiftId,
                            openedDate: nextDate,
                            openedAt: new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }),
                            openedBy: currentUser.name,
                            startingFloat: effectiveCountedCash, // Physical note total becomes the new opening float
                            status: 'OPEN',
                            payouts: []
                          });

                          // NOTE: We deliberately do NOT call setDenominations(...) to reset.
                          // The physical note counts stay intact in the drawer inputs.

                          const shiftVarianceText = currentShiftVariance === 0 
                            ? 'Balanced ($0.00)' 
                            : currentShiftVariance > 0 
                            ? `Overage of ${settings.currency} ${currentShiftVariance.toFixed(2)}` 
                            : `Shortage of ${settings.currency} ${Math.abs(currentShiftVariance).toFixed(2)}`;

                          recordAuditLog(
                            'SHIFT_CLOSED_Z_REPORT',
                            closedShift.shiftId,
                            `Shift closed by ${currentUser.name}. Physical count of ${settings.currency} ${effectiveCountedCash.toFixed(2)} retained as opening float for ${newShiftId}. Note counts preserved. Result: ${shiftVarianceText}.`
                          );
                        }}
                        className="px-4 py-2 bg-[#ff5500] hover:bg-orange-600 text-white font-bold rounded-xl text-xs flex items-center gap-2 shadow-xs cursor-pointer shrink-0"
                      >
                        <Lock className="h-4 w-4" />
                        <span>Close Shift &amp; Print Z-Report</span>
                      </button>
                    </div>
                  </div>

                  {/* 5-Column Equation KPI Cards */}
                  <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-5 gap-3.5">
                    {/* A. Opening Cash Float */}
                    <div className="bg-white p-4 rounded-2xl border border-slate-200 shadow-xs">
                      <p className="text-[10px] font-black uppercase text-slate-400">1. Opening Float</p>
                      <p className="text-xl font-black font-mono text-slate-900 mt-1">
                        {settings.currency} {openingFloat.toFixed(2)}
                      </p>
                      <span className="text-[10px] text-slate-400 font-medium">Drawer base float</span>
                    </div>

                    {/* B. Cash Sales Inflow */}
                    <div className="bg-white p-4 rounded-2xl border border-slate-200 shadow-xs">
                      <p className="text-[10px] font-black uppercase text-slate-400">2. + Cash Sales</p>
                      <p className="text-xl font-black font-mono text-emerald-600 mt-1">
                        +{settings.currency} {shiftCashSales.toFixed(2)}
                      </p>
                      <span className="text-[10px] text-slate-400 font-medium">{shiftTotalBills} orders settled</span>
                    </div>

                    {/* C. Cash Out Disbursements */}
                    <div className="bg-white p-4 rounded-2xl border border-slate-200 shadow-xs">
                      <div className="flex items-center justify-between">
                        <p className="text-[10px] font-black uppercase text-slate-400">3. - Cash Out</p>
                        {pendingPayouts.length > 0 && (
                          <span className="px-1.5 py-0.2 rounded-full text-[9px] font-bold bg-amber-100 text-amber-800">
                            {pendingPayouts.length} Pending
                          </span>
                        )}
                      </div>
                      <p className="text-xl font-black font-mono text-rose-600 mt-1">
                        -{settings.currency} {totalApprovedCashOut.toFixed(2)}
                      </p>
                      <span className="text-[10px] text-slate-400 font-medium">Approved payouts</span>
                    </div>

                    {/* D. Expected Cash in Drawer */}
                    <div className="bg-white p-4 rounded-2xl border-2 border-[#ff5500]/40 bg-orange-50/20 shadow-xs">
                      <p className="text-[10px] font-black uppercase text-[#ff5500]">4. = Expected Cash</p>
                      <p className="text-2xl font-black font-mono text-slate-950 mt-1">
                        {settings.currency} {expectedCash.toFixed(2)}
                      </p>
                      <span className="text-[10px] text-slate-500 font-bold">Float + Sales - CashOut</span>
                    </div>

                    {/* E. Drawer Variance: Counted Cash - Expected Cash */}
                    <div className={`p-4 rounded-2xl border shadow-xs ${
                      !hasCounted
                        ? 'bg-slate-50 border-slate-200 text-slate-700'
                        : currentShiftVariance === 0
                        ? 'bg-emerald-50/70 border-emerald-300 text-emerald-950'
                        : currentShiftVariance > 0
                        ? 'bg-blue-50/70 border-blue-300 text-blue-950'
                        : 'bg-rose-50/70 border-rose-300 text-rose-950'
                    }`}>
                      <div className="flex items-center justify-between">
                        <p className="text-[10px] font-black uppercase tracking-wider opacity-75">
                          5. Drawer Variance
                        </p>
                        {hasCounted && (
                          <span className="text-[9px] font-mono font-bold">
                            Count: {settings.currency}{countedCash.toFixed(0)}
                          </span>
                        )}
                      </div>

                      <p className="text-xl font-black font-mono mt-1">
                        {!hasCounted ? (
                          <span className="text-slate-400 text-sm">Enter Counts Below</span>
                        ) : (
                          `${currentShiftVariance > 0 ? '+' : ''}${settings.currency} ${currentShiftVariance.toFixed(2)}`
                        )}
                      </p>

                      <div className="flex items-center justify-between mt-0.5">
                        <span className="text-[10px] font-bold">
                          {!hasCounted 
                            ? 'Awaiting note count' 
                            : currentShiftVariance === 0 
                            ? '✓ Exactly Balanced' 
                            : currentShiftVariance > 0 
                            ? 'Cash Overage' 
                            : 'Cash Shortage'}
                        </span>
                        {currentUser.role === 'Administrator' && carriedShortage !== 0 && (
                          <button
                            type="button"
                            onClick={() => {
                              if (window.confirm(`Clear carried discrepancy of ${settings.currency} ${carriedShortage.toFixed(2)}?`)) {
                                setUnsettledVariance(0);
                                recordAuditLog(
                                  'VARIANCE_SETTLED',
                                  currentShift.shiftId,
                                  `Admin ${currentUser.name} cleared carried variance of ${settings.currency} ${carriedShortage.toFixed(2)}`
                                );
                              }
                            }}
                            className="text-[9px] font-bold text-indigo-700 hover:text-indigo-900 underline cursor-pointer"
                          >
                            Clear Carry
                          </button>
                        )}
                      </div>
                    </div>
                  </div>

                  {/* Cash Out & Payout Management Section */}
                  <div className="grid grid-cols-1 lg:grid-cols-12 gap-6">
                    {/* Disburse Drawer Cash Out Form */}
                    <div className="lg:col-span-5 bg-white p-5 rounded-2xl border border-slate-200 shadow-xs space-y-4">
                      <div className="flex items-center justify-between border-b border-slate-100 pb-3">
                        <div className="flex items-center gap-2">
                          <div className="p-2 bg-rose-50 text-rose-600 rounded-xl">
                            <Banknote className="h-4 w-4" />
                          </div>
                          <div>
                            <h3 className="text-xs font-black uppercase tracking-wider text-slate-900">
                              Disburse Drawer Cash Out
                            </h3>
                            <p className="text-[10px] text-slate-400">Prints voucher immediately on submission</p>
                          </div>
                        </div>
                      </div>

                      <form
                        onSubmit={(e) => {
                          e.preventDefault();
                          const amt = parseFloat(cashOutForm.amount);
                          if (isNaN(amt) || amt <= 0 || !cashOutForm.reason.trim()) return;

                          const isManager = currentUser.role === 'Administrator' || currentUser.role === 'Manager';
                          const newCashOut = {
                            id: getNextCashOutNumber(),
                            shiftId: currentShift.shiftId,
                            amount: amt,
                            category: cashOutForm.category,
                            reason: cashOutForm.reason.trim(),
                            recipient: cashOutForm.recipient.trim() || 'General Expense',
                            requestedBy: currentUser.name,
                            requestedRole: currentUser.role,
                            createdAt: new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }),
                            date: getLocalDateStr(),
                            status: isManager ? 'APPROVED' : 'PENDING',
                            approvedBy: isManager ? currentUser.name : null,
                            approvedAt: isManager ? new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }) : null
                          };

                          setCurrentShift(prev => ({
                            ...prev,
                            payouts: [newCashOut, ...(prev.payouts || [])]
                          }));

                          setExpenses(prev => [newCashOut, ...prev]);

                          // Trigger immediate thermal voucher print
                          triggerAutoPrint({
                            type: 'CASH_OUT_VOUCHER',
                            data: newCashOut
                          }, `Cash Out Voucher #${newCashOut.id}`);

                          recordAuditLog(
                            isManager ? 'CASH_OUT_APPROVED_DIRECT' : 'CASH_OUT_REQUESTED',
                            newCashOut.id,
                            `Cash out of ${settings.currency} ${amt.toFixed(2)} for "${newCashOut.reason}" (${newCashOut.category}) by ${currentUser.name}`
                          );

                          setCashOutForm({
                            amount: '',
                            category: 'Supplier / Vendor',
                            reason: '',
                            recipient: ''
                          });
                        }}
                        className="space-y-3"
                      >
                        <div className="grid grid-cols-2 gap-3">
                          <div>
                            <label className="block text-[10px] font-extrabold uppercase text-slate-500 mb-1">
                              Amount ({settings.currency})
                            </label>
                            <input
                              type="number"
                              step="0.01"
                              min="1"
                              required
                              value={cashOutForm.amount}
                              onChange={e => setCashOutForm(prev => ({ ...prev, amount: e.target.value }))}
                              placeholder="e.g. 2500.00"
                              className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-mono font-bold text-slate-900 focus:bg-white focus:outline-none focus:border-[#ff5500]"
                            />
                          </div>

                          <div>
                            <label className="block text-[10px] font-extrabold uppercase text-slate-500 mb-1">
                              Category
                            </label>
                            <select
                              value={cashOutForm.category}
                              onChange={e => setCashOutForm(prev => ({ ...prev, category: e.target.value }))}
                              className="w-full px-2.5 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-bold text-slate-900 focus:bg-white focus:outline-none focus:border-[#ff5500]"
                            >
                              <option value="Supplier / Vendor">Supplier / Vendor</option>
                              <option value="Market Produce">Market Produce</option>
                              <option value="Safe Drop">Safe Drop (Bank Deposit)</option>
                              <option value="Petty Cash / Store Supplies">Petty Cash / Supplies</option>
                              <option value="Staff Tip Out">Staff Tip Out</option>
                              <option value="Utilities">Utilities</option>
                              <option value="Other Emergency">Other Emergency</option>
                            </select>
                          </div>
                        </div>

                        <div>
                          <label className="block text-[10px] font-extrabold uppercase text-slate-500 mb-1">
                            Reason &amp; Description
                          </label>
                          <input
                            type="text"
                            required
                            value={cashOutForm.reason}
                            onChange={e => setCashOutForm(prev => ({ ...prev, reason: e.target.value }))}
                            placeholder="e.g. Paid seafood driver for 5kg reef fish"
                            className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs text-slate-900 focus:bg-white focus:outline-none focus:border-[#ff5500]"
                          />
                        </div>

                        <div>
                          <label className="block text-[10px] font-extrabold uppercase text-slate-500 mb-1">
                            Paid To / Recipient
                          </label>
                          <input
                            type="text"
                            value={cashOutForm.recipient}
                            onChange={e => setCashOutForm(prev => ({ ...prev, recipient: e.target.value }))}
                            placeholder="e.g. Fish Vendor Sunil / Safe Vault"
                            className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs text-slate-900 focus:bg-white focus:outline-none focus:border-[#ff5500]"
                          />
                        </div>

                        <button
                          type="submit"
                          className="w-full py-2.5 bg-slate-900 hover:bg-slate-800 text-white font-bold rounded-xl text-xs flex items-center justify-center gap-2 shadow-sm transition-all cursor-pointer"
                        >
                          <Send className="h-3.5 w-3.5" />
                          <span>Submit &amp; Print Cash Out Voucher</span>
                        </button>
                      </form>
                    </div>

                    {/* Cash Out Log List */}
                    <div className="lg:col-span-7 bg-white p-5 rounded-2xl border border-slate-200 shadow-xs flex flex-col justify-between space-y-4">
                      <div>
                        <div className="flex items-center justify-between border-b border-slate-100 pb-3">
                          <div className="flex items-center gap-2">
                            <div className="p-2 bg-indigo-50 text-indigo-600 rounded-xl">
                              <ShieldCheck className="h-4 w-4" />
                            </div>
                            <div>
                              <h3 className="text-xs font-black uppercase tracking-wider text-slate-900">
                                Shift Disbursements &amp; Approvals
                              </h3>
                              <p className="text-[10px] text-slate-400">Current shift drawer payouts</p>
                            </div>
                          </div>
                          <span className="text-[10px] font-bold font-mono px-2 py-0.5 rounded-full bg-slate-100 text-slate-600">
                            {(currentShift.payouts || []).length} Records
                          </span>
                        </div>

                        <div className="mt-3 divide-y divide-slate-100 max-h-72 overflow-y-auto pr-1">
                          {(currentShift.payouts || []).length === 0 ? (
                            <div className="py-8 text-center text-slate-400 text-xs italic">
                              No cash out transactions requested during this shift.
                            </div>
                          ) : (
                            (currentShift.payouts || []).map((item) => {
                              const isApproved = item.status === 'APPROVED' || !item.status;
                              const isPending = item.status === 'PENDING';
                              const canDirectApprove = currentUser.role === 'Administrator' || currentUser.role === 'Manager';

                              return (
                                <div key={item.id} className="py-2.5 flex items-center justify-between gap-3 hover:bg-slate-50/50 transition-colors">
                                  <div className="space-y-0.5">
                                    <div className="flex items-center gap-2">
                                      <span className="font-bold text-xs text-slate-900">{item.reason}</span>
                                      <span className={`px-1.5 py-0.2 rounded text-[9px] font-black uppercase ${
                                        isApproved
                                          ? 'bg-emerald-100 text-emerald-800'
                                          : isPending
                                          ? 'bg-amber-100 text-amber-800'
                                          : 'bg-rose-100 text-rose-800'
                                      }`}>
                                        {item.status || 'APPROVED'}
                                      </span>
                                    </div>
                                    <p className="text-[10px] text-slate-500">
                                      <span className="font-semibold text-slate-700">{item.category}</span> • Recipient: {item.recipient || 'N/A'} • Req by {item.requestedBy}
                                    </p>
                                    {isApproved && item.approvedBy && (
                                      <p className="text-[9px] font-mono text-emerald-700">
                                        ✓ Approved by {item.approvedBy} ({item.approvedAt || 'Verified'})
                                      </p>
                                    )}
                                  </div>

                                  <div className="flex items-center gap-2 shrink-0">
                                    <span className="font-mono font-black text-sm text-slate-900">
                                      {settings.currency} {(parseFloat(item.amount) || 0).toFixed(2)}
                                    </span>

                                    {isPending && (
                                      <div className="flex items-center gap-1">
                                        <button
                                          type="button"
                                          onClick={() => {
                                            if (canDirectApprove) {
                                              const updatedItem = {
                                                ...item,
                                                status: 'APPROVED',
                                                approvedBy: currentUser.name,
                                                approvedAt: new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })
                                              };

                                              setCurrentShift(prev => ({
                                                ...prev,
                                                payouts: prev.payouts.map(p => p.id === item.id ? updatedItem : p)
                                              }));

                                              setExpenses(prev => prev.map(p => p.id === item.id ? updatedItem : p));

                                              triggerAutoPrint({
                                                type: 'CASH_OUT_VOUCHER',
                                                data: updatedItem
                                              }, `Authorized Cash Out #${item.id}`);

                                              recordAuditLog('CASH_OUT_APPROVED', item.id, `Manager ${currentUser.name} approved cash out ${item.id} of ${settings.currency} ${item.amount}`);
                                            } else {
                                              setCashOutApprovalModal({
                                                open: true,
                                                item,
                                                managerPin: '',
                                                error: ''
                                              });
                                            }
                                          }}
                                          className="px-2.5 py-1 bg-emerald-600 hover:bg-emerald-700 text-white rounded-lg text-[10px] font-bold flex items-center gap-1 shadow-xs cursor-pointer"
                                        >
                                          <Check className="h-3 w-3 stroke-[3]" />
                                          <span>Approve</span>
                                        </button>

                                        <button
                                          type="button"
                                          onClick={() => {
                                            const updatedItem = {
                                              ...item,
                                              status: 'REJECTED',
                                              rejectedBy: currentUser.name
                                            };

                                            setCurrentShift(prev => ({
                                              ...prev,
                                              payouts: prev.payouts.map(p => p.id === item.id ? updatedItem : p)
                                            }));

                                            setExpenses(prev => prev.map(p => p.id === item.id ? updatedItem : p));

                                            recordAuditLog('CASH_OUT_REJECTED', item.id, `${currentUser.name} rejected cash out request ${item.id}`);
                                          }}
                                          className="px-2 py-1 bg-rose-50 hover:bg-rose-100 text-rose-700 rounded-lg text-[10px] font-bold cursor-pointer"
                                        >
                                          Reject
                                        </button>
                                      </div>
                                    )}

                                    {isApproved && (
                                      <button
                                        type="button"
                                        onClick={() => {
                                          triggerAutoPrint({
                                            type: 'CASH_OUT_VOUCHER',
                                            data: item
                                          }, `Cash Out Ref ${item.id}`);
                                        }}
                                        className="p-1 text-slate-400 hover:text-slate-800 rounded hover:bg-slate-100 cursor-pointer"
                                        title="Reprint Cash Out Voucher"
                                      >
                                        <Printer className="h-3.5 w-3.5" />
                                      </button>
                                    )}
                                  </div>
                                </div>
                              );
                            })
                          )}
                        </div>
                      </div>
                    </div>
                  </div>

                  {/* Physical Note Breakdown Section: Updates Variance in Real-Time */}
                  <div className="bg-white p-5 rounded-2xl border border-slate-200 space-y-4">
                    <div className="flex items-center justify-between">
                      <h3 className="text-xs font-black uppercase tracking-wider text-slate-900 flex items-center gap-2">
                        <Coins className="h-4 w-4 text-[#ff5500]" /> Physical Cash Note Count
                      </h3>
                      <div className="flex items-center gap-3">
                        <span className="text-xs font-mono font-black px-3 py-1 rounded-xl bg-slate-900 text-white shadow-xs">
                          Counted Total: {settings.currency} {countedCash.toFixed(2)}
                        </span>
                        {hasCounted && (
                          <span className={`text-xs font-mono font-black px-3 py-1 rounded-xl ${
                            currentShiftVariance === 0 ? 'bg-emerald-100 text-emerald-800' : currentShiftVariance > 0 ? 'bg-blue-100 text-blue-800' : 'bg-rose-100 text-rose-800'
                          }`}>
                            Variance: {currentShiftVariance > 0 ? '+' : ''}{settings.currency} {currentShiftVariance.toFixed(2)}
                          </span>
                        )}
                      </div>
                    </div>

                    <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-6 gap-2 sm:gap-3">
                      {[5000, 1000, 500, 100, 50, 20].map(denom => (
                        <div key={denom} className="flex flex-col justify-between p-2.5 bg-slate-50 rounded-xl border border-slate-200">
                          <div className="flex justify-between items-center text-xs">
                            <span className="font-mono font-bold text-slate-700">{settings.currency} {denom}</span>
                            <span className="text-[10px] font-mono text-slate-400">
                              {settings.currency} {((Number(denominations[denom]) || 0) * denom).toFixed(0)}
                            </span>
                          </div>
                          <input
                            type="number"
                            min="0"
                            value={denominations[denom] !== undefined && denominations[denom] !== 0 ? denominations[denom] : ''}
                            onChange={e => {
                              const val = e.target.value === '' ? '' : Math.max(0, parseInt(e.target.value) || 0);
                              setDenominations(prev => ({ ...prev, [denom]: val }));
                            }}
                            placeholder="0 notes"
                            className="w-full mt-2 px-2 py-1.5 bg-white border border-slate-300 rounded-lg text-center text-xs font-mono font-bold text-slate-900 focus:outline-none focus:border-[#ff5500]"
                          />
                        </div>
                      ))}
                    </div>
                  </div>
                </>
              );
            })()}
          </div>
        )}

        {/* VIEW 8: KITCHEN DISPLAY (KDS) */}
        {activeTab === 'kds' && (
          <div className="flex-1 overflow-y-auto p-6 space-y-6">
            <h2 className="text-xl font-black text-slate-900">Live Kitchen Display (KOT)</h2>
            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 xl:grid-cols-4 gap-3 sm:gap-4">
              {activeOrders.filter(o => o.items.some(i => i.department === 'Kitchen')).map(order => (
                <div key={order.orderId} className="bg-white rounded-2xl border-2 border-rose-200 p-4 shadow-xs">
                  <div className="flex justify-between items-center border-b border-slate-100 pb-2 mb-3">
                    <div>
                      <h4 className="font-extrabold text-sm text-slate-900">{order.tableName}</h4>
                      <span className="text-[10px] text-slate-400 font-mono">Order #{order.orderId}</span>
                    </div>
                    <span className="px-2 py-0.5 bg-rose-100 text-rose-700 font-bold text-xs rounded-full">
                      {order.sentAt}
                    </span>
                  </div>
                  <div className="space-y-2">
                    {order.items.filter(i => i.department === 'Kitchen').map((item, idx) => (
                      <div key={idx} className="p-2 bg-slate-50 rounded-lg">
                        <p className="font-bold text-xs text-slate-900">{item.qty}x {item.name}</p>
                        {item.notes && <p className="text-[10px] text-rose-600 font-semibold mt-0.5">&gt; {item.notes}</p>}
                      </div>
                    ))}
                  </div>
                </div>
              ))}
            </div>
          </div>
        )}

        {/* VIEW 9: BAR DISPLAY (BOT) */}
        {activeTab === 'bar' && (
          <div className="flex-1 overflow-y-auto p-6 space-y-6">
            <h2 className="text-xl font-black text-slate-900">Live Bar Display (BOT)</h2>
            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 xl:grid-cols-4 gap-3 sm:gap-4">
              {activeOrders.filter(o => o.items.some(i => i.department === 'Bar')).map(order => (
                <div key={order.orderId} className="bg-white rounded-2xl border-2 border-indigo-200 p-4 shadow-xs">
                  <div className="flex justify-between items-center border-b border-slate-100 pb-2 mb-3">
                    <div>
                      <h4 className="font-extrabold text-sm text-slate-900">{order.tableName}</h4>
                      <span className="text-[10px] text-slate-400 font-mono">Order #{order.orderId}</span>
                    </div>
                    <span className="px-2 py-0.5 bg-indigo-100 text-indigo-700 font-bold text-xs rounded-full">
                      {order.sentAt}
                    </span>
                  </div>
                  <div className="space-y-2">
                    {order.items.filter(i => i.department === 'Bar').map((item, idx) => (
                      <div key={idx} className="p-2 bg-slate-50 rounded-lg">
                        <p className="font-bold text-xs text-slate-900">{item.qty}x {item.name}</p>
                        {item.notes && <p className="text-[10px] text-indigo-600 font-semibold mt-0.5">&gt; {item.notes}</p>}
                      </div>
                    ))}
                  </div>
                </div>
              ))}
            </div>
          </div>
        )}

        {/* VIEW 10: MENU MANAGEMENT */}
        {activeTab === 'menu_admin' && (
          <div className="flex-1 overflow-y-auto p-6 space-y-6">
            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4">
              <div>
                <h2 className="text-xl font-black text-slate-900">Menu Management &amp; Dishes</h2>
                <p className="text-xs text-slate-500 mt-0.5">Manage 150+ dishes, upload food photos, configure categories, and link recipes.</p>
              </div>
              <div className="flex items-center gap-3">
                <div className="flex items-center gap-1.5 bg-white border border-slate-200 rounded-xl px-2.5 py-1 shadow-xs">
                  <span className="text-[11px] font-bold text-slate-500">Category:</span>
                  <select
                    value={adminMenuCategory}
                    onChange={e => setAdminMenuCategory(e.target.value)}
                    className="text-xs font-bold text-slate-900 bg-white focus:outline-none cursor-pointer py-1 pr-1"
                    style={{ color: '#0f172a', backgroundColor: '#ffffff' }}
                  >
                    {categoriesList.map(cat => (
                      <option key={cat} value={cat} style={{ color: '#0f172a', backgroundColor: '#ffffff' }}>
                        {cat === 'All' ? `All (${menuItems.length})` : `${cat} (${menuItems.filter(m => m.category === cat).length})`}
                      </option>
                    ))}
                  </select>
                </div>

                <label className="px-3.5 py-2.5 bg-slate-800 hover:bg-slate-700 text-white font-bold rounded-xl text-xs flex items-center gap-2 shadow-xs shrink-0 cursor-pointer transition-colors">
                  <Upload className="h-4 w-4" />
                  <span>Import from PDF</span>
                  <input
                    type="file"
                    accept="application/pdf"
                    className="hidden"
                    onChange={async (e) => {
                      const file = e.target.files?.[0];
                      if (!file) return;

                      try {
                        const imported = await extractMenuFromPDF(file);
                        if (imported.length === 0) {
                          alert("No dishes with readable prices were detected. Ensure the PDF contains selectable text rather than flattened image scans.");
                          return;
                        }

                        setMenuItems(prev => [...prev, ...imported]);
                        recordAuditLog('MENU_IMPORTED', 'PDF', `Imported ${imported.length} items from ${file.name}`);
                        alert(`Successfully imported ${imported.length} items from ${file.name}!`);
                      } catch (err) {
                        console.error("PDF Parsing error:", err);
                        alert("Failed to parse PDF document. Please verify the file integrity.");
                      } finally {
                        e.target.value = '';
                      }
                    }}
                  />
                </label>

                <button
                  type="button"
                  onClick={() => setAddItemModalOpen(true)}
                  className="px-4 py-2.5 bg-[#ff5500] hover:bg-orange-600 text-white font-bold rounded-xl text-xs flex items-center gap-2 shadow-xs shrink-0 cursor-pointer"
                >
                  <Plus className="h-4 w-4" /> Add New Menu Item
                </button>
              </div>
            </div>

            <div className="bg-white rounded-2xl border border-slate-200 overflow-hidden shadow-xs">
              <table className="w-full text-left text-xs">
                <thead className="bg-slate-50 text-[10px] font-black uppercase text-slate-400 border-b border-slate-200">
                  <tr>
                    <th className="py-3 px-4">Item &amp; Photo</th>
                    <th className="py-3 px-4">Dept</th>
                    <th className="py-3 px-4">Category</th>
                    <th className="py-3 px-4">Selling Price</th>
                    <th className="py-3 px-4">BOM Raw Cost</th>
                    <th className="py-3 px-4 text-right">Actions</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {menuItems
                    .filter(item => adminMenuCategory === 'All' || item.category === adminMenuCategory)
                    .map(item => {
                      const { cogs } = calculateDishAvailability(item.recipe);
                      return (
                        <tr key={item.id} className="hover:bg-slate-50/70">
                          <td className="py-3 px-4">
                            <div className="flex items-center gap-3">
                              <label className="relative group cursor-pointer shrink-0" title="Click to upload/change photo">
                                {item.imageUrl ? (
                                  <img src={item.imageUrl} alt={item.name} className="h-11 w-11 rounded-xl object-cover border border-slate-200 group-hover:opacity-75 transition-opacity" />
                                ) : (
                                  <div className="h-11 w-11 rounded-xl bg-orange-50 border border-orange-200 flex flex-col items-center justify-center text-orange-600 font-bold text-[9px] group-hover:bg-orange-100 transition-colors">
                                    <Upload className="h-3.5 w-3.5 mb-0.5" />
                                    <span>ADD</span>
                                  </div>
                                )}
                                <span className="absolute inset-0 bg-black/40 text-white rounded-xl text-[9px] font-bold flex items-center justify-center opacity-0 group-hover:opacity-100 transition-opacity">
                                  Change
                                </span>
                                <input
                                  type="file"
                                  accept="image/*"
                                  className="hidden"
                                  onChange={e => {
                                    const file = e.target.files?.[0];
                                    if (file) {
                                      const reader = new FileReader();
                                      reader.onload = ev => {
                                        if (ev.target?.result) {
                                          setMenuItems(prev => prev.map(m => m.id === item.id ? { ...m, imageUrl: ev.target.result } : m));
                                          recordAuditLog('UPDATE_DISH_PHOTO', item.id, `Uploaded new photo for ${item.name}`);
                                        }
                                      };
                                      reader.readAsDataURL(file);
                                    }
                                  }}
                                />
                              </label>
                              <div>
                                <p className="font-extrabold text-slate-900">{item.name}</p>
                                <span className="text-[10px] text-slate-400 line-clamp-1">{item.description}</span>
                              </div>
                            </div>
                          </td>
                          <td className="py-3 px-4">
                            <span className={`px-2 py-0.5 rounded text-[10px] font-bold ${
                              item.department === 'Kitchen' ? 'bg-rose-100 text-rose-700' : 'bg-indigo-100 text-indigo-700'
                            }`}>
                              {item.department}
                            </span>
                          </td>
                          <td className="py-3 px-4 text-slate-600 font-medium">{item.category}</td>
                          <td className="py-3 px-4 font-mono font-bold text-[#ff5500]">{settings.currency} {item.price.toFixed(2)}</td>
                          <td className="py-3 px-4 font-mono text-slate-500">{settings.currency} {cogs.toFixed(2)}</td>
                          <td className="py-3 px-4 text-right">
                            <div className="flex items-center justify-end gap-1">
                              <button
                                type="button"
                                onClick={() => {
                                  setEditingMenuItem({ ...item });
                                  setIsEditModalOpen(true);
                                }}
                                className="p-1.5 text-slate-400 hover:text-slate-700 hover:bg-slate-100 rounded-lg transition-colors cursor-pointer"
                                title="Edit Item"
                              >
                                <Edit3 className="h-4 w-4" />
                              </button>

                              <button
                                type="button"
                                onClick={() => {
                                  if (window.confirm(`Are you sure you want to delete "${item.name}" from the menu?`)) {
                                    setMenuItems(prev => prev.filter(m => m.id !== item.id));
                                    recordAuditLog('DELETE_MENU_ITEM', item.id, `Removed ${item.name} from menu.`);
                                  }
                                }}
                                className="p-1.5 text-slate-400 hover:text-rose-600 hover:bg-rose-50 rounded-lg transition-colors cursor-pointer"
                                title="Delete Item"
                              >
                                <Trash2 className="h-4 w-4" />
                              </button>
                            </div>
                          </td>
                        </tr>
                      );
                    })}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {/* VIEW 12: STAFF MANAGEMENT */}
        {activeTab === 'staff' && (
          <div className="flex-1 overflow-y-auto p-6 space-y-6">
            <div className="flex items-center justify-between">
              <div>
                <h2 className="text-xl font-black text-slate-900">Staff Management, Roles &amp; Wages</h2>
                <p className="text-xs text-slate-500 mt-0.5">
                  Configure employee credentials, security PINs, and base salary structures linked directly to payroll.
                </p>
              </div>
              <button
                type="button"
                onClick={() => {
                  setEditingStaffId(null);
                  setNewStaffForm({
                    name: '',
                    role: 'Cashier',
                    pin: '',
                    email: '',
                    basicSalary: 35000,
                    budgetaryAllowance: 2500,
                    otherAllowances: 0,
                    fixedBonus: 0,
                    overtimeRate: 250,
                    epfEtfEnabled: true
                  });
                  setAddStaffModalOpen(true);
                }}
                className="px-4 py-2 bg-[#ff5500] hover:bg-orange-600 text-white font-bold rounded-xl text-xs flex items-center gap-2 shadow-xs cursor-pointer transition-all"
              >
                <Plus className="h-4 w-4" /> Add Employee
              </button>
            </div>

            <div className="bg-white rounded-2xl border border-slate-200 overflow-hidden shadow-xs">
              <table className="w-full text-left text-xs">
                <thead className="bg-slate-50 text-[10px] font-black uppercase text-slate-400 border-b border-slate-200">
                  <tr>
                    <th className="py-3 px-4">Staff Member</th>
                    <th className="py-3 px-4">Role</th>
                    <th className="py-3 px-4 text-right">Basic Salary</th>
                    <th className="py-3 px-4 text-right">BRA / Allowances</th>
                    <th className="py-3 px-4 text-center">EPF / ETF</th>
                    <th className="py-3 px-4">Security PIN</th>
                    <th className="py-3 px-4 text-right">Actions</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {staffList.map(member => (
                    <tr key={member.id} className="hover:bg-slate-50 transition-colors">
                      <td className="py-3 px-4">
                        <div className="flex items-center gap-2.5">
                          <div className="h-8 w-8 rounded-full bg-slate-900 text-white font-black text-xs flex items-center justify-center">
                            {member.avatar}
                          </div>
                          <div>
                            <p className="font-extrabold text-slate-900">{member.name}</p>
                            <span className="text-[10px] text-slate-400">{member.email}</span>
                          </div>
                        </div>
                      </td>

                      <td className="py-3 px-4">
                        <span className="px-2 py-0.5 bg-orange-100 text-[#ff5500] rounded font-bold text-[10px]">
                          {member.role}
                        </span>
                      </td>

                      <td className="py-3 px-4 text-right font-mono font-bold text-slate-900">
                        {settings.currency} {(member.basicSalary ?? 35000).toLocaleString('en-US', { minimumFractionDigits: 2 })}
                      </td>

                      <td className="py-3 px-4 text-right font-mono text-slate-600">
                        +{settings.currency} {((member.budgetaryAllowance ?? 2500) + (member.otherAllowances ?? 0)).toFixed(2)}
                      </td>

                      <td className="py-3 px-4 text-center">
                        <span className={`px-2 py-0.5 rounded-full text-[10px] font-bold ${
                          member.epfEtfEnabled !== false ? 'bg-emerald-100 text-emerald-800' : 'bg-slate-100 text-slate-500'
                        }`}>
                          {member.epfEtfEnabled !== false ? 'Enrolled' : 'Exempt'}
                        </span>
                      </td>

                      <td className="py-3 px-4 font-mono text-slate-500">•••• ({member.pin})</td>

                      <td className="py-3 px-4 text-right">
                        <div className="flex items-center justify-end gap-1.5">
  {/* DIRECT RUN PAYROLL BUTTON */}
  <button
    type="button"
    onClick={() => {
      const curPeriod = `${new Date().getFullYear()}-${String(new Date().getMonth() + 1).padStart(2, '0')}`;
      setPayrollInputForm({
        staffId: member.id,
        period: curPeriod,
        epfEtfEnabled: member.epfEtfEnabled !== false,
        basicSalary: member.basicSalary ?? 35000,
        budgetaryAllowance: member.budgetaryAllowance ?? 2500,
        otherAllowances: member.otherAllowances ?? 0,
        serviceChargeBonus: Math.round((salesMetrics.serviceCharge || 0) / Math.max(1, staffList.length)),
        incentiveBonus: member.fixedBonus ?? 0,
        overtimeHours: 0,
        overtimeRate: member.overtimeRate ?? 250,
        salaryAdvance: 0,
        otherDeductions: 0,
        standardWorkingDays: 26,
        workedDays: 26,
        paidLeaves: 0,
        unpaidLeaves: 0,
        holidaysCount: 4,
        shortShiftsCount: 0,
        notes: `Standard wages for ${member.name}`
      });
      setEditingPayrollId(null);
      setActiveTab('payroll');
      setPayrollSubTab('payslips');
      setProcessPayModalOpen(true);
    }}
    className="px-2.5 py-1 bg-emerald-50 hover:bg-emerald-100 text-emerald-700 font-bold rounded-lg text-xs flex items-center gap-1 transition-colors cursor-pointer"
    title="Issue Pay Slip with this employee's defaults"
  >
    <Briefcase className="h-3 w-3" />
    <span>Issue Pay</span>
  </button>

                          {/* EDIT PROFILE */}
                          <button
                            type="button"
                            onClick={() => {
                              setEditingStaffId(member.id);
                              setNewStaffForm({
                                name: member.name,
                                role: member.role,
                                pin: member.pin,
                                email: member.email || '',
                                basicSalary: member.basicSalary ?? 35000,
                                budgetaryAllowance: member.budgetaryAllowance ?? 2500,
                                otherAllowances: member.otherAllowances ?? 0,
                                fixedBonus: member.fixedBonus ?? 0,
                                overtimeRate: member.overtimeRate ?? 250,
                                epfEtfEnabled: member.epfEtfEnabled !== false
                              });
                              setAddStaffModalOpen(true);
                            }}
                            className="p-1.5 text-slate-400 hover:text-indigo-600 rounded-lg hover:bg-slate-100 transition-colors cursor-pointer"
                            title="Edit Staff Member"
                          >
                            <Edit3 className="h-4 w-4" />
                          </button>

                          {/* DELETE */}
                          {currentUser.role === 'Administrator' && member.id !== currentUser.id && (
                            <button
                              type="button"
                              onClick={() => {
                                if (window.confirm(`Remove staff member ${member.name}?`)) {
                                  setStaffList(prev => prev.filter(s => s.id !== member.id));
                                  recordAuditLog('STAFF_DELETED', member.id, `Removed staff member ${member.name} (${member.role})`);
                                }
                              }}
                              className="p-1.5 text-slate-400 hover:text-rose-600 rounded-lg hover:bg-slate-100 transition-colors cursor-pointer"
                              title="Remove Employee"
                            >
                              <Trash2 className="h-4 w-4" />
                            </button>
                          )}
                        </div>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}
        {/* VIEW: EXTERNAL INVOICES & VENDOR BILLS */}
        {activeTab === 'vendor_bills' && (
          <div className="flex-1 overflow-y-auto p-4 sm:p-6 space-y-6">
            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4">
              <div>
                <h2 className="text-xl font-black text-slate-900 flex items-center gap-2">
                  <FileText className="h-5 w-5 text-[#ff5500]" />
                  Vendor Invoices &amp; Accounts Payable
                </h2>
                <p className="text-xs text-slate-500 mt-0.5">
                  Track external supplier invoices paid via bank transfer, cheque, or card outside of cashier cash drawers.
                </p>
              </div>

              <div className="flex flex-wrap items-center gap-2 shrink-0">
                {/* 1. EXCEL EXPORT BUTTON */}
                <button
                  type="button"
                  onClick={() => {
                    const billsList = Array.isArray(vendorBills) ? vendorBills : [];
                    if (billsList.length === 0) {
                      alert('No vendor bills available to export.');
                      return;
                    }

                    const exportRows = billsList.map(bill => ({
                      'Invoice #': bill.invoiceNumber || bill.id,
                      'Bill Date': bill.billDate || 'N/A',
                      'Due Date': bill.dueDate || 'N/A',
                      'Vendor / Payee': bill.vendorName,
                      'Category': bill.category,
                      'Payment Method': bill.paymentMethod,
                      'Amount': Number(bill.amount) || 0,
                      'Payment Status': bill.paymentStatus,
                      'Recorded By': bill.recordedBy || 'N/A',
                      'Notes': bill.notes || ''
                    }));

                    exportReportToExcel('Vendor_Bills_Accounts_Payable', exportRows, 'Vendor_Invoices');
                  }}
                  className="px-3.5 py-2 bg-emerald-600 hover:bg-emerald-700 text-white font-bold rounded-xl text-xs flex items-center gap-1.5 shadow-xs cursor-pointer transition-all active:scale-95"
                  title="Export Vendor Invoices to Excel (.xlsx)"
                >
                  <Download className="h-3.5 w-3.5" />
                  <span>Export Excel</span>
                </button>

                {/* 2. PDF / PRINT BUTTON */}
                <button
                  type="button"
                  onClick={() => {
                    window.print();
                  }}
                  className="px-3.5 py-2 bg-slate-900 hover:bg-slate-800 text-white font-bold rounded-xl text-xs flex items-center gap-1.5 shadow-xs cursor-pointer transition-all active:scale-95"
                  title="Print or Save Vendor Bills as PDF"
                >
                  <Printer className="h-3.5 w-3.5 text-orange-400" />
                  <span>PDF / Print</span>
                </button>

                {/* 3. RECORD VENDOR BILL BUTTON */}
                <button
                  type="button"
                  onClick={() => {
                    setVendorBillForm({
                      invoiceNumber: '',
                      vendorName: '',
                      category: 'Food & Beverage Supply',
                      billDate: getLocalDateStr(),
                      dueDate: getLocalDateStr(),
                      amount: '',
                      paymentMethod: 'BANK_TRANSFER',
                      paymentStatus: 'PAID',
                      notes: ''
                    });
                    setAddVendorBillModalOpen(true);
                  }}
                  className="px-4 py-2 bg-[#ff5500] hover:bg-orange-600 text-white font-bold rounded-xl text-xs flex items-center gap-2 shadow-xs cursor-pointer transition-all active:scale-95"
                >
                  <Plus className="h-4 w-4" />
                  <span>Record Vendor Bill</span>
                </button>
              </div>
            </div>
            
            {/* KPI Cards */}
            {(() => {
              const totalBills = vendorBills.reduce((acc, b) => acc + (Number(b.amount) || 0), 0);
              const paidBills = vendorBills.filter(b => b.paymentStatus === 'PAID').reduce((acc, b) => acc + (Number(b.amount) || 0), 0);
              const unpaidBills = vendorBills.filter(b => b.paymentStatus !== 'PAID').reduce((acc, b) => acc + (Number(b.amount) || 0), 0);

              return (
                <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
                  <div className="bg-white rounded-2xl border border-slate-200 p-4 shadow-xs">
                    <span className="text-[10px] font-black uppercase text-slate-400">Total Invoiced</span>
                    <p className="text-2xl font-black font-mono text-slate-900 mt-1">{settings.currency} {totalBills.toFixed(2)}</p>
                    <span className="text-[11px] text-slate-500">{vendorBills.length} recorded invoices</span>
                  </div>

                  <div className="bg-white rounded-2xl border border-slate-200 p-4 shadow-xs">
                    <span className="text-[10px] font-black uppercase text-slate-400">Paid Invoices (OPEX)</span>
                    <p className="text-2xl font-black font-mono text-emerald-600 mt-1">{settings.currency} {paidBills.toFixed(2)}</p>
                    <span className="text-[11px] text-slate-500">Linked to P&amp;L expenses</span>
                  </div>

                  <div className="bg-white rounded-2xl border border-slate-200 p-4 shadow-xs">
                    <span className="text-[10px] font-black uppercase text-slate-400">Outstanding Accounts Payable</span>
                    <p className="text-2xl font-black font-mono text-rose-600 mt-1">{settings.currency} {unpaidBills.toFixed(2)}</p>
                    <span className="text-[11px] text-slate-500">Unsettled credit terms</span>
                  </div>
                </div>
              );
            })()}

            {/* Invoices Table */}
            <div className="bg-white rounded-2xl border border-slate-200 overflow-hidden shadow-xs">
              <table className="w-full text-left text-xs">
                <thead className="bg-slate-50 text-[10px] font-black uppercase text-slate-400 border-b border-slate-200">
                  <tr>
                    <th className="py-2.5 px-3">Invoice #</th>
                    <th className="py-2.5 px-3">Date</th>
                    <th className="py-2.5 px-3">Vendor / Payee</th>
                    <th className="py-2.5 px-3">Category</th>
                    <th className="py-2.5 px-3">Payment Method</th>
                    <th className="py-2.5 px-3 text-right">Amount</th>
                    <th className="py-2.5 px-3 text-center">Status</th>
                    <th className="py-2.5 px-3 text-right">Actions</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {vendorBills.length === 0 ? (
                    <tr>
                      <td colSpan={8} className="py-8 text-center text-slate-400 italic">
                        No external invoices recorded yet. Click &ldquo;Record Vendor Bill&rdquo; to add distributor or utility bills.
                      </td>
                    </tr>
                  ) : (
                    vendorBills.map(bill => (
                      <tr key={bill.id} className="hover:bg-slate-50 transition-colors">
                        <td className="py-3 px-3 font-mono font-bold text-slate-800">{bill.invoiceNumber || bill.id}</td>
                        <td className="py-3 px-3 text-slate-500">{bill.billDate}</td>
                        <td className="py-3 px-3 font-bold text-slate-900">{bill.vendorName}</td>
                        <td className="py-3 px-3">
                          <span className="px-2 py-0.5 bg-slate-100 rounded text-[10px] font-semibold text-slate-700">
                            {bill.category}
                          </span>
                        </td>
                        <td className="py-3 px-3 font-mono text-slate-600">{bill.paymentMethod}</td>
                        <td className="py-3 px-3 text-right font-mono font-black text-slate-900">
                          {settings.currency} {Number(bill.amount).toFixed(2)}
                        </td>
                        <td className="py-3 px-3 text-center">
                          <span className={`px-2 py-0.5 rounded-full text-[10px] font-bold ${
                            bill.paymentStatus === 'PAID'
                              ? 'bg-emerald-100 text-emerald-800'
                              : 'bg-amber-100 text-amber-800'
                          }`}>
                            {bill.paymentStatus}
                          </span>
                        </td>
                        <td className="py-3 px-3 text-right">
                          <button
                            type="button"
                            onClick={() => {
                              if (window.confirm(`Delete invoice ${bill.invoiceNumber || bill.id} from ${bill.vendorName}?`)) {
                                setVendorBills(prev => prev.filter(b => b.id !== bill.id));
                                recordAuditLog('VENDOR_BILL_DELETED', bill.id, `Deleted invoice ${bill.invoiceNumber} for ${settings.currency} ${bill.amount}`);
                              }
                            }}
                            className="p-1.5 text-slate-400 hover:text-rose-600 hover:bg-rose-50 rounded-lg cursor-pointer transition-colors"
                            title="Delete Invoice"
                          >
                            <Trash2 className="h-4 w-4" />
                          </button>
                        </td>
                      </tr>
                    ))
                  )}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {/* VIEW: COMPREHENSIVE ACCOUNTING & P&L ANALYTICS */}
        {activeTab === 'accounting' && (
          <div className="flex-1 overflow-y-auto p-4 sm:p-6 space-y-6">
            {/* Header & Date Range Filter */}
            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4">
              <div>
                <h2 className="text-xl font-black text-slate-900 flex items-center gap-2">
                  <PieChart className="h-5 w-5 text-[#ff5500]" />
                  Financial Accounting, P&amp;L &amp; Expense Analytics
                </h2>
                <p className="text-xs text-slate-500 mt-0.5">
                  Track revenue vs operational expenditures, cash-out outflow categories, cost of inventory, and net profits.
                </p>
              </div>

              {/* Filter Buttons & Cloud Archive Action */}
              <div className="flex flex-wrap items-center gap-2">
                <div className="flex items-center gap-1.5 bg-white border border-slate-200 p-1 rounded-xl shadow-xs text-xs font-bold">
                  {[
                    { id: 'ALL', label: 'All Time' },
                    { id: 'TODAY', label: 'Today' },
                    { id: 'THIS_MONTH', label: 'This Month' }
                  ].map(f => (
                    <button
                      key={f.id}
                      type="button"
                      onClick={() => setAccountingPeriod(f.id)}
                      className={`px-3 py-1.5 rounded-lg cursor-pointer transition-colors ${
                        accountingPeriod === f.id
                          ? 'bg-[#ff5500] text-white shadow-xs'
                          : 'text-slate-600 hover:text-slate-900 hover:bg-slate-50'
                      }`}
                    >
                      {f.label}
                    </button>
                  ))}
                </div>

                {/* Lock & Push Financial Snapshot to Firebase */}
                <button
                  type="button"
                  onClick={async () => {
                    const todayStr = getLocalDateStr();
                    const currentMonthStr = todayStr.slice(0, 7);

                    const snapshotRecord = {
                      id: `ACC-SNAP-${Date.now().toString().slice(-6)}`,
                      period: accountingPeriod,
                      timestamp: new Date().toLocaleString(),
                      savedBy: currentUser.name,
                      financials: {
                        netSalesRevenue,
                        grossSalesRevenue,
                        totalDiscountsGiven,
                        totalBOMCostOfGoodsSold,
                        grossProfit,
                        grossMarginPercent: `${grossMarginPercent}%`,
                        totalOperationalCashOut,
                        totalSafeDropBanking,
                        totalVendorBillsPaid,
                        totalPayrollDisbursed,
                        totalEmployerEpfEtf,
                        totalOperatingExpenses,
                        netProfit,
                        profitMargin: `${netProfitMargin}%`,
                        expenseCategories
                      }
                    };

                    // 1. Sync live state to Firebase
                    await syncToCloud('latest_accounting_summary', snapshotRecord);

                    // 2. Write permanent historical copy to Firebase pos_archives
                    if (typeof appendCloudArchive === 'function') {
                      await appendCloudArchive('accounting_periods', snapshotRecord);
                    }

                    recordAuditLog(
                      'ACCOUNTING_SNAPSHOT_SAVED',
                      snapshotRecord.id,
                      `Saved financial snapshot for ${accountingPeriod}. Net Revenue: ${settings.currency} ${netSalesRevenue.toFixed(2)}, Net Profit: ${settings.currency} ${netProfit.toFixed(2)}`
                    );

                    setSettingsNotice({
                      title: 'Accounting Snapshot Synced',
                      detail: `P&L statement for ${accountingPeriod} uploaded and permanently archived to Firebase.`
                    });
                    setTimeout(() => setSettingsNotice(null), 3500);
                  }}
                  className="px-3.5 py-2 bg-slate-900 hover:bg-black text-white font-bold rounded-xl text-xs flex items-center gap-1.5 shadow-xs cursor-pointer transition-all active:scale-95"
                  title="Upload financial snapshot to Firebase"
                >
                  <Upload className="h-3.5 w-3.5 text-[#ff5500]" />
                  <span>Sync to Cloud</span>
                </button>
              </div>
            </div>

           {/* Calculations Engine (Mapped to transactions, BOM COGS, expenses, vendorBills & payroll) */}
            {(() => {
              const todayStr = getLocalDateStr();
              const currentMonthStr = todayStr.slice(0, 7); // 'YYYY-MM'

              // 1. Data Sources
              const salesData = Array.isArray(transactions) ? transactions : [];
              const expenseData = Array.isArray(expenses) ? expenses : (currentShift?.payouts || []);
              const externalBillsData = Array.isArray(vendorBills) ? vendorBills : [];
              const payrollData = Array.isArray(payrollRecords) ? payrollRecords : [];

              // 2. Filter Sales / Revenue
              const filteredSales = salesData.filter(inv => {
                if (!inv) return false;
                const dStr = extractDateStr(inv.date) || String(inv.date || '');
                if (accountingPeriod === 'TODAY') return dStr.startsWith(todayStr);
                if (accountingPeriod === 'THIS_MONTH') return dStr.startsWith(currentMonthStr);
                return true;
              });

              // 3. Filter All Drawer Cash-Out Disbursements
              const filteredDisbursements = expenseData.filter(v => {
                if (!v || v.status === 'REJECTED') return false;
                const dStr = extractDateStr(v.date || v.createdAt) || String(v.date || '');
                if (accountingPeriod === 'TODAY') return dStr.startsWith(todayStr);
                if (accountingPeriod === 'THIS_MONTH') return dStr.startsWith(currentMonthStr);
                return true;
              });

              // Helper: Distinguish internal banking transfers from actual operating expenses
              const isBankingTransfer = (item) => {
                const category = String(item.category || '').toLowerCase();
                const reason = String(item.reason || '').toLowerCase();
                return (
                  category.includes('safe drop') ||
                  category.includes('bank deposit') ||
                  category.includes('banking') ||
                  reason.includes('safe drop') ||
                  reason.includes('bank deposit')
                );
              };

              // True drawer operational expenditures (vendor cash payouts, petty cash, supplies)
              const operationalExpenses = filteredDisbursements.filter(v => !isBankingTransfer(v));

              // Internal drawer-to-safe / banking transfers (Non-OPEX asset movements)
              const bankingTransfers = filteredDisbursements.filter(v => isBankingTransfer(v));

              // 4. Filter External Vendor Bills (Non-Drawer Invoices: Bank transfer, Cheque, Card)
              const filteredVendorBills = externalBillsData.filter(b => {
                if (!b || b.paymentStatus !== 'PAID') return false;
                const bDate = extractDateStr(b.billDate || b.date) || String(b.billDate || '');
                if (accountingPeriod === 'TODAY') return bDate.startsWith(todayStr);
                if (accountingPeriod === 'THIS_MONTH') return bDate.startsWith(currentMonthStr);
                return true;
              });

              // 5. Filter Payroll Disbursed
              const filteredPayroll = payrollData.filter(p => {
                if (!p) return false;
                if (accountingPeriod === 'THIS_MONTH') return p.period === currentMonthStr;
                if (accountingPeriod === 'TODAY') return p.processedAt && String(p.processedAt).startsWith(todayStr);
                return true;
              });

              // --- REVENUE METRICS ---
              const grossSalesRevenue = filteredSales.reduce((acc, inv) => acc + (Number(inv.subtotal) || 0), 0);
              const totalTaxCollected = filteredSales.reduce((acc, inv) => acc + (Number(inv.tax) || 0), 0);
              const totalServiceCharge = filteredSales.reduce((acc, inv) => acc + (Number(inv.serviceCharge) || 0), 0);
              const totalDiscountsGiven = filteredSales.reduce((acc, inv) => acc + (Number(inv.discount) || 0), 0);
              const calculatedNet = grossSalesRevenue - totalDiscountsGiven;
              const netSalesRevenue = calculatedNet > 0 ? calculatedNet : filteredSales.reduce((acc, inv) => acc + (Number(inv.total) || 0), 0);

              // --- COST OF GOODS SOLD (BOM INGREDIENT COST TRACKING) ---
              let totalBOMCostOfGoodsSold = 0;
              filteredSales.forEach(inv => {
                // If invoice already has cogs computed at settlement time, use it
                if (Number(inv.cogs) > 0) {
                  totalBOMCostOfGoodsSold += Number(inv.cogs);
                } else if (Array.isArray(inv.items)) {
                  // Fallback: Recompute line-by-line from recipes (counts 0 if not configured)
                  inv.items.forEach(item => {
                    const dish = menuItems.find(m => m.id === item.id || m.name === item.name);
                    const qty = Number(item.qty) || 1;
                    if (dish && Array.isArray(dish.recipe) && dish.recipe.length > 0) {
                      let itemUnitCogs = 0;
                      dish.recipe.forEach(r => {
                        const ing = inventoryMap[r.ingredientId];
                        if (ing && Number(ing.cost) > 0) {
                          itemUnitCogs += (Number(ing.cost) * (Number(r.amount) || 0));
                        }
                      });
                      totalBOMCostOfGoodsSold += (itemUnitCogs * qty);
                    }
                  });
                }
              });

              // Gross Profit after raw food/beverage ingredient cost
              const grossProfit = netSalesRevenue - totalBOMCostOfGoodsSold;
              const grossMarginPercent = netSalesRevenue > 0 ? ((grossProfit / netSalesRevenue) * 100).toFixed(1) : 0;

              // --- EXPENSE CATEGORIZATION ---
              const expenseCategories = {};
              operationalExpenses.forEach(exp => {
                const cat = (exp.category || exp.reason || 'General Purchases').trim();
                const amt = Number(exp.amount) || 0;
                expenseCategories[cat] = (expenseCategories[cat] || 0) + amt;
              });

              filteredVendorBills.forEach(b => {
                const cat = (b.category || 'Vendor Invoices').trim();
                const amt = Number(b.amount) || 0;
                expenseCategories[cat] = (expenseCategories[cat] || 0) + amt;
              });

              // --- SUBTOTALS ---
              const totalOperationalCashOut = operationalExpenses.reduce((acc, exp) => acc + (Number(exp.amount) || 0), 0);
              const totalSafeDropBanking = bankingTransfers.reduce((acc, exp) => acc + (Number(exp.amount) || 0), 0);
              const totalDrawerCashOutflow = totalOperationalCashOut + totalSafeDropBanking;
              const totalVendorBillsPaid = filteredVendorBills.reduce((acc, b) => acc + (Number(b.amount) || 0), 0);

              // --- PAYROLL TOTALS ---
              const totalPayrollDisbursed = filteredPayroll.reduce((acc, p) => acc + (Number(p.breakdown?.netSalary) || Number(p.netPay) || 0), 0);
              const totalEmployerEpfEtf = filteredPayroll.reduce((acc, p) => acc + (Number(p.breakdown?.epfEmployer) || 0) + (Number(p.breakdown?.etfEmployer) || 0), 0);

              // --- TOTAL OPEX (Operating Expenses) ---
              const totalOperatingExpenses = totalOperationalCashOut + totalVendorBillsPaid + totalPayrollDisbursed + totalEmployerEpfEtf;

              // --- NET OPERATING PROFIT (Net Revenue - BOM COGS - OPEX) ---
              const netProfit = grossProfit - totalOperatingExpenses;
              const profitMargin = netSalesRevenue > 0 ? ((netProfit / netSalesRevenue) * 100).toFixed(1) : 0;

              // --- TOP 10 ITEMS SOLD ---
              const itemSalesMap = {};
              filteredSales.forEach(inv => {
                (inv.items || []).forEach(item => {
                  if (!item || !item.name) return;
                  if (!itemSalesMap[item.name]) {
                    itemSalesMap[item.name] = { name: item.name, qty: 0, revenue: 0 };
                  }
                  const qty = Number(item.qty) || 1;
                  const price = Number(item.price) || 0;
                  itemSalesMap[item.name].qty += qty;
                  itemSalesMap[item.name].revenue += (price * qty);
                });
              });
              const topSoldItems = Object.values(itemSalesMap).sort((a, b) => b.qty - a.qty).slice(0, 10);
              const maxItemQty = topSoldItems.length > 0 ? Math.max(...topSoldItems.map(i => i.qty)) : 1;

              return (
                <div className="space-y-6">
                  {/* Top Key Performance Cards */}
                  <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-5 gap-3.5">
                    {/* Gross Revenue */}
                    <div className="bg-white rounded-2xl border border-slate-200 p-4 shadow-xs">
                      <div className="flex items-center justify-between">
                        <span className="text-[10px] font-black uppercase text-slate-400">Total Net Revenue</span>
                        <span className="p-1.5 bg-emerald-50 text-emerald-600 rounded-lg">
                          <TrendingUp className="h-4 w-4" />
                        </span>
                      </div>
                      <p className="text-xl font-black font-mono text-slate-900 mt-1">
                        {settings.currency} {netSalesRevenue.toFixed(2)}
                      </p>
                      <p className="text-[10px] text-slate-400 mt-0.5">
                        Gross {settings.currency} {grossSalesRevenue.toFixed(0)} - Disc {settings.currency} {totalDiscountsGiven.toFixed(0)}
                      </p>
                    </div>

                    {/* BOM Cost of Goods Sold (COGS) */}
                    <div className="bg-white rounded-2xl border border-slate-200 p-4 shadow-xs">
                      <div className="flex items-center justify-between">
                        <span className="text-[10px] font-black uppercase text-slate-400">Cost of Goods (COGS)</span>
                        <span className="p-1.5 bg-amber-50 text-amber-600 rounded-lg">
                          <BookOpen className="h-4 w-4" />
                        </span>
                      </div>
                      <p className="text-xl font-black font-mono text-amber-700 mt-1">
                        {settings.currency} {totalBOMCostOfGoodsSold.toFixed(2)}
                      </p>
                      <p className="text-[10px] text-slate-400 mt-0.5">
                        BOM Recipes • Gross Margin {grossMarginPercent}%
                      </p>
                    </div>

                    {/* Operational Cash Out (Excludes Safe Drops) */}
                    <div className="bg-white rounded-2xl border border-slate-200 p-4 shadow-xs">
                      <div className="flex items-center justify-between">
                        <span className="text-[10px] font-black uppercase text-slate-400">Total OPEX</span>
                        <span className="p-1.5 bg-rose-50 text-rose-600 rounded-lg">
                          <TrendingDown className="h-4 w-4" />
                        </span>
                      </div>
                      <p className="text-xl font-black font-mono text-rose-600 mt-1">
                        {settings.currency} {totalOperatingExpenses.toFixed(2)}
                      </p>
                      <p className="text-[10px] text-slate-400 mt-0.5">
                        Cashouts + Bills + Wages (Safe Drops Excluded)
                      </p>
                    </div>

                    {/* Total Wages & Statutory */}
                    <div className="bg-white rounded-2xl border border-slate-200 p-4 shadow-xs">
                      <div className="flex items-center justify-between">
                        <span className="text-[10px] font-black uppercase text-slate-400">Salaries &amp; EPF/ETF</span>
                        <span className="p-1.5 bg-indigo-50 text-indigo-600 rounded-lg">
                          <Briefcase className="h-4 w-4" />
                        </span>
                      </div>
                      <p className="text-xl font-black font-mono text-indigo-600 mt-1">
                        {settings.currency} {(totalPayrollDisbursed + totalEmployerEpfEtf).toFixed(2)}
                      </p>
                      <p className="text-[10px] text-slate-400 mt-0.5">
                        Take-home + {settings.currency} {totalEmployerEpfEtf.toFixed(2)} EPF/ETF
                      </p>
                    </div>

                    {/* Net Profit */}
                    <div className={`rounded-2xl border p-4 shadow-xs ${
                      netProfit >= 0 ? 'bg-emerald-950 text-white border-emerald-900' : 'bg-rose-950 text-white border-rose-900'
                    }`}>
                      <div className="flex items-center justify-between">
                        <span className="text-[10px] font-black uppercase tracking-wider text-emerald-300">
                          Net Operating Profit
                        </span>
                        <span className="px-2 py-0.5 rounded-full text-[9px] font-black font-mono bg-white/10">
                          {profitMargin}% Margin
                        </span>
                      </div>
                      <p className="text-xl font-black font-mono mt-1">
                        {settings.currency} {netProfit.toFixed(2)}
                      </p>
                      <p className="text-[10px] opacity-80 mt-0.5">
                        {netProfit >= 0 ? 'Profitable operation' : 'Operating at a loss'}
                      </p>
                    </div>
                  </div>

                  {/* Main Grid: Comprehensive P&L Statement and Expense Category Breakdown */}
                  <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
                    {/* PROFIT & LOSS STATEMENT */}
                    <div className="bg-white rounded-2xl border border-slate-200 p-5 shadow-xs flex flex-col justify-between">
                      <div>
                        <div className="flex items-center justify-between pb-3 border-b border-slate-100">
                          <h3 className="text-xs font-black uppercase tracking-wider text-slate-900 flex items-center gap-1.5">
                            <Receipt className="h-4 w-4 text-[#ff5500]" />
                            Official Profit &amp; Loss Statement
                          </h3>
                          <span className="text-[11px] font-mono font-bold text-slate-400">
                            {accountingPeriod}
                          </span>
                        </div>

                        <div className="divide-y divide-slate-100 text-xs py-2 space-y-2">
                          {/* REVENUE SECTION */}
                          <div className="pt-2">
                            <span className="text-[10px] font-black uppercase text-emerald-600 block mb-1">
                              1. Operating Revenue &amp; Gross Margin
                            </span>
                            <div className="flex justify-between py-1 text-slate-700">
                              <span>Gross Menu &amp; Bar Invoiced Sales</span>
                              <span className="font-mono font-bold">{settings.currency} {grossSalesRevenue.toFixed(2)}</span>
                            </div>
                            <div className="flex justify-between py-1 text-rose-600">
                              <span>Less: Promotional Discounts &amp; Vouchers</span>
                              <span className="font-mono font-bold">-{settings.currency} {totalDiscountsGiven.toFixed(2)}</span>
                            </div>
                            <div className="flex justify-between py-1 text-slate-900 font-bold bg-slate-50 px-2 rounded-lg">
                              <span>Net Sales Revenue</span>
                              <span className="font-mono">{settings.currency} {netSalesRevenue.toFixed(2)}</span>
                            </div>
                            <div className="flex justify-between py-1 text-amber-700 px-2">
                              <span>Less: Cost of Goods Sold (BOM Recipe Ingredients)</span>
                              <span className="font-mono font-bold">-{settings.currency} {totalBOMCostOfGoodsSold.toFixed(2)}</span>
                            </div>
                            <div className="flex justify-between py-1.5 font-black text-slate-950 bg-emerald-50 px-2 rounded-lg mt-1 border border-emerald-200">
                              <span>Gross Operating Profit</span>
                              <span className="font-mono text-emerald-800">{settings.currency} {grossProfit.toFixed(2)} ({grossMarginPercent}%)</span>
                            </div>
                          </div>

                          {/* OPERATING EXPENDITURES (Excludes Safe Drop) */}
                          <div className="pt-3">
                            <span className="text-[10px] font-black uppercase text-rose-600 block mb-1">
                              2. Operating Expenditures (OPEX)
                            </span>
                            <div className="flex justify-between py-1 text-slate-700">
                              <span>Cash Out Purchases &amp; Petty Cash</span>
                              <span className="font-mono">{settings.currency} {totalOperationalCashOut.toFixed(2)}</span>
                            </div>
                            <div className="flex justify-between py-1 text-slate-700">
                              <span>External Vendor Bills &amp; Banked Invoices</span>
                              <span className="font-mono">{settings.currency} {totalVendorBillsPaid.toFixed(2)}</span>
                            </div>
                            <div className="flex justify-between py-1 text-slate-700">
                              <span>Staff Net Salaries &amp; Wage Disbursed</span>
                              <span className="font-mono">{settings.currency} {totalPayrollDisbursed.toFixed(2)}</span>
                            </div>
                            <div className="flex justify-between py-1 text-slate-700">
                              <span>Employer EPF (12%) &amp; ETF (3%) Remittance</span>
                              <span className="font-mono">{settings.currency} {totalEmployerEpfEtf.toFixed(2)}</span>
                            </div>
                            <div className="flex justify-between py-1.5 font-black text-rose-700 bg-rose-50 px-2 rounded-lg mt-1">
                              <span>Total Operating Expenses</span>
                              <span className="font-mono">-{settings.currency} {totalOperatingExpenses.toFixed(2)}</span>
                            </div>
                          </div>

                          {/* INTERNAL ASSET TRANSFERS (Safe Drop / Bank Deposit) */}
                          <div className="pt-3">
                            <span className="text-[10px] font-black uppercase text-indigo-600 block mb-1">
                              3. Internal Asset Transfers &amp; Banking (Non-Expense)
                            </span>
                            <div className="flex justify-between py-1 text-slate-600">
                              <span>Safe Drops &amp; Bank Deposits</span>
                              <span className="font-mono font-bold text-indigo-700">+{settings.currency} {totalSafeDropBanking.toFixed(2)}</span>
                            </div>
                          </div>

                          {/* TAX & GRATUITY ESCROW */}
                          <div className="pt-3">
                            <span className="text-[10px] font-black uppercase text-slate-500 block mb-1">
                              4. Escrow Fiduciary Liabilities (Collected &amp; Held)
                            </span>
                            <div className="flex justify-between py-1 text-slate-500">
                              <span>Government Tax / VAT Invoiced</span>
                              <span className="font-mono">{settings.currency} {totalTaxCollected.toFixed(2)}</span>
                            </div>
                            <div className="flex justify-between py-1 text-slate-500">
                              <span>Service Charge Pool Collected</span>
                              <span className="font-mono">{settings.currency} {totalServiceCharge.toFixed(2)}</span>
                            </div>
                          </div>
                        </div>
                      </div>

                      {/* BOTTOM NET PROFIT TOTAL */}
                      <div className="pt-3 border-t border-slate-200 mt-2">
                        <div className="flex justify-between items-center text-sm font-black p-3 bg-slate-900 text-white rounded-xl">
                          <span>NET OPERATING SURPLUS / (DEFICIT)</span>
                          <span className={`font-mono text-base ${netProfit >= 0 ? 'text-emerald-400' : 'text-rose-400'}`}>
                            {settings.currency} {netProfit.toFixed(2)}
                          </span>
                        </div>
                      </div>
                    </div>

                    {/* EXPENSE CATEGORY BREAKDOWN CHART */}
                    <div className="bg-white rounded-2xl border border-slate-200 p-5 shadow-xs flex flex-col justify-between">
                      <div>
                        <div className="flex items-center justify-between pb-3 border-b border-slate-100">
                          <h3 className="text-xs font-black uppercase tracking-wider text-slate-900 flex items-center gap-1.5">
                            <Coins className="h-4 w-4 text-[#ff5500]" />
                            Operational Spending Categories
                          </h3>
                          <span className="text-[11px] font-mono text-slate-500 font-bold">
                            {Object.keys(expenseCategories).length} Categories
                          </span>
                        </div>

                        {/* Visual Bar Breakdown (Operational Expenses Only) */}
                        <div className="py-3 space-y-3.5">
                          {Object.keys(expenseCategories).length === 0 ? (
                            <p className="text-xs text-slate-400 italic py-8 text-center">
                              No operational expenditures recorded for this period.
                            </p>
                          ) : (
                            Object.entries(expenseCategories)
                              .sort((a, b) => b[1] - a[1])
                              .map(([category, amount]) => {
                                const percentage = totalOperationalCashOut > 0 
                                  ? ((amount / totalOperationalCashOut) * 100).toFixed(1) 
                                  : 0;

                                return (
                                  <div key={category} className="space-y-1">
                                    <div className="flex justify-between text-xs font-bold">
                                      <span className="text-slate-800">{category}</span>
                                      <span className="font-mono text-slate-900">
                                        {settings.currency} {amount.toFixed(2)} <span className="text-slate-400 text-[10px]">({percentage}%)</span>
                                      </span>
                                    </div>
                                    <div className="w-full bg-slate-100 h-2.5 rounded-full overflow-hidden">
                                      <div
                                        className="bg-orange-500 h-full rounded-full transition-all duration-500"
                                        style={{ width: `${percentage}%` }}
                                      />
                                    </div>
                                  </div>
                                );
                              })
                          )}
                        </div>
                      </div>

                      {/* Outflow Breakdown Footer */}
                      <div className="pt-3 border-t border-slate-100 space-y-1.5 text-xs">
                        <div className="flex justify-between items-center text-slate-600">
                          <span>Operational Expenses (OPEX):</span>
                          <span className="font-mono font-bold text-rose-600">
                            {settings.currency} {totalOperationalCashOut.toFixed(2)}
                          </span>
                        </div>
                        <div className="flex justify-between items-center text-slate-600">
                          <span>Bank &amp; Safe Transfers:</span>
                          <span className="font-mono font-bold text-indigo-600">
                            {settings.currency} {totalSafeDropBanking.toFixed(2)}
                          </span>
                        </div>
                        <div className="flex justify-between items-center pt-1.5 border-t border-slate-100 font-bold text-slate-800">
                          <span>Total Drawer Outflow:</span>
                          <span className="font-mono font-black text-slate-900 text-sm">
                            {settings.currency} {totalDrawerCashOutflow.toFixed(2)}
                          </span>
                        </div>
                      </div>
                    </div>
                  </div>

                  {/* TOP 10 ITEMS SOLD ANALYTICS CHART */}
                  <div className="bg-white rounded-2xl border border-slate-200 p-5 shadow-xs">
                    <div className="flex items-center justify-between pb-3 border-b border-slate-100">
                      <div>
                        <h3 className="text-xs font-black uppercase tracking-wider text-slate-900 flex items-center gap-1.5">
                          <BarChart3 className="h-4 w-4 text-[#ff5500]" />
                          Top 10 High Volume Menu Items Sold
                        </h3>
                        <p className="text-[10px] text-slate-400">Item units dispensed vs total revenue generated</p>
                      </div>
                      <span className="text-xs font-mono font-bold text-slate-600">
                        {topSoldItems.length} Products
                      </span>
                    </div>

                    <div className="mt-4 space-y-3">
                      {topSoldItems.length === 0 ? (
                        <p className="text-xs text-slate-400 italic py-8 text-center">
                          No settled menu items found in the selected period.
                        </p>
                      ) : (
                        topSoldItems.map((item, index) => {
                          const barWidth = Math.max(5, (item.qty / maxItemQty) * 100);

                          return (
                            <div key={item.name} className="flex flex-col sm:flex-row sm:items-center gap-2 text-xs">
                              {/* Item rank and title */}
                              <div className="w-48 shrink-0 flex items-center gap-2">
                                <span className="h-5 w-5 rounded bg-slate-900 text-white text-[10px] font-black flex items-center justify-center">
                                  {index + 1}
                                </span>
                                <span className="font-bold text-slate-800 truncate" title={item.name}>
                                  {item.name}
                                </span>
                              </div>

                              {/* Graphical Bar */}
                              <div className="flex-1 bg-slate-100 h-6 rounded-lg overflow-hidden relative flex items-center px-2">
                                <div
                                  className="absolute left-0 top-0 bottom-0 bg-gradient-to-r from-orange-400 to-[#ff5500] rounded-lg opacity-85 transition-all duration-500"
                                  style={{ width: `${barWidth}%` }}
                                />
                                <span className="relative z-10 text-[10px] font-mono font-black text-white drop-shadow-xs">
                                  {item.qty} units
                                </span>
                              </div>

                              {/* Invoiced Revenue */}
                              <div className="w-28 text-right font-mono font-bold text-slate-900">
                                {settings.currency} {item.revenue.toFixed(2)}
                              </div>
                            </div>
                          );
                        })
                      )}
                    </div>
                  </div>
                </div>
              );
            })()}
          </div>
        )}

        {/* VIEW: COMPREHENSIVE EMPLOYMENT, TIME CLOCK & SRI LANKAN PAYROLL */}
        {activeTab === 'payroll' && (
          <div className="flex-1 overflow-y-auto p-4 sm:p-6 space-y-6">
            {/* Header and Action Bar */}
            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4">
              <div>
                <h2 className="text-xl font-black text-slate-900 flex items-center gap-2">
                  <Briefcase className="h-5 w-5 text-[#ff5500]" />
                  Employment, Attendance &amp; Sri Lanka Statutory Payroll
                </h2>
                <p className="text-xs text-slate-500 mt-0.5">
                  Daily Clock-In/Out biometric time sheet, EPF (8%/12%), ETF (3%), Service Pool bonus, and official pay slips.
                </p>
              </div>

              <div className="flex flex-wrap items-center gap-2">
                {/* 1. ISSUE ADVANCE BUTTON */}
                <button
                  type="button"
                  onClick={() => {
                    const firstStaff = staffList[0];
                    setAdvanceForm({
                      staffId: firstStaff ? firstStaff.id : '',
                      period: `${new Date().getFullYear()}-${String(new Date().getMonth() + 1).padStart(2, '0')}`,
                      amount: '',
                      paymentMethod: 'CASH',
                      reason: 'Emergency advance on salary',
                      notes: ''
                    });
                    setIssueAdvanceModalOpen(true);
                  }}
                  className="px-3.5 py-2 bg-amber-500 hover:bg-amber-600 text-white font-bold rounded-xl text-xs flex items-center gap-1.5 shadow-xs cursor-pointer transition-all active:scale-95"
                  title="Disburse cash or bank advance to employee"
                >
                  <Banknote className="h-4 w-4" />
                  <span>Issue Advance</span>
                </button>

                {/* 2. EXPORT ARCHIVE BUTTON */}
                <button
                  type="button"
                  onClick={() => {
                    const exportPayload = {
                      exportedAt: new Date().toISOString(),
                      restaurant: settings.restaurantName || 'Restaurant POS',
                      attendanceLogs,
                      payrollRecords,
                      salaryAdvances
                    };
                    const dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(exportPayload, null, 2));
                    const downloadAnchor = document.createElement('a');
                    downloadAnchor.setAttribute("href", dataStr);
                    downloadAnchor.setAttribute("download", `payroll_archive_${getLocalDateStr()}.json`);
                    document.body.appendChild(downloadAnchor);
                    downloadAnchor.click();
                    downloadAnchor.remove();
                  }}
                  className="px-3.5 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold rounded-xl text-xs flex items-center gap-1.5 cursor-pointer transition-colors shadow-xs"
                >
                  <Download className="h-4 w-4 text-slate-600" />
                  <span>Export Archive</span>
                </button>

                {/* 3. PROCESS PAY SLIP BUTTON */}
                <button
                  type="button"
                  onClick={() => {
                    const firstStaff = staffList[0];
                    const curPeriod = `${new Date().getFullYear()}-${String(new Date().getMonth() + 1).padStart(2, '0')}`;
                    const advancesForStaff = (salaryAdvances || [])
                      .filter(a => a.staffId === firstStaff?.id && a.period === curPeriod && a.status !== 'REJECTED')
                      .reduce((acc, a) => acc + (Number(a.amount) || 0), 0);

                    setEditingPayrollId(null);
                    setPayrollInputForm({
                      staffId: firstStaff ? firstStaff.id : '',
                      period: curPeriod,
                      epfEtfEnabled: firstStaff ? firstStaff.epfEtfEnabled !== false : true,
                      basicSalary: firstStaff?.basicSalary ?? 35000,
                      budgetaryAllowance: firstStaff?.budgetaryAllowance ?? 2500,
                      otherAllowances: firstStaff?.otherAllowances ?? 0,
                      serviceChargeBonus: Math.round((salesMetrics?.serviceCharge || 0) / Math.max(1, staffList.length)),
                      incentiveBonus: firstStaff?.fixedBonus ?? 0,
                      overtimeHours: 0,
                      overtimeRate: firstStaff?.overtimeRate ?? 250,
                      salaryAdvance: advancesForStaff,
                      otherDeductions: 0,
                      standardWorkingDays: 26,
                      workedDays: 26,
                      paidLeaves: 0,
                      unpaidLeaves: 0,
                      holidaysCount: 4,
                      shortShiftsCount: 0,
                      notes: ''
                    });
                    setProcessPayModalOpen(true);
                  }}
                  className="px-4 py-2 bg-[#ff5500] hover:bg-orange-600 text-white font-bold rounded-xl text-xs flex items-center gap-2 shadow-xs cursor-pointer transition-all active:scale-95"
                >
                  <Plus className="h-4 w-4" />
                  <span>Process Pay Slip</span>
                </button>
              </div>
            </div>

            {/* Navigation Sub-Tabs */}
            <div className="flex items-center gap-4 border-b border-slate-200 pb-2 text-xs font-bold overflow-x-auto">
              {[
                { id: 'attendance', label: 'Time Sheet & Clock In/Out' },
                { id: 'payslips', label: 'Payroll & Pay Slips' },
                { id: 'epf_etf', label: 'Sri Lanka EPF / ETF Return (Form C)' },
                { id: 'profiles', label: 'Employee Registry & Wages' }
              ].map(tab => (
                <button
                  key={tab.id}
                  type="button"
                  onClick={() => setPayrollSubTab(tab.id)}
                  className={`pb-2 transition-all relative whitespace-nowrap cursor-pointer ${
                    payrollSubTab === tab.id
                      ? 'text-[#ff5500] after:absolute after:bottom-0 after:left-0 after:w-full after:h-0.5 after:bg-[#ff5500]'
                      : 'text-slate-500 hover:text-slate-900'
                  }`}
                >
                  {tab.label}
                </button>
              ))}
            </div>

            {/* SUB-TAB 1: TIME SHEET & CLOCK IN / CLOCK OUT */}
            {payrollSubTab === 'attendance' && (
              <div className="space-y-6">
                {/* Employee Cards Grid */}
                <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
                  {staffList.map(member => {
                    const { clockedIn, log } = getStaffClockStatus(member.id);

                    return (
                      <div key={member.id} className="bg-white rounded-2xl border border-slate-200 p-4 shadow-xs flex flex-col justify-between">
                        <div className="flex items-start justify-between">
                          <div className="flex items-center gap-2.5">
                            <div className="h-10 w-10 rounded-xl bg-slate-900 text-white font-black text-xs flex items-center justify-center">
                              {member.avatar}
                            </div>
                            <div>
                              <h4 className="font-extrabold text-xs text-slate-900 leading-tight">{member.name}</h4>
                              <span className="text-[10px] text-slate-500 block">{member.role}</span>
                            </div>
                          </div>

                          <span className={`px-2 py-0.5 rounded-full text-[9px] font-black uppercase ${
                            clockedIn ? 'bg-emerald-100 text-emerald-800 animate-pulse' : 'bg-slate-100 text-slate-500'
                          }`}>
                            {clockedIn ? '● On Duty' : 'Off Duty'}
                          </span>
                        </div>

                        <div className="my-3 py-2 border-y border-slate-100 text-[11px] text-slate-600 space-y-1">
                          <div className="flex justify-between">
                            <span>Today Clock In:</span>
                            <span className="font-mono font-bold">{log ? log.clockInTime : '--:--'}</span>
                          </div>
                          <div className="flex justify-between">
                            <span>Today Clock Out:</span>
                            <span className="font-mono font-bold">{log && log.clockOutTime ? log.clockOutTime : '--:--'}</span>
                          </div>
                        </div>

                        {/* Clock In / Out Button with Permanent Local & Cloud Archiving */}
                        <button
                          type="button"
                          onClick={() => {
                            const now = new Date();
                            const nowTime = now.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit', second: '2-digit' });
                            const todayStr = getLocalDateStr();
                            const nowIso = now.toISOString();

                            if (clockedIn && log) {
                              // Clock Out - Update active session and archive
                              const start = new Date(`${todayStr} ${log.clockInTime}`);
                              const diffHours = Math.max(0.1, Number(((now - start) / (1000 * 60 * 60)).toFixed(2)));

                              const updatedLog = {
                                ...log,
                                clockOutTime: nowTime,
                                clockOutISO: nowIso,
                                totalHours: diffHours,
                                isOvertime: diffHours > 8,
                                status: 'COMPLETED'
                              };

                              setAttendanceLogs(prev => prev.map(a => a.id === log.id ? updatedLog : a));

                              // Append permanent completed session to cloud archive
                              if (typeof appendCloudArchive === 'function') {
                                appendCloudArchive('attendance_logs', updatedLog);
                              }

                              recordAuditLog(
                                'STAFF_CLOCK_OUT_PERMANENT',
                                member.id,
                                `${member.name} clocked out at ${nowTime} (${diffHours} hrs). Log ID: ${log.id}`
                              );
                            } else {
                              // Clock In - Generate unique permanent entry
                              const newLog = {
                                id: `ATT-${todayStr.replace(/-/g, '')}-${member.id}-${Date.now().toString().slice(-4)}`,
                                date: todayStr,
                                staffId: member.id,
                                staffName: member.name,
                                role: member.role,
                                clockInTime: nowTime,
                                clockInISO: nowIso,
                                clockOutTime: null,
                                clockOutISO: null,
                                totalHours: 0,
                                isOvertime: false,
                                status: 'ON_DUTY',
                                createdByDevice: typeof navigator !== 'undefined' ? navigator.userAgent : 'POS Terminal'
                              };

                              setAttendanceLogs(prev => [newLog, ...prev]);

                              // Append permanent clock-in event to cloud archive
                              if (typeof appendCloudArchive === 'function') {
                                appendCloudArchive('attendance_logs', newLog);
                              }

                              recordAuditLog(
                                'STAFF_CLOCK_IN_PERMANENT',
                                member.id,
                                `${member.name} clocked in at ${nowTime}. Log ID: ${newLog.id}`
                              );
                            }
                          }}
                          className={`w-full py-2.5 rounded-xl text-xs font-bold transition-all cursor-pointer flex items-center justify-center gap-1.5 shadow-xs ${
                            clockedIn
                              ? 'bg-rose-50 hover:bg-rose-100 text-rose-700 border border-rose-200'
                              : 'bg-emerald-600 hover:bg-emerald-700 text-white'
                          }`}
                        >
                          <Clock className="h-3.5 w-3.5" />
                          <span>{clockedIn ? 'Clock Out' : 'Clock In Now'}</span>
                        </button>
                      </div>
                    );
                  })}
                </div>

                {/* Attendance History Table */}
                <div className="bg-white rounded-2xl border border-slate-200 overflow-hidden shadow-xs">
                  <div className="p-4 border-b border-slate-100 flex items-center justify-between">
                    <div>
                      <h3 className="text-xs font-black uppercase tracking-wider text-slate-900">
                        Attendance History &amp; Work Hours Sheet
                      </h3>
                      <p className="text-[10px] text-slate-400">Permanently saved time clock punches</p>
                    </div>
                    <span className="text-xs font-mono text-slate-600 font-bold">{attendanceLogs.length} Records</span>
                  </div>

                  <div className="overflow-x-auto">
                    <table className="w-full text-left text-xs">
                      <thead className="bg-slate-50 text-[10px] font-black uppercase text-slate-400 border-b border-slate-200">
                        <tr>
                          <th className="py-2.5 px-4">Date</th>
                          <th className="py-2.5 px-4">Employee</th>
                          <th className="py-2.5 px-4">Role</th>
                          <th className="py-2.5 px-4">Clock In</th>
                          <th className="py-2.5 px-4">Clock Out</th>
                          <th className="py-2.5 px-4 text-center">Total Hours</th>
                          <th className="py-2.5 px-4 text-center">Overtime (&gt;8h)</th>
                          <th className="py-2.5 px-4 text-right">Action</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-100">
                        {attendanceLogs.length === 0 ? (
                          <tr>
                            <td colSpan={8} className="py-8 text-center text-slate-400 italic">
                              No attendance punches recorded yet.
                            </td>
                          </tr>
                        ) : (
                          attendanceLogs.map(att => (
                            <tr key={att.id} className="hover:bg-slate-50">
                              <td className="py-3 px-4 font-mono font-bold text-slate-700">{att.date}</td>
                              <td className="py-3 px-4 font-bold text-slate-900">{att.staffName}</td>
                              <td className="py-3 px-4 text-slate-500">{att.role}</td>
                              <td className="py-3 px-4 font-mono text-emerald-700 font-bold">{att.clockInTime}</td>
                              <td className="py-3 px-4 font-mono text-slate-700 font-bold">{att.clockOutTime || '--:--'}</td>
                              <td className="py-3 px-4 text-center font-mono font-bold">{att.totalHours ? `${att.totalHours} hrs` : '--'}</td>
                              <td className="py-3 px-4 text-center">
                                {att.isOvertime ? (
                                  <span className="px-2 py-0.5 rounded-full text-[9px] font-black uppercase bg-amber-100 text-amber-800">
                                    Overtime
                                  </span>
                                ) : (
                                  <span className="text-slate-400 font-mono text-[11px]">Normal</span>
                                )}
                              </td>
                              <td className="py-3 px-4 text-right">
                                {currentUser.role === 'Administrator' && (
                                  <button
                                    type="button"
                                    onClick={() => {
                                      if (window.confirm('Delete attendance punch record? Note: This will only remove it locally; permanent archives remain on cloud.')) {
                                        setAttendanceLogs(prev => prev.filter(a => a.id !== att.id));
                                        recordAuditLog(
                                          'ADMIN_DELETE_ATTENDANCE_LOG',
                                          att.id,
                                          `Admin ${currentUser.name} removed attendance record ${att.id} for ${att.staffName}`
                                        );
                                      }
                                    }}
                                    className="p-1 text-slate-400 hover:text-rose-600 transition-colors"
                                    title="Delete Record"
                                  >
                                    <Trash2 className="h-3.5 w-3.5" />
                                  </button>
                                )}
                              </td>
                            </tr>
                          ))
                        )}
                      </tbody>
                    </table>
                  </div>
                </div>
              </div>
            )}

            {/* SUB-TAB 2: PAYROLL RECORDS & PAY SLIPS */}
            {payrollSubTab === 'payslips' && (
              <div className="space-y-6">
                <div className="bg-white rounded-2xl border border-slate-200 overflow-hidden shadow-xs">
                  <div className="p-4 border-b border-slate-100 flex items-center justify-between">
                    <div>
                      <h3 className="text-xs font-black uppercase tracking-wider text-slate-900">
                        Processed Salary &amp; Wage Slips
                      </h3>
                      <p className="text-[10px] text-slate-400">Includes Basic, Allowances, EPF 8% &amp; Service Gratuity pool</p>
                    </div>
                    <span className="text-xs font-mono text-slate-600 font-bold">{payrollRecords.length} Slips</span>
                  </div>

                  <div className="overflow-x-auto">
                    <table className="w-full text-left text-xs">
                      <thead className="bg-slate-50 text-[10px] font-black uppercase text-slate-400 border-b border-slate-200">
                        <tr>
                          <th className="py-2.5 px-3">Slip ID</th>
                          <th className="py-2.5 px-3">Employee</th>
                          <th className="py-2.5 px-3">Period</th>
                          <th className="py-2.5 px-3 text-right">Basic + Allowances</th>
                          <th className="py-2.5 px-3 text-right">EPF Base</th>
                          <th className="py-2.5 px-3 text-right text-rose-600">EPF 8%</th>
                          <th className="py-2.5 px-3 text-right text-amber-700">Advance</th>
                          <th className="py-2.5 px-3 text-right text-emerald-600">Service Pool</th>
                          <th className="py-2.5 px-3 text-right font-black">Net Take-Home</th>
                          <th className="py-2.5 px-3 text-right">Actions</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-100">
                        {payrollRecords.length === 0 ? (
                          <tr>
                            <td colSpan={9} className="py-8 text-center text-slate-400 italic">
                              No payroll records processed yet. Click &ldquo;Process Pay Slip&rdquo; above.
                            </td>
                          </tr>
                        ) : (
                          payrollRecords.map(rec => {
                            const canManage = currentUser.role === 'Administrator' || currentUser.role === 'Manager';

                            return (
                              <tr key={rec.id} className="hover:bg-slate-50 transition-colors">
                                <td className="py-3 px-3 font-mono font-bold text-slate-800">{rec.id}</td>
                                <td className="py-3 px-3">
                                  <span className="font-bold text-slate-900 block">{rec.staffName}</span>
                                  <span className="text-[10px] text-slate-400">{rec.role}</span>
                                </td>
                                <td className="py-3 px-3 font-mono font-bold text-slate-700">{rec.period}</td>
                                <td className="py-3 px-3 text-right font-mono">
                                  {settings.currency} {((rec.breakdown?.basic || 0) + (rec.breakdown?.bra || 0) + (rec.breakdown?.allowances || 0)).toFixed(2)}
                                </td>
                                <td className="py-3 px-3 text-right font-mono font-semibold text-slate-800">
                                  {rec.breakdown?.epfEtfEnabled !== false 
                                    ? `${settings.currency} ${(rec.breakdown?.epfLiableEarnings || 0).toFixed(2)}`
                                    : <span className="text-slate-400 italic text-[10px]">Exempt</span>}
                                </td>
                                <td className="py-3 px-3 text-right font-mono text-rose-600 font-bold">
                                  {rec.breakdown?.epfEtfEnabled !== false 
                                    ? `-${settings.currency} ${(rec.breakdown?.epfEmployee || 0).toFixed(2)}`
                                    : '0.00'}
                                </td>
                                <td className="py-3 px-3 text-right font-mono text-amber-700 font-bold">
                                  {(rec.breakdown?.salaryAdvance || 0) > 0 
                                    ? `-${settings.currency} ${(rec.breakdown?.salaryAdvance || 0).toFixed(2)}` 
                                    : '0.00'}
                                </td>
                                <td className="py-3 px-3 text-right font-mono text-emerald-600 font-bold">
                                  +{(rec.breakdown?.serviceChargeBonus || 0) > 0 ? `${settings.currency} ${(rec.breakdown?.serviceChargeBonus || 0).toFixed(2)}` : '0.00'}
                                </td>
                                <td className="py-3 px-3 text-right font-mono font-black text-slate-950 text-sm">
                                  {settings.currency} {(rec.breakdown?.netSalary || rec.netPay || 0).toFixed(2)}
                                </td>
                                <td className="py-3 px-3 text-right">
                                  <div className="flex items-center justify-end gap-1.5">
                                    {/* Print Advance Slip if advance was given */}
                                    {Number(rec.breakdown?.salaryAdvance) > 0 && (
                                      <button
                                        type="button"
                                        onClick={() => {
                                          triggerAutoPrint({
                                            type: 'SALARY_ADVANCE_VOUCHER',
                                            data: {
                                              id: `ADV-${rec.id.slice(-6)}`,
                                              staffName: rec.staffName,
                                              role: rec.role,
                                              period: rec.period,
                                              amount: rec.breakdown.salaryAdvance,
                                              notes: rec.notes
                                            }
                                          }, `Reprint Advance: ${rec.staffName}`);
                                        }}
                                        className="p-1.5 text-amber-600 hover:text-amber-800 hover:bg-amber-50 rounded-lg cursor-pointer transition-colors"
                                        title="Reprint Advance Voucher"
                                      >
                                        <Banknote className="h-4 w-4" />
                                      </button>
                                    )}

                                    {/* Print Full Payslip */}
                                    <button
                                      type="button"
                                      onClick={() => {
                                        triggerAutoPrint({
                                          type: 'PAYSLIP_PRINT',
                                          data: rec
                                        }, `Payslip ${rec.id} - ${rec.staffName}`);
                                      }}
                                      className="p-1.5 text-slate-400 hover:text-slate-900 hover:bg-slate-100 rounded-lg cursor-pointer transition-colors"
                                      title="Print Payslip"
                                    >
                                      <Printer className="h-4 w-4" />
                                    </button>

                                    {/* Edit */}
                                    {canManage && (
                                      <button
                                        type="button"
                                        onClick={() => {
                                          setPayrollInputForm({
                                            staffId: rec.staffId,
                                            period: rec.period,
                                            epfEtfEnabled: rec.breakdown?.epfEtfEnabled !== false,
                                            basicSalary: rec.breakdown?.basic || 35000,
                                            budgetaryAllowance: rec.breakdown?.bra || 2500,
                                            otherAllowances: rec.breakdown?.allowances || 0,
                                            serviceChargeBonus: rec.breakdown?.serviceChargeBonus || 0,
                                            incentiveBonus: rec.breakdown?.incentiveBonus || 0,
                                            overtimeHours: (rec.breakdown?.overtimePay && rec.breakdown?.overtimeRate) 
                                              ? (rec.breakdown.overtimePay / rec.breakdown.overtimeRate) 
                                              : 0,
                                            overtimeRate: 250,
                                            otherDeductions: rec.breakdown?.otherDeductions || 0,
                                            notes: rec.notes || ''
                                          });
                                          setEditingPayrollId(rec.id);
                                          setProcessPayModalOpen(true);
                                        }}
                                        className="p-1.5 text-slate-400 hover:text-indigo-600 hover:bg-indigo-50 rounded-lg cursor-pointer transition-colors"
                                        title="Edit Payslip"
                                      >
                                        <Edit3 className="h-4 w-4" />
                                      </button>
                                    )}

                                    {/* Delete */}
                                    {currentUser.role === 'Administrator' && (
                                      <button
                                        type="button"
                                        onClick={() => {
                                          if (window.confirm(`Delete pay slip ${rec.id} for ${rec.staffName}? This action will be recorded in the audit log.`)) {
                                            setPayrollRecords(prev => prev.filter(r => r.id !== rec.id));
                                            recordAuditLog(
                                              'ADMIN_DELETE_PAYSLIP',
                                              rec.id,
                                              `Admin ${currentUser.name} deleted pay slip ${rec.id} for ${rec.staffName} (${rec.period})`
                                            );
                                          }
                                        }}
                                        className="p-1.5 text-slate-400 hover:text-rose-600 hover:bg-rose-50 rounded-lg cursor-pointer transition-colors"
                                        title="Delete Payslip"
                                      >
                                        <Trash2 className="h-4 w-4" />
                                      </button>
                                    )}
                                  </div>
                                </td>
                              </tr>
                            );
                          })
                        )}
                      </tbody>
                    </table>
                  </div>
                </div>
                
                {/* 2. SALARY ADVANCES DISBURSED LEDGER */}
                <div className="bg-white rounded-2xl border border-slate-200 overflow-hidden shadow-xs mt-6">
                  <div className="p-4 border-b border-slate-100 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                    <div>
                      <h3 className="text-xs font-black uppercase tracking-wider text-slate-900 flex items-center gap-1.5">
                        <Banknote className="h-4 w-4 text-amber-600" />
                        Salary Advance Disbursements &amp; Recovery Ledger
                      </h3>
                      <p className="text-[10px] text-slate-400">
                        Tracks advances issued to workers and the scheduled monthly payroll recovery
                      </p>
                    </div>

                    <div className="flex items-center gap-2">
                      <span className="text-xs font-mono font-bold text-amber-900 bg-amber-50 px-2.5 py-1 rounded-xl border border-amber-200">
                        Total Advances: {settings.currency} {(salaryAdvances || []).reduce((acc, a) => acc + (Number(a.amount) || 0), 0).toFixed(2)}
                      </span>
                      <button
                        type="button"
                        onClick={() => {
                          const firstStaff = staffList[0];
                          setAdvanceForm({
                            staffId: firstStaff ? firstStaff.id : '',
                            period: `${new Date().getFullYear()}-${String(new Date().getMonth() + 1).padStart(2, '0')}`,
                            amount: '',
                            paymentMethod: 'CASH',
                            reason: 'Emergency advance on salary',
                            notes: ''
                          });
                          setIssueAdvanceModalOpen(true);
                        }}
                        className="px-3 py-1 bg-amber-500 hover:bg-amber-600 text-white rounded-lg text-xs font-bold flex items-center gap-1 shadow-xs cursor-pointer transition-colors"
                      >
                        <Plus className="h-3.5 w-3.5" />
                        <span>Issue Advance</span>
                      </button>
                    </div>
                  </div>

                  <div className="overflow-x-auto">
                    <table className="w-full text-left text-xs">
                      <thead className="bg-slate-50 text-[10px] font-black uppercase text-slate-400 border-b border-slate-200">
                        <tr>
                          <th className="py-2.5 px-3">Voucher Ref</th>
                          <th className="py-2.5 px-3">Date &amp; Time</th>
                          <th className="py-2.5 px-3">Employee</th>
                          <th className="py-2.5 px-3">Recovery Month</th>
                          <th className="py-2.5 px-3">Method</th>
                          <th className="py-2.5 px-3">Reason / Notes</th>
                          <th className="py-2.5 px-3 text-right">Amount</th>
                          <th className="py-2.5 px-3 text-center">Status</th>
                          <th className="py-2.5 px-3 text-right">Actions</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-100">
                        {(!salaryAdvances || salaryAdvances.length === 0) ? (
                          <tr>
                            <td colSpan={9} className="py-8 text-center text-slate-400 italic">
                              No salary advance disbursements recorded yet. Click &ldquo;Issue Advance&rdquo; to disburse funds.
                            </td>
                          </tr>
                        ) : (
                          salaryAdvances.map(adv => {
                            // Check if this advance was already recovered on a settled payslip
                            const isRecovered = payrollRecords.some(
                              rec => rec.staffId === adv.staffId && rec.period === adv.period
                            );

                            return (
                              <tr key={adv.id} className="hover:bg-slate-50 transition-colors">
                                <td className="py-3 px-3 font-mono font-bold text-slate-800">{adv.id}</td>
                                <td className="py-3 px-3 text-slate-500 whitespace-nowrap">
                                  {adv.date} {adv.disbursedAt}
                                </td>
                                <td className="py-3 px-3">
                                  <span className="font-bold text-slate-900 block">{adv.staffName}</span>
                                  <span className="text-[10px] text-slate-400">{adv.role}</span>
                                </td>
                                <td className="py-3 px-3 font-mono font-bold text-amber-800 bg-amber-50/50">
                                  {adv.period}
                                </td>
                                <td className="py-3 px-3">
                                  <span className={`px-2 py-0.5 rounded text-[10px] font-mono font-bold ${
                                    adv.paymentMethod === 'CASH'
                                      ? 'bg-emerald-50 text-emerald-800 border border-emerald-200'
                                      : 'bg-indigo-50 text-indigo-800 border border-indigo-200'
                                  }`}>
                                    {adv.paymentMethod}
                                  </span>
                                </td>
                                <td className="py-3 px-3 text-slate-600 max-w-xs truncate" title={adv.notes || adv.reason}>
                                  {adv.reason || adv.notes || 'Advance on wages'}
                                </td>
                                <td className="py-3 px-3 text-right font-mono font-black text-amber-900 text-sm">
                                  {settings.currency} {Number(adv.amount).toFixed(2)}
                                </td>
                                <td className="py-3 px-3 text-center">
                                  <span className={`px-2 py-0.5 rounded-full text-[10px] font-bold ${
                                    isRecovered
                                      ? 'bg-emerald-100 text-emerald-800'
                                      : 'bg-amber-100 text-amber-800'
                                  }`}>
                                    {isRecovered ? 'Deducted (Recovered)' : 'Pending Deduction'}
                                  </span>
                                </td>
                                <td className="py-3 px-3 text-right">
                                  <div className="flex items-center justify-end gap-1.5">
                                    {/* Reprint Advance Voucher */}
                                    <button
                                      type="button"
                                      onClick={() => {
                                        triggerAutoPrint({
                                          type: 'SALARY_ADVANCE_VOUCHER',
                                          data: {
                                            id: adv.id,
                                            staffName: adv.staffName,
                                            role: adv.role,
                                            period: adv.period,
                                            amount: adv.amount,
                                            notes: adv.notes || adv.reason
                                          }
                                        }, `Reprint Advance Voucher: ${adv.staffName}`);
                                      }}
                                      className="p-1.5 text-slate-400 hover:text-slate-900 hover:bg-slate-100 rounded-lg cursor-pointer transition-colors"
                                      title="Reprint Advance Voucher"
                                    >
                                      <Printer className="h-4 w-4" />
                                    </button>

                                    {/* Delete Advance (Admin only) */}
                                    {currentUser.role === 'Administrator' && (
                                      <button
                                        type="button"
                                        onClick={() => {
                                          if (window.confirm(`Delete advance voucher ${adv.id} (${settings.currency} ${adv.amount}) for ${adv.staffName}?`)) {
                                            setSalaryAdvances(prev => prev.filter(a => a.id !== adv.id));
                                            recordAuditLog(
                                              'ADMIN_DELETE_SALARY_ADVANCE',
                                              adv.id,
                                              `Deleted advance of ${settings.currency} ${adv.amount} for ${adv.staffName} (${adv.period})`
                                            );
                                          }
                                        }}
                                        className="p-1.5 text-slate-400 hover:text-rose-600 hover:bg-rose-50 rounded-lg cursor-pointer transition-colors"
                                        title="Delete Advance Record"
                                      >
                                        <Trash2 className="h-4 w-4" />
                                      </button>
                                    )}
                                  </div>
                                </td>
                              </tr>
                            );
                          })
                        )}
                      </tbody>
                    </table>
                  </div>
                </div>
              </div>
            )}

            {/* SUB-TAB 3: SRI LANKA EPF / ETF RETURN (FORM C EQUIVALENT) */}
            {payrollSubTab === 'epf_etf' && (
              <div className="space-y-6">
                {(() => {
                  let totalEpfBase = 0;
                  let totalEpfEmp8 = 0;
                  let totalEpfEmpr12 = 0;
                  let totalEtfEmpr3 = 0;

                  payrollRecords.forEach(r => {
                    if (r.breakdown) {
                      totalEpfBase += (r.breakdown.epfLiableEarnings || 0);
                      totalEpfEmp8 += (r.breakdown.epfEmployee || 0);
                      totalEpfEmpr12 += (r.breakdown.epfEmployer || 0);
                      totalEtfEmpr3 += (r.breakdown.etfEmployer || 0);
                    }
                  });

                  return (
                    <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
                      <div className="bg-white p-4 rounded-2xl border border-slate-200 shadow-xs">
                        <p className="text-[10px] font-black uppercase text-slate-400">Total Liable Earnings (EPF Base)</p>
                        <p className="text-xl font-black font-mono text-slate-900 mt-1">
                          {settings.currency} {totalEpfBase.toFixed(2)}
                        </p>
                        <span className="text-[10px] text-slate-400">Basic + Budgetary Allowances</span>
                      </div>

                      <div className="bg-white p-4 rounded-2xl border border-slate-200 shadow-xs">
                        <p className="text-[10px] font-black uppercase text-slate-400">EPF Employee Share (8%)</p>
                        <p className="text-xl font-black font-mono text-rose-600 mt-1">
                          {settings.currency} {totalEpfEmp8.toFixed(2)}
                        </p>
                        <span className="text-[10px] text-slate-400">Deducted from workers</span>
                      </div>

                      <div className="bg-white p-4 rounded-2xl border border-slate-200 shadow-xs">
                        <p className="text-[10px] font-black uppercase text-slate-400">EPF Employer Share (12%)</p>
                        <p className="text-xl font-black font-mono text-indigo-600 mt-1">
                          {settings.currency} {totalEpfEmpr12.toFixed(2)}
                        </p>
                        <span className="text-[10px] text-slate-400">Paid by business</span>
                      </div>

                      <div className="bg-white p-4 rounded-2xl border border-slate-200 shadow-xs">
                        <p className="text-[10px] font-black uppercase text-slate-400">ETF Employer Share (3%)</p>
                        <p className="text-xl font-black font-mono text-emerald-600 mt-1">
                          {settings.currency} {totalEtfEmpr3.toFixed(2)}
                        </p>
                        <span className="text-[10px] text-slate-400">Remitted to ETF Board</span>
                      </div>
                    </div>
                  );
                })()}

                <div className="bg-white rounded-2xl border border-slate-200 p-5 shadow-xs">
                  <div className="border-b border-slate-100 pb-3 mb-4">
                    <h3 className="text-sm font-black text-slate-900 uppercase">
                      Department of Labour &amp; Central Bank of Sri Lanka Form C Schedule
                    </h3>
                    <p className="text-xs text-slate-500">Official monthly remittance computation schedule</p>
                  </div>

                  <div className="overflow-x-auto">
                    <table className="w-full text-left text-xs">
                      <thead className="bg-slate-50 text-[10px] font-black uppercase text-slate-400 border-b border-slate-200">
                        <tr>
                          <th className="py-2.5 px-3">Employee Name</th>
                          <th className="py-2.5 px-3">EPF Member #</th>
                          <th className="py-2.5 px-3 text-right">Liable Wages</th>
                          <th className="py-2.5 px-3 text-right">Employee 8%</th>
                          <th className="py-2.5 px-3 text-right">Employer 12%</th>
                          <th className="py-2.5 px-3 text-right font-bold">Total EPF (20%)</th>
                          <th className="py-2.5 px-3 text-right">Employer ETF (3%)</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-100">
                        {payrollRecords.map(rec => (
                          <tr key={rec.id} className="hover:bg-slate-50">
                            <td className="py-2.5 px-3 font-bold text-slate-900">{rec.staffName}</td>
                            <td className="py-2.5 px-3 font-mono text-slate-500">EPF-{rec.staffId.slice(-4)}</td>
                            <td className="py-2.5 px-3 text-right font-mono">{settings.currency} {(rec.breakdown?.epfLiableEarnings || 0).toFixed(2)}</td>
                            <td className="py-2.5 px-3 text-right font-mono text-rose-600">{settings.currency} {(rec.breakdown?.epfEmployee || 0).toFixed(2)}</td>
                            <td className="py-2.5 px-3 text-right font-mono text-indigo-600">{settings.currency} {(rec.breakdown?.epfEmployer || 0).toFixed(2)}</td>
                            <td className="py-2.5 px-3 text-right font-mono font-black text-slate-900">
                              {settings.currency} {(rec.breakdown?.totalEpfFund || 0).toFixed(2)}
                            </td>
                            <td className="py-2.5 px-3 text-right font-mono text-emerald-700 font-bold">
                              {settings.currency} {(rec.breakdown?.etfEmployer || 0).toFixed(2)}
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              </div>
            )}

            {/* SUB-TAB 4: EMPLOYEE REGISTRY & BASE SALARIES */}
            {payrollSubTab === 'profiles' && (
              <div className="bg-white rounded-2xl border border-slate-200 overflow-hidden shadow-xs">
                <div className="p-4 border-b border-slate-100 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                  <div>
                    <h3 className="text-xs font-black uppercase tracking-wider text-slate-900">
                      Employee Registry, Wages &amp; Advances
                    </h3>
                    <p className="text-[10px] text-slate-400 mt-0.5">
                      View configured employee profiles, base salary structures, and active unrecovered salary advances
                    </p>
                  </div>

                  <span className="text-xs font-mono font-bold text-slate-600 bg-slate-50 px-2.5 py-1 rounded-xl border border-slate-200 self-start sm:self-auto">
                    {staffList.length} Registered Staff
                  </span>
                </div>

                <div className="overflow-x-auto">
                  <table className="w-full text-left text-xs">
                    <thead className="bg-slate-50 text-[10px] font-black uppercase text-slate-400 border-b border-slate-200">
                      <tr>
                        <th className="py-3 px-4">Employee</th>
                        <th className="py-3 px-4">Role</th>
                        <th className="py-3 px-4 text-right">Base Salary</th>
                        <th className="py-3 px-4 text-right">Pending Advances</th>
                        <th className="py-3 px-4">Contact</th>
                        <th className="py-3 px-4">NIC / National ID</th>
                        <th className="py-3 px-4 text-center">Status</th>
                        <th className="py-3 px-4 text-right">Quick Actions</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100">
                      {staffList.map(member => {
                        // Calculate total unrecovered salary advances for this employee
                        const unrecoveredAdvance = (salaryAdvances || [])
                          .filter(adv => 
                            adv.staffId === member.id && 
                            adv.status !== 'REJECTED' &&
                            !payrollRecords.some(r => r.staffId === member.id && r.period === adv.period)
                          )
                          .reduce((sum, a) => sum + (Number(a.amount) || 0), 0);

                        return (
                          <tr key={member.id} className="hover:bg-slate-50 transition-colors">
                            <td className="py-3 px-4">
                              <span className="font-extrabold text-slate-900 block">{member.name}</span>
                              <span className="text-[10px] text-slate-400 font-mono">{member.id}</span>
                            </td>
                            <td className="py-3 px-4">
                              <span className="px-2 py-0.5 bg-orange-100 text-[#ff5500] rounded font-bold text-[10px]">
                                {member.role}
                              </span>
                            </td>
                            <td className="py-3 px-4 text-right font-mono font-bold text-slate-900">
                              {settings.currency} {(member.basicSalary ?? 35000).toLocaleString('en-US', { minimumFractionDigits: 2 })}
                            </td>
                            <td className="py-3 px-4 text-right font-mono font-bold">
                              {unrecoveredAdvance > 0 ? (
                                <span className="text-amber-800 bg-amber-50 px-2 py-0.5 rounded-lg border border-amber-200 inline-block text-xs">
                                  {settings.currency} {unrecoveredAdvance.toFixed(2)}
                                </span>
                              ) : (
                                <span className="text-slate-400 text-[11px] font-normal">None</span>
                              )}
                            </td>
                            <td className="py-3 px-4 text-slate-600">{member.email}</td>
                            <td className="py-3 px-4 font-mono text-slate-700">1992{member.pin}402V</td>
                            <td className="py-3 px-4 text-center">
                              <span className="px-2 py-0.5 rounded-full text-[9px] font-bold bg-emerald-100 text-emerald-800">
                                Active
                              </span>
                            </td>
                            <td className="py-3 px-4 text-right">
                              <div className="flex items-center justify-end gap-1.5">
                                {/* Direct Advance Disbursement Button */}
                                <button
                                  type="button"
                                  onClick={() => {
                                    setAdvanceForm({
                                      staffId: member.id,
                                      period: `${new Date().getFullYear()}-${String(new Date().getMonth() + 1).padStart(2, '0')}`,
                                      amount: '',
                                      paymentMethod: 'CASH',
                                      reason: 'Mid-month advance on salary',
                                      notes: ''
                                    });
                                    setIssueAdvanceModalOpen(true);
                                  }}
                                  className="px-2.5 py-1 bg-amber-50 hover:bg-amber-100 text-amber-800 font-bold rounded-lg text-[11px] flex items-center gap-1 transition-colors cursor-pointer border border-amber-200"
                                  title="Disburse salary advance to this employee"
                                >
                                  <Banknote className="h-3 w-3" />
                                  <span>Advance</span>
                                </button>

                                {/* Direct Issue Pay Button */}
                                <button
                                  type="button"
                                  onClick={() => {
                                    const curPeriod = `${new Date().getFullYear()}-${String(new Date().getMonth() + 1).padStart(2, '0')}`;
                                    const advancesForMonth = (salaryAdvances || [])
                                      .filter(a => a.staffId === member.id && a.period === curPeriod && a.status !== 'REJECTED')
                                      .reduce((acc, a) => acc + (Number(a.amount) || 0), 0);

                                    setPayrollInputForm({
                                      staffId: member.id,
                                      period: curPeriod,
                                      epfEtfEnabled: member.epfEtfEnabled !== false,
                                      basicSalary: member.basicSalary ?? 35000,
                                      budgetaryAllowance: member.budgetaryAllowance ?? 2500,
                                      otherAllowances: member.otherAllowances ?? 0,
                                      serviceChargeBonus: Math.round((salesMetrics?.serviceCharge || 0) / Math.max(1, staffList.length)),
                                      incentiveBonus: member.fixedBonus ?? 0,
                                      overtimeHours: 0,
                                      overtimeRate: member.overtimeRate ?? 250,
                                      salaryAdvance: advancesForMonth,
                                      otherDeductions: 0,
                                      standardWorkingDays: 26,
                                      workedDays: 26,
                                      paidLeaves: 0,
                                      unpaidLeaves: 0,
                                      holidaysCount: 4,
                                      shortShiftsCount: 0,
                                      notes: `Standard wages for ${member.name}`
                                    });
                                    setEditingPayrollId(null);
                                    setActiveTab('payroll');
                                    setPayrollSubTab('payslips');
                                    setProcessPayModalOpen(true);
                                  }}
                                  className="px-2.5 py-1 bg-emerald-50 hover:bg-emerald-100 text-emerald-700 font-bold rounded-lg text-[11px] flex items-center gap-1 transition-colors cursor-pointer border border-emerald-200"
                                  title="Run monthly payslip for this employee"
                                >
                                  <Briefcase className="h-3 w-3" />
                                  <span>Issue Pay</span>
                                </button>
                              </div>
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              </div>
            )}
          </div>
        )}

        {/* VIEW 13: SYSTEM SETTINGS */}
        {activeTab === 'settings' && (
          <div className="flex-1 overflow-y-auto p-6 lg:p-8 space-y-6 max-w-5xl mx-auto w-full">
            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 border-b border-slate-200 pb-5">
              <div>
                <h2 className="text-xl font-black text-slate-900 flex items-center gap-2">
                  <Settings className="h-5 w-5 text-[#ff5500]" />
                  System, Business &amp; Peripheral Settings
                </h2>
                <p className="text-xs text-slate-500 mt-1">
                  Manage company identity, thermal printing options, automated cash drawer solenoid, and database backup files.
                </p>
              </div>

              <div className="flex items-center gap-2">
                <button
                  type="button"
                  onClick={handleExportBackup}
                  className="px-3.5 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold rounded-xl text-xs flex items-center gap-1.5 transition-colors cursor-pointer"
                  title="Export complete system database to JSON file"
                >
                  <Download className="h-3.5 w-3.5 text-slate-600" />
                  <span>Download Backup</span>
                </button>

                <button
                  type="button"
                  onClick={() => {
                    setSettingsNotice({
                      title: 'Settings Saved',
                      detail: `Configuration for ${settings.restaurantName} updated successfully.`
                    });
                    setTimeout(() => setSettingsNotice(null), 3500);
                  }}
                  className="px-4 py-2 bg-[#008f5d] hover:bg-emerald-700 text-white font-bold rounded-xl text-xs flex items-center gap-1.5 shadow-sm transition-all cursor-pointer"
                >
                  <CheckCircle2 className="h-4 w-4" />
                  <span>Save Changes</span>
                </button>
              </div>
            </div>

            {/* SECTION 1: FULL COMPANY & BUSINESS DETAILS */}
            <div className="bg-white rounded-2xl border border-slate-200 p-6 shadow-xs space-y-4">
              <div className="flex items-center justify-between border-b border-slate-100 pb-3">
                <div className="flex items-center gap-2">
                  <div className="p-2 bg-orange-50 text-[#ff5500] rounded-xl">
                    <Building className="h-4 w-4" />
                  </div>
                  <div>
                    <h3 className="text-xs font-black uppercase tracking-wider text-slate-900">
                      Company &amp; Business Information
                    </h3>
                    <p className="text-[10px] text-slate-400">Printed on official receipts, tax invoices, and Z-reports</p>
                  </div>
                </div>
              </div>

              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">Trading / Brand Name</label>
                  <input
                    type="text"
                    value={settings.restaurantName}
                    onChange={e => setSettings(prev => ({ ...prev, restaurantName: e.target.value }))}
                    placeholder="e.g. Linoli Cove Midigama"
                    className="w-full px-3.5 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-semibold text-slate-900 focus:bg-white focus:outline-none focus:border-[#ff5500]"
                  />
                </div>

                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">Brand Tagline / Slogan</label>
                  <input
                    type="text"
                    value={settings.tagline}
                    onChange={e => setSettings(prev => ({ ...prev, tagline: e.target.value }))}
                    placeholder="e.g. RESTAURANT & BAR"
                    className="w-full px-3.5 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-semibold text-slate-900 focus:bg-white focus:outline-none focus:border-[#ff5500]"
                  />
                </div>

                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">Registered Legal Entity Name</label>
                  <input
                    type="text"
                    value={settings.legalName || ''}
                    onChange={e => setSettings(prev => ({ ...prev, legalName: e.target.value }))}
                    placeholder="e.g. Linoli Cove Leisure (Pvt) Ltd"
                    className="w-full px-3.5 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-medium text-slate-900 focus:bg-white focus:outline-none focus:border-[#ff5500]"
                  />
                </div>

                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">Business Registration No. (BRN / Company ID)</label>
                  <input
                    type="text"
                    value={settings.businessRegNo || ''}
                    onChange={e => setSettings(prev => ({ ...prev, businessRegNo: e.target.value }))}
                    placeholder="e.g. PV-00289144"
                    className="w-full px-3.5 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-mono text-slate-900 focus:bg-white focus:outline-none focus:border-[#ff5500]"
                  />
                </div>

                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">Tax Identification / VAT / GST No.</label>
                  <input
                    type="text"
                    value={settings.taxId || ''}
                    onChange={e => setSettings(prev => ({ ...prev, taxId: e.target.value }))}
                    placeholder="e.g. TIN-109284719"
                    className="w-full px-3.5 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-mono text-slate-900 focus:bg-white focus:outline-none focus:border-[#ff5500]"
                  />
                </div>

                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">Terminal Hardware Identifier</label>
                  <input
                    type="text"
                    value={settings.terminalId}
                    onChange={e => setSettings(prev => ({ ...prev, terminalId: e.target.value }))}
                    placeholder="e.g. LINOLI-MAIN-01"
                    className="w-full px-3.5 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-mono font-bold text-slate-900 focus:bg-white focus:outline-none focus:border-[#ff5500]"
                  />
                </div>

                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">Contact Phone Number</label>
                  <input
                    type="text"
                    value={settings.phone || ''}
                    onChange={e => setSettings(prev => ({ ...prev, phone: e.target.value }))}
                    placeholder="e.g. +94 74 036 6741"
                    className="w-full px-3.5 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-medium text-slate-900 focus:bg-white focus:outline-none focus:border-[#ff5500]"
                  />
                </div>

                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">Business Email Address</label>
                  <input
                    type="email"
                    value={settings.email || ''}
                    onChange={e => setSettings(prev => ({ ...prev, email: e.target.value }))}
                    placeholder="e.g. info@linolicove.me"
                    className="w-full px-3.5 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-medium text-slate-900 focus:bg-white focus:outline-none focus:border-[#ff5500]"
                  />
                </div>

                <div className="md:col-span-2">
                  <label className="block text-xs font-bold text-slate-700 mb-1">Official Website or Social Link</label>
                  <input
                    type="text"
                    value={settings.website || ''}
                    onChange={e => setSettings(prev => ({ ...prev, website: e.target.value }))}
                    placeholder="e.g. www.linolicove.me / @linolicove"
                    className="w-full px-3.5 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-medium text-slate-900 focus:bg-white focus:outline-none focus:border-[#ff5500]"
                  />
                </div>

                <div className="md:col-span-2">
                  <label className="block text-xs font-bold text-slate-700 mb-1">Full Physical Street Address</label>
                  <input
                    type="text"
                    value={settings.address || ''}
                    onChange={e => setSettings(prev => ({ ...prev, address: e.target.value }))}
                    placeholder="e.g. 380 A Matara Road, Midigama, 81700, Sri Lanka"
                    className="w-full px-3.5 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-medium text-slate-900 focus:bg-white focus:outline-none focus:border-[#ff5500]"
                  />
                </div>
              </div>
            </div>

            {/* SECTION 2: AUTOMATED DAILY EMAIL DISPATCH */}
            <div className="bg-white rounded-2xl border border-slate-200 p-6 shadow-xs space-y-4">
              <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 border-b border-slate-100 pb-3">
                <div className="flex items-center gap-2">
                  <div className="p-2 bg-orange-50 text-[#ff5500] rounded-xl">
                    <Mail className="h-4 w-4" />
                  </div>
                  <div>
                    <h3 className="text-xs font-black uppercase tracking-wider text-slate-900">
                      Automated Daily 11:30 PM Email Dispatch
                    </h3>
                    <p className="text-[10px] text-slate-400">
                      Auto-transmits complete end-of-day sales, settlements, invoices, voids &amp; audit records
                    </p>
                  </div>
                </div>

                <div className="flex items-center gap-2">
                  <span className={`px-2.5 py-1 rounded-full text-[10px] font-mono font-bold ${
                    emailSettings.enabled
                      ? 'bg-emerald-100 text-emerald-800'
                      : 'bg-slate-100 text-slate-500'
                  }`}>
                    {emailSettings.enabled ? `● Scheduled: ${emailSettings.scheduledTime || '23:30'} Daily` : 'Disabled'}
                  </span>
                  {emailSettings.lastSentDate && (
                    <span className="text-[10px] font-mono text-slate-500">
                      Last Sent: {emailSettings.lastSentDate}
                    </span>
                  )}
                </div>
              </div>

              <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">Target Recipient Email</label>
                  <input
                    type="email"
                    value={emailSettings.recipient}
                    onChange={e => setEmailSettings(prev => ({ ...prev, recipient: e.target.value }))}
                    placeholder="linolicove@gmail.com"
                    className="w-full px-3.5 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-semibold text-slate-900 focus:bg-white focus:outline-none focus:border-[#ff5500]"
                  />
                  <span className="text-[10px] text-slate-400 mt-1 block">Receives complete daily business data</span>
                </div>

                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">Scheduled Time (24h)</label>
                  <input
                    type="text"
                    value={emailSettings.scheduledTime}
                    onChange={e => setEmailSettings(prev => ({ ...prev, scheduledTime: e.target.value }))}
                    placeholder="23:30"
                    className="w-full px-3.5 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-mono font-bold text-slate-900 focus:bg-white focus:outline-none focus:border-[#ff5500]"
                  />
                  <span className="text-[10px] text-slate-400 mt-1 block">Default: 23:30 (11:30 PM)</span>
                </div>

                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">Automation Status</label>
                  <select
                    value={emailSettings.enabled ? 'YES' : 'NO'}
                    onChange={e => setEmailSettings(prev => ({ ...prev, enabled: e.target.value === 'YES' }))}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-bold text-slate-900 focus:bg-white focus:outline-none focus:border-[#ff5500]"
                  >
                    <option value="YES">Enabled (Auto-send at 11:30 PM)</option>
                    <option value="NO">Disabled (Manual trigger only)</option>
                  </select>
                </div>
              </div>

              <div className="p-3.5 bg-slate-50 border border-slate-200 rounded-xl flex flex-col md:flex-row items-stretch md:items-center justify-between gap-3">
                <div className="space-y-0.5 text-xs">
                  <p className="font-bold text-slate-800">What data is transmitted in the 11:30 PM package?</p>
                  <p className="text-[11px] text-slate-500">
                    Gross revenue, net sales, taxes, service pool, discounts, individual invoice ledgers, cashier drawer count &amp; cash-outs, cancelled ticket voids, and the complete day's audit trail.
                  </p>
                </div>

                <div className="flex items-center gap-2 shrink-0">
                  <button
                    type="button"
                    disabled={isSendingEmail}
                    onClick={() => sendDailyEodEmail(true)}
                    className="px-4 py-2 bg-[#ff5500] hover:bg-orange-600 text-white font-bold rounded-xl text-xs flex items-center gap-1.5 shadow-xs transition-all disabled:opacity-50 cursor-pointer"
                  >
                    <Send className="h-3.5 w-3.5" />
                    <span>{isSendingEmail ? 'Dispatching...' : 'Send Daily Report Now'}</span>
                  </button>
                </div>
              </div>

              <details className="text-xs text-slate-600 pt-1">
                <summary className="font-bold cursor-pointer text-slate-700 hover:text-[#ff5500] select-none">
                  Advanced: Direct Silent Webhook or EmailJS API Keys (Optional)
                </summary>
                <div className="mt-3 p-3.5 bg-slate-50 border border-slate-200 rounded-xl grid grid-cols-1 sm:grid-cols-2 gap-3">
                  <div className="sm:col-span-2">
                    <label className="block text-[10px] font-extrabold uppercase text-slate-500 mb-1">
                      Webhook POST URL (Zapier / Make / Cloud Function)
                    </label>
                    <input
                      type="url"
                      value={emailSettings.webhookUrl}
                      onChange={e => setEmailSettings(prev => ({ ...prev, webhookUrl: e.target.value }))}
                      placeholder="https://hook.eu2.make.com/... or https://api.yoursite.com/eod-report"
                      className="w-full px-3 py-1.5 bg-white border border-slate-200 rounded-lg text-xs font-mono text-slate-900"
                    />
                    <p className="text-[10px] text-slate-400 mt-0.5">Posts JSON payload containing complete sales, shifts &amp; audit data directly.</p>
                  </div>
                  <div>
                    <label className="block text-[10px] font-extrabold uppercase text-slate-500 mb-1">EmailJS Service ID</label>
                    <input
                      type="text"
                      value={emailSettings.emailjsServiceId}
                      onChange={e => setEmailSettings(prev => ({ ...prev, emailjsServiceId: e.target.value }))}
                      placeholder="service_..."
                      className="w-full px-3 py-1.5 bg-white border border-slate-200 rounded-lg text-xs font-mono text-slate-900"
                    />
                  </div>
                  <div>
                    <label className="block text-[10px] font-extrabold uppercase text-slate-500 mb-1">EmailJS Template ID</label>
                    <input
                      type="text"
                      value={emailSettings.emailjsTemplateId}
                      onChange={e => setEmailSettings(prev => ({ ...prev, emailjsTemplateId: e.target.value }))}
                      placeholder="template_..."
                      className="w-full px-3 py-1.5 bg-white border border-slate-200 rounded-lg text-xs font-mono text-slate-900"
                    />
                  </div>
                </div>
              </details>
            </div>

            {/* SECTION 3: STREAMLINED AUTO-PRINTER CONFIGURATION */}
            <div className="bg-white rounded-2xl border border-slate-200 p-6 shadow-xs space-y-4">
              <div className="flex items-center justify-between border-b border-slate-100 pb-3">
                <div className="flex items-center gap-2">
                  <div className="p-2 bg-indigo-50 text-indigo-600 rounded-xl">
                    <Printer className="h-4 w-4" />
                  </div>
                  <div>
                    <h3 className="text-xs font-black uppercase tracking-wider text-slate-900">
                      Thermal Auto-Printer Configuration
                    </h3>
                    <p className="text-[10px] text-slate-400">ESC/POS thermal slips for KOT, BOT, proforma bills &amp; tax invoices</p>
                  </div>
                </div>
                <span className="text-[10px] font-mono font-bold px-2.5 py-1 rounded-full bg-slate-100 text-slate-600">
                  {pairedUsbDevice ? `USB: ${pairedUsbDevice.productName || 'Connected'}` : 'System Default Spooler'}
                </span>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">Paper Roll Width</label>
                  <select
                    value={settings.receiptRollWidth || '80mm'}
                    onChange={e => setSettings(prev => ({ ...prev, receiptRollWidth: e.target.value }))}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-bold text-slate-900 focus:bg-white focus:outline-none focus:border-[#ff5500]"
                  >
                    <option value="80mm">80mm Thermal Paper (Standard POS)</option>
                    <option value="58mm">58mm Thermal Paper (Compact)</option>
                  </select>
                </div>

                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">Receipt Font Size</label>
                  <select
                    value={settings.receiptFontSize || '11px'}
                    onChange={e => setSettings(prev => ({ ...prev, receiptFontSize: e.target.value }))}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-bold text-slate-900 focus:bg-white focus:outline-none focus:border-[#ff5500]"
                  >
                    <option value="10px">10px - Compact (Fits More Items)</option>
                    <option value="11px">11px - Standard (Recommended)</option>
                    <option value="12px">12px - Medium Large</option>
                    <option value="13px">13px - Large Text</option>
                    <option value="14px">14px - Extra Bold &amp; Large</option>
                  </select>
                </div>

                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">Receipt Font Type</label>
                  <select
                    value={settings.receiptFontFamily || 'monospace'}
                    onChange={e => setSettings(prev => ({ ...prev, receiptFontFamily: e.target.value }))}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-bold text-slate-900 focus:bg-white focus:outline-none focus:border-[#ff5500]"
                  >
                    <option value="monospace">Monospace (Classic ESC/POS Receipt)</option>
                    <option value="sans-serif">Sans-Serif (Modern Clean Helvetica/Arial)</option>
                    <option value="serif">Serif (Classic Traditional)</option>
                  </select>
                </div>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-3 gap-4 pt-1">
                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">Thermal Slip Margins</label>
                  <select
                    value={settings.receiptMargin || '2mm'}
                    onChange={e => setSettings(prev => ({ ...prev, receiptMargin: e.target.value }))}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-bold text-slate-900 focus:bg-white focus:outline-none focus:border-[#ff5500]"
                  >
                    <option value="0mm">0mm - Full Width (Edge-to-Edge)</option>
                    <option value="2mm">2mm - Standard Thermal Margin</option>
                    <option value="4mm">4mm - Comfortable Margin</option>
                    <option value="6mm">6mm - Wide Margin</option>
                  </select>
                </div>

                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">Auto-Print on Send Order</label>
                  <select
                    value={settings.autoPrintOrder !== false ? 'ENABLED' : 'DISABLED'}
                    onChange={e => setSettings(prev => ({ ...prev, autoPrintOrder: e.target.value === 'ENABLED' }))}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-bold text-slate-900 focus:bg-white focus:outline-none focus:border-[#ff5500]"
                  >
                    <option value="ENABLED">Yes - Print KOT &amp; BOT Slips</option>
                    <option value="DISABLED">No - Manual Print Only</option>
                  </select>
                </div>

                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">Auto-Print on Settlement</label>
                  <select
                    value={settings.autoPrintBill !== false ? 'ENABLED' : 'DISABLED'}
                    onChange={e => setSettings(prev => ({ ...prev, autoPrintBill: e.target.value === 'ENABLED' }))}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-bold text-slate-900 focus:bg-white focus:outline-none focus:border-[#ff5500]"
                  >
                    <option value="ENABLED">Yes - Print Final Tax Invoice</option>
                    <option value="DISABLED">No - Manual Print Only</option>
                  </select>
                </div>
              </div>

              <div className="p-3.5 bg-slate-50 border border-slate-200 rounded-xl flex flex-col sm:flex-row items-stretch sm:items-center justify-between gap-3">
                <div className="space-y-0.5">
                  <p className="text-xs font-bold text-slate-800">Direct WebUSB Thermal Printer Connection</p>
                  <p className="text-[10px] text-slate-500">
                    Pair once with your USB printer for fast ESC/POS output, or run via OS print spooler.
                  </p>
                  {usbStatusMessage && (
                    <p className="text-[10px] font-mono text-[#ff5500] mt-0.5">{usbStatusMessage}</p>
                  )}
                </div>

                <div className="flex items-center gap-2 shrink-0">
                  <button
                    type="button"
                    onClick={() => {
                      if (navigator.usb) {
                        navigator.usb.requestDevice({ filters: [] }).then(dev => {
                          setPairedUsbDevice(dev);
                          setUsbStatusMessage(`Paired with ${dev.productName || 'USB Printer'}`);
                        }).catch(() => {
                          setUsbStatusMessage('Pairing cancelled or printer busy.');
                        });
                      } else {
                        setUsbStatusMessage('WebUSB not supported; standard OS spooler active.');
                      }
                    }}
                    className="px-3 py-1.5 bg-slate-900 hover:bg-slate-800 text-white rounded-lg text-xs font-bold flex items-center gap-1.5 cursor-pointer"
                  >
                    <Usb className="h-3.5 w-3.5" />
                    <span>Pair USB Printer</span>
                  </button>

                  <button
                    type="button"
                    onClick={() => {
                      triggerAutoPrint({
                        type: 'TEMP_BILL',
                        data: {
                          table: 'TEST-PRINTER',
                          server: currentUser.name,
                          items: [
                            { name: 'Diagnostic Test Print', qty: 1, price: 0.00 }
                          ],
                          subtotal: 0,
                          discount: 0,
                          service: 0,
                          tax: 0,
                          total: 0
                        }
                      }, 'Diagnostic Slip');
                    }}
                    className="px-3 py-1.5 bg-white border border-slate-200 text-slate-700 hover:bg-slate-100 rounded-lg text-xs font-bold flex items-center gap-1.5 cursor-pointer"
                  >
                    <Printer className="h-3.5 w-3.5 text-slate-500" />
                    <span>Test Slip</span>
                  </button>
                </div>
              </div>
            </div>

            {/* SECTION 4: AUTOMATED CASH DRAWER SOLENOID */}
            <div className="bg-white rounded-2xl border border-slate-200 p-6 shadow-xs space-y-4">
              <div className="flex items-center justify-between border-b border-slate-100 pb-3">
                <div className="flex items-center gap-2">
                  <div className="p-2 bg-emerald-50 text-emerald-600 rounded-xl">
                    <Zap className="h-4 w-4" />
                  </div>
                  <div>
                    <h3 className="text-xs font-black uppercase tracking-wider text-slate-900">
                      Automated Cash Drawer Solenoid
                    </h3>
                    <p className="text-[10px] text-slate-400">Triggers physical RJ11/RJ12 drawer pop via printer kick pulse</p>
                  </div>
                </div>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">Auto Drawer Kick</label>
                  <select
                    value={settings.autoDrawerKick}
                    onChange={e => setSettings(prev => ({ ...prev, autoDrawerKick: e.target.value }))}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-bold text-slate-900 focus:bg-white focus:outline-none focus:border-[#ff5500]"
                  >
                    <option value="ENABLED">Enabled (Auto-Pop on Payment)</option>
                    <option value="DISABLED">Disabled (Manual Key Only)</option>
                  </select>
                </div>

                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">Drawer Kick Trigger</label>
                  <select
                    value={settings.drawerKickTrigger}
                    onChange={e => setSettings(prev => ({ ...prev, drawerKickTrigger: e.target.value }))}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-bold text-slate-900 focus:bg-white focus:outline-none focus:border-[#ff5500]"
                  >
                    <option value="CASH_ONLY">Cash Payments Only</option>
                    <option value="ALL">All Payments (Cash, Card &amp; Split)</option>
                  </select>
                </div>

                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">RJ11 / RJ12 Pinout</label>
                  <select
                    value={settings.drawerPinout || 'PIN_2'}
                    onChange={e => setSettings(prev => ({ ...prev, drawerPinout: e.target.value }))}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-bold text-slate-900 focus:bg-white focus:outline-none focus:border-[#ff5500]"
                  >
                    <option value="PIN_2">Pin 2 / ESC p 0 (Epson, Rongta, Xprinter)</option>
                    <option value="PIN_5">Pin 5 / ESC p 1 (Star Micronics, Custom)</option>
                  </select>
                </div>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 pt-1">
                <div className="flex items-center justify-between p-3 bg-slate-50 border border-slate-200 rounded-xl">
                  <div>
                    <span className="text-xs font-bold text-slate-800">Register Chime Sound</span>
                    <p className="text-[10px] text-slate-500">Plays brass bell tone upon successful payment settlement</p>
                  </div>
                  <div className="flex items-center gap-1.5">
                    <button
                      type="button"
                      onClick={() => setSettings(prev => ({ ...prev, chimeAudio: !prev.chimeAudio }))}
                      className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all cursor-pointer ${
                        settings.chimeAudio
                          ? 'bg-emerald-100 text-emerald-800 border border-emerald-300'
                          : 'bg-slate-200 text-slate-600'
                      }`}
                    >
                      {settings.chimeAudio ? 'Chime ON' : 'Chime OFF'}
                    </button>
                    <button
                      type="button"
                      onClick={playCashRegisterChime}
                      className="p-1.5 text-slate-500 hover:text-slate-900 rounded-lg hover:bg-slate-200 cursor-pointer"
                      title="Test Audio Chime"
                    >
                      <Volume2 className="h-4 w-4" />
                    </button>
                  </div>
                </div>

                <div className="flex items-center justify-between p-3 bg-slate-50 border border-slate-200 rounded-xl">
                  <div>
                    <span className="text-xs font-bold text-slate-800">Manual Solenoid Kick Test</span>
                    <p className="text-[10px] text-slate-500">Fires test pulse without creating a transaction</p>
                  </div>
                  <button
                    type="button"
                    onClick={() => {
                      if (settings.chimeAudio) playCashRegisterChime();
                      setSettingsNotice({
                        title: 'Cash Drawer Pulse Fired',
                        detail: `Trigger pulse sent via ${settings.drawerPinout || 'Pin 2'} • Register chime sounded.`
                      });
                      setTimeout(() => setSettingsNotice(null), 3500);
                    }}
                    className="px-3.5 py-1.5 bg-[#ff5500] hover:bg-orange-600 text-white rounded-lg text-xs font-bold flex items-center gap-1.5 shadow-xs cursor-pointer"
                  >
                    <Zap className="h-3.5 w-3.5" />
                    <span>Pop Drawer</span>
                  </button>
                </div>
              </div>
            </div>

            {/* SECTION 5: TAX, SERVICE CHARGE & CURRENCY */}
            <div className="bg-white rounded-2xl border border-slate-200 p-6 shadow-xs space-y-4">
              <div className="flex items-center justify-between border-b border-slate-100 pb-3">
                <div className="flex items-center gap-2">
                  <div className="p-2 bg-amber-50 text-amber-600 rounded-xl">
                    <Percent className="h-4 w-4" />
                  </div>
                  <div>
                    <h3 className="text-xs font-black uppercase tracking-wider text-slate-900">
                      Currency, Taxes &amp; Surcharge Rates
                    </h3>
                    <p className="text-[10px] text-slate-400">Default rates applied across tables and receipts</p>
                  </div>
                </div>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">Currency Symbol / Code</label>
                  <input
                    type="text"
                    value={settings.currency}
                    onChange={e => setSettings(prev => ({ ...prev, currency: e.target.value }))}
                    placeholder="e.g. Rs., $, €, LKR"
                    className="w-full px-3.5 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-mono font-bold text-slate-900 focus:bg-white focus:outline-none focus:border-[#ff5500]"
                  />
                  <span className="text-[10px] text-slate-400 mt-1 block">Displayed on menu, POS &amp; receipts</span>
                </div>

                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">Default Service Charge (%)</label>
                  <input
                    type="number"
                    step="0.5"
                    min="0"
                    value={settings.serviceChargeRate}
                    onChange={e => setSettings(prev => ({ ...prev, serviceChargeRate: parseFloat(e.target.value) || 0 }))}
                    className="w-full px-3.5 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-mono font-bold text-slate-900 focus:bg-white focus:outline-none focus:border-[#ff5500]"
                  />
                  <span className="text-[10px] text-slate-400 mt-1 block">Staff gratuity pool</span>
                </div>

                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">Sales Tax / VAT Rate (%)</label>
                  <input
                    type="number"
                    step="0.5"
                    min="0"
                    value={settings.taxRate}
                    onChange={e => setSettings(prev => ({ ...prev, taxRate: parseFloat(e.target.value) || 0 }))}
                    className="w-full px-3.5 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-mono font-bold text-slate-900 focus:bg-white focus:outline-none focus:border-[#ff5500]"
                  />
                  <span className="text-[10px] text-slate-400 mt-1 block">Statutory tax rate</span>
                </div>
              </div>

              <div className="grid grid-cols-1 md:grid-cols-2 gap-4 pt-2">
                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">Thermal Receipt Header Notes</label>
                  <textarea
                    rows={3}
                    value={settings.receiptHeader}
                    onChange={e => setSettings(prev => ({ ...prev, receiptHeader: e.target.value }))}
                    placeholder="Address, phone, tax reg line"
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-medium text-slate-900 focus:bg-white focus:outline-none focus:border-[#ff5500]"
                  />
                </div>

                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">Thermal Receipt Footer Message</label>
                  <textarea
                    rows={3}
                    value={settings.receiptFooter}
                    onChange={e => setSettings(prev => ({ ...prev, receiptFooter: e.target.value }))}
                    placeholder="Thank you message, wifi password, or return policy"
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-medium text-slate-900 focus:bg-white focus:outline-none focus:border-[#ff5500]"
                  />
                </div>
              </div>
            </div>

            {/* SECTION 6: SYSTEM DATABASE BACKUP & RESTORE */}
            <div className="bg-white rounded-2xl border border-slate-200 p-6 shadow-xs space-y-4">
              <div className="flex items-center justify-between border-b border-slate-100 pb-3">
                <div className="flex items-center gap-2">
                  <div className="p-2 bg-purple-50 text-purple-600 rounded-xl">
                    <HardDrive className="h-4 w-4" />
                  </div>
                  <div>
                    <h3 className="text-xs font-black uppercase tracking-wider text-slate-900">
                      System Database Backup &amp; Disaster Recovery
                    </h3>
                    <p className="text-[10px] text-slate-400">Export or restore full store database (Menu, Staff, Inventory, Shifts &amp; Sales)</p>
                  </div>
                </div>
              </div>

              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div className="p-4 bg-slate-50 border border-slate-200 rounded-xl flex flex-col justify-between space-y-3">
                  <div>
                    <span className="text-xs font-bold text-slate-900 flex items-center gap-1.5">
                      <Download className="h-4 w-4 text-[#ff5500]" /> Export JSON Database Backup
                    </span>
                    <p className="text-[11px] text-slate-500 mt-1">
                      Download a complete snapshot of all {menuItems.length} dishes, {inventory.length} ingredients, staff credentials, and {transactions.length} sales records.
                    </p>
                  </div>
                  <button
                    type="button"
                    onClick={handleExportBackup}
                    className="w-full py-2 bg-slate-900 hover:bg-slate-800 text-white font-bold rounded-lg text-xs flex items-center justify-center gap-1.5 cursor-pointer"
                  >
                    <Download className="h-3.5 w-3.5" /> Download System Backup (.json)
                  </button>
                </div>

                <div className="p-4 bg-slate-50 border border-slate-200 rounded-xl flex flex-col justify-between space-y-3">
                  <div>
                    <span className="text-xs font-bold text-slate-900 flex items-center gap-1.5">
                      <Upload className="h-4 w-4 text-emerald-600" /> Restore System from Backup File
                    </span>
                    <p className="text-[11px] text-slate-500 mt-1">
                      Upload a previously exported `.json` file to restore settings, inventory levels, menus, and transaction history.
                    </p>
                  </div>

                  <label className="w-full py-2 bg-white hover:bg-slate-100 border border-slate-300 text-slate-800 font-bold rounded-lg text-xs flex items-center justify-center gap-1.5 cursor-pointer transition-colors">
                    <Upload className="h-3.5 w-3.5 text-slate-600" />
                    <span>Select Backup File (.json)</span>
                    <input
                      type="file"
                      accept=".json,application/json"
                      onChange={handleImportBackup}
                      className="hidden"
                    />
                  </label>
                </div>
              </div>
            </div>

            {/* SECTION 7: ADMIN FACTORY RESET / WIPE TEST DATA */}
            {currentUser.role === 'Administrator' && (
              <div className="bg-rose-50/60 rounded-2xl border border-rose-200 p-6 shadow-xs space-y-3">
                <div className="flex items-center justify-between">
                  <div>
                    <h3 className="text-xs font-black uppercase tracking-wider text-rose-900 flex items-center gap-1.5">
                      <AlertTriangle className="h-4 w-4 text-rose-600" />
                      Administrator Data Purge (Reset Test Data)
                    </h3>
                    <p className="text-[11px] text-rose-700 mt-0.5">
                      Clear test transactions, reset all tables to VACANT, and start with a clean ledger without overriding the cash float.
                    </p>
                  </div>

                  <button
                    type="button"
                    onClick={() => {
                      if (!window.confirm("Are you sure you want to purge all test transactions and reset tables to VACANT? Active cash float and note counts will be preserved.")) {
                        return;
                      }

                      setTransactions([]);
                      setActiveOrders([]);
                      setCancelledTickets([]);
                      setAuditLogs([]);
                      setFloorTables(prev => prev.map(t => ({ ...t, status: 'VACANT', currentOrderRef: null })));
                      
                      // Keep current shift float intact rather than forcing a hardcoded number
                      setCurrentShift(prev => ({
                        shiftId: `SHIFT-${getLocalDateStr().replace(/-/g, '')}-01`,
                        openedDate: getLocalDateStr(),
                        openedAt: new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }),
                        openedBy: currentUser.name,
                        startingFloat: Number(prev.startingFloat) || 0.00,
                        status: 'OPEN',
                        payouts: []
                      }));

                      recordAuditLog('ADMIN_RESET_LEDGER', 'ALL', `Administrator ${currentUser.name} purged all sales and reset tables.`);
                      setSettingsNotice({
                        title: 'All Invoices & Orders Cleared',
                        detail: 'Ledger cleared, floor tables set to VACANT, active cash float retained.'
                      });
                      setTimeout(() => setSettingsNotice(null), 4000);
                    }}
                    className="px-3.5 py-2 bg-rose-600 hover:bg-rose-700 text-white font-bold rounded-xl text-xs flex items-center gap-1.5 shadow-xs shrink-0 cursor-pointer"
                  >
                    <Trash2 className="h-3.5 w-3.5" />
                    <span>Purge Test Records</span>
                  </button>
                </div>
              </div>
            )}
          </div>
        )}

      </main>

      {/* MODAL: EDIT ACTIVE BILL (Billing Queue) */}
      {editBillModalOpen && editingBill && (
        <div className="fixed inset-0 bg-black/75 backdrop-blur-xs z-50 flex items-center justify-center p-4">
          <div className="bg-white rounded-3xl w-full max-w-lg p-6 shadow-2xl border border-slate-200 max-h-[90vh] overflow-y-auto">
            <div className="flex items-center justify-between pb-3 border-b border-slate-100">
              <div>
                <h3 className="text-base font-black text-slate-900">Edit Active Bill: {editingBill.tableName}</h3>
                <p className="text-xs text-slate-500 font-mono">Order #{editingBill.orderId}</p>
              </div>
              <button onClick={() => setEditBillModalOpen(false)} className="text-slate-400 hover:text-slate-900 cursor-pointer">
                <X className="h-5 w-5" />
              </button>
            </div>

            <div className="mt-4 space-y-4">
              <div className="p-3 bg-slate-50 border border-slate-200 rounded-2xl space-y-2">
                <span className="text-xs font-bold text-slate-800">Add Item to Bill</span>
                <div className="flex gap-2">
                  <select
                    id="addDishToBillSelect"
                    className="flex-1 px-3 py-1.5 border border-slate-200 rounded-xl text-xs bg-white text-slate-900"
                  >
                    {menuItems.map(dish => (
                      <option key={dish.id} value={dish.id}>{dish.name} - {settings.currency} {dish.price.toFixed(2)}</option>
                    ))}
                  </select>
                  <button
                    type="button"
                    onClick={() => {
                      const selectEl = document.getElementById('addDishToBillSelect');
                      const selectedDish = menuItems.find(m => m.id === selectEl?.value);
                      if (selectedDish) {
                        setEditingBill(prev => {
                          const existing = prev.items.find(i => i.id === selectedDish.id);
                          const updatedItems = existing
                            ? prev.items.map(i => i.id === selectedDish.id ? { ...i, qty: i.qty + 1 } : i)
                            : [...prev.items, { ...selectedDish, cartItemId: `bill_${Date.now()}_${Math.random().toString(36).substring(2, 5)}`, qty: 1, notes: '' }];
                          return { ...prev, items: updatedItems };
                        });
                      }
                    }}
                    className="px-3 py-1.5 bg-slate-900 text-white rounded-xl text-xs font-bold hover:bg-slate-800 cursor-pointer"
                  >
                    + Add
                  </button>
                </div>
              </div>

              <div className="space-y-2 max-h-56 overflow-y-auto">
                <span className="text-xs font-black uppercase text-slate-400">Current Items</span>
                {editingBill.items.map(item => (
                  <div key={item.cartItemId || item.id} className="p-2.5 bg-slate-50 rounded-xl border border-slate-200 flex items-center justify-between">
                    <div>
                      <p className="font-bold text-xs text-slate-900">{item.name}</p>
                      <span className="font-mono text-xs text-[#ff5500]">{settings.currency} {(item.price * item.qty).toFixed(2)}</span>
                    </div>
                    <div className="flex items-center gap-2">
                      <button
                        type="button"
                        onClick={() => {
                          setEditingBill(prev => ({
                            ...prev,
                            items: prev.items.map(i => (i.cartItemId || i.id) === (item.cartItemId || item.id) ? { ...i, qty: Math.max(1, i.qty - 1) } : i)
                          }));
                        }}
                        className="h-6 w-6 bg-white border border-slate-200 rounded text-xs font-bold cursor-pointer"
                      >
                        -
                      </button>
                      <span className="font-mono font-bold text-xs">{item.qty}</span>
                      <button
                        type="button"
                        onClick={() => {
                          setEditingBill(prev => ({
                            ...prev,
                            items: prev.items.map(i => (i.cartItemId || i.id) === (item.cartItemId || item.id) ? { ...i, qty: i.qty + 1 } : i)
                          }));
                        }}
                        className="h-6 w-6 bg-white border border-slate-200 rounded text-xs font-bold cursor-pointer"
                      >
                        +
                      </button>
                      <button
                        type="button"
                        onClick={() => {
                          setEditingBill(prev => ({
                            ...prev,
                            items: prev.items.filter(i => (i.cartItemId || i.id) !== (item.cartItemId || item.id))
                          }));
                        }}
                        className="text-slate-400 hover:text-rose-600 p-1 cursor-pointer"
                      >
                        <Trash2 className="h-3.5 w-3.5" />
                      </button>
                    </div>
                  </div>
                ))}
              </div>

              <div className="flex gap-2 pt-2">
                <button
                  type="button"
                  onClick={() => setEditBillModalOpen(false)}
                  className="flex-1 py-2.5 bg-slate-100 text-slate-700 font-bold rounded-xl text-xs hover:bg-slate-200 cursor-pointer"
                >
                  Cancel
                </button>
                <button
                  type="button"
                  onClick={() => {
                    const originalOrder = activeOrders.find(o => o.orderId === editingBill.orderId);
                    const originalItems = originalOrder?.items || [];
                    const updatedItems = editingBill.items || [];
                    const changes = [];
                    const newlyAddedOrIncremented = [];

                    // 1. Detect additions and quantity increases
                    updatedItems.forEach(item => {
                      const prev = originalItems.find(i => (i.cartItemId || i.id) === (item.cartItemId || item.id));
                      if (!prev) {
                        changes.push(`ADDED "${item.name}" (Qty: ${item.qty}, ${settings.currency} ${(item.price * item.qty).toFixed(2)})`);
                        newlyAddedOrIncremented.push({ ...item, qty: item.qty });
                      } else if (item.qty > prev.qty) {
                        const diff = item.qty - prev.qty;
                        changes.push(`INCREASED "${item.name}" (+${diff}, now ${item.qty})`);
                        newlyAddedOrIncremented.push({ ...item, qty: diff });
                      } else if (item.qty < prev.qty) {
                        const diff = prev.qty - item.qty;
                        changes.push(`DECREASED "${item.name}" (-${diff}, now ${item.qty})`);
                      }
                    });

                    // 2. Detect deleted/voided items
                    originalItems.forEach(item => {
                      const stillExists = updatedItems.some(i => (i.cartItemId || i.id) === (item.cartItemId || item.id));
                      if (!stillExists) {
                        changes.push(`REMOVED "${item.name}" (was Qty: ${item.qty}, ${settings.currency} ${(item.price * item.qty).toFixed(2)})`);
                      }
                    });

                    const changeSummary = changes.length > 0 ? changes.join(' | ') : 'No line item quantity modifications';

                    // 3. Save updated bill to state
                    setActiveOrders(prev => prev.map(o => o.orderId === editingBill.orderId ? editingBill : o));

                    // 4. Automatically print KOT and BOT for added/incremented items
                    if (newlyAddedOrIncremented.length > 0 && settings.autoPrintOrder !== false) {
                      const kitchenItems = newlyAddedOrIncremented.filter(i => i.department === 'Kitchen');
                      const barItems = newlyAddedOrIncremented.filter(i => i.department === 'Bar');

                      if (kitchenItems.length > 0 || barItems.length > 0) {
                        triggerAutoPrint({
                          type: 'KOT_BOT_DISPATCH',
                          data: {
                            order: {
                              ...editingBill,
                              sentAt: new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }),
                              isAddon: true
                            },
                            kitchenItems,
                            barItems
                          }
                        }, `${editingBill.tableName} • Add-on Ticket Auto-Printed`);
                      }
                    }

                    // 5. Differential audit logging
                    recordAuditLog(
                      'BILL_MODIFIED_DIFF',
                      editingBill.orderId,
                      `Saved changes on Bill #${editingBill.orderId} (${editingBill.tableName}): ${changeSummary}`
                    );

                    setEditBillModalOpen(false);
                  }}
                  className="flex-1 py-2.5 bg-[#ff5500] hover:bg-orange-600 text-white font-bold rounded-xl text-xs shadow-xs cursor-pointer"
                >
                  Save Changes to Bill
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* EDIT MENU ITEM MODAL */}
      {isEditModalOpen && editingMenuItem && (
        <div className="fixed inset-0 bg-black/75 backdrop-blur-xs z-50 flex items-center justify-center p-4">
          <div className="bg-white rounded-3xl w-full max-w-lg p-6 shadow-2xl border border-slate-200 max-h-[90vh] overflow-y-auto">
            <div className="flex items-center justify-between pb-3 border-b border-slate-100">
              <div>
                <h3 className="text-base font-black text-slate-900">Edit Dish / Item</h3>
                <p className="text-xs text-slate-500">Update dish details, pricing, BOM recipes, and category</p>
              </div>
              <button
                type="button"
                onClick={() => {
                  setIsEditModalOpen(false);
                  setEditingMenuItem(null);
                  setEditDishBomInput({ ingredientId: '', amount: '' });
                }}
                className="text-slate-400 hover:text-slate-900 cursor-pointer"
              >
                <X className="h-5 w-5" />
              </button>
            </div>

            <form
              onSubmit={(e) => {
                e.preventDefault();
                const priceNum = parseFloat(editingMenuItem.price);
                if (!editingMenuItem.name.trim() || isNaN(priceNum) || priceNum <= 0) {
                  alert('Please enter a valid item name and price.');
                  return;
                }

                setMenuItems(prev => prev.map(m => m.id === editingMenuItem.id ? {
                  ...editingMenuItem,
                  price: priceNum,
                  recipe: editingMenuItem.recipe || []
                } : m));

                recordAuditLog('MENU_ITEM_UPDATED', editingMenuItem.id, `Updated ${editingMenuItem.name} to ${settings.currency} ${priceNum.toFixed(2)}`);

                setIsEditModalOpen(false);
                setEditingMenuItem(null);
                setEditDishBomInput({ ingredientId: '', amount: '' });
              }}
              className="mt-4 space-y-4"
            >
              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">Item Name</label>
                  <input
                    type="text"
                    required
                    value={editingMenuItem.name}
                    onChange={(e) => setEditingMenuItem(prev => ({ ...prev, name: e.target.value }))}
                    className="w-full px-3 py-2 border border-slate-200 rounded-xl text-xs font-bold text-slate-900 focus:bg-white focus:outline-none focus:ring-2 focus:ring-orange-500/20"
                  />
                </div>

                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">Selling Price ({settings.currency})</label>
                  <input
                    type="number"
                    step="0.01"
                    min="0"
                    required
                    value={editingMenuItem.price}
                    onChange={(e) => setEditingMenuItem(prev => ({ ...prev, price: e.target.value }))}
                    className="w-full px-3 py-2 border border-slate-200 rounded-xl text-xs font-mono font-bold text-slate-900 focus:bg-white focus:outline-none focus:ring-2 focus:ring-orange-500/20"
                  />
                </div>
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">Department</label>
                  <select
                    value={editingMenuItem.department || 'Kitchen'}
                    onChange={(e) => setEditingMenuItem(prev => ({ ...prev, department: e.target.value }))}
                    className="w-full px-3 py-2 border border-slate-200 rounded-xl text-xs font-bold bg-white text-slate-900 focus:outline-none focus:border-[#ff5500] cursor-pointer"
                  >
                    <option value="Kitchen">Kitchen (Sends KOT)</option>
                    <option value="Bar">Bar (Sends BOT)</option>
                  </select>
                </div>

                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">Category</label>
                  <input
                    type="text"
                    value={editingMenuItem.category || ''}
                    onChange={(e) => setEditingMenuItem(prev => ({ ...prev, category: e.target.value }))}
                    placeholder="e.g. Rice & Curry, Kottu, Starters"
                    className="w-full px-3 py-2 border border-slate-200 rounded-xl text-xs font-bold text-slate-900 focus:bg-white focus:outline-none focus:ring-2 focus:ring-orange-500/20"
                  />
                </div>
              </div>

              <div>
                <label className="block text-xs font-bold text-slate-700 mb-1">Prep Time</label>
                <input
                  type="text"
                  value={editingMenuItem.prepTime || '10m'}
                  onChange={(e) => setEditingMenuItem(prev => ({ ...prev, prepTime: e.target.value }))}
                  placeholder="e.g. 10m"
                  className="w-full px-3 py-2 border border-slate-200 rounded-xl text-xs font-bold text-slate-900 focus:bg-white focus:outline-none focus:ring-2 focus:ring-orange-500/20"
                />
              </div>

              <div>
                <label className="block text-xs font-bold text-slate-700 mb-1">Photo (Upload File or Enter URL)</label>
                <div className="space-y-2">
                  <div className="flex gap-2 items-center">
                    <label className="cursor-pointer px-3 py-2 bg-slate-100 hover:bg-slate-200 border border-slate-200 text-slate-700 font-bold rounded-xl text-xs flex items-center gap-1.5 transition-colors">
                      <Upload className="h-4 w-4" />
                      <span>Upload Photo</span>
                      <input
                        type="file"
                        accept="image/*"
                        className="hidden"
                        onChange={(e) => {
                          const file = e.target.files?.[0];
                          if (file) {
                            const reader = new FileReader();
                            reader.onload = ev => {
                              if (ev.target?.result) {
                                setEditingMenuItem(prev => ({ ...prev, imageUrl: ev.target.result }));
                              }
                            };
                            reader.readAsDataURL(file);
                          }
                        }}
                      />
                    </label>

                    {editingMenuItem.imageUrl && (
                      <button
                        type="button"
                        onClick={() => setEditingMenuItem(prev => ({ ...prev, imageUrl: '' }))}
                        className="p-2 text-slate-400 hover:text-rose-600 rounded-xl border border-slate-200 hover:bg-rose-50 transition-colors cursor-pointer"
                        title="Remove Photo"
                      >
                        <Trash2 className="h-4 w-4" />
                      </button>
                    )}
                  </div>

                  <input
                    type="url"
                    value={editingMenuItem.imageUrl || ''}
                    onChange={(e) => setEditingMenuItem(prev => ({ ...prev, imageUrl: e.target.value }))}
                    placeholder="https://..."
                    className="w-full px-3 py-2 border border-slate-200 rounded-xl text-xs text-slate-900 focus:bg-white focus:outline-none"
                  />
                </div>
              </div>

              {/* RECIPE INGREDIENT BOM BUILDER (FIXED) */}
              <div className="p-3 bg-slate-50 border border-slate-200 rounded-2xl space-y-2">
                <span className="text-xs font-bold text-slate-800 flex items-center gap-1.5">
                  <BookOpen className="h-3.5 w-3.5 text-[#ff5500]" /> Link Recipe Ingredients (BOM)
                </span>
                
                <div className="flex gap-2">
                  <select
                    value={editDishBomInput.ingredientId || (inventory[0]?.id || '')}
                    onChange={(e) => setEditDishBomInput(prev => ({ ...prev, ingredientId: e.target.value }))}
                    className="flex-1 px-2.5 py-1.5 border border-slate-200 rounded-xl text-xs bg-white text-slate-900 focus:outline-none focus:border-[#ff5500]"
                    style={{ color: '#0f172a', backgroundColor: '#ffffff' }}
                  >
                    {inventory.map(ing => (
                      <option key={ing.id} value={ing.id} style={{ color: '#0f172a', backgroundColor: '#ffffff' }}>
                        {ing.name} ({ing.unit})
                      </option>
                    ))}
                  </select>

                  <input
                    type="number"
                    min="0.01"
                    step="any"
                    value={editDishBomInput.amount}
                    onChange={(e) => setEditDishBomInput(prev => ({ ...prev, amount: e.target.value }))}
                    placeholder="Qty/portion"
                    className="w-28 px-2.5 py-1.5 border border-slate-200 rounded-xl text-xs font-mono bg-white text-slate-900 focus:outline-none focus:border-[#ff5500]"
                  />

                  <button
                    type="button"
                    onClick={(e) => {
                      e.preventDefault();
                      e.stopPropagation();

                      const ingId = editDishBomInput.ingredientId || inventory[0]?.id;
                      const amt = parseFloat(editDishBomInput.amount);

                      if (!ingId) {
                        alert('Please select an ingredient.');
                        return;
                      }
                      if (isNaN(amt) || amt <= 0) {
                        alert('Please enter a valid amount greater than 0.');
                        return;
                      }

                      setEditingMenuItem(prev => {
                        const currentRecipe = Array.isArray(prev.recipe) ? [...prev.recipe] : [];
                        const existingIdx = currentRecipe.findIndex(r => r.ingredientId === ingId);

                        if (existingIdx >= 0) {
                          currentRecipe[existingIdx] = {
                            ...currentRecipe[existingIdx],
                            amount: Number((currentRecipe[existingIdx].amount + amt).toFixed(2))
                          };
                          return { ...prev, recipe: currentRecipe };
                        } else {
                          return { ...prev, recipe: [...currentRecipe, { ingredientId: ingId, amount: amt }] };
                        }
                      });

                      setEditDishBomInput(prev => ({ ...prev, amount: '' }));
                    }}
                    className="px-4 py-1.5 bg-slate-900 hover:bg-slate-800 text-white rounded-xl text-xs font-bold transition-all cursor-pointer shrink-0 active:scale-95"
                  >
                    + Add
                  </button>
                </div>

                {editingMenuItem.recipe && editingMenuItem.recipe.length > 0 ? (
                  <div className="space-y-1.5 max-h-36 overflow-y-auto pt-1">
                    {editingMenuItem.recipe.map((r, i) => {
                      const matchedIng = inventory.find(inv => inv.id === r.ingredientId);
                      const unitCost = matchedIng ? (matchedIng.cost * r.amount) : 0;
                      return (
                        <div key={`${r.ingredientId}_${i}`} className="flex justify-between items-center text-xs p-2 bg-white rounded-xl border border-slate-200">
                          <div>
                            <span className="font-bold text-slate-900">{matchedIng?.name || r.ingredientId}</span>
                            <span className="text-[10px] text-slate-400 block font-mono">
                              Cost: {settings.currency} {unitCost.toFixed(2)}
                            </span>
                          </div>
                          <div className="flex items-center gap-2">
                            <span className="font-mono font-bold text-slate-700 bg-slate-100 px-2 py-0.5 rounded-lg text-xs">
                              {r.amount} {matchedIng?.unit || 'units'}
                            </span>
                            <button
                              type="button"
                              onClick={(e) => {
                                e.preventDefault();
                                e.stopPropagation();
                                setEditingMenuItem(prev => ({
                                  ...prev,
                                  recipe: prev.recipe.filter((_, idx) => idx !== i)
                                }));
                              }}
                              className="text-slate-400 hover:text-rose-600 p-1 rounded hover:bg-rose-50 transition-colors cursor-pointer"
                              title="Remove ingredient"
                            >
                              <X className="h-3.5 w-3.5" />
                            </button>
                          </div>
                        </div>
                      );
                    })}
                  </div>
                ) : (
                  <p className="text-[11px] text-slate-400 italic pt-1">No ingredients linked yet. Select an ingredient, enter qty, and click &ldquo;+ Add&rdquo;.</p>
                )}
              </div>

              <div className="flex gap-2 pt-2 border-t border-slate-100">
                <button
                  type="button"
                  onClick={() => {
                    setIsEditModalOpen(false);
                    setEditingMenuItem(null);
                    setEditDishBomInput({ ingredientId: '', amount: '' });
                  }}
                  className="flex-1 py-2.5 bg-slate-100 text-slate-700 font-bold rounded-xl text-xs hover:bg-slate-200 cursor-pointer"
                >
                  Cancel
                </button>
                <button
                  type="submit"
                  className="flex-1 py-2.5 bg-[#ff5500] hover:bg-orange-600 text-white font-bold rounded-xl text-xs shadow-xs cursor-pointer"
                >
                  Save Changes
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* MODAL: CHECKOUT & SETTLEMENT */}
      {checkoutModalOpen && (
        <div className="fixed inset-0 bg-black/75 backdrop-blur-xs z-50 flex items-center justify-center p-4">
          <div className="bg-white rounded-3xl w-full max-w-md p-6 shadow-2xl border border-slate-200">
            <div className="flex items-center justify-between pb-3 border-b border-slate-100">
              <h3 className="text-base font-black text-slate-900">
                Settle Bill: {settlingOrder ? settlingOrder.tableName : selectedTable.name}
              </h3>
              <button onClick={() => setCheckoutModalOpen(false)} className="text-slate-400 hover:text-slate-900 cursor-pointer">
                <X className="h-5 w-5" />
              </button>
            </div>

            {(() => {
              const currentRef = settlingOrder || {
                items: cart,
                serviceChargeActive,
                taxActive,
                discountPercent
              };
              const fin = calculateOrderFinancials(currentRef);

              return (
                <div className="mt-4 space-y-4">
                  <div className="grid grid-cols-3 gap-2">
                    {['CASH', 'CARD', 'SPLIT'].map(type => (
                      <button
                        key={type}
                        type="button"
                        onClick={() => setPaymentMethod(type)}
                        className={`py-2 rounded-xl text-xs font-bold border transition-all cursor-pointer ${
                          paymentMethod === type
                            ? 'bg-[#ff5500] text-white border-[#ff5500]'
                            : 'bg-slate-50 border-slate-200 text-slate-600'
                        }`}
                      >
                        {type}
                      </button>
                    ))}
                  </div>

                  {paymentMethod === 'CASH' && (
                    <div className="space-y-2">
                      <div className="flex items-center justify-between">
                        <label className="block text-xs font-bold text-slate-700">Cash Tendered ({settings.currency})</label>
                        <button
                          type="button"
                          onClick={() => setCashTendered(fin.total.toFixed(2))}
                          className="text-[10px] font-bold text-[#ff5500] hover:underline cursor-pointer"
                        >
                          Exact Amount
                        </button>
                      </div>
                      <input
                        type="number"
                        step="0.01"
                        value={cashTendered}
                        onChange={e => setCashTendered(e.target.value)}
                        placeholder={fin.total.toFixed(2)}
                        className="w-full px-3 py-2 border border-slate-200 rounded-xl text-slate-900 font-mono text-sm focus:outline-none focus:border-[#ff5500]"
                      />

                      <div className="flex gap-1.5 overflow-x-auto pt-0.5">
                        {[
                          Math.ceil(fin.total / 100) * 100,
                          Math.ceil(fin.total / 500) * 500,
                          Math.ceil(fin.total / 1000) * 1000,
                          5000
                        ]
                          .filter((val, idx, arr) => val >= fin.total && arr.indexOf(val) === idx)
                          .slice(0, 3)
                          .map(val => (
                            <button
                              key={val}
                              type="button"
                              onClick={() => setCashTendered(val.toString())}
                              className="px-2.5 py-1 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-lg text-[10px] font-mono font-bold transition-colors cursor-pointer"
                            >
                              {settings.currency} {val}
                            </button>
                          ))}
                      </div>

                      {parseFloat(cashTendered) >= fin.total && (
                        <div className="p-2.5 bg-emerald-50 border border-emerald-200 rounded-xl text-xs font-mono font-bold text-emerald-800 flex justify-between items-center">
                          <span>Balance / Change to Return:</span>
                          <span className="text-sm text-emerald-700">
                            {settings.currency} {(parseFloat(cashTendered) - fin.total).toFixed(2)}
                          </span>
                        </div>
                      )}
                    </div>
                  )}

                  <div className="p-3 bg-slate-50 rounded-2xl border border-slate-200 space-y-1 text-xs">
                    <div className="flex justify-between text-slate-500">
                      <span>Subtotal</span>
                      <span className="font-mono">{settings.currency} {fin.subtotal.toFixed(2)}</span>
                    </div>
                    {fin.discount > 0 && (
                      <div className="flex justify-between text-rose-600">
                        <span>Discount</span>
                        <span className="font-mono">-{settings.currency} {fin.discount.toFixed(2)}</span>
                      </div>
                    )}
                    {fin.service > 0 && (
                      <div className="flex justify-between text-emerald-700 font-medium">
                        <span>Service Charge</span>
                        <span className="font-mono">+{settings.currency} {fin.service.toFixed(2)}</span>
                      </div>
                    )}
                    {fin.tax > 0 && (
                      <div className="flex justify-between text-indigo-700 font-medium">
                        <span>Taxes</span>
                        <span className="font-mono">+{settings.currency} {fin.tax.toFixed(2)}</span>
                      </div>
                    )}
                    <div className="flex justify-between text-sm font-black text-slate-900 pt-2 border-t border-slate-200">
                      <span>Total Due</span>
                      <span className="text-base font-mono text-[#ff5500]">{settings.currency} {fin.total.toFixed(2)}</span>
                    </div>
                  </div>

                  <button
                    type="button"
                    onClick={handleCompleteSettlement}
                    className="w-full py-3 bg-[#008f5d] hover:bg-emerald-700 text-white font-extrabold rounded-xl text-xs uppercase tracking-wider shadow-xs cursor-pointer active:scale-95"
                  >
                    Confirm Settlement &amp; Deduct BOM Stock
                  </button>
                </div>
              );
            })()}
          </div>
        </div>
      )}

      {/* STREAMLINED NON-BLOCKING PRINT NOTIFICATION TOAST */}
      {printNotice && (
        <div className="fixed bottom-6 right-6 z-50 bg-slate-900/95 text-white px-4 py-3 rounded-2xl shadow-2xl border border-slate-700 flex items-center gap-3 backdrop-blur-md transition-all">
          <div className="h-9 w-9 rounded-xl bg-orange-500/20 text-[#ff5500] flex items-center justify-center shrink-0">
            <Printer className="h-5 w-5 animate-pulse" />
          </div>
          <div>
            <p className="text-xs font-bold text-white">{printNotice.title}</p>
            <p className="text-[10px] text-slate-400 font-mono mt-0.5">{printNotice.detail}</p>
          </div>
        </div>
      )}

     {/* MODAL: EDIT ACTIVE BILL (Billing Queue) */}
      {editBillModalOpen && editingBill && (
        <div className="fixed inset-0 bg-black/75 backdrop-blur-xs z-50 flex items-center justify-center p-4">
          <div className="bg-white rounded-3xl w-full max-w-lg p-6 shadow-2xl border border-slate-200 max-h-[90vh] overflow-y-auto">
            <div className="flex items-center justify-between pb-3 border-b border-slate-100">
              <div>
                <h3 className="text-base font-black text-slate-900">Edit Active Bill: {editingBill.tableName}</h3>
                <p className="text-xs text-slate-500 font-mono">Order #{editingBill.orderId}</p>
              </div>
              <button onClick={() => setEditBillModalOpen(false)} className="text-slate-400 hover:text-slate-900 cursor-pointer">
                <X className="h-5 w-5" />
              </button>
            </div>

            <div className="mt-4 space-y-4">
              <div className="p-3 bg-slate-50 border border-slate-200 rounded-2xl space-y-2">
                <span className="text-xs font-bold text-slate-800">Add Item to Bill</span>
                <div className="flex gap-2">
                  <select
                    id="addDishToBillSelect"
                    className="flex-1 px-3 py-1.5 border border-slate-200 rounded-xl text-xs bg-white text-slate-900"
                  >
                    {menuItems.map(dish => (
                      <option key={dish.id} value={dish.id}>{dish.name} - {settings.currency} {dish.price.toFixed(2)}</option>
                    ))}
                  </select>
                  <button
                    type="button"
                    onClick={() => {
                      const selectEl = document.getElementById('addDishToBillSelect');
                      const selectedDish = menuItems.find(m => m.id === selectEl?.value);
                      if (selectedDish) {
                        setEditingBill(prev => {
                          const existing = prev.items.find(i => i.id === selectedDish.id);
                          const updatedItems = existing
                            ? prev.items.map(i => i.id === selectedDish.id ? { ...i, qty: i.qty + 1 } : i)
                            : [...prev.items, { ...selectedDish, cartItemId: `bill_${Date.now()}_${Math.random().toString(36).substring(2, 5)}`, qty: 1, notes: '' }];
                          return { ...prev, items: updatedItems };
                        });
                      }
                    }}
                    className="px-3 py-1.5 bg-slate-900 text-white rounded-xl text-xs font-bold hover:bg-slate-800 cursor-pointer"
                  >
                    + Add
                  </button>
                </div>
              </div>

              <div className="space-y-2 max-h-56 overflow-y-auto">
                <span className="text-xs font-black uppercase text-slate-400">Current Items</span>
                {editingBill.items.map(item => (
                  <div key={item.cartItemId || item.id} className="p-2.5 bg-slate-50 rounded-xl border border-slate-200 flex items-center justify-between">
                    <div>
                      <p className="font-bold text-xs text-slate-900">{item.name}</p>
                      <span className="font-mono text-xs text-[#ff5500]">{settings.currency} {(item.price * item.qty).toFixed(2)}</span>
                    </div>
                    <div className="flex items-center gap-2">
                      <button
                        type="button"
                        onClick={() => {
                          setEditingBill(prev => ({
                            ...prev,
                            items: prev.items.map(i => (i.cartItemId || i.id) === (item.cartItemId || item.id) ? { ...i, qty: Math.max(1, i.qty - 1) } : i)
                          }));
                        }}
                        className="h-6 w-6 bg-white border border-slate-200 rounded text-xs font-bold cursor-pointer"
                      >
                        -
                      </button>
                      <span className="font-mono font-bold text-xs">{item.qty}</span>
                      <button
                        type="button"
                        onClick={() => {
                          setEditingBill(prev => ({
                            ...prev,
                            items: prev.items.map(i => (i.cartItemId || i.id) === (item.cartItemId || item.id) ? { ...i, qty: i.qty + 1 } : i)
                          }));
                        }}
                        className="h-6 w-6 bg-white border border-slate-200 rounded text-xs font-bold cursor-pointer"
                      >
                        +
                      </button>
                      <button
                        type="button"
                        onClick={() => {
                          setEditingBill(prev => ({
                            ...prev,
                            items: prev.items.filter(i => (i.cartItemId || i.id) !== (item.cartItemId || item.id))
                          }));
                        }}
                        className="text-slate-400 hover:text-rose-600 p-1 cursor-pointer"
                      >
                        <Trash2 className="h-3.5 w-3.5" />
                      </button>
                    </div>
                  </div>
                ))}
              </div>

              <div className="flex gap-2 pt-2">
                <button
                  type="button"
                  onClick={() => setEditBillModalOpen(false)}
                  className="flex-1 py-2.5 bg-slate-100 text-slate-700 font-bold rounded-xl text-xs hover:bg-slate-200 cursor-pointer"
                >
                  Cancel
                </button>
                <button
                  type="button"
                  onClick={() => {
                    const originalOrder = activeOrders.find(o => o.orderId === editingBill.orderId);
                    const originalItems = originalOrder?.items || [];
                    const updatedItems = editingBill.items || [];
                    const changes = [];
                    const newlyAddedOrIncremented = [];

                    // 1. Detect additions and quantity increases
                    updatedItems.forEach(item => {
                      const prev = originalItems.find(i => (i.cartItemId || i.id) === (item.cartItemId || item.id));
                      if (!prev) {
                        changes.push(`ADDED "${item.name}" (Qty: ${item.qty}, ${settings.currency} ${(item.price * item.qty).toFixed(2)})`);
                        newlyAddedOrIncremented.push({ ...item, qty: item.qty });
                      } else if (item.qty > prev.qty) {
                        const diff = item.qty - prev.qty;
                        changes.push(`INCREASED "${item.name}" (+${diff}, now ${item.qty})`);
                        newlyAddedOrIncremented.push({ ...item, qty: diff });
                      } else if (item.qty < prev.qty) {
                        const diff = prev.qty - item.qty;
                        changes.push(`DECREASED "${item.name}" (-${diff}, now ${item.qty})`);
                      }
                    });

                    // 2. Detect removed items
                    originalItems.forEach(item => {
                      const stillExists = updatedItems.some(i => (i.cartItemId || i.id) === (item.cartItemId || item.id));
                      if (!stillExists) {
                        changes.push(`REMOVED "${item.name}" (was Qty: ${item.qty}, ${settings.currency} ${(item.price * item.qty).toFixed(2)})`);
                      }
                    });

                    const changeSummary = changes.length > 0 ? changes.join(' | ') : 'No line item quantity modifications';

                    // 3. Save updated bill to state
                    setActiveOrders(prev => prev.map(o => o.orderId === editingBill.orderId ? editingBill : o));

                    // 4. Auto-print supplementary KOT/BOT tickets if new items or increased counts exist
                    if (newlyAddedOrIncremented.length > 0 && settings.autoPrintOrder !== false) {
                      const kitchenItems = newlyAddedOrIncremented.filter(i => i.department === 'Kitchen');
                      const barItems = newlyAddedOrIncremented.filter(i => i.department === 'Bar');

                      if (kitchenItems.length > 0 || barItems.length > 0) {
                        triggerAutoPrint({
                          type: 'KOT_BOT_DISPATCH',
                          data: {
                            order: {
                              ...editingBill,
                              sentAt: new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }),
                              isAddon: true
                            },
                            kitchenItems,
                            barItems
                          }
                        }, `${editingBill.tableName} • Add-on KOT/BOT Printed`);
                      }
                    }

                    // 5. Differential audit logging
                    recordAuditLog(
                      'BILL_MODIFIED_DIFF',
                      editingBill.orderId,
                      `Saved changes on Bill #${editingBill.orderId} (${editingBill.tableName}): ${changeSummary}`
                    );

                    setEditBillModalOpen(false);
                  }}
                  className="flex-1 py-2.5 bg-[#ff5500] hover:bg-orange-600 text-white font-bold rounded-xl text-xs shadow-xs cursor-pointer"
                >
                  Save Changes to Bill
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* EDIT MENU ITEM MODAL */}
      {isEditModalOpen && editingMenuItem && (
        <div className="fixed inset-0 bg-black/75 backdrop-blur-xs z-50 flex items-center justify-center p-4">
          <div className="bg-white rounded-3xl w-full max-w-lg p-6 shadow-2xl border border-slate-200 max-h-[90vh] overflow-y-auto">
            <div className="flex items-center justify-between pb-3 border-b border-slate-100">
              <div>
                <h3 className="text-base font-black text-slate-900">Edit Dish / Item</h3>
                <p className="text-xs text-slate-500">Update dish details, pricing, BOM recipes, and category</p>
              </div>
              <button
                type="button"
                onClick={() => {
                  setIsEditModalOpen(false);
                  setEditingMenuItem(null);
                }}
                className="text-slate-400 hover:text-slate-900 cursor-pointer"
              >
                <X className="h-5 w-5" />
              </button>
            </div>

            <form
              onSubmit={(e) => {
                e.preventDefault();
                const priceNum = parseFloat(editingMenuItem.price);
                if (!editingMenuItem.name.trim() || isNaN(priceNum) || priceNum <= 0) {
                  alert('Please enter a valid item name and price.');
                  return;
                }

                setMenuItems(prev => prev.map(m => m.id === editingMenuItem.id ? {
                  ...editingMenuItem,
                  price: priceNum,
                  recipe: editingMenuItem.recipe || []
                } : m));

                recordAuditLog('MENU_ITEM_UPDATED', editingMenuItem.id, `Updated ${editingMenuItem.name} to ${settings.currency} ${priceNum.toFixed(2)}`);

                setIsEditModalOpen(false);
                setEditingMenuItem(null);
              }}
              className="mt-4 space-y-4"
            >
              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">Item Name</label>
                  <input
                    type="text"
                    required
                    value={editingMenuItem.name}
                    onChange={(e) => setEditingMenuItem(prev => ({ ...prev, name: e.target.value }))}
                    className="w-full px-3 py-2 border border-slate-200 rounded-xl text-xs font-bold text-slate-900 focus:bg-white focus:outline-none focus:ring-2 focus:ring-orange-500/20"
                  />
                </div>

                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">Selling Price ({settings.currency})</label>
                  <input
                    type="number"
                    step="0.01"
                    min="0"
                    required
                    value={editingMenuItem.price}
                    onChange={(e) => setEditingMenuItem(prev => ({ ...prev, price: e.target.value }))}
                    className="w-full px-3 py-2 border border-slate-200 rounded-xl text-xs font-mono font-bold text-slate-900 focus:bg-white focus:outline-none focus:ring-2 focus:ring-orange-500/20"
                  />
                </div>
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">Department</label>
                  <select
                    value={editingMenuItem.department || 'Kitchen'}
                    onChange={(e) => setEditingMenuItem(prev => ({ ...prev, department: e.target.value }))}
                    className="w-full px-3 py-2 border border-slate-200 rounded-xl text-xs font-bold bg-white text-slate-900 focus:outline-none focus:border-[#ff5500] cursor-pointer"
                  >
                    <option value="Kitchen">Kitchen (Sends KOT)</option>
                    <option value="Bar">Bar (Sends BOT)</option>
                  </select>
                </div>

                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">Category</label>
                  <input
                    type="text"
                    value={editingMenuItem.category || ''}
                    onChange={(e) => setEditingMenuItem(prev => ({ ...prev, category: e.target.value }))}
                    placeholder="e.g. Rice & Curry, Kottu, Starters"
                    className="w-full px-3 py-2 border border-slate-200 rounded-xl text-xs font-bold text-slate-900 focus:bg-white focus:outline-none focus:ring-2 focus:ring-orange-500/20"
                  />
                </div>
              </div>

              <div>
                <label className="block text-xs font-bold text-slate-700 mb-1">Prep Time</label>
                <input
                  type="text"
                  value={editingMenuItem.prepTime || '10m'}
                  onChange={(e) => setEditingMenuItem(prev => ({ ...prev, prepTime: e.target.value }))}
                  placeholder="e.g. 10m"
                  className="w-full px-3 py-2 border border-slate-200 rounded-xl text-xs font-bold text-slate-900 focus:bg-white focus:outline-none focus:ring-2 focus:ring-orange-500/20"
                />
              </div>

              <div>
                <label className="block text-xs font-bold text-slate-700 mb-1">Photo (Upload File or Enter URL)</label>
                <div className="space-y-2">
                  <div className="flex gap-2 items-center">
                    <label className="cursor-pointer px-3 py-2 bg-slate-100 hover:bg-slate-200 border border-slate-200 text-slate-700 font-bold rounded-xl text-xs flex items-center gap-1.5 transition-colors">
                      <Upload className="h-4 w-4" />
                      <span>Upload Photo</span>
                      <input
                        type="file"
                        accept="image/*"
                        className="hidden"
                        onChange={(e) => {
                          const file = e.target.files?.[0];
                          if (file) {
                            const reader = new FileReader();
                            reader.onload = ev => {
                              if (ev.target?.result) {
                                setEditingMenuItem(prev => ({ ...prev, imageUrl: ev.target.result }));
                              }
                            };
                            reader.readAsDataURL(file);
                          }
                        }}
                      />
                    </label>

                    {editingMenuItem.imageUrl && (
                      <button
                        type="button"
                        onClick={() => setEditingMenuItem(prev => ({ ...prev, imageUrl: '' }))}
                        className="p-2 text-slate-400 hover:text-rose-600 rounded-xl border border-slate-200 hover:bg-rose-50 transition-colors cursor-pointer"
                        title="Remove Photo"
                      >
                        <Trash2 className="h-4 w-4" />
                      </button>
                    )}
                  </div>

                  <input
                    type="url"
                    value={editingMenuItem.imageUrl || ''}
                    onChange={(e) => setEditingMenuItem(prev => ({ ...prev, imageUrl: e.target.value }))}
                    placeholder="https://..."
                    className="w-full px-3 py-2 border border-slate-200 rounded-xl text-xs text-slate-900 focus:bg-white focus:outline-none"
                  />
                </div>
              </div>

              {/* RECIPE INGREDIENT BOM BUILDER */}
              <div className="p-3 bg-slate-50 border border-slate-200 rounded-2xl space-y-2">
                <span className="text-xs font-bold text-slate-800 flex items-center gap-1.5">
                  <BookOpen className="h-3.5 w-3.5 text-[#ff5500]" /> Link Recipe Ingredients (BOM)
                </span>
                <div className="flex gap-2">
                  <select id="editDishIngSelect" className="flex-1 px-2.5 py-1.5 border border-slate-200 rounded-xl text-xs bg-white text-slate-900">
                    {inventory.map(ing => (
                      <option key={ing.id} value={ing.id}>{ing.name} ({ing.unit})</option>
                    ))}
                  </select>
                  <input
                    id="editDishIngAmount"
                    type="number"
                    min="0.1"
                    step="any"
                    placeholder="Qty/portion"
                    className="w-28 px-2.5 py-1.5 border border-slate-200 rounded-xl text-xs font-mono bg-white text-slate-900"
                  />
                  <button
                    type="button"
                    onClick={() => {
                      const sel = document.getElementById('editDishIngSelect');
                      const amtInput = document.getElementById('editDishIngAmount');
                      const ingId = sel?.value;
                      const amt = parseFloat(amtInput?.value);
                      if (!ingId || isNaN(amt) || amt <= 0) return;

                      setEditingMenuItem(prev => ({
                        ...prev,
                        recipe: [...(prev.recipe || []), { ingredientId: ingId, amount: amt }]
                      }));
                      if (amtInput) amtInput.value = '';
                    }}
                    className="px-3 py-1.5 bg-slate-900 hover:bg-slate-800 text-white rounded-xl text-xs font-bold transition-colors cursor-pointer"
                  >
                    + Add
                  </button>
                </div>

                {editingMenuItem.recipe && editingMenuItem.recipe.length > 0 && (
                  <div className="space-y-1 max-h-28 overflow-y-auto">
                    {editingMenuItem.recipe.map((r, i) => {
                      const matchedIng = inventory.find(inv => inv.id === r.ingredientId);
                      return (
                        <div key={i} className="flex justify-between items-center text-xs p-1.5 bg-white rounded-lg border border-slate-200">
                          <span className="font-bold text-slate-800">{matchedIng?.name || r.ingredientId}</span>
                          <div className="flex items-center gap-2">
                            <span className="font-mono text-slate-500">{r.amount} {matchedIng?.unit}</span>
                            <button
                              type="button"
                              onClick={() => setEditingMenuItem(prev => ({
                                ...prev,
                                recipe: prev.recipe.filter((_, idx) => idx !== i)
                              }))}
                              className="text-slate-400 hover:text-rose-600 transition-colors cursor-pointer"
                            >
                              <X className="h-3.5 w-3.5" />
                            </button>
                          </div>
                        </div>
                      );
                    })}
                  </div>
                )}
              </div>

              <div className="flex gap-2 pt-2 border-t border-slate-100">
                <button
                  type="button"
                  onClick={() => {
                    setIsEditModalOpen(false);
                    setEditingMenuItem(null);
                  }}
                  className="flex-1 py-2.5 bg-slate-100 text-slate-700 font-bold rounded-xl text-xs hover:bg-slate-200 cursor-pointer"
                >
                  Cancel
                </button>
                <button
                  type="submit"
                  className="flex-1 py-2.5 bg-[#ff5500] hover:bg-orange-600 text-white font-bold rounded-xl text-xs shadow-xs cursor-pointer"
                >
                  Save Changes
                </button>
              </div>
            </form>
          </div>
        </div>
      )}
      {/* MODAL: ADD NEW MENU ITEM */}
      {addItemModalOpen && (
        <div 
          className="fixed inset-0 bg-black/80 backdrop-blur-xs z-[9999] flex items-center justify-center p-4 text-slate-900"
          onClick={(e) => {
            if (e.target === e.currentTarget) setAddItemModalOpen(false);
          }}
        >
          <div className="bg-white rounded-3xl w-full max-w-lg p-6 shadow-2xl border border-slate-200 relative max-h-[90vh] overflow-y-auto">
            <div className="flex items-center justify-between pb-3 border-b border-slate-100">
              <div className="flex items-center gap-2">
                <Plus className="h-5 w-5 text-[#ff5500]" />
                <h3 className="text-base font-black text-slate-900">Add New Dish / Menu Item</h3>
              </div>
              <button
                type="button"
                onClick={() => setAddItemModalOpen(false)}
                className="text-slate-400 hover:text-slate-900 p-1 cursor-pointer"
              >
                <X className="h-5 w-5" />
              </button>
            </div>

            <form
              onSubmit={(e) => {
                e.preventDefault();
                const priceNum = parseFloat(newDishForm.price);
                if (!newDishForm.name.trim() || isNaN(priceNum) || priceNum <= 0) {
                  alert('Please enter a valid dish name and selling price.');
                  return;
                }

                const resolvedCategory = newDishForm.category === 'Other' && newDishForm.customCategory?.trim()
                  ? newDishForm.customCategory.trim()
                  : newDishForm.category;

                const newItem = {
                  id: `dish_${Date.now()}_${Math.floor(Math.random() * 1000)}`,
                  name: newDishForm.name.trim(),
                  department: newDishForm.department || 'Kitchen',
                  category: resolvedCategory || 'Main Menu',
                  price: priceNum,
                  prepTime: newDishForm.prepTime || '15m',
                  imageUrl: newDishForm.imageUrl || '',
                  description: newDishForm.description || '',
                  recipe: newDishForm.recipeIngredients || []
                };

                setMenuItems(prev => [...prev, newItem]);
                recordAuditLog('MENU_ITEM_CREATED', newItem.id, `Created new dish "${newItem.name}" (${newItem.department}) for ${settings.currency} ${newItem.price.toFixed(2)}`);

                setNewDishForm({
                  name: '',
                  department: 'Kitchen',
                  category: 'Rice & Noodles',
                  customCategory: '',
                  price: '',
                  prepTime: '10m',
                  description: '',
                  imageUrl: '',
                  recipeIngredients: []
                });
                setAddItemModalOpen(false);
              }}
              className="mt-4 space-y-4"
            >
              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">Dish Name *</label>
                  <input
                    type="text"
                    required
                    autoFocus
                    value={newDishForm.name}
                    onChange={(e) => setNewDishForm(prev => ({ ...prev, name: e.target.value }))}
                    placeholder="e.g. Seafood Fried Rice"
                    className="w-full px-3.5 py-2.5 bg-slate-50 border border-slate-200 rounded-xl text-xs font-bold text-slate-900 focus:bg-white focus:outline-none focus:border-[#ff5500]"
                  />
                </div>

                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">Selling Price ({settings.currency}) *</label>
                  <input
                    type="number"
                    step="0.01"
                    min="0.01"
                    required
                    value={newDishForm.price}
                    onChange={(e) => setNewDishForm(prev => ({ ...prev, price: e.target.value }))}
                    placeholder="e.g. 1850.00"
                    className="w-full px-3.5 py-2.5 bg-slate-50 border border-slate-200 rounded-xl text-xs font-mono font-bold text-slate-900 focus:bg-white focus:outline-none focus:border-[#ff5500]"
                  />
                </div>
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">Department</label>
                  <select
                    value={newDishForm.department}
                    onChange={(e) => setNewDishForm(prev => ({ ...prev, department: e.target.value }))}
                    className="w-full px-3.5 py-2.5 bg-white border border-slate-200 rounded-xl text-xs font-bold text-slate-900 focus:outline-none focus:border-[#ff5500]"
                    style={{ color: '#0f172a', backgroundColor: '#ffffff' }}
                  >
                    <option value="Kitchen">Kitchen (Sends KOT)</option>
                    <option value="Bar">Bar (Sends BOT)</option>
                  </select>
                </div>

                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">Category</label>
                  <input
                    type="text"
                    required
                    value={newDishForm.category}
                    onChange={(e) => setNewDishForm(prev => ({ ...prev, category: e.target.value }))}
                    placeholder="e.g. Starters, Mains, Cocktails"
                    className="w-full px-3.5 py-2.5 bg-slate-50 border border-slate-200 rounded-xl text-xs font-bold text-slate-900 focus:bg-white focus:outline-none focus:border-[#ff5500]"
                  />
                </div>
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">Prep Time</label>
                  <input
                    type="text"
                    value={newDishForm.prepTime}
                    onChange={(e) => setNewDishForm(prev => ({ ...prev, prepTime: e.target.value }))}
                    placeholder="e.g. 15m"
                    className="w-full px-3.5 py-2.5 bg-slate-50 border border-slate-200 rounded-xl text-xs font-bold text-slate-900 focus:bg-white focus:outline-none focus:border-[#ff5500]"
                  />
                </div>

                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">Dish Photo</label>
                  <div className="flex gap-2">
                    <label className="cursor-pointer px-3 py-2 bg-slate-100 hover:bg-slate-200 border border-slate-200 text-slate-700 font-bold rounded-xl text-xs flex items-center gap-1.5 transition-colors">
                      <Upload className="h-3.5 w-3.5" />
                      <span>Upload</span>
                      <input
                        type="file"
                        accept="image/*"
                        className="hidden"
                        onChange={(e) => {
                          const file = e.target.files?.[0];
                          if (file) {
                            const reader = new FileReader();
                            reader.onload = ev => {
                              if (ev.target?.result) {
                                setNewDishForm(prev => ({ ...prev, imageUrl: ev.target.result }));
                              }
                            };
                            reader.readAsDataURL(file);
                          }
                        }}
                      />
                    </label>
                    <input
                      type="url"
                      value={newDishForm.imageUrl}
                      onChange={(e) => setNewDishForm(prev => ({ ...prev, imageUrl: e.target.value }))}
                      placeholder="or paste URL"
                      className="flex-1 px-2.5 py-1.5 bg-slate-50 border border-slate-200 rounded-xl text-xs text-slate-900 focus:bg-white focus:outline-none"
                    />
                  </div>
                </div>
              </div>

              <div>
                <label className="block text-xs font-bold text-slate-700 mb-1">Description</label>
                <textarea
                  rows={2}
                  value={newDishForm.description}
                  onChange={(e) => setNewDishForm(prev => ({ ...prev, description: e.target.value }))}
                  placeholder="Ingredients, dietary notes or flavors..."
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs text-slate-900 focus:bg-white focus:outline-none focus:border-[#ff5500]"
                />
              </div>

              {/* RECIPE BOM BUILDER */}
              <div className="p-3 bg-slate-50 border border-slate-200 rounded-2xl space-y-2">
                <span className="text-xs font-bold text-slate-800 flex items-center gap-1.5">
                  <BookOpen className="h-3.5 w-3.5 text-[#ff5500]" /> Link Recipe Ingredients (BOM)
                </span>
                <div className="flex gap-2">
                  <select id="newDishIngSelect" className="flex-1 px-2.5 py-1.5 border border-slate-200 rounded-xl text-xs bg-white text-slate-900">
                    {inventory.map(ing => (
                      <option key={ing.id} value={ing.id}>{ing.name} ({ing.unit})</option>
                    ))}
                  </select>
                  <input
                    id="newDishIngAmount"
                    type="number"
                    min="0.1"
                    step="any"
                    placeholder="Qty/portion"
                    className="w-28 px-2.5 py-1.5 border border-slate-200 rounded-xl text-xs font-mono bg-white text-slate-900"
                  />
                  <button
                    type="button"
                    onClick={() => {
                      const sel = document.getElementById('newDishIngSelect');
                      const amtInput = document.getElementById('newDishIngAmount');
                      const ingId = sel?.value;
                      const amt = parseFloat(amtInput?.value);
                      if (!ingId || isNaN(amt) || amt <= 0) return;

                      setNewDishForm(prev => ({
                        ...prev,
                        recipeIngredients: [...(prev.recipeIngredients || []), { ingredientId: ingId, amount: amt }]
                      }));
                      if (amtInput) amtInput.value = '';
                    }}
                    className="px-3 py-1.5 bg-slate-900 hover:bg-slate-800 text-white rounded-xl text-xs font-bold transition-colors cursor-pointer"
                  >
                    + Add
                  </button>
                </div>

                {newDishForm.recipeIngredients && newDishForm.recipeIngredients.length > 0 && (
                  <div className="space-y-1 max-h-28 overflow-y-auto">
                    {newDishForm.recipeIngredients.map((r, i) => {
                      const matchedIng = inventory.find(inv => inv.id === r.ingredientId);
                      return (
                        <div key={i} className="flex justify-between items-center text-xs p-1.5 bg-white rounded-lg border border-slate-200">
                          <span className="font-bold text-slate-800">{matchedIng?.name || r.ingredientId}</span>
                          <div className="flex items-center gap-2">
                            <span className="font-mono text-slate-500">{r.amount} {matchedIng?.unit}</span>
                            <button
                              type="button"
                              onClick={() => setNewDishForm(prev => ({
                                ...prev,
                                recipeIngredients: prev.recipeIngredients.filter((_, idx) => idx !== i)
                              }))}
                              className="text-slate-400 hover:text-rose-600 transition-colors cursor-pointer"
                            >
                              <X className="h-3.5 w-3.5" />
                            </button>
                          </div>
                        </div>
                      );
                    })}
                  </div>
                )}
              </div>

              <div className="flex gap-2 pt-2 border-t border-slate-100">
                <button
                  type="button"
                  onClick={() => setAddItemModalOpen(false)}
                  className="flex-1 py-2.5 bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold rounded-xl text-xs transition-colors cursor-pointer"
                >
                  Cancel
                </button>
                <button
                  type="submit"
                  className="flex-1 py-2.5 bg-[#ff5500] hover:bg-orange-600 text-white font-bold rounded-xl text-xs shadow-xs transition-all cursor-pointer active:scale-95"
                >
                  Create Menu Item
                </button>
              </div>
            </form>
          </div>
        </div>
      )}
      
      {/* MODAL: CHECKOUT & SETTLEMENT */}
      {checkoutModalOpen && (
        <div className="fixed inset-0 bg-black/75 backdrop-blur-xs z-50 flex items-center justify-center p-4">
          <div className="bg-white rounded-3xl w-full max-w-md p-6 shadow-2xl border border-slate-200">
            <div className="flex items-center justify-between pb-3 border-b border-slate-100">
              <h3 className="text-base font-black text-slate-900">
                Settle Bill: {settlingOrder ? settlingOrder.tableName : selectedTable.name}
              </h3>
              <button onClick={() => setCheckoutModalOpen(false)} className="text-slate-400 hover:text-slate-900 cursor-pointer">
                <X className="h-5 w-5" />
              </button>
            </div>

            {(() => {
              const currentRef = settlingOrder || {
                items: cart,
                serviceChargeActive,
                taxActive,
                discountPercent
              };
              const fin = calculateOrderFinancials(currentRef);

              return (
                <div className="mt-4 space-y-4">
                  <div className="grid grid-cols-3 gap-2">
                    {['CASH', 'CARD', 'SPLIT'].map(type => (
                      <button
                        key={type}
                        type="button"
                        onClick={() => setPaymentMethod(type)}
                        className={`py-2 rounded-xl text-xs font-bold border transition-all cursor-pointer ${
                          paymentMethod === type
                            ? 'bg-[#ff5500] text-white border-[#ff5500]'
                            : 'bg-slate-50 border-slate-200 text-slate-600'
                        }`}
                      >
                        {type}
                      </button>
                    ))}
                  </div>

                  {paymentMethod === 'CASH' && (
                    <div className="space-y-2">
                      <div className="flex items-center justify-between">
                        <label className="block text-xs font-bold text-slate-700">Cash Tendered ({settings.currency})</label>
                        <button
                          type="button"
                          onClick={() => setCashTendered(fin.total.toFixed(2))}
                          className="text-[10px] font-bold text-[#ff5500] hover:underline cursor-pointer"
                        >
                          Exact Amount
                        </button>
                      </div>
                      <input
                        type="number"
                        step="0.01"
                        value={cashTendered}
                        onChange={e => setCashTendered(e.target.value)}
                        placeholder={fin.total.toFixed(2)}
                        className="w-full px-3 py-2 border border-slate-200 rounded-xl text-slate-900 font-mono text-sm focus:outline-none focus:border-[#ff5500]"
                      />

                      <div className="flex gap-1.5 overflow-x-auto pt-0.5">
                        {[
                          Math.ceil(fin.total / 100) * 100,
                          Math.ceil(fin.total / 500) * 500,
                          Math.ceil(fin.total / 1000) * 1000,
                          5000
                        ]
                          .filter((val, idx, arr) => val >= fin.total && arr.indexOf(val) === idx)
                          .slice(0, 3)
                          .map(val => (
                            <button
                              key={val}
                              type="button"
                              onClick={() => setCashTendered(val.toString())}
                              className="px-2.5 py-1 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-lg text-[10px] font-mono font-bold transition-colors cursor-pointer"
                            >
                              {settings.currency} {val}
                            </button>
                          ))}
                      </div>

                      {parseFloat(cashTendered) >= fin.total && (
                        <div className="p-2.5 bg-emerald-50 border border-emerald-200 rounded-xl text-xs font-mono font-bold text-emerald-800 flex justify-between items-center">
                          <span>Balance / Change to Return:</span>
                          <span className="text-sm text-emerald-700">
                            {settings.currency} {(parseFloat(cashTendered) - fin.total).toFixed(2)}
                          </span>
                        </div>
                      )}
                    </div>
                  )}

                  <div className="p-3 bg-slate-50 rounded-2xl border border-slate-200 space-y-1 text-xs">
                    <div className="flex justify-between text-slate-500">
                      <span>Subtotal</span>
                      <span className="font-mono">{settings.currency} {fin.subtotal.toFixed(2)}</span>
                    </div>
                    {fin.discount > 0 && (
                      <div className="flex justify-between text-rose-600">
                        <span>Discount</span>
                        <span className="font-mono">-{settings.currency} {fin.discount.toFixed(2)}</span>
                      </div>
                    )}
                    {fin.service > 0 && (
                      <div className="flex justify-between text-emerald-700 font-medium">
                        <span>Service Charge</span>
                        <span className="font-mono">+{settings.currency} {fin.service.toFixed(2)}</span>
                      </div>
                    )}
                    {fin.tax > 0 && (
                      <div className="flex justify-between text-indigo-700 font-medium">
                        <span>Taxes</span>
                        <span className="font-mono">+{settings.currency} {fin.tax.toFixed(2)}</span>
                      </div>
                    )}
                    <div className="flex justify-between text-sm font-black text-slate-900 pt-2 border-t border-slate-200">
                      <span>Total Due</span>
                      <span className="text-base font-mono text-[#ff5500]">{settings.currency} {fin.total.toFixed(2)}</span>
                    </div>
                  </div>

                  <button
                    type="button"
                    onClick={handleCompleteSettlement}
                    className="w-full py-3 bg-[#008f5d] hover:bg-emerald-700 text-white font-extrabold rounded-xl text-xs uppercase tracking-wider shadow-xs cursor-pointer active:scale-95"
                  >
                    Confirm Settlement &amp; Deduct BOM Stock
                  </button>
                </div>
              );
            })()}
          </div>
        </div>
      )}

      {/* STREAMLINED NON-BLOCKING PRINT NOTIFICATION TOAST */}
      {printNotice && (
        <div className="fixed bottom-6 right-6 z-50 bg-slate-900/95 text-white px-4 py-3 rounded-2xl shadow-2xl border border-slate-700 flex items-center gap-3 backdrop-blur-md transition-all">
          <div className="h-9 w-9 rounded-xl bg-orange-500/20 text-[#ff5500] flex items-center justify-center shrink-0">
            <Printer className="h-5 w-5 animate-pulse" />
          </div>
          <div>
            <p className="text-xs font-bold text-white">{printNotice.title}</p>
            <p className="text-[10px] text-slate-400 font-mono mt-0.5">{printNotice.detail}</p>
          </div>
        </div>
      )}

      {/* HIDDEN OFF-SCREEN THERMAL PRINT AREA */}
      <div
        id="thermal-print-area"
        style={{
          fontSize: settings.receiptFontSize || '11px',
          fontFamily: settings.receiptFontFamily || 'monospace',
          padding: settings.receiptMargin || '2mm'
        }}
        className="hidden print:block w-full bg-white text-slate-900 leading-tight space-y-4"
      >
        {activePrintSlip && (
          <>
            {/* 1. KOT / BOT Order & Add-On Dispatch Slip */}
            {activePrintSlip.type === 'KOT_BOT_DISPATCH' && (
              <div className="space-y-4">
                {activePrintSlip.data.kitchenItems?.length > 0 && (
                  <div className="border-b-2 border-dashed border-slate-800 pb-3 text-center">
                    <p className="font-black text-xs">
                      {activePrintSlip.data.order?.isAddon ? '** KITCHEN ADD-ON TICKET (KOT) **' : '** KITCHEN ORDER TICKET (KOT) **'}
                    </p>
                    <p className="font-bold text-xs mt-1">{activePrintSlip.data.order.tableName}</p>
                    <p className="text-[10px]">Time: {activePrintSlip.data.order.sentAt}</p>
                    <div className="text-left py-2 space-y-1">
                      {activePrintSlip.data.kitchenItems.map((item, idx) => (
                        <div key={idx}>
                          <p className="font-bold">{item.qty}x {item.name}</p>
                          {item.notes && <p className="text-[10px] pl-2 italic">&gt; {item.notes}</p>}
                        </div>
                      ))}
                    </div>
                  </div>
                )}

                {activePrintSlip.data.barItems?.length > 0 && (
                  <div className="border-b-2 border-dashed border-slate-800 pb-3 text-center">
                    <p className="font-black text-xs">
                      {activePrintSlip.data.order?.isAddon ? '** BAR ADD-ON TICKET (BOT) **' : '** BAR ORDER TICKET (BOT) **'}
                    </p>
                    <p className="font-bold text-xs mt-1">{activePrintSlip.data.order.tableName}</p>
                    <p className="text-[10px]">Time: {activePrintSlip.data.order.sentAt}</p>
                    <div className="text-left py-2 space-y-1">
                      {activePrintSlip.data.barItems.map((item, idx) => (
                        <div key={idx}>
                          <p className="font-bold">{item.qty}x {item.name}</p>
                          {item.notes && <p className="text-[10px] pl-2 italic">&gt; {item.notes}</p>}
                        </div>
                      ))}
                    </div>
                  </div>
                )}
              </div>
            )}

            {/* 1. Sri Lankan Statutory Payslip Print Slip */}
            {activePrintSlip.type === 'PAYSLIP_PRINT' && (
              <div className="space-y-3 font-mono text-xs">
                {/* Header */}
                <div className="text-center border-b-2 border-dashed border-black pb-2">
                  <h2 className="font-black text-sm uppercase">{settings.restaurantName}</h2>
                  <p className="text-[10px]">{settings.address}</p>
                  <p className="font-bold text-xs uppercase mt-1">*** SALARY PAY SLIP ***</p>
                  <p className="text-[10px]">Period: {activePrintSlip.data.period} • Slip #{activePrintSlip.data.id}</p>
                </div>

                {/* Employee Info */}
                <div className="space-y-1 border-b border-black pb-2 text-[11px]">
                  <div className="flex justify-between">
                    <span>Employee:</span>
                    <span className="font-bold">{activePrintSlip.data.staffName}</span>
                  </div>
                  <div className="flex justify-between">
                    <span>Designation:</span>
                    <span>{activePrintSlip.data.role}</span>
                  </div>
                  <div className="flex justify-between text-[10px]">
                    <span>Statutory Status:</span>
                    <span className="font-bold uppercase">
                      {activePrintSlip.data.breakdown?.epfEtfEnabled !== false ? 'EPF / ETF Enrolled' : 'Statutory Exempt (No EPF)'}
                    </span>
                  </div>
                </div>
                {/* Attendance & Shift Breakdown */}
                {activePrintSlip.data.breakdown?.standardWorkingDays && (
                  <div className="space-y-1 border-b border-dashed border-black pb-2 text-[10px]">
                    <p className="font-bold uppercase">=== ATTENDANCE &amp; LEAVES ===</p>
                    <div className="flex justify-between">
                      <span>Standard / Worked Days:</span>
                      <span className="font-bold">
                        {activePrintSlip.data.breakdown.workedDays} / {activePrintSlip.data.breakdown.standardWorkingDays} Days
                      </span>
                    </div>
                    <div className="flex justify-between">
                      <span>Paid Leaves / Holidays:</span>
                      <span>
                        {activePrintSlip.data.breakdown.paidLeaves} Paid / {activePrintSlip.data.breakdown.holidaysCount} Holidays
                      </span>
                    </div>
                    {Number(activePrintSlip.data.breakdown.shortShiftsCount) > 0 && (
                      <div className="flex justify-between">
                        <span>Short Shifts (&lt;5h):</span>
                        <span>{activePrintSlip.data.breakdown.shortShiftsCount} Recorded</span>
                      </div>
                    )}
                    {Number(activePrintSlip.data.breakdown.unpaidLeaves) > 0 && (
                      <div className="flex justify-between font-bold text-rose-800">
                        <span>Unpaid Leave Days:</span>
                        <span>{activePrintSlip.data.breakdown.unpaidLeaves} Days</span>
                      </div>
                    )}
                  </div>
                )}

                {/* Earnings Section */}
                <div className="space-y-1 border-b border-dashed border-black pb-2 text-[11px]">
                  <p className="font-bold text-[10px] uppercase">=== EARNINGS ===</p>
                  <div className="flex justify-between">
                    <span>Basic Salary:</span>
                    <span>{settings.currency} {activePrintSlip.data.breakdown?.basic.toFixed(2)}</span>
                  </div>
                  <div className="flex justify-between">
                    <span>Budgetary Allowance (BRA):</span>
                    <span>+{settings.currency} {activePrintSlip.data.breakdown?.bra.toFixed(2)}</span>
                  </div>
                  {activePrintSlip.data.breakdown?.allowances > 0 && (
                    <div className="flex justify-between">
                      <span>Other Allowances:</span>
                      <span>+{settings.currency} {activePrintSlip.data.breakdown.allowances.toFixed(2)}</span>
                    </div>
                  )}
                  {activePrintSlip.data.breakdown?.serviceChargeBonus > 0 && (
                    <div className="flex justify-between font-bold">
                      <span>Service Pool Share:</span>
                      <span>+{settings.currency} {activePrintSlip.data.breakdown.serviceChargeBonus.toFixed(2)}</span>
                    </div>
                  )}
                  {activePrintSlip.data.breakdown?.incentiveBonus > 0 && (
                    <div className="flex justify-between font-bold">
                      <span>Performance Bonus:</span>
                      <span>+{settings.currency} {activePrintSlip.data.breakdown.incentiveBonus.toFixed(2)}</span>
                    </div>
                  )}
                  {activePrintSlip.data.breakdown?.overtimePay > 0 && (
                    <div className="flex justify-between">
                      <span>Overtime Pay:</span>
                      <span>+{settings.currency} {activePrintSlip.data.breakdown.overtimePay.toFixed(2)}</span>
                    </div>
                  )}
                  <div className="flex justify-between font-black border-t border-black pt-1">
                    <span>GROSS EARNINGS:</span>
                    <span>{settings.currency} {activePrintSlip.data.breakdown?.grossEarnings.toFixed(2)}</span>
                  </div>
                </div>

                {/* Deductions Section */}
                <div className="space-y-1 border-b border-dashed border-black pb-2 text-[11px]">
                  <p className="font-bold text-[10px] uppercase">=== DEDUCTIONS ===</p>
                  {activePrintSlip.data.breakdown?.epfEtfEnabled !== false ? (
                    <>
                      <div className="flex justify-between text-[10px] text-slate-600">
                        <span>EPF Base Earnings:</span>
                        <span>{settings.currency} {activePrintSlip.data.breakdown?.epfLiableEarnings.toFixed(2)}</span>
                      </div>
                      <div className="flex justify-between font-bold text-rose-700">
                        <span>EPF (Employee 8%):</span>
                        <span>-{settings.currency} {activePrintSlip.data.breakdown?.epfEmployee.toFixed(2)}</span>
                      </div>
                      {Number(activePrintSlip.data.breakdown?.unpaidLeaveDeduction) > 0 && (
                    <div className="flex justify-between font-bold text-rose-700">
                      <span>Unpaid Leave ({activePrintSlip.data.breakdown.unpaidLeaves}d @ {settings.currency}{activePrintSlip.data.breakdown.perDayBasicRate}):</span>
                      <span>-{settings.currency} {Number(activePrintSlip.data.breakdown.unpaidLeaveDeduction).toFixed(2)}</span>
                    </div>
                  )}
                    </>
                  ) : (
                    <div className="flex justify-between italic text-[10px] text-slate-600">
                      <span>EPF (Employee 8%):</span>
                      <span>Exempt (0.00)</span>
                    </div>
                  )}

                  {/* ITEMIZE SALARY ADVANCE DEDUCTION */}
                  {Number(activePrintSlip.data.breakdown?.salaryAdvance) > 0 && (
                    <div className="flex justify-between font-bold text-amber-800">
                      <span>Salary Advance Deducted:</span>
                      <span>-{settings.currency} {Number(activePrintSlip.data.breakdown.salaryAdvance).toFixed(2)}</span>
                    </div>
                  )}

                  {activePrintSlip.data.breakdown?.otherDeductions > 0 && (
                    <div className="flex justify-between text-rose-700">
                      <span>Other Deductions:</span>
                      <span>-{settings.currency} {activePrintSlip.data.breakdown.otherDeductions.toFixed(2)}</span>
                    </div>
                  )}
                </div>

                {/* Take-Home Pay */}
                <div className="space-y-1 border-b-2 border-black pb-2 text-xs">
                  <div className="flex justify-between font-black text-sm">
                    <span>NET TAKE-HOME PAY:</span>
                    <span>{settings.currency} {activePrintSlip.data.breakdown?.netSalary.toFixed(2)}</span>
                  </div>
                </div>

                {/* Employer Statutory Remittances */}
                {activePrintSlip.data.breakdown?.epfEtfEnabled !== false && (
                  <div className="space-y-1 border-b border-dotted border-black pb-2 text-[10px] text-slate-700">
                    <p className="font-bold uppercase">=== EMPLOYER STATUTORY CONTRIBUTIONS ===</p>
                    <div className="flex justify-between">
                      <span>EPF (Employer 12%):</span>
                      <span>{settings.currency} {activePrintSlip.data.breakdown?.epfEmployer.toFixed(2)}</span>
                    </div>
                    <div className="flex justify-between">
                      <span>ETF (Employer 3%):</span>
                      <span>{settings.currency} {activePrintSlip.data.breakdown?.etfEmployer.toFixed(2)}</span>
                    </div>
                  </div>
                )}
                
                {/* Signatures */}
                <div className="pt-3 text-[9px] flex justify-between">
                  <div className="text-center">
                    <p>___________________</p>
                    <p>Employee Signature</p>
                  </div>
                  <div className="text-center">
                    <p>___________________</p>
                    <p>Authorized Officer</p>
                  </div>
                </div>
              </div>
            )}

            {/* 2. STANDALONE SALARY ADVANCE VOUCHER SLIP */}
            {activePrintSlip.type === 'SALARY_ADVANCE_VOUCHER' && (
              <div className="space-y-3 font-mono text-xs">
                <div className="text-center border-b-2 border-dashed border-black pb-2">
                  <h2 className="font-black text-sm uppercase">{settings.restaurantName}</h2>
                  <p className="text-[10px]">{settings.address}</p>
                  <p className="font-black text-xs uppercase mt-1 tracking-wider">*** SALARY ADVANCE DISBURSEMENT VOUCHER ***</p>
                  <p className="text-[10px]">Voucher Ref: {activePrintSlip.data.id || 'ADV-NEW'} • Month: {activePrintSlip.data.period}</p>
                  <p className="text-[9px]">Date: {new Date().toLocaleDateString()} {new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}</p>
                </div>

                <div className="space-y-1.5 border-b border-black pb-2 text-[11px]">
                  <div className="flex justify-between">
                    <span>Employee Name:</span>
                    <span className="font-bold">{activePrintSlip.data.staffName}</span>
                  </div>
                  <div className="flex justify-between">
                    <span>Designation / Role:</span>
                    <span>{activePrintSlip.data.role}</span>
                  </div>
                  <div className="flex justify-between">
                    <span>Salary Month:</span>
                    <span className="font-bold">{activePrintSlip.data.period}</span>
                  </div>
                </div>

                <div className="p-2 border border-black rounded space-y-1 text-center bg-slate-50">
                  <span className="text-[10px] uppercase font-bold block">Cash Advance Amount</span>
                  <p className="text-base font-black font-mono">
                    {settings.currency} {(Number(activePrintSlip.data.amount) || 0).toFixed(2)}
                  </p>
                </div>

                <div className="text-[10px] text-slate-700 leading-snug space-y-1 border-b border-dashed border-black pb-2">
                  <p>
                    I, <strong>{activePrintSlip.data.staffName}</strong>, hereby acknowledge the receipt of {settings.currency} {(Number(activePrintSlip.data.amount) || 0).toFixed(2)} in cash as an advance towards my wages for {activePrintSlip.data.period}, and consent to this amount being deducted on the final monthly payroll.
                  </p>
                  {activePrintSlip.data.notes && (
                    <p className="italic text-[9px] pt-1">Note: {activePrintSlip.data.notes}</p>
                  )}
                </div>

                <div className="pt-4 text-[9px] flex justify-between">
                  <div className="text-center">
                    <p>___________________</p>
                    <p className="mt-0.5">Employee Signature</p>
                  </div>
                  <div className="text-center">
                    <p>___________________</p>
                    <p className="mt-0.5">Authorized By (Manager)</p>
                  </div>
                </div>
              </div>
            )}

            {/* 2. Proforma Temporary Bill Slip */}
            {activePrintSlip.type === 'TEMP_BILL' && (
              <div className="space-y-2">
                <div className="text-center border-b-2 border-dashed border-slate-800 pb-2">
                  <p className="font-black text-sm">{settings.restaurantName}</p>
                  <p className="font-bold text-xs mt-1">*** PROFORMA TEMPORARY BILL ***</p>
                  <p className="text-[10px]">Table: {activePrintSlip.data.table} • Server: {activePrintSlip.data.server}</p>
                  <p className="text-[9px]">{new Date().toLocaleString()}</p>
                </div>
                <div className="py-1 border-b border-slate-300 space-y-1">
                  {activePrintSlip.data.items.map((item, idx) => (
                    <div key={idx} className="flex justify-between">
                      <span>{item.qty}x {item.name}</span>
                      <span>{settings.currency} {(item.price * item.qty).toFixed(2)}</span>
                    </div>
                  ))}
                </div>
                <div className="space-y-1 text-[10px]">
                  <div className="flex justify-between">
                    <span>Subtotal:</span>
                    <span>{settings.currency} {activePrintSlip.data.subtotal.toFixed(2)}</span>
                  </div>
                  {activePrintSlip.data.discount > 0 && (
                    <div className="flex justify-between text-rose-600">
                      <span>Discount:</span>
                      <span>-{settings.currency} {activePrintSlip.data.discount.toFixed(2)}</span>
                    </div>
                  )}
                  {activePrintSlip.data.service > 0 && (
                    <div className="flex justify-between">
                      <span>Service Charge:</span>
                      <span>+{settings.currency} {activePrintSlip.data.service.toFixed(2)}</span>
                    </div>
                  )}
                  {activePrintSlip.data.tax > 0 && (
                    <div className="flex justify-between">
                      <span>Taxes:</span>
                      <span>+{settings.currency} {activePrintSlip.data.tax.toFixed(2)}</span>
                    </div>
                  )}
                  <div className="flex justify-between font-black text-xs pt-1 border-t border-slate-800">
                    <span>ESTIMATED TOTAL:</span>
                    <span>{settings.currency} {activePrintSlip.data.total.toFixed(2)}</span>
                  </div>
                </div>
                <p className="text-center text-[9px] italic pt-2">Not a tax invoice • For guest review only</p>
              </div>
            )}

            {/* 3. Final Settlement Tax Invoice */}
            {activePrintSlip.type === 'FINAL_BILL' && (
              <div className="space-y-3 font-mono">
                <div className="text-center border-b-2 border-dashed border-black pb-3 space-y-1">
                  <h1 className="font-black text-xl tracking-tight uppercase leading-tight">
                    {settings.restaurantName}
                  </h1>
                  {settings.tagline && (
                    <p className="font-bold text-xs uppercase tracking-wider">
                      {settings.tagline}
                    </p>
                  )}
                  <div className="text-xs font-semibold leading-snug whitespace-pre-line text-black pt-1">
                    {settings.receiptHeader}
                  </div>
                  <div className="pt-2">
                    <p className="font-black text-sm tracking-wide uppercase border-y border-black py-1 inline-block w-full">
                      TAX INVOICE #{activePrintSlip.data.invoiceNo}
                    </p>
                  </div>
                  <p className="text-xs font-bold tracking-tight pt-1">
                    {activePrintSlip.data.date} &bull; {activePrintSlip.data.table}
                  </p>
                </div>

                <div className="py-1 border-b border-black space-y-1 text-xs">
                  {activePrintSlip.data.items.map((item, idx) => (
                    <div key={idx} className="flex justify-between font-semibold">
                      <span>{item.qty}x {item.name}</span>
                      <span>{settings.currency} {(item.price * item.qty).toFixed(2)}</span>
                    </div>
                  ))}
                </div>

                <div className="space-y-1 text-xs">
                  <div className="flex justify-between">
                    <span>Subtotal:</span>
                    <span>{settings.currency} {(activePrintSlip.data.subtotal || 0).toFixed(2)}</span>
                  </div>
                  {activePrintSlip.data.discount > 0 && (
                    <div className="flex justify-between text-rose-600 font-semibold">
                      <span>Discount ({activePrintSlip.data.discountPercent || 0}%):</span>
                      <span>-{settings.currency} {activePrintSlip.data.discount.toFixed(2)}</span>
                    </div>
                  )}
                  {activePrintSlip.data.serviceCharge > 0 && (
                    <div className="flex justify-between font-medium">
                      <span>Service Charge ({settings.serviceChargeRate}%):</span>
                      <span>+{settings.currency} {activePrintSlip.data.serviceCharge.toFixed(2)}</span>
                    </div>
                  )}
                  {activePrintSlip.data.tax > 0 && (
                    <div className="flex justify-between font-medium">
                      <span>Taxes ({settings.taxRate}%):</span>
                      <span>+{settings.currency} {activePrintSlip.data.tax.toFixed(2)}</span>
                    </div>
                  )}
                  <div className="flex justify-between font-black text-sm pt-1.5 border-t border-black">
                    <span>TOTAL AMOUNT DUE:</span>
                    <span>{settings.currency} {activePrintSlip.data.total.toFixed(2)}</span>
                  </div>
                  <div className="flex justify-between text-xs font-bold pt-0.5">
                    <span>PAYMENT METHOD:</span>
                    <span>{activePrintSlip.data.paymentMethod}</span>
                  </div>

                  {activePrintSlip.data.paymentMethod === 'CASH' && (
                    <div className="pt-2 mt-1 border-t border-dashed border-black space-y-1">
                      <div className="flex justify-between text-xs font-bold">
                        <span>CASH TENDERED (GIVEN):</span>
                        <span>
                          {settings.currency} {(activePrintSlip.data.cashTendered !== undefined ? activePrintSlip.data.cashTendered : activePrintSlip.data.total).toFixed(2)}
                        </span>
                      </div>
                      <div className="flex justify-between text-sm font-black">
                        <span>BALANCE / CHANGE DUE:</span>
                        <span>
                          {settings.currency} {(activePrintSlip.data.changeDue || 0).toFixed(2)}
                        </span>
                      </div>
                    </div>
                  )}
                </div>

                <p className="text-center font-bold text-xs pt-3 whitespace-pre-line border-t border-dashed border-black">
                  {settings.receiptFooter}
                </p>
              </div>
            )}
            
            {/* 4. Cash Out Voucher Slip */}
            {activePrintSlip.type === 'CASH_OUT_VOUCHER' && (
              <div className="space-y-2">
                <div className="text-center border-b-2 border-dashed border-slate-800 pb-2">
                  <p className="font-black text-sm">{settings.restaurantName}</p>
                  <p className="font-bold text-xs mt-1">*** CASH OUT VOUCHER ***</p>
                  <p className="text-[10px]">Ref: {activePrintSlip.data.id} • Terminal: {settings.terminalId}</p>
                  <p className="text-[9px]">{activePrintSlip.data.date || getLocalDateStr()} {activePrintSlip.data.createdAt || activePrintSlip.data.time}</p>
                </div>

                <div className="py-2 border-b border-slate-300 space-y-1 text-xs">
                  <div className="flex justify-between">
                    <span className="font-bold">CATEGORY:</span>
                    <span>{activePrintSlip.data.category}</span>
                  </div>
                  <div className="flex justify-between">
                    <span className="font-bold">PAID TO:</span>
                    <span>{activePrintSlip.data.recipient || 'General Expense'}</span>
                  </div>
                  <div className="text-left pt-1">
                    <span className="font-bold">DESCRIPTION:</span>
                    <p className="italic text-[10px]">{activePrintSlip.data.reason}</p>
                  </div>
                </div>

                <div className="py-1 border-b border-slate-800 space-y-1 text-xs">
                  <div className="flex justify-between font-black text-sm">
                    <span>AMOUNT DISBURSED:</span>
                    <span>{settings.currency} {(parseFloat(activePrintSlip.data.amount) || 0).toFixed(2)}</span>
                  </div>
                </div>

                <div className="pt-2 text-[10px] space-y-3">
                  <div className="flex justify-between">
                    <span>Requested By: {activePrintSlip.data.requestedBy}</span>
                    <span>Approved By: {activePrintSlip.data.approvedBy || 'Supervisor'}</span>
                  </div>

                  <div className="pt-4 border-t border-dotted border-slate-400 flex justify-between text-[9px]">
                    <div>
                      <p>___________________</p>
                      <p>Cashier Signature</p>
                    </div>
                    <div className="text-right">
                      <p>___________________</p>
                      <p>Manager Signature</p>
                    </div>
                  </div>
                </div>

                <p className="text-center text-[9px] italic pt-2">Official Cash Drawer Disbursement Voucher</p>
              </div>
            )}

            {/* 5. Z-Report Shift Balancing Slip */}
            {activePrintSlip.type === 'Z_REPORT' && (
              <div className="space-y-3 font-mono">
                <div className="text-center border-b-2 border-dashed border-slate-800 pb-2">
                  <p className="font-black text-sm uppercase">{settings.restaurantName}</p>
                  <p className="text-[9px] uppercase tracking-wider">{settings.tagline}</p>
                  <p className="font-black text-xs mt-1.5">*** END OF SHIFT Z-REPORT ***</p>
                  <p className="text-[10px] font-bold mt-0.5">SHIFT ID: {activePrintSlip.data.shiftId}</p>
                  <p className="text-[9px]">Terminal: {settings.terminalId}</p>
                </div>

                <div className="text-[10px] space-y-0.5 border-b border-dashed border-slate-400 pb-2">
                  <div className="flex justify-between">
                    <span>Opened:</span>
                    <span>{activePrintSlip.data.openedDate} {activePrintSlip.data.openedAt}</span>
                  </div>
                  <div className="flex justify-between">
                    <span>Opened By:</span>
                    <span>{activePrintSlip.data.openedBy}</span>
                  </div>
                  <div className="flex justify-between">
                    <span>Closed:</span>
                    <span>{activePrintSlip.data.closedDate || activePrintSlip.data.openedDate} {activePrintSlip.data.closedAt}</span>
                  </div>
                  <div className="flex justify-between font-bold">
                    <span>Closed By:</span>
                    <span>{activePrintSlip.data.closedBy}</span>
                  </div>
                </div>

                <div className="text-[10px] space-y-1 border-b border-dashed border-slate-400 pb-2">
                  <p className="font-black text-[11px] uppercase">=== SALES SUMMARY ===</p>
                  <div className="flex justify-between">
                    <span>Total Invoices Settled:</span>
                    <span className="font-bold">{activePrintSlip.data.metrics?.totalBills || 0}</span>
                  </div>
                  <div className="flex justify-between">
                    <span>Cash Sales:</span>
                    <span className="font-bold">+{settings.currency} {(activePrintSlip.data.metrics?.cashSales || 0).toFixed(2)}</span>
                  </div>
                  <div className="flex justify-between">
                    <span>Card / Digital Sales:</span>
                    <span>+{settings.currency} {(activePrintSlip.data.metrics?.cardSales || 0).toFixed(2)}</span>
                  </div>
                  {activePrintSlip.data.metrics?.otherSales > 0 && (
                    <div className="flex justify-between">
                      <span>Split / Other Sales:</span>
                      <span>+{settings.currency} {activePrintSlip.data.metrics.otherSales.toFixed(2)}</span>
                    </div>
                  )}
                  <div className="flex justify-between font-black text-xs pt-1 border-t border-slate-300">
                    <span>TOTAL GROSS SALES:</span>
                    <span>{settings.currency} {(activePrintSlip.data.metrics?.grossSales || 0).toFixed(2)}</span>
                  </div>
                </div>

                <div className="text-[10px] space-y-1 border-b border-dashed border-slate-400 pb-2">
                  <div className="flex justify-between font-black text-[11px] uppercase">
                    <span>=== CASH DISBURSEMENTS ===</span>
                    <span>-{settings.currency} {(activePrintSlip.data.metrics?.cashOutTotal || 0).toFixed(2)}</span>
                  </div>
                  {(!activePrintSlip.data.metrics?.approvedPayouts || activePrintSlip.data.metrics.approvedPayouts.length === 0) ? (
                    <p className="italic text-[9px] text-slate-500">No cash out payouts recorded.</p>
                  ) : (
                    activePrintSlip.data.metrics.approvedPayouts.map((payout, idx) => (
                      <div key={idx} className="flex justify-between text-[9px]">
                        <span className="truncate max-w-[180px]">{payout.reason} ({payout.category})</span>
                        <span className="font-bold">-{settings.currency} {(parseFloat(payout.amount) || 0).toFixed(2)}</span>
                      </div>
                    ))
                  )}
                </div>

                <div className="text-[10px] space-y-1 border-b border-dashed border-slate-400 pb-2">
                  <p className="font-black text-[11px] uppercase">=== CASH NOTE BREAKDOWN ===</p>
                  {activePrintSlip.data.metrics?.denominations && Object.keys(activePrintSlip.data.metrics.denominations).length > 0 ? (
                    Object.entries(activePrintSlip.data.metrics.denominations)
                      .sort((a, b) => Number(b[0]) - Number(a[0]))
                      .map(([denom, count]) => {
                        const noteCount = Number(count) || 0;
                        const subtotal = Number(denom) * noteCount;
                        return (
                          <div key={denom} className="flex justify-between text-[9px]">
                            <span>{settings.currency} {denom} x {noteCount}</span>
                            <span className="font-mono">{settings.currency} {subtotal.toFixed(2)}</span>
                          </div>
                        );
                      })
                  ) : (
                    <p className="italic text-[9px] text-slate-500">No denomination breakdown entered.</p>
                  )}
                </div>
                

                <div className="text-xs space-y-1 border-b-2 border-dashed border-slate-800 pb-2.5">
                  <p className="font-black text-[11px] uppercase">=== DRAWER BALANCING ===</p>
                  <div className="flex justify-between text-[10px]">
                    <span>Opening Cash Float:</span>
                    <span>{settings.currency} {(activePrintSlip.data.metrics?.startingFloat || 0).toFixed(2)}</span>
                  </div>
                  <div className="flex justify-between text-[10px]">
                    <span>+ Cash Sales Added:</span>
                    <span>+{settings.currency} {(activePrintSlip.data.metrics?.cashSales || 0).toFixed(2)}</span>
                  </div>
                  <div className="flex justify-between text-[10px]">
                    <span>- Cash Out Payouts:</span>
                    <span>-{settings.currency} {(activePrintSlip.data.metrics?.cashOutTotal || 0).toFixed(2)}</span>
                  </div>
                  <div className="flex justify-between font-bold text-[11px] pt-1 border-t border-slate-300">
                    <span>EXPECTED IN DRAWER:</span>
                    <span>{settings.currency} {(activePrintSlip.data.metrics?.expectedCash || 0).toFixed(2)}</span>
                  </div>
                  <div className="flex justify-between font-black text-xs">
                    <span>ACTUAL COUNTED CASH:</span>
                    <span>{settings.currency} {(activePrintSlip.data.metrics?.countedCash || 0).toFixed(2)}</span>
                  </div>
                  <div className="flex justify-between font-black text-xs pt-1 border-t border-slate-800">
                    <span>DRAWER VARIANCE:</span>
                    <span>
                      {(activePrintSlip.data.metrics?.variance || 0) > 0 ? '+' : ''}
                      {settings.currency} {(activePrintSlip.data.metrics?.variance || 0).toFixed(2)}{' '}
                      ({(activePrintSlip.data.metrics?.variance || 0) === 0
                        ? 'BALANCED'
                        : (activePrintSlip.data.metrics?.variance || 0) > 0
                        ? 'OVERAGE'
                        : 'SHORTAGE'})
                    </span>
                  </div>

                  {Boolean(activePrintSlip.data.metrics?.carriedDiscrepancy) && (
                    <div className="flex justify-between font-black text-[11px] pt-1 border-t border-dashed border-slate-600">
                      <span>CUMULATIVE UNSETTLED CARRYOVER:</span>
                      <span>
                        {(activePrintSlip.data.metrics?.carriedDiscrepancy || 0) > 0 ? '+' : ''}
                        {settings.currency} {(activePrintSlip.data.metrics?.carriedDiscrepancy || 0).toFixed(2)}
                      </span>
                    </div>
                  )}
                </div>

                <div className="pt-3 text-[9px] space-y-4">
                  <div className="flex justify-between">
                    <div>
                      <p>_______________________</p>
                      <p className="mt-0.5">Cashier: {activePrintSlip.data.closedBy}</p>
                    </div>
                    <div className="text-right">
                      <p>_______________________</p>
                      <p className="mt-0.5">Manager / Supervisor</p>
                    </div>
                  </div>
                  <p className="text-center italic text-[8px] pt-1">
                    System Generated Shift Audit Record • Retain for Accounts
                  </p>
                </div>
              </div>
            )}
          </>
        )}
      </div>

      {/* MODAL: CASH OUT MANAGER APPROVAL */}
      {cashOutApprovalModal.open && cashOutApprovalModal.item && (
        <div className="fixed inset-0 bg-black/75 backdrop-blur-xs z-50 flex items-center justify-center p-4">
          <div className="bg-white rounded-3xl w-full max-w-sm p-6 shadow-2xl border border-slate-200">
            <div className="flex items-center justify-between pb-3 border-b border-slate-100">
              <div className="flex items-center gap-2">
                <ShieldCheck className="h-5 w-5 text-indigo-600" />
                <h3 className="text-base font-black text-slate-900">Manager Authorization</h3>
              </div>
              <button
                type="button"
                onClick={() => setCashOutApprovalModal({ open: false, item: null, managerPin: '', error: '' })}
                className="text-slate-400 hover:text-slate-900 cursor-pointer"
              >
                <X className="h-5 w-5" />
              </button>
            </div>

            <div className="mt-4 space-y-4">
              <div className="p-3 bg-slate-50 rounded-2xl border border-slate-200 text-xs space-y-1">
                <div className="flex justify-between">
                  <span className="text-slate-500">Expense:</span>
                  <span className="font-bold text-slate-900">{cashOutApprovalModal.item.reason}</span>
                </div>
                <div className="flex justify-between">
                  <span className="text-slate-500">Amount:</span>
                  <span className="font-mono font-black text-[#ff5500]">
                    {settings.currency} {(parseFloat(cashOutApprovalModal.item.amount) || 0).toFixed(2)}
                  </span>
                </div>
                <div className="flex justify-between">
                  <span className="text-slate-500">Requested By:</span>
                  <span className="font-medium text-slate-700">{cashOutApprovalModal.item.requestedBy}</span>
                </div>
              </div>

              <div>
                <label className="block text-xs font-bold text-slate-700 mb-1 text-center">
                  Enter Manager / Admin 4-Digit PIN
                </label>
                <input
                  type="password"
                  maxLength={4}
                  autoFocus
                  value={cashOutApprovalModal.managerPin}
                  onChange={(e) => setCashOutApprovalModal(prev => ({ ...prev, managerPin: e.target.value, error: '' }))}
                  placeholder="••••"
                  className="w-full px-3 py-2 text-center text-xl font-mono tracking-widest border border-slate-200 rounded-xl bg-slate-50 focus:bg-white focus:outline-none focus:border-[#ff5500]"
                />
                {cashOutApprovalModal.error && (
                  <p className="text-xs text-rose-600 font-bold mt-1 text-center">{cashOutApprovalModal.error}</p>
                )}
                <p className="text-[10px] text-slate-400 text-center mt-1">Default Admin PIN: 1234</p>
              </div>

              <div className="grid grid-cols-2 gap-2 pt-1">
                {/* REJECT BUTTON */}
                <button
                  type="button"
                  onClick={() => {
                    const manager = staffList.find(s => (s.role === 'Administrator' || s.role === 'Manager') && s.pin === cashOutApprovalModal.managerPin.trim());
                    if (!manager) {
                      setCashOutApprovalModal(prev => ({ ...prev, error: 'Invalid Manager/Admin PIN' }));
                      return;
                    }

                    const targetItem = cashOutApprovalModal.item;
                    const updated = {
                      ...targetItem,
                      status: 'REJECTED',
                      rejectedBy: manager.name
                    };

                    setCurrentShift(prev => ({
                      ...prev,
                      payouts: prev.payouts.map(p => p.id === targetItem.id ? updated : p)
                    }));

                    setExpenses(prev => prev.map(p => p.id === targetItem.id ? updated : p));

                    recordAuditLog(
                      'CASH_OUT_REJECTED',
                      targetItem.id,
                      `Manager ${manager.name} rejected cash out request ${targetItem.id} of ${settings.currency} ${targetItem.amount}`
                    );

                    setCashOutApprovalModal({ open: false, item: null, managerPin: '', error: '' });
                  }}
                  className="py-2.5 bg-rose-50 hover:bg-rose-100 text-rose-700 border border-rose-200 rounded-xl text-xs font-bold transition-colors cursor-pointer"
                >
                  ✕ Reject
                </button>

                {/* APPROVE BUTTON */}
                <button
                  type="button"
                  onClick={() => {
                    const manager = staffList.find(s => (s.role === 'Administrator' || s.role === 'Manager') && s.pin === cashOutApprovalModal.managerPin.trim());
                    if (!manager) {
                      setCashOutApprovalModal(prev => ({ ...prev, error: 'Invalid Manager/Admin PIN' }));
                      return;
                    }

                    const targetItem = cashOutApprovalModal.item;
                    const updated = {
                      ...targetItem,
                      status: 'APPROVED',
                      approvedBy: manager.name,
                      approvedAt: new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })
                    };

                    setCurrentShift(prev => ({
                      ...prev,
                      payouts: prev.payouts.map(p => p.id === targetItem.id ? updated : p)
                    }));

                    setExpenses(prev => prev.map(p => p.id === targetItem.id ? updated : p));

                    // Auto-print thermal disbursement voucher upon manager sign-off
                    triggerAutoPrint({
                      type: 'CASH_OUT_VOUCHER',
                      data: updated
                    }, `Authorized Voucher #${targetItem.id}`);

                    recordAuditLog(
                      'CASH_OUT_APPROVED_PIN',
                      targetItem.id,
                      `Manager ${manager.name} authorized cash out ${targetItem.id} for ${settings.currency} ${targetItem.amount}`
                    );

                    setCashOutApprovalModal({ open: false, item: null, managerPin: '', error: '' });
                  }}
                  className="py-2.5 bg-emerald-600 hover:bg-emerald-700 text-white font-bold rounded-xl text-xs shadow-xs cursor-pointer transition-colors"
                >
                  ✓ Approve &amp; Print
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* MODAL: TABLE ALLOCATION WITH VISUAL TICKS */}
      {allocationModalOpen && (
        <div className="fixed inset-0 bg-black/75 backdrop-blur-xs z-50 flex items-center justify-center p-4">
          <div className="bg-white rounded-3xl w-full max-w-lg p-6 shadow-2xl border border-slate-200 max-h-[90vh] overflow-y-auto">
            <div className="flex items-center justify-between pb-3 border-b border-slate-100">
              <h3 className="text-base font-black text-slate-900 flex items-center gap-2">
                <Grid className="h-4 w-4 text-[#ff5500]" /> Assign Table or Order Mode
              </h3>
              <button onClick={() => setAllocationModalOpen(false)} className="text-slate-400 hover:text-slate-900 cursor-pointer">
                <X className="h-5 w-5" />
              </button>
            </div>

            <div className="mt-4 grid grid-cols-2 gap-3">
              <button
                type="button"
                onClick={() => setOrderMode('DINING')}
                className={`p-3 rounded-2xl border flex items-center justify-between transition-all cursor-pointer ${
                  orderMode === 'DINING'
                    ? 'border-[#ff5500] bg-orange-50/70 shadow-xs'
                    : 'border-slate-200 bg-white hover:bg-slate-50'
                }`}
              >
                <div className="flex items-center gap-2">
                  <Grid className="h-4 w-4 text-[#ff5500]" />
                  <span className="font-bold text-xs text-slate-900">Dine-In Tables</span>
                </div>
                <div className={`h-5 w-5 rounded-full flex items-center justify-center border ${
                  orderMode === 'DINING' ? 'bg-[#ff5500] border-[#ff5500] text-white' : 'border-slate-300'
                }`}>
                  {orderMode === 'DINING' && <Check className="h-3 w-3 stroke-[3]" />}
                </div>
              </button>

              <button
                type="button"
                onClick={() => setOrderMode('TAKEAWAY')}
                className={`p-3 rounded-2xl border flex items-center justify-between transition-all cursor-pointer ${
                  orderMode === 'TAKEAWAY'
                    ? 'border-[#ff5500] bg-orange-50/70 shadow-xs'
                    : 'border-slate-200 bg-white hover:bg-slate-50'
                }`}
              >
                <div className="flex items-center gap-2">
                  <ShoppingBag className="h-4 w-4 text-[#ff5500]" />
                  <span className="font-bold text-xs text-slate-900">Takeaway Express</span>
                </div>
                <div className={`h-5 w-5 rounded-full flex items-center justify-center border ${
                  orderMode === 'TAKEAWAY' ? 'bg-[#ff5500] border-[#ff5500] text-white' : 'border-slate-300'
                }`}>
                  {orderMode === 'TAKEAWAY' && <Check className="h-3 w-3 stroke-[3]" />}
                </div>
              </button>
            </div>

            {orderMode === 'DINING' ? (
              <div className="mt-4 space-y-3">
                <div className="flex items-center justify-between">
                  <span className="text-xs font-bold text-slate-700">Select Table:</span>
                  <div className="flex items-center gap-1.5 text-xs">
                    <span className="text-slate-500">Guests:</span>
                    <button
                      type="button"
                      onClick={() => setGuestCount(Math.max(1, guestCount - 1))}
                      className="h-6 w-6 rounded bg-slate-100 font-bold cursor-pointer"
                    >
                      -
                    </button>
                    <span className="font-mono font-bold w-5 text-center">{guestCount}</span>
                    <button
                      type="button"
                      onClick={() => setGuestCount(guestCount + 1)}
                      className="h-6 w-6 rounded bg-slate-100 font-bold cursor-pointer"
                    >
                      +
                    </button>
                  </div>
                </div>

                <div className="grid grid-cols-2 gap-2.5 max-h-60 overflow-y-auto pr-1">
                  {floorTables.map(tbl => {
                    const isSelected = selectedTable.id === tbl.id;
                    const isOccupied = tbl.status === 'OCCUPIED';
                    return (
                      <button
                        key={tbl.id}
                        type="button"
                        onClick={() => {
                          setSelectedTable(tbl);
                          setAllocationModalOpen(false);
                        }}
                        className={`p-3 rounded-2xl border text-left flex items-start justify-between transition-all cursor-pointer ${
                          isSelected
                            ? 'border-[#ff5500] bg-orange-50/80 ring-2 ring-orange-500/20 shadow-xs'
                            : 'border-slate-200 bg-white hover:bg-slate-50'
                        }`}
                      >
                        <div>
                          <div className="flex items-center gap-1.5">
                            <span className="font-black text-xs text-slate-900">{tbl.name}</span>
                            <span className={`text-[9px] font-bold px-1.5 py-0.2 rounded-full ${
                              isOccupied ? 'bg-orange-100 text-[#ff5500]' : 'bg-emerald-100 text-emerald-700'
                            }`}>
                              {tbl.status}
                            </span>
                          </div>
                          <p className="text-[10px] text-slate-500 mt-0.5">{tbl.zone}</p>
                          <span className="text-[10px] text-slate-400 font-mono mt-1 block">Capacity: {tbl.capacity} Seats</span>
                        </div>
                        <div className={`h-5 w-5 rounded-full flex items-center justify-center border shrink-0 transition-all ${
                          isSelected ? 'bg-[#ff5500] border-[#ff5500] text-white' : 'border-slate-300 bg-white'
                        }`}>
                          {isSelected && <Check className="h-3 w-3 stroke-[3]" />}
                        </div>
                      </button>
                    );
                  })}
                </div>
              </div>
            ) : (
              <div className="mt-4 p-4 bg-slate-50 rounded-2xl border border-slate-200 space-y-3">
                <div className="flex justify-between items-center">
                  <span className="text-xs font-bold text-slate-800">Express Token:</span>
                  <span className="font-mono font-bold text-xs bg-orange-100 text-[#ff5500] px-2 py-0.5 rounded-md">
                    {takeawayInfo.token}
                  </span>
                </div>
                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">Customer / Guest Name</label>
                  <input
                    type="text"
                    value={takeawayInfo.name}
                    onChange={(e) => setTakeawayInfo(prev => ({ ...prev, name: e.target.value }))}
                    className="w-full px-3 py-2 border border-slate-200 rounded-xl text-xs bg-white text-slate-900 focus:outline-none focus:border-[#ff5500]"
                  />
                </div>
                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">Contact Phone</label>
                  <input
                    type="text"
                    value={takeawayInfo.phone}
                    onChange={(e) => setTakeawayInfo(prev => ({ ...prev, phone: e.target.value }))}
                    placeholder="+94..."
                    className="w-full px-3 py-2 border border-slate-200 rounded-xl text-xs font-mono bg-white text-slate-900 focus:outline-none focus:border-[#ff5500]"
                  />
                </div>
                <button
                  type="button"
                  onClick={() => setAllocationModalOpen(false)}
                  className="w-full py-2.5 bg-[#ff5500] text-white font-bold rounded-xl text-xs cursor-pointer"
                >
                  Confirm Takeaway Order
                </button>
              </div>
            )}
          </div>
        </div>
      )}

      {/* MODAL: ADD / EDIT STAFF MEMBER */}
      {addStaffModalOpen && (
        <div 
          className="fixed inset-0 bg-black/80 backdrop-blur-xs z-[9999] flex items-center justify-center p-4 text-slate-900"
          onClick={(e) => {
            if (e.target === e.currentTarget) {
              setAddStaffModalOpen(false);
              setEditingStaffId(null);
              setStaffFormError('');
            }
          }}
        >
          <div className="bg-white rounded-3xl w-full max-w-lg p-6 shadow-2xl border border-slate-200 relative max-h-[90vh] overflow-y-auto">
            <div className="flex items-center justify-between pb-3 border-b border-slate-100">
              <div className="flex items-center gap-2">
                <Users className="h-5 w-5 text-[#ff5500]" />
                <h3 className="text-base font-black text-slate-900">
                  {editingStaffId ? 'Edit Staff Profile & Wages' : 'Add New Staff Member'}
                </h3>
              </div>
              <button
                type="button"
                onClick={() => {
                  setAddStaffModalOpen(false);
                  setEditingStaffId(null);
                  setStaffFormError('');
                }}
                className="text-slate-400 hover:text-slate-900 p-1 cursor-pointer"
              >
                <X className="h-5 w-5" />
              </button>
            </div>

            <form onSubmit={handleCreateStaff} className="mt-4 space-y-4">
              {/* Profile Details */}
              <div>
                <label className="block text-xs font-bold text-slate-700 mb-1">Full Name *</label>
                <input
                  type="text"
                  required
                  autoFocus
                  value={newStaffForm.name}
                  onChange={e => setNewStaffForm(prev => ({ ...prev, name: e.target.value }))}
                  placeholder="e.g. Kasun Fernando"
                  className="w-full px-3.5 py-2.5 bg-slate-50 border border-slate-200 rounded-xl text-xs font-semibold text-slate-900 focus:bg-white focus:outline-none focus:border-[#ff5500]"
                />
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">Role / Access Level</label>
                  <select
                    value={newStaffForm.role}
                    onChange={e => setNewStaffForm(prev => ({ ...prev, role: e.target.value }))}
                    className="w-full px-3 py-2.5 bg-white border border-slate-200 rounded-xl text-xs font-bold text-slate-900 focus:outline-none focus:border-[#ff5500]"
                  >
                    {Object.keys(ROLE_PERMISSIONS).map(role => (
                      <option key={role} value={role}>{role}</option>
                    ))}
                  </select>
                </div>

                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">4-Digit Security PIN *</label>
                  <input
                    type="password"
                    inputMode="numeric"
                    maxLength={4}
                    required
                    value={newStaffForm.pin}
                    onChange={e => {
                      const val = e.target.value.replace(/\D/g, '').slice(0, 4);
                      setNewStaffForm(prev => ({ ...prev, pin: val }));
                    }}
                    placeholder="e.g. 4321"
                    className="w-full px-3 py-2.5 bg-slate-50 border border-slate-200 rounded-xl text-center text-sm font-mono font-bold text-slate-900 tracking-widest focus:bg-white focus:outline-none focus:border-[#ff5500]"
                  />
                </div>
              </div>

              <div>
                <label className="block text-xs font-bold text-slate-700 mb-1">Email Address</label>
                <input
                  type="email"
                  value={newStaffForm.email}
                  onChange={e => setNewStaffForm(prev => ({ ...prev, email: e.target.value }))}
                  placeholder="e.g. kasun@linolicove.me"
                  className="w-full px-3.5 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs text-slate-900 focus:bg-white focus:outline-none focus:border-[#ff5500]"
                />
              </div>

              {/* SALARY & COMPENSATION SECTION */}
              <div className="p-3.5 bg-slate-50 rounded-2xl border border-slate-200 space-y-3">
                <span className="text-[10px] font-black uppercase tracking-wider text-slate-500 block">
                  Default Wage &amp; Allowance Structure ({settings.currency})
                </span>

                <div className="grid grid-cols-2 gap-3">
                  <div>
                    <label className="block text-[11px] font-bold text-slate-700 mb-1">Base Monthly Salary *</label>
                    <input
                      type="number"
                      step="0.01"
                      required
                      value={newStaffForm.basicSalary}
                      onChange={e => setNewStaffForm(prev => ({ ...prev, basicSalary: e.target.value }))}
                      className="w-full px-2.5 py-1.5 bg-white border border-slate-200 rounded-lg text-xs font-mono font-bold"
                    />
                  </div>

                  <div>
                    <label className="block text-[11px] font-bold text-slate-700 mb-1">Budgetary Relief (BRA)</label>
                    <input
                      type="number"
                      step="0.01"
                      value={newStaffForm.budgetaryAllowance}
                      onChange={e => setNewStaffForm(prev => ({ ...prev, budgetaryAllowance: e.target.value }))}
                      className="w-full px-2.5 py-1.5 bg-white border border-slate-200 rounded-lg text-xs font-mono font-bold"
                    />
                  </div>
                </div>

                <div className="grid grid-cols-3 gap-2.5">
                  <div>
                    <label className="block text-[10px] font-bold text-slate-600 mb-1">Fixed Allowances</label>
                    <input
                      type="number"
                      step="0.01"
                      value={newStaffForm.otherAllowances}
                      onChange={e => setNewStaffForm(prev => ({ ...prev, otherAllowances: e.target.value }))}
                      className="w-full px-2 py-1.5 bg-white border border-slate-200 rounded-lg text-xs font-mono font-bold"
                    />
                  </div>

                  <div>
                    <label className="block text-[10px] font-bold text-slate-600 mb-1">Fixed Bonus</label>
                    <input
                      type="number"
                      step="0.01"
                      value={newStaffForm.fixedBonus}
                      onChange={e => setNewStaffForm(prev => ({ ...prev, fixedBonus: e.target.value }))}
                      className="w-full px-2 py-1.5 bg-white border border-slate-200 rounded-lg text-xs font-mono font-bold"
                    />
                  </div>

                  <div>
                    <label className="block text-[10px] font-bold text-slate-600 mb-1">OT Rate / Hour</label>
                    <input
                      type="number"
                      step="1"
                      value={newStaffForm.overtimeRate}
                      onChange={e => setNewStaffForm(prev => ({ ...prev, overtimeRate: e.target.value }))}
                      className="w-full px-2 py-1.5 bg-white border border-slate-200 rounded-lg text-xs font-mono font-bold"
                    />
                  </div>
                </div>

                {/* EPF/ETF Enrolled Toggle */}
                <div className="flex items-center justify-between pt-2 border-t border-slate-200">
                  <div>
                    <span className="text-xs font-bold text-slate-800 block">Sri Lanka EPF / ETF Deduction</span>
                    <span className="text-[10px] text-slate-400">Employee 8% + Employer 12%/3%</span>
                  </div>
                  <button
                    type="button"
                    onClick={() => setNewStaffForm(prev => ({ ...prev, epfEtfEnabled: !prev.epfEtfEnabled }))}
                    className={`px-3 py-1 rounded-lg text-xs font-bold cursor-pointer transition-colors ${
                      newStaffForm.epfEtfEnabled !== false ? 'bg-emerald-600 text-white' : 'bg-slate-200 text-slate-700'
                    }`}
                  >
                    {newStaffForm.epfEtfEnabled !== false ? 'Enrolled' : 'Exempt'}
                  </button>
                </div>
              </div>

              {staffFormError && (
                <div className="p-2.5 rounded-xl bg-rose-50 border border-rose-200 text-rose-600 text-xs font-semibold text-center">
                  {staffFormError}
                </div>
              )}

              <div className="flex gap-2 pt-2 border-t border-slate-100">
                <button
                  type="button"
                  onClick={() => {
                    setAddStaffModalOpen(false);
                    setEditingStaffId(null);
                    setStaffFormError('');
                  }}
                  className="flex-1 py-2.5 bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold rounded-xl text-xs transition-colors cursor-pointer"
                >
                  Cancel
                </button>
                <button
                  type="submit"
                  className="flex-1 py-2.5 bg-[#ff5500] hover:bg-orange-600 text-white font-bold rounded-xl text-xs shadow-xs transition-all cursor-pointer active:scale-95"
                >
                  {editingStaffId ? 'Save Profile Changes' : 'Save Employee'}
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* MODAL: EDIT RAW MATERIAL INVENTORY */}
      {editInventoryModalOpen && editingInventoryItem && (
        <div className="fixed inset-0 bg-black/75 backdrop-blur-xs z-50 flex items-center justify-center p-4">
          <div className="bg-white rounded-3xl w-full max-w-md p-6 shadow-2xl border border-slate-200 max-h-[90vh] overflow-y-auto">
            <div className="flex items-center justify-between pb-3 border-b border-slate-100">
              <div className="flex items-center gap-2">
                <Package className="h-5 w-5 text-[#ff5500]" />
                <h3 className="text-base font-black text-slate-900">Edit Raw Material</h3>
              </div>
              <button
                type="button"
                onClick={() => {
                  setEditInventoryModalOpen(false);
                  setEditingInventoryItem(null);
                }}
                className="text-slate-400 hover:text-slate-900 cursor-pointer"
              >
                <X className="h-5 w-5" />
              </button>
            </div>

            <form
              onSubmit={(e) => {
                e.preventDefault();
                const costVal = parseFloat(editingInventoryItem.cost);
                const stockVal = parseFloat(editingInventoryItem.stock);
                const thresholdVal = parseFloat(editingInventoryItem.threshold);

                if (!editingInventoryItem.name.trim() || isNaN(costVal) || costVal <= 0) {
                  alert('Please enter a valid material name and unit cost.');
                  return;
                }

                const originalItem = inventoryMap[editingInventoryItem.id];
                const oldStock = originalItem ? Number(originalItem.stock) : Number(editingInventoryItem.stock);
                const targetStock = isNaN(stockVal) ? oldStock : stockVal;

                setInventory(prev => prev.map(item => item.id === editingInventoryItem.id ? {
                  ...editingInventoryItem,
                  name: editingInventoryItem.name.trim(),
                  cost: costVal,
                  stock: targetStock,
                  threshold: isNaN(thresholdVal) ? item.threshold : thresholdVal
                } : item));

                recordAuditLog(
                  'INVENTORY_ITEM_MODIFIED',
                  editingInventoryItem.id,
                  `Updated raw material "${editingInventoryItem.name.trim()}": Stock: ${targetStock} ${editingInventoryItem.unit}, Cost: ${settings.currency} ${costVal.toFixed(2)}/${editingInventoryItem.unit}, Alert: ${thresholdVal} ${editingInventoryItem.unit}`
                );

                if (oldStock !== targetStock && typeof recordStockMovement === 'function') {
                  const diff = Number((targetStock - oldStock).toFixed(2));
                  recordStockMovement(
                    'MANUAL_EDIT',
                    editingInventoryItem.id,
                    editingInventoryItem.name.trim(),
                    diff,
                    oldStock,
                    targetStock,
                    editingInventoryItem.unit,
                    costVal,
                    'Manual stock adjustment in inventory editor',
                    'ADMIN_EDIT'
                  );
                }

                setEditInventoryModalOpen(false);
                setEditingInventoryItem(null);
              }}
              className="mt-4 space-y-4"
            >
              <div>
                <label className="block text-xs font-bold text-slate-700 mb-1">Material Name</label>
                <input
                  type="text"
                  required
                  value={editingInventoryItem.name}
                  onChange={e => setEditingInventoryItem(prev => ({ ...prev, name: e.target.value }))}
                  className="w-full px-3.5 py-2.5 bg-slate-50 border border-slate-200 rounded-xl text-xs font-semibold text-slate-900 focus:bg-white focus:outline-none focus:border-[#ff5500]"
                />
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">Category</label>
                  <select
                    value={editingInventoryItem.category}
                    onChange={e => setEditingInventoryItem(prev => ({ ...prev, category: e.target.value }))}
                    className="w-full px-3 py-2.5 bg-slate-50 border border-slate-200 rounded-xl text-xs font-bold text-slate-900 focus:bg-white focus:outline-none focus:border-[#ff5500]"
                    style={{ color: '#0f172a', backgroundColor: '#ffffff' }}
                  >
                    <option value="Dry Goods">Dry Goods</option>
                    <option value="Dairy & Eggs">Dairy &amp; Eggs</option>
                    <option value="Meat">Meat</option>
                    <option value="Poultry">Poultry</option>
                    <option value="Seafood">Seafood</option>
                    <option value="Beverages">Beverages</option>
                    <option value="Bar Supplies">Bar Supplies</option>
                    <option value="Bakery">Bakery</option>
                    <option value="Produce">Produce</option>
                  </select>
                </div>

                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">Unit of Measure</label>
                  <select
                    value={editingInventoryItem.unit}
                    onChange={e => setEditingInventoryItem(prev => ({ ...prev, unit: e.target.value }))}
                    className="w-full px-3 py-2.5 bg-slate-50 border border-slate-200 rounded-xl text-xs font-bold text-slate-900 focus:bg-white focus:outline-none focus:border-[#ff5500]"
                    style={{ color: '#0f172a', backgroundColor: '#ffffff' }}
                  >
                    <option value="g">Grams (g)</option>
                    <option value="ml">Milliliters (ml)</option>
                    <option value="pcs">Pieces (pcs)</option>
                    <option value="kg">Kilograms (kg)</option>
                    <option value="l">Liters (l)</option>
                  </select>
                </div>
              </div>

              <div className="grid grid-cols-3 gap-3">
                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">Unit Cost ({settings.currency})</label>
                  <input
                    type="number"
                    step="0.01"
                    min="0.01"
                    required
                    value={editingInventoryItem.cost}
                    onChange={e => setEditingInventoryItem(prev => ({ ...prev, cost: e.target.value }))}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-mono font-bold text-slate-900 focus:bg-white focus:outline-none focus:border-[#ff5500]"
                  />
                </div>

                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">Stock Level</label>
                  <input
                    type="number"
                    step="any"
                    min="0"
                    required
                    value={editingInventoryItem.stock}
                    onChange={e => setEditingInventoryItem(prev => ({ ...prev, stock: e.target.value }))}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-mono font-bold text-slate-900 focus:bg-white focus:outline-none focus:border-[#ff5500]"
                  />
                </div>

                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">Low Alert</label>
                  <input
                    type="number"
                    step="any"
                    min="1"
                    required
                    value={editingInventoryItem.threshold}
                    onChange={e => setEditingInventoryItem(prev => ({ ...prev, threshold: e.target.value }))}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-mono font-bold text-slate-900 focus:bg-white focus:outline-none focus:border-[#ff5500]"
                  />
                </div>
              </div>

              <div className="flex gap-2 pt-2 border-t border-slate-100">
                <button
                  type="button"
                  onClick={() => {
                    setEditInventoryModalOpen(false);
                    setEditingInventoryItem(null);
                  }}
                  className="flex-1 py-2.5 bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold rounded-xl text-xs transition-colors cursor-pointer"
                >
                  Cancel
                </button>
                <button
                  type="submit"
                  className="flex-1 py-2.5 bg-[#ff5500] hover:bg-orange-600 text-white font-bold rounded-xl text-xs shadow-xs transition-all cursor-pointer"
                >
                  Save Material Changes
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* MODAL: RECEIVE STOCK INTAKE (GRN) */}
      {receiveStockModalOpen && (
        <div className="fixed inset-0 bg-black/75 backdrop-blur-xs z-50 flex items-center justify-center p-4">
          <div className="bg-white rounded-3xl w-full max-w-md p-6 shadow-2xl border border-slate-200 max-h-[90vh] overflow-y-auto">
            <div className="flex items-center justify-between pb-3 border-b border-slate-100">
              <div className="flex items-center gap-2">
                <Package className="h-5 w-5 text-emerald-600" />
                <h3 className="text-base font-black text-slate-900">Receive Stock Intake (GRN)</h3>
              </div>
              <button
                type="button"
                onClick={() => setReceiveStockModalOpen(false)}
                className="text-slate-400 hover:text-slate-900 cursor-pointer"
              >
                <X className="h-5 w-5" />
              </button>
            </div>

            <form onSubmit={handleReceiveStock} className="mt-4 space-y-4">
              <div>
                <label className="block text-xs font-bold text-slate-700 mb-1">Select Material / Ingredient</label>
                <select
                  required
                  value={receiveStockForm.ingredientId}
                  onChange={e => {
                    const id = e.target.value;
                    const item = inventoryMap[id];
                    setReceiveStockForm(prev => ({
                      ...prev,
                      ingredientId: id,
                      newCost: item ? item.cost.toString() : ''
                    }));
                  }}
                  className="w-full px-3 py-2.5 bg-slate-50 border border-slate-200 rounded-xl text-xs font-bold text-slate-900 focus:bg-white focus:outline-none focus:border-[#ff5500]"
                  style={{ color: '#0f172a', backgroundColor: '#ffffff' }}
                >
                  {inventory.map(item => (
                    <option key={item.id} value={item.id} style={{ color: '#0f172a', backgroundColor: '#ffffff' }}>
                      {item.name} ({item.stock} {item.unit} available)
                    </option>
                  ))}
                </select>
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">
                    Qty Received ({inventoryMap[receiveStockForm.ingredientId]?.unit || 'units'})
                  </label>
                  <input
                    type="number"
                    step="any"
                    min="0.01"
                    required
                    value={receiveStockForm.quantity}
                    onChange={e => setReceiveStockForm(prev => ({ ...prev, quantity: e.target.value }))}
                    placeholder="e.g. 5000"
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-mono font-bold text-slate-900 focus:bg-white focus:outline-none focus:border-[#ff5500]"
                  />
                </div>

                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">
                    Updated Unit Cost ({settings.currency})
                  </label>
                  <input
                    type="number"
                    step="0.01"
                    min="0.01"
                    value={receiveStockForm.newCost}
                    onChange={e => setReceiveStockForm(prev => ({ ...prev, newCost: e.target.value }))}
                    placeholder="Current cost"
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-mono font-bold text-slate-900 focus:bg-white focus:outline-none focus:border-[#ff5500]"
                  />
                </div>
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">Supplier / Vendor</label>
                  <input
                    type="text"
                    value={receiveStockForm.supplier}
                    onChange={e => setReceiveStockForm(prev => ({ ...prev, supplier: e.target.value }))}
                    placeholder="e.g. Colombo Central Market"
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs text-slate-900 focus:bg-white focus:outline-none focus:border-[#ff5500]"
                  />
                </div>

                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">Invoice / Ref #</label>
                  <input
                    type="text"
                    value={receiveStockForm.invoiceRef}
                    onChange={e => setReceiveStockForm(prev => ({ ...prev, invoiceRef: e.target.value }))}
                    placeholder="e.g. GRN-804"
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-mono text-slate-900 focus:bg-white focus:outline-none focus:border-[#ff5500]"
                  />
                </div>
              </div>

              {receiveStockForm.ingredientId && receiveStockForm.quantity && (
                <div className="p-3 bg-emerald-50 rounded-xl border border-emerald-200 text-xs text-emerald-800 space-y-1">
                  <p className="font-bold flex items-center gap-1">
                    <CheckCircle2 className="h-3.5 w-3.5 text-emerald-600" />
                    Intake Projection:
                  </p>
                  <p>
                    Material: <strong>{inventoryMap[receiveStockForm.ingredientId]?.name}</strong>
                  </p>
                  <p className="font-mono">
                    Stock: {inventoryMap[receiveStockForm.ingredientId]?.stock} →{' '}
                    <strong>
                      {Number(((inventoryMap[receiveStockForm.ingredientId]?.stock || 0) + (parseFloat(receiveStockForm.quantity) || 0)).toFixed(2))}{' '}
                      {inventoryMap[receiveStockForm.ingredientId]?.unit}
                    </strong>
                  </p>
                </div>
              )}

              <div className="flex gap-2 pt-2">
                <button
                  type="button"
                  onClick={() => setReceiveStockModalOpen(false)}
                  className="flex-1 py-2.5 bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold rounded-xl text-xs transition-colors cursor-pointer"
                >
                  Cancel
                </button>
                <button
                  type="submit"
                  className="flex-1 py-2.5 bg-emerald-600 hover:bg-emerald-700 text-white font-bold rounded-xl text-xs shadow-xs transition-all cursor-pointer"
                >
                  Confirm Intake
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* MODAL: ADD NEW FLOOR TABLE */}
      {addTableModalOpen && (
        <div className="fixed inset-0 bg-black/75 backdrop-blur-xs z-50 flex items-center justify-center p-4">
          <div className="bg-white rounded-3xl w-full max-w-sm p-6 shadow-2xl border border-slate-200">
            <div className="flex items-center justify-between pb-3 border-b border-slate-100">
              <div className="flex items-center gap-2">
                <Grid className="h-5 w-5 text-[#ff5500]" />
                <h3 className="text-base font-black text-slate-900">Add New Floor Table</h3>
              </div>
              <button
                type="button"
                onClick={() => {
                  setAddTableModalOpen(false);
                  setNewTableForm({ name: '', zone: 'Indoor Main Hall', capacity: 4 });
                }}
                className="text-slate-400 hover:text-slate-900 cursor-pointer"
              >
                <X className="h-5 w-5" />
              </button>
            </div>

            <form onSubmit={handleCreateTable} className="mt-4 space-y-4">
              <div>
                <label className="block text-xs font-bold text-slate-700 mb-1">Table Name / Identifier</label>
                <input
                  type="text"
                  required
                  autoFocus
                  value={newTableForm.name}
                  onChange={e => setNewTableForm(prev => ({ ...prev, name: e.target.value }))}
                  placeholder="e.g. Table 6, Cabana 2, Bar Seat 03"
                  className="w-full px-3.5 py-2.5 bg-slate-50 border border-slate-200 rounded-xl text-xs font-bold text-slate-900 focus:bg-white focus:outline-none focus:border-[#ff5500]"
                />
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">Floor Zone</label>
                  <select
                    value={newTableForm.zone}
                    onChange={e => setNewTableForm(prev => ({ ...prev, zone: e.target.value }))}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-bold text-slate-900 focus:bg-white focus:outline-none focus:border-[#ff5500]"
                    style={{ color: '#0f172a', backgroundColor: '#ffffff' }}
                  >
                    <option value="Indoor Main Hall">Indoor Main Hall</option>
                    <option value="Deck Lounge">Deck Lounge</option>
                    <option value="Cocktail Counter">Cocktail Counter</option>
                    <option value="Private Ocean View">Private Ocean View</option>
                    <option value="Beachfront Garden">Beachfront Garden</option>
                    <option value="Rooftop Terrace">Rooftop Terrace</option>
                  </select>
                </div>

                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">Seating Capacity</label>
                  <input
                    type="number"
                    min="1"
                    max="50"
                    required
                    value={newTableForm.capacity}
                    onChange={e => setNewTableForm(prev => ({ ...prev, capacity: e.target.value }))}
                    placeholder="4"
                    className="w-full px-3.5 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-mono font-bold text-slate-900 focus:bg-white focus:outline-none focus:border-[#ff5500]"
                  />
                </div>
              </div>

              <div className="flex gap-2 pt-2">
                <button
                  type="button"
                  onClick={() => {
                    setAddTableModalOpen(false);
                    setNewTableForm({ name: '', zone: 'Indoor Main Hall', capacity: 4 });
                  }}
                  className="flex-1 py-2.5 bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold rounded-xl text-xs transition-colors cursor-pointer"
                >
                  Cancel
                </button>
                <button
                  type="submit"
                  className="flex-1 py-2.5 bg-[#ff5500] hover:bg-orange-600 text-white font-bold rounded-xl text-xs shadow-xs transition-all cursor-pointer active:scale-95"
                >
                  Create Table
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* MODAL: ADD NEW RAW MATERIAL */}
      {addInventoryModalOpen && (
        <div 
          className="fixed inset-0 bg-black/80 backdrop-blur-xs z-[9999] flex items-center justify-center p-4 text-slate-900"
          onClick={(e) => {
            if (e.target === e.currentTarget) setAddInventoryModalOpen(false);
          }}
        >
          <div className="bg-white rounded-3xl w-full max-w-md p-6 shadow-2xl border border-slate-200 relative max-h-[90vh] overflow-y-auto">
            <div className="flex items-center justify-between pb-3 border-b border-slate-100">
              <div className="flex items-center gap-2">
                <Package className="h-5 w-5 text-[#ff5500]" />
                <h3 className="text-base font-black text-slate-900">Add Raw Material to Inventory</h3>
              </div>
              <button
                type="button"
                onClick={() => setAddInventoryModalOpen(false)}
                className="text-slate-400 hover:text-slate-900 cursor-pointer p-1"
              >
                <X className="h-5 w-5" />
              </button>
            </div>

            <form onSubmit={handleCreateInventoryItem} className="mt-4 space-y-4">
              <div>
                <label className="block text-xs font-bold text-slate-700 mb-1">Material Name *</label>
                <input
                  type="text"
                  required
                  autoFocus
                  value={newInventoryForm.name || ''}
                  onChange={e => setNewInventoryForm(prev => ({ ...prev, name: e.target.value }))}
                  placeholder="e.g. Basmati Rice, Fresh Lime, Single Malt"
                  className="w-full px-3.5 py-2.5 bg-slate-50 border border-slate-200 rounded-xl text-xs font-semibold text-slate-900 focus:bg-white focus:outline-none focus:border-[#ff5500]"
                />
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">Category</label>
                  <select
                    value={newInventoryForm.category || 'Dry Goods'}
                    onChange={e => setNewInventoryForm(prev => ({ ...prev, category: e.target.value }))}
                    className="w-full px-3 py-2.5 bg-white border border-slate-200 rounded-xl text-xs font-bold text-slate-900 focus:outline-none focus:border-[#ff5500]"
                    style={{ color: '#0f172a', backgroundColor: '#ffffff' }}
                  >
                    <option value="Dry Goods">Dry Goods</option>
                    <option value="Dairy & Eggs">Dairy &amp; Eggs</option>
                    <option value="Meat">Meat</option>
                    <option value="Poultry">Poultry</option>
                    <option value="Seafood">Seafood</option>
                    <option value="Beverages">Beverages</option>
                    <option value="Bar Supplies">Bar Supplies</option>
                    <option value="Bakery">Bakery</option>
                    <option value="Produce">Produce</option>
                  </select>
                </div>

                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">Unit</label>
                  <select
                    value={newInventoryForm.unit || 'g'}
                    onChange={e => setNewInventoryForm(prev => ({ ...prev, unit: e.target.value }))}
                    className="w-full px-3 py-2.5 bg-white border border-slate-200 rounded-xl text-xs font-bold text-slate-900 focus:outline-none focus:border-[#ff5500]"
                    style={{ color: '#0f172a', backgroundColor: '#ffffff' }}
                  >
                    <option value="g">Grams (g)</option>
                    <option value="ml">Milliliters (ml)</option>
                    <option value="pcs">Pieces (pcs)</option>
                    <option value="kg">Kilograms (kg)</option>
                    <option value="l">Liters (l)</option>
                  </select>
                </div>
              </div>

              <div className="grid grid-cols-3 gap-3">
                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">Unit Cost ({settings.currency}) *</label>
                  <input
                    type="number"
                    step="0.01"
                    min="0.01"
                    required
                    value={newInventoryForm.cost || ''}
                    onChange={e => setNewInventoryForm(prev => ({ ...prev, cost: e.target.value }))}
                    placeholder="150.00"
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-mono font-bold text-slate-900 focus:bg-white focus:outline-none focus:border-[#ff5500]"
                  />
                </div>

                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">Initial Stock</label>
                  <input
                    type="number"
                    step="any"
                    min="0"
                    value={newInventoryForm.stock || ''}
                    onChange={e => setNewInventoryForm(prev => ({ ...prev, stock: e.target.value }))}
                    placeholder="1000"
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-mono font-bold text-slate-900 focus:bg-white focus:outline-none focus:border-[#ff5500]"
                  />
                </div>

                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">Low Alert</label>
                  <input
                    type="number"
                    step="any"
                    min="1"
                    value={newInventoryForm.threshold || ''}
                    onChange={e => setNewInventoryForm(prev => ({ ...prev, threshold: e.target.value }))}
                    placeholder="100"
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-mono font-bold text-slate-900 focus:bg-white focus:outline-none focus:border-[#ff5500]"
                  />
                </div>
              </div>

              <div className="flex gap-2 pt-2 border-t border-slate-100">
                <button
                  type="button"
                  onClick={() => setAddInventoryModalOpen(false)}
                  className="flex-1 py-2.5 bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold rounded-xl text-xs transition-colors cursor-pointer"
                >
                  Cancel
                </button>
                <button
                  type="submit"
                  className="flex-1 py-2.5 bg-[#ff5500] hover:bg-orange-600 text-white font-bold rounded-xl text-xs shadow-md transition-all cursor-pointer active:scale-95"
                >
                  Save Material
                </button>
              </div>
            </form>
          </div>
        </div>
      )}
      
      {/* MODAL: CONFIGURE DISH RECIPE & BOM PORTIONS */}
      {recipeConfigModalOpen && editingDishForRecipe && (
        <div
          className="fixed inset-0 bg-black/80 backdrop-blur-xs z-[9999] flex items-center justify-center p-4 text-slate-900"
          onClick={(e) => {
            if (e.target === e.currentTarget) {
              setRecipeConfigModalOpen(false);
              setEditingDishForRecipe(null);
            }
          }}
        >
          <div className="bg-white rounded-3xl w-full max-w-lg p-6 shadow-2xl border border-slate-200 relative max-h-[90vh] overflow-y-auto space-y-4">
            <div className="flex items-center justify-between pb-3 border-b border-slate-100">
              <div className="flex items-center gap-2">
                <Sliders className="h-5 w-5 text-[#ff5500]" />
                <div>
                  <h3 className="text-base font-black text-slate-900">
                    Configure Recipe: {editingDishForRecipe.name}
                  </h3>
                  <p className="text-xs text-slate-500 font-medium">
                    Link raw inventory ingredients and set portion requirements
                  </p>
                </div>
              </div>
              <button
                type="button"
                onClick={() => {
                  setRecipeConfigModalOpen(false);
                  setEditingDishForRecipe(null);
                }}
                className="text-slate-400 hover:text-slate-900 p-1 cursor-pointer"
              >
                <X className="h-5 w-5" />
              </button>
            </div>

            {/* Live Financial & Portion Projection */}
            {(() => {
              const { cogs, portions } = calculateDishAvailability(currentRecipeIngredients);
              const margin =
                editingDishForRecipe.price > 0
                  ? (((editingDishForRecipe.price - cogs) / editingDishForRecipe.price) * 100).toFixed(1)
                  : 0;

              return (
                <div className="grid grid-cols-3 gap-2.5 p-3 bg-slate-50 border border-slate-200 rounded-2xl text-xs">
                  <div>
                    <span className="text-[10px] font-bold text-slate-400 uppercase">Selling Price</span>
                    <p className="font-mono font-black text-[#ff5500] text-sm mt-0.5">
                      {settings.currency} {editingDishForRecipe.price.toFixed(2)}
                    </p>
                  </div>
                  <div>
                    <span className="text-[10px] font-bold text-slate-400 uppercase">Total COGS</span>
                    <p className="font-mono font-black text-slate-900 text-sm mt-0.5">
                      {settings.currency} {cogs.toFixed(2)}
                    </p>
                  </div>
                  <div>
                    <span className="text-[10px] font-bold text-slate-400 uppercase">Gross Margin</span>
                    <p className="font-mono font-black text-emerald-600 text-sm mt-0.5">
                      {margin}% ({portions} left)
                    </p>
                  </div>
                </div>
              );
            })()}

            {/* Add Ingredient Bar */}
            <div className="p-3.5 bg-slate-50 border border-slate-200 rounded-2xl space-y-2">
              <span className="text-xs font-bold text-slate-800 flex items-center gap-1.5">
                <BookOpen className="h-3.5 w-3.5 text-[#ff5500]" /> Select Raw Material to Add
              </span>
              <div className="flex gap-2">
                <select
                  value={tempIngredientSelect.ingredientId || inventory[0]?.id || ''}
                  onChange={(e) =>
                    setTempIngredientSelect((prev) => ({ ...prev, ingredientId: e.target.value }))
                  }
                  className="flex-1 px-3 py-2 border border-slate-200 rounded-xl text-xs bg-white text-slate-900 focus:outline-none focus:border-[#ff5500]"
                  style={{ color: '#0f172a', backgroundColor: '#ffffff' }}
                >
                  {inventory.map((ing) => (
                    <option key={ing.id} value={ing.id} style={{ color: '#0f172a', backgroundColor: '#ffffff' }}>
                      {ing.name} ({ing.stock} {ing.unit} in stock)
                    </option>
                  ))}
                </select>

                <input
                  type="number"
                  min="0.01"
                  step="any"
                  value={tempIngredientSelect.amount}
                  onChange={(e) =>
                    setTempIngredientSelect((prev) => ({ ...prev, amount: e.target.value }))
                  }
                  placeholder="Qty/portion"
                  className="w-28 px-3 py-2 border border-slate-200 rounded-xl text-xs font-mono bg-white text-slate-900 focus:outline-none focus:border-[#ff5500]"
                />

                <button
                  type="button"
                  onClick={handleAddIngredientToRecipe}
                  className="px-4 py-2 bg-slate-900 hover:bg-slate-800 text-white rounded-xl text-xs font-bold transition-all cursor-pointer shadow-xs active:scale-95 shrink-0"
                >
                  + Add
                </button>
              </div>
            </div>

            {/* Ingredients Table */}
            <div className="space-y-2">
              <span className="font-black text-slate-400 uppercase tracking-wider text-[10px]">
                Linked BOM Ingredients ({currentRecipeIngredients.length})
              </span>

              {currentRecipeIngredients.length === 0 ? (
                <div className="p-6 text-center text-slate-400 italic text-xs border border-dashed border-slate-200 rounded-2xl">
                  No ingredients configured for this dish yet. Select an ingredient above and tap &ldquo;+ Add&rdquo;.
                </div>
              ) : (
                <div className="space-y-2 max-h-56 overflow-y-auto pr-1">
                  {currentRecipeIngredients.map((item, idx) => {
                    const ing = inventoryMap[item.ingredientId];
                    const lineCost = ing ? ing.cost * item.amount : 0;
                    return (
                      <div
                        key={`${item.ingredientId}_${idx}`}
                        className="flex items-center justify-between p-3 bg-slate-50 border border-slate-200 rounded-xl text-xs"
                      >
                        <div>
                          <p className="font-bold text-slate-900">{ing?.name || item.ingredientId}</p>
                          <span className="text-[10px] text-slate-500 font-mono">
                            Cost: {settings.currency} {lineCost.toFixed(2)} ({settings.currency} {ing?.cost || 0}/{ing?.unit || 'unit'})
                          </span>
                        </div>

                        <div className="flex items-center gap-2">
                          <span className="px-2.5 py-1 bg-white border border-slate-200 rounded-lg font-mono font-bold text-slate-800">
                            {item.amount} {ing?.unit || 'units'}
                          </span>
                          <button
                            type="button"
                            onClick={() => handleRemoveIngredientFromRecipe(idx)}
                            className="p-1.5 text-slate-400 hover:text-rose-600 rounded-lg hover:bg-rose-50 cursor-pointer transition-colors"
                            title="Remove Ingredient"
                          >
                            <Trash2 className="h-4 w-4" />
                          </button>
                        </div>
                      </div>
                    );
                  })}
                </div>
              )}
            </div>

            <div className="flex gap-2 pt-3 border-t border-slate-100">
              <button
                type="button"
                onClick={() => {
                  setRecipeConfigModalOpen(false);
                  setEditingDishForRecipe(null);
                }}
                className="flex-1 py-2.5 bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold rounded-xl text-xs cursor-pointer transition-colors"
              >
                Cancel
              </button>
              <button
                type="button"
                onClick={handleSaveRecipeConfiguration}
                className="flex-1 py-2.5 bg-[#ff5500] hover:bg-orange-600 text-white font-bold rounded-xl text-xs shadow-xs cursor-pointer active:scale-95 transition-all"
              >
                Save Recipe BOM
              </button>
            </div>
          </div>
        </div>
      )}

      {/* MODAL: ADD VENDOR INVOICE */}
        {addVendorBillModalOpen && (
          <div 
            className="fixed inset-0 bg-black/80 backdrop-blur-xs z-[9999] flex items-center justify-center p-4 text-slate-900"
            onClick={(e) => {
              if (e.target === e.currentTarget) setAddVendorBillModalOpen(false);
            }}
          >
            <div className="bg-white rounded-3xl w-full max-w-lg p-6 shadow-2xl border border-slate-200 relative max-h-[90vh] overflow-y-auto">
              <div className="flex items-center justify-between pb-3 border-b border-slate-100">
                <div className="flex items-center gap-2">
                  <FileText className="h-5 w-5 text-[#ff5500]" />
                  <h3 className="text-base font-black text-slate-900">Record Vendor Bill / External Invoice</h3>
                </div>
                <button
                  type="button"
                  onClick={() => setAddVendorBillModalOpen(false)}
                  className="text-slate-400 hover:text-slate-900 cursor-pointer p-1"
                >
                  <X className="h-5 w-5" />
                </button>
              </div>

              <form
                onSubmit={(e) => {
                  e.preventDefault();
                  const amt = parseFloat(vendorBillForm.amount);
                  if (isNaN(amt) || amt <= 0 || !vendorBillForm.vendorName.trim()) {
                    alert('Please enter a valid vendor name and bill amount.');
                    return;
                  }

                  const newBill = {
                    id: `BILL-${Date.now().toString().slice(-6)}`,
                    invoiceNumber: vendorBillForm.invoiceNumber.trim() || `INV-${Date.now().toString().slice(-4)}`,
                    vendorName: vendorBillForm.vendorName.trim(),
                    category: vendorBillForm.category,
                    billDate: vendorBillForm.billDate,
                    dueDate: vendorBillForm.dueDate,
                    amount: amt,
                    paymentMethod: vendorBillForm.paymentMethod,
                    paymentStatus: vendorBillForm.paymentStatus,
                    notes: vendorBillForm.notes.trim(),
                    recordedBy: currentUser.name,
                    recordedAt: new Date().toISOString()
                  };

                  setVendorBills(prev => [newBill, ...prev]);

                  if (typeof appendCloudArchive === 'function') {
                    appendCloudArchive('vendor_bills_archive', newBill);
                  }

                  recordAuditLog(
                    'VENDOR_BILL_CREATED',
                    newBill.id,
                    `Recorded invoice #${newBill.invoiceNumber} from ${newBill.vendorName} for ${settings.currency} ${amt.toFixed(2)} (${newBill.paymentStatus})`
                  );

                  setAddVendorBillModalOpen(false);
                }}
                className="mt-4 space-y-4"
              >
                <div className="grid grid-cols-2 gap-3">
                  <div>
                    <label className="block text-xs font-bold text-slate-700 mb-1">Invoice / Ref #</label>
                    <input
                      type="text"
                      value={vendorBillForm.invoiceNumber}
                      onChange={e => setVendorBillForm(prev => ({ ...prev, invoiceNumber: e.target.value }))}
                      placeholder="e.g. INV-90412"
                      className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-mono font-bold text-slate-900 focus:bg-white focus:outline-none"
                    />
                  </div>

                  <div>
                    <label className="block text-xs font-bold text-slate-700 mb-1">Vendor / Payee *</label>
                    <input
                      type="text"
                      required
                      value={vendorBillForm.vendorName}
                      onChange={e => setVendorBillForm(prev => ({ ...prev, vendorName: e.target.value }))}
                      placeholder="e.g. Ceylon Cold Stores / CEB"
                      className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-bold text-slate-900 focus:bg-white focus:outline-none"
                    />
                  </div>
                </div>

                <div className="grid grid-cols-2 gap-3">
                  <div>
                    <label className="block text-xs font-bold text-slate-700 mb-1">Category</label>
                    <select
                      value={vendorBillForm.category}
                      onChange={e => setVendorBillForm(prev => ({ ...prev, category: e.target.value }))}
                      className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-bold text-slate-900 focus:bg-white focus:outline-none"
                    >
                      <option value="Food & Beverage Supply">Food &amp; Beverage Supply</option>
                      <option value="Liquor & Bar Supply">Liquor &amp; Bar Supply</option>
                      <option value="Electricity / Utilities">Electricity / Utilities</option>
                      <option value="Water Services">Water Services</option>
                      <option value="Rent & Property">Rent &amp; Property</option>
                      <option value="Equipment Maintenance">Equipment Maintenance</option>
                      <option value="Marketing & Software">Marketing &amp; Software</option>
                      <option value="Other Operating Expense">Other Operating Expense</option>
                    </select>
                  </div>

                  <div>
                    <label className="block text-xs font-bold text-slate-700 mb-1">Total Bill Amount ({settings.currency}) *</label>
                    <input
                      type="number"
                      step="0.01"
                      min="1"
                      required
                      value={vendorBillForm.amount}
                      onChange={e => setVendorBillForm(prev => ({ ...prev, amount: e.target.value }))}
                      placeholder="e.g. 45000.00"
                      className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-mono font-bold text-slate-900 focus:bg-white focus:outline-none"
                    />
                  </div>
                </div>

                <div className="grid grid-cols-2 gap-3">
                  <div>
                    <label className="block text-xs font-bold text-slate-700 mb-1">Invoice Date</label>
                    <input
                      type="date"
                      required
                      value={vendorBillForm.billDate}
                      onChange={e => setVendorBillForm(prev => ({ ...prev, billDate: e.target.value }))}
                      className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-mono text-slate-900 focus:bg-white focus:outline-none"
                    />
                  </div>

                  <div>
                    <label className="block text-xs font-bold text-slate-700 mb-1">Payment Method</label>
                    <select
                      value={vendorBillForm.paymentMethod}
                      onChange={e => setVendorBillForm(prev => ({ ...prev, paymentMethod: e.target.value }))}
                      className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-bold text-slate-900 focus:bg-white focus:outline-none"
                    >
                      <option value="BANK_TRANSFER">Direct Bank Transfer</option>
                      <option value="CHEQUE">Company Cheque</option>
                      <option value="CREDIT_CARD">Company Credit Card</option>
                      <option value="ONLINE_PAYMENT">Online / Portal Payment</option>
                    </select>
                  </div>
                </div>

                <div className="flex items-center justify-between p-3 bg-slate-50 border border-slate-200 rounded-xl">
                  <div>
                    <span className="text-xs font-bold text-slate-800 block">Payment Status</span>
                    <span className="text-[10px] text-slate-400">Paid bills are immediately booked into P&amp;L OPEX</span>
                  </div>
                  <div className="flex gap-1.5">
                    {['PAID', 'UNPAID'].map(status => (
                      <button
                        key={status}
                        type="button"
                        onClick={() => setVendorBillForm(prev => ({ ...prev, paymentStatus: status }))}
                        className={`px-3 py-1 rounded-lg text-xs font-bold transition-colors cursor-pointer ${
                          vendorBillForm.paymentStatus === status
                            ? 'bg-[#ff5500] text-white'
                            : 'bg-white border border-slate-200 text-slate-700'
                        }`}
                      >
                        {status}
                      </button>
                    ))}
                  </div>
                </div>

                <div className="flex gap-2 pt-2 border-t border-slate-100">
                  <button
                    type="button"
                    onClick={() => setAddVendorBillModalOpen(false)}
                    className="flex-1 py-2.5 bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold rounded-xl text-xs cursor-pointer"
                  >
                    Cancel
                  </button>
                  <button
                    type="submit"
                    className="flex-1 py-2.5 bg-[#ff5500] hover:bg-orange-600 text-white font-bold rounded-xl text-xs shadow-xs cursor-pointer"
                  >
                    Save &amp; Sync to Cloud
                  </button>
                </div>
              </form>
            </div>
          </div>
        )}
        
      {/* MODAL: SRI LANKAN STATUTORY PAYROLL GENERATOR */}
      {processPayModalOpen && (
        <div 
          className="fixed inset-0 bg-black/80 backdrop-blur-xs z-[9999] flex items-center justify-center p-4 text-slate-900"
          onClick={(e) => {
            if (e.target === e.currentTarget) {
              setProcessPayModalOpen(false);
              setEditingPayrollId(null);
            }
          }}
        >
          <div className="bg-white rounded-3xl w-full max-w-lg p-6 shadow-2xl border border-slate-200 relative max-h-[90vh] overflow-y-auto">
            <div className="flex items-center justify-between pb-3 border-b border-slate-100">
              <div className="flex items-center gap-2">
                <Briefcase className="h-5 w-5 text-[#ff5500]" />
                <h3 className="text-base font-black text-slate-900">
                  {editingPayrollId ? 'Edit Processed Pay Slip' : 'Sri Lanka Compliant Pay Slip Generator'}
                </h3>
              </div>
              <button
                type="button"
                onClick={() => {
                  setProcessPayModalOpen(false);
                  setEditingPayrollId(null);
                }}
                className="text-slate-400 hover:text-slate-900 cursor-pointer p-1"
              >
                <X className="h-5 w-5" />
              </button>
            </div>

            {(() => {
              const breakdown = calculateSriLankanPayroll(payrollInputForm);
              const selectedStaff = staffList.find(s => s.id === payrollInputForm.staffId) || staffList[0];
              const isEpfActive = payrollInputForm.epfEtfEnabled !== false;

              return (
                <form
                  onSubmit={(e) => {
                    e.preventDefault();
                    if (!selectedStaff) return;

                    const now = new Date();
                    const nowIso = now.toISOString();
                    const fullDateStr = getLocalDateStr();
                    const timeStr = now.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });

                    if (editingPayrollId) {
                      // 1. UPDATE EXISTING RECORD (PRESERVE ORIGINAL ID)
                      const targetExisting = payrollRecords.find(r => r.id === editingPayrollId);
                      const updatedSlipRecord = {
                        ...(targetExisting || {}),
                        id: editingPayrollId,
                        staffId: selectedStaff.id,
                        staffName: selectedStaff.name,
                        role: selectedStaff.role,
                        period: payrollInputForm.period,
                        epfEtfEnabled: isEpfActive,
                        breakdown,
                        notes: payrollInputForm.notes || 'Salary Record Updated',
                        lastEditedBy: currentUser.name,
                        lastEditedByRole: currentUser.role,
                        lastEditedAt: `${fullDateStr} ${timeStr}`,
                        lastEditedAtISO: nowIso,
                        revisions: [
                          ...((targetExisting && targetExisting.revisions) || []),
                          {
                            editedAt: nowIso,
                            editor: currentUser.name,
                            previousNet: targetExisting ? (targetExisting.breakdown?.netSalary || targetExisting.netPay) : 0,
                            newNet: breakdown.netSalary
                          }
                        ]
                      };

                      setPayrollRecords(prev => prev.map(r => r.id === editingPayrollId ? updatedSlipRecord : r));

                      // Auto-print updated version
                      triggerAutoPrint({
                        type: 'PAYSLIP_PRINT',
                        data: updatedSlipRecord
                      }, `Updated Payslip ${updatedSlipRecord.id} - ${selectedStaff.name}`);

                      recordAuditLog(
                        'PAYROLL_RECORD_EDITED',
                        updatedSlipRecord.id,
                        `Updated ${isEpfActive ? 'EPF-Liable' : 'EPF-Exempt'} payslip for ${selectedStaff.name} (${payrollInputForm.period}). New Net: ${settings.currency} ${breakdown.netSalary.toFixed(2)}. Cloud Sync Queued.`
                      );
                    } else {
                      // 2. CREATE NEW PERMANENT RECORD
                      const permanentSlipId = `PAY-${payrollInputForm.period.replace(/-/g, '')}-${selectedStaff.id.slice(-4)}-${Date.now().toString().slice(-4)}`;

                      const permanentSlipRecord = {
                        id: permanentSlipId,
                        staffId: selectedStaff.id,
                        staffName: selectedStaff.name,
                        role: selectedStaff.role,
                        period: payrollInputForm.period,
                        epfEtfEnabled: isEpfActive,
                        breakdown,
                        notes: payrollInputForm.notes || 'Monthly Salary Disbursed',
                        processedBy: currentUser.name,
                        processedByRole: currentUser.role,
                        processedAt: `${fullDateStr} ${timeStr}`,
                        processedAtISO: nowIso,
                        revisions: []
                      };

                      setPayrollRecords(prev => [permanentSlipRecord, ...prev]);

                      // Direct cloud append if using atomic append
                      if (typeof appendCloudArchive === 'function') {
                        appendCloudArchive('payroll_records', permanentSlipRecord);
                      }

                      // Auto-print newly issued slip
                      triggerAutoPrint({
                        type: 'PAYSLIP_PRINT',
                        data: permanentSlipRecord
                      }, `Payslip ${permanentSlipRecord.id} - ${selectedStaff.name}`);

                      recordAuditLog(
                        'PAYROLL_PERMANENT_RECORD_SAVED',
                        permanentSlipRecord.id,
                        `Issued ${isEpfActive ? 'EPF-Liable' : 'EPF-Exempt'} payslip for ${selectedStaff.name} (${payrollInputForm.period}). Net: ${settings.currency} ${breakdown.netSalary.toFixed(2)}. Cloud Sync Queued.`
                      );
                    }

                    setEditingPayrollId(null);
                    setProcessPayModalOpen(false);
                  }}
                  className="mt-4 space-y-4"
                >
                  <div className="grid grid-cols-2 gap-3">
                    <div>
                      <label className="block text-xs font-bold text-slate-700 mb-1">Employee</label>
                      <select
                        value={payrollInputForm.staffId}
                        onChange={(e) => {
                          const targetId = e.target.value;
                          const staff = staffList.find(s => s.id === targetId);

                          setPayrollInputForm(prev => ({
                            ...prev,
                            staffId: targetId,
                            basicSalary: staff?.basicSalary ?? 35000,
                            budgetaryAllowance: staff?.budgetaryAllowance ?? 2500,
                            otherAllowances: staff?.otherAllowances ?? 0,
                            incentiveBonus: staff?.fixedBonus ?? 0,
                            overtimeRate: staff?.overtimeRate ?? 250,
                            epfEtfEnabled: staff?.epfEtfEnabled !== false
                          }));
                        }}
                        className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-bold text-slate-900 focus:bg-white focus:outline-none focus:border-[#ff5500]"
                      >
                        {staffList.map(s => (
                          <option key={s.id} value={s.id}>
                            {s.name} ({s.role}) - Base: {settings.currency} {s.basicSalary ?? 35000}
                          </option>
                        ))}
                      </select>
                    </div>

                    <div>
                      <label className="block text-xs font-bold text-slate-700 mb-1">Salary Month</label>
                      <input
                        type="month"
                        required
                        value={payrollInputForm.period}
                        onChange={e => setPayrollInputForm(prev => ({ ...prev, period: e.target.value }))}
                        className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-mono font-bold text-slate-900 focus:bg-white focus:outline-none"
                      />
                    </div>
                  </div>

                  {/* EPF / ETF APPLICABILITY TOGGLE */}
                  <div className="p-3 bg-slate-50 border border-slate-200 rounded-2xl flex items-center justify-between gap-3">
                    <div>
                      <span className="text-xs font-black text-slate-900 block">
                        Sri Lanka EPF / ETF Deductions
                      </span>
                      <p className="text-[10px] text-slate-500">
                        {isEpfActive 
                          ? 'Employee 8% deduction + Employer 12% EPF & 3% ETF active' 
                          : 'Exempt / Non-EPF (Casual, Probationary, or Contractor)'}
                      </p>
                    </div>

                    <button
                      type="button"
                      onClick={() => setPayrollInputForm(prev => ({ ...prev, epfEtfEnabled: !isEpfActive }))}
                      className={`px-3 py-1.5 rounded-xl text-xs font-bold transition-all cursor-pointer shrink-0 flex items-center gap-1.5 ${
                        isEpfActive
                          ? 'bg-emerald-600 hover:bg-emerald-700 text-white shadow-xs'
                          : 'bg-slate-200 hover:bg-slate-300 text-slate-700'
                      }`}
                    >
                      {isEpfActive ? (
                        <>
                          <Check className="h-3.5 w-3.5 stroke-[3]" />
                          <span>EPF Active</span>
                        </>
                      ) : (
                        <span>Exempt (No EPF)</span>
                      )}
                    </button>
                  </div>
                  {/* ATTENDANCE, WORKING DAYS & LEAVES (AUTO-SYNCED & MANUALLY ADJUSTABLE) */}
                  <div className="p-3.5 bg-slate-50 rounded-2xl border border-slate-200 space-y-3">
                    <div className="flex items-center justify-between">
                      <span className="text-[10px] font-black uppercase tracking-wider text-slate-500">
                        Attendance, Leave &amp; Shift Register
                      </span>
                      <span className="text-[10px] font-bold text-emerald-700 bg-emerald-50 px-2 py-0.5 rounded-full border border-emerald-200">
                        ⚡ Linked to Time Clock
                      </span>
                    </div>

                    <div className="grid grid-cols-3 gap-2.5">
                      <div>
                        <label className="block text-[11px] font-bold text-slate-700 mb-1">Standard Work Days</label>
                        <input
                          type="number"
                          min="1"
                          max="31"
                          value={payrollInputForm.standardWorkingDays}
                          onChange={e => setPayrollInputForm(prev => ({ ...prev, standardWorkingDays: e.target.value }))}
                          className="w-full px-2.5 py-1.5 bg-white border border-slate-200 rounded-lg text-xs font-mono font-bold text-slate-900"
                        />
                      </div>
                      <div>
                        <label className="block text-[11px] font-bold text-slate-700 mb-1">Actual Worked Days</label>
                        <input
                          type="number"
                          min="0"
                          max="31"
                          value={payrollInputForm.workedDays}
                          onChange={e => setPayrollInputForm(prev => ({ ...prev, workedDays: e.target.value }))}
                          className="w-full px-2.5 py-1.5 bg-white border border-slate-200 rounded-lg text-xs font-mono font-bold text-emerald-700"
                        />
                      </div>
                      <div>
                        <label className="block text-[11px] font-bold text-slate-700 mb-1">Poya / Public Holidays</label>
                        <input
                          type="number"
                          min="0"
                          max="10"
                          value={payrollInputForm.holidaysCount}
                          onChange={e => setPayrollInputForm(prev => ({ ...prev, holidaysCount: e.target.value }))}
                          className="w-full px-2.5 py-1.5 bg-white border border-slate-200 rounded-lg text-xs font-mono font-bold text-slate-900"
                        />
                      </div>
                    </div>

                    <div className="grid grid-cols-3 gap-2.5 pt-1 border-t border-slate-200">
                      <div>
                        <label className="block text-[10px] font-bold text-slate-600 mb-1">Allowed Paid Leaves</label>
                        <input
                          type="number"
                          min="0"
                          value={payrollInputForm.paidLeaves}
                          onChange={e => setPayrollInputForm(prev => ({ ...prev, paidLeaves: e.target.value }))}
                          className="w-full px-2 py-1.5 bg-white border border-slate-200 rounded-lg text-xs font-mono font-bold text-slate-900"
                        />
                      </div>
                      <div>
                        <label className="block text-[10px] font-bold text-rose-700 mb-1">Unpaid Leaves (Deducted)</label>
                        <input
                          type="number"
                          min="0"
                          value={payrollInputForm.unpaidLeaves}
                          onChange={e => setPayrollInputForm(prev => ({ ...prev, unpaidLeaves: e.target.value }))}
                          className="w-full px-2 py-1.5 bg-rose-50 border border-rose-300 rounded-lg text-xs font-mono font-bold text-rose-900"
                        />
                      </div>
                      <div>
                        <label className="block text-[10px] font-bold text-amber-700 mb-1">Short Shifts (&lt;5h)</label>
                        <input
                          type="number"
                          min="0"
                          value={payrollInputForm.shortShiftsCount}
                          onChange={e => setPayrollInputForm(prev => ({ ...prev, shortShiftsCount: e.target.value }))}
                          className="w-full px-2 py-1.5 bg-white border border-slate-200 rounded-lg text-xs font-mono font-bold text-amber-900"
                        />
                      </div>
                    </div>

                    {Number(breakdown.unpaidLeaveDeduction) > 0 && (
                      <div className="p-2 bg-rose-50 border border-rose-200 rounded-xl text-[11px] text-rose-800 flex justify-between font-mono">
                        <span>Unpaid Leave Deduction ({payrollInputForm.unpaidLeaves} days @ {settings.currency}{breakdown.perDayBasicRate}/day):</span>
                        <span className="font-bold">-{settings.currency} {breakdown.unpaidLeaveDeduction.toFixed(2)}</span>
                      </div>
                    )}
                  </div>

                  {/* EARNINGS & ALLOWANCES */}
                  <div className="p-3.5 bg-slate-50 rounded-2xl border border-slate-200 space-y-3">
                    <span className="text-[10px] font-black uppercase tracking-wider text-slate-400 block">
                      1. Earnings &amp; Allowances ({settings.currency})
                    </span>

                    <div className="grid grid-cols-2 gap-3">
                      <div>
                        <label className="block text-[11px] font-bold text-slate-600 mb-1">Basic Salary</label>
                        <input
                          type="number"
                          step="0.01"
                          required
                          value={payrollInputForm.basicSalary}
                          onChange={e => setPayrollInputForm(prev => ({ ...prev, basicSalary: e.target.value }))}
                          className="w-full px-2.5 py-1.5 bg-white border border-slate-200 rounded-lg text-xs font-mono font-bold"
                        />
                      </div>
                      <div>
                        <label className="block text-[11px] font-bold text-slate-600 mb-1">Budgetary Relief (BRA)</label>
                        <input
                          type="number"
                          step="0.01"
                          value={payrollInputForm.budgetaryAllowance}
                          onChange={e => setPayrollInputForm(prev => ({ ...prev, budgetaryAllowance: e.target.value }))}
                          className="w-full px-2.5 py-1.5 bg-white border border-slate-200 rounded-lg text-xs font-mono font-bold"
                        />
                      </div>
                    </div>

                    <div className="grid grid-cols-2 gap-3">
                      <div>
                        <label className="block text-[11px] font-bold text-slate-600 mb-1">
                          Fixed Allowances {isEpfActive && '(Liable to EPF)'}
                        </label>
                        <input
                          type="number"
                          step="0.01"
                          value={payrollInputForm.otherAllowances}
                          onChange={e => setPayrollInputForm(prev => ({ ...prev, otherAllowances: e.target.value }))}
                          className="w-full px-2.5 py-1.5 bg-white border border-slate-200 rounded-lg text-xs font-mono font-bold"
                        />
                      </div>
                      <div>
                        <label className="block text-[11px] font-bold text-slate-600 mb-1">Service Gratuity Share</label>
                        <input
                          type="number"
                          step="0.01"
                          value={payrollInputForm.serviceChargeBonus}
                          onChange={e => setPayrollInputForm(prev => ({ ...prev, serviceChargeBonus: e.target.value }))}
                          className="w-full px-2.5 py-1.5 bg-white border border-slate-200 rounded-lg text-xs font-mono font-bold text-emerald-700"
                        />
                      </div>
                    </div>
                  </div>

                  {/* OVERTIME, ADVANCE & DEDUCTIONS */}
                  <div className="p-3.5 bg-slate-50 rounded-2xl border border-slate-200 space-y-3">
                    <span className="text-[10px] font-black uppercase tracking-wider text-slate-400 block">
                      2. Additions, Overtime &amp; Advances ({settings.currency})
                    </span>

                    <div className="grid grid-cols-2 gap-3">
                      <div>
                        <label className="block text-[11px] font-bold text-slate-600 mb-1">Overtime Hours (hrs)</label>
                        <input
                          type="number"
                          step="0.5"
                          min="0"
                          value={payrollInputForm.overtimeHours}
                          onChange={e => setPayrollInputForm(prev => ({ ...prev, overtimeHours: e.target.value }))}
                          className="w-full px-2.5 py-1.5 bg-white border border-slate-200 rounded-lg text-xs font-mono font-bold"
                        />
                      </div>
                      <div>
                        <label className="block text-[11px] font-bold text-slate-600 mb-1">Performance Bonus</label>
                        <input
                          type="number"
                          step="0.01"
                          min="0"
                          value={payrollInputForm.incentiveBonus}
                          onChange={e => setPayrollInputForm(prev => ({ ...prev, incentiveBonus: e.target.value }))}
                          className="w-full px-2.5 py-1.5 bg-white border border-slate-200 rounded-lg text-xs font-mono font-bold"
                        />
                      </div>
                    </div>

                    <div className="grid grid-cols-2 gap-3 pt-1 border-t border-slate-200">
                      {/* SALARY ADVANCE INPUT */}
                      <div>
                        <div className="flex items-center justify-between mb-1">
                          <label className="block text-[11px] font-black text-amber-800">
                            Salary Advance ({settings.currency})
                          </label>
                          {Number(payrollInputForm.salaryAdvance) > 0 && (
                            <button
                              type="button"
                              onClick={() => {
                                triggerAutoPrint({
                                  type: 'SALARY_ADVANCE_VOUCHER',
                                  data: {
                                    id: editingPayrollId || `ADV-${payrollInputForm.period.replace(/-/g, '')}-${selectedStaff?.id.slice(-4)}`,
                                    staffName: selectedStaff?.name || 'Staff',
                                    role: selectedStaff?.role || 'Staff',
                                    period: payrollInputForm.period,
                                    amount: Number(payrollInputForm.salaryAdvance) || 0,
                                    notes: payrollInputForm.notes || 'Salary Advance Payment'
                                  }
                                }, `Advance Voucher: ${selectedStaff?.name} (${settings.currency} ${payrollInputForm.salaryAdvance})`);
                              }}
                              className="text-[9px] font-bold text-amber-900 bg-amber-100 hover:bg-amber-200 px-1.5 py-0.5 rounded cursor-pointer transition-colors"
                              title="Print advance receipt voucher now"
                            >
                              🖨️ Print Slip
                            </button>
                          )}
                        </div>
                        <input
                          type="number"
                          step="0.01"
                          min="0"
                          value={payrollInputForm.salaryAdvance}
                          onChange={e => setPayrollInputForm(prev => ({ ...prev, salaryAdvance: e.target.value }))}
                          placeholder="e.g. 10000.00"
                          className="w-full px-2.5 py-1.5 bg-amber-50/70 border border-amber-300 rounded-lg text-xs font-mono font-bold text-amber-950 focus:bg-white focus:outline-none focus:border-amber-600"
                        />
                      </div>

                      <div>
                        <label className="block text-[11px] font-bold text-slate-600 mb-1">Other Deductions</label>
                        <input
                          type="number"
                          step="0.01"
                          min="0"
                          value={payrollInputForm.otherDeductions}
                          onChange={e => setPayrollInputForm(prev => ({ ...prev, otherDeductions: e.target.value }))}
                          className="w-full px-2.5 py-1.5 bg-white border border-slate-200 rounded-lg text-xs font-mono font-bold text-rose-700"
                        />
                      </div>
                    </div>
                  </div>

                  {/* LIVE BREAKDOWN PREVIEW */}
                  <div className="p-3.5 bg-slate-900 text-white rounded-2xl space-y-1.5 text-xs font-mono">
                    <div className="flex justify-between text-slate-400">
                      <span>Gross Earnings:</span>
                      <span>{settings.currency} {breakdown.grossEarnings.toFixed(2)}</span>
                    </div>

                    {isEpfActive ? (
                      <>
                        <div className="flex justify-between text-slate-400 text-[11px]">
                          <span>EPF Liable Base:</span>
                          <span>{settings.currency} {breakdown.epfLiableEarnings.toFixed(2)}</span>
                        </div>
                        <div className="flex justify-between text-rose-400">
                          <span>EPF Employee (8%):</span>
                          <span>-{settings.currency} {breakdown.epfEmployee.toFixed(2)}</span>
                        </div>
                      </>
                    ) : (
                      <div className="py-1 text-[11px] text-amber-400/90 italic flex items-center justify-between">
                        <span>Statutory EPF / ETF:</span>
                        <span>Exempt (0.00)</span>
                      </div>
                    )}

                    {Number(breakdown.salaryAdvance) > 0 && (
                      <div className="flex justify-between text-amber-300 font-bold">
                        <span>Salary Advance Deducted:</span>
                        <span>-{settings.currency} {breakdown.salaryAdvance.toFixed(2)}</span>
                      </div>
                    )}

                    {Number(payrollInputForm.otherDeductions) > 0 && (
                      <div className="flex justify-between text-rose-400">
                        <span>Other Deductions:</span>
                        <span>-{settings.currency} {Number(payrollInputForm.otherDeductions).toFixed(2)}</span>
                      </div>
                    )}

                    <div className="flex justify-between pt-2 border-t border-slate-700 text-base font-black text-emerald-400">
                      <span>Net Employee Take-Home:</span>
                      <span>{settings.currency} {breakdown.netSalary.toFixed(2)}</span>
                    </div>
                  </div>

                  {/* MODAL ACTION BUTTONS (INCLUDES DIRECT PRINT ADVANCE VOUCHER) */}
                  <div className="flex flex-col sm:flex-row gap-2 pt-1">
                    <button
                      type="button"
                      onClick={() => {
                        setProcessPayModalOpen(false);
                        setEditingPayrollId(null);
                      }}
                      className="py-2.5 px-3 bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold rounded-xl text-xs cursor-pointer transition-colors"
                    >
                      Cancel
                    </button>

                    {/* PRINT SALARY ADVANCE VOUCHER BUTTON */}
                    {Number(payrollInputForm.salaryAdvance) > 0 && (
                      <button
                        type="button"
                        onClick={() => {
                          triggerAutoPrint({
                            type: 'SALARY_ADVANCE_VOUCHER',
                            data: {
                              id: editingPayrollId || `ADV-${payrollInputForm.period.replace(/-/g, '')}-${selectedStaff?.id.slice(-4)}`,
                              staffName: selectedStaff?.name || 'Staff',
                              role: selectedStaff?.role || 'Staff',
                              period: payrollInputForm.period,
                              amount: Number(payrollInputForm.salaryAdvance) || 0,
                              notes: payrollInputForm.notes || 'Salary Advance Payment'
                            }
                          }, `Advance Voucher: ${selectedStaff?.name}`);
                        }}
                        className="py-2.5 px-3 bg-amber-500 hover:bg-amber-600 text-white font-bold rounded-xl text-xs flex items-center justify-center gap-1.5 shadow-xs cursor-pointer transition-colors"
                      >
                        <Printer className="h-3.5 w-3.5" />
                        <span>Print Advance Slip</span>
                      </button>
                    )}

                    <button
                      type="submit"
                      className="flex-1 py-2.5 bg-[#ff5500] hover:bg-orange-600 text-white font-bold rounded-xl text-xs shadow-xs cursor-pointer transition-colors flex items-center justify-center gap-1.5"
                    >
                      <Printer className="h-3.5 w-3.5" />
                      <span>{editingPayrollId ? 'Save Changes & Print Payslip' : 'Issue & Print Payslip'}</span>
                    </button>
                  </div>
                </form>
              );
            })()}
          </div>
        </div>
      )}
      
      {/* MODAL: DISBURSE STANDALONE SALARY ADVANCE */}
      {issueAdvanceModalOpen && (
        <div
          className="fixed inset-0 bg-black/80 backdrop-blur-xs z-[9999] flex items-center justify-center p-4 text-slate-900"
          onClick={(e) => {
            if (e.target === e.currentTarget) setIssueAdvanceModalOpen(false);
          }}
        >
          <div className="bg-white rounded-3xl w-full max-w-md p-6 shadow-2xl border border-slate-200 relative max-h-[90vh] overflow-y-auto space-y-4">
            <div className="flex items-center justify-between pb-3 border-b border-slate-100">
              <div className="flex items-center gap-2">
                <Banknote className="h-5 w-5 text-amber-600" />
                <div>
                  <h3 className="text-base font-black text-slate-900">Issue Salary Advance Payment</h3>
                  <p className="text-[10px] text-slate-500">Disburses advance funds and prints thermal receipt voucher</p>
                </div>
              </div>
              <button
                type="button"
                onClick={() => setIssueAdvanceModalOpen(false)}
                className="text-slate-400 hover:text-slate-900 cursor-pointer p-1"
              >
                <X className="h-5 w-5" />
              </button>
            </div>

            <form
              onSubmit={(e) => {
                e.preventDefault();
                const amt = parseFloat(advanceForm.amount);
                const selectedEmployee = staffList.find(s => s.id === advanceForm.staffId) || staffList[0];

                if (!selectedEmployee || isNaN(amt) || amt <= 0) {
                  alert('Please select a valid employee and enter a positive advance amount.');
                  return;
                }

                const advanceId = `ADV-${advanceForm.period.replace(/-/g, '')}-${selectedEmployee.id.slice(-4)}-${Date.now().toString().slice(-4)}`;

                const newAdvanceRecord = {
                  id: advanceId,
                  staffId: selectedEmployee.id,
                  staffName: selectedEmployee.name,
                  role: selectedEmployee.role,
                  period: advanceForm.period,
                  amount: amt,
                  paymentMethod: advanceForm.paymentMethod,
                  reason: advanceForm.reason.trim() || 'Salary Advance',
                  notes: advanceForm.notes.trim(),
                  disbursedBy: currentUser.name,
                  disbursedAt: new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }),
                  date: getLocalDateStr(),
                  timestamp: new Date().toISOString(),
                  status: 'APPROVED'
                };

                // 1. Add to state
                setSalaryAdvances(prev => [newAdvanceRecord, ...prev]);

                // 2. Also register as a drawer cash out payout if paid via CASH
                if (advanceForm.paymentMethod === 'CASH') {
                  const cashOutEntry = {
                    id: getNextCashOutNumber(),
                    shiftId: currentShift.shiftId,
                    amount: amt,
                    category: 'Staff Advance',
                    reason: `Salary advance for ${selectedEmployee.name} (${advanceForm.period})`,
                    recipient: selectedEmployee.name,
                    requestedBy: currentUser.name,
                    requestedRole: currentUser.role,
                    createdAt: new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }),
                    date: getLocalDateStr(),
                    status: 'APPROVED',
                    approvedBy: currentUser.name,
                    approvedAt: new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })
                  };

                  setCurrentShift(prev => ({
                    ...prev,
                    payouts: [cashOutEntry, ...(prev.payouts || [])]
                  }));
                  setExpenses(prev => [cashOutEntry, ...prev]);
                }

                // 3. Print the thermal advance voucher
                triggerAutoPrint({
                  type: 'SALARY_ADVANCE_VOUCHER',
                  data: {
                    id: newAdvanceRecord.id,
                    staffName: newAdvanceRecord.staffName,
                    role: newAdvanceRecord.role,
                    period: newAdvanceRecord.period,
                    amount: newAdvanceRecord.amount,
                    notes: newAdvanceRecord.notes || newAdvanceRecord.reason
                  }
                }, `Advance Disbursed: ${selectedEmployee.name} (${settings.currency} ${amt.toFixed(2)})`);

                // 4. Record audit trail
                recordAuditLog(
                  'SALARY_ADVANCE_ISSUED',
                  newAdvanceRecord.id,
                  `Issued ${settings.currency} ${amt.toFixed(2)} salary advance (${newAdvanceRecord.paymentMethod}) to ${selectedEmployee.name} for period ${advanceForm.period}`
                );

                setIssueAdvanceModalOpen(false);
              }}
              className="space-y-3.5"
            >
              <div>
                <label className="block text-xs font-bold text-slate-700 mb-1">Select Employee *</label>
                <select
                  value={advanceForm.staffId}
                  onChange={e => setAdvanceForm(prev => ({ ...prev, staffId: e.target.value }))}
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-bold text-slate-900 focus:bg-white focus:outline-none focus:border-[#ff5500]"
                >
                  {staffList.map(s => (
                    <option key={s.id} value={s.id}>
                      {s.name} ({s.role}) - Base: {settings.currency} {s.basicSalary ?? 35000}
                    </option>
                  ))}
                </select>
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">Advance Amount ({settings.currency}) *</label>
                  <input
                    type="number"
                    step="0.01"
                    min="1"
                    required
                    value={advanceForm.amount}
                    onChange={e => setAdvanceForm(prev => ({ ...prev, amount: e.target.value }))}
                    placeholder="e.g. 10000.00"
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-mono font-bold text-slate-900 focus:bg-white focus:outline-none focus:border-[#ff5500]"
                  />
                </div>

                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">Deduct On Month *</label>
                  <input
                    type="month"
                    required
                    value={advanceForm.period}
                    onChange={e => setAdvanceForm(prev => ({ ...prev, period: e.target.value }))}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-mono font-bold text-slate-900 focus:bg-white focus:outline-none focus:border-[#ff5500]"
                  />
                </div>
              </div>

              <div>
                <label className="block text-xs font-bold text-slate-700 mb-1">Payment Method</label>
                <select
                  value={advanceForm.paymentMethod}
                  onChange={e => setAdvanceForm(prev => ({ ...prev, paymentMethod: e.target.value }))}
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-bold text-slate-900 focus:bg-white focus:outline-none focus:border-[#ff5500]"
                >
                  <option value="CASH">Cash Drawer Outflow (Deducts Float &amp; Pops Drawer)</option>
                  <option value="BANK_TRANSFER">Direct Bank Transfer / Online</option>
                </select>
              </div>

              <div>
                <label className="block text-xs font-bold text-slate-700 mb-1">Reason / Purpose</label>
                <input
                  type="text"
                  value={advanceForm.reason}
                  onChange={e => setAdvanceForm(prev => ({ ...prev, reason: e.target.value }))}
                  placeholder="e.g. Medical emergency / travel advance"
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs text-slate-900 focus:bg-white focus:outline-none focus:border-[#ff5500]"
                />
              </div>

              <div className="flex gap-2 pt-2 border-t border-slate-100">
                <button
                  type="button"
                  onClick={() => setIssueAdvanceModalOpen(false)}
                  className="flex-1 py-2.5 bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold rounded-xl text-xs cursor-pointer transition-colors"
                >
                  Cancel
                </button>
                <button
                  type="submit"
                  className="flex-1 py-2.5 bg-amber-500 hover:bg-amber-600 text-white font-bold rounded-xl text-xs shadow-xs cursor-pointer transition-colors flex items-center justify-center gap-1.5"
                >
                  <Printer className="h-3.5 w-3.5" />
                  <span>Disburse &amp; Print Slip</span>
                </button>
              </div>
            </form>
          </div>
        </div>
      )}
    </div>
  );
}
// End of component