import React, { useState, useEffect } from 'react';
import axios from 'axios';
import * as XLSX from 'xlsx';

// Dynamic Cloud / Local Backend URL Configuration
const BACKEND_URL = import.meta.env.VITE_API_BASE_URL || 'https://msedcl-billing-backend.onrender.com';

// Master Units List
const UNITS_LIST = ['NO', 'SET', 'KG', 'M', 'LS', 'EA', 'SQ.M', 'LTR'];

// Helper Function: Convert Number to Indian Currency Words
const convertNumberToIndianWords = (amount) => {
  if (isNaN(amount) || amount === null || amount === undefined) return '';
  const num = Math.round(Number(amount) * 100) / 100;
  if (num === 0) return 'INR Zero Rupees Only';

  const singleDigits = ['', 'One', 'Two', 'Three', 'Four', 'Five', 'Six', 'Seven', 'Eight', 'Nine'];
  const twoDigits = ['Ten', 'Eleven', 'Twelve', 'Thirteen', 'Fourteen', 'Fifteen', 'Sixteen', 'Seventeen', 'Eighteen', 'Nineteen'];
  const tensMultiple = ['', '', 'Twenty', 'Thirty', 'Forty', 'Fifty', 'Sixty', 'Seventy', 'Eighty', 'Ninety'];

  const convertTwoDigits = (n) => {
    if (n < 10) return singleDigits[n];
    if (n >= 10 && n < 20) return twoDigits[n - 10];
    return tensMultiple[Math.floor(n / 10)] + (n % 10 !== 0 ? ' ' + singleDigits[n % 10] : '');
  };

  const convertThreeDigits = (n) => {
    let str = '';
    if (n >= 100) {
      str += singleDigits[Math.floor(n / 100)] + ' Hundred ';
      n %= 100;
    }
    if (n > 0) {
      str += convertTwoDigits(n);
    }
    return str.trim();
  };

  const rupees = Math.floor(num);
  const paise = Math.round((num - rupees) * 100);

  let result = '';

  const crore = Math.floor(rupees / 10000000);
  let rem = rupees % 10000000;
  const lakh = Math.floor(rem / 100000);
  rem %= 100000;
  const thousand = Math.floor(rem / 1000);
  const hundreds = rem % 1000;

  if (crore > 0) result += convertTwoDigits(crore) + ' Crore ';
  if (lakh > 0) result += convertTwoDigits(lakh) + ' Lakh ';
  if (thousand > 0) result += convertTwoDigits(thousand) + ' Thousand ';
  if (hundreds > 0) result += convertThreeDigits(hundreds);

  result = result.trim();
  if (!result) result = 'Zero';

  let finalWords = `INR ${result} Rupees`;
  if (paise > 0) {
    finalWords += ` and ${convertTwoDigits(paise)} Paise`;
  }
  return finalWords + ' Only';
};

// Standard Parts Catalog Master (Rudra Electricals)
const RUDRA_PARTS_CATALOG = [
  { id: 1, code: '16402030015', name: '11KV 45 KN DISC INSULATORS', unit: 'NO', rate: 295, category: 'Material MIGO' },
  { id: 2, code: '21020501025', name: 'HT STAY SETS', unit: 'SET', rate: 807, category: 'Material MIGO' },
  { id: 3, code: '49055171005', name: 'G.I. STAY WIRE 7/3.15 MM (10 SWG)', unit: 'KG', rate: 72, category: 'Material MIGO' },
  { id: 4, code: '86999817629', name: 'CRIMPING OF JUMPERS AT CUT POINTS', unit: 'NO', rate: 1696, category: 'Material MIGO' },
  { id: 5, code: '86999821499', name: 'SUNDRIES SUCH AS N/B, WASHERS, WELDING ETC.', unit: 'LS', rate: 3163, category: 'Material MIGO' },
  { id: 6, code: '16251017075', name: '11 KV PIN INSULATORS WITH G.I. PINS', unit: 'NO', rate: 120, category: 'Material MIGO' },
  { id: 7, code: '16402040085', name: '11 KV DISC INSULATORS.N.70 KN', unit: 'NO', rate: 326, category: 'Material MIGO' },
  { id: 8, code: '21808563185', name: '11KV TOP FITTING', unit: 'NO', rate: 123, category: 'Material MIGO' },
  { id: 9, code: '33286328101', name: 'PSC POLE 9 MTR / 200KG', unit: 'NO', rate: 2983, category: 'Material MIGO' },
  { id: 10, code: '33401009041', name: 'M.S.FLAT50X6MM', unit: 'KG', rate: 65, category: 'Material MIGO' },
  { id: 11, code: '21807299615', name: '11 KV V CROSS ARM WITH CLAMP', unit: 'NO', rate: 863, category: 'Material MIGO' },
  { id: 12, code: '11862015015', name: 'AAAC 34 SQMM COND', unit: 'M', rate: 17.8, category: 'Material MIGO' },
  { id: 13, code: '11862035065', name: 'AAAC 100 SQMM COND', unit: 'M', rate: 70, category: 'Material MIGO' },
  { id: 14, code: '16403050095', name: '11 KV DISCINSUT&C 4500 KG.', unit: 'NO', rate: 281, category: 'Material MIGO' },
  { id: 15, code: '86999821069', name: 'SUNDRIES', unit: 'LS', rate: 886, category: 'Material MIGO' },
  { id: 16, code: '86999822549', name: 'EARTHING MATERIALS', unit: 'LS', rate: 500, category: 'Material MIGO' },
  { id: 17, code: '21998199035', name: 'STRAIN HARDWARE FOR 34 SQ.MM AAAC', unit: 'NO', rate: 161, category: 'Material MIGO' },
  { id: 18, code: '11062015025', name: 'ALL ALUMINIUM ALLOY CONDUCTOR OF 55 SQ.MM.', unit: 'M', rate: 36, category: 'Material MIGO' },
  { id: 19, code: '49055170885', name: 'G.I.STAY WIRE7/4 MM (8 SWG)', unit: 'KG', rate: 72, category: 'Material MIGO' },
  { id: 20, code: '49105100835', name: 'G.I.WIRE 4 MM(8 SWG)', unit: 'KG', rate: 72, category: 'Material MIGO' },
  { id: 21, code: 'BM.HT.OH.004', name: 'Repl of faulty pin insulator', unit: 'NO', rate: 183, category: 'Services SES' },
  { id: 22, code: 'BM.HT.OH.002', name: 'Repl of faulty jumper', unit: 'NO', rate: 183, category: 'Services SES' },
  { id: 23, code: 'BM.HT11KV.OH.008', name: 'Repl of tilted / damaged Cross Arm/Channel', unit: 'NO', rate: 366, category: 'Services SES' },
  { id: 24, code: 'BM.HT11KV.OH.001', name: 'Repl / Repair of snapped conductor', unit: 'SET', rate: 366, category: 'Services SES' },
  { id: 25, code: 'BM.HT11KV.OH.009', name: 'Repl of tilted / damaged Top clip', unit: 'NO', rate: 122, category: 'Services SES' },
  { id: 26, code: 'BM.HT11KV.OH.003', name: 'Repl of faulty disc insulator', unit: 'NO', rate: 488, category: 'Services SES' },
  { id: 27, code: 'BM.HT11KV.OH.007', name: 'Trimming Tree Branches fallen HT lines', unit: 'LS', rate: 610, category: 'Services SES' },
  { id: 28, code: 'BM.HT11KV.OH.006', name: 'Repl PSC poleincl diging & concreting', unit: 'NO', rate: 1220, category: 'Services SES' },
  { id: 29, code: 'BM.HT.OH.023', name: 'Concreting & muffing of PSC poles 1:4:8', unit: 'EA', rate: 1464, category: 'Services SES' },
  { id: 30, code: 'CENTAGE-PM', name: 'PM Centages', unit: 'LS', rate: 1222, category: 'Centages SES' }
];

const INITIAL_AGENCIES = [
  {
    id: 1,
    agencyName: 'M/S. RUDRA ELECTRICALS',
    address: 'At Shelgaon Tq. Badnapur, Dist. Jalna-431203',
    email: 'nileshsonwane40@gmail.com',
    mobile: '9421103111/ 9421103777',
    panNo: 'KTWPS1687C',
    gstNo: '27KTWPS1687C1ZR',
    division: 'Jalna I',
    subDivision: 'Badnapur S/Dn.',
    circleOffice: 'Jalna',
    loaNo: 'EE/R&M/BM/HT/LOE No. 139856',
    loaDate: '2026-05-13',
    workOrderNo: 'R&M/BM/HT/WO/ 4780459015,4780459017 4780461833',
    workOrderDates: '08.09.2026 & 24.09.2026',
    schemeName: 'Empanelment HT Line',
    subject: 'Submission of bill against tender for Breakdown Maintenance of HT Over Head Line & Underground Cable under Badnapur Sub Division under Jalna-I Division on Full Turnkey Basis.',
    justification: 'FULL TRUNKY ON EE.P.O ORDER JALNA',
    catalog: RUDRA_PARTS_CATALOG
  },
  {
    id: 2,
    agencyName: 'M/S. AYUSHI AIR EXPRESS PVT. LTD.',
    address: 'RJ Public School, Plot No. 484, Street No. 15, Mahipalpur Ext., New Delhi-110037',
    email: 'info@ayushiairexpress.com',
    mobile: '9312390182',
    panNo: 'AAJCA9764F',
    gstNo: '07AAJCA9764F1ZC',
    division: 'Pune / Jalna',
    subDivision: 'Logistics Division',
    circleOffice: 'Jalna',
    loaNo: 'LR.NO. 63276',
    loaDate: '2026-10-14',
    workOrderNo: 'INV/4138',
    workOrderDates: '14.10.2026',
    schemeName: 'Transport & Logistics Consignment',
    subject: 'Transportation of Breakdown Maintenance Material Consignment.',
    justification: 'DOOR DELIVERY AS PER CONSIGNMENT NOTE',
    catalog: []
  }
];

export default function App() {
  const [agencies, setAgencies] = useState(INITIAL_AGENCIES);
  const [selectedAgencyId, setSelectedAgencyId] = useState(INITIAL_AGENCIES[0].id);
  const [activeTab, setActiveTab] = useState('billing');
  const [printCopyMode, setPrintCopyMode] = useState('ALL');
  const [targetDivisionName, setTargetDivisionName] = useState('Division No.I');

  const currentAgency = agencies.find((a) => a.id === Number(selectedAgencyId)) || agencies[0];

  const [activeAgencyData, setActiveAgencyData] = useState(() => {
    const saved = localStorage.getItem('draft_agency_data');
    if (saved) {
      try { return JSON.parse(saved); } catch (e) {}
    }
    return {
      agencyName: currentAgency?.agencyName || '',
      gstNo: currentAgency?.gstNo || '',
      panNo: currentAgency?.panNo || '',
      address: currentAgency?.address || '',
      email: currentAgency?.email || '',
      mobile: currentAgency?.mobile || '',
      division: currentAgency?.division || 'Jalna I',
      subDivision: currentAgency?.subDivision || '',
      circleOffice: currentAgency?.circleOffice || 'Jalna',
      schemeName: currentAgency?.schemeName || '',
      subject: currentAgency?.subject || '',
      justification: currentAgency?.justification || ''
    };
  });

  const [gstLoading, setGstLoading] = useState(false);
  const [gstStatusMsg, setGstStatusMsg] = useState('');
  const [addAgencyLoading, setAddAgencyLoading] = useState(false);

  const [invoiceHeader, setInvoiceHeader] = useState(() => {
    const saved = localStorage.getItem('draft_invoice_header');
    if (saved) {
      try { return JSON.parse(saved); } catch (e) {}
    }
    return {
      invoiceNo: 'RE/2026-27/01',
      invoiceDate: '2026-10-23',
      sch33kvAmount: 0.0,
      belowRatePercent: 0.0,
      loaNo: currentAgency?.loaNo || '',
      loaDate: currentAgency?.loaDate || '',
      workOrderNo: currentAgency?.workOrderNo || '',
      workOrderDates: currentAgency?.workOrderDates || '',
      empanelmentNo: '139856',
      empanelmentDate: '2026-05-13',
      workStartedDate: '2026-09-08',
      workCompletedDate: '2026-09-24',
      commissioningDate: '2026-09-24'
    };
  });

  const [measurementItems, setMeasurementItems] = useState(() => {
    const saved = localStorage.getItem('draft_measurement_items');
    if (saved) {
      try { return JSON.parse(saved); } catch (e) {}
    }
    return [
      { erpPoId: '4780459015', erpMoId: '400003601315', locationName: '013405|203-AMBADGAON-AG', itemCategory: 'Material MIGO', itemCode: '16402030015', description: '11KV 45 KN DISC INSULATORS', unit: 'NO', qty: 6, rate: 295, amount: 1770 },
      { erpPoId: '4780459015', erpMoId: '400003601315', locationName: '013405|203-AMBADGAON-AG', itemCategory: 'Material MIGO', itemCode: '21020501025', description: 'HT STAY SETS', unit: 'SET', qty: 5, rate: 807, amount: 4035 },
      { erpPoId: '4780459015', erpMoId: '400003601315', locationName: '013405|203-AMBADGAON-AG', itemCategory: 'Services SES', itemCode: 'BM.HT.OH.004', description: 'Repl of faulty pin insulator', unit: 'NO', qty: 8, rate: 183, amount: 1464 },
      { erpPoId: '4780459015', erpMoId: '400003601315', locationName: '013405|203-AMBADGAON-AG', itemCategory: 'Centages SES', itemCode: 'CENTAGE-PM', description: 'PM Centages', unit: 'LS', qty: 1, rate: 1222, amount: 1222 }
    ];
  });

  useEffect(() => {
    localStorage.setItem('draft_agency_data', JSON.stringify(activeAgencyData));
  }, [activeAgencyData]);

  useEffect(() => {
    localStorage.setItem('draft_invoice_header', JSON.stringify(invoiceHeader));
  }, [invoiceHeader]);

  useEffect(() => {
    localStorage.setItem('draft_measurement_items', JSON.stringify(measurementItems));
  }, [measurementItems]);

  const [savedInvoicesList, setSavedInvoicesList] = useState([]);
  const [historySearch, setHistorySearch] = useState('');
  
  // Date-Range Filter State
  const [filterStartDate, setFilterStartDate] = useState('');
  const [filterEndDate, setFilterEndDate] = useState('');
  
  const [loadingHistory, setLoadingHistory] = useState(false);

  const [newAgencyForm, setNewAgencyForm] = useState({
    agencyName: '',
    address: '',
    email: '',
    mobile: '',
    panNo: '',
    gstNo: '',
    division: 'Jalna I',
    subDivision: 'Badnapur S/Dn.',
    circleOffice: 'Jalna',
    loaNo: '',
    loaDate: '',
    workOrderNo: '',
    workOrderDates: '',
    schemeName: 'HT Breakdown Maintenance Work',
    subject: 'Submission of bill against tender for Breakdown Maintenance on Full Turnkey Basis.',
    justification: 'FULL TRUNKY ON EE.P.O ORDER JALNA'
  });

  const [newPartForm, setNewPartForm] = useState({
    code: '',
    name: '',
    unit: 'NO',
    rate: '',
    category: 'Material MIGO'
  });

  const [saving, setSaving] = useState(false);
  const [saveSuccess, setSaveSuccess] = useState(null);

  const generateNextInvoiceNo = (currentNo) => {
    if (!currentNo) return 'RE/2026-27/01';
    const parts = currentNo.split('/');
    if (parts.length > 0) {
      const lastPart = parts[parts.length - 1];
      const num = parseInt(lastPart, 10);
      if (!isNaN(num)) {
        const nextNum = String(num + 1).padStart(lastPart.length, '0');
        parts[parts.length - 1] = nextNum;
        return parts.join('/');
      }
    }
    return `${currentNo}-1`;
  };

  const calculatedSch11kv = (measurementItems || []).reduce((acc, c) => acc + (Number(c?.amount) || 0), 0);
  const totalBreakdown = Number(invoiceHeader?.sch33kvAmount || 0) + calculatedSch11kv;
  const belowRateDeduction = (totalBreakdown * Number(invoiceHeader?.belowRatePercent || 0)) / 100;
  const subTotal = totalBreakdown - belowRateDeduction;
  const cgst = Math.round(subTotal * 0.09 * 100) / 100;
  const sgst = Math.round(subTotal * 0.09 * 100) / 100;
  const grandTotal = Math.round((subTotal + cgst + sgst) * 100) / 100;

  const totalInvoicesCount = (savedInvoicesList || []).length;
  const totalGrossTurnover = (savedInvoicesList || []).reduce((acc, inv) => acc + (Number(inv?.grandTotal) || 0), 0);
  const totalGstCollected = (savedInvoicesList || []).reduce(
    (acc, inv) => acc + (Number(inv?.cgstAmount) || 0) + (Number(inv?.sgstAmount) || 0),
    0
  );

  const fetchAllSavedInvoices = async () => {
    setLoadingHistory(true);
    try {
      const res = await axios.get(`${BACKEND_URL}/api/invoices`);
      const invoices = Array.isArray(res?.data) ? res.data : [];
      setSavedInvoicesList(invoices);

      if (invoices.length > 0 && !localStorage.getItem('draft_invoice_header')) {
        const latestInv = invoices[invoices.length - 1];
        if (latestInv && latestInv.invoiceNo) {
          const nextSuggested = generateNextInvoiceNo(latestInv.invoiceNo);
          setInvoiceHeader((prev) => ({ ...prev, invoiceNo: nextSuggested }));
        }
      }
    } catch (err) {
      console.error("Fetch Error:", err);
      setSavedInvoicesList([]);
    } finally {
      setLoadingHistory(false);
    }
  };

  useEffect(() => {
    fetchAllSavedInvoices();
  }, []);

  useEffect(() => {
    if (activeTab === 'history') {
      fetchAllSavedInvoices();
    }
  }, [activeTab]);

  const handleClearDraft = () => {
    if (!window.confirm("चालू असलेला अनसेव्हड ड्राफ्ट क्लिअर करून फॉर्म रिसेट करायचा का?")) return;
    localStorage.removeItem('draft_agency_data');
    localStorage.removeItem('draft_invoice_header');
    localStorage.removeItem('draft_measurement_items');
    window.location.reload();
  };

  const handleLoadInvoice = (inv, mode = 'view') => {
    if (!inv) return;
    setActiveAgencyData((prev) => ({
      ...prev,
      agencyName: inv.agencyName || prev.agencyName,
      gstNo: inv.gstNo || prev.gstNo,
      panNo: inv.panNo || prev.panNo,
      division: inv.division || prev.division,
      subDivision: inv.subDivision || prev.subDivision
    }));

    setInvoiceHeader((prev) => ({
      ...prev,
      invoiceNo: inv.invoiceNo || prev.invoiceNo,
      invoiceDate: inv.invoiceDate || prev.invoiceDate,
      loaNo: inv.loaNo || prev.loaNo,
      loaDate: inv.loaDate || prev.loaDate
    }));

    if (inv.division && String(inv.division).includes('II')) {
      setTargetDivisionName('Division No.II');
    } else {
      setTargetDivisionName('Division No.I');
    }

    if (inv.items && Array.isArray(inv.items) && inv.items.length > 0) {
      setMeasurementItems(
        inv.items.map((i) => ({
          erpPoId: i.erpPoId || '',
          erpMoId: i.erpMoId || '',
          locationName: i.locationName || '',
          itemCategory: i.itemCategory || 'Material MIGO',
          itemCode: i.itemCode || '',
          description: i.description || '',
          unit: i.unit || 'NO',
          qty: i.qty || 1,
          rate: i.rate || 0,
          amount: i.amount || 0
        }))
      );
    }

    if (mode === 'print') {
      setTimeout(() => {
        window.print();
      }, 300);
    } else {
      setActiveTab('billing');
      window.scrollTo({ top: 0, behavior: 'smooth' });
    }
  };

  const handleDeleteInvoice = async (id, invNo) => {
    if (!window.confirm(`खरोखर Invoice #${invNo || id} डिलीट करायचे आहे का?`)) return;
    try {
      await axios.delete(`${BACKEND_URL}/api/invoices/${id}`);
      setSaveSuccess(`Invoice #${invNo || id} यशस्वीरीत्या काढून टाकले!`);
      fetchAllSavedInvoices();
    } catch (err) {
      alert('Delete failed. Backend चालू आहे का ते तपासा.');
    }
  };

  // Safe search and Date-range filter
  const filteredHistory = (savedInvoicesList || []).filter((inv) => {
    if (!inv) return false;
    const search = String(historySearch || '').toLowerCase();
    const invNo = String(inv?.invoiceNo || '').toLowerCase();
    const agName = String(inv?.agencyName || '').toLowerCase();
    const gst = String(inv?.gstNo || '').toLowerCase();
    
    const matchesSearch = invNo.includes(search) || agName.includes(search) || gst.includes(search);

    let matchesDate = true;
    if (filterStartDate) {
      matchesDate = matchesDate && (inv?.invoiceDate >= filterStartDate);
    }
    if (filterEndDate) {
      matchesDate = matchesDate && (inv?.invoiceDate <= filterEndDate);
    }

    return matchesSearch && matchesDate;
  });

  const handleExportToExcel = () => {
    if (!filteredHistory || filteredHistory.length === 0) {
      alert('Export करण्यासाठी कोणतेही बिल उपलब्ध नाही!');
      return;
    }

    const excelData = filteredHistory.map((inv, index) => ({
      'Sr. No': index + 1,
      'Invoice ID': inv?.id || '',
      'Invoice No': inv?.invoiceNo || '',
      'Invoice Date': inv?.invoiceDate || '',
      'Agency Name': inv?.agencyName || '',
      'GSTIN': inv?.gstNo || '',
      'PAN': inv?.panNo || '',
      'Division': inv?.division || 'Jalna I',
      'Sub-Division': inv?.subDivision || '',
      'LOA No': inv?.loaNo || '',
      'Taxable Value (₹)': Number(inv?.subTotal || 0).toFixed(2),
      'CGST 9% (₹)': Number(inv?.cgstAmount || 0).toFixed(2),
      'SGST 9% (₹)': Number(inv?.sgstAmount || 0).toFixed(2),
      'Grand Total (₹)': Number(inv?.grandTotal || 0).toFixed(2)
    }));

    const worksheet = XLSX.utils.json_to_sheet(excelData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Filtered_Invoices');

    const fileName = filterStartDate && filterEndDate 
      ? `MSEDCL_Invoices_${filterStartDate}_to_${filterEndDate}.xlsx`
      : `MSEDCL_Invoice_Records_${new Date().toISOString().slice(0, 10)}.xlsx`;

    XLSX.writeFile(workbook, fileName);
  };

  const fetchOnlineGstDetails = async (gstin, target = 'billing') => {
    const isBilling = target === 'billing';
    if (isBilling) {
      setGstLoading(true);
      setGstStatusMsg('Connecting to GST Registry...');
    } else {
      setAddAgencyLoading(true);
    }

    try {
      const response = await axios.get(`https://sheetlabs.com/BSIN/gstin?gstin=${gstin}`);
      if (response?.data && response.data.length > 0) {
        const gstInfo = response.data[0];
        const tradeName = gstInfo.tradeNam || gstInfo.lgnm || '';
        const regAddr = [
          gstInfo.bno, gstInfo.bnm, gstInfo.st, gstInfo.loc, gstInfo.dst, gstInfo.stcd, gstInfo.pncd
        ].filter(Boolean).join(', ');

        const formattedName = tradeName ? `M/S. ${tradeName.toUpperCase()}` : '';
        const panVal = gstin.substring(2, 12);

        if (isBilling) {
          setActiveAgencyData((prev) => ({
            ...prev,
            agencyName: formattedName || prev.agencyName,
            address: regAddr || prev.address,
            panNo: panVal
          }));
          setGstStatusMsg(`✅ Verified: ${formattedName || 'Registered Entity'}`);
        } else {
          setNewAgencyForm((prev) => ({
            ...prev,
            agencyName: formattedName || prev.agencyName,
            address: regAddr || prev.address,
            panNo: panVal
          }));
        }
      } else {
        if (isBilling) setGstStatusMsg('ℹ️ GSTIN valid format, PAN extracted.');
      }
    } catch (err) {
      if (isBilling) setGstStatusMsg('ℹ️ PAN extracted. (Network lookup offline)');
    } finally {
      if (isBilling) setGstLoading(false);
      else setAddAgencyLoading(false);
    }
  };

  const handleGstInputChange = (e) => {
    const enteredGst = e.target.value.toUpperCase().trim();
    const panExtracted = enteredGst.length >= 10 ? enteredGst.substring(2, 12) : '';
    const matchedAgency = agencies.find((a) => a.gstNo.toUpperCase() === enteredGst);

    if (matchedAgency) {
      setSelectedAgencyId(matchedAgency.id);
      setActiveAgencyData({ ...matchedAgency });
      setGstStatusMsg(`✅ Auto-filled: ${matchedAgency.agencyName}`);
    } else {
      setActiveAgencyData((prev) => ({
        ...prev,
        gstNo: enteredGst,
        panNo: panExtracted || prev.panNo
      }));

      if (enteredGst.length === 15) {
        fetchOnlineGstDetails(enteredGst, 'billing');
      } else {
        setGstStatusMsg('');
      }
    }
  };

  const handleAddAgencyGstChange = (e) => {
    const enteredGst = e.target.value.toUpperCase().trim();
    const panExtracted = enteredGst.length >= 10 ? enteredGst.substring(2, 12) : '';

    const matchedAgency = agencies.find((a) => a.gstNo.toUpperCase() === enteredGst);
    if (matchedAgency) {
      setNewAgencyForm({ ...matchedAgency });
      return;
    }

    setNewAgencyForm((prev) => ({
      ...prev,
      gstNo: enteredGst,
      panNo: panExtracted || prev.panNo
    }));

    if (enteredGst.length === 15) {
      fetchOnlineGstDetails(enteredGst, 'addAgency');
    }
  };

  const handleSelectAgency = (aid) => {
    const id = Number(aid);
    setSelectedAgencyId(id);
    const ag = agencies.find((a) => a.id === id);
    if (ag) {
      setActiveAgencyData({ ...ag });
      setInvoiceHeader((prev) => ({
        ...prev,
        loaNo: ag.loaNo,
        loaDate: ag.loaDate,
        workOrderNo: ag.workOrderNo,
        workOrderDates: ag.workOrderDates,
        empanelmentDate: ag.loaDate
      }));
      setGstStatusMsg(`✅ Switched to: ${ag.agencyName}`);
    }
  };

  const handleActiveAgencyChange = (field, val) => {
    setActiveAgencyData((prev) => ({ ...prev, [field]: val }));
  };

  const handleAddAgencySubmit = (e) => {
    e.preventDefault();
    if (!newAgencyForm.agencyName || !newAgencyForm.gstNo) {
      alert('Agency Name & GST No required!');
      return;
    }
    const createdAgency = { ...newAgencyForm, id: Date.now(), catalog: [] };
    const updated = [...agencies, createdAgency];
    setAgencies(updated);
    setSelectedAgencyId(createdAgency.id);
    setActiveAgencyData({ ...createdAgency });
    setInvoiceHeader((prev) => ({
      ...prev,
      loaNo: createdAgency.loaNo,
      loaDate: createdAgency.loaDate,
      workOrderNo: createdAgency.workOrderNo,
      workOrderDates: createdAgency.workOrderDates
    }));
    setActiveTab('manageParts');
    alert(`Agency "${createdAgency.agencyName}" saved! You can now add parts.`);
  };

  const handleAddPartToCurrentAgency = (e) => {
    e.preventDefault();
    if (!newPartForm.code || !newPartForm.name || !newPartForm.rate) return;
    const part = {
      id: Date.now(),
      code: newPartForm.code,
      name: newPartForm.name,
      unit: newPartForm.unit,
      rate: Number(newPartForm.rate),
      category: newPartForm.category
    };

    const updated = agencies.map((a) => {
      if (a.id === currentAgency.id) {
        return { ...a, catalog: [...(a.catalog || []), part] };
      }
      return a;
    });
    setAgencies(updated);
    setNewPartForm({ code: '', name: '', unit: 'NO', rate: '', category: 'Material MIGO' });
    alert(`Part added to catalog!`);
  };

  const handleUpdateCatalogRate = (partId, newRate) => {
    const updated = agencies.map((a) => {
      if (a.id === currentAgency.id) {
        return {
          ...a,
          catalog: a.catalog.map((p) => (p.id === partId ? { ...p, rate: Number(newRate) } : p))
        };
      }
      return a;
    });
    setAgencies(updated);
  };

  const handleDeleteCatalogPart = (partId) => {
    if (!window.confirm('Delete this part from catalog?')) return;
    const updated = agencies.map((a) => {
      if (a.id === currentAgency.id) {
        return { ...a, catalog: a.catalog.filter((p) => p.id !== partId) };
      }
      return a;
    });
    setAgencies(updated);
  };

  const handleMeasurementChange = (index, field, val) => {
    const updated = [...measurementItems];
    updated[index][field] = val;
    if (field === 'qty' || field === 'rate') {
      const q = Number(field === 'qty' ? val : updated[index].qty) || 0;
      const r = Number(field === 'rate' ? val : updated[index].rate) || 0;
      updated[index].amount = Math.round(q * r * 100) / 100;
    }
    setMeasurementItems(updated);
  };

  const handleAutoSuggestPart = (index, partCode) => {
    const matched = currentAgency.catalog?.find((p) => p.code === partCode);
    if (!matched) return;
    const updated = [...measurementItems];
    updated[index].itemCode = matched.code;
    updated[index].description = matched.name;
    updated[index].unit = matched.unit;
    updated[index].rate = matched.rate;
    updated[index].itemCategory = matched.category;
    const q = Number(updated[index].qty) || 0;
    updated[index].amount = Math.round(q * matched.rate * 100) / 100;
    setMeasurementItems(updated);
  };

  const addMeasurementRow = () => {
    const last = measurementItems[measurementItems.length - 1] || {};
    setMeasurementItems([
      ...measurementItems,
      {
        erpPoId: last.erpPoId || '',
        erpMoId: last.erpMoId || '',
        locationName: last.locationName || '',
        itemCategory: 'Material MIGO',
        itemCode: '',
        description: '',
        unit: 'NO',
        qty: 1,
        rate: 0,
        amount: 0
      }
    ]);
  };

  const removeMeasurementRow = (idx) => {
    if (measurementItems.length === 1) return;
    setMeasurementItems(measurementItems.filter((_, i) => i !== idx));
  };

  const handleSaveToBackend = async () => {
    setSaving(true);
    setSaveSuccess(null);
    try {
      const payload = {
        agencyName: activeAgencyData.agencyName,
        gstNo: activeAgencyData.gstNo,
        panNo: activeAgencyData.panNo,
        invoiceNo: invoiceHeader.invoiceNo,
        invoiceDate: invoiceHeader.invoiceDate,
        division: targetDivisionName === 'Division No.II' ? 'Jalna II' : 'Jalna I',
        subDivision: activeAgencyData.subDivision,
        loaNo: invoiceHeader.loaNo,
        loaDate: invoiceHeader.loaDate,
        subTotal,
        cgstPercent: 9.0,
        sgstPercent: 9.0,
        cgstAmount: cgst,
        sgstAmount: sgst,
        grandTotal,
        items: measurementItems
      };
      const res = await axios.post(`${BACKEND_URL}/api/invoices`, payload);
      
      const savedNo = invoiceHeader.invoiceNo;
      setSaveSuccess(`✅ Invoice #${res?.data?.id || ''} (${savedNo}) successfully saved to MySQL database!`);
      
      localStorage.removeItem('draft_agency_data');
      localStorage.removeItem('draft_invoice_header');
      localStorage.removeItem('draft_measurement_items');

      const nextNo = generateNextInvoiceNo(savedNo);
      setInvoiceHeader((prev) => ({ ...prev, invoiceNo: nextNo }));

      fetchAllSavedInvoices();
    } catch (e) {
      alert('Save failed: Render Backend चालू आहे का ते तपासा (किंवा 30 सेकंद थांबून पुन्हा प्रयत्न करा).');
    } finally {
      setSaving(false);
    }
  };

  const renderDocumentSet = (copyStamp) => (
    <div key={copyStamp} className="document-set" style={{ pageBreakAfter: 'always' }}>
      {/* PAGE 1: TAX INVOICE */}
      <div className="a4-page page-1" contentEditable suppressContentEditableWarning>
        <div className="page-inner">
          <div className="page-top-content">
            <div style={{ textAlign: 'right', fontSize: '11px', fontWeight: 'bold' }}>{copyStamp}</div>
            <div style={{ textAlign: 'center', borderBottom: '2px solid #000', paddingBottom: '6px', marginBottom: '6px' }}>
              <h1 style={{ margin: 0, fontSize: '24px', letterSpacing: '0.5px' }}>{activeAgencyData.agencyName}</h1>
              <p style={{ margin: '3px 0', fontSize: '12px' }}>{activeAgencyData.address}</p>
              <p style={{ margin: '2px 0', fontSize: '11px' }}>E-mail : {activeAgencyData.email} | Mob: {activeAgencyData.mobile}</p>
            </div>
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', borderBottom: '1px solid #000', fontSize: '12px', padding: '6px', fontWeight: 'bold' }}>
              <div>PAN :- {activeAgencyData.panNo}</div>
              <div style={{ textAlign: 'right' }}>GST No. :- {activeAgencyData.gstNo}</div>
            </div>
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', borderBottom: '1px solid #000', fontSize: '11.5px' }}>
              <div style={{ borderRight: '1px solid #000', padding: '8px' }}>
                <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: '8px' }}>
                  <span>Invoice No: <strong>{invoiceHeader.invoiceNo}</strong></span>
                  <span>Date: <strong>{invoiceHeader.invoiceDate}</strong></span>
                </div>
                <p style={{ margin: '2px 0' }}><strong>To,</strong></p>
                <p style={{ margin: '2px 0' }}><strong>The Executive Engineer,</strong></p>
                <p style={{ margin: '2px 0' }}><strong>M.S.E.D.C.L. {targetDivisionName}, Jalna.</strong></p>
              </div>
              <div style={{ padding: '8px' }}>
                <p style={{ margin: '0 0 6px 0', fontWeight: 'bold' }}>LOA.No: {invoiceHeader.loaNo}, Date: {invoiceHeader.loaDate}</p>
                <p style={{ margin: 0, fontSize: '11px', lineHeight: '1.4' }}><strong>Sub - </strong> {activeAgencyData.subject}</p>
                <div style={{ display: 'grid', gridTemplateColumns: '120px 1fr', borderTop: '1px solid #000', marginTop: '8px', paddingTop: '6px' }}>
                  <div>Division:-</div>
                  <div>{targetDivisionName === 'Division No.II' ? 'Jalna II' : 'Jalna I'}</div>
                  <div>Sub-Division:-</div>
                  <div>{activeAgencyData.subDivision}</div>
                </div>
              </div>
            </div>
            <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '12px', color: '#000' }} border="1">
              <thead>
                <tr style={{ background: '#f8fafc', textAlign: 'center', fontWeight: 'bold' }}>
                  <th style={{ width: '40px', padding: '10px 6px' }}>Sr.</th>
                  <th style={{ padding: '10px 6px' }}>Particulars / Description of Work</th>
                  <th style={{ width: '120px', padding: '10px 6px' }}>Rate / %</th>
                  <th style={{ width: '130px', padding: '10px 6px', textAlign: 'right' }}>Amount (₹)</th>
                </tr>
              </thead>
              <tbody>
                <tr>
                  <td style={{ textAlign: 'center', padding: '10px 6px' }}>1</td>
                  <td style={{ padding: '10px 6px' }}>Sch H4 - 33 KV Breakdown Over Head Line</td>
                  <td style={{ textAlign: 'center', padding: '10px 6px' }}>-</td>
                  <td style={{ textAlign: 'right', padding: '10px 6px' }}>{Number(invoiceHeader.sch33kvAmount).toFixed(2)}</td>
                </tr>
                <tr>
                  <td style={{ textAlign: 'center', padding: '10px 6px' }}>2</td>
                  <td style={{ padding: '10px 6px' }}>Sch H6 - 11 KV Breakdown Over Head Line</td>
                  <td style={{ textAlign: 'center', padding: '10px 6px' }}>-</td>
                  <td style={{ textAlign: 'right', padding: '10px 6px' }}>{calculatedSch11kv.toFixed(2)}</td>
                </tr>
                <tr style={{ background: '#f8fafc', fontWeight: 'bold' }}>
                  <td style={{ textAlign: 'center', padding: '10px 6px' }}>3</td>
                  <td style={{ padding: '10px 6px' }}>Gross Total of Works</td>
                  <td style={{ textAlign: 'center', padding: '10px 6px' }}>-</td>
                  <td style={{ textAlign: 'right', padding: '10px 6px' }}>{totalBreakdown.toFixed(2)}</td>
                </tr>
                {Number(invoiceHeader.belowRatePercent) > 0 && (
                  <tr>
                    <td style={{ textAlign: 'center', padding: '10px 6px' }}>4</td>
                    <td style={{ padding: '10px 6px' }}>Less: Below Tender Rate Rebate</td>
                    <td style={{ textAlign: 'center', padding: '10px 6px' }}>{Number(invoiceHeader.belowRatePercent).toFixed(2)}%</td>
                    <td style={{ textAlign: 'right', padding: '10px 6px' }}>- {belowRateDeduction.toFixed(2)}</td>
                  </tr>
                )}
                <tr style={{ fontWeight: 'bold' }}>
                  <td colSpan="3" style={{ textAlign: 'right', padding: '10px 6px' }}>Total Taxable Value (Sub-Total):</td>
                  <td style={{ textAlign: 'right', padding: '10px 6px' }}>{subTotal.toFixed(2)}</td>
                </tr>
                <tr>
                  <td colSpan="2" style={{ textAlign: 'right', padding: '9px 6px' }}>Central GST (CGST)</td>
                  <td style={{ textAlign: 'center', padding: '9px 6px' }}>9.00%</td>
                  <td style={{ textAlign: 'right', padding: '9px 6px' }}>{cgst.toFixed(2)}</td>
                </tr>
                <tr>
                  <td colSpan="2" style={{ textAlign: 'right', padding: '9px 6px' }}>State GST (SGST)</td>
                  <td style={{ textAlign: 'center', padding: '9px 6px' }}>9.00%</td>
                  <td style={{ textAlign: 'right', padding: '9px 6px' }}>{sgst.toFixed(2)}</td>
                </tr>
                <tr style={{ fontWeight: 'bold', background: '#f1f5f9', fontSize: '13px' }}>
                  <td colSpan="3" style={{ textAlign: 'right', padding: '12px 6px' }}>Grand Total (Total Invoice Value):</td>
                  <td style={{ textAlign: 'right', padding: '12px 6px', fontSize: '14px' }}>₹{grandTotal.toFixed(2)}</td>
                </tr>
              </tbody>
            </table>
            
            {/* Amount in words */}
            <div style={{ border: '1px solid #000', borderTop: 'none', padding: '12px 14px', fontSize: '12px', background: '#fff' }}>
              <div style={{ fontWeight: 'bold', color: '#334155', textTransform: 'uppercase', fontSize: '11px' }}>Amount Chargeable (in words):</div>
              <div style={{ fontWeight: 'bold', color: '#000000', marginTop: '4px', fontSize: '13px' }}>
                {convertNumberToIndianWords(grandTotal)}
              </div>
            </div>
          </div>
          <div style={{ display: 'grid', gridTemplateColumns: '3fr 1.5fr', border: '1.5px solid #000', fontSize: '10px', minHeight: '130px', marginTop: '20px' }}>
            <div style={{ padding: '10px', borderRight: '1px solid #000', lineHeight: '1.5' }}>
              <strong>Declaration :- </strong>I/WE HEREBY CERTIFY MY/OUR REGISTRATION CERTIFICATE UNDR MAHARASHTRA VALUE ADDED TAX ACT 2005 IS IN FORCE...
            </div>
            <div style={{ textAlign: 'center', padding: '10px', display: 'flex', flexDirection: 'column', justifyContent: 'space-between' }}>
              <div style={{ fontWeight: 'bold', fontSize: '11px' }}>For {activeAgencyData.agencyName}</div>
              <div style={{ borderTop: '1px solid #000', paddingTop: '6px', fontSize: '11px', fontWeight: 'bold' }}>Authorised Signatory</div>
            </div>
          </div>
        </div>
      </div>

      {/* PAGE 2: CERTIFICATE */}
      <div className="a4-page page-2" contentEditable suppressContentEditableWarning>
        <div className="page-inner">
          <div className="page-top-content">
            <div style={{ textAlign: 'center', marginBottom: '14px' }}>
              <h2 style={{ margin: 0, fontSize: '24px', fontWeight: 'bold', textDecoration: 'underline' }}>Certificate</h2>
            </div>
            <div style={{ background: '#ffff00', color: '#000000', padding: '10px 14px', border: '1.5px solid #000', textAlign: 'center', fontWeight: 'bold', fontSize: '12.5px', marginBottom: '14px' }}>
              <div style={{ marginBottom: '4px' }}>1) LOA NO:- {invoiceHeader.loaNo} Dt.{invoiceHeader.loaDate}</div>
              <div>2) Work Order No:- {invoiceHeader.workOrderNo} Date: {invoiceHeader.workOrderDates}</div>
            </div>
            <div style={{ fontSize: '12px', fontWeight: 'bold', marginBottom: '8px' }}>This is to certify that -</div>
            <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '11px' }} border="1">
              <tbody>
                <tr><td style={{ width: '35px', textAlign: 'center', padding: '5px' }}>1</td><td colSpan="2" style={{ padding: '5px' }}>The work is completed as Per M.S.E.D.C.L standard method of Construction</td></tr>
                <tr><td style={{ textAlign: 'center', padding: '5px' }}>2</td><td colSpan="2" style={{ padding: '5px' }}>Work is Completed Within time limit as given by concern Section Officer.</td></tr>
                <tr><td style={{ textAlign: 'center', padding: '5px' }}>3</td><td colSpan="2" style={{ padding: '5px' }}>No T & P issued by MSEDCL to Contractor.</td></tr>
                <tr><td style={{ textAlign: 'center', padding: '5px' }}>4</td><td colSpan="2" style={{ padding: '5px' }}>No material is issued by MSEDCL to Contractor.</td></tr>
                <tr><td style={{ textAlign: 'center', padding: '5px' }}>5</td><td colSpan="2" style={{ padding: '5px' }}>No T & P has been provided to the contractor.</td></tr>
                <tr><td style={{ textAlign: 'center', padding: '5px' }}>6</td><td colSpan="2" style={{ padding: '5px' }}>No material balance with Contractor. Claim is correct to the best of knowledge.</td></tr>
                <tr><td style={{ textAlign: 'center', padding: '5px' }}>7</td><td colSpan="2" style={{ padding: '5px' }}>No Damages occurred to MSEDCL Properties during execution</td></tr>
                <tr><td style={{ textAlign: 'center', padding: '5px' }}>8</td><td colSpan="2" style={{ padding: '5px' }}>Work has been executed under the supervision of Section Officer & SDO</td></tr>
                <tr><td style={{ textAlign: 'center', padding: '5px' }}>9</td><td colSpan="2" style={{ padding: '5px' }}>Material Supplied by Agency is as per Standard Specifications</td></tr>
                <tr><td style={{ textAlign: 'center', padding: '5px' }}>10</td><td colSpan="2" style={{ padding: '5px' }}>Material Supplied by Agency is inspected & approved by undersigned SDO</td></tr>
                <tr><td style={{ textAlign: 'center', padding: '5px' }}>11</td><td colSpan="2" style={{ padding: '5px' }}>The measurement taken jointly by contractor & Section Officer</td></tr>
                <tr><td style={{ textAlign: 'center', padding: '5px' }}>12</td><td colSpan="2" style={{ padding: '5px' }}>Photographs uploaded on MSEDCL Website.</td></tr>
                <tr><td style={{ textAlign: 'center', padding: '5px' }}>13</td><td colSpan="2" style={{ padding: '5px' }}>Old removed material deposited in section office store</td></tr>
                <tr><td style={{ textAlign: 'center', padding: '5px' }}>14</td><td style={{ padding: '5px' }}>Work Started: <strong>{invoiceHeader.workStartedDate}</strong></td><td style={{ padding: '5px' }}>Completed: <strong>{invoiceHeader.workCompletedDate}</strong></td></tr>
                <tr><td style={{ textAlign: 'center', padding: '5px' }}>15</td><td colSpan="2" style={{ padding: '5px' }}>Variation in claimed quantities in the bill are technically justified.</td></tr>
                <tr><td style={{ textAlign: 'center', padding: '5px' }}>16</td><td colSpan="2" style={{ padding: '5px' }}>These bills are not claimed previously.</td></tr>
                <tr><td style={{ textAlign: 'center', padding: '5px' }}>17</td><td colSpan="2" style={{ padding: '5px' }}>Sites 100% verified and found work satisfactory.</td></tr>
              </tbody>
            </table>
          </div>
          <div style={{ marginTop: '50px', marginBottom: '15px', display: 'flex', justifyContent: 'space-between', fontSize: '11.5px', textAlign: 'center', fontWeight: 'bold' }}>
            <div><p>_____________________</p><p>Section Engineer</p></div>
            <div><p>_____________________</p><p>Assistant Engineer</p></div>
            <div><p>_____________________</p><p>Executive Engineer</p></div>
          </div>
        </div>
      </div>

      {/* PAGE 3: WORK COMPLETION REPORT */}
      <div className="a4-page page-3" contentEditable suppressContentEditableWarning>
        <div className="page-inner">
          <div className="page-top-content">
            <div style={{ textAlign: 'center', borderBottom: '2px solid #000', paddingBottom: '8px', marginBottom: '10px' }}>
              <h3 style={{ margin: 0, fontSize: '16px', letterSpacing: '0.5px' }}>MAHARASHTRA STATE ELECTRICTY DISTRIBUTION CO.Ltd</h3>
              <h4 style={{ margin: '4px 0 0', fontSize: '15px' }}>O&M JALNA {targetDivisionName === 'Division No.II' ? 'DIVISION NO II' : 'DIVISION NO I'} JALNA</h4>
            </div>
            <div style={{ textAlign: 'center', fontStyle: 'italic', textDecoration: 'underline', fontWeight: 'bold', fontSize: '14px', margin: '12px 0' }}>
              WORK COMPLETION REPORT
            </div>
            <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '12px' }} border="1">
              <tbody>
                <tr><td style={{ width: '45%', padding: '7px 6px', fontWeight: 'bold' }}>Circle Office :-</td><td style={{ padding: '7px 6px' }}>{activeAgencyData.circleOffice}</td></tr>
                <tr><td style={{ width: '45%', padding: '7px 6px', fontWeight: 'bold' }}>Division :-</td><td style={{ padding: '7px 6px' }}>{targetDivisionName === 'Division No.II' ? 'Division II' : 'Division I'} {activeAgencyData.circleOffice}</td></tr>
                <tr><td style={{ width: '45%', padding: '7px 6px', fontWeight: 'bold' }}>Name Of Sub Division :-</td><td style={{ padding: '7px 6px' }}>{activeAgencyData.subDivision}</td></tr>
                <tr><td style={{ padding: '7px 6px', fontWeight: 'bold' }}>1. Name Of Agency :-</td><td style={{ padding: '7px 6px' }}>{activeAgencyData.agencyName} {activeAgencyData.circleOffice}</td></tr>
                <tr><td style={{ padding: '7px 6px', fontWeight: 'bold' }}>2. LOA / Award No. {invoiceHeader.empanelmentNo} :-</td><td style={{ padding: '7px 6px' }}>Dtd {invoiceHeader.empanelmentDate}</td></tr>
                <tr><td style={{ padding: '7px 6px', fontWeight: 'bold' }}>3. Name Of Scheme :-</td><td style={{ padding: '7px 6px' }}>{activeAgencyData.schemeName}</td></tr>
                <tr><td style={{ padding: '7px 6px', fontWeight: 'bold' }}>4. Work Order No/Po No:-</td><td style={{ padding: '7px 6px' }}>{invoiceHeader.workOrderNo}</td></tr>
                <tr><td style={{ padding: '7px 6px', fontWeight: 'bold' }}>DATE OF WORK STARTED</td><td style={{ padding: '7px 6px' }}>{invoiceHeader.workStartedDate}</td></tr>
                <tr><td style={{ padding: '7px 6px', fontWeight: 'bold' }}>DATE OF COMPLETION</td><td style={{ padding: '7px 6px' }}>{invoiceHeader.workCompletedDate}</td></tr>
                <tr><td style={{ padding: '7px 6px', fontWeight: 'bold' }}>DATE OF COMMISSIONING</td><td style={{ padding: '7px 6px' }}>{invoiceHeader.commissioningDate}</td></tr>
                <tr><td colSpan="2" style={{ padding: '10px 6px', fontWeight: 'bold' }}>JUSTIFICATION FOR EXCESS/SAVING :- {activeAgencyData.justification}</td></tr>
                <tr><td colSpan="2" style={{ padding: '12px 6px' }}>Above work is completed by this Agency satisfactory and stipulated time limit by order.</td></tr>
              </tbody>
            </table>
          </div>
          <div style={{ marginTop: '50px', marginBottom: '15px', display: 'flex', justifyContent: 'space-between', fontSize: '11.5px', textAlign: 'center', fontWeight: 'bold' }}>
            <div><p>_____________________</p><p>Junior Engineer / Section Head</p></div>
            <div><p>_____________________</p><p>Deputy Executive Engineer</p></div>
            <div><p>_____________________</p><p>Executive Engineer</p></div>
          </div>
        </div>
      </div>
    </div>
  );

  return (
    <div style={{ fontFamily: 'Segoe UI, Arial, sans-serif', background: '#f8fafc', minHeight: '100vh', color: '#0f172a' }}>
      
      {/* ZERO-CONFIGURATION FULL A4 PRINT CSS */}
      <style>{`
        input, select, textarea {
          background-color: #ffffff !important;
          color: #000000 !important;
          border: 1px solid #cbd5e1 !important;
          border-radius: 6px !important;
          font-family: inherit !important;
          outline: none !important;
        }

        select option {
          background-color: #ffffff !important;
          color: #000000 !important;
          font-weight: 500 !important;
          padding: 8px !important;
        }

        input:focus, select:focus {
          border-color: #0284c7 !important;
          box-shadow: 0 0 0 2px rgba(2, 132, 199, 0.2) !important;
        }

        @media screen {
          .print-area {
            display: none !important;
          }
        }

        @page {
          size: A4 portrait;
          margin: 0mm !important;
        }

        @media print {
          html, body {
            width: 210mm !important;
            height: 297mm !important;
            margin: 0 !important;
            padding: 0 !important;
            background: #ffffff !important;
            -webkit-print-color-adjust: exact !important;
            print-color-adjust: exact !important;
          }

          .no-print, header, nav, .tips-header {
            display: none !important;
          }

          .print-area {
            display: block !important;
            width: 210mm !important;
            margin: 0 !important;
            padding: 0 !important;
          }

          .document-set {
            page-break-after: always !important;
          }

          .a4-page {
            width: 200mm !important;
            height: 287mm !important;
            max-height: 287mm !important;
            margin: 5mm auto !important;
            padding: 7mm 6mm !important;
            box-sizing: border-box !important;
            border: 2px solid #000000 !important;
            page-break-before: always !important;
            page-break-inside: avoid !important;
            page-break-after: always !important;
            overflow: hidden !important;
            background: #ffffff !important;
            color: #000000 !important;
            display: block !important;
          }

          .page-inner {
            height: 100% !important;
            display: flex !important;
            flex-direction: column !important;
            justify-content: space-between !important;
          }

          .page-1 table th, .page-1 table td {
            padding: 10px 6px !important;
            font-size: 12.5px !important;
          }

          .page-2 table td {
            padding: 6.5px 6px !important;
            font-size: 11px !important;
          }

          .page-3 table td {
            padding: 9px 8px !important;
            font-size: 12.5px !important;
          }

          table {
            page-break-inside: avoid !important;
          }
        }
      `}</style>

      {/* TOP CONTROLS & NAVIGATION */}
      <div className="no-print" style={{ maxWidth: '1320px', margin: '0 auto 20px', display: 'flex', justifyContent: 'space-between', alignItems: 'center', background: '#0f172a', padding: '16px 20px', borderRadius: '10px', color: '#fff' }}>
        <div>
          <h1 style={{ margin: 0, fontSize: '18px', fontWeight: 'bold' }}>⚡ MSEDCL Billing ERP</h1>
          <p style={{ margin: '4px 0 0', color: '#94a3b8', fontSize: '12px' }}>
            Current Agency: <strong style={{ color: '#38bdf8' }}>{activeAgencyData?.agencyName || 'No Agency Selected'}</strong> ({currentAgency?.catalog?.length || 0} Parts)
          </p>
        </div>

        <div style={{ display: 'flex', gap: '8px', alignItems: 'center' }}>
          <button onClick={() => setActiveTab('billing')} style={{ background: activeTab === 'billing' ? '#0284c7' : '#334155', color: '#fff', border: 'none', padding: '8px 12px', borderRadius: '6px', cursor: 'pointer', fontWeight: 'bold' }}>
            📋 Billing Form
          </button>
          <button onClick={() => setActiveTab('history')} style={{ background: activeTab === 'history' ? '#0284c7' : '#334155', color: '#fff', border: 'none', padding: '8px 12px', borderRadius: '6px', cursor: 'pointer', fontWeight: 'bold' }}>
            📜 Bill Records
          </button>
          <button onClick={() => setActiveTab('manageParts')} style={{ background: activeTab === 'manageParts' ? '#0284c7' : '#334155', color: '#fff', border: 'none', padding: '8px 12px', borderRadius: '6px', cursor: 'pointer', fontWeight: 'bold' }}>
            ⚙️ Parts & Rates
          </button>
          <button onClick={() => setActiveTab('manageAgency')} style={{ background: activeTab === 'manageAgency' ? '#0284c7' : '#334155', color: '#fff', border: 'none', padding: '8px 12px', borderRadius: '6px', cursor: 'pointer', fontWeight: 'bold' }}>
            🏢 + Add Agency
          </button>

          <select value={printCopyMode} onChange={(e) => setPrintCopyMode(e.target.value)} style={{ padding: '8px', borderRadius: '6px', fontWeight: 'bold', fontSize: '12px', background: '#1e293b', color: '#fff', border: '1px solid #334155' }}>
            <option value="ALL">🖨️ Both Sets (6 Pages)</option>
            <option value="ORIGINAL">📄 Original Only (3 Pages)</option>
            <option value="DUPLICATE">📄 Duplicate Only (3 Pages)</option>
          </select>

          <button onClick={handleSaveToBackend} disabled={saving} style={{ background: '#059669', color: '#fff', border: 'none', padding: '8px 14px', borderRadius: '6px', cursor: 'pointer', fontWeight: 'bold' }}>
            {saving ? '...' : '💾 Save DB'}
          </button>

          <button onClick={() => window.print()} style={{ background: '#2563eb', color: '#fff', border: 'none', padding: '8px 16px', borderRadius: '6px', cursor: 'pointer', fontWeight: 'bold' }}>
            📄 Save as PDF / Print
          </button>
        </div>
      </div>

      {saveSuccess && (
        <div className="no-print" style={{ maxWidth: '1320px', margin: '0 auto 16px', background: '#dcfce7', border: '1px solid #86efac', color: '#166534', padding: '10px 14px', borderRadius: '8px', fontWeight: 'bold' }}>
          {saveSuccess}
        </div>
      )}

      {/* DASHBOARD METRICS CARDS */}
      <div className="no-print" style={{ maxWidth: '1320px', margin: '0 auto 20px', display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(240px, 1fr))', gap: '16px' }}>
        <div style={{ background: '#fff', padding: '16px 20px', borderRadius: '10px', border: '1px solid #e2e8f0', boxShadow: '0 1px 3px rgba(0,0,0,0.05)', display: 'flex', alignItems: 'center', gap: '14px' }}>
          <div style={{ background: '#e0f2fe', color: '#0284c7', width: '48px', height: '48px', borderRadius: '10px', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: '22px' }}>
            📄
          </div>
          <div>
            <div style={{ fontSize: '12px', color: '#64748b', fontWeight: 'bold', textTransform: 'uppercase' }}>एकूण तयार बिले</div>
            <div style={{ fontSize: '22px', fontWeight: 'bold', color: '#0f172a', marginTop: '2px' }}>{totalInvoicesCount}</div>
          </div>
        </div>

        <div style={{ background: '#fff', padding: '16px 20px', borderRadius: '10px', border: '1px solid #e2e8f0', boxShadow: '0 1px 3px rgba(0,0,0,0.05)', display: 'flex', alignItems: 'center', gap: '14px' }}>
          <div style={{ background: '#dcfce7', color: '#166534', width: '48px', height: '48px', borderRadius: '10px', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: '22px' }}>
            💰
          </div>
          <div>
            <div style={{ fontSize: '12px', color: '#64748b', fontWeight: 'bold', textTransform: 'uppercase' }}>एकूण उलाढाल (Turnover)</div>
            <div style={{ fontSize: '22px', fontWeight: 'bold', color: '#166534', marginTop: '2px' }}>₹{totalGrossTurnover.toLocaleString('en-IN', { minimumFractionDigits: 2 })}</div>
          </div>
        </div>

        <div style={{ background: '#fff', padding: '16px 20px', borderRadius: '10px', border: '1px solid #e2e8f0', boxShadow: '0 1px 3px rgba(0,0,0,0.05)', display: 'flex', alignItems: 'center', gap: '14px' }}>
          <div style={{ background: '#fef3c7', color: '#b45309', width: '48px', height: '48px', borderRadius: '10px', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: '22px' }}>
            🏛️
          </div>
          <div>
            <div style={{ fontSize: '12px', color: '#64748b', fontWeight: 'bold', textTransform: 'uppercase' }}>एकूण जीएसटी (CGST+SGST)</div>
            <div style={{ fontSize: '22px', fontWeight: 'bold', color: '#b45309', marginTop: '2px' }}>₹{totalGstCollected.toLocaleString('en-IN', { minimumFractionDigits: 2 })}</div>
          </div>
        </div>
      </div>

      {/* ================= TAB: BILL RECORDS (HISTORY) ================= */}
      {activeTab === 'history' && (
        <div className="no-print" style={{ maxWidth: '1320px', margin: '0 auto 30px', background: '#fff', padding: '24px', borderRadius: '10px', border: '1px solid #e2e8f0' }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '16px', flexWrap: 'wrap', gap: '12px' }}>
            <div>
              <h2 style={{ fontSize: '16px', margin: 0, fontWeight: 'bold', color: '#0f172a' }}>
                📜 Saved Invoices in MySQL Database ({(filteredHistory || []).length} of {(savedInvoicesList || []).length})
              </h2>
              <p style={{ margin: '4px 0 0', fontSize: '12px', color: '#64748b' }}>
                तारीख श्रेणी किंवा शोध वापरून बिले फिल्टर करा आणि एक्सेल अहवाल डाऊनलोड करा.
              </p>
            </div>
            
            {/* Date-Range Filter Controls */}
            <div style={{ display: 'flex', gap: '8px', alignItems: 'center', flexWrap: 'wrap' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: '4px' }}>
                <span style={{ fontSize: '11px', fontWeight: 'bold', color: '#64748b' }}>From:</span>
                <input
                  type="date"
                  value={filterStartDate}
                  onChange={(e) => setFilterStartDate(e.target.value)}
                  style={{ padding: '6px 8px', borderRadius: '6px', fontSize: '12px' }}
                />
              </div>

              <div style={{ display: 'flex', alignItems: 'center', gap: '4px' }}>
                <span style={{ fontSize: '11px', fontWeight: 'bold', color: '#64748b' }}>To:</span>
                <input
                  type="date"
                  value={filterEndDate}
                  onChange={(e) => setFilterEndDate(e.target.value)}
                  style={{ padding: '6px 8px', borderRadius: '6px', fontSize: '12px' }}
                />
              </div>

              {(filterStartDate || filterEndDate) && (
                <button
                  onClick={() => { setFilterStartDate(''); setFilterEndDate(''); }}
                  title="Clear Date Filter"
                  style={{ background: '#fee2e2', color: '#dc2626', border: 'none', padding: '6px 10px', borderRadius: '6px', cursor: 'pointer', fontWeight: 'bold', fontSize: '11px' }}
                >
                  ✕ Clear
                </button>
              )}

              <input
                type="text"
                placeholder="🔍 Search Invoice / Agency..."
                value={historySearch}
                onChange={(e) => setHistorySearch(e.target.value)}
                style={{ padding: '6px 10px', borderRadius: '6px', fontSize: '12px', width: '180px' }}
              />
              <button onClick={handleExportToExcel} style={{ background: '#15803d', color: '#fff', border: 'none', padding: '7px 12px', borderRadius: '6px', cursor: 'pointer', fontWeight: 'bold', fontSize: '12px' }}>
                📊 Export to Excel
              </button>
              <button onClick={fetchAllSavedInvoices} style={{ background: '#0284c7', color: '#fff', border: 'none', padding: '7px 12px', borderRadius: '6px', cursor: 'pointer', fontWeight: 'bold', fontSize: '12px' }}>
                🔄 Refresh
              </button>
            </div>
          </div>

          {loadingHistory ? (
            <div style={{ textAlign: 'center', padding: '30px', color: '#0284c7', fontWeight: 'bold' }}>
              ⏳ Loading bills from MySQL...
            </div>
          ) : filteredHistory.length === 0 ? (
            <div style={{ textAlign: 'center', padding: '30px', color: '#64748b', fontSize: '13px' }}>
              निवडलेल्या कालावधीत किंवा शोधामध्ये कोणतेही बिल सापडले नाही!
            </div>
          ) : (
            <div style={{ overflowX: 'auto' }}>
              <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '12.5px' }}>
                <thead>
                  <tr style={{ background: '#f8fafc', borderBottom: '2px solid #cbd5e1', textAlign: 'left' }}>
                    <th style={{ padding: '10px 8px' }}>ID</th>
                    <th style={{ padding: '10px 8px' }}>Invoice No</th>
                    <th style={{ padding: '10px 8px' }}>Date</th>
                    <th style={{ padding: '10px 8px' }}>Agency Name</th>
                    <th style={{ padding: '10px 8px' }}>Division</th>
                    <th style={{ padding: '10px 8px', textAlign: 'right' }}>Total (₹)</th>
                    <th style={{ padding: '10px 8px', textAlign: 'center' }}>Actions</th>
                  </tr>
                </thead>
                <tbody>
                  {filteredHistory.map((inv) => (
                    <tr key={inv.id} style={{ borderBottom: '1px solid #f1f5f9' }}>
                      <td style={{ padding: '8px', fontWeight: 'bold' }}>#{inv?.id}</td>
                      <td style={{ padding: '8px', color: '#0284c7', fontWeight: 'bold' }}>{inv?.invoiceNo || '-'}</td>
                      <td style={{ padding: '8px' }}>{inv?.invoiceDate || '-'}</td>
                      <td style={{ padding: '8px' }}>{inv?.agencyName || '-'}</td>
                      <td style={{ padding: '8px' }}>{inv?.division || 'Jalna I'}</td>
                      <td style={{ padding: '8px', textAlign: 'right', fontWeight: 'bold', color: '#059669' }}>
                        ₹{Number(inv?.grandTotal || 0).toFixed(2)}
                      </td>
                      <td style={{ padding: '8px', textAlign: 'center', display: 'flex', gap: '6px', justifyContent: 'center' }}>
                        <button
                          onClick={() => handleLoadInvoice(inv, 'print')}
                          title="View & Print"
                          style={{ background: '#0284c7', color: '#fff', border: 'none', padding: '5px 10px', borderRadius: '4px', cursor: 'pointer', fontWeight: 'bold', fontSize: '11px' }}
                        >
                          🖨️ Direct Print
                        </button>
                        <button
                          onClick={() => handleLoadInvoice(inv, 'edit')}
                          title="Edit In Billing Form"
                          style={{ background: '#10b981', color: '#fff', border: 'none', padding: '5px 10px', borderRadius: '4px', cursor: 'pointer', fontWeight: 'bold', fontSize: '11px' }}
                        >
                          ✏️ Edit in Form
                        </button>
                        <button
                          onClick={() => handleDeleteInvoice(inv.id, inv.invoiceNo)}
                          title="Delete from MySQL"
                          style={{ background: '#ef4444', color: '#fff', border: 'none', padding: '5px 10px', borderRadius: '4px', cursor: 'pointer', fontWeight: 'bold', fontSize: '11px' }}
                        >
                          🗑️ Delete
                        </button>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
        </div>
      )}

      {/* ================= TAB 1: ADD NEW AGENCY ================= */}
      {activeTab === 'manageAgency' && (
        <div className="no-print" style={{ maxWidth: '1320px', margin: '0 auto 30px', background: '#fff', padding: '24px', borderRadius: '10px', border: '1px solid #e2e8f0' }}>
          <h2 style={{ fontSize: '16px', margin: '0 0 16px 0', color: '#0f172a' }}>
            🏢 Add New Agency (Type GSTIN for Auto-Fetch)
          </h2>
          <form onSubmit={handleAddAgencySubmit} style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(220px, 1fr))', gap: '12px' }}>
            <div>
              <label style={{ fontSize: '11px', fontWeight: 'bold', color: '#0284c7' }}>
                GSTIN (15 Digits) *
              </label>
              <div style={{ position: 'relative' }}>
                <input
                  type="text"
                  maxLength={15}
                  placeholder="e.g. 07AAJCA9764F1ZC"
                  value={newAgencyForm.gstNo}
                  onChange={handleAddAgencyGstChange}
                  required
                  style={{ width: '100%', padding: '8px', borderRadius: '4px', fontWeight: 'bold', background: '#f0f9ff' }}
                />
                {addAgencyLoading && (
                  <span style={{ position: 'absolute', right: '8px', top: '8px', fontSize: '11px', color: '#0284c7', fontWeight: 'bold' }}>
                    ⏳ Fetching...
                  </span>
                )}
              </div>
            </div>

            <div>
              <label style={{ fontSize: '11px', fontWeight: 'bold', color: '#334155' }}>
                Agency Name (Auto-Fetched / Editable) *
              </label>
              <input
                type="text"
                placeholder="Company Name"
                value={newAgencyForm.agencyName}
                onChange={(e) => setNewAgencyForm({ ...newAgencyForm, agencyName: e.target.value })}
                required
                style={{ width: '100%', padding: '8px', borderRadius: '4px', fontWeight: 'bold' }}
              />
            </div>

            <div>
              <label style={{ fontSize: '11px', fontWeight: 'bold', color: '#334155' }}>
                PAN Number (Auto-Extracted)
              </label>
              <input
                type="text"
                placeholder="PAN Number"
                value={newAgencyForm.panNo}
                onChange={(e) => setNewAgencyForm({ ...newAgencyForm, panNo: e.target.value })}
                style={{ width: '100%', padding: '8px', borderRadius: '4px' }}
              />
            </div>

            <div>
              <label style={{ fontSize: '11px', fontWeight: 'bold', color: '#334155' }}>Address (Auto-Fetched)</label>
              <input
                type="text"
                placeholder="Full Address"
                value={newAgencyForm.address}
                onChange={(e) => setNewAgencyForm({ ...newAgencyForm, address: e.target.value })}
                style={{ width: '100%', padding: '8px', borderRadius: '4px' }}
              />
            </div>

            <div>
              <label style={{ fontSize: '11px', fontWeight: 'bold', color: '#334155' }}>Email Address</label>
              <input type="email" placeholder="agency@mail.com" value={newAgencyForm.email} onChange={(e) => setNewAgencyForm({ ...newAgencyForm, email: e.target.value })} style={{ width: '100%', padding: '8px', borderRadius: '4px' }} />
            </div>

            <div>
              <label style={{ fontSize: '11px', fontWeight: 'bold', color: '#334155' }}>Mobile Number</label>
              <input type="text" placeholder="9876543210" value={newAgencyForm.mobile} onChange={(e) => setNewAgencyForm({ ...newAgencyForm, mobile: e.target.value })} style={{ width: '100%', padding: '8px', borderRadius: '4px' }} />
            </div>

            <div>
              <label style={{ fontSize: '11px', fontWeight: 'bold', color: '#334155' }}>LOA Number</label>
              <input type="text" placeholder="LOA / LOE Number" value={newAgencyForm.loaNo} onChange={(e) => setNewAgencyForm({ ...newAgencyForm, loaNo: e.target.value })} style={{ width: '100%', padding: '8px', borderRadius: '4px' }} />
            </div>

            <div>
              <label style={{ fontSize: '11px', fontWeight: 'bold', color: '#334155' }}>LOA Date</label>
              <input type="date" value={newAgencyForm.loaDate} onChange={(e) => setNewAgencyForm({ ...newAgencyForm, loaDate: e.target.value })} style={{ width: '100%', padding: '8px', borderRadius: '4px' }} />
            </div>

            <div>
              <label style={{ fontSize: '11px', fontWeight: 'bold', color: '#334155' }}>Work Order / PO Numbers</label>
              <input type="text" placeholder="PO Numbers" value={newAgencyForm.workOrderNo} onChange={(e) => setNewAgencyForm({ ...newAgencyForm, workOrderNo: e.target.value })} style={{ width: '100%', padding: '8px', borderRadius: '4px' }} />
            </div>

            <div>
              <label style={{ fontSize: '11px', fontWeight: 'bold', color: '#334155' }}>Work Order Dates</label>
              <input type="text" placeholder="e.g. 08.09.2026 & 24.09.2026" value={newAgencyForm.workOrderDates} onChange={(e) => setNewAgencyForm({ ...newAgencyForm, workOrderDates: e.target.value })} style={{ width: '100%', padding: '8px', borderRadius: '4px' }} />
            </div>

            <div>
              <label style={{ fontSize: '11px', fontWeight: 'bold', color: '#334155' }}>Division</label>
              <input type="text" value={newAgencyForm.division} onChange={(e) => setNewAgencyForm({ ...newAgencyForm, division: e.target.value })} style={{ width: '100%', padding: '8px', borderRadius: '4px' }} />
            </div>

            <div>
              <label style={{ fontSize: '11px', fontWeight: 'bold', color: '#334155' }}>Sub-Division</label>
              <input type="text" value={newAgencyForm.subDivision} onChange={(e) => setNewAgencyForm({ ...newAgencyForm, subDivision: e.target.value })} style={{ width: '100%', padding: '8px', borderRadius: '4px' }} />
            </div>

            <div style={{ gridColumn: '1 / -1', marginTop: '10px' }}>
              <button type="submit" style={{ background: '#0284c7', color: '#fff', border: 'none', padding: '10px 18px', borderRadius: '6px', fontWeight: 'bold', cursor: 'pointer' }}>
                ➕ Save Agency & Activate
              </button>
            </div>
          </form>
        </div>
      )}

      {/* ================= TAB 2: MANAGE PARTS & PRICES ================= */}
      {activeTab === 'manageParts' && (
        <div className="no-print" style={{ maxWidth: '1320px', margin: '0 auto 30px', background: '#fff', padding: '24px', borderRadius: '10px', border: '1px solid #e2e8f0' }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '16px' }}>
            <h2 style={{ fontSize: '16px', margin: 0 }}>
              ⚙️ Parts & Tender Rates for: <span style={{ color: '#0284c7' }}>{activeAgencyData.agencyName}</span>
            </h2>
            <div style={{ display: 'flex', gap: '8px', alignItems: 'center' }}>
              <span style={{ fontSize: '12px', fontWeight: 'bold' }}>Select Agency:</span>
              <select value={selectedAgencyId} onChange={(e) => handleSelectAgency(e.target.value)} style={{ padding: '6px 10px', borderRadius: '6px' }}>
                {agencies.map((a) => <option key={a.id} value={a.id}>{a.agencyName}</option>)}
              </select>
            </div>
          </div>

          <form onSubmit={handleAddPartToCurrentAgency} style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(140px, 1fr)) 120px', gap: '10px', background: '#f8fafc', padding: '14px', borderRadius: '8px', marginBottom: '20px' }}>
            <input type="text" placeholder="Part Code" value={newPartForm.code} onChange={(e) => setNewPartForm({ ...newPartForm, code: e.target.value })} required style={{ padding: '8px', borderRadius: '4px' }} />
            <input type="text" placeholder="Part Description" value={newPartForm.name} onChange={(e) => setNewPartForm({ ...newPartForm, name: e.target.value })} required style={{ padding: '8px', borderRadius: '4px' }} />
            <select value={newPartForm.category} onChange={(e) => setNewPartForm({ ...newPartForm, category: e.target.value })} style={{ padding: '8px', borderRadius: '4px' }}>
              <option value="Material MIGO">Material MIGO</option>
              <option value="Services SES">Services SES</option>
              <option value="Centages SES">Centages SES</option>
            </select>
            <select value={newPartForm.unit} onChange={(e) => setNewPartForm({ ...newPartForm, unit: e.target.value })} style={{ padding: '8px', borderRadius: '4px' }}>
              {UNITS_LIST.map((u) => <option key={u} value={u}>{u}</option>)}
            </select>
            <input type="number" placeholder="Rate (₹)" value={newPartForm.rate} onChange={(e) => setNewPartForm({ ...newPartForm, rate: e.target.value })} required style={{ padding: '8px', borderRadius: '4px' }} />
            <button type="submit" style={{ background: '#0284c7', color: '#fff', border: 'none', borderRadius: '4px', fontWeight: 'bold', cursor: 'pointer' }}>
              ➕ Add Part
            </button>
          </form>

          <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '12px' }}>
            <thead>
              <tr style={{ background: '#f1f5f9', borderBottom: '2px solid #cbd5e1', textAlign: 'left' }}>
                <th style={{ padding: '8px' }}>Code</th>
                <th style={{ padding: '8px' }}>Description</th>
                <th style={{ padding: '8px' }}>Category</th>
                <th style={{ padding: '8px' }}>Unit</th>
                <th style={{ padding: '8px', width: '110px' }}>Default Rate (₹)</th>
                <th style={{ padding: '8px', textAlign: 'center' }}>Action</th>
              </tr>
            </thead>
            <tbody>
              {(currentAgency?.catalog || []).map((part) => (
                <tr key={part.id} style={{ borderBottom: '1px solid #f1f5f9' }}>
                  <td style={{ padding: '7px', fontWeight: 'bold' }}>{part.code}</td>
                  <td style={{ padding: '7px' }}>{part.name}</td>
                  <td style={{ padding: '7px', color: '#64748b' }}>{part.category}</td>
                  <td style={{ padding: '7px' }}>{part.unit}</td>
                  <td style={{ padding: '7px' }}>
                    <input
                      type="number"
                      value={part.rate}
                      onChange={(e) => handleUpdateCatalogRate(part.id, e.target.value)}
                      style={{ width: '80px', padding: '4px', borderRadius: '4px', fontWeight: 'bold' }}
                    />
                  </td>
                  <td style={{ padding: '7px', textAlign: 'center' }}>
                    <button onClick={() => handleDeleteCatalogPart(part.id)} style={{ background: '#fee2e2', color: '#dc2626', border: 'none', padding: '4px 8px', borderRadius: '4px', cursor: 'pointer' }}>
                      ✕ Delete
                    </button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}

      {/* ================= TAB 3: BILLING SCREEN ================= */}
      {activeTab === 'billing' && (
        <div className="no-print" style={{ maxWidth: '1320px', margin: '0 auto 30px' }}>
          
          {/* Section 1: Agency Details */}
          <div style={{ background: '#fff', padding: '20px', borderRadius: '10px', border: '1px solid #e2e8f0', marginBottom: '20px' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '14px' }}>
              <div>
                <h2 style={{ fontSize: '15px', margin: 0, fontWeight: 'bold' }}>
                  1. Agency & Work Order Header (Live Online GST Registry Search)
                </h2>
                {gstStatusMsg && (
                  <div style={{ fontSize: '12px', fontWeight: '600', color: gstStatusMsg.includes('✅') ? '#15803d' : '#0284c7', marginTop: '4px' }}>
                    {gstStatusMsg}
                  </div>
                )}
              </div>
              <div style={{ display: 'flex', gap: '8px', alignItems: 'center' }}>
                <button 
                  onClick={handleClearDraft} 
                  title="फॉर्म रिसेट करून नवीन बिल भरणे" 
                  style={{ background: '#f1f5f9', color: '#475569', border: '1px solid #cbd5e1', padding: '6px 10px', borderRadius: '6px', fontSize: '11.5px', fontWeight: 'bold', cursor: 'pointer' }}
                >
                  🔄 Reset Form
                </button>
                <span style={{ fontSize: '12px', fontWeight: 'bold', color: '#0284c7' }}>Agency Quick Select:</span>
                <select
                  value={selectedAgencyId}
                  onChange={(e) => handleSelectAgency(e.target.value)}
                  style={{ padding: '6px 12px', borderRadius: '6px', border: '1.5px solid #0284c7', fontWeight: 'bold' }}
                >
                  {agencies.map((a) => (
                    <option key={a.id} value={a.id}>{a.agencyName} ({a.gstNo})</option>
                  ))}
                </select>
              </div>
            </div>

            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(210px, 1fr))', gap: '10px' }}>
              <div>
                <label style={{ fontSize: '11px', fontWeight: 'bold', color: '#0284c7' }}>
                  GSTIN (Type 15 Digits for Live Lookup) *
                </label>
                <div style={{ position: 'relative' }}>
                  <input
                    type="text"
                    maxLength={15}
                    value={activeAgencyData.gstNo}
                    onChange={handleGstInputChange}
                    placeholder="e.g. 27KTWPS1687C1ZR"
                    style={{ width: '100%', padding: '6px', borderRadius: '4px', fontWeight: 'bold', background: '#f0f9ff' }}
                  />
                  {gstLoading && (
                    <span style={{ position: 'absolute', right: '8px', top: '6px', fontSize: '11px', color: '#0284c7', fontWeight: 'bold' }}>
                      ⏳ Fetching...
                    </span>
                  )}
                </div>
              </div>

              <div>
                <label style={{ fontSize: '11px', fontWeight: 'bold' }}>Agency Name (Auto-Fetched / Editable)</label>
                <input
                  type="text"
                  value={activeAgencyData.agencyName}
                  onChange={(e) => handleActiveAgencyChange('agencyName', e.target.value)}
                  style={{ width: '100%', padding: '6px', borderRadius: '4px', fontWeight: 'bold' }}
                />
              </div>

              <div>
                <label style={{ fontSize: '11px', fontWeight: 'bold', color: '#0284c7' }}>MSEDCL Bill Address Division *</label>
                <select
                  value={targetDivisionName}
                  onChange={(e) => setTargetDivisionName(e.target.value)}
                  style={{ width: '100%', padding: '6px', borderRadius: '4px', fontWeight: 'bold', background: '#f0f9ff' }}
                >
                  <option value="Division No.I">M.S.E.D.C.L. Division No.I</option>
                  <option value="Division No.II">M.S.E.D.C.L. Division No.II</option>
                </select>
              </div>

              <div>
                <label style={{ fontSize: '11px', fontWeight: 'bold' }}>PAN Number (Auto-Extracted)</label>
                <input
                  type="text"
                  value={activeAgencyData.panNo}
                  onChange={(e) => handleActiveAgencyChange('panNo', e.target.value)}
                  style={{ width: '100%', padding: '6px', borderRadius: '4px' }}
                />
              </div>

              <div>
                <label style={{ fontSize: '11px', fontWeight: 'bold' }}>Contractor Sub-Division</label>
                <input
                  type="text"
                  value={activeAgencyData.subDivision}
                  onChange={(e) => handleActiveAgencyChange('subDivision', e.target.value)}
                  style={{ width: '100%', padding: '6px', borderRadius: '4px' }}
                />
              </div>

              <div>
                <label style={{ fontSize: '11px', fontWeight: 'bold', color: '#0284c7' }}>
                  Invoice No (Auto-Increments) *
                </label>
                <input
                  type="text"
                  value={invoiceHeader.invoiceNo}
                  onChange={(e) => setInvoiceHeader({ ...invoiceHeader, invoiceNo: e.target.value })}
                  style={{ width: '100%', padding: '6px', borderRadius: '4px', fontWeight: 'bold', background: '#f0f9ff' }}
                />
              </div>

              <div>
                <label style={{ fontSize: '11px', fontWeight: 'bold' }}>Invoice Date</label>
                <input
                  type="date"
                  value={invoiceHeader.invoiceDate}
                  onChange={(e) => setInvoiceHeader({ ...invoiceHeader, invoiceDate: e.target.value })}
                  style={{ width: '100%', padding: '6px', borderRadius: '4px' }}
                />
              </div>

              <div>
                <label style={{ fontSize: '11px', fontWeight: 'bold' }}>33 KV Amount (₹)</label>
                <input
                  type="number"
                  value={invoiceHeader.sch33kvAmount}
                  onChange={(e) => setInvoiceHeader({ ...invoiceHeader, sch33kvAmount: Number(e.target.value) })}
                  style={{ width: '100%', padding: '6px', borderRadius: '4px' }}
                />
              </div>

              <div>
                <label style={{ fontSize: '11px', fontWeight: 'bold' }}>Below Rate %</label>
                <input
                  type="number"
                  value={invoiceHeader.belowRatePercent}
                  onChange={(e) => setInvoiceHeader({ ...invoiceHeader, belowRatePercent: Number(e.target.value) })}
                  style={{ width: '100%', padding: '6px', borderRadius: '4px' }}
                />
              </div>

              <div>
                <label style={{ fontSize: '11px', fontWeight: 'bold' }}>LOA Number</label>
                <input
                  type="text"
                  value={invoiceHeader.loaNo}
                  onChange={(e) => setInvoiceHeader({ ...invoiceHeader, loaNo: e.target.value })}
                  style={{ width: '100%', padding: '6px', borderRadius: '4px' }}
                />
              </div>

              <div>
                <label style={{ fontSize: '11px', fontWeight: 'bold' }}>LOA Date</label>
                <input
                  type="date"
                  value={invoiceHeader.loaDate}
                  onChange={(e) => setInvoiceHeader({ ...invoiceHeader, loaDate: e.target.value })}
                  style={{ width: '100%', padding: '6px', borderRadius: '4px' }}
                />
              </div>

              <div>
                <label style={{ fontSize: '11px', fontWeight: 'bold' }}>Work Order / PO Numbers</label>
                <input
                  type="text"
                  value={invoiceHeader.workOrderNo}
                  onChange={(e) => setInvoiceHeader({ ...invoiceHeader, workOrderNo: e.target.value })}
                  style={{ width: '100%', padding: '6px', borderRadius: '4px' }}
                />
              </div>

              <div>
                <label style={{ fontSize: '11px', fontWeight: 'bold' }}>Work Order Dates</label>
                <input
                  type="text"
                  value={invoiceHeader.workOrderDates}
                  onChange={(e) => setInvoiceHeader({ ...invoiceHeader, workOrderDates: e.target.value })}
                  style={{ width: '100%', padding: '6px', borderRadius: '4px' }}
                />
              </div>

              <div>
                <label style={{ fontSize: '11px', fontWeight: 'bold' }}>Work Started Date</label>
                <input
                  type="date"
                  value={invoiceHeader.workStartedDate}
                  onChange={(e) => setInvoiceHeader({ ...invoiceHeader, workStartedDate: e.target.value })}
                  style={{ width: '100%', padding: '6px', borderRadius: '4px' }}
                />
              </div>

              <div>
                <label style={{ fontSize: '11px', fontWeight: 'bold' }}>Work Completed Date</label>
                <input
                  type="date"
                  value={invoiceHeader.workCompletedDate}
                  onChange={(e) => setInvoiceHeader({ ...invoiceHeader, workCompletedDate: e.target.value })}
                  style={{ width: '100%', padding: '6px', borderRadius: '4px' }}
                />
              </div>

              <div>
                <label style={{ fontSize: '11px', fontWeight: 'bold' }}>Commissioning Date</label>
                <input
                  type="date"
                  value={invoiceHeader.commissioningDate}
                  onChange={(e) => setInvoiceHeader({ ...invoiceHeader, commissioningDate: e.target.value })}
                  style={{ width: '100%', padding: '6px', borderRadius: '4px' }}
                />
              </div>
            </div>
          </div>

          {/* Section 2: Detailed Measurement Items */}
          <div style={{ background: '#fff', padding: '20px', borderRadius: '10px', border: '1px solid #e2e8f0' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '14px' }}>
              <h2 style={{ fontSize: '15px', margin: 0, fontWeight: 'bold' }}>
                2. Measurement Line Items (11 KV Total Auto-Sum: ₹{calculatedSch11kv.toFixed(2)})
              </h2>
              <button onClick={addMeasurementRow} style={{ background: '#0284c7', color: '#fff', border: 'none', padding: '8px 14px', borderRadius: '6px', cursor: 'pointer', fontWeight: 'bold' }}>
                ➕ Add Line Item
              </button>
            </div>

            <div style={{ overflowX: 'auto' }}>
              <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '12px' }}>
                <thead>
                  <tr style={{ background: '#f8fafc', borderBottom: '2px solid #e2e8f0', textAlign: 'left' }}>
                    <th style={{ padding: '8px' }}>PO & MO ID</th>
                    <th style={{ padding: '8px' }}>Location</th>
                    <th style={{ padding: '8px' }}>Category</th>
                    <th style={{ padding: '8px' }}>Select Agency Part (Auto-Suggest)</th>
                    <th style={{ padding: '8px' }}>Description (Editable)</th>
                    <th style={{ padding: '8px', width: '70px' }}>Unit</th>
                    <th style={{ padding: '8px', width: '60px' }}>Qty</th>
                    <th style={{ padding: '8px', width: '80px' }}>Rate (₹)</th>
                    <th style={{ padding: '8px', width: '90px' }}>Amount (₹)</th>
                    <th style={{ padding: '8px', width: '40px' }}>Del</th>
                  </tr>
                </thead>
                <tbody>
                  {measurementItems.map((item, idx) => (
                    <tr key={idx} style={{ borderBottom: '1px solid #f1f5f9' }}>
                      <td style={{ padding: '6px' }}>
                        <input type="text" placeholder="PO ID" value={item.erpPoId} onChange={(e) => handleMeasurementChange(idx, 'erpPoId', e.target.value)} style={{ width: '85px', padding: '6px', marginBottom: '3px', borderRadius: '4px' }} />
                        <input type="text" placeholder="MO ID" value={item.erpMoId} onChange={(e) => handleMeasurementChange(idx, 'erpMoId', e.target.value)} style={{ width: '85px', padding: '6px', borderRadius: '4px' }} />
                      </td>
                      <td style={{ padding: '6px' }}>
                        <input type="text" value={item.locationName} onChange={(e) => handleMeasurementChange(idx, 'locationName', e.target.value)} style={{ width: '120px', padding: '6px', borderRadius: '4px' }} />
                      </td>
                      <td style={{ padding: '6px' }}>
                        <select value={item.itemCategory} onChange={(e) => handleMeasurementChange(idx, 'itemCategory', e.target.value)} style={{ width: '110px', padding: '6px', borderRadius: '4px' }}>
                          <option value="Material MIGO">Material MIGO</option>
                          <option value="Services SES">Services SES</option>
                          <option value="Centages SES">Centages SES</option>
                        </select>
                      </td>
                      <td style={{ padding: '6px' }}>
                        <select
                          onChange={(e) => handleAutoSuggestPart(idx, e.target.value)}
                          defaultValue=""
                          style={{ width: '150px', padding: '6px', fontWeight: 'bold', borderRadius: '4px' }}
                        >
                          <option value="" disabled>-- Auto Suggest --</option>
                          {currentAgency.catalog?.filter((p) => p.category === item.itemCategory).map((p) => (
                            <option key={p.id} value={p.code}>{p.code} - {p.name.substring(0, 16)}...</option>
                          ))}
                        </select>
                      </td>
                      <td style={{ padding: '6px' }}>
                        <input type="text" value={item.description} onChange={(e) => handleMeasurementChange(idx, 'description', e.target.value)} style={{ width: '160px', padding: '6px', borderRadius: '4px' }} />
                      </td>
                      <td style={{ padding: '6px' }}>
                        <select value={item.unit} onChange={(e) => handleMeasurementChange(idx, 'unit', e.target.value)} style={{ width: '65px', padding: '6px', borderRadius: '4px' }}>
                          {UNITS_LIST.map((u) => <option key={u} value={u}>{u}</option>)}
                        </select>
                      </td>
                      <td style={{ padding: '6px' }}>
                        <input type="number" value={item.qty} onChange={(e) => handleMeasurementChange(idx, 'qty', e.target.value)} style={{ width: '55px', padding: '6px', fontWeight: 'bold', borderRadius: '4px' }} />
                      </td>
                      <td style={{ padding: '6px' }}>
                        <input type="number" value={item.rate} onChange={(e) => handleMeasurementChange(idx, 'rate', e.target.value)} style={{ width: '75px', padding: '6px', borderRadius: '4px' }} />
                      </td>
                      <td style={{ padding: '6px', fontWeight: 'bold', color: '#000' }}>₹{Number(item.amount || 0).toFixed(2)}</td>
                      <td style={{ padding: '6px', textAlign: 'center' }}>
                        <button onClick={() => removeMeasurementRow(idx)} style={{ background: '#fee2e2', color: '#dc2626', border: 'none', padding: '4px 6px', borderRadius: '4px', cursor: 'pointer' }}>✕</button>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>

        </div>
      )}

      {/* ================= PRINT BUFFER ================= */}
      <div className="print-area">
        {(printCopyMode === 'ALL' || printCopyMode === 'ORIGINAL') && renderDocumentSet('ORIGINAL')}
        {(printCopyMode === 'ALL' || printCopyMode === 'DUPLICATE') && renderDocumentSet('DUPLICATE')}
      </div>

    </div>
  );
}