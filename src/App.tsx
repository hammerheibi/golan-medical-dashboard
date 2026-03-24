import React, { useState, useMemo, useEffect } from 'react';
import * as XLSX from 'xlsx';
import html2canvas from 'html2canvas';
import jsPDF from 'jspdf';
import { 
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer,
  PieChart, Pie, Cell
} from 'recharts';
import { 
  UploadCloud, Users, Activity, Briefcase, Stethoscope, 
  MapPin, AlertCircle, ShieldAlert, HeartPulse, FileText, Filter, XCircle,
  ChevronRight, ChevronLeft, AlertTriangle, Siren, Crosshair, Award, Download,
  X, User
} from 'lucide-react';

const COLORS = ['#3b82f6', '#10b981', '#f59e0b', '#ef4444', '#8b5cf6', '#14b8a6', '#f97316'];

function App() {
  const [data, setData] = useState<any[]>([]);
  const [fileName, setFileName] = useState<string>('');
  const [loading, setLoading] = useState(false);
  const [isExporting, setIsExporting] = useState(false);

  // Filters State
  const [filterSettlement, setFilterSettlement] = useState<string>('all');
  const [filterRole, setFilterRole] = useState<string>('all');
  const [filterAvailability, setFilterAvailability] = useState<string>('all');
  const [filterAssigned, setFilterAssigned] = useState<string>('all');

  // Drill-down State
  const [drillDown, setDrillDown] = useState<{type: string, value: string} | null>(null);
  
  // Row Detail Modal State
  const [selectedRowData, setSelectedRowData] = useState<any | null>(null);

  // Pagination State
  const [currentPage, setCurrentPage] = useState(1);
  const rowsPerPage = 15;

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setLoading(true);
    setFileName(file.name);

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const bstr = evt.target?.result;
        const wb = XLSX.read(bstr, { type: 'binary' });
        const wsname = wb.SheetNames[0];
        const ws = wb.Sheets[wsname];
        const jsonData = XLSX.utils.sheet_to_json(ws, { defval: '', raw: false, dateNF: 'dd/MM/yyyy' });
        setData(jsonData);
        
        clearFilters();
      } catch (error) {
        console.error("Error reading excel file", error);
        alert("שגיאה בקריאת הקובץ. אנא ודא שזהו קובץ אקסל תקין.");
      } finally {
        setLoading(false);
      }
    };
    reader.readAsBinaryString(file);
  };

  const clearFilters = () => {
    setFilterSettlement('all');
    setFilterRole('all');
    setFilterAvailability('all');
    setFilterAssigned('all');
    setDrillDown(null);
  };

  const handleDrillDown = (type: string, value: string) => {
    setDrillDown({ type, value });
    setCurrentPage(1);
    setTimeout(() => {
      document.getElementById('data-table')?.scrollIntoView({ behavior: 'smooth', block: 'start' });
    }, 100);
  };

  const downloadPDF = async () => {
    const element = document.getElementById('dashboard-content');
    if (!element) return;
    
    setIsExporting(true);
    try {
      const canvas = await html2canvas(element, { scale: 2, useCORS: true });
      const imgData = canvas.toDataURL('image/png');
      
      const pdf = new jsPDF('p', 'px', [canvas.width, canvas.height]);
      pdf.addImage(imgData, 'PNG', 0, 0, canvas.width, canvas.height);
      pdf.save('Golan_Medical_Dashboard.pdf');
    } catch (error) {
      console.error('Error generating PDF', error);
      alert('שגיאה ביצירת קובץ ה-PDF');
    } finally {
      setIsExporting(false);
    }
  };

  // --- Data Processors ---
  const getCol = (row: any, keywords: string[]) => {
    const key = Object.keys(row).find(k => keywords.some(kw => k.includes(kw)));
    return key ? String(row[key]).trim() : '';
  };

  const getName = (row: any) => {
    let name = getCol(row, ['שם מלא']);
    if (name) return name;
    
    const first = getCol(row, ['פרטי']);
    const last = getCol(row, ['משפחה']);
    if (first || last) return `${first} ${last}`.trim();
    
    return getCol(row, ['שם', 'זיהוי']);
  };

  const getRoleCategory = (rawRole: string) => {
    if (!rawRole) return 'אחר/לא מוגדר';
    if (rawRole.includes('רופא')) return 'רופאים';
    if (rawRole.includes('פרמדיק') || rawRole.includes('פראמדיק')) return 'פרמדיקים';
    if (rawRole.includes('אח/ות') || rawRole.includes('אחות') || rawRole.includes('אחים') || rawRole === 'אח' || rawRole.includes('אח ')) return 'אחים/אחיות';
    if (rawRole.includes('חובש') || rawRole.includes('מע"ר') || rawRole.includes('עזרה ראשונה') || rawRole.includes('מגיש')) return 'חובשים/מע"רים';
    return 'אחר/לא מוגדר';
  };

  const isAssignedToEmergency = (val: string) => {
    return (val.includes('כן') || val === '1' || val === 'True') ? 'yes' : 'no';
  };

  // Extract unique options for filters
  const filterOptions = useMemo(() => {
    if (!data.length) return { settlements: [], roles: [], availabilities: [] };
    
    const settlements = new Set<string>();
    const roles = new Set<string>();
    const availabilities = new Set<string>();

    data.forEach(row => {
      const s = getCol(row, ['יישוב', 'ישוב']);
      if (s) settlements.add(s);
      
      const r = getRoleCategory(getCol(row, ['הכשרה', 'מקצוע', 'רופא', 'פרמדיק', 'אח', 'חובש', 'מע"ר']));
      if (r) roles.add(r);

      const a = getCol(row, ['זמינות']);
      if (a) availabilities.add(a);
    });

    return {
      settlements: Array.from(settlements).sort(),
      roles: Array.from(roles).sort(),
      availabilities: Array.from(availabilities).sort()
    };
  }, [data]);

  // Apply filters to data (Main top filters)
  const filteredData = useMemo(() => {
    return data.filter(row => {
      const s = getCol(row, ['יישוב', 'ישוב']);
      const r = getRoleCategory(getCol(row, ['הכשרה', 'מקצוע', 'רופא', 'פרמדיק', 'אח', 'חובש', 'מע"ר']));
      const a = getCol(row, ['זמינות']);
      const assigned = isAssignedToEmergency(getCol(row, ['שובץ', 'מכלול בחירום']));

      if (filterSettlement !== 'all' && s !== filterSettlement) return false;
      if (filterRole !== 'all' && r !== filterRole) return false;
      if (filterAvailability !== 'all' && a !== filterAvailability) return false;
      if (filterAssigned !== 'all' && assigned !== filterAssigned) return false;

      return true;
    });
  }, [data, filterSettlement, filterRole, filterAvailability, filterAssigned]);

  // Apply Drill-down to filtered data for the Table
  const tableData = useMemo(() => {
    if (!drillDown) return filteredData;
    return filteredData.filter(row => {
      const rowString = Object.values(row).join(' ').toLowerCase();
      const roleCat = getRoleCategory(getCol(row, ['הכשרה', 'מקצוע', 'רופא', 'פרמדיק', 'אח', 'חובש', 'מע"ר']));
      
      switch (drillDown.type) {
        case 'role': 
          return roleCat === drillDown.value;
        case 'missing_equipment': 
          const eq = getCol(row, ['תיק ציוד', 'תיק רפואי אישי']);
          return !(eq.includes('כן') || eq === '1' || eq.includes('יש'));
        case 'availability':
          return getCol(row, ['זמינות']) === drillDown.value;
        case 'high_availability':
          const av = getCol(row, ['זמינות']);
          return av === '4' || av === '5';
        case 'assigned':
          return isAssignedToEmergency(getCol(row, ['שובץ', 'מכלול בחירום'])) === 'yes';
        case 'bound':
          const bd = getCol(row, ['מרותק']);
          return bd.includes('כן') || bd === '1' || bd === 'True';
        case 'settlement':
          return getCol(row, ['יישוב', 'ישוב']) === drillDown.value;
        case 'organization':
          if (drillDown.value === 'מד"א') return rowString.includes('מד"א') || rowString.includes('מגן דוד');
          if (drillDown.value === 'איחוד הצלה') return rowString.includes('איחוד');
          if (drillDown.value === 'צה"ל') return rowString.includes('צה"ל') || rowString.includes('צבא') || rowString.includes('צבאית');
          return !rowString.includes('מד"א') && !rowString.includes('מגן דוד') && !rowString.includes('איחוד') && !rowString.includes('צה"ל') && !rowString.includes('צבא');
        case 'constraint':
          const constraints = getCol(row, ['אילוצ', 'כוננות', 'מילואים']);
          if (drillDown.value === 'כיתת כוננות') return constraints.includes('כוננות') || rowString.includes('כוננות');
          if (drillDown.value === 'מילואים') return constraints.includes('מילואים') || rowString.includes('מילואים');
          return false;
        case 'active_ready':
          return rowString.includes('acls') || (rowString.includes('פעיל') && !rowString.includes('לא פעיל') && !rowString.includes('שאינו פעיל'));
        case 'specialty':
          if (roleCat !== 'רופאים') return false;
          const isTrauma = rowString.includes('טראומה') || rowString.includes('נמרץ') || rowString.includes('כירורגי') || rowString.includes('הרדמה') || rowString.includes('אורתופדיה');
          const isFam = rowString.includes('משפחה') || rowString.includes('פנימי');
          const isPed = rowString.includes('ילדים');
          const isPsych = rowString.includes('פסיכיאטריה');
          const isWom = rowString.includes('נשים') || rowString.includes('ילודה');
          if (drillDown.value === 'טראומה, נמרץ וכירורגיה') return isTrauma;
          if (drillDown.value === 'משפחה ופנימית') return !isTrauma && isFam;
          if (drillDown.value === 'ילדים') return !isTrauma && !isFam && isPed;
          if (drillDown.value === 'פסיכיאטריה') return !isTrauma && !isFam && !isPed && isPsych;
          if (drillDown.value === 'נשים וילודה') return !isTrauma && !isFam && !isPed && !isPsych && isWom;
          return !isTrauma && !isFam && !isPed && !isPsych && !isWom;
        default: return true;
      }
    });
  }, [filteredData, drillDown]);

  // Reset pagination when table data changes
  useEffect(() => {
    setCurrentPage(1);
  }, [tableData]);

  // Compute stats based on FILTERED data (Not drilled down data, stats stay stable!)
  const stats = useMemo(() => {
    if (!filteredData.length && !data.length) return null;

    let totalPersonnel = filteredData.length;
    let settlementsCount: Record<string, number> = {};
    let availabilityScores: Record<string, number> = { '1':0, '2':0, '3':0, '4':0, '5':0 };
    let rolesCount: Record<string, number> = { 'רופאים': 0, 'פרמדיקים': 0, 'אחים/אחיות': 0, 'חובשים/מע"רים': 0, 'אחר/לא מוגדר': 0 };
    
    let assignedToEmergency = 0;
    let boundToWorkplace = 0;
    let highAvailability = 0;
    let missingEquipment = 0;

    let equipmentStatus = { 'יש תיק אישי': 0, 'אין תיק': 0 };
    
    let specialtiesCount: Record<string, number> = {};
    let constraintsCount: Record<string, number> = { 'כיתת כוננות': 0, 'מילואים': 0 };
    let orgsCount: Record<string, number> = { 'מד"א': 0, 'איחוד הצלה': 0, 'צה"ל': 0, 'אחר / לא שויך': 0 };
    let aclsOrActiveCount = 0;

    filteredData.forEach(row => {
      const rowString = Object.values(row).join(' ').toLowerCase();

      const settlement = getCol(row, ['יישוב', 'ישוב']);
      if (settlement) settlementsCount[settlement] = (settlementsCount[settlement] || 0) + 1;

      const availability = getCol(row, ['זמינות']);
      if (availability && availabilityScores[availability] !== undefined) {
         availabilityScores[availability]++;
         if (availability === '4' || availability === '5') highAvailability++;
      }

      const roleCat = getRoleCategory(getCol(row, ['הכשרה', 'מקצוע', 'רופא', 'פרמדיק', 'אח', 'חובש', 'מע"ר']));
      rolesCount[roleCat]++;

      const assigned = getCol(row, ['שובץ', 'מכלול בחירום']);
      if (isAssignedToEmergency(assigned) === 'yes') assignedToEmergency++;

      const bound = getCol(row, ['מרותק']);
      if (bound.includes('כן') || bound === '1' || bound === 'True') boundToWorkplace++;

      const equipment = getCol(row, ['תיק ציוד', 'תיק רפואי אישי']);
      if (equipment.includes('כן') || equipment === '1' || equipment.includes('יש')) {
        equipmentStatus['יש תיק אישי']++;
      } else {
        equipmentStatus['אין תיק']++;
        missingEquipment++;
      }

      const constraints = getCol(row, ['אילוצ', 'כוננות', 'מילואים']);
      if (constraints.includes('כוננות') || rowString.includes('כוננות')) constraintsCount['כיתת כוננות']++;
      if (constraints.includes('מילואים') || rowString.includes('מילואים')) constraintsCount['מילואים']++;
      
      if (roleCat === 'רופאים') {
         if (rowString.includes('טראומה') || rowString.includes('נמרץ') || rowString.includes('כירורגי') || rowString.includes('הרדמה') || rowString.includes('אורתופדיה')) {
            specialtiesCount['טראומה, נמרץ וכירורגיה'] = (specialtiesCount['טראומה, נמרץ וכירורגיה'] || 0) + 1;
         } else if (rowString.includes('משפחה') || rowString.includes('פנימי')) {
            specialtiesCount['משפחה ופנימית'] = (specialtiesCount['משפחה ופנימית'] || 0) + 1;
         } else if (rowString.includes('ילדים')) {
            specialtiesCount['ילדים'] = (specialtiesCount['ילדים'] || 0) + 1;
         } else if (rowString.includes('פסיכיאטריה')) {
            specialtiesCount['פסיכיאטריה'] = (specialtiesCount['פסיכיאטריה'] || 0) + 1;
         } else if (rowString.includes('נשים') || rowString.includes('ילודה')) {
            specialtiesCount['נשים וילודה'] = (specialtiesCount['נשים וילודה'] || 0) + 1;
         } else {
            specialtiesCount['אחר / לא מוגדר'] = (specialtiesCount['אחר / לא מוגדר'] || 0) + 1;
         }
      }

      if (['פרמדיקים', 'חובשים/מע"רים'].includes(roleCat)) {
         if (rowString.includes('מד"א') || rowString.includes('מגן דוד')) orgsCount['מד"א']++;
         else if (rowString.includes('איחוד')) orgsCount['איחוד הצלה']++;
         else if (rowString.includes('צה"ל') || rowString.includes('צבא') || rowString.includes('צבאית')) orgsCount['צה"ל']++;
         else orgsCount['אחר / לא שויך']++; 
      }

      if (rowString.includes('acls') || (rowString.includes('פעיל') && !rowString.includes('לא פעיל') && !rowString.includes('שאינו פעיל'))) {
         aclsOrActiveCount++;
      }
    });

    const settlementsData = Object.keys(settlementsCount)
      .map(name => ({ name, ערך: settlementsCount[name] }))
      .sort((a, b) => b.ערך - a.ערך)
      .slice(0, 10);

    const rolesData = Object.keys(rolesCount)
      .map(name => ({ name, ערך: rolesCount[name] }))
      .filter(item => item.ערך > 0);

    const availabilityData = Object.keys(availabilityScores).map(score => ({
      name: `רמה ${score}`, ערך: availabilityScores[score]
    }));

    const equipmentData = Object.keys(equipmentStatus).map(name => ({
      name, ערך: equipmentStatus[name as keyof typeof equipmentStatus]
    }));

    const specData = Object.keys(specialtiesCount).map(k => ({ name: k, ערך: specialtiesCount[k] })).filter(item => item.ערך > 0);
    const orgsData = Object.keys(orgsCount).map(k => ({ name: k, ערך: orgsCount[k] })).filter(item => item.ערך > 0);
    const constData = Object.keys(constraintsCount).map(k => ({ name: k, ערך: constraintsCount[k] })).filter(item => item.ערך > 0);

    return {
      totalPersonnel,
      assignedToEmergency,
      boundToWorkplace,
      highAvailability,
      missingEquipment,
      settlementsData,
      rolesData,
      availabilityData,
      equipmentData,
      specData,
      orgsData,
      constData,
      aclsOrActiveCount
    };
  }, [filteredData, data.length]);

  const Card = ({ title, value, icon, subtitle, colorClass = "bg-blue-50 text-blue-600", onClick, isSelected }: any) => (
    <div 
      onClick={onClick}
      className={`group bg-white rounded-xl shadow-sm border ${isSelected ? 'border-blue-500 ring-2 ring-blue-200' : 'border-slate-100 hover:border-blue-300'} p-6 flex flex-col items-start transition-all ${onClick ? 'cursor-pointer hover:shadow-md' : ''} relative overflow-hidden`}
    >
      <div className="flex items-center justify-between w-full mb-4">
        <h3 className="text-slate-500 font-medium relative z-10">{title}</h3>
        <div className={`p-2 rounded-lg ${colorClass} relative z-10`}>{icon}</div>
      </div>
      <p className="text-3xl font-bold text-slate-800 relative z-10">{value}</p>
      {subtitle && <p className="text-sm text-slate-400 mt-2 relative z-10">{subtitle}</p>}
      
      {onClick && (
        <div className="absolute bottom-0 left-0 w-full bg-blue-50 py-1 text-center translate-y-full group-hover:translate-y-0 transition-transform duration-300">
          <span className="text-xs text-blue-600 font-bold">לחץ לסינון הרשומות 👇</span>
        </div>
      )}
    </div>
  );

  const totalPages = Math.max(1, Math.ceil(tableData.length / rowsPerPage));
  const currentTableData = tableData.slice((currentPage - 1) * rowsPerPage, currentPage * rowsPerPage);

  return (
    <div className="min-h-screen bg-slate-50 p-8 font-sans" dir="rtl">
      
      {/* Header */}
      <header className="mb-8 flex flex-col md:flex-row justify-between items-center bg-white p-6 rounded-2xl shadow-sm border border-slate-100">
        <div>
          <h1 className="text-2xl font-bold text-slate-800 flex items-center gap-2">
            <HeartPulse className="text-red-500" /> 
            מועצה אזורית גולן - מיפוי כוח אדם רפואי בחירום
          </h1>
          <p className="text-slate-500 mt-1 text-sm">מערכת תמונת מצב רשותית - נתונים מאובטחים לוקאלית</p>
        </div>
        
        <div className="mt-4 md:mt-0 flex gap-3">
          {stats && (
            <button 
              onClick={downloadPDF}
              disabled={isExporting}
              className="bg-white border border-slate-200 text-slate-700 hover:bg-slate-50 px-4 py-3 rounded-xl shadow-sm transition-all flex items-center gap-2 font-medium disabled:opacity-50"
            >
              {isExporting ? <div className="animate-spin rounded-full h-5 w-5 border-b-2 border-slate-700"></div> : <Download size={20} />}
              ייצא ל-PDF
            </button>
          )}
          <label className="cursor-pointer bg-blue-600 hover:bg-blue-700 text-white px-6 py-3 rounded-xl shadow-md transition-all flex items-center gap-2 font-medium">
            <UploadCloud size={20} />
            {fileName ? 'החלף קובץ' : 'טען אקסל'}
            <input 
              type="file" 
              accept=".xlsx, .xls" 
              className="hidden" 
              onChange={handleFileUpload} 
            />
          </label>
        </div>
      </header>

      {/* Main Content */}
      {loading ? (
        <div className="flex justify-center items-center h-64">
          <div className="animate-spin rounded-full h-12 w-12 border-b-2 border-blue-600"></div>
        </div>
      ) : !stats ? (
        <div className="bg-white rounded-2xl border border-dashed border-slate-300 h-96 flex flex-col justify-center items-center text-slate-400">
          <UploadCloud size={64} className="mb-4 text-slate-300" />
          <h2 className="text-xl font-medium mb-2">אנא העלה קובץ נתונים</h2>
          <p className="text-sm text-center max-w-md">העלה את קובץ האקסל המעודכן. הנתונים אינם נשמרים בשרת ומעובדים ישירות בדפדפן שלך בלבד, לשמירה על סודיות ופרטיות.</p>
        </div>
      ) : (
        <div id="dashboard-content" className="space-y-6 bg-slate-50 p-2 -m-2 rounded-xl">
          
          {/* Filters Bar */}
          <div className="bg-white p-4 rounded-xl shadow-sm border border-slate-200 flex flex-wrap items-center gap-4" data-html2canvas-ignore="true">
             <div className="flex items-center gap-2 text-slate-700 font-bold ml-2">
               <Filter size={18} />
               סינונים מתקדמים:
             </div>
             
             <select 
               value={filterSettlement} 
               onChange={e => setFilterSettlement(e.target.value)}
               className="bg-slate-50 border border-slate-200 text-slate-700 rounded-lg px-4 py-2 outline-none focus:ring-2 focus:ring-blue-500"
             >
               <option value="all">כל היישובים</option>
               {filterOptions.settlements.map(s => <option key={s} value={s}>{s}</option>)}
             </select>

             <select 
               value={filterRole} 
               onChange={e => setFilterRole(e.target.value)}
               className="bg-slate-50 border border-slate-200 text-slate-700 rounded-lg px-4 py-2 outline-none focus:ring-2 focus:ring-blue-500"
             >
               <option value="all">כל ההכשרות</option>
               {filterOptions.roles.map(r => <option key={r} value={r}>{r}</option>)}
             </select>

             <select 
               value={filterAvailability} 
               onChange={e => setFilterAvailability(e.target.value)}
               className="bg-slate-50 border border-slate-200 text-slate-700 rounded-lg px-4 py-2 outline-none focus:ring-2 focus:ring-blue-500"
             >
               <option value="all">כל רמות הזמינות</option>
               {filterOptions.availabilities.map(a => <option key={a} value={a}>זמינות: {a}</option>)}
             </select>

             <select 
               value={filterAssigned} 
               onChange={e => setFilterAssigned(e.target.value)}
               className="bg-slate-50 border border-slate-200 text-slate-700 rounded-lg px-4 py-2 outline-none focus:ring-2 focus:ring-blue-500"
             >
               <option value="all">שיבוץ בחירום (הכל)</option>
               <option value="yes">שובצו למכלול</option>
               <option value="no">טרם שובצו</option>
             </select>

             <button 
               onClick={clearFilters}
               className={`flex items-center gap-1 text-sm font-medium mr-auto px-3 py-2 rounded-lg transition-colors ${
                 filterSettlement !== 'all' || filterRole !== 'all' || filterAvailability !== 'all' || filterAssigned !== 'all' || drillDown !== null
                 ? 'text-red-500 hover:text-red-700 bg-red-50' : 'text-slate-400 opacity-50 cursor-not-allowed'
               }`}
             >
               <XCircle size={16} /> נקה את כל הסינונים
             </button>
          </div>

          {filteredData.length === 0 ? (
            <div className="bg-amber-50 text-amber-800 p-8 rounded-xl text-center border border-amber-200 font-medium">
              לא נמצאו תוצאות התואמות לסינון שבחרת. אנא שנה את הפילטרים.
            </div>
          ) : (
            <>
              {/* Top Stats Row */}
              <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 xl:grid-cols-6 gap-6">
                <Card 
                  title='סה"כ כוח אדם' 
                  value={stats.totalPersonnel} 
                  icon={<Users />} 
                  subtitle={`מתוך ${data.length} בסה"כ`} 
                />
                
                <Card 
                  title='שובצו למכלולים' 
                  value={stats.assignedToEmergency} 
                  icon={<ShieldAlert />} 
                  colorClass="bg-indigo-50 text-indigo-600"
                  subtitle={`${((stats.assignedToEmergency / stats.totalPersonnel) * 100).toFixed(1)}% מסך הצוות המוצג`} 
                  onClick={() => handleDrillDown('assigned', 'yes')}
                  isSelected={drillDown?.type === 'assigned' && drillDown?.value === 'yes'}
                />
                
                <Card 
                  title='מרותקים לעבודה' 
                  value={stats.boundToWorkplace} 
                  icon={<Activity />} 
                  colorClass="bg-slate-100 text-slate-600"
                  subtitle="לא יהיו זמינים ביישוב" 
                  onClick={() => handleDrillDown('bound', 'yes')}
                  isSelected={drillDown?.type === 'bound' && drillDown?.value === 'yes'}
                />
                
                <Card 
                  title='מוכנות תיק אישי' 
                  value={stats.equipmentData.find((d:any) => d.name === 'יש תיק אישי')?.ערך || 0} 
                  icon={<Briefcase />} 
                  colorClass="bg-emerald-50 text-emerald-600"
                  subtitle="אנשי צוות מצוידים" 
                />

                <Card 
                  title='זמינות גבוהה (4-5)' 
                  value={stats.highAvailability} 
                  icon={<HeartPulse />} 
                  colorClass="bg-amber-50 text-amber-600"
                  subtitle={`${((stats.highAvailability / stats.totalPersonnel) * 100).toFixed(1)}% מהצוות זמינים מיידית`} 
                  onClick={() => handleDrillDown('high_availability', 'yes')}
                  isSelected={drillDown?.type === 'high_availability'}
                />

                <Card 
                  title='פער: חוסר בתיק' 
                  value={stats.missingEquipment} 
                  icon={<AlertTriangle />} 
                  colorClass="bg-red-50 text-red-600"
                  subtitle={`${((stats.missingEquipment / stats.totalPersonnel) * 100).toFixed(1)}% ללא ציוד אישי`} 
                  onClick={() => handleDrillDown('missing_equipment', 'yes')}
                  isSelected={drillDown?.type === 'missing_equipment'}
                />
              </div>

              {/* General Charts Section */}
              <div className="grid grid-cols-1 lg:grid-cols-2 gap-8">
                {/* Professions Pie Chart */}
                <div className="bg-white p-6 rounded-2xl shadow-sm border border-slate-100 relative group">
                  <h3 className="text-lg font-bold text-slate-800 mb-6 flex items-center gap-2">
                    <Stethoscope className="text-blue-500" size={20} /> פילוח לפי הכשרה רפואית
                  </h3>
                  <p className="text-xs text-blue-500 absolute top-6 left-6 opacity-0 group-hover:opacity-100 transition-opacity">לחץ על פלח לסינון</p>
                  <div className="h-80 w-full" dir="ltr">
                    <ResponsiveContainer width="100%" height="100%">
                      <PieChart>
                        <Pie
                          data={stats.rolesData}
                          cx="50%" cy="50%"
                          innerRadius={80}
                          outerRadius={120}
                          paddingAngle={5}
                          dataKey="ערך"
                          nameKey="name"
                          label={({name, percent}) => `${name} ${((percent || 0) * 100).toFixed(0)}%`}
                          onClick={(entry) => handleDrillDown('role', entry.name || '')}
                          className="cursor-pointer"
                        >
                          {stats.rolesData.map((_, index) => (
                            <Cell key={`cell-${index}`} fill={COLORS[index % COLORS.length]} />
                          ))}
                        </Pie>
                        <Tooltip />
                        <Legend />
                      </PieChart>
                    </ResponsiveContainer>
                  </div>
                </div>

                {/* Settlements Bar Chart */}
                <div className="bg-white p-6 rounded-2xl shadow-sm border border-slate-100 relative group">
                  <h3 className="text-lg font-bold text-slate-800 mb-6 flex items-center gap-2">
                    <MapPin className="text-emerald-500" size={20} /> צוותים רפואיים לפי יישוב (טופ 10)
                  </h3>
                  <p className="text-xs text-blue-500 absolute top-6 left-6 opacity-0 group-hover:opacity-100 transition-opacity">לחץ על עמודה לסינון</p>
                  <div className="h-80 w-full" dir="ltr">
                    <ResponsiveContainer width="100%" height="100%">
                      <BarChart data={stats.settlementsData} layout="vertical" margin={{ top: 5, right: 30, left: 40, bottom: 5 }}>
                        <CartesianGrid strokeDasharray="3 3" horizontal={false} />
                        <XAxis type="number" />
                        <YAxis dataKey="name" type="category" width={80} tick={{fontSize: 12, fill: '#475569'}} />
                        <Tooltip cursor={{fill: '#f1f5f9'}} />
                        <Bar 
                          dataKey="ערך" 
                          fill="#10b981" 
                          radius={[0, 4, 4, 0]} 
                          barSize={24} 
                          onClick={(data) => handleDrillDown('settlement', data.name || '')}
                          className="cursor-pointer hover:opacity-80"
                        />
                      </BarChart>
                    </ResponsiveContainer>
                  </div>
                </div>
                
                {/* Availability Chart */}
                <div className="bg-white p-6 rounded-2xl shadow-sm border border-slate-100 relative group lg:col-span-2">
                  <h3 className="text-lg font-bold text-slate-800 mb-6 flex items-center gap-2">
                    <AlertCircle className="text-amber-500" size={20} /> רמת זמינות (1 = לא זמין, 5 = זמין מיידית)
                  </h3>
                  <p className="text-xs text-blue-500 absolute top-6 left-6 opacity-0 group-hover:opacity-100 transition-opacity">לחץ על עמודה לסינון</p>
                  <div className="h-64 w-full" dir="ltr">
                    <ResponsiveContainer width="100%" height="100%">
                      <BarChart data={stats.availabilityData}>
                        <CartesianGrid strokeDasharray="3 3" vertical={false} />
                        <XAxis dataKey="name" />
                        <YAxis />
                        <Tooltip cursor={{fill: '#f1f5f9'}} />
                        <Bar 
                          dataKey="ערך" 
                          fill="#f59e0b" 
                          radius={[4, 4, 0, 0]} 
                          barSize={40} 
                          onClick={(data) => handleDrillDown('availability', (data.name || '').replace('רמה ', ''))}
                          className="cursor-pointer hover:opacity-80"
                        />
                      </BarChart>
                    </ResponsiveContainer>
                  </div>
                </div>
              </div>

              {/* NEW SECTION: Operational Readiness & Constraints */}
              <div className="mt-12 space-y-6">
                <div className="flex items-center gap-3 border-b border-slate-200 pb-4">
                  <Siren className="text-red-500" size={28} />
                  <h2 className="text-2xl font-bold text-slate-800">כשירות מבצעית ואילוצים</h2>
                </div>
                
                <div className="grid grid-cols-1 lg:grid-cols-3 gap-8">
                  {/* Specialties Chart */}
                  <div className="bg-white p-6 rounded-2xl shadow-sm border border-slate-100 flex flex-col relative group">
                    <h3 className="text-lg font-bold text-slate-800 mb-6 flex items-center gap-2">
                      <Stethoscope className="text-blue-500" size={20} /> התמחויות קריטיות (רופאים)
                    </h3>
                    <p className="text-xs text-blue-500 absolute top-6 left-6 opacity-0 group-hover:opacity-100 transition-opacity">לחץ לסינון</p>
                    {stats.specData.length > 0 ? (
                      <div className="h-64 w-full flex-grow" dir="ltr">
                        <ResponsiveContainer width="100%" height="100%">
                          <BarChart data={stats.specData} layout="vertical" margin={{ left: 10, right: 20 }}>
                            <CartesianGrid strokeDasharray="3 3" horizontal={false} />
                            <XAxis type="number" />
                            <YAxis dataKey="name" type="category" width={100} tick={{fontSize: 11, fill: '#475569'}} />
                            <Tooltip cursor={{fill: '#f1f5f9'}} />
                            <Bar 
                              dataKey="ערך" 
                              fill="#3b82f6" 
                              radius={[0, 4, 4, 0]} 
                              barSize={20} 
                              onClick={(data) => handleDrillDown('specialty', data.name || '')}
                              className="cursor-pointer hover:opacity-80"
                            />
                          </BarChart>
                        </ResponsiveContainer>
                      </div>
                    ) : (
                      <div className="flex-grow flex items-center justify-center text-slate-400 text-sm">אין מספיק נתונים על התמחויות</div>
                    )}
                  </div>

                  {/* Organizations Chart */}
                  <div className="bg-white p-6 rounded-2xl shadow-sm border border-slate-100 flex flex-col relative group">
                    <h3 className="text-lg font-bold text-slate-800 mb-6 flex items-center gap-2">
                      <Crosshair className="text-emerald-500" size={20} /> שיוך בארגוני הצלה
                    </h3>
                    <p className="text-xs text-blue-500 absolute top-6 left-6 opacity-0 group-hover:opacity-100 transition-opacity">לחץ לסינון</p>
                    {stats.orgsData.length > 0 ? (
                      <div className="h-64 w-full flex-grow" dir="ltr">
                        <ResponsiveContainer width="100%" height="100%">
                          <PieChart>
                            <Pie
                              data={stats.orgsData}
                              cx="50%" cy="50%"
                              innerRadius={60}
                              outerRadius={90}
                              paddingAngle={5}
                              dataKey="ערך"
                              nameKey="name"
                              label={({name, value}) => `${name}: ${value}`}
                              onClick={(entry) => handleDrillDown('organization', entry.name || '')}
                              className="cursor-pointer"
                            >
                              {stats.orgsData.map((_, index) => (
                                <Cell key={`cell-${index}`} fill={['#ef4444', '#f97316', '#10b981', '#8b5cf6'][index % 4]} />
                              ))}
                            </Pie>
                            <Tooltip />
                            <Legend />
                          </PieChart>
                        </ResponsiveContainer>
                      </div>
                    ) : (
                      <div className="flex-grow flex items-center justify-center text-slate-400 text-sm">לא צוין שיוך לארגוני הצלה</div>
                    )}
                  </div>

                  {/* Constraints & Readiness Chart */}
                  <div className="bg-white p-6 rounded-2xl shadow-sm border border-slate-100 flex flex-col justify-between">
                    <div>
                      <h3 className="text-lg font-bold text-slate-800 mb-6 flex items-center gap-2">
                        <AlertTriangle className="text-amber-500" size={20} /> אילוצים מיוחדים
                      </h3>
                      {stats.constData.length > 0 ? (
                        <div className="h-40 w-full" dir="ltr">
                          <ResponsiveContainer width="100%" height="100%">
                            <BarChart data={stats.constData} layout="vertical" margin={{ top: 5, right: 30, left: 40, bottom: 5 }}>
                              <CartesianGrid strokeDasharray="3 3" horizontal={false} />
                              <XAxis type="number" />
                              <YAxis dataKey="name" type="category" width={80} tick={{fontSize: 12, fill: '#475569'}} />
                              <Tooltip cursor={{fill: '#f1f5f9'}} />
                              <Bar 
                                dataKey="ערך" 
                                fill="#f59e0b" 
                                radius={[0, 4, 4, 0]} 
                                barSize={24} 
                                onClick={(data) => handleDrillDown('constraint', data.name || '')}
                                className="cursor-pointer hover:opacity-80"
                              />
                            </BarChart>
                          </ResponsiveContainer>
                        </div>
                      ) : (
                        <div className="h-40 w-full flex items-center justify-center text-slate-400 text-sm">אין אילוצים מתועדים</div>
                      )}
                    </div>
                    
                    <div 
                      onClick={() => handleDrillDown('active_ready', 'פעילים ומוכשרים (ACLS)')}
                      className={`mt-4 pt-4 border-t border-slate-100 flex items-center gap-3 cursor-pointer group transition-colors p-2 -mx-2 rounded-xl ${drillDown?.type === 'active_ready' ? 'bg-blue-50 ring-1 ring-blue-200' : 'hover:bg-slate-50'}`}
                    >
                      <div className="p-2 bg-indigo-50 text-indigo-600 rounded-lg group-hover:bg-indigo-100 transition-colors"><Award size={20} /></div>
                      <div>
                        <p className="text-sm text-slate-500 font-medium group-hover:text-blue-600 transition-colors">כשירות רפואית פעילה (ACLS)</p>
                        <p className="text-xl font-bold text-slate-800">{stats.aclsOrActiveCount} <span className="text-sm font-normal text-slate-500">מוכשרים (לחץ לפירוט)</span></p>
                      </div>
                    </div>
                  </div>
                </div>
              </div>
              
              {/* Data Table with Pagination & Drilldown State */}
              <div id="data-table" className="bg-white rounded-2xl shadow-sm border border-slate-100 overflow-hidden mt-12 scroll-mt-24" data-html2canvas-ignore="true">
                 <div className="p-6 border-b border-slate-100 flex flex-wrap gap-4 justify-between items-center bg-slate-50">
                   <h3 className="text-lg font-bold text-slate-800 flex items-center gap-2">
                     <FileText className="text-slate-500" size={20} /> רשימת כוח אדם <span className="text-xs font-normal text-blue-500 mr-2">(לחץ על רשומה כדי לראות את כרטיס המידע המלא)</span>
                   </h3>
                   
                   {drillDown && (
                     <div className="flex items-center gap-2 bg-blue-600 text-white px-4 py-2 rounded-full text-sm font-bold shadow-sm animate-in zoom-in duration-300">
                       מציג תוצאות עבור: {drillDown.value}
                       <button onClick={() => setDrillDown(null)} className="hover:text-red-200 ml-2 transition-colors"><XCircle size={18} /></button>
                     </div>
                   )}

                   <span className="text-sm bg-white border border-slate-200 text-slate-600 px-4 py-2 rounded-full font-bold shadow-sm">
                     סה"כ {tableData.length} רשומות
                   </span>
                 </div>
                 
                 <div className="overflow-x-auto">
                   <table className="w-full text-right text-sm">
                     <thead className="bg-white text-slate-500 border-b border-slate-100">
                       <tr>
                         <th className="px-6 py-4 font-bold text-slate-700">שם מלא</th>
                         <th className="px-6 py-4 font-bold text-slate-700">טלפון / נייד</th>
                         <th className="px-6 py-4 font-bold text-slate-700">יישוב</th>
                         <th className="px-6 py-4 font-bold text-slate-700">הכשרה רפואית</th>
                         <th className="px-6 py-4 font-bold text-slate-700">זמינות בחירום</th>
                       </tr>
                     </thead>
                     <tbody className="divide-y divide-slate-50 bg-white">
                       {currentTableData.length > 0 ? currentTableData.map((row, idx) => (
                         <tr 
                           key={idx} 
                           onClick={() => setSelectedRowData(row)}
                           className="hover:bg-blue-50 transition-colors cursor-pointer group"
                         >
                           <td className="px-6 py-4 text-slate-800 font-bold group-hover:text-blue-700 transition-colors">
                             {getName(row)}
                           </td>
                           <td className="px-6 py-4 text-blue-600 font-medium" dir="ltr">
                             {getCol(row, ['טלפון', 'נייד', 'phone', 'mobile']) || '-'}
                           </td>
                           <td className="px-6 py-4 text-slate-600 font-medium">{getCol(row, ['יישוב', 'ישוב'])}</td>
                           <td className="px-6 py-4 text-slate-600">
                             {getCol(row, ['הכשרה', 'מקצוע', 'רופא', 'פרמדיק', 'אח', 'חובש', 'מע"ר']) || 'לא הוגדר'}
                           </td>
                           <td className="px-6 py-4">
                             <span className={`inline-flex items-center justify-center w-8 h-8 rounded-full font-bold ${
                               ['4','5'].includes(getCol(row, ['זמינות'])) ? 'bg-emerald-100 text-emerald-700' : 'bg-amber-100 text-amber-700'
                             }`}>
                               {getCol(row, ['זמינות']) || '-'}
                             </span>
                           </td>
                         </tr>
                       )) : (
                         <tr>
                           <td colSpan={5} className="px-6 py-12 text-center text-slate-500 font-medium">לא נמצאו רשומות בחיתוך זה</td>
                         </tr>
                       )}
                     </tbody>
                   </table>
                 </div>
                 
                 {/* Pagination Controls */}
                 {totalPages > 1 && (
                   <div className="flex items-center justify-between px-6 py-4 border-t border-slate-100 bg-white">
                     <button 
                       onClick={() => setCurrentPage(p => Math.max(1, p - 1))}
                       disabled={currentPage === 1}
                       className="flex items-center gap-1 px-4 py-2 rounded-lg font-medium text-slate-600 hover:bg-slate-100 disabled:opacity-50 disabled:cursor-not-allowed transition-colors border border-slate-200"
                     >
                       <ChevronRight size={18} />
                       הקודם
                     </button>
                     
                     <span className="text-sm font-bold text-slate-700">
                       עמוד {currentPage} מתוך {totalPages}
                     </span>
                     
                     <button 
                       onClick={() => setCurrentPage(p => Math.min(totalPages, p + 1))}
                       disabled={currentPage === totalPages}
                       className="flex items-center gap-1 px-4 py-2 rounded-lg font-medium text-slate-600 hover:bg-slate-100 disabled:opacity-50 disabled:cursor-not-allowed transition-colors border border-slate-200"
                     >
                       הבא
                       <ChevronLeft size={18} />
                     </button>
                   </div>
                 )}
              </div>
            </>
          )}

        </div>
      )}

      {/* Row Detail Modal (Smart View) */}
      {selectedRowData && (
        <div 
          className="fixed inset-0 z-[100] flex items-center justify-center bg-slate-900/60 p-4 backdrop-blur-sm transition-opacity" 
          onClick={() => setSelectedRowData(null)}
          dir="rtl"
        >
          <div 
            className="bg-white rounded-2xl shadow-2xl w-full max-w-2xl max-h-[85vh] overflow-hidden flex flex-col animate-in zoom-in-95 duration-200"
            onClick={e => e.stopPropagation()}
          >
            <div className="p-6 border-b border-slate-100 flex justify-between items-center bg-slate-50">
              <h2 className="text-xl font-bold text-slate-800 flex items-center gap-3">
                <div className="p-2 bg-blue-100 text-blue-600 rounded-full"><User size={24} /></div>
                כרטיס מידע אישי: {getName(selectedRowData)}
              </h2>
              <button 
                onClick={() => setSelectedRowData(null)} 
                className="p-2 hover:bg-slate-200 rounded-full text-slate-500 transition-colors"
              >
                <X size={20} />
              </button>
            </div>
            
            <div className="p-6 overflow-y-auto">
              <div className="grid grid-cols-1 md:grid-cols-2 gap-x-8 gap-y-6">
                {Object.entries(selectedRowData).map(([key, value]) => {
                  if (value === '' || value === null || value === undefined) return null;
                  
                  const cleanKey = key.trim().toLowerCase();
                  
                  // Filter out technical fields robustly
                  if (
                    cleanKey === 'מס' || 
                    cleanKey === "מס'" || 
                    cleanKey === 'מספר' || 
                    cleanKey.includes('crm') || 
                    cleanKey.includes('פנייה') ||
                    cleanKey === 'תאריך' || 
                    cleanKey.includes('טופל')
                  ) return null;

                  // Handle raw Excel date numbers (e.g., 44256)
                  let displayValue = String(value);
                  if (typeof value === 'number' && value > 30000 && value < 60000 && cleanKey.includes('תאריך')) {
                    // Excel epoch starts 1900-01-01
                    const date = new Date((value - 25569) * 86400 * 1000);
                    displayValue = date.toLocaleDateString('he-IL');
                  }

                  // Make headers look nicer
                  const isLongText = displayValue.length > 50;
                  
                  return (
                    <div key={key} className={`border-b border-slate-100 pb-3 ${isLongText ? 'md:col-span-2' : ''}`}>
                      <p className="text-xs font-bold text-slate-400 mb-1">{key}</p>
                      <p className="text-sm text-slate-800 font-medium" dir="auto">{displayValue}</p>
                    </div>
                  );
                })}
              </div>
            </div>
            
            <div className="p-4 bg-slate-50 border-t border-slate-100 flex justify-end gap-3">
              <button 
                onClick={() => setSelectedRowData(null)}
                className="bg-slate-800 hover:bg-slate-900 text-white px-6 py-2 rounded-lg font-bold transition-colors shadow-sm"
              >
                סגור כרטיס
              </button>
            </div>
          </div>
        </div>
      )}
      
    </div>
  );
}

export default App;