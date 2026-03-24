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
  ChevronRight, ChevronLeft, AlertTriangle, Siren, Crosshair, Award, Download
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
        const jsonData = XLSX.utils.sheet_to_json(ws, { defval: '' });
        setData(jsonData);
        
        // Reset filters on new file
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

  const getRoleCategory = (rawRole: string) => {
    if (!rawRole) return 'אחר/לא הוגדר';
    if (rawRole.includes('רופא')) return 'רופאים';
    if (rawRole.includes('פרמדיק') || rawRole.includes('פראמדיק')) return 'פרמדיקים';
    if (rawRole.includes('אח ') || rawRole.includes('אחות')) return 'אחים/אחיות';
    if (rawRole.includes('חובש') || rawRole.includes('מע"ר') || rawRole.includes('עזרה ראשונה')) return 'חובשים/מע"רים';
    return 'אחר/לא הוגדר';
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

  // Apply filters to data
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

  // Reset pagination when filters change
  useEffect(() => {
    setCurrentPage(1);
  }, [filteredData]);

  // Compute stats based on FILTERED data
  const stats = useMemo(() => {
    if (!filteredData.length && !data.length) return null;

    let totalPersonnel = filteredData.length;
    let settlementsCount: Record<string, number> = {};
    let availabilityScores: Record<string, number> = { '1':0, '2':0, '3':0, '4':0, '5':0 };
    let rolesCount: Record<string, number> = { 'רופאים': 0, 'פרמדיקים': 0, 'אחים/אחיות': 0, 'חובשים/מע"רים': 0, 'אחר/לא הוגדר': 0 };
    
    let assignedToEmergency = 0;
    let boundToWorkplace = 0;
    let highAvailability = 0;
    let missingEquipment = 0;

    let equipmentStatus = { 'יש תיק אישי': 0, 'אין תיק': 0 };
    
    // New metrics
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

      // Constraints
      const constraints = getCol(row, ['אילוצ', 'כוננות', 'מילואים']);
      if (constraints.includes('כוננות') || rowString.includes('כוננות')) constraintsCount['כיתת כוננות']++;
      if (constraints.includes('מילואים') || rowString.includes('מילואים')) constraintsCount['מילואים']++;
      
      // Specialties - Fuzzy match over whole row string for max robustness
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

      // Organizations - Fuzzy match for max robustness
      if (['פרמדיקים', 'חובשים/מע"רים'].includes(roleCat)) {
         if (rowString.includes('מד"א') || rowString.includes('מגן דוד')) orgsCount['מד"א']++;
         else if (rowString.includes('איחוד')) orgsCount['איחוד הצלה']++;
         else if (rowString.includes('צה"ל') || rowString.includes('צבא')) orgsCount['צה"ל']++;
         else orgsCount['אחר / לא שויך']++; 
      }

      // ACLS / Active - Fuzzy
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

  const Card = ({ title, value, icon, subtitle, colorClass = "bg-blue-50 text-blue-600" }: any) => (
    <div className="bg-white rounded-xl shadow-sm border border-slate-100 p-6 flex flex-col items-start">
      <div className="flex items-center justify-between w-full mb-4">
        <h3 className="text-slate-500 font-medium">{title}</h3>
        <div className={`p-2 rounded-lg ${colorClass}`}>{icon}</div>
      </div>
      <p className="text-3xl font-bold text-slate-800">{value}</p>
      {subtitle && <p className="text-sm text-slate-400 mt-2">{subtitle}</p>}
    </div>
  );

  // Pagination Logic
  const totalPages = Math.max(1, Math.ceil(filteredData.length / rowsPerPage));
  const currentTableData = filteredData.slice((currentPage - 1) * rowsPerPage, currentPage * rowsPerPage);

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

             {(filterSettlement !== 'all' || filterRole !== 'all' || filterAvailability !== 'all' || filterAssigned !== 'all') && (
               <button 
                 onClick={clearFilters}
                 className="flex items-center gap-1 text-sm text-red-500 hover:text-red-700 font-medium mr-auto bg-red-50 px-3 py-2 rounded-lg transition-colors"
               >
                 <XCircle size={16} /> נקה סינונים
               </button>
             )}
          </div>

          {/* Alert if no data after filter */}
          {filteredData.length === 0 ? (
            <div className="bg-amber-50 text-amber-800 p-8 rounded-xl text-center border border-amber-200 font-medium">
              לא נמצאו תוצאות התואמות לסינון שבחרת. אנא שנה את הפילטרים.
            </div>
          ) : (
            <>
              {/* Top Stats Row */}
              <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 xl:grid-cols-6 gap-6">
                <Card title='סה"כ כוח אדם' value={stats.totalPersonnel} icon={<Users />} subtitle={`מתוך ${data.length} בסה"כ`} />
                
                <Card 
                  title='שובצו למכלולים' 
                  value={stats.assignedToEmergency} 
                  icon={<ShieldAlert />} 
                  colorClass="bg-indigo-50 text-indigo-600"
                  subtitle={`${((stats.assignedToEmergency / stats.totalPersonnel) * 100).toFixed(1)}% מסך הצוות המוצג`} 
                />
                
                <Card 
                  title='מרותקים לעבודה' 
                  value={stats.boundToWorkplace} 
                  icon={<Activity />} 
                  colorClass="bg-slate-100 text-slate-600"
                  subtitle="לא יהיו זמינים ביישוב" 
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
                />

                <Card 
                  title='פער: חוסר בתיק' 
                  value={stats.missingEquipment} 
                  icon={<AlertTriangle />} 
                  colorClass="bg-red-50 text-red-600"
                  subtitle={`${((stats.missingEquipment / stats.totalPersonnel) * 100).toFixed(1)}% ללא ציוד אישי`} 
                />
              </div>

              {/* General Charts Section */}
              <div className="grid grid-cols-1 lg:grid-cols-2 gap-8">
                {/* Professions Pie Chart */}
                <div className="bg-white p-6 rounded-2xl shadow-sm border border-slate-100">
                  <h3 className="text-lg font-bold text-slate-800 mb-6 flex items-center gap-2">
                    <Stethoscope className="text-blue-500" size={20} /> פילוח לפי הכשרה רפואית
                  </h3>
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
                <div className="bg-white p-6 rounded-2xl shadow-sm border border-slate-100">
                  <h3 className="text-lg font-bold text-slate-800 mb-6 flex items-center gap-2">
                    <MapPin className="text-emerald-500" size={20} /> צוותים רפואיים לפי יישוב (טופ 10)
                  </h3>
                  <div className="h-80 w-full" dir="ltr">
                    <ResponsiveContainer width="100%" height="100%">
                      <BarChart data={stats.settlementsData} layout="vertical" margin={{ top: 5, right: 30, left: 40, bottom: 5 }}>
                        <CartesianGrid strokeDasharray="3 3" horizontal={false} />
                        <XAxis type="number" />
                        <YAxis dataKey="name" type="category" width={80} tick={{fontSize: 12, fill: '#475569'}} />
                        <Tooltip cursor={{fill: '#f1f5f9'}} />
                        <Bar dataKey="ערך" fill="#10b981" radius={[0, 4, 4, 0]} barSize={24} />
                      </BarChart>
                    </ResponsiveContainer>
                  </div>
                </div>
                
                {/* Availability Chart */}
                <div className="bg-white p-6 rounded-2xl shadow-sm border border-slate-100 lg:col-span-2">
                  <h3 className="text-lg font-bold text-slate-800 mb-6 flex items-center gap-2">
                    <AlertCircle className="text-amber-500" size={20} /> רמת זמינות (1 = לא זמין, 5 = זמין מיידית)
                  </h3>
                  <div className="h-64 w-full" dir="ltr">
                    <ResponsiveContainer width="100%" height="100%">
                      <BarChart data={stats.availabilityData}>
                        <CartesianGrid strokeDasharray="3 3" vertical={false} />
                        <XAxis dataKey="name" />
                        <YAxis />
                        <Tooltip cursor={{fill: '#f1f5f9'}} />
                        <Bar dataKey="ערך" fill="#f59e0b" radius={[4, 4, 0, 0]} barSize={40} />
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
                  <div className="bg-white p-6 rounded-2xl shadow-sm border border-slate-100 flex flex-col">
                    <h3 className="text-lg font-bold text-slate-800 mb-6 flex items-center gap-2">
                      <Stethoscope className="text-blue-500" size={20} /> התמחויות קריטיות (רופאים)
                    </h3>
                    {stats.specData.length > 0 ? (
                      <div className="h-64 w-full flex-grow" dir="ltr">
                        <ResponsiveContainer width="100%" height="100%">
                          <BarChart data={stats.specData} layout="vertical" margin={{ left: 10, right: 20 }}>
                            <CartesianGrid strokeDasharray="3 3" horizontal={false} />
                            <XAxis type="number" />
                            <YAxis dataKey="name" type="category" width={100} tick={{fontSize: 11, fill: '#475569'}} />
                            <Tooltip cursor={{fill: '#f1f5f9'}} />
                            <Bar dataKey="ערך" fill="#3b82f6" radius={[0, 4, 4, 0]} barSize={20} />
                          </BarChart>
                        </ResponsiveContainer>
                      </div>
                    ) : (
                      <div className="flex-grow flex items-center justify-center text-slate-400 text-sm">אין מספיק נתונים על התמחויות</div>
                    )}
                  </div>

                  {/* Organizations Chart */}
                  <div className="bg-white p-6 rounded-2xl shadow-sm border border-slate-100 flex flex-col">
                    <h3 className="text-lg font-bold text-slate-800 mb-6 flex items-center gap-2">
                      <Crosshair className="text-emerald-500" size={20} /> שיוך מבצעי בארגוני הצלה
                    </h3>
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
                              <Bar dataKey="ערך" fill="#f59e0b" radius={[0, 4, 4, 0]} barSize={24} />
                            </BarChart>
                          </ResponsiveContainer>
                        </div>
                      ) : (
                        <div className="h-40 w-full flex items-center justify-center text-slate-400 text-sm">אין אילוצים מתועדים</div>
                      )}
                    </div>
                    
                    <div className="mt-4 pt-4 border-t border-slate-100 flex items-center gap-3">
                      <div className="p-2 bg-indigo-50 text-indigo-600 rounded-lg"><Award size={20} /></div>
                      <div>
                        <p className="text-sm text-slate-500 font-medium">כשירות רפואית פעילה (ACLS / פעילים)</p>
                        <p className="text-xl font-bold text-slate-800">{stats.aclsOrActiveCount} <span className="text-sm font-normal text-slate-500">אנשי צוות מוכשרים</span></p>
                      </div>
                    </div>
                  </div>
                </div>
              </div>
              
              {/* Data Table with Pagination */}
              <div className="bg-white rounded-2xl shadow-sm border border-slate-100 overflow-hidden mt-12" data-html2canvas-ignore="true">
                 <div className="p-6 border-b border-slate-100 flex justify-between items-center">
                   <h3 className="text-lg font-bold text-slate-800 flex items-center gap-2">
                     <FileText className="text-slate-500" size={20} /> נתונים גולמיים
                   </h3>
                   <span className="text-sm bg-slate-100 text-slate-600 px-3 py-1 rounded-full font-medium">
                     סה"כ {filteredData.length} רשומות
                   </span>
                 </div>
                 
                 <div className="overflow-x-auto">
                   <table className="w-full text-right text-sm">
                     <thead className="bg-slate-50 text-slate-500">
                       <tr>
                         <th className="px-6 py-4 font-medium">שם / זיהוי (חלקי)</th>
                         <th className="px-6 py-4 font-medium">יישוב</th>
                         <th className="px-6 py-4 font-medium">הכשרה רפואית</th>
                         <th className="px-6 py-4 font-medium">זמינות בחירום</th>
                         <th className="px-6 py-4 font-medium">מרותק למקום עבודה</th>
                       </tr>
                     </thead>
                     <tbody className="divide-y divide-slate-100">
                       {currentTableData.map((row, idx) => (
                         <tr key={idx} className="hover:bg-slate-50 transition-colors">
                           <td className="px-6 py-4 text-slate-800 font-medium">
                             {getCol(row, ['שם', 'זיהוי']) || `רשומה #${(currentPage - 1) * rowsPerPage + idx + 1}`}
                           </td>
                           <td className="px-6 py-4 text-slate-600">{getCol(row, ['יישוב', 'ישוב'])}</td>
                           <td className="px-6 py-4 text-slate-600">{getCol(row, ['הכשרה', 'מקצוע', 'רופא', 'פרמדיק', 'אח', 'חובש', 'מע"ר'])}</td>
                           <td className="px-6 py-4">
                             <span className="inline-flex items-center justify-center bg-amber-100 text-amber-700 w-8 h-8 rounded-full font-bold">
                               {getCol(row, ['זמינות']) || '-'}
                             </span>
                           </td>
                           <td className="px-6 py-4 text-slate-600">
                             {getCol(row, ['מרותק'])?.includes('כן') || getCol(row, ['מרותק']) === '1' ? 
                               <span className="text-red-500 font-medium">כן</span> : 
                               <span className="text-emerald-500 font-medium">לא</span>
                             }
                           </td>
                         </tr>
                       ))}
                     </tbody>
                   </table>
                 </div>
                 
                 {/* Pagination Controls */}
                 {totalPages > 1 && (
                   <div className="flex items-center justify-between px-6 py-4 border-t border-slate-100 bg-slate-50">
                     <button 
                       onClick={() => setCurrentPage(p => Math.max(1, p - 1))}
                       disabled={currentPage === 1}
                       className="flex items-center gap-1 px-4 py-2 rounded-lg font-medium text-slate-600 hover:bg-slate-200 disabled:opacity-50 disabled:cursor-not-allowed transition-colors"
                     >
                       <ChevronRight size={18} />
                       הקודם
                     </button>
                     
                     <span className="text-sm font-medium text-slate-600">
                       עמוד {currentPage} מתוך {totalPages}
                     </span>
                     
                     <button 
                       onClick={() => setCurrentPage(p => Math.min(totalPages, p + 1))}
                       disabled={currentPage === totalPages}
                       className="flex items-center gap-1 px-4 py-2 rounded-lg font-medium text-slate-600 hover:bg-slate-200 disabled:opacity-50 disabled:cursor-not-allowed transition-colors"
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
    </div>
  );
}

export default App;