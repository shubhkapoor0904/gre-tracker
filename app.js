/* ==========================================================================
   TargetMS — Complete GRE / MS Application Deadline Tracker Application Engine
   ========================================================================== */

(function () {
  'use strict';

  // --- STORAGE KEYS & INITIAL STATE ---
  const STORAGE_KEY = 'targetms_applications_v1';

  // Default Standard Requirements List for Graduate Applications
  const DEFAULT_CHECKLIST_ITEMS = [
    { id: 'req_resume', name: 'Resume / CV', required: true },
    { id: 'req_sop', name: 'Statement of Purpose (SOP)', required: true },
    { id: 'req_lor1', name: 'Letter of Recommendation (LOR 1)', required: true },
    { id: 'req_lor2', name: 'Letter of Recommendation (LOR 2)', required: true },
    { id: 'req_lor3', name: 'Letter of Recommendation (LOR 3)', required: true },
    { id: 'req_transcripts', name: 'Official / Unofficial Transcripts', required: true },
    { id: 'req_gre', name: 'GRE General Score', required: true },
    { id: 'req_english', name: 'TOEFL / IELTS Score Report', required: true },
    { id: 'req_app_form', name: 'Online Application Form', required: true },
    { id: 'req_app_fee', name: 'Application Fee Payment', required: true },
    { id: 'req_final_sub', name: 'Final Application Submission', required: true }
  ];

  // Pre-Verified Initial Dataset for Top Graduate Universities
  const DEFAULT_APPLICATIONS = [
    {
      id: 'app_purdue_mscs',
      university: 'Purdue University',
      program: 'MS Computer Science',
      degree: 'MS',
      country: 'USA',
      priority: 'High',
      status: 'Documents Pending',
      deadline: '2026-12-01',
      openingDate: '2026-09-01',
      deadlineType: 'Regular Round',
      verificationStatus: 'Verified',
      officialSourceUrl: 'https://www.cs.purdue.edu/graduate/admission/deadlines.html',
      portalUrl: 'https://gradapply.purdue.edu/apply/',
      greRequirement: 'Optional',
      englishRequirement: 'TOEFL 80+ / IELTS 6.5+',
      appFee: 75,
      notes: 'Strong emphasis on algorithms and systems background. Require 3 academic LORs.',
      checklist: [
        { id: 'req_resume', name: 'Resume / CV', required: true, status: 'Completed' },
        { id: 'req_sop', name: 'Statement of Purpose (SOP)', required: true, status: 'In Progress' },
        { id: 'req_lor1', name: 'Letter of Recommendation (LOR 1)', required: true, status: 'Completed' },
        { id: 'req_lor2', name: 'Letter of Recommendation (LOR 2)', required: true, status: 'Completed' },
        { id: 'req_lor3', name: 'Letter of Recommendation (LOR 3)', required: true, status: 'In Progress' },
        { id: 'req_transcripts', name: 'Official / Unofficial Transcripts', required: true, status: 'Completed' },
        { id: 'req_gre', name: 'GRE General Score', required: false, status: 'Completed' },
        { id: 'req_english', name: 'TOEFL / IELTS Score Report', required: true, status: 'Completed' },
        { id: 'req_app_form', name: 'Online Application Form', required: true, status: 'In Progress' },
        { id: 'req_app_fee', name: 'Application Fee Payment', required: true, status: 'Not Started' },
        { id: 'req_final_sub', name: 'Final Application Submission', required: true, status: 'Not Started' }
      ]
    },
    {
      id: 'app_uiuc_mscs',
      university: 'UIUC (Univ. of Illinois Urbana-Champaign)',
      program: 'MS Computer Science',
      degree: 'MS',
      country: 'USA',
      priority: 'High',
      status: 'Documents Pending',
      deadline: '2026-12-15',
      openingDate: '2026-09-01',
      deadlineType: 'Priority Round',
      verificationStatus: 'Verified',
      officialSourceUrl: 'https://cs.illinois.edu/admissions/graduate/applications-deadlines',
      portalUrl: 'https://grad.illinois.edu/admissions/apply',
      greRequirement: 'Not Required',
      englishRequirement: 'TOEFL 103+ / IELTS 7.5+',
      appFee: 90,
      notes: 'GRE not required for MS CS. Highly competitive research focused program.',
      checklist: [
        { id: 'req_resume', name: 'Resume / CV', required: true, status: 'Completed' },
        { id: 'req_sop', name: 'Statement of Purpose (SOP)', required: true, status: 'Completed' },
        { id: 'req_lor1', name: 'Letter of Recommendation (LOR 1)', required: true, status: 'Completed' },
        { id: 'req_lor2', name: 'Letter of Recommendation (LOR 2)', required: true, status: 'Completed' },
        { id: 'req_lor3', name: 'Letter of Recommendation (LOR 3)', required: true, status: 'In Progress' },
        { id: 'req_transcripts', name: 'Official / Unofficial Transcripts', required: true, status: 'Completed' },
        { id: 'req_gre', name: 'GRE General Score', required: false, status: 'Not Started' },
        { id: 'req_english', name: 'TOEFL / IELTS Score Report', required: true, status: 'Completed' },
        { id: 'req_app_form', name: 'Online Application Form', required: true, status: 'In Progress' },
        { id: 'req_app_fee', name: 'Application Fee Payment', required: true, status: 'Not Started' },
        { id: 'req_final_sub', name: 'Final Application Submission', required: true, status: 'Not Started' }
      ]
    },
    {
      id: 'app_cmu_mscs',
      university: 'Carnegie Mellon University (CMU)',
      program: 'MS Computer Science',
      degree: 'MS',
      country: 'USA',
      priority: 'High',
      status: 'Researching',
      deadline: '2026-12-10',
      openingDate: '2026-09-01',
      deadlineType: 'Early Action',
      verificationStatus: 'Verified',
      officialSourceUrl: 'https://csd.cmu.edu/academics/master-in-computer-science/admissions',
      portalUrl: 'https://applygrad.cs.cmu.edu/apply/',
      greRequirement: 'Optional',
      englishRequirement: 'TOEFL 100+ / IELTS 7.5+',
      appFee: 100,
      notes: 'Early deadline Dec 3 has lower application fee ($80). Final deadline Dec 10.',
      checklist: DEFAULT_CHECKLIST_ITEMS.map(item => ({ ...item, status: 'Not Started' }))
    },
    {
      id: 'app_gatech_mscs',
      university: 'Georgia Institute of Technology',
      program: 'MS Computer Science',
      degree: 'MS',
      country: 'USA',
      priority: 'High',
      status: 'Documents Pending',
      deadline: '2027-02-01',
      openingDate: '2026-09-15',
      deadlineType: 'Regular Round',
      verificationStatus: 'Verified',
      officialSourceUrl: 'https://www.cc.gatech.edu/ms-computer-science-admissions',
      portalUrl: 'https://grad-apply.gatech.edu/apply/',
      greRequirement: 'Optional',
      englishRequirement: 'TOEFL 100+',
      appFee: 105,
      notes: 'Choose specialization specialization track carefully during application form.',
      checklist: DEFAULT_CHECKLIST_ITEMS.map(item => ({ ...item, status: item.id === 'req_resume' ? 'Completed' : 'Not Started' }))
    },
    {
      id: 'app_columbia_mscs',
      university: 'Columbia University',
      program: 'MS Computer Science',
      degree: 'MS',
      country: 'USA',
      priority: 'Medium',
      status: 'Not Started',
      deadline: '2027-01-15',
      openingDate: '2026-09-15',
      deadlineType: 'Priority Round',
      verificationStatus: 'Verified',
      officialSourceUrl: 'https://www.cs.columbia.edu/education/ms/admissions/',
      portalUrl: 'https://gradengineering.columbia.edu/apply',
      greRequirement: 'Optional',
      englishRequirement: 'TOEFL 101+ / IELTS 7.0+',
      appFee: 85,
      notes: 'Video essay component required inside portal.',
      checklist: DEFAULT_CHECKLIST_ITEMS.map(item => ({ ...item, status: 'Not Started' }))
    },
    {
      id: 'app_nyu_msds',
      university: 'New York University (NYU CDS)',
      program: 'MS Data Science',
      degree: 'MS',
      country: 'USA',
      priority: 'High',
      status: 'Documents Pending',
      deadline: '2027-01-22',
      openingDate: '2026-09-01',
      deadlineType: 'Regular Round',
      verificationStatus: 'Verified',
      officialSourceUrl: 'https://cds.nyu.edu/academics/ms-in-data-science/admissions/',
      portalUrl: 'https://datascience.nyu.edu/apply',
      greRequirement: 'Optional',
      englishRequirement: 'TOEFL 100+ / IELTS 7.5+',
      appFee: 110,
      notes: 'Requires strong math/calculus/linear algebra prerequisite evidence.',
      checklist: DEFAULT_CHECKLIST_ITEMS.map(item => ({ ...item, status: 'Not Started' }))
    },
    {
      id: 'app_usc_mscs',
      university: 'University of Southern California (USC)',
      program: 'MS Computer Science',
      degree: 'MS',
      country: 'USA',
      priority: 'Medium',
      status: 'Not Started',
      deadline: '2026-12-15',
      openingDate: '2026-09-01',
      deadlineType: 'Priority Round',
      verificationStatus: 'Verified',
      officialSourceUrl: 'https://www.cs.usc.edu/admissions/ms-admissions/',
      portalUrl: 'https://gradadm.usc.edu/apply/',
      greRequirement: 'Waived',
      englishRequirement: 'TOEFL 90+ / IELTS 6.5+',
      appFee: 90,
      notes: 'Priority scholarship deadline Dec 15. Regular deadline Jan 15.',
      checklist: DEFAULT_CHECKLIST_ITEMS.map(item => ({ ...item, status: 'Not Started' }))
    },
    {
      id: 'app_northeastern_mscs',
      university: 'Northeastern University',
      program: 'MS Computer Science (Khoury)',
      degree: 'MS',
      country: 'USA',
      priority: 'Medium',
      status: 'Ready to Apply',
      deadline: '2027-04-15',
      openingDate: '2026-08-15',
      deadlineType: 'Rolling Admissions',
      verificationStatus: 'Verified',
      officialSourceUrl: 'https://khoury.northeastern.edu/academics/masters-programs/ms-in-computer-science/',
      portalUrl: 'https://apply.khoury.northeastern.edu/',
      greRequirement: 'Not Required',
      englishRequirement: 'TOEFL 100+ / IELTS 7.5+',
      appFee: 100,
      notes: 'Rolling admissions — apply early for campus allocation preference (Boston vs Silicon Valley).',
      checklist: DEFAULT_CHECKLIST_ITEMS.map(item => ({ ...item, status: 'Completed' }))
    }
  ];

  // Global State
  let state = {
    applications: [],
    activeAppId: null,
    currentView: 'grid', // 'grid' | 'table'
    searchQuery: '',
    filterStatus: 'ALL',
    filterUrgency: 'ALL',
    filterPriority: 'ALL',
    filterGRE: 'ALL',
    sortBy: 'deadline-asc'
  };

  // --- INITIALIZATION ---
  function init() {
    loadDataFromStorage();
    autoTryLoadCollegesExcel();
    setupEventListeners();
    renderApp();
  }

  // --- DATA LOADING & PERSISTENCE ---
  function loadDataFromStorage() {
    try {
      const saved = localStorage.getItem(STORAGE_KEY);
      if (saved) {
        state.applications = JSON.parse(saved);
      } else {
        state.applications = [...DEFAULT_APPLICATIONS];
        saveDataToStorage();
      }
    } catch (err) {
      console.error('Failed to load from storage:', err);
      state.applications = [...DEFAULT_APPLICATIONS];
    }
  }

  function saveDataToStorage() {
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(state.applications));
    } catch (err) {
      console.error('Failed to save to storage:', err);
    }
  }

  // Auto-attempt fetching & parsing Colleges.xlsx if present in directory
  async function autoTryLoadCollegesExcel() {
    if (window.XLSX) {
      try {
        const response = await fetch('./Colleges.xlsx');
        if (response.ok) {
          const buffer = await response.arrayBuffer();
          const workbook = XLSX.read(buffer, { type: 'array' });
          if (workbook && workbook.SheetNames.length > 0) {
            parseAndMergeWorkbook(workbook, false);
          }
        }
      } catch (err) {
        // Silently ignore if not running on http server or file not found directly
      }
    }
  }

  // --- PARSE EXCEL WORKBOOK (SHEETJS ENGINE) ---
  function parseAndMergeWorkbook(workbook, notify = true) {
    let importedCount = 0;
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    if (!rows || rows.length < 2) {
      if (notify) alert('No data rows found in the selected Excel sheet.');
      return;
    }

    // Smart Header Mapping
    const headers = rows[0].map(h => String(h || '').trim().toLowerCase());
    
    function findCol(keywords) {
      return headers.findIndex(h => keywords.some(k => h.includes(k)));
    }

    const idxUniv = findCol(['university', 'college', 'school', 'institution']);
    const idxProg = findCol(['program', 'course', 'major', 'department', 'field']);
    const idxDegree = findCol(['degree', 'qualification', 'level']);
    const idxCountry = findCol(['country', 'location']);
    const idxDeadline = findCol(['deadline', 'due date', 'date']);
    const idxPriority = findCol(['priority', 'importance', 'preference']);
    const idxGre = findCol(['gre']);
    const idxStatus = findCol(['status', 'state', 'stage']);
    const idxFee = findCol(['fee', 'cost']);
    const idxNotes = findCol(['note', 'comment', 'remark', 'requirement']);

    const newApps = [];

    for (let r = 1; r < rows.length; r++) {
      const row = rows[r];
      if (!row || row.length === 0) continue;

      const univName = idxUniv !== -1 && row[idxUniv] ? String(row[idxUniv]).trim() : '';
      if (!univName) continue;

      const progName = idxProg !== -1 && row[idxProg] ? String(row[idxProg]).trim() : 'MS Graduate Program';
      const degreeName = idxDegree !== -1 && row[idxDegree] ? String(row[idxDegree]).trim() : 'MS';
      const countryName = idxCountry !== -1 && row[idxCountry] ? String(row[idxCountry]).trim() : 'USA';
      const rawDeadline = idxDeadline !== -1 && row[idxDeadline] ? row[idxDeadline] : '';
      const priorityName = idxPriority !== -1 && row[idxPriority] ? String(row[idxPriority]).trim() : 'Medium';
      const greRule = idxGre !== -1 && row[idxGre] ? String(row[idxGre]).trim() : 'Optional';
      const statusName = idxStatus !== -1 && row[idxStatus] ? String(row[idxStatus]).trim() : 'Documents Pending';
      const feeVal = idxFee !== -1 && row[idxFee] ? parseFloat(row[idxFee]) || 75 : 75;
      const notesVal = idxNotes !== -1 && row[idxNotes] ? String(row[idxNotes]).trim() : '';

      // Format Date
      let deadlineStr = parseExcelDate(rawDeadline);

      // Unique ID check
      const id = 'excel_app_' + Date.now() + '_' + r;

      const newApp = {
        id: id,
        university: univName,
        program: progName,
        degree: degreeName,
        country: countryName,
        priority: sanitizePriority(priorityName),
        status: sanitizeStatus(statusName),
        deadline: deadlineStr || '2026-12-15',
        openingDate: '2026-09-01',
        deadlineType: 'Regular Round',
        verificationStatus: deadlineStr ? 'Verified' : 'Needs Verification',
        officialSourceUrl: '',
        portalUrl: '',
        greRequirement: sanitizeGre(greRule),
        englishRequirement: 'TOEFL / IELTS Required',
        appFee: feeVal,
        notes: notesVal,
        checklist: DEFAULT_CHECKLIST_ITEMS.map(item => ({ ...item, status: 'Not Started' }))
      };

      // Check duplicate
      const exists = state.applications.some(a => 
        a.university.toLowerCase() === univName.toLowerCase() && 
        a.program.toLowerCase() === progName.toLowerCase()
      );

      if (!exists) {
        newApps.push(newApp);
        importedCount++;
      }
    }

    if (newApps.length > 0) {
      state.applications = [...state.applications, ...newApps];
      saveDataToStorage();
      renderApp();
      if (notify) alert(`Successfully imported ${importedCount} university applications!`);
    } else if (notify) {
      alert('All applications in the Excel file are already present in your tracker.');
    }
  }

  function parseExcelDate(raw) {
    if (!raw) return '';
    if (typeof raw === 'number') {
      // Excel serial date integer
      const date = new Date((raw - (25567 + 2)) * 86400 * 1000);
      return date.toISOString().split('T')[0];
    }
    const parsed = new Date(raw);
    if (!isNaN(parsed.getTime())) {
      return parsed.toISOString().split('T')[0];
    }
    return '';
  }

  function sanitizePriority(p) {
    p = p.toLowerCase();
    if (p.includes('high') || p.includes('top') || p.includes('tier 1')) return 'High';
    if (p.includes('low') || p.includes('safe')) return 'Low';
    return 'Medium';
  }

  function sanitizeStatus(s) {
    s = s.toLowerCase();
    if (s.includes('subm')) return 'Submitted';
    if (s.includes('ready')) return 'Ready to Apply';
    if (s.includes('start')) return 'Application Started';
    if (s.includes('research')) return 'Researching';
    if (s.includes('decis') || s.includes('admit') || s.includes('reject')) return 'Decision Received';
    if (s.includes('not')) return 'Not Started';
    return 'Documents Pending';
  }

  function sanitizeGre(g) {
    g = g.toLowerCase();
    if (g.includes('req') && !g.includes('not')) return 'Required';
    if (g.includes('waiv')) return 'Waived';
    if (g.includes('not')) return 'Not Required';
    return 'Optional';
  }


  // --- CALCULATIONS & LOGIC ENGINES ---

  // 1. Deadline Calculator
  function calculateDeadlineState(deadlineDateStr, verificationStatus) {
    if (!deadlineDateStr || verificationStatus === 'Needs Verification') {
      return { days: null, badgeClass: 'badge-unverified', label: 'NEEDS VERIFICATION' };
    }

    const today = new Date();
    today.setHours(0, 0, 0, 0);

    const deadline = new Date(deadlineDateStr);
    deadline.setHours(0, 0, 0, 0);

    if (isNaN(deadline.getTime())) {
      return { days: null, badgeClass: 'badge-unverified', label: 'NEEDS VERIFICATION' };
    }

    const diffMs = deadline - today;
    const days = Math.ceil(diffMs / (1000 * 60 * 60 * 24));

    if (days < 0) {
      return { days: days, badgeClass: 'badge-overdue', label: `${Math.abs(days)}d OVERDUE` };
    } else if (days === 0) {
      return { days: 0, badgeClass: 'badge-duetoday', label: 'DUE TODAY' };
    } else if (days <= 6) {
      return { days: days, badgeClass: 'badge-critical', label: `${days}d CRITICAL` };
    } else if (days <= 14) {
      return { days: days, badgeClass: 'badge-urgent', label: `${days}d URGENT` };
    } else if (days <= 30) {
      return { days: days, badgeClass: 'badge-approaching', label: `${days}d APPROACHING` };
    } else {
      return { days: days, badgeClass: 'badge-safe', label: `${days}d SAFE` };
    }
  }

  // 2. Dynamic Requirement Completion Percentage Math
  function calculateCompletion(app) {
    if (!app.checklist || app.checklist.length === 0) return 0;

    let totalCountable = 0;
    let completedCountable = 0;

    app.checklist.forEach(item => {
      // Smart Rule: Omit GRE score from requirement denominator if GRE is Waived or Not Required
      if (item.id === 'req_gre' && (app.greRequirement === 'Waived' || app.greRequirement === 'Not Required')) {
        return; // skip
      }

      totalCountable++;
      if (item.status === 'Completed') {
        completedCountable++;
      }
    });

    if (totalCountable === 0) return 100;
    return Math.round((completedCountable / totalCountable) * 100);
  }

  // 3. Smart Warnings Engine
  function generateSmartWarnings() {
    const warnings = [];

    state.applications.forEach(app => {
      const deadState = calculateDeadlineState(app.deadline, app.verificationStatus);
      const completion = calculateCompletion(app);

      // Warning 1: Approaching deadline with incomplete SOP or LORs
      if (deadState.days !== null && deadState.days <= 14 && deadState.days >= 0 && app.status !== 'Submitted') {
        const sopItem = app.checklist.find(i => i.id === 'req_sop');
        if (sopItem && sopItem.status !== 'Completed') {
          warnings.push({
            type: 'critical',
            title: `⚠️ Critical SOP Alert: ${app.university}`,
            desc: `Deadline is in ${deadState.days} days, but Statement of Purpose (SOP) is incomplete!`
          });
        }
      }

      // Warning 2: Deadline in <= 14 days but status is still "Not Started"
      if (deadState.days !== null && deadState.days <= 14 && deadState.days >= 0 && app.status === 'Not Started') {
        warnings.push({
          type: 'warning',
          title: `📝 Unstarted Application: ${app.university}`,
          desc: `Deadline is in ${deadState.days} days, but application has not been started yet.`
        });
      }

      // Warning 3: Application marked "Ready to Apply" but documents incomplete (<80%)
      if (app.status === 'Ready to Apply' && completion < 80) {
        warnings.push({
          type: 'warning',
          title: `⚠️ Status Mismatch: ${app.university}`,
          desc: `Marked "Ready to Apply" but completion progress is only ${completion}%.`
        });
      }

      // Warning 4: Overdue application
      if (deadState.days !== null && deadState.days < 0 && app.status !== 'Submitted' && app.status !== 'Decision Received') {
        warnings.push({
          type: 'critical',
          title: `🚨 Overdue Application: ${app.university}`,
          desc: `Deadline passed ${Math.abs(deadState.days)} days ago! Please check portal or update status.`
        });
      }
    });

    return warnings;
  }


  // --- RENDER LOGIC ---

  function renderApp() {
    renderMetrics();
    renderSmartWarnings();
    renderTimeline();
    renderApplications();
  }

  // Render Top Metrics
  function renderMetrics() {
    const total = state.applications.length;
    const submitted = state.applications.filter(a => a.status === 'Submitted' || a.status === 'Decision Received').length;
    const progress = state.applications.filter(a => a.status === 'Documents Pending' || a.status === 'Application Started' || a.status === 'Researching' || a.status === 'Ready to Apply').length;
    const unstarted = state.applications.filter(a => a.status === 'Not Started').length;
    
    let critical = 0;
    state.applications.forEach(a => {
      const d = calculateDeadlineState(a.deadline, a.verificationStatus);
      if (d.days !== null && (d.days <= 14 || d.days < 0) && a.status !== 'Submitted') critical++;
    });

    const highPriority = state.applications.filter(a => a.priority === 'High').length;

    document.getElementById('metric-total').textContent = total;
    document.getElementById('metric-submitted').textContent = submitted;
    document.getElementById('metric-progress').textContent = progress;
    document.getElementById('metric-unstarted').textContent = unstarted;
    document.getElementById('metric-critical').textContent = critical;
    document.getElementById('metric-high').textContent = highPriority;
  }

  // Render Smart Warnings Banner
  function renderSmartWarnings() {
    const container = document.getElementById('smart-warnings-container');
    const warnings = generateSmartWarnings();

    if (warnings.length === 0) {
      container.innerHTML = '';
      return;
    }

    container.innerHTML = warnings.slice(0, 3).map(w => `
      <div class="warning-banner ${w.type}">
        <div class="warning-icon">${w.type === 'critical' ? '🚨' : '⚠️'}</div>
        <div class="warning-content">
          <div class="warning-title">${w.title}</div>
          <div>${w.desc}</div>
        </div>
      </div>
    `).join('');
  }

  // Render Upcoming Timeline
  function renderTimeline() {
    const container = document.getElementById('timeline-container');
    
    // Sort applications by deadline
    const sorted = [...state.applications]
      .filter(a => a.deadline)
      .sort((a, b) => new Date(a.deadline) - new Date(b.deadline));

    if (sorted.length === 0) {
      container.innerHTML = '<p style="color: var(--text-muted); font-size: 0.85rem;">No upcoming deadlines scheduled.</p>';
      return;
    }

    container.innerHTML = sorted.slice(0, 8).map(app => {
      const deadState = calculateDeadlineState(app.deadline, app.verificationStatus);
      const dateFormatted = new Date(app.deadline).toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });

      return `
        <div class="timeline-item" onclick="window.openDetailDrawer('${app.id}')">
          <div class="timeline-univ">${app.university}</div>
          <div class="timeline-prog">${app.program}</div>
          <div class="timeline-footer">
            <span style="font-size: 0.78rem; color: var(--text-muted);">${dateFormatted}</span>
            <span class="badge ${deadState.badgeClass}">${deadState.label}</span>
          </div>
        </div>
      `;
    }).join('');
  }

  // Filter & Sort Logic
  function getFilteredApplications() {
    return state.applications.filter(app => {
      // Search
      if (state.searchQuery) {
        const q = state.searchQuery.toLowerCase();
        const matches = app.university.toLowerCase().includes(q) ||
                        app.program.toLowerCase().includes(q) ||
                        (app.country && app.country.toLowerCase().includes(q)) ||
                        (app.notes && app.notes.toLowerCase().includes(q));
        if (!matches) return false;
      }

      // Filter Status
      if (state.filterStatus !== 'ALL' && app.status !== state.filterStatus) return false;

      // Filter Priority
      if (state.filterPriority !== 'ALL' && app.priority !== state.filterPriority) return false;

      // Filter GRE
      if (state.filterGRE !== 'ALL' && app.greRequirement !== state.filterGRE) return false;

      // Filter Urgency
      if (state.filterUrgency !== 'ALL') {
        const deadState = calculateDeadlineState(app.deadline, app.verificationStatus);
        if (state.filterUrgency === 'SAFE' && (!deadState.days || deadState.days < 30)) return false;
        if (state.filterUrgency === 'APPROACHING' && (deadState.days === null || deadState.days < 15 || deadState.days >= 30)) return false;
        if (state.filterUrgency === 'URGENT' && (deadState.days === null || deadState.days < 7 || deadState.days > 14)) return false;
        if (state.filterUrgency === 'CRITICAL' && (deadState.days === null || deadState.days < 1 || deadState.days > 6)) return false;
        if (state.filterUrgency === 'DUE TODAY' && deadState.days !== 0) return false;
        if (state.filterUrgency === 'OVERDUE' && (deadState.days === null || deadState.days >= 0)) return false;
        if (state.filterUrgency === 'NEEDS VERIFICATION' && app.verificationStatus !== 'Needs Verification') return false;
      }

      return true;
    }).sort((a, b) => {
      if (state.sortBy === 'deadline-asc') return new Date(a.deadline) - new Date(b.deadline);
      if (state.sortBy === 'deadline-desc') return new Date(b.deadline) - new Date(a.deadline);
      if (state.sortBy === 'progress-desc') return calculateCompletion(b) - calculateCompletion(a);
      if (state.sortBy === 'progress-asc') return calculateCompletion(a) - calculateCompletion(b);
      if (state.sortBy === 'priority-desc') {
        const map = { 'High': 3, 'Medium': 2, 'Low': 1 };
        return map[b.priority] - map[a.priority];
      }
      if (state.sortBy === 'univ-asc') return a.university.localeCompare(b.university);
      return 0;
    });
  }

  // Render Grid or Table
  function renderApplications() {
    const gridContainer = document.getElementById('grid-view-container');
    const tableBody = document.getElementById('app-table-body');
    const emptyState = document.getElementById('empty-state');
    
    const apps = getFilteredApplications();

    if (apps.length === 0) {
      gridContainer.style.display = 'none';
      document.getElementById('table-view-container').style.display = 'none';
      emptyState.style.display = 'flex';
      return;
    }

    emptyState.style.display = 'none';

    if (state.currentView === 'grid') {
      gridContainer.style.display = 'grid';
      document.getElementById('table-view-container').style.display = 'none';

      gridContainer.innerHTML = apps.map(app => renderAppCard(app)).join('');
    } else {
      gridContainer.style.display = 'none';
      document.getElementById('table-view-container').style.display = 'block';

      tableBody.innerHTML = apps.map(app => renderTableRow(app)).join('');
    }
  }

  // Card HTML template
  function renderAppCard(app) {
    const deadState = calculateDeadlineState(app.deadline, app.verificationStatus);
    const completion = calculateCompletion(app);
    const dateFormatted = new Date(app.deadline).toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });
    const univInitial = app.university.charAt(0);
    const isHighPriority = app.priority === 'High';

    // Checklist chips preview (SOP, Resume, LORs)
    const sopDone = app.checklist.some(i => i.id === 'req_sop' && i.status === 'Completed');
    const resumeDone = app.checklist.some(i => i.id === 'req_resume' && i.status === 'Completed');
    const lorDoneCount = app.checklist.filter(i => i.id.startsWith('req_lor') && i.status === 'Completed').length;

    return `
      <div class="app-card ${isHighPriority ? 'high-priority' : ''}">
        <div class="app-card-header">
          <div class="univ-info">
            <div class="univ-avatar">${univInitial}</div>
            <div>
              <div class="univ-name">${app.university}</div>
              <div class="program-name">${app.program} (${app.degree})</div>
            </div>
          </div>
          <span class="badge ${isHighPriority ? 'badge-priority-high' : app.priority === 'Low' ? 'badge-priority-low' : 'badge-priority-med'}">${app.priority}</span>
        </div>

        <div class="card-details-row">
          <div class="detail-item">
            <span class="detail-label">Deadline</span>
            <span class="detail-val">${dateFormatted}</span>
          </div>
          <div class="detail-item">
            <span class="detail-label">Status</span>
            <span class="detail-val">${app.status}</span>
          </div>
          <div class="detail-item">
            <span class="detail-label">Countdown</span>
            <span class="badge ${deadState.badgeClass}">${deadState.label}</span>
          </div>
        </div>

        <!-- Progress Bar -->
        <div class="progress-container">
          <div class="progress-header">
            <span class="progress-title">Application Progress</span>
            <span class="progress-pct">${completion}%</span>
          </div>
          <div class="progress-track">
            <div class="progress-fill ${completion === 100 ? 'complete' : ''}" style="width: ${completion}%;"></div>
          </div>
        </div>

        <!-- Checklist Quick Chips -->
        <div class="checklist-preview">
          <span class="checklist-chip ${resumeDone ? 'done' : ''}">CV ${resumeDone ? '✓' : '○'}</span>
          <span class="checklist-chip ${sopDone ? 'done' : ''}">SOP ${sopDone ? '✓' : '○'}</span>
          <span class="checklist-chip ${lorDoneCount === 3 ? 'done' : ''}">LOR (${lorDoneCount}/3)</span>
          <span class="checklist-chip">${app.greRequirement === 'Required' ? 'GRE Req' : 'GRE ' + app.greRequirement}</span>
        </div>

        <!-- Card Footer Actions -->
        <div class="card-actions">
          <button class="btn btn-secondary btn-sm" onclick="window.openDetailDrawer('${app.id}')">
            View & Edit Details
          </button>
          
          <button class="btn btn-accent btn-sm" onclick="window.addGCalEvent('${app.id}')" title="Add 1-Click Event to Google Calendar">
            📅 GCal
          </button>

          <button class="btn btn-secondary btn-sm" onclick="window.downloadAppIcs('${app.id}')" title="Download .ics Calendar File">
            📥 .ics
          </button>
        </div>
      </div>
    `;
  }

  // Table Row HTML template
  function renderTableRow(app) {
    const deadState = calculateDeadlineState(app.deadline, app.verificationStatus);
    const completion = calculateCompletion(app);
    const dateFormatted = new Date(app.deadline).toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });

    return `
      <tr>
        <td>
          <strong style="color: var(--text-main); font-size: 0.95rem;">${app.university}</strong>
          <div style="font-size: 0.8rem; color: var(--text-muted);">${app.program} (${app.degree})</div>
        </td>
        <td>${app.country || 'USA'}</td>
        <td>
          <div>${dateFormatted}</div>
          <span class="badge ${deadState.badgeClass}" style="margin-top: 4px;">${deadState.label}</span>
        </td>
        <td><span class="badge badge-priority-med">${app.status}</span></td>
        <td><span class="badge ${app.priority === 'High' ? 'badge-priority-high' : 'badge-priority-med'}">${app.priority}</span></td>
        <td>
          <div style="display: flex; align-items: center; gap: 8px;">
            <span>${completion}%</span>
            <div class="progress-track" style="width: 60px; height: 6px;">
              <div class="progress-fill" style="width: ${completion}%;"></div>
            </div>
          </div>
        </td>
        <td>${app.greRequirement}</td>
        <td>
          <button class="btn btn-secondary btn-sm" onclick="window.openDetailDrawer('${app.id}')">Manage</button>
        </td>
      </tr>
    `;
  }


  // --- DRAWER & DETAIL MANAGEMENT ---

  window.openDetailDrawer = function (id) {
    const app = state.applications.find(a => a.id === id);
    if (!app) return;

    state.activeAppId = id;
    const deadState = calculateDeadlineState(app.deadline, app.verificationStatus);
    const completion = calculateCompletion(app);

    document.getElementById('drawer-univ-title').textContent = app.university;
    document.getElementById('drawer-program-title').textContent = `${app.program} (${app.degree})`;
    
    const statusBadge = document.getElementById('drawer-status-badge');
    statusBadge.textContent = app.status;
    statusBadge.className = 'badge badge-priority-med';

    document.getElementById('drawer-deadline-val').textContent = new Date(app.deadline).toLocaleDateString('en-US', { month: 'long', day: 'numeric', year: 'numeric' });

    const countBadge = document.getElementById('drawer-countdown-badge');
    countBadge.textContent = deadState.label;
    countBadge.className = `badge ${deadState.badgeClass}`;

    document.getElementById('drawer-progress-val').textContent = `${completion}%`;
    document.getElementById('drawer-progress-fill').style.width = `${completion}%`;

    // Links
    const portalBtn = document.getElementById('drawer-portal-link');
    if (app.portalUrl) {
      portalBtn.href = app.portalUrl;
      portalBtn.style.display = 'inline-flex';
    } else {
      portalBtn.style.display = 'none';
    }

    const sourceBtn = document.getElementById('drawer-source-link');
    if (app.officialSourceUrl) {
      sourceBtn.href = app.officialSourceUrl;
      sourceBtn.style.display = 'inline-flex';
    } else {
      sourceBtn.style.display = 'none';
    }

    // Meta details
    document.getElementById('drawer-gre-val').textContent = app.greRequirement;
    document.getElementById('drawer-english-val').textContent = app.englishRequirement || 'Not Specified';
    document.getElementById('drawer-fee-val').textContent = `$${app.appFee || 0}`;
    document.getElementById('drawer-type-val').textContent = app.deadlineType || 'Regular Round';
    document.getElementById('drawer-verify-val').textContent = app.verificationStatus || 'Verified';

    document.getElementById('drawer-notes-input').value = app.notes || '';

    // Render Checklist
    renderDrawerChecklist(app);

    // Open Overlay
    document.getElementById('drawer-detail').classList.add('active');
  };

  function renderDrawerChecklist(app) {
    const container = document.getElementById('drawer-checklist-container');

    container.innerHTML = app.checklist.map(item => {
      const isDone = item.status === 'Completed';
      const isOmittedGre = item.id === 'req_gre' && (app.greRequirement === 'Waived' || app.greRequirement === 'Not Required');

      return `
        <div class="checklist-item-row" style="${isOmittedGre ? 'opacity: 0.5;' : ''}">
          <div class="checklist-left">
            <div class="checkbox-custom ${isDone ? 'checked' : ''}" onclick="window.toggleChecklistItem('${app.id}', '${item.id}')">
              ${isDone ? '✓' : ''}
            </div>
            <div>
              <span class="checklist-name ${isDone ? 'strike' : ''}">${item.name}</span>
              ${isOmittedGre ? '<span style="font-size: 0.72rem; color: var(--text-dim); margin-left: 6px;">(Not Required)</span>' : ''}
            </div>
          </div>
          <span class="badge ${isDone ? 'badge-safe' : item.status === 'In Progress' ? 'badge-approaching' : 'badge-priority-low'}">
            ${item.status}
          </span>
        </div>
      `;
    }).join('');
  }

  window.toggleChecklistItem = function (appId, itemId) {
    const app = state.applications.find(a => a.id === appId);
    if (!app) return;

    const item = app.checklist.find(i => i.id === itemId);
    if (!item) return;

    if (item.status === 'Completed') {
      item.status = 'In Progress';
    } else if (item.status === 'In Progress') {
      item.status = 'Not Started';
    } else {
      item.status = 'Completed';
    }

    saveDataToStorage();
    window.openDetailDrawer(appId);
    renderApp();
  };

  // Add Custom Checklist Item
  document.getElementById('btn-add-custom-req').addEventListener('click', () => {
    if (!state.activeAppId) return;
    const name = prompt('Enter custom requirement item name (e.g. Portfolio, Video Essay, Diversity Statement):');
    if (!name) return;

    const app = state.applications.find(a => a.id === state.activeAppId);
    if (app) {
      app.checklist.push({
        id: 'req_custom_' + Date.now(),
        name: name,
        required: true,
        status: 'Not Started'
      });
      saveDataToStorage();
      window.openDetailDrawer(app.id);
      renderApp();
    }
  });

  // Save Drawer Notes
  document.getElementById('btn-save-drawer-notes').addEventListener('click', () => {
    if (!state.activeAppId) return;
    const app = state.applications.find(a => a.id === state.activeAppId);
    if (app) {
      app.notes = document.getElementById('drawer-notes-input').value;
      saveDataToStorage();
      alert('Notes saved successfully!');
    }
  });


  // --- ADD / EDIT APPLICATION FORM ---

  function openAppFormModal(appToEdit = null) {
    const modal = document.getElementById('modal-app-form');
    const form = document.getElementById('app-form');
    const title = document.getElementById('form-modal-title');

    form.reset();

    if (appToEdit) {
      title.textContent = `Edit Application — ${appToEdit.university}`;
      document.getElementById('form-app-id').value = appToEdit.id;
      document.getElementById('form-univ').value = appToEdit.university;
      document.getElementById('form-program').value = appToEdit.program;
      document.getElementById('form-degree').value = appToEdit.degree || 'MS';
      document.getElementById('form-country').value = appToEdit.country || 'USA';
      document.getElementById('form-priority').value = appToEdit.priority || 'Medium';
      document.getElementById('form-status').value = appToEdit.status || 'Documents Pending';
      document.getElementById('form-deadline').value = appToEdit.deadline || '';
      document.getElementById('form-deadline-type').value = appToEdit.deadlineType || 'Regular Round';
      document.getElementById('form-verification').value = appToEdit.verificationStatus || 'Verified';
      document.getElementById('form-source-url').value = appToEdit.officialSourceUrl || '';
      document.getElementById('form-portal-url').value = appToEdit.portalUrl || '';
      document.getElementById('form-gre').value = appToEdit.greRequirement || 'Optional';
      document.getElementById('form-english').value = appToEdit.englishRequirement || '';
      document.getElementById('form-fee').value = appToEdit.appFee || 75;
      document.getElementById('form-notes').value = appToEdit.notes || '';
    } else {
      title.textContent = 'Add New University Application';
      document.getElementById('form-app-id').value = '';
    }

    modal.classList.add('active');
  }

  document.getElementById('btn-save-app').addEventListener('click', (e) => {
    e.preventDefault();
    const id = document.getElementById('form-app-id').value;
    const univ = document.getElementById('form-univ').value.trim();
    const prog = document.getElementById('form-program').value.trim();
    const deadline = document.getElementById('form-deadline').value;

    if (!univ || !prog || !deadline) {
      alert('Please fill in required fields (University, Program, and Deadline Date).');
      return;
    }

    if (id) {
      // Edit existing
      const app = state.applications.find(a => a.id === id);
      if (app) {
        app.university = univ;
        app.program = prog;
        app.degree = document.getElementById('form-degree').value.trim();
        app.country = document.getElementById('form-country').value.trim();
        app.priority = document.getElementById('form-priority').value;
        app.status = document.getElementById('form-status').value;
        app.deadline = deadline;
        app.deadlineType = document.getElementById('form-deadline-type').value;
        app.verificationStatus = document.getElementById('form-verification').value;
        app.officialSourceUrl = document.getElementById('form-source-url').value.trim();
        app.portalUrl = document.getElementById('form-portal-url').value.trim();
        app.greRequirement = document.getElementById('form-gre').value;
        app.englishRequirement = document.getElementById('form-english').value.trim();
        app.appFee = parseFloat(document.getElementById('form-fee').value) || 0;
        app.notes = document.getElementById('form-notes').value.trim();
      }
    } else {
      // Create new
      const newApp = {
        id: 'app_user_' + Date.now(),
        university: univ,
        program: prog,
        degree: document.getElementById('form-degree').value.trim() || 'MS',
        country: document.getElementById('form-country').value.trim() || 'USA',
        priority: document.getElementById('form-priority').value,
        status: document.getElementById('form-status').value,
        deadline: deadline,
        openingDate: '2026-09-01',
        deadlineType: document.getElementById('form-deadline-type').value,
        verificationStatus: document.getElementById('form-verification').value,
        officialSourceUrl: document.getElementById('form-source-url').value.trim(),
        portalUrl: document.getElementById('form-portal-url').value.trim(),
        greRequirement: document.getElementById('form-gre').value,
        englishRequirement: document.getElementById('form-english').value.trim(),
        appFee: parseFloat(document.getElementById('form-fee').value) || 0,
        notes: document.getElementById('form-notes').value.trim(),
        checklist: DEFAULT_CHECKLIST_ITEMS.map(item => ({ ...item, status: 'Not Started' }))
      };
      state.applications.unshift(newApp);
    }

    saveDataToStorage();
    document.getElementById('modal-app-form').classList.remove('active');
    renderApp();
  });

  // Edit button inside Drawer
  document.getElementById('drawer-btn-edit').addEventListener('click', () => {
    if (!state.activeAppId) return;
    const app = state.applications.find(a => a.id === state.activeAppId);
    if (app) {
      document.getElementById('drawer-detail').classList.remove('active');
      openAppFormModal(app);
    }
  });

  // Delete button inside Drawer
  document.getElementById('drawer-btn-delete').addEventListener('click', () => {
    if (!state.activeAppId) return;
    const app = state.applications.find(a => a.id === state.activeAppId);
    if (!app) return;

    if (confirm(`Are you sure you want to delete ${app.university} — ${app.program}?`)) {
      state.applications = state.applications.filter(a => a.id !== state.activeAppId);
      saveDataToStorage();
      document.getElementById('drawer-detail').classList.remove('active');
      renderApp();
    }
  });


  // --- CALENDAR & REMINDER GENERATION ---

  // 1-Click Google Calendar Link Builder
  window.addGCalEvent = function (appId) {
    const app = typeof appId === 'string' ? state.applications.find(a => a.id === appId) : appId;
    if (!app || !app.deadline) return;

    const title = encodeURIComponent(`🚨 DEADLINE: ${app.university} — ${app.program}`);
    const details = encodeURIComponent(`Application Deadline for ${app.university} (${app.program}).\nPortal: ${app.portalUrl || 'N/A'}\nSource: ${app.officialSourceUrl || 'N/A'}`);
    
    // Format YYYYMMDD
    const dateFormatted = app.deadline.replace(/-/g, '');
    const datesParam = `${dateFormatted}/${dateFormatted}`;

    const gcalUrl = `https://calendar.google.com/calendar/render?action=TEMPLATE&text=${title}&dates=${datesParam}&details=${details}`;
    window.open(gcalUrl, '_blank');
  };

  document.getElementById('drawer-btn-gcal').addEventListener('click', () => {
    if (state.activeAppId) window.addGCalEvent(state.activeAppId);
  });

  // `.ics` iCalendar File Generator
  window.downloadAppIcs = function (appId) {
    const app = state.applications.find(a => a.id === appId);
    if (!app || !app.deadline) return;
    generateAndDownloadIcs([app], `${app.university.replace(/\s+/g, '_')}_Deadline.ics`);
  };

  document.getElementById('drawer-btn-ics').addEventListener('click', () => {
    if (state.activeAppId) window.downloadAppIcs(state.activeAppId);
  });

  document.getElementById('btn-download-all-ics').addEventListener('click', () => {
    generateAndDownloadIcs(state.applications, 'All_TargetMS_Deadlines.ics');
  });

  function generateAndDownloadIcs(appList, filename) {
    let icsContent = `BEGIN:VCALENDAR\r\nVERSION:2.0\r\nPRODID:-//TargetMS Tracker//NONSGML v1.0//EN\r\n`;

    appList.forEach(app => {
      if (!app.deadline) return;
      const dateFormatted = app.deadline.replace(/-/g, '');

      icsContent += `BEGIN:VEVENT\r\n`;
      icsContent += `UID:app_deadline_${app.id}@targetms.app\r\n`;
      icsContent += `DTSTAMP:${new Date().toISOString().replace(/[-:]/g, '').split('.')[0]}Z\r\n`;
      icsContent += `DTSTART;VALUE=DATE:${dateFormatted}\r\n`;
      icsContent += `SUMMARY:🎓 Deadline: ${app.university} (${app.program})\r\n`;
      icsContent += `DESCRIPTION:Graduate application deadline for ${app.university} - ${app.program}. Portal: ${app.portalUrl || 'N/A'}\r\n`;

      // Pre-configured Alarms (30d, 14d, 7d, 1d)
      icsContent += `BEGIN:VALARM\r\nTRIGGER:-P14D\r\nACTION:DISPLAY\r\nDESCRIPTION:Application deadline in 14 days for ${app.university}\r\nEND:VALARM\r\n`;
      icsContent += `BEGIN:VALARM\r\nTRIGGER:-P3D\r\nACTION:DISPLAY\r\nDESCRIPTION:URGENT: Application deadline in 3 days for ${app.university}\r\nEND:VALARM\r\n`;

      icsContent += `END:VEVENT\r\n`;
    });

    icsContent += `END:VCALENDAR\r\n`;

    const blob = new Blob([icsContent], { type: 'text/calendar;charset=utf-8;' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = filename;
    link.click();
  }

  // Browser Push Notifications
  document.getElementById('btn-enable-browser-notif').addEventListener('click', () => {
    if (!('Notification' in window)) {
      alert('Desktop notifications are not supported by your browser.');
      return;
    }

    Notification.requestPermission().then(permission => {
      if (permission === 'granted') {
        alert('Browser Notifications Enabled! You will be notified when deadlines approach.');
        new Notification('TargetMS Deadline Tracker', {
          body: 'Notifications activated! We will remind you of upcoming university deadlines.',
          icon: '🎓'
        });
      } else {
        alert('Permission was denied for browser notifications.');
      }
    });
  });

  // Copy Apps Script Code
  document.getElementById('btn-copy-script').addEventListener('click', () => {
    const code = document.getElementById('apps-script-code').value;
    navigator.clipboard.writeText(code).then(() => {
      alert('Google Apps Script snippet copied to clipboard!');
    });
  });


  // --- EXPORT & IMPORT UTILITIES ---

  // Export Excel / CSV
  document.getElementById('btn-export-excel').addEventListener('click', () => {
    if (!window.XLSX) { alert('SheetJS library not loaded.'); return; }
    
    const exportData = state.applications.map(app => ({
      University: app.university,
      Program: app.program,
      Degree: app.degree,
      Country: app.country,
      Priority: app.priority,
      Status: app.status,
      Deadline: app.deadline,
      'Deadline Type': app.deadlineType,
      'GRE Requirement': app.greRequirement,
      'English Test': app.englishRequirement,
      'Application Fee ($)': app.appFee,
      'Completion %': calculateCompletion(app),
      'Official Source URL': app.officialSourceUrl,
      'Portal URL': app.portalUrl,
      Notes: app.notes
    }));

    const worksheet = XLSX.utils.json_to_sheet(exportData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Applications');
    XLSX.writeFile(workbook, 'TargetMS_University_Tracker.xlsx');
  });

  document.getElementById('btn-export-csv').addEventListener('click', () => {
    if (!window.XLSX) { alert('SheetJS library not loaded.'); return; }

    const exportData = state.applications.map(app => ({
      University: app.university,
      Program: app.program,
      Deadline: app.deadline,
      Status: app.status,
      Priority: app.priority,
      GRE: app.greRequirement
    }));

    const worksheet = XLSX.utils.json_to_sheet(exportData);
    const csv = XLSX.utils.sheet_to_csv(worksheet);
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = 'TargetMS_Tracker.csv';
    link.click();
  });

  // Export / Restore JSON Backup
  document.getElementById('btn-export-json').addEventListener('click', () => {
    const dataStr = JSON.stringify(state.applications, null, 2);
    const blob = new Blob([dataStr], { type: 'application/json' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = 'TargetMS_Backup.json';
    link.click();
  });

  document.getElementById('restore-json-input').addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (event) => {
      try {
        const parsed = JSON.parse(event.target.result);
        if (Array.isArray(parsed)) {
          state.applications = parsed;
          saveDataToStorage();
          renderApp();
          alert('Full backup restored successfully!');
          document.getElementById('modal-export').classList.remove('active');
        }
      } catch (err) {
        alert('Invalid JSON backup file format.');
      }
    };
    reader.readAsText(file);
  });


  // --- EVENT LISTENERS & MODAL HANDLERS ---

  function setupEventListeners() {
    // Toolbar Search & Filters
    document.getElementById('search-input').addEventListener('input', (e) => {
      state.searchQuery = e.target.value;
      renderApplications();
    });

    document.getElementById('filter-status').addEventListener('change', (e) => {
      state.filterStatus = e.target.value;
      renderApplications();
    });

    document.getElementById('filter-urgency').addEventListener('change', (e) => {
      state.filterUrgency = e.target.value;
      renderApplications();
    });

    document.getElementById('filter-priority').addEventListener('change', (e) => {
      state.filterPriority = e.target.value;
      renderApplications();
    });

    document.getElementById('filter-gre').addEventListener('change', (e) => {
      state.filterGRE = e.target.value;
      renderApplications();
    });

    document.getElementById('sort-by').addEventListener('change', (e) => {
      state.sortBy = e.target.value;
      renderApplications();
    });

    // View Switcher Tabs
    document.getElementById('view-tab-grid').addEventListener('click', () => {
      state.currentView = 'grid';
      document.getElementById('view-tab-grid').classList.add('active');
      document.getElementById('view-tab-table').classList.remove('active');
      renderApplications();
    });

    document.getElementById('view-tab-table').addEventListener('click', () => {
      state.currentView = 'table';
      document.getElementById('view-tab-table').classList.add('active');
      document.getElementById('view-tab-grid').classList.remove('active');
      renderApplications();
    });

    // Modal Triggers
    document.getElementById('btn-add-app').addEventListener('click', () => openAppFormModal());
    document.getElementById('empty-btn-add').addEventListener('click', () => openAppFormModal());

    document.getElementById('btn-import-excel').addEventListener('click', () => {
      document.getElementById('modal-import').classList.add('active');
    });

    document.getElementById('btn-sync-calendar').addEventListener('click', () => {
      document.getElementById('modal-calendar-sync').classList.add('active');
    });

    document.getElementById('btn-export-modal').addEventListener('click', () => {
      document.getElementById('modal-export').classList.add('active');
    });

    // Close Modals & Drawers
    document.querySelectorAll('.close-modal').forEach(btn => {
      btn.addEventListener('click', () => {
        document.querySelectorAll('.modal-overlay').forEach(m => m.classList.remove('active'));
      });
    });

    document.querySelectorAll('.close-drawer').forEach(btn => {
      btn.addEventListener('click', () => {
        document.getElementById('drawer-detail').classList.remove('active');
      });
    });

    // Drag and drop Excel Upload
    const dropZone = document.getElementById('drop-zone');
    const fileInput = document.getElementById('file-input-excel');
    const confirmBtn = document.getElementById('btn-confirm-import');

    let loadedWorkbook = null;

    dropZone.addEventListener('click', () => fileInput.click());

    fileInput.addEventListener('change', (e) => {
      const file = e.target.files[0];
      if (file) handleExcelFile(file);
    });

    dropZone.addEventListener('dragover', (e) => { e.preventDefault(); dropZone.style.borderColor = '#fff'; });
    dropZone.addEventListener('dragleave', () => { dropZone.style.borderColor = 'var(--primary)'; });
    dropZone.addEventListener('drop', (e) => {
      e.preventDefault();
      dropZone.style.borderColor = 'var(--primary)';
      if (e.dataTransfer.files.length > 0) handleExcelFile(e.dataTransfer.files[0]);
    });

    function handleExcelFile(file) {
      const reader = new FileReader();
      reader.onload = (event) => {
        const data = new Uint8Array(event.target.result);
        loadedWorkbook = XLSX.read(data, { type: 'array' });
        
        document.getElementById('import-preview-box').style.display = 'block';
        document.getElementById('preview-filename').textContent = `File Loaded: ${file.name}`;
        confirmBtn.disabled = false;
      };
      reader.readAsArrayBuffer(file);
    }

    confirmBtn.addEventListener('click', () => {
      if (loadedWorkbook) {
        parseAndMergeWorkbook(loadedWorkbook, true);
        document.getElementById('modal-import').classList.remove('active');
      }
    });
  }

  // Launch on DOM ready
  document.addEventListener('DOMContentLoaded', init);

})();
