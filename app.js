/* ==========================================================================
   TargetMS — Complete GRE / MS Application Deadline Tracker Application Engine
   ========================================================================== */

(function () {
  'use strict';

  // --- STORAGE KEYS & INITIAL STATE ---
  const STORAGE_KEY = 'targetms_applications_v5';
  const PROFILE_STORAGE_KEY = 'targetms_user_profile_v1';
  const SENT_MILESTONES_KEY = 'targetms_sent_milestones_v1';
  const LAST_WEEKLY_DIGEST_KEY = 'targetms_last_weekly_digest_v1';

  // Default User Profile
  const DEFAULT_USER_PROFILE = {
    name: 'Sarah Jenkins',
    email: 'sarah.gre2026@gmail.com',
    autoEmail: true,
    emailStrategy: 'weekly_digest', // Weekly Top 7 Approaching Digest
    emailJsService: 'service_targetms',
    emailJsTemplate: 'template_deadline',
    emailJsKey: 'user_targetms_free_key'
  };

  // Milestone Days for Strategy 1
  const MILESTONE_DAYS = [30, 14, 7, 3, 1];

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

  // Global State
  let state = {
    applications: [],
    userProfile: { ...DEFAULT_USER_PROFILE },
    sentMilestones: {},
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
    loadProfileFromStorage();
    loadSentMilestones();
    initEmailJsSDK();
    
    // Attempt auto-loading Colleges.xlsx as primary source of truth
    const loadedFromStorage = loadDataFromStorage();
    if (!loadedFromStorage) {
      autoTryLoadCollegesExcel();
    } else {
      renderApp();
      checkWeeklyDigestTrigger();
    }

    setupEventListeners();
  }

  // --- DATA LOADING & PERSISTENCE ---
  function loadDataFromStorage() {
    try {
      const saved = localStorage.getItem(STORAGE_KEY);
      if (saved) {
        state.applications = JSON.parse(saved);
        return true;
      }
    } catch (err) {
      console.error('Failed to load from storage:', err);
    }
    return false;
  }

  function saveDataToStorage() {
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(state.applications));
    } catch (err) {
      console.error('Failed to save to storage:', err);
    }
  }

  function loadProfileFromStorage() {
    try {
      const saved = localStorage.getItem(PROFILE_STORAGE_KEY);
      if (saved) {
        state.userProfile = { ...DEFAULT_USER_PROFILE, ...JSON.parse(saved) };
      }
    } catch (err) {
      console.error('Failed to load profile:', err);
    }
    updateProfileUI();
  }

  function saveProfileToStorage() {
    try {
      localStorage.setItem(PROFILE_STORAGE_KEY, JSON.stringify(state.userProfile));
      updateProfileUI();
    } catch (err) {
      console.error('Failed to save profile:', err);
    }
  }

  function loadSentMilestones() {
    try {
      const saved = localStorage.getItem(SENT_MILESTONES_KEY);
      if (saved) state.sentMilestones = JSON.parse(saved);
    } catch (err) {
      state.sentMilestones = {};
    }
  }

  function saveSentMilestones() {
    try {
      localStorage.setItem(SENT_MILESTONES_KEY, JSON.stringify(state.sentMilestones));
    } catch (err) {
      console.error('Failed to save sent milestones:', err);
    }
  }

  function updateProfileUI() {
    document.getElementById('user-display-name').textContent = state.userProfile.name || 'Applicant';
    document.getElementById('user-display-email').textContent = state.userProfile.email || 'Configure Email';
    
    // Avatar initials
    const nameParts = (state.userProfile.name || 'A P').split(' ');
    const initials = (nameParts[0].charAt(0) + (nameParts[1] ? nameParts[1].charAt(0) : '')).toUpperCase();
    document.getElementById('user-avatar-initials').textContent = initials;

    // Form values
    document.getElementById('profile-name').value = state.userProfile.name;
    document.getElementById('profile-email').value = state.userProfile.email;
    document.getElementById('emailjs-service-id').value = state.userProfile.emailJsService;
    document.getElementById('emailjs-template-id').value = state.userProfile.emailJsTemplate;
    document.getElementById('emailjs-public-key').value = state.userProfile.emailJsKey;
    document.getElementById('toggle-auto-email').checked = state.userProfile.autoEmail;
  }

  function initEmailJsSDK() {
    if (window.emailjs && state.userProfile.emailJsKey) {
      try {
        emailjs.init(state.userProfile.emailJsKey);
      } catch (e) {
        console.warn('EmailJS init warning:', e);
      }
    }
  }

  // Auto-attempt fetching & parsing Colleges.xlsx as primary source of truth
  async function autoTryLoadCollegesExcel() {
    if (window.XLSX) {
      try {
        const response = await fetch('./Colleges.xlsx');
        if (response.ok) {
          const buffer = await response.arrayBuffer();
          const workbook = XLSX.read(buffer, { type: 'array' });
          if (workbook && workbook.SheetNames.length > 0) {
            parseAndMergeWorkbook(workbook, false, true); // replaceMode = true
          }
        }
      } catch (err) {
        console.warn('Colleges.xlsx auto-load notice:', err);
      }
    }
  }

  // --- PARSE EXCEL WORKBOOK (EXACT ACCURACY ENGINE) ---
  function parseAndMergeWorkbook(workbook, notify = true, replaceMode = false) {
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
    const idxDeadline = findCol(['deadline', 'due date', 'date', 'due']);
    const idxPriority = findCol(['priority', 'importance', 'preference']);
    const idxGre = findCol(['gre']);
    const idxStatus = findCol(['status', 'state', 'stage']);
    const idxFee = findCol(['fee', 'cost']);
    const idxNotes = findCol(['note', 'comment', 'remark', 'requirement']);

    const parsedApps = [];

    for (let r = 1; r < rows.length; r++) {
      const row = rows[r];
      if (!row || row.length === 0) continue;

      const univName = idxUniv !== -1 && row[idxUniv] ? String(row[idxUniv]).trim() : '';
      if (!univName) continue;

      const progName = idxProg !== -1 && row[idxProg] ? String(row[idxProg]).trim() : 'MS Program';
      const degreeName = idxDegree !== -1 && row[idxDegree] ? String(row[idxDegree]).trim() : 'MS';
      const countryName = idxCountry !== -1 && row[idxCountry] ? String(row[idxCountry]).trim() : 'USA';
      const rawDeadline = idxDeadline !== -1 && row[idxDeadline] ? row[idxDeadline] : '';
      const priorityName = idxPriority !== -1 && row[idxPriority] ? String(row[idxPriority]).trim() : 'Medium';
      const greRule = idxGre !== -1 && row[idxGre] ? String(row[idxGre]).trim() : 'Optional';
      const statusName = idxStatus !== -1 && row[idxStatus] ? String(row[idxStatus]).trim() : 'Documents Pending';
      const feeVal = idxFee !== -1 && row[idxFee] ? parseFloat(row[idxFee]) || 75 : 75;
      const notesVal = idxNotes !== -1 && row[idxNotes] ? String(row[idxNotes]).trim() : '';

      // Parse Date: Extract first date from range & ensure ALL dates are for the FUTURE
      const dateResult = parseExcelDate(rawDeadline);
      const deadlineStr = dateResult.dateStr;
      const verificationStatus = dateResult.verified ? 'Verified' : 'Needs Verification';

      const id = 'excel_app_' + Date.now() + '_' + r;

      const appRecord = {
        id: id,
        university: univName,
        program: progName,
        degree: degreeName,
        country: countryName,
        priority: sanitizePriority(priorityName),
        status: sanitizeStatus(statusName),
        deadline: deadlineStr,
        rawDeadlineText: dateResult.rawText,
        openingDate: '2026-09-01',
        deadlineType: 'Regular Round',
        verificationStatus: verificationStatus,
        officialSourceUrl: '',
        portalUrl: '',
        greRequirement: sanitizeGre(greRule),
        englishRequirement: 'TOEFL / IELTS Required',
        appFee: feeVal,
        notes: notesVal,
        checklist: DEFAULT_CHECKLIST_ITEMS.map(item => ({ ...item, status: 'Not Started' }))
      };

      parsedApps.push(appRecord);
    }

    if (replaceMode || state.applications.length === 0) {
      state.applications = parsedApps;
    } else {
      // Merge unique
      parsedApps.forEach(newApp => {
        const exists = state.applications.some(a => 
          a.university.toLowerCase() === newApp.university.toLowerCase() && 
          a.program.toLowerCase() === newApp.program.toLowerCase()
        );
        if (!exists) state.applications.push(newApp);
      });
    }

    saveDataToStorage();
    renderApp();

    if (notify) {
      alert(`Loaded ${parsedApps.length} universities directly from your Excel spreadsheet!`);
    }
  }

  // Exact Date Parsing Engine: Ensures ALL dates are for the FUTURE (no accidental overdue dates!)
  function parseExcelDate(raw) {
    if (!raw) return { dateStr: '', rawText: '', verified: false };

    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const currentYear = today.getFullYear();

    // 1. Handle Excel numeric serial dates
    if (typeof raw === 'number') {
      let date = new Date((raw - (25567 + 2)) * 86400 * 1000);
      if (!isNaN(date.getTime())) {
        date.setHours(0, 0, 0, 0);
        if (date < today) {
          date.setFullYear(currentYear + (date.getMonth() < today.getMonth() ? 1 : 0));
        }
        if (date < today) {
          date.setFullYear(date.getFullYear() + 1);
        }
        const y = date.getFullYear();
        const m = String(date.getMonth() + 1).padStart(2, '0');
        const d = String(date.getDate()).padStart(2, '0');
        return { dateStr: `${y}-${m}-${d}`, rawText: `${y}-${m}-${d}`, verified: true };
      }
    }

    const str = String(raw).trim();
    if (!str || str.toLowerCase().includes('tbd') || str.toLowerCase().includes('n/a') || str.toLowerCase().includes('unknown')) {
      return { dateStr: '', rawText: str, verified: false };
    }

    // 2. Tentative Range Extraction: Take the FIRST segment before range delimiters (-, to, /, or, ,)
    const firstSegment = str.split(/\s*(-|–|to|\/|or|,)\s*/i)[0].trim();

    // 3. Try parsing direct standard JS date string
    const parsedDirect = new Date(firstSegment);
    if (!isNaN(parsedDirect.getTime()) && parsedDirect.getFullYear() > 2000) {
      parsedDirect.setHours(0, 0, 0, 0);
      let year = parsedDirect.getFullYear();
      
      // RULE: If date would fall in past, advance to next year!
      if (parsedDirect < today) {
        year = currentYear + 1;
        parsedDirect.setFullYear(year);
      }

      const month = String(parsedDirect.getMonth() + 1).padStart(2, '0');
      const day = String(parsedDirect.getDate()).padStart(2, '0');
      return { dateStr: `${year}-${month}-${day}`, rawText: str, verified: true };
    }

    // 4. Regex Pattern Matching for Month & Day in text
    const monthNames = "(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*";
    
    // Pattern A: "Dec 1" or "December 15, 2026"
    const patternA = new RegExp(`${monthNames}\\s+(\\d{1,2})(?:st|nd|rd|th)?(?:,?\\s+(\\d{4}))?`, "i");
    // Pattern B: "15 Dec" or "1st December 2026"
    const patternB = new RegExp(`(\\d{1,2})(?:st|nd|rd|th)?\\s+${monthNames}(?:,?\\s+(\\d{4}))?`, "i");

    let match = str.match(patternA);
    let monthStr = '', dayStr = '', yearStr = '';

    if (match) {
      monthStr = match[1];
      dayStr = match[2];
      yearStr = match[3];
    } else {
      match = str.match(patternB);
      if (match) {
        dayStr = match[1];
        monthStr = match[2];
        yearStr = match[3];
      }
    }

    if (monthStr && dayStr) {
      const monthMap = { jan: 0, feb: 1, mar: 2, apr: 3, may: 4, jun: 5, jul: 6, aug: 7, sep: 8, oct: 9, nov: 10, dec: 11 };
      const m = monthMap[monthStr.toLowerCase().substring(0, 3)];
      const d = parseInt(dayStr, 10);
      
      let y = yearStr ? parseInt(yearStr, 10) : currentYear;

      let dateObj = new Date(y, m, d);
      dateObj.setHours(0, 0, 0, 0);

      // RULE: If candidate date is before today AND year was not explicitly specified, advance to NEXT YEAR!
      if (dateObj < today && !yearStr) {
        y = currentYear + 1;
        dateObj = new Date(y, m, d);
      }

      if (!isNaN(dateObj.getTime())) {
        const monthPad = String(m + 1).padStart(2, '0');
        const dayPad = String(d).padStart(2, '0');
        return { dateStr: `${y}-${monthPad}-${dayPad}`, rawText: str, verified: true };
      }
    }

    // 5. Unstated or Unparseable Date: Return empty dateStr & mark unverified
    return { dateStr: '', rawText: str, verified: false };
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

  // 1. Live Countdown Calculator
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

      // Warning 5: Unverified deadline date
      if (app.verificationStatus === 'Needs Verification') {
        warnings.push({
          type: 'warning',
          title: `🔍 Date Unverified: ${app.university}`,
          desc: `Deadline was not clearly stated in Excel (${app.rawDeadlineText || 'Unstated'}). Click to verify date!`
        });
      }
    });

    return warnings;
  }

  // 4. AUTOMATED WEEKLY DIGEST ENGINE (DYNAMICALY HANDLES ANY NUMBER OF VALID DEADLINES)
  window.getApproachingApps = function (maxLimit = 7) {
    const valid = state.applications.filter(a => {
      if (a.status === 'Submitted' || a.status === 'Decision Received') return false;
      return Boolean(a.deadline || a.rawDeadlineText);
    });

    valid.sort((a, b) => {
      if (!a.deadline) return 1;
      if (!b.deadline) return -1;
      return new Date(a.deadline) - new Date(b.deadline);
    });

    return maxLimit ? valid.slice(0, maxLimit) : valid;
  };

  // Backwards compatibility alias
  window.getTop7ApproachingApps = function () {
    return window.getApproachingApps(7);
  };

  window.sendWeeklyTop7Digest = function (forceManual = false) {
    const recipient = state.userProfile.email;
    if (!recipient) {
      if (forceManual) {
        alert('Please configure your Notification Recipient Email in the Email Alerts Dashboard first!');
        document.getElementById('modal-email-center').classList.add('active');
      }
      return;
    }

    const approachingApps = window.getApproachingApps(7);
    if (approachingApps.length === 0) {
      if (forceManual) alert('No active upcoming applications found to include in the weekly digest.');
      return;
    }

    const appCount = approachingApps.length;
    const headerTitle = appCount === 1 ? '1 APPROACHING DEADLINE' : `${appCount} APPROACHING DEADLINES`;

    // Format Digest Body
    let digestSummary = `🎓 TARGETMS WEEKLY DIGEST — ${headerTitle}\n\n`;
    approachingApps.forEach((app, index) => {
      const deadState = calculateDeadlineState(app.deadline, app.verificationStatus);
      const completion = calculateCompletion(app);
      const displayDeadline = app.deadline || app.rawDeadlineText || 'Needs Verification';
      digestSummary += `${index + 1}. ${app.university} — ${app.program}\n`;
      digestSummary += `   📅 Deadline: ${displayDeadline} (${deadState.label})\n`;
      digestSummary += `   📊 Completion: ${completion}%\n`;
      digestSummary += `   🔗 Portal: ${app.portalUrl || 'N/A'}\n\n`;
    });

    const subjectText = appCount === 1 
      ? `📊 TargetMS Weekly Digest: 1 Approaching Application Deadline`
      : `📊 TargetMS Weekly Digest: ${appCount} Approaching Application Deadlines`;

    const emailParams = {
      to_name: state.userProfile.name,
      to_email: recipient,
      subject: subjectText,
      university_name: `${appCount} Approaching Application(s)`,
      program_name: `Weekly Digest Summary (${new Date().toLocaleDateString()})`,
      deadline_date: approachingApps[0].deadline || approachingApps[0].rawDeadlineText || 'N/A',
      days_remaining: calculateDeadlineState(approachingApps[0].deadline, approachingApps[0].verificationStatus).label,
      completion_pct: calculateCompletion(approachingApps[0]) + '%',
      portal_url: approachingApps[0].portalUrl || 'N/A'
    };

    if (window.emailjs && state.userProfile.emailJsKey && state.userProfile.emailJsKey !== 'user_targetms_free_key') {
      emailjs.send(state.userProfile.emailJsService, state.userProfile.emailJsTemplate, emailParams)
        .then(() => {
          localStorage.setItem(LAST_WEEKLY_DIGEST_KEY, new Date().toISOString());
          alert(`📧 Weekly Digest email successfully sent to ${recipient} (${appCount} application(s) included)!`);
        })
        .catch(err => {
          console.warn('EmailJS error:', err);
          simulateWeeklyDigestDispatch(recipient, subjectText, digestSummary, appCount);
        });
    } else {
      simulateWeeklyDigestDispatch(recipient, subjectText, digestSummary, appCount);
    }
  };

  let currentPreviewEmail = { recipient: '', subject: '', body: '' };

  function openEmailPreviewModal(recipient, subjectText, bodyText, modalTitle = '📧 Application Email Alert Ready') {
    currentPreviewEmail = { recipient: recipient, subject: subjectText, body: bodyText };

    const modalTitleEl = document.getElementById('preview-email-modal-title');
    const recipientEl = document.getElementById('preview-email-recipient');
    const subjectEl = document.getElementById('preview-email-subject');
    const bodyEl = document.getElementById('preview-email-body');

    if (modalTitleEl) modalTitleEl.textContent = modalTitle;
    if (recipientEl) recipientEl.textContent = recipient;
    if (subjectEl) subjectEl.value = subjectText;
    if (bodyEl) bodyEl.value = bodyText;

    const modalPreview = document.getElementById('modal-email-preview');
    if (modalPreview) modalPreview.classList.add('active');
  }

  function simulateWeeklyDigestDispatch(recipient, subjectText, digestSummary, appCount) {
    const rawBody = `Hi ${state.userProfile.name},\n\nHere is your automated weekly digest of your active approaching graduate application deadlines:\n\n${digestSummary}\nGood luck with your application prep!`;
    localStorage.setItem(LAST_WEEKLY_DIGEST_KEY, new Date().toISOString());
    openEmailPreviewModal(recipient, subjectText, rawBody, `📧 TargetMS Weekly Digest (${appCount} Applications)`);
  }

  function simulateEmailDispatch(app, recipient, subjectText) {
    const rawSubject = subjectText || `🚨 Deadline Reminder: ${app.university} (${app.program})`;
    const rawBody = `Hi ${state.userProfile.name},\n\nApplication Alert for ${app.university} - ${app.program}!\n\nDeadline Date: ${app.deadline || app.rawDeadlineText || 'Needs Verification'}\nStatus: ${app.status}\nProgress: ${calculateCompletion(app)}%\nPortal: ${app.portalUrl || 'N/A'}\n\nGood luck with your application prep!`;
    openEmailPreviewModal(recipient, rawSubject, rawBody, `📧 Email Alert: ${app.university}`);
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
      container.innerHTML = '<p style="color: var(--text-muted); font-size: 0.85rem;">No verified upcoming deadlines scheduled.</p>';
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
      if (state.sortBy === 'deadline-asc') {
        if (!a.deadline) return 1;
        if (!b.deadline) return -1;
        return new Date(a.deadline) - new Date(b.deadline);
      }
      if (state.sortBy === 'deadline-desc') {
        if (!a.deadline) return 1;
        if (!b.deadline) return -1;
        return new Date(b.deadline) - new Date(a.deadline);
      }
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
    const dateFormatted = app.deadline 
      ? new Date(app.deadline).toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' })
      : (app.rawDeadlineText || 'Needs Verification');
    
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
            View Details
          </button>

          <button class="btn btn-accent btn-sm" onclick="window.sendApplicationEmail('${app.id}')" title="Send Email Alert">
            📧 Email
          </button>
          
          <button class="btn btn-secondary btn-sm" onclick="window.addGCalEvent('${app.id}')" title="Add 1-Click Event to Google Calendar">
            📅 GCal
          </button>
        </div>
      </div>
    `;
  }

  // Table Row HTML template
  function renderTableRow(app) {
    const deadState = calculateDeadlineState(app.deadline, app.verificationStatus);
    const completion = calculateCompletion(app);
    const dateFormatted = app.deadline 
      ? new Date(app.deadline).toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' })
      : (app.rawDeadlineText || 'Needs Verification');

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
          <button class="btn btn-accent btn-sm" onclick="window.sendApplicationEmail('${app.id}')">📧 Email</button>
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

    const dateFormatted = app.deadline 
      ? new Date(app.deadline).toLocaleDateString('en-US', { month: 'long', day: 'numeric', year: 'numeric' })
      : (app.rawDeadlineText || 'Needs Verification');
    document.getElementById('drawer-deadline-val').textContent = dateFormatted;

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

  // --- ADD / EDIT APPLICATION FORM ---

  window.openAppFormModal = function (appToEdit = null) {
    const modal = document.getElementById('modal-app-form');
    const form = document.getElementById('app-form');
    const title = document.getElementById('form-modal-title');

    if (!modal) return;
    if (form) form.reset();

    if (appToEdit) {
      if (title) title.textContent = `Edit Application — ${appToEdit.university}`;
      const idEl = document.getElementById('form-app-id'); if (idEl) idEl.value = appToEdit.id;
      const univEl = document.getElementById('form-univ'); if (univEl) univEl.value = appToEdit.university;
      const progEl = document.getElementById('form-program'); if (progEl) progEl.value = appToEdit.program;
      const degEl = document.getElementById('form-degree'); if (degEl) degEl.value = appToEdit.degree || 'MS';
      const ctyEl = document.getElementById('form-country'); if (ctyEl) ctyEl.value = appToEdit.country || 'USA';
      const prioEl = document.getElementById('form-priority'); if (prioEl) prioEl.value = appToEdit.priority || 'Medium';
      const statEl = document.getElementById('form-status'); if (statEl) statEl.value = appToEdit.status || 'Documents Pending';
      const deadEl = document.getElementById('form-deadline'); if (deadEl) deadEl.value = appToEdit.deadline || '';
      const typeEl = document.getElementById('form-deadline-type'); if (typeEl) typeEl.value = appToEdit.deadlineType || 'Regular Round';
      const verEl = document.getElementById('form-verification'); if (verEl) verEl.value = appToEdit.verificationStatus || 'Verified';
      const srcEl = document.getElementById('form-source-url'); if (srcEl) srcEl.value = appToEdit.officialSourceUrl || '';
      const portEl = document.getElementById('form-portal-url'); if (portEl) portEl.value = appToEdit.portalUrl || '';
      const greEl = document.getElementById('form-gre'); if (greEl) greEl.value = appToEdit.greRequirement || 'Optional';
      const engEl = document.getElementById('form-english'); if (engEl) engEl.value = appToEdit.englishRequirement || '';
      const feeEl = document.getElementById('form-fee'); if (feeEl) feeEl.value = appToEdit.appFee || 75;
      const noteEl = document.getElementById('form-notes'); if (noteEl) noteEl.value = appToEdit.notes || '';
    } else {
      if (title) title.textContent = 'Add New University Application';
      const idEl = document.getElementById('form-app-id');
      if (idEl) idEl.value = '';
    }

    modal.classList.add('active');
  };

  window.closeDetailDrawer = function () {
    const d = document.getElementById('drawer-detail');
    if (d) d.classList.remove('active');
  };

  window.closeAllModals = function () {
    document.querySelectorAll('.modal-overlay').forEach(m => m.classList.remove('active'));
    document.querySelectorAll('.drawer-overlay').forEach(d => d.classList.remove('active'));
  };

  // 1-Click Google Calendar Link Builder
  window.addGCalEvent = function (appId) {
    const app = typeof appId === 'string' ? state.applications.find(a => a.id === appId) : appId;
    if (!app || !app.deadline) {
      alert('Cannot add to Google Calendar without a verified deadline date. Please edit/verify deadline first!');
      return;
    }

    const title = encodeURIComponent(`🚨 DEADLINE: ${app.university} — ${app.program}`);
    const details = encodeURIComponent(`Application Deadline for ${app.university} (${app.program}).\nPortal: ${app.portalUrl || 'N/A'}\nSource: ${app.officialSourceUrl || 'N/A'}`);
    
    // Format YYYYMMDD
    const dateFormatted = app.deadline.replace(/-/g, '');
    const datesParam = `${dateFormatted}/${dateFormatted}`;

    const gcalUrl = `https://calendar.google.com/calendar/render?action=TEMPLATE&text=${title}&dates=${datesParam}&details=${details}`;
    window.open(gcalUrl, '_blank');
  };

  // `.ics` iCalendar File Generator
  window.downloadAppIcs = function (appId) {
    const app = state.applications.find(a => a.id === appId);
    if (!app || !app.deadline) {
      alert('Cannot export .ics calendar without a verified deadline date.');
      return;
    }
    generateAndDownloadIcs([app], `${app.university.replace(/\s+/g, '_')}_Deadline.ics`);
  };

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

  // Helper function to safely bind event listeners without throwing errors if elements are missing
  function bindEvent(id, event, handler) {
    const el = document.getElementById(id);
    if (el) {
      el.addEventListener(event, handler);
    }
  }

  // --- EVENT LISTENERS & MODAL HANDLERS ---

  function setupEventListeners() {
    // Toolbar Search & Filters
    bindEvent('search-input', 'input', (e) => {
      state.searchQuery = e.target.value;
      renderApplications();
    });

    bindEvent('filter-status', 'change', (e) => {
      state.filterStatus = e.target.value;
      renderApplications();
    });

    bindEvent('filter-urgency', 'change', (e) => {
      state.filterUrgency = e.target.value;
      renderApplications();
    });

    bindEvent('filter-priority', 'change', (e) => {
      state.filterPriority = e.target.value;
      renderApplications();
    });

    bindEvent('filter-gre', 'change', (e) => {
      state.filterGRE = e.target.value;
      renderApplications();
    });

    bindEvent('sort-by', 'change', (e) => {
      state.sortBy = e.target.value;
      renderApplications();
    });

    // View Switcher Tabs
    bindEvent('view-tab-grid', 'click', () => {
      state.currentView = 'grid';
      const gridBtn = document.getElementById('view-tab-grid');
      const tableBtn = document.getElementById('view-tab-table');
      if (gridBtn) gridBtn.classList.add('active');
      if (tableBtn) tableBtn.classList.remove('active');
      renderApplications();
    });

    bindEvent('view-tab-table', 'click', () => {
      state.currentView = 'table';
      const gridBtn = document.getElementById('view-tab-grid');
      const tableBtn = document.getElementById('view-tab-table');
      if (tableBtn) tableBtn.classList.add('active');
      if (gridBtn) gridBtn.classList.remove('active');
      renderApplications();
    });

    // Modal Triggers
    bindEvent('btn-add-app', 'click', () => openAppFormModal());
    bindEvent('empty-btn-add', 'click', () => openAppFormModal());

    bindEvent('btn-email-center', 'click', () => {
      const m = document.getElementById('modal-email-center');
      if (m) m.classList.add('active');
    });

    bindEvent('user-profile-pill', 'click', () => {
      const m = document.getElementById('modal-email-center');
      if (m) m.classList.add('active');
    });

    bindEvent('btn-import-excel', 'click', () => {
      const m = document.getElementById('modal-import');
      if (m) m.classList.add('active');
    });

    bindEvent('btn-sync-calendar', 'click', () => {
      const m = document.getElementById('modal-calendar-sync');
      if (m) m.classList.add('active');
    });

    bindEvent('btn-export-modal', 'click', () => {
      const m = document.getElementById('modal-export');
      if (m) m.classList.add('active');
    });

    // Calendar & Notifications
    bindEvent('btn-download-all-ics', 'click', () => {
      generateAndDownloadIcs(state.applications, 'TargetMS_Deadlines.ics');
    });

    bindEvent('btn-enable-browser-notif', 'click', () => {
      if (!('Notification' in window)) {
        alert('Desktop notifications are not supported by your browser.');
        return;
      }
      Notification.requestPermission().then(permission => {
        if (permission === 'granted') {
          alert('Browser Notifications Enabled!');
          new Notification('TargetMS Deadline Tracker', {
            body: 'Notifications activated! We will remind you of upcoming deadlines.',
            icon: '🎓'
          });
        } else {
          alert('Permission was denied for browser notifications.');
        }
      });
    });

    // Profile & Email JS Handlers
    bindEvent('btn-save-profile', 'click', () => {
      const nameEl = document.getElementById('profile-name');
      const emailEl = document.getElementById('profile-email');
      const nameVal = nameEl ? nameEl.value.trim() : '';
      const emailVal = emailEl ? emailEl.value.trim() : '';

      if (emailVal) {
        state.userProfile.name = nameVal || 'Applicant';
        state.userProfile.email = emailVal;
        saveDataToStorage();
        const dispName = document.getElementById('user-display-name');
        const dispEmail = document.getElementById('user-display-email');
        if (dispName) dispName.textContent = state.userProfile.name;
        if (dispEmail) dispEmail.textContent = state.userProfile.email;
        alert('Profile saved successfully!');
      } else {
        alert('Please enter a valid email address.');
      }
    });

    bindEvent('btn-save-emailjs', 'click', () => {
      const sEl = document.getElementById('emailjs-service-id');
      const tEl = document.getElementById('emailjs-template-id');
      const kEl = document.getElementById('emailjs-public-key');

      if (sEl) state.userProfile.emailJsService = sEl.value.trim();
      if (tEl) state.userProfile.emailJsTemplate = tEl.value.trim();
      if (kEl) state.userProfile.emailJsKey = kEl.value.trim();
      saveDataToStorage();
      alert('EmailJS settings saved!');
    });

    bindEvent('btn-send-test-email', 'click', () => {
      window.sendWeeklyTop7Digest(true);
    });

    bindEvent('toggle-auto-email', 'change', (e) => {
      state.userProfile.autoEmail = e.target.checked;
      saveDataToStorage();
    });

    // Email Dispatch Modal Actions
    bindEvent('btn-open-web-gmail', 'click', () => {
      if (!currentPreviewEmail.recipient) return;
      const su = encodeURIComponent(currentPreviewEmail.subject);
      const body = encodeURIComponent(currentPreviewEmail.body);
      const to = encodeURIComponent(currentPreviewEmail.recipient);
      const gmailUrl = `https://mail.google.com/mail/?view=cm&fs=1&to=${to}&su=${su}&body=${body}`;
      window.open(gmailUrl, '_blank');
    });

    bindEvent('btn-copy-email-text', 'click', () => {
      const fullText = `Subject: ${currentPreviewEmail.subject}\n\n${currentPreviewEmail.body}`;
      if (navigator.clipboard && navigator.clipboard.writeText) {
        navigator.clipboard.writeText(fullText).then(() => {
          alert('📋 Email digest copied to clipboard!');
        }).catch(() => copyTextFallback());
      } else {
        copyTextFallback();
      }
    });

    function copyTextFallback() {
      const copyBox = document.getElementById('preview-email-body');
      if (copyBox) {
        copyBox.select();
        document.execCommand('copy');
        alert('📋 Email digest copied to clipboard!');
      }
    }

    bindEvent('btn-open-desktop-mail', 'click', () => {
      if (!currentPreviewEmail.recipient) return;
      const su = encodeURIComponent(currentPreviewEmail.subject);
      const body = encodeURIComponent(currentPreviewEmail.body);
      window.location.href = `mailto:${currentPreviewEmail.recipient}?subject=${su}&body=${body}`;
    });

    // Export & Restore Handlers
    bindEvent('btn-export-excel', 'click', () => {
      if (!window.XLSX) { alert('SheetJS library not loaded.'); return; }
      const exportData = state.applications.map(app => ({
        University: app.university,
        Program: app.program,
        Degree: app.degree,
        Country: app.country,
        Priority: app.priority,
        Status: app.status,
        Deadline: app.deadline || app.rawDeadlineText || 'Needs Verification',
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

    bindEvent('btn-export-csv', 'click', () => {
      if (!window.XLSX) { alert('SheetJS library not loaded.'); return; }
      const exportData = state.applications.map(app => ({
        University: app.university,
        Program: app.program,
        Deadline: app.deadline || app.rawDeadlineText || 'Needs Verification',
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

    bindEvent('btn-export-json', 'click', () => {
      const dataStr = JSON.stringify(state.applications, null, 2);
      const blob = new Blob([dataStr], { type: 'application/json' });
      const link = document.createElement('a');
      link.href = URL.createObjectURL(blob);
      link.download = 'TargetMS_Backup.json';
      link.click();
    });

    bindEvent('restore-json-input', 'change', (e) => {
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
            const modalExp = document.getElementById('modal-export');
            if (modalExp) modalExp.classList.remove('active');
          }
        } catch (err) {
          alert('Invalid JSON backup file format.');
        }
      };
      reader.readAsText(file);
    });

    // Drawer Action Handlers
    bindEvent('btn-add-custom-req', 'click', () => {
      if (!state.activeAppId) return;
      const name = prompt('Enter custom requirement item name (e.g. Portfolio, Video Essay):');
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

    bindEvent('btn-save-drawer-notes', 'click', () => {
      if (!state.activeAppId) return;
      const app = state.applications.find(a => a.id === state.activeAppId);
      if (app) {
        const input = document.getElementById('drawer-notes-input');
        if (input) app.notes = input.value;
        saveDataToStorage();
        alert('Notes saved successfully!');
      }
    });

    bindEvent('drawer-btn-send-email', 'click', () => {
      if (state.activeAppId) window.sendApplicationEmail(state.activeAppId);
    });

    bindEvent('drawer-btn-gcal', 'click', () => {
      if (state.activeAppId) window.addGCalEvent(state.activeAppId);
    });

    bindEvent('drawer-btn-ics', 'click', () => {
      if (state.activeAppId) window.downloadAppIcs(state.activeAppId);
    });

    bindEvent('drawer-btn-edit', 'click', () => {
      if (!state.activeAppId) return;
      const app = state.applications.find(a => a.id === state.activeAppId);
      if (app) {
        window.closeDetailDrawer();
        window.openAppFormModal(app);
      }
    });

    bindEvent('drawer-btn-delete', 'click', () => {
      if (!state.activeAppId) return;
      const app = state.applications.find(a => a.id === state.activeAppId);
      if (!app) return;
      if (confirm(`Are you sure you want to delete ${app.university} — ${app.program}?`)) {
        state.applications = state.applications.filter(a => a.id !== state.activeAppId);
        saveDataToStorage();
        window.closeDetailDrawer();
        renderApp();
      }
    });

    // Save Application Form Listener
    bindEvent('btn-save-app', 'click', (e) => {
      e.preventDefault();
      const id = document.getElementById('form-app-id') ? document.getElementById('form-app-id').value : '';
      const univEl = document.getElementById('form-univ');
      const progEl = document.getElementById('form-program');
      const univ = univEl ? univEl.value.trim() : '';
      const prog = progEl ? progEl.value.trim() : '';
      const deadline = document.getElementById('form-deadline') ? document.getElementById('form-deadline').value : '';

      if (!univ || !prog) {
        alert('Please fill in required fields (University and Program).');
        return;
      }

      const verificationStatus = deadline ? 'Verified' : 'Needs Verification';

      if (id) {
        const app = state.applications.find(a => a.id === id);
        if (app) {
          app.university = univ;
          app.program = prog;
          app.degree = document.getElementById('form-degree') ? document.getElementById('form-degree').value.trim() : 'MS';
          app.country = document.getElementById('form-country') ? document.getElementById('form-country').value.trim() : 'USA';
          app.priority = document.getElementById('form-priority') ? document.getElementById('form-priority').value : 'Medium';
          app.status = document.getElementById('form-status') ? document.getElementById('form-status').value : 'Documents Pending';
          app.deadline = deadline;
          app.deadlineType = document.getElementById('form-deadline-type') ? document.getElementById('form-deadline-type').value : 'Regular Round';
          app.verificationStatus = (document.getElementById('form-verification') && document.getElementById('form-verification').value) || verificationStatus;
          app.officialSourceUrl = document.getElementById('form-source-url') ? document.getElementById('form-source-url').value.trim() : '';
          app.portalUrl = document.getElementById('form-portal-url') ? document.getElementById('form-portal-url').value.trim() : '';
          app.greRequirement = document.getElementById('form-gre') ? document.getElementById('form-gre').value : 'Optional';
          app.englishRequirement = document.getElementById('form-english') ? document.getElementById('form-english').value.trim() : '';
          app.appFee = document.getElementById('form-fee') ? (parseFloat(document.getElementById('form-fee').value) || 0) : 0;
          app.notes = document.getElementById('form-notes') ? document.getElementById('form-notes').value.trim() : '';
        }
      } else {
        const newApp = {
          id: 'app_user_' + Date.now(),
          university: univ,
          program: prog,
          degree: (document.getElementById('form-degree') && document.getElementById('form-degree').value.trim()) || 'MS',
          country: (document.getElementById('form-country') && document.getElementById('form-country').value.trim()) || 'USA',
          priority: document.getElementById('form-priority') ? document.getElementById('form-priority').value : 'Medium',
          status: document.getElementById('form-status') ? document.getElementById('form-status').value : 'Documents Pending',
          deadline: deadline,
          rawDeadlineText: deadline,
          openingDate: '2026-09-01',
          deadlineType: document.getElementById('form-deadline-type') ? document.getElementById('form-deadline-type').value : 'Regular Round',
          verificationStatus: verificationStatus,
          officialSourceUrl: document.getElementById('form-source-url') ? document.getElementById('form-source-url').value.trim() : '',
          portalUrl: document.getElementById('form-portal-url') ? document.getElementById('form-portal-url').value.trim() : '',
          greRequirement: document.getElementById('form-gre') ? document.getElementById('form-gre').value : 'Optional',
          englishRequirement: document.getElementById('form-english') ? document.getElementById('form-english').value.trim() : '',
          appFee: document.getElementById('form-fee') ? (parseFloat(document.getElementById('form-fee').value) || 0) : 0,
          notes: document.getElementById('form-notes') ? document.getElementById('form-notes').value.trim() : '',
          checklist: DEFAULT_CHECKLIST_ITEMS.map(item => ({ ...item, status: 'Not Started' }))
        };
        state.applications.unshift(newApp);
      }

      saveDataToStorage();
      const modalApp = document.getElementById('modal-app-form');
      if (modalApp) modalApp.classList.remove('active');
      renderApp();
    });

    // Close Modals & Drawers
    document.querySelectorAll('.close-modal').forEach(btn => {
      btn.addEventListener('click', () => window.closeAllModals());
    });

    document.querySelectorAll('.close-drawer').forEach(btn => {
      btn.addEventListener('click', () => window.closeDetailDrawer());
    });

    // Click backdrop overlay to close
    document.querySelectorAll('.modal-overlay').forEach(overlay => {
      overlay.addEventListener('click', (e) => {
        if (e.target === overlay) overlay.classList.remove('active');
      });
    });

    document.querySelectorAll('.drawer-overlay').forEach(drawer => {
      drawer.addEventListener('click', (e) => {
        if (e.target === drawer) drawer.classList.remove('active');
      });
    });


    // Drag and drop Excel Upload
    const dropZone = document.getElementById('drop-zone');
    const fileInput = document.getElementById('file-input-excel');
    const confirmBtn = document.getElementById('btn-confirm-import');

    let loadedWorkbook = null;

    if (dropZone && fileInput) {
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
    }

    function handleExcelFile(file) {
      const reader = new FileReader();
      reader.onload = (event) => {
        const data = new Uint8Array(event.target.result);
        loadedWorkbook = XLSX.read(data, { type: 'array' });
        
        const previewBox = document.getElementById('import-preview-box');
        const filenameEl = document.getElementById('preview-filename');
        if (previewBox) previewBox.style.display = 'block';
        if (filenameEl) filenameEl.textContent = `File Loaded: ${file.name}`;
        if (confirmBtn) confirmBtn.disabled = false;
      };
      reader.readAsArrayBuffer(file);
    }

    if (confirmBtn) {
      confirmBtn.addEventListener('click', () => {
        if (loadedWorkbook) {
          parseAndMergeWorkbook(loadedWorkbook, true, true);
          const m = document.getElementById('modal-import');
          if (m) m.classList.remove('active');
        }
      });
    }
  }

  // Launch on DOM ready
  document.addEventListener('DOMContentLoaded', init);

})();

