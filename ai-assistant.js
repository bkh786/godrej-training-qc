/* ============================================================================
   GODREJ AI COPILOT & LEFT-SIDEBAR FILTER CONTROLLER
   Live SharePoint Datasource-Connected Conversational Analyst (All 5 Pages)
   ============================================================================ */

/* ----------------------------------------------------------------------------
   1. LEFT-SIDEBAR FILTER CONTROLLER (Shared across all 5 Dashboard Pages)
   ---------------------------------------------------------------------------- */
window.toggleFiltersSidebar = function (forceOpen) {
  const sidebar = document.getElementById('filterContainer');
  const backdrop = document.getElementById('filtersSidebarBackdrop');
  if (!sidebar) return;
  const shouldOpen = typeof forceOpen === 'boolean' ? forceOpen : !sidebar.classList.contains('open');
  sidebar.classList.toggle('open', shouldOpen);
  if (backdrop) backdrop.classList.toggle('open', shouldOpen);
};

window.updateFiltersSidebarBadge = function () {
  const badge = document.getElementById('filtersToggleBadge');
  const sidebar = document.getElementById('filterContainer');
  if (!badge || !sidebar) return;

  let activeCount = 0;

  // Count active <select> filters inside #filterContainer
  const selects = sidebar.querySelectorAll('select');
  selects.forEach((sel) => {
    const val = (sel.value || '').trim();
    if (!val || val.toLowerCase() === 'all' || val.toLowerCase() === 'loading...') return;
    // On VM & Branch pages, filterMonth always has a selected month; count it only if user switched away from the first option
    if (sel.id === 'filterMonth') {
      if (sel.selectedIndex > 0) activeCount++;
      return;
    }
    activeCount++;
  });

  // Count active <input type="date"> filters inside #filterContainer
  const dateInputs = sidebar.querySelectorAll('input[type="date"]');
  dateInputs.forEach((inp) => {
    const val = (inp.value || '').trim();
    if (!val) return;
    // Check if it matches default full date bounds on m_score or qc
    if (window.State && window.State.dateBounds) {
      if (inp.id === 'fDateFrom' && val === window.State.dateBounds.min) return;
      if (inp.id === 'fDateTo' && val === window.State.dateBounds.max) return;
    }
    activeCount++;
  });

  badge.textContent = String(activeCount);
  badge.classList.toggle('visible', activeCount > 0);
};

(function initSidebarListeners() {
  function bindSidebarEvents() {
    const sidebar = document.getElementById('filterContainer');
    if (!sidebar) return;

    sidebar.addEventListener('change', () => {
      setTimeout(window.updateFiltersSidebarBadge, 30);
    });
    sidebar.addEventListener('input', () => {
      setTimeout(window.updateFiltersSidebarBadge, 30);
    });

    const resetBtns = document.querySelectorAll('#btnFilterReset, #btnResetFilters, .compact-filter-reset');
    resetBtns.forEach((btn) => {
      btn.addEventListener('click', () => {
        setTimeout(window.updateFiltersSidebarBadge, 60);
      });
    });

    const chips = document.getElementById('activeChips');
    if (chips && typeof MutationObserver !== 'undefined') {
      const obs = new MutationObserver(() => {
        window.updateFiltersSidebarBadge();
      });
      obs.observe(chips, { childList: true, subtree: true });
    }

    const filterBar = document.getElementById('filterBar');
    if (filterBar && typeof MutationObserver !== 'undefined') {
      const obs2 = new MutationObserver(() => {
        window.updateFiltersSidebarBadge();
      });
      obs2.observe(filterBar, { childList: true, subtree: true });
    }

    window.updateFiltersSidebarBadge();
  }

  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') {
      window.toggleFiltersSidebar(false);
    }
  });

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', bindSidebarEvents);
  } else {
    bindSidebarEvents();
  }
})();

/* ----------------------------------------------------------------------------
   2. GODREJ AI COPILOT — LIVE SHAREPOINT DATASOURCE CONNECTOR & ENGINE
   ---------------------------------------------------------------------------- */
(function () {
  if (window.__godrejAiCopilotInitialized) return;
  window.__godrejAiCopilotInitialized = true;

  const state = {
    isOpen: false,
    isExpanded: false,
    isLoading: false,
    messages: [],
    lastSyncTimestamp: null,
    datasourceStatus: 'Connecting to SharePoint...'
  };

  function detectPageType() {
    const path = (window.location.pathname || '').toLowerCase();
    const title = (document.title || '').toLowerCase();
    if (path.includes('program_performance') || title.includes('vm performance')) return 'program';
    if (path.includes('branch') || title.includes('branch performance')) return 'branch';
    if (path.includes('m_score') || title.includes('merchandising score')) return 'm_score';
    if (path.includes('training') || title.includes('training performance')) return 'training';
    return 'qc';
  }

  function getPageLabel(pageType) {
    switch (pageType) {
      case 'program': return 'VM Performance';
      case 'branch': return 'Branch Performance';
      case 'm_score': return 'Merchandising Score';
      case 'training': return 'Training Performance';
      default: return 'QC Performance';
    }
  }

  function getSuggestions(pageType) {
    switch (pageType) {
      case 'program':
        return [
          'Summarize overall VM program coverage & visits',
          'Rank AOMs by coverage & visit achievement %',
          'Who are the top 5 and bottom 5 VMs?',
          'Which states have the highest pending coverage?',
          'Show product display & POSM deployment summary'
        ];
      case 'branch':
        return [
          'Summarize overall branch coverage & visit achievement',
          'Compare Zone-wise coverage and visit performance',
          'Which branches have the highest and lowest achievement?',
          'Rank AOMs and TMMs by branch coverage %',
          'Break down Cat A vs Cat B store coverage'
        ];
      case 'm_score':
        return [
          'What is the overall Merchandising Score & store count?',
          'Compare AOM-wise average Merchandising Score %',
          'Which branches and VMs have the highest M-Score?',
          'Show product category-wise M-Score breakdown',
          'List stores with the lowest Merchandising Scores'
        ];
      case 'training':
        return [
          'Summarize training attendance, pre-test & post-test scores',
          'Compare AOM-wise knowledge gain & pass rate',
          'Which employees have not attempted the post-test?',
          'Compare training categories by post-test score %',
          'Which branches have the highest training completion?'
        ];
      default:
        return [
          'Summarize overall QC score, total audits & rejections',
          'Which QC parameters have the highest rejections?',
          'Compare AOM-wise & Branch-wise QC score %',
          'List stores that scored below 100% with rejection reasons',
          'Who are the top performing VMs in QC audits?'
        ];
    }
  }

  /* --------------------------------------------------------------------------
     LIVE SHAREPOINT DATASOURCE ACCESSORS
     -------------------------------------------------------------------------- */
  function getAllRecords(pageType) {
    try {
      if (pageType === 'program' || pageType === 'branch') {
        if (typeof State !== 'undefined' && Array.isArray(State.records)) return State.records;
        if (window.State && Array.isArray(window.State.records)) return window.State.records;
      } else {
        if (typeof State !== 'undefined' && State.raw && Array.isArray(State.raw.records)) return State.raw.records;
        if (window.State && window.State.raw && Array.isArray(window.State.raw.records)) return window.State.raw.records;
      }
    } catch (e) {}
    return [];
  }

  function getActiveRecords(pageType) {
    try {
      if (pageType === 'program' || pageType === 'branch') {
        if (typeof App !== 'undefined' && typeof App.getFilteredRecords === 'function') return App.getFilteredRecords();
        if (window.App && typeof window.App.getFilteredRecords === 'function') return window.App.getFilteredRecords();
      } else {
        if (typeof getFilteredRecords === 'function') return getFilteredRecords();
        if (typeof window.getFilteredRecords === 'function') return window.getFilteredRecords();
      }
    } catch (e) {}
    return getAllRecords(pageType);
  }

  function getActiveFilterSummary() {
    const parts = [];
    const sidebar = document.getElementById('filterContainer');
    if (sidebar) {
      const selects = sidebar.querySelectorAll('select');
      selects.forEach((sel) => {
        const val = (sel.value || '').trim();
        if (!val || val.toLowerCase() === 'all' || val.toLowerCase() === 'loading...') return;
        const labelEl = sel.closest('.compact-filter-item, .filter-field')?.querySelector('.compact-filter-label, label');
        const label = labelEl ? labelEl.textContent.trim().replace(/\s+/g, ' ') : sel.id.replace(/^filter|^f/, '');
        const text = sel.options[sel.selectedIndex] ? sel.options[sel.selectedIndex].text.trim() : val;
        parts.push(`${label}: ${text}`);
      });
      const dates = sidebar.querySelectorAll('input[type="date"]');
      dates.forEach((inp) => {
        const val = (inp.value || '').trim();
        if (!val) return;
        if (typeof State !== 'undefined' && State.dateBounds) {
          if (inp.id === 'fDateFrom' && val === State.dateBounds.min) return;
          if (inp.id === 'fDateTo' && val === State.dateBounds.max) return;
        }
        const label = inp.id === 'fDateFrom' ? 'From' : 'To';
        parts.push(`${label}: ${val}`);
      });
    }
    return parts.length ? parts.join(' · ') : 'All Data (No Filters)';
  }

  function getDatasourceUpdatedText() {
    const dataUpdatedEl = document.getElementById('dataUpdatedTillMeta');
    const lastSyncEl = document.getElementById('lastUpdatedMeta');
    const sourceEl = document.getElementById('dataSourceLabel');
    const pieces = [];
    if (dataUpdatedEl && dataUpdatedEl.textContent && dataUpdatedEl.textContent.trim() !== '--') {
      const du = dataUpdatedEl.textContent.trim();
      pieces.push(du.toLowerCase().startsWith('data') ? du : 'Data Upto ' + du);
    }
    if (lastSyncEl && lastSyncEl.textContent && !lastSyncEl.textContent.toLowerCase().includes('connecting')) {
      const ls = lastSyncEl.textContent.trim();
      pieces.push(ls.toLowerCase().startsWith('synced') ? ls : 'Synced ' + ls);
    } else if (sourceEl && sourceEl.textContent) {
      pieces.push(sourceEl.textContent.trim());
    }
    return pieces.length ? pieces.join(' · ') : 'SharePoint Live Connected';
  }

  /* --------------------------------------------------------------------------
     REAL-TIME PIPE-TABLE COMPILER FOR ALL 5 PAGES
     -------------------------------------------------------------------------- */
  function buildLiveClientPipeTable(pageType) {
    const allRecs = getAllRecords(pageType);
    const activeRecs = getActiveRecords(pageType);
    const filterSummary = getActiveFilterSummary();
    const updatedMeta = getDatasourceUpdatedText();

    if (!allRecs || allRecs.length === 0) {
      return `PAGE: ${getPageLabel(pageType)}\nSTATUS: SharePoint dataset loading (0 records in memory).`;
    }

    if (pageType === 'program') {
      return buildProgramPipeTable(allRecs, activeRecs, filterSummary, updatedMeta);
    } else if (pageType === 'branch') {
      return buildBranchPipeTable(allRecs, activeRecs, filterSummary, updatedMeta);
    } else if (pageType === 'm_score') {
      return buildMScorePipeTable(allRecs, activeRecs, filterSummary, updatedMeta);
    } else if (pageType === 'training') {
      return buildTrainingPipeTable(allRecs, activeRecs, filterSummary, updatedMeta);
    } else {
      return buildQCPipeTable(allRecs, activeRecs, filterSummary, updatedMeta);
    }
  }

  function pct(num, den) {
    if (!den || den <= 0) return '0.0%';
    return ((num / den) * 100).toFixed(1) + '%';
  }

  /* --- 1. VM PERFORMANCE (program_performance.html) --- */
  function buildProgramPipeTable(allRecs, activeRecs, filterSummary, updatedMeta) {
    const recs = activeRecs.length ? activeRecs : allRecs;
    const sum = (arr, fn) => arr.reduce((acc, r) => acc + (Number(fn(r)) || 0), 0);

    const covTgt = sum(recs, r => r.coverageTarget);
    const covAch = sum(recs, r => r.coverageAch);
    const visTgt = sum(recs, r => r.visitsTarget);
    const visAch = sum(recs, r => r.visitAch);
    const prodVis = sum(recs, r => r.productiveVisits);
    const disp = sum(recs, r => r.productDisplay);
    const posm = sum(recs, r => r.posmDeployment);
    const manAch = sum(recs, r => r.manDaysAchieved);
    const manTgt = sum(recs, r => r.targetManDays);
    const eboTgt = sum(recs, r => r.targetedEBO);
    const eboAch = sum(recs, r => r.eboCoverage);
    const chainTgt = sum(recs, r => r.targetedChain);
    const chainAch = sum(recs, r => r.chainCoverage);
    const coolTgt = sum(recs, r => r.targetedCool);
    const coolAch = sum(recs, r => r.coolCoverage);
    const othTgt = sum(recs, r => r.targetedOther);
    const othAch = sum(recs, r => r.otherCoverage);
    const avgFinalAch = recs.length ? (sum(recs, r => r.finalAch) / recs.length) * 100 : 0;

    // AOM Breakdown
    const aomMap = {};
    recs.forEach(r => {
      const k = r.aom || 'Unassigned';
      if (!aomMap[k]) aomMap[k] = { vms: 0, covTgt: 0, covAch: 0, visTgt: 0, visAch: 0, prodVis: 0, disp: 0, posm: 0, achSum: 0 };
      const g = aomMap[k];
      g.vms++;
      g.covTgt += r.coverageTarget || 0;
      g.covAch += r.coverageAch || 0;
      g.visTgt += r.visitsTarget || 0;
      g.visAch += r.visitAch || 0;
      g.prodVis += r.productiveVisits || 0;
      g.disp += r.productDisplay || 0;
      g.posm += r.posmDeployment || 0;
      g.achSum += (r.finalAch || 0) * 100;
    });
    const aomLines = ['AOM | VMs | CovTgt | CovAch | Cov% | VisTgt | VisAch | Vis% | ProdVis | Prod% | Display | POSM | AvgFinalAch%'];
    Object.entries(aomMap)
      .sort((a, b) => (b[1].covTgt ? b[1].covAch / b[1].covTgt : 0) - (a[1].covTgt ? a[1].covAch / a[1].covTgt : 0))
      .forEach(([aom, g]) => {
        aomLines.push(`${aom} | ${g.vms} | ${g.covTgt} | ${g.covAch} | ${pct(g.covAch, g.covTgt)} | ${g.visTgt} | ${g.visAch} | ${pct(g.visAch, g.visTgt)} | ${g.prodVis} | ${pct(g.prodVis, g.visAch)} | ${g.disp} | ${g.posm} | ${(g.achSum / g.vms).toFixed(1)}%`);
      });

    // State Breakdown
    const stateMap = {};
    recs.forEach(r => {
      const k = r.state || 'Unknown';
      if (!stateMap[k]) stateMap[k] = { vms: 0, covTgt: 0, covAch: 0, visTgt: 0, visAch: 0, achSum: 0 };
      const g = stateMap[k];
      g.vms++;
      g.covTgt += r.coverageTarget || 0;
      g.covAch += r.coverageAch || 0;
      g.visTgt += r.visitsTarget || 0;
      g.visAch += r.visitAch || 0;
      g.achSum += (r.finalAch || 0) * 100;
    });
    const stateLines = ['State | VMs | CovTgt | CovAch | CovPending | Cov% | VisTgt | VisAch | Vis% | AvgFinalAch%'];
    Object.entries(stateMap)
      .sort((a, b) => (b[1].covTgt - b[1].covAch) - (a[1].covTgt - a[1].covAch))
      .forEach(([st, g]) => {
        stateLines.push(`${st} | ${g.vms} | ${g.covTgt} | ${g.covAch} | ${Math.max(0, g.covTgt - g.covAch)} | ${pct(g.covAch, g.covTgt)} | ${g.visTgt} | ${g.visAch} | ${pct(g.visAch, g.visTgt)} | ${(g.achSum / g.vms).toFixed(1)}%`);
      });

    // Complete VM Roster (Sorted by FinalAch% desc)
    const sortedVMs = [...recs].sort((a, b) => (b.finalAch || 0) - (a.finalAch || 0));
    const vmLines = ['Rank | VM Name | EmpCode | AOM | Branch | State | CovTgt | CovAch | Cov% | VisTgt | VisAch | Vis% | ProdVis | Display | POSM | ManDays | FinalAch%'];
    sortedVMs.forEach((r, idx) => {
      vmLines.push(`${idx + 1} | ${r.userName || '-'} | ${r.userCode || '-'} | ${r.aom || '-'} | ${r.branch || '-'} | ${r.state || '-'} | ${r.coverageTarget || 0} | ${r.coverageAch || 0} | ${pct(r.coverageAch, r.coverageTarget)} | ${r.visitsTarget || 0} | ${r.visitAch || 0} | ${pct(r.visitAch, r.visitsTarget)} | ${r.productiveVisits || 0} | ${r.productDisplay || 0} | ${r.posmDeployment || 0} | ${r.manDaysAchieved || 0}/${r.targetManDays || 0} | ${((r.finalAch || 0) * 100).toFixed(1)}%`);
    });

    return [
      `=== GODREJ VM PERFORMANCE LIVE SHAREPOINT DATASOURCE ===`,
      `Status: ${updatedMeta} | Active Filters: ${filterSummary}`,
      `Active Scope: ${recs.length} VMs | Overall Coverage: ${covAch}/${covTgt} (${pct(covAch, covTgt)}, Pending: ${Math.max(0, covTgt - covAch)}) | Total Visits: ${visAch}/${visTgt} (${pct(visAch, visTgt)}) | Productive Visits: ${prodVis} (${pct(prodVis, visAch)}) | Product Display: ${disp} | POSM Deployed: ${posm} | Man-Days: ${manAch}/${manTgt} (${pct(manAch, manTgt)}) | Avg Final Achievement: ${avgFinalAch.toFixed(1)}%`,
      `Store Category Coverage: EBO: ${eboAch}/${eboTgt} (${pct(eboAch, eboTgt)}) | Chain Stores: ${chainAch}/${chainTgt} (${pct(chainAch, chainTgt)}) | Cool Studio: ${coolAch}/${coolTgt} (${pct(coolAch, coolTgt)}) | Others: ${othAch}/${othTgt} (${pct(othAch, othTgt)})`,
      ``,
      `--- AOM-WISE VM PERFORMANCE SUMMARY ---`,
      aomLines.join('\n'),
      ``,
      `--- STATE-WISE VM PERFORMANCE SUMMARY ---`,
      stateLines.join('\n'),
      ``,
      `--- COMPLETE FIELD VM ROSTER (${sortedVMs.length} VMs) ---`,
      vmLines.join('\n')
    ].join('\n');
  }

  /* --- 2. BRANCH PERFORMANCE (branch.html) --- */
  function buildBranchPipeTable(allRecs, activeRecs, filterSummary, updatedMeta) {
    const recs = activeRecs.length ? activeRecs : allRecs;
    const sum = (arr, fn) => arr.reduce((acc, r) => acc + (Number(fn(r)) || 0), 0);

    const covTgt = sum(recs, r => r.coverageTarget);
    const covAch = sum(recs, r => r.coverageAch);
    const visTgt = sum(recs, r => r.visitsTarget);
    const visAch = sum(recs, r => r.visitAch);
    const prodVis = sum(recs, r => r.productiveVisits);
    const eboTgt = sum(recs, r => r.targetedEBO);
    const eboAch = sum(recs, r => r.eboCoverage);
    const chainTgt = sum(recs, r => r.targetedChain);
    const chainAch = sum(recs, r => r.chainCoverage);
    const coolTgt = sum(recs, r => r.targetedCool);
    const coolAch = sum(recs, r => r.coolCoverage);
    const othTgt = sum(recs, r => r.targetedOther);
    const othAch = sum(recs, r => r.otherCoverage);
    const avgAch = recs.length ? (sum(recs, r => r.finalAch) / recs.length) * 100 : 0;

    // Zone Breakdown
    const zoneMap = {};
    recs.forEach(r => {
      const k = r.zone || 'Unknown';
      if (!zoneMap[k]) zoneMap[k] = { branches: 0, covTgt: 0, covAch: 0, visTgt: 0, visAch: 0, prodVis: 0, achSum: 0 };
      const g = zoneMap[k];
      g.branches++;
      g.covTgt += r.coverageTarget || 0;
      g.covAch += r.coverageAch || 0;
      g.visTgt += r.visitsTarget || 0;
      g.visAch += r.visitAch || 0;
      g.prodVis += r.productiveVisits || 0;
      g.achSum += (r.finalAch || 0) * 100;
    });
    const zoneLines = ['Zone | Branches | CovTgt | CovAch | Cov% | VisTgt | VisAch | Vis% | ProdVis | Prod% | AvgAch%'];
    Object.entries(zoneMap)
      .sort((a, b) => (b[1].covTgt ? b[1].covAch / b[1].covTgt : 0) - (a[1].covTgt ? a[1].covAch / a[1].covTgt : 0))
      .forEach(([zn, g]) => {
        zoneLines.push(`${zn} | ${g.branches} | ${g.covTgt} | ${g.covAch} | ${pct(g.covAch, g.covTgt)} | ${g.visTgt} | ${g.visAch} | ${pct(g.visAch, g.visTgt)} | ${g.prodVis} | ${pct(g.prodVis, g.visAch)} | ${(g.achSum / g.branches).toFixed(1)}%`);
      });

    // AOM Breakdown
    const aomMap = {};
    recs.forEach(r => {
      const k = r.aom || 'Unassigned';
      if (!aomMap[k]) aomMap[k] = { branches: 0, covTgt: 0, covAch: 0, visTgt: 0, visAch: 0, achSum: 0 };
      const g = aomMap[k];
      g.branches++;
      g.covTgt += r.coverageTarget || 0;
      g.covAch += r.coverageAch || 0;
      g.visTgt += r.visitsTarget || 0;
      g.visAch += r.visitAch || 0;
      g.achSum += (r.finalAch || 0) * 100;
    });
    const aomLines = ['AOM | Branches | CovTgt | CovAch | Cov% | VisTgt | VisAch | Vis% | AvgAch%'];
    Object.entries(aomMap)
      .sort((a, b) => (b[1].covTgt ? b[1].covAch / b[1].covTgt : 0) - (a[1].covTgt ? a[1].covAch / a[1].covTgt : 0))
      .forEach(([aom, g]) => {
        aomLines.push(`${aom} | ${g.branches} | ${g.covTgt} | ${g.covAch} | ${pct(g.covAch, g.covTgt)} | ${g.visTgt} | ${g.visAch} | ${pct(g.visAch, g.visTgt)} | ${(g.achSum / g.branches).toFixed(1)}%`);
      });

    // Branch Roster
    const sortedBranches = [...recs].sort((a, b) => (b.finalAch || 0) - (a.finalAch || 0));
    const branchLines = ['Rank | Branch | Zone | State | TMM | AOM | CovTgt | CovAch | Cov% | VisTgt | VisAch | Vis% | ProdVis | EBOCov% | ChainCov% | FinalAch%'];
    sortedBranches.forEach((r, idx) => {
      branchLines.push(`${idx + 1} | ${r.branch} | ${r.zone || '-'} | ${r.state || '-'} | ${r.tmm || '-'} | ${r.aom || '-'} | ${r.coverageTarget || 0} | ${r.coverageAch || 0} | ${pct(r.coverageAch, r.coverageTarget)} | ${r.visitsTarget || 0} | ${r.visitAch || 0} | ${pct(r.visitAch, r.visitsTarget)} | ${r.productiveVisits || 0} | ${pct(r.eboCoverage, r.targetedEBO)} | ${pct(r.chainCoverage, r.targetedChain)} | ${((r.finalAch || 0) * 100).toFixed(1)}%`);
    });

    return [
      `=== GODREJ BRANCH PERFORMANCE LIVE SHAREPOINT DATASOURCE ===`,
      `Status: ${updatedMeta} | Active Filters: ${filterSummary}`,
      `Active Scope: ${recs.length} Branches | Coverage: ${covAch}/${covTgt} (${pct(covAch, covTgt)}, Pending: ${Math.max(0, covTgt - covAch)}) | Visits: ${visAch}/${visTgt} (${pct(visAch, visTgt)}) | Productive Visits: ${prodVis} (${pct(prodVis, visAch)}) | EBO Coverage: ${eboAch}/${eboTgt} (${pct(eboAch, eboTgt)}) | Chain Coverage: ${chainAch}/${chainTgt} (${pct(chainAch, chainTgt)}) | Cool Studio: ${coolAch}/${coolTgt} (${pct(coolAch, coolTgt)}) | Others: ${othAch}/${othTgt} (${pct(othAch, othTgt)}) | Avg Final Achievement: ${avgAch.toFixed(1)}%`,
      ``,
      `--- ZONE-WISE BRANCH PERFORMANCE ---`,
      zoneLines.join('\n'),
      ``,
      `--- AOM-WISE BRANCH PERFORMANCE ---`,
      aomLines.join('\n'),
      ``,
      `--- COMPLETE BRANCH ROSTER (${sortedBranches.length} Branches) ---`,
      branchLines.join('\n')
    ].join('\n');
  }

  /* --- 3. MERCHANDISING SCORE (m_score.html) --- */
  function buildMScorePipeTable(allRecs, activeRecs, filterSummary, updatedMeta) {
    const recs = activeRecs.length ? activeRecs : allRecs;
    let totDeployed = 0, totTarget = 0, pctSum = 0;
    const stores = new Set(), vms = new Set(), branches = new Set();
    let exc = 0, good = 0, warn = 0, low = 0;

    const aomMap = {}, branchMap = {}, prodCatMap = {}, storeCatMap = {}, vmMap = {};

    recs.forEach(r => {
      const sc = (Number(r.mScore) || 0) * 100;
      const obt = Number(r.finalDeployed) || 0;
      const mx = Number(r.targetDeployment) || 0;
      totDeployed += obt;
      totTarget += mx;
      pctSum += sc;
      if (r.storeCode || r.storeName) stores.add(r.storeCode || r.storeName);
      if (r.vmName) vms.add(r.vmName);
      if (r.branch) branches.add(r.branch);

      if (sc >= 85) exc++;
      else if (sc >= 70) good++;
      else if (sc >= 50) warn++;
      else low++;

      const addGrp = (map, key, extra) => {
        const k = key || 'Unknown';
        if (!map[k]) map[k] = { count: 0, pctSum: 0, obt: 0, mx: 0, stores: new Set(), extra: extra || '' };
        map[k].count++;
        map[k].pctSum += sc;
        map[k].obt += obt;
        map[k].mx += mx;
        if (r.storeCode || r.storeName) map[k].stores.add(r.storeCode || r.storeName);
      };

      addGrp(aomMap, r.aomName);
      addGrp(branchMap, r.branch, r.aomName);
      addGrp(prodCatMap, r.prodCat);
      addGrp(storeCatMap, r.storeCat);
      addGrp(vmMap, r.vmName, `${r.aomName || '-'} | ${r.branch || '-'}`);
    });

    const avgScore = recs.length ? (pctSum / recs.length).toFixed(1) : '0.0';
    const weightedScore = totTarget > 0 ? ((totDeployed / totTarget) * 100).toFixed(1) : avgScore;

    const fmtGrp = (map, header, limit = 40) => {
      const lines = [header];
      Object.entries(map)
        .sort((a, b) => (b[1].pctSum / b[1].count) - (a[1].pctSum / a[1].count))
        .slice(0, limit)
        .forEach(([k, g]) => {
          const avg = (g.pctSum / g.count).toFixed(1) + '%';
          lines.push(`${k}${g.extra ? ' | ' + g.extra : ''} | ${g.count} | ${g.stores.size} | ${g.obt}/${g.mx} | ${avg}`);
        });
      return lines.join('\n');
    };

    const lowStores = [...recs]
      .sort((a, b) => (a.mScore || 0) - (b.mScore || 0))
      .slice(0, 35)
      .map(r => `${r.visitDate || '-'} | ${r.storeName || '-'} (${r.storeCode || '-'}) | ${r.branch || '-'} | ${r.aomName || '-'} | ${r.vmName || '-'} | ${r.prodCat || '-'} | ${r.prodName || '-'} | ${((r.mScore || 0) * 100).toFixed(1)}% (${r.finalDeployed || 0}/${r.targetDeployment || 0})`);

    return [
      `=== GODREJ MERCHANDISING SCORE (M-SCORE) LIVE SHAREPOINT DATASOURCE ===`,
      `Status: ${updatedMeta} | Active Filters: ${filterSummary}`,
      `Active Scope: ${recs.length} Product-Store Audits | Unique Stores: ${stores.size} | Unique VMs: ${vms.size} | Branches: ${branches.size} | Avg M-Score: ${avgScore}% (Weighted Deployed/Target: ${totDeployed}/${totTarget} = ${weightedScore}%) | Score Bands: Excellent(>=85%): ${exc}, Good(70-84%): ${good}, Warning(50-69%): ${warn}, Critical(<50%): ${low}`,
      ``,
      `--- AOM-WISE M-SCORE SUMMARY ---`,
      fmtGrp(aomMap, 'AOM | Audits | UniqueStores | Deployed/Target | AvgMScore%'),
      ``,
      `--- PRODUCT CATEGORY & STORE CATEGORY M-SCORE ---`,
      fmtGrp(prodCatMap, 'ProductCategory | Audits | UniqueStores | Deployed/Target | AvgMScore%'),
      fmtGrp(storeCatMap, 'StoreCategory | Audits | UniqueStores | Deployed/Target | AvgMScore%'),
      ``,
      `--- BRANCH-WISE M-SCORE SUMMARY ---`,
      fmtGrp(branchMap, 'Branch | AOM | Audits | UniqueStores | Deployed/Target | AvgMScore%', 35),
      ``,
      `--- VM-WISE M-SCORE SUMMARY (TOP & BOTTOM) ---`,
      fmtGrp(vmMap, 'VM Name | AOM | Branch | Audits | UniqueStores | Deployed/Target | AvgMScore%', 80),
      ``,
      `--- LOWEST SCORING STORE AUDITS (BOTTOM 35) ---`,
      `Date | Store | Branch | AOM | VM | ProdCat | Product | M-Score%`,
      ...lowStores
    ].join('\n');
  }

  /* --- 4. QC PERFORMANCE (index.html) --- */
  function buildQCPipeTable(allRecs, activeRecs, filterSummary, updatedMeta) {
    const recs = activeRecs.length ? activeRecs : allRecs;
    let scoreSum = 0, rejAudits = 0, paramRejTotal = 0, perfect = 0;
    const stores = new Set(), vms = new Set();
    const paramRej = {}, aomMap = {}, branchMap = {}, vmMap = {}, catMap = {};

    recs.forEach(r => {
      const sc = (Number(r.qcPct) || 0) * 100;
      const isRej = Boolean(r.isRejection) || sc < 99.9;
      scoreSum += sc;
      if (isRej) rejAudits++;
      else perfect++;
      if (r.outletCode || r.store) stores.add(r.outletCode || r.store);
      if (r.empName) vms.add(r.empName);

      let rowRejCount = 0;
      if (Array.isArray(r.display)) {
        r.display.forEach(p => {
          if (p && !p.ok) {
            paramRej[p.name] = (paramRej[p.name] || 0) + 1;
            paramRejTotal++;
            rowRejCount++;
          }
        });
      }
      if (Array.isArray(r.campaign)) {
        r.campaign.forEach(p => {
          if (p && !p.ok) {
            paramRej[p.name] = (paramRej[p.name] || 0) + 1;
            paramRejTotal++;
            rowRejCount++;
          }
        });
      }

      const addG = (map, key, extra) => {
        const k = key || 'Unknown';
        if (!map[k]) map[k] = { audits: 0, scSum: 0, rej: 0, paramFails: 0, extra: extra || '' };
        map[k].audits++;
        map[k].scSum += sc;
        if (isRej) map[k].rej++;
        map[k].paramFails += rowRejCount;
      };

      addG(aomMap, r.aom);
      addG(branchMap, r.branch, r.aom);
      addG(vmMap, r.empName, `${r.aom || '-'} | ${r.branch || '-'}`);
      addG(catMap, `${r.category || '-'} (${r.classification || '-'})`);
    });

    const avgQC = recs.length ? (scoreSum / recs.length).toFixed(1) : '0.0';

    const paramLines = ['QC Parameter | FailedAudits | ShareOfFailures%'];
    Object.entries(paramRej)
      .sort((a, b) => b[1] - a[1])
      .forEach(([p, c]) => {
        paramLines.push(`${p} | ${c} | ${pct(c, paramRejTotal)}`);
      });

    const fmtQCGrp = (map, hdr, limit = 60) => {
      const lines = [hdr];
      Object.entries(map)
        .sort((a, b) => (b[1].scSum / b[1].audits) - (a[1].scSum / a[1].audits))
        .slice(0, limit)
        .forEach(([k, g]) => {
          lines.push(`${k}${g.extra ? ' | ' + g.extra : ''} | ${g.audits} | ${(g.scSum / g.audits).toFixed(1)}% | ${g.rej} | ${g.paramFails}`);
        });
      return lines.join('\n');
    };

    const rejectedAudits = recs
      .filter(r => r.isRejection || (r.qcPct || 1) < 1)
      .sort((a, b) => (a.qcPct || 0) - (b.qcPct || 0))
      .slice(0, 60)
      .map(r => `${r.visitDate || r.date || '-'} | ${r.store || '-'} (${r.outletCode || '-'}) | ${r.branch || '-'} | ${r.aom || '-'} | ${r.empName || '-'} | ${r.category || '-'} | ${((r.qcPct || 0) * 100).toFixed(1)}% | Mistakes: ${r.mistakes || '-'} | Remarks: ${r.trainerRemarks || '-'}`);

    return [
      `=== GODREJ QC PERFORMANCE LIVE SHAREPOINT DATASOURCE ===`,
      `Status: ${updatedMeta} | Active Filters: ${filterSummary}`,
      `Active Scope: ${recs.length} QC Audits | Unique Stores: ${stores.size} | Unique VMs: ${vms.size} | Avg QC Score: ${avgQC}% | Perfect 100% Audits: ${perfect} (${pct(perfect, recs.length)}) | Rejected Audits: ${rejAudits} | Total Parameter Failures: ${paramRejTotal}`,
      ``,
      `--- QC PARAMETER FAILURE BREAKDOWN ---`,
      paramLines.join('\n'),
      ``,
      `--- AOM-WISE QC PERFORMANCE ---`,
      fmtQCGrp(aomMap, 'AOM | Audits | AvgQCScore% | RejectedAudits | ParamFailures'),
      ``,
      `--- CATEGORY & CLASSIFICATION QC PERFORMANCE ---`,
      fmtQCGrp(catMap, 'Category (Class) | Audits | AvgQCScore% | RejectedAudits | ParamFailures'),
      ``,
      `--- BRANCH-WISE QC PERFORMANCE ---`,
      fmtQCGrp(branchMap, 'Branch | AOM | Audits | AvgQCScore% | RejectedAudits | ParamFailures', 40),
      ``,
      `--- VM-WISE QC PERFORMANCE ---`,
      fmtQCGrp(vmMap, 'VM Name | AOM | Branch | Audits | AvgQCScore% | RejectedAudits | ParamFailures', 80),
      ``,
      `--- REJECTED / UNDER-100% STORE AUDITS (TOP 60) ---`,
      `Date | Store | Branch | AOM | VM | Category | QCScore% | Mistakes | TrainerRemarks`,
      ...rejectedAudits
    ].join('\n');
  }

  /* --- 5. TRAINING PERFORMANCE (training.html) --- */
  function buildTrainingPipeTable(allRecs, activeRecs, filterSummary, updatedMeta) {
    const recs = activeRecs.length ? activeRecs : allRecs;
    let present = 0, attempted = 0, unattempted = 0, scoreSum = 0, ptsAch = 0, ptsTot = 0, passed = 0;
    const topics = new Set(), vms = new Set(), branches = new Set();
    const aomMap = {}, catMap = {}, branchMap = {}, empMap = {};

    recs.forEach(r => {
      if (r.topic) topics.add(r.topic);
      if (r.empCode || r.empName) vms.add(r.empCode || r.empName);
      if (r.branch) branches.add(r.branch);
      const isPres = String(r.attendance || '').toLowerCase() === 'present';
      if (isPres) present++;
      if (r.isAttempted) {
        attempted++;
        const sc = (Number(r.scorePct) || 0) * 100;
        scoreSum += sc;
        ptsAch += Number(r.achievedPoints) || 0;
        ptsTot += Number(r.totalPoints) || 0;
        if (sc >= 80) passed++;
      } else if (r.isUnattempted) {
        unattempted++;
      }

      const addT = (map, key, extra) => {
        const k = key || 'Unknown';
        if (!map[k]) map[k] = { total: 0, pres: 0, att: 0, unatt: 0, scSum: 0, pass: 0, extra: extra || '' };
        const g = map[k];
        g.total++;
        if (isPres) g.pres++;
        if (r.isAttempted) {
          g.att++;
          const sc = (Number(r.scorePct) || 0) * 100;
          g.scSum += sc;
          if (sc >= 80) g.pass++;
        } else if (r.isUnattempted) {
          g.unatt++;
        }
      };

      addT(aomMap, r.rh);
      addT(catMap, `${r.category || '-'} (${r.weekLabel || '-'})`);
      addT(branchMap, r.branch, r.rh);
      addT(empMap, r.empName, `${r.empCode || '-'} | ${r.rh || '-'} | ${r.branch || '-'}`);
    });

    const avgScore = attempted ? (scoreSum / attempted).toFixed(1) : '0.0';

    const fmtTGrp = (map, hdr, limit = 60) => {
      const lines = [hdr];
      Object.entries(map)
        .sort((a, b) => (b[1].att ? b[1].scSum / b[1].att : 0) - (a[1].att ? a[1].scSum / a[1].att : 0))
        .slice(0, limit)
        .forEach(([k, g]) => {
          const avg = g.att ? (g.scSum / g.att).toFixed(1) : '0.0';
          lines.push(`${k}${g.extra ? ' | ' + g.extra : ''} | ${g.total} | ${g.pres} (${pct(g.pres, g.total)}) | ${g.att} | ${g.unatt} | ${avg}% | ${pct(g.pass, g.att)}`);
        });
      return lines.join('\n');
    };

    const unattemptedList = recs
      .filter(r => r.isUnattempted || (r.isAttempted && (r.scorePct || 0) < 0.75))
      .slice(0, 60)
      .map(r => `${r.empName || '-'} (${r.empCode || '-'}) | AOM:${r.rh || '-'} | Branch:${r.branch || '-'} | Cat:${r.category || '-'} | Topic:${r.topic || '-'} | Week:${r.weekLabel || '-'} | Att:${r.attendance || '-'} | Attempted:${r.attempted || '-'} | Score:${r.scorePct != null ? (r.scorePct * 100).toFixed(1) + '%' : '-'}`);

    return [
      `=== GODREJ TRAINING PERFORMANCE LIVE SHAREPOINT DATASOURCE ===`,
      `Status: ${updatedMeta} | Active Filters: ${filterSummary}`,
      `Active Scope: ${recs.length} Training Records | Unique Topics: ${topics.size} | Unique Employees: ${vms.size} | Present: ${present}/${recs.length} (${pct(present, recs.length)}) | Assessment Attempted: ${attempted} | Unattempted: ${unattempted} | Avg Assessment Score: ${avgScore}% (Points: ${ptsAch}/${ptsTot}) | Pass Rate (>=80%): ${pct(passed, attempted)}`,
      ``,
      `--- AOM-WISE TRAINING SUMMARY ---`,
      fmtTGrp(aomMap, 'AOM | Enrolled | Present | Attempted | Unattempted | AvgScore% | PassRate%'),
      ``,
      `--- CATEGORY & WEEK TRAINING SUMMARY ---`,
      fmtTGrp(catMap, 'Category (Week) | Enrolled | Present | Attempted | Unattempted | AvgScore% | PassRate%'),
      ``,
      `--- BRANCH-WISE TRAINING SUMMARY ---`,
      fmtTGrp(branchMap, 'Branch | AOM | Enrolled | Present | Attempted | Unattempted | AvgScore% | PassRate%', 35),
      ``,
      `--- EMPLOYEE-WISE TRAINING SUMMARY ---`,
      fmtTGrp(empMap, 'Employee | EmpCode | AOM | Branch | Enrolled | Present | Attempted | Unattempted | AvgScore% | PassRate%', 85),
      ``,
      `--- UNATTEMPTED OR LOW ASSESSMENT SCORE (<75%) RECORDS ---`,
      ...unattemptedList
    ].join('\n');
  }

  /* --------------------------------------------------------------------------
     SMART CLIENT-SIDE FALLBACK ANALYTICAL ENGINE (Works on Static GitHub Pages)
     -------------------------------------------------------------------------- */
  function answerLocallyFromLiveDatasource(question, pageType) {
    const q = (question || '').toLowerCase();
    const allRecs = getAllRecords(pageType);
    const activeRecs = getActiveRecords(pageType);
    const recs = activeRecs.length ? activeRecs : allRecs;
    const filterSummary = getActiveFilterSummary();
    const updatedMeta = getDatasourceUpdatedText();

    if (!recs || !recs.length) {
      return `The SharePoint datasource is still loading on **${getPageLabel(pageType)}**. Please wait a moment and try again.`;
    }

    const sum = (arr, fn) => arr.reduce((a, r) => a + (Number(fn(r)) || 0), 0);

    // 1. VM PERFORMANCE (program)
    if (pageType === 'program') {
      const covTgt = sum(recs, r => r.coverageTarget);
      const covAch = sum(recs, r => r.coverageAch);
      const visTgt = sum(recs, r => r.visitsTarget);
      const visAch = sum(recs, r => r.visitAch);
      const prodVis = sum(recs, r => r.productiveVisits);
      const disp = sum(recs, r => r.productDisplay);
      const posm = sum(recs, r => r.posmDeployment);
      const manAch = sum(recs, r => r.manDaysAchieved);
      const manTgt = sum(recs, r => r.targetManDays);

      // Check if user asked about a specific VM, AOM, Branch, or State
      const matchedVMs = recs.filter(r =>
        (r.userName && q.includes(r.userName.toLowerCase())) ||
        (r.aom && q.includes(r.aom.toLowerCase())) ||
        (r.branch && q.includes(r.branch.toLowerCase())) ||
        (r.state && q.includes(r.state.toLowerCase()))
      );

      if (matchedVMs.length > 0 && matchedVMs.length <= 25 && !q.includes('rank aom') && !q.includes('overall') && !q.includes('top 5')) {
        const mCovTgt = sum(matchedVMs, r => r.coverageTarget);
        const mCovAch = sum(matchedVMs, r => r.coverageAch);
        const mVisTgt = sum(matchedVMs, r => r.visitsTarget);
        const mVisAch = sum(matchedVMs, r => r.visitAch);
        const rows = matchedVMs
          .sort((a, b) => (b.finalAch || 0) - (a.finalAch || 0))
          .slice(0, 15)
          .map(r => `| ${r.userName} | ${r.aom} | ${r.branch} | ${r.coverageAch}/${r.coverageTarget} (**${pct(r.coverageAch, r.coverageTarget)}**) | ${r.visitAch}/${r.visitsTarget} (**${pct(r.visitAch, r.visitsTarget)}**) | **${((r.finalAch || 0) * 100).toFixed(1)}%** |`);
        return [
          `**Live SharePoint Match (${matchedVMs.length} VMs)** · *${updatedMeta}*`,
          `- **Coverage Achievement:** **${mCovAch.toLocaleString()} / ${mCovTgt.toLocaleString()}** (**${pct(mCovAch, mCovTgt)}**, Pending: **${Math.max(0, mCovTgt - mCovAch)}**)`,
          `- **Visit Achievement:** **${mVisAch.toLocaleString()} / ${mVisTgt.toLocaleString()}** (**${pct(mVisAch, mVisTgt)}**)`,
          ``,
          `| VM Name | AOM | Branch | Coverage | Visits | Final Ach % |`,
          `|---|---|---|---|---|---|`,
          ...rows
        ].join('\n');
      }

      if (q.includes('top') || q.includes('bottom') || q.includes('best') || q.includes('lowest') || q.includes('vm')) {
        const sorted = [...recs].sort((a, b) => (b.finalAch || 0) - (a.finalAch || 0));
        const top5 = sorted.slice(0, 5).map((r, i) => `| #${i + 1} | ${r.userName} | ${r.aom} | ${r.branch} | ${r.coverageAch}/${r.coverageTarget} (${pct(r.coverageAch, r.coverageTarget)}) | ${r.visitAch}/${r.visitsTarget} (${pct(r.visitAch, r.visitsTarget)}) | **${((r.finalAch || 0) * 100).toFixed(1)}%** |`);
        const bot5 = sorted.slice(-5).reverse().map((r, i) => `| #${sorted.length - i} | ${r.userName} | ${r.aom} | ${r.branch} | ${r.coverageAch}/${r.coverageTarget} (${pct(r.coverageAch, r.coverageTarget)}) | ${r.visitAch}/${r.visitsTarget} (${pct(r.visitAch, r.visitsTarget)}) | **${((r.finalAch || 0) * 100).toFixed(1)}%** |`);
        return [
          `**Top 5 & Bottom 5 Field VMs** (*${filterSummary} · ${updatedMeta}*)`,
          ``,
          `**Top 5 Performers:**`,
          `| Rank | VM Name | AOM | Branch | Coverage | Visits | Final Ach % |`,
          `|---|---|---|---|---|---|---|`,
          ...top5,
          ``,
          `**Bottom 5 Performers (Need Attention):**`,
          `| Rank | VM Name | AOM | Branch | Coverage | Visits | Final Ach % |`,
          `|---|---|---|---|---|---|---|`,
          ...bot5
        ].join('\n');
      }

      if (q.includes('state') || q.includes('pending')) {
        const stMap = {};
        recs.forEach(r => {
          const k = r.state || 'Unknown';
          if (!stMap[k]) stMap[k] = { vms: 0, covTgt: 0, covAch: 0, visTgt: 0, visAch: 0 };
          stMap[k].vms++;
          stMap[k].covTgt += r.coverageTarget || 0;
          stMap[k].covAch += r.coverageAch || 0;
          stMap[k].visTgt += r.visitsTarget || 0;
          stMap[k].visAch += r.visitAch || 0;
        });
        const rows = Object.entries(stMap)
          .sort((a, b) => (b[1].covTgt - b[1].covAch) - (a[1].covTgt - a[1].covAch))
          .slice(0, 12)
          .map(([st, g]) => `| ${st} | ${g.vms} | ${g.covTgt} | ${g.covAch} | **${Math.max(0, g.covTgt - g.covAch)}** | **${pct(g.covAch, g.covTgt)}** | ${pct(g.visAch, g.visTgt)} |`);
        return [
          `**State-wise Coverage & Pending Outlets** (*${filterSummary} · ${updatedMeta}*)`,
          ``,
          `| State | VMs | Target | Achieved | Pending | Coverage % | Visit % |`,
          `|---|---|---|---|---|---|---|`,
          ...rows
        ].join('\n');
      }

      // Default / AOM & Executive Summary
      const aomMap = {};
      recs.forEach(r => {
        const k = r.aom || 'Unassigned';
        if (!aomMap[k]) aomMap[k] = { vms: 0, covTgt: 0, covAch: 0, visTgt: 0, visAch: 0, prod: 0, disp: 0, posm: 0, achSum: 0 };
        aomMap[k].vms++;
        aomMap[k].covTgt += r.coverageTarget || 0;
        aomMap[k].covAch += r.coverageAch || 0;
        aomMap[k].visTgt += r.visitsTarget || 0;
        aomMap[k].visAch += r.visitAch || 0;
        aomMap[k].prod += r.productiveVisits || 0;
        aomMap[k].disp += r.productDisplay || 0;
        aomMap[k].posm += r.posmDeployment || 0;
        aomMap[k].achSum += (r.finalAch || 0) * 100;
      });
      const aomRows = Object.entries(aomMap)
        .sort((a, b) => (b[1].covTgt ? b[1].covAch / b[1].covTgt : 0) - (a[1].covTgt ? a[1].covAch / a[1].covTgt : 0))
        .map(([aom, g], idx) => `| #${idx + 1} | **${aom}** | ${g.vms} | ${g.covAch}/${g.covTgt} (**${pct(g.covAch, g.covTgt)}**) | ${g.visAch}/${g.visTgt} (**${pct(g.visAch, g.visTgt)}**) | ${g.prod.toLocaleString()} | **${(g.achSum / g.vms).toFixed(1)}%** |`);

      return [
        `**VM Program Performance Summary** (*${filterSummary} · ${updatedMeta}*)`,
        `- **Overall Coverage:** **${covAch.toLocaleString()} / ${covTgt.toLocaleString()}** (**${pct(covAch, covTgt)}**, Pending: **${Math.max(0, covTgt - covAch).toLocaleString()}**) across **${recs.length} VMs**`,
        `- **Total Visits:** **${visAch.toLocaleString()} / ${visTgt.toLocaleString()}** (**${pct(visAch, visTgt)}**) · **Productive Visits:** **${prodVis.toLocaleString()}** (**${pct(prodVis, visAch)}**)`,
        `- **Field Execution:** **${disp.toLocaleString()}** Product Displays · **${posm.toLocaleString()}** POSM Deployed · **${manAch}/${manTgt}** Man-Days (**${pct(manAch, manTgt)}**)`,
        ``,
        `| Rank | AOM | VMs | Coverage (Ach/Tgt) | Visits (Ach/Tgt) | Prod Visits | Avg Final Ach % |`,
        `|---|---|---|---|---|---|---|`,
        ...aomRows
      ].join('\n');
    }

    // 2. BRANCH PERFORMANCE (branch)
    if (pageType === 'branch') {
      const covTgt = sum(recs, r => r.coverageTarget);
      const covAch = sum(recs, r => r.coverageAch);
      const visTgt = sum(recs, r => r.visitsTarget);
      const visAch = sum(recs, r => r.visitAch);
      const prodVis = sum(recs, r => r.productiveVisits);

      if (q.includes('branch') || q.includes('top') || q.includes('bottom') || q.includes('lowest') || q.includes('highest')) {
        const sorted = [...recs].sort((a, b) => (b.finalAch || 0) - (a.finalAch || 0));
        const topRows = sorted.slice(0, 8).map((r, i) => `| #${i + 1} | **${r.branch}** | ${r.zone} | ${r.aom} | ${r.coverageAch}/${r.coverageTarget} (${pct(r.coverageAch, r.coverageTarget)}) | ${r.visitAch}/${r.visitsTarget} (${pct(r.visitAch, r.visitsTarget)}) | **${((r.finalAch || 0) * 100).toFixed(1)}%** |`);
        const botRows = sorted.slice(-5).reverse().map((r, i) => `| #${sorted.length - i} | **${r.branch}** | ${r.zone} | ${r.aom} | ${r.coverageAch}/${r.coverageTarget} (${pct(r.coverageAch, r.coverageTarget)}) | ${r.visitAch}/${r.visitsTarget} (${pct(r.visitAch, r.visitsTarget)}) | **${((r.finalAch || 0) * 100).toFixed(1)}%** |`);
        return [
          `**Branch Performance Ranking** (*${filterSummary} · ${updatedMeta}*)`,
          ``,
          `**Top Performing Branches:**`,
          `| Rank | Branch | Zone | AOM | Coverage | Visits | Final Ach % |`,
          `|---|---|---|---|---|---|---|`,
          ...topRows,
          ``,
          `**Lowest Performing Branches:**`,
          `| Rank | Branch | Zone | AOM | Coverage | Visits | Final Ach % |`,
          `|---|---|---|---|---|---|---|`,
          ...botRows
        ].join('\n');
      }

      const zMap = {};
      recs.forEach(r => {
        const k = r.zone || 'Unknown';
        if (!zMap[k]) zMap[k] = { b: 0, ct: 0, ca: 0, vt: 0, va: 0, pv: 0 };
        zMap[k].b++;
        zMap[k].ct += r.coverageTarget || 0;
        zMap[k].ca += r.coverageAch || 0;
        zMap[k].vt += r.visitsTarget || 0;
        zMap[k].va += r.visitAch || 0;
        zMap[k].pv += r.productiveVisits || 0;
      });
      const zRows = Object.entries(zMap)
        .sort((a, b) => (b[1].ct ? b[1].ca / b[1].ct : 0) - (a[1].ct ? a[1].ca / a[1].ct : 0))
        .map(([z, g]) => `| **${z}** | ${g.b} | ${g.ca}/${g.ct} (**${pct(g.ca, g.ct)}**) | ${g.va}/${g.vt} (**${pct(g.va, g.vt)}**) | ${g.pv} (${pct(g.pv, g.va)}) |`);

      return [
        `**Branch Performance Summary** (*${filterSummary} · ${updatedMeta}*)`,
        `- **Total Branches:** **${recs.length}** · **Coverage:** **${covAch.toLocaleString()} / ${covTgt.toLocaleString()}** (**${pct(covAch, covTgt)}**)`,
        `- **Total Visits:** **${visAch.toLocaleString()} / ${visTgt.toLocaleString()}** (**${pct(visAch, visTgt)}**) · **Productive Visits:** **${prodVis.toLocaleString()}** (**${pct(prodVis, visAch)}**)`,
        ``,
        `| Zone | Branches | Coverage (Ach/Tgt) | Visits (Ach/Tgt) | Productive Visits |`,
        `|---|---|---|---|---|`,
        ...zRows
      ].join('\n');
    }

    // 3. MERCHANDISING SCORE (m_score)
    if (pageType === 'm_score') {
      const avg = ((sum(recs, r => r.mScore) / recs.length) * 100).toFixed(1);
      const stores = new Set(recs.map(r => r.storeCode || r.storeName)).size;
      const aomMap = {};
      recs.forEach(r => {
        const k = r.aomName || 'Unknown';
        if (!aomMap[k]) aomMap[k] = { c: 0, s: 0, st: new Set() };
        aomMap[k].c++;
        aomMap[k].s += (Number(r.mScore) || 0) * 100;
        aomMap[k].st.add(r.storeCode || r.storeName);
      });
      const aomRows = Object.entries(aomMap)
        .sort((a, b) => (b[1].s / b[1].c) - (a[1].s / a[1].c))
        .map(([k, g]) => `| **${k}** | ${g.c} | ${g.st.size} | **${(g.s / g.c).toFixed(1)}%** |`);

      if (q.includes('low') || q.includes('store') || q.includes('bottom')) {
        const lowRows = [...recs]
          .sort((a, b) => (a.mScore || 0) - (b.mScore || 0))
          .slice(0, 10)
          .map(r => `| ${r.visitDate || '-'} | ${r.storeName} | ${r.branch} | ${r.vmName} | ${r.prodCat} | **${((r.mScore || 0) * 100).toFixed(1)}%** |`);
        return [
          `**Lowest Merchandising Score Store Audits** (*${filterSummary} · ${updatedMeta}*)`,
          ``,
          `| Date | Store Name | Branch | VM | Product Category | M-Score % |`,
          `|---|---|---|---|---|---|`,
          ...lowRows
        ].join('\n');
      }

      return [
        `**Merchandising Score (M-Score) Summary** (*${filterSummary} · ${updatedMeta}*)`,
        `- **Overall Average M-Score:** **${avg}%** across **${recs.length.toLocaleString()}** product-store audits in **${stores}** unique stores`,
        ``,
        `| AOM | Product Audits | Unique Stores | Avg M-Score % |`,
        `|---|---|---|---|`,
        ...aomRows
      ].join('\n');
    }

    // 4. TRAINING PERFORMANCE (training)
    if (pageType === 'training') {
      const pres = recs.filter(r => String(r.attendance || '').toLowerCase() === 'present').length;
      const attRecs = recs.filter(r => r.isAttempted);
      const unattRecs = recs.filter(r => r.isUnattempted);
      const scoreAvg = attRecs.length ? ((sum(attRecs, r => r.scorePct) / attRecs.length) * 100).toFixed(1) : '0.0';

      if (q.includes('unattempt') || q.includes('not attempt') || q.includes('pending')) {
        const rows = unattRecs.slice(0, 15).map(r => `| ${r.empName} | ${r.rh} | ${r.branch} | ${r.category} | ${r.weekLabel} | ${r.attendance} |`);
        return [
          `**Employees With Unattempted Assessment (${unattRecs.length} records)** (*${filterSummary} · ${updatedMeta}*)`,
          ``,
          `| Employee Name | AOM | Branch | Category | Week | Attendance |`,
          `|---|---|---|---|---|---|`,
          ...rows
        ].join('\n');
      }

      const aomMap = {};
      recs.forEach(r => {
        const k = r.rh || 'Unknown';
        if (!aomMap[k]) aomMap[k] = { tot: 0, pr: 0, at: 0, unat: 0, sc: 0 };
        aomMap[k].tot++;
        if (String(r.attendance || '').toLowerCase() === 'present') aomMap[k].pr++;
        if (r.isAttempted) {
          aomMap[k].at++;
          aomMap[k].sc += (Number(r.scorePct) || 0) * 100;
        } else if (r.isUnattempted) {
          aomMap[k].unat++;
        }
      });
      const aomRows = Object.entries(aomMap)
        .sort((a, b) => (b[1].at ? b[1].sc / b[1].at : 0) - (a[1].at ? a[1].sc / a[1].at : 0))
        .map(([k, g]) => {
          const avg = g.at ? (g.sc / g.at).toFixed(1) : '0.0';
          return `| **${k}** | ${g.tot} | ${g.pr} (${pct(g.pr, g.tot)}) | ${g.at} | ${g.unat} | **${avg}%** |`;
        });

      return [
        `**Training Performance Summary** (*${filterSummary} · ${updatedMeta}*)`,
        `- **Attendance:** **${pres} / ${recs.length}** (**${pct(pres, recs.length)}**) · **Assessment Attempted:** **${attRecs.length}** (Unattempted: **${unattRecs.length}**)`,
        `- **Average Assessment Score:** **${scoreAvg}%**`,
        ``,
        `| AOM | Enrolled | Present | Attempted | Unattempted | Avg Score % |`,
        `|---|---|---|---|---|---|`,
        ...aomRows
      ].join('\n');
    }

    // 5. QC PERFORMANCE (qc)
    const avgQC = ((sum(recs, r => r.qcPct) / recs.length) * 100).toFixed(1);
    let totRej = 0;
    const paramRej = {};
    recs.forEach(r => {
      if (Array.isArray(r.display)) {
        r.display.forEach(p => { if (p && !p.ok) { paramRej[p.name] = (paramRej[p.name] || 0) + 1; totRej++; } });
      }
      if (Array.isArray(r.campaign)) {
        r.campaign.forEach(p => { if (p && !p.ok) { paramRej[p.name] = (paramRej[p.name] || 0) + 1; totRej++; } });
      }
    });
    const pRows = Object.entries(paramRej)
      .sort((a, b) => b[1] - a[1])
      .map(([p, c]) => `| **${p}** | ${c} | ${pct(c, totRej)} |`);

    const aomMap = {};
    recs.forEach(r => {
      const k = r.aom || 'Unknown';
      if (!aomMap[k]) aomMap[k] = { c: 0, s: 0, rj: 0 };
      aomMap[k].c++;
      aomMap[k].s += (Number(r.qcPct) || 0) * 100;
      if (r.isRejection) aomMap[k].rj++;
    });
    const aomRows = Object.entries(aomMap)
      .sort((a, b) => (b[1].s / b[1].c) - (a[1].s / a[1].c))
      .map(([k, g]) => `| **${k}** | ${g.c} | **${(g.s / g.c).toFixed(1)}%** | ${g.rj} |`);

    return [
      `**QC Performance Summary** (*${filterSummary} · ${updatedMeta}*)`,
      `- **Total Audits:** **${recs.length}** · **Average QC Score:** **${avgQC}%** · **Total Parameter Failures:** **${totRej}**`,
      ``,
      `**AOM-wise QC Breakdown:**`,
      `| AOM | Audits | Avg QC Score % | Rejected Audits |`,
      `|---|---|---|---|`,
      ...aomRows,
      ``,
      `**Parameter-wise Failures:**`,
      `| QC Parameter | Failures | Share % |`,
      `|---|---|---|`,
      ...pRows
    ].join('\n');
  }

  /* --------------------------------------------------------------------------
     MARKDOWN TO SAFE HTML RENDERER
     -------------------------------------------------------------------------- */
  function escapeHtml(str) {
    return String(str || '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  function formatInline(text) {
    return escapeHtml(text)
      .replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>')
      .replace(/\*(.+?)\*/g, '<em>$1</em>')
      .replace(/`([^`]+)`/g, '<code>$1</code>');
  }

  function renderMarkdown(md) {
    if (!md) return '';
    const lines = String(md).split(/\r?\n/);
    let html = '';
    let inList = false;
    let inTable = false;

    const closeList = () => {
      if (inList) { html += '</ul>'; inList = false; }
    };
    const closeTable = () => {
      if (inTable) { html += '</tbody></table></div>'; inTable = false; }
    };

    for (let i = 0; i < lines.length; i++) {
      const raw = lines[i];
      const trimmed = raw.trim();

      if (!trimmed) {
        closeList();
        closeTable();
        continue;
      }

      // Table row
      if (trimmed.startsWith('|') && trimmed.endsWith('|')) {
        closeList();
        const cells = trimmed.slice(1, -1).split('|').map(c => c.trim());
        // Skip separator row |---|---|
        if (cells.every(c => /^[:\-]+$/.test(c))) {
          continue;
        }
        if (!inTable) {
          html += '<div class="ai-table-wrap"><table class="ai-chat-table"><thead><tr>' + cells.map(c => `<th>${formatInline(c)}</th>`).join('') + '</tr></thead><tbody>';
          inTable = true;
        } else {
          html += '<tr>' + cells.map(c => `<td>${formatInline(c)}</td>`).join('') + '</tr>';
        }
        continue;
      } else {
        closeTable();
      }

      // Headings
      if (/^#{1,4}\s+/.test(trimmed)) {
        closeList();
        const content = trimmed.replace(/^#{1,4}\s+/, '');
        html += `<p><strong>${formatInline(content)}</strong></p>`;
        continue;
      }

      // Bullet list
      if (/^[\-\*\•]\s+/.test(trimmed) || /^\d+\.\s+/.test(trimmed)) {
        if (!inList) { html += '<ul>'; inList = true; }
        const content = trimmed.replace(/^([\-\*\•]|\d+\.)\s+/, '');
        html += `<li>${formatInline(content)}</li>`;
        continue;
      } else {
        closeList();
      }

      html += `<p>${formatInline(trimmed)}</p>`;
    }

    closeList();
    closeTable();
    return html;
  }

  /* --------------------------------------------------------------------------
     UI BUILDER & SHAREPOINT LIVE SYNC HOOK
     -------------------------------------------------------------------------- */
  function createWidgetDOM() {
    const pageType = detectPageType();
    const pageLabel = getPageLabel(pageType);

    // Right-Edge Vertical Tab Launcher Button (Opposite to Left-Edge Filters Tab)
    const launcher = document.createElement('button');
    launcher.className = 'ai-copilot-launcher';
    launcher.id = 'aiCopilotLauncher';
    launcher.setAttribute('aria-label', 'Open Godrej AI Copilot');
    launcher.title = 'Toggle Godrej AI Copilot';
    launcher.innerHTML = `
      <svg class="ai-copilot-chevron" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="15 18 9 12 15 6"/></svg>
      <span class="ai-copilot-launcher-icon">
        <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.3" stroke-linecap="round" stroke-linejoin="round">
          <path d="M12 2v4M12 18v4M4.93 4.93l2.83 2.83M16.24 16.24l2.83 2.83M2 12h4M18 12h4M4.93 19.07l2.83-2.83M16.24 7.76l2.83-2.83"/>
        </svg>
      </span>
      <span class="ai-copilot-launcher-text">AI Copilot</span>
      <span class="ai-copilot-status-dot" title="Connected to live SharePoint datasource"></span>
    `;

    // Chat Drawer Panel
    const panel = document.createElement('section');
    panel.className = 'ai-copilot-panel';
    panel.id = 'aiCopilotPanel';
    panel.setAttribute('aria-label', 'Godrej AI Copilot Drawer');
    panel.innerHTML = `
      <div class="ai-copilot-header">
        <div class="ai-copilot-brand">
          <div class="ai-copilot-avatar">
            <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#ffffff" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round">
              <path d="M12 2L14.8 8.5L22 9.3L16.5 14L18.2 21L12 17.3L5.8 21L7.5 14L2 9.3L9.2 8.5L12 2Z"/>
            </svg>
          </div>
          <div class="ai-copilot-title-wrap">
            <div class="ai-copilot-title">
              <span>Godrej AI Copilot</span>
              <span class="ai-copilot-badge" id="aiCopilotLiveBadge">SharePoint Live</span>
            </div>
            <div class="ai-copilot-subtitle" id="aiCopilotSubtitle">${pageLabel} · Connecting to SharePoint...</div>
          </div>
        </div>
        <div class="ai-copilot-actions">
          <button class="ai-copilot-icon-btn" id="aiCopilotSyncBtn" title="Refresh SharePoint Datasource Now">
            <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2">
              <path d="M21 12a9 9 0 0 0-9-9 9.75 9.75 0 0 0-6.74 2.74L3 8"/><path d="M3 3v5h5"/><path d="M3 12a9 9 0 0 0 9 9 9.75 9.75 0 0 0 6.74-2.74L21 16"/><path d="M16 16h5v5"/>
            </svg>
          </button>
          <button class="ai-copilot-icon-btn" id="aiCopilotExpandBtn" title="Expand / Collapse Drawer">
            <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2">
              <polyline points="15 3 21 3 21 9"/><polyline points="9 21 3 21 3 15"/>
              <line x1="21" y1="3" x2="14" y2="10"/><line x1="3" y1="21" x2="10" y2="14"/>
            </svg>
          </button>
          <button class="ai-copilot-icon-btn" id="aiCopilotClearBtn" title="Reset Conversation">
            <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2">
              <path d="M3 6h18"/><path d="M8 6V4h8v2"/><path d="M19 6l-1 14H6L5 6"/>
            </svg>
          </button>
          <button class="ai-copilot-icon-btn" id="aiCopilotCloseBtn" title="Close">
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4">
              <line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/>
            </svg>
          </button>
        </div>
      </div>

      <div class="ai-copilot-context-bar">
        <span>Active Scope:</span>
        <span class="ai-copilot-context-pill" id="aiCopilotContextPill">Loading filters...</span>
      </div>

      <div class="ai-copilot-messages" id="aiCopilotMessages"></div>

      <div class="ai-copilot-suggestions" id="aiCopilotSuggestions"></div>

      <form class="ai-copilot-composer" id="aiCopilotForm">
        <input
          type="text"
          class="ai-copilot-input"
          id="aiCopilotInput"
          placeholder="Ask anything about the live SharePoint data..."
          autocomplete="off"
        />
        <button type="submit" class="ai-copilot-send" id="aiCopilotSendBtn" title="Send">
          <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round">
            <line x1="22" y1="2" x2="11" y2="13"/><polygon points="22 2 15 22 11 13 2 9 22 2"/>
          </svg>
        </button>
      </form>
    `;

    document.body.appendChild(launcher);
    document.body.appendChild(panel);

    renderSuggestions(pageType);
    addWelcomeMessage(pageType);
    updateContextUI();

    launcher.addEventListener('click', () => toggleDrawer());
    document.getElementById('aiCopilotCloseBtn').addEventListener('click', () => toggleDrawer(false));
    document.getElementById('aiCopilotExpandBtn').addEventListener('click', () => {
      state.isExpanded = !state.isExpanded;
      panel.classList.toggle('expanded', state.isExpanded);
    });
    document.getElementById('aiCopilotClearBtn').addEventListener('click', () => {
      state.messages = [];
      document.getElementById('aiCopilotMessages').innerHTML = '';
      addWelcomeMessage(pageType);
    });
    document.getElementById('aiCopilotSyncBtn').addEventListener('click', () => {
      triggerSharePointRefresh();
    });

    document.getElementById('aiCopilotForm').addEventListener('submit', (e) => {
      e.preventDefault();
      const input = document.getElementById('aiCopilotInput');
      const text = (input.value || '').trim();
      if (!text || state.isLoading) return;
      input.value = '';
      sendUserMessage(text);
    });

    // Keep context bar & subtitle synchronized when filters change or SharePoint finishes syncing
    document.addEventListener('change', () => setTimeout(updateContextUI, 60));
    document.addEventListener('click', (e) => {
      if (e.target && e.target.closest('#btnFilterReset, #btnResetFilters, .compact-filter-reset, #liveStatusBadge')) {
        setTimeout(updateContextUI, 120);
      }
    });

    hookSharePointDatasourceOnLoad();
  }

  function triggerSharePointRefresh() {
    const subtitle = document.getElementById('aiCopilotSubtitle');
    if (subtitle) subtitle.textContent = 'Syncing latest SharePoint data...';
    try {
      if (typeof App !== 'undefined' && typeof App.syncFromSharePoint === 'function') {
        Promise.resolve(App.syncFromSharePoint(true)).finally(() => setTimeout(updateContextUI, 200));
        return;
      }
      if (typeof window.syncFromSharePoint === 'function') {
        Promise.resolve(window.syncFromSharePoint(true)).finally(() => setTimeout(updateContextUI, 200));
        return;
      }
    } catch (e) {}
    setTimeout(updateContextUI, 300);
  }

  function hookSharePointDatasourceOnLoad() {
    // Wrap App.syncFromSharePoint or window.syncFromSharePoint so AI Copilot updates immediately when live SharePoint data arrives
    try {
      if (typeof App !== 'undefined' && typeof App.syncFromSharePoint === 'function' && !App.__aiHooked) {
        const orig = App.syncFromSharePoint.bind(App);
        App.syncFromSharePoint = async function (...args) {
          const res = await orig(...args);
          state.lastSyncTimestamp = new Date();
          updateContextUI();
          return res;
        };
        App.__aiHooked = true;
      } else if (typeof window.syncFromSharePoint === 'function' && !window.__aiSyncHooked) {
        const orig = window.syncFromSharePoint;
        window.syncFromSharePoint = async function (...args) {
          const res = await orig.apply(this, args);
          state.lastSyncTimestamp = new Date();
          updateContextUI();
          return res;
        };
        window.__aiSyncHooked = true;
      }
    } catch (e) {}

    // Observe header sync metadata changes so any background SharePoint sync updates the AI Copilot banner
    const syncTargets = ['lastUpdatedMeta', 'dataUpdatedTillMeta', 'liveStatusText', 'dataSourceLabel'];
    syncTargets.forEach(id => {
      const el = document.getElementById(id);
      if (el && typeof MutationObserver !== 'undefined') {
        const obs = new MutationObserver(() => updateContextUI());
        obs.observe(el, { childList: true, characterData: true, subtree: true });
      }
    });

    setTimeout(updateContextUI, 500);
    setTimeout(updateContextUI, 2000);
  }

  function renderSuggestions(pageType) {
    const wrap = document.getElementById('aiCopilotSuggestions');
    if (!wrap) return;
    const items = getSuggestions(pageType);
    wrap.innerHTML = items
      .map(s => `<button type="button" class="ai-suggestion-chip">${escapeHtml(s)}</button>`)
      .join('');
    wrap.querySelectorAll('.ai-suggestion-chip').forEach((btn, idx) => {
      btn.addEventListener('click', () => {
        if (!state.isLoading) sendUserMessage(items[idx]);
      });
    });
  }

  function addWelcomeMessage(pageType) {
    const pageLabel = getPageLabel(pageType);
    const allRecs = getAllRecords(pageType);
    const countStr = allRecs.length ? `**${allRecs.length.toLocaleString()} live records**` : 'the live SharePoint workbook';
    const welcome = `Hello! I'm your **Godrej AI Copilot** for **${pageLabel}**, connected directly to ${countStr} from SharePoint.\n\nAsk me about **coverage, visits, AOM/branch rankings, M-Score, QC rejections, training scores, or specific stores & VMs** — I always answer from the latest synced datasource and your active filters.`;
    appendBubble('assistant', welcome);
  }

  const MONTHS_SHORT = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

  function formatDateDDMMMYYYY(input) {
    if (!input) return '';
    if (input instanceof Date && !isNaN(input.getTime())) {
      const dd = String(input.getDate()).padStart(2, '0');
      return `${dd} ${MONTHS_SHORT[input.getMonth()]} ${input.getFullYear()}`;
    }
    const str = String(input).trim();
    // Already formatted like "30 Oct 2026"
    if (/^\d{1,2}\s+[A-Za-z]{3}\s+\d{4}$/.test(str)) return str;
    // YYYY-MM-DD
    const mIso = str.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (mIso) {
      const yr = Number(mIso[1]);
      const mo = Number(mIso[2]) - 1;
      const dy = mIso[3];
      if (MONTHS_SHORT[mo]) return `${dy} ${MONTHS_SHORT[mo]} ${yr}`;
    }
    const parsed = new Date(str);
    if (!isNaN(parsed.getTime())) {
      const dd = String(parsed.getDate()).padStart(2, '0');
      return `${dd} ${MONTHS_SHORT[parsed.getMonth()]} ${parsed.getFullYear()}`;
    }
    return str;
  }

  function formatDateTimeSync(dt) {
    const d = dt instanceof Date && !isNaN(dt.getTime()) ? dt : new Date();
    const dd = String(d.getDate()).padStart(2, '0');
    const mon = MONTHS_SHORT[d.getMonth()];
    const yyyy = d.getFullYear();
    const timeStr = d.toLocaleTimeString('en-US', { hour: '2-digit', minute: '2-digit', hour12: true });
    return `${dd} ${mon} ${yyyy}, ${timeStr}`;
  }

  let __isNormalizingHeaderMeta = false;
  function normalizeHeaderSyncMeta(pageType) {
    if (__isNormalizingHeaderMeta) return;
    __isNormalizingHeaderMeta = true;
    try {
      // 1. Ensure status pill reads "Live SharePoint connected" when connected
      const statusTextEl = document.getElementById('liveStatusText');
      if (statusTextEl && statusTextEl.textContent) {
        const cur = statusTextEl.textContent.trim();
        if (cur.toLowerCase() === 'live data connected') {
          statusTextEl.textContent = 'Live SharePoint connected';
        }
      }

      // 2. Format #dataUpdatedTillMeta as "DD MMM YYYY" (e.g. "30 Oct 2026")
      const uptoEl = document.getElementById('dataUpdatedTillMeta');
      if (uptoEl) {
        let rawUpto = '';
        if (typeof State !== 'undefined' && State.metadata && State.metadata.dataUpdatedTill) {
          rawUpto = State.metadata.dataUpdatedTill;
        } else if (typeof State !== 'undefined' && State.dateBounds && State.dateBounds.max) {
          rawUpto = State.dateBounds.max;
        } else {
          const allRecs = getAllRecords(pageType);
          if (allRecs && allRecs.length) {
            let maxIso = '';
            allRecs.forEach(r => {
              const dVal = r.visitDate || r.qcDate || r.dateStr || '';
              if (typeof dVal === 'string' && /^\d{4}-\d{2}-\d{2}/.test(dVal)) {
                if (dVal > maxIso) maxIso = dVal;
              } else if (r.date instanceof Date && !isNaN(r.date.getTime())) {
                const iso = r.date.toISOString().slice(0, 10);
                if (iso > maxIso) maxIso = iso;
              }
            });
            if (maxIso) rawUpto = maxIso;
          }
        }
        const candidate = rawUpto || uptoEl.textContent.trim();
        const formattedUpto = formatDateDDMMMYYYY(candidate);
        if (formattedUpto && uptoEl.textContent !== formattedUpto) {
          uptoEl.textContent = formattedUpto;
        }
      }

      // 3. Format #lastUpdatedMeta as "DD MMM YYYY, hh:mm AM/PM" (e.g. "01 Nov 2026, 02:04 PM")
      const syncEl = document.getElementById('lastUpdatedMeta');
      if (syncEl) {
        const curSync = (syncEl.textContent || '').trim();
        if (!/^\d{2}\s+[A-Za-z]{3}\s+\d{4},\s+\d{1,2}:\d{2}\s*(AM|PM)$/i.test(curSync)) {
          const formattedSync = formatDateTimeSync(state.lastSyncTimestamp || new Date());
          if (syncEl.textContent !== formattedSync) {
            syncEl.textContent = formattedSync;
          }
        }
      }
    } catch (e) {} finally {
      __isNormalizingHeaderMeta = false;
    }
  }

  function updateContextUI() {
    const pageType = detectPageType();
    normalizeHeaderSyncMeta(pageType);
    const pageLabel = getPageLabel(pageType);
    const pill = document.getElementById('aiCopilotContextPill');
    const sub = document.getElementById('aiCopilotSubtitle');
    const allRecs = getAllRecords(pageType);
    const activeRecs = getActiveRecords(pageType);

    if (pill) {
      const fText = getActiveFilterSummary();
      pill.textContent = `${fText} (${activeRecs.length || allRecs.length} rows)`;
      pill.title = `${fText} (${activeRecs.length || allRecs.length} rows)`;
    }
    if (sub) {
      sub.textContent = `${pageLabel} · ${getDatasourceUpdatedText()}`;
    }
  }

  function toggleDrawer(forceState) {
    const panel = document.getElementById('aiCopilotPanel');
    const launcher = document.getElementById('aiCopilotLauncher');
    if (!panel) return;
    state.isOpen = typeof forceState === 'boolean' ? forceState : !state.isOpen;
    panel.classList.toggle('open', state.isOpen);
    if (launcher) launcher.classList.toggle('open', state.isOpen);
    if (state.isOpen) {
      updateContextUI();
      setTimeout(() => {
        const inp = document.getElementById('aiCopilotInput');
        if (inp) inp.focus();
      }, 120);
    }
  }

  function appendBubble(role, markdownText) {
    const container = document.getElementById('aiCopilotMessages');
    if (!container) return null;
    const wrap = document.createElement('div');
    wrap.className = `ai-msg ${role}`;
    const bubble = document.createElement('div');
    bubble.className = 'ai-msg-bubble';
    if (role === 'user') {
      bubble.textContent = markdownText;
    } else {
      bubble.innerHTML = renderMarkdown(markdownText);
    }
    wrap.appendChild(bubble);
    container.appendChild(wrap);
    container.scrollTop = container.scrollHeight;
    return bubble;
  }

  function showTypingIndicator() {
    const container = document.getElementById('aiCopilotMessages');
    const wrap = document.createElement('div');
    wrap.className = 'ai-msg assistant';
    wrap.id = 'aiCopilotTyping';
    wrap.innerHTML = `
      <div class="ai-msg-bubble">
        <div class="ai-typing-dots"><span></span><span></span><span></span></div>
      </div>
    `;
    container.appendChild(wrap);
    container.scrollTop = container.scrollHeight;
  }

  function removeTypingIndicator() {
    const el = document.getElementById('aiCopilotTyping');
    if (el) el.remove();
  }

  async function sendUserMessage(text) {
    const pageType = detectPageType();
    updateContextUI();
    appendBubble('user', text);
    state.messages.push({ role: 'user', content: text });

    state.isLoading = true;
    const sendBtn = document.getElementById('aiCopilotSendBtn');
    if (sendBtn) sendBtn.disabled = true;
    showTypingIndicator();

    const clientPipeTable = buildLiveClientPipeTable(pageType);

    try {
      const response = await fetch('/api/chat', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          page: pageType,
          messages: state.messages.slice(-8),
          activeFilters: { summary: getActiveFilterSummary(), updated: getDatasourceUpdatedText() },
          clientPipeTable
        })
      });

      let data = null;
      try {
        data = await response.json();
      } catch (jsonErr) {
        data = null;
      }

      removeTypingIndicator();

      if (response.ok && data && data.reply) {
        appendBubble('assistant', data.reply);
        state.messages.push({ role: 'assistant', content: data.reply });
      } else {
        // Automatic fallback to live in-memory SharePoint analytical engine (e.g., on static GitHub Pages)
        const localReply = answerLocallyFromLiveDatasource(text, pageType);
        appendBubble('assistant', localReply);
        state.messages.push({ role: 'assistant', content: localReply });
      }
    } catch (err) {
      removeTypingIndicator();
      const localReply = answerLocallyFromLiveDatasource(text, pageType);
      appendBubble('assistant', localReply);
      state.messages.push({ role: 'assistant', content: localReply });
    } finally {
      state.isLoading = false;
      if (sendBtn) sendBtn.disabled = false;
    }
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', createWidgetDOM);
  } else {
    createWidgetDOM();
  }
})();

