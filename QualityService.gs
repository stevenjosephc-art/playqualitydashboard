// ============================================================
// QualityService.gs — Aggregations and logic for Quality Audits
// Source: 'PLX Raw data' tab
// ============================================================

var QUALITY_SHEET_NAME = 'QualityAudits';

// ── SCHEMA MAPPING ────────────────────────────────────────────────────────

var Q_COLS = {
  CASE_ID: 0,            // A
  ENTITY_GROUP: 1,       // B
  AGENT_LDAP: 2,         // C
  OPENING_CHANNEL: 3,    // D
  REVIEW_DATE: 4,        // E
  REVIEW_WEEK: 5,        // F
  REVIEW_MONTH: 6,       // G
  CASE_DATE: 7,          // H
  CASE_WEEK: 8,          // I
  CASE_MONTH: 9,         // J
  CUSTOMER_CRITICAL: 10, // K
  BUSINESS_CRITICAL: 11, // L
  COMPLIANCE_CRITICAL: 12,// M
  REVIEWER_COMMENTS: 13, // N

  // Customer Critical Parameters (O-W)
  LISTENING: 14,
  PROBING: 15,
  COMPLETE_RESOLUTION: 16,
  TROUBLESHOOTING: 17,
  USER_EXPECTATIONS: 18,
  EMPATHY: 19,
  OWNERSHIP: 20,
  REFUNDS: 21,
  RESPONSIVENESS: 22,

  // Business Critical Parameters (X-AD)
  CONSULTS_ESCALATIONS: 23,
  CASE_DETAILS: 24,
  CATEGORIZATION: 25,
  CSAT_REMINDER: 26,
  CASE_STATE: 27,
  OPENING_CLOSING: 28,
  LANGUAGE_PROFICIENCY: 29,

  // Compliance Critical Parameters (AE-AH)
  AUTHENTICATION: 30,
  GOOGLE_ONLY_INFO: 31,
  PROFESSIONAL_CONDUCT: 32,
  PAYMENT_COMPLAINTS: 33,

  TEAM: 36,              // AK
  AGENT_WORKFLOW: 37,    // AL

  SUPERVISOR: 66,        // BO
  MANAGER: 67,           // BP
  LOB: 69                // BR
};

var Q_TARGETS = {
  CUSTOMER: 95,
  BUSINESS: 90,
  COMPLIANCE: 99.50
};

var Q_PARAM_GROUPS = {
  customer: ['LISTENING', 'PROBING', 'COMPLETE_RESOLUTION', 'TROUBLESHOOTING', 'USER_EXPECTATIONS', 'EMPATHY', 'OWNERSHIP', 'REFUNDS', 'RESPONSIVENESS'],
  business: ['CONSULTS_ESCALATIONS', 'CASE_DETAILS', 'CATEGORIZATION', 'CSAT_REMINDER', 'CASE_STATE', 'OPENING_CLOSING', 'LANGUAGE_PROFICIENCY'],
  compliance: ['AUTHENTICATION', 'GOOGLE_ONLY_INFO', 'PROFESSIONAL_CONDUCT', 'PAYMENT_COMPLAINTS']
};

var Q_PARAM_COLS = [].concat(Q_PARAM_GROUPS.customer, Q_PARAM_GROUPS.business, Q_PARAM_GROUPS.compliance);

var _MEMOIZED_RAW_DATA = null;

// ── APPS SCRIPT WEB APP ───────────────────────────────────────────────────

function doGet() {
  var template = HtmlService.createTemplateFromFile('QualityView');

  try {
    // Only fetch session for instant shell loading
    var session = clientGetSession();

    template.bootstrap = JSON.stringify({
      session: session
    });
  } catch(e) {
    Logger.log('doGet Error: ' + e.message);
    template.bootstrap = JSON.stringify({ error: e.message });
  }

  return template.evaluate()
    .setTitle('Quality Dashboard')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ── DATA LOADING ──────────────────────────────────────────────────────────

function getRawQualityData(forceRefresh) {
  if (_MEMOIZED_RAW_DATA && !forceRefresh) return _MEMOIZED_RAW_DATA;

  var cache = CacheService.getScriptCache();
  var cacheKey = 'quality_raw_v1';

  if (!forceRefresh) {
    try {
      var chunkCount = cache.get(cacheKey + '_chunks');
      if (chunkCount) {
        var assembled = '';
        for (var c = 0; c < parseInt(chunkCount); c++) {
          var chunk = cache.get(cacheKey + '_chunk_' + c);
          if (!chunk) { assembled = null; break; }
          assembled += chunk;
        }
        if (assembled) {
          _MEMOIZED_RAW_DATA = JSON.parse(assembled);
          return _MEMOIZED_RAW_DATA;
        }
      }
    } catch(e) {
      Logger.log('Cache read error: ' + e.message);
    }
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(QUALITY_SHEET_NAME);
  if (!sheet) return [];

  var raw = sheet.getDataRange().getValues();
  if (raw.length < 2) return [];

  var data = raw.slice(1);
  _MEMOIZED_RAW_DATA = data;

  try {
    var serialized = JSON.stringify(data);
    var chunkSize = 90000;
    var chunks = [];
    for (var ci = 0; ci < serialized.length; ci += chunkSize) {
      chunks.push(serialized.slice(ci, ci + chunkSize));
    }
    for (var j = 0; j < chunks.length; j++) {
      cache.put(cacheKey + '_chunk_' + j, chunks[j], 600);
    }
    cache.put(cacheKey + '_chunks', String(chunks.length), 600);
  } catch(e) {
    Logger.log('Cache write error: ' + e.message);
  }

  return data;
}

function getAvailableQualityMonths() {
  var rows = getRawQualityData();
  var seen = {};
  for (var i = 0; i < rows.length; i++) {
    var month = rows[i][Q_COLS.REVIEW_MONTH];
    if (month) {
      if (month instanceof Date) {
        month = Utilities.formatDate(month, Session.getScriptTimeZone(), 'yyyy-MM');
      }
      seen[month] = true;
    }
  }
  return Object.keys(seen).sort().reverse();
}

function normalizeQualityMonth(val) {
  if (!val) return '';
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM');
  }
  return String(val).trim();
}

function normalizeLdap(val) {
  if (!val) return '';
  return String(val).trim().toLowerCase().split('@')[0];
}

function formatDate(val) {
  if (!val) return '';
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(val);
}

// ── AGGREGATION ───────────────────────────────────────────────────────────

function aggregateQualityRows(rows) {
  if (!rows || rows.length === 0) return null;

  var customerSum = 0, businessSum = 0, complianceSum = 0;
  var params = {};

  for (var i = 0; i < Q_PARAM_COLS.length; i++) {
    params[Q_PARAM_COLS[i]] = { yes: 0, total: 0 };
  }

  for (var j = 0; j < rows.length; j++) {
    var r = rows[j];
    customerSum += (parseFloat(r[Q_COLS.CUSTOMER_CRITICAL]) || 0);
    businessSum += (parseFloat(r[Q_COLS.BUSINESS_CRITICAL]) || 0);
    complianceSum += (parseFloat(r[Q_COLS.COMPLIANCE_CRITICAL]) || 0);

    for (var k = 0; k < Q_PARAM_COLS.length; k++) {
      var p = Q_PARAM_COLS[k];
      var val = String(r[Q_COLS[p]]).trim().toLowerCase();
      if (val === 'yes' || val === 'no' || val === '1' || val === '0') {
        params[p].total++;
        if (val === 'yes' || val === '1') params[p].yes++;
      }
    }
  }

  var count = rows.length;
  var paramScores = {};
  for (var l = 0; l < Q_PARAM_COLS.length; l++) {
    var pCol = Q_PARAM_COLS[l];
    paramScores[pCol] = params[pCol].total > 0 ? (params[pCol].yes / params[pCol].total) * 100 : null;
  }

  var groupedParams = {
    customer: Q_PARAM_GROUPS.customer.map(function(p) { return { name: p, score: paramScores[p] }; }),
    business: Q_PARAM_GROUPS.business.map(function(p) { return { name: p, score: paramScores[p] }; }),
    compliance: Q_PARAM_GROUPS.compliance.map(function(p) { return { name: p, score: paramScores[p] }; })
  };

  return {
    customer: (customerSum / count) * 100,
    business: (businessSum / count) * 100,
    compliance: (complianceSum / count) * 100,
    count: count,
    params: paramScores,
    groupedParams: groupedParams,
    targets: Q_TARGETS
  };
}

function aggregateTrends(rows) {
  var daily = {};
  var weekly = {};

  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    var dateRaw = r[Q_COLS.REVIEW_DATE];
    var date = formatDate(dateRaw);
    var week = r[Q_COLS.REVIEW_WEEK];

    var trends = [ {obj: daily, key: date}, {obj: weekly, key: week} ];
    for (var j = 0; j < trends.length; j++) {
      var t = trends[j];
      if (!t.key) continue;
      if (!t.obj[t.key]) t.obj[t.key] = { customer: 0, business: 0, compliance: 0, count: 0 };
      t.obj[t.key].customer += (parseFloat(r[Q_COLS.CUSTOMER_CRITICAL]) || 0);
      t.obj[t.key].business += (parseFloat(r[Q_COLS.BUSINESS_CRITICAL]) || 0);
      t.obj[t.key].compliance += (parseFloat(r[Q_COLS.COMPLIANCE_CRITICAL]) || 0);
      t.obj[t.key].count++;
    }
  }

  var formatTrend = function(obj) {
    return Object.keys(obj).sort().map(function(k) {
      var d = obj[k];
      return {
        label: k,
        customer: (d.customer / d.count) * 100,
        business: (d.business / d.count) * 100,
        compliance: (d.compliance / d.count) * 100,
        avg: ((d.customer + d.business + d.compliance) / (d.count * 3)) * 100
      };
    });
  };

  return { daily: formatTrend(daily), weekly: formatTrend(weekly) };
}

// ── CLIENT WRAPPERS ───────────────────────────────────────────────────────

function clientGetAvailableQualityMonths(forceRefresh) {
  if (forceRefresh) {
    _MEMOIZED_RAW_DATA = null;
    var cache = CacheService.getScriptCache();
    cache.remove('quality_raw_v1_chunks');
  }
  return getAvailableQualityMonths();
}

function clientGetMyQuality(ldap, month) {
  if (!ldap) ldap = Session.getActiveUser().getEmail().split('@')[0];

  var allRows = getRawQualityData();
  var filtered = [];
  for (var i = 0; i < allRows.length; i++) {
    var r = allRows[i];
    if (normalizeLdap(r[Q_COLS.AGENT_LDAP]) === normalizeLdap(ldap) &&
        normalizeQualityMonth(r[Q_COLS.REVIEW_MONTH]) === month) {
      filtered.push(r);
    }
  }

  var stats = aggregateQualityRows(filtered);
  var trends = aggregateTrends(filtered);

  var caseLog = filtered.map(function(r) {
    return {
      caseId: r[Q_COLS.CASE_ID],
      reviewDate: formatDate(r[Q_COLS.REVIEW_DATE]),
      customer: r[Q_COLS.CUSTOMER_CRITICAL],
      business: r[Q_COLS.BUSINESS_CRITICAL],
      compliance: r[Q_COLS.COMPLIANCE_CRITICAL],
      comments: r[Q_COLS.REVIEWER_COMMENTS],
      details: {
        customer: {
          LISTENING: r[Q_COLS.LISTENING],
          PROBING: r[Q_COLS.PROBING],
          COMPLETE_RESOLUTION: r[Q_COLS.COMPLETE_RESOLUTION],
          TROUBLESHOOTING: r[Q_COLS.TROUBLESHOOTING],
          USER_EXPECTATIONS: r[Q_COLS.USER_EXPECTATIONS],
          EMPATHY: r[Q_COLS.EMPATHY],
          OWNERSHIP: r[Q_COLS.OWNERSHIP],
          REFUNDS: r[Q_COLS.REFUNDS],
          RESPONSIVENESS: r[Q_COLS.RESPONSIVENESS]
        },
        business: {
          CONSULTS_ESCALATIONS: r[Q_COLS.CONSULTS_ESCALATIONS],
          CASE_DETAILS: r[Q_COLS.CASE_DETAILS],
          CATEGORIZATION: r[Q_COLS.CATEGORIZATION],
          CSAT_REMINDER: r[Q_COLS.CSAT_REMINDER],
          CASE_STATE: r[Q_COLS.CASE_STATE],
          OPENING_CLOSING: r[Q_COLS.OPENING_CLOSING],
          LANGUAGE_PROFICIENCY: r[Q_COLS.LANGUAGE_PROFICIENCY]
        },
        compliance: {
          AUTHENTICATION: r[Q_COLS.AUTHENTICATION],
          GOOGLE_ONLY_INFO: r[Q_COLS.GOOGLE_ONLY_INFO],
          PROFESSIONAL_CONDUCT: r[Q_COLS.PROFESSIONAL_CONDUCT],
          PAYMENT_COMPLAINTS: r[Q_COLS.PAYMENT_COMPLAINTS]
        }
      }
    };
  });

  return {
    ldap: ldap,
    month: month,
    stats: stats,
    trends: trends,
    caseLog: caseLog,
    hasData: filtered.length > 0
  };
}

function clientGetTeamQuality(supervisor, month) {
  var allRows = getRawQualityData();
  var filtered = [];
  for (var i = 0; i < allRows.length; i++) {
    var r = allRows[i];
    if (String(r[Q_COLS.SUPERVISOR]).trim() === String(supervisor).trim() &&
        normalizeQualityMonth(r[Q_COLS.REVIEW_MONTH]) === month) {
      filtered.push(r);
    }
  }

  var stats = aggregateQualityRows(filtered);
  var trends = aggregateTrends(filtered);

  var agentStats = {};
  var uniqueLdaps = [];
  for (var j = 0; j < filtered.length; j++) {
    var row = filtered[j];
    var ldap = normalizeLdap(row[Q_COLS.AGENT_LDAP]);
    if (!agentStats[ldap]) {
      agentStats[ldap] = [];
      uniqueLdaps.push(ldap);
    }
    agentStats[ldap].push(row);
  }

  var agents = uniqueLdaps.map(function(ldap) {
    return {
      ldap: ldap,
      stats: aggregateQualityRows(agentStats[ldap])
    };
  }).sort(function(a, b) {
    return (b.stats.customer + b.stats.business + b.stats.compliance) - (a.stats.customer + a.stats.business + a.stats.compliance);
  });

  return {
    supervisor: supervisor,
    month: month,
    stats: stats,
    trends: trends,
    agents: agents,
    hasData: filtered.length > 0
  };
}

function clientGetClusterQuality(manager, month) {
  var allRows = getRawQualityData();
  var filtered = [];
  for (var i = 0; i < allRows.length; i++) {
    var r = allRows[i];
    if (String(r[Q_COLS.MANAGER]).trim() === String(manager).trim() &&
        normalizeQualityMonth(r[Q_COLS.REVIEW_MONTH]) === month) {
      filtered.push(r);
    }
  }

  var stats = aggregateQualityRows(filtered);
  var trends = aggregateTrends(filtered);

  var supervisorStats = {};
  var uniqueSupervisors = [];
  for (var j = 0; j < filtered.length; j++) {
    var row = filtered[j];
    var sup = String(row[Q_COLS.SUPERVISOR]).trim();
    if (!supervisorStats[sup]) {
      supervisorStats[sup] = [];
      uniqueSupervisors.push(sup);
    }
    supervisorStats[sup].push(row);
  }

  var supervisors = uniqueSupervisors.map(function(sup) {
    return {
      name: sup,
      stats: aggregateQualityRows(supervisorStats[sup])
    };
  }).sort(function(a, b) {
    return (b.stats.customer + b.stats.business + b.stats.compliance) - (a.stats.customer + a.stats.business + a.stats.compliance);
  });

  return {
    manager: manager,
    month: month,
    stats: stats,
    trends: trends,
    supervisors: supervisors,
    hasData: filtered.length > 0
  };
}

function clientGetAllAgents() {
  var rows = getRawQualityData();
  var seen = {};
  for (var i = 0; i < rows.length; i++) {
    var ldap = normalizeLdap(rows[i][Q_COLS.AGENT_LDAP]);
    if (ldap) seen[ldap] = true;
  }
  return Object.keys(seen).sort().map(function(ldap) {
    return { ldap: ldap };
  });
}

function clientGetAllSupervisors() {
  var rows = getRawQualityData();
  var seen = {};
  for (var i = 0; i < rows.length; i++) {
    var sup = String(rows[i][Q_COLS.SUPERVISOR]).trim();
    if (sup) seen[sup] = true;
  }
  return Object.keys(seen).sort();
}

function clientGetAllManagers() {
  var rows = getRawQualityData();
  var seen = {};
  for (var i = 0; i < rows.length; i++) {
    var mgr = String(rows[i][Q_COLS.MANAGER]).trim();
    if (mgr) seen[mgr] = true;
  }
  return Object.keys(seen).sort();
}

function clientGetSession() {
  var email = Session.getActiveUser().getEmail();
  return {
    ldap: email.split('@')[0],
    email: email
  };
}

function clientGetInitialData(forceRefresh) {
  return {
    months: clientGetAvailableQualityMonths(forceRefresh),
    hierarchy: clientGetHierarchy(forceRefresh),
    session: clientGetSession(),
    targets: Q_TARGETS,
    cols: Q_COLS,
    paramGroups: Q_PARAM_GROUPS,
    paramCols: Q_PARAM_COLS
  };
}

function clientGetHierarchy(forceRefresh) {
  var cache = CacheService.getScriptCache();
  var cacheKey = 'quality_hierarchy_v1';

  if (forceRefresh) {
    cache.remove(cacheKey + '_chunks');
  }

  try {
    var chunkCount = cache.get(cacheKey + '_chunks');
    if (chunkCount && !forceRefresh) {
      var assembled = '';
      for (var c = 0; c < parseInt(chunkCount); c++) {
        var chunk = cache.get(cacheKey + '_chunk_' + c);
        if (!chunk) { assembled = null; break; }
        assembled += chunk;
      }
      if (assembled) return JSON.parse(assembled);
    }
  } catch(e) {
    Logger.log('Hierarchy cache read error: ' + e.message);
  }

  var rows = getRawQualityData();
  var hierarchy = {};

  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    var lob = String(r[Q_COLS.LOB] || 'Unknown LOB').trim();
    var sup = String(r[Q_COLS.SUPERVISOR] || 'Unknown Supervisor').trim();
    var agent = normalizeLdap(r[Q_COLS.AGENT_LDAP]);

    if (!agent) continue;

    if (!hierarchy[lob]) hierarchy[lob] = {};
    if (!hierarchy[lob][sup]) hierarchy[lob][sup] = {};

    // Using an object as a set for agents to ensure uniqueness
    hierarchy[lob][sup][agent] = true;
  }

  // Convert Agent sets to sorted arrays
  var result = {};
  var lobs = Object.keys(hierarchy).sort();
  for (var j = 0; j < lobs.length; j++) {
    var l = lobs[j];
    result[l] = {};
    var sups = Object.keys(hierarchy[l]).sort();
    for (var k = 0; k < sups.length; k++) {
      var s = sups[k];
      result[l][s] = Object.keys(hierarchy[l][s]).sort();
    }
  }

  try {
    var serialized = JSON.stringify(result);
    var chunkSize = 90000;
    var chunks = [];
    for (var ci = 0; ci < serialized.length; ci += chunkSize) {
      chunks.push(serialized.slice(ci, ci + chunkSize));
    }
    for (var j = 0; j < chunks.length; j++) {
      cache.put(cacheKey + '_chunk_' + j, chunks[j], 21600);
    }
    cache.put(cacheKey + '_chunks', String(chunks.length), 21600);
  } catch(e) {
    Logger.log('Hierarchy cache write error: ' + e.message);
  }

  return result;
}
