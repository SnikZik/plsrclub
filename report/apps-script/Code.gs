// ─────────────────────────────────────────────
//  SNIR. Client Dashboard — PLSRClub
//  GA4-only · no Sheets · no country filter
// ─────────────────────────────────────────────

var CONFIG = {
  GA4_PROPERTY_ID: "properties/395616952",
  DAYS_BACK: 30,
};

// ─────────────────────────────────────────────
//  doGet — ?startDate=YYYY-MM-DD&endDate=YYYY-MM-DD or ?days=N
// ─────────────────────────────────────────────
function doGet(e) {
  try {
    var startDate, endDate, label;
    var today = formatDate(new Date());

    if (e && e.parameter) {
      if (e.parameter.startDate && e.parameter.endDate) {
        startDate = e.parameter.startDate;
        endDate   = e.parameter.endDate;
        label     = e.parameter.label || null;
      } else if (e.parameter.days) {
        var d = parseInt(e.parameter.days, 10);
        if (d > 0 && d <= 365) {
          startDate = formatDate(daysAgo(d));
          endDate   = today;
        }
      }
    }

    if (!startDate) {
      startDate = formatDate(daysAgo(CONFIG.DAYS_BACK));
      endDate   = today;
    }

    var data = {
      ga4:     getGA4Data(startDate, endDate),
      updated: new Date().toISOString(),
      config:  { startDate: startDate, endDate: endDate, label: label },
    };
    return respond(data);
  } catch (err) {
    return respond({ error: err.message });
  }
}

// ─────────────────────────────────────────────
//  GA4
// ─────────────────────────────────────────────
function getGA4Data(startDate, endDate) {
  var token   = ScriptApp.getOAuthToken();
  var baseUrl = "https://analyticsdata.googleapis.com/v1beta/" + CONFIG.GA4_PROPERTY_ID + ":runReport";

  // Previous period of equal length
  var d1       = new Date(startDate);
  var d2       = new Date(endDate);
  var msDay    = 86400000;
  var periodMs = (d2 - d1) + msDay;
  var prevEnd   = formatDate(new Date(d1.getTime() - msDay));
  var prevStart = formatDate(new Date(d1.getTime() - periodMs));

  var overviewRes = apiPost(baseUrl, {
    dateRanges: [
      { startDate: startDate, endDate: endDate,   name: "current"  },
      { startDate: prevStart, endDate: prevEnd,    name: "previous" },
    ],
    metrics: [
      { name: "sessions" },
      { name: "totalUsers" },
      { name: "newUsers" },
      { name: "bounceRate" },
      { name: "averageSessionDuration" },
    ],
  }, token);

  var ecomRes = apiPost(baseUrl, {
    dateRanges: [
      { startDate: startDate, endDate: endDate,   name: "current"  },
      { startDate: prevStart, endDate: prevEnd,    name: "previous" },
    ],
    dimensions: [{ name: "eventName" }],
    metrics:    [{ name: "eventCount" }, { name: "purchaseRevenue" }],
    dimensionFilter: {
      filter: {
        fieldName: "eventName",
        inListFilter: { values: ["add_to_cart", "view_item", "purchase", "form_submit"] },
      },
    },
  }, token);

  var sourceRes = apiPost(baseUrl, {
    dateRanges: [{ startDate: startDate, endDate: endDate }],
    dimensions: [{ name: "sessionDefaultChannelGroup" }],
    metrics:    [{ name: "sessions" }],
    orderBys:   [{ metric: { metricName: "sessions" }, desc: true }],
    limit: 8,
  }, token);

  var convSourceRes = apiPost(baseUrl, {
    dateRanges: [{ startDate: startDate, endDate: endDate }],
    dimensions: [{ name: "sessionDefaultChannelGroup" }, { name: "eventName" }],
    metrics:    [{ name: "eventCount" }],
    dimensionFilter: {
      filter: {
        fieldName: "eventName",
        inListFilter: { values: ["view_item", "add_to_cart", "form_submit", "purchase"] },
      },
    },
  }, token);

  var durationBySourceRes = apiPost(baseUrl, {
    dateRanges: [{ startDate: startDate, endDate: endDate }],
    dimensions: [{ name: "sessionDefaultChannelGroup" }],
    metrics:    [{ name: "averageSessionDuration" }],
  }, token);

  var pagesRes = apiPost(baseUrl, {
    dateRanges: [{ startDate: startDate, endDate: endDate }],
    dimensions: [{ name: "pageTitle" }, { name: "pagePath" }],
    metrics:    [{ name: "sessions" }, { name: "screenPageViews" }],
    orderBys:   [{ metric: { metricName: "sessions" }, desc: true }],
    limit: 7,
  }, token);

  var sparkRes = apiPost(baseUrl, {
    dateRanges: [{ startDate: startDate, endDate: endDate }],
    dimensions: [{ name: "date" }],
    metrics:    [{ name: "sessions" }],
    orderBys:   [{ dimension: { dimensionName: "date" } }],
  }, token);

  var rows = overviewRes.rows || [];
  return {
    overview:            { current: extractMetricRow(rows, "current"), previous: extractMetricRow(rows, "previous") },
    ecommerce:           formatEcommerce(ecomRes),
    sources:             formatDimMetric(sourceRes),
    conversionsBySource: formatConversionsBySource(convSourceRes),
    durationBySource:    formatDurationBySource(durationBySourceRes),
    pages:               formatPages(pagesRes),
    sparkline:           formatSparkline(sparkRes),
  };
}

// ─────────────────────────────────────────────
//  Helpers
// ─────────────────────────────────────────────
function respond(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function apiPost(url, body, token) {
  var res = UrlFetchApp.fetch(url, {
    method: "POST", contentType: "application/json",
    headers: { Authorization: "Bearer " + token },
    payload: JSON.stringify(body), muteHttpExceptions: true,
  });
  return JSON.parse(res.getContentText());
}

function extractMetricRow(rows, name) {
  for (var i = 0; i < rows.length; i++) {
    var dv = rows[i].dimensionValues;
    if (dv && dv[0] && dv[0].value === name) {
      var mv = rows[i].metricValues;
      return {
        sessions: safeNum(mv,0), totalUsers: safeNum(mv,1), newUsers: safeNum(mv,2),
        bounceRate: safeFloat(mv,3), avgSessionDuration: safeFloat(mv,4),
      };
    }
  }
  return {};
}

function formatEcommerce(res) {
  var cur = {}, prev = {};
  (res.rows || []).forEach(function(r) {
    var name    = r.dimensionValues[0].value;
    var period  = r.dimensionValues[1] ? r.dimensionValues[1].value : "current";
    var count   = parseInt(r.metricValues[0].value, 10) || 0;
    var revenue = parseFloat(r.metricValues[1].value) || 0;
    if (period === "current")  { cur[name]  = { count: count, revenue: revenue }; }
    if (period === "previous") { prev[name] = { count: count, revenue: revenue }; }
  });
  return {
    addToCart:  { current: (cur["add_to_cart"]  || {}).count   || 0, previous: (prev["add_to_cart"]  || {}).count   || 0 },
    viewItem:   { current: (cur["view_item"]    || {}).count   || 0, previous: (prev["view_item"]    || {}).count   || 0 },
    formSubmit: { current: (cur["form_submit"]  || {}).count   || 0, previous: (prev["form_submit"]  || {}).count   || 0 },
    purchases:  { current: (cur["purchase"]     || {}).count   || 0, previous: (prev["purchase"]     || {}).count   || 0 },
    revenue:    { current: (cur["purchase"]     || {}).revenue || 0, previous: (prev["purchase"]     || {}).revenue || 0 },
  };
}

function formatDimMetric(res) {
  var rows  = res.rows || [];
  var total = rows.reduce(function(s,r){ return s+(parseInt(r.metricValues[0].value,10)||0); }, 0);
  return rows.map(function(r) {
    var v = parseInt(r.metricValues[0].value,10)||0;
    return { label: r.dimensionValues[0].value, value: v, pct: total>0?Math.round(v/total*100):0 };
  });
}

function formatConversionsBySource(res) {
  var result = {};
  (res.rows || []).forEach(function(r) {
    var channel   = r.dimensionValues[0].value;
    var eventName = r.dimensionValues[1].value;
    var count     = parseInt(r.metricValues[0].value, 10) || 0;
    if (!result[channel]) result[channel] = {};
    result[channel][eventName] = count;
  });
  return result;
}

function formatDurationBySource(res) {
  var result = {};
  (res.rows || []).forEach(function(r) {
    result[r.dimensionValues[0].value] = parseFloat(r.metricValues[0].value) || 0;
  });
  return result;
}

function formatPages(res) {
  return (res.rows||[]).map(function(r) {
    return { title: r.dimensionValues[0].value, path: r.dimensionValues[1].value,
             sessions: parseInt(r.metricValues[0].value,10)||0, views: parseInt(r.metricValues[1].value,10)||0 };
  });
}

function formatSparkline(res) {
  var dates=[], sessions=[];
  (res.rows||[]).forEach(function(r) {
    var d = r.dimensionValues[0].value;
    dates.push(d.slice(4,6)+"/"+d.slice(6,8));
    sessions.push(parseInt(r.metricValues[0].value,10)||0);
  });
  return { dates:dates, sessions:sessions };
}

function safeNum(arr,idx)   { return arr&&arr[idx]?(parseInt(arr[idx].value,10)||0):0; }
function safeFloat(arr,idx) { return arr&&arr[idx]?(parseFloat(arr[idx].value)||0):0; }
function daysAgo(n)         { var d=new Date(); d.setDate(d.getDate()-n); return d; }
function formatDate(d) {
  return d.getFullYear()+"-"+String(d.getMonth()+1).padStart(2,"0")+"-"+String(d.getDate()).padStart(2,"0");
}
