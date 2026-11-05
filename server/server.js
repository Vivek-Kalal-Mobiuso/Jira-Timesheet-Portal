require('dotenv').config();
const express = require('express');
const axios = require('axios');
const cors = require('cors');

const app = express();
app.use(cors());
app.use(express.json());

const JIRA_BASE = 'https://e-emphasys.atlassian.net';

// Tempo work attribute key that holds the "Task Type" of a worklog
const TASK_TYPE_ATTRIBUTE_KEY = '_TaskType_';

// Jira / Tempo task type -> timesheet task type. Unmapped values are sent as blank.
const TASK_TYPE_MAP = {
  'Holiday': 'Leave',
  'Leave': 'Leave',
  'Managerial-InternalCalls': 'Internal Meeting',
  'Tech-Dev-Coding': 'Development',
  'Tech-Dev-UnitTesting': 'Manual Testing',
  'Research-Feasibility': 'Research',
  'Research-Exploration': 'Research',
  'Research-Exploration': 'Research',
  'Tech-Dev-Code Review': 'Code Review',
  'Functional-Dev-Requirement Analysis': 'Requirement Gathering',
  'All-Customer Calls': 'Client Meeting',
  'QA-Dev-TestCasePrep': 'Manual Testing',
};

// Tempo/Jira date -> YYYY-MM-DD (plain string, no locale formatting)
function normalizeDateYYYYMMDD(value) {
  if (value == null || value === '') return '';
  const iso = String(value).slice(0, 10);
  if (/^\d{4}-\d{2}-\d{2}$/.test(iso)) return iso;
  const d = new Date(value);
  if (Number.isNaN(d.getTime())) return iso;
  return d.toISOString().slice(0, 10);
}

// "HH:MM:SS" -> seconds since midnight (null when missing)
function timeToSeconds(timeStr) {
  if (!timeStr) return null;
  const [h = 0, m = 0, s = 0] = timeStr.split(':').map(Number);
  return h * 3600 + m * 60 + s;
}

// seconds since midnight -> "HH:MM"
function secondsToTime(totalSeconds) {
  const secs = ((totalSeconds % 86400) + 86400) % 86400;
  const pad = (n) => String(n).padStart(2, '0');
  return `${pad(Math.floor(secs / 3600))}:${pad(Math.floor((secs % 3600) / 60))}`;
}

// Raw Task Type value from the worklog's Tempo work attributes, mapped to the timesheet value
function getMappedTaskType(log) {
  const raw = log.attributes?.values?.find((v) => v.key === TASK_TYPE_ATTRIBUTE_KEY)?.value;
  return (raw && TASK_TYPE_MAP[raw]) || '';
}

const DEFAULT_DESCRIPTION_INCLUDE = {
  url: false,
  issueKey: false,
  summary: false,
  description: true,
};

// Builds Task Details from checklist flags; parts joined with blank lines in fixed order
function normalizeDescriptionInclude(raw) {
  if (!raw || typeof raw !== 'object') return { ...DEFAULT_DESCRIPTION_INCLUDE };
  return {
    url: Boolean(raw.url),
    issueKey: Boolean(raw.issueKey),
    summary: Boolean(raw.summary),
    description: raw.description === undefined
      ? DEFAULT_DESCRIPTION_INCLUDE.description
      : Boolean(raw.description),
  };
}

// Order in cell: Summary, Issue URL, Issue Key, Description (blank line between blocks)
function buildDescription(descriptionInclude, issue, description) {
  const flags = normalizeDescriptionInclude(descriptionInclude);
  const desc = (description ?? '').trim();
  const parts = [];

  if (flags.summary && issue.summary?.trim()) parts.push(issue.summary.trim());
  if (flags.url && issue.url?.trim()) parts.push(issue.url.trim());
  if (flags.issueKey && issue.key?.trim()) parts.push(issue.key.trim());
  if (flags.description && desc) parts.push(desc);

  return parts.join('\n\n');
}

// Merges the two worklogs of the same issue that surround a single real break.
// Rules (per issue, per date):
//   - a "gap" is time between two consecutive worklogs of the issue (next start > previous end)
//   - the gap is a "real break" only if no worklog of another issue overlaps it
//   - exactly one real break  -> the two worklogs around it become one row with break start/end
//   - zero or more than one   -> nothing is merged, no breaks are reported
//   - back-to-back worklogs (no gap) are never merged
function groupAndMerge(entries) {
  const rows = [];
  const mergedEntries = new Set();

  const timed = entries.filter((e) => e.startSec !== null);
  const groups = new Map();
  timed.forEach((e) => {
    const k = `${e.date}|${e.issueId}`;
    if (!groups.has(k)) groups.set(k, []);
    groups.get(k).push(e);
  });

  groups.forEach((group) => {
    group.sort((a, b) => a.startSec - b.startSec);

    const realBreaks = [];
    for (let i = 0; i < group.length - 1; i++) {
      const gapStart = group[i].endSec;
      const gapEnd = group[i + 1].startSec;
      if (gapEnd <= gapStart) continue; // back-to-back or overlapping, not a gap

      const otherIssueInGap = timed.some(
        (o) =>
          o.date === group[i].date &&
          o.issueId !== group[i].issueId &&
          o.startSec < gapEnd &&
          o.endSec > gapStart
      );
      if (!otherIssueInGap) realBreaks.push({ index: i, gapStart, gapEnd });
    }

    if (realBreaks.length === 1) {
      const { index, gapStart, gapEnd } = realBreaks[0];
      const first = group[index];
      const second = group[index + 1];
      rows.push({
        sortStart: first.startSec,
        date: first.date,
        issue: first.issue,
        logs: [first.log, second.log],
        timeSpentSeconds: first.log.timeSpentSeconds + second.log.timeSpentSeconds,
        startSec: first.startSec,
        endSec: second.endSec,
        breakStartSec: gapStart,
        breakEndSec: gapEnd,
      });
      mergedEntries.add(first);
      mergedEntries.add(second);
    }
  });

  entries.forEach((e) => {
    if (mergedEntries.has(e)) return;
    rows.push({
      sortStart: e.startSec,
      date: e.date,
      issue: e.issue,
      logs: [e.log],
      timeSpentSeconds: e.log.timeSpentSeconds,
      startSec: e.startSec,
      endSec: e.endSec,
      breakStartSec: null,
      breakEndSec: null,
    });
  });

  // Date, then start time (worklogs without a start time go last for that date)
  rows.sort((a, b) => {
    if (a.date !== b.date) return a.date < b.date ? -1 : 1;
    if (a.sortStart === null && b.sortStart === null) return 0;
    if (a.sortStart === null) return 1;
    if (b.sortStart === null) return -1;
    return a.sortStart - b.sortStart;
  });

  return rows;
}
app.post('/api/get-timesheet', async (req, res) => {
  const { code, email, startDate, endDate, descriptionInclude } = req.body;

  if (!code) {
    return res.status(400).json({ error: 'Authorization code is required' });
  }

  try {
    // ====================================================================
    // STEP 1: GET THE KEY (Exchange Code for Tempo Token)
    // ====================================================================
    const tokenResponse = await axios.post('https://api.tempo.io/oauth/token/', new URLSearchParams({
      grant_type: 'authorization_code',
      client_id: process.env.CLIENT_ID,
      client_secret: process.env.CLIENT_SECRET,
      redirect_uri: process.env.REDIRECT_URI,
      code: code,
    }), { headers: { 'Content-Type': 'application/x-www-form-urlencoded' } }
    );

    const tempoToken = tokenResponse?.data?.access_token;

    // ====================================================================
    // STEP 2: FETCH AccountId using email id
    // ====================================================================
    const jiraAuth = Buffer.from(`${process.env.JIRA_EMAIL}:${process.env.JIRA_API_TOKEN}`).toString('base64');
    const userSearchResponse = await axios.get(`https://e-emphasys.atlassian.net/rest/api/3/user/search?query=${email}`, {
      headers: { 'Authorization': `Basic ${jiraAuth}`, 'Accept': 'application/json' }
    });
    
    if (!userSearchResponse.data || userSearchResponse.data.length === 0) {
      // Return 404 to React so you can show a friendly error to the user
      return res.status(404).json({ error: 'Jira user not found. Ensure the email/name is public or the Service Account is an Admin.' });
    }

    const accountId = userSearchResponse.data[0].accountId;

    const tempoResponse = await axios.get(
      `https://api.tempo.io/4/worklogs/user/${accountId}`,
      {
        params: {
          from: startDate,
          to: endDate,
          limit: 300,   // ← max allowed by Tempo API
          offset: 0,
        },
        headers: { Authorization: `Bearer ${tempoToken}` },
      }
    );

    const worklogs = tempoResponse.data.results;
    if (!worklogs || worklogs.length === 0) return res.json([]);
    // STEP 3: Collect unique issue IDs (de-duplicated)
    const uniqueIssueIds = [...new Set(
      worklogs
        .filter(log => log.issue?.id)
        .map(log => log.issue.id)
    )];


    // STEP 4: Fetch issue details from Jira in parallel
    //         ⚠️ Use issue ID from Tempo — Jira resolves numeric IDs fine
    //            IF the service account has Browse Projects permission
    const jiraBase = JIRA_BASE;
    const issueResults = await Promise.allSettled(
      uniqueIssueIds.map(id =>
        axios.get(`${jiraBase}/rest/api/2/issue/${id}`, {
          params: { fields: 'summary,status,issuetype,assignee' }, // only fetch what you need
          headers: {
            Authorization: `Basic ${jiraAuth}`,
            Accept: 'application/json',
          },
        })
      )
    );

    // Build a lookup map: issueId -> { key, url, summary, ... }
    const issueMap = {};
    issueResults.forEach((result, idx) => {
      const id = uniqueIssueIds[idx];
      if (result.status === 'fulfilled') {
        const { key, fields } = result.value.data;
        issueMap[id] = {
          key,
          url: `${jiraBase}/browse/${key}`,
          summary: fields.summary,
          status: fields.status?.name,
          issueType: fields.issuetype?.name
        };
      } else {
        // Log but don't crash — some issues may be in restricted projects
        console.warn(`⚠️ Could not fetch issue ${id}:`, result.reason?.response?.status);
        issueMap[id] = { key: `ID-${id}`, url: '', summary: 'Access restricted', status: 'Unknown' };
      }
    });

    // STEP 5: Enrich worklogs with issue details
    const fallbackIssue = { key: 'N/A', url: '', summary: 'N/A', status: 'N/A' };
    const entries = worklogs.map(log => {
      const startSec = timeToSeconds(log.startTime);
      return {
        log,
        date: normalizeDateYYYYMMDD(log.startDate),
        issueId: log.issue?.id,
        issue: issueMap[log.issue?.id] ?? fallbackIssue,
        startSec,
        endSec: startSec === null ? null : startSec + log.timeSpentSeconds,
      };
    });

    const rows = groupAndMerge(entries);

    const enrichedLogs = rows.map(row => {
      // De-duplicated worklog descriptions, one per line (relevant for merged rows)
      const descriptions = [...new Set(
        row.logs.map(l => (l.description ?? '').trim()).filter(Boolean)
      )].join('\n');

      return {
        date: row.date,
        timeSpentHours: (row.timeSpentSeconds / 3600).toFixed(2),
        issueKey: row.issue.key,
        issueUrl: row.issue.url ?? '',
        issueSummary: row.issue.summary ?? 'N/A',
        issueStatus: row.issue.status ?? 'N/A',
        worklogDescription: descriptions,
        description: buildDescription(descriptionInclude, row.issue, descriptions),
        taskType: getMappedTaskType(row.logs[0]),
        startTime: row.startSec === null ? 'N/A' : secondsToTime(row.startSec),
        endTime: row.endSec === null ? 'N/A' : secondsToTime(row.endSec),
        breakStart: row.breakStartSec === null ? '' : secondsToTime(row.breakStartSec),
        breakEnd: row.breakEndSec === null ? '' : secondsToTime(row.breakEndSec),
      };
    });

    res.json(enrichedLogs);
  } catch (error) {
    console.error('Tempo API Error:', error.response?.data || error.message);
    res.status(500).json({ error: 'Failed to fetch timesheet' });
  }
});

const PORT = process.env.PORT || 5000;
app.listen(PORT, () => {
  console.log(`✓ Server running on http://localhost:${PORT}`);
});
