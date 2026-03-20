require('dotenv').config();
const express = require('express');
const axios = require('axios');
const cors = require('cors');

const app = express();
app.use(cors());
app.use(express.json());

// Helper function to calculate end time
function calculateEndTime(startTimeStr, timeSpentSeconds) {
  if (!startTimeStr) return 'N/A'; // Sometimes users don't specify a start time in Tempo

  const [hours, minutes, seconds] = startTimeStr.split(':').map(Number);
  const date = new Date();
  date.setHours(hours, minutes, seconds, 0);

  // Add the spent seconds
  date.setSeconds(date.getSeconds() + timeSpentSeconds);

  // Format back to HH:MM:SS
  return date.toTimeString().split(' ')[0];
}
app.post('/api/get-timesheet', async (req, res) => {
  const { code, email, startDate, endDate, format } = req.body;

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
    console.log("worklogs : ", worklogs.length);

    // STEP 3: Collect unique issue IDs (de-duplicated)
    const uniqueIssueIds = [...new Set(
      worklogs
        .filter(log => log.issue?.id)
        .map(log => log.issue.id)
    )];


    // STEP 4: Fetch issue details from Jira in parallel
    //         ⚠️ Use issue ID from Tempo — Jira resolves numeric IDs fine
    //            IF the service account has Browse Projects permission
    const jiraBase = 'https://e-emphasys.atlassian.net';
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

    // Build a lookup map: issueId -> { key, summary, ... }
    const issueMap = {};
    issueResults.forEach((result, idx) => {
      const id = uniqueIssueIds[idx];
      if (result.status === 'fulfilled') {
        const { key, fields } = result.value.data;
        issueMap[id] = {
          key,
          summary: fields.summary,
          status: fields.status?.name,
          issueType: fields.issuetype?.name
        };
      } else {
        // Log but don't crash — some issues may be in restricted projects
        console.warn(`⚠️ Could not fetch issue ${id}:`, result.reason?.response?.status);
        issueMap[id] = { key: `ID-${id}`, summary: 'Access restricted', status: 'Unknown' };
      }
    });

    // STEP 5: Enrich worklogs with issue details
    const enrichedLogs = worklogs.map(log => ({
      date: log.startDate,
      timeSpentHours: (log.timeSpentSeconds / 3600).toFixed(2),
      issueKey: issueMap[log.issue?.id]?.key ?? 'N/A',
      issueSummary: issueMap[log.issue?.id]?.summary ?? 'N/A',
      issueStatus: issueMap[log.issue?.id]?.status ?? 'N/A',
      description: format === 2
        ? (issueMap[log.issue?.id]?.key + " - " + (log.description ?? ''))
        : (log.description ?? ''),
      endTime: calculateEndTime(log.startTime, log.timeSpentSeconds),
      startTime: log.startTime ?? 'N/A',
    }));

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
