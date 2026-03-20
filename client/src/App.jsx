import React, { useState, useEffect, useMemo } from 'react';
import * as XLSX from 'xlsx';
import { saveAs } from 'file-saver';

/* ─── Floating Bubble ─── */
const Bubble = ({ size, left, delay, duration }) => (
  <div
    className="bubble"
    style={{
      width: size,
      height: size,
      left: `${left}%`,
      animationDelay: `${delay}s`,
      animationDuration: `${duration}s`,
    }}
  />
);

/* ─── Jira Icon (SVG) ─── */
const JiraIcon = () => (
  <svg width="22" height="22" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
    <path d="M23.323 11.33L13.4 1.406 12 0 4.567 7.433.677 11.33a.96.96 0 000 1.347l6.76 6.76L12 24l7.433-7.433.193-.193 3.697-3.697a.96.96 0 000-1.347zM12 15.56L8.44 12 12 8.44 15.56 12 12 15.56z" fill="currentColor" />
  </svg>
);

/* ─── Download Icon (SVG) ─── */
const DownloadIcon = () => (
  <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    <path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4" />
    <polyline points="7 10 12 15 17 10" />
    <line x1="12" y1="15" x2="12" y2="3" />
  </svg>
);

/* ─── Check Icon (SVG) ─── */
const CheckIcon = () => (
  <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="3" strokeLinecap="round" strokeLinejoin="round">
    <polyline points="20 6 9 17 4 12" />
  </svg>
);

/* ─── Main App ─── */
/* ─── Email Icon (SVG) ─── */
const EmailIcon = () => (
  <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    <rect x="2" y="4" width="20" height="16" rx="2" />
    <path d="M22 7l-10 7L2 7" />
  </svg>
);

/* ─── Calendar Icon (SVG) ─── */
const CalendarIcon = () => (
  <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    <rect x="3" y="4" width="18" height="18" rx="2" ry="2" />
    <line x1="16" y1="2" x2="16" y2="6" />
    <line x1="8" y1="2" x2="8" y2="6" />
    <line x1="3" y1="10" x2="21" y2="10" />
  </svg>
);

const App = () => {
  const [isAuthorized, setIsAuthorized] = useState(false);
  const [loading, setLoading] = useState(false);
  const [downloadSuccess, setDownloadSuccess] = useState(false);
  const [error, setError] = useState('');
  const [email, setEmail] = useState('');
  const [startDate, setStartDate] = useState('');
  const [endDate, setEndDate] = useState('');
  const [format, setFormat] = useState(1);

  // Generate random bubbles once
  const bubbles = useMemo(() => {
    return Array.from({ length: 12 }, (_, i) => ({
      id: i,
      size: Math.random() * 120 + 40,
      left: Math.random() * 100,
      delay: Math.random() * 10,
      duration: Math.random() * 12 + 10,
    }));
  }, []);

  // Check for OAuth code on mount
  useEffect(() => {
    const urlParams = new URLSearchParams(window.location.search);
    const code = urlParams.get('code');

    if (code) {
      window.history.replaceState({}, document.title, '/');
      setIsAuthorized(true);
      sessionStorage.setItem('tempo_auth_code', code);
    }
  }, []);

  // Redirect to Jira OAuth
  const handleLogin = () => {
    const clientId = 'DmZVzB9mAQCsW4QsXCORYYiUzOE0GRdVsVNmgWB6dRHEfD0gDg';
    const redirectUri = 'http://localhost:5173/callback';
    const jiraUrl = 'https://e-emphasys.atlassian.net';

    const authUrl = `https://api.tempo.io/oauth/authorize/redirect?client_id=${clientId}&redirect_uri=${redirectUri}&jira_url=${jiraUrl}`;
    window.location.href = authUrl;
  };

  // Validate form
  const isFormValid = email.trim() && startDate && endDate && startDate <= endDate;

  // Fetch timesheet and generate Excel
  const handleDownload = async () => {
    if (!isFormValid) {
      setError('Please fill in all fields and ensure dates are valid.');
      setTimeout(() => setError(''), 5000);
      return;
    }

    setLoading(true);
    setError('');
    setDownloadSuccess(false);
    const authCode = sessionStorage.getItem('tempo_auth_code');

    try {
      const baseUrl = import.meta.env.REACT_APP_API_URL || '/api';
      const response = await fetch(`${baseUrl}/get-timesheet`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          code: authCode,
          email: email.trim(),
          startDate,
          endDate,
          format,
        }),
      });

      if (!response.ok) {
        const errData = await response.json().catch(() => ({}));
        throw new Error(errData.error || 'Failed to fetch timesheet data');
      }

      const data = await response.json();

      if (!data || data.length === 0) {
        setError('No worklogs found for the selected period.');
        setTimeout(() => setError(''), 5000);
        return;
      }

      const excelData = data.map((log) => ({
        Date: log.date,
        'Issue': log.issueKey || 'Unknown',
        Description: log.description || '',
        Hours: log.timeSpentHours,
        'Start Time': log.startTime || 'N/A',
        'End Time': log.endTime || 'N/A',
      }));

      const worksheet = XLSX.utils.json_to_sheet(excelData);
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, 'Timesheet');

      const excelBuffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
      const blob = new Blob([excelBuffer], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8',
      });
      saveAs(blob, `Timesheet_${startDate}_to_${endDate}.xlsx`);
      setDownloadSuccess(true);
      // Auth code is single-use — refresh after a short delay so user can re-authorize
      setTimeout(() => {
        sessionStorage.removeItem('tempo_auth_code');
        window.location.reload();
      }, 2000);
    } catch (err) {
      console.error('Error downloading timesheet:', err);
      setError(err.message || 'Something went wrong. Please try again.');
      setTimeout(() => setError(''), 5000);
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="ocean-bg min-h-screen flex items-center justify-center p-4">
      {/* Floating Bubbles */}
      {bubbles.map((b) => (
        <Bubble key={b.id} size={b.size} left={b.left} delay={b.delay} duration={b.duration} />
      ))}

      {/* Glassmorphism Card */}
      <div className="glass-card rounded-2xl p-5 w-full max-w-md relative z-10 fade-in">
        {/* Header */}
        <div className="text-center mb-2">
          <div className="inline-flex items-center justify-center w-16 h-16 rounded-2xl bg-white/10 border border-white/20 mb-5">
            <svg width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="white" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round">
              <rect x="3" y="4" width="18" height="18" rx="2" ry="2" />
              <line x1="16" y1="2" x2="16" y2="6" />
              <line x1="8" y1="2" x2="8" y2="6" />
              <line x1="3" y1="10" x2="21" y2="10" />
              <line x1="8" y1="14" x2="12" y2="14" />
              <line x1="8" y1="18" x2="16" y2="18" />
            </svg>
          </div>
          <h1 className="text-2xl font-bold text-white tracking-tight mb-2">
            Timesheet Portal
          </h1>
          <p className="text-sm text-white/50 font-medium">
            Jira &middot; Tempo Worklogs
          </p>
        </div>

        {/* Divider */}
        <div className="w-full h-px bg-gradient-to-r from-transparent via-white/20 to-transparent mb-4" />

        {/* Content */}
        {!isAuthorized ? (
          <div className="text-center fade-in">
            <div className="mb-2">
              <span className="status-badge status-disconnected">
                <span className="w-2 h-2 rounded-full bg-white/40" />
                Not Connected
              </span>
            </div>
            <p className="text-white/70 text-sm leading-relaxed mb-6">
              Connect your Jira account to generate and download your timesheet as an Excel file.
            </p>
            <div className="flex justify-center mt-2">
              <button
                id="authorize-btn"
                onClick={handleLogin}
                className="glass-btn py-3 px-10 rounded-2xl text-lg flex items-center justify-center gap-2.5 tracking-wide"
              >
                <JiraIcon />
                Authorize with Jira
              </button>
            </div>
          </div>
        ) : (
          <div className="fade-in">
            <div className="text-center mb-5">
              <span className="status-badge status-connected">
                <span className="pulse-dot" />
                Connected
              </span>
            </div>

            {/* Email Input */}
            <div className="mb-5">
              <label className="flex items-center gap-1.5 text-white/60 text-xs font-medium mb-2 tracking-wide uppercase">
                <EmailIcon />
                Jira Email
              </label>
              <input
                id="email-input"
                type="email"
                value={email}
                onChange={(e) => setEmail(e.target.value)}
                placeholder="name@company.com"
                className="glass-input w-full py-3 px-4 rounded-2xl text-sm text-white placeholder-white/30 outline-none"
              />
            </div>

            {/* Date Range */}
            <div className="grid grid-cols-2 gap-4 mb-8">
              <div>
                <label className="flex items-center gap-1.5 text-white/60 text-xs font-medium mb-2 tracking-wide uppercase">
                  <CalendarIcon />
                  From
                </label>
                <input
                  id="start-date-input"
                  type="date"
                  value={startDate}
                  onChange={(e) => setStartDate(e.target.value)}
                  className="glass-input w-full py-3 px-4 rounded-2xl text-sm text-white outline-none"
                />
              </div>
              <div>
                <label className="flex items-center gap-1.5 text-white/60 text-xs font-medium mb-2 tracking-wide uppercase">
                  <CalendarIcon />
                  To
                </label>
                <input
                  id="end-date-input"
                  type="date"
                  value={endDate}
                  onChange={(e) => setEndDate(e.target.value)}
                  min={startDate}
                  className="glass-input w-full py-3 px-4 rounded-2xl text-sm text-white outline-none"
                />
              </div>
            </div>

            {/* Format Selector */}
            <div className="mb-8">
              <label className="flex items-center gap-1.5 text-white/60 text-xs font-medium mb-2.5 tracking-wide uppercase">
                Description Format
              </label>
              <div className="flex gap-3">
                <button
                  type="button"
                  onClick={() => setFormat(1)}
                  className={`format-chip ${format === 1 ? 'format-chip-active' : ''
                    }`}
                >
                  Format 1
                </button>
                <button
                  type="button"
                  onClick={() => setFormat(2)}
                  className={`format-chip ${format === 2 ? 'format-chip-active' : ''
                    }`}
                >
                  Format 2
                </button>
              </div>
              <p className="text-white/30 text-[11px] mt-2">
                {format === 1
                  ? 'Description only'
                  : 'Issue Key - Description'}
              </p>
            </div>

            {/* Generate Button */}
            <div className="flex justify-center mt-2">
              <button
                id="generate-btn"
                onClick={handleDownload}
                disabled={loading || !isFormValid}
                className="glass-btn py-3.5 px-10 rounded-2xl text-sm flex items-center justify-center gap-2.5 tracking-wide"
              >
                {loading ? (
                  <>
                    <span className="spinner" />
                    Generating...
                  </>
                ) : (
                  <>
                    <DownloadIcon />
                    Generate Timesheet
                  </>
                )}
              </button>
            </div>
          </div>
        )}

        {/* Success Toast */}
        {downloadSuccess && (
          <div className="mt-5 flex items-center gap-2 justify-center text-emerald-300 text-sm fade-in">
            <CheckIcon />
            Timesheet downloaded successfully!
          </div>
        )}

        {/* Error Toast */}
        {error && (
          <div className="mt-5 text-center text-red-300 text-sm fade-in">
            {error}
          </div>
        )}

        {/* Footer */}
        <div className="mt-4 pt-5 border-t border-white/10 text-center">
          <p className="text-[11px] text-white/30 tracking-wide uppercase">
            Powered by Tempo &middot; Atlassian <br />
            <span className="text-white">Made By Vivek Kalal</span>
          </p>
        </div>
      </div>
    </div>
  );
};

export default App;
