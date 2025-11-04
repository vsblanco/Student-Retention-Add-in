import React, { useState } from 'react';
import Modal from '../utility/Modal';
import Calendar from '../utility/Calendar';
import { formatExcelDate } from '../utility/Conversion';
// Outreach trigger phrases (case-insensitive substring match)
const Contacted_Keywords = [
  "hung up",
  "hanged up",
  "promise",
  "requested",
  "up to date",
  "will catch up",
  "will come",
  "will complete",
  "will engage",
  "will pass",
  "will submit",
  "will work",
  "will be in class",
  "waiting for instructor",
  "waiting for professor",
  "waiting for teacher",
  "waiting on instructor",
  "waiting on professor",
  "waiting on teacher"
];

export const isOutreachTrigger = (text) => {
  if (!text || typeof text !== 'string') return { matched: false, tag: null, lda: { dates: [], keywords: [] } };

  const lower = text.toLowerCase();

  // Check Contacted keywords
  const contactedMatched = Contacted_Keywords.some(k => lower.includes(k.toLowerCase()));

  // Check LDA date-like keywords (now returns { dates: [], keywords: [] })
  let ldaResult = ldaKeywords(text || '');
  let ldaMatches = Array.isArray(ldaResult.dates) ? ldaResult.dates : [];
  let ldaFoundKeywords = Array.isArray(ldaResult.keywords) ? ldaResult.keywords : [];

  // If the text only contains numbers, try interpreting it as an Excel serial date
  const isNumericOnly = /^\s*\d+(\.\d+)?\s*$/.test(text || '');
  if (isNumericOnly) {
    try {
      const serial = Number(text.trim());
      if (!Number.isNaN(serial)) {
        const maybeDate = formatExcelDate(serial);

        // Normalize possible Date or string result into a string suitable for ldaKeywords
        let normalized = null;
        if (maybeDate instanceof Date && !isNaN(maybeDate.getTime())) {
          const mm = String(maybeDate.getMonth() + 1).padStart(2, '0');
          const dd = String(maybeDate.getDate()).padStart(2, '0');
          const yy = String(maybeDate.getFullYear()).slice(-2);
          normalized = `${mm}/${dd}/${yy}`;
        } else if (typeof maybeDate === 'string' && maybeDate.trim().length > 0) {
          normalized = maybeDate.trim();
        }

        if (normalized) {
          const parsedFromExcel = ldaKeywords(normalized);
          if (parsedFromExcel && Array.isArray(parsedFromExcel.dates) && parsedFromExcel.dates.length > 0) {
            ldaMatches = Array.from(new Set([...(Array.isArray(ldaMatches) ? ldaMatches : []), ...parsedFromExcel.dates]));
          }
          if (parsedFromExcel && Array.isArray(parsedFromExcel.keywords) && parsedFromExcel.keywords.length > 0) {
            ldaFoundKeywords = Array.from(new Set([...(Array.isArray(ldaFoundKeywords) ? ldaFoundKeywords : []), ...parsedFromExcel.keywords]));
          }
        }
      }
    } catch (e) {
      // ignore errors from formatExcelDate
    }
  }

  const tags = [];
  if (contactedMatched) tags.push('Contacted');
  if (Array.isArray(ldaMatches) && ldaMatches.length > 0) {
    // prefix each LDA date with "LDA "
    tags.push(...ldaMatches.map(d => `LDA ${d}`));
  }

  if (tags.length === 0) return { matched: false, tag: null, lda: { dates: ldaMatches, keywords: ldaFoundKeywords } };

  return { matched: true, tag: tags.join(', '), lda: { dates: ldaMatches, keywords: ldaFoundKeywords } };
};

// New helper: extract LDA match strings (normalized as MM/DD/YY) and the matched keywords separately
export function ldaKeywords(text) {
  if (!text || typeof text !== 'string') return { dates: [], keywords: [] };

  // If the text only contains numbers, try interpreting it as an Excel serial date first.
  const isNumericOnly = /^\s*\d+(\.\d+)?\s*$/.test(text);
  if (isNumericOnly) {
    try {
      const serial = Number(text.trim());
      if (!Number.isNaN(serial)) {
        const maybeDate = formatExcelDate(serial, 'short');
        // If formatExcelDate returns a string, trust it and return immediately.
        if (typeof maybeDate === 'string' && maybeDate.trim().length > 0) {
          return { dates: [maybeDate.trim()], keywords: [text.trim()] };
        }
      }
    } catch (e) {
      // ignore and continue with original text
    }
  }

  const keywords = [
    "Tomorrow", "next week",
    "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday",
    "weekend", "weekends"
  ];
  const shortDatePart = '\\b\\d{1,2}[\\/\\-]\\d{1,2}(?:[\\/\\-]\\d{2,4})?\\b';

  // long month formats: "November 8", "Nov 8th", "November 8, 2025", "Nov 8 25"
  const longMonthPart = '\\b(?:January|February|March|April|May|June|July|August|September|October|November|December|Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec)\\s+\\d{1,2}(?:st|nd|rd|th)?(?:[,\\s]+\\d{2,4})?\\b';

  const keywordPart = `\\b(?:${keywords.map(k => k.replace(/[.*+?^${}()|[\\]\\\\]/g, '\\$&')).join('|')})\\b`;
  const combinedRegex = new RegExp(`${shortDatePart}|${longMonthPart}|${keywordPart}`, 'gi');

  const results = [];
  const keywordsFound = [];
  let match;
  const now = new Date();

  const formatTwoDigit = d => {
    const mm = String(d.getMonth() + 1).padStart(2, '0');
    const dd = String(d.getDate()).padStart(2, '0');
    const yy = String(d.getFullYear()).slice(-2);
    return `${mm}/${dd}/${yy}`;
  };

  const weekdayIndex = name => {
    const map = { sunday:0, monday:1, tuesday:2, wednesday:3, thursday:4, friday:5, saturday:6 };
    return map[name.toLowerCase()];
  };

  const monthMap = {
    january:1, jan:1, february:2, feb:2, march:3, mar:3, april:4, apr:4, may:5,
    june:6, jun:6, july:7, jul:7, august:8, aug:8, september:9, sep:9, sept:9,
    october:10, oct:10, november:11, nov:11, december:12, dec:12
  };

  while ((match = combinedRegex.exec(text)) !== null) {
    const token = match[0];
    // record the raw token as a found keyword (dedupe later)
    keywordsFound.push(token);

    // numeric date e.g., 10/7, 10/7/25 or 10-7-2025
    const numeric = token.match(/^(\d{1,2})[\/\-](\d{1,2})(?:[\/\-](\d{2,4}))?$/);
    if (numeric) {
      let month = parseInt(numeric[1], 10);
      let day = parseInt(numeric[2], 10);
      let yearPart = numeric[3];
      let year;
      if (!yearPart) {
        year = now.getFullYear();
      } else if (yearPart.length === 2) {
        year = 2000 + parseInt(yearPart, 10);
      } else {
        year = parseInt(yearPart, 10);
      }
      const d = new Date(year, month - 1, day);
      if (!isNaN(d.getTime())) results.push(formatTwoDigit(d));
      continue;
    }

    // long month format e.g., "November 8th, 2025" or "Nov 8"
    const longMonthMatch = token.match(/^([A-Za-z]+)\s+(\d{1,2})(?:st|nd|rd|th)?(?:[,\s]+(\d{2,4}))?$/i);
    if (longMonthMatch) {
      const rawMonth = longMonthMatch[1].toLowerCase();
      const day = parseInt(longMonthMatch[2], 10);
      const yearPart = longMonthMatch[3];
      const monthNum = monthMap[rawMonth];
      if (monthNum && !Number.isNaN(day)) {
        let year;
        if (!yearPart) {
          year = now.getFullYear();
        } else if (yearPart.length === 2) {
          year = 2000 + parseInt(yearPart, 10);
        } else {
          year = parseInt(yearPart, 10);
        }
        const d = new Date(year, monthNum - 1, day);
        if (!isNaN(d.getTime())) results.push(formatTwoDigit(d));
      }
      continue;
    }

    const lc = token.toLowerCase();
    if (lc === 'tomorrow') {
      const d = new Date(now);
      d.setDate(d.getDate() + 1);
      results.push(formatTwoDigit(d));
      continue;
    }
    if (lc === 'next week') {
      const d = new Date(now);
      d.setDate(d.getDate() + 7);
      results.push(formatTwoDigit(d));
      continue;
    }
    if (lc === 'weekend' || lc === 'weekends') {
      // return next Saturday
      const d = new Date(now);
      const target = 6; // Saturday
      const diff = (target + 7 - d.getDay()) % 7 || 7;
      d.setDate(d.getDate() + diff);
      results.push(formatTwoDigit(d));
      continue;
    }
    // weekday names -> next occurrence of that weekday
    const wIdx = weekdayIndex(token);
    if (typeof wIdx === 'number') {
      const d = new Date(now);
      const diff = (wIdx + 7 - d.getDay()) % 7 || 7;
      d.setDate(d.getDate() + diff);
      results.push(formatTwoDigit(d));
      continue;
    }
  }

  // dedupe and return normalized strings and keywords
  return {
    dates: Array.from(new Set(results)),
    keywords: Array.from(new Set(keywordsFound.map(k => k)))
  };
}

function DNCModal({ isOpen, onClose, phone, otherPhone, email, onSelect }) {
  const handleSelect = value => {
    if (onSelect) onSelect(value);
    if (onClose) onClose();
  };

  return (
    <Modal
      isOpen={isOpen}
      onClose={onClose}
      padding="12px"
      overlayStyle={{
        background: "rgba(0,0,0,0.35)", // darker background
        backdropFilter: "blur(5px) saturate(140%)",
        WebkitBackdropFilter: "blur(5px) saturate(140%)"
      }}
    >
      <div className="w-full h-full flex flex-col rounded-lg p-4">
        <h3 className="text-lg font-medium text-gray-900 mb-2">Select DNC Type</h3>
        <div id="dnc-options-container" className="space-y-2">
          <button
            className="block w-full text-left px-4 py-2 text-sm text-gray-700 hover:bg-gray-100 rounded-md border"
            onClick={() => handleSelect('DNC - Phone')}
          >
            DNC - <span className="font-bold">Phone:</span> {phone}
          </button>
          <button
            className="block w-full text-left px-4 py-2 text-sm text-gray-700 hover:bg-gray-100 rounded-md border"
            onClick={() => handleSelect('DNC - Other Phone')}
          >
            DNC - <span className="font-bold">Other Phone:</span> {otherPhone}
          </button>
          <button
            className="block w-full text-left px-4 py-2 text-sm text-gray-700 hover:bg-gray-100 rounded-md border"
            onClick={() => handleSelect('DNC - Email')}
          >
            DNC - <span className="font-bold">Email:</span> {email}
          </button>
          <button
            className="block w-full text-left px-4 py-2 text-sm text-gray-700 hover:bg-gray-100 rounded-md border"
            onClick={() => handleSelect('DNC')}
          >
            DNC - <span className="font-bold">All:</span> All Contact Methods
          </button>
        </div>
      </div>
    </Modal>
  );
}

function LDAModal({ isOpen, onClose, onSelect }) {
  // Accepts JS Date object from Calendar, converts to MM/DD/YY
  const handleDateSelect = (date) => {
    if (date && onSelect) {
      const mm = String(date.getMonth() + 1).padStart(2, '0');
      const dd = String(date.getDate()).padStart(2, '0');
      const yy = String(date.getFullYear()).slice(-2);
      onSelect(`LDA ${mm}/${dd}/${yy}`);
    }
    if (onClose) onClose();
  };

  return (
    <Modal isOpen={isOpen} onClose={onClose} padding="12px">
      <div className="rounded-lg w-auto mx-auto">
        <h3 className="font-bold mb-2" id="month-year">
          Select LDA Date
        </h3>
        <Calendar onDateSelect={handleDateSelect} />
        {/* OK button removed */}
      </div>
    </Modal>
  );
}

export { DNCModal, LDAModal };

/*
Examples of what ldaKeywords now returns:

- ldaKeywords("Appointment on 10/7/2025")
  -> { dates: ["10/07/25"], keywords: ["10/7/2025"] }

- ldaKeywords("Call tomorrow")
  (Given current date 2025-11-03, "tomorrow" => 2025-11-04)
  -> { dates: ["11/04/25"], keywords: ["tomorrow"] }

- ldaKeywords("Let's meet Monday")
  (Given current date 2025-11-03 (Mon), next Monday => 2025-11-10)
  -> { dates: ["11/10/25"], keywords: ["Monday"] }

- ldaKeywords("weekend")
  (Next Saturday from 2025-11-03 => 2025-11-08)
  -> { dates: ["11/08/25"], keywords: ["weekend"] }

Notes:
- Returned object has two arrays: dates (normalized MM/DD/YY strings) and keywords (matched tokens).
- isOutreachTrigger now returns an added 'lda' property: { dates:[], keywords:[] } along with existing matched/tag fields.
*/
