import React, { useState } from 'react';
import Modal from '../utility/Modal';
import Datepicker from 'react-tailwindcss-datepicker';
import Calendar from '../utility/Calendar';
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
  if (!text || typeof text !== 'string') return { matched: false, tag: null };

  const lower = text.toLowerCase();

  // Check Contacted keywords
  const contactedMatched = Contacted_Keywords.some(k => lower.includes(k.toLowerCase()));

  // Check LDA date-like keywords (returns array of M/D/YY strings)
  const ldaMatches = ldaKeywords(text || '');

  const tags = [];
  if (contactedMatched) tags.push('Contacted');
  if (Array.isArray(ldaMatches) && ldaMatches.length > 0) {
    // prefix each LDA date with "LDA "
    tags.push(...ldaMatches.map(d => `LDA ${d}`));
  }

  if (tags.length === 0) return { matched: false, tag: null };

  return { matched: true, tag: tags.join(', ') };
};

// New helper: extract LDA match strings (normalized as M/D/YY)
export function ldaKeywords(text) {
  if (!text || typeof text !== 'string') return [];
  const keywords = [
    "Tomorrow", "next week",
    "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday",
    "weekend", "weekends"
  ];
  const shortDatePart = '\\b\\d{1,2}[\\/\\-]\\d{1,2}(?:[\\/\\-]\\d{2,4})?\\b';
  const keywordPart = `\\b(?:${keywords.map(k => k.replace(/[.*+?^${}()|[\\]\\\\]/g, '\\$&')).join('|')})\\b`;
  const combinedRegex = new RegExp(`${shortDatePart}|${keywordPart}`, 'gi');

  const results = [];
  let match;
  const now = new Date();

  const formatTwoDigit = d => {
    const mm = d.getMonth() + 1;
    const dd = d.getDate();
    const yy = String(d.getFullYear()).slice(-2);
    return `${mm}/${dd}/${yy}`;
  };

  const weekdayIndex = name => {
    const map = { sunday:0, monday:1, tuesday:2, wednesday:3, thursday:4, friday:5, saturday:6 };
    return map[name.toLowerCase()];
  };

  while ((match = combinedRegex.exec(text)) !== null) {
    const token = match[0];
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

  // dedupe and return normalized strings
  return Array.from(new Set(results));
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
