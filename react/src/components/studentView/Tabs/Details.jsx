// 2025-12-06 13:40 EST - Version 3.1.0 - Added clipboard fallback support
import React, { useState } from 'react';
import {
  IdCardLanyard,
  Phone,
  Mail,
  CalendarDays
} from 'lucide-react';
import { formatExcelDate } from '../../utility/Conversion';

// A small reusable component for displaying a single detail item.
const DetailItem = ({ label, value }) => {
  const itemStyles = {
    marginBottom: '10px'
  };
  const labelStyles = {
    fontWeight: '600',
    color: '#34495e',
    display: 'block',
    fontSize: '14px'
  };
  const valueStyles = {
    color: '#7f8c8d',
    fontSize: '14px',
    marginLeft: '5px'
  };

  return (
    <div style={itemStyles}>
      <span style={labelStyles}>{label}:</span>
      <span style={valueStyles}>{value || 'N/A'}</span>
    </div>
  );
};

// Utility for copy-to-clipboard with Fallback mechanism
// This ensures functionality in environments where navigator.clipboard is restricted (e.g. Office Add-ins)
const copyToClipboard = async (text) => {
  if (!text) return;

  // 1. Try Modern Async API
  if (navigator && navigator.clipboard && navigator.clipboard.writeText) {
    try {
      await navigator.clipboard.writeText(text);
      // Success, exit function
      return;
    } catch (err) {
      console.warn('Clipboard API failed, attempting fallback...', err);
    }
  }

  // 2. Fallback: document.execCommand('copy')
  try {
    const ta = document.createElement('textarea');
    ta.value = text;
    
    // Ensure element is not visible but part of DOM
    ta.style.position = 'fixed';
    ta.style.left = '-9999px';
    ta.style.top = '0';
    ta.setAttribute('readonly', ''); // Prevent keyboard popping up on mobile
    
    document.body.appendChild(ta);
    ta.focus();
    ta.select();
    
    // Execute copy
    const successful = document.execCommand('copy');
    document.body.removeChild(ta);
    
    if (!successful) {
        console.error('Fallback copy failed.');
    }
  } catch (err) {
    console.error('All copy methods failed', err);
  }
};

const style = `
.hover-bg:hover {
  background: #e5e7eb;
}
`;

function CopyField({ label, value, id, Icon }) {
  const [copied, setCopied] = useState(false);

  const handleCopy = async () => {
    if (!value) return;
    
    // Use the robust copy utility
    await copyToClipboard(value);
    
    setCopied(true);
    setTimeout(() => setCopied(false), 1200);
  };

  return (
    <>
      {/* Inject the hover style once */}
      <style>{style}</style>
      <div
        id={id}
        style={{
          padding: '0.5rem',
          borderRadius: '0.5rem',
          cursor: value ? 'pointer' : 'default',
          position: 'relative',
          transition: 'background 0.15s',
          display: 'flex',
          alignItems: 'center',
          gap: '0.5rem',
          overflow: 'hidden'
        }}
        className="hover-bg"
        onClick={value ? handleCopy : undefined}
      >
        {/* Green overlay for copied feedback */}
        <div
          style={{
            position: 'absolute',
            inset: 0,
            background: 'rgba(34,197,94,0.25)', // green overlay
            opacity: copied ? 1 : 0,
            pointerEvents: 'none',
            transition: 'opacity 0.7s ease'
          }}
        />
        {Icon && <Icon size={18} color="#6b7280" style={{ flexShrink: 0 }} />}
        <div style={{ flex: 1 }}>
          <label style={{ fontSize: '0.75rem', color: '#6b7280' }}>{label}</label>
          <p style={{ fontWeight: 600, color: '#1f2937', margin: 0 }}>{value || 'N/A'}</p>
        </div>
      </div>
    </>
  );
}

function StudentDetails({ student }) {
  // --- STYLES ---
  const studentNameStyles = {
    fontSize: '22px',
    fontWeight: '600',
    color: '#2c3e50',
    marginBottom: '15px',
    borderBottom: '1px solid #ecf0f1',
    paddingBottom: '10px'
  };

  const sectionStyles = {
    marginBottom: '20px'
  };

  // Check for both name variations from the student object
  const studentName = student["StudentName"] || student["Student Name"];
  return (
    <div
      id="panel-details"
      style={{
        padding: '1rem 1rem 1rem 0.25rem', // reduced left padding
        display: 'flex',
        flexDirection: 'column',
        gap: '1rem'
      }}
    >
      <div style={{ display: 'flex', flexDirection: 'column', gap: '0.75rem' }}>
        <CopyField
          label="Student ID"
          value={student.ID}
          id="copy-student-id"
          Icon={IdCardLanyard}
        />
        <CopyField
          label="Primary Phone"
          value={student.Phone}
          id="copy-primary-phone"
          Icon={Phone}
        />
        <CopyField
          label="Other Phone"
          value={student.OtherPhone}
          id="copy-other-phone"
          Icon={Phone}
        />
        <CopyField
          label="Student Email"
          value={student.StudentEmail}
          id="copy-student-email"
          Icon={Mail}
        />
        <CopyField
          label="Personal Email"
          value={student.PersonalEmail}
          id="copy-personal-email"
          Icon={Mail}
        />
        <CopyField
          label="Last LDA"
          value={formatExcelDate(student["LDA"], "long")}
          id="copy-last-lda"
          Icon={CalendarDays}
        />
      </div>
    </div>
  );
}

export default StudentDetails;