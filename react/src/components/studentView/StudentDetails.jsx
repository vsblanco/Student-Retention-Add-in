// Timestamp: 2025-10-02 04:37 PM | Version: 3.0.0
import React, { useState } from 'react';
import {
  BadgeInfo,
  Phone,
  Mail,
  CalendarDays
} from 'lucide-react';
import { formatExcelDate } from '../utility/Conversion';

// A small reusable component for displaying a single detail item.
// This keeps our main component clean.
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

// Utility for copy-to-clipboard
const copyToClipboard = (text) => {
  if (navigator && navigator.clipboard) {
    navigator.clipboard.writeText(text);
  }
};

const style = `
.hover-bg:hover {
  background: #e5e7eb;
}
`;

function CopyField({ label, value, id, Icon }) {
  const [copied, setCopied] = useState(false);

  const handleCopy = () => {
    if (!value) return;
    copyToClipboard(value);
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
          gap: '0.5rem'
        }}
        className="hover-bg"
        onClick={value ? handleCopy : undefined}
      >
        {Icon && <Icon size={18} color="#6b7280" style={{ flexShrink: 0 }} />}
        <div style={{ flex: 1 }}>
          <label style={{ fontSize: '0.75rem', color: '#6b7280' }}>{label}</label>
          <p style={{ fontWeight: 600, color: '#1f2937', margin: 0 }}>{value || 'N/A'}</p>
        </div>
        <span
          className={`copy-feedback${copied ? '' : ' hidden'}`}
          style={{
            position: 'absolute',
            right: '0.5rem',
            top: '0.5rem',
            fontSize: '0.75rem',
            background: '#22c55e',
            color: 'white',
            padding: '0.25rem 0.5rem',
            borderRadius: '0.375rem',
            display: copied ? 'inline' : 'none',
            zIndex: 2
          }}
        >
          Copied!
        </span>
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
          Icon={BadgeInfo}
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
        <div
          style={{
            padding: '0.5rem 0.rem 0.5rem 0.25rem', // reduced left padding
            borderRadius: '0.5rem',
            display: 'flex',
            alignItems: 'center',
            gap: '0.5rem'
          }}
        >
          <CalendarDays size={18} color="#6b7280" style={{ flexShrink: 0 }} />
          <div>
            <label style={{ fontSize: '0.75rem', color: '#6b7280' }}>Last LDA</label>
            <p style={{ fontWeight: 600, color: '#1f2937', margin: 0 }}>
              {formatExcelDate(student["LDA"], "long")}
            </p>
          </div>
        </div>
      </div>
    </div>
  );
}

export default StudentDetails;

