import React, { useRef, useEffect, useState } from 'react';
import { createPortal } from 'react-dom';

// =======================
// TagPill Component
// =======================
export const TagPill = ({ label, spanClass }) => (
  <span className={`relative px-2 py-1 rounded-full ${spanClass}`}>
    <span className="relative z-10">{label}</span>
  </span>
);

// =======================
// TagButton Component
// =======================
export const TagButton = ({ tag, onClick }) => (
  <a
    href="#"
    className={`px-3 py-2 text-sm text-gray-700 rounded-md flex items-center hover:brightness-95 ${tag.spanClass}`}
    title={tag.title}
    onClick={e => {
      e.preventDefault();
      onClick(tag.label);
    }}
  >
    <TagPill label={tag.label} spanClass={tag.spanClass} />
  </a>
);

// =======================
// TagDropdownModal Component
// =======================
export const TagDropdownModal = ({
  show,
  onTagClick,
  id,
  tags,
  anchorRef,
  onClose
}) => {
  const [pos, setPos] = useState({ top: 0, left: 0, width: 220 });
  const dropdownRef = useRef(null);
  const [visible, setVisible] = useState(false);
  const [fade, setFade] = useState(false);

  useEffect(() => {
    if (show) {
      setVisible(true);
      // Trigger fade-in after mount
      setTimeout(() => setFade(true), 10);
    } else {
      setFade(false);
      // Delay unmount for fade-out
      const timeout = setTimeout(() => setVisible(false), 180);
      return () => clearTimeout(timeout);
    }
  }, [show]);

  useEffect(() => {
    if (show && anchorRef.current) {
      const r = anchorRef.current.getBoundingClientRect();
      setPos({
        top: r.bottom + window.scrollY + 4,
        left: r.left + window.scrollX,
        width: r.width || 220,
      });
    }
  }, [show, anchorRef]);

  // Click-away handler
  useEffect(() => {
    if (!show) return;
    function handleClick(e) {
      if (
        dropdownRef.current &&
        !dropdownRef.current.contains(e.target) &&
        anchorRef.current &&
        !anchorRef.current.contains(e.target)
      ) {
        onClose?.();
      }
    }
    document.addEventListener('mousedown', handleClick);
    return () => document.removeEventListener('mousedown', handleClick);
  }, [show, onClose, anchorRef]);

  if (!visible) return null;

  return createPortal(
    <div
      ref={dropdownRef}
      id={id}
      className={`glass-dropdown rounded-lg flex flex-col gap-1 items-stretch bg-gray-200/90 transition-opacity duration-180 ${fade ? 'opacity-100' : 'opacity-0 pointer-events-none'}`}
      style={{
        position: 'absolute',
        minWidth: '120px',
        maxHeight: '260px',
        boxShadow: '0 8px 24px 0 rgb(0 0 0 / 0.08)',
        overflowY: 'auto',
        borderRadius: '1rem',
        padding: '0.5rem 0',
        zIndex: 9999,
        ...pos,
      }}
    >
      <style>
        {`
          .glass-dropdown {
            scrollbar-width: none;
          }
          .glass-dropdown:hover {
            scrollbar-width: thin;
            scrollbar-color: rgba(0,0,0,0.1) transparent;
          }
          .glass-dropdown::-webkit-scrollbar {
            width: 0px;
            background: transparent;
          }
          .glass-dropdown:hover::-webkit-scrollbar {
            width: 6px;
            background: transparent;
          }
          .glass-dropdown::-webkit-scrollbar-thumb {
            background: rgba(0,0,0,0.1);
            border-radius: 8px;
          }
        `}
      </style>
      {tags.map(tag => (
        <TagButton key={tag.label} tag={tag} onClick={onTagClick} />
      ))}
    </div>,
    document.body
  );
};

// =======================
// Main InsertTagButton
// =======================
const InsertTagButton = ({
  dropdownId,
  onTagClick,
  showDropdown,
  setShowDropdown,
  dropdownTarget,
  setDropdownTarget,
  targetName,
  tags
}) => {
  const buttonRef = useRef(null);

  return (
    <div className="relative w-full">
      <button
        ref={buttonRef}
        id={`${dropdownId}-button`}
        type="button"
        className="px-2 py-1 text-xs font-semibold rounded-full bg-gray-200 text-gray-700 hover:bg-gray-300"
        onClick={() => {
          setShowDropdown(v => !v);
          setDropdownTarget(targetName);
        }}
        aria-expanded={showDropdown && dropdownTarget === targetName}
        aria-controls={dropdownId}
      >
        Insert Tag
      </button>
      <TagDropdownModal
        show={showDropdown && dropdownTarget === targetName}
        onTagClick={onTagClick}
        id={dropdownId}
        tags={tags}
        anchorRef={buttonRef}
        onClose={() => setShowDropdown(false)}
      />
    </div>
  );
};

export default InsertTagButton;