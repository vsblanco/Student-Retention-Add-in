import React, { useRef, useEffect, useState } from 'react';

// tagAnimDelays stays internal, but tagOptions is now passed as a prop
const tagAnimDelays = ['0ms', '60ms', '120ms', '180ms', '240ms'];

// TagDropdown component
function TagDropdown({ show, onTagClick, id, style, className, tags }) {
  const pillsContainerRef = useRef(null);
  const [dropdownWidth, setDropdownWidth] = useState(undefined);

  useEffect(() => {
    if (pillsContainerRef.current) {
      setDropdownWidth(pillsContainerRef.current.offsetWidth);
    }
  }, [tags, show]);

  return (
    <>
      <style>
        {`
          .glass-dropdown {
            border-radius: 1.5rem !important;
            overflow-x: auto;
            max-width: 80%; /* Remove limit, or set to a specific value like 90vw or 400px */
            scrollbar-width: none;
          }
          .glass-dropdown::-webkit-scrollbar {
            height: 6px;
            background: transparent;
          }
          .glass-dropdown:hover {
            scrollbar-width: thin;
          }
          .glass-dropdown:hover::-webkit-scrollbar {
            height: 6px;
            background: #e5e7eb;
          }
          .glass-dropdown::-webkit-scrollbar-thumb {
            background: #cbd5e1;
            border-radius: 4px;
          }
          .glass-dropdown:not(:hover)::-webkit-scrollbar-thumb {
            background: transparent;
          }
          .tag-pills-container {
            flex: 1 1 0%;
            height: 100%;
          }
        `}
      </style>
      <div
        id={id}
        key={`tag-dropdown-key-${show ? 'open' : 'closed'}`}
        className={`glass-dropdown rounded-lg z-20 p-2 transition-all duration-300 flex flex-row gap-0.5 items-center justify-center ${
          show
            ? 'opacity-100 pointer-events-auto'
            : 'opacity-0 pointer-events-none'
        } ${className || ''}${id === 'filter-tag-dropdown' ? ' bg-transparent' : ' bg-gray-200/90'}`}
        style={{
          position: 'absolute',
          left: '60px',
          top: '-5px', // Move up by 20px
          width: '500px',
          minWidth: '180px',
          height: '50px',
          boxShadow: '0 8px 24px 0 rgb(0 0 0 / 0.08)',
          ...(id === 'filter-tag-dropdown' ? { background: 'transparent' } : {}),
          ...style
        }}
        ref={pillsContainerRef}
      >
        {(tags || []).map((tag, i) => (
          <a
            key={tag.label}
            href="#"
            className={`px-1 py-0 text-sm text-gray-700 rounded-md transition-all duration-200 group tag-anim h-full`}
            title={tag.title}
            style={{
              animationDelay: show ? tagAnimDelays[i % tagAnimDelays.length] : '0ms'
            }}
            onClick={e => {
              e.preventDefault();
              onTagClick(tag.label);
            }}
          >
            <span className={tag.spanClass}>
              {tag.label}
            </span>
          </a>
        ))}
      </div>
    </>
  );
}

// InsertTagButton component
function InsertTagButton({ 
  dropdownId, 
  onTagClick, 
  showDropdown, 
  setShowDropdown, 
  dropdownTarget, 
  setDropdownTarget, 
  targetName, 
  dropdownClassName, 
  dropdownStyle,
  tags
}) {
  const containerRef = useRef(null);

  return (
    <div className="relative w-full" ref={containerRef}>
      <button
        id={`${dropdownId}-button`}
        type="button"
        className="px-2 py-1 text-xs font-semibold rounded-full bg-gray-200 text-gray-2700 hover:bg-gray-300 transition-colors"
        onClick={() => {
          setShowDropdown(v => !v);
          setDropdownTarget(targetName);
        }}
        aria-expanded={showDropdown && dropdownTarget === targetName}
        aria-controls={dropdownId}
      >
        Insert Tag
      </button>
      <TagDropdown
        show={showDropdown && dropdownTarget === targetName}
        onTagClick={onTagClick}
        id={dropdownId}
        className={dropdownClassName}
        style={dropdownStyle}
        tags={tags}
      />
    </div>
  );
}

export default InsertTagButton;