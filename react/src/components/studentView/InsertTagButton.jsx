import React, { useRef, useEffect, useState } from 'react';

// Inject custom scrollbar styles globally
if (typeof document !== "undefined" && !document.getElementById("custom-scrollbar-style")) {
  const style = document.createElement("style");
  style.id = "custom-scrollbar-style";
  style.innerHTML = `
    .glass-dropdown {
      border-radius: 1.5rem !important;
      overflow-x: auto;
      max-width: 80%;
      scrollbar-width: none;
    }
    .glass-dropdown::-webkit-scrollbar {
      height: 10px;
      background: transparent;
    }
    .glass-dropdown:hover {
      scrollbar-width: thin;
      overflow-x: scroll;
    }
    .glass-dropdown:hover::-webkit-scrollbar {
      height: 10px;
      background: #f3f4f6;
    }
    .glass-dropdown::-webkit-scrollbar-thumb {
      background: linear-gradient(90deg, #cbd5e1 0%, #94a3b8 100%);
      border-radius: 10px;
      box-shadow: 0 2px 8px rgba(0,0,0,0.12);
      border: 2px solid #f3f4f6;
      transition: background 0.2s;
    }
    .glass-dropdown:hover::-webkit-scrollbar-thumb {
      background: linear-gradient(90deg, #94a3b8 0%, #64748b 100%);
    }
    .glass-dropdown::-webkit-scrollbar-track {
      background: #f3f4f6;
      border-radius: 10px;
    }
    .glass-dropdown:not(:hover)::-webkit-scrollbar-thumb {
      background: transparent;
    }
    .tag-pills-container {
      flex: 1 1 0%;
      height: 100%;
    }
  `;
  document.head.appendChild(style);
}

const tagAnimDelays = ['0ms', '60ms', '120ms', '180ms', '240ms'];

const dropdownBaseStyle = {
  position: 'absolute',
  left: '60px',
  top: '-5px',
  width: '500px',
  minWidth: '180px',
  height: '50px',
  boxShadow: '0 8px 24px 0 rgb(0 0 0 / 0.08)',
};

function TagDropdown({ show, onTagClick, id, style = {}, className = '', tags = [] }) {
  const pillsContainerRef = useRef(null);
  const [dropdownWidth, setDropdownWidth] = useState();

  useEffect(() => {
    if (pillsContainerRef.current) {
      setDropdownWidth(pillsContainerRef.current.offsetWidth);
    }
  }, [tags, show]);

  const combinedStyle = {
    ...dropdownBaseStyle,
    ...(id === 'filter-tag-dropdown' ? { background: 'transparent' } : {}),
    ...style
  };

  return (
    <>
      <div
        id={id}
        key={`tag-dropdown-key-${show ? 'open' : 'closed'}`}
        className={`glass-dropdown rounded-lg z-20 p-2 transition-all duration-300 flex flex-row gap-0.5 items-center justify-center ${
          show
            ? 'opacity-100 pointer-events-auto'
            : 'opacity-0 pointer-events-none'
        } ${className}${id === 'filter-tag-dropdown' ? ' bg-transparent' : ' bg-gray-200/90'}`}
        style={combinedStyle}
        ref={pillsContainerRef}
      >
        {tags.map((tag, i) => (
          <a
            key={tag.label}
            href="#"
            className="px-1 py-0 text-sm text-gray-700 rounded-md transition-all duration-200 group tag-anim h-full"
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

function InsertTagButton({
  dropdownId,
  onTagClick,
  showDropdown,
  setShowDropdown,
  dropdownTarget,
  setDropdownTarget,
  targetName,
  dropdownClassName = '',
  dropdownStyle = {},
  tags = []
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