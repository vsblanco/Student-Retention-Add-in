import React from 'react';

// tagAnimDelays stays internal, but tagOptions is now passed as a prop
const tagAnimDelays = ['0ms', '60ms', '120ms', '180ms', '240ms'];

// TagDropdown component
function TagDropdown({ show, onTagClick, id, style, className, tags }) {
  return (
    <div
      id={id}
      key={`tag-dropdown-key-${show ? 'open' : 'closed'}`}
      className={`glass-dropdown absolute left-full top-0 -translate-y-3 ml-2 rounded-lg z-20 p-2 transition-all duration-300 flex flex-row gap-0.5 ${
        show
          ? 'opacity-100 -translate-y-3 pointer-events-auto'
          : 'opacity-0 -translate-y-2 pointer-events-none'
      } ${className || ''}${id === 'filter-tag-dropdown' ? ' bg-transparent' : ' bg-gray-200/90'}`}
      style={{
        boxShadow: '0 8px 24px 0 rgb(0 0 0 / 0.08)',
        minWidth: '100%',
        width: 'max-content',
        maxWidth: '220px',
        ...(id === 'filter-tag-dropdown' ? { background: 'transparent' } : {}),
        ...style
      }}
    >
      {(tags || []).map((tag, i) => (
        <a
          key={tag.label}
          href="#"
          className={`px-1 py-2 text-sm text-gray-700 rounded-md transition-all duration-200 group tag-anim`}
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
  tags // <-- new prop
}) {
  return (
    <div className="relative">
      <button
        id={`${dropdownId}-button`}
        type="button"
        className="px-2 py-1 text-xs font-semibold rounded-full bg-gray-200 text-gray-700 hover:bg-gray-300 transition-colors"
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
