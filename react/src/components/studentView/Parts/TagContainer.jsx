import React, { useEffect } from 'react';

const TagContainer = ({ tags = [], tag = null, className = '', onRemove, onChange }) => {
  // normalize input:
  // - if `tags` is a string, split by commas
  // - if `tags` is an array, use as-is
  // - if `tag` provided, treat as single-item
  let list = [];
  if (typeof tags === 'string') {
    list = tags
      .split(',')
      .map(s => s.trim())
      .filter(Boolean)
      .map(label => ({ label }));
  } else if (Array.isArray(tags)) {
    list = tags;
  } else if (tag) {
    list = [tag];
  }

  // derive normalized label array and comma-separated string
  const normalizedLabels = (list || [])
    .map((t) => (typeof t === 'string' ? t : (t && t.label) || ''))
    .filter(Boolean);
  const tagString = normalizedLabels.join(', ');

  // notify parent of current tags as a comma-separated string
  useEffect(() => {
    if (typeof onChange === 'function') {
      onChange(tagString);
    }
  }, [tagString, onChange]);

  if (!list || list.length === 0) return null;

  // container style copied from CommentModal modalTagViewPills/modalTagPills
  const containerStyle = {
    display: 'flex',
    flexWrap: 'wrap',
    gap: 4,
    padding: '2px 10px',
    borderRadius: 9999,
    border: '1.5px solid #cfcfcf',
    marginTop: 0,
    minHeight: 32,
    alignItems: 'center',
    width: '100%',
    boxSizing: 'border-box',
    maxWidth: '100%',
    background: 'transparent',
  };

  // individual pill base classes and inline style similar to CommentModal pills
  const pillBaseClass = 'px-1 py-0.5 text-xs font-semibold rounded-full';
  const pillInlineStyle = {
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap',
    fontSize: '0.8em',
    padding: '2px 6px',
    display: 'inline-flex',
    alignItems: 'center',
    gap: 0,
    borderRadius: 9999,
    marginRight: 0,
  };

  const removeBtnStyle = {
    margin: 0,
    background: 'transparent',
    border: 'none',
    cursor: 'pointer',
    fontSize: '0.85em',
    lineHeight: 1,
    padding: '0 4px',
    color: '#374151',
    display: 'inline-flex',
    alignItems: 'center',
    justifyContent: 'center',
  };

  return (
    <div className={`${className}`} style={containerStyle}>
      {list.map((t, i) => {
        const label = typeof t === 'string' ? t : t.label || String(i);
        const spanClass = typeof t === 'object' ? (t.spanClass || '') : '';
        return (
          <span
            key={label + i}
            className={`${pillBaseClass} ${spanClass}`}
            style={pillInlineStyle}
            title={label}
          >
            <span style={{ overflow: 'hidden', textOverflow: 'ellipsis' }}>{label}</span>
            {typeof onRemove === 'function' && (
              <button
                type="button"
                aria-label={`Remove ${label}`}
                onClick={(e) => {
                  e.stopPropagation();
                  onRemove(label);
                }}
                style={removeBtnStyle}
              >
                ×
              </button>
            )}
          </span>
        );
      })}
    </div>
  );
};

export default TagContainer;
