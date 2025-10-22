import React from 'react';

function TabButton({ tabId, label, isActive, onTabChange }) {
  return (
    <button
      type="button"
      className={`studentview-tab ${isActive ? 'active' : ''}`}
      onClick={() => onTabChange(tabId)}
    >
      {label}
    </button>
  );
}

export default React.memo(TabButton);