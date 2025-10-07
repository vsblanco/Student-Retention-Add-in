import React, { useEffect, useRef } from 'react';

const ContextMenu = ({
  items = [],
  position = { x: 0, y: 0 },
  visible = false,
  onClose = () => {},
  onSelect = () => {},
}) => {
  const menuRef = useRef(null);

  useEffect(() => {
    const handleClickOutside = (event) => {
      if (menuRef.current && !menuRef.current.contains(event.target)) {
        onClose();
      }
    };
    if (visible) {
      document.addEventListener('mousedown', handleClickOutside);
    }
    return () => {
      document.removeEventListener('mousedown', handleClickOutside);
    };
  }, [visible, onClose]);

  if (!visible) return null;

  return (
    <ul
      ref={menuRef}
      style={{
        position: 'absolute',
        top: position.y,
        left: position.x,
        background: '#fff',
        border: '1px solid #ccc',
        boxShadow: '0 2px 8px rgba(0,0,0,0.15)',
        listStyle: 'none',
        padding: '8px 0',
        margin: 0,
        zIndex: 1000,
        minWidth: '160px',
      }}
    >
      {items.map((item, idx) => (
        <li
          key={idx}
          style={{
            padding: '8px 16px',
            cursor: 'pointer',
            userSelect: 'none',
            ...(item.disabled ? { color: '#aaa', cursor: 'not-allowed' } : {}),
          }}
          onClick={() => {
            if (!item.disabled) {
              onSelect(item);
              onClose();
            }
          }}
        >
          {item.label}
        </li>
      ))}
    </ul>
  );
};

export default ContextMenu;
