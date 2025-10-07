import React, { useRef, useEffect, useState } from 'react';
import PropTypes from 'prop-types';

export const ContextMenu = ({
  position = { x: 0, y: 0 },
  visible = false,
  onClose = () => {},
  children,
}) => {
  const menuRef = useRef(null);
  const [fadeIn, setFadeIn] = useState(false);

  useEffect(() => {
    if (visible) {
      setFadeIn(false);
      // Trigger fade-in after mount
      setTimeout(() => setFadeIn(true), 10);
    } else {
      setFadeIn(false);
    }
  }, [visible]);

  if (!visible) return null;

  return (
    <ul
      ref={menuRef}
      style={{
        position: 'absolute',
        top: position.y,
        left: position.x,
        background: 'rgba(255, 255, 255, 0.25)',
        border: '1px solid rgba(200, 200, 255, 0.35)',
        boxShadow: '0 4px 24px 0 rgba(30, 60, 120, 0.18)',
        listStyle: 'none',
        padding: '8px 0',
        margin: 0,
        zIndex: 1000,
        minWidth: '160px',
        borderRadius: '16px',
        backdropFilter: 'blur(12px)',
        WebkitBackdropFilter: 'blur(12px)',
        overflow: 'hidden',
        opacity: fadeIn ? 1 : 0,
        transition: 'opacity 180ms cubic-bezier(.4,0,.2,1)',
      }}
      className="context-menu-fade"
    >
      {/* Place any buttons or elements as children */}
      {children}
    </ul>
  );
};

ContextMenu.propTypes = {
  position: PropTypes.shape({
    x: PropTypes.number,
    y: PropTypes.number,
  }),
  visible: PropTypes.bool,
  onClose: PropTypes.func,
  children: PropTypes.node,
};

ContextMenu.propTypes = {
  items: PropTypes.arrayOf(
    PropTypes.shape({
      label: PropTypes.string.isRequired,
      disabled: PropTypes.bool,
    })
  ),
  position: PropTypes.shape({
    x: PropTypes.number,
    y: PropTypes.number,
  }),
  visible: PropTypes.bool,
  onClose: PropTypes.func,
  onSelect: PropTypes.func,
};

export default ContextMenu;
