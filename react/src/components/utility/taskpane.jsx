import React, { useRef, useState } from "react";

const taskpaneStyle = {
  position: "fixed",
  top: 0,
  right: 0,
  height: "100vh",
  width: "360px",
  background: "#fff",
  boxShadow: "-2px 0 8px rgba(0,0,0,0.15)",
  zIndex: 1000,
  transition: "transform 0.3s cubic-bezier(.4,0,.2,1)",
  display: "flex",
  flexDirection: "column",
};

const headerStyle = {
  padding: "16px",
  borderBottom: "1px solid #eee",
  display: "flex",
  justifyContent: "flex-end",
  alignItems: "center",
  background: "#f3f3f3",
};

const closeBtnStyle = {
  background: "none",
  border: "none",
  fontSize: "1.5rem",
  cursor: "pointer",
  color: "#666",
};

export default function Taskpane({ open, onClose, header, children }) {
  const [width, setWidth] = useState(360);
  const resizing = useRef(false);
  const startX = useRef(0);
  const startWidth = useRef(360);

  // Mouse event handlers for resizing
  const onMouseDown = (e) => {
    resizing.current = true;
    startX.current = e.clientX;
    startWidth.current = width;
    document.addEventListener("mousemove", onMouseMove);
    document.addEventListener("mouseup", onMouseUp);
    e.preventDefault();
  };

  const onMouseMove = (e) => {
    if (!resizing.current) return;
    const delta = startX.current - e.clientX;
    let newWidth = startWidth.current + delta;
    if (newWidth < 240) newWidth = 240;
    if (newWidth > 700) newWidth = 700;
    setWidth(newWidth);
  };

  const onMouseUp = () => {
    resizing.current = false;
    document.removeEventListener("mousemove", onMouseMove);
    document.removeEventListener("mouseup", onMouseUp);
  };

  return (
    <div
      style={{
        ...taskpaneStyle,
        width: width + "px",
        transform: open ? "translateX(0)" : "translateX(100%)",
        pointerEvents: open ? "auto" : "none",
        userSelect: resizing.current ? "none" : "auto",
      }}
      aria-hidden={!open}
    >
      {/* Resize handle */}
      <div
        style={{
          position: "absolute",
          left: 0,
          top: 0,
          width: "6px",
          height: "100%",
          cursor: "ew-resize",
          zIndex: 1001,
        }}
        onMouseDown={onMouseDown}
        title="Resize"
      />
      {/* Combined header and close button */}
      <div
        style={{
          display: "flex",
          alignItems: "center",
          justifyContent: "space-between",
          padding: "6px 12px",
          borderBottom: "1px solid #eee",
          background: "#f3f3f3",
          minHeight: "40px",
        }}
      >
        <div>{header}</div>
        <button style={closeBtnStyle} onClick={onClose} aria-label="Close taskpane">
          &times;
        </button>
      </div>
      <div style={{ flex: 1, overflowY: "auto", padding: "16px", paddingTop: 0 }}>
        {children}
      </div>
    </div>
  );
}
