import React from "react";

const Modal = ({
  isOpen,
  onClose,
  children,
  borderRadius = "20px",
  className = "",
  padding,
  style,
  overlayStyle,
  ...props
}) => {
  if (!isOpen) return null;

  const handleOverlayClick = (e) => {
    if (e.target === e.currentTarget && onClose) {
      onClose(e);
    }
  };

  return (
    <div
      className={`modal-overlay ${className}`}
      style={{
        position: "fixed",
        top: 0,
        left: 0,
        right: 0,
        bottom: 0,
        zIndex: 9999,
        background: "rgba(30,30,30,0.5)", // updated to dark gray, 50% opacity
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        width: "100vw",
        height: "100vh",
        backdropFilter: "blur(2px) saturate(140%)",
        WebkitBackdropFilter: "blur(2px) saturate(140%)",
        ...overlayStyle,
      }}
      onClick={handleOverlayClick}
      {...props}
    >
      <div
        className="modal-container"
        style={{
          minWidth: 280,
          maxWidth: 500,
          maxHeight: 500,
          borderRadius,
          padding,
          background: "rgba(255,255,255,0.7)",
          boxShadow: "0 8px 32px 0 rgba(31, 38, 135, 0.37)",
          backdropFilter: "blur(15px) saturate(140%)",
          WebkitBackdropFilter: "blur(15px) saturate(140%)",
          border: "1px solid rgba(255,255,255,0.18)",
          display: "flex",
          flexDirection: "column",
          justifyContent: "center",
          alignItems: "center",
          margin: "0 auto",
          overflow: "auto",
          position: "relative",
          ...style,
        }}
        onClick={(e) => e.stopPropagation()}
      >
        {children}
      </div>
    </div>
  );
};

export default Modal;