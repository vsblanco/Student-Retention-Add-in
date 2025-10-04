import React from "react";
import AboutIcon from "../../assets/icons/about-icon.png";
import StudentViewIcon from "../../assets/icons/details-icon.png";

// Simple Excel-like Ribbon UI
const ribbonStyle = {
  display: "flex",
  flexDirection: "column",
  background: "#f3f3f3",
  borderBottom: "1px solid #d4d4d4",
  fontFamily: "Segoe UI, Arial, sans-serif",
  fontSize: "14px",
};

const tabsStyle = {
  display: "flex",
  alignItems: "center",
  height: "36px",
  padding: "0 8px",
  borderBottom: "1px solid #e0e0e0",
  background: "#f8f8f8",
};

const tabStyle = (active) => ({
  padding: "0 16px",
  height: "100%",
  display: "flex",
  alignItems: "center",
  cursor: "pointer",
  borderBottom: active ? "2px solid #217346" : "2px solid transparent",
  color: active ? "#217346" : "#222",
  fontWeight: active ? "bold" : "normal",
  background: active ? "#fff" : "transparent",
  transition: "background 0.2s, border-bottom 0.2s",
});


const buttonStyle = {
  border: "none", // removed border
  borderRadius: "3px",
  padding: "6px 10px",
  margin: "2px 0",
  cursor: "pointer",
  fontSize: "13px",
  minWidth: "40px",
  transition: "background 0.2s, border 0.2s",
  background: "#f8f8f8",
};

const buttonHoverStyle = {
  background: "#e6f2ec",
};

export default function Ribbon({ activeTab = "retention" }) {
  const [hovered, setHovered] = React.useState(null);

  return (
    <div style={ribbonStyle}>
      <div style={tabsStyle}>
        {tabs.map((tab) => (
          <div key={tab.key} style={tabStyle(tab.key === activeTab)}>
            {tab.label}
          </div>
        ))}
      </div>
      {activeTab === "retention" && (
        <div style={{ display: "flex", alignItems: "center", padding: "8px 12px", background: "#fff", minHeight: "48px", gap: "12px" }}>
          <button
            style={hovered === "about"
              ? { ...buttonStyle, ...buttonHoverStyle, display: "flex", alignItems: "center" }
              : { ...buttonStyle, display: "flex", alignItems: "center" }
            }
            onMouseEnter={() => setHovered("about")}
            onMouseLeave={() => setHovered(null)}
          >
            <img src={AboutIcon} alt="" style={{ width: 18, height: 18, marginRight: 6 }} />
            <span>About</span>
          </button>
          <button
            style={hovered === "student"
              ? { ...buttonStyle, ...buttonHoverStyle, display: "flex", alignItems: "center" }
              : { ...buttonStyle, display: "flex", alignItems: "center" }
            }
            onMouseEnter={() => setHovered("student")}
            onMouseLeave={() => setHovered(null)}
          >
            <img src={StudentViewIcon} alt="" style={{ width: 18, height: 18, marginRight: 6 }} />
            <span>Student View</span>
          </button>
        </div>
      )}
    </div>
  );
}

const tabs = [
  { key: "home", label: "Home" },
  { key: "insert", label: "Insert" },
  { key: "draw", label: "Draw" },
  { key: "data", label: "Data" },
  { key: "review", label: "Review" },
  { key: "view", label: "View" },
  { key: "retention", label: "Retention" }, // Added Retention tab
];

const groups = [
  {
    label: "Clipboard",
    buttons: ["Paste", "Cut", "Copy", "Format Painter"],
  },
  {
    label: "Font",
    buttons: ["Bold", "Italic", "Underline"],
  },
  {
    label: "Alignment",
    buttons: ["Left", "Center", "Right"],
  },
  {
    label: "Number",
    buttons: ["%", "$", "123"],
  },
];

const groupBarStyle = {
  display: "flex",
  alignItems: "flex-end",
  padding: "8px 8px 0 8px",
  background: "#fff",
};

const groupStyle = {
  display: "flex",
  flexDirection: "column",
  alignItems: "center",
  marginRight: "24px",
};

const groupLabelStyle = {
  fontSize: "11px",
  color: "#888",
  marginTop: "4px",
};
