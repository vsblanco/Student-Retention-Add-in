import React, { useEffect, useState } from "react";
import PropTypes from "prop-types";
import { marked } from "marked";
import "github-markdown-css/github-markdown.css"; // <-- Add this import

const MarkdownViewer = ({ file, theme = "light" }) => {
  const [content, setContent] = useState("");
  const [error, setError] = useState(null);

  useEffect(() => {
    if (!file) return;
    fetch(file)
      .then((res) => {
        if (!res.ok) throw new Error("Failed to load markdown file");
        return res.text();
      })
      .then((text) => setContent(marked.parse(text)))
      .catch((err) => setError(err.message));
  }, [file]);

  const themeStyles =
    theme === "dark"
      ? { background: "#181a1b", color: "#eee" }
      : { background: "#fff", color: "#111" };

  if (error) return <div>Error: {error}</div>;
  return (
    <div
      className="markdown-body"
      style={themeStyles}
      dangerouslySetInnerHTML={{ __html: content }}
    />
  );
};

MarkdownViewer.propTypes = {
  file: PropTypes.string.isRequired,
  theme: PropTypes.oneOf(["light", "dark"]),
};

export default MarkdownViewer;
