// Timestamp: 2025-11-22 | Version: 4.0.0
import React, { useEffect, useState } from "react";
import PropTypes from "prop-types";
import ReactMarkdown from "react-markdown";
import remarkGfm from "remark-gfm";
import "github-markdown-css/github-markdown.css";

const MarkdownViewer = ({ file, onLinkClick, theme = "light" }) => {
  const [content, setContent] = useState("");
  const [error, setError] = useState(null);

  useEffect(() => {
    if (!file) return;
    fetch(file)
      .then((res) => {
        if (!res.ok) throw new Error(`Failed to load ${file}`);
        return res.text();
      })
      .then((text) => setContent(text))
      .catch((err) => setError(err.message));
  }, [file]);

  const themeStyles =
    theme === "dark"
      ? { background: "#181a1b", color: "#eee" }
      : { background: "#fff", color: "#111" };

  if (error) return <div className="p-4 text-red-500">Error: {error}</div>;

  return (
    <div className="markdown-body" style={themeStyles}>
      <ReactMarkdown
        remarkPlugins={[remarkGfm]}
        components={{
          a: ({ node, href, children, ...props }) => {
            const isInternal = href && href.endsWith(".md");
            return (
              <a
                href={href}
                onClick={(e) => {
                  if (isInternal && onLinkClick) {
                    e.preventDefault(); // Stop browser navigation
                    onLinkClick(href); // Tell parent to switch file
                  }
                }}
                {...props}
              >
                {children}
              </a>
            );
          },
        }}
      >
        {content}
      </ReactMarkdown>
    </div>
  );
};

MarkdownViewer.propTypes = {
  file: PropTypes.string.isRequired,
  onLinkClick: PropTypes.func, // Function to handle internal navigation
  theme: PropTypes.oneOf(["light", "dark"]),
};

export default MarkdownViewer;