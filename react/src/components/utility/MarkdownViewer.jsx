import React, { useEffect, useState } from "react";
import PropTypes from "prop-types";
import { marked } from "marked";

const MarkdownViewer = ({ file }) => {
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

  if (error) return <div>Error: {error}</div>;
  return (
    <div
      className="markdown-body"
      dangerouslySetInnerHTML={{ __html: content }}
    />
  );
};

MarkdownViewer.propTypes = {
  file: PropTypes.string.isRequired,
};

export default MarkdownViewer;
