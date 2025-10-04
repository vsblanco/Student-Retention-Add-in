import React from "react";
import MarkdownViewer from "../utility/MarkdownViewer";

const About = () => {
  console.log("About component rendered");

  return <MarkdownViewer file="/about.md" theme="light" />;
};

export default About;
