import React from 'react';

// Tag definitions for comments
export const COMMENT_TAGS = [
  {
    label: "Urgent",
    bgClass: "bg-red-100",
    tagClass: "px-2 py-0.5 font-semibold rounded-full bg-red-100 text-red-800"
  },
  {
    label: "Note",
    bgClass: "bg-gray-200",
    tagClass: "px-2 py-0.5 font-semibold rounded-full bg-gray-200 text-gray-800"
  },
  {
    label: "DNC",
    bgClass: "bg-red-200",
    tagClass: "px-2 py-0.5 font-semibold rounded-full bg-red-200 text-black"
  },
  {
    label: "LDA",
    bgClass: "bg-orange-100",
    tagClass: "px-2 py-0.5 font-semibold rounded-full bg-orange-100 text-orange-800"
  },
  {
    label: "Contacted",
    bgClass: "bg-yellow-100",
    tagClass: "px-2 py-0.5 font-semibold rounded-full bg-yellow-200 text-yellow-800"
  },
  {
    label: "Outreach",
    bgClass: "bg-gray-200",
    tagClass: "px-2 py-0.5 font-semibold rounded-full bg-blue-200 text-gray-800"
  }
];

function Comment({ entry, searchTerm, index }) {
  // Find tag info or fallback
  const tagInfo = COMMENT_TAGS.find(t => t.label === entry.tag);

  // Determine background and tag classes
  const bgClass = tagInfo ? tagInfo.bgClass : "bg-gray-200";
  const tagClass = tagInfo ? tagInfo.tagClass : "px-2 py-0.5 font-semibold rounded-full bg-blue-100 text-blue-800";

  // Highlight search term in comment
  let commentContent = entry.comment;
  if (searchTerm && entry.comment) {
    const regex = new RegExp(`(${searchTerm.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')})`, 'gi');
    commentContent = entry.comment.split(regex).map((part, i) =>
      regex.test(part) ? (
        <mark key={i} className="bg-yellow-200 px-0.5 rounded">{part}</mark>
      ) : (
        part
      )
    );
  }

  return (
    <li
      className={`p-3 rounded-lg shadow-sm relative ${bgClass}`}
      data-row-index={entry.studentId || index}
    >
      <p className="text-sm text-gray-800">{commentContent}</p>
      <div className="text-xs text-gray-500 mt-2 pt-2 border-t border-gray-200 flex justify-between items-center">
        <div className="flex items-center gap-2">
          {entry.tag && (
            <span className={tagClass}>
              {entry.tag}
            </span>
          )}
          <span className="font-medium">{entry.createdBy}</span>
        </div>
        <span>
          {entry.timestamp}
        </span>
      </div>
    </li>
  );
}

export default Comment;
