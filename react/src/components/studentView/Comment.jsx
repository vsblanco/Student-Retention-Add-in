import React, { useState } from 'react';
import { formatExcelDate } from '../utility/Conversion';

// Tag definitions for comments
export const COMMENT_TAGS = [
  {
    label: "Urgent",
    bgClass: "bg-red-100",
    tagClass: "px-2 py-0.5 font-semibold rounded-full bg-red-100 text-red-800",
    priority: 2
  },
  {
    label: "Note",
    bgClass: "bg-gray-200",
    tagClass: "px-2 py-0.5 font-semibold rounded-full bg-gray-700 text-gray-200",
    pinned: true,
    priority: 3
  },
  {
    label: "DNC",
    bgClass: "bg-red-200",
    tagClass: "px-2 py-0.5 font-semibold rounded-full bg-red-200 text-black",
    pinned: true,
    priority: 4,
    subtags: [
      {
        label: "DNC - Phone",
        bgClass: "bg-red-200",
        tagClass: "px-2 py-0.5 font-semibold rounded-full bg-red-200 text-black",
        pinned: true,
        priority: 4
      },
      {
        label: "DNC - Other Phone",
        bgClass: "bg-red-200",
        tagClass: "px-2 py-0.5 font-semibold rounded-full bg-red-200 text-black",
        pinned: true,
        priority: 4
      },
      {
        label: "DNC - Email",
        bgClass: "bg-red-200",
        tagClass: "px-2 py-0.5 font-semibold rounded-full bg-red-200 text-black",
        pinned: true,
        priority: 4
      }
    ]
  },
  {
    label: "LDA",
    bgClass: "bg-orange-100",
    tagClass: "px-2 py-0.5 font-semibold rounded-full bg-orange-100 text-orange-800",
    priority: 3
  },
  {
    label: "Contacted",
    bgClass: "bg-yellow-100",
    tagClass: "px-2 py-0.5 font-semibold rounded-full bg-yellow-200 text-yellow-800",
    priority: 2
  },
  {
    label: "Outreach",
    bgClass: "bg-gray-100",
    tagClass: "px-2 py-0.5 font-semibold rounded-full bg-blue-100 text-blue-800",
    priority: 1
  },
  {
    label: "Quote",
    bgClass: "bg-blue-100",
    tagClass: "px-2 py-0.5 font-semibold rounded-full bg-blue-50 text-blue-800",
    priority: 2
  }
];

// Helper to find tag or subtag by label
function findTagInfo(label) {
  for (const tag of COMMENT_TAGS) {
    if (tag.label === label) return tag;
    if (tag.subtags) {
      const subtag = tag.subtags.find(st => st.label === label);
      if (subtag) return subtag;
    }
  }
  // Special handling for LDA subclass: "LDA" + date (e.g., "LDA 10/3/25")
  if (typeof label === "string" && label.startsWith("LDA ")) {
    // Return LDA tag but override label, bgClass, and priority for display
    const ldaTag = COMMENT_TAGS.find(t => t.label === "LDA");
    if (ldaTag) {
      return {
        ...ldaTag,
        label, // keep the full label, e.g., "LDA 10/3/25"
        bgClass: "bg-orange-100",
        priority: 3
      };
    }
  }
  return null;
}

function Comment({ entry, searchTerm, index }) {
  // Support multiple tags separated by commas
  let tags = entry.tag
    ? entry.tag.split(',').map(t => t.trim()).filter(Boolean)
    : [];
  // Remove "Comment" tags
  tags = tags.filter(t => t !== "Comment");

  // Find tag info for all tags
  const tagInfos = tags.map(findTagInfo);

  // Determine which tag has the *highest* priority value (largest number)
  // If multiple tags have the same priority, prefer the first in the list
  let tagInfo = null;
  let maxPriority = -Infinity;
  tagInfos.forEach((info, idx) => {
    if (info && typeof info.priority === "number") {
      if (info.priority > maxPriority) {
        maxPriority = info.priority;
        tagInfo = info;
      }
    }
  });
  if (!tagInfo) tagInfo = tagInfos[0] || null;

  // Determine background class from the highest priority tag
  // If no tagInfo, use default
  const bgClass = tagInfo && tagInfo.bgClass ? tagInfo.bgClass : "bg-gray-200";
  const tagClass = tagInfo ? tagInfo.tagClass : "px-2 py-0.5 font-semibold rounded-full bg-blue-100 text-blue-800";

  // Check if Quote tag is present
  const hasQuoteTag = tags.includes("Quote");

  // Special handling for LDA tag with date in label
  let isLdaWithDate = false;
  let ldaDate = null;
  let ldaDateRegex = null;
  if (tagInfo && tagInfo.label && /^LDA\s+(.+)/i.test(tagInfo.label)) {
    isLdaWithDate = true;
    ldaDate = tagInfo.label.replace(/^LDA\s+/i, '').trim();

    // Try to parse numeric date (e.g., 10/7/25) to possible month/day/year
    let datePatterns = [];
    const numericDateMatch = ldaDate.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
    if (numericDateMatch) {
      // Numeric date: MM/DD/YY(YY)
      const monthNum = parseInt(numericDateMatch[1], 10);
      const dayNum = parseInt(numericDateMatch[2], 10);
      // Accept both 2-digit and 4-digit years
      const yearNum = numericDateMatch[3];

      // Build patterns for "October 7", "Oct 7", "October 7th", etc.
      const monthNames = [
        "", "January", "February", "March", "April", "May", "June",
        "July", "August", "September", "October", "November", "December"
      ];
      const monthFull = monthNames[monthNum] || "";
      const monthAbbr = monthFull ? monthFull.slice(0, 3) : "";

      // Patterns: October 7, October 7th, Oct 7, Oct 7th (case-insensitive)
      if (monthFull) {
        datePatterns.push(`${monthFull}\\s+${dayNum}(?:st|nd|rd|th)?`);
        datePatterns.push(`${monthAbbr}\\.?\\s+${dayNum}(?:st|nd|rd|th)?`);
      }
      // Also match numeric date as fallback
      datePatterns.push(`${monthNum}\\/${dayNum}\\/${yearNum}`);
      // Allow for 2-digit year in text (e.g., 25 or 2025)
      if (yearNum.length === 2) {
        datePatterns.push(`${monthNum}\\/${dayNum}\\/20${yearNum}`);
      }
    } else {
      // Fallback: match the string as-is
      datePatterns.push(ldaDate.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&'));
    }
    ldaDateRegex = new RegExp(datePatterns.join('|'), 'gi');
  }

  // Parse comment for quote logic
  let beforeQuote = null, quoteText = null, afterQuote = null;
  if (hasQuoteTag && typeof entry.comment === "string") {
    // Match text between the first pair of straight or curly quotes
    const match = entry.comment.match(/^(.*?)["“”](.*?)["“”](.*)$/s);
    if (match) {
      beforeQuote = match[1]?.trim() ? <p className="text-sm text-gray-800 mb-2">{match[1].trim()}</p> : null;
      quoteText = match[2]?.trim();
      afterQuote = match[3]?.trim() ? <p className="text-sm text-gray-800 mt-2">{match[3].trim()}</p> : null;
    } else {
      quoteText = entry.comment;
    }
  }

  // Highlight search term in comment (for non-quote or fallback)
  let commentContent = entry.comment;
  if (!hasQuoteTag || !quoteText) {
    if (isLdaWithDate && entry.comment && ldaDateRegex) {
      // Bold the date in the comment if found (any format)
      let lastIndex = 0;
      let parts = [];
      let match;
      ldaDateRegex.lastIndex = 0;
      while ((match = ldaDateRegex.exec(entry.comment)) !== null) {
        if (match.index > lastIndex) {
          parts.push(entry.comment.slice(lastIndex, match.index));
        }
        parts.push(<b key={`lda-date-${match.index}`}>{match[0]}</b>);
        lastIndex = ldaDateRegex.lastIndex;
      }
      if (lastIndex < entry.comment.length) {
        parts.push(entry.comment.slice(lastIndex));
      }
      // Highlight "Tomorrow", "next week", weekdays, and weekends in the resulting parts
      const highlightLdaKeywords = part => {
        if (typeof part !== "string") return part;
        // Regex for "Tomorrow", "next week", weekdays, "weekend", "weekends" (case-insensitive)
        const keywordRegex = /\b(Tomorrow|next week|Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday|weekend|weekends)\b/gi;
        let keywordParts = [];
        let lastIdx = 0;
        let kwMatch;
        while ((kwMatch = keywordRegex.exec(part)) !== null) {
          if (kwMatch.index > lastIdx) {
            keywordParts.push(part.slice(lastIdx, kwMatch.index));
          }
          keywordParts.push(
            <b key={`lda-keyword-${kwMatch.index}`}>{kwMatch[0]}</b>
          );
          lastIdx = keywordRegex.lastIndex;
        }
        if (lastIdx < part.length) {
          keywordParts.push(part.slice(lastIdx));
        }
        return keywordParts.length > 0 ? keywordParts : part;
      };
      commentContent = parts.flatMap(highlightLdaKeywords);
    } else if (searchTerm && entry.comment) {
      const regex = new RegExp(`(${searchTerm.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')})`, 'gi');
      commentContent = entry.comment.split(regex).map((part, i) =>
        regex.test(part) ? (
          <mark key={i} className="bg-yellow-200 px-0.5 rounded">{part}</mark>
        ) : (
          part
        )
      );
    }
  }

  // Format timestamp if it's a number/string Excel serial
  let formattedTimestamp = entry.timestamp;
  if (!isNaN(entry.timestamp)) {
    formattedTimestamp = formatExcelDate(entry.timestamp);
  }

  // State for expanding/collapsing long comments and quotes
  const [expanded, setExpanded] = useState(false);

  // Ref and state for regular comment
  const commentRef = React.useRef(null);
  const [isLong, setIsLong] = useState(false);

  // Ref and state for quote text
  const quoteRef = React.useRef(null);
  const [isQuoteLong, setIsQuoteLong] = useState(false);

  React.useEffect(() => {
    if (commentRef.current) {
      const el = commentRef.current;
      const style = window.getComputedStyle(el);
      const lineHeight = parseFloat(style.lineHeight);
      const lines = el.scrollHeight / lineHeight;
      setIsLong(lines > 3.1);
    }
    if (quoteRef.current) {
      const el = quoteRef.current;
      const style = window.getComputedStyle(el);
      const lineHeight = parseFloat(style.lineHeight);
      const lines = el.scrollHeight / lineHeight;
      setIsQuoteLong(lines > 3.1);
    }
  }, [entry.comment, commentContent, quoteText, expanded]);

  return (
    <li
      className={`p-3 rounded-lg shadow-sm relative ${bgClass}`}
      data-row-index={entry.studentId || index}
    >
      {hasQuoteTag && quoteText ? (
        <>
          {beforeQuote}
          <blockquote className="relative bg-blue-50 border-l-4 border-blue-400 pl-6 pr-2 py-3 mb-2 rounded">
            <span className="absolute left-2 top-2 text-4xl text-blue-200 leading-none select-none" aria-hidden="true">“</span>
            <span
              ref={quoteRef}
              className={`text-base text-blue-900 font-serif ${!expanded ? 'line-clamp-3' : ''}`}
              style={
                !expanded
                  ? {
                      display: '-webkit-box',
                      WebkitLineClamp: 3,
                      WebkitBoxOrient: 'vertical',
                      overflow: 'hidden'
                    }
                  : {}
              }
            >
              {quoteText}
            </span>
            <span className="absolute right-2 bottom-2 text-4xl text-blue-200 leading-none select-none" aria-hidden="true">”</span>
          </blockquote>
          {isQuoteLong && (
            <button
              className="text-xs text-gray-600 mt-1 rounded bg-gray-100 bg-opacity-0 hover:bg-opacity-100 transition duration-150 px-2 py-1"
              onClick={() => setExpanded(e => !e)}
              type="button"
            >
              {expanded ? 'Show less' : 'Show more'}
            </button>
          )}
          {afterQuote}
        </>
      ) : (
        <>
          <p
            ref={commentRef}
            className={`text-sm text-gray-800 ${!expanded ? 'line-clamp-3' : ''}`}
            style={
              !expanded
                ? {
                    display: '-webkit-box',
                    WebkitLineClamp: 3,
                    WebkitBoxOrient: 'vertical',
                    overflow: 'hidden'
                  }
                : {}
            }
          >
            {commentContent}
          </p>
          {isLong && (
            <button
              className="text-xs text-gray-600 mt-1 rounded bg-gray-200 bg-opacity-0 hover:bg-opacity-100 transition duration-150 px-2 py-1"
              onClick={() => setExpanded(e => !e)}
              type="button"
            >
              {expanded ? 'Show less' : 'Show more'}
            </button>
          )}
        </>
      )}
      <div className="text-xs text-gray-500 mt-2 pt-2 border-t border-gray-200 flex justify-between items-center">
        <div className="flex items-center gap-2">
          {tags.length > 0 && tags.map((tag, idx) => {
            // "Comment" tags already filtered out above
            const tagInfo = findTagInfo(tag);
            const tagClass = tagInfo
              ? tagInfo.tagClass
              : "px-2 py-0.5 font-semibold rounded-full bg-blue-100 text-blue-800";
            // For LDA subclass, display the full label (e.g., "LDA 10/3/25")
            const tagLabel = tagInfo && tagInfo.label ? tagInfo.label : tag;
            return (
              <span key={tag + idx} className={tagClass}>
                {tagLabel}
              </span>
            );
          })}
          <span className="font-medium">{entry.createdBy}</span>
        </div>
        <span>
          {formattedTimestamp}
        </span>
      </div>
    </li>
  );
}

export default Comment;
