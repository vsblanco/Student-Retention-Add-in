import React, { useState } from 'react';
import { formatExcelDate, formatName } from '../../utility/Conversion';
import BounceAnimation from '../../utility/BounceAnimation';
import CommentModal from '../Modal/CommentModal';
import { toast } from 'react-toastify';

// CSS class constants
const liStyle = "p-3 rounded-lg shadow-sm relative";
const borderLeftStyle = "border-l-4 pl-4";
const tagPillStyle = "px-2 py-0.5 font-semibold rounded-full";
const tagDefaultStyle = `${tagPillStyle} bg-gray-100 text-gray-800`;
const plusPillStyle = `${tagPillStyle} bg-white text-gray-700 text-xs opacity-50`;
const createdByStyle = "font-medium whitespace-nowrap";
const tagsRowStyle = "flex items-center gap-1";
const timestampRowStyle = "text-xs text-gray-500 mt-2 pt-2 border-t border-gray-200 flex justify-between items-center";
const commentTextStyle = "text-sm text-gray-800";
const commentClampStyle = "line-clamp-3";
const quoteBlockStyle = "relative bg-gray-100 border-l-4 border-blue-200 pl-6 pr-2 py-3 mb-2 rounded";
const quoteTextStyle = "text-sm text-gray-600 font-serif";
const quoteClampStyle = "line-clamp-3";
const quoteMarkStyle = "absolute left-2 top-2 text-4xl text-gray-400 leading-none select-none";
const quoteMarkRightStyle = "absolute right-2 bottom-2 text-4xl text-gray-400 leading-none select-none";
const showMoreBtnStyle = "text-xs text-gray-600 mt-1 rounded bg-gray-100 bg-opacity-0 hover:bg-opacity-100 transition duration-150 px-2 py-1";
const showMoreBtnStyleAlt = "text-xs text-gray-600 mt-1 rounded bg-gray-200 bg-opacity-0 hover:bg-opacity-100 transition duration-150 px-2 py-1";

// Tag definitions for comments
export const COMMENT_TAGS = [
  {
    label: "Urgent",
    bgClass: "bg-red-50",
    tagClass: "px-2 py-0.5 font-semibold rounded-full bg-red-100 text-red-800",
    priority: 2,
    borderColor: "border-red-400",
    about: 'Reserved for urgent attention',
  },
  {
    label: "Note",
    bgClass: "bg-gray-200",
    tagClass: "px-2 py-0.5 font-semibold rounded-full bg-gray-700 text-gray-200",
    pinned: true,
    priority: 3,
    borderColor: "border-gray-400",
    about: 'A pinned note for general information',
  },
  {
    label: "DNC",
    bgClass: "bg-red-100",
    tagClass: "px-2 py-0.5 font-semibold rounded-full bg-red-200 text-black",
    pinned: true,
    priority: 4,
    borderColor: "border-red-600",
    about: 'Student requested not to be contacted',
    subtags: [
      {
        label: "DNC - Phone",
        bgClass: "bg-red-100",
        tagClass: "px-2 py-0.5 font-semibold rounded-full bg-red-200 text-black",
        pinned: true,
        priority: 4,
        borderColor: "border-red-600",
        about: 'Student requested no phone contact'
      },
      {
        label: "DNC - Other Phone",
        bgClass: "bg-red-100",
        tagClass: "px-2 py-0.5 font-semibold rounded-full bg-red-200 text-black",
        pinned: true,
        priority: 4,
        borderColor: "border-red-600",
        about: 'Student requested no other phone contact'
      },
      {
        label: "DNC - Email",
        bgClass: "bg-red-200",
        tagClass: "px-2 py-0.5 font-semibold rounded-full bg-red-200 text-black",
        pinned: true,
        priority: 4,
        borderColor: "border-red-600",
        about: 'Student requested no email contact'
      }
    ]
  },
  {
    label: "LDA",
    bgClass: "bg-orange-100",
    tagClass: "px-2 py-0.5 font-semibold rounded-full bg-orange-200 text-orange-800",
    priority: 3,
    borderColor: "border-orange-400",
    about: 'Student plans to submit LDA'
  },
  {
    label: "Contacted",
    bgClass: "bg-yellow-100",
    tagClass: "px-2 py-0.5 font-semibold rounded-full bg-yellow-200 text-yellow-800",
    priority: 2,
    borderColor: "border-yellow-400",
    about: 'Student gave a response'
  },
  {
    label: "Outreach",
    bgClass: "bg-gray-100",
    tagClass: "px-2 py-0.5 font-semibold rounded-full bg-gray-200 text-gray-800",
    priority: 1,
    borderColor: "border-gray-300",
    about: 'Sourced from the Outreach Column'
  },
  {
    label: "Quote",
    bgClass: "bg-gray-100",
    tagClass: "px-2 py-0.5 font-semibold rounded-full bg-blue-50 text-blue-800",
    priority: 2,
    borderColor: "border-blue-400",
    about: 'Contains quoted text'
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

function Comment({ entry, searchTerm, index, onContextMenu, delete: deleteFunction, save: saveFunction }) {
  // Support multiple tags separated by commas
  let tags = entry.tag
    ? entry.tag.split(',').map(t => t.trim()).filter(Boolean)
    : [];
  // Remove "Comment" tags
  tags = tags.filter(t => t !== "Comment");
  // Find tag info for all tags
  const tagInfos = tags.map(findTagInfo);

  // Sort tags by descending priority
  const sortedTagInfos = [...tagInfos].sort((a, b) => {
    const pa = a?.priority ?? -Infinity;
    const pb = b?.priority ?? -Infinity;
    return pb - pa;
  });

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
  const bgClass = tagInfo && tagInfo.bgClass ? tagInfo.bgClass : "bg-gray-100";
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
      commentContent = parts.flatMap(part => highlightLdaKeywords(part));
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

  // New helper: produce a date-only string (prefer Excel serial -> JS Date if numeric)
  const getDateOnlyString = () => {
    if (entry.timestamp == null) return formattedTimestamp;
    // If numeric Excel serial, convert to JS date
    if (!isNaN(entry.timestamp)) {
      try {
        // Excel's epoch: 1899-12-30 (approx). For simplicity we use this base.
        const excelSerial = Number(entry.timestamp);
        const excelEpoch = new Date(1899, 11, 30);
        const jsDate = new Date(excelEpoch.getTime() + excelSerial * 86400000);
        return jsDate.toLocaleDateString();
      } catch (e) {
        // fallback to formatted string
      }
    }
    // Try to extract a common date pattern (MM/DD/YYYY or M/D/YY etc.)
    const dateMatch = String(formattedTimestamp).match(/\b\d{1,2}[\/-]\d{1,2}[\/-]\d{2,4}\b/);
    if (dateMatch) return dateMatch[0];
    // Fallback: strip time portions like " at 3:00 PM" or after comma
    return String(formattedTimestamp).split(' at ')[0].split(',')[0];
  };

  const formattedTimestampDateOnly = getDateOnlyString();

  // State for expanding/collapsing long comments and quotes
  const [expanded, setExpanded] = useState(false);

  // Local removal animation state for immediate visual feedback on delete
  const [removing, setRemoving] = useState(false);
  const [removed, setRemoved] = useState(false);

  // Ref and state for regular comment
  const commentRef = React.useRef(null);
  const [isLong, setIsLong] = useState(false);

  // Ref and state for quote text
  const quoteRef = React.useRef(null);
  const [isQuoteLong, setIsQuoteLong] = useState(false);

  // Bounce state for each tag pill (array of booleans)
  const [bounceTags, setBounceTags] = useState([]);

  // Modal state
  const [modalOpen, setModalOpen] = useState(false);

  // Ref for the tags container to decide how many pills we can show
  const tagsContainerRef = React.useRef(null);
  // Ref for the root comment <li> so we measure the whole comment container width
  const rootRef = React.useRef(null);

  // Max number of visible tag pills (1 for narrow, 2 for normal)
  const [maxPills, setMaxPills] = useState(2);

  // Keep previous width to log only when it actually changes
  const prevWidthRef = React.useRef(null);

  // Observe container size and switch pill count based on whole comment width
  React.useEffect(() => {
    const el = rootRef.current;
    if (!el) return;

    // Helper that applies measurement and logging based on the full comment container
    const applyWidth = (rawWidth) => {
      const w = Math.round(rawWidth || el.offsetWidth || el.clientWidth || 0);

      // Estimate how many pills can fit using the whole comment width.
      // Reserve space for created-by, timestamp and paddings then divide by approximated pill width.
      const RESERVED_SPACE = 150; // adjust if createdBy/timestamp layout changes
      const PILL_APPROX_WIDTH = 92; // approx per-tag pill width
      const available = Math.max(0, w - RESERVED_SPACE);
      const fitCount = Math.max(1, Math.floor(available / PILL_APPROX_WIDTH));
      const newMaxPills = Math.min(2, fitCount); // cap to 1..2

      setMaxPills(newMaxPills);

      // Log only when rounded width changes to keep logs meaningful, include comment index
      if (prevWidthRef.current !== w) {
        prevWidthRef.current = w;
      }
    };

    // ResizeObserver callback prefers entry.contentRect when available
    const update = (entries) => {
      if (entries && entries.length) {
        const e = entries[0];
        const w = e.contentRect ? e.contentRect.width : e.target.getBoundingClientRect().width;
        applyWidth(w);
      } else {
        // fallback (called on initial run or window resize fallback)
        const rect = el.getBoundingClientRect();
        applyWidth(rect.width);
      }
    };

    // Initial measurement
    update();

    let ro;
    if (typeof ResizeObserver !== 'undefined') {
      ro = new ResizeObserver(update);
      ro.observe(el);
    } else {
      window.addEventListener('resize', update);
    }
    return () => {
      if (ro) ro.disconnect();
      else window.removeEventListener('resize', update);
    };
  }, [index]);

  // Helper to trigger bounce for a tag index
  const triggerBounce = idx => {
    setBounceTags(prev => {
      const arr = [...prev];
      arr[idx] = true;
      return arr;
    });
    setTimeout(() => {
      setBounceTags(prev => {
        const arr = [...prev];
        arr[idx] = false;
        return arr;
      });
    }, 500);
  };

  // Show toast with tag about info
  const showTagAbout = tagInfo => {
    if (tagInfo && tagInfo.about) {
      toast.info(tagInfo.about, {
        position: "bottom-left",
        autoClose: 2000,
        hideProgressBar: false,
        closeOnClick: true,
        pauseOnHover: false,
        draggable: false,
        theme: "light",
        style: { fontSize: '1rem' }
      });
    }
  };

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

  // If completely removed, render nothing — placed after all hooks so hook order is stable
  if (removed) return null;

  // Determine border color from the highest priority tag
  const borderColorClass = tagInfo && tagInfo.borderColor ? tagInfo.borderColor : "border-gray-300";

  // Map bgClass to a slightly darker hover color
  const hoverBgMap = {
    'bg-red-50': '#FFE5E5',
    'bg-gray-200': '#D7DBDE',
    'bg-red-100': '#FAD7D7',
    'bg-orange-100': '#FFDEB8',
    'bg-yellow-100': '#F7ECC1',
    'bg-gray-100': '#EBEBEB',
    'bg-blue-100': '#bfdbfe',
    'bg-blue-50': '#EBEBEB',
  };
  const hoverBgColor = hoverBgMap[bgClass] || '#e0e7ff';

  // Helper for LDA tag display
  function renderTagLabel(tagInfo) {
    if (
      tagInfo &&
      /^LDA\s+(.+)/i.test(tagInfo.label)
    ) {
      // Extract date part
      const ldaDate = tagInfo.label.replace(/^LDA\s+/i, '').trim();
      // Responsive: show only date on small screens or if container is too narrow
      // Use CSS: hide "LDA" prefix on xs/sm screens
      return (
        <span className="lda-tag-label">
          <span className="hidden sm:inline">LDA&nbsp;</span>
          {ldaDate}
        </span>
      );
    }
    return tagInfo.label;
  }

  return (
    <>
      <li
        ref={rootRef}
        className={`${liStyle} ${bgClass} ${borderLeftStyle} ${borderColorClass}`}
        data-row-index={entry.studentId || index}
        onContextMenu={onContextMenu}
        style={{
          position: 'relative',
          transition: 'background 0.15s, box-shadow 0.15s, opacity 300ms ease, transform 300ms ease, height 300ms ease, margin 300ms ease, padding 300ms ease',
          // animate out when removing
          opacity: removing ? 0 : 1,
          transform: removing ? 'translateX(8px) scale(0.995)' : 'none',
          // collapse visually by letting inline styles override paddings/margins/height when removing
          overflow: removing ? 'hidden' : undefined,
          paddingTop: removing ? 0 : undefined,
          paddingBottom: removing ? 0 : undefined,
          marginTop: removing ? 0 : undefined,
          marginBottom: removing ? 0 : undefined,
          height: removing ? 0 : undefined,
          pointerEvents: removing ? 'none' : undefined
        }}
        onClick={() => setModalOpen(true)}
        tabIndex={0}
        role="button"
        aria-label="Open comment modal"
        onMouseEnter={e => {
          e.currentTarget.style.background = hoverBgColor;
          e.currentTarget.style.boxShadow = '0 4px 16px rgba(59,130,246,0.10)';
        }}
        onMouseLeave={e => {
          e.currentTarget.style.background = '';
          e.currentTarget.style.boxShadow = '';
        }}
      >
        <BounceAnimation />
        {hasQuoteTag && quoteText ? (
          <>
            {beforeQuote}
            <blockquote className={quoteBlockStyle}>
              <span className={quoteMarkStyle} aria-hidden="true">“</span>
              <span
                ref={quoteRef}
                className={`${quoteTextStyle} ${!expanded ? quoteClampStyle : ""}`}
                style={
                  !expanded
                    ? {
                        display: '-webkit-box',
                        WebkitLineClamp: 3,
                        WebkitBoxOrient: 'vertical',
                        overflow: 'hidden',
                        userSelect: 'none'
                      }
                    : { userSelect: 'none' }
                }
              >
                {quoteText}
              </span>
              <span className={quoteMarkRightStyle} aria-hidden="true">”</span>

              {/* Moved attribution inside the quote container but outside the actual quoted text */}
              <div className={`${quoteTextStyle} mt-1 text-right`}>
                - {formatName(entry.studentName || entry.student || entry.createdby) || 'Unknown'}
              </div>
            </blockquote>

            {afterQuote}
          </>
        ) : (
          <>
            <p
              ref={commentRef}
              className={`${commentTextStyle} ${!expanded ? commentClampStyle : ""}`}
              style={
                !expanded
                  ? {
                      display: '-webkit-box',
                      WebkitLineClamp: 3,
                      WebkitBoxOrient: 'vertical',
                      overflow: 'hidden',
                      userSelect: 'none'
                    }
                  : { userSelect: 'none' }
              }
            >
              {commentContent}
            </p>
          </>
        )}
        <div className={timestampRowStyle}>
          {/* attach the ref to the real tags container so we measure a visible box */}
          <div ref={tagsContainerRef} className={tagsRowStyle} style={{ flex: 1 }}>
            {sortedTagInfos.slice(0, maxPills).map((tagInfo, idx) => {
               if (!tagInfo) return null;
               let tagClass = tagInfo.tagClass
                 ? tagInfo.tagClass
                 : tagDefaultStyle;
               // If more than one tag and Outreach, set opacity
               if (sortedTagInfos.length > 1 && tagInfo.label === "Outreach") {
                 tagClass += " opacity-75";
               }
               // Use renderTagLabel for LDA tags
               return (
                 <span
                   key={tagInfo.label + idx}
                   className={`${tagClass}${bounceTags[idx] ? ' bounce' : ''}`}
                   style={{
                     maxWidth: 90,
                     overflow: 'hidden',
                     textOverflow: 'ellipsis',
                     whiteSpace: 'nowrap',
                     cursor: 'pointer',
                     userSelect: 'none' // prevent highlighting
                   }}
                   onClick={e => {
                     e.stopPropagation(); // Prevent opening modal
                     triggerBounce(idx);
                     showTagAbout(tagInfo);
                   }}
                   tabIndex={0}
                   aria-label={`Bounce ${tagInfo.label} tag`}
                 >
                   {renderTagLabel(tagInfo)}
                 </span>
               );
             })}
            {sortedTagInfos.length > maxPills && (
              <span className={plusPillStyle}>
                +{sortedTagInfos.length - maxPills}
              </span>
            )}
          </div>
          <span className={createdByStyle} style={{ marginLeft: 'auto' }}>
            {entry.createdby ? entry.createdby : "Unknown"}
          </span>
          <span className="mx-2 text-gray-400">|</span>
          <span>
            {/* Show date-only when we've reduced visible pills to 1, otherwise full timestamp */}
            {maxPills === 1 ? formattedTimestampDateOnly : formattedTimestamp}
          </span>
        </div>
      </li>
      <CommentModal
        isOpen={modalOpen}
        onSaved={(updatedEntry) => {
          saveFunction(updatedEntry);
        }}
        onClose={() => setModalOpen(false)}
        // onDeleted gives immediate feedback: animate out locally then unmount
        onDeleted={(commentID) => {
          deleteFunction(commentID);
          // close modal immediately (safety)
          setModalOpen(false);
          // start removal animation
         // setRemoving(true);
          // unmount after animation completes
         // setTimeout(() => setRemoved(true), 350);
        }}
        entry={entry}
        COMMENT_TAGS={COMMENT_TAGS}
        findTagInfo={findTagInfo}
        hasQuoteTag={hasQuoteTag}
        quoteText={quoteText}
        beforeQuote={beforeQuote}
        afterQuote={afterQuote}
        formatExcelDate={formatExcelDate}
        quoteStyles={{
          block: quoteBlockStyle,
          text: quoteTextStyle,
          markLeft: quoteMarkStyle,
          markRight: quoteMarkRightStyle
        }}
      />
    </>
  );
}

// CommentSkeleton: lightweight placeholder for loading states.
// Usage: <CommentSkeleton /> or <CommentSkeleton showAvatar={false} />
export function CommentSkeleton({ lines = 1, showAvatar = true, showTags = true }) {
  // shimmer gradient style applied inline; keyframes defined inside the returned fragment
  const gradientBase = 'linear-gradient(90deg, rgba(0,0,0,0.06) 0%, rgba(0,0,0,0.035) 20%, rgba(0,0,0,0.06) 40%, rgba(0,0,0,0.035) 60%, rgba(0,0,0,0.06) 100%)';
  const baseStyle = {
    background: gradientBase,
    backgroundSize: '200% 100%',
    animation: 'shimmer 3s linear infinite'
  };

  const lineCount = Math.max(1, lines);
  const contentLines = Array.from({ length: lineCount }).map((_, i) => (
    <div
      key={i}
      style={{
        ...baseStyle,
        width: i === 0 ? '90%' : (i === lineCount - 1 ? '60%' : '85%')
      }}
    />
  ));

  return (
    <>
      <style>{`
        @keyframes shimmer {
          0% { background-position: 200% 0; }
          100% { background-position: -200% 0; }
        }
      `}</style>

      <li className={`${liStyle} bg-gray-50 ${borderLeftStyle} border-gray-300`} aria-hidden="true" style={{ overflow: 'hidden' }}>
        <div className="flex items-start">
          {showAvatar && (
            <div
              style={{
                ...baseStyle,
                width: 36,
                height: 36,
                borderRadius: '50%',
                marginRight: 12
              }}
            />
          )}
          <div style={{ flex: 1 }}>
            {/* headline / short line */}
            <div style={{ ...baseStyle, height: 12, width: '45%', marginBottom: 10, borderRadius: 6 }} />

            {/* content placeholder lines */}
            {contentLines}

            {/* tags / timestamp row visually matching Comment layout */}
            <div className="mt-2 flex items-center gap-2" style={{ marginTop: 8 }}>
              {showTags && (
                <>
                  <div style={{ ...baseStyle, width: 72, height: 22, borderRadius: 9999 }} />
                  <div style={{ ...baseStyle, width: 56, height: 22, borderRadius: 9999 }} />
                </>
              )}
              <div style={{ flex: 1 }} />
              <div style={{ ...baseStyle, width: 80, height: 12, borderRadius: 6 }} />
            </div>
          </div>
        </div>

        {/* small timestamp line to mirror real comment footer */}
        <div className={timestampRowStyle} style={{ marginTop: 12 }}>
          <div style={{ ...baseStyle, width: 120, height: 10, borderRadius: 6 }} />
        </div>
      </li>
    </>
  );
}

// Helper to highlight LDA keywords (exported for reuse)
export function highlightLdaKeywords(part) {
  if (typeof part !== "string") return part;

  // Keywords to highlight
  const keywords = [
    "Tomorrow", "next week",
    "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday",
    "weekend", "weekends"
  ];

  // Build regex parts:
  // - shortDate: matches MM/DD or M/D and allows optional year (MM/DD/YY or MM/DD/YYYY), accepts / or - separators
  // - keyword group: word-boundary match for listed keywords (case-insensitive)
  const shortDatePart = '\\b\\d{1,2}[\\/\\-]\\d{1,2}(?:[\\/\\-]\\d{2,4})?\\b';
  const keywordPart = `\\b(?:${keywords.map(k => k.replace(/[.*+?^${}()|[\\]\\\\]/g, '\\$&')).join('|')})\\b`;
  const combinedRegex = new RegExp(`${shortDatePart}|${keywordPart}`, 'gi');

  const parts = [];
  let lastIdx = 0;
  let match;
  while ((match = combinedRegex.exec(part)) !== null) {
    const idx = match.index;
    if (idx > lastIdx) {
      parts.push(part.slice(lastIdx, idx));
    }
    // Bold the matched keyword or date
    parts.push(<b key={`lda-match-${idx}`}>{match[0]}</b>);
    lastIdx = combinedRegex.lastIndex;
  }
  if (lastIdx < part.length) {
    parts.push(part.slice(lastIdx));
  }

  return parts.length > 0 ? parts : part;
}

// New helper: extract LDA match strings (normalized as M/D/YY)
export function extractLdaMatches(text) {
  if (!text || typeof text !== 'string') return [];
  const keywords = [
    "Tomorrow", "next week",
    "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday",
    "weekend", "weekends"
  ];
  const shortDatePart = '\\b\\d{1,2}[\\/\\-]\\d{1,2}(?:[\\/\\-]\\d{2,4})?\\b';
  const keywordPart = `\\b(?:${keywords.map(k => k.replace(/[.*+?^${}()|[\\]\\\\]/g, '\\$&')).join('|')})\\b`;
  const combinedRegex = new RegExp(`${shortDatePart}|${keywordPart}`, 'gi');

  const results = [];
  let match;
  const now = new Date();

  const formatTwoDigit = d => {
    const mm = d.getMonth() + 1;
    const dd = d.getDate();
    const yy = String(d.getFullYear()).slice(-2);
    return `${mm}/${dd}/${yy}`;
  };

  const weekdayIndex = name => {
    const map = { sunday:0, monday:1, tuesday:2, wednesday:3, thursday:4, friday:5, saturday:6 };
    return map[name.toLowerCase()];
  };

  while ((match = combinedRegex.exec(text)) !== null) {
    const token = match[0];
    // numeric date e.g., 10/7, 10/7/25 or 10-7-2025
    const numeric = token.match(/^(\d{1,2})[\/\-](\d{1,2})(?:[\/\-](\d{2,4}))?$/);
    if (numeric) {
      let month = parseInt(numeric[1], 10);
      let day = parseInt(numeric[2], 10);
      let yearPart = numeric[3];
      let year;
      if (!yearPart) {
        year = now.getFullYear();
      } else if (yearPart.length === 2) {
        year = 2000 + parseInt(yearPart, 10);
      } else {
        year = parseInt(yearPart, 10);
      }
      const d = new Date(year, month - 1, day);
      if (!isNaN(d.getTime())) results.push(formatTwoDigit(d));
      continue;
    }

    const lc = token.toLowerCase();
    if (lc === 'tomorrow') {
      const d = new Date(now);
      d.setDate(d.getDate() + 1);
      results.push(formatTwoDigit(d));
      continue;
    }
    if (lc === 'next week') {
      const d = new Date(now);
      d.setDate(d.getDate() + 7);
      results.push(formatTwoDigit(d));
      continue;
    }
    if (lc === 'weekend' || lc === 'weekends') {
      // return next Saturday
      const d = new Date(now);
      const target = 6; // Saturday
      const diff = (target + 7 - d.getDay()) % 7 || 7;
      d.setDate(d.getDate() + diff);
      results.push(formatTwoDigit(d));
      continue;
    }
    // weekday names -> next occurrence of that weekday
    const wIdx = weekdayIndex(token);
    if (typeof wIdx === 'number') {
      const d = new Date(now);
      const diff = (wIdx + 7 - d.getDay()) % 7 || 7;
      d.setDate(d.getDate() + diff);
      results.push(formatTwoDigit(d));
      continue;
    }
  }

  // dedupe and return normalized strings
  return Array.from(new Set(results));
}

export default Comment;