import React, { useState } from 'react';
import { formatExcelDate } from '../utility/Conversion';
import BounceAnimation from '../utility/BounceAnimation';
import Modal from '../utility/Modal'; // Add this import
import InsertTagButton from './InsertTagButton'; // <-- import InsertTagButton

// CSS class constants
const liStyle = "comment-li p-3 rounded-lg shadow-sm relative transition-colors";
const borderLeftStyle = "border-l-4 pl-4";
const tagPillStyle = "px-2 py-0.5 font-semibold rounded-full";
const tagDefaultStyle = `${tagPillStyle} bg-gray-100 text-gray-800`;
const plusPillStyle = `${tagPillStyle} bg-gray-200 text-gray-700 text-xs`;
const createdByStyle = "font-medium whitespace-nowrap";
const tagsRowStyle = "flex items-center gap-2";
const timestampRowStyle = "text-xs text-gray-500 mt-2 pt-2 border-t flex justify-between items-center";
const commentTextStyle = "text-sm text-gray-800";
const commentClampStyle = "line-clamp-3";
const quoteBlockStyle = "relative bg-blue-50 border-l-4 border-blue-200 pl-6 pr-2 py-3 mb-2 rounded";
const quoteTextStyle = "text-base text-blue-900 font-serif";
const quoteClampStyle = "line-clamp-3";
const quoteMarkStyle = "absolute left-2 top-2 text-4xl text-blue-200 leading-none select-none";
const quoteMarkRightStyle = "absolute right-2 bottom-2 text-4xl text-blue-200 leading-none select-none";
const showMoreBtnStyle = "text-xs text-gray-600 mt-1 rounded bg-gray-100 bg-opacity-0 hover:bg-opacity-100 transition duration-150 px-2 py-1";
const showMoreBtnStyleAlt = "text-xs text-gray-600 mt-1 rounded bg-gray-200 bg-opacity-0 hover:bg-opacity-100 transition duration-150 px-2 py-1";

// Tag definitions for comments
export const COMMENT_TAGS = [
  {
    label: "Urgent",
    bgClass: "bg-red-50",
    tagClass: "px-2 py-0.5 font-semibold rounded-full bg-red-100 text-red-800",
    priority: 2,
    borderColor: "border-red-400"
  },
  {
    label: "Note",
    bgClass: "bg-gray-200",
    tagClass: "px-2 py-0.5 font-semibold rounded-full bg-gray-700 text-gray-200",
    pinned: true,
    priority: 3,
    borderColor: "border-gray-400"
  },
  {
    label: "DNC",
    bgClass: "bg-red-100",
    tagClass: "px-2 py-0.5 font-semibold rounded-full bg-red-200 text-black",
    pinned: true,
    priority: 4,
    borderColor: "border-red-600",
    subtags: [
      {
        label: "DNC - Phone",
        bgClass: "bg-red-100",
        tagClass: "px-2 py-0.5 font-semibold rounded-full bg-red-200 text-black",
        pinned: true,
        priority: 4,
        borderColor: "border-red-600"
      },
      {
        label: "DNC - Other Phone",
        bgClass: "bg-red-100",
        tagClass: "px-2 py-0.5 font-semibold rounded-full bg-red-200 text-black",
        pinned: true,
        priority: 4,
        borderColor: "border-red-600"
      },
      {
        label: "DNC - Email",
        bgClass: "bg-red-200",
        tagClass: "px-2 py-0.5 font-semibold rounded-full bg-red-200 text-black",
        pinned: true,
        priority: 4,
        borderColor: "border-red-600"
      }
    ]
  },
  {
    label: "LDA",
    bgClass: "bg-orange-100",
    tagClass: "px-2 py-0.5 font-semibold rounded-full bg-orange-200 text-orange-800",
    priority: 3,
    borderColor: "border-orange-400"
  },
  {
    label: "Contacted",
    bgClass: "bg-yellow-100",
    tagClass: "px-2 py-0.5 font-semibold rounded-full bg-yellow-200 text-yellow-800",
    priority: 2,
    borderColor: "border-yellow-400"
  },
  {
    label: "Outreach",
    bgClass: "bg-gray-100",
    tagClass: "px-2 py-0.5 font-semibold rounded-full bg-gray-200 text-gray-800",
    priority: 1,
    borderColor: "border-gray-300"
  },
  {
    label: "Quote",
    bgClass: "bg-blue-100",
    tagClass: "px-2 py-0.5 font-semibold rounded-full bg-blue-50 text-blue-800",
    priority: 2,
    borderColor: "border-blue-400"
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

function Comment({ entry, searchTerm, index, onContextMenu, onEditComment }) {
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
  const tagClass = tagInfo ? tagInfo.tagClass : "px-2 py-0.5 font-semibold rounded-full bg-gray-100 text-gray-800";

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

  // Bounce state for each tag pill (array of booleans)
  const [bounceTags, setBounceTags] = useState([]);

  // Modal state
  const [modalOpen, setModalOpen] = useState(false);
  const [modalMode, setModalMode] = useState('view'); // 'view' or 'edit'

  // Textarea state for editing comment in modal
  const [modalComment, setModalComment] = useState(entry.comment || "");

  // Modal tag container object: { tagLabel: true }
  const [modalTagContainer, setModalTagContainer] = useState({});

  // State for delete confirmation in edit mode
  const [confirmDelete, setConfirmDelete] = useState(false);

  // Prefill modalTagContainer and comment when modal opens or entry changes
  React.useEffect(() => {
    setModalComment(entry.comment || "");
    let tags = entry.tag
      ? entry.tag.split(',').map(t => t.trim()).filter(Boolean)
      : [];
    tags = tags.filter(t => t !== "Comment");
    const tagObj = {};
    tags.forEach(t => { tagObj[t] = true; });
    setModalTagContainer(tagObj);
    setModalMode('view'); // Always start in view mode
  }, [entry.comment, entry.tag, modalOpen]);

  // Reset confirmDelete when modal or mode changes
  React.useEffect(() => {
    setConfirmDelete(false);
  }, [modalOpen, modalMode]);

  // Handler for opening modal
  const handleOpenModal = () => setModalOpen(true);
  const handleCloseModal = () => setModalOpen(false);

  // Handler for saving edited comment and tags
  const handleSaveComment = async () => {
    if (onEditComment) {
      // Join tags for saving
      const newTagString = Object.keys(modalTagContainer).join(', ');
      const success = await onEditComment(
        { ...entry, tag: newTagString },
        modalComment
      );
      if (success) setModalOpen(false);
    } else {
      setModalOpen(false);
    }
  };

  // Handler for deleting comment (implement your logic here)
  const handleDeleteComment = async () => {
    // Example: call a prop or show a confirmation dialog
    if (window.confirm("Are you sure you want to delete this comment?")) {
      if (typeof onEditComment === 'function') {
        await onEditComment({ ...entry, deleted: true }, null); // Or your delete logic
      }
      setModalOpen(false);
    }
  };

  // State for InsertTagButton dropdown
  const [showTagDropdown, setShowTagDropdown] = useState(false);
  const [dropdownTarget, setDropdownTarget] = useState(null);

  // Example tags for InsertTagButton (use COMMENT_TAGS for demo)
  const insertTagButtonTags = COMMENT_TAGS.map(tag => ({
    label: tag.label,
    spanClass: tag.tagClass,
    title: tag.label
  }));

  // Handler for tag click (add tag to container)
  const handleInsertTag = tagLabel => {
    setModalTagContainer(prev => ({
      ...prev,
      [tagLabel]: true
    }));
    setShowTagDropdown(false);
  };

  // Handler for removing a tag pill
  const handleRemoveTag = tagLabel => {
    setModalTagContainer(prev => {
      const newObj = { ...prev };
      delete newObj[tagLabel];
      return newObj;
    });
  };

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

  // Determine border color from the highest priority tag
  const borderColorClass = tagInfo && tagInfo.borderColor ? tagInfo.borderColor : "border-gray-300";

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

  // --- Tag pills rendering (all tags inside one pill-shaped outline box) ---
  const modalTagPillStyle = "px-1 py-0.5 text-xs font-semibold rounded-full";
  const modalTagDefaultStyle = `${modalTagPillStyle} bg-gray-100 text-gray-800`;
  const modalPlusPillStyle = `${modalTagPillStyle} bg-gray-200 text-gray-700 text-xs`;

  const modalTagPills = (
    <div
      style={{
        display: 'flex',
        flexWrap: 'wrap',
        gap: 4,
        padding: '6px 10px',
        borderRadius: 9999,
        border: '1px solid #cfcfcf',
        marginTop: 8,
        minHeight: 32,
        alignItems: 'center',
        width: '100%',
        boxSizing: 'border-box',
        maxWidth: '100%',
      }}
    >
      {Object.keys(modalTagContainer).map((tagLabel, idx) => {
        const tagInfo = findTagInfo(tagLabel);
        let tagClass = tagInfo?.tagClass
          ? tagInfo.tagClass.replace(/px-2|py-0\.5|text-[^\s]+/g, '')
          : "";
        tagClass = `${modalTagPillStyle} ${tagClass} ${tagInfo?.bgClass || ""} ${tagInfo?.textClass || ""}`;

        // Extract text color from tagClass (e.g., text-red-800)
        let textColor;
        if (tagInfo?.tagClass) {
          const match = tagInfo.tagClass.match(/text-([a-z]+-\d{3,4}|[a-z]+)/);
          if (match) {
            // Map Tailwind color to CSS color
            const tailwindToCss = {
              'text-red-800': '#991b1b',
              'text-gray-800': '#1f2937',
              'text-gray-200': '#e5e7eb',
              'text-black': '#000',
              'text-orange-800': '#9a3412',
              'text-yellow-800': '#854d0e',
              'text-blue-800': '#1e40af',
              'text-blue-900': '#1e3a8a',
              // Add more mappings as needed
            };
            textColor = tailwindToCss[match[0]] || undefined;
          }
        }

        return (
          <span
            key={tagLabel}
            className={tagClass}
            style={{
              overflow: 'hidden',
              textOverflow: 'ellipsis',
              whiteSpace: 'nowrap',
              fontSize: '0.8em',
              padding: '2px 6px',
              display: 'inline-flex',
              alignItems: 'center',
              background: tagInfo?.bgClass ? undefined : undefined,
              borderRadius: 9999,
              marginRight: 2,
              color: textColor, // <-- ensure correct text color
            }}
          >
            {tagLabel}
            <button
              type="button"
              aria-label={`Remove ${tagLabel}`}
              onClick={() => handleRemoveTag(tagLabel)}
              style={{
                marginLeft: 4,
                background: 'transparent',
                border: 'none',
                color: '#888',
                cursor: 'pointer',
                fontSize: '1em',
                lineHeight: 1
              }}
            >
              ×
            </button>
          </span>
        );
      })}
    </div>
  );

  // --- Modal tag pills for view mode (not editable) ---
  const modalTagViewPills = (
    <div
      style={{
        display: 'flex',
        flexWrap: 'wrap',
        gap: 4,
        padding: '6px 10px',
        borderRadius: 9999,
        border: '1px solid #dcdcdc',
        marginTop: 8,
        minHeight: 32,
        alignItems: 'center',
        width: '100%',
        boxSizing: 'border-box',
        maxWidth: '100%',
        background: 'transparent',
      }}
    >
      {Object.keys(modalTagContainer).map((tagLabel, idx) => {
        const tagInfo = findTagInfo(tagLabel);
        let tagClass = tagInfo?.tagClass
          ? tagInfo.tagClass.replace(/px-2|py-0\.5|text-[^\s]+/g, '')
          : "";
        tagClass = `${modalTagPillStyle} ${tagClass} ${tagInfo?.bgClass || ""} ${tagInfo?.textClass || ""}`;

        // Extract text color from tagClass (e.g., text-red-800)
        let textColor;
        if (tagInfo?.tagClass) {
          const match = tagInfo.tagClass.match(/text-([a-z]+-\d{3,4}|[a-z]+)/);
          if (match) {
            const tailwindToCss = {
              'text-red-800': '#991b1b',
              'text-gray-800': '#1f2937',
              'text-gray-200': '#e5e7eb',
              'text-black': '#000',
              'text-orange-800': '#9a3412',
              'text-yellow-800': '#854d0e',
              'text-blue-800': '#1e40af',
              'text-blue-900': '#1e3a8a',
              // Add more mappings as needed
            };
            textColor = tailwindToCss[match[0]] || undefined;
          }
        }

        return (
          <span
            key={tagLabel}
            className={tagClass}
            style={{
              overflow: 'hidden',
              textOverflow: 'ellipsis',
              whiteSpace: 'nowrap',
              fontSize: '0.8em',
              padding: '2px 6px',
              display: 'inline-flex',
              alignItems: 'center',
              background: tagInfo?.bgClass ? undefined : undefined,
              borderRadius: 9999,
              marginRight: 2,
              color: textColor, // <-- ensure correct text color
            }}
          >
            {tagLabel}
          </span>
        );
      })}
    </div>
  );

  // --- Modal content ---
  const modalContentView = (
    <div style={{ width: '100%' }}>
      {/* Tags (view only) */}
      {modalTagViewPills}
      {/* Comment (view only) */}
      <div style={{ marginTop: 12, marginBottom: 12 }}>
        <div
          style={{
            fontSize: '1rem',
            color: '#222',
            whiteSpace: 'pre-wrap',
            wordBreak: 'break-word',
            minHeight: 60,
            paddingLeft: 16,
            paddingTop: 6, // <-- add top padding
            borderRadius: 8,
            background: 'transparent',
            border: '1px solid #d5d5d5',
          }}
        >
          {/* Use quote rendering logic in view mode */}
          {hasQuoteTag && quoteText ? (
            <>
              {beforeQuote}
              <blockquote className={quoteBlockStyle} style={{ marginLeft: 0 }}>
                <span className={quoteMarkStyle} aria-hidden="true">“</span>
                <span className={quoteTextStyle}>
                  {quoteText}
                </span>
                <span className={quoteMarkRightStyle} aria-hidden="true">”</span>
              </blockquote>
              {afterQuote}
            </>
          ) : (
            <span>
              {entry.comment}
            </span>
          )}
        </div>
      </div>
      {/* Bottom row: created by and timestamp (bottom left), Edit button (bottom right) */}
      <div style={{
        display: 'flex',
        justifyContent: 'space-between',
        alignItems: 'center',
        marginTop: 8,
        width: '100%',
      }}>
        <div style={{ color: '#666', fontSize: '0.95em', paddingLeft: 4 }}>
          <span>
            {entry.createdby ? entry.createdby : "Unknown"}
          </span>
          <span style={{ margin: '0 8px', color: '#bbb' }}>|</span>
          <span>
            {!isNaN(entry.timestamp) ? formatExcelDate(entry.timestamp) : entry.timestamp}
          </span>
        </div>
        <button
          type="button"
          onClick={() => setModalMode('edit')}
          style={{
            padding: '6px 16px',
            borderRadius: 6,
            background: '#2563eb',
            color: 'white',
            fontWeight: 500,
            border: 'none',
            cursor: 'pointer'
          }}
        >
          Edit
        </button>
      </div>
    </div>
  );

  const modalContentEdit = (
    <div style={{ width: '100%' }}>
      {/* InsertTagButton above outline pill box */}
      <InsertTagButton
        dropdownId="modal-insert-tag-dropdown"
        onTagClick={handleInsertTag}
        showDropdown={showTagDropdown}
        setShowDropdown={setShowTagDropdown}
        dropdownTarget={dropdownTarget}
        setDropdownTarget={setDropdownTarget}
        targetName="modal-comment"
        tags={insertTagButtonTags}
      />
      {/* Modal tag pills below InsertTagButton, above textarea */}
      {modalTagPills}
      {/* Textarea for editing comment */}
      <textarea
        value={modalComment}
        onChange={e => setModalComment(e.target.value)}
        style={{
          width: '100%',
          minHeight: 80,
          fontSize: '1rem',
          marginBottom: 12,
          padding: 8,
          borderRadius: 6,
          border: '1px solid #cccccc',
          resize: 'vertical',
          boxSizing: 'border-box'
        }}
        placeholder="Edit comment..."
      />
      {/* Update, Cancel, and Delete buttons */}
      <div style={{
        display: 'flex',
        justifyContent: 'space-between',
        alignItems: 'center',
        marginTop: 16,
        width: '100%',
        gap: 8
      }}>
        <button
          type="button"
          onClick={() => setConfirmDelete(true)}
          style={{
            padding: '6px 16px',
            borderRadius: 6,
            background: confirmDelete ? '#fee2e2' : '#ef4444',
            color: confirmDelete ? '#b91c1c' : 'white',
            fontWeight: 500,
            border: 'none',
            cursor: 'pointer'
          }}
          disabled={confirmDelete}
        >
          {confirmDelete ? 'Delete...' : 'Delete'}
        </button>
        <div style={{ display: 'flex', gap: 8 }}>
          <button
            type="button"
            onClick={() => {
              if (confirmDelete) setConfirmDelete(false);
              else setModalMode('view');
            }}
            style={{
              padding: '6px 16px',
              borderRadius: 6,
              background: confirmDelete ? '#fee2e2' : '#e5e7eb',
              color: confirmDelete ? '#b91c1c' : '#222',
              fontWeight: 500,
              border: 'none',
              cursor: 'pointer'
            }}
          >
            {confirmDelete ? 'Cancel Delete' : 'Cancel'}
          </button>
          <button
            type="button"
            onClick={confirmDelete ? handleDeleteComment : handleSaveComment}
            disabled={
              confirmDelete
                ? false
                : modalComment === (entry.comment || "") &&
                  Object.keys(modalTagContainer).join(', ') === (entry.tag || "").split(',').map(t => t.trim()).filter(Boolean).filter(t => t !== "Comment").join(', ')
            }
            style={{
              padding: '6px 16px',
              borderRadius: 6,
              background: confirmDelete
                ? '#ef4444'
                : (
                  modalComment === (entry.comment || "") &&
                  Object.keys(modalTagContainer).join(', ') === (entry.tag || "").split(',').map(t => t.trim()).filter(Boolean).filter(t => t !== "Comment").join(', ')
                )
                  ? '#d1d5db'
                  : '#2563eb',
              color: confirmDelete
                ? 'white'
                : (
                  modalComment === (entry.comment || "") &&
                  Object.keys(modalTagContainer).join(', ') === (entry.tag || "").split(',').map(t => t.trim()).filter(Boolean).filter(t => t !== "Comment").join(', ')
                )
                  ? '#888'
                  : 'white',
              fontWeight: 500,
              border: 'none',
              cursor: confirmDelete
                ? 'pointer'
                : (
                  modalComment === (entry.comment || "") &&
                  Object.keys(modalTagContainer).join(', ') === (entry.tag || "").split(',').map(t => t.trim()).filter(Boolean).filter(t => t !== "Comment").join(', ')
                )
                  ? 'not-allowed'
                  : 'pointer',
              opacity: confirmDelete
                ? 1
                : (
                  modalComment === (entry.comment || "") &&
                  Object.keys(modalTagContainer).join(', ') === (entry.tag || "").split(',').map(t => t.trim()).filter(Boolean).filter(t => t !== "Comment").join(', ')
                ) ? 0.7 : 1
            }}
          >
            {confirmDelete ? 'Confirm Delete' : 'Update'}
          </button>
        </div>
      </div>
    </div>
  );

  return (
    <>
      <style>
        {`
          .comment-li:hover {
            filter: brightness(0.95);
          }
        `}
      </style>
      <li
        className={`${liStyle} ${bgClass} ${borderLeftStyle} ${borderColorClass}`}
        data-row-index={entry.studentId || index}
        onContextMenu={onContextMenu}
        style={{ position: 'relative', cursor: 'pointer' }}
        tabIndex={0}
        aria-label="Open comment modal"
        onClick={handleOpenModal} // Add this line
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
                        overflow: 'hidden'
                      }
                    : {}
                }
              >
                {quoteText}
              </span>
              <span className={quoteMarkRightStyle} aria-hidden="true">”</span>
            </blockquote>
            {isQuoteLong && (
              <button
                className={showMoreBtnStyle}
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
              className={`${commentTextStyle} ${!expanded ? commentClampStyle : ""}`}
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
                className={showMoreBtnStyleAlt}
                onClick={() => setExpanded(e => !e)}
                type="button"
              >
                {expanded ? 'Show less' : 'Show more'}
              </button>
            )}
          </>
        )}
        <div
          className={`${timestampRowStyle} ${borderColorClass}`}
          style={{
            borderTopWidth: '1px',
            borderTopStyle: 'solid',
            borderTopColor:
              borderColorClass === "border-red-400" ? "rgba(248,113,113,0.5)"
              : borderColorClass === "border-gray-400" ? "rgba(156,163,175,0.5)"
              : borderColorClass === "border-red-600" ? "rgba(220,38,38,0.5)"
              : borderColorClass === "border-orange-400" ? "rgba(251,146,60,0.5)"
              : borderColorClass === "border-yellow-400" ? "rgba(250,204,21,0.5)"
              : borderColorClass === "border-gray-300" ? "rgba(209,213,219,0.5)"
              : borderColorClass === "border-blue-400" ? "rgba(96,165,250,0.5)"
              : undefined
          }}
        >
          <div className={tagsRowStyle} style={{ flex: 1 }}>
            {sortedTagInfos.slice(0, 2).map((tagInfo, idx) => {
              if (!tagInfo) return null;
              // Hide Outreach tag if any other tag has higher priority
              if (
                tagInfo.label === "Outreach" &&
                sortedTagInfos.some(
                  t => t && t.label !== "Outreach" && (t.priority ?? -Infinity) > (tagInfo.priority ?? -Infinity)
                )
              ) {
                return null;
              }
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
                    userSelect: 'none'
                  }}
                  onClick={() => triggerBounce(idx)}
                  tabIndex={0}
                  aria-label={`Bounce ${tagInfo.label} tag`}
                >
                  {renderTagLabel(tagInfo)}
                </span>
              );
            })}
            {sortedTagInfos.length > 2 && (
              <span className={plusPillStyle}>
                +{sortedTagInfos.length - 2}
              </span>
            )}
          </div>
          <span className={createdByStyle} style={{ marginLeft: 'auto' }}>
            {entry.createdby ? entry.createdby : "Unknown"}
          </span>
          <span className="mx-2 text-gray-400">|</span>
          <span>
            {formattedTimestamp}
          </span>
        </div>
      </li>
      <Modal
        isOpen={modalOpen}
        onClose={handleCloseModal}
        padding="12px"
        style={{ minHeight: 200, width: '100%' }}
      >
        {modalMode === 'view' ? modalContentView : modalContentEdit}
      </Modal>
    </>
  );
}

export default Comment;
