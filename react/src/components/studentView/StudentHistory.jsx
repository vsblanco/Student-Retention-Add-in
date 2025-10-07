import React, { useState } from 'react';
import Comment, { COMMENT_TAGS } from './Comment';
import NewComment from './NewComment';

// Add styles constant
const styles = `
  @keyframes fadeInDrop {
    from {
      opacity: 0;
      transform: translateY(-24px);
    }
    to {
      opacity: 1;
      transform: translateY(0);
    }
  }
  .animate-fadein {
    animation: fadeInDrop 0.4s cubic-bezier(0.4,0,0.2,1);
  }
  /* See-through scrollbar styles */
  #history-content {
    scrollbar-width: thin; /* Firefox */
    scrollbar-color: rgba(0,0,0,0.15) rgba(0,0,0,0.03);
  }
  #history-content::-webkit-scrollbar {
    width: 8px;
    background: transparent;
  }
  #history-content::-webkit-scrollbar-thumb {
    background: rgba(0,0,0,0.15);
    border-radius: 8px;
  }
  #history-content::-webkit-scrollbar-track {
    background: rgba(0,0,0,0.03);
  }
`;

function StudentHistory({ history }) {
  // Normalize all keys in each history entry to lowercase and trimmed (e.g., "Student Name" -> "studentname")
  const normalizedHistory = Array.isArray(history)
    ? history.map(entry => {
        const normalized = {};
        Object.keys(entry || {}).forEach(key => {
          const normKey = String(key).toLowerCase().replace(/\s+/g, '');
          normalized[normKey] = entry[key];
        });
        return normalized;
      })
    : [];

  const [showSearch, setShowSearch] = useState(false);
  const [searchTerm, setSearchTerm] = useState('');
  const [showNewComment, setShowNewComment] = useState(false);

  // Filter history by search term (case-insensitive, matches comment)
  const filteredHistory = Array.isArray(normalizedHistory)
    ? normalizedHistory.filter(
        entry =>
          !searchTerm ||
          (entry.comment && entry.comment.toLowerCase().includes(searchTerm.toLowerCase()))
      )
    : [];

  // Helper to check if any tag or subtag in entry is pinned
  function isEntryPinned(entry) {
    if (!entry.tag) return false;
    const entryTags = entry.tag.split(',').map(t => t.trim());
    return entryTags.some(tagLabel => {
      const tagObj = COMMENT_TAGS.find(t => t.label === tagLabel);
      if (tagObj) {
        if (tagObj.pinned) return true;
        // Check subtags if present
        if (Array.isArray(tagObj.subtags)) {
          return tagObj.subtags.some(subtag => subtag.pinned);
        }
      }
      return false;
    });
  }

  // Split filteredHistory into pinned and unpinned, then sort unpinned by recency (reverse order)
  const pinnedComments = [];
  const unpinnedComments = [];
  [...filteredHistory].forEach(entry => {
    if (isEntryPinned(entry)) {
      pinnedComments.push(entry);
    } else {
      unpinnedComments.push(entry);
    }
  });

  if (!Array.isArray(normalizedHistory) || normalizedHistory.length === 0) {
    return (
      <div id="history-content">
        <ul className="space-y-4">
          <li className="p-3 bg-gray-100 rounded-lg shadow-sm relative">
            <p className="text-sm text-gray-800">No history found for this student.</p>
          </li>
        </ul>
      </div>
    );
  }

  return (
    <div>
      <style>{styles}</style>
      {/* History Header */}
      <div className="sticky-header space-y-4">
        <div className="flex justify-between items-center">
          <h3 className="text-lg font-bold text-gray-800">History</h3>
          <div className="flex items-center space-x-2">
            <button
              id="search-history-button"
              className="bg-gray-600 text-white w-8 h-8 rounded-full shadow-lg flex items-center justify-center hover:bg-gray-700"
              onClick={() => setShowSearch(v => !v)}
              aria-label="Search history"
            >
              <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z"></path></svg>
            </button>
            <button
              id="add-comment-button"
              className="bg-blue-600 text-white w-8 h-8 rounded-full shadow-lg flex items-center justify-center hover:bg-blue-700"
              onClick={() => setShowNewComment(v => !v)}
              aria-label="Add comment"
            >
              <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M12 6v6m0 0v6m0-6h6m-6 0H6"></path></svg>
            </button>
          </div>
        </div>

        <div
          id="search-container"
          className={`${showSearch ? '' : 'hidden'} space-y-2`}
        >
          <div id="tag-filter-container" className="flex flex-wrap items-center gap-2 pb-2 border-b border-gray-200">
            {/* Removed InsertTagButton for filtering history by tag */}
          </div>
          <div className="relative">
            <input
              type="text"
              id="search-input"
              className="w-full p-2 pl-8 border rounded-md"
              placeholder="Search comments..."
              value={searchTerm}
              onChange={e => setSearchTerm(e.target.value)}
              autoFocus={showSearch}
            />
            <div className="absolute inset-y-0 left-0 pl-2 flex items-center pointer-events-none">
              <svg className="h-5 w-5 text-gray-400" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z"></path></svg>
            </div>
            <button
              id="clear-search-button"
              className="absolute inset-y-0 right-0 pr-3 flex items-center text-gray-500 hover:text-gray-700"
              type="button"
              onClick={() => setSearchTerm('')}
              aria-label="Clear search"
            >
              ×
            </button>
          </div>
        </div>

        {/* Move new comment box to NewComment component */}
        <NewComment
          show={showNewComment}
          onClose={() => setShowNewComment(false)}
        />
      </div>
      {/* End History Header */}

      <div
        id="history-content"
        className="overflow-y-auto"
        style={{
          height: 'calc(100vh - 260px)'
        }}
      >
        <ul className="space-y-4">
          {/* Pinned comments first (keep their original order) */}
          {pinnedComments.map((entry, index) => (
            <Comment
              key={`pinned-${index}`}
              entry={entry}
              searchTerm={searchTerm}
              index={index}
            />
          ))}
          {/* Then unpinned comments, most recent first */}
          {[...unpinnedComments].reverse().map((entry, index) => (
            <Comment
              key={`unpinned-${index}`}
              entry={entry}
              searchTerm={searchTerm}
              index={index}
            />
          ))}
        </ul>
      </div>
    </div>
  );
}
export default StudentHistory;


