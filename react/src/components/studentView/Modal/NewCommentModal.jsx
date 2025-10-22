import React, { useRef, useState } from 'react';
import { useRef as useRefHook, useEffect as useEffectHook } from 'react';
import InsertTagButton from '../Parts/InsertTagButton';
import TagContainer from '../Parts/TagContainer';
import { COMMENT_TAGS } from '../Parts/Comment.jsx';

const styles = `
  @keyframes fadeInDrop {
    from {
      opacity: 0.5;
      transform: translateY(-1px);
    }
    to {
      opacity: 1;
      transform: translateY(0);
    }
  }

  .animate-fadein {
    animation: fadeInDrop 0.4s cubic-bezier(0.4, 0, 0.2, 1);
  }
`;

function NewComment({ show, onClose, addCommentToHistory, initialComment = "" }) {
  const newCommentInputRef = useRef(null);
  const [loading, setLoading] = useState(false);
  const isMountedRef = useRef(null);

  // dropdown state for InsertTagButton
  const [showDropdown, setShowDropdown] = useState(false);
  const [dropdownTarget, setDropdownTarget] = useState(null);
  // tags inserted by the user (displayed as pills)
  const [insertedTags, setInsertedTags] = useState([]);
  // remove an inserted tag (used by TagContainer "×")
  const handleRemoveInsertedTag = (label) => {
    setInsertedTags(prev => prev.filter(t => (typeof t === 'string' ? t : t.label) !== label));
  };

  // use shared COMMENT_TAGS only (empty array if unavailable)
  const tags = (Array.isArray(COMMENT_TAGS) ? COMMENT_TAGS : []).map(t => ({
    label: t.label,
    title: t.title || t.label,
    spanClass: t.tagClass || t.spanClass || t.bgClass || ''
  }));

  // insert tag label at cursor position inside textarea
  const handleTagClick = (label) => {
    // Do NOT insert the tag label into the textarea.
    // Only record the selected tag in the insertedTags pill list and close the dropdown.
    setShowDropdown(false);
    setDropdownTarget(null);
    const tagObj = tags.find(t => t.label === label) || { label, spanClass: 'bg-gray-100 text-gray-800' };
    setInsertedTags(prev => (prev.some(t => (typeof t === 'string' ? t : t.label) === tagObj.label) ? prev : [...prev, tagObj]));
  };

  React.useEffect(() => {
    isMountedRef.current = true;
    return () => {
      isMountedRef.current = false;
    };
  }, []);

  React.useEffect(() => {
    if (newCommentInputRef.current) {
      newCommentInputRef.current.value = initialComment;
    }
  }, [initialComment, show]);

  const handleSubmit = async () => {
    const comment = newCommentInputRef.current.value.trim();
    if (!comment) {
      // keep UX silent on empty comment per request
      return;
    }
    setLoading(true);
    // close modal immediately
    if (onClose) onClose();

    // build comma-separated tag string from insertedTags
    const tagString = (insertedTags || [])
      .map(t => (typeof t === 'string' ? t : (t && t.label) || ''))
      .filter(Boolean)
      .join(', ');

    // continue submission in background; guard state updates if unmounted
    try {
      console.log('Submitting comment with tags:', tagString);
      const success = await addCommentToHistory(comment, tagString);
      if (success) {
        // clear input if still mounted
        if (isMountedRef.current && newCommentInputRef.current) {
          newCommentInputRef.current.value = '';
          // clear inserted tags after successful submit
          setInsertedTags([]);
        }
      } else {
        console.error('Failed to add comment');
      }
    } catch (err) {
      console.error('Error adding comment', err);
    } finally {
      if (isMountedRef.current) setLoading(false);
    }
  };

  return (
    <>
      <style>{styles}</style>
      <div
        id="new-comment-section"
        className={`relative p-1 bg-gradient-to-br from-white via-gray-50 to-gray-100 border border-gray-200 rounded-2xl shadow-xl transition-all duration-200 ${
          show ? 'animate-fadein opacity-100' : 'hidden opacity-0'
        }`}
      >
        {/* Display inserted tags (pills) then the InsertTagButton */}
        <div className="flex flex-col gap-2 mb-3">
          <div className="flex items-center justify-start">
            <InsertTagButton
              dropdownId="new-comment-tags"
              onTagClick={handleTagClick}
              showDropdown={showDropdown}
              setShowDropdown={setShowDropdown}
              dropdownTarget={dropdownTarget}
              setDropdownTarget={setDropdownTarget}
              targetName="newComment"
              tags={tags}
            />
          </div>
          <TagContainer tags={insertedTags} onRemove={handleRemoveInsertedTag} />
        </div>

        <textarea
          id="new-comment-input"
          ref={newCommentInputRef}
          className="w-full p-3 border border-gray-300 rounded-xl bg-white shadow-inner focus:ring-2 focus:ring-blue-400 focus:border-blue-400 transition-all duration-150 text-base placeholder-gray-400 resize-vertical"
          rows={3}
          placeholder="Add a new comment..."
        ></textarea>
        <div className="flex justify-end items-center mt-4">
           <button
             id="submit-comment-button"
             className="px-6 py-2 bg-gradient-to-r from-blue-600 to-blue-500 text-white font-semibold rounded-xl shadow-md hover:from-blue-700 hover:to-blue-600 focus:outline-none focus:ring-2 focus:ring-blue-400 transition-all duration-150 disabled:bg-gray-400 disabled:from-gray-400 disabled:to-gray-400"
             onClick={handleSubmit}
             disabled={loading}
           >
             Submit
           </button>
         </div>
       </div>
     </>
   );
}

export default NewComment;
