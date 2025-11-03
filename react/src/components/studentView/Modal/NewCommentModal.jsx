import React, { useRef, useState } from 'react';
import { useRef as useRefHook, useEffect as useEffectHook } from 'react';
import InsertTagButton from '../Parts/InsertTagButton';
import TagContainer from '../Parts/TagContainer';
import { COMMENT_TAGS, extractLdaMatches } from '../Parts/Comment.jsx'; // added extractLdaMatches
import { DNCModal, LDAModal } from '../Tag.jsx';

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

function NewComment({ show, onClose, addCommentToHistory, initialComment = "", phone = "", otherPhone = "", email = "" }) {
  const newCommentInputRef = useRef(null);
  const [loading, setLoading] = useState(false);
  const isMountedRef = useRef(null);

  // DNC / LDA modal state (copied behavior from CommentModal.jsx)
  const [showDNCModal, setShowDNCModal] = useState(false);
  const [pendingDncTag, setPendingDncTag] = useState(false);
  const [showLDAModal, setShowLDAModal] = useState(false);

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

  // Refs for debouncing LDA removals and remembering last matches
  const ldaRemovalTimerRef = React.useRef(null);
  const lastLdaMatchesRef = React.useRef([]);

  // Helper: update insertedTags to reflect a concrete set of LDA matches (immediate apply)
  const updateInsertedTagsForMatches = (matches) => {
    const newLdaLabels = (matches || []).map(m => `LDA ${m}`);
    const parent = tags.find(t => (t.label || '').toUpperCase() === 'LDA');
    setInsertedTags(prev => {
      // Keep non-LDA tags
      const nonLda = prev.filter(t => {
        const lbl = (typeof t === 'string' ? t : t.label) || '';
        return !lbl.toUpperCase().startsWith('LDA');
      });

      // Build existing labels set to avoid duplicates
      const existing = new Set(prev.map(t => (typeof t === 'string' ? t : t.label)));

      // Add only new LDA labels that aren't already present
      const additions = newLdaLabels
        .filter(l => !existing.has(l))
        .map(l => ({
          label: l,
          title: l,
          spanClass: parent ? parent.spanClass : 'bg-gray-100 text-gray-800'
        }));

      // Return non-LDA tags followed by new (or retained) LDA tags
      return [...nonLda, ...additions];
    });
  };

  // Schedule removal of all LDA tags after a short pause (to avoid oscillation while typing).
  const scheduleRemoveLdaTags = (delay = 800) => {
    if (ldaRemovalTimerRef.current) clearTimeout(ldaRemovalTimerRef.current);
    ldaRemovalTimerRef.current = setTimeout(() => {
      // remove all LDA tags if no matches persisted
      setInsertedTags(prev => prev.filter(t => {
        const lbl = (typeof t === 'string' ? t : t.label) || '';
        return !lbl.toUpperCase().startsWith('LDA');
      }));
      lastLdaMatchesRef.current = [];
      ldaRemovalTimerRef.current = null;
    }, delay);
  };

  // insert tag label at cursor position inside textarea
  const handleTagClick = (label) => {
    // Special handling for DNC / LDA (open modals as in CommentModal)
    if (label === 'DNC') {
      setShowDropdown(false);
      setDropdownTarget(null);
      setShowDNCModal(true);
      setPendingDncTag(true);
      return;
    }
    if (label === 'LDA') {
      setShowDropdown(false);
      setDropdownTarget(null);
      setShowLDAModal(true);
      return;
    }

    // Default behavior: record the selected tag in the insertedTags pill list and close the dropdown.
    setShowDropdown(false);
    setDropdownTarget(null);
    const tagObj = tags.find(t => t.label === label) || { label, spanClass: 'bg-gray-100 text-gray-800' };
    setInsertedTags(prev => (prev.some(t => (typeof t === 'string' ? t : t.label) === tagObj.label) ? prev : [...prev, tagObj]));
  };

  // DNC selection callback: insert the chosen DNC string into the pill list (tentative)
  const handleDncSelect = (dncString) => {
    setShowDNCModal(false);
    setPendingDncTag(false);
    if (dncString) {
      // Copy styling from the parent "DNC" tag so the new DNC variant keeps the same color
      const parent = tags.find(t => (t.label || '').toUpperCase() === 'DNC');
      const tagObj = {
        label: dncString,
        title: dncString,
        spanClass: parent ? parent.spanClass : 'bg-gray-100 text-gray-800'
      };
      setInsertedTags(prev => (prev.some(t => (typeof t === 'string' ? t : t.label) === tagObj.label) ? prev : [...prev, tagObj]));
      setShowDropdown(false);
    }
  };

  // LDA selection callback: LDAModal returns the formatted LDA string (e.g. "LDA MM/DD/YY")
  const handleLdaSelect = (ldaString) => {
    setShowLDAModal(false);
    if (ldaString) {
      // Copy styling from the parent "LDA" tag so the LDA variant keeps the same color
      const parent = tags.find(t => (t.label || '').toUpperCase() === 'LDA');
      const tagObj = {
        label: ldaString,
        title: ldaString,
        spanClass: parent ? parent.spanClass : 'bg-gray-100 text-gray-800'
      };
      setInsertedTags(prev => (prev.some(t => (typeof t === 'string' ? t : t.label) === tagObj.label) ? prev : [...prev, tagObj]));
    }
  };

  // New: detect LDA-like text as user types and auto-insert/update/remove LDA tag(s)
  const handleInputChange = (e) => {
    const val = e.target.value;
    const matches = extractLdaMatches(val || '') || [];

    // If we found at least one match, cancel any pending removal and apply changes immediately.
    if (matches.length > 0) {
      if (ldaRemovalTimerRef.current) {
        clearTimeout(ldaRemovalTimerRef.current);
        ldaRemovalTimerRef.current = null;
      }
      // If matches changed from last seen, update tags immediately
      const last = lastLdaMatchesRef.current || [];
      const same = matches.length === last.length && matches.every((m, i) => m === last[i]);
      if (!same) {
        updateInsertedTagsForMatches(matches);
        lastLdaMatchesRef.current = matches;
      }
    } else {
      // No matches found currently: don't remove immediately (avoids oscillation while typing).
      // Schedule removal after short debounce so quick typing doesn't cause flicker.
      scheduleRemoveLdaTags(800);
    }
  };

  React.useEffect(() => {
    isMountedRef.current = true;
    return () => {
      isMountedRef.current = false;
      // cleanup timer on unmount
      if (ldaRemovalTimerRef.current) {
        clearTimeout(ldaRemovalTimerRef.current);
        ldaRemovalTimerRef.current = null;
      }
    };
  }, []);

  React.useEffect(() => {
    if (newCommentInputRef.current) {
      newCommentInputRef.current.value = initialComment;
      // Ensure insertedTags reflect any LDA content in the initial comment (apply immediately)
      const initialMatches = extractLdaMatches(initialComment || '') || [];
      if (initialMatches.length > 0) {
        updateInsertedTagsForMatches(initialMatches);
        lastLdaMatchesRef.current = initialMatches;
      }
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
              forceCloseDropdown={showDNCModal || showLDAModal || pendingDncTag}
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
          onChange={handleInputChange} // attach watcher for LDA detection
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

       {/* Render DNC/LDA modals alongside the component so they can return tags into insertedTags */}
       {showDNCModal && (
         <DNCModal
           isOpen={showDNCModal}
           onClose={() => { setShowDNCModal(false); setPendingDncTag(false); }}
           phone={phone}
           otherPhone={otherPhone}
           email={email}
           onSelect={handleDncSelect}
         />
       )}
       {showLDAModal && (
         <LDAModal
           isOpen={showLDAModal}
           onClose={() => setShowLDAModal(false)}
           onSelect={handleLdaSelect}
         />
       )}
     </>
   );
 }

 export default NewComment;
