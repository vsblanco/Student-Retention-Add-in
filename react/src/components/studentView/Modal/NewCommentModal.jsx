import React, { useRef, useState } from 'react';
import { useRef as useRefHook, useEffect as useEffectHook } from 'react';

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
    // continue submission in background; guard state updates if unmounted
    try {
      const success = await addCommentToHistory(comment);
      if (success) {
        // clear input if still mounted
        if (isMountedRef.current && newCommentInputRef.current) {
          newCommentInputRef.current.value = '';
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
