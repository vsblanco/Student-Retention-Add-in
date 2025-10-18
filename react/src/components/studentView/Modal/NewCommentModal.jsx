import React, { useRef, useState } from 'react';

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
  const [status, setStatus] = useState('');

  React.useEffect(() => {
    if (newCommentInputRef.current) {
      newCommentInputRef.current.value = initialComment;
    }
  }, [initialComment, show]);

  const handleSubmit = async () => {
    const comment = newCommentInputRef.current.value.trim();
    if (!comment) {
      setStatus('Comment cannot be empty.');
      return;
    }
    setStatus('Submitting...');
    const success = await addCommentToHistory(comment);
    if (success) {
      setStatus('Comment added!');
      newCommentInputRef.current.value = '';
      if (onClose) onClose();
    } else {
      setStatus('Failed to add comment.');
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
        <div className="flex justify-between items-center mt-4">
          <span id="comment-status" className="text-sm text-green-600">{status}</span>
          <button
            id="submit-comment-button"
            className="px-6 py-2 bg-gradient-to-r from-blue-600 to-blue-500 text-white font-semibold rounded-xl shadow-md hover:from-blue-700 hover:to-blue-600 focus:outline-none focus:ring-2 focus:ring-blue-400 transition-all duration-150 disabled:bg-gray-400 disabled:from-gray-400 disabled:to-gray-400"
            onClick={handleSubmit}
            disabled={status === 'Submitting...'}
          >
            Submit
          </button>
        </div>
      </div>
    </>
  );
}

export default NewComment;
