import React, { useState, useEffect, useRef } from 'react';
import Modal from '../utility/Modal';
import InsertTagButton from './InsertTagButton';

function CommentModal({
  isOpen,
  onClose,
  entry,
  onEditComment,
  COMMENT_TAGS,
  findTagInfo,
  hasQuoteTag,
  quoteText,
  beforeQuote,
  afterQuote,
  formatExcelDate
}) {
  // Modal state and logic moved from Comment.jsx
  const [modalMode, setModalMode] = useState('view');
  const [modalComment, setModalComment] = useState(entry.comment || "");
  const [modalTagContainer, setModalTagContainer] = useState({});
  const [confirmDelete, setConfirmDelete] = useState(false);
  const [showTagDropdown, setShowTagDropdown] = useState(false);
  const [dropdownTarget, setDropdownTarget] = useState(null);

  useEffect(() => {
    setModalComment(entry.comment || "");
    let tags = entry.tag
      ? entry.tag.split(',').map(t => t.trim()).filter(Boolean)
      : [];
    tags = tags.filter(t => t !== "Comment");
    const tagObj = {};
    tags.forEach(t => { tagObj[t] = true; });
    setModalTagContainer(tagObj);
    setModalMode('view');
  }, [entry.comment, entry.tag, isOpen]);

  useEffect(() => {
    setConfirmDelete(false);
  }, [isOpen, modalMode]);

  const handleSaveComment = async () => {
    if (onEditComment) {
      const newTagString = Object.keys(modalTagContainer).join(', ');
      const success = await onEditComment(
        { ...entry, tag: newTagString },
        modalComment
      );
      if (success) onClose();
    } else {
      onClose();
    }
  };

  const handleDeleteComment = async () => {
    if (window.confirm("Are you sure you want to delete this comment?")) {
      if (typeof onEditComment === 'function') {
        await onEditComment({ ...entry, deleted: true }, null);
      }
      onClose();
    }
  };

  const insertTagButtonTags = COMMENT_TAGS.map(tag => ({
    label: tag.label,
    spanClass: tag.tagClass,
    title: tag.label
  }));

  const handleInsertTag = tagLabel => {
    setModalTagContainer(prev => ({
      ...prev,
      [tagLabel]: true
    }));
    setShowTagDropdown(false);
  };

  const handleRemoveTag = tagLabel => {
    setModalTagContainer(prev => {
      const newObj = { ...prev };
      delete newObj[tagLabel];
      return newObj;
    });
  };

  // --- Modal tag pills for edit mode ---
  const modalTagPillStyle = "px-1 py-0.5 text-xs font-semibold rounded-full";
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
              color: textColor,
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

  // --- Modal tag pills for view mode ---
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
              color: textColor,
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
      {modalTagViewPills}
      <div style={{ marginTop: 12, marginBottom: 12 }}>
        <div
          style={{
            fontSize: '1rem',
            color: '#222',
            whiteSpace: 'pre-wrap',
            wordBreak: 'break-word',
            minHeight: 60,
            paddingLeft: 16,
            paddingTop: 6,
            borderRadius: 8,
            background: 'transparent',
            border: '1px solid #d5d5d5',
          }}
        >
          {hasQuoteTag && quoteText ? (
            <>
              {beforeQuote}
              <blockquote className="relative bg-blue-50 border-l-4 border-blue-200 pl-6 pr-2 py-3 mb-2 rounded" style={{ marginLeft: 0 }}>
                <span className="absolute left-2 top-2 text-4xl text-blue-200 leading-none select-none" aria-hidden="true">“</span>
                <span className="text-base text-blue-900 font-serif">
                  {quoteText}
                </span>
                <span className="absolute right-2 bottom-2 text-4xl text-blue-200 leading-none select-none" aria-hidden="true">”</span>
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
      {modalTagPills}
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
    <Modal
      isOpen={isOpen}
      onClose={onClose}
      padding="12px"
      style={{ minHeight: 200, width: '100%' }}
    >
      {modalMode === 'view' ? modalContentView : modalContentEdit}
    </Modal>
  );
}

export default CommentModal;
