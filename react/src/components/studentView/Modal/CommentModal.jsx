import React, { useState, useEffect, useRef } from 'react';
import Modal from '../../utility/Modal.jsx';
import InsertTagButton from '../Parts/InsertTagButton.jsx';
import { highlightLdaKeywords } from '../Parts/Comment.jsx';
import { DNCModal, LDAModal } from '../Tag.jsx';
import { Pencil, ArrowLeft, Check, Trash2, Clipboard } from 'lucide-react';
import { toast } from 'react-toastify';
import { deleteComment, editComment } from '../../utility/EditStudentHistory.jsx';

function CommentModal({
  isOpen,
  onClose,
  // called by modal after a successful delete to allow immediate local removal/animation
  onDeleted,
  onSaved,
  entry,
  COMMENT_TAGS,
  findTagInfo,
  hasQuoteTag,
  quoteText,
  beforeQuote,
  afterQuote,
  formatExcelDate,
  quoteStyles = {},
}) {
  // Modal state and logic moved from Comment.jsx
  const [modalMode, setModalMode] = useState('view');
  const [modalComment, setModalComment] = useState(entry.comment || "");
  // saved comment shown in view mode (finalized when Update is clicked)
  const [modalSavedComment, setModalSavedComment] = useState(entry.comment || "");
  const [modalTagContainer, setModalTagContainer] = useState({});
  // Temporary edit container used only while in edit mode.
  const [editTagContainer, setEditTagContainer] = useState({});
  const [confirmDelete, setConfirmDelete] = useState(false);
  const [showTagDropdown, setShowTagDropdown] = useState(false);
  const [dropdownTarget, setDropdownTarget] = useState(null);
  const [showDNCModal, setShowDNCModal] = useState(false);
  const [pendingDncTag, setPendingDncTag] = useState(null);
  const [showLDAModal, setShowLDAModal] = useState(false);

  // Ensure modal has class strings for quote elements (use passed-in styles with fallbacks)
  const blockClass = quoteStyles.block || "relative bg-blue-50 border-l-4 border-blue-200 pl-6 pr-2 py-3 mb-2 rounded";
  const textClass = quoteStyles.text || "text-base text-blue-900 font-serif";
  const markLeftClass = quoteStyles.markLeft || "absolute left-2 top-2 text-4xl text-blue-200 leading-none select-none";
  const markRightClass = quoteStyles.markRight || "absolute right-2 bottom-2 text-4xl text-blue-200 leading-none select-none";

  useEffect(() => {
    setModalComment(entry.comment || "");
    setModalSavedComment(entry.comment || "");
    let tags = entry.tag
      ? entry.tag.split(',').map(t => t.trim()).filter(Boolean)
      : [];
    tags = tags.filter(t => t !== "Comment");
    const tagObj = {};
    tags.forEach(t => { tagObj[t] = true; });
    // initialize both saved and edit containers from the entry
    setModalTagContainer(tagObj);
    setEditTagContainer(tagObj);
    setModalMode('view');
  }, [entry.comment, entry.tag, isOpen]);

  // When entering edit mode, ensure edit container is a copy of the saved tags
  useEffect(() => {
    if (modalMode === 'edit') {
      setEditTagContainer(modalTagContainer || {});
      // start editing from the saved comment
      setModalComment(modalSavedComment || "");
    }
  }, [modalMode, modalTagContainer, modalSavedComment]);

  // When returning to view mode without saving, discard tentative edits
  useEffect(() => {
    if (modalMode === 'view') {
      setEditTagContainer(modalTagContainer || {});
      // ensure view shows the saved comment
      setModalComment(modalSavedComment || "");
    }
  }, [modalMode, modalTagContainer, modalSavedComment]);

  useEffect(() => {
    setConfirmDelete(false);
  }, [isOpen, modalMode]);

  const handleSaveComment = async () => {
    // Finalize edit tags into saved tags and switch back to view mode (do not close)
    setModalTagContainer(editTagContainer);
    // finalize the edited text so view mode shows it
    setModalSavedComment(modalComment);

    // Log the new comment entry object when Update is pressed
    const newCommentEntry = {
      commentid: entry.commentid,
      comment: modalComment,
      tag: Object.keys(editTagContainer).join(', '),
      createdby: entry.createdby,
      timestamp: entry.timestamp
    };
    console.log('Updated comment entry:', newCommentEntry);

    // send update to shared edit implementation
    try {
      await editComment(entry.commentid, newCommentEntry);
      onSaved(entry.commentid);
    } catch (err) {
      try { console.error('Edit comment failed:', err); } catch (_) {}
      return;
    }
    // Switch back to view mode so the newly saved tags appear in the view
    setModalMode('view');
  };

  const handleDeleteComment = () => {
    console.log(`Deleting comment with ID: ${entry.commentid}`);
    onDeleted(entry.commentid);
  };

  const insertTagButtonTags = COMMENT_TAGS.map(tag => ({
    label: tag.label,
    spanClass: tag.tagClass,
    title: tag.label
  }));

  const handleInsertTag = tagLabel => {
    if (tagLabel === "DNC") {
      setShowTagDropdown(false); // Close dropdown before opening DNC modal
      setShowDNCModal(true);
      setPendingDncTag(true); // mark that we're waiting for DNCModal
      return;
    }
    if (tagLabel === "LDA") {
      setShowTagDropdown(false);
      setShowLDAModal(true);
      return;
    }
    // Only mutate the edit container so changes are tentative until Update is clicked
    setEditTagContainer(prev => ({
      ...prev,
      [tagLabel]: true
    }));
    setShowTagDropdown(false);
  };

  const handleRemoveTag = tagLabel => {
    // Remove from the edit container only
    setEditTagContainer(prev => {
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
      {Object.keys(editTagContainer).map((tagLabel, idx) => {
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
                color: textColor || '#888', // use tag text color if available
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
        border: '2px solid #cfcfcf',
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

  // helper to compare tag containers
  const tagKeys = obj => Object.keys(obj || {}).sort().join(', ');
  const noTagChanges = tagKeys(editTagContainer) === tagKeys(modalTagContainer);

  // --- Modal content ---
  const [clipboardHover, setClipboardHover] = useState(false);

  const modalContentView = (
    <div
      style={{ width: '100%' }}
      onClick={e => {
        // Only close if NOT clicking edit or clipboard button
        if (
          e.target.closest('button[aria-label="Edit"]') ||
          e.target.closest('button[aria-label="Copy comment"]')
        ) {
          return;
        }
        if (onClose) onClose();
      }}
    >
      {Object.keys(modalTagContainer).length > 0 && modalTagViewPills}
      <div
        style={{ marginTop: 12, marginBottom: 12, position: 'relative' }}
        onMouseEnter={() => setClipboardHover(true)}
        onMouseLeave={() => setClipboardHover(false)}
      >
        {/* Clipboard icon */}
        <button
          type="button"
          aria-label="Copy comment"
          title="Copy comment"
          style={{
            position: 'absolute',
            top: 8,
            right: 8,
            background: clipboardHover ? '#e0e0e0' : '#e0f2fe',
            border: 'none',
            borderRadius: 6,
            width: 32,
            height: 32,
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            cursor: 'pointer',
            boxShadow: clipboardHover
              ? '0 2px 8px rgba(2,132,199,0.12)'
              : '0 1px 4px rgba(0,0,0,0.08)',
            transition: 'background 0.15s, box-shadow 0.15s, opacity 0.15s',
            opacity: clipboardHover ? 0.5 : 0,
            zIndex: 2
          }}
          onClick={() => {
            let textToCopy;
            if (hasQuoteTag && quoteText) {
              textToCopy = `${beforeQuote || ''}${quoteText}${afterQuote || ''}`;
            } else {
              textToCopy = modalSavedComment || '';
            }
            navigator.clipboard.writeText(textToCopy);
          }}
        >
          <Clipboard size={18} />
        </button>
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
            border: '1px solid #cfcfcf',
            position: 'relative'
          }}
        >
          {hasQuoteTag && quoteText ? (
            <>
              {beforeQuote}
              <blockquote className={blockClass} style={{ marginLeft: 0 }}>
                <span className={markLeftClass} aria-hidden="true">“</span>
                <span className={textClass}>
                  {quoteText}
                </span>
                <span className={markRightClass} aria-hidden="true">”</span>
              </blockquote>
              {afterQuote}
            </>
          ) : (
            highlightLdaKeywords(modalSavedComment)
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
          aria-label="Edit"
          style={{
            width: 36,
            height: 36,
            borderRadius: '50%',
            background: '#2563eb',
            color: 'white',
            border: 'none',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            cursor: 'pointer',
            boxShadow: '0 1px 4px rgba(0,0,0,0.08)',
            transition: 'background 0.15s, box-shadow 0.15s'
          }}
          onMouseEnter={e => {
            e.currentTarget.style.background = '#1e40af';
            e.currentTarget.style.boxShadow = '0 2px 8px rgba(37,99,235,0.15)';
          }}
          onMouseLeave={e => {
            e.currentTarget.style.background = '#2563eb';
            e.currentTarget.style.boxShadow = '0 1px 4px rgba(0,0,0,0.08)';
          }}
        >
          <Pencil size={20} />
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
        forceCloseDropdown={showDNCModal}
      />
      {Object.keys(editTagContainer).length > 0 && modalTagPills}
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
          border: '1px solid #cfcfcf',
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
        <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
          <button
            type="button"
            onClick={() => setConfirmDelete(confirmDelete ? false : true)}
            style={{
              width: 36,
              height: 36,
              borderRadius: '50%',
              background: confirmDelete ? '#fee2e2' : '#ef4444',
              color: confirmDelete ? '#b91c1c' : 'white',
              border: 'none',
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'center',
              cursor: 'pointer',
              boxShadow: '0 1px 4px rgba(0,0,0,0.08)',
              transition: 'background 0.15s, box-shadow 0.15s'
            }}
            disabled={false}
            aria-label={confirmDelete ? "Cancel Delete" : "Delete"}
            title={confirmDelete ? "Cancel Delete" : "Delete"}
            onMouseEnter={e => {
              e.currentTarget.style.background = confirmDelete ? '#fecaca' : '#dc2626';
              e.currentTarget.style.boxShadow = confirmDelete
                ? '0 2px 8px rgba(220,38,38,0.10)'
                : '0 2px 8px rgba(220,38,38,0.15)';
            }}
            onMouseLeave={e => {
              e.currentTarget.style.background = confirmDelete ? '#fee2e2' : '#ef4444';
              e.currentTarget.style.boxShadow = '0 1px 4px rgba(0,0,0,0.08)';
            }}
          >
            <Trash2 size={20} />
          </button>

          {/* Moved CommentID to sit to the right of the delete button */}
          <span style={{ fontSize: 12, color: '#6b7280', userSelect: 'text' }}>
            Comment ID: {entry.commentid}
          </span>
        </div>

        <div style={{ display: 'flex', gap: 8 }}>
          <button
            type="button"
            onClick={() => {
              // Discard edits and return to view mode
              setEditTagContainer(modalTagContainer || {});
              setModalMode('view');
            }}
            style={{
              width: 36,
              height: 36,
              borderRadius: '50%',
              background: '#e5e7eb',
              color: '#222',
              border: 'none',
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'center',
              cursor: 'pointer',
              boxShadow: '0 1px 4px rgba(0,0,0,0.08)',
              transition: 'background 0.15s, box-shadow 0.15s'
            }}
            aria-label="Back"
            title="Back"
            onMouseEnter={e => {
              e.currentTarget.style.background = '#d1d5db';
              e.currentTarget.style.boxShadow = '0 2px 8px rgba(31,41,55,0.12)';
            }}
            onMouseLeave={e => {
              e.currentTarget.style.background = '#e5e7eb';
              e.currentTarget.style.boxShadow = '0 1px 4px rgba(0,0,0,0.08)';
            }}
          >
            <ArrowLeft size={20} />
          </button>
          <button
            type="button"
            onClick={confirmDelete ? handleDeleteComment : handleSaveComment}
            disabled={
              confirmDelete
                ? false
                : (modalComment === (entry.comment || "") && noTagChanges)
            }
            style={{
              width: 36,
              height: 36,
              borderRadius: '50%',
              background: confirmDelete
                ? '#ef4444'
                : (modalComment === (entry.comment || "") && noTagChanges)
                  ? '#d1d5db'
                  : '#2563eb',
              color: confirmDelete
                ? 'white'
                : (modalComment === (entry.comment || "") && noTagChanges)
                  ? '#888'
                  : 'white',
              fontWeight: 500,
              border: 'none',
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'center',
              cursor: confirmDelete
                ? 'pointer'
                : (modalComment === (entry.comment || "") && noTagChanges ? 'not-allowed' : 'pointer'),
              opacity: confirmDelete ? 1 : (modalComment === (entry.comment || "") && noTagChanges ? 0.7 : 1),
              transition: 'background 0.15s, box-shadow 0.15s'
            }}
            aria-label={confirmDelete ? 'Confirm Delete' : 'Update'}
            title={confirmDelete ? 'Confirm Delete' : 'Update'}
            onMouseEnter={e => {
              e.currentTarget.style.background = confirmDelete ? '#dc2626' : '#1e40af';
              e.currentTarget.style.boxShadow = confirmDelete
                ? '0 2px 8px rgba(220,38,38,0.15)'
                : '0 2px 8px rgba(37,99,235,0.15)';
            }}
            onMouseLeave={e => {
              e.currentTarget.style.background = confirmDelete
                ? '#ef4444'
                : (modalComment === (entry.comment || "") && noTagChanges)
                ? '#d1d5db'
                : '#2563eb';
              e.currentTarget.style.boxShadow = '0 1px 4px rgba(0,0,0,0.08)';
            }}
          >
            <Check size={20} />
          </button>
         </div>
       </div>
     </div>
   );

  // DNCModal callback
  const handleDncSelect = (dncString) => {
    setShowDNCModal(false);
    setPendingDncTag(false);
    if (dncString) {
      // Add to the edit container so the tag is tentative and will be discarded on cancel
      setEditTagContainer(prev => ({
        ...prev,
        [dncString]: true
      }));
      setShowTagDropdown(false);
    }
  };

  return (
    <>
      <Modal
        isOpen={isOpen}
        onClose={onClose}
        padding="12px"
        style={{ minHeight: 200, width: '100%' }}
      >
        {modalMode === 'view' ? modalContentView : modalContentEdit}
      </Modal>
      {showDNCModal && (
        <DNCModal
          isOpen={showDNCModal}
          onClose={() => { setShowDNCModal(false); setPendingDncTag(false); }}
          phone={entry.phone}
          otherPhone={entry.otherPhone}
          email={entry.email}
          onSelect={handleDncSelect}
        />
      )}
      {showLDAModal && (
        <LDAModal
          isOpen={showLDAModal}
          onClose={() => setShowLDAModal(false)}
          keywords={entry.ldaKeywords || []}
          onSelect={kw => {
            setShowLDAModal(false);
            if (kw) {
              // Add to the edit container so the tag is tentative and will be discarded on cancel
              setEditTagContainer(prev => ({
                ...prev,
                [kw]: true
              }));
            }
          }}
        />
      )}
    </>
  );
}

export default CommentModal;
