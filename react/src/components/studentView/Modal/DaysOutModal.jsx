import React from 'react';
import Modal from '../../utility/Modal';

const DaysOutModal = ({ isOpen, onClose, daysOut }) => {
  const numeric = typeof daysOut === 'number' && !isNaN(daysOut);
  const result = numeric ? 14 - daysOut : null;
  const resultText = result === null ? 'N/A' : `${result} ${result === 1 ? 'day' : 'days'}`;

  // compute deadline: today + result days (result may be 0 or negative)
  const deadlineText =
    result === null
      ? 'N/A'
      : (() => {
          const d = new Date();
          d.setDate(d.getDate() + result);
          return d.toLocaleDateString('en-US', {
            month: 'long',
            day: 'numeric',
            year: 'numeric'
          });
        })();

  return (
    <Modal
      isOpen={isOpen}
      onClose={onClose}
      padding="16px"
      borderRadius="12px"
      style={{ maxWidth: 420, textAlign: 'center' }}
    >
      <div style={{ width: '100%' }}>
        <h3 style={{ margin: 0, marginBottom: 8, fontSize: 18, fontWeight: 700 }}>
          {numeric ? `${daysOut} Days Out` : 'N/A Days Out'}
        </h3>
        <p style={{ margin: '8px 0' }}>
          The student has <strong>{resultText}</strong> left.
        </p>
        <p style={{ margin: '8px 0' }}>
          They have until <strong>{deadlineText}</strong> to submit work.
        </p>
      </div>
    </Modal>
  );
};

export default DaysOutModal;
