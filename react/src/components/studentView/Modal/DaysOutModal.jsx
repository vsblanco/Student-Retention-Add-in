import React from 'react';
import Modal from '../../utility/Modal';

const DaysOutModal = ({ isOpen, onClose, daysOut }) => {
  const numeric = typeof daysOut === 'number' && !isNaN(daysOut);
  const result = numeric ? 14 - daysOut : null;
  const noDaysLeft = numeric && result < 0;
  const isDueToday = numeric && result === 0;
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
          {numeric ? (
            noDaysLeft ? (
              <span style={{ textDecoration: 'line-through' }}>{`${daysOut} Days Out`}</span>
            ) : (
              `${daysOut} Days Out`
            )
          ) : (
            'N/A Days Out'
          )}
        </h3>
        <p style={{ margin: '8px 0' }}>
          {noDaysLeft ? (
            <>The Student has <strong>no days</strong> left to submit work.</>
          ) : isDueToday ? (
            <>The student has to submit work by <strong>TODAY</strong>.</>
          ) : (
            <>
              The student has <strong>{resultText}</strong> left.
            </>
          )}
        </p>
        {numeric && result > 0 && (
          <p style={{ margin: '8px 0' }}>
            They have until <strong>{deadlineText}</strong> to submit work.
          </p>
        )}
      </div>
    </Modal>
  );
};

export default DaysOutModal;
