import React, { useState } from 'react';
import Modal from '../utility/Modal';
import Datepicker from 'react-tailwindcss-datepicker';
import Calendar from '../utility/Calendar';

function DNCModal({ isOpen, onClose, phone, otherPhone, email, onSelect }) {
  const handleSelect = value => {
    if (onSelect) onSelect(value);
    if (onClose) onClose();
  };

  return (
    <Modal
      isOpen={isOpen}
      onClose={onClose}
      padding="12px"
      overlayStyle={{
        background: "rgba(0,0,0,0.35)", // darker background
        backdropFilter: "blur(5px) saturate(140%)",
        WebkitBackdropFilter: "blur(5px) saturate(140%)"
      }}
    >
      <div className="w-full h-full flex flex-col rounded-lg p-4">
        <h3 className="text-lg font-medium text-gray-900 mb-2">Select DNC Type</h3>
        <div id="dnc-options-container" className="space-y-2">
          <button
            className="block w-full text-left px-4 py-2 text-sm text-gray-700 hover:bg-gray-100 rounded-md border"
            onClick={() => handleSelect('DNC - Phone')}
          >
            DNC - <span className="font-bold">Phone:</span> {phone}
          </button>
          <button
            className="block w-full text-left px-4 py-2 text-sm text-gray-700 hover:bg-gray-100 rounded-md border"
            onClick={() => handleSelect('DNC - Other Phone')}
          >
            DNC - <span className="font-bold">Other Phone:</span> {otherPhone}
          </button>
          <button
            className="block w-full text-left px-4 py-2 text-sm text-gray-700 hover:bg-gray-100 rounded-md border"
            onClick={() => handleSelect('DNC - Email')}
          >
            DNC - <span className="font-bold">Email:</span> {email}
          </button>
          <button
            className="block w-full text-left px-4 py-2 text-sm text-gray-700 hover:bg-gray-100 rounded-md border"
            onClick={() => handleSelect('DNC')}
          >
            DNC - <span className="font-bold">All:</span> All Contact Methods
          </button>
        </div>
      </div>
    </Modal>
  );
}

function LDAModal({ isOpen, onClose, onSelect }) {
  // Accepts JS Date object from Calendar, converts to MM/DD/YY
  const handleDateSelect = (date) => {
    if (date && onSelect) {
      const mm = String(date.getMonth() + 1).padStart(2, '0');
      const dd = String(date.getDate()).padStart(2, '0');
      const yy = String(date.getFullYear()).slice(-2);
      onSelect(`LDA ${mm}/${dd}/${yy}`);
    }
    if (onClose) onClose();
  };

  return (
    <Modal isOpen={isOpen} onClose={onClose} padding="12px">
      <div className="rounded-lg w-auto mx-auto">
        <h3 className="font-bold mb-2" id="month-year">
          Select LDA Date
        </h3>
        <Calendar onDateSelect={handleDateSelect} />
        {/* OK button removed */}
      </div>
    </Modal>
  );
}

export { DNCModal, LDAModal };
