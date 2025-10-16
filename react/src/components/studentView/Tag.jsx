import React, { useState } from 'react';
import Modal from '../utility/Modal';
import Datepicker from 'react-tailwindcss-datepicker';
import Calendar from 'react-calendar';
import 'react-calendar/dist/Calendar.css';

function DNCModal({ isOpen, onClose, phone, otherPhone, email, onSelect }) {
  const handleSelect = value => {
    if (onSelect) onSelect(value);
    if (onClose) onClose();
  };

  return (
    <Modal isOpen={isOpen} onClose={onClose} padding="12px">
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
  const [selectedDate, setSelectedDate] = useState(null);

  const handleDayChange = (date) => {
    setSelectedDate(date);
  };

  const handleOk = () => {
    if (selectedDate && onSelect) {
      const mm = String(selectedDate.getMonth() + 1).padStart(2, '0');
      const dd = String(selectedDate.getDate()).padStart(2, '0');
      const yy = String(selectedDate.getFullYear()).slice(-2);
      onSelect(`LDA ${mm}/${dd}/${yy}`);
    }
    if (onClose) onClose();
  };

  return (
    <Modal isOpen={isOpen} onClose={onClose} padding="12px">
      <div className="bg-white rounded-lg shadow-xl p-4 w-72 mx-auto">
        <h3 className="font-bold mb-2" id="month-year">
          Select LDA Date
        </h3>
        <Calendar
          onChange={handleDayChange}
          value={selectedDate}
          className="mb-4"
        />
        <div className="flex justify-end mt-4">
          <button
            type="button"
            className="px-3 py-1 bg-blue-600 text-white rounded disabled:bg-gray-400"
            onClick={handleOk}
            id="confirm-date"
            disabled={!selectedDate}
          >
            OK
          </button>
        </div>
      </div>
    </Modal>
  );
}

export { DNCModal, LDAModal };
