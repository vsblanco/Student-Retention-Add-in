import React, { useState } from "react";

const MONTHS = [
  "January",
  "February",
  "March",
  "April",
  "May",
  "June",
  "July",
  "August",
  "September",
  "October",
  "November",
  "December",
];

const Calendar = ({ onDateSelect }) => {
  const today = new Date();
  const [month, setMonth] = useState(9); // October (0-indexed)
  const [year, setYear] = useState(2025);
  const [selectedDate, setSelectedDate] = useState(null);

  // Get number of days in month
  const daysInMonth = new Date(year, month + 1, 0).getDate();
  // Get which day of week the 1st falls on (0=Sun)
  const firstDayOfWeek = new Date(year, month, 1).getDay();

  // Generate array for calendar grid
  const calendarCells = [];
  for (let i = 0; i < firstDayOfWeek; i++) {
    calendarCells.push(null); // empty cells
  }
  for (let day = 1; day <= daysInMonth; day++) {
    calendarCells.push(day);
  }

  const handleDateClick = (day) => {
    setSelectedDate(new Date(year, month, day));
  };

  const handleOkClick = () => {
    if (selectedDate && onDateSelect) {
      onDateSelect(selectedDate);
    }
  };

  const handlePrevMonth = () => {
    if (month === 0) {
      setMonth(11);
      setYear(year - 1);
    } else {
      setMonth(month - 1);
    }
    setSelectedDate(null);
  };

  const handleNextMonth = () => {
    if (month === 11) {
      setMonth(0);
      setYear(year + 1);
    } else {
      setMonth(month + 1);
    }
    setSelectedDate(null);
  };

  return (
    <div className="bg-white rounded-lg shadow-xl p-4 w-72">
      <div className="flex justify-between items-center mb-2">
        <button
          className="p-1 rounded-full hover:bg-gray-200"
          onClick={handlePrevMonth}
          aria-label="Previous Month"
        >
          &lt;
        </button>
        <h3 className="font-bold">
          {MONTHS[month]} {year}
        </h3>
        <button
          className="p-1 rounded-full hover:bg-gray-200"
          onClick={handleNextMonth}
          aria-label="Next Month"
        >
          &gt;
        </button>
      </div>
      <div className="grid grid-cols-7 gap-1 text-center text-sm">
        <div className="font-bold text-gray-500">S</div>
        <div className="font-bold text-gray-500">M</div>
        <div className="font-bold text-gray-500">T</div>
        <div className="font-bold text-gray-500">W</div>
        <div className="font-bold text-gray-500">T</div>
        <div className="font-bold text-gray-500">F</div>
        <div className="font-bold text-gray-500">S</div>
        {calendarCells.map((day, idx) =>
          day ? (
            <button
              key={day}
              className={`p-1 rounded-full hover:bg-blue-200 ${
                selectedDate &&
                selectedDate.getDate() === day &&
                selectedDate.getMonth() === month &&
                selectedDate.getFullYear() === year
                  ? "bg-blue-600 text-white"
                  : ""
              }`}
              onClick={() => handleDateClick(day)}
            >
              {day}
            </button>
          ) : (
            <div key={`empty-${idx}`}></div>
          )
        )}
      </div>
      <div className="flex justify-end mt-4">
        <button
          className="px-3 py-1 bg-blue-600 text-white rounded disabled:bg-gray-400"
          disabled={!selectedDate}
          onClick={handleOkClick}
        >
          OK
        </button>
      </div>
    </div>
  );
};

export default Calendar;
