// Conversion.jsx

// Convert Excel serial date to JavaScript Date object
// Excel uses a date system where 1 corresponds to 1900-01-01
// This function converts an Excel serial date to a JavaScript Date object

export function formatExcelDate(serial, format = "default") {
  const excelEpoch = new Date(Date.UTC(1899, 11, 30));
  const days = Math.floor(Number(serial));
  const ms = Math.round((Number(serial) - days) * 24 * 60 * 60 * 1000);
  const date = new Date(excelEpoch.getTime() + days * 86400000 + ms);

  if (isNaN(date.getTime())) return "N/A";

  if (format === "long") {
    // "Month Day, Year"
    return date.toLocaleDateString(undefined, {
      year: 'numeric',
      month: 'long',
      day: 'numeric'
    });
  }

  // Default: mm/dd/yy H:MM AM/PM
  const mm = String(date.getMonth() + 1).padStart(2, '0');
  const dd = String(date.getDate()).padStart(2, '0');
  const yy = String(date.getFullYear()).slice(-2);
  let hours = date.getHours();
  const minutes = String(date.getMinutes()).padStart(2, '0');
  const ampm = hours >= 12 ? 'PM' : 'AM';
  hours = hours % 12;
  hours = hours ? hours : 12; // 0 should be 12

  return `${mm}/${dd}/${yy} ${hours}:${minutes} ${ampm}`;
}

// Converts "First Last" <-> "Last, First"
export function formatName(name) {
  if (!name || typeof name !== 'string') return name;

  // Detect "Last, First"
  if (name.includes(',')) {
    const [last, first] = name.split(',').map(s => s.trim());
    if (first && last) {
      return `${first} ${last}`;
    }
  } else {
    // Assume "First Last"
    const parts = name.trim().split(/\s+/);
    if (parts.length === 2) {
      return `${parts[1]}, ${parts[0]}`;
    }
  }
  // If format not recognized, return as is
  return name;
}

// Usage example:
// const date = formatExcelDate(44561); // Converts Excel date serial to JavaScript Date
// console.log(date); // Outputs: "10/07/25 12:00 AM" (depending on the input serial)