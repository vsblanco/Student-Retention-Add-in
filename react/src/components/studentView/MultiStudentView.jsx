// Multi-Student Selection View Component
import React, { useMemo, useState } from 'react';
import chromeExtensionService from '../../services/chromeExtensionService';

// Helper to parse grade strings into numbers
const parseGrade = (grade) => {
  if (grade === null || grade === undefined) return null;
  if (typeof grade === 'number') return grade;
  if (typeof grade === 'string') {
    const match = grade.match(/^(\d+(?:\.\d+)?)\s*%?$/);
    if (match) return Number(match[1]);
  }
  return null;
};

// Helper to parse numeric values (for Days Out and Missing Assignments)
const parseNumber = (value) => {
  if (value === null || value === undefined) return null;
  if (typeof value === 'number') return value;
  if (typeof value === 'string') {
    const parsed = parseFloat(value);
    return isNaN(parsed) ? null : parsed;
  }
  return null;
};

// Calculate whisker plot statistics (min, Q1, median, Q3, max)
const calculateWhiskerStats = (values) => {
  if (!values || values.length === 0) return null;

  const sorted = [...values].sort((a, b) => a - b);
  const n = sorted.length;

  const min = sorted[0];
  const max = sorted[n - 1];

  // Calculate median
  const median = n % 2 === 0
    ? (sorted[n / 2 - 1] + sorted[n / 2]) / 2
    : sorted[Math.floor(n / 2)];

  // Calculate quartiles
  const q1Index = Math.floor(n * 0.25);
  const q3Index = Math.floor(n * 0.75);
  const q1 = sorted[q1Index];
  const q3 = sorted[q3Index];

  const mean = values.reduce((sum, g) => sum + g, 0) / n;

  return { min, q1, median, q3, max, mean, count: n };
};

function MultiStudentView({ students, hiddenRowCount = 0 }) {
  const [distributionType, setDistributionType] = useState('grade');

  // Calculate stats for all distribution types
  const gradeStats = useMemo(() => {
    if (!students || students.length === 0) return null;
    const validGrades = students
      .map(s => parseGrade(s.Grade))
      .filter(g => g !== null);
    if (validGrades.length === 0) return null;
    return calculateWhiskerStats(validGrades);
  }, [students]);

  const daysOutStats = useMemo(() => {
    if (!students || students.length === 0) return null;
    const validDaysOut = students
      .map(s => parseNumber(s.DaysOut))
      .filter(d => d !== null);
    if (validDaysOut.length === 0) return null;
    return calculateWhiskerStats(validDaysOut);
  }, [students]);

  const missingAssignmentsStats = useMemo(() => {
    if (!students || students.length === 0) return null;
    const validMissing = students
      .map(s => parseNumber(s.MissingAssignments))
      .filter(m => m !== null);
    if (validMissing.length === 0) return null;
    return calculateWhiskerStats(validMissing);
  }, [students]);

  const gradebookLinks = useMemo(() => {
    if (!students) return [];
    return students
      .map(s => s.Gradebook)
      .filter(link => link && typeof link === 'string' && /^https?:\/\/\S+$/i.test(link));
  }, [students]);

  // Auto-switch distribution type when current selection has no data
  React.useEffect(() => {
    let currentStats = null;

    // Get stats for current distribution type
    if (distributionType === 'grade') {
      currentStats = gradeStats;
    } else if (distributionType === 'daysOut') {
      currentStats = daysOutStats;
    } else if (distributionType === 'missingAssignments') {
      currentStats = missingAssignmentsStats;
    }

    // If current distribution has no data, switch to next available
    if (!currentStats) {
      if (daysOutStats) {
        console.log('📊 MultiStudentView: No data for', distributionType, '- switching to daysOut');
        setDistributionType('daysOut');
      } else if (missingAssignmentsStats) {
        console.log('📊 MultiStudentView: No data for', distributionType, '- switching to missingAssignments');
        setDistributionType('missingAssignments');
      } else if (gradeStats && distributionType !== 'grade') {
        console.log('📊 MultiStudentView: No data for', distributionType, '- switching to grade');
        setDistributionType('grade');
      }
    }
  }, [distributionType, gradeStats, daysOutStats, missingAssignmentsStats]);

  const openAllGradebooks = () => {
    console.log('🔵 MultiStudentView: openAllGradebooks called');
    console.log('🔵 MultiStudentView: Gradebook links count:', gradebookLinks.length);
    console.log('🔵 MultiStudentView: Gradebook links:', gradebookLinks);

    if (gradebookLinks.length === 0) {
      console.warn('⚠️ MultiStudentView: No valid gradebook links found');
      alert('No valid gradebook links found for selected students.');
      return;
    }

    // Send SRK_PING to ensure Chrome extension is active and listening
    const pingMessage = {
      type: "SRK_PING",
      timestamp: new Date().toISOString(),
      source: "excel-addin-multistudent"
    };
    console.log('🏓 MultiStudentView: Sending SRK_PING to Chrome extension:', pingMessage);
    chromeExtensionService.sendMessage(pingMessage);

    // Send links to chrome extension via chromeExtensionService
    const linksMessage = {
      type: "SRK_LINKS",
      links: gradebookLinks
    };
    console.log('🔗 MultiStudentView: Sending SRK_LINKS to Chrome extension:', linksMessage);
    chromeExtensionService.sendMessage(linksMessage);

    console.log('✅ MultiStudentView: Both messages sent to Chrome extension');
  };

  // Get current stats and config based on selected distribution
  const getCurrentDistribution = () => {
    switch (distributionType) {
      case 'grade':
        return {
          stats: gradeStats,
          title: 'Grade Distribution',
          unit: '%',
          scaleLabels: ['0%', '50%', '100%'],
          isPercentage: true,
          reverseColors: false  // Higher is better
        };
      case 'daysOut':
        return {
          stats: daysOutStats,
          title: 'Days Out Distribution',
          unit: ' days',
          scaleLabels: daysOutStats ? [
            '0',
            Math.round(daysOutStats.max / 2).toString(),
            Math.round(daysOutStats.max).toString()
          ] : ['0', '0', '0'],
          isPercentage: false,
          reverseColors: true  // Lower is better
        };
      case 'missingAssignments':
        return {
          stats: missingAssignmentsStats,
          title: 'Missing Assignments Distribution',
          unit: '',
          scaleLabels: missingAssignmentsStats ? [
            '0',
            Math.round(missingAssignmentsStats.max / 2).toString(),
            Math.round(missingAssignmentsStats.max).toString()
          ] : ['0', '0', '0'],
          isPercentage: false,
          reverseColors: true  // Lower is better
        };
      default:
        return {
          stats: gradeStats,
          title: 'Grade Distribution',
          unit: '%',
          scaleLabels: ['0%', '50%', '100%'],
          isPercentage: true,
          reverseColors: false
        };
    }
  };

  // Generalized Whisker plot component
  const WhiskerPlot = ({ stats, title, unit, scaleLabels, isPercentage, reverseColors }) => {
    if (!stats) {
      return (
        <div className="p-6 bg-gray-50 border border-gray-200 rounded-lg text-center text-gray-500">
          No {title.toLowerCase()} data available for selected students
        </div>
      );
    }

    const { min, q1, median, q3, max, mean, count } = stats;
    const range = max - min;

    // Color scheme based on whether lower or higher is better
    const minColor = reverseColors ? 'text-green-600' : 'text-red-600';
    const maxColor = reverseColors ? 'text-red-600' : 'text-green-600';
    const minStroke = reverseColors ? '#10b981' : '#ef4444';
    const maxStroke = reverseColors ? '#ef4444' : '#10b981';

    // Calculate positions (0-100 scale) - for percentage, use absolute positioning
    const getPosition = (value) => {
      if (range === 0) return 50; // All same value
      if (isPercentage) {
        // For percentages (0-100), map directly
        return value * 0.8 + 10; // 10% padding on each side
      } else {
        // For other metrics, scale relative to min/max
        return ((value - min) / range) * 80 + 10;
      }
    };

    const minPos = isPercentage ? 10 : 10;
    const q1Pos = getPosition(q1);
    const medianPos = getPosition(median);
    const q3Pos = getPosition(q3);
    const maxPos = isPercentage ? 90 : getPosition(max);
    const meanPos = getPosition(mean);

    return (
      <div className="p-6 bg-white border border-gray-200 rounded-lg">
        {/* Header with dropdown as title */}
        <div className="mb-4">
          <select
            value={distributionType}
            onChange={(e) => setDistributionType(e.target.value)}
            className="w-full text-md font-bold text-gray-800 bg-transparent border-none cursor-pointer hover:text-blue-600 focus:outline-none appearance-none"
            style={{
              backgroundImage: `url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='12' viewBox='0 0 12 12'%3E%3Cpath fill='%23666' d='M6 9L1 4h10z'/%3E%3C/svg%3E")`,
              backgroundRepeat: 'no-repeat',
              backgroundPosition: 'right center',
              paddingRight: '20px'
            }}
          >
            <option value="grade">Grade Distribution</option>
            <option value="daysOut">Days Out Distribution</option>
            <option value="missingAssignments">Missing Assignments Distribution</option>
          </select>
        </div>

        {/* Stats summary */}
        <div className="grid grid-cols-3 gap-2 mb-6 text-sm">
          <div className="text-center">
            <div className="text-gray-500">Min</div>
            <div className={`font-bold ${minColor}`}>
              {Math.round(min)}{unit}
            </div>
          </div>
          <div className="text-center">
            <div className="text-gray-500">Average</div>
            <div className="font-bold text-blue-600">
              {Math.round(mean)}{unit}
            </div>
          </div>
          <div className="text-center">
            <div className="text-gray-500">Max</div>
            <div className={`font-bold ${maxColor}`}>
              {Math.round(max)}{unit}
            </div>
          </div>
          <div className="text-center">
            <div className="text-gray-500">Q1</div>
            <div className="font-bold text-gray-700">
              {Math.round(q1)}{unit}
            </div>
          </div>
          <div className="text-center">
            <div className="text-gray-500">Median</div>
            <div className="font-bold text-gray-700">
              {Math.round(median)}{unit}
            </div>
          </div>
          <div className="text-center">
            <div className="text-gray-500">Q3</div>
            <div className="font-bold text-gray-700">
              {Math.round(q3)}{unit}
            </div>
          </div>
        </div>

        {/* Whisker plot visualization */}
        <div className="relative h-24 mb-2">
          <svg width="100%" height="100%" className="overflow-visible">
            {/* Scale reference lines */}
            <line x1="10%" y1="50%" x2="90%" y2="50%" stroke="#e5e7eb" strokeWidth="1" />

            {/* Min whisker */}
            <line
              x1={`${minPos}%`}
              y1="40%"
              x2={`${minPos}%`}
              y2="60%"
              stroke={minStroke}
              strokeWidth="2"
            />

            {/* Left whisker line */}
            <line
              x1={`${minPos}%`}
              y1="50%"
              x2={`${q1Pos}%`}
              y2="50%"
              stroke="#6b7280"
              strokeWidth="2"
            />

            {/* Box (Q1 to Q3) */}
            <rect
              x={`${q1Pos}%`}
              y="30%"
              width={`${q3Pos - q1Pos}%`}
              height="40%"
              fill="#93c5fd"
              stroke="#3b82f6"
              strokeWidth="2"
              rx="4"
            />

            {/* Median line */}
            <line
              x1={`${medianPos}%`}
              y1="30%"
              x2={`${medianPos}%`}
              y2="70%"
              stroke="#1e40af"
              strokeWidth="3"
            />

            {/* Mean marker (circle) */}
            <circle
              cx={`${meanPos}%`}
              cy="50%"
              r="4"
              fill="#3b82f6"
              stroke="#1e40af"
              strokeWidth="2"
            />

            {/* Right whisker line */}
            <line
              x1={`${q3Pos}%`}
              y1="50%"
              x2={`${maxPos}%`}
              y2="50%"
              stroke="#6b7280"
              strokeWidth="2"
            />

            {/* Max whisker */}
            <line
              x1={`${maxPos}%`}
              y1="40%"
              x2={`${maxPos}%`}
              y2="60%"
              stroke={maxStroke}
              strokeWidth="2"
            />

            {/* Scale labels */}
            <text x="10%" y="90%" fontSize="12" fill="#6b7280" textAnchor="middle">
              {scaleLabels[0]}
            </text>
            <text x="50%" y="90%" fontSize="12" fill="#6b7280" textAnchor="middle">
              {scaleLabels[1]}
            </text>
            <text x="90%" y="90%" fontSize="12" fill="#6b7280" textAnchor="middle">
              {scaleLabels[2]}
            </text>
          </svg>
        </div>

        <div className="text-xs text-gray-500 text-center mt-4">
          {count} student{count !== 1 ? 's' : ''} with data
        </div>
      </div>
    );
  };

  const currentDistribution = getCurrentDistribution();

  return (
    <div className="p-4 space-y-4">
      {/* Distribution Whisker Plot */}
      <WhiskerPlot
        stats={currentDistribution.stats}
        title={currentDistribution.title}
        unit={currentDistribution.unit}
        scaleLabels={currentDistribution.scaleLabels}
        isPercentage={currentDistribution.isPercentage}
        reverseColors={currentDistribution.reverseColors}
      />

      {/* Open All Gradebooks Button */}
      <div className="flex flex-col gap-2">
        <button
          type="button"
          onClick={openAllGradebooks}
          disabled={gradebookLinks.length === 0}
          className={`w-full py-3 px-4 rounded-lg font-semibold text-white transition-colors ${
            gradebookLinks.length > 0
              ? 'bg-blue-600 hover:bg-blue-700 active:bg-blue-800'
              : 'bg-gray-300 cursor-not-allowed'
          }`}
          title={
            gradebookLinks.length > 0
              ? `Open ${gradebookLinks.length} gradebook${gradebookLinks.length !== 1 ? 's' : ''}`
              : 'No valid gradebook links'
          }
        >
          {gradebookLinks.length > 0
            ? `Open All Gradebooks (${gradebookLinks.length})`
            : 'No Gradebook Links'}
        </button>
      </div>

      {/* Hidden row count info */}
      {hiddenRowCount > 0 && (
        <div className="text-xs text-gray-500 text-center pt-2 border-t border-gray-200">
          {hiddenRowCount} row{hiddenRowCount !== 1 ? 's' : ''} hidden by filters
        </div>
      )}
    </div>
  );
}

export default MultiStudentView;
