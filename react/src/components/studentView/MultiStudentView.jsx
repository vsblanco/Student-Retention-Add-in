// Multi-Student Selection View Component
import React, { useMemo, useState } from 'react';

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

function MultiStudentView({ students }) {
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

  const openAllGradebooks = () => {
    if (gradebookLinks.length === 0) {
      alert('No valid gradebook links found for selected students.');
      return;
    }

    gradebookLinks.forEach((link, index) => {
      setTimeout(() => {
        if (window.Office && window.Office.context && window.Office.context.ui && window.Office.context.ui.openBrowserWindow) {
          window.Office.context.ui.openBrowserWindow(link);
        } else {
          window.open(link, '_blank');
        }
      }, index * 100); // Stagger the opens slightly to avoid popup blockers
    });
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
          isPercentage: true
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
          isPercentage: false
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
          isPercentage: false
        };
      default:
        return {
          stats: gradeStats,
          title: 'Grade Distribution',
          unit: '%',
          scaleLabels: ['0%', '50%', '100%'],
          isPercentage: true
        };
    }
  };

  // Generalized Whisker plot component
  const WhiskerPlot = ({ stats, title, unit, scaleLabels, isPercentage }) => {
    if (!stats) {
      return (
        <div className="p-6 bg-gray-50 border border-gray-200 rounded-lg text-center text-gray-500">
          No {title.toLowerCase()} data available for selected students
        </div>
      );
    }

    const { min, q1, median, q3, max, mean, count } = stats;
    const range = max - min;

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
        {/* Header with dropdown */}
        <div className="flex items-center justify-between mb-4">
          <h3 className="text-md font-bold text-gray-800">{title}</h3>
          <select
            value={distributionType}
            onChange={(e) => setDistributionType(e.target.value)}
            className="px-3 py-1 text-sm border border-gray-300 rounded-md bg-white text-gray-700 hover:border-blue-400 focus:outline-none focus:border-blue-500 cursor-pointer"
          >
            <option value="grade">Grade</option>
            <option value="daysOut">Days Out</option>
            <option value="missingAssignments">Missing Assignments</option>
          </select>
        </div>

        {/* Stats summary */}
        <div className="grid grid-cols-3 gap-2 mb-6 text-sm">
          <div className="text-center">
            <div className="text-gray-500">Min</div>
            <div className="font-bold text-red-600">
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
            <div className="font-bold text-green-600">
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
              stroke="#ef4444"
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
              stroke="#10b981"
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

        {gradebookLinks.length > 0 && (
          <div className="text-xs text-gray-500 text-center">
            Opens {gradebookLinks.length} gradebook link{gradebookLinks.length !== 1 ? 's' : ''} in new window{gradebookLinks.length !== 1 ? 's' : ''}
          </div>
        )}
      </div>

      {/* Student count info */}
      <div className="text-sm text-gray-600 text-center pt-2 border-t border-gray-200">
        {students.length} student{students.length !== 1 ? 's' : ''} selected
      </div>
    </div>
  );
}

export default MultiStudentView;
