import { HISTORY_SHEET } from '../../../../shared/constants.js';
import {
  STUDENT_NAME_ALIASES,
  STUDENT_ID_ALIASES,
  STUDENT_NUMBER_ALIASES,
  GENDER_ALIASES,
  PRIMARY_PHONE_ALIASES,
  OTHER_PHONE_ALIASES,
  STUDENT_EMAIL_ALIASES,
  PERSONAL_EMAIL_ALIASES,
  ASSIGNED_ALIASES,
  CREATED_BY_ALIASES,
  EXPECTED_START_DATE_ALIASES,
  GRADE_ALIASES,
  LAST_LDA_ALIASES,
  DAYS_OUT_ALIASES,
  GRADE_BOOK_ALIASES,
  MISSING_ASSIGNMENTS_ALIASES,
  OUTREACH_ALIASES,
  TITLE_ALIASES,
  DUE_DATE_ALIASES,
  SCORE_ALIASES,
  SUBMISSION_LINK_ALIASES,
  ASSIGNMENT_LINK_ALIASES,
  TIMESTAMP_ALIASES,
  COMMENT_ALIASES,
  TAGS_ALIASES,
  COMMENT_ID_ALIASES,
} from '../../../../shared/columnAliases.js';

export const Sheets = {
  HISTORY: HISTORY_SHEET,
  MISSING_ASSIGNMENT: "Missing Assignments"
};

// Canonical column concepts → alias arrays. All alias arrays come from
// /shared/columnAliases.js — see that file to add or change a column alias.
// Matching is case- and whitespace-insensitive (handled by normalizeHeader
// in /shared/excel-helpers.js).
//
// "ID" composes SyStudentId aliases first (preferred), then Student Number
// as fallback.
export const COLUMN_ALIASES = {
  StudentName: STUDENT_NAME_ALIASES,
  ID: [...STUDENT_ID_ALIASES, ...STUDENT_NUMBER_ALIASES],
  Gender: GENDER_ALIASES,
  Phone: PRIMARY_PHONE_ALIASES,
  CreatedBy: CREATED_BY_ALIASES,
  OtherPhone: OTHER_PHONE_ALIASES,
  StudentEmail: STUDENT_EMAIL_ALIASES,
  PersonalEmail: PERSONAL_EMAIL_ALIASES,
  Assigned: ASSIGNED_ALIASES,
  ExpectedStartDate: EXPECTED_START_DATE_ALIASES,
  Grade: GRADE_ALIASES,
  LDA: LAST_LDA_ALIASES,
  DaysOut: DAYS_OUT_ALIASES,
  Gradebook: GRADE_BOOK_ALIASES,
  MissingAssignments: MISSING_ASSIGNMENTS_ALIASES,
  Outreach: OUTREACH_ALIASES,
};

export const COLUMN_ALIASES_ASSIGNMENTS = {
  StudentName: STUDENT_NAME_ALIASES,
  title: TITLE_ALIASES,
  dueDate: DUE_DATE_ALIASES,
  score: SCORE_ALIASES,
  submissionLink: SUBMISSION_LINK_ALIASES,
  assignmentLink: ASSIGNMENT_LINK_ALIASES,
  Gradebook: GRADE_BOOK_ALIASES,
};

export const COLUMN_ALIASES_HISTORY = {
  timestamp: TIMESTAMP_ALIASES,
  comment: COMMENT_ALIASES,
  createdBy: CREATED_BY_ALIASES,
  tag: TAGS_ALIASES,
  StudentID: [...STUDENT_ID_ALIASES, ...STUDENT_NUMBER_ALIASES],
  commentID: COMMENT_ID_ALIASES,
};
