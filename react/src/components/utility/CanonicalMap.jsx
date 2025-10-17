import { COLUMN_ALIASES, COLUMN_ALIASES_ASSIGNMENTS, COLUMN_ALIASES_HISTORY } from './ColumnMapping.jsx';

export const normalizeHeader = (value = '') =>
  String(value).toLowerCase().replace(/\s+/g, '');

const createCanonicalMap = (aliasesRecord) => {
  const map = {};
  Object.entries(aliasesRecord).forEach(([canonicalName, aliases]) => {
    const aliasList = Array.isArray(aliases) ? aliases : [aliases];
    aliasList.forEach(alias => {
      map[normalizeHeader(alias)] = canonicalName;
    });
    map[normalizeHeader(canonicalName)] = canonicalName;
  });
  return map;
};

export const canonicalHeaderMap = createCanonicalMap(COLUMN_ALIASES);
export const canonicalAssignmentsHeaderMap = createCanonicalMap(COLUMN_ALIASES_ASSIGNMENTS);
export const canonicalHistoryHeaderMap = createCanonicalMap(COLUMN_ALIASES_HISTORY);

export const getCanonicalName = (map, header) =>
  map[normalizeHeader(header)] ?? (typeof header === 'string' ? header.trim() : header);

export const getCanonicalColIdx = (
  headers,
  colName,
  canonicalMap = canonicalHeaderMap,
  aliasMap = COLUMN_ALIASES
) => {
  if (!Array.isArray(headers) || !headers.length) return -1;

  // 1) direct canonical match (preferred)
  for (let i = 0; i < headers.length; i++) {
    if (getCanonicalName(canonicalMap, headers[i]) === colName) return i;
  }

  // 2) alias / normalized header match
  const aliases = aliasMap[colName] ? [colName, ...aliasMap[colName]] : [colName];
  const normAliases = new Set(aliases.map(normalizeHeader));
  for (let i = 0; i < headers.length; i++) {
    if (normAliases.has(normalizeHeader(headers[i]))) return i;
  }
  return -1;
};
