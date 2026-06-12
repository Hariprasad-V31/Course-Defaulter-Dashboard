import { Fragment, useEffect, useMemo, useRef, useState } from 'react'
import * as XLSX from 'xlsx'
import {
  Alert,
  Box,
  Button,
  Chip,
  Container,
  Dialog,
  DialogContent,
  DialogTitle,
  Divider,
  FormControl,
  Grid,
  IconButton,
  InputAdornment,
  InputLabel,
  MenuItem,
  Paper,
  Select,
  Stack,
  Table,
  TableBody,
  TableCell,
  TableContainer,
  TableHead,
  TableRow,
  TextField,
  Typography,
  createTheme,
  ThemeProvider,
} from '@mui/material'
import AnalyticsRoundedIcon from '@mui/icons-material/AnalyticsRounded'
import BusinessCenterRoundedIcon from '@mui/icons-material/BusinessCenterRounded'
import CheckCircleRoundedIcon from '@mui/icons-material/CheckCircleRounded'
import FileDownloadRoundedIcon from '@mui/icons-material/FileDownloadRounded'
import GroupsRoundedIcon from '@mui/icons-material/GroupsRounded'
import InsightsRoundedIcon from '@mui/icons-material/InsightsRounded'
import PlayArrowRoundedIcon from '@mui/icons-material/PlayArrowRounded'
import ReportProblemRoundedIcon from '@mui/icons-material/ReportProblemRounded'
import RestartAltRoundedIcon from '@mui/icons-material/RestartAltRounded'
import CloseRoundedIcon from '@mui/icons-material/CloseRounded'
import SearchRoundedIcon from '@mui/icons-material/SearchRounded'
import UploadFileRoundedIcon from '@mui/icons-material/UploadFileRounded'
import './App.css'

const COURSES = [
  { key: 'secuware', label: 'Secuware' },
  { key: 'learningCoach', label: 'Learning Coach' },
  { key: 'e2', label: 'E2' },
  { key: 'mandatory', label: 'Mandatory' },
]

const ACCEPTED_EXTENSIONS = ['.xlsx', '.xls', '.xlsm']

const HEADER_CANDIDATES = {
  employeeId: [
    'employeeid',
    'employeenumber',
    'employeeno',
    'employee',
    'employeecode',
    'colleagueid',
    'empid',
    'empno',
    'empnumber',
    'id',
    'userid',
    'personid',
    'personnumber',
    'personno',
    'associateid',
    'associatecode',
    'resourceid',
    'personnelnumber',
    'personnelno',
  ],
  employeeName: ['employeename', 'name', 'fullname', 'colleaguename'],
  portfolio: ['portfolio', 'portfolioname', 'vertical', 'departmentportfolio'],
  statusGeneric: [
    'iscompliant',
    'finalstatus',
    'interviewstatus',
    'compliance',
    'compliant',
    'completionstatus',
    'coursecompletionstatus',
    'trainingstatus',
    'learningstatus',
    'mandatorystatus',
    'status',
  ],
  mandatoryStatus: [
    'iscompliant',
    'compliant',
    'compliancestatus',
    'coursecompletionstatus',
    'completionstatus',
    'mandatorystatus',
    'trainingstatus',
    'status',
  ],
  e2Status: [
    'heldlevel',
    'currentlevel',
    'acquiredlevel',
    'highestlevelacquired',
    'highestlevelheld',
    'highestcompetencylevel',
    'currentcompetencylevel',
    'competencyacquired',
    'competencyheld',
    'competencylevel',
    'proficiencylevel',
    'level',
    'scoreband',
    'rating',
    'status',
  ],
}

const defaultUploadState = {
  secuware: null,
  learningCoach: null,
  e2: null,
  mandatory: null,
}

const defaultHcUploadState = {
  previous: null,
  current: null,
}

const EXCLUDED_PORTFOLIOS = new Set(['Sourcing', 'Selling Experience'])
const MAX_EMPLOYEE_RESULTS = 200

const MAJOR_PORTFOLIOS = [
  'Foods',
  'Online',
  'Infra',
  'Logistics',
  'Retail',
  'International',
  'Engineering Office',
  'C&H',
]

function sortPortfoliosByPriority(portfolios) {
  const major = []
  const other = []

  portfolios.forEach((portfolio) => {
    if (MAJOR_PORTFOLIOS.includes(portfolio)) {
      major.push(portfolio)
    } else {
      other.push(portfolio)
    }
  })

  major.sort(
    (a, b) => MAJOR_PORTFOLIOS.indexOf(a) - MAJOR_PORTFOLIOS.indexOf(b),
  )
  other.sort((a, b) => a.localeCompare(b))

  return [...major, ...other]
}

const theme = createTheme({
  palette: {
    mode: 'light',
    primary: {
      main: '#0f766e',
    },
    secondary: {
      main: '#1d4ed8',
    },
    success: {
      main: '#15803d',
    },
    error: {
      main: '#b91c1c',
    },
    background: {
      default: 'transparent',
      paper: 'rgba(255, 255, 255, 0.78)',
    },
  },
  shape: {
    borderRadius: 14,
  },
  typography: {
    fontFamily: 'var(--font-sans)',
    h3: {
      fontSize: 'var(--fs-h1)',
      fontWeight: 'var(--fw-bold)',
      lineHeight: 'var(--lh-h1)',
      letterSpacing: 0,
    },
    h6: {
      fontSize: 'var(--fs-h3)',
      fontWeight: 'var(--fw-semibold)',
      lineHeight: 'var(--lh-h3)',
      letterSpacing: 0,
    },
    body1: {
      fontSize: 'var(--fs-body)',
      fontWeight: 'var(--fw-regular)',
      lineHeight: 'var(--lh-body)',
    },
    body2: {
      fontSize: 'var(--fs-small)',
      fontWeight: 'var(--fw-regular)',
      lineHeight: 'var(--lh-small)',
    },
    caption: {
      fontSize: 'var(--fs-caption)',
      fontWeight: 'var(--fw-medium)',
      lineHeight: 'var(--lh-caption)',
    },
    button: {
      fontSize: 'var(--fs-small)',
      fontWeight: 'var(--fw-semibold)',
      lineHeight: 'var(--lh-small)',
      letterSpacing: 0,
      textTransform: 'none',
    },
  },
  components: {
    MuiPaper: {
      styleOverrides: {
        root: {
          backgroundImage: 'none',
        },
      },
    },
    MuiButton: {
      styleOverrides: {
        root: {
          borderRadius: 10,
          minHeight: 42,
        },
        outlined: {
          backgroundColor: 'rgba(255, 255, 255, 0.72)',
          borderColor: '#cbdde2',
        },
        contained: {
          boxShadow: '0 12px 28px rgba(15, 118, 110, 0.22)',
        },
      },
    },
    MuiTableCell: {
      styleOverrides: {
        head: {
          color: '#0f172a',
          fontWeight: 700,
          backgroundColor: '#eef7f5',
          borderBottomColor: '#cfe3df',
        },
        body: {
          borderBottomColor: '#e6eef1',
        },
      },
    },
  },
})

function normalizeKey(value) {
  return String(value ?? '')
    .toLowerCase()
    .replace(/[^a-z0-9]/g, '')
}

function normalizeCellValue(value) {
  return String(value ?? '').trim()
}

function normalizePortfolioName(value) {
  const portfolio = normalizeCellValue(value)

  if (!portfolio) {
    return 'Unmapped'
  }

  // Compact, lowercase, alphanumeric-only key for matching variants
  // (handles "C&H", " - ", extra spaces, casing, etc.)
  const compact = portfolio
    .toLowerCase()
    .replace(/&/g, 'and')
    .replace(/[^a-z0-9]+/g, '')

  // C&H Commercial Trading + C&H E2E Planning -> "C&H Commercial Trading/E2E Planning"
  if (
    compact === 'candhcommercialtrading' ||
    compact === 'candhe2eplanning' ||
    compact === 'chcommercialtrading' ||
    compact === 'che2eplanning'
  ) {
    return 'C&H Commercial Trading/E2E Planning'
  }

  // Retail + Retail - RTS -> "Retail"
  if (compact === 'retail' || compact === 'retailrts') {
    return 'Retail'
  }

  // BPS + BPS-DOS -> "BPS"
  if (compact === 'bps' || compact === 'bpsdos') {
    return 'BPS'
  }

  // Roll these portfolios into "Others"
  const othersKeys = new Set([
    'dws',
    'serviceassurancepractice',
    'cis',
    'certificatemanagementassessmentphase',
    'certificatemgmtassessmentphase',
    'cybersecurity',
    'dwa',
    'operatingsecurelyprogram',
    'customer',
    'customeroando', // "Customer O&O"
  ])
  if (othersKeys.has(compact)) {
    return 'Others'
  }

  return portfolio
}

function isExcludedPortfolio(portfolio) {
  return EXCLUDED_PORTFOLIOS.has(portfolio)
}

function normalizeEmployeeId(value) {
  return normalizeCellValue(value).replace(/\s+/g, '')
}

function ensureExcelFile(file) {
  if (!file) {
    return false
  }

  const lowered = file.name.toLowerCase()
  return ACCEPTED_EXTENSIONS.some((ext) => lowered.endsWith(ext))
}

function findHeaderByCandidates(headers, candidates, options = {}) {
  const exactOnlyCandidates = new Set(options.exactOnlyCandidates || [])
  const excludedParts = options.excludedParts || []
  const normalizedHeaders = headers.map((item) => ({
    raw: item,
    normalized: normalizeKey(item),
  })).filter(
    (header) =>
      header.normalized &&
      !excludedParts.some((part) => header.normalized.includes(part)),
  )

  for (const candidate of candidates) {
    const direct = normalizedHeaders.find(
      (header) => header.normalized === candidate,
    )
    if (direct) {
      return direct.raw
    }

    if (exactOnlyCandidates.has(candidate)) {
      continue
    }

    const includeMatch = normalizedHeaders.find((header) =>
      header.normalized.includes(candidate),
    )
    if (includeMatch) {
      return includeMatch.raw
    }
  }

  return null
}

function findEmployeeIdHeader(headers) {
  return findHeaderByCandidates(headers, HEADER_CANDIDATES.employeeId, {
    exactOnlyCandidates: ['employee', 'id'],
    excludedParts: [
      'name',
      'status',
      'portfolio',
      'customer',
      'month',
      'country',
      'geography',
      'branch',
      'grade',
      'type',
      'assignment',
      'policy',
      'competency',
      'category',
      'date',
      'geo',
      'iou',
      'project',
    ],
  })
}

function classifyGenericStatus(rawValue) {
  const normalized = normalizeCellValue(rawValue).toLowerCase()

  if (!normalized) {
    return '-'
  }

  if (
    normalized.includes('non-compliant') ||
    normalized.includes('non compliant') ||
    normalized.includes('noncompliant') ||
    normalized.includes('non-complaint') ||
    normalized.includes('non complaint') ||
    normalized.includes('noncomplaint') ||
    normalized.includes('not compliant') ||
    normalized.includes('not completed') ||
    normalized.includes('yet to complete') ||
    normalized.includes('incomplete')
  ) {
    return 'Defaulter'
  }

  if (normalized === 'e0' || normalized === 'e1') {
    return 'Defaulter'
  }

  if (
    normalized.includes('compliant') ||
    normalized.includes('completed') ||
    normalized === 'e2' ||
    normalized === 'e3'
  ) {
    return 'Completed'
  }

  return '-'
}

function classifyE2Status(rawValue) {
  const normalized = normalizeCellValue(rawValue).toLowerCase()

  if (!normalized) {
    return '-'
  }

  if (/(^|\b)e0(\b|$)/.test(normalized) || /(^|\b)e1(\b|$)/.test(normalized)) {
    return 'Defaulter'
  }

  if (/(^|\b)e2(\b|$)/.test(normalized) || /(^|\b)e3(\b|$)/.test(normalized)) {
    return 'Completed'
  }

  return classifyGenericStatus(normalized)
}

function mergeStatuses(existingStatus, nextStatus) {
  if (existingStatus === 'Defaulter' || nextStatus === 'Defaulter') {
    return 'Defaulter'
  }

  if (existingStatus === 'Completed' || nextStatus === 'Completed') {
    return 'Completed'
  }

  return '-'
}

function createBlankStatusMap() {
  return Object.fromEntries(
    COURSES.map((course) => [
      course.key,
      {
        previous: '-',
        current: '-',
      },
    ]),
  )
}

// Tokens we expect to see in a real header row. Used to score candidate header
// rows when reports include title/metadata rows above the actual table
// (common in TCS / HR exports such as the E2+ competency report).
const HEADER_DETECTION_TOKENS = [
  'employeeid',
  'employee',
  'employeenumber',
  'empid',
  'colleagueid',
  'personnumber',
  'associateid',
  'employeename',
  'name',
  'fullname',
  'portfolio',
  'department',
  'offdo',
  'status',
  'heldlevel',
  'acquiredlevel',
  'currentlevel',
  'competencylevel',
  'level',
  'compliance',
  'compliant',
  'iscompliant',
  'completionstatus',
  'coursecompletionstatus',
  'trainingstatus',
]

function detectHeaderRowIndex(matrix) {
  const limit = Math.min(matrix.length, 30)
  let bestIndex = 0
  let bestScore = -1

  for (let i = 0; i < limit; i++) {
    const row = matrix[i] || []
    const normalizedCells = row.map((cell) => normalizeKey(cell))
    const nonEmptyCount = normalizedCells.filter((cell) => cell).length

    // A header row should have at least a couple of populated columns.
    if (nonEmptyCount < 2) {
      continue
    }

    let score = 0
    for (const cell of normalizedCells) {
      if (!cell) continue
      for (const token of HEADER_DETECTION_TOKENS) {
        if (cell === token) {
          score += 3
          break
        }
        if (cell.includes(token)) {
          score += 1
          break
        }
      }
    }

    // Mild bonus for wider rows so a real header beats a stray metadata line
    // that happens to contain one matching word.
    score += Math.min(nonEmptyCount, 12) * 0.1

    if (score > bestScore) {
      bestScore = score
      bestIndex = i
    }
  }

  return bestIndex
}

async function parseRowsFromFile(file) {
  const arrayBuffer = await file.arrayBuffer()
  const workbook = XLSX.read(arrayBuffer, { type: 'array', cellDates: false })
  const firstSheetName = workbook.SheetNames[0]
  if (!firstSheetName) {
    return []
  }

  const worksheet = workbook.Sheets[firstSheetName]

  // Read as a 2D array first so we can locate the real header row even when
  // the sheet starts with title/metadata/filter rows above the table.
  const matrix = XLSX.utils.sheet_to_json(worksheet, {
    header: 1,
    defval: '',
    raw: false,
    blankrows: false,
  })

  if (!matrix.length) {
    return []
  }

  const headerRowIndex = detectHeaderRowIndex(matrix)
  const headerRow = matrix[headerRowIndex] || []

  // Build unique, non-empty header names. Empty cells get a synthetic label so
  // they don't collapse together; duplicates get a numeric suffix.
  const headers = []
  const seen = new Map()
  headerRow.forEach((cell, columnIndex) => {
    let name = normalizeCellValue(cell)
    if (!name) {
      name = `Column ${columnIndex + 1}`
    }
    const baseName = name
    let suffix = seen.get(baseName) || 0
    while (seen.has(name)) {
      suffix += 1
      name = `${baseName} (${suffix})`
    }
    seen.set(baseName, suffix)
    seen.set(name, 0)
    headers.push(name)
  })

  const rows = []
  for (let i = headerRowIndex + 1; i < matrix.length; i++) {
    const row = matrix[i]
    if (!row) continue
    const isBlank = headers.every(
      (_, idx) => normalizeCellValue(row[idx]) === '',
    )
    if (isBlank) continue

    const rowObject = {}
    headers.forEach((header, idx) => {
      rowObject[header] = row[idx] ?? ''
    })
    rows.push(rowObject)
  }

  return rows
}

async function parseHcMap(file, label) {
  const hcRows = await parseRowsFromFile(file)
  const hcHeaders = Object.keys(hcRows[0] || {})
  const hcIdHeader = findEmployeeIdHeader(hcHeaders)
  const hcPortfolioHeader = findHeaderByCandidates(
    hcHeaders,
    HEADER_CANDIDATES.portfolio,
  )
  const hcNameHeader = findHeaderByCandidates(
    hcHeaders,
    HEADER_CANDIDATES.employeeName,
  )
  const hcOffDoHeader = findHeaderByCandidates(
    hcHeaders,
    ['offdo'],
    { exactOnlyCandidates: ['offdo'] },
  )

  if (!hcIdHeader || !hcPortfolioHeader) {
    throw new Error(
      `Unable to detect Employee ID and Portfolio columns in ${label}. Found columns: ${hcHeaders.join(', ')}`,
    )
  }

  return hcRows.reduce((map, row) => {
    const employeeId = normalizeEmployeeId(row[hcIdHeader])
    if (!employeeId) {
      return map
    }

    map[employeeId] = {
      portfolio: normalizePortfolioName(row[hcPortfolioHeader]),
      employeeName:
        normalizeCellValue(row[hcNameHeader]) || map[employeeId]?.employeeName || '-',
      offDo: hcOffDoHeader
        ? normalizeCellValue(row[hcOffDoHeader])
        : '',
    }

    return map
  }, {})
}

function buildPortfolioOffDoMap(hcMaps) {
  const result = {}
  const sources = [hcMaps.current, hcMaps.previous]
  sources.forEach((source) => {
    if (!source) return
    Object.values(source).forEach((entry) => {
      const portfolio = entry?.portfolio
      const offDo = entry?.offDo
      if (!portfolio || !offDo) return
      // 'Others' is a catch-all bucket; never assign it an Off Do.
      if (portfolio === 'Others') return
      if (!result[portfolio]) {
        result[portfolio] = offDo
      }
    })
  })
  return result
}

function refreshEmployeePortfolio(employee) {
  employee.portfolio =
    employee.portfolios.current !== '-'
      ? employee.portfolios.current
      : employee.portfolios.previous !== '-'
        ? employee.portfolios.previous
        : 'Unmapped'
}

function getOrCreateEmployee(employeeMap, employeeId) {
  if (!employeeMap[employeeId]) {
    employeeMap[employeeId] = {
      employeeId,
      employeeName: '-',
      portfolio: 'Unmapped',
      portfolios: {
        previous: '-',
        current: '-',
      },
      statuses: createBlankStatusMap(),
    }
  }

  return employeeMap[employeeId]
}

function applyHcEntryToEmployee(employeeMap, employeeId, period, hcEntry) {
  const employee = getOrCreateEmployee(employeeMap, employeeId)

  if (hcEntry?.employeeName && (employee.employeeName === '-' || period === 'current')) {
    employee.employeeName = hcEntry.employeeName
  }

  employee.portfolios[period] = hcEntry?.portfolio || 'Unmapped'
  refreshEmployeePortfolio(employee)

  return employee
}

function employeeMatchesPortfolio(employee, portfolio) {
  if (portfolio === 'All') {
    return true
  }

  return (
    employee.portfolios.current === portfolio ||
    employee.portfolios.previous === portfolio ||
    employee.portfolio === portfolio
  )
}

function buildDeltaStyles(previousValue, currentValue) {
  if (previousValue === null || currentValue === null) {
    return { color: 'text.primary', weight: 600, trend: '' }
  }

  const delta = currentValue - previousValue
  if (delta > 0) {
    return { color: 'error.main', weight: 700, trend: `+${delta}` }
  }

  if (delta < 0) {
    return { color: 'success.main', weight: 700, trend: `${delta}` }
  }

  return { color: 'text.primary', weight: 600, trend: '0' }
}

function getPortfolioDefaulterTotal(row) {
  return COURSES.reduce((total, course) => {
    const previousCount = row.counts[course.key].previous
    const currentCount = row.counts[course.key].current

    return (
      total +
      (typeof previousCount === 'number' ? previousCount : 0) +
      (typeof currentCount === 'number' ? currentCount : 0)
    )
  }, 0)
}

function getOverallStatusForPeriod(employee, period) {
  const statuses = COURSES.map((course) => employee.statuses[course.key][period])

  if (statuses.some((status) => status === 'Defaulter')) {
    return 'Defaulter'
  }

  if (statuses.some((status) => status === 'Completed')) {
    return 'Completed'
  }

  return '-'
}

function getOverallStatusChipColor(status) {
  if (status === 'Defaulter') {
    return 'error'
  }

  if (status === 'Completed') {
    return 'success'
  }

  return 'default'
}

function getCourseDefaultersForPeriod(employee, period) {
  return COURSES.filter(
    (course) => employee.statuses[course.key][period] === 'Defaulter',
  ).map((course) => course.label)
}

function App() {
  const [currentFiles, setCurrentFiles] = useState(defaultUploadState)
  const [previousFiles, setPreviousFiles] = useState(defaultUploadState)
  const [hcFiles, setHcFiles] = useState(defaultHcUploadState)

  const [dashboardData, setDashboardData] = useState(null)
  const [searchQuery, setSearchQuery] = useState('')
  const [debouncedSearchQuery, setDebouncedSearchQuery] = useState('')
  const [selectedPortfolio, setSelectedPortfolio] = useState('All')
  const [errorText, setErrorText] = useState('')
  const [loading, setLoading] = useState(false)
  const [missingHcDialogOpen, setMissingHcDialogOpen] = useState(false)
  const [missingHcDialogPeriod, setMissingHcDialogPeriod] = useState('current')
  const [portfolioDialogOpen, setPortfolioDialogOpen] = useState(false)
  const [portfolioDialogPeriod, setPortfolioDialogPeriod] = useState('current')
  const [portfolioViewMode, setPortfolioViewMode] = useState('matrix')
  const [activePortfolioForEmployees, setActivePortfolioForEmployees] = useState(null)
  const debounceTimer = useRef(null)

  useEffect(() => {
    if (debounceTimer.current) {
      clearTimeout(debounceTimer.current)
    }

    if (!searchQuery) {
      debounceTimer.current = setTimeout(() => {
        setDebouncedSearchQuery('')
      }, 0)
      return
    }

    debounceTimer.current = setTimeout(() => {
      setDebouncedSearchQuery(searchQuery)
    }, 150)

    return () => {
      if (debounceTimer.current) {
        clearTimeout(debounceTimer.current)
      }
    }
  }, [searchQuery])

  const availablePortfolios = useMemo(() => {
    if (!dashboardData?.portfolioRows) {
      return ['All']
    }

    const portfolioNames = dashboardData.portfolioRows.map((row) => row.portfolio)
    const sorted = sortPortfoliosByPriority(portfolioNames)
    return ['All', ...sorted]
  }, [dashboardData])

  const filteredPortfolioRows = useMemo(() => {
    if (!dashboardData?.portfolioRows) {
      return []
    }

    let rows = dashboardData.portfolioRows

    if (selectedPortfolio !== 'All') {
      rows = rows.filter((row) => row.portfolio === selectedPortfolio)
    }

    // Preserve order from dashboardData.portfolioRows so Off Do groups remain contiguous.
    return rows
  }, [dashboardData, selectedPortfolio])

  const offDoGroupInfo = useMemo(() => {
    const sizes = {}
    filteredPortfolioRows.forEach((row) => {
      // 'Others' rows always stand alone — never bucketed under any Off Do.
      if (row.portfolio === 'Others') return
      const key = row.offDo || '-'
      sizes[key] = (sizes[key] || 0) + 1
    })

    const seen = new Set()
    return filteredPortfolioRows.map((row, idx) => {
      if (row.portfolio === 'Others') {
        return {
          row,
          offDoKey: `__others__${idx}`,
          isFirstInGroup: true,
          groupSize: 1,
        }
      }

      const key = row.offDo || '-'
      const isFirstInGroup = !seen.has(key)
      if (isFirstInGroup) seen.add(key)
      return {
        row,
        offDoKey: key,
        isFirstInGroup,
        groupSize: sizes[key],
      }
    })
  }, [filteredPortfolioRows])

  const portfolioSummaryRows = useMemo(() => {
    if (!dashboardData?.employees?.length || !filteredPortfolioRows.length) {
      return []
    }

    const employees = dashboardData.employees
    const hasPrevious = dashboardData.hasPrevious

    return filteredPortfolioRows.map((row) => {
      const portfolio = row.portfolio

      const currentTotal = employees.filter(
        (employee) => employee.portfolios.current === portfolio,
      ).length
      const currentDefaulters = employees.filter(
        (employee) =>
          employee.portfolios.current === portfolio &&
          COURSES.some(
            (course) => employee.statuses[course.key].current === 'Defaulter',
          ),
      ).length

      const previousTotal = hasPrevious
        ? employees.filter(
            (employee) => employee.portfolios.previous === portfolio,
          ).length
        : null
      const previousDefaulters = hasPrevious
        ? employees.filter(
            (employee) =>
              employee.portfolios.previous === portfolio &&
              COURSES.some(
                (course) => employee.statuses[course.key].previous === 'Defaulter',
              ),
          ).length
        : null

      const currentPercent =
        currentTotal > 0 ? (currentDefaulters / currentTotal) * 100 : 0
      const previousPercent =
        hasPrevious && previousTotal > 0
          ? (previousDefaulters / previousTotal) * 100
          : null

      return {
        portfolio,
        offDo: row.offDo,
        previousTotal,
        previousDefaulters,
        previousPercent,
        currentTotal,
        currentDefaulters,
        currentPercent,
      }
    })
  }, [dashboardData, filteredPortfolioRows])

  const employeePool = useMemo(() => {
    if (!dashboardData?.employees) {
      return []
    }

    if (selectedPortfolio === 'All') {
      return dashboardData.employees
    }

    return dashboardData.employees.filter((employee) =>
      employeeMatchesPortfolio(employee, selectedPortfolio),
    )
  }, [dashboardData, selectedPortfolio])

  const normalizedSearch = useMemo(
    () => normalizeCellValue(debouncedSearchQuery).toLowerCase(),
    [debouncedSearchQuery],
  )

  const normalizedSearchId = useMemo(
    () => normalizeEmployeeId(debouncedSearchQuery).toLowerCase(),
    [debouncedSearchQuery],
  )

  const portfolioEmployeeRows = useMemo(() => {
    if (!dashboardData?.employees?.length || !activePortfolioForEmployees) {
      return []
    }

    const portfolioEmployees = dashboardData.employees.filter(
      (employee) =>
        employee.portfolios.current === activePortfolioForEmployees ||
        employee.portfolios.previous === activePortfolioForEmployees ||
        employee.portfolio === activePortfolioForEmployees,
    )

    if (!debouncedSearchQuery) {
      return portfolioEmployees.slice(0, MAX_EMPLOYEE_RESULTS)
    }

    return portfolioEmployees
      .filter((employee) => {
        const byId = employee.employeeId.toLowerCase().includes(normalizedSearchId)
        const byName = employee.employeeName.toLowerCase().includes(normalizedSearch)
        return byId || byName
      })
      .slice(0, MAX_EMPLOYEE_RESULTS)
  }, [
    activePortfolioForEmployees,
    dashboardData,
    debouncedSearchQuery,
    normalizedSearch,
    normalizedSearchId,
  ])

  const exactEmployeeMatch = useMemo(() => {
    if (!normalizedSearchId || !employeePool.length) {
      return null
    }

    return (
      employeePool.find(
        (employee) => employee.employeeId.toLowerCase() === normalizedSearchId,
      ) || null
    )
  }, [employeePool, normalizedSearchId])

  function updateCourseFile(period, courseKey, file) {
    setErrorText('')

    if (file && !ensureExcelFile(file)) {
      setErrorText('Only Excel files are supported (.xlsx, .xls, .xlsm).')
      return
    }

    if (period === 'current') {
      setCurrentFiles((prev) => ({ ...prev, [courseKey]: file || null }))
      return
    }

    setPreviousFiles((prev) => ({ ...prev, [courseKey]: file || null }))
  }

  function updateHcFile(period, file) {
    setErrorText('')

    if (file && !ensureExcelFile(file)) {
      setErrorText('Headcount file must be an Excel file (.xlsx, .xls, .xlsm).')
      return
    }

    setHcFiles((prev) => ({ ...prev, [period]: file || null }))
  }

  function resetDashboard() {
    setCurrentFiles(defaultUploadState)
    setPreviousFiles(defaultUploadState)
    setHcFiles(defaultHcUploadState)
    setDashboardData(null)
    setSearchQuery('')
    setSelectedPortfolio('All')
    setActivePortfolioForEmployees(null)
    setErrorText('')
  }

  async function processFiles() {
    setErrorText('')

    const allCurrentAvailable = COURSES.every((course) => currentFiles[course.key])
    if (!allCurrentAvailable) {
      setErrorText('Upload all 4 Current course files before processing.')
      return
    }

    if (!hcFiles.current) {
      setErrorText('Upload the Current Headcount file before processing.')
      return
    }

    const previousFileCount = COURSES.filter(
      (course) => previousFiles[course.key],
    ).length

    const hasPrevious = previousFileCount > 0
    const previousCourseAvailability = Object.fromEntries(
      COURSES.map((course) => [course.key, Boolean(previousFiles[course.key])]),
    )

    setLoading(true)

    try {
      const currentHcMap = await parseHcMap(hcFiles.current, 'Current Headcount file')
      const hcMaps = {
        current: currentHcMap,
        previous:
          hasPrevious && hcFiles.previous
            ? await parseHcMap(hcFiles.previous, 'Previous Headcount file')
            : hasPrevious
              ? currentHcMap
              : {},
      }
      const portfolioOffDoMap = buildPortfolioOffDoMap(hcMaps)
      const employeeMap = {}

      Object.entries(hcMaps.current).forEach(([employeeId, hcEntry]) => {
        applyHcEntryToEmployee(employeeMap, employeeId, 'current', hcEntry)
      })

      if (hasPrevious) {
        Object.entries(hcMaps.previous).forEach(([employeeId, hcEntry]) => {
          applyHcEntryToEmployee(employeeMap, employeeId, 'previous', hcEntry)
        })
      }

      const missingFromHcIds = {
        previous: new Set(),
        current: new Set(),
      }

      const filesToProcess = []
      for (const course of COURSES) {
        filesToProcess.push({
          file: currentFiles[course.key],
          courseKey: course.key,
          period: 'current',
        })

        if (previousFiles[course.key]) {
          filesToProcess.push({
            file: previousFiles[course.key],
            courseKey: course.key,
            period: 'previous',
          })
        }
      }

      for (const item of filesToProcess) {
        if (!item.file) {
          continue
        }

        const rows = await parseRowsFromFile(item.file)
        const headers = Object.keys(rows[0] || {})

        const idHeader = findEmployeeIdHeader(headers)
        const nameHeader = findHeaderByCandidates(
          headers,
          HEADER_CANDIDATES.employeeName,
        )

        const statusHeader = findHeaderByCandidates(
          headers,
          item.courseKey === 'e2'
            ? HEADER_CANDIDATES.e2Status
            : item.courseKey === 'mandatory'
              ? HEADER_CANDIDATES.mandatoryStatus
              : HEADER_CANDIDATES.statusGeneric,
        )

        if (!idHeader || !statusHeader) {
          const missing = []
          if (!idHeader) missing.push('Employee ID')
          if (!statusHeader) missing.push('Status')
          const detectedHeaders = headers.filter((h) => h && !/^Column \d+$/.test(h))
          throw new Error(
            `Unable to detect ${missing.join(' and ')} column${missing.length > 1 ? 's' : ''} in ${item.file.name}. ` +
              `Detected columns: ${detectedHeaders.length ? detectedHeaders.join(', ') : '(none)'}.`,
          )
        }

        rows.forEach((row) => {
          const employeeId = normalizeEmployeeId(row[idHeader])
          if (!employeeId) {
            return
          }

          const periodHcEntry = hcMaps[item.period][employeeId]
          if (!periodHcEntry) {
            missingFromHcIds[item.period].add(employeeId)
            return
          }

          const entry = applyHcEntryToEmployee(
            employeeMap,
            employeeId,
            item.period,
            periodHcEntry,
          )

          if (nameHeader) {
            const parsedName = normalizeCellValue(row[nameHeader])
            if (parsedName) {
              entry.employeeName = parsedName
            }
          }

          const rawStatus = row[statusHeader]
          const nextStatus =
            item.courseKey === 'e2'
              ? classifyE2Status(rawStatus)
              : classifyGenericStatus(rawStatus)

          const currentStatus = entry.statuses[item.courseKey][item.period]
          entry.statuses[item.courseKey][item.period] = mergeStatuses(
            currentStatus,
            nextStatus,
          )
        })
      }

      if (!hasPrevious) {
        Object.values(employeeMap).forEach((employee) => {
          COURSES.forEach((course) => {
            employee.statuses[course.key].previous = '-'
          })
        })
      }

      const employees = Object.values(employeeMap).sort((a, b) => {
        const byPortfolio = a.portfolio.localeCompare(b.portfolio)
        if (byPortfolio !== 0) {
          return byPortfolio
        }

        return a.employeeId.localeCompare(b.employeeId)
      })

      const portfolioNames = new Set()

      employees.forEach((employee) => {
        if (
          employee.portfolios.current !== '-' &&
          !isExcludedPortfolio(employee.portfolios.current)
        ) {
          portfolioNames.add(employee.portfolios.current)
        }

        if (
          hasPrevious &&
          employee.portfolios.previous !== '-' &&
          !isExcludedPortfolio(employee.portfolios.previous)
        ) {
          portfolioNames.add(employee.portfolios.previous)
        }
      })

      const portfolioRows = sortPortfoliosByPriority(Array.from(portfolioNames))
        .map((portfolio) => {
          const counts = {}

          COURSES.forEach((course) => {
            const previousCount = hasPrevious
              && previousCourseAvailability[course.key]
              ? employees.filter(
                  (employee) =>
                    employee.portfolios.previous === portfolio &&
                    employee.statuses[course.key].previous === 'Defaulter',
                ).length
              : null

            const currentCount = employees.filter(
              (employee) =>
                employee.portfolios.current === portfolio &&
                employee.statuses[course.key].current === 'Defaulter',
            ).length

            counts[course.key] = {
              previous: previousCount,
              current: currentCount,
            }
          })

          return {
            portfolio,
            counts,
            offDo: portfolioOffDoMap[portfolio] || '-',
          }
        })
        .filter((row) =>
          COURSES.some((course) => {
            const previousCount = row.counts[course.key].previous
            const currentCount = row.counts[course.key].current

            return (
              (typeof previousCount === 'number' && previousCount > 0) ||
              (typeof currentCount === 'number' && currentCount > 0)
            )
          }),
        )
        .sort((a, b) => {
          const aHasOffDo = a.offDo && a.offDo !== '-'
          const bHasOffDo = b.offDo && b.offDo !== '-'

          if (aHasOffDo && !bHasOffDo) return -1
          if (!aHasOffDo && bHasOffDo) return 1
          if (aHasOffDo && bHasOffDo && a.offDo !== b.offDo) {
            return a.offDo.localeCompare(b.offDo)
          }

          const aIsMajor = MAJOR_PORTFOLIOS.includes(a.portfolio)
          const bIsMajor = MAJOR_PORTFOLIOS.includes(b.portfolio)

          if (aIsMajor && !bIsMajor) return -1
          if (!aIsMajor && bIsMajor) return 1

          if (aIsMajor && bIsMajor) {
            const aMajorIndex = MAJOR_PORTFOLIOS.indexOf(a.portfolio)
            const bMajorIndex = MAJOR_PORTFOLIOS.indexOf(b.portfolio)
            return aMajorIndex - bMajorIndex
          }

          const totalDifference = getPortfolioDefaulterTotal(b) - getPortfolioDefaulterTotal(a)
          if (totalDifference !== 0) {
            return totalDifference
          }

          return a.portfolio.localeCompare(b.portfolio)
        })

      setDashboardData({
        hasPrevious,
        hasPreviousHc: Boolean(hcFiles.previous),
        previousFileCount,
        previousCourseAvailability,
        employees,
        portfolioRows,
        missingFromHc: {
          previous: hasPrevious ? missingFromHcIds.previous.size : 0,
          current: missingFromHcIds.current.size,
          total: new Set([
            ...missingFromHcIds.current,
            ...(hasPrevious ? missingFromHcIds.previous : []),
          ]).size,
          currentIds: Array.from(missingFromHcIds.current).sort(),
          previousIds: hasPrevious ? Array.from(missingFromHcIds.previous).sort() : [],
        },
        processedAt: new Date().toISOString(),
      })

      setSelectedPortfolio('All')
      setActivePortfolioForEmployees(null)
    } catch (error) {
      setDashboardData(null)
      setErrorText(error.message || 'Failed to parse uploaded files.')
    } finally {
      setLoading(false)
    }
  }

  function exportDashboard() {
    if (!dashboardData?.portfolioRows?.length) {
      setErrorText('Process files first to export dashboard data.')
      return
    }

    const header = [
      'Off Do',
      'Portfolio',
      'Secuware (Prev)',
      'Secuware (Curr)',
      'Learning Coach (Prev)',
      'Learning Coach (Curr)',
      'E2 (Prev)',
      'E2 (Curr)',
      'Mandatory (Prev)',
      'Mandatory (Curr)',
    ]

    const rows = dashboardData.portfolioRows.map((row) => [
      row.offDo || '-',
      row.portfolio,
      row.counts.secuware.previous ?? '-',
      row.counts.secuware.current ?? '-',
      row.counts.learningCoach.previous ?? '-',
      row.counts.learningCoach.current ?? '-',
      row.counts.e2.previous ?? '-',
      row.counts.e2.current ?? '-',
      row.counts.mandatory.previous ?? '-',
      row.counts.mandatory.current ?? '-',
    ])

    const worksheet = XLSX.utils.aoa_to_sheet([header, ...rows])
    const workbook = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Defaulter Dashboard')
    XLSX.writeFile(workbook, `defaulter-dashboard-${Date.now()}.xlsx`)
  }

  function exportPortfolioSummary() {
    if (!dashboardData?.portfolioRows?.length) {
      setErrorText('Process files first to export dashboard data.')
      return
    }

    const employees = dashboardData.employees
    const hasPrevious = dashboardData.hasPrevious

    const summary = dashboardData.portfolioRows.map((row) => {
      const portfolio = row.portfolio
      const offDo = row.offDo || '-'

      const currentTotal = employees.filter(
        (employee) => employee.portfolios.current === portfolio,
      ).length
      const currentDefaulters = employees.filter(
        (employee) =>
          employee.portfolios.current === portfolio &&
          COURSES.some(
            (course) => employee.statuses[course.key].current === 'Defaulter',
          ),
      ).length

      const previousTotal = hasPrevious
        ? employees.filter(
            (employee) => employee.portfolios.previous === portfolio,
          ).length
        : null
      const previousDefaulters = hasPrevious
        ? employees.filter(
            (employee) =>
              employee.portfolios.previous === portfolio &&
              COURSES.some(
                (course) => employee.statuses[course.key].previous === 'Defaulter',
              ),
          ).length
        : null

      const currentPercent =
        currentTotal > 0 ? (currentDefaulters / currentTotal) * 100 : 0
      const previousPercent =
        hasPrevious && previousTotal !== null && previousTotal > 0
          ? (previousDefaulters / previousTotal) * 100
          : null

      return {
        offDo,
        portfolio,
        previousTotal,
        previousDefaulters,
        previousPercent,
        currentTotal,
        currentDefaulters,
        currentPercent,
      }
    })

    const header = hasPrevious
      ? [
          'Off Do',
          'Portfolio',
          'Total Employees (Prev)',
          'Total Employees (Curr)',
          'Total Defaulters (Prev)',
          'Total Defaulters (Curr)',
          'Defaulter % (Prev)',
          'Defaulter % (Curr)',
        ]
      : [
          'Off Do',
          'Portfolio',
          'Total Employees',
          'Total Defaulters',
          'Defaulter %',
        ]

    const formatPercent = (value) =>
      value === null || value === undefined ? '-' : `${value.toFixed(2)}%`

    const rows = summary.map((row) =>
      hasPrevious
        ? [
            row.offDo,
            row.portfolio,
            row.previousTotal ?? '-',
            row.currentTotal,
            row.previousDefaulters ?? '-',
            row.currentDefaulters,
            formatPercent(row.previousPercent),
            formatPercent(row.currentPercent),
          ]
        : [
            row.offDo,
            row.portfolio,
            row.currentTotal,
            row.currentDefaulters,
            formatPercent(row.currentPercent),
          ],
    )

    const worksheet = XLSX.utils.aoa_to_sheet([header, ...rows])
    const workbook = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Portfolio Summary')
    XLSX.writeFile(workbook, `portfolio-defaulter-summary-${Date.now()}.xlsx`)
  }

  const stats = useMemo(() => {
    if (!dashboardData) {
      return {
        portfolios: 0,
        employees: 0,
        defaultersNow: 0,
        missingFromHc: { previous: 0, current: 0, total: 0 },
      }
    }

    const totalCurrentDefaulters = dashboardData.employees.reduce((acc, employee) => {
      const hasDefaulter = COURSES.some(
        (course) => employee.statuses[course.key].current === 'Defaulter',
      )
      return acc + (hasDefaulter ? 1 : 0)
    }, 0)

    return {
      portfolios: dashboardData.portfolioRows.length,
      employees: dashboardData.employees.length,
      defaultersNow: totalCurrentDefaulters,
      missingFromHc: dashboardData.missingFromHc,
    }
  }, [dashboardData])

  return (
    <ThemeProvider theme={theme}>
      <Box className="dashboard-page">
        <Container maxWidth="xl" className="fade-in-up">
          <Paper className="dashboard-header" elevation={0}>
            <Stack
              direction={{ xs: 'column', md: 'row' }}
              spacing={2}
              alignItems={{ xs: 'flex-start', md: 'center' }}
              justifyContent="space-between"
            >
              <Stack direction="row" spacing={1.5} alignItems="center">
                <Box className="dashboard-logo-badge">
                  <AnalyticsRoundedIcon />
                </Box>
                <Box>
                  <Typography variant="h3" className="dashboard-title">
                    Course Defaulter Portfolio Dashboard
                  </Typography>
                </Box>
              </Stack>
              <Chip
                label="Excel Portfolio Analytics"
                color="primary"
                variant="outlined"
                className="dashboard-status-chip"
              />
            </Stack>
          </Paper>

          <Grid container spacing={2.5} mb={3.75}>
            <Grid size={{ xs: 12, md: 4 }}>
              <Paper className="file-upload-card file-upload-card--current-courses" elevation={0}>
                <Stack direction="row" alignItems="center" justifyContent="space-between" mb={1.5}>
                  <Typography variant="h6">Current Data</Typography>
                  <Chip
                    size="small"
                    label={`${COURSES.filter((course) => currentFiles[course.key]).length}/4`}
                    color={COURSES.every((course) => currentFiles[course.key]) ? 'success' : 'default'}
                    variant="outlined"
                  />
                </Stack>
                <Stack spacing={1.5}>
                  {COURSES.map((course) => (
                    <Button
                      key={`current-${course.key}`}
                      variant="outlined"
                      component="label"
                      color="primary"
                      startIcon={
                        currentFiles[course.key] ? (
                          <CheckCircleRoundedIcon />
                        ) : (
                          <UploadFileRoundedIcon />
                        )
                      }
                      className={`file-select-button ${
                        currentFiles[course.key] ? 'file-select-button--has-file' : ''
                      }`}
                      sx={{ justifyContent: 'space-between' }}
                    >
                      <span>{course.label} (Current)</span>
                      <Chip
                        size="small"
                        label={currentFiles[course.key] ? 'Ready' : 'Upload'}
                        color={currentFiles[course.key] ? 'success' : 'default'}
                        variant={currentFiles[course.key] ? 'filled' : 'outlined'}
                      />
                      <input
                        hidden
                        type="file"
                        accept=".xlsx,.xls,.xlsm"
                        onChange={(event) =>
                          updateCourseFile(
                            'current',
                            course.key,
                            event.target.files?.[0] || null,
                          )
                        }
                      />
                    </Button>
                  ))}
                </Stack>
              </Paper>
            </Grid>

            <Grid size={{ xs: 12, md: 4 }}>
              <Paper className="file-upload-card file-upload-card--previous-courses" elevation={0}>
                <Stack direction="row" alignItems="center" justifyContent="space-between" mb={1.5}>
                  <Typography variant="h6">Previous Data</Typography>
                  <Chip
                    size="small"
                    label={`${COURSES.filter((course) => previousFiles[course.key]).length}/4`}
                    color={COURSES.some((course) => previousFiles[course.key]) ? 'secondary' : 'default'}
                    variant="outlined"
                  />
                </Stack>
                <Stack spacing={1.5}>
                  {COURSES.map((course) => (
                    <Button
                      key={`previous-${course.key}`}
                      variant="outlined"
                      component="label"
                      color="secondary"
                      startIcon={
                        previousFiles[course.key] ? (
                          <CheckCircleRoundedIcon />
                        ) : (
                          <UploadFileRoundedIcon />
                        )
                      }
                      className={`file-select-button ${
                        previousFiles[course.key] ? 'file-select-button--has-file' : ''
                      }`}
                      sx={{ justifyContent: 'space-between' }}
                    >
                      <span>{course.label} (Previous)</span>
                      <Chip
                        size="small"
                        label={previousFiles[course.key] ? 'Ready' : 'Optional'}
                        color={previousFiles[course.key] ? 'secondary' : 'default'}
                        variant={previousFiles[course.key] ? 'filled' : 'outlined'}
                      />
                      <input
                        hidden
                        type="file"
                        accept=".xlsx,.xls,.xlsm"
                        onChange={(event) =>
                          updateCourseFile(
                            'previous',
                            course.key,
                            event.target.files?.[0] || null,
                          )
                        }
                      />
                    </Button>
                  ))}
                </Stack>
              </Paper>
            </Grid>

            <Grid size={{ xs: 12, md: 4 }}>
              <Paper className="file-upload-card file-upload-card--headcount" elevation={0}>
                <Stack direction="row" alignItems="center" justifyContent="space-between" mb={1.5}>
                  <Typography variant="h6">Headcount</Typography>
                  <Chip
                    size="small"
                    label={`${Number(Boolean(hcFiles.current)) + Number(Boolean(hcFiles.previous))}/2`}
                    color={hcFiles.current ? 'warning' : 'default'}
                    variant="outlined"
                  />
                </Stack>

                <Stack spacing={1.5}>
                  <Button
                    variant="outlined"
                    component="label"
                    color="warning"
                    startIcon={
                      hcFiles.current ? <CheckCircleRoundedIcon /> : <UploadFileRoundedIcon />
                    }
                    className={`file-select-button ${hcFiles.current ? 'file-select-button--has-file' : ''}`}
                    sx={{ width: '100%', justifyContent: 'space-between' }}
                  >
                    <span>Headcount (Current)</span>
                    <Chip
                      size="small"
                      label={hcFiles.current ? 'Ready' : 'Required'}
                      color={hcFiles.current ? 'success' : 'warning'}
                      variant={hcFiles.current ? 'filled' : 'outlined'}
                    />
                    <input
                      hidden
                      type="file"
                      accept=".xlsx,.xls,.xlsm"
                      onChange={(event) =>
                        updateHcFile('current', event.target.files?.[0] || null)
                      }
                    />
                  </Button>
                  <Button
                    variant="outlined"
                    component="label"
                    color="secondary"
                    startIcon={
                      hcFiles.previous ? <CheckCircleRoundedIcon /> : <UploadFileRoundedIcon />
                    }
                    className={`file-select-button ${hcFiles.previous ? 'file-select-button--has-file' : ''}`}
                    sx={{ width: '100%', justifyContent: 'space-between' }}
                  >
                    <span>Headcount (Previous)</span>
                    <Chip
                      size="small"
                      label={hcFiles.previous ? 'Ready' : 'Optional'}
                      color={hcFiles.previous ? 'secondary' : 'default'}
                      variant={hcFiles.previous ? 'filled' : 'outlined'}
                    />
                    <input
                      hidden
                      type="file"
                      accept=".xlsx,.xls,.xlsm"
                      onChange={(event) =>
                        updateHcFile('previous', event.target.files?.[0] || null)
                      }
                    />
                  </Button>
                </Stack>

                <Box sx={{ mt: 2 }}>
                  <Typography variant="body2" color="text.secondary">
                    Current Headcount: {hcFiles.current ? hcFiles.current.name : 'Not selected'}
                  </Typography>
                  <Typography variant="body2" color="text.secondary">
                    Previous Headcount: {hcFiles.previous ? hcFiles.previous.name : 'Not selected'}
                  </Typography>
                </Box>

                <Divider sx={{ my: 2 }} />

                <Stack direction="row" spacing={1.2}>
                  <Button
                    variant="contained"
                    onClick={processFiles}
                    disabled={loading}
                    startIcon={<PlayArrowRoundedIcon />}
                    sx={{ flex: 1 }}
                  >
                    {loading ? 'Processing...' : 'Process Files'}
                  </Button>
                  <Button
                    variant="outlined"
                    onClick={resetDashboard}
                    startIcon={<RestartAltRoundedIcon />}
                  >
                    Reset
                  </Button>
                </Stack>
              </Paper>
            </Grid>
          </Grid>

          {errorText && (
            <Alert severity="error" sx={{ mb: 2 }}>
              {errorText}
            </Alert>
          )}

          {dashboardData && (
            <Stack spacing={3} className="fade-in-up" sx={{ mt: '30px' }}>
              <Grid container spacing={2}>
                <Grid size={{ xs: 12, sm: 6, lg: 2.4 }}>
                  <Paper
                    className="summary-tile summary-tile--clickable summary-tile--teal-clickable"
                    elevation={0}
                    onClick={() => {
                      setPortfolioDialogPeriod('current')
                      setPortfolioDialogOpen(true)
                    }}
                    role="button"
                    tabIndex={0}
                    onKeyDown={(e) => {
                      if (e.key === 'Enter' || e.key === ' ') {
                        setPortfolioDialogPeriod('current')
                        setPortfolioDialogOpen(true)
                      }
                    }}
                  >
                    <Stack direction="row" alignItems="center" justifyContent="space-between">
                      <Box>
                        <Typography color="text.secondary">Defaulter Portfolios</Typography>
                        <Typography variant="h6">{stats.portfolios}</Typography>
                      </Box>
                      <Box className="summary-tile-icon summary-tile-icon--teal">
                        <BusinessCenterRoundedIcon fontSize="small" />
                      </Box>
                    </Stack>
                  </Paper>
                </Grid>
                <Grid size={{ xs: 12, sm: 6, lg: 2.4 }}>
                  <Paper className="summary-tile" elevation={0}>
                    <Stack direction="row" alignItems="center" justifyContent="space-between">
                      <Box>
                        <Typography color="text.secondary">Headcount Employees</Typography>
                        <Typography variant="h6">{stats.employees}</Typography>
                      </Box>
                      <Box className="summary-tile-icon summary-tile-icon--blue">
                        <GroupsRoundedIcon fontSize="small" />
                      </Box>
                    </Stack>
                  </Paper>
                </Grid>
                <Grid size={{ xs: 12, sm: 6, lg: 2.4 }}>
                  <Paper className="summary-tile" elevation={0}>
                    <Stack direction="row" alignItems="center" justifyContent="space-between">
                      <Box>
                        <Typography color="text.secondary">
                          Current Defaulters (Any Course)
                        </Typography>
                        <Typography variant="h6">{stats.defaultersNow}</Typography>
                      </Box>
                      <Box className="summary-tile-icon summary-tile-icon--amber">
                        <InsightsRoundedIcon fontSize="small" />
                      </Box>
                    </Stack>
                  </Paper>
                </Grid>
                <Grid size={{ xs: 12, sm: 6, lg: 2.4 }}>
                  <Paper
                    className="summary-tile summary-tile--warning summary-tile--clickable"
                    elevation={0}
                    onClick={() => {
                      setMissingHcDialogPeriod('current')
                      setMissingHcDialogOpen(true)
                    }}
                    role="button"
                    tabIndex={0}
                    onKeyDown={(e) => {
                      if (e.key === 'Enter' || e.key === ' ') {
                        setMissingHcDialogPeriod('current')
                        setMissingHcDialogOpen(true)
                      }
                    }}
                  >
                    <Stack direction="row" alignItems="center" justifyContent="space-between">
                      <Box>
                        <Typography color="text.secondary">Missing From Headcount</Typography>
                        <Typography variant="h6">{stats.missingFromHc.total}</Typography>
                        <Typography variant="caption" color="text.secondary">
                          Current {stats.missingFromHc.current}
                          {dashboardData.hasPrevious
                            ? ` / Previous ${stats.missingFromHc.previous}`
                            : ''}
                        </Typography>
                      </Box>
                      <Box className="summary-tile-icon summary-tile-icon--orange">
                        <ReportProblemRoundedIcon fontSize="small" />
                      </Box>
                    </Stack>
                  </Paper>
                </Grid>
                <Grid size={{ xs: 12, sm: 6, lg: 2.4 }}>
                  <Paper className="summary-tile" elevation={0}>
                    <Stack direction="row" alignItems="center" justifyContent="space-between">
                      <Box>
                        <Typography color="text.secondary">Mode</Typography>
                        <Typography variant="h6">
                          {dashboardData.hasPrevious
                            ? `Previous ${dashboardData.previousFileCount}/4`
                            : 'Current only'}
                        </Typography>
                      </Box>
                      <Box className="summary-tile-icon summary-tile-icon--slate">
                        <AnalyticsRoundedIcon fontSize="small" />
                      </Box>
                    </Stack>
                  </Paper>
                </Grid>
              </Grid>

              <Paper className="controls-toolbar" elevation={0}>
                <Stack
                  direction={{ xs: 'column', md: 'row' }}
                  spacing={1.2}
                  justifyContent="space-between"
                >
                  <Stack direction={{ xs: 'column', md: 'row' }} spacing={1.2} sx={{ flex: 1 }}>
                    <TextField
                      label="Search Employee ID or Name"
                      value={searchQuery}
                      onChange={(event) => setSearchQuery(event.target.value)}
                      size="small"
                      InputProps={{
                        startAdornment: (
                          <InputAdornment position="start">
                            <SearchRoundedIcon fontSize="small" />
                          </InputAdornment>
                        ),
                      }}
                      sx={{ flex: 1, minWidth: 220 }}
                    />

                    <FormControl size="small" sx={{ minWidth: 220 }}>
                      <InputLabel>Portfolio Filter</InputLabel>
                      <Select
                        value={selectedPortfolio}
                        label="Portfolio Filter"
                        onChange={(event) => setSelectedPortfolio(event.target.value)}
                      >
                        {availablePortfolios.map((item) => (
                          <MenuItem key={item} value={item}>
                            {item}
                          </MenuItem>
                        ))}
                      </Select>
                    </FormControl>
                  </Stack>

                  <Stack direction={{ xs: 'column', sm: 'row' }} spacing={1.2}>
                    <Button
                      variant="contained"
                      color="secondary"
                      onClick={exportDashboard}
                      startIcon={<FileDownloadRoundedIcon />}
                    >
                      Export Matrix
                    </Button>
                    <Button
                      variant="outlined"
                      color="secondary"
                      onClick={exportPortfolioSummary}
                      startIcon={<FileDownloadRoundedIcon />}
                    >
                      Export Summary
                    </Button>
                  </Stack>
                </Stack>
              </Paper>

              <Paper className="employee-lookup-panel" elevation={0}>
                <Stack
                  direction={{ xs: 'column', md: 'row' }}
                  spacing={1.4}
                  alignItems={{ xs: 'flex-start', md: 'center' }}
                  justifyContent="space-between"
                >
                  <Box>
                    <Typography variant="h6">Employee Quick Status</Typography>
                    {exactEmployeeMatch ? (
                      <Typography variant="body2" color="text.secondary">
                        {exactEmployeeMatch.employeeName} ({exactEmployeeMatch.employeeId})
                      </Typography>
                    ) : (
                      <Typography variant="body2" color="text.secondary">
                        Enter an exact Employee ID to get direct status.
                      </Typography>
                    )}
                  </Box>

                  {exactEmployeeMatch ? (() => {
                    const currentOverall = getOverallStatusForPeriod(
                      exactEmployeeMatch,
                      'current',
                    )
                    const previousOverall = getOverallStatusForPeriod(
                      exactEmployeeMatch,
                      'previous',
                    )
                    const currentDefaulters = getCourseDefaultersForPeriod(
                      exactEmployeeMatch,
                      'current',
                    )
                    const previousDefaulters = getCourseDefaultersForPeriod(
                      exactEmployeeMatch,
                      'previous',
                    )

                    return (
                      <Stack direction="column" spacing={1} alignItems="flex-start">
                        <Stack direction="row" spacing={1} flexWrap="wrap" useFlexGap>
                          <Chip
                            label={
                              currentOverall === 'Defaulter'
                                ? `Current: Defaulter in ${currentDefaulters.join(', ')}`
                                : `Current: ${currentOverall}`
                            }
                            color={getOverallStatusChipColor(currentOverall)}
                          />
                          {dashboardData.hasPrevious && (
                            <Chip
                              label={
                                previousOverall === 'Defaulter'
                                  ? `Previous: Defaulter in ${previousDefaulters.join(', ')}`
                                  : `Previous: ${previousOverall}`
                              }
                              color={getOverallStatusChipColor(previousOverall)}
                              variant="outlined"
                            />
                          )}
                        </Stack>
                      </Stack>
                    )
                  })() : null}
                </Stack>
              </Paper>

              <Paper className="data-table-panel" elevation={0}>
                <Stack
                  direction={{ xs: 'column', sm: 'row' }}
                  spacing={1.5}
                  justifyContent="space-between"
                  alignItems={{ xs: 'flex-start', sm: 'center' }}
                  mb={1.5}
                >
                  <Typography variant="h6">
                    {portfolioViewMode === 'matrix'
                      ? 'Portfolio Defaulter Matrix'
                      : 'Portfolio Defaulter Summary'}
                  </Typography>
                  <Stack direction="row" spacing={1}>
                    <Chip
                      label="Matrix"
                      color={portfolioViewMode === 'matrix' ? 'primary' : 'default'}
                      variant={portfolioViewMode === 'matrix' ? 'filled' : 'outlined'}
                      onClick={() => setPortfolioViewMode('matrix')}
                      clickable
                    />
                    <Chip
                      label="Summary"
                      color={portfolioViewMode === 'summary' ? 'primary' : 'default'}
                      variant={portfolioViewMode === 'summary' ? 'filled' : 'outlined'}
                      onClick={() => setPortfolioViewMode('summary')}
                      clickable
                    />
                  </Stack>
                </Stack>

                {portfolioViewMode === 'matrix' ? (
                <TableContainer>
                  <Table size="small" stickyHeader>
                    <TableHead>
                      <TableRow>
                        <TableCell>Off Do</TableCell>
                        <TableCell>Portfolio</TableCell>
                        {COURSES.map((course) => (
                          <Fragment key={`${course.key}-headers`}>
                            <TableCell align="right">
                              {course.label} (Prev)
                            </TableCell>
                            <TableCell align="right">
                              {course.label} (Curr)
                            </TableCell>
                          </Fragment>
                        ))}
                      </TableRow>
                    </TableHead>
                    <TableBody>
                      {offDoGroupInfo.map(({ row, isFirstInGroup, groupSize }) => (
                        <TableRow key={row.portfolio} hover>
                          {isFirstInGroup && (
                            <TableCell
                              rowSpan={groupSize}
                              sx={{
                                fontWeight: 700,
                                verticalAlign: 'top',
                                backgroundColor: 'rgba(15, 118, 110, 0.06)',
                                borderRight: '1px solid rgba(0,0,0,0.08)',
                                whiteSpace: 'nowrap',
                              }}
                            >
                              {row.offDo || '-'}
                            </TableCell>
                          )}
                          <TableCell
                            sx={{
                              fontWeight: 700,
                              cursor: 'pointer',
                              color: activePortfolioForEmployees === row.portfolio
                                ? 'primary.main'
                                : 'text.primary',
                              textDecoration:
                                activePortfolioForEmployees === row.portfolio
                                  ? 'underline'
                                  : 'none',
                            }}
                            onClick={() => setActivePortfolioForEmployees(row.portfolio)}
                          >
                            {row.portfolio}
                          </TableCell>

                          {COURSES.map((course) => {
                            const previousValue = row.counts[course.key].previous
                            const currentValue = row.counts[course.key].current
                            const styles = buildDeltaStyles(previousValue, currentValue)

                            return (
                              <Fragment key={`${row.portfolio}-${course.key}`}>
                                <TableCell
                                  key={`${row.portfolio}-${course.key}-prev`}
                                  align="right"
                                >
                                  {previousValue === null ? '-' : previousValue}
                                </TableCell>
                                <TableCell
                                  key={`${row.portfolio}-${course.key}-curr`}
                                  align="right"
                                  sx={{ color: styles.color, fontWeight: styles.weight }}
                                >
                                  {currentValue}
                                  {styles.trend && styles.trend !== '0' && (
                                    <Typography
                                      component="span"
                                      sx={{
                                        ml: 0.8,
                                        fontSize: '0.75rem',
                                        color: styles.color,
                                      }}
                                    >
                                      ({styles.trend})
                                    </Typography>
                                  )}
                                </TableCell>
                              </Fragment>
                            )
                          })}
                        </TableRow>
                      ))}
                    </TableBody>
                  </Table>
                </TableContainer>
                ) : (
                <TableContainer>
                  <Table size="small" stickyHeader>
                    <TableHead>
                      <TableRow>
                        <TableCell>Off Do</TableCell>
                        <TableCell>Portfolio</TableCell>
                        {dashboardData.hasPrevious && (
                          <>
                            <TableCell align="right">Total Employees (Prev)</TableCell>
                            <TableCell align="right">Total Employees (Curr)</TableCell>
                            <TableCell align="right">Total Defaulters (Prev)</TableCell>
                            <TableCell align="right">Total Defaulters (Curr)</TableCell>
                            <TableCell align="right">Defaulter % (Prev)</TableCell>
                            <TableCell align="right">Defaulter % (Curr)</TableCell>
                          </>
                        )}
                        {!dashboardData.hasPrevious && (
                          <>
                            <TableCell align="right">Total Employees</TableCell>
                            <TableCell align="right">Total Defaulters</TableCell>
                            <TableCell align="right">Defaulter %</TableCell>
                          </>
                        )}
                      </TableRow>
                    </TableHead>
                    <TableBody>
                      {portfolioSummaryRows.map((row, index) => {
                        const groupMeta = offDoGroupInfo[index]
                        const percentStyles = buildDeltaStyles(
                          row.previousPercent,
                          row.currentPercent,
                        )
                        const defaulterStyles = buildDeltaStyles(
                          row.previousDefaulters,
                          row.currentDefaulters,
                        )

                        return (
                          <TableRow key={row.portfolio} hover>
                            {groupMeta?.isFirstInGroup && (
                              <TableCell
                                rowSpan={groupMeta.groupSize}
                                sx={{
                                  fontWeight: 700,
                                  verticalAlign: 'top',
                                  backgroundColor: 'rgba(15, 118, 110, 0.06)',
                                  borderRight: '1px solid rgba(0,0,0,0.08)',
                                  whiteSpace: 'nowrap',
                                }}
                              >
                                {row.offDo || '-'}
                              </TableCell>
                            )}
                            <TableCell
                              sx={{
                                fontWeight: 700,
                                cursor: 'pointer',
                                color:
                                  activePortfolioForEmployees === row.portfolio
                                    ? 'primary.main'
                                    : 'text.primary',
                                textDecoration:
                                  activePortfolioForEmployees === row.portfolio
                                    ? 'underline'
                                    : 'none',
                              }}
                              onClick={() =>
                                setActivePortfolioForEmployees(row.portfolio)
                              }
                            >
                              {row.portfolio}
                            </TableCell>

                            {dashboardData.hasPrevious ? (
                              <>
                                <TableCell align="right">
                                  {row.previousTotal === null ? '-' : row.previousTotal}
                                </TableCell>
                                <TableCell align="right" sx={{ fontWeight: 600 }}>
                                  {row.currentTotal}
                                </TableCell>
                                <TableCell align="right">
                                  {row.previousDefaulters === null
                                    ? '-'
                                    : row.previousDefaulters}
                                </TableCell>
                                <TableCell
                                  align="right"
                                  sx={{
                                    color: defaulterStyles.color,
                                    fontWeight: defaulterStyles.weight,
                                  }}
                                >
                                  {row.currentDefaulters}
                                  {defaulterStyles.trend &&
                                    defaulterStyles.trend !== '0' && (
                                      <Typography
                                        component="span"
                                        sx={{
                                          ml: 0.8,
                                          fontSize: '0.75rem',
                                          color: defaulterStyles.color,
                                        }}
                                      >
                                        ({defaulterStyles.trend})
                                      </Typography>
                                    )}
                                </TableCell>
                                <TableCell align="right">
                                  {row.previousPercent === null
                                    ? '-'
                                    : `${row.previousPercent.toFixed(2)}%`}
                                </TableCell>
                                <TableCell
                                  align="right"
                                  sx={{
                                    color: percentStyles.color,
                                    fontWeight: percentStyles.weight,
                                  }}
                                >
                                  {`${row.currentPercent.toFixed(2)}%`}
                                </TableCell>
                              </>
                            ) : (
                              <>
                                <TableCell align="right" sx={{ fontWeight: 600 }}>
                                  {row.currentTotal}
                                </TableCell>
                                <TableCell align="right" sx={{ fontWeight: 600 }}>
                                  {row.currentDefaulters}
                                </TableCell>
                                <TableCell align="right" sx={{ fontWeight: 600 }}>
                                  {`${row.currentPercent.toFixed(2)}%`}
                                </TableCell>
                              </>
                            )}
                          </TableRow>
                        )
                      })}
                    </TableBody>
                  </Table>
                </TableContainer>
                )}
              </Paper>

              <Paper className="data-table-panel" elevation={0}>
                <Stack
                  direction={{ xs: 'column', md: 'row' }}
                  spacing={1.5}
                  justifyContent="space-between"
                  alignItems={{ xs: 'flex-start', md: 'center' }}
                  mb={1.5}
                >
                  <Typography variant="h6">
                    {activePortfolioForEmployees
                      ? `${activePortfolioForEmployees} - Employee Details`
                      : 'Portfolio Employee Details'}
                  </Typography>
                  <Typography variant="body2" color="text.secondary">
                    {activePortfolioForEmployees
                      ? `Showing ${portfolioEmployeeRows.length} employee(s)`
                      : 'Click a portfolio in the matrix to view employees'}
                  </Typography>
                </Stack>

                {!activePortfolioForEmployees ? (
                  <Typography color="text.secondary" sx={{ py: 2 }}>
                    Select any portfolio name in the "Portfolio Defaulter Matrix" table above to
                    load employee-wise defaulter details.
                  </Typography>
                ) : (
                <TableContainer>
                  <Table size="small" stickyHeader>
                    <TableHead>
                      <TableRow>
                        <TableCell>Employee ID</TableCell>
                        <TableCell>Employee Name</TableCell>
                        <TableCell>
                          {dashboardData.hasPrevious ? 'Portfolio (Prev/Curr)' : 'Portfolio'}
                        </TableCell>
                        {COURSES.map((course) => (
                          <TableCell key={`${course.key}-search`} align="center">
                            {course.label} (Prev/Curr)
                          </TableCell>
                        ))}
                      </TableRow>
                    </TableHead>
                    <TableBody>
                      {portfolioEmployeeRows.map((employee) => (
                        <TableRow key={employee.employeeId}>
                          <TableCell>{employee.employeeId}</TableCell>
                          <TableCell>{employee.employeeName || '-'}</TableCell>
                          <TableCell>
                            {dashboardData.hasPrevious ? (
                              <Stack spacing={0.4}>
                                <Typography variant="body2">
                                  P: {employee.portfolios.previous}
                                </Typography>
                                <Typography variant="body2">
                                  C: {employee.portfolios.current}
                                </Typography>
                              </Stack>
                            ) : (
                              employee.portfolio
                            )}
                          </TableCell>
                          {COURSES.map((course) => (
                            <TableCell key={`${employee.employeeId}-${course.key}`} align="center">
                              <Stack
                                direction="row"
                                spacing={0.7}
                                justifyContent="center"
                                flexWrap="wrap"
                              >
                                <Chip
                                  size="small"
                                  label={`P: ${employee.statuses[course.key].previous}`}
                                  color={
                                    employee.statuses[course.key].previous === 'Defaulter'
                                      ? 'error'
                                      : 'success'
                                  }
                                  variant={
                                    employee.statuses[course.key].previous === '-'
                                      ? 'outlined'
                                      : 'filled'
                                  }
                                />
                                <Chip
                                  size="small"
                                  label={`C: ${employee.statuses[course.key].current}`}
                                  color={
                                    employee.statuses[course.key].current === 'Defaulter'
                                      ? 'error'
                                      : 'success'
                                  }
                                  variant={
                                    employee.statuses[course.key].current === '-'
                                      ? 'outlined'
                                      : 'filled'
                                  }
                                />
                              </Stack>
                            </TableCell>
                          ))}
                        </TableRow>
                      ))}
                    </TableBody>
                  </Table>
                </TableContainer>
                )}
              </Paper>
            </Stack>
          )}
        </Container>

        {dashboardData && (
          <Dialog
            open={missingHcDialogOpen}
            onClose={() => setMissingHcDialogOpen(false)}
            maxWidth="sm"
            fullWidth
            PaperProps={{ sx: { borderRadius: 3 } }}
          >
            <DialogTitle>
              <Stack direction="row" alignItems="center" justifyContent="space-between">
                <Stack direction="row" spacing={1.2} alignItems="center">
                  <ReportProblemRoundedIcon color="warning" />
                  <Typography variant="h6">Missing From Headcount</Typography>
                </Stack>
                <IconButton size="small" onClick={() => setMissingHcDialogOpen(false)}>
                  <CloseRoundedIcon fontSize="small" />
                </IconButton>
              </Stack>
            </DialogTitle>

            <DialogContent dividers>
              {dashboardData.hasPrevious && (
                <Stack direction="row" spacing={1} mb={2}>
                  <Chip
                    label={`Current (${dashboardData.missingFromHc.currentIds.length})`}
                    color={missingHcDialogPeriod === 'current' ? 'warning' : 'default'}
                    variant={missingHcDialogPeriod === 'current' ? 'filled' : 'outlined'}
                    onClick={() => setMissingHcDialogPeriod('current')}
                    clickable
                  />
                  <Chip
                    label={`Previous (${dashboardData.missingFromHc.previousIds.length})`}
                    color={missingHcDialogPeriod === 'previous' ? 'secondary' : 'default'}
                    variant={missingHcDialogPeriod === 'previous' ? 'filled' : 'outlined'}
                    onClick={() => setMissingHcDialogPeriod('previous')}
                    clickable
                  />
                </Stack>
              )}

              {(() => {
                const ids =
                  missingHcDialogPeriod === 'current'
                    ? dashboardData.missingFromHc.currentIds
                    : dashboardData.missingFromHc.previousIds

                if (!ids.length) {
                  return (
                    <Typography color="text.secondary" sx={{ py: 2, textAlign: 'center' }}>
                      No employees missing from Headcount for this period.
                    </Typography>
                  )
                }

                return (
                  <>
                    <Typography variant="body2" color="text.secondary" mb={1.5}>
                      {ids.length} employee{ids.length !== 1 ? 's' : ''} found in course files but
                      not in the Headcount file.
                    </Typography>
                    <TableContainer sx={{ maxHeight: 420 }}>
                      <Table size="small" stickyHeader>
                        <TableHead>
                          <TableRow>
                            <TableCell>#</TableCell>
                            <TableCell>Employee ID</TableCell>
                          </TableRow>
                        </TableHead>
                        <TableBody>
                          {ids.map((id, index) => (
                            <TableRow key={id} hover>
                              <TableCell sx={{ color: 'text.secondary', width: 48 }}>
                                {index + 1}
                              </TableCell>
                              <TableCell sx={{ fontWeight: 600, fontFamily: 'monospace' }}>
                                {id}
                              </TableCell>
                            </TableRow>
                          ))}
                        </TableBody>
                      </Table>
                    </TableContainer>
                  </>
                )
              })()}
            </DialogContent>
          </Dialog>
        )}

        {dashboardData && (
          <Dialog
            open={portfolioDialogOpen}
            onClose={() => setPortfolioDialogOpen(false)}
            maxWidth="md"
            fullWidth
            PaperProps={{ sx: { borderRadius: 3 } }}
          >
            <DialogTitle>
              <Stack direction="row" alignItems="center" justifyContent="space-between">
                <Stack direction="row" spacing={1.2} alignItems="center">
                  <BusinessCenterRoundedIcon color="primary" />
                  <Typography variant="h6">Defaulter Count by Portfolio</Typography>
                </Stack>
                <IconButton size="small" onClick={() => setPortfolioDialogOpen(false)}>
                  <CloseRoundedIcon fontSize="small" />
                </IconButton>
              </Stack>
            </DialogTitle>

            <DialogContent dividers>
              {dashboardData.hasPrevious && (
                <Stack direction="row" spacing={1} mb={2}>
                  <Chip
                    label="Current"
                    color={portfolioDialogPeriod === 'current' ? 'primary' : 'default'}
                    variant={portfolioDialogPeriod === 'current' ? 'filled' : 'outlined'}
                    onClick={() => setPortfolioDialogPeriod('current')}
                    clickable
                  />
                  <Chip
                    label="Previous"
                    color={portfolioDialogPeriod === 'previous' ? 'secondary' : 'default'}
                    variant={portfolioDialogPeriod === 'previous' ? 'filled' : 'outlined'}
                    onClick={() => setPortfolioDialogPeriod('previous')}
                    clickable
                  />
                </Stack>
              )}

              {(() => {
                const period = portfolioDialogPeriod
                const rows = dashboardData.portfolioRows
                  .map((row) => {
                    const total = COURSES.reduce((sum, course) => {
                      const val = row.counts[course.key][period]
                      return sum + (typeof val === 'number' ? val : 0)
                    }, 0)
                    const perCourse = COURSES.map((course) => ({
                      label: course.label,
                      count: row.counts[course.key][period],
                    }))
                    return { portfolio: row.portfolio, total, perCourse }
                  })
                  .filter((r) => r.total > 0)
                  .sort((a, b) => b.total - a.total)

                if (!rows.length) {
                  return (
                    <Typography color="text.secondary" sx={{ py: 2, textAlign: 'center' }}>
                      No defaulters found for this period.
                    </Typography>
                  )
                }

                const grandTotal = rows.reduce((s, r) => s + r.total, 0)

                return (
                  <>
                    <Typography variant="body2" color="text.secondary" mb={1.5}>
                      {rows.length} portfolios with defaulters — {grandTotal} total defaulters
                    </Typography>
                    <TableContainer sx={{ maxHeight: 480 }}>
                      <Table size="small" stickyHeader>
                        <TableHead>
                          <TableRow>
                            <TableCell>#</TableCell>
                            <TableCell>Portfolio</TableCell>
                            {COURSES.map((course) => (
                              <TableCell key={course.key} align="right">
                                {course.label}
                              </TableCell>
                            ))}
                            <TableCell align="right" sx={{ fontWeight: 700 }}>
                              Total
                            </TableCell>
                          </TableRow>
                        </TableHead>
                        <TableBody>
                          {rows.map((row, index) => (
                            <TableRow key={row.portfolio} hover>
                              <TableCell sx={{ color: 'text.secondary', width: 40 }}>
                                {index + 1}
                              </TableCell>
                              <TableCell sx={{ fontWeight: 700 }}>
                                {row.portfolio}
                              </TableCell>
                              {row.perCourse.map((c) => (
                                <TableCell key={c.label} align="right">
                                  {c.count === null ? (
                                    <Typography variant="caption" color="text.disabled">—</Typography>
                                  ) : c.count > 0 ? (
                                    <Chip
                                      label={c.count}
                                      size="small"
                                      color="error"
                                      variant="filled"
                                      sx={{ fontWeight: 700, minWidth: 36 }}
                                    />
                                  ) : (
                                    <Typography variant="caption" color="success.main" sx={{ fontWeight: 600 }}>0</Typography>
                                  )}
                                </TableCell>
                              ))}
                              <TableCell align="right">
                                <Chip
                                  label={row.total}
                                  size="small"
                                  color="error"
                                  variant="outlined"
                                  sx={{ fontWeight: 700, minWidth: 40 }}
                                />
                              </TableCell>
                            </TableRow>
                          ))}
                        </TableBody>
                      </Table>
                    </TableContainer>
                  </>
                )
              })()}
            </DialogContent>
          </Dialog>
        )}
      </Box>
    </ThemeProvider>
  )
}

export default App
