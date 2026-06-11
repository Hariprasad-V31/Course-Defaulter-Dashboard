import { Fragment, useEffect, useMemo, useRef, useState } from 'react'
import * as XLSX from 'xlsx'
import {
  Alert,
  Box,
  Button,
  Chip,
  Container,
  Divider,
  FormControl,
  Grid,
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

  const lowerPortfolio = portfolio.toLowerCase()
  const searchable = lowerPortfolio
    .replace(/&/g, ' and ')
    .replace(/[^a-z0-9]+/g, ' ')
    .trim()
  const compact = searchable.replace(/\s+/g, '')

  if (compact === 'bps' || compact.startsWith('bps')) {
    return 'BPS'
  }

  if (lowerPortfolio.includes('c&h') || compact.startsWith('candh')) {
    return 'C&H'
  }

  if (searchable.includes('customer')) {
    return 'Customer Engagement'
  }

  if (compact === 'data' || compact === 'datacentre' || compact === 'datacenter') {
    return 'Data & Datacentre'
  }

  if (compact === 'dwa' || compact === 'dws') {
    return 'DWA/DWS'
  }

  if (searchable.startsWith('retail')) {
    return 'Retail'
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

async function parseRowsFromFile(file) {
  const arrayBuffer = await file.arrayBuffer()
  const workbook = XLSX.read(arrayBuffer, { type: 'array', cellDates: false })
  const firstSheetName = workbook.SheetNames[0]
  if (!firstSheetName) {
    return []
  }

  const worksheet = workbook.Sheets[firstSheetName]
  return XLSX.utils.sheet_to_json(worksheet, {
    defval: '',
    raw: false,
  })
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
    }

    return map
  }, {})
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

    return sortPortfoliosByPriority(rows.map((r) => r.portfolio)).map((portfolio) =>
      rows.find((r) => r.portfolio === portfolio),
    )
  }, [dashboardData, selectedPortfolio])

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

  const filteredEmployeeRows = useMemo(() => {
    if (!employeePool.length) {
      return []
    }

    if (!debouncedSearchQuery) {
      return employeePool.slice(0, 50)
    }

    return employeePool
      .filter((employee) => {
        const byId = employee.employeeId
          .toLowerCase()
          .includes(normalizedSearchId)
        const byName = employee.employeeName
          .toLowerCase()
          .includes(normalizedSearch)
        return byId || byName
      })
      .slice(0, MAX_EMPLOYEE_RESULTS)
  }, [employeePool, normalizedSearch, normalizedSearchId, debouncedSearchQuery])

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

  const hasMoreSearchResults = useMemo(() => {
    if (!debouncedSearchQuery) {
      return employeePool.length > 50
    }

    const totalMatches = employeePool.filter((employee) => {
      const byId = employee.employeeId
        .toLowerCase()
        .includes(normalizedSearchId)
      const byName = employee.employeeName
        .toLowerCase()
        .includes(normalizedSearch)
      return byId || byName
    }).length

    return totalMatches > filteredEmployeeRows.length
  }, [employeePool, filteredEmployeeRows.length, normalizedSearch, normalizedSearchId, debouncedSearchQuery])



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
      setErrorText('HC file must be an Excel file (.xlsx, .xls, .xlsm).')
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
      setErrorText('Upload the Current HC Count file before processing.')
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
      const currentHcMap = await parseHcMap(hcFiles.current, 'Current HC Count file')
      const hcMaps = {
        current: currentHcMap,
        previous:
          hasPrevious && hcFiles.previous
            ? await parseHcMap(hcFiles.previous, 'Previous HC Count file')
            : hasPrevious
              ? currentHcMap
              : {},
      }
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
          throw new Error(
            `Unable to detect required columns in ${item.file.name}. Ensure Employee ID and status columns are present.`,
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

          return { portfolio, counts }
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
        },
        processedAt: new Date().toISOString(),
      })

      setSelectedPortfolio('All')
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
      <Box className="dashboard-shell">
        <Container maxWidth="xl" className="soft-rise">
          <Paper className="app-header" elevation={0}>
            <Stack
              direction={{ xs: 'column', md: 'row' }}
              spacing={2}
              alignItems={{ xs: 'flex-start', md: 'center' }}
              justifyContent="space-between"
            >
              <Stack direction="row" spacing={1.5} alignItems="center">
                <Box className="header-mark">
                  <AnalyticsRoundedIcon />
                </Box>
                <Box>
                  <Typography variant="h3" className="page-title">
                    Course Defaulter Portfolio Dashboard
                  </Typography>
                </Box>
              </Stack>
              <Chip
                label="Excel Portfolio Analytics"
                color="primary"
                variant="outlined"
                className="header-chip"
              />
            </Stack>
          </Paper>

          <Grid container spacing={2.2} mb={2.5}>
            <Grid size={{ xs: 12, md: 4 }}>
              <Paper className="upload-card upload-card--current" elevation={0}>
                <Stack direction="row" alignItems="center" justifyContent="space-between" mb={1.4}>
                  <Typography variant="h6">Current Data</Typography>
                  <Chip
                    size="small"
                    label={`${COURSES.filter((course) => currentFiles[course.key]).length}/4`}
                    color={COURSES.every((course) => currentFiles[course.key]) ? 'success' : 'default'}
                    variant="outlined"
                  />
                </Stack>
                <Stack spacing={1.3}>
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
                      className={`file-button ${
                        currentFiles[course.key] ? 'file-button--selected' : ''
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
              <Paper className="upload-card upload-card--previous" elevation={0}>
                <Stack direction="row" alignItems="center" justifyContent="space-between" mb={1.4}>
                  <Typography variant="h6">Previous Data</Typography>
                  <Chip
                    size="small"
                    label={`${COURSES.filter((course) => previousFiles[course.key]).length}/4`}
                    color={COURSES.some((course) => previousFiles[course.key]) ? 'secondary' : 'default'}
                    variant="outlined"
                  />
                </Stack>
                <Stack spacing={1.3}>
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
                      className={`file-button ${
                        previousFiles[course.key] ? 'file-button--selected' : ''
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
              <Paper className="upload-card upload-card--hc" elevation={0}>
                <Stack direction="row" alignItems="center" justifyContent="space-between" mb={1.4}>
                  <Typography variant="h6">HC Count</Typography>
                  <Chip
                    size="small"
                    label={`${Number(Boolean(hcFiles.current)) + Number(Boolean(hcFiles.previous))}/2`}
                    color={hcFiles.current ? 'warning' : 'default'}
                    variant="outlined"
                  />
                </Stack>

                <Stack spacing={1.3}>
                  <Button
                    variant="outlined"
                    component="label"
                    color="warning"
                    startIcon={
                      hcFiles.current ? <CheckCircleRoundedIcon /> : <UploadFileRoundedIcon />
                    }
                    className={`file-button ${hcFiles.current ? 'file-button--selected' : ''}`}
                    sx={{ width: '100%', justifyContent: 'space-between' }}
                  >
                    <span>HC Count (Current)</span>
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
                    className={`file-button ${hcFiles.previous ? 'file-button--selected' : ''}`}
                    sx={{ width: '100%', justifyContent: 'space-between' }}
                  >
                    <span>HC Count (Previous)</span>
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

                <Box sx={{ mt: 1.5 }}>
                  <Typography variant="body2" color="text.secondary">
                    Current HC: {hcFiles.current ? hcFiles.current.name : 'Not selected'}
                  </Typography>
                  <Typography variant="body2" color="text.secondary">
                    Previous HC: {hcFiles.previous ? hcFiles.previous.name : 'Not selected'}
                  </Typography>
                </Box>

                <Divider sx={{ my: 1.6 }} />

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
            <Stack spacing={2.4} className="soft-rise">
              <Grid container spacing={1.6}>
                <Grid size={{ xs: 12, sm: 6, lg: 2.4 }}>
                  <Paper className="metric-card" elevation={0}>
                    <Stack direction="row" alignItems="center" justifyContent="space-between">
                      <Box>
                        <Typography color="text.secondary">Defaulter Portfolios</Typography>
                        <Typography variant="h6">{stats.portfolios}</Typography>
                      </Box>
                      <Box className="metric-icon metric-icon--teal">
                        <BusinessCenterRoundedIcon fontSize="small" />
                      </Box>
                    </Stack>
                  </Paper>
                </Grid>
                <Grid size={{ xs: 12, sm: 6, lg: 2.4 }}>
                  <Paper className="metric-card" elevation={0}>
                    <Stack direction="row" alignItems="center" justifyContent="space-between">
                      <Box>
                        <Typography color="text.secondary">HC Employees</Typography>
                        <Typography variant="h6">{stats.employees}</Typography>
                      </Box>
                      <Box className="metric-icon metric-icon--blue">
                        <GroupsRoundedIcon fontSize="small" />
                      </Box>
                    </Stack>
                  </Paper>
                </Grid>
                <Grid size={{ xs: 12, sm: 6, lg: 2.4 }}>
                  <Paper className="metric-card" elevation={0}>
                    <Stack direction="row" alignItems="center" justifyContent="space-between">
                      <Box>
                        <Typography color="text.secondary">
                          Current Defaulters (Any Course)
                        </Typography>
                        <Typography variant="h6">{stats.defaultersNow}</Typography>
                      </Box>
                      <Box className="metric-icon metric-icon--amber">
                        <InsightsRoundedIcon fontSize="small" />
                      </Box>
                    </Stack>
                  </Paper>
                </Grid>
                <Grid size={{ xs: 12, sm: 6, lg: 2.4 }}>
                  <Paper className="metric-card metric-card--warning" elevation={0}>
                    <Stack direction="row" alignItems="center" justifyContent="space-between">
                      <Box>
                        <Typography color="text.secondary">Missing From HC</Typography>
                        <Typography variant="h6">{stats.missingFromHc.total}</Typography>
                        <Typography variant="caption" color="text.secondary">
                          Curr {stats.missingFromHc.current}
                          {dashboardData.hasPrevious
                            ? ` / Prev ${stats.missingFromHc.previous}`
                            : ''}
                        </Typography>
                      </Box>
                      <Box className="metric-icon metric-icon--orange">
                        <ReportProblemRoundedIcon fontSize="small" />
                      </Box>
                    </Stack>
                  </Paper>
                </Grid>
                <Grid size={{ xs: 12, sm: 6, lg: 2.4 }}>
                  <Paper className="metric-card" elevation={0}>
                    <Stack direction="row" alignItems="center" justifyContent="space-between">
                      <Box>
                        <Typography color="text.secondary">Mode</Typography>
                        <Typography variant="h6">
                          {dashboardData.hasPrevious
                            ? `Previous ${dashboardData.previousFileCount}/4`
                            : 'Current only'}
                        </Typography>
                      </Box>
                      <Box className="metric-icon metric-icon--slate">
                        <AnalyticsRoundedIcon fontSize="small" />
                      </Box>
                    </Stack>
                  </Paper>
                </Grid>
              </Grid>

              <Paper className="toolbar-card" elevation={0}>
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

                  <Button
                    variant="contained"
                    color="secondary"
                    onClick={exportDashboard}
                    startIcon={<FileDownloadRoundedIcon />}
                  >
                    Export to Excel
                  </Button>
                </Stack>
              </Paper>

              <Paper className="lookup-card" elevation={0}>
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

              <Paper className="table-card" elevation={0}>
                <Typography variant="h6" mb={1}>
                  Portfolio Defaulter Matrix
                </Typography>

                <TableContainer>
                  <Table size="small" stickyHeader>
                    <TableHead>
                      <TableRow>
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
                      {filteredPortfolioRows.map((row) => (
                        <TableRow key={row.portfolio} hover>
                          <TableCell sx={{ fontWeight: 700 }}>{row.portfolio}</TableCell>

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
              </Paper>

              <Paper className="table-card" elevation={0}>
                <Stack
                  direction={{ xs: 'column', md: 'row' }}
                  spacing={0.8}
                  justifyContent="space-between"
                  alignItems={{ xs: 'flex-start', md: 'center' }}
                  mb={1}
                >
                  <Typography variant="h6">Employee Search Results</Typography>
                  <Typography variant="body2" color="text.secondary">
                    Showing {filteredEmployeeRows.length}
                    {hasMoreSearchResults ? ' (limited for performance)' : ''}
                  </Typography>
                </Stack>

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
                      {filteredEmployeeRows.map((employee) => (
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
              </Paper>
            </Stack>
          )}
        </Container>
      </Box>
    </ThemeProvider>
  )
}

export default App
