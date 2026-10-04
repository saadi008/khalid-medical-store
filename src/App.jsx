console.log("STEP 12 APP LOADED")

import React, { useEffect, useMemo, useState } from 'react'
import * as XLSX from 'xlsx'
import { jsPDF } from 'jspdf'
import autoTable from 'jspdf-autotable'

import {
  saveCurrentSession,
  getCurrentSession,
  clearCurrentSession,
  saveTransaction,
  getTransactionsBySession,
  saveDailySession,
  getDailySessions,
  completeDailySession,
  authenticateUser,
  saveActiveUser,
  getActiveUser,
  clearActiveUser,
  addAuditLog,
  getAuditLogs,
  generateCustomerId,
  hasRegisteredUsers,
  createUser,
  getUsers,
  generateEncryptedBackupFile,
  previewEncryptedBackupRestore,
  confirmEncryptedBackupRestore,
  getBackupHistory,
} from './db'

/* --------------------------------
   UI Foundation
-------------------------------- */

const ui = {
  card:
    'bg-white border border-slate-200/80 rounded-2xl shadow-[0_12px_32px_rgba(15,23,42,0.07)]',

  cardHover:
    'bg-white border border-slate-200/80 rounded-2xl shadow-[0_12px_32px_rgba(15,23,42,0.07)] transition-all duration-200 hover:-translate-y-1 hover:border-blue-300 hover:shadow-[0_18px_42px_rgba(37,99,235,0.14)]',

  input:
    'w-full rounded-xl border border-slate-200 bg-white px-4 py-3 text-sm text-slate-900 shadow-[inset_0_1px_2px_rgba(15,23,42,0.04)] outline-none transition duration-200 placeholder:text-slate-400 hover:border-slate-300 focus:border-blue-500 focus:ring-4 focus:ring-blue-100/70 focus:shadow-[0_0_0_1px_rgba(59,130,246,0.10)]',

  inputCompact:
    'w-full rounded-xl border border-slate-200 bg-white px-3 py-2.5 text-sm text-slate-900 shadow-sm outline-none transition-all duration-200 placeholder:text-slate-400 hover:border-slate-300 focus:border-blue-500 focus:ring-4 focus:ring-blue-500/10',

  label:
    'block text-sm font-semibold text-slate-700 mb-2',

  primaryButton:
    'bg-gradient-to-b from-blue-500 to-blue-700 text-white rounded-xl px-4 py-3 font-semibold shadow-[0_8px_20px_rgba(37,99,235,0.24)] transition-all duration-200 hover:-translate-y-0.5 hover:from-blue-600 hover:to-blue-800 hover:shadow-[0_14px_28px_rgba(37,99,235,0.30)] active:translate-y-0 active:shadow-md focus:outline-none focus:ring-4 focus:ring-blue-100 disabled:opacity-50 disabled:cursor-not-allowed',

  secondaryButton:
    'bg-white border border-slate-200 text-slate-700 rounded-xl px-4 py-3 font-semibold shadow-sm transition-all duration-200 hover:-translate-y-0.5 hover:border-blue-300 hover:bg-blue-50/70 hover:text-blue-700 hover:shadow-md focus:outline-none focus:ring-4 focus:ring-blue-100',

  dangerButton:
    'bg-gradient-to-r from-red-500 to-rose-600 text-white rounded-xl px-4 py-3 font-semibold shadow-[0_8px_18px_rgba(239,68,68,0.18)] transition-all duration-200 hover:-translate-y-0.5 hover:from-red-600 hover:to-rose-700 hover:shadow-[0_12px_24px_rgba(239,68,68,0.24)] focus:outline-none focus:ring-4 focus:ring-red-500/10 disabled:opacity-50 disabled:cursor-not-allowed',

  mutedButton:
    'bg-slate-100 text-slate-700 rounded-xl px-4 py-3 font-semibold transition-all duration-200 hover:bg-slate-200 hover:text-slate-900 hover:-translate-y-0.5 focus:outline-none focus:ring-4 focus:ring-slate-100',

  section:
    'bg-white border border-slate-200/80 rounded-2xl shadow-[0_12px_32px_rgba(15,23,42,0.07)]',

  sectionHeader:
    'border-b border-slate-200/80',

  tableHeader:
    'bg-gradient-to-r from-slate-50 to-blue-50/70 text-xs font-bold uppercase tracking-wide text-slate-600',
}

/* --------------------------------
   Helpers
-------------------------------- */

function getCurrentDate() {
  return new Date().toISOString().split('T')[0]
}

function getCurrentTime() {
  return new Date().toISOString()
}

function createId() {
  if (
    typeof crypto !== 'undefined' &&
    typeof crypto.randomUUID === 'function'
  ) {
    return crypto.randomUUID()
  }

  return `${Date.now()}-${Math.random()
    .toString(36)
    .substring(2, 10)}`
}

function formatMoney(amount) {
  return `Rs. ${Number(amount || 0).toLocaleString()}`
}

function formatDate(dateString) {
  if (!dateString) return '-'

  const date = new Date(`${dateString}T00:00:00`)

  return date.toLocaleDateString('en-PK', {
    day: '2-digit',
    month: 'short',
    year: 'numeric',
  })
}

function formatMonth(monthValue) {
  if (!monthValue) return '-'

  const date = new Date(`${monthValue}-01T00:00:00`)

  return date.toLocaleDateString('en-PK', {
    month: 'long',
    year: 'numeric',
  })
}

function formatTime(dateString) {
  if (!dateString) return '-'

  return new Date(dateString).toLocaleTimeString('en-PK', {
    hour: '2-digit',
    minute: '2-digit',
  })
}

function getCategoryName(category) {
  const names = {
    medicine: 'Medicine',
    general: 'General Items',
    dispensing: 'Dispensing',
  }

  return names[category] || category
}

function getRefundedAmountForSale(sale, transactions) {
  if (!sale || !Array.isArray(transactions)) {
    return 0
  }

  return transactions
    .filter((transaction) => {
      if (transaction.type !== 'refund') {
        return false
      }

      // New refunds: exact original sale link
      if (transaction.originalSaleId) {
        return transaction.originalSaleId === sale.id
      }

      // Legacy refunds created before originalSaleId was added
      return (
        sale.customerRef &&
        transaction.customerRef === sale.customerRef
      )
    })
    .reduce(
      (total, transaction) =>
        total + Number(transaction.amount || 0),
      0
    )
}

/* --------------------------------
   Category Icon
-------------------------------- */

function CategoryIcon({ category, size = 22 }) {
  if (category === 'medicine') {
    return (
      <svg
        width={size}
        height={size}
        viewBox="0 0 24 24"
        fill="none"
        stroke="currentColor"
        strokeWidth="2"
        strokeLinecap="round"
        strokeLinejoin="round"
      >
        <path d="M10.5 20.5 3.5 13.5a4.95 4.95 0 0 1 7-7l7 7a4.95 4.95 0 0 1-7 7Z" />
        <path d="m8 8 8 8" />
      </svg>
    )
  }

  if (category === 'general') {
    return (
      <svg
        width={size}
        height={size}
        viewBox="0 0 24 24"
        fill="none"
        stroke="currentColor"
        strokeWidth="2"
        strokeLinecap="round"
        strokeLinejoin="round"
      >
        <path d="m21 8-9-5-9 5 9 5 9-5Z" />
        <path d="m3 8 9 5 9-5" />
        <path d="M3 8v8l9 5 9-5V8" />
        <path d="M12 13v8" />
      </svg>
    )
  }

  return (
    <svg
      width={size}
      height={size}
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth="2"
      strokeLinecap="round"
      strokeLinejoin="round"
    >
      <path d="M8 3h8" />
      <path d="M9 3v4l-2 3v8a3 3 0 0 0 3 3h4a3 3 0 0 0 3-3v-8l-2-3V3" />
      <path d="M7 10h10" />
      <path d="M10 14h4" />
    </svg>
  )
}

/* --------------------------------
   Refund Method Icon
-------------------------------- */
function RefundMethodIcon({ type, size = 20 }) {
  if (type === 'customerId') {
    return (
      <svg
        width={size}
        height={size}
        viewBox="0 0 24 24"
        fill="none"
        stroke="currentColor"
        strokeWidth="1.8"
        strokeLinecap="round"
        strokeLinejoin="round"
      >
        <path d="M20 21a8 8 0 0 0-16 0" />
        <circle cx="12" cy="7" r="4" />
      </svg>
    )
  }

  return (
    <svg
      width={size}
      height={size}
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth="1.8"
      strokeLinecap="round"
      strokeLinejoin="round"
    >
      <path d="M6 2h9l3 3v17H6z" />
      <path d="M14 2v4h4" />
      <path d="M9 10h6" />
      <path d="M9 14h6" />
      <path d="M9 18h4" />
    </svg>
  )
}

/* --------------------------------
   Main App
-------------------------------- */

export default function App() {
  const [isLoading, setIsLoading] = useState(true)
  const [currentUser, setCurrentUser] = useState(null)

  const [notification, setNotification] = useState({
    show: false,
    type: 'info',
    message: '',
  })

  function showNotification(type, message) {
    setNotification({
      show: true,
      type,
      message,
    })

    setTimeout(() => {
      setNotification({
        show: false,
        type: 'info',
        message: '',
      })
    }, 4000)
  }

  function hideNotification() {
    setNotification({
      show: false,
      type: 'info',
      message: '',
    })
  }

  /* --------------------------------
     Daily Session State
  -------------------------------- */

  const [isDayOpen, setIsDayOpen] = useState(false)
  const [isDayClosed, setIsDayClosed] = useState(false)

  const [sessionId, setSessionId] = useState(null)
  const [openingCash, setOpeningCash] = useState(0)
  const [openingCashInput, setOpeningCashInput] = useState('')
  const [sessionDate, setSessionDate] = useState('')

  const [sales, setSales] = useState({
    medicine: 0,
    general: 0,
    dispensing: 0,
  })

  const [transactions, setTransactions] = useState([])

  const [categorySaleAmounts, setCategorySaleAmounts] =
    useState({
      medicine: '',
      general: '',
      dispensing: '',
    })

  const [showRefundModal, setShowRefundModal] = useState(false)

  const [refundForm, setRefundForm] = useState({
    category: 'all',
    amount: '',
    reason: '',
    customerRef: '',
  })

  const [refundCustomerNumber, setRefundCustomerNumber] =
    useState('')
  const [refundSale, setRefundSale] = useState(null)
  const [showRefundDetails, setShowRefundDetails] = useState(false)
  const [refundSearchMethod, setRefundSearchMethod] = useState('saleHistory')
  const [saleHistoryFilters, setSaleHistoryFilters] = useState({
    date: '',
    category: 'all',
    amount: '',
    time: '',
  })

  const [saleHistoryResults, setSaleHistoryResults] =
    useState([])

  const [saleHistorySearched, setSaleHistorySearched] =
    useState(false)
  const [saleHistoryPage, setSaleHistoryPage] = useState(1)
  const [showClosingModal, setShowClosingModal] = useState(false)

  const [actualCash, setActualCash] = useState('')
  const [closingReason, setClosingReason] = useState('')

  const [closingData, setClosingData] = useState(null)

  const [dailyRecords, setDailyRecords] = useState([])

  const [selectedRecord, setSelectedRecord] = useState(null)
  const [currentView, setCurrentView] = useState('dashboard')
  const [mobileMenuOpen, setMobileMenuOpen] = useState(false)
  const [auditLogs, setAuditLogs] = useState([])
  const [transactionPage, setTransactionPage] = useState(1)

  /* --------------------------------
   Refund Modal Reset
-------------------------------- */

  useEffect(() => {
    if (showRefundModal) {
      setRefundSearchMethod('customerId')
      setRefundSale(null)
      setRefundCustomerNumber('')

      setRefundForm({
        category: 'all',
        amount: '',
        reason: '',
        customerRef: '',
      })

      setSaleHistoryFilters({
        date: '',
        category: 'all',
        amount: '',
        time: '',
      })

      setSaleHistoryResults([])
      setSaleHistorySearched(false)
      setSaleHistoryPage(1)
    }
  }, [showRefundModal])

  /* --------------------------------
     Calculations
  -------------------------------- */

  const grossSales =
    sales.medicine +
    sales.general +
    sales.dispensing

  const totalRefunds = transactions
    .filter(
      (transaction) => transaction.type === 'refund'
    )
    .reduce(
      (total, transaction) =>
        total + Number(transaction.amount || 0),
      0
    )

  const netSales = grossSales - totalRefunds

  const expectedClosingCash =
    openingCash + netSales

  /* --------------------------------
     Load Authentication + App Data
  -------------------------------- */

  useEffect(() => {
    async function initializeApp() {
      try {
        const activeUser = await getActiveUser()

        if (!activeUser) {
          setCurrentUser(null)
          setIsLoading(false)
          return
        }

        setCurrentUser(activeUser)

        await loadAppData()
      } catch (error) {
        console.error(
          'Failed to initialize application:',
          error
        )

        showNotification('error',
          'Unable to load local database. Please refresh the application.'
        )

        setIsLoading(false)
      }
    }

    async function loadAppData() {
      const [currentSession, history, logs] =
        await Promise.all([
          getCurrentSession(),
          getDailySessions(),
          getAuditLogs(),
        ])

      setDailyRecords(history)
      setAuditLogs(logs)

      if (currentSession) {
        setCurrentView('dashboard')

        const sessionTransactions =
          await getTransactionsBySession(
            currentSession.sessionId
          )

        setSessionId(currentSession.sessionId)
        setSessionDate(currentSession.date)
        setOpeningCash(
          Number(currentSession.openingCash || 0)
        )

        setSales(
          currentSession.sales || {
            medicine: 0,
            general: 0,
            dispensing: 0,
          }
        )

        setTransactions(sessionTransactions)

        setTransactionPage(1)

        setIsDayOpen(true)
        setIsDayClosed(false)
      } else {
        setCurrentView('dashboard')
        setIsDayOpen(false)
        setIsDayClosed(false)
      }

      setIsLoading(false)
    }

    initializeApp()
  }, [])

  /* --------------------------------
     Login
  -------------------------------- */

  async function handleLogin(username, password) {
    try {
      const user = await authenticateUser(
        username,
        password
      )

      if (!user) {
        return {
          success: false,
          message: 'Invalid username or password.',
        }
      }

      await saveActiveUser(user)

      setCurrentUser(user)

      await addAuditLog({
        userId: user.id,
        username: user.username,
        action: 'LOGIN',
        description: `User ${user.username} logged in successfully.`,
      })

      /*
       * Load existing daily session/history
       * after successful login.
       */
      const [currentSession, history, logs] =
        await Promise.all([
          getCurrentSession(),
          getDailySessions(),
          getAuditLogs(),
        ])

      setDailyRecords(history)
      setAuditLogs(logs)

      if (currentSession) {
        const sessionTransactions =
          await getTransactionsBySession(
            currentSession.sessionId
          )

        setSessionId(currentSession.sessionId)
        setSessionDate(currentSession.date)
        setOpeningCash(
          Number(currentSession.openingCash || 0)
        )

        setSales(
          currentSession.sales || {
            medicine: 0,
            general: 0,
            dispensing: 0,
          }
        )

        setTransactions(sessionTransactions)
        setTransactionPage(1)

        setIsDayOpen(true)
        setIsDayClosed(false)
      } else {
        setIsDayOpen(false)
        setIsDayClosed(false)
        setCurrentView('dashboard')
      }

      return {
        success: true,
      }
    } catch (error) {
      console.error('Login failed:', error)

      return {
        success: false,
        message:
          'Unable to login. Please try again.',
      }
    }
  }

  /* --------------------------------
     Logout
  -------------------------------- */

  async function handleLogout() {
    try {
      if (currentUser) {
        await addAuditLog({
          userId: currentUser.id,
          username: currentUser.username,
          action: 'LOGOUT',
          description: `User ${currentUser.username} logged out.`,
          sessionId: sessionId || null,
        })
      }
      await clearActiveUser()

      /*
       * Important:
       * Logout does NOT clear the current daily
       * session or financial records.
       */
      setCurrentUser(null)

      setMobileMenuOpen(false)
      setSelectedRecord(null)
      setShowSaleModal(false)
      setShowRefundModal(false)
      setShowClosingModal(false)
    } catch (error) {
      console.error('Logout failed:', error)

      showNotification('error',
        'Could not logout. Please try again.'
      )
    }
  }

  /* --------------------------------
     Start Daily Session
  -------------------------------- */

  async function handleStartDay() {
    const amount = Number(openingCashInput)

    if (
      !openingCashInput ||
      Number.isNaN(amount) ||
      amount < 0
    ) {
      showNotification('error', 'Please enter a valid opening cash amount.')
      return
    }

    if (!currentUser) {
      showNotification('error', 'Please login first.')
      return
    }

    try {
      const newSessionId = createId()
      const date = getCurrentDate()

      const newSession = {
        sessionId: newSessionId,
        date,
        openingCash: amount,
        sales: {
          medicine: 0,
          general: 0,
          dispensing: 0,
        },
        status: 'open',
        openedAt: getCurrentTime(),

        openedBy: {
          userId: currentUser.id,
          username: currentUser.username,
          name: currentUser.name,
          role: currentUser.role,
        },
      }

      await saveCurrentSession(newSession)

      await addAuditLog({
        userId: currentUser.id,
        username: currentUser.username,
        action: 'START_DAILY_SESSION',
        description: `User ${currentUser.username} started the daily session.`,
        sessionId: newSession.sessionId,
      })

      setSessionId(newSessionId)
      setOpeningCash(amount)
      setSessionDate(date)

      setSales({
        medicine: 0,
        general: 0,
        dispensing: 0,
      })

      setTransactions([])
      setTransactionPage(1)
      setClosingData(null)

      setActualCash('')
      setClosingReason('')

      setOpeningCashInput('')

      setIsDayOpen(true)
      setIsDayClosed(false)
      setCurrentView('dashboard')
    } catch (error) {
      console.error(error)

      showNotification('error',
        'Could not start the daily session. Please try again.'
      )
    }
  }

  /* --------------------------------
     Update Current Session
  -------------------------------- */

  async function persistCurrentSession(
    updatedSales = sales
  ) {
    if (!sessionId) return

    await saveCurrentSession({
      key: 'currentSession',
      sessionId,
      date: sessionDate,
      openingCash,
      sales: updatedSales,
      status: 'open',
      openedAt: getCurrentTime(),

      openedBy: {
        userId: currentUser?.id || null,
        username: currentUser?.username || '',
        name: currentUser?.name || '',
        role: currentUser?.role || '',
      },
    })
  }

  /* --------------------------------
     Add Sale
  -------------------------------- */

  async function handleAddSale(
    category = saleForm.category,
    saleAmount = saleForm.amount
  ) {
    const amount = Number(saleAmount)

    if (
      !saleAmount ||
      !Number.isFinite(amount) ||
      amount <= 0
    ) {
      showNotification('error', 'Please enter a valid sale amount.')
      return
    }

    if (!sessionId) {
      showNotification('error', 'Please start a daily session first.')
      return
    }

    if (!currentUser) {
      showNotification('error', 'Please login first.')
      return
    }

    try {
      const customerRef = await generateCustomerId()

      const transaction = {
        id: createId(),
        sessionId,
        type: 'sale',
        category: category,
        amount,
        reason: '',
        time: getCurrentTime(),
        date: getCurrentDate(),

        /* Step 12 transaction attribution */
        userId: currentUser.id,
        username: currentUser.username,
        userName: currentUser.name,
        userRole: currentUser.role,
        customerRef,
      }

      await saveTransaction(transaction)

      await addAuditLog({
        userId: currentUser.id,
        username: currentUser.username,
        action: 'ADD_SALE',
        description: `User ${currentUser.username} added a ${getCategoryName(category)} sale of ${formatMoney(amount)}. Customer ID: ${customerRef}`,
        sessionId: sessionId,
        transactionId: transaction.id,
      })

      const updatedSales = {
        ...sales,
        [category]: sales[category] + amount,
      }

      await persistCurrentSession(updatedSales)

      setSales(updatedSales)

      setTransactions((previous) => [
        transaction,
        ...previous,
      ])

      setTransactionPage(1)

      setCategorySaleAmounts((previous) => ({
        ...previous,
        [category]: '',
      }))
    } catch (error) {
      console.error(error)

      showNotification('error', 'Sale could not be saved.')
    }
  }

  async function handleCustomerIdSearch(customerNumber, category) {
    const rawCustomerId = String(customerNumber || '').trim()

    // Empty Customer ID
    if (!rawCustomerId) {
      setRefundSale(null)

      setRefundForm((previous) => ({
        ...previous,
        category: 'all',
        customerRef: '',
        amount: '',
        reason: '',
      }))

      return
    }

    try {
      const allTransactions = []

      /* Current active session transactions */
      allTransactions.push(...(transactions || []))

      /* Completed daily sessions */
      const completedSessions = await getDailySessions()

      for (const session of completedSessions || []) {
        if (
          session.sessionId &&
          session.sessionId !== sessionId
        ) {
          const sessionTransactions =
            await getTransactionsBySession(
              session.sessionId
            )

          allTransactions.push(...sessionTransactions)
        }
      }

      /*
        Normalize Customer ID input.
  
        Supported formats:
  
        CUST-001
        cust-001
        CUST-01
        1
        01
        001
        12
        012
        345
  
        Numeric input is matched against the numeric
        portion of the generated Customer ID.
      */
      const normalizeCustomerId = (value) => {
        const normalizedValue = String(value || '')
          .trim()
          .toUpperCase()

        if (!normalizedValue) {
          return ''
        }

        /*
          Full Customer ID:
          CUST-001 -> 1
          CUST-012 -> 12
          CUST-345 -> 345
        */
        const customerIdMatch =
          normalizedValue.match(/^CUST-(\d+)$/)

        if (customerIdMatch) {
          return String(
            Number(customerIdMatch[1])
          )
        }

        /*
          Numeric-only input:
          1 -> 1
          12 -> 12
          345 -> 345
        */
        if (/^\d+$/.test(normalizedValue)) {
          return String(
            Number(normalizedValue)
          )
        }

        /*
          For any other text, keep it as-is so that
          invalid input simply produces no match.
        */
        return normalizedValue
      }

      const searchedCustomerId =
        normalizeCustomerId(rawCustomerId)

      /*
        Find sales using Customer ID.
  
        Category "all" means all categories.
      */
      const matchingSales = allTransactions
        .filter((transaction) => {
          if (transaction.type !== 'sale') {
            return false
          }

          const transactionCustomerId =
            normalizeCustomerId(
              transaction.customerRef
            )

          if (
            transactionCustomerId !==
            searchedCustomerId
          ) {
            return false
          }

          if (
            category !== 'all' &&
            transaction.category !== category
          ) {
            return false
          }

          return true
        })
        .sort(
          (a, b) =>
            new Date(b.time) - new Date(a.time)
        )

      /*
        Customer IDs are generated uniquely,
        so the latest matching sale is selected.
      */
      const matchedSale =
        matchingSales[0] || null

      if (!matchedSale) {
        setRefundSale(null)

        setRefundForm((previous) => ({
          ...previous,
          customerRef: rawCustomerId,
          amount: '',
          reason: '',
        }))

        return
      }

      /*
        Sale found:
        automatically select the original sale
        and switch category to the sale's
        actual category.
      */
      setRefundSale(matchedSale)
      setRefundForm((previous) => ({
        ...previous,
        category: matchedSale.category,
        customerRef: matchedSale.customerRef || rawCustomerId,
        amount: '',
        reason: '',
      }))
    } catch (error) {
      console.error(
        'Customer ID search error:',
        error
      )

      setRefundSale(null)

      setRefundForm((previous) => ({
        ...previous,
        customerRef: rawCustomerId,
        amount: '',
        reason: '',
      }))

      showNotification(
        'error',
        'Could not search Customer ID.'
      )
    }
  }

  /* --------------------------------
     Add Refund
  -------------------------------- */
  async function handleSaleHistorySearch() {
    if (!saleHistoryFilters.date) {
      showNotification(
        'error',
        'Please select the sale date.'
      )
      return
    }

    try {
      const allTransactions = []

      /* Current active session transactions */
      allTransactions.push(
        ...(transactions || [])
      )

      /* Completed daily sessions */
      const completedSessions =
        await getDailySessions()

      for (const session of completedSessions || []) {
        if (
          session.sessionId &&
          session.sessionId !== sessionId
        ) {
          const sessionTransactions =
            await getTransactionsBySession(
              session.sessionId
            )

          allTransactions.push(
            ...sessionTransactions
          )
        }
      }

      console.log('SALE HISTORY DEBUG')
      console.log('Selected Search Date:', saleHistoryFilters.date)
      console.log('Current Session ID:', sessionId)
      console.log('All Transactions:', allTransactions)

      const saleTransactions =
        allTransactions.filter(
          (transaction) => {
            if (transaction.type !== 'sale') {
              return false
            }

            if (
              transaction.date !==
              saleHistoryFilters.date
            ) {
              return false
            }

            if (
              saleHistoryFilters.category !==
              'all' &&
              transaction.category !==
              saleHistoryFilters.category
            ) {
              return false
            }

            if (
              saleHistoryFilters.amount &&
              Number(transaction.amount) !==
              Number(
                saleHistoryFilters.amount
              )
            ) {
              return false
            }

            if (
              saleHistoryFilters.time
            ) {
              const saleTime =
                new Date(
                  transaction.time
                )

              const saleHours =
                String(
                  saleTime.getHours()
                ).padStart(2, '0')

              const saleMinutes =
                String(
                  saleTime.getMinutes()
                ).padStart(2, '0')

              const [
                filterHours,
                filterMinutes,
              ] =
                saleHistoryFilters.time
                  .split(':')
                  .map(Number)

              const saleTotalMinutes =
                Number(saleHours) * 60 +
                Number(saleMinutes)

              const filterTotalMinutes =
                filterHours * 60 +
                filterMinutes

              const difference =
                Math.abs(
                  saleTotalMinutes -
                  filterTotalMinutes
                )

              if (difference > 30) {
                return false
              }
            }

            return true
          }
        )

      const uniqueSales =
        Array.from(
          new Map(
            saleTransactions.map(
              (transaction) => [
                transaction.id,
                transaction,
              ]
            )
          ).values()
        )

      uniqueSales.sort(
        (a, b) =>
          new Date(b.time) -
          new Date(a.time)
      )

      setSaleHistoryResults(
        uniqueSales
      )

      setSaleHistoryPage(1)

      setSaleHistorySearched(true)

      if (uniqueSales.length === 0) {
        showNotification(
          'warning',
          'No matching sales found. Try changing the search filters.'
        )
      }
    } catch (error) {
      console.error(error)

      showNotification(
        'error',
        'Could not search sale history.'
      )
    }
  }

  async function handleAddRefund() {
    const amount = Number(refundForm.amount)
    if (refundSale) {
      const previousRefunds = getRefundedAmountForSale(
        refundSale,
        transactions
      )

      const remainingRefundable =
        Number(refundSale.amount || 0) -
        previousRefunds

      if (amount > remainingRefundable) {
        showNotification(
          'error',
          `Refund amount cannot exceed the remaining refundable amount of ${formatMoney(
            Math.max(remainingRefundable, 0)
          )}.`
        )
        return
      }
    }

    if (
      !refundForm.amount ||
      Number.isNaN(amount) ||
      amount <= 0
    ) {
      showNotification('error', 'Please enter a valid refund amount.')
      return
    }

    if (!refundForm.reason.trim()) {
      showNotification('error', 'Please enter the refund reason.')
      return
    }

    if (!refundForm.customerRef?.trim()) {
      showNotification('error', 'Please enter the Customer ID.')
      return
    }

    if (!refundSale) {
      showNotification('error', 'Please enter a valid Customer ID.')
      return
    }

    const previousRefunds =
      getRefundedAmountForSale(
        refundSale,
        transactions
      )

    const remainingRefundable =
      Number(refundSale.amount || 0) -
      previousRefunds

    if (amount > remainingRefundable) {
      showNotification('error',
        `Refund amount cannot be greater than the remaining refundable amount of ${formatMoney(
          remainingRefundable
        )}.`
      )
      return
    }

    if (amount > refundSale.amount) {
      showNotification('error',
        `Refund amount cannot be greater than the original sale amount of ${formatMoney(
          refundSale.amount
        )}.`
      )
      return
    }

    if (amount > grossSales) {
      showNotification('error',
        'Refund amount cannot be greater than total sales.'
      )
      return
    }

    if (!sessionId) {
      showNotification('error', 'Please start a daily session first.')
      return
    }

    if (!currentUser) {
      showNotification('error', 'Please login first.')
      return
    }

    try {
      const transaction = {
        id: createId(),
        sessionId,
        type: 'refund',
        category: refundForm.category,
        amount,
        reason: refundForm.reason.trim(),
        time: getCurrentTime(),

        /* Step 12 transaction attribution */
        userId: currentUser.id,
        username: currentUser.username,
        userName: currentUser.name,
        userRole: currentUser.role,
        customerRef: refundForm.customerRef.trim(),

        /* Step 2.4.5 original sale linking */
        originalSaleId: refundSale.id,
      }

      await saveTransaction(transaction)

      await addAuditLog({
        userId: currentUser.id,
        username: currentUser.username,
        action: 'ADD_REFUND',
        description: `User ${currentUser.username} added a ${getCategoryName(refundForm.category)} refund of ${formatMoney(amount)}. Customer ID: ${refundForm.customerRef.trim()}. Reason: ${refundForm.reason.trim()}`,
        sessionId: sessionId,
        transactionId: transaction.id,
      })

      setTransactions((previous) => [
        transaction,
        ...previous,
      ])

      setTransactionPage(1)

      setRefundForm({
        category: 'all',
        amount: '',
        reason: '',
        customerRef: '',
      })

      setRefundCustomerNumber('')
      setRefundSale(null)
      setShowRefundDetails(false)
      setRefundSearchMethod('saleHistory')

      setCurrentView('dashboard')

      showNotification(
        'success',
        'Refund saved successfully.'
      )
    } catch (error) {
      console.error(error)

      showNotification('error', 'Refund could not be saved.')
    }
  }

  /* --------------------------------
     Open Closing
  -------------------------------- */

  function handleOpenClosing() {
    if (transactions.length === 0) {
      const confirmClose = window.confirm(
        'No sales or refunds have been recorded today. Do you still want to close the day?'
      )

      if (!confirmClose) {
        return
      }
    }

    setActualCash('')
    setClosingReason('')
    setShowClosingModal(true)
  }

  /* --------------------------------
     Complete Closing
  -------------------------------- */

  async function handleCompleteClosing() {
    const actual = Number(actualCash)

    if (
      actualCash === '' ||
      Number.isNaN(actual) ||
      actual < 0
    ) {
      showNotification('error', 'Please enter a valid actual cash amount.')
      return
    }

    const difference =
      actual - expectedClosingCash

    if (
      difference !== 0 &&
      !closingReason.trim()
    ) {
      showNotification('error',
        'Please provide a reason for the cash difference.'
      )
      return
    }

    let status = 'Matched'

    if (difference < 0) {
      status = 'Shortage'
    }

    if (difference > 0) {
      status = 'Extra Cash'
    }

    const result = {
      actualCash: actual,
      expectedClosingCash,
      difference,
      status,
    }

    const dailyRecord = {
      id: sessionId,
      sessionId,
      date: sessionDate || getCurrentDate(),

      openingCash,

      sales: {
        medicine: sales.medicine,
        general: sales.general,
        dispensing: sales.dispensing,
      },

      grossSales,
      totalRefunds,
      netSales,

      expectedClosingCash,
      actualCash: actual,
      difference,

      status,

      reason:
        difference === 0
          ? ''
          : closingReason.trim(),

      closedAt: getCurrentTime(),

      /* Step 12 closing attribution */
      closedBy: currentUser
        ? {
          userId: currentUser.id,
          username: currentUser.username,
          name: currentUser.name,
          role: currentUser.role,
        }
        : null,
    }

    try {
      await completeDailySession(dailyRecord)

      await addAuditLog({
        userId: currentUser.id,
        username: currentUser.username,
        action: 'DAILY_CLOSING',
        description: `User ${currentUser.username} closed the daily session. Actual cash: ${formatMoney(actual)}, Expected cash: ${formatMoney(expectedClosingCash)}, Status: ${status}.`,
        sessionId: sessionId,
      })

      setDailyRecords((previous) => [
        dailyRecord,
        ...previous.filter(
          (record) => record.id !== dailyRecord.id
        ),
      ])

      setClosingData(result)

      setIsDayClosed(true)
      setIsDayOpen(false)

      setShowClosingModal(false)

      setActualCash('')
      setClosingReason('')
    } catch (error) {
      console.error(error)

      showNotification('error',
        'Daily closing could not be saved. Please try again.'
      )
    }
  }

  /* --------------------------------
     Start New Day
  -------------------------------- */

  async function handleStartNewDay() {
    try {
      await clearCurrentSession()

      await addAuditLog({
        userId: currentUser.id,
        username: currentUser.username,
        action: 'START_NEW_DAY',
        description: `User ${currentUser.username} started a new day.`,
        sessionId: sessionId,
      })

      setSessionId(null)
      setOpeningCash(0)
      setOpeningCashInput('')
      setSessionDate('')

      setSales({
        medicine: 0,
        general: 0,
        dispensing: 0,
      })

      setTransactions([])
      setTransactionPage(1)

      setClosingData(null)

      setActualCash('')
      setClosingReason('')

      setIsDayOpen(false)
      setIsDayClosed(false)
      setCurrentView('dashboard')
      setMobileMenuOpen(false)

      const history = await getDailySessions()
      setDailyRecords(history)

      const logs = await getAuditLogs()
      setAuditLogs(logs)

    } catch (error) {
      console.error(error)

      showNotification('error',
        'Could not prepare the new day.'
      )
    }
  }

  /* --------------------------------
     Refresh History
  -------------------------------- */

  async function refreshHistory() {
    try {
      const history = await getDailySessions()
      setDailyRecords(history)
    } catch (error) {
      console.error(error)
    }
  }

  /* --------------------------------
     Loading
  -------------------------------- */

  if (isLoading) {
    return (
      <div className="min-h-screen bg-slate-50 flex items-center justify-center">
        <div className="text-center">
          <div className="w-10 h-10 border-4 border-slate-200 border-t-slate-700 rounded-full animate-spin mx-auto mb-4" />

          <p className="text-slate-500 font-medium">
            Loading local database...
          </p>
        </div>
      </div>
    )
  }

  /* --------------------------------
     Authentication Gate
  -------------------------------- */

  if (!currentUser) {
    return (
      <LoginScreen
        onLogin={handleLogin}
      />
    )
  }

  /* --------------------------------
     User Permission
  -------------------------------- */

  const isAdmin = currentUser.role === 'admin'

  /* --------------------------------
   User Management Page
-------------------------------- */

  const userManagementPage = (
    <PageContainer
      title="User Management"
      subtitle="Manage staff accounts and system access."
    >
      <UserManagement
        showNotification={showNotification}
        currentUser={currentUser}
      />
    </PageContainer>
  )

  if (currentView === 'userManagement' && isAdmin) {
    return (
      <AppShell
        currentView={currentView}
        onNavigate={(view) => {
          if (
            (view === 'monthly' ||
              view === 'auditLogs' ||
              view === 'userManagement' ||
              view === 'backupRestore') &&
            !isAdmin
          ) {
            return
          }

          setCurrentView(view)
          setMobileMenuOpen(false)
        }}
        mobileMenuOpen={mobileMenuOpen}
        setMobileMenuOpen={setMobileMenuOpen}
        isDayOpen={isDayOpen}
        sessionDate={sessionDate}
        currentUser={currentUser}
        onLogout={handleLogout}
        isAdmin={isAdmin}
      >
        {notification.show && (
          <Notification
            type={notification.type}
            message={notification.message}
            onClose={hideNotification}
          />
        )}

        {userManagementPage}
      </AppShell>
    )
  }

  if (currentView === 'backupRestore' && isAdmin) {
    return (
      <AppShell
        currentView={currentView}
        onNavigate={(view) => {
          if (
            (view === 'monthly' ||
              view === 'auditLogs' ||
              view === 'userManagement' ||
              view === 'backupRestore') &&
            !isAdmin
          ) {
            return
          }

          setCurrentView(view)
          setMobileMenuOpen(false)
        }}
        isDayOpen={isDayOpen}
        currentUser={currentUser}
        onLogout={handleLogout}
        mobileMenuOpen={mobileMenuOpen}
        setMobileMenuOpen={setMobileMenuOpen}
      >
        <BackupRestore
          showNotification={showNotification}
        />
      </AppShell>
    )
  }

  if (currentView === 'dataExport' && isAdmin) {
    return (
      <AppShell
        currentView={currentView}
        onNavigate={(view) => {
          if (
            (view === 'monthly' ||
              view === 'auditLogs' ||
              view === 'userManagement' ||
              view === 'backupRestore' ||
              view === 'dataExport') &&
            !isAdmin
          ) {
            return
          }

          setCurrentView(view)
          setMobileMenuOpen(false)
        }}
        mobileMenuOpen={mobileMenuOpen}
        setMobileMenuOpen={setMobileMenuOpen}
        isDayOpen={isDayOpen}
        sessionDate={sessionDate}
        currentUser={currentUser}
        onLogout={handleLogout}
        isAdmin={isAdmin}
      >
        <DataExport />
      </AppShell>
    )
  }

  /* --------------------------------
     Start Day Screen
  -------------------------------- */

  if (!isDayOpen && !isDayClosed) {
    return (
      <AppShell

        currentView={currentView}
        onNavigate={(view) => {
          if (
            (view === 'monthly' ||
              view === 'auditLogs' ||
              view === 'userManagement' ||
              view === 'backupRestore') &&
            !isAdmin
          ) {
            return
          }

          setCurrentView(view)
          setMobileMenuOpen(false)
        }}
        mobileMenuOpen={mobileMenuOpen}
        setMobileMenuOpen={setMobileMenuOpen}
        isDayOpen={false}
        currentUser={currentUser}
        onLogout={handleLogout}
        isAdmin={isAdmin}
      >


        {currentView === 'refund' ? (
          <PageContainer>
            <div className="space-y-5">
              {/* Refund Page Header */}
              <div className="flex flex-col gap-4 sm:flex-row sm:items-center sm:justify-between">
                <div>
                  <div className="flex items-center gap-3">
                    <div className="flex h-11 w-11 items-center justify-center rounded-xl bg-blue-50 text-blue-600">
                      <svg
                        width="22"
                        height="22"
                        viewBox="0 0 24 24"
                        fill="none"
                        stroke="currentColor"
                        strokeWidth="2"
                        strokeLinecap="round"
                        strokeLinejoin="round"
                      >
                        <path d="M3 12h18" />
                        <path d="m12 3 9 9-9 9" />
                      </svg>
                    </div>

                    <div>
                      <h1 className="text-xl font-bold tracking-tight text-slate-900 sm:text-2xl">
                        Refund Management
                      </h1>
                      <p className="mt-1 text-sm text-slate-500">
                        Find the original sale and process a customer refund.
                      </p>
                    </div>
                  </div>
                </div>

                <button
                  type="button"
                  onClick={() => setCurrentView('dashboard')}
                  className={`${ui.secondaryButton} inline-flex items-center justify-center gap-2 px-4 py-2.5 text-sm`}
                >
                  <span>←</span>
                  Back to Dashboard
                </button>
              </div>

              {/* Refund Workspace Placeholder */}
              <div className={`${ui.card} min-h-[420px] p-5 sm:p-7`}>
                <div className="flex min-h-[380px] items-center justify-center">
                  <div className="max-w-md text-center">
                    <div className="mx-auto flex h-16 w-16 items-center justify-center rounded-2xl bg-blue-50 text-blue-600">
                      <svg
                        width="30"
                        height="30"
                        viewBox="0 0 24 24"
                        fill="none"
                        stroke="currentColor"
                        strokeWidth="2"
                        strokeLinecap="round"
                        strokeLinejoin="round"
                      >
                        <path d="M9 14 4 9l5-5" />
                        <path d="M4 9h10a6 6 0 0 1 6 6v1" />
                      </svg>
                    </div>

                    <h2 className="mt-5 text-lg font-bold text-slate-900">
                      Refund Workspace
                    </h2>

                    <p className="mt-2 text-sm leading-6 text-slate-500">
                      Customer ID and Sale History refund options will appear here.
                    </p>
                  </div>
                </div>
              </div>
            </div>
          </PageContainer>
        ) : currentView === 'recordDetails' && selectedRecord ? (
          <DailyRecordDetails
            record={selectedRecord}
            onBack={() => {
              setSelectedRecord(null)
              setCurrentView('records')
            }}
          />
        ) : currentView === 'records' ? (
          <PageContainer
            title="Daily Records"
            subtitle="View and search completed daily sessions."
          >
            <DailyHistory
              records={dailyRecords}
              onViewDetails={(record) => {
                setSelectedRecord(record)
                setCurrentView('recordDetails')
              }}
              onRefresh={refreshHistory}
            />
          </PageContainer>
        ) : currentView === 'auditLogs' && isAdmin ? (
          <PageContainer title="Audit Logs">
            <AuditLogsTable logs={auditLogs} />
          </PageContainer>
        ) : currentView === 'monthly' && isAdmin ? (
          <PageContainer
            title="Monthly Dashboard & Reports"
            subtitle="Review sales, cash reconciliation, and daily performance by month."
          >
            <MonthlyReports records={dailyRecords} />
          </PageContainer>
        ) : (
          <div className="max-w-2xl mx-auto">
            <div className="bg-white rounded-2xl shadow-sm border border-slate-200/80 p-6 sm:p-8">
              <div className="text-center mb-7">
                <div className="w-14 h-14 bg-blue-50 rounded-2xl flex items-center justify-center mx-auto mb-4 text-blue-600">
                  <svg
                    width="28"
                    height="28"
                    viewBox="0 0 24 24"
                    fill="none"
                    stroke="currentColor"
                    strokeWidth="2"
                  >
                    <path d="M3 10h18" />
                    <path d="M5 10v10h14V10" />
                    <path d="M4 10V6h16v4" />
                    <path d="M8 6V3h8v3" />
                  </svg>
                </div>

                <h2 className="text-xl sm:text-2xl font-bold text-slate-900">
                  Start Daily Session
                </h2>

                <p className="text-sm text-slate-500 mt-2 max-w-md mx-auto">
                  Enter the cash available in the box at the beginning of the day.
                </p>
              </div>

              <label className={ui.label}>
                Opening Cash
              </label>

              <div className="relative">
                <span className="absolute left-4 top-1/2 -translate-y-1/2 text-slate-500 font-medium">
                  Rs.
                </span>

                <input
                  type="number"
                  min="0"
                  value={openingCashInput}
                  onChange={(e) =>
                    setOpeningCashInput(
                      e.target.value
                    )
                  }
                  onKeyDown={(e) => {
                    if (e.key === 'Enter') {
                      handleStartDay()
                    }
                  }}
                  placeholder="1000"
                  className="w-full border border-slate-300 rounded-xl py-3.5 pl-12 pr-4 outline-none transition focus:border-blue-500 focus:ring-4 focus:ring-blue-50"
                />
              </div>

              <button
                onClick={handleStartDay}
                className={`${ui.primaryButton} w-full mt-5 py-3.5`}
              >
                Start Day
              </button>
            </div>
          </div>
        )}
      </AppShell>
    )
  }

  /* --------------------------------
     Closed Day Screen
  -------------------------------- */

  if (isDayClosed) {
    return (
      <AppShell
        currentView={currentView}
        onNavigate={(view) => {
          if (
            (view === 'monthly' ||
              view === 'auditLogs' ||
              view === 'userManagement' ||
              view === 'backupRestore') &&
            !isAdmin
          ) {
            return
          }

          setCurrentView(view)
          setMobileMenuOpen(false)
        }}
        mobileMenuOpen={mobileMenuOpen}
        setMobileMenuOpen={setMobileMenuOpen}
        isDayOpen={false}
        currentUser={currentUser}
        onLogout={handleLogout}
        isAdmin={isAdmin}
      >
        {currentView === 'refund' ? (
          <PageContainer
            title="Refund Management"
            subtitle="Find a sale and process a customer refund."
          >
            <div className="space-y-5">
              <div className="flex flex-col gap-3 sm:flex-row sm:items-center sm:justify-between">
                <div>
                  <h2 className="text-lg font-bold text-slate-900">
                    Refund Management
                  </h2>

                  <p className="mt-1 text-sm text-slate-500">
                    Select how you want to find the original sale.
                  </p>
                </div>

                <button
                  type="button"
                  onClick={() => setCurrentView('dashboard')}
                  className={`${ui.secondaryButton} inline-flex w-fit items-center justify-center gap-2 px-4 py-2.5 text-sm`}
                >
                  <span>←</span>
                  Back to Dashboard
                </button>
              </div>

              <div className={`${ui.card} min-h-[420px] p-5 sm:p-7`}>
                <div className="grid grid-cols-1 gap-5 lg:grid-cols-[220px_minmax(0,1fr)]">
                  {/* Refund Methods */}
                  <div className={`${ui.card} self-start lg:sticky lg:top-6 p-3`}>
                    <div className="px-3 pb-3 pt-2">
                      <p className="text-xs font-semibold uppercase tracking-wider text-slate-400">
                        Refund Method
                      </p>
                    </div>

                    <div className="space-y-2">
                      <button
                        type="button"
                        onClick={() => setRefundSearchMethod('customerId')}
                        className={`w-full rounded-xl px-4 py-3 text-left transition ${refundSearchMethod === 'customerId'
                          ? 'bg-blue-50 text-blue-700 shadow-sm ring-1 ring-blue-100'
                          : 'text-slate-600 hover:bg-slate-50'
                          }`}
                      >
                        <div className="flex items-center gap-3">
                          <span
                            className={`flex h-9 w-9 shrink-0 items-center justify-center rounded-lg ${refundSearchMethod === 'customerId'
                              ? 'bg-blue-100 text-blue-600'
                              : 'bg-slate-100 text-slate-500'
                              }`}
                          >
                            <RefundMethodIcon type="customerId" />
                          </span>

                          <div className="min-w-0">
                            <p className="text-sm font-semibold">
                              Customer ID
                            </p>
                            <p className="mt-0.5 text-xs text-slate-400">
                              Find by customer
                            </p>
                          </div>
                        </div>
                      </button>

                      <button
                        type="button"
                        onClick={() => setRefundSearchMethod('saleHistory')}
                        className={`w-full rounded-xl px-4 py-3 text-left transition ${refundSearchMethod === 'saleHistory'
                          ? 'bg-blue-50 text-blue-700 shadow-sm ring-1 ring-blue-100'
                          : 'text-slate-600 hover:bg-slate-50'
                          }`}
                      >
                        <div className="flex items-center gap-3">
                          <span
                            className={`flex h-9 w-9 shrink-0 items-center justify-center rounded-lg ${refundSearchMethod === 'saleHistory'
                              ? 'bg-blue-100 text-blue-600'
                              : 'bg-slate-100 text-slate-500'
                              }`}
                          >
                            <RefundMethodIcon type="saleHistory" />
                          </span>

                          <div className="min-w-0">
                            <p className="text-sm font-semibold">
                              Sale History
                            </p>
                            <p className="mt-0.5 text-xs text-slate-400">
                              Find by sale
                            </p>
                          </div>
                        </div>
                      </button>
                    </div>
                  </div>

                  {/* Refund Workspace */}
                  <div className={`${ui.card} min-w-0 p-5 sm:p-6`}>
                    <div className="border-b border-slate-100 pb-4">
                      <div className="flex items-center gap-3">
                        <div className="flex h-10 w-10 items-center justify-center rounded-xl bg-blue-50 text-blue-600">
                          <RefundMethodIcon
                            type={
                              refundSearchMethod === 'customerId'
                                ? 'customerId'
                                : 'saleHistory'
                            }
                          />
                        </div>

                        <div>
                          <h3 className="text-base font-bold text-slate-900 sm:text-lg">
                            {refundSearchMethod === 'customerId'
                              ? 'Customer ID'
                              : 'Sale History'}
                          </h3>

                          <p className="mt-0.5 text-xs text-slate-500 sm:text-sm">
                            {refundSearchMethod === 'customerId'
                              ? 'Find a customer sale using their Customer ID.'
                              : 'Search and select a previous sale to refund.'}
                          </p>
                        </div>
                      </div>
                    </div>

                    <div className="py-6">
                      <div className="rounded-xl border border-dashed border-slate-200 bg-slate-50/60 p-6 text-center">
                        <p className="text-sm font-medium text-slate-600">
                          {refundSearchMethod === 'customerId'
                            ? 'Customer ID workspace'
                            : 'Sale History workspace'}
                        </p>

                        <p className="mt-1 text-xs text-slate-400">
                          Existing refund controls will be moved here in the next step.
                        </p>
                      </div>
                    </div>
                  </div>
                </div>
              </div>
            </div>
          </PageContainer>
        ) : currentView === 'recordDetails' && selectedRecord ? (
          <DailyRecordDetails
            record={selectedRecord}
            onBack={() => {
              setSelectedRecord(null)
              setCurrentView('records')
            }}
          />
        ) : currentView === 'records' ? (
          <PageContainer
            title="Daily Records"
            subtitle="View and search completed daily sessions."
          >
            <DailyHistory
              records={dailyRecords}
              onViewDetails={(record) => {
                setSelectedRecord(record)
                setCurrentView('recordDetails')
              }}
              onRefresh={refreshHistory}
            />
          </PageContainer>
        ) : currentView === 'auditLogs' && isAdmin ? (
          <PageContainer
            title="Audit Logs"
          >
            <AuditLogsTable logs={auditLogs} />
          </PageContainer>
        ) : currentView === 'monthly' &&
          isAdmin ? (
          <PageContainer
            title="Monthly Dashboard & Reports"
            subtitle="Review sales, cash reconciliation, and daily performance by month."
          >
            <MonthlyReports
              records={dailyRecords}
            />
          </PageContainer>
        ) : (
          <div className="max-w-4xl mx-auto">
            <div className="bg-white border border-slate-200 rounded-2xl shadow-sm p-6 sm:p-8 text-center">
              <div className="w-16 h-16 bg-emerald-100 text-emerald-600 rounded-full flex items-center justify-center mx-auto mb-5">
                <svg
                  width="32"
                  height="32"
                  viewBox="0 0 24 24"
                  fill="none"
                  stroke="currentColor"
                  strokeWidth="2.5"
                >
                  <path d="m5 12 4 4L19 6" />
                </svg>
              </div>

              <h2 className="text-2xl font-bold tracking-tight text-slate-900">
                Day Closed Successfully
              </h2>

              <p className="text-slate-500 mt-2">
                The daily record has been saved to the local database.
              </p>

              {closingData && (
                <div className="grid grid-cols-1 sm:grid-cols-3 gap-4 mt-8 text-left">
                  <DetailCard
                    title="Expected Cash"
                    value={formatMoney(
                      closingData.expectedClosingCash
                    )}
                  />

                  <DetailCard
                    title="Actual Cash"
                    value={formatMoney(
                      closingData.actualCash
                    )}
                  />

                  <DetailCard
                    title="Difference"
                    value={formatMoney(
                      Math.abs(
                        closingData.difference
                      )
                    )}
                    status={
                      closingData.status
                    }
                  />
                </div>
              )}

              <div className="flex flex-col sm:flex-row gap-3 mt-8 justify-center">
                <button
                  onClick={handleStartNewDay}
                  className={`${ui.primaryButton} px-6`}
                >
                  Start New Day
                </button>

                <button
                  onClick={() => {
                    const record = dailyRecords.find((item) => item.id === sessionId)

                    if (record) {
                      setSelectedRecord(record)
                      setCurrentView('recordDetails')
                    }
                  }}
                  className={`${ui.secondaryButton} px-6`}
                >
                  View Today's Record
                </button>
              </div>
            </div>
          </div>
        )}
      </AppShell>
    )
  }

  /* --------------------------------
     Main Dashboard
  -------------------------------- */

  const transactionPageSize = 5

  const totalTransactionPages = Math.max(
    1,
    Math.ceil(
      transactions.length /
      transactionPageSize
    )
  )

  const visibleTransactions =
    transactions.slice(
      (transactionPage - 1) *
      transactionPageSize,
      transactionPage *
      transactionPageSize
    )

  const currentMonth = new Date()
    .toISOString()
    .slice(0, 7)

  const monthlyCustomerCount = transactions.filter(
    (transaction) =>
      transaction.type === 'sale' &&
      transaction.customerRef &&
      new Date(transaction.time)
        .toISOString()
        .slice(0, 7) === currentMonth
  ).length

  const recentRefundSales = (transactions || [])
    .filter((transaction) => transaction.type === 'sale')
    .sort((a, b) => new Date(b.time) - new Date(a.time))
    .slice(0, 10)

  return (
    <AppShell
      currentView={currentView}
      onNavigate={(view) => {
        if (
          (view === 'monthly' ||
            view === 'auditLogs' ||
            view === 'userManagement' ||
            view === 'backupRestore') &&
          !isAdmin
        ) {
          return
        }

        setCurrentView(view)
        setMobileMenuOpen(false)
      }}
      mobileMenuOpen={mobileMenuOpen}
      setMobileMenuOpen={setMobileMenuOpen}
      isDayOpen={true}
      sessionDate={sessionDate}
      currentUser={currentUser}
      onLogout={handleLogout}
      isAdmin={isAdmin}
    >
      {notification.show && (
        <Notification
          type={notification.type}
          message={notification.message}
          onClose={hideNotification}
        />
      )}
      {currentView === 'refund' ? (
        <PageContainer
          title="Refund Management"
          subtitle="Find a previous sale and process a refund."
        >
          <div className="space-y-5">
            {/* Refund Page Header */}
            <div className="flex flex-col gap-3 sm:flex-row sm:items-center sm:justify-between">
              <div className="min-w-0">
                <div className="flex items-center gap-3">
                  <div className="flex h-10 w-10 shrink-0 items-center justify-center rounded-xl bg-blue-50 text-blue-600 shadow-sm ring-1 ring-blue-100">
                    <RefundMethodIcon type="saleHistory" size={19} />
                  </div>

                  <div className="min-w-0">
                    <p className="text-[10px] font-bold uppercase tracking-[0.14em] text-blue-600">
                      Sale History
                    </p>
                    <h2 className="mt-0.5 text-lg font-bold tracking-tight text-slate-900 sm:text-xl">
                      Refund a Previous Sale
                    </h2>
                  </div>
                </div>
              </div>

              <button
                type="button"
                onClick={() => setCurrentView('dashboard')}
                className={`${ui.secondaryButton} inline-flex w-fit items-center justify-center gap-2 px-4 py-2.5 text-sm`}
              >
                <span>←</span>
                Back to Dashboard
              </button>
            </div>

            {/* Search Workspace */}
            <div className={`${ui.card} overflow-hidden p-0`}>
              <div className="border-b border-blue-100 bg-blue-50/50 px-4 py-3.5 sm:px-5">
                <div className="flex items-center gap-2.5">
                  <div className="flex h-8 w-8 shrink-0 items-center justify-center rounded-lg bg-white text-blue-600 shadow-sm ring-1 ring-blue-100">
                    <span className="text-sm font-bold">⌕</span>
                  </div>

                  <div>
                    <p className="text-sm font-bold text-slate-900">
                      Find Original Sale
                    </p>
                    <p className="mt-0.5 text-[11px] text-slate-500 sm:text-xs">
                      Sale date is required. The other filters help narrow the results.
                    </p>
                  </div>
                </div>
              </div>

              <div className="p-4 sm:p-5">
                <div className="grid grid-cols-1 gap-3 sm:grid-cols-2 lg:grid-cols-[1.15fr_1.15fr_1fr_1fr_auto] lg:items-end">
                  {/* Sale Date */}
                  <div className="min-w-0">
                    <label className={ui.label}>
                      Sale Date
                      <span className="ml-1 text-xs font-normal text-red-500">
                        Required
                      </span>
                    </label>

                    <input
                      type="date"
                      value={saleHistoryFilters.date}
                      onChange={(e) => {
                        setSaleHistoryFilters((previous) => ({
                          ...previous,
                          date: e.target.value,
                        }))
                        setSaleHistorySearched(false)
                        setSaleHistoryPage(1)
                      }}
                      className={ui.input}
                    />
                  </div>

                  {/* Category */}
                  <div className="min-w-0">
                    <label className={ui.label}>
                      Category
                    </label>

                    <select
                      value={saleHistoryFilters.category}
                      onChange={(e) => {
                        setSaleHistoryFilters((previous) => ({
                          ...previous,
                          category: e.target.value,
                        }))
                        setSaleHistorySearched(false)
                        setSaleHistoryPage(1)
                      }}
                      className={ui.input}
                    >
                      <option value="all">All Categories</option>
                      <option value="medicine">Medicine</option>
                      <option value="general">General Items</option>
                      <option value="dispensing">Dispensing</option>
                    </select>
                  </div>

                  {/* Sale Amount */}
                  <div className="min-w-0">
                    <label className={ui.label}>
                      Sale Amount
                    </label>

                    <div className="relative">
                      <span className="absolute left-4 top-1/2 -translate-y-1/2 text-xs font-bold text-slate-500">
                        Rs.
                      </span>

                      <input
                        type="number"
                        min="0"
                        value={saleHistoryFilters.amount}
                        onChange={(e) => {
                          setSaleHistoryFilters((previous) => ({
                            ...previous,
                            amount: e.target.value,
                          }))
                          setSaleHistorySearched(false)
                          setSaleHistoryPage(1)
                        }}
                        placeholder="e.g. 1200"
                        className={`${ui.input} pl-12`}
                      />
                    </div>
                  </div>

                  {/* Approximate Time */}
                  <div className="min-w-0">
                    <label className={ui.label}>
                      Approximate Time
                      <span className="ml-1 text-xs font-normal text-slate-400">
                        Optional
                      </span>
                    </label>

                    <input
                      type="time"
                      value={saleHistoryFilters.time}
                      onChange={(e) => {
                        setSaleHistoryFilters((previous) => ({
                          ...previous,
                          time: e.target.value,
                        }))
                        setSaleHistorySearched(false)
                        setSaleHistoryPage(1)
                      }}
                      className={ui.input}
                    />
                  </div>

                  {/* Find Button */}
                  <button
                    type="button"
                    onClick={handleSaleHistorySearch}
                    className={`${ui.primaryButton} inline-flex h-[46px] w-full items-center justify-center gap-2 px-5 lg:w-auto`}
                  >
                    <span className="text-base">⌕</span>
                    <span>Find</span>
                  </button>
                </div>

                <div className="mt-3 flex flex-col gap-1 text-[11px] text-slate-400 sm:flex-row sm:items-center sm:justify-between">
                  <span>Tip: Use amount or approximate time when several sales were made on the same date.</span>
                  <span>Time matches sales within approximately 30 minutes.</span>
                </div>
              </div>
            </div>

            {/* Search Results */}
            {saleHistorySearched && !showRefundDetails && (
              <div className={`${ui.card} overflow-hidden p-0`}>
                <div className="flex flex-col gap-2 border-b border-slate-100 bg-white px-4 py-3.5 sm:flex-row sm:items-center sm:justify-between sm:px-5">
                  <div>
                    <p className="text-[10px] font-bold uppercase tracking-[0.14em] text-slate-400">
                      Search Results
                    </p>
                    <h3 className="mt-0.5 text-sm font-bold text-slate-900 sm:text-base">
                      {saleHistoryResults.length > 0
                        ? `${saleHistoryResults.length} matching sale${saleHistoryResults.length === 1 ? '' : 's'} found`
                        : 'No matching sales found'}
                    </h3>
                  </div>

                  {saleHistoryResults.length > 0 && (
                    <span className="w-fit rounded-full bg-blue-50 px-2.5 py-1 text-[10px] font-bold text-blue-700 ring-1 ring-blue-100">
                      {saleHistoryResults.length} Result{saleHistoryResults.length === 1 ? '' : 's'}
                    </span>
                  )}
                </div>

                {saleHistoryResults.length > 0 ? (
                  <>
                    <div className="overflow-x-auto">
                      <table className="w-full min-w-[980px] border-collapse">
                        <thead>
                          <tr className="border-b border-slate-100 bg-slate-50/70">
                            <th className="px-4 py-3 text-left text-[9px] font-bold uppercase tracking-wide text-slate-500">Category</th>
                            <th className="px-3 py-3 text-left text-[9px] font-bold uppercase tracking-wide text-slate-500">Date</th>
                            <th className="px-3 py-3 text-left text-[9px] font-bold uppercase tracking-wide text-slate-500">Time</th>
                            <th className="px-3 py-3 text-right text-[9px] font-bold uppercase tracking-wide text-slate-500">Sale Amount</th>
                            <th className="px-3 py-3 text-left text-[9px] font-bold uppercase tracking-wide text-slate-500">Customer ID</th>
                            <th className="px-3 py-3 text-right text-[9px] font-bold uppercase tracking-wide text-slate-500">Refundable Amount</th>
                            <th className="px-3 py-3 text-center text-[9px] font-bold uppercase tracking-wide text-slate-500">Status</th>
                            <th className="px-4 py-3 text-right text-[9px] font-bold uppercase tracking-wide text-slate-500">Action</th>
                          </tr>
                        </thead>

                        <tbody>
                          {saleHistoryResults
                            .slice(
                              (saleHistoryPage - 1) * 5,
                              saleHistoryPage * 5
                            )
                            .map((sale) => {
                              const refundedAmount = getRefundedAmountForSale(sale, transactions)
                              const remainingRefundable = Math.max(
                                Number(sale.amount || 0) - refundedAmount,
                                0
                              )
                              const isRefundable = remainingRefundable > 0

                              return (
                                <tr
                                  key={sale.id}
                                  className="border-b border-slate-100 bg-white transition last:border-b-0 hover:bg-blue-50/30"
                                >
                                  <td className="px-4 py-3">
                                    <div className="flex items-center gap-2">
                                      <span className="flex h-7 w-7 shrink-0 items-center justify-center rounded-lg bg-blue-50 text-blue-600 ring-1 ring-blue-100">
                                        <RefundMethodIcon type="saleHistory" size={14} />
                                      </span>
                                      <span className="text-[11px] font-bold text-slate-800">
                                        {getCategoryName(sale.category)}
                                      </span>
                                    </div>
                                  </td>
                                  <td className="px-3 py-3 text-[10px] font-semibold text-slate-700">{formatDate(sale.date)}</td>
                                  <td className="px-3 py-3 text-[10px] font-semibold text-slate-700">{formatTime(sale.time)}</td>
                                  <td className="px-3 py-3 text-right text-[11px] font-extrabold text-slate-900">{formatMoney(sale.amount)}</td>
                                  <td className="px-3 py-3 text-[10px] font-medium text-slate-600">{sale.customerRef || 'Not available'}</td>
                                  <td className="px-3 py-3 text-right text-[11px] font-bold text-blue-700">{formatMoney(remainingRefundable)}</td>
                                  <td className="px-3 py-3 text-center">
                                    {isRefundable ? (
                                      <span className="inline-flex rounded-full bg-emerald-50 px-2.5 py-1 text-[9px] font-bold text-emerald-700 ring-1 ring-emerald-100">
                                        Refundable
                                      </span>
                                    ) : (
                                      <span className="inline-flex rounded-full bg-red-50 px-2.5 py-1 text-[9px] font-bold text-red-700 ring-1 ring-red-100">
                                        Fully Refunded
                                      </span>
                                    )}
                                  </td>
                                  <td className="px-4 py-3 text-right">
                                    <button
                                      type="button"
                                      disabled={!isRefundable}
                                      onClick={() => {
                                        setRefundSale(sale)
                                        setRefundCustomerNumber(sale.customerRef || '')
                                        setRefundForm((previous) => ({
                                          ...previous,
                                          category: sale.category,
                                          customerRef: sale.customerRef || '',
                                          amount: '',
                                          reason: '',
                                        }))
                                        setShowRefundDetails(true)
                                      }}
                                      className={`inline-flex min-w-[108px] items-center justify-center rounded-lg px-3 py-2 text-[10px] font-bold transition ${isRefundable
                                        ? ui.primaryButton
                                        : 'cursor-not-allowed bg-slate-100 text-slate-400'
                                        }`}
                                    >
                                      {isRefundable ? 'Refund' : 'Unavailable'}
                                    </button>
                                  </td>
                                </tr>
                              )
                            })}
                        </tbody>
                      </table>
                    </div>

                    {Math.ceil(saleHistoryResults.length / 5) > 1 && (
                      <div className="flex items-center justify-between gap-3 border-t border-slate-100 px-4 py-3 sm:px-5">
                        <button
                          type="button"
                          disabled={saleHistoryPage === 1}
                          onClick={() =>
                            setSaleHistoryPage((page) => Math.max(page - 1, 1))
                          }
                          className="rounded-lg border border-slate-200 bg-white px-3 py-2 text-xs font-semibold text-slate-600 transition hover:border-blue-200 hover:bg-blue-50 disabled:cursor-not-allowed disabled:opacity-40"
                        >
                          <span className="sm:hidden">←</span>
                          <span className="hidden sm:inline">← Previous</span>
                        </button>

                        <p className="text-[10px] font-semibold text-slate-500 sm:text-xs">
                          Page {saleHistoryPage} of {Math.ceil(saleHistoryResults.length / 5)}
                        </p>

                        <button
                          type="button"
                          disabled={saleHistoryPage >= Math.ceil(saleHistoryResults.length / 5)}
                          onClick={() =>
                            setSaleHistoryPage((page) =>
                              Math.min(page + 1, Math.ceil(saleHistoryResults.length / 5))
                            )
                          }
                          className="rounded-lg border border-slate-200 bg-white px-3 py-2 text-xs font-semibold text-slate-600 transition hover:border-blue-200 hover:bg-blue-50 disabled:cursor-not-allowed disabled:opacity-40"
                        >
                          <span className="sm:hidden">→</span>
                          <span className="hidden sm:inline">Next →</span>
                        </button>
                      </div>
                    )}
                  </>
                ) : (
                  <div className="px-5 py-10 text-center">
                    <div className="mx-auto flex h-11 w-11 items-center justify-center rounded-xl bg-slate-50 text-slate-400 ring-1 ring-slate-200">
                      🔎
                    </div>
                    <p className="mt-3 text-sm font-semibold text-slate-700">
                      No matching sales found
                    </p>
                    <p className="mt-1 text-xs text-slate-400">
                      Try changing the date or other search filters.
                    </p>
                  </div>
                )}
              </div>
            )}
          </div>

          {/* Refund Details Popup */}
          {refundSale && showRefundDetails && (
            <div className="fixed inset-0 z-[80] flex items-center justify-center bg-slate-950/45 p-3 backdrop-blur-[2px] sm:p-5">
              <div
                className="absolute inset-0"
                aria-hidden="true"
              />

              <div className="relative z-10 flex max-h-[92vh] w-full max-w-3xl flex-col overflow-hidden rounded-2xl border border-blue-200 bg-white shadow-[0_24px_80px_rgba(15,23,42,0.22)]">
                {/* Popup Header */}
                <div className="flex items-center justify-between gap-3 border-b border-blue-100 bg-blue-50/60 px-4 py-3 sm:px-5">
                  <div className="flex min-w-0 items-center gap-3">
                    <div className="flex h-9 w-9 shrink-0 items-center justify-center rounded-xl bg-blue-600 text-white shadow-sm ring-2 ring-blue-100">
                      ↩
                    </div>

                    <div className="min-w-0">
                      <p className="text-[10px] font-bold uppercase tracking-[0.14em] text-blue-600">
                        Refund Details
                      </p>
                      <h3 className="mt-0.5 truncate text-sm font-bold text-slate-900 sm:text-base">
                        {getCategoryName(refundSale.category)}
                      </h3>
                    </div>
                  </div>

                  <button
                    type="button"
                    onClick={() => {
                      setShowRefundDetails(false)
                      setRefundForm((previous) => ({
                        ...previous,
                        amount: '',
                        reason: '',
                      }))
                    }}
                    className="flex h-8 w-8 shrink-0 items-center justify-center rounded-lg border border-slate-200 bg-white text-slate-500 transition hover:border-blue-200 hover:bg-blue-50 hover:text-blue-700"
                    aria-label="Back to sale history"
                  >
                    ×
                  </button>
                </div>

                <div className="overflow-y-auto">
                  {/* Original Sale */}
                  <div className="px-4 pt-4 sm:px-5">
                    <div className="rounded-xl border border-blue-100 bg-blue-50/70 p-3.5 sm:p-4">
                      <div className="flex flex-col gap-3 sm:flex-row sm:items-center sm:justify-between">
                        <div>
                          <p className="text-[9px] font-bold uppercase tracking-[0.12em] text-blue-500">
                            Original Sale
                          </p>
                          <p className="mt-1 text-base font-extrabold text-slate-900">
                            {formatMoney(refundSale.amount)}
                          </p>
                        </div>

                        {(() => {
                          const remainingRefundable = Math.max(
                            Number(refundSale.amount || 0) -
                            getRefundedAmountForSale(refundSale, transactions),
                            0
                          )

                          return (
                            <span className={`w-fit rounded-full px-2.5 py-1 text-[9px] font-bold ring-1 ${remainingRefundable > 0
                              ? 'bg-emerald-50 text-emerald-700 ring-emerald-100'
                              : 'bg-red-50 text-red-700 ring-red-100'
                              }`}>
                              {remainingRefundable > 0 ? 'Refundable' : 'Fully Refunded'}
                            </span>
                          )
                        })()}
                      </div>
                    </div>
                  </div>

                  {/* Sale Information */}
                  <div className="px-4 py-4 sm:px-5">
                    <p className="mb-2 text-[10px] font-bold uppercase tracking-[0.12em] text-slate-400">
                      Sale Information
                    </p>

                    <div className="overflow-x-auto rounded-xl border border-slate-200">
                      <table className="w-full min-w-[620px] border-collapse">
                        <thead>
                          <tr className="border-b border-slate-100 bg-slate-50">
                            <th className="px-3 py-2.5 text-left text-[9px] font-bold uppercase tracking-wide text-slate-500">Category</th>
                            <th className="px-3 py-2.5 text-left text-[9px] font-bold uppercase tracking-wide text-slate-500">Date</th>
                            <th className="px-3 py-2.5 text-left text-[9px] font-bold uppercase tracking-wide text-slate-500">Time</th>
                            <th className="px-3 py-2.5 text-right text-[9px] font-bold uppercase tracking-wide text-slate-500">Amount</th>
                            <th className="px-3 py-2.5 text-left text-[9px] font-bold uppercase tracking-wide text-slate-500">Customer ID</th>
                          </tr>
                        </thead>
                        <tbody>
                          <tr>
                            <td className="px-3 py-3 text-[10px] font-bold text-slate-800">{getCategoryName(refundSale.category)}</td>
                            <td className="px-3 py-3 text-[10px] font-semibold text-slate-700">{formatDate(refundSale.date)}</td>
                            <td className="px-3 py-3 text-[10px] font-semibold text-slate-700">{formatTime(refundSale.time)}</td>
                            <td className="px-3 py-3 text-right text-[10px] font-bold text-slate-800">{formatMoney(refundSale.amount)}</td>
                            <td className="px-3 py-3 text-[10px] font-medium text-slate-600">{refundSale.customerRef || 'Not available'}</td>
                          </tr>
                        </tbody>
                      </table>
                    </div>
                  </div>

                  {/* Refund Summary */}
                  {(() => {
                    const previousRefunds = getRefundedAmountForSale(refundSale, transactions)
                    const remainingRefundable = Math.max(
                      Number(refundSale.amount || 0) - previousRefunds,
                      0
                    )

                    return (
                      <div className="px-4 pb-4 sm:px-5">
                        <p className="mb-2 text-[10px] font-bold uppercase tracking-[0.12em] text-slate-400">
                          Refund Summary
                        </p>

                        <div className="grid grid-cols-1 overflow-hidden rounded-xl border border-slate-200 bg-slate-50 sm:grid-cols-3">
                          <div className="border-b border-slate-200 px-3 py-3 sm:border-b-0 sm:border-r">
                            <p className="text-[9px] font-semibold uppercase tracking-wide text-slate-400">Already Refunded</p>
                            <p className="mt-1 text-sm font-bold text-slate-700">{formatMoney(previousRefunds)}</p>
                          </div>
                          <div className="border-b border-slate-200 px-3 py-3 sm:border-b-0 sm:border-r">
                            <p className="text-[9px] font-semibold uppercase tracking-wide text-slate-400">Remaining</p>
                            <p className="mt-1 text-sm font-bold text-blue-700">{formatMoney(remainingRefundable)}</p>
                          </div>
                          <div className="px-3 py-3">
                            <p className="text-[9px] font-semibold uppercase tracking-wide text-slate-400">Status</p>
                            <p className={`mt-1 inline-flex rounded-full px-2 py-1 text-[9px] font-bold ring-1 ${remainingRefundable > 0
                              ? 'bg-emerald-50 text-emerald-700 ring-emerald-100'
                              : 'bg-red-50 text-red-700 ring-red-100'
                              }`}>
                              {remainingRefundable > 0 ? 'Refundable' : 'Fully Refunded'}
                            </p>
                          </div>
                        </div>
                      </div>
                    )
                  })()}

                  {/* Refund Form */}
                  <div className="border-t border-blue-100 bg-white px-4 py-4 sm:px-5">
                    <div className="grid grid-cols-1 gap-4 sm:grid-cols-2">
                      <div>
                        <label className={ui.label}>Refund Amount</label>
                        <div className="relative">
                          <span className="absolute left-4 top-1/2 -translate-y-1/2 text-xs font-bold text-slate-500">Rs.</span>
                          <input
                            type="number"
                            min="0"
                            max={Math.max(
                              Number(refundSale.amount || 0) -
                              getRefundedAmountForSale(refundSale, transactions),
                              0
                            )}
                            value={refundForm.amount}
                            onChange={(e) =>
                              setRefundForm((previous) => ({
                                ...previous,
                                amount: e.target.value,
                              }))
                            }
                            placeholder="Enter refund amount"
                            className={`${ui.input} pl-12`}
                          />
                        </div>
                        <p className="mt-1.5 text-[11px] text-slate-400">
                          Maximum: <span className="font-semibold text-slate-600">{formatMoney(Math.max(Number(refundSale.amount || 0) - getRefundedAmountForSale(refundSale, transactions), 0))}</span>
                        </p>
                      </div>

                      <div>
                        <label className={ui.label}>Refund Reason</label>
                        <select
                          value={refundForm.reason}
                          onChange={(e) =>
                            setRefundForm((previous) => ({
                              ...previous,
                              reason: e.target.value,
                            }))
                          }
                          className={ui.input}
                        >
                          <option value="">Select refund reason</option>
                          <option value="Medicine Already Available At Home">Medicine Already Available At Home</option>
                          <option value="Wrong Medicine Purchased">Wrong Medicine Purchased</option>
                          <option value="Duplicate Purchase">Duplicate Purchase</option>
                          <option value="Customer Changed Mind">Customer Changed Mind</option>
                          <option value="Medicine Not Required">Medicine Not Required</option>
                          <option value="Doctor Changed Prescription">Doctor Changed Prescription</option>
                          <option value="Purchased By Mistake">Purchased By Mistake</option>
                          <option value="Other Reason">Other Reason</option>
                        </select>
                      </div>
                    </div>

                    <div className="mt-4 flex flex-col-reverse gap-2.5 border-t border-slate-100 pt-4 sm:flex-row sm:items-center sm:justify-between">
                      <button
                        type="button"
                        onClick={() => {
                          setShowRefundDetails(false)
                          setRefundForm((previous) => ({
                            ...previous,
                            amount: '',
                            reason: '',
                          }))
                        }}
                        className={`${ui.secondaryButton} w-full px-5 py-2.5 transition-colors duration-200 hover:!border-blue-600 hover:!bg-blue-600 hover:!text-white sm:w-36`}
                      >
                        ← Back
                      </button>

                      <button
                        type="button"
                        onClick={handleAddRefund}
                        disabled={!refundForm.amount || !refundForm.reason}
                        className={`${ui.primaryButton} w-full px-5 py-2.5 sm:w-auto`}
                      >
                        Process Refund
                      </button>
                    </div>
                  </div>
                </div>
              </div>
            </div>
          )}
        </PageContainer>
      ) : currentView === 'recordDetails' && selectedRecord ? (
        <DailyRecordDetails
          record={selectedRecord}
          onBack={() => {
            setSelectedRecord(null)
            setCurrentView('records')
          }}
        />
      ) : currentView === 'records' ? (
        <PageContainer
          title="Daily Records"
          subtitle="View and search completed daily sessions."
        >
          <DailyHistory
            records={dailyRecords}
            onViewDetails={(record) => {
              setSelectedRecord(record)
              setCurrentView('recordDetails')
            }}
            onRefresh={refreshHistory}
          />
        </PageContainer>
      ) : currentView === 'auditLogs' && isAdmin ? (
        <PageContainer
          title="Audit Logs"
        >
          <AuditLogsTable logs={auditLogs} />
        </PageContainer>
      ) : currentView === 'monthly' && isAdmin ? (
        <PageContainer
          title="Monthly Dashboard & Reports"
          subtitle="Review sales, cash reconciliation, and daily performance by month."
        >
          <MonthlyReports
            records={dailyRecords}
          />
        </PageContainer>
      ) : (
        <>
          <div className="mt-5 grid grid-cols-1 xl:grid-cols-[minmax(0,1fr)_360px] gap-5 items-start">
            <div className="min-w-0 space-y-5">

              <div className="mb-3 flex justify-start">
                <button
                  type="button"
                  onClick={() => {
                    setRefundForm({
                      category: 'all',
                      amount: '',
                      reason: '',
                      customerRef: '',
                    })
                    setRefundCustomerNumber('')
                    setRefundSale(null)
                    setShowRefundDetails(false)
                    setRefundSearchMethod('saleHistory')
                    setSaleHistoryFilters({
                      date: '',
                      category: 'all',
                      amount: '',
                      time: '',
                    })
                    setSaleHistoryResults([])
                    setSaleHistorySearched(false)
                    setSaleHistoryPage(1)
                    setShowRefundModal(false)
                    setCurrentView('refund')
                  }}
                  className={`${ui.primaryButton} px-4 py-2 text-sm`}
                >
                  + Add Refund
                </button>
              </div>
              <div className="grid grid-cols-2 gap-3 lg:grid-cols-4">
                <div className="animate-[summaryCardEnter_280ms_ease-out_80ms_both]">
                  <SummaryCard
                    title="Opening Cash"
                    value={formatMoney(
                      openingCash
                    )}
                    icon="cash"
                  />
                </div>

                <div className="animate-[summaryCardEnter_280ms_ease-out_140ms_both]">
                  <SummaryCard
                    title="Gross Sales"
                    value={formatMoney(
                      grossSales
                    )}
                    icon="sales"
                  />
                </div>

                <div className="animate-[summaryCardEnter_280ms_ease-out_200ms_both]">
                  <SummaryCard
                    title="Refunds"
                    value={formatMoney(
                      totalRefunds
                    )}
                    icon="refund"
                  />
                </div>

                <div className="animate-[summaryCardEnter_280ms_ease-out_260ms_both]">
                  <SummaryCard
                    title="Expected Cash"
                    value={formatMoney(
                      expectedClosingCash
                    )}
                    icon="total"
                  />
                </div>
              </div>

              <section>
                <div className="mb-3">
                  <h2 className="text-lg font-bold tracking-tight text-slate-900">
                    Sales by Category
                  </h2>
                </div>


                <div className="grid grid-cols-1 gap-4 md:grid-cols-3">
                  {[
                    {
                      category: 'medicine',
                      title: 'Medicine',
                      amount: sales.medicine,
                    },
                    {
                      category: 'general',
                      title: 'General Items',
                      amount: sales.general,
                    },
                    {
                      category: 'dispensing',
                      title: 'Dispensing',
                      amount: sales.dispensing,
                    },
                  ].map((item) => (
                    <SaleCategory
                      key={item.category}
                      category={item.category}
                      title={item.title}
                      amount={item.amount}
                    >
                      <div className="mt-4 border-t border-slate-100 pt-3">
                        <label
                          htmlFor={`sale-${item.category}`}
                          className="mb-1.5 block text-xs font-semibold text-slate-600"
                        >
                          Add Sale Amount (Rs.)
                        </label>

                        <div className="flex min-w-0 items-stretch gap-2">
                          <input
                            id={`sale-${item.category}`}
                            type="number"
                            min="0"
                            step="any"
                            value={categorySaleAmounts[item.category]}
                            onChange={(e) =>
                              setCategorySaleAmounts((previous) => ({
                                ...previous,
                                [item.category]: e.target.value,
                              }))
                            }
                            onKeyDown={(e) => {
                              if (e.key === 'Enter') {
                                handleAddSale(
                                  item.category,
                                  categorySaleAmounts[item.category]
                                )
                              }
                            }}
                            placeholder="Enter amount"
                            disabled={!isDayOpen || isDayClosed}
                            className="min-w-0 flex-1 rounded-lg border border-slate-200 bg-white px-2.5 py-2 text-sm text-slate-900 outline-none transition focus:border-blue-500 focus:ring-2 focus:ring-blue-100 disabled:cursor-not-allowed disabled:bg-slate-50"
                          />

                          <button
                            type="button"
                            onClick={() =>
                              handleAddSale(
                                item.category,
                                categorySaleAmounts[item.category]
                              )
                            }
                            disabled={
                              !isDayOpen ||
                              isDayClosed ||
                              !categorySaleAmounts[item.category]
                            }
                            className="shrink-0 rounded-lg bg-gradient-to-b from-blue-500 to-blue-700 px-3 py-2 text-xs font-semibold text-white shadow-sm transition hover:from-blue-600 hover:to-blue-800 focus:outline-none focus:ring-2 focus:ring-blue-200 disabled:cursor-not-allowed disabled:opacity-50"
                          >
                            + Add
                          </button>
                        </div>
                      </div>
                    </SaleCategory>
                  ))}
                </div>
              </section>

              <section>
                <div className={`${ui.card} p-5`}>
                  <div className="flex flex-col md:flex-row md:items-center md:justify-between gap-4">
                    <div>
                      <h2 className="text-base font-bold tracking-tight text-slate-900 sm:text-lg">
                        Daily Closing
                      </h2>
                    </div>

                    <button
                      onClick={
                        handleOpenClosing
                      }
                      className={`${ui.primaryButton} px-5`}
                    >
                      Close Day
                    </button>
                  </div>

                  <div className="mt-5 grid grid-cols-1 sm:grid-cols-3 gap-4">
                    <div className="rounded-xl border border-slate-200/80 bg-slate-50 p-4">
                      <p className="text-[11px] font-semibold uppercase tracking-[0.08em] text-slate-500 sm:text-xs">
                        Opening Cash
                      </p>
                      <p className="mt-2 text-base font-bold tracking-tight text-slate-900 sm:text-lg">
                        {formatMoney(openingCash)}
                      </p>
                    </div>

                    <div className="rounded-xl border border-blue-100 bg-blue-50 p-4">
                      <p className="text-[11px] font-semibold uppercase tracking-[0.08em] text-blue-600 sm:text-xs">
                        Net Sales
                      </p>
                      <p className="mt-2 text-base font-bold tracking-tight text-blue-700 sm:text-lg">
                        {formatMoney(netSales)}
                      </p>
                    </div>

                    <div className="rounded-xl border border-emerald-100 bg-emerald-50 p-4">
                      <p className="text-[11px] font-semibold uppercase tracking-[0.08em] text-emerald-600 sm:text-xs">
                        Expected Closing Cash
                      </p>
                      <p className="mt-2 text-base font-bold tracking-tight text-emerald-700 sm:text-lg">
                        {formatMoney(expectedClosingCash)}
                      </p>
                    </div>
                  </div>
                </div>
              </section>
            </div>

            <aside className="xl:sticky xl:top-5">
              <div>
                <section className={`${ui.section} overflow-hidden`}>
                  <div className="border-b border-slate-200/80 bg-slate-50/60 p-5">
                    <div className="flex items-center justify-between gap-3">
                      <div>
                        <h2 className="text-base font-bold tracking-tight text-slate-900 sm:text-lg">
                          Today's Activity
                        </h2>

                        <div className="mt-1 space-y-0.5">
                          <p className="text-sm text-slate-500">
                            {transactions.length}{' '}
                            transaction
                            {transactions.length !==
                              1
                              ? 's'
                              : ''}{' '}
                            recorded
                          </p>

                          <p className="text-xs text-blue-600 font-medium">
                            {new Date().toLocaleString(
                              'en-US',
                              { month: 'long' }
                            )}{' '}
                            • {monthlyCustomerCount}{' '}
                            customer reference
                            {monthlyCustomerCount !==
                              1
                              ? 's'
                              : ''}
                          </p>
                        </div>
                      </div>

                      <div className="flex h-10 w-10 items-center justify-center rounded-xl bg-blue-50 text-blue-600">
                        <svg
                          width="19"
                          height="19"
                          viewBox="0 0 24 24"
                          fill="none"
                          stroke="currentColor"
                          strokeWidth="2"
                        >
                          <path d="M6 3h12v18H6z" />
                          <path
                            d="M9 7h6M9 11h6M9 15h4"
                            strokeLinecap="round"
                          />
                        </svg>
                      </div>
                    </div>
                  </div>

                  {transactions.length ===
                    0 ? (
                    <div className="px-5 py-12 text-center">
                      <div className="mx-auto flex h-12 w-12 items-center justify-center rounded-2xl bg-slate-50 text-slate-400 ring-1 ring-slate-200">
                        <svg
                          width="22"
                          height="22"
                          viewBox="0 0 24 24"
                          fill="none"
                          stroke="currentColor"
                          strokeWidth="1.8"
                          strokeLinecap="round"
                          strokeLinejoin="round"
                        >
                          <path d="M6 3h12v18H6z" />
                          <path d="M9 7h6M9 11h6M9 15h4" />
                        </svg>
                      </div>

                      <p className="mt-4 text-sm font-bold text-slate-800">
                        No transactions yet
                      </p>

                      <p className="mx-auto mt-1 max-w-sm text-sm leading-6 text-slate-500">
                        Added sales and refunds will appear here.
                      </p>
                    </div>
                  ) : (
                    <>
                      <div className="divide-y divide-slate-100">
                        {visibleTransactions.map(
                          (transaction) => (
                            <div
                              key={transaction.id}
                              className="flex items-start justify-between gap-3 px-5 py-4 transition hover:bg-slate-50"
                            >
                              <div className="flex min-w-0 items-start gap-3">
                                <div
                                  className={`inline-flex items-center rounded-full border px-2.5 py-1 text-[10px] font-bold uppercase tracking-wide ${transaction.type === 'sale'
                                    ? 'border-blue-200 bg-blue-50 text-blue-700'
                                    : 'border-red-200 bg-red-50 text-red-700'
                                    }`}
                                >
                                  <CategoryIcon
                                    category={transaction.category}
                                    size={18}
                                  />
                                </div>

                                <div className="min-w-0">
                                  <div className="flex flex-wrap items-center gap-2">
                                    <p className="truncate text-sm font-semibold text-slate-900">
                                      {getCategoryName(
                                        transaction.category
                                      )}
                                    </p>

                                    <span
                                      className={`rounded-full px-2 py-0.5 text-[9px] font-bold uppercase tracking-wider ${transaction.type === 'sale'
                                        ? 'bg-emerald-50 text-emerald-700'
                                        : 'bg-red-50 text-red-700'
                                        }`}
                                    >
                                      {transaction.type === 'sale'
                                        ? 'SALE'
                                        : 'REFUND'}
                                    </span>
                                  </div>

                                  <p className="mt-1 text-xs text-slate-500">
                                    {formatTime(transaction.time)}
                                  </p>

                                  {transaction.userName && (
                                    <p className="mt-0.5 text-[11px] text-slate-400">
                                      By {transaction.userName}
                                    </p>
                                  )}

                                  {transaction.type === 'sale' &&
                                    transaction.customerRef && (
                                      <p className="mt-1 text-[11px] font-medium text-blue-600">
                                        Customer ID: {transaction.customerRef}
                                      </p>
                                    )}
                                </div>
                              </div>

                              <div
                                className={`shrink-0 whitespace-nowrap pt-0.5 text-sm font-bold tracking-tight ${transaction.type === 'sale'
                                  ? 'text-emerald-600'
                                  : 'text-red-600'
                                  }`}
                              >
                                {transaction.type === 'sale'
                                  ? '+'
                                  : '-'}
                                {formatMoney(transaction.amount)}
                              </div>
                            </div>
                          )
                        )}
                      </div>

                      <div className="flex flex-col gap-3 border-t border-slate-200/80 bg-slate-50/60 px-5 py-3 sm:flex-row sm:items-center sm:justify-between">
                        <p className="text-xs font-semibold text-slate-500">
                          Page{' '}
                          <span className="text-slate-900">
                            {transactionPage}
                          </span>{' '}
                          of{' '}
                          <span className="text-slate-900">
                            {totalTransactionPages}
                          </span>
                        </p>

                        <div className="flex items-center justify-between gap-2 sm:justify-end">
                          {/* Previous arrow: Desktop */}
                          <button
                            type="button"
                            onClick={() =>
                              setTransactionPage(
                                (page) => Math.max(1, page - 1)
                              )
                            }
                            disabled={transactionPage <= 1}
                            aria-label="Go to previous page"
                            className="hidden h-9 w-9 items-center justify-center rounded-lg border border-slate-300 bg-white text-lg font-semibold text-slate-700 transition hover:border-blue-200 hover:bg-blue-50 hover:text-blue-700 focus:outline-none focus:ring-2 focus:ring-blue-100 disabled:cursor-not-allowed disabled:opacity-40 sm:inline-flex"
                          >
                            ‹
                          </button>

                          {/* Previous button: Small devices */}
                          <button
                            type="button"
                            onClick={() =>
                              setTransactionPage(
                                (page) => Math.max(1, page - 1)
                              )
                            }
                            disabled={transactionPage <= 1}
                            className="inline-flex flex-1 items-center justify-center rounded-lg border border-slate-300 bg-white px-4 py-2 text-xs font-semibold text-slate-700 transition hover:border-blue-200 hover:bg-blue-50 hover:text-blue-700 focus:outline-none focus:ring-2 focus:ring-blue-100 disabled:cursor-not-allowed disabled:opacity-40 sm:hidden"
                          >
                            ‹ Previous
                          </button>

                          {/* Next arrow: Desktop */}
                          <button
                            type="button"
                            onClick={() =>
                              setTransactionPage(
                                (page) =>
                                  Math.min(
                                    totalTransactionPages,
                                    page + 1
                                  )
                              )
                            }
                            disabled={
                              transactionPage >=
                              totalTransactionPages
                            }
                            aria-label="Go to next page"
                            className="hidden h-9 w-9 items-center justify-center rounded-lg border border-slate-300 bg-white text-lg font-semibold text-slate-700 transition hover:border-blue-200 hover:bg-blue-50 hover:text-blue-700 focus:outline-none focus:ring-2 focus:ring-blue-100 disabled:cursor-not-allowed disabled:opacity-40 sm:inline-flex"
                          >
                            ›
                          </button>

                          {/* Next button: Small devices */}
                          <button
                            type="button"
                            onClick={() =>
                              setTransactionPage(
                                (page) =>
                                  Math.min(
                                    totalTransactionPages,
                                    page + 1
                                  )
                              )
                            }
                            disabled={
                              transactionPage >=
                              totalTransactionPages
                            }
                            className="inline-flex flex-1 items-center justify-center rounded-lg border border-slate-300 bg-white px-4 py-2 text-xs font-semibold text-slate-700 transition hover:border-blue-200 hover:bg-blue-50 hover:text-blue-700 focus:outline-none focus:ring-2 focus:ring-blue-100 disabled:cursor-not-allowed disabled:opacity-40 sm:hidden"
                          >
                            Next ›
                          </button>
                        </div>
                      </div>
                    </>
                  )}
                </section>
              </div>
            </aside>
          </div>
        </>
      )
      }

      {/* Refund Modal */}

      {
        showRefundModal && (
          <Modal
            title="Add Refund"
            onClose={() => {
              setShowRefundModal(false)
              setRefundSearchMethod('customerId')
            }}
          >
            {/* Search Method Toggle */}

            <div className="mb-5">
              <p className="mb-2 text-sm font-semibold text-slate-700">
                Search Original Sale
              </p>

              <div className="grid grid-cols-2 gap-2 rounded-xl bg-slate-100 p-1">
                <button
                  type="button"
                  onClick={() => {
                    setRefundSearchMethod('customerId')
                    setRefundSale(null)
                    setRefundForm({
                      ...refundForm,
                      amount: '',
                      reason: '',
                      customerRef: '',
                    })
                  }}
                  className={`rounded-lg px-3 py-2.5 text-sm font-semibold transition ${refundSearchMethod === 'customerId'
                    ? 'bg-white text-blue-700 shadow-sm'
                    : 'text-slate-500 hover:text-slate-700'
                    }`}
                >
                  Customer ID
                </button>

                <button
                  type="button"
                  onClick={() => {
                    setRefundSearchMethod('saleHistory')
                    setRefundSale(null)
                    setRefundForm({
                      ...refundForm,
                      amount: '',
                      reason: '',
                      customerRef: '',
                    })
                  }}
                  className={`rounded-lg px-3 py-2.5 text-sm font-semibold transition ${refundSearchMethod === 'saleHistory'
                    ? 'bg-white text-blue-700 shadow-sm'
                    : 'text-slate-500 hover:text-slate-700'
                    }`}
                >
                  Sale History
                </button>
              </div>
            </div>

            {/* Customer ID Search Method */}

            {refundSearchMethod === 'customerId' && (
              <>
                <label className={ui.label}>
                  Category
                </label>

                <select
                  value={refundForm.category}
                  disabled={!!refundSale}
                  onChange={(e) => {
                    const selectedCategory = e.target.value

                    setRefundForm((previous) => ({
                      ...previous,
                      category: selectedCategory,
                      amount: '',
                      reason: '',
                    }))

                    if (refundCustomerNumber.trim()) {
                      setRefundSale(null)

                      handleCustomerIdSearch(
                        refundCustomerNumber,
                        selectedCategory
                      )
                    }
                  }}
                  className={`${ui.input} disabled:cursor-not-allowed disabled:bg-slate-100 disabled:text-slate-500`}
                >
                  <option value="all">
                    All Categories
                  </option>

                  <option value="medicine">
                    Medicine
                  </option>

                  <option value="general">
                    General Items
                  </option>

                  <option value="dispensing">
                    Dispensing
                  </option>
                </select>

                <div>
                  <label className={ui.label}>
                    Customer ID
                    {!refundSale && (
                      <span className="ml-1 text-xs font-normal text-red-500">
                        Required
                      </span>
                    )}
                  </label>

                  <div className="relative">
                    <input
                      type="text"
                      value={refundCustomerNumber}
                      onChange={(e) => {
                        const customerNumber = e.target.value

                        setRefundCustomerNumber(customerNumber)

                        if (refundSale) {
                          setRefundSale(null)
                        }
                        setRefundForm((previous) => ({
                          ...previous,
                          customerRef: customerNumber,
                          amount: '',
                          reason: '',
                        }))

                        handleCustomerIdSearch(
                          customerNumber,
                          'all'
                        )
                      }}
                      placeholder="Enter Customer ID"
                      className={ui.input}
                    />
                  </div>
                </div>

                {refundForm.customerRef && (
                  <div className="mt-4">
                    {refundSale ? (
                      <>
                        {(() => {
                          const previousRefunds = transactions
                            .filter(
                              (transaction) =>
                                transaction.type === 'refund' &&
                                transaction.customerRef ===
                                refundSale.customerRef
                            )
                            .reduce(
                              (total, transaction) =>
                                total + Number(transaction.amount || 0),
                              0
                            )

                          const remainingRefundable =
                            Number(refundSale.amount || 0) - previousRefunds

                          if (remainingRefundable <= 0) {
                            return (
                              <div className="rounded-xl border border-red-100 bg-red-50/80 p-4">
                                <div className="flex items-start gap-3">
                                  <div className="flex h-9 w-9 shrink-0 items-center justify-center rounded-xl bg-white text-red-500 ring-1 ring-red-100">
                                    <svg
                                      width="18"
                                      height="18"
                                      viewBox="0 0 24 24"
                                      fill="none"
                                      stroke="currentColor"
                                      strokeWidth="1.8"
                                      strokeLinecap="round"
                                      strokeLinejoin="round"
                                    >
                                      <circle cx="12" cy="12" r="8" />
                                      <path d="M12 8v5" />
                                      <path d="M12 16h.01" />
                                    </svg>
                                  </div>

                                  <div className="min-w-0">
                                    <p className="text-sm font-semibold text-red-600">
                                      Sale Already Fully Refunded
                                    </p>

                                    <p className="mt-1 text-xs leading-5 text-red-500">
                                      This sale has already been fully refunded and cannot be refunded again.
                                    </p>
                                  </div>
                                </div>
                              </div>
                            )
                          }

                          return (
                            <div className="rounded-xl border border-emerald-100 bg-emerald-50/80 p-4">
                              <div className="flex items-start gap-3">
                                <div className="flex h-9 w-9 shrink-0 items-center justify-center rounded-xl bg-white text-emerald-600 ring-1 ring-emerald-100">
                                  <svg
                                    width="18"
                                    height="18"
                                    viewBox="0 0 24 24"
                                    fill="none"
                                    stroke="currentColor"
                                    strokeWidth="1.8"
                                    strokeLinecap="round"
                                    strokeLinejoin="round"
                                  >
                                    <path d="m5 12 4 4L19 6" />
                                  </svg>
                                </div>

                                <div className="min-w-0 flex-1">
                                  <p className="text-xs font-bold uppercase tracking-wide text-emerald-700">
                                    Sale Found
                                  </p>

                                  <p className="mt-2 text-sm font-semibold text-slate-800">
                                    {getCategoryName(refundSale.category)}
                                    {' • '}
                                    {formatMoney(refundSale.amount)}
                                  </p>

                                  <p className="mt-1 text-xs text-slate-500">
                                    Customer ID: {refundSale.customerRef}
                                  </p>

                                  <p className="mt-2 text-xs font-bold text-blue-600">
                                    Remaining Refundable:{' '}
                                    {formatMoney(remainingRefundable)}
                                  </p>

                                  <div className="mt-3 flex justify-end">
                                    <button
                                      type="button"
                                      onClick={() => {
                                        setShowRefundDetails(true)
                                      }}
                                      className="rounded-lg bg-blue-600 px-4 py-2 text-xs font-semibold text-white transition hover:bg-blue-700"
                                    >
                                      Select
                                    </button>
                                  </div>
                                </div>
                              </div>
                            </div>
                          )
                        })()}
                      </>
                    ) : (
                      <div className="rounded-xl border border-red-100 bg-red-50/80 p-4">
                        <div className="flex items-start gap-3">
                          <div className="flex h-9 w-9 shrink-0 items-center justify-center rounded-xl bg-white text-red-500 ring-1 ring-red-100">
                            <svg
                              width="18"
                              height="18"
                              viewBox="0 0 24 24"
                              fill="none"
                              stroke="currentColor"
                              strokeWidth="1.8"
                              strokeLinecap="round"
                              strokeLinejoin="round"
                            >
                              <circle cx="12" cy="12" r="8" />
                              <path d="M12 8v5" />
                              <path d="M12 16h.01" />
                            </svg>
                          </div>

                          <div className="min-w-0">
                            <p className="text-xs font-semibold text-red-600">
                              Customer ID not found
                            </p>

                            <p className="mt-1 text-xs leading-5 text-red-500">
                              Please enter a valid Customer ID.
                            </p>
                          </div>
                        </div>
                      </div>
                    )}
                  </div>
                )}
              </>
            )}

            {/* Sale History Search Method */}

            {refundSearchMethod === 'saleHistory' && (
              <div className="space-y-4">
                {/* Sale Date */}

                <div>
                  <label className={ui.label}>
                    Sale Date
                    <span className="ml-1 text-xs font-normal text-red-500">
                      Required
                    </span>
                  </label>

                  <input
                    type="date"
                    value={saleHistoryFilters.date}
                    onChange={(e) =>
                      setSaleHistoryFilters({
                        ...saleHistoryFilters,
                        date: e.target.value,
                      })
                    }
                    className={ui.input}
                  />
                </div>

                {/* Category */}

                <div>
                  <label className={ui.label}>
                    Category
                  </label>

                  <select
                    value={saleHistoryFilters.category}
                    onChange={(e) =>
                      setSaleHistoryFilters({
                        ...saleHistoryFilters,
                        category: e.target.value,
                      })
                    }
                    className={ui.input}
                  >
                    <option value="all">
                      All Categories
                    </option>

                    <option value="medicine">
                      Medicine
                    </option>

                    <option value="general">
                      General Items
                    </option>

                    <option value="dispensing">
                      Dispensing
                    </option>
                  </select>
                </div>

                {/* Sale Amount */}

                <div>
                  <label className={ui.label}>
                    Sale Amount
                  </label>

                  <div className="relative">
                    <span className="absolute left-4 top-1/2 -translate-y-1/2 text-xs font-bold text-slate-500">
                      Rs.
                    </span>

                    <input
                      type="number"
                      min="0"
                      value={saleHistoryFilters.amount}
                      onChange={(e) =>
                        setSaleHistoryFilters({
                          ...saleHistoryFilters,
                          amount: e.target.value,
                        })
                      }
                      placeholder="e.g. 1200"
                      className="w-full rounded-xl border border-slate-300 bg-white py-3 pl-12 pr-4 text-sm text-slate-900 outline-none transition focus:border-blue-500 focus:ring-4 focus:ring-blue-50"
                    />
                  </div>
                </div>

                {/* Approximate Time */}

                <div>
                  <label className={ui.label}>
                    Approximate Time
                    <span className="ml-1 text-xs font-normal text-slate-400">
                      Optional
                    </span>
                  </label>

                  <input
                    type="time"
                    value={saleHistoryFilters.time}
                    onChange={(e) =>
                      setSaleHistoryFilters({
                        ...saleHistoryFilters,
                        time: e.target.value,
                      })
                    }
                    className={ui.input}
                  />
                </div>

                {/* Search Button */}

                <button
                  type="button"
                  onClick={handleSaleHistorySearch}
                  className={`${ui.primaryButton} w-full`}
                >
                  Find Original Sale
                </button>

                {saleHistorySearched && (
                  <div className="mt-5">
                    <div className="mb-3 flex items-center justify-between gap-3">
                      <p className="text-sm font-bold text-slate-800">
                        Matching Sales
                      </p>

                      <span className="rounded-full bg-blue-50 px-2.5 py-1 text-xs font-bold text-blue-700">
                        {saleHistoryResults.length}{' '}
                        {saleHistoryResults.length === 1
                          ? 'sale'
                          : 'sales'}
                      </span>
                    </div>

                    {saleHistoryResults.length === 0 ? (
                      <div className="rounded-xl border border-slate-200 bg-slate-50/80 p-4">
                        <div className="flex items-start gap-3">
                          <div className="flex h-9 w-9 shrink-0 items-center justify-center rounded-xl bg-white text-slate-400 ring-1 ring-slate-200">
                            <svg
                              width="18"
                              height="18"
                              viewBox="0 0 24 24"
                              fill="none"
                              stroke="currentColor"
                              strokeWidth="1.8"
                              strokeLinecap="round"
                              strokeLinejoin="round"
                            >
                              <circle cx="11" cy="11" r="6" />
                              <path d="m16 16 4 4" />
                            </svg>
                          </div>

                          <div className="min-w-0">
                            <p className="text-sm font-semibold text-slate-700">
                              No matching sales found.
                            </p>

                            <p className="mt-1 text-xs leading-5 text-slate-500">
                              Try changing the date, category, amount, or approximate time.
                            </p>
                          </div>
                        </div>
                      </div>
                    ) : (
                      <div className="space-y-2">
                        {saleHistoryResults
                          .slice(
                            (saleHistoryPage - 1) * 5,
                            saleHistoryPage * 5
                          )
                          .map((sale) => {
                            const previousRefunds = getRefundedAmountForSale(
                              sale,
                              transactions
                            )

                            const remainingRefundable =
                              Number(sale.amount || 0) - previousRefunds

                            const isSelected =
                              refundSale?.id === sale.id

                            const isFullyRefunded =
                              remainingRefundable <= 0

                            return (
                              <div
                                key={sale.id}
                                className={`rounded-xl border p-3 transition ${isSelected
                                  ? 'border-blue-200 bg-blue-50/80 ring-1 ring-blue-100'
                                  : 'border-slate-200 bg-white hover:border-blue-200 hover:bg-slate-50'
                                  }`}
                              >
                                <div className="flex flex-col gap-2.5 sm:flex-row sm:items-center sm:justify-between">
                                  <div className="min-w-0">
                                    <div className="flex flex-wrap items-center gap-2">
                                      <p className="text-xs font-bold text-slate-900">
                                        {getCategoryName(sale.category)}
                                      </p>

                                      {sale.customerRef && (
                                        <span className="rounded-full bg-slate-100 px-2 py-0.5 text-[9px] font-bold text-slate-600">
                                          {sale.customerRef}
                                        </span>
                                      )}
                                    </div>

                                    <div className="mt-1 flex flex-wrap items-center gap-x-2 gap-y-0.5 text-[10px] text-slate-500">
                                      <span>
                                        {formatDate(sale.date)}
                                      </span>

                                      <span className="text-slate-300">
                                        •
                                      </span>

                                      <span>
                                        {formatTime(sale.time)}
                                      </span>
                                    </div>

                                    <div className="mt-1.5 flex flex-wrap items-center gap-x-3 gap-y-1">
                                      <p className="text-sm font-bold text-slate-900">
                                        {formatMoney(sale.amount)}
                                      </p>

                                      <div
                                        className={`inline-flex items-center gap-1.5 rounded-full px-2 py-1 ${isFullyRefunded
                                          ? 'bg-red-50 text-red-600 ring-1 ring-red-100'
                                          : 'bg-blue-50 text-blue-600 ring-1 ring-blue-100'
                                          }`}
                                      >
                                        <span
                                          className={`h-1.5 w-1.5 rounded-full ${isFullyRefunded
                                            ? 'bg-red-500'
                                            : 'bg-blue-500'
                                            }`}
                                        />

                                        <p className="text-[10px] font-semibold">
                                          {isFullyRefunded
                                            ? 'Fully Refunded'
                                            : `Remaining: ${formatMoney(
                                              remainingRefundable
                                            )}`}
                                        </p>
                                      </div>
                                    </div>
                                  </div>

                                  <button
                                    type="button"
                                    disabled={isFullyRefunded}
                                    onClick={() => {
                                      setRefundSale(sale)
                                      setShowRefundDetails(true)

                                      setRefundForm({
                                        category: sale.category,
                                        amount: '',
                                        reason: '',
                                        customerRef:
                                          sale.customerRef || '',
                                      })

                                      setRefundCustomerNumber(
                                        sale.customerRef || ''
                                      )
                                    }}
                                    className={`shrink-0 rounded-lg px-3 py-2 text-xs font-semibold transition ${isFullyRefunded
                                      ? 'cursor-not-allowed bg-slate-100 text-slate-400'
                                      : isSelected
                                        ? 'bg-blue-600 text-white'
                                        : 'border border-slate-300 bg-white text-slate-700 hover:border-blue-200 hover:bg-blue-50 hover:text-blue-700'
                                      }`}
                                  >
                                    {isSelected
                                      ? 'Selected'
                                      : isFullyRefunded
                                        ? 'Unavailable'
                                        : 'Select Sale'}
                                  </button>
                                </div>
                              </div>
                            )
                          })}

                        {saleHistoryResults.length > 5 && (
                          <div className="mt-3 border-t border-slate-100 pt-3">
                            <div className="flex items-center justify-between gap-2">
                              {/* Desktop Previous */}

                              <button
                                type="button"
                                onClick={() =>
                                  setSaleHistoryPage(
                                    (page) => Math.max(page - 1, 1)
                                  )
                                }
                                disabled={saleHistoryPage === 1}
                                className="hidden h-9 w-9 items-center justify-center rounded-lg border border-slate-200 bg-white text-slate-600 transition hover:border-blue-200 hover:bg-blue-50 hover:text-blue-700 disabled:cursor-not-allowed disabled:opacity-40 sm:flex"
                                aria-label="Previous page"
                              >
                                ←
                              </button>

                              {/* Mobile Previous */}

                              <button
                                type="button"
                                onClick={() =>
                                  setSaleHistoryPage(
                                    (page) => Math.max(page - 1, 1)
                                  )
                                }
                                disabled={saleHistoryPage === 1}
                                className="rounded-lg border border-slate-200 bg-white px-3 py-2 text-xs font-semibold text-slate-600 transition hover:border-blue-200 hover:bg-blue-50 hover:text-blue-700 disabled:cursor-not-allowed disabled:opacity-40 sm:hidden"
                              >
                                Previous
                              </button>

                              {/* Page Indicator */}

                              <span className="text-[10px] font-semibold text-slate-500 sm:text-xs">
                                Page {saleHistoryPage} of{' '}
                                {Math.ceil(
                                  saleHistoryResults.length / 5
                                )}
                              </span>

                              {/* Mobile Next */}

                              <button
                                type="button"
                                onClick={() =>
                                  setSaleHistoryPage(
                                    (page) =>
                                      Math.min(
                                        page + 1,
                                        Math.ceil(
                                          saleHistoryResults.length / 5
                                        )
                                      )
                                  )
                                }
                                disabled={
                                  saleHistoryPage >=
                                  Math.ceil(
                                    saleHistoryResults.length / 5
                                  )
                                }
                                className="rounded-lg border border-slate-200 bg-white px-3 py-2 text-xs font-semibold text-slate-600 transition hover:border-blue-200 hover:bg-blue-50 hover:text-blue-700 disabled:cursor-not-allowed disabled:opacity-40 sm:hidden"
                              >
                                Next
                              </button>

                              {/* Desktop Next */}

                              <button
                                type="button"
                                onClick={() =>
                                  setSaleHistoryPage(
                                    (page) =>
                                      Math.min(
                                        page + 1,
                                        Math.ceil(
                                          saleHistoryResults.length / 5
                                        )
                                      )
                                  )
                                }
                                disabled={
                                  saleHistoryPage >=
                                  Math.ceil(
                                    saleHistoryResults.length / 5
                                  )
                                }
                                className="hidden h-9 w-9 items-center justify-center rounded-lg border border-slate-200 bg-white text-slate-600 transition hover:border-blue-200 hover:bg-blue-50 hover:text-blue-700 disabled:cursor-not-allowed disabled:opacity-40 sm:flex"
                                aria-label="Next page"
                              >
                                →
                              </button>
                            </div>
                          </div>
                        )}
                      </div>
                    )}
                  </div>
                )}
              </div>
            )}

            {/* Selected Sale Summary */}

            {refundSale && showRefundDetails && (
              <div className="mt-4 overflow-hidden rounded-2xl border border-slate-200 bg-white shadow-sm">

                {/* Header */}
                <div className="flex items-center justify-between gap-3 border-b border-slate-100 px-4 py-4 sm:px-5">
                  <div className="flex min-w-0 items-center gap-3">
                    <div className="flex h-10 w-10 shrink-0 items-center justify-center rounded-xl bg-blue-600 text-base text-white ring-2 ring-blue-100">
                      ↩
                    </div>

                    <div className="min-w-0">
                      <div className="flex items-center gap-2">
                        <p className="text-xs font-bold uppercase tracking-widest text-blue-600">
                          Original Sale
                        </p>

                        <span className="inline-flex items-center gap-1 rounded-full bg-emerald-50 px-2 py-0.5 text-[9px] font-bold uppercase tracking-wide text-emerald-700 ring-1 ring-emerald-100">
                          <span className="text-[8px]">
                            ✓
                          </span>
                          Verified
                        </span>
                      </div>

                      <p className="mt-1 truncate text-sm font-bold text-slate-900">
                        {getCategoryName(refundSale.category)}
                      </p>
                    </div>
                  </div>

                  <div className="flex shrink-0 items-center gap-1.5 rounded-full border border-blue-100 bg-blue-50 px-2.5 py-1.5">
                    <span className="text-[10px] font-bold text-blue-700">
                      ✓
                    </span>

                    <span className="hidden text-[9px] font-bold uppercase tracking-wide text-blue-700 sm:inline">
                      Selected
                    </span>
                  </div>
                </div>


                {/* Main Sale Amount */}
                <div className="px-4 pt-4 sm:px-5">

                  <div className="rounded-xl bg-blue-50 px-4 py-3.5">
                    <div className="flex items-center justify-between gap-4">

                      <div>
                        <p className="text-[9px] font-bold uppercase tracking-[0.1em] text-blue-500">
                          Original Sale Amount
                        </p>
                      </div>

                      <p className="shrink-0 text-xl font-extrabold tracking-tight text-blue-700 sm:text-2xl">
                        {formatMoney(refundSale.amount)}
                      </p>

                    </div>
                  </div>

                </div>


                {/* Sale Information */}
                <div className="px-4 py-4 sm:px-5">

                  <p className="mb-3 text-[10px] font-bold uppercase tracking-[0.1em] text-slate-400">
                    Sale Information
                  </p>

                  <div className="w-full overflow-x-auto rounded-xl border border-slate-200">
                    <table className="w-full table-fixed border-collapse">

                      <thead>
                        <tr className="border-b border-slate-100 bg-slate-50">
                          <th className="w-auto px-2 py-2.5 text-left text-[8px] font-bold uppercase tracking-wide text-slate-400 sm:text-[10px]">
                            Sale Date
                          </th>

                          <th className="w-auto px-2 py-2.5 text-left text-[8px] font-bold uppercase tracking-wide text-slate-400 sm:text-[10px]">
                            Sale Time
                          </th>

                          <th className="w-auto px-2 py-2.5 text-left text-[8px] font-bold uppercase tracking-wide text-slate-400 sm:text-[10px]">
                            Original Amount
                          </th>

                          <th className="w-auto px-2 py-2.5 text-left text-[8px] font-bold uppercase tracking-wide text-slate-400 sm:text-[10px]">
                            Customer ID
                          </th>
                        </tr>
                      </thead>

                      <tbody>
                        <tr className="bg-white">

                          <td className="break-words px-2 py-2.5 text-[10px] font-bold leading-4 text-slate-800 sm:text-xs">
                            {formatDate(refundSale.date)}
                          </td>

                          <td className="break-words px-2 py-2.5 text-[10px] font-bold leading-4 text-slate-800 sm:text-xs">
                            {formatTime(refundSale.time)}
                          </td>

                          <td className="break-words px-2 py-2.5 text-[10px] font-extrabold leading-4 text-slate-900 sm:text-xs">
                            {formatMoney(refundSale.amount)}
                          </td>

                          <td className="min-w-0 break-all px-2 py-2.5 text-left text-[10px] font-bold leading-4 text-blue-700 sm:text-xs">
                            {refundSale.customerRef || 'Not available'}
                          </td>

                        </tr>
                      </tbody>

                    </table>
                  </div>

                </div>


                {/* Refund Financial Summary */}
                {(() => {
                  const previousRefunds = getRefundedAmountForSale(
                    refundSale,
                    transactions
                  )

                  const remainingRefundable =
                    Number(refundSale.amount || 0) - previousRefunds

                  const isFullyRefunded = remainingRefundable <= 0

                  return (
                    <>

                      <div className="mx-4 border-t border-slate-100 sm:mx-5" />

                      <div className="px-4 py-4 sm:px-5">

                        <p className="mb-2 text-[9px] font-bold uppercase tracking-[0.08em] text-slate-400">
                          Refund Summary
                        </p>

                        <div className="grid grid-cols-3 divide-x divide-slate-200 rounded-xl border border-slate-100 bg-slate-50">

                          <div className="px-3 py-3">
                            <p className="whitespace-nowrap text-[9px] font-semibold uppercase tracking-wide text-slate-400">
                              Already Refunded
                            </p>

                            <p className="mt-1.5 text-sm font-bold text-slate-700">
                              {formatMoney(previousRefunds)}
                            </p>
                          </div>

                          <div className="px-3 py-3">
                            <p className="whitespace-nowrap text-[9px] font-semibold uppercase tracking-wide text-slate-400">
                              Remaining
                            </p>

                            <p
                              className={`mt-1.5 text-sm font-bold ${!isFullyRefunded
                                ? 'text-blue-700'
                                : 'text-red-600'
                                }`}
                            >
                              {formatMoney(
                                Math.max(remainingRefundable, 0)
                              )}
                            </p>
                          </div>

                          <div className="px-3 py-3">
                            <p className="whitespace-nowrap text-[9px] font-semibold uppercase tracking-wide text-slate-400">
                              Refundable
                            </p>

                            <p
                              className={`mt-1.5 inline-flex items-center gap-1 rounded-full px-2 py-1 text-xs font-bold ${!isFullyRefunded
                                ? 'bg-emerald-50 text-emerald-600 ring-1 ring-emerald-100'
                                : 'bg-red-50 text-red-600 ring-1 ring-red-100'
                                }`}
                            >
                              <span className="text-[9px]">
                                {!isFullyRefunded ? '✓' : '×'}
                              </span>

                              {!isFullyRefunded ? 'Yes' : 'No'}
                            </p>
                          </div>
                        </div>
                      </div>
                    </>
                  )
                })()}
              </div>
            )}

            {refundSale && showRefundDetails && (
              <div className="mt-4 rounded-2xl border border-slate-200 bg-white p-4 shadow-sm">

                <div className="mb-4 flex items-center gap-3">
                  <div className="flex h-9 w-9 shrink-0 items-center justify-center rounded-xl bg-red-50 text-red-600 ring-1 ring-red-100">
                    <span className="text-sm font-bold">
                      ↩
                    </span>
                  </div>

                  <div className="min-w-0">
                    <p className="text-[10px] font-bold uppercase tracking-[0.08em] text-red-600">
                      Refund Details
                    </p>

                    <p className="mt-0.5 text-sm font-bold text-slate-900">
                      Enter refund information
                    </p>

                    <p className="mt-1 text-[11px] leading-4 text-slate-500">
                      Enter the amount and select the reason for this refund.
                    </p>
                  </div>
                </div>

                {/* Refund Amount */}

                <div>
                  <label className="mb-2 block text-sm font-semibold text-slate-700">
                    Refund Amount
                  </label>

                  <div className="relative">
                    <span className="absolute left-4 top-1/2 -translate-y-1/2 text-slate-500">
                      Rs.
                    </span>

                    <input
                      type="number"
                      min="0"
                      max={
                        Math.max(
                          Number(refundSale.amount || 0) -
                          getRefundedAmountForSale(
                            refundSale,
                            transactions
                          ),
                          0
                        )
                      }
                      value={refundForm.amount}
                      onChange={(e) =>
                        setRefundForm({
                          ...refundForm,
                          amount: e.target.value,
                        })
                      }
                      placeholder="500"
                      className="w-full rounded-xl border border-slate-300 bg-white py-3 pl-12 pr-4 text-sm font-medium text-slate-900 outline-none transition focus:border-blue-500 focus:ring-4 focus:ring-blue-50"
                    />
                  </div>

                  <p className="mt-1.5 text-xs text-slate-400">
                    Maximum refundable amount:{' '}
                    <span className="font-semibold text-slate-600">
                      {formatMoney(
                        Math.max(
                          Number(refundSale.amount || 0) -
                          getRefundedAmountForSale(
                            refundSale,
                            transactions
                          ),
                          0
                        )
                      )}
                    </span>
                  </p>
                </div>

                {/* Refund Reason */}

                <div className="mt-5">
                  <label className="mb-2 block text-sm font-semibold text-slate-700">
                    Refund Reason
                  </label>

                  <select
                    value={refundForm.reason}
                    onChange={(e) =>
                      setRefundForm({
                        ...refundForm,
                        reason: e.target.value,
                      })
                    }
                    className={ui.input}
                  >
                    <option value="">
                      Select refund reason
                    </option>

                    <option value="Medicine Already Available At Home">
                      Medicine Already Available At Home
                    </option>

                    <option value="Wrong Medicine Purchased">
                      Wrong Medicine Purchased
                    </option>

                    <option value="Duplicate Purchase">
                      Duplicate Purchase
                    </option>

                    <option value="Customer Changed Mind">
                      Customer Changed Mind
                    </option>

                    <option value="Medicine Not Required">
                      Medicine Not Required
                    </option>

                    <option value="Doctor Changed Prescription">
                      Doctor Changed Prescription
                    </option>

                    <option value="Purchased By Mistake">
                      Purchased By Mistake
                    </option>

                    <option value="Other Reason">
                      Other Reason
                    </option>
                  </select>
                </div>

              </div>
            )}

            {/* Modal Actions */}

            <div className="mt-7 border-t border-slate-100 pt-5">
              <button
                type="button"
                onClick={() => setShowRefundModal(false)}
                className={`${ui.secondaryButton} w-full`}
              >
                Cancel
              </button>
            </div>
          </Modal>
        )
      }

      {
        showRefundModal && showRefundDetails && refundSale && (
          <Modal
            title="Confirm Refund"
            onClose={() => {
              setShowRefundDetails(false)
            }}
          >
            <div className="rounded-2xl border border-slate-200 bg-white p-4 shadow-sm">
              <div className="flex flex-col gap-3 sm:flex-row sm:items-start sm:justify-between">
                <div className="min-w-0">
                  <div className="flex items-center gap-3">
                    <div className="flex h-11 w-11 shrink-0 items-center justify-center rounded-2xl bg-gradient-to-br from-blue-50 to-blue-100 text-blue-600 shadow-inner ring-1 ring-blue-100">
                      <span className="text-base font-bold">
                        ↩
                      </span>
                    </div>

                    <div className="min-w-0">
                      <p className="text-[10px] font-bold uppercase tracking-[0.08em] text-blue-600">
                        Original Sale
                      </p>

                      <p className="mt-0.5 text-sm font-bold text-slate-900">
                        {getCategoryName(refundSale.category)}
                      </p>
                    </div>
                  </div>

                  <p className="mt-2 text-xs leading-5 text-slate-500">
                    This sale has been selected as the original sale for this refund.
                  </p>
                </div>

                <div className="flex w-fit shrink-0 items-center gap-1.5 rounded-full border border-emerald-200 bg-emerald-50 px-3 py-1">
                  <span className="flex h-4 w-4 items-center justify-center rounded-full bg-emerald-100 text-[9px] font-bold text-emerald-700">
                    ✓
                  </span>

                  <span className="text-[10px] font-bold uppercase tracking-wide text-emerald-700">
                    Sale Selected
                  </span>
                </div>
              </div>

              <div className="mt-4 rounded-xl border border-slate-100 overflow-hidden">
                <div className="grid grid-cols-4">
                  <div className="min-w-0 border-r border-slate-100 bg-slate-50 px-1.5 py-2 sm:px-2.5">
                    <p className="text-[8px] sm:text-[9px] font-semibold leading-3 text-slate-500">
                      Customer ID
                    </p>

                    <p className="mt-1 break-words text-[9px] sm:text-[10px] font-bold leading-3.5 text-slate-800">
                      {refundSale.customerRef || 'Not available'}
                    </p>
                  </div>

                  <div className="min-w-0 border-r border-slate-100 bg-slate-50 px-1.5 py-2 sm:px-2.5">
                    <p className="text-[8px] sm:text-[9px] font-semibold leading-3 text-slate-500">
                      Category
                    </p>

                    <p className="mt-1 break-words text-[9px] sm:text-[10px] font-bold leading-3.5 text-slate-800">
                      {getCategoryName(refundSale.category)}
                    </p>
                  </div>

                  <div className="min-w-0 border-r border-slate-100 bg-slate-50 px-1.5 py-2 sm:px-2.5">
                    <p className="text-[8px] sm:text-[9px] font-semibold leading-3 text-slate-500">
                      Sale Date
                    </p>

                    <p className="mt-1 break-words text-[9px] sm:text-[10px] font-bold leading-3.5 text-slate-800">
                      {formatDate(refundSale.date)}
                    </p>
                  </div>

                  <div className="min-w-0 bg-slate-50 px-1.5 py-2 sm:px-2.5">
                    <p className="text-[8px] sm:text-[9px] font-semibold leading-3 text-slate-500">
                      Sale Time
                    </p>

                    <p className="mt-1 break-words text-[9px] sm:text-[10px] font-bold leading-3.5 text-slate-800">
                      {formatTime(refundSale.time)}
                    </p>
                  </div>
                </div>
              </div>

              {(() => {
                const previousRefunds = getRefundedAmountForSale(
                  refundSale,
                  transactions
                )

                const remainingRefundable =
                  Number(refundSale.amount || 0) - previousRefunds

                return (
                  <div className="mt-3 rounded-xl border border-slate-100 overflow-hidden">
                    <div className="grid grid-cols-3">
                      <div className="min-w-0 border-r border-slate-100 bg-slate-50 px-1.5 py-2 sm:px-2.5 sm:py-2.5">
                        <p className="whitespace-nowrap text-[7px] font-semibold tracking-tight text-slate-500 sm:text-[9px]">
                          Original Amount
                        </p>
                        <p className="mt-1 whitespace-nowrap text-[8px] font-bold tracking-tight text-slate-800 sm:text-[10px]">
                          {formatMoney(refundSale.amount)}
                        </p>
                      </div>

                      <div className="min-w-0 border-r border-slate-100 bg-slate-50 px-1.5 py-2 sm:px-2.5 sm:py-2.5">
                        <p className="whitespace-nowrap text-[7px] font-semibold tracking-tight text-slate-500 sm:text-[9px]">
                          Already Refunded
                        </p>
                        <p className="mt-1 whitespace-nowrap text-[8px] font-bold tracking-tight text-slate-800 sm:text-[10px]">
                          {formatMoney(previousRefunds)}
                        </p>
                      </div>

                      <div className="min-w-0 bg-blue-50/60 px-1.5 py-2 sm:px-2.5 sm:py-2.5">
                        <p className="whitespace-nowrap text-[7px] font-semibold tracking-tight text-slate-500 sm:text-[9px]">
                          Remaining Refundable
                        </p>
                        <p className="mt-1 whitespace-nowrap text-[8px] font-bold tracking-tight text-blue-700 sm:text-[10px]">
                          {formatMoney(Math.max(remainingRefundable, 0))}
                        </p>
                      </div>
                    </div>
                  </div>
                )
              })()}
            </div>
            <div className="mt-4 rounded-2xl border border-slate-200 bg-white p-4 shadow-sm">
              <div className="mb-4 flex items-center gap-3">
                <div className="flex h-9 w-9 shrink-0 items-center justify-center rounded-xl bg-red-50 text-red-600">
                  <span className="text-sm font-bold">
                    ↩
                  </span>
                </div>

                <div className="min-w-0">
                  <p className="text-[10px] font-bold uppercase tracking-[0.08em] text-red-600">
                    Refund Details
                  </p>

                  <p className="mt-0.5 text-sm font-bold text-slate-900">
                    Enter refund information
                  </p>
                </div>
              </div>

              {/* Refund Amount */}

              <div>
                <label className="mb-2 block text-sm font-semibold text-slate-700">
                  Refund Amount
                </label>

                <div className="relative">
                  <span className="absolute left-4 top-1/2 -translate-y-1/2 text-slate-500">
                    Rs.
                  </span>

                  <input
                    type="number"
                    min="0"
                    max={
                      Math.max(
                        Number(refundSale.amount || 0) -
                        getRefundedAmountForSale(
                          refundSale,
                          transactions
                        ),
                        0
                      )
                    }
                    value={refundForm.amount}
                    onChange={(e) =>
                      setRefundForm({
                        ...refundForm,
                        amount: e.target.value,
                      })
                    }
                    placeholder="500"
                    className="w-full rounded-xl border border-slate-300 bg-white py-3 pl-12 pr-4 text-sm font-medium text-slate-900 outline-none transition focus:border-blue-500 focus:ring-4 focus:ring-blue-50"
                  />
                </div>

                <p className="mt-1.5 text-xs text-slate-400">
                  Maximum refundable amount:{' '}
                  <span className="font-semibold text-slate-600">
                    {formatMoney(
                      Math.max(
                        Number(refundSale.amount || 0) -
                        getRefundedAmountForSale(
                          refundSale,
                          transactions
                        ),
                        0
                      )
                    )}
                  </span>
                </p>
              </div>

              {/* Refund Reason */}

              <div className="mt-5">
                <label className="mb-2 block text-sm font-semibold text-slate-700">
                  Refund Reason
                </label>

                <select
                  value={refundForm.reason}
                  onChange={(e) =>
                    setRefundForm({
                      ...refundForm,
                      reason: e.target.value,
                    })
                  }
                  className={ui.input}
                >
                  <option value="">
                    Select refund reason
                  </option>

                  <option value="Medicine Already Available At Home">
                    Medicine Already Available At Home
                  </option>

                  <option value="Wrong Medicine Purchased">
                    Wrong Medicine Purchased
                  </option>

                  <option value="Duplicate Purchase">
                    Duplicate Purchase
                  </option>

                  <option value="Customer Changed Mind">
                    Customer Changed Mind
                  </option>

                  <option value="Medicine Not Required">
                    Medicine Not Required
                  </option>

                  <option value="Doctor Changed Prescription">
                    Doctor Changed Prescription
                  </option>

                  <option value="Purchased By Mistake">
                    Purchased By Mistake
                  </option>

                  <option value="Other Reason">
                    Other Reason
                  </option>
                </select>
              </div>
            </div>

            <div className="mt-5 flex gap-3 border-t border-slate-100 pt-5">
              <button
                type="button"
                onClick={() => {
                  setShowRefundDetails(false)
                }}
                className={`${ui.secondaryButton} flex-1`}
              >
                Cancel
              </button>

              <button
                type="button"
                onClick={handleAddRefund}
                className={`${ui.secondaryButton} w-fit px-5 py-3 text-sm shadow-sm hover:-translate-y-0.5 hover:border-blue-200 hover:text-blue-700 hover:shadow-md`}
              >
                Add Refund
              </button>
            </div>
          </Modal>
        )
      }
      {/* Closing Modal */}

      {
        showClosingModal && (
          <Modal
            title="Daily Closing"
            wide
            onClose={() =>
              setShowClosingModal(false)
            }
          >
            <div className="grid grid-cols-1 gap-3 sm:grid-cols-2 lg:grid-cols-5">
              <DetailCard
                title="Opening Cash"
                value={formatMoney(
                  openingCash
                )}
                icon="cash"
                tone="blue"
              />

              <DetailCard
                title="Gross Sales"
                value={formatMoney(
                  grossSales
                )}
                icon="sales"
                tone="emerald"
              />

              <DetailCard
                title="Refunds"
                value={formatMoney(
                  totalRefunds
                )}
                icon="refund"
                tone="red"
              />

              <DetailCard
                title="Net Sales"
                value={formatMoney(
                  netSales
                )}
                icon="net"
                tone="violet"
              />

              <DetailCard
                title="Expected Cash"
                value={formatMoney(
                  expectedClosingCash
                )}
                icon="calculator"
                tone="sky"
              />
            </div>

            <div className="mt-6">
              <label className="block text-sm font-semibold text-slate-700 mb-2">
                Actual Cash
              </label>

              <div className="relative">
                <span className="absolute left-4 top-1/2 -translate-y-1/2 text-slate-500">
                  Rs.
                </span>

                <input
                  type="number"
                  min="0"
                  value={actualCash}
                  onChange={(e) =>
                    setActualCash(
                      e.target.value
                    )
                  }
                  placeholder="1900"
                  className="w-full rounded-xl border border-slate-300 bg-white py-3 pl-12 pr-4 text-sm text-slate-900 outline-none transition focus:border-blue-500 focus:ring-4 focus:ring-blue-50"
                />
              </div>
            </div>

            {actualCash !== '' &&
              !Number.isNaN(
                Number(actualCash)
              ) && (
                <div
                  className={`mt-4 rounded-xl border p-4 animate-[closingResultEnter_250ms_ease-out] ${Number(actualCash) -
                    expectedClosingCash ===
                    0
                    ? 'border-emerald-100 bg-emerald-50'
                    : Number(actualCash) -
                      expectedClosingCash <
                      0
                      ? 'border-red-100 bg-red-50'
                      : 'border-amber-100 bg-amber-50'
                    }`}
                >
                  <div className="flex flex-col gap-3 sm:flex-row sm:items-center sm:justify-between">
                    <div>
                      <p className="text-xs font-semibold uppercase tracking-wide text-slate-500">
                        Difference
                      </p>

                      <p
                        className={`mt-1 text-xl font-bold tracking-tight ${Number(actualCash) -
                          expectedClosingCash ===
                          0
                          ? 'text-emerald-700'
                          : Number(actualCash) -
                            expectedClosingCash <
                            0
                            ? 'text-red-700'
                            : 'text-amber-700'
                          }`}
                      >
                        {formatMoney(
                          Math.abs(
                            Number(actualCash) -
                            expectedClosingCash
                          )
                        )}
                      </p>
                    </div>

                    <span
                      className={`rounded-full px-3 py-1.5 text-xs font-bold ${Number(actualCash) -
                        expectedClosingCash ===
                        0
                        ? 'bg-emerald-100 text-emerald-700'
                        : Number(actualCash) -
                          expectedClosingCash <
                          0
                          ? 'bg-red-100 text-red-700'
                          : 'bg-amber-100 text-amber-700'
                        }`}
                    >
                      {Number(actualCash) -
                        expectedClosingCash ===
                        0
                        ? 'Matched'
                        : Number(actualCash) -
                          expectedClosingCash <
                          0
                          ? 'Shortage'
                          : 'Extra Cash'}
                    </span>
                  </div>
                </div>
              )}

            {actualCash !== '' &&
              !Number.isNaN(
                Number(actualCash)
              ) &&
              Number(actualCash) !==
              expectedClosingCash && (
                <div className="mt-5">
                  <label className="block text-sm font-semibold text-slate-700 mb-2">
                    Reason for Difference
                  </label>

                  <textarea
                    value={closingReason}
                    onChange={(e) =>
                      setClosingReason(
                        e.target.value
                      )
                    }
                    placeholder="Enter reason for shortage or extra cash"
                    rows="3"
                    className={`${ui.input} resize-none`}
                  />
                </div>
              )}

            <div className="mt-7 flex gap-3 border-t border-slate-100 pt-5">
              <button
                onClick={() =>
                  setShowClosingModal(false)
                }
                className={`${ui.secondaryButton} min-w-0 flex-1 whitespace-nowrap px-3 py-2.5 text-xs sm:text-sm`}
              >
                Cancel
              </button>

              <button
                onClick={handleCompleteClosing}
                className={`${ui.primaryButton} min-w-0 flex-1 whitespace-nowrap px-3 py-2.5 text-xs sm:text-sm flex items-center justify-center`}
              >
                Complete Closing
              </button>
            </div>
          </Modal>
        )
      }
    </AppShell >

  )
}

/* --------------------------------
   Login Screen
-------------------------------- */

function LoginScreen({ onLogin }) {
  const [mode, setMode] = useState('login')

  const [hasUsers, setHasUsers] = useState(null)
  const [checkingUsers, setCheckingUsers] = useState(true)

  const [username, setUsername] = useState('')
  const [password, setPassword] = useState('')
  const [showPassword, setShowPassword] = useState(false)

  const [name, setName] = useState('')
  const [confirmPassword, setConfirmPassword] =
    useState('')
  const [showConfirmPassword, setShowConfirmPassword] =
    useState(false)

  const [loggingIn, setLoggingIn] = useState(false)
  const [creatingAccount, setCreatingAccount] =
    useState(false)

  const [error, setError] = useState('')
  const [success, setSuccess] = useState('')

  useEffect(() => {
    checkRegisteredUsers()
  }, [])

  async function checkRegisteredUsers() {
    try {
      const registered =
        await hasRegisteredUsers()

      setHasUsers(registered)

      if (!registered) {
        setMode('signup')
      } else {
        setMode('login')
      }
    } catch (error) {
      console.error(
        'Failed to check registered users:',
        error
      )

      setError(
        'Unable to check local account data.'
      )
    } finally {
      setCheckingUsers(false)
    }
  }

  function switchMode(nextMode) {
    setMode(nextMode)
    setError('')
    setSuccess('')

    setUsername('')
    setPassword('')
    setName('')
    setConfirmPassword('')
    setShowPassword(false)
    setShowConfirmPassword(false)
  }

  async function handleLogin(e) {
    e.preventDefault()

    setError('')
    setSuccess('')

    if (!username.trim()) {
      setError(
        'Please enter your username.'
      )
      return
    }

    if (!password) {
      setError(
        'Please enter your password.'
      )
      return
    }

    setLoggingIn(true)

    try {
      const result = await onLogin(
        username,
        password
      )

      if (!result.success) {
        setError(result.message)
      }
    } catch (error) {
      console.error(
        'Login error:',
        error
      )

      setError(
        'Something went wrong while signing in.'
      )
    } finally {
      setLoggingIn(false)
    }
  }

  async function handleCreateAccount(e) {
    e.preventDefault()

    setError('')
    setSuccess('')

    if (!name.trim()) {
      setError(
        'Please enter your full name.'
      )
      return
    }

    if (!username.trim()) {
      setError(
        'Please enter a username.'
      )
      return
    }

    if (username.trim().length < 3) {
      setError(
        'Username must be at least 3 characters.'
      )
      return
    }

    if (!password) {
      setError(
        'Please create a password.'
      )
      return
    }

    const passwordRequirements = {
      minLength: password.length >= 8,
      uppercase: /[A-Z]/.test(password),
      lowercase: /[a-z]/.test(password),
      number: /[0-9]/.test(password),
      special: /[^A-Za-z0-9]/.test(password),
    }

    if (!passwordRequirements.minLength) {
      setError(
        'Password must be at least 8 characters long.'
      )
      return
    }

    if (!passwordRequirements.uppercase) {
      setError(
        'Password must contain at least one uppercase letter.'
      )
      return
    }

    if (!passwordRequirements.lowercase) {
      setError(
        'Password must contain at least one lowercase letter.'
      )
      return
    }

    if (!passwordRequirements.number) {
      setError(
        'Password must contain at least one number.'
      )
      return
    }

    if (!passwordRequirements.special) {
      setError(
        'Password must contain at least one special character.'
      )
      return
    }

    if (!confirmPassword) {
      setError(
        'Please confirm your password.'
      )
      return
    }

    if (password !== confirmPassword) {
      setError(
        'Passwords do not match.'
      )
      return
    }

    setCreatingAccount(true)

    try {
      await createUser({
        name,
        username,
        password,
        role: 'admin',
      })

      setSuccess(
        'Admin account created successfully. You can now sign in.'
      )

      setHasUsers(true)

      setPassword('')
      setConfirmPassword('')

      setTimeout(() => {
        setMode('login')
        setSuccess('')
      }, 1200)
    } catch (error) {
      console.error(
        'Account creation error:',
        error
      )

      setError(
        error.message ||
        'Unable to create the account.'
      )
    } finally {
      setCreatingAccount(false)
    }
  }

  if (checkingUsers) {
    return (
      <div className="min-h-screen bg-slate-50 flex items-center justify-center px-3 py-4 sm:px-4 sm:py-5">
        <div className={`${ui.card} w-full max-w-2xl`}>
          <div className="p-8 text-center">
            <div className="mx-auto mb-5 flex h-16 w-16 items-center justify-center rounded-2xl bg-blue-600 text-white shadow-sm">
              <svg
                width="32"
                height="32"
                viewBox="0 0 24 24"
                fill="none"
                stroke="currentColor"
                strokeWidth="2"
                strokeLinecap="round"
                strokeLinejoin="round"
              >
                <path d="M3 10h18" />
                <path d="M5 10v10h14V10" />
                <path d="M4 10V6h16v4" />
                <path d="M8 6V3h8v3" />
              </svg>
            </div>

            <h1 className="text-xl font-bold text-slate-900">
              Khalid Medical POS
            </h1>

            <p className="mt-2 text-sm text-slate-500">
              Checking local account...
            </p>

            <div className="mt-5 flex justify-center">
              <div className="h-6 w-6 animate-spin rounded-full border-2 border-slate-200 border-t-blue-600"></div>
            </div>
          </div>
        </div>
      </div>
    )
  }

  const isSignup = mode === 'signup'

  return (
    <div className="min-h-screen bg-slate-950 flex items-center justify-center px-3 py-6 sm:px-6 sm:py-8 relative overflow-hidden">
      <div className="relative z-10 w-full max-w-2xl">
        <div className="overflow-hidden rounded-3xl border border-slate-800 bg-slate-900/95 shadow-[0_24px_80px_rgba(0,0,0,0.45)] backdrop-blur-sm">
          <div className="px-5 py-5 sm:px-8 sm:py-6">
            <div className="mb-7 text-center sm:mb-8">

              {/* Logo */}

              <div className="relative mx-auto mb-5 h-28 w-28 flex items-center justify-center sm:h-32 sm:w-32">
                <div className="absolute inset-[-7px] rounded-full border border-blue-400/30 border-t-blue-400 border-r-cyan-300 animate-spin [animation-duration:3.5s]" />

                <div className="absolute inset-[-3px] rounded-full border border-cyan-400/20 border-b-cyan-300 animate-spin [animation-duration:2.5s] [animation-direction:reverse]" />

                <div className="relative mx-auto mb-5 h-28 w-28 flex items-center justify-center sm:h-32 sm:w-32">
                  <div className="absolute inset-[-7px] rounded-full border border-blue-400/30 border-t-blue-400 border-r-cyan-300 animate-spin [animation-duration:3.5s]" />

                  <div className="absolute inset-[-3px] rounded-full border border-cyan-400/20 border-b-cyan-300 animate-spin [animation-duration:2.5s] [animation-direction:reverse]" />

                  <img
                    src="/km-store-logo.png"
                    alt="KM Store logo"
                    className="relative z-10 h-28 w-28 rounded-full object-contain sm:h-32 sm:w-32"
                  />
                </div>
              </div>

              {/* App Name */}
              <h1 className="text-[25px] font-bold tracking-tight text-white sm:text-[29px]">
                Khalid Medical POS
              </h1>

              {/* App Type */}
              <p className="mt-2 text-[11px] font-bold uppercase tracking-[0.22em] text-blue-400 sm:text-xs">
                Management System
              </p>

              {/* Page Description */}
              <p className="mt-3 text-[13px] leading-5 text-slate-400 sm:text-sm">
                {isSignup
                  ? 'Create your administrator account'
                  : 'Sign in to continue'}
              </p>
            </div>

            {isSignup ? (
              <form onSubmit={handleCreateAccount} className="space-y-3.5">
                <div className="grid grid-cols-1 gap-3 sm:grid-cols-2">
                  <div>
                    <label className="mb-1.5 block text-[12px] font-semibold text-slate-200">
                      Full Name
                    </label>
                    <input
                      type="text"
                      value={name}
                      onChange={(e) => setName(e.target.value)}
                      placeholder="Enter full name"
                      className="w-full rounded-xl border border-slate-700 bg-slate-800/80 px-3 py-3 text-sm text-white shadow-sm outline-none transition-all duration-200 placeholder:text-slate-500 hover:border-slate-600 focus:border-blue-500 focus:ring-4 focus:ring-blue-500/10"
                    />
                  </div>

                  <div>
                    <label className="mb-1.5 block text-[12px] font-semibold text-slate-200">
                      Username
                    </label>
                    <input
                      type="text"
                      value={username}
                      onChange={(e) => setUsername(e.target.value)}
                      placeholder="Choose username"
                      className="w-full rounded-xl border border-slate-700 bg-slate-800/80 px-3 py-3 text-sm text-white shadow-sm outline-none transition-all duration-200 placeholder:text-slate-500 hover:border-slate-600 focus:border-blue-500 focus:ring-4 focus:ring-blue-500/10"
                    />
                  </div>
                </div>

                <div className="grid grid-cols-1 gap-3 sm:grid-cols-2">
                  <div>
                    <label className="mb-1.5 block text-[12px] font-semibold text-slate-200">
                      Password
                    </label>

                    <div className="relative">
                      <input
                        type={showPassword ? 'text' : 'password'}
                        value={password}
                        onChange={(e) => setPassword(e.target.value)}
                        placeholder="Create password"
                        className="w-full rounded-xl border border-slate-700 bg-slate-800/80 px-3 py-3 pr-12 text-sm text-white shadow-sm outline-none transition-all duration-200 placeholder:text-slate-500 hover:border-slate-600 focus:border-blue-500 focus:ring-4 focus:ring-blue-500/10"
                        autoComplete="new-password"
                      />

                      <button
                        type="button"
                        onClick={() =>
                          setShowPassword((value) => !value)
                        }
                        className="absolute right-3 top-1/2 -translate-y-1/2 rounded-lg p-1.5 text-slate-400 transition hover:bg-slate-700 hover:text-blue-400 focus:outline-none focus:ring-2 focus:ring-blue-500/30"
                        aria-label={
                          showPassword
                            ? 'Hide password'
                            : 'Show password'
                        }
                      >
                        {showPassword ? (
                          <svg
                            xmlns="http://www.w3.org/2000/svg"
                            viewBox="0 0 24 24"
                            fill="none"
                            stroke="currentColor"
                            strokeWidth="2"
                            className="h-5 w-5"
                          >
                            <path d="M3 3l18 18" />
                            <path d="M10.58 10.58a2 2 0 0 0 2.83 2.83" />
                            <path d="M9.88 4.24A10.94 10.94 0 0 1 12 4c5 0 8.5 4 10 8a16.6 16.6 0 0 1-3.03 4.68" />
                            <path d="M6.61 6.61C4.62 7.88 3.25 9.72 2 12c1.5 4 5 8 10 8a10.94 10.94 0 0 0 4.12-.76" />
                          </svg>
                        ) : (
                          <svg
                            xmlns="http://www.w3.org/2000/svg"
                            viewBox="0 0 24 24"
                            fill="none"
                            stroke="currentColor"
                            strokeWidth="2"
                            className="h-5 w-5"
                          >
                            <path d="M2 12s3.5-7 10-7 10 7 10 7-3.5 7-10 7S2 12 2 12Z" />
                            <circle cx="12" cy="12" r="3" />
                          </svg>
                        )}
                      </button>
                    </div>
                  </div>

                  <div>
                    <label className="mb-1.5 block text-[12px] font-semibold text-slate-200">
                      Confirm Password
                    </label>

                    <div className="relative">
                      <input
                        type={
                          showConfirmPassword
                            ? 'text'
                            : 'password'
                        }
                        value={confirmPassword}
                        onChange={(e) =>
                          setConfirmPassword(e.target.value)
                        }
                        placeholder="Confirm password"
                        className="w-full rounded-xl border border-slate-700 bg-slate-800/80 px-3 py-3 pr-12 text-sm text-white shadow-sm outline-none transition-all duration-200 placeholder:text-slate-500 hover:border-slate-600 focus:border-blue-500 focus:ring-4 focus:ring-blue-500/10"
                        autoComplete="new-password"
                      />

                      <button
                        type="button"
                        onClick={() =>
                          setShowConfirmPassword(
                            (value) => !value
                          )
                        }
                        className="absolute right-3 top-1/2 -translate-y-1/2 rounded-lg p-1.5 text-slate-400 transition hover:bg-slate-700 hover:text-blue-400 focus:outline-none focus:ring-2 focus:ring-blue-500/30"
                        aria-label={
                          showConfirmPassword
                            ? 'Hide confirm password'
                            : 'Show confirm password'
                        }
                      >
                        {showConfirmPassword ? (
                          <svg
                            xmlns="http://www.w3.org/2000/svg"
                            viewBox="0 0 24 24"
                            fill="none"
                            stroke="currentColor"
                            strokeWidth="2"
                            className="h-5 w-5"
                          >
                            <path d="M3 3l18 18" />
                            <path d="M10.58 10.58a2 2 0 0 0 2.83 2.83" />
                            <path d="M9.88 4.24A10.94 10.94 0 0 1 12 4c5 0 8.5 4 10 8a16.6 16.6 0 0 1-3.03 4.68" />
                            <path d="M6.61 6.61C4.62 7.88 3.25 9.72 2 12c1.5 4 5 8 10 8a10.94 10.94 0 0 0 4.12-.76" />
                          </svg>
                        ) : (
                          <svg
                            xmlns="http://www.w3.org/2000/svg"
                            viewBox="0 0 24 24"
                            fill="none"
                            stroke="currentColor"
                            strokeWidth="2"
                            className="h-5 w-5"
                          >
                            <path d="M2 12s3.5-7 10-7 10 7 10 7-3.5 7-10 7S2 12 2 12Z" />
                            <circle cx="12" cy="12" r="3" />
                          </svg>
                        )}
                      </button>
                    </div>
                  </div>
                </div>

                <div className="rounded-xl border border-slate-200 bg-slate-50 px-4 py-3 shadow-sm">
                  <div className="mb-2 flex items-center gap-2">
                    <span className="flex h-5 w-5 items-center justify-center rounded-full bg-blue-100 text-[10px] font-bold text-blue-600">
                      ✓
                    </span>

                    <p className="text-xs font-semibold text-slate-700">
                      Password must contain:
                    </p>
                  </div>

                  <div className="grid grid-cols-1 gap-x-6 gap-y-1.5 sm:grid-cols-2">
                    <PasswordRequirement
                      met={password.length >= 8}
                      text="8+ characters"
                    />

                    <PasswordRequirement
                      met={/[A-Z]/.test(password)}
                      text="Uppercase letter"
                    />

                    <PasswordRequirement
                      met={/[a-z]/.test(password)}
                      text="Lowercase letter"
                    />

                    <PasswordRequirement
                      met={/[0-9]/.test(password)}
                      text="Number"
                    />

                    <PasswordRequirement
                      met={/[^A-Za-z0-9]/.test(password)}
                      text="Special character"
                    />
                  </div>
                </div>

                <div className="flex items-start gap-3 rounded-xl border border-blue-100 bg-blue-50/70 px-4 py-3 shadow-sm">
                  <div className="flex h-8 w-8 shrink-0 items-center justify-center rounded-lg bg-blue-100 text-blue-600">
                    <svg
                      xmlns="http://www.w3.org/2000/svg"
                      className="h-4 w-4"
                      viewBox="0 0 24 24"
                      fill="none"
                      stroke="currentColor"
                      strokeWidth="2"
                      strokeLinecap="round"
                      strokeLinejoin="round"
                    >
                      <path d="M12 3l8 4v5c0 5-3.5 8.5-8 10-4.5-1.5-8-5-8-10V7l8-4z" />
                      <path d="M9 12l2 2 4-4" />
                    </svg>
                  </div>

                  <div className="min-w-0">
                    <p className="text-sm font-semibold text-blue-900">
                      Administrator Account
                    </p>

                    <p className="mt-0.5 text-xs leading-4 text-blue-700">
                      This account will have full access to the Medical Store Management System.
                    </p>
                  </div>
                </div>

                {error && (
                  <div className="flex items-start gap-3 rounded-xl border border-red-200 bg-red-50 px-4 py-3">
                    <div className="mt-0.5 flex h-5 w-5 shrink-0 items-center justify-center rounded-full bg-red-100 text-red-600">
                      <span className="text-xs font-bold">
                        !
                      </span>
                    </div>

                    <p className="text-sm font-medium leading-5 text-red-700">
                      {error}
                    </p>
                  </div>
                )}

                {success && (
                  <div className="flex items-start gap-3 rounded-xl border border-emerald-200 bg-emerald-50 px-4 py-3">
                    <div className="mt-0.5 flex h-5 w-5 shrink-0 items-center justify-center rounded-full bg-emerald-100 text-emerald-600">
                      <span className="text-xs font-bold">
                        ✓
                      </span>
                    </div>

                    <p className="text-sm font-medium leading-5 text-emerald-700">
                      {success}
                    </p>
                  </div>
                )}

                <button
                  type="submit"
                  disabled={creatingAccount}
                  className="w-full rounded-xl bg-gradient-to-r from-blue-500 via-blue-600 to-blue-700 px-4 py-3 text-sm font-semibold text-white shadow-[0_10px_24px_rgba(37,99,235,0.28)] transition-all duration-200 hover:-translate-y-0.5 hover:from-blue-600 hover:via-blue-700 hover:to-blue-800 hover:shadow-[0_14px_30px_rgba(37,99,235,0.35)] active:translate-y-0 active:shadow-[0_6px_16px_rgba(37,99,235,0.22)] focus:outline-none focus:ring-4 focus:ring-blue-500/20 disabled:cursor-not-allowed disabled:opacity-50"
                >
                  {creatingAccount
                    ? 'Creating account...'
                    : 'Create Admin Account'}
                </button>
              </form>
            ) : (
              <form
                onSubmit={handleLogin}
                className="mx-auto w-full max-w-lg space-y-4"
              >
                <div>
                  <label className="mb-1.5 block text-[12px] font-semibold text-slate-200">
                    Username
                  </label>

                  <input
                    type="text"
                    value={username}
                    onChange={(e) =>
                      setUsername(e.target.value)
                    }
                    placeholder="Enter username"
                    autoComplete="username"
                    className="w-full rounded-xl border border-slate-700 bg-slate-800/80 px-3 py-3 text-sm text-white shadow-sm outline-none transition-all duration-200 placeholder:text-slate-500 hover:border-slate-600 focus:border-blue-500 focus:ring-4 focus:ring-blue-500/10"
                  />
                </div>

                <div>
                  <label className="mb-1.5 block text-[12px] font-semibold text-slate-200">
                    Password
                  </label>

                  <div className="relative">
                    <input
                      type={showPassword ? 'text' : 'password'}
                      value={password}
                      onChange={(e) => setPassword(e.target.value)}
                      placeholder="Create a password"
                      autoComplete="new-password"
                      className="w-full rounded-xl border border-slate-700 bg-slate-800/80 px-3 py-3 pr-12 text-sm text-white shadow-sm outline-none transition-all duration-200 placeholder:text-slate-500 hover:border-slate-600 focus:border-blue-500 focus:ring-4 focus:ring-blue-500/10"
                    />

                    <button
                      type="button"
                      onClick={() => setShowPassword((value) => !value)}
                      className="absolute right-3 top-1/2 -translate-y-1/2 rounded-lg p-1.5 text-slate-400 transition hover:bg-slate-700 hover:text-blue-400 focus:outline-none focus:ring-2 focus:ring-blue-100"
                      aria-label={showPassword ? 'Hide password' : 'Show password'}
                    >
                      {showPassword ? (
                        <svg
                          xmlns="http://www.w3.org/2000/svg"
                          viewBox="0 0 24 24"
                          fill="none"
                          stroke="currentColor"
                          strokeWidth="2"
                          className="h-5 w-5"
                        >
                          <path d="M3 3l18 18" />
                          <path d="M10.58 10.58a2 2 0 0 0 2.83 2.83" />
                          <path d="M9.88 4.24A10.94 10.94 0 0 1 12 4c5 0 8.5 4 10 8a16.6 16.6 0 0 1-3.03 4.68" />
                          <path d="M6.61 6.61C4.62 7.88 3.25 9.72 2 12c1.5 4 5 8 10 8a10.94 10.94 0 0 0 4.12-.76" />
                        </svg>
                      ) : (
                        <svg
                          xmlns="http://www.w3.org/2000/svg"
                          viewBox="0 0 24 24"
                          fill="none"
                          stroke="currentColor"
                          strokeWidth="2"
                          className="h-5 w-5"
                        >
                          <path d="M2 12s3.5-7 10-7 10 7 10 7-3.5 7-10 7S2 12 2 12Z" />
                          <circle cx="12" cy="12" r="3" />
                        </svg>
                      )}
                    </button>
                  </div>
                </div>

                {error && (
                  <div className="flex items-start gap-3 rounded-xl border border-red-200 bg-red-50 px-4 py-3">
                    <div className="mt-0.5 flex h-5 w-5 shrink-0 items-center justify-center rounded-full bg-red-100 text-red-600">
                      <span className="text-xs font-bold">
                        !
                      </span>
                    </div>

                    <p className="text-sm font-medium leading-5 text-red-700">
                      {error}
                    </p>
                  </div>
                )}

                {success && (
                  <div className="flex items-start gap-3 rounded-xl border border-emerald-200 bg-emerald-50 px-4 py-3">
                    <div className="mt-0.5 flex h-5 w-5 shrink-0 items-center justify-center rounded-full bg-emerald-100 text-emerald-600">
                      <span className="text-xs font-bold">
                        ✓
                      </span>
                    </div>

                    <p className="text-sm font-medium leading-5 text-emerald-700">
                      {success}
                    </p>
                  </div>
                )}

                <button
                  type="submit"
                  disabled={loggingIn}
                  className="w-full rounded-xl bg-gradient-to-r from-blue-500 via-blue-600 to-blue-700 px-4 py-3 text-sm font-semibold text-white shadow-[0_10px_24px_rgba(37,99,235,0.28)] transition-all duration-200 hover:-translate-y-0.5 hover:from-blue-600 hover:via-blue-700 hover:to-blue-800 hover:shadow-[0_14px_30px_rgba(37,99,235,0.35)] active:translate-y-0 active:shadow-[0_6px_16px_rgba(37,99,235,0.22)] focus:outline-none focus:ring-4 focus:ring-blue-500/20 disabled:cursor-not-allowed disabled:opacity-50"
                >
                  {loggingIn
                    ? 'Signing in...'
                    : 'Sign In'}
                </button>
              </form>
            )}

            {!isSignup && !hasUsers && (
              <div className="mx-auto mt-4 w-full max-w-lg text-center">
                <p className="text-xs text-slate-500">
                  No account has been registered
                  yet.
                </p>

                <button
                  type="button"
                  onClick={() =>
                    switchMode('signup')
                  }
                  className="mt-2 text-sm font-semibold transition-colors hover:text-blue-700"
                >
                  Create the first account
                </button>
              </div>
            )}
          </div>

          <div className="border-t border-slate-200/80 bg-slate-50/80 p-5">
            <div className="flex items-center justify-center gap-2">
              <span className="h-2 w-2 rounded-full bg-emerald-500"></span>

              <p className="text-xs font-semibold text-slate-600">
                Offline Authentication
              </p>
            </div>

            <p className="mt-1.5 text-center text-xs leading-5 text-slate-400">
              Your account and login data are
              handled locally on this device.
            </p>
          </div>
        </div>
      </div>
    </div>
  )
}

/* --------------------------------
   Backup & Restore
-------------------------------- */
function BackupRestore({
  showNotification,
}) {
  const [backupPassword, setBackupPassword] = useState('')
  const [showBackupPassword, setShowBackupPassword] = useState(false)
  const [selectedFile, setSelectedFile] = useState(null)

  const [backupHistory, setBackupHistory] = useState([])
  const [backupHistoryPage, setBackupHistoryPage] = useState(1)
  const [isLoadingBackupHistory, setIsLoadingBackupHistory] = useState(false)
  const [restorePreview, setRestorePreview] = useState(null)
  const [isCreatingBackup, setIsCreatingBackup] = useState(false)
  const [isPreviewing, setIsPreviewing] = useState(false)
  const [isRestoring, setIsRestoring] = useState(false)

  const BACKUPS_PER_PAGE = 5

  const totalBackupPages = Math.ceil(
    backupHistory.length / BACKUPS_PER_PAGE
  )

  const backupHistoryStartIndex =
    (backupHistoryPage - 1) * BACKUPS_PER_PAGE

  const paginatedBackupHistory = backupHistory.slice(
    backupHistoryStartIndex,
    backupHistoryStartIndex + BACKUPS_PER_PAGE
  )

  const loadBackupHistory = async () => {
    try {
      setIsLoadingBackupHistory(true)

      const history = await getBackupHistory()

      setBackupHistory(history || [])
      setBackupHistoryPage(1)
    } catch (error) {
      console.error('Failed to load backup history:', error)
    } finally {
      setIsLoadingBackupHistory(false)
    }
  }
  useEffect(() => {
    loadBackupHistory()
  }, [])

  function handleFileChange(event) {
    const file = event.target.files?.[0] || null

    setSelectedFile(file)
    setRestorePreview(null)
  }

  async function handleCreateBackup() {
    if (!backupPassword) {
      showNotification(
        'error',
        'Please enter a backup password.'
      )
      return
    }

    if (backupPassword.length < 8) {
      showNotification(
        'error',
        'Backup password must be at least 8 characters.'
      )
      return
    }

    try {
      setIsCreatingBackup(true)

      await generateEncryptedBackupFile(
        backupPassword
      )

      await loadBackupHistory()

      showNotification(
        'success',
        'Encrypted backup created successfully.'
      )
    } catch (error) {
      console.error(
        'Backup creation failed:',
        error
      )

      showNotification(
        'error',
        'Backup could not be created. Please try again.'
      )
    } finally {
      setIsCreatingBackup(false)
    }
  }

  async function handlePreviewRestore() {
    if (!selectedFile) {
      showNotification(
        'error',
        'Please select a .kmsbackup file first.'
      )
      return
    }

    if (!backupPassword) {
      showNotification(
        'error',
        'Please enter the backup password.'
      )
      return
    }

    try {
      setIsPreviewing(true)
      setRestorePreview(null)

      const result =
        await previewEncryptedBackupRestore(
          selectedFile,
          backupPassword
        )

      setRestorePreview(result)

      if (!result?.valid) {
        showNotification(
          'error',
          'The selected backup is not valid.'
        )
      }
    } catch (error) {
      console.error(
        'Backup preview failed:',
        error
      )

      showNotification(
        'error',
        error?.message ||
        'Could not read the backup file.'
      )
    } finally {
      setIsPreviewing(false)
    }
  }

  async function handleRestore() {
    if (!selectedFile || !restorePreview?.valid) {
      showNotification(
        'error',
        'Please preview a valid backup before restoring.'
      )
      return
    }

    const confirmed = window.confirm(
      'WARNING: Restoring this backup will replace the current local data. A safety backup will be created first. Continue?'
    )

    if (!confirmed) {
      return
    }

    try {
      setIsRestoring(true)

      await confirmEncryptedBackupRestore(
        selectedFile,
        backupPassword,
        true
      )

      showNotification(
        'success',
        'Backup restored successfully. Please login again.'
      )

      setTimeout(() => {
        window.location.reload()
      }, 1200)
    } catch (error) {
      console.error(
        'Backup restore failed:',
        error
      )

      showNotification(
        'error',
        error?.message ||
        'Backup could not be restored.'
      )
    } finally {
      setIsRestoring(false)
    }
  }

  return (
    <PageContainer>
      <div className="space-y-5">

        {/* Header */}
        <div>
          <div className="flex items-center gap-3">
            <div className="flex h-11 w-11 items-center justify-center rounded-xl bg-blue-50 text-blue-600">
              <svg
                width="22"
                height="22"
                viewBox="0 0 24 24"
                fill="none"
                stroke="currentColor"
                strokeWidth="2"
                strokeLinecap="round"
                strokeLinejoin="round"
              >
                <path d="M12 3v12" />
                <path d="m7 10 5 5 5-5" />
                <path d="M5 21h14" />
                <path d="M5 17v4" />
                <path d="M19 17v4" />
              </svg>
            </div>

            <div>
              <h1 className="text-xl font-bold tracking-tight text-slate-900 sm:text-2xl">
                Backup & Restore
              </h1>

              <p className="mt-1 text-sm text-slate-500">
                Create a secure backup or restore your
                medical store data.
              </p>
            </div>
          </div>
        </div>

        {/* Backup Password */}
        <div className={`${ui.card} p-5 sm:p-6`}>
          <div className="flex items-start gap-3">
            <div className="flex h-10 w-10 shrink-0 items-center justify-center rounded-xl bg-blue-50 text-blue-600">
              🔐
            </div>

            <div>
              <h2 className="text-base font-bold text-slate-900">
                Backup Password
              </h2>

              <p className="mt-1 text-sm leading-6 text-slate-500">
                This password encrypts and protects your
                backup file. Keep it somewhere safe.
              </p>
            </div>
          </div>

          <div className="mt-5 max-w-xl">
            <label className={ui.label}>
              Password
            </label>

            <div className="relative">
              <input
                type={showBackupPassword ? 'text' : 'password'}
                value={backupPassword}
                onChange={(event) =>
                  setBackupPassword(event.target.value)
                }
                placeholder="Enter backup password"
                className={`${ui.input} pr-12`}
              />

              <button
                type="button"
                onClick={() =>
                  setShowBackupPassword((previous) => !previous)
                }
                className="absolute right-3 top-1/2 -translate-y-1/2 text-slate-400 transition hover:text-slate-600"
                aria-label={
                  showBackupPassword
                    ? 'Hide password'
                    : 'Show password'
                }
              >
                {showBackupPassword ? '🙈' : '👁️'}
              </button>
            </div>

            <p className="mt-2 text-xs text-slate-400">
              Minimum 8 characters.
            </p>
          </div>
        </div>

        {/* Create Backup */}
        <div className={`${ui.card} overflow-hidden`}>
          <div className="p-5 sm:p-6">
            <div className="flex flex-col gap-5 lg:flex-row lg:items-center lg:justify-between">

              <div className="flex items-start gap-4">
                <div className="flex h-12 w-12 shrink-0 items-center justify-center rounded-xl bg-blue-50 text-xl">
                  💾
                </div>

                <div>
                  <h2 className="text-lg font-bold text-slate-900">
                    Create Backup
                  </h2>

                  <p className="mt-1 max-w-2xl text-sm leading-6 text-slate-500">
                    Create an encrypted backup of your medical
                    store data and save it as a secure
                    <span className="font-semibold text-slate-600">
                      {' '}
                      .kmsbackup
                    </span>{' '}
                    file.
                  </p>
                </div>
              </div>

              <button
                type="button"
                onClick={handleCreateBackup}
                disabled={
                  isCreatingBackup ||
                  !backupPassword
                }
                className={`${ui.primaryButton} inline-flex w-full items-center justify-center gap-2 lg:w-auto lg:min-w-[190px] ${isCreatingBackup
                  ? 'animate-[processingPulse_1.2s_ease-in-out_infinite]'
                  : ''
                  }`}
              >
                <span className="text-base">↓</span>

                {isCreatingBackup
                  ? 'Creating...'
                  : 'Create Backup'}
              </button>
            </div>
          </div>

          <div className="border-t border-slate-100 bg-slate-50/70 px-5 py-3 sm:px-6">
            <div className="flex flex-wrap items-center gap-x-5 gap-y-2 text-xs text-slate-500">
              <span className="inline-flex items-center gap-1.5">
                <span className="text-emerald-600">✓</span>
                Encrypted
              </span>

              <span className="inline-flex items-center gap-1.5">
                <span className="text-emerald-600">✓</span>
                Local file
              </span>

              <span className="inline-flex items-center gap-1.5">
                <span className="text-emerald-600">✓</span>
                No active session included
              </span>
            </div>
          </div>
        </div>

        {/* Backup History */}
        <div className={`${ui.card} overflow-hidden`}>
          <div className="p-5 sm:p-6">
            <div className="flex flex-col gap-4 sm:flex-row sm:items-start sm:justify-between">
              <div className="flex items-start gap-4">
                <div className="flex h-12 w-12 shrink-0 items-center justify-center rounded-xl bg-blue-50 text-xl">
                  🕘
                </div>

                <div>
                  <div className="flex flex-wrap items-center gap-2">
                    <h2 className="text-lg font-bold text-slate-900">
                      Backup History
                    </h2>

                    <span className="inline-flex items-center rounded-full bg-blue-50 px-2.5 py-1 text-xs font-bold text-blue-700">
                      {backupHistory.length}{' '}
                      {backupHistory.length === 1
                        ? 'Backup'
                        : 'Backups'}
                    </span>
                  </div>

                  <p className="mt-1 text-sm leading-6 text-slate-500">
                    Previously created encrypted backups from this
                    application.
                  </p>
                </div>
              </div>

              <button
                type="button"
                onClick={loadBackupHistory}
                disabled={isLoadingBackupHistory}
                className="inline-flex w-full items-center justify-center gap-2 rounded-xl border border-blue-200 bg-white px-4 py-2.5 text-sm font-semibold text-blue-700 transition hover:bg-blue-50 disabled:cursor-not-allowed disabled:opacity-60 sm:w-auto"
              >
                <span>
                  {isLoadingBackupHistory ? '↻' : '⟳'}
                </span>

                {isLoadingBackupHistory
                  ? 'Refreshing...'
                  : 'Refresh'}
              </button>
            </div>

            <div className="mt-5">
              {isLoadingBackupHistory ? (
                <div className="rounded-xl border border-slate-200 bg-slate-50 px-4 py-6 text-center">
                  <p className="text-sm font-medium text-slate-500">
                    Loading backup history...
                  </p>
                </div>
              ) : backupHistory.length === 0 ? (
                <div className="rounded-2xl border border-dashed border-blue-200 bg-blue-50/40 px-5 py-9 text-center">
                  <div className="mx-auto flex h-14 w-14 items-center justify-center rounded-2xl bg-white text-2xl shadow-sm">
                    💾
                  </div>

                  <p className="mt-4 text-sm font-bold text-slate-800">
                    No backup history yet
                  </p>

                  <p className="mx-auto mt-1.5 max-w-md text-xs leading-5 text-slate-500">
                    Once you create an encrypted backup, its file name,
                    date, size, and backup type will appear here.
                  </p>
                </div>
              ) : (
                <>
                  <div className="overflow-x-auto rounded-xl border border-slate-200">
                    <table className="min-w-[720px] w-full divide-y divide-slate-200">
                      <thead className="bg-slate-50">
                        <tr>
                          <th className="px-4 py-3 text-left text-xs font-semibold uppercase tracking-wide text-slate-500">
                            Backup File
                          </th>

                          <th className="px-4 py-3 text-left text-xs font-semibold uppercase tracking-wide text-slate-500">
                            Created
                          </th>

                          <th className="px-4 py-3 text-left text-xs font-semibold uppercase tracking-wide text-slate-500">
                            Size
                          </th>

                          <th className="px-4 py-3 text-left text-xs font-semibold uppercase tracking-wide text-slate-500">
                            Type
                          </th>
                        </tr>
                      </thead>

                      <tbody className="divide-y divide-slate-100 bg-white">
                        {paginatedBackupHistory.map((backup, index) => (
                          <tr
                            key={backup.id}
                            className={
                              backupHistoryStartIndex + index === 0
                                ? 'bg-blue-50/40'
                                : 'bg-white'
                            }
                          >
                            <td className="px-4 py-3">
                              <div className="flex items-center gap-3">
                                <div className="flex h-9 w-9 shrink-0 items-center justify-center rounded-lg bg-blue-50 text-blue-600">
                                  💾
                                </div>

                                <div className="min-w-0">
                                  <div className="flex flex-wrap items-center gap-2">
                                    <p className="max-w-[220px] truncate text-sm font-semibold text-slate-800 sm:max-w-[260px]">
                                      {backup.fileName || 'Backup file'}
                                    </p>

                                    {backupHistoryStartIndex + index === 0 && (
                                      <span className="shrink-0 rounded-full bg-blue-100 px-2 py-0.5 text-[10px] font-bold uppercase tracking-wide text-blue-700">
                                        Latest
                                      </span>
                                    )}
                                  </div>

                                  <p className="mt-0.5 text-xs text-slate-400">
                                    {backup.type === 'safety'
                                      ? 'Created automatically before restore'
                                      : 'Created manually'}
                                  </p>
                                </div>
                              </div>
                            </td>

                            <td className="whitespace-nowrap px-4 py-3">
                              {backup.createdAt ? (
                                <div>
                                  <p className="text-sm font-semibold text-slate-700">
                                    {new Date(
                                      backup.createdAt
                                    ).toLocaleDateString(undefined, {
                                      day: '2-digit',
                                      month: 'short',
                                      year: 'numeric',
                                    })}
                                  </p>

                                  <p className="mt-0.5 text-xs text-slate-400">
                                    {new Date(
                                      backup.createdAt
                                    ).toLocaleTimeString(undefined, {
                                      hour: '2-digit',
                                      minute: '2-digit',
                                    })}
                                  </p>
                                </div>
                              ) : (
                                <span className="text-sm text-slate-400">
                                  Unknown
                                </span>
                              )}
                            </td>

                            <td className="whitespace-nowrap px-4 py-3">
                              {typeof backup.size === 'number' ? (
                                backup.size >= 1024 * 1024 ? (
                                  <span className="text-sm font-semibold text-slate-700">
                                    {(backup.size / (1024 * 1024)).toFixed(2)} MB
                                  </span>
                                ) : (
                                  <span className="text-sm font-semibold text-slate-700">
                                    {(backup.size / 1024).toFixed(1)} KB
                                  </span>
                                )
                              ) : (
                                <span className="text-sm text-slate-400">
                                  Unknown
                                </span>
                              )}
                            </td>

                            <td className="px-4 py-3">
                              {backup.type === 'safety' ? (
                                <span className="inline-flex rounded-full bg-amber-50 px-2.5 py-1 text-xs font-semibold text-amber-700">
                                  Safety Backup
                                </span>
                              ) : (
                                <span className="inline-flex rounded-full bg-emerald-50 px-2.5 py-1 text-xs font-semibold text-emerald-700">
                                  Encrypted Backup
                                </span>
                              )}
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                  {totalBackupPages > 1 ? (
                    <div className="mt-4 flex items-center justify-between gap-3">
                      <button
                        type="button"
                        onClick={() =>
                          setBackupHistoryPage((page) =>
                            Math.max(page - 1, 1)
                          )
                        }
                        disabled={backupHistoryPage === 1}
                        className="inline-flex items-center gap-2 rounded-xl border border-slate-200 bg-white px-3 py-2.5 text-sm font-semibold text-slate-700 transition hover:border-blue-300 hover:bg-blue-50 hover:text-blue-700 disabled:cursor-not-allowed disabled:opacity-40"
                        aria-label="Previous page"
                      >
                        <span className="text-base">←</span>
                        <span className="hidden sm:inline">
                          Previous
                        </span>
                      </button>

                      <span className="text-xs font-semibold text-slate-500 sm:text-sm">
                        Page {backupHistoryPage} of {totalBackupPages}
                      </span>

                      <button
                        type="button"
                        onClick={() =>
                          setBackupHistoryPage((page) =>
                            Math.min(page + 1, totalBackupPages)
                          )
                        }
                        disabled={
                          backupHistoryPage === totalBackupPages
                        }
                        className="inline-flex items-center gap-2 rounded-xl border border-blue-200 bg-blue-50 px-3 py-2.5 text-sm font-semibold text-blue-700 transition hover:bg-blue-100 disabled:cursor-not-allowed disabled:opacity-40"
                        aria-label="Next page"
                      >
                        <span className="hidden sm:inline">
                          Next
                        </span>
                        <span className="text-base">→</span>
                      </button>
                    </div>
                  ) : null}
                </>
              )}
            </div>
          </div>

          <div className="border-t border-slate-100 bg-slate-50/70 px-5 py-3 sm:px-6">
            <p className="text-xs leading-5 text-slate-500">
              Backup history stores metadata only. It records when a backup
              download was initiated. Backup passwords and encrypted backup
              contents are never stored in backup history.
            </p>
          </div>
        </div>

        {/* Restore */}
        <div className={`${ui.card} p-5 sm:p-6`}>
          <div className="flex items-start gap-4">
            <div className="flex h-12 w-12 shrink-0 items-center justify-center rounded-xl bg-slate-100 text-xl">
              📂
            </div>

            <div>
              <h2 className="text-lg font-bold text-slate-900">
                Restore Backup
              </h2>

              <p className="mt-1 max-w-2xl text-sm leading-6 text-slate-500">
                Select a previously created
                <span className="font-semibold text-slate-600">
                  {' '}
                  .kmsbackup
                </span>{' '}
                file, verify its contents, and restore your
                store data safely.
              </p>
            </div>
          </div>

          <div className="mt-5">
            <label className={ui.label}>
              Backup File
            </label>

            <input
              type="file"
              accept=".kmsbackup,application/json"
              onChange={handleFileChange}
              className="block w-full rounded-xl border border-slate-200 bg-white px-3 py-3 text-sm text-slate-700 file:mr-4 file:rounded-lg file:border-0 file:bg-blue-50 file:px-3 file:py-2 file:text-sm file:font-semibold file:text-blue-700 hover:file:bg-blue-100"
            />

            {selectedFile && (
              <div className="mt-3 flex items-center gap-3 rounded-xl border border-blue-200 bg-blue-50/60 px-4 py-3">
                <div className="flex h-10 w-10 shrink-0 items-center justify-center rounded-xl bg-blue-100 text-blue-600">
                  📁
                </div>

                <div className="min-w-0 flex-1">
                  <p className="text-xs font-semibold uppercase tracking-wide text-blue-600">
                    Selected Backup
                  </p>

                  <p className="mt-1 truncate text-sm font-semibold text-slate-800">
                    {selectedFile.name}
                  </p>

                  <p className="mt-1 text-xs text-slate-500">
                    {(selectedFile.size / 1024).toFixed(1)} KB
                  </p>
                </div>

                <button
                  type="button"
                  onClick={() => {
                    setSelectedFile(null)
                    setRestorePreview(null)
                  }}
                  className="shrink-0 rounded-lg px-3 py-2 text-xs font-semibold text-slate-500 transition hover:bg-white hover:text-red-600"
                >
                  Remove
                </button>
              </div>
            )}
          </div>

          <div className="mt-5 flex flex-wrap gap-3">
            <button
              type="button"
              onClick={handlePreviewRestore}
              disabled={
                isPreviewing ||
                !selectedFile ||
                !backupPassword
              }
              className={`${ui.secondaryButton} inline-flex items-center justify-center gap-2`}
            >
              🔍
              {isPreviewing
                ? 'Checking...'
                : 'Preview Backup'}
            </button>
          </div>

          {/* Preview Result */}
          {restorePreview && (
            <div className="mt-6 rounded-2xl border border-blue-200 bg-blue-50/60 p-5">
              {restorePreview.valid ? (
                <>
                  <div className="flex items-start justify-between gap-4">
                    <div className="flex items-start gap-3">
                      <div className="flex h-10 w-10 shrink-0 items-center justify-center rounded-xl bg-emerald-100 text-emerald-700">
                        ✓
                      </div>

                      <div>
                        <h3 className="font-bold text-slate-900">
                          Backup is valid
                        </h3>

                        <p className="mt-1 text-sm text-slate-500">
                          The backup passed all validation checks and
                          is ready for restore.
                        </p>
                      </div>
                    </div>

                    <span className="shrink-0 rounded-full bg-emerald-100 px-3 py-1 text-xs font-bold text-emerald-700">
                      Verified
                    </span>
                  </div>

                  <div className="mt-5 grid grid-cols-1 gap-3 sm:grid-cols-2 lg:grid-cols-4">
                    <div className="rounded-xl border border-slate-200 bg-white p-4">
                      <p className="text-xs font-semibold uppercase tracking-wide text-slate-400">
                        Daily Sessions
                      </p>

                      <p className="mt-1 text-xl font-bold text-slate-900">
                        {restorePreview.dailySessionsCount ?? 0}
                      </p>
                    </div>

                    <div className="rounded-xl border border-slate-200 bg-white p-4">
                      <p className="text-xs font-semibold uppercase tracking-wide text-slate-400">
                        Transactions
                      </p>

                      <p className="mt-1 text-xl font-bold text-slate-900">
                        {restorePreview.transactionsCount ?? 0}
                      </p>
                    </div>

                    <div className="rounded-xl border border-slate-200 bg-white p-4">
                      <p className="text-xs font-semibold uppercase tracking-wide text-slate-400">
                        Users
                      </p>

                      <p className="mt-1 text-xl font-bold text-slate-900">
                        {restorePreview.usersCount ?? 0}
                      </p>
                    </div>

                    <div className="rounded-xl border border-slate-200 bg-white p-4">
                      <p className="text-xs font-semibold uppercase tracking-wide text-slate-400">
                        Audit Logs
                      </p>

                      <p className="mt-1 text-xl font-bold text-slate-900">
                        {restorePreview.auditLogsCount ?? 0}
                      </p>
                    </div>
                  </div>

                  <div className="mt-5 overflow-hidden rounded-2xl border border-amber-200 bg-amber-50">
                    <div className="flex items-start gap-3 p-4">
                      <div className="flex h-10 w-10 shrink-0 items-center justify-center rounded-xl bg-amber-100 text-amber-700">
                        ⚠️
                      </div>

                      <div>
                        <p className="text-sm font-bold text-amber-900">
                          Before restoring this backup
                        </p>

                        <p className="mt-1 text-sm leading-6 text-amber-800">
                          Restoring will replace the current local
                          database with the data contained in this
                          backup.
                        </p>
                      </div>
                    </div>

                    <div className="border-t border-amber-200/70 bg-amber-100/40 px-4 py-3">
                      <div className="flex flex-wrap gap-x-5 gap-y-2 text-xs font-medium text-amber-800">
                        <span>✓ Safety backup created first</span>
                        <span>✓ Backup must pass validation</span>
                        <span>✓ Manual login required after restore</span>
                      </div>
                    </div>
                  </div>

                  <button
                    type="button"
                    onClick={handleRestore}
                    disabled={isRestoring}
                    className={`${ui.dangerButton} mt-5 inline-flex items-center justify-center gap-2 ${isRestoring
                      ? 'animate-[processingPulse_1.2s_ease-in-out_infinite]'
                      : ''
                      }`}
                  >
                    {isRestoring
                      ? 'Restoring...'
                      : 'Restore This Backup'}
                  </button>
                </>
              ) : (
                <div className="flex items-start gap-3">
                  <div className="flex h-10 w-10 shrink-0 items-center justify-center rounded-xl bg-red-100 text-red-600">
                    !
                  </div>

                  <div>
                    <h3 className="font-bold text-red-800">
                      Invalid Backup
                    </h3>

                    <p className="mt-1 text-sm leading-6 text-red-700">
                      This backup could not be validated.
                      Check the file and password and try
                      again.
                    </p>
                  </div>
                </div>
              )}
            </div>
          )}
        </div>

        {/* Security Note */}
        <div className="overflow-hidden rounded-2xl border border-blue-100 bg-blue-50/50">
          <div className="flex items-start gap-4 p-5 sm:p-6">
            <div className="flex h-11 w-11 shrink-0 items-center justify-center rounded-xl bg-blue-100 text-lg">
              🛡️
            </div>

            <div>
              <h3 className="text-sm font-bold text-slate-900">
                Backup Security
              </h3>

              <p className="mt-1 text-sm leading-6 text-slate-600">
                Your backup is encrypted before it is saved.
                The active login session is intentionally
                excluded from the backup.
              </p>
            </div>
          </div>

          <div className="grid grid-cols-1 border-t border-blue-100 sm:grid-cols-3">
            <div className="px-5 py-4 sm:border-r sm:border-blue-100">
              <p className="text-xs font-semibold uppercase tracking-wide text-blue-600">
                Encryption
              </p>

              <p className="mt-1 text-sm font-semibold text-slate-700">
                AES-GCM 256-bit
              </p>
            </div>

            <div className="px-5 py-4 sm:border-r sm:border-blue-100">
              <p className="text-xs font-semibold uppercase tracking-wide text-blue-600">
                Key Protection
              </p>

              <p className="mt-1 text-sm font-semibold text-slate-700">
                Password-based
              </p>
            </div>

            <div className="px-5 py-4">
              <p className="text-xs font-semibold uppercase tracking-wide text-blue-600">
                Active Session
              </p>

              <p className="mt-1 text-sm font-semibold text-slate-700">
                Not included
              </p>
            </div>
          </div>
        </div>

      </div>
    </PageContainer>
  )
}


async function exportDailyRecordsCSV() {
  try {
    const sessions = await getDailySessions()

    const recordsBySession = await Promise.all(
      sessions.map(async (session) => {
        const transactions = await getTransactionsBySession(
          session.sessionId
        )

        if (transactions.length === 0) {
          return [{ session, transaction: {} }]
        }

        return transactions.map((transaction) => ({
          session,
          transaction,
        }))
      })
    )

    const flattenObject = (object, prefix = '') => {
      const result = {}

      Object.entries(object || {}).forEach(([key, value]) => {
        const column = prefix ? `${prefix}.${key}` : key

        if (
          value !== null &&
          typeof value === 'object' &&
          !Array.isArray(value) &&
          !(value instanceof Date)
        ) {
          Object.assign(
            result,
            flattenObject(value, column)
          )
        } else {
          result[column] =
            value instanceof Date
              ? value.toISOString()
              : Array.isArray(value)
                ? JSON.stringify(value)
                : value ?? ''
        }
      })

      return result
    }

    const rows = recordsBySession
      .flat()
      .map(({ session, transaction }) => ({
        ...flattenObject(session, 'Session'),
        ...flattenObject(transaction, 'Transaction'),
      }))

    if (rows.length === 0) {
      alert('No daily records are available to export.')
      return
    }

    const headers = [
      ...new Set(rows.flatMap((row) => Object.keys(row))),
    ]

    const escapeCSV = (value) => {
      const text = String(value ?? '')
      return `"${text.replace(/"/g, '""')}"`
    }

    const csvContent = [
      headers.map(escapeCSV).join(','),
      ...rows.map((row) =>
        headers
          .map((header) => escapeCSV(row[header]))
          .join(',')
      ),
    ].join('\r\n')

    const blob = new Blob(
      ['\uFEFF', csvContent],
      { type: 'text/csv;charset=utf-8;' }
    )

    const url = URL.createObjectURL(blob)
    const link = document.createElement('a')

    link.href = url
    link.download = `daily-records-${new Date()
      .toISOString()
      .slice(0, 10)}.csv`

    document.body.appendChild(link)
    link.click()
    link.remove()
    URL.revokeObjectURL(url)
  } catch (error) {
    console.error('Daily Records export failed:', error)
    alert('Daily Records export failed. Please check the console.')
  }
}

async function exportDailyRecordsXLSX() {
  try {
    const sessions = await getDailySessions()

    const recordsBySession = await Promise.all(
      sessions.map(async (session) => {
        const transactions =
          await getTransactionsBySession(
            session.sessionId
          )

        if (transactions.length === 0) {
          return [{ session, transaction: {} }]
        }

        return transactions.map((transaction) => ({
          session,
          transaction,
        }))
      })
    )

    const flattenObject = (
      object,
      prefix = ''
    ) => {
      const result = {}

      Object.entries(object || {}).forEach(
        ([key, value]) => {
          const column = prefix
            ? `${prefix}.${key}`
            : key

          if (
            value !== null &&
            typeof value === 'object' &&
            !Array.isArray(value) &&
            !(value instanceof Date)
          ) {
            Object.assign(
              result,
              flattenObject(value, column)
            )
          } else {
            result[column] =
              value instanceof Date
                ? value.toISOString()
                : Array.isArray(value)
                  ? JSON.stringify(value)
                  : value ?? ''
          }
        }
      )

      return result
    }

    const rows = recordsBySession
      .flat()
      .map(({ session, transaction }) => ({
        ...flattenObject(
          session,
          'Session'
        ),
        ...flattenObject(
          transaction,
          'Transaction'
        ),
      }))

    if (rows.length === 0) {
      alert(
        'No daily records are available to export.'
      )
      return
    }

    const worksheet =
      XLSX.utils.json_to_sheet(rows)

    const workbook =
      XLSX.utils.book_new()

    XLSX.utils.book_append_sheet(
      workbook,
      worksheet,
      'Daily Records'
    )

    XLSX.writeFile(
      workbook,
      `daily-records-${new Date()
        .toISOString()
        .slice(0, 10)}.xlsx`
    )
  } catch (error) {
    console.error(
      'Daily Records XLSX export failed:',
      error
    )

    alert(
      'Daily Records Excel export failed. Please check the console.'
    )
  }
}

async function exportDailyRecordsPDF() {
  try {
    const sessions = await getDailySessions()

    const recordsBySession = await Promise.all(
      sessions.map(async (session) => {
        const transactions =
          await getTransactionsBySession(
            session.sessionId
          )

        if (transactions.length === 0) {
          return [{ session, transaction: {} }]
        }

        return transactions.map((transaction) => ({
          session,
          transaction,
        }))
      })
    )

    const rows = recordsBySession
      .flat()
      .map(({ session, transaction }) => ({
        Date: session.date || '',
        Session:
          session.sessionId || '',
        'Opening Cash':
          session.openingCash ?? '',
        'Gross Sales':
          session.grossSales ?? '',
        Refunds:
          session.totalRefunds ?? '',
        'Net Sales':
          session.netSales ?? '',
        'Expected Cash':
          session.expectedClosingCash ?? '',
        'Actual Cash':
          session.actualCash ?? '',
        Status:
          session.status || '',
        Transaction:
          transaction.type || '',
        Category:
          transaction.category || '',
        Amount:
          transaction.amount ?? '',
        Reason:
          transaction.reason || '',
        Customer:
          transaction.customerRef || '',
      }))

    if (rows.length === 0) {
      alert(
        'No daily records are available to export.'
      )
      return
    }

    const doc = new jsPDF({
      orientation: 'landscape',
      unit: 'mm',
      format: 'a4',
    })

    doc.setFontSize(16)
    doc.text(
      'Daily Records',
      14,
      15
    )

    doc.setFontSize(9)
    doc.text(
      `Generated: ${new Date().toLocaleString()}`,
      14,
      21
    )

    const columns = [
      'Date',
      'Session',
      'Opening Cash',
      'Gross Sales',
      'Refunds',
      'Net Sales',
      'Expected Cash',
      'Actual Cash',
      'Status',
      'Transaction',
      'Category',
      'Amount',
      'Reason',
      'Customer',
    ]

    const body = rows.map((row) =>
      columns.map(
        (column) =>
          String(row[column] ?? '')
      )
    )

    autoTable(doc, {
      head: [columns],
      body,
      startY: 27,
      styles: {
        fontSize: 6,
        cellPadding: 2,
        overflow: 'linebreak',
      },
      headStyles: {
        fontSize: 6,
        fontStyle: 'bold',
      },
      margin: {
        left: 8,
        right: 8,
      },
    })

    doc.save(
      `daily-records-${new Date()
        .toISOString()
        .slice(0, 10)}.pdf`
    )
  } catch (error) {
    console.error(
      'Daily Records PDF export failed:',
      error
    )

    alert(
      'Daily Records PDF export failed. Please check the console.'
    )
  }
}

async function exportMonthlyReportsCSV() {
  try {
    const sessions = await getDailySessions()

    if (!sessions || sessions.length === 0) {
      alert('No monthly records are available to export.')
      return
    }

    const monthlyMap = {}

    sessions.forEach((record) => {
      const month = String(record.date || '').slice(0, 7)

      if (!month) return

      if (!monthlyMap[month]) {
        monthlyMap[month] = {
          month,
          workingDays: 0,
          grossSales: 0,
          totalRefunds: 0,
          netSales: 0,
          openingCash: 0,
          expectedClosingCash: 0,
          actualCash: 0,
          medicine: 0,
          general: 0,
          dispensing: 0,
          matchedDays: 0,
          shortageDays: 0,
          extraCashDays: 0,
          totalShortage: 0,
          totalExtraCash: 0,
          overallDifference: 0,
        }
      }

      const monthly = monthlyMap[month]

      monthly.workingDays += 1

      monthly.grossSales += Number(
        record.grossSales || 0
      )

      monthly.totalRefunds += Number(
        record.totalRefunds || 0
      )

      monthly.netSales += Number(
        record.netSales || 0
      )

      monthly.openingCash += Number(
        record.openingCash || 0
      )

      monthly.expectedClosingCash += Number(
        record.expectedClosingCash || 0
      )

      monthly.actualCash += Number(
        record.actualCash || 0
      )

      monthly.medicine += Number(
        record.sales?.medicine || 0
      )

      monthly.general += Number(
        record.sales?.general || 0
      )

      monthly.dispensing += Number(
        record.sales?.dispensing || 0
      )

      if (record.status === 'Matched') {
        monthly.matchedDays += 1
      }

      if (record.status === 'Shortage') {
        monthly.shortageDays += 1
      }

      if (record.status === 'Extra Cash') {
        monthly.extraCashDays += 1
      }

      const difference = Number(
        record.difference || 0
      )

      monthly.overallDifference += difference

      if (difference < 0) {
        monthly.totalShortage += Math.abs(
          difference
        )
      }

      if (difference > 0) {
        monthly.totalExtraCash += difference
      }
    })

    const rows = Object.values(monthlyMap)
      .sort((a, b) =>
        a.month.localeCompare(b.month)
      )
      .map((monthly) => ({
        Month: monthly.month,
        'Working Days': monthly.workingDays,
        'Gross Sales': monthly.grossSales,
        'Total Refunds': monthly.totalRefunds,
        'Net Sales': monthly.netSales,
        'Opening Cash': monthly.openingCash,
        'Expected Cash': monthly.expectedClosingCash,
        'Actual Cash': monthly.actualCash,
        'Medicine Sales': monthly.medicine,
        'General Items Sales': monthly.general,
        'Dispensing Sales': monthly.dispensing,
        'Matched Days': monthly.matchedDays,
        'Shortage Days': monthly.shortageDays,
        'Extra Cash Days': monthly.extraCashDays,
        'Total Shortage': monthly.totalShortage,
        'Total Extra Cash': monthly.totalExtraCash,
        'Overall Difference': monthly.overallDifference,
      }))

    if (rows.length === 0) {
      alert('No monthly records are available to export.')
      return
    }

    const headers = Object.keys(rows[0])

    const escapeCSV = (value) => {
      const text = String(value ?? '')
      return `"${text.replace(/"/g, '""')}"`
    }

    const csvContent = [
      headers.map(escapeCSV).join(','),
      ...rows.map((row) =>
        headers
          .map((header) =>
            escapeCSV(row[header])
          )
          .join(',')
      ),
    ].join('\r\n')

    const blob = new Blob(
      ['\uFEFF', csvContent],
      {
        type: 'text/csv;charset=utf-8;',
      }
    )

    const url = URL.createObjectURL(blob)
    const link = document.createElement('a')

    link.href = url

    link.download = `monthly-reports-${new Date()
      .toISOString()
      .slice(0, 10)}.csv`

    document.body.appendChild(link)
    link.click()
    link.remove()

    URL.revokeObjectURL(url)
  } catch (error) {
    console.error(
      'Monthly Reports export failed:',
      error
    )

    alert(
      'Monthly Reports export failed. Please check the console.'
    )
  }
}

async function exportMonthlyReportsXLSX() {
  try {
    const sessions = await getDailySessions()

    if (!sessions || sessions.length === 0) {
      alert('No monthly records are available to export.')
      return
    }

    const monthlyMap = {}

    sessions.forEach((record) => {
      const month = String(record.date || '').slice(0, 7)

      if (!month) return

      if (!monthlyMap[month]) {
        monthlyMap[month] = {
          month,
          workingDays: 0,
          grossSales: 0,
          totalRefunds: 0,
          netSales: 0,
          openingCash: 0,
          expectedClosingCash: 0,
          actualCash: 0,
          medicine: 0,
          general: 0,
          dispensing: 0,
          matchedDays: 0,
          shortageDays: 0,
          extraCashDays: 0,
          totalShortage: 0,
          totalExtraCash: 0,
          overallDifference: 0,
        }
      }

      const monthly = monthlyMap[month]

      monthly.workingDays += 1

      monthly.grossSales += Number(
        record.grossSales || 0
      )

      monthly.totalRefunds += Number(
        record.totalRefunds || 0
      )

      monthly.netSales += Number(
        record.netSales || 0
      )

      monthly.openingCash += Number(
        record.openingCash || 0
      )

      monthly.expectedClosingCash += Number(
        record.expectedClosingCash || 0
      )

      monthly.actualCash += Number(
        record.actualCash || 0
      )

      monthly.medicine += Number(
        record.sales?.medicine || 0
      )

      monthly.general += Number(
        record.sales?.general || 0
      )

      monthly.dispensing += Number(
        record.sales?.dispensing || 0
      )

      if (record.status === 'Matched') {
        monthly.matchedDays += 1
      }

      if (record.status === 'Shortage') {
        monthly.shortageDays += 1
      }

      if (record.status === 'Extra Cash') {
        monthly.extraCashDays += 1
      }

      const difference = Number(
        record.difference || 0
      )

      monthly.overallDifference += difference

      if (difference < 0) {
        monthly.totalShortage += Math.abs(
          difference
        )
      }

      if (difference > 0) {
        monthly.totalExtraCash += difference
      }
    })

    const rows = Object.values(monthlyMap)
      .sort((a, b) =>
        a.month.localeCompare(b.month)
      )
      .map((monthly) => ({
        Month: monthly.month,
        'Working Days': monthly.workingDays,
        'Gross Sales': monthly.grossSales,
        'Total Refunds': monthly.totalRefunds,
        'Net Sales': monthly.netSales,
        'Opening Cash': monthly.openingCash,
        'Expected Cash':
          monthly.expectedClosingCash,
        'Actual Cash': monthly.actualCash,
        'Medicine Sales': monthly.medicine,
        'General Items Sales':
          monthly.general,
        'Dispensing Sales':
          monthly.dispensing,
        'Matched Days':
          monthly.matchedDays,
        'Shortage Days':
          monthly.shortageDays,
        'Extra Cash Days':
          monthly.extraCashDays,
        'Total Shortage':
          monthly.totalShortage,
        'Total Extra Cash':
          monthly.totalExtraCash,
        'Overall Difference':
          monthly.overallDifference,
      }))

    if (rows.length === 0) {
      alert('No monthly records are available to export.')
      return
    }

    const worksheet =
      XLSX.utils.json_to_sheet(rows)

    const workbook =
      XLSX.utils.book_new()

    XLSX.utils.book_append_sheet(
      workbook,
      worksheet,
      'Monthly Reports'
    )

    XLSX.writeFile(
      workbook,
      `monthly-reports-${new Date()
        .toISOString()
        .slice(0, 10)}.xlsx`
    )
  } catch (error) {
    console.error(
      'Monthly Reports XLSX export failed:',
      error
    )

    alert(
      'Monthly Reports Excel export failed. Please check the console.'
    )
  }
}

async function exportMonthlyReportsPDF() {
  try {
    const sessions = await getDailySessions()

    const rows = sessions.map((session) => ({
      Date: session.date || '',
      Session: session.sessionId || '',
      'Opening Cash': session.openingCash ?? '',
      'Gross Sales': session.grossSales ?? '',
      Refunds: session.totalRefunds ?? '',
      'Net Sales': session.netSales ?? '',
      'Expected Cash': session.expectedClosingCash ?? '',
      'Actual Cash': session.actualCash ?? '',
      Status: session.status || '',
    }))

    if (rows.length === 0) {
      alert(
        'No monthly report data is available to export.'
      )
      return
    }

    const doc = new jsPDF({
      orientation: 'landscape',
      unit: 'mm',
      format: 'a4',
    })

    doc.setFontSize(16)
    doc.text(
      'Monthly Reports',
      14,
      15
    )

    doc.setFontSize(9)
    doc.text(
      `Generated: ${new Date().toLocaleString()}`,
      14,
      21
    )

    const columns = [
      'Date',
      'Session',
      'Opening Cash',
      'Gross Sales',
      'Refunds',
      'Net Sales',
      'Expected Cash',
      'Actual Cash',
      'Status',
    ]

    const body = rows.map((row) =>
      columns.map(
        (column) =>
          String(row[column] ?? '')
      )
    )

    autoTable(doc, {
      head: [columns],
      body,
      startY: 27,
      styles: {
        fontSize: 7,
        cellPadding: 2,
        overflow: 'linebreak',
      },
      headStyles: {
        fontSize: 7,
        fontStyle: 'bold',
      },
      margin: {
        left: 10,
        right: 10,
      },
    })

    doc.save(
      `monthly-reports-${new Date()
        .toISOString()
        .slice(0, 10)}.pdf`
    )
  } catch (error) {
    console.error(
      'Monthly Reports PDF export failed:',
      error
    )

    alert(
      'Monthly Reports PDF export failed. Please check the console.'
    )
  }
}

async function exportAuditLogsCSV() {
  try {
    const logs = await getAuditLogs()

    if (!logs || logs.length === 0) {
      alert('No audit logs are available to export.')
      return
    }

    const flattenObject = (object, prefix = '') => {
      const result = {}

      Object.entries(object || {}).forEach(([key, value]) => {
        const column = prefix
          ? `${prefix}.${key}`
          : key

        if (
          value !== null &&
          typeof value === 'object' &&
          !Array.isArray(value) &&
          !(value instanceof Date)
        ) {
          Object.assign(
            result,
            flattenObject(value, column)
          )
        } else {
          result[column] =
            value instanceof Date
              ? value.toISOString()
              : Array.isArray(value)
                ? JSON.stringify(value)
                : value ?? ''
        }
      })

      return result
    }

    const rows = logs.map((log) =>
      flattenObject(log)
    )

    const headers = [
      ...new Set(
        rows.flatMap((row) =>
          Object.keys(row)
        )
      ),
    ]

    const escapeCSV = (value) => {
      const text = String(value ?? '')
      return `"${text.replace(/"/g, '""')}"`
    }

    const csvContent = [
      headers
        .map(escapeCSV)
        .join(','),

      ...rows.map((row) =>
        headers
          .map((header) =>
            escapeCSV(row[header])
          )
          .join(',')
      ),
    ].join('\r\n')

    const blob = new Blob(
      ['\uFEFF', csvContent],
      {
        type: 'text/csv;charset=utf-8;',
      }
    )

    const url =
      URL.createObjectURL(blob)

    const link =
      document.createElement('a')

    link.href = url

    link.download =
      `audit-logs-${new Date()
        .toISOString()
        .slice(0, 10)}.csv`

    document.body.appendChild(link)

    link.click()

    link.remove()

    URL.revokeObjectURL(url)
  } catch (error) {
    console.error(
      'Audit Logs export failed:',
      error
    )

    alert(
      'Audit Logs export failed. Please check the console.'
    )
  }
}

async function exportAuditLogsXLSX() {
  try {
    const logs = await getAuditLogs()

    if (!logs || logs.length === 0) {
      alert('No audit logs are available to export.')
      return
    }

    const flattenObject = (
      object,
      prefix = ''
    ) => {
      const result = {}

      Object.entries(object || {}).forEach(
        ([key, value]) => {
          const column = prefix
            ? `${prefix}.${key}`
            : key

          if (
            value !== null &&
            typeof value === 'object' &&
            !Array.isArray(value) &&
            !(value instanceof Date)
          ) {
            Object.assign(
              result,
              flattenObject(value, column)
            )
          } else {
            result[column] =
              value instanceof Date
                ? value.toISOString()
                : Array.isArray(value)
                  ? JSON.stringify(value)
                  : value ?? ''
          }
        }
      )

      return result
    }

    const rows = logs.map((log) =>
      flattenObject(log)
    )

    const headers = [
      ...new Set(
        rows.flatMap((row) =>
          Object.keys(row)
        )
      ),
    ]

    const worksheet =
      XLSX.utils.json_to_sheet(rows)

    const workbook =
      XLSX.utils.book_new()

    XLSX.utils.book_append_sheet(
      workbook,
      worksheet,
      'Audit Logs'
    )

    XLSX.writeFile(
      workbook,
      `audit-logs-${new Date()
        .toISOString()
        .slice(0, 10)}.xlsx`
    )
  } catch (error) {
    console.error(
      'Audit Logs XLSX export failed:',
      error
    )

    alert(
      'Audit Logs Excel export failed. Please check the console.'
    )
  }
}

async function exportAuditLogsPDF() {
  try {
    const logs = await getAuditLogs()

    const rows = logs.map((log) => ({
      Date:
        log.timestamp
          ? new Date(log.timestamp).toLocaleString()
          : '',
      Action: log.action || '',
      User:
        log.userName ||
        log.username ||
        log.userId ||
        '',
      Details:
        log.details ||
        log.description ||
        '',
      Customer:
        log.customerRef ||
        '',
    }))

    if (rows.length === 0) {
      alert(
        'No audit logs are available to export.'
      )
      return
    }

    const doc = new jsPDF({
      orientation: 'landscape',
      unit: 'mm',
      format: 'a4',
    })

    doc.setFontSize(16)
    doc.text(
      'Audit Logs',
      14,
      15
    )

    doc.setFontSize(9)
    doc.text(
      `Generated: ${new Date().toLocaleString()}`,
      14,
      21
    )

    const columns = [
      'Date',
      'Action',
      'User',
      'Details',
      'Customer',
    ]

    const body = rows.map((row) =>
      columns.map(
        (column) =>
          String(row[column] ?? '')
      )
    )

    autoTable(doc, {
      head: [columns],
      body,
      startY: 27,
      styles: {
        fontSize: 7,
        cellPadding: 2,
        overflow: 'linebreak',
      },
      headStyles: {
        fontSize: 7,
        fontStyle: 'bold',
      },
      margin: {
        left: 10,
        right: 10,
      },
    })

    doc.save(
      `audit-logs-${new Date()
        .toISOString()
        .slice(0, 10)}.pdf`
    )
  } catch (error) {
    console.error(
      'Audit Logs PDF export failed:',
      error
    )

    alert(
      'Audit Logs PDF export failed. Please check the console.'
    )
  }
}

function DataExport() {
  const [isExporting, setIsExporting] = useState(false)
  const handleDailyPDFExport = async () => {
    setIsExporting(true)

    try {
      await exportDailyRecordsPDF()
    } finally {
      setIsExporting(false)
    }
  }

  const handleDailyExcelExport = async () => {
    setIsExporting(true)

    try {
      await exportDailyRecordsXLSX()
    } finally {
      setIsExporting(false)
    }
  }
  const handleMonthlyPDFExport = async () => {
    setIsExporting(true)

    try {
      await exportMonthlyReportsPDF()
    } finally {
      setIsExporting(false)
    }
  }
  const handleMonthlyExcelExport = async () => {
    setIsExporting(true)

    try {
      await exportMonthlyReportsXLSX()
    } finally {
      setIsExporting(false)
    }
  }
  const handleAuditPDFExport = async () => {
    setIsExporting(true)

    try {
      await exportAuditLogsPDF()
    } finally {
      setIsExporting(false)
    }
  }
  const handleAuditExcelExport = async () => {
    setIsExporting(true)

    try {
      await exportAuditLogsXLSX()
    } finally {
      setIsExporting(false)
    }
  }
  return (
    <PageContainer>
      <div className="space-y-5">

        {/* Header */}
        <div>
          <div className="flex items-center gap-3">
            <div className="flex h-11 w-11 items-center justify-center rounded-xl bg-blue-50 text-blue-600">
              <svg
                width="22"
                height="22"
                viewBox="0 0 24 24"
                fill="none"
                stroke="currentColor"
                strokeWidth="2"
                strokeLinecap="round"
                strokeLinejoin="round"
              >
                <path d="M12 3v12" />
                <path d="m7 10 5 5 5-5" />
                <path d="M5 21h14" />
              </svg>
            </div>

            <div>
              <h1 className="text-xl font-bold tracking-tight text-slate-900 sm:text-2xl">
                Data Export
              </h1>

              <p className="mt-1 text-sm text-slate-500">
                Export your medical store records for reporting and offline use.
              </p>
            </div>
          </div>
        </div>

        {/* Export Center */}
        <div className={`${ui.card} p-5 sm:p-6`}>
          <div className="flex items-start gap-4">
            <div className="flex h-12 w-12 shrink-0 items-center justify-center rounded-xl bg-blue-50 text-xl">
              📤
            </div>

            <div>
              <h2 className="text-lg font-bold text-slate-900">
                Export Center
              </h2>

              <p className="mt-1 max-w-2xl text-sm leading-6 text-slate-500">
                Choose the type of records you want to export and select
                your preferred file format.
              </p>
            </div>
          </div>

          {/* Export Options */}
          <div className="mt-6 grid grid-cols-1 gap-4 md:grid-cols-2 lg:grid-cols-3">

            {/* Daily Records */}
            <div className="rounded-2xl border border-slate-200 bg-white p-5">
              <div className="flex h-11 w-11 items-center justify-center rounded-xl bg-blue-50 text-xl">
                📅
              </div>

              <h3 className="mt-4 text-base font-bold text-slate-900">
                Daily Records
              </h3>

              <p className="mt-1 text-sm leading-6 text-slate-500">
                Export daily sessions, sales, refunds, and closing records.
              </p>

              <div className="mt-5 flex flex-wrap gap-2">
                <button
                  type="button"
                  onClick={handleDailyPDFExport}
                  className={`rounded-xl bg-blue-600 px-4 py-2 text-sm font-semibold text-white transition hover:bg-blue-700 cursor-pointer ${isExporting
                    ? 'animate-[processingPulse_1.2s_ease-in-out_infinite]'
                    : ''
                    }`}
                >
                  PDF
                </button>


                <button
                  type="button"
                  onClick={handleDailyExcelExport}
                  className={`rounded-xl bg-blue-600 px-4 py-2 text-sm font-semibold text-white transition hover:bg-blue-700 cursor-pointer ${isExporting
                    ? 'animate-[processingPulse_1.2s_ease-in-out_infinite]'
                    : ''
                    }`}
                >
                  Excel
                </button>
              </div>
            </div>

            {/* Monthly Reports */}
            <div className="rounded-2xl border border-slate-200 bg-white p-5">
              <div className="flex h-11 w-11 items-center justify-center rounded-xl bg-blue-50 text-xl">
                📊
              </div>

              <h3 className="mt-4 text-base font-bold text-slate-900">
                Monthly Reports
              </h3>

              <p className="mt-1 text-sm leading-6 text-slate-500">
                Export monthly sales, refunds, category totals, and summaries.
              </p>

              <div className="mt-5 flex flex-wrap gap-2">
                <button
                  type="button"
                  onClick={handleMonthlyPDFExport}
                  disabled={isExporting}
                  className={`rounded-xl bg-blue-600 px-4 py-2 text-sm font-semibold text-white transition hover:bg-blue-700 cursor-pointer ${isExporting
                    ? 'animate-[processingPulse_1.2s_ease-in-out_infinite]'
                    : ''
                    }`}
                >
                  PDF
                </button>

                <button
                  type="button"
                  onClick={handleMonthlyExcelExport}
                  disabled={isExporting}
                  className={`rounded-xl bg-blue-600 px-4 py-2 text-sm font-semibold text-white transition hover:bg-blue-700 cursor-pointer ${isExporting
                    ? 'animate-[processingPulse_1.2s_ease-in-out_infinite]'
                    : ''
                    }`}
                >
                  Excel
                </button>
              </div>
            </div>

            {/* Audit Logs */}
            <div className="rounded-2xl border border-slate-200 bg-white p-5">
              <div className="flex h-11 w-11 items-center justify-center rounded-xl bg-blue-50 text-xl">
                🛡️
              </div>

              <h3 className="mt-4 text-base font-bold text-slate-900">
                Audit Logs
              </h3>

              <p className="mt-1 text-sm leading-6 text-slate-500">
                Export user activity and system audit records.
              </p>

              <div className="mt-5 flex flex-wrap gap-2">
                <button
                  type="button"
                  onClick={handleAuditPDFExport}
                  disabled={isExporting}
                  className={`rounded-xl bg-blue-600 px-4 py-2 text-sm font-semibold text-white transition hover:bg-blue-700 cursor-pointer ${isExporting
                    ? 'animate-[processingPulse_1.2s_ease-in-out_infinite]'
                    : ''
                    }`}
                >
                  PDF
                </button>

                <button
                  type="button"
                  onClick={handleAuditExcelExport}
                  disabled={isExporting}
                  className={`rounded-xl bg-blue-600 px-4 py-2 text-sm font-semibold text-white transition hover:bg-blue-700 cursor-pointer ${isExporting
                      ? 'animate-[processingPulse_1.2s_ease-in-out_infinite]'
                      : ''
                    }`}
                >
                  Excel
                </button>
              </div>
            </div>

          </div>
        </div>

        {/* Information Note */}
        <div className="overflow-hidden rounded-2xl border border-blue-100 bg-blue-50/50">
          <div className="flex items-start gap-4 p-5 sm:p-6">
            <div className="flex h-11 w-11 shrink-0 items-center justify-center rounded-xl bg-blue-100 text-lg">
              ℹ️
            </div>

            <div>
              <h3 className="text-sm font-bold text-slate-900">
                Export Information
              </h3>

              <p className="mt-1 text-sm leading-6 text-slate-600">
                Exports will be generated locally from your current
                application data. No data will be uploaded to an external
                server.
              </p>
            </div>
          </div>
        </div>

      </div>
    </PageContainer>
  )
}

/* --------------------------------
   App Shell / Navigation
-------------------------------- */

function PageContainer({ children }) {
  return (
    <div className="max-w-7xl mx-auto">
      {children}
    </div>
  )
}

function NavIcon({ type }) {
  const common = {
    width: 19,
    height: 19,
    viewBox: '0 0 24 24',
    fill: 'none',
    stroke: 'currentColor',
    strokeWidth: 2,
    strokeLinecap: 'round',
    strokeLinejoin: 'round',
  }

  if (type === 'dashboard') {
    return (
      <svg {...common}>
        <rect
          x="3"
          y="3"
          width="7"
          height="7"
          rx="1"
        />
        <rect
          x="14"
          y="3"
          width="7"
          height="7"
          rx="1"
        />
        <rect
          x="3"
          y="14"
          width="7"
          height="7"
          rx="1"
        />
        <rect
          x="14"
          y="14"
          width="7"
          height="7"
          rx="1"
        />
      </svg>
    )
  }

  if (type === 'records') {
    return (
      <svg {...common}>
        <path d="M6 3h12v18H6z" />
        <path d="M9 7h6M9 11h6M9 15h4" />
      </svg>
    )
  }

  if (type === 'monthly') {
    return (
      <svg {...common}>
        <path d="M4 19V5" />
        <path d="M4 19h16" />
        <path d="m7 15 3-4 3 2 5-6" />
      </svg>
    )
  }

  if (type === 'audit') {
    return (
      <svg {...common}>
        <path d="M3 12a9 9 0 1 0 3-6.7" />
        <path d="M3 4v5h5" />
        <path d="M12 7v5l3 2" />
      </svg>
    )
  }
  if (type === 'users') {
    return (
      <svg {...common}>
        <path d="M16 21v-2a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v2" />
        <circle cx="9" cy="7" r="4" />
        <path d="M22 21v-2a4 4 0 0 0-3-3.87" />
        <path d="M16 3.13a4 4 0 0 1 0 7.75" />
      </svg>
    )
  }

  if (type === 'backup') {
    return (
      <svg {...common}>
        <path d="M12 3v12" />
        <path d="m7 10 5 5 5-5" />
        <path d="M5 21h14" />
        <path d="M5 17v4" />
        <path d="M19 17v4" />
      </svg>
    )
  }

  if (type === 'export') {
    return (
      <svg {...common}>
        <path d="M12 3v12" />
        <path d="m7 10 5 5 5-5" />
        <path d="M5 21h14" />
      </svg>
    )
  }

  return (
    <svg {...common}>
      <path d="M3 10h18" />
      <path d="M5 10v10h14V10" />
      <path d="M4 10V6h16v4" />
      <path d="M8 6V3h8v3" />
    </svg>
  )
}

function AppShell({
  currentView,
  onNavigate,
  mobileMenuOpen,
  setMobileMenuOpen,
  isDayOpen,
  sessionDate,
  currentUser,
  onLogout,
  isAdmin,
  children,
}) {
  const menuItems = [
    {
      id: 'dashboard',
      label: isDayOpen
        ? 'Daily Dashboard'
        : 'Start Daily',
      icon: isDayOpen
        ? 'dashboard'
        : 'start',
    },
    {
      id: 'records',
      label: 'Daily Records',
      icon: 'records',
    },
    ...(isAdmin
      ? [
        {
          id: 'auditLogs',
          label: 'Audit Logs',
          icon: 'audit',
        },
        {
          id: 'monthly',
          label: 'Monthly Reports',
          icon: 'monthly',
        },
        {
          id: 'userManagement',
          label: 'User Management',
          icon: 'users',
        },
        {
          id: 'backupRestore',
          label: 'Backup & Restore',
          icon: 'backup',
        },
        {
          id: 'dataExport',
          label: 'Data Export',
          icon: 'export',
        },
      ]
      : []),
  ]

  const pageTitle =
    currentView === 'dashboard'
      ? isDayOpen
        ? 'Daily Dashboard'
        : 'Start Daily'
      : currentView === 'records'
        ? 'Daily Records'
        : currentView === 'recordDetails'
          ? 'Daily Record Details'
          : currentView === 'auditLogs'
            ? 'Audit Logs'
            : currentView === 'monthly'
              ? 'Monthly Reports'
              : currentView === 'userManagement'
                ? 'User Management'
                : currentView === 'backupRestore'
                  ? 'Backup & Restore'
                  : currentView === 'dataExport'
                    ? 'Data Export'
                    : currentView === 'refund'
                      ? 'Refund Management'
                      : ''

  const handleNavigate = (view) => {
    onNavigate(view)
    setMobileMenuOpen(false)
  }

  const sidebar = (
    <div className="h-full min-h-0 flex flex-col overflow-hidden bg-slate-950 border-r border-slate-800 shadow-[8px_0_30px_rgba(15,23,42,0.18)]">
      <div className="p-5 border-b border-slate-800 bg-gradient-to-br from-slate-950 via-slate-950 to-slate-900">
        <div className="flex items-center gap-3">

          <div className="relative h-16 w-16 shrink-0 flex items-center justify-center">
            <div className="absolute inset-[-5px] rounded-full border border-blue-400/30 border-t-blue-400 border-r-cyan-300 animate-spin [animation-duration:3.5s]" />

            <div className="absolute inset-[-2px] rounded-full border border-cyan-400/20 border-b-cyan-300 animate-spin [animation-duration:2.5s] [animation-direction:reverse]" />

            <img
              src="/km-store-logo.png"
              alt="KM Store logo"
              className="relative z-10 h-16 w-16 rounded-full object-contain"
            />
          </div>

          <div className="min-w-0">
            <h1 className="font-bold text-white -900 leading-tight">
              Khalid Medical
            </h1>

            <p className="text-[10px] text-slate-400 mt-0.5 whitespace-nowrap">
              POS & Cash Management
            </p>
          </div>
        </div>
      </div>

      <nav className="min-h-0 flex-1 overflow-y-auto p-3 space-y-1 scrollbar-thin scrollbar-thumb-slate-700 scrollbar-track-transparent">
        <p className="px-3 pt-2 pb-2 text-[11px] font-bold uppercase tracking-wider text-slate-500">
          Menu
        </p>

        {menuItems.map((item) => {
          const active =
            currentView === item.id

          return (
            <button
              key={item.id}
              onClick={() =>
                handleNavigate(
                  item.id
                )
              }
              className={`w-full flex items-center gap-3 px-3 py-3 rounded-xl text-sm font-semibold text-left transition focus:outline-none focus:ring-2 focus:ring-blue-100 ${active
                ? 'bg-blue-600/15 text-blue-300 shadow-sm ring-1 ring-blue-500/20'
                : 'text-slate-300 hover:bg-slate-900 hover:text-white'
                }`}
            >
              <NavIcon
                type={item.icon}
              />

              <span>
                {item.label}
              </span>
            </button>
          )
        })}
      </nav>

      <div className="shrink-0 p-4 border-t border-slate-800">
        <div className="rounded-2xl border border-slate-800 bg-gradient-to-br from-slate-900 to-slate-800 p-3 mb-3 shadow-sm">
          <div className="flex items-center justify-between gap-2">
            <div className="min-w-0">
              <p className="text-[11px] font-semibold uppercase tracking-wide text-slate-400">
                Logged In
              </p>

              <p className="text-sm font-bold text-white mt-1 truncate">
                {currentUser?.name}
              </p>
            </div>

            <span
              className={`shrink-0 px-2 py-1 rounded-full text-[10px] font-bold ${isAdmin
                ? 'bg-blue-100 text-blue-700'
                : 'bg-slate-200 text-slate-400'
                }`}
            >
              {isAdmin
                ? 'ADMIN'
                : 'USER'}
            </span>
          </div>

          <p className="text-xs text-slate-400 mt-1">
            @{currentUser?.username}
          </p>
        </div>

        {isDayOpen ? (
          <div className="rounded-2xl border border-emerald-500/20 bg-gradient-to-br from-emerald-500/10 to-slate-900 p-3 shadow-sm">
            <p className="text-[11px] font-semibold uppercase tracking-wide text-emerald-300">
              Current Day
            </p>

            <p className="text-sm font-bold text-slate-400 mt-1">
              {formatDate(
                sessionDate
              )}
            </p>

            <p className="text-xs text-slate-500 mt-1">
              Session is open
            </p>
          </div>
        ) : (
          <div className="rounded-xl bg-slate-900 p-3" border-slate-800 shadow-sm>
            <p className="text-xs font-medium text-white">
              No active day
            </p>

            <p className="text-xs text-slate-400 mt-1">
              Start a session to begin today's work.
            </p>
          </div>
        )}

        <button
          onClick={onLogout}
          className="w-full mt-4 flex items-center justify-center gap-2 rounded-xl border border-slate-700 bg-slate-900 px-3 py-2.5 text-sm font-semibold text-slate-300 shadow-sm transition-all hover:-translate-y-0.5 hover:border-blue-200  hover:bg-slate-800 hover:text-white hover:shadow-md focus:outline-none focus:ring-4 focus:ring-blue-100"
        >
          <svg
            width="17"
            height="17"
            viewBox="0 0 24 24"
            fill="none"
            stroke="currentColor"
            strokeWidth="2"
            strokeLinecap="round"
            strokeLinejoin="round"
          >
            <path d="M10 17l5-5-5-5" />
            <path d="M15 12H3" />
            <path d="M21 19V5a2 2 0 0 0-2-2h-6" />
          </svg>

          Logout
        </button>
      </div>
    </div>
  )

  return (
    <div className="min-h-screen bg-[radial-gradient(circle_at_top_right,_rgba(37,99,235,0.14),_transparent_30%),radial-gradient(circle_at_bottom_left,_rgba(15,23,42,0.06),_transparent_28%),linear-gradient(180deg,#ffffff_0%,#f1f5f9_52%,#e8eef7_100%)]">

      <style>
        {`
    @keyframes pageEnter {
      from {
        opacity: 0;
        transform: translateY(8px);
      }

      to {
        opacity: 1;
        transform: translateY(0);
      }
    }

    @keyframes summaryCardEnter {
      from {
        opacity: 0;
        transform: translateY(4px);
      }

      to {
        opacity: 1;
        transform: translateY(0);
      }
    }
      @keyframes notificationEnter {
  from {
    opacity: 0;
    transform: translateX(10px);
  }

  to {
    opacity: 1;
    transform: translateX(0);
  }
}
  @keyframes closingResultEnter {
  from {
    opacity: 0;
    transform: scale(0.98);
  }

  to {
    opacity: 1;
    transform: scale(1);
  }
}
  @keyframes processingPulse {
  0%,
  100% {
    opacity: 0.55;
    transform: scale(0.98);
  }

  50% {
    opacity: 1;
    transform: scale(1);
  }
}
  `}
      </style>

      <div className="hidden lg:block fixed inset-y-0 left-0 w-64 z-40">
        {sidebar}
      </div>

      {mobileMenuOpen && (
        <div className="fixed inset-0 z-50 lg:hidden">
          <button
            aria-label="Close menu"
            className="absolute inset-0 bg-black/30"
            onClick={() =>
              setMobileMenuOpen(
                false
              )
            }
          />

          <div className="relative w-72 max-w-[85vw] h-full shadow-xl">
            {sidebar}
          </div>
        </div>
      )}

      <div className="lg:pl-64 min-h-screen">
        <header className="sticky top-0 z-40 bg-white/90 backdrop-blur-xl border-b border-slate-200/80 shadow-[0_4px_20px_rgba(15,23,42,0.06)]">
          <div className="px-4 sm:px-6 py-3.5 flex items-center justify-between gap-4">
            <div className="flex items-center gap-3 min-w-0">
              {/* Mobile Menu Button */}
              <button
                onClick={() =>
                  setMobileMenuOpen(
                    true
                  )
                }
                className="relative flex h-10 w-10 shrink-0 items-center justify-center focus:outline-none focus:ring-4 focus:ring-blue-100 lg:hidden"
                aria-label="Open menu"
              >

                <div className="relative h-10 w-10 shrink-0 flex items-center justify-center">
                  <div className="absolute inset-[-4px] rounded-full border border-blue-400/30 border-t-blue-400 border-r-cyan-300 animate-spin [animation-duration:3.5s]" />

                  <div className="absolute inset-[-2px] rounded-full border border-cyan-400/20 border-b-cyan-300 animate-spin [animation-duration:2.5s] [animation-direction:reverse]" />

                  <img
                    src="/km-store-logo.png"
                    alt="KM Store logo"
                    className="relative z-10 h-10 w-10 rounded-full object-contain"
                  />
                </div>
              </button>

              {/* Mobile App Name */}
              <div className="lg:hidden min-w-0">
                <h1 className="truncate text-sm font-bold tracking-tight text-slate-800">
                  Khalid Medical POS
                </h1>
              </div>

              {/* Desktop Page Title */}
              <div className="hidden lg:block min-w-0">
                <h1 className="text-base sm:text-lg font-bold text-blue-600 truncate">
                  {pageTitle}
                </h1>
              </div>
            </div>

            <div className="flex items-center gap-2 sm:gap-3">

              {/* Mobile Navbar Info */}
              <div className="flex sm:hidden items-center gap-2 min-w-0">

                {/* User / Role */}
                <div className="min-w-0 text-right">
                  <p className="max-w-[78px] truncate text-[11px] font-bold text-slate-800">
                    {currentUser?.name}
                  </p>

                  <p className="text-[9px] font-medium text-slate-500">
                    {isAdmin
                      ? 'Administrator'
                      : 'Staff User'}
                  </p>
                </div>

                {/* Current Session */}
                {isDayOpen &&
                  sessionDate && (
                    <div className="shrink-0 border-l border-slate-200 pl-2">
                      <div className="flex items-center justify-end gap-1">
                        <span className="h-1.5 w-1.5 rounded-full bg-green-500" />

                        <p className="text-[9px] font-bold uppercase tracking-wide text-green-600">
                          Session
                        </p>
                      </div>

                      <p className="mt-0.5 text-[10px] font-semibold text-slate-700">
                        {formatDate(sessionDate)}
                      </p>
                    </div>
                  )}
              </div>

              {/* Desktop Navbar Info */}
              <div className="hidden sm:block text-right">
                <p className="text-xs font-semibold text-slate-800">
                  {currentUser?.name}
                </p>

                <p className="text-[11px] text-slate-500">
                  {isAdmin
                    ? 'Administrator'
                    : 'Staff User'}
                </p>
              </div>

              {isDayOpen &&
                sessionDate && (
                  <div className="hidden md:block text-right">
                    <div className="flex items-center justify-end gap-1.5">
                      <span className="h-2 w-2 rounded-full bg-green-500" />

                      <p className="text-[11px] font-semibold uppercase tracking-wide text-green-600">
                        Current Session
                      </p>
                    </div>

                    <p className="mt-0.5 text-sm font-semibold text-slate-800">
                      {formatDate(sessionDate)}
                    </p>
                  </div>
                )}
            </div>
          </div>
        </header>

        <main
          key={currentView}
          className="px-4 sm:px-6 py-5 sm:py-6 animate-[pageEnter_220ms_ease-out]"
        >
          {children}
        </main>
      </div>
    </div>
  )
}

/* --------------------------------
   Summary Card
-------------------------------- */

function SummaryCard({
  title,
  value,
  icon,
}) {
  return (
    <div className={`${ui.cardHover} min-w-0 p-3`}>
      <div className="flex items-center justify-between gap-2">
        <p className="min-w-0 truncate text-xs font-semibold text-slate-500 sm:text-sm">
          {title}
        </p>

        {icon && (
          <div
            className={`flex h-8 w-8 shrink-0 items-center justify-center rounded-xl shadow-inner ring-1 ${icon === 'cash'
              ? 'bg-gradient-to-br from-emerald-50 to-emerald-100 text-emerald-600 ring-emerald-100'
              : icon === 'sales'
                ? 'bg-gradient-to-br from-blue-50 to-blue-100 text-blue-600 ring-blue-100'
                : icon === 'refund'
                  ? 'bg-gradient-to-br from-red-50 to-red-100 text-red-600 ring-red-100'
                  : 'bg-gradient-to-br from-slate-50 to-slate-100 text-slate-600 ring-slate-100'
              }`}
          >
            <svg
              width="15"
              height="15"
              viewBox="0 0 24 24"
              fill="none"
              stroke="currentColor"
              strokeWidth="2"
              strokeLinecap="round"
              strokeLinejoin="round"
            >
              {icon === 'cash' && (
                <>
                  <rect x="3" y="6" width="18" height="12" rx="2" />
                  <circle cx="12" cy="12" r="2.5" />
                  <path d="M7 9h.01M17 15h.01" />
                </>
              )}

              {icon === 'sales' && (
                <>
                  <path d="M4 19V9" />
                  <path d="M10 19V5" />
                  <path d="M16 19v-8" />
                  <path d="M22 19V3" />
                </>
              )}

              {icon === 'refund' && (
                <>
                  <path d="M3 12a9 9 0 1 0 3-6.7" />
                  <path d="M3 4v6h6" />
                  <path d="M12 8v4l3 2" />
                </>
              )}

              {icon === 'total' && (
                <>
                  <rect x="4" y="3" width="16" height="18" rx="2" />
                  <path d="M8 7h8M8 11h8M8 15h5" />
                </>
              )}
            </svg>
          </div>
        )}
      </div>

      <p className="mt-2 text-lg font-bold tracking-tight text-slate-900 sm:text-xl">
        {value}
      </p>
    </div>
  )
}

/* --------------------------------
   Sale Category
-------------------------------- */

function SaleCategory({
  category,
  title,
  amount,
  children,
}) {
  return (
    <div className={`${ui.cardHover} h-full min-w-0 p-4`}>
      <div className="flex items-center justify-between gap-3">
        <div className="flex items-center gap-3 min-w-0">
          <div
            className={`flex h-10 w-10 shrink-0 items-center justify-center rounded-xl ring-1 ${category === 'medicine'
              ? 'bg-blue-50 text-blue-600 ring-blue-100'
              : category === 'general' || category === 'general_items'
                ? 'bg-slate-50 text-slate-600 ring-slate-200'
                : 'bg-emerald-50 text-emerald-600 ring-emerald-100'
              }`}
          >
            <CategoryIcon
              category={category}
            />
          </div>

          <div className="min-w-0">
            <p className="truncate font-semibold text-slate-900">
              {title}
            </p>

            <p className="mt-0.5 text-sm text-slate-500">
              Today's sales
            </p>
          </div>
        </div>

        <div className="h-2 w-2 shrink-0 rounded-full bg-emerald-500" />
      </div>

      <p className="mt-4 text-2xl font-bold tracking-tight text-slate-900">
        {formatMoney(amount)}
      </p>

      {children}
    </div>
  )
}

/* --------------------------------
   Notification
-------------------------------- */

function Notification({
  type = 'info',
  message,
  onClose,
}) {
  const styles = {
    success: {
      container:
        'border-emerald-200 bg-emerald-50 text-emerald-800',
      icon: 'text-emerald-600',
      title: 'Success',
    },

    error: {
      container:
        'border-red-200 bg-red-50 text-red-800',
      icon: 'text-red-600',
      title: 'Error',
    },

    warning: {
      container:
        'border-amber-200 bg-amber-50 text-amber-800',
      icon: 'text-amber-600',
      title: 'Warning',
    },

    info: {
      container:
        'border-blue-200 bg-blue-50 text-blue-800',
      icon: 'text-blue-600',
      title: 'Info',
    },
  }

  const currentStyle =
    styles[type] || styles.info

  return (
    <div
      className={`fixed right-5 top-5 z-100 w-[calc(100%-2.5rem)] max-w-sm rounded-2xl border p-4 shadow-lg animate-[notificationEnter_220ms_ease-out] ${currentStyle.container}`}
    >
      <div className="flex items-start gap-3">
        {/* Icon */}
        <div
          className={`mt-0.5 shrink-0 ${currentStyle.icon}`}
        >
          {type === 'success' && (
            <svg
              className="h-5 w-5"
              viewBox="0 0 24 24"
              fill="none"
              stroke="currentColor"
              strokeWidth="2.5"
            >
              <path d="m5 12 4 4L19 6" />
            </svg>
          )}

          {type === 'error' && (
            <svg
              className="h-5 w-5"
              viewBox="0 0 24 24"
              fill="none"
              stroke="currentColor"
              strokeWidth="2.5"
            >
              <circle
                cx="12"
                cy="12"
                r="9"
              />
              <path d="M12 8v4" />
              <path d="M12 16h.01" />
            </svg>
          )}

          {type === 'warning' && (
            <svg
              className="h-5 w-5"
              viewBox="0 0 24 24"
              fill="none"
              stroke="currentColor"
              strokeWidth="2.5"
            >
              <path d="M10.3 3.7 2.2 18a2 2 0 0 0 1.7 3h16.2a2 2 0 0 0 1.7-3L13.7 3.7a2 2 0 0 0-3.4 0Z" />
              <path d="M12 9v4" />
              <path d="M12 17h.01" />
            </svg>
          )}

          {type === 'info' && (
            <svg
              className="h-5 w-5"
              viewBox="0 0 24 24"
              fill="none"
              stroke="currentColor"
              strokeWidth="2.5"
            >
              <circle
                cx="12"
                cy="12"
                r="9"
              />
              <path d="M12 11v5" />
              <path d="M12 8h.01" />
            </svg>
          )}
        </div>

        {/* Content */}
        <div className="min-w-0 flex-1">
          <p className="text-sm font-bold">
            {currentStyle.title}
          </p>

          <p className="mt-1 text-sm leading-5">
            {message}
          </p>
        </div>

        {/* Close Button */}
        <button
          type="button"
          onClick={onClose}
          className="shrink-0 rounded-lg p-1 opacity-60 transition hover:bg-black/5 hover:opacity-100"
          aria-label="Close notification"
        >
          <svg
            className="h-4 w-4"
            viewBox="0 0 24 24"
            fill="none"
            stroke="currentColor"
            strokeWidth="2"
          >
            <path d="M6 6l12 12" />
            <path d="M18 6 6 18" />
          </svg>
        </button>
      </div>
    </div>
  )
}

/* --------------------------------
   Modal
-------------------------------- */

function Modal({
  title,
  children,
  onClose,
  wide = false,
}) {
  return (
    <div className="fixed inset-0 z-50 bg-slate-950/45 backdrop-blur-md flex items-center justify-center p-4">
      <div
        className={`bg-white border border-blue-100 rounded-3xl shadow-[0_30px_80px_rgba(15,23,42,0.25)] w-full ${wide ? 'max-w-5xl' : 'max-w-lg'
          } max-h-[90vh] overflow-y-auto`}
      >
        <div className="flex items-center justify-between p-5 border-b border-slate-200/80">
          <h2 className="text-lg font-bold tracking-tight text-slate-900">
            {title}
          </h2>

          <button
            onClick={onClose}
            className="w-9 h-9 rounded-xl hover:bg-slate-100 text-slate-500 transition focus:outline-none focus:ring-4 focus:ring-slate-100">
            ✕
          </button>
        </div>

        <div className="p-5 bg-slate-50/60">
          {children}
        </div>
      </div>
    </div>
  )
}

/* --------------------------------
   Detail Card
-------------------------------- */
function DetailCard({
  title,
  value,
  status,
  icon,
  tone = 'slate',
}) {
  const toneStyles = {
    blue: {
      card:
        'border-blue-100 bg-gradient-to-br from-blue-50/80 via-white to-blue-50/40',
      icon:
        'border-blue-100 bg-blue-50/90 text-blue-600',
      value:
        'text-blue-700',
    },

    emerald: {
      card:
        'border-emerald-100 bg-gradient-to-br from-emerald-50/80 via-white to-emerald-50/40',
      icon:
        'border-emerald-100 bg-emerald-50/90 text-emerald-600',
      value:
        'text-emerald-700',
    },

    red: {
      card:
        'border-red-100 bg-gradient-to-br from-red-50/80 via-white to-red-50/40',
      icon:
        'border-red-100 bg-red-50/90 text-red-600',
      value:
        'text-red-700',
    },

    violet: {
      card:
        'border-violet-100 bg-gradient-to-br from-violet-50/80 via-white to-violet-50/40',
      icon:
        'border-violet-100 bg-violet-50/90 text-violet-600',
      value:
        'text-violet-700',
    },

    sky: {
      card:
        'border-sky-100 bg-gradient-to-br from-sky-50/80 via-white to-sky-50/40',
      icon:
        'border-sky-100 bg-sky-50/90 text-sky-600',
      value:
        'text-sky-700',
    },

    indigo: {
      card:
        'border-indigo-100 bg-gradient-to-br from-indigo-50/80 via-white to-indigo-50/40',
      icon:
        'border-indigo-100 bg-indigo-50/90 text-indigo-600',
      value:
        'text-indigo-700',
    },

    amber: {
      card:
        'border-amber-100 bg-gradient-to-br from-amber-50/80 via-white to-amber-50/40',
      icon:
        'border-amber-100 bg-amber-50/90 text-amber-600',
      value:
        'text-amber-700',
    },

    slate: {
      card:
        'border-slate-200 bg-gradient-to-br from-slate-50/80 via-white to-slate-50/40',
      icon:
        'border-slate-200 bg-slate-50/90 text-slate-600',
      value:
        'text-slate-700',
    },
  }

  const statusStyles = {
    Matched: {
      card:
        'border-emerald-100 bg-gradient-to-br from-emerald-50/80 via-white to-emerald-50/40',
      icon:
        'border-emerald-100 bg-emerald-50/90 text-emerald-600',
      value:
        'text-emerald-700',
    },

    Shortage: {
      card:
        'border-red-100 bg-gradient-to-br from-red-50/80 via-white to-red-50/40',
      icon:
        'border-red-100 bg-red-50/90 text-red-600',
      value:
        'text-red-700',
    },

    'Extra Cash': {
      card:
        'border-amber-100 bg-gradient-to-br from-amber-50/80 via-white to-amber-50/40',
      icon:
        'border-amber-100 bg-amber-50/90 text-amber-600',
      value:
        'text-amber-700',
    },
  }

  const activeTone =
    status && statusStyles[status]
      ? statusStyles[status]
      : toneStyles[tone] || toneStyles.slate

  return (
    <div
      className={`
        group relative min-w-0 overflow-hidden
        rounded-2xl border
        p-4
        shadow-[0_8px_24px_rgba(15,23,42,0.06)]
        transition-all duration-200
        hover:-translate-y-0.5
        hover:shadow-[0_14px_32px_rgba(15,23,42,0.10)]
        ${activeTone.card}
      `}
    >
      {/* Soft background glow */}
      <div className="pointer-events-none absolute -right-5 -top-5 h-16 w-16 rounded-full bg-white/60 blur-xl" />

      <div className="relative flex items-start justify-between gap-3">
        <div className="min-w-0">
          <p className="truncate text-[10px] font-bold uppercase tracking-[0.08em] text-slate-500 sm:text-[11px]">
            {title}
          </p>

          <p
            className={`
              mt-2
              min-w-0
              break-words
              text-lg
              font-extrabold
              tracking-tight
              sm:text-xl
              ${activeTone.value}
            `}
          >
            {value}
          </p>

          {status && (
            <p className="mt-1 text-[10px] font-semibold text-slate-500 sm:text-[11px]">
              {status}
            </p>
          )}
        </div>

        {/* Relatable Icon */}
        <div
          className={`
            flex h-9 w-9 shrink-0 items-center justify-center
            rounded-xl border
            shadow-sm
            transition-transform duration-200
            group-hover:scale-110
            ${activeTone.icon}
          `}
        >
          <svg
            className="h-4 w-4"
            viewBox="0 0 24 24"
            fill="none"
            stroke="currentColor"
            strokeWidth="1.8"
            strokeLinecap="round"
            strokeLinejoin="round"
          >
            {icon === 'calendar' && (
              <>
                <rect
                  x="3"
                  y="4"
                  width="18"
                  height="17"
                  rx="2"
                />
                <path d="M16 2v4M8 2v4M3 10h18" />
                <path d="M8 14h.01M12 14h.01M16 14h.01M8 18h.01M12 18h.01" />
              </>
            )}

            {icon === 'sales' && (
              <>
                <path d="M4 19V5" />
                <path d="M4 19h17" />
                <path d="m7 15 4-4 3 2 6-7" />
              </>
            )}

            {icon === 'refund' && (
              <>
                <path d="M4 7h5V2" />
                <path d="M4 7a8 8 0 1 1-1 9" />
                <path d="M20 17h-5v5" />
                <path d="M20 17a8 8 0 1 1 1-9" />
              </>
            )}

            {icon === 'net' && (
              <>
                <rect x="3" y="5" width="18" height="14" rx="2" />
                <path d="M7 9h10M7 13h4M7 16h3M15 13h2" />
              </>
            )}

            {icon === 'wallet' && (
              <>
                <path d="M4 6h16v14H4a2 2 0 0 1-2-2V6a2 2 0 0 1 2-2h14" />
                <path d="M16 11h5v5h-5a2.5 2.5 0 0 1 0-5Z" />
                <path d="M17 13.5h.01" />
              </>
            )}

            {icon === 'calculator' && (
              <>
                <rect
                  x="4"
                  y="2"
                  width="16"
                  height="20"
                  rx="2"
                />
                <path d="M8 6h8M8 11h2M12 11h2M16 11h.01M8 15h2M12 15h2M16 15h.01M8 19h2M12 19h2M16 19h.01" />
              </>
            )}

            {icon === 'cash' && (
              <>
                <rect
                  x="3"
                  y="6"
                  width="18"
                  height="12"
                  rx="2"
                />
                <circle
                  cx="12"
                  cy="12"
                  r="3"
                />
                <path d="M7 9h.01M17 15h.01" />
              </>
            )}

            {icon === 'difference' && (
              <>
                <path d="M4 7h16" />
                <path d="M4 17h16" />
                <path d="M8 4v6" />
                <path d="M16 14v6" />
                <circle cx="8" cy="7" r="2" />
                <circle cx="16" cy="17" r="2" />
              </>
            )}
          </svg>
        </div>
      </div>
    </div>
  )
}

/* --------------------------------
   Monthly Dashboard & Reports
-------------------------------- */

function MonthlyReports({ records }) {
  const getDefaultMonth = () => {
    const now = new Date()

    return `${now.getFullYear()}-${String(
      now.getMonth() + 1
    ).padStart(2, '0')}`
  }

  const [selectedMonth, setSelectedMonth] =
    useState(getDefaultMonth())
  const [currentPage, setCurrentPage] = useState(1)
  const [isExporting, setIsExporting] = useState(false)

  const monthlyData = useMemo(() => {
    const monthRecords = records.filter(
      (record) =>
        String(
          record.date || ''
        ).slice(0, 7) ===
        selectedMonth
    )

    const workingDays =
      monthRecords.length

    const grossSales =
      monthRecords.reduce(
        (total, record) =>
          total +
          Number(
            record.grossSales || 0
          ),
        0
      )

    const totalRefunds =
      monthRecords.reduce(
        (total, record) =>
          total +
          Number(
            record.totalRefunds || 0
          ),
        0
      )

    const netSales =
      monthRecords.reduce(
        (total, record) =>
          total +
          Number(
            record.netSales || 0
          ),
        0
      )

    const openingCash =
      monthRecords.reduce(
        (total, record) =>
          total +
          Number(
            record.openingCash || 0
          ),
        0
      )

    const expectedClosingCash =
      monthRecords.reduce(
        (total, record) =>
          total +
          Number(
            record.expectedClosingCash ||
            0
          ),
        0
      )

    const actualCash =
      monthRecords.reduce(
        (total, record) =>
          total +
          Number(
            record.actualCash || 0
          ),
        0
      )

    const categoryTotals = {
      medicine:
        monthRecords.reduce(
          (total, record) =>
            total +
            Number(
              record.sales?.medicine ||
              0
            ),
          0
        ),

      general:
        monthRecords.reduce(
          (total, record) =>
            total +
            Number(
              record.sales?.general ||
              0
            ),
          0
        ),

      dispensing:
        monthRecords.reduce(
          (total, record) =>
            total +
            Number(
              record.sales?.dispensing ||
              0
            ),
          0
        ),
    }

    const matchedDays =
      monthRecords.filter(
        (record) =>
          record.status === 'Matched'
      ).length

    const shortageDays =
      monthRecords.filter(
        (record) =>
          record.status === 'Shortage'
      ).length

    const extraCashDays =
      monthRecords.filter(
        (record) =>
          record.status === 'Extra Cash'
      ).length

    const totalShortage =
      monthRecords.reduce(
        (total, record) => {
          const difference =
            Number(
              record.difference || 0
            )

          return difference < 0
            ? total +
            Math.abs(
              difference
            )
            : total
        },
        0
      )

    const totalExtraCash =
      monthRecords.reduce(
        (total, record) => {
          const difference =
            Number(
              record.difference || 0
            )

          return difference > 0
            ? total + difference
            : total
        },
        0
      )

    const overallDifference =
      monthRecords.reduce(
        (total, record) =>
          total +
          Number(
            record.difference || 0
          ),
        0
      )

    return {
      monthRecords,
      workingDays,
      grossSales,
      totalRefunds,
      netSales,
      openingCash,
      expectedClosingCash,
      actualCash,
      categoryTotals,
      matchedDays,
      shortageDays,
      extraCashDays,
      totalShortage,
      totalExtraCash,
      overallDifference,
    }
  }, [records, selectedMonth])

  const {
    monthRecords,
    workingDays,
    grossSales,
    totalRefunds,
    netSales,
    openingCash,
    expectedClosingCash,
    actualCash,
    categoryTotals,
    matchedDays,
    shortageDays,
    extraCashDays,
    totalShortage,
    totalExtraCash,
    overallDifference,
  } = monthlyData

  const overallStatus =
    overallDifference === 0
      ? 'Matched'
      : overallDifference < 0
        ? 'Shortage'
        : 'Extra Cash'

  const recordsPerPage = 7

  const totalPages = Math.max(
    1,
    Math.ceil(
      monthRecords.length /
      recordsPerPage
    )
  )

  const paginatedRecords =
    monthRecords.slice(
      (currentPage - 1) *
      recordsPerPage,
      currentPage *
      recordsPerPage
    )

  return (
    <section
      className={`${ui.section} mt-2 w-full min-w-0 overflow-hidden px-4 py-5 sm:px-6 sm:py-7`}
    >
      <div className="flex min-w-0 flex-col gap-4 sm:flex-row sm:items-end sm:justify-between">
        <div className="min-w-0 w-full sm:w-auto">
          <label className="mb-2 block text-[11px] font-bold uppercase tracking-wide text-slate-500">
            Select Month
          </label>

          <input
            type="month"
            value={selectedMonth}
            onChange={(e) => {
              setSelectedMonth(
                e.target.value
              )
              setCurrentPage(1)
            }}
            className={`${ui.inputCompact} w-full sm:w-52`}
          />
        </div>

        <div className="inline-flex w-fit max-w-full shrink-0 items-center gap-2 rounded-xl border border-blue-100 bg-blue-50 px-3 py-2 text-xs font-bold text-blue-700">
          <span className="h-1.5 w-1.5 rounded-full bg-blue-500"></span>

          {formatMonth(
            selectedMonth
          )}
        </div>
      </div>

      {monthRecords.length ===
        0 ? (
        <div className="p-10 text-center">
          <div className="mx-auto flex h-12 w-12 items-center justify-center rounded-2xl bg-slate-50 text-slate-400 ring-1 ring-slate-200">
            <svg
              width="22"
              height="22"
              viewBox="0 0 24 24"
              fill="none"
              stroke="currentColor"
              strokeWidth="1.8"
              strokeLinecap="round"
              strokeLinejoin="round"
            >
              <path d="M6 3h12v18H6z" />
              <path d="M9 7h6M9 11h6M9 15h4" />
            </svg>
          </div>

          <h3 className="mt-4 text-base font-bold text-slate-900">
            No completed records for this month
          </h3>

          <p className="mx-auto mt-1 max-w-sm text-sm leading-6 text-slate-500">
            Complete a daily closing to see its data in the monthly report.
          </p>
        </div>
      ) : (
        <div className="space-y-6 p-4 sm:space-y-7 sm:p-6">
          <div>
            <h3 className="text-base font-bold tracking-tight text-slate-900 mb-3">
              Monthly Summary
            </h3>

            <div className="grid min-w-0 grid-cols-1 gap-3 sm:grid-cols-2 lg:grid-cols-4">
              <DetailCard
                title="Working Days"
                value={workingDays}
                icon="calendar"
                tone="blue"
              />

              <DetailCard
                title="Gross Sales"
                value={formatMoney(grossSales)}
                icon="sales"
                tone="emerald"
              />

              <DetailCard
                title="Total Refunds"
                value={formatMoney(totalRefunds)}
                icon="refund"
                tone="red"
              />

              <DetailCard
                title="Net Sales"
                value={formatMoney(netSales)}
                icon="net"
                tone="violet"
              />
            </div>
          </div>

          <div>
            <h3 className="text-base font-bold tracking-tight text-slate-900 mb-3">
              Monthly Cash Summary
            </h3>

            <div className="grid min-w-0 grid-cols-1 gap-3 sm:grid-cols-2 lg:grid-cols-4">
              <DetailCard
                title="Opening Cash"
                value={formatMoney(openingCash)}
                icon="wallet"
                tone="blue"
              />

              <DetailCard
                title="Expected Cash"
                value={formatMoney(expectedClosingCash)}
                icon="calculator"
                tone="sky"
              />

              <DetailCard
                title="Actual Cash"
                value={formatMoney(actualCash)}
                icon="cash"
                tone="indigo"
              />

              <DetailCard
                title="Overall Difference"
                value={formatMoney(
                  Math.abs(overallDifference)
                )}
                icon="difference"
                status={overallStatus}
              />
            </div>

            <p className="mt-3 max-w-full text-[10px] font-medium leading-4 tracking-tight text-slate-500 sm:text-xs sm:leading-5">
              Opening, expected, and actual cash are monthly totals across completed sessions.
            </p>
          </div>

          <div className="min-w-0">
            <div className="mb-3 flex min-w-0 flex-col gap-1 sm:flex-row sm:items-end sm:justify-between">
              <div className="min-w-0">
                <h3 className="text-base font-bold tracking-tight text-slate-900">
                  Sales by Category
                </h3>
              </div>

              <span className="w-fit max-w-full shrink-0 rounded-lg bg-slate-100 px-2.5 py-1 text-[11px] font-bold uppercase tracking-wide text-slate-600">
                {formatMonth(selectedMonth)}
              </span>
            </div>

            <div className="grid min-w-0 grid-cols-1 gap-3 sm:grid-cols-3">
              <MonthlyCategoryCard
                category="medicine"
                title="Medicine"
                amount={
                  categoryTotals.medicine
                }
              />

              <MonthlyCategoryCard
                category="general"
                title="General Items"
                amount={
                  categoryTotals.general
                }
              />

              <MonthlyCategoryCard
                category="dispensing"
                title="Dispensing"
                amount={
                  categoryTotals.dispensing
                }
              />
            </div>
          </div>

          <div className="min-w-0">
            <div className="mb-3 flex min-w-0 flex-col gap-1 sm:flex-row sm:items-end sm:justify-between">
              <div className="min-w-0">
                <h3 className="text-base font-bold tracking-tight text-slate-900">
                  Daily Status Overview
                </h3>

                <p className="mt-1 whitespace-nowrap text-[11px] font-medium tracking-tight text-slate-500 sm:text-xs">
                  Completed daily sessions by closing status.
                </p>
              </div>

              <span className="w-fit max-w-full shrink-0 text-xs font-semibold text-slate-500">
                {workingDays} {workingDays === 1 ? 'day' : 'days'} recorded
              </span>
            </div>

            <div className="grid min-w-0 grid-cols-1 gap-3 sm:grid-cols-3">
              <MonthlyStatusCard
                title="Matched Days"
                count={matchedDays}
                status="Matched"
              />

              <MonthlyStatusCard
                title="Shortage Days"
                count={
                  shortageDays
                }
                status="Shortage"
              />

              <MonthlyStatusCard
                title="Extra Cash Days"
                count={
                  extraCashDays
                }
                status="Extra Cash"
              />
            </div>
          </div>

          <div className="min-w-0">
            <div className="mb-3 flex min-w-0 flex-col gap-1">
              <div className="min-w-0">
                <h3 className="text-base font-bold tracking-tight text-slate-900">
                  Shortage & Extra Cash
                </h3>

                <p className="mt-1 whitespace-nowrap text-[11px] font-medium tracking-tight text-slate-500 sm:text-xs">
                  Monthly cash differences from daily closing.
                </p>
              </div>
            </div>

            <div className="grid min-w-0 grid-cols-1 gap-3 sm:grid-cols-2">
              <div className="min-w-0 rounded-2xl border border-red-100 bg-red-50 p-4 shadow-sm transition hover:-translate-y-0.5 hover:shadow-md sm:p-5">
                <div className="flex min-w-0 items-center justify-between gap-3">
                  <p className="min-w-0 truncate text-xs font-bold uppercase tracking-wide text-red-600">
                    Total Shortage
                  </p>

                  <span className="h-2 w-2 shrink-0 rounded-full bg-red-500" />
                </div>

                <p className="mt-3 break-words text-xl font-bold tracking-tight text-red-700 sm:text-2xl">
                  {formatMoney(
                    totalShortage
                  )}
                </p>

                <p className="mt-1 whitespace-nowrap text-[11px] font-medium text-red-600/70 sm:text-xs">
                  Across {shortageDays} {shortageDays === 1 ? 'day' : 'days'}
                </p>
              </div>

              <div className="min-w-0 rounded-2xl border border-amber-100 bg-amber-50 p-4 shadow-sm transition hover:-translate-y-0.5 hover:shadow-md sm:p-5">
                <div className="flex min-w-0 items-center justify-between gap-3">
                  <p className="min-w-0 truncate text-xs font-bold uppercase tracking-wide text-amber-600">
                    Total Extra Cash
                  </p>

                  <span className="h-2 w-2 shrink-0 rounded-full bg-amber-500" />
                </div>

                <p className="mt-3 break-words text-xl font-bold tracking-tight text-amber-700 sm:text-2xl">
                  {formatMoney(
                    totalExtraCash
                  )}
                </p>

                <p className="mt-1 whitespace-nowrap text-[11px] font-medium text-amber-600/70 sm:text-xs">
                  Across {extraCashDays} {extraCashDays === 1 ? 'day' : 'days'}
                </p>
              </div>
            </div>
          </div>

          <div className="min-w-0">
            <div className="mb-3 flex min-w-0 flex-col gap-1">
              <h3 className="text-lg font-bold tracking-tight text-slate-900">
                Daily Breakdown
              </h3>

              <p className="whitespace-nowrap text-[11px] font-medium tracking-tight text-slate-500 sm:text-xs">
                Daily records for{' '}
                {formatMonth(
                  selectedMonth
                )}
                .
              </p>
            </div>

            <div className="overflow-hidden rounded-xl border border-slate-200 shadow-sm sm:rounded-2xl">

              {/* Table */}
              <div className="min-w-0 overflow-x-auto scroll-smooth">
                <table className="min-w-[760px] w-full">
                  <thead className="border-b border-slate-200/80 bg-slate-50">
                    <tr>
                      <th className="whitespace-nowrap px-3 py-2.5 text-left align-middle text-[11px] font-bold uppercase tracking-[0.06em] text-slate-500">
                        Date
                      </th>

                      <th className="whitespace-nowrap px-3 py-2.5 text-left align-middle text-[11px] font-bold uppercase tracking-[0.06em] text-slate-500">
                        Gross Sales
                      </th>

                      <th className="whitespace-nowrap px-3 py-2.5 text-left align-middle text-[11px] font-bold uppercase tracking-[0.06em] text-slate-500">
                        Refunds
                      </th>

                      <th className="whitespace-nowrap px-3 py-2.5 text-left align-middle text-[11px] font-bold uppercase tracking-[0.06em] text-slate-500">
                        Net Sales
                      </th>

                      <th className="whitespace-nowrap px-3 py-2.5 text-left align-middle text-[11px] font-bold uppercase tracking-[0.06em] text-slate-500">
                        Difference
                      </th>

                      <th className="whitespace-nowrap px-3 py-2.5 text-left align-middle text-[11px] font-bold uppercase tracking-[0.06em] text-slate-500">
                        Status
                      </th>
                    </tr>
                  </thead>

                  <tbody className="divide-y divide-slate-200/60 bg-white">
                    {paginatedRecords.map(
                      (record) => (
                        <tr key={record.id} className="align-middle transition-colors duration-150 hover:bg-slate-50/80">
                          <td className="whitespace-nowrap px-3 py-3 text-[12px] font-bold text-slate-900">
                            {formatDate(
                              record.date
                            )}
                          </td>

                          <td className="whitespace-nowrap px-3 py-2.5 text-[12px] font-medium text-slate-700">
                            {formatMoney(
                              record.grossSales
                            )}
                          </td>

                          <td className="whitespace-nowrap px-3 py-2.5 text-[12px] font-medium text-slate-700">
                            {formatMoney(
                              record.totalRefunds
                            )}
                          </td>

                          <td className="whitespace-nowrap px-3 py-2.5 text-[12px]  font-bold text-slate-900">
                            {formatMoney(
                              record.netSales
                            )}
                          </td>

                          <td
                            className={`whitespace-nowrap px-3 py-2.5 text-[12px] font-bold tracking-tight ${Number(
                              record.difference || 0
                            ) < 0
                              ? 'text-red-600'
                              : Number(
                                record.difference || 0
                              ) > 0
                                ? 'text-amber-600'
                                : 'text-emerald-600'
                              }`}
                          >
                            {Number(
                              record.difference || 0
                            ) > 0
                              ? '+'
                              : Number(
                                record.difference || 0
                              ) < 0
                                ? '-'
                                : ''}

                            {formatMoney(
                              Math.abs(
                                Number(
                                  record.difference || 0
                                )
                              )
                            )}
                          </td>

                          <td className="whitespace-nowrap px-3 py-3 text-left align-middle">
                            <StatusBadge
                              status={
                                record.status
                              }
                            />
                          </td>
                        </tr>
                      )
                    )}
                  </tbody>
                </table>
              </div>

              {/* Pagination */}
              {totalPages > 1 && (
                <div className="border-t border-slate-200 bg-white px-4 py-3 sm:px-5 sm:py-4">
                  <div className="flex items-center justify-between gap-4">

                    {/* Left Side */}
                    <div className="shrink-0">

                      {/* Desktop Previous */}
                      <button
                        type="button"
                        onClick={() =>
                          setCurrentPage(
                            (page) =>
                              Math.max(
                                1,
                                page - 1
                              )
                          )
                        }
                        disabled={
                          currentPage === 1
                        }
                        className={`${ui.secondaryButton} hidden rounded-lg px-3 py-2 text-xs sm:inline-flex`}
                        aria-label="Previous page"
                      >
                        Previous
                      </button>

                      {/* Mobile Previous */}
                      <button
                        type="button"
                        onClick={() =>
                          setCurrentPage(
                            (page) =>
                              Math.max(
                                1,
                                page - 1
                              )
                          )
                        }
                        disabled={
                          currentPage === 1
                        }
                        className={`${ui.secondaryButton} flex h-9 w-9 items-center justify-center rounded-lg p-0 text-lg sm:hidden`}
                        aria-label="Previous page"
                      >
                        ‹
                      </button>

                    </div>

                    {/* Page Information */}
                    <p className="whitespace-nowrap text-xs font-semibold text-slate-500">
                      Page {currentPage} of {totalPages}
                    </p>

                    {/* Right Side */}
                    <div className="shrink-0">

                      {/* Desktop Next */}
                      <button
                        type="button"
                        onClick={() =>
                          setCurrentPage(
                            (page) =>
                              Math.min(
                                totalPages,
                                page + 1
                              )
                          )
                        }
                        disabled={
                          currentPage === totalPages
                        }
                        className={`${ui.secondaryButton} hidden rounded-lg px-3 py-2 text-xs sm:inline-flex`}
                        aria-label="Next page"
                      >
                        Next
                      </button>

                      {/* Mobile Next */}
                      <button
                        type="button"
                        onClick={() =>
                          setCurrentPage(
                            (page) =>
                              Math.min(
                                totalPages,
                                page + 1
                              )
                          )
                        }
                        disabled={
                          currentPage === totalPages
                        }
                        className={`${ui.secondaryButton} flex h-9 w-9 items-center justify-center rounded-lg p-0 text-lg sm:hidden`}
                        aria-label="Next page"
                      >
                        ›
                      </button>

                    </div>

                  </div>
                </div>
              )}

            </div>
          </div>
        </div>
      )}
    </section>
  )
}

/* --------------------------------
   Monthly Category Card
-------------------------------- */
function MonthlyCategoryCard({
  category,
  title,
  amount,
}) {
  const categoryStyles = {
    medicine: {
      container:
        'border-blue-100/80 bg-gradient-to-br from-white via-blue-50/40 to-blue-50/70 hover:border-blue-200 hover:shadow-[0_12px_28px_rgba(37,99,235,0.11)]',
      iconBox:
        'border-blue-100 bg-blue-50/90 text-blue-600',
      glow:
        'bg-blue-100/60',
      amount:
        'text-blue-700',
    },

    general: {
      container:
        'border-slate-200/80 bg-gradient-to-br from-white via-slate-50/40 to-slate-100/60 hover:border-slate-300 hover:shadow-[0_12px_28px_rgba(71,85,105,0.10)]',
      iconBox:
        'border-slate-200 bg-slate-50/90 text-slate-600',
      glow:
        'bg-slate-200/60',
      amount:
        'text-slate-700',
    },

    dispensing: {
      container:
        'border-violet-100/80 bg-gradient-to-br from-white via-violet-50/40 to-violet-50/70 hover:border-violet-200 hover:shadow-[0_12px_28px_rgba(124,58,237,0.10)]',
      iconBox:
        'border-violet-100 bg-violet-50/90 text-violet-600',
      glow:
        'bg-violet-100/60',
      amount:
        'text-violet-700',
    },
  }

  const styles =
    categoryStyles[category] ||
    categoryStyles.medicine

  return (
    <div
      className={`
        group relative overflow-hidden
        rounded-2xl
        border
        px-4 py-4
        shadow-[0_8px_24px_rgba(15,23,42,0.06)]
        backdrop-blur-md
        transition-all duration-200
        hover:-translate-y-0.5
        ${styles.container}
      `}
    >
      {/* Soft Background Glow */}
      <div
        className={`
          pointer-events-none absolute
          -right-6 -top-6
          h-20 w-20
          rounded-full
          opacity-70
          blur-2xl
          transition-transform duration-300
          group-hover:scale-125
          ${styles.glow}
        `}
      />

      {/* Card Header */}
      <div className="relative flex items-center justify-between gap-3">
        <div className="flex min-w-0 items-center gap-3">
          {/* Category Icon */}
          <div
            className={`
              flex h-10 w-10 shrink-0
              items-center justify-center
              rounded-xl
              border
              shadow-[0_4px_10px_rgba(15,23,42,0.06)]
              transition-all duration-200
              group-hover:scale-105
              ${styles.iconBox}
            `}
          >
            <CategoryIcon
              category={category}
              size={20}
            />
          </div>

          <div className="min-w-0">
            <p className="text-[9px] font-bold uppercase tracking-[0.11em] text-slate-400">
              Category
            </p>

            <p className="mt-0.5 truncate text-sm font-bold text-slate-800">
              {title}
            </p>
          </div>
        </div>

        {/* Small Category Icon */}
        <div
          className={`
            flex h-7 w-7 shrink-0
            items-center justify-center
            rounded-lg
            border
            bg-white/70
            shadow-sm
            backdrop-blur-sm
            transition-all duration-200
            group-hover:scale-110
            ${styles.iconBox}
          `}
        >
          <CategoryIcon
            category={category}
            size={14}
          />
        </div>
      </div>

      {/* Total Sales */}
      <div className="relative mt-3 border-t border-slate-100/80 pt-3">
        <p className="text-[9px] font-semibold uppercase tracking-[0.08em] text-slate-400">
          Total Sales
        </p>

        <p
          className={`
            mt-0.5
            break-words
            text-xl
            font-extrabold
            tracking-tight
            sm:text-2xl
            ${styles.amount}
          `}
        >
          {formatMoney(amount)}
        </p>
      </div>
    </div>
  )
}

function MonthlyStatusCard({
  title,
  count,
  status,
}) {
  let statusClass =
    'bg-slate-50 border-slate-200 text-slate-900'

  let statusDot =
    'bg-slate-400'

  if (status === 'Matched') {
    statusClass =
      'bg-emerald-50 border-emerald-100 text-emerald-700'

    statusDot =
      'bg-emerald-500'
  }

  if (status === 'Shortage') {
    statusClass =
      'bg-red-50 border-red-100 text-red-700'

    statusDot =
      'bg-red-500'
  }

  if (status === 'Extra Cash') {
    statusClass =
      'bg-amber-50 border-amber-100 text-amber-700'

    statusDot =
      'bg-amber-500'
  }

  return (
    <div
      className={`min-w-0 rounded-2xl border p-4 shadow-sm transition hover:-translate-y-0.5 hover:shadow-md sm:p-5 ${statusClass}`}
    >
      <div className="flex min-w-0 items-center justify-between gap-3">
        <p className="whitespace-nowrap text-[11px] font-bold uppercase tracking-tight opacity-75 sm:text-xs">
          {title}
        </p>

        <span
          className={`h-2 w-2 shrink-0 rounded-full ring-2 ring-white/70 ${statusDot}`}
        />
      </div>

      <p className="mt-3 text-2xl font-bold tracking-tight sm:text-3xl">
        {count}
      </p>

      <p className="mt-1 whitespace-nowrap text-[11px] font-semibold opacity-70 sm:text-xs">
        {count === 1 ? 'completed day' : 'completed days'}
      </p>
    </div>
  )
}

/* --------------------------------
   Daily History
-------------------------------- */

function DailyHistory({
  records,
  onRefresh,
  onViewDetails,
}) {
  const [searchTerm, setSearchTerm] =
    useState('')

  const [statusFilter, setStatusFilter] =
    useState('all')

  const [
    categoryFilter,
    setCategoryFilter,
  ] = useState('all')

  const [dateFilter, setDateFilter] =
    useState('')

  const [currentPage, setCurrentPage] =
    useState(1)

  const recordsPerPage = 10

  const filteredRecords = useMemo(() => {
    const search =
      searchTerm
        .trim()
        .toLowerCase()

    return records.filter(
      (record) => {
        const formattedDate =
          formatDate(
            record.date
          ).toLowerCase()

        const rawDate = String(
          record.date || ''
        ).toLowerCase()

        const recordId = String(
          record.id || ''
        ).toLowerCase()

        const sessionId = String(
          record.sessionId || ''
        ).toLowerCase()

        const matchesSearch =
          !search ||
          formattedDate.includes(
            search
          ) ||
          rawDate.includes(search) ||
          recordId.includes(search) ||
          sessionId.includes(search)

        const matchesStatus =
          statusFilter === 'all' ||
          record.status ===
          statusFilter

        const matchesCategory =
          categoryFilter ===
          'all' ||
          Number(
            record.sales?.[
            categoryFilter
            ] || 0
          ) > 0

        const matchesDate =
          !dateFilter ||
          record.date === dateFilter

        return (
          matchesSearch &&
          matchesStatus &&
          matchesCategory &&
          matchesDate
        )
      }
    )
  }, [
    records,
    searchTerm,
    statusFilter,
    categoryFilter,
    dateFilter,
  ])

  const hasActiveFilters =
    searchTerm.trim() !== '' ||
    statusFilter !== 'all' ||
    categoryFilter !== 'all' ||
    dateFilter !== ''

  useEffect(() => {
    setCurrentPage(1)
  }, [
    searchTerm,
    statusFilter,
    categoryFilter,
    dateFilter,
  ])

  const totalPages =
    Math.max(
      1,
      Math.ceil(
        filteredRecords.length /
        recordsPerPage
      )
    )

  const paginatedRecords =
    filteredRecords.slice(
      (currentPage - 1) *
      recordsPerPage,
      currentPage *
      recordsPerPage
    )

  const clearFilters = () => {
    setSearchTerm('')
    setStatusFilter('all')
    setCategoryFilter('all')
    setDateFilter('')
  }

  return (
    <section className={`${ui.section} overflow-hidden`}>
      <div className="border-b border-slate-200/80 bg-white p-5">
        <div className="flex flex-col gap-4 lg:flex-row lg:items-center lg:justify-between">

          <button
            onClick={onRefresh}
            className={`${ui.primaryButton} inline-flex items-center justify-center gap-2 px-4 py-2.5 text-sm shadow-sm hover:-translate-y-0.5 hover:shadow-md`}
          >
            <svg
              className="h-4 w-4"
              viewBox="0 0 24 24"
              fill="none"
              stroke="currentColor"
              strokeWidth="2"
            >
              <path
                d="M20 11a8.1 8.1 0 0 0-15.5-2M4 5v4h4"
                strokeLinecap="round"
                strokeLinejoin="round"
              />

              <path
                d="M4 13a8.1 8.1 0 0 0 15.5 2M20 19v-4h-4"
                strokeLinecap="round"
                strokeLinejoin="round"
              />
            </svg>

            Refresh
          </button>
        </div>
      </div>

      <div className="border-b border-slate-200/80 bg-slate-50/70 p-5">
        <div className="grid grid-cols-1 gap-4 md:grid-cols-2 xl:grid-cols-4">
          <div className="xl:col-span-1">
            <label className="mb-1.5 block text-xs font-bold uppercase tracking-wide text-slate-600">
              Search
            </label>

            <div className="relative">
              <svg
                className="pointer-events-none absolute left-3 top-1/2 h-4 w-4 -translate-y-1/2 text-slate-500"
                viewBox="0 0 24 24"
                fill="none"
                stroke="currentColor"
                strokeWidth="2"
              >
                <circle
                  cx="11"
                  cy="11"
                  r="7"
                />

                <path
                  d="m20 20-4-4"
                  strokeLinecap="round"
                />
              </svg>

              <input
                type="text"
                value={searchTerm}
                onChange={(e) =>
                  setSearchTerm(
                    e.target.value
                  )
                }
                placeholder="Search date or session ID..."
                className={`${ui.inputCompact} pl-9`}
              />
            </div>
          </div>

          <div>
            <label className="mb-1.5 block text-xs font-bold uppercase tracking-wide text-slate-600">
              Date
            </label>

            <input
              type="date"
              value={dateFilter}
              onChange={(e) =>
                setDateFilter(
                  e.target.value
                )
              }
              className={ui.inputCompact}
            />
          </div>

          <div>
            <label className="mb-1.5 block text-xs font-bold uppercase tracking-wide text-slate-600">
              Status
            </label>

            <select
              value={statusFilter}
              onChange={(e) =>
                setStatusFilter(
                  e.target.value
                )
              }
              className={ui.inputCompact}
            >
              <option value="all">
                All Statuses
              </option>

              <option value="Matched">
                Matched
              </option>

              <option value="Shortage">
                Shortage
              </option>

              <option value="Extra Cash">
                Extra Cash
              </option>
            </select>
          </div>

          <div>
            <label className="mb-1.5 block text-xs font-bold uppercase tracking-wide text-slate-600">
              Category
            </label>

            <select
              value={categoryFilter}
              onChange={(e) =>
                setCategoryFilter(
                  e.target.value
                )
              }
              className={ui.inputCompact}
            >
              <option value="all">
                All Categories
              </option>

              <option value="medicine">
                Medicine
              </option>

              <option value="general">
                General Items
              </option>

              <option value="dispensing">
                Dispensing
              </option>
            </select>
          </div>
        </div>

        <div className="mt-4 flex flex-col gap-3 sm:flex-row sm:items-center sm:justify-between">
          <div className="text-sm font-medium text-slate-500">
            Showing{' '}
            <span className="font-bold text-slate-900">
              {
                filteredRecords.length
              }
            </span>{' '}
            of{' '}
            <span className="font-bold text-slate-900">
              {records.length}
            </span>{' '}
            records
          </div>

          {hasActiveFilters && (
            <button
              onClick={clearFilters}
              className={`${ui.secondaryButton} inline-flex items-center justify-center gap-2 px-4 py-2 text-sm`}
            >
              Clear Filters
            </button>
          )}
        </div>
      </div>

      {records.length === 0 ? (
        <div className="p-10 text-center">
          <div className="mx-auto flex h-12 w-12 items-center justify-center rounded-2xl bg-slate-50 text-slate-400 ring-1 ring-slate-200">
            <svg
              width="22"
              height="22"
              viewBox="0 0 24 24"
              fill="none"
              stroke="currentColor"
              strokeWidth="1.8"
              strokeLinecap="round"
              strokeLinejoin="round"
            >
              <path d="M6 3h12v18H6z" />
              <path d="M9 7h6M9 11h6M9 15h4" />
            </svg>
          </div>

          <h3 className="mt-4 text-base font-bold text-slate-900">
            No daily records yet
          </h3>

          <p className="mx-auto mt-1 max-w-sm text-sm leading-6 text-slate-500">
            Completed daily sessions will appear here.
          </p>
        </div>
      ) : filteredRecords.length ===
        0 ? (
        <div className="p-10 text-center">
          <div className="mx-auto mb-4 flex h-12 w-12 items-center justify-center rounded-2xl bg-slate-100 text-slate-400">
            <svg
              className="h-7 w-7 text-slate-500"
              viewBox="0 0 24 24"
              fill="none"
              stroke="currentColor"
              strokeWidth="2"
            >
              <circle
                cx="11"
                cy="11"
                r="7"
              />

              <path d="m20 20-4-4" />
            </svg>
          </div>

          <h3 className="mt-4 text-base font-bold text-slate-900">
            No records match your filters
          </h3>

          <p className="mx-auto mt-1.5 max-w-md text-sm leading-6 text-slate-500">
            Try changing your search or filter settings.
          </p>

          <button
            onClick={clearFilters}
            className={`${ui.primaryButton} mt-5 px-4 py-2.5 text-sm shadow-sm`}
          >
            Clear Filters
          </button>
        </div>
      ) : (
        <><div className="overflow-x-auto">
          <table className="min-w-full">
            <thead className="border-b border-slate-200 bg-slate-50">
              <tr>
                <th className="px-5 py-3 text-left text-xs font-bold uppercase tracking-wide text-slate-500">
                  Date
                </th>

                <th className="px-5 py-3 text-left text-xs font-bold uppercase tracking-wide text-slate-500">
                  Opening
                </th>

                <th className="px-5 py-3 text-left text-xs font-bold uppercase tracking-wide text-slate-500">
                  Net Sales
                </th>

                <th className="px-5 py-3 text-left text-xs font-bold uppercase tracking-wide text-slate-500">
                  Actual Cash
                </th>

                <th className="px-5 py-3 text-left text-xs font-bold uppercase tracking-wide text-slate-500">
                  Status
                </th>

                <th className="px-5 py-3 text-right text-xs font-bold uppercase tracking-wide text-slate-500">
                  Action
                </th>
              </tr>
            </thead>

            <tbody className="divide-y divide-slate-100">
              {paginatedRecords.map(
                (record) => (
                  <tr
                    key={record.id}
                    className="hover:bg-slate-50"
                  >
                    <td className="whitespace-nowrap px-5 py-4 text-sm font-semibold text-slate-900">
                      {formatDate(
                        record.date
                      )}
                    </td>

                    <td className="whitespace-nowrap px-5 py-4 text-sm text-slate-700">
                      {formatMoney(
                        record.openingCash
                      )}
                    </td>

                    <td className="whitespace-nowrap px-5 py-4 text-sm font-semibold text-slate-900">
                      {formatMoney(
                        record.netSales
                      )}
                    </td>

                    <td className="whitespace-nowrap px-5 py-4 text-sm text-slate-700">
                      {formatMoney(
                        record.actualCash
                      )}
                    </td>

                    <td className="whitespace-nowrap px-5 py-4">
                      <StatusBadge
                        status={record.status} />
                    </td>

                    <td className="whitespace-nowrap px-5 py-4 text-right">
                      <button
                        onClick={() => onViewDetails(
                          record
                        )}
                        className="rounded-xl border border-slate-300 bg-white px-3 py-2 text-sm font-semibold text-slate-700 transition hover:border-blue-200 hover:bg-blue-50 hover:text-blue-700 focus:outline-none focus:ring-2 focus:ring-blue-100"
                      >
                        View Details
                      </button>
                    </td>
                  </tr>
                )
              )}
            </tbody>
          </table>
        </div>
          <div className="mt-4 border-t border-slate-200/80 px-4 pb-4 pt-4 sm:px-5 sm:pb-5 sm:pt-4">
            <div className="flex items-center justify-between gap-4">

              {/* Previous */}
              <div className="shrink-0">
                {/* Desktop Previous */}
                <button
                  type="button"
                  onClick={() =>
                    setCurrentPage(
                      (page) =>
                        Math.max(
                          1,
                          page - 1
                        )
                    )
                  }
                  disabled={
                    currentPage === 1
                  }
                  className={`${ui.secondaryButton} hidden rounded-lg px-3 py-2 text-xs sm:inline-flex`}
                >
                  Previous
                </button>

                {/* Mobile Previous */}
                <button
                  type="button"
                  onClick={() =>
                    setCurrentPage(
                      (page) =>
                        Math.max(
                          1,
                          page - 1
                        )
                    )
                  }
                  disabled={
                    currentPage === 1
                  }
                  className={`${ui.secondaryButton} flex h-9 w-9 items-center justify-center rounded-lg p-0 text-lg sm:hidden`}
                  aria-label="Previous page"
                >
                  ‹
                </button>
              </div>

              {/* Page Information */}
              <p className="whitespace-nowrap text-xs font-semibold text-slate-500 sm:text-sm">
                Page {currentPage} of {totalPages}
              </p>

              {/* Next */}
              <div className="shrink-0">
                {/* Desktop Next */}
                <button
                  type="button"
                  onClick={() =>
                    setCurrentPage(
                      (page) =>
                        Math.min(
                          totalPages,
                          page + 1
                        )
                    )
                  }
                  disabled={
                    currentPage === totalPages
                  }
                  className={`${ui.secondaryButton} hidden rounded-lg px-3 py-2 text-xs sm:inline-flex`}
                >
                  Next
                </button>

                {/* Mobile Next */}
                <button
                  type="button"
                  onClick={() =>
                    setCurrentPage(
                      (page) =>
                        Math.min(
                          totalPages,
                          page + 1
                        )
                    )
                  }
                  disabled={
                    currentPage === totalPages
                  }
                  className={`${ui.secondaryButton} flex h-9 w-9 items-center justify-center rounded-lg p-0 text-lg sm:hidden`}
                  aria-label="Next page"
                >
                  ›
                </button>
              </div>

            </div>
          </div></>
      )}
    </section>
  )
}

/* --------------------------------
   Audit Logs Table
-------------------------------- */

function AuditLogsTable({ logs }) {
  const [searchTerm, setSearchTerm] = useState('')
  const [actionFilter, setActionFilter] = useState('ALL')
  const [fromDate, setFromDate] = useState('')
  const [toDate, setToDate] = useState('')
  const [currentPage, setCurrentPage] = useState(1)

  const logsPerPage = 10

  const filteredLogs = (logs || []).filter((log) => {
    const search = searchTerm.trim().toLowerCase()

    const matchesSearch =
      !search ||
      (log.username || '')
        .toLowerCase()
        .includes(search) ||
      (log.action || '')
        .toLowerCase()
        .includes(search) ||
      (log.description || '')
        .toLowerCase()
        .includes(search)

    const matchesAction =
      actionFilter === 'ALL' ||
      log.action === actionFilter

    const logDate = new Date(log.timestamp)

    const matchesFromDate =
      !fromDate ||
      logDate >= new Date(`${fromDate}T00:00:00`)

    const matchesToDate =
      !toDate ||
      logDate <= new Date(`${toDate}T23:59:59.999`)

    return (
      matchesSearch &&
      matchesAction &&
      matchesFromDate &&
      matchesToDate
    )
  })

  const hasActiveFilters =
    searchTerm.trim() !== '' ||
    actionFilter !== 'ALL' ||
    fromDate !== '' ||
    toDate !== ''

  const totalPages = Math.max(
    1,
    Math.ceil(
      filteredLogs.length / logsPerPage
    )
  )

  const paginatedLogs = filteredLogs.slice(
    (currentPage - 1) * logsPerPage,
    currentPage * logsPerPage
  )

  useEffect(() => {
    setCurrentPage(1)
  }, [
    searchTerm,
    actionFilter,
    fromDate,
    toDate,
  ])

  useEffect(() => {
    if (currentPage > totalPages) {
      setCurrentPage(totalPages)
    }
  }, [
    currentPage,
    totalPages,
  ])

  const handleClearFilters = () => {
    setSearchTerm('')
    setActionFilter('ALL')
    setFromDate('')
    setToDate('')
    setCurrentPage(1)
  }

  if (!logs || logs.length === 0) {
    return (
      <div className="rounded-2xl border border-slate-200 bg-white p-10 text-center shadow-sm">
        <div className="mx-auto flex h-14 w-14 items-center justify-center rounded-full bg-slate-100">
          <svg
            className="h-7 w-7 text-slate-500"
            viewBox="0 0 24 24"
            fill="none"
            stroke="currentColor"
            strokeWidth="2"
          >
            <path d="M4 4h16v16H4z" />
            <path d="M8 8h8M8 12h8M8 16h5" />
          </svg>
        </div>

        <h3 className="mt-4 text-base font-bold text-slate-900">
          No audit logs yet
        </h3>

        <p className="mt-1 text-sm text-slate-500">
          Important system activities will appear here.
        </p>
      </div>
    )
  }

  return (
    <div className="rounded-2xl border border-slate-200 bg-white shadow-sm">
      {/* Filters */}
      <div className="border-b border-slate-200 p-5">
        <div className="grid grid-cols-1 gap-4 lg:grid-cols-5">

          {/* Search */}
          <div className="lg:col-span-2">
            <label
              htmlFor="audit-search"
              className="mb-1.5 block text-xs font-semibold text-slate-500"
            >
              Search
            </label>

            <div className="relative">
              <svg
                className="absolute left-3 top-1/2 h-5 w-5 -translate-y-1/2 text-slate-400"
                viewBox="0 0 24 24"
                fill="none"
                stroke="currentColor"
                strokeWidth="2"
              >
                <circle
                  cx="11"
                  cy="11"
                  r="7"
                />

                <path d="m20 20-4-4" />
              </svg>

              <input
                id="audit-search"
                type="text"
                value={searchTerm}
                onChange={(event) =>
                  setSearchTerm(event.target.value)
                }
                placeholder="Search audit logs..."
                className={`${ui.inputCompact} pl-10`}
              />
            </div>
          </div>

          {/* Action Filter */}
          <div>
            <label
              htmlFor="audit-action-filter"
              className="mb-1.5 block text-xs font-semibold text-slate-500"
            >
              Filter by Action
            </label>

            <select
              id="audit-action-filter"
              value={actionFilter}
              onChange={(event) =>
                setActionFilter(event.target.value)
              }
              className={ui.inputCompact}
            >
              <option value="ALL">
                All Actions
              </option>

              <option value="LOGIN">
                LOGIN
              </option>

              <option value="LOGOUT">
                LOGOUT
              </option>

              <option value="START_DAILY_SESSION">
                START_DAILY_SESSION
              </option>

              <option value="ADD_SALE">
                ADD_SALE
              </option>

              <option value="ADD_REFUND">
                ADD_REFUND
              </option>

              <option value="DAILY_CLOSING">
                DAILY_CLOSING
              </option>

              <option value="START_NEW_DAY">
                START_NEW_DAY
              </option>
            </select>
          </div>

          {/* From Date */}
          <div>
            <label
              htmlFor="audit-from-date"
              className="mb-1.5 block text-xs font-semibold text-slate-500"
            >
              From Date
            </label>

            <input
              id="audit-from-date"
              type="date"
              value={fromDate}
              onChange={(event) =>
                setFromDate(event.target.value)
              }
              className={ui.inputCompact}
            />
          </div>

          {/* To Date */}
          <div>
            <label
              htmlFor="audit-to-date"
              className="mb-1.5 block text-xs font-semibold text-slate-500"
            >
              To Date
            </label>

            <input
              id="audit-to-date"
              type="date"
              value={toDate}
              onChange={(event) =>
                setToDate(event.target.value)
              }
              className={ui.inputCompact}
            />
          </div>

          {/* Clear Filters */}
          <div className="flex items-end">
            <button
              type="button"
              onClick={handleClearFilters}
              disabled={!hasActiveFilters}
              className={`${ui.secondaryButton} w-full py-2.5 text-sm disabled:cursor-not-allowed disabled:opacity-50`}
            >
              Clear Filters
            </button>
          </div>
        </div>

        {/* Export + Result Count */}
        <div className="mt-5 flex flex-col gap-3 sm:flex-row sm:items-center sm:justify-between">
          <p className="text-xs font-medium text-slate-500">
            Showing {filteredLogs.length} of {logs.length} audit logs
          </p>
        </div>
      </div>

      {/* Audit Logs Table */}
      <div className="overflow-x-auto">
        <table className="min-w-[900px] w-full table-fixed">

          <thead className="bg-slate-50/90">
            <tr>
              <th className="w-[150px] whitespace-nowrap px-4 py-3 text-left text-[10px] font-bold uppercase tracking-wide text-slate-500 sm:px-5 sm:text-[11px]">
                Date & Time
              </th>

              <th className="w-[125px] whitespace-nowrap px-4 py-3 text-left text-[10px] font-bold uppercase tracking-wide text-slate-500 sm:px-5 sm:text-[11px]">
                User
              </th>

              <th className="w-[150px] whitespace-nowrap px-4 py-3 text-left text-[10px] font-bold uppercase tracking-wide text-slate-500 sm:px-5 sm:text-[11px]">
                Action
              </th>

              <th className="px-4 py-3 text-left text-[10px] font-bold uppercase tracking-wide text-slate-500 sm:px-5 sm:text-[11px]">
                Description
              </th>
            </tr>
          </thead>

          <tbody className="divide-y divide-slate-200/70">

            {paginatedLogs.map((log) => (
              <tr
                key={log.id}
                className="transition hover:bg-slate-50"
              >

                {/* Date & Time */}
                <td className="whitespace-nowrap px-4 py-3 align-middle font-mono text-[10px] text-slate-700 sm:px-5 sm:py-3.5 sm:text-xs">
                  <div className="whitespace-nowrap font-semibold text-slate-800">
                    {new Date(
                      log.timestamp
                    ).toLocaleDateString()}
                  </div>

                  <div className="mt-0.5 whitespace-nowrap text-[9px] text-slate-500 sm:text-[10px]">
                    {formatTime(
                      log.timestamp
                    )}
                  </div>
                </td>

                {/* User */}
                <td className="w-[190px] max-w-[190px] px-4 py-3 align-middle sm:px-5 sm:py-3.5">
                  <div className="min-w-0">
                    <div className="truncate font-mono text-[10px] font-semibold text-slate-800 sm:text-xs">
                      {log.username || 'System'}
                    </div>

                    {log.userId && (
                      <div
                        className="mt-0.5 max-w-full break-all font-mono text-[9px] leading-4 text-slate-400 sm:text-[10px]"
                        title={`ID: ${log.userId}`}
                      >
                        ID: {log.userId}
                      </div>
                    )}
                  </div>
                </td>

                {/* Action */}
                <td className="w-[150px] max-w-[150px] px-4 py-3 align-middle sm:px-5 sm:py-3.5">
                  <span className="inline-flex max-w-full whitespace-nowrap rounded-md border border-blue-100 bg-blue-50 px-2 py-1 font-mono text-[9px] font-bold tracking-tight text-blue-700 sm:text-[10px]">
                    {log.action}
                  </span>
                </td>

                {/* Description */}
                <td className="min-w-[360px] px-4 py-3 align-middle sm:px-5 sm:py-3.5">
                  <p className="font-mono text-[10px] leading-5 text-slate-600 sm:text-[11px] sm:leading-5">
                    {log.description ||
                      'No description available.'}
                  </p>
                </td>

              </tr>
            ))}

          </tbody>
        </table>
      </div>

      {totalPages > 1 && (
        <div className="border-t border-slate-200/80 px-4 pb-4 pt-4 sm:px-5 sm:pb-5 sm:pt-4">
          <div className="flex items-center justify-between gap-4">

            {/* Previous */}
            <div className="shrink-0">
              <button
                type="button"
                onClick={() =>
                  setCurrentPage((page) =>
                    Math.max(1, page - 1)
                  )
                }
                disabled={currentPage === 1}
                className={`${ui.secondaryButton} hidden px-3 py-2 text-xs sm:inline-flex`}
                aria-label="Previous audit logs page"
              >
                Previous
              </button>

              <button
                type="button"
                onClick={() =>
                  setCurrentPage((page) =>
                    Math.max(1, page - 1)
                  )
                }
                disabled={currentPage === 1}
                className={`${ui.secondaryButton} flex h-9 w-9 items-center justify-center p-0 text-lg sm:hidden`}
                aria-label="Previous audit logs page"
              >
                ‹
              </button>
            </div>

            {/* Page Information */}
            <p className="whitespace-nowrap text-[10px] font-semibold text-slate-500 sm:text-xs">
              Page {currentPage} of {totalPages}
            </p>

            {/* Next */}
            <div className="shrink-0">
              <button
                type="button"
                onClick={() =>
                  setCurrentPage((page) =>
                    Math.min(totalPages, page + 1)
                  )
                }
                disabled={currentPage === totalPages}
                className={`${ui.secondaryButton} hidden px-3 py-2 text-xs sm:inline-flex`}
                aria-label="Next audit logs page"
              >
                Next
              </button>

              <button
                type="button"
                onClick={() =>
                  setCurrentPage((page) =>
                    Math.min(totalPages, page + 1)
                  )
                }
                disabled={currentPage === totalPages}
                className={`${ui.secondaryButton} flex h-9 w-9 items-center justify-center p-0 text-lg sm:hidden`}
                aria-label="Next audit logs page"
              >
                ›
              </button>
            </div>

          </div>
        </div>
      )}
    </div>
  )
}

function NotificationToast({ type, message, onClose }) {
  const styles = {
    success: {
      container:
        'border-emerald-200 bg-emerald-50 text-emerald-800',
      icon: 'text-emerald-600',
      title: 'Success',
    },

    error: {
      container:
        'border-red-200 bg-red-50 text-red-800',
      icon: 'text-red-600',
      title: 'Error',
    },

    warning: {
      container:
        'border-amber-200 bg-amber-50 text-amber-800',
      icon: 'text-amber-600',
      title: 'Warning',
    },

    info: {
      container:
        'border-blue-200 bg-blue-50 text-blue-800',
      icon: 'text-blue-600',
      title: 'Info',
    },
  }

  const currentStyle =
    styles[type] || styles.info

  return (
    <div
      className={`fixed right-5 top-5 z-100 w-[calc(100%-2.5rem)] max-w-sm rounded-2xl border p-4 shadow-[0_18px_45px_rgba(15,23,42,0.16)] backdrop-blur-xl ${currentStyle.container}`}
    >
      <div className="flex items-start gap-3">

        {/* Icon */}
        <div
          className={`mt-0.5 shrink-0 ${currentStyle.icon}`}
        >
          {type === 'success' && (
            <svg
              className="h-5 w-5"
              viewBox="0 0 24 24"
              fill="none"
              stroke="currentColor"
              strokeWidth="2.5"
            >
              <path d="m5 12 4 4L19 6" />
            </svg>
          )}

          {type === 'error' && (
            <svg
              className="h-5 w-5"
              viewBox="0 0 24 24"
              fill="none"
              stroke="currentColor"
              strokeWidth="2.5"
            >
              <circle
                cx="12"
                cy="12"
                r="9"
              />
              <path d="M12 8v4" />
              <path d="M12 16h.01" />
            </svg>
          )}

          {type === 'warning' && (
            <svg
              className="h-5 w-5"
              viewBox="0 0 24 24"
              fill="none"
              stroke="currentColor"
              strokeWidth="2.5"
            >
              <path d="M10.3 3.7 2.2 18a2 2 0 0 0 1.7 3h16.2a2 2 0 0 0 1.7-3L13.7 3.7a2 2 0 0 0-3.4 0Z" />
              <path d="M12 9v4" />
              <path d="M12 17h.01" />
            </svg>
          )}

          {type === 'info' && (
            <svg
              className="h-5 w-5"
              viewBox="0 0 24 24"
              fill="none"
              stroke="currentColor"
              strokeWidth="2.5"
            >
              <circle
                cx="12"
                cy="12"
                r="9"
              />
              <path d="M12 11v5" />
              <path d="M12 8h.01" />
            </svg>
          )}
        </div>

        {/* Content */}
        <div className="min-w-0 flex-1">
          <p className="text-sm font-bold">
            {currentStyle.title}
          </p>

          <p className="mt-1 text-sm leading-5">
            {message}
          </p>
        </div>

        {/* Close Button */}
        <button
          type="button"
          onClick={onClose}
          className="shrink-0 rounded-lg p-1 opacity-60 transition hover:bg-black/5 hover:opacity-100"
          aria-label="Close notification"
        >
          <svg
            className="h-4 w-4"
            viewBox="0 0 24 24"
            fill="none"
            stroke="currentColor"
            strokeWidth="2"
          >
            <path d="M6 6l12 12" />
            <path d="M18 6 6 18" />
          </svg>
        </button>

      </div>
    </div>
  )
}

/* --------------------------------
   Status Badge
-------------------------------- */

function StatusBadge({ status }) {
  const classes = {
    Matched:
      'border border-emerald-200 bg-emerald-50 text-emerald-700',

    Shortage:
      'border border-red-200 bg-red-50 text-red-700',

    'Extra Cash':
      'border border-amber-200 bg-amber-50 text-amber-700',
  }

  return (
    <span
      className={`inline-flex items-center rounded-full border px-2.5 py-1 text-xs font-semibold ${classes[status] ||
        'border border-slate-200 bg-slate-100 text-slate-700'
        }`}
    >
      {status}
    </span>
  )
}

/* --------------------------------
   Daily Record Details Screen
-------------------------------- */
function DailyRecordDetails({
  record,
  onBack,
}) {
  const [transactions, setTransactions] = useState([])
  const [allTransactions, setAllTransactions] = useState([])
  const [loading, setLoading] = useState(true)
  const [transactionPage, setTransactionPage] = useState(1)

  const transactionsPerPage = 5

  useEffect(() => {
    async function loadTransactions() {
      try {
        const currentRecordTransactions =
          await getTransactionsBySession(
            record.sessionId || record.id
          )

        const currentTransactions =
          currentRecordTransactions || []

        setTransactions(currentTransactions)

        /*
         * Load transactions from completed sessions as well.
         *
         * This is required because a refund can be created
         * in today's session while its original sale belongs
         * to an older/completed session.
         */
        const completedSessions =
          await getDailySessions()

        const linkedTransactions = [
          ...currentTransactions,
        ]

        for (const session of completedSessions || []) {
          if (
            session?.sessionId &&
            session.sessionId !==
            (record.sessionId || record.id)
          ) {
            const sessionTransactions =
              await getTransactionsBySession(
                session.sessionId
              )

            linkedTransactions.push(
              ...(sessionTransactions || [])
            )
          }
        }

        setAllTransactions(linkedTransactions)
      } catch (error) {
        console.error(error)

        setTransactions([])
        setAllTransactions([])
      } finally {
        setLoading(false)
      }
    }

    loadTransactions()
  }, [record])

  useEffect(() => {
    setTransactionPage(1)
  }, [record])

  const totalTransactionPages = Math.max(
    1,
    Math.ceil(
      transactions.length / transactionsPerPage
    )
  )

  const paginatedTransactions = transactions.slice(
    (transactionPage - 1) * transactionsPerPage,
    transactionPage * transactionsPerPage
  )

  const difference = Number(record.difference || 0)

  const status =
    record.status ||
    (difference === 0
      ? 'Matched'
      : difference < 0
        ? 'Shortage'
        : 'Extra Cash')

  const categoryTotals = {
    medicine: Number(record.sales?.medicine || 0),
    general: Number(record.sales?.general || 0),
    dispensing: Number(record.sales?.dispensing || 0),
  }

  const getTransactionAmount = (transaction) => {
    const amount = Number(transaction.amount || 0)

    return transaction.type === 'refund'
      ? -amount
      : amount
  }

  return (
    <PageContainer>
      {/* --------------------------------
         Header
      -------------------------------- */}
      <div className="mb-6">
        <div className="flex flex-col gap-3 sm:flex-row sm:items-center sm:justify-between">
          <button
            type="button"
            onClick={onBack}
            className={ui.secondaryButton}
          >
            ← Back to Daily Records
          </button>

          <StatusBadge status={status} />
        </div>

        <div className="mt-5">
          <h1 className="text-2xl font-bold text-slate-900">
            Daily Record Details
          </h1>

          <p className="mt-1 text-sm text-slate-500">
            {formatDate(record.date)}
          </p>
        </div>
      </div>

      {/* --------------------------------
   Financial Summary
-------------------------------- */}
      <section
        className="
    relative overflow-hidden rounded-2xl
    border border-blue-100/80
    bg-gradient-to-br from-white via-blue-50/40 to-slate-50
    shadow-[0_14px_35px_rgba(15,23,42,0.09)]
  "
      >
        {/* Subtle glass highlight */}
        <div className="pointer-events-none absolute -right-20 -top-20 h-40 w-40 rounded-full bg-blue-200/20 blur-3xl" />
        <div className="pointer-events-none absolute -bottom-24 -left-16 h-40 w-40 rounded-full bg-cyan-200/15 blur-3xl" />

        {/* Header */}
        <div className="relative flex items-center justify-between border-b border-blue-100/70 px-4 py-3 sm:px-5">
          <div>
            <h2 className="text-sm font-bold tracking-tight text-slate-900 sm:text-base">
              Financial Summary
            </h2>
            <p className="mt-0.5 text-[10px] font-medium text-slate-500 sm:text-xs">
              Daily cash and sales overview
            </p>
          </div>

          <div className="flex h-8 w-8 items-center justify-center rounded-xl border border-blue-100 bg-white/70 text-blue-600 shadow-[0_4px_12px_rgba(37,99,235,0.10)] backdrop-blur-md">
            <svg
              className="h-4 w-4"
              viewBox="0 0 24 24"
              fill="none"
              stroke="currentColor"
              strokeWidth="1.8"
            >
              <path d="M4 19V5" />
              <path d="M4 19h16" />
              <path d="m7 15 3-4 3 2 5-6" />
            </svg>
          </div>
        </div>

        {/* Financial Metrics */}
        <div className="relative grid grid-cols-2 gap-2.5 p-3 sm:grid-cols-4 sm:gap-3 sm:p-4">

          {/* Opening Cash */}
          <div
            className="
        group relative overflow-hidden rounded-xl
        border border-blue-100/80
        bg-white/75
        px-3 py-2.5
        shadow-[0_5px_14px_rgba(15,23,42,0.06)]
        backdrop-blur-md
        transition-all duration-200
        hover:-translate-y-0.5
        hover:border-blue-200
        hover:shadow-[0_9px_20px_rgba(37,99,235,0.10)]
        sm:px-4 sm:py-3
      "
          >
            <div className="absolute right-2 top-2 flex h-8 w-8 items-center justify-center rounded-lg border border-blue-100 bg-blue-50/80 text-blue-600 shadow-sm transition-all duration-200 group-hover:scale-110 group-hover:bg-blue-100">
              <svg
                className="h-4 w-4"
                viewBox="0 0 24 24"
                fill="none"
                stroke="currentColor"
                strokeWidth="1.8"
                strokeLinecap="round"
                strokeLinejoin="round"
              >
                <rect x="3" y="6" width="18" height="13" rx="2" />
                <path d="M7 6V4h10v2" />
                <path d="M8 12h8" />
                <path d="M10 15h4" />
              </svg>
            </div>

            <p className="relative text-[10px] font-semibold uppercase tracking-wide text-slate-500 sm:text-[11px]">
              Opening Cash
            </p>

            <p className="relative mt-1 text-sm font-extrabold tracking-tight text-slate-900 sm:text-base">
              {formatMoney(record.openingCash)}
            </p>
          </div>

          {/* Gross Sales */}
          <div
            className="
        group relative overflow-hidden rounded-xl
        border border-emerald-100/80
        bg-gradient-to-br from-white/85 to-emerald-50/50
        px-3 py-2.5
        shadow-[0_5px_14px_rgba(16,185,129,0.07)]
        backdrop-blur-md
        transition-all duration-200
        hover:-translate-y-0.5
        hover:border-emerald-200
        hover:shadow-[0_9px_20px_rgba(16,185,129,0.11)]
        sm:px-4 sm:py-3
      "
          >
            <div className="absolute right-2 top-2 flex h-8 w-8 items-center justify-center rounded-lg border border-emerald-100 bg-emerald-50/80 text-emerald-600 shadow-sm transition-all duration-200 group-hover:scale-110 group-hover:bg-emerald-100">
              <svg
                className="h-4 w-4"
                viewBox="0 0 24 24"
                fill="none"
                stroke="currentColor"
                strokeWidth="1.8"
                strokeLinecap="round"
                strokeLinejoin="round"
              >
                <path d="M4 19V5" />
                <path d="M4 19h16" />
                <path d="m7 15 3-4 3 2 5-6" />
              </svg>
            </div>

            <p className="relative text-[10px] font-semibold uppercase tracking-wide text-slate-500 sm:text-[11px]">
              Gross Sales
            </p>

            <p className="relative mt-1 text-sm font-extrabold tracking-tight text-emerald-700 sm:text-base">
              {formatMoney(record.grossSales)}
            </p>
          </div>

          {/* Refunds */}
          <div
            className="
        group relative overflow-hidden rounded-xl
        border border-red-100/80
        bg-gradient-to-br from-white/85 to-red-50/45
        px-3 py-2.5
        shadow-[0_5px_14px_rgba(239,68,68,0.06)]
        backdrop-blur-md
        transition-all duration-200
        hover:-translate-y-0.5
        hover:border-red-200
        hover:shadow-[0_9px_20px_rgba(239,68,68,0.10)]
        sm:px-4 sm:py-3
      "
          >
            <div className="absolute right-2 top-2 flex h-8 w-8 items-center justify-center rounded-lg border border-red-100 bg-red-50/80 text-red-600 shadow-sm transition-all duration-200 group-hover:scale-110 group-hover:bg-red-100">
              <svg
                className="h-4 w-4"
                viewBox="0 0 24 24"
                fill="none"
                stroke="currentColor"
                strokeWidth="1.8"
                strokeLinecap="round"
                strokeLinejoin="round"
              >
                <path d="M9 7H5l3-3" />
                <path d="M5 7a7 7 0 1 1-1 8" />
                <path d="M15 17h4l-3 3" />
                <path d="M19 17a7 7 0 1 1 1-8" />
              </svg>
            </div>

            <p className="relative text-[10px] font-semibold uppercase tracking-wide text-slate-500 sm:text-[11px]">
              Refunds
            </p>

            <p className="relative mt-1 text-sm font-extrabold tracking-tight text-red-600 sm:text-base">
              {formatMoney(record.totalRefunds)}
            </p>
          </div>

          {/* Net Sales */}
          <div
            className="
        group relative overflow-hidden rounded-xl
        border border-violet-100/80
        bg-gradient-to-br from-white/85 to-violet-50/45
        px-3 py-2.5
        shadow-[0_5px_14px_rgba(124,58,237,0.06)]
        backdrop-blur-md
        transition-all duration-200
        hover:-translate-y-0.5
        hover:border-violet-200
        hover:shadow-[0_9px_20px_rgba(124,58,237,0.10)]
        sm:px-4 sm:py-3
      "
          >
            <div className="absolute right-2 top-2 flex h-8 w-8 items-center justify-center rounded-lg border border-violet-100 bg-violet-50/80 text-violet-600 shadow-sm transition-all duration-200 group-hover:scale-110 group-hover:bg-violet-100">
              <svg
                className="h-4 w-4"
                viewBox="0 0 24 24"
                fill="none"
                stroke="currentColor"
                strokeWidth="1.8"
                strokeLinecap="round"
                strokeLinejoin="round"
              >
                <rect x="3" y="5" width="18" height="14" rx="2" />
                <path d="M7 9h10" />
                <path d="M7 13h4" />
                <path d="M7 16h3" />
                <path d="M15 13h2" />
              </svg>
            </div>

            <p className="relative text-[10px] font-semibold uppercase tracking-wide text-slate-500 sm:text-[11px]">
              Net Sales
            </p>

            <p className="relative mt-1 text-sm font-extrabold tracking-tight text-violet-700 sm:text-base">
              {formatMoney(record.netSales)}
            </p>
          </div>

          {/* Expected Cash */}
          <div
            className="
        group relative overflow-hidden rounded-xl
        border border-sky-100/80
        bg-gradient-to-br from-white/85 to-sky-50/50
        px-3 py-2.5
        shadow-[0_5px_14px_rgba(14,165,233,0.06)]
        backdrop-blur-md
        transition-all duration-200
        hover:-translate-y-0.5
        hover:border-sky-200
        hover:shadow-[0_9px_20px_rgba(14,165,233,0.10)]
        sm:px-4 sm:py-3
      "
          >
            <div className="absolute right-2 top-2 flex h-8 w-8 items-center justify-center rounded-lg border border-sky-100 bg-sky-50/80 text-sky-600 shadow-sm transition-all duration-200 group-hover:scale-110 group-hover:bg-sky-100">
              <svg
                className="h-4 w-4"
                viewBox="0 0 24 24"
                fill="none"
                stroke="currentColor"
                strokeWidth="1.8"
                strokeLinecap="round"
                strokeLinejoin="round"
              >
                <rect x="4" y="3" width="16" height="18" rx="2" />
                <path d="M8 7h8" />
                <path d="M8 11h2" />
                <path d="M12 11h2" />
                <path d="M16 11h0" />
                <path d="M8 15h2" />
                <path d="M12 15h2" />
                <path d="M16 15h0" />
              </svg>
            </div>

            <p className="relative text-[10px] font-semibold uppercase tracking-wide text-slate-500 sm:text-[11px]">
              Expected Cash
            </p>

            <p className="relative mt-1 text-sm font-extrabold tracking-tight text-sky-700 sm:text-base">
              {formatMoney(record.expectedClosingCash)}
            </p>
          </div>

          {/* Actual Cash */}
          <div
            className="
        group relative overflow-hidden rounded-xl
        border border-indigo-100/80
        bg-gradient-to-br from-white/85 to-indigo-50/45
        px-3 py-2.5
        shadow-[0_5px_14px_rgba(79,70,229,0.06)]
        backdrop-blur-md
        transition-all duration-200
        hover:-translate-y-0.5
        hover:border-indigo-200
        hover:shadow-[0_9px_20px_rgba(79,70,229,0.10)]
        sm:px-4 sm:py-3
      "
          >
            <div className="absolute right-2 top-2 flex h-8 w-8 items-center justify-center rounded-lg border border-indigo-100 bg-indigo-50/80 text-indigo-600 shadow-sm transition-all duration-200 group-hover:scale-110 group-hover:bg-indigo-100">
              <svg
                className="h-4 w-4"
                viewBox="0 0 24 24"
                fill="none"
                stroke="currentColor"
                strokeWidth="1.8"
                strokeLinecap="round"
                strokeLinejoin="round"
              >
                <rect x="3" y="6" width="18" height="12" rx="2" />
                <circle cx="12" cy="12" r="3" />
                <path d="M7 10h.01" />
                <path d="M17 14h.01" />
              </svg>
            </div>

            <p className="relative text-[10px] font-semibold uppercase tracking-wide text-slate-500 sm:text-[11px]">
              Actual Cash
            </p>

            <p className="relative mt-1 text-sm font-extrabold tracking-tight text-indigo-700 sm:text-base">
              {formatMoney(record.actualCash)}
            </p>
          </div>

          {/* Difference */}
          <div
            className={`
        group relative overflow-hidden rounded-xl
        border
        px-3 py-2.5
        backdrop-blur-md
        shadow-[0_5px_14px_rgba(15,23,42,0.06)]
        transition-all duration-200
        hover:-translate-y-0.5
        sm:px-4 sm:py-3
        ${difference === 0
                ? 'border-emerald-100/80 bg-gradient-to-br from-white/85 to-emerald-50/50 hover:border-emerald-200 hover:shadow-[0_9px_20px_rgba(16,185,129,0.10)]'
                : difference < 0
                  ? 'border-red-100/80 bg-gradient-to-br from-white/85 to-red-50/50 hover:border-red-200 hover:shadow-[0_9px_20px_rgba(239,68,68,0.10)]'
                  : 'border-amber-100/80 bg-gradient-to-br from-white/85 to-amber-50/50 hover:border-amber-200 hover:shadow-[0_9px_20px_rgba(245,158,11,0.10)]'
              }
      `}
          >
            <div
              className={`
    absolute right-2 top-2 flex h-8 w-8 items-center justify-center rounded-lg border shadow-sm transition-all duration-200 group-hover:scale-110
    ${difference === 0
                  ? 'border-emerald-100 bg-emerald-50/80 text-emerald-600 group-hover:bg-emerald-100'
                  : difference < 0
                    ? 'border-red-100 bg-red-50/80 text-red-600 group-hover:bg-red-100'
                    : 'border-amber-100 bg-amber-50/80 text-amber-600 group-hover:bg-amber-100'
                }
  `}
            >
              <svg
                className="h-4 w-4"
                viewBox="0 0 24 24"
                fill="none"
                stroke="currentColor"
                strokeWidth="1.8"
                strokeLinecap="round"
                strokeLinejoin="round"
              >
                <path d="M4 7h16" />
                <path d="M4 17h16" />
                <path d="M8 4v6" />
                <path d="M16 14v6" />
                <circle cx="8" cy="7" r="2" />
                <circle cx="16" cy="17" r="2" />
              </svg>
            </div>

            <p className="relative text-[10px] font-semibold uppercase tracking-wide text-slate-500 sm:text-[11px]">
              Difference
            </p>

            <p
              className={`relative mt-1 text-sm font-extrabold tracking-tight sm:text-base ${difference === 0
                ? 'text-emerald-600'
                : difference < 0
                  ? 'text-red-600'
                  : 'text-amber-600'
                }`}
            >
              {formatMoney(Math.abs(difference))}
            </p>
          </div>

          {/* Final Status */}
          <div
            className={`
        group relative overflow-hidden rounded-xl
        border
        px-3 py-2.5
        backdrop-blur-md
        shadow-[0_5px_14px_rgba(15,23,42,0.06)]
        transition-all duration-200
        hover:-translate-y-0.5
        sm:px-4 sm:py-3
        ${status === 'Matched'
                ? 'border-emerald-100/80 bg-gradient-to-br from-white/85 to-emerald-50/50 hover:border-emerald-200'
                : status === 'Shortage'
                  ? 'border-red-100/80 bg-gradient-to-br from-white/85 to-red-50/50 hover:border-red-200'
                  : 'border-amber-100/80 bg-gradient-to-br from-white/85 to-amber-50/50 hover:border-amber-200'
              }
      `}

          >
            <div
              className={`
    absolute right-2 top-2 flex h-8 w-8 items-center justify-center rounded-lg border shadow-sm transition-all duration-200 group-hover:scale-110
    ${status === 'Matched'
                  ? 'border-emerald-100 bg-emerald-50/80 text-emerald-600 group-hover:bg-emerald-100'
                  : status === 'Shortage'
                    ? 'border-red-100 bg-red-50/80 text-red-600 group-hover:bg-red-100'
                    : 'border-amber-100 bg-amber-50/80 text-amber-600 group-hover:bg-amber-100'
                }
  `}
            >
              <svg
                className="h-4 w-4"
                viewBox="0 0 24 24"
                fill="none"
                stroke="currentColor"
                strokeWidth="1.8"
                strokeLinecap="round"
                strokeLinejoin="round"
              >
                {status === 'Matched' ? (
                  <>
                    <circle cx="12" cy="12" r="8.5" />
                    <path d="m8.5 12 2.3 2.3 4.7-5" />
                  </>
                ) : status === 'Shortage' ? (
                  <>
                    <circle cx="12" cy="12" r="8.5" />
                    <path d="M8 8l8 8" />
                    <path d="m16 8-8 8" />
                  </>
                ) : (
                  <>
                    <circle cx="12" cy="12" r="8.5" />
                    <path d="M12 7v5" />
                    <path d="M12 16h.01" />
                  </>
                )}
              </svg>
            </div>
            <p className="text-[10px] font-semibold uppercase tracking-wide text-slate-500 sm:text-[11px]">
              Final Status
            </p>

            <div className="mt-1.5">
              <StatusBadge status={status} />
            </div>
          </div>
        </div>
      </section>

      {/* --------------------------------
   Closing Information
-------------------------------- */}
      <section
        className="
    relative mt-6 overflow-hidden rounded-2xl
    border border-blue-100/80
    bg-gradient-to-br from-white via-blue-50/30 to-slate-50
    shadow-[0_14px_35px_rgba(15,23,42,0.08)]
  "
      >
        {/* Subtle glass highlights */}
        <div className="pointer-events-none absolute -right-16 -top-16 h-32 w-32 rounded-full bg-blue-200/20 blur-3xl" />
        <div className="pointer-events-none absolute -bottom-20 -left-12 h-32 w-32 rounded-full bg-cyan-200/15 blur-3xl" />

        {/* Header */}
        <div
          className="
      relative flex flex-col gap-3
      border-b border-blue-100/70
      px-4 py-3
      sm:flex-row sm:items-start sm:justify-between
      sm:px-5 sm:py-3.5
    "
        >
          {/* Title */}
          <div className="min-w-0">
            <div className="flex items-center gap-2.5">
              <div
                className="
            flex h-8 w-8 shrink-0 items-center justify-center
            rounded-xl
            border border-blue-100
            bg-white/75
            text-blue-600
            shadow-[0_4px_12px_rgba(37,99,235,0.10)]
            backdrop-blur-md
          "
              >
                <svg
                  className="h-4 w-4"
                  viewBox="0 0 24 24"
                  fill="none"
                  stroke="currentColor"
                  strokeWidth="1.8"
                  strokeLinecap="round"
                  strokeLinejoin="round"
                >
                  <path d="M12 3 5 6v5c0 4.5 2.8 7.8 7 10 4.2-2.2 7-5.5 7-10V6l-7-3Z" />
                  <path d="m9.5 12 1.7 1.7 3.5-3.5" />
                </svg>
              </div>

              <div>
                <h2 className="text-sm font-bold tracking-tight text-slate-900 sm:text-base">
                  Closing Information
                </h2>

                <p className="mt-0.5 text-[10px] font-medium text-slate-500 sm:text-xs">
                  Session closing details
                </p>
              </div>
            </div>
          </div>

          {/* Closing Reason - Top Right */}
          <div
            className="
        relative min-w-0
        rounded-xl
        border border-slate-200/80
        bg-white/70
        px-3 py-2
        shadow-[0_5px_14px_rgba(15,23,42,0.05)]
        backdrop-blur-md
        sm:max-w-[52%]
        sm:px-3.5 sm:py-2.5
      "
          >
            <div className="flex items-start gap-2.5">
              <div
                className="
            mt-0.5 flex h-6 w-6 shrink-0 items-center justify-center
            rounded-lg
            bg-blue-50
            text-blue-600
          "
              >
                <svg
                  className="h-3.5 w-3.5"
                  viewBox="0 0 24 24"
                  fill="none"
                  stroke="currentColor"
                  strokeWidth="1.8"
                  strokeLinecap="round"
                  strokeLinejoin="round"
                >
                  <path d="M4 5h16v11H8l-4 4V5Z" />
                  <path d="M8 9h8" />
                  <path d="M8 12h5" />
                </svg>
              </div>

              <div className="min-w-0">
                <p className="text-[9px] font-bold uppercase tracking-[0.12em] text-slate-400 sm:text-[10px]">
                  Closing Reason
                </p>

                <p className="mt-0.5 text-xs font-medium leading-5 text-slate-700 sm:text-sm">
                  {record.reason || 'No closing reason provided.'}
                </p>
              </div>
            </div>
          </div>
        </div>

        {/* Closing Details */}
        <div className="relative grid grid-cols-1 gap-2.5 p-3 sm:grid-cols-2 lg:grid-cols-4 sm:p-4">

          {/* Closed By */}
          <div
            className="
        group relative overflow-hidden rounded-xl
        border border-blue-100/80
        bg-white/75
        px-3 py-2.5
        shadow-[0_5px_14px_rgba(15,23,42,0.05)]
        backdrop-blur-md
        transition-all duration-200
        hover:-translate-y-0.5
        hover:border-blue-200
        hover:shadow-[0_9px_20px_rgba(37,99,235,0.09)]
        sm:px-3.5 sm:py-3
      "
          >
            <div
              className="
    absolute right-2 top-2
    flex h-7 w-7 items-center justify-center
    rounded-lg
    border border-blue-100
    bg-blue-50/80
    text-blue-600
    shadow-sm
    transition-all duration-200
    group-hover:scale-110
    group-hover:bg-blue-100
  "
            >
              <svg
                className="h-3.5 w-3.5"
                viewBox="0 0 24 24"
                fill="none"
                stroke="currentColor"
                strokeWidth="1.8"
                strokeLinecap="round"
                strokeLinejoin="round"
              >
                <path d="M20 21a8 8 0 0 0-16 0" />
                <circle cx="12" cy="7" r="4" />
              </svg>
            </div>

            <p className="relative text-[10px] font-semibold uppercase tracking-wide text-slate-400">
              Closed By
            </p>

            <p className="relative mt-1 text-sm font-bold text-slate-900">
              {record.closedBy?.name || 'N/A'}
            </p>
          </div>

          {/* Username */}
          <div
            className="
        group relative overflow-hidden rounded-xl
        border border-indigo-100/80
        bg-gradient-to-br from-white/80 to-indigo-50/35
        px-3 py-2.5
        shadow-[0_5px_14px_rgba(79,70,229,0.05)]
        backdrop-blur-md
        transition-all duration-200
        hover:-translate-y-0.5
        hover:border-indigo-200
        hover:shadow-[0_9px_20px_rgba(79,70,229,0.09)]
        sm:px-3.5 sm:py-3
      "
          >
            <div
              className="
    absolute right-2 top-2
    flex h-7 w-7 items-center justify-center
    rounded-lg
    border border-indigo-100
    bg-indigo-50/80
    text-indigo-600
    shadow-sm
    transition-all duration-200
    group-hover:scale-110
    group-hover:bg-indigo-100
  "
            >
              <svg
                className="h-3.5 w-3.5"
                viewBox="0 0 24 24"
                fill="none"
                stroke="currentColor"
                strokeWidth="1.8"
                strokeLinecap="round"
                strokeLinejoin="round"
              >
                <path d="M20 21a8 8 0 0 0-16 0" />
                <circle cx="12" cy="7" r="4" />
              </svg>
            </div>

            <p className="relative text-[10px] font-semibold uppercase tracking-wide text-slate-400">
              Username
            </p>

            <p className="relative mt-1 text-sm font-bold text-slate-900">
              {record.closedBy?.username
                ? `@${record.closedBy.username}`
                : 'N/A'}
            </p>
          </div>

          {/* Role */}
          <div
            className="
        group relative overflow-hidden rounded-xl
        border border-violet-100/80
        bg-gradient-to-br from-white/80 to-violet-50/35
        px-3 py-2.5
        shadow-[0_5px_14px_rgba(124,58,237,0.05)]
        backdrop-blur-md
        transition-all duration-200
        hover:-translate-y-0.5
        hover:border-violet-200
        hover:shadow-[0_9px_20px_rgba(124,58,237,0.09)]
        sm:px-3.5 sm:py-3
      "
          >
            <div
              className="
    absolute right-2 top-2
    flex h-7 w-7 items-center justify-center
    rounded-lg
    border border-violet-100
    bg-violet-50/80
    text-violet-600
    shadow-sm
    transition-all duration-200
    group-hover:scale-110
    group-hover:bg-violet-100
  "
            >
              <svg
                className="h-3.5 w-3.5"
                viewBox="0 0 24 24"
                fill="none"
                stroke="currentColor"
                strokeWidth="1.8"
                strokeLinecap="round"
                strokeLinejoin="round"
              >
                <path d="M12 3 5 6v5c0 4.5 2.8 7.8 7 10 4.2-2.2 7-5.5 7-10V6l-7-3Z" />
                <path d="m9.5 12 1.7 1.7 3.5-3.5" />
              </svg>
            </div>

            <p className="relative text-[10px] font-semibold uppercase tracking-wide text-slate-400">
              Role
            </p>

            <p className="relative mt-1 text-sm font-bold capitalize text-slate-900">
              {record.closedBy?.role || 'N/A'}
            </p>
          </div>

          {/* Closed At */}
          <div
            className="
        group relative overflow-hidden rounded-xl
        border border-sky-100/80
        bg-gradient-to-br from-white/80 to-sky-50/35
        px-3 py-2.5
        shadow-[0_5px_14px_rgba(14,165,233,0.05)]
        backdrop-blur-md
        transition-all duration-200
        hover:-translate-y-0.5
        hover:border-sky-200
        hover:shadow-[0_9px_20px_rgba(14,165,233,0.09)]
        sm:px-3.5 sm:py-3
      "
          >
            <div
              className="
    absolute right-2 top-2
    flex h-7 w-7 items-center justify-center
    rounded-lg
    border border-sky-100
    bg-sky-50/80
    text-sky-600
    shadow-sm
    transition-all duration-200
    group-hover:scale-110
    group-hover:bg-sky-100
  "
            >
              <svg
                className="h-3.5 w-3.5"
                viewBox="0 0 24 24"
                fill="none"
                stroke="currentColor"
                strokeWidth="1.8"
                strokeLinecap="round"
                strokeLinejoin="round"
              >
                <circle cx="12" cy="12" r="8.5" />
                <path d="M12 7v5l3 2" />
              </svg>
            </div>

            <p className="relative text-[10px] font-semibold uppercase tracking-wide text-slate-400">
              Closed At
            </p>

            <p className="relative mt-1 text-sm font-bold leading-5 text-slate-900">
              {record.closedAt
                ? `${new Date(
                  record.closedAt
                ).toLocaleDateString('en-PK', {
                  day: '2-digit',
                  month: 'short',
                  year: 'numeric',
                })} ${formatTime(record.closedAt)}`
                : 'N/A'}
            </p>
          </div>
        </div>
      </section>

      {/* --------------------------------
         Sales by Category
      -------------------------------- */}
      <section className={`${ui.section} mt-6`}>
        <div className={`${ui.sectionHeader} px-5 py-4`}>
          <h2 className="text-base font-semibold text-slate-900">
            Sales by Category
          </h2>
        </div>

        <div className="relative grid grid-cols-1 gap-3 p-3 sm:grid-cols-3 sm:p-4">
          <HistoryCategory
            category="medicine"
            amount={categoryTotals.medicine}
          />

          <HistoryCategory
            category="general"
            amount={categoryTotals.general}
          />

          <HistoryCategory
            category="dispensing"
            amount={categoryTotals.dispensing}
          />
        </div>
      </section>

      {/* --------------------------------
   Transaction History
-------------------------------- */}
      <section
        className="
    relative mt-6 overflow-hidden rounded-2xl
    border border-blue-100/80
    bg-gradient-to-br from-white via-blue-50/25 to-slate-50
    shadow-[0_14px_35px_rgba(15,23,42,0.08)]
  "
      >
        {/* Soft glass highlights */}
        <div className="pointer-events-none absolute -right-20 -top-20 h-40 w-40 rounded-full bg-blue-200/20 blur-3xl" />
        <div className="pointer-events-none absolute -bottom-24 -left-16 h-40 w-40 rounded-full bg-cyan-200/10 blur-3xl" />

        {/* Header */}
        <div
          className="
      relative flex items-center justify-between
      border-b border-blue-100/70
      px-4 py-3
      sm:px-5 sm:py-3.5
    "
        >
          <div className="flex min-w-0 items-center gap-2.5">
            {/* Header Icon */}
            <div
              className="
          flex h-8 w-8 shrink-0 items-center justify-center
          rounded-xl
          border border-blue-100
          bg-white/75
          text-blue-600
          shadow-[0_4px_12px_rgba(37,99,235,0.10)]
          backdrop-blur-md
        "
            >
              <svg
                className="h-4 w-4"
                viewBox="0 0 24 24"
                fill="none"
                stroke="currentColor"
                strokeWidth="1.8"
                strokeLinecap="round"
                strokeLinejoin="round"
              >
                <path d="M6 3h12v18H6z" />
                <path d="M9 7h6" />
                <path d="M9 11h6" />
                <path d="M9 15h4" />
              </svg>
            </div>

            <div className="min-w-0">
              <h2 className="text-sm font-bold tracking-tight text-slate-900 sm:text-base">
                Transaction History
              </h2>

              <p className="mt-0.5 truncate text-[10px] font-medium text-slate-500 sm:text-xs">
                Complete sales and refund records for this session
              </p>
            </div>
          </div>

          {/* Record Count */}
          <div
            className="
        ml-3 shrink-0
        rounded-full
        border border-blue-100
        bg-white/75
        px-2.5 py-1
        text-[10px] font-bold
        text-blue-700
        shadow-sm
        backdrop-blur-md
        sm:px-3 sm:py-1.5 sm:text-xs
      "
          >
            {transactions.length} record
            {transactions.length !== 1 ? 's' : ''}
          </div>
        </div>

        {loading ? (
          <div className="relative px-5 py-8 text-center text-sm text-slate-500">
            Loading transactions...
          </div>
        ) : transactions.length === 0 ? (
          <div className="relative px-5 py-8 text-center text-sm text-slate-500">
            No transactions found for this daily record.
          </div>
        ) : (
          <>
            {/* Transaction List */}
            <div className="relative space-y-2.5 p-3 sm:p-4">
              {paginatedTransactions.map((transaction) => {
                const signedAmount =
                  getTransactionAmount(transaction)

                const originalSale =
                  transaction.type === 'refund' &&
                    transaction.originalSaleId
                    ? allTransactions.find(
                      (item) =>
                        item.type === 'sale' &&
                        item.id === transaction.originalSaleId
                    )
                    : null

                const isRefund =
                  transaction.type === 'refund'

                return (
                  <div
                    key={transaction.id}
                    className={`
                group relative overflow-hidden rounded-xl
                border
                bg-white/75
                shadow-[0_5px_15px_rgba(15,23,42,0.055)]
                backdrop-blur-md
                transition-all duration-200
                hover:-translate-y-0.5
                ${isRefund
                        ? 'border-red-100/90 hover:border-red-200 hover:shadow-[0_9px_20px_rgba(239,68,68,0.09)]'
                        : 'border-emerald-100/90 hover:border-emerald-200 hover:shadow-[0_9px_20px_rgba(16,185,129,0.09)]'
                      }
              `}
                  >
                    {/* Small status glow */}
                    <div
                      className={`
                  pointer-events-none absolute -right-6 -top-6
                  h-16 w-16 rounded-full blur-2xl opacity-60
                  ${isRefund
                          ? 'bg-red-100'
                          : 'bg-emerald-100'
                        }
                `}
                    />

                    {/* Transaction Content */}
                    {!isRefund ? (
                      /* --------------------------------
                         SALE TRANSACTION - COMPACT ROW
                      -------------------------------- */
                      <div
                        className="
      relative flex flex-col gap-3
      px-3 py-2.5
      sm:flex-row sm:items-center sm:gap-3
      sm:px-4 sm:py-3
    "
                      >
                        {/* Category + Sale + Date/Time */}
                        <div className="flex min-w-0 shrink-0 items-center gap-2.5 sm:w-[22%]">
                          {/* Category Icon */}
                          <div
                            className="
          flex h-9 w-9 shrink-0 items-center justify-center
          rounded-lg border border-blue-100
          bg-blue-50 text-blue-600 shadow-sm
        "
                          >
                            <CategoryIcon
                              category={transaction.category}
                              size={17}
                            />
                          </div>

                          <div className="min-w-0">
                            <div className="flex flex-wrap items-center gap-1.5">
                              <p className="truncate text-xs font-bold text-slate-800 sm:text-sm">
                                {getCategoryName(transaction.category)}
                              </p>

                              <span
                                className="
              rounded-full bg-emerald-100
              px-1.5 py-0.5
              text-[9px] font-bold tracking-wide
              text-emerald-700
            "
                              >
                                SALE
                              </span>
                            </div>

                            <p className="mt-0.5 whitespace-nowrap text-[10px] font-medium text-slate-400 sm:text-[11px]">
                              {transaction.date
                                ? formatDate(transaction.date)
                                : '-'}
                              {' · '}
                              {transaction.time
                                ? formatTime(transaction.time)
                                : '-'}
                            </p>
                          </div>
                        </div>

                        {/* Customer ID */}
                        <div
                          className="
        min-w-0 rounded-lg
        border border-slate-100
        bg-white/65
        px-2.5 py-2
        sm:flex-1
      "
                        >
                          <p className="text-[9px] font-bold uppercase tracking-wide text-slate-400">
                            Customer ID
                          </p>

                          <p className="mt-0.5 truncate text-[11px] font-semibold text-slate-700 sm:text-xs">
                            {transaction.customerRef || 'N/A'}
                          </p>
                        </div>

                        {/* Recorded By */}
                        <div
                          className="
        min-w-0 rounded-lg
        border border-slate-100
        bg-white/65
        px-2.5 py-2
        sm:flex-1
      "
                        >
                          <p className="text-[9px] font-bold uppercase tracking-wide text-slate-400">
                            Recorded By
                          </p>

                          <p className="mt-0.5 truncate text-[11px] font-semibold text-slate-700 sm:text-xs">
                            {transaction.userName ||
                              transaction.username ||
                              'N/A'}
                          </p>
                        </div>

                        {/* Role */}
                        <div
                          className="
        min-w-0 rounded-lg
        border border-slate-100
        bg-white/65
        px-2.5 py-2
        sm:flex-1
      "
                        >
                          <p className="text-[9px] font-bold uppercase tracking-wide text-slate-400">
                            Role
                          </p>

                          <p className="mt-0.5 truncate text-[11px] font-semibold capitalize text-slate-700 sm:text-xs">
                            {transaction.userRole || 'N/A'}
                          </p>
                        </div>

                        {/* Amount */}
                        <div
                          className="
        shrink-0
        rounded-lg
        border border-emerald-100
        bg-emerald-50/50
        px-3 py-2
        text-left
        sm:min-w-[125px]
        sm:text-right
      "
                        >
                          <p
                            className="
          text-sm font-extrabold tracking-tight
          text-emerald-600
          sm:text-base
        "
                          >
                            +
                            {formatMoney(Math.abs(signedAmount))}
                          </p>

                          <p className="mt-0.5 text-[9px] font-medium uppercase tracking-wide text-slate-400">
                            Amount
                          </p>
                        </div>
                      </div>
                    ) : (
                      /* --------------------------------
                         REFUND TRANSACTION - EXISTING LAYOUT
                      -------------------------------- */
                      <>
                        {/* Main Transaction Row */}
                        <div
                          className="
    relative flex flex-col gap-3
    px-3 py-2.5
    sm:flex-row sm:items-center sm:gap-3
    sm:px-4 sm:py-3
  "
                        >
                          {/* Category + Sale/Refund + Date/Time */}
                          <div className="flex min-w-0 shrink-0 items-center gap-2.5 sm:w-[22%]">
                            {/* Category Icon */}
                            <div
                              className={`
        flex h-9 w-9 shrink-0 items-center justify-center
        rounded-lg border shadow-sm
        ${isRefund
                                  ? 'border-red-100 bg-red-50 text-red-600'
                                  : 'border-blue-100 bg-blue-50 text-blue-600'
                                }
      `}
                            >
                              <CategoryIcon
                                category={transaction.category}
                                size={17}
                              />
                            </div>

                            <div className="min-w-0">
                              <div className="flex flex-wrap items-center gap-1.5">
                                <p className="truncate text-xs font-bold text-slate-800 sm:text-sm">
                                  {getCategoryName(transaction.category)}
                                </p>

                                <span
                                  className={`
            rounded-full
            px-1.5 py-0.5
            text-[9px] font-bold tracking-wide
            ${isRefund
                                      ? 'bg-red-100 text-red-700'
                                      : 'bg-emerald-100 text-emerald-700'
                                    }
          `}
                                >
                                  {isRefund ? 'REFUND' : 'SALE'}
                                </span>
                              </div>

                              <p className="mt-0.5 whitespace-nowrap text-[10px] font-medium text-slate-400 sm:text-[11px]">
                                {transaction.date
                                  ? formatDate(transaction.date)
                                  : '-'}
                                {' · '}
                                {transaction.time
                                  ? formatTime(transaction.time)
                                  : '-'}
                              </p>
                            </div>
                          </div>

                          {/* Customer ID */}
                          <div
                            className="
      min-w-0 rounded-lg
      border border-slate-100
      bg-white/65
      px-2.5 py-2
      sm:flex-1
    "
                          >
                            <p className="text-[9px] font-bold uppercase tracking-wide text-slate-400">
                              Customer ID
                            </p>

                            <p className="mt-0.5 truncate text-[11px] font-semibold text-slate-700 sm:text-xs">
                              {transaction.customerRef || 'N/A'}
                            </p>
                          </div>

                          {/* Recorded By */}
                          <div
                            className="
      min-w-0 rounded-lg
      border border-slate-100
      bg-white/65
      px-2.5 py-2
      sm:flex-1
    "
                          >
                            <p className="text-[9px] font-bold uppercase tracking-wide text-slate-400">
                              Recorded By
                            </p>

                            <p className="mt-0.5 truncate text-[11px] font-semibold text-slate-700 sm:text-xs">
                              {transaction.userName ||
                                transaction.username ||
                                'N/A'}
                            </p>
                          </div>

                          {/* Role */}
                          <div
                            className="
      min-w-0 rounded-lg
      border border-slate-100
      bg-white/65
      px-2.5 py-2
      sm:flex-1
    "
                          >
                            <p className="text-[9px] font-bold uppercase tracking-wide text-slate-400">
                              Role
                            </p>

                            <p className="mt-0.5 truncate text-[11px] font-semibold capitalize text-slate-700 sm:text-xs">
                              {transaction.userRole || 'N/A'}
                            </p>
                          </div>

                          {/* Amount */}
                          <div
                            className={`
      shrink-0 rounded-lg
      border px-3 py-2
      text-left
      sm:min-w-[125px]
      sm:text-right
      ${isRefund
                                ? 'border-red-100 bg-red-50/60'
                                : 'border-emerald-100 bg-emerald-50/50'
                              }
    `}
                          >
                            <p
                              className={`
        text-sm font-extrabold tracking-tight
        sm:text-base
        ${isRefund ? 'text-red-600' : 'text-emerald-600'}
      `}
                            >
                              {signedAmount < 0 ? '-' : '+'}
                              {formatMoney(Math.abs(signedAmount))}
                            </p>

                            <p className="mt-0.5 text-[9px] font-medium uppercase tracking-wide text-slate-400">
                              Amount
                            </p>
                          </div>
                        </div>
                      </>
                    )}

                    {/* Refund Details */}
                    {(transaction.reason ||
                      transaction.type === 'refund') && (
                        <div
                          className="
        relative
        border-t border-slate-100/80
        px-3 py-2.5
        sm:px-4 sm:py-3
      "
                        >
                          {/* Original Sale */}
                          {transaction.type === 'refund' && (
                            <div>
                              <div className="mb-2 flex items-center gap-2">
                                <div className="h-px flex-1 bg-slate-100" />

                                <p className="shrink-0 text-[9px] font-bold uppercase tracking-[0.12em] text-slate-400">
                                  Original Sale
                                </p>

                                <div className="h-px flex-1 bg-slate-100" />
                              </div>

                              {originalSale ? (
                                <div
                                  className="
                grid grid-cols-1 gap-2
                sm:grid-cols-2
                lg:grid-cols-5
              "
                                >
                                  {/* Date */}
                                  <div
                                    className="
                  rounded-lg
                  border border-slate-100
                  bg-white/65
                  px-2.5 py-2
                "
                                  >
                                    <p className="text-[9px] font-semibold uppercase tracking-wide text-slate-400">
                                      Date
                                    </p>

                                    <p className="mt-0.5 text-[11px] font-medium text-slate-700 sm:text-xs">
                                      {originalSale.date
                                        ? formatDate(originalSale.date)
                                        : 'N/A'}
                                    </p>
                                  </div>

                                  {/* Time */}
                                  <div
                                    className="
                  rounded-lg
                  border border-slate-100
                  bg-white/65
                  px-2.5 py-2
                "
                                  >
                                    <p className="text-[9px] font-semibold uppercase tracking-wide text-slate-400">
                                      Time
                                    </p>

                                    <p className="mt-0.5 text-[11px] font-medium text-slate-700 sm:text-xs">
                                      {originalSale.time
                                        ? formatTime(originalSale.time)
                                        : 'N/A'}
                                    </p>
                                  </div>

                                  {/* Category */}
                                  <div
                                    className="
                  rounded-lg
                  border border-slate-100
                  bg-white/65
                  px-2.5 py-2
                "
                                  >
                                    <p className="text-[9px] font-semibold uppercase tracking-wide text-slate-400">
                                      Category
                                    </p>

                                    <p className="mt-0.5 truncate text-[11px] font-medium text-slate-700 sm:text-xs">
                                      {getCategoryName(originalSale.category)}
                                    </p>
                                  </div>

                                  {/* Amount */}
                                  <div
                                    className="
                  rounded-lg
                  border border-slate-100
                  bg-white/65
                  px-2.5 py-2
                "
                                  >
                                    <p className="text-[9px] font-semibold uppercase tracking-wide text-slate-400">
                                      Amount
                                    </p>

                                    <p className="mt-0.5 text-[11px] font-bold text-slate-800 sm:text-xs">
                                      {formatMoney(originalSale.amount)}
                                    </p>
                                  </div>

                                  {/* Customer ID */}
                                  <div
                                    className="
                  rounded-lg
                  border border-slate-100
                  bg-white/65
                  px-2.5 py-2
                "
                                  >
                                    <p className="text-[9px] font-semibold uppercase tracking-wide text-slate-400">
                                      Customer ID
                                    </p>

                                    <p className="mt-0.5 truncate text-[11px] font-medium text-slate-700 sm:text-xs">
                                      {originalSale.customerRef ||
                                        transaction.customerRef ||
                                        'N/A'}
                                    </p>
                                  </div>
                                </div>
                              ) : (
                                <div
                                  className="
                rounded-lg
                border border-amber-100
                bg-amber-50/60
                px-3 py-2
              "
                                >
                                  <p className="text-xs text-amber-700">
                                    {transaction.originalSaleId
                                      ? `Sale reference: ${transaction.originalSaleId}`
                                      : 'Original sale record unavailable.'}
                                  </p>
                                </div>
                              )}
                            </div>
                          )}

                          {/* Refund Reason */}
                          <div
                            className="
          mt-2.5
          rounded-lg
          border border-slate-100
          bg-slate-50/55
          px-3 py-2.5
        "
                          >
                            <p className="text-[9px] font-bold uppercase tracking-wide text-slate-400">
                              Refund Reason
                            </p>

                            <p className="mt-0.5 text-xs leading-5 text-slate-700">
                              {transaction.reason ||
                                'No reason provided.'}
                            </p>
                          </div>
                        </div>
                      )}
                  </div>
                )
              })}
            </div>

            {/* Pagination */}
            {
              totalTransactionPages > 1 && (
                <div
                  className="
            relative flex items-center justify-between
            border-t border-blue-100/70
            px-4 py-3
            sm:px-5
          "
                >
                  <p className="text-[10px] font-medium text-slate-500 sm:text-xs">
                    Page {transactionPage} of{' '}
                    {totalTransactionPages}
                  </p>

                  <div className="flex items-center gap-2">
                    <button
                      type="button"
                      onClick={() =>
                        setTransactionPage((page) =>
                          Math.max(1, page - 1)
                        )
                      }
                      disabled={transactionPage === 1}
                      className="
                flex h-8 w-8 items-center justify-center
                rounded-lg
                border border-slate-200
                bg-white/80
                text-slate-600
                shadow-sm
                transition-all duration-200
                hover:border-blue-200
                hover:bg-blue-50
                hover:text-blue-600
                disabled:cursor-not-allowed
                disabled:opacity-40
              "
                    >
                      <svg
                        className="h-3.5 w-3.5"
                        viewBox="0 0 24 24"
                        fill="none"
                        stroke="currentColor"
                        strokeWidth="2"
                        strokeLinecap="round"
                        strokeLinejoin="round"
                      >
                        <path d="m15 18-6-6 6-6" />
                      </svg>
                    </button>

                    <button
                      type="button"
                      onClick={() =>
                        setTransactionPage((page) =>
                          Math.min(
                            totalTransactionPages,
                            page + 1
                          )
                        )
                      }
                      disabled={
                        transactionPage ===
                        totalTransactionPages
                      }
                      className="
                flex h-8 w-8 items-center justify-center
                rounded-lg
                border border-slate-200
                bg-white/80
                text-slate-600
                shadow-sm
                transition-all duration-200
                hover:border-blue-200
                hover:bg-blue-50
                hover:text-blue-600
                disabled:cursor-not-allowed
                disabled:opacity-40
              "
                    >
                      <svg
                        className="h-3.5 w-3.5"
                        viewBox="0 0 24 24"
                        fill="none"
                        stroke="currentColor"
                        strokeWidth="2"
                        strokeLinecap="round"
                        strokeLinejoin="round"
                      >
                        <path d="m9 18 6-6-6-6" />
                      </svg>
                    </button>
                  </div>
                </div>
              )
            }
          </>
        )}
      </section>
    </PageContainer>
  )
}


/* --------------------------------
   History Category
-------------------------------- */
function HistoryCategory({
  category,
  amount,
}) {
  const categoryStyles = {
    medicine: {
      container:
        'border-blue-100/80 bg-gradient-to-br from-white/85 to-blue-50/55 hover:border-blue-200 hover:shadow-[0_10px_24px_rgba(37,99,235,0.11)]',
      iconBox:
        'border-blue-100 bg-blue-50/80 text-blue-600',
      iconGlow:
        'bg-blue-100/60',
      amount:
        'text-blue-700',
    },

    general: {
      container:
        'border-indigo-100/80 bg-gradient-to-br from-white/85 to-indigo-50/50 hover:border-indigo-200 hover:shadow-[0_10px_24px_rgba(79,70,229,0.10)]',
      iconBox:
        'border-indigo-100 bg-indigo-50/80 text-indigo-600',
      iconGlow:
        'bg-indigo-100/60',
      amount:
        'text-indigo-700',
    },

    dispensing: {
      container:
        'border-cyan-100/80 bg-gradient-to-br from-white/85 to-cyan-50/50 hover:border-cyan-200 hover:shadow-[0_10px_24px_rgba(6,182,212,0.10)]',
      iconBox:
        'border-cyan-100 bg-cyan-50/80 text-cyan-600',
      iconGlow:
        'bg-cyan-100/60',
      amount:
        'text-cyan-700',
    },
  }

  const styles =
    categoryStyles[category] || categoryStyles.medicine

  return (
    <div
      className={`
        group relative overflow-hidden rounded-2xl
        border
        px-4 py-3.5
        shadow-[0_6px_16px_rgba(15,23,42,0.06)]
        backdrop-blur-md
        transition-all duration-200
        hover:-translate-y-0.5
        ${styles.container}
      `}
    >
      {/* Soft decorative glow */}
      <div
        className={`
          pointer-events-none absolute
          -right-5 -top-5
          h-20 w-20
          rounded-full
          opacity-70
          blur-2xl
          transition-transform duration-300
          group-hover:scale-125
          ${styles.iconGlow}
        `}
      />

      {/* Category Icon + Name */}
      <div className="relative flex items-start justify-between gap-3">
        <div className="flex min-w-0 items-center gap-3">
          {/* Main Category Icon */}
          <div
            className={`
              flex h-10 w-10 shrink-0
              items-center justify-center
              rounded-xl
              border
              shadow-[0_4px_10px_rgba(15,23,42,0.06)]
              transition-all duration-200
              group-hover:scale-105
              ${styles.iconBox}
            `}
          >
            <CategoryIcon
              category={category}
              size={20}
            />
          </div>

          <div className="min-w-0">
            <p className="text-[10px] font-bold uppercase tracking-[0.11em] text-slate-400">
              Category
            </p>

            <p className="mt-0.5 truncate text-sm font-bold text-slate-800">
              {getCategoryName(category)}
            </p>
          </div>
        </div>

        {/* Small Premium Icon */}
        <div
          className={`
            flex h-7 w-7 shrink-0
            items-center justify-center
            rounded-lg
            border
            bg-white/65
            shadow-sm
            backdrop-blur-sm
            transition-all duration-200
            group-hover:scale-110
            ${styles.iconBox}
          `}
        >
          <svg
            className="h-3.5 w-3.5"
            viewBox="0 0 24 24"
            fill="none"
            stroke="currentColor"
            strokeWidth="1.8"
            strokeLinecap="round"
            strokeLinejoin="round"
          >
            {category === 'medicine' ? (
              <>
                <path d="M10.5 20.5 3.5 13.5a4.95 4.95 0 0 1 7-7l7 7a4.95 4.95 0 0 1-7 7Z" />
                <path d="m8 8 8 8" />
              </>
            ) : category === 'general' ? (
              <>
                <path d="m21 8-9-5-9 5 9 5 9-5Z" />
                <path d="M3 8v8l9 5 9-5V8" />
              </>
            ) : (
              <>
                <path d="M8 3h8" />
                <path d="M9 3v4l-2 3v8a3 3 0 0 0 3 3h4a3 3 0 0 0 3-3v-8l-2-3V3" />
                <path d="M7 10h10" />
              </>
            )}
          </svg>
        </div>
      </div>

      {/* Amount */}
      <div className="relative mt-3 border-t border-slate-100/80 pt-2.5">
        <p className="text-[10px] font-semibold uppercase tracking-wide text-slate-400">
          Total Sales
        </p>

        <p
          className={`
            mt-0.5 text-lg font-extrabold
            tracking-tight
            ${styles.amount}
          `}
        >
          {formatMoney(amount)}
        </p>
      </div>
    </div>
  )
}
/* --------------------------------
   Password Requirement
-------------------------------- */
function PasswordRequirement({ met, text }) {
  return (
    <div
      className={`flex items-center gap-2 text-xs ${met
        ? 'text-emerald-600'
        : 'text-slate-500'
        }`}
    >
      <span
        className={`flex h-4 w-4 shrink-0 items-center justify-center rounded-full text-[10px] font-bold ${met
          ? 'bg-emerald-100 text-emerald-600'
          : 'bg-slate-200 text-slate-400'
          }`}
      >
        {met ? '✓' : '•'}
      </span>

      <span>{text}</span>
    </div>
  )
}

/* --------------------------------
   User Management
-------------------------------- */
function UserManagement({
  showNotification,
  currentUser,
}) {
  const isUserCurrentlyLoggedIn = (user) => {
    return (
      currentUser?.id &&
      user?.id &&
      currentUser.id === user.id
    )
  }
  const [users, setUsers] = useState([])
  const [showCreateUserModal, setShowCreateUserModal] =
    useState(false)

  const [userForm, setUserForm] = useState({
    name: '',
    username: '',
    password: '',
    confirmPassword: '',
  })

  const [showUserPassword, setShowUserPassword] = useState(false)
  const [showUserConfirmPassword, setShowUserConfirmPassword] = useState(false)

  const [isSavingUser, setIsSavingUser] =
    useState(false)

  /* --------------------------------
     Load Users
  -------------------------------- */
  useEffect(() => {
    async function loadUsers() {
      try {
        const data = await getUsers()
        setUsers(data || [])
      } catch (error) {
        console.error(
          'Failed to load users:',
          error
        )

        showNotification(
          'error',
          'Unable to load system users.'
        )
      }
    }

    loadUsers()
  }, [showNotification])

  /* --------------------------------
     Reset Create User Form
  -------------------------------- */
  function resetUserForm() {
    setUserForm({
      name: '',
      username: '',
      password: '',
      confirmPassword: '',
    })
    setShowUserPassword(false)
    setShowUserConfirmPassword(false)
  }

  /* --------------------------------
     Open Create User Modal
  -------------------------------- */
  function handleOpenCreateUser() {
    resetUserForm()
    setShowCreateUserModal(true)
  }

  /* --------------------------------
     Close Create User Modal
  -------------------------------- */
  function handleCloseCreateUser() {
    if (isSavingUser) return

    setShowCreateUserModal(false)
    resetUserForm()
  }

  /* --------------------------------
     Create User
  -------------------------------- */
  async function handleCreateUser() {
    const name = userForm.name.trim()
    const username =
      userForm.username.trim()
    const password = userForm.password
    const confirmPassword =
      userForm.confirmPassword

    if (!name) {
      showNotification(
        'error',
        'Please enter the staff/user name.'
      )
      return
    }

    if (!username) {
      showNotification(
        'error',
        'Please enter a username.'
      )
      return
    }

    if (!password) {
      showNotification(
        'error',
        'Please enter a password.'
      )
      return
    }

    const passwordRequirements = {
      minLength: password.length >= 8,
      uppercase: /[A-Z]/.test(password),
      lowercase: /[a-z]/.test(password),
      number: /[0-9]/.test(password),
      special: /[^A-Za-z0-9]/.test(password),
    }

    if (!passwordRequirements.minLength) {
      showNotification(
        'error',
        'Password must be at least 8 characters long.'
      )
      return
    }

    if (!passwordRequirements.uppercase) {
      showNotification(
        'error',
        'Password must contain at least one uppercase letter.'
      )
      return
    }

    if (!passwordRequirements.lowercase) {
      showNotification(
        'error',
        'Password must contain at least one lowercase letter.'
      )
      return
    }

    if (!passwordRequirements.number) {
      showNotification(
        'error',
        'Password must contain at least one number.'
      )
      return
    }

    if (!passwordRequirements.special) {
      showNotification(
        'error',
        'Password must contain at least one special character.'
      )
      return
    }

    if (!confirmPassword) {
      showNotification(
        'error',
        'Please confirm the password.'
      )
      return
    }

    if (password !== confirmPassword) {
      showNotification(
        'error',
        'Passwords do not match.'
      )
      return
    }

    try {
      setIsSavingUser(true)

      const newUser = await createUser({
        name,
        username,
        password,
        role: 'user',
      })

      setUsers((previousUsers) => [
        ...previousUsers,
        newUser,
      ])

      setShowCreateUserModal(false)
      resetUserForm()

      showNotification(
        'success',
        `User "${newUser.name}" was created successfully.`
      )
    } catch (error) {
      console.error(
        'Failed to create user:',
        error
      )

      if (
        error?.message ===
        'Username already exists.'
      ) {
        showNotification(
          'error',
          'This username already exists. Please choose another username.'
        )
      } else {
        showNotification(
          'error',
          error?.message ||
          'Unable to create user. Please try again.'
        )
      }
    } finally {
      setIsSavingUser(false)
    }
  }

  return (
    <>
      <section
        className={`${ui.section} overflow-hidden`}
      >
        <div className="p-6 sm:p-8">
          {/* Header */}
          <div className="flex flex-col gap-4 sm:flex-row sm:items-center sm:justify-between">
            <div>
              <div className="flex items-center gap-3">
                <div className="flex h-11 w-11 items-center justify-center rounded-xl bg-blue-50 text-blue-600 ring-1 ring-blue-100">
                  <NavIcon type="users" />
                </div>

                <div>
                  <h2 className="text-lg font-bold text-slate-900">
                    User Management
                  </h2>

                  <p className="mt-1 text-sm text-slate-500">
                    Manage system users and staff access.
                  </p>
                </div>
              </div>
            </div>

            <button
              type="button"
              onClick={handleOpenCreateUser}
              className={`${ui.primaryButton} whitespace-nowrap`}
            >
              + Create User
            </button>
          </div>

          {/* User List */}
          <div className="mt-6 overflow-hidden rounded-2xl border border-slate-200 bg-white">
            {users.length === 0 ? (
              <div className="p-8 text-center">
                <div className="mx-auto flex h-12 w-12 items-center justify-center rounded-xl bg-slate-100 text-slate-500">
                  <NavIcon type="users" />
                </div>

                <h3 className="mt-4 text-sm font-bold text-slate-900">
                  No users found
                </h3>

                <p className="mt-1 text-sm text-slate-500">
                  Create a staff account to get started.
                </p>
              </div>
            ) : (
              <>
                {/* Desktop / Tablet Table */}
                <div className="hidden overflow-x-auto sm:block">
                  <table className="w-full min-w-[620px]">
                    <thead>
                      <tr className={ui.tableHeader}>
                        <th className="px-5 py-3 text-left">
                          Name
                        </th>

                        <th className="px-5 py-3 text-left">
                          Username
                        </th>

                        <th className="px-5 py-3 text-left">
                          Role
                        </th>

                        <th className="px-5 py-3 text-left">
                          Status
                        </th>
                      </tr>
                    </thead>

                    <tbody className="divide-y divide-slate-100">
                      {users.map((user) => (
                        <tr
                          key={user.id}
                          className="transition hover:bg-blue-50/30"
                        >
                          <td className="px-5 py-4">
                            <div className="flex items-center gap-3">
                              <div className="flex h-9 w-9 shrink-0 items-center justify-center rounded-xl bg-blue-50 text-sm font-bold text-blue-600 ring-1 ring-blue-100">
                                {user.name
                                  ?.charAt(0)
                                  ?.toUpperCase() || 'U'}
                              </div>

                              <div className="min-w-0">
                                <p className="truncate text-sm font-semibold text-slate-900">
                                  {user.name}
                                </p>
                              </div>
                            </div>
                          </td>

                          <td className="px-5 py-4">
                            <p className="text-sm text-slate-600">
                              @{user.username}
                            </p>
                          </td>

                          <td className="px-5 py-4">
                            <span
                              className={`inline-flex items-center gap-1.5 rounded-full px-2.5 py-1 text-[10px] font-bold uppercase tracking-wide ${user.role === 'admin'
                                ? isUserCurrentlyLoggedIn(user)
                                  ? 'border border-blue-300 bg-blue-100 text-blue-800 ring-2 ring-blue-200 shadow-sm'
                                  : 'border border-blue-200 bg-blue-50 text-blue-700'
                                : 'border border-slate-200 bg-slate-100 text-slate-600'
                                }`}
                            >
                              {user.role === 'admin' && (
                                <span
                                  className={`h-1.5 w-1.5 rounded-full ${isUserCurrentlyLoggedIn(user)
                                    ? 'bg-emerald-500'
                                    : 'bg-blue-400'
                                    }`}
                                />
                              )}

                              {user.role === 'admin'
                                ? 'Admin'
                                : 'User'}
                            </span>
                          </td>

                          <td className="px-5 py-4">
                            <span
                              className={`inline-flex items-center gap-1.5 rounded-full px-2.5 py-1 text-xs font-semibold ${isUserCurrentlyLoggedIn(user)
                                ? 'bg-emerald-50 text-emerald-700 ring-1 ring-emerald-200'
                                : 'bg-red-50 text-red-600 ring-1 ring-red-200'
                                }`}
                            >
                              <span
                                className={`h-1.5 w-1.5 rounded-full ${isUserCurrentlyLoggedIn(user)
                                  ? 'bg-emerald-500'
                                  : 'bg-red-500'
                                  }`}
                              />

                              {isUserCurrentlyLoggedIn(user)
                                ? 'Active'
                                : 'Inactive'}
                            </span>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>

                {/* Mobile User Cards */}
                <div className="divide-y divide-slate-100 sm:hidden">
                  {users.map((user) => (
                    <div
                      key={user.id}
                      className="p-4"
                    >
                      <div className="flex items-start gap-3">
                        <div className="flex h-10 w-10 shrink-0 items-center justify-center rounded-xl bg-blue-50 text-sm font-bold text-blue-600 ring-1 ring-blue-100">
                          {user.name
                            ?.charAt(0)
                            ?.toUpperCase() || 'U'}
                        </div>

                        <div className="min-w-0 flex-1">
                          <div className="flex items-start justify-between gap-3">
                            <div className="min-w-0">
                              <p className="truncate text-sm font-bold text-slate-900">
                                {user.name}
                              </p>

                              <p className="mt-0.5 truncate text-xs text-slate-500">
                                @{user.username}
                              </p>
                            </div>

                            <span
                              className={`shrink-0 inline-flex items-center gap-1.5 rounded-full border px-2.5 py-1 text-[10px] font-bold uppercase tracking-wide ${user.role === 'admin'
                                ? isUserCurrentlyLoggedIn(user)
                                  ? 'border-blue-300 bg-blue-100 text-blue-800 ring-2 ring-blue-200 shadow-sm'
                                  : 'border-blue-200 bg-blue-50 text-blue-700'
                                : 'border-slate-200 bg-slate-100 text-slate-600'
                                }`}
                            >
                              {user.role === 'admin' && (
                                <span
                                  className={`h-1.5 w-1.5 rounded-full ${isUserCurrentlyLoggedIn(user)
                                    ? 'bg-emerald-500'
                                    : 'bg-blue-400'
                                    }`}
                                />
                              )}

                              {user.role === 'admin'
                                ? 'Admin'
                                : 'User'}
                            </span>
                          </div>

                          <div
                            className={`mt-3 flex items-center gap-1.5 text-xs font-semibold ${isUserCurrentlyLoggedIn(user)
                              ? 'text-emerald-600'
                              : 'text-red-600'
                              }`}
                          >
                            <span
                              className={`h-1.5 w-1.5 rounded-full ${isUserCurrentlyLoggedIn(user)
                                ? 'bg-emerald-500'
                                : 'bg-red-500'
                                }`}
                            />

                            {isUserCurrentlyLoggedIn(user)
                              ? 'Active'
                              : 'Inactive'}
                          </div>
                        </div>
                      </div>
                    </div>
                  ))}
                </div>
              </>
            )}
          </div>
        </div>
      </section>

      {/* Create User Modal */}
      {showCreateUserModal && (
        <Modal
          title="Create User Account"
          onClose={handleCloseCreateUser}
        >
          <div className="space-y-5">
            {/* Name */}
            <div>
              <label className={ui.label}>
                Staff/User Name
              </label>

              <input
                type="text"
                value={userForm.name}
                onChange={(e) =>
                  setUserForm((previous) => ({
                    ...previous,
                    name: e.target.value,
                  }))
                }
                placeholder="Enter staff name"
                className={ui.input}
                autoFocus
              />
            </div>

            {/* Username */}
            <div>
              <label className={ui.label}>
                Username
              </label>

              <input
                type="text"
                value={userForm.username}
                onChange={(e) =>
                  setUserForm((previous) => ({
                    ...previous,
                    username: e.target.value,
                  }))
                }
                placeholder="Enter username"
                className={ui.input}
                autoComplete="off"
              />
            </div>

            {/* Password */}
            <div>
              <label className={ui.label}>
                Password
              </label>

              <div className="relative mt-2">
                <input
                  type={showUserPassword ? 'text' : 'password'}
                  value={userForm.password}
                  onChange={(e) =>
                    setUserForm((previous) => ({
                      ...previous,
                      password: e.target.value,
                    }))
                  }
                  placeholder="Enter password"
                  className={`${ui.input} pr-12`}
                  autoComplete="new-password"
                />

                <button
                  type="button"
                  onClick={() =>
                    setShowUserPassword((value) => !value)
                  }
                  className="absolute right-3 top-1/2 flex h-8 w-8 -translate-y-1/2 items-center justify-center rounded-lg text-slate-400 transition hover:bg-blue-50 hover:text-blue-600 focus:outline-none focus:ring-2 focus:ring-blue-100"
                  aria-label={
                    showUserPassword
                      ? 'Hide password'
                      : 'Show password'
                  }
                >
                  {showUserPassword ? (
                    <svg
                      xmlns="http://www.w3.org/2000/svg"
                      viewBox="0 0 24 24"
                      fill="none"
                      stroke="currentColor"
                      strokeWidth="2"
                      className="h-5 w-5"
                    >
                      <path
                        strokeLinecap="round"
                        strokeLinejoin="round"
                        d="M3 3l18 18"
                      />
                      <path
                        strokeLinecap="round"
                        strokeLinejoin="round"
                        d="M10.58 10.58a2 2 0 0 0 2.83 2.83"
                      />
                      <path
                        strokeLinecap="round"
                        strokeLinejoin="round"
                        d="M9.88 4.24A10.94 10.94 0 0 1 12 4c5 0 8.5 4 10 8a16.6 16.6 0 0 1-3.03 4.68"
                      />
                      <path
                        strokeLinecap="round"
                        strokeLinejoin="round"
                        d="M6.61 6.61C4.62 7.88 3.25 9.72 2 12c1.5 4 5 8 10 8a10.94 10.94 0 0 0 4.12-.76"
                      />
                    </svg>
                  ) : (
                    <svg
                      xmlns="http://www.w3.org/2000/svg"
                      viewBox="0 0 24 24"
                      fill="none"
                      stroke="currentColor"
                      strokeWidth="2"
                      className="h-5 w-5"
                    >
                      <path
                        strokeLinecap="round"
                        strokeLinejoin="round"
                        d="M2 12s3.5-7 10-7 10 7 10 7-3.5 7-10 7S2 12 2 12Z"
                      />
                      <circle cx="12" cy="12" r="3" />
                    </svg>
                  )}
                </button>
              </div>

              <div className="mt-2 rounded-xl border border-slate-200 bg-slate-50/70 px-3 py-2.5">
                <p className="mb-2 text-[11px] font-semibold text-slate-600">
                  Password must contain:
                </p>

                <div className="grid grid-cols-1 gap-x-4 gap-y-1 sm:grid-cols-2">
                  <PasswordRequirement
                    met={userForm.password.length >= 8}
                    text="8+ characters"
                  />

                  <PasswordRequirement
                    met={/[A-Z]/.test(userForm.password)}
                    text="Uppercase letter"
                  />

                  <PasswordRequirement
                    met={/[a-z]/.test(userForm.password)}
                    text="Lowercase letter"
                  />

                  <PasswordRequirement
                    met={/[0-9]/.test(userForm.password)}
                    text="Number"
                  />

                  <PasswordRequirement
                    met={/[^A-Za-z0-9]/.test(userForm.password)}
                    text="Special character"
                  />
                </div>
              </div>
            </div>

            {/* Confirm Password */}
            <div>
              <label className={ui.label}>
                Confirm Password
              </label>

              <div className="relative mt-2">
                <input
                  type={
                    showUserConfirmPassword
                      ? 'text'
                      : 'password'
                  }
                  value={userForm.confirmPassword}
                  onChange={(e) =>
                    setUserForm((previous) => ({
                      ...previous,
                      confirmPassword: e.target.value,
                    }))
                  }
                  placeholder="Confirm password"
                  className={`${ui.input} pr-12`}
                  autoComplete="new-password"
                />

                <button
                  type="button"
                  onClick={() =>
                    setShowUserConfirmPassword(
                      (value) => !value
                    )
                  }
                  className="absolute right-3 top-1/2 flex h-8 w-8 -translate-y-1/2 items-center justify-center rounded-lg text-slate-400 transition hover:bg-blue-50 hover:text-blue-600 focus:outline-none focus:ring-2 focus:ring-blue-100"
                  aria-label={
                    showUserConfirmPassword
                      ? 'Hide confirm password'
                      : 'Show confirm password'
                  }
                >
                  {showUserConfirmPassword ? (
                    <svg
                      xmlns="http://www.w3.org/2000/svg"
                      viewBox="0 0 24 24"
                      fill="none"
                      stroke="currentColor"
                      strokeWidth="2"
                      className="h-5 w-5"
                    >
                      <path
                        strokeLinecap="round"
                        strokeLinejoin="round"
                        d="M3 3l18 18"
                      />
                      <path
                        strokeLinecap="round"
                        strokeLinejoin="round"
                        d="M10.58 10.58a2 2 0 0 0 2.83 2.83"
                      />
                      <path
                        strokeLinecap="round"
                        strokeLinejoin="round"
                        d="M9.88 4.24A10.94 10.94 0 0 1 12 4c5 0 8.5 4 10 8a16.6 16.6 0 0 1-3.03 4.68"
                      />
                      <path
                        strokeLinecap="round"
                        strokeLinejoin="round"
                        d="M6.61 6.61C4.62 7.88 3.25 9.72 2 12c1.5 4 5 8 10 8a10.94 10.94 0 0 0 4.12-.76"
                      />
                    </svg>
                  ) : (
                    <svg
                      xmlns="http://www.w3.org/2000/svg"
                      viewBox="0 0 24 24"
                      fill="none"
                      stroke="currentColor"
                      strokeWidth="2"
                      className="h-5 w-5"
                    >
                      <path
                        strokeLinecap="round"
                        strokeLinejoin="round"
                        d="M2 12s3.5-7 10-7 10 7 10 7-3.5 7-10 7S2 12 2 12Z"
                      />
                      <circle cx="12" cy="12" r="3" />
                    </svg>
                  )}
                </button>
              </div>
            </div>

            {/* Role */}
            <div>
              <label className={ui.label}>
                Role
              </label>

              <div className="flex items-center justify-between rounded-xl border border-slate-200 bg-slate-50 px-4 py-3">
                <div>
                  <p className="text-sm font-semibold text-slate-900">
                    User
                  </p>

                  <p className="mt-0.5 text-xs text-slate-500">
                    Staff access
                  </p>
                </div>

                <span className="inline-flex items-center rounded-full border border-slate-200 bg-slate-100 px-2.5 py-1 text-[10px] font-bold uppercase tracking-wide text-slate-600">
                  USER
                </span>
              </div>
            </div>

            {/* Actions */}
            <div className="flex flex-col-reverse gap-3 border-t border-slate-200 pt-5 sm:flex-row sm:justify-end">
              <button
                type="button"
                onClick={handleCloseCreateUser}
                disabled={isSavingUser}
                className={`${ui.secondaryButton} w-full sm:w-auto`}
              >
                Cancel
              </button>

              <button
                type="button"
                onClick={handleCreateUser}
                disabled={isSavingUser}
                className={`${ui.primaryButton} w-full sm:w-auto`}
              >
                {isSavingUser
                  ? 'Creating User...'
                  : 'Create User'}
              </button>
            </div>
          </div>
        </Modal>
      )}
    </>
  )
}