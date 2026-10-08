const DB_NAME = 'MedicalStoreDB'
const DB_VERSION = 5

const BACKUP_FORMAT_VERSION = 1
const BACKUP_FILE_EXTENSION = '.kmsbackup'
const BACKUP_APP_NAME = 'Khalid Medical Store'

const BACKUP_ENCRYPTION_ALGORITHM = 'AES-GCM'
const BACKUP_KEY_LENGTH = 256
const BACKUP_SALT_LENGTH = 16
const BACKUP_IV_LENGTH = 12
const BACKUP_PBKDF2_ITERATIONS = 250000

const STORES = {
    DAILY_SESSIONS: 'dailySessions',
    TRANSACTIONS: 'transactions',
    APP_STATE: 'appState',
    USERS: 'users',
    AUDIT_LOGS: 'auditLogs',
    BACKUP_HISTORY: 'backupHistory',
}

const DEFAULT_USERS = [
    {
        id: 'admin-user',
        username: 'admin',
        password: 'admin123',
        name: 'Administrator',
        role: 'admin',
    },
    {
        id: 'staff-user',
        username: 'user',
        password: 'user123',
        name: 'Staff User',
        role: 'user',
    },
]

function openDB() {
    return new Promise((resolve, reject) => {
        const request = indexedDB.open(DB_NAME, DB_VERSION)

        request.onupgradeneeded = (event) => {
            const db = event.target.result

            // Daily Sessions Store
            if (!db.objectStoreNames.contains(STORES.DAILY_SESSIONS)) {
                const dailySessions =
                    db.createObjectStore(
                        STORES.DAILY_SESSIONS,
                        { keyPath: 'id' }
                    )

                dailySessions.createIndex(
                    'date',
                    'date',
                    { unique: false }
                )

                dailySessions.createIndex(
                    'status',
                    'status',
                    { unique: false }
                )
            }
            if (!db.objectStoreNames.contains(STORES.BACKUP_HISTORY)) {
                db.createObjectStore(STORES.BACKUP_HISTORY, {
                    keyPath: 'id',
                    autoIncrement: true,
                })
            }

            if (!db.objectStoreNames.contains(STORES.AUDIT_LOGS)) {
                const auditLogs = db.createObjectStore(
                    STORES.AUDIT_LOGS,
                    { keyPath: 'id' }
                )

                auditLogs.createIndex('timestamp', 'timestamp', {
                    unique: false,
                })

                auditLogs.createIndex('userId', 'userId', {
                    unique: false,
                })

                auditLogs.createIndex('action', 'action', {
                    unique: false,
                })

                auditLogs.createIndex('sessionId', 'sessionId', {
                    unique: false,
                })
            }

            // Transactions Store
            if (!db.objectStoreNames.contains(STORES.TRANSACTIONS)) {
                const transactions = db.createObjectStore(
                    STORES.TRANSACTIONS,
                    { keyPath: 'id' }
                )

                transactions.createIndex('sessionId', 'sessionId', {
                    unique: false,
                })

                transactions.createIndex('type', 'type', {
                    unique: false,
                })

                transactions.createIndex('category', 'category', {
                    unique: false,
                })
            }

            // App State Store
            if (!db.objectStoreNames.contains(STORES.APP_STATE)) {
                db.createObjectStore(STORES.APP_STATE, {
                    keyPath: 'key',
                })
            }

            // Users Store
            if (!db.objectStoreNames.contains(STORES.USERS)) {
                db.createObjectStore(STORES.USERS, {
                    keyPath: 'id',
                })
            }
        }

        request.onsuccess = () => {
            resolve(request.result)
        }

        request.onerror = () => {
            reject(request.error)
        }
    })
}

/* =========================================================
   LEGACY DEFAULT USER MIGRATION
========================================================= */

async function migrateLegacyDefaultUsers() {
    const db = await openDB()

    return new Promise((resolve, reject) => {
        const transaction = db.transaction(
            STORES.USERS,
            'readwrite'
        )

        const store = transaction.objectStore(
            STORES.USERS
        )

        const request = store.getAll()

        request.onsuccess = () => {
            const users = request.result || []

            const legacyUserIds = [
                'admin-user',
                'staff-user',
            ]

            const legacyUsers = users.filter(
                (user) =>
                    legacyUserIds.includes(user.id)
            )

            if (legacyUsers.length === 0) {
                return
            }

            legacyUsers.forEach((user) => {
                store.delete(user.id)
            })
        }

        request.onerror = () => {
            db.close()
            reject(request.error)
        }

        transaction.oncomplete = () => {
            db.close()
            resolve()
        }

        transaction.onerror = () => {
            db.close()
            reject(transaction.error)
        }
    })
}

/* =========================================================
   USER MANAGEMENT
========================================================= */

export async function createUser({
    name,
    username,
    password,
    role = 'user',
}) {
    const db = await openDB()

    const cleanName = name.trim()
    const cleanUsername = username.trim().toLowerCase()

    if (!cleanName) {
        db.close()
        throw new Error('Name is required.')
    }

    if (!cleanUsername) {
        db.close()
        throw new Error('Username is required.')
    }

    if (!password) {
        db.close()
        throw new Error('Password is required.')
    }

    if (!['admin', 'user'].includes(role)) {
        db.close()
        throw new Error('Invalid user role.')
    }

    return new Promise((resolve, reject) => {
        const transaction = db.transaction(
            STORES.USERS,
            'readwrite'
        )

        const store = transaction.objectStore(
            STORES.USERS
        )

        const request = store.getAll()

        request.onsuccess = () => {
            const users = request.result || []

            const usernameExists = users.some(
                (user) =>
                    user.username?.trim().toLowerCase() ===
                    cleanUsername
            )

            if (usernameExists) {
                db.close()
                reject(
                    new Error(
                        'Username already exists.'
                    )
                )
                return
            }

            const newUser = {
                id: crypto.randomUUID(),
                name: cleanName,
                username: cleanUsername,
                password,
                role,
                createdAt: new Date().toISOString(),
            }

            store.put(newUser)

            transaction.oncomplete = () => {
                db.close()

                const {
                    password: _password,
                    ...safeUser
                } = newUser

                resolve(safeUser)
            }
        }

        request.onerror = () => {
            db.close()
            reject(request.error)
        }

        transaction.onerror = () => {
            db.close()
            reject(transaction.error)
        }
    })
}


/* =========================================================
   GET USERS
========================================================= */

export async function getUsers() {
    const db = await openDB()

    return new Promise((resolve, reject) => {
        const transaction = db.transaction(
            STORES.USERS,
            'readonly'
        )

        const store = transaction.objectStore(
            STORES.USERS
        )

        const request = store.getAll()

        request.onsuccess = () => {
            db.close()

            const users = (request.result || []).map(
                (user) => {
                    const {
                        password: _password,
                        ...safeUser
                    } = user

                    return safeUser
                }
            )

            resolve(users)
        }

        request.onerror = () => {
            db.close()
            reject(request.error)
        }
    })
}


/* =========================================================
   CHECK REGISTERED USERS
========================================================= */

export async function hasRegisteredUsers() {
    await migrateLegacyDefaultUsers()

    const db = await openDB()

    return new Promise((resolve, reject) => {
        const transaction = db.transaction(
            STORES.USERS,
            'readonly'
        )

        const store = transaction.objectStore(
            STORES.USERS
        )

        const request = store.count()

        request.onsuccess = () => {
            db.close()
            resolve(request.result > 0)
        }

        request.onerror = () => {
            db.close()
            reject(request.error)
        }
    })
}


/* =========================================================
   USER AUTHENTICATION
========================================================= */

export async function authenticateUser(
    username,
    password
) {
    const db = await openDB()

    const cleanUsername =
        username.trim().toLowerCase()

    return new Promise((resolve, reject) => {
        const transaction = db.transaction(
            STORES.USERS,
            'readonly'
        )

        const store = transaction.objectStore(
            STORES.USERS
        )

        const request = store.getAll()

        request.onsuccess = () => {
            db.close()

            const user = (request.result || []).find(
                (item) =>
                    item.username?.trim().toLowerCase() ===
                    cleanUsername &&
                    item.password === password
            )

            if (!user) {
                resolve(null)
                return
            }

            const {
                password: _password,
                ...safeUser
            } = user

            resolve(safeUser)
        }

        request.onerror = () => {
            db.close()
            reject(request.error)
        }
    })
}

/* =========================================================
   ACTIVE USER
========================================================= */

export async function saveActiveUser(user) {
    const db = await openDB()

    return new Promise((resolve, reject) => {
        const transaction = db.transaction(
            STORES.APP_STATE,
            'readwrite'
        )

        const store = transaction.objectStore(
            STORES.APP_STATE
        )

        store.put({
            key: 'activeUser',
            ...user,
        })

        transaction.oncomplete = () => {
            db.close()
            resolve()
        }

        transaction.onerror = () => {
            db.close()
            reject(transaction.error)
        }
    })
}

export async function getActiveUser() {
    const db = await openDB()

    return new Promise((resolve, reject) => {
        const transaction = db.transaction(
            STORES.APP_STATE,
            'readonly'
        )

        const store = transaction.objectStore(
            STORES.APP_STATE
        )

        const request = store.get('activeUser')

        request.onsuccess = () => {
            db.close()
            resolve(request.result || null)
        }

        request.onerror = () => {
            db.close()
            reject(request.error)
        }
    })
}

export async function clearActiveUser() {
    const db = await openDB()

    return new Promise((resolve, reject) => {
        const transaction = db.transaction(
            STORES.APP_STATE,
            'readwrite'
        )

        const store = transaction.objectStore(
            STORES.APP_STATE
        )

        store.delete('activeUser')

        transaction.oncomplete = () => {
            db.close()
            resolve()
        }

        transaction.onerror = () => {
            db.close()
            reject(transaction.error)
        }
    })
}

/* =========================================================
   CURRENT SESSION
========================================================= */

export async function saveCurrentSession(session) {
    const db = await openDB()

    return new Promise((resolve, reject) => {
        const transaction = db.transaction(
            STORES.APP_STATE,
            'readwrite'
        )

        const store = transaction.objectStore(
            STORES.APP_STATE
        )

        store.put({
            key: 'currentSession',
            ...session,
        })

        transaction.oncomplete = () => {
            db.close()
            resolve()
        }

        transaction.onerror = () => {
            db.close()
            reject(transaction.error)
        }
    })
}

export async function getCurrentSession() {
    const db = await openDB()

    return new Promise((resolve, reject) => {
        const transaction = db.transaction(
            STORES.APP_STATE,
            'readonly'
        )

        const store = transaction.objectStore(
            STORES.APP_STATE
        )

        const request = store.get('currentSession')

        request.onsuccess = () => {
            db.close()
            resolve(request.result || null)
        }

        request.onerror = () => {
            db.close()
            reject(request.error)
        }
    })
}

export async function clearCurrentSession() {
    const db = await openDB()

    return new Promise((resolve, reject) => {
        const transaction = db.transaction(
            STORES.APP_STATE,
            'readwrite'
        )

        const store = transaction.objectStore(
            STORES.APP_STATE
        )

        store.delete('currentSession')

        transaction.oncomplete = () => {
            db.close()
            resolve()
        }

        transaction.onerror = () => {
            db.close()
            reject(transaction.error)
        }
    })
}

/* =========================================================
   TRANSACTIONS
========================================================= */

export async function saveTransaction(transactionData) {
    const db = await openDB()

    return new Promise((resolve, reject) => {
        const transaction = db.transaction(
            STORES.TRANSACTIONS,
            'readwrite'
        )

        const store = transaction.objectStore(
            STORES.TRANSACTIONS
        )

        store.put(transactionData)

        transaction.oncomplete = () => {
            db.close()
            resolve()
        }

        transaction.onerror = () => {
            db.close()
            reject(transaction.error)
        }
    })
}

/* --------------------------------
   Delete Transaction
-------------------------------- */
export async function deleteTransaction(
    transactionId
) {
    const db = await openDB()

    return new Promise((resolve, reject) => {
        const transaction = db.transaction(
            STORES.TRANSACTIONS,
            'readwrite'
        )

        const store = transaction.objectStore(
            STORES.TRANSACTIONS
        )

        store.delete(transactionId)

        transaction.oncomplete = () => {
            db.close()
            resolve()
        }

        transaction.onerror = () => {
            db.close()
            reject(transaction.error)
        }
    })
}

export async function getTransactionsBySession(sessionId) {
    const db = await openDB()

    return new Promise((resolve, reject) => {
        const transaction = db.transaction(
            STORES.TRANSACTIONS,
            'readonly'
        )

        const store = transaction.objectStore(
            STORES.TRANSACTIONS
        )

        const index = store.index('sessionId')

        const request = index.getAll(sessionId)

        request.onsuccess = () => {
            db.close()

            const records = request.result || []

            records.sort(
                (a, b) =>
                    new Date(b.time) -
                    new Date(a.time)
            )

            resolve(records)
        }

        request.onerror = () => {
            db.close()
            reject(request.error)
        }
    })
}

/* =========================================================
   DAILY SESSIONS
========================================================= */

export async function saveDailySession(record) {
    const db = await openDB()

    return new Promise((resolve, reject) => {
        const transaction = db.transaction(
            STORES.DAILY_SESSIONS,
            'readwrite'
        )

        const store = transaction.objectStore(
            STORES.DAILY_SESSIONS
        )

        store.put(record)

        transaction.oncomplete = () => {
            db.close()
            resolve()
        }

        transaction.onerror = () => {
            db.close()
            reject(transaction.error)
        }
    })
}

export async function getDailySessions() {
    const db = await openDB()

    return new Promise((resolve, reject) => {
        const transaction = db.transaction(
            STORES.DAILY_SESSIONS,
            'readonly'
        )

        const store = transaction.objectStore(
            STORES.DAILY_SESSIONS
        )

        const request = store.getAll()

        request.onsuccess = () => {
            db.close()

            const records = request.result || []

            records.sort((a, b) => {
                const dateA = new Date(
                    `${a.date}T00:00:00`
                )

                const dateB = new Date(
                    `${b.date}T00:00:00`
                )

                if (dateB - dateA !== 0) {
                    return dateB - dateA
                }

                return (
                    new Date(b.closedAt) -
                    new Date(a.closedAt)
                )
            })

            resolve(records)
        }

        request.onerror = () => {
            db.close()
            reject(request.error)
        }
    })
}

/* =========================================================
   COMPLETE DAILY SESSION
========================================================= */

export async function completeDailySession(record) {
    const db = await openDB()

    return new Promise((resolve, reject) => {
        const transaction = db.transaction(
            [
                STORES.DAILY_SESSIONS,
                STORES.APP_STATE,
            ],
            'readwrite'
        )

        const dailySessionsStore =
            transaction.objectStore(
                STORES.DAILY_SESSIONS
            )

        const appStateStore =
            transaction.objectStore(
                STORES.APP_STATE
            )

        dailySessionsStore.put(record)

        appStateStore.delete(
            'currentSession'
        )

        transaction.oncomplete = () => {
            db.close()
            resolve()
        }

        transaction.onerror = () => {
            db.close()
            reject(transaction.error)
        }
    })
}

/* =========================================================
   CLEAR DATABASE
========================================================= */

export async function clearDatabase() {
    return new Promise((resolve, reject) => {
        const request =
            indexedDB.deleteDatabase(DB_NAME)

        request.onsuccess = () => {
            resolve()
        }

        request.onerror = () => {
            reject(request.error)
        }

        request.onblocked = () => {
            reject(
                new Error(
                    'Database is currently in use. Close other app tabs first.'
                )
            )
        }
    })
}

/* --------------------------------
   Audit Logs
-------------------------------- */

export async function addAuditLog({
    userId = null,
    username = null,
    action,
    description,
    sessionId = null,
    transactionId = null,
}) {
    const db = await openDB()

    const auditLog = {
        id: crypto.randomUUID(),
        timestamp: new Date().toISOString(),
        userId,
        username,
        action,
        description,
        sessionId,
        transactionId,
    }

    return new Promise((resolve, reject) => {
        const transaction = db.transaction(
            STORES.AUDIT_LOGS,
            'readwrite'
        )

        const store = transaction.objectStore(
            STORES.AUDIT_LOGS
        )

        store.put(auditLog)

        transaction.oncomplete = () => {
            db.close()
            resolve(auditLog)
        }

        transaction.onerror = () => {
            db.close()
            reject(transaction.error)
        }
    })
}

export async function getAuditLogs() {
    const db = await openDB()

    return new Promise((resolve, reject) => {
        const transaction = db.transaction(
            STORES.AUDIT_LOGS,
            'readonly'
        )

        const store = transaction.objectStore(
            STORES.AUDIT_LOGS
        )

        const request = store.getAll()

        request.onsuccess = () => {
            db.close()

            const logs = request.result || []

            logs.sort(
                (a, b) =>
                    new Date(b.timestamp) -
                    new Date(a.timestamp)
            )

            resolve(logs)
        }

        request.onerror = () => {
            db.close()
            reject(request.error)
        }
    })
}

/* --------------------------------
   Monthly Customer ID
-------------------------------- */

export async function generateCustomerId() {
    const db = await openDB()

    const currentMonth = new Date()
        .toISOString()
        .slice(0, 7)

    return new Promise((resolve, reject) => {
        const transaction = db.transaction(
            STORES.APP_STATE,
            'readwrite'
        )

        const store = transaction.objectStore(
            STORES.APP_STATE
        )

        const request = store.get(
            'customerIdCounter'
        )

        request.onsuccess = () => {
            const existing = request.result

            let nextCount = 1

            if (
                existing &&
                existing.month === currentMonth
            ) {
                nextCount =
                    Number(existing.count || 0) + 1
            }

            const customerId =
                `CUST-${String(nextCount).padStart(3, '0')}`

            store.put({
                key: 'customerIdCounter',
                month: currentMonth,
                count: nextCount,
            })

            transaction.oncomplete = () => {
                db.close()
                resolve(customerId)
            }
        }

        request.onerror = () => {
            db.close()
            reject(request.error)
        }

        transaction.onerror = () => {
            db.close()
            reject(transaction.error)
        }
    })
}

/* =========================================================
   RESEQUENCE CUSTOMER IDS AFTER SALE DELETE
========================================================= */

export async function resequenceCustomerIds() {
    const db = await openDB()

    return new Promise((resolve, reject) => {
        const transaction = db.transaction(
            STORES.TRANSACTIONS,
            'readwrite'
        )

        const store = transaction.objectStore(
            STORES.TRANSACTIONS
        )

        const request = store.getAll()

        request.onsuccess = () => {
            const allTransactions =
                request.result || []

            /*
              Only sales with Customer IDs are considered.
              Each month has its own independent sequence.
            */
            const monthlySales = {}

            allTransactions
                .filter(
                    (transaction) =>
                        transaction.type === 'sale' &&
                        transaction.customerRef
                )
                .forEach((sale) => {
                    const saleDate =
                        sale.date ||
                        sale.time ||
                        ''

                    const month =
                        String(saleDate).slice(0, 7)

                    if (!monthlySales[month]) {
                        monthlySales[month] = []
                    }

                    monthlySales[month].push(sale)
                })

            /*
              Keep the original sale order.
              The earliest sale in each month gets
              CUST-001, next gets CUST-002, etc.
            */
            const saleCustomerIdMap = new Map()
            const monthlyCounters = {}

            Object.keys(monthlySales).forEach(
                (month) => {
                    const salesForMonth =
                        monthlySales[month].sort(
                            (a, b) => {
                                const timeA =
                                    new Date(
                                        a.time || 0
                                    ).getTime()

                                const timeB =
                                    new Date(
                                        b.time || 0
                                    ).getTime()

                                if (
                                    timeA !==
                                    timeB
                                ) {
                                    return (
                                        timeA -
                                        timeB
                                    )
                                }

                                return String(
                                    a.id
                                ).localeCompare(
                                    String(b.id)
                                )
                            }
                        )

                    monthlyCounters[month] =
                        salesForMonth.length

                    salesForMonth.forEach(
                        (sale, index) => {
                            const newCustomerRef =
                                `CUST-${String(
                                    index + 1
                                ).padStart(3, '0')}`

                            saleCustomerIdMap.set(
                                sale.id,
                                newCustomerRef
                            )

                            /*
                              Only update the record if
                              its Customer ID actually changed.
                            */
                            if (
                                sale.customerRef !==
                                newCustomerRef
                            ) {
                                store.put({
                                    ...sale,
                                    customerRef:
                                        newCustomerRef,
                                })
                            }
                        }
                    )
                }
            )

            /*
              Update refund Customer IDs so they remain
              linked to their original sale.
            */
            allTransactions
                .filter(
                    (transaction) =>
                        transaction.type ===
                        'refund' &&
                        transaction.originalSaleId
                )
                .forEach((refund) => {
                    const newCustomerRef =
                        saleCustomerIdMap.get(
                            refund.originalSaleId
                        )

                    if (
                        newCustomerRef &&
                        refund.customerRef !==
                        newCustomerRef
                    ) {
                        store.put({
                            ...refund,
                            customerRef:
                                newCustomerRef,
                        })
                    }
                })

            transaction.oncomplete =
                async () => {
                    try {
                        db.close()

                        /*
                          Update the Customer ID counter
                          for the current month only.
                        */
                        const appDb =
                            await openDB()

                        await new Promise(
                            (
                                resolveState,
                                rejectState
                            ) => {
                                const stateTransaction =
                                    appDb.transaction(
                                        STORES.APP_STATE,
                                        'readwrite'
                                    )

                                const stateStore =
                                    stateTransaction.objectStore(
                                        STORES.APP_STATE
                                    )

                                const currentMonth =
                                    new Date()
                                        .toISOString()
                                        .slice(
                                            0,
                                            7
                                        )

                                const currentMonthCount =
                                    monthlyCounters[
                                    currentMonth
                                    ] || 0

                                stateStore.put({
                                    key:
                                        'customerIdCounter',
                                    month:
                                        currentMonth,
                                    count:
                                        currentMonthCount,
                                })

                                stateTransaction.oncomplete =
                                    () => {
                                        appDb.close()
                                        resolveState()
                                    }

                                stateTransaction.onerror =
                                    () => {
                                        appDb.close()
                                        rejectState(
                                            stateTransaction.error
                                        )
                                    }
                            }
                        )

                        resolve({
                            success: true,
                            monthlyCounters,
                        })
                    } catch (error) {
                        reject(error)
                    }
                }

            request.onerror = () => {
                db.close()
                reject(request.error)
            }

            transaction.onerror = () => {
                db.close()
                reject(transaction.error)
            }

            transaction.onabort = () => {
                db.close()

                reject(
                    transaction.error ||
                    new Error(
                        'Customer ID resequencing failed.'
                    )
                )
            }
        }
    })
}

/* =========================================================
   BACKUP DATA COLLECTOR
========================================================= */

export async function createBackupData() {
    const db = await openDB()

    return new Promise((resolve, reject) => {
        const transaction = db.transaction(
            [
                STORES.DAILY_SESSIONS,
                STORES.TRANSACTIONS,
                STORES.USERS,
                STORES.AUDIT_LOGS,
                STORES.APP_STATE,
            ],
            'readonly'
        )

        const dailySessionsStore =
            transaction.objectStore(
                STORES.DAILY_SESSIONS
            )

        const transactionsStore =
            transaction.objectStore(
                STORES.TRANSACTIONS
            )

        const usersStore =
            transaction.objectStore(
                STORES.USERS
            )

        const auditLogsStore =
            transaction.objectStore(
                STORES.AUDIT_LOGS
            )

        const appStateStore =
            transaction.objectStore(
                STORES.APP_STATE
            )

        const dailySessionsRequest =
            dailySessionsStore.getAll()

        const transactionsRequest =
            transactionsStore.getAll()

        const usersRequest =
            usersStore.getAll()

        const auditLogsRequest =
            auditLogsStore.getAll()

        const appStateRequest =
            appStateStore.getAll()

        transaction.oncomplete = () => {
            db.close()

            const appStateRecords =
                appStateRequest.result || []

            const currentSession =
                appStateRecords.find(
                    (item) =>
                        item.key ===
                        'currentSession'
                ) || null

            const customerIdCounter =
                appStateRecords.find(
                    (item) =>
                        item.key ===
                        'customerIdCounter'
                ) || null

            console.log('BACKUP RECORD COUNTS:', {
                dailySessions:
                    dailySessionsRequest.result?.length || 0,

                transactions:
                    transactionsRequest.result?.length || 0,

                users:
                    usersRequest.result?.length || 0,

                auditLogs:
                    auditLogsRequest.result?.length || 0,
            })

            const backupData = {
                dailySessions:
                    dailySessionsRequest.result || [],

                transactions:
                    transactionsRequest.result || [],

                users:
                    usersRequest.result || [],

                auditLogs:
                    auditLogsRequest.result || [],

                appState: {
                    currentSession,
                    customerIdCounter,
                },
            }

            resolve(backupData)
        }

        transaction.onerror = () => {
            db.close()
            reject(transaction.error)
        }

        transaction.onabort = () => {
            db.close()
            reject(
                transaction.error ||
                new Error(
                    'Backup data collection was aborted.'
                )
            )
        }
    })
}

/* =========================================================
   BACKUP PACKAGE BUILDER
========================================================= */

export async function buildBackupPackage() {
    const backupData = await createBackupData()

    const backupPackage = {
        metadata: {
            appName: BACKUP_APP_NAME,
            backupVersion: BACKUP_FORMAT_VERSION,
            databaseName: DB_NAME,
            databaseVersion: DB_VERSION,
            createdAt: new Date().toISOString(),
        },

        data: backupData,
    }

    return backupPackage
}

/* =========================================================
   BACKUP FILE GENERATION
========================================================= */

export async function generateBackupFile() {
    const backupPackage =
        await buildBackupPackage()

    const jsonData =
        JSON.stringify(
            backupPackage,
            null,
            2
        )

    const blob = new Blob(
        [jsonData],
        {
            type: 'application/json',
        }
    )

    const timestamp =
        new Date()
            .toISOString()
            .replace(/[:.]/g, '-')
            .replace('T', '_')
            .slice(0, 19)

    const fileName =
        `${BACKUP_APP_NAME.replace(
            /\s+/g,
            ''
        )}_Backup_${timestamp}${BACKUP_FILE_EXTENSION}`

    const downloadUrl =
        URL.createObjectURL(blob)

    const link =
        document.createElement('a')

    link.href = downloadUrl
    link.download = fileName

    document.body.appendChild(link)

    link.click()

    document.body.removeChild(link)

    URL.revokeObjectURL(downloadUrl)

    return {
        fileName,
        size: blob.size,
    }
}

/* =========================================================
   BACKUP VALIDATION
========================================================= */

export function validateBackupPackage(
    backupPackage
) {
    if (!backupPackage) {
        return {
            valid: false,
            error: 'Backup data is empty.',
        }
    }

    if (
        typeof backupPackage !== 'object' ||
        Array.isArray(backupPackage)
    ) {
        return {
            valid: false,
            error: 'Invalid backup structure.',
        }
    }

    const metadata =
        backupPackage.metadata

    const data =
        backupPackage.data

    if (!metadata || typeof metadata !== 'object') {
        return {
            valid: false,
            error: 'Backup metadata is missing.',
        }
    }

    if (!data || typeof data !== 'object') {
        return {
            valid: false,
            error: 'Backup data section is missing.',
        }
    }

    if (
        metadata.appName !==
        BACKUP_APP_NAME
    ) {
        return {
            valid: false,
            error: 'This backup does not belong to Khalid Medical Store.',
        }
    }

    if (
        metadata.backupVersion !==
        BACKUP_FORMAT_VERSION
    ) {
        return {
            valid: false,
            error: 'Unsupported backup version.',
        }
    }

    if (!Array.isArray(data.dailySessions)) {
        return {
            valid: false,
            error: 'Daily sessions data is invalid.',
        }
    }

    if (!Array.isArray(data.transactions)) {
        return {
            valid: false,
            error: 'Transactions data is invalid.',
        }
    }

    if (!Array.isArray(data.users)) {
        return {
            valid: false,
            error: 'Users data is invalid.',
        }
    }

    if (!Array.isArray(data.auditLogs)) {
        return {
            valid: false,
            error: 'Audit logs data is invalid.',
        }
    }

    if (
        !data.appState ||
        typeof data.appState !== 'object'
    ) {
        return {
            valid: false,
            error: 'App state data is invalid.',
        }
    }

    if (
        Object.prototype.hasOwnProperty.call(
            data.appState,
            'activeUser'
        )
    ) {
        return {
            valid: false,
            error:
                'Backup contains active user session data.',
        }
    }

    return {
        valid: true,
        error: null,
    }
}

if (typeof window !== 'undefined') {
    window.testValidateMedicalStoreBackup =
        async () => {
            const backupPackage =
                await buildBackupPackage()

            const result =
                validateBackupPackage(
                    backupPackage
                )

            console.log(
                'BACKUP VALIDATION RESULT:',
                result
            )

            return result
        }
}

/* =========================================================
   BACKUP ENCRYPTION - KEY DERIVATION
========================================================= */

async function deriveBackupKey(
    password,
    salt
) {
    const encoder =
        new TextEncoder()

    const passwordBytes =
        encoder.encode(password)

    const baseKey =
        await crypto.subtle.importKey(
            'raw',
            passwordBytes,
            'PBKDF2',
            false,
            ['deriveKey']
        )

    return crypto.subtle.deriveKey(
        {
            name: 'PBKDF2',
            salt,
            iterations:
                BACKUP_PBKDF2_ITERATIONS,
            hash: 'SHA-256',
        },
        baseKey,
        {
            name:
                BACKUP_ENCRYPTION_ALGORITHM,
            length:
                BACKUP_KEY_LENGTH,
        },
        false,
        [
            'encrypt',
            'decrypt',
        ]
    )
}

/* =========================================================
   BACKUP ENCRYPTION - ENCRYPT DATA
========================================================= */

async function encryptBackupData(
    plainText,
    password
) {
    if (!password) {
        throw new Error(
            'Backup password is required.'
        )
    }

    const encoder =
        new TextEncoder()

    const plainTextBytes =
        encoder.encode(plainText)

    const salt =
        crypto.getRandomValues(
            new Uint8Array(
                BACKUP_SALT_LENGTH
            )
        )

    const iv =
        crypto.getRandomValues(
            new Uint8Array(
                BACKUP_IV_LENGTH
            )
        )

    const key =
        await deriveBackupKey(
            password,
            salt
        )

    const encryptedData =
        await crypto.subtle.encrypt(
            {
                name:
                    BACKUP_ENCRYPTION_ALGORITHM,
                iv,
            },
            key,
            plainTextBytes
        )

    return {
        salt,
        iv,
        encryptedData:
            new Uint8Array(
                encryptedData
            ),
    }
}

/* =========================================================
   BACKUP ENCRYPTION - DECRYPT DATA
========================================================= */

async function decryptBackupData(
    encryptedData,
    password
) {
    if (!password) {
        throw new Error(
            'Backup password is required.'
        )
    }

    const key =
        await deriveBackupKey(
            password,
            encryptedData.salt
        )

    try {
        const decryptedData =
            await crypto.subtle.decrypt(
                {
                    name:
                        BACKUP_ENCRYPTION_ALGORITHM,
                    iv:
                        encryptedData.iv,
                },
                key,
                encryptedData.encryptedData
            )

        const decoder =
            new TextDecoder()

        return decoder.decode(
            decryptedData
        )
    } catch {
        throw new Error(
            'Unable to decrypt backup. The password may be incorrect or the backup may be corrupted.'
        )
    }
}



/* =========================================================
   SECURE BACKUP PACKAGE
========================================================= */

export async function buildEncryptedBackupPackage(
    password
) {
    if (!password) {
        throw new Error(
            'Backup password is required.'
        )
    }

    const backupPackage =
        await buildBackupPackage()

    const plainText =
        JSON.stringify(
            backupPackage
        )

    const encrypted =
        await encryptBackupData(
            plainText,
            password
        )

    const encryptedBytes =
        Array.from(
            encrypted.encryptedData
        )

    const saltBytes =
        Array.from(
            encrypted.salt
        )

    const ivBytes =
        Array.from(
            encrypted.iv
        )

    return {
        metadata: {
            appName:
                BACKUP_APP_NAME,

            backupVersion:
                BACKUP_FORMAT_VERSION,

            databaseName:
                DB_NAME,

            databaseVersion:
                DB_VERSION,

            encryption: {
                algorithm:
                    BACKUP_ENCRYPTION_ALGORITHM,

                keyLength:
                    BACKUP_KEY_LENGTH,

                saltLength:
                    BACKUP_SALT_LENGTH,

                ivLength:
                    BACKUP_IV_LENGTH,

                kdf: 'PBKDF2',

                hash: 'SHA-256',

                iterations:
                    BACKUP_PBKDF2_ITERATIONS,
            },

            createdAt:
                new Date().toISOString(),
        },

        encryptedData: {
            salt: saltBytes,
            iv: ivBytes,
            data: encryptedBytes,
        },
    }
}


/* =========================================================
   ENCRYPTED BACKUP FILE GENERATION
========================================================= */

export async function generateEncryptedBackupFile(
    password
) {
    if (!password) {
        throw new Error(
            'Backup password is required.'
        )
    }

    const encryptedPackage =
        await buildEncryptedBackupPackage(
            password
        )

    const jsonData =
        JSON.stringify(
            encryptedPackage
        )

    const blob =
        new Blob(
            [jsonData],
            {
                type: 'application/octet-stream',
            }
        )

    const timestamp =
        new Date()
            .toISOString()
            .replace(/[:.]/g, '-')
            .replace('T', '_')
            .slice(0, 19)

    const fileName =
        `${BACKUP_APP_NAME.replace(
            /\s+/g,
            ''
        )}_Backup_${timestamp}${BACKUP_FILE_EXTENSION}`

    const downloadUrl =
        URL.createObjectURL(blob)

    const link =
        document.createElement('a')

    link.href =
        downloadUrl

    link.download =
        fileName

    document.body.appendChild(link)

    link.click()

    document.body.removeChild(link)

    URL.revokeObjectURL(
        downloadUrl
    )

    /* --------------------------------
       Save Backup History
    -------------------------------- */

    await saveBackupHistory({
        fileName,
        size: blob.size,
        type: 'encrypted',
    })

    return {
        fileName,
        size: blob.size,
    }
}

/* =========================================================
   ENCRYPTED BACKUP FILE READER
========================================================= */

export async function readEncryptedBackupFile(
    file,
    password
) {
    if (!file) {
        throw new Error(
            'Backup file is required.'
        )
    }

    if (!password) {
        throw new Error(
            'Backup password is required.'
        )
    }

    const fileText =
        await file.text()

    let encryptedPackage

    try {
        encryptedPackage =
            JSON.parse(fileText)
    } catch {
        throw new Error(
            'Backup file is not a valid backup file.'
        )
    }

    if (
        !encryptedPackage ||
        typeof encryptedPackage !==
        'object'
    ) {
        throw new Error(
            'Invalid backup package.'
        )
    }

    const metadata =
        encryptedPackage.metadata

    const encryptedData =
        encryptedPackage.encryptedData

    if (
        !metadata ||
        typeof metadata !== 'object'
    ) {
        throw new Error(
            'Backup metadata is missing.'
        )
    }

    if (
        metadata.appName !==
        BACKUP_APP_NAME
    ) {
        throw new Error(
            'This backup does not belong to Khalid Medical Store.'
        )
    }

    if (
        metadata.backupVersion !==
        BACKUP_FORMAT_VERSION
    ) {
        throw new Error(
            'Unsupported backup version.'
        )
    }

    if (
        !encryptedData ||
        !Array.isArray(
            encryptedData.salt
        ) ||
        !Array.isArray(
            encryptedData.iv
        ) ||
        !Array.isArray(
            encryptedData.data
        )
    ) {
        throw new Error(
            'Encrypted backup data is invalid.'
        )
    }

    if (
        encryptedData.salt.length !==
        BACKUP_SALT_LENGTH
    ) {
        throw new Error(
            'Backup salt is invalid.'
        )
    }

    if (
        encryptedData.iv.length !==
        BACKUP_IV_LENGTH
    ) {
        throw new Error(
            'Backup IV is invalid.'
        )
    }

    const decryptedText =
        await decryptBackupData(
            {
                salt:
                    new Uint8Array(
                        encryptedData.salt
                    ),

                iv:
                    new Uint8Array(
                        encryptedData.iv
                    ),

                encryptedData:
                    new Uint8Array(
                        encryptedData.data
                    ),
            },
            password
        )

    let backupPackage

    try {
        backupPackage =
            JSON.parse(
                decryptedText
            )
    } catch {
        throw new Error(
            'Decrypted backup data is invalid.'
        )
    }
    const validation =
        validateBackupPackage(
            backupPackage
        )

    if (!validation.valid) {
        throw new Error(
            validation.error ||
            'Backup validation failed.'
        )
    }

    return backupPackage
}

/* =========================================================
   RESTORE DATA VALIDATION
========================================================= */

export function validateRestoreData(
    backupPackage
) {
    const validation =
        validateBackupPackage(
            backupPackage
        )

    if (!validation.valid) {
        return validation
    }

    const data =
        backupPackage.data

    if (
        !data.appState ||
        typeof data.appState !==
        'object'
    ) {
        return {
            valid: false,
            error:
                'Restore app state is invalid.',
        }
    }

    if (
        data.appState.activeUser !==
        undefined
    ) {
        return {
            valid: false,
            error:
                'Restore data contains active user session.',
        }
    }

    return {
        valid: true,
        error: null,
    }
}

/* =========================================================
   RESTORE DATA PREPARATION
========================================================= */

export function prepareRestoreData(
    backupPackage
) {
    const validation =
        validateRestoreData(
            backupPackage
        )

    if (!validation.valid) {
        throw new Error(
            validation.error ||
            'Restore data validation failed.'
        )
    }

    const data =
        backupPackage.data

    return {
        dailySessions:
            [...data.dailySessions],

        transactions:
            [...data.transactions],

        users:
            [...data.users],

        auditLogs:
            [...data.auditLogs],

        currentSession:
            data.appState.currentSession
                ? {
                    ...data.appState
                        .currentSession,
                }
                : null,

        customerIdCounter:
            data.appState.customerIdCounter
                ? {
                    ...data.appState
                        .customerIdCounter,
                }
                : null,
    }
}

/* =========================================================
   RESTORE DATABASE
========================================================= */

export async function restoreDatabase(
    backupPackage
) {
    const restoreData =
        prepareRestoreData(
            backupPackage
        )

    const db =
        await openDB()

    return new Promise(
        (resolve, reject) => {
            const transaction =
                db.transaction(
                    [
                        STORES.DAILY_SESSIONS,
                        STORES.TRANSACTIONS,
                        STORES.USERS,
                        STORES.AUDIT_LOGS,
                        STORES.APP_STATE,
                    ],
                    'readwrite'
                )

            const dailySessionsStore =
                transaction.objectStore(
                    STORES.DAILY_SESSIONS
                )

            const transactionsStore =
                transaction.objectStore(
                    STORES.TRANSACTIONS
                )

            const usersStore =
                transaction.objectStore(
                    STORES.USERS
                )

            const auditLogsStore =
                transaction.objectStore(
                    STORES.AUDIT_LOGS
                )

            const appStateStore =
                transaction.objectStore(
                    STORES.APP_STATE
                )

            dailySessionsStore.clear()

            transactionsStore.clear()

            usersStore.clear()

            auditLogsStore.clear()

            appStateStore.clear()

            restoreData.dailySessions.forEach(
                (record) => {
                    dailySessionsStore.put(
                        record
                    )
                }
            )

            restoreData.transactions.forEach(
                (record) => {
                    transactionsStore.put(
                        record
                    )
                }
            )

            restoreData.users.forEach(
                (user) => {
                    usersStore.put(
                        user
                    )
                }
            )

            restoreData.auditLogs.forEach(
                (log) => {
                    auditLogsStore.put(
                        log
                    )
                }
            )

            if (
                restoreData.currentSession
            ) {
                appStateStore.put({
                    ...restoreData.currentSession,
                    key:
                        'currentSession',
                })
            }

            if (
                restoreData.customerIdCounter
            ) {
                appStateStore.put({
                    ...restoreData.customerIdCounter,
                    key:
                        'customerIdCounter',
                })
            }

            transaction.oncomplete =
                () => {
                    db.close()

                    resolve({
                        success: true,

                        dailySessionsCount:
                            restoreData
                                .dailySessions
                                .length,

                        transactionsCount:
                            restoreData
                                .transactions
                                .length,

                        usersCount:
                            restoreData
                                .users
                                .length,

                        auditLogsCount:
                            restoreData
                                .auditLogs
                                .length,

                        hasCurrentSession:
                            restoreData
                                .currentSession !==
                            null,

                        hasCustomerIdCounter:
                            restoreData
                                .customerIdCounter !==
                            null,
                    })
                }

            transaction.onerror =
                () => {
                    db.close()

                    reject(
                        transaction.error ||
                        new Error(
                            'Database restore failed.'
                        )
                    )
                }

            transaction.onabort =
                () => {
                    db.close()

                    reject(
                        transaction.error ||
                        new Error(
                            'Database restore was aborted.'
                        )
                    )
                }
        }
    )
}

/* =========================================================
   RESTORE ENCRYPTED BACKUP FILE
========================================================= */

export async function restoreEncryptedBackupFile(
    file,
    password
) {
    if (!file) {
        throw new Error(
            'Backup file is required.'
        )
    }

    if (!password) {
        throw new Error(
            'Backup password is required.'
        )
    }

    const backupPackage =
        await readEncryptedBackupFile(
            file,
            password
        )

    const validation =
        validateRestoreData(
            backupPackage
        )

    if (!validation.valid) {
        throw new Error(
            validation.error ||
            'Backup cannot be restored.'
        )
    }

    const safetyBackup =
        await generateSafetyBackupFile(
            password
        )

    const restoreResult =
        await restoreDatabase(
            backupPackage
        )

    return {
        success:
            restoreResult.success,

        restoreResult,

        safetyBackup,
    }
}

/* =========================================================
   CREATE SAFETY BACKUP BEFORE RESTORE
========================================================= */

export async function createSafetyBackup(
    password
) {
    if (!password) {
        throw new Error(
            'Backup password is required.'
        )
    }

    const encryptedPackage =
        await buildEncryptedBackupPackage(
            password
        )

    return encryptedPackage
}

/* =========================================================
   SAFETY BACKUP FILE GENERATION
========================================================= */

export async function generateSafetyBackupFile(
    password
) {
    if (!password) {
        throw new Error(
            'Backup password is required.'
        )
    }

    const encryptedPackage =
        await createSafetyBackup(
            password
        )

    const jsonData =
        JSON.stringify(
            encryptedPackage
        )

    const blob =
        new Blob(
            [jsonData],
            {
                type:
                    'application/octet-stream',
            }
        )

    const timestamp =
        new Date()
            .toISOString()
            .replace(/[:.]/g, '-')
            .replace('T', '_')
            .slice(0, 19)

    const fileName =
        `${BACKUP_APP_NAME.replace(
            /\s+/g,
            ''
        )}_SafetyBackup_${timestamp}${BACKUP_FILE_EXTENSION}`

    const downloadUrl =
        URL.createObjectURL(blob)

    const link =
        document.createElement('a')

    link.href =
        downloadUrl

    link.download =
        fileName

    document.body.appendChild(link)

    link.click()

    document.body.removeChild(link)

    URL.revokeObjectURL(
        downloadUrl
    )

    await saveBackupHistory({
        fileName,
        size: blob.size,
        type: 'safety',
    })

    return {
        fileName,
        size: blob.size,
    }
}

/* =========================================================
   RESTORE PREVIEW
========================================================= */

export async function previewEncryptedBackupRestore(
    file,
    password
) {
    if (!file) {
        throw new Error(
            'Backup file is required.'
        )
    }

    if (!password) {
        throw new Error(
            'Backup password is required.'
        )
    }

    const backupPackage =
        await readEncryptedBackupFile(
            file,
            password
        )

    const validation =
        validateRestoreData(
            backupPackage
        )

    if (!validation.valid) {
        throw new Error(
            validation.error ||
            'Backup cannot be restored.'
        )
    }

    const restoreData =
        prepareRestoreData(
            backupPackage
        )

    return {
        valid: true,

        backupCreatedAt:
            backupPackage.metadata.createdAt,

        dailySessionsCount:
            restoreData.dailySessions.length,

        transactionsCount:
            restoreData.transactions.length,

        usersCount:
            restoreData.users.length,

        auditLogsCount:
            restoreData.auditLogs.length,

        hasCurrentSession:
            restoreData.currentSession !==
            null,

        hasCustomerIdCounter:
            restoreData.customerIdCounter !==
            null,

        activeUserIncluded: false,
    }
}

/* =========================================================
   CONFIRMED BACKUP RESTORE
========================================================= */

export async function confirmEncryptedBackupRestore(
    file,
    password,
    confirmed
) {
    if (!confirmed) {
        throw new Error(
            'Restore confirmation is required.'
        )
    }

    if (!file) {
        throw new Error(
            'Backup file is required.'
        )
    }

    if (!password) {
        throw new Error(
            'Backup password is required.'
        )
    }

    const backupPackage =
        await readEncryptedBackupFile(
            file,
            password
        )

    const validation =
        validateRestoreData(
            backupPackage
        )

    if (!validation.valid) {
        throw new Error(
            validation.error ||
            'Backup cannot be restored.'
        )
    }

    /*
       Create the safety backup package first.
       This keeps the current data protected without
       making the restore dependent on the browser
       download/history operation.
    */
    const safetyBackup =
        await createSafetyBackup(
            password
        )

    /*
       Now perform the actual database restore.
    */
    const restoreResult =
        await restoreDatabase(
            backupPackage
        )

    return {
        success:
            restoreResult.success,

        restoreResult,

        safetyBackup,
    }
}

export async function saveBackupHistory(backupRecord) {
    const db = await openDB()

    return new Promise((resolve, reject) => {
        const transaction = db.transaction(
            STORES.BACKUP_HISTORY,
            'readwrite'
        )

        const store = transaction.objectStore(
            STORES.BACKUP_HISTORY
        )

        const request = store.add({
            ...backupRecord,
            createdAt:
                backupRecord?.createdAt ||
                new Date().toISOString(),
        })

        request.onsuccess = () => {
            const getAllRequest = store.getAll()

            getAllRequest.onsuccess = () => {
                const records =
                    getAllRequest.result || []

                if (records.length <= 20) {
                    resolve(request.result)
                    return
                }

                records.sort(
                    (a, b) =>
                        new Date(b.createdAt).getTime() -
                        new Date(a.createdAt).getTime()
                )

                const recordsToDelete =
                    records.slice(20)

                let deletedCount = 0

                recordsToDelete.forEach((record) => {
                    const deleteRequest =
                        store.delete(record.id)

                    deleteRequest.onsuccess = () => {
                        deletedCount += 1

                        if (
                            deletedCount ===
                            recordsToDelete.length
                        ) {
                            resolve(request.result)
                        }
                    }

                    deleteRequest.onerror = () => {
                        reject(deleteRequest.error)
                    }
                })
            }

            getAllRequest.onerror = () => {
                reject(getAllRequest.error)
            }
        }

        request.onerror = () => {
            reject(request.error)
        }
    })
}

export async function getBackupHistory() {
    const db = await openDB()

    return new Promise((resolve, reject) => {
        const transaction = db.transaction(
            STORES.BACKUP_HISTORY,
            'readonly'
        )

        const store = transaction.objectStore(
            STORES.BACKUP_HISTORY
        )

        const request = store.getAll()

        request.onsuccess = () => {
            const records = request.result || []

            records.sort((a, b) => {
                const dateDifference =
                    new Date(b.createdAt).getTime() -
                    new Date(a.createdAt).getTime()

                if (dateDifference !== 0) {
                    return dateDifference
                }

                return (b.id || 0) - (a.id || 0)
            })

            resolve(records)
        }

        request.onerror = () => {
            reject(request.error)
        }
    })
}

export async function clearBackupHistory() {
    const db = await openDB()

    return new Promise((resolve, reject) => {
        const transaction = db.transaction(
            STORES.BACKUP_HISTORY,
            'readwrite'
        )

        const store = transaction.objectStore(
            STORES.BACKUP_HISTORY
        )

        const request = store.clear()

        request.onsuccess = () => {
            resolve(true)
        }

        request.onerror = () => {
            reject(request.error)
        }
    })
}