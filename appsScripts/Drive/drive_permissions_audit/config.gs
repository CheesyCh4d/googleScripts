/**
 * CONFIGURATION
 * Global constants used across the project.
 */

// 4.5 minutes (in ms). Safety buffer for the 6-minute execution limit.
const MAX_EXECUTION_TIME = 4.5 * 60 * 1000; 

// How many rows to collect in memory before writing to the Sheet.
const BATCH_SIZE = 50; 

// Max time (in ms) to wait before forcing a write to the sheet.
const FLUSH_INTERVAL = 10 * 1000;