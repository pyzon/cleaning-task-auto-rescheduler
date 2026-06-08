/**
 * ============================================================
 * CLEANING TASK AUTO-RESCHEDULER
 * ============================================================
 *
 * This Apps Script synchronizes a Google Tasks list with a
 * spreadsheet that tracks historical task completion data.
 *
 * For each defined task:
 * - If it does not exist in Google Tasks, it is created.
 * - If it exists and was completed, the completion date is logged.
 * - A new due date is calculated based on the spreadsheet's
 *   computed moving average interval.
 * - The task is rescheduled and marked as incomplete.
 *
 * The moving average mechanism gradually adapts the due date
 * based on actual completion behavior. If a task has been done
 * more frequently in recent cycles, the calculated interval
 * decreases, bringing the next due date closer. If it has been
 * completed less frequently, the interval increases gradually,
 * pushing the next due date further out. Over time, this creates
 * a self-adjusting rhythm aligned with real-world usage patterns.
 *
 * The name "cleaning" appears throughout the script because
 * this approach is particularly suitable for recurring cleaning
 * tasks.
 *
 * The spreadsheet acts as the authoritative source for
 * historical data and interval calculations, while the script
 * handles automation and synchronization with Google Tasks.
 */

/**
 * Main entry point.
 *
 * Iterates through all task titles defined in the first row of the Dates sheet.
 * For each task:
 *  - Ensures the task exists in the Google Tasks list (creates it if missing)
 *  - If completed, logs the completion date into the Dates sheet
 *  - Reads the calculated moving average interval from the Intervals sheet
 *  - Sets a new due date based on the computed interval
 *  - Marks the task as incomplete
 *
 * Requires:
 *  - Advanced Google Tasks API enabled
 *  - Google Sheets API enabled
 *  - Script properties properly configured
 */
function rescheduleCleaningTasks() {
  const config = getConfig();
  const spreadsheet = SpreadsheetApp.openById(config.spreadsheetId);

  const datesSheet = spreadsheet.getSheetById(config.datesSheetId);
  const intervalsSheet = spreadsheet.getSheetById(config.intervalsSheetId);

  const taskTitles = gettaskTitles(datesSheet);

  const tasksByTitle = getTasksByTitle(config.taskListId);

  taskTitles.forEach((taskTitle, index) => {
    processTask(
      taskTitle,
      index,
      config.taskListId,
      tasksByTitle,
      datesSheet,
      intervalsSheet,
    );
  });
}

/**
 * Reads and validates required configuration values from Script Properties.
 *
 * Required properties:
 *  - CLEANING_DATA_SPREADSHEET_ID
 *  - DATES_SHEET_ID
 *  - INTERVALS_SHEET_ID
 *  - CLEANING_TASKLIST_ID
 *
 * @returns {Object} Configuration object.
 * @throws {Error} If any required property is missing or empty.
 */
function getConfig() {
  const props = PropertiesService.getScriptProperties();

  function requireProperty(key) {
    const value = props.getProperty(key);
    if (!value) {
      throw new Error('Missing required Script Property: ' + key);
    }
    return value;
  }

  return {
    spreadsheetId: requireProperty('CLEANING_DATA_SPREADSHEET_ID'),
    datesSheetId: Number(requireProperty('DATES_SHEET_ID')),
    intervalsSheetId: Number(requireProperty('INTERVALS_SHEET_ID')),
    taskListId: requireProperty('CLEANING_TASKLIST_ID')
  };
}

/**
 * Retrieves task titles from the first row of the Dates sheet.
 *
 * The first row is expected to contain task titles in consecutive columns,
 * starting from column 1, without gaps. The column order determines the
 * mapping between:
 *   - completion history in the Dates sheet
 *   - calculated intervals in the Intervals sheet
 *   - Google Tasks list items
 *
 * Empty cells are filtered out as a defensive safeguard. Under the intended
 * spreadsheet structure, there should be no empty cells between task titles.
 * The filter prevents accidental blank headers from being treated as tasks.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The Dates sheet.
 * @returns {string[]} Ordered array of task titles.
 */
function gettaskTitles(sheet) {
  const firstRow = sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .getValues()[0];
  return firstRow.filter((title) => title && title.toString().trim() !== "");
}

/**
 * Processes a single task.
 *
 * The function assumes that all tasks in the list have already been fetched
 * and indexed by title (tasksByTitle), avoiding repeated API calls.
 *
 * Behavior:
 *  - If the task does not exist in the task list → creates it and assigns a due date.
 *  - If the task exists but is not completed (needsAction) → does nothing.
 *  - If the task exists and is completed:
 *      - Logs completion date into Dates sheet
 *      - Forces spreadsheet recalculation
 *      - Reads computed interval from Intervals sheet
 *      - Sets new due date
 *      - Marks task as incomplete
 *
 * @param {string} taskTitle - Title of the task.
 * @param {number} columnIndex - Zero-based column index in the spreadsheet.
 * @param {string} taskListId - Google Tasks list ID.
 * @param {Map<string, Object>} tasksByTitle - Preloaded map of task title → Google Tasks task object.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} datesSheet - Sheet storing completion history.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} intervalsSheet - Sheet storing computed intervals.
 */
function processTask(
  taskTitle,
  columnIndex,
  taskListId,
  tasksByTitle,
  datesSheet,
  intervalsSheet,
) {
  const task = tasksByTitle.get(taskTitle);

  if (!task) {
    console.log(taskTitle);
    const newTask = createTask(taskListId, taskTitle);
    const interval = getIntervalForTask(intervalsSheet, columnIndex);
    setTaskDueDate(newTask.id, taskListId, interval);
    return;
  }

  if (task.status !== 'completed') {
    // The only other possible value is needsAction
    return;
  }

  const completionDate = extractCompletionDate(task.completed);
  appendCompletionDate(datesSheet, columnIndex, completionDate);

  SpreadsheetApp.flush(); // ensure calculations are applied

  const interval = getIntervalForTask(intervalsSheet, columnIndex);
  setTaskDueDate(task.id, taskListId, interval);
  markTaskIncomplete(task.id, taskListId);
}

/**
 * Retrieves all tasks from a Google Tasks list and indexes them by exact title.
 *
 * Includes completed and hidden tasks in the result set. The function performs
 * full pagination over the task list and builds an in-memory lookup map where
 * each task title maps to its corresponding task object.
 *
 * If multiple tasks share the same title, only one will be stored in the map
 * (the last one encountered during traversal).
 *
 * @param {string} taskListId - Google Tasks list ID.
 * @returns {Map<string, Object>} Map of task title → task object.
 */
function getTasksByTitle(taskListId) {
  const tasksByTitle = new Map();
  let pageToken = null;

  do {
    const response = Tasks.Tasks.list(taskListId, {
      showCompleted: true,
      showHidden: true,
      maxResults: 100,
      pageToken,
    });

    for (const task of response.items || []) {
      tasksByTitle.set(task.title, task);
    }

    pageToken = response.nextPageToken;
  } while (pageToken);

  return tasksByTitle;
}

/**
 * Creates a new task in the specified Google Tasks list.
 *
 * The task is created without a due date. Due date must be set separately.
 *
 * @param {string} taskListId - Google Tasks list ID.
 * @param {string} title - Task title.
 * @returns {Object} The created task object.
 */
function createTask(taskListId, title) {
  return Tasks.Tasks.insert({
    title: title
  }, taskListId);
}

/**
 * Reads the computed interval value for a task from the Intervals sheet.
 *
 * Assumes:
 *  - Row 1 contains task titles
 *  - Row 2 contains the calculated moving average interval
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} intervalsSheet - Intervals sheet.
 * @param {number} columnIndex - Zero-based column index.
 * @returns {number} Interval in days.
 */
function getIntervalForTask(intervalsSheet, columnIndex) {
  const column = columnIndex + 1;
  return intervalsSheet.getRange(2, column).getValue();
}

/**
 * Sets a task's due date based on an interval in days from today.
 *
 * If intervalDays is invalid or empty, no update is performed.
 *
 * @param {string} taskId - ID of the task.
 * @param {string} taskListId - Google Tasks list ID.
 * @param {number} intervalDays - Number of days to add to today's date.
 */
function setTaskDueDate(taskId, taskListId, intervalDays) {
  if (!intervalDays || isNaN(intervalDays)) return;

  const dueDate = new Date();
  dueDate.setDate(dueDate.getDate() + Number(intervalDays));

  Tasks.Tasks.patch({
    due: dueDate.toISOString()
  }, taskListId, taskId);
}

/**
 * Extracts and formats the completion date from a Google Tasks
 * completion timestamp.
 *
 * @param {string} completedTimestamp - ISO timestamp from Google Tasks.
 * @returns {string} Date formatted as 'yyyy-MM-dd'.
 */
function extractCompletionDate(completedTimestamp) {
  const date = new Date(completedTimestamp);
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

/**
 * Appends a completion date to the first empty cell in a task's column.
 *
 * Data starts at row 2 (row 1 contains headers).
 * If no empty cell exists, the date is appended at the bottom.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Dates sheet.
 * @param {number} columnIndex - Zero-based column index.
 * @param {string} dateString - Date formatted as 'yyyy-MM-dd'.
 */
function appendCompletionDate(sheet, columnIndex, dateString) {
  const column = columnIndex + 1; // convert 0-based to 1-based
  const lastRow = sheet.getLastRow();
  const columnValues = sheet.getRange(2, column, lastRow).getValues();

  for (let i = 0; i < columnValues.length; i++) {
    if (!columnValues[i][0]) {
      sheet.getRange(i + 2, column).setValue(dateString);
      return;
    }
  }

  // If no empty cell found, append at bottom
  sheet.getRange(lastRow + 1, column).setValue(dateString);
}

/**
 * Marks a completed task as incomplete (needsAction).
 *
 * Clears the completed timestamp and updates the status.
 *
 * @param {string} taskId - ID of the task.
 * @param {string} taskListId - Google Tasks list ID.
 */
function markTaskIncomplete(taskId, taskListId) {
  Tasks.Tasks.patch({
    status: 'needsAction',
    completed: null
  }, taskListId, taskId);
}
