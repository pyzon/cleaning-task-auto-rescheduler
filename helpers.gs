/**
 * ============================================================
 * HELPER FUNCTIONS
 * ============================================================
 */

/***
 * Persists constants into PropertiesService
 *
 * Only needs to be run once, then properties are stored indefinitely
 * (regardless of executions, script edits, deployments)
 * Of course, directly deleting the properties, or the values
 * having changed are some of the reasons to run this.
 * You can review your Script Properties in the Project Settings.
 *
 * TODO: replace placeholders with your IDs
 */
function initProperties() {
  const props = PropertiesService.getScriptProperties();
  props.setProperty(
    "CLEANING_DATA_SPREADSHEET_ID",
    "00000000000000000000000000000000000000000000",
  );
  props.setProperty("DATES_SHEET_ID", "0");
  props.setProperty("INTERVALS_SHEET_ID", "000000000");
  props.setProperty("CLEANING_TASKLIST_ID", "0000000000000000000000");
}

/**
 * Deletes all properties of this script
 *
 * Use it if you made a mistake and want to start over.
 */
function deleteAllProperties() {
  const props = PropertiesService.getScriptProperties();
  props.deleteAllProperties();
  //Logger.log(props.getProperties());
}

/**
 * Creates a daily time-driven trigger for rescheduleCleaningTasks().
 * Runs once per day around 23:00.
 * Warning: not idempotent, if you run it twice, it creates two triggers.
 * You can review your triggers in the Triggers tab.
 */
function createDailyTrigger() {
  ScriptApp.newTrigger("rescheduleCleaningTasks")
    .timeBased()
    .everyDays(1)
    .atHour(23)
    .nearMinute(0)
    .create();
}

/**
 * Lists task lists.
 * Can be used to find ids.
 */
function listTaskLists() {
  const taskLists = Tasks.Tasklists.list().items || [];

  for (const taskList of taskLists) {
    Logger.log("Task List: " + taskList);
  }
}

/**
 * Removes duplicate tasks by title from the cleaning Google Tasks list.
 * Keeps the first task encountered and deletes the others.
 *
 * @returns {number} Number of tasks deleted.
 */
function removeDuplicateTasks() {
  const config = getConfig();
  const taskListId = config.taskListId;
  const seenTitles = new Set();
  const duplicateTaskIds = [];

  let pageToken;

  do {
    const response = Tasks.Tasks.list(taskListId, {
      showCompleted: true,
      showHidden: true,
      maxResults: 100,
      pageToken,
    });

    for (const task of response.items || []) {
      if (seenTitles.has(task.title)) {
        duplicateTaskIds.push(task.id);
        console.log(`Duplicate found: ${task.title} (Due: ${task.due})`);
      } else {
        seenTitles.add(task.title);
      }
    }

    pageToken = response.nextPageToken;
  } while (pageToken);

  for (const taskId of duplicateTaskIds) {
    Tasks.Tasks.remove(taskListId, taskId);
  }

  console.log(`Deleted ${duplicateTaskIds.length} duplicate tasks.`);
}
