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
 */
function initProperties() {
  const props = PropertiesService.getScriptProperties();
  props.setProperty(
    "CLEANING_DATA_SPREADSHEET_ID",
    "1af0IVB08hem7Kxvqw1c-eypuC-SUbXi6lA0EJv_ILig"
  );
  props.setProperty(
    "DATES_SHEET_ID",
    "0"
  );
  props.setProperty(
    "INTERVALS_SHEET_ID",
    "445196433"
  );
  props.setProperty(
    "CLEANING_TASKLIST_ID",
    "NDJKMktMd1dFZ2hwYV9iVQ"
  );
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
  ScriptApp.newTrigger('rescheduleCleaningTasks')
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
