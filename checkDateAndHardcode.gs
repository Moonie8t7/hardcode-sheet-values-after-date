/**
 * @author u/IAmMoonie <https://www.reddit.com/user/IAmMoonie/>
 * @desc The script checks the date in a cell; if the date is the same OR past the date of the cell, it will take the target cell value and enter it as a static number rather than a formula.
 * @license MIT
 * @version 1.0
 */

/* A constant variable that stores the spreadsheet ID. */
const SPREADSHEET_ID = "your sheetID goes here";

/* A constant variable that stores the email address of the user. */
const MY_EMAIL = "put your email here for notifications when the script errors";

/**
 * Example:
 * If the date value in cell D1 is today or later, then copy the value of A1 and rewrite it into A1.
 */
function run() {
  checkDateAndHardcode_("D1", "A1", "some reason");
}

/**
 * It checks if the date in the dateCell is after it or is today, and if so, it clears the content of the
 * targetCell and sets its value to the number resulting from the formula in that cell.
 * @param dateCell - The cell that contains the date.
 * @param targetCell - The cell that you want to hardcode.
 * @param identifier - This is a string that can be used to easily identify this function.
 */
function checkDateAndHardcode_(dateCell, targetCell, identifier) {
  try {
    /* Checking if the dateCell and targetCell are empty. If they are empty, it will throw an error. */
    if (!dateCell || !targetCell) {
      throw new Error("Both date cell and target cell must be specified");
    }
    /* Checking if the dateCell and targetCell are a string. If they are not, it will throw an error. */
    if (typeof dateCell !== "string" || typeof targetCell !== "string") {
      throw new Error("Both date cell and target cell must be strings");
    }
    console.info(
      `Starting script execution: dateCell=${dateCell}, targetCell=${targetCell}, identifier=${identifier}`
    );
    /* Getting the value of the date cell. */
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const dateValue = ss.getRange(dateCell).getValue();
    console.info(`date cell value: ${dateValue}`);
    /* Checking if the date value is a valid date. If it is not, it will throw an error. */
    if (!(dateValue instanceof Date)) {
      throw new Error(
        `The date value in cell ${dateCell} is not in a valid date format`
      );
    }
    /* Creating a new date object with the current date. */
    const today = new Date();
    /* Checking if the date in the dateCell is after or is today, and if so, it clears the content
    of the
     * targetCell and sets its value to the number resulting from the formula in that cell. */
    if (today >= dateValue) {
      const targetRange = ss.getRange(targetCell);
      const formula = targetRange.getFormula();
      /* Checking if the target cell has a formula. If it does, it will log the formula to the console.
      If it does not, it will log the value of the target cell to the console. */
      if (formula) {
        console.info(`target cell FORMULA: ${formula}`);
      } else {
        console.info(`target cell VALUE: ${targetRange.getValue()}`);
      }
      /* Clearing the content of the target cell and then setting the value of the target cell to the
      value of the target cell. */
      targetRange.clearContent();
      targetRange.setValue(targetRange.getValue());
      /* Sending an email to the email address specified in the MY_EMAIL variable. */
      MailApp.sendEmail(
        MY_EMAIL,
        `${identifier} hardcoded`,
        `You can remove the function:\n checkDateAndHardcode_(${dateCell},${targetCell},${identifier})\n\n from your run function`
      );
    }
  } catch (error) {
    /* Logging the error to the console. */
    console.error(error);
    /* Sending an email to the email address specified in the MY_EMAIL variable. */
    MailApp.sendEmail(MY_EMAIL, "Error in script", error);
  }
}
