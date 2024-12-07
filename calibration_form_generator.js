function createFormFromSheet() {
  // Get the active spreadsheet and sheet.
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const data = sheet.getDataRange().getValues(); // Get all data
  const headers = data.shift(); // Remove and store the header row
  const levelInfo = {
    "P2": "Engineering: Tackle work that may take more than a week and may need collaboration.\n\nIT: Takes on solo tasks and completes them in a few days.",
    "P3": "Engineering: Break down multi week tasks and work with others to deliver them.\n\nIT: Tackles work that may take more than a week and may need collaboration."
  };

  // Create a new form.
  const form = FormApp.create(sheet.getName() + " Form Test");
  form.setDescription("Ahead of the calibration meeting, managers will be required to review the proposed ratings and blurbs ahead of our synchronous calibration discussions, and provide feedback.\n\nThe below form lists each employee for the appropriate calibration session. Please respond for each person so we can get an idea of where time will need to be spent in our session together. This session focuses on our IC cohort (P2-P5). Please review the Engineering CLG & IT CLG as you go through.\n\nThis form will remain open until Tuesday, July 2 at 12pm PT for reviewers to submit their feedback, ask clarifying questions, or suggest additional colleague feedback or supporting documentation.\n\nNOTE: this form will automatically collect respondents email addresses and will be shared with PBP and session facilitators.")

  const levels = Object.keys(levelInfo);

  levels.forEach((level, index) => {
    const rows = getRowsByHeaderValue(data, headers, "Job Level", level);
    // Add section for level description
    form.addPageBreakItem().setTitle(level).setHelpText(levelInfo[level]);
    addPromptsForLevel(form,rows,headers);
  });
}

function addPromptsForLevel(form, rows, headers) {
  // Loop through the headers to create form items.
  rows.forEach((row, index) => {
    const page = form.addPageBreakItem();
    var item;
    const name = getValueByHeader(row,headers,"Worker");
    const whatRating = getValueByHeader(row,headers,"What Rating");
    const howRating = getValueByHeader(row,headers,"How Rating");
    const impactRating = getValueByHeader(row,headers,"Impact Rating");
    const title = name + " [" + whatRating + " / " + howRating + " / " + impactRating +"]"
    page.setTitle(title);
    page.setHelpText(getValueByHeader(row,headers,"Blurb"));
    item = form.addMultipleChoiceItem(); 
    item.setTitle(title)
    .setChoices([item.createChoice('I have some some familiarity on this person’s work, and I support the ratings and blurb'), item.createChoice('I don’t have any context on this person or their work, but the rating/blurb seem reasonable'), item.createChoice("I would like additional information and/or I have questions about the ratings/blurb")])
    .showOtherOption(false);
    item = form.addParagraphTextItem();
    item.setTitle("If you answered that you'd like more information or have questions, please elaborate.");
  });
}

function getValueByHeader(row, headers, headerName) {
  const columnIndex = headers.indexOf(headerName);

  if (columnIndex !== -1) {
    return row[columnIndex];
  } else {
    // Handle case where header is not found
    Logger.log(`Header "${headerName}" not found.`);
    return null; // Or throw an error, or return a default value
  }
}

function getRowsByHeaderValue(data, headers, headerName, targetValue) {
  const columnIndex = headers.indexOf(headerName);

  if (columnIndex === -1) {
    Logger.log(`Header "${headerName}" not found.`);
    return []; // Or throw an error
  }

  const matchingRows = [];

  data.forEach((row, rowIndex) => {
    if (row[columnIndex] === targetValue) {
      matchingRows.push(row); 
    }
  });

  return matchingRows;
}