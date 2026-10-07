/**
 * Creates the "Student Wellbeing Survey" Google Form and a linked response Google Sheet
 * in the account of whoever runs it.
 *
 * How to use:
 *   1. Open https://script.google.com and click "New project".
 *   2. Replace the editor contents with this file and click Save.
 *   3. Select the function createStudentWellbeingForm and click Run.
 *      (The first run asks you to authorise access to Google Forms and Sheets.)
 *   4. Open View > Logs (or the Execution log) to get the form and sheet links.
 *
 * The questions match the columns of data/student_sample_317.csv, so responses collected with this form
 * can be analysed with the same notebook and Stata do-file.
 */
function createStudentWellbeingForm() {
  var title = 'Student Wellbeing Survey';
  var form = FormApp.create(title);
  form.setDescription(
    'This short, anonymous survey is part of an econometrics course project at the Department of Humanities ' +
    'and Social Sciences, IIT Kharagpur. It asks about your background, studies, social life and wellbeing. ' +
    'It takes about 3 minutes. Participation is voluntary, no names or email addresses are collected, and ' +
    'results are reported only as group statistics.'
  );
  form.setCollectEmail(false);
  form.setAllowResponseEdits(false);
  form.setLimitOneResponsePerUser(false); // keeps the form open without requiring sign-in
  form.setProgressBar(true);
  form.setConfirmationMessage(
    'Thank you for taking part. If you are feeling low or overwhelmed, please reach out to your university ' +
    'counselling centre or someone you trust.'
  );

  // ---- Consent: "No" ends the form
  var consent = form.addMultipleChoiceItem()
    .setTitle('I am 18 years or older and I agree to take part in this anonymous survey.')
    .setRequired(true);
  consent.setChoices([
    consent.createChoice('Yes, I agree', FormApp.PageNavigationType.CONTINUE),
    consent.createChoice('No', FormApp.PageNavigationType.SUBMIT)
  ]);

  // ---- Section 1: about you
  form.addPageBreakItem().setTitle('About you');
  addChoice(form, 'Gender', ['Male', 'Female', 'Other / Prefer not to say']);
  addNumber(form, 'Age (in years)', 16, 40, 'Please enter your age in whole years.');
  addChoice(form, 'Family background', ['Lower Middle class', 'Middle class', 'Rich class']);
  addChoice(form, 'Relationship status', ['Single', 'In a Relationship', 'Complicated']);
  addChoice(form, 'Living style', ['Bachelor', 'Family'],
    'Bachelor = hostel, mess or shared accommodation; Family = living with your family.');

  // ---- Section 2: studies
  form.addPageBreakItem().setTitle('Your studies');
  addNumber(form, 'Current semester', 0, 14, 'Enter 0 if you have not completed your first semester yet.');
  addNumber(form, 'Current CGPA (on a 10.00 scale)', 0, 10, 'Example: 8.45');

  // ---- Section 3: social life, health and wellbeing
  form.addPageBreakItem().setTitle('Social life, health and wellbeing');
  addChoice(form, 'Do you have a friend circle?', ['Yes', 'No']);
  addChoice(form, 'How would you describe your current health condition?', ['Good', 'Normal', 'Abnormal', 'Critical']);
  addChoice(form, 'Monthly living expense, excluding study costs (in local currency)',
    ['Low', 'Mid', 'High'],
    'Low = below 10,000; Mid = 10,000 to 14,999; High = 15,000 or more. ' +
    'Count food, rent, transport and personal expenses — not tuition or books.');
  addChoice(form, 'Over the past few months, have you felt depressed (persistently low mood or loss of interest)?',
    ['Yes', 'No']);

  // ---- Linked response sheet
  var sheet = SpreadsheetApp.create(title + ' (Responses)');
  form.setDestination(FormApp.DestinationType.SPREADSHEET, sheet.getId());

  Logger.log('Share this link with respondents: ' + form.getPublishedUrl());
  Logger.log('Edit the form here: ' + form.getEditUrl());
  Logger.log('Responses are collected in: ' + sheet.getUrl());
}

function addChoice(form, title, choices, help) {
  var item = form.addMultipleChoiceItem().setTitle(title).setChoiceValues(choices).setRequired(true);
  if (help) item.setHelpText(help);
  return item;
}

function addNumber(form, title, min, max, help) {
  var validation = FormApp.createTextValidation()
    .setHelpText('Please enter a number between ' + min + ' and ' + max + '.')
    .requireNumberBetween(min, max)
    .build();
  var item = form.addTextItem().setTitle(title).setValidation(validation).setRequired(true);
  if (help) item.setHelpText(help);
  return item;
}
