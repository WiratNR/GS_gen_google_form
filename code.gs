const CONFIG = {
  FILE_NAME: "EX-sample",
  HEAD_TITLE: "แบบทดสอบก่อนเก็บคะแนน",
  DESCRIPTION: "คำอธิบาย ข้อสอบมีทั้งหมด ๑๐๐ ข้อ (๑๐๐ คะแนน) ให้นักเรียนเลือกคำตอบที่ถูกต้องที่สุดเพียงคำตอบเดียว",
  SHEET_NAME: "Sheet1",
  FOLDER_ID: "xxxxx",
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('สร้างฟอร์ม')
    .addItem('สร้างแบบทดสอบบน Google Form', 'createGGForm')
  menu.addToUi();
}

function createGGForm() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  const [, ...values] = sheet
    .getDataRange()
    .getDisplayValues()
    .filter((r) => r.join("") != "");
  const obj = values.map(([a, b, c]) => {
    const answers = b
      .split("\n")
      .map((e) => e.trim())
      .filter(String);
    const correct = c
      .split("\n")
      .map((e) => e.trim())
      .filter(String);
    return {
      question: a,
      answers,
      correct,
      point: 1,
      type: correct.length == 1 ? "addMultipleChoiceItem" : "addCheckboxItem",
    };
  });

  // Create the form
  const form = FormApp.create(CONFIG.FILE_NAME)
    .setIsQuiz(true)
    .setTitle(CONFIG.HEAD_TITLE)

  obj.forEach(({ question, answers, correct, point, type }) => {
    const choice = form[type]();
    const choices = answers.map((e) =>
      choice.createChoice(e, correct.includes(e) ? true : false)
    );
    choice.setTitle(question).setPoints(point).setChoices(choices);
  });

  form.setDescription(CONFIG.DESCRIPTION)

  const formFile = DriveApp.getFileById(form.getId());
  formFile.setTrashed(true);

  const newId = formFile.makeCopy().setName(CONFIG.FILE_NAME).getId()
  DriveApp.getFileById(newId).moveTo(DriveApp.getFolderById(CONFIG.FOLDER_ID));
}
