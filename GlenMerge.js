/**
 * GlenMerge
 * Louille Glen Benatiro
 * June 2024
 * glenbenatiro@gmail.com
 * */

// -----------------------------------------------------------------------------

// ESLint Globals
/* global Session */
/* global DriveApp */
/* global GmailApp */
/* global DocumentApp */
/* global SpreadsheetApp */

// -----------------------------------------------------------------------------

// eslint-disable-next-line vars-on-top, no-unused-vars, no-var
var GoogleMIMEType = {
  DOCS: 'application/vnd.google-apps.document',
  SHEETS: 'application/vnd.google-apps.spreadsheet',
};

// eslint-disable-next-line vars-on-top, no-unused-vars, no-var
var ColumnSelectorType = {
  SPECIFY: 'Specify',
  SELECT_COL_HEADER: 'Select Column Header',
  SELECT_COL_LETTER: 'Select Column Letter',
  SELECT_COL_NUMBER: 'Select Column Number',
};

// eslint-disable-next-line vars-on-top, no-unused-vars, no-var
var RowFilterOperator = {
  EQUAL_TO: 'Equal To',
  NOT_EQUAL_TO: 'Not Equal To',
  CONTAINS: 'Contains',
  DOES_NOT_CONTAIN: 'Does Not Contain',
  IS_EMPTY: 'Is Empty',
  IS_NOT_EMPTY: 'Is Not Empty',
  LESS_THAN: 'Less Than',
  GREATER_THAN: 'Greater Than',
};

// -----------------------------------------------------------------------------

const GLENMERGE = {
  DATA_SOURCE_SHEET_HEADER_ROW_NAMES: {
    MERGE_STATUS: 'G Merge Status',
    EMAIL_TRACKING_STATUS: 'Email Tracking Status',
    DOCUMENT_URL: 'Document URL',
    MERGE_ID: 'G Merge Id',
  },
  COLORS: {
    GLEN_DARK_BLUE: '#1b578d',
  },
  DATA_SOURCE_SHEET_HEADER_ROW: 1,
};

GLENMERGE.DEFAULT_ROW_FILTERS = [
  {
    type: ColumnSelectorType.SELECT_COL_HEADER,
    input: GLENMERGE.DATA_SOURCE_SHEET_HEADER_ROW_NAMES.MERGE_STATUS,
    operator: RowFilterOperator.IS_EMPTY,
    rowContent: '',
  },
];

// -----------------------------------------------------------------------------

function getGoogleEntityIDFromURL_(url) {
  let id = '';

  const parts = url.split(
    /^(([^:/?#]+):)?(\/\/([^/?#]*))?([^?#]*)(\?([^#]*))?(#(.*))?/,
  );

  if (url.indexOf('?id=') >= 0) {
    id = parts[6].split('=')[1].replace('&usp', '');
  } else {
    id = parts[5].split('/');

    const [sortArr] = id.sort((a, b) => b.length - a.length);

    id = sortArr;
  }

  return id;
}

function sheetColumnNumberToLetters_(number) {
  let numberCopy = number;
  let letters = '';

  while (numberCopy > 0) {
    const temp = (numberCopy - 1) % 26;
    letters = String.fromCharCode(temp + 65) + letters;
    numberCopy = (numberCopy - temp - 1) / 26;
  }

  return letters;
}

function createSheetHeaderRowObject_(sheet, row) {
  const data = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];

  const hdrObj = data.reduce((accumulator, curr, index) => {
    accumulator[curr] = {
      num: index + 1,
      let: sheetColumnNumberToLetters_(index + 1),
    };

    return accumulator;
  }, {});

  return hdrObj;
}

function sheetColumnLettersToNumber_(letters) {
  let n = 0;

  for (let p = 0; p < letters.length; p += 1) {
    n = letters[p].charCodeAt() - 64 + n * 26;
  }

  return n;
}

function convertDocsToPDF_(doc, destFolder) {
  const file = DriveApp.getFileById(doc.getId());
  const folder = destFolder ?? file.getParents().next();
  const pdf = DriveApp.createFile(doc.getAs('application/pdf'))
    .moveTo(folder)
    .setName(file.getName());

  return pdf;
}

function validateEnum(enumeration, input, enumerationName = 'enum') {
  if (Object.values(enumeration).includes(input)) {
    return input;
  }

  throw new Error(`Invalid ${enumerationName} input: ${input}`);
}

// -----------------------------------------------------------------------------

function initGlenMergeHeaderRowColumns_(sheet, row) {
  const rowData = sheet
    .getRange(row, 1, 1, sheet.getLastColumn())
    .getValues()
    .flat();

  Object.values(GLENMERGE.DATA_SOURCE_SHEET_HEADER_ROW_NAMES).forEach(
    (name) => {
      let index = rowData.indexOf(name);

      if (index < 0) {
        index = sheet.getLastColumn() + 1;
        sheet.getRange(row, index).setValue(name);
      } else {
        index += 1;
      }

      sheet
        .getRange(row, index)
        .setBackground(GLENMERGE.COLORS.GLEN_DARK_BLUE)
        .setFontColor('white');
    },
  );
}

function doesRowPassFilters_(sheetRow, sheetData, runtimeRowFilters) {
  const filterCheck = (sheetValue, operator, rowContent) => {
    switch (operator) {
      case RowFilterOperator.EQUAL_TO:
        return sheetValue === rowContent;

      case RowFilterOperator.NOT_EQUAL_TO:
        return sheetValue !== rowContent;

      case RowFilterOperator.CONTAINS:
        return sheetValue.includes(rowContent);

      case RowFilterOperator.DOES_NOT_CONTAIN:
        return !sheetValue.includes(rowContent);

      case RowFilterOperator.IS_EMPTY:
        return sheetValue.trim().length === 0;

      case RowFilterOperator.IS_NOT_EMPTY:
        return sheetValue.trim().length !== 0;

      case RowFilterOperator.LESS_THAN:
        return sheetValue < rowContent;

      case RowFilterOperator.GREATER_THAN:
        return sheetValue > rowContent;

      default:
        throw new Error(`Invalid row filter operator: ${operator}`);
    }
  };

  return runtimeRowFilters.every((filter) => {
    const sheetValue = sheetData[sheetRow - 1][filter.columnNumber - 1];

    return filterCheck(sheetValue, filter.operator, filter.rowContent);
  });
}

function getDocMergeDocsTemplateTags_(file) {
  const tags = {};

  const doc = file;
  const body = doc.getBody();
  const text = body.getText();
  const regex = /{{(.*?)}}/g;
  const matches = Array.from(text.matchAll(regex), (match) => match[1]);

  matches.forEach((x, index) => {
    if (!Object.prototype.hasOwnProperty.call(tags, x)) {
      tags[x] = [];
    }

    tags[x].push(index);
  });

  return tags;
}

function getDocMergeSheetsTemplateTags_(file) {
  const tags = {};

  const spreadsheet = file;
  const sheets = spreadsheet.getSheets();
  const regex = /{{(.*?)}}/g;

  sheets.forEach((sheet) => {
    const sheetName = sheet.getName();
    const sheetData = sheet.getDataRange().getValues();

    tags[sheetName] = {};

    sheetData.forEach((rowData, row) => {
      rowData.forEach((cellValue, col) => {
        if (typeof cellValue === 'string') {
          const matches = cellValue.match(regex);

          if (matches) {
            matches.forEach((match) => {
              tags[sheetName][match] = {
                row: row + 1,
                col: col + 1,
              };
            });
          }
        }
      });
    });
  });

  return tags;
}

function convertMergeFileToPDF_(mergeFile) {
  let pdf = null;
  const type = mergeFile.getMimeType();

  switch (type) {
    case GoogleMIMEType.DOCS: {
      const doc = DocumentApp.openByUrl(mergeFile.getUrl());
      pdf = convertDocsToPDF_(doc);

      break;
    }

    case GoogleMIMEType.SHEETS: {
      const ss = SpreadsheetApp.openByUrl(mergeFile.getUrl());
      pdf = GlenSheetsToPDF.convert(ss);
      break;
    }

    default:
      throw new Error(`Invalid document merge file type: ${type}`);
  }

  mergeFile.setTrashed(true);

  return pdf;
}

function setColSecStateObj_(input, columnSelectorType, colSecStateObj) {
  const temp = colSecStateObj;

  if (columnSelectorType === null || input === null) {
    temp.type = null;
    temp.input = null;
  } else {
    temp.type = validateEnum(ColumnSelectorType, columnSelectorType);
    temp.input = input;
  }
}

function setColSelEnableStateObj_(
  input,
  columnSelectorType,
  colSecEnableStateObj,
) {
  const temp = colSecEnableStateObj;

  if (columnSelectorType === null || input === null) {
    temp.isEnabled = false;
  } else {
    temp.isEnabled = true;
  }

  setColSecStateObj_(input, columnSelectorType, temp);
}

function getGoogleFileMIMETypeByURL_(url) {
  // Docs URL format
  // https://docs.google.com/document/d/1ZKoK393SofUvDeibTitZinwo4NONZK-jPvfiHccK9-w/edit
  // Sheets URL format
  // https://docs.google.com/spreadsheets/d/1Jxa6Gd7rt_4VnJbFhICKvrq37Ph3HS_8Vob6CXQn9Ac/edit#gid=0

  const parts = url.trim().split('/');
  const index = parts.indexOf('docs.google.com');

  if (index >= 0) {
    switch (parts[index + 1]) {
      case 'document':
        return GoogleMIMEType.DOCS;
      case 'spreadsheets':
        return GoogleMIMEType.SHEETS;
      default:
        break;
    }
  }

  throw new Error(
    `Can't identify Google file MIME type from provided URL: ${url}`,
  );
}

function getMergeFormattedDateTime_(date) {
  const year = date.getFullYear();
  const month = (date.getMonth() + 1).toString().padStart(2, '0'); // Months are zero-based
  const day = date.getDate().toString().padStart(2, '0');
  const hours = date.getHours().toString().padStart(2, '0');
  const minutes = date.getMinutes().toString().padStart(2, '0');

  // Get timezone offset in minutes and convert to hours and minutes
  const timezoneOffset = -date.getTimezoneOffset(); // In minutes, negative for positive offsets
  const offsetHours = Math.floor(Math.abs(timezoneOffset) / 60);
  const sign = timezoneOffset >= 0 ? '+' : '-';

  // Format timezone offset
  const timezoneString = `GMT${sign}${offsetHours.toString().padStart(2, '0')}`;

  const formattedDateTime = `${year}/${month}/${day} ${hours}:${minutes} ${timezoneString}`;
  return formattedDateTime;
}

function addGlenMergeColsOnSheetIfNotPresent(sheet) {
  initGlenMergeHeaderRowColumns_(sheet, GLENMERGE.DATA_SOURCE_SHEET_HEADER_ROW);
}

function getMailMergeSenderEmailAddresses() {
  return [Session.getActiveUser().getEmail(), ...GmailApp.getAliases()];
}

// -----------------------------------------------------------------------------

class GlenMerge {
  constructor() {
    this.dataSourceSheet_ = {
      sheet: null,
      headerRow: null,
      hdrObj: null,
      rowFilters: GLENMERGE.DEFAULT_ROW_FILTERS,
    };

    this.docMerge_ = {
      isEnabled: false,
      template: {
        file: null,
        type: null,
        tags: null,
      },
      destinationFolder: null,
      documentTitle: {
        input: null,
        type: null,
      },
      sharedTo: {
        isEnabled: false,
        input: null,
        type: null,
      },
      mergeAsPDF: false,
    };

    this.mailMerge_ = {
      isEnabled: false,
      template: {
        file: null,
        tags: null,
        folder: null,
      },
      sender: 'me',
      subject: {
        input: null,
        type: null,
      },
      recipients: {
        input: null,
        type: null,
      },
      cc: {
        isEnabled: false,
        input: null,
        type: null,
      },
      bcc: {
        isEnabled: false,
        input: null,
        type: null,
      },
      sendAsAttachment: true,
    };
  }

  // ---------------------------------------------------------------------------

  getStateInfo() {
    return {
      dataSourceSheet: { ...this.dataSourceSheet_ },
      docMerge: { ...this.docMerge_ },
      mailMerge: { ...this.mailMerge_ },
    };
  }

  getColSecStateObjSheetCol_(colSecStateObj) {
    const { input, type } = colSecStateObj;

    switch (type) {
      case ColumnSelectorType.SPECIFY:
        throw new Error(
          `ColumnSelectorType ${ColumnSelectorType.SPECIFY} is not a valid argument for this function.`,
        );

      case ColumnSelectorType.SELECT_COL_HEADER:
        return this.getDataSourceSheetHeaderRowObject()[input].num;

      case ColumnSelectorType.SELECT_COL_LETTER:
        return sheetColumnLettersToNumber_(input);

      case ColumnSelectorType.SELECT_COL_NUMBER:
        return input;

      default:
        throw new Error(`Invalid columnSelectorType parameter: ${type}`);
    }
  }

  getColSecStateObjSheetData_(colSecStateObj, sheetRow, sheetData) {
    const { input, type } = colSecStateObj;

    switch (type) {
      case ColumnSelectorType.SPECIFY:
        return input;

      case ColumnSelectorType.SELECT_COL_HEADER:
      case ColumnSelectorType.SELECT_COL_LETTER:
      case ColumnSelectorType.SELECT_COL_NUMBER: {
        const col = this.getColSecStateObjSheetCol_(colSecStateObj);

        return sheetData[sheetRow - 1][col - 1];
      }

      default:
        throw new Error(`Invalid columnSelectorType parameter: ${type}`);
    }
  }

  // checks
  isReadyToRun() {
    const errors = [];

    if (!this.dataSourceSheet_.sheet) {
      errors.push(`No data source sheet set.`);
    }

    if (this.docMerge_.isEnabled) {
      if (!this.docMerge_.template.file) {
        errors.push(`No document merge template file set.`);
      }

      if (!this.docMerge_.template.type) {
        errors.push(`No document merge template type set.`);
      }

      if (!this.docMerge_.template.tags) {
        errors.push(`No document merge template tags set.`);
      }

      if (!this.docMerge_.destinationFolder) {
        errors.push(`No document merge destination folder set.`);
      }

      if (!this.docMerge_.documentTitle.input) {
        errors.push(`No document merge document title column input set.`);
      }

      if (!this.docMerge_.documentTitle.type) {
        errors.push(`No document merge document title column type set.`);
      }
    }

    if (this.mailMerge_.isEnabled) {
      if (!this.mailMerge_.template.file) {
        errors.push(`No mail merge template set.`);
      }

      if (!this.mailMerge_.sender) {
        errors.push(`No mail merge email sender set.`);
      }

      if (!this.mailMerge_.subject.input) {
        errors.push(`No mail merge email subject column input set.`);
      }

      if (!this.mailMerge_.subject.type) {
        errors.push(`No mail merge email subject column type set.`);
      }

      if (!this.mailMerge_.recipients.input) {
        errors.push(`No mail merge email recipients column input set.`);
      }

      if (!this.mailMerge_.recipients.type) {
        errors.push(`No mail merge email recipients column type set.`);
      }
    }

    return {
      state: errors.length === 0,
      errors,
    };
  }

  // data source sheet
  getDataSourceSheetHeaderRowObject() {
    return this.dataSourceSheet_.hdrObj;
  }

  setDataSourceSheet(sheet) {
    this.dataSourceSheet_.sheet = sheet;
    this.dataSourceSheet_.headerRow = GLENMERGE.DATA_SOURCE_SHEET_HEADER_ROW;

    initGlenMergeHeaderRowColumns_(
      this.dataSourceSheet_.sheet,
      this.dataSourceSheet_.headerRow,
    );

    this.dataSourceSheet_.hdrObj = createSheetHeaderRowObject_(
      this.dataSourceSheet_.sheet,
      this.dataSourceSheet_.headerRow,
    );
  }

  // mail merge
  enableDocMerge(bool) {
    this.docMerge_.isEnabled = bool;
  }

  setDocMergeTemplateByURL(url) {
    const type = getGoogleFileMIMETypeByURL_(url);
    const { template } = this.docMerge_;

    switch (type) {
      case GoogleMIMEType.DOCS:
        template.file = DocumentApp.openByUrl(url);
        template.tags = getDocMergeDocsTemplateTags_(template.file);
        break;

      case GoogleMIMEType.SHEETS:
        template.file = SpreadsheetApp.openByUrl(url);
        template.tags = getDocMergeSheetsTemplateTags_(template.file);
        break;

      default:
        throw new Error(
          `Invalid document merge template file provided. Only Docs or Sheets files are allowed.`,
        );
    }

    template.type = type;
  }

  setDocMergeTemplate(file) {
    return this.setDocMergeTemplateByURL(file.getUrl());
  }

  setDocMergeOutputFolder(folder) {
    this.docMerge_.destinationFolder = folder;
  }

  setDocMergeOutputFolderByURL(url) {
    this.setDocMergeOutputFolder(
      DriveApp.getFolderById(getGoogleEntityIDFromURL_(url)),
    );
  }

  setDocMergeTitle(input, columnSelectorType) {
    setColSecStateObj_(input, columnSelectorType, this.docMerge_.documentTitle);
  }

  setDocMergeSharedTo(input, columnSelectorType) {
    setColSelEnableStateObj_(
      input,
      columnSelectorType,
      this.docMerge_.sharedTo,
    );
  }

  setDocMergeAsPDF(bool) {
    this.docMerge_.mergeAsPDF = bool;
  }

  doDocMergeDocsTemplate_(context) {
    const {
      sheetRow,
      sheetData,
      docTemplateFile,
      docTemplateTags,
      hdrObj,
      destinationFolder,
    } = context;

    const templateFile = DriveApp.getFileById(docTemplateFile.getId());
    const mergeFile = templateFile.makeCopy('%temp%', destinationFolder);
    const mergeDocs = DocumentApp.openByUrl(mergeFile.getUrl());
    const body = mergeDocs.getBody();

    Object.entries(hdrObj).forEach(([colName, colObj]) => {
      if (Object.prototype.hasOwnProperty.call(docTemplateTags, colName)) {
        const searchPattern = `{{${colName}}}`;
        const replacement = sheetData[sheetRow - 1][colObj.num - 1];

        if (body.findText(searchPattern)) {
          body.replaceText(searchPattern, replacement);
        } else {
          const checkElement = (element) => {
            if (element.getType() === DocumentApp.ElementType.TEXT) {
              if (element.getText().includes(`{{${colName}}`)) {
                element.setText(
                  element.getText().replace(searchPattern, replacement),
                );
              }
            }
          };

          const iterateElements = (element) => {
            checkElement(element);

            if (element.getNumChildren) {
              const numChildren = element.getNumChildren();

              for (let i = 0; i < numChildren; i += 1) {
                iterateElements(element.getChild(i));
              }
            }
          };

          iterateElements(body);
        }
      }
    });

    mergeDocs.saveAndClose();

    const docTitle = this.getColSecStateObjSheetData_(
      this.docMerge_.documentTitle,
      sheetRow,
      sheetData,
    );

    mergeDocs.setName(docTitle);

    return mergeDocs;
  }

  doDocMergeSheetsTemplate_(sheetRow, sheetData) {
    const sheetsTemplateFile = this.docMerge_.template.file;
    const sheetsTemplateTags = this.docMerge_.template.tags;
    const hdrObj = this.getDataSourceSheetHeaderRowObject();
    const { destinationFolder } = this.docMerge_;

    const templateFile = DriveApp.getFileById(sheetsTemplateFile.getId());
    const mergeFile = templateFile.makeCopy('%temp%', destinationFolder);
    const mergeSpreadsheet = SpreadsheetApp.openByUrl(mergeFile.getUrl());

    Object.keys(sheetsTemplateTags).forEach((sheetName) => {
      const mergeSheet = mergeSpreadsheet.getSheetByName(sheetName);
      const mergeSheetData = mergeSheet.getDataRange().getValues();

      Object.entries(sheetsTemplateTags[sheetName]).forEach(
        ([key, { row, col }]) => {
          const sourceKey = key.substring(2, key.length - 2);

          if (hdrObj[sourceKey]) {
            const sourceValue =
              sheetData[sheetRow - 1][hdrObj[sourceKey].num - 1];

            const cellValue = mergeSheetData[row - 1][col - 1];
            mergeSheetData[row - 1][col - 1] = cellValue.replace(
              key,
              sourceValue,
            );
          }
        },
      );

      mergeSheet
        .getRange(1, 1, mergeSheetData.length, mergeSheetData[0].length)
        .setValues(mergeSheetData);
    });

    const title = this.getColSecStateObjSheetData_(
      this.docMerge_.documentTitle,
      sheetRow,
      sheetData,
    );

    mergeSpreadsheet.setName(title);

    return mergeSpreadsheet;
  }

  // mail merge
  enableMailMerge(bool) {
    this.mailMerge_.isEnabled = bool;
  }

  setMailMergeTemplateByURL(url) {
    const type = getGoogleFileMIMETypeByURL_(url);

    if (type !== GoogleMIMEType.DOCS) {
      throw new Error(
        `Invalid mail merge template file provided. Only Docs files are allowed.`,
      );
    } else {
      const doc = DocumentApp.openByUrl(url);
      const file = DriveApp.getFileById(doc.getId());
      const folder = file.getParents().next();

      this.mailMerge_.template.file = doc;
      this.mailMerge_.template.tags = getDocMergeDocsTemplateTags_(doc);
      this.mailMerge_.template.folder = folder;
    }
  }

  setMailMergeSender(email = 'me') {
    this.mailMerge_.sender =
      email === 'me' ? Session.getActiveUser().getEmail() : email;
  }

  setMailMergeSubject(input, columnSelectorType) {
    setColSecStateObj_(input, columnSelectorType, this.mailMerge_.subject);
  }

  setMailMergeRecipient(input, columnSelectorType) {
    setColSecStateObj_(input, columnSelectorType, this.mailMerge_.recipients);
  }

  setMailMergeCC(input, columnSelectorType) {
    setColSelEnableStateObj_(input, columnSelectorType, this.mailMerge_.cc);
  }

  setMailMergeBCC(input, columnSelectorType) {
    setColSelEnableStateObj_(input, columnSelectorType, this.mailMerge_.bcc);
  }

  setMailMergeSendAsAttachment(bool) {
    this.mailMerge_.sendAsAttachment = bool;
  }

  getMailMergeBody_(sheetRow, sheetData) {
    const doc = this.doDocMergeDocsTemplate_({
      sheetRow,
      sheetData,
      docTemplateFile: this.mailMerge_.template.file,
      docTemplateTags: this.mailMerge_.template.tags,
      hdrObj: this.getDataSourceSheetHeaderRowObject(),
      destinationFolder: this.mailMerge_.template.folder,
    });

    const text = doc.getBody().getText();

    DriveApp.getFileById(doc.getId()).setTrashed(true);

    return text;
  }

  doMailMerge_(sheetRow, sheetData, mergeFile) {
    const body = this.getMailMergeBody_(sheetRow, sheetData);

    const recipient = this.getColSecStateObjSheetData_(
      this.mailMerge_.recipients,
      sheetRow,
      sheetData,
    );
    const subject = this.getColSecStateObjSheetData_(
      this.mailMerge_.subject,
      sheetRow,
      sheetData,
    );

    const options = {};

    if (this.mailMerge_.cc.isEnabled) {
      options.cc = this.getColSecStateObjSheetData_(
        this.mailMerge_.cc,
        sheetRow,
        sheetData,
      );
    }

    if (this.mailMerge_.bcc.isEnabled) {
      options.bcc = this.getColSecStateObjSheetData_(
        this.mailMerge_.bcc,
        sheetRow,
        sheetData,
      );
    }

    if (this.mailMerge_.sendAsAttachment && mergeFile) {
      options.attachments = [mergeFile];
    } else {
      // to-do: add code here
    }

    options.from =
      this.mailMerge_.sender === 'me'
        ? Session.getActiveUser().getEmail()
        : this.mailMerge_.sender;

    GmailApp.sendEmail(recipient, subject, body, options);
  }

  // row filters
  addRowFilter(input, columnSelectorType, operator, rowContent) {
    const rowFilter = {
      type: columnSelectorType,
      input,
      operator,
      rowContent,
    };

    this.dataSourceSheet_.rowFilters.push(rowFilter);
  }

  addRowFilters(filters) {
    filters.forEach((filter) => {
      this.addRowFilter(
        filter.input,
        filter.type,
        filter.operator,
        filter.rowContent,
      );
    });
  }

  getRowFilters() {
    return this.dataSourceSheet_.rowFilters;
  }

  resetRowFilters() {
    this.dataSourceSheet_.rowFilters = GLENMERGE.DEFAULT_ROW_FILTERS;
  }

  getRuntimeRowFilters() {
    return this.dataSourceSheet_.rowFilters.reduce((accumulator, curr) => {
      const obj = {
        columnNumber: this.getColSecStateObjSheetCol_(curr),
        operator: curr.operator,
        rowContent: curr.rowContent,
      };

      accumulator.push(obj);

      return accumulator;
    }, []);
  }

  // run
  createMergeFile_(sheetRow, sheetData) {
    let file = null;

    const { type } = this.docMerge_.template;

    switch (type) {
      case GoogleMIMEType.DOCS:
        file = this.doDocMergeDocsTemplate_({
          sheetRow,
          sheetData,
          docTemplateFile: this.docMerge_.template.file,
          docTemplateTags: this.docMerge_.template.tags,
          hdrObj: this.getDataSourceSheetHeaderRowObject(),
          destinationFolder: this.docMerge_.destinationFolder,
        });
        break;

      case GoogleMIMEType.SHEETS:
        file = this.doDocMergeSheetsTemplate_(sheetRow, sheetData);
        break;

      default:
        throw new Error(`Invalid document merge file type: ${type}`);
    }

    return DriveApp.getFileById(file.getId());
  }

  writeDoneTimestamp_(sheetRow, sheet, mergeFile) {
    const mergeStatusCol =
      this.getDataSourceSheetHeaderRowObject()[
        GLENMERGE.DATA_SOURCE_SHEET_HEADER_ROW_NAMES.MERGE_STATUS
      ].num;
    const docURLCol =
      this.getDataSourceSheetHeaderRowObject()[
        GLENMERGE.DATA_SOURCE_SHEET_HEADER_ROW_NAMES.DOCUMENT_URL
      ].num;

    sheet
      .getRange(sheetRow, mergeStatusCol)
      .setValue(
        `Done\nProcessed on: ${getMergeFormattedDateTime_(new Date())}`,
      );

    if (mergeFile) {
      sheet.getRange(sheetRow, docURLCol).setValue(mergeFile.getUrl());
    }

    SpreadsheetApp.flush();
  }

  doMerge_() {
    console.log(`Running merge...`);

    const { sheet } = this.dataSourceSheet_;
    const lastRow = sheet.getLastRow();
    const sheetData = sheet.getDataRange().getDisplayValues();
    const runtimeRowFilters = this.getRuntimeRowFilters();

    for (
      let sheetRow = this.dataSourceSheet_.headerRow + 1;
      sheetRow <= lastRow;
      sheetRow += 1
    ) {
      console.log(`\tProcessing row ${sheetRow}...`);

      if (!doesRowPassFilters_(sheetRow, sheetData, runtimeRowFilters)) {
        console.log(`\t\tRow does not pass filters. Skipping.`);
      } else {
        let mergeFile = null;

        if (this.docMerge_.isEnabled) {
          console.log(`\t\tRunning doc merge...`);

          mergeFile = this.createMergeFile_(sheetRow, sheetData);

          if (this.docMerge_.mergeAsPDF) {
            mergeFile = convertMergeFileToPDF_(mergeFile);
          }
        }

        if (this.mailMerge_.isEnabled) {
          console.log(`\t\tRunning mail merge...`);

          this.doMailMerge_(sheetRow, sheetData, mergeFile);
        }

        this.writeDoneTimestamp_(sheetRow, sheet, mergeFile);

        console.log(`\t\tRow merge done.`);
      }
    }

    console.log(`Done running merge.`);
  }

  run() {
    const check = this.isReadyToRun();

    if (!check.state) {
      const str = `\n${check.errors.join('\n')}`;

      throw new Error(str);
    } else {
      this.doMerge_();
    }
  }
}

// =============================================================================

function createInstance() {
  return new GlenMerge();
}

// EOF
