  // Replace with the ID of your specific spreadsheet
  const TARGET_SPREADSHEET_ID = '1D3Y8DCMOq7b3X7FYuaFqYuGHfbgU_WtUiX0IYU4XHGI';

  // Generates a unique GUID
  function generateGUID() {
    return Utilities.getUuid();
  }

  // Ensures the "Event Log" sheet exists, creates it if not
  function ensureEventLogSheetExists() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = 'Event Log';
    let eventLogSheet = spreadsheet.getSheetByName(sheetName);

    // Check if the "Event Log" sheet exists
    if (!eventLogSheet) {
      // Create the "Event Log" sheet
      eventLogSheet = spreadsheet.insertSheet(sheetName);

      // Set up headers for the audit log
      const headers = ['Timestamp', 'User', 'Edited Row', 'Edited Column', 'New Value'];
      eventLogSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    }
  }

  // Triggered when a new form response is submitted
  function onFormSubmit(e) {
    if (e.source.getId() !== TARGET_SPREADSHEET_ID) return; // Check if the event is triggered from the correct spreadsheet
    ensureEventLogSheetExists();
    const formResponsesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Form Responses');
    const shadowerAdminsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Shadower Admins');
    const lastRow = formResponsesSheet.getLastRow();

    // Get the latest row data from Form Responses
    const rowData = formResponsesSheet.getRange(lastRow, 1, 1, formResponsesSheet.getLastColumn()).getValues()[0];

    // Add Created (EST) as the current date/time and a new GUID for ID
    const newId = generateGUID();

    // Append the row to Shadower Admins starting from row 3
    shadowerAdminsSheet.appendRow([...rowData, newId, '', '', '', '', false]);
  }

  // ******************************************************
  // Triggered when any cell is edited in the spreadsheet
  // ******************************************************

  function onEdit(e) {
    if (e.source.getId() !== TARGET_SPREADSHEET_ID) return; // Check if the event is triggered from the correct spreadsheet
    ensureEventLogSheetExists();
    const editedRow = e.range.getRow();

    // Skip edits in rows 1 and 2
    if (editedRow < 3) return;

    logEditEvent(e);
    handleCategoryEdit(e);
    onCategoryEdit(e);
  }

  // Logs edits to the "Event Log" sheet
  function logEditEvent(e) {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const eventLogSheet = spreadsheet.getSheetByName('Event Log');
    const editedRange = e.range;
    const editedRow = editedRange.getRow();
    const editedCol = editedRange.getColumn();
    const editedValue = e.value;
    const user = Session.getActiveUser().getEmail();
    const timestamp = new Date();

    // Get the name of the edited sheet
    const editedSheet = e.source.getActiveSheet().getName();

    // Skip logging for specific sheets and header rows
    if (editedSheet !== 'Form Responses' && editedSheet !== 'Event Log' && editedRow >= 3) {
      const sheetId = spreadsheet.getId();
      const rowLink = generateRowLink(sheetId, editedSheet, editedRow);
      
      // Log the event with timestamp, user, edited row, column, and value, and a link to the edited row
      eventLogSheet.appendRow([timestamp, user, editedRow, editedCol, editedValue, rowLink]);
    }
  }

  // Generates a hyperlink to the specific row in the sheet
  function generateRowLink(sheetId, sheetName, row) {
    const link = `https://docs.google.com/spreadsheets/d/${sheetId}/edit#gid=${getSheetGid(sheetName)}&range=A${row}`;
    return `=HYPERLINK("${link}", "Go to Row ${row}")`;
  }

  // Gets the GID (sheet ID) of the specified sheet name
  function getSheetGid(sheetName) {
    const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    for (let sheet of sheets) {
      if (sheet.getName() === sheetName) {
        return sheet.getSheetId();
      }
    }
    return null; // Return null if the sheet name doesn't match
  }

  // Handles modifications in "Shadower Admins" and category sheets
  function handleCategoryEdit(e) {
    const shadowerAdminsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Shadower Admins');
    const editedRange = e.range;
    const editedRow = editedRange.getRow();
    const editedCol = editedRange.getColumn();
    const categoryCol = 9; // "Category" is in column I
    const recordIdCol = 7; // ID is in column I (index 7)

    if (e.source.getActiveSheet().getName() === 'Shadower Admins' && editedRow >= 3) {
      const rowData = shadowerAdminsSheet.getRange(editedRow, 1, 1, shadowerAdminsSheet.getLastColumn()).getValues()[0];
      const oldCategory = shadowerAdminsSheet.getRange(editedRow, categoryCol).getValue();
      const newCategory = editedCol === categoryCol ? rowData[categoryCol - 1] : oldCategory;

      Logger.log(`Old Category: ${oldCategory}, New Category: ${newCategory}`);

      // Handle creation of new category sheet if needed
      if (editedCol === categoryCol || oldCategory !== newCategory) {
        let categorySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(newCategory);

        if (!categorySheet) {
          categorySheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(newCategory);

          const headerRange = shadowerAdminsSheet.getRange(1, 1, 2, shadowerAdminsSheet.getLastColumn());
          const headerValues = headerRange.getValues();
          const headerFormats = headerRange.getTextStyles();
          const headerValidations = headerRange.getDataValidations();

          categorySheet.getRange(1, 1, 2, shadowerAdminsSheet.getLastColumn()).setValues(headerValues);
          categorySheet.getRange(1, 1, 2, shadowerAdminsSheet.getLastColumn()).setTextStyles(headerFormats);
          categorySheet.getRange(1, 1, 2, shadowerAdminsSheet.getLastColumn()).setDataValidations(headerValidations);
        }
      }

      // Handle updates to the category sheet
      if (oldCategory) {
        const oldCategorySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(oldCategory);

        if (oldCategorySheet) {
          const oldRow = findRowInCategorySheet(oldCategorySheet, rowData[recordIdCol - 1]); 

          if (oldRow) {
            Logger.log(`Deleting row ${oldRow} from old category sheet ${oldCategory}`);
            oldCategorySheet.deleteRow(oldRow);
          }
        }
      }

      if (newCategory) {
        let categorySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(newCategory);
        const rowInCategory = findRowInCategorySheet(categorySheet, rowData[recordIdCol - 1]);

        if (rowInCategory) {
          Logger.log(`Updating row ${rowInCategory} in new category sheet ${newCategory}`);
          categorySheet.getRange(rowInCategory, 1, 1, rowData.length).setValues([rowData]);
        } else {
          Logger.log(`Appending row to new category sheet ${newCategory}`);
          categorySheet.appendRow(rowData);
        }
      }

      // Update "Shadower Admins" with any changes
      shadowerAdminsSheet.getRange(editedRow, 1, 1, rowData.length).setValues([rowData]);
    }
  }

  // Finds the row in a category sheet by ID
  function findRowInCategorySheet(sheet, id) {
    const recordIdCol = 7; // ID is in column G (index 6)
    const data = sheet.getDataRange().getValues();
    for (let i = 2; i < data.length; i++) { // Start searching from row 3
      if (data[i][recordIdCol-1] === id) {
        return i + 1;
      }
    }
    return null;
  }

  // Syncs edits made in category sheets back to "Shadower Admins"
  function onCategoryEdit(e) {
    const editedSheet = e.source.getActiveSheet();
    const shadowerAdminsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Shadower Admins');

    // Skip if the edit is not in a category sheet
    if (['Shadower Admins', 'Form Responses', 'Event Log'].includes(editedSheet.getName())) return;

    const editedRange = e.range;
    const editedRow = editedRange.getRow();
    if (editedRow < 3) return; // Skip header rows

    const rowData = editedSheet.getRange(editedRow, 1, 1, editedSheet.getLastColumn()).getValues()[0];
    const rowId = rowData[8]; // ID is in the 9th column (I index 8)

    // Find the corresponding row in "Shadower Admins"
    const shadowerAdminsData = shadowerAdminsSheet.getDataRange().getValues();
    const rowToUpdate = shadowerAdminsData.findIndex((row, index) => index >= 2 && row[8] === rowId);

    if (rowToUpdate !== -1) {
      // Update the corresponding row in "Shadower Admins"
      shadowerAdminsSheet.getRange(rowToUpdate + 1, 1, 1, rowData.length).setValues([rowData]);
    }
  }

  // Function to backfill existing rows without an ID and creation date
  function backfillMissingData() {
    if (SpreadsheetApp.getActiveSpreadsheet().getId() !== TARGET_SPREADSHEET_ID) return; // Verifica que el ID de la hoja de cálculo coincida
    const shadowerAdminsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Shadower Admins');
    const dataRange = shadowerAdminsSheet.getRange(3, 1, shadowerAdminsSheet.getLastRow() - 2, shadowerAdminsSheet.getLastColumn());
    const data = dataRange.getValues();
    const idUpdates = [];
    const urlUpdates = [];
    const recordIdCol = 7; // ID is in column I (index 7)

    data.forEach((row, index) => {
      const rowIndex = index + 3; // Ajuste para la posición real de la fila
      const id = row[recordIdCol - 1]; // Columna I (ID)
      const createdDate = row[0]; // Columna A (Created EST)
      const name = row[1]; // Columna B (Admin name)
      const questionURL = row[3]; // Columna D (URL de la pregunta)

      if (!name) {
        return;
      }

      // Backfill de ID si está vacío
      if (!id) {
        const newId = generateGUID();
        idUpdates.push([rowIndex, 9, newId]); // Guarda la fila y columna para la actualización del ID
      }

      // Backfill de Created Date si está vacío
      if (!createdDate) {
        const newCreatedDate = new Date();
        idUpdates.push([rowIndex, 1, newCreatedDate]); // Guarda la fila y columna para la actualización del Created Date
      }

      // Si la URL de la pregunta está vacía pero la Columna 3 tiene un hipervínculo
      if (!questionURL && shadowerAdminsSheet.getRange(rowIndex, 3).getRichTextValue()) {
        const richTextValue = shadowerAdminsSheet.getRange(rowIndex, 3).getRichTextValue();
        const url = richTextValue.getLinkUrl(); // Obtiene la URL del hipervínculo
        const plainText = richTextValue.getText(); // Obtiene el texto plano

        // Actualiza el texto y la URL en las columnas correspondientes
        if (url) {
          urlUpdates.push([rowIndex, 4, url]); // Actualiza la Columna D con la URL
          shadowerAdminsSheet.getRange(rowIndex, 3).setValue(plainText); // Actualiza la Columna C con solo el texto
        }
      }
    });

    // Actualiza los IDs y Created Dates faltantes
    idUpdates.forEach(([rowIndex, colIndex, value]) => {
      shadowerAdminsSheet.getRange(rowIndex, colIndex).setValue(value);
    });

    // Actualiza las URLs en la Columna 4
    urlUpdates.forEach(([rowIndex, colIndex, url]) => {
      shadowerAdminsSheet.getRange(rowIndex, colIndex).setValue(url);
    });
  }
