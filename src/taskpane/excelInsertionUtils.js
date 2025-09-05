/* global Excel console */

/**
 * Insert narrative text into the current worksheet
 * @param {string} text - The narrative text to insert
 * @returns {Promise<Object>} - Result of the operation
 */
export async function insertNarrativeInCurrentSheet(text) {
  try {
    await Excel.run(async (context) => {
      // Get the current worksheet
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      
      // Get the current selection
      const range = context.workbook.getSelectedRange();
      range.load("address");
      await context.sync();
      
      // If no selection, use cell A1
      let targetRange;
      if (!range.address) {
        targetRange = sheet.getRange("A1");
      } else {
        targetRange = range;
      }
      
      // Insert the text
      targetRange.values = [[text]];
      targetRange.format.wrapText = true;
      targetRange.format.autofitColumns();
      
      await context.sync();
    });
    
    return { success: true };
  } catch (error) {
    console.error("Error inserting narrative in current sheet:", error);
    return { success: false, error: error.message };
  }
}

/**
 * Insert narrative text into a new worksheet
 * @param {string} text - The narrative text to insert
 * @returns {Promise<Object>} - Result of the operation
 */
export async function insertNarrativeInNewSheet(text) {
  try {
    await Excel.run(async (context) => {
      // Add a new worksheet
      const newSheet = context.workbook.worksheets.add();
      
      // Get cell A1 in the new sheet
      const range = newSheet.getRange("A1");
      
      // Insert the text
      range.values = [[text]];
      range.format.wrapText = true;
      range.format.autofitColumns();
      
      // Activate the new sheet
      newSheet.activate();
      
      await context.sync();
    });
    
    return { success: true };
  } catch (error) {
    console.error("Error inserting narrative in new sheet:", error);
    return { success: false, error: error.message };
  }
}

/**
 * Insert a table into the current worksheet
 * @param {Object} tableData - The table data to insert
 * @returns {Promise<Object>} - Result of the operation
 */
export async function insertTableInCurrentSheet(tableData) {
  try {
    await Excel.run(async (context) => {
      // Get the current worksheet
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      
      // Get the current selection
      const selection = context.workbook.getSelectedRange();
      selection.load("address");
      await context.sync();
      
      // If no selection, use cell A1
      let startCell;
      if (!selection.address) {
        startCell = sheet.getRange("A1");
      } else {
        startCell = selection;
      }
      
      // Determine the dimensions of the table
      const rowCount = tableData.rows.length + 2; // Headers + data rows + totals
      const columnCount = tableData.headers.length;
      
      // Get the range for the entire table
      const tableRange = startCell.getResizedRange(rowCount - 1, columnCount - 1);
      
      // Create the data array for the table
      const tableValues = [
        tableData.headers, // Headers row
        ...tableData.rows, // Data rows
        tableData.totals   // Totals row
      ];
      
      // Insert the table data
      tableRange.values = tableValues;
      
      // Format the table
      const headerRange = tableRange.getRow(0);
      headerRange.format.font.bold = true;
      
      const totalsRange = tableRange.getRow(tableValues.length - 1);
      totalsRange.format.font.bold = true;
      
      // Apply number formatting if needed
      if (tableData.formatting && tableData.formatting.numberFormat === "currency") {
        for (let i = 1; i < columnCount; i++) {
          const numberColumn = tableRange.getColumn(i);
          numberColumn.numberFormat = [["$#,##0.00"]];
        }
      }
      
      // Autofit columns
      tableRange.format.autofitColumns();
      
      await context.sync();
    });
    
    return { success: true };
  } catch (error) {
    console.error("Error inserting table in current sheet:", error);
    return { success: false, error: error.message };
  }
}

/**
 * Insert a table into a new worksheet
 * @param {Object} tableData - The table data to insert
 * @returns {Promise<Object>} - Result of the operation
 */
export async function insertTableInNewSheet(tableData) {
  try {
    await Excel.run(async (context) => {
      // Add a new worksheet
      const newSheet = context.workbook.worksheets.add();
      
      // Get cell A1 in the new sheet
      const startCell = newSheet.getRange("A1");
      
      // Determine the dimensions of the table
      const rowCount = tableData.rows.length + 2; // Headers + data rows + totals
      const columnCount = tableData.headers.length;
      
      // Get the range for the entire table
      const tableRange = startCell.getResizedRange(rowCount - 1, columnCount - 1);
      
      // Create the data array for the table
      const tableValues = [
        tableData.headers, // Headers row
        ...tableData.rows, // Data rows
        tableData.totals   // Totals row
      ];
      
      // Insert the table data
      tableRange.values = tableValues;
      
      // Format the table
      const headerRange = tableRange.getRow(0);
      headerRange.format.font.bold = true;
      
      const totalsRange = tableRange.getRow(tableValues.length - 1);
      totalsRange.format.font.bold = true;
      
      // Apply number formatting if needed
      if (tableData.formatting && tableData.formatting.numberFormat === "currency") {
        for (let i = 1; i < columnCount; i++) {
          const numberColumn = tableRange.getColumn(i);
          numberColumn.numberFormat = [["$#,##0.00"]];
        }
      }
      
      // Autofit columns
      tableRange.format.autofitColumns();
      
      // Activate the new sheet
      newSheet.activate();
      
      await context.sync();
    });
    
    return { success: true };
  } catch (error) {
    console.error("Error inserting table in new sheet:", error);
    return { success: false, error: error.message };
  }
}

/**
 * Copy text to clipboard
 * @param {string} text - The text to copy to clipboard
 * @returns {Promise<Object>} - Result of the operation
 */
export async function copyToClipboard(text) {
  try {
    // Use the Clipboard API if available
    if (navigator.clipboard && navigator.clipboard.writeText) {
      await navigator.clipboard.writeText(text);
      return { success: true };
    } else {
      // Fallback method for browsers that don't support the Clipboard API
      const textArea = document.createElement("textarea");
      textArea.value = text;
      
      // Make the textarea out of viewport
      textArea.style.position = "fixed";
      textArea.style.left = "-999999px";
      textArea.style.top = "-999999px";
      document.body.appendChild(textArea);
      
      textArea.focus();
      textArea.select();
      
      const successful = document.execCommand("copy");
      document.body.removeChild(textArea);
      
      if (successful) {
        return { success: true };
      } else {
        return { 
          success: false, 
          error: "Unable to copy to clipboard using fallback method" 
        };
      }
    }
  } catch (error) {
    console.error("Error copying to clipboard:", error);
    return { success: false, error: error.message };
  }
}

/**
 * Format table data for clipboard
 * @param {Object} tableData - The table data to format
 * @returns {string} - Formatted table text
 */
export function formatTableForClipboard(tableData) {
  // Create a string representation of the table
  let tableText = "";
  
  // Add headers
  tableText += tableData.headers.join("\t") + "\n";
  
  // Add rows
  tableData.rows.forEach(row => {
    tableText += row.join("\t") + "\n";
  });
  
  // Add totals
  tableText += tableData.totals.join("\t");
  
  return tableText;
}
