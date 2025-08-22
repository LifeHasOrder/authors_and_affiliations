/**
 * Author and Affiliation Manager for Google Docs and Sheets
 * Last updated: 2025-08-21
 * @author LifeHasOrder
 */

// Global variables for author data
let authorData = [];
let affiliationData = [];

/**
 * Creates custom menu items when the document is opened.
 */
function onOpen() {
  let ui;
  try {
    // Try to get DocumentApp UI (for Google Docs)
    ui = DocumentApp.getUi();
  } catch (e) {
    try {
      // If that fails, try SpreadsheetApp UI (for Google Sheets)
      ui = SpreadsheetApp.getUi();
    } catch (e2) {
      Logger.log('Unable to create UI menu');
      return;
    }
  }
  
  ui.createMenu('Author Manager')
    .addItem('Manage Authors & Affiliations', 'showAuthorManager')
    .addSeparator()
    .addItem('Insert Author List', 'insertAuthorList')
    .addItem('Insert Affiliation List', 'insertAffiliationList')
    .addItem('Insert Full Citation Block', 'insertFullCitation')
    .addSeparator()
    .addItem('Export to New Sheet', 'exportToSheet')
    .addItem('Update Linked Sheet', 'updateLinkedSheet')
    .addSeparator()
    .addItem('Clear All Data', 'clearAllData')
    .addToUi();
}

/**
 * Shows the main author management interface.
 */
function showAuthorManager() {
  const html = HtmlService.createHtmlOutputFromFile('author_manager_v3')
    .setWidth(800)
    .setHeight(600);
    
  let ui;
  try {
    ui = DocumentApp.getUi();
  } catch (e) {
    ui = SpreadsheetApp.getUi();
  }
  
  ui.showModalDialog(html, 'Author & Affiliation Manager');
}

/**
 * Saves author and affiliation data.
 * @param {Array} authors - Array of author objects
 * @param {Array} affiliations - Array of affiliation objects
 * @return {Object} Result object with success status and message
 */
function saveAuthorData(authors, affiliations) {
  const properties = PropertiesService.getDocumentProperties();
  try {
    properties.setProperties({
      'authors': JSON.stringify(authors),
      'affiliations': JSON.stringify(affiliations)
    });
    
    // Update global variables
    authorData = authors;
    affiliationData = affiliations;
    
    return { success: true, message: 'Data saved successfully!' };
  } catch (e) {
    return { success: false, message: 'Error saving data: ' + e.toString() };
  }
}

/**
 * Loads saved author and affiliation data.
 * @return {Object} Object containing authors and affiliations arrays
 */
function loadAuthorData() {
  const properties = PropertiesService.getDocumentProperties();
  try {
    const authorsJson = properties.getProperty('authors');
    const affiliationsJson = properties.getProperty('affiliations');
    
    const authors = authorsJson ? JSON.parse(authorsJson) : [];
    const affiliations = affiliationsJson ? JSON.parse(affiliationsJson) : [];
    
    return { authors, affiliations };
  } catch (e) {
    Logger.log('Error loading data: ' + e.toString());
    return { authors: [], affiliations: [] };
  }
}

/**
 * Gets ordered affiliations based on author order.
 * @param {Array} authors - Array of author objects
 * @param {Array} affiliations - Array of affiliation objects
 * @return {Array} Ordered array of affiliations
 */
function getOrderedAffiliations(authors, affiliations) {
  const usedAffiliationIds = new Set();
  const ordered = [];
  
  authors.forEach(author => {
    if (author.affiliationIds && author.affiliationIds.length > 0) {
      // Sort affiliations by text before processing to ensure consistent order
      const sortedAffiliationIds = author.affiliationIds
        .map(id => affiliations.find(aff => aff.id === id))
        .filter(Boolean)
        .sort((a, b) => a.text.localeCompare(b.text))
        .map(aff => aff.id);

      sortedAffiliationIds.forEach(id => {
        if (!usedAffiliationIds.has(id)) {
          const aff = affiliations.find(a => a.id === id);
          if (aff) {
            usedAffiliationIds.add(id);
            ordered.push(aff);
          }
        }
      });
    }
  });
  
  return ordered;
}

/**
 * Formats the author list according to specified style.
 * @param {Array} authors - Array of author objects
 * @param {Array} affiliations - Array of affiliation objects
 * @param {Object} style - Formatting style options
 * @return {Object} Formatted text and elements for superscript
 */
function formatAuthorList(authors, affiliations, style) {
  if (!authors || authors.length === 0) {
    return { text: '', elements: [] };
  }

  // Ensure style has all required properties
  style = style || {};
  style.separator = style.separator || ',';
  style.useAnd = style.useAnd !== false;
  style.affiliationStyle = style.affiliationStyle || 'letters';
  style.markerStyle = style.markerStyle || 'superscript';
  style.markerPosition = style.markerPosition || 'after';

  const orderedAffiliations = getOrderedAffiliations(authors, affiliations);
  const affiliationMap = {};
  let result = '';
  const textElements = [];
  
  // Create affiliation mapping
  orderedAffiliations.forEach((aff, index) => {
    affiliationMap[aff.id] = {
      text: aff.text,
      marker: style.affiliationStyle === 'numbers' ? 
             (index + 1).toString() : 
             String.fromCharCode(97 + index)
    };
  });

  // Format each author
  for (let i = 0; i < authors.length; i++) {
    const author = authors[i];
    let markerText = '';
    let currentOffset;
    
    // Get affiliation markers
    if (author.affiliationIds && author.affiliationIds.length > 0) {
      const markers = author.affiliationIds
        .map(id => affiliationMap[id] ? affiliationMap[id].marker : '')
        .filter(marker => marker !== '')
        .sort();
      
      if (markers.length > 0) {
        markerText = markers.join(',');
      }
    }

    // Add separator and handle marker position
    if (i > 0) {
      if (i === authors.length - 1 && style.useAnd) {
        if (style.markerPosition === 'after-separator' && markerText) {
          result += style.separator; // Remove extra space before marker
          currentOffset = result.length;
          result += markerText + ' and '; // Add space after marker
          if (style.markerStyle === 'superscript') {
            textElements.push({
              start: currentOffset,
              end: currentOffset + markerText.length
            });
          }
        } else {
          result += style.separator === ',' ? ', and ' : ' and ';
        }
      } else {
        if (style.markerPosition === 'after-separator' && markerText) {
          result += style.separator; // Remove extra space before marker
          currentOffset = result.length;
          result += markerText + ' '; // Add space after marker
          if (style.markerStyle === 'superscript') {
            textElements.push({
              start: currentOffset,
              end: currentOffset + markerText.length
            });
          }
        } else {
          result += style.separator + ' ';
        }
      }
    }

    // Add author name and marker
    currentOffset = result.length;
    if (style.markerPosition === 'before' && markerText) {
      if (style.markerStyle === 'superscript') {
        result += markerText;
        textElements.push({
          start: currentOffset,
          end: currentOffset + markerText.length
        });
        result += author.name;
      } else {
        result += '(' + markerText + ')' + author.name;
      }
    } else if (style.markerPosition === 'after' && markerText) {
      result += author.name;
      currentOffset = result.length;
      if (style.markerStyle === 'superscript') {
        result += markerText;
        textElements.push({
          start: currentOffset,
          end: currentOffset + markerText.length
        });
      } else {
        result += '(' + markerText + ')';
      }
    } else {
      result += author.name;
    }
  }

  return { text: result, elements: textElements };
}

/**
 * Formats the affiliation list based on author order.
 * @param {Array} authors - Array of author objects
 * @param {Array} affiliations - Array of affiliation objects
 * @param {Object} style - Formatting style options
 * @return {Object} Formatted text and elements for superscript
 */
function formatAffiliationList(authors, affiliations, style) {
  const orderedAffiliations = getOrderedAffiliations(authors, affiliations);
  if (orderedAffiliations.length === 0) {
    return { text: '', elements: [] };
  }

  let formatted = '';
  const textElements = [];

  orderedAffiliations.forEach((aff, i) => {
    const marker = style.affiliationStyle === 'numbers' ? 
                  (i + 1).toString() : 
                  String.fromCharCode(97 + i);
    
    if (i > 0) {
      formatted += '\n';
    }
    
    const currentOffset = formatted.length;
    formatted += marker + ' ' + aff.text;
    
    if (style.markerStyle === 'superscript') {
      textElements.push({
        start: currentOffset,
        end: currentOffset + marker.length
      });
    }
  });

  return { text: formatted, elements: textElements };
}

/**
 * Helper function to get default style if none provided.
 * @return {Object} Default style settings
 */
function getCurrentStyle() {
  return {
    separator: ',',
    useAnd: true,
    affiliationStyle: 'letters',
    markerStyle: 'superscript',
    markerPosition: 'after'
  };
}

/**
 * Inserts formatted text with superscripts into the document.
 * @param {Array} authors - Array of author objects
 * @param {Array} affiliations - Array of affiliation objects
 * @param {Object} style - Formatting style options
 * @param {boolean} includeAffiliations - Whether to include affiliation list
 * @return {Object} Result object with success status
 */
function insertFormattedText(authors, affiliations, style, includeAffiliations) {
  try {
    const doc = DocumentApp.getActiveDocument();
    const cursor = doc.getCursor();
    const body = doc.getBody();
    
    // Format the text first
    const authorResult = formatAuthorList(authors, affiliations, style);
    let fullText = authorResult.text;
    const allElements = [...authorResult.elements];
    
    if (includeAffiliations) {
      const affiliationResult = formatAffiliationList(authors, affiliations, style);
      if (affiliationResult.text) {
        const affiliationOffset = fullText.length + 2; // Account for \n\n
        fullText += '\n\n' + affiliationResult.text;
        // Adjust offset for affiliation markers
        affiliationResult.elements.forEach(el => {
          allElements.push({
            start: el.start + affiliationOffset,
            end: el.end + affiliationOffset
          });
        });
      }
    }

    // Insert and format the text
    let insertedElement;
    if (cursor) {
      const element = cursor.getElement();
      const offset = cursor.getOffset();
      
      if (element.editAsText) {
        const text = element.editAsText();
        text.insertText(offset, fullText);
        insertedElement = text;
        
        // Apply superscript formatting to each marker
        if (style.markerStyle === 'superscript') {
          allElements.forEach(elem => {
            insertedElement.setTextStyle(
              offset + elem.start,
              offset + elem.end - 1,
              DocumentApp.newTextStyle()
                .setBaselineOffset(DocumentApp.TextBaselineOffset.SUPERSCRIPT)
                .build()
            );
          });
        }
      }
    } else {
      // Append to end of document
      const para = body.appendParagraph('');
      insertedElement = para.editAsText();
      insertedElement.setText(fullText);
      
      // Apply superscript formatting to each marker
      if (style.markerStyle === 'superscript') {
        allElements.forEach(elem => {
          insertedElement.setTextStyle(
            elem.start,
            elem.end - 1,
            DocumentApp.newTextStyle()
              .setBaselineOffset(DocumentApp.TextBaselineOffset.SUPERSCRIPT)
              .build()
          );
        });
      }
    }
    
    return { success: true };
  } catch (e) {
    Logger.log('Error inserting formatted text: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

/**
 * Inserts formatted author list into the document.
 * @param {Object} style - Optional formatting style options
 * @return {Object} Result object with success status and message
 */
function insertAuthorList(style) {
  const data = loadAuthorData();
  if (!data.authors || data.authors.length === 0) {
    return { success: false, message: 'No authors found. Please add authors first.' };
  }
  
  return insertFormattedText(
    data.authors,
    data.affiliations,
    style || getCurrentStyle(),
    false
  );
}

/**
 * Inserts formatted affiliation list into the document.
 * @param {Object} style - Optional formatting style options
 * @return {Object} Result object with success status and message
 */
function insertAffiliationList(style) {
  const data = loadAuthorData();
  if (!data.affiliations || data.affiliations.length === 0) {
    return { success: false, message: 'No affiliations found. Please add affiliations first.' };
  }
  
  const formatted = formatAffiliationList(data.authors, data.affiliations, style || getCurrentStyle());
  try {
    insertTextAtCursor(formatted.text);
    return { success: true };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

/**
 * Inserts complete formatted citation into the document.
 * @param {Object} style - Optional formatting style options
 * @return {Object} Result object with success status and message
 */
function insertFullCitation(style) {
  const data = loadAuthorData();
  if (!data.authors || data.authors.length === 0) {
    return { success: false, message: 'No authors found. Please add authors first.' };
  }
  
  return insertFormattedText(
    data.authors,
    data.affiliations,
    style || getCurrentStyle(),
    true
  );
}

/**
 * Exports data to a new Google Sheet.
 * @return {Object} Result object with success status, URL, and any error message
 */
function exportToSheet() {
  const data = loadAuthorData();
  if (data.authors.length === 0 && data.affiliations.length === 0) {
    return { success: false, error: 'No data to export.' };
  }
  
  try {
    const sheet = SpreadsheetApp.create('Author & Affiliation Data - ' + new Date().toISOString().split('T')[0]);
    const authorSheet = sheet.getActiveSheet();
    authorSheet.setName('Authors');
    
    // Set up authors sheet
    authorSheet.getRange(1, 1, 1, 3)
      .setValues([['Author Name', 'Affiliation IDs', 'Order']])
      .setFontWeight('bold');
    
    if (data.authors.length > 0) {
      const authorRows = data.authors.map((author, index) => {
        const affiliationTexts = author.affiliationIds
          .map(id => {
            const aff = data.affiliations.find(a => a.id === id);
            return aff ? aff.text : 'Unknown';
          })
          .join('; ');
        return [author.name, affiliationTexts, index + 1];
      });
      authorSheet.getRange(2, 1, authorRows.length, 3).setValues(authorRows);
    }
    
    // Set up affiliations sheet
    const affSheet = sheet.insertSheet('Affiliations');
    affSheet.getRange(1, 1, 1, 3)
      .setValues([['Affiliation ID', 'Affiliation Text', 'Order in Authors']])
      .setFontWeight('bold');
    
    const orderedAffiliations = getOrderedAffiliations(data.authors, data.affiliations);
    if (orderedAffiliations.length > 0) {
      const affRows = orderedAffiliations.map((aff, index) => [aff.id, aff.text, index + 1]);
      affSheet.getRange(2, 1, affRows.length, 3).setValues(affRows);
    }
    
    // Format sheets
    authorSheet.autoResizeColumns(1, 3);
    affSheet.autoResizeColumns(1, 3);

    return { success: true, url: sheet.getUrl() };
  } catch (e) {
    Logger.log('Export error: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

/**
 * Links to an existing Google Sheet.
 * @param {string} sheetUrl - URL of the Google Sheet to link
 * @return {Object} Result object with success status and message
 */
function linkToSheet(sheetUrl) {
  try {
    const sheetId = extractSheetId(sheetUrl);
    if (!sheetId) {
      return { success: false, message: 'Invalid Google Sheets URL' };
    }
    
    const properties = PropertiesService.getDocumentProperties();
    properties.setProperty('linkedSheetId', sheetId);
    
    return updateLinkedSheet();
  } catch (e) {
    Logger.log('Link error: ' + e.toString());
    return { success: false, message: 'Error linking to sheet: ' + e.toString() };
  }
}

/**
 * Updates the linked Google Sheet with current data.
 * @return {Object} Result object with success status and message
 */
function updateLinkedSheet() {
  try {
    const properties = PropertiesService.getDocumentProperties();
    const sheetId = properties.getProperty('linkedSheetId');
    
    if (!sheetId) {
      return { 
        success: false, 
        message: 'No linked sheet found. Please link to a sheet first.' 
      };
    }
    
    const data = loadAuthorData();
    const sheet = SpreadsheetApp.openById(sheetId);
    
    // Update authors sheet
    let authorSheet = sheet.getSheetByName('Authors') || sheet.insertSheet('Authors');
    authorSheet.clear();
    authorSheet.getRange(1, 1, 1, 3)
      .setValues([['Author Name', 'Affiliations', 'Order']])
      .setFontWeight('bold');
    
    if (data.authors.length > 0) {
      const authorRows = data.authors.map((author, index) => {
        const affiliationTexts = author.affiliationIds
          .map(id => {
            const aff = data.affiliations.find(a => a.id === id);
            return aff ? aff.text : 'Unknown';
          })
          .join('; ');
        return [author.name, affiliationTexts, index + 1];
      });
      authorSheet.getRange(2, 1, authorRows.length, 3).setValues(authorRows);
    }
    
    // Update affiliations sheet
    let affSheet = sheet.getSheetByName('Affiliations') || sheet.insertSheet('Affiliations');
    affSheet.clear();
    affSheet.getRange(1, 1, 1, 2)
      .setValues([['Affiliation Text', 'Order in Authors']])
      .setFontWeight('bold');
    
    const orderedAffiliations = getOrderedAffiliations(data.authors, data.affiliations);
    if (orderedAffiliations.length > 0) {
      const affRows = orderedAffiliations.map((aff, index) => [aff.text, index + 1]);
      affSheet.getRange(2, 1, affRows.length, 2).setValues(affRows);
    }
    
    return { success: true, message: 'Linked sheet updated successfully!' };
  } catch (e) {
    Logger.log('Update linked sheet error: ' + e.toString());
    return { success: false, message: 'Error updating linked sheet: ' + e.toString() };
  }
}

/**
 * Extracts Google Sheets ID from URL.
 * @param {string} url - Google Sheets URL
 * @return {string|null} Sheet ID if found, null otherwise
 */
function extractSheetId(url) {
  const match = url.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  return match ? match[1] : null;
}

/**
 * Inserts plain text at cursor position (works for both Docs and Sheets).
 * @param {string} text - Text to insert
 */
function insertTextAtCursor(text) {
  try {
    const doc = DocumentApp.getActiveDocument();
    const cursor = doc.getCursor();
    if (cursor) {
      cursor.insertText(text);
    } else {
      doc.getBody().appendParagraph(text);
    }
  } catch (e) {
    try {
      const sheet = SpreadsheetApp.getActiveSheet();
      const range = sheet.getActiveRange();
      if (range) {
        range.setValue(text);
      }
    } catch (e2) {
      Logger.log('Could not insert text: ' + e2.toString());
    }
  }
}

/**
 * Shows a message to the user via an alert.
 * @param {string} message - Message to display
 */
function showMessage(message) {
  let ui;
  try {
    ui = DocumentApp.getUi();
  } catch (e) {
    ui = SpreadsheetApp.getUi();
  }
  ui.alert(message);
}

/**
 * Clears all saved author and affiliation data.
 */
function clearAllData() {
  let ui;
  try {
    ui = DocumentApp.getUi();
  } catch (e) {
    ui = SpreadsheetApp.getUi();
  }
  
  const response = ui.alert(
    'Clear All Data',
    'Are you sure you want to clear all saved authors and affiliations?',
    ui.ButtonSet.YES_NO
  );
  
  if (response === ui.Button.YES) {
    const properties = PropertiesService.getDocumentProperties();
    properties.deleteProperty('authors');
    properties.deleteProperty('affiliations');
    
    authorData = [];
    affiliationData = [];
    showMessage('All data has been cleared.');
  }
}