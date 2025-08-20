/**
 * Author and Affiliation Manager for Google Docs and Sheets
 * This script creates a user interface to manage author lists and their affiliations
 * with customizable formatting for different journal requirements
 */

// Global variables for author data
var authorData = [];
var affiliationData = [];

/**
 * Creates custom menu items when the document is opened
 */
function onOpen() {
  var ui;
  
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
 * Shows the main author management interface
 */
function showAuthorManager() {
  var html = HtmlService.createHtmlOutputFromFile('AuthorManager')
    .setWidth(800)
    .setHeight(600)
    .setTitle('Author & Affiliation Manager');
  
  var ui;
  try {
    ui = DocumentApp.getUi();
  } catch (e) {
    ui = SpreadsheetApp.getUi();
  }
  
  ui.showModalDialog(html, 'Author & Affiliation Manager');
}

/**
 * Saves author and affiliation data
 */
function saveAuthorData(authors, affiliations) {
  var properties = PropertiesService.getDocumentProperties();
  
  try {
    properties.setProperties({
      'authors': JSON.stringify(authors),
      'affiliations': JSON.stringify(affiliations)
    });
    
    // Update global variables
    authorData = authors;
    affiliationData = affiliations;
    
    return {success: true, message: 'Data saved successfully!'};
  } catch (e) {
    return {success: false, message: 'Error saving data: ' + e.toString()};
  }
}

/**
 * Loads saved author and affiliation data
 */
function loadAuthorData() {
  var properties = PropertiesService.getDocumentProperties();
  
  try {
    var authorsJson = properties.getProperty('authors');
    var affiliationsJson = properties.getProperty('affiliations');
    
    var authors = authorsJson ? JSON.parse(authorsJson) : [];
    var affiliations = affiliationsJson ? JSON.parse(affiliationsJson) : [];
    
    return {
      authors: authors,
      affiliations: affiliations
    };
  } catch (e) {
    Logger.log('Error loading data: ' + e.toString());
    return {
      authors: [],
      affiliations: []
    };
  }
}

/**
 * Gets ordered affiliations based on author order
 */
function getOrderedAffiliations(authors, affiliations) {
  var usedAffiliationIds = [];
  var affiliationOrder = {};
  
  // Go through authors in order and collect their affiliations
  authors.forEach(function(author) {
    if (author.affiliationIds && author.affiliationIds.length > 0) {
      // Sort author's affiliations alphabetically by text for consistency
      var authorAffs = author.affiliationIds.map(function(id) {
        return affiliations.find(function(aff) { return aff.id === id; });
      }).filter(function(aff) { return aff; }).sort(function(a, b) {
        return a.text.localeCompare(b.text);
      });
      
      authorAffs.forEach(function(aff) {
        if (usedAffiliationIds.indexOf(aff.id) === -1) {
          usedAffiliationIds.push(aff.id);
          affiliationOrder[aff.id] = usedAffiliationIds.length - 1;
        }
      });
    }
  });
  
  // Return ordered affiliations
  return usedAffiliationIds.map(function(id) {
    return affiliations.find(function(aff) { return aff.id === id; });
  }).filter(function(aff) { return aff; });
}

/**
 * Formats the author list according to specified style
 */
function formatAuthorList(authors, affiliations, style) {
  if (!authors || authors.length === 0) {
    return '';
  }
  
  var orderedAffiliations = getOrderedAffiliations(authors, affiliations);
  var affiliationMap = {};
  
  // Create affiliation mapping based on author order
  orderedAffiliations.forEach(function(aff, index) {
    affiliationMap[aff.id] = {
      text: aff.text,
      marker: style.affiliationStyle === 'numbers' ? (index + 1).toString() : 
              String.fromCharCode(97 + index) // a, b, c, etc.
    };
  });
  
  var result = '';
  
  // Format authors
  for (var i = 0; i < authors.length; i++) {
    var author = authors[i];
    var authorText = author.name;
    var markerText = '';
    
    // Get affiliation markers
    if (author.affiliationIds && author.affiliationIds.length > 0) {
      var markers = author.affiliationIds.map(function(id) {
        return affiliationMap[id] ? affiliationMap[id].marker : '';
      }).filter(function(marker) {
        return marker !== '';
      }).sort(); // Sort markers for consistency
      
      if (markers.length > 0) {
        var joinedMarkers = markers.join(',');
        if (style.markerStyle === 'superscript') {
          markerText = joinedMarkers;
        } else {
          markerText = '(' + joinedMarkers + ')';
        }
      }
    }
    
    // Add separator before author (except first)
    if (i > 0) {
      if (i === authors.length - 1 && style.useAnd) {
        result += style.separator === ',' ? ', and ' : ' and ';
      } else {
        result += style.separator + ' ';
      }
    }
    
    // Add author and marker based on position preference
    if (style.markerPosition === 'before' && markerText) {
      result += markerText + authorText;
    } else {
      result += authorText + markerText;
    }
  }
  
  return result;
}

/**
 * Formats the affiliation list based on author order
 */
function formatAffiliationList(authors, affiliations, style) {
  var orderedAffiliations = getOrderedAffiliations(authors, affiliations);
  
  if (orderedAffiliations.length === 0) {
    return '';
  }
  
  var formatted = '';
  
  for (var i = 0; i < orderedAffiliations.length; i++) {
    var aff = orderedAffiliations[i];
    var marker = style.affiliationStyle === 'numbers' ? (i + 1).toString() : 
                String.fromCharCode(97 + i);
    
    if (i > 0) {
      formatted += '\n';
    }
    
    formatted += marker + ' ' + aff.text;
  }
  
  return formatted;
}

/**
 * Inserts text with proper formatting (including superscripts for Docs)
 */
function insertFormattedText(authors, affiliations, style, includeAffiliations) {
  try {
    var doc = DocumentApp.getActiveDocument();
    var cursor = doc.getCursor();
    var body = doc.getBody();
    var insertionPoint;
    
    if (cursor) {
      insertionPoint = cursor.getElement();
      var offset = cursor.getOffset();
    } else {
      var para = body.appendParagraph('');
      insertionPoint = para;
      var offset = 0;
    }
    
    // Format authors with proper superscripts
    if (authors && authors.length > 0) {
      var orderedAffiliations = getOrderedAffiliations(authors, affiliations);
      var affiliationMap = {};
      
      // Create affiliation mapping based on author order
      orderedAffiliations.forEach(function(aff, index) {
        affiliationMap[aff.id] = {
          text: aff.text,
          marker: style.affiliationStyle === 'numbers' ? (index + 1).toString() : 
                  String.fromCharCode(97 + index)
        };
      });
      
      var fullText = '';
      var textElements = [];
      
      // Build the full text and track where markers should be
      for (var i = 0; i < authors.length; i++) {
        var author = authors[i];
        
        // Add separator
        if (i > 0) {
          if (i === authors.length - 1 && style.useAnd) {
            fullText += style.separator === ',' ? ', and ' : ' and ';
          } else {
            fullText += style.separator + ' ';
          }
        }
        
        // Get affiliation markers
        var markerText = '';
        if (author.affiliationIds && author.affiliationIds.length > 0) {
          var markers = author.affiliationIds.map(function(id) {
            return affiliationMap[id] ? affiliationMap[id].marker : '';
          }).filter(function(marker) {
            return marker !== '';
          }).sort();
          
          if (markers.length > 0) {
            markerText = markers.join(',');
          }
        }
        
        // Add author name and marker based on position
        if (style.markerPosition === 'before' && markerText) {
          if (style.markerStyle === 'parentheses') {
            fullText += '(' + markerText + ')' + author.name;
          } else {
            var markerStart = fullText.length;
            fullText += markerText + author.name;
            textElements.push({
              start: markerStart,
              end: markerStart + markerText.length,
              superscript: true
            });
          }
        } else {
          var nameStart = fullText.length;
          fullText += author.name;
          
          if (markerText) {
            var markerStart = fullText.length;
            if (style.markerStyle === 'parentheses') {
              fullText += '(' + markerText + ')';
            } else {
              fullText += markerText;
              textElements.push({
                start: markerStart,
                end: markerStart + markerText.length,
                superscript: true
              });
            }
          }
        }
      }
      
      // Insert the text
      if (cursor) {
        var textElement = insertionPoint.asText();
        textElement.insertText(offset, fullText);
        
        // Apply superscript formatting to markers
        textElements.forEach(function(element) {
          if (element.superscript) {
            textElement.setTextStyle(offset + element.start, offset + element.end - 1, 
              DocumentApp.newTextStyle().setBaselineOffset(DocumentApp.TextBaselineOffset.SUPERSCRIPT).build());
          }
        });
      } else {
        var para = insertionPoint.asParagraph();
        para.setText(fullText);
        
        // Apply superscript formatting
        textElements.forEach(function(element) {
          if (element.superscript) {
            para.setTextStyle(element.start, element.end - 1, 
              DocumentApp.newTextStyle().setBaselineOffset(DocumentApp.TextBaselineOffset.SUPERSCRIPT).build());
          }
        });
      }
      
      // Add affiliations if requested
      if (includeAffiliations && orderedAffiliations.length > 0) {
        var affText = '\n\n';
        var affElements = [];
        
        for (var i = 0; i < orderedAffiliations.length; i++) {
          var aff = orderedAffiliations[i];
          var marker = style.affiliationStyle === 'numbers' ? (i + 1).toString() : 
                      String.fromCharCode(97 + i);
          
          if (i > 0) {
            affText += '\n';
          }
          
          var markerStart = affText.length;
          if (style.markerStyle === 'parentheses') {
            affText += '(' + marker + ') ' + aff.text;
          } else {
            affText += marker + ' ' + aff.text;
            affElements.push({
              start: markerStart,
              end: markerStart + marker.length,
              superscript: true
            });
          }
        }
        
        // Insert affiliation text
        if (cursor) {
          var textElement = insertionPoint.asText();
          var affOffset = textElement.getText().length;
          textElement.insertText(affOffset, affText);
          
          affElements.forEach(function(element) {
            if (element.superscript) {
              textElement.setTextStyle(affOffset + element.start, affOffset + element.end - 1, 
                DocumentApp.newTextStyle().setBaselineOffset(DocumentApp.TextBaselineOffset.SUPERSCRIPT).build());
            }
          });
        } else {
          var affPara = body.appendParagraph(affText);
          affElements.forEach(function(element) {
            if (element.superscript) {
              affPara.setTextStyle(element.start, element.end - 1, 
                DocumentApp.newTextStyle().setBaselineOffset(DocumentApp.TextBaselineOffset.SUPERSCRIPT).build());
            }
          });
        }
      }
    }
  } catch (e) {
    Logger.log('Formatting error: ' + e.toString());
    // Fallback to simple text insertion for Sheets
    var authorList = formatAuthorList(authors, affiliations, style);
    var affiliationList = formatAffiliationList(authors, affiliations, style);
    
    var text = authorList;
    if (includeAffiliations && affiliationList) {
      text += '\n\n' + affiliationList;
    }
    
    insertTextAtCursor(text);
  }
}

/**
 * Inserts formatted author list into document
 */
function insertAuthorList() {
  var data = loadAuthorData();
  
  if (data.authors.length === 0) {
    showMessage('No authors found. Please use Author Manager to add authors first.');
    return;
  }
  
  var style = {
    separator: ',',
    useAnd: true,
    affiliationStyle: 'letters', // 'letters' or 'numbers'
    markerStyle: 'superscript', // 'superscript' or 'parentheses'
    markerPosition: 'after' // 'before' or 'after'
  };
  
  insertFormattedText(data.authors, data.affiliations, style, false);
}

/**
 * Inserts formatted affiliation list into document
 */
function insertAffiliationList() {
  var data = loadAuthorData();
  
  if (data.affiliations.length === 0) {
    showMessage('No affiliations found. Please use Author Manager to add affiliations first.');
    return;
  }
  
  var style = {
    affiliationStyle: 'letters',
    markerStyle: 'superscript',
    markerPosition: 'after'
  };
  
  var affiliationList = formatAffiliationList(data.authors, data.affiliations, style);
  insertTextAtCursor(affiliationList);
}

/**
 * Inserts complete citation block (authors + affiliations)
 */
function insertFullCitation() {
  var data = loadAuthorData();
  
  if (data.authors.length === 0) {
    showMessage('No authors found. Please use Author Manager to add authors first.');
    return;
  }
  
  var style = {
    separator: ',',
    useAnd: true,
    affiliationStyle: 'letters',
    markerStyle: 'superscript',
    markerPosition: 'after'
  };
  
  insertFormattedText(data.authors, data.affiliations, style, true);
}

/**
 * Exports data to a new Google Sheet
 */
function exportToSheet() {
  var data = loadAuthorData();
  
  if (data.authors.length === 0 && data.affiliations.length === 0) {
    showMessage('No data to export.');
    return;
  }
  
  try {
    // Create new spreadsheet
    var sheet = SpreadsheetApp.create('Author & Affiliation Data - ' + new Date().toISOString().split('T')[0]);
    var authorSheet = sheet.getActiveSheet();
    authorSheet.setName('Authors');
    
    // Set up author sheet headers
    authorSheet.getRange(1, 1, 1, 3).setValues([['Author Name', 'Affiliation IDs', 'Order']]);
    authorSheet.getRange(1, 1, 1, 3).setFontWeight('bold');
    
    // Add author data with actual affiliation text
    if (data.authors.length > 0) {
      var authorData = data.authors.map(function(author, index) {
        var affiliationTexts = '';
        if (author.affiliationIds && author.affiliationIds.length > 0) {
          affiliationTexts = author.affiliationIds.map(function(id) {
            var aff = data.affiliations.find(function(a) { return a.id === id; });
            return aff ? aff.text : 'Unknown';
          }).join('; ');
        }
        
        return [
          author.name,
          affiliationTexts,
          index + 1
        ];
      });
      authorSheet.getRange(2, 1, authorData.length, 3).setValues(authorData);
    }
    
    // Create affiliations sheet
    var affSheet = sheet.insertSheet('Affiliations');
    affSheet.getRange(1, 1, 1, 3).setValues([['Affiliation ID', 'Affiliation Text', 'Order in Authors']]);
    affSheet.getRange(1, 1, 1, 3).setFontWeight('bold');
    
    // Add affiliation data in author order
    var orderedAffiliations = getOrderedAffiliations(data.authors, data.affiliations);
    if (orderedAffiliations.length > 0) {
      var affData = orderedAffiliations.map(function(aff, index) {
        return [
          aff.id,
          aff.text,
          index + 1
        ];
      });
      affSheet.getRange(2, 1, affData.length, 3).setValues(affData);
    }
    
    // Create formatted output sheet
    var outputSheet = sheet.insertSheet('Formatted Output');
    
    // Generate different format examples
    var styles = [
      {name: 'Nature Style (letters, superscript)', separator: ',', useAnd: true, affiliationStyle: 'letters', markerStyle: 'superscript', markerPosition: 'after'},
      {name: 'Cell Style (numbers, parentheses)', separator: ',', useAnd: false, affiliationStyle: 'numbers', markerStyle: 'parentheses', markerPosition: 'after'},
      {name: 'Semicolon Style (letters, parentheses)', separator: ';', useAnd: false, affiliationStyle: 'letters', markerStyle: 'parentheses', markerPosition: 'after'}
    ];
    
    var outputData = [['Format Style', 'Author List', 'Affiliation List']];
    
    styles.forEach(function(style) {
      var authorList = formatAuthorList(data.authors, data.affiliations, style);
      var affiliationList = formatAffiliationList(data.authors, data.affiliations, style);
      outputData.push([style.name, authorList, affiliationList]);
    });
    
    outputSheet.getRange(1, 1, outputData.length, 3).setValues(outputData);
    outputSheet.getRange(1, 1, 1, 3).setFontWeight('bold');
    
    // Auto-resize columns
    authorSheet.autoResizeColumns(1, 3);
    affSheet.autoResizeColumns(1, 3);
    outputSheet.autoResizeColumns(1, 3);
    
    var url = sheet.getUrl();
    showMessage('Data exported successfully! Sheet URL: ' + url);
    return {success: true, url: url};
    
  } catch (e) {
    Logger.log('Export error: ' + e.toString());
    showMessage('Error exporting data: ' + e.toString());
    return {success: false, error: e.toString()};
  }
}

/**
 * Links to an existing Google Sheet (saves the sheet ID for future updates)
 */
function linkToSheet(sheetUrl) {
  try {
    var sheetId = extractSheetId(sheetUrl);
    if (!sheetId) {
      return {success: false, message: 'Invalid Google Sheets URL'};
    }
    
    // Save the linked sheet ID
    var properties = PropertiesService.getDocumentProperties();
    properties.setProperty('linkedSheetId', sheetId);
    
    // Update the linked sheet with current data
    return updateLinkedSheet();
    
  } catch (e) {
    Logger.log('Link error: ' + e.toString());
    return {success: false, message: 'Error linking to sheet: ' + e.toString()};
  }
}

/**
 * Updates the linked Google Sheet with current data
 */
function updateLinkedSheet() {
  try {
    var properties = PropertiesService.getDocumentProperties();
    var sheetId = properties.getProperty('linkedSheetId');
    
    if (!sheetId) {
      return {success: false, message: 'No linked sheet found. Please link to a sheet first.'};
    }
    
    var data = loadAuthorData();
    var sheet = SpreadsheetApp.openById(sheetId);
    
    // Update or create Authors sheet
    var authorSheet;
    try {
      authorSheet = sheet.getSheetByName('Authors');
    } catch (e) {
      authorSheet = sheet.insertSheet('Authors');
    }
    
    // Clear existing data and add headers
    authorSheet.clear();
    authorSheet.getRange(1, 1, 1, 3).setValues([['Author Name', 'Affiliations', 'Order']]);
    authorSheet.getRange(1, 1, 1, 3).setFontWeight('bold');
    
    // Add author data with actual affiliation text
    if (data.authors.length > 0) {
      var authorData = data.authors.map(function(author, index) {
        var affiliationTexts = '';
        if (author.affiliationIds && author.affiliationIds.length > 0) {
          affiliationTexts = author.affiliationIds.map(function(id) {
            var aff = data.affiliations.find(function(a) { return a.id === id; });
            return aff ? aff.text : 'Unknown';
          }).join('; ');
        }
        
        return [
          author.name,
          affiliationTexts,
          index + 1
        ];
      });
      authorSheet.getRange(2, 1, authorData.length, 3).setValues(authorData);
    }
    
    // Update Affiliations sheet
    var affSheet;
    try {
      affSheet = sheet.getSheetByName('Affiliations');
    } catch (e) {
      affSheet = sheet.insertSheet('Affiliations');
    }
    
    affSheet.clear();
    affSheet.getRange(1, 1, 1, 2).setValues([['Affiliation Text', 'Order in Authors']]);
    affSheet.getRange(1, 1, 1, 2).setFontWeight('bold');
    
    var orderedAffiliations = getOrderedAffiliations(data.authors, data.affiliations);
    if (orderedAffiliations.length > 0) {
      var affData = orderedAffiliations.map(function(aff, index) {
        return [
          aff.text,
          index + 1
        ];
      });
      affSheet.getRange(2, 1, affData.length, 2).setValues(affData);
    }
    
    return {success: true, message: 'Linked sheet updated successfully!'};
    
  } catch (e) {
    Logger.log('Update linked sheet error: ' + e.toString());
    return {success: false, message: 'Error updating linked sheet: ' + e.toString()};
  }
}

/**
 * Extracts Google Sheets ID from URL
 */
function extractSheetId(url) {
  var match = url.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  return match ? match[1] : null;
}

/**
 * Inserts text at cursor position (works for both Docs and Sheets)
 */
function insertTextAtCursor(text) {
  try {
    // Try Google Docs first
    var doc = DocumentApp.getActiveDocument();
    var cursor = doc.getCursor();
    
    if (cursor) {
      cursor.insertText(text);
    } else {
      var body = doc.getBody();
      body.appendParagraph(text);
    }
  } catch (e) {
    try {
      // If Docs fails, try Sheets
      var sheet = SpreadsheetApp.getActiveSheet();
      var range = sheet.getActiveRange();
      
      if (range) {
        range.setValue(text);
      }
    } catch (e2) {
      Logger.log('Could not insert text: ' + e2.toString());
      showMessage('Could not insert text. Please try selecting a location first.');
    }
  }
}

/**
 * Shows a message to the user
 */
function showMessage(message) {
  var ui;
  try {
    ui = DocumentApp.getUi();
  } catch (e) {
    ui = SpreadsheetApp.getUi();
  }
  
  ui.alert('Author Manager', message, ui.ButtonSet.OK);
}

/**
 * Clears all saved author and affiliation data
 */
function clearAllData() {
  var ui;
  try {
    ui = DocumentApp.getUi();
  } catch (e) {
    ui = SpreadsheetApp.getUi();
  }
  
  var response = ui.alert('Clear All Data', 
    'Are you sure you want to clear all saved authors and affiliations?', 
    ui.ButtonSet.YES_NO);
  
  if (response === ui.Button.YES) {
    var properties = PropertiesService.getDocumentProperties();
    properties.deleteProperty('authors');
    properties.deleteProperty('affiliations');
    
    authorData = [];
    affiliationData = [];
    
    showMessage('All data has been cleared.');
  }
}

/**
 * Utility function to generate unique IDs
 */
function generateId() {
  return 'id_' + Math.random().toString(36).substr(2, 9);
}