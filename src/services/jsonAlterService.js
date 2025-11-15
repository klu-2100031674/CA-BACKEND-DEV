const fs = require('fs');

function a1ToRC(a1) {
  const m = /^([A-Za-z]+)(\d+)$/.exec(a1);
  if (!m) throw new Error(`Bad A1 address: ${a1}`);
  const colStr = m[1].toUpperCase(), row = parseInt(m[2], 10) - 1;
  let col = 0;
  for (let i = 0; i < colStr.length; i++) col = col * 26 + (colStr.charCodeAt(i) - 64);
  return { r: row, c: col - 1 };
}

function findCell(sheet, r, c) {
  return (sheet.celldata || []).find(item => item.r === r && item.c === c) || null;
}

function upsertCell(sheet, r, c) {
  const existing = findCell(sheet, r, c);
  if (existing) return existing;
  if (!sheet.celldata) sheet.celldata = [];
  const newCell = { r, c, v: {} };
  sheet.celldata.push(newCell);
  return newCell;
}

function updateLuckysheetCell(book, sheetName, a1, value) {
  const sheet = (book.data || []).find(s => s.name === sheetName);
  if (!sheet) throw new Error(`Sheet '${sheetName}' not found`);
  const { r, c } = a1ToRC(a1);
  const cell = upsertCell(sheet, r, c);
  const isNum = !isNaN(value);
  
  cell.v = {
    ...(cell.v || {}),
    v: isNum ? Number(value) : value,
    m: String(value),
    ct: isNum ? { t: 'n' } : { t: 's', fa: '@' }
  };
}

function updateAllCellStructures(book, sheetName, r, c, value) {
  const sheet = (book.data || []).find(s => s.name === sheetName);
  if (!sheet) throw new Error(`Sheet '${sheetName}' not found`);
  
  const isNum = !isNaN(value);
  let updatedCount = 0;

  if (sheet.celldata) {
    const cellDataItem = sheet.celldata.find(item => item.r === r && item.c === c);
    if (cellDataItem && cellDataItem.v) {
      cellDataItem.v = {
        ...cellDataItem.v,
        v: isNum ? Number(value) : value,
        m: String(value),
        ct: isNum ? { t: 'n' } : { t: 's', fa: '@' }
      };
      updatedCount++;
      // console.log(`  Updated celldata structure at (${r},${c})`);
    }
  }

  if (sheet.data && Array.isArray(sheet.data[r]) && sheet.data[r][c]) {
    const directCell = sheet.data[r][c];
    if (typeof directCell === 'object' && directCell !== null) {
      Object.assign(directCell, {
        v: isNum ? Number(value) : value,
        m: String(value),
        ct: isNum ? { t: 'n' } : { t: 's', fa: '@' }
      });
      updatedCount++;
      // console.log(`  Updated direct data structure at (${r},${c})`);
    }
  }

  function searchAndUpdateNested(obj, path = '', visited = new WeakSet(), depth = 0) {
    if (!obj || typeof obj !== 'object' || depth > 50) return;
    
    if (visited.has(obj)) return;
    visited.add(obj);
    
    try {
      for (const [key, value] of Object.entries(obj)) {
        if (Array.isArray(value)) {
          value.forEach((item, index) => {
            if (item && typeof item === 'object') {
              if (item.r === r && item.c === c && item.v) {
                item.v = {
                  ...item.v,
                  v: isNum ? Number(value) : value,
                  m: String(value),
                  ct: isNum ? { t: 'n' } : { t: 's', fa: '@' }
                };
                updatedCount++;
                // console.log(`  Updated nested structure at ${path}.${key}[${index}]`);
              }
              searchAndUpdateNested(item, `${path}.${key}[${index}]`, visited, depth + 1);
            }
          });
        } else if (value && typeof value === 'object') {
          searchAndUpdateNested(value, `${path}.${key}`, visited, depth + 1);
        }
      }
    } catch (error) {
      console.warn(`Error in searchAndUpdateNested at ${path}:`, error.message);
    } finally {
      visited.delete(obj);
    }
  }

  searchAndUpdateNested(sheet, `sheet.${sheetName}`);

  return updatedCount;
}

function run() {
  try {
    const book = JSON.parse(fs.readFileSync('workbook.json', 'utf-8'));
    console.log('Loaded workbook.json successfully.');

    if (book.data) {
      book.data.forEach(sheet => {
        if (sheet.celldata) {
          sheet.celldata = sheet.celldata.filter(cell => cell.r !== undefined && cell.c !== undefined);
        }
      });
    }

    const updates = {
      "H3": "15000000",
      "I11": "1000000"
    };

    for (const [a1, val] of Object.entries(updates)) {
      console.log(`\nUpdating ${a1} with ${val}:`);
      const { r, c } = a1ToRC(a1);
      const updatedCount = updateAllCellStructures(book, 'Assumptions.1', r, c, val);
      console.log(`  Total structures updated: ${updatedCount}`);
      
      updateLuckysheetCell(book, 'Assumptions.1', a1, val);
    }

    fs.writeFileSync('output.json', JSON.stringify(book, null, 2));
    console.log('\nSaved to output.json');
    console.log('Note: Formulas are preserved, only values and display text updated.');
  } catch (error) {
    console.error('Error:', error.message);
  }
}

function safeDeepCopy(obj, visited = new WeakMap(), depth = 0) {
  if (obj === null || typeof obj !== 'object') return obj;
  
  if (depth > 100) {
    console.warn('Deep copy depth limit reached, returning null');
    return null;
  }
  
  if (visited.has(obj)) return visited.get(obj);
  
  if (obj instanceof Date) return new Date(obj);
  
  if (Array.isArray(obj)) {
    const arrCopy = [];
    visited.set(obj, arrCopy);
    try {
      obj.forEach((item, index) => {
        arrCopy[index] = safeDeepCopy(item, visited, depth + 1);
      });
    } catch (error) {
      console.warn('Error copying array:', error.message);
    }
    return arrCopy;
  }
  
  const objCopy = {};
  visited.set(obj, objCopy);
  try {
    Object.keys(obj).forEach(key => {
      objCopy[key] = safeDeepCopy(obj[key], visited, depth + 1);
    });
  } catch (error) {
    console.warn('Error copying object:', error.message);
  }
  
  return objCopy;
}

function alterTemplateJson(templateJson, formData, sheetName = 'Assumptions.1') {
  try {
    const book = safeDeepCopy(templateJson);
    
    if (book.data) {
      book.data.forEach(sheet => {
        if (sheet.celldata) {
          sheet.celldata = sheet.celldata.filter(cell => cell.r !== undefined && cell.c !== undefined);
        }
      });
    }

    for (const [cellAddress, value] of Object.entries(formData)) {
      if (value !== undefined && value !== '') {
        const upperCellAddress = cellAddress.toUpperCase();
        // console.log(`Updating ${upperCellAddress} with ${value}`);
        
        try {
          const { r, c } = a1ToRC(upperCellAddress);
          updateAllCellStructures(book, sheetName, r, c, value);
          updateLuckysheetCell(book, sheetName, upperCellAddress, value);
        } catch (error) {
          console.warn(`Invalid cell address ${upperCellAddress}: ${error.message}`);
        }
      }
    }

    return book;
  } catch (error) {
    console.error('Error altering template JSON:', error.message);
    throw error;
  }
}

module.exports = {
  alterTemplateJson,
  a1ToRC,
  updateLuckysheetCell,
  updateAllCellStructures,
  safeDeepCopy
};

if (require.main === module) {
  run();
}
     