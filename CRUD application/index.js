const ExcelJS = require('exceljs');

const workbook = new ExcelJS.Workbook();

const sheet = workbook.addWorksheet('Sheet1');

const saveWorkbook = async () => {
  try {
    await workbook.xlsx.writeFile('data.xlsx');
    console.log('Workbook saved successfully.');
  } catch (error) {
    console.error('Failed to save the workbook:', error);
  }
};

const createEntry = async (data) => {
  try {

    const newRow = sheet.addRow(data);
    await saveWorkbook();
    console.log('Entry created successfully.');
  } catch (error) {
    console.error('Failed to create entry:', error);
  }
};

const readData = async () => {
  try {
    const rows = sheet.getRows();
    const data = rows.map((row) => row.values);
    console.log('Data read successfully:', data);
    return data;
  } catch (error) {
    console.error('Failed to read data:', error);
    return [];
  }
};

const updateEntry = async (rowIndex, newData) => {
  try {
    const row = sheet.getRow(rowIndex);
    Object.keys(newData).forEach((key) => {
      row.getCell(key).value = newData[key];
    });
    await saveWorkbook();
    console.log('Entry updated successfully.');
  } catch (error) {
    console.error('Failed to update entry:', error);
  }
};

const deleteEntry = async (rowIndex) => {
  try {
    sheet.spliceRows(rowIndex, 1);
    await saveWorkbook();
    console.log('Entry deleted successfully.');
  } catch (error) {
    console.error('Failed to delete entry:', error);
  }
};

createEntry({ field1: 'Value 1', field2: 'Value 2' });

readData();

updateEntry(2, { field1: 'Updated Value 1' });

deleteEntry(3);
