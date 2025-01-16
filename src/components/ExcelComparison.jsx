import React, { useState } from 'react';
import * as XLSX from 'xlsx';

const ExcelComparison = () => {
  const [sheets, setSheets] = useState([]);
  const [firstSheet, setFirstSheet] = useState('');
  const [secondSheet, setSecondSheet] = useState('');
  const [firstColumn, setFirstColumn] = useState('');
  const [secondColumn, setSecondColumn] = useState('');
  const [differences, setDifferences] = useState({ inFirstOnly: [], inSecondOnly: [] });
  const [columnCounts, setColumnCounts] = useState({});
  const [debugInfo, setDebugInfo] = useState('');

  const getColumnLetter = (columnNumber) => {
    let dividend = columnNumber;
    let columnName = '';
    let modulo;

    while (dividend > 0) {
      modulo = (dividend - 1) % 26;
      columnName = String.fromCharCode(65 + modulo) + columnName;
      dividend = Math.floor((dividend - modulo) / 26);
    }

    return columnName;
  };

  const getColumnOptions = (count) => {
    return Array.from({ length: count }, (_, i) => getColumnLetter(i + 1));
  };

  const handleFileUpload = (event) => {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, {
          type: 'array',
          cellDates: true,
          cellNF: true,
          cellText: true
        });

        const sheetNames = workbook.SheetNames;
        setSheets(sheetNames);

        const counts = {};
        sheetNames.forEach(sheet => {
          const worksheet = workbook.Sheets[sheet];
          const range = XLSX.utils.decode_range(worksheet['!ref']);
          counts[sheet] = range.e.c + 1;
        });
        
        setColumnCounts(counts);
        window.currentWorkbook = workbook;
        setDebugInfo('File loaded successfully. Sheets found: ' + sheetNames.join(', '));
      } catch (error) {
        setDebugInfo('Error loading file: ' + error.message);
        console.error('Error loading workbook:', error);
      }
    };

    reader.readAsArrayBuffer(file);
  };

  const compareColumns = () => {
    try {
      setDebugInfo('Starting comparison...');

      if (!window.currentWorkbook) {
        setDebugInfo('No workbook found. Please upload a file first.');
        return;
      }

      const workbook = window.currentWorkbook;
      const col1Index = XLSX.utils.decode_col(firstColumn);
      const col2Index = XLSX.utils.decode_col(secondColumn);

      const worksheet1 = workbook.Sheets[firstSheet];
      const worksheet2 = workbook.Sheets[secondSheet];

      if (!worksheet1 || !worksheet2) {
        setDebugInfo('Could not find one or both worksheets');
        return;
      }

      const range1 = XLSX.utils.decode_range(worksheet1['!ref']);
      const range2 = XLSX.utils.decode_range(worksheet2['!ref']);

      const values1 = new Set();
      const values2 = new Set();

      // Get values from first sheet
      for (let row = range1.s.r; row <= range1.e.r; row++) {
        const cellAddress = XLSX.utils.encode_cell({ r: row, c: col1Index });
        const cell = worksheet1[cellAddress];
        if (cell && cell.v !== undefined) {
          values1.add(cell.v.toString().trim());
        }
      }

      // Get values from second sheet
      for (let row = range2.s.r; row <= range2.e.r; row++) {
        const cellAddress = XLSX.utils.encode_cell({ r: row, c: col2Index });
        const cell = worksheet2[cellAddress];
        if (cell && cell.v !== undefined) {
          values2.add(cell.v.toString().trim());
        }
      }

      setDebugInfo(`Found ${values1.size} values in first column and ${values2.size} values in second column`);

      // Find differences
      const inFirstOnly = Array.from(values1).filter(value => !values2.has(value));
      const inSecondOnly = Array.from(values2).filter(value => !values1.has(value));

      setDifferences({
        inFirstOnly,
        inSecondOnly
      });

      setDebugInfo(`Comparison complete. Found ${inFirstOnly.length} unique to first column and ${inSecondOnly.length} unique to second column`);
    } catch (error) {
      setDebugInfo('Error during comparison: ' + error.message);
      console.error('Error during comparison:', error);
    }
  };

  return (
    <div className="max-w-4xl mx-auto p-6 bg-white rounded-lg shadow-lg">
      <h1 className="text-2xl font-bold mb-6">Excel Sheet Comparison Tool</h1>
      
      <div className="space-y-6">
        <input
          type="file"
          accept=".xlsx,.xls"
          onChange={handleFileUpload}
          className="block w-full text-sm text-gray-500 file:mr-4 file:py-2 file:px-4 file:rounded-full file:border-0 file:text-sm file:font-semibold file:bg-blue-50 file:text-blue-700 hover:file:bg-blue-100"
        />

        {debugInfo && (
          <div className="bg-gray-100 p-4 rounded-md text-sm">
            <pre>{debugInfo}</pre>
          </div>
        )}

        {sheets.length > 0 && (
          <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
            <div className="space-y-2">
              <label className="block text-sm font-medium text-gray-700">First Sheet</label>
              <select 
                value={firstSheet}
                onChange={(e) => setFirstSheet(e.target.value)}
                className="mt-1 block w-full rounded-md border border-gray-300 p-2"
              >
                <option value="">Select Sheet</option>
                {sheets.map(sheet => (
                  <option key={sheet} value={sheet}>{sheet}</option>
                ))}
              </select>

              {firstSheet && (
                <>
                  <label className="block text-sm font-medium text-gray-700">Column</label>
                  <select
                    value={firstColumn}
                    onChange={(e) => setFirstColumn(e.target.value)}
                    className="mt-1 block w-full rounded-md border border-gray-300 p-2"
                  >
                    <option value="">Select Column</option>
                    {getColumnOptions(columnCounts[firstSheet]).map(col => (
                      <option key={col} value={col}>Column {col}</option>
                    ))}
                  </select>
                </>
              )}
            </div>

            <div className="space-y-2">
              <label className="block text-sm font-medium text-gray-700">Second Sheet</label>
              <select
                value={secondSheet}
                onChange={(e) => setSecondSheet(e.target.value)}
                className="mt-1 block w-full rounded-md border border-gray-300 p-2"
              >
                <option value="">Select Sheet</option>
                {sheets.map(sheet => (
                  <option key={sheet} value={sheet}>{sheet}</option>
                ))}
              </select>

              {secondSheet && (
                <>
                  <label className="block text-sm font-medium text-gray-700">Column</label>
                  <select
                    value={secondColumn}
                    onChange={(e) => setSecondColumn(e.target.value)}
                    className="mt-1 block w-full rounded-md border border-gray-300 p-2"
                  >
                    <option value="">Select Column</option>
                    {getColumnOptions(columnCounts[secondSheet]).map(col => (
                      <option key={col} value={col}>Column {col}</option>
                    ))}
                  </select>
                </>
              )}
            </div>
          </div>
        )}

        {firstSheet && secondSheet && firstColumn && secondColumn && (
          <button
            onClick={compareColumns}
            className="w-full px-4 py-2 bg-blue-600 text-white rounded-md hover:bg-blue-700 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:ring-offset-2"
          >
            Compare Columns
          </button>
        )}

        {(differences.inFirstOnly.length > 0 || differences.inSecondOnly.length > 0) && (
          <div className="space-y-4">
            {differences.inFirstOnly.length > 0 && (
              <div className="bg-blue-50 border border-blue-200 rounded-md p-4">
                <h3 className="font-medium mb-2">Values only in {firstSheet} - Column {firstColumn}:</h3>
                <div className="max-h-40 overflow-y-auto">
                  {differences.inFirstOnly.map((value, index) => (
                    <div key={index} className="text-sm text-gray-600">{value}</div>
                  ))}
                </div>
              </div>
            )}

            {differences.inSecondOnly.length > 0 && (
              <div className="bg-blue-50 border border-blue-200 rounded-md p-4">
                <h3 className="font-medium mb-2">Values only in {secondSheet} - Column {secondColumn}:</h3>
                <div className="max-h-40 overflow-y-auto">
                  {differences.inSecondOnly.map((value, index) => (
                    <div key={index} className="text-sm text-gray-600">{value}</div>
                  ))}
                </div>
              </div>
            )}
          </div>
        )}
      </div>
    </div>
  );
};

export default ExcelComparison;
