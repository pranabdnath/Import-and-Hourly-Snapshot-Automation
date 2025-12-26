function importAndSnapshot() {
  const sourceId = "your_source_sheet_id_here";
  const sourceSheetName = "Active Projects";
  const sourceRange = "A:L";
  const targetSheetName = "ActiveWallet";

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheet = ss.getSheetByName(targetSheetName);

  const sourceSS = SpreadsheetApp.openById(sourceId);
  const sourceSheet = sourceSS.getSheetByName(sourceSheetName);

  const data = sourceSheet.getRange(sourceRange).getValues();

  targetSheet.getRange(1, 1, targetSheet.getMaxRows(), 12).clearContent();
  targetSheet.getRange(1, 1, data.length, data[0].length).setValues(data);

  hourlySnapshotFromImported(targetSheet);
}

function hourlySnapshotFromImported(sheet) {
  const now = new Date();
  const hour = now.getHours();
  const minute = now.getMinutes();

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const data = sheet.getRange(2, 4, lastRow - 1, 4).getValues();

  const cities = [];
  const values = [];

  data.forEach(row => {
    const city = row[0];
    const value = Number(row[3]) || 0;
    if (city) {
      cities.push(city);
      values.push(value);
    }
  });

  const uniqueCities = [...new Set(cities)];

  const cityTotal = uniqueCities.map(
    city => cities.filter(c => c === city).length
  );

  const cityAbove15 = uniqueCities.map(city =>
    cities.reduce(
      (count, c, i) => (c === city && values[i] >= 15 ? count + 1 : count),
      0
    )
  );

  let colStart = 18;
  if (hour === 9 && minute <= 10) colStart = 14;

  const rowsToClear = uniqueCities.length + 5;
  sheet.getRange(4, colStart, rowsToClear, 3).clearContent();

  sheet.getRange(4, colStart, uniqueCities.length, 1)
    .setValues(uniqueCities.map(c => [c]));

  sheet.getRange(4, colStart + 1, cityTotal.length, 1)
    .setValues(cityTotal.map(v => [v]));

  sheet.getRange(4, colStart + 2, cityAbove15.length, 1)
    .setValues(cityAbove15.map(v => [v]));

  const totalRow = 4 + uniqueCities.length;
  sheet.getRange(totalRow, colStart).setValue("TOTAL");
  sheet.getRange(totalRow, colStart + 1).setValue(cities.length);
  sheet.getRange(totalRow, colStart + 2).setValue(
    cityAbove15.reduce((a, v) => a + v, 0)
  );
}
