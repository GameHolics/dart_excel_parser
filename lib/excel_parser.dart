import 'dart:convert';
import 'dart:typed_data';

import 'package:excel/excel.dart';
import 'package:tuple/tuple.dart';

typedef DataCell = Tuple2<CellType, dynamic>;
typedef DataRow = List<DataCell>;

bool canParse(String fileName) => fileName.endsWith('.xlsx');

List<DataRow> parseRaw(Uint8List data) {
  final excel = Excel.decodeBytes(data);
  if (excel.sheets.isEmpty) throw 'File has no sheets!';
  final sheetName = excel.getDefaultSheet();
  var sheet = excel.sheets[sheetName];
  if (sheet == null) throw '$sheetName not exist!';
  final rowNum = sheet.maxRows;
  if (rowNum == 0) throw '$sheetName has not rows!';
  final rows = sheet.rows;
  final columnNum = sheet.maxCols;
  final parsed = <DataRow>[];
  for (var r = 0; r < rowNum; r++) {
    var row = rows[r];
    if (r >= parsed.length) {
      parsed.add(<DataCell>[]);
    }
    var dataRow = parsed[r];
    for (var c = 0; c < columnNum; c++) {
      final cell = c < row.length ? row[c] : null;
      dataRow.add(_parseCell(cell));
    }
  }
  return parsed;
}

DataCell _parseCell(Data? data) => data == null ? DataCell(CellType.String, '') : DataCell(data.cellType, data.value);

dynamic _validateDataCell(DataCell dataCell) => dataCell.item1 == CellType.Formula ? '' : dataCell.item2;

String rawDataToJson(List<DataRow> rawData, bool prettify) {
  return JsonEncoder.withIndent(
      prettify ? '  ' : null, (dynamic obj) => obj is DataCell ? _validateDataCell(obj) : obj.toString()).convert(rawData);
}

String rawDataToSeparatedText(List<DataRow> rawData, String cellSeparator, {String lineSeparator = '\n'}) {
  final sb = StringBuffer();
  for (var dataRow in rawData) {
    sb.writeAll(dataRow.map(_validateDataCell), cellSeparator);
    sb.write(lineSeparator);
  }
  return sb.toString();
}
