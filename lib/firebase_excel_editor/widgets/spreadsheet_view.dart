import 'package:flutter/material.dart';
import 'package:excel/excel.dart';
import 'editable_cell.dart';

class SpreadsheetView extends StatelessWidget {
  final Sheet sheet;
  final Function(int row, int col, String value) onCellChanged;

  const SpreadsheetView({
    Key? key,
    required this.sheet,
    required this.onCellChanged,
  }) : super(key: key);

  @override
  Widget build(BuildContext context) {
    final maxRows = sheet.maxRows;
    final maxCols = sheet.maxColumns;

    return SingleChildScrollView(
      scrollDirection: Axis.horizontal,
      child: SingleChildScrollView(
        child: DataTable(
          columns: List.generate(
            maxCols,
                (i) => DataColumn(
              label: Text(String.fromCharCode(65 + i)),
            ),
          ),
          rows: List.generate(
            maxRows,
                (rowIndex) => DataRow(
              cells: List.generate(
                maxCols,
                    (colIndex) {
                  final cell = sheet.cell(
                    CellIndex.indexByColumnRow(
                      columnIndex: colIndex,
                      rowIndex: rowIndex,
                    ),
                  );
                  return DataCell(
                    EditableCell(
                      value: cell.value?.toString() ?? '',
                      onChanged: (value) => onCellChanged(rowIndex, colIndex, value),
                    ),
                  );
                },
              ),
            ),
          ),
        ),
      ),
    );
  }
}