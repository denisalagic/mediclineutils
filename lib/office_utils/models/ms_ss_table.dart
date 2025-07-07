///Classes for storing details of SpreadSheet tables.
class MsSsTable {
  ///List of columns
  List<MsSsCol> cols = [];

  ///List of rows
  List<MsSsRow> rows = [];
}

class MsSsCol {
  int min;
  int max;
  double width;
  int customWidth;
  MsSsCol(this.min, this.max, this.width, this.customWidth);
}

class MsSsRow {
  int rowId;
  String spans;
  double height;
  String? style;
  List<MsSsCell> cells = [];
  MsSsRow(this.rowId, this.spans, this.height);
}

class MsSsCell {
  int colNo;
  String type;
  String value;
  String? style;
  int colSpan;
  MsSsCell(this.colNo, this.type, this.value,this.colSpan);
}
