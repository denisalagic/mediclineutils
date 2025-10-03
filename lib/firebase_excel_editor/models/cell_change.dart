class CellChange {
  final String sheet;
  final int row;
  final int col;
  final dynamic value;
  final int timestamp;

  CellChange({
    required this.sheet,
    required this.row,
    required this.col,
    required this.value,
    required this.timestamp,
  });

  Map<String, dynamic> toJson() => {
    'sheet': sheet,
    'row': row,
    'col': col,
    'value': value,
    'timestamp': timestamp,
  };

  factory CellChange.fromJson(Map<dynamic, dynamic> json) => CellChange(
    sheet: json['sheet'],
    row: json['row'],
    col: json['col'],
    value: json['value'],
    timestamp: json['timestamp'],
  );
}
