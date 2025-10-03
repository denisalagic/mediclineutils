import 'dart:async';
import 'package:firebase_database/firebase_database.dart';
import 'package:firebase_storage/firebase_storage.dart';
import 'package:flutter/material.dart';
import 'package:excel/excel.dart';
import 'models/cell_change.dart';
import 'services/excel_sync_manager.dart';
import 'widgets/editor_toolbar.dart';
import 'widgets/spreadsheet_view.dart';

class FirebaseExcelEditor extends StatefulWidget {
  final String fileStoragePath;
  final String sessionId;
  final FirebaseStorage storage;
  final FirebaseDatabase database;

  const FirebaseExcelEditor({
    super.key,
    required this.fileStoragePath,
    required this.sessionId,
    required this.database,
    required this.storage
  });

  @override
  State<FirebaseExcelEditor> createState() => _FirebaseExcelEditorState();
}

class _FirebaseExcelEditorState extends State<FirebaseExcelEditor> {
  Excel? _excel;
  String? _currentSheet;
  bool _isLoading = true;
  String? _error;
  late ExcelSyncManager _syncManager;
  StreamSubscription? _changesSubscription;

  @override
  void initState() {
    super.initState();
    _syncManager = ExcelSyncManager(
      storagePath: widget.fileStoragePath,
      sessionId: widget.sessionId,
      database: widget.database,
      storage: widget.storage,
    );
    _loadExcelFile();
  }

  Future<void> _loadExcelFile() async {
    try {
      setState(() {
        _isLoading = true;
        _error = null;
      });

      _excel = await _syncManager.loadExcel();

      if (_excel != null && _excel!.tables.isNotEmpty) {
        _currentSheet = _excel!.tables.keys.first;
      }

      _changesSubscription = _syncManager.changesStream.listen((change) {
        _applyRemoteChange(change);
      });

      setState(() {
        _isLoading = false;
      });
    } catch (e) {
      setState(() {
        _error = e.toString();
        _isLoading = false;
      });
    }
  }

  void _applyRemoteChange(CellChange change) {
    if (_excel == null || change.sheet != _currentSheet) return;

    final sheet = _excel!.tables[change.sheet];
    if (sheet == null) return;

    // Convert the dynamic value to appropriate CellValue type
    CellValue? cellValue;
    if (change.value == null || change.value == '') {
      cellValue = null;
    } else if (change.value is num) {
      cellValue = IntCellValue(change.value as int);
    } else if (change.value is double) {
      cellValue = DoubleCellValue(change.value as double);
    } else if (change.value is bool) {
      cellValue = BoolCellValue(change.value as bool);
    } else if (change.value.toString().startsWith('=')) {
      cellValue = FormulaCellValue(change.value.toString());
    } else {
      cellValue = TextCellValue(change.value.toString());
    }

    sheet.updateCell(
      CellIndex.indexByColumnRow(
        columnIndex: change.col,
        rowIndex: change.row,
      ),
      cellValue,
    );

    setState(() {});
  }

  Future<void> _updateCell(int row, int col, String value) async {
    if (_excel == null || _currentSheet == null) return;

    final sheet = _excel!.tables[_currentSheet!];
    if (sheet == null) return;

    // Convert the string value to appropriate CellValue type
    CellValue? cellValue;
    if (value.isEmpty) {
      cellValue = null;
    } else if (value.startsWith('=')) {
      cellValue = FormulaCellValue(value);
    } else {
      // Try to parse as number
      final numValue = num.tryParse(value);
      if (numValue != null) {
        cellValue = numValue is int
            ? IntCellValue(numValue)
            : DoubleCellValue(numValue as double);
      } else if (value.toLowerCase() == 'true' || value.toLowerCase() == 'false') {
        cellValue = BoolCellValue(value.toLowerCase() == 'true');
      } else {
        cellValue = TextCellValue(value);
      }
    }

    sheet.updateCell(
      CellIndex.indexByColumnRow(columnIndex: col, rowIndex: row),
      cellValue,
    );

    setState(() {});

    await _syncManager.syncCellChange(
      CellChange(
        sheet: _currentSheet!,
        row: row,
        col: col,
        value: value,
        timestamp: DateTime.now().millisecondsSinceEpoch,
      ),
    );
  }

  Future<void> _saveToStorage() async {
    if (_excel == null) return;

    try {
      await _syncManager.saveExcel(_excel!);
      if (mounted) {
        ScaffoldMessenger.of(context).showSnackBar(
          const SnackBar(content: Text('Saved successfully!')),
        );
      }
    } catch (e) {
      if (mounted) {
        ScaffoldMessenger.of(context).showSnackBar(
          SnackBar(content: Text('Error saving: $e')),
        );
      }
    }
  }

  @override
  void dispose() {
    _changesSubscription?.cancel();
    _syncManager.dispose();
    super.dispose();
  }

  @override
  Widget build(BuildContext context) {
    if (_isLoading) {
      return const Center(child: CircularProgressIndicator());
    }

    if (_error != null) {
      return Center(child: Text('Error: $_error'));
    }

    if (_excel == null || _currentSheet == null) {
      return const Center(child: Text('No data available'));
    }

    final sheet = _excel!.tables[_currentSheet!];
    if (sheet == null) {
      return const Center(child: Text('Sheet not found'));
    }

    return Column(
      children: [
        EditorToolbar(
          currentSheet: _currentSheet,
          availableSheets: _excel!.tables.keys.toList(),
          onSheetChanged: (value) {
            setState(() {
              _currentSheet = value;
            });
          },
          onSave: _saveToStorage,
        ),
        Expanded(
          child: SpreadsheetView(
            sheet: sheet,
            onCellChanged: _updateCell,
          ),
        ),
      ],
    );
  }
}