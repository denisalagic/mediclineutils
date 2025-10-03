import 'dart:async';
import 'dart:io';
import 'package:firebase_storage/firebase_storage.dart';
import 'package:firebase_database/firebase_database.dart';
import 'package:excel/excel.dart';
import 'package:path_provider/path_provider.dart';
import '../models/cell_change.dart';

class ExcelSyncManager {
  final String storagePath;
  final String sessionId;
  final FirebaseStorage storage;
  final FirebaseDatabase database;

  final _changesController = StreamController<CellChange>.broadcast();
  Stream<CellChange> get changesStream => _changesController.stream;

  ExcelSyncManager({
    required this.storagePath,
    required this.sessionId,
    required this.database,
    required this.storage
  }) {
    _listenToChanges();
  }

  void _listenToChanges() {
    database.ref('excel_changes/$sessionId').onChildAdded.listen((event) {
      final data = event.snapshot.value as Map<dynamic, dynamic>?;
      if (data != null) {
        _changesController.add(CellChange.fromJson(data));
      }
    });
  }

  Future<Excel> loadExcel() async {
    final ref = storage.ref(storagePath);
    final dir = await getTemporaryDirectory();
    final file = File('${dir.path}/temp_excel.xlsx');

    await ref.writeToFile(file);
    final bytes = await file.readAsBytes();
    return Excel.decodeBytes(bytes);
  }

  Future<void> saveExcel(Excel excel) async {
    final bytes = excel.encode();
    if (bytes == null) throw Exception('Failed to encode Excel');

    final dir = await getTemporaryDirectory();
    final file = File('${dir.path}/temp_excel.xlsx');
    await file.writeAsBytes(bytes);

    final ref = storage.ref(storagePath);
    await ref.putFile(file);
  }

  Future<void> syncCellChange(CellChange change) async {
    await database
        .ref('excel_changes/$sessionId')
        .push()
        .set(change.toJson());
  }

  void dispose() {
    _changesController.close();
  }
}