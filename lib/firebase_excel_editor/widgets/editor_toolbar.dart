import 'package:flutter/material.dart';

class EditorToolbar extends StatelessWidget {
  final String? currentSheet;
  final List<String> availableSheets;
  final Function(String?) onSheetChanged;
  final VoidCallback onSave;

  const EditorToolbar({
    Key? key,
    required this.currentSheet,
    required this.availableSheets,
    required this.onSheetChanged,
    required this.onSave,
  }) : super(key: key);

  @override
  Widget build(BuildContext context) {
    return Container(
      padding: const EdgeInsets.all(8),
      color: Colors.grey[200],
      child: Row(
        children: [
          DropdownButton<String>(
            value: currentSheet,
            items: availableSheets.map((name) {
              return DropdownMenuItem(value: name, child: Text(name));
            }).toList(),
            onChanged: onSheetChanged,
          ),
          const Spacer(),
          ElevatedButton.icon(
            onPressed: onSave,
            icon: const Icon(Icons.cloud_upload),
            label: const Text('Save'),
          ),
        ],
      ),
    );
  }
}