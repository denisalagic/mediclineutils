import 'package:flutter/material.dart';

class EditableCell extends StatefulWidget {
  final String value;
  final Function(String) onChanged;

  const EditableCell({
    super.key,
    required this.value,
    required this.onChanged,
  });

  @override
  State<EditableCell> createState() => _EditableCellState();
}

class _EditableCellState extends State<EditableCell> {
  late TextEditingController _controller;
  bool _isEditing = false;

  @override
  void initState() {
    super.initState();
    _controller = TextEditingController(text: widget.value);
  }

  @override
  void didUpdateWidget(EditableCell oldWidget) {
    super.didUpdateWidget(oldWidget);
    if (widget.value != oldWidget.value && !_isEditing) {
      _controller.text = widget.value;
    }
  }

  @override
  Widget build(BuildContext context) {
    return GestureDetector(
      onDoubleTap: () {
        setState(() {
          _isEditing = true;
        });
      },
      child: _isEditing
          ? TextField(
        controller: _controller,
        autofocus: true,
        onSubmitted: (value) {
          widget.onChanged(value);
          setState(() {
            _isEditing = false;
          });
        },
        onTapOutside: (_) {
          widget.onChanged(_controller.text);
          setState(() {
            _isEditing = false;
          });
        },
      )
          : Text(widget.value),
    );
  }

  @override
  void dispose() {
    _controller.dispose();
    super.dispose();
  }
}