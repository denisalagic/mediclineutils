import 'dart:ui';

import 'package:flutter/material.dart';

class ProgressIndicatorView extends StatefulWidget {
  final double height;
  final double width;
  const ProgressIndicatorView(this.height,this.width,{super.key});

  @override
  State<StatefulWidget> createState() => ProgressIndicatorState();
}

class ProgressIndicatorState extends State<ProgressIndicatorView>
    with SingleTickerProviderStateMixin {
  @override
  Widget build(BuildContext context) {

    return Container(
      color: Colors.grey.shade50,
      height: widget.height,
      width: widget.width,
      child: Column(
        mainAxisSize: MainAxisSize.max,
        mainAxisAlignment: MainAxisAlignment.center,
        children: [
          const CircularProgressIndicator(),
           SizedBox(
            height: widget.height*.1,
          ),
        ],
      ),
    );
  }
}
