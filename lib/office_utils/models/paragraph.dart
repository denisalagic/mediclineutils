import 'ms_image.dart';
import 'ms_text_span.dart';

///Class for storing word paragraph details.
class Paragraph {
  int seqNo;
  String style;
  List<MsTextSpan> textSpans = [];
  List<MsImage> images = [];
  Map<String, String> tabDetails = {};
  String? shadingColor;
  Map<String, String>? formats;
  Paragraph(this.seqNo, this.style);
}
