import 'dart:typed_data';

///Class for storing image details
class MsImage {
  ///Image sequence number
  int pSeqNo;

  ///Image path
  String imagePath;

  ///Image type
  String type;

  ///Image width
  int cx;

  ///Image height
  int cy;

  ///Image bytes (for web)
  Uint8List? bytes;

  MsImage(this.pSeqNo, this.imagePath, this.type, this.cx, this.cy, [this.bytes]);
}
