import 'presentation_shape.dart';
import 'presentation_text_box.dart';

///Class for storing slide details in pptx.
class Slide {
  int id;
  String rId;
  String fileName;
  String backgroundImagePath = "";
  List<PresentationShape> presentationShapes = [];
  List<PresentationTextBox> presentationTextBoxes = [];
  Slide(this.id, this.rId, this.fileName);
}
