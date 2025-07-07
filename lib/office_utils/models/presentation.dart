import 'slide.dart';

///Class for storing presentation details.
class Presentation {
  String name;
  List<Slide> slides = [];
  List<Slide> masterSlides = [];
  Presentation(this.name);
}
