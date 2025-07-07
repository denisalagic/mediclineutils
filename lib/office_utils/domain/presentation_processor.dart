import 'dart:convert';
import 'dart:io';

import 'package:archive/archive.dart';
import 'package:collection/collection.dart';
import 'package:flutter/foundation.dart';
import 'package:flutter/material.dart';
import 'package:xml/xml.dart' as xml;

import '../models/presentation.dart';
import '../models/presentation_paragraph.dart';
import '../models/presentation_shape.dart';
import '../models/presentation_text.dart';
import '../models/presentation_text_box.dart';
import '../models/relationship.dart';
import '../models/slide.dart';

///Class for processing .pptx files
class PresentationProcessor {
  ///Function for processing .pptx file
  void getPresentationDetails(
      ArchiveFile presentationFile, Presentation presentation) {
    final fileContent = utf8.decode(presentationFile.content);
    final presentationDoc = xml.XmlDocument.parse(fileContent);
    var slidesRoot = presentationDoc.findAllElements("p:sldIdLst");
    if (slidesRoot.isNotEmpty) {
      var slides = slidesRoot.first.findAllElements("p:sldId");
      if (slides.isNotEmpty) {
        for (var slide in slides) {
          int id = 0;
          String rId = "";
          var tempId = slide.getAttribute("id");
          if (tempId != null) {
            id = int.parse(tempId);
          }
          var tempRid = slide.getAttribute("r:id");
          if (tempRid != null) {
            rId = tempRid;
          }
          presentation.slides.add(Slide(id, rId, ""));
        }
      }
    }
    var masterSlidesRoot = presentationDoc.findAllElements("p:sldMasterIdLst");
    if (masterSlidesRoot.isNotEmpty) {
      var masterSlides =
          masterSlidesRoot.first.findAllElements("p:sldMasterId");
      if (masterSlides.isNotEmpty) {
        for (var slide in masterSlides) {
          int id = 0;
          String rId = "";
          var tempId = slide.getAttribute("id");
          if (tempId != null) {
            id = int.parse(tempId);
          }
          var tempRid = slide.getAttribute("r:id");
          if (tempRid != null) {
            rId = tempRid;
          }
          presentation.masterSlides.add(Slide(id, rId, ""));
        }
      }
    }
  }

  ///Function for processing all shapes
  void getAllShapes(ArchiveFile presentationFile, Slide slide) {
    final fileContent = utf8.decode(presentationFile.content);
    final diagramDoc = xml.XmlDocument.parse(fileContent);
    var diagramsRoot = diagramDoc.findAllElements("dsp:sp");
    if (diagramsRoot.isNotEmpty) {
      for (var diagram in diagramsRoot) {
        String id = "";
        String text = "";
        double offsety = 0;
        Offset offset = const Offset(0, 0);
        Size size = const Size(0, 0);
        var tempId = diagram.getAttribute("modelId");
        if (tempId != null) {
          id = tempId;
        }
        var checkTxtBody = diagram.findAllElements("dsp:txBody");
        if (checkTxtBody.isNotEmpty) {
          var checkParaElement = checkTxtBody.first.findAllElements("a:p");
          if (checkParaElement.isNotEmpty) {
            var txtElement = checkParaElement.first.findAllElements("a:t");
            if (txtElement.isNotEmpty) {
              text = txtElement.first.innerText;
            }
          }
        }
        var checkSlFrm = diagram.findAllElements("a:xfrm");
        if (checkSlFrm.isNotEmpty) {
          var checkOffE = checkSlFrm.first.findAllElements("a:off");
          if (checkOffE.isNotEmpty) {
            var tempX = checkOffE.first.getAttribute("x");
            if (tempX != null) {
            }
            var tempY = checkOffE.first.getAttribute("y");
            if (tempY != null) {
              offsety = double.parse(tempY);
            }
          }
        }

        var checkTxFrm = diagram.findAllElements("dsp:txXfrm");
        if (checkTxFrm.isNotEmpty) {
          var checkOffE = checkTxFrm.first.findAllElements("a:off");
          if (checkOffE.isNotEmpty) {
            double x = 0;
            double y = 0;
            var tempX = checkOffE.first.getAttribute("x");
            if (tempX != null) {
              x = double.parse(tempX);
            }
            var tempY = checkOffE.first.getAttribute("y");
            if (tempY != null) {
              y = double.parse(tempY);
            }
            offset = Offset(x, y + offsety);
          }
          var checkExtE = checkTxFrm.first.findAllElements("a:ext");
          if (checkExtE.isNotEmpty) {
            double x = 0;
            double y = 0;
            var tempX = checkExtE.first.getAttribute("cx");
            if (tempX != null) {
              x = double.parse(tempX);
            }
            var tempY = checkExtE.first.getAttribute("cy");
            if (tempY != null) {
              y = double.parse(tempY);
            }
            size = Size(x, y);
          }
        }
        slide.presentationShapes.add(PresentationShape(id, text, offset, size));
      }
    }
  }

  ///Function for processing slides
  Future<void> readAllSlides(
      Presentation presentation,
      List<Relationship> relationShips,
      Archive archive,
      String presentationOutputDirectory)async {
    for (int i = 0; i < presentation.slides.length; i++) {
      var slideRelation = relationShips.firstWhereOrNull((rel) {
        return rel.id == presentation.slides[i].rId;
      });
      if (slideRelation != null) {
        var slideFile = archive.singleWhere((archiveFile) {
          return archiveFile.name.endsWith(slideRelation.target);
        });
        if (slideFile.isFile) {
          final fileContent = utf8.decode(slideFile.content);
          final slideDoc = xml.XmlDocument.parse(fileContent);
          presentation.slides[i].fileName = slideFile.name.split("/").last;
          var spElement = slideDoc.findAllElements("p:sp");
          if (spElement.isNotEmpty) {
            for (int j = 0; j < spElement.length; j++) {
              var tempTextBox=await compute(getPresentationTextBox, spElement.elementAt(j));
              if(tempTextBox!=null){
                presentation.slides[i].presentationTextBoxes.add(tempTextBox);
              }
            }
          }

          var checkSlideRel = archive.singleWhereOrNull((archiveFile) {
            return archiveFile.name
                .endsWith("${presentation.slides[i].fileName}.rels");
          });
          if (checkSlideRel != null) {
            List<Relationship> slideLevelRelations = [];
            final fileContent = utf8.decode(checkSlideRel.content);
            String drawingTarget = "";
            String layoutTarget = "";
            final document = xml.XmlDocument.parse(fileContent);
            final relationshipsElement =
                document.findAllElements("Relationship");
            for (var rel in relationshipsElement) {
              if (rel.getAttribute("Id") != null) {
                slideLevelRelations.add(Relationship(
                    rel.getAttribute("Id").toString(),
                    rel.getAttribute("Target").toString()));
              }
              if (rel.getAttribute("Type") != null &&
                  rel
                      .getAttribute("Type")!
                      .endsWith("relationships/diagramDrawing")) {
                drawingTarget =
                    rel.getAttribute("Target").toString().replaceAll("../", "");
              }
              if (rel.getAttribute("Type") != null &&
                  rel
                      .getAttribute("Type")!
                      .endsWith("relationships/slideLayout")) {
                layoutTarget =
                    rel.getAttribute("Target").toString().replaceAll("../", "");
              }
            }
            if (drawingTarget.isNotEmpty) {
              var diagramFile = archive.singleWhereOrNull((archiveFile) {
                return archiveFile.name.endsWith(drawingTarget);
              });
              if (diagramFile != null) {
                getAllShapes(diagramFile, presentation.slides[i]);
              }
            }
            if (layoutTarget.isNotEmpty) {
              var checkLayoutRel = archive.singleWhereOrNull((archiveFile) {
                return archiveFile.name
                    .endsWith("${layoutTarget.split("/").last}.rels");
              });
              List<Relationship> layoutRelations = [];
              if (checkLayoutRel != null) {
                final fileContent2 = utf8.decode(checkLayoutRel.content);
                final document2 = xml.XmlDocument.parse(fileContent2);
                final relationshipsElement2 =
                    document2.findAllElements("Relationship");

                for (var rel in relationshipsElement2) {
                  if (rel.getAttribute("Id") != null) {
                    layoutRelations.add(Relationship(
                        rel.getAttribute("Id").toString(),
                        rel.getAttribute("Target").toString()));
                  }
                }
              }
              var layoutFile = archive.singleWhereOrNull((archiveFile) {
                return archiveFile.name.endsWith(layoutTarget);
              });
              if (layoutFile != null) {
                final fileContent3 = utf8.decode(layoutFile.content);
                final document3 = xml.XmlDocument.parse(fileContent3);
                var chkBg = document3.findAllElements("p:bg");
                if (chkBg.isNotEmpty) {
                  var chkBlip = chkBg.first.findAllElements("a:blip");
                  if (chkBlip.isNotEmpty) {
                    var chkEmbed = chkBlip.first.getAttribute("r:embed");
                    if (chkEmbed != null) {
                      var layoutRelTarget =
                          layoutRelations.firstWhereOrNull((rel) {
                        return rel.id == chkEmbed;
                      });
                      if (layoutRelTarget != null) {
                        presentation.slides[i].backgroundImagePath =
                            "$presentationOutputDirectory/${layoutRelTarget.target.split("/").last}";
                      }
                    }
                  }
                }
              }
            }
          }
        }
      }
    }
  }

  ///Function for displaying the presentations
  Future<List<Widget>> displayPresentation(Presentation presentation) async {
    List<Widget> tempList = [];
    List<Widget> slideWidgets = [];
    for (int i = 0; i < presentation.slides.length; i++) {
      List<Widget> tempSlideWidget =
          await compute(getSlideDetails, presentation.slides[i]);
      slideWidgets.addAll(tempSlideWidget);
    }
    tempList.add(Container(
      color: Colors.grey,
      width: double.infinity,
      margin: const EdgeInsets.all(8),
      child: Column(
        children: slideWidgets,
      ),
    ));
    return tempList;
  }

  ///Function for getting slide details
  static List<Widget> getSlideDetails(Slide slide) {
    List<Widget> tempSlide = [];
    List<Widget> tempShapes = [];
    List<Widget> slideWidget = [];
    double maxWidth = 600;
    double maxHeight = 450;
    int divisionFactor = 12700;
    for (int j = 0; j < slide.presentationTextBoxes.length; j++) {
      if (slide.presentationTextBoxes[j].offset.dx / divisionFactor +
              slide.presentationTextBoxes[j].size.width / divisionFactor >
          maxWidth) {
        maxWidth = slide.presentationTextBoxes[j].offset.dx / divisionFactor +
            slide.presentationTextBoxes[j].size.width / divisionFactor;
      }
      if (slide.presentationTextBoxes[j].offset.dy / divisionFactor +
              slide.presentationTextBoxes[j].size.height / divisionFactor >
          maxHeight) {
        maxHeight = slide.presentationTextBoxes[j].offset.dy / divisionFactor +
            slide.presentationTextBoxes[j].size.height / divisionFactor;
      }
      List<Widget> textBoxTexts = [];
      for (int k = 0;
          k < slide.presentationTextBoxes[j].presentationParas.length;
          k++) {
        List<TextSpan> textSpans = [];
        for (int l = 0;
            l <
                slide.presentationTextBoxes[j].presentationParas[k].textSpans
                    .length;
            l++) {
          textSpans.add(TextSpan(
              text: slide.presentationTextBoxes[j].presentationParas[k]
                  .textSpans[l].text,
              style: TextStyle(
                  fontSize: slide.presentationTextBoxes[j].presentationParas[k]
                      .textSpans[l].fontSize,
                  color: Colors.black)));
        }
        textBoxTexts.add(RichText(text: TextSpan(children: textSpans)));
      }
      if (slide.presentationTextBoxes[j].size.height != 0 &&
          slide.presentationTextBoxes[j].size.width != 0) {
        tempShapes.add(Positioned(
            top: slide.presentationTextBoxes[j].offset.dy / divisionFactor,
            left: slide.presentationTextBoxes[j].offset.dx / divisionFactor,
            child: SizedBox(
              height:
                  slide.presentationTextBoxes[j].size.height / divisionFactor,
              width: slide.presentationTextBoxes[j].size.width / divisionFactor,
              child: Column(
                children: textBoxTexts,
              ),
            )));
      } else {
        tempShapes.add(Positioned(
            top: slide.presentationTextBoxes[j].offset.dy / divisionFactor,
            left: slide.presentationTextBoxes[j].offset.dx / divisionFactor,
            child: Column(
              children: textBoxTexts,
            )));
      }
    }
    for (int j = 0; j < slide.presentationShapes.length; j++) {
      if (slide.presentationShapes[j].offset.dx / divisionFactor +
              slide.presentationShapes[j].size.width / divisionFactor >
          maxWidth) {
        maxWidth = slide.presentationShapes[j].offset.dx / divisionFactor +
            slide.presentationShapes[j].size.width / divisionFactor;
      }
      if (slide.presentationShapes[j].offset.dy / divisionFactor +
              slide.presentationShapes[j].size.height / divisionFactor >
          maxHeight) {
        maxHeight = slide.presentationShapes[j].offset.dy / divisionFactor +
            slide.presentationShapes[j].size.height / divisionFactor;
      }
      tempShapes.add(Positioned(
          top: slide.presentationShapes[j].offset.dy / divisionFactor,
          left: slide.presentationShapes[j].offset.dx / divisionFactor,
          child: Container(
            decoration: BoxDecoration(border: Border.all(color: Colors.blue)),
            height: slide.presentationShapes[j].size.height / divisionFactor,
            width: slide.presentationShapes[j].size.width / divisionFactor,
            child: Center(
                child: Text(
              slide.presentationShapes[j].text,
            )),
          )));
    }
    if (tempShapes.isNotEmpty) {
      tempSlide.add(SizedBox(
          height: maxHeight,
          width: maxWidth,
          child: Stack(
            children: tempShapes,
          )));
    }
    slideWidget.add(SingleChildScrollView(
      scrollDirection: Axis.horizontal,
      child: Container(
        constraints: const BoxConstraints(minHeight: 450),
        decoration: slide.backgroundImagePath != ""
            ? BoxDecoration(
                color: Colors.white,
                image: DecorationImage(
                  image: FileImage(File(slide.backgroundImagePath)),
                  fit: BoxFit.fill,
                ),
              )
            : const BoxDecoration(
                color: Colors.white,
              ),
        width: maxWidth,
        margin: const EdgeInsets.all(8),
        child: Column(
          children: tempSlide,
        ),
      ),
    ));
    return slideWidget;
  }

  static PresentationTextBox? getPresentationTextBox(xml.XmlElement spElement){
    Offset offset = const Offset(0, 0);
    Size size = const Size(0, 0);
    double offsetY = 0;
    double offsetX = 0;
    if (spElement.parentElement != null &&
        spElement.parentElement?.name.toString() ==
            "p:grpSp") {
      var grpSpPr = spElement.parentElement
          ?.findAllElements("p:grpSpPr");
      if (grpSpPr != null && grpSpPr.isNotEmpty) {
        var chckOff = grpSpPr.first.findAllElements("a:off");
        if (chckOff.isNotEmpty) {
          var offX = chckOff.first.getAttribute("x");
          if (offX != null) {
            offsetX = double.parse(offX);
          }
          var offY = chckOff.first.getAttribute("y");
          if (offY != null) {
            offsetY = double.parse(offY);
          }
        }
      }
    }
    var xfrmElement =
    spElement.findAllElements("a:xfrm");
    if (xfrmElement.isNotEmpty) {
      var chkOff = xfrmElement.first.findAllElements("a:off");
      if (chkOff.isNotEmpty) {
        var offX = chkOff.first.getAttribute("x");
        var offY = chkOff.first.getAttribute("y");
        if (offX != null && offY != null) {
          offset = Offset(double.parse(offX) + offsetX,
              double.parse(offY) + offsetY);
        }
      }
      var chkExt = xfrmElement.first.findAllElements("a:ext");
      if (chkExt.isNotEmpty) {
        var extX = chkExt.first.getAttribute("cx");
        var extY = chkExt.first.getAttribute("cy");
        if (extX != null && extY != null) {
          size = Size(double.parse(extX), double.parse(extY));
        }
      }
    }
    List<PresentationParagraph> presentationParagraphs = [];
    spElement.findAllElements("p:txBody").forEach((txt) {
      var chkPara = txt.findAllElements("a:p");
      List<PresentationText> presentationTexts = [];
      if (chkPara.isNotEmpty) {
        for (var para in chkPara) {
          presentationTexts = [];
          var chkR = para.findAllElements("a:r");
          if (chkR.isNotEmpty) {
            for (var r in chkR) {
              double fontSize = 20;
              var rPr = r.findAllElements("a:rPr");
              if (rPr.isNotEmpty) {
                var tempSize = rPr.first.getAttribute("sz");
                if (tempSize != null) {
                  fontSize = double.parse(tempSize) / 150;
                }
              }
              var text = "";
              r.findAllElements("a:t").forEach((txt2) {
                text += txt2.innerText;
              });

              if (text.isNotEmpty) {
                presentationTexts
                    .add(PresentationText(text, fontSize));
              }
            }
          }
          if (presentationTexts.isNotEmpty) {
            PresentationParagraph paragraph = PresentationParagraph();
            paragraph.textSpans = presentationTexts;
            presentationParagraphs.add(paragraph);
          }
        }
      }
    });
    if (presentationParagraphs.isNotEmpty) {
      PresentationTextBox presentationTextBox =
      PresentationTextBox(offset, size);
      presentationTextBox.presentationParas = presentationParagraphs;
      return presentationTextBox;
    }else {
      return null;
    }


  }
}
