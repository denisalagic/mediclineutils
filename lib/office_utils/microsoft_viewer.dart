import 'dart:convert';
import 'dart:io';

import 'package:archive/archive.dart';
import 'package:collection/collection.dart';
import 'package:flutter/foundation.dart';
import 'package:flutter/material.dart';
import 'package:flutter/services.dart';
import './domain/common_processor.dart';
import './domain/presentation_processor.dart';
import './domain/spreadsheet_processor.dart';
import './domain/word_processor.dart';
import './models/document.dart';
import './models/font_details.dart';
import './models/foot_end_note.dart';
import './models/presentation.dart';
import './models/relationship.dart';
import './models/spreadsheet.dart';
import './models/ss_color_schemes.dart';
import './models/ss_style.dart';
import './models/styles.dart';
import './utils/odttf.dart';
import './widget/progress_indicator.dart';
import 'package:path_provider/path_provider.dart';
import 'package:xml/xml.dart' as xml;
import 'models/shared_string.dart';

///The main dart file that takes the bytes data passed and then parses the files and displays the information.
class MicrosoftViewer extends StatefulWidget {
  final List<int> fileBytes;
  final bool fixedHeight;
  const MicrosoftViewer(this.fileBytes,this.fixedHeight, {super.key});

  @override
  State<StatefulWidget> createState() => MicrosoftViewerState();
}

class MicrosoftViewerState extends State<MicrosoftViewer> {
  ZipDecoder? _zipDecoder;
  String wordOutputDirectory = "";
  String spreadSheetOutputDirectory = "";
  String presentationOutputDirectory = "";
  String fileType = "";
  late Archive archive;
  late Document wordDocument;
  List<Relationship> relationShips = [];
  List<SharedString> sharedStrings = [];
  int elementDepth = 0;
  int seqNo = 0;
  Presentation presentation = Presentation("empty presentation document");
  SpreadSheet spreadSheet = SpreadSheet("empty spread sheet");
  List<Styles> stylesList = [];
  List<Widget> wordWidgets = [];

  List<Widget> presentationWidgets = [];
  List<FontDetails> fontList = [];
  List<FootEndNote> footNotes = [];
  List<FootEndNote> endNotes = [];

  List<SSStyle> spreadSheetStyles=[];
  List<Widget> spreadSheetWidgets = [];
  List<SSColorSchemes> spreadSheetColorSchemes=[];

  bool showProgressBar = true;
  @override
  void initState() {
    _zipDecoder ??= ZipDecoder();
    archive = _zipDecoder!.decodeBytes(widget.fileBytes);
    wordDocument = Document("empty word document", archive: archive);
    parseAndShowData();
    super.initState();
  }

  Future<void> parseAndShowData() async {
    await setupDirectory();
    if (archive.any((archiveFile) {
      return archiveFile.name == 'word/document.xml';
    })) {
      fileType = "word";
    } else if (archive.any((archiveFile) {
      return archiveFile.name == 'xl/workbook.xml';
    })) {
      setState(() {
        fileType = "spreadsheet";
      });
    } else if (archive.any((archiveFile) {
      return archiveFile.name == 'ppt/presentation.xml';
    })) {
      setState(() {
        fileType = "presentation";
      });
    }
    if (fileType == "word") {
      var relFile = archive.singleWhere((archiveFile) {
        return archiveFile.name.endsWith("document.xml.rels");
      });
      getRelationships(relFile);
      var mediaFile = archive.where((archiveFile) {
        return archiveFile.name.startsWith('word/media/');
      });
      for (var medFile in mediaFile) {
        extractMedia(medFile, wordOutputDirectory);
      }
      var stylesFile = archive.singleWhere((archiveFile) {
        return archiveFile.name.endsWith("word/styles.xml");
      });
      Map<String, String> defaultValues = {};
      CommonProcessor()
          .processStylesFile(stylesFile, stylesList, defaultValues);
      if (defaultValues.isNotEmpty) {
        if (defaultValues["fontSize"] != null) {
          wordDocument.defaultFontSize = int.parse(defaultValues["fontSize"]!);
        }
        if (defaultValues["lineSpacing"] != null) {
          wordDocument.defaultLineSpacing =
              int.parse(defaultValues["lineSpacing"]!);
        }
      }
      var fontTable = archive.singleWhereOrNull((archiveFile) {
        return archiveFile.name.endsWith("word/fontTable.xml");
      });
      if (fontTable != null) {
        var fontTableRel = archive.singleWhereOrNull((archiveFile) {
          return archiveFile.name.endsWith("_rels/fontTable.xml.rels");
        });
        if (fontTableRel != null) {
          CommonProcessor().processFonts(fontList, fontTable, fontTableRel);
          for (int i = 0; i < fontList.length; i++) {
            var fontFile = archive.singleWhereOrNull((archiveFile) {
              return archiveFile.name.endsWith(fontList[i].fileName);
            });
            if (fontFile != null) {
              String? fontKey =
                  fontList[i].fontKey.replaceAll("{", "").replaceAll("}", "");
              await loadFonts(fontFile, fontList[i].name, fontKey);
            }
          }
        }
      }
      var footNoteFile = archive.singleWhereOrNull((archiveFile) {
        return archiveFile.name.endsWith("footnotes.xml");
      });
      var endNoteFile = archive.singleWhereOrNull((archiveFile) {
        return archiveFile.name.endsWith("endnotes.xml");
      });
      WordProcessor()
          .processFootEndNotes(footNoteFile, endNoteFile, footNotes, endNotes);
      var wordFile = archive.singleWhere((archiveFile) {
        return archiveFile.name == 'word/document.xml';
      });
      await WordProcessor().processWordFile(wordFile, elementDepth, relationShips,
          wordOutputDirectory, stylesList, wordDocument, fileType);
      List<Widget> tempWidgets = await WordProcessor().displayWordFile(
          fileType, wordDocument, stylesList, footNotes, endNotes);
      setState(() {
        wordWidgets = tempWidgets;
        showProgressBar = false;
      });
    } else if (fileType == "spreadsheet") {
      var relFile = archive.singleWhere((archiveFile) {
        return archiveFile.name.endsWith("workbook.xml.rels");
      });
      getRelationships(relFile);
      var stylesFile=archive.singleWhereOrNull((archiveFile){
        return archiveFile.name.endsWith("xl/styles.xml");
      });
      if(stylesFile!=null) {
        spreadSheetStyles=[];
        SpreadsheetProcessor().processSpreadSheetStyles(stylesFile, spreadSheetStyles);
      }
      var themRel=relationShips.firstWhereOrNull((rel){
        return rel.target.startsWith("theme/theme");
      });
      if(themRel!=null){
        var themeFile=archive.singleWhereOrNull((archiveFile){
          return archiveFile.name.endsWith(themRel.target);
        });
        if(themeFile!=null) {
          spreadSheetColorSchemes = [];
          SpreadsheetProcessor().processColorSchemes(themeFile, spreadSheetColorSchemes);
        }
      }

      var shareStringsFile = archive.singleWhere((archiveFile) {
        return archiveFile.name.endsWith("sharedStrings.xml");
      });
      getSharedStrings(shareStringsFile);
      var workbookFile = archive.singleWhere((archiveFile) {
        return archiveFile.name.endsWith("xl/workbook.xml");
      });

      SpreadsheetProcessor().getSpreadSheetDetails(workbookFile, spreadSheet);
      await SpreadsheetProcessor().readAllSheets(spreadSheet, relationShips, archive);
      List<Widget> tempWidgets = await SpreadsheetProcessor()
          .displaySpreadSheet(spreadSheet, sharedStrings,spreadSheetStyles,spreadSheetColorSchemes);
      setState(() {
        spreadSheetWidgets = tempWidgets;
        showProgressBar = false;
      });
    } else if (fileType == "presentation") {
      var relFile = archive.singleWhere((archiveFile) {
        return archiveFile.name.endsWith("presentation.xml.rels");
      });
      getRelationships(relFile);
      var mediaFile = archive.where((archiveFile) {
        return archiveFile.name.startsWith('ppt/media/');
      });
      for (var medFile in mediaFile) {
        extractMedia(medFile, presentationOutputDirectory);
      }
      var presentationFile = archive.singleWhere((archiveFile) {
        return archiveFile.name.endsWith("ppt/presentation.xml");
      });
      PresentationProcessor()
          .getPresentationDetails(presentationFile, presentation);
      await PresentationProcessor().readAllSlides(
          presentation, relationShips, archive, presentationOutputDirectory);
      List<Widget> tempWidgets =
          await PresentationProcessor().displayPresentation(presentation);
      setState(() {
        presentationWidgets = tempWidgets;
        showProgressBar = false;
      });
    }
  }

  Future<void> setupDirectory() async {
    if (kIsWeb) {
      wordOutputDirectory = "";
      spreadSheetOutputDirectory = "";
      presentationOutputDirectory = "";
      return;
    }
    var applicationSupportDirectory = await getApplicationSupportDirectory();
    wordOutputDirectory = "${applicationSupportDirectory.path}/word/";
    spreadSheetOutputDirectory =
        "${applicationSupportDirectory.path}/spreadSheet/";
    presentationOutputDirectory =
        "${applicationSupportDirectory.path}/presentation/";
    var wordDir = Directory(wordOutputDirectory);
    if (wordDir.existsSync()) {
      wordDir.deleteSync(recursive: true);
    }
    wordDir.createSync(recursive: true);
    var xlsDir = Directory(spreadSheetOutputDirectory);
    if (xlsDir.existsSync()) {
      xlsDir.deleteSync(recursive: true);
    }
    xlsDir.createSync(recursive: true);
    var pptDir = Directory(presentationOutputDirectory);
    if (pptDir.existsSync()) {
      pptDir.deleteSync(recursive: true);
    }
    pptDir.createSync(recursive: true);
  }

  void getRelationships(ArchiveFile relFile) {
    final fileContent = utf8.decode(relFile.content);

    final document = xml.XmlDocument.parse(fileContent);
    final relationshipsElement = document.findAllElements("Relationship");
    relationShips = [];
    for (var rel in relationshipsElement) {
      if (rel.getAttribute("Id") != null) {
        relationShips.add(Relationship(rel.getAttribute("Id").toString(),
            rel.getAttribute("Target").toString()));
      }
    }
  }

  void getSharedStrings(ArchiveFile shareStringsFile) {
    final fileContent = utf8.decode(shareStringsFile.content);
    final document = xml.XmlDocument.parse(fileContent);
    sharedStrings = [];
    int index = 0;
    document.findAllElements('si').forEach((node) {
      sharedStrings
          .add(SharedString(index, node.getElement("t")?.innerText ?? ""));
      index++;
    });
  }

  Future<void> extractMedia(ArchiveFile mediaFile, String dirPath) async {
    if (kIsWeb) return;
    final String outputFilePath = dirPath + mediaFile.name.split("/").last;
    final File outFile = File(outputFilePath);
    await outFile.writeAsBytes(mediaFile.content as List<int>);
  }

  Future<void> loadFonts(
      ArchiveFile fontFile, String fontFamily, String fileName) async {
    ODTTF().deobfuscate(fontFile.content, fileName);
    var fontLoader = FontLoader(fontFamily)..addFont(getBytes(fontFile));
    await fontLoader.load();
  }

  Future<ByteData> getBytes(ArchiveFile fontFile) async {
    return ByteData.view(fontFile.content.buffer);
  }

  @override
  Widget build(BuildContext context) {
    return customWidget(
      LayoutBuilder(builder: (BuildContext context, BoxConstraints constraints){
        return Stack(
          children: [
            InteractiveViewer(
              minScale: 0.1,
              maxScale: 1.6,
              child: SingleChildScrollView(
                    child: Column(
                      crossAxisAlignment: CrossAxisAlignment.start,
                      children: fileType == "word"
                          ? wordWidgets
                          : fileType == "spreadsheet"
                          ? spreadSheetWidgets
                          : presentationWidgets,
                    ),
              ),
            ),
            showProgressBar ? ProgressIndicatorView(constraints.maxHeight,constraints.maxWidth) : Container()
          ],
        );
      }),
    );

  }
  Widget customWidget(Widget child) {
    // Always wrap in a Column to ensure Expanded has a proper parent
    return Column(
      children: [
        Expanded(child: child),
      ],
    );
  }
}
