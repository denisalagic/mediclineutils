import 'package:collection/collection.dart';

import '../data/indexed_color.dart';
import '../models/ss_color_schemes.dart';

class ColorUtil {
  static String? resolveExcelColor(
      String bgClrIndex,
      String themeId,
      List<SSColorSchemes> colorSchemes
      ) {
    if (bgClrIndex.isNotEmpty) {
      try {
        int index = int.parse(bgClrIndex);
        if (index < 64) {
          return IndexedColor().colors[index];
        } else if (index == 64 && themeId.isNotEmpty) {
          themeId = themeId.toLowerCase().replaceAll("a:", "").trim();

          var clrScheme = colorSchemes.firstWhereOrNull((clrSch) => clrSch.id == themeId);
          if (clrScheme != null) {
            String? hex = clrScheme.srgbClr.isNotEmpty
                ? clrScheme.srgbClr
                : clrScheme.sysClrLast;
            if (hex.isNotEmpty) {
              hex = hex.replaceAll("#", "");
              if (hex.length > 6) hex = hex.substring(hex.length - 6);
              return "#$hex";
            }
          }
        }
      } catch (_) {
        return "#FFFFFF";
      }
    }
    return "#FFFFFF";
  }


}