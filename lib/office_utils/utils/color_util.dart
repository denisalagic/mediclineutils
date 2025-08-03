import 'package:collection/collection.dart';

import '../data/indexed_color.dart';
import '../models/ss_color_schemes.dart';

class ColorUtil {
  static String? resolveExcelColor(
      String bgClrIndex,
      String themeId,
      List<SSColorSchemes> colorSchemes
      ) {
    // First, try indexed color
    if (bgClrIndex.isNotEmpty) {
      try {
        int index = int.parse(bgClrIndex);
        if (index < 64) {
          return IndexedColor().colors[index]; // always #xxxxxx
        } else if (index == 64) {
          // Now try theme
          var clrScheme = colorSchemes.firstWhereOrNull((clrSch) => clrSch.id == themeId);
          if (clrScheme != null) {
            String? hex = clrScheme.srgbClr.isNotEmpty
                ? clrScheme.srgbClr
                : clrScheme.sysClrLast;

            if (hex.isNotEmpty) {
              // Remove any extra '#' just in case
              hex = hex.replaceAll("#", "").toUpperCase();
              // Take last 6 digits (Excel colors sometimes include alpha)
              if (hex.length > 6) hex = hex.substring(hex.length - 6);
              return "#$hex";
            }
          }
          return "#FFFFFF"; // fallback to white
        }
      } catch (_) {
        return "#FFFFFF"; // fallback
      }
    }

    return null;
  }

}