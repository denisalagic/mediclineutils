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

  static String applyTint(String hexColor, double tint) {
    // Convert hex color to RGB
    int r = int.parse(hexColor.substring(1, 3), radix: 16);
    int g = int.parse(hexColor.substring(3, 5), radix: 16);
    int b = int.parse(hexColor.substring(5, 7), radix: 16);

    // Apply tint adjustment
    if (tint > 0) {
      r = r + ((255 - r) * tint).toInt();
      g = g + ((255 - g) * tint).toInt();
      b = b + ((255 - b) * tint).toInt();
    } else {
      r = (r * (1 + tint)).toInt();
      g = (g * (1 + tint)).toInt();
      b = (b * (1 + tint)).toInt();
    }

    // Ensure values are within bounds
    r = r.clamp(0, 255);
    g = g.clamp(0, 255);
    b = b.clamp(0, 255);

    // Convert back to hex and return
    return "#${r.toRadixString(16).padLeft(2, '0')}${g.toRadixString(16).padLeft(2, '0')}${b.toRadixString(16).padLeft(2, '0')}";
  }
}