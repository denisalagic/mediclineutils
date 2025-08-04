import 'package:collection/collection.dart';
import 'dart:math' as math;

import '../data/indexed_color.dart';
import '../models/ss_color_schemes.dart';

class ColorUtil {
  static String resolveExcelColor(
      String bgClrIndex,
      String themeId,
      List<SSColorSchemes> colorSchemes, {
      String? tint,
      String? rgbValue,
      String? autoColor,
  }) {
    // Handle direct RGB values first
    if (rgbValue != null && rgbValue.isNotEmpty) {
      String hex = rgbValue.replaceAll("#", "").toUpperCase();
      if (hex.length == 6 || hex.length == 8) {
        return "#${hex.substring(hex.length >= 6 ? hex.length - 6 : 0)}";
      }
    }

    // Handle auto colors (usually black for text, white for background)
    if (autoColor != null && autoColor.toLowerCase() == "1") {
      return "#000000"; // Default to black for auto colors
    }

    // Handle indexed colors
    if (bgClrIndex.isNotEmpty) {
      try {
        int index = int.parse(bgClrIndex);

        // Standard indexed colors (0-63)
        if (index >= 0 && index < 64) {
          String color = IndexedColor().colors[index];
          // Only apply tint if present and not '0'
          if (tint != null && tint.isNotEmpty && tint != "0") {
            return _applyTint(color, tint);
          } else {
            return color;
          }
        }

        // Theme colors (index 64+ or when themeId is provided)
        if (index >= 64 || themeId.isNotEmpty) {
          String resolvedColor = _resolveThemeColor(themeId, colorSchemes);
          // Only apply tint if present and not '0'
          if (tint != null && tint.isNotEmpty && tint != "0") {
            return _applyTint(resolvedColor, tint);
          } else {
            return resolvedColor;
          }
        }
      } catch (e) {
        print("Error parsing color index '$bgClrIndex': $e");
      }
    }

    // Handle theme colors without index
    if (themeId.isNotEmpty) {
      String resolvedColor = _resolveThemeColor(themeId, colorSchemes);
      // Only apply tint if present and not '0'
      if (tint != null && tint.isNotEmpty && tint != "0") {
        return _applyTint(resolvedColor, tint);
      } else {
        return resolvedColor;
      }
    }

    // Default fallback - return transparent/white
    return "#FFFFFF";
  }

  static String _resolveThemeColor(String themeId, List<SSColorSchemes> colorSchemes) {
    if (themeId.isEmpty) return "#FFFFFF";

    // Clean up theme ID - remove prefixes and normalize
    String cleanThemeId = themeId.toLowerCase()
        .replaceAll("a:", "")
        .replaceAll("theme", "")
        .trim();

    // Try exact match first
    var clrScheme = colorSchemes.firstWhereOrNull((clrSch) =>
        clrSch.id.toLowerCase() == cleanThemeId);

    // If no exact match, try partial matches
    if (clrScheme == null) {
      clrScheme = colorSchemes.firstWhereOrNull((clrSch) =>
          clrSch.id.toLowerCase().contains(cleanThemeId) ||
          cleanThemeId.contains(clrSch.id.toLowerCase()));
    }

    // Try matching by numeric ID
    if (clrScheme == null) {
      try {
        int numericId = int.parse(cleanThemeId);
        // Map common theme IDs to theme names
        String themeName = _getThemeNameFromId(numericId);
        if (themeName.isNotEmpty) {
          clrScheme = colorSchemes.firstWhereOrNull((clrSch) =>
              clrSch.id.toLowerCase() == themeName);
        }
      } catch (_) {
        // Not a numeric ID, continue
      }
    }

    if (clrScheme != null) {
      // Prioritize srgbClr over sysClrLast
      String hex = "";
      if (clrScheme.srgbClr.isNotEmpty) {
        hex = clrScheme.srgbClr;
      } else if (clrScheme.sysClrLast.isNotEmpty) {
        hex = clrScheme.sysClrLast;
      }

      if (hex.isNotEmpty) {
        hex = hex.replaceAll("#", "").toUpperCase();
        // Handle ARGB format (8 chars) by taking last 6 chars (RGB)
        if (hex.length > 6) {
          hex = hex.substring(hex.length - 6);
        }
        if (hex.length == 6) {
          return "#$hex";
        }
      }
    }

    // Fallback to default theme colors if no match found
    return _getDefaultThemeColor(cleanThemeId);
  }

  static String _getThemeNameFromId(int id) {
    switch (id) {
      case 0:
      case 1:
        return "lt1";
      case 2:
      case 3:
        return "dk1";
      case 4:
        return "accent1";
      case 5:
        return "accent2";
      case 6:
        return "accent3";
      case 7:
        return "accent4";
      case 8:
        return "accent5";
      case 9:
        return "accent6";
      case 10:
        return "hlink";
      case 11:
        return "folhlink";
      default:
        return "";
    }
  }

  static String _getDefaultThemeColor(String themeId) {
    switch (themeId.toLowerCase()) {
      case "0":
      case "1":
      case "lt1":
        return "#FFFFFF";
      case "2":
      case "3":
      case "dk1":
        return "#000000";
      case "4":
      case "accent1":
        return "#4F81BD";
      case "5":
      case "accent2":
        return "#F79646";
      case "6":
      case "accent3":
        return "#9BBB59";
      case "7":
      case "accent4":
        return "#8064A2";
      case "8":
      case "accent5":
        return "#4BACC6";
      case "9":
      case "accent6":
        return "#F24C4C";
      case "10":
      case "hlink":
        return "#0000FF";
      case "11":
      case "folhlink":
        return "#800080";
      default:
        return "#FFFFFF";
    }
  }

  static String _applyTint(String color, String? tint) {
    if (tint == null || tint.isEmpty || tint == "0") {
      return color;
    }

    try {
      double tintValue = double.parse(tint);
      if (tintValue == 0) return color;

      // Remove # if present
      String hex = color.replaceAll("#", "");
      if (hex.length != 6) return color;

      // Parse RGB values
      int r = int.parse(hex.substring(0, 2), radix: 16);
      int g = int.parse(hex.substring(2, 4), radix: 16);
      int b = int.parse(hex.substring(4, 6), radix: 16);

      // Apply tint
      if (tintValue > 0) {
        // Tint towards white
        r = (r + (255 - r) * tintValue).round();
        g = (g + (255 - g) * tintValue).round();
        b = (b + (255 - b) * tintValue).round();
      } else {
        // Shade towards black
        double shade = 1 + tintValue;
        r = (r * shade).round();
        g = (g * shade).round();
        b = (b * shade).round();
      }

      // Clamp values
      r = math.max(0, math.min(255, r));
      g = math.max(0, math.min(255, g));
      b = math.max(0, math.min(255, b));

      // Convert back to hex
      String newHex = r.toRadixString(16).padLeft(2, '0') +
                     g.toRadixString(16).padLeft(2, '0') +
                     b.toRadixString(16).padLeft(2, '0');

      return "#${newHex.toUpperCase()}";
    } catch (e) {
      print("Error applying tint '$tint' to color '$color': $e");
      return color;
    }
  }
}