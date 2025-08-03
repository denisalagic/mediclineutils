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
        } else if (index == 64) {
          var clrScheme = colorSchemes.firstWhereOrNull((clrSch) => clrSch.id == themeId);
          if (clrScheme != null) {
            // Try system color (used for accessibility or themes)
            if (clrScheme.sysClrLast.isNotEmpty) {
              return "#${clrScheme.sysClrLast.padLeft(6, '0')}";
            }
            // Try RGB override
            if (clrScheme.srgbClr.isNotEmpty) {
              return "#${clrScheme.srgbClr.padLeft(6, '0')}";
            }
          }
          return "#FFFFFF"; // fallback
        }
      } catch (_) {
        return "#FFFFFF";
      }
    }

    // Handle edge case: no index, just RGB theme
    if (themeId.isNotEmpty) {
      var clrScheme = colorSchemes.firstWhereOrNull((clrSch) => clrSch.id == themeId);
      if (clrScheme != null) {
        if (clrScheme.sysClrLast.isNotEmpty) {
          return "#${clrScheme.sysClrLast.padLeft(6, '0')}";
        }
        if (clrScheme.srgbClr.isNotEmpty) {
          return "#${clrScheme.srgbClr.padLeft(6, '0')}";
        }
      }
    }

    return null;
  }
}