part of excel;

///Self correct the spanning of rows and columns by checking their cross-sectional relationship between if exists.
_selfCorrectSpanMap(Excel _excel) {
  _excel._mergeChangeLook.forEach((key) {
    if (_isContain(_excel._sheetMap['$key']) &&
        _excel._sheetMap['$key']!._spanList.isNotEmpty) {
      List<_Span?> spanList =
          List<_Span?>.from(_excel._sheetMap['$key']!._spanList);

      for (int i = 0; i < spanList.length; i++) {
        if (spanList[i] != null) {
          _Span checkerPos = spanList[i]!;
          int startRow = checkerPos.rowSpanStart,
              startColumn = checkerPos.columnSpanStart,
              endRow = checkerPos.rowSpanEnd,
              endColumn = checkerPos.columnSpanEnd;

          for (int j = i + 1; j < spanList.length; j++) {
            if (spanList[j] != null) {
              _Span spanObj = spanList[j]!;

              List gotMap = _isLocationChangeRequired(
                  startColumn, startRow, endColumn, endRow, spanObj);
              bool changeValue = gotMap[0] as bool;
              List<int> gotPosition = gotMap[1] as List<int>;

              if (changeValue) {
                startColumn = gotPosition[0];
                startRow = gotPosition[1];
                endColumn = gotPosition[2];
                endRow = gotPosition[3];
                spanList[j] = null;
              } else {
                List gotMap2 = _isLocationChangeRequired(
                    spanObj.columnSpanStart,
                    spanObj.rowSpanStart,
                    spanObj.columnSpanEnd,
                    spanObj.rowSpanEnd,
                    checkerPos);
                bool changeValue2 = gotMap2[0] as bool;
                List<int> gotPosition2 = gotMap2[1] as List<int>;

                if (changeValue2) {
                  startColumn = gotPosition2[0];
                  startRow = gotPosition2[1];
                  endColumn = gotPosition2[2];
                  endRow = gotPosition2[3];
                  spanList[j] = null;
                }
              }
            }
          }
          _Span spanObj1 = _Span();
          spanObj1._start = [startRow, startColumn];
          spanObj1._end = [endRow, endColumn];
          spanList[i] = spanObj1;
        }
      }
      _excel._sheetMap['$key']!._spanList = List<_Span?>.from(spanList);
      _excel._sheetMap['$key']!._cleanUpSpanMap();
    }
  });

  _excel._mergeChangeLook.forEach((key) {
    if (_isContain(_excel._sheetMap['$key'])) {
      _excel._sheetMap['$key']!.spannedItems;
    }
  });
}
