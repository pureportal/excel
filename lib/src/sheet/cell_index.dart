part of excel;

class CellIndex {
  CellIndex._({required int col, required int row}) {
    this._columnIndex = col;
    this._rowIndex = row;
  }

  ///```
  ///CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: 0 ); // A1
  ///CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: 1 ); // A2
  ///
  ///```
  static CellIndex indexByColumnRow(
      {required int columnIndex, required int rowIndex}) {
    return CellIndex._(col: columnIndex, row: rowIndex);
  }

  ///```
  /// CellIndex.indexByColumnRow('A1'); // columnIndex: 0, rowIndex: 0
  /// CellIndex.indexByColumnRow('A2'); // columnIndex: 0, rowIndex: 1
  ///```
  static CellIndex indexByString(String cellIndex) {
    var li = _cellCoordsFromCellId(cellIndex);
    return CellIndex._(row: li[0], col: li[1]);
  }

  /// `Pretty Expensive function`, it might slow your program `Avoid Usage`
  String get cellId {
    return getCellId(this.columnIndex, this.rowIndex);
  }

  int _rowIndex = 0;

  int get rowIndex {
    return this._rowIndex;
  }

  int _columnIndex = 0;

  int get columnIndex {
    return this._columnIndex;
  }
}
