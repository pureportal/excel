part of excel;

/// Position of Cell - Indexes
// ignore: must_be_immutable
class CellIndex extends Equatable {
  ///```
  ///CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: 0 ); // A1
  ///CellIndex.indexByColumnRow(columnIndex: 0, rowIndex: 1 ); // A2
  ///
  ///```
  CellIndex.indexByColumnRow(
      {required int columnIndex, required int rowIndex}) {
    this._columnIndex = columnIndex;
    this._rowIndex = rowIndex;
  }

  ///```
  /// CellIndex.indexByColumnRow('A1'); // columnIndex: 0, rowIndex: 0
  /// CellIndex.indexByColumnRow('A2'); // columnIndex: 0, rowIndex: 1
  ///```
  static CellIndex indexByString(String cellIndex) {
    var li = _cellCoordsFromCellId(cellIndex);
    return CellIndex.indexByColumnRow(rowIndex: li[0], columnIndex: li[1]);
  }

  /// `Expensive function`, it might slow your program `Avoid Usage`
  String get cellId {
    return getCellId(this.columnIndex, this.rowIndex);
  }

  int _rowIndex = 0;

  /// gets rowIndex of the cell
  int get rowIndex {
    return this._rowIndex;
  }

  int _columnIndex = 0;

  /// gets columnIndex of the cell
  int get columnIndex {
    return this._columnIndex;
  }

  @override
  List<Object> get props => [_columnIndex, _rowIndex];
}
