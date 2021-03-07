part of excel;

// ignore: must_be_immutable
class Data extends Equatable {
  CellStyle? _cellStyle;
  dynamic? _value;
  CellType _cellType = CellType.String;
  bool _isFormula = false;
  late int _rowIndex, _colIndex;
  late Sheet _sheet;
  late String _sheetName;

  ///
  ///It will clone the object by changing the `this` reference of previous DataObject and putting `new this` reference, with copying the values too
  ///
  Data._clone(Sheet sheet, Data dataObject)
      : this._(
          sheet,
          dataObject._rowIndex,
          dataObject.colIndex,
          value_: dataObject._value,
          cellStyle_: dataObject._cellStyle,
          isFormula_: dataObject._isFormula,
          cellType_: dataObject._cellType,
        );

  ///
  ///Initializes the new `Data Object`
  ///
  Data._(
    Sheet sheet,
    int row,
    int col, {
    dynamic? value_,
    CellStyle? cellStyle_,
    bool isFormula_ = false,
    CellType cellType_ = CellType.String,
  }) {
    this._sheet = sheet;
    this._value = value_;
    this._cellStyle = cellStyle_;
    this._isFormula = isFormula_;
    this._cellType = cellType_;
    this._sheetName = sheet.sheetName;
    this._rowIndex = row;
    this._colIndex = col;
  }

  /// returns the newData object when called from Sheet Class
  static Data newData(Sheet sheet, int row, int col) {
    return Data._(sheet, row, col);
  }

  /// returns the cell type
  CellType get cellType {
    return this._cellType;
  }

  /// returns true is the cellType is CellType.Formula
  bool get isFormula {
    return this._isFormula;
  }

  /// returns the row Index
  int get rowIndex {
    return this._rowIndex;
  }

  /// returns the column Index
  int get colIndex {
    return this._colIndex;
  }

  /// returns the sheet-name
  String get sheetName {
    return this._sheetName;
  }

  /// returns the string based `cellId` as `A1`, `A2` or `Z5`
  String get cellId {
    return getCellId(this._colIndex, this._rowIndex);
  }

  set value(dynamic _value) {
    _sheet.updateCell(
        CellIndex.indexByColumnRow(
            columnIndex: this._colIndex, rowIndex: this._rowIndex),
        _value);
  }

  /// returns the value stored in this cell;
  ///
  /// It will return `null` if no value is stored in this cell.
  dynamic get value {
    return _value;
  }

  /// sets the user defined CellStyle in this current cell
  set style(CellStyle _style) {
    this._cellStyle = _style;
  }

  /// returns the user-defined CellStyle
  ///
  /// if `no` cellStyle is set then it returns `null`
  CellStyle? get cellStyle {
    return this._cellStyle;
  }

  set cellStyle(CellStyle? updateCellStyle) {
    _sheet._excel._colorChanges = true;
    this._cellStyle = updateCellStyle;
  }

  @override
  List<Object?> get props => throw UnimplementedError();
}
