part of excel;

/// Styling class for cells
// ignore: must_be_immutable
class CellStyle extends Equatable {
  String? _fontColorHex, _backgroundColorHex, _fontFamily;
  HorizontalAlign _horizontalAlign = HorizontalAlign.Left;
  VerticalAlign _verticalAlign = VerticalAlign.Bottom;
  TextWrapping? _textWrapping;
  bool _bold = false, _italic = false;
  Underline _underline = Underline.None;
  int? _fontSize;
  int _rotation = 0;

  CellStyle({
    String? fontColorHex = "FF000000",
    String? backgroundColorHex = "none",
    int? fontSize,
    String? fontFamily,
    HorizontalAlign horizontalAlign = HorizontalAlign.Left,
    VerticalAlign verticalAlign = VerticalAlign.Bottom,
    TextWrapping? textWrapping,
    bool bold = false,
    Underline underline = Underline.None,
    bool italic = false,
    int rotation = 0,
  }) {
    this._textWrapping = textWrapping;

    this._bold = bold;

    this.fontSize = fontSize;

    this._italic = italic;

    this.fontFamily = fontFamily;

    this._rotation = rotation;

    if (fontColorHex != null) {
      this._fontColorHex = _isColorAppropriate(fontColorHex);
    } else {
      this._fontColorHex = "FF000000";
    }

    if (backgroundColorHex != null) {
      this._backgroundColorHex = _isColorAppropriate(backgroundColorHex);
    } else {
      this._backgroundColorHex = "none";
    }

    this._verticalAlign = verticalAlign;

    this._horizontalAlign = horizontalAlign;
  }

  ///
  ///Get Font Color
  ///
  String? get fontColor {
    return this._fontColorHex;
  }

  ///
  ///Set Font Color
  ///
  set fontColor(String? fontColorHex) {
    if (fontColorHex != null) {
      this._fontColorHex = _isColorAppropriate(fontColorHex);
    } else {
      this._fontColorHex = "FF000000";
    }
  }

  ///
  ///Get Background Color
  ///
  String? get backgroundColor {
    return this._backgroundColorHex;
  }

  ///
  ///Set Background Color
  ///
  set backgroundColor(String? backgroundColorHex) {
    if (backgroundColorHex != null) {
      this._backgroundColorHex = _isColorAppropriate(backgroundColorHex);
    } else {
      this._backgroundColorHex = "none";
    }
  }

  ///
  ///Get Horizontal Alignment
  ///
  HorizontalAlign? get horizontalAlignment {
    return this._horizontalAlign;
  }

  ///
  ///Set Horizontal Alignment
  ///
  set horizontalAlignment(HorizontalAlign? horizontalAlign) {
    this._horizontalAlign = horizontalAlign ?? HorizontalAlign.Left;
  }

  ///
  ///Get Vertical Alignment
  ///
  VerticalAlign get verticalAlignment {
    return this._verticalAlign;
  }

  ///
  ///Set Vertical Alignment
  ///
  set verticalAlignment(VerticalAlign _) {
    this._verticalAlign = _;
  }

  ///
  ///`Get Wrapping`
  ///
  TextWrapping? get wrap {
    return this._textWrapping;
  }

  ///
  ///`Set Wrapping`
  ///
  set wrap(TextWrapping? _) {
    this._textWrapping = _;
  }

  ///
  ///`Get FontFamily`
  ///
  String? get fontFamily {
    return this._fontFamily;
  }

  ///
  ///`Set FontFamily`
  ///
  set fontFamily(String? _) {
    this._fontFamily = _;
  }

  ///
  ///Get Font Size
  ///
  int? get fontSize {
    return this._fontSize;
  }

  ///
  ///Set Font Size
  ///
  set fontSize(int? _) {
    this._fontSize = _;
  }

  ///
  ///Get Rotation
  ///
  int get rotation {
    return this._rotation;
  }

  ///
  ///Rotation varies from [90 to -90]
  ///
  set rotation(int _) {
    if (_ > 90 || _ < -90) {
      _ = 0;
    }
    if (_ < 0) {
      /// The value is from 0 to -90 so now make it absolute and add it to 90
      ///
      /// -(_) + 90
      _ = -(_) + 90;
    }
    this._rotation = _;
  }

  ///
  ///Get `Underline`
  ///
  Underline get underline {
    return this._underline;
  }

  ///
  ///set `Underline`
  ///
  set underline(Underline _) {
    this._underline = _;
  }

  ///
  ///Get `Bold`
  ///
  bool get isBold {
    return this._bold;
  }

  ///
  ///Set `Bold`
  ///
  set isBold(bool _) {
    this._bold = _;
  }

  ///
  ///Get `Italic`
  ///
  bool get isItalic {
    return this._italic;
  }

  ///
  ///Set `Italic`
  ///
  set isItalic(bool _) {
    this._italic = _;
  }

  @override
  List<Object?> get props => [
        _bold,
        _rotation,
        _italic,
        _fontSize,
        _fontFamily,
        _textWrapping,
        _fontColorHex,
        _verticalAlign,
        _horizontalAlign,
        _backgroundColorHex,
      ];
}
