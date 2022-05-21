package word

// enum WdMailSystem
var WdMailSystem = struct {
	WdNoMailSystem int32
	WdMAPI int32
	WdPowerTalk int32
	WdMAPIandPowerTalk int32
}{
	WdNoMailSystem: 0,
	WdMAPI: 1,
	WdPowerTalk: 2,
	WdMAPIandPowerTalk: 3,
}

// enum WdTemplateType
var WdTemplateType = struct {
	WdNormalTemplate int32
	WdGlobalTemplate int32
	WdAttachedTemplate int32
}{
	WdNormalTemplate: 0,
	WdGlobalTemplate: 1,
	WdAttachedTemplate: 2,
}

// enum WdContinue
var WdContinue = struct {
	WdContinueDisabled int32
	WdResetList int32
	WdContinueList int32
}{
	WdContinueDisabled: 0,
	WdResetList: 1,
	WdContinueList: 2,
}

// enum WdIMEMode
var WdIMEMode = struct {
	WdIMEModeNoControl int32
	WdIMEModeOn int32
	WdIMEModeOff int32
	WdIMEModeHiragana int32
	WdIMEModeKatakana int32
	WdIMEModeKatakanaHalf int32
	WdIMEModeAlphaFull int32
	WdIMEModeAlpha int32
	WdIMEModeHangulFull int32
	WdIMEModeHangul int32
}{
	WdIMEModeNoControl: 0,
	WdIMEModeOn: 1,
	WdIMEModeOff: 2,
	WdIMEModeHiragana: 4,
	WdIMEModeKatakana: 5,
	WdIMEModeKatakanaHalf: 6,
	WdIMEModeAlphaFull: 7,
	WdIMEModeAlpha: 8,
	WdIMEModeHangulFull: 9,
	WdIMEModeHangul: 10,
}

// enum WdBaselineAlignment
var WdBaselineAlignment = struct {
	WdBaselineAlignTop int32
	WdBaselineAlignCenter int32
	WdBaselineAlignBaseline int32
	WdBaselineAlignFarEast50 int32
	WdBaselineAlignAuto int32
}{
	WdBaselineAlignTop: 0,
	WdBaselineAlignCenter: 1,
	WdBaselineAlignBaseline: 2,
	WdBaselineAlignFarEast50: 3,
	WdBaselineAlignAuto: 4,
}

// enum WdIndexFilter
var WdIndexFilter = struct {
	WdIndexFilterNone int32
	WdIndexFilterAiueo int32
	WdIndexFilterAkasatana int32
	WdIndexFilterChosung int32
	WdIndexFilterLow int32
	WdIndexFilterMedium int32
	WdIndexFilterFull int32
}{
	WdIndexFilterNone: 0,
	WdIndexFilterAiueo: 1,
	WdIndexFilterAkasatana: 2,
	WdIndexFilterChosung: 3,
	WdIndexFilterLow: 4,
	WdIndexFilterMedium: 5,
	WdIndexFilterFull: 6,
}

// enum WdIndexSortBy
var WdIndexSortBy = struct {
	WdIndexSortByStroke int32
	WdIndexSortBySyllable int32
}{
	WdIndexSortByStroke: 0,
	WdIndexSortBySyllable: 1,
}

// enum WdJustificationMode
var WdJustificationMode = struct {
	WdJustificationModeExpand int32
	WdJustificationModeCompress int32
	WdJustificationModeCompressKana int32
}{
	WdJustificationModeExpand: 0,
	WdJustificationModeCompress: 1,
	WdJustificationModeCompressKana: 2,
}

// enum WdFarEastLineBreakLevel
var WdFarEastLineBreakLevel = struct {
	WdFarEastLineBreakLevelNormal int32
	WdFarEastLineBreakLevelStrict int32
	WdFarEastLineBreakLevelCustom int32
}{
	WdFarEastLineBreakLevelNormal: 0,
	WdFarEastLineBreakLevelStrict: 1,
	WdFarEastLineBreakLevelCustom: 2,
}

// enum WdMultipleWordConversionsMode
var WdMultipleWordConversionsMode = struct {
	WdHangulToHanja int32
	WdHanjaToHangul int32
}{
	WdHangulToHanja: 0,
	WdHanjaToHangul: 1,
}

// enum WdColorIndex
var WdColorIndex = struct {
	WdAuto int32
	WdBlack int32
	WdBlue int32
	WdTurquoise int32
	WdBrightGreen int32
	WdPink int32
	WdRed int32
	WdYellow int32
	WdWhite int32
	WdDarkBlue int32
	WdTeal int32
	WdGreen int32
	WdViolet int32
	WdDarkRed int32
	WdDarkYellow int32
	WdGray50 int32
	WdGray25 int32
	WdByAuthor int32
	WdNoHighlight int32
}{
	WdAuto: 0,
	WdBlack: 1,
	WdBlue: 2,
	WdTurquoise: 3,
	WdBrightGreen: 4,
	WdPink: 5,
	WdRed: 6,
	WdYellow: 7,
	WdWhite: 8,
	WdDarkBlue: 9,
	WdTeal: 10,
	WdGreen: 11,
	WdViolet: 12,
	WdDarkRed: 13,
	WdDarkYellow: 14,
	WdGray50: 15,
	WdGray25: 16,
	WdByAuthor: -1,
	WdNoHighlight: 0,
}

// enum WdTextureIndex
var WdTextureIndex = struct {
	WdTextureNone int32
	WdTexture2Pt5Percent int32
	WdTexture5Percent int32
	WdTexture7Pt5Percent int32
	WdTexture10Percent int32
	WdTexture12Pt5Percent int32
	WdTexture15Percent int32
	WdTexture17Pt5Percent int32
	WdTexture20Percent int32
	WdTexture22Pt5Percent int32
	WdTexture25Percent int32
	WdTexture27Pt5Percent int32
	WdTexture30Percent int32
	WdTexture32Pt5Percent int32
	WdTexture35Percent int32
	WdTexture37Pt5Percent int32
	WdTexture40Percent int32
	WdTexture42Pt5Percent int32
	WdTexture45Percent int32
	WdTexture47Pt5Percent int32
	WdTexture50Percent int32
	WdTexture52Pt5Percent int32
	WdTexture55Percent int32
	WdTexture57Pt5Percent int32
	WdTexture60Percent int32
	WdTexture62Pt5Percent int32
	WdTexture65Percent int32
	WdTexture67Pt5Percent int32
	WdTexture70Percent int32
	WdTexture72Pt5Percent int32
	WdTexture75Percent int32
	WdTexture77Pt5Percent int32
	WdTexture80Percent int32
	WdTexture82Pt5Percent int32
	WdTexture85Percent int32
	WdTexture87Pt5Percent int32
	WdTexture90Percent int32
	WdTexture92Pt5Percent int32
	WdTexture95Percent int32
	WdTexture97Pt5Percent int32
	WdTextureSolid int32
	WdTextureDarkHorizontal int32
	WdTextureDarkVertical int32
	WdTextureDarkDiagonalDown int32
	WdTextureDarkDiagonalUp int32
	WdTextureDarkCross int32
	WdTextureDarkDiagonalCross int32
	WdTextureHorizontal int32
	WdTextureVertical int32
	WdTextureDiagonalDown int32
	WdTextureDiagonalUp int32
	WdTextureCross int32
	WdTextureDiagonalCross int32
}{
	WdTextureNone: 0,
	WdTexture2Pt5Percent: 25,
	WdTexture5Percent: 50,
	WdTexture7Pt5Percent: 75,
	WdTexture10Percent: 100,
	WdTexture12Pt5Percent: 125,
	WdTexture15Percent: 150,
	WdTexture17Pt5Percent: 175,
	WdTexture20Percent: 200,
	WdTexture22Pt5Percent: 225,
	WdTexture25Percent: 250,
	WdTexture27Pt5Percent: 275,
	WdTexture30Percent: 300,
	WdTexture32Pt5Percent: 325,
	WdTexture35Percent: 350,
	WdTexture37Pt5Percent: 375,
	WdTexture40Percent: 400,
	WdTexture42Pt5Percent: 425,
	WdTexture45Percent: 450,
	WdTexture47Pt5Percent: 475,
	WdTexture50Percent: 500,
	WdTexture52Pt5Percent: 525,
	WdTexture55Percent: 550,
	WdTexture57Pt5Percent: 575,
	WdTexture60Percent: 600,
	WdTexture62Pt5Percent: 625,
	WdTexture65Percent: 650,
	WdTexture67Pt5Percent: 675,
	WdTexture70Percent: 700,
	WdTexture72Pt5Percent: 725,
	WdTexture75Percent: 750,
	WdTexture77Pt5Percent: 775,
	WdTexture80Percent: 800,
	WdTexture82Pt5Percent: 825,
	WdTexture85Percent: 850,
	WdTexture87Pt5Percent: 875,
	WdTexture90Percent: 900,
	WdTexture92Pt5Percent: 925,
	WdTexture95Percent: 950,
	WdTexture97Pt5Percent: 975,
	WdTextureSolid: 1000,
	WdTextureDarkHorizontal: -1,
	WdTextureDarkVertical: -2,
	WdTextureDarkDiagonalDown: -3,
	WdTextureDarkDiagonalUp: -4,
	WdTextureDarkCross: -5,
	WdTextureDarkDiagonalCross: -6,
	WdTextureHorizontal: -7,
	WdTextureVertical: -8,
	WdTextureDiagonalDown: -9,
	WdTextureDiagonalUp: -10,
	WdTextureCross: -11,
	WdTextureDiagonalCross: -12,
}

// enum WdUnderline
var WdUnderline = struct {
	WdUnderlineNone int32
	WdUnderlineSingle int32
	WdUnderlineWords int32
	WdUnderlineDouble int32
	WdUnderlineDotted int32
	WdUnderlineThick int32
	WdUnderlineDash int32
	WdUnderlineDotDash int32
	WdUnderlineDotDotDash int32
	WdUnderlineWavy int32
	WdUnderlineWavyHeavy int32
	WdUnderlineDottedHeavy int32
	WdUnderlineDashHeavy int32
	WdUnderlineDotDashHeavy int32
	WdUnderlineDotDotDashHeavy int32
	WdUnderlineDashLong int32
	WdUnderlineDashLongHeavy int32
	WdUnderlineWavyDouble int32
}{
	WdUnderlineNone: 0,
	WdUnderlineSingle: 1,
	WdUnderlineWords: 2,
	WdUnderlineDouble: 3,
	WdUnderlineDotted: 4,
	WdUnderlineThick: 6,
	WdUnderlineDash: 7,
	WdUnderlineDotDash: 9,
	WdUnderlineDotDotDash: 10,
	WdUnderlineWavy: 11,
	WdUnderlineWavyHeavy: 27,
	WdUnderlineDottedHeavy: 20,
	WdUnderlineDashHeavy: 23,
	WdUnderlineDotDashHeavy: 25,
	WdUnderlineDotDotDashHeavy: 26,
	WdUnderlineDashLong: 39,
	WdUnderlineDashLongHeavy: 55,
	WdUnderlineWavyDouble: 43,
}

// enum WdEmphasisMark
var WdEmphasisMark = struct {
	WdEmphasisMarkNone int32
	WdEmphasisMarkOverSolidCircle int32
	WdEmphasisMarkOverComma int32
	WdEmphasisMarkOverWhiteCircle int32
	WdEmphasisMarkUnderSolidCircle int32
}{
	WdEmphasisMarkNone: 0,
	WdEmphasisMarkOverSolidCircle: 1,
	WdEmphasisMarkOverComma: 2,
	WdEmphasisMarkOverWhiteCircle: 3,
	WdEmphasisMarkUnderSolidCircle: 4,
}

// enum WdInternationalIndex
var WdInternationalIndex = struct {
	WdListSeparator int32
	WdDecimalSeparator int32
	WdThousandsSeparator int32
	WdCurrencyCode int32
	Wd24HourClock int32
	WdInternationalAM int32
	WdInternationalPM int32
	WdTimeSeparator int32
	WdDateSeparator int32
	WdProductLanguageID int32
}{
	WdListSeparator: 17,
	WdDecimalSeparator: 18,
	WdThousandsSeparator: 19,
	WdCurrencyCode: 20,
	Wd24HourClock: 21,
	WdInternationalAM: 22,
	WdInternationalPM: 23,
	WdTimeSeparator: 24,
	WdDateSeparator: 25,
	WdProductLanguageID: 26,
}

// enum WdAutoMacros
var WdAutoMacros = struct {
	WdAutoExec int32
	WdAutoNew int32
	WdAutoOpen int32
	WdAutoClose int32
	WdAutoExit int32
	WdAutoSync int32
}{
	WdAutoExec: 0,
	WdAutoNew: 1,
	WdAutoOpen: 2,
	WdAutoClose: 3,
	WdAutoExit: 4,
	WdAutoSync: 5,
}

// enum WdCaptionPosition
var WdCaptionPosition = struct {
	WdCaptionPositionAbove int32
	WdCaptionPositionBelow int32
}{
	WdCaptionPositionAbove: 0,
	WdCaptionPositionBelow: 1,
}

// enum WdCountry
var WdCountry = struct {
	WdUS int32
	WdCanada int32
	WdLatinAmerica int32
	WdNetherlands int32
	WdFrance int32
	WdSpain int32
	WdItaly int32
	WdUK int32
	WdDenmark int32
	WdSweden int32
	WdNorway int32
	WdGermany int32
	WdPeru int32
	WdMexico int32
	WdArgentina int32
	WdBrazil int32
	WdChile int32
	WdVenezuela int32
	WdJapan int32
	WdTaiwan int32
	WdChina int32
	WdKorea int32
	WdFinland int32
	WdIceland int32
}{
	WdUS: 1,
	WdCanada: 2,
	WdLatinAmerica: 3,
	WdNetherlands: 31,
	WdFrance: 33,
	WdSpain: 34,
	WdItaly: 39,
	WdUK: 44,
	WdDenmark: 45,
	WdSweden: 46,
	WdNorway: 47,
	WdGermany: 49,
	WdPeru: 51,
	WdMexico: 52,
	WdArgentina: 54,
	WdBrazil: 55,
	WdChile: 56,
	WdVenezuela: 58,
	WdJapan: 81,
	WdTaiwan: 886,
	WdChina: 86,
	WdKorea: 82,
	WdFinland: 358,
	WdIceland: 354,
}

// enum WdHeadingSeparator
var WdHeadingSeparator = struct {
	WdHeadingSeparatorNone int32
	WdHeadingSeparatorBlankLine int32
	WdHeadingSeparatorLetter int32
	WdHeadingSeparatorLetterLow int32
	WdHeadingSeparatorLetterFull int32
}{
	WdHeadingSeparatorNone: 0,
	WdHeadingSeparatorBlankLine: 1,
	WdHeadingSeparatorLetter: 2,
	WdHeadingSeparatorLetterLow: 3,
	WdHeadingSeparatorLetterFull: 4,
}

// enum WdSeparatorType
var WdSeparatorType = struct {
	WdSeparatorHyphen int32
	WdSeparatorPeriod int32
	WdSeparatorColon int32
	WdSeparatorEmDash int32
	WdSeparatorEnDash int32
}{
	WdSeparatorHyphen: 0,
	WdSeparatorPeriod: 1,
	WdSeparatorColon: 2,
	WdSeparatorEmDash: 3,
	WdSeparatorEnDash: 4,
}

// enum WdPageNumberAlignment
var WdPageNumberAlignment = struct {
	WdAlignPageNumberLeft int32
	WdAlignPageNumberCenter int32
	WdAlignPageNumberRight int32
	WdAlignPageNumberInside int32
	WdAlignPageNumberOutside int32
}{
	WdAlignPageNumberLeft: 0,
	WdAlignPageNumberCenter: 1,
	WdAlignPageNumberRight: 2,
	WdAlignPageNumberInside: 3,
	WdAlignPageNumberOutside: 4,
}

// enum WdBorderType
var WdBorderType = struct {
	WdBorderTop int32
	WdBorderLeft int32
	WdBorderBottom int32
	WdBorderRight int32
	WdBorderHorizontal int32
	WdBorderVertical int32
	WdBorderDiagonalDown int32
	WdBorderDiagonalUp int32
}{
	WdBorderTop: -1,
	WdBorderLeft: -2,
	WdBorderBottom: -3,
	WdBorderRight: -4,
	WdBorderHorizontal: -5,
	WdBorderVertical: -6,
	WdBorderDiagonalDown: -7,
	WdBorderDiagonalUp: -8,
}

// enum WdBorderTypeHID
var WdBorderTypeHID = struct {
	Emptyenum int32
}{
	Emptyenum: 0,
}

// enum WdFramePosition
var WdFramePosition = struct {
	WdFrameTop int32
	WdFrameLeft int32
	WdFrameBottom int32
	WdFrameRight int32
	WdFrameCenter int32
	WdFrameInside int32
	WdFrameOutside int32
}{
	WdFrameTop: -999999,
	WdFrameLeft: -999998,
	WdFrameBottom: -999997,
	WdFrameRight: -999996,
	WdFrameCenter: -999995,
	WdFrameInside: -999994,
	WdFrameOutside: -999993,
}

// enum WdAnimation
var WdAnimation = struct {
	WdAnimationNone int32
	WdAnimationLasVegasLights int32
	WdAnimationBlinkingBackground int32
	WdAnimationSparkleText int32
	WdAnimationMarchingBlackAnts int32
	WdAnimationMarchingRedAnts int32
	WdAnimationShimmer int32
}{
	WdAnimationNone: 0,
	WdAnimationLasVegasLights: 1,
	WdAnimationBlinkingBackground: 2,
	WdAnimationSparkleText: 3,
	WdAnimationMarchingBlackAnts: 4,
	WdAnimationMarchingRedAnts: 5,
	WdAnimationShimmer: 6,
}

// enum WdCharacterCase
var WdCharacterCase = struct {
	WdNextCase int32
	WdLowerCase int32
	WdUpperCase int32
	WdTitleWord int32
	WdTitleSentence int32
	WdToggleCase int32
	WdHalfWidth int32
	WdFullWidth int32
	WdKatakana int32
	WdHiragana int32
}{
	WdNextCase: -1,
	WdLowerCase: 0,
	WdUpperCase: 1,
	WdTitleWord: 2,
	WdTitleSentence: 4,
	WdToggleCase: 5,
	WdHalfWidth: 6,
	WdFullWidth: 7,
	WdKatakana: 8,
	WdHiragana: 9,
}

// enum WdCharacterCaseHID
var WdCharacterCaseHID = struct {
	Emptyenum int32
}{
	Emptyenum: 0,
}

// enum WdSummaryMode
var WdSummaryMode = struct {
	WdSummaryModeHighlight int32
	WdSummaryModeHideAllButSummary int32
	WdSummaryModeInsert int32
	WdSummaryModeCreateNew int32
}{
	WdSummaryModeHighlight: 0,
	WdSummaryModeHideAllButSummary: 1,
	WdSummaryModeInsert: 2,
	WdSummaryModeCreateNew: 3,
}

// enum WdSummaryLength
var WdSummaryLength = struct {
	Wd10Sentences int32
	Wd20Sentences int32
	Wd100Words int32
	Wd500Words int32
	Wd10Percent int32
	Wd25Percent int32
	Wd50Percent int32
	Wd75Percent int32
}{
	Wd10Sentences: -2,
	Wd20Sentences: -3,
	Wd100Words: -4,
	Wd500Words: -5,
	Wd10Percent: -6,
	Wd25Percent: -7,
	Wd50Percent: -8,
	Wd75Percent: -9,
}

// enum WdStyleType
var WdStyleType = struct {
	WdStyleTypeParagraph int32
	WdStyleTypeCharacter int32
	WdStyleTypeTable int32
	WdStyleTypeList int32
	WdStyleTypeParagraphOnly int32
	WdStyleTypeLinked int32
}{
	WdStyleTypeParagraph: 1,
	WdStyleTypeCharacter: 2,
	WdStyleTypeTable: 3,
	WdStyleTypeList: 4,
	WdStyleTypeParagraphOnly: 5,
	WdStyleTypeLinked: 6,
}

// enum WdUnits
var WdUnits = struct {
	WdCharacter int32
	WdWord int32
	WdSentence int32
	WdParagraph int32
	WdLine int32
	WdStory int32
	WdScreen int32
	WdSection int32
	WdColumn int32
	WdRow int32
	WdWindow int32
	WdCell int32
	WdCharacterFormatting int32
	WdParagraphFormatting int32
	WdTable int32
	WdItem int32
}{
	WdCharacter: 1,
	WdWord: 2,
	WdSentence: 3,
	WdParagraph: 4,
	WdLine: 5,
	WdStory: 6,
	WdScreen: 7,
	WdSection: 8,
	WdColumn: 9,
	WdRow: 10,
	WdWindow: 11,
	WdCell: 12,
	WdCharacterFormatting: 13,
	WdParagraphFormatting: 14,
	WdTable: 15,
	WdItem: 16,
}

// enum WdGoToItem
var WdGoToItem = struct {
	WdGoToBookmark int32
	WdGoToSection int32
	WdGoToPage int32
	WdGoToTable int32
	WdGoToLine int32
	WdGoToFootnote int32
	WdGoToEndnote int32
	WdGoToComment int32
	WdGoToField int32
	WdGoToGraphic int32
	WdGoToObject int32
	WdGoToEquation int32
	WdGoToHeading int32
	WdGoToPercent int32
	WdGoToSpellingError int32
	WdGoToGrammaticalError int32
	WdGoToProofreadingError int32
}{
	WdGoToBookmark: -1,
	WdGoToSection: 0,
	WdGoToPage: 1,
	WdGoToTable: 2,
	WdGoToLine: 3,
	WdGoToFootnote: 4,
	WdGoToEndnote: 5,
	WdGoToComment: 6,
	WdGoToField: 7,
	WdGoToGraphic: 8,
	WdGoToObject: 9,
	WdGoToEquation: 10,
	WdGoToHeading: 11,
	WdGoToPercent: 12,
	WdGoToSpellingError: 13,
	WdGoToGrammaticalError: 14,
	WdGoToProofreadingError: 15,
}

// enum WdGoToDirection
var WdGoToDirection = struct {
	WdGoToFirst int32
	WdGoToLast int32
	WdGoToNext int32
	WdGoToRelative int32
	WdGoToPrevious int32
	WdGoToAbsolute int32
}{
	WdGoToFirst: 1,
	WdGoToLast: -1,
	WdGoToNext: 2,
	WdGoToRelative: 2,
	WdGoToPrevious: 3,
	WdGoToAbsolute: 1,
}

// enum WdCollapseDirection
var WdCollapseDirection = struct {
	WdCollapseStart int32
	WdCollapseEnd int32
}{
	WdCollapseStart: 1,
	WdCollapseEnd: 0,
}

// enum WdRowHeightRule
var WdRowHeightRule = struct {
	WdRowHeightAuto int32
	WdRowHeightAtLeast int32
	WdRowHeightExactly int32
}{
	WdRowHeightAuto: 0,
	WdRowHeightAtLeast: 1,
	WdRowHeightExactly: 2,
}

// enum WdFrameSizeRule
var WdFrameSizeRule = struct {
	WdFrameAuto int32
	WdFrameAtLeast int32
	WdFrameExact int32
}{
	WdFrameAuto: 0,
	WdFrameAtLeast: 1,
	WdFrameExact: 2,
}

// enum WdInsertCells
var WdInsertCells = struct {
	WdInsertCellsShiftRight int32
	WdInsertCellsShiftDown int32
	WdInsertCellsEntireRow int32
	WdInsertCellsEntireColumn int32
}{
	WdInsertCellsShiftRight: 0,
	WdInsertCellsShiftDown: 1,
	WdInsertCellsEntireRow: 2,
	WdInsertCellsEntireColumn: 3,
}

// enum WdDeleteCells
var WdDeleteCells = struct {
	WdDeleteCellsShiftLeft int32
	WdDeleteCellsShiftUp int32
	WdDeleteCellsEntireRow int32
	WdDeleteCellsEntireColumn int32
}{
	WdDeleteCellsShiftLeft: 0,
	WdDeleteCellsShiftUp: 1,
	WdDeleteCellsEntireRow: 2,
	WdDeleteCellsEntireColumn: 3,
}

// enum WdListApplyTo
var WdListApplyTo = struct {
	WdListApplyToWholeList int32
	WdListApplyToThisPointForward int32
	WdListApplyToSelection int32
}{
	WdListApplyToWholeList: 0,
	WdListApplyToThisPointForward: 1,
	WdListApplyToSelection: 2,
}

// enum WdAlertLevel
var WdAlertLevel = struct {
	WdAlertsNone int32
	WdAlertsMessageBox int32
	WdAlertsAll int32
}{
	WdAlertsNone: 0,
	WdAlertsMessageBox: -2,
	WdAlertsAll: -1,
}

// enum WdCursorType
var WdCursorType = struct {
	WdCursorWait int32
	WdCursorIBeam int32
	WdCursorNormal int32
	WdCursorNorthwestArrow int32
}{
	WdCursorWait: 0,
	WdCursorIBeam: 1,
	WdCursorNormal: 2,
	WdCursorNorthwestArrow: 3,
}

// enum WdEnableCancelKey
var WdEnableCancelKey = struct {
	WdCancelDisabled int32
	WdCancelInterrupt int32
}{
	WdCancelDisabled: 0,
	WdCancelInterrupt: 1,
}

// enum WdRulerStyle
var WdRulerStyle = struct {
	WdAdjustNone int32
	WdAdjustProportional int32
	WdAdjustFirstColumn int32
	WdAdjustSameWidth int32
}{
	WdAdjustNone: 0,
	WdAdjustProportional: 1,
	WdAdjustFirstColumn: 2,
	WdAdjustSameWidth: 3,
}

// enum WdParagraphAlignment
var WdParagraphAlignment = struct {
	WdAlignParagraphLeft int32
	WdAlignParagraphCenter int32
	WdAlignParagraphRight int32
	WdAlignParagraphJustify int32
	WdAlignParagraphDistribute int32
	WdAlignParagraphJustifyMed int32
	WdAlignParagraphJustifyHi int32
	WdAlignParagraphJustifyLow int32
	WdAlignParagraphThaiJustify int32
}{
	WdAlignParagraphLeft: 0,
	WdAlignParagraphCenter: 1,
	WdAlignParagraphRight: 2,
	WdAlignParagraphJustify: 3,
	WdAlignParagraphDistribute: 4,
	WdAlignParagraphJustifyMed: 5,
	WdAlignParagraphJustifyHi: 7,
	WdAlignParagraphJustifyLow: 8,
	WdAlignParagraphThaiJustify: 9,
}

// enum WdParagraphAlignmentHID
var WdParagraphAlignmentHID = struct {
	Emptyenum int32
}{
	Emptyenum: 0,
}

// enum WdListLevelAlignment
var WdListLevelAlignment = struct {
	WdListLevelAlignLeft int32
	WdListLevelAlignCenter int32
	WdListLevelAlignRight int32
}{
	WdListLevelAlignLeft: 0,
	WdListLevelAlignCenter: 1,
	WdListLevelAlignRight: 2,
}

// enum WdRowAlignment
var WdRowAlignment = struct {
	WdAlignRowLeft int32
	WdAlignRowCenter int32
	WdAlignRowRight int32
}{
	WdAlignRowLeft: 0,
	WdAlignRowCenter: 1,
	WdAlignRowRight: 2,
}

// enum WdTabAlignment
var WdTabAlignment = struct {
	WdAlignTabLeft int32
	WdAlignTabCenter int32
	WdAlignTabRight int32
	WdAlignTabDecimal int32
	WdAlignTabBar int32
	WdAlignTabList int32
}{
	WdAlignTabLeft: 0,
	WdAlignTabCenter: 1,
	WdAlignTabRight: 2,
	WdAlignTabDecimal: 3,
	WdAlignTabBar: 4,
	WdAlignTabList: 6,
}

// enum WdVerticalAlignment
var WdVerticalAlignment = struct {
	WdAlignVerticalTop int32
	WdAlignVerticalCenter int32
	WdAlignVerticalJustify int32
	WdAlignVerticalBottom int32
}{
	WdAlignVerticalTop: 0,
	WdAlignVerticalCenter: 1,
	WdAlignVerticalJustify: 2,
	WdAlignVerticalBottom: 3,
}

// enum WdCellVerticalAlignment
var WdCellVerticalAlignment = struct {
	WdCellAlignVerticalTop int32
	WdCellAlignVerticalCenter int32
	WdCellAlignVerticalBottom int32
}{
	WdCellAlignVerticalTop: 0,
	WdCellAlignVerticalCenter: 1,
	WdCellAlignVerticalBottom: 3,
}

// enum WdTrailingCharacter
var WdTrailingCharacter = struct {
	WdTrailingTab int32
	WdTrailingSpace int32
	WdTrailingNone int32
}{
	WdTrailingTab: 0,
	WdTrailingSpace: 1,
	WdTrailingNone: 2,
}

// enum WdListGalleryType
var WdListGalleryType = struct {
	WdBulletGallery int32
	WdNumberGallery int32
	WdOutlineNumberGallery int32
}{
	WdBulletGallery: 1,
	WdNumberGallery: 2,
	WdOutlineNumberGallery: 3,
}

// enum WdListNumberStyle
var WdListNumberStyle = struct {
	WdListNumberStyleArabic int32
	WdListNumberStyleUppercaseRoman int32
	WdListNumberStyleLowercaseRoman int32
	WdListNumberStyleUppercaseLetter int32
	WdListNumberStyleLowercaseLetter int32
	WdListNumberStyleOrdinal int32
	WdListNumberStyleCardinalText int32
	WdListNumberStyleOrdinalText int32
	WdListNumberStyleKanji int32
	WdListNumberStyleKanjiDigit int32
	WdListNumberStyleAiueoHalfWidth int32
	WdListNumberStyleIrohaHalfWidth int32
	WdListNumberStyleArabicFullWidth int32
	WdListNumberStyleKanjiTraditional int32
	WdListNumberStyleKanjiTraditional2 int32
	WdListNumberStyleNumberInCircle int32
	WdListNumberStyleAiueo int32
	WdListNumberStyleIroha int32
	WdListNumberStyleArabicLZ int32
	WdListNumberStyleBullet int32
	WdListNumberStyleGanada int32
	WdListNumberStyleChosung int32
	WdListNumberStyleGBNum1 int32
	WdListNumberStyleGBNum2 int32
	WdListNumberStyleGBNum3 int32
	WdListNumberStyleGBNum4 int32
	WdListNumberStyleZodiac1 int32
	WdListNumberStyleZodiac2 int32
	WdListNumberStyleZodiac3 int32
	WdListNumberStyleTradChinNum1 int32
	WdListNumberStyleTradChinNum2 int32
	WdListNumberStyleTradChinNum3 int32
	WdListNumberStyleTradChinNum4 int32
	WdListNumberStyleSimpChinNum1 int32
	WdListNumberStyleSimpChinNum2 int32
	WdListNumberStyleSimpChinNum3 int32
	WdListNumberStyleSimpChinNum4 int32
	WdListNumberStyleHanjaRead int32
	WdListNumberStyleHanjaReadDigit int32
	WdListNumberStyleHangul int32
	WdListNumberStyleHanja int32
	WdListNumberStyleHebrew1 int32
	WdListNumberStyleArabic1 int32
	WdListNumberStyleHebrew2 int32
	WdListNumberStyleArabic2 int32
	WdListNumberStyleHindiLetter1 int32
	WdListNumberStyleHindiLetter2 int32
	WdListNumberStyleHindiArabic int32
	WdListNumberStyleHindiCardinalText int32
	WdListNumberStyleThaiLetter int32
	WdListNumberStyleThaiArabic int32
	WdListNumberStyleThaiCardinalText int32
	WdListNumberStyleVietCardinalText int32
	WdListNumberStyleLowercaseRussian int32
	WdListNumberStyleUppercaseRussian int32
	WdListNumberStyleLowercaseGreek int32
	WdListNumberStyleUppercaseGreek int32
	WdListNumberStyleArabicLZ2 int32
	WdListNumberStyleArabicLZ3 int32
	WdListNumberStyleArabicLZ4 int32
	WdListNumberStyleLowercaseTurkish int32
	WdListNumberStyleUppercaseTurkish int32
	WdListNumberStyleLowercaseBulgarian int32
	WdListNumberStyleUppercaseBulgarian int32
	WdListNumberStylePictureBullet int32
	WdListNumberStyleLegal int32
	WdListNumberStyleLegalLZ int32
	WdListNumberStyleNone int32
}{
	WdListNumberStyleArabic: 0,
	WdListNumberStyleUppercaseRoman: 1,
	WdListNumberStyleLowercaseRoman: 2,
	WdListNumberStyleUppercaseLetter: 3,
	WdListNumberStyleLowercaseLetter: 4,
	WdListNumberStyleOrdinal: 5,
	WdListNumberStyleCardinalText: 6,
	WdListNumberStyleOrdinalText: 7,
	WdListNumberStyleKanji: 10,
	WdListNumberStyleKanjiDigit: 11,
	WdListNumberStyleAiueoHalfWidth: 12,
	WdListNumberStyleIrohaHalfWidth: 13,
	WdListNumberStyleArabicFullWidth: 14,
	WdListNumberStyleKanjiTraditional: 16,
	WdListNumberStyleKanjiTraditional2: 17,
	WdListNumberStyleNumberInCircle: 18,
	WdListNumberStyleAiueo: 20,
	WdListNumberStyleIroha: 21,
	WdListNumberStyleArabicLZ: 22,
	WdListNumberStyleBullet: 23,
	WdListNumberStyleGanada: 24,
	WdListNumberStyleChosung: 25,
	WdListNumberStyleGBNum1: 26,
	WdListNumberStyleGBNum2: 27,
	WdListNumberStyleGBNum3: 28,
	WdListNumberStyleGBNum4: 29,
	WdListNumberStyleZodiac1: 30,
	WdListNumberStyleZodiac2: 31,
	WdListNumberStyleZodiac3: 32,
	WdListNumberStyleTradChinNum1: 33,
	WdListNumberStyleTradChinNum2: 34,
	WdListNumberStyleTradChinNum3: 35,
	WdListNumberStyleTradChinNum4: 36,
	WdListNumberStyleSimpChinNum1: 37,
	WdListNumberStyleSimpChinNum2: 38,
	WdListNumberStyleSimpChinNum3: 39,
	WdListNumberStyleSimpChinNum4: 40,
	WdListNumberStyleHanjaRead: 41,
	WdListNumberStyleHanjaReadDigit: 42,
	WdListNumberStyleHangul: 43,
	WdListNumberStyleHanja: 44,
	WdListNumberStyleHebrew1: 45,
	WdListNumberStyleArabic1: 46,
	WdListNumberStyleHebrew2: 47,
	WdListNumberStyleArabic2: 48,
	WdListNumberStyleHindiLetter1: 49,
	WdListNumberStyleHindiLetter2: 50,
	WdListNumberStyleHindiArabic: 51,
	WdListNumberStyleHindiCardinalText: 52,
	WdListNumberStyleThaiLetter: 53,
	WdListNumberStyleThaiArabic: 54,
	WdListNumberStyleThaiCardinalText: 55,
	WdListNumberStyleVietCardinalText: 56,
	WdListNumberStyleLowercaseRussian: 58,
	WdListNumberStyleUppercaseRussian: 59,
	WdListNumberStyleLowercaseGreek: 60,
	WdListNumberStyleUppercaseGreek: 61,
	WdListNumberStyleArabicLZ2: 62,
	WdListNumberStyleArabicLZ3: 63,
	WdListNumberStyleArabicLZ4: 64,
	WdListNumberStyleLowercaseTurkish: 65,
	WdListNumberStyleUppercaseTurkish: 66,
	WdListNumberStyleLowercaseBulgarian: 67,
	WdListNumberStyleUppercaseBulgarian: 68,
	WdListNumberStylePictureBullet: 249,
	WdListNumberStyleLegal: 253,
	WdListNumberStyleLegalLZ: 254,
	WdListNumberStyleNone: 255,
}

// enum WdListNumberStyleHID
var WdListNumberStyleHID = struct {
	Emptyenum int32
}{
	Emptyenum: 0,
}

// enum WdNoteNumberStyle
var WdNoteNumberStyle = struct {
	WdNoteNumberStyleArabic int32
	WdNoteNumberStyleUppercaseRoman int32
	WdNoteNumberStyleLowercaseRoman int32
	WdNoteNumberStyleUppercaseLetter int32
	WdNoteNumberStyleLowercaseLetter int32
	WdNoteNumberStyleSymbol int32
	WdNoteNumberStyleArabicFullWidth int32
	WdNoteNumberStyleKanji int32
	WdNoteNumberStyleKanjiDigit int32
	WdNoteNumberStyleKanjiTraditional int32
	WdNoteNumberStyleNumberInCircle int32
	WdNoteNumberStyleHanjaRead int32
	WdNoteNumberStyleHanjaReadDigit int32
	WdNoteNumberStyleTradChinNum1 int32
	WdNoteNumberStyleTradChinNum2 int32
	WdNoteNumberStyleSimpChinNum1 int32
	WdNoteNumberStyleSimpChinNum2 int32
	WdNoteNumberStyleHebrewLetter1 int32
	WdNoteNumberStyleArabicLetter1 int32
	WdNoteNumberStyleHebrewLetter2 int32
	WdNoteNumberStyleArabicLetter2 int32
	WdNoteNumberStyleHindiLetter1 int32
	WdNoteNumberStyleHindiLetter2 int32
	WdNoteNumberStyleHindiArabic int32
	WdNoteNumberStyleHindiCardinalText int32
	WdNoteNumberStyleThaiLetter int32
	WdNoteNumberStyleThaiArabic int32
	WdNoteNumberStyleThaiCardinalText int32
	WdNoteNumberStyleVietCardinalText int32
}{
	WdNoteNumberStyleArabic: 0,
	WdNoteNumberStyleUppercaseRoman: 1,
	WdNoteNumberStyleLowercaseRoman: 2,
	WdNoteNumberStyleUppercaseLetter: 3,
	WdNoteNumberStyleLowercaseLetter: 4,
	WdNoteNumberStyleSymbol: 9,
	WdNoteNumberStyleArabicFullWidth: 14,
	WdNoteNumberStyleKanji: 10,
	WdNoteNumberStyleKanjiDigit: 11,
	WdNoteNumberStyleKanjiTraditional: 16,
	WdNoteNumberStyleNumberInCircle: 18,
	WdNoteNumberStyleHanjaRead: 41,
	WdNoteNumberStyleHanjaReadDigit: 42,
	WdNoteNumberStyleTradChinNum1: 33,
	WdNoteNumberStyleTradChinNum2: 34,
	WdNoteNumberStyleSimpChinNum1: 37,
	WdNoteNumberStyleSimpChinNum2: 38,
	WdNoteNumberStyleHebrewLetter1: 45,
	WdNoteNumberStyleArabicLetter1: 46,
	WdNoteNumberStyleHebrewLetter2: 47,
	WdNoteNumberStyleArabicLetter2: 48,
	WdNoteNumberStyleHindiLetter1: 49,
	WdNoteNumberStyleHindiLetter2: 50,
	WdNoteNumberStyleHindiArabic: 51,
	WdNoteNumberStyleHindiCardinalText: 52,
	WdNoteNumberStyleThaiLetter: 53,
	WdNoteNumberStyleThaiArabic: 54,
	WdNoteNumberStyleThaiCardinalText: 55,
	WdNoteNumberStyleVietCardinalText: 56,
}

// enum WdNoteNumberStyleHID
var WdNoteNumberStyleHID = struct {
	Emptyenum int32
}{
	Emptyenum: 0,
}

// enum WdCaptionNumberStyle
var WdCaptionNumberStyle = struct {
	WdCaptionNumberStyleArabic int32
	WdCaptionNumberStyleUppercaseRoman int32
	WdCaptionNumberStyleLowercaseRoman int32
	WdCaptionNumberStyleUppercaseLetter int32
	WdCaptionNumberStyleLowercaseLetter int32
	WdCaptionNumberStyleArabicFullWidth int32
	WdCaptionNumberStyleKanji int32
	WdCaptionNumberStyleKanjiDigit int32
	WdCaptionNumberStyleKanjiTraditional int32
	WdCaptionNumberStyleNumberInCircle int32
	WdCaptionNumberStyleGanada int32
	WdCaptionNumberStyleChosung int32
	WdCaptionNumberStyleZodiac1 int32
	WdCaptionNumberStyleZodiac2 int32
	WdCaptionNumberStyleHanjaRead int32
	WdCaptionNumberStyleHanjaReadDigit int32
	WdCaptionNumberStyleTradChinNum2 int32
	WdCaptionNumberStyleTradChinNum3 int32
	WdCaptionNumberStyleSimpChinNum2 int32
	WdCaptionNumberStyleSimpChinNum3 int32
	WdCaptionNumberStyleHebrewLetter1 int32
	WdCaptionNumberStyleArabicLetter1 int32
	WdCaptionNumberStyleHebrewLetter2 int32
	WdCaptionNumberStyleArabicLetter2 int32
	WdCaptionNumberStyleHindiLetter1 int32
	WdCaptionNumberStyleHindiLetter2 int32
	WdCaptionNumberStyleHindiArabic int32
	WdCaptionNumberStyleHindiCardinalText int32
	WdCaptionNumberStyleThaiLetter int32
	WdCaptionNumberStyleThaiArabic int32
	WdCaptionNumberStyleThaiCardinalText int32
	WdCaptionNumberStyleVietCardinalText int32
}{
	WdCaptionNumberStyleArabic: 0,
	WdCaptionNumberStyleUppercaseRoman: 1,
	WdCaptionNumberStyleLowercaseRoman: 2,
	WdCaptionNumberStyleUppercaseLetter: 3,
	WdCaptionNumberStyleLowercaseLetter: 4,
	WdCaptionNumberStyleArabicFullWidth: 14,
	WdCaptionNumberStyleKanji: 10,
	WdCaptionNumberStyleKanjiDigit: 11,
	WdCaptionNumberStyleKanjiTraditional: 16,
	WdCaptionNumberStyleNumberInCircle: 18,
	WdCaptionNumberStyleGanada: 24,
	WdCaptionNumberStyleChosung: 25,
	WdCaptionNumberStyleZodiac1: 30,
	WdCaptionNumberStyleZodiac2: 31,
	WdCaptionNumberStyleHanjaRead: 41,
	WdCaptionNumberStyleHanjaReadDigit: 42,
	WdCaptionNumberStyleTradChinNum2: 34,
	WdCaptionNumberStyleTradChinNum3: 35,
	WdCaptionNumberStyleSimpChinNum2: 38,
	WdCaptionNumberStyleSimpChinNum3: 39,
	WdCaptionNumberStyleHebrewLetter1: 45,
	WdCaptionNumberStyleArabicLetter1: 46,
	WdCaptionNumberStyleHebrewLetter2: 47,
	WdCaptionNumberStyleArabicLetter2: 48,
	WdCaptionNumberStyleHindiLetter1: 49,
	WdCaptionNumberStyleHindiLetter2: 50,
	WdCaptionNumberStyleHindiArabic: 51,
	WdCaptionNumberStyleHindiCardinalText: 52,
	WdCaptionNumberStyleThaiLetter: 53,
	WdCaptionNumberStyleThaiArabic: 54,
	WdCaptionNumberStyleThaiCardinalText: 55,
	WdCaptionNumberStyleVietCardinalText: 56,
}

// enum WdCaptionNumberStyleHID
var WdCaptionNumberStyleHID = struct {
	Emptyenum int32
}{
	Emptyenum: 0,
}

// enum WdPageNumberStyle
var WdPageNumberStyle = struct {
	WdPageNumberStyleArabic int32
	WdPageNumberStyleUppercaseRoman int32
	WdPageNumberStyleLowercaseRoman int32
	WdPageNumberStyleUppercaseLetter int32
	WdPageNumberStyleLowercaseLetter int32
	WdPageNumberStyleArabicFullWidth int32
	WdPageNumberStyleKanji int32
	WdPageNumberStyleKanjiDigit int32
	WdPageNumberStyleKanjiTraditional int32
	WdPageNumberStyleNumberInCircle int32
	WdPageNumberStyleHanjaRead int32
	WdPageNumberStyleHanjaReadDigit int32
	WdPageNumberStyleTradChinNum1 int32
	WdPageNumberStyleTradChinNum2 int32
	WdPageNumberStyleSimpChinNum1 int32
	WdPageNumberStyleSimpChinNum2 int32
	WdPageNumberStyleHebrewLetter1 int32
	WdPageNumberStyleArabicLetter1 int32
	WdPageNumberStyleHebrewLetter2 int32
	WdPageNumberStyleArabicLetter2 int32
	WdPageNumberStyleHindiLetter1 int32
	WdPageNumberStyleHindiLetter2 int32
	WdPageNumberStyleHindiArabic int32
	WdPageNumberStyleHindiCardinalText int32
	WdPageNumberStyleThaiLetter int32
	WdPageNumberStyleThaiArabic int32
	WdPageNumberStyleThaiCardinalText int32
	WdPageNumberStyleVietCardinalText int32
	WdPageNumberStyleNumberInDash int32
}{
	WdPageNumberStyleArabic: 0,
	WdPageNumberStyleUppercaseRoman: 1,
	WdPageNumberStyleLowercaseRoman: 2,
	WdPageNumberStyleUppercaseLetter: 3,
	WdPageNumberStyleLowercaseLetter: 4,
	WdPageNumberStyleArabicFullWidth: 14,
	WdPageNumberStyleKanji: 10,
	WdPageNumberStyleKanjiDigit: 11,
	WdPageNumberStyleKanjiTraditional: 16,
	WdPageNumberStyleNumberInCircle: 18,
	WdPageNumberStyleHanjaRead: 41,
	WdPageNumberStyleHanjaReadDigit: 42,
	WdPageNumberStyleTradChinNum1: 33,
	WdPageNumberStyleTradChinNum2: 34,
	WdPageNumberStyleSimpChinNum1: 37,
	WdPageNumberStyleSimpChinNum2: 38,
	WdPageNumberStyleHebrewLetter1: 45,
	WdPageNumberStyleArabicLetter1: 46,
	WdPageNumberStyleHebrewLetter2: 47,
	WdPageNumberStyleArabicLetter2: 48,
	WdPageNumberStyleHindiLetter1: 49,
	WdPageNumberStyleHindiLetter2: 50,
	WdPageNumberStyleHindiArabic: 51,
	WdPageNumberStyleHindiCardinalText: 52,
	WdPageNumberStyleThaiLetter: 53,
	WdPageNumberStyleThaiArabic: 54,
	WdPageNumberStyleThaiCardinalText: 55,
	WdPageNumberStyleVietCardinalText: 56,
	WdPageNumberStyleNumberInDash: 57,
}

// enum WdPageNumberStyleHID
var WdPageNumberStyleHID = struct {
	Emptyenum int32
}{
	Emptyenum: 0,
}

// enum WdStatistic
var WdStatistic = struct {
	WdStatisticWords int32
	WdStatisticLines int32
	WdStatisticPages int32
	WdStatisticCharacters int32
	WdStatisticParagraphs int32
	WdStatisticCharactersWithSpaces int32
	WdStatisticFarEastCharacters int32
}{
	WdStatisticWords: 0,
	WdStatisticLines: 1,
	WdStatisticPages: 2,
	WdStatisticCharacters: 3,
	WdStatisticParagraphs: 4,
	WdStatisticCharactersWithSpaces: 5,
	WdStatisticFarEastCharacters: 6,
}

// enum WdStatisticHID
var WdStatisticHID = struct {
	Emptyenum int32
}{
	Emptyenum: 0,
}

// enum WdBuiltInProperty
var WdBuiltInProperty = struct {
	WdPropertyTitle int32
	WdPropertySubject int32
	WdPropertyAuthor int32
	WdPropertyKeywords int32
	WdPropertyComments int32
	WdPropertyTemplate int32
	WdPropertyLastAuthor int32
	WdPropertyRevision int32
	WdPropertyAppName int32
	WdPropertyTimeLastPrinted int32
	WdPropertyTimeCreated int32
	WdPropertyTimeLastSaved int32
	WdPropertyVBATotalEdit int32
	WdPropertyPages int32
	WdPropertyWords int32
	WdPropertyCharacters int32
	WdPropertySecurity int32
	WdPropertyCategory int32
	WdPropertyFormat int32
	WdPropertyManager int32
	WdPropertyCompany int32
	WdPropertyBytes int32
	WdPropertyLines int32
	WdPropertyParas int32
	WdPropertySlides int32
	WdPropertyNotes int32
	WdPropertyHiddenSlides int32
	WdPropertyMMClips int32
	WdPropertyHyperlinkBase int32
	WdPropertyCharsWSpaces int32
}{
	WdPropertyTitle: 1,
	WdPropertySubject: 2,
	WdPropertyAuthor: 3,
	WdPropertyKeywords: 4,
	WdPropertyComments: 5,
	WdPropertyTemplate: 6,
	WdPropertyLastAuthor: 7,
	WdPropertyRevision: 8,
	WdPropertyAppName: 9,
	WdPropertyTimeLastPrinted: 10,
	WdPropertyTimeCreated: 11,
	WdPropertyTimeLastSaved: 12,
	WdPropertyVBATotalEdit: 13,
	WdPropertyPages: 14,
	WdPropertyWords: 15,
	WdPropertyCharacters: 16,
	WdPropertySecurity: 17,
	WdPropertyCategory: 18,
	WdPropertyFormat: 19,
	WdPropertyManager: 20,
	WdPropertyCompany: 21,
	WdPropertyBytes: 22,
	WdPropertyLines: 23,
	WdPropertyParas: 24,
	WdPropertySlides: 25,
	WdPropertyNotes: 26,
	WdPropertyHiddenSlides: 27,
	WdPropertyMMClips: 28,
	WdPropertyHyperlinkBase: 29,
	WdPropertyCharsWSpaces: 30,
}

// enum WdLineSpacing
var WdLineSpacing = struct {
	WdLineSpaceSingle int32
	WdLineSpace1pt5 int32
	WdLineSpaceDouble int32
	WdLineSpaceAtLeast int32
	WdLineSpaceExactly int32
	WdLineSpaceMultiple int32
}{
	WdLineSpaceSingle: 0,
	WdLineSpace1pt5: 1,
	WdLineSpaceDouble: 2,
	WdLineSpaceAtLeast: 3,
	WdLineSpaceExactly: 4,
	WdLineSpaceMultiple: 5,
}

// enum WdNumberType
var WdNumberType = struct {
	WdNumberParagraph int32
	WdNumberListNum int32
	WdNumberAllNumbers int32
}{
	WdNumberParagraph: 1,
	WdNumberListNum: 2,
	WdNumberAllNumbers: 3,
}

// enum WdListType
var WdListType = struct {
	WdListNoNumbering int32
	WdListListNumOnly int32
	WdListBullet int32
	WdListSimpleNumbering int32
	WdListOutlineNumbering int32
	WdListMixedNumbering int32
	WdListPictureBullet int32
}{
	WdListNoNumbering: 0,
	WdListListNumOnly: 1,
	WdListBullet: 2,
	WdListSimpleNumbering: 3,
	WdListOutlineNumbering: 4,
	WdListMixedNumbering: 5,
	WdListPictureBullet: 6,
}

// enum WdStoryType
var WdStoryType = struct {
	WdMainTextStory int32
	WdFootnotesStory int32
	WdEndnotesStory int32
	WdCommentsStory int32
	WdTextFrameStory int32
	WdEvenPagesHeaderStory int32
	WdPrimaryHeaderStory int32
	WdEvenPagesFooterStory int32
	WdPrimaryFooterStory int32
	WdFirstPageHeaderStory int32
	WdFirstPageFooterStory int32
	WdFootnoteSeparatorStory int32
	WdFootnoteContinuationSeparatorStory int32
	WdFootnoteContinuationNoticeStory int32
	WdEndnoteSeparatorStory int32
	WdEndnoteContinuationSeparatorStory int32
	WdEndnoteContinuationNoticeStory int32
}{
	WdMainTextStory: 1,
	WdFootnotesStory: 2,
	WdEndnotesStory: 3,
	WdCommentsStory: 4,
	WdTextFrameStory: 5,
	WdEvenPagesHeaderStory: 6,
	WdPrimaryHeaderStory: 7,
	WdEvenPagesFooterStory: 8,
	WdPrimaryFooterStory: 9,
	WdFirstPageHeaderStory: 10,
	WdFirstPageFooterStory: 11,
	WdFootnoteSeparatorStory: 12,
	WdFootnoteContinuationSeparatorStory: 13,
	WdFootnoteContinuationNoticeStory: 14,
	WdEndnoteSeparatorStory: 15,
	WdEndnoteContinuationSeparatorStory: 16,
	WdEndnoteContinuationNoticeStory: 17,
}

// enum WdSaveFormat
var WdSaveFormat = struct {
	WdFormatDocument int32
	WdFormatDocument97 int32
	WdFormatTemplate int32
	WdFormatTemplate97 int32
	WdFormatText int32
	WdFormatTextLineBreaks int32
	WdFormatDOSText int32
	WdFormatDOSTextLineBreaks int32
	WdFormatRTF int32
	WdFormatUnicodeText int32
	WdFormatEncodedText int32
	WdFormatHTML int32
	WdFormatWebArchive int32
	WdFormatFilteredHTML int32
	WdFormatXML int32
	WdFormatXMLDocument int32
	WdFormatXMLDocumentMacroEnabled int32
	WdFormatXMLTemplate int32
	WdFormatXMLTemplateMacroEnabled int32
	WdFormatDocumentDefault int32
	WdFormatPDF int32
	WdFormatXPS int32
	WdFormatFlatXML int32
	WdFormatFlatXMLMacroEnabled int32
	WdFormatFlatXMLTemplate int32
	WdFormatFlatXMLTemplateMacroEnabled int32
	WdFormatOpenDocumentText int32
}{
	WdFormatDocument: 0,
	WdFormatDocument97: 0,
	WdFormatTemplate: 1,
	WdFormatTemplate97: 1,
	WdFormatText: 2,
	WdFormatTextLineBreaks: 3,
	WdFormatDOSText: 4,
	WdFormatDOSTextLineBreaks: 5,
	WdFormatRTF: 6,
	WdFormatUnicodeText: 7,
	WdFormatEncodedText: 7,
	WdFormatHTML: 8,
	WdFormatWebArchive: 9,
	WdFormatFilteredHTML: 10,
	WdFormatXML: 11,
	WdFormatXMLDocument: 12,
	WdFormatXMLDocumentMacroEnabled: 13,
	WdFormatXMLTemplate: 14,
	WdFormatXMLTemplateMacroEnabled: 15,
	WdFormatDocumentDefault: 16,
	WdFormatPDF: 17,
	WdFormatXPS: 18,
	WdFormatFlatXML: 19,
	WdFormatFlatXMLMacroEnabled: 20,
	WdFormatFlatXMLTemplate: 21,
	WdFormatFlatXMLTemplateMacroEnabled: 22,
	WdFormatOpenDocumentText: 23,
}

// enum WdOpenFormat
var WdOpenFormat = struct {
	WdOpenFormatAuto int32
	WdOpenFormatDocument int32
	WdOpenFormatTemplate int32
	WdOpenFormatRTF int32
	WdOpenFormatText int32
	WdOpenFormatUnicodeText int32
	WdOpenFormatEncodedText int32
	WdOpenFormatAllWord int32
	WdOpenFormatWebPages int32
	WdOpenFormatXML int32
	WdOpenFormatXMLDocument int32
	WdOpenFormatXMLDocumentMacroEnabled int32
	WdOpenFormatXMLTemplate int32
	WdOpenFormatXMLTemplateMacroEnabled int32
	WdOpenFormatDocument97 int32
	WdOpenFormatTemplate97 int32
	WdOpenFormatAllWordTemplates int32
	WdOpenFormatXMLDocumentSerialized int32
	WdOpenFormatXMLDocumentMacroEnabledSerialized int32
	WdOpenFormatXMLTemplateSerialized int32
	WdOpenFormatXMLTemplateMacroEnabledSerialized int32
	WdOpenFormatOpenDocumentText int32
}{
	WdOpenFormatAuto: 0,
	WdOpenFormatDocument: 1,
	WdOpenFormatTemplate: 2,
	WdOpenFormatRTF: 3,
	WdOpenFormatText: 4,
	WdOpenFormatUnicodeText: 5,
	WdOpenFormatEncodedText: 5,
	WdOpenFormatAllWord: 6,
	WdOpenFormatWebPages: 7,
	WdOpenFormatXML: 8,
	WdOpenFormatXMLDocument: 9,
	WdOpenFormatXMLDocumentMacroEnabled: 10,
	WdOpenFormatXMLTemplate: 11,
	WdOpenFormatXMLTemplateMacroEnabled: 12,
	WdOpenFormatDocument97: 1,
	WdOpenFormatTemplate97: 2,
	WdOpenFormatAllWordTemplates: 13,
	WdOpenFormatXMLDocumentSerialized: 14,
	WdOpenFormatXMLDocumentMacroEnabledSerialized: 15,
	WdOpenFormatXMLTemplateSerialized: 16,
	WdOpenFormatXMLTemplateMacroEnabledSerialized: 17,
	WdOpenFormatOpenDocumentText: 18,
}

// enum WdHeaderFooterIndex
var WdHeaderFooterIndex = struct {
	WdHeaderFooterPrimary int32
	WdHeaderFooterFirstPage int32
	WdHeaderFooterEvenPages int32
}{
	WdHeaderFooterPrimary: 1,
	WdHeaderFooterFirstPage: 2,
	WdHeaderFooterEvenPages: 3,
}

// enum WdTocFormat
var WdTocFormat = struct {
	WdTOCTemplate int32
	WdTOCClassic int32
	WdTOCDistinctive int32
	WdTOCFancy int32
	WdTOCModern int32
	WdTOCFormal int32
	WdTOCSimple int32
}{
	WdTOCTemplate: 0,
	WdTOCClassic: 1,
	WdTOCDistinctive: 2,
	WdTOCFancy: 3,
	WdTOCModern: 4,
	WdTOCFormal: 5,
	WdTOCSimple: 6,
}

// enum WdTofFormat
var WdTofFormat = struct {
	WdTOFTemplate int32
	WdTOFClassic int32
	WdTOFDistinctive int32
	WdTOFCentered int32
	WdTOFFormal int32
	WdTOFSimple int32
}{
	WdTOFTemplate: 0,
	WdTOFClassic: 1,
	WdTOFDistinctive: 2,
	WdTOFCentered: 3,
	WdTOFFormal: 4,
	WdTOFSimple: 5,
}

// enum WdToaFormat
var WdToaFormat = struct {
	WdTOATemplate int32
	WdTOAClassic int32
	WdTOADistinctive int32
	WdTOAFormal int32
	WdTOASimple int32
}{
	WdTOATemplate: 0,
	WdTOAClassic: 1,
	WdTOADistinctive: 2,
	WdTOAFormal: 3,
	WdTOASimple: 4,
}

// enum WdLineStyle
var WdLineStyle = struct {
	WdLineStyleNone int32
	WdLineStyleSingle int32
	WdLineStyleDot int32
	WdLineStyleDashSmallGap int32
	WdLineStyleDashLargeGap int32
	WdLineStyleDashDot int32
	WdLineStyleDashDotDot int32
	WdLineStyleDouble int32
	WdLineStyleTriple int32
	WdLineStyleThinThickSmallGap int32
	WdLineStyleThickThinSmallGap int32
	WdLineStyleThinThickThinSmallGap int32
	WdLineStyleThinThickMedGap int32
	WdLineStyleThickThinMedGap int32
	WdLineStyleThinThickThinMedGap int32
	WdLineStyleThinThickLargeGap int32
	WdLineStyleThickThinLargeGap int32
	WdLineStyleThinThickThinLargeGap int32
	WdLineStyleSingleWavy int32
	WdLineStyleDoubleWavy int32
	WdLineStyleDashDotStroked int32
	WdLineStyleEmboss3D int32
	WdLineStyleEngrave3D int32
	WdLineStyleOutset int32
	WdLineStyleInset int32
}{
	WdLineStyleNone: 0,
	WdLineStyleSingle: 1,
	WdLineStyleDot: 2,
	WdLineStyleDashSmallGap: 3,
	WdLineStyleDashLargeGap: 4,
	WdLineStyleDashDot: 5,
	WdLineStyleDashDotDot: 6,
	WdLineStyleDouble: 7,
	WdLineStyleTriple: 8,
	WdLineStyleThinThickSmallGap: 9,
	WdLineStyleThickThinSmallGap: 10,
	WdLineStyleThinThickThinSmallGap: 11,
	WdLineStyleThinThickMedGap: 12,
	WdLineStyleThickThinMedGap: 13,
	WdLineStyleThinThickThinMedGap: 14,
	WdLineStyleThinThickLargeGap: 15,
	WdLineStyleThickThinLargeGap: 16,
	WdLineStyleThinThickThinLargeGap: 17,
	WdLineStyleSingleWavy: 18,
	WdLineStyleDoubleWavy: 19,
	WdLineStyleDashDotStroked: 20,
	WdLineStyleEmboss3D: 21,
	WdLineStyleEngrave3D: 22,
	WdLineStyleOutset: 23,
	WdLineStyleInset: 24,
}

// enum WdLineWidth
var WdLineWidth = struct {
	WdLineWidth025pt int32
	WdLineWidth050pt int32
	WdLineWidth075pt int32
	WdLineWidth100pt int32
	WdLineWidth150pt int32
	WdLineWidth225pt int32
	WdLineWidth300pt int32
	WdLineWidth450pt int32
	WdLineWidth600pt int32
}{
	WdLineWidth025pt: 2,
	WdLineWidth050pt: 4,
	WdLineWidth075pt: 6,
	WdLineWidth100pt: 8,
	WdLineWidth150pt: 12,
	WdLineWidth225pt: 18,
	WdLineWidth300pt: 24,
	WdLineWidth450pt: 36,
	WdLineWidth600pt: 48,
}

// enum WdBreakType
var WdBreakType = struct {
	WdSectionBreakNextPage int32
	WdSectionBreakContinuous int32
	WdSectionBreakEvenPage int32
	WdSectionBreakOddPage int32
	WdLineBreak int32
	WdPageBreak int32
	WdColumnBreak int32
	WdLineBreakClearLeft int32
	WdLineBreakClearRight int32
	WdTextWrappingBreak int32
}{
	WdSectionBreakNextPage: 2,
	WdSectionBreakContinuous: 3,
	WdSectionBreakEvenPage: 4,
	WdSectionBreakOddPage: 5,
	WdLineBreak: 6,
	WdPageBreak: 7,
	WdColumnBreak: 8,
	WdLineBreakClearLeft: 9,
	WdLineBreakClearRight: 10,
	WdTextWrappingBreak: 11,
}

// enum WdTabLeader
var WdTabLeader = struct {
	WdTabLeaderSpaces int32
	WdTabLeaderDots int32
	WdTabLeaderDashes int32
	WdTabLeaderLines int32
	WdTabLeaderHeavy int32
	WdTabLeaderMiddleDot int32
}{
	WdTabLeaderSpaces: 0,
	WdTabLeaderDots: 1,
	WdTabLeaderDashes: 2,
	WdTabLeaderLines: 3,
	WdTabLeaderHeavy: 4,
	WdTabLeaderMiddleDot: 5,
}

// enum WdTabLeaderHID
var WdTabLeaderHID = struct {
	Emptyenum int32
}{
	Emptyenum: 0,
}

// enum WdMeasurementUnits
var WdMeasurementUnits = struct {
	WdInches int32
	WdCentimeters int32
	WdMillimeters int32
	WdPoints int32
	WdPicas int32
}{
	WdInches: 0,
	WdCentimeters: 1,
	WdMillimeters: 2,
	WdPoints: 3,
	WdPicas: 4,
}

// enum WdMeasurementUnitsHID
var WdMeasurementUnitsHID = struct {
	Emptyenum int32
}{
	Emptyenum: 0,
}

// enum WdDropPosition
var WdDropPosition = struct {
	WdDropNone int32
	WdDropNormal int32
	WdDropMargin int32
}{
	WdDropNone: 0,
	WdDropNormal: 1,
	WdDropMargin: 2,
}

// enum WdNumberingRule
var WdNumberingRule = struct {
	WdRestartContinuous int32
	WdRestartSection int32
	WdRestartPage int32
}{
	WdRestartContinuous: 0,
	WdRestartSection: 1,
	WdRestartPage: 2,
}

// enum WdFootnoteLocation
var WdFootnoteLocation = struct {
	WdBottomOfPage int32
	WdBeneathText int32
}{
	WdBottomOfPage: 0,
	WdBeneathText: 1,
}

// enum WdEndnoteLocation
var WdEndnoteLocation = struct {
	WdEndOfSection int32
	WdEndOfDocument int32
}{
	WdEndOfSection: 0,
	WdEndOfDocument: 1,
}

// enum WdSortSeparator
var WdSortSeparator = struct {
	WdSortSeparateByTabs int32
	WdSortSeparateByCommas int32
	WdSortSeparateByDefaultTableSeparator int32
}{
	WdSortSeparateByTabs: 0,
	WdSortSeparateByCommas: 1,
	WdSortSeparateByDefaultTableSeparator: 2,
}

// enum WdTableFieldSeparator
var WdTableFieldSeparator = struct {
	WdSeparateByParagraphs int32
	WdSeparateByTabs int32
	WdSeparateByCommas int32
	WdSeparateByDefaultListSeparator int32
}{
	WdSeparateByParagraphs: 0,
	WdSeparateByTabs: 1,
	WdSeparateByCommas: 2,
	WdSeparateByDefaultListSeparator: 3,
}

// enum WdSortFieldType
var WdSortFieldType = struct {
	WdSortFieldAlphanumeric int32
	WdSortFieldNumeric int32
	WdSortFieldDate int32
	WdSortFieldSyllable int32
	WdSortFieldJapanJIS int32
	WdSortFieldStroke int32
	WdSortFieldKoreaKS int32
}{
	WdSortFieldAlphanumeric: 0,
	WdSortFieldNumeric: 1,
	WdSortFieldDate: 2,
	WdSortFieldSyllable: 3,
	WdSortFieldJapanJIS: 4,
	WdSortFieldStroke: 5,
	WdSortFieldKoreaKS: 6,
}

// enum WdSortFieldTypeHID
var WdSortFieldTypeHID = struct {
	Emptyenum int32
}{
	Emptyenum: 0,
}

// enum WdSortOrder
var WdSortOrder = struct {
	WdSortOrderAscending int32
	WdSortOrderDescending int32
}{
	WdSortOrderAscending: 0,
	WdSortOrderDescending: 1,
}

// enum WdTableFormat
var WdTableFormat = struct {
	WdTableFormatNone int32
	WdTableFormatSimple1 int32
	WdTableFormatSimple2 int32
	WdTableFormatSimple3 int32
	WdTableFormatClassic1 int32
	WdTableFormatClassic2 int32
	WdTableFormatClassic3 int32
	WdTableFormatClassic4 int32
	WdTableFormatColorful1 int32
	WdTableFormatColorful2 int32
	WdTableFormatColorful3 int32
	WdTableFormatColumns1 int32
	WdTableFormatColumns2 int32
	WdTableFormatColumns3 int32
	WdTableFormatColumns4 int32
	WdTableFormatColumns5 int32
	WdTableFormatGrid1 int32
	WdTableFormatGrid2 int32
	WdTableFormatGrid3 int32
	WdTableFormatGrid4 int32
	WdTableFormatGrid5 int32
	WdTableFormatGrid6 int32
	WdTableFormatGrid7 int32
	WdTableFormatGrid8 int32
	WdTableFormatList1 int32
	WdTableFormatList2 int32
	WdTableFormatList3 int32
	WdTableFormatList4 int32
	WdTableFormatList5 int32
	WdTableFormatList6 int32
	WdTableFormatList7 int32
	WdTableFormatList8 int32
	WdTableFormat3DEffects1 int32
	WdTableFormat3DEffects2 int32
	WdTableFormat3DEffects3 int32
	WdTableFormatContemporary int32
	WdTableFormatElegant int32
	WdTableFormatProfessional int32
	WdTableFormatSubtle1 int32
	WdTableFormatSubtle2 int32
	WdTableFormatWeb1 int32
	WdTableFormatWeb2 int32
	WdTableFormatWeb3 int32
}{
	WdTableFormatNone: 0,
	WdTableFormatSimple1: 1,
	WdTableFormatSimple2: 2,
	WdTableFormatSimple3: 3,
	WdTableFormatClassic1: 4,
	WdTableFormatClassic2: 5,
	WdTableFormatClassic3: 6,
	WdTableFormatClassic4: 7,
	WdTableFormatColorful1: 8,
	WdTableFormatColorful2: 9,
	WdTableFormatColorful3: 10,
	WdTableFormatColumns1: 11,
	WdTableFormatColumns2: 12,
	WdTableFormatColumns3: 13,
	WdTableFormatColumns4: 14,
	WdTableFormatColumns5: 15,
	WdTableFormatGrid1: 16,
	WdTableFormatGrid2: 17,
	WdTableFormatGrid3: 18,
	WdTableFormatGrid4: 19,
	WdTableFormatGrid5: 20,
	WdTableFormatGrid6: 21,
	WdTableFormatGrid7: 22,
	WdTableFormatGrid8: 23,
	WdTableFormatList1: 24,
	WdTableFormatList2: 25,
	WdTableFormatList3: 26,
	WdTableFormatList4: 27,
	WdTableFormatList5: 28,
	WdTableFormatList6: 29,
	WdTableFormatList7: 30,
	WdTableFormatList8: 31,
	WdTableFormat3DEffects1: 32,
	WdTableFormat3DEffects2: 33,
	WdTableFormat3DEffects3: 34,
	WdTableFormatContemporary: 35,
	WdTableFormatElegant: 36,
	WdTableFormatProfessional: 37,
	WdTableFormatSubtle1: 38,
	WdTableFormatSubtle2: 39,
	WdTableFormatWeb1: 40,
	WdTableFormatWeb2: 41,
	WdTableFormatWeb3: 42,
}

// enum WdTableFormatApply
var WdTableFormatApply = struct {
	WdTableFormatApplyBorders int32
	WdTableFormatApplyShading int32
	WdTableFormatApplyFont int32
	WdTableFormatApplyColor int32
	WdTableFormatApplyAutoFit int32
	WdTableFormatApplyHeadingRows int32
	WdTableFormatApplyLastRow int32
	WdTableFormatApplyFirstColumn int32
	WdTableFormatApplyLastColumn int32
}{
	WdTableFormatApplyBorders: 1,
	WdTableFormatApplyShading: 2,
	WdTableFormatApplyFont: 4,
	WdTableFormatApplyColor: 8,
	WdTableFormatApplyAutoFit: 16,
	WdTableFormatApplyHeadingRows: 32,
	WdTableFormatApplyLastRow: 64,
	WdTableFormatApplyFirstColumn: 128,
	WdTableFormatApplyLastColumn: 256,
}

// enum WdLanguageID
var WdLanguageID = struct {
	WdLanguageNone int32
	WdNoProofing int32
	WdAfrikaans int32
	WdAlbanian int32
	WdAmharic int32
	WdArabicAlgeria int32
	WdArabicBahrain int32
	WdArabicEgypt int32
	WdArabicIraq int32
	WdArabicJordan int32
	WdArabicKuwait int32
	WdArabicLebanon int32
	WdArabicLibya int32
	WdArabicMorocco int32
	WdArabicOman int32
	WdArabicQatar int32
	WdArabic int32
	WdArabicSyria int32
	WdArabicTunisia int32
	WdArabicUAE int32
	WdArabicYemen int32
	WdArmenian int32
	WdAssamese int32
	WdAzeriCyrillic int32
	WdAzeriLatin int32
	WdBasque int32
	WdByelorussian int32
	WdBengali int32
	WdBulgarian int32
	WdBurmese int32
	WdCatalan int32
	WdCherokee int32
	WdChineseHongKongSAR int32
	WdChineseMacaoSAR int32
	WdSimplifiedChinese int32
	WdChineseSingapore int32
	WdTraditionalChinese int32
	WdCroatian int32
	WdCzech int32
	WdDanish int32
	WdDivehi int32
	WdBelgianDutch int32
	WdDutch int32
	WdEdo int32
	WdEnglishAUS int32
	WdEnglishBelize int32
	WdEnglishCanadian int32
	WdEnglishCaribbean int32
	WdEnglishIreland int32
	WdEnglishJamaica int32
	WdEnglishNewZealand int32
	WdEnglishPhilippines int32
	WdEnglishSouthAfrica int32
	WdEnglishTrinidadTobago int32
	WdEnglishUK int32
	WdEnglishUS int32
	WdEnglishZimbabwe int32
	WdEnglishIndonesia int32
	WdEstonian int32
	WdFaeroese int32
	WdPersian int32
	WdFilipino int32
	WdFinnish int32
	WdFulfulde int32
	WdBelgianFrench int32
	WdFrenchCameroon int32
	WdFrenchCanadian int32
	WdFrenchCotedIvoire int32
	WdFrench int32
	WdFrenchLuxembourg int32
	WdFrenchMali int32
	WdFrenchMonaco int32
	WdFrenchReunion int32
	WdFrenchSenegal int32
	WdFrenchMorocco int32
	WdFrenchHaiti int32
	WdSwissFrench int32
	WdFrenchWestIndies int32
	WdFrenchCongoDRC int32
	WdFrisianNetherlands int32
	WdGaelicIreland int32
	WdGaelicScotland int32
	WdGalician int32
	WdGeorgian int32
	WdGermanAustria int32
	WdGerman int32
	WdGermanLiechtenstein int32
	WdGermanLuxembourg int32
	WdSwissGerman int32
	WdGreek int32
	WdGuarani int32
	WdGujarati int32
	WdHausa int32
	WdHawaiian int32
	WdHebrew int32
	WdHindi int32
	WdHungarian int32
	WdIbibio int32
	WdIcelandic int32
	WdIgbo int32
	WdIndonesian int32
	WdInuktitut int32
	WdItalian int32
	WdSwissItalian int32
	WdJapanese int32
	WdKannada int32
	WdKanuri int32
	WdKashmiri int32
	WdKazakh int32
	WdKhmer int32
	WdKirghiz int32
	WdKonkani int32
	WdKorean int32
	WdKyrgyz int32
	WdLao int32
	WdLatin int32
	WdLatvian int32
	WdLithuanian int32
	WdMacedonianFYROM int32
	WdMalaysian int32
	WdMalayBruneiDarussalam int32
	WdMalayalam int32
	WdMaltese int32
	WdManipuri int32
	WdMarathi int32
	WdMongolian int32
	WdNepali int32
	WdNorwegianBokmol int32
	WdNorwegianNynorsk int32
	WdOriya int32
	WdOromo int32
	WdPashto int32
	WdPolish int32
	WdPortugueseBrazil int32
	WdPortuguese int32
	WdPunjabi int32
	WdRhaetoRomanic int32
	WdRomanianMoldova int32
	WdRomanian int32
	WdRussianMoldova int32
	WdRussian int32
	WdSamiLappish int32
	WdSanskrit int32
	WdSerbianCyrillic int32
	WdSerbianLatin int32
	WdSinhalese int32
	WdSindhi int32
	WdSindhiPakistan int32
	WdSlovak int32
	WdSlovenian int32
	WdSomali int32
	WdSorbian int32
	WdSpanishArgentina int32
	WdSpanishBolivia int32
	WdSpanishChile int32
	WdSpanishColombia int32
	WdSpanishCostaRica int32
	WdSpanishDominicanRepublic int32
	WdSpanishEcuador int32
	WdSpanishElSalvador int32
	WdSpanishGuatemala int32
	WdSpanishHonduras int32
	WdMexicanSpanish int32
	WdSpanishNicaragua int32
	WdSpanishPanama int32
	WdSpanishParaguay int32
	WdSpanishPeru int32
	WdSpanishPuertoRico int32
	WdSpanishModernSort int32
	WdSpanish int32
	WdSpanishUruguay int32
	WdSpanishVenezuela int32
	WdSesotho int32
	WdSutu int32
	WdSwahili int32
	WdSwedishFinland int32
	WdSwedish int32
	WdSyriac int32
	WdTajik int32
	WdTamazight int32
	WdTamazightLatin int32
	WdTamil int32
	WdTatar int32
	WdTelugu int32
	WdThai int32
	WdTibetan int32
	WdTigrignaEthiopic int32
	WdTigrignaEritrea int32
	WdTsonga int32
	WdTswana int32
	WdTurkish int32
	WdTurkmen int32
	WdUkrainian int32
	WdUrdu int32
	WdUzbekCyrillic int32
	WdUzbekLatin int32
	WdVenda int32
	WdVietnamese int32
	WdWelsh int32
	WdXhosa int32
	WdYi int32
	WdYiddish int32
	WdYoruba int32
	WdZulu int32
}{
	WdLanguageNone: 0,
	WdNoProofing: 1024,
	WdAfrikaans: 1078,
	WdAlbanian: 1052,
	WdAmharic: 1118,
	WdArabicAlgeria: 5121,
	WdArabicBahrain: 15361,
	WdArabicEgypt: 3073,
	WdArabicIraq: 2049,
	WdArabicJordan: 11265,
	WdArabicKuwait: 13313,
	WdArabicLebanon: 12289,
	WdArabicLibya: 4097,
	WdArabicMorocco: 6145,
	WdArabicOman: 8193,
	WdArabicQatar: 16385,
	WdArabic: 1025,
	WdArabicSyria: 10241,
	WdArabicTunisia: 7169,
	WdArabicUAE: 14337,
	WdArabicYemen: 9217,
	WdArmenian: 1067,
	WdAssamese: 1101,
	WdAzeriCyrillic: 2092,
	WdAzeriLatin: 1068,
	WdBasque: 1069,
	WdByelorussian: 1059,
	WdBengali: 1093,
	WdBulgarian: 1026,
	WdBurmese: 1109,
	WdCatalan: 1027,
	WdCherokee: 1116,
	WdChineseHongKongSAR: 3076,
	WdChineseMacaoSAR: 5124,
	WdSimplifiedChinese: 2052,
	WdChineseSingapore: 4100,
	WdTraditionalChinese: 1028,
	WdCroatian: 1050,
	WdCzech: 1029,
	WdDanish: 1030,
	WdDivehi: 1125,
	WdBelgianDutch: 2067,
	WdDutch: 1043,
	WdEdo: 1126,
	WdEnglishAUS: 3081,
	WdEnglishBelize: 10249,
	WdEnglishCanadian: 4105,
	WdEnglishCaribbean: 9225,
	WdEnglishIreland: 6153,
	WdEnglishJamaica: 8201,
	WdEnglishNewZealand: 5129,
	WdEnglishPhilippines: 13321,
	WdEnglishSouthAfrica: 7177,
	WdEnglishTrinidadTobago: 11273,
	WdEnglishUK: 2057,
	WdEnglishUS: 1033,
	WdEnglishZimbabwe: 12297,
	WdEnglishIndonesia: 14345,
	WdEstonian: 1061,
	WdFaeroese: 1080,
	WdPersian: 1065,
	WdFilipino: 1124,
	WdFinnish: 1035,
	WdFulfulde: 1127,
	WdBelgianFrench: 2060,
	WdFrenchCameroon: 11276,
	WdFrenchCanadian: 3084,
	WdFrenchCotedIvoire: 12300,
	WdFrench: 1036,
	WdFrenchLuxembourg: 5132,
	WdFrenchMali: 13324,
	WdFrenchMonaco: 6156,
	WdFrenchReunion: 8204,
	WdFrenchSenegal: 10252,
	WdFrenchMorocco: 14348,
	WdFrenchHaiti: 15372,
	WdSwissFrench: 4108,
	WdFrenchWestIndies: 7180,
	WdFrenchCongoDRC: 9228,
	WdFrisianNetherlands: 1122,
	WdGaelicIreland: 2108,
	WdGaelicScotland: 1084,
	WdGalician: 1110,
	WdGeorgian: 1079,
	WdGermanAustria: 3079,
	WdGerman: 1031,
	WdGermanLiechtenstein: 5127,
	WdGermanLuxembourg: 4103,
	WdSwissGerman: 2055,
	WdGreek: 1032,
	WdGuarani: 1140,
	WdGujarati: 1095,
	WdHausa: 1128,
	WdHawaiian: 1141,
	WdHebrew: 1037,
	WdHindi: 1081,
	WdHungarian: 1038,
	WdIbibio: 1129,
	WdIcelandic: 1039,
	WdIgbo: 1136,
	WdIndonesian: 1057,
	WdInuktitut: 1117,
	WdItalian: 1040,
	WdSwissItalian: 2064,
	WdJapanese: 1041,
	WdKannada: 1099,
	WdKanuri: 1137,
	WdKashmiri: 1120,
	WdKazakh: 1087,
	WdKhmer: 1107,
	WdKirghiz: 1088,
	WdKonkani: 1111,
	WdKorean: 1042,
	WdKyrgyz: 1088,
	WdLao: 1108,
	WdLatin: 1142,
	WdLatvian: 1062,
	WdLithuanian: 1063,
	WdMacedonianFYROM: 1071,
	WdMalaysian: 1086,
	WdMalayBruneiDarussalam: 2110,
	WdMalayalam: 1100,
	WdMaltese: 1082,
	WdManipuri: 1112,
	WdMarathi: 1102,
	WdMongolian: 1104,
	WdNepali: 1121,
	WdNorwegianBokmol: 1044,
	WdNorwegianNynorsk: 2068,
	WdOriya: 1096,
	WdOromo: 1138,
	WdPashto: 1123,
	WdPolish: 1045,
	WdPortugueseBrazil: 1046,
	WdPortuguese: 2070,
	WdPunjabi: 1094,
	WdRhaetoRomanic: 1047,
	WdRomanianMoldova: 2072,
	WdRomanian: 1048,
	WdRussianMoldova: 2073,
	WdRussian: 1049,
	WdSamiLappish: 1083,
	WdSanskrit: 1103,
	WdSerbianCyrillic: 3098,
	WdSerbianLatin: 2074,
	WdSinhalese: 1115,
	WdSindhi: 1113,
	WdSindhiPakistan: 2137,
	WdSlovak: 1051,
	WdSlovenian: 1060,
	WdSomali: 1143,
	WdSorbian: 1070,
	WdSpanishArgentina: 11274,
	WdSpanishBolivia: 16394,
	WdSpanishChile: 13322,
	WdSpanishColombia: 9226,
	WdSpanishCostaRica: 5130,
	WdSpanishDominicanRepublic: 7178,
	WdSpanishEcuador: 12298,
	WdSpanishElSalvador: 17418,
	WdSpanishGuatemala: 4106,
	WdSpanishHonduras: 18442,
	WdMexicanSpanish: 2058,
	WdSpanishNicaragua: 19466,
	WdSpanishPanama: 6154,
	WdSpanishParaguay: 15370,
	WdSpanishPeru: 10250,
	WdSpanishPuertoRico: 20490,
	WdSpanishModernSort: 3082,
	WdSpanish: 1034,
	WdSpanishUruguay: 14346,
	WdSpanishVenezuela: 8202,
	WdSesotho: 1072,
	WdSutu: 1072,
	WdSwahili: 1089,
	WdSwedishFinland: 2077,
	WdSwedish: 1053,
	WdSyriac: 1114,
	WdTajik: 1064,
	WdTamazight: 1119,
	WdTamazightLatin: 2143,
	WdTamil: 1097,
	WdTatar: 1092,
	WdTelugu: 1098,
	WdThai: 1054,
	WdTibetan: 1105,
	WdTigrignaEthiopic: 1139,
	WdTigrignaEritrea: 2163,
	WdTsonga: 1073,
	WdTswana: 1074,
	WdTurkish: 1055,
	WdTurkmen: 1090,
	WdUkrainian: 1058,
	WdUrdu: 1056,
	WdUzbekCyrillic: 2115,
	WdUzbekLatin: 1091,
	WdVenda: 1075,
	WdVietnamese: 1066,
	WdWelsh: 1106,
	WdXhosa: 1076,
	WdYi: 1144,
	WdYiddish: 1085,
	WdYoruba: 1130,
	WdZulu: 1077,
}

// enum WdFieldType
var WdFieldType = struct {
	WdFieldEmpty int32
	WdFieldRef int32
	WdFieldIndexEntry int32
	WdFieldFootnoteRef int32
	WdFieldSet int32
	WdFieldIf int32
	WdFieldIndex int32
	WdFieldTOCEntry int32
	WdFieldStyleRef int32
	WdFieldRefDoc int32
	WdFieldSequence int32
	WdFieldTOC int32
	WdFieldInfo int32
	WdFieldTitle int32
	WdFieldSubject int32
	WdFieldAuthor int32
	WdFieldKeyWord int32
	WdFieldComments int32
	WdFieldLastSavedBy int32
	WdFieldCreateDate int32
	WdFieldSaveDate int32
	WdFieldPrintDate int32
	WdFieldRevisionNum int32
	WdFieldEditTime int32
	WdFieldNumPages int32
	WdFieldNumWords int32
	WdFieldNumChars int32
	WdFieldFileName int32
	WdFieldTemplate int32
	WdFieldDate int32
	WdFieldTime int32
	WdFieldPage int32
	WdFieldExpression int32
	WdFieldQuote int32
	WdFieldInclude int32
	WdFieldPageRef int32
	WdFieldAsk int32
	WdFieldFillIn int32
	WdFieldData int32
	WdFieldNext int32
	WdFieldNextIf int32
	WdFieldSkipIf int32
	WdFieldMergeRec int32
	WdFieldDDE int32
	WdFieldDDEAuto int32
	WdFieldGlossary int32
	WdFieldPrint int32
	WdFieldFormula int32
	WdFieldGoToButton int32
	WdFieldMacroButton int32
	WdFieldAutoNumOutline int32
	WdFieldAutoNumLegal int32
	WdFieldAutoNum int32
	WdFieldImport int32
	WdFieldLink int32
	WdFieldSymbol int32
	WdFieldEmbed int32
	WdFieldMergeField int32
	WdFieldUserName int32
	WdFieldUserInitials int32
	WdFieldUserAddress int32
	WdFieldBarCode int32
	WdFieldDocVariable int32
	WdFieldSection int32
	WdFieldSectionPages int32
	WdFieldIncludePicture int32
	WdFieldIncludeText int32
	WdFieldFileSize int32
	WdFieldFormTextInput int32
	WdFieldFormCheckBox int32
	WdFieldNoteRef int32
	WdFieldTOA int32
	WdFieldTOAEntry int32
	WdFieldMergeSeq int32
	WdFieldPrivate int32
	WdFieldDatabase int32
	WdFieldAutoText int32
	WdFieldCompare int32
	WdFieldAddin int32
	WdFieldSubscriber int32
	WdFieldFormDropDown int32
	WdFieldAdvance int32
	WdFieldDocProperty int32
	WdFieldOCX int32
	WdFieldHyperlink int32
	WdFieldAutoTextList int32
	WdFieldListNum int32
	WdFieldHTMLActiveX int32
	WdFieldBidiOutline int32
	WdFieldAddressBlock int32
	WdFieldGreetingLine int32
	WdFieldShape int32
	WdFieldCitation int32
	WdFieldBibliography int32
}{
	WdFieldEmpty: -1,
	WdFieldRef: 3,
	WdFieldIndexEntry: 4,
	WdFieldFootnoteRef: 5,
	WdFieldSet: 6,
	WdFieldIf: 7,
	WdFieldIndex: 8,
	WdFieldTOCEntry: 9,
	WdFieldStyleRef: 10,
	WdFieldRefDoc: 11,
	WdFieldSequence: 12,
	WdFieldTOC: 13,
	WdFieldInfo: 14,
	WdFieldTitle: 15,
	WdFieldSubject: 16,
	WdFieldAuthor: 17,
	WdFieldKeyWord: 18,
	WdFieldComments: 19,
	WdFieldLastSavedBy: 20,
	WdFieldCreateDate: 21,
	WdFieldSaveDate: 22,
	WdFieldPrintDate: 23,
	WdFieldRevisionNum: 24,
	WdFieldEditTime: 25,
	WdFieldNumPages: 26,
	WdFieldNumWords: 27,
	WdFieldNumChars: 28,
	WdFieldFileName: 29,
	WdFieldTemplate: 30,
	WdFieldDate: 31,
	WdFieldTime: 32,
	WdFieldPage: 33,
	WdFieldExpression: 34,
	WdFieldQuote: 35,
	WdFieldInclude: 36,
	WdFieldPageRef: 37,
	WdFieldAsk: 38,
	WdFieldFillIn: 39,
	WdFieldData: 40,
	WdFieldNext: 41,
	WdFieldNextIf: 42,
	WdFieldSkipIf: 43,
	WdFieldMergeRec: 44,
	WdFieldDDE: 45,
	WdFieldDDEAuto: 46,
	WdFieldGlossary: 47,
	WdFieldPrint: 48,
	WdFieldFormula: 49,
	WdFieldGoToButton: 50,
	WdFieldMacroButton: 51,
	WdFieldAutoNumOutline: 52,
	WdFieldAutoNumLegal: 53,
	WdFieldAutoNum: 54,
	WdFieldImport: 55,
	WdFieldLink: 56,
	WdFieldSymbol: 57,
	WdFieldEmbed: 58,
	WdFieldMergeField: 59,
	WdFieldUserName: 60,
	WdFieldUserInitials: 61,
	WdFieldUserAddress: 62,
	WdFieldBarCode: 63,
	WdFieldDocVariable: 64,
	WdFieldSection: 65,
	WdFieldSectionPages: 66,
	WdFieldIncludePicture: 67,
	WdFieldIncludeText: 68,
	WdFieldFileSize: 69,
	WdFieldFormTextInput: 70,
	WdFieldFormCheckBox: 71,
	WdFieldNoteRef: 72,
	WdFieldTOA: 73,
	WdFieldTOAEntry: 74,
	WdFieldMergeSeq: 75,
	WdFieldPrivate: 77,
	WdFieldDatabase: 78,
	WdFieldAutoText: 79,
	WdFieldCompare: 80,
	WdFieldAddin: 81,
	WdFieldSubscriber: 82,
	WdFieldFormDropDown: 83,
	WdFieldAdvance: 84,
	WdFieldDocProperty: 85,
	WdFieldOCX: 87,
	WdFieldHyperlink: 88,
	WdFieldAutoTextList: 89,
	WdFieldListNum: 90,
	WdFieldHTMLActiveX: 91,
	WdFieldBidiOutline: 92,
	WdFieldAddressBlock: 93,
	WdFieldGreetingLine: 94,
	WdFieldShape: 95,
	WdFieldCitation: 96,
	WdFieldBibliography: 97,
}

// enum WdBuiltinStyle
var WdBuiltinStyle = struct {
	WdStyleNormal int32
	WdStyleEnvelopeAddress int32
	WdStyleEnvelopeReturn int32
	WdStyleBodyText int32
	WdStyleHeading1 int32
	WdStyleHeading2 int32
	WdStyleHeading3 int32
	WdStyleHeading4 int32
	WdStyleHeading5 int32
	WdStyleHeading6 int32
	WdStyleHeading7 int32
	WdStyleHeading8 int32
	WdStyleHeading9 int32
	WdStyleIndex1 int32
	WdStyleIndex2 int32
	WdStyleIndex3 int32
	WdStyleIndex4 int32
	WdStyleIndex5 int32
	WdStyleIndex6 int32
	WdStyleIndex7 int32
	WdStyleIndex8 int32
	WdStyleIndex9 int32
	WdStyleTOC1 int32
	WdStyleTOC2 int32
	WdStyleTOC3 int32
	WdStyleTOC4 int32
	WdStyleTOC5 int32
	WdStyleTOC6 int32
	WdStyleTOC7 int32
	WdStyleTOC8 int32
	WdStyleTOC9 int32
	WdStyleNormalIndent int32
	WdStyleFootnoteText int32
	WdStyleCommentText int32
	WdStyleHeader int32
	WdStyleFooter int32
	WdStyleIndexHeading int32
	WdStyleCaption int32
	WdStyleTableOfFigures int32
	WdStyleFootnoteReference int32
	WdStyleCommentReference int32
	WdStyleLineNumber int32
	WdStylePageNumber int32
	WdStyleEndnoteReference int32
	WdStyleEndnoteText int32
	WdStyleTableOfAuthorities int32
	WdStyleMacroText int32
	WdStyleTOAHeading int32
	WdStyleList int32
	WdStyleListBullet int32
	WdStyleListNumber int32
	WdStyleList2 int32
	WdStyleList3 int32
	WdStyleList4 int32
	WdStyleList5 int32
	WdStyleListBullet2 int32
	WdStyleListBullet3 int32
	WdStyleListBullet4 int32
	WdStyleListBullet5 int32
	WdStyleListNumber2 int32
	WdStyleListNumber3 int32
	WdStyleListNumber4 int32
	WdStyleListNumber5 int32
	WdStyleTitle int32
	WdStyleClosing int32
	WdStyleSignature int32
	WdStyleDefaultParagraphFont int32
	WdStyleBodyTextIndent int32
	WdStyleListContinue int32
	WdStyleListContinue2 int32
	WdStyleListContinue3 int32
	WdStyleListContinue4 int32
	WdStyleListContinue5 int32
	WdStyleMessageHeader int32
	WdStyleSubtitle int32
	WdStyleSalutation int32
	WdStyleDate int32
	WdStyleBodyTextFirstIndent int32
	WdStyleBodyTextFirstIndent2 int32
	WdStyleNoteHeading int32
	WdStyleBodyText2 int32
	WdStyleBodyText3 int32
	WdStyleBodyTextIndent2 int32
	WdStyleBodyTextIndent3 int32
	WdStyleBlockQuotation int32
	WdStyleHyperlink int32
	WdStyleHyperlinkFollowed int32
	WdStyleStrong int32
	WdStyleEmphasis int32
	WdStyleNavPane int32
	WdStylePlainText int32
	WdStyleHtmlNormal int32
	WdStyleHtmlAcronym int32
	WdStyleHtmlAddress int32
	WdStyleHtmlCite int32
	WdStyleHtmlCode int32
	WdStyleHtmlDfn int32
	WdStyleHtmlKbd int32
	WdStyleHtmlPre int32
	WdStyleHtmlSamp int32
	WdStyleHtmlTt int32
	WdStyleHtmlVar int32
	WdStyleNormalTable int32
	WdStyleNormalObject int32
	WdStyleTableLightShading int32
	WdStyleTableLightList int32
	WdStyleTableLightGrid int32
	WdStyleTableMediumShading1 int32
	WdStyleTableMediumShading2 int32
	WdStyleTableMediumList1 int32
	WdStyleTableMediumList2 int32
	WdStyleTableMediumGrid1 int32
	WdStyleTableMediumGrid2 int32
	WdStyleTableMediumGrid3 int32
	WdStyleTableDarkList int32
	WdStyleTableColorfulShading int32
	WdStyleTableColorfulList int32
	WdStyleTableColorfulGrid int32
	WdStyleTableLightShadingAccent1 int32
	WdStyleTableLightListAccent1 int32
	WdStyleTableLightGridAccent1 int32
	WdStyleTableMediumShading1Accent1 int32
	WdStyleTableMediumShading2Accent1 int32
	WdStyleTableMediumList1Accent1 int32
	WdStyleListParagraph int32
	WdStyleQuote int32
	WdStyleIntenseQuote int32
	WdStyleSubtleEmphasis int32
	WdStyleIntenseEmphasis int32
	WdStyleSubtleReference int32
	WdStyleIntenseReference int32
	WdStyleBookTitle int32
	WdStyleBibliography int32
	WdStyleTocHeading int32
}{
	WdStyleNormal: -1,
	WdStyleEnvelopeAddress: -37,
	WdStyleEnvelopeReturn: -38,
	WdStyleBodyText: -67,
	WdStyleHeading1: -2,
	WdStyleHeading2: -3,
	WdStyleHeading3: -4,
	WdStyleHeading4: -5,
	WdStyleHeading5: -6,
	WdStyleHeading6: -7,
	WdStyleHeading7: -8,
	WdStyleHeading8: -9,
	WdStyleHeading9: -10,
	WdStyleIndex1: -11,
	WdStyleIndex2: -12,
	WdStyleIndex3: -13,
	WdStyleIndex4: -14,
	WdStyleIndex5: -15,
	WdStyleIndex6: -16,
	WdStyleIndex7: -17,
	WdStyleIndex8: -18,
	WdStyleIndex9: -19,
	WdStyleTOC1: -20,
	WdStyleTOC2: -21,
	WdStyleTOC3: -22,
	WdStyleTOC4: -23,
	WdStyleTOC5: -24,
	WdStyleTOC6: -25,
	WdStyleTOC7: -26,
	WdStyleTOC8: -27,
	WdStyleTOC9: -28,
	WdStyleNormalIndent: -29,
	WdStyleFootnoteText: -30,
	WdStyleCommentText: -31,
	WdStyleHeader: -32,
	WdStyleFooter: -33,
	WdStyleIndexHeading: -34,
	WdStyleCaption: -35,
	WdStyleTableOfFigures: -36,
	WdStyleFootnoteReference: -39,
	WdStyleCommentReference: -40,
	WdStyleLineNumber: -41,
	WdStylePageNumber: -42,
	WdStyleEndnoteReference: -43,
	WdStyleEndnoteText: -44,
	WdStyleTableOfAuthorities: -45,
	WdStyleMacroText: -46,
	WdStyleTOAHeading: -47,
	WdStyleList: -48,
	WdStyleListBullet: -49,
	WdStyleListNumber: -50,
	WdStyleList2: -51,
	WdStyleList3: -52,
	WdStyleList4: -53,
	WdStyleList5: -54,
	WdStyleListBullet2: -55,
	WdStyleListBullet3: -56,
	WdStyleListBullet4: -57,
	WdStyleListBullet5: -58,
	WdStyleListNumber2: -59,
	WdStyleListNumber3: -60,
	WdStyleListNumber4: -61,
	WdStyleListNumber5: -62,
	WdStyleTitle: -63,
	WdStyleClosing: -64,
	WdStyleSignature: -65,
	WdStyleDefaultParagraphFont: -66,
	WdStyleBodyTextIndent: -68,
	WdStyleListContinue: -69,
	WdStyleListContinue2: -70,
	WdStyleListContinue3: -71,
	WdStyleListContinue4: -72,
	WdStyleListContinue5: -73,
	WdStyleMessageHeader: -74,
	WdStyleSubtitle: -75,
	WdStyleSalutation: -76,
	WdStyleDate: -77,
	WdStyleBodyTextFirstIndent: -78,
	WdStyleBodyTextFirstIndent2: -79,
	WdStyleNoteHeading: -80,
	WdStyleBodyText2: -81,
	WdStyleBodyText3: -82,
	WdStyleBodyTextIndent2: -83,
	WdStyleBodyTextIndent3: -84,
	WdStyleBlockQuotation: -85,
	WdStyleHyperlink: -86,
	WdStyleHyperlinkFollowed: -87,
	WdStyleStrong: -88,
	WdStyleEmphasis: -89,
	WdStyleNavPane: -90,
	WdStylePlainText: -91,
	WdStyleHtmlNormal: -95,
	WdStyleHtmlAcronym: -96,
	WdStyleHtmlAddress: -97,
	WdStyleHtmlCite: -98,
	WdStyleHtmlCode: -99,
	WdStyleHtmlDfn: -100,
	WdStyleHtmlKbd: -101,
	WdStyleHtmlPre: -102,
	WdStyleHtmlSamp: -103,
	WdStyleHtmlTt: -104,
	WdStyleHtmlVar: -105,
	WdStyleNormalTable: -106,
	WdStyleNormalObject: -158,
	WdStyleTableLightShading: -159,
	WdStyleTableLightList: -160,
	WdStyleTableLightGrid: -161,
	WdStyleTableMediumShading1: -162,
	WdStyleTableMediumShading2: -163,
	WdStyleTableMediumList1: -164,
	WdStyleTableMediumList2: -165,
	WdStyleTableMediumGrid1: -166,
	WdStyleTableMediumGrid2: -167,
	WdStyleTableMediumGrid3: -168,
	WdStyleTableDarkList: -169,
	WdStyleTableColorfulShading: -170,
	WdStyleTableColorfulList: -171,
	WdStyleTableColorfulGrid: -172,
	WdStyleTableLightShadingAccent1: -173,
	WdStyleTableLightListAccent1: -174,
	WdStyleTableLightGridAccent1: -175,
	WdStyleTableMediumShading1Accent1: -176,
	WdStyleTableMediumShading2Accent1: -177,
	WdStyleTableMediumList1Accent1: -178,
	WdStyleListParagraph: -180,
	WdStyleQuote: -181,
	WdStyleIntenseQuote: -182,
	WdStyleSubtleEmphasis: -261,
	WdStyleIntenseEmphasis: -262,
	WdStyleSubtleReference: -263,
	WdStyleIntenseReference: -264,
	WdStyleBookTitle: -265,
	WdStyleBibliography: -266,
	WdStyleTocHeading: -267,
}

// enum WdWordDialogTab
var WdWordDialogTab = struct {
	WdDialogToolsOptionsTabView int32
	WdDialogToolsOptionsTabGeneral int32
	WdDialogToolsOptionsTabEdit int32
	WdDialogToolsOptionsTabPrint int32
	WdDialogToolsOptionsTabSave int32
	WdDialogToolsOptionsTabProofread int32
	WdDialogToolsOptionsTabTrackChanges int32
	WdDialogToolsOptionsTabUserInfo int32
	WdDialogToolsOptionsTabCompatibility int32
	WdDialogToolsOptionsTabTypography int32
	WdDialogToolsOptionsTabFileLocations int32
	WdDialogToolsOptionsTabFuzzy int32
	WdDialogToolsOptionsTabHangulHanjaConversion int32
	WdDialogToolsOptionsTabBidi int32
	WdDialogToolsOptionsTabSecurity int32
	WdDialogFilePageSetupTabMargins int32
	WdDialogFilePageSetupTabPaper int32
	WdDialogFilePageSetupTabLayout int32
	WdDialogFilePageSetupTabCharsLines int32
	WdDialogInsertSymbolTabSymbols int32
	WdDialogInsertSymbolTabSpecialCharacters int32
	WdDialogNoteOptionsTabAllFootnotes int32
	WdDialogNoteOptionsTabAllEndnotes int32
	WdDialogInsertIndexAndTablesTabIndex int32
	WdDialogInsertIndexAndTablesTabTableOfContents int32
	WdDialogInsertIndexAndTablesTabTableOfFigures int32
	WdDialogInsertIndexAndTablesTabTableOfAuthorities int32
	WdDialogOrganizerTabStyles int32
	WdDialogOrganizerTabAutoText int32
	WdDialogOrganizerTabCommandBars int32
	WdDialogOrganizerTabMacros int32
	WdDialogFormatFontTabFont int32
	WdDialogFormatFontTabCharacterSpacing int32
	WdDialogFormatFontTabAnimation int32
	WdDialogFormatBordersAndShadingTabBorders int32
	WdDialogFormatBordersAndShadingTabPageBorder int32
	WdDialogFormatBordersAndShadingTabShading int32
	WdDialogToolsEnvelopesAndLabelsTabEnvelopes int32
	WdDialogToolsEnvelopesAndLabelsTabLabels int32
	WdDialogFormatParagraphTabIndentsAndSpacing int32
	WdDialogFormatParagraphTabTextFlow int32
	WdDialogFormatParagraphTabTeisai int32
	WdDialogFormatDrawingObjectTabColorsAndLines int32
	WdDialogFormatDrawingObjectTabSize int32
	WdDialogFormatDrawingObjectTabPosition int32
	WdDialogFormatDrawingObjectTabWrapping int32
	WdDialogFormatDrawingObjectTabPicture int32
	WdDialogFormatDrawingObjectTabTextbox int32
	WdDialogFormatDrawingObjectTabWeb int32
	WdDialogFormatDrawingObjectTabHR int32
	WdDialogToolsAutoCorrectExceptionsTabFirstLetter int32
	WdDialogToolsAutoCorrectExceptionsTabInitialCaps int32
	WdDialogToolsAutoCorrectExceptionsTabHangulAndAlphabet int32
	WdDialogToolsAutoCorrectExceptionsTabIac int32
	WdDialogFormatBulletsAndNumberingTabBulleted int32
	WdDialogFormatBulletsAndNumberingTabNumbered int32
	WdDialogFormatBulletsAndNumberingTabOutlineNumbered int32
	WdDialogLetterWizardTabLetterFormat int32
	WdDialogLetterWizardTabRecipientInfo int32
	WdDialogLetterWizardTabOtherElements int32
	WdDialogLetterWizardTabSenderInfo int32
	WdDialogToolsAutoManagerTabAutoCorrect int32
	WdDialogToolsAutoManagerTabAutoFormatAsYouType int32
	WdDialogToolsAutoManagerTabAutoText int32
	WdDialogToolsAutoManagerTabAutoFormat int32
	WdDialogToolsAutoManagerTabSmartTags int32
	WdDialogTablePropertiesTabTable int32
	WdDialogTablePropertiesTabRow int32
	WdDialogTablePropertiesTabColumn int32
	WdDialogTablePropertiesTabCell int32
	WdDialogEmailOptionsTabSignature int32
	WdDialogEmailOptionsTabStationary int32
	WdDialogEmailOptionsTabQuoting int32
	WdDialogWebOptionsBrowsers int32
	WdDialogWebOptionsGeneral int32
	WdDialogWebOptionsFiles int32
	WdDialogWebOptionsPictures int32
	WdDialogWebOptionsEncoding int32
	WdDialogWebOptionsFonts int32
	WdDialogToolsOptionsTabAcetate int32
	WdDialogTemplates int32
	WdDialogTemplatesXMLSchema int32
	WdDialogTemplatesXMLExpansionPacks int32
	WdDialogTemplatesLinkedCSS int32
	WdDialogStyleManagementTabEdit int32
	WdDialogStyleManagementTabRecommend int32
	WdDialogStyleManagementTabRestrict int32
}{
	WdDialogToolsOptionsTabView: 204,
	WdDialogToolsOptionsTabGeneral: 203,
	WdDialogToolsOptionsTabEdit: 224,
	WdDialogToolsOptionsTabPrint: 208,
	WdDialogToolsOptionsTabSave: 209,
	WdDialogToolsOptionsTabProofread: 211,
	WdDialogToolsOptionsTabTrackChanges: 386,
	WdDialogToolsOptionsTabUserInfo: 213,
	WdDialogToolsOptionsTabCompatibility: 525,
	WdDialogToolsOptionsTabTypography: 739,
	WdDialogToolsOptionsTabFileLocations: 225,
	WdDialogToolsOptionsTabFuzzy: 790,
	WdDialogToolsOptionsTabHangulHanjaConversion: 786,
	WdDialogToolsOptionsTabBidi: 1029,
	WdDialogToolsOptionsTabSecurity: 1361,
	WdDialogFilePageSetupTabMargins: 150000,
	WdDialogFilePageSetupTabPaper: 150001,
	WdDialogFilePageSetupTabLayout: 150003,
	WdDialogFilePageSetupTabCharsLines: 150004,
	WdDialogInsertSymbolTabSymbols: 200000,
	WdDialogInsertSymbolTabSpecialCharacters: 200001,
	WdDialogNoteOptionsTabAllFootnotes: 300000,
	WdDialogNoteOptionsTabAllEndnotes: 300001,
	WdDialogInsertIndexAndTablesTabIndex: 400000,
	WdDialogInsertIndexAndTablesTabTableOfContents: 400001,
	WdDialogInsertIndexAndTablesTabTableOfFigures: 400002,
	WdDialogInsertIndexAndTablesTabTableOfAuthorities: 400003,
	WdDialogOrganizerTabStyles: 500000,
	WdDialogOrganizerTabAutoText: 500001,
	WdDialogOrganizerTabCommandBars: 500002,
	WdDialogOrganizerTabMacros: 500003,
	WdDialogFormatFontTabFont: 600000,
	WdDialogFormatFontTabCharacterSpacing: 600001,
	WdDialogFormatFontTabAnimation: 600002,
	WdDialogFormatBordersAndShadingTabBorders: 700000,
	WdDialogFormatBordersAndShadingTabPageBorder: 700001,
	WdDialogFormatBordersAndShadingTabShading: 700002,
	WdDialogToolsEnvelopesAndLabelsTabEnvelopes: 800000,
	WdDialogToolsEnvelopesAndLabelsTabLabels: 800001,
	WdDialogFormatParagraphTabIndentsAndSpacing: 1000000,
	WdDialogFormatParagraphTabTextFlow: 1000001,
	WdDialogFormatParagraphTabTeisai: 1000002,
	WdDialogFormatDrawingObjectTabColorsAndLines: 1200000,
	WdDialogFormatDrawingObjectTabSize: 1200001,
	WdDialogFormatDrawingObjectTabPosition: 1200002,
	WdDialogFormatDrawingObjectTabWrapping: 1200003,
	WdDialogFormatDrawingObjectTabPicture: 1200004,
	WdDialogFormatDrawingObjectTabTextbox: 1200005,
	WdDialogFormatDrawingObjectTabWeb: 1200006,
	WdDialogFormatDrawingObjectTabHR: 1200007,
	WdDialogToolsAutoCorrectExceptionsTabFirstLetter: 1400000,
	WdDialogToolsAutoCorrectExceptionsTabInitialCaps: 1400001,
	WdDialogToolsAutoCorrectExceptionsTabHangulAndAlphabet: 1400002,
	WdDialogToolsAutoCorrectExceptionsTabIac: 1400003,
	WdDialogFormatBulletsAndNumberingTabBulleted: 1500000,
	WdDialogFormatBulletsAndNumberingTabNumbered: 1500001,
	WdDialogFormatBulletsAndNumberingTabOutlineNumbered: 1500002,
	WdDialogLetterWizardTabLetterFormat: 1600000,
	WdDialogLetterWizardTabRecipientInfo: 1600001,
	WdDialogLetterWizardTabOtherElements: 1600002,
	WdDialogLetterWizardTabSenderInfo: 1600003,
	WdDialogToolsAutoManagerTabAutoCorrect: 1700000,
	WdDialogToolsAutoManagerTabAutoFormatAsYouType: 1700001,
	WdDialogToolsAutoManagerTabAutoText: 1700002,
	WdDialogToolsAutoManagerTabAutoFormat: 1700003,
	WdDialogToolsAutoManagerTabSmartTags: 1700004,
	WdDialogTablePropertiesTabTable: 1800000,
	WdDialogTablePropertiesTabRow: 1800001,
	WdDialogTablePropertiesTabColumn: 1800002,
	WdDialogTablePropertiesTabCell: 1800003,
	WdDialogEmailOptionsTabSignature: 1900000,
	WdDialogEmailOptionsTabStationary: 1900001,
	WdDialogEmailOptionsTabQuoting: 1900002,
	WdDialogWebOptionsBrowsers: 2000000,
	WdDialogWebOptionsGeneral: 2000000,
	WdDialogWebOptionsFiles: 2000001,
	WdDialogWebOptionsPictures: 2000002,
	WdDialogWebOptionsEncoding: 2000003,
	WdDialogWebOptionsFonts: 2000004,
	WdDialogToolsOptionsTabAcetate: 1266,
	WdDialogTemplates: 2100000,
	WdDialogTemplatesXMLSchema: 2100001,
	WdDialogTemplatesXMLExpansionPacks: 2100002,
	WdDialogTemplatesLinkedCSS: 2100003,
	WdDialogStyleManagementTabEdit: 2200000,
	WdDialogStyleManagementTabRecommend: 2200001,
	WdDialogStyleManagementTabRestrict: 2200002,
}

// enum WdWordDialogTabHID
var WdWordDialogTabHID = struct {
	WdDialogFilePageSetupTabPaperSize int32
	WdDialogFilePageSetupTabPaperSource int32
}{
	WdDialogFilePageSetupTabPaperSize: 150001,
	WdDialogFilePageSetupTabPaperSource: 150002,
}

// enum WdWordDialog
var WdWordDialog = struct {
	WdDialogHelpAbout int32
	WdDialogHelpWordPerfectHelp int32
	WdDialogDocumentStatistics int32
	WdDialogFileNew int32
	WdDialogFileOpen int32
	WdDialogMailMergeOpenDataSource int32
	WdDialogMailMergeOpenHeaderSource int32
	WdDialogFileSaveAs int32
	WdDialogFileSummaryInfo int32
	WdDialogToolsTemplates int32
	WdDialogFilePrint int32
	WdDialogFilePrintSetup int32
	WdDialogFileFind int32
	WdDialogFormatAddrFonts int32
	WdDialogEditPasteSpecial int32
	WdDialogEditFind int32
	WdDialogEditReplace int32
	WdDialogEditStyle int32
	WdDialogEditLinks int32
	WdDialogEditObject int32
	WdDialogTableToText int32
	WdDialogTextToTable int32
	WdDialogTableInsertTable int32
	WdDialogTableInsertCells int32
	WdDialogTableInsertRow int32
	WdDialogTableDeleteCells int32
	WdDialogTableSplitCells int32
	WdDialogTableRowHeight int32
	WdDialogTableColumnWidth int32
	WdDialogToolsCustomize int32
	WdDialogInsertBreak int32
	WdDialogInsertSymbol int32
	WdDialogInsertPicture int32
	WdDialogInsertFile int32
	WdDialogInsertDateTime int32
	WdDialogInsertField int32
	WdDialogInsertMergeField int32
	WdDialogInsertBookmark int32
	WdDialogMarkIndexEntry int32
	WdDialogInsertIndex int32
	WdDialogInsertTableOfContents int32
	WdDialogInsertObject int32
	WdDialogToolsCreateEnvelope int32
	WdDialogFormatFont int32
	WdDialogFormatParagraph int32
	WdDialogFormatSectionLayout int32
	WdDialogFormatColumns int32
	WdDialogFileDocumentLayout int32
	WdDialogFilePageSetup int32
	WdDialogFormatTabs int32
	WdDialogFormatStyle int32
	WdDialogFormatDefineStyleFont int32
	WdDialogFormatDefineStylePara int32
	WdDialogFormatDefineStyleTabs int32
	WdDialogFormatDefineStyleFrame int32
	WdDialogFormatDefineStyleBorders int32
	WdDialogFormatDefineStyleLang int32
	WdDialogFormatPicture int32
	WdDialogToolsLanguage int32
	WdDialogFormatBordersAndShading int32
	WdDialogFormatFrame int32
	WdDialogToolsThesaurus int32
	WdDialogToolsHyphenation int32
	WdDialogToolsBulletsNumbers int32
	WdDialogToolsHighlightChanges int32
	WdDialogToolsRevisions int32
	WdDialogToolsCompareDocuments int32
	WdDialogTableSort int32
	WdDialogToolsOptionsGeneral int32
	WdDialogToolsOptionsView int32
	WdDialogToolsAdvancedSettings int32
	WdDialogToolsOptionsPrint int32
	WdDialogToolsOptionsSave int32
	WdDialogToolsOptionsSpellingAndGrammar int32
	WdDialogToolsOptionsUserInfo int32
	WdDialogToolsMacroRecord int32
	WdDialogToolsMacro int32
	WdDialogWindowActivate int32
	WdDialogFormatRetAddrFonts int32
	WdDialogOrganizer int32
	WdDialogToolsOptionsEdit int32
	WdDialogToolsOptionsFileLocations int32
	WdDialogToolsWordCount int32
	WdDialogControlRun int32
	WdDialogInsertPageNumbers int32
	WdDialogFormatPageNumber int32
	WdDialogCopyFile int32
	WdDialogFormatChangeCase int32
	WdDialogUpdateTOC int32
	WdDialogInsertDatabase int32
	WdDialogTableFormula int32
	WdDialogFormFieldOptions int32
	WdDialogInsertCaption int32
	WdDialogInsertCaptionNumbering int32
	WdDialogInsertAutoCaption int32
	WdDialogFormFieldHelp int32
	WdDialogInsertCrossReference int32
	WdDialogInsertFootnote int32
	WdDialogNoteOptions int32
	WdDialogToolsAutoCorrect int32
	WdDialogToolsOptionsTrackChanges int32
	WdDialogConvertObject int32
	WdDialogInsertAddCaption int32
	WdDialogConnect int32
	WdDialogToolsCustomizeKeyboard int32
	WdDialogToolsCustomizeMenus int32
	WdDialogToolsMergeDocuments int32
	WdDialogMarkTableOfContentsEntry int32
	WdDialogFileMacPageSetupGX int32
	WdDialogFilePrintOneCopy int32
	WdDialogEditFrame int32
	WdDialogMarkCitation int32
	WdDialogTableOfContentsOptions int32
	WdDialogInsertTableOfAuthorities int32
	WdDialogInsertTableOfFigures int32
	WdDialogInsertIndexAndTables int32
	WdDialogInsertFormField int32
	WdDialogFormatDropCap int32
	WdDialogToolsCreateLabels int32
	WdDialogToolsProtectDocument int32
	WdDialogFormatStyleGallery int32
	WdDialogToolsAcceptRejectChanges int32
	WdDialogHelpWordPerfectHelpOptions int32
	WdDialogToolsUnprotectDocument int32
	WdDialogToolsOptionsCompatibility int32
	WdDialogTableOfCaptionsOptions int32
	WdDialogTableAutoFormat int32
	WdDialogMailMergeFindRecord int32
	WdDialogReviewAfmtRevisions int32
	WdDialogViewZoom int32
	WdDialogToolsProtectSection int32
	WdDialogFontSubstitution int32
	WdDialogInsertSubdocument int32
	WdDialogNewToolbar int32
	WdDialogToolsEnvelopesAndLabels int32
	WdDialogFormatCallout int32
	WdDialogTableFormatCell int32
	WdDialogToolsCustomizeMenuBar int32
	WdDialogFileRoutingSlip int32
	WdDialogEditTOACategory int32
	WdDialogToolsManageFields int32
	WdDialogDrawSnapToGrid int32
	WdDialogDrawAlign int32
	WdDialogMailMergeCreateDataSource int32
	WdDialogMailMergeCreateHeaderSource int32
	WdDialogMailMerge int32
	WdDialogMailMergeCheck int32
	WdDialogMailMergeHelper int32
	WdDialogMailMergeQueryOptions int32
	WdDialogFileMacPageSetup int32
	WdDialogListCommands int32
	WdDialogEditCreatePublisher int32
	WdDialogEditSubscribeTo int32
	WdDialogEditPublishOptions int32
	WdDialogEditSubscribeOptions int32
	WdDialogFileMacCustomPageSetupGX int32
	WdDialogToolsOptionsTypography int32
	WdDialogToolsAutoCorrectExceptions int32
	WdDialogToolsOptionsAutoFormatAsYouType int32
	WdDialogMailMergeUseAddressBook int32
	WdDialogToolsHangulHanjaConversion int32
	WdDialogToolsOptionsFuzzy int32
	WdDialogEditGoToOld int32
	WdDialogInsertNumber int32
	WdDialogLetterWizard int32
	WdDialogFormatBulletsAndNumbering int32
	WdDialogToolsSpellingAndGrammar int32
	WdDialogToolsCreateDirectory int32
	WdDialogTableWrapping int32
	WdDialogFormatTheme int32
	WdDialogTableProperties int32
	WdDialogEmailOptions int32
	WdDialogCreateAutoText int32
	WdDialogToolsAutoSummarize int32
	WdDialogToolsGrammarSettings int32
	WdDialogEditGoTo int32
	WdDialogWebOptions int32
	WdDialogInsertHyperlink int32
	WdDialogToolsAutoManager int32
	WdDialogFileVersions int32
	WdDialogToolsOptionsAutoFormat int32
	WdDialogFormatDrawingObject int32
	WdDialogToolsOptions int32
	WdDialogFitText int32
	WdDialogEditAutoText int32
	WdDialogPhoneticGuide int32
	WdDialogToolsDictionary int32
	WdDialogFileSaveVersion int32
	WdDialogToolsOptionsBidi int32
	WdDialogFrameSetProperties int32
	WdDialogTableTableOptions int32
	WdDialogTableCellOptions int32
	WdDialogIMESetDefault int32
	WdDialogTCSCTranslator int32
	WdDialogHorizontalInVertical int32
	WdDialogTwoLinesInOne int32
	WdDialogFormatEncloseCharacters int32
	WdDialogConsistencyChecker int32
	WdDialogToolsOptionsSmartTag int32
	WdDialogFormatStylesCustom int32
	WdDialogCSSLinks int32
	WdDialogInsertWebComponent int32
	WdDialogToolsOptionsEditCopyPaste int32
	WdDialogToolsOptionsSecurity int32
	WdDialogSearch int32
	WdDialogShowRepairs int32
	WdDialogMailMergeInsertAsk int32
	WdDialogMailMergeInsertFillIn int32
	WdDialogMailMergeInsertIf int32
	WdDialogMailMergeInsertNextIf int32
	WdDialogMailMergeInsertSet int32
	WdDialogMailMergeInsertSkipIf int32
	WdDialogMailMergeFieldMapping int32
	WdDialogMailMergeInsertAddressBlock int32
	WdDialogMailMergeInsertGreetingLine int32
	WdDialogMailMergeInsertFields int32
	WdDialogMailMergeRecipients int32
	WdDialogMailMergeFindRecipient int32
	WdDialogMailMergeSetDocumentType int32
	WdDialogLabelOptions int32
	WdDialogXMLElementAttributes int32
	WdDialogSchemaLibrary int32
	WdDialogPermission int32
	WdDialogMyPermission int32
	WdDialogXMLOptions int32
	WdDialogFormattingRestrictions int32
	WdDialogSourceManager int32
	WdDialogCreateSource int32
	WdDialogDocumentInspector int32
	WdDialogStyleManagement int32
	WdDialogInsertSource int32
	WdDialogOMathRecognizedFunctions int32
	WdDialogInsertPlaceholder int32
	WdDialogBuildingBlockOrganizer int32
	WdDialogContentControlProperties int32
	WdDialogCompatibilityChecker int32
	WdDialogExportAsFixedFormat int32
	WdDialogFileNew2007 int32
}{
	WdDialogHelpAbout: 9,
	WdDialogHelpWordPerfectHelp: 10,
	WdDialogDocumentStatistics: 78,
	WdDialogFileNew: 79,
	WdDialogFileOpen: 80,
	WdDialogMailMergeOpenDataSource: 81,
	WdDialogMailMergeOpenHeaderSource: 82,
	WdDialogFileSaveAs: 84,
	WdDialogFileSummaryInfo: 86,
	WdDialogToolsTemplates: 87,
	WdDialogFilePrint: 88,
	WdDialogFilePrintSetup: 97,
	WdDialogFileFind: 99,
	WdDialogFormatAddrFonts: 103,
	WdDialogEditPasteSpecial: 111,
	WdDialogEditFind: 112,
	WdDialogEditReplace: 117,
	WdDialogEditStyle: 120,
	WdDialogEditLinks: 124,
	WdDialogEditObject: 125,
	WdDialogTableToText: 128,
	WdDialogTextToTable: 127,
	WdDialogTableInsertTable: 129,
	WdDialogTableInsertCells: 130,
	WdDialogTableInsertRow: 131,
	WdDialogTableDeleteCells: 133,
	WdDialogTableSplitCells: 137,
	WdDialogTableRowHeight: 142,
	WdDialogTableColumnWidth: 143,
	WdDialogToolsCustomize: 152,
	WdDialogInsertBreak: 159,
	WdDialogInsertSymbol: 162,
	WdDialogInsertPicture: 163,
	WdDialogInsertFile: 164,
	WdDialogInsertDateTime: 165,
	WdDialogInsertField: 166,
	WdDialogInsertMergeField: 167,
	WdDialogInsertBookmark: 168,
	WdDialogMarkIndexEntry: 169,
	WdDialogInsertIndex: 170,
	WdDialogInsertTableOfContents: 171,
	WdDialogInsertObject: 172,
	WdDialogToolsCreateEnvelope: 173,
	WdDialogFormatFont: 174,
	WdDialogFormatParagraph: 175,
	WdDialogFormatSectionLayout: 176,
	WdDialogFormatColumns: 177,
	WdDialogFileDocumentLayout: 178,
	WdDialogFilePageSetup: 178,
	WdDialogFormatTabs: 179,
	WdDialogFormatStyle: 180,
	WdDialogFormatDefineStyleFont: 181,
	WdDialogFormatDefineStylePara: 182,
	WdDialogFormatDefineStyleTabs: 183,
	WdDialogFormatDefineStyleFrame: 184,
	WdDialogFormatDefineStyleBorders: 185,
	WdDialogFormatDefineStyleLang: 186,
	WdDialogFormatPicture: 187,
	WdDialogToolsLanguage: 188,
	WdDialogFormatBordersAndShading: 189,
	WdDialogFormatFrame: 190,
	WdDialogToolsThesaurus: 194,
	WdDialogToolsHyphenation: 195,
	WdDialogToolsBulletsNumbers: 196,
	WdDialogToolsHighlightChanges: 197,
	WdDialogToolsRevisions: 197,
	WdDialogToolsCompareDocuments: 198,
	WdDialogTableSort: 199,
	WdDialogToolsOptionsGeneral: 203,
	WdDialogToolsOptionsView: 204,
	WdDialogToolsAdvancedSettings: 206,
	WdDialogToolsOptionsPrint: 208,
	WdDialogToolsOptionsSave: 209,
	WdDialogToolsOptionsSpellingAndGrammar: 211,
	WdDialogToolsOptionsUserInfo: 213,
	WdDialogToolsMacroRecord: 214,
	WdDialogToolsMacro: 215,
	WdDialogWindowActivate: 220,
	WdDialogFormatRetAddrFonts: 221,
	WdDialogOrganizer: 222,
	WdDialogToolsOptionsEdit: 224,
	WdDialogToolsOptionsFileLocations: 225,
	WdDialogToolsWordCount: 228,
	WdDialogControlRun: 235,
	WdDialogInsertPageNumbers: 294,
	WdDialogFormatPageNumber: 298,
	WdDialogCopyFile: 300,
	WdDialogFormatChangeCase: 322,
	WdDialogUpdateTOC: 331,
	WdDialogInsertDatabase: 341,
	WdDialogTableFormula: 348,
	WdDialogFormFieldOptions: 353,
	WdDialogInsertCaption: 357,
	WdDialogInsertCaptionNumbering: 358,
	WdDialogInsertAutoCaption: 359,
	WdDialogFormFieldHelp: 361,
	WdDialogInsertCrossReference: 367,
	WdDialogInsertFootnote: 370,
	WdDialogNoteOptions: 373,
	WdDialogToolsAutoCorrect: 378,
	WdDialogToolsOptionsTrackChanges: 386,
	WdDialogConvertObject: 392,
	WdDialogInsertAddCaption: 402,
	WdDialogConnect: 420,
	WdDialogToolsCustomizeKeyboard: 432,
	WdDialogToolsCustomizeMenus: 433,
	WdDialogToolsMergeDocuments: 435,
	WdDialogMarkTableOfContentsEntry: 442,
	WdDialogFileMacPageSetupGX: 444,
	WdDialogFilePrintOneCopy: 445,
	WdDialogEditFrame: 458,
	WdDialogMarkCitation: 463,
	WdDialogTableOfContentsOptions: 470,
	WdDialogInsertTableOfAuthorities: 471,
	WdDialogInsertTableOfFigures: 472,
	WdDialogInsertIndexAndTables: 473,
	WdDialogInsertFormField: 483,
	WdDialogFormatDropCap: 488,
	WdDialogToolsCreateLabels: 489,
	WdDialogToolsProtectDocument: 503,
	WdDialogFormatStyleGallery: 505,
	WdDialogToolsAcceptRejectChanges: 506,
	WdDialogHelpWordPerfectHelpOptions: 511,
	WdDialogToolsUnprotectDocument: 521,
	WdDialogToolsOptionsCompatibility: 525,
	WdDialogTableOfCaptionsOptions: 551,
	WdDialogTableAutoFormat: 563,
	WdDialogMailMergeFindRecord: 569,
	WdDialogReviewAfmtRevisions: 570,
	WdDialogViewZoom: 577,
	WdDialogToolsProtectSection: 578,
	WdDialogFontSubstitution: 581,
	WdDialogInsertSubdocument: 583,
	WdDialogNewToolbar: 586,
	WdDialogToolsEnvelopesAndLabels: 607,
	WdDialogFormatCallout: 610,
	WdDialogTableFormatCell: 612,
	WdDialogToolsCustomizeMenuBar: 615,
	WdDialogFileRoutingSlip: 624,
	WdDialogEditTOACategory: 625,
	WdDialogToolsManageFields: 631,
	WdDialogDrawSnapToGrid: 633,
	WdDialogDrawAlign: 634,
	WdDialogMailMergeCreateDataSource: 642,
	WdDialogMailMergeCreateHeaderSource: 643,
	WdDialogMailMerge: 676,
	WdDialogMailMergeCheck: 677,
	WdDialogMailMergeHelper: 680,
	WdDialogMailMergeQueryOptions: 681,
	WdDialogFileMacPageSetup: 685,
	WdDialogListCommands: 723,
	WdDialogEditCreatePublisher: 732,
	WdDialogEditSubscribeTo: 733,
	WdDialogEditPublishOptions: 735,
	WdDialogEditSubscribeOptions: 736,
	WdDialogFileMacCustomPageSetupGX: 737,
	WdDialogToolsOptionsTypography: 739,
	WdDialogToolsAutoCorrectExceptions: 762,
	WdDialogToolsOptionsAutoFormatAsYouType: 778,
	WdDialogMailMergeUseAddressBook: 779,
	WdDialogToolsHangulHanjaConversion: 784,
	WdDialogToolsOptionsFuzzy: 790,
	WdDialogEditGoToOld: 811,
	WdDialogInsertNumber: 812,
	WdDialogLetterWizard: 821,
	WdDialogFormatBulletsAndNumbering: 824,
	WdDialogToolsSpellingAndGrammar: 828,
	WdDialogToolsCreateDirectory: 833,
	WdDialogTableWrapping: 854,
	WdDialogFormatTheme: 855,
	WdDialogTableProperties: 861,
	WdDialogEmailOptions: 863,
	WdDialogCreateAutoText: 872,
	WdDialogToolsAutoSummarize: 874,
	WdDialogToolsGrammarSettings: 885,
	WdDialogEditGoTo: 896,
	WdDialogWebOptions: 898,
	WdDialogInsertHyperlink: 925,
	WdDialogToolsAutoManager: 915,
	WdDialogFileVersions: 945,
	WdDialogToolsOptionsAutoFormat: 959,
	WdDialogFormatDrawingObject: 960,
	WdDialogToolsOptions: 974,
	WdDialogFitText: 983,
	WdDialogEditAutoText: 985,
	WdDialogPhoneticGuide: 986,
	WdDialogToolsDictionary: 989,
	WdDialogFileSaveVersion: 1007,
	WdDialogToolsOptionsBidi: 1029,
	WdDialogFrameSetProperties: 1074,
	WdDialogTableTableOptions: 1080,
	WdDialogTableCellOptions: 1081,
	WdDialogIMESetDefault: 1094,
	WdDialogTCSCTranslator: 1156,
	WdDialogHorizontalInVertical: 1160,
	WdDialogTwoLinesInOne: 1161,
	WdDialogFormatEncloseCharacters: 1162,
	WdDialogConsistencyChecker: 1121,
	WdDialogToolsOptionsSmartTag: 1395,
	WdDialogFormatStylesCustom: 1248,
	WdDialogCSSLinks: 1261,
	WdDialogInsertWebComponent: 1324,
	WdDialogToolsOptionsEditCopyPaste: 1356,
	WdDialogToolsOptionsSecurity: 1361,
	WdDialogSearch: 1363,
	WdDialogShowRepairs: 1381,
	WdDialogMailMergeInsertAsk: 4047,
	WdDialogMailMergeInsertFillIn: 4048,
	WdDialogMailMergeInsertIf: 4049,
	WdDialogMailMergeInsertNextIf: 4053,
	WdDialogMailMergeInsertSet: 4054,
	WdDialogMailMergeInsertSkipIf: 4055,
	WdDialogMailMergeFieldMapping: 1304,
	WdDialogMailMergeInsertAddressBlock: 1305,
	WdDialogMailMergeInsertGreetingLine: 1306,
	WdDialogMailMergeInsertFields: 1307,
	WdDialogMailMergeRecipients: 1308,
	WdDialogMailMergeFindRecipient: 1326,
	WdDialogMailMergeSetDocumentType: 1339,
	WdDialogLabelOptions: 1367,
	WdDialogXMLElementAttributes: 1460,
	WdDialogSchemaLibrary: 1417,
	WdDialogPermission: 1469,
	WdDialogMyPermission: 1437,
	WdDialogXMLOptions: 1425,
	WdDialogFormattingRestrictions: 1427,
	WdDialogSourceManager: 1920,
	WdDialogCreateSource: 1922,
	WdDialogDocumentInspector: 1482,
	WdDialogStyleManagement: 1948,
	WdDialogInsertSource: 2120,
	WdDialogOMathRecognizedFunctions: 2165,
	WdDialogInsertPlaceholder: 2348,
	WdDialogBuildingBlockOrganizer: 2067,
	WdDialogContentControlProperties: 2394,
	WdDialogCompatibilityChecker: 2439,
	WdDialogExportAsFixedFormat: 2349,
	WdDialogFileNew2007: 1116,
}

// enum WdWordDialogHID
var WdWordDialogHID = struct {
	Emptyenum int32
}{
	Emptyenum: 0,
}

// enum WdFieldKind
var WdFieldKind = struct {
	WdFieldKindNone int32
	WdFieldKindHot int32
	WdFieldKindWarm int32
	WdFieldKindCold int32
}{
	WdFieldKindNone: 0,
	WdFieldKindHot: 1,
	WdFieldKindWarm: 2,
	WdFieldKindCold: 3,
}

// enum WdTextFormFieldType
var WdTextFormFieldType = struct {
	WdRegularText int32
	WdNumberText int32
	WdDateText int32
	WdCurrentDateText int32
	WdCurrentTimeText int32
	WdCalculationText int32
}{
	WdRegularText: 0,
	WdNumberText: 1,
	WdDateText: 2,
	WdCurrentDateText: 3,
	WdCurrentTimeText: 4,
	WdCalculationText: 5,
}

// enum WdChevronConvertRule
var WdChevronConvertRule = struct {
	WdNeverConvert int32
	WdAlwaysConvert int32
	WdAskToNotConvert int32
	WdAskToConvert int32
}{
	WdNeverConvert: 0,
	WdAlwaysConvert: 1,
	WdAskToNotConvert: 2,
	WdAskToConvert: 3,
}

// enum WdMailMergeMainDocType
var WdMailMergeMainDocType = struct {
	WdNotAMergeDocument int32
	WdFormLetters int32
	WdMailingLabels int32
	WdEnvelopes int32
	WdCatalog int32
	WdEMail int32
	WdFax int32
	WdDirectory int32
}{
	WdNotAMergeDocument: -1,
	WdFormLetters: 0,
	WdMailingLabels: 1,
	WdEnvelopes: 2,
	WdCatalog: 3,
	WdEMail: 4,
	WdFax: 5,
	WdDirectory: 3,
}

// enum WdMailMergeState
var WdMailMergeState = struct {
	WdNormalDocument int32
	WdMainDocumentOnly int32
	WdMainAndDataSource int32
	WdMainAndHeader int32
	WdMainAndSourceAndHeader int32
	WdDataSource int32
}{
	WdNormalDocument: 0,
	WdMainDocumentOnly: 1,
	WdMainAndDataSource: 2,
	WdMainAndHeader: 3,
	WdMainAndSourceAndHeader: 4,
	WdDataSource: 5,
}

// enum WdMailMergeDestination
var WdMailMergeDestination = struct {
	WdSendToNewDocument int32
	WdSendToPrinter int32
	WdSendToEmail int32
	WdSendToFax int32
}{
	WdSendToNewDocument: 0,
	WdSendToPrinter: 1,
	WdSendToEmail: 2,
	WdSendToFax: 3,
}

// enum WdMailMergeActiveRecord
var WdMailMergeActiveRecord = struct {
	WdNoActiveRecord int32
	WdNextRecord int32
	WdPreviousRecord int32
	WdFirstRecord int32
	WdLastRecord int32
	WdFirstDataSourceRecord int32
	WdLastDataSourceRecord int32
	WdNextDataSourceRecord int32
	WdPreviousDataSourceRecord int32
}{
	WdNoActiveRecord: -1,
	WdNextRecord: -2,
	WdPreviousRecord: -3,
	WdFirstRecord: -4,
	WdLastRecord: -5,
	WdFirstDataSourceRecord: -6,
	WdLastDataSourceRecord: -7,
	WdNextDataSourceRecord: -8,
	WdPreviousDataSourceRecord: -9,
}

// enum WdMailMergeDefaultRecord
var WdMailMergeDefaultRecord = struct {
	WdDefaultFirstRecord int32
	WdDefaultLastRecord int32
}{
	WdDefaultFirstRecord: 1,
	WdDefaultLastRecord: -16,
}

// enum WdMailMergeDataSource
var WdMailMergeDataSource = struct {
	WdNoMergeInfo int32
	WdMergeInfoFromWord int32
	WdMergeInfoFromAccessDDE int32
	WdMergeInfoFromExcelDDE int32
	WdMergeInfoFromMSQueryDDE int32
	WdMergeInfoFromODBC int32
	WdMergeInfoFromODSO int32
}{
	WdNoMergeInfo: -1,
	WdMergeInfoFromWord: 0,
	WdMergeInfoFromAccessDDE: 1,
	WdMergeInfoFromExcelDDE: 2,
	WdMergeInfoFromMSQueryDDE: 3,
	WdMergeInfoFromODBC: 4,
	WdMergeInfoFromODSO: 5,
}

// enum WdMailMergeComparison
var WdMailMergeComparison = struct {
	WdMergeIfEqual int32
	WdMergeIfNotEqual int32
	WdMergeIfLessThan int32
	WdMergeIfGreaterThan int32
	WdMergeIfLessThanOrEqual int32
	WdMergeIfGreaterThanOrEqual int32
	WdMergeIfIsBlank int32
	WdMergeIfIsNotBlank int32
}{
	WdMergeIfEqual: 0,
	WdMergeIfNotEqual: 1,
	WdMergeIfLessThan: 2,
	WdMergeIfGreaterThan: 3,
	WdMergeIfLessThanOrEqual: 4,
	WdMergeIfGreaterThanOrEqual: 5,
	WdMergeIfIsBlank: 6,
	WdMergeIfIsNotBlank: 7,
}

// enum WdBookmarkSortBy
var WdBookmarkSortBy = struct {
	WdSortByName int32
	WdSortByLocation int32
}{
	WdSortByName: 0,
	WdSortByLocation: 1,
}

// enum WdWindowState
var WdWindowState = struct {
	WdWindowStateNormal int32
	WdWindowStateMaximize int32
	WdWindowStateMinimize int32
}{
	WdWindowStateNormal: 0,
	WdWindowStateMaximize: 1,
	WdWindowStateMinimize: 2,
}

// enum WdPictureLinkType
var WdPictureLinkType = struct {
	WdLinkNone int32
	WdLinkDataInDoc int32
	WdLinkDataOnDisk int32
}{
	WdLinkNone: 0,
	WdLinkDataInDoc: 1,
	WdLinkDataOnDisk: 2,
}

// enum WdLinkType
var WdLinkType = struct {
	WdLinkTypeOLE int32
	WdLinkTypePicture int32
	WdLinkTypeText int32
	WdLinkTypeReference int32
	WdLinkTypeInclude int32
	WdLinkTypeImport int32
	WdLinkTypeDDE int32
	WdLinkTypeDDEAuto int32
	WdLinkTypeChart int32
}{
	WdLinkTypeOLE: 0,
	WdLinkTypePicture: 1,
	WdLinkTypeText: 2,
	WdLinkTypeReference: 3,
	WdLinkTypeInclude: 4,
	WdLinkTypeImport: 5,
	WdLinkTypeDDE: 6,
	WdLinkTypeDDEAuto: 7,
	WdLinkTypeChart: 8,
}

// enum WdWindowType
var WdWindowType = struct {
	WdWindowDocument int32
	WdWindowTemplate int32
}{
	WdWindowDocument: 0,
	WdWindowTemplate: 1,
}

// enum WdViewType
var WdViewType = struct {
	WdNormalView int32
	WdOutlineView int32
	WdPrintView int32
	WdPrintPreview int32
	WdMasterView int32
	WdWebView int32
	WdReadingView int32
	WdConflictView int32
}{
	WdNormalView: 1,
	WdOutlineView: 2,
	WdPrintView: 3,
	WdPrintPreview: 4,
	WdMasterView: 5,
	WdWebView: 6,
	WdReadingView: 7,
	WdConflictView: 8,
}

// enum WdSeekView
var WdSeekView = struct {
	WdSeekMainDocument int32
	WdSeekPrimaryHeader int32
	WdSeekFirstPageHeader int32
	WdSeekEvenPagesHeader int32
	WdSeekPrimaryFooter int32
	WdSeekFirstPageFooter int32
	WdSeekEvenPagesFooter int32
	WdSeekFootnotes int32
	WdSeekEndnotes int32
	WdSeekCurrentPageHeader int32
	WdSeekCurrentPageFooter int32
}{
	WdSeekMainDocument: 0,
	WdSeekPrimaryHeader: 1,
	WdSeekFirstPageHeader: 2,
	WdSeekEvenPagesHeader: 3,
	WdSeekPrimaryFooter: 4,
	WdSeekFirstPageFooter: 5,
	WdSeekEvenPagesFooter: 6,
	WdSeekFootnotes: 7,
	WdSeekEndnotes: 8,
	WdSeekCurrentPageHeader: 9,
	WdSeekCurrentPageFooter: 10,
}

// enum WdSpecialPane
var WdSpecialPane = struct {
	WdPaneNone int32
	WdPanePrimaryHeader int32
	WdPaneFirstPageHeader int32
	WdPaneEvenPagesHeader int32
	WdPanePrimaryFooter int32
	WdPaneFirstPageFooter int32
	WdPaneEvenPagesFooter int32
	WdPaneFootnotes int32
	WdPaneEndnotes int32
	WdPaneFootnoteContinuationNotice int32
	WdPaneFootnoteContinuationSeparator int32
	WdPaneFootnoteSeparator int32
	WdPaneEndnoteContinuationNotice int32
	WdPaneEndnoteContinuationSeparator int32
	WdPaneEndnoteSeparator int32
	WdPaneComments int32
	WdPaneCurrentPageHeader int32
	WdPaneCurrentPageFooter int32
	WdPaneRevisions int32
	WdPaneRevisionsHoriz int32
	WdPaneRevisionsVert int32
}{
	WdPaneNone: 0,
	WdPanePrimaryHeader: 1,
	WdPaneFirstPageHeader: 2,
	WdPaneEvenPagesHeader: 3,
	WdPanePrimaryFooter: 4,
	WdPaneFirstPageFooter: 5,
	WdPaneEvenPagesFooter: 6,
	WdPaneFootnotes: 7,
	WdPaneEndnotes: 8,
	WdPaneFootnoteContinuationNotice: 9,
	WdPaneFootnoteContinuationSeparator: 10,
	WdPaneFootnoteSeparator: 11,
	WdPaneEndnoteContinuationNotice: 12,
	WdPaneEndnoteContinuationSeparator: 13,
	WdPaneEndnoteSeparator: 14,
	WdPaneComments: 15,
	WdPaneCurrentPageHeader: 16,
	WdPaneCurrentPageFooter: 17,
	WdPaneRevisions: 18,
	WdPaneRevisionsHoriz: 19,
	WdPaneRevisionsVert: 20,
}

// enum WdPageFit
var WdPageFit = struct {
	WdPageFitNone int32
	WdPageFitFullPage int32
	WdPageFitBestFit int32
	WdPageFitTextFit int32
}{
	WdPageFitNone: 0,
	WdPageFitFullPage: 1,
	WdPageFitBestFit: 2,
	WdPageFitTextFit: 3,
}

// enum WdBrowseTarget
var WdBrowseTarget = struct {
	WdBrowsePage int32
	WdBrowseSection int32
	WdBrowseComment int32
	WdBrowseFootnote int32
	WdBrowseEndnote int32
	WdBrowseField int32
	WdBrowseTable int32
	WdBrowseGraphic int32
	WdBrowseHeading int32
	WdBrowseEdit int32
	WdBrowseFind int32
	WdBrowseGoTo int32
}{
	WdBrowsePage: 1,
	WdBrowseSection: 2,
	WdBrowseComment: 3,
	WdBrowseFootnote: 4,
	WdBrowseEndnote: 5,
	WdBrowseField: 6,
	WdBrowseTable: 7,
	WdBrowseGraphic: 8,
	WdBrowseHeading: 9,
	WdBrowseEdit: 10,
	WdBrowseFind: 11,
	WdBrowseGoTo: 12,
}

// enum WdPaperTray
var WdPaperTray = struct {
	WdPrinterDefaultBin int32
	WdPrinterUpperBin int32
	WdPrinterOnlyBin int32
	WdPrinterLowerBin int32
	WdPrinterMiddleBin int32
	WdPrinterManualFeed int32
	WdPrinterEnvelopeFeed int32
	WdPrinterManualEnvelopeFeed int32
	WdPrinterAutomaticSheetFeed int32
	WdPrinterTractorFeed int32
	WdPrinterSmallFormatBin int32
	WdPrinterLargeFormatBin int32
	WdPrinterLargeCapacityBin int32
	WdPrinterPaperCassette int32
	WdPrinterFormSource int32
}{
	WdPrinterDefaultBin: 0,
	WdPrinterUpperBin: 1,
	WdPrinterOnlyBin: 1,
	WdPrinterLowerBin: 2,
	WdPrinterMiddleBin: 3,
	WdPrinterManualFeed: 4,
	WdPrinterEnvelopeFeed: 5,
	WdPrinterManualEnvelopeFeed: 6,
	WdPrinterAutomaticSheetFeed: 7,
	WdPrinterTractorFeed: 8,
	WdPrinterSmallFormatBin: 9,
	WdPrinterLargeFormatBin: 10,
	WdPrinterLargeCapacityBin: 11,
	WdPrinterPaperCassette: 14,
	WdPrinterFormSource: 15,
}

// enum WdOrientation
var WdOrientation = struct {
	WdOrientPortrait int32
	WdOrientLandscape int32
}{
	WdOrientPortrait: 0,
	WdOrientLandscape: 1,
}

// enum WdSelectionType
var WdSelectionType = struct {
	WdNoSelection int32
	WdSelectionIP int32
	WdSelectionNormal int32
	WdSelectionFrame int32
	WdSelectionColumn int32
	WdSelectionRow int32
	WdSelectionBlock int32
	WdSelectionInlineShape int32
	WdSelectionShape int32
}{
	WdNoSelection: 0,
	WdSelectionIP: 1,
	WdSelectionNormal: 2,
	WdSelectionFrame: 3,
	WdSelectionColumn: 4,
	WdSelectionRow: 5,
	WdSelectionBlock: 6,
	WdSelectionInlineShape: 7,
	WdSelectionShape: 8,
}

// enum WdCaptionLabelID
var WdCaptionLabelID = struct {
	WdCaptionFigure int32
	WdCaptionTable int32
	WdCaptionEquation int32
}{
	WdCaptionFigure: -1,
	WdCaptionTable: -2,
	WdCaptionEquation: -3,
}

// enum WdReferenceType
var WdReferenceType = struct {
	WdRefTypeNumberedItem int32
	WdRefTypeHeading int32
	WdRefTypeBookmark int32
	WdRefTypeFootnote int32
	WdRefTypeEndnote int32
}{
	WdRefTypeNumberedItem: 0,
	WdRefTypeHeading: 1,
	WdRefTypeBookmark: 2,
	WdRefTypeFootnote: 3,
	WdRefTypeEndnote: 4,
}

// enum WdReferenceKind
var WdReferenceKind = struct {
	WdContentText int32
	WdNumberRelativeContext int32
	WdNumberNoContext int32
	WdNumberFullContext int32
	WdEntireCaption int32
	WdOnlyLabelAndNumber int32
	WdOnlyCaptionText int32
	WdFootnoteNumber int32
	WdEndnoteNumber int32
	WdPageNumber int32
	WdPosition int32
	WdFootnoteNumberFormatted int32
	WdEndnoteNumberFormatted int32
}{
	WdContentText: -1,
	WdNumberRelativeContext: -2,
	WdNumberNoContext: -3,
	WdNumberFullContext: -4,
	WdEntireCaption: 2,
	WdOnlyLabelAndNumber: 3,
	WdOnlyCaptionText: 4,
	WdFootnoteNumber: 5,
	WdEndnoteNumber: 6,
	WdPageNumber: 7,
	WdPosition: 15,
	WdFootnoteNumberFormatted: 16,
	WdEndnoteNumberFormatted: 17,
}

// enum WdIndexFormat
var WdIndexFormat = struct {
	WdIndexTemplate int32
	WdIndexClassic int32
	WdIndexFancy int32
	WdIndexModern int32
	WdIndexBulleted int32
	WdIndexFormal int32
	WdIndexSimple int32
}{
	WdIndexTemplate: 0,
	WdIndexClassic: 1,
	WdIndexFancy: 2,
	WdIndexModern: 3,
	WdIndexBulleted: 4,
	WdIndexFormal: 5,
	WdIndexSimple: 6,
}

// enum WdIndexType
var WdIndexType = struct {
	WdIndexIndent int32
	WdIndexRunin int32
}{
	WdIndexIndent: 0,
	WdIndexRunin: 1,
}

// enum WdRevisionsWrap
var WdRevisionsWrap = struct {
	WdWrapNever int32
	WdWrapAlways int32
	WdWrapAsk int32
}{
	WdWrapNever: 0,
	WdWrapAlways: 1,
	WdWrapAsk: 2,
}

// enum WdRevisionType
var WdRevisionType = struct {
	WdNoRevision int32
	WdRevisionInsert int32
	WdRevisionDelete int32
	WdRevisionProperty int32
	WdRevisionParagraphNumber int32
	WdRevisionDisplayField int32
	WdRevisionReconcile int32
	WdRevisionConflict int32
	WdRevisionStyle int32
	WdRevisionReplace int32
	WdRevisionParagraphProperty int32
	WdRevisionTableProperty int32
	WdRevisionSectionProperty int32
	WdRevisionStyleDefinition int32
	WdRevisionMovedFrom int32
	WdRevisionMovedTo int32
	WdRevisionCellInsertion int32
	WdRevisionCellDeletion int32
	WdRevisionCellMerge int32
	WdRevisionCellSplit int32
	WdRevisionConflictInsert int32
	WdRevisionConflictDelete int32
}{
	WdNoRevision: 0,
	WdRevisionInsert: 1,
	WdRevisionDelete: 2,
	WdRevisionProperty: 3,
	WdRevisionParagraphNumber: 4,
	WdRevisionDisplayField: 5,
	WdRevisionReconcile: 6,
	WdRevisionConflict: 7,
	WdRevisionStyle: 8,
	WdRevisionReplace: 9,
	WdRevisionParagraphProperty: 10,
	WdRevisionTableProperty: 11,
	WdRevisionSectionProperty: 12,
	WdRevisionStyleDefinition: 13,
	WdRevisionMovedFrom: 14,
	WdRevisionMovedTo: 15,
	WdRevisionCellInsertion: 16,
	WdRevisionCellDeletion: 17,
	WdRevisionCellMerge: 18,
	WdRevisionCellSplit: 19,
	WdRevisionConflictInsert: 20,
	WdRevisionConflictDelete: 21,
}

// enum WdRoutingSlipDelivery
var WdRoutingSlipDelivery = struct {
	WdOneAfterAnother int32
	WdAllAtOnce int32
}{
	WdOneAfterAnother: 0,
	WdAllAtOnce: 1,
}

// enum WdRoutingSlipStatus
var WdRoutingSlipStatus = struct {
	WdNotYetRouted int32
	WdRouteInProgress int32
	WdRouteComplete int32
}{
	WdNotYetRouted: 0,
	WdRouteInProgress: 1,
	WdRouteComplete: 2,
}

// enum WdSectionStart
var WdSectionStart = struct {
	WdSectionContinuous int32
	WdSectionNewColumn int32
	WdSectionNewPage int32
	WdSectionEvenPage int32
	WdSectionOddPage int32
}{
	WdSectionContinuous: 0,
	WdSectionNewColumn: 1,
	WdSectionNewPage: 2,
	WdSectionEvenPage: 3,
	WdSectionOddPage: 4,
}

// enum WdSaveOptions
var WdSaveOptions = struct {
	WdDoNotSaveChanges int32
	WdSaveChanges int32
	WdPromptToSaveChanges int32
}{
	WdDoNotSaveChanges: 0,
	WdSaveChanges: -1,
	WdPromptToSaveChanges: -2,
}

// enum WdDocumentKind
var WdDocumentKind = struct {
	WdDocumentNotSpecified int32
	WdDocumentLetter int32
	WdDocumentEmail int32
}{
	WdDocumentNotSpecified: 0,
	WdDocumentLetter: 1,
	WdDocumentEmail: 2,
}

// enum WdDocumentType
var WdDocumentType = struct {
	WdTypeDocument int32
	WdTypeTemplate int32
	WdTypeFrameset int32
}{
	WdTypeDocument: 0,
	WdTypeTemplate: 1,
	WdTypeFrameset: 2,
}

// enum WdOriginalFormat
var WdOriginalFormat = struct {
	WdWordDocument int32
	WdOriginalDocumentFormat int32
	WdPromptUser int32
}{
	WdWordDocument: 0,
	WdOriginalDocumentFormat: 1,
	WdPromptUser: 2,
}

// enum WdRelocate
var WdRelocate = struct {
	WdRelocateUp int32
	WdRelocateDown int32
}{
	WdRelocateUp: 0,
	WdRelocateDown: 1,
}

// enum WdInsertedTextMark
var WdInsertedTextMark = struct {
	WdInsertedTextMarkNone int32
	WdInsertedTextMarkBold int32
	WdInsertedTextMarkItalic int32
	WdInsertedTextMarkUnderline int32
	WdInsertedTextMarkDoubleUnderline int32
	WdInsertedTextMarkColorOnly int32
	WdInsertedTextMarkStrikeThrough int32
	WdInsertedTextMarkDoubleStrikeThrough int32
}{
	WdInsertedTextMarkNone: 0,
	WdInsertedTextMarkBold: 1,
	WdInsertedTextMarkItalic: 2,
	WdInsertedTextMarkUnderline: 3,
	WdInsertedTextMarkDoubleUnderline: 4,
	WdInsertedTextMarkColorOnly: 5,
	WdInsertedTextMarkStrikeThrough: 6,
	WdInsertedTextMarkDoubleStrikeThrough: 7,
}

// enum WdRevisedLinesMark
var WdRevisedLinesMark = struct {
	WdRevisedLinesMarkNone int32
	WdRevisedLinesMarkLeftBorder int32
	WdRevisedLinesMarkRightBorder int32
	WdRevisedLinesMarkOutsideBorder int32
}{
	WdRevisedLinesMarkNone: 0,
	WdRevisedLinesMarkLeftBorder: 1,
	WdRevisedLinesMarkRightBorder: 2,
	WdRevisedLinesMarkOutsideBorder: 3,
}

// enum WdDeletedTextMark
var WdDeletedTextMark = struct {
	WdDeletedTextMarkHidden int32
	WdDeletedTextMarkStrikeThrough int32
	WdDeletedTextMarkCaret int32
	WdDeletedTextMarkPound int32
	WdDeletedTextMarkNone int32
	WdDeletedTextMarkBold int32
	WdDeletedTextMarkItalic int32
	WdDeletedTextMarkUnderline int32
	WdDeletedTextMarkDoubleUnderline int32
	WdDeletedTextMarkColorOnly int32
	WdDeletedTextMarkDoubleStrikeThrough int32
}{
	WdDeletedTextMarkHidden: 0,
	WdDeletedTextMarkStrikeThrough: 1,
	WdDeletedTextMarkCaret: 2,
	WdDeletedTextMarkPound: 3,
	WdDeletedTextMarkNone: 4,
	WdDeletedTextMarkBold: 5,
	WdDeletedTextMarkItalic: 6,
	WdDeletedTextMarkUnderline: 7,
	WdDeletedTextMarkDoubleUnderline: 8,
	WdDeletedTextMarkColorOnly: 9,
	WdDeletedTextMarkDoubleStrikeThrough: 10,
}

// enum WdRevisedPropertiesMark
var WdRevisedPropertiesMark = struct {
	WdRevisedPropertiesMarkNone int32
	WdRevisedPropertiesMarkBold int32
	WdRevisedPropertiesMarkItalic int32
	WdRevisedPropertiesMarkUnderline int32
	WdRevisedPropertiesMarkDoubleUnderline int32
	WdRevisedPropertiesMarkColorOnly int32
	WdRevisedPropertiesMarkStrikeThrough int32
	WdRevisedPropertiesMarkDoubleStrikeThrough int32
}{
	WdRevisedPropertiesMarkNone: 0,
	WdRevisedPropertiesMarkBold: 1,
	WdRevisedPropertiesMarkItalic: 2,
	WdRevisedPropertiesMarkUnderline: 3,
	WdRevisedPropertiesMarkDoubleUnderline: 4,
	WdRevisedPropertiesMarkColorOnly: 5,
	WdRevisedPropertiesMarkStrikeThrough: 6,
	WdRevisedPropertiesMarkDoubleStrikeThrough: 7,
}

// enum WdFieldShading
var WdFieldShading = struct {
	WdFieldShadingNever int32
	WdFieldShadingAlways int32
	WdFieldShadingWhenSelected int32
}{
	WdFieldShadingNever: 0,
	WdFieldShadingAlways: 1,
	WdFieldShadingWhenSelected: 2,
}

// enum WdDefaultFilePath
var WdDefaultFilePath = struct {
	WdDocumentsPath int32
	WdPicturesPath int32
	WdUserTemplatesPath int32
	WdWorkgroupTemplatesPath int32
	WdUserOptionsPath int32
	WdAutoRecoverPath int32
	WdToolsPath int32
	WdTutorialPath int32
	WdStartupPath int32
	WdProgramPath int32
	WdGraphicsFiltersPath int32
	WdTextConvertersPath int32
	WdProofingToolsPath int32
	WdTempFilePath int32
	WdCurrentFolderPath int32
	WdStyleGalleryPath int32
	WdBorderArtPath int32
}{
	WdDocumentsPath: 0,
	WdPicturesPath: 1,
	WdUserTemplatesPath: 2,
	WdWorkgroupTemplatesPath: 3,
	WdUserOptionsPath: 4,
	WdAutoRecoverPath: 5,
	WdToolsPath: 6,
	WdTutorialPath: 7,
	WdStartupPath: 8,
	WdProgramPath: 9,
	WdGraphicsFiltersPath: 10,
	WdTextConvertersPath: 11,
	WdProofingToolsPath: 12,
	WdTempFilePath: 13,
	WdCurrentFolderPath: 14,
	WdStyleGalleryPath: 15,
	WdBorderArtPath: 19,
}

// enum WdCompatibility
var WdCompatibility = struct {
	WdNoTabHangIndent int32
	WdNoSpaceRaiseLower int32
	WdPrintColBlack int32
	WdWrapTrailSpaces int32
	WdNoColumnBalance int32
	WdConvMailMergeEsc int32
	WdSuppressSpBfAfterPgBrk int32
	WdSuppressTopSpacing int32
	WdOrigWordTableRules int32
	WdTransparentMetafiles int32
	WdShowBreaksInFrames int32
	WdSwapBordersFacingPages int32
	WdLeaveBackslashAlone int32
	WdExpandShiftReturn int32
	WdDontULTrailSpace int32
	WdDontBalanceSingleByteDoubleByteWidth int32
	WdSuppressTopSpacingMac5 int32
	WdSpacingInWholePoints int32
	WdPrintBodyTextBeforeHeader int32
	WdNoLeading int32
	WdNoSpaceForUL int32
	WdMWSmallCaps int32
	WdNoExtraLineSpacing int32
	WdTruncateFontHeight int32
	WdSubFontBySize int32
	WdUsePrinterMetrics int32
	WdWW6BorderRules int32
	WdExactOnTop int32
	WdSuppressBottomSpacing int32
	WdWPSpaceWidth int32
	WdWPJustification int32
	WdLineWrapLikeWord6 int32
	WdShapeLayoutLikeWW8 int32
	WdFootnoteLayoutLikeWW8 int32
	WdDontUseHTMLParagraphAutoSpacing int32
	WdDontAdjustLineHeightInTable int32
	WdForgetLastTabAlignment int32
	WdAutospaceLikeWW7 int32
	WdAlignTablesRowByRow int32
	WdLayoutRawTableWidth int32
	WdLayoutTableRowsApart int32
	WdUseWord97LineBreakingRules int32
	WdDontBreakWrappedTables int32
	WdDontSnapTextToGridInTableWithObjects int32
	WdSelectFieldWithFirstOrLastCharacter int32
	WdApplyBreakingRules int32
	WdDontWrapTextWithPunctuation int32
	WdDontUseAsianBreakRulesInGrid int32
	WdUseWord2002TableStyleRules int32
	WdGrowAutofit int32
	WdUseNormalStyleForList int32
	WdDontUseIndentAsNumberingTabStop int32
	WdFELineBreak11 int32
	WdAllowSpaceOfSameStyleInTable int32
	WdWW11IndentRules int32
	WdDontAutofitConstrainedTables int32
	WdAutofitLikeWW11 int32
	WdUnderlineTabInNumList int32
	WdHangulWidthLikeWW11 int32
	WdSplitPgBreakAndParaMark int32
	WdDontVertAlignCellWithShape int32
	WdDontBreakConstrainedForcedTables int32
	WdDontVertAlignInTextbox int32
	WdWord11KerningPairs int32
	WdCachedColBalance int32
	WdDisableOTKerning int32
	WdFlipMirrorIndents int32
	WdDontOverrideTableStyleFontSzAndJustification int32
}{
	WdNoTabHangIndent: 1,
	WdNoSpaceRaiseLower: 2,
	WdPrintColBlack: 3,
	WdWrapTrailSpaces: 4,
	WdNoColumnBalance: 5,
	WdConvMailMergeEsc: 6,
	WdSuppressSpBfAfterPgBrk: 7,
	WdSuppressTopSpacing: 8,
	WdOrigWordTableRules: 9,
	WdTransparentMetafiles: 10,
	WdShowBreaksInFrames: 11,
	WdSwapBordersFacingPages: 12,
	WdLeaveBackslashAlone: 13,
	WdExpandShiftReturn: 14,
	WdDontULTrailSpace: 15,
	WdDontBalanceSingleByteDoubleByteWidth: 16,
	WdSuppressTopSpacingMac5: 17,
	WdSpacingInWholePoints: 18,
	WdPrintBodyTextBeforeHeader: 19,
	WdNoLeading: 20,
	WdNoSpaceForUL: 21,
	WdMWSmallCaps: 22,
	WdNoExtraLineSpacing: 23,
	WdTruncateFontHeight: 24,
	WdSubFontBySize: 25,
	WdUsePrinterMetrics: 26,
	WdWW6BorderRules: 27,
	WdExactOnTop: 28,
	WdSuppressBottomSpacing: 29,
	WdWPSpaceWidth: 30,
	WdWPJustification: 31,
	WdLineWrapLikeWord6: 32,
	WdShapeLayoutLikeWW8: 33,
	WdFootnoteLayoutLikeWW8: 34,
	WdDontUseHTMLParagraphAutoSpacing: 35,
	WdDontAdjustLineHeightInTable: 36,
	WdForgetLastTabAlignment: 37,
	WdAutospaceLikeWW7: 38,
	WdAlignTablesRowByRow: 39,
	WdLayoutRawTableWidth: 40,
	WdLayoutTableRowsApart: 41,
	WdUseWord97LineBreakingRules: 42,
	WdDontBreakWrappedTables: 43,
	WdDontSnapTextToGridInTableWithObjects: 44,
	WdSelectFieldWithFirstOrLastCharacter: 45,
	WdApplyBreakingRules: 46,
	WdDontWrapTextWithPunctuation: 47,
	WdDontUseAsianBreakRulesInGrid: 48,
	WdUseWord2002TableStyleRules: 49,
	WdGrowAutofit: 50,
	WdUseNormalStyleForList: 51,
	WdDontUseIndentAsNumberingTabStop: 52,
	WdFELineBreak11: 53,
	WdAllowSpaceOfSameStyleInTable: 54,
	WdWW11IndentRules: 55,
	WdDontAutofitConstrainedTables: 56,
	WdAutofitLikeWW11: 57,
	WdUnderlineTabInNumList: 58,
	WdHangulWidthLikeWW11: 59,
	WdSplitPgBreakAndParaMark: 60,
	WdDontVertAlignCellWithShape: 61,
	WdDontBreakConstrainedForcedTables: 62,
	WdDontVertAlignInTextbox: 63,
	WdWord11KerningPairs: 64,
	WdCachedColBalance: 65,
	WdDisableOTKerning: 66,
	WdFlipMirrorIndents: 67,
	WdDontOverrideTableStyleFontSzAndJustification: 68,
}

// enum WdPaperSize
var WdPaperSize = struct {
	WdPaper10x14 int32
	WdPaper11x17 int32
	WdPaperLetter int32
	WdPaperLetterSmall int32
	WdPaperLegal int32
	WdPaperExecutive int32
	WdPaperA3 int32
	WdPaperA4 int32
	WdPaperA4Small int32
	WdPaperA5 int32
	WdPaperB4 int32
	WdPaperB5 int32
	WdPaperCSheet int32
	WdPaperDSheet int32
	WdPaperESheet int32
	WdPaperFanfoldLegalGerman int32
	WdPaperFanfoldStdGerman int32
	WdPaperFanfoldUS int32
	WdPaperFolio int32
	WdPaperLedger int32
	WdPaperNote int32
	WdPaperQuarto int32
	WdPaperStatement int32
	WdPaperTabloid int32
	WdPaperEnvelope9 int32
	WdPaperEnvelope10 int32
	WdPaperEnvelope11 int32
	WdPaperEnvelope12 int32
	WdPaperEnvelope14 int32
	WdPaperEnvelopeB4 int32
	WdPaperEnvelopeB5 int32
	WdPaperEnvelopeB6 int32
	WdPaperEnvelopeC3 int32
	WdPaperEnvelopeC4 int32
	WdPaperEnvelopeC5 int32
	WdPaperEnvelopeC6 int32
	WdPaperEnvelopeC65 int32
	WdPaperEnvelopeDL int32
	WdPaperEnvelopeItaly int32
	WdPaperEnvelopeMonarch int32
	WdPaperEnvelopePersonal int32
	WdPaperCustom int32
}{
	WdPaper10x14: 0,
	WdPaper11x17: 1,
	WdPaperLetter: 2,
	WdPaperLetterSmall: 3,
	WdPaperLegal: 4,
	WdPaperExecutive: 5,
	WdPaperA3: 6,
	WdPaperA4: 7,
	WdPaperA4Small: 8,
	WdPaperA5: 9,
	WdPaperB4: 10,
	WdPaperB5: 11,
	WdPaperCSheet: 12,
	WdPaperDSheet: 13,
	WdPaperESheet: 14,
	WdPaperFanfoldLegalGerman: 15,
	WdPaperFanfoldStdGerman: 16,
	WdPaperFanfoldUS: 17,
	WdPaperFolio: 18,
	WdPaperLedger: 19,
	WdPaperNote: 20,
	WdPaperQuarto: 21,
	WdPaperStatement: 22,
	WdPaperTabloid: 23,
	WdPaperEnvelope9: 24,
	WdPaperEnvelope10: 25,
	WdPaperEnvelope11: 26,
	WdPaperEnvelope12: 27,
	WdPaperEnvelope14: 28,
	WdPaperEnvelopeB4: 29,
	WdPaperEnvelopeB5: 30,
	WdPaperEnvelopeB6: 31,
	WdPaperEnvelopeC3: 32,
	WdPaperEnvelopeC4: 33,
	WdPaperEnvelopeC5: 34,
	WdPaperEnvelopeC6: 35,
	WdPaperEnvelopeC65: 36,
	WdPaperEnvelopeDL: 37,
	WdPaperEnvelopeItaly: 38,
	WdPaperEnvelopeMonarch: 39,
	WdPaperEnvelopePersonal: 40,
	WdPaperCustom: 41,
}

// enum WdCustomLabelPageSize
var WdCustomLabelPageSize = struct {
	WdCustomLabelLetter int32
	WdCustomLabelLetterLS int32
	WdCustomLabelA4 int32
	WdCustomLabelA4LS int32
	WdCustomLabelA5 int32
	WdCustomLabelA5LS int32
	WdCustomLabelB5 int32
	WdCustomLabelMini int32
	WdCustomLabelFanfold int32
	WdCustomLabelVertHalfSheet int32
	WdCustomLabelVertHalfSheetLS int32
	WdCustomLabelHigaki int32
	WdCustomLabelHigakiLS int32
	WdCustomLabelB4JIS int32
}{
	WdCustomLabelLetter: 0,
	WdCustomLabelLetterLS: 1,
	WdCustomLabelA4: 2,
	WdCustomLabelA4LS: 3,
	WdCustomLabelA5: 4,
	WdCustomLabelA5LS: 5,
	WdCustomLabelB5: 6,
	WdCustomLabelMini: 7,
	WdCustomLabelFanfold: 8,
	WdCustomLabelVertHalfSheet: 9,
	WdCustomLabelVertHalfSheetLS: 10,
	WdCustomLabelHigaki: 11,
	WdCustomLabelHigakiLS: 12,
	WdCustomLabelB4JIS: 13,
}

// enum WdProtectionType
var WdProtectionType = struct {
	WdNoProtection int32
	WdAllowOnlyRevisions int32
	WdAllowOnlyComments int32
	WdAllowOnlyFormFields int32
	WdAllowOnlyReading int32
}{
	WdNoProtection: -1,
	WdAllowOnlyRevisions: 0,
	WdAllowOnlyComments: 1,
	WdAllowOnlyFormFields: 2,
	WdAllowOnlyReading: 3,
}

// enum WdPartOfSpeech
var WdPartOfSpeech = struct {
	WdAdjective int32
	WdNoun int32
	WdAdverb int32
	WdVerb int32
	WdPronoun int32
	WdConjunction int32
	WdPreposition int32
	WdInterjection int32
	WdIdiom int32
	WdOther int32
}{
	WdAdjective: 0,
	WdNoun: 1,
	WdAdverb: 2,
	WdVerb: 3,
	WdPronoun: 4,
	WdConjunction: 5,
	WdPreposition: 6,
	WdInterjection: 7,
	WdIdiom: 8,
	WdOther: 9,
}

// enum WdSubscriberFormats
var WdSubscriberFormats = struct {
	WdSubscriberBestFormat int32
	WdSubscriberRTF int32
	WdSubscriberText int32
	WdSubscriberPict int32
}{
	WdSubscriberBestFormat: 0,
	WdSubscriberRTF: 1,
	WdSubscriberText: 2,
	WdSubscriberPict: 4,
}

// enum WdEditionType
var WdEditionType = struct {
	WdPublisher int32
	WdSubscriber int32
}{
	WdPublisher: 0,
	WdSubscriber: 1,
}

// enum WdEditionOption
var WdEditionOption = struct {
	WdCancelPublisher int32
	WdSendPublisher int32
	WdSelectPublisher int32
	WdAutomaticUpdate int32
	WdManualUpdate int32
	WdChangeAttributes int32
	WdUpdateSubscriber int32
	WdOpenSource int32
}{
	WdCancelPublisher: 0,
	WdSendPublisher: 1,
	WdSelectPublisher: 2,
	WdAutomaticUpdate: 3,
	WdManualUpdate: 4,
	WdChangeAttributes: 5,
	WdUpdateSubscriber: 6,
	WdOpenSource: 7,
}

// enum WdRelativeHorizontalPosition
var WdRelativeHorizontalPosition = struct {
	WdRelativeHorizontalPositionMargin int32
	WdRelativeHorizontalPositionPage int32
	WdRelativeHorizontalPositionColumn int32
	WdRelativeHorizontalPositionCharacter int32
	WdRelativeHorizontalPositionLeftMarginArea int32
	WdRelativeHorizontalPositionRightMarginArea int32
	WdRelativeHorizontalPositionInnerMarginArea int32
	WdRelativeHorizontalPositionOuterMarginArea int32
}{
	WdRelativeHorizontalPositionMargin: 0,
	WdRelativeHorizontalPositionPage: 1,
	WdRelativeHorizontalPositionColumn: 2,
	WdRelativeHorizontalPositionCharacter: 3,
	WdRelativeHorizontalPositionLeftMarginArea: 4,
	WdRelativeHorizontalPositionRightMarginArea: 5,
	WdRelativeHorizontalPositionInnerMarginArea: 6,
	WdRelativeHorizontalPositionOuterMarginArea: 7,
}

// enum WdRelativeVerticalPosition
var WdRelativeVerticalPosition = struct {
	WdRelativeVerticalPositionMargin int32
	WdRelativeVerticalPositionPage int32
	WdRelativeVerticalPositionParagraph int32
	WdRelativeVerticalPositionLine int32
	WdRelativeVerticalPositionTopMarginArea int32
	WdRelativeVerticalPositionBottomMarginArea int32
	WdRelativeVerticalPositionInnerMarginArea int32
	WdRelativeVerticalPositionOuterMarginArea int32
}{
	WdRelativeVerticalPositionMargin: 0,
	WdRelativeVerticalPositionPage: 1,
	WdRelativeVerticalPositionParagraph: 2,
	WdRelativeVerticalPositionLine: 3,
	WdRelativeVerticalPositionTopMarginArea: 4,
	WdRelativeVerticalPositionBottomMarginArea: 5,
	WdRelativeVerticalPositionInnerMarginArea: 6,
	WdRelativeVerticalPositionOuterMarginArea: 7,
}

// enum WdHelpType
var WdHelpType = struct {
	WdHelp int32
	WdHelpAbout int32
	WdHelpActiveWindow int32
	WdHelpContents int32
	WdHelpExamplesAndDemos int32
	WdHelpIndex int32
	WdHelpKeyboard int32
	WdHelpPSSHelp int32
	WdHelpQuickPreview int32
	WdHelpSearch int32
	WdHelpUsingHelp int32
	WdHelpIchitaro int32
	WdHelpPE2 int32
	WdHelpHWP int32
}{
	WdHelp: 0,
	WdHelpAbout: 1,
	WdHelpActiveWindow: 2,
	WdHelpContents: 3,
	WdHelpExamplesAndDemos: 4,
	WdHelpIndex: 5,
	WdHelpKeyboard: 6,
	WdHelpPSSHelp: 7,
	WdHelpQuickPreview: 8,
	WdHelpSearch: 9,
	WdHelpUsingHelp: 10,
	WdHelpIchitaro: 11,
	WdHelpPE2: 12,
	WdHelpHWP: 13,
}

// enum WdHelpTypeHID
var WdHelpTypeHID = struct {
	Emptyenum int32
}{
	Emptyenum: 0,
}

// enum WdKeyCategory
var WdKeyCategory = struct {
	WdKeyCategoryNil int32
	WdKeyCategoryDisable int32
	WdKeyCategoryCommand int32
	WdKeyCategoryMacro int32
	WdKeyCategoryFont int32
	WdKeyCategoryAutoText int32
	WdKeyCategoryStyle int32
	WdKeyCategorySymbol int32
	WdKeyCategoryPrefix int32
}{
	WdKeyCategoryNil: -1,
	WdKeyCategoryDisable: 0,
	WdKeyCategoryCommand: 1,
	WdKeyCategoryMacro: 2,
	WdKeyCategoryFont: 3,
	WdKeyCategoryAutoText: 4,
	WdKeyCategoryStyle: 5,
	WdKeyCategorySymbol: 6,
	WdKeyCategoryPrefix: 7,
}

// enum WdKey
var WdKey = struct {
	WdNoKey int32
	WdKeyShift int32
	WdKeyControl int32
	WdKeyCommand int32
	WdKeyAlt int32
	WdKeyOption int32
	WdKeyA int32
	WdKeyB int32
	WdKeyC int32
	WdKeyD int32
	WdKeyE int32
	WdKeyF int32
	WdKeyG int32
	WdKeyH int32
	WdKeyI int32
	WdKeyJ int32
	WdKeyK int32
	WdKeyL int32
	WdKeyM int32
	WdKeyN int32
	WdKeyO int32
	WdKeyP int32
	WdKeyQ int32
	WdKeyR int32
	WdKeyS int32
	WdKeyT int32
	WdKeyU int32
	WdKeyV int32
	WdKeyW int32
	WdKeyX int32
	WdKeyY int32
	WdKeyZ int32
	WdKey0 int32
	WdKey1 int32
	WdKey2 int32
	WdKey3 int32
	WdKey4 int32
	WdKey5 int32
	WdKey6 int32
	WdKey7 int32
	WdKey8 int32
	WdKey9 int32
	WdKeyBackspace int32
	WdKeyTab int32
	WdKeyNumeric5Special int32
	WdKeyReturn int32
	WdKeyPause int32
	WdKeyEsc int32
	WdKeySpacebar int32
	WdKeyPageUp int32
	WdKeyPageDown int32
	WdKeyEnd int32
	WdKeyHome int32
	WdKeyInsert int32
	WdKeyDelete int32
	WdKeyNumeric0 int32
	WdKeyNumeric1 int32
	WdKeyNumeric2 int32
	WdKeyNumeric3 int32
	WdKeyNumeric4 int32
	WdKeyNumeric5 int32
	WdKeyNumeric6 int32
	WdKeyNumeric7 int32
	WdKeyNumeric8 int32
	WdKeyNumeric9 int32
	WdKeyNumericMultiply int32
	WdKeyNumericAdd int32
	WdKeyNumericSubtract int32
	WdKeyNumericDecimal int32
	WdKeyNumericDivide int32
	WdKeyF1 int32
	WdKeyF2 int32
	WdKeyF3 int32
	WdKeyF4 int32
	WdKeyF5 int32
	WdKeyF6 int32
	WdKeyF7 int32
	WdKeyF8 int32
	WdKeyF9 int32
	WdKeyF10 int32
	WdKeyF11 int32
	WdKeyF12 int32
	WdKeyF13 int32
	WdKeyF14 int32
	WdKeyF15 int32
	WdKeyF16 int32
	WdKeyScrollLock int32
	WdKeySemiColon int32
	WdKeyEquals int32
	WdKeyComma int32
	WdKeyHyphen int32
	WdKeyPeriod int32
	WdKeySlash int32
	WdKeyBackSingleQuote int32
	WdKeyOpenSquareBrace int32
	WdKeyBackSlash int32
	WdKeyCloseSquareBrace int32
	WdKeySingleQuote int32
}{
	WdNoKey: 255,
	WdKeyShift: 256,
	WdKeyControl: 512,
	WdKeyCommand: 512,
	WdKeyAlt: 1024,
	WdKeyOption: 1024,
	WdKeyA: 65,
	WdKeyB: 66,
	WdKeyC: 67,
	WdKeyD: 68,
	WdKeyE: 69,
	WdKeyF: 70,
	WdKeyG: 71,
	WdKeyH: 72,
	WdKeyI: 73,
	WdKeyJ: 74,
	WdKeyK: 75,
	WdKeyL: 76,
	WdKeyM: 77,
	WdKeyN: 78,
	WdKeyO: 79,
	WdKeyP: 80,
	WdKeyQ: 81,
	WdKeyR: 82,
	WdKeyS: 83,
	WdKeyT: 84,
	WdKeyU: 85,
	WdKeyV: 86,
	WdKeyW: 87,
	WdKeyX: 88,
	WdKeyY: 89,
	WdKeyZ: 90,
	WdKey0: 48,
	WdKey1: 49,
	WdKey2: 50,
	WdKey3: 51,
	WdKey4: 52,
	WdKey5: 53,
	WdKey6: 54,
	WdKey7: 55,
	WdKey8: 56,
	WdKey9: 57,
	WdKeyBackspace: 8,
	WdKeyTab: 9,
	WdKeyNumeric5Special: 12,
	WdKeyReturn: 13,
	WdKeyPause: 19,
	WdKeyEsc: 27,
	WdKeySpacebar: 32,
	WdKeyPageUp: 33,
	WdKeyPageDown: 34,
	WdKeyEnd: 35,
	WdKeyHome: 36,
	WdKeyInsert: 45,
	WdKeyDelete: 46,
	WdKeyNumeric0: 96,
	WdKeyNumeric1: 97,
	WdKeyNumeric2: 98,
	WdKeyNumeric3: 99,
	WdKeyNumeric4: 100,
	WdKeyNumeric5: 101,
	WdKeyNumeric6: 102,
	WdKeyNumeric7: 103,
	WdKeyNumeric8: 104,
	WdKeyNumeric9: 105,
	WdKeyNumericMultiply: 106,
	WdKeyNumericAdd: 107,
	WdKeyNumericSubtract: 109,
	WdKeyNumericDecimal: 110,
	WdKeyNumericDivide: 111,
	WdKeyF1: 112,
	WdKeyF2: 113,
	WdKeyF3: 114,
	WdKeyF4: 115,
	WdKeyF5: 116,
	WdKeyF6: 117,
	WdKeyF7: 118,
	WdKeyF8: 119,
	WdKeyF9: 120,
	WdKeyF10: 121,
	WdKeyF11: 122,
	WdKeyF12: 123,
	WdKeyF13: 124,
	WdKeyF14: 125,
	WdKeyF15: 126,
	WdKeyF16: 127,
	WdKeyScrollLock: 145,
	WdKeySemiColon: 186,
	WdKeyEquals: 187,
	WdKeyComma: 188,
	WdKeyHyphen: 189,
	WdKeyPeriod: 190,
	WdKeySlash: 191,
	WdKeyBackSingleQuote: 192,
	WdKeyOpenSquareBrace: 219,
	WdKeyBackSlash: 220,
	WdKeyCloseSquareBrace: 221,
	WdKeySingleQuote: 222,
}

// enum WdOLEType
var WdOLEType = struct {
	WdOLELink int32
	WdOLEEmbed int32
	WdOLEControl int32
}{
	WdOLELink: 0,
	WdOLEEmbed: 1,
	WdOLEControl: 2,
}

// enum WdOLEVerb
var WdOLEVerb = struct {
	WdOLEVerbPrimary int32
	WdOLEVerbShow int32
	WdOLEVerbOpen int32
	WdOLEVerbHide int32
	WdOLEVerbUIActivate int32
	WdOLEVerbInPlaceActivate int32
	WdOLEVerbDiscardUndoState int32
}{
	WdOLEVerbPrimary: 0,
	WdOLEVerbShow: -1,
	WdOLEVerbOpen: -2,
	WdOLEVerbHide: -3,
	WdOLEVerbUIActivate: -4,
	WdOLEVerbInPlaceActivate: -5,
	WdOLEVerbDiscardUndoState: -6,
}

// enum WdOLEPlacement
var WdOLEPlacement = struct {
	WdInLine int32
	WdFloatOverText int32
}{
	WdInLine: 0,
	WdFloatOverText: 1,
}

// enum WdEnvelopeOrientation
var WdEnvelopeOrientation = struct {
	WdLeftPortrait int32
	WdCenterPortrait int32
	WdRightPortrait int32
	WdLeftLandscape int32
	WdCenterLandscape int32
	WdRightLandscape int32
	WdLeftClockwise int32
	WdCenterClockwise int32
	WdRightClockwise int32
}{
	WdLeftPortrait: 0,
	WdCenterPortrait: 1,
	WdRightPortrait: 2,
	WdLeftLandscape: 3,
	WdCenterLandscape: 4,
	WdRightLandscape: 5,
	WdLeftClockwise: 6,
	WdCenterClockwise: 7,
	WdRightClockwise: 8,
}

// enum WdLetterStyle
var WdLetterStyle = struct {
	WdFullBlock int32
	WdModifiedBlock int32
	WdSemiBlock int32
}{
	WdFullBlock: 0,
	WdModifiedBlock: 1,
	WdSemiBlock: 2,
}

// enum WdLetterheadLocation
var WdLetterheadLocation = struct {
	WdLetterTop int32
	WdLetterBottom int32
	WdLetterLeft int32
	WdLetterRight int32
}{
	WdLetterTop: 0,
	WdLetterBottom: 1,
	WdLetterLeft: 2,
	WdLetterRight: 3,
}

// enum WdSalutationType
var WdSalutationType = struct {
	WdSalutationInformal int32
	WdSalutationFormal int32
	WdSalutationBusiness int32
	WdSalutationOther int32
}{
	WdSalutationInformal: 0,
	WdSalutationFormal: 1,
	WdSalutationBusiness: 2,
	WdSalutationOther: 3,
}

// enum WdSalutationGender
var WdSalutationGender = struct {
	WdGenderFemale int32
	WdGenderMale int32
	WdGenderNeutral int32
	WdGenderUnknown int32
}{
	WdGenderFemale: 0,
	WdGenderMale: 1,
	WdGenderNeutral: 2,
	WdGenderUnknown: 3,
}

// enum WdMovementType
var WdMovementType = struct {
	WdMove int32
	WdExtend int32
}{
	WdMove: 0,
	WdExtend: 1,
}

// enum WdConstants
var WdConstants = struct {
	WdUndefined int32
	WdToggle int32
	WdForward int32
	WdBackward int32
	WdAutoPosition int32
	WdFirst int32
	WdCreatorCode int32
}{
	WdUndefined: 9999999,
	WdToggle: 9999998,
	WdForward: 1073741823,
	WdBackward: -1073741823,
	WdAutoPosition: 0,
	WdFirst: 1,
	WdCreatorCode: 1297307460,
}

// enum WdPasteDataType
var WdPasteDataType = struct {
	WdPasteOLEObject int32
	WdPasteRTF int32
	WdPasteText int32
	WdPasteMetafilePicture int32
	WdPasteBitmap int32
	WdPasteDeviceIndependentBitmap int32
	WdPasteHyperlink int32
	WdPasteShape int32
	WdPasteEnhancedMetafile int32
	WdPasteHTML int32
}{
	WdPasteOLEObject: 0,
	WdPasteRTF: 1,
	WdPasteText: 2,
	WdPasteMetafilePicture: 3,
	WdPasteBitmap: 4,
	WdPasteDeviceIndependentBitmap: 5,
	WdPasteHyperlink: 7,
	WdPasteShape: 8,
	WdPasteEnhancedMetafile: 9,
	WdPasteHTML: 10,
}

// enum WdPrintOutItem
var WdPrintOutItem = struct {
	WdPrintDocumentContent int32
	WdPrintProperties int32
	WdPrintComments int32
	WdPrintMarkup int32
	WdPrintStyles int32
	WdPrintAutoTextEntries int32
	WdPrintKeyAssignments int32
	WdPrintEnvelope int32
	WdPrintDocumentWithMarkup int32
}{
	WdPrintDocumentContent: 0,
	WdPrintProperties: 1,
	WdPrintComments: 2,
	WdPrintMarkup: 2,
	WdPrintStyles: 3,
	WdPrintAutoTextEntries: 4,
	WdPrintKeyAssignments: 5,
	WdPrintEnvelope: 6,
	WdPrintDocumentWithMarkup: 7,
}

// enum WdPrintOutPages
var WdPrintOutPages = struct {
	WdPrintAllPages int32
	WdPrintOddPagesOnly int32
	WdPrintEvenPagesOnly int32
}{
	WdPrintAllPages: 0,
	WdPrintOddPagesOnly: 1,
	WdPrintEvenPagesOnly: 2,
}

// enum WdPrintOutRange
var WdPrintOutRange = struct {
	WdPrintAllDocument int32
	WdPrintSelection int32
	WdPrintCurrentPage int32
	WdPrintFromTo int32
	WdPrintRangeOfPages int32
}{
	WdPrintAllDocument: 0,
	WdPrintSelection: 1,
	WdPrintCurrentPage: 2,
	WdPrintFromTo: 3,
	WdPrintRangeOfPages: 4,
}

// enum WdDictionaryType
var WdDictionaryType = struct {
	WdSpelling int32
	WdGrammar int32
	WdThesaurus int32
	WdHyphenation int32
	WdSpellingComplete int32
	WdSpellingCustom int32
	WdSpellingLegal int32
	WdSpellingMedical int32
	WdHangulHanjaConversion int32
	WdHangulHanjaConversionCustom int32
}{
	WdSpelling: 0,
	WdGrammar: 1,
	WdThesaurus: 2,
	WdHyphenation: 3,
	WdSpellingComplete: 4,
	WdSpellingCustom: 5,
	WdSpellingLegal: 6,
	WdSpellingMedical: 7,
	WdHangulHanjaConversion: 8,
	WdHangulHanjaConversionCustom: 9,
}

// enum WdDictionaryTypeHID
var WdDictionaryTypeHID = struct {
	Emptyenum int32
}{
	Emptyenum: 0,
}

// enum WdSpellingWordType
var WdSpellingWordType = struct {
	WdSpellword int32
	WdWildcard int32
	WdAnagram int32
}{
	WdSpellword: 0,
	WdWildcard: 1,
	WdAnagram: 2,
}

// enum WdSpellingErrorType
var WdSpellingErrorType = struct {
	WdSpellingCorrect int32
	WdSpellingNotInDictionary int32
	WdSpellingCapitalization int32
}{
	WdSpellingCorrect: 0,
	WdSpellingNotInDictionary: 1,
	WdSpellingCapitalization: 2,
}

// enum WdProofreadingErrorType
var WdProofreadingErrorType = struct {
	WdSpellingError int32
	WdGrammaticalError int32
}{
	WdSpellingError: 0,
	WdGrammaticalError: 1,
}

// enum WdInlineShapeType
var WdInlineShapeType = struct {
	WdInlineShapeEmbeddedOLEObject int32
	WdInlineShapeLinkedOLEObject int32
	WdInlineShapePicture int32
	WdInlineShapeLinkedPicture int32
	WdInlineShapeOLEControlObject int32
	WdInlineShapeHorizontalLine int32
	WdInlineShapePictureHorizontalLine int32
	WdInlineShapeLinkedPictureHorizontalLine int32
	WdInlineShapePictureBullet int32
	WdInlineShapeScriptAnchor int32
	WdInlineShapeOWSAnchor int32
	WdInlineShapeChart int32
	WdInlineShapeDiagram int32
	WdInlineShapeLockedCanvas int32
	WdInlineShapeSmartArt int32
}{
	WdInlineShapeEmbeddedOLEObject: 1,
	WdInlineShapeLinkedOLEObject: 2,
	WdInlineShapePicture: 3,
	WdInlineShapeLinkedPicture: 4,
	WdInlineShapeOLEControlObject: 5,
	WdInlineShapeHorizontalLine: 6,
	WdInlineShapePictureHorizontalLine: 7,
	WdInlineShapeLinkedPictureHorizontalLine: 8,
	WdInlineShapePictureBullet: 9,
	WdInlineShapeScriptAnchor: 10,
	WdInlineShapeOWSAnchor: 11,
	WdInlineShapeChart: 12,
	WdInlineShapeDiagram: 13,
	WdInlineShapeLockedCanvas: 14,
	WdInlineShapeSmartArt: 15,
}

// enum WdArrangeStyle
var WdArrangeStyle = struct {
	WdTiled int32
	WdIcons int32
}{
	WdTiled: 0,
	WdIcons: 1,
}

// enum WdSelectionFlags
var WdSelectionFlags = struct {
	WdSelStartActive int32
	WdSelAtEOL int32
	WdSelOvertype int32
	WdSelActive int32
	WdSelReplace int32
}{
	WdSelStartActive: 1,
	WdSelAtEOL: 2,
	WdSelOvertype: 4,
	WdSelActive: 8,
	WdSelReplace: 16,
}

// enum WdAutoVersions
var WdAutoVersions = struct {
	WdAutoVersionOff int32
	WdAutoVersionOnClose int32
}{
	WdAutoVersionOff: 0,
	WdAutoVersionOnClose: 1,
}

// enum WdOrganizerObject
var WdOrganizerObject = struct {
	WdOrganizerObjectStyles int32
	WdOrganizerObjectAutoText int32
	WdOrganizerObjectCommandBars int32
	WdOrganizerObjectProjectItems int32
}{
	WdOrganizerObjectStyles: 0,
	WdOrganizerObjectAutoText: 1,
	WdOrganizerObjectCommandBars: 2,
	WdOrganizerObjectProjectItems: 3,
}

// enum WdFindMatch
var WdFindMatch = struct {
	WdMatchParagraphMark int32
	WdMatchTabCharacter int32
	WdMatchCommentMark int32
	WdMatchAnyCharacter int32
	WdMatchAnyDigit int32
	WdMatchAnyLetter int32
	WdMatchCaretCharacter int32
	WdMatchColumnBreak int32
	WdMatchEmDash int32
	WdMatchEnDash int32
	WdMatchEndnoteMark int32
	WdMatchField int32
	WdMatchFootnoteMark int32
	WdMatchGraphic int32
	WdMatchManualLineBreak int32
	WdMatchManualPageBreak int32
	WdMatchNonbreakingHyphen int32
	WdMatchNonbreakingSpace int32
	WdMatchOptionalHyphen int32
	WdMatchSectionBreak int32
	WdMatchWhiteSpace int32
}{
	WdMatchParagraphMark: 65551,
	WdMatchTabCharacter: 9,
	WdMatchCommentMark: 5,
	WdMatchAnyCharacter: 65599,
	WdMatchAnyDigit: 65567,
	WdMatchAnyLetter: 65583,
	WdMatchCaretCharacter: 11,
	WdMatchColumnBreak: 14,
	WdMatchEmDash: 8212,
	WdMatchEnDash: 8211,
	WdMatchEndnoteMark: 65555,
	WdMatchField: 19,
	WdMatchFootnoteMark: 65554,
	WdMatchGraphic: 1,
	WdMatchManualLineBreak: 65551,
	WdMatchManualPageBreak: 65564,
	WdMatchNonbreakingHyphen: 30,
	WdMatchNonbreakingSpace: 160,
	WdMatchOptionalHyphen: 31,
	WdMatchSectionBreak: 65580,
	WdMatchWhiteSpace: 65655,
}

// enum WdFindWrap
var WdFindWrap = struct {
	WdFindStop int32
	WdFindContinue int32
	WdFindAsk int32
}{
	WdFindStop: 0,
	WdFindContinue: 1,
	WdFindAsk: 2,
}

// enum WdInformation
var WdInformation = struct {
	WdActiveEndAdjustedPageNumber int32
	WdActiveEndSectionNumber int32
	WdActiveEndPageNumber int32
	WdNumberOfPagesInDocument int32
	WdHorizontalPositionRelativeToPage int32
	WdVerticalPositionRelativeToPage int32
	WdHorizontalPositionRelativeToTextBoundary int32
	WdVerticalPositionRelativeToTextBoundary int32
	WdFirstCharacterColumnNumber int32
	WdFirstCharacterLineNumber int32
	WdFrameIsSelected int32
	WdWithInTable int32
	WdStartOfRangeRowNumber int32
	WdEndOfRangeRowNumber int32
	WdMaximumNumberOfRows int32
	WdStartOfRangeColumnNumber int32
	WdEndOfRangeColumnNumber int32
	WdMaximumNumberOfColumns int32
	WdZoomPercentage int32
	WdSelectionMode int32
	WdCapsLock int32
	WdNumLock int32
	WdOverType int32
	WdRevisionMarking int32
	WdInFootnoteEndnotePane int32
	WdInCommentPane int32
	WdInHeaderFooter int32
	WdAtEndOfRowMarker int32
	WdReferenceOfType int32
	WdHeaderFooterType int32
	WdInMasterDocument int32
	WdInFootnote int32
	WdInEndnote int32
	WdInWordMail int32
	WdInClipboard int32
}{
	WdActiveEndAdjustedPageNumber: 1,
	WdActiveEndSectionNumber: 2,
	WdActiveEndPageNumber: 3,
	WdNumberOfPagesInDocument: 4,
	WdHorizontalPositionRelativeToPage: 5,
	WdVerticalPositionRelativeToPage: 6,
	WdHorizontalPositionRelativeToTextBoundary: 7,
	WdVerticalPositionRelativeToTextBoundary: 8,
	WdFirstCharacterColumnNumber: 9,
	WdFirstCharacterLineNumber: 10,
	WdFrameIsSelected: 11,
	WdWithInTable: 12,
	WdStartOfRangeRowNumber: 13,
	WdEndOfRangeRowNumber: 14,
	WdMaximumNumberOfRows: 15,
	WdStartOfRangeColumnNumber: 16,
	WdEndOfRangeColumnNumber: 17,
	WdMaximumNumberOfColumns: 18,
	WdZoomPercentage: 19,
	WdSelectionMode: 20,
	WdCapsLock: 21,
	WdNumLock: 22,
	WdOverType: 23,
	WdRevisionMarking: 24,
	WdInFootnoteEndnotePane: 25,
	WdInCommentPane: 26,
	WdInHeaderFooter: 28,
	WdAtEndOfRowMarker: 31,
	WdReferenceOfType: 32,
	WdHeaderFooterType: 33,
	WdInMasterDocument: 34,
	WdInFootnote: 35,
	WdInEndnote: 36,
	WdInWordMail: 37,
	WdInClipboard: 38,
}

// enum WdWrapType
var WdWrapType = struct {
	WdWrapSquare int32
	WdWrapTight int32
	WdWrapThrough int32
	WdWrapNone int32
	WdWrapTopBottom int32
	WdWrapBehind int32
	WdWrapFront int32
	WdWrapInline int32
}{
	WdWrapSquare: 0,
	WdWrapTight: 1,
	WdWrapThrough: 2,
	WdWrapNone: 3,
	WdWrapTopBottom: 4,
	WdWrapBehind: 5,
	WdWrapFront: 3,
	WdWrapInline: 7,
}

// enum WdWrapSideType
var WdWrapSideType = struct {
	WdWrapBoth int32
	WdWrapLeft int32
	WdWrapRight int32
	WdWrapLargest int32
}{
	WdWrapBoth: 0,
	WdWrapLeft: 1,
	WdWrapRight: 2,
	WdWrapLargest: 3,
}

// enum WdOutlineLevel
var WdOutlineLevel = struct {
	WdOutlineLevel1 int32
	WdOutlineLevel2 int32
	WdOutlineLevel3 int32
	WdOutlineLevel4 int32
	WdOutlineLevel5 int32
	WdOutlineLevel6 int32
	WdOutlineLevel7 int32
	WdOutlineLevel8 int32
	WdOutlineLevel9 int32
	WdOutlineLevelBodyText int32
}{
	WdOutlineLevel1: 1,
	WdOutlineLevel2: 2,
	WdOutlineLevel3: 3,
	WdOutlineLevel4: 4,
	WdOutlineLevel5: 5,
	WdOutlineLevel6: 6,
	WdOutlineLevel7: 7,
	WdOutlineLevel8: 8,
	WdOutlineLevel9: 9,
	WdOutlineLevelBodyText: 10,
}

// enum WdTextOrientation
var WdTextOrientation = struct {
	WdTextOrientationHorizontal int32
	WdTextOrientationUpward int32
	WdTextOrientationDownward int32
	WdTextOrientationVerticalFarEast int32
	WdTextOrientationHorizontalRotatedFarEast int32
	WdTextOrientationVertical int32
}{
	WdTextOrientationHorizontal: 0,
	WdTextOrientationUpward: 2,
	WdTextOrientationDownward: 3,
	WdTextOrientationVerticalFarEast: 1,
	WdTextOrientationHorizontalRotatedFarEast: 4,
	WdTextOrientationVertical: 5,
}

// enum WdTextOrientationHID
var WdTextOrientationHID = struct {
	Emptyenum int32
}{
	Emptyenum: 0,
}

// enum WdPageBorderArt
var WdPageBorderArt = struct {
	WdArtApples int32
	WdArtMapleMuffins int32
	WdArtCakeSlice int32
	WdArtCandyCorn int32
	WdArtIceCreamCones int32
	WdArtChampagneBottle int32
	WdArtPartyGlass int32
	WdArtChristmasTree int32
	WdArtTrees int32
	WdArtPalmsColor int32
	WdArtBalloons3Colors int32
	WdArtBalloonsHotAir int32
	WdArtPartyFavor int32
	WdArtConfettiStreamers int32
	WdArtHearts int32
	WdArtHeartBalloon int32
	WdArtStars3D int32
	WdArtStarsShadowed int32
	WdArtStars int32
	WdArtSun int32
	WdArtEarth2 int32
	WdArtEarth1 int32
	WdArtPeopleHats int32
	WdArtSombrero int32
	WdArtPencils int32
	WdArtPackages int32
	WdArtClocks int32
	WdArtFirecrackers int32
	WdArtRings int32
	WdArtMapPins int32
	WdArtConfetti int32
	WdArtCreaturesButterfly int32
	WdArtCreaturesLadyBug int32
	WdArtCreaturesFish int32
	WdArtBirdsFlight int32
	WdArtScaredCat int32
	WdArtBats int32
	WdArtFlowersRoses int32
	WdArtFlowersRedRose int32
	WdArtPoinsettias int32
	WdArtHolly int32
	WdArtFlowersTiny int32
	WdArtFlowersPansy int32
	WdArtFlowersModern2 int32
	WdArtFlowersModern1 int32
	WdArtWhiteFlowers int32
	WdArtVine int32
	WdArtFlowersDaisies int32
	WdArtFlowersBlockPrint int32
	WdArtDecoArchColor int32
	WdArtFans int32
	WdArtFilm int32
	WdArtLightning1 int32
	WdArtCompass int32
	WdArtDoubleD int32
	WdArtClassicalWave int32
	WdArtShadowedSquares int32
	WdArtTwistedLines1 int32
	WdArtWaveline int32
	WdArtQuadrants int32
	WdArtCheckedBarColor int32
	WdArtSwirligig int32
	WdArtPushPinNote1 int32
	WdArtPushPinNote2 int32
	WdArtPumpkin1 int32
	WdArtEggsBlack int32
	WdArtCup int32
	WdArtHeartGray int32
	WdArtGingerbreadMan int32
	WdArtBabyPacifier int32
	WdArtBabyRattle int32
	WdArtCabins int32
	WdArtHouseFunky int32
	WdArtStarsBlack int32
	WdArtSnowflakes int32
	WdArtSnowflakeFancy int32
	WdArtSkyrocket int32
	WdArtSeattle int32
	WdArtMusicNotes int32
	WdArtPalmsBlack int32
	WdArtMapleLeaf int32
	WdArtPaperClips int32
	WdArtShorebirdTracks int32
	WdArtPeople int32
	WdArtPeopleWaving int32
	WdArtEclipsingSquares2 int32
	WdArtHypnotic int32
	WdArtDiamondsGray int32
	WdArtDecoArch int32
	WdArtDecoBlocks int32
	WdArtCirclesLines int32
	WdArtPapyrus int32
	WdArtWoodwork int32
	WdArtWeavingBraid int32
	WdArtWeavingRibbon int32
	WdArtWeavingAngles int32
	WdArtArchedScallops int32
	WdArtSafari int32
	WdArtCelticKnotwork int32
	WdArtCrazyMaze int32
	WdArtEclipsingSquares1 int32
	WdArtBirds int32
	WdArtFlowersTeacup int32
	WdArtNorthwest int32
	WdArtSouthwest int32
	WdArtTribal6 int32
	WdArtTribal4 int32
	WdArtTribal3 int32
	WdArtTribal2 int32
	WdArtTribal5 int32
	WdArtXIllusions int32
	WdArtZanyTriangles int32
	WdArtPyramids int32
	WdArtPyramidsAbove int32
	WdArtConfettiGrays int32
	WdArtConfettiOutline int32
	WdArtConfettiWhite int32
	WdArtMosaic int32
	WdArtLightning2 int32
	WdArtHeebieJeebies int32
	WdArtLightBulb int32
	WdArtGradient int32
	WdArtTriangleParty int32
	WdArtTwistedLines2 int32
	WdArtMoons int32
	WdArtOvals int32
	WdArtDoubleDiamonds int32
	WdArtChainLink int32
	WdArtTriangles int32
	WdArtTribal1 int32
	WdArtMarqueeToothed int32
	WdArtSharksTeeth int32
	WdArtSawtooth int32
	WdArtSawtoothGray int32
	WdArtPostageStamp int32
	WdArtWeavingStrips int32
	WdArtZigZag int32
	WdArtCrossStitch int32
	WdArtGems int32
	WdArtCirclesRectangles int32
	WdArtCornerTriangles int32
	WdArtCreaturesInsects int32
	WdArtZigZagStitch int32
	WdArtCheckered int32
	WdArtCheckedBarBlack int32
	WdArtMarquee int32
	WdArtBasicWhiteDots int32
	WdArtBasicWideMidline int32
	WdArtBasicWideOutline int32
	WdArtBasicWideInline int32
	WdArtBasicThinLines int32
	WdArtBasicWhiteDashes int32
	WdArtBasicWhiteSquares int32
	WdArtBasicBlackSquares int32
	WdArtBasicBlackDashes int32
	WdArtBasicBlackDots int32
	WdArtStarsTop int32
	WdArtCertificateBanner int32
	WdArtHandmade1 int32
	WdArtHandmade2 int32
	WdArtTornPaper int32
	WdArtTornPaperBlack int32
	WdArtCouponCutoutDashes int32
	WdArtCouponCutoutDots int32
}{
	WdArtApples: 1,
	WdArtMapleMuffins: 2,
	WdArtCakeSlice: 3,
	WdArtCandyCorn: 4,
	WdArtIceCreamCones: 5,
	WdArtChampagneBottle: 6,
	WdArtPartyGlass: 7,
	WdArtChristmasTree: 8,
	WdArtTrees: 9,
	WdArtPalmsColor: 10,
	WdArtBalloons3Colors: 11,
	WdArtBalloonsHotAir: 12,
	WdArtPartyFavor: 13,
	WdArtConfettiStreamers: 14,
	WdArtHearts: 15,
	WdArtHeartBalloon: 16,
	WdArtStars3D: 17,
	WdArtStarsShadowed: 18,
	WdArtStars: 19,
	WdArtSun: 20,
	WdArtEarth2: 21,
	WdArtEarth1: 22,
	WdArtPeopleHats: 23,
	WdArtSombrero: 24,
	WdArtPencils: 25,
	WdArtPackages: 26,
	WdArtClocks: 27,
	WdArtFirecrackers: 28,
	WdArtRings: 29,
	WdArtMapPins: 30,
	WdArtConfetti: 31,
	WdArtCreaturesButterfly: 32,
	WdArtCreaturesLadyBug: 33,
	WdArtCreaturesFish: 34,
	WdArtBirdsFlight: 35,
	WdArtScaredCat: 36,
	WdArtBats: 37,
	WdArtFlowersRoses: 38,
	WdArtFlowersRedRose: 39,
	WdArtPoinsettias: 40,
	WdArtHolly: 41,
	WdArtFlowersTiny: 42,
	WdArtFlowersPansy: 43,
	WdArtFlowersModern2: 44,
	WdArtFlowersModern1: 45,
	WdArtWhiteFlowers: 46,
	WdArtVine: 47,
	WdArtFlowersDaisies: 48,
	WdArtFlowersBlockPrint: 49,
	WdArtDecoArchColor: 50,
	WdArtFans: 51,
	WdArtFilm: 52,
	WdArtLightning1: 53,
	WdArtCompass: 54,
	WdArtDoubleD: 55,
	WdArtClassicalWave: 56,
	WdArtShadowedSquares: 57,
	WdArtTwistedLines1: 58,
	WdArtWaveline: 59,
	WdArtQuadrants: 60,
	WdArtCheckedBarColor: 61,
	WdArtSwirligig: 62,
	WdArtPushPinNote1: 63,
	WdArtPushPinNote2: 64,
	WdArtPumpkin1: 65,
	WdArtEggsBlack: 66,
	WdArtCup: 67,
	WdArtHeartGray: 68,
	WdArtGingerbreadMan: 69,
	WdArtBabyPacifier: 70,
	WdArtBabyRattle: 71,
	WdArtCabins: 72,
	WdArtHouseFunky: 73,
	WdArtStarsBlack: 74,
	WdArtSnowflakes: 75,
	WdArtSnowflakeFancy: 76,
	WdArtSkyrocket: 77,
	WdArtSeattle: 78,
	WdArtMusicNotes: 79,
	WdArtPalmsBlack: 80,
	WdArtMapleLeaf: 81,
	WdArtPaperClips: 82,
	WdArtShorebirdTracks: 83,
	WdArtPeople: 84,
	WdArtPeopleWaving: 85,
	WdArtEclipsingSquares2: 86,
	WdArtHypnotic: 87,
	WdArtDiamondsGray: 88,
	WdArtDecoArch: 89,
	WdArtDecoBlocks: 90,
	WdArtCirclesLines: 91,
	WdArtPapyrus: 92,
	WdArtWoodwork: 93,
	WdArtWeavingBraid: 94,
	WdArtWeavingRibbon: 95,
	WdArtWeavingAngles: 96,
	WdArtArchedScallops: 97,
	WdArtSafari: 98,
	WdArtCelticKnotwork: 99,
	WdArtCrazyMaze: 100,
	WdArtEclipsingSquares1: 101,
	WdArtBirds: 102,
	WdArtFlowersTeacup: 103,
	WdArtNorthwest: 104,
	WdArtSouthwest: 105,
	WdArtTribal6: 106,
	WdArtTribal4: 107,
	WdArtTribal3: 108,
	WdArtTribal2: 109,
	WdArtTribal5: 110,
	WdArtXIllusions: 111,
	WdArtZanyTriangles: 112,
	WdArtPyramids: 113,
	WdArtPyramidsAbove: 114,
	WdArtConfettiGrays: 115,
	WdArtConfettiOutline: 116,
	WdArtConfettiWhite: 117,
	WdArtMosaic: 118,
	WdArtLightning2: 119,
	WdArtHeebieJeebies: 120,
	WdArtLightBulb: 121,
	WdArtGradient: 122,
	WdArtTriangleParty: 123,
	WdArtTwistedLines2: 124,
	WdArtMoons: 125,
	WdArtOvals: 126,
	WdArtDoubleDiamonds: 127,
	WdArtChainLink: 128,
	WdArtTriangles: 129,
	WdArtTribal1: 130,
	WdArtMarqueeToothed: 131,
	WdArtSharksTeeth: 132,
	WdArtSawtooth: 133,
	WdArtSawtoothGray: 134,
	WdArtPostageStamp: 135,
	WdArtWeavingStrips: 136,
	WdArtZigZag: 137,
	WdArtCrossStitch: 138,
	WdArtGems: 139,
	WdArtCirclesRectangles: 140,
	WdArtCornerTriangles: 141,
	WdArtCreaturesInsects: 142,
	WdArtZigZagStitch: 143,
	WdArtCheckered: 144,
	WdArtCheckedBarBlack: 145,
	WdArtMarquee: 146,
	WdArtBasicWhiteDots: 147,
	WdArtBasicWideMidline: 148,
	WdArtBasicWideOutline: 149,
	WdArtBasicWideInline: 150,
	WdArtBasicThinLines: 151,
	WdArtBasicWhiteDashes: 152,
	WdArtBasicWhiteSquares: 153,
	WdArtBasicBlackSquares: 154,
	WdArtBasicBlackDashes: 155,
	WdArtBasicBlackDots: 156,
	WdArtStarsTop: 157,
	WdArtCertificateBanner: 158,
	WdArtHandmade1: 159,
	WdArtHandmade2: 160,
	WdArtTornPaper: 161,
	WdArtTornPaperBlack: 162,
	WdArtCouponCutoutDashes: 163,
	WdArtCouponCutoutDots: 164,
}

// enum WdBorderDistanceFrom
var WdBorderDistanceFrom = struct {
	WdBorderDistanceFromText int32
	WdBorderDistanceFromPageEdge int32
}{
	WdBorderDistanceFromText: 0,
	WdBorderDistanceFromPageEdge: 1,
}

// enum WdReplace
var WdReplace = struct {
	WdReplaceNone int32
	WdReplaceOne int32
	WdReplaceAll int32
}{
	WdReplaceNone: 0,
	WdReplaceOne: 1,
	WdReplaceAll: 2,
}

// enum WdFontBias
var WdFontBias = struct {
	WdFontBiasDontCare int32
	WdFontBiasDefault int32
	WdFontBiasFareast int32
}{
	WdFontBiasDontCare: 255,
	WdFontBiasDefault: 0,
	WdFontBiasFareast: 1,
}

// enum WdBrowserLevel
var WdBrowserLevel = struct {
	WdBrowserLevelV4 int32
	WdBrowserLevelMicrosoftInternetExplorer5 int32
	WdBrowserLevelMicrosoftInternetExplorer6 int32
}{
	WdBrowserLevelV4: 0,
	WdBrowserLevelMicrosoftInternetExplorer5: 1,
	WdBrowserLevelMicrosoftInternetExplorer6: 2,
}

// enum WdEnclosureType
var WdEnclosureType = struct {
	WdEnclosureCircle int32
	WdEnclosureSquare int32
	WdEnclosureTriangle int32
	WdEnclosureDiamond int32
}{
	WdEnclosureCircle: 0,
	WdEnclosureSquare: 1,
	WdEnclosureTriangle: 2,
	WdEnclosureDiamond: 3,
}

// enum WdEncloseStyle
var WdEncloseStyle = struct {
	WdEncloseStyleNone int32
	WdEncloseStyleSmall int32
	WdEncloseStyleLarge int32
}{
	WdEncloseStyleNone: 0,
	WdEncloseStyleSmall: 1,
	WdEncloseStyleLarge: 2,
}

// enum WdHighAnsiText
var WdHighAnsiText = struct {
	WdHighAnsiIsFarEast int32
	WdHighAnsiIsHighAnsi int32
	WdAutoDetectHighAnsiFarEast int32
}{
	WdHighAnsiIsFarEast: 0,
	WdHighAnsiIsHighAnsi: 1,
	WdAutoDetectHighAnsiFarEast: 2,
}

// enum WdLayoutMode
var WdLayoutMode = struct {
	WdLayoutModeDefault int32
	WdLayoutModeGrid int32
	WdLayoutModeLineGrid int32
	WdLayoutModeGenko int32
}{
	WdLayoutModeDefault: 0,
	WdLayoutModeGrid: 1,
	WdLayoutModeLineGrid: 2,
	WdLayoutModeGenko: 3,
}

// enum WdDocumentMedium
var WdDocumentMedium = struct {
	WdEmailMessage int32
	WdDocument int32
	WdWebPage int32
}{
	WdEmailMessage: 0,
	WdDocument: 1,
	WdWebPage: 2,
}

// enum WdMailerPriority
var WdMailerPriority = struct {
	WdPriorityNormal int32
	WdPriorityLow int32
	WdPriorityHigh int32
}{
	WdPriorityNormal: 1,
	WdPriorityLow: 2,
	WdPriorityHigh: 3,
}

// enum WdDocumentViewDirection
var WdDocumentViewDirection = struct {
	WdDocumentViewRtl int32
	WdDocumentViewLtr int32
}{
	WdDocumentViewRtl: 0,
	WdDocumentViewLtr: 1,
}

// enum WdArabicNumeral
var WdArabicNumeral = struct {
	WdNumeralArabic int32
	WdNumeralHindi int32
	WdNumeralContext int32
	WdNumeralSystem int32
}{
	WdNumeralArabic: 0,
	WdNumeralHindi: 1,
	WdNumeralContext: 2,
	WdNumeralSystem: 3,
}

// enum WdMonthNames
var WdMonthNames = struct {
	WdMonthNamesArabic int32
	WdMonthNamesEnglish int32
	WdMonthNamesFrench int32
}{
	WdMonthNamesArabic: 0,
	WdMonthNamesEnglish: 1,
	WdMonthNamesFrench: 2,
}

// enum WdCursorMovement
var WdCursorMovement = struct {
	WdCursorMovementLogical int32
	WdCursorMovementVisual int32
}{
	WdCursorMovementLogical: 0,
	WdCursorMovementVisual: 1,
}

// enum WdVisualSelection
var WdVisualSelection = struct {
	WdVisualSelectionBlock int32
	WdVisualSelectionContinuous int32
}{
	WdVisualSelectionBlock: 0,
	WdVisualSelectionContinuous: 1,
}

// enum WdTableDirection
var WdTableDirection = struct {
	WdTableDirectionRtl int32
	WdTableDirectionLtr int32
}{
	WdTableDirectionRtl: 0,
	WdTableDirectionLtr: 1,
}

// enum WdFlowDirection
var WdFlowDirection = struct {
	WdFlowLtr int32
	WdFlowRtl int32
}{
	WdFlowLtr: 0,
	WdFlowRtl: 1,
}

// enum WdDiacriticColor
var WdDiacriticColor = struct {
	WdDiacriticColorBidi int32
	WdDiacriticColorLatin int32
}{
	WdDiacriticColorBidi: 0,
	WdDiacriticColorLatin: 1,
}

// enum WdGutterStyle
var WdGutterStyle = struct {
	WdGutterPosLeft int32
	WdGutterPosTop int32
	WdGutterPosRight int32
}{
	WdGutterPosLeft: 0,
	WdGutterPosTop: 1,
	WdGutterPosRight: 2,
}

// enum WdGutterStyleOld
var WdGutterStyleOld = struct {
	WdGutterStyleLatin int32
	WdGutterStyleBidi int32
}{
	WdGutterStyleLatin: -10,
	WdGutterStyleBidi: 2,
}

// enum WdSectionDirection
var WdSectionDirection = struct {
	WdSectionDirectionRtl int32
	WdSectionDirectionLtr int32
}{
	WdSectionDirectionRtl: 0,
	WdSectionDirectionLtr: 1,
}

// enum WdDateLanguage
var WdDateLanguage = struct {
	WdDateLanguageBidi int32
	WdDateLanguageLatin int32
}{
	WdDateLanguageBidi: 10,
	WdDateLanguageLatin: 1033,
}

// enum WdCalendarTypeBi
var WdCalendarTypeBi = struct {
	WdCalendarTypeBidi int32
	WdCalendarTypeGregorian int32
}{
	WdCalendarTypeBidi: 99,
	WdCalendarTypeGregorian: 100,
}

// enum WdCalendarType
var WdCalendarType = struct {
	WdCalendarWestern int32
	WdCalendarArabic int32
	WdCalendarHebrew int32
	WdCalendarTaiwan int32
	WdCalendarJapan int32
	WdCalendarThai int32
	WdCalendarKorean int32
	WdCalendarSakaEra int32
	WdCalendarTranslitEnglish int32
	WdCalendarTranslitFrench int32
	WdCalendarUmalqura int32
}{
	WdCalendarWestern: 0,
	WdCalendarArabic: 1,
	WdCalendarHebrew: 2,
	WdCalendarTaiwan: 3,
	WdCalendarJapan: 4,
	WdCalendarThai: 5,
	WdCalendarKorean: 6,
	WdCalendarSakaEra: 7,
	WdCalendarTranslitEnglish: 8,
	WdCalendarTranslitFrench: 9,
	WdCalendarUmalqura: 13,
}

// enum WdReadingOrder
var WdReadingOrder = struct {
	WdReadingOrderRtl int32
	WdReadingOrderLtr int32
}{
	WdReadingOrderRtl: 0,
	WdReadingOrderLtr: 1,
}

// enum WdHebSpellStart
var WdHebSpellStart = struct {
	WdFullScript int32
	WdPartialScript int32
	WdMixedScript int32
	WdMixedAuthorizedScript int32
}{
	WdFullScript: 0,
	WdPartialScript: 1,
	WdMixedScript: 2,
	WdMixedAuthorizedScript: 3,
}

// enum WdAraSpeller
var WdAraSpeller = struct {
	WdNone int32
	WdInitialAlef int32
	WdFinalYaa int32
	WdBoth int32
}{
	WdNone: 0,
	WdInitialAlef: 1,
	WdFinalYaa: 2,
	WdBoth: 3,
}

// enum WdColor
var WdColor = struct {
	WdColorAutomatic int32
	WdColorBlack int32
	WdColorBlue int32
	WdColorTurquoise int32
	WdColorBrightGreen int32
	WdColorPink int32
	WdColorRed int32
	WdColorYellow int32
	WdColorWhite int32
	WdColorDarkBlue int32
	WdColorTeal int32
	WdColorGreen int32
	WdColorViolet int32
	WdColorDarkRed int32
	WdColorDarkYellow int32
	WdColorBrown int32
	WdColorOliveGreen int32
	WdColorDarkGreen int32
	WdColorDarkTeal int32
	WdColorIndigo int32
	WdColorOrange int32
	WdColorBlueGray int32
	WdColorLightOrange int32
	WdColorLime int32
	WdColorSeaGreen int32
	WdColorAqua int32
	WdColorLightBlue int32
	WdColorGold int32
	WdColorSkyBlue int32
	WdColorPlum int32
	WdColorRose int32
	WdColorTan int32
	WdColorLightYellow int32
	WdColorLightGreen int32
	WdColorLightTurquoise int32
	WdColorPaleBlue int32
	WdColorLavender int32
	WdColorGray05 int32
	WdColorGray10 int32
	WdColorGray125 int32
	WdColorGray15 int32
	WdColorGray20 int32
	WdColorGray25 int32
	WdColorGray30 int32
	WdColorGray35 int32
	WdColorGray375 int32
	WdColorGray40 int32
	WdColorGray45 int32
	WdColorGray50 int32
	WdColorGray55 int32
	WdColorGray60 int32
	WdColorGray625 int32
	WdColorGray65 int32
	WdColorGray70 int32
	WdColorGray75 int32
	WdColorGray80 int32
	WdColorGray85 int32
	WdColorGray875 int32
	WdColorGray90 int32
	WdColorGray95 int32
}{
	WdColorAutomatic: -16777216,
	WdColorBlack: 0,
	WdColorBlue: 16711680,
	WdColorTurquoise: 16776960,
	WdColorBrightGreen: 65280,
	WdColorPink: 16711935,
	WdColorRed: 255,
	WdColorYellow: 65535,
	WdColorWhite: 16777215,
	WdColorDarkBlue: 8388608,
	WdColorTeal: 8421376,
	WdColorGreen: 32768,
	WdColorViolet: 8388736,
	WdColorDarkRed: 128,
	WdColorDarkYellow: 32896,
	WdColorBrown: 13209,
	WdColorOliveGreen: 13107,
	WdColorDarkGreen: 13056,
	WdColorDarkTeal: 6697728,
	WdColorIndigo: 10040115,
	WdColorOrange: 26367,
	WdColorBlueGray: 10053222,
	WdColorLightOrange: 39423,
	WdColorLime: 52377,
	WdColorSeaGreen: 6723891,
	WdColorAqua: 13421619,
	WdColorLightBlue: 16737843,
	WdColorGold: 52479,
	WdColorSkyBlue: 16763904,
	WdColorPlum: 6697881,
	WdColorRose: 13408767,
	WdColorTan: 10079487,
	WdColorLightYellow: 10092543,
	WdColorLightGreen: 13434828,
	WdColorLightTurquoise: 16777164,
	WdColorPaleBlue: 16764057,
	WdColorLavender: 16751052,
	WdColorGray05: 15987699,
	WdColorGray10: 15132390,
	WdColorGray125: 14737632,
	WdColorGray15: 14277081,
	WdColorGray20: 13421772,
	WdColorGray25: 12632256,
	WdColorGray30: 11776947,
	WdColorGray35: 10921638,
	WdColorGray375: 10526880,
	WdColorGray40: 10066329,
	WdColorGray45: 9211020,
	WdColorGray50: 8421504,
	WdColorGray55: 7566195,
	WdColorGray60: 6710886,
	WdColorGray625: 6316128,
	WdColorGray65: 5855577,
	WdColorGray70: 5000268,
	WdColorGray75: 4210752,
	WdColorGray80: 3355443,
	WdColorGray85: 2500134,
	WdColorGray875: 2105376,
	WdColorGray90: 1644825,
	WdColorGray95: 789516,
}

// enum WdShapePosition
var WdShapePosition = struct {
	WdShapeTop int32
	WdShapeLeft int32
	WdShapeBottom int32
	WdShapeRight int32
	WdShapeCenter int32
	WdShapeInside int32
	WdShapeOutside int32
}{
	WdShapeTop: -999999,
	WdShapeLeft: -999998,
	WdShapeBottom: -999997,
	WdShapeRight: -999996,
	WdShapeCenter: -999995,
	WdShapeInside: -999994,
	WdShapeOutside: -999993,
}

// enum WdTablePosition
var WdTablePosition = struct {
	WdTableTop int32
	WdTableLeft int32
	WdTableBottom int32
	WdTableRight int32
	WdTableCenter int32
	WdTableInside int32
	WdTableOutside int32
}{
	WdTableTop: -999999,
	WdTableLeft: -999998,
	WdTableBottom: -999997,
	WdTableRight: -999996,
	WdTableCenter: -999995,
	WdTableInside: -999994,
	WdTableOutside: -999993,
}

// enum WdDefaultListBehavior
var WdDefaultListBehavior = struct {
	WdWord8ListBehavior int32
	WdWord9ListBehavior int32
	WdWord10ListBehavior int32
}{
	WdWord8ListBehavior: 0,
	WdWord9ListBehavior: 1,
	WdWord10ListBehavior: 2,
}

// enum WdDefaultTableBehavior
var WdDefaultTableBehavior = struct {
	WdWord8TableBehavior int32
	WdWord9TableBehavior int32
}{
	WdWord8TableBehavior: 0,
	WdWord9TableBehavior: 1,
}

// enum WdAutoFitBehavior
var WdAutoFitBehavior = struct {
	WdAutoFitFixed int32
	WdAutoFitContent int32
	WdAutoFitWindow int32
}{
	WdAutoFitFixed: 0,
	WdAutoFitContent: 1,
	WdAutoFitWindow: 2,
}

// enum WdPreferredWidthType
var WdPreferredWidthType = struct {
	WdPreferredWidthAuto int32
	WdPreferredWidthPercent int32
	WdPreferredWidthPoints int32
}{
	WdPreferredWidthAuto: 1,
	WdPreferredWidthPercent: 2,
	WdPreferredWidthPoints: 3,
}

// enum WdFarEastLineBreakLanguageID
var WdFarEastLineBreakLanguageID = struct {
	WdLineBreakJapanese int32
	WdLineBreakKorean int32
	WdLineBreakSimplifiedChinese int32
	WdLineBreakTraditionalChinese int32
}{
	WdLineBreakJapanese: 1041,
	WdLineBreakKorean: 1042,
	WdLineBreakSimplifiedChinese: 2052,
	WdLineBreakTraditionalChinese: 1028,
}

// enum WdViewTypeOld
var WdViewTypeOld = struct {
	WdPageView int32
	WdOnlineView int32
}{
	WdPageView: 3,
	WdOnlineView: 6,
}

// enum WdFramesetType
var WdFramesetType = struct {
	WdFramesetTypeFrameset int32
	WdFramesetTypeFrame int32
}{
	WdFramesetTypeFrameset: 0,
	WdFramesetTypeFrame: 1,
}

// enum WdFramesetSizeType
var WdFramesetSizeType = struct {
	WdFramesetSizeTypePercent int32
	WdFramesetSizeTypeFixed int32
	WdFramesetSizeTypeRelative int32
}{
	WdFramesetSizeTypePercent: 0,
	WdFramesetSizeTypeFixed: 1,
	WdFramesetSizeTypeRelative: 2,
}

// enum WdFramesetNewFrameLocation
var WdFramesetNewFrameLocation = struct {
	WdFramesetNewFrameAbove int32
	WdFramesetNewFrameBelow int32
	WdFramesetNewFrameRight int32
	WdFramesetNewFrameLeft int32
}{
	WdFramesetNewFrameAbove: 0,
	WdFramesetNewFrameBelow: 1,
	WdFramesetNewFrameRight: 2,
	WdFramesetNewFrameLeft: 3,
}

// enum WdScrollbarType
var WdScrollbarType = struct {
	WdScrollbarTypeAuto int32
	WdScrollbarTypeYes int32
	WdScrollbarTypeNo int32
}{
	WdScrollbarTypeAuto: 0,
	WdScrollbarTypeYes: 1,
	WdScrollbarTypeNo: 2,
}

// enum WdTwoLinesInOneType
var WdTwoLinesInOneType = struct {
	WdTwoLinesInOneNone int32
	WdTwoLinesInOneNoBrackets int32
	WdTwoLinesInOneParentheses int32
	WdTwoLinesInOneSquareBrackets int32
	WdTwoLinesInOneAngleBrackets int32
	WdTwoLinesInOneCurlyBrackets int32
}{
	WdTwoLinesInOneNone: 0,
	WdTwoLinesInOneNoBrackets: 1,
	WdTwoLinesInOneParentheses: 2,
	WdTwoLinesInOneSquareBrackets: 3,
	WdTwoLinesInOneAngleBrackets: 4,
	WdTwoLinesInOneCurlyBrackets: 5,
}

// enum WdHorizontalInVerticalType
var WdHorizontalInVerticalType = struct {
	WdHorizontalInVerticalNone int32
	WdHorizontalInVerticalFitInLine int32
	WdHorizontalInVerticalResizeLine int32
}{
	WdHorizontalInVerticalNone: 0,
	WdHorizontalInVerticalFitInLine: 1,
	WdHorizontalInVerticalResizeLine: 2,
}

// enum WdHorizontalLineAlignment
var WdHorizontalLineAlignment = struct {
	WdHorizontalLineAlignLeft int32
	WdHorizontalLineAlignCenter int32
	WdHorizontalLineAlignRight int32
}{
	WdHorizontalLineAlignLeft: 0,
	WdHorizontalLineAlignCenter: 1,
	WdHorizontalLineAlignRight: 2,
}

// enum WdHorizontalLineWidthType
var WdHorizontalLineWidthType = struct {
	WdHorizontalLinePercentWidth int32
	WdHorizontalLineFixedWidth int32
}{
	WdHorizontalLinePercentWidth: -1,
	WdHorizontalLineFixedWidth: -2,
}

// enum WdPhoneticGuideAlignmentType
var WdPhoneticGuideAlignmentType = struct {
	WdPhoneticGuideAlignmentCenter int32
	WdPhoneticGuideAlignmentZeroOneZero int32
	WdPhoneticGuideAlignmentOneTwoOne int32
	WdPhoneticGuideAlignmentLeft int32
	WdPhoneticGuideAlignmentRight int32
	WdPhoneticGuideAlignmentRightVertical int32
}{
	WdPhoneticGuideAlignmentCenter: 0,
	WdPhoneticGuideAlignmentZeroOneZero: 1,
	WdPhoneticGuideAlignmentOneTwoOne: 2,
	WdPhoneticGuideAlignmentLeft: 3,
	WdPhoneticGuideAlignmentRight: 4,
	WdPhoneticGuideAlignmentRightVertical: 5,
}

// enum WdNewDocumentType
var WdNewDocumentType = struct {
	WdNewBlankDocument int32
	WdNewWebPage int32
	WdNewEmailMessage int32
	WdNewFrameset int32
	WdNewXMLDocument int32
}{
	WdNewBlankDocument: 0,
	WdNewWebPage: 1,
	WdNewEmailMessage: 2,
	WdNewFrameset: 3,
	WdNewXMLDocument: 4,
}

// enum WdKana
var WdKana = struct {
	WdKanaKatakana int32
	WdKanaHiragana int32
}{
	WdKanaKatakana: 8,
	WdKanaHiragana: 9,
}

// enum WdCharacterWidth
var WdCharacterWidth = struct {
	WdWidthHalfWidth int32
	WdWidthFullWidth int32
}{
	WdWidthHalfWidth: 6,
	WdWidthFullWidth: 7,
}

// enum WdNumberStyleWordBasicBiDi
var WdNumberStyleWordBasicBiDi = struct {
	WdListNumberStyleBidi1 int32
	WdListNumberStyleBidi2 int32
	WdCaptionNumberStyleBidiLetter1 int32
	WdCaptionNumberStyleBidiLetter2 int32
	WdNoteNumberStyleBidiLetter1 int32
	WdNoteNumberStyleBidiLetter2 int32
	WdPageNumberStyleBidiLetter1 int32
	WdPageNumberStyleBidiLetter2 int32
}{
	WdListNumberStyleBidi1: 49,
	WdListNumberStyleBidi2: 50,
	WdCaptionNumberStyleBidiLetter1: 49,
	WdCaptionNumberStyleBidiLetter2: 50,
	WdNoteNumberStyleBidiLetter1: 49,
	WdNoteNumberStyleBidiLetter2: 50,
	WdPageNumberStyleBidiLetter1: 49,
	WdPageNumberStyleBidiLetter2: 50,
}

// enum WdTCSCConverterDirection
var WdTCSCConverterDirection = struct {
	WdTCSCConverterDirectionSCTC int32
	WdTCSCConverterDirectionTCSC int32
	WdTCSCConverterDirectionAuto int32
}{
	WdTCSCConverterDirectionSCTC: 0,
	WdTCSCConverterDirectionTCSC: 1,
	WdTCSCConverterDirectionAuto: 2,
}

// enum WdDisableFeaturesIntroducedAfter
var WdDisableFeaturesIntroducedAfter = struct {
	Wd70 int32
	Wd70FE int32
	Wd80 int32
}{
	Wd70: 0,
	Wd70FE: 1,
	Wd80: 2,
}

// enum WdWrapTypeMerged
var WdWrapTypeMerged = struct {
	WdWrapMergeInline int32
	WdWrapMergeSquare int32
	WdWrapMergeTight int32
	WdWrapMergeBehind int32
	WdWrapMergeFront int32
	WdWrapMergeThrough int32
	WdWrapMergeTopBottom int32
}{
	WdWrapMergeInline: 0,
	WdWrapMergeSquare: 1,
	WdWrapMergeTight: 2,
	WdWrapMergeBehind: 3,
	WdWrapMergeFront: 4,
	WdWrapMergeThrough: 5,
	WdWrapMergeTopBottom: 6,
}

// enum WdRecoveryType
var WdRecoveryType = struct {
	WdPasteDefault int32
	WdSingleCellText int32
	WdSingleCellTable int32
	WdListContinueNumbering int32
	WdListRestartNumbering int32
	WdTableInsertAsRows int32
	WdTableAppendTable int32
	WdTableOriginalFormatting int32
	WdChartPicture int32
	WdChart int32
	WdChartLinked int32
	WdFormatOriginalFormatting int32
	WdFormatSurroundingFormattingWithEmphasis int32
	WdFormatPlainText int32
	WdTableOverwriteCells int32
	WdListCombineWithExistingList int32
	WdListDontMerge int32
	WdUseDestinationStylesRecovery int32
}{
	WdPasteDefault: 0,
	WdSingleCellText: 5,
	WdSingleCellTable: 6,
	WdListContinueNumbering: 7,
	WdListRestartNumbering: 8,
	WdTableInsertAsRows: 11,
	WdTableAppendTable: 10,
	WdTableOriginalFormatting: 12,
	WdChartPicture: 13,
	WdChart: 14,
	WdChartLinked: 15,
	WdFormatOriginalFormatting: 16,
	WdFormatSurroundingFormattingWithEmphasis: 20,
	WdFormatPlainText: 22,
	WdTableOverwriteCells: 23,
	WdListCombineWithExistingList: 24,
	WdListDontMerge: 25,
	WdUseDestinationStylesRecovery: 19,
}

// enum WdLineEndingType
var WdLineEndingType = struct {
	WdCRLF int32
	WdCROnly int32
	WdLFOnly int32
	WdLFCR int32
	WdLSPS int32
}{
	WdCRLF: 0,
	WdCROnly: 1,
	WdLFOnly: 2,
	WdLFCR: 3,
	WdLSPS: 4,
}

// enum WdStyleSheetLinkType
var WdStyleSheetLinkType = struct {
	WdStyleSheetLinkTypeLinked int32
	WdStyleSheetLinkTypeImported int32
}{
	WdStyleSheetLinkTypeLinked: 0,
	WdStyleSheetLinkTypeImported: 1,
}

// enum WdStyleSheetPrecedence
var WdStyleSheetPrecedence = struct {
	WdStyleSheetPrecedenceHigher int32
	WdStyleSheetPrecedenceLower int32
	WdStyleSheetPrecedenceHighest int32
	WdStyleSheetPrecedenceLowest int32
}{
	WdStyleSheetPrecedenceHigher: -1,
	WdStyleSheetPrecedenceLower: -2,
	WdStyleSheetPrecedenceHighest: 1,
	WdStyleSheetPrecedenceLowest: 0,
}

// enum WdEmailHTMLFidelity
var WdEmailHTMLFidelity = struct {
	WdEmailHTMLFidelityLow int32
	WdEmailHTMLFidelityMedium int32
	WdEmailHTMLFidelityHigh int32
}{
	WdEmailHTMLFidelityLow: 1,
	WdEmailHTMLFidelityMedium: 2,
	WdEmailHTMLFidelityHigh: 3,
}

// enum WdMailMergeMailFormat
var WdMailMergeMailFormat = struct {
	WdMailFormatPlainText int32
	WdMailFormatHTML int32
}{
	WdMailFormatPlainText: 0,
	WdMailFormatHTML: 1,
}

// enum WdMappedDataFields
var WdMappedDataFields = struct {
	WdUniqueIdentifier int32
	WdCourtesyTitle int32
	WdFirstName int32
	WdMiddleName int32
	WdLastName int32
	WdSuffix int32
	WdNickname int32
	WdJobTitle int32
	WdCompany int32
	WdAddress1 int32
	WdAddress2 int32
	WdCity int32
	WdState int32
	WdPostalCode int32
	WdCountryRegion int32
	WdBusinessPhone int32
	WdBusinessFax int32
	WdHomePhone int32
	WdHomeFax int32
	WdEmailAddress int32
	WdWebPageURL int32
	WdSpouseCourtesyTitle int32
	WdSpouseFirstName int32
	WdSpouseMiddleName int32
	WdSpouseLastName int32
	WdSpouseNickname int32
	WdRubyFirstName int32
	WdRubyLastName int32
	WdAddress3 int32
	WdDepartment int32
}{
	WdUniqueIdentifier: 1,
	WdCourtesyTitle: 2,
	WdFirstName: 3,
	WdMiddleName: 4,
	WdLastName: 5,
	WdSuffix: 6,
	WdNickname: 7,
	WdJobTitle: 8,
	WdCompany: 9,
	WdAddress1: 10,
	WdAddress2: 11,
	WdCity: 12,
	WdState: 13,
	WdPostalCode: 14,
	WdCountryRegion: 15,
	WdBusinessPhone: 16,
	WdBusinessFax: 17,
	WdHomePhone: 18,
	WdHomeFax: 19,
	WdEmailAddress: 20,
	WdWebPageURL: 21,
	WdSpouseCourtesyTitle: 22,
	WdSpouseFirstName: 23,
	WdSpouseMiddleName: 24,
	WdSpouseLastName: 25,
	WdSpouseNickname: 26,
	WdRubyFirstName: 27,
	WdRubyLastName: 28,
	WdAddress3: 29,
	WdDepartment: 30,
}

// enum WdConditionCode
var WdConditionCode = struct {
	WdFirstRow int32
	WdLastRow int32
	WdOddRowBanding int32
	WdEvenRowBanding int32
	WdFirstColumn int32
	WdLastColumn int32
	WdOddColumnBanding int32
	WdEvenColumnBanding int32
	WdNECell int32
	WdNWCell int32
	WdSECell int32
	WdSWCell int32
}{
	WdFirstRow: 0,
	WdLastRow: 1,
	WdOddRowBanding: 2,
	WdEvenRowBanding: 3,
	WdFirstColumn: 4,
	WdLastColumn: 5,
	WdOddColumnBanding: 6,
	WdEvenColumnBanding: 7,
	WdNECell: 8,
	WdNWCell: 9,
	WdSECell: 10,
	WdSWCell: 11,
}

// enum WdCompareTarget
var WdCompareTarget = struct {
	WdCompareTargetSelected int32
	WdCompareTargetCurrent int32
	WdCompareTargetNew int32
}{
	WdCompareTargetSelected: 0,
	WdCompareTargetCurrent: 1,
	WdCompareTargetNew: 2,
}

// enum WdMergeTarget
var WdMergeTarget = struct {
	WdMergeTargetSelected int32
	WdMergeTargetCurrent int32
	WdMergeTargetNew int32
}{
	WdMergeTargetSelected: 0,
	WdMergeTargetCurrent: 1,
	WdMergeTargetNew: 2,
}

// enum WdUseFormattingFrom
var WdUseFormattingFrom = struct {
	WdFormattingFromCurrent int32
	WdFormattingFromSelected int32
	WdFormattingFromPrompt int32
}{
	WdFormattingFromCurrent: 0,
	WdFormattingFromSelected: 1,
	WdFormattingFromPrompt: 2,
}

// enum WdRevisionsView
var WdRevisionsView = struct {
	WdRevisionsViewFinal int32
	WdRevisionsViewOriginal int32
}{
	WdRevisionsViewFinal: 0,
	WdRevisionsViewOriginal: 1,
}

// enum WdRevisionsMode
var WdRevisionsMode = struct {
	WdBalloonRevisions int32
	WdInLineRevisions int32
	WdMixedRevisions int32
}{
	WdBalloonRevisions: 0,
	WdInLineRevisions: 1,
	WdMixedRevisions: 2,
}

// enum WdRevisionsBalloonWidthType
var WdRevisionsBalloonWidthType = struct {
	WdBalloonWidthPercent int32
	WdBalloonWidthPoints int32
}{
	WdBalloonWidthPercent: 0,
	WdBalloonWidthPoints: 1,
}

// enum WdRevisionsBalloonPrintOrientation
var WdRevisionsBalloonPrintOrientation = struct {
	WdBalloonPrintOrientationAuto int32
	WdBalloonPrintOrientationPreserve int32
	WdBalloonPrintOrientationForceLandscape int32
}{
	WdBalloonPrintOrientationAuto: 0,
	WdBalloonPrintOrientationPreserve: 1,
	WdBalloonPrintOrientationForceLandscape: 2,
}

// enum WdRevisionsBalloonMargin
var WdRevisionsBalloonMargin = struct {
	WdLeftMargin int32
	WdRightMargin int32
}{
	WdLeftMargin: 0,
	WdRightMargin: 1,
}

// enum WdTaskPanes
var WdTaskPanes = struct {
	WdTaskPaneFormatting int32
	WdTaskPaneRevealFormatting int32
	WdTaskPaneMailMerge int32
	WdTaskPaneTranslate int32
	WdTaskPaneSearch int32
	WdTaskPaneXMLStructure int32
	WdTaskPaneDocumentProtection int32
	WdTaskPaneDocumentActions int32
	WdTaskPaneSharedWorkspace int32
	WdTaskPaneHelp int32
	WdTaskPaneResearch int32
	WdTaskPaneFaxService int32
	WdTaskPaneXMLDocument int32
	WdTaskPaneDocumentUpdates int32
	WdTaskPaneSignature int32
	WdTaskPaneStyleInspector int32
	WdTaskPaneDocumentManagement int32
	WdTaskPaneApplyStyles int32
	WdTaskPaneNav int32
	WdTaskPaneSelection int32
}{
	WdTaskPaneFormatting: 0,
	WdTaskPaneRevealFormatting: 1,
	WdTaskPaneMailMerge: 2,
	WdTaskPaneTranslate: 3,
	WdTaskPaneSearch: 4,
	WdTaskPaneXMLStructure: 5,
	WdTaskPaneDocumentProtection: 6,
	WdTaskPaneDocumentActions: 7,
	WdTaskPaneSharedWorkspace: 8,
	WdTaskPaneHelp: 9,
	WdTaskPaneResearch: 10,
	WdTaskPaneFaxService: 11,
	WdTaskPaneXMLDocument: 12,
	WdTaskPaneDocumentUpdates: 13,
	WdTaskPaneSignature: 14,
	WdTaskPaneStyleInspector: 15,
	WdTaskPaneDocumentManagement: 16,
	WdTaskPaneApplyStyles: 17,
	WdTaskPaneNav: 18,
	WdTaskPaneSelection: 19,
}

// enum WdShowFilter
var WdShowFilter = struct {
	WdShowFilterStylesAvailable int32
	WdShowFilterStylesInUse int32
	WdShowFilterStylesAll int32
	WdShowFilterFormattingInUse int32
	WdShowFilterFormattingAvailable int32
	WdShowFilterFormattingRecommended int32
}{
	WdShowFilterStylesAvailable: 0,
	WdShowFilterStylesInUse: 1,
	WdShowFilterStylesAll: 2,
	WdShowFilterFormattingInUse: 3,
	WdShowFilterFormattingAvailable: 4,
	WdShowFilterFormattingRecommended: 5,
}

// enum WdMergeSubType
var WdMergeSubType = struct {
	WdMergeSubTypeOther int32
	WdMergeSubTypeAccess int32
	WdMergeSubTypeOAL int32
	WdMergeSubTypeOLEDBWord int32
	WdMergeSubTypeWorks int32
	WdMergeSubTypeOLEDBText int32
	WdMergeSubTypeOutlook int32
	WdMergeSubTypeWord int32
	WdMergeSubTypeWord2000 int32
}{
	WdMergeSubTypeOther: 0,
	WdMergeSubTypeAccess: 1,
	WdMergeSubTypeOAL: 2,
	WdMergeSubTypeOLEDBWord: 3,
	WdMergeSubTypeWorks: 4,
	WdMergeSubTypeOLEDBText: 5,
	WdMergeSubTypeOutlook: 6,
	WdMergeSubTypeWord: 7,
	WdMergeSubTypeWord2000: 8,
}

// enum WdDocumentDirection
var WdDocumentDirection = struct {
	WdLeftToRight int32
	WdRightToLeft int32
}{
	WdLeftToRight: 0,
	WdRightToLeft: 1,
}

// enum WdLanguageID2000
var WdLanguageID2000 = struct {
	WdChineseHongKong int32
	WdChineseMacao int32
	WdEnglishTrinidad int32
}{
	WdChineseHongKong: 3076,
	WdChineseMacao: 5124,
	WdEnglishTrinidad: 11273,
}

// enum WdRectangleType
var WdRectangleType = struct {
	WdTextRectangle int32
	WdShapeRectangle int32
	WdMarkupRectangle int32
	WdMarkupRectangleButton int32
	WdPageBorderRectangle int32
	WdLineBetweenColumnRectangle int32
	WdSelection int32
	WdSystem int32
	WdMarkupRectangleArea int32
	WdReadingModeNavigation int32
	WdMarkupRectangleMoveMatch int32
	WdReadingModePanningArea int32
	WdMailNavArea int32
	WdDocumentControlRectangle int32
}{
	WdTextRectangle: 0,
	WdShapeRectangle: 1,
	WdMarkupRectangle: 2,
	WdMarkupRectangleButton: 3,
	WdPageBorderRectangle: 4,
	WdLineBetweenColumnRectangle: 5,
	WdSelection: 6,
	WdSystem: 7,
	WdMarkupRectangleArea: 8,
	WdReadingModeNavigation: 9,
	WdMarkupRectangleMoveMatch: 10,
	WdReadingModePanningArea: 11,
	WdMailNavArea: 12,
	WdDocumentControlRectangle: 13,
}

// enum WdLineType
var WdLineType = struct {
	WdTextLine int32
	WdTableRow int32
}{
	WdTextLine: 0,
	WdTableRow: 1,
}

// enum WdXMLNodeType
var WdXMLNodeType = struct {
	WdXMLNodeElement int32
	WdXMLNodeAttribute int32
}{
	WdXMLNodeElement: 1,
	WdXMLNodeAttribute: 2,
}

// enum WdXMLSelectionChangeReason
var WdXMLSelectionChangeReason = struct {
	WdXMLSelectionChangeReasonMove int32
	WdXMLSelectionChangeReasonInsert int32
	WdXMLSelectionChangeReasonDelete int32
}{
	WdXMLSelectionChangeReasonMove: 0,
	WdXMLSelectionChangeReasonInsert: 1,
	WdXMLSelectionChangeReasonDelete: 2,
}

// enum WdXMLNodeLevel
var WdXMLNodeLevel = struct {
	WdXMLNodeLevelInline int32
	WdXMLNodeLevelParagraph int32
	WdXMLNodeLevelRow int32
	WdXMLNodeLevelCell int32
}{
	WdXMLNodeLevelInline: 0,
	WdXMLNodeLevelParagraph: 1,
	WdXMLNodeLevelRow: 2,
	WdXMLNodeLevelCell: 3,
}

// enum WdSmartTagControlType
var WdSmartTagControlType = struct {
	WdControlSmartTag int32
	WdControlLink int32
	WdControlHelp int32
	WdControlHelpURL int32
	WdControlSeparator int32
	WdControlButton int32
	WdControlLabel int32
	WdControlImage int32
	WdControlCheckbox int32
	WdControlTextbox int32
	WdControlListbox int32
	WdControlCombo int32
	WdControlActiveX int32
	WdControlDocumentFragment int32
	WdControlDocumentFragmentURL int32
	WdControlRadioGroup int32
}{
	WdControlSmartTag: 1,
	WdControlLink: 2,
	WdControlHelp: 3,
	WdControlHelpURL: 4,
	WdControlSeparator: 5,
	WdControlButton: 6,
	WdControlLabel: 7,
	WdControlImage: 8,
	WdControlCheckbox: 9,
	WdControlTextbox: 10,
	WdControlListbox: 11,
	WdControlCombo: 12,
	WdControlActiveX: 13,
	WdControlDocumentFragment: 14,
	WdControlDocumentFragmentURL: 15,
	WdControlRadioGroup: 16,
}

// enum WdEditorType
var WdEditorType = struct {
	WdEditorEveryone int32
	WdEditorOwners int32
	WdEditorEditors int32
	WdEditorCurrent int32
}{
	WdEditorEveryone: -1,
	WdEditorOwners: -4,
	WdEditorEditors: -5,
	WdEditorCurrent: -6,
}

// enum WdXMLValidationStatus
var WdXMLValidationStatus = struct {
	WdXMLValidationStatusOK int32
	WdXMLValidationStatusCustom int32
}{
	WdXMLValidationStatusOK: 0,
	WdXMLValidationStatusCustom: -1072898048,
}

// enum WdStyleSort
var WdStyleSort = struct {
	WdStyleSortByName int32
	WdStyleSortRecommended int32
	WdStyleSortByFont int32
	WdStyleSortByBasedOn int32
	WdStyleSortByType int32
}{
	WdStyleSortByName: 0,
	WdStyleSortRecommended: 1,
	WdStyleSortByFont: 2,
	WdStyleSortByBasedOn: 3,
	WdStyleSortByType: 4,
}

// enum WdRemoveDocInfoType
var WdRemoveDocInfoType = struct {
	WdRDIComments int32
	WdRDIRevisions int32
	WdRDIVersions int32
	WdRDIRemovePersonalInformation int32
	WdRDIEmailHeader int32
	WdRDIRoutingSlip int32
	WdRDISendForReview int32
	WdRDIDocumentProperties int32
	WdRDITemplate int32
	WdRDIDocumentWorkspace int32
	WdRDIInkAnnotations int32
	WdRDIDocumentServerProperties int32
	WdRDIDocumentManagementPolicy int32
	WdRDIContentType int32
	WdRDIAll int32
}{
	WdRDIComments: 1,
	WdRDIRevisions: 2,
	WdRDIVersions: 3,
	WdRDIRemovePersonalInformation: 4,
	WdRDIEmailHeader: 5,
	WdRDIRoutingSlip: 6,
	WdRDISendForReview: 7,
	WdRDIDocumentProperties: 8,
	WdRDITemplate: 9,
	WdRDIDocumentWorkspace: 10,
	WdRDIInkAnnotations: 11,
	WdRDIDocumentServerProperties: 14,
	WdRDIDocumentManagementPolicy: 15,
	WdRDIContentType: 16,
	WdRDIAll: 99,
}

// enum WdCheckInVersionType
var WdCheckInVersionType = struct {
	WdCheckInMinorVersion int32
	WdCheckInMajorVersion int32
	WdCheckInOverwriteVersion int32
}{
	WdCheckInMinorVersion: 0,
	WdCheckInMajorVersion: 1,
	WdCheckInOverwriteVersion: 2,
}

// enum WdMoveToTextMark
var WdMoveToTextMark = struct {
	WdMoveToTextMarkNone int32
	WdMoveToTextMarkBold int32
	WdMoveToTextMarkItalic int32
	WdMoveToTextMarkUnderline int32
	WdMoveToTextMarkDoubleUnderline int32
	WdMoveToTextMarkColorOnly int32
	WdMoveToTextMarkStrikeThrough int32
	WdMoveToTextMarkDoubleStrikeThrough int32
}{
	WdMoveToTextMarkNone: 0,
	WdMoveToTextMarkBold: 1,
	WdMoveToTextMarkItalic: 2,
	WdMoveToTextMarkUnderline: 3,
	WdMoveToTextMarkDoubleUnderline: 4,
	WdMoveToTextMarkColorOnly: 5,
	WdMoveToTextMarkStrikeThrough: 6,
	WdMoveToTextMarkDoubleStrikeThrough: 7,
}

// enum WdMoveFromTextMark
var WdMoveFromTextMark = struct {
	WdMoveFromTextMarkHidden int32
	WdMoveFromTextMarkDoubleStrikeThrough int32
	WdMoveFromTextMarkStrikeThrough int32
	WdMoveFromTextMarkCaret int32
	WdMoveFromTextMarkPound int32
	WdMoveFromTextMarkNone int32
	WdMoveFromTextMarkBold int32
	WdMoveFromTextMarkItalic int32
	WdMoveFromTextMarkUnderline int32
	WdMoveFromTextMarkDoubleUnderline int32
	WdMoveFromTextMarkColorOnly int32
}{
	WdMoveFromTextMarkHidden: 0,
	WdMoveFromTextMarkDoubleStrikeThrough: 1,
	WdMoveFromTextMarkStrikeThrough: 2,
	WdMoveFromTextMarkCaret: 3,
	WdMoveFromTextMarkPound: 4,
	WdMoveFromTextMarkNone: 5,
	WdMoveFromTextMarkBold: 6,
	WdMoveFromTextMarkItalic: 7,
	WdMoveFromTextMarkUnderline: 8,
	WdMoveFromTextMarkDoubleUnderline: 9,
	WdMoveFromTextMarkColorOnly: 10,
}

// enum WdOMathFunctionType
var WdOMathFunctionType = struct {
	WdOMathFunctionAcc int32
	WdOMathFunctionBar int32
	WdOMathFunctionBox int32
	WdOMathFunctionBorderBox int32
	WdOMathFunctionDelim int32
	WdOMathFunctionEqArray int32
	WdOMathFunctionFrac int32
	WdOMathFunctionFunc int32
	WdOMathFunctionGroupChar int32
	WdOMathFunctionLimLow int32
	WdOMathFunctionLimUpp int32
	WdOMathFunctionMat int32
	WdOMathFunctionNary int32
	WdOMathFunctionPhantom int32
	WdOMathFunctionScrPre int32
	WdOMathFunctionRad int32
	WdOMathFunctionScrSub int32
	WdOMathFunctionScrSubSup int32
	WdOMathFunctionScrSup int32
	WdOMathFunctionText int32
	WdOMathFunctionNormalText int32
	WdOMathFunctionLiteralText int32
}{
	WdOMathFunctionAcc: 1,
	WdOMathFunctionBar: 2,
	WdOMathFunctionBox: 3,
	WdOMathFunctionBorderBox: 4,
	WdOMathFunctionDelim: 5,
	WdOMathFunctionEqArray: 6,
	WdOMathFunctionFrac: 7,
	WdOMathFunctionFunc: 8,
	WdOMathFunctionGroupChar: 9,
	WdOMathFunctionLimLow: 10,
	WdOMathFunctionLimUpp: 11,
	WdOMathFunctionMat: 12,
	WdOMathFunctionNary: 13,
	WdOMathFunctionPhantom: 14,
	WdOMathFunctionScrPre: 15,
	WdOMathFunctionRad: 16,
	WdOMathFunctionScrSub: 17,
	WdOMathFunctionScrSubSup: 18,
	WdOMathFunctionScrSup: 19,
	WdOMathFunctionText: 20,
	WdOMathFunctionNormalText: 21,
	WdOMathFunctionLiteralText: 22,
}

// enum WdOMathHorizAlignType
var WdOMathHorizAlignType = struct {
	WdOMathHorizAlignCenter int32
	WdOMathHorizAlignLeft int32
	WdOMathHorizAlignRight int32
}{
	WdOMathHorizAlignCenter: 0,
	WdOMathHorizAlignLeft: 1,
	WdOMathHorizAlignRight: 2,
}

// enum WdOMathVertAlignType
var WdOMathVertAlignType = struct {
	WdOMathVertAlignCenter int32
	WdOMathVertAlignTop int32
	WdOMathVertAlignBottom int32
}{
	WdOMathVertAlignCenter: 0,
	WdOMathVertAlignTop: 1,
	WdOMathVertAlignBottom: 2,
}

// enum WdOMathFracType
var WdOMathFracType = struct {
	WdOMathFracBar int32
	WdOMathFracNoBar int32
	WdOMathFracSkw int32
	WdOMathFracLin int32
}{
	WdOMathFracBar: 0,
	WdOMathFracNoBar: 1,
	WdOMathFracSkw: 2,
	WdOMathFracLin: 3,
}

// enum WdOMathSpacingRule
var WdOMathSpacingRule = struct {
	WdOMathSpacingSingle int32
	WdOMathSpacing1pt5 int32
	WdOMathSpacingDouble int32
	WdOMathSpacingExactly int32
	WdOMathSpacingMultiple int32
}{
	WdOMathSpacingSingle: 0,
	WdOMathSpacing1pt5: 1,
	WdOMathSpacingDouble: 2,
	WdOMathSpacingExactly: 3,
	WdOMathSpacingMultiple: 4,
}

// enum WdOMathType
var WdOMathType = struct {
	WdOMathDisplay int32
	WdOMathInline int32
}{
	WdOMathDisplay: 0,
	WdOMathInline: 1,
}

// enum WdOMathShapeType
var WdOMathShapeType = struct {
	WdOMathShapeCentered int32
	WdOMathShapeMatch int32
}{
	WdOMathShapeCentered: 0,
	WdOMathShapeMatch: 1,
}

// enum WdOMathJc
var WdOMathJc = struct {
	WdOMathJcCenterGroup int32
	WdOMathJcCenter int32
	WdOMathJcLeft int32
	WdOMathJcRight int32
	WdOMathJcInline int32
}{
	WdOMathJcCenterGroup: 1,
	WdOMathJcCenter: 2,
	WdOMathJcLeft: 3,
	WdOMathJcRight: 4,
	WdOMathJcInline: 7,
}

// enum WdOMathBreakBin
var WdOMathBreakBin = struct {
	WdOMathBreakBinBefore int32
	WdOMathBreakBinAfter int32
	WdOMathBreakBinRepeat int32
}{
	WdOMathBreakBinBefore: 0,
	WdOMathBreakBinAfter: 1,
	WdOMathBreakBinRepeat: 2,
}

// enum WdOMathBreakSub
var WdOMathBreakSub = struct {
	WdOMathBreakSubMinusMinus int32
	WdOMathBreakSubPlusMinus int32
	WdOMathBreakSubMinusPlus int32
}{
	WdOMathBreakSubMinusMinus: 0,
	WdOMathBreakSubPlusMinus: 1,
	WdOMathBreakSubMinusPlus: 2,
}

// enum WdReadingLayoutMargin
var WdReadingLayoutMargin = struct {
	WdAutomaticMargin int32
	WdSuppressMargin int32
	WdFullMargin int32
}{
	WdAutomaticMargin: 0,
	WdSuppressMargin: 1,
	WdFullMargin: 2,
}

// enum WdContentControlType
var WdContentControlType = struct {
	WdContentControlRichText int32
	WdContentControlText int32
	WdContentControlPicture int32
	WdContentControlComboBox int32
	WdContentControlDropdownList int32
	WdContentControlBuildingBlockGallery int32
	WdContentControlDate int32
	WdContentControlGroup int32
	WdContentControlCheckBox int32
}{
	WdContentControlRichText: 0,
	WdContentControlText: 1,
	WdContentControlPicture: 2,
	WdContentControlComboBox: 3,
	WdContentControlDropdownList: 4,
	WdContentControlBuildingBlockGallery: 5,
	WdContentControlDate: 6,
	WdContentControlGroup: 7,
	WdContentControlCheckBox: 8,
}

// enum WdCompareDestination
var WdCompareDestination = struct {
	WdCompareDestinationOriginal int32
	WdCompareDestinationRevised int32
	WdCompareDestinationNew int32
}{
	WdCompareDestinationOriginal: 0,
	WdCompareDestinationRevised: 1,
	WdCompareDestinationNew: 2,
}

// enum WdGranularity
var WdGranularity = struct {
	WdGranularityCharLevel int32
	WdGranularityWordLevel int32
}{
	WdGranularityCharLevel: 0,
	WdGranularityWordLevel: 1,
}

// enum WdMergeFormatFrom
var WdMergeFormatFrom = struct {
	WdMergeFormatFromOriginal int32
	WdMergeFormatFromRevised int32
	WdMergeFormatFromPrompt int32
}{
	WdMergeFormatFromOriginal: 0,
	WdMergeFormatFromRevised: 1,
	WdMergeFormatFromPrompt: 2,
}

// enum WdShowSourceDocuments
var WdShowSourceDocuments = struct {
	WdShowSourceDocumentsNone int32
	WdShowSourceDocumentsOriginal int32
	WdShowSourceDocumentsRevised int32
	WdShowSourceDocumentsBoth int32
}{
	WdShowSourceDocumentsNone: 0,
	WdShowSourceDocumentsOriginal: 1,
	WdShowSourceDocumentsRevised: 2,
	WdShowSourceDocumentsBoth: 3,
}

// enum WdPasteOptions
var WdPasteOptions = struct {
	WdKeepSourceFormatting int32
	WdMatchDestinationFormatting int32
	WdKeepTextOnly int32
	WdUseDestinationStyles int32
}{
	WdKeepSourceFormatting: 0,
	WdMatchDestinationFormatting: 1,
	WdKeepTextOnly: 2,
	WdUseDestinationStyles: 3,
}

// enum WdBuildingBlockTypes
var WdBuildingBlockTypes = struct {
	WdTypeQuickParts int32
	WdTypeCoverPage int32
	WdTypeEquations int32
	WdTypeFooters int32
	WdTypeHeaders int32
	WdTypePageNumber int32
	WdTypeTables int32
	WdTypeWatermarks int32
	WdTypeAutoText int32
	WdTypeTextBox int32
	WdTypePageNumberTop int32
	WdTypePageNumberBottom int32
	WdTypePageNumberPage int32
	WdTypeTableOfContents int32
	WdTypeCustomQuickParts int32
	WdTypeCustomCoverPage int32
	WdTypeCustomEquations int32
	WdTypeCustomFooters int32
	WdTypeCustomHeaders int32
	WdTypeCustomPageNumber int32
	WdTypeCustomTables int32
	WdTypeCustomWatermarks int32
	WdTypeCustomAutoText int32
	WdTypeCustomTextBox int32
	WdTypeCustomPageNumberTop int32
	WdTypeCustomPageNumberBottom int32
	WdTypeCustomPageNumberPage int32
	WdTypeCustomTableOfContents int32
	WdTypeCustom1 int32
	WdTypeCustom2 int32
	WdTypeCustom3 int32
	WdTypeCustom4 int32
	WdTypeCustom5 int32
	WdTypeBibliography int32
	WdTypeCustomBibliography int32
}{
	WdTypeQuickParts: 1,
	WdTypeCoverPage: 2,
	WdTypeEquations: 3,
	WdTypeFooters: 4,
	WdTypeHeaders: 5,
	WdTypePageNumber: 6,
	WdTypeTables: 7,
	WdTypeWatermarks: 8,
	WdTypeAutoText: 9,
	WdTypeTextBox: 10,
	WdTypePageNumberTop: 11,
	WdTypePageNumberBottom: 12,
	WdTypePageNumberPage: 13,
	WdTypeTableOfContents: 14,
	WdTypeCustomQuickParts: 15,
	WdTypeCustomCoverPage: 16,
	WdTypeCustomEquations: 17,
	WdTypeCustomFooters: 18,
	WdTypeCustomHeaders: 19,
	WdTypeCustomPageNumber: 20,
	WdTypeCustomTables: 21,
	WdTypeCustomWatermarks: 22,
	WdTypeCustomAutoText: 23,
	WdTypeCustomTextBox: 24,
	WdTypeCustomPageNumberTop: 25,
	WdTypeCustomPageNumberBottom: 26,
	WdTypeCustomPageNumberPage: 27,
	WdTypeCustomTableOfContents: 28,
	WdTypeCustom1: 29,
	WdTypeCustom2: 30,
	WdTypeCustom3: 31,
	WdTypeCustom4: 32,
	WdTypeCustom5: 33,
	WdTypeBibliography: 34,
	WdTypeCustomBibliography: 35,
}

// enum WdAlignmentTabRelative
var WdAlignmentTabRelative = struct {
	WdMargin int32
	WdIndent int32
}{
	WdMargin: 0,
	WdIndent: 1,
}

// enum WdAlignmentTabAlignment
var WdAlignmentTabAlignment = struct {
	WdLeft int32
	WdCenter int32
	WdRight int32
}{
	WdLeft: 0,
	WdCenter: 1,
	WdRight: 2,
}

// enum WdCellColor
var WdCellColor = struct {
	WdCellColorByAuthor int32
	WdCellColorNoHighlight int32
	WdCellColorPink int32
	WdCellColorLightBlue int32
	WdCellColorLightYellow int32
	WdCellColorLightPurple int32
	WdCellColorLightOrange int32
	WdCellColorLightGreen int32
	WdCellColorLightGray int32
}{
	WdCellColorByAuthor: -1,
	WdCellColorNoHighlight: 0,
	WdCellColorPink: 1,
	WdCellColorLightBlue: 2,
	WdCellColorLightYellow: 3,
	WdCellColorLightPurple: 4,
	WdCellColorLightOrange: 5,
	WdCellColorLightGreen: 6,
	WdCellColorLightGray: 7,
}

// enum WdTextboxTightWrap
var WdTextboxTightWrap = struct {
	WdTightNone int32
	WdTightAll int32
	WdTightFirstAndLastLines int32
	WdTightFirstLineOnly int32
	WdTightLastLineOnly int32
}{
	WdTightNone: 0,
	WdTightAll: 1,
	WdTightFirstAndLastLines: 2,
	WdTightFirstLineOnly: 3,
	WdTightLastLineOnly: 4,
}

// enum WdShapePositionRelative
var WdShapePositionRelative = struct {
	WdShapePositionRelativeNone int32
}{
	WdShapePositionRelativeNone: -999999,
}

// enum WdShapeSizeRelative
var WdShapeSizeRelative = struct {
	WdShapeSizeRelativeNone int32
}{
	WdShapeSizeRelativeNone: -999999,
}

// enum WdRelativeHorizontalSize
var WdRelativeHorizontalSize = struct {
	WdRelativeHorizontalSizeMargin int32
	WdRelativeHorizontalSizePage int32
	WdRelativeHorizontalSizeLeftMarginArea int32
	WdRelativeHorizontalSizeRightMarginArea int32
	WdRelativeHorizontalSizeInnerMarginArea int32
	WdRelativeHorizontalSizeOuterMarginArea int32
}{
	WdRelativeHorizontalSizeMargin: 0,
	WdRelativeHorizontalSizePage: 1,
	WdRelativeHorizontalSizeLeftMarginArea: 2,
	WdRelativeHorizontalSizeRightMarginArea: 3,
	WdRelativeHorizontalSizeInnerMarginArea: 4,
	WdRelativeHorizontalSizeOuterMarginArea: 5,
}

// enum WdRelativeVerticalSize
var WdRelativeVerticalSize = struct {
	WdRelativeVerticalSizeMargin int32
	WdRelativeVerticalSizePage int32
	WdRelativeVerticalSizeTopMarginArea int32
	WdRelativeVerticalSizeBottomMarginArea int32
	WdRelativeVerticalSizeInnerMarginArea int32
	WdRelativeVerticalSizeOuterMarginArea int32
}{
	WdRelativeVerticalSizeMargin: 0,
	WdRelativeVerticalSizePage: 1,
	WdRelativeVerticalSizeTopMarginArea: 2,
	WdRelativeVerticalSizeBottomMarginArea: 3,
	WdRelativeVerticalSizeInnerMarginArea: 4,
	WdRelativeVerticalSizeOuterMarginArea: 5,
}

// enum WdThemeColorIndex
var WdThemeColorIndex = struct {
	WdNotThemeColor int32
	WdThemeColorMainDark1 int32
	WdThemeColorMainLight1 int32
	WdThemeColorMainDark2 int32
	WdThemeColorMainLight2 int32
	WdThemeColorAccent1 int32
	WdThemeColorAccent2 int32
	WdThemeColorAccent3 int32
	WdThemeColorAccent4 int32
	WdThemeColorAccent5 int32
	WdThemeColorAccent6 int32
	WdThemeColorHyperlink int32
	WdThemeColorHyperlinkFollowed int32
	WdThemeColorBackground1 int32
	WdThemeColorText1 int32
	WdThemeColorBackground2 int32
	WdThemeColorText2 int32
}{
	WdNotThemeColor: -1,
	WdThemeColorMainDark1: 0,
	WdThemeColorMainLight1: 1,
	WdThemeColorMainDark2: 2,
	WdThemeColorMainLight2: 3,
	WdThemeColorAccent1: 4,
	WdThemeColorAccent2: 5,
	WdThemeColorAccent3: 6,
	WdThemeColorAccent4: 7,
	WdThemeColorAccent5: 8,
	WdThemeColorAccent6: 9,
	WdThemeColorHyperlink: 10,
	WdThemeColorHyperlinkFollowed: 11,
	WdThemeColorBackground1: 12,
	WdThemeColorText1: 13,
	WdThemeColorBackground2: 14,
	WdThemeColorText2: 15,
}

// enum WdExportFormat
var WdExportFormat = struct {
	WdExportFormatPDF int32
	WdExportFormatXPS int32
}{
	WdExportFormatPDF: 17,
	WdExportFormatXPS: 18,
}

// enum WdExportOptimizeFor
var WdExportOptimizeFor = struct {
	WdExportOptimizeForPrint int32
	WdExportOptimizeForOnScreen int32
}{
	WdExportOptimizeForPrint: 0,
	WdExportOptimizeForOnScreen: 1,
}

// enum WdExportCreateBookmarks
var WdExportCreateBookmarks = struct {
	WdExportCreateNoBookmarks int32
	WdExportCreateHeadingBookmarks int32
	WdExportCreateWordBookmarks int32
}{
	WdExportCreateNoBookmarks: 0,
	WdExportCreateHeadingBookmarks: 1,
	WdExportCreateWordBookmarks: 2,
}

// enum WdExportItem
var WdExportItem = struct {
	WdExportDocumentContent int32
	WdExportDocumentWithMarkup int32
}{
	WdExportDocumentContent: 0,
	WdExportDocumentWithMarkup: 7,
}

// enum WdExportRange
var WdExportRange = struct {
	WdExportAllDocument int32
	WdExportSelection int32
	WdExportCurrentPage int32
	WdExportFromTo int32
}{
	WdExportAllDocument: 0,
	WdExportSelection: 1,
	WdExportCurrentPage: 2,
	WdExportFromTo: 3,
}

// enum WdFrenchSpeller
var WdFrenchSpeller = struct {
	WdFrenchBoth int32
	WdFrenchPreReform int32
	WdFrenchPostReform int32
}{
	WdFrenchBoth: 0,
	WdFrenchPreReform: 1,
	WdFrenchPostReform: 2,
}

// enum WdDocPartInsertOptions
var WdDocPartInsertOptions = struct {
	WdInsertContent int32
	WdInsertParagraph int32
	WdInsertPage int32
}{
	WdInsertContent: 0,
	WdInsertParagraph: 1,
	WdInsertPage: 2,
}

// enum WdContentControlDateStorageFormat
var WdContentControlDateStorageFormat = struct {
	WdContentControlDateStorageText int32
	WdContentControlDateStorageDate int32
	WdContentControlDateStorageDateTime int32
}{
	WdContentControlDateStorageText: 0,
	WdContentControlDateStorageDate: 1,
	WdContentControlDateStorageDateTime: 2,
}

// enum XlChartSplitType
var XlChartSplitType = struct {
	XlSplitByPosition int32
	XlSplitByPercentValue int32
	XlSplitByCustomSplit int32
	XlSplitByValue int32
}{
	XlSplitByPosition: 1,
	XlSplitByPercentValue: 3,
	XlSplitByCustomSplit: 4,
	XlSplitByValue: 2,
}

// enum XlSizeRepresents
var XlSizeRepresents = struct {
	XlSizeIsWidth int32
	XlSizeIsArea int32
}{
	XlSizeIsWidth: 2,
	XlSizeIsArea: 1,
}

// enum XlAxisGroup
var XlAxisGroup = struct {
	XlPrimary int32
	XlSecondary int32
}{
	XlPrimary: 1,
	XlSecondary: 2,
}

// enum XlBackground
var XlBackground = struct {
	XlBackgroundAutomatic int32
	XlBackgroundOpaque int32
	XlBackgroundTransparent int32
}{
	XlBackgroundAutomatic: -4105,
	XlBackgroundOpaque: 3,
	XlBackgroundTransparent: 2,
}

// enum XlChartGallery
var XlChartGallery = struct {
	XlBuiltIn int32
	XlUserDefined int32
	XlAnyGallery int32
}{
	XlBuiltIn: 21,
	XlUserDefined: 22,
	XlAnyGallery: 23,
}

// enum XlChartPicturePlacement
var XlChartPicturePlacement = struct {
	XlSides int32
	XlEnd int32
	XlEndSides int32
	XlFront int32
	XlFrontSides int32
	XlFrontEnd int32
	XlAllFaces int32
}{
	XlSides: 1,
	XlEnd: 2,
	XlEndSides: 3,
	XlFront: 4,
	XlFrontSides: 5,
	XlFrontEnd: 6,
	XlAllFaces: 7,
}

// enum XlDataLabelSeparator
var XlDataLabelSeparator = struct {
	XlDataLabelSeparatorDefault int32
}{
	XlDataLabelSeparatorDefault: 1,
}

// enum XlPattern
var XlPattern = struct {
	XlPatternAutomatic int32
	XlPatternChecker int32
	XlPatternCrissCross int32
	XlPatternDown int32
	XlPatternGray16 int32
	XlPatternGray25 int32
	XlPatternGray50 int32
	XlPatternGray75 int32
	XlPatternGray8 int32
	XlPatternGrid int32
	XlPatternHorizontal int32
	XlPatternLightDown int32
	XlPatternLightHorizontal int32
	XlPatternLightUp int32
	XlPatternLightVertical int32
	XlPatternNone int32
	XlPatternSemiGray75 int32
	XlPatternSolid int32
	XlPatternUp int32
	XlPatternVertical int32
	XlPatternLinearGradient int32
	XlPatternRectangularGradient int32
}{
	XlPatternAutomatic: -4105,
	XlPatternChecker: 9,
	XlPatternCrissCross: 16,
	XlPatternDown: -4121,
	XlPatternGray16: 17,
	XlPatternGray25: -4124,
	XlPatternGray50: -4125,
	XlPatternGray75: -4126,
	XlPatternGray8: 18,
	XlPatternGrid: 15,
	XlPatternHorizontal: -4128,
	XlPatternLightDown: 13,
	XlPatternLightHorizontal: 11,
	XlPatternLightUp: 14,
	XlPatternLightVertical: 12,
	XlPatternNone: -4142,
	XlPatternSemiGray75: 10,
	XlPatternSolid: 1,
	XlPatternUp: -4162,
	XlPatternVertical: -4166,
	XlPatternLinearGradient: 4000,
	XlPatternRectangularGradient: 4001,
}

// enum XlPictureAppearance
var XlPictureAppearance = struct {
	XlPrinter int32
	XlScreen int32
}{
	XlPrinter: 2,
	XlScreen: 1,
}

// enum XlCopyPictureFormat
var XlCopyPictureFormat = struct {
	XlBitmap int32
	XlPicture int32
}{
	XlBitmap: 2,
	XlPicture: -4147,
}

// enum XlRgbColor
var XlRgbColor = struct {
	XlAliceBlue int32
	XlAntiqueWhite int32
	XlAqua int32
	XlAquamarine int32
	XlAzure int32
	XlBeige int32
	XlBisque int32
	XlBlack int32
	XlBlanchedAlmond int32
	XlBlue int32
	XlBlueViolet int32
	XlBrown int32
	XlBurlyWood int32
	XlCadetBlue int32
	XlChartreuse int32
	XlCoral int32
	XlCornflowerBlue int32
	XlCornsilk int32
	XlCrimson int32
	XlDarkBlue int32
	XlDarkCyan int32
	XlDarkGoldenrod int32
	XlDarkGreen int32
	XlDarkGray int32
	XlDarkGrey int32
	XlDarkKhaki int32
	XlDarkMagenta int32
	XlDarkOliveGreen int32
	XlDarkOrange int32
	XlDarkOrchid int32
	XlDarkRed int32
	XlDarkSalmon int32
	XlDarkSeaGreen int32
	XlDarkSlateBlue int32
	XlDarkSlateGray int32
	XlDarkSlateGrey int32
	XlDarkTurquoise int32
	XlDarkViolet int32
	XlDeepPink int32
	XlDeepSkyBlue int32
	XlDimGray int32
	XlDimGrey int32
	XlDodgerBlue int32
	XlFireBrick int32
	XlFloralWhite int32
	XlForestGreen int32
	XlFuchsia int32
	XlGainsboro int32
	XlGhostWhite int32
	XlGold int32
	XlGoldenrod int32
	XlGray int32
	XlGreen int32
	XlGrey int32
	XlGreenYellow int32
	XlHoneydew int32
	XlHotPink int32
	XlIndianRed int32
	XlIndigo int32
	XlIvory int32
	XlKhaki int32
	XlLavender int32
	XlLavenderBlush int32
	XlLawnGreen int32
	XlLemonChiffon int32
	XlLightBlue int32
	XlLightCoral int32
	XlLightCyan int32
	XlLightGoldenrodYellow int32
	XlLightGray int32
	XlLightGreen int32
	XlLightGrey int32
	XlLightPink int32
	XlLightSalmon int32
	XlLightSeaGreen int32
	XlLightSkyBlue int32
	XlLightSlateGray int32
	XlLightSlateGrey int32
	XlLightSteelBlue int32
	XlLightYellow int32
	XlLime int32
	XlLimeGreen int32
	XlLinen int32
	XlMaroon int32
	XlMediumAquamarine int32
	XlMediumBlue int32
	XlMediumOrchid int32
	XlMediumPurple int32
	XlMediumSeaGreen int32
	XlMediumSlateBlue int32
	XlMediumSpringGreen int32
	XlMediumTurquoise int32
	XlMediumVioletRed int32
	XlMidnightBlue int32
	XlMintCream int32
	XlMistyRose int32
	XlMoccasin int32
	XlNavajoWhite int32
	XlNavy int32
	XlNavyBlue int32
	XlOldLace int32
	XlOlive int32
	XlOliveDrab int32
	XlOrange int32
	XlOrangeRed int32
	XlOrchid int32
	XlPaleGoldenrod int32
	XlPaleGreen int32
	XlPaleTurquoise int32
	XlPaleVioletRed int32
	XlPapayaWhip int32
	XlPeachPuff int32
	XlPeru int32
	XlPink int32
	XlPlum int32
	XlPowderBlue int32
	XlPurple int32
	XlRed int32
	XlRosyBrown int32
	XlRoyalBlue int32
	XlSalmon int32
	XlSandyBrown int32
	XlSeaGreen int32
	XlSeashell int32
	XlSienna int32
	XlSilver int32
	XlSkyBlue int32
	XlSlateBlue int32
	XlSlateGray int32
	XlSlateGrey int32
	XlSnow int32
	XlSpringGreen int32
	XlSteelBlue int32
	XlTan int32
	XlTeal int32
	XlThistle int32
	XlTomato int32
	XlTurquoise int32
	XlYellow int32
	XlYellowGreen int32
	XlViolet int32
	XlWheat int32
	XlWhite int32
	XlWhiteSmoke int32
}{
	XlAliceBlue: 16775408,
	XlAntiqueWhite: 14150650,
	XlAqua: 16776960,
	XlAquamarine: 13959039,
	XlAzure: 16777200,
	XlBeige: 14480885,
	XlBisque: 12903679,
	XlBlack: 0,
	XlBlanchedAlmond: 13495295,
	XlBlue: 16711680,
	XlBlueViolet: 14822282,
	XlBrown: 2763429,
	XlBurlyWood: 8894686,
	XlCadetBlue: 10526303,
	XlChartreuse: 65407,
	XlCoral: 5275647,
	XlCornflowerBlue: 15570276,
	XlCornsilk: 14481663,
	XlCrimson: 3937500,
	XlDarkBlue: 9109504,
	XlDarkCyan: 9145088,
	XlDarkGoldenrod: 755384,
	XlDarkGreen: 25600,
	XlDarkGray: 11119017,
	XlDarkGrey: 11119017,
	XlDarkKhaki: 7059389,
	XlDarkMagenta: 9109643,
	XlDarkOliveGreen: 3107669,
	XlDarkOrange: 36095,
	XlDarkOrchid: 13382297,
	XlDarkRed: 139,
	XlDarkSalmon: 8034025,
	XlDarkSeaGreen: 9419919,
	XlDarkSlateBlue: 9125192,
	XlDarkSlateGray: 5197615,
	XlDarkSlateGrey: 5197615,
	XlDarkTurquoise: 13749760,
	XlDarkViolet: 13828244,
	XlDeepPink: 9639167,
	XlDeepSkyBlue: 16760576,
	XlDimGray: 6908265,
	XlDimGrey: 6908265,
	XlDodgerBlue: 16748574,
	XlFireBrick: 2237106,
	XlFloralWhite: 15792895,
	XlForestGreen: 2263842,
	XlFuchsia: 16711935,
	XlGainsboro: 14474460,
	XlGhostWhite: 16775416,
	XlGold: 55295,
	XlGoldenrod: 2139610,
	XlGray: 8421504,
	XlGreen: 32768,
	XlGrey: 8421504,
	XlGreenYellow: 3145645,
	XlHoneydew: 15794160,
	XlHotPink: 11823615,
	XlIndianRed: 6053069,
	XlIndigo: 8519755,
	XlIvory: 15794175,
	XlKhaki: 9234160,
	XlLavender: 16443110,
	XlLavenderBlush: 16118015,
	XlLawnGreen: 64636,
	XlLemonChiffon: 13499135,
	XlLightBlue: 15128749,
	XlLightCoral: 8421616,
	XlLightCyan: 9145088,
	XlLightGoldenrodYellow: 13826810,
	XlLightGray: 13882323,
	XlLightGreen: 9498256,
	XlLightGrey: 13882323,
	XlLightPink: 12695295,
	XlLightSalmon: 8036607,
	XlLightSeaGreen: 11186720,
	XlLightSkyBlue: 16436871,
	XlLightSlateGray: 10061943,
	XlLightSlateGrey: 10061943,
	XlLightSteelBlue: 14599344,
	XlLightYellow: 14745599,
	XlLime: 65280,
	XlLimeGreen: 3329330,
	XlLinen: 15134970,
	XlMaroon: 128,
	XlMediumAquamarine: 11206502,
	XlMediumBlue: 13434880,
	XlMediumOrchid: 13850042,
	XlMediumPurple: 14381203,
	XlMediumSeaGreen: 7451452,
	XlMediumSlateBlue: 15624315,
	XlMediumSpringGreen: 10156544,
	XlMediumTurquoise: 13422920,
	XlMediumVioletRed: 8721863,
	XlMidnightBlue: 7346457,
	XlMintCream: 16449525,
	XlMistyRose: 14804223,
	XlMoccasin: 11920639,
	XlNavajoWhite: 11394815,
	XlNavy: 8388608,
	XlNavyBlue: 8388608,
	XlOldLace: 15136253,
	XlOlive: 32896,
	XlOliveDrab: 2330219,
	XlOrange: 42495,
	XlOrangeRed: 17919,
	XlOrchid: 14053594,
	XlPaleGoldenrod: 7071982,
	XlPaleGreen: 10025880,
	XlPaleTurquoise: 15658671,
	XlPaleVioletRed: 9662683,
	XlPapayaWhip: 14020607,
	XlPeachPuff: 12180223,
	XlPeru: 4163021,
	XlPink: 13353215,
	XlPlum: 14524637,
	XlPowderBlue: 15130800,
	XlPurple: 8388736,
	XlRed: 255,
	XlRosyBrown: 9408444,
	XlRoyalBlue: 14772545,
	XlSalmon: 7504122,
	XlSandyBrown: 6333684,
	XlSeaGreen: 5737262,
	XlSeashell: 15660543,
	XlSienna: 2970272,
	XlSilver: 12632256,
	XlSkyBlue: 15453831,
	XlSlateBlue: 13458026,
	XlSlateGray: 9470064,
	XlSlateGrey: 9470064,
	XlSnow: 16448255,
	XlSpringGreen: 8388352,
	XlSteelBlue: 11829830,
	XlTan: 9221330,
	XlTeal: 8421376,
	XlThistle: 14204888,
	XlTomato: 4678655,
	XlTurquoise: 13688896,
	XlYellow: 65535,
	XlYellowGreen: 3329434,
	XlViolet: 15631086,
	XlWheat: 11788021,
	XlWhite: 16777215,
	XlWhiteSmoke: 16119285,
}

// enum XlConstants
var XlConstants = struct {
	XlAutomatic int32
	XlCombination int32
	XlCustom int32
	XlBar int32
	XlColumn int32
	Xl3DBar int32
	Xl3DSurface int32
	XlDefaultAutoFormat int32
	XlNone int32
	XlAbove int32
	XlBelow int32
	XlBoth int32
	XlBottom int32
	XlCenter int32
	XlChecker int32
	XlCircle int32
	XlCorner int32
	XlCrissCross int32
	XlCross int32
	XlDiamond int32
	XlDistributed int32
	XlFill int32
	XlFixedValue int32
	XlGeneral int32
	XlGray16 int32
	XlGray25 int32
	XlGray50 int32
	XlGray75 int32
	XlGray8 int32
	XlGrid int32
	XlHigh int32
	XlInside int32
	XlJustify int32
	XlLeft int32
	XlLightDown int32
	XlLightHorizontal int32
	XlLightUp int32
	XlLightVertical int32
	XlLow int32
	XlMaximum int32
	XlMinimum int32
	XlMinusValues int32
	XlNextToAxis int32
	XlOpaque int32
	XlOutside int32
	XlPercent int32
	XlPlus int32
	XlPlusValues int32
	XlRight int32
	XlScale int32
	XlSemiGray75 int32
	XlShowLabel int32
	XlShowLabelAndPercent int32
	XlShowPercent int32
	XlShowValue int32
	XlSingle int32
	XlSolid int32
	XlSquare int32
	XlStar int32
	XlStError int32
	XlTop int32
	XlTransparent int32
	XlTriangle int32
}{
	XlAutomatic: -4105,
	XlCombination: -4111,
	XlCustom: -4114,
	XlBar: 2,
	XlColumn: 3,
	Xl3DBar: -4099,
	Xl3DSurface: -4103,
	XlDefaultAutoFormat: -1,
	XlNone: -4142,
	XlAbove: 0,
	XlBelow: 1,
	XlBoth: 1,
	XlBottom: -4107,
	XlCenter: -4108,
	XlChecker: 9,
	XlCircle: 8,
	XlCorner: 2,
	XlCrissCross: 16,
	XlCross: 4,
	XlDiamond: 2,
	XlDistributed: -4117,
	XlFill: 5,
	XlFixedValue: 1,
	XlGeneral: 1,
	XlGray16: 17,
	XlGray25: -4124,
	XlGray50: -4125,
	XlGray75: -4126,
	XlGray8: 18,
	XlGrid: 15,
	XlHigh: -4127,
	XlInside: 2,
	XlJustify: -4130,
	XlLeft: -4131,
	XlLightDown: 13,
	XlLightHorizontal: 11,
	XlLightUp: 14,
	XlLightVertical: 12,
	XlLow: -4134,
	XlMaximum: 2,
	XlMinimum: 4,
	XlMinusValues: 3,
	XlNextToAxis: 4,
	XlOpaque: 3,
	XlOutside: 3,
	XlPercent: 2,
	XlPlus: 9,
	XlPlusValues: 2,
	XlRight: -4152,
	XlScale: 3,
	XlSemiGray75: 10,
	XlShowLabel: 4,
	XlShowLabelAndPercent: 5,
	XlShowPercent: 3,
	XlShowValue: 2,
	XlSingle: 2,
	XlSolid: 1,
	XlSquare: 1,
	XlStar: 5,
	XlStError: 4,
	XlTop: -4160,
	XlTransparent: 2,
	XlTriangle: 3,
}

// enum XlReadingOrder
var XlReadingOrder = struct {
	XlContext int32
	XlLTR int32
	XlRTL int32
}{
	XlContext: -5002,
	XlLTR: -5003,
	XlRTL: -5004,
}

// enum XlBorderWeight
var XlBorderWeight = struct {
	XlHairline int32
	XlMedium int32
	XlThick int32
	XlThin int32
}{
	XlHairline: 1,
	XlMedium: -4138,
	XlThick: 4,
	XlThin: 2,
}

// enum XlLegendPosition
var XlLegendPosition = struct {
	XlLegendPositionBottom int32
	XlLegendPositionCorner int32
	XlLegendPositionLeft int32
	XlLegendPositionRight int32
	XlLegendPositionTop int32
	XlLegendPositionCustom int32
}{
	XlLegendPositionBottom: -4107,
	XlLegendPositionCorner: 2,
	XlLegendPositionLeft: -4131,
	XlLegendPositionRight: -4152,
	XlLegendPositionTop: -4160,
	XlLegendPositionCustom: -4161,
}

// enum XlUnderlineStyle
var XlUnderlineStyle = struct {
	XlUnderlineStyleDouble int32
	XlUnderlineStyleDoubleAccounting int32
	XlUnderlineStyleNone int32
	XlUnderlineStyleSingle int32
	XlUnderlineStyleSingleAccounting int32
}{
	XlUnderlineStyleDouble: -4119,
	XlUnderlineStyleDoubleAccounting: 5,
	XlUnderlineStyleNone: -4142,
	XlUnderlineStyleSingle: 2,
	XlUnderlineStyleSingleAccounting: 4,
}

// enum XlColorIndex
var XlColorIndex = struct {
	XlColorIndexAutomatic int32
	XlColorIndexNone int32
}{
	XlColorIndexAutomatic: -4105,
	XlColorIndexNone: -4142,
}

// enum XlMarkerStyle
var XlMarkerStyle = struct {
	XlMarkerStyleAutomatic int32
	XlMarkerStyleCircle int32
	XlMarkerStyleDash int32
	XlMarkerStyleDiamond int32
	XlMarkerStyleDot int32
	XlMarkerStyleNone int32
	XlMarkerStylePicture int32
	XlMarkerStylePlus int32
	XlMarkerStyleSquare int32
	XlMarkerStyleStar int32
	XlMarkerStyleTriangle int32
	XlMarkerStyleX int32
}{
	XlMarkerStyleAutomatic: -4105,
	XlMarkerStyleCircle: 8,
	XlMarkerStyleDash: -4115,
	XlMarkerStyleDiamond: 2,
	XlMarkerStyleDot: -4118,
	XlMarkerStyleNone: -4142,
	XlMarkerStylePicture: -4147,
	XlMarkerStylePlus: 9,
	XlMarkerStyleSquare: 1,
	XlMarkerStyleStar: 5,
	XlMarkerStyleTriangle: 3,
	XlMarkerStyleX: -4168,
}

// enum XlRowCol
var XlRowCol = struct {
	XlColumns int32
	XlRows int32
}{
	XlColumns: 2,
	XlRows: 1,
}

// enum XlDataLabelsType
var XlDataLabelsType = struct {
	XlDataLabelsShowNone int32
	XlDataLabelsShowValue int32
	XlDataLabelsShowPercent int32
	XlDataLabelsShowLabel int32
	XlDataLabelsShowLabelAndPercent int32
	XlDataLabelsShowBubbleSizes int32
}{
	XlDataLabelsShowNone: -4142,
	XlDataLabelsShowValue: 2,
	XlDataLabelsShowPercent: 3,
	XlDataLabelsShowLabel: 4,
	XlDataLabelsShowLabelAndPercent: 5,
	XlDataLabelsShowBubbleSizes: 6,
}

// enum XlErrorBarInclude
var XlErrorBarInclude = struct {
	XlErrorBarIncludeBoth int32
	XlErrorBarIncludeMinusValues int32
	XlErrorBarIncludeNone int32
	XlErrorBarIncludePlusValues int32
}{
	XlErrorBarIncludeBoth: 1,
	XlErrorBarIncludeMinusValues: 3,
	XlErrorBarIncludeNone: -4142,
	XlErrorBarIncludePlusValues: 2,
}

// enum XlErrorBarType
var XlErrorBarType = struct {
	XlErrorBarTypeCustom int32
	XlErrorBarTypeFixedValue int32
	XlErrorBarTypePercent int32
	XlErrorBarTypeStDev int32
	XlErrorBarTypeStError int32
}{
	XlErrorBarTypeCustom: -4114,
	XlErrorBarTypeFixedValue: 1,
	XlErrorBarTypePercent: 2,
	XlErrorBarTypeStDev: -4155,
	XlErrorBarTypeStError: 4,
}

// enum XlErrorBarDirection
var XlErrorBarDirection = struct {
	XlChartX int32
	XlChartY int32
}{
	XlChartX: -4168,
	XlChartY: 1,
}

// enum XlChartPictureType
var XlChartPictureType = struct {
	XlStackScale int32
	XlStack int32
	XlStretch int32
}{
	XlStackScale: 3,
	XlStack: 2,
	XlStretch: 1,
}

// enum XlChartItem
var XlChartItem = struct {
	XlDataLabel int32
	XlChartArea int32
	XlSeries int32
	XlChartTitle int32
	XlWalls int32
	XlCorners int32
	XlDataTable int32
	XlTrendline int32
	XlErrorBars int32
	XlXErrorBars int32
	XlYErrorBars int32
	XlLegendEntry int32
	XlLegendKey int32
	XlShape int32
	XlMajorGridlines int32
	XlMinorGridlines int32
	XlAxisTitle int32
	XlUpBars int32
	XlPlotArea int32
	XlDownBars int32
	XlAxis int32
	XlSeriesLines int32
	XlFloor int32
	XlLegend int32
	XlHiLoLines int32
	XlDropLines int32
	XlRadarAxisLabels int32
	XlNothing int32
	XlLeaderLines int32
	XlDisplayUnitLabel int32
	XlPivotChartFieldButton int32
	XlPivotChartDropZone int32
}{
	XlDataLabel: 0,
	XlChartArea: 2,
	XlSeries: 3,
	XlChartTitle: 4,
	XlWalls: 5,
	XlCorners: 6,
	XlDataTable: 7,
	XlTrendline: 8,
	XlErrorBars: 9,
	XlXErrorBars: 10,
	XlYErrorBars: 11,
	XlLegendEntry: 12,
	XlLegendKey: 13,
	XlShape: 14,
	XlMajorGridlines: 15,
	XlMinorGridlines: 16,
	XlAxisTitle: 17,
	XlUpBars: 18,
	XlPlotArea: 19,
	XlDownBars: 20,
	XlAxis: 21,
	XlSeriesLines: 22,
	XlFloor: 23,
	XlLegend: 24,
	XlHiLoLines: 25,
	XlDropLines: 26,
	XlRadarAxisLabels: 27,
	XlNothing: 28,
	XlLeaderLines: 29,
	XlDisplayUnitLabel: 30,
	XlPivotChartFieldButton: 31,
	XlPivotChartDropZone: 32,
}

// enum XlBarShape
var XlBarShape = struct {
	XlBox int32
	XlPyramidToPoint int32
	XlPyramidToMax int32
	XlCylinder int32
	XlConeToPoint int32
	XlConeToMax int32
}{
	XlBox: 0,
	XlPyramidToPoint: 1,
	XlPyramidToMax: 2,
	XlCylinder: 3,
	XlConeToPoint: 4,
	XlConeToMax: 5,
}

// enum XlEndStyleCap
var XlEndStyleCap = struct {
	XlCap int32
	XlNoCap int32
}{
	XlCap: 1,
	XlNoCap: 2,
}

// enum XlTrendlineType
var XlTrendlineType = struct {
	XlExponential int32
	XlLinear int32
	XlLogarithmic int32
	XlMovingAvg int32
	XlPolynomial int32
	XlPower int32
}{
	XlExponential: 5,
	XlLinear: -4132,
	XlLogarithmic: -4133,
	XlMovingAvg: 6,
	XlPolynomial: 3,
	XlPower: 4,
}

// enum XlAxisType
var XlAxisType = struct {
	XlCategory int32
	XlSeriesAxis int32
	XlValue int32
}{
	XlCategory: 1,
	XlSeriesAxis: 3,
	XlValue: 2,
}

// enum XlAxisCrosses
var XlAxisCrosses = struct {
	XlAxisCrossesAutomatic int32
	XlAxisCrossesCustom int32
	XlAxisCrossesMaximum int32
	XlAxisCrossesMinimum int32
}{
	XlAxisCrossesAutomatic: -4105,
	XlAxisCrossesCustom: -4114,
	XlAxisCrossesMaximum: 2,
	XlAxisCrossesMinimum: 4,
}

// enum XlTickMark
var XlTickMark = struct {
	XlTickMarkCross int32
	XlTickMarkInside int32
	XlTickMarkNone int32
	XlTickMarkOutside int32
}{
	XlTickMarkCross: 4,
	XlTickMarkInside: 2,
	XlTickMarkNone: -4142,
	XlTickMarkOutside: 3,
}

// enum XlScaleType
var XlScaleType = struct {
	XlScaleLinear int32
	XlScaleLogarithmic int32
}{
	XlScaleLinear: -4132,
	XlScaleLogarithmic: -4133,
}

// enum XlTickLabelPosition
var XlTickLabelPosition = struct {
	XlTickLabelPositionHigh int32
	XlTickLabelPositionLow int32
	XlTickLabelPositionNextToAxis int32
	XlTickLabelPositionNone int32
}{
	XlTickLabelPositionHigh: -4127,
	XlTickLabelPositionLow: -4134,
	XlTickLabelPositionNextToAxis: 4,
	XlTickLabelPositionNone: -4142,
}

// enum XlTimeUnit
var XlTimeUnit = struct {
	XlDays int32
	XlMonths int32
	XlYears int32
}{
	XlDays: 0,
	XlMonths: 1,
	XlYears: 2,
}

// enum XlCategoryType
var XlCategoryType = struct {
	XlCategoryScale int32
	XlTimeScale int32
	XlAutomaticScale int32
}{
	XlCategoryScale: 2,
	XlTimeScale: 3,
	XlAutomaticScale: -4105,
}

// enum XlDisplayUnit
var XlDisplayUnit = struct {
	XlHundreds int32
	XlThousands int32
	XlTenThousands int32
	XlHundredThousands int32
	XlMillions int32
	XlTenMillions int32
	XlHundredMillions int32
	XlThousandMillions int32
	XlMillionMillions int32
}{
	XlHundreds: -2,
	XlThousands: -3,
	XlTenThousands: -4,
	XlHundredThousands: -5,
	XlMillions: -6,
	XlTenMillions: -7,
	XlHundredMillions: -8,
	XlThousandMillions: -9,
	XlMillionMillions: -10,
}

// enum XlOrientation
var XlOrientation = struct {
	XlDownward int32
	XlHorizontal int32
	XlUpward int32
	XlVertical int32
}{
	XlDownward: -4170,
	XlHorizontal: -4128,
	XlUpward: -4171,
	XlVertical: -4166,
}

// enum XlTickLabelOrientation
var XlTickLabelOrientation = struct {
	XlTickLabelOrientationAutomatic int32
	XlTickLabelOrientationDownward int32
	XlTickLabelOrientationHorizontal int32
	XlTickLabelOrientationUpward int32
	XlTickLabelOrientationVertical int32
}{
	XlTickLabelOrientationAutomatic: -4105,
	XlTickLabelOrientationDownward: -4170,
	XlTickLabelOrientationHorizontal: -4128,
	XlTickLabelOrientationUpward: -4171,
	XlTickLabelOrientationVertical: -4166,
}

// enum XlDisplayBlanksAs
var XlDisplayBlanksAs = struct {
	XlInterpolated int32
	XlNotPlotted int32
	XlZero int32
}{
	XlInterpolated: 3,
	XlNotPlotted: 1,
	XlZero: 2,
}

// enum XlDataLabelPosition
var XlDataLabelPosition = struct {
	XlLabelPositionCenter int32
	XlLabelPositionAbove int32
	XlLabelPositionBelow int32
	XlLabelPositionLeft int32
	XlLabelPositionRight int32
	XlLabelPositionOutsideEnd int32
	XlLabelPositionInsideEnd int32
	XlLabelPositionInsideBase int32
	XlLabelPositionBestFit int32
	XlLabelPositionMixed int32
	XlLabelPositionCustom int32
}{
	XlLabelPositionCenter: -4108,
	XlLabelPositionAbove: 0,
	XlLabelPositionBelow: 1,
	XlLabelPositionLeft: -4131,
	XlLabelPositionRight: -4152,
	XlLabelPositionOutsideEnd: 2,
	XlLabelPositionInsideEnd: 3,
	XlLabelPositionInsideBase: 4,
	XlLabelPositionBestFit: 5,
	XlLabelPositionMixed: 6,
	XlLabelPositionCustom: 7,
}

// enum XlPivotFieldOrientation
var XlPivotFieldOrientation = struct {
	XlColumnField int32
	XlDataField int32
	XlHidden int32
	XlPageField int32
	XlRowField int32
}{
	XlColumnField: 2,
	XlDataField: 4,
	XlHidden: 0,
	XlPageField: 3,
	XlRowField: 1,
}

// enum XlHAlign
var XlHAlign = struct {
	XlHAlignCenter int32
	XlHAlignCenterAcrossSelection int32
	XlHAlignDistributed int32
	XlHAlignFill int32
	XlHAlignGeneral int32
	XlHAlignJustify int32
	XlHAlignLeft int32
	XlHAlignRight int32
}{
	XlHAlignCenter: -4108,
	XlHAlignCenterAcrossSelection: 7,
	XlHAlignDistributed: -4117,
	XlHAlignFill: 5,
	XlHAlignGeneral: 1,
	XlHAlignJustify: -4130,
	XlHAlignLeft: -4131,
	XlHAlignRight: -4152,
}

// enum XlVAlign
var XlVAlign = struct {
	XlVAlignBottom int32
	XlVAlignCenter int32
	XlVAlignDistributed int32
	XlVAlignJustify int32
	XlVAlignTop int32
}{
	XlVAlignBottom: -4107,
	XlVAlignCenter: -4108,
	XlVAlignDistributed: -4117,
	XlVAlignJustify: -4130,
	XlVAlignTop: -4160,
}

// enum XlLineStyle
var XlLineStyle = struct {
	XlContinuous int32
	XlDash int32
	XlDashDot int32
	XlDashDotDot int32
	XlDot int32
	XlDouble int32
	XlSlantDashDot int32
	XlLineStyleNone int32
}{
	XlContinuous: 1,
	XlDash: -4115,
	XlDashDot: 4,
	XlDashDotDot: 5,
	XlDot: -4118,
	XlDouble: -4119,
	XlSlantDashDot: 13,
	XlLineStyleNone: -4142,
}

// enum XlChartElementPosition
var XlChartElementPosition = struct {
	XlChartElementPositionAutomatic int32
	XlChartElementPositionCustom int32
}{
	XlChartElementPositionAutomatic: -4105,
	XlChartElementPositionCustom: -4114,
}

// enum WdUpdateStyleListBehavior
var WdUpdateStyleListBehavior = struct {
	WdListBehaviorKeepPreviousPattern int32
	WdListBehaviorAddBulletsNumbering int32
}{
	WdListBehaviorKeepPreviousPattern: 0,
	WdListBehaviorAddBulletsNumbering: 1,
}

// enum WdApplyQuickStyleSets
var WdApplyQuickStyleSets = struct {
	WdSessionStartSet int32
	WdTemplateSet int32
}{
	WdSessionStartSet: 1,
	WdTemplateSet: 2,
}

// enum WdLigatures
var WdLigatures = struct {
	WdLigaturesNone int32
	WdLigaturesStandard int32
	WdLigaturesContextual int32
	WdLigaturesHistorical int32
	WdLigaturesDiscretional int32
	WdLigaturesStandardContextual int32
	WdLigaturesStandardHistorical int32
	WdLigaturesContextualHistorical int32
	WdLigaturesStandardDiscretional int32
	WdLigaturesContextualDiscretional int32
	WdLigaturesHistoricalDiscretional int32
	WdLigaturesStandardContextualHistorical int32
	WdLigaturesStandardContextualDiscretional int32
	WdLigaturesStandardHistoricalDiscretional int32
	WdLigaturesContextualHistoricalDiscretional int32
	WdLigaturesAll int32
}{
	WdLigaturesNone: 0,
	WdLigaturesStandard: 1,
	WdLigaturesContextual: 2,
	WdLigaturesHistorical: 4,
	WdLigaturesDiscretional: 8,
	WdLigaturesStandardContextual: 3,
	WdLigaturesStandardHistorical: 5,
	WdLigaturesContextualHistorical: 6,
	WdLigaturesStandardDiscretional: 9,
	WdLigaturesContextualDiscretional: 10,
	WdLigaturesHistoricalDiscretional: 12,
	WdLigaturesStandardContextualHistorical: 7,
	WdLigaturesStandardContextualDiscretional: 11,
	WdLigaturesStandardHistoricalDiscretional: 13,
	WdLigaturesContextualHistoricalDiscretional: 14,
	WdLigaturesAll: 15,
}

// enum WdNumberForm
var WdNumberForm = struct {
	WdNumberFormDefault int32
	WdNumberFormLining int32
	WdNumberFormOldStyle int32
}{
	WdNumberFormDefault: 0,
	WdNumberFormLining: 1,
	WdNumberFormOldStyle: 2,
}

// enum WdNumberSpacing
var WdNumberSpacing = struct {
	WdNumberSpacingDefault int32
	WdNumberSpacingProportional int32
	WdNumberSpacingTabular int32
}{
	WdNumberSpacingDefault: 0,
	WdNumberSpacingProportional: 1,
	WdNumberSpacingTabular: 2,
}

// enum WdStylisticSet
var WdStylisticSet = struct {
	WdStylisticSetDefault int32
	WdStylisticSet01 int32
	WdStylisticSet02 int32
	WdStylisticSet03 int32
	WdStylisticSet04 int32
	WdStylisticSet05 int32
	WdStylisticSet06 int32
	WdStylisticSet07 int32
	WdStylisticSet08 int32
	WdStylisticSet09 int32
	WdStylisticSet10 int32
	WdStylisticSet11 int32
	WdStylisticSet12 int32
	WdStylisticSet13 int32
	WdStylisticSet14 int32
	WdStylisticSet15 int32
	WdStylisticSet16 int32
	WdStylisticSet17 int32
	WdStylisticSet18 int32
	WdStylisticSet19 int32
	WdStylisticSet20 int32
}{
	WdStylisticSetDefault: 0,
	WdStylisticSet01: 1,
	WdStylisticSet02: 2,
	WdStylisticSet03: 4,
	WdStylisticSet04: 8,
	WdStylisticSet05: 16,
	WdStylisticSet06: 32,
	WdStylisticSet07: 64,
	WdStylisticSet08: 128,
	WdStylisticSet09: 256,
	WdStylisticSet10: 512,
	WdStylisticSet11: 1024,
	WdStylisticSet12: 2048,
	WdStylisticSet13: 4096,
	WdStylisticSet14: 8192,
	WdStylisticSet15: 16384,
	WdStylisticSet16: 32768,
	WdStylisticSet17: 65536,
	WdStylisticSet18: 131072,
	WdStylisticSet19: 262144,
	WdStylisticSet20: 524288,
}

// enum WdSpanishSpeller
var WdSpanishSpeller = struct {
	WdSpanishTuteoOnly int32
	WdSpanishTuteoAndVoseo int32
	WdSpanishVoseoOnly int32
}{
	WdSpanishTuteoOnly: 0,
	WdSpanishTuteoAndVoseo: 1,
	WdSpanishVoseoOnly: 2,
}

// enum WdLockType
var WdLockType = struct {
	WdLockNone int32
	WdLockReservation int32
	WdLockEphemeral int32
	WdLockChanged int32
}{
	WdLockNone: 0,
	WdLockReservation: 1,
	WdLockEphemeral: 2,
	WdLockChanged: 3,
}

// enum XlPieSliceLocation
var XlPieSliceLocation = struct {
	XlHorizontalCoordinate int32
	XlVerticalCoordinate int32
}{
	XlHorizontalCoordinate: 1,
	XlVerticalCoordinate: 2,
}

// enum XlPieSliceIndex
var XlPieSliceIndex = struct {
	XlOuterCounterClockwisePoint int32
	XlOuterCenterPoint int32
	XlOuterClockwisePoint int32
	XlMidClockwiseRadiusPoint int32
	XlCenterPoint int32
	XlMidCounterClockwiseRadiusPoint int32
	XlInnerClockwisePoint int32
	XlInnerCenterPoint int32
	XlInnerCounterClockwisePoint int32
}{
	XlOuterCounterClockwisePoint: 1,
	XlOuterCenterPoint: 2,
	XlOuterClockwisePoint: 3,
	XlMidClockwiseRadiusPoint: 4,
	XlCenterPoint: 5,
	XlMidCounterClockwiseRadiusPoint: 6,
	XlInnerClockwisePoint: 7,
	XlInnerCenterPoint: 8,
	XlInnerCounterClockwisePoint: 9,
}

// enum WdCompatibilityMode
var WdCompatibilityMode = struct {
	WdWord2003 int32
	WdWord2007 int32
	WdWord2010 int32
	WdCurrent int32
}{
	WdWord2003: 11,
	WdWord2007: 12,
	WdWord2010: 14,
	WdCurrent: 65535,
}

// enum WdProtectedViewCloseReason
var WdProtectedViewCloseReason = struct {
	WdProtectedViewCloseNormal int32
	WdProtectedViewCloseEdit int32
	WdProtectedViewCloseForced int32
}{
	WdProtectedViewCloseNormal: 0,
	WdProtectedViewCloseEdit: 1,
	WdProtectedViewCloseForced: 2,
}

// enum WdPortugueseReform
var WdPortugueseReform = struct {
	WdPortuguesePreReform int32
	WdPortuguesePostReform int32
	WdPortugueseBoth int32
}{
	WdPortuguesePreReform: 1,
	WdPortuguesePostReform: 2,
	WdPortugueseBoth: 3,
}

