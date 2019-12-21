unit Word_TLB_Constants;
(*//*************************************************************************
  This file is created as a subset of Word_TLB.pas.
  We do not need the Types, Classes and Interfaces, so we have copied out the
  Constants part of the file. It can be updated at any time.  The file needs
  to have all declarations of TOLEEnum removed from the constants section.

  All Values are HEX no Decimal.

  I guess this is probably in some way Copyright Microsoft.

//*************************************************************************)
interface


// *********************************************************************//
// Declaration of Enumerations defined in Type Library
// *********************************************************************//
// Constants for enum WdMailSystem
const
  wdNoMailSystem = $00000000;
  wdMAPI = $00000001;
  wdPowerTalk = $00000002;
  wdMAPIandPowerTalk = $00000003;


// Constants for enum WdTemplateType
const
  wdNormalTemplate = $00000000;
  wdGlobalTemplate = $00000001;
  wdAttachedTemplate = $00000002;


// Constants for enum WdContinue
const
  wdContinueDisabled = $00000000;
  wdResetList = $00000001;
  wdContinueList = $00000002;


// Constants for enum WdIMEMode
const
  wdIMEModeNoControl = $00000000;
  wdIMEModeOn = $00000001;
  wdIMEModeOff = $00000002;
  wdIMEModeHiragana = $00000004;
  wdIMEModeKatakana = $00000005;
  wdIMEModeKatakanaHalf = $00000006;
  wdIMEModeAlphaFull = $00000007;
  wdIMEModeAlpha = $00000008;
  wdIMEModeHangulFull = $00000009;
  wdIMEModeHangul = $0000000A;


// Constants for enum WdBaselineAlignment
const
  wdBaselineAlignTop = $00000000;
  wdBaselineAlignCenter = $00000001;
  wdBaselineAlignBaseline = $00000002;
  wdBaselineAlignFarEast50 = $00000003;
  wdBaselineAlignAuto = $00000004;


// Constants for enum WdIndexFilter
const
  wdIndexFilterNone = $00000000;
  wdIndexFilterAiueo = $00000001;
  wdIndexFilterAkasatana = $00000002;
  wdIndexFilterChosung = $00000003;
  wdIndexFilterLow = $00000004;
  wdIndexFilterMedium = $00000005;
  wdIndexFilterFull = $00000006;


// Constants for enum WdIndexSortBy
const
  wdIndexSortByStroke = $00000000;
  wdIndexSortBySyllable = $00000001;


// Constants for enum WdJustificationMode
const
  wdJustificationModeExpand = $00000000;
  wdJustificationModeCompress = $00000001;
  wdJustificationModeCompressKana = $00000002;


// Constants for enum WdFarEastLineBreakLevel
const
  wdFarEastLineBreakLevelNormal = $00000000;
  wdFarEastLineBreakLevelStrict = $00000001;
  wdFarEastLineBreakLevelCustom = $00000002;


// Constants for enum WdMultipleWordConversionsMode
const
  wdHangulToHanja = $00000000;
  wdHanjaToHangul = $00000001;


// Constants for enum WdColorIndex
const
  wdAuto = $00000000;
  wdBlack = $00000001;
  wdBlue = $00000002;
  wdTurquoise = $00000003;
  wdBrightGreen = $00000004;
  wdPink = $00000005;
  wdRed = $00000006;
  wdYellow = $00000007;
  wdWhite = $00000008;
  wdDarkBlue = $00000009;
  wdTeal = $0000000A;
  wdGreen = $0000000B;
  wdViolet = $0000000C;
  wdDarkRed = $0000000D;
  wdDarkYellow = $0000000E;
  wdGray50 = $0000000F;
  wdGray25 = $00000010;
  wdClassicRed = $00000011;
  wdClassicBlue = $00000012;
  wdByAuthor = $FFFFFFFF;
  wdNoHighlight = $00000000;


// Constants for enum WdTextureIndex
const
  wdTextureNone = $00000000;
  wdTexture2Pt5Percent = $00000019;
  wdTexture5Percent = $00000032;
  wdTexture7Pt5Percent = $0000004B;
  wdTexture10Percent = $00000064;
  wdTexture12Pt5Percent = $0000007D;
  wdTexture15Percent = $00000096;
  wdTexture17Pt5Percent = $000000AF;
  wdTexture20Percent = $000000C8;
  wdTexture22Pt5Percent = $000000E1;
  wdTexture25Percent = $000000FA;
  wdTexture27Pt5Percent = $00000113;
  wdTexture30Percent = $0000012C;
  wdTexture32Pt5Percent = $00000145;
  wdTexture35Percent = $0000015E;
  wdTexture37Pt5Percent = $00000177;
  wdTexture40Percent = $00000190;
  wdTexture42Pt5Percent = $000001A9;
  wdTexture45Percent = $000001C2;
  wdTexture47Pt5Percent = $000001DB;
  wdTexture50Percent = $000001F4;
  wdTexture52Pt5Percent = $0000020D;
  wdTexture55Percent = $00000226;
  wdTexture57Pt5Percent = $0000023F;
  wdTexture60Percent = $00000258;
  wdTexture62Pt5Percent = $00000271;
  wdTexture65Percent = $0000028A;
  wdTexture67Pt5Percent = $000002A3;
  wdTexture70Percent = $000002BC;
  wdTexture72Pt5Percent = $000002D5;
  wdTexture75Percent = $000002EE;
  wdTexture77Pt5Percent = $00000307;
  wdTexture80Percent = $00000320;
  wdTexture82Pt5Percent = $00000339;
  wdTexture85Percent = $00000352;
  wdTexture87Pt5Percent = $0000036B;
  wdTexture90Percent = $00000384;
  wdTexture92Pt5Percent = $0000039D;
  wdTexture95Percent = $000003B6;
  wdTexture97Pt5Percent = $000003CF;
  wdTextureSolid = $000003E8;
  wdTextureDarkHorizontal = $FFFFFFFF;
  wdTextureDarkVertical = $FFFFFFFE;
  wdTextureDarkDiagonalDown = $FFFFFFFD;
  wdTextureDarkDiagonalUp = $FFFFFFFC;
  wdTextureDarkCross = $FFFFFFFB;
  wdTextureDarkDiagonalCross = $FFFFFFFA;
  wdTextureHorizontal = $FFFFFFF9;
  wdTextureVertical = $FFFFFFF8;
  wdTextureDiagonalDown = $FFFFFFF7;
  wdTextureDiagonalUp = $FFFFFFF6;
  wdTextureCross = $FFFFFFF5;
  wdTextureDiagonalCross = $FFFFFFF4;


// Constants for enum WdUnderline
const
  wdUnderlineNone = $00000000;
  wdUnderlineSingle = $00000001;
  wdUnderlineWords = $00000002;
  wdUnderlineDouble = $00000003;
  wdUnderlineDotted = $00000004;
  wdUnderlineThick = $00000006;
  wdUnderlineDash = $00000007;
  wdUnderlineDotDash = $00000009;
  wdUnderlineDotDotDash = $0000000A;
  wdUnderlineWavy = $0000000B;
  wdUnderlineWavyHeavy = $0000001B;
  wdUnderlineDottedHeavy = $00000014;
  wdUnderlineDashHeavy = $00000017;
  wdUnderlineDotDashHeavy = $00000019;
  wdUnderlineDotDotDashHeavy = $0000001A;
  wdUnderlineDashLong = $00000027;
  wdUnderlineDashLongHeavy = $00000037;
  wdUnderlineWavyDouble = $0000002B;


// Constants for enum WdEmphasisMark
const
  wdEmphasisMarkNone = $00000000;
  wdEmphasisMarkOverSolidCircle = $00000001;
  wdEmphasisMarkOverComma = $00000002;
  wdEmphasisMarkOverWhiteCircle = $00000003;
  wdEmphasisMarkUnderSolidCircle = $00000004;


// Constants for enum WdInternationalIndex
const
  wdListSeparator = $00000011;
  wdDecimalSeparator = $00000012;
  wdThousandsSeparator = $00000013;
  wdCurrencyCode = $00000014;
  wd24HourClock = $00000015;
  wdInternationalAM = $00000016;
  wdInternationalPM = $00000017;
  wdTimeSeparator = $00000018;
  wdDateSeparator = $00000019;
  wdProductLanguageID = $0000001A;


// Constants for enum WdAutoMacros
const
  wdAutoExec = $00000000;
  wdAutoNew = $00000001;
  wdAutoOpen = $00000002;
  wdAutoClose = $00000003;
  wdAutoExit = $00000004;
  wdAutoSync = $00000005;


// Constants for enum WdCaptionPosition
const
  wdCaptionPositionAbove = $00000000;
  wdCaptionPositionBelow = $00000001;


// Constants for enum WdCountry
const
  wdUS = $00000001;
  wdCanada = $00000002;
  wdLatinAmerica = $00000003;
  wdNetherlands = $0000001F;
  wdFrance = $00000021;
  wdSpain = $00000022;
  wdItaly = $00000027;
  wdUK = $0000002C;
  wdDenmark = $0000002D;
  wdSweden = $0000002E;
  wdNorway = $0000002F;
  wdGermany = $00000031;
  wdPeru = $00000033;
  wdMexico = $00000034;
  wdArgentina = $00000036;
  wdBrazil = $00000037;
  wdChile = $00000038;
  wdVenezuela = $0000003A;
  wdJapan = $00000051;
  wdTaiwan = $00000376;
  wdChina = $00000056;
  wdKorea = $00000052;
  wdFinland = $00000166;
  wdIceland = $00000162;


// Constants for enum WdHeadingSeparator
const
  wdHeadingSeparatorNone = $00000000;
  wdHeadingSeparatorBlankLine = $00000001;
  wdHeadingSeparatorLetter = $00000002;
  wdHeadingSeparatorLetterLow = $00000003;
  wdHeadingSeparatorLetterFull = $00000004;


// Constants for enum WdSeparatorType
const
  wdSeparatorHyphen = $00000000;
  wdSeparatorPeriod = $00000001;
  wdSeparatorColon = $00000002;
  wdSeparatorEmDash = $00000003;
  wdSeparatorEnDash = $00000004;


// Constants for enum WdPageNumberAlignment
const
  wdAlignPageNumberLeft = $00000000;
  wdAlignPageNumberCenter = $00000001;
  wdAlignPageNumberRight = $00000002;
  wdAlignPageNumberInside = $00000003;
  wdAlignPageNumberOutside = $00000004;


// Constants for enum WdBorderType
const
  wdBorderTop = $FFFFFFFF;
  wdBorderLeft = $FFFFFFFE;
  wdBorderBottom = $FFFFFFFD;
  wdBorderRight = $FFFFFFFC;
  wdBorderHorizontal = $FFFFFFFB;
  wdBorderVertical = $FFFFFFFA;
  wdBorderDiagonalDown = $FFFFFFF9;
  wdBorderDiagonalUp = $FFFFFFF8;


// Constants for enum WdBorderTypeHID
const
  emptyenum = $00000000;


// Constants for enum WdFramePosition
const
  wdFrameTop = $FFF0BDC1;
  wdFrameLeft = $FFF0BDC2;
  wdFrameBottom = $FFF0BDC3;
  wdFrameRight = $FFF0BDC4;
  wdFrameCenter = $FFF0BDC5;
  wdFrameInside = $FFF0BDC6;
  wdFrameOutside = $FFF0BDC7;


// Constants for enum WdAnimation
const
  wdAnimationNone = $00000000;
  wdAnimationLasVegasLights = $00000001;
  wdAnimationBlinkingBackground = $00000002;
  wdAnimationSparkleText = $00000003;
  wdAnimationMarchingBlackAnts = $00000004;
  wdAnimationMarchingRedAnts = $00000005;
  wdAnimationShimmer = $00000006;


// Constants for enum WdCharacterCase
const
  wdNextCase = $FFFFFFFF;
  wdLowerCase = $00000000;
  wdUpperCase = $00000001;
  wdTitleWord = $00000002;
  wdTitleSentence = $00000004;
  wdToggleCase = $00000005;
  wdHalfWidth = $00000006;
  wdFullWidth = $00000007;
  wdKatakana = $00000008;
  wdHiragana = $00000009;


// Constants for enum WdCharacterCaseHID
const
  emptyenum_ = $00000000;


// Constants for enum WdSummaryMode
const
  wdSummaryModeHighlight = $00000000;
  wdSummaryModeHideAllButSummary = $00000001;
  wdSummaryModeInsert = $00000002;
  wdSummaryModeCreateNew = $00000003;


// Constants for enum WdSummaryLength
const
  wd10Sentences = $FFFFFFFE;
  wd20Sentences = $FFFFFFFD;
  wd100Words = $FFFFFFFC;
  wd500Words = $FFFFFFFB;
  wd10Percent = $FFFFFFFA;
  wd25Percent = $FFFFFFF9;
  wd50Percent = $FFFFFFF8;
  wd75Percent = $FFFFFFF7;


// Constants for enum WdStyleType
const
  wdStyleTypeParagraph = $00000001;
  wdStyleTypeCharacter = $00000002;
  wdStyleTypeTable = $00000003;
  wdStyleTypeList = $00000004;
  wdStyleTypeParagraphOnly = $00000005;
  wdStyleTypeLinked = $00000006;


// Constants for enum WdUnits
const
  wdCharacter = $00000001;
  wdWord = $00000002;
  wdSentence = $00000003;
  wdParagraph = $00000004;
  wdLine = $00000005;
  wdStory = $00000006;
  wdScreen = $00000007;
  wdSection = $00000008;
  wdColumn = $00000009;
  wdRow = $0000000A;
  wdWindow = $0000000B;
  wdCell = $0000000C;
  wdCharacterFormatting = $0000000D;
  wdParagraphFormatting = $0000000E;
  wdTable = $0000000F;
  wdItem = $00000010;


// Constants for enum WdGoToItem
const
  wdGoToBookmark = $FFFFFFFF;
  wdGoToSection = $00000000;
  wdGoToPage = $00000001;
  wdGoToTable = $00000002;
  wdGoToLine = $00000003;
  wdGoToFootnote = $00000004;
  wdGoToEndnote = $00000005;
  wdGoToComment = $00000006;
  wdGoToField = $00000007;
  wdGoToGraphic = $00000008;
  wdGoToObject = $00000009;
  wdGoToEquation = $0000000A;
  wdGoToHeading = $0000000B;
  wdGoToPercent = $0000000C;
  wdGoToSpellingError = $0000000D;
  wdGoToGrammaticalError = $0000000E;
  wdGoToProofreadingError = $0000000F;


// Constants for enum WdGoToDirection
const
  wdGoToFirst = $00000001;
  wdGoToLast = $FFFFFFFF;
  wdGoToNext = $00000002;
  wdGoToRelative = $00000002;
  wdGoToPrevious = $00000003;
  wdGoToAbsolute = $00000001;


// Constants for enum WdCollapseDirection
const
  wdCollapseStart = $00000001;
  wdCollapseEnd = $00000000;


// Constants for enum WdRowHeightRule
const
  wdRowHeightAuto = $00000000;
  wdRowHeightAtLeast = $00000001;
  wdRowHeightExactly = $00000002;


// Constants for enum WdFrameSizeRule
const
  wdFrameAuto = $00000000;
  wdFrameAtLeast = $00000001;
  wdFrameExact = $00000002;


// Constants for enum WdInsertCells
const
  wdInsertCellsShiftRight = $00000000;
  wdInsertCellsShiftDown = $00000001;
  wdInsertCellsEntireRow = $00000002;
  wdInsertCellsEntireColumn = $00000003;


// Constants for enum WdDeleteCells
const
  wdDeleteCellsShiftLeft = $00000000;
  wdDeleteCellsShiftUp = $00000001;
  wdDeleteCellsEntireRow = $00000002;
  wdDeleteCellsEntireColumn = $00000003;


// Constants for enum WdListApplyTo
const
  wdListApplyToWholeList = $00000000;
  wdListApplyToThisPointForward = $00000001;
  wdListApplyToSelection = $00000002;


// Constants for enum WdAlertLevel
const
  wdAlertsNone = $00000000;
  wdAlertsMessageBox = $FFFFFFFE;
  wdAlertsAll = $FFFFFFFF;


// Constants for enum WdCursorType
const
  wdCursorWait = $00000000;
  wdCursorIBeam = $00000001;
  wdCursorNormal = $00000002;
  wdCursorNorthwestArrow = $00000003;


// Constants for enum WdEnableCancelKey
const
  wdCancelDisabled = $00000000;
  wdCancelInterrupt = $00000001;


// Constants for enum WdRulerStyle
const
  wdAdjustNone = $00000000;
  wdAdjustProportional = $00000001;
  wdAdjustFirstColumn = $00000002;
  wdAdjustSameWidth = $00000003;


// Constants for enum WdParagraphAlignment
const
  wdAlignParagraphLeft = $00000000;
  wdAlignParagraphCenter = $00000001;
  wdAlignParagraphRight = $00000002;
  wdAlignParagraphJustify = $00000003;
  wdAlignParagraphDistribute = $00000004;
  wdAlignParagraphJustifyMed = $00000005;
  wdAlignParagraphJustifyHi = $00000007;
  wdAlignParagraphJustifyLow = $00000008;
  wdAlignParagraphThaiJustify = $00000009;


// Constants for enum WdParagraphAlignmentHID
const
  emptyenum__ = $00000000;


// Constants for enum WdListLevelAlignment
const
  wdListLevelAlignLeft = $00000000;
  wdListLevelAlignCenter = $00000001;
  wdListLevelAlignRight = $00000002;


// Constants for enum WdRowAlignment
const
  wdAlignRowLeft = $00000000;
  wdAlignRowCenter = $00000001;
  wdAlignRowRight = $00000002;


// Constants for enum WdTabAlignment
const
  wdAlignTabLeft = $00000000;
  wdAlignTabCenter = $00000001;
  wdAlignTabRight = $00000002;
  wdAlignTabDecimal = $00000003;
  wdAlignTabBar = $00000004;
  wdAlignTabList = $00000006;


// Constants for enum WdVerticalAlignment
const
  wdAlignVerticalTop = $00000000;
  wdAlignVerticalCenter = $00000001;
  wdAlignVerticalJustify = $00000002;
  wdAlignVerticalBottom = $00000003;


// Constants for enum WdCellVerticalAlignment
const
  wdCellAlignVerticalTop = $00000000;
  wdCellAlignVerticalCenter = $00000001;
  wdCellAlignVerticalBottom = $00000003;


// Constants for enum WdTrailingCharacter
const
  wdTrailingTab = $00000000;
  wdTrailingSpace = $00000001;
  wdTrailingNone = $00000002;


// Constants for enum WdListGalleryType
const
  wdBulletGallery = $00000001;
  wdNumberGallery = $00000002;
  wdOutlineNumberGallery = $00000003;


// Constants for enum WdListNumberStyle
const
  wdListNumberStyleArabic = $00000000;
  wdListNumberStyleUppercaseRoman = $00000001;
  wdListNumberStyleLowercaseRoman = $00000002;
  wdListNumberStyleUppercaseLetter = $00000003;
  wdListNumberStyleLowercaseLetter = $00000004;
  wdListNumberStyleOrdinal = $00000005;
  wdListNumberStyleCardinalText = $00000006;
  wdListNumberStyleOrdinalText = $00000007;
  wdListNumberStyleKanji = $0000000A;
  wdListNumberStyleKanjiDigit = $0000000B;
  wdListNumberStyleAiueoHalfWidth = $0000000C;
  wdListNumberStyleIrohaHalfWidth = $0000000D;
  wdListNumberStyleArabicFullWidth = $0000000E;
  wdListNumberStyleKanjiTraditional = $00000010;
  wdListNumberStyleKanjiTraditional2 = $00000011;
  wdListNumberStyleNumberInCircle = $00000012;
  wdListNumberStyleAiueo = $00000014;
  wdListNumberStyleIroha = $00000015;
  wdListNumberStyleArabicLZ = $00000016;
  wdListNumberStyleBullet = $00000017;
  wdListNumberStyleGanada = $00000018;
  wdListNumberStyleChosung = $00000019;
  wdListNumberStyleGBNum1 = $0000001A;
  wdListNumberStyleGBNum2 = $0000001B;
  wdListNumberStyleGBNum3 = $0000001C;
  wdListNumberStyleGBNum4 = $0000001D;
  wdListNumberStyleZodiac1 = $0000001E;
  wdListNumberStyleZodiac2 = $0000001F;
  wdListNumberStyleZodiac3 = $00000020;
  wdListNumberStyleTradChinNum1 = $00000021;
  wdListNumberStyleTradChinNum2 = $00000022;
  wdListNumberStyleTradChinNum3 = $00000023;
  wdListNumberStyleTradChinNum4 = $00000024;
  wdListNumberStyleSimpChinNum1 = $00000025;
  wdListNumberStyleSimpChinNum2 = $00000026;
  wdListNumberStyleSimpChinNum3 = $00000027;
  wdListNumberStyleSimpChinNum4 = $00000028;
  wdListNumberStyleHanjaRead = $00000029;
  wdListNumberStyleHanjaReadDigit = $0000002A;
  wdListNumberStyleHangul = $0000002B;
  wdListNumberStyleHanja = $0000002C;
  wdListNumberStyleHebrew1 = $0000002D;
  wdListNumberStyleArabic1 = $0000002E;
  wdListNumberStyleHebrew2 = $0000002F;
  wdListNumberStyleArabic2 = $00000030;
  wdListNumberStyleHindiLetter1 = $00000031;
  wdListNumberStyleHindiLetter2 = $00000032;
  wdListNumberStyleHindiArabic = $00000033;
  wdListNumberStyleHindiCardinalText = $00000034;
  wdListNumberStyleThaiLetter = $00000035;
  wdListNumberStyleThaiArabic = $00000036;
  wdListNumberStyleThaiCardinalText = $00000037;
  wdListNumberStyleVietCardinalText = $00000038;
  wdListNumberStyleLowercaseRussian = $0000003A;
  wdListNumberStyleUppercaseRussian = $0000003B;
  wdListNumberStyleLowercaseGreek = $0000003C;
  wdListNumberStyleUppercaseGreek = $0000003D;
  wdListNumberStyleArabicLZ2 = $0000003E;
  wdListNumberStyleArabicLZ3 = $0000003F;
  wdListNumberStyleArabicLZ4 = $00000040;
  wdListNumberStyleLowercaseTurkish = $00000041;
  wdListNumberStyleUppercaseTurkish = $00000042;
  wdListNumberStyleLowercaseBulgarian = $00000043;
  wdListNumberStyleUppercaseBulgarian = $00000044;
  wdListNumberStylePictureBullet = $000000F9;
  wdListNumberStyleLegal = $000000FD;
  wdListNumberStyleLegalLZ = $000000FE;
  wdListNumberStyleNone = $000000FF;


// Constants for enum WdListNumberStyleHID
const
  emptyenum___ = $00000000;


// Constants for enum WdNoteNumberStyle
const
  wdNoteNumberStyleArabic = $00000000;
  wdNoteNumberStyleUppercaseRoman = $00000001;
  wdNoteNumberStyleLowercaseRoman = $00000002;
  wdNoteNumberStyleUppercaseLetter = $00000003;
  wdNoteNumberStyleLowercaseLetter = $00000004;
  wdNoteNumberStyleSymbol = $00000009;
  wdNoteNumberStyleArabicFullWidth = $0000000E;
  wdNoteNumberStyleKanji = $0000000A;
  wdNoteNumberStyleKanjiDigit = $0000000B;
  wdNoteNumberStyleKanjiTraditional = $00000010;
  wdNoteNumberStyleNumberInCircle = $00000012;
  wdNoteNumberStyleHanjaRead = $00000029;
  wdNoteNumberStyleHanjaReadDigit = $0000002A;
  wdNoteNumberStyleTradChinNum1 = $00000021;
  wdNoteNumberStyleTradChinNum2 = $00000022;
  wdNoteNumberStyleSimpChinNum1 = $00000025;
  wdNoteNumberStyleSimpChinNum2 = $00000026;
  wdNoteNumberStyleHebrewLetter1 = $0000002D;
  wdNoteNumberStyleArabicLetter1 = $0000002E;
  wdNoteNumberStyleHebrewLetter2 = $0000002F;
  wdNoteNumberStyleArabicLetter2 = $00000030;
  wdNoteNumberStyleHindiLetter1 = $00000031;
  wdNoteNumberStyleHindiLetter2 = $00000032;
  wdNoteNumberStyleHindiArabic = $00000033;
  wdNoteNumberStyleHindiCardinalText = $00000034;
  wdNoteNumberStyleThaiLetter = $00000035;
  wdNoteNumberStyleThaiArabic = $00000036;
  wdNoteNumberStyleThaiCardinalText = $00000037;
  wdNoteNumberStyleVietCardinalText = $00000038;


// Constants for enum WdNoteNumberStyleHID
const
  emptyenum____ = $00000000;


// Constants for enum WdCaptionNumberStyle
const
  wdCaptionNumberStyleArabic = $00000000;
  wdCaptionNumberStyleUppercaseRoman = $00000001;
  wdCaptionNumberStyleLowercaseRoman = $00000002;
  wdCaptionNumberStyleUppercaseLetter = $00000003;
  wdCaptionNumberStyleLowercaseLetter = $00000004;
  wdCaptionNumberStyleArabicFullWidth = $0000000E;
  wdCaptionNumberStyleKanji = $0000000A;
  wdCaptionNumberStyleKanjiDigit = $0000000B;
  wdCaptionNumberStyleKanjiTraditional = $00000010;
  wdCaptionNumberStyleNumberInCircle = $00000012;
  wdCaptionNumberStyleGanada = $00000018;
  wdCaptionNumberStyleChosung = $00000019;
  wdCaptionNumberStyleZodiac1 = $0000001E;
  wdCaptionNumberStyleZodiac2 = $0000001F;
  wdCaptionNumberStyleHanjaRead = $00000029;
  wdCaptionNumberStyleHanjaReadDigit = $0000002A;
  wdCaptionNumberStyleTradChinNum2 = $00000022;
  wdCaptionNumberStyleTradChinNum3 = $00000023;
  wdCaptionNumberStyleSimpChinNum2 = $00000026;
  wdCaptionNumberStyleSimpChinNum3 = $00000027;
  wdCaptionNumberStyleHebrewLetter1 = $0000002D;
  wdCaptionNumberStyleArabicLetter1 = $0000002E;
  wdCaptionNumberStyleHebrewLetter2 = $0000002F;
  wdCaptionNumberStyleArabicLetter2 = $00000030;
  wdCaptionNumberStyleHindiLetter1 = $00000031;
  wdCaptionNumberStyleHindiLetter2 = $00000032;
  wdCaptionNumberStyleHindiArabic = $00000033;
  wdCaptionNumberStyleHindiCardinalText = $00000034;
  wdCaptionNumberStyleThaiLetter = $00000035;
  wdCaptionNumberStyleThaiArabic = $00000036;
  wdCaptionNumberStyleThaiCardinalText = $00000037;
  wdCaptionNumberStyleVietCardinalText = $00000038;


// Constants for enum WdCaptionNumberStyleHID
const
  emptyenum_____ = $00000000;


// Constants for enum WdPageNumberStyle
const
  wdPageNumberStyleArabic = $00000000;
  wdPageNumberStyleUppercaseRoman = $00000001;
  wdPageNumberStyleLowercaseRoman = $00000002;
  wdPageNumberStyleUppercaseLetter = $00000003;
  wdPageNumberStyleLowercaseLetter = $00000004;
  wdPageNumberStyleArabicFullWidth = $0000000E;
  wdPageNumberStyleKanji = $0000000A;
  wdPageNumberStyleKanjiDigit = $0000000B;
  wdPageNumberStyleKanjiTraditional = $00000010;
  wdPageNumberStyleNumberInCircle = $00000012;
  wdPageNumberStyleHanjaRead = $00000029;
  wdPageNumberStyleHanjaReadDigit = $0000002A;
  wdPageNumberStyleTradChinNum1 = $00000021;
  wdPageNumberStyleTradChinNum2 = $00000022;
  wdPageNumberStyleSimpChinNum1 = $00000025;
  wdPageNumberStyleSimpChinNum2 = $00000026;
  wdPageNumberStyleHebrewLetter1 = $0000002D;
  wdPageNumberStyleArabicLetter1 = $0000002E;
  wdPageNumberStyleHebrewLetter2 = $0000002F;
  wdPageNumberStyleArabicLetter2 = $00000030;
  wdPageNumberStyleHindiLetter1 = $00000031;
  wdPageNumberStyleHindiLetter2 = $00000032;
  wdPageNumberStyleHindiArabic = $00000033;
  wdPageNumberStyleHindiCardinalText = $00000034;
  wdPageNumberStyleThaiLetter = $00000035;
  wdPageNumberStyleThaiArabic = $00000036;
  wdPageNumberStyleThaiCardinalText = $00000037;
  wdPageNumberStyleVietCardinalText = $00000038;
  wdPageNumberStyleNumberInDash = $00000039;


// Constants for enum WdPageNumberStyleHID
const
  emptyenum______ = $00000000;


// Constants for enum WdStatistic
const
  wdStatisticWords = $00000000;
  wdStatisticLines = $00000001;
  wdStatisticPages = $00000002;
  wdStatisticCharacters = $00000003;
  wdStatisticParagraphs = $00000004;
  wdStatisticCharactersWithSpaces = $00000005;
  wdStatisticFarEastCharacters = $00000006;


// Constants for enum WdStatisticHID
const
  emptyenum_______ = $00000000;


// Constants for enum WdBuiltInProperty
const
  wdPropertyTitle = $00000001;
  wdPropertySubject = $00000002;
  wdPropertyAuthor = $00000003;
  wdPropertyKeywords = $00000004;
  wdPropertyComments = $00000005;
  wdPropertyTemplate = $00000006;
  wdPropertyLastAuthor = $00000007;
  wdPropertyRevision = $00000008;
  wdPropertyAppName = $00000009;
  wdPropertyTimeLastPrinted = $0000000A;
  wdPropertyTimeCreated = $0000000B;
  wdPropertyTimeLastSaved = $0000000C;
  wdPropertyVBATotalEdit = $0000000D;
  wdPropertyPages = $0000000E;
  wdPropertyWords = $0000000F;
  wdPropertyCharacters = $00000010;
  wdPropertySecurity = $00000011;
  wdPropertyCategory = $00000012;
  wdPropertyFormat = $00000013;
  wdPropertyManager = $00000014;
  wdPropertyCompany = $00000015;
  wdPropertyBytes = $00000016;
  wdPropertyLines = $00000017;
  wdPropertyParas = $00000018;
  wdPropertySlides = $00000019;
  wdPropertyNotes = $0000001A;
  wdPropertyHiddenSlides = $0000001B;
  wdPropertyMMClips = $0000001C;
  wdPropertyHyperlinkBase = $0000001D;
  wdPropertyCharsWSpaces = $0000001E;


// Constants for enum WdLineSpacing
const
  wdLineSpaceSingle = $00000000;
  wdLineSpace1pt5 = $00000001;
  wdLineSpaceDouble = $00000002;
  wdLineSpaceAtLeast = $00000003;
  wdLineSpaceExactly = $00000004;
  wdLineSpaceMultiple = $00000005;


// Constants for enum WdNumberType
const
  wdNumberParagraph = $00000001;
  wdNumberListNum = $00000002;
  wdNumberAllNumbers = $00000003;


// Constants for enum WdListType
const
  wdListNoNumbering = $00000000;
  wdListListNumOnly = $00000001;
  wdListBullet = $00000002;
  wdListSimpleNumbering = $00000003;
  wdListOutlineNumbering = $00000004;
  wdListMixedNumbering = $00000005;
  wdListPictureBullet = $00000006;


// Constants for enum WdStoryType
const
  wdMainTextStory = $00000001;
  wdFootnotesStory = $00000002;
  wdEndnotesStory = $00000003;
  wdCommentsStory = $00000004;
  wdTextFrameStory = $00000005;
  wdEvenPagesHeaderStory = $00000006;
  wdPrimaryHeaderStory = $00000007;
  wdEvenPagesFooterStory = $00000008;
  wdPrimaryFooterStory = $00000009;
  wdFirstPageHeaderStory = $0000000A;
  wdFirstPageFooterStory = $0000000B;
  wdFootnoteSeparatorStory = $0000000C;
  wdFootnoteContinuationSeparatorStory = $0000000D;
  wdFootnoteContinuationNoticeStory = $0000000E;
  wdEndnoteSeparatorStory = $0000000F;
  wdEndnoteContinuationSeparatorStory = $00000010;
  wdEndnoteContinuationNoticeStory = $00000011;


// Constants for enum WdSaveFormat
const
  wdFormatDocument = $00000000;
  wdFormatDocument97 = $00000000;
  wdFormatTemplate = $00000001;
  wdFormatTemplate97 = $00000001;
  wdFormatText = $00000002;
  wdFormatTextLineBreaks = $00000003;
  wdFormatDOSText = $00000004;
  wdFormatDOSTextLineBreaks = $00000005;
  wdFormatRTF = $00000006;
  wdFormatUnicodeText = $00000007;
  wdFormatEncodedText = $00000007;
  wdFormatHTML = $00000008;
  wdFormatWebArchive = $00000009;
  wdFormatFilteredHTML = $0000000A;
  wdFormatXML = $0000000B;
  wdFormatXMLDocument = $0000000C;
  wdFormatXMLDocumentMacroEnabled = $0000000D;
  wdFormatXMLTemplate = $0000000E;
  wdFormatXMLTemplateMacroEnabled = $0000000F;
  wdFormatDocumentDefault = $00000010;
  wdFormatPDF = $00000011;
  wdFormatXPS = $00000012;
  wdFormatFlatXML = $00000013;
  wdFormatFlatXMLMacroEnabled = $00000014;
  wdFormatFlatXMLTemplate = $00000015;
  wdFormatFlatXMLTemplateMacroEnabled = $00000016;
  wdFormatOpenDocumentText = $00000017;
  wdFormatStrictOpenXMLDocument = $00000018;


// Constants for enum WdOpenFormat
const
  wdOpenFormatAuto = $00000000;
  wdOpenFormatDocument = $00000001;
  wdOpenFormatTemplate = $00000002;
  wdOpenFormatRTF = $00000003;
  wdOpenFormatText = $00000004;
  wdOpenFormatUnicodeText = $00000005;
  wdOpenFormatEncodedText = $00000005;
  wdOpenFormatAllWord = $00000006;
  wdOpenFormatWebPages = $00000007;
  wdOpenFormatXML = $00000008;
  wdOpenFormatXMLDocument = $00000009;
  wdOpenFormatXMLDocumentMacroEnabled = $0000000A;
  wdOpenFormatXMLTemplate = $0000000B;
  wdOpenFormatXMLTemplateMacroEnabled = $0000000C;
  wdOpenFormatDocument97 = $00000001;
  wdOpenFormatTemplate97 = $00000002;
  wdOpenFormatAllWordTemplates = $0000000D;
  wdOpenFormatXMLDocumentSerialized = $0000000E;
  wdOpenFormatXMLDocumentMacroEnabledSerialized = $0000000F;
  wdOpenFormatXMLTemplateSerialized = $00000010;
  wdOpenFormatXMLTemplateMacroEnabledSerialized = $00000011;
  wdOpenFormatOpenDocumentText = $00000012;


// Constants for enum WdHeaderFooterIndex
const
  wdHeaderFooterPrimary = $00000001;
  wdHeaderFooterFirstPage = $00000002;
  wdHeaderFooterEvenPages = $00000003;


// Constants for enum WdTocFormat
const
  wdTOCTemplate = $00000000;
  wdTOCClassic = $00000001;
  wdTOCDistinctive = $00000002;
  wdTOCFancy = $00000003;
  wdTOCModern = $00000004;
  wdTOCFormal = $00000005;
  wdTOCSimple = $00000006;


// Constants for enum WdTofFormat
const
  wdTOFTemplate = $00000000;
  wdTOFClassic = $00000001;
  wdTOFDistinctive = $00000002;
  wdTOFCentered = $00000003;
  wdTOFFormal = $00000004;
  wdTOFSimple = $00000005;


// Constants for enum WdToaFormat
const
  wdTOATemplate = $00000000;
  wdTOAClassic = $00000001;
  wdTOADistinctive = $00000002;
  wdTOAFormal = $00000003;
  wdTOASimple = $00000004;


// Constants for enum WdLineStyle
const
  wdLineStyleNone = $00000000;
  wdLineStyleSingle = $00000001;
  wdLineStyleDot = $00000002;
  wdLineStyleDashSmallGap = $00000003;
  wdLineStyleDashLargeGap = $00000004;
  wdLineStyleDashDot = $00000005;
  wdLineStyleDashDotDot = $00000006;
  wdLineStyleDouble = $00000007;
  wdLineStyleTriple = $00000008;
  wdLineStyleThinThickSmallGap = $00000009;
  wdLineStyleThickThinSmallGap = $0000000A;
  wdLineStyleThinThickThinSmallGap = $0000000B;
  wdLineStyleThinThickMedGap = $0000000C;
  wdLineStyleThickThinMedGap = $0000000D;
  wdLineStyleThinThickThinMedGap = $0000000E;
  wdLineStyleThinThickLargeGap = $0000000F;
  wdLineStyleThickThinLargeGap = $00000010;
  wdLineStyleThinThickThinLargeGap = $00000011;
  wdLineStyleSingleWavy = $00000012;
  wdLineStyleDoubleWavy = $00000013;
  wdLineStyleDashDotStroked = $00000014;
  wdLineStyleEmboss3D = $00000015;
  wdLineStyleEngrave3D = $00000016;
  wdLineStyleOutset = $00000017;
  wdLineStyleInset = $00000018;


// Constants for enum WdLineWidth
const
  wdLineWidth025pt = $00000002;
  wdLineWidth050pt = $00000004;
  wdLineWidth075pt = $00000006;
  wdLineWidth100pt = $00000008;
  wdLineWidth150pt = $0000000C;
  wdLineWidth225pt = $00000012;
  wdLineWidth300pt = $00000018;
  wdLineWidth450pt = $00000024;
  wdLineWidth600pt = $00000030;


// Constants for enum WdBreakType
const
  wdSectionBreakNextPage = $00000002;
  wdSectionBreakContinuous = $00000003;
  wdSectionBreakEvenPage = $00000004;
  wdSectionBreakOddPage = $00000005;
  wdLineBreak = $00000006;
  wdPageBreak = $00000007;
  wdColumnBreak = $00000008;
  wdLineBreakClearLeft = $00000009;
  wdLineBreakClearRight = $0000000A;
  wdTextWrappingBreak = $0000000B;


// Constants for enum WdTabLeader
const
  wdTabLeaderSpaces = $00000000;
  wdTabLeaderDots = $00000001;
  wdTabLeaderDashes = $00000002;
  wdTabLeaderLines = $00000003;
  wdTabLeaderHeavy = $00000004;
  wdTabLeaderMiddleDot = $00000005;


// Constants for enum WdTabLeaderHID
const
  emptyenum________ = $00000000;


// Constants for enum WdMeasurementUnits
const
  wdInches = $00000000;
  wdCentimeters = $00000001;
  wdMillimeters = $00000002;
  wdPoints = $00000003;
  wdPicas = $00000004;


// Constants for enum WdMeasurementUnitsHID
const
  emptyenum_________ = $00000000;


// Constants for enum WdDropPosition
const
  wdDropNone = $00000000;
  wdDropNormal = $00000001;
  wdDropMargin = $00000002;


// Constants for enum WdNumberingRule
const
  wdRestartContinuous = $00000000;
  wdRestartSection = $00000001;
  wdRestartPage = $00000002;


// Constants for enum WdFootnoteLocation
const
  wdBottomOfPage = $00000000;
  wdBeneathText = $00000001;


// Constants for enum WdEndnoteLocation
const
  wdEndOfSection = $00000000;
  wdEndOfDocument = $00000001;


// Constants for enum WdSortSeparator
const
  wdSortSeparateByTabs = $00000000;
  wdSortSeparateByCommas = $00000001;
  wdSortSeparateByDefaultTableSeparator = $00000002;


// Constants for enum WdTableFieldSeparator
const
  wdSeparateByParagraphs = $00000000;
  wdSeparateByTabs = $00000001;
  wdSeparateByCommas = $00000002;
  wdSeparateByDefaultListSeparator = $00000003;


// Constants for enum WdSortFieldType
const
  wdSortFieldAlphanumeric = $00000000;
  wdSortFieldNumeric = $00000001;
  wdSortFieldDate = $00000002;
  wdSortFieldSyllable = $00000003;
  wdSortFieldJapanJIS = $00000004;
  wdSortFieldStroke = $00000005;
  wdSortFieldKoreaKS = $00000006;


// Constants for enum WdSortFieldTypeHID
const
  emptyenum__________ = $00000000;


// Constants for enum WdSortOrder
const
  wdSortOrderAscending = $00000000;
  wdSortOrderDescending = $00000001;


// Constants for enum WdTableFormat
const
  wdTableFormatNone = $00000000;
  wdTableFormatSimple1 = $00000001;
  wdTableFormatSimple2 = $00000002;
  wdTableFormatSimple3 = $00000003;
  wdTableFormatClassic1 = $00000004;
  wdTableFormatClassic2 = $00000005;
  wdTableFormatClassic3 = $00000006;
  wdTableFormatClassic4 = $00000007;
  wdTableFormatColorful1 = $00000008;
  wdTableFormatColorful2 = $00000009;
  wdTableFormatColorful3 = $0000000A;
  wdTableFormatColumns1 = $0000000B;
  wdTableFormatColumns2 = $0000000C;
  wdTableFormatColumns3 = $0000000D;
  wdTableFormatColumns4 = $0000000E;
  wdTableFormatColumns5 = $0000000F;
  wdTableFormatGrid1 = $00000010;
  wdTableFormatGrid2 = $00000011;
  wdTableFormatGrid3 = $00000012;
  wdTableFormatGrid4 = $00000013;
  wdTableFormatGrid5 = $00000014;
  wdTableFormatGrid6 = $00000015;
  wdTableFormatGrid7 = $00000016;
  wdTableFormatGrid8 = $00000017;
  wdTableFormatList1 = $00000018;
  wdTableFormatList2 = $00000019;
  wdTableFormatList3 = $0000001A;
  wdTableFormatList4 = $0000001B;
  wdTableFormatList5 = $0000001C;
  wdTableFormatList6 = $0000001D;
  wdTableFormatList7 = $0000001E;
  wdTableFormatList8 = $0000001F;
  wdTableFormat3DEffects1 = $00000020;
  wdTableFormat3DEffects2 = $00000021;
  wdTableFormat3DEffects3 = $00000022;
  wdTableFormatContemporary = $00000023;
  wdTableFormatElegant = $00000024;
  wdTableFormatProfessional = $00000025;
  wdTableFormatSubtle1 = $00000026;
  wdTableFormatSubtle2 = $00000027;
  wdTableFormatWeb1 = $00000028;
  wdTableFormatWeb2 = $00000029;
  wdTableFormatWeb3 = $0000002A;


// Constants for enum WdTableFormatApply
const
  wdTableFormatApplyBorders = $00000001;
  wdTableFormatApplyShading = $00000002;
  wdTableFormatApplyFont = $00000004;
  wdTableFormatApplyColor = $00000008;
  wdTableFormatApplyAutoFit = $00000010;
  wdTableFormatApplyHeadingRows = $00000020;
  wdTableFormatApplyLastRow = $00000040;
  wdTableFormatApplyFirstColumn = $00000080;
  wdTableFormatApplyLastColumn = $00000100;


// Constants for enum WdLanguageID
const
  wdLanguageNone = $00000000;
  wdNoProofing = $00000400;
  wdAfrikaans = $00000436;
  wdAlbanian = $0000041C;
  wdAmharic = $0000045E;
  wdArabicAlgeria = $00001401;
  wdArabicBahrain = $00003C01;
  wdArabicEgypt = $00000C01;
  wdArabicIraq = $00000801;
  wdArabicJordan = $00002C01;
  wdArabicKuwait = $00003401;
  wdArabicLebanon = $00003001;
  wdArabicLibya = $00001001;
  wdArabicMorocco = $00001801;
  wdArabicOman = $00002001;
  wdArabicQatar = $00004001;
  wdArabic = $00000401;
  wdArabicSyria = $00002801;
  wdArabicTunisia = $00001C01;
  wdArabicUAE = $00003801;
  wdArabicYemen = $00002401;
  wdArmenian = $0000042B;
  wdAssamese = $0000044D;
  wdAzeriCyrillic = $0000082C;
  wdAzeriLatin = $0000042C;
  wdBasque = $0000042D;
  wdByelorussian = $00000423;
  wdBengali = $00000445;
  wdBulgarian = $00000402;
  wdBurmese = $00000455;
  wdCatalan = $00000403;
  wdCherokee = $0000045C;
  wdChineseHongKongSAR = $00000C04;
  wdChineseMacaoSAR = $00001404;
  wdSimplifiedChinese = $00000804;
  wdChineseSingapore = $00001004;
  wdTraditionalChinese = $00000404;
  wdCroatian = $0000041A;
  wdCzech = $00000405;
  wdDanish = $00000406;
  wdDivehi = $00000465;
  wdBelgianDutch = $00000813;
  wdDutch = $00000413;
  wdEdo = $00000466;
  wdEnglishAUS = $00000C09;
  wdEnglishBelize = $00002809;
  wdEnglishCanadian = $00001009;
  wdEnglishCaribbean = $00002409;
  wdEnglishIreland = $00001809;
  wdEnglishJamaica = $00002009;
  wdEnglishNewZealand = $00001409;
  wdEnglishPhilippines = $00003409;
  wdEnglishSouthAfrica = $00001C09;
  wdEnglishTrinidadTobago = $00002C09;
  wdEnglishUK = $00000809;
  wdEnglishUS = $00000409;
  wdEnglishZimbabwe = $00003009;
  wdEnglishIndonesia = $00003809;
  wdEstonian = $00000425;
  wdFaeroese = $00000438;
  wdPersian = $00000429;
  wdFilipino = $00000464;
  wdFinnish = $0000040B;
  wdFulfulde = $00000467;
  wdBelgianFrench = $0000080C;
  wdFrenchCameroon = $00002C0C;
  wdFrenchCanadian = $00000C0C;
  wdFrenchCotedIvoire = $0000300C;
  wdFrench = $0000040C;
  wdFrenchLuxembourg = $0000140C;
  wdFrenchMali = $0000340C;
  wdFrenchMonaco = $0000180C;
  wdFrenchReunion = $0000200C;
  wdFrenchSenegal = $0000280C;
  wdFrenchMorocco = $0000380C;
  wdFrenchHaiti = $00003C0C;
  wdSwissFrench = $0000100C;
  wdFrenchWestIndies = $00001C0C;
  wdFrenchCongoDRC = $0000240C;
  wdFrisianNetherlands = $00000462;
  wdGaelicIreland = $0000083C;
  wdGaelicScotland = $0000043C;
  wdGalician = $00000456;
  wdGeorgian = $00000437;
  wdGermanAustria = $00000C07;
  wdGerman = $00000407;
  wdGermanLiechtenstein = $00001407;
  wdGermanLuxembourg = $00001007;
  wdSwissGerman = $00000807;
  wdGreek = $00000408;
  wdGuarani = $00000474;
  wdGujarati = $00000447;
  wdHausa = $00000468;
  wdHawaiian = $00000475;
  wdHebrew = $0000040D;
  wdHindi = $00000439;
  wdHungarian = $0000040E;
  wdIbibio = $00000469;
  wdIcelandic = $0000040F;
  wdIgbo = $00000470;
  wdIndonesian = $00000421;
  wdInuktitut = $0000045D;
  wdItalian = $00000410;
  wdSwissItalian = $00000810;
  wdJapanese = $00000411;
  wdKannada = $0000044B;
  wdKanuri = $00000471;
  wdKashmiri = $00000460;
  wdKazakh = $0000043F;
  wdKhmer = $00000453;
  wdKirghiz = $00000440;
  wdKonkani = $00000457;
  wdKorean = $00000412;
  wdKyrgyz = $00000440;
  wdLao = $00000454;
  wdLatin = $00000476;
  wdLatvian = $00000426;
  wdLithuanian = $00000427;
  wdMacedonianFYROM = $0000042F;
  wdMalaysian = $0000043E;
  wdMalayBruneiDarussalam = $0000083E;
  wdMalayalam = $0000044C;
  wdMaltese = $0000043A;
  wdManipuri = $00000458;
  wdMarathi = $0000044E;
  wdMongolian = $00000450;
  wdNepali = $00000461;
  wdNorwegianBokmol = $00000414;
  wdNorwegianNynorsk = $00000814;
  wdOriya = $00000448;
  wdOromo = $00000472;
  wdPashto = $00000463;
  wdPolish = $00000415;
  wdPortugueseBrazil = $00000416;
  wdPortuguese = $00000816;
  wdPunjabi = $00000446;
  wdRhaetoRomanic = $00000417;
  wdRomanianMoldova = $00000818;
  wdRomanian = $00000418;
  wdRussianMoldova = $00000819;
  wdRussian = $00000419;
  wdSamiLappish = $0000043B;
  wdSanskrit = $0000044F;
  wdSerbianCyrillic = $00000C1A;
  wdSerbianLatin = $0000081A;
  wdSinhalese = $0000045B;
  wdSindhi = $00000459;
  wdSindhiPakistan = $00000859;
  wdSlovak = $0000041B;
  wdSlovenian = $00000424;
  wdSomali = $00000477;
  wdSorbian = $0000042E;
  wdSpanishArgentina = $00002C0A;
  wdSpanishBolivia = $0000400A;
  wdSpanishChile = $0000340A;
  wdSpanishColombia = $0000240A;
  wdSpanishCostaRica = $0000140A;
  wdSpanishDominicanRepublic = $00001C0A;
  wdSpanishEcuador = $0000300A;
  wdSpanishElSalvador = $0000440A;
  wdSpanishGuatemala = $0000100A;
  wdSpanishHonduras = $0000480A;
  wdMexicanSpanish = $0000080A;
  wdSpanishNicaragua = $00004C0A;
  wdSpanishPanama = $0000180A;
  wdSpanishParaguay = $00003C0A;
  wdSpanishPeru = $0000280A;
  wdSpanishPuertoRico = $0000500A;
  wdSpanishModernSort = $00000C0A;
  wdSpanish = $0000040A;
  wdSpanishUruguay = $0000380A;
  wdSpanishVenezuela = $0000200A;
  wdSesotho = $00000430;
  wdSutu = $00000430;
  wdSwahili = $00000441;
  wdSwedishFinland = $0000081D;
  wdSwedish = $0000041D;
  wdSyriac = $0000045A;
  wdTajik = $00000428;
  wdTamazight = $0000045F;
  wdTamazightLatin = $0000085F;
  wdTamil = $00000449;
  wdTatar = $00000444;
  wdTelugu = $0000044A;
  wdThai = $0000041E;
  wdTibetan = $00000451;
  wdTigrignaEthiopic = $00000473;
  wdTigrignaEritrea = $00000873;
  wdTsonga = $00000431;
  wdTswana = $00000432;
  wdTurkish = $0000041F;
  wdTurkmen = $00000442;
  wdUkrainian = $00000422;
  wdUrdu = $00000420;
  wdUzbekCyrillic = $00000843;
  wdUzbekLatin = $00000443;
  wdVenda = $00000433;
  wdVietnamese = $0000042A;
  wdWelsh = $00000452;
  wdXhosa = $00000434;
  wdYi = $00000478;
  wdYiddish = $0000043D;
  wdYoruba = $0000046A;
  wdZulu = $00000435;


// Constants for enum WdFieldType
const
  wdFieldEmpty = $FFFFFFFF;
  wdFieldRef = $00000003;
  wdFieldIndexEntry = $00000004;
  wdFieldFootnoteRef = $00000005;
  wdFieldSet = $00000006;
  wdFieldIf = $00000007;
  wdFieldIndex = $00000008;
  wdFieldTOCEntry = $00000009;
  wdFieldStyleRef = $0000000A;
  wdFieldRefDoc = $0000000B;
  wdFieldSequence = $0000000C;
  wdFieldTOC = $0000000D;
  wdFieldInfo = $0000000E;
  wdFieldTitle = $0000000F;
  wdFieldSubject = $00000010;
  wdFieldAuthor = $00000011;
  wdFieldKeyWord = $00000012;
  wdFieldComments = $00000013;
  wdFieldLastSavedBy = $00000014;
  wdFieldCreateDate = $00000015;
  wdFieldSaveDate = $00000016;
  wdFieldPrintDate = $00000017;
  wdFieldRevisionNum = $00000018;
  wdFieldEditTime = $00000019;
  wdFieldNumPages = $0000001A;
  wdFieldNumWords = $0000001B;
  wdFieldNumChars = $0000001C;
  wdFieldFileName = $0000001D;
  wdFieldTemplate = $0000001E;
  wdFieldDate = $0000001F;
  wdFieldTime = $00000020;
  wdFieldPage = $00000021;
  wdFieldExpression = $00000022;
  wdFieldQuote = $00000023;
  wdFieldInclude = $00000024;
  wdFieldPageRef = $00000025;
  wdFieldAsk = $00000026;
  wdFieldFillIn = $00000027;
  wdFieldData = $00000028;
  wdFieldNext = $00000029;
  wdFieldNextIf = $0000002A;
  wdFieldSkipIf = $0000002B;
  wdFieldMergeRec = $0000002C;
  wdFieldDDE = $0000002D;
  wdFieldDDEAuto = $0000002E;
  wdFieldGlossary = $0000002F;
  wdFieldPrint = $00000030;
  wdFieldFormula = $00000031;
  wdFieldGoToButton = $00000032;
  wdFieldMacroButton = $00000033;
  wdFieldAutoNumOutline = $00000034;
  wdFieldAutoNumLegal = $00000035;
  wdFieldAutoNum = $00000036;
  wdFieldImport = $00000037;
  wdFieldLink = $00000038;
  wdFieldSymbol = $00000039;
  wdFieldEmbed = $0000003A;
  wdFieldMergeField = $0000003B;
  wdFieldUserName = $0000003C;
  wdFieldUserInitials = $0000003D;
  wdFieldUserAddress = $0000003E;
  wdFieldBarCode = $0000003F;
  wdFieldDocVariable = $00000040;
  wdFieldSection = $00000041;
  wdFieldSectionPages = $00000042;
  wdFieldIncludePicture = $00000043;
  wdFieldIncludeText = $00000044;
  wdFieldFileSize = $00000045;
  wdFieldFormTextInput = $00000046;
  wdFieldFormCheckBox = $00000047;
  wdFieldNoteRef = $00000048;
  wdFieldTOA = $00000049;
  wdFieldTOAEntry = $0000004A;
  wdFieldMergeSeq = $0000004B;
  wdFieldPrivate = $0000004D;
  wdFieldDatabase = $0000004E;
  wdFieldAutoText = $0000004F;
  wdFieldCompare = $00000050;
  wdFieldAddin = $00000051;
  wdFieldSubscriber = $00000052;
  wdFieldFormDropDown = $00000053;
  wdFieldAdvance = $00000054;
  wdFieldDocProperty = $00000055;
  wdFieldOCX = $00000057;
  wdFieldHyperlink = $00000058;
  wdFieldAutoTextList = $00000059;
  wdFieldListNum = $0000005A;
  wdFieldHTMLActiveX = $0000005B;
  wdFieldBidiOutline = $0000005C;
  wdFieldAddressBlock = $0000005D;
  wdFieldGreetingLine = $0000005E;
  wdFieldShape = $0000005F;
  wdFieldCitation = $00000060;
  wdFieldBibliography = $00000061;
  wdFieldMergeBarcode = $00000062;
  wdFieldDisplayBarcode = $00000063;


// Constants for enum WdBuiltinStyle
const
  wdStyleNormal = $FFFFFFFF;
  wdStyleEnvelopeAddress = $FFFFFFDB;
  wdStyleEnvelopeReturn = $FFFFFFDA;
  wdStyleBodyText = $FFFFFFBD;
  wdStyleHeading1 = $FFFFFFFE;
  wdStyleHeading2 = $FFFFFFFD;
  wdStyleHeading3 = $FFFFFFFC;
  wdStyleHeading4 = $FFFFFFFB;
  wdStyleHeading5 = $FFFFFFFA;
  wdStyleHeading6 = $FFFFFFF9;
  wdStyleHeading7 = $FFFFFFF8;
  wdStyleHeading8 = $FFFFFFF7;
  wdStyleHeading9 = $FFFFFFF6;
  wdStyleIndex1 = $FFFFFFF5;
  wdStyleIndex2 = $FFFFFFF4;
  wdStyleIndex3 = $FFFFFFF3;
  wdStyleIndex4 = $FFFFFFF2;
  wdStyleIndex5 = $FFFFFFF1;
  wdStyleIndex6 = $FFFFFFF0;
  wdStyleIndex7 = $FFFFFFEF;
  wdStyleIndex8 = $FFFFFFEE;
  wdStyleIndex9 = $FFFFFFED;
  wdStyleTOC1 = $FFFFFFEC;
  wdStyleTOC2 = $FFFFFFEB;
  wdStyleTOC3 = $FFFFFFEA;
  wdStyleTOC4 = $FFFFFFE9;
  wdStyleTOC5 = $FFFFFFE8;
  wdStyleTOC6 = $FFFFFFE7;
  wdStyleTOC7 = $FFFFFFE6;
  wdStyleTOC8 = $FFFFFFE5;
  wdStyleTOC9 = $FFFFFFE4;
  wdStyleNormalIndent = $FFFFFFE3;
  wdStyleFootnoteText = $FFFFFFE2;
  wdStyleCommentText = $FFFFFFE1;
  wdStyleHeader = $FFFFFFE0;
  wdStyleFooter = $FFFFFFDF;
  wdStyleIndexHeading = $FFFFFFDE;
  wdStyleCaption = $FFFFFFDD;
  wdStyleTableOfFigures = $FFFFFFDC;
  wdStyleFootnoteReference = $FFFFFFD9;
  wdStyleCommentReference = $FFFFFFD8;
  wdStyleLineNumber = $FFFFFFD7;
  wdStylePageNumber = $FFFFFFD6;
  wdStyleEndnoteReference = $FFFFFFD5;
  wdStyleEndnoteText = $FFFFFFD4;
  wdStyleTableOfAuthorities = $FFFFFFD3;
  wdStyleMacroText = $FFFFFFD2;
  wdStyleTOAHeading = $FFFFFFD1;
  wdStyleList = $FFFFFFD0;
  wdStyleListBullet = $FFFFFFCF;
  wdStyleListNumber = $FFFFFFCE;
  wdStyleList2 = $FFFFFFCD;
  wdStyleList3 = $FFFFFFCC;
  wdStyleList4 = $FFFFFFCB;
  wdStyleList5 = $FFFFFFCA;
  wdStyleListBullet2 = $FFFFFFC9;
  wdStyleListBullet3 = $FFFFFFC8;
  wdStyleListBullet4 = $FFFFFFC7;
  wdStyleListBullet5 = $FFFFFFC6;
  wdStyleListNumber2 = $FFFFFFC5;
  wdStyleListNumber3 = $FFFFFFC4;
  wdStyleListNumber4 = $FFFFFFC3;
  wdStyleListNumber5 = $FFFFFFC2;
  wdStyleTitle = $FFFFFFC1;
  wdStyleClosing = $FFFFFFC0;
  wdStyleSignature = $FFFFFFBF;
  wdStyleDefaultParagraphFont = $FFFFFFBE;
  wdStyleBodyTextIndent = $FFFFFFBC;
  wdStyleListContinue = $FFFFFFBB;
  wdStyleListContinue2 = $FFFFFFBA;
  wdStyleListContinue3 = $FFFFFFB9;
  wdStyleListContinue4 = $FFFFFFB8;
  wdStyleListContinue5 = $FFFFFFB7;
  wdStyleMessageHeader = $FFFFFFB6;
  wdStyleSubtitle = $FFFFFFB5;
  wdStyleSalutation = $FFFFFFB4;
  wdStyleDate = $FFFFFFB3;
  wdStyleBodyTextFirstIndent = $FFFFFFB2;
  wdStyleBodyTextFirstIndent2 = $FFFFFFB1;
  wdStyleNoteHeading = $FFFFFFB0;
  wdStyleBodyText2 = $FFFFFFAF;
  wdStyleBodyText3 = $FFFFFFAE;
  wdStyleBodyTextIndent2 = $FFFFFFAD;
  wdStyleBodyTextIndent3 = $FFFFFFAC;
  wdStyleBlockQuotation = $FFFFFFAB;
  wdStyleHyperlink = $FFFFFFAA;
  wdStyleHyperlinkFollowed = $FFFFFFA9;
  wdStyleStrong = $FFFFFFA8;
  wdStyleEmphasis = $FFFFFFA7;
  wdStyleNavPane = $FFFFFFA6;
  wdStylePlainText = $FFFFFFA5;
  wdStyleHtmlNormal = $FFFFFFA1;
  wdStyleHtmlAcronym = $FFFFFFA0;
  wdStyleHtmlAddress = $FFFFFF9F;
  wdStyleHtmlCite = $FFFFFF9E;
  wdStyleHtmlCode = $FFFFFF9D;
  wdStyleHtmlDfn = $FFFFFF9C;
  wdStyleHtmlKbd = $FFFFFF9B;
  wdStyleHtmlPre = $FFFFFF9A;
  wdStyleHtmlSamp = $FFFFFF99;
  wdStyleHtmlTt = $FFFFFF98;
  wdStyleHtmlVar = $FFFFFF97;
  wdStyleNormalTable = $FFFFFF96;
  wdStyleNormalObject = $FFFFFF62;
  wdStyleTableLightShading = $FFFFFF61;
  wdStyleTableLightList = $FFFFFF60;
  wdStyleTableLightGrid = $FFFFFF5F;
  wdStyleTableMediumShading1 = $FFFFFF5E;
  wdStyleTableMediumShading2 = $FFFFFF5D;
  wdStyleTableMediumList1 = $FFFFFF5C;
  wdStyleTableMediumList2 = $FFFFFF5B;
  wdStyleTableMediumGrid1 = $FFFFFF5A;
  wdStyleTableMediumGrid2 = $FFFFFF59;
  wdStyleTableMediumGrid3 = $FFFFFF58;
  wdStyleTableDarkList = $FFFFFF57;
  wdStyleTableColorfulShading = $FFFFFF56;
  wdStyleTableColorfulList = $FFFFFF55;
  wdStyleTableColorfulGrid = $FFFFFF54;
  wdStyleTableLightShadingAccent1 = $FFFFFF53;
  wdStyleTableLightListAccent1 = $FFFFFF52;
  wdStyleTableLightGridAccent1 = $FFFFFF51;
  wdStyleTableMediumShading1Accent1 = $FFFFFF50;
  wdStyleTableMediumShading2Accent1 = $FFFFFF4F;
  wdStyleTableMediumList1Accent1 = $FFFFFF4E;
  wdStyleListParagraph = $FFFFFF4C;
  wdStyleQuote = $FFFFFF4B;
  wdStyleIntenseQuote = $FFFFFF4A;
  wdStyleSubtleEmphasis = $FFFFFEFB;
  wdStyleIntenseEmphasis = $FFFFFEFA;
  wdStyleSubtleReference = $FFFFFEF9;
  wdStyleIntenseReference = $FFFFFEF8;
  wdStyleBookTitle = $FFFFFEF7;
  wdStyleBibliography = $FFFFFEF6;
  wdStyleTocHeading = $FFFFFEF5;


// Constants for enum WdWordDialogTab
const
  wdDialogToolsOptionsTabView = $000000CC;
  wdDialogToolsOptionsTabGeneral = $000000CB;
  wdDialogToolsOptionsTabEdit = $000000E0;
  wdDialogToolsOptionsTabPrint = $000000D0;
  wdDialogToolsOptionsTabSave = $000000D1;
  wdDialogToolsOptionsTabProofread = $000000D3;
  wdDialogToolsOptionsTabTrackChanges = $00000182;
  wdDialogToolsOptionsTabUserInfo = $000000D5;
  wdDialogToolsOptionsTabCompatibility = $0000020D;
  wdDialogToolsOptionsTabTypography = $000002E3;
  wdDialogToolsOptionsTabFileLocations = $000000E1;
  wdDialogToolsOptionsTabFuzzy = $00000316;
  wdDialogToolsOptionsTabHangulHanjaConversion = $00000312;
  wdDialogToolsOptionsTabBidi = $00000405;
  wdDialogToolsOptionsTabSecurity = $00000551;
  wdDialogFilePageSetupTabMargins = $000249F0;
  wdDialogFilePageSetupTabPaper = $000249F1;
  wdDialogFilePageSetupTabLayout = $000249F3;
  wdDialogFilePageSetupTabCharsLines = $000249F4;
  wdDialogInsertSymbolTabSymbols = $00030D40;
  wdDialogInsertSymbolTabSpecialCharacters = $00030D41;
  wdDialogNoteOptionsTabAllFootnotes = $000493E0;
  wdDialogNoteOptionsTabAllEndnotes = $000493E1;
  wdDialogInsertIndexAndTablesTabIndex = $00061A80;
  wdDialogInsertIndexAndTablesTabTableOfContents = $00061A81;
  wdDialogInsertIndexAndTablesTabTableOfFigures = $00061A82;
  wdDialogInsertIndexAndTablesTabTableOfAuthorities = $00061A83;
  wdDialogOrganizerTabStyles = $0007A120;
  wdDialogOrganizerTabAutoText = $0007A121;
  wdDialogOrganizerTabCommandBars = $0007A122;
  wdDialogOrganizerTabMacros = $0007A123;
  wdDialogFormatFontTabFont = $000927C0;
  wdDialogFormatFontTabCharacterSpacing = $000927C1;
  wdDialogFormatFontTabAnimation = $000927C2;
  wdDialogFormatBordersAndShadingTabBorders = $000AAE60;
  wdDialogFormatBordersAndShadingTabPageBorder = $000AAE61;
  wdDialogFormatBordersAndShadingTabShading = $000AAE62;
  wdDialogToolsEnvelopesAndLabelsTabEnvelopes = $000C3500;
  wdDialogToolsEnvelopesAndLabelsTabLabels = $000C3501;
  wdDialogFormatParagraphTabIndentsAndSpacing = $000F4240;
  wdDialogFormatParagraphTabTextFlow = $000F4241;
  wdDialogFormatParagraphTabTeisai = $000F4242;
  wdDialogFormatDrawingObjectTabColorsAndLines = $00124F80;
  wdDialogFormatDrawingObjectTabSize = $00124F81;
  wdDialogFormatDrawingObjectTabPosition = $00124F82;
  wdDialogFormatDrawingObjectTabWrapping = $00124F83;
  wdDialogFormatDrawingObjectTabPicture = $00124F84;
  wdDialogFormatDrawingObjectTabTextbox = $00124F85;
  wdDialogFormatDrawingObjectTabWeb = $00124F86;
  wdDialogFormatDrawingObjectTabHR = $00124F87;
  wdDialogToolsAutoCorrectExceptionsTabFirstLetter = $00155CC0;
  wdDialogToolsAutoCorrectExceptionsTabInitialCaps = $00155CC1;
  wdDialogToolsAutoCorrectExceptionsTabHangulAndAlphabet = $00155CC2;
  wdDialogToolsAutoCorrectExceptionsTabIac = $00155CC3;
  wdDialogFormatBulletsAndNumberingTabBulleted = $0016E360;
  wdDialogFormatBulletsAndNumberingTabNumbered = $0016E361;
  wdDialogFormatBulletsAndNumberingTabOutlineNumbered = $0016E362;
  wdDialogLetterWizardTabLetterFormat = $00186A00;
  wdDialogLetterWizardTabRecipientInfo = $00186A01;
  wdDialogLetterWizardTabOtherElements = $00186A02;
  wdDialogLetterWizardTabSenderInfo = $00186A03;
  wdDialogToolsAutoManagerTabAutoCorrect = $0019F0A0;
  wdDialogToolsAutoManagerTabAutoFormatAsYouType = $0019F0A1;
  wdDialogToolsAutoManagerTabAutoText = $0019F0A2;
  wdDialogToolsAutoManagerTabAutoFormat = $0019F0A3;
  wdDialogToolsAutoManagerTabSmartTags = $0019F0A4;
  wdDialogTablePropertiesTabTable = $001B7740;
  wdDialogTablePropertiesTabRow = $001B7741;
  wdDialogTablePropertiesTabColumn = $001B7742;
  wdDialogTablePropertiesTabCell = $001B7743;
  wdDialogEmailOptionsTabSignature = $001CFDE0;
  wdDialogEmailOptionsTabStationary = $001CFDE1;
  wdDialogEmailOptionsTabQuoting = $001CFDE2;
  wdDialogWebOptionsBrowsers = $001E8480;
  wdDialogWebOptionsGeneral = $001E8480;
  wdDialogWebOptionsFiles = $001E8481;
  wdDialogWebOptionsPictures = $001E8482;
  wdDialogWebOptionsEncoding = $001E8483;
  wdDialogWebOptionsFonts = $001E8484;
  wdDialogToolsOptionsTabAcetate = $000004F2;
  wdDialogTemplates = $00200B20;
  wdDialogTemplatesXMLSchema = $00200B21;
  wdDialogTemplatesXMLExpansionPacks = $00200B22;
  wdDialogTemplatesLinkedCSS = $00200B23;
  wdDialogStyleManagementTabEdit = $002191C0;
  wdDialogStyleManagementTabRecommend = $002191C1;
  wdDialogStyleManagementTabRestrict = $002191C2;


// Constants for enum WdWordDialogTabHID
const
  wdDialogFilePageSetupTabPaperSize = $000249F1;
  wdDialogFilePageSetupTabPaperSource = $000249F2;


// Constants for enum WdWordDialog
const
  wdDialogHelpAbout = $00000009;
  wdDialogHelpWordPerfectHelp = $0000000A;
  wdDialogDocumentStatistics = $0000004E;
  wdDialogFileNew = $0000004F;
  wdDialogFileOpen = $00000050;
  wdDialogMailMergeOpenDataSource = $00000051;
  wdDialogMailMergeOpenHeaderSource = $00000052;
  wdDialogFileSaveAs = $00000054;
  wdDialogFileSummaryInfo = $00000056;
  wdDialogToolsTemplates = $00000057;
  wdDialogFilePrint = $00000058;
  wdDialogFilePrintSetup = $00000061;
  wdDialogFileFind = $00000063;
  wdDialogFormatAddrFonts = $00000067;
  wdDialogEditPasteSpecial = $0000006F;
  wdDialogEditFind = $00000070;
  wdDialogEditReplace = $00000075;
  wdDialogEditStyle = $00000078;
  wdDialogEditLinks = $0000007C;
  wdDialogEditObject = $0000007D;
  wdDialogTableToText = $00000080;
  wdDialogTextToTable = $0000007F;
  wdDialogTableInsertTable = $00000081;
  wdDialogTableInsertCells = $00000082;
  wdDialogTableInsertRow = $00000083;
  wdDialogTableDeleteCells = $00000085;
  wdDialogTableSplitCells = $00000089;
  wdDialogTableRowHeight = $0000008E;
  wdDialogTableColumnWidth = $0000008F;
  wdDialogToolsCustomize = $00000098;
  wdDialogInsertBreak = $0000009F;
  wdDialogInsertSymbol = $000000A2;
  wdDialogInsertPicture = $000000A3;
  wdDialogInsertFile = $000000A4;
  wdDialogInsertDateTime = $000000A5;
  wdDialogInsertField = $000000A6;
  wdDialogInsertMergeField = $000000A7;
  wdDialogInsertBookmark = $000000A8;
  wdDialogMarkIndexEntry = $000000A9;
  wdDialogInsertIndex = $000000AA;
  wdDialogInsertTableOfContents = $000000AB;
  wdDialogInsertObject = $000000AC;
  wdDialogToolsCreateEnvelope = $000000AD;
  wdDialogFormatFont = $000000AE;
  wdDialogFormatParagraph = $000000AF;
  wdDialogFormatSectionLayout = $000000B0;
  wdDialogFormatColumns = $000000B1;
  wdDialogFileDocumentLayout = $000000B2;
  wdDialogFilePageSetup = $000000B2;
  wdDialogFormatTabs = $000000B3;
  wdDialogFormatStyle = $000000B4;
  wdDialogFormatDefineStyleFont = $000000B5;
  wdDialogFormatDefineStylePara = $000000B6;
  wdDialogFormatDefineStyleTabs = $000000B7;
  wdDialogFormatDefineStyleFrame = $000000B8;
  wdDialogFormatDefineStyleBorders = $000000B9;
  wdDialogFormatDefineStyleLang = $000000BA;
  wdDialogFormatPicture = $000000BB;
  wdDialogToolsLanguage = $000000BC;
  wdDialogFormatBordersAndShading = $000000BD;
  wdDialogFormatFrame = $000000BE;
  wdDialogToolsThesaurus = $000000C2;
  wdDialogToolsHyphenation = $000000C3;
  wdDialogToolsBulletsNumbers = $000000C4;
  wdDialogToolsHighlightChanges = $000000C5;
  wdDialogToolsRevisions = $000000C5;
  wdDialogToolsCompareDocuments = $000000C6;
  wdDialogTableSort = $000000C7;
  wdDialogToolsOptionsGeneral = $000000CB;
  wdDialogToolsOptionsView = $000000CC;
  wdDialogToolsAdvancedSettings = $000000CE;
  wdDialogToolsOptionsPrint = $000000D0;
  wdDialogToolsOptionsSave = $000000D1;
  wdDialogToolsOptionsSpellingAndGrammar = $000000D3;
  wdDialogToolsOptionsUserInfo = $000000D5;
  wdDialogToolsMacroRecord = $000000D6;
  wdDialogToolsMacro = $000000D7;
  wdDialogWindowActivate = $000000DC;
  wdDialogFormatRetAddrFonts = $000000DD;
  wdDialogOrganizer = $000000DE;
  wdDialogToolsOptionsEdit = $000000E0;
  wdDialogToolsOptionsFileLocations = $000000E1;
  wdDialogToolsWordCount = $000000E4;
  wdDialogControlRun = $000000EB;
  wdDialogInsertPageNumbers = $00000126;
  wdDialogFormatPageNumber = $0000012A;
  wdDialogCopyFile = $0000012C;
  wdDialogFormatChangeCase = $00000142;
  wdDialogUpdateTOC = $0000014B;
  wdDialogInsertDatabase = $00000155;
  wdDialogTableFormula = $0000015C;
  wdDialogFormFieldOptions = $00000161;
  wdDialogInsertCaption = $00000165;
  wdDialogInsertCaptionNumbering = $00000166;
  wdDialogInsertAutoCaption = $00000167;
  wdDialogFormFieldHelp = $00000169;
  wdDialogInsertCrossReference = $0000016F;
  wdDialogInsertFootnote = $00000172;
  wdDialogNoteOptions = $00000175;
  wdDialogToolsAutoCorrect = $0000017A;
  wdDialogToolsOptionsTrackChanges = $00000182;
  wdDialogConvertObject = $00000188;
  wdDialogInsertAddCaption = $00000192;
  wdDialogConnect = $000001A4;
  wdDialogToolsCustomizeKeyboard = $000001B0;
  wdDialogToolsCustomizeMenus = $000001B1;
  wdDialogToolsMergeDocuments = $000001B3;
  wdDialogMarkTableOfContentsEntry = $000001BA;
  wdDialogFileMacPageSetupGX = $000001BC;
  wdDialogFilePrintOneCopy = $000001BD;
  wdDialogEditFrame = $000001CA;
  wdDialogMarkCitation = $000001CF;
  wdDialogTableOfContentsOptions = $000001D6;
  wdDialogInsertTableOfAuthorities = $000001D7;
  wdDialogInsertTableOfFigures = $000001D8;
  wdDialogInsertIndexAndTables = $000001D9;
  wdDialogInsertFormField = $000001E3;
  wdDialogFormatDropCap = $000001E8;
  wdDialogToolsCreateLabels = $000001E9;
  wdDialogToolsProtectDocument = $000001F7;
  wdDialogFormatStyleGallery = $000001F9;
  wdDialogToolsAcceptRejectChanges = $000001FA;
  wdDialogHelpWordPerfectHelpOptions = $000001FF;
  wdDialogToolsUnprotectDocument = $00000209;
  wdDialogToolsOptionsCompatibility = $0000020D;
  wdDialogTableOfCaptionsOptions = $00000227;
  wdDialogTableAutoFormat = $00000233;
  wdDialogMailMergeFindRecord = $00000239;
  wdDialogReviewAfmtRevisions = $0000023A;
  wdDialogViewZoom = $00000241;
  wdDialogToolsProtectSection = $00000242;
  wdDialogFontSubstitution = $00000245;
  wdDialogInsertSubdocument = $00000247;
  wdDialogNewToolbar = $0000024A;
  wdDialogToolsEnvelopesAndLabels = $0000025F;
  wdDialogFormatCallout = $00000262;
  wdDialogTableFormatCell = $00000264;
  wdDialogToolsCustomizeMenuBar = $00000267;
  wdDialogFileRoutingSlip = $00000270;
  wdDialogEditTOACategory = $00000271;
  wdDialogToolsManageFields = $00000277;
  wdDialogDrawSnapToGrid = $00000279;
  wdDialogDrawAlign = $0000027A;
  wdDialogMailMergeCreateDataSource = $00000282;
  wdDialogMailMergeCreateHeaderSource = $00000283;
  wdDialogMailMerge = $000002A4;
  wdDialogMailMergeCheck = $000002A5;
  wdDialogMailMergeHelper = $000002A8;
  wdDialogMailMergeQueryOptions = $000002A9;
  wdDialogFileMacPageSetup = $000002AD;
  wdDialogListCommands = $000002D3;
  wdDialogEditCreatePublisher = $000002DC;
  wdDialogEditSubscribeTo = $000002DD;
  wdDialogEditPublishOptions = $000002DF;
  wdDialogEditSubscribeOptions = $000002E0;
  wdDialogFileMacCustomPageSetupGX = $000002E1;
  wdDialogToolsOptionsTypography = $000002E3;
  wdDialogToolsAutoCorrectExceptions = $000002FA;
  wdDialogToolsOptionsAutoFormatAsYouType = $0000030A;
  wdDialogMailMergeUseAddressBook = $0000030B;
  wdDialogToolsHangulHanjaConversion = $00000310;
  wdDialogToolsOptionsFuzzy = $00000316;
  wdDialogEditGoToOld = $0000032B;
  wdDialogInsertNumber = $0000032C;
  wdDialogLetterWizard = $00000335;
  wdDialogFormatBulletsAndNumbering = $00000338;
  wdDialogToolsSpellingAndGrammar = $0000033C;
  wdDialogToolsCreateDirectory = $00000341;
  wdDialogTableWrapping = $00000356;
  wdDialogFormatTheme = $00000357;
  wdDialogTableProperties = $0000035D;
  wdDialogEmailOptions = $0000035F;
  wdDialogCreateAutoText = $00000368;
  wdDialogToolsAutoSummarize = $0000036A;
  wdDialogToolsGrammarSettings = $00000375;
  wdDialogEditGoTo = $00000380;
  wdDialogWebOptions = $00000382;
  wdDialogInsertHyperlink = $0000039D;
  wdDialogToolsAutoManager = $00000393;
  wdDialogFileVersions = $000003B1;
  wdDialogToolsOptionsAutoFormat = $000003BF;
  wdDialogFormatDrawingObject = $000003C0;
  wdDialogToolsOptions = $000003CE;
  wdDialogFitText = $000003D7;
  wdDialogEditAutoText = $000003D9;
  wdDialogPhoneticGuide = $000003DA;
  wdDialogToolsDictionary = $000003DD;
  wdDialogFileSaveVersion = $000003EF;
  wdDialogToolsOptionsBidi = $00000405;
  wdDialogFrameSetProperties = $00000432;
  wdDialogTableTableOptions = $00000438;
  wdDialogTableCellOptions = $00000439;
  wdDialogIMESetDefault = $00000446;
  wdDialogTCSCTranslator = $00000484;
  wdDialogHorizontalInVertical = $00000488;
  wdDialogTwoLinesInOne = $00000489;
  wdDialogFormatEncloseCharacters = $0000048A;
  wdDialogConsistencyChecker = $00000461;
  wdDialogToolsOptionsSmartTag = $00000573;
  wdDialogFormatStylesCustom = $000004E0;
  wdDialogCSSLinks = $000004ED;
  wdDialogInsertWebComponent = $0000052C;
  wdDialogToolsOptionsEditCopyPaste = $0000054C;
  wdDialogToolsOptionsSecurity = $00000551;
  wdDialogSearch = $00000553;
  wdDialogShowRepairs = $00000565;
  wdDialogMailMergeInsertAsk = $00000FCF;
  wdDialogMailMergeInsertFillIn = $00000FD0;
  wdDialogMailMergeInsertIf = $00000FD1;
  wdDialogMailMergeInsertNextIf = $00000FD5;
  wdDialogMailMergeInsertSet = $00000FD6;
  wdDialogMailMergeInsertSkipIf = $00000FD7;
  wdDialogMailMergeFieldMapping = $00000518;
  wdDialogMailMergeInsertAddressBlock = $00000519;
  wdDialogMailMergeInsertGreetingLine = $0000051A;
  wdDialogMailMergeInsertFields = $0000051B;
  wdDialogMailMergeRecipients = $0000051C;
  wdDialogMailMergeFindRecipient = $0000052E;
  wdDialogMailMergeSetDocumentType = $0000053B;
  wdDialogLabelOptions = $00000557;
  wdDialogXMLElementAttributes = $000005B4;
  wdDialogSchemaLibrary = $00000589;
  wdDialogPermission = $000005BD;
  wdDialogMyPermission = $0000059D;
  wdDialogXMLOptions = $00000591;
  wdDialogFormattingRestrictions = $00000593;
  wdDialogSourceManager = $00000780;
  wdDialogCreateSource = $00000782;
  wdDialogDocumentInspector = $000005CA;
  wdDialogStyleManagement = $0000079C;
  wdDialogInsertSource = $00000848;
  wdDialogOMathRecognizedFunctions = $00000875;
  wdDialogInsertPlaceholder = $0000092C;
  wdDialogBuildingBlockOrganizer = $00000813;
  wdDialogContentControlProperties = $0000095A;
  wdDialogCompatibilityChecker = $00000987;
  wdDialogExportAsFixedFormat = $0000092D;
  wdDialogFileNew2007 = $0000045C;


// Constants for enum WdWordDialogHID
const
  emptyenum___________ = $00000000;


// Constants for enum WdFieldKind
const
  wdFieldKindNone = $00000000;
  wdFieldKindHot = $00000001;
  wdFieldKindWarm = $00000002;
  wdFieldKindCold = $00000003;


// Constants for enum WdTextFormFieldType
const
  wdRegularText = $00000000;
  wdNumberText = $00000001;
  wdDateText = $00000002;
  wdCurrentDateText = $00000003;
  wdCurrentTimeText = $00000004;
  wdCalculationText = $00000005;


// Constants for enum WdChevronConvertRule
const
  wdNeverConvert = $00000000;
  wdAlwaysConvert = $00000001;
  wdAskToNotConvert = $00000002;
  wdAskToConvert = $00000003;


// Constants for enum WdMailMergeMainDocType
const
  wdNotAMergeDocument = $FFFFFFFF;
  wdFormLetters = $00000000;
  wdMailingLabels = $00000001;
  wdEnvelopes = $00000002;
  wdCatalog = $00000003;
  wdEMail = $00000004;
  wdFax = $00000005;
  wdDirectory = $00000003;


// Constants for enum WdMailMergeState
const
  wdNormalDocument = $00000000;
  wdMainDocumentOnly = $00000001;
  wdMainAndDataSource = $00000002;
  wdMainAndHeader = $00000003;
  wdMainAndSourceAndHeader = $00000004;
  wdDataSource = $00000005;


// Constants for enum WdMailMergeDestination
const
  wdSendToNewDocument = $00000000;
  wdSendToPrinter = $00000001;
  wdSendToEmail = $00000002;
  wdSendToFax = $00000003;


// Constants for enum WdMailMergeActiveRecord
const
  wdNoActiveRecord = $FFFFFFFF;
  wdNextRecord = $FFFFFFFE;
  wdPreviousRecord = $FFFFFFFD;
  wdFirstRecord = $FFFFFFFC;
  wdLastRecord = $FFFFFFFB;
  wdFirstDataSourceRecord = $FFFFFFFA;
  wdLastDataSourceRecord = $FFFFFFF9;
  wdNextDataSourceRecord = $FFFFFFF8;
  wdPreviousDataSourceRecord = $FFFFFFF7;


// Constants for enum WdMailMergeDefaultRecord
const
  wdDefaultFirstRecord = $00000001;
  wdDefaultLastRecord = $FFFFFFF0;


// Constants for enum WdMailMergeDataSource
const
  wdNoMergeInfo = $FFFFFFFF;
  wdMergeInfoFromWord = $00000000;
  wdMergeInfoFromAccessDDE = $00000001;
  wdMergeInfoFromExcelDDE = $00000002;
  wdMergeInfoFromMSQueryDDE = $00000003;
  wdMergeInfoFromODBC = $00000004;
  wdMergeInfoFromODSO = $00000005;


// Constants for enum WdMailMergeComparison
const
  wdMergeIfEqual = $00000000;
  wdMergeIfNotEqual = $00000001;
  wdMergeIfLessThan = $00000002;
  wdMergeIfGreaterThan = $00000003;
  wdMergeIfLessThanOrEqual = $00000004;
  wdMergeIfGreaterThanOrEqual = $00000005;
  wdMergeIfIsBlank = $00000006;
  wdMergeIfIsNotBlank = $00000007;


// Constants for enum WdBookmarkSortBy
const
  wdSortByName = $00000000;
  wdSortByLocation = $00000001;


// Constants for enum WdWindowState
const
  wdWindowStateNormal = $00000000;
  wdWindowStateMaximize = $00000001;
  wdWindowStateMinimize = $00000002;


// Constants for enum WdPictureLinkType
const
  wdLinkNone = $00000000;
  wdLinkDataInDoc = $00000001;
  wdLinkDataOnDisk = $00000002;


// Constants for enum WdLinkType
const
  wdLinkTypeOLE = $00000000;
  wdLinkTypePicture = $00000001;
  wdLinkTypeText = $00000002;
  wdLinkTypeReference = $00000003;
  wdLinkTypeInclude = $00000004;
  wdLinkTypeImport = $00000005;
  wdLinkTypeDDE = $00000006;
  wdLinkTypeDDEAuto = $00000007;
  wdLinkTypeChart = $00000008;


// Constants for enum WdWindowType
const
  wdWindowDocument = $00000000;
  wdWindowTemplate = $00000001;


// Constants for enum WdViewType
const
  wdNormalView = $00000001;
  wdOutlineView = $00000002;
  wdPrintView = $00000003;
  wdPrintPreview = $00000004;
  wdMasterView = $00000005;
  wdWebView = $00000006;
  wdReadingView = $00000007;
  wdConflictView = $00000008;


// Constants for enum WdSeekView
const
  wdSeekMainDocument = $00000000;
  wdSeekPrimaryHeader = $00000001;
  wdSeekFirstPageHeader = $00000002;
  wdSeekEvenPagesHeader = $00000003;
  wdSeekPrimaryFooter = $00000004;
  wdSeekFirstPageFooter = $00000005;
  wdSeekEvenPagesFooter = $00000006;
  wdSeekFootnotes = $00000007;
  wdSeekEndnotes = $00000008;
  wdSeekCurrentPageHeader = $00000009;
  wdSeekCurrentPageFooter = $0000000A;


// Constants for enum WdSpecialPane
const
  wdPaneNone = $00000000;
  wdPanePrimaryHeader = $00000001;
  wdPaneFirstPageHeader = $00000002;
  wdPaneEvenPagesHeader = $00000003;
  wdPanePrimaryFooter = $00000004;
  wdPaneFirstPageFooter = $00000005;
  wdPaneEvenPagesFooter = $00000006;
  wdPaneFootnotes = $00000007;
  wdPaneEndnotes = $00000008;
  wdPaneFootnoteContinuationNotice = $00000009;
  wdPaneFootnoteContinuationSeparator = $0000000A;
  wdPaneFootnoteSeparator = $0000000B;
  wdPaneEndnoteContinuationNotice = $0000000C;
  wdPaneEndnoteContinuationSeparator = $0000000D;
  wdPaneEndnoteSeparator = $0000000E;
  wdPaneComments = $0000000F;
  wdPaneCurrentPageHeader = $00000010;
  wdPaneCurrentPageFooter = $00000011;
  wdPaneRevisions = $00000012;
  wdPaneRevisionsHoriz = $00000013;
  wdPaneRevisionsVert = $00000014;


// Constants for enum WdPageFit
const
  wdPageFitNone = $00000000;
  wdPageFitFullPage = $00000001;
  wdPageFitBestFit = $00000002;
  wdPageFitTextFit = $00000003;


// Constants for enum WdBrowseTarget
const
  wdBrowsePage = $00000001;
  wdBrowseSection = $00000002;
  wdBrowseComment = $00000003;
  wdBrowseFootnote = $00000004;
  wdBrowseEndnote = $00000005;
  wdBrowseField = $00000006;
  wdBrowseTable = $00000007;
  wdBrowseGraphic = $00000008;
  wdBrowseHeading = $00000009;
  wdBrowseEdit = $0000000A;
  wdBrowseFind = $0000000B;
  wdBrowseGoTo = $0000000C;


// Constants for enum WdPaperTray
const
  wdPrinterDefaultBin = $00000000;
  wdPrinterUpperBin = $00000001;
  wdPrinterOnlyBin = $00000001;
  wdPrinterLowerBin = $00000002;
  wdPrinterMiddleBin = $00000003;
  wdPrinterManualFeed = $00000004;
  wdPrinterEnvelopeFeed = $00000005;
  wdPrinterManualEnvelopeFeed = $00000006;
  wdPrinterAutomaticSheetFeed = $00000007;
  wdPrinterTractorFeed = $00000008;
  wdPrinterSmallFormatBin = $00000009;
  wdPrinterLargeFormatBin = $0000000A;
  wdPrinterLargeCapacityBin = $0000000B;
  wdPrinterPaperCassette = $0000000E;
  wdPrinterFormSource = $0000000F;


// Constants for enum WdOrientation
const
  wdOrientPortrait = $00000000;
  wdOrientLandscape = $00000001;


// Constants for enum WdSelectionType
const
  wdNoSelection = $00000000;
  wdSelectionIP = $00000001;
  wdSelectionNormal = $00000002;
  wdSelectionFrame = $00000003;
  wdSelectionColumn = $00000004;
  wdSelectionRow = $00000005;
  wdSelectionBlock = $00000006;
  wdSelectionInlineShape = $00000007;
  wdSelectionShape = $00000008;


// Constants for enum WdCaptionLabelID
const
  wdCaptionFigure = $FFFFFFFF;
  wdCaptionTable = $FFFFFFFE;
  wdCaptionEquation = $FFFFFFFD;


// Constants for enum WdReferenceType
const
  wdRefTypeNumberedItem = $00000000;
  wdRefTypeHeading = $00000001;
  wdRefTypeBookmark = $00000002;
  wdRefTypeFootnote = $00000003;
  wdRefTypeEndnote = $00000004;


// Constants for enum WdReferenceKind
const
  wdContentText = $FFFFFFFF;
  wdNumberRelativeContext = $FFFFFFFE;
  wdNumberNoContext = $FFFFFFFD;
  wdNumberFullContext = $FFFFFFFC;
  wdEntireCaption = $00000002;
  wdOnlyLabelAndNumber = $00000003;
  wdOnlyCaptionText = $00000004;
  wdFootnoteNumber = $00000005;
  wdEndnoteNumber = $00000006;
  wdPageNumber = $00000007;
  wdPosition = $0000000F;
  wdFootnoteNumberFormatted = $00000010;
  wdEndnoteNumberFormatted = $00000011;


// Constants for enum WdIndexFormat
const
  wdIndexTemplate = $00000000;
  wdIndexClassic = $00000001;
  wdIndexFancy = $00000002;
  wdIndexModern = $00000003;
  wdIndexBulleted = $00000004;
  wdIndexFormal = $00000005;
  wdIndexSimple = $00000006;


// Constants for enum WdIndexType
const
  wdIndexIndent = $00000000;
  wdIndexRunin = $00000001;


// Constants for enum WdRevisionsWrap
const
  wdWrapNever = $00000000;
  wdWrapAlways = $00000001;
  wdWrapAsk = $00000002;


// Constants for enum WdRevisionType
const
  wdNoRevision = $00000000;
  wdRevisionInsert = $00000001;
  wdRevisionDelete = $00000002;
  wdRevisionProperty = $00000003;
  wdRevisionParagraphNumber = $00000004;
  wdRevisionDisplayField = $00000005;
  wdRevisionReconcile = $00000006;
  wdRevisionConflict = $00000007;
  wdRevisionStyle = $00000008;
  wdRevisionReplace = $00000009;
  wdRevisionParagraphProperty = $0000000A;
  wdRevisionTableProperty = $0000000B;
  wdRevisionSectionProperty = $0000000C;
  wdRevisionStyleDefinition = $0000000D;
  wdRevisionMovedFrom = $0000000E;
  wdRevisionMovedTo = $0000000F;
  wdRevisionCellInsertion = $00000010;
  wdRevisionCellDeletion = $00000011;
  wdRevisionCellMerge = $00000012;
  wdRevisionCellSplit = $00000013;
  wdRevisionConflictInsert = $00000014;
  wdRevisionConflictDelete = $00000015;


// Constants for enum WdRoutingSlipDelivery
const
  wdOneAfterAnother = $00000000;
  wdAllAtOnce = $00000001;


// Constants for enum WdRoutingSlipStatus
const
  wdNotYetRouted = $00000000;
  wdRouteInProgress = $00000001;
  wdRouteComplete = $00000002;


// Constants for enum WdSectionStart
const
  wdSectionContinuous = $00000000;
  wdSectionNewColumn = $00000001;
  wdSectionNewPage = $00000002;
  wdSectionEvenPage = $00000003;
  wdSectionOddPage = $00000004;


// Constants for enum WdSaveOptions
const
  wdDoNotSaveChanges = $00000000;
  wdSaveChanges = $FFFFFFFF;
  wdPromptToSaveChanges = $FFFFFFFE;


// Constants for enum WdDocumentKind
const
  wdDocumentNotSpecified = $00000000;
  wdDocumentLetter = $00000001;
  wdDocumentEmail = $00000002;


// Constants for enum WdDocumentType
const
  wdTypeDocument = $00000000;
  wdTypeTemplate = $00000001;
  wdTypeFrameset = $00000002;


// Constants for enum WdOriginalFormat
const
  wdWordDocument = $00000000;
  wdOriginalDocumentFormat = $00000001;
  wdPromptUser = $00000002;


// Constants for enum WdRelocate
const
  wdRelocateUp = $00000000;
  wdRelocateDown = $00000001;


// Constants for enum WdInsertedTextMark
const
  wdInsertedTextMarkNone = $00000000;
  wdInsertedTextMarkBold = $00000001;
  wdInsertedTextMarkItalic = $00000002;
  wdInsertedTextMarkUnderline = $00000003;
  wdInsertedTextMarkDoubleUnderline = $00000004;
  wdInsertedTextMarkColorOnly = $00000005;
  wdInsertedTextMarkStrikeThrough = $00000006;
  wdInsertedTextMarkDoubleStrikeThrough = $00000007;


// Constants for enum WdRevisedLinesMark
const
  wdRevisedLinesMarkNone = $00000000;
  wdRevisedLinesMarkLeftBorder = $00000001;
  wdRevisedLinesMarkRightBorder = $00000002;
  wdRevisedLinesMarkOutsideBorder = $00000003;


// Constants for enum WdDeletedTextMark
const
  wdDeletedTextMarkHidden = $00000000;
  wdDeletedTextMarkStrikeThrough = $00000001;
  wdDeletedTextMarkCaret = $00000002;
  wdDeletedTextMarkPound = $00000003;
  wdDeletedTextMarkNone = $00000004;
  wdDeletedTextMarkBold = $00000005;
  wdDeletedTextMarkItalic = $00000006;
  wdDeletedTextMarkUnderline = $00000007;
  wdDeletedTextMarkDoubleUnderline = $00000008;
  wdDeletedTextMarkColorOnly = $00000009;
  wdDeletedTextMarkDoubleStrikeThrough = $0000000A;


// Constants for enum WdRevisedPropertiesMark
const
  wdRevisedPropertiesMarkNone = $00000000;
  wdRevisedPropertiesMarkBold = $00000001;
  wdRevisedPropertiesMarkItalic = $00000002;
  wdRevisedPropertiesMarkUnderline = $00000003;
  wdRevisedPropertiesMarkDoubleUnderline = $00000004;
  wdRevisedPropertiesMarkColorOnly = $00000005;
  wdRevisedPropertiesMarkStrikeThrough = $00000006;
  wdRevisedPropertiesMarkDoubleStrikeThrough = $00000007;


// Constants for enum WdFieldShading
const
  wdFieldShadingNever = $00000000;
  wdFieldShadingAlways = $00000001;
  wdFieldShadingWhenSelected = $00000002;


// Constants for enum WdDefaultFilePath
const
  wdDocumentsPath = $00000000;
  wdPicturesPath = $00000001;
  wdUserTemplatesPath = $00000002;
  wdWorkgroupTemplatesPath = $00000003;
  wdUserOptionsPath = $00000004;
  wdAutoRecoverPath = $00000005;
  wdToolsPath = $00000006;
  wdTutorialPath = $00000007;
  wdStartupPath = $00000008;
  wdProgramPath = $00000009;
  wdGraphicsFiltersPath = $0000000A;
  wdTextConvertersPath = $0000000B;
  wdProofingToolsPath = $0000000C;
  wdTempFilePath = $0000000D;
  wdCurrentFolderPath = $0000000E;
  wdStyleGalleryPath = $0000000F;
  wdBorderArtPath = $00000013;


// Constants for enum WdCompatibility
const
  wdNoTabHangIndent = $00000001;
  wdNoSpaceRaiseLower = $00000002;
  wdPrintColBlack = $00000003;
  wdWrapTrailSpaces = $00000004;
  wdNoColumnBalance = $00000005;
  wdConvMailMergeEsc = $00000006;
  wdSuppressSpBfAfterPgBrk = $00000007;
  wdSuppressTopSpacing = $00000008;
  wdOrigWordTableRules = $00000009;
  wdTransparentMetafiles = $0000000A;
  wdShowBreaksInFrames = $0000000B;
  wdSwapBordersFacingPages = $0000000C;
  wdLeaveBackslashAlone = $0000000D;
  wdExpandShiftReturn = $0000000E;
  wdDontULTrailSpace = $0000000F;
  wdDontBalanceSingleByteDoubleByteWidth = $00000010;
  wdSuppressTopSpacingMac5 = $00000011;
  wdSpacingInWholePoints = $00000012;
  wdPrintBodyTextBeforeHeader = $00000013;
  wdNoLeading = $00000014;
  wdNoSpaceForUL = $00000015;
  wdMWSmallCaps = $00000016;
  wdNoExtraLineSpacing = $00000017;
  wdTruncateFontHeight = $00000018;
  wdSubFontBySize = $00000019;
  wdUsePrinterMetrics = $0000001A;
  wdWW6BorderRules = $0000001B;
  wdExactOnTop = $0000001C;
  wdSuppressBottomSpacing = $0000001D;
  wdWPSpaceWidth = $0000001E;
  wdWPJustification = $0000001F;
  wdLineWrapLikeWord6 = $00000020;
  wdShapeLayoutLikeWW8 = $00000021;
  wdFootnoteLayoutLikeWW8 = $00000022;
  wdDontUseHTMLParagraphAutoSpacing = $00000023;
  wdDontAdjustLineHeightInTable = $00000024;
  wdForgetLastTabAlignment = $00000025;
  wdAutospaceLikeWW7 = $00000026;
  wdAlignTablesRowByRow = $00000027;
  wdLayoutRawTableWidth = $00000028;
  wdLayoutTableRowsApart = $00000029;
  wdUseWord97LineBreakingRules = $0000002A;
  wdDontBreakWrappedTables = $0000002B;
  wdDontSnapTextToGridInTableWithObjects = $0000002C;
  wdSelectFieldWithFirstOrLastCharacter = $0000002D;
  wdApplyBreakingRules = $0000002E;
  wdDontWrapTextWithPunctuation = $0000002F;
  wdDontUseAsianBreakRulesInGrid = $00000030;
  wdUseWord2002TableStyleRules = $00000031;
  wdGrowAutofit = $00000032;
  wdUseNormalStyleForList = $00000033;
  wdDontUseIndentAsNumberingTabStop = $00000034;
  wdFELineBreak11 = $00000035;
  wdAllowSpaceOfSameStyleInTable = $00000036;
  wdWW11IndentRules = $00000037;
  wdDontAutofitConstrainedTables = $00000038;
  wdAutofitLikeWW11 = $00000039;
  wdUnderlineTabInNumList = $0000003A;
  wdHangulWidthLikeWW11 = $0000003B;
  wdSplitPgBreakAndParaMark = $0000003C;
  wdDontVertAlignCellWithShape = $0000003D;
  wdDontBreakConstrainedForcedTables = $0000003E;
  wdDontVertAlignInTextbox = $0000003F;
  wdWord11KerningPairs = $00000040;
  wdCachedColBalance = $00000041;
  wdDisableOTKerning = $00000042;
  wdFlipMirrorIndents = $00000043;
  wdDontOverrideTableStyleFontSzAndJustification = $00000044;
  wdUseWord2010TableStyleRules = $00000045;
  wdDelayableFloatingTable = $00000046;
  wdAllowHyphenationAtTrackBottom = $00000047;
  wdUseWord2013TrackBottomHyphenation = $00000048;
  wdUsePre2018iOSMacLayout = $00000049;


// Constants for enum WdPaperSize
const
  wdPaper10x14 = $00000000;
  wdPaper11x17 = $00000001;
  wdPaperLetter = $00000002;
  wdPaperLetterSmall = $00000003;
  wdPaperLegal = $00000004;
  wdPaperExecutive = $00000005;
  wdPaperA3 = $00000006;
  wdPaperA4 = $00000007;
  wdPaperA4Small = $00000008;
  wdPaperA5 = $00000009;
  wdPaperB4 = $0000000A;
  wdPaperB5 = $0000000B;
  wdPaperCSheet = $0000000C;
  wdPaperDSheet = $0000000D;
  wdPaperESheet = $0000000E;
  wdPaperFanfoldLegalGerman = $0000000F;
  wdPaperFanfoldStdGerman = $00000010;
  wdPaperFanfoldUS = $00000011;
  wdPaperFolio = $00000012;
  wdPaperLedger = $00000013;
  wdPaperNote = $00000014;
  wdPaperQuarto = $00000015;
  wdPaperStatement = $00000016;
  wdPaperTabloid = $00000017;
  wdPaperEnvelope9 = $00000018;
  wdPaperEnvelope10 = $00000019;
  wdPaperEnvelope11 = $0000001A;
  wdPaperEnvelope12 = $0000001B;
  wdPaperEnvelope14 = $0000001C;
  wdPaperEnvelopeB4 = $0000001D;
  wdPaperEnvelopeB5 = $0000001E;
  wdPaperEnvelopeB6 = $0000001F;
  wdPaperEnvelopeC3 = $00000020;
  wdPaperEnvelopeC4 = $00000021;
  wdPaperEnvelopeC5 = $00000022;
  wdPaperEnvelopeC6 = $00000023;
  wdPaperEnvelopeC65 = $00000024;
  wdPaperEnvelopeDL = $00000025;
  wdPaperEnvelopeItaly = $00000026;
  wdPaperEnvelopeMonarch = $00000027;
  wdPaperEnvelopePersonal = $00000028;
  wdPaperCustom = $00000029;


// Constants for enum WdCustomLabelPageSize
const
  wdCustomLabelLetter = $00000000;
  wdCustomLabelLetterLS = $00000001;
  wdCustomLabelA4 = $00000002;
  wdCustomLabelA4LS = $00000003;
  wdCustomLabelA5 = $00000004;
  wdCustomLabelA5LS = $00000005;
  wdCustomLabelB5 = $00000006;
  wdCustomLabelMini = $00000007;
  wdCustomLabelFanfold = $00000008;
  wdCustomLabelVertHalfSheet = $00000009;
  wdCustomLabelVertHalfSheetLS = $0000000A;
  wdCustomLabelHigaki = $0000000B;
  wdCustomLabelHigakiLS = $0000000C;
  wdCustomLabelB4JIS = $0000000D;


// Constants for enum WdProtectionType
const
  wdNoProtection = $FFFFFFFF;
  wdAllowOnlyRevisions = $00000000;
  wdAllowOnlyComments = $00000001;
  wdAllowOnlyFormFields = $00000002;
  wdAllowOnlyReading = $00000003;


// Constants for enum WdPartOfSpeech
const
  wdAdjective = $00000000;
  wdNoun = $00000001;
  wdAdverb = $00000002;
  wdVerb = $00000003;
  wdPronoun = $00000004;
  wdConjunction = $00000005;
  wdPreposition = $00000006;
  wdInterjection = $00000007;
  wdIdiom = $00000008;
  wdOther = $00000009;


// Constants for enum WdSubscriberFormats
const
  wdSubscriberBestFormat = $00000000;
  wdSubscriberRTF = $00000001;
  wdSubscriberText = $00000002;
  wdSubscriberPict = $00000004;


// Constants for enum WdEditionType
const
  wdPublisher = $00000000;
  wdSubscriber = $00000001;


// Constants for enum WdEditionOption
const
  wdCancelPublisher = $00000000;
  wdSendPublisher = $00000001;
  wdSelectPublisher = $00000002;
  wdAutomaticUpdate = $00000003;
  wdManualUpdate = $00000004;
  wdChangeAttributes = $00000005;
  wdUpdateSubscriber = $00000006;
  wdOpenSource = $00000007;


// Constants for enum WdRelativeHorizontalPosition
const
  wdRelativeHorizontalPositionMargin = $00000000;
  wdRelativeHorizontalPositionPage = $00000001;
  wdRelativeHorizontalPositionColumn = $00000002;
  wdRelativeHorizontalPositionCharacter = $00000003;
  wdRelativeHorizontalPositionLeftMarginArea = $00000004;
  wdRelativeHorizontalPositionRightMarginArea = $00000005;
  wdRelativeHorizontalPositionInnerMarginArea = $00000006;
  wdRelativeHorizontalPositionOuterMarginArea = $00000007;


// Constants for enum WdRelativeVerticalPosition
const
  wdRelativeVerticalPositionMargin = $00000000;
  wdRelativeVerticalPositionPage = $00000001;
  wdRelativeVerticalPositionParagraph = $00000002;
  wdRelativeVerticalPositionLine = $00000003;
  wdRelativeVerticalPositionTopMarginArea = $00000004;
  wdRelativeVerticalPositionBottomMarginArea = $00000005;
  wdRelativeVerticalPositionInnerMarginArea = $00000006;
  wdRelativeVerticalPositionOuterMarginArea = $00000007;


// Constants for enum WdHelpType
const
  wdHelp = $00000000;
  wdHelpAbout = $00000001;
  wdHelpActiveWindow = $00000002;
  wdHelpContents = $00000003;
  wdHelpExamplesAndDemos = $00000004;
  wdHelpIndex = $00000005;
  wdHelpKeyboard = $00000006;
  wdHelpPSSHelp = $00000007;
  wdHelpQuickPreview = $00000008;
  wdHelpSearch = $00000009;
  wdHelpUsingHelp = $0000000A;
  wdHelpIchitaro = $0000000B;
  wdHelpPE2 = $0000000C;
  wdHelpHWP = $0000000D;


// Constants for enum WdHelpTypeHID
const
  emptyenum____________ = $00000000;


// Constants for enum WdKeyCategory
const
  wdKeyCategoryNil = $FFFFFFFF;
  wdKeyCategoryDisable = $00000000;
  wdKeyCategoryCommand = $00000001;
  wdKeyCategoryMacro = $00000002;
  wdKeyCategoryFont = $00000003;
  wdKeyCategoryAutoText = $00000004;
  wdKeyCategoryStyle = $00000005;
  wdKeyCategorySymbol = $00000006;
  wdKeyCategoryPrefix = $00000007;


// Constants for enum WdKey
const
  wdNoKey = $000000FF;
  wdKeyShift = $00000100;
  wdKeyControl = $00000200;
  wdKeyCommand = $00000200;
  wdKeyAlt = $00000400;
  wdKeyOption = $00000400;
  wdKeyA = $00000041;
  wdKeyB = $00000042;
  wdKeyC = $00000043;
  wdKeyD = $00000044;
  wdKeyE = $00000045;
  wdKeyF = $00000046;
  wdKeyG = $00000047;
  wdKeyH = $00000048;
  wdKeyI = $00000049;
  wdKeyJ = $0000004A;
  wdKeyK = $0000004B;
  wdKeyL = $0000004C;
  wdKeyM = $0000004D;
  wdKeyN = $0000004E;
  wdKeyO = $0000004F;
  wdKeyP = $00000050;
  wdKeyQ = $00000051;
  wdKeyR = $00000052;
  wdKeyS = $00000053;
  wdKeyT = $00000054;
  wdKeyU = $00000055;
  wdKeyV = $00000056;
  wdKeyW = $00000057;
  wdKeyX = $00000058;
  wdKeyY = $00000059;
  wdKeyZ = $0000005A;
  wdKey0 = $00000030;
  wdKey1 = $00000031;
  wdKey2 = $00000032;
  wdKey3 = $00000033;
  wdKey4 = $00000034;
  wdKey5 = $00000035;
  wdKey6 = $00000036;
  wdKey7 = $00000037;
  wdKey8 = $00000038;
  wdKey9 = $00000039;
  wdKeyBackspace = $00000008;
  wdKeyTab = $00000009;
  wdKeyNumeric5Special = $0000000C;
  wdKeyReturn = $0000000D;
  wdKeyPause = $00000013;
  wdKeyEsc = $0000001B;
  wdKeySpacebar = $00000020;
  wdKeyPageUp = $00000021;
  wdKeyPageDown = $00000022;
  wdKeyEnd = $00000023;
  wdKeyHome = $00000024;
  wdKeyInsert = $0000002D;
  wdKeyDelete = $0000002E;
  wdKeyNumeric0 = $00000060;
  wdKeyNumeric1 = $00000061;
  wdKeyNumeric2 = $00000062;
  wdKeyNumeric3 = $00000063;
  wdKeyNumeric4 = $00000064;
  wdKeyNumeric5 = $00000065;
  wdKeyNumeric6 = $00000066;
  wdKeyNumeric7 = $00000067;
  wdKeyNumeric8 = $00000068;
  wdKeyNumeric9 = $00000069;
  wdKeyNumericMultiply = $0000006A;
  wdKeyNumericAdd = $0000006B;
  wdKeyNumericSubtract = $0000006D;
  wdKeyNumericDecimal = $0000006E;
  wdKeyNumericDivide = $0000006F;
  wdKeyF1 = $00000070;
  wdKeyF2 = $00000071;
  wdKeyF3 = $00000072;
  wdKeyF4 = $00000073;
  wdKeyF5 = $00000074;
  wdKeyF6 = $00000075;
  wdKeyF7 = $00000076;
  wdKeyF8 = $00000077;
  wdKeyF9 = $00000078;
  wdKeyF10 = $00000079;
  wdKeyF11 = $0000007A;
  wdKeyF12 = $0000007B;
  wdKeyF13 = $0000007C;
  wdKeyF14 = $0000007D;
  wdKeyF15 = $0000007E;
  wdKeyF16 = $0000007F;
  wdKeyScrollLock = $00000091;
  wdKeySemiColon = $000000BA;
  wdKeyEquals = $000000BB;
  wdKeyComma = $000000BC;
  wdKeyHyphen = $000000BD;
  wdKeyPeriod = $000000BE;
  wdKeySlash = $000000BF;
  wdKeyBackSingleQuote = $000000C0;
  wdKeyOpenSquareBrace = $000000DB;
  wdKeyBackSlash = $000000DC;
  wdKeyCloseSquareBrace = $000000DD;
  wdKeySingleQuote = $000000DE;


// Constants for enum WdOLEType
const
  wdOLELink = $00000000;
  wdOLEEmbed = $00000001;
  wdOLEControl = $00000002;


// Constants for enum WdOLEVerb
const
  wdOLEVerbPrimary = $00000000;
  wdOLEVerbShow = $FFFFFFFF;
  wdOLEVerbOpen = $FFFFFFFE;
  wdOLEVerbHide = $FFFFFFFD;
  wdOLEVerbUIActivate = $FFFFFFFC;
  wdOLEVerbInPlaceActivate = $FFFFFFFB;
  wdOLEVerbDiscardUndoState = $FFFFFFFA;


// Constants for enum WdOLEPlacement
const
  wdInLine = $00000000;
  wdFloatOverText = $00000001;


// Constants for enum WdEnvelopeOrientation
const
  wdLeftPortrait = $00000000;
  wdCenterPortrait = $00000001;
  wdRightPortrait = $00000002;
  wdLeftLandscape = $00000003;
  wdCenterLandscape = $00000004;
  wdRightLandscape = $00000005;
  wdLeftClockwise = $00000006;
  wdCenterClockwise = $00000007;
  wdRightClockwise = $00000008;


// Constants for enum WdLetterStyle
const
  wdFullBlock = $00000000;
  wdModifiedBlock = $00000001;
  wdSemiBlock = $00000002;


// Constants for enum WdLetterheadLocation
const
  wdLetterTop = $00000000;
  wdLetterBottom = $00000001;
  wdLetterLeft = $00000002;
  wdLetterRight = $00000003;


// Constants for enum WdSalutationType
const
  wdSalutationInformal = $00000000;
  wdSalutationFormal = $00000001;
  wdSalutationBusiness = $00000002;
  wdSalutationOther = $00000003;


// Constants for enum WdSalutationGender
const
  wdGenderFemale = $00000000;
  wdGenderMale = $00000001;
  wdGenderNeutral = $00000002;
  wdGenderUnknown = $00000003;


// Constants for enum WdMovementType
const
  wdMove = $00000000;
  wdExtend = $00000001;


// Constants for enum WdConstants
const
  wdUndefined = $0098967F;
  wdToggle = $0098967E;
  wdForward = $3FFFFFFF;
  wdBackward = $C0000001;
  wdAutoPosition = $00000000;
  wdFirst = $00000001;
  wdCreatorCode = $4D535744;


// Constants for enum WdPasteDataType
const
  wdPasteOLEObject = $00000000;
  wdPasteRTF = $00000001;
  wdPasteText = $00000002;
  wdPasteMetafilePicture = $00000003;
  wdPasteBitmap = $00000004;
  wdPasteDeviceIndependentBitmap = $00000005;
  wdPasteHyperlink = $00000007;
  wdPasteShape = $00000008;
  wdPasteEnhancedMetafile = $00000009;
  wdPasteHTML = $0000000A;


// Constants for enum WdPrintOutItem
const
  wdPrintDocumentContent = $00000000;
  wdPrintProperties = $00000001;
  wdPrintComments = $00000002;
  wdPrintMarkup = $00000002;
  wdPrintStyles = $00000003;
  wdPrintAutoTextEntries = $00000004;
  wdPrintKeyAssignments = $00000005;
  wdPrintEnvelope = $00000006;
  wdPrintDocumentWithMarkup = $00000007;


// Constants for enum WdPrintOutPages
const
  wdPrintAllPages = $00000000;
  wdPrintOddPagesOnly = $00000001;
  wdPrintEvenPagesOnly = $00000002;


// Constants for enum WdPrintOutRange
const
  wdPrintAllDocument = $00000000;
  wdPrintSelection = $00000001;
  wdPrintCurrentPage = $00000002;
  wdPrintFromTo = $00000003;
  wdPrintRangeOfPages = $00000004;


// Constants for enum WdDictionaryType
const
  wdSpelling = $00000000;
  wdGrammar = $00000001;
  wdThesaurus = $00000002;
  wdHyphenation = $00000003;
  wdSpellingComplete = $00000004;
  wdSpellingCustom = $00000005;
  wdSpellingLegal = $00000006;
  wdSpellingMedical = $00000007;
  wdHangulHanjaConversion = $00000008;
  wdHangulHanjaConversionCustom = $00000009;


// Constants for enum WdDictionaryTypeHID
const
  emptyenum_____________ = $00000000;


// Constants for enum WdSpellingWordType
const
  wdSpellword = $00000000;
  wdWildcard = $00000001;
  wdAnagram = $00000002;


// Constants for enum WdSpellingErrorType
const
  wdSpellingCorrect = $00000000;
  wdSpellingNotInDictionary = $00000001;
  wdSpellingCapitalization = $00000002;


// Constants for enum WdProofreadingErrorType
const
  wdSpellingError = $00000000;
  wdGrammaticalError = $00000001;


// Constants for enum WdInlineShapeType
const
  wdInlineShapeEmbeddedOLEObject = $00000001;
  wdInlineShapeLinkedOLEObject = $00000002;
  wdInlineShapePicture = $00000003;
  wdInlineShapeLinkedPicture = $00000004;
  wdInlineShapeOLEControlObject = $00000005;
  wdInlineShapeHorizontalLine = $00000006;
  wdInlineShapePictureHorizontalLine = $00000007;
  wdInlineShapeLinkedPictureHorizontalLine = $00000008;
  wdInlineShapePictureBullet = $00000009;
  wdInlineShapeScriptAnchor = $0000000A;
  wdInlineShapeOWSAnchor = $0000000B;
  wdInlineShapeChart = $0000000C;
  wdInlineShapeDiagram = $0000000D;
  wdInlineShapeLockedCanvas = $0000000E;
  wdInlineShapeSmartArt = $0000000F;
  wdInlineShapeWebVideo = $00000010;
  wdInlineShapeGraphic = $00000011;
  wdInlineShapeLinkedGraphic = $00000012;
  wdInlineShape3DModel = $00000013;
  wdInlineShapeLinked3DModel = $00000014;


// Constants for enum WdArrangeStyle
const
  wdTiled = $00000000;
  wdIcons = $00000001;


// Constants for enum WdSelectionFlags
const
  wdSelStartActive = $00000001;



// Constants for enum WdAutoVersions
const
  wdAutoVersionOff = $00000000;
  wdAutoVersionOnClose = $00000001;


// Constants for enum WdOrganizerObject
const
  wdOrganizerObjectStyles = $00000000;
  wdOrganizerObjectAutoText = $00000001;
  wdOrganizerObjectCommandBars = $00000002;
  wdOrganizerObjectProjectItems = $00000003;


// Constants for enum WdFindMatch
const
  wdMatchParagraphMark = $0001000F;
  wdMatchTabCharacter = $00000009;
  wdMatchCommentMark = $00000005;
  wdMatchAnyCharacter = $0001003F;
  wdMatchAnyDigit = $0001001F;
  wdMatchAnyLetter = $0001002F;
  wdMatchCaretCharacter = $0000000B;
  wdMatchColumnBreak = $0000000E;
  wdMatchEmDash = $00002014;
  wdMatchEnDash = $00002013;
  wdMatchEndnoteMark = $00010013;
  wdMatchField = $00000013;
  wdMatchFootnoteMark = $00010012;
  wdMatchGraphic = $00000001;
  wdMatchManualLineBreak = $0001000F;
  wdMatchManualPageBreak = $0001001C;
  wdMatchNonbreakingHyphen = $0000001E;
  wdMatchNonbreakingSpace = $000000A0;
  wdMatchOptionalHyphen = $0000001F;
  wdMatchSectionBreak = $0001002C;
  wdMatchWhiteSpace = $00010077;


// Constants for enum WdFindWrap
const
  wdFindStop = $00000000;
  wdFindContinue = $00000001;
  wdFindAsk = $00000002;


// Constants for enum WdInformation
const
  wdActiveEndAdjustedPageNumber = $00000001;
  wdActiveEndSectionNumber = $00000002;
  wdActiveEndPageNumber = $00000003;
  wdNumberOfPagesInDocument = $00000004;
  wdHorizontalPositionRelativeToPage = $00000005;
  wdVerticalPositionRelativeToPage = $00000006;
  wdHorizontalPositionRelativeToTextBoundary = $00000007;
  wdVerticalPositionRelativeToTextBoundary = $00000008;
  wdFirstCharacterColumnNumber = $00000009;
  wdFirstCharacterLineNumber = $0000000A;
  wdFrameIsSelected = $0000000B;
  wdWithInTable = $0000000C;
  wdStartOfRangeRowNumber = $0000000D;
  wdEndOfRangeRowNumber = $0000000E;
  wdMaximumNumberOfRows = $0000000F;
  wdStartOfRangeColumnNumber = $00000010;
  wdEndOfRangeColumnNumber = $00000011;
  wdMaximumNumberOfColumns = $00000012;
  wdZoomPercentage = $00000013;
  wdSelectionMode = $00000014;
  wdCapsLock = $00000015;
  wdNumLock = $00000016;
  wdOverType = $00000017;
  wdRevisionMarking = $00000018;
  wdInFootnoteEndnotePane = $00000019;
  wdInCommentPane = $0000001A;
  wdInHeaderFooter = $0000001C;
  wdAtEndOfRowMarker = $0000001F;
  wdReferenceOfType = $00000020;
  wdHeaderFooterType = $00000021;
  wdInMasterDocument = $00000022;
  wdInFootnote = $00000023;
  wdInEndnote = $00000024;
  wdInWordMail = $00000025;
  wdInClipboard = $00000026;
  wdInCoverPage = $00000029;
  wdInBibliography = $0000002A;
  wdInCitation = $0000002B;
  wdInFieldCode = $0000002C;
  wdInFieldResult = $0000002D;
  wdInContentControl = $0000002E;


// Constants for enum WdWrapType
const
  wdWrapSquare = $00000000;
  wdWrapTight = $00000001;
  wdWrapThrough = $00000002;
  wdWrapNone = $00000003;
  wdWrapTopBottom = $00000004;
  wdWrapBehind = $00000005;
  wdWrapFront = $00000003;
  wdWrapInline = $00000007;


// Constants for enum WdWrapSideType
const
  wdWrapBoth = $00000000;
  wdWrapLeft = $00000001;
  wdWrapRight = $00000002;
  wdWrapLargest = $00000003;


// Constants for enum WdOutlineLevel
const
  wdOutlineLevel1 = $00000001;
  wdOutlineLevel2 = $00000002;
  wdOutlineLevel3 = $00000003;
  wdOutlineLevel4 = $00000004;
  wdOutlineLevel5 = $00000005;
  wdOutlineLevel6 = $00000006;
  wdOutlineLevel7 = $00000007;
  wdOutlineLevel8 = $00000008;
  wdOutlineLevel9 = $00000009;
  wdOutlineLevelBodyText = $0000000A;


// Constants for enum WdTextOrientation
const
  wdTextOrientationHorizontal = $00000000;
  wdTextOrientationUpward = $00000002;
  wdTextOrientationDownward = $00000003;
  wdTextOrientationVerticalFarEast = $00000001;
  wdTextOrientationHorizontalRotatedFarEast = $00000004;
  wdTextOrientationVertical = $00000005;


// Constants for enum WdTextOrientationHID
const
  emptyenum______________ = $00000000;


// Constants for enum WdPageBorderArt
const
  wdArtApples = $00000001;
  wdArtMapleMuffins = $00000002;
  wdArtCakeSlice = $00000003;
  wdArtCandyCorn = $00000004;
  wdArtIceCreamCones = $00000005;
  wdArtChampagneBottle = $00000006;
  wdArtPartyGlass = $00000007;
  wdArtChristmasTree = $00000008;
  wdArtTrees = $00000009;
  wdArtPalmsColor = $0000000A;
  wdArtBalloons3Colors = $0000000B;
  wdArtBalloonsHotAir = $0000000C;
  wdArtPartyFavor = $0000000D;
  wdArtConfettiStreamers = $0000000E;
  wdArtHearts = $0000000F;
  wdArtHeartBalloon = $00000010;
  wdArtStars3D = $00000011;
  wdArtStarsShadowed = $00000012;
  wdArtStars = $00000013;
  wdArtSun = $00000014;
  wdArtEarth2 = $00000015;
  wdArtEarth1 = $00000016;
  wdArtPeopleHats = $00000017;
  wdArtSombrero = $00000018;
  wdArtPencils = $00000019;
  wdArtPackages = $0000001A;
  wdArtClocks = $0000001B;
  wdArtFirecrackers = $0000001C;
  wdArtRings = $0000001D;
  wdArtMapPins = $0000001E;
  wdArtConfetti = $0000001F;
  wdArtCreaturesButterfly = $00000020;
  wdArtCreaturesLadyBug = $00000021;
  wdArtCreaturesFish = $00000022;
  wdArtBirdsFlight = $00000023;
  wdArtScaredCat = $00000024;
  wdArtBats = $00000025;
  wdArtFlowersRoses = $00000026;
  wdArtFlowersRedRose = $00000027;
  wdArtPoinsettias = $00000028;
  wdArtHolly = $00000029;
  wdArtFlowersTiny = $0000002A;
  wdArtFlowersPansy = $0000002B;
  wdArtFlowersModern2 = $0000002C;
  wdArtFlowersModern1 = $0000002D;
  wdArtWhiteFlowers = $0000002E;
  wdArtVine = $0000002F;
  wdArtFlowersDaisies = $00000030;
  wdArtFlowersBlockPrint = $00000031;
  wdArtDecoArchColor = $00000032;
  wdArtFans = $00000033;
  wdArtFilm = $00000034;
  wdArtLightning1 = $00000035;
  wdArtCompass = $00000036;
  wdArtDoubleD = $00000037;
  wdArtClassicalWave = $00000038;
  wdArtShadowedSquares = $00000039;
  wdArtTwistedLines1 = $0000003A;
  wdArtWaveline = $0000003B;
  wdArtQuadrants = $0000003C;
  wdArtCheckedBarColor = $0000003D;
  wdArtSwirligig = $0000003E;
  wdArtPushPinNote1 = $0000003F;
  wdArtPushPinNote2 = $00000040;
  wdArtPumpkin1 = $00000041;
  wdArtEggsBlack = $00000042;
  wdArtCup = $00000043;
  wdArtHeartGray = $00000044;
  wdArtGingerbreadMan = $00000045;
  wdArtBabyPacifier = $00000046;
  wdArtBabyRattle = $00000047;
  wdArtCabins = $00000048;
  wdArtHouseFunky = $00000049;
  wdArtStarsBlack = $0000004A;
  wdArtSnowflakes = $0000004B;
  wdArtSnowflakeFancy = $0000004C;
  wdArtSkyrocket = $0000004D;
  wdArtSeattle = $0000004E;
  wdArtMusicNotes = $0000004F;
  wdArtPalmsBlack = $00000050;
  wdArtMapleLeaf = $00000051;
  wdArtPaperClips = $00000052;
  wdArtShorebirdTracks = $00000053;
  wdArtPeople = $00000054;
  wdArtPeopleWaving = $00000055;
  wdArtEclipsingSquares2 = $00000056;
  wdArtHypnotic = $00000057;
  wdArtDiamondsGray = $00000058;
  wdArtDecoArch = $00000059;
  wdArtDecoBlocks = $0000005A;
  wdArtCirclesLines = $0000005B;
  wdArtPapyrus = $0000005C;
  wdArtWoodwork = $0000005D;
  wdArtWeavingBraid = $0000005E;
  wdArtWeavingRibbon = $0000005F;
  wdArtWeavingAngles = $00000060;
  wdArtArchedScallops = $00000061;
  wdArtSafari = $00000062;
  wdArtCelticKnotwork = $00000063;
  wdArtCrazyMaze = $00000064;
  wdArtEclipsingSquares1 = $00000065;
  wdArtBirds = $00000066;
  wdArtFlowersTeacup = $00000067;
  wdArtNorthwest = $00000068;
  wdArtSouthwest = $00000069;
  wdArtTribal6 = $0000006A;
  wdArtTribal4 = $0000006B;
  wdArtTribal3 = $0000006C;
  wdArtTribal2 = $0000006D;
  wdArtTribal5 = $0000006E;
  wdArtXIllusions = $0000006F;
  wdArtZanyTriangles = $00000070;
  wdArtPyramids = $00000071;
  wdArtPyramidsAbove = $00000072;
  wdArtConfettiGrays = $00000073;
  wdArtConfettiOutline = $00000074;
  wdArtConfettiWhite = $00000075;
  wdArtMosaic = $00000076;
  wdArtLightning2 = $00000077;
  wdArtHeebieJeebies = $00000078;
  wdArtLightBulb = $00000079;
  wdArtGradient = $0000007A;
  wdArtTriangleParty = $0000007B;
  wdArtTwistedLines2 = $0000007C;
  wdArtMoons = $0000007D;
  wdArtOvals = $0000007E;
  wdArtDoubleDiamonds = $0000007F;
  wdArtChainLink = $00000080;
  wdArtTriangles = $00000081;
  wdArtTribal1 = $00000082;
  wdArtMarqueeToothed = $00000083;
  wdArtSharksTeeth = $00000084;
  wdArtSawtooth = $00000085;
  wdArtSawtoothGray = $00000086;
  wdArtPostageStamp = $00000087;
  wdArtWeavingStrips = $00000088;
  wdArtZigZag = $00000089;
  wdArtCrossStitch = $0000008A;
  wdArtGems = $0000008B;
  wdArtCirclesRectangles = $0000008C;
  wdArtCornerTriangles = $0000008D;
  wdArtCreaturesInsects = $0000008E;
  wdArtZigZagStitch = $0000008F;
  wdArtCheckered = $00000090;
  wdArtCheckedBarBlack = $00000091;
  wdArtMarquee = $00000092;
  wdArtBasicWhiteDots = $00000093;
  wdArtBasicWideMidline = $00000094;
  wdArtBasicWideOutline = $00000095;
  wdArtBasicWideInline = $00000096;
  wdArtBasicThinLines = $00000097;
  wdArtBasicWhiteDashes = $00000098;
  wdArtBasicWhiteSquares = $00000099;
  wdArtBasicBlackSquares = $0000009A;
  wdArtBasicBlackDashes = $0000009B;
  wdArtBasicBlackDots = $0000009C;
  wdArtStarsTop = $0000009D;
  wdArtCertificateBanner = $0000009E;
  wdArtHandmade1 = $0000009F;
  wdArtHandmade2 = $000000A0;
  wdArtTornPaper = $000000A1;
  wdArtTornPaperBlack = $000000A2;
  wdArtCouponCutoutDashes = $000000A3;
  wdArtCouponCutoutDots = $000000A4;


// Constants for enum WdBorderDistanceFrom
const
  wdBorderDistanceFromText = $00000000;
  wdBorderDistanceFromPageEdge = $00000001;


// Constants for enum WdReplace
const
  wdReplaceNone = $00000000;
  wdReplaceOne = $00000001;
  wdReplaceAll = $00000002;


// Constants for enum WdFontBias
const
  wdFontBiasDontCare = $000000FF;
  wdFontBiasDefault = $00000000;
  wdFontBiasFareast = $00000001;


// Constants for enum WdBrowserLevel
const
  wdBrowserLevelV4 = $00000000;
  wdBrowserLevelMicrosoftInternetExplorer5 = $00000001;
  wdBrowserLevelMicrosoftInternetExplorer6 = $00000002;


// Constants for enum WdEnclosureType
const
  wdEnclosureCircle = $00000000;
  wdEnclosureSquare = $00000001;
  wdEnclosureTriangle = $00000002;
  wdEnclosureDiamond = $00000003;


// Constants for enum WdEncloseStyle
const
  wdEncloseStyleNone = $00000000;
  wdEncloseStyleSmall = $00000001;
  wdEncloseStyleLarge = $00000002;


// Constants for enum WdHighAnsiText
const
  wdHighAnsiIsFarEast = $00000000;
  wdHighAnsiIsHighAnsi = $00000001;
  wdAutoDetectHighAnsiFarEast = $00000002;


// Constants for enum WdLayoutMode
const
  wdLayoutModeDefault = $00000000;
  wdLayoutModeGrid = $00000001;
  wdLayoutModeLineGrid = $00000002;
  wdLayoutModeGenko = $00000003;


// Constants for enum WdDocumentMedium
const
  wdEmailMessage = $00000000;
  wdDocument = $00000001;
  wdWebPage = $00000002;


// Constants for enum WdMailerPriority
const
  wdPriorityNormal = $00000001;
  wdPriorityLow = $00000002;
  wdPriorityHigh = $00000003;


// Constants for enum WdDocumentViewDirection
const
  wdDocumentViewRtl = $00000000;
  wdDocumentViewLtr = $00000001;


// Constants for enum WdArabicNumeral
const
  wdNumeralArabic = $00000000;
  wdNumeralHindi = $00000001;
  wdNumeralContext = $00000002;
  wdNumeralSystem = $00000003;


// Constants for enum WdMonthNames
const
  wdMonthNamesArabic = $00000000;
  wdMonthNamesEnglish = $00000001;
  wdMonthNamesFrench = $00000002;


// Constants for enum WdCursorMovement
const
  wdCursorMovementLogical = $00000000;
  wdCursorMovementVisual = $00000001;


// Constants for enum WdVisualSelection
const
  wdVisualSelectionBlock = $00000000;
  wdVisualSelectionContinuous = $00000001;


// Constants for enum WdTableDirection
const
  wdTableDirectionRtl = $00000000;
  wdTableDirectionLtr = $00000001;


// Constants for enum WdFlowDirection
const
  wdFlowLtr = $00000000;
  wdFlowRtl = $00000001;


// Constants for enum WdDiacriticColor
const
  wdDiacriticColorBidi = $00000000;
  wdDiacriticColorLatin = $00000001;


// Constants for enum WdGutterStyle
const
  wdGutterPosLeft = $00000000;
  wdGutterPosTop = $00000001;
  wdGutterPosRight = $00000002;


// Constants for enum WdGutterStyleOld
const
  wdGutterStyleLatin = $FFFFFFF6;
  wdGutterStyleBidi = $00000002;


// Constants for enum WdSectionDirection
const
  wdSectionDirectionRtl = $00000000;
  wdSectionDirectionLtr = $00000001;


// Constants for enum WdDateLanguage
const
  wdDateLanguageBidi = $0000000A;
  wdDateLanguageLatin = $00000409;


// Constants for enum WdCalendarTypeBi
const
  wdCalendarTypeBidi = $00000063;
  wdCalendarTypeGregorian = $00000064;


// Constants for enum WdCalendarType
const
  wdCalendarWestern = $00000000;
  wdCalendarArabic = $00000001;
  wdCalendarHebrew = $00000002;
  wdCalendarTaiwan = $00000003;
  wdCalendarJapan = $00000004;
  wdCalendarThai = $00000005;
  wdCalendarKorean = $00000006;
  wdCalendarSakaEra = $00000007;
  wdCalendarTranslitEnglish = $00000008;
  wdCalendarTranslitFrench = $00000009;
  wdCalendarUmalqura = $0000000D;


// Constants for enum WdReadingOrder
const
  wdReadingOrderRtl = $00000000;
  wdReadingOrderLtr = $00000001;


// Constants for enum WdHebSpellStart
const
  wdFullScript = $00000000;
  wdPartialScript = $00000001;
  wdMixedScript = $00000002;
  wdMixedAuthorizedScript = $00000003;


// Constants for enum WdAraSpeller
const
  wdNone = $00000000;
  wdInitialAlef = $00000001;
  wdFinalYaa = $00000002;
  wdBoth = $00000003;


// Constants for enum WdColor
const
  wdColorAutomatic = $FF000000;
  wdColorBlack = $00000000;
  wdColorBlue = $00FF0000;
  wdColorTurquoise = $00FFFF00;
  wdColorBrightGreen = $0000FF00;
  wdColorPink = $00FF00FF;
  wdColorRed = $000000FF;
  wdColorYellow = $0000FFFF;
  wdColorWhite = $00FFFFFF;
  wdColorDarkBlue = $00800000;
  wdColorTeal = $00808000;
  wdColorGreen = $00008000;
  wdColorViolet = $00800080;
  wdColorDarkRed = $00000080;
  wdColorDarkYellow = $00008080;
  wdColorBrown = $00003399;
  wdColorOliveGreen = $00003333;
  wdColorDarkGreen = $00003300;
  wdColorDarkTeal = $00663300;
  wdColorIndigo = $00993333;
  wdColorOrange = $000066FF;
  wdColorBlueGray = $00996666;
  wdColorLightOrange = $000099FF;
  wdColorLime = $0000CC99;
  wdColorSeaGreen = $00669933;
  wdColorAqua = $00CCCC33;
  wdColorLightBlue = $00FF6633;
  wdColorGold = $0000CCFF;
  wdColorSkyBlue = $00FFCC00;
  wdColorPlum = $00663399;
  wdColorRose = $00CC99FF;
  wdColorTan = $0099CCFF;
  wdColorLightYellow = $0099FFFF;
  wdColorLightGreen = $00CCFFCC;
  wdColorLightTurquoise = $00FFFFCC;
  wdColorPaleBlue = $00FFCC99;
  wdColorLavender = $00FF99CC;
  wdColorGray05 = $00F3F3F3;
  wdColorGray10 = $00E6E6E6;
  wdColorGray125 = $00E0E0E0;
  wdColorGray15 = $00D9D9D9;
  wdColorGray20 = $00CCCCCC;
  wdColorGray25 = $00C0C0C0;
  wdColorGray30 = $00B3B3B3;
  wdColorGray35 = $00A6A6A6;
  wdColorGray375 = $00A0A0A0;
  wdColorGray40 = $00999999;
  wdColorGray45 = $008C8C8C;
  wdColorGray50 = $00808080;
  wdColorGray55 = $00737373;
  wdColorGray60 = $00666666;
  wdColorGray625 = $00606060;
  wdColorGray65 = $00595959;
  wdColorGray70 = $004C4C4C;
  wdColorGray75 = $00404040;
  wdColorGray80 = $00333333;
  wdColorGray85 = $00262626;
  wdColorGray875 = $00202020;
  wdColorGray90 = $00191919;
  wdColorGray95 = $000C0C0C;


// Constants for enum WdShapePosition
const
  wdShapeTop = $FFF0BDC1;
  wdShapeLeft = $FFF0BDC2;
  wdShapeBottom = $FFF0BDC3;
  wdShapeRight = $FFF0BDC4;
  wdShapeCenter = $FFF0BDC5;
  wdShapeInside = $FFF0BDC6;
  wdShapeOutside = $FFF0BDC7;


// Constants for enum WdTablePosition
const
  wdTableTop = $FFF0BDC1;
  wdTableLeft = $FFF0BDC2;
  wdTableBottom = $FFF0BDC3;
  wdTableRight = $FFF0BDC4;
  wdTableCenter = $FFF0BDC5;
  wdTableInside = $FFF0BDC6;
  wdTableOutside = $FFF0BDC7;


// Constants for enum WdDefaultListBehavior
const
  wdWord8ListBehavior = $00000000;
  wdWord9ListBehavior = $00000001;
  wdWord10ListBehavior = $00000002;


// Constants for enum WdDefaultTableBehavior
const
  wdWord8TableBehavior = $00000000;
  wdWord9TableBehavior = $00000001;


// Constants for enum WdAutoFitBehavior
const
  wdAutoFitFixed = $00000000;
  wdAutoFitContent = $00000001;
  wdAutoFitWindow = $00000002;


// Constants for enum WdPreferredWidthType
const
  wdPreferredWidthAuto = $00000001;
  wdPreferredWidthPercent = $00000002;
  wdPreferredWidthPoints = $00000003;


// Constants for enum WdFarEastLineBreakLanguageID
const
  wdLineBreakJapanese = $00000411;
  wdLineBreakKorean = $00000412;
  wdLineBreakSimplifiedChinese = $00000804;
  wdLineBreakTraditionalChinese = $00000404;


// Constants for enum WdViewTypeOld
const
  wdPageView = $00000003;
  wdOnlineView = $00000006;


// Constants for enum WdFramesetType
const
  wdFramesetTypeFrameset = $00000000;
  wdFramesetTypeFrame = $00000001;


// Constants for enum WdFramesetSizeType
const
  wdFramesetSizeTypePercent = $00000000;
  wdFramesetSizeTypeFixed = $00000001;
  wdFramesetSizeTypeRelative = $00000002;


// Constants for enum WdFramesetNewFrameLocation
const
  wdFramesetNewFrameAbove = $00000000;
  wdFramesetNewFrameBelow = $00000001;
  wdFramesetNewFrameRight = $00000002;
  wdFramesetNewFrameLeft = $00000003;


// Constants for enum WdScrollbarType
const
  wdScrollbarTypeAuto = $00000000;
  wdScrollbarTypeYes = $00000001;
  wdScrollbarTypeNo = $00000002;


// Constants for enum WdTwoLinesInOneType
const
  wdTwoLinesInOneNone = $00000000;
  wdTwoLinesInOneNoBrackets = $00000001;
  wdTwoLinesInOneParentheses = $00000002;
  wdTwoLinesInOneSquareBrackets = $00000003;
  wdTwoLinesInOneAngleBrackets = $00000004;
  wdTwoLinesInOneCurlyBrackets = $00000005;


// Constants for enum WdHorizontalInVerticalType
const
  wdHorizontalInVerticalNone = $00000000;
  wdHorizontalInVerticalFitInLine = $00000001;
  wdHorizontalInVerticalResizeLine = $00000002;


// Constants for enum WdHorizontalLineAlignment
const
  wdHorizontalLineAlignLeft = $00000000;
  wdHorizontalLineAlignCenter = $00000001;
  wdHorizontalLineAlignRight = $00000002;


// Constants for enum WdHorizontalLineWidthType
const
  wdHorizontalLinePercentWidth = $FFFFFFFF;
  wdHorizontalLineFixedWidth = $FFFFFFFE;


// Constants for enum WdPhoneticGuideAlignmentType
const
  wdPhoneticGuideAlignmentCenter = $00000000;
  wdPhoneticGuideAlignmentZeroOneZero = $00000001;
  wdPhoneticGuideAlignmentOneTwoOne = $00000002;
  wdPhoneticGuideAlignmentLeft = $00000003;
  wdPhoneticGuideAlignmentRight = $00000004;
  wdPhoneticGuideAlignmentRightVertical = $00000005;


// Constants for enum WdNewDocumentType
const
  wdNewBlankDocument = $00000000;
  wdNewWebPage = $00000001;
  wdNewEmailMessage = $00000002;
  wdNewFrameset = $00000003;
  wdNewXMLDocument = $00000004;


// Constants for enum WdKana
const
  wdKanaKatakana = $00000008;
  wdKanaHiragana = $00000009;


// Constants for enum WdCharacterWidth
const
  wdWidthHalfWidth = $00000006;
  wdWidthFullWidth = $00000007;


// Constants for enum WdNumberStyleWordBasicBiDi
const
  wdListNumberStyleBidi1 = $00000031;
  wdListNumberStyleBidi2 = $00000032;
  wdCaptionNumberStyleBidiLetter1 = $00000031;
  wdCaptionNumberStyleBidiLetter2 = $00000032;
  wdNoteNumberStyleBidiLetter1 = $00000031;
  wdNoteNumberStyleBidiLetter2 = $00000032;
  wdPageNumberStyleBidiLetter1 = $00000031;
  wdPageNumberStyleBidiLetter2 = $00000032;


// Constants for enum WdTCSCConverterDirection
const
  wdTCSCConverterDirectionSCTC = $00000000;
  wdTCSCConverterDirectionTCSC = $00000001;
  wdTCSCConverterDirectionAuto = $00000002;


// Constants for enum WdDisableFeaturesIntroducedAfter
const
  wd70 = $00000000;
  wd70FE = $00000001;
  wd80 = $00000002;


// Constants for enum WdWrapTypeMerged
const
  wdWrapMergeInline = $00000000;
  wdWrapMergeSquare = $00000001;
  wdWrapMergeTight = $00000002;
  wdWrapMergeBehind = $00000003;
  wdWrapMergeFront = $00000004;
  wdWrapMergeThrough = $00000005;
  wdWrapMergeTopBottom = $00000006;


// Constants for enum WdRecoveryType
const
  wdPasteDefault = $00000000;
  wdSingleCellText = $00000005;
  wdSingleCellTable = $00000006;
  wdListContinueNumbering = $00000007;
  wdListRestartNumbering = $00000008;
  wdTableInsertAsRows = $0000000B;
  wdTableAppendTable = $0000000A;
  wdTableOriginalFormatting = $0000000C;
  wdChartPicture = $0000000D;
  wdChart = $0000000E;
  wdChartLinked = $0000000F;
  wdFormatOriginalFormatting = $00000010;
  wdFormatSurroundingFormattingWithEmphasis = $00000014;
  wdFormatPlainText = $00000016;
  wdTableOverwriteCells = $00000017;
  wdListCombineWithExistingList = $00000018;
  wdListDontMerge = $00000019;
  wdUseDestinationStylesRecovery = $00000013;


// Constants for enum WdLineEndingType
const
  wdCRLF = $00000000;
  wdCROnly = $00000001;
  wdLFOnly = $00000002;
  wdLFCR = $00000003;
  wdLSPS = $00000004;


// Constants for enum WdStyleSheetLinkType
const
  wdStyleSheetLinkTypeLinked = $00000000;
  wdStyleSheetLinkTypeImported = $00000001;


// Constants for enum WdStyleSheetPrecedence
const
  wdStyleSheetPrecedenceHigher = $FFFFFFFF;
  wdStyleSheetPrecedenceLower = $FFFFFFFE;
  wdStyleSheetPrecedenceHighest = $00000001;
  wdStyleSheetPrecedenceLowest = $00000000;


// Constants for enum WdEmailHTMLFidelity
const
  wdEmailHTMLFidelityLow = $00000001;
  wdEmailHTMLFidelityMedium = $00000002;
  wdEmailHTMLFidelityHigh = $00000003;


// Constants for enum WdMailMergeMailFormat
const
  wdMailFormatPlainText = $00000000;
  wdMailFormatHTML = $00000001;


// Constants for enum WdMappedDataFields
const
  wdUniqueIdentifier = $00000001;
  wdCourtesyTitle = $00000002;
  wdFirstName = $00000003;
  wdMiddleName = $00000004;
  wdLastName = $00000005;
  wdSuffix = $00000006;
  wdNickname = $00000007;
  wdJobTitle = $00000008;
  wdCompany = $00000009;
  wdAddress1 = $0000000A;
  wdAddress2 = $0000000B;
  wdCity = $0000000C;
  wdState = $0000000D;
  wdPostalCode = $0000000E;
  wdCountryRegion = $0000000F;
  wdBusinessPhone = $00000010;
  wdBusinessFax = $00000011;
  wdHomePhone = $00000012;
  wdHomeFax = $00000013;
  wdEmailAddress = $00000014;
  wdWebPageURL = $00000015;
  wdSpouseCourtesyTitle = $00000016;
  wdSpouseFirstName = $00000017;
  wdSpouseMiddleName = $00000018;
  wdSpouseLastName = $00000019;
  wdSpouseNickname = $0000001A;
  wdRubyFirstName = $0000001B;
  wdRubyLastName = $0000001C;
  wdAddress3 = $0000001D;
  wdDepartment = $0000001E;


// Constants for enum WdConditionCode
const
  wdFirstRow = $00000000;
  wdLastRow = $00000001;
  wdOddRowBanding = $00000002;
  wdEvenRowBanding = $00000003;
  wdFirstColumn = $00000004;
  wdLastColumn = $00000005;
  wdOddColumnBanding = $00000006;
  wdEvenColumnBanding = $00000007;
  wdNECell = $00000008;
  wdNWCell = $00000009;
  wdSECell = $0000000A;
  wdSWCell = $0000000B;


// Constants for enum WdCompareTarget
const
  wdCompareTargetSelected = $00000000;
  wdCompareTargetCurrent = $00000001;
  wdCompareTargetNew = $00000002;


// Constants for enum WdMergeTarget
const
  wdMergeTargetSelected = $00000000;
  wdMergeTargetCurrent = $00000001;
  wdMergeTargetNew = $00000002;


// Constants for enum WdUseFormattingFrom
const
  wdFormattingFromCurrent = $00000000;
  wdFormattingFromSelected = $00000001;
  wdFormattingFromPrompt = $00000002;


// Constants for enum WdRevisionsView
const
  wdRevisionsViewFinal = $00000000;
  wdRevisionsViewOriginal = $00000001;


// Constants for enum WdRevisionsMode
const
  wdBalloonRevisions = $00000000;
  wdInLineRevisions = $00000001;
  wdMixedRevisions = $00000002;


// Constants for enum WdRevisionsBalloonWidthType
const
  wdBalloonWidthPercent = $00000000;
  wdBalloonWidthPoints = $00000001;


// Constants for enum WdRevisionsBalloonPrintOrientation
const
  wdBalloonPrintOrientationAuto = $00000000;
  wdBalloonPrintOrientationPreserve = $00000001;
  wdBalloonPrintOrientationForceLandscape = $00000002;


// Constants for enum WdRevisionsBalloonMargin
const
  wdLeftMargin = $00000000;
  wdRightMargin = $00000001;


// Constants for enum WdTaskPanes
const
  wdTaskPaneFormatting = $00000000;
  wdTaskPaneRevealFormatting = $00000001;
  wdTaskPaneMailMerge = $00000002;
  wdTaskPaneTranslate = $00000003;
  wdTaskPaneSearch = $00000004;
  wdTaskPaneXMLStructure = $00000005;
  wdTaskPaneDocumentProtection = $00000006;
  wdTaskPaneDocumentActions = $00000007;
  wdTaskPaneSharedWorkspace = $00000008;
  wdTaskPaneHelp = $00000009;
  wdTaskPaneResearch = $0000000A;
  wdTaskPaneFaxService = $0000000B;
  wdTaskPaneXMLDocument = $0000000C;
  wdTaskPaneDocumentUpdates = $0000000D;
  wdTaskPaneSignature = $0000000E;
  wdTaskPaneStyleInspector = $0000000F;
  wdTaskPaneDocumentManagement = $00000010;
  wdTaskPaneApplyStyles = $00000011;
  wdTaskPaneNav = $00000012;
  wdTaskPaneSelection = $00000013;
  wdTaskPaneProofing = $00000014;
  wdTaskPaneXMLMapping = $00000015;
  wdTaskPaneRevPaneFlex = $00000016;
  wdTaskPaneThesaurus = $00000017;


// Constants for enum WdShowFilter
const
  wdShowFilterStylesAvailable = $00000000;
  wdShowFilterStylesInUse = $00000001;
  wdShowFilterStylesAll = $00000002;
  wdShowFilterFormattingInUse = $00000003;
  wdShowFilterFormattingAvailable = $00000004;
  wdShowFilterFormattingRecommended = $00000005;


// Constants for enum WdMergeSubType
const
  wdMergeSubTypeOther = $00000000;
  wdMergeSubTypeAccess = $00000001;
  wdMergeSubTypeOAL = $00000002;
  wdMergeSubTypeOLEDBWord = $00000003;
  wdMergeSubTypeWorks = $00000004;
  wdMergeSubTypeOLEDBText = $00000005;
  wdMergeSubTypeOutlook = $00000006;
  wdMergeSubTypeWord = $00000007;
  wdMergeSubTypeWord2000 = $00000008;


// Constants for enum WdDocumentDirection
const
  wdLeftToRight = $00000000;
  wdRightToLeft = $00000001;


// Constants for enum WdLanguageID2000
const
  wdChineseHongKong = $00000C04;
  wdChineseMacao = $00001404;
  wdEnglishTrinidad = $00002C09;


// Constants for enum WdRectangleType
const
  wdTextRectangle = $00000000;
  wdShapeRectangle = $00000001;
  wdMarkupRectangle = $00000002;
  wdMarkupRectangleButton = $00000003;
  wdPageBorderRectangle = $00000004;
  wdLineBetweenColumnRectangle = $00000005;
  wdSelection = $00000006;
  wdSystem = $00000007;
  wdMarkupRectangleArea = $00000008;
  wdReadingModeNavigation = $00000009;
  wdMarkupRectangleMoveMatch = $0000000A;
  wdReadingModePanningArea = $0000000B;
  wdMailNavArea = $0000000C;
  wdDocumentControlRectangle = $0000000D;


// Constants for enum WdLineType
const
  wdTextLine = $00000000;
  wdTableRow = $00000001;


// Constants for enum WdXMLNodeType
const
  wdXMLNodeElement = $00000001;
  wdXMLNodeAttribute = $00000002;


// Constants for enum WdXMLSelectionChangeReason
const
  wdXMLSelectionChangeReasonMove = $00000000;
  wdXMLSelectionChangeReasonInsert = $00000001;
  wdXMLSelectionChangeReasonDelete = $00000002;


// Constants for enum WdXMLNodeLevel
const
  wdXMLNodeLevelInline = $00000000;
  wdXMLNodeLevelParagraph = $00000001;
  wdXMLNodeLevelRow = $00000002;
  wdXMLNodeLevelCell = $00000003;


// Constants for enum WdSmartTagControlType
const
  wdControlSmartTag = $00000001;
  wdControlLink = $00000002;
  wdControlHelp = $00000003;
  wdControlHelpURL = $00000004;
  wdControlSeparator = $00000005;
  wdControlButton = $00000006;
  wdControlLabel = $00000007;
  wdControlImage = $00000008;
  wdControlCheckbox = $00000009;
  wdControlTextbox = $0000000A;
  wdControlListbox = $0000000B;
  wdControlCombo = $0000000C;
  wdControlActiveX = $0000000D;
  wdControlDocumentFragment = $0000000E;
  wdControlDocumentFragmentURL = $0000000F;
  wdControlRadioGroup = $00000010;


// Constants for enum WdEditorType
const
  wdEditorEveryone = $FFFFFFFF;
  wdEditorOwners = $FFFFFFFC;
  wdEditorEditors = $FFFFFFFB;
  wdEditorCurrent = $FFFFFFFA;


// Constants for enum WdXMLValidationStatus
const
  wdXMLValidationStatusOK = $00000000;
  wdXMLValidationStatusCustom = $C00CE000;


// Constants for enum WdStyleSort
const
  wdStyleSortByName = $00000000;
  wdStyleSortRecommended = $00000001;
  wdStyleSortByFont = $00000002;
  wdStyleSortByBasedOn = $00000003;
  wdStyleSortByType = $00000004;


// Constants for enum WdRemoveDocInfoType
const
  wdRDIComments = $00000001;
  wdRDIRevisions = $00000002;
  wdRDIVersions = $00000003;
  wdRDIRemovePersonalInformation = $00000004;
  wdRDIEmailHeader = $00000005;
  wdRDIRoutingSlip = $00000006;
  wdRDISendForReview = $00000007;
  wdRDIDocumentProperties = $00000008;
  wdRDITemplate = $00000009;
  wdRDIDocumentWorkspace = $0000000A;
  wdRDIInkAnnotations = $0000000B;
  wdRDIDocumentServerProperties = $0000000E;
  wdRDIDocumentManagementPolicy = $0000000F;
  wdRDIContentType = $00000010;
  wdRDITaskpaneWebExtensions = $00000011;
  wdRDIAtMentions = $00000012;
  wdRDIDocumentTasks = $00000013;
  wdRDIAll = $00000063;


// Constants for enum WdCheckInVersionType
const
  wdCheckInMinorVersion = $00000000;
  wdCheckInMajorVersion = $00000001;
  wdCheckInOverwriteVersion = $00000002;


// Constants for enum WdMoveToTextMark
const
  wdMoveToTextMarkNone = $00000000;
  wdMoveToTextMarkBold = $00000001;
  wdMoveToTextMarkItalic = $00000002;
  wdMoveToTextMarkUnderline = $00000003;
  wdMoveToTextMarkDoubleUnderline = $00000004;
  wdMoveToTextMarkColorOnly = $00000005;
  wdMoveToTextMarkStrikeThrough = $00000006;
  wdMoveToTextMarkDoubleStrikeThrough = $00000007;


// Constants for enum WdMoveFromTextMark
const
  wdMoveFromTextMarkHidden = $00000000;
  wdMoveFromTextMarkDoubleStrikeThrough = $00000001;
  wdMoveFromTextMarkStrikeThrough = $00000002;
  wdMoveFromTextMarkCaret = $00000003;
  wdMoveFromTextMarkPound = $00000004;
  wdMoveFromTextMarkNone = $00000005;
  wdMoveFromTextMarkBold = $00000006;
  wdMoveFromTextMarkItalic = $00000007;
  wdMoveFromTextMarkUnderline = $00000008;
  wdMoveFromTextMarkDoubleUnderline = $00000009;
  wdMoveFromTextMarkColorOnly = $0000000A;


// Constants for enum WdOMathFunctionType
const
  wdOMathFunctionAcc = $00000001;
  wdOMathFunctionBar = $00000002;
  wdOMathFunctionBox = $00000003;
  wdOMathFunctionBorderBox = $00000004;
  wdOMathFunctionDelim = $00000005;
  wdOMathFunctionEqArray = $00000006;
  wdOMathFunctionFrac = $00000007;
  wdOMathFunctionFunc = $00000008;
  wdOMathFunctionGroupChar = $00000009;
  wdOMathFunctionLimLow = $0000000A;
  wdOMathFunctionLimUpp = $0000000B;
  wdOMathFunctionMat = $0000000C;
  wdOMathFunctionNary = $0000000D;
  wdOMathFunctionPhantom = $0000000E;
  wdOMathFunctionScrPre = $0000000F;
  wdOMathFunctionRad = $00000010;
  wdOMathFunctionScrSub = $00000011;
  wdOMathFunctionScrSubSup = $00000012;
  wdOMathFunctionScrSup = $00000013;
  wdOMathFunctionText = $00000014;
  wdOMathFunctionNormalText = $00000015;
  wdOMathFunctionLiteralText = $00000016;


// Constants for enum WdOMathHorizAlignType
const
  wdOMathHorizAlignCenter = $00000000;
  wdOMathHorizAlignLeft = $00000001;
  wdOMathHorizAlignRight = $00000002;


// Constants for enum WdOMathVertAlignType
const
  wdOMathVertAlignCenter = $00000000;
  wdOMathVertAlignTop = $00000001;
  wdOMathVertAlignBottom = $00000002;


// Constants for enum WdOMathFracType
const
  wdOMathFracBar = $00000000;
  wdOMathFracNoBar = $00000001;
  wdOMathFracSkw = $00000002;
  wdOMathFracLin = $00000003;


// Constants for enum WdOMathSpacingRule
const
  wdOMathSpacingSingle = $00000000;
  wdOMathSpacing1pt5 = $00000001;
  wdOMathSpacingDouble = $00000002;
  wdOMathSpacingExactly = $00000003;
  wdOMathSpacingMultiple = $00000004;


// Constants for enum WdOMathType
const
  wdOMathDisplay = $00000000;
  wdOMathInline = $00000001;


// Constants for enum WdOMathShapeType
const
  wdOMathShapeCentered = $00000000;
  wdOMathShapeMatch = $00000001;


// Constants for enum WdOMathJc
const
  wdOMathJcCenterGroup = $00000001;
  wdOMathJcCenter = $00000002;
  wdOMathJcLeft = $00000003;
  wdOMathJcRight = $00000004;
  wdOMathJcInline = $00000007;


// Constants for enum WdOMathBreakBin
const
  wdOMathBreakBinBefore = $00000000;
  wdOMathBreakBinAfter = $00000001;
  wdOMathBreakBinRepeat = $00000002;


// Constants for enum WdOMathBreakSub
const
  wdOMathBreakSubMinusMinus = $00000000;
  wdOMathBreakSubPlusMinus = $00000001;
  wdOMathBreakSubMinusPlus = $00000002;


// Constants for enum WdReadingLayoutMargin
const
  wdAutomaticMargin = $00000000;
  wdSuppressMargin = $00000001;
  wdFullMargin = $00000002;


// Constants for enum WdContentControlType
const
  wdContentControlRichText = $00000000;
  wdContentControlText = $00000001;
  wdContentControlPicture = $00000002;
  wdContentControlComboBox = $00000003;
  wdContentControlDropdownList = $00000004;
  wdContentControlBuildingBlockGallery = $00000005;
  wdContentControlDate = $00000006;
  wdContentControlGroup = $00000007;
  wdContentControlCheckBox = $00000008;
  wdContentControlRepeatingSection = $00000009;


// Constants for enum WdCompareDestination
const
  wdCompareDestinationOriginal = $00000000;
  wdCompareDestinationRevised = $00000001;
  wdCompareDestinationNew = $00000002;


// Constants for enum WdGranularity
const
  wdGranularityCharLevel = $00000000;
  wdGranularityWordLevel = $00000001;


// Constants for enum WdMergeFormatFrom
const
  wdMergeFormatFromOriginal = $00000000;
  wdMergeFormatFromRevised = $00000001;
  wdMergeFormatFromPrompt = $00000002;


// Constants for enum WdShowSourceDocuments
const
  wdShowSourceDocumentsNone = $00000000;
  wdShowSourceDocumentsOriginal = $00000001;
  wdShowSourceDocumentsRevised = $00000002;
  wdShowSourceDocumentsBoth = $00000003;


// Constants for enum WdPasteOptions
const
  wdKeepSourceFormatting = $00000000;
  wdMatchDestinationFormatting = $00000001;
  wdKeepTextOnly = $00000002;
  wdUseDestinationStyles = $00000003;


// Constants for enum WdBuildingBlockTypes
const
  wdTypeQuickParts = $00000001;
  wdTypeCoverPage = $00000002;
  wdTypeEquations = $00000003;
  wdTypeFooters = $00000004;
  wdTypeHeaders = $00000005;
  wdTypePageNumber = $00000006;
  wdTypeTables = $00000007;
  wdTypeWatermarks = $00000008;
  wdTypeAutoText = $00000009;
  wdTypeTextBox = $0000000A;
  wdTypePageNumberTop = $0000000B;
  wdTypePageNumberBottom = $0000000C;
  wdTypePageNumberPage = $0000000D;
  wdTypeTableOfContents = $0000000E;
  wdTypeCustomQuickParts = $0000000F;
  wdTypeCustomCoverPage = $00000010;
  wdTypeCustomEquations = $00000011;
  wdTypeCustomFooters = $00000012;
  wdTypeCustomHeaders = $00000013;
  wdTypeCustomPageNumber = $00000014;
  wdTypeCustomTables = $00000015;
  wdTypeCustomWatermarks = $00000016;
  wdTypeCustomAutoText = $00000017;
  wdTypeCustomTextBox = $00000018;
  wdTypeCustomPageNumberTop = $00000019;
  wdTypeCustomPageNumberBottom = $0000001A;
  wdTypeCustomPageNumberPage = $0000001B;
  wdTypeCustomTableOfContents = $0000001C;
  wdTypeCustom1 = $0000001D;
  wdTypeCustom2 = $0000001E;
  wdTypeCustom3 = $0000001F;
  wdTypeCustom4 = $00000020;
  wdTypeCustom5 = $00000021;
  wdTypeBibliography = $00000022;
  wdTypeCustomBibliography = $00000023;


// Constants for enum WdAlignmentTabRelative
const
  wdMargin = $00000000;
  wdIndent = $00000001;


// Constants for enum WdAlignmentTabAlignment
const
  wdLeft = $00000000;
  wdCenter = $00000001;
  wdRight = $00000002;


// Constants for enum WdCellColor
const
  wdCellColorByAuthor = $FFFFFFFF;
  wdCellColorNoHighlight = $00000000;
  wdCellColorPink = $00000001;
  wdCellColorLightBlue = $00000002;
  wdCellColorLightYellow = $00000003;
  wdCellColorLightPurple = $00000004;
  wdCellColorLightOrange = $00000005;
  wdCellColorLightGreen = $00000006;
  wdCellColorLightGray = $00000007;


// Constants for enum WdTextboxTightWrap
const
  wdTightNone = $00000000;
  wdTightAll = $00000001;
  wdTightFirstAndLastLines = $00000002;
  wdTightFirstLineOnly = $00000003;
  wdTightLastLineOnly = $00000004;


// Constants for enum WdShapePositionRelative
const
  wdShapePositionRelativeNone = $FFF0BDC1;


// Constants for enum WdShapeSizeRelative
const
  wdShapeSizeRelativeNone = $FFF0BDC1;


// Constants for enum WdRelativeHorizontalSize
const
  wdRelativeHorizontalSizeMargin = $00000000;
  wdRelativeHorizontalSizePage = $00000001;
  wdRelativeHorizontalSizeLeftMarginArea = $00000002;
  wdRelativeHorizontalSizeRightMarginArea = $00000003;
  wdRelativeHorizontalSizeInnerMarginArea = $00000004;
  wdRelativeHorizontalSizeOuterMarginArea = $00000005;


// Constants for enum WdRelativeVerticalSize
const
  wdRelativeVerticalSizeMargin = $00000000;
  wdRelativeVerticalSizePage = $00000001;
  wdRelativeVerticalSizeTopMarginArea = $00000002;
  wdRelativeVerticalSizeBottomMarginArea = $00000003;
  wdRelativeVerticalSizeInnerMarginArea = $00000004;
  wdRelativeVerticalSizeOuterMarginArea = $00000005;


// Constants for enum WdThemeColorIndex
const
  wdNotThemeColor = $FFFFFFFF;
  wdThemeColorMainDark1 = $00000000;
  wdThemeColorMainLight1 = $00000001;
  wdThemeColorMainDark2 = $00000002;
  wdThemeColorMainLight2 = $00000003;
  wdThemeColorAccent1 = $00000004;
  wdThemeColorAccent2 = $00000005;
  wdThemeColorAccent3 = $00000006;
  wdThemeColorAccent4 = $00000007;
  wdThemeColorAccent5 = $00000008;
  wdThemeColorAccent6 = $00000009;
  wdThemeColorHyperlink = $0000000A;
  wdThemeColorHyperlinkFollowed = $0000000B;
  wdThemeColorBackground1 = $0000000C;
  wdThemeColorText1 = $0000000D;
  wdThemeColorBackground2 = $0000000E;
  wdThemeColorText2 = $0000000F;


// Constants for enum WdExportFormat
const
  wdExportFormatPDF = $00000011;
  wdExportFormatXPS = $00000012;


// Constants for enum WdExportOptimizeFor
const
  wdExportOptimizeForPrint = $00000000;
  wdExportOptimizeForOnScreen = $00000001;


// Constants for enum WdExportCreateBookmarks
const
  wdExportCreateNoBookmarks = $00000000;
  wdExportCreateHeadingBookmarks = $00000001;
  wdExportCreateWordBookmarks = $00000002;


// Constants for enum WdExportItem
const
  wdExportDocumentContent = $00000000;
  wdExportDocumentWithMarkup = $00000007;


// Constants for enum WdExportRange
const
  wdExportAllDocument = $00000000;
  wdExportSelection = $00000001;
  wdExportCurrentPage = $00000002;
  wdExportFromTo = $00000003;


// Constants for enum WdFrenchSpeller
const
  wdFrenchBoth = $00000000;
  wdFrenchPreReform = $00000001;
  wdFrenchPostReform = $00000002;


// Constants for enum WdDocPartInsertOptions
const
  wdInsertContent = $00000000;
  wdInsertParagraph = $00000001;
  wdInsertPage = $00000002;


// Constants for enum WdContentControlDateStorageFormat
const
  wdContentControlDateStorageText = $00000000;
  wdContentControlDateStorageDate = $00000001;
  wdContentControlDateStorageDateTime = $00000002;


// Constants for enum XlChartSplitType
const
  xlSplitByPosition = $00000001;
  xlSplitByPercentValue = $00000003;
  xlSplitByCustomSplit = $00000004;
  xlSplitByValue = $00000002;


// Constants for enum XlSizeRepresents
const
  xlSizeIsWidth = $00000002;
  xlSizeIsArea = $00000001;


// Constants for enum XlAxisGroup
const
  xlPrimary = $00000001;
  xlSecondary = $00000002;


// Constants for enum XlBackground
const
  xlBackgroundAutomatic = $FFFFEFF7;
  xlBackgroundOpaque = $00000003;
  xlBackgroundTransparent = $00000002;


// Constants for enum XlChartGallery
const
  xlBuiltIn = $00000015;
  xlUserDefined = $00000016;
  xlAnyGallery = $00000017;


// Constants for enum XlChartPicturePlacement
const
  xlSides = $00000001;
  xlEnd = $00000002;
  xlEndSides = $00000003;
  xlFront = $00000004;
  xlFrontSides = $00000005;
  xlFrontEnd = $00000006;
  xlAllFaces = $00000007;


// Constants for enum XlDataLabelSeparator
const
  xlDataLabelSeparatorDefault = $00000001;


// Constants for enum XlPattern
const
  xlPatternAutomatic = $FFFFEFF7;
  xlPatternChecker = $00000009;
  xlPatternCrissCross = $00000010;
  xlPatternDown = $FFFFEFE7;
  xlPatternGray16 = $00000011;
  xlPatternGray25 = $FFFFEFE4;
  xlPatternGray50 = $FFFFEFE3;
  xlPatternGray75 = $FFFFEFE2;
  xlPatternGray8 = $00000012;
  xlPatternGrid = $0000000F;
  xlPatternHorizontal = $FFFFEFE0;
  xlPatternLightDown = $0000000D;
  xlPatternLightHorizontal = $0000000B;
  xlPatternLightUp = $0000000E;
  xlPatternLightVertical = $0000000C;
  xlPatternNone = $FFFFEFD2;
  xlPatternSemiGray75 = $0000000A;
  xlPatternSolid = $00000001;
  xlPatternUp = $FFFFEFBE;
  xlPatternVertical = $FFFFEFBA;
  xlPatternLinearGradient = $00000FA0;
  xlPatternRectangularGradient = $00000FA1;


// Constants for enum XlPictureAppearance
const
  xlPrinter = $00000002;
  xlScreen = $00000001;


// Constants for enum XlCopyPictureFormat
const
  xlBitmap = $00000002;
  xlPicture = $FFFFEFCD;


// Constants for enum XlRgbColor
const
  xlAliceBlue = $00FFF8F0;
  xlAntiqueWhite = $00D7EBFA;
  xlAqua = $00FFFF00;
  xlAquamarine = $00D4FF7F;
  xlAzure = $00FFFFF0;
  xlBeige = $00DCF5F5;
  xlBisque = $00C4E4FF;
  xlBlack = $00000000;
  xlBlanchedAlmond = $00CDEBFF;
  xlBlue = $00FF0000;
  xlBlueViolet = $00E22B8A;
  xlBrown = $002A2AA5;
  xlBurlyWood = $0087B8DE;
  xlCadetBlue = $00A09E5F;
  xlChartreuse = $0000FF7F;
  xlCoral = $00507FFF;
  xlCornflowerBlue = $00ED9564;
  xlCornsilk = $00DCF8FF;
  xlCrimson = $003C14DC;
  xlDarkBlue = $008B0000;
  xlDarkCyan = $008B8B00;
  xlDarkGoldenrod = $000B86B8;
  xlDarkGreen = $00006400;
  xlDarkGray = $00A9A9A9;
  xlDarkGrey = $00A9A9A9;
  xlDarkKhaki = $006BB7BD;
  xlDarkMagenta = $008B008B;
  xlDarkOliveGreen = $002F6B55;
  xlDarkOrange = $00008CFF;
  xlDarkOrchid = $00CC3299;
  xlDarkRed = $0000008B;
  xlDarkSalmon = $007A96E9;
  xlDarkSeaGreen = $008FBC8F;
  xlDarkSlateBlue = $008B3D48;
  xlDarkSlateGray = $004F4F2F;
  xlDarkSlateGrey = $004F4F2F;
  xlDarkTurquoise = $00D1CE00;
  xlDarkViolet = $00D30094;
  xlDeepPink = $009314FF;
  xlDeepSkyBlue = $00FFBF00;
  xlDimGray = $00696969;
  xlDimGrey = $00696969;
  xlDodgerBlue = $00FF901E;
  xlFireBrick = $002222B2;
  xlFloralWhite = $00F0FAFF;
  xlForestGreen = $00228B22;
  xlFuchsia = $00FF00FF;
  xlGainsboro = $00DCDCDC;
  xlGhostWhite = $00FFF8F8;
  xlGold = $0000D7FF;
  xlGoldenrod = $0020A5DA;
  xlGray = $00808080;
  xlGreen = $00008000;
  xlGrey = $00808080;
  xlGreenYellow = $002FFFAD;
  xlHoneydew = $00F0FFF0;
  xlHotPink = $00B469FF;
  xlIndianRed = $005C5CCD;
  xlIndigo = $0082004B;
  xlIvory = $00F0FFFF;
  xlKhaki = $008CE6F0;
  xlLavender = $00FAE6E6;
  xlLavenderBlush = $00F5F0FF;
  xlLawnGreen = $0000FC7C;
  xlLemonChiffon = $00CDFAFF;
  xlLightBlue = $00E6D8AD;
  xlLightCoral = $008080F0;
  xlLightCyan = $008B8B00;
  xlLightGoldenrodYellow = $00D2FAFA;
  xlLightGray = $00D3D3D3;
  xlLightGreen = $0090EE90;
  xlLightGrey = $00D3D3D3;
  xlLightPink = $00C1B6FF;
  xlLightSalmon = $007AA0FF;
  xlLightSeaGreen = $00AAB220;
  xlLightSkyBlue = $00FACE87;
  xlLightSlateGray = $00998877;
  xlLightSlateGrey = $00998877;
  xlLightSteelBlue = $00DEC4B0;
  xlLightYellow = $00E0FFFF;
  xlLime = $0000FF00;
  xlLimeGreen = $0032CD32;
  xlLinen = $00E6F0FA;
  xlMaroon = $00000080;
  xlMediumAquamarine = $00AAFF66;
  xlMediumBlue = $00CD0000;
  xlMediumOrchid = $00D355BA;
  xlMediumPurple = $00DB7093;
  xlMediumSeaGreen = $0071B33C;
  xlMediumSlateBlue = $00EE687B;
  xlMediumSpringGreen = $009AFA00;
  xlMediumTurquoise = $00CCD148;
  xlMediumVioletRed = $008515C7;
  xlMidnightBlue = $00701919;
  xlMintCream = $00FAFFF5;
  xlMistyRose = $00E1E4FF;
  xlMoccasin = $00B5E4FF;
  xlNavajoWhite = $00ADDEFF;
  xlNavy = $00800000;
  xlNavyBlue = $00800000;
  xlOldLace = $00E6F5FD;
  xlOlive = $00008080;
  xlOliveDrab = $00238E6B;
  xlOrange = $0000A5FF;
  xlOrangeRed = $000045FF;
  xlOrchid = $00D670DA;
  xlPaleGoldenrod = $006BE8EE;
  xlPaleGreen = $0098FB98;
  xlPaleTurquoise = $00EEEEAF;
  xlPaleVioletRed = $009370DB;
  xlPapayaWhip = $00D5EFFF;
  xlPeachPuff = $00B9DAFF;
  xlPeru = $003F85CD;
  xlPink = $00CBC0FF;
  xlPlum = $00DDA0DD;
  xlPowderBlue = $00E6E0B0;
  xlPurple = $00800080;
  xlRed = $000000FF;
  xlRosyBrown = $008F8FBC;
  xlRoyalBlue = $00E16941;
  xlSalmon = $007280FA;
  xlSandyBrown = $0060A4F4;
  xlSeaGreen = $00578B2E;
  xlSeashell = $00EEF5FF;
  xlSienna = $002D52A0;
  xlSilver = $00C0C0C0;
  xlSkyBlue = $00EBCE87;
  xlSlateBlue = $00CD5A6A;
  xlSlateGray = $00908070;
  xlSlateGrey = $00908070;
  xlSnow = $00FAFAFF;
  xlSpringGreen = $007FFF00;
  xlSteelBlue = $00B48246;
  xlTan = $008CB4D2;
  xlTeal = $00808000;
  xlThistle = $00D8BFD8;
  xlTomato = $004763FF;
  xlTurquoise = $00D0E040;
  xlYellow = $0000FFFF;
  xlYellowGreen = $0032CD9A;
  xlViolet = $00EE82EE;
  xlWheat = $00B3DEF5;
  xlWhite = $00FFFFFF;
  xlWhiteSmoke = $00F5F5F5;


// Constants for enum XlConstants
const
  xlAutomatic = $FFFFEFF7;
  xlCombination = $FFFFEFF1;
  xlCustom = $FFFFEFEE;
  xlBar = $00000002;
  xlColumn = $00000003;
  xl3DBar = $FFFFEFFD;
  xl3DSurface = $FFFFEFF9;
  xlDefaultAutoFormat = $FFFFFFFF;
  xlNone = $FFFFEFD2;
  xlAbove = $00000000;
  xlBelow = $00000001;
  xlBoth = $00000001;
  xlBottom = $FFFFEFF5;
  xlCenter = $FFFFEFF4;
  xlChecker = $00000009;
  xlCircle = $00000008;
  xlCorner = $00000002;
  xlCrissCross = $00000010;
  xlCross = $00000004;
  xlDiamond = $00000002;
  xlDistributed = $FFFFEFEB;
  xlFill = $00000005;
  xlFixedValue = $00000001;
  xlGeneral = $00000001;
  xlGray16 = $00000011;
  xlGray25 = $FFFFEFE4;
  xlGray50 = $FFFFEFE3;
  xlGray75 = $FFFFEFE2;
  xlGray8 = $00000012;
  xlGrid = $0000000F;
  xlHigh = $FFFFEFE1;
  xlInside = $00000002;
  xlJustify = $FFFFEFDE;
  xlLeft = $FFFFEFDD;
  xlLightDown = $0000000D;
  xlLightHorizontal = $0000000B;
  xlLightUp = $0000000E;
  xlLightVertical = $0000000C;
  xlLow = $FFFFEFDA;
  xlMaximum = $00000002;
  xlMinimum = $00000004;
  xlMinusValues = $00000003;
  xlNextToAxis = $00000004;
  xlOpaque = $00000003;
  xlOutside = $00000003;
  xlPercent = $00000002;
  xlPlus = $00000009;
  xlPlusValues = $00000002;
  xlRight = $FFFFEFC8;
  xlScale = $00000003;
  xlSemiGray75 = $0000000A;
  xlShowLabel = $00000004;
  xlShowLabelAndPercent = $00000005;
  xlShowPercent = $00000003;
  xlShowValue = $00000002;
  xlSingle = $00000002;
  xlSolid = $00000001;
  xlSquare = $00000001;
  xlStar = $00000005;
  xlStError = $00000004;
  xlTop = $FFFFEFC0;
  xlTransparent = $00000002;
  xlTriangle = $00000003;


// Constants for enum XlReadingOrder
const
  xlContext = $FFFFEC76;
  xlLTR = $FFFFEC75;
  xlRTL = $FFFFEC74;


// Constants for enum XlBorderWeight
const
  xlHairline = $00000001;
  xlMedium = $FFFFEFD6;
  xlThick = $00000004;
  xlThin = $00000002;


// Constants for enum XlLegendPosition
const
  xlLegendPositionBottom = $FFFFEFF5;
  xlLegendPositionCorner = $00000002;
  xlLegendPositionLeft = $FFFFEFDD;
  xlLegendPositionRight = $FFFFEFC8;
  xlLegendPositionTop = $FFFFEFC0;
  xlLegendPositionCustom = $FFFFEFBF;


// Constants for enum XlUnderlineStyle
const
  xlUnderlineStyleDouble = $FFFFEFE9;
  xlUnderlineStyleDoubleAccounting = $00000005;
  xlUnderlineStyleNone = $FFFFEFD2;
  xlUnderlineStyleSingle = $00000002;
  xlUnderlineStyleSingleAccounting = $00000004;


// Constants for enum XlColorIndex
const
  xlColorIndexAutomatic = $FFFFEFF7;
  xlColorIndexNone = $FFFFEFD2;


// Constants for enum XlMarkerStyle
const
  xlMarkerStyleAutomatic = $FFFFEFF7;
  xlMarkerStyleCircle = $00000008;
  xlMarkerStyleDash = $FFFFEFED;
  xlMarkerStyleDiamond = $00000002;
  xlMarkerStyleDot = $FFFFEFEA;
  xlMarkerStyleNone = $FFFFEFD2;
  xlMarkerStylePicture = $FFFFEFCD;
  xlMarkerStylePlus = $00000009;
  xlMarkerStyleSquare = $00000001;
  xlMarkerStyleStar = $00000005;
  xlMarkerStyleTriangle = $00000003;
  xlMarkerStyleX = $FFFFEFB8;


// Constants for enum XlRowCol
const
  xlColumns = $00000002;
  xlRows = $00000001;


// Constants for enum XlDataLabelsType
const
  xlDataLabelsShowNone = $FFFFEFD2;
  xlDataLabelsShowValue = $00000002;
  xlDataLabelsShowPercent = $00000003;
  xlDataLabelsShowLabel = $00000004;
  xlDataLabelsShowLabelAndPercent = $00000005;
  xlDataLabelsShowBubbleSizes = $00000006;


// Constants for enum XlErrorBarInclude
const
  xlErrorBarIncludeBoth = $00000001;
  xlErrorBarIncludeMinusValues = $00000003;
  xlErrorBarIncludeNone = $FFFFEFD2;
  xlErrorBarIncludePlusValues = $00000002;


// Constants for enum XlErrorBarType
const
  xlErrorBarTypeCustom = $FFFFEFEE;
  xlErrorBarTypeFixedValue = $00000001;
  xlErrorBarTypePercent = $00000002;
  xlErrorBarTypeStDev = $FFFFEFC5;
  xlErrorBarTypeStError = $00000004;


// Constants for enum XlErrorBarDirection
const
  xlChartX = $FFFFEFB8;
  xlChartY = $00000001;


// Constants for enum XlChartPictureType
const
  xlStackScale = $00000003;
  xlStack = $00000002;
  xlStretch = $00000001;


// Constants for enum XlChartItem
const
  xlDataLabel = $00000000;
  xlChartArea = $00000002;
  xlSeries = $00000003;
  xlChartTitle = $00000004;
  xlWalls = $00000005;
  xlCorners = $00000006;
  xlDataTable = $00000007;
  xlTrendline = $00000008;
  xlErrorBars = $00000009;
  xlXErrorBars = $0000000A;
  xlYErrorBars = $0000000B;
  xlLegendEntry = $0000000C;
  xlLegendKey = $0000000D;
  xlShape = $0000000E;
  xlMajorGridlines = $0000000F;
  xlMinorGridlines = $00000010;
  xlAxisTitle = $00000011;
  xlUpBars = $00000012;
  xlPlotArea = $00000013;
  xlDownBars = $00000014;
  xlAxis = $00000015;
  xlSeriesLines = $00000016;
  xlFloor = $00000017;
  xlLegend = $00000018;
  xlHiLoLines = $00000019;
  xlDropLines = $0000001A;
  xlRadarAxisLabels = $0000001B;
  xlNothing = $0000001C;
  xlLeaderLines = $0000001D;
  xlDisplayUnitLabel = $0000001E;
  xlPivotChartFieldButton = $0000001F;
  xlPivotChartDropZone = $00000020;


// Constants for enum XlBarShape
const
  xlBox = $00000000;
  xlPyramidToPoint = $00000001;
  xlPyramidToMax = $00000002;
  xlCylinder = $00000003;
  xlConeToPoint = $00000004;
  xlConeToMax = $00000005;


// Constants for enum XlEndStyleCap
const
  xlCap = $00000001;
  xlNoCap = $00000002;


// Constants for enum XlTrendlineType
const
  xlExponential = $00000005;
  xlLinear = $FFFFEFDC;
  xlLogarithmic = $FFFFEFDB;
  xlMovingAvg = $00000006;
  xlPolynomial = $00000003;
  xlPower = $00000004;


// Constants for enum XlAxisType
const
  xlCategory = $00000001;
  xlSeriesAxis = $00000003;
  xlValue = $00000002;


// Constants for enum XlAxisCrosses
const
  xlAxisCrossesAutomatic = $FFFFEFF7;
  xlAxisCrossesCustom = $FFFFEFEE;
  xlAxisCrossesMaximum = $00000002;
  xlAxisCrossesMinimum = $00000004;


// Constants for enum XlTickMark
const
  xlTickMarkCross = $00000004;
  xlTickMarkInside = $00000002;
  xlTickMarkNone = $FFFFEFD2;
  xlTickMarkOutside = $00000003;


// Constants for enum XlScaleType
const
  xlScaleLinear = $FFFFEFDC;
  xlScaleLogarithmic = $FFFFEFDB;


// Constants for enum XlTickLabelPosition
const
  xlTickLabelPositionHigh = $FFFFEFE1;
  xlTickLabelPositionLow = $FFFFEFDA;
  xlTickLabelPositionNextToAxis = $00000004;
  xlTickLabelPositionNone = $FFFFEFD2;


// Constants for enum XlTimeUnit
const
  xlDays = $00000000;
  xlMonths = $00000001;
  xlYears = $00000002;


// Constants for enum XlCategoryType
const
  xlCategoryScale = $00000002;
  xlTimeScale = $00000003;
  xlAutomaticScale = $FFFFEFF7;


// Constants for enum XlDisplayUnit
const
  xlHundreds = $FFFFFFFE;
  xlThousands = $FFFFFFFD;
  xlTenThousands = $FFFFFFFC;
  xlHundredThousands = $FFFFFFFB;
  xlMillions = $FFFFFFFA;
  xlTenMillions = $FFFFFFF9;
  xlHundredMillions = $FFFFFFF8;
  xlThousandMillions = $FFFFFFF7;
  xlMillionMillions = $FFFFFFF6;


// Constants for enum XlOrientation
const
  xlDownward = $FFFFEFB6;
  xlHorizontal = $FFFFEFE0;
  xlUpward = $FFFFEFB5;
  xlVertical = $FFFFEFBA;


// Constants for enum XlTickLabelOrientation
const
  xlTickLabelOrientationAutomatic = $FFFFEFF7;
  xlTickLabelOrientationDownward = $FFFFEFB6;
  xlTickLabelOrientationHorizontal = $FFFFEFE0;
  xlTickLabelOrientationUpward = $FFFFEFB5;
  xlTickLabelOrientationVertical = $FFFFEFBA;


// Constants for enum XlDisplayBlanksAs
const
  xlInterpolated = $00000003;
  xlNotPlotted = $00000001;
  xlZero = $00000002;


// Constants for enum XlDataLabelPosition
const
  xlLabelPositionCenter = $FFFFEFF4;
  xlLabelPositionAbove = $00000000;
  xlLabelPositionBelow = $00000001;
  xlLabelPositionLeft = $FFFFEFDD;
  xlLabelPositionRight = $FFFFEFC8;
  xlLabelPositionOutsideEnd = $00000002;
  xlLabelPositionInsideEnd = $00000003;
  xlLabelPositionInsideBase = $00000004;
  xlLabelPositionBestFit = $00000005;
  xlLabelPositionMixed = $00000006;
  xlLabelPositionCustom = $00000007;


// Constants for enum XlPivotFieldOrientation
const
  xlColumnField = $00000002;
  xlDataField = $00000004;
  xlHidden = $00000000;
  xlPageField = $00000003;
  xlRowField = $00000001;


// Constants for enum XlHAlign
const
  xlHAlignCenter = $FFFFEFF4;
  xlHAlignCenterAcrossSelection = $00000007;
  xlHAlignDistributed = $FFFFEFEB;
  xlHAlignFill = $00000005;
  xlHAlignGeneral = $00000001;
  xlHAlignJustify = $FFFFEFDE;
  xlHAlignLeft = $FFFFEFDD;
  xlHAlignRight = $FFFFEFC8;


// Constants for enum XlVAlign
const
  xlVAlignBottom = $FFFFEFF5;
  xlVAlignCenter = $FFFFEFF4;
  xlVAlignDistributed = $FFFFEFEB;
  xlVAlignJustify = $FFFFEFDE;
  xlVAlignTop = $FFFFEFC0;


// Constants for enum XlLineStyle
const
  xlContinuous = $00000001;
  xlDash = $FFFFEFED;
  xlDashDot = $00000004;
  xlDashDotDot = $00000005;
  xlDot = $FFFFEFEA;
  xlDouble = $FFFFEFE9;
  xlSlantDashDot = $0000000D;
  xlLineStyleNone = $FFFFEFD2;


// Constants for enum XlChartElementPosition
const
  xlChartElementPositionAutomatic = $FFFFEFF7;
  xlChartElementPositionCustom = $FFFFEFEE;


// Constants for enum WdUpdateStyleListBehavior
const
  wdListBehaviorKeepPreviousPattern = $00000000;
  wdListBehaviorAddBulletsNumbering = $00000001;


// Constants for enum WdApplyQuickStyleSets
const
  wdSessionStartSet = $00000001;
  wdTemplateSet = $00000002;


// Constants for enum WdLigatures
const
  wdLigaturesNone = $00000000;
  wdLigaturesStandard = $00000001;
  wdLigaturesContextual = $00000002;
  wdLigaturesHistorical = $00000004;
  wdLigaturesDiscretional = $00000008;
  wdLigaturesStandardContextual = $00000003;
  wdLigaturesStandardHistorical = $00000005;
  wdLigaturesContextualHistorical = $00000006;
  wdLigaturesStandardDiscretional = $00000009;
  wdLigaturesContextualDiscretional = $0000000A;
  wdLigaturesHistoricalDiscretional = $0000000C;
  wdLigaturesStandardContextualHistorical = $00000007;
  wdLigaturesStandardContextualDiscretional = $0000000B;
  wdLigaturesStandardHistoricalDiscretional = $0000000D;
  wdLigaturesContextualHistoricalDiscretional = $0000000E;
  wdLigaturesAll = $0000000F;


// Constants for enum WdNumberForm
const
  wdNumberFormDefault = $00000000;
  wdNumberFormLining = $00000001;
  wdNumberFormOldStyle = $00000002;


// Constants for enum WdNumberSpacing
const
  wdNumberSpacingDefault = $00000000;
  wdNumberSpacingProportional = $00000001;
  wdNumberSpacingTabular = $00000002;


// Constants for enum WdStylisticSet
const
  wdStylisticSetDefault = $00000000;
  wdStylisticSet01 = $00000001;
  wdStylisticSet02 = $00000002;
  wdStylisticSet03 = $00000004;
  wdStylisticSet04 = $00000008;
  wdStylisticSet05 = $00000010;
  wdStylisticSet06 = $00000020;
  wdStylisticSet07 = $00000040;
  wdStylisticSet08 = $00000080;
  wdStylisticSet09 = $00000100;
  wdStylisticSet10 = $00000200;
  wdStylisticSet11 = $00000400;
  wdStylisticSet12 = $00000800;
  wdStylisticSet13 = $00001000;
  wdStylisticSet14 = $00002000;
  wdStylisticSet15 = $00004000;
  wdStylisticSet16 = $00008000;
  wdStylisticSet17 = $00010000;
  wdStylisticSet18 = $00020000;
  wdStylisticSet19 = $00040000;
  wdStylisticSet20 = $00080000;


// Constants for enum WdSpanishSpeller
const
  wdSpanishTuteoOnly = $00000000;
  wdSpanishTuteoAndVoseo = $00000001;
  wdSpanishVoseoOnly = $00000002;


// Constants for enum WdLockType
const
  wdLockNone = $00000000;
  wdLockReservation = $00000001;
  wdLockEphemeral = $00000002;
  wdLockChanged = $00000003;


// Constants for enum XlPieSliceLocation
const
  xlHorizontalCoordinate = $00000001;
  xlVerticalCoordinate = $00000002;


// Constants for enum XlPieSliceIndex
const
  xlOuterCounterClockwisePoint = $00000001;
  xlOuterCenterPoint = $00000002;
  xlOuterClockwisePoint = $00000003;
  xlMidClockwiseRadiusPoint = $00000004;
  xlCenterPoint = $00000005;
  xlMidCounterClockwiseRadiusPoint = $00000006;
  xlInnerClockwisePoint = $00000007;
  xlInnerCenterPoint = $00000008;
  xlInnerCounterClockwisePoint = $00000009;


// Constants for enum WdCompatibilityMode
const
  wdWord2003 = $0000000B;
  wdWord2007 = $0000000C;
  wdWord2010 = $0000000E;
  wdWord2013 = $0000000F;
  wdCurrent = $0000FFFF;


// Constants for enum WdProtectedViewCloseReason
const
  wdProtectedViewCloseNormal = $00000000;
  wdProtectedViewCloseEdit = $00000001;
  wdProtectedViewCloseForced = $00000002;


// Constants for enum WdPortugueseReform
const
  wdPortuguesePreReform = $00000001;
  wdPortuguesePostReform = $00000002;
  wdPortugueseBoth = $00000003;


// Constants for enum WdContentControlAppearance
const
  wdContentControlBoundingBox = $00000000;
  wdContentControlTags = $00000001;
  wdContentControlHidden = $00000002;


// Constants for enum WdContentControlLevel
const
  wdContentControlLevelInline = $00000000;
  wdContentControlLevelParagraph = $00000001;
  wdContentControlLevelRow = $00000002;
  wdContentControlLevelCell = $00000003;


// Constants for enum XlCategoryLabelLevel
const
  xlCategoryLabelLevelNone = $FFFFFFFD;
  xlCategoryLabelLevelCustom = $FFFFFFFE;
  xlCategoryLabelLevelAll = $FFFFFFFF;


// Constants for enum XlSeriesNameLevel
const
  xlSeriesNameLevelNone = $FFFFFFFD;
  xlSeriesNameLevelCustom = $FFFFFFFE;
  xlSeriesNameLevelAll = $FFFFFFFF;


// Constants for enum WdPageColor
const
  wdPageColorNone = $00000000;
  wdPageColorSepia = $00000001;
  wdPageColorInverse = $00000002;


// Constants for enum WdColumnWidth
const
  wdColumnWidthNarrow = $00000001;
  wdColumnWidthDefault = $00000002;
  wdColumnWidthWide = $00000003;


// Constants for enum WdRevisionsMarkup
const
  wdRevisionsMarkupNone = $00000000;
  wdRevisionsMarkupSimple = $00000001;
  wdRevisionsMarkupAll = $00000002;


// Constants for enum XlParentDataLabelOptions
const
  xlParentDataLabelOptionsNone = $00000000;
  xlParentDataLabelOptionsBanner = $00000001;
  xlParentDataLabelOptionsOverlapping = $00000002;


// Constants for enum XlBinsType
const
  xlBinsTypeAutomatic = $00000000;
  xlBinsTypeCategorical = $00000001;
  xlBinsTypeManual = $00000002;
  xlBinsTypeBinSize = $00000003;
  xlBinsTypeBinCount = $00000004;


// Constants for enum XlCategorySortOrder
const
  xlIndexAscending = $00000000;
  xlIndexDescending = $00000001;
  xlCategoryAscending = $00000002;
  xlCategoryDescending = $00000003;


// Constants for enum XlValueSortOrder
const
  xlValueNone = $00000000;
  xlValueAscending = $00000001;
  xlValueDescending = $00000002;


// Constants for enum XlGeoProjectionType
const
  xlGeoProjectionTypeAutomatic = $00000000;
  xlGeoProjectionTypeMercator = $00000001;
  xlGeoProjectionTypeMiller = $00000002;
  xlGeoProjectionTypeAlbers = $00000003;
  xlGeoProjectionTypeRobinson = $00000004;


// Constants for enum XlGeoMappingLevel
const
  xlGeoMappingLevelAutomatic = $00000000;
  xlGeoMappingLevelDataOnly = $00000001;
  xlGeoMappingLevelPostalCode = $00000002;
  xlGeoMappingLevelCounty = $00000003;
  xlGeoMappingLevelState = $00000004;
  xlGeoMappingLevelCountryRegion = $00000005;
  xlGeoMappingLevelCountryRegionList = $00000006;
  xlGeoMappingLevelWorld = $00000007;


// Constants for enum XlRegionLabelOptions
const
  xlRegionLabelOptionsNone = $00000000;
  xlRegionLabelOptionsBestFitOnly = $00000001;
  xlRegionLabelOptionsShowAll = $00000002;


// Constants for enum XlSeriesColorGradientStyle
const
  xlSeriesColorGradientStyleSequential = $00000000;
  xlSeriesColorGradientStyleDiverging = $00000001;


// Constants for enum XlGradientStopPositionType
const
  xlGradientStopPositionTypeExtremeValue = $00000000;
  xlGradientStopPositionTypeNumber = $00000001;
  xlGradientStopPositionTypePercent = $00000002;


// Constants for enum WdPageMovementType
const
  wdVertical = $00000001;
  wdSideToSide = $00000002;



implementation

end.

