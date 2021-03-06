unit PowerPoint_TLB_Constants;

// ************************************************************************ //
// WARNING                                                                    
// -------                                                                  
// Type Library, the contents of this file will be regenerated and all        
// manual modifications will be lost.                                         
// ************************************************************************ //

// $Rev: 52393 $
// File generated on 29 Dec 2019 07:49:20 from Type Library described below.

// ************************************************************************  //
// Type Lib: C:\Program Files (x86)\Microsoft Office\Root\Office16\MSPPT.OLB (1)
// LIBID: {91493440-5A91-11CF-8700-00AA0060263B}
// LCID: 0
// Helpfile: C:\Program Files (x86)\Microsoft Office\Root\Office16\VBAPP10.CHM 
// HelpString: Microsoft PowerPoint 16.0 Object Library

interface


// *********************************************************************//
// Declaration of Enumerations defined in Type Library                    
// *********************************************************************//
// Constants for enum PpWindowState
const
  ppWindowNormal = $00000001;
  ppWindowMinimized = $00000002;
  ppWindowMaximized = $00000003;

// Constants for enum PpArrangeStyle
const
  ppArrangeTiled = $00000001;
  ppArrangeCascade = $00000002;

// Constants for enum PpViewType
const
  ppViewSlide = $00000001;
  ppViewSlideMaster = $00000002;
  ppViewNotesPage = $00000003;
  ppViewHandoutMaster = $00000004;
  ppViewNotesMaster = $00000005;
  ppViewOutline = $00000006;
  ppViewSlideSorter = $00000007;
  ppViewTitleMaster = $00000008;
  ppViewNormal = $00000009;
  ppViewPrintPreview = $0000000A;
  ppViewThumbnails = $0000000B;
  ppViewMasterThumbnails = $0000000C;

// Constants for enum PpColorSchemeIndex
const
  ppSchemeColorMixed = $FFFFFFFE;
  ppNotSchemeColor = $00000000;
  ppBackground = $00000001;
  ppForeground = $00000002;
  ppShadow = $00000003;
  ppTitle = $00000004;
  ppFill = $00000005;
  ppAccent1 = $00000006;
  ppAccent2 = $00000007;
  ppAccent3 = $00000008;

// Constants for enum PpSlideSizeType
const
  ppSlideSizeOnScreen = $00000001;
  ppSlideSizeLetterPaper = $00000002;
  ppSlideSizeA4Paper = $00000003;
  ppSlideSize35MM = $00000004;
  ppSlideSizeOverhead = $00000005;
  ppSlideSizeBanner = $00000006;
  ppSlideSizeCustom = $00000007;
  ppSlideSizeLedgerPaper = $00000008;
  ppSlideSizeA3Paper = $00000009;
  ppSlideSizeB4ISOPaper = $0000000A;
  ppSlideSizeB5ISOPaper = $0000000B;
  ppSlideSizeB4JISPaper = $0000000C;
  ppSlideSizeB5JISPaper = $0000000D;
  ppSlideSizeHagakiCard = $0000000E;
  ppSlideSizeOnScreen16x9 = $0000000F;
  ppSlideSizeOnScreen16x10 = $00000010;

// Constants for enum PpSaveAsFileType
const
  ppSaveAsPresentation = $00000001;
  ppSaveAsPowerPoint7 = $00000002;
  ppSaveAsPowerPoint4 = $00000003;
  ppSaveAsPowerPoint3 = $00000004;
  ppSaveAsTemplate = $00000005;
  ppSaveAsRTF = $00000006;
  ppSaveAsShow = $00000007;
  ppSaveAsAddIn = $00000008;
  ppSaveAsPowerPoint4FarEast = $0000000A;
  ppSaveAsDefault = $0000000B;
  ppSaveAsHTML = $0000000C;
  ppSaveAsHTMLv3 = $0000000D;
  ppSaveAsHTMLDual = $0000000E;
  ppSaveAsMetaFile = $0000000F;
  ppSaveAsGIF = $00000010;
  ppSaveAsJPG = $00000011;
  ppSaveAsPNG = $00000012;
  ppSaveAsBMP = $00000013;
  ppSaveAsWebArchive = $00000014;
  ppSaveAsTIF = $00000015;
  ppSaveAsPresForReview = $00000016;
  ppSaveAsEMF = $00000017;
  ppSaveAsOpenXMLPresentation = $00000018;
  ppSaveAsOpenXMLPresentationMacroEnabled = $00000019;
  ppSaveAsOpenXMLTemplate = $0000001A;
  ppSaveAsOpenXMLTemplateMacroEnabled = $0000001B;
  ppSaveAsOpenXMLShow = $0000001C;
  ppSaveAsOpenXMLShowMacroEnabled = $0000001D;
  ppSaveAsOpenXMLAddin = $0000001E;
  ppSaveAsOpenXMLTheme = $0000001F;
  ppSaveAsPDF = $00000020;
  ppSaveAsXPS = $00000021;
  ppSaveAsXMLPresentation = $00000022;
  ppSaveAsOpenDocumentPresentation = $00000023;
  ppSaveAsOpenXMLPicturePresentation = $00000024;
  ppSaveAsWMV = $00000025;
  ppSaveAsStrictOpenXMLPresentation = $00000026;
  ppSaveAsMP4 = $00000027;
  ppSaveAsAnimatedGIF = $00000028;
  ppSaveAsExternalConverter = $0000FA00;

// Constants for enum PpTextStyleType
const
  ppDefaultStyle = $00000001;
  ppTitleStyle = $00000002;
  ppBodyStyle = $00000003;

// Constants for enum PpSlideLayout
const
  ppLayoutMixed = $FFFFFFFE;
  ppLayoutTitle = $00000001;
  ppLayoutText = $00000002;
  ppLayoutTwoColumnText = $00000003;
  ppLayoutTable = $00000004;
  ppLayoutTextAndChart = $00000005;
  ppLayoutChartAndText = $00000006;
  ppLayoutOrgchart = $00000007;
  ppLayoutChart = $00000008;
  ppLayoutTextAndClipart = $00000009;
  ppLayoutClipartAndText = $0000000A;
  ppLayoutTitleOnly = $0000000B;
  ppLayoutBlank = $0000000C;
  ppLayoutTextAndObject = $0000000D;
  ppLayoutObjectAndText = $0000000E;
  ppLayoutLargeObject = $0000000F;
  ppLayoutObject = $00000010;
  ppLayoutTextAndMediaClip = $00000011;
  ppLayoutMediaClipAndText = $00000012;
  ppLayoutObjectOverText = $00000013;
  ppLayoutTextOverObject = $00000014;
  ppLayoutTextAndTwoObjects = $00000015;
  ppLayoutTwoObjectsAndText = $00000016;
  ppLayoutTwoObjectsOverText = $00000017;
  ppLayoutFourObjects = $00000018;
  ppLayoutVerticalText = $00000019;
  ppLayoutClipArtAndVerticalText = $0000001A;
  ppLayoutVerticalTitleAndText = $0000001B;
  ppLayoutVerticalTitleAndTextOverChart = $0000001C;
  ppLayoutTwoObjects = $0000001D;
  ppLayoutObjectAndTwoObjects = $0000001E;
  ppLayoutTwoObjectsAndObject = $0000001F;
  ppLayoutCustom = $00000020;
  ppLayoutSectionHeader = $00000021;
  ppLayoutComparison = $00000022;
  ppLayoutContentWithCaption = $00000023;
  ppLayoutPictureWithCaption = $00000024;

// Constants for enum PpEntryEffect
const
  ppEffectMixed = $FFFFFFFE;
  ppEffectNone = $00000000;
  ppEffectCut = $00000101;
  ppEffectCutThroughBlack = $00000102;
  ppEffectRandom = $00000201;
  ppEffectBlindsHorizontal = $00000301;
  ppEffectBlindsVertical = $00000302;
  ppEffectCheckerboardAcross = $00000401;
  ppEffectCheckerboardDown = $00000402;
  ppEffectCoverLeft = $00000501;
  ppEffectCoverUp = $00000502;
  ppEffectCoverRight = $00000503;
  ppEffectCoverDown = $00000504;
  ppEffectCoverLeftUp = $00000505;
  ppEffectCoverRightUp = $00000506;
  ppEffectCoverLeftDown = $00000507;
  ppEffectCoverRightDown = $00000508;
  ppEffectDissolve = $00000601;
  ppEffectFade = $00000701;
  ppEffectUncoverLeft = $00000801;
  ppEffectUncoverUp = $00000802;
  ppEffectUncoverRight = $00000803;
  ppEffectUncoverDown = $00000804;
  ppEffectUncoverLeftUp = $00000805;
  ppEffectUncoverRightUp = $00000806;
  ppEffectUncoverLeftDown = $00000807;
  ppEffectUncoverRightDown = $00000808;
  ppEffectRandomBarsHorizontal = $00000901;
  ppEffectRandomBarsVertical = $00000902;
  ppEffectStripsUpLeft = $00000A01;
  ppEffectStripsUpRight = $00000A02;
  ppEffectStripsDownLeft = $00000A03;
  ppEffectStripsDownRight = $00000A04;
  ppEffectStripsLeftUp = $00000A05;
  ppEffectStripsRightUp = $00000A06;
  ppEffectStripsLeftDown = $00000A07;
  ppEffectStripsRightDown = $00000A08;
  ppEffectWipeLeft = $00000B01;
  ppEffectWipeUp = $00000B02;
  ppEffectWipeRight = $00000B03;
  ppEffectWipeDown = $00000B04;
  ppEffectBoxOut = $00000C01;
  ppEffectBoxIn = $00000C02;
  ppEffectFlyFromLeft = $00000D01;
  ppEffectFlyFromTop = $00000D02;
  ppEffectFlyFromRight = $00000D03;
  ppEffectFlyFromBottom = $00000D04;
  ppEffectFlyFromTopLeft = $00000D05;
  ppEffectFlyFromTopRight = $00000D06;
  ppEffectFlyFromBottomLeft = $00000D07;
  ppEffectFlyFromBottomRight = $00000D08;
  ppEffectPeekFromLeft = $00000D09;
  ppEffectPeekFromDown = $00000D0A;
  ppEffectPeekFromRight = $00000D0B;
  ppEffectPeekFromUp = $00000D0C;
  ppEffectCrawlFromLeft = $00000D0D;
  ppEffectCrawlFromUp = $00000D0E;
  ppEffectCrawlFromRight = $00000D0F;
  ppEffectCrawlFromDown = $00000D10;
  ppEffectZoomIn = $00000D11;
  ppEffectZoomInSlightly = $00000D12;
  ppEffectZoomOut = $00000D13;
  ppEffectZoomOutSlightly = $00000D14;
  ppEffectZoomCenter = $00000D15;
  ppEffectZoomBottom = $00000D16;
  ppEffectStretchAcross = $00000D17;
  ppEffectStretchLeft = $00000D18;
  ppEffectStretchUp = $00000D19;
  ppEffectStretchRight = $00000D1A;
  ppEffectStretchDown = $00000D1B;
  ppEffectSwivel = $00000D1C;
  ppEffectSpiral = $00000D1D;
  ppEffectSplitHorizontalOut = $00000E01;
  ppEffectSplitHorizontalIn = $00000E02;
  ppEffectSplitVerticalOut = $00000E03;
  ppEffectSplitVerticalIn = $00000E04;
  ppEffectFlashOnceFast = $00000F01;
  ppEffectFlashOnceMedium = $00000F02;
  ppEffectFlashOnceSlow = $00000F03;
  ppEffectAppear = $00000F04;
  ppEffectCircleOut = $00000F05;
  ppEffectDiamondOut = $00000F06;
  ppEffectCombHorizontal = $00000F07;
  ppEffectCombVertical = $00000F08;
  ppEffectFadeSmoothly = $00000F09;
  ppEffectNewsflash = $00000F0A;
  ppEffectPlusOut = $00000F0B;
  ppEffectPushDown = $00000F0C;
  ppEffectPushLeft = $00000F0D;
  ppEffectPushRight = $00000F0E;
  ppEffectPushUp = $00000F0F;
  ppEffectWedge = $00000F10;
  ppEffectWheel1Spoke = $00000F11;
  ppEffectWheel2Spokes = $00000F12;
  ppEffectWheel3Spokes = $00000F13;
  ppEffectWheel4Spokes = $00000F14;
  ppEffectWheel8Spokes = $00000F15;
  ppEffectWheelReverse1Spoke = $00000F16;
  ppEffectVortexLeft = $00000F17;
  ppEffectVortexUp = $00000F18;
  ppEffectVortexRight = $00000F19;
  ppEffectVortexDown = $00000F1A;
  ppEffectRippleCenter = $00000F1B;
  ppEffectRippleRightUp = $00000F1C;
  ppEffectRippleLeftUp = $00000F1D;
  ppEffectRippleLeftDown = $00000F1E;
  ppEffectRippleRightDown = $00000F1F;
  ppEffectGlitterDiamondLeft = $00000F20;
  ppEffectGlitterDiamondUp = $00000F21;
  ppEffectGlitterDiamondRight = $00000F22;
  ppEffectGlitterDiamondDown = $00000F23;
  ppEffectGlitterHexagonLeft = $00000F24;
  ppEffectGlitterHexagonUp = $00000F25;
  ppEffectGlitterHexagonRight = $00000F26;
  ppEffectGlitterHexagonDown = $00000F27;
  ppEffectGalleryLeft = $00000F28;
  ppEffectGalleryRight = $00000F29;
  ppEffectConveyorLeft = $00000F2A;
  ppEffectConveyorRight = $00000F2B;
  ppEffectDoorsVertical = $00000F2C;
  ppEffectDoorsHorizontal = $00000F2D;
  ppEffectWindowVertical = $00000F2E;
  ppEffectWindowHorizontal = $00000F2F;
  ppEffectWarpIn = $00000F30;
  ppEffectWarpOut = $00000F31;
  ppEffectFlyThroughIn = $00000F32;
  ppEffectFlyThroughOut = $00000F33;
  ppEffectFlyThroughInBounce = $00000F34;
  ppEffectFlyThroughOutBounce = $00000F35;
  ppEffectRevealSmoothLeft = $00000F36;
  ppEffectRevealSmoothRight = $00000F37;
  ppEffectRevealBlackLeft = $00000F38;
  ppEffectRevealBlackRight = $00000F39;
  ppEffectHoneycomb = $00000F3A;
  ppEffectFerrisWheelLeft = $00000F3B;
  ppEffectFerrisWheelRight = $00000F3C;
  ppEffectSwitchLeft = $00000F3D;
  ppEffectSwitchUp = $00000F3E;
  ppEffectSwitchRight = $00000F3F;
  ppEffectSwitchDown = $00000F40;
  ppEffectFlipLeft = $00000F41;
  ppEffectFlipUp = $00000F42;
  ppEffectFlipRight = $00000F43;
  ppEffectFlipDown = $00000F44;
  ppEffectFlashbulb = $00000F45;
  ppEffectShredStripsIn = $00000F46;
  ppEffectShredStripsOut = $00000F47;
  ppEffectShredRectangleIn = $00000F48;
  ppEffectShredRectangleOut = $00000F49;
  ppEffectCubeLeft = $00000F4A;
  ppEffectCubeUp = $00000F4B;
  ppEffectCubeRight = $00000F4C;
  ppEffectCubeDown = $00000F4D;
  ppEffectRotateLeft = $00000F4E;
  ppEffectRotateUp = $00000F4F;
  ppEffectRotateRight = $00000F50;
  ppEffectRotateDown = $00000F51;
  ppEffectBoxLeft = $00000F52;
  ppEffectBoxUp = $00000F53;
  ppEffectBoxRight = $00000F54;
  ppEffectBoxDown = $00000F55;
  ppEffectOrbitLeft = $00000F56;
  ppEffectOrbitUp = $00000F57;
  ppEffectOrbitRight = $00000F58;
  ppEffectOrbitDown = $00000F59;
  ppEffectPanLeft = $00000F5A;
  ppEffectPanUp = $00000F5B;
  ppEffectPanRight = $00000F5C;
  ppEffectPanDown = $00000F5D;
  ppEffectFallOverLeft = $00000F5E;
  ppEffectFallOverRight = $00000F5F;
  ppEffectDrapeLeft = $00000F60;
  ppEffectDrapeRight = $00000F61;
  ppEffectCurtains = $00000F62;
  ppEffectWindLeft = $00000F63;
  ppEffectWindRight = $00000F64;
  ppEffectPrestige = $00000F65;
  ppEffectFracture = $00000F66;
  ppEffectCrush = $00000F67;
  ppEffectPeelOffLeft = $00000F68;
  ppEffectPeelOffRight = $00000F69;
  ppEffectPageCurlSingleLeft = $00000F6A;
  ppEffectPageCurlSingleRight = $00000F6B;
  ppEffectPageCurlDoubleLeft = $00000F6C;
  ppEffectPageCurlDoubleRight = $00000F6D;
  ppEffectAirplaneLeft = $00000F6E;
  ppEffectAirplaneRight = $00000F6F;
  ppEffectOrigamiLeft = $00000F70;
  ppEffectOrigamiRight = $00000F71;
  ppEffectMorphByObject = $00000F72;
  ppEffectMorphByWord = $00000F73;
  ppEffectMorphByChar = $00000F74;

// Constants for enum PpTextLevelEffect
const
  ppAnimateLevelMixed = $FFFFFFFE;
  ppAnimateLevelNone = $00000000;
  ppAnimateByFirstLevel = $00000001;
  ppAnimateBySecondLevel = $00000002;
  ppAnimateByThirdLevel = $00000003;
  ppAnimateByFourthLevel = $00000004;
  ppAnimateByFifthLevel = $00000005;
  ppAnimateByAllLevels = $00000010;

// Constants for enum PpTextUnitEffect
const
  ppAnimateUnitMixed = $FFFFFFFE;
  ppAnimateByParagraph = $00000000;
  ppAnimateByWord = $00000001;
  ppAnimateByCharacter = $00000002;

// Constants for enum PpChartUnitEffect
const
  ppAnimateChartMixed = $FFFFFFFE;
  ppAnimateBySeries = $00000001;
  ppAnimateByCategory = $00000002;
  ppAnimateBySeriesElements = $00000003;
  ppAnimateByCategoryElements = $00000004;
  ppAnimateChartAllAtOnce = $00000005;

// Constants for enum PpAfterEffect
const
  ppAfterEffectMixed = $FFFFFFFE;
  ppAfterEffectNothing = $00000000;
  ppAfterEffectHide = $00000001;
  ppAfterEffectDim = $00000002;
  ppAfterEffectHideOnClick = $00000003;

// Constants for enum PpAdvanceMode
const
  ppAdvanceModeMixed = $FFFFFFFE;
  ppAdvanceOnClick = $00000001;
  ppAdvanceOnTime = $00000002;

// Constants for enum PpSoundEffectType
const
  ppSoundEffectsMixed = $FFFFFFFE;
  ppSoundNone = $00000000;
  ppSoundStopPrevious = $00000001;
  ppSoundFile = $00000002;

// Constants for enum PpFollowColors
const
  ppFollowColorsMixed = $FFFFFFFE;
  ppFollowColorsNone = $00000000;
  ppFollowColorsScheme = $00000001;
  ppFollowColorsTextAndBackground = $00000002;

// Constants for enum PpUpdateOption
const
  ppUpdateOptionMixed = $FFFFFFFE;
  ppUpdateOptionManual = $00000001;
  ppUpdateOptionAutomatic = $00000002;

// Constants for enum PpParagraphAlignment
const
  ppAlignmentMixed = $FFFFFFFE;
  ppAlignLeft = $00000001;
  ppAlignCenter = $00000002;
  ppAlignRight = $00000003;
  ppAlignJustify = $00000004;
  ppAlignDistribute = $00000005;
  ppAlignThaiDistribute = $00000006;
  ppAlignJustifyLow = $00000007;

// Constants for enum PpBaselineAlignment
const
  ppBaselineAlignMixed = $FFFFFFFE;
  ppBaselineAlignBaseline = $00000001;
  ppBaselineAlignTop = $00000002;
  ppBaselineAlignCenter = $00000003;
  ppBaselineAlignFarEast50 = $00000004;
  ppBaselineAlignAuto = $00000005;

// Constants for enum PpTabStopType
const
  ppTabStopMixed = $FFFFFFFE;
  ppTabStopLeft = $00000001;
  ppTabStopCenter = $00000002;
  ppTabStopRight = $00000003;
  ppTabStopDecimal = $00000004;

// Constants for enum PpIndentControl
const
  ppIndentControlMixed = $FFFFFFFE;
  ppIndentReplaceAttr = $00000001;
  ppIndentKeepAttr = $00000002;

// Constants for enum PpChangeCase
const
  ppCaseSentence = $00000001;
  ppCaseLower = $00000002;
  ppCaseUpper = $00000003;
  ppCaseTitle = $00000004;
  ppCaseToggle = $00000005;

// Constants for enum PpSlideShowPointerType
const
  ppSlideShowPointerNone = $00000000;
  ppSlideShowPointerArrow = $00000001;
  ppSlideShowPointerPen = $00000002;
  ppSlideShowPointerAlwaysHidden = $00000003;
  ppSlideShowPointerAutoArrow = $00000004;
  ppSlideShowPointerEraser = $00000005;

// Constants for enum PpSlideShowState
const
  ppSlideShowRunning = $00000001;
  ppSlideShowPaused = $00000002;
  ppSlideShowBlackScreen = $00000003;
  ppSlideShowWhiteScreen = $00000004;
  ppSlideShowDone = $00000005;

// Constants for enum PpSlideShowAdvanceMode
const
  ppSlideShowManualAdvance = $00000001;
  ppSlideShowUseSlideTimings = $00000002;
  ppSlideShowRehearseNewTimings = $00000003;

// Constants for enum PpFileDialogType
const
  ppFileDialogOpen = $00000001;
  ppFileDialogSave = $00000002;

// Constants for enum PpPrintOutputType
const
  ppPrintOutputSlides = $00000001;
  ppPrintOutputTwoSlideHandouts = $00000002;
  ppPrintOutputThreeSlideHandouts = $00000003;
  ppPrintOutputSixSlideHandouts = $00000004;
  ppPrintOutputNotesPages = $00000005;
  ppPrintOutputOutline = $00000006;
  ppPrintOutputBuildSlides = $00000007;
  ppPrintOutputFourSlideHandouts = $00000008;
  ppPrintOutputNineSlideHandouts = $00000009;
  ppPrintOutputOneSlideHandouts = $0000000A;

// Constants for enum PpPrintHandoutOrder
const
  ppPrintHandoutVerticalFirst = $00000001;
  ppPrintHandoutHorizontalFirst = $00000002;

// Constants for enum PpPrintColorType
const
  ppPrintColor = $00000001;
  ppPrintBlackAndWhite = $00000002;
  ppPrintPureBlackAndWhite = $00000003;

// Constants for enum PpSelectionType
const
  ppSelectionNone = $00000000;
  ppSelectionSlides = $00000001;
  ppSelectionShapes = $00000002;
  ppSelectionText = $00000003;

// Constants for enum PpDirection
const
  ppDirectionMixed = $FFFFFFFE;
  ppDirectionLeftToRight = $00000001;
  ppDirectionRightToLeft = $00000002;

// Constants for enum PpDateTimeFormat
const
  ppDateTimeFormatMixed = $FFFFFFFE;
  ppDateTimeMdyy = $00000001;
  ppDateTimeddddMMMMddyyyy = $00000002;
  ppDateTimedMMMMyyyy = $00000003;
  ppDateTimeMMMMdyyyy = $00000004;
  ppDateTimedMMMyy = $00000005;
  ppDateTimeMMMMyy = $00000006;
  ppDateTimeMMyy = $00000007;
  ppDateTimeMMddyyHmm = $00000008;
  ppDateTimeMMddyyhmmAMPM = $00000009;
  ppDateTimeHmm = $0000000A;
  ppDateTimeHmmss = $0000000B;
  ppDateTimehmmAMPM = $0000000C;
  ppDateTimehmmssAMPM = $0000000D;
  ppDateTimeFigureOut = $0000000E;
  ppDateTimeUAQ1 = $0000000F;
  ppDateTimeUAQ2 = $00000010;
  ppDateTimeUAQ3 = $00000011;
  ppDateTimeUAQ4 = $00000012;
  ppDateTimeUAQ5 = $00000013;
  ppDateTimeUAQ6 = $00000014;
  ppDateTimeUAQ7 = $00000015;

// Constants for enum PpTransitionSpeed
const
  ppTransitionSpeedMixed = $FFFFFFFE;
  ppTransitionSpeedSlow = $00000001;
  ppTransitionSpeedMedium = $00000002;
  ppTransitionSpeedFast = $00000003;

// Constants for enum PpMouseActivation
const
  ppMouseClick = $00000001;
  ppMouseOver = $00000002;

// Constants for enum PpActionType
const
  ppActionMixed = $FFFFFFFE;
  ppActionNone = $00000000;
  ppActionNextSlide = $00000001;
  ppActionPreviousSlide = $00000002;
  ppActionFirstSlide = $00000003;
  ppActionLastSlide = $00000004;
  ppActionLastSlideViewed = $00000005;
  ppActionEndShow = $00000006;
  ppActionHyperlink = $00000007;
  ppActionRunMacro = $00000008;
  ppActionRunProgram = $00000009;
  ppActionNamedSlideShow = $0000000A;
  ppActionOLEVerb = $0000000B;
  ppActionPlay = $0000000C;

// Constants for enum PpPlaceholderType
const
  ppPlaceholderMixed = $FFFFFFFE;
  ppPlaceholderTitle = $00000001;
  ppPlaceholderBody = $00000002;
  ppPlaceholderCenterTitle = $00000003;
  ppPlaceholderSubtitle = $00000004;
  ppPlaceholderVerticalTitle = $00000005;
  ppPlaceholderVerticalBody = $00000006;
  ppPlaceholderObject = $00000007;
  ppPlaceholderChart = $00000008;
  ppPlaceholderBitmap = $00000009;
  ppPlaceholderMediaClip = $0000000A;
  ppPlaceholderOrgChart = $0000000B;
  ppPlaceholderTable = $0000000C;
  ppPlaceholderSlideNumber = $0000000D;
  ppPlaceholderHeader = $0000000E;
  ppPlaceholderFooter = $0000000F;
  ppPlaceholderDate = $00000010;
  ppPlaceholderVerticalObject = $00000011;
  ppPlaceholderPicture = $00000012;

// Constants for enum PpSlideShowType
const
  ppShowTypeSpeaker = $00000001;
  ppShowTypeWindow = $00000002;
  ppShowTypeKiosk = $00000003;
  ppShowTypeWindow2 = $00000004;

// Constants for enum PpPrintRangeType
const
  ppPrintAll = $00000001;
  ppPrintSelection = $00000002;
  ppPrintCurrent = $00000003;
  ppPrintSlideRange = $00000004;
  ppPrintNamedSlideShow = $00000005;
  ppPrintSection = $00000006;

// Constants for enum PpAutoSize
const
  ppAutoSizeMixed = $FFFFFFFE;
  ppAutoSizeNone = $00000000;
  ppAutoSizeShapeToFitText = $00000001;

// Constants for enum PpMediaType
const
  ppMediaTypeMixed = $FFFFFFFE;
  ppMediaTypeOther = $00000001;
  ppMediaTypeSound = $00000002;
  ppMediaTypeMovie = $00000003;

// Constants for enum PpSoundFormatType
const
  ppSoundFormatMixed = $FFFFFFFE;
  ppSoundFormatNone = $00000000;
  ppSoundFormatWAV = $00000001;
  ppSoundFormatMIDI = $00000002;
  ppSoundFormatCDAudio = $00000003;

// Constants for enum PpFarEastLineBreakLevel
const
  ppFarEastLineBreakLevelNormal = $00000001;
  ppFarEastLineBreakLevelStrict = $00000002;
  ppFarEastLineBreakLevelCustom = $00000003;

// Constants for enum PpSlideShowRangeType
const
  ppShowAll = $00000001;
  ppShowSlideRange = $00000002;
  ppShowNamedSlideShow = $00000003;

// Constants for enum PpFrameColors
const
  ppFrameColorsBrowserColors = $00000001;
  ppFrameColorsPresentationSchemeTextColor = $00000002;
  ppFrameColorsPresentationSchemeAccentColor = $00000003;
  ppFrameColorsWhiteTextOnBlack = $00000004;
  ppFrameColorsBlackTextOnWhite = $00000005;

// Constants for enum PpBorderType
const
  ppBorderTop = $00000001;
  ppBorderLeft = $00000002;
  ppBorderBottom = $00000003;
  ppBorderRight = $00000004;
  ppBorderDiagonalDown = $00000005;
  ppBorderDiagonalUp = $00000006;

// Constants for enum PpHTMLVersion
const
  ppHTMLv3 = $00000001;
  ppHTMLv4 = $00000002;
  ppHTMLDual = $00000003;
  ppHTMLAutodetect = $00000004;

// Constants for enum PpPublishSourceType
const
  ppPublishAll = $00000001;
  ppPublishSlideRange = $00000002;
  ppPublishNamedSlideShow = $00000003;

// Constants for enum PpBulletType
const
  ppBulletMixed = $FFFFFFFE;
  ppBulletNone = $00000000;
  ppBulletUnnumbered = $00000001;
  ppBulletNumbered = $00000002;
  ppBulletPicture = $00000003;

// Constants for enum PpNumberedBulletStyle
const
  ppBulletStyleMixed = $FFFFFFFE;
  ppBulletAlphaLCPeriod = $00000000;
  ppBulletAlphaUCPeriod = $00000001;
  ppBulletArabicParenRight = $00000002;
  ppBulletArabicPeriod = $00000003;
  ppBulletRomanLCParenBoth = $00000004;
  ppBulletRomanLCParenRight = $00000005;
  ppBulletRomanLCPeriod = $00000006;
  ppBulletRomanUCPeriod = $00000007;
  ppBulletAlphaLCParenBoth = $00000008;
  ppBulletAlphaLCParenRight = $00000009;
  ppBulletAlphaUCParenBoth = $0000000A;
  ppBulletAlphaUCParenRight = $0000000B;
  ppBulletArabicParenBoth = $0000000C;
  ppBulletArabicPlain = $0000000D;
  ppBulletRomanUCParenBoth = $0000000E;
  ppBulletRomanUCParenRight = $0000000F;
  ppBulletSimpChinPlain = $00000010;
  ppBulletSimpChinPeriod = $00000011;
  ppBulletCircleNumDBPlain = $00000012;
  ppBulletCircleNumWDWhitePlain = $00000013;
  ppBulletCircleNumWDBlackPlain = $00000014;
  ppBulletTradChinPlain = $00000015;
  ppBulletTradChinPeriod = $00000016;
  ppBulletArabicAlphaDash = $00000017;
  ppBulletArabicAbjadDash = $00000018;
  ppBulletHebrewAlphaDash = $00000019;
  ppBulletKanjiKoreanPlain = $0000001A;
  ppBulletKanjiKoreanPeriod = $0000001B;
  ppBulletArabicDBPlain = $0000001C;
  ppBulletArabicDBPeriod = $0000001D;
  ppBulletThaiAlphaPeriod = $0000001E;
  ppBulletThaiAlphaParenRight = $0000001F;
  ppBulletThaiAlphaParenBoth = $00000020;
  ppBulletThaiNumPeriod = $00000021;
  ppBulletThaiNumParenRight = $00000022;
  ppBulletThaiNumParenBoth = $00000023;
  ppBulletHindiAlphaPeriod = $00000024;
  ppBulletHindiNumPeriod = $00000025;
  ppBulletKanjiSimpChinDBPeriod = $00000026;
  ppBulletHindiNumParenRight = $00000027;
  ppBulletHindiAlpha1Period = $00000028;

// Constants for enum PpShapeFormat
const
  ppShapeFormatGIF = $00000000;
  ppShapeFormatJPG = $00000001;
  ppShapeFormatPNG = $00000002;
  ppShapeFormatBMP = $00000003;
  ppShapeFormatWMF = $00000004;
  ppShapeFormatEMF = $00000005;

// Constants for enum PpExportMode
const
  ppRelativeToSlide = $00000001;
  ppClipRelativeToSlide = $00000002;
  ppScaleToFit = $00000003;
  ppScaleXY = $00000004;

// Constants for enum PpPasteDataType
const
  ppPasteDefault = $00000000;
  ppPasteBitmap = $00000001;
  ppPasteEnhancedMetafile = $00000002;
  ppPasteMetafilePicture = $00000003;
  ppPasteGIF = $00000004;
  ppPasteJPG = $00000005;
  ppPastePNG = $00000006;
  ppPasteText = $00000007;
  ppPasteHTML = $00000008;
  ppPasteRTF = $00000009;
  ppPasteOLEObject = $0000000A;
  ppPasteShape = $0000000B;
  ppPasteSVG = $0000000C;

// Constants for enum MsoAnimEffect
const
  msoAnimEffectCustom = $00000000;
  msoAnimEffectAppear = $00000001;
  msoAnimEffectFly = $00000002;
  msoAnimEffectBlinds = $00000003;
  msoAnimEffectBox = $00000004;
  msoAnimEffectCheckerboard = $00000005;
  msoAnimEffectCircle = $00000006;
  msoAnimEffectCrawl = $00000007;
  msoAnimEffectDiamond = $00000008;
  msoAnimEffectDissolve = $00000009;
  msoAnimEffectFade = $0000000A;
  msoAnimEffectFlashOnce = $0000000B;
  msoAnimEffectPeek = $0000000C;
  msoAnimEffectPlus = $0000000D;
  msoAnimEffectRandomBars = $0000000E;
  msoAnimEffectSpiral = $0000000F;
  msoAnimEffectSplit = $00000010;
  msoAnimEffectStretch = $00000011;
  msoAnimEffectStrips = $00000012;
  msoAnimEffectSwivel = $00000013;
  msoAnimEffectWedge = $00000014;
  msoAnimEffectWheel = $00000015;
  msoAnimEffectWipe = $00000016;
  msoAnimEffectZoom = $00000017;
  msoAnimEffectRandomEffects = $00000018;
  msoAnimEffectBoomerang = $00000019;
  msoAnimEffectBounce = $0000001A;
  msoAnimEffectColorReveal = $0000001B;
  msoAnimEffectCredits = $0000001C;
  msoAnimEffectEaseIn = $0000001D;
  msoAnimEffectFloat = $0000001E;
  msoAnimEffectGrowAndTurn = $0000001F;
  msoAnimEffectLightSpeed = $00000020;
  msoAnimEffectPinwheel = $00000021;
  msoAnimEffectRiseUp = $00000022;
  msoAnimEffectSwish = $00000023;
  msoAnimEffectThinLine = $00000024;
  msoAnimEffectUnfold = $00000025;
  msoAnimEffectWhip = $00000026;
  msoAnimEffectAscend = $00000027;
  msoAnimEffectCenterRevolve = $00000028;
  msoAnimEffectFadedSwivel = $00000029;
  msoAnimEffectDescend = $0000002A;
  msoAnimEffectSling = $0000002B;
  msoAnimEffectSpinner = $0000002C;
  msoAnimEffectStretchy = $0000002D;
  msoAnimEffectZip = $0000002E;
  msoAnimEffectArcUp = $0000002F;
  msoAnimEffectFadedZoom = $00000030;
  msoAnimEffectGlide = $00000031;
  msoAnimEffectExpand = $00000032;
  msoAnimEffectFlip = $00000033;
  msoAnimEffectShimmer = $00000034;
  msoAnimEffectFold = $00000035;
  msoAnimEffectChangeFillColor = $00000036;
  msoAnimEffectChangeFont = $00000037;
  msoAnimEffectChangeFontColor = $00000038;
  msoAnimEffectChangeFontSize = $00000039;
  msoAnimEffectChangeFontStyle = $0000003A;
  msoAnimEffectGrowShrink = $0000003B;
  msoAnimEffectChangeLineColor = $0000003C;
  msoAnimEffectSpin = $0000003D;
  msoAnimEffectTransparency = $0000003E;
  msoAnimEffectBoldFlash = $0000003F;
  msoAnimEffectBlast = $00000040;
  msoAnimEffectBoldReveal = $00000041;
  msoAnimEffectBrushOnColor = $00000042;
  msoAnimEffectBrushOnUnderline = $00000043;
  msoAnimEffectColorBlend = $00000044;
  msoAnimEffectColorWave = $00000045;
  msoAnimEffectComplementaryColor = $00000046;
  msoAnimEffectComplementaryColor2 = $00000047;
  msoAnimEffectContrastingColor = $00000048;
  msoAnimEffectDarken = $00000049;
  msoAnimEffectDesaturate = $0000004A;
  msoAnimEffectFlashBulb = $0000004B;
  msoAnimEffectFlicker = $0000004C;
  msoAnimEffectGrowWithColor = $0000004D;
  msoAnimEffectLighten = $0000004E;
  msoAnimEffectStyleEmphasis = $0000004F;
  msoAnimEffectTeeter = $00000050;
  msoAnimEffectVerticalGrow = $00000051;
  msoAnimEffectWave = $00000052;
  msoAnimEffectMediaPlay = $00000053;
  msoAnimEffectMediaPause = $00000054;
  msoAnimEffectMediaStop = $00000055;
  msoAnimEffectPathCircle = $00000056;
  msoAnimEffectPathRightTriangle = $00000057;
  msoAnimEffectPathDiamond = $00000058;
  msoAnimEffectPathHexagon = $00000059;
  msoAnimEffectPath5PointStar = $0000005A;
  msoAnimEffectPathCrescentMoon = $0000005B;
  msoAnimEffectPathSquare = $0000005C;
  msoAnimEffectPathTrapezoid = $0000005D;
  msoAnimEffectPathHeart = $0000005E;
  msoAnimEffectPathOctagon = $0000005F;
  msoAnimEffectPath6PointStar = $00000060;
  msoAnimEffectPathFootball = $00000061;
  msoAnimEffectPathEqualTriangle = $00000062;
  msoAnimEffectPathParallelogram = $00000063;
  msoAnimEffectPathPentagon = $00000064;
  msoAnimEffectPath4PointStar = $00000065;
  msoAnimEffectPath8PointStar = $00000066;
  msoAnimEffectPathTeardrop = $00000067;
  msoAnimEffectPathPointyStar = $00000068;
  msoAnimEffectPathCurvedSquare = $00000069;
  msoAnimEffectPathCurvedX = $0000006A;
  msoAnimEffectPathVerticalFigure8 = $0000006B;
  msoAnimEffectPathCurvyStar = $0000006C;
  msoAnimEffectPathLoopdeLoop = $0000006D;
  msoAnimEffectPathBuzzsaw = $0000006E;
  msoAnimEffectPathHorizontalFigure8 = $0000006F;
  msoAnimEffectPathPeanut = $00000070;
  msoAnimEffectPathFigure8Four = $00000071;
  msoAnimEffectPathNeutron = $00000072;
  msoAnimEffectPathSwoosh = $00000073;
  msoAnimEffectPathBean = $00000074;
  msoAnimEffectPathPlus = $00000075;
  msoAnimEffectPathInvertedTriangle = $00000076;
  msoAnimEffectPathInvertedSquare = $00000077;
  msoAnimEffectPathLeft = $00000078;
  msoAnimEffectPathTurnRight = $00000079;
  msoAnimEffectPathArcDown = $0000007A;
  msoAnimEffectPathZigzag = $0000007B;
  msoAnimEffectPathSCurve2 = $0000007C;
  msoAnimEffectPathSineWave = $0000007D;
  msoAnimEffectPathBounceLeft = $0000007E;
  msoAnimEffectPathDown = $0000007F;
  msoAnimEffectPathTurnUp = $00000080;
  msoAnimEffectPathArcUp = $00000081;
  msoAnimEffectPathHeartbeat = $00000082;
  msoAnimEffectPathSpiralRight = $00000083;
  msoAnimEffectPathWave = $00000084;
  msoAnimEffectPathCurvyLeft = $00000085;
  msoAnimEffectPathDiagonalDownRight = $00000086;
  msoAnimEffectPathTurnDown = $00000087;
  msoAnimEffectPathArcLeft = $00000088;
  msoAnimEffectPathFunnel = $00000089;
  msoAnimEffectPathSpring = $0000008A;
  msoAnimEffectPathBounceRight = $0000008B;
  msoAnimEffectPathSpiralLeft = $0000008C;
  msoAnimEffectPathDiagonalUpRight = $0000008D;
  msoAnimEffectPathTurnUpRight = $0000008E;
  msoAnimEffectPathArcRight = $0000008F;
  msoAnimEffectPathSCurve1 = $00000090;
  msoAnimEffectPathDecayingWave = $00000091;
  msoAnimEffectPathCurvyRight = $00000092;
  msoAnimEffectPathStairsDown = $00000093;
  msoAnimEffectPathUp = $00000094;
  msoAnimEffectPathRight = $00000095;
  msoAnimEffectMediaPlayFromBookmark = $00000096;
  msoAnimEffect3DArrive = $00000097;
  msoAnimEffect3DTurntable = $00000098;
  msoAnimEffect3DSwing = $00000099;
  msoAnimEffect3DJumpAndTurn = $0000009A;

// Constants for enum MsoAnimateByLevel
const
  msoAnimateLevelMixed = $FFFFFFFF;
  msoAnimateLevelNone = $00000000;
  msoAnimateTextByAllLevels = $00000001;
  msoAnimateTextByFirstLevel = $00000002;
  msoAnimateTextBySecondLevel = $00000003;
  msoAnimateTextByThirdLevel = $00000004;
  msoAnimateTextByFourthLevel = $00000005;
  msoAnimateTextByFifthLevel = $00000006;
  msoAnimateChartAllAtOnce = $00000007;
  msoAnimateChartByCategory = $00000008;
  msoAnimateChartByCategoryElements = $00000009;
  msoAnimateChartBySeries = $0000000A;
  msoAnimateChartBySeriesElements = $0000000B;
  msoAnimateDiagramAllAtOnce = $0000000C;
  msoAnimateDiagramDepthByNode = $0000000D;
  msoAnimateDiagramDepthByBranch = $0000000E;
  msoAnimateDiagramBreadthByNode = $0000000F;
  msoAnimateDiagramBreadthByLevel = $00000010;
  msoAnimateDiagramClockwise = $00000011;
  msoAnimateDiagramClockwiseIn = $00000012;
  msoAnimateDiagramClockwiseOut = $00000013;
  msoAnimateDiagramCounterClockwise = $00000014;
  msoAnimateDiagramCounterClockwiseIn = $00000015;
  msoAnimateDiagramCounterClockwiseOut = $00000016;
  msoAnimateDiagramInByRing = $00000017;
  msoAnimateDiagramOutByRing = $00000018;
  msoAnimateDiagramUp = $00000019;
  msoAnimateDiagramDown = $0000001A;

// Constants for enum MsoAnimTriggerType
const
  msoAnimTriggerMixed = $FFFFFFFF;
  msoAnimTriggerNone = $00000000;
  msoAnimTriggerOnPageClick = $00000001;
  msoAnimTriggerWithPrevious = $00000002;
  msoAnimTriggerAfterPrevious = $00000003;
  msoAnimTriggerOnShapeClick = $00000004;
  msoAnimTriggerOnMediaBookmark = $00000005;

// Constants for enum MsoAnimAfterEffect
const
  msoAnimAfterEffectMixed = $FFFFFFFF;
  msoAnimAfterEffectNone = $00000000;
  msoAnimAfterEffectDim = $00000001;
  msoAnimAfterEffectHide = $00000002;
  msoAnimAfterEffectHideOnNextClick = $00000003;

// Constants for enum MsoAnimTextUnitEffect
const
  msoAnimTextUnitEffectMixed = $FFFFFFFF;
  msoAnimTextUnitEffectByParagraph = $00000000;
  msoAnimTextUnitEffectByCharacter = $00000001;
  msoAnimTextUnitEffectByWord = $00000002;

// Constants for enum MsoAnimEffectRestart
const
  msoAnimEffectRestartAlways = $00000001;
  msoAnimEffectRestartWhenOff = $00000002;
  msoAnimEffectRestartNever = $00000003;

// Constants for enum MsoAnimEffectAfter
const
  msoAnimEffectAfterFreeze = $00000001;
  msoAnimEffectAfterRemove = $00000002;
  msoAnimEffectAfterHold = $00000003;
  msoAnimEffectAfterTransition = $00000004;

// Constants for enum MsoAnimDirection
const
  msoAnimDirectionNone = $00000000;
  msoAnimDirectionUp = $00000001;
  msoAnimDirectionRight = $00000002;
  msoAnimDirectionDown = $00000003;
  msoAnimDirectionLeft = $00000004;
  msoAnimDirectionOrdinalMask = $00000005;
  msoAnimDirectionUpLeft = $00000006;
  msoAnimDirectionUpRight = $00000007;
  msoAnimDirectionDownRight = $00000008;
  msoAnimDirectionDownLeft = $00000009;
  msoAnimDirectionTop = $0000000A;
  msoAnimDirectionBottom = $0000000B;
  msoAnimDirectionTopLeft = $0000000C;
  msoAnimDirectionTopRight = $0000000D;
  msoAnimDirectionBottomRight = $0000000E;
  msoAnimDirectionBottomLeft = $0000000F;
  msoAnimDirectionHorizontal = $00000010;
  msoAnimDirectionVertical = $00000011;
  msoAnimDirectionAcross = $00000012;
  msoAnimDirectionIn = $00000013;
  msoAnimDirectionOut = $00000014;
  msoAnimDirectionClockwise = $00000015;
  msoAnimDirectionCounterclockwise = $00000016;
  msoAnimDirectionHorizontalIn = $00000017;
  msoAnimDirectionHorizontalOut = $00000018;
  msoAnimDirectionVerticalIn = $00000019;
  msoAnimDirectionVerticalOut = $0000001A;
  msoAnimDirectionSlightly = $0000001B;
  msoAnimDirectionCenter = $0000001C;
  msoAnimDirectionInSlightly = $0000001D;
  msoAnimDirectionInCenter = $0000001E;
  msoAnimDirectionInBottom = $0000001F;
  msoAnimDirectionOutSlightly = $00000020;
  msoAnimDirectionOutCenter = $00000021;
  msoAnimDirectionOutBottom = $00000022;
  msoAnimDirectionFontBold = $00000023;
  msoAnimDirectionFontItalic = $00000024;
  msoAnimDirectionFontUnderline = $00000025;
  msoAnimDirectionFontStrikethrough = $00000026;
  msoAnimDirectionFontShadow = $00000027;
  msoAnimDirectionFontAllCaps = $00000028;
  msoAnimDirectionInstant = $00000029;
  msoAnimDirectionGradual = $0000002A;
  msoAnimDirectionCycleClockwise = $0000002B;
  msoAnimDirectionCycleCounterclockwise = $0000002C;

// Constants for enum MsoAnimType
const
  msoAnimTypeMixed = $FFFFFFFE;
  msoAnimTypeNone = $00000000;
  msoAnimTypeMotion = $00000001;
  msoAnimTypeColor = $00000002;
  msoAnimTypeScale = $00000003;
  msoAnimTypeRotation = $00000004;
  msoAnimTypeProperty = $00000005;
  msoAnimTypeCommand = $00000006;
  msoAnimTypeFilter = $00000007;
  msoAnimTypeSet = $00000008;

// Constants for enum MsoAnimAdditive
const
  msoAnimAdditiveAddBase = $00000001;
  msoAnimAdditiveAddSum = $00000002;

// Constants for enum MsoAnimAccumulate
const
  msoAnimAccumulateNone = $00000001;
  msoAnimAccumulateAlways = $00000002;

// Constants for enum MsoAnimProperty
const
  msoAnimNone = $00000000;
  msoAnimX = $00000001;
  msoAnimY = $00000002;
  msoAnimWidth = $00000003;
  msoAnimHeight = $00000004;
  msoAnimOpacity = $00000005;
  msoAnimRotation = $00000006;
  msoAnimColor = $00000007;
  msoAnimVisibility = $00000008;
  msoAnimTextFontBold = $00000064;
  msoAnimTextFontColor = $00000065;
  msoAnimTextFontEmboss = $00000066;
  msoAnimTextFontItalic = $00000067;
  msoAnimTextFontName = $00000068;
  msoAnimTextFontShadow = $00000069;
  msoAnimTextFontSize = $0000006A;
  msoAnimTextFontSubscript = $0000006B;
  msoAnimTextFontSuperscript = $0000006C;
  msoAnimTextFontUnderline = $0000006D;
  msoAnimTextFontStrikeThrough = $0000006E;
  msoAnimTextBulletCharacter = $0000006F;
  msoAnimTextBulletFontName = $00000070;
  msoAnimTextBulletNumber = $00000071;
  msoAnimTextBulletColor = $00000072;
  msoAnimTextBulletRelativeSize = $00000073;
  msoAnimTextBulletStyle = $00000074;
  msoAnimTextBulletType = $00000075;
  msoAnimShapePictureContrast = $000003E8;
  msoAnimShapePictureBrightness = $000003E9;
  msoAnimShapePictureGamma = $000003EA;
  msoAnimShapePictureGrayscale = $000003EB;
  msoAnimShapeFillOn = $000003EC;
  msoAnimShapeFillColor = $000003ED;
  msoAnimShapeFillOpacity = $000003EE;
  msoAnimShapeFillBackColor = $000003EF;
  msoAnimShapeLineOn = $000003F0;
  msoAnimShapeLineColor = $000003F1;
  msoAnimShapeShadowOn = $000003F2;
  msoAnimShapeShadowType = $000003F3;
  msoAnimShapeShadowColor = $000003F4;
  msoAnimShapeShadowOpacity = $000003F5;
  msoAnimShapeShadowOffsetX = $000003F6;
  msoAnimShapeShadowOffsetY = $000003F7;

// Constants for enum PpAlertLevel
const
  ppAlertsNone = $00000001;
  ppAlertsAll = $00000002;

// Constants for enum PpRevisionInfo
const
  ppRevisionInfoNone = $00000000;
  ppRevisionInfoBaseline = $00000001;
  ppRevisionInfoMerged = $00000002;

// Constants for enum MsoAnimCommandType
const
  msoAnimCommandTypeEvent = $00000000;
  msoAnimCommandTypeCall = $00000001;
  msoAnimCommandTypeVerb = $00000002;

// Constants for enum MsoAnimFilterEffectType
const
  msoAnimFilterEffectTypeNone = $00000000;
  msoAnimFilterEffectTypeBarn = $00000001;
  msoAnimFilterEffectTypeBlinds = $00000002;
  msoAnimFilterEffectTypeBox = $00000003;
  msoAnimFilterEffectTypeCheckerboard = $00000004;
  msoAnimFilterEffectTypeCircle = $00000005;
  msoAnimFilterEffectTypeDiamond = $00000006;
  msoAnimFilterEffectTypeDissolve = $00000007;
  msoAnimFilterEffectTypeFade = $00000008;
  msoAnimFilterEffectTypeImage = $00000009;
  msoAnimFilterEffectTypePixelate = $0000000A;
  msoAnimFilterEffectTypePlus = $0000000B;
  msoAnimFilterEffectTypeRandomBar = $0000000C;
  msoAnimFilterEffectTypeSlide = $0000000D;
  msoAnimFilterEffectTypeStretch = $0000000E;
  msoAnimFilterEffectTypeStrips = $0000000F;
  msoAnimFilterEffectTypeWedge = $00000010;
  msoAnimFilterEffectTypeWheel = $00000011;
  msoAnimFilterEffectTypeWipe = $00000012   ;
// Constants for enum PpRemoveDocInfoType
const
  ppRDIComments = $00000001;
  ppRDIRemovePersonalInformation = $00000004;
  ppRDIDocumentProperties = $00000008;
  ppRDIDocumentWorkspace = $0000000A;
  ppRDIInkAnnotations = $0000000B;
  ppRDIPublishPath = $0000000D;
  ppRDIDocumentServerProperties = $0000000E;
  ppRDIDocumentManagementPolicy = $0000000F;
  ppRDIContentType = $00000010;
  ppRDISlideUpdateInformation = $00000011;
  ppRDIAtMentions = $00000012;
  ppRDIAll = $00000063;

// Constants for enum PpCheckInVersionType
const
  ppCheckInMinorVersion = $00000000;
  ppCheckInMajorVersion = $00000001;
  ppCheckInOverwriteVersion = $00000002;

// Constants for enum MsoClickState
const
  msoClickStateAfterAllAnimations = $FFFFFFFE;
  msoClickStateBeforeAutomaticAnimations = $FFFFFFFF;

// Constants for enum PpFixedFormatType
const
  ppFixedFormatTypeXPS = $00000001;
  ppFixedFormatTypePDF = $00000002;

// Constants for enum PpFixedFormatIntent
const
  ppFixedFormatIntentScreen = $00000001;
  ppFixedFormatIntentPrint = $00000002;

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
  rgbAliceBlue = $00FFF8F0;
  rgbAntiqueWhite = $00D7EBFA;
  rgbAqua = $00FFFF00;
  rgbAquamarine = $00D4FF7F;
  rgbAzure = $00FFFFF0;
  rgbBeige = $00DCF5F5;
  rgbBisque = $00C4E4FF;
  rgbBlack = $00000000;
  rgbBlanchedAlmond = $00CDEBFF;
  rgbBlue = $00FF0000;
  rgbBlueViolet = $00E22B8A;
  rgbBrown = $002A2AA5;
  rgbBurlyWood = $0087B8DE;
  rgbCadetBlue = $00A09E5F;
  rgbChartreuse = $0000FF7F;
  rgbCoral = $00507FFF;
  rgbCornflowerBlue = $00ED9564;
  rgbCornsilk = $00DCF8FF;
  rgbCrimson = $003C14DC;
  rgbDarkBlue = $008B0000;
  rgbDarkCyan = $008B8B00;
  rgbDarkGoldenrod = $000B86B8;
  rgbDarkGreen = $00006400;
  rgbDarkGray = $00A9A9A9;
  rgbDarkGrey = $00A9A9A9;
  rgbDarkKhaki = $006BB7BD;
  rgbDarkMagenta = $008B008B;
  rgbDarkOliveGreen = $002F6B55;
  rgbDarkOrange = $00008CFF;
  rgbDarkOrchid = $00CC3299;
  rgbDarkRed = $0000008B;
  rgbDarkSalmon = $007A96E9;
  rgbDarkSeaGreen = $008FBC8F;
  rgbDarkSlateBlue = $008B3D48;
  rgbDarkSlateGray = $004F4F2F;
  rgbDarkSlateGrey = $004F4F2F;
  rgbDarkTurquoise = $00D1CE00;
  rgbDarkViolet = $00D30094;
  rgbDeepPink = $009314FF;
  rgbDeepSkyBlue = $00FFBF00;
  rgbDimGray = $00696969;
  rgbDimGrey = $00696969;
  rgbDodgerBlue = $00FF901E;
  rgbFireBrick = $002222B2;
  rgbFloralWhite = $00F0FAFF;
  rgbForestGreen = $00228B22;
  rgbFuchsia = $00FF00FF;
  rgbGainsboro = $00DCDCDC;
  rgbGhostWhite = $00FFF8F8;
  rgbGold = $0000D7FF;
  rgbGoldenrod = $0020A5DA;
  rgbGray = $00808080;
  rgbGreen = $00008000;
  rgbGrey = $00808080;
  rgbGreenYellow = $002FFFAD;
  rgbHoneydew = $00F0FFF0;
  rgbHotPink = $00B469FF;
  rgbIndianRed = $005C5CCD;
  rgbIndigo = $0082004B;
  rgbIvory = $00F0FFFF;
  rgbKhaki = $008CE6F0;
  rgbLavender = $00FAE6E6;
  rgbLavenderBlush = $00F5F0FF;
  rgbLawnGreen = $0000FC7C;
  rgbLemonChiffon = $00CDFAFF;
  rgbLightBlue = $00E6D8AD;
  rgbLightCoral = $008080F0;
  rgbLightCyan = $008B8B00;
  rgbLightGoldenrodYellow = $00D2FAFA;
  rgbLightGray = $00D3D3D3;
  rgbLightGreen = $0090EE90;
  rgbLightGrey = $00D3D3D3;
  rgbLightPink = $00C1B6FF;
  rgbLightSalmon = $007AA0FF;
  rgbLightSeaGreen = $00AAB220;
  rgbLightSkyBlue = $00FACE87;
  rgbLightSlateGray = $00998877;
  rgbLightSlateGrey = $00998877;
  rgbLightSteelBlue = $00DEC4B0;
  rgbLightYellow = $00E0FFFF;
  rgbLime = $0000FF00;
  rgbLimeGreen = $0032CD32;
  rgbLinen = $00E6F0FA;
  rgbMaroon = $00000080;
  rgbMediumAquamarine = $00AAFF66;
  rgbMediumBlue = $00CD0000;
  rgbMediumOrchid = $00D355BA;
  rgbMediumPurple = $00DB7093;
  rgbMediumSeaGreen = $0071B33C;
  rgbMediumSlateBlue = $00EE687B;
  rgbMediumSpringGreen = $009AFA00;
  rgbMediumTurquoise = $00CCD148;
  rgbMediumVioletRed = $008515C7;
  rgbMidnightBlue = $00701919;
  rgbMintCream = $00FAFFF5;
  rgbMistyRose = $00E1E4FF;
  rgbMoccasin = $00B5E4FF;
  rgbNavajoWhite = $00ADDEFF;
  rgbNavy = $00800000;
  rgbNavyBlue = $00800000;
  rgbOldLace = $00E6F5FD;
  rgbOlive = $00008080;
  rgbOliveDrab = $00238E6B;
  rgbOrange = $0000A5FF;
  rgbOrangeRed = $000045FF;
  rgbOrchid = $00D670DA;
  rgbPaleGoldenrod = $006BE8EE;
  rgbPaleGreen = $0098FB98;
  rgbPaleTurquoise = $00EEEEAF;
  rgbPaleVioletRed = $009370DB;
  rgbPapayaWhip = $00D5EFFF;
  rgbPeachPuff = $00B9DAFF;
  rgbPeru = $003F85CD;
  rgbPink = $00CBC0FF;
  rgbPlum = $00DDA0DD;
  rgbPowderBlue = $00E6E0B0;
  rgbPurple = $00800080;
  rgbRed = $000000FF;
  rgbRosyBrown = $008F8FBC;
  rgbRoyalBlue = $00E16941;
  rgbSalmon = $007280FA;
  rgbSandyBrown = $0060A4F4;
  rgbSeaGreen = $00578B2E;
  rgbSeashell = $00EEF5FF;
  rgbSienna = $002D52A0;
  rgbSilver = $00C0C0C0;
  rgbSkyBlue = $00EBCE87;
  rgbSlateBlue = $00CD5A6A;
  rgbSlateGray = $00908070;
  rgbSlateGrey = $00908070;
  rgbSnow = $00FAFAFF;
  rgbSpringGreen = $007FFF00;
  rgbSteelBlue = $00B48246;
  rgbTan = $008CB4D2;
  rgbTeal = $00808000;
  rgbThistle = $00D8BFD8;
  rgbTomato = $004763FF;
  rgbTurquoise = $00D0E040;
  rgbYellow = $0000FFFF;
  rgbYellowGreen = $0032CD9A;
  rgbViolet = $00EE82EE;
  rgbWheat = $00B3DEF5;
  rgbWhite = $00FFFFFF;
  rgbWhiteSmoke = $00F5F5F5;

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

// Constants for enum XlAxisCrosses
const
  xlAxisCrossesAutomatic = $FFFFEFF7;
  xlAxisCrossesCustom = $FFFFEFEE;
  xlAxisCrossesMaximum = $00000002;
  xlAxisCrossesMinimum = $00000004;

// Constants for enum XlAxisGroup
const
  xlPrimary = $00000001;
  xlSecondary = $00000002;

// Constants for enum XlAxisType
const
  xlCategory = $00000001;
  xlSeriesAxis = $00000003;
  xlValue = $00000002;

// Constants for enum XlBarShape
const
  xlBox = $00000000;
  xlPyramidToPoint = $00000001;
  xlPyramidToMax = $00000002;
  xlCylinder = $00000003;
  xlConeToPoint = $00000004;
  xlConeToMax = $00000005;

// Constants for enum XlBorderWeight
const
  xlHairline = $00000001;
  xlMedium = $FFFFEFD6;
  xlThick = $00000004;
  xlThin = $00000002;

// Constants for enum XlCategoryType
const
  xlCategoryScale = $00000002;
  xlTimeScale = $00000003;
  xlAutomaticScale = $FFFFEFF7;

// Constants for enum XlChartElementPosition
const
  xlChartElementPositionAutomatic = $FFFFEFF7;
  xlChartElementPositionCustom = $FFFFEFEE;

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

// Constants for enum XlOrientation
const
  xlDownward = $FFFFEFB6;
  xlHorizontal = $FFFFEFE0;
  xlUpward = $FFFFEFB5;
  xlVertical = $FFFFEFBA;

// Constants for enum XlChartPictureType
const
  xlStackScale = $00000003;
  xlStack = $00000002;
  xlStretch = $00000001;

// Constants for enum XlChartSplitType
const
  xlSplitByPosition = $00000001;
  xlSplitByPercentValue = $00000003;
  xlSplitByCustomSplit = $00000004;
  xlSplitByValue = $00000002;

// Constants for enum XlColorIndex
const
  xlColorIndexAutomatic = $FFFFEFF7;
  xlColorIndexNone = $FFFFEFD2;

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

// Constants for enum XlDataLabelsType
const
  xlDataLabelsShowNone = $FFFFEFD2;
  xlDataLabelsShowValue = $00000002;
  xlDataLabelsShowPercent = $00000003;
  xlDataLabelsShowLabel = $00000004;
  xlDataLabelsShowLabelAndPercent = $00000005;
  xlDataLabelsShowBubbleSizes = $00000006;

// Constants for enum XlDisplayBlanksAs
const
  xlInterpolated = $00000003;
  xlNotPlotted = $00000001;
  xlZero = $00000002;

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

// Constants for enum XlEndStyleCap
const
  xlCap = $00000001;
  xlNoCap = $00000002;

// Constants for enum XlErrorBarDirection
const
  xlChartX = $FFFFEFB8;
  xlChartY = $00000001;

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

// Constants for enum XlLegendPosition
const
  xlLegendPositionBottom = $FFFFEFF5;
  xlLegendPositionCorner = $00000002;
  xlLegendPositionLeft = $FFFFEFDD;
  xlLegendPositionRight = $FFFFEFC8;
  xlLegendPositionTop = $FFFFEFC0;
  xlLegendPositionCustom = $FFFFEFBF;

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

// Constants for enum XlPivotFieldOrientation
const
  xlColumnField = $00000002;
  xlDataField = $00000004;
  xlHidden = $00000000;
  xlPageField = $00000003;
  xlRowField = $00000001;

// Constants for enum XlReadingOrder
const
  xlContext = $FFFFEC76;
  xlLTR = $FFFFEC75;
  xlRTL = $FFFFEC74;

// Constants for enum XlRowCol
const
  xlColumns = $00000002;
  xlRows = $00000001;

// Constants for enum XlScaleType
const
  xlScaleLinear = $FFFFEFDC;
  xlScaleLogarithmic = $FFFFEFDB;

// Constants for enum XlSizeRepresents
const
  xlSizeIsWidth = $00000002;
  xlSizeIsArea = $00000001;

// Constants for enum XlTickLabelOrientation
const
  xlTickLabelOrientationAutomatic = $FFFFEFF7;
  xlTickLabelOrientationDownward = $FFFFEFB6;
  xlTickLabelOrientationHorizontal = $FFFFEFE0;
  xlTickLabelOrientationUpward = $FFFFEFB5;
  xlTickLabelOrientationVertical = $FFFFEFBA;

// Constants for enum XlTickLabelPosition
const
  xlTickLabelPositionHigh = $FFFFEFE1;
  xlTickLabelPositionLow = $FFFFEFDA;
  xlTickLabelPositionNextToAxis = $00000004;
  xlTickLabelPositionNone = $FFFFEFD2;

// Constants for enum XlTickMark
const
  xlTickMarkCross = $00000004;
  xlTickMarkInside = $00000002;
  xlTickMarkNone = $FFFFEFD2;
  xlTickMarkOutside = $00000003;

// Constants for enum XlTimeUnit
const
  xlDays = $00000000;
  xlMonths = $00000001;
  xlYears = $00000002;

// Constants for enum XlTrendlineType
const
  xlExponential = $00000005;
  xlLinear = $FFFFEFDC;
  xlLogarithmic = $FFFFEFDB;
  xlMovingAvg = $00000006;
  xlPolynomial = $00000003;
  xlPower = $00000004;

// Constants for enum XlUnderlineStyle
const
  xlUnderlineStyleDouble = $FFFFEFE9;
  xlUnderlineStyleDoubleAccounting = $00000005;
  xlUnderlineStyleNone = $FFFFEFD2;
  xlUnderlineStyleSingle = $00000002;
  xlUnderlineStyleSingleAccounting = $00000004;

// Constants for enum XlVAlign
const
  xlVAlignBottom = $FFFFEFF5;
  xlVAlignCenter = $FFFFEFF4;
  xlVAlignDistributed = $FFFFEFEB;
  xlVAlignJustify = $FFFFEFDE;
  xlVAlignTop = $FFFFEFC0;

// Constants for enum PpResampleMediaProfile
const
  ppResampleMediaProfileCustom = $00000001;
  ppResampleMediaProfileSmall = $00000002;
  ppResampleMediaProfileSmaller = $00000003;
  ppResampleMediaProfileSmallest = $00000004;

// Constants for enum PpMediaTaskStatus
const
  ppMediaTaskStatusNone = $00000000;
  ppMediaTaskStatusInProgress = $00000001;
  ppMediaTaskStatusQueued = $00000002;
  ppMediaTaskStatusDone = $00000003;
  ppMediaTaskStatusFailed = $00000004;

// Constants for enum PpPlayerState
const
  ppPlaying = $00000000;
  ppPaused = $00000001;
  ppStopped = $00000002;
  ppNotReady = $00000003;

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

// Constants for enum PpProtectedViewCloseReason
const
  ppProtectedViewCloseNormal = $00000000;
  ppProtectedViewCloseEdit = $00000001;
  ppProtectedViewCloseForced = $00000002;

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

// Constants for enum PpGuideOrientation
const
  ppHorizontalGuide = $00000001;
  ppVerticalGuide = $00000002;

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
   implementation

end.
