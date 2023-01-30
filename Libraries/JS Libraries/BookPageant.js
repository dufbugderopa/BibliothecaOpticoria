// global variables  -->

// global (cross-function) variables
//#region
var DirectoryAnchor = false;
var UserAssistLevel = 3;
var NameOfLibrary = "";
var OutputFileName = "";
var DataHasBeenInput = false;
var ColorSchemeInput = "1";
var ColorScheme = "1";
var dataFromCSVfile;
var RequiredDataFields = [];
var CheckCSVFileAnalysisReportText = "";
var TextModifyStartIndex = 0;
var TextModifyEndIndex = 0;
var TextForSpellchecking
var CellForSpellchecking
var ListeningForSelection = true;
var BookmarkList = "";
var SearchList = "";
var data = [];
var meta = [];
var CurrentElementNumberExpandedData = 0;
var CurrentImageDisplayedInExpandedData = 1;
var NumberOfImagesWithEntry = [];
var ImageDirectory =[];
var ImageCompressionFails = "";
var ImageFileList = [];
var ImageProcessingPaused = false;
var ImageProcessingStopped = false;
var ImageProcessingCancelled = false;
var ImageCompressionNumber = 0.25;
var StartImaageProcessingLoop = 1;
var circleProgress =""; 
var NumberOfImagesToProcess = 0;
var ImageCompressionFails = "";
var ImageCompressionFailsFormatted = "";
var ImageSearchResults = '';
var ImagesZip;
var ImageCompressionLevel = 0.15;
var ImageCompressionCollection = [];
var ImageCompressionAlbum = [];
var ImageCompressionAlbumSizes = [];
var ImageNameInCSVfile = [];
var ImageCompressionCollectionEmpty;
var ImageCompressionAlbumEmpty = true;
var MessagesInImageProcessingContainer = '';
var NumberOfFilesCompressed = 0;
var NumberOfFilesFailed = 0;
var ImageCompressionAlbumSaved = false;
var NumberOfPreviewImageDisplayed = 1;
var ProduceSinglePDF = false;
var NumSpanYearIncrements = 18;
var ModalDialogReturn = '';
var ListSpanOfYears = [.1, .25, .5, 1, 5, 10, 25, 50, 75, 100, 250, 500, 750, 1000, 2500, 5000, 7500, 10000];
// alpha spans expressed in fractions of the full Latin alphabet in base 27. Eg: last value is 26 x 27^10. The first 11 characters (0 -> 10) of a name generate its 'namenumber'
var ListSpanOfAlpha = [23756669087844, 39594448479740, 63351117567584, 102945566047324, 205891132094649, 411782264189298, 617673396283947, 1029455660473245,
  1441237924662543, 2676584717230437, 5353169434460874];
// var AlphaUpper = ["Aa", "z Ba","z Ca", "z Da", "z Ea", "z Fa", "z Ga", "z Ha", "z Ia", "z Ja", "z Ka", "z La", "z Ma", "z Na", "z Oa", "z Pa", "z Qa", "z Ra", "z Sa", "z Ta", "z Ua", "z Va", "z Wa", "z Xa", "z Ya", "z Z"];
var AlphaUpper = ["Aa", " Ba"," Ca", " Da", " Ea", " Fa", " Ga", " Ha", " Ia", " Ja", " Ka", " La", " Ma", " Na", " Oa", " Pa", " Qa", " Ra", " Sa", " Ta", " Ua", " Va", " Wa", " Xa", " Ya", " Za"];
var AlphaNavigate = ["A", "B","C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"];
var AlphaLower = ["", "b","c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z"];
  //BookPageants headers assumed in the CSV input file
var BP_Headers =[
        "BP_Author full name",
        "BP_Author last name first",
        "BP_Author AKA",
        "BP_Author Dates",
        "BP_Translator or Editor",
        "BP_Translator",
        "BP_Editor",
        "BP_Short or Common Title",
        "BP_Brief Summary of the Work",
        "BP_Description_1",
        "BP_Description_2",
        "BP_Reference(s)",
        "BP_Title",
        "BP_Title English Translation",
        "BP_Language",
        "BP_Date Published",
        "BP_Date Range",
        "BP_Biographical Statement",
        "BP_Edition Number",
        "BP_First Edition Date",
        "BP_Provenance",
        "BP_Publisher or Seller",
        "BP_Publication Place",
        "BP_Printer",
        "BP_Engraver",
        "BP_Binder",
        "BP_Multi-Volume Set",
        "BP_Nonce-volume",
        "BP_Collation/Pagination/Foliation",
        "BP_Collation",
        "BP_Pagination",
        "BP_Foliation",
        "BP_Plates",
        "BP_BibliographicNotes",
        "BP_Binding",
        "BP_Format and Size",
        "BP_Format",
        "BP_Size",
        "BP_Image1",
        "BP_Caption1",
        "BP_Image2",
        "BP_Caption2",
        "BP_Image3",
        "BP_Caption3",
        "BP_Image4",
        "BP_Caption4",
        "BP_Image5",
        "BP_Caption5",
        "BP_links",
        "BP_Digital Copy link",
        "BP_Number_In_Catalogue",
        "BP_Source",
        "BP_Dust Jacket",
        "BP_ISBN",
        "BP_Location",
        "BP_Latitude",
        "BP_Longitude"
        ];
var BP_NonAlphabeticHeaders =[          
        "BP_Author Dates",
        "BP_Date Published",
        "BP_Date Range",
        "BP_Edition Number",
        "BP_First Edition Date",
        "BP_Multi-Volume Set",
        "BP_Nonce-volume",
        "BP_Image1",
        "BP_Image2",
        "BP_Image3",
        "BP_Image4",
        "BP_Image5",
        "BP_links",
        "BP_Digital Copy link",
        "BP_Number_In_Catalogue",
        ];
var BP_NumericHeaders =[
        "BP_Author Dates",
        "BP_Date Published",
        "BP_Date Range",
        "BP_Edition Number",
        "BP_First Edition Date",
        "BP_Multi-Volume Set",
        "BP_Number_In_Catalogue",
        "BP_ISBN",
        "BP_Latitude",
        "BP_Longitude"        
        ];
var BP_HeadersBibliographic = [20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,51]
var BP_HeadersDescriptive = [9,10,52,53,54]
var BP_HeadersReferences = [11]
var BP_HeadersDigital = [48,49]
var BP_HeadersImages = [38,40,42,44,46]

var DateDetectionStrings = ['?', 'nd', 'n.d.', 'c', 'ca', 'circa', 'fl', '?', '(', ')', '{', '}', '[', ']']
var DateDetectionFailStrings = [ 'unknown', '']

var BookPageantHomeDirectoryHandle;
var BookPageantFileHandle = [];
var BookPageantCSVFileHandle;
var CSVFileHandle = [];
var BookPageantCSVMappingFileHandle;
var BookPageantBPCSVFileHandl = [];
var BookPageantImagesDirectoryHandle;
var BookPageantPDFDirectoryHandle;
var BookPageantPDFCompositionFileHandle;
var BookPageantBPCSVFileHandle = [];
var DirectoriesDatabaseEstablished = false;
var GenericPDFfull = false;
var NumberOfCompositionFields = 0;
var CSVMappingAlreadyBuilt = false;
var CSVFileMapped = "";
var CSVfileHandleKeep;
var PDFfileHandleKeep;
var CSVFileToMap = "";
var CSVFieldReName = [];
var CSVParseResults;
var SpreadsheetCellDataChanged = false;
var SpreadSheetHasChanged = false;
var SpreadsheetHeaders = [];
var untitledColIds = [];
var CurrentCSVDataRow = 1;
var CurrentCSVDataRowMin = 1;
var CurrentCSVDataRowMax = 1000;
var CurrentCSVDataHeader = 1;
var grid  // data grid  for displaying data input from CSV file
var gridCellsWithProblems = [];
var gridCellsWithMissingImages = [];
var CurrentProblemCellNumberInList = 0;
var DisplayCSVDataFormatted = true;
var CurrentCSVRowSpellQuestion = 0;
var CSVDataSpellClean = [];
var CurrentHighlightedText = null;
var CurrentSpreadsheetCellSpecification;
var CurrentSpreadsheetAllCellSpecification;
var ListOfAuthors = [];
var NumberOfAuthorsInList;
var NumberOfHeadersChoosenForCSVChecking = [];
var ListOfAuthorDates = [];
var ListOfTranslatorsOrEditors = [];
var ListOfPubDates = [];
var ListOfPublishers = [];
var ListOfPrinters = [];
var ListOfFormats = [];
var LocationDataPresent = false;
var LocationDataDescription = '';
var MapBackgroundNumber = 1;
var LocationData = [];
var SiteCluster = [];
var CurrentSiteClusterExpanded = [];
var MultipleMarkersGroup = [];
var MinYear = 10000;
var MaxYear = -10000;
var TimelineStartYear = 0;
var MinAlphaNumber = 9007199254740991;
var MaxAlphaNumber = 0;
var NumOfTimelineElements = 0;
var Year = [];
var LastName = [];
var LastNameNumber = [];
var ShortTitle = [];
var InitialScale = 1;
var CurrentTimelinePixelsPerUnit;
var CurrentSpanOfYears;
var CurrentSpanOfAlphabet;
var CurrentLeftPositionOfTimeline;
var TimeLineScaling;
var TimelineElementHeight = [];
var ElementScalePositionXpixels = [];
var ElementScalePositionYpixels = [];
var CurrentElementScalePositionXpixels = [];
var CurrentElementScalePositionYpixels = [];
var CurrentMaximumXPosition = 0;
var CurrentMinimumXPosition = 0;
var ValueYearsSpan = document.getElementById("ValueYearsSpan");
var ValueAlphaSpan = document.getElementById("ValueAlphaSpan");
var InputFileName
var NavigationScalePixelsPerUnit;
var NavigationWindowWidthInPX;
var BookMarks = [];
var NumberOfBookMarks = 0;
var CurrentBookmarkHighlighted = 0;
var LastBookmarkElementHighlighted = 0;
var BookmarkDisplayOn = false;
var SearchDisplayOn = false;
var FilterDisplayOn = false;
var NumberOfElementSearchHits = -1;
var SearchHits = [];
var BPHeadersOfUserData = [];
var CompositionElementMappedToHeaderField = [];
var ContentXML;
var BPHeadersOfUserDataIsAlphabetic = [];
var BPHeadersOfUserDataIsNumeric = [];
var NumberOfSearchableFields
var FieldToSearch = [];
var FieldToSearchIsAlphabetic = [];
var FieldToSearchIsNumeric = [];
var FieldSearchNumber = [];
var SearchValue = [];
var SearchValueInputBoxNumber = [];
var CurrentSearchFieldListNumber;
var FilterHits = [];
var FieldToFilter = [];
var FieldFilterNumber = [];
var FieldToFilterIsAlphabetic = [];
var FieldToFilterIsNumeric = [];
var FilterValue = [];
var FilterValueInputBoxNumber = [];
var FilterHitMap = [];
var NumberOfFilterFieldBoxesChecked = 0;
var NumberOfSearchFieldBoxesChecked = 0;
var DisplayElement = [];
var NumberOfElementsToDisplay;
var ElementSearchHits = [];
var AlphabeticalSortPointer = [];
var DateSortPointer = [];
var AlphabeticalOrder = [];
var ChronologicalOrder = [];
var HeaderForDateSorting;
var HeaderForAlphasorting;
var DateHeaderOpen = false;
var AlphaHeaderOpen = false;
var ChronologicalSort = false;
var AlphabeticalSort = false;
var ElementSizeHasChanged = false;
var ElementComponentsChanged = false;
var ElementCellScale; 
var style = document.querySelector('[data="NavigateWindowSliderWidth"]');
// 27^n, n=0,10
var AlphaDigitValue = [1, 27, 729, 19683, 531441, 14348907, 387420489, 10460353203, 282429536481, 7625597484987, 205891132094649];
var SearchHitMap = [];
//composition variables
    // paper names and sizes
var PaperSizeName = ["Half Letter", "Letter", "Legal", "Junior-Legal", "Ledger-Tabloid", "AO", "A1", "A2", "A3", "A4", "A5", "A6","A7","A8" ];
var PaperSizeWidth = [5.5, 8.5, 8.5, 5.0, 11.0, 84.1, 59.4, 42.0, 29.7, 21.0, 14.8, 10.5, 7.4, 5.2, ];
var PaperSizeLength = [8.5, 11.0, 14.0, 8.0, 17.0, 118.9, 84.1, 59.4, 42.0, 29.7, 21.0, 14.8, 10.5, 7.4 ];
var PaperName = "Letter";
var PaperNumber;
var PaperWidth = 8.5;
var PaperLength = 11.0;
var PaperUnits = 1;
var PageMargins = [];
var ShowMargins = false;
var PageNumber;
var NumberOfPagesDisplayed;
var ShowPaperSizes = false;
var ShowFonts = false;
var ShowPageNumberChoices = false;
var NumberOfCurrentFont =1;
var FontFacesAvailable;
var NamesOfFonts;
var FontNamesAndFaces = [];
var CurrentFont = "Helvetica";
var CurrentFontFace = "normal";
var CurrentFontSize = 10;
var CompositionBoardAlreadyBuilt = false;
var CurrentElementNumberForComposition = 1;
var CSVHeadersMapped = false;
var NumberOfCompositonElements = 0;
var NumberOfOpenFields = 0;
var MapOfCompositionElementsToFieldNumbers = [];
var MapOfCompositionElementsToFieldNumbersCurrentToOriginal = [];
var MapOfCompositionElementsToBPHeaders = [];
var IdOfCompositionElementContentWithFocus = '';
var CurrentCompositionElementInsertionPointX;
var CurrentCompositionElementInsertionPointY;
var CompositionRulersHaveBeenBuilt = false;
var CurrentVerticalSnap = 1.15*(CurrentFontSize/72.0)*PaperUnits;
var CurrentHorizontalSnap = 0.5;
var PixelsPerScaleUnitWidth;
var PixelsPerScaleUnitHeight;
var NumberRangeText;
var CompositionPrintRange2;
var CompositionPrintRange1;
var PDFImageQuality = 0.15;
var ElementPosition = [];
var PDFpages = [];
var FontChangeToggle = [];
var PagePositionTolerance = 0.1;
var Header = {Present: false, Xposition: 0, Yposition: 0, Width: 0, ContentLink: 0, Content: "This is header text", TextAlign: 'center', pageNumberInHeader: false, pageNumber: 0, pageNumberPosition: "", fontName: '', fontStyle: '', fontSize: 10, LineHeightUnits: 0 }
var Footer = {Present: false, Xposition: 0, Yposition: 0, Width: 0, ContentLink: 0, Content: "This is footer text", TextAlign: 'center', pageNumberInFooter: false, pageNumber: 0, pageNumberPosition: "", fontName: '', fontStyle: '', fontSize: 10, LineHeightUnits: 0 }
var HeaderOrFooterLinking = "";
var CurrentHeaderState = 1;
var PageNumberLayoutNumber = 0;
var CompositionElementWaiting = false;
var NewPageTriggered = false;
var FieldNumberThatOpenedSpecialCharForm;
var SpecialCharactersLoaded = false;
var ElementPending = 0;
var HelpButtonsHaveEventListenerAttached = false;
var CSVFileInitialAnalysis = "";
var ShowFullDataBibliographicDetails = true;
var ShowFullDataDescriptions = true;
var ShowFullDataReferences = true;
var ShowFullDataLinks = true;
var ShowSummaryNumberDisplay = 'block';
var ShowSummaryBriefBookTitleDisplay = 'block';
var ShowSummaryAuthorDisplay = 'block';  
var ShowSummaryPublicationYearDisplay = 'block';
var ShowSummaryBriefDescriptionDisplay = 'block';
var TextForinnerHTMLofCSVCheckterButton = "Headers Found in the CSV File<br><span style = 'color:rgb(0, 179, 255);'>Click on the Header for items' Date</span>"
var BPfileHasBeenLoaded = false;
var BPCSVfileHasBeenLoaded = false;
var CompressedImagesfileHasBeenLoaded = false;
var CSVfileHasBeenLoaded = false;
var WorldMapOpened = false;
var SpecialCharacters = [229, 

    0x00b6, 0x2398, 0x2042, 0x2051, 0x2731, 0x2E33, 0x2248, 0x223C, 0xFF5E, 0x2BB0, 0x2BB1,
    0x2BB2, 0x2BB3, 0x2BB4, 0x2BB5, 0x2BB6, 0x2BB7, 0x2B60, 0x2B61, 0x2B62, 0x2B63,
    0x2B64, 0x2B65, 0x2B66, 0x2B67, 0x2B68, 0x2B69, 0x2B50, 0x2B25, 0x2B26, 0x2B27,
    0x2B28, 0x2B29, 0x2B2A, 0x2B2B, 0x2B2C, 0x2B2D, 0x2B2E, 0x2B2F, 0x2500, 0x2501,
    0x2719, 0x271A, 0x271B, 0x271C, 0x271D, 0x271E, 0x271F, 0x2720, 0x2721, 0x2722,
    0x2723, 0x2723, 0x2724, 0x2725, 0x2726, 0x2727, 0x2728, 0x2729, 0x272A, 0x272B,
    0x272C, 0x272D, 0x272E, 0x272F, 0x2730, 0x2731, 0x2732, 0x2733, 0x2734, 0x2735,
    0x2736, 0x2737, 0x2738, 0x2739, 0x2740, 0x2741, 0x2742, 0x2743, 0x2744, 0x2745,
    0x2746, 0x2747, 0x2748, 0x2766, 0x2767, 0x2619, 0x2580, 0x2581, 0x2582, 0x2583,
    0x2584, 0x2585, 0x2586, 0x2587, 0x2588, 0x2589, 0x258A, 0x258B, 0x258C, 0x258D,
    0x258E, 0x258F, 0x2590, 0x2591, 0x2592, 0x2593, 0x2594, 0x2595, 0x2502, 0x2503, 
    
    0x1F650, 0x1F651, 0x1F652, 0x1F653, 0x1F654, 0x1F655, 0x1F656, 0x1F657, 0x1F658, 0x1F659,
    0x1F65A, 0x1F65B, 0x1F65C, 0x1F65D, 0x1F65E, 0x1F65F, 0x1F660, 0x1F661, 0x1F662, 0x1F663,
    0x1F664, 0x1F665, 0x1F666, 0x1F667, 0x1F668, 0x1F669, 0x1F66A, 0x1F66B, 0x1F66C, 0x1F66D,
    0x1F66E, 0x1F66F,

    0x1F780, 0x1F781, 0x1F782, 0x1F783, 0x1F784, 0x1F785, 0x1F786, 0x1F787, 0x1F788, 0x1F789,
    0x1F78A, 0x1F78B, 0x1F78C, 0x1F78D, 0x1F78E, 0x1F78F, 0x1F790, 0x1F791, 0x1F792, 0x1F793,
    0x1F794, 0x1F795, 0x1F796, 0x1F797, 0x1F798, 0x1F799, 0x1F79A, 0x1F79B, 0x1F79C, 0x1F79D,
    0x1F79E, 0x1F79F, 0x1F7A0, 0x1F7A1, 0x1F7A2, 0x1F7A3, 0x1F7A4, 0x1F7A5, 0x1F7A6, 0x1F7A7,
    0x1F7A8, 0x1F7A9, 0x1F7AA, 0x1F7AB, 0x1F7AC, 0x1F7AD, 0x1F7AE, 0x1F7AF, 0x1F7B0, 0x1F7B1,
    0x1F7B2, 0x1F7B3, 0x1F7B4, 0x1F7B5, 0x1F7B6, 0x1F7B7, 0x1F7B8, 0x1F7B9, 0x1F7BA, 0x1F7BB,
    0x1F7BC, 0x1F7BD, 0x1F7BE, 0x1F7BF, 0x1F7C0, 0x1F7C1, 0x1F7C2, 0x1F7C2, 0x1F7C3, 0x1F7C4,
    0x1F7C5, 0x1F7C6, 0x1F7C7, 0x1F7C8, 0x1F7C9, 0x1F7CA, 0x1F7CB, 0x1F7CC, 0x1F7CD, 0x1F7CE,
    0x1F7CF, 0x1F7D0, 0x1F7D1, 0x1F7D2, 0x1F7D3, 0x1F7D4
]

//#endregion

// User Assistance
//#region
function UserAssistance(Message0, Message1, Message2, ImageLink, TimeInSeconds=30){
  if(UserAssistLevel >= 0){
    document.getElementById("UserAssist").style.height = "3.5vh";
    // document.getElementById("UserAssist").style.margin="-.5vw 0 0 0";
    UserAssistElement = document.getElementById("UserAssist");
    if(UserAssistLevel >= 1 && Message0 != ""){ UserAssistElement.innerHTML = "<strong style='color:red'>"+String.fromCharCode(9658)+"</strong>"+"<strong>"+Message0+"</strong>"}
    if(UserAssistLevel >= 1 && Message1 != ""){ UserAssistElement.innerHTML += " "+"<strong>"+Message1+"</strong>"};
    if(UserAssistLevel >= 2 && Message2 != ""){ UserAssistElement.innerHTML += " "+"<strong>"+Message2+"</strong>"};
    if(ImageLink != "")                       { UserAssistElement.innerHTML += " "+"<img class='UserAssistImage' src='BP Appearance/"+ImageLink+"' alt='Some png'</img>"};
    setTimeout(function(){
      UserAssistElement.innerHTML = "";
      document.getElementById("UserAssist").style.height = "3.5vh"; 
      document.getElementById("UserAssist").style.margin="0 0 0 0";
      }, Number(TimeInSeconds)*1000);  
  }
}
async function InitializeOnOpening(){
  // work in this function is done upon the opening of BookPageant. 
  //  1. LocalStorage is queried to see if any user specified display properties have been stored
  //  2. LocalForage is queried to see if any file/directory handles have been stored.
  LocalStoragePresent = localStorage;
  console.log('Local Storage? ')
  console.log(LocalStoragePresent)
  console.log(getComputedStyle(document.documentElement,null).getPropertyValue('--PDFComposeControlRibbonFontSize'));
  BPlocation =window.location.pathname;
  BPlocation = BPlocation.substring(0, BPlocation.lastIndexOf("/")+1);
    if(LocalStoragePresent.length == 0){
    console.log('Local Storage Empty')
    // set all color, type face, and other syle for initial startmup
    // background transparency
    document.getElementById("SummaryDataBackgroundTransparencySlider").value = 0.0;    
    document.documentElement.style.setProperty('--TimeElementBackgroundTransparency', '0.0', 'important');    
    //default background image
    localStorage.setItem('BackgroundType', 'Image');
    document.documentElement.style.setProperty("--BackgroundType","Image");
    ValueOfBrightnessSlider = 0.5;
    localStorage.setItem("BackgroundImageBrightness",ValueOfBrightnessSlider.toString());    
    document.documentElement.style.setProperty("--TimelineBackgroundFade", Number(ValueOfBrightnessSlider));
    let BackgroundElement = document.getElementById("MainTimeline")      
    BackgroundImageFileName = 'background - LargeBookcase OldBooks Gray.jpg';
    BackgroundElement.style.background = 'linear-gradient(rgba(0, 0, 0, '+ Number(ValueOfBrightnessSlider)+'), rgba(0, 0, 0, '+ Number(ValueOfBrightnessSlider)+')), url("./BP Appearance/'+BackgroundImageFileName+'")';
    BackgroundElement.style.backgroundPosition = 'center';
    BackgroundElement.style.backgroundRepeat = 'no-repeat';
    BackgroundElement.style.backgroundSize = 'cover';
    document.getElementById("RadioButtonImageBackground").checked = true;
    document.getElementById("BackgroundColorPicker").disabled = true;
    document.getElementById("BackgroundImagePicker").disabled = false;
    document.getElementById("BackgroundImagePicker").src = './BP Appearance/ImageFile.png'
    document.getElementById("BackgroundColorPicker").src = './BP Appearance/color-palette grayed.png';
    localStorage.setItem("BackgroundImage", BackgroundImageFileName);

    localStorage.setItem("TimeElementBackgroundTransparency","0.5");
    document.documentElement.style.setProperty('--TimeElementBackgroundTransparency', "0.5", 'important');
    document.getElementById("SummaryDataBackgroundTransparencySlider").value = 0.5;
        
    // summary elememnt background color
    localStorage.setItem("TimeElementBackgroundColor",'rgb(200,200,200)');
    document.documentElement.style.setProperty('--TimeElementBackgroundColor', 'rgb(200,200,200)')
    
    // summary element text color
    localStorage.setItem("TimeElementTextColor",'rgb(0,0,0)');
    document.documentElement.style.setProperty('--SummaryTextColor', 'rgb(0,0,0)', 'important')    

    YearSlider = "0.4"
    localStorage.setItem("YearSlider", YearSlider);

    localStorage.setItem("ElementCellScale",'1.0');
    ElementCellScale = Number('1.0');
    document.documentElement.style.setProperty("--ElementCellScale", ElementCellScale)
    document.getElementById("SummaryDataSizeSlider").value = ElementCellScale;
    document.getElementById("ShowSizeSliderValue").innerHTML = ElementCellScale.toString();
    ElementCellScale = 1.0;
  }

  // see if LocalForage has stored the file system anchor
  try{
    BookPageantHomeDirectoryHandle = await localforage.getItem('AnchorDirectoryHandle');
    if(BookPageantHomeDirectoryHandle){
      DirectoryAnchor = true;
      document.getElementById("FilesAndFunctions").style.opacity = "1.00";
      document.getElementById("FilesAndFunctions").setAttribute("title","Open Files, Perform Functions");
      document.getElementById("OpenDropDownMenu").disabled = false;
    }else{
      DirectoryAnchor = false
      document.getElementById("FilesAndFunctions").style.opacity = "0.50";
      document.getElementById("FilesAndFunctions").setAttribute("title","Disabled Until Anchor is Dropped");      
      document.getElementById("OpenDropDownMenu").disabled = true;
      UserAssistance( "Welcome to BookPageant.",
      "First things first: ",
      "Click the Anchor Button to Specifiy the root of BookPageant's file system",
            "",20.0);      
    }
  }catch(err){
    // there is no file anchor. inactivate and gray-out Files & Functions button
    DirectoryAnchor = false
    document.getElementById("FilesAndFunctions").style.opacity = "0.50";
    document.getElementById("FilesAndFunctions").setAttribute("title","Disabled Until Anchor is Dropped");
    document.getElementById("OpenDropDownMenu").disabled = true;
    UserAssistance( "Welcome to BookPageant.",
    "First things first: ",
    "Click the Anchor Button to Specifiy the root of BookPageant's file system",
    "",20.0);
  }

    // if directory anchor has already been picked, dim the anchor button and see if other file/directory handles have been stored
    if(DirectoryAnchor){
      console.log('Directory Anchor Found')
      document.getElementById("Anchor").style.opacity = "0.25";
      localforage.getItem('AnchorDirectoryHandle').then(function(value) {
        BookPageantHomeDirectoryHandle = value;
        BookPageantFileHandle[0] = BookPageantHomeDirectoryHandle;
        CSVFileHandle[0] = BookPageantHomeDirectoryHandle;
        BookPageantCSVFileHandle = BookPageantHomeDirectoryHandle;
        BookPageantCSVMappingFileHandle = BookPageantHomeDirectoryHandle;
        BookPageantImagesDirectoryHandle = BookPageantHomeDirectoryHandle;
        BookPageantPDFDirectoryHandle = BookPageantHomeDirectoryHandle;
        BookPageantPDFCompositionFileHandle = BookPageantHomeDirectoryHandle
        BookPageantBPCSVFileHandle[0] = BookPageantHomeDirectoryHandle;
      });

      localforage.getItem('BookPageantFileHandle').then(function(value) {
        BookPageantFileHandle = value;
      }).catch(function(err) {
        BookPageantFileHandle = BookPageantHomeDirectoryHandle;        
      });
      
      localforage.getItem('BookPageantCSVFileHandle').then(function(value) {
        BookPageantCSVFileHandle = value;
      }).catch(function(err) {
        BookPageantCSVFileHandle = BookPageantHomeDirectoryHandle;        
      });
      
      localforage.getItem('CSVFileHandle').then(function(value) {
        CSVFileHandle = value;
      }).catch(function(err) {
        CSVFileHandle = BookPageantHomeDirectoryHandle;        
      });

      localforage.getItem('BookPageantCSVMappingFileHandle').then(function(value) {
        BookPageantCSVMappingFileHandle = value;
      }).catch(function(err) {
        BookPageantCSVMappingFileHandle = BookPageantHomeDirectoryHandle;        
      });

      localforage.getItem('BookPageantBPCSVFileHandle').then(function(value) {
        BookPageantBPCSVFileHandle = value;
      }).catch(function(err) {
        BookPageantBPCSVFileHandle = BookPageantHomeDirectoryHandle;        
      });      

      localforage.getItem('BookPageantPDFCompositionFileHandle').then(function(value) {
        BookPageantPDFCompositionFileHandle = value;
      }).catch(function(err) {
        BookPageantPDFCompositionFileHandle = BookPageantHomeDirectoryHandle;        
      });

      localforage.getItem('BookPageantImagesDirectoryHandle').then(function(value) {
        BookPageantImagesDirectoryHandle = value;
      }).catch(function(err) {
        BookPageantImagesDirectoryHandle = BookPageantHomeDirectoryHandle;        
      });

      localforage.getItem('CurrentMapBackgroundNumber').then(function(value) {
        CurrentMapBackgroundNumber = value;
      }).catch(function(err) {
        CurrentMapBackgroundNumber = 1;        
      });

    };
    //background type, color, image 
    let BackgroundType = localStorage.getItem('BackgroundType');
    if(BackgroundType==null){BackgroundType = getComputedStyle(document.documentElement,null).getPropertyValue('--BackgroundType')};
    if(BackgroundType=="Image"){
      //recovering background image data
      document.documentElement.style.setProperty("--BackgroundType","Image");
      let ValueOfBrightnessSlider = localStorage.getItem('BackgroundImageBrightness');
      if(ValueOfBrightnessSlider==null){ValueOfBrightnessSlider = getComputedStyle(document.documentElement,null).getPropertyValue('--TimelineBackgroundFade')};
      document.getElementById('SettingsBackgroundImageBrightnessSlider').value = Number(ValueOfBrightnessSlider);
      document.documentElement.style.setProperty("--TimelineBackgroundFade", Number(ValueOfBrightnessSlider));
      let BackgroundElement = document.getElementById("MainTimeline")      
      let BackgroundImageFileName = localStorage.getItem("BackgroundImage");
      //NB: you can't do this: BackgroundElement.style.background = 'linear-gradient(rgba(0, 0, 0, var(--TimelineBackgroundFade)), rgba(0, 0, 0, var(--TimelineBackgroundFade))), url("./BP Appearance/'+BackgroundImageFileName+'")';
      // you can't assign --TimelineBackgroundFade and use --TimelineBackgroundFade in this immediate way
      BackgroundElement.style.background = 'linear-gradient(rgba(0, 0, 0, '+ Number(ValueOfBrightnessSlider)+'), rgba(0, 0, 0, '+ Number(ValueOfBrightnessSlider)+')), url("./BP Appearance/'+BackgroundImageFileName+'")';
      BackgroundElement.style.backgroundPosition = 'center';
      BackgroundElement.style.backgroundRepeat = 'no-repeat';
      BackgroundElement.style.backgroundSize = 'cover';
      document.getElementById("RadioButtonImageBackground").checked = true;
      document.getElementById("BackgroundColorPicker").disabled = true;
      document.getElementById("BackgroundImagePicker").disabled = false;
      document.getElementById("BackgroundImagePicker").src = './BP Appearance/ImageFile.png'
      document.getElementById("BackgroundColorPicker").src = './BP Appearance/color-palette grayed.png';
    }else{
      //recovering background color data
      // document.documentElement.style.cssText = "--BackgroundType: Color";
      document.getElementById("MainTimeline").style.background = "none";
      let BackgroundColor = localStorage.getItem('BackgroundColor');
      if(BackgroundColor==null)
        {BackgroundColor = getComputedStyle(document.documentElement,null).getPropertyValue('--BackgroundColor')};
      document.getElementById("MainTimeline").style.backgroundColor = BackgroundColor;
      document.getElementById("RadioButtonSolidBackground").checked = true;
      document.getElementById("BackgroundColorPicker").disabled = false;
      document.getElementById("BackgroundImagePicker").disabled = true;
      document.getElementById("BackgroundImagePicker").src = './BP Appearance/ImageFile grayed.png';
      document.getElementById("BackgroundColorPicker").src = './BP Appearance/color-palette.png';
      document.documentElement.style.setProperty('--BackgroundColor', BackgroundColor)
    }

   // time element (book entry) background color and transparency
    let TimeElementBackgroundTransparency = localStorage.getItem("TimeElementBackgroundTransparency");
    if(TimeElementBackgroundTransparency == null){TimeElementBackgroundTransparency = '1.0'};
    // if(TimeElementBackgroundTransparency != null){
      // document.documentElement.style.cssText = "--TimeElementBackgroundTransparency: "+TimeElementBackgroundTransparency.toString();
      document.documentElement.style.setProperty('--TimeElementBackgroundTransparency', TimeElementBackgroundTransparency.toString(), 'important');
      document.getElementById("SummaryDataBackgroundTransparencySlider").value = TimeElementBackgroundTransparency;
    // }
    // summary elememnt background color
    let TimeElementBackgroundColor = localStorage.getItem("TimeElementBackgroundColor");
    if(TimeElementBackgroundColor == null){
      TimeElementBackgroundColor = 'rgb(180,180,180)'
    };
    // if(TimeElementBackgroundColor != null){
      document.documentElement.style.setProperty('--TimeElementBackgroundColor', TimeElementBackgroundColor+", var(--TimeElementBackgroundTransparency)", 'important')
    // }
    // summary element text color
    let TextColor = localStorage.getItem("TimeElementTextColor");
    if(TextColor == null){ TextColor = 'rgb(0,0,0)'};
    // if(TextColor != null){
      document.documentElement.style.setProperty('--SummaryTextColor', TextColor, 'important')    
    // }
    // summary element size
    ElementCellScale = localStorage.getItem("ElementCellScale");
    if(ElementCellScale != null){
      ElementCellScale = Number(ElementCellScale);
      document.documentElement.style.setProperty("--ElementCellScale", ElementCellScale)
      document.getElementById("SummaryDataSizeSlider").value = ElementCellScale;
      document.getElementById("ShowSizeSliderValue").innerHTML = ElementCellScale.toString();
    }else{
      ElementCellScale = 1.0;
    }

    // full data element background color
    let FullDataBackgroundColor = localStorage.getItem("FullDataBackgroundColor");
    if(FullDataBackgroundColor != null){
      document.documentElement.style.cssText = "--FullDataBackgroundColor: ", FullDataBackgroundColor.toString();
      document.getElementById("ExpandedElement").style.backgroundColor = FullDataBackgroundColor;
    }    
    // element box corner roundness
    let BorderRoundness = localStorage.getItem("BorderRadiusFraction");
    if(BorderRoundness != null){
      document.documentElement.style.setProperty("--BorderRadiusFraction: ", BorderRoundness.toString());
      document.getElementById("SettingsFrameRoundness").value = BorderRoundness;
    }
    // system font
    let SystemFontNumber = localStorage.getItem("SystemFont");
    if(SystemFontNumber != null){
      document.getElementById("RadioButtonSystemFont"+SystemFontNumber).checked = true; 
      SetSystemFont(Number(SystemFontNumber));
    }else{
      document.getElementById("RadioButtonSystemFont1").checked = true;       
      SetSystemFont(1);
    }
    // color scheme
    let ColorScheme = localStorage.getItem("ColorScheme");
    if(ColorScheme != null){
      document.documentElement.style.setProperty('--ColorScheme', ColorScheme)
    }else{
      document.documentElement.style.setProperty('--ColorScheme', '1');
      ColorScheme = '1';
    }
    SetColorScheme(ColorScheme)

    // time element summary components
    if(localStorage.getItem("ShowSummaryNumberDisplay") != null){
      ShowSummaryNumberDisplay           = localStorage.getItem("ShowSummaryNumberDisplay");  
      ShowSummaryBriefBookTitleDisplay   = localStorage.getItem("ShowSummaryBriefBookTitleDisplay");
      ShowSummaryAuthorDisplay           = localStorage.getItem("ShowSummaryAuthorDisplay");
      ShowSummaryPublicationYearDisplay  = localStorage.getItem("ShowSummaryPublicationYearDisplay");
      ShowSummaryBriefDescriptionDisplay = localStorage.getItem("ShowSummaryBriefDescriptionDisplay");
      // set governing check boxes accordingly
      document.getElementById("SummaryComponent1").checked = (ShowSummaryNumberDisplay == "block")? true : false;
      document.getElementById("SummaryComponent2").checked = (ShowSummaryBriefBookTitleDisplay == "block")? true : false;
      document.getElementById("SummaryComponent3").checked = (ShowSummaryAuthorDisplay == "block")? true : false;
      document.getElementById("SummaryComponent4").checked = (ShowSummaryPublicationYearDisplay == "block")? true : false;
      document.getElementById("SummaryComponent5").checked = (ShowSummaryBriefDescriptionDisplay == "block")? true : false;
    }else{
      ShowSummaryNumberDisplay           = "block";
      ShowSummaryBriefBookTitleDisplay   = "block";
      ShowSummaryAuthorDisplay           = "block";
      ShowSummaryPublicationYearDisplay  = "block";
      ShowSummaryBriefDescriptionDisplay = "block";
    }

    // full data display components
    if(localStorage.getItem("ShowFullDataBibliographicDetails") != null) {
      ShowFullDataBibliographicDetails   = JSON.parse(localStorage.getItem("ShowFullDataBibliographicDetails"));
      ToggleFullDataComponent('BibliographicDetails', ShowFullDataBibliographicDetails)      
      ShowFullDataDescriptions           = JSON.parse(localStorage.getItem("ShowFullDataDescriptions"));
      ToggleFullDataComponent('Descriptions', ShowFullDataDescriptions)
      ShowFullDataReferences             = JSON.parse(localStorage.getItem("ShowFullDataReferences"));
      ToggleFullDataComponent('References', ShowFullDataReferences)
      ShowFullDataLinks                  = JSON.parse(localStorage.getItem("ShowFullDataLinks"));
      ToggleFullDataComponent('Links', ShowFullDataLinks)
      //set governing check boxes accordingly
      document.getElementById("FullDataComponent1").checked = ShowFullDataBibliographicDetails;
      document.getElementById("FullDataComponent2").checked = ShowFullDataDescriptions;
      document.getElementById("FullDataComponent3").checked = ShowFullDataReferences;
      document.getElementById("FullDataComponent4").checked = ShowFullDataLinks;
    }else{
      ShowFullDataBibliographicDetails   = true;
      ShowFullDataDescriptions           = true;
      ShowFullDataReferences             = true;
      ShowFullDataLinks                  = true;
    }

    // image compression level
    let value = localStorage.getItem("ImageCompressionNumber");
    if(ImageCompressionNumber != null){
      ImageCompressionNumber = Number(value);
    }else{
      ImageCompressionNumber = 0.25;
    }

  // }
}
//#endregion

// Show Progress
//#region
  function asyncShowProgress(PercentComplete){
    return new Promise(resolve=>{
      // circleProgress.value = PercentComplete;
      document.getElementById("BarProgressBar").style.width = PercentComplete.toString()+"%"
      setTimeout(resolve(),100);
      // resolve();
    })
  }
  function sleep(ms){
    return new Promise(resolve => setTimeout(resolve, ms));
  }    

//#endregion

// General Settings
//#region
function LoadSettingsWindow(){
  SettingsElement = document.getElementById('GeneralSettings');
  if(SettingsElement.style.visibility == "visible" ){
    SettingsElement.style.visibility = "hidden";
  }else{
    SettingsElement.style.visibility = "Visible";
    dragElement(document.getElementById("GeneralSettingsDrag"));
  }
}
function CloseGeneralSettings(){
  SettingsElement = document.getElementById('GeneralSettings');
  SettingsElement.style.visibility = "hidden";
  if(ElementSizeHasChanged || ElementComponentsChanged){SetSummaryComponents()}
  ElementSizeHasChanged = false;
  ElementComponentsChanged = false;
}

//global
function SetSystemFont(SystemFontNumber=1){
  switch(SystemFontNumber){
    case 1:
      document.documentElement.style.setProperty('--Font' ,"Times-Roman");
      localStorage.setItem("SystemFont", 1);
      break;
    case 2:
      document.documentElement.style.setProperty('--Font' ,"Scala");
      localStorage.setItem("SystemFont", 2);
      break;
    case 3:
      document.documentElement.style.setProperty('--Font' ,"Helvetica");
      localStorage.setItem("SystemFont", 3);
      break;
    case 4:
      document.documentElement.style.setProperty('--Font' ,"Centaur");
      localStorage.setItem("SystemFont", 4);
      break;
    case 5:
      document.documentElement.style.setProperty('--Font' ,"Cambria");
      localStorage.setItem("SystemFont", 5);
      break;
    case 6:
      document.documentElement.style.setProperty('--Font' ,"BookmanOldStyle");
      localStorage.setItem("SystemFont", 6);
      break;
    case 7:
      document.documentElement.style.setProperty('--Font' ,"Baskerville");
      localStorage.setItem("SystemFont", 7);
      break;
    case 8:
      document.documentElement.style.setProperty('--Font' ,"Arial");
      localStorage.setItem("SystemFont", 8);
      break;
     case 9:
       document.documentElement.style.setProperty('--Font' ,"Garamond");
       localStorage.setItem("SystemFont", 9);
       break;      
  }
} 
function FrameRoundness(ValueOfRoundnessSlider){
  document.documentElement.style.setProperty("--BorderRadiusFraction", ValueOfRoundnessSlider.toString());
  localStorage.setItem("BorderRadiusFraction", ValueOfRoundnessSlider.toString());
}
function SetColorScheme(ColorSchemeInput){
  switch (ColorSchemeInput) {
    case '1':
      document.documentElement.style.setProperty('--AlertColor',     'rgb(0, 179, 255)');
      document.documentElement.style.setProperty('--AlertColorDark', 'rgb(1, 91, 129)');
      document.documentElement.style.setProperty('--AlertColorPale', 'rgb(163, 228, 255)');
      document.documentElement.style.setProperty('--Gray0',          'rgb(255, 255, 255)');
      document.documentElement.style.setProperty('--Gray10',         'rgb(211, 211, 211)');
      document.documentElement.style.setProperty('--Gray15',         'rgb(175, 175, 175)');
      document.documentElement.style.setProperty('--Gray20',         'rgb(139, 139, 139)');
      document.documentElement.style.setProperty('--Gray30',         'rgb(110, 110, 100)');
      document.documentElement.style.setProperty('--Gray40',         'rgb(80, 80, 80)');
      document.documentElement.style.setProperty('--Gray50',         'rgb(39, 39, 39)');
      document.documentElement.style.setProperty('--Gray60',         'rgb(10, 10, 10)');
      document.documentElement.style.setProperty('--HueRotate',      '0deg');
      document.documentElement.style.setProperty('--Brightness',     '100%');
      document.getElementById("RadioButtonScreenColor1").checked = true;
      break;
    case '2':
      document.documentElement.style.setProperty('--AlertColor',     'rgb(32, 129, 195)');
      document.documentElement.style.setProperty('--AlertColorDark', 'rgb(32, 129, 195)');
      document.documentElement.style.setProperty('--AlertColorPale', 'rgba(32, 129, 195, 1)');
      document.documentElement.style.setProperty('--Gray0',          'rgba(247, 249, 249, 1)')
      document.documentElement.style.setProperty('--Gray10',         'rgba(219, 233, 231, 1)')
      document.documentElement.style.setProperty('--Gray15',         'rgba(190, 216, 212, 1)')
      document.documentElement.style.setProperty('--Gray20',         'rgba(120, 213, 215, 1)')
      document.documentElement.style.setProperty('--Gray30',         'rgba(99, 210, 255, 1)')
      document.documentElement.style.setProperty('--Gray40',         'rgba(66, 170, 225, 1)')
      document.documentElement.style.setProperty('--Gray50',         'rgba(49, 150, 210, 1)')
      document.documentElement.style.setProperty('--Gray60',         'rgba(32, 129, 195, 1)')
      document.documentElement.style.setProperty('--HueRotate',      '25deg');
      document.documentElement.style.setProperty('--Brightness',     '100%');
      document.getElementById("RadioButtonScreenColor2").checked = true;
      break;
    case '3':
      document.documentElement.style.setProperty('--AlertColor',     'rgb(145, 255, 231)');
      document.documentElement.style.setProperty('--AlertColorDark', 'rgb(66, 121, 109)');
      document.documentElement.style.setProperty('--AlertColorPale', 'rgba(45, 8, 10, 1)');
      document.documentElement.style.setProperty('--Gray0',          'rgba(238, 241, 239, 1)');
      document.documentElement.style.setProperty('--Gray10',         'rgba(250, 255, 253, 1)');
      document.documentElement.style.setProperty('--Gray15',         'rgba(240, 243, 189, 1)');
      document.documentElement.style.setProperty('--Gray20',         'rgb(41, 215, 148)');
      document.documentElement.style.setProperty('--Gray30',         'rgba(2, 195, 154, 1)');
      document.documentElement.style.setProperty('--Gray40',         'rgba(0, 168, 150, 1)');
      document.documentElement.style.setProperty('--Gray50',         'rgb(2, 144, 144)');
      document.documentElement.style.setProperty('--Gray60',         'rgba(5, 102, 141, 1)');
      document.documentElement.style.setProperty('--HueRotate',      '-75deg');
      document.documentElement.style.setProperty('--Brightness',     '100%');
      document.getElementById("RadioButtonScreenColor3").checked = true;
      break;
    case '4':
      document.documentElement.style.setProperty('--AlertColor',     'rgba(45, 8, 10, 1)');
      document.documentElement.style.setProperty('--AlertColorDark', 'rgb(94, 73, 50)');
      document.documentElement.style.setProperty('--AlertColorPale', 'rgba(45, 8, 10, 1)');
      document.documentElement.style.setProperty('--Gray0',          'rgba(187, 216, 179, 1)');
      document.documentElement.style.setProperty('--Gray10',         'rgb(245, 255, 209)');    
      document.documentElement.style.setProperty('--Gray15',         'rgba(162, 159, 21, 1)');
      document.documentElement.style.setProperty('--Gray20',         'rgba(41, 63, 20, 1)');
      document.documentElement.style.setProperty('--Gray30',         'rgba(40, 89, 67, 1)');
      document.documentElement.style.setProperty('--Gray40',         'rgba(84, 106, 118, 1)');
      document.documentElement.style.setProperty('--Gray50',         'rgba(81, 13, 10, 1)');
      document.documentElement.style.setProperty('--Gray60',         'rgba(25, 17, 2, 1)');
      document.documentElement.style.setProperty('--HueRotate',      '180deg');
      document.documentElement.style.setProperty('--Brightness',     '100%');
      document.getElementById("RadioButtonScreenColor4").checked = true;
      break;
    default:
      break;
  }
  localStorage.setItem('ColorScheme', ColorSchemeInput )
}

//background
function TimelineBackgroundImageBrightness(ValueOfBrightnessSlider) {
  document.documentElement.style.setProperty("--TimelineBackgroundFade", Number(ValueOfBrightnessSlider) );
  if(document.getElementById("RadioButtonImageBackground").checked){
    let BackgroundImageFileName = localStorage.getItem('BackgroundImage');
    let BackgroundElement = document.getElementById("MainTimeline");
    // BackgroundElement.style.background = 'linear-gradient(rgba(0, 0, 0, '+ValueOfBrightnessSlider+'), rgba(0, 0, 0, '+ValueOfBrightnessSlider+')), '+BackgroundImageURL;
    BackgroundElement.style.background = 'linear-gradient(rgba(0, 0, 0, var(--TimelineBackgroundFade)), rgba(0, 0, 0, var(--TimelineBackgroundFade))), url("./BP Appearance/'+BackgroundImageFileName+'")';    
    BackgroundElement.style.backgroundPosition = 'center';
    BackgroundElement.style.backgroundRepeat = 'no-repeat';
    BackgroundElement.style.backgroundSize = 'cover';
    localStorage.setItem("BackgroundImageBrightness",ValueOfBrightnessSlider.toString());
  }
}
function TimelineBackgroundType(input='Solid'){
  if(input=="Solid"){
    //background is solid color
    let element = document.getElementById("BackgroundColorPicker");
    element.disabled = false;
    element.src = './BP Appearance/color-palette.png';
    element = document.getElementById("BackgroundImagePicker");
    element.disabled = true;
    element.src = './BP Appearance/ImageFile grayed.png';
    // console.log(element.id)

    localStorage.setItem("BackgroundType", "Solid");
    //turn background image OFF and set color
    document.getElementById("MainTimeline").style.background = "none";
    // let BackgroundColor = getComputedStyle(document.documentElement,null).getPropertyValue('--BackgroundColor');
    let BackgroundColor = localStorage.getItem('BackgroundColor');
    if(BackgroundColor != null){
      document.getElementById("MainTimeline").style.backgroundColor = BackgroundColor;
    }else{
      OpenColorPicker()
    }
  }else{
    //background is image
    element = document.getElementById("BackgroundColorPicker");
    element.disabled = true;
    element.src = './BP Appearance/color-palette grayed.png';
    element = document.getElementById("BackgroundImagePicker");
    element.disabled = false;
    element.src = './BP Appearance/ImageFile.png';

    localStorage.setItem("BackgroundType", "Image");
    //turn background imge ON if one has previously been chosen. else open image picker
    // let BackgroundFade = getComputedStyle(document.documentElement,null).getPropertyValue('--TimelineBackgroundFade');
    let ValueOfBrightnessSlider = localStorage.getItem('BackgroundImageBrightness');
    if(ValueOfBrightnessSlider==null){ValueOfBrightnessSlider = getComputedStyle(document.documentElement,null).getPropertyValue('--TimelineBackgroundFade')};
    document.documentElement.style.setProperty("--TimelineBackgroundFade", Number(ValueOfBrightnessSlider) );    
    let BackgroundElement = document.getElementById("MainTimeline");
    BackgroundImageFileName = localStorage.getItem("BackgroundImage");
    if(BackgroundImageFileName != null){
      BackgroundElement.style.background = 'linear-gradient(rgba(0, 0, 0, var(--TimelineBackgroundFade)), rgba(0, 0, 0, var(--TimelineBackgroundFade))), url("./BP Appearance/'+BackgroundImageFileName+'")';      BackgroundElement.style.backgroundPosition = 'center';
      BackgroundElement.style.backgroundRepeat = 'no-repeat';
      BackgroundElement.style.backgroundSize = 'cover';
    }else{
      OpenImagePicker();
    }
  }
}

function OpenBackgroundColorPicker(){
  // let ColorPicker = new Huebee( ".ColorPicker",{hues:20, hue0:0, shades:7, saturations:5, notation: 'hex'});
  let BackgrndColorPicker = new Huebee( ".BackgroundColorPicker",{hues:20, hue0:0, shades:7, saturations:5, notation: 'hex', setText:false, setBGColor:false});
  BackgrndColorPicker.open();
  BackgrndColorPicker.on( 'change', function(color) {
    r = "0x" + color[1] + color[2];
    g = "0x" + color[3] + color[4];
    b = "0x" + color[5] + color[6];
    let BackgroundColor = "rgb("+ +r + "," + +g + "," + +b + ")"
    document.getElementById("MainTimeline").style.backgroundColor = BackgroundColor;
    document.documentElement.style.setProperty('--BackgroundColor', BackgroundColor)
    localStorage.setItem("BackgroundColor", BackgroundColor.toString());
  })
}
function OpenImagePicker(){
  GetBackgroundImageFile.click();
}
function ImportBackgroundImageFile(FileList){
  let BackgroundImageFileName = FileList[0].name;
  let element = document.getElementById("MainTimeline");
  element.style.background = 'linear-gradient(rgba(0, 0, 0, var(--TimelineBackgroundFade)), rgba(0, 0, 0, var(--TimelineBackgroundFade))), url("./BP Appearance/'+BackgroundImageFileName+'")';
  element.style.backgroundPosition = 'center';
  element.style.backgroundRepeat = 'no-repeat';
  element.style.backgroundSize = 'cover';
  localStorage.setItem("BackgroundImage", BackgroundImageFileName);
}

//timeline summary elements
function ToggleSummaryComponent(ItemToToggle, State){
  //input is string that names the class of the item.  NB State == checked state AFTER change that brought us here
  //get BookPageant style sheet, find the style, and change it
  let DisplayState = 'block';
  if(State==false){DisplayState = 'none'};
  switch(ItemToToggle){
    case "Number":
      ShowSummaryNumberDisplay = DisplayState;
      ElementComponentsChanged = true;
      break;        
    case "BriefBookTitle":
      ShowSummaryBriefBookTitleDisplay = DisplayState;
      ElementComponentsChanged = true;      
      break;
    case "Author":
      ShowSummaryAuthorDisplay = DisplayState;
      ElementComponentsChanged = true;      
      break;
    case "PublicationYear":
        ShowSummaryPublicationYearDisplay = DisplayState;
        ElementComponentsChanged = true;        
      break;
    case "BriefDescription":
      ShowSummaryBriefDescriptionDisplay = DisplayState;
      ElementComponentsChanged = true;      
      break;
  }
  // store settings for summary components display
  localStorage.setItem("ShowSummaryNumberDisplay", ShowSummaryNumberDisplay);
  localStorage.setItem("ShowSummaryBriefBookTitleDisplay", ShowSummaryBriefBookTitleDisplay);
  localStorage.setItem("ShowSummaryAuthorDisplay", ShowSummaryAuthorDisplay);
  localStorage.setItem("ShowSummaryPublicationYearDisplay", ShowSummaryPublicationYearDisplay);
  localStorage.setItem("ShowSummaryBriefDescriptionDisplay", ShowSummaryBriefDescriptionDisplay);  
}

function SetSummaryComponents(){
  //reproduce items in the timeline using the changed size or components style
  if(ChronologicalSort){
    ProcessDataChronologically();
  }else{
    ProcessDataAlphabetically();
  }
  // store settings for summary components display
  localStorage.setItem("ShowSummaryNumberDisplay", ShowSummaryNumberDisplay);
  localStorage.setItem("ShowSummaryBriefBookTitleDisplay", ShowSummaryBriefBookTitleDisplay);
  localStorage.setItem("ShowSummaryAuthorDisplay", ShowSummaryAuthorDisplay);
  localStorage.setItem("ShowSummaryPublicationYearDisplay", ShowSummaryPublicationYearDisplay);
  localStorage.setItem("ShowSummaryBriefDescriptionDisplay", ShowSummaryBriefDescriptionDisplay);
}

function OpenSummaryBackgroundColorPicker(){
  let ColorPicker = new Huebee( ".SummaryDataColorPicker",{hues:20, hue0:0, shades:7, saturations:5, notation: 'hex', setText:false, setBGColor:false});
  ColorPicker.open();
  ColorPicker.on( 'change', function(color) {
    r = "0x" + color[1] + color[2];
    g = "0x" + color[3] + color[4];
    b = "0x" + color[5] + color[6];
    let BackgroundColor = "rgb("+ +r + "," + +g + "," + +b + ", var(--TimeElementBackgroundTransparency) )"
    //get BookPageant style sheet, find the style, and change it
    document.documentElement.style.setProperty('--TimeElementBackgroundColor', BackgroundColor, 'important')    
    localStorage.setItem("TimeElementBackgroundColor", BackgroundColor.toString());
  })
}
function SummaryDataBackgroundTransparency(ValueOfTransparencySlider){
  document.documentElement.style.setProperty("--TimeElementBackgroundTransparency", ValueOfTransparencySlider.toString());
  localStorage.setItem("TimeElementBackgroundTransparency", ValueOfTransparencySlider.toString());
}
function OpenSummaryTextColorPicker(){
  let ColorPicker = new Huebee( ".SummaryDataTextColorPickerPlace",{hues:20, hue0:0, shades:7, saturations:5, notation: 'hex', setText:false, setBGColor:false});
  ColorPicker.open();
  ColorPicker.on( 'change', function(color) {
    r = "0x" + color[1] + color[2];
    g = "0x" + color[3] + color[4];
    b = "0x" + color[5] + color[6];
    let TextColor = "rgb("+ +r + "," + +g + "," + +b + " )"
    document.documentElement.style.setProperty('--SummaryTextColor', TextColor, 'important')    
    localStorage.setItem("TimeElementTextColor", TextColor.toString());
  })
}
function SummaryDataSize(ValueOfSizeSlider){
  document.documentElement.style.setProperty("--ElementCellScale", ValueOfSizeSlider);
  localStorage.setItem("ElementCellScale", ValueOfSizeSlider);
  ElementCellScale = ValueOfSizeSlider;
  ElementSizeHasChanged = true;
  document.getElementById("ShowSizeSliderValue").innerHTML = ElementCellScale.toString();
}


// full data display
function OpenFullDataBackgroundColorPicker(){
  let ColorPicker = new Huebee( ".FullDataColorPickerPlace",{hues:20, hue0:0, shades:7, saturations:5, notation: 'hex', setText:false, setBGColor:false});
  ColorPicker.open();
  ColorPicker.on( 'change', function(color) {
    r = "0x" + color[1] + color[2];
    g = "0x" + color[3] + color[4];
    b = "0x" + color[5] + color[6];
    let BackgroundColor = "rgb("+ +r + "," + +g + "," + +b + ")"
    document.documentElement.style.setProperty('--FullDataBackgroundColor', BackgroundColor, 'important')    
    localStorage.setItem("FullDataBackgroundColor", BackgroundColor.toString());
    document.getElementById("ExpandedElement").style.backgroundColor = BackgroundColor;
  })
}
function ToggleFullDataComponent(ItemToToggle, State){
  DisplayState = 'block';
  if(State==false){DisplayState = 'none';}

  switch(ItemToToggle){
    case 'BibliographicDetails':
      ShowFullDataBibliographicDetails = State;      
      for(let j=0; j<BP_HeadersBibliographic.length; j++){
        document.getElementById(BP_Headers[BP_HeadersBibliographic[j]]).style.display = DisplayState;
      }
      localStorage.setItem("ShowFullDataBibliographicDetails", State);
      break;
    case 'Descriptions':
      ShowFullDataDescriptions = State;      
      for(let j=0; j<BP_HeadersDescriptive.length; j++){
        document.getElementById(BP_Headers[BP_HeadersDescriptive[j]]).style.display = DisplayState;
      }
      localStorage.setItem("ShowFullDataDescriptions", State);      
      break;      
    case 'References':
      ShowFullDataReferences = State;
      for(let j=0; j<BP_HeadersReferences.length; j++){
        document.getElementById(BP_Headers[BP_HeadersReferences[j]]).style.display = DisplayState;
      }
      localStorage.setItem("ShowFullDataReferences", State);      
      break;      
    case 'Links':
      ShowFullDataLinks = State;
      for(let j=0; j<BP_HeadersDigital.length; j++){
        document.getElementById(BP_Headers[BP_HeadersDigital[j]]).style.display = DisplayState;
      }
      localStorage.setItem("ShowFullDataLinks", State);
      break;      
  }
  // see if we have to add space to bring descriptions below image
  DataElement = document.getElementById("BibliographySeparatorBlank");
  VerticalOffset = pxTOvh(DataElement.offsetTop);
  if(VerticalOffset < pxTOvh(570)){
    DataElement.style.paddingTop = (pxTOvh(570)-VerticalOffset).toString()+"vh";
  }else{
    DataElement.style.paddingTop = "0px";
  }
}

function SetFullComponents(){
  // store settings for full data display
  localStorage.setItem("ShowFullDataBibliographicDetails", ShowFullDataBibliographicDetails);
  localStorage.setItem("ShowFullDataDescriptions", ShowFullDataDescriptions);
  localStorage.setItem("ShowFullDataReferences", ShowFullDataReferences);
  localStorage.setItem("ShowFullDataLinks", ShowFullDataLinks);
}

async function ClearStorage(){
  // localforage.length().then(function(numberOfKeys) {
    // Outputs the length of the database.
    // console.log(numberOfKeys);
  // }).catch(function(err) {
    // This code runs if there were any errors
    // console.log(err);
  // });
// 
  // localforage.keys().then(function(keys) {
    // An array of all the key names.
    // console.log(keys);
// }).catch(function(err) {
    // This code runs if there were any errors
    // console.log(err);
// });

await localforage.iterate(function(value, key, iterationNumber) {
  // Resulting key/value pair -- this callback
  // will be executed for every item in the
  // database.
  console.log([key, value]);
}).then(function() {
  console.log('Iteration has completed');
}).catch(function(err) {
  // This code runs if there were any errors
  console.log(err);
});


await localforage.clear().then(function() {
  // Run this code once the database has been entirely deleted.
  console.log('Database is now empty.');
}).catch(function(err) {
  // This code runs if there were any errors
  console.log(err);
});

await localforage.iterate(function(value, key, iterationNumber) {
  // Resulting key/value pair -- this callback
  // will be executed for every item in the
  // database.
  console.log([key, value]);
}).then(function() {
  console.log('Iteration has completed');
}).catch(function(err) {
  // This code runs if there were any errors
  console.log(err);
});

DirectoriesDatabaseEstablished = false;
localStorage.setItem("DirectoriesDatabaseEstablished", DirectoriesDatabaseEstablished.toString());

}

//#endregion

// move/drag elements -->
//#region
function dragElement(elmnt) {
  var pos1 = 0, pos2 = 0, pos3 = 0, pos4 = 0;
    elmnt.onmousedown = dragMouseDown;

  function dragMouseDown(e) {
    e = e || window.event;
    e.preventDefault();
    // get the mouse cursor position at startup:
    pos3 = e.clientX;
    pos4 = e.clientY;
    document.onmouseup = closeDragElement;
    // call a function whenever the cursor moves:
    document.onmousemove = elementDrag;
  }

  function elementDrag(e) {
    e = e || window.event;
    e.preventDefault();
    // calculate the new cursor position:
    pos1 = pos3 - e.clientX;
    pos2 = pos4 - e.clientY;
    pos3 = e.clientX; 
    pos4 = e.clientY;
    // get the element that is the container of dragger. It has the same ID but with suffix "Drag" - remove "Drag" from the ID.
    ContainerToDrag = document.getElementById(elmnt.id.slice(0,elmnt.id.indexOf("Drag")));
    // set the container's new position:  NB: we don't move the drag element since it remains positioned in the container that is dragged.
    ContainerToDrag.style.top =  ContainerToDrag.offsetTop  - pos2+ "px";
    ContainerToDrag.style.left = ContainerToDrag.offsetLeft - pos1+ "px";
  }

  function closeDragElement() {
    /* stop moving when mouse button is released:*/
    document.onmouseup = null;
    document.onmousemove = null;
    //snap
    if( CurrentVerticalSnap != 0 ){
      SnapWidthOfCompositionWindow  = CurrentHorizontalSnap*PixelsPerScaleUnitWidth; 
      SnapHeightOfCompositionWindow = CurrentVerticalSnap*PixelsPerScaleUnitHeight*PaperUnits;
      //snaps may not align with top margin. calc a fraction that forces this vertical alignment
      let NumberOfSnapsToTopMargin = PageMargins[3]*PixelsPerScaleUnitHeight/SnapHeightOfCompositionWindow;
      let FractionOfSnap = NumberOfSnapsToTopMargin-Math.floor(NumberOfSnapsToTopMargin);
      ContainerToDrag.style.left = SnapWidthOfCompositionWindow *Math.round(ContainerToDrag.offsetLeft/SnapWidthOfCompositionWindow) + "px";
      ContainerToDrag.style.top  = SnapHeightOfCompositionWindow*Math.round(ContainerToDrag.offsetTop/SnapHeightOfCompositionWindow) + FractionOfSnap *SnapHeightOfCompositionWindow + "px";
    }

    //position on page of elements has changed, find current positions of elelemtns on composition board. NB: for sorting purposes only, image are assigned a 1st x position of 0.0
    //this forces images to be first in any row of elememnts that contain an image and allows early detection of end-of-page condition.
    for(let m=1; m<=NumberOfCompositonElements; m++){
      ElementPosition[m-1] = [PaperWidth* document.getElementById('CompositionElement'+MapOfCompositionElementsToFieldNumbers[m].toString()).offsetLeft/document.getElementById("CompositionBoardContent").clientWidth,
                              PaperWidth*(document.getElementById('CompositionElement'+MapOfCompositionElementsToFieldNumbers[m].toString()).offsetTop /document.getElementById("CompositionBoardContent").clientHeight)*(document.getElementById("CompositionBoardContent").clientHeight/document.getElementById("CompositionBoardContent").clientWidth),
                              MapOfCompositionElementsToFieldNumbers[m],
                              MapOfCompositionElementsToBPHeaders[m]];
      FieldNumber = MapOfCompositionElementsToFieldNumbers[m];
      if(FieldNumber >= 0){
        // if(document.getElementById('CompositionElementTitle'+FieldNumber.toString() ).innerHTML.trim() == "Images"){
        if(document.getElementById('CompositionElementTitle'+FieldNumber.toString() ).innerHTML.trim() == "Image local Link1"){
          ElementPosition[m-1][0] = 0.0;
          ElementPosition[m-1][1] = ElementPosition[m-1][1] - 0.*SnapHeightOfCompositionWindow; // move image element up 2 line heights, trumps small changes in positiion and roundoff.
        };
      }
    }
    //sort MapOfCompositionElementsToFieldNumbers list in an order appropriate for PDF printing => y increasing, then x increasing. NB:  sort function re-bases arrays to zero
    ElementPosition.sort(function(a, b){
      if( a[1]>b[1] || (a[1]==b[1] && a[0]>b[0] ) ) return 1;
      if( a[1]==b[1] && a[0]==b[0]) return 0;
      return -1;
    })
    for(let m=1; m<=NumberOfCompositonElements; m++){
      MapOfCompositionElementsToFieldNumbers[m] = ElementPosition[m-1][2]
      MapOfCompositionElementsToBPHeaders[m]= ElementPosition[m-1][3]
    }
    NumProcessed = NumberOfCompositonElements;
  }
}
// move/drag CSV field elements
function dragCSVField(elmnt) {
  var pos1 = 0, pos2 = 0, pos3 = 0, pos4 = 0;
    elmnt.onmousedown = CSVdragMouseDown;

  function CSVdragMouseDown(e) {
    e = e || window.event;
    e.preventDefault();
    // get the mouse cursor position at startup:
    pos3 = e.clientX;
    pos4 = e.clientY;
    document.onmouseup = CSVcloseDragElement;
    // call a function whenever the cursor moves:
    document.onmousemove = CSVelementDrag;
  }

  function CSVelementDrag(e) {
    e = e || window.event;
    e.preventDefault();
    // calculate the new cursor position:
    pos1 = pos3 - e.clientX;
    pos2 = pos4 - e.clientY;
    pos3 = e.clientX; 
    pos4 = e.clientY;
    // get the element that is the container of dragger. It has the same ID but with suffix "Drag" - remove "Drag" from the ID.
    FieldToDrag = document.getElementById(elmnt.id);
    // set the container's new position:  NB: we don't move the drag element since it remains positioned in the container that is dragged.
    // let rect = FieldToDrag.getBoundingClientRect();
    // console.log(rect.top, rect.right, rect.bottom, rect.left);

    FieldToDrag.style.transform = "translate(" + (-pos1).toString() + "px," + (-pos2).toString() + "px)";
    // FieldToDrag.style.top = FieldToDrag.style.top - pos2+ "px";
    // FieldToDrag.style.left = FieldToDrag.style.left - pos1+ "px";
  }

  function CSVcloseDragElement() {
    /* stop moving when mouse button is released:*/
    document.onmouseup = null;
    document.onmousemove = null;
    //snap
    // if( CurrentVerticalSnap != 0 ){
    //   SnapWidthOfCompositionWindow  = CurrentHorizontalSnap*PixelsPerScaleUnitWidth; 
    //   SnapHeightOfCompositionWindow = CurrentVerticalSnap*PixelsPerScaleUnitHeight*PaperUnits;
    //   //snaps may not align with top margin. calc a fraction that forces this vertical alignment
    //   // let NumberOfSnapsToTopMargin = PageMargins[3]*PixelsPerScaleUnitHeight/SnapHeightOfCompositionWindow;
    //   // let FractionOfSnap = NumberOfSnapsToTopMargin-Math.floor(NumberOfSnapsToTopMargin);
    //   // FieldToDrag.style.left = SnapWidthOfCompositionWindow *Math.round(FieldToDrag.offsetLeft/SnapWidthOfCompositionWindow) + "px";
    //   // FieldToDrag.style.top  = SnapHeightOfCompositionWindow*Math.round(FieldToDrag.offsetHeight/SnapHeightOfCompositionWindow) + FractionOfSnap *SnapHeightOfCompositionWindow + "px";
    //   FieldToDrag.style.left = SnapWidthOfCompositionWindow + "px";
    //   FieldToDrag.style.top  = SnapHeightOfCompositionWindow+ "px";
    // }
  }
}
//#endregion

// scoll synchronization -->
//#region
  function SyncScroll(ScrollElementID) {
    // couple timeline drag-scroll with timeline marker and navigation window.
    let verticalScrollBarWitdh = 8;
    ScrollIntegerPx = document.getElementById(ScrollElementID).scrollLeft;
    document.getElementById("TimeLineMarker").scrollLeft = ScrollIntegerPx
    // if(ChronologicalSort == "true"){
      CurrentLeftPositionOfTimeline = MinYear + (ScrollIntegerPx) / CurrentTimelinePixelsPerUnit;
      SliderValue = 50 * ((CurrentLeftPositionOfTimeline - MinYear) * NavigationScalePixelsPerUnit) / (screen.width - verticalScrollBarWitdh - NavigationWindowWidthInPX / 2);
      document.getElementById("NavigateWindowSlider").value = SliderValue;
    // }else{
      // CurrentLeftPositionOfTimeline = MinAlphaNumber + (ScrollIntegerPx) / CurrentTimelinePixelsPerUnit;
      // SliderValue = 50 * ((CurrentLeftPositionOfTimeline - MinAlphaNumber) * NavigationScalePixelsPerUnit) / (screen.width - NavigationWindowWidthInPX / 2);
      // document.getElementById("NavigateWindowSlider").value = SliderValue;
    // }
    // console.log(ChronologicalSort. CurrentLeftPositionOfTimeline, MinYear, ScrollIntegerPx, CurrentTimelinePixelsPerUnit)
  }

  function NavigateWindow(ScrollValue) {
    //  couple navigation slider with timeline and timeline marker
    ScrollIntegerPx = (((Number(ScrollValue) / 50) * (screen.width -12 - NavigationWindowWidthInPX / 2)) / NavigationScalePixelsPerUnit) * CurrentTimelinePixelsPerUnit;
    // console.log("NW", screen.width, NavigationWindowWidthInPX, NavigationScalePixelsPerUnit, CurrentTimelinePixelsPerUnit)
    ScrollIntegerPx = Math.min(ScrollIntegerPx, CurrentMaximumXPosition-screen.width)
    ScrollIntegerPx = Math.max(ScrollIntegerPx, CurrentMinimumXPosition)
    document.getElementById("MainTimeline").scrollLeft = ScrollIntegerPx;
    document.getElementById("TimeLineMarker").scrollLeft = ScrollIntegerPx;
    if(ChronologicalSort == true){
      // check on navigation slider limits
      let CurrentLeftPositionOfTimeline = MinYear + (ScrollIntegerPx) / CurrentTimelinePixelsPerUnit;
      SliderValue = 50 * ((CurrentLeftPositionOfTimeline - MinYear) * NavigationScalePixelsPerUnit) / (screen.width - NavigationWindowWidthInPX / 2);
      document.getElementById("NavigateWindowSlider").value = SliderValue;
      localStorage.setItem("TimeLineScrollValue", SliderValue.toString());
    }else{
      let CurrentLeftPositionOfTimeline = MinAlphaNumber + (ScrollIntegerPx) / CurrentTimelinePixelsPerUnit;
      SliderValue = 50 * ((CurrentLeftPositionOfTimeline - MinAlphaNumber) * NavigationScalePixelsPerUnit) / (screen.width - NavigationWindowWidthInPX / 2);
      document.getElementById("NavigateWindowSlider").value = SliderValue;
      localStorage.setItem("TimeLineScrollValue", SliderValue.toString());
    }
    // console.log('NavWin', CurrentLeftPositionOfTimeline, ScrollIntegerPx, CurrentTimelinePixelsPerUnit)
  }
//#endregion

// drop down menu for settings and functions  -->
//#region
  function OpenDropDownMenu(){
    if( document.getElementById("DropDownMenuContent").style.visibility == 'visible'){
      document.getElementById("DropDownMenuContent").style.visibility = 'hidden';
    }else{
      document.getElementById("DropDownMenuContent").style.visibility = 'visible';
    }
    let location = window.location.pathname;
    let homeDirectoryPath = location.substring(0, location.lastIndexOf("/"));
    console.log(homeDirectoryPath)
  }

  /* Close the dropdown if the user clicks outside of it */
  window.onclick = function(event) {
  if ( ! event.target.matches('.OpenDropDownMenu')) {
    document.getElementById("DropDownMenuContent").style.visibility = 'hidden';
  }
}

function Close(){
  Openedwindow = window.open('', '_self');
  Openedwindow.close();
}


//#endregion

// Data acquisition, population, element placement, etc -->
//#region
// find BookPageant home directory and capture/save the directoryHangle
async function GetHandleToBookPageantDirectory(){
  BookPageantHomeDirectoryHandle = await window.showDirectoryPicker({} );
  // gray out anchor button to show that anchor has been chosen
  document.getElementById("Anchor").style.opacity = "0.5";
  DirectoriesDatabaseEstablished = true;
  localforage.setItem('AnchorDirectoryHandle', BookPageantHomeDirectoryHandle)
  .then(function (value) {
    DirectoriesDatabaseEstablished = true;
  }).catch(function(err) {
    DirectoriesDatabaseEstablished = false;
    alert('Failed to establish anchor')
  });
  BookPageantFileHandle[0] = BookPageantHomeDirectoryHandle;
  BookPageantCSVFileHandle = BookPageantHomeDirectoryHandle;
  CSVFileHandle[0] = BookPageantHomeDirectoryHandle;
  BookPageantCSVMappingFileHandle = BookPageantHomeDirectoryHandle;
  BookPageantImagesDirectoryHandle = BookPageantHomeDirectoryHandle;
  BookPageantPDFDirectoryHandle = BookPageantHomeDirectoryHandle;
  BookPageantPDFCompositionFileHandle = BookPageantHomeDirectoryHandle;
  BookPageantBPCSVFileHandle[0] = BookPageantHomeDirectoryHandle;
  document.getElementById("FilesAndFunctions").style.opacity = "1.00";
  document.getElementById("FilesAndFunctions").setAttribute("title","Open Files, Perform Functions");
  document.getElementById("OpenDropDownMenu").disabled = false;      
}

async function ChooseDataFile() {
  UserAssistance( "Navigate to folder '/BookPageant/CSV files', choose a BP-formatted .csv file and Open.",
                  "File name has the following form: (A Library Name).bp.csv",
                  "Note the double extension '.bp.csv'",
                  "CSVFile.png",5.0);
  document.getElementById("OpenDropDownMenu").click();
  try{CSVFileHandle
    if(CSVFileHandle[0]==null){CSVFileHandle[0]==BookPageantHomeDirectoryHandle}
    CSVFileHandle = await window.showOpenFilePicker({startIn: CSVFileHandle[0], types: [{descriiption: 'BP.CSV file', accept:{'text/plain':['.BP.CSV']}}]});
  }catch{
    CSVFileHandle = await window.showOpenFilePicker({startIn: BookPageantHomeDirectoryHandle, types: [{descriiption: 'BP.CSV file', accept:{'text/plain':['.BP.CSV']}}]});
  }
  InputFile = await CSVFileHandle[0].getFile();
  localforage.setItem('CSVFileHandle', CSVFileHandle);
  BPCSVfileHasBeenLoaded = true;
  BPfileHasBeenLoaded = false;
  CompressedImagesfileHasBeenLoaded = false;
  CSVfileHasBeenLoaded = false;
  // we don't know if there are images available, so allow their display in data expansion if they are.
  ImageCompressionCollectionEmpty = false;
  document.getElementById("BP_Image1").style.display = 'block';
  document.getElementById("BP_ImageCaption").style.display = 'block';
  document.getElementById("ExpandedElementDataImageControl").style.display = 'flex';
  GetCVSFileList(InputFile)    
}

 //read contents of CSV file 
  function GetCVSFileList(InputFile) {
    document.body.style.cursor = "wait";
    // InputFile = FileList[0]
		Reader = new FileReader();
		Reader.onload = function(){
      NameOfLibrary = InputFile.name.slice(0,InputFile.name.indexOf(".bp"));
      document.getElementById("SetLibraryName").value = NameOfLibrary.trim();
      document.getElementById("BPlogoAndLibraryName").innerHTML = '<img class="Logo" src="./BP Appearance/BP Icon.png" alt="NotFound">' + "&nbsp;"+"&nbsp;"+"&nbsp;" + NameOfLibrary.trim() + "&nbsp;"+"&nbsp;"+"&nbsp;" + '<img class="Logo" src="./Images/'+NameOfLibrary.trim()+'.JPG" alt="">' ;
      dataFromCSVfile = Reader.result;
      ParseCSVfile(dataFromCSVfile);
    }
    Reader.readAsText(InputFile);
    document.body.style.cursor = "initial";
  }    
  // parse input CSV file and get basic user data
  function ParseCSVfile(dataFromCSVfile){
    document.body.style.cursor = "wait";
    DataHasBeenInput = false;
    Papa.parse(dataFromCSVfile, {
      delimiter: ",",
      header: true,
      download: false,
      dynamicTyping: false,
      skipEmptyLines: true,
      complete: function (results) {
        // console.log(results);
        data = results.data;
        errors = results.errors;
        meta = results.meta;
        if (results.errors.length !== 0) {
          alert(results.errors[0].type + " \n" + results.errors[0].code + " \n" + results.errors[0].message + "\n In Row: " + results.errors[0].row);
        } else {
          NumUserDataFields = 0
          for (m = 0; m < meta.fields.length; m++) {
            NumUserDataFields = NumUserDataFields + 1;
            BPHeadersOfUserData[NumUserDataFields] = meta.fields[m];
            //find which search fields are alphabetic, numeric, or neither
            if( BP_NonAlphabeticHeaders.indexOf(BPHeadersOfUserData[NumUserDataFields]) == -1 ){
              BPHeadersOfUserDataIsAlphabetic[NumUserDataFields] = true;
              BPHeadersOfUserDataIsNumeric[NumUserDataFields] = false;
            }else{
              BPHeadersOfUserDataIsAlphabetic[NumUserDataFields] = false;
              if( BP_NumericHeaders.indexOf(BPHeadersOfUserData[NumUserDataFields]) == -1 ){
                BPHeadersOfUserDataIsNumeric[NumUserDataFields] = false;
              }else{
                BPHeadersOfUserDataIsNumeric[NumUserDataFields] = true;
              }
            }
          }

          // build table of search fields and search values
          //#region          
          table = document.createElement('TABLE');
          table.setAttribute('class', 'SearchTable')
          table.setAttribute('id', 'SearchTable')
          tableBody = document.createElement('TBODY');
          table.appendChild(tableBody);
          NumberOfSearchableFields = 0
          for (let m = 1; m <= NumUserDataFields; m++) {
            if(data[0][BPHeadersOfUserData[m]]==''){continue}
            NumberOfSearchableFields++
            let tr = document.createElement('tr');
            tableBody.appendChild(tr);
            for (let j = 1; j <= 5; j++) {
              td = document.createElement('td');
              if (j == 1) {
                Input = document.createElement('input');
                Input.type = 'checkbox';
                Input.name = 'Search Field';
                Input.id = 'SearchFieldElement'+m.toString();
                Input.title = 'Click to activate search in this data category'
                Input.onchange = ()=>{                                                          //checks whether search field box is checked
                  if( !document.getElementById('SearchFieldElement'+m.toString()).checked){
                    NumberOfSearchFieldBoxesChecked--                                           //increments on check, decrements on uncheck
                  }else{
                    NumberOfSearchFieldBoxesChecked++
                  }
                  if(NumberOfSearchFieldBoxesChecked!=0){                                       //if at least 1 box is checked, enable start/stop search buttons
                    let ButtonElement = document.getElementById('StartSearchButton');
                    ButtonElement.disabled = false;
                    ButtonElement.style.opacity = '1.0';
                    ButtonElement.style.pointerEvents = 'auto';
                  }else{
                    ButtonElement = document.getElementById('StartSearchButton');
                    ButtonElement.disabled = true;
                    ButtonElement.style.opacity = '0.5';
                    ButtonElement.style.pointerEvents = 'none';
                  }
                };
                td.appendChild(Input);
              };
              if (j == 2) {
                SearchFieldName = document.createElement('span');
                //data[0][~] is the user-supplied name for the BP search field ~.
                if(BPHeadersOfUserData[m]=="BP_Author Dates"){
                  SearchFieldName.innerHTML = "Author Birth Date"; // + data[0][BPHeadersOfUserData[m]] + " ";
                }else{
                  SearchFieldName.innerHTML = " " + data[0][BPHeadersOfUserData[m]] + " ";
                  // limit field name length
                  if( SearchFieldName.innerHTML.length > 17 ){
                    SearchFieldName.innerHTML = SearchFieldName.innerHTML.substring(0,17)+"...";
                  }
                }
                td.appendChild(SearchFieldName);
              };
              if (j == 3) {
                Input = document.createElement('input');
                Input.type = 'search';
                Input.setAttribute( 'name' , 'Search Value');
                Input.setAttribute( 'class'  , 'ASearchValue');
                Input.setAttribute( 'title'  , 'Input search criterion text');
                if(BPHeadersOfUserData[m] == "BP_Nonce-volume" ||
                   BPHeadersOfUserData[m] == "BP_Image1" ||
                   BPHeadersOfUserData[m] == "BP_Image2" ||
                   BPHeadersOfUserData[m] == "BP_Image3" ||
                   BPHeadersOfUserData[m] == "BP_Image4" ||
                   BPHeadersOfUserData[m] == "BP_Image5" ||
                   BPHeadersOfUserData[m] == "BP_links" ||
                   BPHeadersOfUserData[m] == "BP_Digital Copy link"){
                    // Input.pattern = "!\?";
                    // Input.size = 2;
                    Input.maxlength = 2;
                   }
                td.appendChild(Input);
              };
              if (j==4){
                Input = document.createElement('button');
                Input.setAttribute( 'onclick'  , ' GetSearchList(' + m.toString() + ' ) ');
                Input.setAttribute( 'id'  , 'GetSearchListButton');
                Input.setAttribute( 'title'  , 'List Available Entries');
                Input.setAttribute( 'class'  , 'GetSearchListButton');
                Input.name = 'Open';
                Input.innerHTML = "<img class='GetSearchListButtonImage' src='BP Appearance/List.png' alt='Some png'>"
                if(BPHeadersOfUserData[m]=="BP_Author full name" ||
                   BPHeadersOfUserData[m]=="BP_Author last name first" ||
                   BPHeadersOfUserData[m]=="BP_Translator or Editor" ||
                   BPHeadersOfUserData[m]=="BP_Publisher or Seller" ||
                   BPHeadersOfUserData[m]=="BP_Publication Place" ||
                   BPHeadersOfUserData[m]=="BP_Printer" ||
                   BPHeadersOfUserData[m]=="BP_Format and Size" ||
                   BPHeadersOfUserData[m]=="BP_Format"){
                    Input.style.visiblity = "visible";
                   }else{
                    Input.style.visibility = "hidden";
                }
                td.appendChild(Input);
              };                
              if (j==5){
                Input = document.createElement('button');
                Input.setAttribute( 'onclick'  , ' MapSearchFieldToHits(' + m.toString() + ' ) ');
                Input.setAttribute( 'id'  , 'SearchValueMapButton');
                Input.setAttribute( 'title'  , 'Show Search Hits');
                Input.setAttribute( 'class'  , 'SearchValueMapButton');
                Input.name = 'Open';
                Input.innerHTML = "<img class='SearchValueMapButtonImage' src='BP Appearance/Target.png' alt='Some png'>"
                td.appendChild(Input)
              };
              tr.appendChild(td);
            }
          }
          document.getElementById("SearchFieldsContainer").appendChild(table);
          //#endregion
          
          // build table of filter fields and filter values
          //#region
          Filtertable = document.createElement('TABLE');
          Filtertable.setAttribute('class', 'FilterTable')
          Filtertable.setAttribute('id', 'FilterTable')
          FiltertableBody = document.createElement('TBODY');
          Filtertable.appendChild(FiltertableBody);
          for (let m = 1; m <= NumUserDataFields; m++) {
            if(data[0][BPHeadersOfUserData[m]]==''){continue}            
            let tr = document.createElement('tr');
            FiltertableBody.appendChild(tr);
            for (let j = 1; j <= 5; j++) {
              td = document.createElement('td');
              if (j == 1) {
                Input = document.createElement('input');
                Input.type = 'checkbox';
                Input.name = 'Filter Field';
                Input.id = 'FilterFieldElement'+m.toString();
                Input.title = 'Click to activate filtering by this data category'
                Input.onchange = ()=>{                                                          //checks whether filter field box is checked
                  if( !document.getElementById('FilterFieldElement'+m.toString()).checked){
                    NumberOfFilterFieldBoxesChecked--                                           //increments on check, decrements on uncheck
                  }else{
                    NumberOfFilterFieldBoxesChecked++
                  }
                  if(NumberOfFilterFieldBoxesChecked!=0){                                       //if at least 1 box is checked, enable start/stop fileter buttons
                    let ButtonElement = document.getElementById('StartFilterButton');
                    ButtonElement.disabled = false;
                    ButtonElement.style.opacity = '1.0';
                    ButtonElement.style.pointerEvents = 'auto';
                  }else{
                    ButtonElement = document.getElementById('StartFilterButton');
                    ButtonElement.disabled = true;
                    ButtonElement.style.opacity = '0.5';
                    ButtonElement.style.pointerEvents = 'none';
                  }
                };
                td.appendChild(Input);
              };
              if (j == 2) {
                SearchFieldName = document.createElement('span');
                //data[0][~] is the user-supplied name for the BP search field ~.
                if(BPHeadersOfUserData[m]=="BP_Author Dates"){
                  SearchFieldName.innerHTML = "Author Birth Date"; // + data[0][BPHeadersOfUserData[m]] + " ";
                }else{
                  SearchFieldName.innerHTML = " " + data[0][BPHeadersOfUserData[m]] + " ";
                  if( SearchFieldName.innerHTML.length > 17 ){
                    SearchFieldName.innerHTML = SearchFieldName.innerHTML.substring(0,17)+"...";
                  }                  
                }
                td.appendChild(SearchFieldName);
              };
              if (j == 3) {
                Input = document.createElement('input');
                Input.setAttribute( 'name' , 'Fileter Value');
                Input.setAttribute( 'class'  , 'ASearchValue');
                Input.setAttribute( 'title'  , 'Input filter criterion text');
                Input.type = 'search';
                if(BPHeadersOfUserData[m] == "BP_Nonce-volume" ||
                   BPHeadersOfUserData[m] == "BP_Image1" ||
                   BPHeadersOfUserData[m] == "BP_Image2" ||
                   BPHeadersOfUserData[m] == "BP_Image3" ||
                   BPHeadersOfUserData[m] == "BP_Image4" ||
                   BPHeadersOfUserData[m] == "BP_Image5" ||
                   BPHeadersOfUserData[m] == "BP_links" ||
                   BPHeadersOfUserData[m] == "BP_Digital Copy link"){
                    // Input.pattern = "!\?";
                    Input.size = 2;
                    Input.maxlength = 2;
                   }
                td.appendChild(Input);
              };
              if (j==4){
                Input = document.createElement('button');
                Input.setAttribute( 'onclick'  , ' GetFilterList(' + m.toString() + ' ) ');
                Input.setAttribute( 'id'  , 'GetFilterListButton');
                Input.setAttribute( 'title'  , 'List Available Entries');
                Input.setAttribute( 'class'  , 'GetFilterListButton');
                Input.name = 'Open';
                Input.innerHTML = "<img class='GetFilterListButtonImage' src='BP Appearance/List.png' alt='Some png'>"
                if(BPHeadersOfUserData[m]=="BP_Author full name" ||
                   BPHeadersOfUserData[m]=="BP_Author last name first" ||
                   BPHeadersOfUserData[m]=="BP_Translator or Editor" ||
                   BPHeadersOfUserData[m]=="BP_Publisher or Seller" ||
                   BPHeadersOfUserData[m]=="BP_Publication Place" ||
                   BPHeadersOfUserData[m]=="BP_Printer" ||
                   BPHeadersOfUserData[m]=="BP_Format and Size" ||
                   BPHeadersOfUserData[m]=="BP_Format"
                  //  BPHeadersOfUserData[m]=="BP_Size"
                   ){
                    Input.style.visiblity = "visible";
                   }else{
                    Input.style.visibility = "hidden";
                }
                td.appendChild(Input);
              };                
              tr.appendChild(td);
            }
          }
          document.getElementById("FilterFieldsContainer").appendChild(Filtertable);
          //#endregion

          DisplayElement = Array(data.length).fill(true);
          // get publication year (for year-ordering), last name (for lexical ordering) and edit paragraph breaks into user text.
          // HeaderForDateSorting = "BP_Date Published"  //this is the default header for date dorting upon parsing the csv file
          // GetDateDataForSorting(HeaderForDateSorting)
          // HeaderForAlphaSorting = "BP_Author last name first"  //this is the default header for alpha dorting upon parsing the csv file
          // GetAlphaDataForSorting(HeaderForAlphaSorting)
          LocationDataPresent = false;
          LocationDataDescription = '';
          LocationData[0] = 0;

          for (i = 0; i < data.length; i++) {
            if (i == 0) {
              ShortTitle[0] = " ";
              DisplayElement[i] = false;
            } else {
              //while we're here, find and replace all occurances of "|" in the user's text with an HTML "<br>" to force a new line in the text display
              if(data[i]["BP_Description_1"]){
                var text = data[i]["BP_Description_1"];
                text = text.replaceAll('|', "<br>");
                text = text.replaceAll('\n', "<br>");
                data[i]["BP_Description_1"] = text;
              }
              if(data[i]["BP_Description_2"]){
                var text = data[i]["BP_Description_2"];
                text = text.replaceAll('|', "<br>");
                text = text.replaceAll('\n', "<br>");                
                data[i]["BP_Description_2"] = text;
              }
              if(data[i]["BP_Reference(s)"]){
                var text = data[i]["BP_Reference(s)"];
                text = text.replaceAll('|', "<br>"); 
                text = text.replaceAll('\n', "<br>");                
                data[i]["BP_Reference(s)"] = text;
              }
              if(data[i]["BP_Collation/Pagination/Foliation"]){
                var text = data[i]["BP_Collation/Pagination/Foliation"];
                text = text.replaceAll('|', "<br>");
                text = text.replaceAll('\n', "<br>");                
                data[i]["BP_Collation/Pagination/Foliation"] = text;
              }
              if(data[i]["BP_Collation"]){
                var text = data[i]["BP_Collation"];
                text = text.replaceAll('|', "<br>");
                text = text.replaceAll('\n', "<br>");                
                data[i]["BP_Collation"] = text;
              }
              if(data[i]["BP_Pagination"]){
                var text = data[i]["BP_Pagination"];
                text = text.replaceAll('|', "<br>");
                text = text.replaceAll('\n', "<br>");                
                data[i]["BP_Pagination"] = text;
              }
              if(data[i]["BP_Foliation"]){
                var text = data[i]["BP_Foliation"];
                text = text.replaceAll('|', "<br>"); 
                text = text.replaceAll('\n', "<br>");                
                data[i]["BP_Foliation"] = text;
              }
              if(data[i]["BP_Plates"]){
                var text = data[i]["BP_Plates"];
                text = text.replaceAll('|', "<br>");
                text = text.replaceAll('\n', "<br>");                
                data[i]["BP_Plates"] = text;
              }
              if(data[i]["BP_BibliographicNotes"]){
                var text = data[i]["BP_BibliographicNotes"];
                text = text.replaceAll('|', "<br>");
                text = text.replaceAll('\n', "<br>");
                data[i]["BP_BibliographicNotes"] = text;
              }

              if(data[i]["BP_Short or Common Title"]){
                ShortTitle[i] = data[i]["BP_Short or Common Title"]
              }else{
                ShortTitle[i] = data[i]["BP_Title"]; 
              }

              //trim links to images
              NumberOfImagesWithEntry[i] = 0;
              if(data[i]["BP_Image1"]){data[i]["BP_Image1"] = data[i]["BP_Image1"].slice(data[i]["BP_Image1"].lastIndexOf("/")+1).trim();NumberOfImagesWithEntry[i]++}
              if(data[i]["BP_Image2"]){data[i]["BP_Image2"] = data[i]["BP_Image2"].slice(data[i]["BP_Image2"].lastIndexOf("/")+1).trim();NumberOfImagesWithEntry[i]++}
              if(data[i]["BP_Image3"]){data[i]["BP_Image3"] = data[i]["BP_Image3"].slice(data[i]["BP_Image3"].lastIndexOf("/")+1).trim();NumberOfImagesWithEntry[i]++}
              if(data[i]["BP_Image4"]){data[i]["BP_Image4"] = data[i]["BP_Image4"].slice(data[i]["BP_Image4"].lastIndexOf("/")+1).trim();NumberOfImagesWithEntry[i]++}
              if(data[i]["BP_Image5"]){data[i]["BP_Image5"] = data[i]["BP_Image5"].slice(data[i]["BP_Image5"].lastIndexOf("/")+1).trim();NumberOfImagesWithEntry[i]++}
              ImageCompressionCollection[i] = [];  //Array(NumberOfImagesWithEntry[i]);

              // if present, capture location data
              if( data[i]["BP_Latitude"] != null && data[i]["BP_Latitude"] != '' && data[i]["BP_Latitude"] != ' '  &&
                  data[i]["BP_Longitude"] != null && data[i]["BP_Longitude"] != '' && data[i]["BP_Longitude"] != ' ') {
                    LocationDataPresent = true;
                    LocationDataDescription = data[0]["BP_Latitude"].trim()+' '+data[0]["BP_Longitude"].trim();
                    LocationData[0]++;
                    LocationData[LocationData[0]] = [];
                    LocationData[LocationData[0]][0] = i;
                    LocationData[LocationData[0]][1] = Number(data[i]["BP_Latitude"]);
                    LocationData[LocationData[0]][2] = Number(data[i]["BP_Longitude"]);
               };
            }
          }

          // get publication year (for year-ordering), last name (for lexical ordering) and edit paragraph breaks into user text.
          HeaderForDateSorting = "BP_Date Published"  //this is the default header for date dorting upon parsing the csv file
          GetDateDataForSorting(HeaderForDateSorting)
          HeaderForAlphaSorting = "BP_Author last name first"  //this is the default header for alpha dorting upon parsing the csv file
          GetAlphaDataForSorting(HeaderForAlphaSorting)

          // ImageCompressionCollectionEmpty = true;
          FoundInvalidYear = false;
          for (i = 0; i < data.length; i++) {
            //any invalid/nonexistent year data is set to just below the minimum year
            if(Year[i] == -123.123){
              Year[i] = MinYear;
              FoundInvalidYear = true;
            }
          }
          //if necessary, drop min year to leave room for invalid/nonexistant years
          if(FoundInvalidYear){MinYear--}
          // drop 0th data set, since it's the user names for header categories
          NumOfTimelineElements = data.length - 1;

          //search lists
          //#region          
          //generate lists of data to be used for search lists: author
          let SetOfAuthors = new Set();
          for (n = 1; n <= NumOfTimelineElements; n++) {
            Author = data[n]["BP_Author last name first"];
            Author = Author.replace(/\(/g, "");
            Author = Author.replace(/\)/g, "");
            Author = Author.trim();
            SetOfAuthors.add(Author);
          }
          ListOfAuthors = Array.from(SetOfAuthors);
          ListOfAuthors.sort(function (a, b) {return a.toLowerCase().localeCompare(b.toLowerCase());});
          NumberOfAuthorsInList = ListOfAuthors.length

          //generate lists of data to be used for search lists: author dates
          // var ListOfAuthorDates = [];

          //generate lists of data to be used for search lists: translator or editor
          if( data[0]["BP_Translator or Editor"] ||
              data[0]["BP_Translator"] ||
              data[0]["BP_Editor"] ){
            let SetOfTransOrEdit = new Set();
            for (n = 1; n <= NumOfTimelineElements; n++) {
              TransOrEdit = data[n]["BP_Translator or Editor"];
              if(TransOrEdit=='' || TransOrEdit==' '){data[n]["BP_Translator"]};
              if(TransOrEdit=='' || TransOrEdit==' '){data[n]["BP_Editor"]};
              if(TransOrEdit=='' || TransOrEdit==' '){continue};
              FirstIndex = 0;
              if(TransOrEdit.includes("(")){
                LastIndex = TransOrEdit.search(/\(/);
                if(LastIndex==0){
                  FirstIndex = 1;
                  LastIndex = TransOrEdit.search(/\)/);
                }
              }else{
                LastIndex = TransOrEdit.length;
              }
              TransOrEdit = TransOrEdit.substring(FirstIndex,LastIndex);
              TransOrEdit = TransOrEdit.trim();
              // TransOrEdit = TransOrEdit.substring(0,TransOrEdit.search(/[0-9]/));
              SetOfTransOrEdit.add(TransOrEdit);
            }
            ListOfTransOrEdit = Array.from(SetOfTransOrEdit);
            ListOfTransOrEdit.sort(function (a, b) {return a.toLowerCase().localeCompare(b.toLowerCase());});
            NumberOfTransOrEditInList = ListOfTransOrEdit.length
          }else{
            NumberOfTransOrEditInList = 0;
          }

          // //pub date
          // var ListOfPubDates = [];

          //generate lists of data to be used for search lists: publisher
          if(data[0]["BP_Publisher or Seller"]){
            let SetOfPublishers = new Set();
            for (n = 1; n <= NumOfTimelineElements; n++) {
              Publishers = data[n]["BP_Publisher or Seller"];
              FirstIndex = 0;
              if(Publishers.includes("(")){
                LastIndex = Publishers.search(/\(/);
                if(LastIndex==0){
                  FirstIndex = 1;
                  LastIndex = Publishers.search(/\)/);
                }
              }else{
                LastIndex = Publishers.length;
              }
              Publishers = Publishers.substring(FirstIndex,LastIndex);
              Publishers = Publishers.trim();
              // TransOrEdit = TransOrEdit.substring(0,TransOrEdit.search(/[0-9]/));
              SetOfPublishers.add(Publishers);
            }
            ListOfPublishers = Array.from(SetOfPublishers);
            ListOfPublishers.sort(function (a, b) {return a.toLowerCase().localeCompare(b.toLowerCase());});
            NumberOfPublishersInList = ListOfPublishers.length
          }else{
            NumberOfPublishersInList = 0;
          }
          //generate lists of data to be used for search lists: publication place
          if(data[0]["BP_Publication Place"]){
            let SetOfPublicationPlaces = new Set();
            for (n = 1; n <= NumOfTimelineElements; n++) {
              PublicationPlaces = data[n]["BP_Publication Place"];
              FirstIndex = 0;
              if(PublicationPlaces.includes("(")){
                LastIndex = PublicationPlaces.search(/\(/);
                if(LastIndex==0){
                  FirstIndex = 1;
                  LastIndex = PublicationPlaces.search(/\)/);
                }
              }else{
                LastIndex = PublicationPlaces.length;
              }
              PublicationPlaces = PublicationPlaces.substring(FirstIndex,LastIndex);;
              PublicationPlaces = PublicationPlaces.trim();
              // TransOrEdit = TransOrEdit.substring(0,TransOrEdit.search(/[0-9]/));
              SetOfPublicationPlaces.add(PublicationPlaces);
            }
            ListOfPublicationPlaces = Array.from(SetOfPublicationPlaces);
            ListOfPublicationPlaces.sort(function (a, b) {return a.toLowerCase().localeCompare(b.toLowerCase());});
            NumberOfPublicationPlacesInList = ListOfPublicationPlaces.length;
          }else{
            NumberOfPublicationPlacesInList = 0;
          }

          //generate lists of data to be used for search lists: printer
          if(data[0]["BP_Printer"]){
            let SetOfPrinters = new Set();
            for (n = 1; n <= NumOfTimelineElements; n++) {
              Printers = data[n]["BP_Printer"];
              FirstIndex = 0;
              if(Printers.includes("(")){
                LastIndex = Printers.search(/\(/);
                if(LastIndex==0){
                  FirstIndex = 1;
                  LastIndex = Printers.search(/\)/);
                }
              }else{
                LastIndex = Printers.length;
              }
              Printers = Printers.substring(FirstIndex,LastIndex);
              Printers = Printers.trim();
              // TransOrEdit = TransOrEdit.substring(0,TransOrEdit.search(/[0-9]/));
              SetOfPrinters.add(Printers);
            }
            ListOfPrinters = Array.from(SetOfPrinters);
            ListOfPrinters.sort(function (a, b) {return a.toLowerCase().localeCompare(b.toLowerCase());});
            NumberOfPrintersInList = ListOfPrinters.length
          }else{
            NumberOfPrintersInList = 0;
          }

          //generate lists of data to be used for search lists: formats
          ListOfFormats[0] = "1"+String.fromCharCode(176);
          ListOfFormats[1] = "2"+String.fromCharCode(176);
          ListOfFormats[2] = "4"+String.fromCharCode(176);
          ListOfFormats[3] = "6"+String.fromCharCode(176);
          ListOfFormats[4] = "8"+String.fromCharCode(176);
          ListOfFormats[5] = "12"+String.fromCharCode(176);
          ListOfFormats[6] = "16"+String.fromCharCode(176);
          ListOfFormats[7] = "24"+String.fromCharCode(176);
          ListOfFormats[8] = "32"+String.fromCharCode(176);
          ListOfFormats[9] = "64"+String.fromCharCode(176);
          NumberOfFormatsInList = ListOfFormats.length
          //#endregion

          //build a pointer to elements sorted by date published
          heapSortByDate();
          DateSortPointer[0] = 0;
          //year array into sorted order.
          // NB!  the sort method sors in place with the first value of the result at Year[0].
          // Year.sort();
          //build a pointer to elements sorted lexically by last name
          heapSortAlphabetize();
          AlphabeticalSortPointer[0] = 0;
          DataHasBeenInput = true;
          document.getElementById("MakeDOCfromBookPageantButton").style.opacity = "1.0";
          document.getElementById("MakeDOCfromBookPageantButton").disabled = false;

          //intial data display data. chronologically if no saved value
          //data dorting and display type
          ChronologicalSort = localStorage.getItem("ChronologicalSort");
          AlphabeticalSort = localStorage.getItem("AlphabeticalSort");
          if(ChronologicalSort == "true"){
            ChronologicalSort = true;
            AlphabeticalSort = false;
            YearSlider = localStorage.getItem("YearSlider");
            ProcessDataChronologically();
            // if(YearSlider){ScaleOfYears(YearSlider)};
            document.getElementById("ScaleSlider").value = Number(YearSlider);
          }else if(AlphabeticalSort == "true"){
            ChronologicalSort = false;
            AlphabeticalSort = true;
            AlphabetSlider = localStorage.getItem("AlphabetSlider");
            ProcessDataAlphabetically();
            // if(AlphabetSlider){ScaleOfAlphabet(AlphabetSlider)}
            document.getElementById("AlphaSlider").value = Number(AlphabetSlider);
          }else{
            ProcessDataChronologically()
          }
          // wheelzoom(document.getElementById('BP_Image1'));
        }
      }
    })
    //opening default. Can be overridden by saved data from the Composer.
    PaperNumber = 1;
    PageMargins[1] = 1;
    PageMargins[2] = 1;
    PageMargins[3] = 1;
    PageMargins[4] = 1;
    //set maximum number in element spinner on composition board
    document.getElementById("MakeDOCfromBookPageantButton").min=1;
    document.getElementById("MakeDOCfromBookPageantButton").max=NumOfTimelineElements;
    document.body.style.cursor = "initial";
  } 
// clear any .BP or .BP.CSV display that may be present
  function ClearTimeline(){
    data.length = 0;
    meta.length = 0;
    BPCSVfileHasBeenLoaded = false;
    BPfileHasBeenLoaded = false;
    CompressedImagesfileHasBeenLoaded = false;
    CSVfileHasBeenLoaded = false;
    let TimeLine = document.getElementById("MainTimeline");
    while (TimeLine.hasChildNodes()) {
      TimeLine.removeChild(TimeLine.firstChild);
    }
    DateSortPointer.length = 0;
    AlphabeticalSortPointer.length = 0;
    DataHasBeenInput = false;
   }
//#endregion

// csv file checking
//#region
  function CheckCSVFile(){
      //bring up the list of fields from the CSV file
      CSVFieldElements = document.getElementById("AvailableCSVFieldElementHeaders");
      CSVFieldElements.style.zIndex = "1050";
      CSVFieldElements.style.visibility = "visible";
      CSVFieldElements.style.backgroundColor = "var(--Gray10)";
      dragElement(document.getElementById("AvailableCSVFieldElementHeadersDrag"));
      //clear previous header list, if present
      try{
      document.getElementById("AvailableCSVFieldElementHeadersList").remove()
      }catch{}
     // prep header text
      document.getElementById('CSVFieldElementHeaderList').innerHTML = TextForinnerHTMLofCSVCheckterButton;
      // turn off color coding, if present
      if(document.getElementById("ColorCodingCSVProblems")) {
        document.getElementById("ColorCodingCSVProblems").style.visibility = "hidden";
      }
      // clear all previous cell coloring
      // grid._clearCellColors([{x:0,y:0},{x:meta.fields.length,y:data.length}]);      
      clearCellColors([{x:0,y:0},{x:meta.fields.length,y:data.length}]);
      if(gridCellsWithMissingImages[0] != 0){
        // restore color to missing image file cells
        let SpreadsheetCellx = 0
        let SpreadsheetCelly = 0
        for(let m=1; m<=gridCellsWithMissingImages[0]; m++){
          SpreadsheetCellx = gridCellsWithMissingImages[m][0];
          SpreadsheetCelly = gridCellsWithMissingImages[m][1];
          setCellColors([{x:SpreadsheetCellx,y:SpreadsheetCelly},{x:SpreadsheetCellx,y:SpreadsheetCelly}], '#87cefa');
        }
      }

      // color the header row
      // grid._setCellColors([{x:0,y:0},{x:meta.fields.length,y:0}], '#D7E5F0');
      // setCellColors([{x:0,y:0},{x:meta.fields.length-1,y:0}], '#D7E5F0');
      // clear list of problem cells
      gridCellsWithProblems.length = 0;

      let collection = document.getElementsByClassName("MoveToCellButton");
      for( let m=0; m<collection.length; m++){collection[m].style.visibility = "hidden"};
      
      // createe list of available headers 
      CSVFieldElementsContent = document.createElement('div');
      CSVFieldElementsContent.setAttribute( 'id'  , 'AvailableCSVFieldElementHeadersList');
      CSVFieldElementsContent.setAttribute( 'class'  , 'AvailableCSVFieldElementHeadersList');

      NumCSVFields = 0;
      // initialize list of headers to check: minimum is three required: date, title, author.
      NumberOfHeadersChoosenForCSVChecking[0] = 3;
      NumberOfHeadersChoosenForCSVChecking[1] = 0;
      NumberOfHeadersChoosenForCSVChecking[2] = 0;
      NumberOfHeadersChoosenForCSVChecking[3] = 0;
    
      CurrentCSVDataRowMax = data.length-1;
      // initialize spell-checking array and scan data for <i><i> pairs, replacing them with <i></i>.
      for (let m = 0; m < meta.fields.length; m++) {        
        for (let n=1; n<=CurrentCSVDataRowMax; n++){
          // CSVDataSpellClean[m+1][n] = 1;
          // console.log(n.toString()+" "+m.toString());
          if( (data[n][meta.fields[m]].match(/'<\/i>'/gi)||[]).length == 0 ){
            data[n][meta.fields[m]] = data[n][meta.fields[m]].split( /<i>/ ).reduce( (a,b, i) => i %2 == 0 ? a + "</i>" + b : a + "<i>" + b );
          }
        }
      }

      for (m = 0; m < meta.fields.length; m++) {
        NumCSVFields = NumCSVFields + 1;
        CSVFieldElement = document.createElement('div');
        CSVFieldElement.setAttribute( 'id'  , 'CSVFieldElement'+NumCSVFields.toString());
        CSVFieldElement.setAttribute( 'class'  , 'CSVFieldElement');
        CSVFieldElement.setAttribute( 'onclick'  , "GetDateAuthorTitleFieldnumbersForChecking("+NumCSVFields.toString()+")" ) ;
        CSVFieldElement.name = 'notrequired';
        CSVFieldElement.innerHTML = meta.fields[m];
        CSVFieldElementsContent.appendChild(CSVFieldElement)
        RequiredDataFields[NumCSVFields] = "false";
      }  

      CSVFieldElements.appendChild(CSVFieldElementsContent);
      UserAssistance( "Click on a CSV data header from the left-hand list that gives publication/creation date.",
      "This will be used by BookPageant to chronologically sort items.", '','', 20.0);      
  }  
  function FinishCheckCSVFile(){
      document.getElementById("CSVFileChecking").style.visibility = "visible";
      dragElement(document.getElementById("CSVFileCheckingDrag"));
      Analysls = document.getElementById('CSVAnalysis')
      Analysls.innerHTML = '<strong>Analysis of CSV File</strong><br><br>';
      Analysls.innerHTML += 'Number of headers or data fields: <strong>'+SpreadsheetHeaders.length.toString()+'</strong><br>';
      // test for duplicate header names. 
      DuplicateHeaders = false;
      EmptyHeaders = false;
      gridCellsWithProblems.length = 0;
      gridCellsWithProblems[0] = 0
      // clear all previous cell coloring
      let HeadersSoFar = Object.create(null);
      for (var i = 0; i < SpreadsheetHeaders.length; ++i) {
          var value = SpreadsheetHeaders[i];
          if(value=='' || value==' '){
            EmptyHeaders = true;
            EmptyHeaderColumn = i+1;
            break
          }
          if (value in HeadersSoFar) {
              DuplicateHeader = SpreadsheetHeaders[i]
              DuplicateHeaders = true;
              DuplicateHeaderNumber = i
              break
          }
          HeadersSoFar[value] = true;
      }
      if(DuplicateHeaders){
        DuplicateHeaderMessage = '<br><strong>Duplicate Headers found: '+DuplicateHeader+'</strong><br>'
        Analysls.innerHTML += DuplicateHeaderMessage.fontcolor("red")
        // change color is appropriate cells to RED
        SpreadsheetCellx = DuplicateHeaderNumber;
        SpreadsheetCelly = 0;
        // grid._setCellColors([{x:SpreadsheetCellx,y:SpreadsheetCelly},{x:SpreadsheetCellx,y:SpreadsheetCelly}], '#ff0000');
        setCellColors([{x:SpreadsheetCellx,y:SpreadsheetCelly},{x:SpreadsheetCellx,y:SpreadsheetCelly}], '#ff0000');
        gridCellsWithProblems[0]++
        // gridCellsWithProblems[gridCellsWithProblems[0]] = {x:SpreadsheetCellx,y:SpreadsheetCelly}
        gridCellsWithProblems[gridCellsWithProblems[0]] = [SpreadsheetCellx,SpreadsheetCelly];

      }else if(EmptyHeaders){
        EmptryHeaderMessage = '<br><strong>Empty or Missing Header at column: '+EmptyHeaderColumn.toString()+'</strong><br>'
        Analysls.innerHTML += EmptryHeaderMessage.fontcolor("red")
        // change color is appropriate cells to RED
        SpreadsheetCellx = EmptyHeaderColumn-1;
        SpreadsheetCelly = 0;
        setCellColors([{x:SpreadsheetCellx,y:-1},{x:SpreadsheetCellx,y:-1}], '#ff0000');
        gridCellsWithProblems[0]++
        // gridCellsWithProblems[gridCellsWithProblems[0]] = {x:SpreadsheetCellx,y:SpreadsheetCelly}        
        gridCellsWithProblems[gridCellsWithProblems[0]] = [SpreadsheetCellx,SpreadsheetCelly];
      }
      
      Analysls.innerHTML += 'Number of records: <strong>'+data.length.toString()+'</strong><br>';
      Analysls.innerHTML += 'Number of CSV errors of form or structure: <strong>'+errors.length.toString()+'</strong><br>';
      for( let k=0; k<errors.length; k++){
        Analysls.innerHTML += '&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp'+'Record ' + errors[k].row.toString() + ': ' + errors[k].message.toString() + '<br>';
      }
      if(errors.length > 0 || DuplicateHeaders){
        Analysls.innerHTML += '<br><strong>Cannot continue CSV file assessment</strong><br>'
      }else{
        Analysls.innerHTML += '<br><strong>ID critical data fields and analyze CSV file</strong><br>'
      }
      CSVFileInitialAnalysis = Analysls.innerHTML;
      AssessCSVFile()
    // }  
    // Reader.readAsText(InputFile);
    UserAssistance( "From the left-hand list, choose the data to be verified by clicking on the header from the CSV file.",
    "Minimally, headers for data describing <strong>author, title, and date</strong> should be verified. ",
    "Initiate Data verification by clicking the continue arrow",
    "CheckCSVFile.png",
     30.0);
  }
  function GetDateAuthorTitleFieldnumbersForChecking(FieldNumber){
    CSVFieldElement = document.getElementById('CSVFieldElement'+FieldNumber.toString());
    let NewNumberHeaders = 0;
    if(CSVFieldElement.name == 'notrequired'){
      // add user-selected item to header field list
      // first see if this field is being selected as one of the required first three.
      if(        NumberOfHeadersChoosenForCSVChecking[1] == 0){
          NumberOfCheckingHeader = 1;
          document.getElementById('CSVFieldElement'+FieldNumber.toString()).innerHTML += "<small><i> (Header's data will be used for Date)</i></small>";        
        }else if(NumberOfHeadersChoosenForCSVChecking[2] == 0){
          NumberOfCheckingHeader = 2;
          document.getElementById('CSVFieldElement'+FieldNumber.toString()).innerHTML += "<small><i> (Header's data will be used for Title)</i></small>";
        }else if(NumberOfHeadersChoosenForCSVChecking[3] == 0){
          NumberOfCheckingHeader = 3;
          document.getElementById('CSVFieldElement'+FieldNumber.toString()).innerHTML += "<small><i> (Header's data will be used for Author)</i></small>";
        }else{
        // three required headers are indicated, addition checking header is being added
        NumberOfHeadersChoosenForCSVChecking[0]++;
        NumberOfCheckingHeader = NumberOfHeadersChoosenForCSVChecking[0]
      }
      CSVFieldElement.style.color = 'var(--AlertColor)';
      CSVFieldElement.style.backgroundColor = 'var(--Gray0)';
      CSVFieldElement.style.fontWeight = 'bold'
      CSVFieldElement.name = 'required';
      RequiredDataFields[FieldNumber] = true;
      NumberOfHeadersChoosenForCSVChecking[NumberOfCheckingHeader] = FieldNumber;

    }else{
      // remove user-deselected item from header field list
      for(let m=1; m<=NumberOfHeadersChoosenForCSVChecking[0]; m++){
        if(FieldNumber == NumberOfHeadersChoosenForCSVChecking[m]){
          if( m==1 || m==2 || m==3){
            CSVFieldElement = document.getElementById('CSVFieldElement'+FieldNumber.toString());
            CSVFieldElement.innerHTML = CSVFieldElement.innerHTML.substring(0,CSVFieldElement.innerHTML.indexOf('<small>'))
            CSVFieldElement.style.fontWeight = 'normal'
          }
          NumberOfHeadersChoosenForCSVChecking[m] = 0;
          break;
        }
      }
      // pack header field list. NB 1st three (required) header field at not packed
      for(let m=1; m<=NumberOfHeadersChoosenForCSVChecking[0]; m++){
        if(NumberOfHeadersChoosenForCSVChecking[m] != 0 || m <= 3){
          NewNumberHeaders++
          NumberOfHeadersChoosenForCSVChecking[NewNumberHeaders] = NumberOfHeadersChoosenForCSVChecking[m];
        }
      }
      NumberOfHeadersChoosenForCSVChecking[0] = NewNumberHeaders
      CSVFieldElement.style.color =  'black';
      CSVFieldElement.name = 'notrequired';
      CSVFieldElement.style.backgroundColor = 'var(--Gray0)';
      RequiredDataFields[FieldNumber] = false;
      // see if data view window needs to be closed
      if( CurrentCSVDataHeader+1 == FieldNumber){CloseCSVFileDataViewing()}
    }
    let FoundAMissingRequiredHeader = false;
    for( let m=1; m<=3; m++){
      if(NumberOfHeadersChoosenForCSVChecking[1]==0 || !NumberOfHeadersChoosenForCSVChecking[1]){
        // document.getElementById('CSVFieldElement'+FieldNumber.toString()).innerHTML += '<small><i> (Checking will use this header for Date)</i></small>';
        document.getElementById("CSVFieldElementHeaderList").innerHTML = 
          "Headers Found in the CSV File<br><span style = 'color:rgb(0, 179, 255);'>Click on the Header for items' Date</span>"
          FoundAMissingRequiredHeader = true;
          break;
      }
      if(NumberOfHeadersChoosenForCSVChecking[2]==0 || !NumberOfHeadersChoosenForCSVChecking[2]){
        // document.getElementById('CSVFieldElement'+FieldNumber.toString()).innerHTML += '<small><i> (Checking will use this header for Author)</i></small>';
        document.getElementById("CSVFieldElementHeaderList").innerHTML = 
          "Headers Found in the CSV File<br><span style = 'color:rgb(0, 179, 255);'>Click on the Header for items' Title</span>"
          FoundAMissingRequiredHeader = true;
          UserAssistance( "Click on a CSV data header from the left-hand list that gives item title.",
          "This will be used by BookPageant to alphabetially sort items.", '','', 45.0);                
          break;
      }      
      if(NumberOfHeadersChoosenForCSVChecking[3]==0 || !NumberOfHeadersChoosenForCSVChecking[3]){
        // document.getElementById('CSVFieldElement'+FieldNumber.toString()).innerHTML += '<small><i> (Checking will use this header for Title)</i></small>';
        document.getElementById("CSVFieldElementHeaderList").innerHTML = 
          "Headers Found in the CSV File<br><span style = 'color:rgb(0, 179, 255);'>Click on the Header items' Author</span>"
          FoundAMissingRequiredHeader = true;
          UserAssistance( "Click on a CSV data header from the left-hand list that gives item author.",
          "This will be used by BookPageant to alphabetially sort items.", '','', 45.0);                          
          break; 
      }
    }
    if(!FoundAMissingRequiredHeader){
      document.getElementById("CSVFieldElementHeaderList").innerHTML = 
      "Headers Found in the CSV File<br>Click on Additional Headers for Data Checking"
      UserAssistance( "Click on any additional CSV data headers from the left-hand list.",
      "BookPageant will verify that each record has data under these headers.", '','', 45.0);                
    }
    // change state of CSV file assess button, if necessary
    if(NumberOfHeadersChoosenForCSVChecking[1] ==0 ||
       NumberOfHeadersChoosenForCSVChecking[2] ==0 ||
       NumberOfHeadersChoosenForCSVChecking[3] ==0 ){
       document.getElementById("ContinueCSVFileCheckingButton").disabled = true;
       document.getElementById("ContinueCSVFileCheckingButton").style.opacity = 0.25;
       document.getElementById("ContinueCSVFileCheckingButton").style.cursor = 'default';
    }else{
      // the required headers have been identified. turn on continue checking button
      document.getElementById("ContinueCSVFileCheckingButton").disabled = false;
      document.getElementById("ContinueCSVFileCheckingButton").style.opacity = 1.0;
      document.getElementById("ContinueCSVFileCheckingButton").style.cursor = 'pointer';
      document.getElementById("ContinueCSVFileCheckingButton").onclick = function(){FinishCheckCSVFile()};
    }
  }
  function AssessCSVFile(){
    // clear all previous cell coloring
    clearCellColors([{x:0,y:0},{x:meta.fields.length,y:data.length}]);
    if(gridCellsWithMissingImages[0] != 0){
      // restore color to missing image file cells
      let SpreadsheetCellx = 0
      let SpreadsheetCelly = 0
      for(let m=1; m<=gridCellsWithMissingImages[0]; m++){
        SpreadsheetCellx = gridCellsWithMissingImages[m][0];
        SpreadsheetCelly = gridCellsWithMissingImages[m][1];
        setCellColors([{x:SpreadsheetCellx,y:SpreadsheetCelly},{x:SpreadsheetCellx,y:SpreadsheetCelly}], '#87cefa');
      }
    }  
    // move through each record of CSV file, checking data under headings specified by the user (NumberOfHeadersChoosenForCSVChecking)
    let NumCriticalErrors = 0;    
    let MissingDataErrors = 0;
    let NumberOfDataErrorsFound = 0;
    let BadRecordNumbers = [];
    let ErrorList = "";
    let RecordError = "";
    let FoundProblem = false;
    let Note = false;
    let YearDataProvided = "";
    let TextProvided = "";
    let SpreadsheetCellx = 0;
    let SpreadsheetCelly = 0;
    RecordLoop:
    for(let r=0; r<data.length; r++){
      RecordError = "";
      HeaderLoop:
      for( let m=1; m<=NumberOfHeadersChoosenForCSVChecking[0]; m++){
        RecordError = "Row "+(r+1).toString()+": ";
        FoundProblem = false;
        // see if data is missing aloogehter
        if( data[r][meta.fields[NumberOfHeadersChoosenForCSVChecking[m]-1]] == "" ){
          switch(m){
            case 1:
              RecordError += "Missing Date"; NumCriticalErrors++; FoundProblem = true;
              // change color is appropriate cell to RED
              SpreadsheetCellx = NumberOfHeadersChoosenForCSVChecking[m]-1;
              SpreadsheetCelly = r;
              setCellColors([{x:SpreadsheetCellx,y:SpreadsheetCelly},{x:SpreadsheetCellx,y:SpreadsheetCelly}], '#ff0000');
              gridCellsWithProblems[0]++
              gridCellsWithProblems[gridCellsWithProblems[0]] = [SpreadsheetCellx,SpreadsheetCelly];
              break;
            case 2:
              RecordError += "Missing Title"; NumCriticalErrors++; FoundProblem = true;
              // change color is appropriate cell to RED
              SpreadsheetCellx = NumberOfHeadersChoosenForCSVChecking[m]-1;
              SpreadsheetCelly = r;
              setCellColors([{x:SpreadsheetCellx,y:SpreadsheetCelly},{x:SpreadsheetCellx,y:SpreadsheetCelly}], '#ff0000');
              gridCellsWithProblems[0]++
              gridCellsWithProblems[gridCellsWithProblems[0]] = [SpreadsheetCellx,SpreadsheetCelly];
              break;
            case 3:
              RecordError += "Missing Author. Will be assigned 'Anonymous'"; NumCriticalErrors++; FoundProblem = true;
              // change color is appropriate cell to ORANGE
              SpreadsheetCellx = NumberOfHeadersChoosenForCSVChecking[m]-1;
              SpreadsheetCelly = r;
              setCellColors([{x:SpreadsheetCellx,y:SpreadsheetCelly},{x:SpreadsheetCellx,y:SpreadsheetCelly}], '#ff0000');
              gridCellsWithProblems[0]++
              gridCellsWithProblems[gridCellsWithProblems[0]] = [SpreadsheetCellx,SpreadsheetCelly];
              break;
            default:
              RecordError += "Missing data under Header: "+meta.fields[NumberOfHeadersChoosenForCSVChecking[m]-1]; FoundProblem = true;
              // change color is appropriate cell to ORANGE
              SpreadsheetCellx = NumberOfHeadersChoosenForCSVChecking[m]-1;
              SpreadsheetCelly = r;
              setCellColors([{x:SpreadsheetCellx,y:SpreadsheetCelly},{x:SpreadsheetCellx,y:SpreadsheetCelly}], '#ffa500');
              gridCellsWithProblems[0]++
              gridCellsWithProblems[gridCellsWithProblems[0]] = [SpreadsheetCellx,SpreadsheetCelly];
              break;
          }
        }
        // check date formating
        Note = false;
        if(m==1 && data[r][meta.fields[NumberOfHeadersChoosenForCSVChecking[m]-1]] != ""){
          YearDataProvided = data[r][meta.fields[NumberOfHeadersChoosenForCSVChecking[m]-1]];
          if( YearDataProvided.includes(".") ||
              YearDataProvided.includes("-") ||
              YearDataProvided.includes("?") ||
              YearDataProvided.includes("ca") ||
              YearDataProvided.includes("c.") ||
              YearDataProvided.includes("nd") ||
              YearDataProvided.includes("n.d.") ||
              YearDataProvided.includes("[") ||
              YearDataProvided.includes("]") ||
              YearDataProvided.includes("{") ||
              YearDataProvided.includes("}") ||
              YearDataProvided.toLowerCase().includes("unknown") ) {
            Note = true;
            YearExtracted = YearDataProvided.match(/\d+/);
            if(YearExtracted !=null){
              RecordError += "Year given as: "+YearDataProvided+" Will be interpreted as: "+YearExtracted;
              // change color is appropriate cell to YELLOW
              SpreadsheetCellx = NumberOfHeadersChoosenForCSVChecking[m]-1;
              SpreadsheetCelly = r;
              setCellColors([{x:SpreadsheetCellx,y:SpreadsheetCelly},{x:SpreadsheetCellx,y:SpreadsheetCelly}], '#ffff00');
              gridCellsWithProblems[0]++
              gridCellsWithProblems[gridCellsWithProblems[0]] = [SpreadsheetCellx,SpreadsheetCelly];
            }else{
              RecordError += "Year given as: "+YearDataProvided+" Will be assigned 1455";
              // change color is appropriate cell to YELLOW
              SpreadsheetCellx = NumberOfHeadersChoosenForCSVChecking[m]-1;
              SpreadsheetCelly = r;
              setCellColors([{x:SpreadsheetCellx,y:SpreadsheetCelly},{x:SpreadsheetCellx,y:SpreadsheetCelly}], '#ffff00');
              gridCellsWithProblems[0]++
              gridCellsWithProblems[gridCellsWithProblems[0]] = [SpreadsheetCellx,SpreadsheetCelly];
            }
          }
        }
        // check for missing italic tags
        if(m >=3 && data[r][meta.fields[NumberOfHeadersChoosenForCSVChecking[m]-1]] != ""){
          TextProvided = data[r][meta.fields[NumberOfHeadersChoosenForCSVChecking[m]-1]];
          let pos = 0;
          let ItalicTagCount = TextProvided.split("<i>").length-1;
          let ItalicTagCloseCount = TextProvided.split("</i>").length-1;
          if(ItalicTagCount > 0 && (ItalicTagCount != ItalicTagCloseCount) ){
            if( 2*( Math.floor(ItalicTagCount/2)) != ItalicTagCount ){        //fast versiion of Math.floor()
              RecordError += "Mismatched tags for italics in text under Header: "+meta.fields[NumberOfHeadersChoosenForCSVChecking[m]-1];
              Note = true;
            }
            if( 2*(Math.floor(ItalicTagCloseCount/2)) != ItalicTagCloseCount ){
              RecordError += "Mismatched tags for italics in text under Header: "+meta.fields[NumberOfHeadersChoosenForCSVChecking[m]-1];
              Note = true;
            }
          }
        }

        if(FoundProblem || Note){ErrorList += RecordError + "<br>"}        
      }
    }

    Analysls = document.getElementById('CSVAnalysis');
    if(ErrorList != "" || gridCellsWithMissingImages[0]>0){
      Analysls.innerHTML += ErrorList;
      // put color coding explanation into spreadsheet controls div
      let CSVControls = document.getElementById("CSVspreadsheetControls")
      if(!document.getElementById("ColorCodingCSVProblems")) {
        ColorCoding = document.createElement("div");
        ColorCoding.setAttribute( "ID", "ColorCodingCSVProblems");
        ColorCoding.setAttribute( "class", "ColorCodingCSVProblems");
        ColorCoding.style.visibility = "visible";
        ColorCoding.innerHTML = "<span style='color:red'>'███'</span><span style='color:black'><b>Critical, Missing </b></span>";
        ColorCoding.innerHTML += "<span style='color:orange'>' ███'</span><span style='color:black'><b>Optional, Missing </b></span>";
        ColorCoding.innerHTML += "<span style='color:yellow'>' ███'</span><span style='color:black'><b>Uncertain </b></span>";
        ColorCoding.innerHTML += "<span style='color:lightskyblue'>' ███'</span><span style='color:black'><b>No Image File </b></span>";
        ColorCoding.innerHTML +=  "&nbsp&nbsp<button type='button', id='MoveToCellButtonNext', class='MoveToCellButton' onclick = 'MoveGridViewPortPosition(-1)' title = 'Move to Previous Problem Cell'>"+
                                  "<img class='CSVCheckingButtonImages' src='BP Appearance/MoveBackToCell.png' alt='Some png'></img></button>"+
                                  "&nbsp&nbsp&nbsp<button type='button', id='MoveToCellButtonPrev', class='MoveToCellButton' onclick = 'MoveGridViewPortPosition(+1)' title = 'Move to Next Problem Cell'>"+
                                  "<img class='CSVCheckingButtonImages' src='BP Appearance/MoveToCell.png' alt='Some png'></img></button><span style='color:black'><b>&nbsp&nbsp&nbsp&nbspMove To Highlighted Cells</b></span>"
        CSVControls.insertBefore(ColorCoding, CSVControls.firstChild);
      }else{
        document.getElementById("ColorCodingCSVProblems").style.visibility = "visible";
        let collection = document.getElementsByClassName("MoveToCellButton");
        for( let m=0; m<collection.length; m++){
          collection[m].style.visibility = "visible";
          collection[m].disabled = false;
          };        
      }
    }else{
      Analysls.innerHTML += "No Problems Found";
      let collection = document.getElementsByClassName("MoveToCellButton");
      for( let m=0; m<collection.length; m++){
        collection[m].style.visibility = "visible";
        collection[m].disabled = true;
        };              
    }

    CheckCSVFileAnalysisReportText = Analysls.innerHTML;
    UserAssistance( "Use the Trash button to clear verification results.",
    "Use the Download Report button to save a text file of the analysis to the Standard Download Folder. ",
    "",
    "DownloadReport.png",
     30.0);
  }
  async function CheckAllImageFiles(){
    ImageNameInCSVfile.length = 0;
    let ColContainingImageFile = [];
    let FirstRow = 0;
    NumberOfFilesFailed = 0;
    ImageProcessingContainer.innerHTML = '';
    ImageCompressionFails = '';
    gridCellsWithMissingImages[0] = 0;
    // search the data from the CSV for cell content with an image file name
    if( BPCSVfileHasBeenLoaded ){
      FirstRow = 0;
    }else{
      FirstRow = 1;
    }
    NumImageFilesToFind = 0;
    for( let m=0; m<=meta.fields.length-1; m++){
      for (let n=FirstRow; n <= data.length-1; n++){
          let TextToCheck = data[n][meta.fields[m]].toLowerCase();
          if( TextToCheck.includes(".jpg") || 
              TextToCheck.includes(".tif") ||
              TextToCheck.includes(".png") ||
              TextToCheck.includes(".bmp") ){
                NumImageFilesToFind++
                ImageNameInCSVfile[NumImageFilesToFind] = TextToCheck.slice(TextToCheck.lastIndexOf("/")+1).trim()
              }
      }
    }
    
    if(NumImageFilesToFind==0){
      // no image file names found in CSV file. report null search
      // ImageProcessingContainer.innerHTML += "No image file names found in CSV file.<br>No search performed."
      // ImageCompressionFailsFormatted = "No image file names found in CSV file."+String.fromCharCode(13)+"No search performed."
    }else{
      // image file names found in CSV file. check against list of images found in BookPageant/Images folder
      if(BookPageantImagesDirectoryHandle != null){
        // BookPageantImagesDirectoryHandle = await window.showDirectoryPicker({startIn: BookPageantImagesDirectoryHandle[0], types: [{descriiption: 'image files', accept:{'text/plain':['.PDFformat.bp']}}]});
        BookPageantImagesDirectoryHandle = await window.showDirectoryPicker({startIn: BookPageantImagesDirectoryHandle});
      }else{
     // BookPageantImagesDirectoryHandle = await window.showDirectoryPicker({startIn: BookPageantHomeDirectoryHandle, types: [{descriiption: 'image files', accept:{'text/plain':['.PDFformat.bp']}}]});
        BookPageantImagesDirectoryHandle = await window.showDirectoryPicker({startIn: BookPageantHomeDirectoryHandle});
      }
      localforage.setItem('BookPageantImagesDirectoryHandle', BookPageantImagesDirectoryHandle);
      ImageDirectory = await BookPageantImagesDirectoryHandle;
      ImageNumber = 0;
      let ImageNameInDirectory = [];
      for await (const entry of ImageDirectory.values()) {
        ImageNumber++
        ImageNameInDirectory[ImageNumber] = entry.name.toLowerCase();
      }
      // move throgh lists to find missing image files
      ImageProcessingContainer.innerHTML += "Found<strong> "+NumImageFilesToFind.toString()+"</strong> image file names in CSV file.<br>"+
                                            "Search directory: <strong>/"+BookPageantImagesDirectoryHandle.name+"</strong><br><br>";
      ImageCSVLoop:
      for (let m=1; m <= data.length-1; m++){
        if(ImageNameInCSVfile[m]==''){continue};
        ImageNameToFind = ImageNameInCSVfile[m];
        FoundMatch = false;    
        ImageDirectoryLoop:
        for (let n=1; n<=ImageNameInDirectory.length; n++){
          if(ImageNameToFind == ImageNameInDirectory[n]){
            FoundMatch = true;
            break ImageDirectoryLoop;
          }
        }
        if(FoundMatch){ continue ImageCSVLoop}
        // if we get here we failed to find a match
        ErrorRowInCSVfile = m 
        if( BPCSVfileHasBeenLoaded ){ErrorRowInCSVfile++}    
        NumberOfFilesFailed++
        ImageProcessingContainer.innerHTML += "Number: " + ErrorRowInCSVfile.toString() + " in CSV file. Image not Found:<strong> "+ ImageNameToFind + "</strong><br>";
        ImageCompressionFails += ErrorRowInCSVfile.toString()+": "+ImageNameToFind + "<br>";
        // color cells in spreadsheet where images weren't found
        SpreadsheetCellx = ColContainingImageFile[m];
        SpreadsheetCelly = m;
        // grid._setCellColors([{x:SpreadsheetCellx,y:SpreadsheetCelly},{x:SpreadsheetCellx,y:SpreadsheetCelly}], '#568cc9');
        setCellColors([{x:SpreadsheetCellx,y:SpreadsheetCelly},{x:SpreadsheetCellx,y:SpreadsheetCelly}], '#87cefa');
        gridCellsWithMissingImages[0]++
        gridCellsWithMissingImages[gridCellsWithMissingImages[0]] = [SpreadsheetCellx,SpreadsheetCelly];
      }  
      // generate summary of image search
      ImageProcessingContainer.innerHTML += "<br><strong> Done: " + NumImageFilesToFind.toString() + " Images Sought </strong>";
      if(ImageCompressionFails.length>0){
        ImageProcessingContainer.innerHTML +="<br><strong>"+"Files Not Found: "+NumberOfFilesFailed+"</strong><br>"
        ImageCompressionFailsFormatted = "Found "+NumImageFilesToFind.toString()+" image file names in CSV file."+ String.fromCharCode(13)+
                                         "Search directory: /"+BookPageantImagesDirectoryHandle.name+""+String.fromCharCode(13)+String.fromCharCode(13)+
                                         "Files not found: "+NumberOfFilesFailed+String.fromCharCode(13)+
                                         ImageCompressionFails.replace(/<br>/gi, String.fromCharCode(13));
      }else{
        ImageCompressionFails += "Found "+NumImageFilesToFind.toString()+" image file names in CSV file."+ String.fromCharCode(13);
        ImageCompressionFailsFormatted = ImageCompressionFails+"All " + NumImageFilesToFind.toString() +" Files Found " + String.fromCharCode(13);
        ImageProcessingContainer.innerHTML += "<br><strong> All " + NumImageFilesToFind.toString() +" Image Files Found </strong>";    
      }
    }
    let CSVControls = document.getElementById("CSVspreadsheetControls")
    if(!document.getElementById("ColorCodingCSVProblems")) {
      ColorCoding = document.createElement("div");
      ColorCoding.setAttribute( "ID", "ColorCodingCSVProblems");
      ColorCoding.setAttribute( "class", "ColorCodingCSVProblems");
      ColorCoding.style.visibility = "visible";
      ColorCoding.innerHTML = "<span style='color:red'>'███'</span><span style='color:black'><b>Critical, Missing </b></span>";
      ColorCoding.innerHTML += "<span style='color:orange'>' ███'</span><span style='color:black'><b>Optional, Missing </b></span>";
      ColorCoding.innerHTML += "<span style='color:yellow'>' ███'</span><span style='color:black'><b>Uncertain </b></span>";
      ColorCoding.innerHTML += "<span style='color:lightskyblue'>' ███'</span><span style='color:black'><b>No Image File </b></span>";
      ColorCoding.innerHTML +=  "&nbsp&nbsp<button type='button', id='MoveToCellButtonNext', class='MoveToCellButton' onclick = 'MoveGridViewPortPosition(-1)' title = 'Move to Previous Problem Cell'>"+
                                "<img class='img' src='BP Appearance/MoveBackToCell.png' alt='Some png'></img></button>"+
                                "&nbsp<button type='button', id='MoveToCellButtonPrev', class='MoveToCellButton' onclick = 'MoveGridViewPortPosition(+1)' title = 'Move to Next Problem Cell'>"+
                                "<img class='img' src='BP Appearance/MoveToCell.png' alt='Some png'></img></button><span style='color:black'><b>&nbsp&nbspMove To Highlighted Cells</b></span>"
      CSVControls.insertBefore(ColorCoding, CSVControls.firstChild);
    }else{
      document.getElementById("ColorCodingCSVProblems").style.visibility = "visible";
      let collection = document.getElementsByClassName("MoveToCellButton");
      for( let m=0; m<collection.length; m++){collection[m].style.visibility = "visible"};        
    }
    // to (just) report image file search, don't expose all the controls on the image processing window
    document.getElementById("CompressImagesInCSVButton").style.display = 'none';
    document.getElementById("CompressAllImagesButton").style.display = 'none';
    document.getElementById("CompressImagesButton").style.display = 'none';
    document.getElementById("ImageCompressionLevel").style.display = 'none';
    document.getElementById("StopCompressImagesButton").style.display = 'none';
    document.getElementById("CancelCompressImagesButton").style.display = 'none';
    document.getElementById("ZIPImageProcessingButton").style.display = 'none';
    document.getElementById("SaveImageProcessingButton").style.display = 'none';

    ImageCompressionAlbumSaved = true;    //we're borrowing the compress image window. this prevents it from checking whether image files have been saved
    OpenCompressImages()
    document.getElementById("OpenDropDownMenu").click();  
  }  
  function MoveGridViewPortPosition(direction){
    if(!gridCellsWithProblems[0]){gridCellsWithProblems[0]=0}
    if(!gridCellsWithMissingImages[0]){gridCellsWithMissingImages[0]=0}
    TotalNumberOfProblemCells = gridCellsWithProblems[0]+gridCellsWithMissingImages[0];
    
    CurrentProblemCellNumberInList = CurrentProblemCellNumberInList + Number(direction);
    if(CurrentProblemCellNumberInList > TotalNumberOfProblemCells){CurrentProblemCellNumberInList=1};
    if(CurrentProblemCellNumberInList < 1){CurrentProblemCellNumberInList=TotalNumberOfProblemCells};

    if(CurrentProblemCellNumberInList <= gridCellsWithProblems[0]){
      // problem cell number is within 1st list
      grid.setViewportPosition({x:gridCellsWithProblems[CurrentProblemCellNumberInList][0],y:gridCellsWithProblems[CurrentProblemCellNumberInList][1]});
      grid.setCellCursorPosition({x:gridCellsWithProblems[CurrentProblemCellNumberInList][0],y:gridCellsWithProblems[CurrentProblemCellNumberInList][1]});
    }else{
      // problem cell number is within missing image list
      let PlaceInList = CurrentProblemCellNumberInList-gridCellsWithProblems[0];
      grid.setViewportPosition({x:gridCellsWithMissingImages[PlaceInList][0],y:gridCellsWithMissingImages[PlaceInList][1]});
      grid.setCellCursorPosition({x:gridCellsWithMissingImages[PlaceInList][0],y:gridCellsWithMissingImages[PlaceInList][1]});      
    }
  }
  function CloseCSVFileHeaders(){
    document.getElementById("AvailableCSVFieldElementHeaders").style.visibility = "hidden";
    document.getElementById("CSVFileChecking").style.visibility = "hidden";
  }
  function SaveCSVFileAnalysis(){
    CheckCSVFileAnalysisReportText = CheckCSVFileAnalysisReportText.replace(/<br>/gi, String.fromCharCode(13));
    CheckCSVFileAnalysisReportText = CheckCSVFileAnalysisReportText.replace(/<strong>/gi, "");
    CheckCSVFileAnalysisReportText = CheckCSVFileAnalysisReportText.replace(/<\/strong>/gi, "");
    CheckCSVFileAnalysisReportText = CheckCSVFileAnalysisReportText.replace(/<em>/gi, "");
    CheckCSVFileAnalysisReportText = CheckCSVFileAnalysisReportText.replace(/<\/em>/gi, "");
    download(CheckCSVFileAnalysisReportText, "BookPageantCSVFileCheckingReport.txt","text/plain");
  }
  function CloseCSVFileCheckingResults(){
    CSVFieldElements = document.getElementById("CSVAnalysis");
    //clear any existing CSV results
    while (CSVFieldElements.hasChildNodes()) {
      CSVFieldElements.removeChild(CSVFieldElements.firstChild);
    }
    CloseTextModify();
    CSVFieldElements = document.getElementById("CSVFileChecking");    
    CSVFieldElements.style.visibility = "hidden";
    document.getElementById("CSVFileChecking").style.visibility = "hidden";
  }
  function ClearCSVFileChecingResults(){
    Analysis = document.getElementById('CSVAnalysis');
    Analysis.innerHTML = CSVFileInitialAnalysis;
  }
  function ResetCSVHeaderChecking(){
    for(let m=1; m<=NumberOfHeadersChoosenForCSVChecking[0]; m++){
      let FieldNumber = NumberOfHeadersChoosenForCSVChecking[m]; 
      if(FieldNumber > 0){
        CSVFieldElement = document.getElementById('CSVFieldElement'+FieldNumber.toString())
        let Index = CSVFieldElement.innerHTML.indexOf('<small>');
        if(Index != 0){CSVFieldElement.innerHTML = CSVFieldElement.innerHTML.substring(0,Index)};
        CSVFieldElement.style.color = 'black';
        CSVFieldElement.style.fontWeight = 'normal'        
        CSVFieldElement.name = 'notrequired';
        // CSVFieldElement.style.backgroundColor = '';
        RequiredDataFields[FieldNumber] = false;
        NumberOfHeadersChoosenForCSVChecking[m] = 0;
      }
    }
    NumberOfHeadersChoosenForCSVChecking[0] = 3;
    NumberOfHeadersChoosenForCSVChecking[1] = 0;
    NumberOfHeadersChoosenForCSVChecking[2] = 0;
    NumberOfHeadersChoosenForCSVChecking[3] = 0;
    document.getElementById("CSVFieldElementHeader").innerHTML = TextForinnerHTMLofCSVCheckterButton;
    document.getElementById("ContinueCSVFileCheckingButton").disabled = true;
    document.getElementById("ContinueCSVFileCheckingButton").style.opacity = 0.25;    
    document.getElementById("ContinueCSVFileCheckingButton").style.cursor = 'default';
    document.getElementById("CSVFileChecking").style.visibility = "hidden";
  }
//#endregion

// CSV file data viewing & checking
//#region
  function OpenCSVDataViewing(TextToEdit, CellLocation){
    dragElement(document.getElementById("CSVFileDataViewingDrag"));
    document.getElementById("CSVFileDataViewing").style.visibility = "visible";
    // document.getElementById("CSVFileDataViewingHeader").innerHTML = meta.fields[HeaderNumber];
    document.addEventListener('mouseup', HandleMouseUP)
    document.getElementById("CSVData").style.display = 'block'
    TextToEdit = TextToEdit.replace(/[|]/g,"<br>").split( /<i>/ ).reduce( (a,b, i) => i %2 == 0 ? a + "</i>" + b : a + "<i>" + b );
    document.getElementById("CSVData").innerHTML = TextToEdit;
    document.getElementById("CSVData").contenteditable = true;
    document.getElementById("CSVData").spellcheck = true;
    document.getElementById("CSVData").focus();
    TextForSpellchecking = TextToEdit
    CellForSpellchecking = CellLocation
    //enable image preview function iff test is a image file name with appropriate extension
    let ImageFileName = TextToEdit;
    let re = /(?:\.([^.]+))?$/;
    let FileExtension = re.exec(ImageFileName)[1];
    if(!FileExtension){
      document.getElementById("PreviewImageCSVFileDataViewingButton").disabled = true;
      document.getElementById("PreviewImageCSVFileDataViewingButton").style.opacity = 0.35;
      document.getElementById("PreviewImageCSVFileDataViewingButton").style.pointerEvents = 'none';      
    }else{
      if( FileExtension.toLowerCase()=="jpg" || 
      FileExtension.toLowerCase()=="tif" ||
      FileExtension.toLowerCase()=="png" ||
      FileExtension.toLowerCase()=="bmp" ){
        document.getElementById("PreviewImageCSVFileDataViewingButton").disabled = false;       
        document.getElementById("PreviewImageCSVFileDataViewingButton").style.opacity = 1.00;  
        document.getElementById("PreviewImageCSVFileDataViewingButton").style.pointerEvents = 'auto';        
      }else{
        document.getElementById("PreviewImageCSVFileDataViewingButton").disabled = true;
        document.getElementById("PreviewImageCSVFileDataViewingButton").style.opacity = 0.35;  
        document.getElementById("PreviewImageCSVFileDataViewingButton").style.pointerEvents = 'none';
      }
    }
  }
  function HandleMouseUP(){  
    if(window.getSelection().toString().length > 0){
      let highlightedText = window.getSelection().toString();
      if(!ListeningForSelection){return};
      CurrentHighlightedText = window.getSelection();
      let posX = event.clientX + 5;
      let posY = event.clientY + 20;
      TextModifyStartIndex = highlightedText.anchorOffset;
      TextModifyEndIndex = highlightedText.focusOffset;
      let TextModElement = document.getElementById("CSVDataTextModify");
      TextModElement.style.visibility = "visible";
      TextModElement.style.position = "fixed";
      TextModElement.style.top  = posY.toString()+"px";
      TextModElement.style.left = posX.toString()+"px";
      ListeningForSelection = false;
    }
  }  
  function ItalizieText(){
    document.execCommand('italic');
    // grid.setCellValues(CurrentSpreadsheetCellSpecification, document.getElementById("CSVData").innerHTML);
    document.getElementById("CSVDataTextModify").style.visibility = "hidden";
    ListeningForSelection = true;
    SpreadsheetCellDataChanged = true;
    SpreadSheetHasChanged = true;        
  }
  function BoldText(){
    document.execCommand('bold');
    // grid.setCellValues(CurrentSpreadsheetCellSpecification, document.getElementById("CSVData").innerHTML);    
    document.getElementById("CSVDataTextModify").style.visibility = "hidden";
    ListeningForSelection = true; 
    SpreadsheetCellDataChanged = true;       
    SpreadSheetHasChanged = true;        
  }
  function SubscriptText(){
    document.execCommand('subscript');
    // grid.setCellValues(CurrentSpreadsheetCellSpecification, document.getElementById("CSVData").innerHTML);
    document.getElementById("CSVDataTextModify").style.visibility = "hidden";
    ListeningForSelection = true; 
    SpreadsheetCellDataChanged = true;       
    SpreadSheetHasChanged = true;        
  }
  function SuperscriptText(){
    document.execCommand('superscript');
    // grid.setCellValues(CurrentSpreadsheetCellSpecification, document.getElementById("CSVData").innerHTML);    
    document.getElementById("CSVDataTextModify").style.visibility = "hidden";
    ListeningForSelection = true; 
    SpreadsheetCellDataChanged = true;       
    SpreadSheetHasChanged = true;        
  }
  function UnderlineText(){
    document.execCommand('underline');
    // grid.setCellValues(CurrentSpreadsheetCellSpecification, document.getElementById("CSVData").innerHTML);    
    document.getElementById("CSVDataTextModify").style.visibility = "hidden";
    ListeningForSelection = true; 
    SpreadsheetCellDataChanged = true;       
    SpreadSheetHasChanged = true;        
  }
  function PlainText(){
    document.execCommand('removeFormat');
    // grid.setCellValues(CurrentSpreadsheetCellSpecification, document.getElementById("CSVData").innerHTML);    
    document.getElementById("CSVDataTextModify").style.visibility = "hidden";
    ListeningForSelection = true; 
    SpreadsheetCellDataChanged = true;       
    SpreadSheetHasChanged = true;        
  }  
  function UndoLastChangeToText(){
    document.execCommand('undo');
    // grid.setCellValues(CurrentSpreadsheetCellSpecification, document.getElementById("CSVData").innerHTML);    
    document.getElementById("CSVDataTextModify").style.visibility = "hidden";
    ListeningForSelection = true; 
    SpreadsheetCellDataChanged = true;       
    SpreadSheetHasChanged = true;    
  }
  function CloseTextModify(){
    document.getElementById("CSVDataTextModify").style.visibility = "hidden";
    ListeningForSelection = true; 
    SpreadsheetCellDataChanged = true;       
  }

  function FormatUnformatCSVFileData(element){
    DisplayCSVDataFormatted = !DisplayCSVDataFormatted;
    // change button image
    if(DisplayCSVDataFormatted == false) {
      element.innerHTML="<img class=\"img\" src=\"BP Appearance/unformat.png\" alt=\"Some png\">"
       document.getElementById("CSVData").innerHTML = JSON.stringify(data[CurrentCSVDataRow-1][meta.fields[CurrentCSVDataHeader]]);
    }else{
      element.innerHTML="<img class=\"img\" src=\"BP Appearance/format.png\" alt=\"Some png\">"
      document.getElementById("CSVData").innerHTML = data[CurrentCSVDataRow-1][meta.fields[CurrentCSVDataHeader]];
    }  //.replace(/[|]/g,"<br>").split( /<i>/ ).reduce( (a,b, i) => i %2 == 0 ? a + "</i>" + b : a + "<i>" + b );    }
  }

  function SpellCheckCSVFileData(){
    let dictionary = new Typo("en_US", false, false, { dictionaryPath: "./Libraries/JS Libraries/dictionaries" })
    let ListOfWords = [];
    ListOfWords = TextForSpellchecking.replace(/<i>/gi,"").replace(/[.,;:!?'"]/g,"").replace(/[|\-]+/g," ").replace("\u2013"," ").replace("\u2014"," ").split(" ");
    // ListOfWords = TextForSpellchecking.replace(/<i>/gi,"").split(" ");
    // ListOfWords = TextForSpellchecking.split(" ");
    NumberOfWordsWithSpellQuestion = 0;
    WordLoop: 
    for( let n=0; n<ListOfWords.length; n++){
      if( dictionary.check(ListOfWords[n])==false ){
        ListOfWords[n] = "<span><u>"+ListOfWords[n]+"</u></span>";
        // let suggestion = dictionary.suggest(ListOfWords[n])
        NumberOfWordsWithSpellQuestion++
      }
    }
    document.getElementById("CSVData").innerHTML = ListOfWords.join(' ');
    console.log(NumberOfWordsWithSpellQuestion)
  }
  function FindNextSpellCheckErrorCSVFileData(){
    RowLoop:
    for( let m=CurrentCSVRowSpellQuestion+1; m<=CurrentCSVDataRowMax; m++){
      if(CSVDataSpellClean[CurrentCSVDataHeader][m] == -1){
        CurrentCSVDataRow = m
        document.getElementById("CSVFileRow").value = CurrentCSVDataRow;
        if(DisplayCSVDataFormatted){
          document.getElementById("CSVData").innerHTML = data[CurrentCSVDataRow-1][meta.fields[CurrentCSVDataHeader]].replace(/[|]/g,"<br>").split( /<i>/ ).reduce( (a,b, i) => i %2 == 0 ? a + "</i>" + b : a + "<i>" + b );
        }else{
          document.getElementById("CSVData").innerText = data[CurrentCSVDataRow-1][meta.fields[CurrentCSVDataHeader]];      
        }
        document.getElementById("CSVData").contenteditable = true;
        document.getElementById("CSVData").spellcheck = true;
        document.getElementById("CSVData").focus();
        CurrentCSVRowSpellQuestion = m;        
        break RowLoop;   
      }
    }

  }

  function SaveChangesToCSVFileData(){
    grid.setCellValues(CurrentSpreadsheetCellSpecification, document.getElementById("CSVData").innerHTML);
    SpreadsheetCellDataChanged = false;
  }

  function PreviewImageInCSVFile(){
    let ImageFileName = TextToEdit;
    let re = /(?:\.([^.]+))?$/;
    let FileExtension = re.exec(ImageFileName)[1];
    if( FileExtension.toLowerCase()=="jpg" || 
        FileExtension.toLowerCase()=="tif" ||
        FileExtension.toLowerCase()=="png" ||
        FileExtension.toLowerCase()=="bmp" ){
          JustImageName = 'Images/'+ImageFileName.substring(ImageFileName.lastIndexOf('/') + 1);
          document.getElementById("DisplayCSVImagePreview").style.display = 'block';
          document.getElementById("DisplayCSVImagePreview").src = JustImageName;
          wheelzoom(document.getElementById("DisplayCSVImagePreview"));
          document.getElementById("CloseDisplayCSVImagePreviewButton").style.display = 'block';
    }
  }

  function CloseDisplayCSVImagePreview(){
    document.getElementById("DisplayCSVImagePreview").style.display = 'none';
    document.getElementById("DisplayCSVImagePreview").src = '';
    document.getElementById("DisplayCSVImagePreview").dispatchEvent(new CustomEvent('wheelzoom.destroy'));    
    document.getElementById("CloseDisplayCSVImagePreviewButton").style.display = 'none';    
  }

  async function SaveCSVfile(){
    data = grid.getData();
    let CSVfile = Papa.unparse(data, {
      delimiter: ",",
      header: true,
      skipEmptyLines: true,
    })
    if(CSVFileHandle != null){
      SaveCSVFileHandle = await window.showSaveFilePicker({startIn: CSVFileHandle[0], types: [{descriiption: 'CSV file', accept:{'text/plain':['.CSV']}}]});
    }else{
      SaveCSVFileHandle = await window.showSaveFilePicker({startIn: BookPageantHomeDirectoryHandle, types: [{descriiption: 'CSV file', accept:{'text/plain':['.CSV']}}]});
    }
    localforage.setItem('CSVFileHandle', CSVFileHandle);    
    let writable = await SaveCSVFileHandle.createWritable();
    await writable.write(CSVfile);
    await writable.close();
  }

  function CloseCSVFileDataViewing(){
    if(SpreadsheetCellDataChanged){
      // give the user the chance to save changes.
      document.getElementById('CustomDialog').showModal();
    }
    document.getElementById("CSVFileDataViewing").style.visibility = "hidden";
    document.getElementById("CSVDataTextModify").style.visibility = "hidden";
    document.removeEventListener('mouseup', HandleMouseUP);
    ListeningForSelection = true;    

  }
//#endregion

// CSV file mapping
//#region  
  async function ChooseCSVtoBPMappingFile(){
    UserAssistance( "Navigate to folder '/BookPageant/CSV files', choose a .csv file and Open.",
                    "File name has the following form: (Any Name).csv",
                    "",
                    "CSVFile.png", 20.0);
    document.getElementById("OpenDropDownMenu").click();
    try{BookPageantCSVFileHandle
      if( BookPageantCSVFileHandle[0]==null){ BookPageantCSVFileHandle[0]==BookPageantHomeDirectoryHandle}
      BookPageantCSVFileHandle = await window.showOpenFilePicker({startIn: BookPageantCSVFileHandle[0], types: [{descriiption: 'BP file', accept:{'text/plain':['.CSV']}}]});
    }catch{
      BookPageantCSVFileHandle = await window.showOpenFilePicker({startIn: BookPageantHomeDirectoryHandle, types: [{descriiption: 'BP file', accept:{'text/plain':['.CSV']}}]});
    }
    InputFile = await BookPageantCSVFileHandle[0].getFile();
    localforage.setItem('BookPageantCSVFileHandle', BookPageantCSVFileHandle);
    document.getElementById("BPlogoAndLibraryName").innerHTML = '<img class="Logo" src="./BP Appearance/BP Icon.png" alt="NotFound">' + "&nbsp;"+"&nbsp;"+"&nbsp;" + InputFile.name.trim();
    CSVtoBPMapping(InputFile);
  }
  function CSVtoBPMapping(InputFile){
		Reader = new FileReader();
		Reader.onload = function(){
      dataFromCSVfile = Reader.result;
      Papa.parse(dataFromCSVfile, {
        delimiter: ",",
        header: true,
        download: false,
        dynamicTyping: false,
        skipEmptyLines: true,
        complete: function (results) {
          data = results.data;
          errors = results.errors;
          meta = results.meta;
          CSVParseResults = results;
        }
      })
      CSVtoBPMappingMachine(meta);
    }  
    Reader.readAsText(InputFile);
  }

  function CSVtoBPMappingMachine(meta){
    if(CSVMappingAlreadyBuilt && CSVFileMapped.name==CSVFileToMap.name ){
      document.getElementById("AvailableCSVFieldElements").style.visibility = "visible";
      document.getElementById("BPFieldMapping").style.visibility = "visible";
      return};
      UserAssistance( "Drag CSV data headers from the left-hand list to appropriate cells in the 2nd column of the table on the right.",
      "Give the resulting mapped data a name using the Library Name input cell at the upper right of the table.",
      "When all the desired matches are made, use the Dowload CSV button to save the .bsp.csv file.",
      "DownloadCSVFile.png", 30.0);

    // guard lilbrary file name input from invalid characters
    document.getElementById("SetLibraryName").addEventListener("keypress", event => {
      if( event.key == "*"  || 
          event.key == "|"  ||
          event.key == "/"  || 
          event.key == "\\" ||
          event.key == "?"  || 
          event.key == ":"  || 
          event.key == '"'  || 
          event.key == "<"  || 
          event.key == ">"  
        ){event.preventDefault()}});
    CSVFileMapped = CSVFileToMap
    document.getElementById("BPFieldMapping").style.visibility = "visible";
    dragElement(document.getElementById("BPFieldMappingDrag"));
    //bring up the list of available BP field elements
    // BPFieldElements = document.getElementById("AvailableBPFieldElements");
    BPFieldElements = document.createElement("AvailableBPFieldElements");
    BPFieldElements.setAttribute( 'class'  , 'AvailableBPFieldElements');
    BPFieldElements.style.visibility = "inherit";
    NumBPFields = 0;
    //add hearders/labels/instructions
    BPFieldElementHeader = document.createElement('div');
    BPFieldElementHeader.setAttribute( 'id'  , 'BPFieldElementHeader1');
    BPFieldElementHeader.setAttribute( 'class'  , 'BPFieldElementHeader');
    BPFieldElementHeader.innerHTML = "Data Field in BookPageant"
    BPFieldElements.appendChild(BPFieldElementHeader);
    BPFieldElementHeader = document.createElement('div');
    BPFieldElementHeader.setAttribute( 'id'  , 'BPFieldElementHeader2');
    BPFieldElementHeader.setAttribute( 'class'  , 'BPFieldElementHeader');
    BPFieldElementHeader.innerHTML = "Drag CSV Header here to Fill a BookPageant Field"
    BPFieldElements.appendChild(BPFieldElementHeader);
    BPFieldElementHeader = document.createElement('div');
    BPFieldElementHeader.setAttribute( 'id'  , 'BPFieldElementHeader3');
    BPFieldElementHeader.setAttribute( 'class'  , 'BPFieldElementHeader');
    BPFieldElementHeader.innerHTML = "Label that BookPageant Uses"
    BPFieldElements.appendChild(BPFieldElementHeader);

    //add to the list of fields
    for (m = 1; m <= BP_Headers.length; m++) {
      NumBPFields = NumBPFields + 1;
      BPFieldElement = document.createElement('div');
      BPFieldElement.setAttribute( 'id'  , 'BPFieldElement'+NumBPFields.toString());
      BPFieldElement.setAttribute( 'class'  , 'BPFieldElement');
      BPFieldElement.innerHTML = BP_Headers[NumBPFields-1];
      BPFieldElement.addEventListener( "contextmenu", function(e){ 
        e.preventDefault();
        // FoldAllHelp();
        // OpenHelp('Importing CSV Data Into BookPageant',1);
        // OpenHelp('For a NEW .CSV file',2);
        // OpenHelp(e.target.innerHTML, 3);
      } );
      
      if( BP_Headers[NumBPFields-1] == "BP_Author last name first" ||
          BP_Headers[NumBPFields-1] == "BP_Title" ||
          BP_Headers[NumBPFields-1] == "BP_Date Published" ){
        BPFieldElement.style.color = "rgb(9,61,254)";
      }
      // CompositionElement.style.border = "1px dotted blue";
      BPFieldElements.appendChild(BPFieldElement)
      // make a drop-place for field from CSV header elements
      BPFieldElementPair = document.createElement('div');
      BPFieldElementPair.setAttribute( 'id'  , 'BPFieldElementPair'+NumBPFields.toString() );
      BPFieldElementPair.setAttribute( 'class'  , 'BPFieldElementPair');
      BPFieldElementPair.setAttribute( 'ondrop', 'DropCSVElement(event)');
      BPFieldElementPair.setAttribute( 'ondragover', 'allowCSVElementDrop(event)');
      BPFieldElementPair.setAttribute( 'ondragstart'  , 'dragCSVElement(event,'+NumBPFields.toString()+', "BPheader")' ) ;
      //tuck the name of the matching BP name element into the 'title' attribute of the BP drop field
      BPFieldElementPair.setAttribute( 'title' , 'BPFieldElementName'+NumBPFields.toString() );
      BPFieldElementPair.setAttribute( 'name' , '' );
      BPFieldElements.appendChild(BPFieldElementPair)
        //add to the list of fields
      BPFieldElementName = document.createElement('input');
      BPFieldElementName.setAttribute( 'type'  ,  'text');
      BPFieldElementName.setAttribute( 'id' , 'BPFieldElementName'+NumBPFields.toString() );
      BPFieldElementName.setAttribute( 'class'  , 'BPFieldElementName');
      BPFieldElementName.setAttribute( 'ondrop', 'PreventDrop(event)');
      BPFieldElementName.setAttribute( 'title' , 'BPFieldElementName'+NumBPFields.toString() );
      BPFieldElementName.disabled = true;
      BPFieldElementName.style.overflow = "hidden";
      BPFieldElements.appendChild(BPFieldElementName)
    }
    document.getElementById("BPFieldMapping").appendChild(BPFieldElements);

    //bring up the list of fields from the CSV file
    CSVFieldElements = document.getElementById("AvailableCSVFieldElements");
    //clear any existing CSV fields
    while (CSVFieldElements.hasChildNodes()) {
      CSVFieldElements.removeChild(CSVFieldElements.firstChild);
    }    
    CSVFieldElements.style.visibility = "visible";
    // header and controls for CSV file header list
    let FieldElementsControls = document.createElement('div');
    FieldElementsControls.setAttribute( 'id'  , 'FieldElementsControls');
    FieldElementsControls.setAttribute( 'class'  , 'FieldElementsControls');
    FieldElementsControls.style.zIndex = "1050";
    FieldElementsControls.style.visibility = "inherit";
    FieldElementsControls.style.backgroundColor = "var(--Gray10)";
    FieldElementsControls.style.justifyContent = "center";              //override default justification/spacing for this class
    // title and image for controls div
    CSVFieldElementsHeader = document.createElement('div');
    CSVFieldElementsHeader.setAttribute( 'id'  , 'CSVFieldElementHeader');
    CSVFieldElementsHeader.setAttribute( 'class'  , 'CSVFieldElementHeader');
    CSVFieldElementsHeader.setAttribute( 'draggable'  , 'false');
    CSVFieldElementsHeader.style.textAlign = "center";
    CSVFieldElementsHeader.innerHTML = "Data headers found in CSV file <br> "+
                                       "Drag CSV headers to BP data fields. <span style='color:rgb(9,61,254)';>Blue BP fields</span> MUST be filled<br><br>"+  //BPFieldElement.style.color = "rgb(9,61,254)"
                                       "<img class='imgX2' src='BP Appearance/continue.png' alt='Some png'>";
    FieldElementsControls.appendChild(CSVFieldElementsHeader);
    //add controls to FieldElements div
    CSVFieldElements.appendChild(FieldElementsControls);

    CSVFieldElementsContent = document.createElement('div');
    CSVFieldElementsContent.setAttribute( 'id'  , 'AvailableCSVFieldElementsList');
    CSVFieldElementsContent.setAttribute( 'class'  , 'AvailableCSVFieldElementsList');
    CSVFieldElementsContent.setAttribute( 'ondrop', 'DropCSVElement(event,'+NumBPFields.toString()+')');
    CSVFieldElementsContent.setAttribute( 'ondragover', 'allowCSVElementDrop(event)');
    CSVFieldElementsContent.style.visibility = "inherit"
    CSVFieldElementsContent.style.zIndex = 150;

    NumCSVFields = 0;
    //add to the list of fields
    for (m = 0; m < meta.fields.length; m++) {
      NumCSVFields = NumCSVFields + 1;
      // CompositionElementMappedToHeaderField[NumCompositionElements] = BPHeadersOfUserData[m];
      CSVFieldElement = document.createElement('div');
      CSVFieldElement.setAttribute( 'id'  , 'CSVFieldElement'+NumCSVFields.toString());
      CSVFieldElement.setAttribute( 'class'  , 'CSVFieldElement');
      CSVFieldElement.setAttribute( 'draggable'  , 'true');
      CSVFieldElement.setAttribute( 'ondragstart'  , 'dragCSVElement(event,'+NumCSVFields.toString()+', "CSVheader")' ) ;
      CSVFieldElement.style.top = (m*20).toString()+"px";
      CSVFieldElement.style.overflow = "hidden";
      CSVFieldElement.innerHTML = meta.fields[m];
      CSVFieldElementsContent.appendChild(CSVFieldElement)
    }  
    CSVFieldElements.appendChild(CSVFieldElementsContent);
  }
  function dragCSVElement(event, FieldNumber, source){
    event.dataTransfer.setData("text", event.target.id);
    event.dataTransfer.setData("number", FieldNumber);
    event.dataTransfer.setData("source", source);
  }
  function DropCSVElement(ev, el) {
    ev.preventDefault();
    //get the id of the BP field elment name that was stored in the 'name' attribute. NB this is done BEFORE it is overwritten by the mapped CSV header number
    nameInputElement =  ev.target.getAttribute( 'title' );
    BPInputElement = ev.target.getAttribute( 'name' );
    if( ev.dataTransfer.getData('source')!="BPheader" ){
      //get the user CSV field number from
      let CSVFieldNumber = ev.dataTransfer.getData("number");      
      //overwrite the 'name' attribute with the CSV number
      ev.target.setAttribute( 'name', CSVFieldNumber);
    }else{
      //we're moving headers within BP list. get the previous user CSV number
      //let CSVFieldNumber = ev.dataTransfer.getData("number");
      let CSVFieldNumber = ev.dataTransfer.getData("text").substr(15);
      //overwrite the 'name' attribute with the CSV number
      ev.target.setAttribute( 'name', CSVFieldNumber);
    }
    //get the target.id from the dataTransfer gizmo and use it to append the move CSV element to the BP pair element
    let data = ev.dataTransfer.getData("text");
    ev.target.appendChild(document.getElementById(data));
    //if target is a BPFieldElementPair, put the name of the CSV field into the BP field name input cell
    //otherwise the drag/drop is to move a CSV element back to the CSV list, so scan the BP name elements and remove the name of the CSV field.
    if(document.getElementById( nameInputElement )){
      document.getElementById( nameInputElement ).value = document.getElementById(data).innerHTML;
      document.getElementById( nameInputElement ).disabled = false;
      ev.target.removeAttribute('ondragover');
      //if we're moving from a BPheader to another BPheader reset source BP header cells
      if( ev.dataTransfer.getData('source')=="BPheader" ){
        BPFieldNumber = ev.dataTransfer.getData('number')
        let target = document.getElementById( 'BPFieldElementPair'+BPFieldNumber );
        target.setAttribute( 'ondragover', 'allowCSVElementDrop(event)');
        target.setAttribute( 'name', '');
        document.getElementById( 'BPFieldElementName'+BPFieldNumber ).innerHTML = "";
        document.getElementById( 'BPFieldElementName'+BPFieldNumber ).value = "";
      }
    }else{
      for (m = 1; m <= BP_Headers.length; m++) {
        BPnameField = document.getElementById( 'BPFieldElementName'+m.toString() );
        if(BPnameField.value == document.getElementById(data).innerHTML){
          BPnameField.value = null;
          //reset BP mapping cell to allow drag-n-drop
          let target = document.getElementById( 'BPFieldElementPair'+m.toString() );
          target.setAttribute( 'ondragover', 'allowCSVElementDrop(event)');
          break;
        }
      }
    }
    UserAssistance( "When all the desired CSV-to-BP matches are made they can be saved in a Mapping File.",
    "Use the File Download button to save the mapping. It will have the Library Name and a triple extension: '.csv.BPmapped.bp' .",
    "Use the File Upload button to rerieve a previiously saved CSV-to-BP mapping",
    "Save.png", 20.0);

  }
  function allowCSVElementDrop(ev) {
    ev.preventDefault();
  }
  function PreventDrop(ev, el){
    ev.preventDefault();
    var CSVFieldNumber = ev.dataTransfer.getData("number");
    ev.target.setAttribute( 'name', CSVFieldNumber);
    nameInputElement =  ev.target.getAttribute( 'title' );
    document.getElementById( nameInputElement ).value +=""
    return;
  }

  async function SaveBPFieldMapping(){
    DataToSaveBPCSVPairs = "";
    DataToSaveBPCSVNames = "";
    NumCSVFieldsMapped = 0
    NameOfLibrary = document.getElementById("SetLibraryName").value;
    if(NameOfLibrary==""){
      // alert("Give the library a name")
      DayPilot.Modal.alert("Library Name Field is Empty", { theme: "modal_rounded" });
      return
    }

    let CSVNumberMappedToBPHeader = [];
    let BPToCSVMap = [];
    for (m = 1; m <= BP_Headers.length; m++) {
      DataToSaveBPCSVPairs = DataToSaveBPCSVPairs + document.getElementById('BPFieldElementPair'+m.toString()).getAttribute( 'name' )+"|";
      DataToSaveBPCSVNames = DataToSaveBPCSVNames + document.getElementById('BPFieldElementName'+m.toString()).value+"|";
      if(document.getElementById('BPFieldElementPair'+m.toString()).getAttribute( 'name' ) != ""){
        NumCSVFieldsMapped++
        CSVNumberMappedToBPHeader[NumCSVFieldsMapped] = new Array(3);
        CSVNumberMappedToBPHeader[NumCSVFieldsMapped][0] = Number(document.getElementById('BPFieldElementPair'+m.toString()).getAttribute( 'name' ))-1;
        CSVNumberMappedToBPHeader[NumCSVFieldsMapped][1] = document.getElementById('BPFieldElementName'+m.toString()).value
        CSVNumberMappedToBPHeader[NumCSVFieldsMapped][2] = m-1
      }else{
        //if any of the 3 required BP headers are empty, mapping is incomplete, can't save. return
        if(m==2 || m==13 || m==16){
          DayPilot.Modal.alert("One of the 3 required fields is empty", { theme: "modal_rounded" });
          return
        }
      }
    }
    
    DataToSave = NameOfLibrary+"!"+JSON.stringify(CSVFileToMap)+"!"+DataToSaveBPCSVPairs+"!"+DataToSaveBPCSVNames;
    compressedData = LZString.compressToUTF16(DataToSave)

    if(BookPageantCSVMappingFileHandle){
      BookPageantCSVMappingFileHandle = await window.showSaveFilePicker({startIn: BookPageantCSVMappingFileHandle[0], types: [{descriiption: 'BP Mapping file', accept:{'text/plain':['.csv.BPmapped.bp']}}]});
    }else{
      BookPageantCSVMappingFileHandle = await window.showSaveFilePicker({startIn: BookPageantHomeDirectoryHandle, types: [{descriiption: 'BP Mapping file', accept:{'text/plain':['.csv.BPmapped.bp']}}]});
    }
    localforage.setItem('BookPageantCSVMappingFileHandle', BookPageantCSVMappingFileHandle);    
    let writable = await BookPageantCSVMappingFileHandle.createWritable();
    await writable.write(compressedData);
    await writable.close();
  }

  async function SaveBPCSV(){
    DataToSaveBPCSVPairs = "";
    DataToSaveBPCSVNames = "";
    NumCSVFieldsMapped = 0
    NameOfLibrary = document.getElementById("SetLibraryName").value;
    if(NameOfLibrary==""){
      // alert("Give the library a name")
      DayPilot.Modal.alert("Library Name Field is Empty", { theme: "modal_rounded" });
      return
    }

    let CSVNumberMappedToBPHeader = [];
    let BPToCSVMap = [];

    for (m = 1; m <= BP_Headers.length; m++) {
      DataToSaveBPCSVPairs = DataToSaveBPCSVPairs + document.getElementById('BPFieldElementPair'+m.toString()).getAttribute( 'name' )+"|";
      DataToSaveBPCSVNames = DataToSaveBPCSVNames + document.getElementById('BPFieldElementName'+m.toString()).value+"|";
      if(document.getElementById('BPFieldElementPair'+m.toString()).getAttribute( 'name' ) != ""){
        NumCSVFieldsMapped++
        CSVNumberMappedToBPHeader[NumCSVFieldsMapped] = new Array(3);
        // columen number (starting at zero) in input CSV file == CSV header number
        CSVNumberMappedToBPHeader[NumCSVFieldsMapped][0] = Number(document.getElementById('BPFieldElementPair'+m.toString()).getAttribute( 'name' ))-1;
        // name of header
        CSVNumberMappedToBPHeader[NumCSVFieldsMapped][1] = document.getElementById('BPFieldElementName'+m.toString()).value
        // BP header number
        CSVNumberMappedToBPHeader[NumCSVFieldsMapped][2] = m-1
      }else{
        //if any of the 3 required BP headers are empty, mapping is incomplete, can't save. return
        if(m==2 || m==13 || m==16){
          DayPilot.Modal.alert("One of the 3 required fields is empty", { theme: "modal_rounded" });
          return
        }
      }
    }
    
    //modify data retrieved from user CSV file, cast it into BP format, and download it.
    // 1st, pack a list of CSV mapped headers to BP_Headers
    for( let m = 1; m <= NumCSVFieldsMapped; m++){
      BPToCSVMap[m] = CSVNumberMappedToBPHeader[m][2]
    }
    // data key names must be changed. build an object of old and new key names
    let ListOfOldAndNewKeys = new Object;
    for( let m = 1; m <= NumCSVFieldsMapped; m++){
      let OldKey = CSVParseResults.meta.fields[CSVNumberMappedToBPHeader[m][0]]
      ListOfOldAndNewKeys[OldKey] = BP_Headers[BPToCSVMap[m]]
    }
    // meta contains needs to contain appropriate BP field names.
    CSVParseResults.meta.fields.length = NumCSVFieldsMapped
    for( let m = 1; m <= NumCSVFieldsMapped; m++){
      CSVParseResults.meta.fields[m-1] = BP_Headers[BPToCSVMap[m]];
    }
    //use key renameing function on each data set
    for (let k=0; k < CSVParseResults.data.length; k++){
      CSVParseResults.data[k] = renameKeys(CSVParseResults.data[k], ListOfOldAndNewKeys);
    }
    //now move all data record down 1, making room at the top for alternate field names
    for (let k = CSVParseResults.data.length; k > 0; k--){
      CSVParseResults.data[k] = CSVParseResults.data[k-1];
    }
    //while we're here check if record should be amended because of missing data. Pack CSVParseResults.data with only valid records.
    let kk = 0;
    for (let k = 1; k < CSVParseResults.data.length; k++){
      if(CSVParseResults.data[k]["BP_Author last name first"]==""){CSVParseResults.data[k]["BP_Author last name first"]=="Anonymous"}
      if(CSVParseResults.data[k]["BP_Title"]==""){CSVParseResults.data[k]["BP_Title"]=="N.S."}
      if(CSVParseResults.data[k]["BP_Date Published"]==""){CSVParseResults.data[k]["BP_Date Published"]=="N.S."}
      kk++
      CSVParseResults.data[kk] = CSVParseResults.data[k];
    }
    CSVParseResults.data.length = kk+1;
    //new 1st (zeroth) row in the data structure = alternatve field names
    CSVParseResults.data[0] = new Object;
    for( let m = 1; m <= NumCSVFieldsMapped; m++){
      OldKey = CSVParseResults.meta.fields[m-1];
      CSVParseResults.data[0][OldKey] = CSVNumberMappedToBPHeader[m][1];
    }

    compressedData = Papa.unparse(CSVParseResults);
    if(BookPageantBPCSVFileHandle){
      if(BookPageantBPCSVFileHandle==null) {BookPageantBPCSVFileHandle[0]=BookPageantHomeDirectoryHandle;}
      BookPageantBPCSVFileHandle = await window.showSaveFilePicker({startIn: BookPageantBPCSVFileHandle[0], types: [{descriiption: 'BP CSV file', accept:{'text/plain':['.bp.csv']}}]});
    }else{
      BookPageantBPCSVFileHandle = await window.showSaveFilePicker({startIn: BookPageantHomeDirectoryHandle, types: [{descriiption: 'BP CSV file', accept:{'text/plain':['.bp.csv']}}]});
    }
    localforage.setItem('BookPageantBPCSVFileHandle', BookPageantBPCSVFileHandle);    
    let writable = await BookPageantBPCSVFileHandle.createWritable();
    await writable.write(compressedData);
    await writable.close();
    
    function renameKeys(obj, newKeys) {
      const keyValues = Object.keys(obj).map(key => {
        const newKey = newKeys[key] || key;
        return { [newKey]: obj[key] };
      });
      return Object.assign({}, ...keyValues);
    }
  }
  
  async function RetrieveBPMapping(){
    UserAssistance( "Navigate to folder '/BookPageant/BP files'.",
                    "Choose a BPmapped template file: name has the form: (Libraryfilename).csv.BPmapped.bp",
                    "Note the triple extension '.csv.BPmapped.bp'",
                    "",10.0);
    if(BookPageantCSVMappingFileHandle){
      if(BookPageantCSVMappingFileHandle[0]==null) {BookPageantCSVMappingFileHandle[0]==BookPageantHomeDirectoryHandle;}
      BookPageantCSVMappingFileHandle = await window.showOpenFilePicker({startIn: BookPageantCSVMappingFileHandle[0], types: [{descriiption: 'BP Mapping file', accept:{'text/plain':['.csv.BPmapped.bp']}}]});
    }else{
      BookPageantCSVMappingFileHandle = await window.showOpenFilePicker({startIn: BookPageantHomeDirectoryHandle, types: [{descriiption: 'BP Mapping file', accept:{'text/plain':['.csv.BPmapped.bp']}}]});
      console.log('Handle null')      
    }
    InputFile = await BookPageantCSVMappingFileHandle[0].getFile();
    localforage.setItem('BookPageantCSVMappingFileHandle', BookPageantCSVMappingFileHandle);    
    // clear BP CSV mappings
    for (m = 1; m <= BP_Headers.length; m++) {
      document.getElementById('BPFieldElementPair'+m.toString()).setAttribute( 'name', "" );
      document.getElementById('BPFieldElementName'+m.toString()).value = "";
    }
    //read BP CSV mapping data from file
    // InputFile = FileList[0];
    CompressedReader = new FileReader();
    CompressedReader.onload = function(){
      compressedData = CompressedReader.result;
      decompressedData = LZString.decompressFromUTF16(compressedData);
      let SeparatedDecompressedData = decompressedData.split("!");
      NameOfLibrary = SeparatedDecompressedData[0];
      document.getElementById("SetLibraryName").value = NameOfLibrary;
      let CSVFileMappedAsString = SeparatedDecompressedData[1].split("|");
      let SavedBPCSVPairs = SeparatedDecompressedData[2].split("|");
      let SavedBPCSVNames = SeparatedDecompressedData[3].split("|");
      CSVFileMapped = JSON.parse(CSVFileMappedAsString);
      //put retrieved data in appropriate elements
      for (m = 1; m <= BP_Headers.length; m++) {
        if(SavedBPCSVPairs[m-1] != ""){
          CSVElementNumber = Number(SavedBPCSVPairs[m-1]);
          document.getElementById('BPFieldElementPair'+m.toString()).appendChild(document.getElementById('CSVFieldElement'+CSVElementNumber.toString()));
          document.getElementById('BPFieldElementPair'+m.toString()).setAttribute( 'name' , CSVElementNumber);
          document.getElementById('BPFieldElementName'+m.toString()).value = SavedBPCSVNames[m-1];
          document.getElementById('BPFieldElementName'+m.toString()).disabled = false;
        }
      }
      }
    CompressedReader.readAsText(InputFile);
  }

  function ClearAllBPFieldElements(){
    for (m = 1; m <= BP_Headers.length; m++) {
      target = document.getElementById( 'BPFieldElementName'+m.toString() );
      target.setAttribute( 'ondragover', 'allowCSVElementDrop(event)');      
      document.getElementById( 'BPFieldElementName'+m.toString() ).innerHTML = null;
      document.getElementById( 'BPFieldElementName'+m.toString() ).value = null;
      target = document.getElementById( 'BPFieldElementPair'+m.toString() );
      target.setAttribute( 'ondragover', 'allowCSVElementDrop(event)');      
      target.innerHTML = "";
    }
    CSVFieldElements = document.getElementById("AvailableCSVFieldElements");
    //clear any existing CSV fields
    while (CSVFieldElements.hasChildNodes()) {
      CSVFieldElements.removeChild(CSVFieldElements.firstChild);
    } 
    CSVFieldElements.style.visibility = "visible";
    // header and controls for CSV file header list
    let FieldElementsControls = document.createElement('div');
    FieldElementsControls.setAttribute( 'id'  , 'FieldElementsControls');
    FieldElementsControls.setAttribute( 'class'  , 'FieldElementsControls');
    FieldElementsControls.style.zIndex = "1050";
    FieldElementsControls.style.visibility = "inherit";
    FieldElementsControls.style.backgroundColor = "var(--Gray10)";
    FieldElementsControls.style.justifyContent = "center";              //override default justification/spacing for this class
    // title and image for controls div
    CSVFieldElementsHeader = document.createElement('div');
    CSVFieldElementsHeader.setAttribute( 'id'  , 'CSVFieldElementHeader');
    CSVFieldElementsHeader.setAttribute( 'class'  , 'CSVFieldElementHeader');
    CSVFieldElementsHeader.setAttribute( 'draggable'  , 'false');
    CSVFieldElementsHeader.style.textAlign = "center";
    CSVFieldElementsHeader.innerHTML = "Data headers found in CSV file <br> "+
                                       "Drag CSV headers to BP data fields. <span style='color:rgb(9,61,254)';>Blue BP fields</span> MUST be filled<br>"+  //BPFieldElement.style.color = "rgb(9,61,254)"
                                       "<img class='imgX2' src='BP Appearance/continue.png' alt='Some png'>";
    FieldElementsControls.appendChild(CSVFieldElementsHeader);
    //add controls to FieldElements div
    CSVFieldElements.appendChild(FieldElementsControls);

    CSVFieldElementsContent = document.createElement('div');
    CSVFieldElementsContent.setAttribute( 'id'  , 'AvailableCSVFieldElementsList');
    CSVFieldElementsContent.setAttribute( 'class'  , 'AvailableCSVFieldElementsList');
    CSVFieldElementsContent.setAttribute( 'ondrop', 'DropCSVElement(event,'+NumBPFields.toString()+')');
    CSVFieldElementsContent.setAttribute( 'ondragover', 'allowCSVElementDrop(event)');
    CSVFieldElementsContent.style.visibility = "inherit"
    CSVFieldElementsContent.style.zIndex = 150;

    NumCSVFields = 0;
    //add to the list of fields
    for (m = 0; m < meta.fields.length; m++) {
      NumCSVFields = NumCSVFields + 1;
      // CompositionElementMappedToHeaderField[NumCompositionElements] = BPHeadersOfUserData[m];
      CSVFieldElement = document.createElement('div');
      CSVFieldElement.setAttribute( 'id'  , 'CSVFieldElement'+NumCSVFields.toString());
      CSVFieldElement.setAttribute( 'class'  , 'CSVFieldElement');
      CSVFieldElement.setAttribute( 'draggable'  , 'true');
      CSVFieldElement.setAttribute( 'ondragstart'  , 'dragCSVElement(event,'+NumCSVFields.toString()+', "CSVheader")' ) ;
      CSVFieldElement.style.top = (m*20).toString()+"px";
      CSVFieldElement.style.overflow = "hidden";
      CSVFieldElement.innerHTML = meta.fields[m];
      CSVFieldElementsContent.appendChild(CSVFieldElement)
    }  
    CSVFieldElements.appendChild(CSVFieldElementsContent);

  }
  function CloseBPFieldElements(){
    CSVMappingAlreadyBuilt = true;
    CSVFileMapped = CSVFileToMap;
    document.getElementById("BPFieldMapping").style.visibility = "hidden";
    document.getElementById("AvailableCSVFieldElements").style.visibility = "hidden";
    document.getElementById("BPlogoAndLibraryName").innerHTML = '<img class="Logo" src="./BP Appearance/BP Icon.png" alt="NotFound">';
  }
//#endregion

// CSV spreadsheet
//#region
async function CSVspreadsheet(){
  // dismiss dropdown main F&F menu
  OpenDropDownMenu();
  // get csv file to open and display
  UserAssistance( "Navigate to folder '/BookPageant/CSV files', choose a BP-formatted .csv or a simple .csv file and Open.",
  "File name has the following form: (Any Name).csv or (Any Name).bp.csv",
  "",
  "CSVFile.png", 15.0);  
  try{CSVFileHandle
    if(CSVFileHandle[0]==null){CSVFileHandle[0]==BookPageantHomeDirectoryHandle}
    CSVFileHandle = await window.showOpenFilePicker({startIn: CSVFileHandle[0], types: [{descriiption: 'CSV file', accept:{'text/plain':['.CSV']}}]});
  }catch{
    CSVFileHandle = await window.showOpenFilePicker({startIn: BookPageantHomeDirectoryHandle, types: [{descriiption: 'CSV file', accept:{'text/plain':['.CSV']}}]});
  }
  InputFile = await CSVFileHandle[0].getFile();
  localforage.setItem('CSVFileHandle', CSVFileHandle);
  let SpreadsheetElement = document.getElementById("CSVspreadsheet");
  SpreadsheetElement.style.display = 'block';  
  // flag whether .csv or .bp.csv file has been opened enable/disable csv file checking accordingly.
  if( InputFile.name.trim().substr(-7, 7).toLowerCase() == '.bp.csv' ){
    BPCSVfileHasBeenLoaded = true;
    CSVfileHasBeenLoaded = false;
    document.getElementById("CheckSpreadsheetButton").disabled = false;
    document.getElementById("CheckSpreadsheetButton").style.opacity = '1.0'
    document.getElementById("CheckSpreadsheetButton").style.cursor = 'pointer';
    document.getElementById("SaveSpreadsheetCheckingButton").disabled = false;
    document.getElementById("SaveSpreadsheetCheckingButton").style.opacity = '1.0'
    document.getElementById("SaveSpreadsheetCheckingButton").style.cursor = 'pointer';    
  }else{
    BPCSVfileHasBeenLoaded = false;
    CSVfileHasBeenLoaded = true;
    document.getElementById("CheckSpreadsheetButton").disabled = false;
    document.getElementById("CheckSpreadsheetButton").style.opacity = '1.0'
    document.getElementById("CheckSpreadsheetButton").style.cursor = 'pointer';
    document.getElementById("SaveSpreadsheetCheckingButton").disabled = false;
    document.getElementById("SaveSpreadsheetCheckingButton").style.opacity = '1.0'
    document.getElementById("SaveSpreadsheetCheckingButton").style.cursor = 'pointer';    
  }
  NameOfLibrary = InputFile.name.trim();
  document.getElementById("BPlogoAndLibraryName").innerHTML = '<img class="Logo" src="./BP Appearance/BP Icon.png" alt="NotFound">' + "&nbsp;"+"&nbsp;"+"&nbsp;" + NameOfLibrary.trim();
  // csv file input, turn on image checking button on image processing page
  document.getElementById("CheckImagesButton").style.opacity = "1.0";
  document.getElementById("CheckImagesButton").style.cursor = "pointer";
  document.getElementById("CheckImagesButton").disabled = false; 

  ClearTimeline()
 
  document.body.style.cursor = "wait";
  Reader = new FileReader();
  Reader.onload = function(){
    dataFromCSVfile = Reader.result;
    DataHasBeenInput = false;
    Papa.parse(dataFromCSVfile, {
      delimiter: ",",
      header: true,
      download: false,
      dynamicTyping: false,
      skipEmptyLines: true,
      complete: function (results) {
        errors = results.errors;
        data = results.data;
        meta = results.meta;
        document.body.style.cursor = "initial";  
        SpreadsheetCellDataChanged = false;
        for (i = 0; i < data.length; i++) {
          if (i == 0) {
          } else {
            //while we're here, find and replace all occurances of "|" and LF char in the user's text with an HTML "<br>" to force a new line in the text display
            if(data[i]["BP_Description_1"]){
              var text = data[i]["BP_Description_1"];
              text = text.replace(/\|\r\n|\n|\r/g, "<br>");
              data[i]["BP_Description_1"] = text;
            }
            if(data[i]["BP_Description_2"]){
              var text = data[i]["BP_Description_2"];
              text = text.replace(/\|\r\n|\n|\r/g, "<br>");
              data[i]["BP_Description_2"] = text;
            }
            if(data[i]["BP_Reference(s)"]){
              var text = data[i]["BP_Reference(s)"];
              text = text.replace(/\|\r\n|\n|\r/g, "<br>");
              data[i]["BP_Reference(s)"] = text;
            }
            if(data[i]["BP_Collation/Pagination/Foliation"]){
              var text = data[i]["BP_Collation/Pagination/Foliation"];
              text = text.replace(/\|\r\n|\n|\r/g, "<br>");
              data[i]["BP_Collation/Pagination/Foliation"] = text;
            }
            if(data[i]["BP_Collation"]){
              var text = data[i]["BP_Collation"];
              text = text.replace(/\|\r\n|\n|\r/g, "<br>");
              data[i]["BP_Collation"] = text;
            }
            if(data[i]["BP_Pagination"]){
              var text = data[i]["BP_Pagination"];
              text = text.replace(/\|\r\n|\n|\r/g, "<br>");
              data[i]["BP_Pagination"] = text;
            }
            if(data[i]["BP_Foliation"]){
              var text = data[i]["BP_Foliation"];
              text = text.replace(/\|\r\n|\n|\r/g, "<br>");
              data[i]["BP_Foliation"] = text;
            }
            if(data[i]["BP_Plates"]){
              var text = data[i]["BP_Plates"];
              text = text.replace(/\|\r\n|\n|\r/g, "<br>");
              data[i]["BP_Plates"] = text;
            }
            if(data[i]["BP_BibliographicNotes"]){
              var text = data[i]["BP_BibliographicNotes"];
              text = text.replace(/\|\r\n|\n|\r/g, "<br>");
              data[i]["BP_BibliographicNotes"] = text;
            }
            if(data[i]["BP_Short or Common Title"]){
              ShortTitle[i] = data[i]["BP_Short or Common Title"]
            }else{
              ShortTitle[i] = data[i]["BP_Title"]; 
            }
          }
        }

        grid = new DataGridXL("CSVspreadsheetDisplay", {
          data: data,
          contextMenuItems: [
            "copy",
            "cut",
            "paste",
            "delete_rows",
            "insert_rows_before",
            "insert_rows_after",
            "delete_cols",
            "insert_cols_before",
            "insert_cols_after",
            {"label": "Edit Data",                 // custom item
              "method": function(){
              let CellLocation = this.getCellSelection();
              let CellContent = grid.getCellValue(CellLocation[0][0]);
              TextToEdit = CellContent;
              CurrentSpreadsheetCellSpecification = this.getCellSelection();
              CurrentSpreadsheetAllCellSpecification = this.selectAll();
              OpenCSVDataViewing(TextToEdit, CellLocation);
              }
            }
          ],
          colHeaderLabelFunction: function(index, coord, colRef, labels){
            if(!colRef.title){
              // store bad header column numbers
              untitledColIds.push(colRef.id);              
              return "<span style='color:red; background-color:aqua'><b>Missing Header</b></span>";
            }else{
              return "<span ><strong>"+colRef.title+"</strong></span>";
            }
          },
          colHeaderLabelAlign: "center"            
        });  
      }
    })
    grid.events.on("ready", function(){
      for(var i = 0; i < untitledColIds.length; i++){
        this._columnNodes[untitledColIds[i]-1].columnHeader.style.backgroundColor = "aqua"}
      SpreadSheetHasChanged = false;  
    });
    grid.events.on("any", function(gridEvent, eventName){
      // ignore any events that do not change document
      if(
        eventName.includes("ready") ||
        eventName.includes("selection") ||
        eventName.includes("resize") ||
        eventName.includes("contextmenu") ||
        eventName.includes("fullscreen")){
        return;
      }else{
        SpreadSheetHasChanged = true;
        console.log('spreadsheet has changed');
        return;
      }
    });
    // color code the header rows
    if(BPCSVfileHasBeenLoaded){
      // grid._setCellColors([{x:0,y:0},{x:meta.fields.length,y:0}], '#D7E5F0');
      setCellColors([{x:0,y:0},{x:meta.fields.length-1,y:0}], '#D7E5F0');
    }else{
      // grid._setCellColors([{x:0,y:0},{x:SpreadsheetHeaders.length,y:0}], '#D7E5F0');
      // setCellColors([{x:0,y:0},{x:SpreadsheetHeaders.length,y:0}], '#D7E5F0');
    }    
    // reaize grid columns to fit header text. put text onto canvas to get width in pixels.
    let canvas = document.createElement("canvas");
    let context = canvas.getContext("2d");
    let WidthOfHeader = 0;
    context.font = "16px Arial";
    for( let m=0; m<meta.fields.length; m++){
      if(meta.fields[m]=="" || meta.fields[m]==" "){
        WidthOfHeader = parseInt(context.measureText("Missing Header").width)+5;
      }else{
        WidthOfHeader = parseInt(context.measureText(meta.fields[m]).width)+5;
      }
      grid.resizeCols(m,WidthOfHeader);
    }
    canvas = null;
    grid.redraw();

    
  }
  Reader.readAsText(InputFile);

}

// this substitutes for the intrinsic function available in V1 of DataGridXL
function setCellColors(cellRange, color){
  if(!grid) return;
  var cellRef;
  for(var y = cellRange[0].y; y <= cellRange[1].y; y++){
    for(var x = cellRange[0].x; x <= cellRange[1].x; x++){
      cellRef = grid.getCellFromStore(grid.getRowIdByCoord(y), grid.getColIdByCoord(x));
      if(!cellRef){
        asdf =1
      }else{
        cellRef.color = color;
      }
    }
  }
  // do not forget to redraw the grid
  grid.redraw();
}
// ditto for this call color function
function clearCellColors(cellRange){
  setCellColors(cellRange, null);
}

function CloseCSVspreadsheet(){ 
  if(SpreadSheetHasChanged){
    // give the user the chance to save changes.
    document.getElementById('CustomDialog1').showModal();
  }

  document.getElementById("CSVspreadsheet").style.display = 'none';
  document.getElementById("CSVFileDataViewing").style.visibility = "hidden";
  document.getElementById("CSVDataTextModify").style.visibility = "hidden";
  document.getElementById("CSVspreadsheetDisplay").innerHTML = '';
  // turn off color coding, if present
  if(document.getElementById("ColorCodingCSVProblems")) {
    document.getElementById("ColorCodingCSVProblems").style.visibility = "hidden";
  }
  // csv file is being closed, turn oiff image checking button on image processing page
  document.getElementById("CheckImagesButton").style.opacity = "0.5";
  document.getElementById("CheckImagesButton").disabled = true;  
  document.getElementById("BPlogoAndLibraryName").innerHTML = '<img class="Logo" src="./BP Appearance/BP Icon.png" alt="NotFound">';  
}

//#endregion

// data display, time and alphabet ordering  -->
//#region

  // reveal and populate element expanded data window
  function OpenElementDetailData(InputElementNumber) {
    ElementNumber = InputElementNumber; 
    CurrentElementNumberExpandedData = ElementNumber
    document.getElementById('ExpandedElement').style.visibility = "visible";
    document.getElementById('ExpandedElement').style.display = "block";
    document.getElementById('ExpandedElement').setAttribute('name' , ElementNumber.toString())
    document.getElementById('BP_Image1').src = "data:image/gif;base64,R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7"; //data[ElementNumber]["BP_Image1"];
    document.getElementById('CurrentImageNumber').innerHTML = "0/" + NumberOfImagesWithEntry[ElementNumber].toString();

    // process (nearly) all BP headers
    FullAuthorNamePresent = false;
    for(let n=0; n < BP_Headers.length; n++){
      DataElement = document.getElementById(BP_Headers[n]);
      if(DataElement==null){continue};  // some headers (like multiple image headers) don't have a separate element declared for them
      DataElement.style.fontFamily = getComputedStyle(document.documentElement).getPropertyValue('--Font')
      DataElement.style.visibility = "hidden";
      // see if current BP header has data mapped into it or whether the data slot is empty; if so, don't show data header
      if( data[0][BP_Headers[n]] == null || data[ElementNumber][BP_Headers[n]] == '' || data[ElementNumber][BP_Headers[n]] == ' '){
        DataElement.style.display = 'none';
        continue};
      // don't use (additional) author name field if 1st one is present
      if(FullAuthorNamePresent && BP_Headers[n]=="BP_Author last name first"){
        DataElement.style.display = 'none';
        continue};
      // don't display edition data if it's zero or hasn't been given
      if( BP_Headers[n]=="BP_Edition Number" && (data[ElementNumber][BP_Headers[n]] < 1 || isNaN(data[ElementNumber][BP_Headers[n]]) || !isNaN(parseFloat(data[ElementNumber][BP_Headers[n]]) ))){
        DataElement.style.display = 'none';
        continue;
      }
      // some headers involve multi-volume data; if so, add a break after the data title
      if(
        (BP_Headers[n] == "BP_Description_1"                   ||
         BP_Headers[n] == "BP_Description_2"                   ||
         BP_Headers[n] == "BP_Reference(s)"                    ||
         BP_Headers[n] == "BP_Collation/Pagination/Foliation"  ||
         BP_Headers[n] == "BP_Collation"                       ||
         BP_Headers[n] == "BP_Pagination"                      ||
         BP_Headers[n] == "BP_Foliation"                       ||
         BP_Headers[n] == "BP_Plates"                          ||
         BP_Headers[n] == "BP_BibliographicNotes") && 
         (data[ElementNumber][BP_Headers[n]].includes("Vol") ||
          data[ElementNumber][BP_Headers[n]].includes("vol") )
      ){
        DataElement.innerHTML = "<strong>"+data[0][BP_Headers[n]]+"</strong>: <br>" + " " + data[ElementNumber][BP_Headers[n]];
      }else if(BP_Headers[n] != "BP_Digital Copy link"){
        DataElement.innerHTML = "<strong style='color:white'>"+data[0][BP_Headers[n]]+"</strong>:" + " " + data[ElementNumber][BP_Headers[n]];
      }else{
        DataElement.innerHTML='';
        newlink = document.createElement('a');
        newlink.className = 'HTTPlink';
        newlink.innerHTML = '<strong style="color:white">Link To Digital Copy: </strong> <br>' + data[ElementNumber][BP_Headers[n]];
        newlink.setAttribute('title', 'Link To Digital Copy');
        newlink.setAttribute('href', data[ElementNumber][BP_Headers[n]]);
        newlink.setAttribute('target', "_blank");
        DataElement.appendChild(newlink)        
      }
       // special cases, add a preceeding line break
      DataElement.style.display = 'block';
      DataElement.style.visibility = "visible";
      if(BP_Headers[n]=="BP_Author full name"){FullAuthorNamePresent = true};
    }
    // manage blank lines
    if( !data[ElementNumber]["BP_Short or Common Title"] ){
      document.getElementById("ExpandedDescriptionSeparatorBlank1").style.display = 'none';
    }
    if( !data[ElementNumber]["BP_Title English Translation"] ){
      document.getElementById("ExpandedDescriptionSeparatorBlank3").style.display = 'none';
    }
    if( !data[ElementNumber]["BP_Brief Summary of the Work"] ){
      document.getElementById("ExpandedDescriptionSeparatorBlank4").style.display = 'none';
    }
    if( !data[ElementNumber]["BP_Edition Number"] ){
      document.getElementById("ExpandedDescriptionSeparatorBlank5").style.display = 'none';
    }
    if( !data[ElementNumber]["BP_Publisher or Seller"] ){
      document.getElementById("ExpandedDescriptionSeparatorBlank6").style.display = 'none';
    }    
    if( !data[ElementNumber]["BP_Description_2"] ){
      document.getElementById("ExpandedDescriptionSeparatorBlank7").style.display = 'none';
    }
    if( !data[ElementNumber]["BP_Reference(s)"] ){
      document.getElementById("ExpandedDescriptionSeparatorBlank8").style.display = 'none';
    }
    if( !data[ElementNumber]["BP_Digital Copy link"] ){
      document.getElementById("ExpandedDescriptionSeparatorBlank8").style.display = 'none';
    }    
    // process any full data display formatting
    if(!ShowFullDataBibliographicDetails){
      for(let j=0; j<BP_HeadersBibliographic.length; j++){
        document.getElementById(BP_Headers[BP_HeadersBibliographic[j]]).style.display = 'none';
      }
    }
    if(!ShowFullDataDescriptions){
      for(let j=0; j<BP_HeadersDescriptive.length; j++){
        document.getElementById(BP_Headers[BP_HeadersDescriptive[j]]).style.display = 'none';
      }
    }
    if(!ShowFullDataReferences){
      for(let j=0; j<BP_HeadersReferences.length; j++){
        document.getElementById(BP_Headers[BP_HeadersReferences[j]]).style.display = 'none';
      }
    }
    if(!ShowFullDataLinks){
      for(let j=0; j<BP_HeadersDigital.length; j++){
        document.getElementById(BP_Headers[BP_HeadersDigital[j]]).style.display = 'none';
      }
    }

    if(ImageCompressionCollectionEmpty && BPfileHasBeenLoaded){
      // if there are no images in the BP file, remove image and its controls from the display
      document.getElementById("BP_Image1").style.display = 'none';
      document.getElementById("BP_ImageCaption").style.display = 'none';
      document.getElementById("ExpandedElementDataImageControl").style.display = 'none';
      return
    }

    //empty image caption
    document.getElementById('BP_ImageCaption').innerHTML = '';
    //check for images
    if( !data[ElementNumber]["BP_Image1"] ||                        //this element may have no image: not specified or never found/loaded
        ((ImageCompressionCollectionEmpty || BPCSVfileHasBeenLoaded) && data[CurrentElementNumberExpandedData]["BP_Image1"]=='')  ||
        (!(ImageCompressionCollectionEmpty || BPCSVfileHasBeenLoaded) && ImageCompressionCollection[CurrentElementNumberExpandedData][1]=='')){    
      document.getElementById("BP_Image1").style.display = 'none';
      document.getElementById("NextImageButton").style.display = 'none';
      document.getElementById("PrevImageButton").style.display = 'none';
      document.getElementById("CurrentImageNumber" ).style.display = 'none';
      document.getElementById('BP_ImageCaption').innerHTML = '';
      document.getElementById('ExpandedElement').scrollBy(0, -10000)
      return
    }else{
      // avoid multiple attachments of wheelzoom, otherwise it builds flashing/covering backgrounds
      document.getElementById("BP_Image1").dispatchEvent(new CustomEvent('wheelzoom.destroy'));
      wheelzoom(document.getElementById('BP_Image1'));
      document.getElementById("BP_Image1").style.display = 'block';
      document.getElementById("NextImageButton").style.display = 'block';
      document.getElementById("PrevImageButton").style.display = 'block';
      document.getElementById("CurrentImageNumber" ).style.display = 'block';      
      // set if we have to insert blank lines to get first descritpion section below image
      DataElement = document.getElementById("BibliographySeparatorBlank");
      VerticalOffset = pxTOvh(DataElement.offsetTop);
      if(VerticalOffset < pxTOvh(625)){
        DataElement.style.paddingTop = (pxTOvh(625)-VerticalOffset).toString()+"px";
      }else{
        DataElement.style.paddingTop = "0px";
      }
    }
    
    document.getElementById('ExpandedElement').scrollBy(0, -10000)
    // reset displayed image number
    CurrentImageDisplayedInExpandedData = 0;
    NextImageInExpandedData('+'); //start with the first image
  
    document.getElementById('NextImageButton').src = String.fromCharCode(175);
    document.getElementById('PrevImageButton').src = String.fromCharCode(174);
  }

  function CloseElementDetailData() {
    document.getElementById('ExpandedElement').style.visibility = "hidden";
    document.getElementById('ExpandedElement').style.display = "none";
    document.getElementById('BP_Image1').src = "data:image/gif;base64,R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7";
  }

  function NextImageInExpandedData(Step){
    if(Step=='+'){
      CurrentImageDisplayedInExpandedData++
    }else{
      CurrentImageDisplayedInExpandedData--
    }
    if(CurrentImageDisplayedInExpandedData > NumberOfImagesWithEntry[CurrentElementNumberExpandedData]){CurrentImageDisplayedInExpandedData=0}
    if(CurrentImageDisplayedInExpandedData < 0){CurrentImageDisplayedInExpandedData=NumberOfImagesWithEntry[CurrentElementNumberExpandedData]}
    if( ImageCompressionCollectionEmpty || BPCSVfileHasBeenLoaded){
      switch(CurrentImageDisplayedInExpandedData){
        case 0:
          document.getElementById('BP_Image1').src = "data:image/gif;base64,R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7";  //data[CurrentElementNumberExpandedData]["BP_Image1"];
          document.getElementById('CurrentImageNumber').innerHTML = "0/" + NumberOfImagesWithEntry[ElementNumber].toString();
          document.getElementById('BP_ImageCaption').innerHTML = '';
          break;
        case 1:
          document.getElementById('BP_Image1').src = data[CurrentElementNumberExpandedData]["BP_Image1"];
          document.getElementById('CurrentImageNumber').innerHTML = "1/" + NumberOfImagesWithEntry[ElementNumber].toString();
          document.getElementById('BP_ImageCaption').innerHTML = data[CurrentElementNumberExpandedData]["BP_Caption1"];
          break;
        case 2:
          document.getElementById('BP_Image1').src = data[CurrentElementNumberExpandedData]["BP_Image2"];
          document.getElementById('CurrentImageNumber').innerHTML = "2/" + NumberOfImagesWithEntry[ElementNumber].toString();
          document.getElementById('BP_ImageCaption').innerHTML = data[CurrentElementNumberExpandedData]["BP_Caption2"];
          break;
        case 3:
          document.getElementById('BP_Image1').src = data[CurrentElementNumberExpandedData]["BP_Image3"];
          document.getElementById('CurrentImageNumber').innerHTML = "3/" + NumberOfImagesWithEntry[ElementNumber].toString();
          document.getElementById('BP_ImageCaption').innerHTML = data[CurrentElementNumberExpandedData]["BP_Caption3"];
          break;
        case 4:
          document.getElementById('BP_Image1').src = data[CurrentElementNumberExpandedData]["BP_Image4"];
          document.getElementById('CurrentImageNumber').innerHTML = "4/" + NumberOfImagesWithEntry[ElementNumber].toString();
          document.getElementById('BP_ImageCaption').innerHTML = data[CurrentElementNumberExpandedData]["BP_Caption4"];
          break;
        case 5:
          document.getElementById('BP_Image1').src = data[CurrentElementNumberExpandedData]["BP_Image5"];
          document.getElementById('CurrentImageNumber').innerHTML = "5/" + NumberOfImagesWithEntry[ElementNumber].toString();
          document.getElementById('BP_ImageCaption').innerHTML = data[CurrentElementNumberExpandedData]["BP_Caption5"];
          break;
        }
    }else{
      switch(CurrentImageDisplayedInExpandedData){
        case 0:
          document.getElementById('BP_Image1').src = "data:image/gif;base64,R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7";  //data[CurrentElementNumberExpandedData]["BP_Image1"];
          document.getElementById('CurrentImageNumber').innerHTML = "0/" + NumberOfImagesWithEntry[ElementNumber].toString();
          document.getElementById('BP_ImageCaption').innerHTML = '';
          break;
        case 1:
          document.getElementById('BP_Image1').src = ImageCompressionCollection[CurrentElementNumberExpandedData][1]
          document.getElementById('CurrentImageNumber').innerHTML = "1/" + NumberOfImagesWithEntry[ElementNumber].toString();
          if(data[CurrentElementNumberExpandedData]["BP_Caption1"]){document.getElementById('BP_ImageCaption').innerHTML = data[CurrentElementNumberExpandedData]["BP_Caption1"]};
          break;
        case 2:
          document.getElementById('BP_Image1').src = ImageCompressionCollection[CurrentElementNumberExpandedData][2]
          document.getElementById('CurrentImageNumber').innerHTML = "2/" + NumberOfImagesWithEntry[ElementNumber].toString();
          if(data[CurrentElementNumberExpandedData]["BP_Caption2"]){document.getElementById('BP_ImageCaption').innerHTML = data[CurrentElementNumberExpandedData]["BP_Caption2"]};
          break;
        case 3:
          document.getElementById('BP_Image1').src = ImageCompressionCollection[CurrentElementNumberExpandedData][3]
          document.getElementById('CurrentImageNumber').innerHTML = "3/" + NumberOfImagesWithEntry[ElementNumber].toString();
          if(data[CurrentElementNumberExpandedData]["BP_Caption3"]){document.getElementById('BP_ImageCaption').innerHTML = data[CurrentElementNumberExpandedData]["BP_Caption3"]};
          break;
        case 4:
          document.getElementById('BP_Image1').src = ImageCompressionCollection[CurrentElementNumberExpandedData][4]
          document.getElementById('CurrentImageNumber').innerHTML = "4/" + NumberOfImagesWithEntry[ElementNumber].toString();
          if(data[CurrentElementNumberExpandedData]["BP_Caption4"]){document.getElementById('BP_ImageCaption').innerHTML = data[CurrentElementNumberExpandedData]["BP_Caption4"]};
          break;
        case 5:
          document.getElementById('BP_Image1').src = ImageCompressionCollection[CurrentElementNumberExpandedData][5]
          document.getElementById('CurrentImageNumber').innerHTML = "5/" + NumberOfImagesWithEntry[ElementNumber].toString();
          if(data[CurrentElementNumberExpandedData]["BP_Caption5"]){document.getElementById('BP_ImageCaption').innerHTML = data[CurrentElementNumberExpandedData]["BP_Caption5"]};
          break;
        }        
      }
  }

  function MakePDFOfExpandedData(){
    ElementNumber = document.getElementById('ExpandedElement').getAttribute('name')
    UserAssistance( "Navigate to folder '/BookPageant/BP files', choose a PDF format file and Open.",
    "File name has the following form: (A Library and Composition Name).PDFformat.bp",
    "Note the double extension '.PDFformat.bp'",
    "", 15.0);
    RetrieveCompositionBoard(CurrentElementNumberExpandedData.toString());
  }

  function FetchExpandedDataComposition(FileList){
    // set composition page
    //Compose();
    // call to retrieve composition file. optional 2nd parameter triggers PDF generation within RetrieveCompositionBoard() and when done, closes the composition page.
    RetrieveCompositionBoard(FileList, CurrentElementNumberExpandedData.toString());
    nowWhat = 1.0
  }
  
  async function ProcessDataChronologically(Task="Main") {
    if(!DataHasBeenInput){return};
    let ProgressElement = document.getElementById("BlinkingLoading");
    // disable alphabet span slider & enbable year span slider
    document.getElementById("ScaleSlider").disabled = false;
    document.getElementById("ScaleSlider").style.pointerEvents = 'auto';

    document.getElementById("AlphaSlider").disabled = true;
    document.getElementById("AlphaSlider").style.pointerEvents = 'none';
    // dim/bright display headers used for sorting
    document.getElementById('DateSortHeader').style.opacity = '1.0';    
    document.getElementById('AlphaSortHeader').style.opacity = '0.25';
    // move time line to the min year
    CurrentLeftPositionOfTimeline = MinYear;
    document.documentElement.style.setProperty("--ElementCellScale", ElementCellScale);
    ProgressElement.style.display = "block";
    ProgressElement.innerHTML="Building Timeline";
    ProgressElement.animate([{opacity:0.50},{opacity:0.75},{opacity:1},{opacity:0.75},{opacity:0.50}],{duration:1500,iterations:Infinity})
    document.body.style.cursor = "progress"
    //toggle display order buttons' and sliders' background colors
    document.getElementById("ProcessDataChronologically").style.background="var(--Gray0)";
    document.getElementById("ProcessDataAlphabetically").style.background="var(--Gray10)";
    document.getElementById("ScaleSlider").style.setProperty('--SliderThumbColor', 'var(--AlertColor)')
    document.getElementById("AlphaSlider").style.setProperty('--SliderThumbColor', 'var(--Gray40)')
    ChronologicalSort = true;
    AlphabeticalSort = false;
    localStorage.setItem("ChronologicalSort", ChronologicalSort.toString());
    localStorage.setItem("AlphabeticalSort", AlphabeticalSort.toString());
    // establish chronology timeline start, extent, steps, and markers
    SetNavigationScaleInYears();
    PlaceMarkersByDate();

    // initial scaling spreads elements across available screen width, minus the element size, 22px padding on the left and the 15px scroll bar width on the right
    // InitialScale = (screen.width - 400*ElementCellScale - 22 - 15) / (MaxYear - MinYear)
    InitialScale = (screen.width - 400*ElementCellScale ) / (MaxYear - MinYear)
    SpanOfYears = (screen.width ) / InitialScale
    CurrentSpanOfYears = SpanOfYears;
    ValueYearsSpan.innerHTML = SpanOfYears.toFixed(0)
    Order = -1;

    // clear elements
    let TimeLine = document.getElementById("MainTimeline");
    while (TimeLine.hasChildNodes()) {
      TimeLine.removeChild(TimeLine.firstChild);
    }
    // position of progress circle depends on function that called for element ordering
    if(Task=="Main"){
      document.getElementById("CircleProgressDiv").setAttribute( 'style', 'display: block');
      await sleep(20);
      circleProgress = new CircleProgress('.CircleProgressDiv',{constrain: true, textFormat: 'percent' });  
    }else if(Task=="Filter"){
      document.getElementById("FilterProgress").setAttribute( 'style', 'display: block');
      await sleep(20);
      circleProgress = new CircleProgress('.FilterProgress',{constrain: true, textFormat: 'percent' });  
    }
    
    circleProgress.max = 100;
    circleProgress.value = 0;
    LastPositionY = 0;
    DocFrag = new DocumentFragment();

    // for (let ii = 1; ii <= NumOfTimelineElements; ii++) {
      // let j = ChronologicalOrder[ii];
      
    for (let j = 1; j <= NumOfTimelineElements; j++) {
      let ii = ChronologicalOrder[j];

      //new time line element appened at div with ID MainTimeline.  NB: TimeLineElements' numberering orders them by PubDate.
      TimeLineElement = document.createElement('div');
      TimeLineElement.setAttribute('class', 'TimeLineElement');
      TimeLineElement.setAttribute('id', 'TimeLineElement'+ii.toString());
      TimeLineElement.style.visibility = 'hidden';

        // this section adds detection of right mouse up/down to temporarily magnify timeline element while right button is down (context menu is suppressed)
        TimeLineElement.addEventListener( "contextmenu", function(e){ 
          e.preventDefault();        
        });
        TimeLineElement.addEventListener( "mousedown", function(e){ 
          e.preventDefault();
          if(e.button == 2){
            let matrix = (window.getComputedStyle(document.getElementById('TimeLineElement'+ii.toString())).transform).split('(')[1].split(')')[0].split(',');
            let leftPos = (parseFloat(matrix[4])).toString().trim()+'px';
            let topPos = (parseFloat(matrix[5])).toString().trim()+'px';
            document.getElementById('TimeLineElement'+ii.toString()).style.transformOrigin = 'top left';
            document.getElementById('TimeLineElement'+ii.toString()).style.transform = 'translate('+leftPos+','+topPos+')'+' scale(1.5)';
            console.log('Chron right mouse down'+leftPos+topPos)
          }
        } );
        TimeLineElement.addEventListener( "mouseup", function(e){ 
          e.preventDefault();
          if(e.button == 2){
            let matrix = (window.getComputedStyle(document.getElementById('TimeLineElement'+ii.toString())).transform).split('(')[1].split(')')[0].split(',');
            let leftPos = (parseFloat(matrix[4])).toString().trim()+'px';
            let topPos = (parseFloat(matrix[5])).toString().trim()+'px';
            document.getElementById('TimeLineElement'+ii.toString()).style.transformOrigin = 'top left';
            document.getElementById('TimeLineElement'+ii.toString()).style.transform = 'translate('+leftPos+','+topPos+')'+' scale(1.0)';    
            // these two line are diagnostics
            // document.getElementById('TimeLineElement'+ii.toString()).style.border =  '3px solid seashell';
            // document.getElementById('TimeLineElement'+ii.toString()).style.zIndex =  '10';
            // 
          }
        });

      //to the time line element div just created, append its child div's.

      //button for opening expanded element data
      NewButton = document.createElement('button');
      NewButton.innerHTML = "&#xf139";
      NewButton.setAttribute("class", "ExpandElementButton");
      NewButton.setAttribute("title", "Expand to Full Data");
      NewButton.setAttribute("onclick", "OpenElementDetailData(" + ii.toString() + ")");
      NewButton.style.transform = 'translate('+ (ElementCellScale*400-46).toString()+'px, -3px)';
      NewButton.innerHTML = "<img class='imgForDataExpand' src='BP Appearance/Full.png' alt='Some png'>"
      TimeLineElement.appendChild(NewButton)

      //button for book marking element
      NewButton = document.createElement('button');
      NewButton.setAttribute("class", "BookmarkElementButton");
      NewButton.setAttribute("id", "BookmarkElementButton"+ ii.toString() + ")");
      NewButton.setAttribute("title", "Bookmark this Element");
      Order = Order + 1;
      NewButton.setAttribute("onclick", "BookmarkThisElement(" + ii.toString() + ")");
      NewButton.style.transform = 'translate('+ (ElementCellScale*400-45  ).toString()+'px, 26px)';
      NewButton.innerHTML = "<img class='img' src='BP Appearance/bookmarkFull.png' alt='Some png'>"
      TimeLineElement.appendChild(NewButton)

      NewDiv = document.createElement('div');
      NewDiv.setAttribute('class', 'ElementMarker');
      NewDiv.style.transform = 'translate(-13px, -10px)';
      NewDiv.textContent = String.fromCharCode(9670);
      TimeLineElement.appendChild(NewDiv)

      // add elements of summary data
      if(ShowSummaryNumberDisplay=='block'){
        EandTNewDiv = document.createElement('div');
        EandTNewDiv.setAttribute('class', 'ElementNumberAndTitle');
        EandTNewDiv.setAttribute('id', 'ElementNumberAndTitle'+ii.toString());
        NewDiv = document.createElement('div');
        NewDiv.setAttribute('class', 'ElementNumber');
        NewDiv.setAttribute('id', 'ElementNumber'+ii.toString());
        NewDiv.innerHTML = j.toString() + String.fromCharCode(160) + String.fromCharCode(8226) + String.fromCharCode(160) + String.fromCharCode(160);
        if(ShowSummaryBriefBookTitleDisplay=='block'){
          if( data[ii]["BP_Short or Common Title"] ){
            NewDiv.innerHTML += data[ii]["BP_Short or Common Title"].toUpperCase().bold();
          }else{
            NewDiv.innerHTML += data[ii]["BP_Title"].toUpperCase().bold();
          }
        }
        EandTNewDiv.appendChild(NewDiv);
        TimeLineElement.appendChild(EandTNewDiv);
      }

      if(ShowSummaryBriefBookTitleDisplay=='block'){
        NewDiv = document.createElement('div');
        NewDiv.setAttribute('class', 'BriefBookTitle');
        NewDiv.setAttribute('id', 'BriefBookTitle'+ii.toString());
        NewDiv.style.display = 'none';
        if( data[ii]["BP_Short or Common Title"] ){
          NewDiv.innerHTML = data[ii]["BP_Short or Common Title"];
        }else{
          NewDiv.innerHTML = data[ii]["BP_Title"];
        }
        TimeLineElement.appendChild(NewDiv);
      } 
      if(ShowSummaryAuthorDisplay=='block'){
        NewDiv = document.createElement('div');
        NewDiv.setAttribute('class', 'Author');
        NewDiv.setAttribute('id', 'Author'+ii.toString());
        if(data[ii]["BP_Author full name"]){
          NewDiv.innerHTML = data[ii]["BP_Author full name"];
        }else{
          NewDiv.innerHTML = data[ii]["BP_Author last name first"];
        }
        TimeLineElement.appendChild(NewDiv);
      }
      if(ShowSummaryPublicationYearDisplay=='block'){
        NewDiv = document.createElement('div');
        NewDiv.setAttribute('class', 'PublicationYear');
        NewDiv.setAttribute('id', 'PublicationYear'+ii.toString());
        NewDiv.innerHTML = data[ii]["BP_Date Published"];
        TimeLineElement.appendChild(NewDiv)
      }
      if(ShowSummaryBriefDescriptionDisplay=='block'){
        NewDiv = document.createElement('div');
        NewDiv.setAttribute('class', 'DescriptionSeparator');
        NewDiv.textContent = "--------------------------------------"
        TimeLineElement.appendChild(NewDiv);

        NewDiv = document.createElement('div');
        NewDiv.setAttribute('class', 'BriefDescription');
        NewDiv.setAttribute('id', 'BriefDescription'+ii.toString());
        if(data[ii]["BP_Brief Summary of the Work"]){
          NewDiv.innerHTML = data[ii]["BP_Brief Summary of the Work"];
        }else{
          NewDiv.innerHTML = "";
        }
        TimeLineElement.appendChild(NewDiv)
      }
      // append the completed time line element to the fragment
      DocFrag.appendChild(TimeLineElement);
      await sleep(0);
      circleProgress.value = 100*j/NumOfTimelineElements;
    }
    // append the entire fragment to the main time line
    TimeLine.appendChild(DocFrag)

    if(Task=="Main"){
      document.getElementById("CircleProgressDiv").style.display = "none";
    }else if(Task=="Filter"){
      document.getElementById("FilterProgress").style.visibility = "hidden";
    }
    
    // capture the order and  height of each element. NB: .offsetHeight isn't available until elements are appended to the main timeline
    for (let ii = 1; ii <= NumOfTimelineElements; ii++) {
      TimelineElementHeight[ii] = document.getElementById('TimeLineElement'+ii.toString()).offsetHeight + 10;
    }

    // calculate initial positions of elements based on sorting date
    LastPositionY = 0;
    for (let i = 0; i <= NumOfTimelineElements; i++) {
      if (i == 0) {continue};
      if( !DisplayElement[ChronologicalOrder[i]]){continue};
      CurrentYear = Year[ChronologicalOrder[i]];
      PositionX = (CurrentYear - TimelineStartYear) * InitialScale;
      // CurrentTimeLineElement = document.getElementById("TimeLineElement"+i.toString());
      CurrentTimeLineElement = document.getElementById("TimeLineElement"+ChronologicalOrder[i].toString());
      CurrentTimeLineElement.style.transform = "translate(" + PositionX.toString() + "px," + LastPositionY.toString() + "px)";
      CurrentTimeLineElement.style.visibility = "visible";
      ElementScalePositionYpixels[i] = LastPositionY;
      LastPositionY = LastPositionY + TimelineElementHeight[ChronologicalOrder[i]];
    }

    // finish initial data processing: compress time elements into cascading columns, establish scroll and navigate timelines, put element placemarkers on navigate timeline.
    //force a scaling change to set the navigation slider width
    YearSlider = localStorage.getItem("YearSlider");
    if(!YearSlider){
      ScaleOfYears("0.1");
    }else{
      ScaleOfYears(YearSlider);
    }
    document.getElementById("ScaleSlider").value = Number(YearSlider);
    document.body.style.cursor = "auto";

    // if stored, fetch the current value of timeline element styling
    let TimeElementBackgroundTransparency = localStorage.getItem("TimeElementBackgroundTransparency");
    if(TimeElementBackgroundTransparency != null){
      document.documentElement.style.setProperty('--TimeElementBackgroundTransparency', TimeElementBackgroundTransparency.toString(), 'important');
    }
    let TimeElementBackgroundColor = localStorage.getItem("TimeElementBackgroundColor");
    if(TimeElementBackgroundColor != null){
      document.documentElement.style.setProperty('--TimeElementBackgroundColor', TimeElementBackgroundColor.toString(), 'important');
    }
    let TimeElementTextColor = localStorage.getItem("TimeElementTextColor");
    if(TimeElementTextColor != null){
      document.documentElement.style.setProperty('--SummaryTextColor', TimeElementTextColor.toString(), 'important');
    }

    // update visibility of full data display component groups. NB: JASON parsing is required since localStorage converts booleans to strings (!)
    for( let m=1; m<=4; m++){
      switch (m){
        case 1:
          ShowFullDataBibliographicDetails = JSON.parse(localStorage.getItem("ShowFullDataBibliographicDetails"));
          if(ShowFullDataBibliographicDetails == null){ShowFullDataBibliographicDetails = true}
          ToggleFullDataComponent('BibliographicDetails', ShowFullDataBibliographicDetails)
          break;
        case 2:
          ShowFullDataDescriptions = JSON.parse(localStorage.getItem("ShowFullDataDescriptions"));
          if(ShowFullDataDescriptions == null){ShowFullDataDescriptions = true}
          ToggleFullDataComponent('Descriptions', ShowFullDataDescriptions)
          break;
        case 3:
          ShowFullDataReferences = JSON.parse(localStorage.getItem("ShowFullDataReferences"));
          if(ShowFullDataReferences == null){ShowFullDataReferences = true}          
          ToggleFullDataComponent('References', ShowFullDataReferences)
          break;
        case 4:
          ShowFullDataLinks = JSON.parse(localStorage.getItem("ShowFullDataLinks"));
          if(ShowFullDataLinks == null){ShowFullDataLinks = true}          
          ToggleFullDataComponent('Links', ShowFullDataLinks)
          break;
      }
    }
    ProgressElement.style.display = "none";
  }
  
  async function ProcessDataAlphabetically(Task="Main") {
    if(!DataHasBeenInput){return}
    let ProgressElement = document.getElementById("BlinkingLoading");
    // enable alphabet span slider & disbable year span slider    
    document.getElementById("ScaleSlider").disabled = true;
    document.getElementById("ScaleSlider").style.pointerEvents = 'none';

    document.getElementById("AlphaSlider").disabled = false;
    document.getElementById("AlphaSlider").style.pointerEvents = 'auto';
    // dim/bright display headers used for sorting
    document.getElementById('DateSortHeader').style.opacity = '0.25';
    document.getElementById('AlphaSortHeader').style.opacity = '1.0';
    // move time line to the letter A
    CurrentLeftPositionOfTimeline = MinAlphaNumber;    
    document.documentElement.style.setProperty("--ElementCellScale", ElementCellScale);    
    ProgressElement.style.display = "block";
    ProgressElement.innerHTML="Alphabetizing";
    ProgressElement.animate([{opacity:0.50},{opacity:0.75},{opacity:1},{opacity:0.75},{opacity:0.50}],{duration:1500,iterations:Infinity})    
    //toggle display order buttons' background
    document.getElementById("ProcessDataChronologically").style.background="var(--Gray10)";
    document.getElementById("ProcessDataAlphabetically").style.background="var(--Gray0)";
    document.getElementById("ScaleSlider").style.setProperty('--SliderThumbColor', 'var(--Gray40)')
    document.getElementById("AlphaSlider").style.setProperty('--SliderThumbColor', 'var(--AlertColor)')
    ChronologicalSort = false;
    AlphabeticalSort = true;
    localStorage.setItem("ChronologicalSort", ChronologicalSort.toString());
    localStorage.setItem("AlphabeticalSort", AlphabeticalSort.toString());

    // initial scaling spreads elements across available screen width, minus the element size, 22px padding on the left and the 15px scroll bar width on the right
    InitialScaleInv = (MaxAlphaNumber - MinAlphaNumber) / (screen.width - 400 - 22 - 15)
    SpanOfAlphabet = (screen.width - 22 - 15) * InitialScaleInv
    CurrentSpanOfAlphabet = SpanOfAlphabet;
    ValueAlphaSpan.innerHTML = SpanOfAlphabet.toFixed(0)
    Order = -1;
    // clear elements
    TimeLine = document.getElementById("MainTimeline");
    while (TimeLine.hasChildNodes()) {
      TimeLine.removeChild(TimeLine.firstChild);
    }
    // position of progress circle depends on function that called for element ordering
    if(Task=="Main"){
      document.getElementById("CircleProgressDiv").setAttribute( 'style', 'display: block');
      await sleep(20);
      circleProgress = new CircleProgress('.CircleProgressDiv',{constrain: true, textFormat: 'percent' });  
    }else if(Task=="Filter"){
      document.getElementById("FilterProgress").setAttribute( 'style', 'display: block');
      await sleep(20);
      circleProgress = new CircleProgress('.FilterProgress',{constrain: true, textFormat: 'percent' });  
    }    
    // establish alphabetic elements
    document.getElementById("CircleProgressDiv").setAttribute( 'style', 'display: block');
    await sleep(20);
    circleProgress = new CircleProgress('.CircleProgressDiv',{constrain: true, textFormat: 'percent' });  
    circleProgress.max = 100;
    circleProgress.value = 0;
    LastPositionY = 0;
    DocFrag = new DocumentFragment();

    for (let i = 1; i <= NumOfTimelineElements; i++) {
      //order by author last name
      let j = AlphabeticalOrder[i];
      //new time line element appened at div with ID MainTimeline
      TimeLineElement = document.createElement('div');
      TimeLineElement.setAttribute('class', 'TimeLineElement');
      TimeLineElement.setAttribute('id', 'TimeLineElement'+i.toString());
      TimeLineElement.style.visibility = 'hidden';

        // this section adds detions of right mouse up/down to temporarily magnify timeline element while right button is down (context menu is suppressed)
        TimeLineElement.addEventListener( "contextmenu", function(e){ 
          e.preventDefault();        
        });
        TimeLineElement.addEventListener( "mousedown", function(e){ 
          e.preventDefault();
          if(e.button == 2){
            let matrix = (window.getComputedStyle(document.getElementById('TimeLineElement'+i.toString())).transform).split('(')[1].split(')')[0].split(',');
            let leftPos = (parseFloat(matrix[4])).toString().trim()+'px';
            let topPos = (parseFloat(matrix[5])).toString().trim()+'px';
            document.getElementById('TimeLineElement'+i.toString()).style.transformOrigin = 'top left';
            document.getElementById('TimeLineElement'+i.toString()).style.transform = 'translate('+leftPos+','+topPos+')'+' scale(1.5)';
            console.log(i.toString()+' Alpha right mouse down '+leftPos+' '+topPos)
          }
        } );
        TimeLineElement.addEventListener( "mouseup", function(e){ 
          e.preventDefault();
          if(e.button == 2){
            let matrix = (window.getComputedStyle(document.getElementById('TimeLineElement'+i.toString())).transform).split('(')[1].split(')')[0].split(',');
            let leftPos = (parseFloat(matrix[4])).toString().trim()+'px';
            let topPos = (parseFloat(matrix[5])).toString().trim()+'px';
            document.getElementById('TimeLineElement'+i.toString()).style.transformOrigin = 'top left';
            document.getElementById('TimeLineElement'+i.toString()).style.transform = 'translate('+leftPos+','+topPos+')'+' scale(1.0)';        
          }
        });

      //button for opening expanded element data
      NewButton = document.createElement('button');
      NewButton.innerHTML = "&#xf139";
      NewButton.setAttribute("class", "ExpandElementButton");
      NewButton.setAttribute("title", "Expand to Full Data");
      NewButton.setAttribute("onclick", "OpenElementDetailData(" + i.toString() + ")");
      NewButton.style.transform = 'translate('+ (ElementCellScale*400-43).toString()+'px, -6px)';      
      NewButton.innerHTML = "<img class='imgForDataExpand' src='BP Appearance/Full.png' alt='Some png'>"
      TimeLineElement.appendChild(NewButton)
      //button for book marking element
      NewButton = document.createElement('button');
      NewButton.innerHTML = "&#xf139";
      NewButton.setAttribute("class", "BookmarkElementButton");
      NewButton.setAttribute("id", "BookmarkElementButton"+ i.toString() + ")");      
      NewButton.setAttribute("title", "Bookmark this Element");
      Order = Order + 1
      NewButton.setAttribute("onclick", "BookmarkThisElement(" + i.toString() + ")");
      NewButton.style.transform = 'translate('+ (ElementCellScale*400-45).toString()+'px, 24px)';      
      NewButton.innerHTML = "<img class='img' src='BP Appearance/bookmarkFull.png' alt='Some png'>"
      TimeLineElement.appendChild(NewButton)

      NewDiv = document.createElement('div');
      NewDiv.setAttribute('class', 'ElementMarker');
      NewDiv.style.transform = 'translate(-13px, -10px)';      
      NewDiv.textContent = String.fromCharCode(9670);
      TimeLineElement.appendChild(NewDiv)
      
      if(ShowSummaryNumberDisplay=='block'){
        EandTNewDiv = document.createElement('div');
        EandTNewDiv.setAttribute('class', 'ElementNumberAndTitle');
        EandTNewDiv.setAttribute('id', 'ElementNumberAndTitle'+i.toString());
        NewDiv = document.createElement('div');
        NewDiv.setAttribute('class', 'ElementNumber');
        NewDiv.setAttribute('id', 'ElementNumber'+i.toString());
        NewDiv.innerHTML = j.toString() + String.fromCharCode(160) + String.fromCharCode(8226) + String.fromCharCode(160) + String.fromCharCode(160);
        if(ShowSummaryBriefBookTitleDisplay=='block'){
          if( data[i]["BP_Short or Common Title"] ){
            NewDiv.innerHTML += data[i]["BP_Short or Common Title"].toUpperCase().bold();
          }else{
            NewDiv.innerHTML += data[i]["BP_Title"].toUpperCase().bold();
          }
        }
        EandTNewDiv.appendChild(NewDiv);
        TimeLineElement.appendChild(EandTNewDiv);
      }
      if(ShowSummaryBriefBookTitleDisplay=='block'){
        NewDiv = document.createElement('div');
        NewDiv.setAttribute('class', 'BriefBookTitle');
        NewDiv.setAttribute('id', 'BriefBookTitle'+i.toString());
        NewDiv.style.display = 'none';      
        if( data[i]["BP_Short or Common Title"] ){
          NewDiv.innerHTML = data[i]["BP_Short or Common Title"];
        }else{
          NewDiv.innerHTML = data[i]["BP_Title"];
        }
        TimeLineElement.appendChild(NewDiv);
      }
      if(ShowSummaryAuthorDisplay=='block'){
        NewDiv = document.createElement('div');
        NewDiv.setAttribute('class', 'Author');
        NewDiv.setAttribute('id', 'Author'+i.toString());
        if(data[j]["BP_Author full name"]){
          NewDiv.innerHTML = data[i]["BP_Author full name"];
        }else{
          NewDiv.innerHTML = data[i]["BP_Author last name first"];
        }      
        TimeLineElement.appendChild(NewDiv);
      }
      if(ShowSummaryPublicationYearDisplay=='block'){
        NewDiv = document.createElement('div');
        NewDiv.setAttribute('class', 'PublicationYear');
        NewDiv.setAttribute('id', 'PublicationYear'+i.toString());
        NewDiv.innerHTML = data[i]["BP_Date Published"];      
        TimeLineElement.appendChild(NewDiv)
      }
      if(ShowSummaryBriefDescriptionDisplay=='block'){
        NewDiv = document.createElement('div');
        NewDiv.setAttribute('class', 'DescriptionSeparator');
        NewDiv.textContent = "--------------------------------------"
        TimeLineElement.appendChild(NewDiv);

        NewDiv = document.createElement('div');
        NewDiv.setAttribute('class', 'BriefDescription');
        NewDiv.setAttribute('id', 'BriefDescription'+i.toString());
        if(data[i]["BP_Brief Summary of the Work"]){
          NewDiv.innerHTML = data[i]["BP_Brief Summary of the Work"];
        }else{
          NewDiv.innerHTML = "";
        }
        TimeLineElement.appendChild(NewDiv)
      }

      DocFrag.appendChild(TimeLineElement);
      await sleep(0);
      circleProgress.value = 100*i/NumOfTimelineElements;
    }
    // append the entire fragment to the main time line
    TimeLine.appendChild(DocFrag);

    if(Task=="Main"){
      document.getElementById("CircleProgressDiv").style.display = "none";
    }else if(Task=="Filter"){
      document.getElementById("FilterProgress").style.visibility = "hidden";
    }

    // calculate initial positions of elements based on alphabetical order of author's last name
    LastPositionY = 0;
    for (i = 0; i <= NumOfTimelineElements; i++) {
      if (i == 0) { continue }
      PositionX = (LastNameNumber[AlphabeticalSortPointer[i]] - 205891132094639) / InitialScaleInv; //NB:  205891132094639 = number for "A" = start of alpha line
      // PositionX = Math.floor(PositionX);
      CurrentTimeLilneElement =  document.getElementById("TimeLineElement"+AlphabeticalSortPointer[i].toString());
      CurrentTimeLilneElement.style.transform = "translate(" + PositionX.toString() + "px," + LastPositionY.toString() + "px)";
      TimelineElementHeight[AlphabeticalSortPointer[i]] = CurrentTimeLilneElement.offsetHeight + 10;
      ElementScalePositionYpixels[AlphabeticalSortPointer[i]] = LastPositionY
      LastPositionY = LastPositionY + TimelineElementHeight[AlphabeticalSortPointer[i]];
    }
    // finish initial data processing: compress time elements into cascading columns, establish scroll and navigate timelines, put element placemarkers on navigate timeline.
    SetNavigationScaleByAlpha();
    PlaceMarkersByAlpha();
    AlphabetSlider = localStorage.getItem("AlphabetSlider");
    if(!AlphabetSlider){
      AlphabetSlider = "0"
      ScaleOfAlphabet("0");
    }else{
      ScaleOfAlphabet(AlphabetSlider)
    }
    document.getElementById("AlphaSlider").value = Number(AlphabetSlider);
    // if stored, fetch the current value of timeline element background styling
    let TimeElementBackgroundTransparency = localStorage.getItem("TimeElementBackgroundTransparency");
    if(TimeElementBackgroundTransparency != null){
      document.documentElement.style.setProperty('--TimeElementBackgroundTransparency', TimeElementBackgroundTransparency.toString(), 'important');
    }
    let TimeElementBackgroundColor = localStorage.getItem("TimeElementBackgroundColor");
    if(TimeElementBackgroundColor != null){
      document.documentElement.style.setProperty('--TimeElementBackgroundColor', TimeElementBackgroundColor.toString(), 'important');
    }
    let TimeElementTextColor = localStorage.getItem("TimeElementTextColor");
    if(TimeElementTextColor != null){
      document.documentElement.style.setProperty('--SummaryTextColor', TimeElementTextColor.toString(), 'important');
    }

    // update visibility of full data display component groups. NB: JASON parsing is required since localStorage converts booleans to strings (!)
    for( let m=1; m<=4; m++){
      switch (m){
        case 1:
          ShowFullDataBibliographicDetails = JSON.parse(localStorage.getItem("ShowFullDataBibliographicDetails"));
          if(ShowFullDataBibliographicDetails == null){ShowFullDataBibliographicDetails = true}
          ToggleFullDataComponent('BibliographicDetails', ShowFullDataBibliographicDetails)
          break;
        case 2:
          ShowFullDataDescriptions = JSON.parse(localStorage.getItem("ShowFullDataDescriptions"));
          if(ShowFullDataDescriptions == null){ShowFullDataDescriptions = true}
          ToggleFullDataComponent('Descriptions', ShowFullDataDescriptions)
          break;
        case 3:
          ShowFullDataReferences = JSON.parse(localStorage.getItem("ShowFullDataReferences"));
          if(ShowFullDataReferences == null){ShowFullDataReferences = true}          
          ToggleFullDataComponent('References', ShowFullDataReferences)
          break;
        case 4:
          ShowFullDataLinks = JSON.parse(localStorage.getItem("ShowFullDataLinks"));
          if(ShowFullDataLinks == null){ShowFullDataLinks = true}          
          ToggleFullDataComponent('Links', ShowFullDataLinks)
          break;
      }
    }
    ProgressElement.style.display = "none";
  }

  // get slider data and modify positions of time line elements
  function ScaleOfYears(SliderValue) {
    localStorage.setItem("YearSlider", SliderValue.toString());
    Span = MaxYear - MinYear;
    for (m = NumSpanYearIncrements - 1; m >= 0; m--) {
      if (Span > ListSpanOfYears[m]) {
        LastInList = m;
        break;
      }
    }
    //given total span of years, get apppropiate Span of Years for the 10 stops on span-of-years slider
    switch (parseFloat(SliderValue)) {
      case 0:
        SpanOfYears = MaxYear - MinYear; break;
      case 0.1:
        SpanOfYears = ListSpanOfYears[LastInList]; break;
      case 0.2:
        SpanOfYears = ListSpanOfYears[LastInList - 1]; break;
      case 0.3:
        SpanOfYears = ListSpanOfYears[LastInList - 2]; break;
      case 0.4:
        SpanOfYears = ListSpanOfYears[LastInList - 3]; break;
      case 0.5:
        SpanOfYears = ListSpanOfYears[LastInList - 4]; break;
      case 0.6:
        SpanOfYears = ListSpanOfYears[LastInList - 5]; break;
      case 0.7:
        SpanOfYears = ListSpanOfYears[LastInList - 6]; break;
      case 0.8:
        SpanOfYears = ListSpanOfYears[LastInList - 7]; break;
      case 0.9:
        SpanOfYears = ListSpanOfYears[LastInList - 8]; break;
      default:
        SpanOfYears = ListSpanOfYears[LastInList - 9]; break;
    }
    Scale = (screen.width) / SpanOfYears;
    ValueYearsSpan.innerHTML = SpanOfYears.toFixed(0) + ' Years';
    LastPositionY = 0;
    NumCornersInCurrentColumn = 0;
    //first, each element's positiion in a single cascading chain is determined. The x-position depends on the element's year and the current value of (year) Scale.
    //the y-position depends on all the previous elements' heights.
    LastPositionY = 0;
    CurrentMaximumXPosition = 0;
    CurrentMinimumXPosition = 100000;
    for (i = 0; i <= NumOfTimelineElements; i++) {
      if (i == 0) {continue};
      CurrentYear = Year[ChronologicalOrder[i]];
      PositionX = ((CurrentYear - TimelineStartYear) * Scale );
      ElementScalePositionXpixels[ChronologicalOrder[i]] = PositionX;
      ElementScalePositionYpixels[ChronologicalOrder[i]] = LastPositionY;
      LastPositionY = LastPositionY + TimelineElementHeight[ChronologicalOrder[i]];
      if( DisplayElement[ChronologicalOrder[i]] ){
        if(i==NumOfTimelineElements){
          asdf = i;
        }
        if(PositionX+400*ElementCellScale > CurrentMaximumXPosition){CurrentMaximumXPosition = PositionX+400*ElementCellScale};
        if(PositionX     < CurrentMinimumXPosition){CurrentMinimumXPosition = PositionX};
      }
    }

    // now break the single cascade into multiple branches, setting each branch at the top and 
    //fitted so that the elements of each branch never overlap in x or y with those from a previous branch.
    // ResetY is the value substracted from an element's y-positiion to properly place it vertically (i.e, in y)
    //either because it starts a new cascade column OR there are interveneing elements not displayed (i.e. filetered out)
    ResetY = 0;
    NumElementsInCurrentColumn = 0;
    CurrentUpperCornerX = [];
    CurrentUpperCornerY = [];
    for (i = 1; i <= NumOfTimelineElements; i++) {
      if( !DisplayElement[ChronologicalOrder[i]]){
        //capture the latest y-coordinate of skipped element
        ResetY = ElementScalePositionYpixels[ChronologicalOrder[i+1]];
        continue;
      };
      //NB: using ResetY moves current column of elements to top (that is, Y = 0)
      document.getElementById("TimeLineElement"+ChronologicalOrder[i].toString()).style.transform = "translate(" + ElementScalePositionXpixels[ChronologicalOrder[i]].toString() + "px," + (ElementScalePositionYpixels[ChronologicalOrder[i]]-ResetY).toString() + "px)";
      NumElementsInCurrentColumn = NumElementsInCurrentColumn + 1;
      CurrentUpperCornerX[NumElementsInCurrentColumn] = ElementScalePositionXpixels[ChronologicalOrder[i]] + 400*ElementCellScale;
      CurrentUpperCornerY[NumElementsInCurrentColumn] = ElementScalePositionYpixels[ChronologicalOrder[i]]
      // capture absolute position of elements 
      CurrentElementScalePositionXpixels[ChronologicalOrder[i]] = ElementScalePositionXpixels[ChronologicalOrder[i]];
      CurrentElementScalePositionYpixels[ChronologicalOrder[i]] = ElementScalePositionYpixels[ChronologicalOrder[i]] - ResetY;

      // see if new column can be started. we need the next displayable element
      for(ii = i+1; ii <= NumOfTimelineElements; ii++ ){
        if( DisplayElement[ChronologicalOrder[ii]] ){
          NextValidi = ii;
          //move outerloop index to next valid i-1 (loop mechanism increments i) and leave the loop
          i=NextValidi-1;
          break;
        }else{
          // if element [ii] isn't displayed, increment ResetY by its y-size (height)
          ResetY += TimelineElementHeight[ChronologicalOrder[ii]];
        }
      }

      //move down the current column until there are now time element overlaps in x
      NewColumn = true;
      CurrentColumnLoop:
      for (let j = 1; j <= NumElementsInCurrentColumn; j++) {
        TimeElementLoop:
        for (let k = NextValidi; k <= NumOfTimelineElements; k++) {
          if( !DisplayElement[ChronologicalOrder[k]]){continue TimeElementLoop};
          if (ElementScalePositionXpixels[ChronologicalOrder[k]] < CurrentUpperCornerX[j]) {
            if (ElementScalePositionYpixels[ChronologicalOrder[k]] + TimelineElementHeight[ChronologicalOrder[k]] - ElementScalePositionYpixels[ChronologicalOrder[NextValidi]] > CurrentUpperCornerY[j] -ResetY) {
              NewColumn = false;
              break CurrentColumnLoop;
            }
          }
        }
        asdf = j;
      }
      if (NewColumn == true) {
        ResetY = ElementScalePositionYpixels[ChronologicalOrder[NextValidi]]
        // ResetY = ElementScalePositionYpixels[i+1]
        NumElementsInCurrentColumn = 0;
      }
    }
    // bring scroll window to its top if the year-span has changed
    MainTimeline.scrollBy(0, -10000);

    // FullSpanOfYears = MaxYear - MinYear;
    var TimeLineMarkerYear = Year[1];
    var DeltaYear = parseInt(SpanOfYears / 10);
    if (DeltaYear < 2) { DeltaYear = 1; }
    else if (DeltaYear < 5) { DeltaYear = 2; }
    else if (DeltaYear < 10) { DeltaYear = 5; }
    else if (DeltaYear < 25) { DeltaYear = 10; }
    else if (DeltaYear < 50) { DeltaYear = 25; }
    else if (DeltaYear < 100) { DeltaYear = 50; }
    else if (DeltaYear < 500) { DeltaYear = 100; }
    else if (DeltaYear < 1000) { DeltaYear = 500; }
    else { DeltaYear = 1000; }
    MarkerYear = (Math.floor(TimelineStartYear / DeltaYear) + 1) * DeltaYear
    MarkerYear = MarkerYear - DeltaYear;
    //remove old time line markers
    Node = document.getElementById("TimeLineMarker");
    while (Node.hasChildNodes()) {
      Node.removeChild(Node.firstChild);
    }

    // pad the last year to make sure time scale spans the last time element, which is 400px wide
    LastYear = Number(MaxYear) + 4 * DeltaYear;
    // MarkerYear = TimeLineMarkerYear;

    for (i = 0; i < 1000; i++) {
      if (MarkerYear > LastYear) { break };
      NewDiv = document.createElement('div');
      NewDiv.setAttribute('class', 'UnitOnTimeLineMarker');
      NewDiv.innerHTML = "|" + String(MarkerYear);
      ScaledPosition = (MarkerYear - TimelineStartYear) * Scale //+ 27;                               // the offset differes from 22 (see above) to make symbols align with beginning of timeline element
      NewDiv.style.transform = "translate(" + ScaledPosition.toString() + "px," + "0px)";
      document.getElementById("TimeLineMarker").appendChild(NewDiv);

      NewDiv = document.createElement('div');
      NewDiv.setAttribute('class', 'UnitOnTimeLineMarker');
      NewDiv.innerHTML = "|";
      ScaledPosition = (MarkerYear - TimelineStartYear) * Scale //+ 27;                               // the offset differes from 22 (see above) to make symbols align with beginning of timeline element
      NewDiv.style.transform = "translate(" + ScaledPosition.toString() + "px," + "8px)";
      document.getElementById("TimeLineMarker").appendChild(NewDiv);

      //sub markers
      Step = Math.max(Math.floor(DeltaYear / 5), 1)

      for (j = 0; j < DeltaYear; j = j + Step) { 
        if(j==0 && DeltaYear > 1){continue}  // submarker is not place at full year mark
        NewDiv = document.createElement('div');
        NewDiv.setAttribute('class', 'UnitOnTimeLineMarker');
        NewDiv.innerHTML = String.fromCharCode(9132);
        ScaledPosition = (MarkerYear + j - TimelineStartYear) * Scale //+ 26.5;                         // the offset differes from 22 (see above) to make symbols align with beginning of timeline element
        NewDiv.style.fontSize = "0.5vh";
        NewDiv.style.transform = "translate(" + ScaledPosition.toString() + "px," + "1.5vh)";
        document.getElementById("TimeLineMarker").appendChild(NewDiv);
      }
      MarkerYear = MarkerYear + DeltaYear;
    }
    CurrentTimelinePixelsPerUnit = Scale;
    //set width of navigation slider to show SpaneOfYears
    NavigationWindowWidthInPX = NavigationScalePixelsPerUnit * SpanOfYears;
    style.innerHTML = ".NavigateWindowSlider::-webkit-slider-thumb{ width: " + NavigationWindowWidthInPX.toString() + "px !important; height: 15px !important; }";
    // after changing year-span, reposition time-line and slider to same left-most position 
    if(CurrentLeftPositionOfTimeline){
      // console.log( 'SoY', CurrentLeftPositionOfTimeline, MinYear, CurrentTimelinePixelsPerUnit, NavigationScalePixelsPerUnit, screen.width, NavigationWindowWidthInPX, CurrentTimelinePixelsPerUnit)
      let ScrollValue = (CurrentLeftPositionOfTimeline-MinYear)*CurrentTimelinePixelsPerUnit*NavigationScalePixelsPerUnit*50/((screen.width-NavigationWindowWidthInPX/2)*CurrentTimelinePixelsPerUnit)
    NavigateWindow(ScrollValue);
    }
  }

  function ScaleOfAlphabet(SliderValue) {
    localStorage.setItem("AlphabetSlider", SliderValue.toString());
    // Span = MaxAlphaNumber - MinAlphaNumber;
    // Span = (26*(27**10) - 1*(27**10));
    for (m = 10 - 1; m >= 0; m--) {
      ListSpanOfAlpha[m] = (m+1)*27**10;
      // if (Span > ListSpanOfAlpha[m]) {
        // LastInList = m;
        // break;
      // }
    }
    //given total alphabet span, get apppropiate Span for the 10 stops on alphabet span slider
    switch (parseFloat(SliderValue*10)){
      case 0:
        SpanOfAlpha = (26*(27**10)+1*(27**9)); NumberOfLettersSpanned = 26; ValueAlphaSpan.innerHTML = "A-Z"; break;
      case 1:
        SpanOfAlpha = (13*(27**10)+1*(27**9)); NumberOfLettersSpanned = 13; ValueAlphaSpan.innerHTML = "13 Letters"; break;
      case 2:
        SpanOfAlpha = ( 7*(27**10)+1*(27**9)); NumberOfLettersSpanned = 7; ValueAlphaSpan.innerHTML = "7 Letters"; break;
      case 3:
        SpanOfAlpha = ( 5*(27**10)+1*(27**9)); NumberOfLettersSpanned = 5; ValueAlphaSpan.innerHTML = "5 Letters"; break;
      case 4:
        SpanOfAlpha = ( 3*(27**10)+1*(27**9)); NumberOfLettersSpanned = 3; ValueAlphaSpan.innerHTML = "3 Letters"; break;
      case 5:
        SpanOfAlpha = ( 2*(27**10)+1*(27**9)); NumberOfLettersSpanned = 2; ValueAlphaSpan.innerHTML = "2 Letters"; break;
      case 6:
        SpanOfAlpha = ( 1*(27**10)+1*(27**9)); NumberOfLettersSpanned = 1; ValueAlphaSpan.innerHTML = "1 Letter"; break;
      case 7:
        SpanOfAlpha = ( 0*(27**10)+13*(27**9)); NumberOfLettersSpanned = 13/26; ValueAlphaSpan.innerHTML = "1/2 Letter"; break;
      case 8:
        SpanOfAlpha = ( 0*(27**10)+7*(27**9)); NumberOfLettersSpanned = 7/26; ValueAlphaSpan.innerHTML = "1/3 Letter"; break;
      case 9:
        SpanOfAlpha = ( 0*(27**10)+5*(27**9)); NumberOfLettersSpanned = 5/26; ValueAlphaSpan.innerHTML = "1/5 Letter"; break;
      default:
        SpanOfAlpha = ( 0*(27**10)+3*(27**9)); NumberOfLettersSpanned = 3/26; ValueAlphaSpan.innerHTML = "1/7 Letter"; break;
    }
    Scale = (screen.width) / SpanOfAlpha;
    // ValueAlphaSpan.innerHTML = SpanOfAlpha.toFixed(1);
    LastPositionY = 0;
    NumCornersInCurrentColumn = 0;
    CurrentMaximumXPosition = 0;
    CurrentMinimumXPosition = 100000;
    for (i = 1; i <= NumOfTimelineElements; i++) {
      // (27**10) + (27**9) is the numerical value of Aa_________ and scales all LastNameNumbers to start at 'A'
      ElementScalePositionXpixels[AlphabeticalSortPointer[i]] = (LastNameNumber[AlphabeticalSortPointer[i]] - (27**10) - (27**9)) * Scale
      PositionX = Math.max(1,ElementScalePositionXpixels[AlphabeticalSortPointer[i]]);
      ElementScalePositionYpixels[AlphabeticalSortPointer[i]] = LastPositionY;
      LastPositionY = LastPositionY + TimelineElementHeight[AlphabeticalSortPointer[i]];
      if( DisplayElement[AlphabeticalSortPointer[i]] ){
        if(PositionX+400*ElementCellScale > CurrentMaximumXPosition){CurrentMaximumXPosition = PositionX+400*ElementCellScale};
        if(PositionX+400*ElementCellScale     < CurrentMinimumXPosition){CurrentMinimumXPosition = PositionX};
      }
    }
    ResetY = 0;
    NumElementsInCurrentColumn = 0;
    CurrentUpperCornerX = [];
    CurrentUpperCornerY = [];
    // NB!!  ElementScalePositionYpixels is already in date order, so data pointer is not used for it.
    for (i = 1; i <= NumOfTimelineElements; i++) {
      if( !DisplayElement[AlphabeticalSortPointer[i]]){
        //capture the latest y-coordinate of skipped element
        // document.getElementById("TimeLineElement"+AlphabeticalSortPointer[i].toString()).remove();
        ResetY = ElementScalePositionYpixels[AlphabeticalSortPointer[i+1]];
        continue;
      };
      //NB: using ResetY moves current column of elements to top (that is, Y = 0)
      CurrentTimelineElement = document.getElementById("TimeLineElement"+AlphabeticalSortPointer[i].toString());
      CurrentTimelineElement.style.transform = "translate(" + ElementScalePositionXpixels[AlphabeticalSortPointer[i]].toString() + "px," + String(ElementScalePositionYpixels[AlphabeticalSortPointer[i]] - ResetY) + "px)";
      CurrentTimelineElement.style.visibility = "visible";
      NumElementsInCurrentColumn = NumElementsInCurrentColumn + 1;
      CurrentUpperCornerX[NumElementsInCurrentColumn] = ElementScalePositionXpixels[AlphabeticalSortPointer[i]] + 400*ElementCellScale;
      CurrentUpperCornerY[NumElementsInCurrentColumn] = ElementScalePositionYpixels[AlphabeticalSortPointer[i]]
      // capture absolute position of elements 
      CurrentElementScalePositionXpixels[AlphabeticalSortPointer[i]] = ElementScalePositionXpixels[AlphabeticalSortPointer[i]];
      CurrentElementScalePositionYpixels[AlphabeticalSortPointer[i]] = ElementScalePositionYpixels[AlphabeticalSortPointer[i]] - ResetY;

      // see if new column can be started
      // we need the next displayed element
      for(ii = i+1; ii <= NumOfTimelineElements; ii++ ){
        if( DisplayElement[AlphabeticalSortPointer[ii]] ){
          NextValidi = ii;
          //move outerloop index to next valid i-1 (loop mechanism increments i) and leave the loop
          i=NextValidi-1;
          break;
        }else{
          //if element ii isn't displayed, increment ResetY by its y-size (height)
          ResetY += TimelineElementHeight[AlphabeticalSortPointer[ii]];
        }
      }

      NewColumn = true;
      CurrentColumnLoop2:
      for (let j = 1; j <= NumElementsInCurrentColumn; j++) {
        TimeElementLoop2:
        for (let k = NextValidi; k <= NumOfTimelineElements; k++) {
          if( !DisplayElement[AlphabeticalSortPointer[k]]){continue TimeElementLoop2};
          if (ElementScalePositionXpixels[AlphabeticalSortPointer[k]] < CurrentUpperCornerX[j]) {
            if (ElementScalePositionYpixels[AlphabeticalSortPointer[k]] + TimelineElementHeight[AlphabeticalSortPointer[k]] - ElementScalePositionYpixels[AlphabeticalSortPointer[NextValidi]] > CurrentUpperCornerY[j] - ResetY) {
            // if (ElementScalePositionYpixels[AlphabeticalSortPointer[k]] + TimelineElementHeight[AlphabeticalSortPointer[k]]  > ElementScalePositionYpixels[AlphabeticalSortPointer[NextValidi]] + CurrentUpperCornerY[j]) {
              NewColumn = false;
              break CurrentColumnLoop2;
            }
          }
        }
      }
      if (NewColumn == true) {
        ResetY = ElementScalePositionYpixels[AlphabeticalSortPointer[NextValidi]];
        NumElementsInCurrentColumn = 0;
      }
    }
    // bring scroll window to its top if the year-span has changed
    MainTimeline.scrollBy(0, -10000);
    SeekAndCenterElement(1);

    //remove old time line markers
    Node = document.getElementById("TimeLineMarker");
    while (Node.hasChildNodes()) {
      Node.removeChild(Node.firstChild);
    }

    // since we're using characters to form alphabet place marks, we need to shift to the center of the marker character
    let canvas = document.createElement("canvas");
    let context = canvas.getContext("2d");
    context.font = CurrentFont;
    let OneHalfWidthOfMarkerCharacter = context.measureText('z Ba').width*0.5;
    canvas = null;

    for (i = 0; i < 30; i++) {
      NewDiv = document.createElement('div');
      NewDiv.setAttribute('class', 'UnitOnTimeLineMarker');
      if(i<26){
        // NewDiv.innerHTML = "|" + AlphaUpper[i];
        NewDiv.innerHTML = AlphaUpper[i];
      }else{
        NewDiv.innerHTML = "|";
      }
      
      if(i==0){
        ScaledPosition =  (i*(27**10)) * Scale - OneHalfWidthOfMarkerCharacter;
        ScaledPosition = ScaledPosition + OneHalfWidthOfMarkerCharacter;}
      else{
          ScaledPosition =  (i*(27**10)+0*27**9) * Scale + OneHalfWidthOfMarkerCharacter;
      }
      NewDiv.style.transform = "translate(" + ScaledPosition.toString() + "px," + "0px)";
      document.getElementById("TimeLineMarker").appendChild(NewDiv);
      //sub markers
      if(i<26){
        if(NumberOfLettersSpanned >= 13){Step=26}
        else if(NumberOfLettersSpanned >= 7){Step=13}
        else if(NumberOfLettersSpanned >= 5){Step=7}
        else if(NumberOfLettersSpanned >= 3){Step=5}
        else if(NumberOfLettersSpanned >= 2){Step=2}
        else if(NumberOfLettersSpanned >= 1){Step=1}
        else {Step=1}
        // for (j = Math.floor(Step/2); j < 26; j = j + Step) {
        for (j = 0; j < 26; j = j + Step) {
          NewDiv = document.createElement('div');
          NewDiv.setAttribute('class', 'UnitOnTimeLineMarker');
          NewDiv.innerHTML = String.fromCharCode(9132)+AlphaLower[j];
          NewDiv.innerHTML = AlphaLower[j];
          ScaledPosition = ((i+j/26)*(27**10)) * Scale;
          NewDiv.style.transform = "translate(" + ScaledPosition.toString() + "px," + "0px)";
          document.getElementById("TimeLineMarker").appendChild(NewDiv);
        }
      }  
    }
    CurrentTimelinePixelsPerUnit = Scale;
    //set width of navigation slider to show SpaneOfLetters
    NavigationWindowWidthInPX = NavigationScalePixelsPerUnit * SpanOfAlpha;
    style.innerHTML = ".NavigateWindowSlider::-webkit-slider-thumb{ width: " + NavigationWindowWidthInPX.toString() + "px !important; height: 15px !important; }";
    // after changing letter-span, reposition time-line and slider to same left-most position     
    if(CurrentLeftPositionOfTimeline){
      // console.log( 'SoA', CurrentLeftPositionOfTimeline, MinAlphaNumber, CurrentTimelinePixelsPerUnit, NavigationScalePixelsPerUnit, screen.width, NavigationWindowWidthInPX, CurrentTimelinePixelsPerUnit)
      let ScrollValue = (CurrentLeftPositionOfTimeline-MinAlphaNumber)*CurrentTimelinePixelsPerUnit*NavigationScalePixelsPerUnit*50/((screen.width-NavigationWindowWidthInPX/2)*CurrentTimelinePixelsPerUnit)
      NavigateWindow(ScrollValue);
    }
  }

  function SetNavigationScaleInYears() {
    // set out markets on full time line at bottom of screen for each element
    //cleear navigation scale (but not the slider!)
    Node = document.getElementById("Navigate");
    while (Node.hasChildNodes() && Node.children.length > 1) {
      Node.removeChild(Node.lastChild);
    }
    //populate with year markers
    SpanOfYears = MaxYear - MinYear;
    TimeLineMarkerYear = MinYear;
    DeltaYear = parseInt(SpanOfYears / 10);
    if (DeltaYear < 2) { DeltaYear = 1; }
    else if (DeltaYear < 5) { DeltaYear = 2; }
    else if (DeltaYear < 10) { DeltaYear = 5; }
    else if (DeltaYear < 25) { DeltaYear = 10; }
    else if (DeltaYear < 50) { DeltaYear = 25; }
    else if (DeltaYear < 100) { DeltaYear = 50; }
    else if (DeltaYear < 500) { DeltaYear = 100; }
    else if (DeltaYear < 1000) { DeltaYear = 500; }
    else { DeltaYear = 1000; }
    Scale = ((screen.width) / (SpanOfYears + DeltaYear));
    TimeLineMarkerYear = (Math.floor(TimeLineMarkerYear / DeltaYear) + 1) * DeltaYear
    // TimeLineMarkerYear = (Math.floor(TimeLineMarkerYear / DeltaYear) ) * DeltaYear
    TimeLineMarkerYear = TimeLineMarkerYear - DeltaYear;
    // pad the last year to make sure time scale spans the last time element, which is 400px wide
    LastYear = Number(MaxYear); // - DeltaYear;
    TimelineStartYear = TimeLineMarkerYear;
    MarkerYear = TimeLineMarkerYear
    // since we're using characters to form naviation slider marks, we need to shift to the center of the marker character
    let canvas = document.createElement("canvas");
    let context = canvas.getContext("2d");
    context.font = CurrentFont;
    let OneHalfWidthOfMarkerCharacter = context.measureText(String.fromCodePoint(0x20D2).trim()).width*0.5;
    canvas = null;

    for (i = 0; i < 1000; i++) {
      if (MarkerYear > LastYear) { break };
      // ScaledPosition = (MarkerYear - MinYear) * Scale - 3;
      ScaledPosition = (MarkerYear - TimeLineMarkerYear) * Scale - OneHalfWidthOfMarkerCharacter;
      ScaledPosition = ScaledPosition*1.005; //tweak the placement of timeline navigation slider characters into close alignment with scrolling

      NewDiv = document.createElement('div');
      NewDiv.setAttribute('class', 'YearOnNavigateMarker');
      NewDiv.innerHTML = String.fromCodePoint(0x20D2).trim();      
      NewDiv.style.transform = "translate(" + ScaledPosition.toString() + "px," + "1vh)";
      NewDiv.style.zIndex = 100;
      document.getElementById("Navigate").appendChild(NewDiv);

      NewDiv = document.createElement('div');
      NewDiv.setAttribute('class', 'YearOnNavigateMarker');
      NewDiv.innerHTML = String.fromCodePoint(0x20D2).trim()+String(MarkerYear).trim();
      NewDiv.style.transform = "translate(" + ScaledPosition.toString() + "px," + "0px)";
      document.getElementById("Navigate").appendChild(NewDiv);
      MarkerYear = MarkerYear + DeltaYear;
    }
    NavigationScalePixelsPerUnit = Scale;
  }

  function SetNavigationScaleByAlpha() {
    // sets down marker for each entry on alphabet scale at bottom of screen
    //cleear navigation scale (but not the slider!)
    Node = document.getElementById("Navigate");
    while (Node.hasChildNodes() && Node.children.length > 1) {
      Node.removeChild(Node.lastChild);
    }
    // since we're using characters to form naviation slider marks, we need to shift to the center of the marker character    
    let canvas = document.createElement("canvas");
    let context = canvas.getContext("2d");
    context.font = CurrentFont;
    let OneHalfWidthOfMarkerCharacter = context.measureText('z Ba').width*0.5;

    //populate with Letter markers
    SpanOfAlpha = MaxAlphaNumber - 205891132094639;
    Scale = (screen.width) / (SpanOfAlpha); // + DeltaYear);
    for (i = 0; i < 26; i++) {
      NewDiv = document.createElement('div');
      NewDiv.setAttribute('class', 'YearOnNavigateMarker');
      NewDiv.innerHTML = AlphaNavigate[i];
      ScaledPosition = (i*205891132094639 ) * Scale //- OneHalfWidthOfMarkerCharacter;
      if(i==0){ScaledPosition = ScaledPosition + OneHalfWidthOfMarkerCharacter*0.5}
      NewDiv.style.transform = "translate(" + ScaledPosition.toString() + "px," + "0px)";
      document.getElementById("Navigate").appendChild(NewDiv);
    }
    NavigationScalePixelsPerUnit = Scale;
  }
  
  function PlaceMarkersByDate() {
    //clear current markers
    Node = document.getElementById("PlaceMarkers");
    while (Node.hasChildNodes()) {
      Node.removeChild(Node.lastChild);
    }
    //add null, hidden placemarker at zero
    NewDiv = document.createElement('div');
    NewDiv.setAttribute('class', 'PlaceMarkerElement');
    ScaledPosition = (Year[ChronologicalOrder[0]] - MinYear) * NavigationScalePixelsPerUnit + 15;
    NewDiv.style.transform = "translate(" + ScaledPosition.toString() + "px," + "0px)";
    document.getElementById("PlaceMarkers").appendChild(NewDiv);
    //placemarker for each element
    let MinMark = 100000;
    let MaxMark = 0;
    for (let i = 1; i <= NumOfTimelineElements; i++) {
      if( !DisplayElement[ChronologicalOrder[i]]){continue};
      NewDiv = document.createElement('div');
      NewDiv.setAttribute('class', 'PlaceMarkerElement');
      NewDiv.setAttribute('id', 'PlaceMarkerElement'+ChronologicalOrder[i].toString());
      NewDiv.innerHTML = String.fromCharCode(9670);
      ScaledPosition = (Year[ChronologicalOrder[i]] - TimelineStartYear) * NavigationScalePixelsPerUnit + 15;
      NewDiv.style.transform = "translate(" + ScaledPosition.toString() + "px," + "0px)";
      document.getElementById("PlaceMarkers").appendChild(NewDiv);
      MinMark = Math.min(ScaledPosition,MinMark)
      MaxMark = Math.max(ScaledPosition,MaxMark)
    }
    // add bounding bracket symbols marking the current span of elements on the time line
    NewDiv = document.createElement('div');
    NewDiv.setAttribute('class', 'PlaceMarkerElementBoundary');
    NewDiv.setAttribute('id', 'PlaceMarkerElementMin');
    NewDiv.innerHTML = String.fromCharCode(9665);
    NewDiv.style.transform = "translate(" + (MinMark-10).toString() + "px," + "0px)";
    document.getElementById("PlaceMarkers").appendChild(NewDiv);

    NewDiv = document.createElement('div');
    NewDiv.setAttribute('class', 'PlaceMarkerElementBoundary');
    NewDiv.setAttribute('id', 'PlaceMarkerElementMax');
    NewDiv.innerHTML = String.fromCharCode(9655);
    NewDiv.style.transform = "translate(" + (MaxMark+10).toString() + "px," + "0px)";
    document.getElementById("PlaceMarkers").appendChild(NewDiv);
  }
  
  function PlaceMarkersByAlpha() {
    //clear current markers
    Node = document.getElementById("PlaceMarkers");
    // alert(document.getElementById("PlaceMarkers").childElementCount)
    while (Node.firstChild) {
      Node.removeChild(Node.firstChild);
    }
    //add null, hidden placemarker at zero
    NewDiv = document.createElement('div');
    NewDiv.setAttribute('class', 'PlaceMarkerElement');
    NewDiv.setAttribute('id', 'PlaceMarkerElement'+"0");
    ScaledPosition = (LastNameNumber[AlphabeticalOrder[0]] - 205891132094639) * NavigationScalePixelsPerUnit + 15;
    NewDiv.style.transform = "translate(" + ScaledPosition.toString() + "px," + "0px)";
    document.getElementById("PlaceMarkers").appendChild(NewDiv);
    //placemarker for each element
    let MinMark = 100000;
    let MaxMark = 0;
    for (let i = 1; i <= NumOfTimelineElements; i++) {
      if( !DisplayElement[AlphabeticalOrder[i]]){continue};
      NewDiv = document.createElement('div');
      NewDiv.setAttribute('class', 'PlaceMarkerElement');
      NewDiv.setAttribute('id', 'PlaceMarkerElement'+AlphabeticalOrder[i].toString());
      NewDiv.innerHTML = String.fromCharCode(9670);
      // NB we use LastNameNumber[i] directly here (rather than AlphabeticalOrder) since that is how LastNameNumber is ordered.  {unlike the use of Year[ChronologicalOrder[]] above}
      ScaledPosition = (LastNameNumber[i] - 205891132094639)* NavigationScalePixelsPerUnit; //205891132094639 = number for "A" = start of alpha line
      NewDiv.style.transform = "translate(" + ScaledPosition.toString() + "px," + "0px)";
      document.getElementById("PlaceMarkers").appendChild(NewDiv);
      MinMark = Math.min(ScaledPosition,MinMark)
      MaxMark = Math.max(ScaledPosition,MaxMark)
    }
    // add bounding bracket symbols marking the current span of elements on the time line
    NewDiv = document.createElement('div');
    NewDiv.setAttribute('class', 'PlaceMarkerElementBoundary');
    NewDiv.setAttribute('id', 'PlaceMarkerElementMin');
    NewDiv.innerHTML = String.fromCharCode(9665);
    NewDiv.style.transform = "translate(" + (MinMark-10).toString() + "px," + "0px)";
    document.getElementById("PlaceMarkers").appendChild(NewDiv);

    NewDiv = document.createElement('div');
    NewDiv.setAttribute('class', 'PlaceMarkerElementBoundary');
    NewDiv.setAttribute('id', 'PlaceMarkerElementMax');
    NewDiv.innerHTML = String.fromCharCode(9655);
    NewDiv.style.transform = "translate(" + (MaxMark+10).toString() + "px," + "0px)";
    document.getElementById("PlaceMarkers").appendChild(NewDiv);
  }
  
  function ChooseHeaderForDateSorting(){
    if( BPfileHasBeenLoaded == false && BPCSVfileHasBeenLoaded == false){return}
    document.getElementById('DataHeaders').style.display = 'block';
    WidthOfCellText = document.getElementById('DataHeaders').offsetWidth;
    WidthOfCellText = (Number(WidthOfCellText)*0.85).toString()+'px';
    dragElement(document.getElementById("DataHeadersDrag"));
    let element = document.getElementById('DataHeadersContainer');
    element.style.display = 'flex';
    HeaderTable = document.createElement('TABLE');
    for (let m = 1; m <= NumUserDataFields; m++) {
      if(data[0][BPHeadersOfUserData[m]]==''){continue}
      NewRow = HeaderTable.insertRow();

      NewCell = NewRow.insertCell(-1);
      CheckBox = document.createElement('input');
      CheckBox.type = 'radio';      
      CheckBox.name = 'HeaderName';
      CheckBox.id = 'HeaderName';
      NewCell.appendChild(CheckBox);

      HeaderName = NewRow.insertCell(-1);
      HeaderName.style.display = 'block';
      HeaderName.style.width = WidthOfCellText;
      HeaderName.style.whiteSpace = 'nowrap';
      HeaderName.style.overflow = 'hidden';
      HeaderName.style.textOverflow = 'ellipsis';
      HeaderName.title = data[0][BPHeadersOfUserData[m]];
      HeaderName.innerHTML = " " + data[0][BPHeadersOfUserData[m]];
    } 
    element.appendChild(HeaderTable);    
    // check box for the currrnet data header for date
    InputBoxes = HeaderTable.getElementsByTagName("INPUT");
    for (m = 0; m < InputBoxes.length; m++) {
      if(HeaderForDateSorting == meta.fields[m]){
        InputBoxes[m].checked == true;
        InputBoxes[m].click();
        break
      }
    }
    DateHeaderOpen = true;
  }
  function GetDateDataForSorting(DataHeader){
    DisplayElement.fill(true);
    MinYear =  10000;
    MaxYear = -10000;
    Year.length = 0;
    YearParsingLoop: 
    for (i = 0; i < data.length; i++) {
      if (i == 0) {
        Year[i] = -10000;
        LastName[i] = "          ";
        LastNameNumber[i] = 0;
        ShortTitle[0] = " ";
        DisplayElement[i] = false;
      } else {
        let DateText = data[i][DataHeader]
        if(DateText.toLowerCase().includes('bc')){
          BC = true;
        }else{
          BC = false;
        }
        // parse this date string. extract year digits, if present.
        if (DateText.includes(".")) {
          DateText = DateText.split(".", 1)[0];
        }
        if (DateText.includes("-")) {
          DateText = DateText.split("-", 1)[0];
        }
        if(DateText.match(/\d+/)==null){Year[i] = 0.0; DisplayElement[i] = false; continue YearParsingLoop;}

        //['?', 'nd', 'n.d.', 'c', 'ca', 'circa', 'fl', '?', '(', ')', '{', '}', '[', ']']
        for (let m=0; m<DateDetectionStrings.length; m++){
          if(DateText==null){Year[i] = 0.0; DisplayElement[i] = false; continue YearParsingLoop;}
          // console.log(DateText);
          if (DateText.toLowerCase().includes(DateDetectionStrings[m])) {DateText = DateText.match(/\-?\d+/)[0]};
          if(DateText.match(/\d+/)==null){Year[i] = 0.0; DisplayElement[i] = false; continue YearParsingLoop;}
        }
        try{
          Year[i] = parseFloat(DateText)
        }
        catch(err){
          DateText = DateText.match(/\-?\d+/)[0];
          Year[i] = parseFloat(DateText)
        }
        if(BC){Year[i] = -Year[i]};
        if( Number.isNaN(Year[i]) ){Year[i] = 0.0; DisplayElement[i] = false};

        if(DisplayElement[i]==true){
          MinYear = Math.min(MinYear, Year[i]);
          MaxYear = Math.max(MaxYear, Year[i]);
        }
      }
    }    
    //build a pointer to elements sorted by year
    heapSortByDate();
    DateSortPointer[0] = 0;
    document.getElementById('DateSortHeader').innerText = data[0][DataHeader]; //put user data header name into date sort label
    document.getElementById('DateSortHeader').title = data[0][DataHeader]; //put user data header name hover title
  }
  
  function ChooseHeaderForAlphaSorting(){
    if( BPfileHasBeenLoaded == false && BPCSVfileHasBeenLoaded == false){return}
    document.getElementById('DataHeaders').style.display = 'block';
    WidthOfCellText = document.getElementById('DataHeaders').offsetWidth;
    WidthOfCellText = (Number(WidthOfCellText)*0.85).toString()+'px';    
    dragElement(document.getElementById("DataHeadersDrag"));
    let element = document.getElementById('DataHeadersContainer');
    element.style.display = 'flex';
    HeaderTable = document.createElement('TABLE');
    for (let m = 1; m <= NumUserDataFields; m++) {
      NewRow = HeaderTable.insertRow();

      NewCell = NewRow.insertCell(-1);
      CheckBox = document.createElement('input');
      CheckBox.type = 'radio';      
      CheckBox.name = 'HeaderName';
      CheckBox.id = 'HeaderName';
      NewCell.appendChild(CheckBox);

      HeaderName = NewRow.insertCell(-1);
      HeaderName.style.display = 'block';
      HeaderName.style.width = WidthOfCellText;
      HeaderName.style.whiteSpace = 'nowrap';
      HeaderName.style.overflow = 'hidden';
      HeaderName.style.textOverflow = 'ellipsis';
      HeaderName.title = data[0][BPHeadersOfUserData[m]];
      HeaderName.innerHTML = " " + data[0][BPHeadersOfUserData[m]];
    } 
    element.appendChild(HeaderTable);    
    // check box for the currrnet data header for date
    InputBoxes = HeaderTable.getElementsByTagName("INPUT");
    for (m = 0; m < InputBoxes.length; m++) {
      if(HeaderForAlphaSorting == meta.fields[m]){
        InputBoxes[m].checked == true;
        InputBoxes[m].click();
        break
      }
    }
    AlphaHeaderOpen = true;
  }
  function GetAlphaDataForSorting(DataHeader){
    //for alphabetical sorting, build alphabetical data with diacritical marks removed and all uppercase.
    DisplayElement.fill(true);
    for (i = 0; i < data.length; i++) {
      NormalizedField = data[i][DataHeader].normalize('NFD').replace(/[\u0300-\u036f]/g, "");
      NormalizedField = NormalizedField.replace(/\(/g, "");
      NormalizedField = NormalizedField.replace(/\)/g, "");
      NormalizedField = NormalizedField.replace(/´/g, "");
      NormalizedField = NormalizedField.replace(/\?/g, "");
      NormalizedField = NormalizedField.replace(/" "/g, "@"); // @ has uncode == 64, one less than 'A', so blank becomes 1 less than 'A'
      NormalizedField = NormalizedField.toUpperCase()
      NumChars = NormalizedField.length;

      if(NumChars > 0){
        if (NumChars > 11) { NormalizedField = NormalizedField.slice(0, 11) } else if (NumChars < 11) { NormalizedField = NormalizedField.padEnd(11) }
        NormalizedField = NormalizedField.split(',')[0]     //get chars just up to any comma
        Sum = 0;
        for (m = 0; m <= NormalizedField.length-1; m++) {
          // AlphaDigitValue = 27^(n-1) where n = char position in the last name beginning from the left
          // blank -> 0, A -> 1, B ->2, . . . Z -> 26
          Sum = Sum + (NormalizedField.charCodeAt(m) - 64) * AlphaDigitValue[10 - m];
        }
        LastNameNumber[i] = Sum;
      }else{
        DisplayElement[i] = false;
        LastNameNumber[i] = 0;
      }
      LastName[i] = NormalizedField;
      MinAlphaNumber = Math.min(MinAlphaNumber, LastNameNumber[i]);
      MaxAlphaNumber = Math.max(MaxAlphaNumber, LastNameNumber[i]);
    }
    heapSortAlphabetize()
    document.getElementById('AlphaSortHeader').innerText = data[0][DataHeader]; //put user data header name into alpha sort label
    document.getElementById('AlphaSortHeader').title = data[0][DataHeader]; //put user data header name into hover title
  }

  function CloseChooseHeaderForSorting(){
    // save latest radio-button pick for date sort header
    InputBoxes = HeaderTable.getElementsByTagName("INPUT");
    for (m = 0; m < InputBoxes.length; m++) {
      if( InputBoxes[m].checked == true ){
        HeaderForSorting = meta.fields[m];
        break
      }
    }
    // empty header table
    let element = document.getElementById('DataHeadersContainer')
    while (element.firstChild) {
      element.removeChild(element.lastChild);
    }
    document.getElementById('DataHeaders').style.display = 'none';
    if(DateHeaderOpen){
      HeaderForDateSorting = HeaderForSorting;
      GetDateDataForSorting(HeaderForDateSorting);   
      DateHeaderOpen = false;
    }else if(AlphaHeaderOpen){
      HeaderForAlphaSorting = HeaderForSorting;
      GetAlphaDataForSorting(HeaderForAlphaSorting)
      AlphaHeaderOpen = false;
    }
    
  }

function pxTOvw(value) {
  var w = window,
    d = document,
    e = d.documentElement,
    g = d.getElementsByTagName('body')[0],
    x = w.innerWidth || e.clientWidth || g.clientWidth,
    y = w.innerHeight|| e.clientHeight|| g.clientHeight;
  var result = (100*value)/x;
  return result;
}

function pxTOvh(value) {
  var w = window,
    d = document,
    e = d.documentElement,
    g = d.getElementsByTagName('body')[0],
    x = w.innerWidth || e.clientWidth || g.clientWidth,
    y = w.innerHeight|| e.clientHeight|| g.clientHeight;
  var result = (100*value)/y;
  return result;
}

function vwTOpx(value) {
  var w = window,
    d = document,
    e = d.documentElement,
    g = d.getElementsByTagName('body')[0],
    x = w.innerWidth || e.clientWidth || g.clientWidth,
    y = w.innerHeight|| e.clientHeight|| g.clientHeight;
  var result = (x*value)/100;
  return result;
}

function vhTOpx(value) {
  var w = window,
    d = document,
    e = d.documentElement,
    g = d.getElementsByTagName('body')[0],
    x = w.innerWidth || e.clientWidth || g.clientWidth,
    y = w.innerHeight|| e.clientHeight|| g.clientHeight;
  var result = (y*value)/100;
  return result;
}
//#endregion

// Bookmarking management -->
//#region
  function BookmarkThisElement(ElementNumber, Source) {
   // don't duplicate bookmarks
    for (m = 1; m <= NumberOfBookMarks; m++) {
      if (BookMarks[m] == ElementNumber) {
        RemoveBookmark(ElementNumber, Source);
        CloseBookmarks();
        // return color of timeline bookmark button to alert color
        document.getElementById("BookmarkElementButton"+ ElementNumber.toString() + ")").innerHTML = "<img class='img' src='BP Appearance/bookmarkFull.png' alt='Some png'>";
        return;
      }
    }
    // change color of timeline element bookmark button image
    document.getElementById("BookmarkElementButton"+ ElementNumber.toString() + ")").innerHTML = "<img class='img' src='BP Appearance/BookmarkFullAlertComp.png' alt='Some png'>";

    // build array of bookmark timeline element numbers
    NumberOfBookMarks = NumberOfBookMarks + 1
    BookMarks[NumberOfBookMarks] = ElementNumber;
    // if(ChronologicalSort){OrderElementNumber = ChronologicalOrder[ElementNumber]}
    // else if(AlphabeticalSort){OrderElementNumber = AlphabeticalOrder[ElementNumber]};

    if(!isNaN(Source)){  
      //if bookmarking comes from search results, change search marker image to white
      ImgHeight = vhTOpx(2);
      ImgWidth = pxTOvw(ImgHeight);
      ImgWidth = vwTOpx(ImgWidth);
      document.getElementById("PlaceMarkerElement"+ElementNumber.toString()).innerHTML = "<img src=\"BP Appearance/magnifying-glass WHITE.png\" width="+ImgWidth.toString()+"px height="+ImgHeight.toString()+"px >";
      document.getElementById("PlaceMarkerElement"+ElementNumber.toString()).style.opacity = "1";
      document.getElementById("PlaceMarkerElement"+ElementNumber.toString()).style.zIndex = "10";
      //for bookmarking in the list of search hits, the bookmark button is flagged as already clicked by graying it      
      let ButtonNumber = Number(Source);     
      document.getElementById("BookmarkSearchHitButton"+ButtonNumber.toString()).style.opacity = "0.5";
      
    }else{
      //if not, put a bookmarker on the timeline at screen bottom; i.e. change element's timeline marker to opaque white.
      document.getElementById("PlaceMarkerElement"+ElementNumber.toString()).style.opacity = "1";
      document.getElementById("PlaceMarkerElement"+ElementNumber.toString()).style.color = "var(--Gray0)";
      document.getElementById("PlaceMarkerElement"+ElementNumber.toString()).style.zIndex = "10";
      document.getElementById("PlaceMarkerElement"+ElementNumber.toString()).innerHTML = String.fromCharCode(9670);
    }

    // add bookmark element to bookmark container 
    NewDiv = document.createElement('div');
    NewDiv.setAttribute('class', 'CurrentBookmark');
    NewDiv.setAttribute('id', 'CurrentBookmark'+NumberOfBookMarks.toString());
    NewDiv.innerHTML = document.getElementById("BriefBookTitle"+ElementNumber.toString()).innerHTML.toUpperCase() + " - " +
                       document.getElementById("PublicationYear"+ElementNumber.toString()).innerHTML + "<br>" +
                       document.getElementById("Author"+ElementNumber.toString()).innerHTML;
    document.getElementById("BookmarkContainer").appendChild(NewDiv);
    AddItHere = document.getElementById("CurrentBookmark"+NumberOfBookMarks.toString());

    //add bookmark delete button to bookmark entry
    NewButton = document.createElement('button');
    NewButton.innerHTML = "&#xf139";
    NewButton.setAttribute("class", "RemoveBookmarkButton");
    NewButton.setAttribute("onclick", "RemoveBookmark(" + ElementNumber.toString() + ", ' ')");
    NewButton.innerHTML = "<img class='ListButtonImages' src='BP Appearance/Bin.png' alt='Some png'>";
    AddItHere.appendChild(NewButton);

    //change styling of a bookmakred element: border, background.
    //get textcolor, invert it, and use that as the background color for a bookmakred element
    TextColor = localStorage.getItem("TimeElementTextColor");
    let rgb = TextColor.match(/\d+/g);
    r = 255-Number(rgb[0]);
    g = 255-Number(rgb[1]);
    b = 255-Number(rgb[2]);
    TextColor = "rgb("+ +r + "," + +g + "," + +b + " )"
    document.documentElement.style.setProperty('--TimeLineElementBookmarkedBackground', TextColor);
    document.getElementById("TimeLineElement"+ElementNumber.toString()).setAttribute("class", "TimeLineElementBookmarked");

    // make bookmark seek buttons active
    document.getElementById("BookmarkSeekBackward").disabled = false;
    document.getElementById("BookmarkSeekForward").disabled = false;
    document.getElementsByClassName("BookmarkSeekBackwardButtonImage")[0].style.opacity = "1.0";
    document.getElementsByClassName("BookmarkSeekForwardButtonImage")[0].style.opacity = "1.0";
    document.getElementById("BookmarkSeekBackward").style.cursor = "pointer";
    document.getElementById("BookmarkSeekForward").style.cursor = "pointer";
  }

  function ShowBookmarks() {
    if (BookmarkDisplayOn == false) {
      document.getElementById('BookmarkManagement').style.visibility = "visible";
      BookmarkDisplayOn = true;
      dragElement(document.getElementById("BookmarkManagementDrag"));
    }
    else {
      document.getElementById('BookmarkManagement').style.visibility = "hidden";
      BookmarkDisplayOn = false;
    }
  }

  function CloseBookmarks() {
    document.getElementById('BookmarkManagement').style.visibility = "hidden";
    BookmarkDisplayOn = false;
  }

  function RemoveBookmark(ElementNumber, Source) {
    // remove entry from Bookmark array.
    for (m = 1; m <= NumberOfBookMarks; m++) {
      if (BookMarks[m] == ElementNumber) {
        BookMarks.splice(m, 1);
        DeleteHere = document.getElementById("BookmarkContainer");
        DeleteHere.removeChild(DeleteHere.childNodes[m-1]);
        //turn off bookmark highlight
        if(ChronologicalSort){OrderElementNumber = ChronologicalOrder[ElementNumber]}
        else if(AlphabeticalSort){OrderElementNumber = AlphabeticalOrder[ElementNumber]};

        document.getElementById("PlaceMarkerElement"+ElementNumber.toString()).style.opacity = "0.15";
        document.getElementById("PlaceMarkerElement"+ElementNumber.toString()).style.color = "rgb(9, 182, 254)";
        document.getElementById("PlaceMarkerElement"+ElementNumber.toString()).style.zIndex = "0";
        //revert colors and border to bookmakred elements
        document.getElementById("TimeLineElement"+ElementNumber.toString()).setAttribute("class", "TimeLineElement");
        document.getElementById("BookmarkElementButton"+ ElementNumber.toString() + ")").innerHTML = "<img class='img' src='BP Appearance/BookmarkFull.png' alt='Some png'>"
        break;
      }
    }
    NumberOfBookMarks = NumberOfBookMarks - 1;
    if (NumberOfBookMarks < 1) {
      // make bookmark seek buttons inactive
      document.getElementById("BookmarkSeekBackward").disabled = true;
      document.getElementById("BookmarkSeekForward").disabled = true;
      document.getElementsByClassName("BookmarkSeekBackwardButtonImage")[0].style.opacity = "0.25";
      document.getElementsByClassName("BookmarkSeekForwardButtonImage")[0].style.opacity = "0.25";
      document.getElementById("BookmarkSeekBackward").style.cursor = "default";
      document.getElementById("BookmarkSeekForward").style.cursor = "default";
      LastBookmarkElementHighlighted = 0;
    }
    if(!isNaN(Source)){
      let ButtonNumber = Number(Source);
      if(ButtonNumber > 0){
        document.getElementById("BookmarkSearchHitButton"+ButtonNumber.toString()).style.opacity = "1";
        document.getElementById("PlaceMarkerElement"+ElementNumber.toString()).innerHTML = "<img src='BP Appearance/magnifying-glass RED.png' width="+ImgWidth.toString()+"px height="+ImgHeight.toString()+"px >";      
        document.getElementById("PlaceMarkerElement"+ElementNumber.toString()).style.opacity = "1";
        document.getElementById("PlaceMarkerElement"+ElementNumber.toString()).style.zIndex = "10";
      }
    }
    //bookmark page is already open, so we need to reset the toggle state to bookmark page stays open
    BookmarkDisplayOn = false;
    ShowBookmarks();
  }

  function SeekNextBookmark() {
    CurrentBookmarkHighlighted = CurrentBookmarkHighlighted + 1
    if (CurrentBookmarkHighlighted > NumberOfBookMarks) { CurrentBookmarkHighlighted = 1; }
    SeekAndCenterElement(BookMarks[CurrentBookmarkHighlighted]);
  }

  function SeekPreviousBookmark() {
    CurrentBookmarkHighlighted = CurrentBookmarkHighlighted - 1;
    if (CurrentBookmarkHighlighted <= 0) { CurrentBookmarkHighlighted = NumberOfBookMarks; }
    SeekAndCenterElement(BookMarks[CurrentBookmarkHighlighted]);
  }

  function ClearAllBookmarks() {
    DeleteHere = document.getElementById("BookmarkContainer");
    while (DeleteHere.firstChild) {
      DeleteHere.firstChild.remove();
    }
    for (m = 1; m <= NumberOfBookMarks; m++) {
      // BookMarks.splice(m,1);
      ElementNumber = BookMarks[m];
      BookMarks[m] = 0;
      if(ChronologicalSort){OrderElementNumber = ChronologicalOrder[ElementNumber]}
        else if(AlphabeticalSort){OrderElementNumber = AlphabeticalOrder[ElementNumber]};
      //turn off bookmark highlight
      document.getElementById("PlaceMarkerElement"+ElementNumber.toString()).style.opacity = "0.15";
      document.getElementById("PlaceMarkerElement"+ElementNumber.toString()).style.color = "var(--AlertColor)";
      document.getElementById("PlaceMarkerElement"+ElementNumber.toString()).style.zIndex = "0";
      //revert colors and border to bookmakred elements
      document.getElementById("TimeLineElement"+ElementNumber.toString()).setAttribute("class", "TimeLineElement");
      document.getElementById("BookmarkElementButton"+ ElementNumber.toString() + ")").innerHTML = "<img class='img' src='BP Appearance/BookmarkFull.png' alt='Some png'>"      
    }
    NumberOfBookMarks = 0;
    LastBookmarkElementHighlighted = 0;
    // make bookmark seek buttons inactive
    document.getElementById("BookmarkSeekBackward").disabled = true;
    document.getElementById("BookmarkSeekForward").disabled = true;
    document.getElementsByClassName("BookmarkSeekBackwardButtonImage")[0].style.opacity = "0.25";
    document.getElementsByClassName("BookmarkSeekForwardButtonImage")[0].style.opacity = "0.25";
    document.getElementById("BookmarkSeekBackward").style.cursor = "default";
    document.getElementById("BookmarkSeekForward").style.cursor = "default";
    //bookmark page is already open, so we need to reset the toggle state to bookmark page stays open
    BookmarkDisplayOn = false;
    ShowBookmarks();
  }

  function SaveBookmarkList(){
    // build text string for output
    //Analysls = document.getElementById('BookmarkList')
    BookmarkList = ""
    BookmarkList = '<strong>List of Bookmarked Entries</strong><br><br>';
    for (m = 1; m <= NumberOfBookMarks; m++) {
      ElementNumber = BookMarks[m];
      // if(ChronologicalSort){OrderElementNumber = ChronologicalOrder[ElementNumber]}
        // else if(AlphabeticalSort){OrderElementNumber = AlphabeticalOrder[ElementNumber]};
        BookmarkList += document.getElementById("BriefBookTitle"+ElementNumber.toString()).innerHTML.toUpperCase() + " - " +
                        document.getElementById("PublicationYear"+ElementNumber.toString()).innerHTML + "<br>" +
                        document.getElementById("Author"+ElementNumber.toString()).innerHTML+"<br><br>";      
    }
    BookmarkList = BookmarkList.replace(/<br>/gi, String.fromCharCode(13));
    BookmarkList = BookmarkList.replace(/<strong>/gi, "");
    BookmarkList = BookmarkList.replace(/<\/strong>/gi, "");
    BookmarkList = BookmarkList.replace(/<em>/gi, "");
    BookmarkList = BookmarkList.replace(/<\/em>/gi, "");    
    download(BookmarkList, "ListOfBookmarkedEntries.txt","text/plain");
  }
//#endregion

// place seek and center -->
//#region
  function SeekAndCenterElement(ElementNumber) {

    if(ChronologicalSort){
      // SeekPositionX = CurrentElementScalePositionXpixels[ChronologicalOrder[ElementNumber]] - screen.width * 0.5 + 200;
      // SeekPositionY = CurrentElementScalePositionYpixels[ChronologicalOrder[ElementNumber]] - screen.height * 0.25;
      SeekPositionX = CurrentElementScalePositionXpixels[ElementNumber] - screen.width * 0.5 + 200;
      SeekPositionY = CurrentElementScalePositionYpixels[ElementNumber] - screen.height * 0.25;      
      document.getElementById("MainTimeline").scrollTo(SeekPositionX, SeekPositionY);
      document.getElementById("TimeLineMarker").scrollTo(SeekPositionX, 0);
    }
    else if(AlphabeticalSort){
      // SeekPositionX = CurrentElementScalePositionXpixels[AlphabeticalOrder[ElementNumber]] - screen.width * 0.5 + 200;
      // SeekPositionY = CurrentElementScalePositionYpixels[AlphabeticalOrder[ElementNumber]] - screen.height * 0.25;
      SeekPositionX = CurrentElementScalePositionXpixels[ElementNumber] - screen.width * 0.5 + 200;
      SeekPositionY = CurrentElementScalePositionYpixels[ElementNumber] - screen.height * 0.25;      
      document.getElementById("MainTimeline").scrollTo(SeekPositionX, SeekPositionY);
      document.getElementById("TimeLineMarker").scrollTo(SeekPositionX, 0);
    }

    // highlight in read the current seek element in the full timeline
    let OrderNumber = 0;
    if(ChronologicalSort){OrderNumber= ChronologicalOrder[ElementNumber]}
    else if(AlphabeticalSort){OrderNumber= AlphabeticalOrder[ElementNumber]}

    document.getElementById("PlaceMarkerElement"+ElementNumber.toString()).style.color = "rgb(255, 0, 0)";
    document.getElementById("PlaceMarkerElement"+ElementNumber.toString()).style.zIndex = "11";
    if (LastBookmarkElementHighlighted != 0) {
      document.getElementById("PlaceMarkerElement"+LastBookmarkElementHighlighted.toString()).style.color = "rgb(255, 255, 255)";
      document.getElementById("PlaceMarkerElement"+LastBookmarkElementHighlighted.toString()).style.zIndex = "10";
    }
    LastBookmarkElementHighlighted = ElementNumber;
  }
//#endregion

// General help system -->
//#region

   function LoadGeneralDocumentation(){
    HelpElement = document.getElementById('Help');
    if(HelpElement.style.visibility == "visible" ){
      HelpElement.style.visibility = "hidden";
      document.getElementById('HelpToC').style.visibility = "hidden";
    }else{
      HelpElement.style.visibility = "visible";
      dragElement(document.getElementById("HelpDrag"));
      document.getElementById('HelpToC').style.visibility = "visible";
      document.getElementById("HelpIframe").src = "Instructions/BookPageant User Guide and Documentation.pdf#page=2&view=fit"
    }
  }

  async function OpenHelp(SectionNumber){
    document.getElementById('HelpToC').style.visibility = "hidden";    
    HelpElement = document.getElementById('Help');
    HelpElement.style.visibility = "visible";
    dragElement(document.getElementById("HelpDrag"));
    // get appropriate page number in documentation
    switch(SectionNumber){
      case '1.0':
        DocPage = 6;        
       break; 
      case '2.0':
        DocPage = 10;        
        break; 
      case '3.0':
        DocPage = 19;        
        break; 
      case '4.0':
        DocPage = 23;        
        break;
      case '5.0':
        DocPage = 25;
        break;
      case '6.0':
        DocPage = 30;        
        break;
      case '7.0':
        DocPage = 32;        
        break; 
      case '8.0':
        DocPage = 35;        
        break; 
      case '9.0':
        DocPage = 39;        
        break; 
      case '10.0':
        DocPage = 40;
        break;
        case '10.2':
          DocPage = 48;
          break;                
        case '10.4':
          DocPage = 52;
          break;        
      case '11.0':
        DocPage = 57;        
        break;
      case '12.0':
        DocPage = 60;        
        break;
      case '13.0':
        DocPage = 63;        
        break;
        case '13.3':
          DocPage = 64;
          break;        
      case '14.0':
        DocPage = 75;        
        break;                  
    } 
    document.getElementById("HelpIframe").src = 'about:blank';
    await new Promise(r => setTimeout(r, 125));                     //wait for iframe to refresch
    document.getElementById("HelpIframe").src = "Instructions/BookPageant User Guide and Documentation.pdf#page="+DocPage.toString()+"&view=fit";
    await new Promise(r => setTimeout(r, 125));                      //wait for iframe to refresch
  }
  
  function CloseHelp(){
    document.getElementById('Help').style.visibility = "hidden";
    document.getElementById('HelpToC').style.visibility = "hidden";    
    document.getElementById("HelpIframe").src = "";
  }
    
//#endregion

// search management -->
//#region
  function ShowSearchManagement() {
    if (SearchDisplayOn == false) {
      document.getElementById('SearchManagement').style.visibility = "visible";
      SearchDisplayOn = true;
      dragElement(document.getElementById("SearchManagementDrag"));
      document.getElementById("SearchFieldsContainer").focus();
    }
    else {
      document.getElementById('SearchManagement').style.visibility = "hidden";
      SearchDisplayOn = false;
    }
  }
  
  function CloseSearchManagement() {
    document.getElementById('SearchManagement').style.visibility = "hidden";
    CloseAvailableSearchLists('Search');
    SearchDisplayOn = false;
  }
  
  function StartSearch() {
    SearchTable = document.getElementById("SearchTable");
    InputBoxes = SearchTable.getElementsByTagName("INPUT");
    NumSearchItems = 0
    for (m = 0; m < InputBoxes.length; m++) {
      if (InputBoxes[m].getAttribute('type') == 'checkbox' && InputBoxes[m].checked == true && InputBoxes[m + 1].value.length != 0) {
        NumSearchItems = NumSearchItems + 1;
        // in each row of search filed table:  check box, field name, input box. so the input boxes are [0,1], [2,3], [4,5], ...
        // the checkbox and search field alternate in the container, so the checkboxes are the even-numbered elements.        
        FieldToSearch[NumSearchItems] = BPHeadersOfUserData[(m + 2) / 2];
        FieldToSearchIsAlphabetic[NumSearchItems] = BPHeadersOfUserDataIsAlphabetic[(m + 2) / 2];
        FieldToSearchIsNumeric[NumSearchItems] = BPHeadersOfUserDataIsNumeric[(m + 2) / 2];
        SearchValue[NumSearchItems] = InputBoxes[m + 1].value;
        SearchValueInputBoxNumber[NumSearchItems] = m + 1;
        FieldSearchNumber[NumSearchItems] = m; //(m + 2) / 2
      }
    }
    SearchBookData(NumSearchItems);
  }
  
  function areParenthesesBalanced(TestString){
      //increment or decrement uptoPrevChar according to open or closing barckets. Balanced if ends at zero.
      return !TestString.split('').reduce((uptoPrevChar, thisChar) => {
      if(thisChar === '(' || thisChar === '{' || thisChar === '[' ) {
          return ++uptoPrevChar;
      } else if (thisChar === ')' || thisChar === '}' || thisChar === ']') {
          return --uptoPrevChar;
      }
      return uptoPrevChar
    }, 0);
  }
  
  function SearchBookData(NumSearchItems) {
    if(NumSearchItems==0){return}

    //check search phrase(s) for balanced bracketing.
    SearchTable = document.getElementById("SearchTable");
    InputBoxes = SearchTable.getElementsByTagName("INPUT");
    for (k = 1; k <= NumSearchItems; k++) {
      TestString = SearchValue[k];
      if(!areParenthesesBalanced(TestString)){
        InputBoxes[SearchValueInputBoxNumber[k]].style.color = "rgb(255,0,0)";
        NumberOfElementSearchHits = 0;
        return;        
      }else{
        InputBoxes[SearchValueInputBoxNumber[k]].style.color = "rgb(0,0,0)";
      }
    }
    // clear last search hits, if any, and make an empty node at zero
    ClearAllSearchHits();
    Node = document.getElementById("SearchResultsContainer");
    while (Node.hasChildNodes()) {
      Node.removeChild(Node.lastChild);
    }
    NewDiv = document.createElement('div');
    NewDiv.setAttribute('class', 'CurrentSearchHit');
    NewDiv.style.visibility = "hidden";
    NewDiv.style.height = "0px";
    document.getElementById("SearchResultsContainer").appendChild(NewDiv);
    NumberOfElementSearchHits = 0
    NumberOfElementSearchHitsTop = 0
    NumberOfElementSearchHitsBottom = 0
    SearchHits[NumberOfElementSearchHits] = 0;
    SearchHit = false;
    NumSortedSearchItems = 0
    let SortedSearchItems = [];
    let SearchItemHits = [];
    //build pointer that puts all "OR" search fields at the of the list of search elements and all "AND" at the bottom.
    //This allows early out of the SearchItemLoop if either: a search hit on an "OR" field or a fail on any "AND" field
    for (k = 1; k <= NumSearchItems; k++) {
      SearchValue[k] = SearchValue[k].trim();
      if(SearchValue[k].charAt(0) == "|" ){
        NumberOfElementSearchHitsTop = NumberOfElementSearchHitsTop + 1;
        SortedSearchItems[NumberOfElementSearchHitsTop] = k;
      }else{
        NumberOfElementSearchHitsBottom = NumberOfElementSearchHitsBottom + 1;
        SortedSearchItems[NumSearchItems+1-NumberOfElementSearchHitsBottom] = k
      }
    }
    // alert(SearchValue+"\n"+SearchValue.length+"\n"+SortedSearchItems+"\n"+SortedSearchItems.length)

    ElementLoop:
    for (let SearchElementNumber = 1; SearchElementNumber <= NumOfTimelineElements; SearchElementNumber++) {
      SearchHit = false;
      NumberOfHits = 0;

      SearchItemLoop:
      for (kk = 1; kk <= NumSearchItems; kk++) {
        k = SortedSearchItems[kk]
        // parse appropriate search element for name and numbers
        let ElementsFromSearchFieldFound = [];
        UserSearchField = SearchValue[k]; 
        // see if user has specified this as an "AND" search by testing for the the presence of "&" as the 1st character.
        // If NOT, then failure to find this search value in the current element is enough to disqualify current element. Search can be halted
        // If so, then "OR" means if search value is in current element, that's enough to qualify current element. Search can be halted.
        ANDfield = true
        if(UserSearchField.charAt(0) == "&"){
          ANDfield = true;
        }else if(UserSearchField.charAt(0) == "|"){
          //remove the "|" from the front of the search field and set ANDfield to false; that is, this field is not required
          UserSearchField = UserSearchField.slice(1);
          ANDfield = false;
        }
        //if the user search field is ONLY "?", the simply detect the presence of anything but null/blank/empty in the data
        if(UserSearchField==="?"){
          for( m = 0; m < UserSearchField.length; m++){
            if(data[SearchElementNumber][FieldToSearch[k]] != ""){
              SearchHit = true;
            }
          }
        }else if(UserSearchField==="!?"){
          for( m = 0; m < UserSearchField.length; m++){
            if(data[SearchElementNumber][FieldToSearch[k]] == ""){
              SearchHit = true;
            }
          }
        }else{
          if( FieldToSearchIsAlphabetic[k] ){
            //if current search item is alphabetic, parse it to isolate user-supplied names/text/strings in preparation for evaluation 
            ElementsFromSearchField = UserSearchField; 
            // ElementsFromSearchField = ElementsFromSearchField.split(/[.:;?!~,&|()<>{}\[\]\r\n/\\]+/);
            ElementsFromSearchField = ElementsFromSearchField.split(/[:;?!~,&|()<>{}\[\]\r\n/\\]+/);  // no "." in regex ??
            //remove empty strings from array resulting from the regexp split
            ElementsFromSearchField = ElementsFromSearchField.filter(Boolean);
            // alert(ElementsFromSearchField)
            for( m = 0; m < ElementsFromSearchField.length; m++){
              NormalizedFieldToSearch = data[SearchElementNumber][FieldToSearch[k]].normalize('NFD').replace(/[\u0300-\u036f]/g, "");
              // alert(m+" "+ElementsFromSearchField[m]+" "+FieldToSearch[k]+" "+NormalizedFieldToSearch+" "+UserSearchField);
              //test each ElementsFromSearchField, flagging found/not found as true/false, reconstructing parsed search field with true/false substituted for user-specified words
              if (RegExp(ElementsFromSearchField[m].trim()).test(NormalizedFieldToSearch)) {
                UserSearchField = UserSearchField.replace(ElementsFromSearchField[m], "true");
                // alert(SearchElementNumber+" "+m+" true "+ElementsFromSearchField[m]+" "+NormalizedFieldToSearch+" "+UserSearchField)
              } else {
                UserSearchField = UserSearchField.replace(ElementsFromSearchField[m], "false");
              }
            }
            SearchHit = eval(UserSearchField);
            if(!SearchHit && ANDfield){
              //current element fails required "AND" search field, move to the next element
              continue ElementLoop;
            }else if(SearchHit && !ANDfield){
              //current element meats an "OR" search field, leave search loop, and move to code that logs current element to the search hit list
              NumberOfHits = NumberOfHits + 1;
              SearchItemHits[NumberOfHits] = SearchValueInputBoxNumber[k];
              break SearchItemLoop;
            }else if(SearchHit){
              NumberOfHits = NumberOfHits + 1;
              SearchItemHits[NumberOfHits] = SearchValueInputBoxNumber[k];
              continue SearchItemLoop;
            }
          }else if(FieldToSearchIsNumeric[k]){
            //if current search item is numeric, parse it to isolate user-supplied names/text/strings in preparation for evaluation 
            if( 
              UserSearchField.includes("<") ||
              UserSearchField.includes(">") ||
              UserSearchField.includes("=") ||
              UserSearchField.includes("-") ||
              UserSearchField.includes("(") ||
              UserSearchField.includes(")") ||
              UserSearchField.includes("&") ||
              UserSearchField.includes("|") 
              ){
              //user search field contains numerical operatiions, form expression for evaluation
              NumericData = data[SearchElementNumber][FieldToSearch[k]].normalize('NFD').replace(/[\u0300-\u036f]/g, "");
              if (NumericData.includes(".")) {
                // if this string isn't a valid integer (contains one or more "." or "-"" ), extract leading digits
                NumericData = NumericData.split(".", 1)[0];
              } else if (NumericData.includes("-")) {
                NumericData = NumericData.split("-", 1)[0];
              }
              NumAsString = NumericData.toString();
              //first, remove anything but numbers and operators, then hide any "<=" or ">=" (eg: user may put "1st" or "2nd" or "1st English" in the edition number field).
              UserSearchField = UserSearchField.replace(/[a-zA-Z]/g, '');
              UserSearchField = UserSearchField.replace(/<=/g, "LTE");
              UserSearchField = UserSearchField.replace(/>=/g, "GTE");
              //process >, <, =, then return "<=" or ">="
              UserSearchField = UserSearchField.replace(/</g, NumAsString+"<");
              UserSearchField = UserSearchField.replace(/>/g, NumAsString+">");
              UserSearchField = UserSearchField.replace(/=/g, NumAsString+"==");
              UserSearchField = UserSearchField.replace(/LTE/g, NumAsString+"<=");
              UserSearchField = UserSearchField.replace(/GTE/g, NumAsString+">=");
              UserSearchField = UserSearchField.replace(/-/g, "<="+NumAsString+"&"+NumAsString+"<=");
              try{SearchHit = eval(UserSearchField);}
              catch(err){};
              if(!SearchHit && ANDfield){
                //current element fails required "AND" search field, move to the next element
                continue ElementLoop;
              }else if(SearchHit && !ANDfield){
                //current element meats an "OR" search field, leave search loop, and move to code that logs current element to the search hit list
                NumberOfHits = NumberOfHits + 1;
              SearchItemHits[NumberOfHits] = SearchValueInputBoxNumber[k];
                break SearchItemLoop;
              }else if(SearchHit){
                NumberOfHits = NumberOfHits + 1;
                SearchItemHits[NumberOfHits] = SearchValueInputBoxNumber[k];
                continue SearchItemLoop;
              }
            }else{
              //user search item is a single number, so look only for an exact match
              NormalizedFieldToSearch = data[SearchElementNumber][FieldToSearch[k]].normalize('NFD').replace(/[\u0300-\u036f]/g, "");
              if (RegExp(UserSearchField).test(NormalizedFieldToSearch)) {
                SearchHit = true;
              } else {
                SearchHit = false;
              }
              if(!SearchHit && ANDfield){
                //current element fails required "AND" search field, move to the next element
                continue ElementLoop;
              }else if(SearchHit && !ANDfield){
                //current element meats an "OR" search field, leave search loop, and move to code that logs current element to the search hit list
                NumberOfHits = NumberOfHits + 1;
                SearchItemHits[NumberOfHits] = SearchValueInputBoxNumber[k];
                break SearchItemLoop;
              }else if(SearchHit){
                NumberOfHits = NumberOfHits + 1;
                SearchItemHits[NumberOfHits] = SearchValueInputBoxNumber[k];
                continue SearchItemLoop;
              }
            }
          }else{
            //if search field is neither alphabetic nor numeric, only data presence is detected
            // if(FieldToSearch[k]=="BP_Image"){
            //   //specical processing to see if supplied image URL actually fetches an image
            //     // var img = new Image();
            //     // var good;
            //     // good = true;
            //     // img.src = data[SearchElementNumber][FieldToSearch[k]];
            //     // img.onerror = ()=>{if(UserSearchField == "?"){SearchItemLoop}}
            //     // SearchHit = true;
            //     // NumberOfHits = NumberOfHits + 1;
            //     // SearchItemHits[NumberOfHits] = SearchValueInputBoxNumber[k];
            // }else{
              if(UserSearchField == "?" && data[SearchElementNumber][FieldToSearch[k]] != ""){
                SearchHit = true;
                NumberOfHits = NumberOfHits + 1;
                SearchItemHits[NumberOfHits] = SearchValueInputBoxNumber[k];
              }else if(UserSearchField == "!?" && data[SearchElementNumber][FieldToSearch[k]] == ""){
                SearchHit = true;
                NumberOfHits = NumberOfHits + 1;
                SearchItemHits[NumberOfHits] = SearchValueInputBoxNumber[k];
              }
            // }
          }
        }
      }
      //if we've hit on all the search items required, mark element as a search hit and its (summary) properties to the search results content
      if (SearchHit) {
        let OrderNumber = 0;
        if(ChronologicalSort){OrderNumber= ChronologicalOrder[SearchElementNumber]}
        else if(AlphabeticalSort){OrderNumber= AlphabeticalOrder[SearchElementNumber]}

        // put a SearchMarker on the timeline at screen bottom. The display ordering has to be used to get the proper locations.
        CurrentPlaceMarker = document.getElementById("PlaceMarkerElement"+SearchElementNumber.toString());
        CurrentPlaceMarker.style.opacity = "1";
        CurrentPlaceMarker.style.zIndex = "10";
        // search marker is image, size must be expressed in pixels. Derive from vw and vh and make sure it's square.
        ImgHeight = vhTOpx(2);
        ImgWidth = pxTOvw(ImgHeight);
        ImgWidth = vwTOpx(ImgWidth);
        CurrentPlaceMarker.innerHTML = "<img src=\"BP Appearance/magnifying-glass RED.png\" width="+ImgWidth.toString()+"px height="+ImgHeight.toString()+"px >";

        // increase hit count
        NumberOfElementSearchHits = NumberOfElementSearchHits + 1;
        // log the search hit element number
        SearchHits[NumberOfElementSearchHits] = SearchElementNumber;
        // add bookmark element to search container 
        NewDiv = document.createElement('div');
        NewDiv.setAttribute('class', 'CurrentSearchHit');

        if( data[SearchElementNumber]["BP_Short or Common Title"] != null && data[SearchElementNumber]["BP_Author full name"] != null){
          NewDiv.innerHTML = OrderNumber + "  " + data[SearchElementNumber]["BP_Short or Common Title"].toUpperCase() + " - " + data[SearchElementNumber]["BP_Date Published"] + "<br>" + data[SearchElementNumber]["BP_Author full name"];
        }else{
          NewDiv.innerHTML = OrderNumber + "  " + data[SearchElementNumber]["BP_Title"].toUpperCase() + " - " + data[SearchElementNumber]["BP_Date Published"] + "<br>" + data[SearchElementNumber]["BP_Author last name first"];
        }
        document.getElementById("SearchResultsContainer").appendChild(NewDiv);

        let ButtonDiv = document.createElement('div');
        ButtonDiv.setAttribute('class', 'CurrentSearchHitButtons');
        document.getElementsByClassName("CurrentSearchHit")[NumberOfElementSearchHits].appendChild(ButtonDiv);
        
        //add delete search hit button to search hit entry
        NewButton = document.createElement('button');
        NewButton.setAttribute("class", "RemoveSearchHitButton");
        NewButton.setAttribute("onclick", "RemoveSearchHit(" + SearchElementNumber.toString() + ")");
        NewButton.innerHTML = "<img class='ListButtonImages' src='BP Appearance/Bin.png' alt='Some png'>";
        ButtonDiv.appendChild(NewButton);

        //add bookmark button to search hit entry
        NewButton = document.createElement('button');
        NewButton.setAttribute("id", "BookmarkSearchHitButton"+NumberOfElementSearchHits.toString());
        NewButton.setAttribute("class", "BookmarkSearchHitButton");
        NewButton.setAttribute("onclick", "BookmarkThisElement(" + SearchElementNumber.toString() +","+ NumberOfElementSearchHits.toString() +" )" );
        NewButton.innerHTML = "<img class='ListButtonImages' src='BP Appearance/bookmarkFull.png' alt='Some png'>"
        ButtonDiv.appendChild(NewButton)

        //add ID search field button
        NewButton = document.createElement('button');
        NewButton.setAttribute("class", "ShowSearchFieldsHitButton");
        NewButton.setAttribute("onclick", "MapHitsToSearchField(" + NumberOfElementSearchHits.toString() + ")" );
        NewButton.setAttribute("name", "Open"); //attribute name is used to toggle this button on/off in the called function
        NewButton.innerHTML = "<img class='ListButtonImages' src='BP Appearance/Target.png' alt='Some png'>"
        ButtonDiv.appendChild(NewButton)

        //add GOTO details button
        NewButton = document.createElement('button');
        NewButton.setAttribute("class", "ShowSearchFullElementButton");
        NewButton.setAttribute("onclick", "OpenElementDetailData(" + SearchElementNumber.toString() + ")" );
        NewButton.innerHTML = "<img class='ListButtonImages' src='BP Appearance/Full.png' alt='Some png'>"
        ButtonDiv.appendChild(NewButton)

        //record search hit fields.
        SearchHitMap[NumberOfElementSearchHits] = [];
        SearchHitMap[NumberOfElementSearchHits][0] = NumberOfHits;
        for( mm=1; mm<=NumberOfHits; mm++){
          SearchHitMap[NumberOfElementSearchHits][mm] = SearchItemHits[mm];
        }

      }
    }
    //remove dummy zero search hit. NB: array SearchHits and class array CurrentSearchHit are still zero-based.
    Node = document.getElementById("SearchResultsContainer");
    Node.removeChild(Node.firstChild);
    // SearchHits.splice(0, 1);
    document.getElementById("NumberOfSearchResults").innerHTML = NumberOfElementSearchHits.toString() + " Matches";
    
  }

  function RemoveSearchHit(SearchElementNumber) {
    // remove entry from Bookmark array.
    if(ChronologicalSort){OrderNumber= ChronologicalOrder[SearchElementNumber]}
        else if(AlphabeticalSort){OrderNumber= AlphabeticalOrder[SearchElementNumber]}

    for (m = 1; m <= NumberOfElementSearchHits; m++) {
      if (SearchHits[m] == SearchElementNumber) {
        SearchHits.splice(m, 1);
        DeleteHere = document.getElementById("SearchResultsContainer");
        DeleteHere.removeChild(DeleteHere.childNodes[m-1]);
        ElementHasBeenBookmarked = false;
        for (n = 1; n <= NumberOfBookMarks; n++) {
          if (BookMarks[n] == OrderNumber) {
            var ElementHasBeenBookmarked = true;
            break;
          }
        }
        // CurrentTransform = document.getElementsByClassName("PlaceMarkerElement")[OrderNumber].style.transform
        // document.getElementsByClassName("PlaceMarkerElement")[OrderNumber].style.transform = CurrentTransform + "translateY(6px) scale(0.666666,0.666666)";
        
        //turn off SearchMark highlight but convert to bookmark if search element was bookmarked
        CurrentPlaceMarkerElement = document.getElementById("PlaceMarkerElement"+OrderNumber.toString())
        if (ElementHasBeenBookmarked) {
          CurrentPlaceMarkerElement.style.opacity = "1.0";
          CurrentPlaceMarkerElement.style.color = "rgb(255, 255, 255)";
          CurrentPlaceMarkerElement.style.zIndex = "10";
          CurrentPlaceMarkerElement.innerHTML = String.fromCharCode(9670);
        } else {
          CurrentPlaceMarkerElement.style.opacity = "0.15";
          CurrentPlaceMarkerElement.style.color = "rgb(9, 182, 254)";
          CurrentPlaceMarkerElement.style.zIndex = "0";
          CurrentPlaceMarkerElement.innerHTML = String.fromCharCode(9670);
        }
        break;
      }
    }
    NumberOfElementSearchHits = NumberOfElementSearchHits - 1;
    if(NumberOfElementSearchHits > 0){
      document.getElementById("NumberOfSearchResults").innerHTML = "Found: "+NumberOfElementSearchHits.toString()+" Matches";
    }else{
      document.getElementById("NumberOfSearchResults").innerHTML = "Found: ";
    }
    if (NumberOfElementSearchHits < 0) {
    }
    //Search page is already open, so we need to reset the toggle state to bookmark page stays open
    SearchDisplayOn = false;
    ShowSearchManagement();
  }

  function BookmarkAllSearchHits() {
    if( BPfileHasBeenLoaded == false && BPCSVfileHasBeenLoaded == false){return}
    for (hits = 1; hits <= NumberOfElementSearchHits; hits++) {
      ElementNumber = SearchHits[hits];
      BookmarkThisElement(ElementNumber, hits.toString());
    }
  }

  function ClearAllSearchHits() {
    if( BPfileHasBeenLoaded == false && BPCSVfileHasBeenLoaded == false){return}
    for (m = 1; m <= NumberOfElementSearchHits; m++) {
      SearchElementNumber = SearchHits[m];
      ElementHasBeenBookmarked = false;
      if     (ChronologicalSort){OrderNumber = ChronologicalOrder[SearchElementNumber]}
      else if(AlphabeticalSort) {OrderNumber = AlphabeticalOrder[SearchElementNumber]};

      for (n = 1; n <= NumberOfBookMarks; n++) {
        if (BookMarks[n] == SearchElementNumber) {
          ElementHasBeenBookmarked = true;
          break;
        }
      }
      //turn off SearchMark highlight but convert to bookmark if search element was bookmarked
      CurrentPlaceMarker = document.getElementById("PlaceMarkerElement"+SearchElementNumber.toString());
      if (ElementHasBeenBookmarked) {
        CurrentPlaceMarker.style.opacity = "1.0";
        CurrentPlaceMarker.style.color = "rgb(255, 255, 255)";
        CurrentPlaceMarker.style.zIndex = "10";
        CurrentPlaceMarker.innerHTML = String.fromCharCode(9670);
      } else {
        CurrentPlaceMarker.style.opacity = "0.15";
        CurrentPlaceMarker.style.color = "rgb(9, 182, 254)";
        CurrentPlaceMarker.style.zIndex = "0";
        CurrentPlaceMarker.innerHTML = String.fromCharCode(9670);
        // remove border to searchhit elements
        CurrentPlaceMarker.style.border = "3px solid transparent";
      }
    }
    DeleteHere = document.getElementById("SearchResultsContainer");
    while (DeleteHere.firstChild) {
      DeleteHere.firstChild.remove();
    }
    NumberOfElementSearchHits = -1;
    document.getElementById("NumberOfSearchResults").innerHTML = "Found: ";
    //bookmark page is already open, so we need to reset the toggle state to bookmark page stays open
    SearchDisplayOn = false;
    SearchTable = document.getElementById("SearchTable");
    InputBoxes = SearchTable.getElementsByTagName("INPUT");
    ButtonBoxes = SearchTable.getElementsByTagName("Button");
    for (m = 1; m <= NumberOfSearchableFields; m++) {
      InputBoxes[m*2-1].style.backgroundColor = "rgb(255,255,255";
      ButtonBoxes[m*2-1].innerHTML = "<img class='SearchValueMapButtonImage' src='BP Appearance/Target.png' alt='Some png'>"
      ButtonBoxes[m*2-1].name = "Open";
    }
    ShowSearchManagement();
  }

  function ClearAllSearchFields() {
    if( BPfileHasBeenLoaded == false && BPCSVfileHasBeenLoaded == false){return}
    NumberOfElementSearchHits = -1;
    // document.getElementById("NumberOfSearchResults").innerHTML = "Filtered Items: ";
    //bookmark page is already open, so we need to reset the toggle state to bookmark page stays open
    SearchTable = document.getElementById("SearchTable");
    InputBoxes = SearchTable.getElementsByTagName("INPUT");
    ButtonBoxes = SearchTable.getElementsByTagName("Button");
    // empty all the filter input
    for (m = 0; m < NumberOfSearchableFields; m++) {
      InputBoxes[m*2].style.backgroundColor = "rgb(255,255,255";
      InputBoxes[m*2].checked = false;
      InputBoxes[m*2+1].value = '';
    }
    NumberOfElementFilterHits = 0;
    NumberOfSpaces = 30-16-NumberOfElementFilterHits.toString().length
    NumberOfSearchFieldBoxesChecked = 0;
    let ButtonElement = document.getElementById('StartSearchButton');
    ButtonElement.disabled = true;
    ButtonElement.style.opacity = '0.5';
    ButtonElement.style.pointerEvents = 'none';
  }
  function MapHitsToSearchField(HitNumber){
    //toggle button image and background color of those search field check boxes that specfied element has hit
    Button = document.getElementsByClassName("ShowSearchFieldsHitButton")[HitNumber-1]; // index has to account for zero-based array of ShowSearchFieldsHitButton class
    if( Button.name == "Open") {
      Button.innerHTML = "<img class='img' src='BP Appearance/TargetRed.png' alt='Some png'>"
      SearchTable = document.getElementById("SearchTable");
      InputBoxes = SearchTable.getElementsByTagName("INPUT");
      let NumberOfSearchHits = SearchHitMap[HitNumber][0];
      for (m = 1; m <= NumberOfSearchHits ; m++) {
        InputBoxNumber = SearchHitMap[HitNumber][m];
        InputBoxes[InputBoxNumber].style.backgroundColor = "var(--InputBoxChangedColor)";
      }
      Button.name = "Closed";
     }else{
      Button.innerHTML = "<img class='ListButtonImages' src='BP Appearance/Target.png' alt='Some png'>"
      SearchTable = document.getElementById("SearchTable");
      InputBoxes = SearchTable.getElementsByTagName("INPUT");
      let NumberOfSearchHits = SearchHitMap[HitNumber][0];
      for (m = 1; m <= NumberOfSearchHits ; m++) {
        InputBoxNumber = SearchHitMap[HitNumber][m]; 
        InputBoxes[InputBoxNumber].style.backgroundColor = "rgb(255,255,255";
      }
      Button.name = "Open";
    } 
  }  

  function MapSearchFieldToHits(SearchFieldNumber){
    //find any matches of search hits results with current search field number
    SearchTable = document.getElementById("SearchTable");
    InputBoxes = SearchTable.getElementsByTagName("Button");
    SearchFieldButton = InputBoxes[SearchFieldNumber*2-1]; // each field has 2 boxes. we want the odd-numbered one here. index has to account for zero-based array of SearchFields
    if(SearchFieldButton.name == "Open") {
      AnyHits = false;
      for (hits = 1; hits <= NumberOfElementSearchHits; hits++) {
        let NumberOfSearchHits = SearchHitMap[hits][0];
        for (m = 1; m <= NumberOfSearchHits ; m++) {
          if( SearchFieldNumber*2-1 == SearchHitMap[hits][m] ){   //SearchHitMap[hits][m] contains a search field input box number. have to convert to an input box number
            //match found, toggle color of target button for matched item in search hit results list
            SearchResult =  document.getElementsByClassName("CurrentSearchHit")[hits-1] // index has to account for zero-based array of ShowSearchFieldsHitButton class
            SearchResult.style.border = "3px solid var(--AlertColor)";
            AnyHits = true;
          };
        }
      }
      if(AnyHits){
        SearchFieldButton.innerHTML = "<img class='SearchValueMapButtonImage' src='BP Appearance/TargetRed.png' alt='Some png'>"
        SearchFieldButton.name = "Closed";
      }
    }else{
      AnyHits = false;
      for (hits = 1; hits <= NumberOfElementSearchHits; hits++) {
          let NumberOfSearchHits = SearchHitMap[hits][0];
          for (m = 1; m <= NumberOfSearchHits ; m++) {
            if( SearchFieldNumber*2-1 == SearchHitMap[hits][m] ){
              //match found, toggle color of target button for matched item in search hit results list
              SearchResult =  document.getElementsByClassName("CurrentSearchHit")[hits-1] // index has to account for zero-based array of ShowSearchFieldsHitButton class
              SearchResult.style.border = "2px solid var(--Gray50)";
              AnyHits = true;
            };
          }
        }
      if(AnyHits){
        SearchFieldButton.innerHTML = "<img class='SearchValueMapButtonImage' src='BP Appearance/Target.png' alt='Some png'>"
        SearchFieldButton.name = "Open";
      }
    }
  }  

  function GetSearchList(FieldNumber){
    let SearchTable = document.getElementById("SearchTable");
    let InputBoxes = SearchTable.getElementsByTagName("Button");
    CurrentSearchFieldListNumber = FieldNumber;

    if(BPHeadersOfUserData[FieldNumber]=="BP_Author full name" || BPHeadersOfUserData[FieldNumber]=="BP_Author last name first"){
      ListTitle = "Authors";
      ListName = ListOfAuthors;
      NumInList = NumberOfAuthorsInList;
    }else if(BPHeadersOfUserData[FieldNumber]=="BP_Translator or Editor" || BPHeadersOfUserData[FieldNumber]=="BP_Translator" || BPHeadersOfUserData[FieldNumber]=="BP_Editor") {
      ListTitle = "Translators or Editors";
      ListName = ListOfTransOrEdit;
      NumInList = NumberOfTransOrEditInList;
    }else if(BPHeadersOfUserData[FieldNumber]=="BP_Publisher or Seller"){
      ListTitle = "Publishers";
      ListName = ListOfPublishers;
      NumInList = NumberOfPublishersInList;
    }else if(BPHeadersOfUserData[FieldNumber]=="BP_Publication Place"){
      ListTitle = "Publication Places";
      ListName = ListOfPublicationPlaces;
      NumInList = NumberOfPublicationPlacesInList;
    }else if(BPHeadersOfUserData[FieldNumber]=="BP_Printer"){
      ListTitle = "Printers";
      ListName = ListOfPrinters;
      NumInList = NumberOfPrintersInList;
    }else if(BPHeadersOfUserData[FieldNumber]=="BP_Format and Size" || BPHeadersOfUserData[FieldNumber]=="BP_Format"){
      ListTitle = "Formats";
      ListName = ListOfFormats;
      NumInList = NumberOfFormatsInList;
    }else{return}

    SearchFieldButton = InputBoxes[FieldNumber*2-2]; // each field has 2 boxes. we want the even-numbered one here. index has to account for zero-based array of SearchFields
    if( SearchFieldButton.name=='Open'){
      //toggle on: populate available search list and make list window visible
      SearchFieldButton.innerHTML = "<img class='GetSearchListButtonImage' src='BP Appearance/ListRed.png' alt='Some png'>";
      SearchFieldButton.name='Closed';
      document.getElementById('AvailableSearchLists').style.visibility = "visible";
      Listname = document.getElementById('AvailableSearchListsTitle');
      Listname.innerHTML = ListTitle
      SearchListElement = document.getElementById('AvailableSearchListsContent');
      List = document.createElement("ul");
      List.style = "list-style-type:none";
      List.style = "margin: 0";
      List.style = "padding: 0 0 0 10px";
      for (n = 0; n < NumInList; n++) {
        ListItem = document.createElement("li");
        ListItem.innerHTML =  ListName[n];
        ListItem.setAttribute('onclick' , 'PutTextListInSearchField('+FieldNumber.toString()+',\"'+ListName[n]+'\")' );
        List.appendChild(ListItem);
      } 
      SearchListElement.appendChild(List);
      dragElement(document.getElementById("AvailableSearchListsDrag"));
    }else{
      //toggle off: remove list and hide list window
      SearchFieldButton.innerHTML = "<img class='GetSearchListButtonImage' src='BP Appearance/List.png' alt='Some png'>";
      SearchFieldButton.name='Open';
      document.getElementById('AvailableSearchLists').style.visibility = "hidden";
      SearchListElement = document.getElementById('AvailableSearchListsContent');
      while (SearchListElement.hasChildNodes() && SearchListElement.children.length > 0) {
        SearchListElement.removeChild(SearchListElement.lastChild);
      }
    }
  }
  function CloseAvailableSearchLists(List){
    if(List=='Search'){
      InputBoxes = SearchTable.getElementsByTagName("Button");
      for (let m = 1; m <= NumUserDataFields; m++){
        SearchFieldButton = InputBoxes[m*2-2]; // each field has 2 boxes. we want the even-numbered one here. index has to account for zero-based array of SearchFields        
        if(BPHeadersOfUserData[m]=="BP_Author full name" || BPHeadersOfUserData[m]=="BP_Author last name first"){
        }else if(BPHeadersOfUserData[m]=="BP_Translator or Editor" || BPHeadersOfUserData[m]=="BP_Translator" || BPHeadersOfUserData[m]=="BP_Editor"){
        }else if(BPHeadersOfUserData[m]=="BP_Publisher or Seller"){
        }else if(BPHeadersOfUserData[m]=="BP_Publication Place"){
        }else if(BPHeadersOfUserData[m]=="BP_Printer"){
        }else if(BPHeadersOfUserData[m]=="BP_Format and Size" || BPHeadersOfUserData[m]=="BP_Format"){
        }else{continue}
        SearchFieldButton.innerHTML = "<img class='GetSearchListButtonImage' src='BP Appearance/List.png' alt='Some png'>";
        SearchFieldButton.name='Open';
      }
      document.getElementById('AvailableSearchLists').style.visibility = "hidden";      
      SearchListElement = document.getElementById('AvailableSearchListsContent');
      while (SearchListElement.hasChildNodes() && SearchListElement.children.length > 0) {
        SearchListElement.removeChild(SearchListElement.lastChild);
      }      
    }else{
      InputBoxes = FilterTable.getElementsByTagName("Button");
      for (let m = 1; m <= NumUserDataFields; m++){
        FilterFieldButton = InputBoxes[m-1]; // each field has 2 boxes. we want the even-numbered one here. index has to account for zero-based array of FilterFields
        if(BPHeadersOfUserData[m]=="BP_Author full name" || BPHeadersOfUserData[m]=="BP_Author last name first"){
        }else if(BPHeadersOfUserData[m]=="BP_Translator or Editor" || BPHeadersOfUserData[m]=="BP_Translator" || BPHeadersOfUserData[m]=="BP_Editor"){
        }else if(BPHeadersOfUserData[m]=="BP_Publisher or Seller"){
        }else if(BPHeadersOfUserData[m]=="BP_Publication Place"){
        }else if(BPHeadersOfUserData[m]=="BP_Printer"){
        }else if(BPHeadersOfUserData[m]=="BP_Format and Size" || BPHeadersOfUserData[m]=="BP_Format"){
        }else{continue}
        FilterFieldButton.innerHTML = "<img class='GetFilterListButtonImage' src='BP Appearance/List.png' alt='Some png'>";
        FilterFieldButton.name='Open';
      }
      document.getElementById('AvailableSearchLists').style.visibility = "hidden";
      FilterListElement = document.getElementById('AvailableSearchListsContent');
      while (FilterListElement.hasChildNodes() && FilterListElement.children.length > 0) {
        FilterListElement.removeChild(FilterListElement.lastChild);
      }      
    }
  }

  function PutTextListInSearchField(FieldNumber,TextFromList){
    SearchTable = document.getElementById("SearchTable");
    InputBoxes = SearchTable.getElementsByTagName("INPUT");
      ListText = TextFromList.trim().split(",")
      ListText[0] = ListText[0].normalize('NFD').replace(/[\u0300-\u036f]/g, "");
      InputBoxes[FieldNumber*2-1].value += ListText[0]
  }

  function SaveSearchList(){
    // build text string for output
    //Analysls = document.getElementById('BookmarkList')
    SearchList = ""
    SearchList = '<strong>List of Search Entries</strong><br><br>';
    for (m = 1; m <= NumberOfElementSearchHits; m++) {
      SearchElementNumber = SearchHits[m];
      // if(ChronologicalSort){OrderElementNumber = ChronologicalOrder[SearchElementNumber]}
      // else if(AlphabeticalSort){OrderElementNumber = AlphabeticalOrder[SearchElementNumber]};
        SearchList += document.getElementById("BriefBookTitle"+SearchElementNumber.toString()).innerHTML.toUpperCase() + " - " +
                      document.getElementById("PublicationYear"+SearchElementNumber.toString()).innerHTML + "<br>" +
                      document.getElementById("Author"+SearchElementNumber.toString()).innerHTML+"<br><br>";      
    }
    SearchList = SearchList.replace(/<br>/gi, String.fromCharCode(13));
    SearchList = SearchList.replace(/<strong>/gi, "");
    SearchList = SearchList.replace(/<\/strong>/gi, "");
    SearchList = SearchList.replace(/<em>/gi, "");
    SearchList = SearchList.replace(/<\/em>/gi, "");    
    download(SearchList, "ListOfSearchEntries.txt","text/plain");
  }

  //#endregion

// filter management -->
//#region
function ShowFilterManagement() {
  if (FilterDisplayOn == false) {
    document.getElementById('FilterManagement').style.visibility = "visible";
    FilterDisplayOn = true;
    dragElement(document.getElementById("FilterManagementDrag"));
  }
  else {
    document.getElementById('FilterManagement').style.visibility = "hidden";
    // document.getElementById('SearchHelp').style.visibility = "hidden";      
    FilterDisplayOn = false;
  }
}

function CloseFilterManagement() {
  document.getElementById('FilterManagement').style.visibility = "hidden";
  FilterDisplayOn = false;
  ScrollIntegerPx = CurrentMinimumXPosition;
  document.getElementById("MainTimeline").scrollLeft = ScrollIntegerPx;
  document.getElementById("TimeLineMarker").scrollLeft = ScrollIntegerPx;
  CloseAvailableSearchLists('Filter');
  // check on navigation slider limits
  if(ChronologicalSort == "true"){    
    // check on navigation slider limits
    CurrentLeftPositionOfTimeline = MinYear + (ScrollIntegerPx) / CurrentTimelinePixelsPerUnit;
    SliderValue = 50 * ((CurrentLeftPositionOfTimeline - MinYear) * NavigationScalePixelsPerUnit) / (screen.width - NavigationWindowWidthInPX / 2);
    document.getElementById("NavigateWindowSlider").value = SliderValue;
    localStorage.setItem("TimeLineScrollValue", SliderValue.toString());
  }else{
    CurrentLeftPositionOfTimeline = MinAlphaNumber + (ScrollIntegerPx) / CurrentTimelinePixelsPerUnit;
    SliderValue = 50 * ((CurrentLeftPositionOfTimeline - MinAlphaNumber) * NavigationScalePixelsPerUnit) / (screen.width - NavigationWindowWidthInPX / 2);
    document.getElementById("NavigateWindowSlider").value = SliderValue;
    localStorage.setItem("TimeLineScrollValue", SliderValue.toString());
  }
}

function StartFilter() {
  FilterTable = document.getElementById("FilterTable");
  InputBoxes = FilterTable.getElementsByTagName("INPUT");
  NumFilterItems = 0
  for (m = 0; m < InputBoxes.length; m++) {
    if (InputBoxes[m].getAttribute('type') == 'checkbox' && InputBoxes[m].checked == true && InputBoxes[m + 1].value.length != 0) {
      NumFilterItems = NumFilterItems + 1;
      // in each row of search filed table:  check box, field name, input box. so the input boxes are [0,1], [2,3], [4,5], ...
      // the checkbox and search field alternate in the container, so the checkboxes are the even-numbered elements.        
      FieldToFilter[NumFilterItems] = BPHeadersOfUserData[(m + 2) / 2];
      FieldToFilterIsAlphabetic[NumFilterItems] = BPHeadersOfUserDataIsAlphabetic[(m + 2) / 2];
      FieldToFilterIsNumeric[NumFilterItems] = BPHeadersOfUserDataIsNumeric[(m + 2) / 2];
      FilterValue[NumFilterItems] = InputBoxes[m + 1].value;
      FilterValueInputBoxNumber[NumFilterItems] = m + 1;
      FieldFilterNumber[NumFilterItems] = m; //(m + 2) / 2
    }
  }
  FilterBookData(NumFilterItems);
}

function EndFilter(){
  DisplayElement.fill(true);
  if(ChronologicalSort){
    ProcessDataChronologically("Filter");
  }
  else if(AlphabeticalSort){
    ProcessDataAlphabetically();
  }
  document.getElementById("ScaleSlider").value = 0;
  document.getElementById("NumberOfFilterResults").innerHTML = "Filter Fields";
  //NumberOfFilterFieldBoxesChecked = 0;
  //let ButtonElement = document.getElementById('StartFilterButton');
  //ButtonElement.disabled = true;
  //ButtonElement.style.opacity = '0.5';
  //ButtonElement.style.pointerEvents = 'none';
}

function FilterBookData(NumFilterItems) {
  if(NumFilterItems==0){return}
  //check search phrase(s) for balanced bracketing.
  FilterTable = document.getElementById("FilterTable");
  InputBoxes = FilterTable.getElementsByTagName("INPUT");
  for (k = 1; k <= NumFilterItems; k++) {
    TestString = FilterValue[k];
    if(!areParenthesesBalanced(TestString)){
      InputBoxes[FilterValueInputBoxNumber[k]].style.color = "rgb(255,0,0)";
      NumberOfElementFilterHits = 0;
      return;        
    }else{
      InputBoxes[FilterValueInputBoxNumber[k]].style.color = "rgb(0,0,0)";
    }
  }

    NumSortedFilterItems = 0
    NumberOfElementFilterHits = 0
    NumberOfElementFilterHitsTop = 0
    NumberOfElementFilterHitsBottom = 0
    // SearchHits[NumberOfElementFilterHits] = 0;
    let SortedFilterItems = [];
    let FilterItemHits = [];
    //build pointer that puts all "OR" filter fields at the of the list of filter elements and all "AND" at the bottom.
    //This allows early out of the FilterItemLoop if either: a filter hit on an "OR" field or a fail on any "AND" field
    for (k = 1; k <= NumFilterItems; k++) {
      FilterValue[k] = FilterValue[k].trim();
      if(FilterValue[k].charAt(0) == "|" ){
        NumberOfElementFilterHitsTop = NumberOfElementFilterHitsTop + 1;
        SortedFilterItems[NumberOfElementFilterHitsTop] = k;
      }else{
        NumberOfElementFilterHitsBottom = NumberOfElementFilterHitsBottom + 1;
        SortedFilterItems[NumFilterItems+1-NumberOfElementFilterHitsBottom] = k
      }
    }
  // clear last Filter hits, if any, and make an empty node at zero
  FilterElementLoop:
  for (let FilterElementNumber = 1; FilterElementNumber <= NumOfTimelineElements; FilterElementNumber++) {
    FilterHit = false;
    NumberOfHits = 0;
    let localData = data[FilterElementNumber]

    FilterItemLoop:
    for (kk = 1; kk <= NumFilterItems; kk++) {
      k = SortedFilterItems[kk]
      // parse appropriate filter element for name and numbers
      let ElementsFromFilterFieldFound = [];
      UserFilterField = FilterValue[k]; 
      // see if user has specified this as an "AND" filter by testing for the the presence of "!" as the 1st character.
      // If NOT, then failure to find this filter value in the current element is enough to disqualify current element. Filter can be halted
      // If so, then "OR" means if filter value is in current element, that's enough to qualify current element. Filter can be halted.
      if(UserFilterField.charAt(0) != "|"){
        ANDfield = true;
      }else{
        //remove the "|" from the front of the filter field
        UserFilterField = UserFilterField.slice(1);
        ANDfield = false;
      }
      //if the user filter field is ONLY "?", then simply detect the presence of anything but null/blank/empty in the data
      if(UserFilterField==="?"){
        for( m = 0; m < UserFilterField.length; m++){
          if(localData[FieldToFilter[k]] != ""){
            FilterHit = true;
          }
        }
      }else if(UserFilterField==="!?"){
        for( m = 0; m < UserFilterField.length; m++){
          if(localData[FieldToFilter[k]] == ""){
            FilterHit = true;
          }
        }
      }else{
        if( FieldToFilterIsAlphabetic[k] ){
          var t0=performance.now();
          //if current filter item is alphabetic, parse it to isolate user-supplied names/text/strings in preparation for evaluation 
          ElementsFromFilterField = UserFilterField; 
          // ElementsFromFilterField = ElementsFromFilterField.split(/[.:;?!~,&|()<>{}\[\]\r\n/\\]+/);
          ElementsFromFilterField = ElementsFromFilterField.split(/[:;?!~,&|()<>{}\[\]\r\n/\\]+/);  // no "." in regex ??
          //remove empty strings from array resulting from the regexp split
          ElementsFromFilterField = ElementsFromFilterField.filter(Boolean);
          // alert(ElementsFromFilterField)
          NormalizedFieldToFilter = localData[FieldToFilter[k]].normalize('NFD').replace(/[\u0300-\u036f]/g, "");
          for( m = 0; m < ElementsFromFilterField.length; m++){
            //test each ElementsFromFilterField, flagging found/not found as true/false, reconstructing parsed filter field with true/false substituted for user-specified words
            if (RegExp(ElementsFromFilterField[m].trim()).test(NormalizedFieldToFilter)) {
              UserFilterField = UserFilterField.replace(ElementsFromFilterField[m], "true");
            } else {
              UserFilterField = UserFilterField.replace(ElementsFromFilterField[m], "false");
            }
          }
          FilterHit = eval(UserFilterField);
          if(!FilterHit && ANDfield){
            //current element fails required "AND" filter field, move to the next element
            DisplayElement[FilterElementNumber] = false;
              continue FilterElementLoop;
          }else if(FilterHit && !ANDfield){
            //current element meats an "OR" filter field, leave filter loop, and move to code that logs current element to the filter hit list
            NumberOfHits = NumberOfHits + 1;
            FilterItemHits[NumberOfHits] = FilterValueInputBoxNumber[k];
              break FilterItemLoop;
          }else if(FilterHit){
            NumberOfHits = NumberOfHits + 1;
            FilterItemHits[NumberOfHits] = FilterValueInputBoxNumber[k];
              continue FilterItemLoop;
          }
        }else if(FieldToFilterIsNumeric[k]){
          //if current filter item is numeric, parse it to isolate user-supplied names/text/strings in preparation for evaluation 
          if( 
            UserFilterField.includes("<") ||
            UserFilterField.includes(">") ||
            UserFilterField.includes("=") ||
            UserFilterField.includes("-") ||
            UserFilterField.includes("(") ||
            UserFilterField.includes(")") ||
            UserFilterField.includes("&") ||
            UserFilterField.includes("|") 
            ){
            //user filter field contains numerical operatiions, form expression for evaluation
            NumericData = localData[FieldToFilter[k]].normalize('NFD').replace(/[\u0300-\u036f]/g, "");
            if (NumericData.includes(".")) {
              // if this string isn't a valid integer (contains one or more "." or "-"" ), extract leading digits
              NumericData = NumericData.split(".", 1)[0];
            } else if (NumericData.includes("-")) {
              NumericData = NumericData.split("-", 1)[0];
            }
            NumAsString = NumericData.toString();
            //first, remove anything but numbers and operators, then hide any "<=" or ">=" (eg: user may put "1st" or "2nd" or "1st English" in the edition number field).
            UserFilterField = UserFilterField.replace(/[a-zA-Z]/g, '');
            UserFilterField = UserFilterField.replace(/<=/g, "LTE");
            UserFilterField = UserFilterField.replace(/>=/g, "GTE");
            //process >, <, =, then return "<=" or ">="
            UserFilterField = UserFilterField.replace(/</g, NumAsString+"<");
            UserFilterField = UserFilterField.replace(/>/g, NumAsString+">");
            UserFilterField = UserFilterField.replace(/=/g, NumAsString+"==");
            UserFilterField = UserFilterField.replace(/LTE/g, NumAsString+"<=");
            UserFilterField = UserFilterField.replace(/GTE/g, NumAsString+">=");
            UserFilterField = UserFilterField.replace(/-/g, "<="+NumAsString+"&"+NumAsString+"<=");
            try{FilterHit = eval(UserFilterField);}
            catch(err){};
            if(!FilterHit && ANDfield){
              //current element fails required "AND" filter field, move to the next element
              DisplayElement[FilterElementNumber] = false;
              continue FilterElementLoop;
            }else if(FilterHit && !ANDfield){
              //current element meats an "OR" filter field, leave filter loop, and move to code that logs current element to the filter hit list
              NumberOfHits = NumberOfHits + 1;
            FilterItemHits[NumberOfHits] = FilterValueInputBoxNumber[k];
              break FilterItemLoop;
            }else if(FilterHit){
              NumberOfHits = NumberOfHits + 1;
              FilterItemHits[NumberOfHits] = FilterValueInputBoxNumber[k];
              continue FilterItemLoop;
            }
          }else{
            //user filter item is a single number, so look only for an exact match
            NormalizedFieldToFilter = localData[FieldToFilter[k]].normalize('NFD').replace(/[\u0300-\u036f]/g, "");
            if (RegExp(UserFilterField).test(NormalizedFieldToFilter)) {
              FilterHit = true;
            } else {
              FilterHit = false;
            }
            if(!FilterHit && ANDfield){
              //current element fails required "AND" filter field, move to the next element
              DisplayElement[FilterElementNumber] = false;
              continue FilterElementLoop;
            }else if(FilterHit && !ANDfield){
              //current element meats an "OR" filter field, leave filter loop, and move to code that logs current element to the filter hit list
              NumberOfHits = NumberOfHits + 1;
              FilterItemHits[NumberOfHits] = FilterValueInputBoxNumber[k];
              break FilterItemLoop;
            }else if(FilterHit){
              NumberOfHits = NumberOfHits + 1;
              FilterItemHits[NumberOfHits] = FilterValueInputBoxNumber[k];
              continue FilterItemLoop;
            }
          }
        }else{
          //if filter field is neither alphabetic nor numeric, only data presence is detected
          // if(FieldToFilter[k]=="BP_Image"){
          //   //specical processing to see if supplied image URL actually fetches an image
          //     // var img = new Image();
          //     // var good;
          //     // good = true;
          //     // img.src = data[FilterElementNumber][FieldToFilter[k]];
          //     // img.onerror = ()=>{if(UserFilterField == "?"){FilterItemLoop}}
          //     // FilterHit = true;
          //     // NumberOfHits = NumberOfHits + 1;
          //     // FilterItemHits[NumberOfHits] = FilterValueInputBoxNumber[k];
          // }else{
            if(UserFilterField == "?" && localData[FieldToFilter[k]] != ""){
              FilterHit = true;
              NumberOfHits = NumberOfHits + 1;
              FilterItemHits[NumberOfHits] = FilterValueInputBoxNumber[k];
            }else if(UserFilterField == "!?" && localData[FieldToFilter[k]] == ""){
              FilterHit = true;
              NumberOfHits = NumberOfHits + 1;
              FilterItemHits[NumberOfHits] = FilterValueInputBoxNumber[k];
            }
          // }
        }
      }
    }
    //if we've hit on all the filter items required, mark element as a filter hit and its (summary) properties to the filter results content
    if (FilterHit) {
      // if(ChronologicalSort){OrderNumber= ChronologicalOrder[FilterElementNumber]}
      // else if(AlphabeticalSort){OrderNumber= AlphabeticalOrder[FilterElementNumber]}
      DisplayElement[FilterElementNumber] = true;
      // increase hit count
      NumberOfElementFilterHits = NumberOfElementFilterHits + 1;
      // log the filter hit element number
      FilterHits[NumberOfElementFilterHits] = FilterElementNumber;

      //record filter hit fields.
      FilterHitMap[NumberOfElementFilterHits] = [];
      FilterHitMap[NumberOfElementFilterHits][0] = NumberOfHits;
      for( mm=1; mm<=NumberOfHits; mm++){
        FilterHitMap[NumberOfElementFilterHits][mm] = FilterItemHits[mm];
      }
    }else{
      DisplayElement[FilterElementNumber] = false;
    }
  }
  //remove dummy zero filter hit. NB: array FilterHits and class array CurrentFilterHit are still zero-based.
  // Node = document.getElementById("FilterResultsContainer");
  // Node.removeChild(Node.firstChild);
  FilterHits.splice(0, 1);
  NumberOfSpaces = 30-16-NumberOfElementFilterHits.toString().length
  document.getElementById("NumberOfFilterResults").innerHTML = NumberOfElementFilterHits.toString()+" Items";  //+ "&nbsp".repeat(NumberOfSpaces)
  // refresh display using filtered element list
  if(ChronologicalSort){
    ProcessDataChronologically("Filter");
  }else{
    ProcessDataAlphabetically("Filter");
  }
}

function ClearAllFilterFields() {
  if( BPfileHasBeenLoaded == false && BPCSVfileHasBeenLoaded == false){return}  
  NumberOfElementSearchHits = -1;
  // document.getElementById("NumberOfSearchResults").innerHTML = "Filtered Items: ";
  //bookmark page is already open, so we need to reset the toggle state to bookmark page stays open
  FilterTable = document.getElementById("FilterTable");
  InputBoxes = FilterTable.getElementsByTagName("INPUT");
  ButtonBoxes = FilterTable.getElementsByTagName("Button");
  // empty all the filter input
  for (m = 0; m < NumberOfSearchableFields; m++) {
    InputBoxes[m*2].style.backgroundColor = "rgb(255,255,255";
    InputBoxes[m*2].checked = false;
    InputBoxes[m*2+1].value = '';
  }
  NumberOfElementFilterHits = 0;
  NumberOfSpaces = 30-16-NumberOfElementFilterHits.toString().length
  document.getElementById("NumberOfFilterResults").innerHTML = "Filter Fields";
  NumberOfFilterFieldBoxesChecked = 0;
  let ButtonElement = document.getElementById('StartFilterButton');
  ButtonElement.disabled = true;
  ButtonElement.style.opacity = '0.5';
  ButtonElement.style.pointerEvents = 'none';
}

function GetFilterList(FieldNumber){
  if( BPfileHasBeenLoaded == false && BPCSVfileHasBeenLoaded == false){return}
  let FilterTable = document.getElementById("FilterTable");
  let InputBoxes = FilterTable.getElementsByTagName("Button");
  CurrentSearchFieldListNumber = -FieldNumber;

  if(BPHeadersOfUserData[FieldNumber]=="BP_Author full name" || BPHeadersOfUserData[FieldNumber]=="BP_Author last name first"){
    ListTitle = "Authors";
    ListName = ListOfAuthors;
    NumInList = NumberOfAuthorsInList;
  }else if(BPHeadersOfUserData[FieldNumber]=="BP_Translator or Editor" || BPHeadersOfUserData[FieldNumber]=="BP_Translator" || BPHeadersOfUserData[FieldNumber]=="BP_Editor"){
    ListTitle = "Translators or Editors";
    ListName = ListOfTransOrEdit;
    NumInList = NumberOfTransOrEditInList;
  }else if(BPHeadersOfUserData[FieldNumber]=="BP_Publisher or Seller"){
    ListTitle = "Publishers";
    ListName = ListOfPublishers;
    NumInList = NumberOfPublishersInList;
  }else if(BPHeadersOfUserData[FieldNumber]=="BP_Publication Place"){
    ListTitle = "Publication Places";
    ListName = ListOfPublicationPlaces;
    NumInList = NumberOfPublicationPlacesInList;
  }else if(BPHeadersOfUserData[FieldNumber]=="BP_Printer"){
    ListTitle = "Printers";
    ListName = ListOfPrinters;
    NumInList = NumberOfPrintersInList;
  }else if(BPHeadersOfUserData[FieldNumber]=="BP_Format and Size" || BPHeadersOfUserData[FieldNumber]=="BP_Format"){
    ListTitle = "Formats";
    ListName = ListOfFormats;
    NumInList = NumberOfFormatsInList;
  }else{return}

  FilterFieldButton = InputBoxes[FieldNumber-1]; // each field has 2 boxes. we want the even-numbered one here. index has to account for zero-based array of FilterFields
  if( FilterFieldButton.name=='Open'){
    //toggle on: populate available filter list and make list window visible
    FilterFieldButton.innerHTML = "<img class='GetFilterListButtonImage' src='BP Appearance/ListRed.png' alt='Some png'>";
    FilterFieldButton.name='Closed';
    document.getElementById('AvailableSearchLists').style.visibility = "visible";
    Listname = document.getElementById('AvailableSearchListsTitle');
    Listname.innerHTML = ListTitle;
    FilterListElement = document.getElementById('AvailableSearchListsContent');
    List = document.createElement("ul");
    List.style = "list-style-type:none";
    List.style = "margin: 0";
    List.style = "padding: 0 0 0 10px";
    for (n = 0; n < NumInList; n++) {
      ListItem = document.createElement("li");
      ListItem.innerHTML =  ListName[n];
      ListItem.setAttribute('onclick' , 'PutTextListInFilterField('+FieldNumber.toString()+',\"'+ListName[n]+'\")' );
      List.appendChild(ListItem);
    } 
    FilterListElement.appendChild(List);
    dragElement(document.getElementById("AvailableSearchListsDrag"));
  }else{
    //toggle off: remove list and hide list window
    FilterFieldButton.innerHTML = "<img class='GetFilterListButtonImage' src='BP Appearance/List.png' alt='Some png'>";
    FilterFieldButton.name='Open';
    document.getElementById('AvailableSearchLists').style.visibility = "hidden";
    FilterListElement = document.getElementById('AvailableSearchListsContent');
    // FilterListElement.innerHTML = ""
    //revmoe filte list, but not the 1st child, which is the drag element
    while (FilterListElement.hasChildNodes() && FilterListElement.children.length > 0) {
      FilterListElement.removeChild(FilterListElement.lastChild);
    }
  }
}

function PutTextListInFilterField(FieldNumber,TextFromList){
  SearchTable = document.getElementById("FilterTable");
  InputBoxes = SearchTable.getElementsByTagName("INPUT");
    ListText = TextFromList.trim().split(",")
    ListText[0] = ListText[0].normalize('NFD').replace(/[\u0300-\u036f]/g, "");
    InputBoxes[FieldNumber*2-1].value += ListText[0]
}
//#endregion

// alphabetical and date ordering -->
//#region
  function heapSortAlphabetize() {
    //order by author name, short title, date. In that order. NB: this groups titles together and orders them by date.
    aval = [];
    for (n = 0; n <= NumOfTimelineElements; n++) {
      AlphabeticalSortPointer[n] = n;
    }
    if(n==1){
      AlphabeticalOrder[1] = 1;
      return
    }
    k = Math.floor(NumOfTimelineElements / 2) + 1;
    ir = NumOfTimelineElements;

    for (m = 1; m < 10000; m++) {
      if (1 < k) {
        k = k - 1;
        indxt = AlphabeticalSortPointer[k];
        // aval = LastName[indxt];
        aval[1] = LastNameNumber[indxt];
        aval[2] = ShortTitle[indxt];
        aval[3] = Year[indxt];
      } else {
        indxt = AlphabeticalSortPointer[ir];
        // aval = LastName[indxt];
        aval[1] = LastNameNumber[indxt];
        aval[2] = ShortTitle[indxt];
        aval[3] = Year[indxt];
        AlphabeticalSortPointer[ir] = AlphabeticalSortPointer[1];
        ir = ir - 1;
        if (ir == 1) {
          AlphabeticalSortPointer[1] = indxt;
          AlphabeticalOrder[indxt] = 1;
          break;
        }
      }
      i = k;
      j = k + k;
      while (j <= ir) {
        if (j < ir) {
          // if (LastName[AlphabeticalSortPointer[j]].localeCompare(LastName[AlphabeticalSortPointer[j + 1]]) == -1) {
          if ( LastNameNumber[AlphabeticalSortPointer[j]] < LastNameNumber[AlphabeticalSortPointer[j + 1]] ||
              (LastNameNumber[AlphabeticalSortPointer[j]] == LastNameNumber[AlphabeticalSortPointer[j + 1]] &&
              ShortTitle[AlphabeticalSortPointer[j]].localeCompare(ShortTitle[AlphabeticalSortPointer[j + 1]]) == -1) ||
              (LastNameNumber[AlphabeticalSortPointer[j]] == LastNameNumber[AlphabeticalSortPointer[j + 1]] &&
               ShortTitle[AlphabeticalSortPointer[j]] == ShortTitle[AlphabeticalSortPointer[j + 1]] &&
              Year[AlphabeticalSortPointer[j]] < Year[AlphabeticalSortPointer[j + 1]] ) ) {
            j = j + 1;
          }
        }
        // if (aval.localeCompare(LastName[AlphabeticalSortPointer[j]]) == -1) {
        if ( aval[1] < LastNameNumber[AlphabeticalSortPointer[j]] ||
            (aval[1] == LastNameNumber[AlphabeticalSortPointer[j]] &&
             aval[2].localeCompare(ShortTitle[AlphabeticalSortPointer[j]]) == -1) ||
            (aval[1] == LastNameNumber[AlphabeticalSortPointer[j]] &&
             aval[2] == ShortTitle[AlphabeticalSortPointer[j]] &&
             aval[3] < Year[AlphabeticalSortPointer[j]]) ) {
          AlphabeticalSortPointer[i] = AlphabeticalSortPointer[j];
          i = j;
          j = j + j;
        } else {
          j = ir + 1;
        }
      }
      AlphabeticalSortPointer[i] = indxt;
    }
    //mapping back to user element numbers
    for (n = 1; n <= NumOfTimelineElements; n++) {
      AlphabeticalOrder[n] = AlphabeticalSortPointer[n];
    }
    AlphabeticalOrder[0] = 0;
  }

  function heapSortByDate() {
    //sorts by publication date then by alphabetical order of author last name
    let n, k, ir, m, j, indxt;
    let aval=[];
    for (n = 1; n <= NumOfTimelineElements; n++) {
      DateSortPointer[n] = n;
    }
    if(n==1){
      ChronologicalOrder[1] = 1;
      ChronologicalOrder[0] = 0;
      return
    }
    k = Math.floor(NumOfTimelineElements / 2) + 1;
    ir = NumOfTimelineElements;

    DateSortLoop:
    while (1==1) {
      if (k > 1) {
        k = k - 1;
        indxt = DateSortPointer[k];
        aval[1] = Year[indxt];
        // aval[2] = LastNameNumber[indxt];
      } else {
        indxt = DateSortPointer[ir];
        aval[1] = Year[indxt];
        // aval[2] = LastNameNumber[indxt];
        DateSortPointer[ir] = DateSortPointer[1];
        ir = ir - 1;
        if (ir == 1) {
          DateSortPointer[1] = indxt;
          break DateSortLoop;
        }
      }
      i = k;
      j = k + k;
      while (j <= ir) {
        if (j < ir) {
          if ( Year[DateSortPointer[j]] < Year[DateSortPointer[j + 1]] // ||
              //  (Year[DateSortPointer[j]] == Year[DateSortPointer[j + 1]] &&
                // LastNameNumber[DateSortPointer[j]] < LastNameNumber[DateSortPointer[j+1]]) 
              ) {
            j = j + 1;
          }
        }
        if ( aval[1] < Year[DateSortPointer[j]] // ||
            // (aval[1] == Year[DateSortPointer[j]] &&
            //  aval[2] < LastNameNumber[DateSortPointer[j]] ) 
             ) {
          DateSortPointer[i] = DateSortPointer[j];
          i = j;
          j = j + j;
        } else {
          j = ir + 1;
        }
      }
      DateSortPointer[i] = indxt;
    }
    //mapping back to user element numbers
    for (n = 1; n <= NumOfTimelineElements; n++) {
      // ChronologicalOrder[DateSortPointer[n]] = n;
      ChronologicalOrder[n] = DateSortPointer[n];
    }
    ChronologicalOrder[0] = 0;
  }
//#endregion

// Images compression and saving -->
//#region

function OpenCompressImages(){
  ImageProcessCounter = document.getElementById("ImageProcessingWindow");
  ImageProcessCounter.style.visibility = 'visible';
  ImageProcessCounter.style.zIndex = 1100;    // make sure this is on top of the spreadsheet, if its open and displayed
  dragElement(document.getElementById("ImageProcessingWindowDrag"));
  ImageCompressionNumber = Number(localStorage.getItem("ImageCompressionNumber"));
  document.getElementById("ImageCompressionLevel").value = ImageCompressionNumber;
  document.getElementById("OpenDropDownMenu").click();
  ImageCompressionAlbumSaved = true;
  document.getElementById("PreviewImage").style.display = 'none';
  document.getElementById("PreviewImage").src = '';  
  document.getElementById("ClosePreviewImagesButton").style.display = 'none';  
  document.getElementById("PreviewImageLabeling").style.display = 'none';
}

async function CompressImagesInCSV(){
  // get .bp.csv file,
  UserAssistance( "Navigate to folder '/BookPageant/CSV files', choose a BP-formatted .csv file and Open.", ' ' , ' ',
                  "CSVFile.png",5.0);
  document.getElementById("BPCSVspinner").style.display = "block";                  
  document.getElementById("PreviewImage").style.display = 'none';
  document.getElementById("PreviewImage").src = '';
  document.getElementById("BPCSVdone").style.display = "none";
  document.getElementById("PreviewImageLabeling").style.display = 'none';  
  try{CSVFileHandle
    if(CSVFileHandle[0]==null){CSVFileHandle[0]==BookPageantHomeDirectoryHandle}    
    CSVFileHandle = await window.showOpenFilePicker({startIn: CSVFileHandle[0], types: [{descriiption: 'BP file', accept:{'text/plain':['.BP.CSV']}}]});
  }catch{
    CSVFileHandle = await window.showOpenFilePicker({startIn: BookPageantHomeDirectoryHandle, types: [{descriiption: 'BP file', accept:{'text/plain':['.BP.CSV']}}]});
  }
  localforage.setItem('CSVFileHandle', CSVFileHandle);  
  BPCSVFile = await CSVFileHandle[0].getFile();
  dataFromCSVfile = await BPCSVFile.text();
  // parse the CSVfile data: the data array is needed map images
  Papa.parse(dataFromCSVfile, {
    delimiter: ",",
    header: true,
    download: false,
    dynamicTyping: false,
    skipEmptyLines: true,
    complete: function (results) {
      errors = results.errors;
      data = results.data;
      meta = results.meta;
    }
  });

  ImageNameInCSVfile.length = 0;
  let FirstRow = 0;
  NumberOfFilesFailed = 0;
  ImageProcessingContainer.innerHTML = '';
  ImageCompressionFails = '';
  gridCellsWithMissingImages[0] = 0;
  // search the data from the CSV for cell content with an image file name
  if( BPCSVfileHasBeenLoaded ){
    FirstRow = 0;
  }else{
    FirstRow = 1;
  }
  NumImageFilesToFind = 0;
  for( let m=0; m<=meta.fields.length-1; m++){
    for (let n=FirstRow; n <= data.length-1; n++){
        let TextToCheck = data[n][meta.fields[m]].toLowerCase();
        if( TextToCheck.includes(".jpg") || 
            TextToCheck.includes(".tif") ||
            TextToCheck.includes(".png") ||
            TextToCheck.includes(".bmp") ){
              NumImageFilesToFind++
              ImageNameInCSVfile[NumImageFilesToFind] = TextToCheck.slice(TextToCheck.lastIndexOf("/")+1).trim()
            }
    }
  }
  if(NumImageFilesToFind==0){
    // no image file names found in CSV file. report null search
    document.getElementById('CustomDialog3').showModal();
  }else{
    ImageProcessingContainer.innerHTML = '<strong>'+NumImageFilesToFind.toString()+
                                        ' Image File Names Found in file: <i>' +BPCSVFile.name+'</i><br> Choose a folder to search for Images, or Cancel:'+
                                        '</strong><br>'+
                                        '<button class="CompressImagesButton" onclick="CompressCSVFileImages()"><img class="img" src="BP Appearance/Folder.png" alt="Some png"></button>&nbsp&nbsp&nbsp'+
                                        '<button class="CompressImagesButton" onclick="(function(){ImageProcessingContainer.innerHTML = null}())"><img class="img" src="BP Appearance/Cancel.png" alt="Some png"></button>'
  }
}
async function CompressCSVFileImages(){
    ImageSearchResults = ''
    if(BookPageantImagesDirectoryHandle != null){
      BookPageantImagesDirectoryHandle = await window.showDirectoryPicker({startIn: BookPageantImagesDirectoryHandle});
    }else{
      BookPageantImagesDirectoryHandle = await window.showDirectoryPicker({startIn: BookPageantHomeDirectoryHandle});
    }
    localforage.setItem('BookPageantImagesDirectoryHandle', BookPageantImagesDirectoryHandle);
    ImageDirectory = await BookPageantImagesDirectoryHandle;
    let ImagesFound = [];
    for (let m=1; m < ImageNameInCSVfile.length; m++){
      ImagesFound[m] = false;
    }
    let NumImages = 0
    for await (const entry of ImageDirectory.values()) {
      // test current entry in image directory against list of image file names in CSV file
      for (let m=1; m < ImageNameInCSVfile.length; m++){
        if( ImageNameInCSVfile[m].toLowerCase() == entry.name.toLowerCase() ){
          const file = await entry.getFile();
          ImageFileList[NumImages] = file;
          NumImages++
          ImagesFound[m] = true;
          break;
        }
      }  
    }
    ImageProcessingContainer.innerHTML = '';
    for (let m=1; m < ImageNameInCSVfile.length; m++){
      if(!ImagesFound[m]){
        ImageProcessingContainer.innerHTML += '<strong><span style="color:yellow; font-size: X-large;"> &#9888</span> Image File Not Found: </strong><i> '+ImageNameInCSVfile[m]+ '</i><br>'
        ImageSearchResults += 'Image File Not Found: '+ImageNameInCSVfile[m]+'\n';
      }
    }
    MessagesInImageProcessingContainer = ImageProcessingContainer.innerHTML
    CompressImages();
}

async function CompressAllImages() {
  document.getElementById("PreviewImage").style.display = 'none';
  document.getElementById("PreviewImage").src = '';
  document.getElementById("ClosePreviewImagesButton").style.display = 'none';
  document.getElementById("PreviewImageLabeling").style.display = 'none';
  ImageFileList.length = 0;
  if(BookPageantImagesDirectoryHandle != null){
    BookPageantImagesDirectoryHandle = await window.showDirectoryPicker({startIn: BookPageantImagesDirectoryHandle});
  }else{
    BookPageantImagesDirectoryHandle = await window.showDirectoryPicker({startIn: BookPageantHomeDirectoryHandle});
  }
  if(BookPageantImagesDirectoryHandle != null){
    localforage.setItem('BookPageantImagesDirectoryHandle', BookPageantImagesDirectoryHandle);
  }
  // BookPageantImagesDirectoryHandle == handle to directory. move through each entry and get any file handles
  NumberOfImages = 0
  let extention = '';
  for await (entry of BookPageantImagesDirectoryHandle.values()) {
    if( entry.kind == 'file'){
      extention = (entry.name.substr(entry.name.indexOf('.'))).toLowerCase();
      if( extention =='.bmp' ||
          extention =='.gif' ||
          extention =='.jpeg'||
          extention =='.jpg' ||
          extention =='.tif' ||
          extention =='.tiff'||
          extention =='.png'    ){      
      const file = await entry.getFile();
      ImageFileList[NumberOfImages] = file;      
      NumberOfImages++
      } 
    } 
  }
  StartImaageProcessingLoop = 1;
  CompressImages();
}

async function CompressSomeImages() {
  document.getElementById("PreviewImage").style.display = 'none';
  document.getElementById("PreviewImage").src = '';
  document.getElementById("ClosePreviewImagesButton").style.display = 'none';
  document.getElementById("PreviewImageLabeling").style.display = 'none';
  ImageFileList.length = 0;
  try{BookPageantImagesDirectoryHandle;
    ImageFilesChosen = await window.showOpenFilePicker({multiple: true, 
                                                        startIn: BookPageantImagesDirectoryHandle,
                                                        types: [{ description: 'Images', accept: {'image/*': ['.bmp', '.gif', '.jpeg', '.jpg', '.tif', '.tiff', '.png']}},],
                                                        excludeAcceptAllOption: true
                                                        });
  }catch{
    ImageFilesChosen = await window.showOpenFilePicker({multiple: true, 
                                                        startIn: BookPageantHomeDirectoryHandle,
                                                        types: [{ description: 'Images', accept: {'image/*': ['.bmp', '.gif', '.jpeg', '.jpg', '.tif', '.tiff', '.png']}},],                                                                        
                                                        excludeAcceptAllOption: true
                                                        });      
  }
  ImageDirectory = await ImageFilesChosen;
  if(BookPageantImagesDirectoryHandle != null){
    localforage.setItem('BookPageantImagesDirectoryHandle', BookPageantImagesDirectoryHandle);
  }
  let NumImages = 0
  for await (entry of ImageDirectory.values()) {
    if (entry.kind === "file"){
      const file = await entry.getFile();
      ImageFileList[NumImages] = file;
      NumImages++
    };
  };
  CompressImages();
}

async function CompressImages(){
  // var wmf = require('wmf');  
  // ProgressElement = document.getElementById('ImageProcessingProgress');
document.getElementById("ImageProcessingProgress").style.display = "block";    
circleProgress = new CircleProgress('.ImageProcessingProgress',{textFormat: 'percent', max:100, animation: 'none' });  
// circleProgress.max = 100;
if(StartImaageProcessingLoop==1){
  circleProgress.value = 0;
  NumberOfFilesCompressed = 0;
  NumberOfFilesFailed = 0;
  ImageCompressionFails = "";
  ImageCompressionFailsFormatted = ImageSearchResults;
  ImagesZip = new JSZip();
  NumberOfImagesToProcess = ImageFileList.length;   //compress
}else{
  circleProgress.value = 100*StartImaageProcessingLoop/NumberOfImagesToProcess;
}

let TotalNumImagesProcessed = 0;
let FileSizeOriginal = 0;
let FileSizeOriginalTotal = 0;
let FileSizeCompressed = 0;
let FileSizeCompressedTotal = 0;
document.body.style.cursor = "wait";
ImageElementLoop: 
for (let ElementNumber = StartImaageProcessingLoop; ElementNumber <= NumberOfImagesToProcess; ElementNumber++) {
  circleProgress.value = 100*(ElementNumber/NumberOfImagesToProcess);
  ImageName = ImageFileList[ElementNumber-1].name.toLowerCase();
  TotalNumImagesProcessed++  
  FullImageName = ImageName //for compression    

  ImageLoaded = new Promise( function(resolve, reject) {
    img = new Image();
    img.onerror = ()=>reject();
    img.onload = ()=>resolve(
    ImageComp = new Promise( function(resolve, reject){
    Canvas = document.createElement('canvas');
    Canvas.width = img.width;
    Canvas.height = img.height;
    ctx = Canvas.getContext('2d');
    ctx.drawImage(img, 0, 0, img.width, img.height);
    imgComp = new Image();
    imgComp.onload = () => resolve();
    imgComp.onerror = () => reject();
    imgComp.src = Canvas.toDataURL("image/jpeg", Number(ImageCompressionLevel));

    //parse DataURL and return base64 to binary
    byteString = atob(imgComp.src.split(',')[1]);
    let mimeString = imgComp.src.split(',')[0].split(':')[1].split(';')[0];
    let arrayBuffer = new ArrayBuffer(byteString.length);
    let bytes = new Uint8Array(arrayBuffer);
    for (let i = 0; i < byteString.length; i++) {bytes[i] = byteString.charCodeAt(i)};
   // the file output from the compression function is always a jpg. change the file name extension to reflect this
    ImageName = ImageName.slice(0,ImageName.lastIndexOf(".")+1).trim()+"jpg";
    //arrayBuffer contains binary jpeg data. pass it to the ZIP file generator.
    ImagesZip.file(ImageName,arrayBuffer);

    }).then( ()=>{
      NumberOfFilesCompressed++
     ImageCompressionAlbum[NumberOfFilesCompressed] = [ImageName,imgComp.src];
      // for report file
      FileSizeOriginal = (ImageFileList[ElementNumber-1].size/1000000).toFixed(3).toString();
      FileSizeOriginalTotal = FileSizeOriginalTotal + (ImageFileList[ElementNumber-1].size/1000000); 
      FileSizeCompressed = (byteString.length/1000000).toFixed(3).toString();
      FileSizeCompressedTotal = FileSizeCompressedTotal + (byteString.length/1000000);
      ImageCompressionFailsFormatted += ImageName + '  '+FileSizeOriginal+'Mb '+String.fromCharCode(0x2192)+' '+FileSizeCompressed+' Mb'+String.fromCharCode(13);
      ImageCompressionAlbumSizes[NumberOfFilesCompressed] = FileSizeOriginal+'Mb '+String.fromCharCode(0x2192)+' '+FileSizeCompressed+' Mb'+String.fromCharCode(13);
      if(ElementNumber==NumberOfImagesToProcess){
        ImageProcessingContainer.innerHTML = MessagesInImageProcessingContainer+"<strong> Done: " + ElementNumber + " Images Processed </strong>";
      }else{
        ImageProcessingContainer.innerHTML = MessagesInImageProcessingContainer+"<strong> Compressing Image: "+ElementNumber+" of "+NumberOfImagesToProcess+"</strong>";
      }
      if(ImageCompressionFails.length>0){ImageProcessingContainer.innerHTML +="<br>"+"Files Not Found: "+NumberOfFilesFailed+"<br>"}
      ImageProcessingContainer.innerHTML += ImageCompressionFails;
      }, ()=>{console.log(ElementNumber, 'Failed')})
    );
    }).then( ()=>{}, ()=>{
      //code for image loading failure
      NumberOfFilesCompressed++
      NumberOfFilesFailed++
      if(ElementNumber==NumberOfImagesToProcess){
        ImageProcessingContainer.innerHTML = MessagesInImageProcessingContainer+"<strong> Done: " + ElementNumber + " Images Processed </strong>";
        if(ImageCompressionFails.length>0){ImageProcessingContainer.innerHTML +="<br><strong>"+"Files Not Found: "+NumberOfFilesFailed+"</strong><br>"}
      }else{
        ImageProcessingContainer.innerHTML = MessagesInImageProcessingContainer+"Compressing Image: "+ElementNumber;
        if(ImageCompressionFails.length>0){ImageProcessingContainer.innerHTML +="<br>"+"Files Not Found: "+NumberOfFilesFailed+"<br>"}
      }
      ImageCompressionFails += ElementNumber+": "+ FullImageName + "<br>";
      ImageProcessingContainer.innerHTML += ImageCompressionFails;
    }
  )

  // ImageFile = await ImageFileList[ElementNumber-1].getFile()
  ImageFile = ImageFileList[ElementNumber-1];
  FullImageName = ImageFile.name
  ImageReader = new FileReader();
  ImageReader.readAsDataURL(ImageFile);
  ImageReader.onloadend = ()=>{img.src = ImageReader.result;}
  

  // img.src = FullImageName;
  await ImageLoaded // wait here until image is loaded from file and has been compressed.
  if(ImageProcessingStopped){
    StartImaageProcessingLoop = ElementNumber+1;
    break;
  };
} 

if(ImageProcessingStopped){
  if(!ImageProcessingCancelled){
      ImageProcessingContainer.innerHTML = MessagesInImageProcessingContainer+"Compressing Images <strong>Paused</strong> ";
  }else{
    ImageProcessingContainer.innerHTML = MessagesInImageProcessingContainer+"Processing Stopped<br>"+
                                         "Number of Images processed: "+StartImaageProcessingLoop.toString()
    ImageCompressionFailsFormatted = ImageSearchResults +
                                     "Processing Stopped"+String.fromCharCode(13)
                                     "Number of Images processed: "+StartImaageProcessingLoop.toString()                                           
    circleProgress.value = 0;
    ImageCompressionAlbumSizes[0] = FileSizeOriginalTotal.toFixed(3).toString()+'Mb '+String.fromCharCode(0x2192)+' '+FileSizeCompressedTotal.toFixed(3).toString()+' Mb'+String.fromCharCode(13);
    //reset image processing status
    StartImaageProcessingLoop = 1;
    ImageProcessingStopped = false;
    ImageProcessingCancelled = false;
  }
}else{
  //clean up ImageCompressionFails string for reading/printing
  if(ImageCompressionFails.length>0){
    ImageCompressionFailsFormatted += ImageCompressionFails.replace(/<br>/gi, String.fromCharCode(13));
    ImageCompressionFailsFormatted = "Number of Files Compressed: " + NumberOfFilesCompressed+String.fromCharCode(10,13)+"Files not found: "+NumberOfFilesFailed+String.fromCharCode(13)+ImageCompressionFailsFormatted;
  }else{
    ImageCompressionFailsFormatted = "Number of Files Compressed: " + NumberOfFilesCompressed+String.fromCharCode(13)+ImageCompressionFailsFormatted;
  }
  ImageCompressionAlbumSizes[0] = FileSizeOriginalTotal.toFixed(3).toString()+'Mb '+String.fromCharCode(0x2192)+' '+FileSizeCompressedTotal.toFixed(3).toString()+' Mb'+String.fromCharCode(13);
  ImageCompressionAlbum[0] = TotalNumImagesProcessed;      
  ImageCompressionAlbum.length = TotalNumImagesProcessed+1;
  ImageCompressionAlbumEmpty = false;

}
circleProgress.value = 100.0*(NumberOfFilesCompressed/NumberOfImagesToProcess);      
document.body.style.cursor = "initial";
// document.getElementById("SaveImageProcessingReportButton").style.opacity = "1.0";
document.getElementById("SaveImageProcessingReportButton").disabled = false;  
// document.getElementById("ZIPImageProcessingButton").style.opacity = "1.0";
document.getElementById("ZIPImageProcessingButton").disabled = false;  
// document.getElementById("SaveImageProcessingButton").style.opacity = "1.0";
document.getElementById("SaveImageProcessingButton").disabled = false;    
document.getElementById("PreviewCompressImagesButton").disabled = false;    

document.getElementById("ImageProcessingProgress").style.display = "none";

ImageCompressionAlbumSaved = false;
}

async function SaveImageProcessing(){
  ImageCompressionAlbumSaved = true;
  if(BookPageantImagesDirectoryHandle != null){
    BookPageantImagesDirectoryHandle = await window.showSaveFilePicker({startIn: BookPageantImagesDirectoryHandle, excludeAcceptAllOption: true, types: [{descriiption: 'Compressed Images file', accept:{'text/plain':['.compImages.bp']}}]});
  }else{
    BookPageantImagesDirectoryHandle = await window.showSaveFilePicker({startIn: BookPageantHomeDirectoryHandle, excludeAcceptAllOption: true, types: [{descriiption: 'Compressed Images file', accept:{'text/plain':['.compImages.bp']}}]});
  }
  localforage.setItem('BookPageantImagesDirectoryHandle', BookPageantImagesDirectoryHandle);    
  let writableStream = await BookPageantImagesDirectoryHandle.createWritable();
  // string length is limited to < 2^29 ~ 500,000,000 bytes. if image collection is >, then it must saved in multiple parts
  for(let m=0; m<=ImageCompressionAlbum.length-1; m++){
      await writableStream.write(JSON.stringify(ImageCompressionAlbum[m])+"!%!%!");
  }
  await writableStream.close();

}
function SetImageCompressionLevel(ImageCompressionNumber){
  ImageCompressionLevel = ImageCompressionNumber;
  document.getElementById("ImageCompressionLevel").value = ImageCompressionNumber;
  localStorage.setItem("ImageCompressionNumber", ImageCompressionNumber);
}
function SaveImageProcessingAnalysis(){
  download(ImageCompressionFailsFormatted, "BookPageantImageCompressionReport.txt","text/plain");
}
function DownloadZipOfCompressedImages(){
  document.body.style.cursor = "wait";
  ImageProcessingContainer.innerHTML += "<br><strong>"+"Building ZIP File. This will take a few minutes ..."+"<strong>";
  ImagesZip.generateAsync({type:"blob"}).then( 
    function(blob){
      if(NameOfLibrary !=""){
        download(blob, NameOfLibrary+" CompressedImages.ZIP" ,"octet/stream")
      }else{
        download(blob, "CompressedImages.ZIP" ,"octet/stream")
      }
      ImageProcessingContainer.innerHTML +="<br><strong>"+"ZIP File Built"+"</strong> and placed in the Stardard Download Folder"
      document.body.style.cursor = "initial";
    } );  
}
function PauseProcessingImageFiles(){
  // toggle pause function & change button image
  if(!ImageProcessingPaused){
    ImageProcessingPaused = true;
    document.getElementById("StopCompressImagesButtonImage").src = 'BP Appearance/GoResume.png';
    document.getElementById("StopCompressImagesButtonImage").title ='Resume Compressing Images';
    ImageProcessingStopped = true;
  }else{
    ImageProcessingPaused = false;
    document.getElementById("StopCompressImagesButtonImage").src = 'BP Appearance/Pause.png';
    document.getElementById("StopCompressImagesButtonImage").title ='Pause Compressing Images';
    ImageProcessingStopped = false;
    // continue processing images
    CompressImages()
  }

}
function StopProcessingImageFiles(){
  // reset pause function
  ImageProcessingPaused = false;
  document.getElementById("StopCompressImagesButton").src = 'BP Appearance/Pause.png';
  ImageProcessingStopped = true;
  ImageProcessingCancelled = true;
}
function PreviewCompressedImages(){
  // avoid multiple attachments of wheelzoom, otherwise it builds flashing/covering backgrounds
  document.getElementById("PreviewImage").style.display = 'block';
  document.getElementById("PreviewImage").dispatchEvent(new CustomEvent('wheelzoom.destroy'));
  wheelzoom(document.getElementById('PreviewImage'));
  document.getElementById("ClosePreviewImagesButton").style.display = 'block';
  document.getElementById("PreviewImageLabeling").style.display = 'block';
  PreviewImage.src = ImageCompressionAlbum[1][1];
  PreviewImageFileName.innerHTML = ImageCompressionAlbum[1][0];
  NumberOfPreviewImageDisplayed = 1
  CurrentPreviewImageNumber.innerHTML = '<span style="font-size: large;"><strong>'+'&nbsp&nbsp&nbsp1/'+ImageCompressionAlbum[0].toString()+'&nbsp&nbsp&nbsp</strong></span>';
  ImageCompression.innerHTML = ImageCompressionAlbumSizes[NumberOfFilesCompressed]+'&nbsp&nbsp Files total: '+ImageCompressionAlbumSizes[0];
}
function NextPreviewImage(Step){
  let ImagesInGallery = ImageCompressionAlbum[0];
  if(Step=='+'){
    NumberOfPreviewImageDisplayed++
  }else{
    NumberOfPreviewImageDisplayed--
  }
  if(NumberOfPreviewImageDisplayed > ImagesInGallery){NumberOfPreviewImageDisplayed=1}
  else if(NumberOfPreviewImageDisplayed < 1){NumberOfPreviewImageDisplayed=ImagesInGallery};
  PreviewImage.src = ImageCompressionAlbum[NumberOfPreviewImageDisplayed][1];
  PreviewImageFileName.innerHTML = ImageCompressionAlbum[NumberOfPreviewImageDisplayed][0];
  ImageCompression.innerHTML = ImageCompressionAlbumSizes[NumberOfPreviewImageDisplayed]+'&nbsp&nbsp Files total: '+ImageCompressionAlbumSizes[0];
  CurrentPreviewImageNumber.innerHTML = '<span style="font-size: large;"><strong>&nbsp&nbsp&nbsp'+NumberOfPreviewImageDisplayed.toString()+'/'+ImageCompressionAlbum[0].toString()+'&nbsp&nbsp&nbsp</strong></span>';
}
function ClosePreviewImages(){
  document.getElementById("PreviewImage").style.display = 'none';
  document.getElementById("PreviewImage").src = '';
  document.getElementById("ClosePreviewImagesButton").style.display = 'none';
  document.getElementById("PreviewImageLabeling").style.display = 'none';
}

function CloseCompressImages(){
  if(!ImageCompressionAlbumSaved){
    // give the user the chance to save compressed images.
    document.getElementById('CustomDialog2').showModal();
  }
  ImageProcessingContainer.innerHTML = '';
  ImageProcessCounter = document.getElementById("ImageProcessingWindow");
  ImageProcessCounter.style.visibility = 'hidden';
  // always turn on all the controls on the image processing window when closing it
  document.getElementById("CompressAllImagesButton").style.display = 'block';
  document.getElementById("CompressImagesButton").style.display = 'block';
  document.getElementById("ImageCompressionLevel").style.display = 'block';
  document.getElementById("StopCompressImagesButton").style.display = 'block';
  document.getElementById("CancelCompressImagesButton").style.display = 'block';
}
//#endregion

// BP file generation, save, and open -->
//#region

function MakeBookPageantFileWithImages(){
  // opens bp file make window with upload buttons and BP file save button
  document.getElementById("OpenDropDownMenu").click();  
  document.getElementById("MakeBPfile").style.visibility = 'visible';
  dragElement(document.getElementById("MakeBPfileDrag"));
  // remove display of wait/loading circles and done checks on the three buttons
  document.getElementById("BPCSVspinner").style.display = "none";
  document.getElementById("BPCSVdone").style.display = "none";  
  document.getElementById("Imagesspinner").style.display = "none";
  document.getElementById("NoImagesspinner").style.display = "none";
  document.getElementById("Imagesdone").style.display = "none";
  document.getElementById("NoImagesdone").style.display = "none";
  document.getElementById("BPspinner").style.display = "none";
  document.getElementById("BPdone").style.display = "none";
}

function CloseMakeBPfile(){
  document.getElementById("MakeBPfile").style.visibility = 'hidden';
}

async function UploadBPCSVfile(){
  // get .bp.csv file,
  UserAssistance( "Navigate to folder '/BookPageant/CSV files', choose a BP-formatted .csv file and Open.",
                  "CSVFile.png",5.0);
  document.getElementById("BPCSVspinner").style.display = "block";
  document.getElementById("BPCSVdone").style.display = "none";    
  try{CSVFileHandle
    if(CSVFileHandle[0]==null){CSVFileHandle[0]==BookPageantHomeDirectoryHandle}    
    CSVFileHandle = await window.showOpenFilePicker({startIn: CSVFileHandle[0], types: [{descriiption: 'BP file', accept:{'text/plain':['.BP.CSV']}}]});
  }catch{
    CSVFileHandle = await window.showOpenFilePicker({startIn: BookPageantHomeDirectoryHandle, types: [{descriiption: 'BP file', accept:{'text/plain':['.BP.CSV']}}]});
  }
  localforage.setItem('CSVFileHandle', CSVFileHandle);  
  BPCSVFile = await CSVFileHandle[0].getFile();
  dataFromCSVfile = await BPCSVFile.text();
  // parse the CSVfile data: the data array is needed map images
  Papa.parse(dataFromCSVfile, {
    delimiter: ",",
    header: true,
    download: false,
    dynamicTyping: false,
    skipEmptyLines: true,
    complete: function (results) {
      errors = results.errors;
      data = results.data;
      meta = results.meta;
    }
  });

  BPCSVfileHasBeenLoaded = true;
  if(BPCSVfileHasBeenLoaded && CompressedImagesfileHasBeenLoaded){
    // if bp and images have been uploaded, activate compress and save buttons
    let element = document.getElementById("MakeBP");
    element.style.opacity = '1.0';
    element.style.cursor = 'pointer';
    element.disabled = false;
  }
  document.getElementById("BPCSVspinner").style.display = "none";
  document.getElementById("BPCSVdone").style.display = "block";  
}

async function UploadCompressedImagefile(){
  // get .bp.csv file,
  UserAssistance( "Navigate to folder '/BookPageant/Images', choose a compressed image file ( .compImages.bp ) and Open.",
                  "Images.png",5.0);
  document.getElementById("Imagesspinner").style.display = "block";
  document.getElementById("Imagesdone").style.display = "none";    
  let ImagesDirectoryHandle;
  try{BookPageantImagesDirectoryHandle
    ImagesDirectoryHandle = await window.showOpenFilePicker({startIn: BookPageantImagesDirectoryHandle, excludeAcceptAllOption: true, types: [{descriiption: 'Compressed Images file', accept:{'text/plain':['.compImages.bp']}}]});
  }catch{
    ImagesDirectoryHandle= await window.showOpenFilePicker({startIn: BookPageantHomeDirectoryHandle, excludeAcceptAllOption: true, types: [{descriiption: 'Compressed Images file', accept:{'text/plain':['.compImages.bp']}}]});
  }
  if(BookPageantImagesDirectoryHandle != null){
    localforage.setItem('BookPageantImagesDirectoryHandle', BookPageantImagesDirectoryHandle);
  }
  let readableFile = await ImagesDirectoryHandle[0].getFile();
  let NewBlob = await readableFile.slice(0);
  // get size of blob and determine the number of discretizations
  let NumSteps = parseInt(NewBlob.size/500000000) + 1
  NumBytesPerStep = parseInt(NewBlob.size/NumSteps)
  Start = 0
  Stop = NumBytesPerStep
  NumberOfImages = 0
  LeftoverString = '';  
  for (let n = 1; n<=NumSteps; n++){
    Text1Blob = await NewBlob.slice(Start,Stop);
    Text1 = await Text1Blob.text();
    StringArray = Text1.split("!%!%!")
    for (let m=0; m <= StringArray.length-1; m++ ){
      if( m == 0){
        ImageCompressionAlbum[NumberOfImages] = JSON.parse(LeftoverString+StringArray[m]) //pick up last piece of previous slice of blob
        NumberOfImages++        
      }else if( m == StringArray.length-1) {
        LeftoverString = StringArray[m];              // last parsed piece of blob probably doesn't contain an entire image entry
      }else{
        ImageCompressionAlbum[NumberOfImages] = JSON.parse(StringArray[m])
        NumberOfImages++        
      }
    }
    Start = Stop;
    Stop = Stop + NumBytesPerStep;
  }

  CompressedImagesfileHasBeenLoaded = true;
  if(BPCSVfileHasBeenLoaded && CompressedImagesfileHasBeenLoaded){
    // if bp and images have been uploaded, activate compress and save buttons
    let element = document.getElementById("MakeBP");
    element.style.opacity = '1.0';
    element.style.cursor = 'pointer';
    element.disabled = false;
  }
  document.getElementById("Imagesspinner").style.display = "none";
  document.getElementById("Imagesdone").style.display = "block";  
}
  
function NoCompressedImagefile(){
  // user chooses not to include images in BP file
  CompressedImagesfileHasBeenLoaded = true;
  if(BPCSVfileHasBeenLoaded && CompressedImagesfileHasBeenLoaded){
    // if bp and images have been uploaded, activate compress and save buttons
    let element = document.getElementById("MakeBP");
    element.style.opacity = '1.0';
    element.style.cursor = 'pointer';
    element.disabled = false;
  }  
  document.getElementById("NoImagesdone").style.display = "block";
  ImageCompressionAlbum.length = 0;
  // change image that appears in make BP file button
  document.getElementById('MakeBPFileImages').src='BP Appearance/NoImages.png';
}

async function MakeBP(){
  document.getElementById("BPspinner").style.display = "block";
  document.getElementById("BPdone").style.display = "none";
  NameOfLibrary = document.getElementById("SetLibraryName").value;
  // if ImageCompressionCollectionEmpty is filled, produce BP file with compressed images attached
  compressedData = LZString.compressToUTF16(NameOfLibrary+"|"+dataFromCSVfile)+"!%!%!"+JSON.stringify(ImageCompressionCollection)
  
  if(BookPageantFileHandle != null){
    BookPageantFileHandle = await window.showSaveFilePicker({startIn: BookPageantFileHandle[0], types: [{descriiption: 'BP file', accept:{'text/plain':['.BP']}}]});
  }else{
    BookPageantFileHandle = await window.showSaveFilePicker({startIn: BookPageantHomeDirectoryHandle, types: [{descriiption: 'BP file', accept:{'text/plain':['.BP']}}]});      
  }
  localforage.setItem('BookPageantFileHandle', BookPageantFileHandle);
  NameOfLibrary = BookPageantFileHandle.name.slice(0,BookPageantFileHandle.name.lastIndexOf("."));
  let writableStream = await BookPageantFileHandle.createWritable();
  // write the data in pieces; compressed data followed by compressed images
  compressedData = LZString.compressToUTF16(NameOfLibrary+"|"+dataFromCSVfile)+"!%!%!";
  await writableStream.write(compressedData);
  for(let m=0; m<=ImageCompressionAlbum.length-1; m++){
    await writableStream.write(JSON.stringify(ImageCompressionAlbum[m])+"!%!%!");
  }
  await writableStream.close();
  document.getElementById("BPspinner").style.display = "none";
  document.getElementById("BPdone").style.display = "block";  
}
  
async function openBookPageantFile(){
  document.getElementById("OpenDropDownMenu").click();
  UserAssistance( "Navigate to folder '/BookPageant/BP files', choose a .bp file and Open.",
                  "",
                  "",
                  "BP Icon.png",5.0);  

  // let options={accepts: [ {extensions: ['.BP'], directories: false, path: '/BookPagent/BP Files/'} ]};
  // BookPageantFileHandle = await window.showOpenFilePicker(options);

  try{BookPageantFileHandle
    if(BookPageantFileHandle[0]==null){BookPageantFileHandle[0]==BookPageantHomeDirectoryHandle}
    BookPageantFileHandle = await window.showOpenFilePicker({startIn: BookPageantFileHandle[0], types: [{descriiption: 'BP file', accept:{'text/plain':['.BP']}}]});
  }catch{
    BookPageantFileHandle = await window.showOpenFilePicker({startIn: BookPageantHomeDirectoryHandle, types: [{descriiption: 'BP file', accept:{'text/plain':['.BP']}}]});      
  }
  readableFile = await BookPageantFileHandle[0].getFile();
  localforage.setItem('BookPageantFileHandle', BookPageantFileHandle);

  BPCSVfileHasBeenLoaded = false;
  BPfileHasBeenLoaded = true;
  CompressedImagesfileHasBeenLoaded = false;
  CSVfileHasBeenLoaded = false;
  let TimeLine = document.getElementById("MainTimeline");
  while (TimeLine.hasChildNodes()) {
    TimeLine.removeChild(TimeLine.firstChild);
  }  
  let ProgressElement = document.getElementById("BlinkingLoading");  
  ProgressElement.style.display = "block";
  ProgressElement.innerHTML="Loading BookPageant File"
  ProgressElement.animate([{opacity:0.50},{opacity:0.75},{opacity:1},{opacity:0.75},{opacity:0.50}],{duration:1500,iterations:Infinity})  
  let NewBlob = await readableFile.slice(0);
  // get size of blob and determine the number of discretizations
  let NumSteps = parseInt(NewBlob.size/500000000) + 1
  NumBytesPerStep = parseInt(NewBlob.size/NumSteps)
  Start = 0
  Stop = NumBytesPerStep
  NumberOfImages = 0
  LeftoverString = '';  
  for (let n = 1; n<=NumSteps; n++){
    Text1Blob = await NewBlob.slice(Start,Stop);
    Text1 = await Text1Blob.text();
    StringArray = Text1.split("!%!%!")
    for (let m=0; m <= StringArray.length-1; m++ ){
      if(StringArray[m]=='undefined'){continue};
      if( m == 0 && n == 1){
        // if this is the 1st slice of the blob, then the 1st section is the LZcompressed data
        decompressedData = LZString.decompressFromUTF16(StringArray[m]);
        NameOfLibrary = decompressedData.slice(0,decompressedData.indexOf("|"));
        element = document.getElementById("BPlogoAndLibraryName").innerHTML = '<img class="Logo" src="./BP Appearance/BP Icon.png" alt="NotFound">' + "&nbsp;"+"&nbsp;"+"&nbsp;" + NameOfLibrary.trim() + "&nbsp;"+"&nbsp;"+"&nbsp;" + '<img class="Logo" src="./Images/'+NameOfLibrary.trim()+'.JPG" alt="">' ;
        decompressedCSVData = decompressedData.slice(decompressedData.indexOf("|")+1);
      }else if( m == 0){
        NumberOfImages++
        ImageCompressionAlbum[NumberOfImages] = JSON.parse(LeftoverString+StringArray[m]) //pick up last piece of previous slice of blob
      }else if( m == StringArray.length-1) {
        LeftoverString = StringArray[m];              // last parsed piece of blob probably doesn't contain an entire image entry
      }else{
        NumberOfImages++
        ImageCompressionAlbum[NumberOfImages] = JSON.parse(StringArray[m])
      }
    }
    Start = Stop;
    Stop = Stop + NumBytesPerStep;
  }

  ParseCSVfile(decompressedCSVData);

  for (let n=0; n <= data.length-1; n++){
    ImageCompressionCollection[n] = []; //new Array('','','','','','');
  }
  DataFieldLoop:
  for( let m=0; m<=meta.fields.length-1; m++){
    switch(meta.fields[m]){
    case "BP_Image1":
      BPFieldIs = "BP_Image1"
      ImageNum = 1;
      break;
    case "BP_Image2":
      BPFieldIs = "BP_Image2"
      ImageNum = 2;      
      break;      
    case "BP_Image3":
      BPFieldIs = "BP_Image3"
      ImageNum = 3;      
      break;      
    case "BP_Image4":
      BPFieldIs = "BP_Image4"
      ImageNum = 4;      
      break;      
    case "BP_Image5":
      BPFieldIs = "BP_Image5"
      ImageNum = 5;      
      break;      
    default: 
      continue DataFieldLoop
    }  
    // search all data under this header for an image file name
    for (let n=1; n <= data.length-1; n++){
      let TextToCheck = data[n][BPFieldIs].toLowerCase();
      // let TextToCheck = data[n][BPFieldIs]
      // if( TextToCheck.includes(".jpg") || 
          // TextToCheck.includes(".tif") ||
          // TextToCheck.includes(".png") ||
          // TextToCheck.includes(".bmp") ){
          // find matching image, if any
          TextToCheck = TextToCheck.slice(TextToCheck.lastIndexOf("/")+1).trim()
          for( let p=1; p<ImageCompressionAlbum.length; p++ ){
            if( TextToCheck==ImageCompressionAlbum[p][0]){
              ImageCompressionCollection[n][ImageNum] = ImageCompressionAlbum[p][1];
              NumberOfImagesWithEntry[n] = ImageNum                               //NB: this contains the total number of images found for this data entry
            }
          }
      // }
    }
  }
  if(NumberOfImages==0){
    ImageCompressionCollectionEmpty = true;
  }else{
    ImageCompressionCollectionEmpty = false;
  }
}

//#endregion

// Composing PDF output -->
//  Basic composition settings
//#region

  function Compose(event){
    // let FieldNumber = -1; //this is only a place holder. If an open element is chosen for the composition page, it's field number will be < 0.
    document.getElementById("OpenDropDownMenu").click();    
    //bring up the list of available composition elements
    document.getElementById("AvailableCompositionElements").style.visibility = "visible";
    dragElement(document.getElementById("AvailableCompositionElementsDrag"));
    //after making sure current list is empty, (re)populate the list
    CompElements = document.getElementById("AvailableCompositionElementsContent")
    while (CompElements.hasChildNodes()) {
      CompElements.removeChild(CompElements.lastChild);
    }

    NumCompositionElements = 0
    //always list a text/symbol open element first
    NumCompositionElements = NumCompositionElements + 1;
    CompositionElementMappedToHeaderField[NumCompositionElements] = 0;
    let FieldNumber = -1; //this is only a place holder. If an open element is chosen for the composition page, it's field number will be < 0.
    CompositionElement = document.createElement('button');
    CompositionElement.setAttribute( 'onclick'  , ' CopyToCompositionBoard('+FieldNumber.toString()+' ) ');
    CompositionElement.setAttribute( 'id'  , 'CopyToCompositionBoardButton'+FieldNumber.toString());
    CompositionElement.setAttribute( 'class'  , 'CopyToCompositionBoardButton');
    CompositionElement.style.cursor = 'pointer';
    CompositionElement.innerHTML = "Open Field";
    CompositionElement.style.border = "1px dotted blue";
    CompElements.appendChild(CompositionElement)

    //list of fields
    let canvas = document.createElement("canvas");
    let context = canvas.getContext("2d");
    context.font = CurrentFont;

    NumberOfCompositionFields = BP_Headers.length;
    for (let m = 0; m < NumberOfCompositionFields; m++) {
      NumCompositionElements = NumCompositionElements + 1;
      BP_FieldNumber =  m;
      let UserNameOfDataField = '';
      //if user data is present, capture names from user data for BP_fields
      if(DataHasBeenInput){
        for (let k = 1; k < NumUserDataFields; k++) {
          // find user data name that matches BP field number
          if( BPHeadersOfUserData[k]==BP_Headers[m] ){
            UserNameOfDataField = data[0][BPHeadersOfUserData[k]];
            // trim user name if necessary
            if( UserNameOfDataField.length > 30 ){
              while(UserNameOfDataField.length > 30){
                UserNameOfDataField = UserNameOfDataField.slice(0,UserNameOfDataField.lastIndexOf(" "))
              };
            }
            UserNameOfDataField = '['+UserNameOfDataField+']';
            UserNameOfDataField = "<span style='color:blue; font-size:1.5vh'>" + UserNameOfDataField + "</span>"; 
            break;
          }
        }
      }

      CompositionElementMappedToHeaderField[NumCompositionElements] = BP_Headers[m].slice(0,19);
      CompositionElement = document.createElement('button');
      CompositionElement.setAttribute( 'onclick'  , ' CopyToCompositionBoard('+BP_FieldNumber.toString()+' ) ');
      CompositionElement.setAttribute( 'id'  , 'CopyToCompositionBoardButton'+BP_FieldNumber.toString());
      CompositionElement.setAttribute( 'class'  , 'CopyToCompositionBoardButton');
      CompositionElement.style.cursor = 'pointer';
      CompElements.appendChild(CompositionElement)

      // alert(CompositionElement.getBoundingClientRect().width)

      //make sure name isn't too long
      let Name = BP_Headers[m].trim().slice(3);
      // alert(context.measureText(Name).width)
      if( Name.length > 25 ){
        while(Name.length > 25){
          Name = Name.slice(0,Name.lastIndexOf(" "))
        }
        CompositionElement.innerHTML = "<span class='alignleft'> "+Name+" ...</span>";
      }else{
        CompositionElement.innerHTML = "<span class='alignleft'> "+Name+"</span>";
      }
      CompositionElement.style.border = "1px dotted blue";

      asdf = CompositionElement.offsetWidth
    }
    //open compositiion board
    document.getElementById("CompositionBoard").style.visibility = "visible";
    dragElement(document.getElementById("CompositionBoardDrag"));

    // defaults on opening:
    let DefalutPaperNumber = 1;  // letter 8.5/11.0
    PageMargins[1] = 1.0; //LeftMargin
    PageMargins[2] = 0.5; //RightMargin
    PageMargins[3] = 1.0; //TopMargin
    PageMargins[4] = 1.0; //BottomMargin
    NumFontFace = 0;      //normal

    //reach into jsPDF to see what fonts are avilable / have been loaded
    let PDFdoc = new jsPDF({orientation: 'p', unit: 'in', format: 'letter', });
    let fontlist = PDFdoc.getFontList();    
    NamesOfFonts = Object.keys(fontlist);
    FontFacesAvailable = Object.values(fontlist)
    let NumberOfFontsAvailable = NamesOfFonts.length
    PDFdoc = null;
    //build list of font families and faces, and dropdown button for each font family. DO NOT include the 1st 5 (fonts internal to jsPDF)
    NumberInFontList = 0;
    for (let mFont = 5; mFont < NumberOfFontsAvailable; mFont++){
      if( NamesOfFonts[mFont]=="Times-Roman"){NumFontFamily=mFont}
      for( let nFace=0; nFace < FontFacesAvailable[mFont].length; nFace++){        
        //main font menu lists only type families. leave room in font numbering for avaiable type faces.
        NumberInFontList = NumberInFontList + 1;
        if(nFace==0){
          Font = document.createElement('button');
          Font.setAttribute( 'onclick'  , 'ChooseFontFamily('+ mFont.toString() + ' ) ');
          Font.setAttribute( 'id'  , 'ChooseFontFamilyButton'+ mFont.toString() ); 
          Font.setAttribute( 'class'  , 'ChooseFontFamilyButton');
          Font.style.cursor = 'pointer';
          Font.style.left = '10px';
          Font.innerHTML = NamesOfFonts[mFont];
          Font.style.fontFamily = NamesOfFonts[mFont]
          //special cases of names that differ from OS recognized fonts (due to font handling in jsPDF)
          if(NamesOfFonts[mFont]=="Arial-Narrow"){
            Font.innerHTML = "Arial Narrow";
            Font.style.fontFamily = "Arial_Narrow";
          }else if(NamesOfFonts[mFont]=="BookmanOldStyle"){
            Font.innerHTML = "Bookman_Old_Style";
            Font.style.fontFamily = "Bookman_Old_Style";
          }else if(NamesOfFonts[mFont]=="Times-Roman"){
            Font.innerHTML = "Times_New_Roman";
            Font.style.fontFamily = "Times_New_Roman";
          }
        }
        FontNamesAndFaces[NumberInFontList] = NamesOfFonts[mFont] +" "+ FontFacesAvailable[mFont][nFace];        
      }
    }
    FontStyleSubstitute = "normal";
    SetFont(NumFontFamily, NumFontFace);

    BuildRulers();

    SetPaperSize(DefalutPaperNumber);

    // CompositionRulersHaveBeenBuilt = false;
    // // CompositionPrintRange1 = 1;
    // // CompositionPrintRange2 = 1;
    document.getElementById("FontChoose").innerHTML = CurrentFont+" "+CurrentFontFace;  
    document.getElementById("FontSize").value = CurrentFontSize;
    // // document.getElementById("ElementNumberForContent").value = CurrentElementNumberForComposition;
    // if( typeof NumberRangeText==='undefined' || NumberRangeText==='null'){NumberRangeText = "1"};
    // if( typeof PDFImageQuality==='undefined' || PDFImageQuality==='null'){PDFImageQuality = "1"};
    // PopulateCompositionElements(NumberRangeText);
    // document.getElementById("ElementNumberForContent").value = NumberRangeText;
    // document.getElementById("QualityForContent").value = PDFImageQuality;

    // SetPDFDocQuality(PDFImageQuality);
    // SetPDFDocQuality(PDFImageQuality);
    document.getElementById('HeaderFooterButton' ).setAttribute("name", "0")
    document.getElementById("SetVertSnap").value = CurrentVerticalSnap;
    document.getElementById("SetHorizSnap").value = CurrentHorizontalSnap;
    document.getElementById("PageNumberChoose").innerHTML = "None"
    NumberOfOpenFields = 0;

    document.getElementById("CompositionBoard").style.visibility = "visible";
    document.getElementById("CompositionBoardControls").style.visibility = "visible";
    document.getElementById("PaperMargins").style.visibility = 'visible';
    document.getElementById("FontChoose").style.visibility = 'visible';
    document.getElementById("PageNumberChoose").style.visibility = 'visible'; 
    
    for(let k=1; k <= NumberOfCompositonElements; k++){
      let FieldNumber = MapOfCompositionElementsToFieldNumbers[k];
      if(!document.getElementById('CompositionElement'+FieldNumber.toString())){continue}
      SettingsElementComposition(FieldNumber, 'restore');
      // shade header elements that have already been used
      if(FieldNumber >= 0){document.getElementById('CopyToCompositionBoardButton'+FieldNumber.toString()).style.opacity = "0.5"};    
    }
  }

  function CloseCompositionBoard(){
    document.getElementById("AvailableCompositionElements").style.visibility = "hidden";
    document.getElementById("CompositionBoard").style.visibility = "hidden";
    document.getElementById("CompositionBoardControls").style.visibility = "hidden";
    document.getElementById("PaperMargins").style.visibility = 'hidden';
    document.getElementById("FontChoose").style.visibility = 'hidden';
    document.getElementById("PageNumberChoose").style.visibility = 'hidden';        
    
    for(let k=1; k <= NumberOfCompositonElements; k++){
      let FieldNumber = MapOfCompositionElementsToFieldNumbers[k];
      if(!document.getElementById('CompositionElement'+FieldNumber.toString())){continue}
      SettingsElementComposition(FieldNumber, 'store&hide');
    }
    // if present, hide header composition element buttons
    if(document.getElementById("HeaderLinkContentButton")){document.getElementById("HeaderLinkContentButton").style.visibility = 'hidden'}
    if(document.getElementById("FontFaceHeaderContentButton")){document.getElementById("FontFaceHeaderContentButton").style.visibility = 'hidden'}
    if(document.getElementById("AlignHeaderContentButton")){document.getElementById("AlignHeaderContentButton").style.visibility = 'hidden'}
    if(document.getElementById("RemoveHeaderLinkContentButton")){document.getElementById("RemoveHeaderLinkContentButton").style.visibility = 'hidden'}
    if(document.getElementById("HeaderElementDrag")){document.getElementById("HeaderElementDrag").style.visibility = 'hidden'}
    // if present, hide footer composition element buttons
    if(document.getElementById("FooterLinkContentButton")){document.getElementById("FooterLinkContentButton").style.visibility = 'hidden'}
    if(document.getElementById("FontFaceFooterContentButton")){document.getElementById("FontFaceFooterContentButton").style.visibility = 'hidden'}
    if(document.getElementById("AlignFooterContentButton")){document.getElementById("AlignFooterContentButton").style.visibility = 'hidden'}
    if(document.getElementById("RemoveFooterLinkContentButton")){document.getElementById("RemoveFooterLinkContentButton").style.visibility = 'hidden'}
    if(document.getElementById("FooterElementDrag")){document.getElementById("FooterElementDrag").style.visibility = 'hidden'}

    SpecialCharactersLoaded = false;
  }

  function PaperCompositionBoard(){
    // let user choose paper size and set size/proportions for composition board and rulers
    let PaperSizeDropDownElement = document.getElementById("PaperSizeDropDown")    
    if(!ShowPaperSizes){
      PaperSizeDropDownElement.style.visiblity = "visible";
      PaperSizeDropDownElement.style.zIndex = "50";
      ShowPaperSizes = true;
      for (m=0; m<14; m++){
        PaperSize = document.createElement('button');
        PaperSize.setAttribute( 'onclick'  , ' SetPaperSize('+m.toString()+' ) ');
        PaperSize.setAttribute( 'id'  , 'SetPaperSizeButton' + m.toString()+' ) ');
        PaperSize.setAttribute( 'class'  , 'SetPaperSizeButton');
        PaperSize.style.cursor = 'pointer';
        PaperSize.style.left = '10px';
        if(PaperSizeName[m].charAt(0) != "A") {
          //Imperial, USA standard paper sizes, decimal inches
          PaperSize.innerHTML = PaperSizeName[m] +"   "+PaperSizeWidth[m]+" x "+PaperSizeLength[m]+" in";
        }else{
          //ISO metric paper sizes, given in mee converted to cm.
          PaperSize.innerHTML = PaperSizeName[m] +"   "+PaperSizeWidth[m]+" x "+PaperSizeLength[m]+" cm";
        }
        PaperSizeDropDownElement. appendChild(PaperSize);
      }
      //find paper select button and position dropdown w.r.t to its lower left corner
      let PaperButtonPlace = document.getElementById("PaperCompositionBoardButton").getBoundingClientRect();
      xTrans = PaperButtonPlace.left.toString()+'px';
      yTrans = (PaperButtonPlace.bottom).toString()+'px';
      // console.log(document.getElementById("PaperCompositionBoardButton").getBoundingClientRect())
      PaperSizeDropDownElement.style.top = yTrans;      
      PaperSizeDropDownElement.style.left = xTrans;      
    }else{
      while (PaperSizeDropDownElement.hasChildNodes() && PaperSizeDropDownElement.children.length > 0) {
        PaperSizeDropDownElement.removeChild(PaperSizeDropDownElement.lastChild);
      }
      PaperSizeDropDownElement.style.visiblity = "hidden";
      PaperSizeDropDownElement.style.zIndex = "0";
      ShowPaperSizes = false;
    }
  }

  function SetPaperSize(NewPaperNumber){
    PaperNumber = NewPaperNumber;
    PaperName = PaperSizeName[PaperNumber];
    PaperWidth = PaperSizeWidth[PaperNumber];
    PaperLength = PaperSizeLength[PaperNumber];
    //display paper name/size in compose board control header
    if(PaperName.charAt(0) != "A") {
      //Imperial, USA standard paper sizes, decimal inches
      PaperText = PaperName +"<br>   <b>"+PaperWidth+"</b> x <b>"+PaperLength+"</b> in";
      document.documentElement.style.setProperty('--ComposerPaperWidth', PaperWidth.toString());
      document.documentElement.style.setProperty('--ComposerPaperLength', PaperLength.toString());
      PaperUnits = 1;
    }else{
      //ISO metric paper sizes, given in mee converted to cm.
      PaperText = PaperName +"<br>   <b>"+(Number(PaperWidth)).toString()+"</b> x <b>"+(Number(PaperLength)).toString()+"</b> cm";
      document.documentElement.style.setProperty('--ComposerPaperWidth', (Number(PaperWidth)).toString());
      document.documentElement.style.setProperty('--ComposerPaperLength', (Number(PaperLength)).toString());
      PaperUnits = 2.54;  // cm per inch
    }
    ShowPaperSizes = true;
    PaperCompositionBoard();
    document.getElementById("PaperForCompositionBoard").innerHTML = PaperText;
    // if(!document.getElementById("rulerH")){
      // CompositionRulersHaveBeenBuilt = false;
      // BuildRulers();
    // }
    document.getElementById("rulerH").style.width = (100*(10/PaperWidth)).toString()+"%" ;
    document.getElementById("rulerV").style.height = (100*(10/PaperWidth)*(Element.clientWidth/Element.clientHeight)).toString()+"%" ; 
    SetMargins();
    BuildRulers();
      // // remove multiple horizontal rulers. PITA.
      let rulerElements = document.getElementsByClassName('rulerH');
      while(rulerElements.length > 1){
        rulerElements[0].parentNode.removeChild(rulerElements[0]);
      }    
    //update screen fonts for composition elements
    for(let m=1; m<=NumberOfCompositonElements; m++){
      FieldNumber = MapOfCompositionElementsToFieldNumbers[m];
      element = document.getElementById('CompositionElementContent'+FieldNumber.toString());
      if(element){
        fontHeightInPX = (CurrentFontSize/72.0)*PixelsPerScaleUnitHeight*PaperUnits;
        if(CurrentFontFace.includes("Oblique")){FontStyleSubstitute = "italic"} else if (CurrentFontFace.includes("italic")){FontStyleSubstitute = "italic"} else {FontStyleSubstitute = "normal"}
        if(CurrentFontFace.includes("bold") || CurrentFontFace.includes("bold")){FontWeightSubstitute = "bold"} else {FontWeightSubstitute = "normal"}
        if(CurrentFont=="Arial-Narrow"){
          FontFamilySubstitute = "Arial_Narrow";
        }else if(CurrentFont=="BookmanOldStyle"){
          FontFamilySubstitute = "Bookman_Old_Style";
        }else if(CurrentFont=="Times-Roman"){
          FontFamilySubstitute = "Times_New_Roman";
        }else{FontFamilySubstitute = CurrentFont}
        if(FieldNumber >= 0){element.style.font = FontStyleSubstitute+" "+FontWeightSubstitute+" "+fontHeightInPX.toString()+"px "+FontFamilySubstitute};
      }
    }
  }

  function PaperMargins(){
    if(!ShowMargins){
      PaperName = PaperSizeName[PaperNumber];
      PaperWidth = PaperSizeWidth[PaperNumber];
      PaperLength = PaperSizeLength[PaperNumber];
      ShowMargins = true;
      DropDown = document.getElementById("PaperMarginsDropDown");
      DropDown.style.visibility = "visible";
      DropDown.style.zIndex = "50";
      //build dropdown input only once
      if( !DropDown.hasChildNodes() ){
        for(let m=1; m<=4; m++){
          
          let margin = document.createElement('input');
          if      (m==1){margin.setAttribute("id","LeftMarginInput");   margin.setAttribute("type", "number"); margin.setAttribute("min", "0"); margin.setAttribute("max", PaperWidth.toString());  margin.setAttribute("step", "0.05");
          }else if(m==2){margin.setAttribute("id","RightMarginInput");  margin.setAttribute("type", "number"); margin.setAttribute("min", "0"); margin.setAttribute("max", PaperWidth.toString());  margin.setAttribute("step", "0.05");
          }else if(m==3){margin.setAttribute("id","TopMarginInput");    margin.setAttribute("type", "number"); margin.setAttribute("min", "0"); margin.setAttribute("max", PaperLength.toString()); margin.setAttribute("step", "0.05");
          }else if(m==4){margin.setAttribute("id","BottomMarginInput"); margin.setAttribute("type", "number"); margin.setAttribute("min", "0"); margin.setAttribute("max", PaperLength.toString()); margin.setAttribute("step", "0.05");
          }
          margin.setAttribute("class", "MarginsInput");
          margin.style.height = "15px";
          margin.style.width = "50px";
          margin.style.border = "1px solid black";
          margin.style.visibility = "inherit";
          margin.style.padding = "0 0 0 0";
          DropDown.appendChild(margin);

          if      (m==1){
            t = document.createTextNode("Left");
          }else if(m==2){
            t = document.createTextNode("Right");
          }else if(m==3){
            t = document.createTextNode("Top");
          }else if(m==4){
            t = document.createTextNode("Bottom");
          }
          DropDown.appendChild(t);
        }
        document.getElementById("LeftMarginInput").value = PageMargins[1];
        document.getElementById("RightMarginInput").value = PageMargins[2];
        document.getElementById("TopMarginInput").value = PageMargins[3]; 
        document.getElementById("BottomMarginInput").value = PageMargins[4];
      }else{
        document.getElementById("LeftMarginInput").value = PageMargins[1];
        document.getElementById("RightMarginInput").value = PageMargins[2];
        document.getElementById("TopMarginInput").value = PageMargins[3];
        document.getElementById("BottomMarginInput").value = PageMargins[4];
      }
      let PaperMarginPlace = document.getElementById("PaperMarginsButton").getBoundingClientRect();
      xTrans = PaperMarginPlace.left.toString()+'px';
      yTrans = (PaperMarginPlace.bottom).toString()+'px';
      DropDown.style.top = yTrans;      
      DropDown.style.left = xTrans;      
    }else{
      PageMargins[1] = parseFloat(document.getElementById("LeftMarginInput").value);
      PageMargins[2] = parseFloat(document.getElementById("RightMarginInput").value);
      PageMargins[3] = parseFloat(document.getElementById("TopMarginInput").value);
      PageMargins[4] = parseFloat(document.getElementById("BottomMarginInput").value);
      DropDown = document.getElementById("PaperMarginsDropDown");
      DropDown.style.visibility = "hidden";
      DropDown.style.zIndex = "0";
      ShowMargins = false;
      SetMargins();
    }
  }  

  function SetMargins(){
    PaperName = PaperSizeName[PaperNumber];
    PaperWidth = PaperSizeWidth[PaperNumber];
    PaperLength = PaperSizeLength[PaperNumber];
    LeftMargin = PageMargins[1];
    RightMargin = PageMargins[2];
    TopMargin = PageMargins[3];
    BottomMargin = PageMargins[4];
    MarginWidth = PaperWidth-LeftMargin-RightMargin;
    MarginHeight = PaperLength-TopMargin-BottomMargin;
    if(PaperName.charAt(0) != "A") {
      //Imperial, USA standard paper sizes, decimal inches
      MarginText = "L:<b>"+LeftMargin.toFixed(2)+"</b>  R:<b>"+RightMargin.toFixed(2)+"</b><br>  T:<b>"+TopMargin.toFixed(2)+"</b>  B:<b>"+BottomMargin.toFixed(2)+"</b>";
    }else{
      //ISO metric paper sizes, given in mee converted to cm.
      MarginText = "L:<b>"+LeftMargin.toFixed(2)+"</b>  R:<b>"+RightMargin.toFixed(2)+"</b><br>  T:<b>"+TopMargin.toFixed(2)+"</b>  B:<b>"+BottomMargin.toFixed(2)+"</b>";
    }
    document.getElementById("PaperMargins").innerHTML = MarginText;

    if(!document.getElementById("MarginsContainer")){
      CompositionBoard = document.getElementById("CompositionBoardContent");
      MarginsContainer = document.createElement('div');
      MarginsContainer.setAttribute( 'id', 'MarginsContainer' );
      MarginsContainer.setAttribute( 'class', 'MarginsContainer');
      MarginsContainer.style.backgroundColor = "transparent";
      MarginsContainer.style.position = "absolute";
      MarginsContainer.style.visibility = "inherit";
      MarginsContainer.style.width = "100%"
      MarginsContainer.style.zIndex = "0";
      // the height of the margins container == the vertical ruler extent: 121 10ths as determined by paper WIDTH. Element.
      // .clientWidth/.clientHeight accounts for the aspect ratio 
      MarginsContainer.style.height = (100*(10/PaperWidth)*(CompositionBoard.clientWidth/CompositionBoard.clientHeight)).toString()+"%";
      MarginsContainer.style.left = "0%";
      MarginsContainer.style.top = "0%";
      CompositionBoard.appendChild(MarginsContainer);
    }else{
      CompositionBoard = document.getElementById("CompositionBoardContent");
      document.getElementById("MarginsContainer").style.height = (100*(10/PaperWidth)*(CompositionBoard.clientWidth/CompositionBoard.clientHeight)).toString()+"%";
    }
    //remove any existing margin outlines
    MarginsContainer = document.getElementById("MarginsContainer");
    while (MarginsContainer.hasChildNodes() && MarginsContainer.children.length > 0) {
      MarginsContainer.removeChild(MarginsContainer.lastChild);
      }
    // put margins in the composition board as a composition element
    for(m=0; m<NumberOfPagesDisplayed-1; m++){
      MarginOutline = document.createElement('div');
      MarginOutline.setAttribute( 'id'  , 'Margins' );
      MarginOutline.setAttribute( 'class'  , 'Margins');
      MarginOutline.style.backgroundColor = "transparent";
      MarginOutline.style.color = "rgb(0, 0, 0)"; 
      MarginOutline.style.width = (100*(MarginWidth/PaperWidth)).toString()+"%";
      MarginOutline.style.left = (100*(LeftMargin/PaperWidth)).toString()+"%";
      // margin outlines are in the MarginContainer which is sized so that: 100% -> 10 units high. So, divinging a hieght dimension by 10 scales it properly
      MarginOutline.style.height = (100*(MarginHeight/10)).toString()+"%";
      MarginOutline.style.top = (100*(TopMargin/10. + m*PaperLength/10.)).toString()+"%";
      MarginOutline.style.border ="1px dotted var(--Gray10)";
      MarginOutline.style.position = "absolute";
      MarginOutline.style.visibility = "inherit";
      MarginOutline.style.zIndex = "0";
      MarginOutline.style.overflow = "auto";
      MarginOutline.style.resize = "both";
      MarginsContainer.appendChild(MarginOutline);
    }
  }

  function FontChoose(){
    // let user choose font and face
    let FontChooseDropDownElement = document.getElementById("FontChooseDropDown");    
    if(!ShowFonts){
      FontChooseDropDownElement.style.visiblity = "inherit";
      FontChooseDropDownElement.style.zIndex = "50";
      let FontButtonPlace = document.getElementById("FontChooseButton").getBoundingClientRect();
      console.log(FontButtonPlace)
      xTrans = FontButtonPlace.left.toString()+'px';
      yTrans = FontButtonPlace.bottom.toString()+'px';
      FontChooseDropDownElement.style.top = yTrans;
      FontChooseDropDownElement.style.left = xTrans;
      ShowFonts = true;
      //reach into jsPDF to see what fonts are avilable / have been loaded
      let PDFdoc = new jsPDF({orientation: 'p', unit: 'in', format: 'letter', });
      let fontlist = PDFdoc.getFontList();
      NamesOfFonts = Object.keys(fontlist);
      FontFacesAvailable = Object.values(fontlist)
      let NumberOfFontsAvailable = NamesOfFonts.length
      PDFdoc = null;
      //build list of font families and faces, and dropdown button for each font family. DO NOT include the 1st 5 (fonts internal to jsPDF)
      NumberInFontList = 0;
      for (mFont = 5; mFont < NumberOfFontsAvailable; mFont++){
        for( let nFace=0; nFace < FontFacesAvailable[mFont].length; nFace++){        
          //main font menu lists only type families. leave room in font numbering for avaiable type faces.
          NumberInFontList = NumberInFontList + 1;
          if(nFace==0){
            Font = document.createElement('button');
            Font.setAttribute( 'onclick'  , 'ChooseFontFamily('+ mFont.toString() + ' ) ');
            Font.setAttribute( 'id'  , 'ChooseFontFamilyButton'+ mFont.toString() ); 
            Font.setAttribute( 'class'  , 'ChooseFontFamilyButton');
            Font.style.cursor = 'pointer';
            Font.style.left = '10px';
            Font.innerHTML = NamesOfFonts[mFont];
            Font.style.fontFamily = NamesOfFonts[mFont]
            //special cases of names that differ from OS recognized fonts (due to font handling in jsPDF)
            if(NamesOfFonts[mFont]=="Arial-Narrow"){
              Font.innerHTML = "Arial Narrow";
              Font.style.fontFamily = "Arial_Narrow";
            }else if(NamesOfFonts[mFont]=="BookmanOldStyle"){
              Font.innerHTML = "Bookman_Old_Style";
              Font.style.fontFamily = "Bookman_Old_Style";
            }else if(NamesOfFonts[mFont]=="Times-Roman"){
              Font.innerHTML = "Times_New_Roman";
              Font.style.fontFamily = "Times_New_Roman";
                  }
          }
          FontNamesAndFaces[NumberInFontList] = NamesOfFonts[mFont] +" "+ FontFacesAvailable[mFont][nFace];
          FontChooseDropDownElement.appendChild(Font);
        }
      }
      //make font face list place
      FontFaceChooseDropDown = document.createElement('div');
      FontFaceChooseDropDown.setAttribute( 'id'  , 'FontFaceChooseDropDown' ); 
      FontFaceChooseDropDown.setAttribute( 'class'  , 'FontFaceChooseDropDown' );
      FontFaceChooseDropDown.style.visibility = "hidden";
      FontChooseDropDownElement.appendChild(FontFaceChooseDropDown);
    }else{
      while (FontChooseDropDownElement.hasChildNodes() && FontChooseDropDownElement.children.length > 0) {
        FontChooseDropDownElement.removeChild(FontChooseDropDownElement.lastChild);
      }
      FontChooseDropDownElement.style.visiblity = "hidden";
      FontChooseDropDownElement.style.zIndex = "0";
      ShowFonts = false;
      //update screen fonts for composition elements
      for(let m=1; m<=NumberOfCompositonElements; m++){
        FieldNumber = MapOfCompositionElementsToFieldNumbers[m];
        element = document.getElementById('CompositionElementContent'+FieldNumber.toString());
        fontHeightInPX = (CurrentFontSize/72.0)*PixelsPerScaleUnitHeight*PaperUnits;
        //fussing to match requirements for CSS font style specification.
        FontStyleSubstitute = "normal";
        if( CurrentFontFace.includes("Italic") || CurrentFontFace.includes("italic") ){FontStyleSubstitute = "italic"};
        FontWeightSubstitute = "normal";
        if( CurrentFontFace.includes("Bold") || CurrentFontFace.includes("bold") ){FontWeightSubstitute = "bold"}
        //special cases of names that differ from OS recognized fonts (due to font handling in jsPDF)
        if(CurrentFont=="Arial-Narrow"){
          FontFamilySubstitute = "Arial_Narrow";
        }else if(CurrentFont=="BookmanOldStyle"){
          FontFamilySubstitute = "Bookman_Old_Style";
        }else if(CurrentFont=="Times-Roman"){
          FontFamilySubstitute = "Times_New_Roman";
          }
        element.style.font = FontStyleSubstitute+" "+FontWeightSubstitute+" "+fontHeightInPX.toString()+"px "+FontFamilySubstitute;
      }
    }
  }

  function ChooseFontFamily(NumFontFamily){
    Font = document.getElementById("FontFaceChooseDropDown");
    while (Font.hasChildNodes() && Font.children.length > 0) {
      Font.removeChild(Font.lastChild);
    }
    //if this font family button has been clicked, build a dropdown for its available faces. clicking on the face button will set the font family and face.
    for( let nFace=0; nFace < FontFacesAvailable[NumFontFamily].length; nFace++){        
      Face = document.createElement('button');
      Face.setAttribute( 'onclick'  , ' SetFont('+ NumFontFamily.toString() + "," + nFace.toString() + ' ) ');
      Face.setAttribute( 'id'  , 'SetFontButton'+ mFont.toString() ); 
      Face.setAttribute( 'class'  , 'SetFontButton');
      Face.style.cursor = 'pointer';
      Face.innerHTML = FontFacesAvailable[NumFontFamily][nFace];
      document.getElementById('FontFaceChooseDropDown').appendChild(Face);
    }
    // position font face drop down
    let FontButtonPlace = document.getElementById("FontChooseDropDown").getBoundingClientRect();
    xTrans = FontButtonPlace.right.toString()+'px';
    yTrans = (pxTOvh(FontButtonPlace.top)+(NumFontFamily-5)*2.5).toString()+'vh';   
    document.getElementById('FontFaceChooseDropDown').style.top = yTrans
    document.getElementById('FontFaceChooseDropDown').style.left = xTrans

    document.getElementById('FontFaceChooseDropDown').style.visibility = "visible";
  }

  function SetFont(NumFontFamily, NumFontFace){
    CurrentFont = NamesOfFonts[NumFontFamily];
    CurrentFontFace = FontFacesAvailable[NumFontFamily][NumFontFace]; 
    for(nfont=1; nfont <= NumberInFontList; nfont++){
      if( CurrentFont == FontNamesAndFaces[nfont].slice(0,FontNamesAndFaces[nfont].lastIndexOf(" ")).trim() && 
          CurrentFontFace == FontNamesAndFaces[nfont].slice(FontNamesAndFaces[nfont].lastIndexOf(" ")+1).trim() ){
        NumberOfCurrentFont = nfont; 
        break
      }
    }
    if(CurrentFont=="Arial-Narrow"){
      document.getElementById("FontChoose").innerHTML = "Arial_Narrow " + CurrentFontFace;
    }else if(CurrentFont=="BookmanOldStyle"){
      document.getElementById("FontChoose").innerHTML = "Bookman_Old_Style "+CurrentFontFace;
    }else if(CurrentFont=="Times-Roman"){
      document.getElementById("FontChoose").innerHTML = "Times_New_Roman "+CurrentFontFace;
    }else{document.getElementById("FontChoose").innerHTML = FontNamesAndFaces[NumberOfCurrentFont];}
  //show chosen font on composition board header
    
    //close font drop down after user choice
    Node = document.getElementById("FontChooseDropDown");
    while (Node.hasChildNodes() && Node.children.length > 0) {
      Node.removeChild(Node.lastChild);
    }
    document.getElementById("FontChooseDropDown").style.visiblity = "hidden";
    // document.getElementById('FontFaceChooseDropDown').style.visibility = "hidden";
    ShowFonts = false;
    //update screen fonts for composition elements
    for(let m=1; m<=NumberOfCompositonElements; m++){
      FieldNumber = MapOfCompositionElementsToFieldNumbers[m];
      element = document.getElementById('CompositionElementContent'+FieldNumber.toString());
      fontHeightInPX = (CurrentFontSize/72.0)*PixelsPerScaleUnitHeight*PaperUnits;
      //fussing to match requirements for CSS font style specification.
      if(CurrentFontFace.includes("Oblique")){FontStyleSubstitute = "italic"} else if (CurrentFontFace.includes("italic")){FontStyleSubstitute = "italic"} else {FontStyleSubstitute = "normal"}
      if(CurrentFontFace.includes("Bold") || CurrentFontFace.includes("bold")){FontWeightSubstitute = "bold"} else {FontWeightSubstitute = "normal"}
      if(CurrentFont=='Arial-Narrow'){
        FontFamilySubstitute = 'Arial_Narrow';
      }else if(CurrentFont=='BookmanOldStyle'){
        FontFamilySubstitute = 'Bookman_Old_Style';
      }else if(CurrentFont=='Times-Roman'){
        FontFamilySubstitute = 'Times_New_Roman';
      }else{FontFamilySubstitute = CurrentFont}
      NewFontSpec = FontStyleSubstitute+" "+FontWeightSubstitute+" "+fontHeightInPX.toString()+"px "+FontFamilySubstitute;
      element.style.font = NewFontSpec;
    }
    Header.fontName = CurrentFont;
    Header.fontStyle = FontStyleSubstitute;
    Header.fontSize = CurrentFontSize;
    if(Header.Present){document.getElementById("HeaderContent").style.font = FontStyleSubstitute+" "+FontWeightSubstitute+" "+fontHeightInPX.toString()+"px "+FontFamilySubstitute};
    Footer.fontName = CurrentFont;
    Footer.fontStyle = FontStyleSubstitute;
    Footer.fontSize = CurrentFontSize;
    if(Footer.Present){document.getElementById("FooterContent").style.font = FontStyleSubstitute+" "+FontWeightSubstitute+" "+fontHeightInPX.toString()+"px "+FontFamilySubstitute};
  }

  function FontSize(size){
    CurrentFontSize = size;
    //update screen fonts for composition elements
    for(let m=1; m<=NumberOfCompositonElements; m++){
      FieldNumber = MapOfCompositionElementsToFieldNumbers[m];
      element = document.getElementById('CompositionElementContent'+FieldNumber.toString());
      fontHeightInPX = (CurrentFontSize/72.0)*PixelsPerScaleUnitHeight*PaperUnits;
      if(CurrentFontFace.includes("Oblique")){FontStyleSubstitute = "italic"} else if (CurrentFontFace.includes("italic")){FontStyleSubstitute = "italic"} else {FontStyleSubstitute = "normal"}
      if(CurrentFontFace.includes("Bold") || CurrentFontFace.includes("bold")){FontWeightSubstitute = "bold"} else {FontWeightSubstitute = "normal"}
      if(CurrentFont=="Arial-Narrow"){
        FontFamilySubstitute = "Arial_Narrow";
      }else if(CurrentFont=="BookmanOldStyle"){
        FontFamilySubstitute = "Bookman_Old_Style";
      }else if(CurrentFont=="Times-Roman"){
        FontFamilySubstitute = "Times_New_Roman";
      }else{FontFamilySubstitute = CurrentFont}
      element.style.font = FontStyleSubstitute+" "+FontWeightSubstitute+" "+fontHeightInPX.toString()+"px "+FontFamilySubstitute;
    }
  }
  
  function SetPageNumbering(){
    // let user choose page numbering
    let PageDropdownElement = document.getElementById("PageNumberChooseDropDown");
    if(!ShowPageNumberChoices){
      PageDropdownElement.style.visiblity = "visible";
      PageDropdownElement.style.zIndex = "50";
      // place page number dropdown under button
      let PageDropdown = document.getElementById("PageNumberingButton").getBoundingClientRect();
      xTrans = PageDropdown.left.toString()+'px';
      yTrans = (PageDropdown.bottom).toString()+'px';
      PageDropdownElement.style.top = yTrans;      
      PageDropdownElement.style.left = xTrans;      

      ShowPageNumberChoices = true;
      //build list of page numbering options. 
      NumberInFontList = 0;
      for (PageNumberingChoice = 0; PageNumberingChoice <= 8; PageNumberingChoice++){
          PageNumbering = document.createElement('button');
          PageNumbering.setAttribute( 'onclick'  , 'HeaderFooterPageNumber(' + PageNumberingChoice.toString() +' )' );
          PageNumbering.setAttribute( 'id'  , 'PageNumberChooseButton' ); 
          PageNumbering.setAttribute( 'class'  , 'PageNumberChooseButton');
          PageNumbering.style.cursor = 'pointer';
          PageNumbering.style.left = '10px';
          if(PageNumberingChoice == 0){PageNumbering.innerHTML = "None"}
          else if(PageNumberingChoice == 1){PageNumbering.innerHTML = "Top Center"}
          else if(PageNumberingChoice == 2){PageNumbering.innerHTML = "Top Left"}
          else if(PageNumberingChoice == 3){PageNumbering.innerHTML = "Top Right"}
          else if(PageNumberingChoice == 4){PageNumbering.innerHTML = "Top Even/Odd"}
          else if(PageNumberingChoice == 5){PageNumbering.innerHTML = "Bottom Center"}
          else if(PageNumberingChoice == 6){PageNumbering.innerHTML = "Bottom Left"}
          else if(PageNumberingChoice == 7){PageNumbering.innerHTML = "Bottom Right"}
          else if(PageNumberingChoice == 8){PageNumbering.innerHTML = "Bottom Even/Odd"}
          PageDropdownElement.appendChild(PageNumbering);
      }
    }else{
      PageDropdownElement.visiblity = "hidden";
      PageDropdownElement.style.zIndex = "0";
      while (PageDropdownElement.hasChildNodes() && PageDropdownElement.children.length > 0) {
        PageDropdownElement.removeChild(PageDropdownElement.lastChild);
      }
      ShowPageNumberChoices = false;
    }
  }

  function SetVertSnap(value){
    CurrentVerticalSnap = value;
  }
  function SetHorizSnap(value){
    CurrentHorizontalSnap = value;
  }
//#endregion

// header and footer management
//#region
  function SetHeaderFooter(){
    // change state of header/footer
    HFElement = document.getElementById('HeaderFooterButton');
    //get current header state
    CurrentHeaderState = parseInt(HFElement.getAttribute('name')) ;
    (CurrentHeaderState == 3)?  CurrentHeaderState = 0 : CurrentHeaderState = CurrentHeaderState + 1;
    HFElement.setAttribute( 'name'  , CurrentHeaderState.toString());
    //get rid of old image
    while (HFElement.hasChildNodes()) {
      HFElement.removeChild(HFElement.firstChild);
    }
    //replace with new
    img = document.createElement('img');
    img.src = (CurrentHeaderState==0)? 'BP Appearance/Header0.png':
              (CurrentHeaderState==1)? 'BP Appearance/Header1.png':
              (CurrentHeaderState==2)? 'BP Appearance/Header2.png' : 'BP Appearance/Header3.png' ;
    img.alt = 'Some png';
    img.width = '25';
    img.height = '30';
    img.style.border = "none"
    HFElement.appendChild(img);
    //change header/footer paramters as required
    //  Header = {Present: false, Xposition: 0, Yposition: 0, Width: 0, Content: "This is header text", TextAlign: 'center', pageNumberInHeader: false, pageNumber: 0, pageNumberPosition: "", fontName: '', fontStyle: '', fontSize: 10, LineHeightUnits: 0 }
    //  Footer = {Present: false, Xposition: 0, Yposition: 0, Width: 0, Content: "This is footer text", TextAlign: 'center', pageNumberInFooter: false, pageNumber: 0, pageNumberPosition: "", fontName: '', fontStyle: '', fontSize: 10, LineHeightUnits: 0 }
    CompositionElementContent = document.getElementById("CompositionBoardContent");
      if(CurrentHeaderState==0){
        Header.Present = false;
        Footer.Present = false;
        removeHeader() //if( document.getElementById('HeaderElement')!=null){document.getElementById('HeaderElement').style.visibility = "hidden"}
        removeFooter() //if( document.getElementById('FooterElement')!=null){document.getElementById('FooterElement').style.visibility = "hidden"}
      }
    else if(CurrentHeaderState==1){
      //header only
      Header.Present = true;
      if( document.getElementById('PageHeader')==null){makeHeader()};
      Footer.Present = false;
      removeFooter() //if( document.getElementById('FooterElement')!=null){document.getElementById('FooterElement').style.visibility = "hidden"}
    }
    else if(CurrentHeaderState==2){
      //footer only
      Header.Present = false;
      Footer.Present = true;
      removeHeader() //if( document.getElementById('HeaderElement')!=null){document.getElementById('HeaderElement').style.visibility = "hidden"}
      if( document.getElementById('FooterElement')==null){makeFooter()};
    }
    else if(CurrentHeaderState==3){ 
      // header and footer
      Header.Present = true;
      Footer.Present = true;
      if( document.getElementById('HeaderElement')==null){makeHeader()};
      if( document.getElementById('FooterElement')==null){makeFooter()};
    };
  }

  function makeHeader(){
    PaperName = PaperSizeName[PaperNumber];
    PaperWidth = PaperSizeWidth[PaperNumber];
    PaperLength = PaperSizeLength[PaperNumber];
    LeftMargin = PageMargins[1];
    RightMargin = PageMargins[2];
    TopMargin = PageMargins[3];
    BottomMargin = PageMargins[4];
    MarginWidth = PaperWidth-LeftMargin-RightMargin;
    MarginHeight = PaperLength-TopMargin-BottomMargin;

    HeaderElement = document.createElement('div');
    HeaderElement.setAttribute( 'id', 'HeaderElement' );
    HeaderElement.setAttribute( 'class', 'HeaderElement');
    HeaderElement.style.backgroundColor = "transparent";
    HeaderElement.addEventListener('mouseup', function(){ResizeHeaderFooter("HeaderElement")}); 
    HeaderElement.style.position = "absolute";
    HeaderElement.style.left = (LeftMargin*PixelsPerScaleUnitWidth).toString()+'px';
    HeaderElement.style.resize = "both";
    HeaderElement.oncontextmenu = function(e){e.preventDefault(); SettingsHeaderFooter("HeaderSettings") };
    //page numbering position indicators
    HeaderElementPageNumberLeft = document.createElement('div');
    HeaderElementPageNumberLeft.setAttribute( 'id', 'HeaderElementPageNumberLeft' );
    HeaderElementPageNumberLeft.setAttribute( 'class', 'HeaderElementPageNumberLeft');
    HeaderElementPageNumberLeft.style.backgroundColor = "transparent";
    HeaderElementPageNumberLeft.style.position = "absolute";
    HeaderElementPageNumberLeft.style.left = "0px";
    HeaderElementPageNumberLeft.style.width = "11px";
    HeaderElementPageNumberLeft.innerHTML = " "
    HeaderElementPageNumberRight = document.createElement('div');
    HeaderElementPageNumberRight.setAttribute( 'id', 'HeaderElementPageNumberRight' );
    HeaderElementPageNumberRight.setAttribute( 'class', 'HeaderElementPageNumberRight');
    HeaderElementPageNumberRight.style.backgroundColor = "transparent";
    HeaderElementPageNumberRight.style.position = "absolute";
    HeaderElementPageNumberRight.style.right = "0px";
    HeaderElementPageNumberRight.style.width = "11px";
    HeaderElementPageNumberRight.innerHTML = " "
    HeaderElementPageNumberCenter = document.createElement('div');
    HeaderElementPageNumberCenter.setAttribute( 'id', 'HeaderElementPageNumberCenter' );
    HeaderElementPageNumberCenter.setAttribute( 'class', 'HeaderElementPageNumberCenter');
    HeaderElementPageNumberCenter.style.backgroundColor = "transparent";
    HeaderElementPageNumberCenter.style.position = "absolute";
    HeaderElementPageNumberCenter.style.right = "50%";
    HeaderElementPageNumberCenter.style.width = "11px";
    HeaderElementPageNumberCenter.innerHTML = ""
    //add control to header
    HeaderElement.appendChild(HeaderElementPageNumberLeft);
    HeaderElement.appendChild(HeaderElementPageNumberRight);
    HeaderElement.appendChild(HeaderElementPageNumberCenter);

    let SnapHeightOfCompositionWindow = CurrentVerticalSnap*PaperUnits;
    let NumberOfSnapsToTopMargin = PageMargins[3]/SnapHeightOfCompositionWindow;
    let FractionOfSnap = NumberOfSnapsToTopMargin-Math.floor(NumberOfSnapsToTopMargin);
    SnapTop = SnapHeightOfCompositionWindow*Math.round(0.5/SnapHeightOfCompositionWindow) + FractionOfSnap *SnapHeightOfCompositionWindow;
    HeaderElement.style.top =  (100*(SnapTop/PaperWidth)*(CompositionBoard.clientWidth/CompositionBoard.clientHeight)).toString()+"%";
    HeaderElement.style.width = (MarginWidth*PixelsPerScaleUnitWidth).toString()+'px';
    HeaderElement.style.height = (CurrentVerticalSnap*PixelsPerScaleUnitHeight*PaperUnits).toString()+'px';

    //add controls to header: toggle (drag handle, font face, align, unlink, link)
    HeaderElementControl = document.createElement('div');
    HeaderElementControl.setAttribute( 'id'  , 'HeaderElementControl' );
    HeaderElementControl.setAttribute( 'class'  , 'HeaderElementControl');
    HeaderElementControl.setAttribute( 'name'  , 'off');
    HeaderElementControl.style.visibility = 'hidden'
    HeaderElementControl.style.zIndex = "0";
    HeaderElementControl.style.right = "0px";

    //add drag element to the header element control
    HeaderElementDrag = document.createElement('div');
    HeaderElementDrag.setAttribute( 'id'  , 'HeaderElementDrag');
    HeaderElementDrag.setAttribute( 'class'  , 'MoveImageCompElement');
    HeaderElementDrag.style.width = '20px';
    HeaderElementDrag.style.height = '20px';
    HeaderElementDrag.style.right = '130px';
    HeaderElementDrag.style.position = "absolute";
    HeaderElementDrag.style.visibility = 'hidden';
    HeaderElementDrag.style.cursor = "grabbing";
    HeaderElementDrag.style.zIndex = "0";
    //image for dragging
    img = document.createElement('img');
    img.src = 'BP Appearance/Move.png';
    img.alt = 'Some png';
    img.width = '22';
    img.height = '22';
    HeaderElementDrag.appendChild(img);
    HeaderElementControl.appendChild(HeaderElementDrag);

    //add remove button to header control
    HeaderElementRemove = document.createElement('button');
    HeaderElementRemove.setAttribute( 'id'  , 'RemoveHeaderLinkContentButton' );
    HeaderElementRemove.setAttribute( 'class'  , 'RemoveHeaderLinkContentButton');
    HeaderElementRemove.setAttribute( 'onclick'  , 'RemoveHeaderLinkContent() ');
    HeaderElementRemove.style.visibility = 'hidden';
    HeaderElementRemove.style.cursor = "pointer";
    HeaderElementRemove.style.zIndex = "0";
    img = document.createElement('img');
    img.src = 'BP Appearance/BreakLink.png';
    img.alt = 'Some png';
    img.width = '22';
    img.height = '22';
    HeaderElementRemove.appendChild(img);
    HeaderElementControl.appendChild(HeaderElementRemove);

    //add align button to header control
    HeaderElementAlign = document.createElement('button');
    HeaderElementAlign.setAttribute( 'id'  , 'AlignHeaderContentButton');
    HeaderElementAlign.setAttribute( 'class'  , 'AlignHeaderFooterContentButton');
    HeaderElementAlign.setAttribute( 'onclick'  , 'AlignHeaderFooterContent('+'"AlignHeaderContentButton"'+')');
    HeaderElementAlign.style.visibility = 'hidden';
    HeaderElementAlign.style.cursor = "pointer";
    HeaderElementAlign.style.zIndex = "0";
    HeaderElementAlign.style.width = '20px';
    HeaderElementAlign.style.height = '20px';
    //initial condition for header text is: left
    HeaderElementAlign.setAttribute( 'name'  , '1');
    img = document.createElement('img');
    img.src = 'BP Appearance/AlignLeft.png';
    img.alt = 'Some png';
    img.width = '22';
    img.height = '22';
    HeaderElementAlign.appendChild(img);
    HeaderElementControl.appendChild(HeaderElementAlign);

    //add font face button to header control
    HeaderElementFontFace = document.createElement('button');
    HeaderElementFontFace.setAttribute( 'id'  , 'FontFaceHeaderContentButton');
    HeaderElementFontFace.setAttribute( 'class'  , 'FontFaceHeaderContentButton');
    HeaderElementFontFace.setAttribute( 'onclick'  , 'FontFaceHeaderFooterContent('+'"FontFaceHeaderContentButton"'+','+'"HeaderContent"'+')');
    HeaderElementFontFace.style.visibility = 'hidden';
    HeaderElementFontFace.style.cursor = "pointer";
    HeaderElementFontFace.style.zIndex = "0";
    HeaderElementFontFace.style.width = '20px';
    HeaderElementFontFace.style.height = '20px';
    //initial condition for font face is default to that for the whole composition
    if( CurrentFontFace=="normal"){HeaderElementFontFace.setAttribute( 'name'  , '1')}
    else if(CurrentFontFace=="italic"){HeaderElementFontFace.setAttribute( 'name'  , '2')}
    else if(CurrentFontFace=="bold"){HeaderElementFontFace.setAttribute( 'name'  , '3')}
    else if(CurrentFontFace=='bolditalic'){HeaderElementFontFace.setAttribute( 'name'  , '4')}
    else{HeaderElementFontFace.setAttribute( 'name'  , '1')}
    img = document.createElement('img');
    img.src = 'BP Appearance/FontFace1.png';
    img.alt = 'Some png';
    img.width = '22';
    img.height = '22';
    HeaderElementFontFace.appendChild(img);
    HeaderElementControl.appendChild(HeaderElementFontFace);

    //add content link button to header control
    HeaderElementLink = document.createElement('button');
    HeaderElementLink.setAttribute( 'id'  , 'HeaderLinkContentButton');
    HeaderElementLink.setAttribute( 'class'  , 'HeaderLinkContentButton');
    HeaderElementLink.setAttribute( 'onclick'  , 'HeaderLinkCursor()'); //clicking changes the cursor, which in turn is detected when element list is clicked
    HeaderElementLink.style.visibility = 'hidden';
    HeaderElementLink.style.cursor = "pointer";
    HeaderElementLink.style.zIndex = "0";
    HeaderElementLink.style.width = '20px';
    HeaderElementLink.style.height = '20px';
    img = document.createElement('img');
    img.src = 'BP Appearance/Link.png';
    img.alt = 'Some png';
    img.width = '22';
    img.height = '22';
    HeaderElementLink.appendChild(img);
    HeaderElementControl.appendChild(HeaderElementLink);

    //add control to header
    HeaderElement.appendChild(HeaderElementControl);
    //add header to composition board
    document.getElementById("CompositionBoardContent").appendChild(HeaderElement)
    // trigger header dragging
    dragElement(document.getElementById('HeaderElementDrag'));
    // update header object
    // Header = {Present: false, Xposition: 0, Yposition: 0, Width: 0, Content: "This is header text", TextAlign: 'center', pageNumberInHeader: false, pageNumber: 0, pageNumberPosition: "", fontName: '', fontStyle: '', fontSize: 10, LineHeightUnits: 0 }
    Header.Present = true;
    Header.Xposition = LeftMargin;
    Header.Yposition = SnapTop;
    Header.Width = MarginWidth;
    Header.TextAlign = "left";
    Header.fontName = CurrentFont;
    Header.fontStyle = CurrentFontFace;
    Header.fontSize = CurrentFontSize;

    //set header page numbering 
    SetHeaderFooterPageNumberSymbol()
}

  function makeFooter(){
    PaperName = PaperSizeName[PaperNumber];
    PaperWidth = PaperSizeWidth[PaperNumber];
    PaperLength = PaperSizeLength[PaperNumber];
    LeftMargin = PageMargins[1];
    RightMargin = PageMargins[2];
    TopMargin = PageMargins[3];
    BottomMargin = PageMargins[4];
    MarginWidth = PaperWidth-LeftMargin-RightMargin;
    MarginHeight = PaperLength-TopMargin-BottomMargin;

    FooterElement = document.createElement('div');
    FooterElement.setAttribute( 'id', 'FooterElement' );
    FooterElement.setAttribute( 'class', 'FooterElement');
    FooterElement.style.backgroundColor = "transparent";
    FooterElement.addEventListener('mouseup', function(){ResizeHeaderFooter("FooterElement")}); 
    FooterElement.style.position = "absolute";
    FooterElement.style.left = (LeftMargin*PixelsPerScaleUnitWidth).toString()+'px';
    FooterElement.style.resize = "both";
    // FooterElement.style.zIndex = "0";
    FooterElement.oncontextmenu = function(e){e.preventDefault(); SettingsHeaderFooter("FooterSettings") };
    //page numbering position indicators
    FooterElementPageNumberLeft = document.createElement('div');
    FooterElementPageNumberLeft.setAttribute( 'id', 'FooterElementPageNumberLeft' );
    FooterElementPageNumberLeft.setAttribute( 'class', 'FooterElementPageNumberLeft');
    FooterElementPageNumberLeft.style.backgroundColor = "transparent";
    FooterElementPageNumberLeft.style.position = "absolute";
    FooterElementPageNumberLeft.style.left = "0px";
    FooterElementPageNumberLeft.style.width = "11px";
    FooterElementPageNumberLeft.innerHTML = " "
    FooterElementPageNumberRight = document.createElement('div');
    FooterElementPageNumberRight.setAttribute( 'id', 'FooterElementPageNumberRight' );
    FooterElementPageNumberRight.setAttribute( 'class', 'FooterElementPageNumberRight');
    FooterElementPageNumberRight.style.backgroundColor = "transparent";
    FooterElementPageNumberRight.style.position = "absolute";
    FooterElementPageNumberRight.style.right = "0px";
    FooterElementPageNumberRight.style.width = "11px";
    FooterElementPageNumberRight.innerHTML = " "
    FooterElementPageNumberCenter = document.createElement('div');
    FooterElementPageNumberCenter.setAttribute( 'id', 'FooterElementPageNumberCenter' );
    FooterElementPageNumberCenter.setAttribute( 'class', 'FooterElementPageNumberCenter');
    FooterElementPageNumberCenter.style.backgroundColor = "transparent";
    FooterElementPageNumberCenter.style.position = "absolute";
    FooterElementPageNumberCenter.style.right = "50%";
    FooterElementPageNumberCenter.style.width = "11px";
    FooterElementPageNumberCenter.innerHTML = " "
    //add control to header
    FooterElement.appendChild(FooterElementPageNumberLeft);
    FooterElement.appendChild(FooterElementPageNumberRight);
    FooterElement.appendChild(FooterElementPageNumberCenter);

    let SnapHeightOfCompositionWindow = CurrentVerticalSnap*PaperUnits;
    let NumberOfSnapsToTopMargin = PageMargins[3]/SnapHeightOfCompositionWindow;
    let FractionOfSnap = NumberOfSnapsToTopMargin-Math.floor(NumberOfSnapsToTopMargin);
    SnapBottom = SnapHeightOfCompositionWindow*Math.round((PaperLength-0.5-SnapHeightOfCompositionWindow  )/SnapHeightOfCompositionWindow) //+ FractionOfSnap *SnapHeightOfCompositionWindow;
    FooterElement.style.top =  (100*(SnapBottom/PaperWidth)*(CompositionBoard.clientWidth/CompositionBoard.clientHeight)).toString()+"%";
    FooterElement.style.width = (MarginWidth*PixelsPerScaleUnitWidth).toString()+'px';
    FooterElement.style.height = (CurrentVerticalSnap*PixelsPerScaleUnitHeight*PaperUnits).toString()+'px';

    //add controls to footer: toggle (drag handle, font face, align, unlink, link)
    FooterElementControl = document.createElement('div');
    FooterElementControl.setAttribute( 'id'  , 'FooterElementControl' );
    FooterElementControl.setAttribute( 'class'  , 'FooterElementControl');
    FooterElementControl.setAttribute( 'name'  , 'off');
    FooterElementControl.style.visibility = 'hidden'
    FooterElementControl.style.zIndex = "0";
    FooterElementControl.style.right = "0px";

    //add drag element to the footer element control
    FooterElementDrag = document.createElement('div');
    FooterElementDrag.setAttribute( 'id'  , 'FooterElementDrag');
    FooterElementDrag.setAttribute( 'class'  , 'MoveImageCompElement');
    FooterElementDrag.style.width = '20px';
    FooterElementDrag.style.height = '20px';
    FooterElementDrag.style.right = '130px';
    FooterElementDrag.style.position = "absolute";
    FooterElementDrag.style.visibility = 'hidden';
    FooterElementDrag.style.cursor = "grabbing";
    FooterElementDrag.style.zIndex = "0";
    //image for dragging
    img = document.createElement('img');
    img.src = 'BP Appearance/Move.png';
    img.alt = 'Some png';
    img.width = '22';
    img.height = '22';
    FooterElementDrag.appendChild(img);
    FooterElementControl.appendChild(FooterElementDrag);

    //add remove button to footer control
    FooterElementRemove = document.createElement('button');
    FooterElementRemove.setAttribute( 'id'  , 'RemoveFooterLinkContentButton' );
    FooterElementRemove.setAttribute( 'class'  , 'RemoveFooterLinkContentButton');
    FooterElementRemove.setAttribute( 'onclick'  , 'RemoveFooterLinkContent() ');
    FooterElementRemove.style.visibility = 'hidden';
    FooterElementRemove.style.cursor = "pointer";
    FooterElementRemove.style.zIndex = "0";
    img = document.createElement('img');
    img.src = 'BP Appearance/BreakLink.png';
    img.alt = 'Some png';
    img.width = '22';
    img.height = '22';
    FooterElementRemove.appendChild(img);
    FooterElementControl.appendChild(FooterElementRemove);

    //add align button to footer control
    FooterElementAlign = document.createElement('button');
    FooterElementAlign.setAttribute( 'id'  , 'AlignFooterContentButton');
    FooterElementAlign.setAttribute( 'class'  , 'AlignHeaderFooterContentButton');
    FooterElementAlign.setAttribute( 'onclick'  , 'AlignHeaderFooterContent('+'"AlignFooterContentButton"'+')');
    FooterElementAlign.style.visibility = 'hidden';
    FooterElementAlign.style.cursor = "pointer";
    FooterElementAlign.style.zIndex = "0";
    FooterElementAlign.style.width = '20px';
    FooterElementAlign.style.height = '20px';

    //initial condition for footer is: left
    FooterElementAlign.setAttribute( 'name'  , '1');
    img = document.createElement('img');
    img.src = 'BP Appearance/AlignLeft.png';
    img.alt = 'Some png';
    img.width = '22';
    img.height = '22';
    FooterElementAlign.appendChild(img);
    FooterElementControl.appendChild(FooterElementAlign);

    //add font face button to footer control
    FooterElementFontFace = document.createElement('button');
    FooterElementFontFace.setAttribute( 'id'  , 'FontFaceFooterContentButton');
    FooterElementFontFace.setAttribute( 'class'  , 'FontFaceFooterContentButton');
    FooterElementFontFace.setAttribute( 'onclick'  , 'FontFaceHeaderFooterContent('+'"FontFaceFooterContentButton"'+','+'"FooterContent"'+')');
    FooterElementFontFace.style.visibility = 'hidden';
    FooterElementFontFace.style.cursor = "pointer";
    FooterElementFontFace.style.zIndex = "0";
    FooterElementFontFace.style.width = '20px';
    FooterElementFontFace.style.height = '20px';

    //initial condition for font face is default to that for the whole composition
    if( CurrentFontFace=="normal"){FooterElementFontFace.setAttribute( 'name'  , '1')}
    else if(CurrentFontFace=="italic"){FooterElementFontFace.setAttribute( 'name'  , '2')}
    else if(CurrentFontFace=="bold"){FooterElementFontFace.setAttribute( 'name'  , '3')}
    else if(CurrentFontFace=='bolditalic'){FooterElementFontFace.setAttribute( 'name'  , '4')}
    else{FooterElementFontFace.setAttribute( 'name'  , '1')}
    img = document.createElement('img');
    img.src = 'BP Appearance/FontFace1.png';
    img.alt = 'Some png';
    img.width = '22';
    img.height = '22';
    FooterElementFontFace.appendChild(img);
    FooterElementControl.appendChild(FooterElementFontFace);

    //add content link button to footer control
    FooterElementLink = document.createElement('button');
    FooterElementLink.setAttribute( 'id'  , 'FooterLinkContentButton');
    FooterElementLink.setAttribute( 'class'  , 'FooterLinkContentButton');
    FooterElementLink.setAttribute( 'onclick'  ,  'FooterLinkCursor()');  //clicking changes the cursor, which in turn is detected when element list is clicked
    FooterElementLink.style.visibility = 'hidden';
    FooterElementLink.style.cursor = "pointer";
    FooterElementLink.style.zIndex = "0";
    FooterElementLink.style.width = '20px';
    FooterElementLink.style.height = '20px';
    img = document.createElement('img');
    img.src = 'BP Appearance/Link.png';
    img.alt = 'Some png';
    img.width = '22';
    img.height = '22';
    FooterElementLink.appendChild(img);
    FooterElementControl.appendChild(FooterElementLink);

    //add control to footer
    FooterElement.appendChild(FooterElementControl);
    //add footer to composition board
    document.getElementById("CompositionBoardContent").appendChild(FooterElement)
    // trigger footer dragging
    dragElement(document.getElementById('FooterElementDrag'));
    // update footer object
    // Footer = {Present: false, Xposition: 0, Yposition: 0, Width: 0, Content: "This is footer text", TextAlign: 'center', pageNumberInFooter: false, pageNumber: 0, pageNumberPosition: "", fontName: '', fontStyle: '', fontSize: 10, LineHeightUnits: 0 }
    Footer.Present = true;
    Footer.Xposition = LeftMargin;
    Footer.Yposition = SnapBottom;
    Footer.Width = MarginWidth;
    Footer.TextAlign = "left";
    Footer.fontName = CurrentFont;
    Footer.fontStyle = CurrentFontFace;
    Footer.fontSize = CurrentFontSize;
    //set header page numbering 
    SetHeaderFooterPageNumberSymbol()
  }

  function HeaderLinkCursor(){
    document.body.style.cursor = 'url("./BP Appearance/CursorLink.png"), auto';
    HeaderOrFooterLinking = "Header";
    // change default cursor 'pointer' to CursorLink for all copy header elements
    let Headers = document.getElementsByClassName('CopyToCompositionBoardButton');
    for(i = 0; i < Headers.length; i++) {
      Headers[i].style.cursor = 'url("./BP Appearance/CursorLink.png"), auto';
    }
  }
  function FooterLinkCursor(){
    document.body.style.cursor = 'url("./BP Appearance/CursorLink.png"), auto';
    HeaderOrFooterLinking = "Footer";
    // change default cursor 'pointer' to CursorLink for all copy header elements
    let Headers = document.getElementsByClassName('CopyToCompositionBoardButton');
    for(i = 0; i < Headers.length; i++) {
      Headers[i].style.cursor = 'url("./BP Appearance/CursorLink.png"), auto';
    }    
  }

  function HeaderLinkContent(FieldNumber){
    // content of element
    document.body.style.cursor = "default";
    if(FieldNumber>0){
      //header linked to data 
      HeaderElementContent = document.createElement('div');
      HeaderElementContent.setAttribute( 'id'  , 'HeaderContent');
      HeaderElementContent.setAttribute( 'class'  , 'HeaderContent');
      HeaderElementContent.innerHTML = BP_Headers[FieldNumber].slice(3);
      HeaderElementContent.style.opacity = "0.5";
      HeaderElementContent.style.borderStyle = "groove";
      HeaderElementContent.style.width = "100%";
      HeaderElementContent.style.height = "100%";
      HeaderElementContent.style.position = "absolute";
      HeaderElementContent.style.top = '0%';
      document.getElementById("HeaderElement").appendChild(HeaderElementContent);
    }else{
      //text input box into header
      HeaderElementContent = document.createElement('input');
      HeaderElementContent.setAttribute( 'id'  , 'HeaderContent' );
      HeaderElementContent.setAttribute( 'type', 'text');
      HeaderElementContent.innerHTML = 'Open';
      HeaderElementContent.style.opacity = "0.5";
      HeaderElementContent.style.width = "99.5%";
      HeaderElementContent.style.padding = 0, 0, 0, 0;
      HeaderElementContent.style.margins = 0, 0, 0, 0;
      HeaderElementContent.style.borderStyle = "groove";
      HeaderElementContent.style.backgroundColor = 'transparent';
      HeaderElementContent.style.position = "absolute";
      HeaderElementContent.style.top = '0%';      
      document.getElementById("HeaderElement").appendChild(HeaderElementContent);
    }
    fontHeightInPX = (CurrentFontSize/72.0)*PixelsPerScaleUnitHeight*PaperUnits;
    if(CurrentFontFace.includes("Oblique")){FontStyleSubstitute = "italic"} else if (CurrentFontFace.includes("italic")){FontStyleSubstitute = "italic"} else {FontStyleSubstitute = "normal"}
    if(CurrentFontFace.includes("Bold") || CurrentFontFace.includes("bold")){FontWeightSubstitute = "bold"} else {FontWeightSubstitute = "normal"}
    //special cases of names that differ from OS recognized fonts (due to font handling in jsPDF)
    if(CurrentFont=="Arial-Narrow"){
      FontFamilySubstitute = "Arial_Narrow";
    }else if(CurrentFont=="BookmanOldStyle"){
      FontFamilySubstitute = "Bookman_Old_Style";
    }else if(CurrentFont=="Times-Roman"){
      FontFamilySubstitute = "Times_New_Roman";
    }else{FontFamilySubstitute = CurrentFont}
    HeaderElementContent.style.font = FontStyleSubstitute+" "+FontWeightSubstitute+" "+fontHeightInPX.toString()+"px "+FontFamilySubstitute;
    checkfontspec = CompositionElementContent.style.font;
    HeaderElementContent.style.columnCount = "1";

    // update header object
    // Header = {Present: false, Xposition: 0, Yposition: 0, Width: 0, Content: "This is header text", TextAlign: 'center', pageNumberInHeader: false, pageNumber: 0, pageNumberPosition: "", fontName: '', fontStyle: '', fontSize: 10, LineHeightUnits: 0 }
    Header.ContentLink = FieldNumber;
    Header.TextAlign = "center";
    Header.fontName = CurrentFont;
    Header.fontStyle = FontStyleSubstitute;
    Header.fontSize = CurrentFontSize;
  }
  function FooterLinkContent(FieldNumber){
    // content of element
    document.body.style.cursor = "default";
    if(FieldNumber>0){
      //footer linked to data 
      FooterElementContent = document.createElement('div');
      FooterElementContent.setAttribute( 'id'  , 'FooterContent');
      FooterElementContent.setAttribute( 'class'  , 'FooterContent');
      FooterElementContent.innerHTML = BP_Headers[FieldNumber].slice(3);
      FooterElementContent.style.opacity = "0.5";  
      FooterElementContent.style.borderStyle = "groove";
      FooterElementContent.style.width = "100%";
      FooterElementContent.style.height = "100%";
      FooterElementContent.style.position = "absolute";
      FooterElementContent.style.top = '0%';          
      document.getElementById("FooterElement").appendChild(FooterElementContent);      
    }else{
      //text input box into footer
      FooterElementContent = document.createElement('input');
      FooterElementContent.setAttribute( 'id'  , 'FooterContent' );
      FooterElementContent.setAttribute( 'type', 'text');
      FooterElementContent.innerHTML = 'Open';
      FooterElementContent.style.opacity = "0.5";
      FooterElementContent.style.width = "99.5%";
      FooterElementContent.style.padding = 0, 0, 0, 0;
      FooterElementContent.style.margins = 0, 0, 0, 0;
      FooterElementContent.style.borderStyle = "groove";
      FooterElementContent.style.backgroundColor = 'transparent';
      FooterElementContent.style.position = "absolute";
      FooterElementContent.style.top = '0%';            
      document.getElementById("FooterElement").appendChild(FooterElementContent);
    }
    fontHeightInPX = (CurrentFontSize/72.0)*PixelsPerScaleUnitHeight*PaperUnits;
    if(CurrentFontFace.includes("Oblique")){FontStyleSubstitute = "italic"} else if (CurrentFontFace.includes("italic")){FontStyleSubstitute = "italic"} else {FontStyleSubstitute = "normal"}
    if(CurrentFontFace.includes("Bold") || CurrentFontFace.includes("bold")){FontWeightSubstitute = "bold"} else {FontWeightSubstitute = "normal"}
    if(CurrentFont=="Arial-Narrow"){
      FontFamilySubstitute = "Arial_Narrow";
    }else if(CurrentFont=="BookmanOldStyle"){
      FontFamilySubstitute = "Bookman_Old_Style";
    }else if(CurrentFont=="Times-Roman"){
      FontFamilySubstitute = "Times_New_Roman";
    }else{FontFamilySubstitute = CurrentFont}
    FooterElementContent.style.font = FontStyleSubstitute+" "+FontWeightSubstitute+" "+fontHeightInPX.toString()+"px "+FontFamilySubstitute;
    checkfontspec = CompositionElementContent.style.font;
    FooterElementContent.style.columnCount = "1";

    // update footer object
    // Footer = {Present: false, Xposition: 0, Yposition: 0, Width: 0, Content: "This is footer text", TextAlign: 'center', pageNumberInFooter: false, pageNumber: 0, pageNumberPosition: "", fontName: '', fontStyle: '', fontSize: 10, LineHeightUnits: 0 }
    Footer.ContentLink = FieldNumber;
    Footer.TextAlign = "left";
    Footer.fontName = CurrentFont;
    Footer.fontStyle = FontStyleSubstitute;
    Footer.fontSize = CurrentFontSize;
  }

  function RemoveHeaderLinkContent(){
    if(document.getElementById("HeaderContent")){document.getElementById("HeaderContent").remove()}
    Header.ContentLink = 0;
  }
  function RemoveFooterLinkContent(){
    if(document.getElementById("FooterContent")){document.getElementById("FooterContent").remove()}
    Footer.ContentLink = 0;
  }

  function removeHeader(){
    if(document.getElementById('HeaderElement')!=null){
       document.getElementById('HeaderElement').removeEventListener('mouseup', function(){ResizeHeaderFooter("HeaderElement")});
       while (document.getElementById('HeaderElement').firstChild) {
        document.getElementById('HeaderElement').firstChild.remove();
       }
       document.getElementById('HeaderElement').remove()
       Header.Present = false;
    }
  }
  function removeFooter(){
    if( document.getElementById('FooterElement')!=null){
        document.getElementById('FooterElement').removeEventListener('mouseup', function(){ResizeHeaderFooter("FooterElement")});             
        while (document.getElementById('FooterElement').firstChild) {
          document.getElementById('FooterElement').firstChild.remove();
         }
          document.getElementById('FooterElement' ).remove()
        Footer.Present = false;
      };
  }
  
  function HeaderFooterPageNumber(ParameterValue){
    // Header = {Xposition: 0, Yposition: 0, Width: 0, Content: "This is header text", TextAlign: 'center', pageNumberInHeader: false, pageNumber: 0, pageNumberPosition: "", fontName: '', fontStyle: '', fontSize: 10, LineHeightUnits: 0 }
    // Footer = {Xposition: 0, Yposition: 0, Width: 0, Content: "This is footer text", TextAlign: 'center', pageNumberInFooter: false, pageNumber: 0, pageNumberPosition: "", fontName: '', fontStyle: '', fontSize: 10, LineHeightUnits: 0 }
    if     (ParameterValue==0){Header.pageNumberInHeader=false;                                         Footer.pageNumberInFooter=false; document.getElementById("PageNumberChoose").innerHTML = "None"}
    else if(ParameterValue==1){Header.pageNumberInHeader=true;  Header.pageNumberPosition="TopCenter";  Footer.pageNumberInFooter=false; Footer.pageNumberPosition="";             document.getElementById("PageNumberChoose").innerHTML = "Top Center"}
    else if(ParameterValue==2){Header.pageNumberInHeader=true;  Header.pageNumberPosition="TopLeft";    Footer.pageNumberInFooter=false; Footer.pageNumberPosition="";             document.getElementById("PageNumberChoose").innerHTML = "Top Left"}
    else if(ParameterValue==3){Header.pageNumberInHeader=true;  Header.pageNumberPosition="TopRight";   Footer.pageNumberInFooter=false; Footer.pageNumberPosition="";             document.getElementById("PageNumberChoose").innerHTML = "Top Right"}
    else if(ParameterValue==4){Header.pageNumberInHeader=true;  Header.pageNumberPosition="TopEvenOdd"; Footer.pageNumberInFooter=false; Footer.pageNumberPosition="";             document.getElementById("PageNumberChoose").innerHTML = "Top Even Odd"}
    else if(ParameterValue==5){Header.pageNumberInHeader=false; Header.pageNumberPosition="";           Footer.pageNumberInFooter=true;  Footer.pageNumberPosition="BottomCenter"; document.getElementById("PageNumberChoose").innerHTML = "Bottom Center"}
    else if(ParameterValue==6){Header.pageNumberInHeader=false; Header.pageNumberPosition="";           Footer.pageNumberInFooter=true;  Footer.pageNumberPosition="BottomLeft";   document.getElementById("PageNumberChoose").innerHTML = "Bottom Left"}
    else if(ParameterValue==7){Header.pageNumberInHeader=false; Header.pageNumberPosition="";           Footer.pageNumberInFooter=true;  Footer.pageNumberPosition="BottomRight";  document.getElementById("PageNumberChoose").innerHTML = "Bottom Right"}
    else if(ParameterValue==8){Header.pageNumberInHeader=false; Header.pageNumberPosition="";           Footer.pageNumberInFooter=true;  Footer.pageNumberPosition="BottomEvenOdd";document.getElementById("PageNumberChoose").innerHTML = "Bottom Even Odd"}
    PageNumberLayoutNumber = ParameterValue;
    SetHeaderFooterPageNumberSymbol()
    SetPageNumbering()
  }

  function SetHeaderFooterPageNumberSymbol(){
    if(Header.Present){
      document.getElementById("HeaderElementPageNumberLeft").innerHTML = " ";
      document.getElementById("HeaderElementPageNumberRight").innerHTML = " ";
      document.getElementById("HeaderElementPageNumberCenter").innerHTML = " ";
      if(Header.pageNumberInHeader){
             if(Header.pageNumberPosition == "TopCenter"){document.getElementById("HeaderElementPageNumberCenter").innerHTML = "#";}
        else if(Header.pageNumberPosition == "TopLeft"  ){document.getElementById("HeaderElementPageNumberLeft").innerHTML = "#";}
        else if(Header.pageNumberPosition == "TopRight" ){document.getElementById("HeaderElementPageNumberRight").innerHTML = "#";}
        else{document.getElementById("HeaderElementPageNumberLeft").innerHTML = "#"; document.getElementById("HeaderElementPageNumberRight").innerHTML = "#";}
      }
    }

    if(Footer.Present){
      document.getElementById("FooterElementPageNumberLeft").innerHTML = " ";
      document.getElementById("FooterElementPageNumberRight").innerHTML = " ";
      document.getElementById("FooterElementPageNumberCenter").innerHTML = " ";
      if(Footer.pageNumberInFooter){
             if(Footer.pageNumberPosition == "BottomCenter"){document.getElementById("FooterElementPageNumberCenter").innerHTML = "#"}
        else if(Footer.pageNumberPosition == "BottomLeft"  ){document.getElementById("FooterElementPageNumberLeft").innerHTML = "#"}
        else if(Footer.pageNumberPosition == "BottomRight" ){document.getElementById("FooterElementPageNumberRight").innerHTML = "#"}
        else {document.getElementById("FooterElementPageNumberLeft").innerHTML = "#"; document.getElementById("FooterElementPageNumberRight").innerHTML = "#"}
      }
    }
  }

  function SettingsHeaderFooter(HeaderFooter){
    // toggle settings of header control on/off
    if(HeaderFooter.includes("Header")){
      Element = document.getElementById('HeaderElementControl');
      let ControlsState = Element.getAttribute('name');
      if(ControlsState=='off'){
        // turn setting for this compositiion element on
        Element.setAttribute('name', 'on')
        Element.style.visibility = 'visible'
        Element.style.zIndex = "50";
        document.getElementById('HeaderLinkContentButton').style.visibility   = 'visible';
        document.getElementById('HeaderLinkContentButton').style.zIndex = "40";
        document.getElementById('AlignHeaderContentButton').style.visibility   = 'visible';
        document.getElementById('AlignHeaderContentButton').style.zIndex = "40";
        document.getElementById('RemoveHeaderLinkContentButton').style.visibility  = 'visible';
        document.getElementById('RemoveHeaderLinkContentButton').style.zIndex = "40";
        document.getElementById('FontFaceHeaderContentButton').style.visibility= 'visible';
        document.getElementById('FontFaceHeaderContentButton').style.zIndex = "40";
        document.getElementById('HeaderElement'+'Drag').style.visibility   = 'visible';
        document.getElementById('HeaderElement'+'Drag').style.zIndex = "40";
      }else{
        // turn setting for this compositiion element off
        Element.setAttribute('name', 'off')
        Element.style.visibility = 'hdden'
        Element.style.zIndex = "0";
        document.getElementById('HeaderLinkContentButton').style.visibility   = 'hidden';
        document.getElementById('HeaderLinkContentButton').style.zIndex = "0";
        document.getElementById('AlignHeaderContentButton').style.visibility   = 'hidden';
        document.getElementById('AlignHeaderContentButton').style.zIndex = "0";
        document.getElementById('RemoveHeaderLinkContentButton').style.visibility  = 'hidden';
        document.getElementById('RemoveHeaderLinkContentButton').style.zIndex = "0";
        document.getElementById('FontFaceHeaderContentButton').style.visibility= 'hidden';
        document.getElementById('FontFaceHeaderContentButton').style.zIndex = "0";
        document.getElementById('HeaderElement'+'Drag').style.visibility   = 'hidden';
        document.getElementById('HeaderElement'+'Drag').style.zIndex = "0";
      }
    }else{
      Element = document.getElementById('FooterElementControl');
      let ControlsState = Element.getAttribute('name');
      if(ControlsState=='off'){
        // turn setting for this compositiion element on
        Element.setAttribute('name', 'on')
        Element.style.visibility = 'visible'
        Element.style.zIndex = "50";
        document.getElementById('FooterLinkContentButton').style.visibility   = 'visible';
        document.getElementById('FooterLinkContentButton').style.zIndex = "40";
        document.getElementById('AlignFooterContentButton').style.visibility   = 'visible';
        document.getElementById('AlignFooterContentButton').style.zIndex = "40";
        document.getElementById('RemoveFooterLinkContentButton').style.visibility  = 'visible';
        document.getElementById('RemoveFooterLinkContentButton').style.zIndex = "40";
        document.getElementById('FontFaceFooterContentButton').style.visibility= 'visible';
        document.getElementById('FontFaceFooterContentButton').style.zIndex = "40";
        document.getElementById('FooterElement'+'Drag').style.visibility   = 'visible';
        document.getElementById('FooterElement'+'Drag').style.zIndex = "40";
      }else{
        // turn setting for this compositiion element off
        Element.setAttribute('name', 'off')
        Element.style.visibility = 'hdden'
        Element.style.zIndex = "0";
        document.getElementById('FooterLinkContentButton').style.visibility   = 'hidden';
        document.getElementById('FooterLinkContentButton').style.zIndex = "0";
        document.getElementById('AlignFooterContentButton').style.visibility   = 'hidden';
        document.getElementById('AlignFooterContentButton').style.zIndex = "0";
        document.getElementById('RemoveFooterLinkContentButton').style.visibility  = 'hidden';
        document.getElementById('RemoveFooterLinkContentButton').style.zIndex = "0";
        document.getElementById('FontFaceFooterContentButton').style.visibility= 'hidden';
        document.getElementById('FontFaceFooterContentButton').style.zIndex = "0";
        document.getElementById('FooterElement'+'Drag').style.visibility   = 'hidden';
        document.getElementById('FooterElement'+'Drag').style.zIndex = "0";
      }
    }
  }
  function AlignHeaderFooterContent(HeaderFooter){
    // change heading state for this composition element
    Element = document.getElementById(HeaderFooter);
    //get current header state
    CurrentAlignState = parseInt(Element.getAttribute('name'));
    (CurrentAlignState == 4)?  CurrentAlignState = 1 : CurrentAlignState = CurrentAlignState + 1;
    Element.setAttribute( 'name'  , CurrentAlignState.toString());
    //get rid of old image
    while (Element.hasChildNodes()) {
      Element.removeChild(Element.firstChild);
    }
    //replace with new
    img = document.createElement('img');
    img.src = (CurrentAlignState==1)? 'BP Appearance/AlignLeft.png':
              (CurrentAlignState==2)? 'BP Appearance/AlignRight.png':
              (CurrentAlignState==3)? 'BP Appearance/AlignJustified.png' : 'BP Appearance/AlignCentered.png' ;
    img.alt = 'Some png';
    img.width = '25';
    img.height = '25';
    Element.appendChild(img);
    //change content as required
    if(HeaderFooter.includes("Header")){
      CompositionElementContent = document.getElementById('HeaderContent');
           if(CurrentAlignState==1){HeaderContent.style.textAlign = "left"}
      else if(CurrentAlignState==2){HeaderContent.style.textAlign = "right"}
      else if(CurrentAlignState==3){HeaderContent.style.textAlign = "justify"}
      else if(CurrentAlignState==4){HeaderContent.style.textAlign = "center"};
      // update header object
      // Header = {Present: false, Xposition: 0, Yposition: 0, Width: 0, Content: "This is header text", TextAlign: 'center', pageNumberInHeader: false, pageNumber: 0, pageNumberPosition: "", fontName: '', fontStyle: '', fontSize: 10, LineHeightUnits: 0 }
      Header.TextAlign = HeaderContent.style.textAlign
    }else{
      CompositionElementContent = document.getElementById('FooterContent');
           if(CurrentAlignState==1){FooterContent.style.textAlign = "left"}
      else if(CurrentAlignState==2){FooterContent.style.textAlign = "right"}
      else if(CurrentAlignState==3){FooterContent.style.textAlign = "justify"}
      else if(CurrentAlignState==4){FooterContent.style.textAlign = "center"};
      // update header object
      // Header = {Present: false, Xposition: 0, Yposition: 0, Width: 0, Content: "This is header text", TextAlign: 'center', pageNumberInHeader: false, pageNumber: 0, pageNumberPosition: "", fontName: '', fontStyle: '', fontSize: 10, LineHeightUnits: 0 }
      Footer.TextAlign = FooterContent.style.textAlign
    }
  }
  function FontFaceHeaderFooterContent(HeaderFooter, Content){
    // change heading state for this composition element
    if(HeaderFooter.includes("Header")){
      Element = document.getElementById('FontFaceHeaderContentButton');
      //get current header state
      CurrentFontFaceNumber = parseInt(Element.getAttribute('name'));
      (CurrentFontFaceNumber == 4)?  CurrentFontFaceNumber = 1 : CurrentFontFaceNumber = CurrentFontFaceNumber + 1;
      Element.setAttribute( 'name'  , CurrentFontFaceNumber.toString());
      //get rid of old image
      while (Element.hasChildNodes()) {
        Element.removeChild(Element.firstChild);
      }
      //replace with new
      img = document.createElement('img');
      img.src = (CurrentFontFaceNumber==1)? 'BP Appearance/FontFace1.png':
                (CurrentFontFaceNumber==2)? 'BP Appearance/FontFace2.png':
                (CurrentFontFaceNumber==3)? 'BP Appearance/FontFace3.png' : 'BP Appearance/FontFace4.png' ;
      img.alt = 'Some png';
      img.width = '25';
      img.height = '25';
      Element.appendChild(img);
      //change content as required
      let CompositionElementContent = document.getElementById(Content);
      let CurrentFontSpec = CompositionElementContent.style.font;
      //the font style is:   FontStyleSubstitute+" "+FontWeightSubstitute+" "+fontHeightInPX.toString()+"px "+FontFamilySubstitute;
      if(CurrentFontFaceNumber==1){
        FontStyleSubstitute = "normal";
        FontWeightSubstitute = "normal";
        Header.fontStyle = "normal"; // update header object
      }else if(CurrentFontFaceNumber==2){
        FontStyleSubstitute = "italic";
        FontWeightSubstitute = "normal";
        Header.fontStyle = "italic"; // update header object
      }else if(CurrentFontFaceNumber==3){
        FontStyleSubstitute = "normal";
        FontWeightSubstitute = "bold";
        Header.fontStyle = "bold"; // update header object
      }else if(CurrentFontFaceNumber==4){
        FontStyleSubstitute = "italic";
        FontWeightSubstitute = "bold";
        Header.fontStyle = "bolditalic"; // update header object
      };
      fontHeightInPX = (CurrentFontSize/72.0)*PixelsPerScaleUnitHeight*PaperUnits;
      //the font style is:   FontStyleSubstitute+" "+FontWeightSubstitute+" "+fontHeightInPX.toString()+"px "+FontFamilySubstitute;
      CompositionElementContent.style.font = FontStyleSubstitute+" "+FontWeightSubstitute+" "+fontHeightInPX.toString()+"px "+CurrentFont;
    }else{
      Element = document.getElementById('FontFaceFooterContentButton');
      //get current footer state
      CurrentFontFaceNumber = parseInt(Element.getAttribute('name'));
      (CurrentFontFaceNumber == 4)?  CurrentFontFaceNumber = 1 : CurrentFontFaceNumber = CurrentFontFaceNumber + 1;
      Element.setAttribute( 'name'  , CurrentFontFaceNumber.toString());
      //get rid of old image
      while (Element.hasChildNodes()) {
        Element.removeChild(Element.firstChild);
      }
      //replace with new
      img = document.createElement('img');
      img.src = (CurrentFontFaceNumber==1)? 'BP Appearance/FontFace1.png':
                (CurrentFontFaceNumber==2)? 'BP Appearance/FontFace2.png':
                (CurrentFontFaceNumber==3)? 'BP Appearance/FontFace3.png' : 'BP Appearance/FontFace4.png' ;
      img.alt = 'Some png';
      img.width = '25';
      img.height = '25';
      Element.appendChild(img);
      //change content as required
      let CompositionElementContent = document.getElementById(Content);
      let CurrentFontSpec = CompositionElementContent.style.font;
      //the font style is:   FontStyleSubstitute+" "+FontWeightSubstitute+" "+fontHeightInPX.toString()+"px "+FontFamilySubstitute;
      if(CurrentFontFaceNumber==1){
        FontStyleSubstitute = "normal";
        FontWeightSubstitute = "normal";
        Footer.fontStyle = "normal"; // update footer object
      }else if(CurrentFontFaceNumber==2){
        FontStyleSubstitute = "italic";
        FontWeightSubstitute = "normal";
        Footer.fontStyle = "italic"; // update footer object
      }else if(CurrentFontFaceNumber==3){
        FontStyleSubstitute = "normal";
        FontWeightSubstitute = "bold";
        Footer.fontStyle = "bold"; // update footer object
      }else if(CurrentFontFaceNumber==4){
        FontStyleSubstitute = "italic";
        FontWeightSubstitute = "bold";
        Footer.fontStyle = "bolditalic"; // update footer object
      };
      fontHeightInPX = (CurrentFontSize/72.0)*PixelsPerScaleUnitHeight*PaperUnits;
      //the font style is:   FontStyleSubstitute+" "+FontWeightSubstitute+" "+fontHeightInPX.toString()+"px "+FontFamilySubstitute;
      CompositionElementContent.style.font = FontStyleSubstitute+" "+FontWeightSubstitute+" "+fontHeightInPX.toString()+"px "+CurrentFont;
    }
  }
  function ResizeHeaderFooter(HeaderFooter) {
    Element = document.getElementById(HeaderFooter.toString());
    SnapWidthOfCompositionWindow  = CurrentHorizontalSnap*PixelsPerScaleUnitWidth; 
    SnapHeightOfCompositionWindow = CurrentVerticalSnap*PixelsPerScaleUnitHeight*PaperUnits;
    ElementWidth = Number(Element.style.width.slice(0,Element.style.width.indexOf("px")));
    ElementHeight = Number(Element.style.height.slice(0,Element.style.height.indexOf("px")));
    Element.style.width = Math.max(SnapWidthOfCompositionWindow,SnapWidthOfCompositionWindow*Math.round(ElementWidth/SnapWidthOfCompositionWindow)) + "px";
    Element.style.height  = Math.max(SnapHeightOfCompositionWindow,SnapHeightOfCompositionWindow*Math.round(ElementHeight/SnapHeightOfCompositionWindow)) + "px";
  }
//#endregion

// composition board functions
//#region
 function PopulateCompositionElements(InputNumberRangeText){
    //if the field has two numbers, use the 1st as the element number to fill the composition page
    let NumberRangeTextSplit = InputNumberRangeText.split("-")
    CurrentElementNumberForComposition = Number(NumberRangeTextSplit[0]);
    CurrentElementNumberForComposition = Math.min(NumOfTimelineElements,Math.max(1,Math.round(CurrentElementNumberForComposition)));
    CompositionPrintRange1 = Number(NumberRangeTextSplit[0]);
    CompositionPrintRange1 = Math.min(NumOfTimelineElements,Math.max(1,Math.round(CompositionPrintRange1)));
    if(NumberRangeTextSplit.length==2){
      CompositionPrintRange2 = Number(NumberRangeTextSplit[1]);
      CompositionPrintRange2 = Math.min(NumOfTimelineElements,Math.max(1,Math.round(CompositionPrintRange2)));
    }else{
      CompositionPrintRange2 = CompositionPrintRange1;
    }
    return
  }

  function SetPDFDocQuality(QualityNumber){
    PDFImageQuality = QualityNumber;
    document.getElementById("QualityForContent").value = PDFImageQuality;
  }

  function ToggleSpecialCharacters(){
    if(!SpecialCharactersLoaded){
      // populate special characters picker for open fields
      document.getElementById('SpecialCharactersForOpenFields').style.display= 'block'; 
      document.getElementById('SpecialCharactersForOpenFields').style.zIndex= '1500'; 
      dragElement(document.getElementById("SpecialCharactersForOpenFieldsDrag"));     
      for(let m=1; m<=SpecialCharacters[0]; m++ ){
        Unit = document.createElement('button');
        if(m <= 2){
          Unit.setAttribute("class","SpecialCharacterButtonActionSymbol");          
        }else{
          Unit.setAttribute("class","SpecialCharacterButton");
        }        
        Unit.setAttribute( 'onclick'  , 'PutCharInOpenField('+m.toString()+')');      
        Unit.type = "button";
        Unit.innerHTML = String.fromCodePoint(SpecialCharacters[m]);
        document.getElementById("SpecialCharacters").appendChild(Unit); 
      }
      SpecialCharactersLoaded = true;
    }else{
      Node = document.getElementById("SpecialCharacters");
      while (Node.hasChildNodes()) {
        Node.removeChild(Node.firstChild);
      }
      document.getElementById('SpecialCharactersForOpenFields').style.display= 'none';
      SpecialCharactersLoaded = false;
    }
  }

  function CopyToCompositionBoard(FieldNumber){
    // detect if user is picker header content
    if( HeaderOrFooterLinking == "Header" || HeaderOrFooterLinking == "Footer" ){
      //we're picking an element for linking  to the header. call function that does that, rather than add element to the composition board
      if(HeaderOrFooterLinking == "Header"){
        HeaderLinkContent(FieldNumber);
      }else{
        FooterLinkContent(FieldNumber);
      }
      HeaderOrFooterLinking = '';                             //reset flag for heading linking and cursor and set copytocompositionboard cursor to pointer
      document.body.style.cursor = "default";
      let Headers = document.getElementsByClassName('CopyToCompositionBoardButton');
      for(i = 0; i < Headers.length; i++) {
        Headers[i].style.cursor = 'pointer';
      }      
      return
    }

    // getSurrogatePair(astralCodePoint)

    if(!CurrentElementNumberForComposition){
      CurrentElementNumberForComposition = 1;
    }

    //don't add element if it's already on the composition board and NOT a free field/text input for user use
    for(let m=1;m<=NumberOfCompositonElements;m++){
      if(FieldNumber==MapOfCompositionElementsToFieldNumbers[m] && FieldNumber >= 0){return}
    }

    CompositionBoard = document.getElementById("CompositionBoardContent");
    if(NumberOfCompositonElements==0){
      CurrentCompositionElementInsertionPointX = (100*(PageMargins[1]/PaperWidth));
      CurrentMaxYonPage = PageMargins[3];     //1st composition element is at the top of the composition board
    }else{
      //find position/size of lowest element on the composition board
      CurrentMaxYonPage = 0
      for( let m=1; m<=NumberOfCompositonElements; m++){
        let LastElementAdded = document.getElementById('CompositionElement'+MapOfCompositionElementsToFieldNumbers[NumberOfCompositonElements].toString());
        let LastElementBottomPosition = (LastElementAdded.offsetTop+LastElementAdded.offsetHeight)/PixelsPerScaleUnitHeight;
        if(LastElementBottomPosition > CurrentMaxYonPage){CurrentMaxYonPage = LastElementBottomPosition};
      }
    }
    NumberOfCompositonElements++;

    if(FieldNumber == -1){
      if( NumberOfOpenFields == 0 ){
        FieldNumber = -1;
      }else{
        FieldNumber = Math.min(...MapOfCompositionElementsToFieldNumbers.slice(1)) - 1
      }
      NumberOfOpenFields = NumberOfOpenFields - 1
      MapOfCompositionElementsToFieldNumbers[NumberOfCompositonElements]=FieldNumber;
    }else{
      MapOfCompositionElementsToFieldNumbers[NumberOfCompositonElements]=FieldNumber;
      MapOfCompositionElementsToBPHeaders[NumberOfCompositonElements]=BP_Headers[FieldNumber];
    }

    // make a new, numbered composition element
    CompositionElement = document.createElement('div');
    CompositionElement.setAttribute( 'id'  , 'CompositionElement'+FieldNumber.toString() );
    CompositionElement.setAttribute( 'class'  , 'CompositionElement');
    CompositionElement.value = '';
    //add listener to call element size snap function
    CompositionElement.addEventListener('mouseup', function(){ResizeCompositionElement(FieldNumber)});
    CompositionElement.addEventListener('mouseup', function(){IdOfCompositionElementContentWithFocus='CompositionElementContent'+FieldNumber.toString()});
    CompositionElement.oncontextmenu = function(e){e.preventDefault(); SettingsElementComposition(FieldNumber, '') };
                                                                        
    // CompositionElement.style.position = "absolute";
    CompositionElement.style.left = CurrentCompositionElementInsertionPointX.toString()+"%";
    // CompositionElement.style.resize = "both";
    // CompositionElement.style.zIndex = "40";
    let SnapHeightOfCompositionWindow = CurrentVerticalSnap*PaperUnits;
    let NumberOfSnapsToTopMargin = PageMargins[3]/SnapHeightOfCompositionWindow;
    let FractionOfSnap = NumberOfSnapsToTopMargin-Math.floor(NumberOfSnapsToTopMargin);
    SnapTop = SnapHeightOfCompositionWindow*Math.round(CurrentMaxYonPage/SnapHeightOfCompositionWindow) + FractionOfSnap*SnapHeightOfCompositionWindow;
    CompositionElement.style.top =  (100*(SnapTop/PaperWidth)*(CompositionBoard.clientWidth/CompositionBoard.clientHeight)).toString()+"%";
    //add listener to call element size snap function
    // document.getElementById('CompositionElement'+FieldNumber.toString()).addEventListener('mouseup', function(){ResizeCompositionElement(FieldNumber)}); 

    //initial width depends on paper units
    if(PaperUnits==1){
      CompositionElement.style.width = (4.5*PixelsPerScaleUnitWidth).toString()+'px';
    }else{
      CompositionElement.style.width = (10.0*PixelsPerScaleUnitWidth).toString()+'px';
    }      
    CompositionElement.style.height = (CurrentVerticalSnap*PixelsPerScaleUnitHeight*PaperUnits).toString()+'px';
    
    // compositiion element settings toggle (drag handle, title, settings, remove, align, header, columns, fonts, fontsize)
    CompositionElementControl = document.createElement('div');
    CompositionElementControl.setAttribute( 'id'  , 'CompositionElementControl'+FieldNumber.toString() );
    if(FieldNumber<0){
      CompositionElementControl.setAttribute( 'class'  , 'CompositionElementControlNarrow');
    }else{
      CompositionElementControl.setAttribute( 'class'  , 'CompositionElementControl');
    }
    CompositionElementControl.setAttribute( 'name'  , 'off');
    CompositionElementControl.style.display = 'none';
    while (CompositionElementControl.hasChildNodes()) {
      CompositionElementControl.removeChild(CompositionElementControl.lastChild);
    }


    // add a drag element to the composition element control
    CompositionElementDrag = document.createElement('div');
    CompositionElementDrag.setAttribute( 'id'  , 'CompositionElement'+FieldNumber.toString()+'Drag');
    CompositionElementDrag.setAttribute( 'class'  , 'CompositionElementDrag');
    //image for dragging
    img = document.createElement('img');
    img.src = 'BP Appearance/Move.png';
    img.alt = 'Some png';
    img.width = '25';
    img.height = '25';
    CompositionElementDrag.appendChild(img);
    CompositionElementControl.appendChild(CompositionElementDrag);

    //add font face button to the composition element control
    CompositionElementFontFace = document.createElement('button');
    CompositionElementFontFace.setAttribute( 'id'  , 'FontFaceElementContentButton'+FieldNumber.toString() );
    CompositionElementFontFace.setAttribute( 'class'  , 'FontFaceElementContentButton');
    CompositionElementFontFace.setAttribute( 'title'  , 'Choose a font face');
    CompositionElementFontFace.setAttribute( 'onclick'  , 'FontFaceElementComposition('+FieldNumber.toString()+' ) ');
        //initial condition for font face is default to that for the whole composition
    if( CurrentFontFace=="normal"){CompositionElementFontFace.setAttribute( 'name'  , '1')}
    else if(CurrentFontFace=="italic"){CompositionElementFontFace.setAttribute( 'name'  , '2')}
    else if(CurrentFontFace=="bold"){CompositionElementFontFace.setAttribute( 'name'  , '3')}
    else if(CurrentFontFace=='bolditalic'){CompositionElementFontFace.setAttribute( 'name'  , '4')}
    else{CompositionElementFontFace.setAttribute( 'name'  , '1')}
    img = document.createElement('img');
    img.src = 'BP Appearance/FontFace1.png';
    img.alt = 'Some png';
    img.width = '22';
    img.height = '22';
    CompositionElementFontFace.appendChild(img);
    CompositionElementControl.appendChild(CompositionElementFontFace);

    //add spacer button to the composition element control
    CompositionElementAddSpace = document.createElement('button');
    CompositionElementAddSpace.setAttribute( 'id'  , 'AddSpaceElementContentButton'+FieldNumber.toString() );
    CompositionElementAddSpace.setAttribute( 'class'  , 'AddSpaceElementContentButton');
    CompositionElementAddSpace.setAttribute( 'title'  , 'Add 1/2 line spaces above the element');
    CompositionElementAddSpace.setAttribute( 'onclick'  , 'AddSpaceElementComposition('+FieldNumber.toString()+' ) ');
    //initial condition for header is: left
    CompositionElementAddSpace.setAttribute( 'name'  , '0');
    img = document.createElement('img');
    img.src = 'BP Appearance/AddSpace0.png';
    img.alt = 'Some png';
    img.width = '22';
    img.height = '22';
    CompositionElementAddSpace.appendChild(img);
    CompositionElementControl.appendChild(CompositionElementAddSpace);

    //add column button to the composition element control
    CompositionElementColumns = document.createElement('button');
    CompositionElementColumns.setAttribute( 'id'  , 'ColumnsElementContentButton'+FieldNumber.toString() );
    CompositionElementColumns.setAttribute( 'class'  , 'ColumnsElementContentButton');
    CompositionElementColumns.setAttribute( 'title'  , 'Number of columns');
    CompositionElementColumns.setAttribute( 'onclick'  , 'ColumnsElementComposition('+FieldNumber.toString()+' ) ');
    //initial condition for header is: left
    CompositionElementColumns.setAttribute( 'name'  , '1');
    img = document.createElement('img');
    img.src = 'BP Appearance/Columns1.png';
    img.alt = 'Some png';
    img.width = '22';
    img.height = '22';
    CompositionElementColumns.appendChild(img);
    CompositionElementControl.appendChild(CompositionElementColumns);

    //add align button to the composition element control
    CompositionElementAlign = document.createElement('button');
    CompositionElementAlign.setAttribute( 'id'  , 'AlignElementContentButton'+FieldNumber.toString() );
    CompositionElementAlign.setAttribute( 'class'  , 'AlignElementContentButton');
    CompositionElementAlign.setAttribute( 'title'  , 'Text alignment and justification');
    CompositionElementAlign.setAttribute( 'onclick'  , 'AlignElementComposition('+FieldNumber.toString()+' ) ');
    //initial condition for header is: left
    CompositionElementAlign.setAttribute( 'name'  , '1');
    img = document.createElement('img');
    img.src = 'BP Appearance/AlignLeft.png';
    img.alt = 'Some png';
    img.width = '22';
    img.height = '22';
    CompositionElementAlign.appendChild(img);
    CompositionElementControl.appendChild(CompositionElementAlign);

    //add header button to the composition element control
    CompositionElementHeader = document.createElement('button');
    CompositionElementHeader.setAttribute( 'id'  , 'HeaderElementContentButton'+FieldNumber.toString() );
    CompositionElementHeader.setAttribute( 'class'  , 'HeaderElementContentButton');
    CompositionElementHeader.setAttribute( 'title'  , 'Heading for this section');
    CompositionElementHeader.setAttribute( 'onclick'  , 'HeaderElementComposition('+FieldNumber.toString()+' ) ');
    //initial condition for header is: InLine
    CompositionElementHeader.setAttribute( 'name'  , '1');
    img = document.createElement('img');
    img.src = 'BP Appearance/HeaderInLine.png';
    img.alt = 'Some png';
    img.width = '22';
    img.height = '22';
    CompositionElementHeader.appendChild(img);
    CompositionElementControl.appendChild(CompositionElementHeader);

    //add remove button to the composition element control
    CompositionElementRemove = document.createElement('button');
    CompositionElementRemove.setAttribute( 'id'  , 'RemoveElementContentButton'+FieldNumber.toString() );
    CompositionElementRemove.setAttribute( 'class'  , 'RemoveElementContentButton');
    CompositionElementRemove.setAttribute( 'title'  , 'Remove this element');
    CompositionElementRemove.setAttribute( 'onclick'  , 'RemoveElementFromComposition('+FieldNumber.toString()+' ) ');
    img = document.createElement('img');
    img.src = 'BP Appearance/Bin.png';
    img.alt = 'Some png';
    img.width = '22';
    img.height = '22';
    CompositionElementRemove.appendChild(img);
    CompositionElementControl.appendChild(CompositionElementRemove);

   //add title to the composition element control
   CompositionElementTitle = document.createElement('div');
   CompositionElementTitle.setAttribute( 'id'  , 'CompositionElementTitle'+FieldNumber.toString() );
   CompositionElementTitle.setAttribute( 'class'  , 'CompositionElementTitle');
   // get trimed BP header name for title of composition element
   CompositionElementTitle.innerHTML = "";
   if(FieldNumber >= 0){ CompositionElementTitle.innerHTML = " " + BP_Headers[FieldNumber].trim().slice(3) + " "};
   CompositionElementControl.appendChild(CompositionElementTitle);    

    //header for content of element (excepting negative field)
    if(FieldNumber >= 0){
      CompositionElementContentHeader = document.createElement('div');
      CompositionElementContentHeader.setAttribute( 'id'  , 'CompositionElementContentHeader'+FieldNumber.toString() );
      CompositionElementContentHeader.setAttribute( 'class'  , 'CompositionElementContentHeader');
      CompositionElementContentHeader.innerHTML = "";   //initially, no content in header, it is in-line and included with content
      CompositionElementContentHeader.style.zIndex = "0";
      fontHeightInPX = (CurrentFontSize/72.0)*PixelsPerScaleUnitHeight*PaperUnits;
      if(CurrentFontFace.includes("Oblique")){FontStyleSubstitute = "italic"} else if (CurrentFontFace.includes("italic")){FontStyleSubstitute = "italic"} else {FontStyleSubstitute = "normal"};
      if(CurrentFontFace.includes("Bold") || CurrentFontFace.includes("bold")){FontWeightSubstitute = "bold"} else {FontWeightSubstitute = "normal"};
      //special cases of names that differ from OS recognized fonts (due to font handling in jsPDF)
      if(CurrentFont=="Arial-Narrow"){
        FontFamilySubstitute = "Arial_Narrow";
      }else if(CurrentFont=="BookmanOldStyle"){
        FontFamilySubstitute = "Bookman_Old_Style";
      }else if(CurrentFont=="Times-Roman"){
        FontFamilySubstitute = "Times_New_Roman";
      }else{FontFamilySubstitute = CurrentFont}
      CompositionElementContentHeader.style.font = FontStyleSubstitute+" "+FontWeightSubstitute+" "+fontHeightInPX.toString()+"px "+FontFamilySubstitute;
      if(!BP_Headers[FieldNumber].trim().includes("Image")){
          CompositionElementContent = document.createElement('div');
          CompositionElementContent.setAttribute( 'id'  , 'CompositionElementContent'+FieldNumber.toString() );
          CompositionElementContent.setAttribute( 'class'  , 'CompositionElementContent');
          CompositionElementContent.innerHTML = BP_Headers[FieldNumber].trim().slice(3) + ":  " // + data[CurrentElementNumberForComposition][BPHeadersOfUserData[FieldNumber]];          
          CompositionElementContent.style.zIndex = "0";
        }else{
          //content is image
          CompositionElementContent = document.createElement('img');
          CompositionElementContent.setAttribute( 'id'  , 'CompositionElementContent'+FieldNumber.toString() );
          CompositionElementContent.setAttribute( 'class'  , 'CompositionElementContent');
          CompositionElementContent.src = "Image";
          CompositionElementContent.style.opacity = "1.0";
        }
    }else{
      //content for negative field number is text input; that is, user input
      CompositionElementContent = document.createElement('input');
      CompositionElementContent.setAttribute( 'id'  , 'CompositionElementContent'+FieldNumber.toString() );
      CompositionElementContent.setAttribute( 'type', 'text');
      CompositionElementContent.style.zIndex = "0";
      CompositionElementContent.style.padding = 0, 0, 0, 0;
      CompositionElementContent.style.margins = 0, 0, 0, 0;
      CompositionElementContent.style.backgroundColor = 'transparent';
    }

    fontHeightInPX = (CurrentFontSize/72.0)*PixelsPerScaleUnitHeight*PaperUnits;
    if(CurrentFontFace.includes("Oblique")){FontStyleSubstitute = "italic"} else if (CurrentFontFace.includes("italic")){FontStyleSubstitute = "italic"} else {FontStyleSubstitute = "normal"}
    if(CurrentFontFace.includes("Bold") || CurrentFontFace.includes("bold")){FontWeightSubstitute = "bold"} else {FontWeightSubstitute = "normal"}
    //special cases of names that differ from OS recognized fonts (due to font handling in jsPDF)
    if(CurrentFont=="Arial-Narrow"){
      FontFamilySubstitute = "Arial_Narrow";
    }else if(CurrentFont=="BookmanOldStyle"){
      FontFamilySubstitute = "Bookman_Old_Style";
    }else if(CurrentFont=="Times-Roman"){
      FontFamilySubstitute = "Times_New_Roman";
    }else{FontFamilySubstitute = CurrentFont}
    CompositionElementContent.style.font = FontStyleSubstitute+" "+FontWeightSubstitute+" "+fontHeightInPX.toString()+"px "+FontFamilySubstitute;
    checkfontspec = CompositionElementContent.style.font;
    CompositionElementContent.style.columnCount = "1";

    //add header and content to composition element
    if(FieldNumber >= 0){CompositionElement.appendChild(CompositionElementContentHeader)};
    CompositionElement.appendChild(CompositionElementContent);
    //add control(s) to composition element AFTER content (puts it at the bottom of the element. sticky scroll keeps it there)
    CompositionElement.appendChild(CompositionElementControl);
    //add composition element to composition board
    CompositionBoard.appendChild(CompositionElement)

    // trigger dragging
    dragElement(document.getElementById('CompositionElement'+FieldNumber.toString()+'Drag'));// gray out selection
    //gray out element from palette, except for open field text/image elements
    if(FieldNumber >= 0){document.getElementById('CopyToCompositionBoardButton'+FieldNumber.toString()).style.opacity = "0.5"};
  }

  function HeaderElementComposition(FieldNumber){
    if(FieldNumber < 0){return};
    // change heading state for this composition element
    Element = document.getElementById('HeaderElementContentButton'+FieldNumber.toString() );
    //get current heading state
    CurrentHeaderState = parseInt(Element.getAttribute('name'));
    (CurrentHeaderState == 5)?  CurrentHeaderState = 1 : CurrentHeaderState = CurrentHeaderState + 1;
    Element.setAttribute( 'name'  , CurrentHeaderState.toString());
    //get rid of old image
    while (Element.hasChildNodes()) {
      Element.removeChild(Element.firstChild);
    }
    //replace with new
    img = document.createElement('img');
    img.src = (CurrentHeaderState==1)? 'BP Appearance/HeaderInLine.png':
              (CurrentHeaderState==2)? 'BP Appearance/HeaderInLineCaps.png' :
              (CurrentHeaderState==3)? 'BP Appearance/HeaderAbove.png' :
              (CurrentHeaderState==4)? 'BP Appearance/HeaderAboveCaps.png' : 'BP Appearance/HeaderNone.png' ;
    img.alt = 'Some png';
    img.width = '25';
    img.height = '25';
    Element.appendChild(img);
    //change content as required
         if(CurrentHeaderState==1){HeaderText = "HeaderInLine:  "    }
    else if(CurrentHeaderState==2){HeaderText = "HeaderInLineCaps:  "}
    else if(CurrentHeaderState==3){HeaderText = "HeaderAbove"        }
    else if(CurrentHeaderState==4){HeaderText = "HeaderAboveCaps"    }
    else if(CurrentHeaderState==5){HeaderText = "NoHeader"           }
    // the CompositionElementContentHeader element only contains text if the header is ABOVE the body of the element text
    // that is, header format states 3 and 4.
    CompositionElementContentHeader = document.getElementById('CompositionElementContentHeader'+FieldNumber.toString() );
         if(CurrentHeaderState==1){CompositionElementContentHeader.innerHTML = ""}
    else if(CurrentHeaderState==2){CompositionElementContentHeader.innerHTML = ""}
    else if(CurrentHeaderState==3){CompositionElementContentHeader.innerHTML = HeaderText}
    else if(CurrentHeaderState==4){CompositionElementContentHeader.innerHTML = HeaderText}
    else if(CurrentHeaderState==5){CompositionElementContentHeader.innerHTML = ""}

         if(CurrentHeaderState==1){CompositionElementContentHeader.style.height = "0"}
    else if(CurrentHeaderState==2){CompositionElementContentHeader.style.height = "0"}
    else if(CurrentHeaderState==3){CompositionElementContentHeader.style.height = (CurrentFontSize/72.0)*PixelsPerScaleUnitHeight*PaperUnits.toString()+"px"}
    else if(CurrentHeaderState==4){CompositionElementContentHeader.style.height = (CurrentFontSize/72.0)*PixelsPerScaleUnitHeight*PaperUnits.toString()+"px"}
    else if(CurrentHeaderState==5){CompositionElementContentHeader.style.height = "0"}
  }

  function AlignElementComposition(FieldNumber){
    // change heading state for this composition element
    Element = document.getElementById('AlignElementContentButton'+FieldNumber.toString() );
    //get current header state
    CurrentAlignState = parseInt(Element.getAttribute('name'));
    (CurrentAlignState == 4)?  CurrentAlignState = 1 : CurrentAlignState = CurrentAlignState + 1;
    Element.setAttribute( 'name'  , CurrentAlignState.toString());
    //get rid of old image
    while (Element.hasChildNodes()) {
      Element.removeChild(Element.firstChild);
    }
    //replace with new
    img = document.createElement('img');
    img.src = (CurrentAlignState==1)? 'BP Appearance/AlignLeft.png':
              (CurrentAlignState==2)? 'BP Appearance/AlignRight.png':
              (CurrentAlignState==3)? 'BP Appearance/AlignJustified.png' : 'BP Appearance/AlignCentered.png' ;
    img.alt = 'Some png';
    img.width = '25';
    img.height = '25';
    Element.appendChild(img);
    //change content as required
    CompositionElementContent = document.getElementById('CompositionElementContent'+FieldNumber.toString() );
         if(CurrentAlignState==1){CompositionElementContent.style.textAlign = "left"}
    else if(CurrentAlignState==2){CompositionElementContent.style.textAlign = "right"}
    else if(CurrentAlignState==3){CompositionElementContent.style.textAlign = "justify"}
    else if(CurrentAlignState==4){CompositionElementContent.style.textAlign = "center"};
  }

  function ColumnsElementComposition(FieldNumber){
    // change heading state for this composition element
    Element = document.getElementById('ColumnsElementContentButton'+FieldNumber.toString() );
    //get current header state
    CurrentColumnsState = parseInt(Element.getAttribute('name'));
    (CurrentColumnsState == 4)?  CurrentColumnsState = 1 : CurrentColumnsState = CurrentColumnsState + 1;
    Element.setAttribute( 'name'  , CurrentColumnsState.toString());
    //get rid of old image
    while (Element.hasChildNodes()) {
      Element.removeChild(Element.firstChild);
    }
    //replace with new
    img = document.createElement('img');
    img.src = (CurrentColumnsState==1)? 'BP Appearance/Columns1.png':
              (CurrentColumnsState==2)? 'BP Appearance/Columns2.png':
              (CurrentColumnsState==3)? 'BP Appearance/Columns3.png' : 'BP Appearance/Columns4.png' ;
    img.alt = 'Some png';
    img.width = '25';
    img.height = '25';
    Element.appendChild(img);
    //change content as required
    CompositionElementContent = document.getElementById('CompositionElementContent'+FieldNumber.toString() );
         if(CurrentColumnsState==1){CompositionElementContent.style.columnCount = "1"}
    else if(CurrentColumnsState==2){CompositionElementContent.style.columnCount = "2"}
    else if(CurrentColumnsState==3){CompositionElementContent.style.columnCount = "3"}
    else if(CurrentColumnsState==4){CompositionElementContent.style.columnCount = "4"};
  }

  function FontFaceElementComposition(FieldNumber){
    // change heading state for this composition element
    Element = document.getElementById('FontFaceElementContentButton'+FieldNumber.toString() );
    //get current header state
    CurrentFontFaceNumber = parseInt(Element.getAttribute('name'));
    (CurrentFontFaceNumber == 4)?  CurrentFontFaceNumber = 1 : CurrentFontFaceNumber = CurrentFontFaceNumber + 1;
    Element.setAttribute( 'name'  , CurrentFontFaceNumber.toString());
    //get rid of old image
    while (Element.hasChildNodes()) {
      Element.removeChild(Element.firstChild);
    }
    //replace with new
    img = document.createElement('img');
    img.src = (CurrentFontFaceNumber==1)? 'BP Appearance/FontFace1.png':
              (CurrentFontFaceNumber==2)? 'BP Appearance/FontFace2.png':
              (CurrentFontFaceNumber==3)? 'BP Appearance/FontFace3.png' : 'BP Appearance/FontFace4.png' ;
    img.alt = 'Some png';
    img.width = '25';
    img.height = '25';
    Element.appendChild(img);
    //change content as required
    let CompositionElementContent = document.getElementById('CompositionElementContent'+FieldNumber.toString() );
    let CurrentFontSpec = CompositionElementContent.style.font;
    //the font style is:   FontStyleSubstitute+" "+FontWeightSubstitute+" "+fontHeightInPX.toString()+"px "+FontFamilySubstitute;
    if(CurrentFontFaceNumber==1){
      FontStyleSubstitute = "normal";
      FontWeightSubstitute = "normal";
    }else if(CurrentFontFaceNumber==2){
      FontStyleSubstitute = "italic";
      FontWeightSubstitute = "normal";
    }else if(CurrentFontFaceNumber==3){
      FontStyleSubstitute = "normal";
      FontWeightSubstitute = "bold";
    }else if(CurrentFontFaceNumber==4){
      FontStyleSubstitute = "italic";
      FontWeightSubstitute = "bold";
    };
    fontHeightInPX = (CurrentFontSize/72.0)*PixelsPerScaleUnitHeight*PaperUnits;
    //the font style is:   FontStyleSubstitute+" "+FontWeightSubstitute+" "+fontHeightInPX.toString()+"px "+FontFamilySubstitute;
    CompositionElementContent.style.font = FontStyleSubstitute+" "+FontWeightSubstitute+" "+fontHeightInPX.toString()+"px "+CurrentFont;
  }

  function AddSpaceElementComposition(FieldNumber){
    // change spacingt state for this composition element
    Element = document.getElementById('AddSpaceElementContentButton'+FieldNumber.toString() );
    //get current spacing state
    CurrentAddSPaceState = parseInt(Element.getAttribute('name'));
    (CurrentAddSPaceState == 4)?  CurrentAddSPaceState = 0 : CurrentAddSPaceState = CurrentAddSPaceState + 1;
    Element.setAttribute( 'name'  , CurrentAddSPaceState.toString());
    // get rid of old image
    while (Element.hasChildNodes()) {
      Element.removeChild(Element.firstChild);
    }
    //replace with new
    img = document.createElement('img');
    img.src = (CurrentAddSPaceState==0)? 'BP Appearance/AddSpace0.png':
              (CurrentAddSPaceState==1)? 'BP Appearance/AddSpace1.png':
              (CurrentAddSPaceState==2)? 'BP Appearance/AddSpace2.png':
              (CurrentAddSPaceState==3)? 'BP Appearance/AddSpace3.png': 'BP Appearance/AddSpace4.png' ;
    img.alt = 'Some png';
    img.width = '25';
    img.height = '25';
    Element.appendChild(img);
  }

  function RemoveElementFromComposition(FieldNumber){
    Element = document.getElementById('CompositionElement'+FieldNumber.toString());
    Element.removeEventListener('mouseup', function(){ResizeCompositionElement(FieldNumber)});             
    Element.remove();
    for(m=1;m<=NumberOfCompositonElements;m++){
      if(MapOfCompositionElementsToFieldNumbers[m]==FieldNumber){
        NumberToRemove = m;
        break;
      }
    }
    MapOfCompositionElementsToFieldNumbers.splice(NumberToRemove,1);
    MapOfCompositionElementsToBPHeaders.splice(NumberToRemove,1);
    NumberOfCompositonElements = NumberOfCompositonElements - 1;
    // ungray out selection
    if(FieldNumber >= 0 ){document.getElementById('CopyToCompositionBoardButton'+FieldNumber.toString()).style.opacity = "1.0"};
    if(FieldNumber < 0 ){NumberOfOpenFields = NumberOfOpenFields + 1.0};
  }

  function SettingsElementComposition(FieldNumber, action){
    // toggle settings of this compositiion element on/off
    Element = document.getElementById('CompositionElementControl'+FieldNumber.toString() );
    // if(FieldNumber<0){Element.setAttribute('class','CompositionElementControlNarrow')}
    let SettingsState = Element.getAttribute('name');
    if( action == 'store&hide' ){
      Element.style.display  = 'none';
      document.getElementById('CompositionElementContent'+FieldNumber.toString() ).style.visibility = "hidden";
      if( SettingsState=='off' ){ //since we need to reset these if we retores, reverse the toggle setting for visibilty.
        Element.setAttribute('name', 'on')
      }else{
        Element.setAttribute('name', 'off')
      }      
    }else{
      if( SettingsState=='on' ){
         // turn setting for this compositiion element off
         Element.setAttribute('name', 'off');
         Element.style.display  = 'none';
         document.getElementById('CompositionElementContent'+FieldNumber.toString() ).style.visibility = "visible";
      }else if(SettingsState=='off'){
         // turn setting for this compositiion element on
         Element.setAttribute('name', 'on')
         Element.style.display  = 'flex';
         document.getElementById('CompositionElementContent'+FieldNumber.toString() ).style.visibility = "hidden";         
      }
    }
  }

  function RemoveAllElementsFromComposition(){
    for(m=1;m<=NumberOfCompositonElements;m++){
      FieldNumber = MapOfCompositionElementsToFieldNumbers[m]
        if(document.getElementById('CompositionElement'+FieldNumber.toString())){
        document.getElementById('CompositionElement'+FieldNumber.toString()).removeEventListener('mouseup', function(){ResizeCompositionElement(FieldNumber)});             
        document.getElementById('CompositionElement'+FieldNumber.toString() ).remove();
        if(FieldNumber >= 0 ){document.getElementById('CopyToCompositionBoardButton'+FieldNumber.toString()).style.opacity = "1.0"};
      }
    }
    //reactiveate resizing and contextmenu trapping
    if(Header.Present){
      HeaderElement = document.getElementById('HeaderElement');
      HeaderElement.addEventListener('mouseup', function(){ResizeHeaderFooter("HeaderElement")}); 
      HeaderElement.oncontextmenu = function(e){e.preventDefault(); SettingsHeaderFooter("HeaderSettings") };
      dragElement(document.getElementById('HeaderElementDrag'));
      if( Header.ContentLink == 0 ){
          if(document.getElementById('HeaderContent') != null){
            document.getElementById('HeaderContent').value = Header.Content;
          }
        }
    }
    if(Footer.Present){
      FooterElement = document.getElementById('FooterElement');
      FooterElement.addEventListener('mouseup', function(){ResizeHeaderFooter("FooterElement")}); 
      FooterElement.oncontextmenu = function(e){e.preventDefault(); SettingsHeaderFooter("FooterSettings") };
      dragElement(document.getElementById('FooterElementDrag'));
      if( Footer.ContentLink <= 0 ){
        if(document.getElementById('FooterContent') != null){
          document.getElementById('FooterContent').value = Footer.Content;
        }        
      }
    }
    NumberOfCompositonElements = 0;
    NumberOfOpenFields = 0;
    MapOfCompositionElementsToFieldNumbers = [];
    MapOfCompositionElementsToBPHeaders = [];

  }

  function BuildRulers(){
    // if(!CompositionRulersHaveBeenBuilt){
      //build horizontal ruler only once 
      Element = document.getElementById("CompositionBoardContent");
      Ruler = document.createElement('div');
      Ruler.setAttribute('class','rulerH');
      Ruler.setAttribute('id','rulerH');
      Ruler.style.width =  (100*(10/PaperWidth)).toString()+"%" ;
      Ruler.style.positiion = "absolute";
      Ruler.style.top = "0%";
      Ruler.style.left = "0%"
      Ruler.style.padding = "0 0 0 0";
      Ruler.style.margin = "0 0 0 0";
      NumberOfUnits = 90;
      for( m=0; m<=NumberOfUnits; m++){
        Unit = document.createElement('div');
        Unit.setAttribute("class","unitH");
        Unit.style.left = (m*10).toString()+"%";
        if(m>0){
          UnitName = document.createElement('div');
          UnitName.style.position = "absolute";
          UnitName.style.bottom = "-10px";
          UnitName.style.font = "10px sans-serif";
          UnitName.innerHTML = m.toString();
          Unit.appendChild(UnitName);
        }
        for( p=1; p<=9; p++){
          Tenths = document.createElement('div');
          Tenths.setAttribute("class","tenthsH");
          Tenths.style.left = (p*10).toString()+"%";
          if(p==5){Tenths.style.height = "10px";} // midway mark in unit is taller
          Unit.appendChild(Tenths);
        }
        Ruler.appendChild(Unit);
      }
      //add unit div without tenths divs
      Unit = document.createElement('div');
      Unit.setAttribute("class","unitH");
      UnitName = document.createElement('div');
      Unit.style.left = ((NumberOfUnits+1)*10).toString()+"%";
        UnitName.style.position = "absolute";
        UnitName.style.bottom = "-10px";
        UnitName.style.font = "10px sans-serif";
        UnitName.innerHTML = (NumberOfUnits+1).toString();
        Unit.appendChild(UnitName);
      Ruler.appendChild(Unit);
      Element.appendChild(Ruler);
    // }

    //always build the vertical ruler.
    if( typeof document.getElementById("rulerV") != "undefined" && document.getElementById("rulerV") != null){document.getElementById("rulerV").remove()}

    // if(!document.getElementById("rulerV")){}else{document.getElementById("rulerV").remove()}


    Ruler = document.createElement('div');
    Ruler.setAttribute('class','rulerV');
    Ruler.setAttribute('id','rulerV');
    Ruler.style.height = (100*(10/PaperWidth)*(Element.clientWidth/Element.clientHeight)).toString()+"%" ;
    Ruler.style.position = "absolute";
    // Ruler.style.top = "0%";
    Ruler.style.left = "0%"
    Ruler.style.padding = "0 0 0 0";
    Ruler.style.margin = "0 0 0 0";
    NumberOfUnits = 120;
    NumberOfTenthsInPaperLength = PaperLength*10;
    NumberOfTenthsSoFar = 0;
    NumberOfTenths = 0;
    PageNumber = 1;
    m = 0;
    for( mm=0; mm<=NumberOfUnits; mm++){
      Unit = document.createElement('div');
      Unit.setAttribute("class","unitV");
      Unit.style.top = (NumberOfTenthsSoFar).toString()+"%";
      if(m==0 && mm>0){
        Unit.style.width = "200px";
        Unit.style.borderColor = "red";
        Unit.style.color = "var(--AlertColor)";
        Unit.innerHTML = PageNumber;
        PageNumber = PageNumber + 1;
        }else if(m==0){
          PageNumber = PageNumber + 1;
        }else{
        UnitName = document.createElement('div');
        UnitName.style.position = "absolute";
        UnitName.style.right = "-8px";
        UnitName.style.font = "10px sans-serif";
        UnitName.innerHTML = m.toString();
        Unit.appendChild(UnitName);
      }
      for( p=1; p<=9; p++){
        Tenths = document.createElement('div');
        Tenths.setAttribute("class","tenthsV");
        Tenths.style.top = (p*10).toString()+"%";
        if(p==5){Tenths.style.width = "10px";} // midway mark in unit is wider
        NumberOfTenths = NumberOfTenths + 1;
        if(NumberOfTenths==NumberOfTenthsInPaperLength){
          Tenths.style.width = "200px";
          Tenths.style.borderColor = "red";
          Unit.appendChild(Tenths);
          NumberOfTenths = 0;
          m = -1; // need to rest m, but we increment at loop's end; thus -1
          break;
        }else{
          NumberOfTenthsSoFar = NumberOfTenthsSoFar + 1;
          Unit.appendChild(Tenths);
        }
      }
      Ruler.appendChild(Unit);
      NumberOfTenths = NumberOfTenths + 1;
      NumberOfTenthsSoFar = NumberOfTenthsSoFar + 1;
      if(NumberOfTenths==NumberOfTenthsInPaperLength){
          NumberOfTenths = 0;
          m = 0;
      }else{ m=m+1}
    }
    //add unit div without tenths divs
    Unit = document.createElement('div');
    Unit.setAttribute("class","unitV");
    UnitName = document.createElement('div');
    Unit.style.top = ((NumberOfUnits+1)*10).toString()+"%";
      UnitName.style.position = "absolute";
      UnitName.style.right = "-8px";
      UnitName.style.font = "10px sans-serif";
      UnitName.innerHTML = (NumberOfUnits+1).toString();
      Unit.appendChild(UnitName);
    Ruler.appendChild(Unit);
    document.getElementById("CompositionBoardContent").appendChild(Ruler);
    NumberOfPagesDisplayed = PageNumber;
    CompositionRulersHaveBeenBuilt = true;
    // re-establish pixel-to-unit scaling
    let WidthInPixels = Element.clientWidth
    PixelsPerScaleUnitWidth = WidthInPixels/PaperWidth;
    PixelsPerScaleUnitHeight = PixelsPerScaleUnitWidth; //*(Element.clientWidth/Element.clientHeight);
  }

  async function SaveCompositionBoard(){
    for(let m = 1; m <= NumberOfCompositonElements; m++){
      FieldNumber = MapOfCompositionElementsToFieldNumbers[m];
      if(FieldNumber < 0){
        element = document.getElementById('CompositionElementContent'+FieldNumber.toString());
        element.setAttribute("name", element.value);
      }
    }
    if( Header.ContentLink <= 0 ){
      if(document.getElementById('HeaderContent')){
        Header.Content = document.getElementById('HeaderContent').value;
      }
    }
    if( Footer.ContentLink <= 0 ){
      if(document.getElementById('FooterContent')){
        Footer.Content = document.getElementById('FooterContent').value;
      }
    }
    
    DataToSave =  PaperName +"|"+ PaperNumber +"|"+ PaperWidth +"|"+ PaperLength +"|"+ PaperUnits +"|"+ PageMargins[1] +"|"+ PageMargins[2] +"|"+ PageMargins[3] +"|"+ PageMargins[4] +"|"+ 
                  CurrentFont +"|"+ CurrentFontFace +"|"+ CurrentFontSize +"|"+ 
                  CurrentCompositionElementInsertionPointX +"|"+ CurrentCompositionElementInsertionPointY +"|"+ CurrentVerticalSnap +"|"+ CurrentHorizontalSnap +"|"+ NumberRangeText +"|"+ PixelsPerScaleUnitWidth +"|"+ PixelsPerScaleUnitHeight +"|"+ PDFImageQuality+"|"+
                  NumberOfCompositonElements +"!"+ NumberOfOpenFields +"!"+ MapOfCompositionElementsToFieldNumbers +"!" + BPHeadersOfUserData + "!" + MapOfCompositionElementsToBPHeaders + "!" + JSON.stringify(Header) + "!" + JSON.stringify(Footer) + "!" + CurrentHeaderState + "!" + PageNumberLayoutNumber.toString() + "!" +document.getElementById("CompositionBoardContent").innerHTML;
    compressedData = LZString.compressToUTF16(DataToSave)
    //compressedData = DataToSave

    try{BookPageantPDFCompositionFileHandle;
      PDFfileHandleKeep = await window.showSaveFilePicker({startIn: BookPageantPDFCompositionFileHandle[0], suggestedName: NameOfLibrary+".PDFformat.bp", types: [{descriiption: 'BP file', accept:{'text/plain':['.PDFformat.bp']}}]});      
    }catch{
      PDFfileHandleKeep = await window.showSaveFilePicker({startIn: BookPageantHomeDirectoryHandle, suggestedName: NameOfLibrary+".PDFformat.bp", types: [{descriiption: 'BP file', accept:{'text/plain':['.PDFformat.bp']}}]});      
    }
    let writable = await PDFfileHandleKeep.createWritable();
    await writable.write(compressedData);
    await writable.close();
  }

  async function RetrieveCompositionBoard(call=""){
    // get a PDF  composition file
    try{BookPageantPDFCompositionFileHandle;
      BookPageantPDFCompositionFileHandle = await window.showOpenFilePicker({startIn: BookPageantPDFCompositionFileHandle[0], types: [{descriiption: 'BP file', accept:{'text/plain':['.PDFformat.bp']}}]});
    }catch{
      BookPageantPDFCompositionFileHandle = await window.showOpenFilePicker({startIn: BookPageantHomeDirectoryHandle, types: [{descriiption: 'BP file', accept:{'text/plain':['.PDFformat.bp']}}]});
    }
    InputFile = await BookPageantPDFCompositionFileHandle[0].getFile();
    localforage.setItem('BookPageantPDFCompositionFileHandle', BookPageantPDFCompositionFileHandle);
    
    // clear contents of composition board, including elements, margins and rulers
    let Node = document.getElementById("CompositionBoardContent");
    while (Node.hasChildNodes()) {
      Node.removeChild(Node.firstChild);
    }

    CompressedReader = new FileReader();
    CompressedReader.onload = async function(){
      compressedData = CompressedReader.result;
      decompressedData = LZString.decompressFromUTF16(compressedData);
      let SeparatedDecompressedData = decompressedData.split("|");

      PaperName                                 = SeparatedDecompressedData[0];
      PaperNumber                               = Number(SeparatedDecompressedData[1]);
      PaperWidth                                = Number(SeparatedDecompressedData[2]);
      PaperLength                               = Number(SeparatedDecompressedData[3]);
      PaperUnits                                = SeparatedDecompressedData[4];
      PageMargins[1]                            = Number(SeparatedDecompressedData[5]);
      PageMargins[2]                            = Number(SeparatedDecompressedData[6]);
      PageMargins[3]                            = Number(SeparatedDecompressedData[7]);
      PageMargins[4]                            = Number(SeparatedDecompressedData[8]);
      CurrentFont                               = SeparatedDecompressedData[9];
      CurrentFontFace                           = SeparatedDecompressedData[10];
      CurrentFontSize                           = Number(SeparatedDecompressedData[11]);
      CurrentCompositionElementInsertionPointX  = Number(SeparatedDecompressedData[12]);
      CurrentCompositionElementInsertionPointY  = Number(SeparatedDecompressedData[13]);
      CurrentVerticalSnap                       = Number(SeparatedDecompressedData[14]);
      CurrentHorizontalSnap                     = Number(SeparatedDecompressedData[15]);
      NumberRangeText                           = SeparatedDecompressedData[16];
      PixelsPerScaleUnitWidth                   = Number(SeparatedDecompressedData[17]);
      PixelsPerScaleUnitHeight                  = Number(SeparatedDecompressedData[18]);
      PDFImageQualityOld                        = Number(SeparatedDecompressedData[19]);
      // rebuild composition board margins, rulers, header, footer, and then composition elements
      let ContentDecompressedData               = SeparatedDecompressedData[20].split("!");
      NumberOfCompositonElements                              = ContentDecompressedData[0];
      NumberOfOpenFields                                      = parseInt(ContentDecompressedData[1],10);
      MapOfCompositionElementsToFieldNumbersAsOneString       = ContentDecompressedData[2];
      BPHeadersOfUserDataAsOneString                          = ContentDecompressedData[3];
      MapOfCompositionElementsToBPHeadersAsOneString          = ContentDecompressedData[4];
      HeaderAsString                                          = ContentDecompressedData[5];
      FooterAsString                                          = ContentDecompressedData[6];
      CurrentHeaderState                                      = ContentDecompressedData[7];
      PageNumberLayoutNumberAsString                          = ContentDecompressedData[8];

      MapOfCompositionElementsToFieldNumbers = MapOfCompositionElementsToFieldNumbersAsOneString.split(",");
      for(let k=1; k <= NumberOfCompositonElements; k++){
        MapOfCompositionElementsToFieldNumbers[k] = Number(MapOfCompositionElementsToFieldNumbers[k]);
      }
      BPHeadersOfUserData                    = BPHeadersOfUserDataAsOneString.split(",");      
      MapOfCompositionElementsToBPHeaders    = MapOfCompositionElementsToBPHeadersAsOneString.split(",");


      document.getElementById("CompositionBoardContent").innerHTML = ContentDecompressedData[9];
      // remove multiple horizontal rulers
      let rulerElements = document.getElementsByClassName('rulerH');
      while(rulerElements.length > 1){
        rulerElements[0].parentNode.removeChild(rulerElements[0]);
      }  

      SetPaperSize(PaperNumber);
      document.getElementById("ElementNumberForContent").value = NumberRangeText;
      document.getElementById("FontChoose").innerHTML = CurrentFont+" "+CurrentFontFace;
      document.getElementById("FontSize").value = CurrentFontSize;
      document.getElementById("SetVertSnap").value = CurrentVerticalSnap;
      document.getElementById("SetHorizSnap").value = CurrentHorizontalSnap;

      // bring up composition elements iff this composition isn't being used to produce an expanded data PDF or a catalog.
      if(call == -1 ){
        for( let FieldNumber=1; FieldNumber < BP_Headers.length; FieldNumber++){
          document.getElementById('CopyToCompositionBoardButton'+FieldNumber.toString()).style.opacity = "1.0";
        }
      }
      for(let k=1; k <= NumberOfCompositonElements; k++){
        let FieldNumber = MapOfCompositionElementsToFieldNumbers[k];
        if(!document.getElementById('CompositionElement'+FieldNumber.toString())){continue}
        //reactivate element dragging
        dragElement(document.getElementById('CompositionElement'+FieldNumber.toString()+'Drag'));

        //reactivate listening
        document.getElementById('CompositionElement'+FieldNumber.toString()).addEventListener('mouseup', function(){ResizeCompositionElement(FieldNumber)});
        document.getElementById('CompositionElement'+FieldNumber.toString()).oncontextmenu = function(e){e.preventDefault(); SettingsElementComposition(FieldNumber, '') };

        if(call != -1 ){
          // if we're not reloading a composition (that is, we're generating a PDF) don't show the content/description of composition elements.          
          SettingsElementComposition(FieldNumber, 'store&hide')
          // since we're generating a PDF, use user-supplied titles/header names for each composition element
          let titleText = "";
          if(FieldNumber >= 0){
              if(!/\S/.test(data[0][BP_Headers[FieldNumber]]) || !data[0][BP_Headers[FieldNumber]] ) {
                document.getElementById('CompositionElementTitle'+ FieldNumber.toString()).innerHTML = " " + BP_Headers[FieldNumber].trim().slice(3) + " ";
                // document.getElementById('CompositionElementTitle'+ FieldNumber.toString()).innerHTML = " XXXXXXXXXXXXXXxxxXXX";
                // titleText = BP_Headers[FieldNumber].trim().slice(3);
                // 
                // let titleLength = document.getElementById('CompositionElementTitle'+ FieldNumber.toString()).clientWidth;
                // document.getElementById('CompositionElementTitle'+ FieldNumber.toString()).style.left = (-titleLength).toString()+"px";
                // 
              }else{
                document.getElementById('CompositionElementTitle'+ FieldNumber.toString()).innerHTML = " " + data[0][BP_Headers[FieldNumber]].trim() + " ";
                // document.getElementById('CompositionElementTitle'+ FieldNumber.toString()).innerHTML = " XXXXXXXXXXXXXXxxxXXX";
                // titleText = data[0][BP_Headers[FieldNumber]].trim();
                // let titleLength = document.getElementById('CompositionElementTitle'+ FieldNumber.toString()).clientWidth;
                // document.getElementById('CompositionElementTitle'+ FieldNumber.toString()).style.left = (-titleLength).toString()+"px";                
              }
          }
        }else{
          if(FieldNumber >= 0){document.getElementById('CopyToCompositionBoardButton'+FieldNumber.toString()).style.opacity = "0.5"};
        }
        if(FieldNumber < 0){
          let CompositionElement =  document.getElementById('CompositionElement'+FieldNumber.toString())
          CompositionElement.addEventListener('mouseup', function(){ResizeCompositionElement(FieldNumber)});
          CompositionElement.addEventListener('mouseup', function(){IdOfCompositionElementContentWithFocus='CompositionElementContent'+FieldNumber.toString()});
          CompositionElement.oncontextmenu = function(e){e.preventDefault(); SettingsElementComposition(FieldNumber, '') };          
          continue          // if this is an open field, subsequent field processing is not applicable; we're done.
        } 

        //reset font face
        Element = document.getElementById('FontFaceElementContentButton'+FieldNumber.toString() );
        //get current header state
        CurrentFontFaceNumber = parseInt(Element.getAttribute('name'));
        //set content font face as required
        CompositionElementContent = document.getElementById('CompositionElementContent'+FieldNumber.toString() );
        let CurrentFontSpec = CompositionElementContent.style.font;
        //the font style is:   FontStyleSubstitute+" "+FontWeightSubstitute+" "+fontHeightInPX.toString()+"px "+FontFamilySubstitute;
        if(CurrentFontFaceNumber==1){
          fontFace = " normal";
        }else if(CurrentFontFaceNumber==2){
          fontFace = " italic";
        }else if(CurrentFontFaceNumber==3){
          fontFace = " bold";
        }else if(CurrentFontFaceNumber==4){
          fontFace = " italic bold";
        };
        document.getElementById('CompositionElementContent'+FieldNumber.toString()).style.font += fontFace;
        currentFrontString = document.getElementById('CompositionElementContent'+FieldNumber.toString().trim()).getAttribute('style','font')
        // gray out selected elements on the posible element list or reset open field elements, but NOT if this has been called from Expanded Data PDF request
        if(FieldNumber >= 0 ){
          if(call==""){
            if(document.getElementById('CopyToCompositionBoardButton'+FieldNumber.toString())){
              document.getElementById('CopyToCompositionBoardButton'+FieldNumber.toString()).style.opacity = "0.5"
            }
          };
        }else{
          document.getElementById('CompositionElementContent'+FieldNumber.toString()).value=document.getElementById('CompositionElementContent'+FieldNumber.toString()).name;
        }
      }

     for(let m = 1; m <= NumberOfCompositonElements; m++){
        FieldNumber = MapOfCompositionElementsToFieldNumbers[m];
        if(FieldNumber < 0){
          element = document.getElementById('CompositionElementContent'+FieldNumber.toString());
          element.setAttribute('value', element.name);
          element.visibility = "visible";
        }
      }

      //restore header/footer
      Header = JSON.parse(HeaderAsString);
      Footer = JSON.parse(FooterAsString);
      HFElement = document.getElementById('HeaderFooterButton');
      //reset current header state
      HFElement.setAttribute('name', CurrentHeaderState );
      CurrentHeaderState = parseInt(CurrentHeaderState);
      //reset image
      while (HFElement.hasChildNodes()) {
        HFElement.removeChild(HFElement.firstChild);
      }
      img = document.createElement('img');
      img.src = (CurrentHeaderState==0)? 'BP Appearance/Header0.png':
                (CurrentHeaderState==1)? 'BP Appearance/Header1.png':
                (CurrentHeaderState==2)? 'BP Appearance/Header2.png' : 'BP Appearance/Header3.png' ;
      img.alt = 'Some png';
      img.width = '25';
      img.height = '30';
      img.style.border = "none"
      HFElement.appendChild(img);
      //reactiveate resizing and contextmenu trapping
      if(Header.Present){
        HeaderElement = document.getElementById('HeaderElement');
        HeaderElement.addEventListener('mouseup', function(){ResizeHeaderFooter("HeaderElement")}); 
        HeaderElement.oncontextmenu = function(e){e.preventDefault(); SettingsHeaderFooter("HeaderSettings") };
        dragElement(document.getElementById('HeaderElementDrag'));
        if( Header.ContentLink <= 0 ){
            document.getElementById('HeaderContent').value = Header.Content;
          }
      }
      if(Footer.Present){
        FooterElement = document.getElementById('FooterElement');
        FooterElement.addEventListener('mouseup', function(){ResizeHeaderFooter("FooterElement")}); 
        FooterElement.oncontextmenu = function(e){e.preventDefault(); SettingsHeaderFooter("FooterSettings") };
        dragElement(document.getElementById('FooterElementDrag'));
        if( Footer.ContentLink <= 0 ){
          document.getElementById('FooterContent').value = Footer.Content;
        }
      }

      //restore page numbering state
      PageNumberLayoutNumber =  parseInt(PageNumberLayoutNumberAsString);
      if     (PageNumberLayoutNumber==0){Header.pageNumberInHeader=false;                                         Footer.pageNumberInFooter=false; document.getElementById("PageNumberChoose").innerHTML = "None"}
      else if(PageNumberLayoutNumber==1){Header.pageNumberInHeader=true;  Header.pageNumberPosition="TopCenter";  Footer.pageNumberInFooter=false; Footer.pageNumberPosition="";             document.getElementById("PageNumberChoose").innerHTML = "Top Center"}
      else if(PageNumberLayoutNumber==2){Header.pageNumberInHeader=true;  Header.pageNumberPosition="TopLeft";    Footer.pageNumberInFooter=false; Footer.pageNumberPosition="";             document.getElementById("PageNumberChoose").innerHTML = "Top Left"}
      else if(PageNumberLayoutNumber==3){Header.pageNumberInHeader=true;  Header.pageNumberPosition="TopRight";   Footer.pageNumberInFooter=false; Footer.pageNumberPosition="";             document.getElementById("PageNumberChoose").innerHTML = "Top Right"}
      else if(PageNumberLayoutNumber==4){Header.pageNumberInHeader=true;  Header.pageNumberPosition="TopEvenOdd"; Footer.pageNumberInFooter=false; Footer.pageNumberPosition="";             document.getElementById("PageNumberChoose").innerHTML = "Top Even Odd"}
      else if(PageNumberLayoutNumber==5){Header.pageNumberInHeader=false; Header.pageNumberPosition="";           Footer.pageNumberInFooter=true;  Footer.pageNumberPosition="BottomCenter"; document.getElementById("PageNumberChoose").innerHTML = "Bottom Center"}
      else if(PageNumberLayoutNumber==6){Header.pageNumberInHeader=false; Header.pageNumberPosition="";           Footer.pageNumberInFooter=true;  Footer.pageNumberPosition="BottomLeft";   document.getElementById("PageNumberChoose").innerHTML = "Bottom Left"}
      else if(PageNumberLayoutNumber==7){Header.pageNumberInHeader=false; Header.pageNumberPosition="";           Footer.pageNumberInFooter=true;  Footer.pageNumberPosition="BottomRight";  document.getElementById("PageNumberChoose").innerHTML = "Bottom Right"}
      else if(PageNumberLayoutNumber==8){Header.pageNumberInHeader=false; Header.pageNumberPosition="";           Footer.pageNumberInFooter=true;  Footer.pageNumberPosition="BottomEvenOdd";document.getElementById("PageNumberChoose").innerHTML = "Bottom Even Odd"}      

     // if call > 0, this has been called from the request to produce an Expanded Data PDF from an existing composition, set element number and call for the PDF to be made
     // OR if call = -1, then we're just reloading data into a compositiion board, no PDF is produced, so we're done
      if(call != ""){
        if( call > 0){
          document.getElementById("CompositionBoard").style.visibility = "visible";
          document.getElementById("CompositionBoardControls").style.visibility = "visible";
          document.getElementById("PaperMargins").style.visibility = 'visible';
          document.getElementById("FontChoose").style.visibility = 'visible';
          document.getElementById("PageNumberChoose").style.visibility = 'visible';
          CompositionRulersHaveBeenBuilt = false;
          CompositionBoardAlreadyBuilt = false;
          Header.pageNumberInHeader = false;
          Footer.pageNumberInFooter = false;
          PopulateCompositionElements(call);
          OutputFileName = NameOfLibrary;
          if(data[call]["BP_Author full name"])     {OutputFileName += "-"+data[call]["BP_Author full name"]};
          if(data[call]["BP_Short or Common Title"]){OutputFileName += "-"+data[call]["BP_Short or Common Title"]};
          if(data[call]["BP_Date Published"])       {OutputFileName += "-"+data[call]["BP_Date Published"];}
          OutputFileName += ".pdf";
          ProduceSinglePDF = true;
          // document.getElementById("OpenDropDownMenu").click();
          MakeDOC(OutputFileName);
          CloseCompositionBoard();
        }
      }else{
        // call is null, so we're building a PDF from a library from an existing composition
        // document.getElementById("CompositionBoard").style.visibility = "visible";
        // document.getElementById("CompositionBoardControls").style.visibility = "visible";
        // document.getElementById("PaperMargins").style.visibility = 'visible';
        // document.getElementById("FontChoose").style.visibility = 'visible';
        // document.getElementById("PageNumberChoose").style.visibility = 'visible';        
        OutputFileName = NameOfLibrary +".pdf";
        MakePDFofLibrary(OutputFileName);        
      };
    }
  CompressedReader.readAsText(InputFile);
  }

  function OpenBPCompositionBoard() {
    RetrieveCompositionBoard(-1)
  }  

  function ResizeCompositionElement(FieldNumber) {
    Element = document.getElementById('CompositionElement'+FieldNumber.toString());
    SnapWidthOfCompositionWindow  = CurrentHorizontalSnap*PixelsPerScaleUnitWidth; 
    SnapHeightOfCompositionWindow = CurrentVerticalSnap*PixelsPerScaleUnitHeight*PaperUnits;
    ElementWidth = Number(Element.style.width.slice(0,Element.style.width.indexOf("px")));
    ElementHeight = Number(Element.style.height.slice(0,Element.style.height.indexOf("px")));
    Element.style.width = Math.max(SnapWidthOfCompositionWindow,SnapWidthOfCompositionWindow*Math.round(ElementWidth/SnapWidthOfCompositionWindow)) + "px";
    Element.style.height  = Math.max(SnapHeightOfCompositionWindow,SnapHeightOfCompositionWindow*Math.round(ElementHeight/SnapHeightOfCompositionWindow)) + "px";
    SettingsElementComposition(FieldNumber, '');
    SettingsElementComposition(FieldNumber, '');
  }

  function PutCharInOpenField(ButtonNumbmer){
    // special characters are inserted only into opend field content elements
    if(IdOfCompositionElementContentWithFocus.includes('-')){
      // console.log("Char button clicked"+ButtonNumbmer);
      while(ButtonNumbmer > SpecialCharacters[0]){ButtonNumbmer = ButtonNumbmer-SpecialCharacters[0]}
      CharToAdd = String.fromCodePoint(SpecialCharacters[ButtonNumbmer])
      document.getElementById(IdOfCompositionElementContentWithFocus).value += CharToAdd;
    }
  }

  async function ShowMatchingheaders(){
    if(!CSVHeadersMapped){
      UserAssistance( "Navigate to folder '/BookPageant/CSV files', choose a BP-formatted .csv file and Open.",
      "CSVFile.png", 10.0);  
      try{CSVFileHandle;
        if(CSVFileHandle[0]==null){CSVFileHandle[0]==BookPageantHomeDirectoryHandle}
        CSVFileHandle = await window.showOpenFilePicker({startIn: CSVFileHandle[0], types: [{descriiption: 'BP.CSV file', accept:{'text/plain':['.BP.CSV']}}]});
      }catch{
        CSVFileHandle = await window.showOpenFilePicker({startIn: BookPageantHomeDirectoryHandle, types: [{descriiption: 'BP.CSV file', accept:{'text/plain':['.BP.CSV']}}]});
      }
      InputFile = await CSVFileHandle[0].getFile();
      localforage.setItem('CSVFileHandle', CSVFileHandle);
      Reader = new FileReader();
      Reader.onload = function(){
        dataFromCSVfile = Reader.result;
        DataHasBeenInput = false;
        Papa.parse(dataFromCSVfile, {
          delimiter: ",",
          header: true,
          download: false,
          dynamicTyping: false,
          skipEmptyLines: true,
          complete: function (results) {
            errors = results.errors;
            data = results.data;
            meta = results.meta;
            for( let m=0; m<BP_Headers.length; m++){
              if( data[0][BP_Headers[m]] != null){
                let Name = data[0][BP_Headers[m]].trim()
                if( Name.length > 25 ){while(Name.length > 25){Name = Name.slice(0,Name.lastIndexOf(" "))}}
                CurrentButtonText = document.getElementById('CopyToCompositionBoardButton'+m.toString()).innerHTML
                // strip off any previous CSV header mapping
                if( CurrentButtonText.indexOf('&nbsp') != -1){
                  CurrentButtonText = CurrentButtonText.slice(0,CurrentButtonText.indexOf('&nbsp')).trim();
                }
                document.getElementById('CopyToCompositionBoardButton'+m.toString()).innerHTML = CurrentButtonText + "&nbsp<span style='color: blue; float: right; font-size: 15px;'>"+Name+'&nbsp'+"</span>";
              }else{
                // no new CSV header data, remove any previous mapping
                CurrentButtonText = document.getElementById('CopyToCompositionBoardButton'+m.toString()).innerHTML;
                // strip off any previous CSV header mapping
                if( CurrentButtonText.indexOf('&nbsp') != -1){
                  CurrentButtonText = CurrentButtonText.slice(0,CurrentButtonText.indexOf('&nbsp'))
                }
                document.getElementById('CopyToCompositionBoardButton'+m.toString()).innerHTML = CurrentButtonText;
              }
            }          
          }
        })
      }
      Reader.readAsText(InputFile);
      CSVHeadersMapped = true;
    }else{
      for( let m=0; m<BP_Headers.length; m++){
        CurrentButtonText = document.getElementById('CopyToCompositionBoardButton'+m.toString()).innerHTML;
        // strip off any previous CSV header mapping
        if( CurrentButtonText.indexOf('&nbsp') != -1){
          CurrentButtonText = CurrentButtonText.slice(0,CurrentButtonText.indexOf('&nbsp'))
        }
        document.getElementById('CopyToCompositionBoardButton'+m.toString()).innerHTML = CurrentButtonText;
      }
      CSVHeadersMapped = false;
    }
  }
  
  function displayTextWidth(text) {
    let canvas = displayTextWidth.canvas || (displayTextWidth.canvas = document.createElement("canvas"));
    let context = canvas.getContext("2d");
    context.font = CurrentFont;
    let metrics = context.measureText(text);
    return metrics.width;
  }
//#endregion

// Creating PDF files -->
//#region

//make a PDF from 1 or more entries in an open library.
async function PDFofLibrary(){
  RetrieveCompositionBoard("")
}
function MakePDFofLibrary(OutputFileName){
  // if get here, compositiion has already been uploaded/populated
  // open control panel for generating PDFs from current library
  // document.getElementById("CompositionBoard").style.visibility = "hidden";
  document.getElementById('MakePDFfromLibrary').style.visibility = 'visible';  
  dragElement(document.getElementById("MakePDFfromLibraryDrag"));  
  document.getElementById("ElementNumberForContent").addEventListener('keyup', function(event) {
    const regex = RegExp('^[0-9\b\x1A\x1B]*$');
    if (!regex.test(event.key) && event.keyCode != 8 && event.keyCode != 37 && event.keyCode != 39 && event.keyCode != 189) {event.preventDefault()};
    NumberRangeText = document.getElementById("ElementNumberForContent").value;
  });  
  // close files & Functions dropwdown
  document.getElementById("OpenDropDownMenu").click();
  // reload previous page range
  if(CompositionPrintRange1 == null || CompositionPrintRange2 == null){
    CompositionPrintRange1 = 1;
    CompositionPrintRange2 = 2;
  }
  document.getElementById("ElementNumberForContent").value = CompositionPrintRange1.toString().trim()+"-"+CompositionPrintRange2.toString().trim()
  
  if(PDFImageQuality !=null){
    document.getElementById("QualityForContent").value = PDFImageQuality;
  }else{
    PDFImageQuality = 0.15;
  }
  // set the image number (if any are to be included) to 1.
  CurrentImageDisplayedInExpandedData = 1;  
}
function CloseMakePDFofLibrary(){
  document.getElementById('MakePDFfromLibrary').style.visibility = 'hidden';    
  CloseCompositionBoard();
}

async function MakeDOC(OutputFileName = NameOfLibrary) {
  document.getElementById("OpenDropDownMenu").click();
  var PDFdoc = new jsPDF({orientation: 'p', unit: ((PaperUnits==1)? 'in':'mm'), format: PaperSizeName[PaperNumber].toLowerCase(), });
  let fontlist = PDFdoc.getFontList()
  let NumberOfLinesJustWritten = 0;
  let PageNumber =1;
  let VerticalSectionSpacing = 0.25 // in doc units.
  let LocalLineHeightFactor = PDFdoc.getLineHeightFactor(); // in type points.
  let PointConversion = PDFdoc.internal.scaleFactor // number of points in the document units.
  let fontName = CurrentFont;
  let fontStyle = CurrentFontFace;
  let fontSize = CurrentFontSize; // in type points = 1/72.0"
  PDFpages[0] = 1
  PDFpages[1] = 0;
  let lastPageWritten = PDFpages[0];
  CutLines = [];
  CutLines[0] = 0;
  var BottomPageLimit = PaperSizeLength[PaperNumber] - PageMargins[4];
  document.getElementById("OpenDropDownMenu").click();
  placeHeaderFooterOnPage(Header, Footer, PDFdoc, CompositionPrintRange1);
  LineHeightUnits = LocalLineHeightFactor*fontSize/PointConversion;
  yCurrent = PageMargins[3]+LineHeightUnits*0.5;

  document.getElementById("MakePDFProgress").setAttribute( 'style', 'display: block');
  circleProgress = new CircleProgress('.MakePDFProgress',{constrain: true, textFormat: 'percent' });
  await sleep(20);  
  circleProgress.max = 100;
  circleProgress.value = 0;

 for( let DataElementNumber = CompositionPrintRange1;  DataElementNumber <= CompositionPrintRange2; DataElementNumber++ ){
    //update content of composition elements
    circleProgress.value = 100*DataElementNumber/(CompositionPrintRange2-CompositionPrintRange1+1);
    CompositionElementLoop:
    for(let m = 1; m <= NumberOfCompositonElements; m++){
      FieldNumber = MapOfCompositionElementsToFieldNumbers[m];
      var element = document.getElementById('CompositionElementContent'+FieldNumber.toString());        
      if(FieldNumber >= 0){
        HeaderToMatch = MapOfCompositionElementsToBPHeaders[m];
        if(element){
          // get current header state
          headerElement = document.getElementById('HeaderElementContentButton'+FieldNumber.toString());
          if(headerElement){
            let CurrentHeaderState = parseInt(headerElement.getAttribute('name'));
            if(data[0][HeaderToMatch]){
              if     (CurrentHeaderState==1){HeaderText = data[0][HeaderToMatch].trim() + ":  "}
              else if(CurrentHeaderState==2){HeaderText = data[0][HeaderToMatch].trim().toUpperCase() + ":  "}
              else if(CurrentHeaderState==3){HeaderText = ""}
              else if(CurrentHeaderState==4){HeaderText = ""}
              else if(CurrentHeaderState==5){HeaderText = ""}
              element.innerHTML = HeaderText + data[DataElementNumber][HeaderToMatch];
            }else{
              if     (CurrentHeaderState==1){HeaderText = BP_Headers[FieldNumber].trim().slice(3) + ":  "}
              else if(CurrentHeaderState==2){HeaderText = BP_Headers[FieldNumber].trim().slice(3).toUpperCase() + ":  "}
              else if(CurrentHeaderState==3){HeaderText = ""}
              else if(CurrentHeaderState==4){HeaderText = ""}
              else if(CurrentHeaderState==5){HeaderText = ""}
              element.innerHTML = HeaderText;            
            }
          }
          continue CompositionElementLoop;
        }
      }
    }

   CompElementLoop:
    for(var CurrentElementNumber = 1; CurrentElementNumber <= NumberOfCompositonElements; CurrentElementNumber++){
      //if a new page has been triggered, process any pending composition elements
      if(CompositionElementWaiting && NewPageTriggered){
        FieldNumber = MapOfCompositionElementsToFieldNumbers[ElementPending];
        CompositionElementWaiting = false;
        NewPageTriggered = false;
        ElementPending = 0;
        CurrentElementNumber--;
      }else{
        FieldNumber = MapOfCompositionElementsToFieldNumbers[CurrentElementNumber];
      }
       //if the section of element is empty, skip.
      if(!document.getElementById('CompositionElement'+FieldNumber.toString())){continue};
      // if data is null or blank, skip.
      if (!/\S/.test(data[DataElementNumber][BP_Headers[FieldNumber]])){continue};
      // if the current composition component is empty, skip it. Unless shortened title, even if user data does not contain it: use trunccated full title in this case.
      if(data[DataElementNumber][BP_Headers[FieldNumber]]){
        if(BP_Headers[FieldNumber]=="BP_Short or Common Title"){
          if(data[DataElementNumber][BP_Headers[FieldNumber]].trim() == ""){
            let TempShortTitle = data[DataElementNumber]["BP_Title"];
            TempShortTitle = TempShortTitle.slice(0,41);
            TempShortTitle = TempShortTitle.slice(0,TempShortTitle.lastIndexOf(" "))+" ...";
            document.getElementById('CompositionElement'+FieldNumber.toString()).innerHTML = TempShortTitle
          }
        }else{
          if(data[DataElementNumber][BP_Headers[FieldNumber]]==""){continue CompElementLoop}      
        }
      }else{
        // always provide data for shortened title, even if user data does not contain it: use trunccated full title in this case.
        if(BP_Headers[FieldNumber]=="BP_Short or Common Title"){
          let TempShortTitle = data[DataElementNumber]["BP_Title"];
          TempShortTitle = TempShortTitle.slice(0,41);
          TempShortTitle = TempShortTitle.slice(0,TempShortTitle.lastIndexOf(" "))+" ...";
          document.getElementById('CompositionElementContent'+FieldNumber.toString()).innerHTML += TempShortTitle;   
        }else{
          // skip current composition element if it isn't an image
          if(FieldNumber >= 0){continue};
        }
      }
      // skip edition composition element if user data isn't > 0
      if(BP_Headers[FieldNumber]=="BP_Edition Number"){
        if(data[DataElementNumber][BP_Headers[FieldNumber]] < 1 || isNaN(data[DataElementNumber][BP_Headers[FieldNumber]]) || isNaN(parseFloat(data[DataElementNumber][BP_Headers[FieldNumber]]) )){continue CompElementLoop};
      }

     //if we're at the start of one of the elements to be printed, for a new page if we're near the bottom of the current page
      if(CurrentElementNumber==1){
        NumberOfLines = 4;
        for( m=1; m<=Math.min(5,NumberOfCompositonElements); m++){
          if(document.getElementById('CompositionElementTitle'+MapOfCompositionElementsToFieldNumbers[m].toString() )){
            if(document.getElementById('CompositionElementTitle'+MapOfCompositionElementsToFieldNumbers[m].toString() ).innerHTML.trim() == "Image local Link1"){
            NumberOfLines = 20;
            break
            }
          }
        }
        if(yCurrent >= BottomPageLimit - NumberOfLines*LineHeightUnits){
          //new page. reset all parameters
          PDFdoc.addPage(PaperSizeName[PaperNumber].toLowerCase(),'p')
          yAdd = 0;
          //this is to the upper margin of the next page; that is, the y position is to TOP of first line of text. But jsPDF generator assumes the BOTTOM of the first line of text. Adjust
          yCurrent = PageMargins[3]+LineHeightUnits*0.5;
          PDFpages[0]++;
          PDFpages[PDFpages[0]] = 0;
          PDFdoc.setPage(PDFpages[0]);
          placeHeaderFooterOnPage(Header, Footer, PDFdoc, DataElementNumber)
          CutLines[0]=0;
          NewPageTriggered = true;
          // header or footer may contain a data field and the call to placeHeaderFooterOnPage may change FieldNumber. So, reset it
          FieldNumber = MapOfCompositionElementsToFieldNumbers[CurrentElementNumber];          
        }
      }
      
      // if(document.getElementById('CompositionElement'+MapOfCompositionElementsToFieldNumbersCurrentToOriginal[FieldNumber].toString())){ElementToPrint = document.getElementById('CompositionElement'+MapOfCompositionElementsToFieldNumbersCurrentToOriginal[FieldNumber].toString() )};
      // if(document.getElementById('CompositionElementTitle'+MapOfCompositionElementsToFieldNumbersCurrentToOriginal[FieldNumber].toString())){ElementTitle = document.getElementById('CompositionElementTitle'+MapOfCompositionElementsToFieldNumbersCurrentToOriginal[FieldNumber].toString() ).innerHTML};
      if(document.getElementById('CompositionElement'+FieldNumber.toString())){var ElementToPrint = document.getElementById('CompositionElement'+FieldNumber.toString() )};      
      if(document.getElementById('CompositionElementTitle'+FieldNumber.toString())){var ElementTitle = document.getElementById('CompositionElementTitle'+FieldNumber.toString() ).innerHTML};

      if(ElementTitle.trim().includes("image") || ElementTitle.trim().includes("Image")){
        if(!ImageCompressionCollection[DataElementNumber][1]){continue CompElementLoop}
        if(CurrentImageDisplayedInExpandedData==0){continue CompElementLoop}
        //element contains an image
        // FullImageName = "https://cors-anywhere.herokuapp.com/" + data[DataElementNumber][BPHeadersOfUserData[FieldNumber]];
        if(CurrentImageDisplayedInExpandedData==1){
          FullImageName = data[DataElementNumber][BP_Headers[FieldNumber]].slice(data[DataElementNumber][BP_Headers[FieldNumber]].lastIndexOf("/")+1).trim()
        }else if(CurrentImageDisplayedInExpandedData==2){
          FullImageName = data[DataElementNumber][BP_Headers[FieldNumber+1]].slice(data[DataElementNumber][BP_Headers[FieldNumber+1]].lastIndexOf("/")+1).trim()
        }else if(CurrentImageDisplayedInExpandedData==3){
          FullImageName = data[DataElementNumber][BP_Headers[FieldNumber+2]].slice(data[DataElementNumber][BP_Headers[FieldNumber+2]].lastIndexOf("/")+1).trim()
        }else if(CurrentImageDisplayedInExpandedData==4){
          FullImageName = data[DataElementNumber][BP_Headers[FieldNumber+3]].slice(data[DataElementNumber][BP_Headers[FieldNumber+3]].lastIndexOf("/")+1).trim()          
        }
        xCurrent = PaperWidth *ElementToPrint.offsetLeft/ElementToPrint.parentElement.clientWidth;
        LineHeightUnits = LocalLineHeightFactor*fontSize/PointConversion;
        sectionWidth  = ElementToPrint.clientWidth/PixelsPerScaleUnitWidth;
        // unlike other sections of the compositon (with heights determined by content), 
        // images have a maximum height determined by the image element established by the user, 
        // regardless of image size or ascpect ratio.
        sectionHeight = ElementToPrint.clientHeight/PixelsPerScaleUnitHeight;
        scalingFactor = (ElementToPrint.offsetTop /ElementToPrint.parentElement.clientHeight)*(ElementToPrint.parentNode.clientHeight/ElementToPrint.parentNode.clientWidth);
        HoldNumberOfCompositonElements = NumberOfCompositonElements;

        ImageLoaded = new Promise(resolve => {
            img = new Image();
            img.onload = () => resolve() 
        })
        img.src = ImageCompressionCollection[DataElementNumber][1]      //image source is from the compressed image collection array. Assume image number 1
        await ImageLoaded // wait here until image is loaded from file and its properties are available.
        
        ImageComp = new Promise(resolve => {
          imgComp = new Image();
          Canvas = document.createElement('canvas');
          Canvas.width = img.width;
          Canvas.height = img.height;
          ctx = Canvas.getContext('2d');
          ctx.drawImage(img, 0, 0, img.width, img.height);
          imgComp.onload = () => resolve()
          imgComp.src = Canvas.toDataURL("image/jpeg", Number(PDFImageQuality))
        })
        await ImageComp //wait here until a processed/compressed image has been produced.

        NumberOfCompositonElements = HoldNumberOfCompositonElements;
        iWidth = imgComp.width;
        iHeight = imgComp.height;
        iAspectRatio = iWidth/iHeight;
        // see if image fits into image section of composition
        sectionAspectRatio = sectionWidth/sectionHeight
        if(iAspectRatio > sectionAspectRatio){
          // image is too wide to fit in section without distortion. change section height so that its aspect ratio matches that of the image
          sectionHeight = sectionWidth/iAspectRatio;
        }else{
          sectionWidth = sectionHeight*iAspectRatio;
        }
        // sectionHeight = sectionWidth/iAspectRatio;
        // if(CurrentElementNumber==1 && DataElementNumber == CompositionPrintRange1){
          // yCurrent = LineHeightUnits*0.5+PaperWidth*(ElementToPrint.offsetTop /ElementToPrint.parentElement.clientHeight)*(ElementToPrint.parentNode.clientHeight/ElementToPrint.parentNode.clientWidth);
          // yCurrent = LineHeightUnits*0.5+PaperWidth*scalingFactor;
          // console.log('a', DataElementNumber, xCurrent, yCurrent)
        // }else{
          if(CutLines[0]==0){
            yCurrent = PDFpages[PDFpages[0]];
            // yCurrent = LineHeightUnits*0.5+PaperWidth*(ElementToPrint.offsetTop /ElementToPrint.parentElement.clientHeight)*(ElementToPrint.parentNode.clientHeight/ElementToPrint.parentNode.clientWidth);
            yCurrent = LineHeightUnits*0.5+PaperWidth*scalingFactor;
            // console.log('b', DataElementNumber, xCurrent, yCurrent)
          }else{
            // yCurrent = PaperWidth*(ElementToPrint.offsetTop /ElementToPrint.parentElement.clientHeight)*(ElementToPrint.parentNode.clientHeight/ElementToPrint.parentNode.clientWidth);
            yCurrent = PaperWidth*scalingFactor;
            LargestCutLinesY = PageMargins[3]+LineHeightUnits*0.5;
            for( p=1; p<=CutLines[0]; p++){
              if( CutLines[p][1] < xCurrent + sectionWidth - PagePositionTolerance && 
                CutLines[p][2] > xCurrent + PagePositionTolerance ){
              if(CutLines[p][0] > LargestCutLinesY){LargestCutLinesY = CutLines[p][0]};
              } 
            }
            yCurrent = LargestCutLinesY;
            // console.log('c', DataElementNumber, xCurrent, yCurrent)
          }
          // don't start a new section if we're < what's needed for the image (within a tolerance of 1/2 line height)
          if(yCurrent > BottomPageLimit - sectionHeight + LineHeightUnits*0.5){
            //new page. reset all parameters
            PDFdoc.addPage(PaperSizeName[PaperNumber].toLowerCase(),'p')
            yAdd = 0;
            //this is to the upper margin of the next page; that is, the y position is to TOP of first line of text. But jsPDF generator assumes the BOTTOM of the first line of text. Adjust
            yCurrent = PageMargins[3]+LineHeightUnits*0.5;
            // console.log('d', DataElementNumber, xCurrent, yCurrent)
            PDFpages[0]++;
            PDFpages[PDFpages[0]] = 0;
            PDFdoc.setPage(PDFpages[0]);
            placeHeaderFooterOnPage(Header, Footer, PDFdoc, DataElementNumber)
            CutLines[0]=0;
            NewPageTriggered = true;
          }
        // }
        // console.log(CutLines[0], DataElementNumber, xCurrent, yCurrent, sectionWidth, sectionHeight)
        PDFdoc.addImage(imgComp, 'JPEG', xCurrent, yCurrent , sectionWidth, sectionHeight);    //, alias, compression, rotation)
        yCurrent += sectionHeight;
        if(yCurrent > PDFpages[PDFpages[0]]){PDFpages[PDFpages[0]] = yCurrent};
        CutLines[0]++;
        CutLines[CutLines[0]] = [yCurrent+LineHeightUnits, xCurrent, xCurrent+sectionWidth];

      }else{
        //element contains text
        let LineHeightUnits = LocalLineHeightFactor*fontSize/PointConversion;
        if(FieldNumber >= 0){
          if(document.getElementById('CompositionElementContent'+FieldNumber.toString())){
            ContentToPrint = document.getElementById('CompositionElementContent'+FieldNumber.toString());          
            HeaderToPlace = document.getElementById('CompositionElementContentHeader'+FieldNumber.toString() ).innerHTML;
                 if(HeaderToPlace=="HeaderInLine")    {HeaderToPlace = data[0][BP_Headers[FieldNumber]].trim()+': '}
            else if(HeaderToPlace=="HeaderInLineCaps"){HeaderToPlace = data[0][BP_Headers[FieldNumber]].trim().toUpperCase()+': '}
            else if(HeaderToPlace=="HeaderAbove")     {HeaderToPlace = data[0][BP_Headers[FieldNumber]].trim()}
            else if(HeaderToPlace=="HeaderAboveCaps") {HeaderToPlace = data[0][BP_Headers[FieldNumber]].trim()}
            else if(HeaderToPlace=="NoHeader"){HeaderToPlace=''}

            TextToPlace = ContentToPrint.innerHTML;
            HeaderPresent = false;
            if(HeaderToPlace != ""){
              HeaderPresent = true;
              TextToPlace = HeaderToPlace + "<br>" + TextToPlace}
            textAlign = ContentToPrint.style.textAlign;
            NumberOfColumns = Number(ContentToPrint.style.columnCount);
          }
        }else{
          // this is an Open Field
          HeaderPresent = false;
          ContentToPrint = document.getElementById('CompositionElementContent'+FieldNumber.toString());
          // value in open field text field was transfered from .value to .name
          TextToPlace = ContentToPrint.value;
          // if the field contains new page character U+2398, trigger a new page and continue the composition element loop
          if(TextToPlace.includes("⎘")){
            // new page. reset all parameters
            PDFdoc.addPage(PaperSizeName[PaperNumber].toLowerCase(),'p')
            yAdd = 0;
            // this is to the upper margin of the next page; that is, the y position is to TOP of first line of text. But jsPDF generator assumes the BOTTOM of the first line of text. Adjust
            yCurrent = PageMargins[3]+LineHeightUnits*0.5;
            PDFpages[0]++;
            PDFpages[PDFpages[0]] = 0;
            PDFdoc.setPage(PDFpages[0]);
            placeHeaderFooterOnPage(Header, Footer, PDFdoc, DataElementNumber)
            CutLines[0]=0;
            NewPageTriggered = true;
            continue CompElementLoop;
          }
          // if the field has only calls for new line(s) or paragraph AND we're at the page top, skip
          if(yCurrent <= PageMargins[3]+LineHeightUnits){
            while(TextToPlace.slice(0,1)=="¶" && TextToPlace.length > 0){
              TextToPlace = TextToPlace.slice(1);
            }
            while(TextToPlace.slice(0,4)=="<br>" && TextToPlace.length > 0){
              TextToPlace = TextToPlace.slice(4);
            }
            while(TextToPlace.slice(0,3)=="\\n" && TextToPlace.length > 0){
              TextToPlace = TextToPlace.slice(3);
            }
            if(TextToPlace.length == 0){continue CompElementLoop}
          }
            NumberOfColumns = 1;
          textAlign = ContentToPrint.style.textAlign;
        }
        let LocalFontDescription = ContentToPrint.style.font;
        if( (LocalFontDescription.includes("italic") || LocalFontDescription.includes("Italic")) && (LocalFontDescription.includes("bold") || LocalFontDescription.includes("Bold")) ){fontStyle = "bolditalic"}
        else if( (LocalFontDescription.includes("italic") || LocalFontDescription.includes("Italic")) ){fontStyle = "italic"}
        else if( (LocalFontDescription.includes("bold") || LocalFontDescription.includes("Bold")) ){fontStyle = "bold"}
        else{ fontStyle = "normal"}
        FontChangeToggle[0] = 0;
        let ColumnSpacing = 0.125;
        let xCurrent = PaperWidth *ElementToPrint.offsetLeft/ElementToPrint.parentElement.clientWidth;
        let sectionWidth = PaperWidth*ElementToPrint.clientWidth/ElementToPrint.parentElement.clientWidth;
        //given y position is to TOP of first line of text. PDF generator assumes the BOTTOM of the first line of text. Adjust
        // if(CurrentElementNumber==1 && DataElementNumber == CompositionPrintRange1){
          // yCurrent = LineHeightUnits*0.5+PaperWidth*(ElementToPrint.offsetTop /ElementToPrint.parentElement.clientHeight)*(ElementToPrint.parentNode.clientHeight/ElementToPrint.parentNode.clientWidth)
        // }else{
          if(CutLines[0]==0){
            // yCurrent = PDFpages[PDFpages[0]];
            yCurrent = PageMargins[3]+LineHeightUnits*0.5;
          }else{
            yCurrent = PaperWidth*(ElementToPrint.offsetTop /ElementToPrint.parentElement.clientHeight)*(ElementToPrint.parentNode.clientHeight/ElementToPrint.parentNode.clientWidth);
            LargestCutLinesY = PageMargins[3]+LineHeightUnits*0.5;
            for( p=1; p<=CutLines[0]; p++){
              if( CutLines[p][1] < xCurrent + sectionWidth - PagePositionTolerance && 
                  CutLines[p][2] > xCurrent + PagePositionTolerance ){
                if(CutLines[p][0] > LargestCutLinesY){LargestCutLinesY = CutLines[p][0]};
              }
            }
            yCurrent = LargestCutLinesY;
          }
          //don't start a new section if we're <= 2 lines from the bottom
          if(yCurrent >= BottomPageLimit - 2*LineHeightUnits){
            //before moving to new page, look ahead to see if next compositiion element would fit on current page.
            if( CurrentElementNumber < NumberOfCompositonElements){
              let NextFieldNumber = MapOfCompositionElementsToFieldNumbers[CurrentElementNumber+1];
              let NextElementToPrint = document.getElementById('CompositionElement'+NextFieldNumber.toString() );
              let NextElementTitle = document.getElementById('CompositionElementTitle'+NextFieldNumber.toString() ).innerHTML;
              let NextxCurrent = PaperWidth *NextElementToPrint.offsetLeft/NextElementToPrint.parentElement.clientWidth;
              let NextsectionWidth = PaperWidth*NextElementToPrint.clientWidth/NextElementToPrint.parentElement.clientWidth;
              NextyCurrent = PaperWidth*(NextElementToPrint.offsetTop /NextElementToPrint.parentElement.clientHeight)*(NextElementToPrint.parentNode.clientHeight/NextElementToPrint.parentNode.clientWidth);
              NextLargestCutLinesY = PageMargins[3]+LineHeightUnits*0.5;
              for( p=1; p<=CutLines[0]; p++){
                if( CutLines[p][1] < NextxCurrent + NextsectionWidth - PagePositionTolerance && 
                    CutLines[p][2] > NextxCurrent + PagePositionTolerance ){
                  if(CutLines[p][0] > NextLargestCutLinesY){NextLargestCutLinesY = CutLines[p][0]};
                }
              }
              NextyCurrent = NextLargestCutLinesY;
              if(NextyCurrent < BottomPageLimit - 2*LineHeightUnits){
                //there is room for the m+1 composition element. hold current element until next page is triggered.
                CompositionElementWaiting = true;
                NewPageTriggered = false;
                ElementPending = CurrentElementNumber;
                //move to next element
                continue CompElementLoop;
              }
            }
            //new page. reset all parameters
            PDFdoc.addPage(PaperSizeName[PaperNumber].toLowerCase(),'p')
            yAdd = 0;
            //this is to the upper margin of the next page; that is, the y position is to TOP of first line of text. But jsPDF generator assumes the BOTTOM of the first line of text. Adjust
            yCurrent = PageMargins[3]+LineHeightUnits*0.5;
            PDFpages[0]++;
            PDFpages[PDFpages[0]] = 0;
            PDFdoc.setPage(PDFpages[0]); 
            placeHeaderFooterOnPage(Header, Footer, PDFdoc, DataElementNumber)
            CutLines[0] = 0;
            CutLines.length = 1;
            NewPageTriggered = true;
          }
        // }
        //don't start new page with new line, new para, or break
        if(yCurrent <= PageMargins[3]+LineHeightUnits){
          while(TextToPlace.slice(0,1)=="¶" && TextToPlace.length > 0){
            TextToPlace = TextToPlace.slice(1);
          }
          while(TextToPlace.slice(0,4)=="<br>" && TextToPlace.length > 0){
            TextToPlace = TextToPlace.slice(4);
          }
          while(TextToPlace.slice(0,3)=="\\n" && TextToPlace.length > 0){
            TextToPlace = TextToPlace.slice(3);
          }
          if(TextToPlace.length == 0){continue CompElementLoop}
        }

        //increment line down the page to account for user specified spacing iff we're not at the top of the page.
        if(yCurrent > PageMargins[3]+LineHeightUnits){
          // yCurrent = yCurrent + Number(document.getElementById('AddSpaceElementContentButton'+MapOfCompositionElementsToFieldNumbersCurrentToOriginal[FieldNumber].toString() ).getAttribute('name'))*LineHeightUnits*0.5;
          yCurrent = yCurrent + Number(document.getElementById('AddSpaceElementContentButton'+FieldNumber.toString() ).getAttribute('name'))*LineHeightUnits*0.5;
        }

        //insert CRLF or NL in the text, and decode any ampersands and other special characters
        TextToPlace = TextToPlace.replace(/<br><br>/g,"{LF}{NL}");
        TextToPlace = TextToPlace.replace(/<br>/g,"{NL}");
        TextToPlace = TextToPlace.replace(/\\n\\n/g,"{LF}{NL}");
        TextToPlace = TextToPlace.replace(/\\n/g,"{NL}");
        TextToPlace = TextToPlace.replace(/&amp;/g,'&');
        TextToPlace = TextToPlace.replace(/&gt;/g,'>');
        TextToPlace = TextToPlace.replace(/&lt;/g,'<');
        TextToPlace = TextToPlace.replace(/&quot;/g,'"');
        TextToPlace = TextToPlace.replace(/&#39;/g,"'");
        TextToPlace = TextToPlace.replace(/\(/g,"\(");
        TextToPlace = TextToPlace.replace(/\)/g,"\)");
        // inline bold, underlining, superscript, subscript font faces don't/cannot propagate to PDF. remove their tags
        TextToPlace = TextToPlace.replace(/<b>/g,"");
        TextToPlace = TextToPlace.replace(/<u>/g,"");
        TextToPlace = TextToPlace.replace(/<sup>/g,"");
        TextToPlace = TextToPlace.replace(/<sub>/g,"");
        TextToPlace = TextToPlace.replace(/<\/b>/g,"");
        TextToPlace = TextToPlace.replace(/<\/u>/g,"");
        TextToPlace = TextToPlace.replace(/<\/sup>/g,"");
        TextToPlace = TextToPlace.replace(/<\/sub>/g,"");
        //convert emphasis HTML tag to italic and closing tags to same a opening tags. processing with use "<i>" as on/off toggle for italic font face.
        TextToPlace = TextToPlace.replace(/<em>/g,"<i>");
        TextToPlace = TextToPlace.replace(/<\/em>/g,"<i>");
        TextToPlace = TextToPlace.replace(/<\/i>/g,"<i>");

        // if this is a collation composition element, it is necessary to override current font & style and use special 'collation' font which contains all necessary special characters
        if(BP_Headers[FieldNumber]=="BP_Collation/Pagination/Foliation"){
          if( TextToPlace.includes("Vol") ){
            MultipleVolumeCollation = true;
            NumberOfCollations = (TextToPlace.match(/Vol/g) || []).length;
            CollationTextStartIndex = TextToPlace.indexOf("Vol",0);
          }else if( TextToPlace.includes("vol")){
            NumberOfCollations = (TextToPlace.match(/vol/g) || []).length;
            CollationTextStartIndex = TextToPlace.indexOf("vol",0);
            MultipleVolumeCollation = true;
          }else{
            MultipleVolumeCollation = false;
            NumberOfCollations = 1;
            CollationTextStartIndex = 0;
          }
          // collation, pagination, foliation, and plates for each volume
          for(let VolNumber=1; VolNumber <= NumberOfCollations; VolNumber++ ){
            // write a small horizontal delimeter between sections for multiple volumes if there are more than one.
            if(VolNumber > 1){
              Delimiter = String.fromCharCode(8766,8766,8766,8766);
              placeTextOnPage(PDFdoc, xCurrent, yCurrent, sectionWidth, BottomPageLimit, NumberOfColumns, ColumnSpacing, HeaderPresent, Delimiter, textAlign, 'Collation', fontStyle, fontSize, LineHeightUnits, CutLines, DataElementNumber);
              yCurrent += LineHeightUnits
            }
            // collation perhaps mixed with pagination, foliation, and plates.
            // for each volume parse the text to be processed for collation, pagination, foliation, and plates
            CollationTextEndIndex = TextToPlace.indexOf("Vol",CollationTextStartIndex+6);
            if(CollationTextEndIndex == -1){CollationTextEndIndex = TextToPlace.indexOf("vol",CollationTextStartIndex+6)};
            if(CollationTextEndIndex > 0 ){
              TextForProcessing = TextToPlace.slice(CollationTextStartIndex,CollationTextEndIndex)
            }else{
              TextForProcessing = TextToPlace.slice(CollationTextStartIndex)
            }
            CollationTextStartIndex = CollationTextEndIndex;
            // collation 
            if(TextForProcessing.includes("Collation:")){
              IndexFirstCollationChar = TextForProcessing.indexOf("Collation:") + 10;
            }else if(TextForProcessing.includes("collation:")){
              IndexFirstCollationChar = TextForProcessing.indexOf("collation:") + 10;
            }else if(TextForProcessing.includes("Collation")){
              IndexFirstCollationChar = TextForProcessing.indexOf("Collation") + 9;
            }else if(TextForProcessing.includes("collation")){
              IndexFirstCollationChar = TextForProcessing.indexOf("collation") + 9;
            }else{
              IndexFirstCollationChar = 0;
            }
            // pagination
            if(TextForProcessing.includes("Pagination")){
              IndexLastCollationChar = TextForProcessing.indexOf("Pagination");
            }else if(TextForProcessing.includes("pagination")){
              IndexLastCollationChar = TextForProcessing.indexOf("pagination");
            }else if(TextForProcessing.includes("Unpaginated")){
              IndexLastCollationChar = TextForProcessing.indexOf("Unpaginated");
            }else if(TextForProcessing.includes("Not Paginated")){
              IndexLastCollationChar = TextForProcessing.indexOf("Not Paginated");
            }else if(TextForProcessing.includes("Not paginated")){
              IndexLastCollationChar = TextForProcessing.indexOf("Not paginated");
            }else if(TextForProcessing.includes("No Pagination")){
              IndexLastCollationChar = TextForProcessing.indexOf("No Pagination");
            }else if(TextForProcessing.includes("No pagination")){
              IndexLastCollationChar = TextForProcessing.indexOf("No pagination");
            }else{
              IndexLastCollationChar = TextForProcessing.length
            }
            // foliation
            if(TextForProcessing.includes("Foliation")){
              if(TextForProcessing.indexOf("Foliation")<IndexLastCollationChar){IndexLastCollationChar = TextForProcessing.indexOf("Foliation")};
            }else if(TextForProcessing.includes("foliation")){
              if(TextForProcessing.indexOf("foliation")<IndexLastCollationChar){IndexLastCollationChar = TextForProcessing.indexOf("foliation")};
            }else if(TextForProcessing.includes("Unfoliated")){
              if(TextForProcessing.indexOf("Unfoliated")<IndexLastCollationChar){IndexLastCollationChar = TextForProcessing.indexOf("Unfoliated")};
            }else if(TextForProcessing.includes("Not Foliated")){
              if(TextForProcessing.indexOf("Not Foliated")<IndexLastCollationChar){IndexLastCollationChar = TextForProcessing.indexOf("Not Foliated")};
            }else if(TextForProcessing.includes("No Foliation")){
              if(TextForProcessing.indexOf("No Foliation")<IndexLastCollationChar){IndexLastCollationChar = TextForProcessing.indexOf("No Foliation")};
            }
            // plates or notes
            if(TextForProcessing.includes("Plates")){
              if(TextForProcessing.indexOf("Plates")<IndexLastCollationChar){IndexLastCollationChar = TextForProcessing.indexOf("Plates")};
            }else if(TextForProcessing.includes("plates")){
              if(TextForProcessing.indexOf("plates")<IndexLastCollationChar){IndexLastCollationChar = TextForProcessing.indexOf("plates")};
            }else if(TextForProcessing.includes("Notes")){
              if(TextForProcessing.indexOf("Notes")<IndexLastCollationChar){IndexLastCollationChar = TextForProcessing.indexOf("Notes")};
            }else if(TextForProcessing.includes("notes")){
              if(TextForProcessing.indexOf("notes")<IndexLastCollationChar){IndexLastCollationChar = TextForProcessing.indexOf("notes")};
            }

            if(IndexFirstCollationChar>0){
              // any text before the collation characters
              PreText = TextForProcessing.substring(0,IndexFirstCollationChar)
              placeTextOnPage(PDFdoc, xCurrent, yCurrent, sectionWidth, BottomPageLimit, NumberOfColumns, ColumnSpacing, HeaderPresent, PreText, textAlign, fontName, fontStyle, fontSize, LineHeightUnits, CutLines, DataElementNumber);
              WidthOfPreText = PDFdoc.getStringUnitWidth(PreText+' ')*fontSize/72.0;  // fontSize is points. /72 converts to inches
            }else{
              PreText = '';
              WidthOfPreText = 0;
            }
            // collation characters
            CollationText = TextForProcessing.substring(IndexFirstCollationChar, IndexLastCollationChar)
            // need to pad CollationText to leave space for PreText that has already been placed on the page. length depends on PreText size and font family used for collationg PDF text
            CollationfontName = 'Collation';
            CollationfontStyle = 'normal';
            PDFdoc.setFont(CollationfontName, CollationfontStyle);
            PDFdoc.setFontSize(fontSize);
            Padding = String.fromCharCode(0); // 1st char of padding is null character -- prevents trimming by placeTextOnPage routine.
            PaddingWidth = PDFdoc.getStringUnitWidth(Padding)*fontSize/72.0;
            while (PaddingWidth < WidthOfPreText){
              Padding += String.fromCharCode(32);
              PaddingWidth = PDFdoc.getStringUnitWidth(Padding)*fontSize/72.0;
            }
            CollationText = Padding+CollationText;
            placeTextOnPage(PDFdoc, xCurrent, yCurrent, sectionWidth, BottomPageLimit, NumberOfColumns, ColumnSpacing, HeaderPresent, CollationText, textAlign, CollationfontName, CollationfontStyle, fontSize, LineHeightUnits, CutLines, DataElementNumber);
            // for(let j=0; j<=CollationText.length; j++){
            //   let UnicodePoint = CollationText.codePointAt(j)
            //   console.log(CollationText.charAt(j), UnicodePoint)
            // }
            // return to user-specified font and style
            PDFdoc.setFont(fontName, fontStyle);
            PDFdoc.setFontSize(fontSize);
            PostText = TextForProcessing.substring(IndexLastCollationChar)
            // determine yposition
            LargestCutLinesY = yCurrent //- LineHeightUnits;
            for( p=1; p<=CutLines[0]; p++){
              if( CutLines[p][1] < xCurrent + sectionWidth - PagePositionTolerance && 
                  CutLines[p][2] > xCurrent + PagePositionTolerance ){
                if(CutLines[p][0] > LargestCutLinesY){LargestCutLinesY = CutLines[p][0]};
              }
            }
            yCurrent = LargestCutLinesY;
            // remaining text
            if(CollationText.trim().substring(CollationText.trim().length-4) == '{NL}'){yCurrent = yCurrent - LineHeightUnits };
            placeTextOnPage(PDFdoc, xCurrent, yCurrent, sectionWidth, BottomPageLimit, NumberOfColumns, ColumnSpacing, HeaderPresent, PostText, textAlign, fontName, fontStyle, fontSize, LineHeightUnits, CutLines, DataElementNumber);
            // move to new y position on page
            LargestCutLinesY = yCurrent //- LineHeightUnits;
            for( p=1; p<=CutLines[0]; p++){
              if( CutLines[p][1] < xCurrent + sectionWidth - PagePositionTolerance && 
                  CutLines[p][2] > xCurrent + PagePositionTolerance ){
                if(CutLines[p][0] > LargestCutLinesY){LargestCutLinesY = CutLines[p][0]};
              }
            }
            yCurrent = LargestCutLinesY;
            if(PostText.trim().substring(PostText.trim().length-4) == '{NL}'){yCurrent = yCurrent - LineHeightUnits };
          }
        }else if(BP_Headers[FieldNumber]=="BP_Collation" || BP_Headers[FieldNumber]=="BP_Pagination" || BP_Headers[FieldNumber]=="BP_Foliation"){
          // collatiion, pagination, foliation required collation font for special characters
          CollationfontName = 'Collation';
          CollationfontStyle = 'normal';
          PDFdoc.setFont(CollationfontName, CollationfontStyle);
          LargestCutLinesY = yCurrent //- LineHeightUnits;          
          placeTextOnPage(PDFdoc, xCurrent, yCurrent, sectionWidth, BottomPageLimit, NumberOfColumns, ColumnSpacing, HeaderPresent, TextToPlace, textAlign, CollationfontName, CollationfontStyle, fontSize, LineHeightUnits, CutLines, DataElementNumber);          // return to user-specified font and style
          PDFdoc.setFont(fontName, fontStyle);
        }else if(FieldNumber >= 0 ){
          // for text other than collation, trade out paragraph marks
          TextToPlace = TextToPlace.replace(/¶¶/g,"{LF}{NL}");
          TextToPlace = TextToPlace.replace(/¶/g,"{NL}");
          placeTextOnPage(PDFdoc, xCurrent, yCurrent, sectionWidth, BottomPageLimit, NumberOfColumns, ColumnSpacing, HeaderPresent, TextToPlace, textAlign, fontName, fontStyle, fontSize, LineHeightUnits, CutLines, DataElementNumber);
        }else{
          // open fields may contain unicode characters in the 1st Astral plane, reduce the code point to a reference in the BMP. The Symbola font set (used beloe)
          // has had some glyphs in the BMP replaced with those from the 1st Astral plane: Leaves, ornaments, and others. This is necessary since the PDF maker cannot
          // accommodate references above the BMP.
          for( let j=0; j<TextToPlace.length; j++){
            let UnicodePoint = TextToPlace.codePointAt(j);
            if(UnicodePoint >= 128896){
              //replace reference to character on 1st astral plane with special positions in each of the standard font files (following italics) NB: this occupys TWO positions in the javascript string
              //ornament glyphs have been moved to the code points for Braille Pattern Dots.              
              SubstituteChar = String.fromCodePoint(UnicodePoint-118624);
              TextToPlace = TextToPlace.substring(0, j) + SubstituteChar + TextToPlace.substring(j + 2);              
            }else if(UnicodePoint >= 128592){
              //replace reference to character on 1st astral plane with special positions in each of the standard font files (following italics) NB: this occupys TWO positions in the javascript string
              //ornament glyphs have been moved to the code points for Braille Pattern Dots.              
              SubstituteChar = String.fromCodePoint(UnicodePoint-118352);
              TextToPlace = TextToPlace.substring(0, j) + SubstituteChar + TextToPlace.substring(j + 2);              
            }
          }
          TextToPlace = TextToPlace.replace(/<br><br>/g,"{LF}{NL}");
          TextToPlace = TextToPlace.replace(/<br>/g,"{NL}");
          TextToPlace = TextToPlace.replace(/\\n\\n/g,"{LF}{NL}");
          TextToPlace = TextToPlace.replace(/\\n/g,"{NL}");
          TextToPlace = TextToPlace.replace(/&amp;/g,'&');
          TextToPlace = TextToPlace.replace(/&gt;/g,'>');
          TextToPlace = TextToPlace.replace(/&lt;/g,'<');
          TextToPlace = TextToPlace.replace(/&quot;/g,'"');
          TextToPlace = TextToPlace.replace(/&#39;/g,"'");
          TextToPlace = TextToPlace.replace(/\(/g,"\(");
          TextToPlace = TextToPlace.replace(/\)/g,"\)");
          TextToPlace = TextToPlace.replace(/<em>/g,"<i>");
          TextToPlace = TextToPlace.replace(/<\/em>/g,"<i>");
          TextToPlace = TextToPlace.replace(/<\/i>/g,"<i>");
          TextToPlace = TextToPlace.replace(/¶¶/g,"{LF}{NL}");
          TextToPlace = TextToPlace.replace(/¶/g,"{NL}");            
          // for open fields, swith to font set with rich characters and symbols
          OpenFieldfontName = 'Symbola_hint' // 'LatinModernMath';
          OpenFieldfontStyle = 'normal';
          placeTextOnPage(PDFdoc, xCurrent, yCurrent, sectionWidth, BottomPageLimit, NumberOfColumns, ColumnSpacing, HeaderPresent, TextToPlace, textAlign, OpenFieldfontName, OpenFieldfontStyle, fontSize, LineHeightUnits, CutLines, DataElementNumber);
        }
        lastPageWritten = PDFpages[0];
      }
    }
  }
  PDFdoc.save(OutputFileName)
  if(ProduceSinglePDF){
    ProduceSinglePDF = false;
    CloseCompositionBoard()
  }
  document.getElementById("MakePDFProgress").style.display = "none";  

}

function placeTextOnPage(doc, xStart, yStart, sectionWidth, BottomPageLimit, NumberOfColumns, ColumnSpacing, HeaderPresent, text, textAlign, fontName, fontStyle, fontSize, LineHeightUnits, CutLines, DataElementNumber){
  //doc = jspdf generator
  let ConvertedText = text;
  let CollationUnicodeFound = false;
  // see if text contains Fraktur (in the unicode  1st astral plane)
  if( /[\u{1D400}-\u{1D7FF}}]/u.test(ConvertedText) ){
    for( let j=0; j<ConvertedText.length; j++){
      let UnicodePoint = ConvertedText.codePointAt(j);
      // capture the case for faktur letters
      if( UnicodePoint >= 119808 && UnicodePoint <= 120791){   //119808 is decimal value of the start of this group in the 1st astral plane
        let NumberOfAlpabetCycles = Math.floor((UnicodePoint-119808)/52);
        let LetterNumber = (UnicodePoint-119808) - 52*NumberOfAlpabetCycles
        //replace reference to character on 1st astral plane with special positions in each of the standard font files (following italics) NB: this occupys TWO positions in the javascript string
        SubstituteChar = String.fromCharCode(9524 + LetterNumber);  // 9534 is the decimal offset into the standard character set. -> position of upper case faktur A.
        ConvertedText = ConvertedText.substring(0, j) + SubstituteChar + ConvertedText.substring(j + 2);
        CollationUnicodeFound = true;
      }else if(UnicodePoint >= 120792) {  //120792 is decimal value for start of double-strke numbers (replaced with italic numbers)
        let NumberNumber = (UnicodePoint-120792)
        SubstituteChar = String.fromCharCode(9576 + NumberNumber);  // 9576 is the decimal offset into the standard character set. -> position of italic 0
        ConvertedText = ConvertedText.substring(0, j) + SubstituteChar + ConvertedText.substring(j + 2);
        CollationUnicodeFound = true;
      }
    }
  }
  //for the simple font systems, u2018 and u2019 mess up the character spacing for justified alignment. replace with apostrophe.
  if(fontName=="Helvetica" || fontName=="Courier" || fontName=="Times"){
    ConvertedText = ConvertedText.replace(/\u2018/g,"'").replace(/\u2019/g,"'");
  }

  let tempWidth = (sectionWidth-ColumnSpacing*(NumberOfColumns-1))/NumberOfColumns;
  let TextLinesInSections = [];
  let NumberOfLinesToOutput = [];
  doc.setFont(fontName, fontStyle);
  doc.setFontSize(fontSize);

  //split into sections. NB "{NL}" will be the tast characters in a text section that must be followed by a line feed
  TextSections = ConvertedText.split(/\{NL\}/);
  NumberOfTextSections = TextSections.length
  let NumberOfLineFeedsAtEndOfSection = [];
  for( m=0; m<NumberOfTextSections; m++){
    NumberOfLineFeedsAtEndOfSection[m]=0;
    if(TextSections[m].includes("{LF}")){
      NumberOfLineFeedsAtEndOfSection[m]=1
      //remove the "{NL}" flag so it does not distort the line width calc in jsPDF.
      TextSections[m] = TextSections[m].replace(/\{LF\}/g,"").trim();
    }else{TextSections[m] = TextSections[m].trim()}
    //split text to fit column width and accumulate total number of lines to be output. NB: font must be set in order for this work properly
    TextLinesInSections[m] = doc.splitTextToSize(TextSections[m],tempWidth);
    NumberOfLinesToOutput[m] = TextLinesInSections[m].length + NumberOfLineFeedsAtEndOfSection[m]
  }
  let NumberOfLines = NumberOfLinesToOutput.reduce(function(a,b){return a + b;}, 0);
  //if the text section has a header then one additional line is requied for 2nd and subsequent columns (1st line in these is empty/skipped)
  if(HeaderPresent){NumberOfLines = NumberOfLines + NumberOfColumns-1}
  let NumberOfLinesPerColumn = Math.floor(NumberOfLines/NumberOfColumns);
  if(NumberOfLinesPerColumn*NumberOfColumns != NumberOfLines){NumberOfLinesPerColumn += 1}

  //write out the text sections
  let NumberOfLinesWritten = 0;
  let CurrentColumn = 1;
  let CurrentTextSection = 0;
  let i=0;
  let yAdd = 0;
  let yPosition = yStart  //this is to the TOP of the section being written (or the next line to be written)
  //if we're just beginning and we're already at the bottom of a page, start a new page
  if(yPosition >= BottomPageLimit){
    if(yPosition > PDFpages[PDFpages[0]]){PDFpages[PDFpages[0]] = yPosition};
    //new page. reset all parameters
    doc.addPage(PaperSizeName[PaperNumber].toLowerCase(),'p')
    yAdd = 0;
    //this is to the upper margin of the next page; that is, the y position is to TOP of first line of text. But jsPDF generator assumes the BOTTOM of the first line of text. Adjust
    yPosition = PageMargins[3]+LineHeightUnits*0.5;
    yStart = PageMargins[3]+LineHeightUnits*0.5;
    PDFpages[0]++;
    PDFpages[PDFpages[0]] = 0;
    doc.setPage(PDFpages[0]);
    placeHeaderFooterOnPage(Header, Footer, doc, DataElementNumber)
    CutLines[0] = 0;
    CutLines.length = 1;
    NewPageTriggered = true;
  }

  let NumberOfLinesOutput = 0;
  let NumberOfLinesGathered = 0;
  while( CurrentColumn <= NumberOfColumns && CurrentTextSection < NumberOfTextSections && yPosition < BottomPageLimit ){
    ColumnOfText = "";
    while( NumberOfLinesGathered < NumberOfLinesPerColumn && i < TextLinesInSections[CurrentTextSection].length && yPosition < BottomPageLimit ){
      ColumnOfText += TextLinesInSections[CurrentTextSection][i]+" ";
      i++;
      NumberOfLinesGathered++;
      yPosition += LineHeightUnits;
      NumberOfLinesOutput++;
    }
    CutLines[0]++;
    CutLines[CutLines[0]] = [yPosition, xStart+(tempWidth+ColumnSpacing)*(CurrentColumn-1), xStart+tempWidth*CurrentColumn+ColumnSpacing*(CurrentColumn-1)];

    //done gathering lines. determine which alignment to use
    if(HeaderPresent && CurrentTextSection==0){
      TextAlignmentToUse = "left";
    }else if( textAlign != "justify" ){
      TextAlignmentToUse = textAlign;
    }else{
      //different possible justify alignment conditions
      if(i==TextLinesInSections[CurrentTextSection].length){
        //last line of a text section
        TextAlignmentToUse = "justifyBNTL";
      }else if( textAlign == "justify" && (CurrentColumn < NumberOfColumns || NumberOfLinesOutput != NumberOfLines ) ){
        //last line of column, unless all the text is output
        TextAlignmentToUse = "justify";
      }else{
        TextAlignmentToUse = "justifyBNTL";
      }
    }
    //output text into current column
    ColumnOfText = ColumnOfText.trim();
    if(HeaderPresent && CurrentTextSection==0){
      doc.setFont(fontName, "bold");
      doc.setFontSize(fontSize);
      doc.text(ColumnOfText, xStart+(tempWidth+ColumnSpacing)*(CurrentColumn-1), yStart+yAdd, {maxWidth: tempWidth.toString(), align: TextAlignmentToUse})
      doc.setFont(fontName, fontStyle);
    }else{
      if(fontStyle != 'normal' && fontStyle != 'bold'){
        //if font face is already italic (or bold italic) remove "<i>" tags
        ColumnOfText = ColumnOfText.replace(/<i>/g,"");
      }
      //find position(s) of font-face change toggles
      FontChangeToggle[0] = 0;
      let ItalicTogglePlace = ColumnOfText.indexOf("<i>");
      while( ItalicTogglePlace != -1){
        // record the char position of face change AFTER "<i>"" has been removed from the text string.
        FontChangeToggle[0]++;
        FontChangeToggle[FontChangeToggle[0]] = ItalicTogglePlace;
        //remove <i> from string
        ColumnOfText = ColumnOfText.slice(0,ItalicTogglePlace) + ColumnOfText.slice(ItalicTogglePlace+3);
        ItalicTogglePlace = ColumnOfText.indexOf("<i>");
      }
      //replace chars between italic-tags with substitute chars for italic 
      if(FontChangeToggle[0] > 0){
        for(let i=1; i<= FontChangeToggle[0]; i=i+2){
          let Index1 = FontChangeToggle[i];
          let Index2 = FontChangeToggle[i+1];
          for( let p=Index1; p<Index2; p++){
            let Ucode = ColumnOfText.charCodeAt(p)
            if(       Ucode >= 0x0041 && Ucode <= 0x005A){
              SubstituteChar = String.fromCharCode(0x2500 + Ucode - 0x0041);
            }else if( Ucode >= 0x0061 && Ucode <= 0x007A) {
              SubstituteChar = String.fromCharCode(0x251A + Ucode - 0x0061);
            }else{
              SubstituteChar = ColumnOfText.charAt(p)
            }
            ColumnOfText = ColumnOfText.slice(0,p) + SubstituteChar + ColumnOfText.slice(p+1);  
          }
        }
      }
      doc.text(ColumnOfText, xStart+(tempWidth+ColumnSpacing)*(CurrentColumn-1), yStart+yAdd, {maxWidth: tempWidth.toString(), align: TextAlignmentToUse})
    }
    
    // adjust running paramenters as necessary for new column, new text section, new page.
    if( yPosition + LineHeightUnits >= BottomPageLimit || NumberOfLinesGathered==NumberOfLinesPerColumn){
      //new column
      CurrentColumn++;
      if(HeaderPresent){  //skip the 1st line in subsequent columns if this text has a header
        yAdd = LineHeightUnits;
        //iff new column is valid and we're still accounting for the text header
        if(CurrentColumn <= NumberOfColumns){
          NumberOfLinesGathered = 1;
          NumberOfLinesOutput++;
        }
      }else{
        yAdd = 0;
        NumberOfLinesGathered = 0;
      }
      if(yPosition > PDFpages[PDFpages[0]]){PDFpages[PDFpages[0]] = yPosition};
      yPosition = yStart+yAdd;
      //if we're at the 1st line of a column, supress LF at the end of the text section just completed
      if(i==TextLinesInSections[CurrentTextSection].length){NumberOfLineFeedsAtEndOfSection[CurrentTextSection]=0}
    }

    if(i==TextLinesInSections[CurrentTextSection].length){
      //new text section
      i=0;
      yAdd = (NumberOfLinesGathered+NumberOfLineFeedsAtEndOfSection[CurrentTextSection])*LineHeightUnits;
      NumberOfLinesOutput += NumberOfLineFeedsAtEndOfSection[CurrentTextSection];
      NumberOfLinesGathered += NumberOfLineFeedsAtEndOfSection[CurrentTextSection];
      yPosition += NumberOfLineFeedsAtEndOfSection[CurrentTextSection]*LineHeightUnits;
      CurrentTextSection++;
    }

    if(CurrentColumn > NumberOfColumns && CurrentTextSection != NumberOfTextSections ){
      //record position of last line written before we leave current page
      if(yPosition > PDFpages[PDFpages[0]]){PDFpages[PDFpages[0]] = yPosition};
      //new page. reset all parameters
      doc.addPage(PaperSizeName[PaperNumber].toLowerCase(),'p')
      yAdd = 0;
      //this is to the upper margin of the next page; that is, the y position is to TOP of first line of text. But jsPDF generator assumes the BOTTOM of the first line of text. Adjust
      yPosition = PageMargins[3]+LineHeightUnits*0.5;
      yStart = PageMargins[3]+LineHeightUnits*0.5;
      PDFpages[0]++;
      PDFpages[PDFpages[0]] = 0;
      doc.setPage(PDFpages[0]);
      placeHeaderFooterOnPage(Header, Footer, doc, DataElementNumber)
      CurrentColumn = 1;
      NumberOfLinesPerColumn = Math.floor((NumberOfLines-NumberOfLinesOutput)/NumberOfColumns);
      if(NumberOfLinesPerColumn*NumberOfColumns != (NumberOfLines-NumberOfLinesOutput)){NumberOfLinesPerColumn += 1}
      NumberOfLinesGathered = 0;
      //turn off header flag so that next set of columns do not skip 1st line, even if this section of text started with a header on the previous page.
      HeaderPresent = false;
      CutLines[0] = 0;
      CutLines.length = 1;
      NewPageTriggered = true;
    }
  }
}

function placeHeaderFooterOnPage(Header, Footer, doc, DataElementNumber){

  if(Header.Present) {
    //write out header text
    let headerXStart = Header.Xposition;
    let headerYStart = Header.Yposition;
    let headerSectionWidth = Header.Width;
    let headerTextAlign = Header.TextAlign;
    let headerFontName = Header.fontName;
    let headerFontStyle = Header.fontStyle;
    let headerFontSize = Header.fontSize;
    let headerLineHeightUnits = Header.LineHeightUnits;
    let headContentlink = Header.ContentLink;
    if( headContentlink > 0 ){
      // check header for possible broken/missing content links
        HeaderToMatch = BP_Headers[headContentlink];
        FieldNumber = 0;
        for( let k = 1; k < BP_Headers.length; k++){
          if( HeaderToMatch == BP_Headers[k]){
            FieldNumber = k;
            break
          }
        }  
      if(FieldNumber != 0 ){
        headerText = data[DataElementNumber][BP_Headers[FieldNumber]];
      }else{
        headerText = "";
      }
    }else{
      if(!document.getElementById('HeaderContent')){
        headerText = "";
      }else{
        ContentToPrint = document.getElementById('HeaderContent');
        headerText = ContentToPrint.value;
      }
    }
    // //insert CRLF or NL in the text, and decode any ampersands and other special characters
    // headerText = headerText.replace(/¶¶/g,"{LF}{NL}");
    // headerText = headerText.replace(/¶/g,"{NL}");
    // headerText = headerText.replace(/<br><br>/g,"{LF}{NL}");
    // headerText = headerText.replace(/<br>/g,"{NL}");
    // headerText = headerText.replace(/\\n\\n/g,"{LF}{NL}");
    // headerText = headerText.replace(/\\n/g,"{NL}");
    // headerText = headerText.replace(/&amp;/g,'&');
    // headerText = headerText.replace(/&gt;/g,'>');
    // headerText = headerText.replace(/&lt;/g,'<');
    // headerText = headerText.replace(/&quot;/g,'"');
    // headerText = headerText.replace(/&#39;/g,"'");
    // headerText = headerText.replace(/\(/g,"\(");
    // headerText = headerText.replace(/\)/g,"\)");
    // //convert emphasis HTML tag to italic and closing tags to same a opening tags. processing with use "<i>" as on/off toggle for italic font face.
    // headerText = headerText.replace(/<em>/g,"<i>");
    // headerText = headerText.replace(/<\/em>/g,"<i>");
    // headerText = headerText.replace(/<\/i>/g,"<i>");

    doc.setFont(headerFontName, headerFontStyle);
    doc.setFontSize(headerFontSize);
    doc.text(headerText, headerXStart, headerYStart, {maxWidth: headerSectionWidth.toString(), align: headerTextAlign})
  }
  if(Footer.Present){
    //write out footer text
    let footerXStart = Footer.Xposition;
    let footerYStart = Footer.Yposition;
    let footerSectionWidth = Footer.Width;
    let footerText = Footer.Content;
    let footerTextAlign = Footer.TextAlign;
    let footerFontName = Footer.fontName;
    let footerFontStyle = Footer.fontStyle;
    let footerFontSize = Footer.fontSize;
    let footerLineHeightUnits = Footer.LineHeightUnits;
    let footerContentlink = Footer.ContentLink;
    if( footerContentlink > 0 ){
      // check header for possible broken/missing content links
      HeaderToMatch = BP_Headers[footerContentlink];
      FieldNumber = 0;
      for( let k = 1; k < BP_Headers.length; k++){
        if( HeaderToMatch == BP_Headers[k]){
          FieldNumber = k;
          break
        }
      }  
      if(FieldNumber != 0 ){
        FooterText = data[DataElementNumber][BP_Headers[FieldNumber]];
      }else{
        FooterText = "";
      }
    }else{
      if(!document.getElementById('FooterContent')){
        footerText = "";
      }else{
        ContentToPrint = document.getElementById('FooterContent');
        footerText = ContentToPrint.value;
      }
    }
    doc.setFont(footerFontName, footerFontStyle);
    doc.setFontSize(footerFontSize);
    doc.text(footerText, footerXStart, footerYStart, {maxWidth: footerSectionWidth.toString(), align: footerTextAlign})
  }

  //write out page numbers, if requested
  if(Header.Present && Header.pageNumberInHeader !=0) {
    let pageNumber = PDFpages[0];   //PDFpages[0] is the current page number
    let pageNumberPosition = Header.pageNumberPosition;
    let headerXStart = Header.Xposition;
    let headerYStart = Header.Yposition;
    let headerSectionWidth = Header.Width;
    let headerFontName = Header.fontName;
    let headerFontStyle = Header.fontStyle;
    let headerFontSize = Header.fontSize;
    doc.setFont(headerFontName, headerFontStyle);
    doc.setFontSize(headerFontSize);
    if(pageNumberPosition=='TopCenter'){doc.text(pageNumber.toString(), headerXStart, headerYStart, {maxWidth: headerSectionWidth.toString(), align: "center"})}
    else if(pageNumberPosition=='TopLeft'){doc.text(pageNumber.toString(), headerXStart, headerYStart, {maxWidth: headerSectionWidth.toString(), align: "left"})}
    else if(pageNumberPosition=='TopRight'){doc.text(pageNumber.toString(), headerXStart, headerYStart, {maxWidth: headerSectionWidth.toString(), align: "right"})}
    else if(pageNumberPosition=='TopEvenOdd'){
      if(pageNumber%2 == 0){
        //page is even numbered
        doc.text(pageNumber.toString(), headerXStart, headerYStart, {maxWidth: headerSectionWidth.toString(), align: "left"})
      }else{
        //page is odd numbered
        doc.text(pageNumber.toString(), headerXStart, headerYStart, {maxWidth: headerSectionWidth.toString(), align: "right"})
      }
    }
  }
  if(Footer.Present && Footer.pageNumberInFooter  !=0){
    let pageNumber = PDFpages[0];   //PDFpages[0] is the current page number
    let pageNumberPosition = Footer.pageNumberPosition;
    let footerXStart = Footer.Xposition;
    let footerYStart = Footer.Yposition;
    let footerSectionWidth = Footer.Width;
    let footerFontName = Footer.fontName;
    let footerFontStyle = Footer.fontStyle;
    let footerFontSize = Footer.fontSize;
    doc.setFont(footerFontName, footerFontStyle);
    doc.setFontSize(footerFontSize);
    if(pageNumberPosition=='BottomCenter'){doc.text(pageNumber.toString(), footerXStart, footerYStart, {maxWidth: footerSectionWidth.toString(), align: "center"})}
    else if(pageNumberPosition=='BottomLeft'){doc.text(pageNumber.toString(), footerXStart, footerYStart, {maxWidth: footerSectionWidth.toString(), align: "left"})}
    else if(pageNumberPosition=='BottomRight'){doc.text(pageNumber.toString(), footerXStart, footerYStart, {maxWidth: footerSectionWidth.toString(), align: "right"})}
    else if(pageNumberPosition=='BottomEvenOdd'){
      if(pageNumber%2 == 0){
        //page is even numbered
        doc.text(pageNumber.toString(), footerXStart, footerYStart, {maxWidth: footerSectionWidth.toString(), align: "left"})
      }else{
        //page is odd numbered
        doc.text(pageNumber.toString(), footerXStart, footerYStart, {maxWidth: footerSectionWidth.toString(), align: "right"})
      }
    }

  }

  //reset font
  doc.setFont(CurrentFont, CurrentFontFace);
  doc.setFontSize(CurrentFontSize);

}

function ImageLoad(src){
  return new Promise((resolve, reject) => {
    var img = new Image();
    img.addEventListener( "load", ()=> resolve(img) );
    img.addEventListener( "error", err => reject(err))
    img.src = src
  })
}
//#endregion

// Creating Geograpy & Map files -->
//#region

function OpenMap(){
  if(WorldMapOpened==true){
    CloseMap();
    WorldMapOpened = false;
    return
  }
  WorldMapOpened = true;
  // location data formating
  // LocationDataPresent = boolean indicating if any element shave location data
  // LocationData[0] = number of locations
  // LocationData[i][0] = pointer to element of ith entry in location list (not all elements may have locationd data)
  // LocationData[i][1] = data[i]["BP_Latitude"] turned into number
  // LocationData[i][2] = data[i]["BP_Longitude"] turned into number
  if(!LocationDataPresent){
    // no location data, just show empty map
    let MapElement = document.getElementById("Map");
    MapElement.style.display = 'block';  
    map = L.map('MapDisplay').setView([48.8566, 2.3522], 10);    //48.8566° N, 2.3522° E

    localforage.getItem('CurrentMapBackgroundNumber').then(function(value) {
      CurrentMapBackgroundNumber = value;
    }).catch(function(err) {
      CurrentMapBackgroundNumber = 1;        
    });
    SetMapBackground(CurrentMapBackgroundNumber);    
    return
  }
  // mount LocationDataDescription in display bar
  document.getElementById("NameOfLocationData").innerHTML = LocationDataDescription;
  // accounting of currently expanded clusters
  CurrentSiteClusterExpanded[0] = 0;
  //  if we get here, location data is present. find any data clusters
  let AveLat = 0;
  let AveLong = 0;
  let NumberOfClusters = 0
  let LocationUsed = [];
  SiteCluster.length = 0;
  for (let i = 1; i <= LocationData[0]; i++) {LocationUsed[i]=false};

  for (let i = 1; i <= LocationData[0]; i++) {
    if(LocationUsed[i]){continue}
    SiteAnchorNumber = i;
    NumberOfClusters++
    SiteCluster[NumberOfClusters] = [];
    SiteCluster[NumberOfClusters][0] = 1;
    SiteCluster[NumberOfClusters][1] = i;
    LocationUsed[i] = true;
    AveLat = AveLat + LocationData[i][1];
    AveLong = AveLong + LocationData[i][2];
    for (let j = 1; j <= LocationData[0]; j++) {
      if(j==i){continue};
      if(LocationUsed[j]){continue};
      // see if location j is close to location i
      if( Math.abs(LocationData[j][1]-LocationData[i][1]) <= 0.02 && Math.abs(LocationData[j][2]-LocationData[i][2]) <= 0.02 ){
        SiteCluster[NumberOfClusters][0]++
        SiteCluster[NumberOfClusters][SiteCluster[NumberOfClusters][0]] = j;
        LocationUsed[j] = true;
        AveLat = AveLat + LocationData[j][1];
        AveLong = AveLong + LocationData[j][2];        
      }
    }  
  }
  
  AveLat = (90-AveLat/LocationData[0])/2
  AveLong = (90-AveLong/LocationData[0])/2
  //open map
  let MapElement = document.getElementById("Map");
  MapElement.style.display = 'block';  
  map = L.map('MapDisplay').setView([AveLong, AveLat], 3);

  localforage.getItem('CurrentMapBackgroundNumber').then(function(value) {
    CurrentMapBackgroundNumber = value;
  }).catch(function(err) {
    CurrentMapBackgroundNumber = 1;        
  });
  SetMapBackground(CurrentMapBackgroundNumber);    

  //define marker icons, red = single, blue = multiple
  let myRedIcon = L.icon({
    iconUrl: './BP Appearance/BP Icon.png',
    iconSize: [30, 30],
    iconAnchor: [15, 15]
    // popupAnchor: [-3, -76]
    // shadowUrl: 'my-icon-shadow.png',
    // shadowSize: [68, 95],
    // shadowAnchor: [22, 94]
  });
  let myBlueIcon = L.icon({
    iconUrl: './BP Appearance/BP Icon Blue.png',
    iconSize: [30, 30],
    iconAnchor: [15, 15]
    // popupAnchor: [-3, -76]
    // shadowUrl: 'my-icon-shadow.png',
    // shadowSize: [68, 95],
    // shadowAnchor: [22, 94]
  });
 
  // for (let i = 1; i <= NumOfTimelineElements; i++) {
  for (let i = 1; i <= NumberOfClusters; i++) {
      if(SiteCluster[i][0]==1){
        // marker for location of single item
        let MapMarker = L.marker( [LocationData[SiteCluster[i][1]][1], LocationData[SiteCluster[i][1]][2]], 
                                  {icon: myRedIcon, riseOnHover: true, title: data[SiteCluster[i][1]]["BP_Short or Common Title"]+' '+data[SiteCluster[i][1]]["BP_Date Published"]}).on('contextmenu', ()=>OpenElementDetailData(SiteCluster[i][1])).addTo(map);
        MapMarker.bindPopup( data[SiteCluster[i][1]]["BP_Short or Common Title"] + '<br>' + data[SiteCluster[i][1]]["BP_Author full name"] + '<br>' + data[SiteCluster[i][1]]["BP_Date Published"]).openPopup();
      }else{
        // marker for location of multiple items
        let MapMarker = L.marker( [LocationData[SiteCluster[i][1]][1], LocationData[SiteCluster[i][1]][2]], 
                                  {icon: myBlueIcon, riseOnHover: true, title: SiteCluster[i][0].toString()+' Items' }).on('contextmenu', ()=>MultipleMapItemExpand(i)).addTo(map);
        MapMarker.bindPopup( data[SiteCluster[i][1]]["BP_Short or Common Title"] + '<br>' + data[SiteCluster[i][1]]["BP_Author full name"] + '<br>' + data[SiteCluster[i][1]]["BP_Date Published"]).openPopup();
      }
  }

}

function MultipleMapItemExpand(ClusterNumber){
  // if this cluster is currently expanded, contract it: remove layer with cluster markers and pack cluster expanded list
  for( let m = 1; m <= CurrentSiteClusterExpanded[0]; m++){
   if(CurrentSiteClusterExpanded[m] == ClusterNumber){
     map.removeLayer(MultipleMarkersGroup[ClusterNumber]);
    //  pack list of expanded clusters
    NumberExpanded = 0
    for( let p = 1; p <= CurrentSiteClusterExpanded[0]; p++){
      if(CurrentSiteClusterExpanded[p]==ClusterNumber){continue}
      NumberExpanded++
      CurrentSiteClusterExpanded[NumberExpanded]=CurrentSiteClusterExpanded[p]
    }
    CurrentSiteClusterExpanded[0] = NumberExpanded;
    return
   }
  }
  // if we get here, cluster is expanded. add to expanded cluster list
  CurrentSiteClusterExpanded[0]++
  CurrentSiteClusterExpanded[CurrentSiteClusterExpanded[0]] = ClusterNumber;
  // create map layer group that will hold cluster markers
  MultipleMarkersGroup[ClusterNumber] = L.layerGroup().addTo(map);
  let myRedIcon = L.icon({
    iconUrl: './BP Appearance/BP Icon.png',
    iconSize: [30, 30],
    iconAnchor: [15, 15]
  });
  // build concentric circle of icons from the SiteCluster list
  let center =  [LocationData[SiteCluster[ClusterNumber][1]][1], LocationData[SiteCluster[ClusterNumber][1]][2]];
  // grouped in rignt of increasting number
  // SiteCluster[i][0]
  let start = 0;
  let stop  = 0;
  let radius = 0;
  let radius0 = 0.125
  let angle = 0;
  let Steps = 0;
  start = 1;
  for( let ring = 1; ring <= 10; ring++){
    stop  = Math.min(start + 15*ring, SiteCluster[ClusterNumber][0]);
    radius = radius0+0.025*ring
    LatitudeScaling = Math.cos(center[0]*Math.PI/180)
    Steps = stop-start+1
    for( let m = start; m <= stop ; m++){
      angle = (m-1)*2*Math.PI/Steps;            
      RotationScaling = LatitudeScaling/Math.sqrt( (LatitudeScaling*Math.cos(angle))**2 + Math.sin(angle)**2 ) 
      Eradius = radius*RotationScaling
      Eangle = Math.asin(LatitudeScaling*Math.sin(angle))
      Lat = center[0] + Math.sin(angle)*Eradius;
      Lng = center[1] + Math.cos(angle)*Eradius;
      let MapMarker = L.marker([Lat,Lng], 
        {icon: myRedIcon, riseOnHover: true, 
         title: data[SiteCluster[ClusterNumber][m]]["BP_Short or Common Title"]+' '+data[SiteCluster[ClusterNumber][m]]["BP_Date Published"]
        } ).on('contextmenu', ()=>OpenElementDetailData(SiteCluster[ClusterNumber][m])).addTo(MultipleMarkersGroup[ClusterNumber]);
      let Wheelline = L.polyline( [center,[Lat,Lng]], {color: 'black', weight: 1}).addTo(MultipleMarkersGroup[ClusterNumber])
    }
    start = stop + 1;
    // L.circle(center, {radius: 30000}).addTo(map);
    if(stop == SiteCluster[ClusterNumber][0]){break};
  }
}

function CycleTrhoughMapBackgrounds(){
  try{map.removeLayer(MapLay1)}catch{};
  try{map.removeLayer(MapLay2)}catch{};
  MapBackgroundNumber++
  if(MapBackgroundNumber > 8){MapBackgroundNumber = 1}
  CurrentMapBackgroundNumber = MapBackgroundNumber;
  localforage.setItem('CurrentMapBackgroundNumber',CurrentMapBackgroundNumber)  
  SetMapBackground(CurrentMapBackgroundNumber);
  }

function SetMapBackground(CurrentMapBackgroundNumber){

  switch(CurrentMapBackgroundNumber){
   case 1:
    // gray
    //map.eachLayer(function (layer) {map.removeLayer(layer)});
    MapLay1 = L.tileLayer('https://server.arcgisonline.com/ArcGIS/rest/services/Canvas/World_Light_Gray_Base/MapServer/tile/{z}/{y}/{x}', {
    attribution: 'Tiles &copy; Esri &mdash; Esri, DeLorme, NAVTEQ',
    maxZoom: 16
    }).addTo(map);
    MapLay2 = L.tileLayer('https://stamen-tiles-{s}.a.ssl.fastly.net/toner-hybrid/{z}/{x}/{y}{r}.{ext}', {
    attribution: 'Map tiles by <a href="http://stamen.com">Stamen Design</a>, <a href="http://creativecommons.org/licenses/by/3.0">CC BY 3.0</a> &mdash; Map data &copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> contributors',
    subdomains: 'abcd',
    minZoom: 0,
    maxZoom: 20,
    ext: 'png'
    }).addTo(map);
    break

   case 2:
   // pale background, native spelling for labelling
   //map.eachLayer(function (layer) {map.removeLayer(layer)});   
   MapLay1 = L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
     maxZoom: 19,
     attribution: '© OpenStreetMap'
   }).addTo(map);

   case 3:
    // pale white background, light labellling
    MapLay1 = L.tileLayer('https://{s}.basemaps.cartocdn.com/light_all/{z}/{x}/{y}{r}.png', {
    attribution: '&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> contributors &copy; <a href="https://carto.com/attributions">CARTO</a>',
    subdomains: 'abcd',
    maxZoom: 20
    }).addTo(map);
    break

   case 4:
    // pale background, blue labellling
    MapLay1 = L.tileLayer('https://{s}.basemaps.cartocdn.com/rastertiles/voyager/{z}/{x}/{y}{r}.png', {
    attribution: '&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> contributors &copy; <a href="https://carto.com/attributions">CARTO</a>',
    subdomains: 'abcd',
    maxZoom: 20
    }).addTo(map);
    break

   case 5:
    // pale backgrond, pale blue labelling
    MapLay1 = L.tileLayer('https://{s}.basemaps.cartocdn.com/rastertiles/voyager_labels_under/{z}/{x}/{y}{r}.png', {
    attribution: '&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> contributors &copy; <a href="https://carto.com/attributions">CARTO</a>',
    subdomains: 'abcd',
    maxZoom: 20
    }).addTo(map);
    break

   case 6:
    // black & white stark
    MapLay1 = L.tileLayer('https://stamen-tiles-{s}.a.ssl.fastly.net/toner-lite/{z}/{x}/{y}{r}.{ext}', {
    attribution: 'Map tiles by <a href="http://stamen.com">Stamen Design</a>, <a href="http://creativecommons.org/licenses/by/3.0">CC BY 3.0</a> &mdash; Map data &copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> contributors',
    subdomains: 'abcd',
    minZoom: 0,
    maxZoom: 20,
    ext: 'png'
    }).addTo(map);
    MapLay2 = L.tileLayer('https://stamen-tiles-{s}.a.ssl.fastly.net/toner-hybrid/{z}/{x}/{y}{r}.{ext}', {
    attribution: 'Map tiles by <a href="http://stamen.com">Stamen Design</a>, <a href="http://creativecommons.org/licenses/by/3.0">CC BY 3.0</a> &mdash; Map data &copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> contributors',
    subdomains: 'abcd',
    minZoom: 0,
    maxZoom: 20,
    ext: 'png'
    }).addTo(map);
    break

   case 7:
    // light terrain background, black labelling
    MapLay1 = L.tileLayer('https://stamen-tiles-{s}.a.ssl.fastly.net/terrain-background/{z}/{x}/{y}{r}.{ext}', {
    attribution: 'Map tiles by <a href="http://stamen.com">Stamen Design</a>, <a href="http://creativecommons.org/licenses/by/3.0">CC BY 3.0</a> &mdash; Map data &copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> contributors',
    subdomains: 'abcd',
    minZoom: 0,
    maxZoom: 18,
    ext: 'png'
    }).addTo(map);
    MapLay2 = L.tileLayer('https://stamen-tiles-{s}.a.ssl.fastly.net/toner-hybrid/{z}/{x}/{y}{r}.{ext}', {
    attribution: 'Map tiles by <a href="http://stamen.com">Stamen Design</a>, <a href="http://creativecommons.org/licenses/by/3.0">CC BY 3.0</a> &mdash; Map data &copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> contributors',
    subdomains: 'abcd',
    minZoom: 0,
    maxZoom: 20,
    ext: 'png'
    }).addTo(map);
    break

   case 8:
    //lightly colored, black labelling
    MapLay1 = L.tileLayer('https://server.arcgisonline.com/ArcGIS/rest/services/World_Topo_Map/MapServer/tile/{z}/{y}/{x}', {
	  attribution: 'Tiles &copy; Esri &mdash; Esri, DeLorme, NAVTEQ, TomTom, Intermap, iPC, USGS, FAO, NPS, NRCAN, GeoBase, Kadaster NL, Ordnance Survey, Esri Japan, METI, Esri China (Hong Kong), and the GIS User Community'
    }).addTo(map);
    MapLay2 = L.tileLayer('https://stamen-tiles-{s}.a.ssl.fastly.net/toner-hybrid/{z}/{x}/{y}{r}.{ext}', {
	  attribution: 'Map tiles by <a href="http://stamen.com">Stamen Design</a>, <a href="http://creativecommons.org/licenses/by/3.0">CC BY 3.0</a> &mdash; Map data &copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> contributors',
	  subdomains: 'abcd',
	  minZoom: 0,
	  maxZoom: 20,
	  ext: 'png'
    }).addTo(map);
  }

}

function CloseMap(){
  document.getElementById("Map").style.display = 'none';
  document.getElementById("MapDisplay").innerHTML = '';
  try{map.remove()}catch{};
  WorldMapOpened = false;
}

function MakeGoogleEarthFile(){
  //make a Google Earth / Google Map place file of items in library as specified by lat, long for each item
  let BookTitle = '';
  let BookData = '';
  KMLtext =  '<?xml version="1.0" encoding="UTF-8"?>'+'\n'
  KMLtext += '<kml xmlns="http://www.opengis.net/kml/2.2">'+'\n'
  KMLtext += '<Document>'+'\n'
    for (let i = 1; i < data.length; i++) {
  // for (let i = 1; i < 250; i++) {
    if( !DisplayElement[i] ) {continue};                                                                              //include only displayed elements
    if( data[i]["BP_Latitude"] == null || data[i]["BP_Latitude"] == '' || data[i]["BP_Latitude"] == ' ')  {continue}; //include iff has lat/long data
    if( data[i]["BP_Longitude"] == null || data[i]["BP_Longitude"] == '' || data[i]["BP_Longitude"] == ' ')  {continue}; //include iff has lat/long data
    KMLtext += '<Placemark>'+'\n'
    BookTitle = data[i]["BP_Short or Common Title"];
    BookTitle = BookTitle.replace(/&/g, "&amp;");                                                                     // & (isolated) is not allowed in this XML file
    BookDate = " "+data[i]["BP_Date Published"].toString();
    KMLtext += '  <name>'+BookTitle+BookDate+'</name>'+'\n'
    KMLtext += '  <description>'+NameOfLibrary.trim()+'</description>'+'\n'
    // KMLtext += '  <Style>'+'\n'
    // KMLtext += '    <IconStyle>'+'\n'
    // KMLtext += '      <Icon>'+'\n'
    // KMLtext += '        <href>http://maps.google.com/mapfiles/kml/shapes/homegardenbusiness.png</href>'+'\n'
    // KMLtext += '      </Icon>'+'\n'
    // KMLtext += '    </IconStyle>'+'\n'
    // KMLtext += '    <LabelStyle>'+'\n'
    // KMLtext += '       <color>9f0000ff</color>'+'\n'
    // KMLtext += '       <scale>1.5</scale>'+'\n'
    // KMLtext += '    </LabelStyle>'+'\n'
    // KMLtext += '  </Style>'+'\n'
    KMLtext += '    <Point>'+'\n'
    KMLtext += ' <coordinates>'+data[i]["BP_Longitude"].toString()+','+data[i]["BP_Latitude"].toString()+',0</coordinates>'+'\n'
    KMLtext += '  </Point>'+'\n'
    KMLtext += '</Placemark>'+'\n'
  }

  KMLtext += '</Document>'+'\n'
  KMLtext += '</kml>'
  download(KMLtext, "LibraryName GoogleEarth.kml","text/plain");  
}

//#endregion