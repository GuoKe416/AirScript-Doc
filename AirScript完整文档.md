# AirScript 完整文档

## 目录

- [产品介绍](#产品介绍)
  - [概述](#概述)
  - [脚本语言](#脚本语言)
- [快速上手](#快速上手)
  - [开始](#开始)
  - [最佳实践](#最佳实践)
  - [配置视图](#配置视图)
- [脚本令牌](#脚本令牌)
  - [应用场景](#应用场景)
  - [接口说明](#接口说明)
  - [简介](#简介)
- [示范案例](#示范案例)
  - [多维表](#多维表)
  - [表格](#表格)
- [API文档(1.0)](#api文档-1-0)
  - [内置函数](#内置函数)
  - [多维表格](#多维表格)
    - [字段](#字段)
    - [行记录](#行记录)
    - [表](#表)
    - [视图](#视图)
    - [选区](#选区)
    - [附录](#附录)
  - [智能表格](#智能表格)
    - [工作表](#工作表)
      - [API总览](#api总览)
      - [区域(Range)](#区域-range)
      - [图形(Shape)](#图形-shape)
      - [图表(Chart)](#图表-chart)
      - [字体(Font)](#字体-font)
      - [字段(Field)](#字段-field)
      - [工作簿(Workbook)](#工作簿-workbook)
      - [工作表(Sheet)](#工作表-sheet)
      - [工作表函数(WorksheetFunction)](#工作表函数-worksheetfunction)
      - [排序(Sort)](#排序-sort)
      - [排序字段(SortField)](#排序字段-sortfield)
      - [数据有效性规则(Validation)](#数据有效性规则-validation)
      - [数据表(Sheet)](#数据表-sheet)
      - [条件格式(FormatCondition)](#条件格式-formatcondition)
      - [条件格式集合(FormatConditions)](#条件格式集合-formatconditions)
      - [枚举(Enum)](#枚举-enum)
      - [筛选(AutoFilter)](#筛选-autofilter)
      - [行记录(Record)](#行记录-record)
      - [表格实例(Application)](#表格实例-application)
      - [超链接(Hyperlink)](#超链接-hyperlink)
      - [边框(Border)](#边框-border)
      - [附录](#附录)
    - [数据表](#数据表)
      - [字段(Field)](#字段-field)
      - [数据表(Sheet)](#数据表-sheet)
      - [枚举(Enum)](#枚举-enum)
      - [行记录(Record)](#行记录-record)
      - [表格实例(Application)](#表格实例-application)
      - [附录](#附录)
  - [高级服务](#高级服务)
    - [云文档 API](#云文档-api)
    - [数据库 API](#数据库-api)
    - [概述](#概述)
    - [网络 API](#网络-api)
    - [邮件 API](#邮件-api)
- [API文档(2.0)](#api文档-2-0)
  - [概述](#概述)
  - [智能表格](#智能表格)
    - [工作表](#工作表)
      - [AboveAverage 对象](#aboveaverage-对象)
      - [Adjustments 对象](#adjustments-对象)
      - [AllowEditRange 对象](#alloweditrange-对象)
      - [AllowEditRanges 对象](#alloweditranges-对象)
      - [Application 对象](#application-对象)
      - [Areas 对象](#areas-对象)
      - [AutoFilter 对象](#autofilter-对象)
      - [Axes 对象](#axes-对象)
      - [Axis 对象](#axis-对象)
      - [AxisTitle 对象](#axistitle-对象)
      - [Border 对象](#border-对象)
      - [Borders 对象](#borders-对象)
      - [CalculatedFields 对象](#calculatedfields-对象)
      - [CalculatedItems 对象](#calculateditems-对象)
      - [CellFormat 对象](#cellformat-对象)
      - [Characters 对象](#characters-对象)
      - [Chart 对象](#chart-对象)
      - [ChartArea 对象](#chartarea-对象)
      - [ChartCategory 对象](#chartcategory-对象)
      - [ChartFormat 对象](#chartformat-对象)
      - [ChartGroup 对象](#chartgroup-对象)
      - [ChartGroups 对象](#chartgroups-对象)
      - [ChartObject 对象](#chartobject-对象)
      - [ChartObjects 对象](#chartobjects-对象)
      - [ChartTitle 对象](#charttitle-对象)
      - [Charts 对象](#charts-对象)
      - [ColorFormat 对象](#colorformat-对象)
      - [ColorScale 对象](#colorscale-对象)
      - [ColorScaleCriteria 对象](#colorscalecriteria-对象)
      - [ColorStop 对象](#colorstop-对象)
      - [ColorStops 对象](#colorstops-对象)
      - [Comment 对象](#comment-对象)
      - [Comments 对象](#comments-对象)
      - [ConditionValue 对象](#conditionvalue-对象)
      - [ConnectorFormat 对象](#connectorformat-对象)
      - [ControlFormat 对象](#controlformat-对象)
      - [CustomProperties 对象](#customproperties-对象)
      - [CustomProperty 对象](#customproperty-对象)
      - [DataBarBorder 对象](#databarborder-对象)
      - [DataLabel 对象](#datalabel-对象)
      - [DataLabels 对象](#datalabels-对象)
      - [DataTable 对象](#datatable-对象)
      - [Databar 对象](#databar-对象)
      - [DisplayFormat 对象](#displayformat-对象)
      - [DisplayUnitLabel 对象](#displayunitlabel-对象)
      - [DownBars 对象](#downbars-对象)
      - [DropLines 对象](#droplines-对象)
      - [Error 对象](#error-对象)
      - [ErrorBars 对象](#errorbars-对象)
      - [Errors 对象](#errors-对象)
      - [FillFormat 对象](#fillformat-对象)
      - [Filter 对象](#filter-对象)
      - [Filters 对象](#filters-对象)
      - [Font 对象](#font-对象)
      - [FormatColor 对象](#formatcolor-对象)
      - [FormatCondition 对象](#formatcondition-对象)
      - [FormatConditions 对象](#formatconditions-对象)
      - [FreeformBuilder 对象](#freeformbuilder-对象)
      - [Gridlines 对象](#gridlines-对象)
      - [GroupShapes 对象](#groupshapes-对象)
      - [HiLoLines 对象](#hilolines-对象)
      - [Hyperlink 对象](#hyperlink-对象)
      - [Hyperlinks 对象](#hyperlinks-对象)
      - [Icon 对象](#icon-对象)
      - [IconCriteria 对象](#iconcriteria-对象)
      - [IconSet 对象](#iconset-对象)
      - [IconSetCondition 对象](#iconsetcondition-对象)
      - [IconSets 对象](#iconsets-对象)
      - [Interior 对象](#interior-对象)
      - [LeaderLines 对象](#leaderlines-对象)
      - [Legend 对象](#legend-对象)
      - [LegendEntries 对象](#legendentries-对象)
      - [LegendEntry 对象](#legendentry-对象)
      - [LegendKey 对象](#legendkey-对象)
      - [LineFormat 对象](#lineformat-对象)
      - [LinearGradient 对象](#lineargradient-对象)
      - [ListColumn 对象](#listcolumn-对象)
      - [ListColumns 对象](#listcolumns-对象)
      - [ListObject 对象](#listobject-对象)
      - [ListObjects 对象](#listobjects-对象)
      - [ListRow 对象](#listrow-对象)
      - [ListRows 对象](#listrows-对象)
      - [Name 对象](#name-对象)
      - [Names 对象](#names-对象)
      - [Outline 对象](#outline-对象)
      - [PictureFormat 对象](#pictureformat-对象)
      - [PivotAxis 对象](#pivotaxis-对象)
      - [PivotCell 对象](#pivotcell-对象)
      - [PivotField 对象](#pivotfield-对象)
      - [PivotFields 对象](#pivotfields-对象)
      - [PivotFilter 对象](#pivotfilter-对象)
      - [PivotFilters 对象](#pivotfilters-对象)
      - [PivotFormula 对象](#pivotformula-对象)
      - [PivotFormulas 对象](#pivotformulas-对象)
      - [PivotItem 对象](#pivotitem-对象)
      - [PivotItemList 对象](#pivotitemlist-对象)
      - [PivotItems 对象](#pivotitems-对象)
      - [PivotLayout 对象](#pivotlayout-对象)
      - [PivotLine 对象](#pivotline-对象)
      - [PivotLineCells 对象](#pivotlinecells-对象)
      - [PivotLines 对象](#pivotlines-对象)
      - [PivotTable 对象](#pivottable-对象)
      - [PivotTables 对象](#pivottables-对象)
      - [PlotArea 对象](#plotarea-对象)
      - [Point 对象](#point-对象)
      - [Points 对象](#points-对象)
      - [Protection 对象](#protection-对象)
      - [Range 对象](#range-对象)
      - [Ranges 对象](#ranges-对象)
      - [RectangularGradient 对象](#rectangulargradient-对象)
      - [Series 对象](#series-对象)
      - [SeriesCollection 对象](#seriescollection-对象)
      - [SeriesLines 对象](#serieslines-对象)
      - [ShadowFormat 对象](#shadowformat-对象)
      - [Shape 对象](#shape-对象)
      - [ShapeNodes 对象](#shapenodes-对象)
      - [ShapeRange 对象](#shaperange-对象)
      - [Shapes 对象](#shapes-对象)
      - [SheetViews 对象](#sheetviews-对象)
      - [Sheets 对象](#sheets-对象)
      - [Slicer 对象](#slicer-对象)
      - [SlicerCache 对象](#slicercache-对象)
      - [SlicerCaches 对象](#slicercaches-对象)
      - [SlicerItem 对象](#sliceritem-对象)
      - [SlicerItems 对象](#sliceritems-对象)
      - [SlicerPivotTables 对象](#slicerpivottables-对象)
      - [Slicers 对象](#slicers-对象)
      - [Sort 对象](#sort-对象)
      - [SortField 对象](#sortfield-对象)
      - [SortFields 对象](#sortfields-对象)
      - [SparkAxes 对象](#sparkaxes-对象)
      - [SparkColor 对象](#sparkcolor-对象)
      - [SparkHorizontalAxis 对象](#sparkhorizontalaxis-对象)
      - [SparkPoints 对象](#sparkpoints-对象)
      - [Sparkline 对象](#sparkline-对象)
      - [SparklineGroup 对象](#sparklinegroup-对象)
      - [SparklineGroups 对象](#sparklinegroups-对象)
      - [SpellingOptions 对象](#spellingoptions-对象)
      - [Style 对象](#style-对象)
      - [Styles 对象](#styles-对象)
      - [TableStyle 对象](#tablestyle-对象)
      - [TableStyleElement 对象](#tablestyleelement-对象)
      - [TableStyleElements 对象](#tablestyleelements-对象)
      - [TableStyles 对象](#tablestyles-对象)
      - [TextEffectFormat 对象](#texteffectformat-对象)
      - [TextFrame 对象](#textframe-对象)
      - [TextFrame2 对象](#textframe2-对象)
      - [ThreeDFormat 对象](#threedformat-对象)
      - [TickLabels 对象](#ticklabels-对象)
      - [Top10 对象](#top10-对象)
      - [Trendline 对象](#trendline-对象)
      - [Trendlines 对象](#trendlines-对象)
      - [UniqueValues 对象](#uniquevalues-对象)
      - [UpBars 对象](#upbars-对象)
      - [UserAccess 对象](#useraccess-对象)
      - [UserAccessList 对象](#useraccesslist-对象)
      - [Validation 对象](#validation-对象)
      - [Workbook 对象](#workbook-对象)
      - [Worksheet 对象](#worksheet-对象)
      - [WorksheetFunction 对象](#worksheetfunction-对象)
      - [Worksheets 对象](#worksheets-对象)
      - [XlAboveBelow 枚举](#xlabovebelow-枚举)
      - [XlActionType 枚举](#xlactiontype-枚举)
      - [XlAllocation 枚举](#xlallocation-枚举)
      - [XlAllocationMethod 枚举](#xlallocationmethod-枚举)
      - [XlAllocationValue 枚举](#xlallocationvalue-枚举)
      - [XlApplicationInternational 枚举](#xlapplicationinternational-枚举)
      - [XlApplyNamesOrder 枚举](#xlapplynamesorder-枚举)
      - [XlArabicModes 枚举](#xlarabicmodes-枚举)
      - [XlArrangeStyle 枚举](#xlarrangestyle-枚举)
      - [XlArrowHeadLength 枚举](#xlarrowheadlength-枚举)
      - [XlArrowHeadStyle 枚举](#xlarrowheadstyle-枚举)
      - [XlArrowHeadWidth 枚举](#xlarrowheadwidth-枚举)
      - [XlAutoFillType 枚举](#xlautofilltype-枚举)
      - [XlAutoFilterOperator 枚举](#xlautofilteroperator-枚举)
      - [XlBackground 枚举](#xlbackground-枚举)
      - [XlBordersIndex 枚举](#xlbordersindex-枚举)
      - [XlBuiltInDialog 枚举](#xlbuiltindialog-枚举)
      - [XlCVError 枚举](#xlcverror-枚举)
      - [XlCalcFor 枚举](#xlcalcfor-枚举)
      - [XlCalculatedMemberType 枚举](#xlcalculatedmembertype-枚举)
      - [XlCalculation 枚举](#xlcalculation-枚举)
      - [XlCalculationInterruptKey 枚举](#xlcalculationinterruptkey-枚举)
      - [XlCalculationState 枚举](#xlcalculationstate-枚举)
      - [XlCellChangedState 枚举](#xlcellchangedstate-枚举)
      - [XlCellInsertionMode 枚举](#xlcellinsertionmode-枚举)
      - [XlCellType 枚举](#xlcelltype-枚举)
      - [XlChartGallery 枚举](#xlchartgallery-枚举)
      - [XlChartLocation 枚举](#xlchartlocation-枚举)
      - [XlChartPicturePlacement 枚举](#xlchartpictureplacement-枚举)
      - [XlChartType 枚举](#xlcharttype-枚举)
      - [XlCheckInVersionType 枚举](#xlcheckinversiontype-枚举)
      - [XlClipboardFormat 枚举](#xlclipboardformat-枚举)
      - [XlCmdType 枚举](#xlcmdtype-枚举)
      - [XlColumnDataType 枚举](#xlcolumndatatype-枚举)
      - [XlCommandUnderlines 枚举](#xlcommandunderlines-枚举)
      - [XlCommentDisplayMode 枚举](#xlcommentdisplaymode-枚举)
      - [XlConditionValueTypes 枚举](#xlconditionvaluetypes-枚举)
      - [XlConnectionType 枚举](#xlconnectiontype-枚举)
      - [XlConsolidationFunction 枚举](#xlconsolidationfunction-枚举)
      - [XlContainsOperator 枚举](#xlcontainsoperator-枚举)
      - [XlCopyPictureFormat 枚举](#xlcopypictureformat-枚举)
      - [XlCorruptLoad 枚举](#xlcorruptload-枚举)
      - [XlCreator 枚举](#xlcreator-枚举)
      - [XlCredentialsMethod 枚举](#xlcredentialsmethod-枚举)
      - [XlCubeFieldSubType 枚举](#xlcubefieldsubtype-枚举)
      - [XlCubeFieldType 枚举](#xlcubefieldtype-枚举)
      - [XlCutCopyMode 枚举](#xlcutcopymode-枚举)
      - [XlDVAlertStyle 枚举](#xldvalertstyle-枚举)
      - [XlDVType 枚举](#xldvtype-枚举)
      - [XlDataBarAxisPosition 枚举](#xldatabaraxisposition-枚举)
      - [XlDataBarBorderType 枚举](#xldatabarbordertype-枚举)
      - [XlDataBarFillType 枚举](#xldatabarfilltype-枚举)
      - [XlDataBarNegativeColorType 枚举](#xldatabarnegativecolortype-枚举)
      - [XlDataLabelSeparator 枚举](#xldatalabelseparator-枚举)
      - [XlDataSeriesDate 枚举](#xldataseriesdate-枚举)
      - [XlDataSeriesType 枚举](#xldataseriestype-枚举)
      - [XlDeleteShiftDirection 枚举](#xldeleteshiftdirection-枚举)
      - [XlDirection 枚举](#xldirection-枚举)
      - [XlDisplayDrawingObjects 枚举](#xldisplaydrawingobjects-枚举)
      - [XlDisplayUnit 枚举](#xldisplayunit-枚举)
      - [XlDupeUnique 枚举](#xldupeunique-枚举)
      - [XlDynamicFilterCriteria 枚举](#xldynamicfiltercriteria-枚举)
      - [XlEditionFormat 枚举](#xleditionformat-枚举)
      - [XlEditionOptionsOption 枚举](#xleditionoptionsoption-枚举)
      - [XlEditionType 枚举](#xleditiontype-枚举)
      - [XlEnableCancelKey 枚举](#xlenablecancelkey-枚举)
      - [XlEnableSelection 枚举](#xlenableselection-枚举)
      - [XlErrorBarDirection 枚举](#xlerrorbardirection-枚举)
      - [XlErrorChecks 枚举](#xlerrorchecks-枚举)
      - [XlFileAccess 枚举](#xlfileaccess-枚举)
      - [XlFileFormat 枚举](#xlfileformat-枚举)
      - [XlFileValidationPivotMode 枚举](#xlfilevalidationpivotmode-枚举)
      - [XlFillWith 枚举](#xlfillwith-枚举)
      - [XlFilterAction 枚举](#xlfilteraction-枚举)
      - [XlFilterAllDatesInPeriod 枚举](#xlfilteralldatesinperiod-枚举)
      - [XlFindLookIn 枚举](#xlfindlookin-枚举)
      - [XlFixedFormatQuality 枚举](#xlfixedformatquality-枚举)
      - [XlFixedFormatType 枚举](#xlfixedformattype-枚举)
      - [XlFormControl 枚举](#xlformcontrol-枚举)
      - [XlFormatConditionOperator 枚举](#xlformatconditionoperator-枚举)
      - [XlFormatConditionType 枚举](#xlformatconditiontype-枚举)
      - [XlFormatFilterTypes 枚举](#xlformatfiltertypes-枚举)
      - [XlFormulaLabel 枚举](#xlformulalabel-枚举)
      - [XlGenerateTableRefs 枚举](#xlgeneratetablerefs-枚举)
      - [XlGradientFillType 枚举](#xlgradientfilltype-枚举)
      - [XlHebrewModes 枚举](#xlhebrewmodes-枚举)
      - [XlHighlightChangesTime 枚举](#xlhighlightchangestime-枚举)
      - [XlHtmlType 枚举](#xlhtmltype-枚举)
      - [XlIMEMode 枚举](#xlimemode-枚举)
      - [XlIcon 枚举](#xlicon-枚举)
      - [XlIconSet 枚举](#xliconset-枚举)
      - [XlImportDataAs 枚举](#xlimportdataas-枚举)
      - [XlInsertFormatOrigin 枚举](#xlinsertformatorigin-枚举)
      - [XlInsertShiftDirection 枚举](#xlinsertshiftdirection-枚举)
      - [XlLayoutFormType 枚举](#xllayoutformtype-枚举)
      - [XlLayoutRowType 枚举](#xllayoutrowtype-枚举)
      - [XlLineStyle 枚举](#xllinestyle-枚举)
      - [XlLink 枚举](#xllink-枚举)
      - [XlLinkInfo 枚举](#xllinkinfo-枚举)
      - [XlLinkInfoType 枚举](#xllinkinfotype-枚举)
      - [XlLinkStatus 枚举](#xllinkstatus-枚举)
      - [XlLinkType 枚举](#xllinktype-枚举)
      - [XlListConflict 枚举](#xllistconflict-枚举)
      - [XlListDataType 枚举](#xllistdatatype-枚举)
      - [XlListObjectSourceType 枚举](#xllistobjectsourcetype-枚举)
      - [XlLocationInTable 枚举](#xllocationintable-枚举)
      - [XlLookAt 枚举](#xllookat-枚举)
      - [XlLookFor 枚举](#xllookfor-枚举)
      - [XlMSApplication 枚举](#xlmsapplication-枚举)
      - [XlMailSystem 枚举](#xlmailsystem-枚举)
      - [XlMeasurementUnits 枚举](#xlmeasurementunits-枚举)
      - [XlMouseButton 枚举](#xlmousebutton-枚举)
      - [XlMousePointer 枚举](#xlmousepointer-枚举)
      - [XlOLEType 枚举](#xloletype-枚举)
      - [XlOLEVerb 枚举](#xloleverb-枚举)
      - [XlOartHorizontalOverflow 枚举](#xloarthorizontaloverflow-枚举)
      - [XlOartVerticalOverflow 枚举](#xloartverticaloverflow-枚举)
      - [XlObjectSize 枚举](#xlobjectsize-枚举)
      - [XlOrder 枚举](#xlorder-枚举)
      - [XlOrientation 枚举](#xlorientation-枚举)
      - [XlPTSelectionMode 枚举](#xlptselectionmode-枚举)
      - [XlPageBreak 枚举](#xlpagebreak-枚举)
      - [XlPageBreakExtent 枚举](#xlpagebreakextent-枚举)
      - [XlPageOrientation 枚举](#xlpageorientation-枚举)
      - [XlPaperSize 枚举](#xlpapersize-枚举)
      - [XlParameterDataType 枚举](#xlparameterdatatype-枚举)
      - [XlParameterType 枚举](#xlparametertype-枚举)
      - [XlPasteSpecialOperation 枚举](#xlpastespecialoperation-枚举)
      - [XlPasteType 枚举](#xlpastetype-枚举)
      - [XlPattern 枚举](#xlpattern-枚举)
      - [XlPhoneticAlignment 枚举](#xlphoneticalignment-枚举)
      - [XlPhoneticCharacterType 枚举](#xlphoneticcharactertype-枚举)
      - [XlPictureAppearance 枚举](#xlpictureappearance-枚举)
      - [XlPictureConvertorType 枚举](#xlpictureconvertortype-枚举)
      - [XlPivotCellType 枚举](#xlpivotcelltype-枚举)
      - [XlPivotConditionScope 枚举](#xlpivotconditionscope-枚举)
      - [XlPivotFieldCalculation 枚举](#xlpivotfieldcalculation-枚举)
      - [XlPivotFieldDataType 枚举](#xlpivotfielddatatype-枚举)
      - [XlPivotFieldRepeatLabels 枚举](#xlpivotfieldrepeatlabels-枚举)
      - [XlPivotFilterType 枚举](#xlpivotfiltertype-枚举)
      - [XlPivotFormatType 枚举](#xlpivotformattype-枚举)
      - [XlPivotLineType 枚举](#xlpivotlinetype-枚举)
      - [XlPivotTableMissingItems 枚举](#xlpivottablemissingitems-枚举)
      - [XlPivotTableSourceType 枚举](#xlpivottablesourcetype-枚举)
      - [XlPivotTableVersionList 枚举](#xlpivottableversionlist-枚举)
      - [XlPlacement 枚举](#xlplacement-枚举)
      - [XlPlatform 枚举](#xlplatform-枚举)
      - [XlPortugueseReform 枚举](#xlportuguesereform-枚举)
      - [XlPrintErrors 枚举](#xlprinterrors-枚举)
      - [XlPrintLocation 枚举](#xlprintlocation-枚举)
      - [XlPriority 枚举](#xlpriority-枚举)
      - [XlPropertyDisplayedIn 枚举](#xlpropertydisplayedin-枚举)
      - [XlProtectedViewCloseReason 枚举](#xlprotectedviewclosereason-枚举)
      - [XlProtectedViewWindowState 枚举](#xlprotectedviewwindowstate-枚举)
      - [XlQueryType 枚举](#xlquerytype-枚举)
      - [XlRangeAutoFormat 枚举](#xlrangeautoformat-枚举)
      - [XlRangeValueDataType 枚举](#xlrangevaluedatatype-枚举)
      - [XlReferenceStyle 枚举](#xlreferencestyle-枚举)
      - [XlReferenceType 枚举](#xlreferencetype-枚举)
      - [XlRemoveDocInfoType 枚举](#xlremovedocinfotype-枚举)
      - [XlRgbColor 枚举](#xlrgbcolor-枚举)
      - [XlRobustConnect 枚举](#xlrobustconnect-枚举)
      - [XlRoutingSlipDelivery 枚举](#xlroutingslipdelivery-枚举)
      - [XlRoutingSlipStatus 枚举](#xlroutingslipstatus-枚举)
      - [XlRunAutoMacro 枚举](#xlrunautomacro-枚举)
      - [XlSaveAction 枚举](#xlsaveaction-枚举)
      - [XlSaveAsAccessMode 枚举](#xlsaveasaccessmode-枚举)
      - [XlSaveConflictResolution 枚举](#xlsaveconflictresolution-枚举)
      - [XlSearchDirection 枚举](#xlsearchdirection-枚举)
      - [XlSearchOrder 枚举](#xlsearchorder-枚举)
      - [XlSearchWithin 枚举](#xlsearchwithin-枚举)
      - [XlSheetType 枚举](#xlsheettype-枚举)
      - [XlSheetVisibility 枚举](#xlsheetvisibility-枚举)
      - [XlSlicerCrossFilterType 枚举](#xlslicercrossfiltertype-枚举)
      - [XlSlicerSort 枚举](#xlslicersort-枚举)
      - [XlSmartTagControlType 枚举](#xlsmarttagcontroltype-枚举)
      - [XlSmartTagDisplayMode 枚举](#xlsmarttagdisplaymode-枚举)
      - [XlSortDataOption 枚举](#xlsortdataoption-枚举)
      - [XlSortMethod 枚举](#xlsortmethod-枚举)
      - [XlSortMethodOld 枚举](#xlsortmethodold-枚举)
      - [XlSortOn 枚举](#xlsorton-枚举)
      - [XlSortOrder 枚举](#xlsortorder-枚举)
      - [XlSortOrientation 枚举](#xlsortorientation-枚举)
      - [XlSortType 枚举](#xlsorttype-枚举)
      - [XlSourceType 枚举](#xlsourcetype-枚举)
      - [XlSpanishModes 枚举](#xlspanishmodes-枚举)
      - [XlSparkScale 枚举](#xlsparkscale-枚举)
      - [XlSparkType 枚举](#xlsparktype-枚举)
      - [XlSparklineRowCol 枚举](#xlsparklinerowcol-枚举)
      - [XlSpeakDirection 枚举](#xlspeakdirection-枚举)
      - [XlSpecialCellsValue 枚举](#xlspecialcellsvalue-枚举)
      - [XlStdColorScale 枚举](#xlstdcolorscale-枚举)
      - [XlSubscribeToFormat 枚举](#xlsubscribetoformat-枚举)
      - [XlSubtototalLocationType 枚举](#xlsubtototallocationtype-枚举)
      - [XlSummaryColumn 枚举](#xlsummarycolumn-枚举)
      - [XlSummaryReportType 枚举](#xlsummaryreporttype-枚举)
      - [XlSummaryRow 枚举](#xlsummaryrow-枚举)
      - [XlTabPosition 枚举](#xltabposition-枚举)
      - [XlTableStyleElementType 枚举](#xltablestyleelementtype-枚举)
      - [XlTextParsingType 枚举](#xltextparsingtype-枚举)
      - [XlTextQualifier 枚举](#xltextqualifier-枚举)
      - [XlTextVisualLayoutType 枚举](#xltextvisuallayouttype-枚举)
      - [XlThemeColor 枚举](#xlthemecolor-枚举)
      - [XlThemeFont 枚举](#xlthemefont-枚举)
      - [XlThreadMode 枚举](#xlthreadmode-枚举)
      - [XlTimePeriods 枚举](#xltimeperiods-枚举)
      - [XlToolbarProtection 枚举](#xltoolbarprotection-枚举)
      - [XlTopBottom 枚举](#xltopbottom-枚举)
      - [XlTotalsCalculation 枚举](#xltotalscalculation-枚举)
      - [XlUpdateLinks 枚举](#xlupdatelinks-枚举)
      - [XlWBATemplate 枚举](#xlwbatemplate-枚举)
      - [XlWebFormatting 枚举](#xlwebformatting-枚举)
      - [XlWebSelectionType 枚举](#xlwebselectiontype-枚举)
      - [XlWindowState 枚举](#xlwindowstate-枚举)
      - [XlWindowType 枚举](#xlwindowtype-枚举)
      - [XlWindowView 枚举](#xlwindowview-枚举)
      - [XlXLMMacroType 枚举](#xlxlmmacrotype-枚举)
      - [XlXmlExportResult 枚举](#xlxmlexportresult-枚举)
      - [XlXmlImportResult 枚举](#xlxmlimportresult-枚举)
      - [XlXmlLoadOption 枚举](#xlxmlloadoption-枚举)
      - [XlYesNoGuess 枚举](#xlyesnoguess-枚举)
    - [数据表](#数据表)
      - [待开放](#待开放)
  - [高级服务](#高级服务)
    - [云文档 API](#云文档-api)
    - [概述](#概述)
    - [网络 API](#网络-api)

# 产品介绍

## 概述

# [AirScript 概述​](#airscript-概述)

金山文档 AirScript 是一个简单快速的轻量级脚本应用开发平台，它基于云技术构建，可让您快速轻松地创建与金山文档 Office 文件交互的业务应用。

我们在金山文档的组件中提供了代码编辑器，您的代码将会安全地运行在我们的服务端上。您可以使用现代 JavaScript 语言编写逻辑代码，并可以调用 AirScript 内置强大的组件（表格）API 来构建特定场景的解决方案。

目前在国内业界中尚未出现类似的产品，金山文档 AirScript 是国内首家提供该服务的平台。

## [AirScript 能做什么​](#airscript-能做什么)

AirScript 目前主要为在线表格、多维表打造二开平台，通过编程的方式，提供对表格数据的增删查改、单元格式修改、属性设置等能力。

#### [工具优势​](#工具优势)

无需搭建本地环境，直接在文档内进行脚本云开发。
内置定制化的全局 Application 对象，编辑器智能提示，开发、调试、运行一条龙服务。
同步获取属性，同步执行方法，减少传统的异步调用带来的心智负担。
得益于集成化开发环境，无论是创建定时任务，还是批量处理数据，亦或是自动化生成文档，开发者可以在这里尽情发挥自己的想象力。
## [如何使用 AirScript​](#如何使用-airscript)

打开在线表格 KSheet，切换至「效率」Tab，在下方二级工具栏找到「AirScript 编辑工具」，点击即可调起 AirScript 编辑器，如图：

提示

有文件编辑权限的协作者才能打开开发工具

### [编辑器功能介绍​](#编辑器功能介绍)

脚本编辑器分左右两部分，左边部分为脚本文件管理区域；右边部分为编码及运行区域。功能上提供脚本的增删改、运行及运行时生成提示日志的能力。

文档操作 API 方面，编辑器内置了 Application 对象语法提示，可直接根据文档及语法提示书写需要的 OpenApi。

具体功能介绍：

脚本管理
编辑器左侧及右侧顶部操作区域，提供代码文件的新建、保存、删除、运行等常规操作；
脚本文件存储在云端，并分为【我的脚本】和【文档共享脚本】，其中我的脚本跟随账号，文档共享脚本跟随当前文档；
代码编辑
代码编辑器基于开源
monaco editor
进行二开，用户操作基本类似
VS Code
；
代码保存
表格：代码自动保存在本地，但不自动同步云端，需要用户主动点击「保存」或者「Ctrl+S」才会同步云端，未同步的代码文件会在文件名前方展示一个「小绿点」；
多维表：多维表切换脚本及关闭开发工具会自动保存代码并同步云端；
执行日志
显示代码的执行信息，用户可根据返回信息查看脚本的执行状态以及打印结果；
如果代码报错，将会打印错误信息及位置，点击位置编辑器会定位到错误代码块。

## 脚本语言

# [脚本语言​](#脚本语言)

AirScript脚本采用标准JavaScript语言进行编写，支持大部分ES6语法。

## [内置全局对象​](#内置全局对象)

#### [文档OpenApi对象Application​](#文档openapi对象-application)

Application为文档OpenApi对象，基于不同的金山文档类型，会生成不同的Application对象属性，在代码编辑器内可直接引用。

#### [脚本上下文对象Context​](#脚本上下文对象-context)

Context为脚本运行上下文对象。负责挂载一些与脚本运行相关的属性，如视图、配置等产生的相关内容，在代码编辑器内可直接引用，以方便整体应用的构造。

## [不支持的语法​](#不支持的语法)

### [Class​](#class)

javascript
```javascript
class C{
    private a; // Unexpected token
    public b(){
        return this.a;
    }
}
```

### [Object里直接定义方法​](#object里直接定义方法)

javascript
```javascript
let obj = {
    sayHello(){ //Unexpected token 
        console.log("hello")
    }
}
obj.sayHello()
```

### [import、export​](#import、export)

javascript
```javascript
import os from 'os' // SyntaxError
export default os // SyntaxError
```

### ?.​

javascript
```javascript
function f(argv){
    console.log(argv?.toString()) // Unexpected token
}
```

### [await​](#await)

javascript
```javascript
async function f() {
    await new Promise((resolve)=>{ // Unexpected token
        resolve("")
    })
}
```

### [yield​](#yield)

javascript
```javascript
function* foo(index) {
  while (index < 2) {
    yield index; // Unexpected token
    index++;
  }
}

const iterator = foo(0);

console.log(iterator.next().value);

console.log(iterator.next().value);
```


# 快速上手

## 开始

# [开始​](#开始)

## [打开AirScript编辑器​](#打开airscript编辑器)

在金山文档首页新建一个表格并打开来体验AirScript。
打开表格之后，在上方
效率
-
AirScript编辑工具
弹出编辑页面。
将下方的例子，逐个运行，查看效果来快速上手AirScript。
## [HelloWorld​](#helloworld)

### [console.log()​](#console-log)

js
```js
console.log("hello world!")
```

### [console.error()​](#console-error)

js
```js
console.error("hello world!")
```

## [文档操作​](#文档操作)

### [遍历文件的SheetName，切换对应的Sheet，读取单元格，修改单元格​](#遍历文件的sheetname-切换对应的sheet-读取单元格-修改单元格)

js
```js
// 遍历并打印所有工作表的名称
let sheets = Application.Sheets
for (let i = 0; i < sheets.Count; i++) {
    let sheet = sheets.Item(i + 1)
    console.log(sheet.Name) // 打印每个工作表的名称
}

// 打印当前激活Sheet的名称
console.log(Application.ActiveSheet.Name) // 打印当前激活Sheet的名称

// 获取当前激活工作表的A1单元格
let A1 = Application.ActiveSheet.Range("A1")

// 打印A1单元格内容
console.log(A1.Text) // 打印A1单元格的内容

// 修改A1单元格内容
A1.Value2 = "bar"
console.log(A1.Text) // 打印修改后的A1单元格内容

// 修改A1单元格的背景颜色为黄色(255,255,0)
A1.Interior.Color = RGB(255,255,0) // 高亮A1单元格
```

### [将文档所有非空单元格增加指定后缀​](#将文档所有非空单元格增加指定后缀)

js
```js
// 通过Application.ActiveSheet.UsedRange获得用户使用区域
const usedRange = Application.ActiveSheet.UsedRange
const startRow = usedRange.Row // 获取起始行
const startCol = usedRange.Column // 获取起始列
const endRow = startRow + usedRange.Rows.Count // 获取终止行
const endCol = startCol + usedRange.Columns.Count // 获取终止列

for (let i = startRow; i < endRow; i++) {
    const row = Application.ActiveSheet.Rows(i) // 确定行
    for (let j = startCol; j < endCol; j++) {
        const rg = row.Columns(j) // 从行对象中指定列，从而确定单元格
        const text = rg.Text
        // 如果单元格非空，则添加foo后缀
        if (text !== '') {
            rg.Value2 = text + 'foo' // 修改单元格内容
        }
    }
}

// 遍历UsedRange中的每个单元格并打印其内容
for (let i = startRow; i < endRow; i++) {
    for (let j = startCol; j < endCol; j++) {
        const cell = Application.ActiveSheet.Cells(i, j)
        console.log(cell.Text) // 打印所有单元格数据
    }
}
```

### [新建Sheet​](#新建sheet)

js
```js
// 当前工作表名称
const currentSheetName = Application.ActiveSheet.Name

// 在当前Sheet之后新增两个默认名称的工作表
Application.Sheets.Add(null, Application.Sheets(currentSheetName), 2)

// 在当前Sheet之后新增名称为'新工作表(右)'的Sheet
const newSheetRight = Application.Sheets.Add(null, Application.Sheets(currentSheetName), 1)
newSheetRight.Name = '新工作表(右)'

// 在当前Sheet之前新增名称为'新工作表(左)'的Sheet
const newSheetLeft = Application.Sheets.Add(Application.Sheets(currentSheetName), null, 1)
newSheetLeft.Name = '新工作表(左)'
```

## [更多​](#更多)

更多真实例子可通过示范案例尝试。


## 最佳实践

# [最佳实践​](#最佳实践)

## [体验优化​](#体验优化)

#### [通过超链接运行脚本​](#通过超链接运行脚本)

脚本的运行可以通过超链接触发，绑定方法如下

打开
插入
-
链接
弹出插入链接配置页面。
在文本输入框中填写合适的提示文本，描述该脚本的功能。
点击
类型
旁边的下拉框弹出选项选择
AirScript脚本
再指定想绑定的脚本。
点击
确定
。
点击插入的超链接即可运行脚本。

通过这种方式，用户所见即所得，相比于打开脚本编辑器点击运行友好很多。

#### [更符合直觉的运行体验​](#更符合直觉的运行体验)

脚本需要选定某些单元格范围时，有多种实现方式，通过Application.Selection读取并操作用户选中单元格是最符合用户直觉的运行体验。

下面是通过脚本变量划定遍历范围的例子。该例子每次修改范围都需要修改4个变量，非常不方便，体检较差。

javascript
```javascript
// 错误例子
const startRow = 1
const startColumn = 1
const endRow = 3
const endColumn = 3
for(let i = startRow; i <= endRow; i++){
  let row=Application.Rows(i)
  for (let j = startColumn; j <= endColumn; j++){
    console.log(row.Columns(j).Text)
  }
}
```

下面是通过Application.Selection获取用户在文档界面选中范围单元格，使用方便并符合直觉。

javascript
```javascript
// 正确例子
let selection = Application.Selection
let startRow = selection.Row
let startCol = selection.Column
let endRow = startRow + selection.Rows.Count - 1
let endCol = startCol + selection.Columns.Count - 1

for (let i = startRow; i <= endRow; i++) {
    let row = Application.Rows(i)
    for (let j = startCol; j <= endCol; j++) {
        console.log(row.Columns(j).Text)
    }
}
```

## [性能优化​](#性能优化)

#### [尽可能地复用对象​](#尽可能地复用对象)

对Application的每次函数调用会使用到脚本引擎去操作文件数据，重复调用会造成性能浪费。

如同样实现遍历并打印100*100单元格的内容，复用对象能使性能提升一倍。

无复用对象:

javascript
```javascript
// 错误例子
let start = new Date()
for(let i= 1;i<=100;i++){
  for (let j =1;j<=100;j++){
    Application.Rows(i).Columns(j).Text
  }
}
console.log(new Date()-start,'ms')
```

更好的写法:

javascript
```javascript
// 正确例子
let start = new Date()
for(let i= 1;i<=100;i++){
  let row=Application.Rows(i)
  for (let j =1;j<=100;j++){
    row.Columns(j).Text
  }
}
console.log(new Date()-start)
```

#### [使用UsedRange缩小遍历范围​](#使用usedrange缩小遍历范围)

上面更符合直觉的运行体验的例子中， 全选整个表格则会进行上万亿次遍历,即使只选中一整行也需要进行上百万次遍历， 但实际上我们单元表中使用到范围并没有那么大。 配合使用Application.ActiveSheet.UsedRange可以确认工作簿的使用范围，因此大大加快脚本执行时间。

下面是该例子的改进版本：

javascript
```javascript
const selection = Application.Selection
const usedRange = Application.ActiveSheet.UsedRange

// 确定遍历的范围
const rowFrom = Math.max(selection.Row, usedRange.Row)
const rowTo = Math.min(selection.Row + selection.Rows.Count - 1, usedRange.Row + usedRange.Rows.Count - 1)
const colFrom = Math.max(selection.Column, usedRange.Column)
const colTo = Math.min(selection.Column + selection.Columns.Count - 1, usedRange.Column + usedRange.Columns.Count - 1)

for (let i = rowFrom; i <= rowTo; i++) {
    const row = Application.ActiveSheet.Rows(i) // 复用对象
    for (let j = colFrom; j <= colTo; j++) {
        console.log(row.Columns(j).Text)
    }
}
```


## 配置视图

# [配置视图​](#配置视图)

为方便用户将写的脚本交付给无开发经验的第三方使用，我们推出了参数可视化配置。用户通过点击脚本工具栏视图配置，可以打开参数编排工具。通过配置form表单，结合脚本便可生成快捷操作工具。

## [视图与脚本的关联​](#视图与脚本的关联)

视图配置的表单参数，会在脚本编辑器中将参数挂载在
Context.argv
下，编程的过程中，可直接引用此变量。
Context
作为全局变量，类似
Application
。
视图配置的参数会实时在代码编辑器内生成语法提示，输入Context.argv，就可看到已配置的参数。
预览视图，点击视图内
运行脚本
按钮，将携带表单输入参数运行脚本，同时打印日志复用代码编辑器的日志框。
## [简单使用案例​](#简单使用案例)

#### [1、配置参数​](#_1、配置参数)

打开视图配置，添加以下参数：

| 参数名称 | 变量名称 | 变量类型 | 输入提示 |
| --- | --- | --- | --- |
| 选择表 | sheet | string | 输入表名：“sheet1” |
| 选择单元格 | range | string | 输入单元格：“A1” |
| 输入设置值 | value | string | 输入要设置的值 |

#### [2.输入代码​](#_2-输入代码)

js
```js
Application.Sheets.SelectByName(Context.argv.sheet).Activate()
Application.Range(Context.argv.range).Value=Context.argv.value
```

#### [3.编辑效果​](#_3-编辑效果)

1、2操作完成后，编辑效果如下图：

#### [4.运行效果​](#_4-运行效果)

点击预览，测试输入值，运行脚本，效果如下图：


# 脚本令牌

## 应用场景

# [应用场景​](#应用场景)

在脚本令牌的加持下，智能表格的强大能力得到完美释放，开发者能以一种更为高效和准确的方式执行任务。无论是网页爬取、数据分析，还是自动化流程，我都可以借助脚本令牌来完成。

如下为我们根据实际场景写的一些示例和说明，希望能给开发者一定的启发。

## [1. 私密信息查询​](#_1-私密信息查询)

现代社会越来越重视个人的隐私。这种趋势在很多方面都有所体现。

首先，在教育领域，许多学校开始加强对学生在校期间的信息保护，禁止将学生的个人信息出售或分享给他人。这包括学生的家庭信息、教育记录、考试成绩等。

其次，在医疗保健领域，病人的隐私保护成为了一个重要的问题。医生和其他医疗工作者需要遵守严格的隐私规定，确保病人的个人信息不会被泄露。

此外，在社交媒体领域，许多平台也开始加强对用户信息的保护。他们采取了更严格的数据安全措施，以确保用户的数据不会被泄露或滥用。

如下，我们将展示一个学生成绩查询的在线示例，你可以在线体验利用脚本令牌实现的私密信息查询的功能。现有学生成绩表如下图所示：

你可以根据上表提供的学生信息，输入对应学生的学号和姓名后，即可查询到对应学生的成绩。

注意

真实使用场景中，会要求输入密码或者手机验证等。

### 成绩查询系统

立即查询
上述在线示例的底层实现逻辑就是利用了脚本令牌执行 AirScript 脚本，在表格内查询到数据后返回给界面显示，脚本示例代码如下：

js
```js
// 获取学生的学号
const student_id = '1'
// 获取学生的姓名
const student_name = '金小妹'
// 获取已使用区域的行数
const rowCount = Application.ActiveSheet.UsedRange.Rows.Count

// 获取匹配的学生数据
const data = []
for (let i = 1; i <= rowCount; i++) {
  const studentId = Application.ActiveSheet.Range(`A${i}`).Value2
  const studentName = Application.ActiveSheet.Range(`B${i}`).Value2

  if (studentId === student_id && studentName === student_name) {
    const sex = Application.ActiveSheet.Range(`C${i}`).Value2
    const className = Application.ActiveSheet.Range(`D${i}`).Value2
    const language = Application.ActiveSheet.Range(`E${i}`).Value2
    const math = Application.ActiveSheet.Range(`F${i}`).Value2
    const english = Application.ActiveSheet.Range(`G${i}`).Value2
    const literature = Application.ActiveSheet.Range(`H${i}`).Value2
    const total = Application.ActiveSheet.Range(`I${i}`).Value2
    const rate = Application.ActiveSheet.Range(`J${i}`).Value2

    data.push({
      id: studentId,
      name: studentName,
      sex: sex,
      className: className,
      language: language,
      math: math,
      english: english,
      literature: literature,
      total: total,
      rate: rate
    })
    break
  }
}

// 返回匹配的学生数据
return data
```

## [2. 电商数据同步​](#_2-电商数据同步)

电商数据同步是一个重要的环节，确保电商平台在销售和运营方面能够高效运作。在数据来源上，它可能是多个不同的系统，包括数据库、ERP系统、CRM系统等。而开发者首先要做的是手动导出或者 webhook 的形式获取到商品信息、订单信息、物流信息等数据，然后在您的个人服务器内得到这些数据，进行数据清洗和转换，最后再通过脚本令牌写入到金山文档智能表格内，完成数据的同步。

## [3. RPA 数据同步​](#_3-rpa-数据同步)

如果开发者有自己的 RPA 平台，需要从金山文档内获取值班信息，再通过 RPA 发送到他们的工作群内。传统的实现方式很麻烦，需要先利用 AirScript 的邮件服务将内容发送到邮件里，然后新建定时任务定时执行脚本，最后通过 RPA 读取邮件内容后转发到企业微信。

有了脚本令牌后，再也不用这么辛苦的“曲线救国”了，想要什么数据，通过脚本令牌直线获取即可。

## [4. 简易数据库​](#_4-简易数据库)

一些开发者有自己的个人网站，用户量很少或者仅作为个人学习使用，直接购买云数据库成本太高不划算。

此时大家可以想想，数据库里的表是表，智能表格里的表也是表，在某种条件下，有没有可能智能表格可以平替掉云数据库？

当然有可能，使用脚本令牌您可以轻松的完成数据的增删改查，扔掉老爷车 SQL，使用 JavaScript 来进行“为所欲为”的结构化查询，快来体验一下吧。

注意

以上仅提供一个新思路，只适用于个人学习或者非常轻量级的服务，毕竟玩归玩闹归闹，别拿数据开玩笑。

## [5. 作为数据工具使用​](#_5-作为数据工具使用)

有时候我们需要在自己的系统内使用到表格提供的高级功能，来完成对数据的筛选和过滤操作。比如现在有这么一个场景：开发者使用影刀 RPA 进行一个网站的数据爬取，爬取完了之后存到一个 excel，对 excel 做完数据清理之后才能进行下一步操作。但是有了脚本令牌后，就可以先将数据写入金山文档表格中，然后执行 AirScript 基于规则进行一个自动清理，直接就能进行下一步操作。


## 接口说明

# [接口说明​](#接口说明)

成功生成脚本令牌后，就可以通过 HTTP 接口执行脚本了，我们提供了同步执行和异步执行两种脚本执行接口供开发者使用。

相较而言，前者使用更简单，接口调用后会直接返回执行结果，适用于执行耗时一般的场景；而后者则略微复杂一点，接口调用后不会返回最终的执行的结果，但会立即返回一个task_id，您需要根据此task_id轮询脚本执行的日志，而无需同步等待结果阻塞业务流程，该接口适用于执行耗时比较大的场景。

无论使用您使用哪个接口，都必须先获取到文件 ID 和脚本 ID，请先进入脚本编辑器，在侧边栏列表的更多菜单里复制 webhook 链接即可。

## [同步执行脚本​](#同步执行脚本)

POST /api/v3/ide/file/:file_id/script/:script_id/sync_task

### [Header 参数​](#header2)

| 参数 | 必须 | 类型 | 说明 |
| --- | --- | --- | --- |
| Content-Type | 是 | string | application/json |
| AirScript-Token | 是 | string | 传入您通过 AirScript 编辑器生成的脚本令牌（APIToken） |

### [path 参数​](#path2)

| 参数 | 必须 | 类型 | 说明 |
| --- | --- | --- | --- |
| script_id | 是 | string | 脚本的 ID |
| file_id | 是 | string | 运行脚本的文件 ID |

### [body 参数​](#body2)

| 参数 | 必须 | 类型 | 说明 |
| --- | --- | --- | --- |
| Context | 是 | Object | 运行时的上下文参数 |
| Context.argv | 否 | Object | 传入的上下文参数对象，比如传入{name: 'xiaomeng', age: 18}，在 AS 代码中可通过Context.argv.name获取到传入的值 |
| Context.sheet_name | 否 | string | et,ksheet 运行时所在表名 |
| Context.range | 否 | string | et,ksheet 运行时所在区域，例如$B$156 |
| Context.link_from | 否 | string | et,ksheet 点击超链接所在单元格 |
| Context.db_active_view | 否 | string | db 运行时所在 view 名 |
| Context.db_selection | 否 | string | db 运行时所在选区 |

### [返回参数​](#return2)

| 参数 | 必须 | 类型 | 说明 |
| --- | --- | --- | --- |
| data | 是 | Object | 任务执行数据对象 |
| data.result | 是 | string | 任务执行返回的数据 |
| data.logs | 是 | Array | 任务执行日志 |
| data.logs[i].filename | 是 | string | 执行文件的名称 |
| data.logs[i].timestamp | 是 | string | 执行时间 |
| data.logs[i].unix_time | 是 | number | 执行 unix 时间戳 |
| data.logs[i].level | 是 | string | 日志级别 |
| data.logs[i].args | 是 | string[] | 日志打印参数 |
| status | 是 | string | 任务是否执行完毕 |
| error | 是 | string | 任务执行错误信息 |
| error_details | 否 | object | 错误信息详情对象 |
| error_details.name | 否 | string | 错误信息名称 |
| error_details.msg | 否 | string | 错误信息 |
| error_details.stack | 否 | string[] | 错误信息栈 |
| error_details.unix_time | 否 | number | 错误信息 unix 时间 |

### [请求示例​](#request-example2)

Shell
```text
curl --request POST \
	--url https://www.kdocs.cn/api/v3/ide/file/:file_id/script/:script_id/sync_task \
	--header 'AirScript-Token: xxx' \
	--header 'Content-Type: application/json' \
	--data '{"Context":{"argv":{},"sheet_name":"表名","range":"$B$156"}}'
```

Java
```text
OkHttpClient client = new OkHttpClient();

MediaType mediaType = MediaType.parse("application/json");
RequestBody body = RequestBody.create(mediaType, "{\"Context\":{\"argv\":{},\"sheet_name\":\"表名\",\"range\":\"$B$156\"}}");
Request request = new Request.Builder()
	.url("https://www.kdocs.cn/api/v3/ide/file/:file_id/script/:script_id/sync_task")
	.post(body)
	.addHeader("Content-Type", "application/json")
	.addHeader("AirScript-Token", "xxx")
	.build();

Response response = client.newCall(request).execute();
```

Go
```text
package main

import (
	"fmt"
	"strings"
	"net/http"
	"io/ioutil"
)

func main() {

	url := "https://www.kdocs.cn/api/v3/ide/file/:file_id/script/:script_id/sync_task"

	payload := strings.NewReader("{\"Context\":{\"argv\":{},\"sheet_name\":\"表名\",\"range\":\"$B$156\"}}")

	req, _ := http.NewRequest("POST", url, payload)

	req.Header.Add("Content-Type", "application/json")
	req.Header.Add("AirScript-Token", "xxx")

	res, _ := http.DefaultClient.Do(req)

	defer res.Body.Close()
	body, _ := ioutil.ReadAll(res.Body)

	fmt.Println(res)
	fmt.Println(string(body))

}
```

Python
```text
import http.client

conn = http.client.HTTPSConnection("www.kdocs.cn")

payload = "{\"Context\":{\"argv\":{},\"sheet_name\":\"表名\",\"range\":\"$B$156\"}}"

headers = {
    'Content-Type': "application/json",
    'AirScript-Token': "xxx"
    }

conn.request("POST", "/api/v3/ide/file/:file_id/script/:script_id/sync_task", payload, headers)

res = conn.getresponse()
data = res.read()

print(data.decode("utf-8"))
```

PHP
```text
<?php

$curl = curl_init();

curl_setopt_array($curl, [
	CURLOPT_URL => "https://www.kdocs.cn/api/v3/ide/file/:file_id/script/:script_id/sync_task",
	CURLOPT_RETURNTRANSFER => true,
	CURLOPT_ENCODING => "",
	CURLOPT_MAXREDIRS => 10,
	CURLOPT_TIMEOUT => 30,
	CURLOPT_HTTP_VERSION => CURL_HTTP_VERSION_1_1,
	CURLOPT_CUSTOMREQUEST => "POST",
	CURLOPT_POSTFIELDS => "{\"Context\":{\"argv\":{},\"sheet_name\":\"表名\",\"range\":\"$B$156\"}}",
	CURLOPT_HTTPHEADER => [
		"AirScript-Token: xxx",
		"Content-Type: application/json"
	],
]);

$response = curl_exec($curl);
$err = curl_error($curl);

curl_close($curl);

if ($err) {
	echo "cURL Error #:" . $err;
} else {
	echo $response;
}
```

JS
```text
const data = JSON.stringify({
	"Context": {
		"argv": {},
		"sheet_name": "表名",
		"range": "$B$156"
	}
});

const xhr = new XMLHttpRequest();
xhr.withCredentials = true;

xhr.addEventListener("readystatechange", function () {
	if (this.readyState === this.DONE) {
		console.log(this.responseText);
	}
});

xhr.open("POST", "https://www.kdocs.cn/api/v3/ide/file/:file_id/script/:script_id/sync_task");
xhr.setRequestHeader("Content-Type", "application/json");
xhr.setRequestHeader("AirScript-Token", "xxx");

xhr.send(data);
```

Node.js
```text
const http = require("https");

const options = {
	"method": "POST",
	"hostname": "www.kdocs.cn",
	"port": null,
	"path": "/api/v3/ide/file/:file_id/script/:script_id/sync_task",
	"headers": {
		"Content-Type": "application/json",
		"AirScript-Token": "xxx"
	}
};

const req = http.request(options, function (res) {
	const chunks = [];

	res.on("data", function (chunk) {
		chunks.push(chunk);
	});

	res.on("end", function () {
		const body = Buffer.concat(chunks);
		console.log(body.toString());
	});
});

req.write(JSON.stringify({Context: {argv: {}, sheet_name: '表名', range: '$B$156'}}));
req.end();
```

C
```text
CURL *hnd = curl_easy_init();

curl_easy_setopt(hnd, CURLOPT_CUSTOMREQUEST, "POST");
curl_easy_setopt(hnd, CURLOPT_URL, "https://www.kdocs.cn/api/v3/ide/file/:file_id/script/:script_id/sync_task");

struct curl_slist *headers = NULL;
headers = curl_slist_append(headers, "Content-Type: application/json");
headers = curl_slist_append(headers, "AirScript-Token: xxx");
curl_easy_setopt(hnd, CURLOPT_HTTPHEADER, headers);

curl_easy_setopt(hnd, CURLOPT_POSTFIELDS, "{\"Context\":{\"argv\":{},\"sheet_name\":\"表名\",\"range\":\"$B$156\"}}");

CURLcode ret = curl_easy_perform(hnd);
```

C#
```text
var client = new RestClient("https://www.kdocs.cn/api/v3/ide/file/:file_id/script/:script_id/sync_task");
var request = new RestRequest(Method.POST);
request.AddHeader("Content-Type", "application/json");
request.AddHeader("AirScript-Token", "xxx");
request.AddParameter("application/json", "{\"Context\":{\"argv\":{},\"sheet_name\":\"表名\",\"range\":\"$B$156\"}}", ParameterType.RequestBody);
IRestResponse response = client.Execute(request);
```

### [返回示例​](#response-example2)

json
```json
{
  "data": {
    "logs": [
      {
        "filename": "<system>",
        "timestamp": "16:44:08.271",
        "unix_time": 1690274648271,
        "level": "info",
        "args": ["脚本环境初始化..."]
      },
      {
        "filename": "<system>",
        "timestamp": "16:44:08.953",
        "unix_time": 1690274648953,
        "level": "info",
        "args": ["已开始执行"]
      },
      {
        "filename": "未命名脚本.js:1:9",
        "timestamp": "16:44:08.968",
        "unix_time": 1690274648968,
        "level": "info",
        "args": ["打印参数A：111"]
      },
      {
        "filename": "<system>",
        "timestamp": "16:44:08.969",
        "unix_time": 1690274648969,
        "level": "info",
        "args": ["执行完毕"]
      }
    ],
    "result": "[Undefined]"
  },
  "error": "",
  "status": "finished"
}
```

## [异步执行脚本​](#异步执行脚本)

POST /api/v3/ide/file/:file_id/script/:script_id/task

### [Header 参数​](#header1)

| 参数 | 必须 | 类型 | 说明 |
| --- | --- | --- | --- |
| Content-Type | 是 | string | application/json |
| AirScript-Token | 是 | string | 传入您通过 AirScript 编辑器生成的脚本令牌（APIToken） |

### [path 参数​](#path1)

| 参数 | 必须 | 类型 | 说明 |
| --- | --- | --- | --- |
| script_id | 是 | string | 脚本的 ID |
| file_id | 是 | string | 运行脚本的文件 ID |

### [body 参数​](#body1)

| 参数 | 必须 | 类型 | 说明 |
| --- | --- | --- | --- |
| Context | 是 | Object | 运行时的上下文参数 |
| Context.argv | 否 | Object | 传入的上下文参数对象，比如传入{name: 'xiaomeng', age: 18}，在 AS 代码中可通过Context.argv.name获取到传入的值 |
| Context.sheet_name | 否 | string | et,ksheet 运行时所在表名 |
| Context.range | 否 | string | et,ksheet 运行时所在区域，例如$B$156 |
| Context.link_from | 否 | string | et,ksheet 点击超链接所在单元格 |
| Context.db_active_view | 否 | string | db 运行时所在 view 名 |
| Context.db_selection | 否 | string | db 运行时所在选区 |

### [返回参数​](#return1)

| 参数 | 必须 | 类型 | 说明 |
| --- | --- | --- | --- |
| task_id | 是 | string | 运行的任务 Id，用于轮循运行结果 |
| task_type | 是 | string | 任务类型 |

### [请求示例​](#request-example1)

Shell
```json
curl --request POST \
	--url https://www.kdocs.cn/api/v3/ide/file/:file_id/script/:script_id/task \
	--header 'AirScript-Token: xxx' \
	--header 'Content-Type: application/json' \
	--data '{"Context":{"argv":{},"sheet_name":"表名","range":"$B$156"}}'
```

Java
```json
OkHttpClient client = new OkHttpClient();

MediaType mediaType = MediaType.parse("application/json");
RequestBody body = RequestBody.create(mediaType, "{\"Context\":{\"argv\":{},\"sheet_name\":\"表名\",\"range\":\"$B$156\"}}");
Request request = new Request.Builder()
	.url("https://www.kdocs.cn/api/v3/ide/file/:file_id/script/:script_id/task")
	.post(body)
	.addHeader("Content-Type", "application/json")
	.addHeader("AirScript-Token", "xxx")
	.build();

Response response = client.newCall(request).execute();
```

Go
```json
package main

import (
	"fmt"
	"strings"
	"net/http"
	"io/ioutil"
)

func main() {

	url := "https://www.kdocs.cn/api/v3/ide/file/:file_id/script/:script_id/task"

	payload := strings.NewReader("{\"Context\":{\"argv\":{},\"sheet_name\":\"表名\",\"range\":\"$B$156\"}}")

	req, _ := http.NewRequest("POST", url, payload)

	req.Header.Add("Content-Type", "application/json")
	req.Header.Add("AirScript-Token", "xxx")

	res, _ := http.DefaultClient.Do(req)

	defer res.Body.Close()
	body, _ := ioutil.ReadAll(res.Body)

	fmt.Println(res)
	fmt.Println(string(body))

}
```

Python
```json
import http.client

conn = http.client.HTTPSConnection("www.kdocs.cn")

payload = "{\"Context\":{\"argv\":{},\"sheet_name\":\"表名\",\"range\":\"$B$156\"}}"

headers = {
    'Content-Type': "application/json",
    'AirScript-Token': "xxx"
    }

conn.request("POST", "/api/v3/ide/file/:file_id/script/:script_id/task", payload, headers)

res = conn.getresponse()
data = res.read()

print(data.decode("utf-8"))
```

PHP
```json
<?php

$curl = curl_init();

curl_setopt_array($curl, [
	CURLOPT_URL => "https://www.kdocs.cn/api/v3/ide/file/:file_id/script/:script_id/task",
	CURLOPT_RETURNTRANSFER => true,
	CURLOPT_ENCODING => "",
	CURLOPT_MAXREDIRS => 10,
	CURLOPT_TIMEOUT => 30,
	CURLOPT_HTTP_VERSION => CURL_HTTP_VERSION_1_1,
	CURLOPT_CUSTOMREQUEST => "POST",
	CURLOPT_POSTFIELDS => "{\"Context\":{\"argv\":{},\"sheet_name\":\"表名\",\"range\":\"$B$156\"}}",
	CURLOPT_HTTPHEADER => [
		"AirScript-Token: xxx",
		"Content-Type: application/json"
	],
]);

$response = curl_exec($curl);
$err = curl_error($curl);

curl_close($curl);

if ($err) {
	echo "cURL Error #:" . $err;
} else {
	echo $response;
}
```

JS
```json
const data = JSON.stringify({
	"Context": {
		"argv": {},
		"sheet_name": "表名",
		"range": "$B$156"
	}
});

const xhr = new XMLHttpRequest();
xhr.withCredentials = true;

xhr.addEventListener("readystatechange", function () {
	if (this.readyState === this.DONE) {
		console.log(this.responseText);
	}
});

xhr.open("POST", "https://www.kdocs.cn/api/v3/ide/file/:file_id/script/:script_id/task");
xhr.setRequestHeader("Content-Type", "application/json");
xhr.setRequestHeader("AirScript-Token", "xxx");

xhr.send(data);
```

Node.js
```json
const http = require("https");

const options = {
	"method": "POST",
	"hostname": "www.kdocs.cn",
	"port": null,
	"path": "/api/v3/ide/file/:file_id/script/:script_id/task",
	"headers": {
		"Content-Type": "application/json",
		"AirScript-Token": "xxx"
	}
};

const req = http.request(options, function (res) {
	const chunks = [];

	res.on("data", function (chunk) {
		chunks.push(chunk);
	});

	res.on("end", function () {
		const body = Buffer.concat(chunks);
		console.log(body.toString());
	});
});

req.write(JSON.stringify({Context: {argv: {}, sheet_name: '表名', range: '$B$156'}}));
req.end();
```

C
```json
CURL *hnd = curl_easy_init();

curl_easy_setopt(hnd, CURLOPT_CUSTOMREQUEST, "POST");
curl_easy_setopt(hnd, CURLOPT_URL, "https://www.kdocs.cn/api/v3/ide/file/:file_id/script/:script_id/task");

struct curl_slist *headers = NULL;
headers = curl_slist_append(headers, "Content-Type: application/json");
headers = curl_slist_append(headers, "AirScript-Token: xxx");
curl_easy_setopt(hnd, CURLOPT_HTTPHEADER, headers);

curl_easy_setopt(hnd, CURLOPT_POSTFIELDS, "{\"Context\":{\"argv\":{},\"sheet_name\":\"表名\",\"range\":\"$B$156\"}}");

CURLcode ret = curl_easy_perform(hnd);
```

C#
```json
var client = new RestClient("https://www.kdocs.cn/api/v3/ide/file/:file_id/script/:script_id/task");
var request = new RestRequest(Method.POST);
request.AddHeader("Content-Type", "application/json");
request.AddHeader("AirScript-Token", "xxx");
request.AddParameter("application/json", "{\"Context\":{\"argv\":{},\"sheet_name\":\"表名\",\"range\":\"$B$156\"}}", ParameterType.RequestBody);
IRestResponse response = client.Execute(request);
```

### [返回示例​](#response-example1)

json
```json
{
  "data": {
    "task_id": "GN/KU3B3BG84MdCjraN5mukx0Rt5Sp1eJ9k2qClmcaOkkF3PUVNDOYPY7Kz4aQMXSvXn9N08QabldRKjPfzii87fuGYydIuK2la2HMfcxmGK1Pf4WcPEflb5xOOkQQEo8fmEbzcobhurYg=="
  },
  "task_id": "GN/KU3B3BG84MdCjraN5mukx0Rt5Sp1eJ9k2qClmcaOkkF3PUVNDOYPY7Kz4aQMXSvXn9N08QabldRKjPfzii87fuGYydIuK2la2HMfcxmGK1Pf4WcPEflb5xOOkQQEo8fmEbzcobhurYg==",
  "task_type": "open_air_script"
}
```

## [获取任务运行情况​](#获取任务运行情况)

GET /api/v3/script/task

### [query 参数​](#query3)

| 参数 | 必须 | 类型 | 说明 |
| --- | --- | --- | --- |
| task_id | 是 | string | 执行异步任务时返回的 ID |

提示

任务ID为query参数，拼接时请注意先编码下，比如encodeURIComponent(task_id)

### [返回参数​](#return3)

| 参数 | 必须 | 类型 | 说明 |
| --- | --- | --- | --- |
| data | 是 | Object | 任务执行数据对象 |
| data.result | 是 | string | 任务执行返回的数据 |
| data.logs | 是 | Array | 任务执行日志 |
| data.logs[i].filename | 是 | string | 执行文件的名称 |
| data.logs[i].timestamp | 是 | string | 执行时间 |
| data.logs[i].unix_time | 是 | number | 执行 unix 时间戳 |
| data.logs[i].level | 是 | string | 日志级别 |
| data.logs[i].args | 是 | string[] | 日志打印参数 |
| status | 是 | string | 任务是否执行完毕 |
| error | 是 | string | 任务执行错误信息 |
| error_details | 否 | object | 错误信息详情对象 |
| error_details.name | 否 | string | 错误信息名称 |
| error_details.msg | 否 | string | 错误信息 |
| error_details.stack | 否 | string[] | 错误信息栈 |
| error_details.unix_time | 否 | number | 错误信息 unix 时间 |

### [请求示例​](#request-example3)

Shell
```json
curl --request GET \
	--url https://www.kdocs.cn/api/v3/script/task
```

Java
```json
OkHttpClient client = new OkHttpClient();

Request request = new Request.Builder()
	.url("https://www.kdocs.cn/api/v3/script/task")
	.get()
	.build();

Response response = client.newCall(request).execute();
```

Go
```json
package main

import (
	"fmt"
	"net/http"
	"io/ioutil"
)

func main() {

	url := "https://www.kdocs.cn/api/v3/script/task"

	req, _ := http.NewRequest("GET", url, nil)

	res, _ := http.DefaultClient.Do(req)

	defer res.Body.Close()
	body, _ := ioutil.ReadAll(res.Body)

	fmt.Println(res)
	fmt.Println(string(body))

}
```

Python
```json
import http.client

conn = http.client.HTTPSConnection("www.kdocs.cn")

conn.request("GET", "/api/v3/script/task")

res = conn.getresponse()
data = res.read()

print(data.decode("utf-8"))
```

PHP
```json
<?php

$curl = curl_init();

curl_setopt_array($curl, [
	CURLOPT_URL => "https://www.kdocs.cn/api/v3/script/task",
	CURLOPT_RETURNTRANSFER => true,
	CURLOPT_ENCODING => "",
	CURLOPT_MAXREDIRS => 10,
	CURLOPT_TIMEOUT => 30,
	CURLOPT_HTTP_VERSION => CURL_HTTP_VERSION_1_1,
	CURLOPT_CUSTOMREQUEST => "GET",
]);

$response = curl_exec($curl);
$err = curl_error($curl);

curl_close($curl);

if ($err) {
	echo "cURL Error #:" . $err;
} else {
	echo $response;
}
```

JS
```json
const data = null;

const xhr = new XMLHttpRequest();
xhr.withCredentials = true;

xhr.addEventListener("readystatechange", function () {
	if (this.readyState === this.DONE) {
		console.log(this.responseText);
	}
});

xhr.open("GET", "https://www.kdocs.cn/api/v3/script/task");

xhr.send(data);
```

Node.js
```json
const http = require("https");

const options = {
	"method": "GET",
	"hostname": "www.kdocs.cn",
	"port": null,
	"path": "/api/v3/script/task",
	"headers": {}
};

const req = http.request(options, function (res) {
	const chunks = [];

	res.on("data", function (chunk) {
		chunks.push(chunk);
	});

	res.on("end", function () {
		const body = Buffer.concat(chunks);
		console.log(body.toString());
	});
});

req.end();
```

C
```json
CURL *hnd = curl_easy_init();

curl_easy_setopt(hnd, CURLOPT_CUSTOMREQUEST, "GET");
curl_easy_setopt(hnd, CURLOPT_URL, "https://www.kdocs.cn/api/v3/script/task");

CURLcode ret = curl_easy_perform(hnd);
```

C#
```json
var client = new RestClient("https://www.kdocs.cn/api/v3/script/task");
var request = new RestRequest(Method.GET);
IRestResponse response = client.Execute(request);
```

### [返回示例​](#response-example3)

json
```json
{
  "data": {
    "logs": [
      {
        "filename": "<system>",
        "timestamp": "17:05:16.164",
        "unix_time": 1692090316164,
        "level": "info",
        "args": ["脚本环境初始化..."]
      }
    ],
    "result": null
  },
  "error": "Unexpected token (1:91)",
  "error_details": {
    "name": "SyntaxError",
    "msg": "Unexpected token (1:91)",
    "stack": ["    at 未命名脚本.js:1:91"],
    "unix_time": 1692090318372
  },
  "status": "finished"
}
```


## 简介

# [脚本令牌（APIToken）​](#脚本令牌-apitoken)

开发者通过 AirScript 编辑器编写的脚本，可以直接在编辑器内运行，也可以粘贴链接在单元格运行，或者是通过定时任务面板自动运行。

但是上述几种运行方式均集成在我们的平台方，如果开发者希望在自身的业务系统内使用到 AirScript 的能力，则需要借助我们的脚本令牌。

## [什么是脚本令牌？​](#什么是脚本令牌)

脚本令牌即 APIToken，是我们为外部系统引入 AirScript 能力而专门设计的。通过脚本令牌，您可以轻松使用到金山文档 AirScript 提供的能力，执行脚本获取文档数据或者是写入文档内容。

## [如何创建脚本令牌？​](#如何创建脚本令牌)

进入智能表格后，打开脚本编辑器，在工具栏点击【脚本令牌（beta）】按钮
如果之前未创建过脚本令牌，会提示脚本令牌创建所需要注意的点，勾选【我已知晓】，然后点击【创建脚本令牌】即可
为保证一定的安全，如果您还未进行过实名认证，需要先完成实名认证流程
创建成功后即可获取到您的个人脚本令牌，复制令牌信息然后妥善保存
## [如何使用脚本令牌？​](#如何使用脚本令牌)

脚本令牌是外部执行脚本的凭证，在您成功生成自己的脚本令牌后，便可以开始着手使用脚本令牌进行脚本的执行调用。

首先，打开脚本编辑器，在侧边栏任意一个文档脚本的更多菜单里复制脚本的 webhook 链接。

注意

暂时只开放了【文档共享脚本】的 webhook 链接复制，后续将会安排【我的脚本】的 webhook。

查看复制到的链接的内容如下所示：

https://www.kdocs.cn/api/v3/ide/file/caEkI6K5RDG2/script/V2-5rSiBiN7y5xdOd5x5ZYI2r/sync_task

链接内已拼接好了当前脚本的脚本 ID 和所在文档的文件 ID，接下来请求该链接即可读取和编辑本人相应的文档，注意调用的时候必须设置请求头AirScript-Token，值为您的脚本令牌，更详细的说明请参阅接口说明。

这里假设目标脚本的代码如下所示，将 A1 单元格的值修改为AirScript，并返回一个对象:

js
```js
Application.Range('A1').Value2 = 'AirScript'
return {
  name: '金小朦',
  age: 17
}
```

通过脚本令牌和 webhook 我们构造了一个 http 请求，如下所示：

Shell
```js
curl --request POST \
	--url https://www.kdocs.cn/api/v3/ide/file/caEkI6K5RDG2/script/V2-5rSiBiN7y5xdOd5x5ZYI2r/sync_task \
	--header 'AirScript-Token: xxx' \
	--header 'Content-Type: application/json' \
	--data '{"Context":{"argv":{}}}'
```

Java
```js
OkHttpClient client = new OkHttpClient();

MediaType mediaType = MediaType.parse("application/json");
RequestBody body = RequestBody.create(mediaType, "{\"Context\":{\"argv\":{}}}");
Request request = new Request.Builder()
	.url("https://www.kdocs.cn/api/v3/ide/file/caEkI6K5RDG2/script/V2-5rSiBiN7y5xdOd5x5ZYI2r/sync_task")
	.post(body)
	.addHeader("Content-Type", "application/json")
	.addHeader("AirScript-Token", "xxx")
	.build();

Response response = client.newCall(request).execute();
```

Go
```js
package main

import (
	"fmt"
	"strings"
	"net/http"
	"io/ioutil"
)

func main() {

	url := "https://www.kdocs.cn/api/v3/ide/file/caEkI6K5RDG2/script/V2-5rSiBiN7y5xdOd5x5ZYI2r/sync_task"

	payload := strings.NewReader("{\"Context\":{\"argv\":{}}}")

	req, _ := http.NewRequest("POST", url, payload)

	req.Header.Add("Content-Type", "application/json")
	req.Header.Add("AirScript-Token", "xxx")

	res, _ := http.DefaultClient.Do(req)

	defer res.Body.Close()
	body, _ := ioutil.ReadAll(res.Body)

	fmt.Println(res)
	fmt.Println(string(body))

}
```

Python
```js
import http.client

conn = http.client.HTTPSConnection("www.kdocs.cn")

payload = "{\"Context\":{\"argv\":{}}}"

headers = {
    'Content-Type': "application/json",
    'AirScript-Token': "xxx"
    }

conn.request("POST", "/api/v3/ide/file/caEkI6K5RDG2/script/V2-5rSiBiN7y5xdOd5x5ZYI2r/sync_task", payload, headers)

res = conn.getresponse()
data = res.read()

print(data.decode("utf-8"))
```

PHP
```js
<?php

$curl = curl_init();

curl_setopt_array($curl, [
	CURLOPT_URL => "https://www.kdocs.cn/api/v3/ide/file/caEkI6K5RDG2/script/V2-5rSiBiN7y5xdOd5x5ZYI2r/sync_task",
	CURLOPT_RETURNTRANSFER => true,
	CURLOPT_ENCODING => "",
	CURLOPT_MAXREDIRS => 10,
	CURLOPT_TIMEOUT => 30,
	CURLOPT_HTTP_VERSION => CURL_HTTP_VERSION_1_1,
	CURLOPT_CUSTOMREQUEST => "POST",
	CURLOPT_POSTFIELDS => "{\"Context\":{\"argv\":{}}}",
	CURLOPT_HTTPHEADER => [
		"AirScript-Token: xxx",
		"Content-Type: application/json"
	],
]);

$response = curl_exec($curl);
$err = curl_error($curl);

curl_close($curl);

if ($err) {
	echo "cURL Error #:" . $err;
} else {
	echo $response;
}
```

JS
```js
const data = JSON.stringify({
	"Context": {
		"argv": {}
	}
});

const xhr = new XMLHttpRequest();
xhr.withCredentials = true;

xhr.addEventListener("readystatechange", function () {
	if (this.readyState === this.DONE) {
		console.log(this.responseText);
	}
});

xhr.open("POST", "https://www.kdocs.cn/api/v3/ide/file/caEkI6K5RDG2/script/V2-5rSiBiN7y5xdOd5x5ZYI2r/sync_task");
xhr.setRequestHeader("Content-Type", "application/json");
xhr.setRequestHeader("AirScript-Token", "xxx");

xhr.send(data);
```

Node.js
```js
const http = require("https");

const options = {
	"method": "POST",
	"hostname": "www.kdocs.cn",
	"port": null,
	"path": "/api/v3/ide/file/caEkI6K5RDG2/script/V2-5rSiBiN7y5xdOd5x5ZYI2r/sync_task",
	"headers": {
		"Content-Type": "application/json",
		"AirScript-Token": "xxx"
	}
};

const req = http.request(options, function (res) {
	const chunks = [];

	res.on("data", function (chunk) {
		chunks.push(chunk);
	});

	res.on("end", function () {
		const body = Buffer.concat(chunks);
		console.log(body.toString());
	});
});

req.write(JSON.stringify({Context: {argv: {}}}));
req.end();
```

C
```js
CURL *hnd = curl_easy_init();

curl_easy_setopt(hnd, CURLOPT_CUSTOMREQUEST, "POST");
curl_easy_setopt(hnd, CURLOPT_URL, "https://www.kdocs.cn/api/v3/ide/file/caEkI6K5RDG2/script/V2-5rSiBiN7y5xdOd5x5ZYI2r/sync_task");

struct curl_slist *headers = NULL;
headers = curl_slist_append(headers, "Content-Type: application/json");
headers = curl_slist_append(headers, "AirScript-Token: xxx");
curl_easy_setopt(hnd, CURLOPT_HTTPHEADER, headers);

curl_easy_setopt(hnd, CURLOPT_POSTFIELDS, "{\"Context\":{\"argv\":{}}}");

CURLcode ret = curl_easy_perform(hnd);
```

C#
```js
var client = new RestClient("https://www.kdocs.cn/api/v3/ide/file/caEkI6K5RDG2/script/V2-5rSiBiN7y5xdOd5x5ZYI2r/sync_task");
var request = new RestRequest(Method.POST);
request.AddHeader("Content-Type", "application/json");
request.AddHeader("AirScript-Token", "xxx");
request.AddParameter("application/json", "{\"Context\":{\"argv\":{}}}", ParameterType.RequestBody);
IRestResponse response = client.Execute(request);
```

如果请求成功，将会返回如下数据，其内容主要包含脚本运行的日志信息和在代码中 return 的数据，如果您的脚本代码书写有误，相应的报错信息也会在日志中有所体现。

json
```json
{
  "data": {
    "logs": [
      {
        "filename": "<system>",
        "timestamp": "12:03:20.711",
        "unix_time": 1691726600711,
        "level": "info",
        "args": ["脚本环境初始化..."]
      },
      {
        "filename": "<system>",
        "timestamp": "12:03:22.129",
        "unix_time": 1691726602129,
        "level": "info",
        "args": ["已开始执行"]
      },
      {
        "filename": "<system>",
        "timestamp": "12:03:22.312",
        "unix_time": 1691726602312,
        "level": "info",
        "args": ["执行完毕"]
      }
    ],
    "result": {
      "age": 17,
      "name": "金小朦"
    }
  },
  "error": "",
  "status": "finished"
}
```

## [注意事项​](#注意事项)

由于脚本令牌允许第三方访问到平台的服务端资源，为提高一定的
安全性
，我们需要您完成实名认证（已认证可忽略）
脚本令牌，是外部执行脚本的凭证，属于
个人隐私信息
，通过脚本令牌配合脚本 webhook，可读取和编辑本人相应的文件，需妥善管理，请勿对外传播
脚本令牌与用户绑定，每个用户最多
有且仅有一个
脚本令牌，创建新的令牌时，需要先对老令牌进行删除（重新创建的脚本令牌需与原令牌不同）
脚本令牌默认 180 天到期，用户可在创建时手动进行延期，不做限制，可多次延期

# 示范案例

## 多维表

# [多维表案例​](#多维表案例)

此处提供了一些多维表脚本开发示范实例，希望能为您快速理解和上手多维表脚本开发提供帮助。

## [选中区域快速批量填值​](#选中区域快速批量填值)

javascript
```javascript
function main() {
  var activeView = Application.Selection.GetActiveView()

  var selectedRecords = Application.Selection.GetSelectionRecords()[0]
  var time = getNowTime()
  var date = getNowDate()

  // 注意，这里的 “日期”， “时间”， “分类”， 需要替换到您表中的响应字段名
  Application.Record.UpdateRecords({
    SheetId: activeView.sheetId,
    Records: selectedRecords.map(item => ({
      id: item.id,
      fields: {
        日期: date,
        时间: time,
        分类: 'B'
      }
    }))
  })
}
// 获取当前时间，格式为 "hh:mm:ss"
function getNowTime() {
  return new Date().toTimeString().split(' ')[0]
}
// 获取当前日期，格式为 "yyyy:MM:dd"
function getNowDate() {
  var date = new Date()
  return date.getFullYear() + '/' + (date.getMonth() + 1) + '/' + date.getDate()
}

main()
```

## [快速实现“一键归档”​](#快速实现-一键归档)

下面代码实现了一个文件中两张数据结构相同的表，把表一中的已完成的数据插入到表二中，并删除表一中数据。

表结构如图所示：

javascript
```javascript
function main() {
  var sheets = Application.Sheet.GetSheets()
  // 筛选出 打钩的 并且是分类 B 中的数据
  var finishedRecords = Application.Record.GetRecords({
    SheetId: sheets[0].id,
    Filter: {
      mode: 'AND',
      criteria: [
        {
          field: '完成',
          op: 'Equals',
          values: ['1']
        },
        {
          field: '分类',
          op: 'Equals',
          values: ['B']
        }
      ]
    }
  })
  // 如果存在 筛选出来的数据
  if (finishedRecords) {
    // 在表二中插入数据
    Application.Record.CreateRecords({
      SheetId: sheets[1].id,
      Records: finishedRecords.records.map(item => ({
        fields: item.fields
      }))
    })
    // 在表一中删除数据
    Application.Record.DeleteRecords({
      SheetId: sheets[0].id,
      RecordIds: finishedRecords.records.map(item => item.id)
    })
  }
}
main()
```

说明

📌 结合上面两个例子，可以实现自动设置归档日期和时间，或者选中记录一键归档等等功能

## [快速创建一张表​](#快速创建一张表)

javascript
```javascript
function main() {
  Application.Sheet.CreateSheet({
    Name: '我的表',
    Fields: [
      { name: '名称', type: 'MultiLineText' },
      { name: '数量', type: 'Number' },
      { name: '日期', type: 'Date' },
      { name: '时间', type: 'Time' },
      { name: '复选框', type: 'Checkbox' },
      { name: '超链接', type: 'Url' },
      { name: '等级', type: 'Rating', max: 5 },
      { name: '电话', type: 'Phone' },
      { name: '身份证', type: 'ID' },
      { name: '货币', type: 'Currency' },
      { name: '百分比', type: 'Percentage' },
      { name: '邮箱', type: 'Email' },
      { name: '进度', type: 'Complete' },
      {
        name: '分类',
        type: 'SingleSelect',
        items: [{ value: 'A' }, { value: 'B' }, { value: 'C' }]
      },
      {
        name: '状态',
        type: 'MultipleSelect',
        items: [{ value: '已完成' }, { value: '未开始' }, { value: '进行中' }]
      },
      {
        name: '联系人',
        type: 'Contact',
        multipleContacts: false,
        noticeNewContact: false
      },
      { name: '富文本', type: 'Note' },
      { name: '附件', type: 'Attachment' },
      { name: '公式', type: 'Formula' },
      { name: '创建时间', type: 'CreatedTime' },
      { name: '创建者', type: 'CreatedBy' },
      { name: 'AutoNumber', type: 'AutoNumber' }
    ],
    Views: [
      { name: '表格', type: 'Grid' },
      { name: '看板', type: 'Kanban' },
      { name: '画册', type: 'Gallery' },
      { name: '表单', type: 'Form' },
      { name: '甘特', type: 'Gantt' }
    ]
  })
}
main()
```

## [格式化数据批量插入​](#格式化数据批量插入)

javascript
```javascript
function main() {
  var template = ['商品', 10]
  var records = []
  // 插入 100 条 A 类商品，
  for (let i = 1; i < 100; i++) {
    records.push({
      名称: template[0] + i,
      数量: template[1],
      分类: 'A'
    })
  }
  for (let i = 1; i < 100; i++) {
    records.push({
      名称: template[0] + (100 + i),
      数量: template[1] + 10,
      分类: 'B'
    })
  }
  for (let i = 1; i < 100; i++) {
    records.push({
      名称: template[0] + (200 + i),
      数量: template[1] + 10,
      分类: 'C'
    })
  }
  var sheet = Application.Selection.GetActiveSheet()
  Application.Record.CreateRecords({
    SheetId: sheet.sheetId,
    Records: records.map(item => ({
      fields: item
    }))
  })
}
main()
```


## 表格

# [表格案例​](#表格案例)

为了帮助您理解AirScript，以及快速上手脚本开发，我们将表格中开始-快捷工具中的部分功能源码展示出来， 希望您能在阅读完这些真实的案例代码后，能有效利用脚本编辑器来解决问题。

## [清除公式仅保留值​](#清除公式仅保留值)

手动选中一个或多个单元格，然后执行脚本，执行完成后会批量清除其公式，仅保留其值。

js
```js
const API = Application

// 用户选区
const selection = API.Selection
// 获取用户激活工作簿的使用范围
const usedRange = API.ActiveSheet.UsedRange

// 取二者的最小集合，即所需遍历最小集合
const rowFrom = Math.max(selection.Row, usedRange.Row)
const rowTo = Math.min(selection.Row + selection.Rows.Count - 1, usedRange.Row + usedRange.Rows.Count - 1)
const colFrom = Math.max(selection.Column, usedRange.Column)
const colTo = Math.min(selection.Column + selection.Columns.Count - 1, usedRange.Column + usedRange.Columns.Count - 1)

for (let i = rowFrom; i <= rowTo; i++) {
  const row = API.ActiveSheet.Rows(i) // 确定行
  for (let j = colFrom; j <= colTo; j++) {
    const rg = row.Columns(j) // 从行对象中指定列，从而确定单元格
    // Text是单元格的显示值，即如果该单元格原先是公式，那Text即公式计算后的结果
    // 如果想获取未计算的结果，使用rg.Formula
    const text = rg.Text
    if (text !== '') {
      rg.Value2 = text // 原先的Value2是公式，通过将Value2重写成计算后的结果从而清除公式
    }
  }
}
```

## [高亮错误手机号​](#高亮错误手机号)

手动选中一个或多个单元格，然后执行脚本，脚本会自动判断单元格的内的手机号是否正确，执行完成后会高亮错误手机号单元格。

js
```js
// 手机号码的正则表达式,匹配第一位是1,第二位是3-9,其余位是数字的11位字符串
const PhoneReg = /^1[3-9](\d{9})$/i

// 高亮颜色为黄色
const color = RGB(255, 255, 0)

// API 的简化引用
const API = Application

// 用户选择的区域
const selection = API.Selection

// 工作簿激活的范围
const usedRange = API.ActiveSheet.UsedRange

// 取二者的最小集合，即所需遍历最小集合
const rowFrom = Math.max(selection.Row, usedRange.Row)
const rowTo = Math.min(selection.Row + selection.Rows.Count - 1, usedRange.Row + usedRange.Rows.Count - 1)
const colFrom = Math.max(selection.Column, usedRange.Column)
const colTo = Math.min(selection.Column + selection.Columns.Count - 1, usedRange.Column + usedRange.Columns.Count - 1)

for (let i = rowFrom; i <= rowTo; i++) {
    const row = API.ActiveSheet.Rows(i) // 确定行对象
    for (let j = colFrom; j <= colTo; j++) {
        const rg = row.Columns(j) // 再确定列对象,即确定单元格
        const text = rg.Text  // 获取单元格的文本内容
        if (text !== '' && !PhoneReg.test(text)) { // 正则表达式匹配不通过的情况下
            rg.Interior.Color = color // 用黄色高亮该单元格
        }
    }
}
```

## [取消合并填充相同内容​](#取消合并填充相同内容)

手动选中一个或多个单元格，然后执行脚本，脚本会自动取消合并单元格并填充相同内容。

js
```js
const API = Application
// 用户选择的区域
const selection = API.Selection

// 如果选区内有合并的单元格
if (selection.MergeCells) {
  // 获取所有合并的单元格区域
  const areas = selection.MergeArea.Areas
  const areaCount = areas.Count
  
  // 遍历每个合并的单元格区域
  for (let areaIndex = 1; areaIndex <= areaCount; areaIndex++) {
    const area = areas.Item(areaIndex)
    const text = area.Formula // 保存取消合并前该合并单元格的内容
    
    // 获取合并单元格区域的范围
    const rowFrom = area.Row
    const rowTo = rowFrom + area.Rows.Count - 1
    const colFrom = area.Column
    const colTo = colFrom + area.Columns.Count - 1

    // 取消合并
    area.UnMerge()
    // 将原先合并单元格的区域内容改写成合并单元格的值
    for (let i = rowFrom; i <= rowTo; i++) {
      const row = API.ActiveSheet.Rows(i) // 确定行
      for (let j = colFrom; j <= colTo; j++) {
        const cell = row.Columns(j) // 确定列
        cell.Value2 = text // 将之前保存的内容写到单元格里
      }
    }
  }
}
```

## [统计重复次数​](#统计重复次数)

手动选中一个或多个单元格，然后执行脚本，脚本会自动计算选区内的重复的值和重复的次数，新建工作表并将重复的值和重复的次数写入新建工作表的A2单元格和B2单元格。

js
```js
const API = Application
// 用户选择的区域
const selection = API.Selection
// 工作簿激活的范围
const usedRange = API.ActiveSheet.UsedRange

// 取二者的最小集合，即所需遍历最小集合
const rowFrom = Math.max(selection.Row, usedRange.Row)
const rowTo = Math.min(selection.Row + selection.Rows.Count - 1, usedRange.Row + usedRange.Rows.Count - 1)
const colFrom = Math.max(selection.Column, usedRange.Column)
const colTo = Math.min(selection.Column + selection.Columns.Count - 1, usedRange.Column + usedRange.Columns.Count - 1)

// countMap用于统计每个字符串出现次数
const countMap = {}
for (let i = rowFrom; i <= rowTo; i++) {
    const row = API.ActiveSheet.Rows(i) // 确定行
    for (let j = colFrom; j <= colTo; j++) {
        const rg = row.Columns(j) // 再确定列对象,即确定单元格
        const text = rg.Text // 取出该单元格的值用于统计
        if (text) {
            if (countMap[text]) {
                countMap[text]++
            } else {
                countMap[text] = 1
            }
        }
    }
}

const Name = API.ActiveSheet.Name
// 在当前工作表之后新建一个工作表，使用默认名称
const newSheet = API.Sheets.Add(null, API.Sheets(Name), 1)

// 在新工作表的A1和B1写入'重复值'和'重复次数'
newSheet.Range('A1').Value2 = '重复值'
newSheet.Range('B1').Value2 = '重复次数'

// 从第二行开始,每行第一列输出字符串,第二列输出该字符串重复出现的次数
let index = 2
for (let k in countMap) {
    const row = newSheet.Rows(index)
    row.Columns(1).Value2 = "'" + k
    row.Columns(2).Value2 = countMap[k] + ''
    index++;
}
```


# API文档(1.0)

## 内置函数

# [概述​](#概述)

内置函数是用来帮助开发者处理字符串编码/解码、信息处理、参数获取和其他杂项任务的实用函数

## [Crypto​](#crypto)

对信息进行加密，摘要处理

### [示例​](#示例)

javascript
```javascript
// 摘要foo这个字符串信息
console.log(Crypto.createHash("md5").update("foo").digest("hex")) // acbd18db4cc2f85cedef654fccc4a4d8
console.log(Crypto.createHmac("sha256", "a secret").update('some data to hash').digest('hex')) //7fd04df92f636fd450bc841c9418e5825c17f33ad9c87c518115a45971f7f77e
```

### [方法列表​](#方法列表)

| 方法 | 返回类型 | 简介 |
| --- | --- | --- |
| createHash(algorithm) | hash | 创建摘要算法实例，允许"md5", "sha1", "sha", "sha256", "sha512" |
| createHmac(algorithm, key) | hmac | 创建HMAC算法实例，允许"md5", "sha1", "sha", "sha256", "sha512" |

## [hash​](#hash)

摘要对象，由Crypto产生

### [方法列表​](#方法列表-1)

| 方法 | 返回类型 | 简介 |
| --- | --- | --- |
| update(data[ ,inputEncoding]) | hash | 使用给定的 data 更新哈希内容，如果未提供 encoding，且 data 是字符串，则强制为 'utf8' 编码，如果 data 是 Buffer,则忽略 inputEncoding,可重复调用添加数据 |
| digest([encoding] | string| Buffer | 计算传给被哈希的所有数据的摘要，如果提供了 encoding，则将返回字符串；否则返回 Buffer。 |

## [hmac​](#hmac)

hmac对象，由Crypto产生

### [方法列表​](#方法列表-2)

| 方法 | 返回类型 | 简介 |
| --- | --- | --- |
| update(data[ ,inputEncoding]) | hash | 使用给定的 data 更新hmac内容，如果未提供 encoding，且 data 是字符串，则强制为 'utf8' 编码，如果 data 是 Buffer,则忽略 inputEncoding,可重复调用添加数据 |
| digest([encoding] | string| Buffer | 计算传给被hmac的所有数据的摘要，如果提供了 encoding，则将返回字符串；否则返回 Buffer。 |

## [Buffer​](#buffer)

产生一个 Buffer 实例

### [示例​](#示例-1)

javascript
```javascript
// 创建包含字符串 'buffer' 的 UTF-8 字节的新缓冲区。
const buf = Buffer.from([0x62, 0x75, 0x66, 0x66, 0x65, 0x72]); 
console.log(buf.toString()) // buffer
```

### [方法列表​](#方法列表-3)

| 方法 | 返回类型 | 简介 |
| --- | --- | --- |
| from(array) | Buffer | 使用 0 – 255 范围内的字节 array 分配新的 Buffer。 |
| from(string[, encoding]) | Buffer | 从字符串转化为Buffer |
| from(arrayBuffer[, byteOffset[, length]]) | Buffer | 截断arrayBuffer的部分字节，生成新的Buffer |

## [Time​](#time)

时间函数，提供如休眠的方法

### [示例​](#示例-2)

javascript
```javascript
Time.sleep(1000) // 休眠一秒
```

### [方法列表​](#方法列表-4)

| 方法 | 返回类型 | 简介 |
| --- | --- | --- |
| sleep(millisecond) | undefined | 休眠指定毫秒数 |

## [Arguments​](#arguments)

方便获取配置的参数数据

### [示例​](#示例-3)

javascript
```javascript
Arguments.get("foo.bar", "defaults") // 如果自定义参数是{foo : {bar : "value"}}，则返回"value"，如果不存在，则返回第二个参数"defaults"
```

### [方法列表​](#方法列表-5)

| 方法 | 返回类型 | 简介 |
| --- | --- | --- |
| get(string[, defaults]) | any | 通过获取自定义参数的值，key支持使用.进行多次查找，如a.b会寻找{a : {b : ""}}这个结构体的b值。可指定默认值，如果找不到key对应的自定义参数，就返回默认值，没有指定默认值也找不到key返回undefined |


## 多维表格

### 字段

# [Field​](#field)

字段操作

## [获取字段信息​](#获取字段信息)

javascript
```javascript
var fields = Application.Field.GetFields({ SheetId: 1 })
```

## [创建字段​](#创建字段)

Field 格式说明见多维表字段类型说明

javascript
```javascript
var field = Application.Field.CreateFields({
  SheetId: 3,
  Fields: [{ name: '等级', type: 'Rating', max: 5 }]
})
```

## [删除字段​](#删除字段)

javascript
```javascript
// result 中会返回是否删除成功
// [{"deleted":false,"id":"P"},{"deleted":false,"id":"Q"}]
var resutlt = Application.Field.DeleteFields({
  SheetId: 8,
  FieldIds: ['P', 'Q']
})
```

## [更新字段​](#更新字段)

javascript
```javascript
Application.Field.UpdateFields({
  SheetId: 3,
  Fields: [{ id: '6', name: '名称' }]
})
```


### 行记录

# [Record​](#record)

行记录

## [获取行记录（多条）​](#获取行记录-多条)

可选参数列表

| 参数 | 说明 |
| --- | --- |
| ViewId | 填写后将从被指定的视图获取该用户所见到的记录；若不填写，则从工作表获取记录 |
| PageSize | 存在分页时，指定本次查询的起始记录（含）。若不填写或填写为空字符串，则从第一条记录开始获取 |
| Offset | 分页查询时，还将返回一个offset值，指向下一页的第一条记录，供后续查询。查询到最后一页或第maxRecords条记录时，返回数据将不再包含offset值 |
| MaxRecords | 指定要获取的“前maxRecords条记录”，若不填写，则默认返回全部记录 |
| Fields | ["数量", "日期"，"时间"]，指定所返回记录中的字段信息，若不填写，则默认返回全部字段内的信息。 |
| Filter | 筛选条件，详细说明见筛选条件说明 |

javascript
```javascript
var records = Application.Record.GetRecords({ SheetId: 3 })
```

## [获取行记录（单条）​](#获取行记录-单条)

javascript
```javascript
var record = Application.Record.GetRecord({ SheetId: 3, RecordId: 'Bz' })
```

## [删除行记录​](#删除行记录)

javascript
```javascript
var result = Application.Record.DeleteRecords({
  SheetId: 3,
  RecordIds: ['J', 'P', 'Q']
})
```

## [更新行记录​](#更新行记录)

Field 格式说明见多维表字段类型说明

javascript
```javascript
var records = recordapi.UpdateRecords({
  SheetId: 5,
  Records: [
    {
      id: 'A',
      fields: {
        邮箱: 'demo@qq.com',
        多选: ['1', '2']
      }
    }
  ]
})
```

## [创建行记录​](#创建行记录)

Field 格式说明见多维表字段类型说明

javascript
```javascript
var records = recordapi.CreateRecords({
  SheetId: 5,
  Records: [
    {
      fields: {
        邮箱: 'demo@qq.com',
        多选: ['1', '2']
      }
    }
  ]
})
```


### 表

# [Sheet​](#sheet)

表操作

## [获取所有表信息​](#获取所有表信息)

javascript
```javascript
var sheets = Application.Sheet.GetSheets()
```

## [创建新的表​](#创建新的表)

参数说明

最少传入一个表和一个字段
Field 格式说明见
多维表字段类型说明
视图类型说明见
多维表视图类型说明
javascript
```javascript
var sheet = Application.Sheet.CreateSheet({
  Name: '商品表',
  Views: [{ name: '表格', type: 'Grid' }],
  Fields: [
    { name: '名称', type: 'MultiLineText' },
    {
      name: '分类',
      type: 'SingleSelect',
      items: [{ value: 'A' }, { value: 'B' }, { value: 'C' }]
    }
  ]
})
```

## [删除表​](#删除表)

javascript
```javascript
Application.Sheet.DeleteSheet({ SheetId: 8 })
```

## [更新表名字​](#更新表名字)

javascript
```javascript
var result = Application.Sheet.UpdateSheet({ SheetId: 9, Name: '备份表' })
```


### 视图

# [View​](#view)

视图操作

## [获取视图信息​](#获取视图信息)

javascript
```javascript
var views = Application.View.GetViews({ SheetId: 1 }）
```

## [创建视图​](#创建视图)

视图类型说明见多维表视图类型说明

javascript
```javascript
var view = Application.View.CreateView({
  SheetId: 1,
  Name: '表格',
  ViewType: 'Grid'
})
```

## [删除视图​](#删除视图)

javascript
```javascript
Application.View.DeleteView({ SheetId: 1, ViewId: 'F' })
```

## [修改视图名称​](#修改视图名称)

javascript
```javascript
Application.View.UpdateView({
  SheetId: 8,
  ViewId: 'E',
  Name: '名单'
})
```


### 选区

# [Selection​](#selection)

选区

## [获取当前选中视图​](#获取当前选中视图)

javascript
```javascript
var view = selectionHandler.GetActiveView()
```

## [获取选中记录信息​](#获取选中记录信息)

a. 这里返回的 records 是一个二维数组，因为选中记录可以是不连续的，如果是连续的记录，那么可以直接去 records[0]

javascript
```javascript
var records = selectionHandler.GetSelectionRecords()
```

## [设置选中的视图​](#设置选中的视图)

javascript
```javascript
selectionHandler.SetActiveView({ SheetId: 27, ViewId: 'Q' })
```

## [获取当前选中表​](#获取当前选中表)

javascript
```javascript
var sheet = selectionHandler.GetActiveSheet()
```


### 附录

# [附录​](#附录)

## [附录 1：多维表字段类型说明​](#附录-1-多维表字段类型说明)

具体使用方式可以参考快速创建一张表

| 字段类型 | Type | 创建字段格式 | 设置字段值传入形式 | 读取字段值传出形式 |
| --- | --- | --- | --- | --- |
| 多行文本 | MultiLineText | 无特殊要求 | 字符串/ 无特殊格式要求 | 字符串 |
| 日期 | Date | 无特殊要求 | 字符串/yyyy/mm/dd | 字符串 |
| 时间 | Time | 无特殊要求 | 字符串/hh:mm:ss | 字符串 |
| 数值 | Number | 无特殊要求 | 数值 / 无格式 | 数值 |
| 货币 | Currency | 无特殊要求 | 数值 / 无格式 | 数值 |
| 百分比 | Percentage | 无特殊要求 | 数值 / 无格式 | 数值 |
| 身份证 | ID | 无特殊要求 | 字符串 / 符合身份证规则 | 字符串 |
| 电话 | Phone | 无特殊要求 | 字符串 / 符合电话规则 | 字符串 |
| 电子邮箱 | Email | 无特殊要求 | 字符串 / 符合邮箱规则 | 字符串 |
| 超链接 | Url | 可以额外传入一个参数。displayText：指定超链接显示文本。{"name":"超链接","type":"Url","displayText":"跳转"} | 字符串 / 符合 Url 规 | 字符串 |
| 复选框 | Checkbox | 无特殊要求 | true / false | 布尔 |
| 单选项 | SingleSelect | 需要额外传入选项值，至少一个。{"name": "单选项","type": "SingleSelect","items": [{ "value": "item1" }]} | 字符串 / 匹配选项内容 | 字符串 |
| 多选项 | MultipleSelect | 需要额外传入选项值，至少一个。{"name": "单选项","type": "SingleSelect","items": [{ "value": "item1" }, { "value": "item2" }]} | 字符串数组 / 匹配选项内容 | 字符串数组 |
| 等级 | Rating | 需要额外传入一个最大等级, 最大等级大于 0 小于等于 5。{"name": "等级","type": "Rating","max": 5} | 数值 / 大于 0 并且 小于 最大等级 | 数值 |
| 进度条 | Complete | 无特殊要求 | 数值 / 大于等于 0 并且 小于 100 | 字符串 |
| 联系人 | Contact | 需要额外传入两个参数：multipleContacts:<bool>是否支持多个联系人noticeNewContact:<bool>是否通知联系人。{"name": "联系人","type": "Contact","multipleContacts": false,"noticeNewContact": false} | 不支持设值 | 对象 |
| 附件 | Attachment | 无特殊要求 | 不支持设值 | 对象 |
| 关联 | Link | 需要额外传入二个参数：linkSheet: 关联表 IDmultipleLinks: 是否关联多条记录{"name": "联系人","type": "Link","multipleContacts": false,"noticeNewContact": false} | 对应关联表的行记录数组 |  |
| 富文本 | Note | 无特殊要求 | 不支持设值 | 对象 |
| 编号 | AutoNumber | 无特殊要求 | 不支持设值 | 数值 |
| 创建者 | CreatedBy | 无特殊要求 | 不支持设值 | 对象 |
| 创建时间 | CreatedTime | 无特殊要求 | 不支持设值 | 字符串 |
| 公式 | Formula | 无特殊要求 | 不支持设值 | 根据公式的值类型 |
| 引用 | Lookup | 无特殊要求 | 不支持设值 | 与被引用形式相同 |

## [附录 2：多维表视图类型说明​](#附录-2-多维表视图类型说明)

具体使用方式可以参考快速创建一张表

| 视图类型 | 说明 |
| --- | --- |
| Grid | 表格视图 |
| Kanban | 看板视图 |
| Gallery | 画册视图 |
| Form | 表单视图 |
| Gantt | 甘特视图 |

## [附录 3：筛选条件说明​](#附录-3-筛选条件说明)

筛选条件用来对行记录进行筛选，由两部分构成：mode为筛选条件关系；creteria为具体筛选条件（fileds op values）。

json
```json
{
  "mode": "AND", // 选填。表示各筛选条件之间的逻辑关系。只能是"AND"或"OR"。缺省值为"AND"
  "criteria": [
    // filter结构体内必填。包含筛选条件的数组。每个字段上只能有一个筛选条件
    {
      "field": "名称", // 必填。根据 preferId 与否，需要填入字段名或字段id
      "op": "Intersected", // 必填。表示具体的筛选规则，见下
      "values": [
        // 必填。表示筛选规则中的值。数组形式。
        "多维表", // 值为字符串时表示文本匹配
        "12345"
      ]
    },
    {
      "field": "数量",
      "op": "greater",
      "values": ["1"]
    }
  ]
}
```

| 筛选条件 | 参数说明 |
| --- | --- |
| Equals | 等于 |
| NotEqu | 不等于 |
| Greater | 大于 |
| GreaterEqu | 大等于 |
| Less | 小于 |
| LessEqu | 小等于 |
| GreaterEquAndLessEqu | 介于（取等） |
| LessOrGreater | 介于（不取等） |
| BeginWith | 开头是 |
| EndWith | 结尾是 |
| Contains | 包含 |
| NotContains | 不包含 |
| Intersected | 指定值 |
| Empty | 为空 |
| NotEmpty | 不为空 |

各筛选规则独立地限制了 values 数组内最多允许填写的元素数，当 values 内元素数超过阈值时，该筛选规则将失效。

为空、不为空不允许填写元素；介于允许最多填写 2 个元素；指定值允许填写 65535 个元素；其他规则允许最多填写 1 个元素。

注意

filter 不是结构体，当 criteria 未指定 field、op/values 参数填写不合法、values 填写过多参数及其他可能导致筛选规则失效等情形，整个请求将直接失败。

目前还支持对日期进行动态筛选，此时 values[]内的元素需以结构体的形式给出：

json
```json
{
  "mode": "AND",
  "criteria": [
    {
      "field": "日期",
      "op": "Equals",
      "values": [
        {
          "dynamicType": "lastMonth",
          "type": "DynamicSimple"
        }
      ]
    }
  ]
}
```

提示

上述示例对应的筛选条件为等于上一个月。

要使用日期动态筛选，values[]内的结构体需要指定type为DynamicSimple，当op为Equals时，dynamicType可以为如下的值（大小写不敏感）。

| 字段 | 说明 |
| --- | --- |
| today | 今天 |
| yesterday | 昨天 |
| tomorrow | 明天 |
| last7Days | 最近 7 天 |
| last30Days | 最近 30 天 |
| thisWeek | 本周 |
| lastWeek | 上周 |
| nextWeek | 下周 |
| thisMonth | 本月 |
| lastMonth | 上月 |
| nextMonth | 次月 |

提示

当op为greater或less时，dynamicType只能是昨天、今天或明天。


## 智能表格

### 工作表

#### API总览

# [API总览​](#api总览)

[
API
全部
[
表格实例(Application)
工作簿(Workbook)
工作表(Sheet)
区域(Range)
筛选(AutoFilter)
排序(Sort)
排序字段(SortField)
字体(Font)
边框(Border)
图形(Shape)
图表(Chart)
超链接(Hyperlink)
条件格式集合(FormatConditions)
条件格式(FormatCondition)
数据有效性规则(Validation)
工作表函数(WorksheetFunction)
枚举(Enum)
]
分类
[
全部
attribute
function
]
[
[
### 表格实例(Application)

| [ActiveSheet](/api/excel/workbook/Application#activesheet) | [Sheet](/api/excel/workbook/Sheet) | 当前的活动工作表 |
| --- | --- | --- |
| [Sheets](/api/excel/workbook/Application#sheets) | [Sheets](/api/excel/workbook/Application#sheets) | 当前文件的所有工作表 |
| [ActiveWorkbook](/api/excel/workbook/Application#activeworkbook) | [Workbook](/api/excel/workbook/Workbook) | 当前的文档 |
| [Selection](/api/excel/workbook/Application#selection) | [Range](/api/excel/workbook/Range) | 当前的选区对象 |
| [Cells](/api/excel/workbook/Application#cells) | [Range](/api/excel/workbook/Range) | 当前工作表所有单元格 |
| [Columns](/api/excel/workbook/Application#columns) | [Range](/api/excel/workbook/Range) | 当前工作表所有列 |
| [Rows](/api/excel/workbook/Application#rows) | [Range](/api/excel/workbook/Range) | 当前工作表所有行 |
| [FileInfo](/api/excel/workbook/Application#fileinfo) | Object | 当前文档的信息 |
| [UserInfo](/api/excel/workbook/Application#userinfo) | Object | 当前文档的用户信息 |
| [Enum](/api/excel/workbook/Application#enum) | [Enum](/api/excel/workbook/Enum) | 所有的枚举类型 |

| [Range(address)](/api/excel/workbook/Application#range) | [Range](/api/excel/workbook/Range) | 获取当前 ActiveSheet 的某个区域(address 指定) |
| --- | --- | --- |
| [Sheets(name)](/api/excel/workbook/Application#sheets-1) | [Sheet](/api/excel/workbook/Sheet) | 获取名称为 name 的工作表 |

]
[
### 工作簿(Workbook)

| [ActiveSheet](/api/excel/workbook/Workbook#activesheet) | [Sheet](/api/excel/workbook/Sheet) | 工作簿中的活动工作表 |
| --- | --- | --- |
| [Sheets](/api/excel/workbook/Workbook#sheets) | [Sheets](/api/excel/workbook/Application#sheets) | 工作表的所有对象集合 |
| [ReadOnly](/api/excel/workbook/Workbook#readonly) | Boolean | 文档是否只读 |
| [ReadOnlyComment](/api/excel/workbook/Workbook#readonlycomment) | Boolean | 文档是否只读可评论的权限 |
| [SupportReadOnlyComment](/api/excel/workbook/Workbook#supportreadonlycomment) | Boolean | 文档是否支持只读可评论权限 |

| [Save()](/api/excel/workbook/Workbook#save) | String(JSON) | 保存文件的改动 |
| --- | --- | --- |
| [GetComments()](/api/excel/workbook/Workbook#getcomments) | String(JSON) | 获取整个 Workbook 的评论 |
| [ExportAsFixedFormat()](/api/excel/workbook/Workbook#exportasfixedformat) | String(JSON) | 导出整个表格的 PDF 或者 Img 图片 |

]
[
### 工作表(Sheet)

| [Id](/api/excel/workbook/Sheet#id) | String | 该工作表的 Id |
| --- | --- | --- |
| [Name](/api/excel/workbook/Sheet#name) | String | 该工作表的名称 |
| [Index](/api/excel/workbook/Sheet#index) | Number | 该工作表在所有工作表的索引值 |
| [Cells](/api/excel/workbook/Sheet#cells) | [Range](/api/excel/workbook/Range) | 该工作表上所有单元格的集合 |
| [Columns](/api/excel/workbook/Sheet#columns) | [Range](/api/excel/workbook/Range) | 该工作表上所有列的集合 |
| [Rows](/api/excel/workbook/Sheet#rows) | [Range](/api/excel/workbook/Range) | 该工作表上所有行的集合 |
| [UsedRange](/api/excel/workbook/Sheet#usedrange) | [Range](/api/excel/workbook/Range) | 该工作表的使用范围 |
| [Visible](/api/excel/workbook/Sheet#visible) | Boolean | 该工作表是否可见 |
| [Type](/api/excel/workbook/Sheet#type) | String | 该工作表的类型 |
| [Hyperlinks](/api/excel/workbook/Sheet#hyperlinks) | [Hyperlinks](/api/excel/workbook/Sheet#hyperlinks) | 该工作表上所有超链接的集合 |
| [Shapes](/api/excel/workbook/Sheet#shapes) | [Shapes](/api/excel/workbook/Sheet#shapes) | 该工作表上所有图形的集合 |
| [Sort](/api/excel/workbook/Sheet#sort) | [Sort](/api/excel/workbook/Sheet#/api/excel/workbook/Sort) | 该工作表上排序对象 |

| [Range()](/api/excel/workbook/Sheet#range) | [Range](/api/excel/workbook/Range) | 一个单元格或单元格区域 |
| --- | --- | --- |
| [Cells()](/api/excel/workbook/Sheet#cells-1) | [Range](/api/excel/workbook/Range) | 该工作表上的某个单元格 |
| [Activate()](/api/excel/workbook/Sheet#activate) | undefined | 切换(激活)工作表 |
| [Move()](/api/excel/workbook/Sheet#move) | undefined | 移动工作表 |
| [Delete()](/api/excel/workbook/Sheet#delete) | undefined | 删除工作表 |

]
[
### 区域(Range)

| [Count](/api/excel/workbook/Range#count) | Number | 区域中单元格的数量 |
| --- | --- | --- |
| [Text](/api/excel/workbook/Range#text) | String | 【只读】读取单元格格式化文本 |
| [Value/Value2](/api/excel/workbook/Range#value) | any/[][]any | 读写单元格中的值 |
| [FormatConditions](/api/excel/workbook/Range#formatConditions) | [FormatConditions](/api/excel/workbook/FormatConditions) | 用于控制 Excel 中的条件格式 |
| [Formula](/api/excel/workbook/Range#formula) | String | 以 A1 样式表示法表示的对象的隐式交叉的公式 |
| [FormulaArray](/api/excel/workbook/Range#formulaarray) | String | 返回或设置区域的数组公式 |
| [NumberFormat](/api/excel/workbook/Range#numberformat) | String | 获取或者设置区域的数字格式 |
| [Hidden](/api/excel/workbook/Range#hidden) | Boolean | 行或者列的隐藏 |
| [Interior.Color](/api/excel/workbook/Range#interior-color) | String | 内部颜色的十六进制 RGB |
| [HorizontalAlignment](/api/excel/workbook/Range#horizontalalignment) | [Enum.XlHAlign](/api/excel/workbook/Enum#xlhalign) | 设置区域的水平对齐方式 |
| [VerticalAlignment](/api/excel/workbook/Range#verticalalignment) | [Enum.XlVAlign](/api/excel/workbook/Enum#xlvalign) | 设置区域的垂直对齐方式 |
| [WrapText](/api/excel/workbook/Range#wraptext) | Boolean | 获取或者设置区域自动换行 |
| [IndentLevel](/api/excel/workbook/Range#indentlevel) | Number | 单元格缩进 |
| [MergeArea](/api/excel/workbook/Range#mergearea) | [Range](/api/excel/workbook/Range) | 单元格的合并区域 |
| [MergeCells](/api/excel/workbook/Range#mergecells) | Boolean | 区域内是否存在合并的单元格 |
| [Cells](/api/excel/workbook/Range#cells) | [Range](/api/excel/workbook/Range) | 区域中的单元格集合 |
| [Rows](/api/excel/workbook/Range#rows) | [Range](/api/excel/workbook/Range) | 区域中的行集合 |
| [Columns](/api/excel/workbook/Range#columns) | [Range](/api/excel/workbook/Range) | 区域中的列集合 |
| [EntireRow](/api/excel/workbook/Range#entirerow) | [Range](/api/excel/workbook/Range) | 区域所在行的整行 |
| [EntireColumn](/api/excel/workbook/Range#entirecolumn) | [Range](/api/excel/workbook/Range) | 区域所在列的整列 |
| [Row](/api/excel/workbook/Range#row) | Number | 区域中第一行的行号 |
| [RowEnd](/api/excel/workbook/Range#rowend) | Number | 区域中最后一行的行号 |
| [Column](/api/excel/workbook/Range#column) | Number | 区域中第一列的列号 |
| [ColumnEnd](/api/excel/workbook/Range#columnend) | Number | 区域中最后一列的列号 |
| [Borders](/api/excel/workbook/Range#borders) | [Borders](/api/excel/workbook/Range#borders) | 边框集合对象 |

| [BorderAround()](/api/excel/workbook/Range#borderaround) | undefined | 向区域添加边框，并为新边框设置 Border 对象的 Color、LineStyle 和 Weight 属性 |
| --- | --- | --- |
| [Each()](/api/excel/workbook/Range#each) | undefined | 遍历选区所选单元格 |
| [Item()](/api/excel/workbook/Range#item) | [Range](/api/excel/workbook/Range) | 表示区域中指定的位置 |
| [Offset()](/api/excel/workbook/Range#offset) | undefined | 对指定区域进行迁移操作 |
| [Replace()](/api/excel/workbook/Range#replace) | undefined | 对单元格内文本执行替换操作 |
| [Delete()](/api/excel/workbook/Range#delete) | undefined | 单元格、行、列的删除 |
| [Insert()](/api/excel/workbook/Range#insert) | undefined | 单元格、行、列的新增 |
| [InsertImage()](/api/excel/workbook/Range#insertimage) | undefined | 插入单元格图片 |
| [Merge()](/api/excel/workbook/Range#merge) | undefined | 合并单元格 |
| [UnMerge()](/api/excel/workbook/Range#unmerge) | undefined | 取消合并单元格 |
| [Address()](/api/excel/workbook/Range#address) | String | 获取表示使用宏语言的区域引用的 String 值 |
| [AddComment()](/api/excel/workbook/Range#addcomment) | undefined | 添加评论 |
| [ClearComments()](/api/excel/workbook/Range#clearcomments) | undefined | 清除区域的评论 |
| [Clear()](/api/excel/workbook/Range#clear) | undefined | 清空指定区域数据和样式 |
| [ClearContents()](/api/excel/workbook/Range#clearcontents) | undefined | 清除区域的内容 |
| [ClearFormats()](/api/excel/workbook/Range#clearformats) | undefined | 清除区域的样式 |
| [ClearHyperlinks()](/api/excel/workbook/Range#clearhyperlinks) | undefined | 清除区域的超链接样式 |
| [Contain()](/api/excel/workbook/Range#contain) | Boolean | 判断区域是否重叠 |
| [Copy()](/api/excel/workbook/Range#copy) | Boolean | 将当前区域对象复制到剪贴板 |
| [Cut()](/api/excel/workbook/Range#cut) | Boolean | 将当前区域对象粘贴到目标区域 |
| [PasteSpecial()](/api/excel/workbook/Range#pastespecial) | undefined | 将剪贴板中的内容粘贴到指定的单元格或范围 |
| [FillLeft()](/api/excel/workbook/Range#fillleft) | undefined | 对指定区域中的单元格执行从右往左填充 |
| [FillRight()](/api/excel/workbook/Range#fillright) | undefined | 对指定区域中的单元格执行从左往右填充 |
| [FillDown()](/api/excel/workbook/Range#filldown) | undefined | 对指定区域中的单元格执行从上往下填充 |
| [FillUp()](/api/excel/workbook/Range#fillup) | undefined | 对指定区域中的单元格执行从下往上填充 |
| [AutoFill()](/api/excel/workbook/Range#autofill) | undefined | 对指定区域中的单元格执行自动填充 |
| [AutoFilter()](/api/excel/workbook/Range#autofilter) | undefined | 对指定区域中的单元格执行自动筛选 |
| [AutoFit()](/api/excel/workbook/Range#autofit) | undefined | 更改区域中的列宽或行高以达到最佳匹配 |
| [Select()](/api/excel/workbook/Range#select) | undefined | 选择区域 |
| [TextToColumns()](/api/excel/workbook/Range#texttocolumns) | undefined | 将包含文本的一列单元格分解为若干列 |

]
[
### 筛选(AutoFilter)

| [Filters](/api/excel/workbook/AutoFilter#filters) | [Filters](/api/excel/workbook/AutoFilter#filters) | 筛选对象集合 |
| --- | --- | --- |
| [Range](/api/excel/workbook/AutoFilter#range) | [Range](/api/excel/workbook/Range) | 筛选区域 |

| [ApplyFilter()](/api/excel/workbook/AutoFilter#applyfilter) | undefined | 应用筛选到当前工作表 |
| --- | --- | --- |
| [ShowAllData()](/api/excel/workbook/AutoFilter#showalldata) | undefined | 清除所有筛选条件，显示所有数据 |

]
[
### 排序(Sort)

| [Header](/api/excel/workbook/Sort#header) | [XlYesNoGuess](/api/excel/workbook/Enum#xlyesnoguess) | 指定第一行是否包含标题信息 |
| --- | --- | --- |
| [MatchCase](/api/excel/workbook/Sort#matchcase) | Boolean | 是否区分大小写 |
| [Orientation](/api/excel/workbook/Sort#orientation) | [XlSortOrientation](/api/excel/workbook/Enum#xlsortorientation) | 指定排序方向 |
| [Rng](/api/excel/workbook/Sort#rng) | [Range](/api/excel/workbook/Range) | 返回要执行排序的值的区域 |
| [SortFields](/api/excel/workbook/Sort#sortfields) | [SortFields](/api/excel/workbook/SortFields) | 该对象代表与 **Sort** 对象关联的排序字段的集合 |
| [SortMethod](/api/excel/workbook/Sort#sortmethod) | [XlSortMethod](/api/excel/workbook/Enum#xlsortmethod) | 中文排序方法 |

| [Apply()](/api/excel/workbook/Sort#apply) | undefined | 根据当前应用的排序状态对区域进行排序 |
| --- | --- | --- |
| [SetRange()](/api/excel/workbook/Sort#setrange) | undefined | 设置排序发生的范围 |

]
[
### 排序字段(SortField)

代表与 **Sort** 对象关联的排序字段对象

| [CustomOrder](/api/excel/workbook/SortField#customorder) | Variant | 指定对字段进行排序的自定义次序 |
| --- | --- | --- |
| [DataOption](/api/excel/workbook/SortField#dataoption) | [XlSortDataOption](/api/excel/workbook/Enum#xlsortdataoption) | 指定如何在 SortField 对象中指定的区域中对文本进行排序 |
| [Key](/api/excel/workbook/SortField#key) | [Range](/api/excel/workbook/Range) | 指定排序字段，该字段确定要排序的值 |
| [Order](/api/excel/workbook/SortField#order) | [XlSortOrder](/api/excel/workbook/Enum#xlsortorder) | 确定关键字所指定的值的排序次序 |
| [Priority](/api/excel/workbook/SortField#priority) | Number | 指定排序字段的优先级 |
| [SortOn](/api/excel/workbook/SortField#sorton) | [XlSortOn](/api/excel/workbook/Enum#xlsorton) | 返回或设置要排序的单元格的属性 |
| [SortOnValue](/api/excel/workbook/SortField#sortonvalue) | Object | 返回针对指定的 SortField 对象执行排序的值 |

| [Delete()](/api/excel/workbook/SortField#delete) | undefined | 从 SortFields 集合中删除指定的 SortField 对象 |
| --- | --- | --- |
| [ModifyKey()](/api/excel/workbook/SortField#modifykey) | undefined | 修改字段中按其排序的键值 |

]
[
### 字体(Font)

单元格内字体的属性，包括加粗，颜色，大小，斜体，删除线和下划线。

| [Bold](/api/excel/workbook/Font#bold) | Boolean | 字体是否加粗 |
| --- | --- | --- |
| [Color](/api/excel/workbook/Font#color) | String | 字体的颜色 |
| [Italic](/api/excel/workbook/Font#italic) | Boolean | 字体是否斜体 |
| [Name](/api/excel/workbook/Font#name) | String | 字体的名称 |
| [Size](/api/excel/workbook/Font#size) | Number | 字体的大小 |
| [Strikethrough](/api/excel/workbook/Font#strikethrough) | Boolean | 字体是否有删除线 |
| [Underline](/api/excel/workbook/Font#underline) | [Enum.XlUnderlineStyle](/api/excel/workbook/Enum#xlunderlinestyle) | 字体是否有下划线 |

]
[
### 边框(Border)

边框对象，Borders 集合里的某一边框

| [Color](/api/excel/workbook/Border#color) | Number | 边框的颜色 |
| --- | --- | --- |
| [Weight](/api/excel/workbook/Border#weight) | [Enum.XlBorderWeight](/api/excel/workbook/Enum#xlborderweight) | 边框的粗细 |
| [LineStyle](/api/excel/workbook/Border#linestyle) | [Enum.XlLineStyle](/api/excel/workbook/Enum#xllinestyle) | 边框的线条样式 |

]
[
### 图形(Shape)

某个工作表上的单个 Shape 图形对象

| [ID](/api/excel/workbook/Shape#id) | String | 图形 ID |
| --- | --- | --- |
| [Name](/api/excel/workbook/Shape#name) | String | 图形名称 |
| [Title](/api/excel/workbook/Shape#title) | String | 图形标题 |
| [Chart](/api/excel/workbook/Shape#chart) | [Chart](/api/excel/workbook/Chart) | 图表对象 |

| [Delete](/api/excel/workbook/Shape#delete) | undefined | 删除Shape |
| --- | --- | --- |

]
[
### 图表(Chart)

单个图表对象

| [ChartTitle.Text](/api/excel/workbook/Chart#charttitle-text) | String | 图表标题 |
| --- | --- | --- |
| [HasTitle](/api/excel/workbook/Chart#hastitle) | Boolean | 图表标题是否可见 |

| [SetSourceData()](/api/excel/workbook/Chart#setsourcedata) | undefined | 为图表设置源数据区域 |
| --- | --- | --- |

]
[
### 超链接(Hyperlink)

单个超链接对象，有转跳地址和显示文本两个属性，两者可以不相等。

| [Address](/api/excel/workbook/Hyperlink#address) | String | 超链接转跳的地址 |
| --- | --- | --- |
| [TextToDisplay](/api/excel/workbook/Hyperlink#texttodisplay) | String | 超链接显示的文本 |

]
[
### 条件格式集合(FormatConditions)

FormatConditions 集合对象用于控制 Excel 中的条件格式。

| [Count](/api/excel/workbook/FormatConditions#count) | Number | 返回 FormatConditions 集合中的对象数 |
| --- | --- | --- |

| [Add()](/api/excel/workbook/FormatConditions#add) | [FormatCondition](/api/excel/workbook/FormatCondition) | 向 FormatConditions 集合中添加一个条件格式 |
| --- | --- | --- |
| [AddAboveAverage()](/api/excel/workbook/FormatConditions#addaboveaverage) | AboveAverage | 返回表示指定区域的条件格式规则的新 AboveAverage 对象 |
| [AddIconSetCondition()](/api/excel/workbook/FormatConditions#addiconsetcondition) | IconSetCondition | 代表指定区域的图标集条件格式规则 |
| [AddColorScale()](/api/excel/workbook/FormatConditions#addcolorscale) | ColorScale | 该条件格式规则使用单元格颜色中的渐变来指示所选区域中包含的单元格值的相对差异 |
| [AddTop10()](/api/excel/workbook/FormatConditions#addtop10) | Top10 | 该条件格式可以根据指定的截止值查找单元格区域中的最高值和最低值 |
| [AddUniqueValues()](/api/excel/workbook/FormatConditions#adduniquevalues) | UniqueValues | 返回表示指定区域的条件格式规则的新 UniqueValues 对象 |
| [Delete()](/api/excel/workbook/FormatConditions#delete) | undefined | 删除该区域下的条件格式 |
| [Item()](/api/excel/workbook/FormatConditions#item) | [FormatCondition](/api/excel/workbook/FormatCondition) | 从条件格式集合中返回一个条件格式对象 |

]
[
### 条件格式(FormatCondition)

区域内的某个条件格式

| [AppliesTo](/api/excel/workbook/FormatCondition#appliesto) | [Range](/api/excel/workbook/Range) | 应用格式规则的单元格区域 |
| --- | --- | --- |
| [Borders](/api/excel/workbook/FormatCondition#borders) | [Border](/api/excel/workbook/Border) | 返回一个 Borders 集合 |
| [Font](/api/excel/workbook/FormatCondition#font) | [Font](/api/excel/workbook/Font) | 返回一个 Font 对象 |
| [Formula1](/api/excel/workbook/FormatCondition#formula1) | String | 返回与条件格式或者数据有效性相关联的值或表达式 |
| [Formula2](/api/excel/workbook/FormatCondition#formula2) | String | 返回与条件格式或数据有效性验证第二部分相关联的值或表达式 |
| [Interior](/api/excel/workbook/FormatCondition#interior) | Interior | 表示指定对象的内部 |
| [NumberFormat](/api/excel/workbook/FormatCondition#numberformat) | String | 单元格的数字格式 |
| [Operator](/api/excel/workbook/FormatCondition#operator) | [XlFormatConditionOperator](/api/excel/workbook/Enum#xlformatconditionoperator) | 条件格式的运算符 |
| [Priority](/api/excel/workbook/FormatCondition#priority) | Number | 返回或设置条件格式规则的优先级值 |
| [Type](/api/excel/workbook/FormatCondition#type) | [XlFormatConditionType](/api/excel/workbook/Enum#xlformatconditiontype) | 条件格式对象类型 |

| [Modify()](/api/excel/workbook/FormatCondition#modify) | undefined | 更改现有条件格式 |
| --- | --- | --- |
| [ModifyAppliesToRange()](/api/excel/workbook/FormatCondition#modifyappliestorange) | undefined | 设置此格式规则所应用于的单元格区域 |
| [SetFirstPriority()](/api/excel/workbook/FormatCondition#setfirstpriority) | undefined | 将此条件格式规则的优先级值设置为“1” |
| [SetLastPriority()](/api/excel/workbook/FormatCondition#setlastpriority) | undefined | 将此条件格式规则的优先级值增加“1” |

]
[
### 数据有效性规则(Validation)

代表工作表区域的数据有效性规则

| [Add()](/api/excel/workbook/Validation#add) | undefined | 新增数据有效性规则 |
| --- | --- | --- |
| [Modify()](/api/excel/workbook/Validation#modify) | undefined | 修改数据有效性规则 |
| [Delete()](/api/excel/workbook/Validation#delete) | undefined | 删除数据有效性规则 |

]
[
### 工作表函数(WorksheetFunction)

工作表函数对象是 Excel 中的一个内置对象，它包含了许多常用的 Excel 函数，例如 Sum、Average、Min、Max 等。使用 WorksheetFunction 对象可以在 VBA 中调用这些 Excel 函数，以实现对工作表数据的处理和分析。

| [Average()](/api/excel/workbook/WorksheetFunction#average) | Number | 用于计算指定区域内数字的平均值 |
| --- | --- | --- |
| [AverageIf()](/api/excel/workbook/WorksheetFunction#averageif) | Number | 用于计算指定区域内满足给定条件的所有单元格的平均值 |
| [Small()](/api/excel/workbook/WorksheetFunction#small) | Number | 返回数据集中第 k 个最小值 |
| [Large()](/api/excel/workbook/WorksheetFunction#large) | Number | 用于在一个数组或一列数据中返回第 k 个最大值 |
| [Min()](/api/excel/workbook/WorksheetFunction#min) | Number | 用于在一个数组或一列数据中返回最小值 |
| [Max()](/api/excel/workbook/WorksheetFunction#max) | Number | 用于在一个数组或一列数据中返回最大值 |
| [Sum()](/api/excel/workbook/WorksheetFunction#sum) | Number | 对某单元格区域中的所有数字求和 |

]
[
### 枚举(Enum)

枚举类型，存放在 Application 下

]
]
]

#### 区域(Range)

# [Range​](#range)

表示一个单元格、一行、一列、一个包含单个或若干连续单元格区域的选定单元格范围

Range 对象是工作表中部分单元格的集合，可以对工作表中某一区域的单元格进行操作

下方是 Range 对象的使用示例，Range 对象的具体属性和方法请参阅下方的列表。

js
```js
// 选择A1:A10单元格
let range = Application.Range('A1:A10')
// 把A1:A10的值设为对应行号
for (let i = 1; i <= 10; i++) {
  range.Item(i).Value = i.toString()
}
// 打印
printText(range) // 1...10
// 给单元格加上前缀和后缀
range.Each(cell => {
  cell.Value = 'prefix_' + cell.Text + '_suffix'
})
printText(range) //prefix_1_suffix...prefix_10_suffix

function printText(r) {
  r.Each(cell => {
    console.log(cell.Text)
  })
}
```

#### [属性列表​](#属性列表)

| 属性名 | 数据类型 | 简介 |
| --- | --- | --- |
| Count | Number | 区域中单元格的数量 |
| Text | String | 【只读】读取单元格格式化文本 |
| Value/Value2 | any/[][]any | 读写单元格中的值 |
| FormatConditions | FormatConditions | 用于控制 Excel 中的条件格式 |
| Formula | String | 以 A1 样式表示法表示的对象的隐式交叉的公式 |
| FormulaArray | String | 返回或设置区域的数组公式 |
| NumberFormat | String | 获取或者设置区域的数字格式 |
| Hidden | Boolean | 行或者列的隐藏 |
| Interior.Color | String | 内部颜色的十六进制 RGB |
| HorizontalAlignment | Enum.XlHAlign | 设置区域的水平对齐方式 |
| VerticalAlignment | Enum.XlVAlign | 设置区域的垂直对齐方式 |
| WrapText | Boolean | 获取或者设置区域自动换行 |
| IndentLevel | Number | 单元格缩进 |
| MergeArea | Range | 单元格的合并区域 |
| MergeCells | Boolean | 区域内是否存在合并的单元格 |
| Cells | Range | 区域中的单元格集合 |
| Rows | Range | 区域中的行集合 |
| Columns | Range | 区域中的列集合 |
| EntireRow | Range | 区域所在行的整行 |
| EntireColumn | Range | 区域所在列的整列 |
| Row | Number | 区域中第一行的行号 |
| RowEnd | Number | 区域中最后一行的行号 |
| Column | Number | 区域中第一列的列号 |
| ColumnEnd | Number | 区域中最后一列的列号 |
| Borders | Borders | 边框集合对象 |

#### [方法列表​](#方法列表)

| 方法名 | 返回类型 | 简介 |
| --- | --- | --- |
| BorderAround() | undefined | 向区域添加边框，并为新边框设置 Border 对象的 Color、LineStyle 和 Weight 属性 |
| Each() | undefined | 遍历选区所选单元格 |
| Item() | Range | 表示区域中指定的位置 |
| Offset() | undefined | 对指定区域进行迁移操作 |
| Replace() | undefined | 对单元格内文本执行替换操作 |
| Delete() | undefined | 单元格、行、列的删除 |
| Insert() | undefined | 单元格、行、列的新增 |
| InsertImage() | undefined | 插入单元格图片 |
| Merge() | undefined | 合并单元格 |
| UnMerge() | undefined | 取消合并单元格 |
| Address() | String | 获取表示使用宏语言的区域引用的 String 值 |
| AddComment() | undefined | 添加评论 |
| ClearComments() | undefined | 清除区域的评论 |
| Clear() | undefined | 清空指定区域数据和样式 |
| ClearContents() | undefined | 清除区域的内容 |
| ClearFormats() | undefined | 清除区域的样式 |
| ClearHyperlinks() | undefined | 清除区域的超链接样式 |
| Contain() | Boolean | 判断区域是否重叠 |
| Copy() | Boolean | 将当前区域对象复制到剪贴板 |
| Cut() | Boolean | 将当前区域对象粘贴到目标区域 |
| PasteSpecial() | undefined | 将剪贴板中的内容粘贴到指定的单元格或范围 |
| FillLeft() | undefined | 对指定区域中的单元格执行从右往左填充 |
| FillRight() | undefined | 对指定区域中的单元格执行从左往右填充 |
| FillDown() | undefined | 对指定区域中的单元格执行从上往下填充 |
| FillUp() | undefined | 对指定区域中的单元格执行从下往上填充 |
| AutoFill() | undefined | 对指定区域中的单元格执行自动填充 |
| AutoFilter() | undefined | 对指定区域中的单元格执行自动筛选 |
| AutoFit() | undefined | 更改区域中的列宽或行高以达到最佳匹配 |
| Select() | undefined | 选择区域 |
| TextToColumns() | undefined | 将包含文本的一列单元格分解为若干列 |

## [Count​](#count)

区域中列/行/单元格的数量，默认为单元格

#### [数据类型​](#数据类型)

Number - 区域中列/行/单元格的数量

#### [示例​](#示例)

js
```js
// 默认为单元格的Range对象，Range.Count等价于Range.Cells.Count
let range = Application.Range('A1:B2')
console.log(range.Count) // 4
console.log(range.Cells.Count) // 4
console.log(range.Rows.Count) // 2
console.log(range.Columns.Count) // 2
```

## [Text​](#text)

【只读】读取单元格格式化文本

#### [数据类型​](#数据类型-1)

String - 单元格格式化文本

#### [示例​](#示例-1)

js
```js
// 区域对象
let range = Application.Range('A1')

// 【只读】读取单元格格式化文本
console.log(range.Text)

// 修改单元格的内容，Text是只读的，修改需要用Value
range.Value = 'NewVaule'
console.log(range.Text) // NewValue
```

## [Value/Value2​](#value-value2)

读写单元格中的值,支持某个区域读取

#### [数据类型​](#数据类型-2)

读取单个单元格时：any

读取区域单元格：[][]any

#### [示例​](#示例-2)

js
```js
// Value
let range = Application.Range('A1')

range.Value = 'A1'
// 打印A1单元格数据
console.log(range.Value) // A1

range = Range('A1:B2')

// 给A1到B2区域赋值
range.Value = 'WebOffice'

// 打印A1到B2单元格数据
console.log(range.Value) // [["WebOffice","WebOffice"],["WebOffice","WebOffice"]]

// 批量赋值
range.Value = [
  ['A1', 'B1'],
  ['A2', 'B2']
]

// 打印A1到B2单元格数据
console.log(range.Value) // [["A1","B1"],["A2","B2"]]
```

## [FormatConditions​](#formatconditions)

FormatConditions 集合对象用于控制 Excel 中的条件格式。条件格式是一种在工作表中格式化单元格的方法，可以根据单元格的值、公式或其他条件自动应用格式，使数据更易于理解和分析

#### [数据类型​](#数据类型-3)

FormatConditions- FormatConditions 集合对象

#### [示例​](#示例-3)

js
```js
// 获取FormatConditions对象
const formatConditions = Range('A:A').FormatConditions
// 获取条件格式对象数量
const count = formatConditions.Count

console.log(count)
```

## [Formula​](#formula)

Formula 支持读取和设置单元格内容，但细节有些不同

读取情况下：Formula 读取的是公式,而 Text 读取的是值。
赋值情况下：Formula 只支持赋值单个单元格,而 Value 支持赋值区域
#### [数据类型​](#数据类型-4)

String - 隐式交叉的公式

#### [示例​](#示例-4)

js
```js
// 设置A1=1,B1=2
Application.Range('A1').Value = '1'
Application.Range('B1').Value = '2'

// 设置A2单元格的公式: =A1+B1
const range = Application.Range('A2')
range.Formula = '=A1+B1'
// 读取A2单元格的值: 3
console.log(range.Text) // 3
// 读取A2单元格的公式: =A1+B1
console.log(range.Formula) // =A1+B1
```

## [FormulaArray​](#formulaarray)

相较于 Formula，支持设置范围数据，其他与 Formula 一致

#### [数据类型​](#数据类型-5)

String - 区域的数组公式

#### [示例​](#示例-5)

js
```js
let range = Application.Range('A1:B2')

// 设置A1到B2的单元格的值: 100
range.FormulaArray = '=Sum(A1:C3)'

range.Each(function (item) {
  console.log(item.Text) //=Sum(A1:C3)，=Sum(A1:C3)，=Sum(A1:C3)，=Sum(A1:C3)
})
```

## [NumberFormat​](#numberformat)

获取或者设置区域的数字格式。

在获取上，如果指定区域中的所有单元格的数字格式不一致，则此属性返回第一个单元格的数字格式。

在设置上，为了让用户设置更方便，我们列举了WebOffice上常用的值：

常规：
G/通用格式
数值：
0.00_);[红色](0.00)
货币：
￥#,##0.00_);[红色](￥#,##0.00)
会计专用：
_ ￥* #,##0.00_ ;_ ￥* -#,##0.00_ ;_ ￥* "-"??_ ;_ @_
短日期：
yyyy/m/d;@
长日期：
yyyy"年"m"月"d"日";@
时间：
h:mm:ss;@
百分比：
0.00%
分数：
## ?/?
科学技术：
0.00E+00
文本：
@
千位分隔样式：
_ * #,##0.00_ ;_ * -#,##0.00_ ;_ * "-"??_ ;_ @_
这样在调用 API 设置数字格式的时候，可以复制上面的值进行设置，例如设置数字格式为常规：Range.NumberFormat = 'G/通用格式'。

#### [数据类型​](#数据类型-6)

String - 区域的数字格式

#### [示例​](#示例-6)

js
```js
// 区域对象
let range = Application.Range('A1')

// 获取区域数字格式
console.log('当前区域数字格式：', range.NumberFormat) //当前区域数字格式： G/通用格式

// 设置区域数字格式为文本
range.NumberFormat = '@'
console.log('当前区域数字格式：', range.NumberFormat) //当前区域数字格式： @
```

## [Hidden​](#hidden)

行或者列的隐藏，该属性只写，无法获取原先的列是否隐藏

#### [数据类型​](#数据类型-7)

Boolean - 行或者列的隐藏

#### [示例​](#示例-7)

js
```js
// 区域对象
let range = Application.Range('A1')

// 获取整列
let entireColumn = range.EntireColumn

// 隐藏该列
entireColumn.Hidden = true
```

## [Interior.Color​](#interior-color)

区域内部颜色

#### [数据类型​](#数据类型-8)

Number - 内部颜色的十六进制 RGB

#### [示例​](#示例-8)

js
```js
// 内部属性对象
let interior = Application.Range('A1').Interior

// 获取内部颜色
console.log(interior.Color) // 16777215

// 设置内部颜色为红色
interior.Color = RGB(255, 0, 0)
```

## [HorizontalAlignment​](#horizontalalignment)

设置区域的水平对齐方式，设置的值可以是Enum.XlHAlign中的值。

#### [数据类型​](#数据类型-9)

Enum.XlHAlign- 区域的水平对齐方式

#### [示例​](#示例-9)

js
```js
// 区域对象
let range = Application.Range('A1')

// 打印对齐方式
console.log(range.HorizontalAlignment) // 1

// 设置对齐方式：居中
range.HorizontalAlignment = Application.Enum.XlHAlign.xlHAlignCenter

// 打印对齐方式
console.log(range.HorizontalAlignment) // -4108，即居中枚举值
```

## [VerticalAlignment​](#verticalalignment)

设置区域的垂直对齐方式，设置的值可以是Enum.XlVAlign中的值。

#### [数据类型​](#数据类型-10)

Enum.XlVAlign- 区域的垂直对齐方式

#### [示例​](#示例-10)

js
```js
// 区域对象
let range = Application.Range('A1')

// 设置对齐方式：底部对齐
range.VerticalAlignment = Application.Enum.XlVAlign.xlVAlignBottom

// 获取对齐方式
console.log(range.VerticalAlignment) //-4107，即向下垂直对齐方式枚举值
```

## [IndentLevel​](#indentlevel)

设置单元格缩进

#### [数据类型​](#数据类型-11)

Number - 单元格缩进

#### [示例​](#示例-11)

js
```js
// 区域对象
let range = Application.Range('A1')

// 写入值到单元格中
console.log(range.IndentLevel) // 0
range.IndentLevel = 10
console.log(range.IndentLevel) // 10
```

## [WrapText​](#wraptext)

获取或者设置区域自动换行

#### [数据类型​](#数据类型-12)

Boolean - 区域是否自动换行

#### [示例​](#示例-12)

js
```js
// 区域对象
let range = Application.Range('A1')
// 获取区域自动换行
console.log('区域是否自动换行：', range.WrapText) // true
range.WrapText = false
console.log('区域是否自动换行：', range.WrapText) // false
```

## [MergeArea​](#mergearea)

单元格的合并区域

#### [数据类型​](#数据类型-13)

Range - 该区域内合并单元格的 Range 对象

#### [示例​](#示例-13)

js
```js
// 区域对象
let range = Application.Range('A1:D2')
// 合并单元格
range.Merge()
// 将该合并区域值设置为merge
range.MergeArea.Value = 'merge'
```

## [MergeCells​](#mergecells)

区域内是否存在合并的单元格

#### [数据类型​](#数据类型-14)

Boolean - 是否存在合并的单元格

#### [示例​](#示例-14)

js
```js
// 区域对象
let range = Application.Range('A1:D2')
// 合并单元格
range.Merge()
// 区域内是否存在合并的单元格
console.log(range.MergeCells) // true
```

## [Cells​](#cells)

区域中的所有单元格集合，返回一个Range对象（可使用 Range 相关的属性、方法）

#### [数据类型​](#数据类型-15)

Range- 区域中的所有单元格集合

#### [示例​](#示例-15)

js
```js
// 区域对象
let range = Application.Range('A1:B2')

// 单元格对象
let cells = range.Cells

// 取第一个单元格
let cell = cells.Item(1)
console.log(cell.Address()) // $A$1
```

## [Rows​](#rows)

区域中的行，返回一个 Range 对象（可使用 Range 相关的属性、方法）

#### [数据类型​](#数据类型-16)

Range- 区域中的所有行集合

#### [示例​](#示例-16)

js
```js
// A1:B2最一共有2行，因此返回2
console.log(Application.Range('A1:B2').Rows.Count) // 2

// A1:D4最一共有4行，因此返回4
console.log(Application.Range('A1:D4').Rows.Count) // 4
```

## [Columns​](#columns)

区域中的所有列，返回一个Range对象（可使用 Range 相关的属性、方法）

#### [数据类型​](#数据类型-17)

Range- 区域中的所有列集合

#### [示例​](#示例-17)

js
```js
// A1:B2最一共有2列，因此返回2
console.log(Application.Range('A1:B2').Columns.Count) // 2

// A1:D4最一共有4列，因此返回4
console.log(Application.Range('A1:D4').Columns.Count) // 4
```

## [EntireRow​](#entirerow)

包含指定区域的整行

#### [数据类型​](#数据类型-18)

Range- 包含指定区域的整行

#### [示例​](#示例-18)

js
```js
// 一行默认长度为16384
console.log(Application.Range('A1').EntireRow.Count) // 16384

// 两行
console.log(Application.Range('A1:B2').EntireRow.Count) // 32768
```

## [EntireColumn​](#entirecolumn)

包含指定区域的整列

#### [数据类型​](#数据类型-19)

Range- 包含指定区域的整列

#### [示例​](#示例-19)

js
```js
// 一列默认长度为1048576
console.log(Application.Range('A1').EntireColumn.Count) // 1048576

// 两列
console.log(Application.Range('A1:B2').EntireColumn.Count) // 2097152
```

## [Row​](#row)

区域中第一行的行号

#### [数据类型​](#数据类型-20)

Number - 区域中第一行的行号

#### [示例​](#示例-20)

js
```js
let range = Application.Range('A1:D4')
// A1:D4第一行是1行，因此返回1
console.log(range.Row) // 1
```

## [RowEnd​](#rowend)

区域中最后一行的行号

#### [数据类型​](#数据类型-21)

Number - 区域中最后一行的行号

#### [示例​](#示例-21)

js
```js
let range = Application.Range('A1:D4')
// A1:D4最后一行是4行，因此返回4
console.log(range.RowEnd) // 4
```

## [Column​](#column)

区域中最左列的列号

#### [数据类型​](#数据类型-22)

Number - 区域中最左列的列号

#### [示例​](#示例-22)

js
```js
let range = Application.Range('A1:D4')
// A1:D4最左列是A列，因此返回1
console.log(range.Column) // 1
```

## [ColumnEnd​](#columnend)

区域中最右列的列号

#### [数据类型​](#数据类型-23)

Number - 区域中最右列的列号

#### [示例​](#示例-23)

js
```js
let range = Application.Range('A1:D4')
// A1:D4最右列是D列，因此返回4
console.log(range.ColumnEnd) // 4
```

## [Borders​](#borders)

边框对象

#### [数据类型​](#数据类型-24)

Border集合对象 - 边框对象集合

#### [示例​](#示例-24)

js
```js
// 边框对象
let borders = Application.Range('A1').Borders

// 将A1的左上角到右下角的边框颜色设置为红色
borders.Item(Application.Enum.XlBordersIndex.xlDiagonalDown).Color = '#FF0000'
```

### [Borders.Item()​](#borders-item)

单个边框对象，代表单元格区域或样式的边框之一

#### [参数​](#参数)

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| Index | Enum |  | 是 | 指定要检索的边框，参考Enum.XlBordersIndex |

#### [返回类型​](#返回类型)

Border- 单个边框对象，代表单元格区域或样式的边框之一

#### [示例​](#示例-25)

js
```js
// 边框对象
let borders = Application.Range('A1').Borders

// 将A1的左上角到右下角的边框颜色设置为黄色
borders.Item(Application.Enum.XlBordersIndex.xlDiagonalDown).Color = '#FFFF00'
```

## [BorderAround()​](#borderaround)

向区域添加边框，并为新边框设置 Border 对象的 Color、LineStyle 和 Weight 属性

#### [参数​](#参数-1)

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| LineStyle | LineStyle | null | 否 | XlLineStyle 的常量之一，指定边框的线条样式 |
| BorderWeight | BorderWeight | null | 否 | 边框粗细 |
| Color | string | null | 否 | 边框颜色，以 RGB 值表示，例如#ff0000 |

#### [示例​](#示例-26)

js
```js
// 为B2单元格设置红色虚线
Range('B2').BorderAround(xlDash, xlHairline, '#ff0000')
// 为C3单元格设置绿色点划相间线
Range('C3').BorderAround(xlDashDot, xlHairline, '#00ff00')
```

## [Each()​](#each)

遍历选区所选单元格，建议使用此函数时不要涉及插入或删除行列，否则可能会导致不符合预期的结果

#### [参数​](#参数-2)

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| callback | Function | null | 是 | 类似 JS 数组的 forEach |

#### [示例​](#示例-27)

js
```js
// 区域对象
let range = Application.Range('A1:D2')

// 编辑选区单元格
range.Each(function (item) {
  //打印单元格的地址
  console.log(item.Address()) // $A$1...$D$2
})
```

## [Item()​](#item)

返回一个 Range 对象，表示区域中指定的位置

#### [参数​](#参数-3)

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| RowIndex | Number |  | 是 | 如果提供了第二个参数，则返回的单元格的相对行号。如果未提供第二个参数，则为要返回的子范围的索引 |
| ColumnIndex | Number |  | 否 | 要返回的单元格的相对列号 |

#### [返回类型​](#返回类型-1)

Range - 对应的单元格对象

#### [示例​](#示例-28)

js
```js
// 区域对象
const range = Application.Range('A1:D2')
// 区域子项：B2
console.log(range.Item(2, 2).Address()) //$B$2
```

## [Offset()​](#offset)

对指定区域进行偏移，返回偏移后的 Range

#### [参数​](#参数-4)

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| RowOffset | Number | 0 | 否 | 区域偏移的行数：可以是正值、负值或零。正值表示向下偏移，负值表示向上偏移 |
| ColumnOffset | Number | 0 | 否 | 区域偏移的列数：可以是正值、负值或零。正值表示向右偏移，负值表示向左偏移 |

#### [示例​](#示例-29)

js
```js
// 区域对象
let range = Application.Range('A1:D2')
// 区域子项：B2
console.log(range.Address()) //$A$1:$D$2
// 对指定区域进行偏移
let newRange = range.Offset(2, 2)
// 区域子项：B2
console.log(newRange.Address()) //$C$3:$F$4
```

## [Replace()​](#replace)

对单元格内文本执行替换操作

#### [参数​](#参数-5)

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| What | String | 无 | 是 | 希望搜索的字符串。 |
| Replacement | String | 无 | 是 | 替换字符串。 |
| LookAt | Enum | xlWhole | 否 | 可以是下列 XlLookAt 常量之一：xlWhole 或 xlPart。详情看示例 |
| SearchOrder | Enum | xlByRows | 否 | 可以是以下 XlSearchOrder 常量之一：xlByRows 或 xlByColumns。 详情看示例 |
| MatchCase | Boolean | false | 否 | 如果为 True，则搜索区分大小写。 |
| MatchByte | Boolean | false | 否 | 如果是 true，区分全半角符号 |

#### [示例​](#示例-30)

js
```js
// 区域对象
let range = Application.Range('A1:D2')
// 把所有小山全部替换成大山
range.Replace('小山', '大山', xlPart)
// 单元格内容为小山的单元格替换成大山
range.Replace('小山', '大山', xlWhole)
// 把JinXiaoShan替换成金小山，区分大小写
range.Replace('JinXiaoShan', '金小山', xlPart, xlByRows, true)
// 把JinXiaoShan替换成金小山，不区分大小写
range.Replace('JinXiaoShan', '金小山', xlPart, xlByRows, false)
// 把JinXiaoShan替换成金小山，区分全角半角
range.Replace('JinXiaoShan', '金小山', xlPart, xlByRows, true, true)
// 把JinXiaoShan替换成金小山，不区分全角半角
range.Replace('JinXiaoShan', '金小山', xlPart, xlByRows, true, false)
```

## [Delete()​](#delete)

单元格、行、列的删除.删除单元格时,右侧单元格左移

#### [示例：删除行​](#示例-删除行)

js
```js
// 区域对象
let range = Application.Range('B5:D10')
// 选择区域的所有行并删除，例子删除5-10行
range.EntireRow.Delete()
```

#### [示例：删除列​](#示例-删除列)

js
```js
let range = Application.Range('B5:D10')
// 选择区域的所有行并删除，例子删除2-4列
range.EntireColumn.Delete()
```

## [Insert()​](#insert)

单元格、行、列的新增

#### [示例：新增行​](#示例-新增行)

js
```js
let range = Application.Range('B5:D10')
// 选择区域的所有行并在最上行新增一行，例子在第5行上方插入一行，原5行下移为6行
range.EntireRow.Insert()
```

#### [示例：新增列​](#示例-新增列)

js
```js
let range = Application.Range('B5:D10')
// 选择区域的所有列并在最左列新增一列，例子在第2列左侧插入一列，原2列右移为3列
range.EntireColumn.Insert()
```

## [InsertImage()​](#insertimage)

插入单元格图片

#### [参数​](#参数-6)

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| dataURL | string | undefined | 是 | base64 字符串形式的图片 |

#### [示例​](#示例-31)

js
```js
// 获取E1单元格
const range = Range('E1')
// 向目标单元格插入图片
range.InsertImage(
  'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAMAAABEpIrGAAACH1BMVEUAAAA9kP8mpv9Fv/8QT94TXeQAfv8trP0xr/0VUdoAcv8LgPs1sf4AgP4QUuURTuEVl/ksq/wAg/84tf4NUt80sf0bnPsPUuAAYfsAgP4UTtYAY/8GXPQRkv1Jxv8AZP86tv9Bvv9Gv/9Atv8KiPUOVesAZ+wbS8wUl/kjo/sAYv8PVekjovxBu/8AY/8trPxCu/8AW/AAcvAxsP0npvsYTdElo/oQT9o+uP8AYv8KVOIHdugAYfwAfPwbfe49uf4XmfkcdOopo/oAYv86tf4AYf8TT9cZaeYVTdUJVOUcnvsAYv8Agv9Iwv8AYv9Jwf8bSs0gofoAY/8Ag/9Gv/8en/oAYf9Fv/8AWuQZS80AdO5BvP8AYv9DvP8Umfo9uv8LVOUAgf8amfk/uf8AY/8Ag/8Ahf8TUdgVT9UNWOknov8trPwxr/wmpvspqfs1sv0io/s5tf0AgP8bnfoMV+4AYPoAYv8AYf4AYfwAe/sen/oAd/gAXvYAff0LW/MLWfEOUN0PT9oSTtZAu/49uP0qqfweoPoXmvkAdPUAXPIAcfEKUuIWTNIXS88KXPYAWusOVeoPVOkPUuYQUeQQT+IMUeAUTdMZSs09uP4Aa+kAWOgAaeQIVOQIU+QAV+MRTd8STd0RTtgSl/kAW+8NWO8Abu0AbuwBZNsAg/8AYPwAc/MIVecAZuEAVeAAWt8AZN0JgvAIfOsJV+sTVuIAU9qil9AQAAAAa3RSTlMAAwb++A3+bkkkGxL8+Pf36tnTubCYjol4dmVjRjEvLysZFA7+/Pv6+fj39/b19fPw8O/u7Ovq6Ofn5+Xj4uDa2trZ1dPS0s/Ny7+9vLi4trapoqGgoJ2XkI+KgoJycGRhX1BNS0pHQj06IcCB3jkAAAIBSURBVDjLdZCHctpAFEWfCMU1ce8tTu+99957770nJFGwkE1sgyMTAginYbkAIQSD494+0G+ltTAacWZnNLt7dO/Og3Tqzl2og8wsvroAudEI+iy5s+6LzMZ7Bp1r5sFOQegRenAJwoEnoCXr6Mc0ihvSrhuK32vJK21Mva00760OWyz0KZYNbzJQmKXUF1o7rZ24rGngweETIGOwbH2nw/ZDUaM6grJVHzSs3R+NoqDy+lKLuwWX2618r03/QRalhNzbewo+qZyuZX4TVMFQubm3f3xy5VeZg49xsL8IC+mUH+1zeLyJ3pHkRCwWK7hvIGc/CYrwvIhtbXZ4hgYwJLm0LFf56QdhDRAq8wdZMypD3kT/SBJo6nfCMmWTc2WQlUO8GDIn2GSAUnuWGERJpIS4LR4HlWdFNITua9pl5s3h5vJ81tzscAAhuyTc3kZQ51CxqasvMmDGENyZyo1BX2iqLRAI0L7qvU2fv6Ex5mVbgana4XIGfb7QsI0KNcc4O4/Gv67I6Jjn5cmOvy4nGuFwaFipqFjN2YlBQkZn/N0dxAhiSfjUKzqHyxwq/0lIX0T0d6OAGb7dD0Gl/gzH8UqN6JcNp7HcBPN5esTO8U2oSCIpcZVkgwamepedxxpJROP4C9DBdHc9vlWSxG1VDOiTc53jpRW3TJCZ+vMXNeWzDz4DoNZyqecAAAAASUVORK5CYII='
)
```

## [Merge()​](#merge)

合并单元格

#### [参数​](#参数-7)

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| Across | boolean | false | 是 | 如果设置为 true，则将指定区域中每一行的单元格合并为一个单独的合并单元格 |

#### [示例​](#示例-32)

js
```js
const range = Application.Range('A1:D2')
// 合并单元格
range.Merge()
```

## [UnMerge()​](#unmerge)

取消合并单元格

#### [参数​](#参数-8)

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| CancelCenter | Boolean | false | 否 | 是否合并居中 |

#### [示例​](#示例-33)

js
```js
const range = Application.Range('A1:D2')

// 取消合并单元格
range.UnMerge()
```

## [Address()​](#address)

获取表示使用宏语言的区域引用的 String 值

#### [参数​](#参数-9)

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| RowAbsolute | Boolean |  | 否 | 若为 true，以绝对引用的形式返回引用的行部分。默认值为 true |
| ColumnAbsolute | Boolean |  | 否 | 若为 true，以绝对引用的形式返回引用的列部分。默认值为 true |
| ReferenceStyle | Enum |  | 否 | 引用样式。默认值为 xlA1，更多可看Enum.XlReferenceStyle |
| External | Boolean |  | 否 | 若为 true，返回外部引用。若为 false，返回本地引用。默认值为 false |
| RelativeTo | Range |  | 否 | 如果 RowAbsolute 和 ColumnAbsolute 为 false，且 ReferenceStyle 是 xlR1C1，则必须为相对引用包含一个起点。此参数是一个定义起点的 Range 对象 |

#### [返回类型​](#返回类型-2)

String - 表示使用宏语言的区域引用的值

#### [示例​](#示例-34)

js
```js
// 区域对象
let range = Application.Range('A1')

// 获取宏语言的区域引用的 String 值
console.log('address1：', range.Address())

console.log('address2：', range.Address(false, false))

console.log(
  'address3：',
  range.Address(true, true, Application.Enum.XlReferenceStyle.xlR1C1)
)
```

## [AddComment()​](#addcomment)

添加评论

#### [参数​](#参数-10)

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| Text | String |  | 否 | 评论文本 |

#### [示例​](#示例-35)

js
```js
let range = Application.Range('A1')
// 给 A1 区域添加评论
range.AddComment('WebOffice')
```

## [ClearComments()​](#clearcomments)

清除区域的评论

#### [示例​](#示例-36)

js
```js
let range = Application.Range('A1')
// 清除区域的评论
range.ClearComments()
```

## [Clear()​](#clear)

清空指定区域数据和样式

#### [示例​](#示例-37)

js
```js
const range = Range('A1:B2')
range.Value = [
  ['A1', 'B1'],
  ['A2', 'B2']
]
range.Clear()
console.log(range.Value) //[["",""],["",""]]
```

## [ClearContents()​](#clearcontents)

清除区域的内容

#### [示例​](#示例-38)

js
```js
let range = Application.Range('A1')
// 清除区域的内容
range.ClearContents()
```

## [ClearFormats()​](#clearformats)

清除区域的样式

#### [示例​](#示例-39)

js
```js
let range = Application.Range('A1')
// 清除区域的样式
range.ClearFormats()
```

## [ClearHyperlinks()​](#clearhyperlinks)

清除区域的超链接样式

#### [示例​](#示例-40)

js
```js
let range = Application.Range('A1')
// 清除区域的超链接样式
range.ClearHyperlinks()
```

## [Contain()​](#contain)

判断区域是否重叠

#### [参数​](#参数-11)

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| Range | Range |  | 是 | Range 对象，另一块区域 |

#### [返回类型​](#返回类型-3)

Boolean - 区域是否重叠

#### [示例​](#示例-41)

js
```js
// 区域对象
let range = Application.Range('A1:D2')

// 第二块区域对象
let newRange = Application.Range('A1:B4')

// 判断是否重叠
console.log(range.Contain(newRange))
```

## [Copy()​](#copy)

将当前区域对象复制到剪贴板

#### [返回类型​](#返回类型-4)

Boolean - 是否复制成功

#### [示例​](#示例-42)

js
```js
// 复制A1
Range('A1').Copy()
// 粘贴到B1
Range('B1').PasteSpecial()
```

## [Cut()​](#cut)

将当前区域对象粘贴到目标区域

#### [参数​](#参数-12)

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| Destination | Range | null | 是 | 目标区域对象 |

#### [返回类型​](#返回类型-5)

Boolean - 是否复制成功

#### [示例​](#示例-43)

js
```js
// 将A1区域的数据粘贴到A2区域
Range('A1').Cut(Range('A2'))
```

## [PasteSpecial()​](#pastespecial)

用于将剪贴板中的内容粘贴到指定的单元格或范围。它可以以多种方式粘贴数据，包括值、格式、公式等。可以通过指定参数来控制粘贴的方式，如粘贴值、粘贴格式、粘贴公式等

注意

在使用 PasteSpecial()粘贴数据之前，请确保已将源区域的值复制到剪贴板，如果剪贴板内没有任何区域对象，则会抛出剪贴板函数调用异常。

您可先通过Copy()函数将区域对象复制到剪贴板，然后再使用 PasteSpecial()函数进行粘贴，并可以选择性地应用不同的粘贴选项。

#### [参数​](#参数-13)

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| Paste | XlPasteType | xlPasteAll | 否 | 粘贴类型，例如 xlPasteAll 或 xlPasteValues |
| Operation | XlPasteSpecialOperation | xlPasteSpecialOperationNone | 否 | 粘贴操作，例如 xlPasteSpecialOperationAdd |
| SkipBlanks | Boolean | false | 否 | 如果为 true，则不将剪贴板上区域中的空白单元格粘贴到目标区域中 |
| Transpose | Boolean | false | 否 | 如果为 true ，则表示在粘贴区域时转置行和列 |

#### [返回类型​](#返回类型-6)

Undefined

#### [示例一​](#示例一)

粘贴值：使用此方法以值形式将剪贴板中的内容粘贴到指定单元格或范围。例如，将 B1 的值粘贴到 A1

js
```js
Range('B1').Copy()
Range('A1').PasteSpecial(xlPasteValues)
```

#### [示例二​](#示例二)

粘贴格式：将 A1:B2 区域的格式粘贴至 C1:D2

js
```js
Range('A1:B2').Copy()
Range('C1:D2').PasteSpecial(xlPasteFormats)
```

#### [示例三​](#示例三)

粘贴值和格式：将 A1:B2 区域的值和格式粘贴至 A3:B4

js
```js
Range('A1:B2').Copy()
Range('A3:B4').PasteSpecial(xlPasteAll)
```

#### [示例四​](#示例四)

转置粘贴：将 A1:B2 区域的值转置粘贴至 C1:D2（转置粘贴后，行和列会互换）

js
```js
Range('A1:B2').Copy()
Range('C1:D2').PasteSpecial(
  xlPasteValues,
  xlPasteSpecialOperationNone,
  false,
  true
)
```

## [FillLeft()​](#fillleft)

对指定区域中的单元格执行从右往左填充,填充值为最右列的值

#### [示例​](#示例-44)

js
```js
const range = Range('A1:B2')
range.Value = [
  ['A1', 'B1'],
  ['A2', 'B2']
]
range.FillLeft()
console.log(range.Value) // [["B1","B1"],["B2","B2"]]
```

## [FillRight()​](#fillright)

对指定区域中的单元格执行从左往右填充,填充值为最左列的值

#### [示例​](#示例-45)

js
```js
const range = Range('A1:B2')
range.Value = [
  ['A1', 'B1'],
  ['A2', 'B2']
]
range.FillRight()
console.log(range.Value) // [["A1","A1"],["A2","A2"]]
```

## [FillDown()​](#filldown)

对指定区域中的单元格执行从上往下填充,填充值为最上列的值

#### [示例​](#示例-46)

js
```js
const range = Range('A1:B2')
range.Value = [
  ['A1', 'B1'],
  ['A2', 'B2']
]
range.FillDown()
console.log(range.Value) // [["A1","B1"],["A1","B1"]]
```

## [FillUp()​](#fillup)

对指定区域中的单元格执行从下往上填充,填充值为最下列的值

#### [示例​](#示例-47)

js
```js
const range = Range('A1:B2')
range.Value = [
  ['A1', 'B1'],
  ['A2', 'B2']
]
range.FillUp()
console.log(range.Value) //[["A2","B2"],["A2","B2"]]
```

## [AutoFill()​](#autofill)

对指定区域中的单元格执行自动填充

#### [参数​](#参数-14)

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| Destination | Range |  | 否 | 目标区域。目标区域必须包含源区域 |
| Type | Enum | Enum.XlAutoFillType.xlFillDefault | 否 | 填充类型，详细可见Enum.XlAutoFillType |

#### [示例​](#示例-48)

js
```js
// 区域对象
let range = Application.Range('A1')
range.Value = 1
// 要填充的单元格
let fillRange = Application.Range('A1:A20')
// 对指定区域中的单元格执行自动填充
range.AutoFill(fillRange)
fillRange.Each(item => {
  console.log(item.Text()) // 自动填充为：1,2,3,4...20
})
```

## [AutoFilter()​](#autofilter)

对指定区域的单元格执行自动筛选

#### [参数​](#参数-15)

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| Field | Variant | null | 否 | 指定想要基于筛选的字段的整数偏移量。从列表的左侧算起，最左侧的字段是 1 |
| Criteria1 | Variant | null | 否 | 指定判断条件。使用“=”查找空字段，或者使用“<>”查找非空字段。如果忽略该参数，那么判断是全部。如果参数 Operator 是 xlTop10Items，那么参数 Criterial1 指定项目的数量 |
| Operator | XlAutoFilterOperator | null | 否 | 指定筛选的类型，通过枚举值XlAutoFilterOperator来指定 |
| Criteria2 | Variant | null | 否 | 第二个判断条件。与 Criteria1 和 Operator 一起组合成复合筛选条件。 也用作日期字段的单一条件（按日、月或年筛选）。 后跟一个数组，该数组用于描述筛选 Array(Level, Date)。 其中，Level 为 0-2（年、月、日），Date 为筛选期内的一个有效日期 |
| SubField | Variant | null | 否 | 对其应用条件的数据类型中的字段（例如，来自地理位置的“人口”字段或来自股票的“交易量”字段）。省略此值目标是“（显示值）” |
| VisibleDropDown | Variant | true | 否 | 如果为true，则显示已筛选字段的AutoFilter下拉箭头。 如果为false，则隐藏已筛选字段的AutoFilter下拉箭头。 默认情况下为true |

#### [示例一​](#示例一-1)

将 B 列数值排在后面的 11 行筛选显示

js
```js
// 未定义筛选区域时，默认使用UsedRange作为筛选区域
const filterRange = ActiveSheet.UsedRange
// 定义筛选区域B列
let filterField = 2
// 获取筛选的数据区域
const filterColumnRange = filterRange.Columns(filterField)
// 计算数据区域的行数，需要减去表头
const filterRangeCount = filterColumnRange.Rows.Count - 1
// 获取后11名数据
const bottom11Value = 11
// 数据区域中可能没有11行数据
const bottomNum =
  filterRangeCount < bottom11Value ? filterRangeCount : bottom11Value
// 获取筛选的数据里第11小的值
let bottom11 = WorksheetFunction.Small(filterColumnRange, bottomNum)
// 应用筛选
filterRange.AutoFilter(filterField, '<=' + bottom11)
```

#### [示例二​](#示例二-1)

将第 1 列为"1 班"，第 5 列的值小于 60 的单元格填充为红色，第 5 列的值大于 90 的单元格填充为绿色

js
```js
// 选择筛选区域为已使用的区域
const filterRange = ActiveSheet.UsedRange
// 设置筛选第1列
const filterColumn1 = 1
// 设置筛选值为"1班"
let filterValue = '1班'
// 应用筛选
filterRange.AutoFilter(filterColumn1, filterValue)
// 选择要筛选的列号为第5列
const filterColumn2 = 5
// 选择筛选的值为小于"60"
filterValue = '<60'
// 应用筛选
filterRange.AutoFilter(filterColumn2, filterValue)
// 获取除标题行外的数据区域
const dataRange = ActiveSheet.UsedRange.Offset(1).Resize(
  ActiveSheet.UsedRange.Rows.Count - 1
)
// 定义列的区域
const columnRange = dataRange.Columns(filterColumn2)
// 获取筛选后显示的数据
const visibleRange = columnRange.SpecialCells(xlCellTypeVisible)
// 定义颜色红色
const redColor = RGB(255, 0, 0)
// 设置颜色为红色
visibleRange.Interior.Color = redColor
// 清除第5列的筛选条件
filterRange.AutoFilter(filterColumn2)
// 选择筛选的值为大于"90"
const filterValue2 = '>90'
// 应用筛选
filterRange.AutoFilter(filterColumn2, filterValue2)
// 获取筛选后显示的数据
const visibleRange2 = columnRange.SpecialCells(xlCellTypeVisible)
// 定义颜色绿色
const greenColor = RGB(0, 255, 0)
// 设置颜色为绿色
visibleRange2.Interior.Color = greenColor
// 清除第1列的筛选条件
filterRange.AutoFilter(filterColumn1)
// 清除第5列筛选条件
filterRange.AutoFilter(filterColumn2)
```

#### [示例三​](#示例三-1)

筛选出 A 列是 2023 年 4 月 10 日和 2022 年 12 月并且 C 列值高于 10000 的数据/信息/情况/商品/行

js
```js
// 未定义筛选区域时，默认使用UsedRange作为筛选区域
const filterRange = ActiveSheet.UsedRange
// 定义筛选区域A列
let filterField = 1
// 定义日期筛选时间
let criteria = new Array()
// 定义筛选年、月、日的枚举值
const FilterDateTimeEnum = { YEAR: 0, MONTH: 1, DAY: 2 }
// 增加筛选日期为2023年
criteria.push(FilterDateTimeEnum.DAY, '2023/4/10')
// 增加筛选日期为2022年12月
criteria.push(FilterDateTimeEnum.MONTH, '2022/12/1')
// 应用筛选
filterRange.AutoFilter(filterField, undefined, xlFilterValues, criteria)
// 定义筛选区域为C列
filterField = 3
// 设置筛选条件为大于10000
let criteria2 = '>10000'
// 应用筛选
filterRange.AutoFilter(filterField, criteria2)
```

#### [示例四​](#示例四-1)

筛选出 A 列日期是 2023 年、2022 年 3 月、2022 年 4 月、2022 年 5 月 1 日、2022 年 5 月 2 日的数据

js
```js
// 未定义筛选区域时，默认使用UsedRange作为筛选区域
const filterRange = ActiveSheet.UsedRange
// 定义筛选区域A列
let filterField = 1
// 定义日期筛选时间
let criteria = new Array()
// 定义筛选年、月、日的枚举值
const FilterDateTimeEnum = { YEAR: 0, MONTH: 1, DAY: 2 }
// 增加筛选日期为2023年
criteria.push(FilterDateTimeEnum.YEAR, '2023/1/1')
// 增加筛选日期为2022年3月
criteria.push(FilterDateTimeEnum.MONTH, '2022/3/1')
// 增加筛选日期为2022年4月
criteria.push(FilterDateTimeEnum.MONTH, '2022/4/1')
// 增加筛选日期为2022年5月1日
criteria.push(FilterDateTimeEnum.DAY, '2022/5/1')
// 增加筛选日期为2022年5月2日
criteria.push(FilterDateTimeEnum.DAY, '2022/5/2')
// 应用筛选
filterRange.AutoFilter(filterField, undefined, xlFilterValues, criteria)
```

#### [示例五​](#示例五)

帮我将 D 列是本季度同时 C 列是去年到现在截止时间为一年，A 列是'商品 A'显示出来

js
```js
// 未定义筛选区域时，默认使用UsedRange作为筛选区域
const filterRange = ActiveSheet.UsedRange
// 定义筛选区域D列
let filterField = 4
// 设置筛选 xlFilterThisQuarter， 可选值xlFilterLastQuarter、xlFilterNextQuarter、xlFilterThisQuarter
let criteria = xlFilterThisQuarter
// 应用筛选
filterRange.AutoFilter(filterField, criteria, xlFilterDynamic)
// 定义筛选字段为C列
filterField = 3
// 设置筛选时间为 xlFilterYearToDate 过去到今天为止一年的时间
criteria = xlFilterYearToDate
// 应用筛选
filterRange.AutoFilter(filterField, criteria, xlFilterDynamic)
// 定义筛选区域为A列
filterField = 1
// 设置筛选条件'商品A'
criteria = '商品A'
// 应用筛选
filterRange.AutoFilter(filterField, criteria)
```

#### [示例六​](#示例六)

筛选第二列包含'张三'或者开头等于'李四'的数据

js
```js
// 未定义筛选区域时，默认使用UsedRange作为筛选区域
const filterRange = ActiveSheet.UsedRange
// 定义筛选区域第2列
let filterField = 2
// 设置筛选包含内容"张三"
let criteria = '=*张三*'
// 设置筛选开头是"李四"
let criteria2 = '=李四*'
// 应用筛选，指定包含关系
filterRange.AutoFilter(filterField, criteria, xlOr, criteria2)
```

#### [示例七​](#示例七)

将 A 列数值排在前 5 的行筛选显示

js
```js
// 未定义筛选区域时，默认使用UsedRange作为筛选区域
const filterRange = ActiveSheet.UsedRange
// 定义筛选区域A列
let filterField = 1
// 获取筛选的数据区域
const filterColumnRange = filterRange.Columns(filterField)
// 计算数据区域的行数，需要减去表头
const filterRangeCount = filterColumnRange.Rows.Count - 1
// 获取前5名数据
const top5Value = 5
// 数据区域中可能没有5行数据
const topNum = filterRangeCount < top5Value ? filterRangeCount : top5Value
// 获取筛选的数据里第5大的值
let top5 = WorksheetFunction.Large(filterColumnRange, topNum)
// 应用筛选
filterRange.AutoFilter(filterField, '>=' + top5)
```

#### [示例八​](#示例八)

将 C 列为"张三"的 F 列设置为楷体字体

js
```js
// 选择筛选区域为已使用的区域
const filterRange = ActiveSheet.UsedRange
// 选择要筛选的列号为3
const filterColumn = 3
// 选择筛选的值为"张三"
const filterValue = ['张三']
// 应用筛选
filterRange.AutoFilter(filterColumn, filterValue)

// 获取筛选结果
const filteredRange =
  ActiveSheet.AutoFilter.Range.SpecialCells(xlCellTypeVisible)

// 获取筛选结果的F列
const filteredRangeFColumn =
  ActiveSheet.AutoFilter.Range.Columns('F').SpecialCells(xlCellTypeVisible)

// 将筛选的结果的F列设置为楷体字体
filteredRangeFColumn.Font.Name = '楷体'
```

#### [示例九​](#示例九)

筛选仅显示以下符合条件的内容：第 2 列内容为'李四’并且第 4 列是上周

js
```js
// 未定义筛选区域时，默认使用UsedRange作为筛选区域
const filterRange = ActiveSheet.UsedRange
// 定义筛选区域第2列
let filterField = 2
// 设置筛选内容为"李四"
let criteria = '李四'
// 应用筛选
filterRange.AutoFilter(filterField, criteria)
// 定义筛选区域为第4列
filterField = 4
// 设置筛选条件为上周，可选值 xlFilterLastWeek，xlFilterThisWeek，xlFilterNextWeek
criteria = xlFilterLastWeek
// 应用筛选
filterRange.AutoFilter(filterField, criteria, xlFilterDynamic)
```

## [AutoFit()​](#autofit)

更改区域中的列宽或行高以达到最佳匹配

被操作的对象必须是Columns或者Rows，否则，该方法将不会有效果。

一个列宽单位等于“常规”样式中一个字符的宽度

#### [示例​](#示例-49)

js
```js
// A列设置自动列宽
const range = Application.Range('A1')
range.Columns.AutoFit()

// 第一行设置自动行高
const range = Application.Range('A1')
range.Rows.AutoFit()
```

## [Select()​](#select)

选择区域，通过该调用可以修改 Application.Selection 的选区访问

#### [示例​](#示例-50)

js
```js
// 区域对象
let range = Application.Range('A1')
range.Select()
let selection = Application.Selection
console.log(selection.Address())
```

## [TextToColumns()​](#texttocolumns)

将单元格中的文本按指定的分隔符分成多个列

#### [参数​](#参数-16)

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| Destination | Range | null | 否 | 指定要将结果插入到的单元格或单元格区域。如果省略此参数，则结果将覆盖源单元格 |
| DataType | XlTextParsingType | xlDelimited | 否 | 指定将被拆分到多列中的文本的格式 |
| TextQualifier | XlTextQualifier | xlTextQualifierDoubleQuote | 否 | 指定文本分隔符 |
| ConsecutiveDelimiter | Boolean | false | 否 | 指定是否将连续分隔符视为一个分隔符 |
| Tab | Boolean | false | 否 | 指定是否使用制表符作为分隔符 |
| Semicolon | Boolean | false | 否 | 指定是否使用分号作为分隔符 |
| Comma | Boolean | false | 否 | 指定是否使用逗号作为分隔符 |
| Space | Boolean | false | 否 | 指定是否使用空格作为分隔符 |
| Other | Boolean | false | 否 | 指定是否使用OtherChar字符的内容作为分隔符 |
| OtherChar | String | '' | 否 | 指定自定义分隔符，默认值为空字符串 |
| FieldInfo | Array | null | 否 | 包含各个数据列解析信息的数组 |
| DecimalSeparator | String | 系统设置 | 否 | 识别数字时，Excel 使用的小数分隔符 |
| ThousandsSeparator | String | 系统设置 | 否 | 识别数字时，Excel 使用的千位分隔符 |
| TrailingMinusNumbers | String | null | 否 | 以减号字符开始的数字 |

提示

关于FieldInfo字段：此数组包含各个数据列解析信息，具体取决于DataType的值。

分隔数据时，此参数是一个双元素数组，每个双元素数组指定特定列的转换选项。 第一个元素是列号 (从 1 开始的) ，第二个元素是指定应该如何解析列，取值是XlColumnDataType枚举。如果输入数据中特定列不存在给定列说明符，则使用 “常规 ”设置分析该列。如果源数据具有固定宽度的列，则每个双元素数组的第一个元素将列的起始字符位置指定为整数。

#### [返回类型​](#返回类型-7)

undefined

#### [示例一​](#示例一-2)

以 tab 为分隔符分隔 A1:A3 里面的内容。

以分号为分隔符分隔 B 列。

以逗号为分隔符分隔 C1。

以空格为分隔符分离 D1 里面的内容。

js
```js
// 1. 以tab为分隔符拆分A1:A3，Tab表示用tab来分隔
Range('A1:A3').TextToColumns(
  undefined,
  xlDelimited,
  xlTextQualifierNone,
  undefined,
  true,
  undefined,
  undefined,
  undefined,
  undefined,
  undefined,
  undefined
)
// 2. 以分号为分隔符拆分B列，Semicolon表示用分号来分隔
Range('B:B').TextToColumns(
  undefined,
  xlDelimited,
  xlTextQualifierNone,
  undefined,
  undefined,
  true,
  undefined,
  undefined,
  undefined,
  undefined,
  undefined
)
// 3. 以逗号为分隔符拆分C1，Comma表示用逗号来分隔
Range('C1').TextToColumns(
  undefined,
  xlDelimited,
  xlTextQualifierNone,
  undefined,
  undefined,
  undefined,
  true,
  undefined,
  undefined,
  undefined,
  undefined
)
// 4. 以空格为分隔符拆分D1，Space表示用空格来分隔
Range('D1').TextToColumns(
  undefined,
  xlDelimited,
  xlTextQualifierNone,
  undefined,
  undefined,
  undefined,
  undefined,
  true,
  undefined,
  undefined,
  undefined
)
```

#### [示例二​](#示例二-2)

以@为分隔符拆分 A1。

以#为分隔符分隔 A2。

以*为分隔符分隔 A3。

js
```js
// 1. 以@为分隔符分隔A1，@不是默认分隔符所以使用Other
Range('A1').TextToColumns(
  undefined,
  xlDelimited,
  xlTextQualifierNone,
  undefined,
  undefined,
  undefined,
  undefined,
  undefined,
  true,
  '@',
  undefined
)
// 2. 以#为分隔符分隔A2，#不是默认分隔符所以使用Other
Range('A2').TextToColumns(
  undefined,
  xlDelimited,
  xlTextQualifierNone,
  undefined,
  undefined,
  undefined,
  undefined,
  undefined,
  true,
  '#',
  undefined
)
// 3. 以*为分隔符分隔A3，*不是默认@分隔符所以使用Other
Range('A3').TextToColumns(
  undefined,
  xlDelimited,
  xlTextQualifierNone,
  undefined,
  undefined,
  undefined,
  undefined,
  undefined,
  true,
  '*',
  undefined
)
```

#### [示例三​](#示例三-2)

以空格和分号分离 A1:A10 的内容，并拆分到 B1。

以逗号、tab 和=为分隔符号分离 B 列的内容并填充到 D 列。

以 tab、逗号、分号和空格为分隔符号拆分 C 列。

js
```js
// 1. 以空格和分号为分隔符拆分A1:A10，拆分到B1单元格里面
Range('A1:A10').TextToColumns(
  Range('B1'),
  xlDelimited,
  xlTextQualifierNone,
  undefined,
  undefined,
  true,
  undefined,
  true,
  undefined,
  undefined,
  undefined
)
// 2. 以逗号、tab和=为分隔符号拆分B列，拆分到D列里面
Range('B:B').TextToColumns(
  Range('D:D'),
  xlDelimited,
  xlTextQualifierNone,
  undefined,
  true,
  undefined,
  true,
  undefined,
  true,
  '=',
  undefined
)
// 3. 以tab、逗号、分号和空格为分隔符号拆分C列
Range('C:C').TextToColumns(
  undefined,
  xlDelimited,
  xlTextQualifierNone,
  undefined,
  true,
  true,
  true,
  true,
  undefined,
  undefined,
  undefined
)
```

#### [示例四​](#示例四-2)

以逗号对 A1 分列，连续的分隔符要视为单个处理。

以&和空格对 B1:B5 区域进行分列，忽略连续的分隔符。

以!对 C 列拆分数据为多列，连续分隔符单独处理。

js
```js
// 1. 以逗号为分隔符拆分A1，ConsecutiveDelimiter设为true表示连续分隔符视为一个分隔符
Range('A1').TextToColumns(
  undefined,
  xlDelimited,
  xlTextQualifierNone,
  true,
  undefined,
  undefined,
  true,
  undefined,
  undefined,
  undefined,
  undefined
)
// 2. 以&和tab为分隔符拆分B1:B5，&不是默认分隔符所以使用Other，ConsecutiveDelimiter设为true表示连续分隔符视为一个分隔符
Range('B1:B5').TextToColumns(
  undefined,
  xlDelimited,
  xlTextQualifierNone,
  true,
  undefined,
  undefined,
  undefined,
  true,
  true,
  '&',
  undefined
)
// 3. 以!为分隔符拆分C列，!不是默认分隔符所以使用Other，ConsecutiveDelimiter设为false表示连续分隔符单独处理
Range('C:C').TextToColumns(
  undefined,
  xlDelimited,
  xlTextQualifierNone,
  false,
  undefined,
  undefined,
  undefined,
  undefined,
  true,
  '!',
  undefined
)
```

#### [示例五​](#示例五-1)

以空格为分隔符拆分 A1 内容到 B1，拆分的内容第一列设为常规类型，第二列设为文本类型，第三列设为日期 YMD 格式。

以分号为分隔符拆分 C1 内容到 C3，拆分的内容第二列设为日期 DMY 式，第三列设为日期 MYD 格式，第四列设为日期 DYM,格式，第五列设为日期 YDM 格式。

js
```js
// 1. 以空格为分隔符拆分A1，拆分到B1, [X, Y] 代表第X列为Y对应的格式，[1, xlGeneralFormat]代表第一列为常规格式，[2, xlTextFormat]代表第二列为文本格式，[3, xlYMDFormat]代表第三列为日期YMD格式，
Range('A1').TextToColumns(
  Range('B1'),
  xlDelimited,
  xlTextQualifierNone,
  undefined,
  undefined,
  undefined,
  undefined,
  true,
  undefined,
  undefined,
  [
    [1, xlGeneralFormat],
    [2, xlTextFormat],
    [3, xlYMDFormat]
  ]
)
// 2. 以分号为分隔符拆分C1，拆分到C3,[2, xlDMYFormat]代表第二列为日期DMY格式，[3, xlMYDFormat]代表第三列为日期MYD格式，[4, xlDYMFormat]代表第四列为日期DYM格式，[5, xlYDMFormat]代表第四列为日期YDM格式，
Range('C1').TextToColumns(
  Range('C3'),
  xlDelimited,
  xlTextQualifierNone,
  undefined,
  undefined,
  true,
  undefined,
  undefined,
  undefined,
  undefined,
  [
    [2, xlDMYFormat],
    [3, xlMYDFormat],
    [4, xlDYMFormat],
    [5, xlYDMFormat]
  ]
)
```


#### 图形(Shape)

# [Shape​](#shape)

某个工作表上的单个 Shape 图形对象

Shape 对象的具体属性和方法请参阅下方的列表。

### [属性列表​](#属性列表)

| 属性名 | 数据类型 | 简介 |
| --- | --- | --- |
| ID | String | 图形 ID |
| Name | String | 图形名称 |
| Title | String | 图形标题 |
| Chart | Chart | 图表对象 |

### [方法列表​](#方法列表)

| 方法名 | 返回类型 | 简介 |
| --- | --- | --- |
| Delete | undefined | 删除Shape |

## [ID​](#id)

单个图形 ID

### [数据类型​](#数据类型)

String - 图形对象的 ID

### [示例​](#示例)

js
```js
// 图形对象集合
const shape = Application.ActiveSheet.Shapes.Item(1)

// 打印该图形对象的ID
console.log(shape.ID)  //94646582190480
```

## [Name​](#name)

单个图形名称

### [数据类型​](#数据类型-1)

String - 图形对象的名称

### [示例​](#示例-1)

js
```js
// 图形对象集合
const shape = Application.ActiveSheet.Shapes.Item(1)

// 打印图形对象的名称
console.log(shape.Name) //图表 1
```

## [Title​](#title)

单个图形标题

该对象只能设置值，无法读取值

### [数据类型​](#数据类型-2)

String - 图形对象的标题

### [示例​](#示例-2)

js
```js
const shape = Application.ActiveSheet.Shapes.Item(1)

// 设置图形对象的标题,该属性只写,效果需要自行观察图表对象
shape.Title = 'WebOffice'
```

## [Chart​](#chart)

单个图表对象

### [数据类型​](#数据类型-3)

Chart- 单个图表对象

### [示例​](#示例-3)

js
```js
const shape = Application.ActiveSheet.Shapes.Item(1)
shape.Chart.HasTitle = false // 不展示标题
```

## [Delete()​](#delete)

删除图标

### [数据类型​](#数据类型-4)

undefined

### [示例​](#示例-4)

js
```js
const shape = Application.ActiveSheet.Shapes.Item(1)
shape.Delete()
```


#### 图表(Chart)

# [Chart​](#chart)

单个图表对象

Chart对象的具体属性和方法请参阅下方的列表。

### [属性列表​](#属性列表)

| 属性名 | 数据类型 | 简介 |
| --- | --- | --- |
| ChartTitle.Text | String | 图表标题 |
| HasTitle | Boolean | 图表标题是否可见 |

### [方法列表​](#方法列表)

| 方法名 | 返回类型 | 简介 |
| --- | --- | --- |
| SetSourceData() | undefined | 为图表设置源数据区域 |

## [ChartTitle.Text​](#charttitle-text)

设置标题

### [数据类型​](#数据类型)

String - 图表标题

设置图形对象的标题,该属性只写,效果需要自行观察图表对象

### [示例​](#示例)

js
```js
const chart = Application.ActiveSheet.Shapes.Item(1).Chart

// 设置标题
chart.ChartTitle.Text = '图表标题'
```

## [HasTitle​](#hastitle)

如果坐标轴或图表有可见标题，则该属性值为 true

设置图形对象的标题,该属性只写,效果需要自行观察图表对象

### [数据类型​](#数据类型-1)

Boolean - 图表标题是否可见

### [示例​](#示例-1)

js
```js
const chart = Application.ActiveSheet.Shapes.Item(1).Chart

// 设置标题不可见
shape.Chart.HasTitle = false
```

## [SetSourceData()​](#setsourcedata)

为指定图表设置源数据区域

### [参数​](#参数)

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| Source | Range |  | 是 | 包含源数据的区域，可用 Range 对象 |
| PlotBy | Enum |  | 否 | 指定图表类型，对应Enum.XlRowCol，可以为 xlColumns 或 xlRows |

### [示例​](#示例-2)

js
```js
const chart = Application.ActiveSheet.Shapes.Item(1).Chart

// 指定图表数据源
let source = Application.Range('A1:D4')

// 设置图表数据源
chart.SetSourceData(source, Application.Enum.XlRowCol.xlRows)
```


#### 字体(Font)

# [Font​](#font)

单元格内字体的属性，包括加粗，颜色，大小，斜体，删除线和下划线。

Font 对象的具体属性和方法请参阅下方的列表。

### [属性列表​](#属性列表)

| 属性名 | 数据类型 | 简介 |
| --- | --- | --- |
| Bold | Boolean | 字体是否加粗 |
| Color | String | 字体的颜色 |
| Italic | Boolean | 字体是否斜体 |
| Name | String | 字体的名称 |
| Size | Number | 字体的大小 |
| Strikethrough | Boolean | 字体是否有删除线 |
| Underline | Enum.XlUnderlineStyle | 字体是否有下划线 |

## [Bold​](#bold)

获取或者设置是否加粗

### [数据类型​](#数据类型)

Boolean - 字体是否加粗

### [示例​](#示例)

js
```js
// 字体对象
let font = Application.Range('A1').Font

// 打印字体是否加粗
console.log('字体是否加粗：', font.Bold)

// 设置字体加粗
font.Bold = true
```

## [Color​](#color)

获取或者设置字体的颜色

### [数据类型​](#数据类型-1)

String - 字体的颜色，16 进制的 RGB 格式

### [示例​](#示例-1)

js
```js
// 字体对象
let font = Application.Range('A1').Font

// 打印字体颜色
console.log('字体颜色：', font.Color)

// 设置字体颜色
font.Color = '#eb5451'
```

## [Italic​](#italic)

获取或者设置字体是否是斜体

### [数据类型​](#数据类型-2)

Boolean - 字体是否是斜体

### [示例​](#示例-2)

js
```js
// 字体对象
let font = Application.Range('A1').Font

// 打印字体是否斜体
console.log('字体是否为斜体：', font.Italic) // 字体是否为斜体：false

// 设置字体斜体
font.Italic = true
```

## [Name​](#name)

获取或者设置字体的名称

### [数据类型​](#数据类型-3)

String - 字体名

### [示例​](#示例-3)

js
```js
// 设置A1单元格的字体为微软雅黑
Range('A1').Font.Name = '微软雅黑'
```

## [Size​](#size)

设置和获取字体的大小

### [数据类型​](#数据类型-4)

Number - 字体大小

### [示例​](#示例-4)

js
```js
// 字体对象
let font = Application.Range('A1').Font

// 打印字体大小
console.log(font.Size) //12

// 设置字体大小
font.Size = 30
```

## [Strikethrough​](#strikethrough)

获取或者设置字体是否有删除线

### [数据类型​](#数据类型-5)

Boolean - 字体是否有删除线

### [示例​](#示例-5)

js
```js
// 字体对象
let font = Application.Range('A1').Font

// 打印字体是否有删除线
console.log('字体是否有删除线：', font.Strikethrough) // 字体是否有删除线：false

// 设置字体删除线
font.Strikethrough = true
```

## [Underline​](#underline)

获取或者设置字体下划线

### [数据类型​](#数据类型-6)

Enum.XlUnderlineStyle- 字体下划线的类型

例如：设置单下划线即设置为Application.Enum.XlUnderlineStyle.xlUnderlineStyleSingle

### [示例​](#示例-6)

js
```js
// 字体对象
const font = Application.Range('A1').Font

// 打印字体是否设置下划线
console.log('字体是否有下划线：', font.Underline) //字体是否有下划线：-4142

const underlineStatus = Application.Enum.XlUnderlineStyle.xlUnderlineStyleSingle

// 设置字体有单下划线
font.Underline = underlineStatus
```


#### 字段(Field)

# [Field​](#field)

字段操作

### [方法列表​](#方法列表)

| 方法名 | 返回类型 | 简介 |
| --- | --- | --- |
| GetFields() | Array | 获取字段信息 |
| CreateFields() | Array | 创建字段 |
| DeleteFields() | Array | 删除字段 |
| UpdateFields() | Array | 更新字段 |

## [GetFields()​](#getfields)

获取字段信息

### [返回值​](#返回值)

Array - 返回获取的表所有字段信息

| 属性 | 数据类型 | 说明 |
| --- | --- | --- |
| id | String | 字段Id |
| name | String | 字段名称 |
| type | String | 字段类型 |

### [示例​](#示例)

javascript
```javascript
const sheet = Application.ActiveSheet
// 获取的表所有字段信息
const fields = sheet.Field.GetFields()
console.log(fields)
// 打印结果：
// [
//  {"id":"Ce","name":"名称","type":"MultiLineText"},
//  {"id":"Cf","name":"数量","type":"Number"},
// ]
```

## [CreateFields()​](#createfields)

创建字段

### [参数​](#参数)

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| Fields | Array |  | 是 | 表的字段信息,格式说明见附录 |

### [返回值​](#返回值-1)

Array - 返回已创建的表所有字段信息

| 属性 | 数据类型 | 说明 |
| --- | --- | --- |
| id | String | 字段Id |
| name | String | 字段名称 |
| type | String | 字段类型 |

### [示例​](#示例-1)

javascript
```javascript
const sheet = Application.ActiveSheet
const field =  sheet.Field.CreateFields({ 
    Fields: [ 
        { name: '等级',  type: 'Rating', max: 5 }
    ] 
})
console.log(field)
// 打印结果：
// [{"id":"LZ","name":"等级","type":"Rating"}]
```

## [DeleteFields()​](#deletefields)

删除字段

### [参数​](#参数-1)

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| Fields | Array |  | 是 | 需要删除的字段Id |

### [返回值​](#返回值-2)

Array - 返回删除的表id以及删除是否成功信息

| 属性 | 数据类型 | 说明 |
| --- | --- | --- |
| id | String | 字段Id |
| deleted | Boolean | 是否删除成功 |

### [示例​](#示例-2)

javascript
```javascript
const sheet = Application.ActiveSheet
// 删除字段
const resutlt = sheet.Field.DeleteFields({ FieldIds: ['P', 'Q'] })
console.log(resutlt)
// 打印结果：
// [{"deleted":false,"id":"P"},{"deleted":false,"id":"Q"}]
```

## [UpdateFields()​](#updatefields)

更新字段

### [参数​](#参数-2)

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| Fields | Array |  | 是 | 更新的字段信息，包含字段Id，字段name,格式说明见附录 |

### [返回值​](#返回值-3)

Array - 返回已更新的字段信息

| 属性 | 数据类型 | 说明 |
| --- | --- | --- |
| id | String | 字段Id |
| name | String | 字段名称 |
| type | String | 字段类型 |

### [示例​](#示例-3)

javascript
```javascript
const sheet = Application.ActiveSheet
// 修改字段名称
sheet.Field.UpdateFields({ 
    Fields: [{ id: 'LG', name: '跳转' }]
})
```


#### 工作簿(Workbook)

# [Workbook​](#workbook)

工作簿，文件的工作区,包含所有 Sheet

Workbook 对象的具体属性和方法请参阅下方的列表。

#### [属性列表​](#属性列表)

| 属性名 | 数据类型 | 简介 |
| --- | --- | --- |
| ActiveSheet | Sheet | 工作簿中的活动工作表 |
| Sheets | Sheets | 工作表的所有对象集合 |
| ReadOnly | Boolean | 文档是否只读 |
| ReadOnlyComment | Boolean | 文档是否只读可评论的权限 |
| SupportReadOnlyComment | Boolean | 文档是否支持只读可评论权限 |

#### [方法列表​](#方法列表)

| 方法名 | 返回类型 | 简介 |
| --- | --- | --- |
| Save() | String(JSON) | 保存文件的改动 |
| GetComments() | String(JSON) | 获取整个 Workbook 的评论 |
| ExportAsFixedFormat() | String(JSON) | 导出整个表格的 PDF 或者 Img 图片 |

## [ActiveSheet​](#activesheet)

工作簿中的活动工作表

可用Application.ActiveSheet代替

#### [数据类型​](#数据类型)

Sheet- 活动工作簿中的活动工作表

#### [示例​](#示例)

js
```js
// 下面两种写法是一样的
console.log(Application.ActiveWorkbook.ActiveSheet.Name) // Sheet1
console.log(Application.ActiveSheet.Name) // Sheet1
```

## [Sheets​](#sheets)

获取当前文件能操作的所有Sheet，返回一个Sheets对象。

#### [数据类型​](#数据类型-1)

Sheets

#### [示例​](#示例-1)

js
```js
// 下面两种写法是一样的
let sheets = Application.ActiveWorkbook.Sheets
sheets = Application.Sheets
// 打印所有工作表的名称
for (let i = 1; i <= sheets.Count; i++) {
    console.log((sheets.Item(i).Name))
}
```

## [ReadOnly​](#readonly)

返回一个值，表示文档是否只读，此属性为只读属性。

#### [数据类型​](#数据类型-2)

Boolean - 文档是否只读

#### [示例​](#示例-2)

js
```js
// 打印当前文档是否只读
console.log(Application.ActiveWorkbook.ReadOnly) //false
```

## [ReadOnlyComment​](#readonlycomment)

返回一个值，表示文档是否只读可评论的权限，此属性为只读属性。

#### [数据类型​](#数据类型-3)

Boolean - 文档是否只读可评论的权限

#### [示例​](#示例-3)

js
```js
// 打印当前文档是否只读可评论
console.log(Application.ActiveWorkbook.ReadOnlyComment) //true
```

## [Save()​](#save)

保存文件的改动

#### [返回类型​](#返回类型)

String(JSON) - 文件的保存状态，具体格式如下：

| 属性 | 数据类型 | 说明 |
| --- | --- | --- |
| result | String | 保存状态 |
| size | Number | 文件大小，单位 byte |
| version | Number | 版本 |

保存状态说明：

| 保存状态 | 说明 |
| --- | --- |
| ok | 版本保存成功，可在历史版本中查看 |
| nochange | 文档无更新，无需保存版本 |
| SavedEmptyFile | 暂不支持保存空文件 触发场景：内核保存完后文件为空 |
| SpaceFull | 空间已满 |
| QueneFull | 保存中请勿频繁操作 触发场景：服务端处理保存队列已满，正在排队 |
| fail | 保存失败 |

#### [示例​](#示例-4)

js
```js
// 保存文件的改动
let res = Application.ActiveWorkbook.Save()
if (res.result === 'ok') {
  console.log('成功')
} else {
  console.error('失败:', res.result)
}
```

## [GetComments()​](#getcomments)

获取整个 Workbook 的评论

#### [返回类型​](#返回类型-1)

String(JSON) - 整个 workbook 的评论集合，具体格式如下：

| 属性 | 数据类型 | 说明 |
| --- | --- | --- |
| CellComments | Array<Object> | 评论信息集合 |
| PosInfo | String | 单元格信息 |
| SheetName | String | Sheet 名称 |
| UserIds | Array<String> | 用户 id 集合 |

评论信息集合：

| 属性 | 数据类型 | 说明 |
| --- | --- | --- |
| DateTime | String | 时间戳 |
| Text | String | 评论文本 |
| Time | String | 转换后的时间 |
| UserId | String | 用户 id |

#### [示例​](#示例-5)

js
```js
// 获取整个 Workbook 的评论
let json = Application.ActiveWorkbook.GetComments()
for (let i = 0; i < json.length; i++) {
  let pos = json[i].PosInfo
  let comments = json[i].CellComments
  let commentList = []
  for (let j = 0; j < comments.length; j++) {
    commentList.push(comments[j].Text)
  }
  console.log(pos, ':', commentList)
}
```

## [ExportAsFixedFormat()​](#exportasfixedformat)

导出整个表格为对应的 PDF 或者 Img 图片，并获取导出后的 url

#### [参数​](#参数)

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| Type | Enum |  | 可选 | 导出的类型，详细可参考Enum.XlFixedFormatType，目前仅支持导出图片和导出 PDF |

#### [返回类型​](#返回类型-2)

String(JSON) - 导出 URL 的 JSON 字符串

#### [示例 1：导出 PDF​](#示例-1-导出-pdf)

js
```js
// 导出整个表格
let json = Application.ActiveWorkbook.ExportAsFixedFormat()
console.log(json.url)
```

#### [示例 2：导出图片​](#示例-2-导出图片)

js
```js
// 导出整个表格
let json = Application.ActiveWorkbook.ExportAsFixedFormat({
  Type: Application.Enum.XlFixedFormatType.xlTypeIMG
})
console.log(json.url)
```


#### 工作表(Sheet)

# [Sheet​](#sheet)

工作簿（Workbook）中单个工作表(Sheet)对象

Sheet 对象的具体属性和方法请参阅下方的列表。

#### [属性列表​](#属性列表)

| 属性名 | 数据类型 | 简介 |
| --- | --- | --- |
| Id | String | 该工作表的 Id |
| Name | String | 该工作表的名称 |
| Index | Number | 该工作表在所有工作表的索引值 |
| Cells | Range | 该工作表上所有单元格的集合 |
| Columns | Range | 该工作表上所有列的集合 |
| Rows | Range | 该工作表上所有行的集合 |
| UsedRange | Range | 该工作表的使用范围 |
| Visible | Boolean | 该工作表是否可见 |
| Type | String | 该工作表的类型 |
| Hyperlinks | Hyperlinks | 该工作表上所有超链接的集合 |
| Shapes | Shapes | 该工作表上所有图形的集合 |
| Sort | Sort | 该工作表上排序对象 |

#### [方法列表​](#方法列表)

| 方法名 | 返回类型 | 简介 |
| --- | --- | --- |
| Range() | Range | 一个单元格或单元格区域 |
| Cells() | Range | 该工作表上的某个单元格 |
| Activate() | undefined | 切换(激活)工作表 |
| Move() | undefined | 移动工作表 |
| Delete() | undefined | 删除工作表 |

## [Id​](#id)

获取工作表 Id

#### [数据类型​](#数据类型)

String - 工作表 Id

#### [示例​](#示例)

js
```js
const sheet = Application.ActiveSheet
// 打印当前活动工作表的id
console.log(sheet.Id)
```

## [Name​](#name)

设置/获取 工作表名称

#### [数据类型​](#数据类型-1)

String - 该工作表在所有工作表的名称

#### [示例​](#示例-1)

js
```js
const sheet = Application.ActiveSheet
// 打印当前活动工作表的名称
console.log(sheet.Name) // Sheet2

// 将当前工作表的名称改为 WPS WebOffice
sheet.Name = 'WPS WebOffice'
```

## [Index​](#index)

工作表的 index,即该工作表在所有工作表的索引值

#### [数据类型​](#数据类型-2)

String - 该工作表在所有工作表的索引值

#### [示例​](#示例-2)

js
```js
const sheet = Application.ActiveSheet
// 打印当前活动工作表的index
console.log(sheet.Index) // 1
```

## [Cells​](#cells)

工作表上所有单元格的集合，返回一个 Range 对象（可使用 Range 相关的属性、方法）

#### [数据类型​](#数据类型-3)

Range- 工作表上所有单元格的集合

#### [示例​](#示例-3)

js
```js
const sheet = Application.ActiveSheet
// 打印活动工作表的全部单元格地址
console.log(sheet.Cells.Address()) // $A$1:$XFD$1048576
```

## [Columns​](#columns)

工作表上的所有列，返回的是一个 Range 对象，可参考使用Range

#### [数据类型​](#数据类型-4)

Range- 工作表上的所有列

#### [示例​](#示例-4)

js
```js
// 打印该工作表列的数量
const sheet = Application.ActiveSheet
console.log(sheet.Columns.Count) //16384
```

## [Rows​](#rows)

工作表上的行，返回一个 Range 对象（可使用 Range 相关的属性、方法）

#### [数据类型​](#数据类型-5)

Range- 工作表上的所有列

#### [示例​](#示例-5)

js
```js
const sheet = Application.ActiveSheet
console.log(sheet.Rows.Count) // 1048576
```

## [UsedRange​](#usedrange)

工作表激活的区域,即是工作表实际使用到的区域，返回一个 Range 对象（可使用 Range 相关的属性、方法）

#### [数据类型​](#数据类型-6)

Range- 工作表里的一个单元格或单元格区域

#### [示例​](#示例-6)

js
```js
const UsedRange = Application.ActiveSheet.UsedRange
// 打印激活区域的范围,此处假设激活区域为10*10的单元格
console.log(
  UsedRange.Row,
  UsedRange.RowEnd,
  UsedRange.Column,
  UsedRange.ColumnEnd
) // 1,10,1,10
```

## [Visible​](#visible)

显示/隐藏 工作表

#### [数据类型​](#数据类型-7)

Boolean - 工作表是否可见

#### [示例​](#示例-7)

js
```js
const sheet = Application.ActiveSheet
// 隐藏工作表
sheet.Visible = false
// 取消工作表隐藏
sheet.Visible = true
```

## [Type​](#type)

工作表类型

#### [数据类型​](#数据类型-8)

Enum.XlSheetType- 工作表的类型

#### [示例​](#示例-8)

js
```js
const sheet = Application.ActiveSheet
// 打印当前活动工作表的类型
console.log(sheet.Type) //xlWorksheet
```

## [AutoFilter​](#autofilter)

当前工作表的自动筛选对象，该对象内封装了当前工作表的所有筛选对象集合，以及应用和取消筛选条件的方法

#### [数据类型​](#数据类型-9)

AutoFilter- 自动筛选对象

#### [示例​](#示例-9)

js
```js
// 获取自动筛选对象
const autoFilter = Application.ActiveSheet.AutoFilter
// 获取筛选对象集合
const filters = autoFilter.Filters
// 获取当前工作表所有筛选对象的数量
const count = filters.Count
console.log(count)
```

## [Hyperlinks​](#hyperlinks)

工作表上的所有超链接的集合

#### [数据类型​](#数据类型-10)

Hyperlinks集合 - 所有超链接的集合对象

#### [示例​](#示例-10)

js
```js
// 打印超链接集合里超链接对象的个数
const hyperlinks = Application.ActiveSheet.Hyperlinks
console.log(hyperlinks.Count) // 3
```

### [Hyperlinks.Count​](#hyperlinks-count)

集合中超链接的数量

#### [数据类型​](#数据类型-11)

Number - 集合中超链接的数量

#### [示例​](#示例-11)

js
```js
// 打印超链接集合里超链接对象的个数
const hyperlinks = Application.ActiveSheet.Hyperlinks
console.log(hyperlinks.Count) // 3
```

### [Hyperlinks.Item()​](#hyperlinks-item)

获取超链接对象

#### [参数​](#参数)

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| Index | Number |  | 是 | 从 1 开始 |

#### [返回类型​](#返回类型)

Hyperlink- 超链接对象

#### [示例​](#示例-12)

js
```js
const hyperlinks = Application.ActiveSheet.Hyperlinks
if (hyperlinks.Count <= 0) {
  throw new Error('当前文档没有超链接')
}
// 打印超链接的地址和文本
console.log(hyperlinks.Item(1).TextToDisplay) // 超链接的显示内容
console.log(hyperlinks.Item(1).Address) // 点击跳转的超链接
```

## [Shapes​](#shapes)

当前工作表上的所有 Shape 对象的集合

#### [数据类型​](#数据类型-12)

Shapes集合对象 - 该工作表上的所有 Shape 对象的集合

#### [示例​](#示例-13)

js
```js
const sheet = Application.ActiveSheet
const shapes = sheet.Shapes
const chartEnum = Application.Enum.XlChartType.xlColumnClustered //簇状柱形图
// 在图形对象集合中添加 300 * 300 的簇状柱形图
shapes.AddChart2(340, chartEnum, 0, 0, 300, 300)
```

### [Shapes.GetActiveShapeImg()​](#shapes-getactiveshapeimg)

获取激活单元格的图片数据

#### [返回类型​](#返回类型-1)

String - 图片原图下载链接

#### [示例​](#示例-14)

js
```js
// 假如A1有图片
Application.Range('A1').Select()
const imgUrl = Application.ActiveSheet.Shapes.GetActiveShapeImg()
console.log(imgUrl) // https://imageUrl 如果没有图片则返回undefined
```

### [Shapes.Item(Index)​](#shapes-item-index)

代表绘图层中的对象，例如自选图形、任意多边形、OLE 对象或图片

#### [参数​](#参数-1)

| 参数名 | 数据类型 | 默认值 | 可选 | 简介 |
| --- | --- | --- | --- | --- |
| index | Number |  | 否 | 对象的索引 |

#### [返回类型​](#返回类型-2)

Shape- 绘图层的一个对象,自选图形、任意多边形、OLE 对象或图片

#### [示例​](#示例-15)

js
```js
const shapes = Application.ActiveSheet.Shapes
// 打印第一个图形对象的ID
console.log(shapes.Item(1).ID)
```

### [Shapes.Count​](#shapes-count)

图形的数量

#### [数据类型​](#数据类型-13)

Nunber - 工作表的图形对象的数量

#### [示例​](#示例-16)

js
```js
const sheet = Application.ActiveSheet
const shapes = sheet.Shapes
// 打印当前图形的数量
console.log(shapes.Count)
```

### [Shapes.AddChart2()​](#shapes-addchart2)

添加图形

#### [参数​](#参数-2)

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| Style | String |  | 否 | 指定图表样式 |
| XlChartType | Enum |  | 否 | 指定图表类型，对应Enum.XlChartType |
| Left | Number |  | 否 | 指定新建图表的左边距，单位 px |
| Top | Number |  | 否 | 指定新建图表的上边距，单位 px |
| Width | Number |  | 否 | 指定新建图表的宽度，单位 px |
| Height | Number |  | 否 | 指定新建图表的高度，单位 px |

#### [返回类型​](#返回类型-3)

Shape- 添加的图形对象

#### [示例​](#示例-17)

js
```js
const chartEnum = Application.Enum.XlChartType.xlColumnClustered //簇状柱形图的枚举
const sheet = Application.ActiveSheet
const shapes = sheet.Shapes
// 在图形对象集合中添加 300 * 300 的簇状柱形图
shapes.AddChart2(340, chartEnum, 0, 0, 300, 300)
```

### [Shapes.AddPicture()​](#shapes-addpicture)

向表格中插入浮动图片

#### [参数​](#参数-3)

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| FileName | String |  | 是 | 要插入的图片（可以是图片数据的 Base64 字符串或者 URL 图片地址） |
| LinkToFile | Enum |  | 是 | 暂不支持，请填0 |
| SaveWithDocument | Enum |  | 是 | 暂不支持，请填0 |
| Left | Number |  | 是 | 图片左边缘相对于表格左边缘的位置 |
| Top | Number |  | 是 | 图片上边缘相对于表格片上边缘的位置 |
| Width | Number |  | 否 | 图片的宽度 |
| Height | Number |  | 否 | 图片的高度 |

#### [示例​](#示例-18)

js
```js
// 获取图形对象
const shapes = ActiveSheet.Shapes
// 插入浮动图片
shapes.AddPicture({
  FileName:
    'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAMAAABEpIrGAAACH1BMVEUAAAA9kP8mpv9Fv/8QT94TXeQAfv8trP0xr/0VUdoAcv8LgPs1sf4AgP4QUuURTuEVl/ksq/wAg/84tf4NUt80sf0bnPsPUuAAYfsAgP4UTtYAY/8GXPQRkv1Jxv8AZP86tv9Bvv9Gv/9Atv8KiPUOVesAZ+wbS8wUl/kjo/sAYv8PVekjovxBu/8AY/8trPxCu/8AW/AAcvAxsP0npvsYTdElo/oQT9o+uP8AYv8KVOIHdugAYfwAfPwbfe49uf4XmfkcdOopo/oAYv86tf4AYf8TT9cZaeYVTdUJVOUcnvsAYv8Agv9Iwv8AYv9Jwf8bSs0gofoAY/8Ag/9Gv/8en/oAYf9Fv/8AWuQZS80AdO5BvP8AYv9DvP8Umfo9uv8LVOUAgf8amfk/uf8AY/8Ag/8Ahf8TUdgVT9UNWOknov8trPwxr/wmpvspqfs1sv0io/s5tf0AgP8bnfoMV+4AYPoAYv8AYf4AYfwAe/sen/oAd/gAXvYAff0LW/MLWfEOUN0PT9oSTtZAu/49uP0qqfweoPoXmvkAdPUAXPIAcfEKUuIWTNIXS88KXPYAWusOVeoPVOkPUuYQUeQQT+IMUeAUTdMZSs09uP4Aa+kAWOgAaeQIVOQIU+QAV+MRTd8STd0RTtgSl/kAW+8NWO8Abu0AbuwBZNsAg/8AYPwAc/MIVecAZuEAVeAAWt8AZN0JgvAIfOsJV+sTVuIAU9qil9AQAAAAa3RSTlMAAwb++A3+bkkkGxL8+Pf36tnTubCYjol4dmVjRjEvLysZFA7+/Pv6+fj39/b19fPw8O/u7Ovq6Ofn5+Xj4uDa2trZ1dPS0s/Ny7+9vLi4trapoqGgoJ2XkI+KgoJycGRhX1BNS0pHQj06IcCB3jkAAAIBSURBVDjLdZCHctpAFEWfCMU1ce8tTu+99957770nJFGwkE1sgyMTAginYbkAIQSD494+0G+ltTAacWZnNLt7dO/Og3Tqzl2og8wsvroAudEI+iy5s+6LzMZ7Bp1r5sFOQegRenAJwoEnoCXr6Mc0ihvSrhuK32vJK21Mva00760OWyz0KZYNbzJQmKXUF1o7rZ24rGngweETIGOwbH2nw/ZDUaM6grJVHzSs3R+NoqDy+lKLuwWX2618r03/QRalhNzbewo+qZyuZX4TVMFQubm3f3xy5VeZg49xsL8IC+mUH+1zeLyJ3pHkRCwWK7hvIGc/CYrwvIhtbXZ4hgYwJLm0LFf56QdhDRAq8wdZMypD3kT/SBJo6nfCMmWTc2WQlUO8GDIn2GSAUnuWGERJpIS4LR4HlWdFNITua9pl5s3h5vJ81tzscAAhuyTc3kZQ51CxqasvMmDGENyZyo1BX2iqLRAI0L7qvU2fv6Ex5mVbgana4XIGfb7QsI0KNcc4O4/Gv67I6Jjn5cmOvy4nGuFwaFipqFjN2YlBQkZn/N0dxAhiSfjUKzqHyxwq/0lIX0T0d6OAGb7dD0Gl/gzH8UqN6JcNp7HcBPN5esTO8U2oSCIpcZVkgwamepedxxpJROP4C9DBdHc9vlWSxG1VDOiTc53jpRW3TJCZ+vMXNeWzDz4DoNZyqecAAAAASUVORK5CYII=',
  LinkToFile: 0,
  SaveWithDocument: 0,
  Left: 30, // 图片距离左边位置
  Top: 30, // 图片距离顶部位置
  Width: 32, // 图片宽度
  Height: 32 // 图片高度
})
```

## [Sort​](#sort)

排序对象，设置好目标区域后即可进行排序操作

#### [数据类型​](#数据类型-14)

Sort- 排序对象

#### [示例​](#示例-19)

js
```js
// 对 C 列降序排序并且对 D 列做升序排序

//获取当前表格区域
const range = ActiveSheet.UsedRange
//获取到排序对象
const sort = ActiveSheet.Sort
//获取排序范围
const sortFields = sort.SortFields
//清除之前的范围
sortFields.Clear()
//基于C列降序排序, xlSortOnValues代表按值排序, xlDescending代表降序排序
sortFields.Add(Range('C:C').Item(1, 1), xlSortOnValues, xlDescending)
//基于D列升序排序, xlSortOnValues代表按值排序, xlAscending代表升序排序
sortFields.Add(Range('D:D').Item(1, 1), xlSortOnValues, xlAscending)
//设置是否包含表头参数，xlGuess为自动，xlYes为包含表头，xlNo为不包含表头。默认设置为xlGuess
sort.Header = xlGuess
//设置是否大小写敏感，true为区分大小写，false为不区分大小写，默认设置false
sort.MatchCase = false
//设置中文排序方法，xlPinYin为拼音排序，xlStroke为比划数排序。默认设置为xlPinYin
sort.SortMethod = xlPinYin
//设置排序的方法，xlSortColumns为按列排序，xlSortRows为按行排序，默认设置为xlSortColumns
sort.Orientation = xlSortColumns
//排序前必须设置SetRange
sort.SetRange(range)
//开始排序
sort.Apply()
```

## [Range()​](#range)

一个单元格或单元格区域，返回一个 Range 对象（可使用 Range 相关的属性、方法）

#### [数据类型​](#数据类型-15)

Range- 工作表里的一个单元格或单元格区域

#### [示例​](#示例-20)

js
```js
const sheet = Application.ActiveSheet
// 打印当前活动工作表D2单元格的内容
console.log(sheet.Range('D2').Text) // D2

// 修改当前活动工作表D2单元格的内容
sheet.Range('D2').Value = 'this is D2'
```

## [Cells()​](#cells-1)

选择工作表上的某个单元格，返回一个 Range 对象（可使用 Range 相关的属性、方法）

#### [返回类型​](#返回类型-4)

Range- 工作表上的某个单元格

#### [示例​](#示例-21)

js
```js
// 打印活动工作表的第一个单元格地址
console.log(Application.Cells(1).Address()) // $A$1
console.log(Application.Cells(2).Address()) // $A$2
```

## [Activate()​](#activate)

激活工作表

#### [示例​](#示例-22)

js
```js
const sheet = Application.Sheets.Item(1)
// 激活第一个工作表，此时的Application.ActiveSheet.Range和此时的Application.Range都指向第一个工作簿
sheet.Activate()
// 修改了第一个工作表的A1单元格
Application.Range('A1').Value = 'foo'
```

## [Move()​](#move)

移动工作表

#### [参数​](#参数-4)

两个参数互斥

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| Before | number | null | 否 | 验将放置移动的工作表之前的工作表 ID。如果指定 After ，则不能指定 Before。 |
| After | number | null | 否 | 将放置移动的工作表后的工作表 ID。如果指定 Before ，则不能指定 After |

#### [示例​](#示例-23)

js
```js
// 将当前工作表移动到第二个工作表之后
const sheet = Application.ActiveSheet
sheet.Move({
  Before: null,
  After: Application.Sheets(2).Id
})
```

## [Delete()​](#delete)

删除工作表

#### [返回值​](#返回值)

undefined

#### [示例​](#示例-24)

js
```js
// 删除名称为“Sheet2”的工作表
Application.Sheets.Item('Sheet2').Delete()
```


#### 工作表函数(WorksheetFunction)

# [WorksheetFunction​](#worksheetfunction)

工作表函数对象是 Excel 中的一个内置对象，它包含了许多常用的 Excel 函数，例如 Sum、Average、Min、Max 等。使用 WorksheetFunction 对象可以在 VBA 中调用这些 Excel 函数，以实现对工作表数据的处理和分析。

通过 WorksheetFunction 对象，可以调用 Excel 函数并将其结果赋值给变量，也可以直接在代码中使用这些函数来进行数值计算、字符串处理等操作，可以实现复杂的数据处理和分析需求。

例如，可以使用 WorksheetFunction 对象的 SUM 函数来计算一列数字的总和，也可以使用 Average 函数来计算这列数字的平均值。使用 WorksheetFunction 对象的 Max 函数可以找到一列数字中的最大值，Min 函数可以找到一列数字中的最小值。

#### [方法列表​](#方法列表)

| 方法名 | 返回类型 | 简介 |
| --- | --- | --- |
| Average() | Number | 用于计算指定区域内数字的平均值 |
| AverageIf() | Number | 用于计算指定区域内满足给定条件的所有单元格的平均值 |
| Small() | Number | 返回数据集中第 k 个最小值 |
| Large() | Number | 用于在一个数组或一列数据中返回第 k 个最大值 |
| Min() | Number | 用于在一个数组或一列数据中返回最小值 |
| Max() | Number | 用于在一个数组或一列数据中返回最大值 |
| Sum() | Number | 对某单元格区域中的所有数字求和 |

## [Average()​](#average)

用于计算指定区域内数字的平均值

### [参数​](#参数)

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| range | Range | null | 是 | 指定要进行计算的区域 |

### [返回值​](#返回值)

Number - 函数返回指定区域内数字的平均值

js
```js
// 假设有一个数据表，其中A列为学生姓名，B列为学生成绩。现在需要计算B列的平均分数，可以使用以下代码：

// 取出B列
const range = Range('B:B')
// 获取平均值
const average = WorksheetFunction.Average(range)
console.log(average)
```

## [AverageIf()​](#averageif)

用于计算指定区域内满足给定条件的所有单元格的平均值

### [参数​](#参数-1)

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| range | Range | null | 是 | 要求其平均值的一个或多个单元格 |
| criteria | String | null | 是 | 定义将对哪些单元格求平均值的条件，其形式可以为数字、表达式、单元格引用或文本。 例如，条件可以表示为 32、“32”、“>32”、“apples”或 B4 |
| average_range | Range | null | 否 | 要求其平均值的实际单元格集合，如果省略，则使用 range |

### [返回值​](#返回值-1)

Number - 函数返回指定区域内满足条件的数字的平均值

js
```js
// 假设有一个数据表，其中A列为学生姓名，B列为学生成绩。现在需要计算B列中A列为“张三”的平均分数，可以使用以下代码：

// 获取A列
const rangeA = Range('A:A')
// 获取B列
const rangeB = Range('B:B')
// 获取平均值
const average = WorksheetFunction.AverageIf(rangeA, '张三', rangeB)
console.log(average)
```

## [Small()​](#small)

返回数据集中第 k 个最小值。 使用此函数可以返回数据集中特定位置上的数值。

如果数组为空， Small() 将返回#NUM！。

如果 k ≤ 0 或 k 超过数据点数， Small 将返回#NUM！。

如果 n 为数组中数据点的个数，则 SMALL(array,1) 等于最小值，SMALL(array,n) 等于最大值。

### [参数​](#参数-2)

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| range | Range | null | 是 | 要从中获取第 k 个最小值的数组或数据列 |
| k | Number | null | 是 | 要返回的第 k 个最小值的位置。k 必须大于 0，小于等于数组或数据列中的元素个数 |

### [返回值​](#返回值-2)

Number - 函数返回第 k 个最小值

### [示例​](#示例)

js
```js
// 假设有一个数据表，其中A列为学生姓名，B列为学生成绩。现在需要从B列中获取第3个最小值，可以使用以下代码：

// 取出B列
const range = Range('B:B')
// 获取第三小的值
const third = WorksheetFunction.Small(range, 3)
console.log(third)
```

## [Large()​](#large)

用于在一个数组或一列数据中返回第 k 个最大值。例如，可以使用 Large 返回最高、亚军或第三名的分数。

如果数组为空， 则 Large 返回#NUM！。

如果 k ≤ 0 或 k 大于数据点数， 则 Large 返回#NUM！。

如果区域中数据点的个数为 n，则函数 LARGE(array,1) 返回最大值，函数 LARGE(array,n) 返回最小值。

### [参数​](#参数-3)

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| range | Range | null | 是 | 要从中获取第 k 个最大值的数组或数据列 |
| k | Number | null | 是 | 要返回的第 k 个最大值的位置。k 必须大于 0，小于等于数组或数据列中的元素个数 |

### [返回值​](#返回值-3)

Number - 函数返回第 k 个最大值

### [示例​](#示例-1)

js
```js
// 假设有一个数据表，其中A列为学生姓名，B列为学生成绩。现在需要从B列中获取第4个最大值，可以使用以下代码：

// 取出B列
const range = Range('B:B')
// 获取第四大的值
const fourth = WorksheetFunction.Large(range, 4)
console.log(fourth)
```

## [Min()​](#min)

用于在一个数组或一列数据中返回最小值

参数可以是数字，也可以是包含数字的名称、数组或引用。

直接键入参数列表的数字的逻辑值和文本表示也包括在内。

如果参数为数组或引用，则只使用其中的数值。 数组或引用中的空白单元格、逻辑值或文本将被忽略。

如果参数不包含数字， 则 Min 返回 0。

如果参数为错误值或不能转换为数字的文本，则将导致错误。

### [参数​](#参数-4)

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| range | Range | null | 是 | 要从中查找最小值的 n 个数字 |

### [返回值​](#返回值-4)

Number - 函数返回数组或数据列中的最小值

### [示例​](#示例-2)

js
```js
// 假设有一个数据表，其中A列为学生姓名，B列为学生成绩。现在需要从B列中获取最小值，可以使用以下代码：

// 取出B列
const range = Range('B:B')
// 获取最小值
const min = WorksheetFunction.Min(range)
console.log(min)
```

## [Max()​](#max)

用于在一个数组或一列数据中返回最大值

参数可以是数字，也可以是包含数字的名称、数组或引用。

直接键入参数列表的数字的逻辑值和文本表示也包括在内。

如果参数为数组或引用，则只使用其中的数值。 数组或引用中的空白单元格、逻辑值或文本将被忽略。

如果参数不包含数字， 则 Max 返回 0 (零) 。

如果参数为错误值或不能转换为数字的文本，则将导致错误。

### [参数​](#参数-5)

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| range | Range | null | 是 | 要从中获取最大值的数组或数据列 |

### [返回值​](#返回值-5)

Number - 最大值

### [示例​](#示例-3)

js
```js
// 假设有一个数据表，其中A列为学生姓名，B列为学生成绩。现在需要从B列中获取最大值，可以使用以下代码：

// 取出B列
const range = Range('B:B')
// 获取最大值
const max = WorksheetFunction.Max(range)
console.log(max)
```

## [Sum()​](#sum)

对某单元格区域中的所有数字求和

直接键入参数列表的数字、逻辑值和数字的文本表示也包括在内。

如果参数为数组或引用，则只有该数组或引用中的数字将被计算在内。 数组或引用中的空单元格、逻辑值或文本将被忽略。

如果参数为错误值或不能转换为数字的文本，则将导致错误。

### [参数​](#参数-6)

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| Array | Range | null | 是 | 要对其求和的 n 个参数 |

### [返回值​](#返回值-6)

Number - 函数返回数组或数据列中所有数值的总和

### [示例​](#示例-4)

js
```js
// 假设有一个数据表，其中A列为学生姓名，B列为学生成绩。现在需要获取B列学生成绩的和，可以使用以下代码：

// 取出B列
const range = Range('B:B')
// 获取成绩的和
const sum = WorksheetFunction.Sum(range)
console.log(sum)
```

## [Match()​](#match)

用于在一个区域中查找某个值，并返回该值在区域中的位置。

如果需要项在某个范围中的位置而不是项本身，请使用 Match 而不是 Lookup(Object, Object, Object) 函数。

Lookup_value 为需要在 Look_array 中查找的数值。 例如，如果要在电话簿中查找某人的电话号码，则应该将姓名作为查找值，但实际上需要的是电话号码。

Lookup_value 可以为值（数字、文本或逻辑值）或对数字、文本或逻辑值的单元格引用。

如果 match_type 为 1， Match 将查找小于或等于 lookup_value 的最大值。 Lookup_array 必须按升序排列：...-2、-1、0、1、2、...、A-Z、 false、 true。

如果 match_type 为 0， Match 将查找与 lookup_value 完全相等的第一个值。 Lookup_array 可以按任何顺序排列。

如果 match_type 为 -1， Match 将查找大于或等于 lookup_value 的最小值。 Lookup_array 必须按降序排列： true、 false、Z-A、...2、1、0、-1、-2、...等。如果省略 match_type，则假定为 1。

Match 返回匹配值在 lookup_array 中的位置，而不是值本身。 例如，MATCH("b",{"a","b","c"},0) 返回 2，即“b”在数组 {"a","b","c"} 中的相应位置。

匹配 文本值时，Match 不区分大写字母和小写字母。

如果 Match 未能找到匹配项，则返回#N/A 错误值。

如果 match_type 为 0 且 lookup_value 为文本，则可以在 lookup_value 中使用通配符、问号 (?) 和星号 (*)。 问号匹配任意单个字符；星号匹配任意字符序列。 如果要查找实际的问号或星号，则请在该字符前键入一个波形符 (~)。

### [参数​](#参数-7)

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| Lookup_value | Object | null | 是 | 需要在表中查找的值 |
| Lookup_array | Range | null | 是 | 可能包含所要查找的值的连续单元格区域。 Lookup_array 必须为数组或数组引用 |
| Match_type | Object | null | 否 | 数字 -1、0 或 1。 Match_type 指明 Microsoft Excel 如何将 lookup_value 与 lookup_array 中的值进行匹配 |

### [返回值​](#返回值-7)

Number - 函数返回 lookup_value 在 lookup_array 中的位置

### [示例​](#示例-5)

js
```js
// 假设有一个数据表，其中A列为学生姓名，B列为学生成绩。现在需要查找名为“张三”的学生成绩

// 取出A列
const range = Range('A:A')
// 查找名为“张三”的学生成绩的索引
const rowIndex = WorksheetFunction.Match('张三', range, 0)
// 获取“张三”的成绩
const score = Range(`B${rowIndex}`).Value
console.log(score)
```


#### 排序(Sort)

# [Sort​](#sort)

排序对象，设置好目标区域后即可进行排序操作。

#### [属性列表​](#属性列表)

| 属性名 | 数据类型 | 简介 |
| --- | --- | --- |
| Header | XlYesNoGuess | 指定第一行是否包含标题信息 |
| MatchCase | Boolean | 是否区分大小写 |
| Orientation | XlSortOrientation | 指定排序方向 |
| Rng | Range | 返回要执行排序的值的区域 |
| SortFields | SortFields | 该对象代表与Sort对象关联的排序字段的集合 |
| SortMethod | XlSortMethod | 中文排序方法 |

#### [方法列表​](#方法列表)

| 方法名 | 返回类型 | 简介 |
| --- | --- | --- |
| Apply() | undefined | 根据当前应用的排序状态对区域进行排序 |
| SetRange() | undefined | 设置排序发生的范围 |

#### [应用示例​](#应用示例)

以下为您展示排序对象的在工作表内的一些常见应用场景：

示例 1. 对 C 列降序排序并且对 D 列做升序排序

js
```js
//获取当前表格区域
const range = ActiveSheet.UsedRange
//获取到排序对象
const sort = ActiveSheet.Sort
//获取排序范围
const sortFields = sort.SortFields
//清除之前的范围
sortFields.Clear()
//基于C列降序排序, xlSortOnValues代表按值排序, xlDescending代表降序排序
sortFields.Add(Range('C:C').Item(1, 1), xlSortOnValues, xlDescending)
//基于D列升序排序, xlSortOnValues代表按值排序, xlAscending代表升序排序
sortFields.Add(Range('D:D').Item(1, 1), xlSortOnValues, xlAscending)
//设置是否包含表头参数，xlGuess为自动，xlYes为包含表头，xlNo为不包含表头。默认设置为xlGuess
sort.Header = xlGuess
//设置是否大小写敏感，true为区分大小写，false为不区分大小写，默认设置false
sort.MatchCase = false
//设置中文排序方法，xlPinYin为拼音排序，xlStroke为比划数排序。默认设置为xlPinYin
sort.SortMethod = xlPinYin
//设置排序的方法，xlSortColumns为按列排序，xlSortRows为按行排序，默认设置为xlSortColumns
sort.Orientation = xlSortColumns
//排序前必须设置SetRange
sort.SetRange(range)
//开始排序
sort.Apply()
```

示例 2. 将 G 列单元格颜色为红色的设为顶部，F 列红色的放在末尾

js
```js
//获取当前表格区域
const range = ActiveSheet.UsedRange
//获取到排序对象
const sort = ActiveSheet.Sort
//获取排序范围
const sortFields = sort.SortFields
//清除之前的范围
sortFields.Clear()
//基于G列把红色放顶部
//增加排序范围。第1个参数为基于排序的单元格区域。第2个参数用来指定排序依据，xlSortOnValues为基于单元格值排序，xlSortOnFontColor为基于字体颜色排序，xlSortOnCellColor基于单元格颜色排序。第3个参数为排序方式，xlAscending为升序，xlDescending为降序
const sortField1 = sortFields.Add(
  Range('G:G').Item(1, 1),
  xlSortOnCellColor,
  xlAscending
)
//设置排序颜色为红色
sortField1.SortOnValue.Color = RGB(255, 0, 0)
//基于F列把红色放末尾
//增加排序范围。第1个参数为基于排序的单元格区域。第2个参数用来指定排序依据，xlSortOnValues为基于单元格值排序，xlSortOnFontColor为基于字体颜色排序，xlSortOnCellColor基于单元格颜色排序。第3个参数为排序方式，xlAscending为升序，xlDescending为降序
const sortField2 = sortFields.Add(
  Range('F:F').Item(1, 1),
  xlSortOnCellColor,
  xlDescending
)
//设置排序颜色为红色
sortField1.SortOnValue.Color = RGB(255, 0, 0)
//设置是否包含表头参数，xlGuess为自动，xlYes为包含表头，xlNo为不包含表头。默认设置为xlGuess
sort.Header = xlGuess
//设置是否大小写敏感，true为区分大小写，flase为不区分大小写，默认设置false
sort.MatchCase = false
//设置中文排序方法，xlPinYin为拼音排序，xlStroke为比划数排序。默认设置为xlPinYin
sort.SortMethod = xlPinYin
//设置排序的方法，xlSortColumns为对列排序，xlSortRows为对行排序，默认设置为xlSortColumns
sort.Orientation = xlSortColumns
//排序前必须设置SetRange
sort.SetRange(range)
//开始排序
sort.Apply()
```

示例 3. 工作表包含 A 列到 H 列，把 B 列中较小的排到前面

js
```js
// 获取sort对象
const sort = ActiveSheet.Sort

const keyColumn = 'B'
const keyBeginRow = ActiveSheet.UsedRange.Row
const keyEndRow =
  ActiveSheet.UsedRange.Row + ActiveSheet.UsedRange.Rows.Count - 1

const sortFields = sort.SortFields
const sortKeyRange = ActiveSheet.Range(
  keyColumn + keyBeginRow + ':' + keyColumn + keyEndRow
)
sortFields.Clear()
//xlSortOnValues代表按值排序, xlAscending代表升序排序
sortFields.Add(sortKeyRange, xlSortOnValues, xlAscending)

//选中排序区域
const applyBeginRow = ActiveSheet.UsedRange.Row
const applyEndRow =
  ActiveSheet.UsedRange.Row + ActiveSheet.UsedRange.Rows.Count - 1
const applyRange = ActiveSheet.Range(`A${applyBeginRow}:H${applyEndRow}`)
// 排序前必须设置范围
sort.SetRange(applyRange)
//应用排序
sort.Apply()
```

示例 4. 工作表包含 A 列到 H 列，按笔画从大到小排列 A 列

js
```js
//获取 sort 对象
const sort = ActiveSheet.Sort

//选中关键区域
const keyColumn = 'A'
const keyBeginRow = ActiveSheet.UsedRange.Row
const keyEndRow =
  ActiveSheet.UsedRange.Row + ActiveSheet.UsedRange.Rows.Count - 1
const keyRangeStr = keyColumn + keyBeginRow + ':' + keyColumn + keyEndRow

//进行数据降序排序
const sortFields = sort.SortFields
sortFields.Clear()
sortFields.Add(
  ActiveSheet.Range(keyRangeStr),
  xlSortOnValues,
  xlDescending,
  '',
  undefined
)
sort.header = xlYes
sort.Orientation = xlSortColumns
//设置为按笔画排列
sort.SortMethod = xlStroke

//设置排序区域
const beginRow = ActiveSheet.UsedRange.Row
const endRow = ActiveSheet.UsedRange.Row + ActiveSheet.UsedRange.Rows.Count - 1

// 排序前必须设置范围
sort.SetRange(ActiveSheet.Range(`A${beginRow}:H${endRow}`))
//应用排序
sort.Apply()
```

示例 5. 工作表包含 A 列到 H 列，按 B 列单元格颜色排列，无色的在顶端

js
```js
//选中关键区域
const keyColumn = 'B'
const keyBeginRow = ActiveSheet.UsedRange.Row
const keyEndRow =
  ActiveSheet.UsedRange.Row + ActiveSheet.UsedRange.Rows.Count - 1
const keyRangeStr = keyColumn + keyBeginRow + ':' + keyColumn + keyEndRow
let keyRange = ActiveSheet.Range(keyRangeStr)

//单元格颜色进行升序排列
let sort = ActiveSheet.Sort
let sortFields = sort.SortFields
sortFields.Clear()
sortFields.Add(keyRange, xlSortOnCellColor, xlDescending, '', undefined)
sort.header = xlYes
sort.Orientation = xlSortColumns
sort.SortMethod = xlPinYin

//设置排序区域
const beginRow = ActiveSheet.UsedRange.Row
const endRow = ActiveSheet.UsedRange.Row + ActiveSheet.UsedRange.Rows.Count - 1

// 排序前必须设置范围
sort.SetRange(ActiveSheet.Range(`A${beginRow}:H${endRow}`))
//应用排序
sort.Apply()
```

## [Header​](#header)

指定第一行是否包含标题信息，可读/写XlYesNoGuess。

默认值为 xlNo，如果希望 Excel 确定标题，可以指定 xlGuess。

#### [数据类型​](#数据类型)

XlYesNoGuess- 第一行是否包含标题

## [MatchCase​](#matchcase)

设置为true以执行区分大小写的排序，或设置为false以执行不区分大小写的排序。可读/写。

#### [数据类型​](#数据类型-1)

Boolean - 是否区分大小写

## [Orientation​](#orientation)

指定排序方向，可读/写XlSortOrientation。

#### [数据类型​](#数据类型-2)

XlSortOrientation- 排序方向

## [Rng​](#rng)

返回要执行排序的值的区域，此为只读属性。

#### [数据类型​](#数据类型-3)

Range- 排序区域对象

## [SortFields​](#sortfields)

代表与Sort对象关联的排序字段的集合，此为只读属性。

### [Count​](#count)

返回集合中对象的数目，只读。

#### [数据类型​](#数据类型-4)

Number - 对象数目

### [Add()​](#add)

创建新的排序字段，并返回一个 SortField 对象

#### [参数​](#参数)

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| Key | Range | null | 是 | 指定用于排序的键值 |
| SortOn | XlSortOn | null | 否 | 要进行排序的字段 |
| Order | XlSortOrder | null | 否 | 指定排序次序 |
| CustomOrder | Variant | null | 否 | 指定是否应使用自定义排序次序 |
| DataOption | XlSortDataOption | null | 否 | 指定数据选项 |

#### [返回值​](#返回值)

SortField- 代表与Sort对象关联的排序字段对象

### [Clear()​](#clear)

清除所有 SortFields 对象，在开始添加排序字段时，最好先调用一次此方法

#### [返回值​](#返回值-1)

undefined

### [Item()​](#item)

返回一个SortField对象，此为只读属性

#### [参数​](#参数-1)

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| Index | Number / String | null | 是 | 指定用于索引，默认从 1 开始 |

#### [返回值​](#返回值-2)

SortField- 代表与Sort对象关联的排序字段对象

## [SortMethod​](#sortmethod)

指定中文排序方法，可读/写XlSortMethod。

#### [数据类型​](#数据类型-5)

XlSortMethod- 中文排序方法

## [Apply()​](#apply)

根据当前应用的排序状态对区域进行排序

注意

应用排序规则前必须先进行SetRange，否则排序不会生效。

#### [返回值​](#返回值-3)

undefined

## [SetRange()​](#setrange)

设置排序发生的范围

#### [参数​](#参数-2)

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| Range | Range | null | 是 | 指定 Sort 对象所表示的排序所依据的范围 |

#### [返回值​](#返回值-4)

undefined


#### 排序字段(SortField)

# [SortField​](#sortfield)

代表与Sort对象关联的排序字段对象

#### [属性列表​](#属性列表)

| 属性名 | 数据类型 | 简介 |
| --- | --- | --- |
| CustomOrder | Variant | 指定对字段进行排序的自定义次序 |
| DataOption | XlSortDataOption | 指定如何在 SortField 对象中指定的区域中对文本进行排序 |
| Key | Range | 指定排序字段，该字段确定要排序的值 |
| Order | XlSortOrder | 确定关键字所指定的值的排序次序 |
| Priority | Number | 指定排序字段的优先级 |
| SortOn | XlSortOn | 返回或设置要排序的单元格的属性 |
| SortOnValue | Object | 返回针对指定的 SortField 对象执行排序的值 |

#### [方法列表​](#方法列表)

| 方法名 | 返回类型 | 简介 |
| --- | --- | --- |
| Delete() | undefined | 从 SortFields 集合中删除指定的 SortField 对象 |
| ModifyKey() | undefined | 修改字段中按其排序的键值 |

## [CustomOrder​](#customorder)

指定对字段进行排序的自定义次序，可读/写

#### [数据类型​](#数据类型)

Variant - 自定义次序对象

## [DataOption​](#dataoption)

指定如何在 SortField 对象中指定的区域中对文本进行排序，可读/写

#### [数据类型​](#数据类型-1)

XlSortDataOption- 文本排序方式枚举

## [Key​](#key)

指定排序字段，该字段确定要排序的值，此为只读属性

#### [数据类型​](#数据类型-2)

Range- 区域对象

## [Order​](#order)

确定关键字所指定的值的排序次序，可读/写

#### [数据类型​](#数据类型-3)

XlSortOrder- 排序次序枚举

## [Priority​](#priority)

指定排序字段的优先级，可读/写

#### [数据类型​](#数据类型-4)

Number

## [SortOn​](#sorton)

返回或设置要排序的单元格的属性，可读/写

#### [数据类型​](#数据类型-5)

XlSortOn

## [SortOnValue​](#sortonvalue)

返回针对指定的 SortField 对象执行排序的值，此为只读属性

#### [数据类型​](#数据类型-6)

Object

## [Delete()​](#delete)

从 SortFields 集合中删除指定的 SortField 对象

#### [返回值​](#返回值)

undefined

## [ModifyKey()​](#modifykey)

修改字段中按其排序的键值

#### [参数​](#参数)

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| Key | Range | null | 是 | 指定要修改的键 |

#### [返回值​](#返回值-1)

undefined


#### 数据有效性规则(Validation)

# [Validation​](#validation)

代表工作表区域的数据有效性规则

Validation 对象的具体属性和方法请参阅下方的列表。

### [方法列表​](#方法列表)

| 方法名 | 返回类型 | 简介 |
| --- | --- | --- |
| Add() | undefined | 新增数据有效性规则 |
| Modify() | undefined | 修改数据有效性规则 |
| Delete() | undefined | 删除数据有效性规则 |

## [Add()​](#add)

新增数据有效性规则

### [参数​](#参数)

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| Type | Enum |  | 是 | 指定要对值进行的有效性测试的类型，可以是Enum.XlDVType中的值 |
| AlertStyle | Enum |  | 是 | 指定验证过程中显示的消息框所用的图表,可以是Enum.XlDVAlertStyle中的值 |
| Operator | Enum |  | 否 | 指定用于将公式与单元格的值或xlBetween和xlNoteBetween中的值进行比较,比较两个公式的运算符，可以是Enum.XlFormatConditionOperator中的值 |
| Formula1 | String |  | 否 | 数据验证公式中的第一部分，值不得超过 255 个字符 |
| Formula2 | String |  | 否 | 当 Operator 参数为 xlBetween 或 xlNotBetween 时, 数据验证等式的第二部分（否则, 将忽略此参数） |

### [示例​](#示例)

js
```js
// 数据有效性对象
let validation = Application.Range('A1').Validation

// 添加数据验证,如果输入的值不是整数并且不在1~5之间（包括1和5），则显示警告样式
validation.Add({
  Type: Application.Enum.XlDVType.xlValidateWholeNumber,
  AlertStyle: Application.Enum.XlDVAlertStyle.xlValidAlertWarning,
  Operator: Application.Enum.XlFormatConditionOperator.xlBetween,
  Formula1: '1',
  Formula2: '5'
})
```

## [Modify()​](#modify)

修改数据有效性规则

### [参数​](#参数-1)

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| Type | Enum |  | 是 | 指定要对值进行的有效性测试的类型，可以是Enum.XlDVType中的值 |
| AlertStyle | Enum |  | 是 | 指定验证过程中显示的消息框所用的图表,可以是Enum.XlDVAlertStyle中的值 |
| Operator | Enum |  | 否 | 指定用于将公式与单元格的值或xlBetween和xlNoteBetween中的值进行比较,比较两个公式的运算符，可以是Enum.XlFormatConditionOperator中的值 |
| Formula1 | String |  | 否 | 数据验证公式中的第一部分，值不得超过 255 个字符 |
| Formula2 | String |  | 否 | 当 Operator 参数为 xlBetween 或 xlNotBetween 时, 数据验证等式的第二部分（否则, 将忽略此参数） |

### [示例​](#示例-1)

js
```js
// 数据有效性对象
let validation = Application.Range('A1').Validation
// 修改数据验证规则
validation.Modify({
  Type: Application.Enum.XlDVType.xlValidateWholeNumber,
  AlertStyle: Application.Enum.XlDVAlertStyle.xlValidAlertWarning,
  Operator: Application.Enum.XlFormatConditionOperator.xlNotBetween,
  Formula1: '23',
  Formula2: '105'
})
```

## [Delete()​](#delete)

删除数据有效性规则

### [示例​](#示例-2)

js
```js
// 数据有效性对象
let validation = Application.Range('A1').Validation

validation.Delete()
```、
```


#### 数据表(Sheet)

# [Sheet​](#sheet)

工作簿（Workbook）中单个数据表(Sheet)对象

Sheet 对象的具体属性和方法请参阅下方的列表。

#### [属性列表​](#属性列表)

| 属性名 | 数据类型 | 简介 |
| --- | --- | --- |
| Id | String | 该数据表的 Id |
| Name | String | 该数据表的名称 |
| Index | Number | 该数据表在所有表的索引值 |
| Visible | Boolean | 该数据表是否可见 |
| Type | String | 该数据表的类型 |
| Field | Field | 该数据表的字段 |
| Record | Record | 该数据表的行记录 |

#### [方法列表​](#方法列表)

| 方法名 | 返回类型 | 简介 |
| --- | --- | --- |
| Activate() | undefined | 切换(激活)数据表 |
| Move() | undefined | 移动数据表 |
| Delete() | undefined | 删除数据表 |
| IsDBSheet() | Boolean | 是否为数据表 |

## [Id​](#id)

获取数据表 Id

#### [数据类型​](#数据类型)

String - 数据表 Id

#### [示例​](#示例)

js
```js
const sheet = Application.ActiveSheet
// 打印当前活动数据表的id
console.log(sheet.Id)
```

## [Name​](#name)

设置/获取 数据表名称

#### [数据类型​](#数据类型-1)

String - 该数据表在所有数据表的名称

#### [示例​](#示例-1)

js
```js
const sheet = Application.ActiveSheet
// 打印当前活动数据表的名称
console.log(sheet.Name) // Sheet2

// 将当前数据表的名称改为 WPS WebOffice
sheet.Name = 'WPS WebOffice'
```

## [Index​](#index)

数据表的 index,即该数据表在所有数据表的索引值

#### [数据类型​](#数据类型-2)

String - 该数据表在所有数据表的索引值

#### [示例​](#示例-2)

js
```js
const sheet = Application.ActiveSheet
// 打印当前活动数据表的index
console.log(sheet.Index) // 1
```

## [Visible​](#visible)

显示/隐藏 数据表

#### [数据类型​](#数据类型-3)

Boolean - 数据表是否可见

#### [示例​](#示例-3)

js
```js
const sheet = Application.ActiveSheet
// 隐藏数据表
sheet.Visible = false
// 取消数据表隐藏
sheet.Visible = true
```

## [Type​](#type)

数据表类型

#### [数据类型​](#数据类型-4)

Enum.xlEtDataBaseSheet- 数据表的类型

#### [示例​](#示例-4)

js
```js
const sheet = Application.ActiveSheet
// 打印当前活动数据表的类型
console.log(sheet.Type) //xlEtDataBaseSheet
```

## [Field​](#field)

数据表的字段， 返回一个Field对象

#### [数据类型​](#数据类型-5)

Field

#### [示例​](#示例-5)

js
```js
const sheet = Application.ActiveSheet
// 获取的表所有字段信息
const fields = sheet.Field.GetFields()
```

## [Record​](#record)

数据表的字段， 返回一个Record对象

#### [数据类型​](#数据类型-6)

Record

#### [示例​](#示例-6)

js
```js
const sheet = Application.ActiveSheet
const record = sheet.Record.GetRecord({  RecordId: 'Bz' })
```

## [Activate()​](#activate)

激活表

#### [示例​](#示例-7)

js
```js
const sheet = Application.Sheets.Item(1)
// 激活第一个表
sheet.Activate()
```

## [Move()​](#move)

移动数据表

#### [参数​](#参数)

两个参数互斥

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| Before | number | null | 否 | 验将放置移动的数据表之前的数据表 ID。如果指定 After ，则不能指定 Before。 |
| After | number | null | 否 | 将放置移动的数据表后的数据表 ID。如果指定 Before ，则不能指定 After |

#### [示例​](#示例-8)

js
```js
// 将当前数据表移动到第二个数据表之后
const sheet = Application.ActiveSheet
sheet.Move({
  Before: null,
  After: Application.Sheets(2).Id
})
```

## [Delete()​](#delete)

删除数据表

#### [返回值​](#返回值)

undefined

#### [示例​](#示例-9)

js
```js
// 删除名称为“Sheet2”的数据表
Application.Sheets.Item('Sheet2').Delete()
```

## [IsDBSheet()​](#isdbsheet)

是否为数据表

#### [返回值​](#返回值-1)

Boolean

#### [示例​](#示例-10)

js
```js
// 判断当前活跃表是否为数据表
Application.ActiveSheet.IsDBSheet()
```


#### 条件格式(FormatCondition)

# [FormatCondition​](#formatcondition)

区域内的某个条件格式

FormatCondition 对象的具体属性和方法请参阅下方的列表。

#### [属性列表​](#属性列表)

| 属性名 | 数据类型 | 简介 |
| --- | --- | --- |
| AppliesTo | Range | 应用格式规则的单元格区域 |
| Borders | Border | 返回一个 Borders 集合 |
| Font | Font | 返回一个 Font 对象 |
| Formula1 | String | 返回与条件格式或者数据有效性相关联的值或表达式 |
| Formula2 | String | 返回与条件格式或数据有效性验证第二部分相关联的值或表达式 |
| Interior | Interior | 表示指定对象的内部 |
| NumberFormat | String | 单元格的数字格式 |
| Operator | XlFormatConditionOperator | 条件格式的运算符 |
| Priority | Number | 返回或设置条件格式规则的优先级值 |
| Type | XlFormatConditionType | 条件格式对象类型 |

#### [方法列表​](#方法列表)

| 方法名 | 返回类型 | 简介 |
| --- | --- | --- |
| Modify() | undefined | 更改现有条件格式 |
| ModifyAppliesToRange() | undefined | 设置此格式规则所应用于的单元格区域 |
| SetFirstPriority() | undefined | 将此条件格式规则的优先级值设置为“1” |
| SetLastPriority() | undefined | 将此条件格式规则的优先级值增加“1” |

## [AppliesTo​](#appliesto)

返回一个 Range 对象，该对象指应用格式规则的单元格区域

#### [数据类型​](#数据类型)

Range- 区域对象

#### [示例​](#示例)

js
```js
// 计算A列的值为2，并且B列的值大于B列中A列的值为2的平均数
const ave = WorksheetFunction.AverageIf(Range('A9:A30'), 2, Range('B9:B30'))
// 定义公式:A列的值为2，并且B列的值大于B列中A列的值为2的平均数
const formula1 = `=AND(A9=2, B9>${ave})`
// 指定B列添加条件格式
const targetFormulaRange = Range('B9:B30')
const formatCondition1 = targetFormulaRange.FormatConditions.Add(
  xlExpression,
  -1,
  formula1,
  ''
)
// 获取条件格式应用的区域
const appliesToRange = formatCondition1.AppliesTo
// 打印目标区域的地址
console.log(appliesToRange.Address())
```

## [Borders​](#borders)

返回一个 Borders 集合，该集合表示样式或单元格区域的边框 (包括定义为条件格式) 的一部分的区域

#### [数据类型​](#数据类型-1)

Border集合 - 边框对象集合

#### [示例​](#示例-1)

js
```js
// 计算A列的值为2，并且B列的值大于B列中A列的值为2的平均数
const ave = WorksheetFunction.AverageIf(Range('A9:A30'), 2, Range('B9:B30'))
// 定义公式:A列的值为2，并且B列的值大于B列中A列的值为2的平均数
const formula1 = `=AND(A9=2, B9>${ave})`
// 指定B列添加条件格式
const targetFormulaRange = Range('B9:B30')
const formatCondition1 = targetFormulaRange.FormatConditions.Add(
  xlExpression,
  -1,
  formula1,
  ''
)
// 设置条件格式边框颜色
formatCondition1.Borders.Item(xlDiagonalDown).Color = '#FFFF00'
```

## [Font​](#font)

返回一个 Font 对象，该对象表示指定对象的字体

#### [数据类型​](#数据类型-2)

Font- 字体对象

#### [示例​](#示例-2)

js
```js
// 计算A列的值为2，并且B列的值大于B列中A列的值为2的平均数
const ave = WorksheetFunction.AverageIf(Range('A9:A30'), 2, Range('B9:B30'))
// 定义公式:A列的值为2，并且B列的值大于B列中A列的值为2的平均数
const formula1 = `=AND(A9=2, B9>${ave})`
// 指定B列添加条件格式
const targetFormulaRange = Range('B9:B30')
const formatCondition1 = targetFormulaRange.FormatConditions.Add(
  xlExpression,
  -1,
  formula1,
  ''
)
// 获取条件格式字体对象
const font = formatCondition1.Font
// 打印字体颜色
console.log(font.Color)
```

## [Formula1​](#formula1)

返回与条件格式或者数据有效性相关联的值或表达式

Formula1 属性可以是常量值、字符串值、单元格引用或公式

#### [数据类型​](#数据类型-3)

String - 值或者表达式

#### [示例​](#示例-3)

js
```js
// 计算A列的值为2，并且B列的值大于B列中A列的值为2的平均数
const ave = WorksheetFunction.AverageIf(Range('A9:A30'), 2, Range('B9:B30'))
// 定义公式:A列的值为2，并且B列的值大于B列中A列的值为2的平均数
const formula1 = `=AND(A9=2, B9>${ave})`
// 指定B列添加条件格式
const targetFormulaRange = Range('B9:B30')
const formatCondition1 = targetFormulaRange.FormatConditions.Add(
  xlExpression,
  -1,
  formula1,
  ''
)
// 打印条件格式条件1
console.log(formatCondition1.Formula1)
```

## [Formula2​](#formula2)

返回与条件格式或数据有效性验证第二部分相关联的值或表达式

仅当数据验证条件格式 Operator 属性为 xlBetween 或 xlNotBetween 时，才使用 Formula2 属性。可为常量值、字符串值、单元格引用或公式

#### [数据类型​](#数据类型-4)

String - 值或者表达式

#### [示例​](#示例-4)

js
```js
//1.获取F列的条件格式
const formatconditions = Range('F:F').FormatConditions
//2.增加介于90到100的条件格式
const betweenCondition = formatconditions.Add(
  xlCellValue,
  xlBetween,
  '90',
  '100'
)
//3.打印条件格式条件二
console.log(betweenCondition.formula2)
```

## [Interior​](#interior)

返回一个 Interior 对象，该对象表示指定对象的内部

#### [数据类型​](#数据类型-5)

Interior - 内部对象

#### [示例​](#示例-5)

js
```js
// 获取B2到G18的range
const range = Range('B2:G18')
// 清除range的条件格式
range.FormatConditions.Delete()
// 通过新增条件格式设置高亮重复项
let formatCondition = range.FormatConditions.AddUniqueValues()
// 将DupeUnique设置为xlDuplicate，代表条件格式显示重复值
formatCondition.DupeUnique = xlDuplicate
const redColor = RGB(255, 0, 0)
// 将重复值的颜色设为红色
formatCondition.Interior.Color = redColor
```

## [NumberFormat​](#numberformat)

在条件格式规则的计算结果为 True 时返回或设置应用于单元格的数字格式

数字格式是使用“单元格格式”对话框的“数字”选项卡上显示的相同格式代码指定的。

您可以使用内置的数字格式，例如 "General" 或者创建自定义数字格式。

#### [数据类型​](#数据类型-6)

String - 数字格式

#### [示例​](#示例-6)

js
```js
let targetFormulaRange = Range('B:B')
let formatCondition1 = targetFormulaRange.FormatConditions
// 设置满足条件时的格式，B列最大值增加小数点位
formatCondition1.NumberFormat = '0.0'
```

## [Operator​](#operator)

返回条件格式的运算符

#### [数据类型​](#数据类型-7)

XlFormatConditionType- 条件格式操作类型枚举

#### [示例​](#示例-7)

js
```js
//1.获取F列的条件格式
const formatconditions = Range('F:F').FormatConditions
//2.增加介于90到100的条件格式
const betweenCondition = formatconditions.Add(xlCellValue, xlBetween, '90', '100')
//3.打印条件格式操作符
console.log(betweenCondition.Operator)
```

## [Priority​](#priority)

返回或设置条件格式规则的优先级值。 当工作表中存在多个条件格式规则时，优先级确定求值的顺序

#### [数据类型​](#数据类型-8)

Number - 优先级

#### [示例​](#示例-8)

js
```js
//1.获取F列的条件格式
const formatconditions = Range('F:F').FormatConditions
//2.增加介于90到100的条件格式
const betweenCondition = formatconditions.Add(xlCellValue, xlBetween, '90', '100')
//3.打印条件格式优先级
console.log(betweenCondition.Priority)
```

## [Type​](#type)

返回对象类型

对象类型可以是以下XlFormatConditionType枚举中的一个:

xlCellValue

xlExpression

#### [数据类型​](#数据类型-9)

XlFormatConditionType- 条件格式类型

#### [示例​](#示例-9)

js
```js
//1.获取F列的条件格式
const formatconditions = Range('F:F').FormatConditions
//2.增加介于90到100的条件格式
const betweenCondition = formatconditions.Add(xlCellValue, xlBetween, '=90', '=100')
// 打印条件格式类型
console.log(betweenCondition.Type)
```

## [Modify()​](#modify)

更改现有条件格式

### [参数​](#参数)

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| Type | XlFormatConditionType | null | 是 | 指定条件格式是基于单元格值还是基于表达式 |
| Operator | XlFormatConditionOperator | null | 否 | 条件格式运算符，如果 Type 为 xlExpression，则忽略参数 Operator |
| Formula1 | XlFormatConditionType | null | 否 | 与条件格式关联的值或表达式。 可为常量值、字符串值、单元格引用或公式 |
| Formula2 | XlFormatConditionType | null | 否 | 与条件格式关联的值或表达式。 可为常量值、字符串值、单元格引用或公式 |

### [返回值​](#返回值)

undefined

#### [示例​](#示例-10)

js
```js
//1.获取F列的条件格式
const formatconditions = Range('F:F').FormatConditions
//2.增加介于90到100的条件格式
const betweenCondition = formatconditions.Add(xlCellValue, xlBetween, '=90', '=100')
//3.更改条件格式，改为90到110
betweenCondition.Modify(xlCellValue, xlBetween, '=90', '=110')
```

## [ModifyAppliesToRange()​](#modifyappliestorange)

设置此格式规则所应用于的单元格区域

该区域必须采用 A1 引用样式，并且完全包含在作为集合父级的工作表中 FormatConditions 。 可包括区域操作符（冒号）、相交区域操作符（空格）或合并区域操作符（逗号），也可以使用货币符号，但会被忽略

### [参数​](#参数-1)

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| Range | Range | null | 是 | 此格式规则将应用于的区域 |

### [返回值​](#返回值-1)

undefined

### [示例​](#示例-11)

js
```js
// 把B列上的条件格式移动到A列上

//获取B列的条件格式
const formatConditions = Range('B:B').FormatConditions
//设置目的区域
const destRange = Range('A:A')
//获取条件格式的个数
const count = formatConditions.Count
for (let i = 1; i <= count; i++) {
  //修改条件格式的range为目的区域
  formatConditions.Item(i).ModifyAppliesToRange(destRange)
}
```

## [SetFirstPriority()​](#setfirstpriority)

将此条件格式规则的优先级值设置为“1”，以便在工作表上的所有其他规则之前计算此规则

当工作表中有多个条件格式规则时，如果该规则以前未设置为优先级“1”，则此方法将导致工作表上所有其他现有规则的优先级增加一个

注意

条件格式规则的优先级基于工作表级别应用。

### [返回值​](#返回值-2)

undefined

### [示例​](#示例-12)

js
```js
/**问题
 将所有A20:A100中值为"张三"的行中，D20:D100日期为本月的单元格填充为红色
 将所有A11:A20中值为"李四"的行中，D11:D20日期为下月的单元格填充绿色
 */

// 定义公式:A列为张三，D列单元格的年为本年，月为本月时
let formula1 =
  '=AND(A20="张三",YEAR(D20)=YEAR(TODAY()),MONTH(D20)=MONTH(TODAY()))'
// 直接指定D2:D10添加条件格式
let targetFormulaRange = Range('D20:D100')
const formatCondition1 = targetFormulaRange.FormatConditions.Add(
  xlExpression,
  -1,
  formula1,
  ''
)
// SetFirstPriority设置最高优先级，如果设置最低优先级:SetLastPriority
formatCondition1.SetFirstPriority()
// 将满足条件的格式单元格填充红色
formatCondition1.Interior.Pattern = xlPatternSolid
formatCondition1.Interior.Color = RGB(255, 0, 0)

// 定义公式:A列为李四，D列单元格的年为下月所在年，月为下月时
let formula2 =
  '=AND(A11="李四",YEAR(D11)=YEAR(EOMONTH(TODAY(), 1)),MONTH(D11)=MONTH(EOMONTH(TODAY(), 1)))'
// 直接指定D列添加条件格式
const formatCondition2 = Range('D11:D20').FormatConditions.Add(
  xlExpression,
  -1,
  formula2,
  ''
)
// 设置最低优先级
formatCondition2.SetLastPriority()
// 将满足条件的格式单元格填充绿色
formatCondition2.Interior.Pattern = xlPatternSolid
formatCondition2.Interior.Color = RGB(0, 255, 0)
```

## [SetLastPriority()​](#setlastpriority)

为此条件格式规则设置求值顺序，以便在工作表上的所有其他规则之后计算此规则

优先级的实际值将等于工作表上条件格式规则的总数。 如果工作表中有多个条件格式规则，此方法将导致优先级值大于此规则的规则的优先级增加 1

### [返回值​](#返回值-3)

undefined

### [示例​](#示例-13)

js
```js
/**问题
 将所有A20:A100中值为"张三"的行中，D20:D100日期为本月的单元格填充为红色
 将所有A11:A20中值为"李四"的行中，D11:D20日期为下月的单元格填充绿色
 */

// 定义公式:A列为张三，D列单元格的年为本年，月为本月时
let formula1 =
  '=AND(A20="张三",YEAR(D20)=YEAR(TODAY()),MONTH(D20)=MONTH(TODAY()))'
// 直接指定D2:D10添加条件格式
let targetFormulaRange = Range('D20:D100')
const formatCondition1 = targetFormulaRange.FormatConditions.Add(
  xlExpression,
  -1,
  formula1,
  ''
)
// SetFirstPriority设置最高优先级，如果设置最低优先级:SetLastPriority
formatCondition1.SetFirstPriority()
// 将满足条件的格式单元格填充红色
formatCondition1.Interior.Pattern = xlPatternSolid
formatCondition1.Interior.Color = RGB(255, 0, 0)

// 定义公式:A列为李四，D列单元格的年为下月所在年，月为下月时
let formula2 =
  '=AND(A11="李四",YEAR(D11)=YEAR(EOMONTH(TODAY(), 1)),MONTH(D11)=MONTH(EOMONTH(TODAY(), 1)))'
// 直接指定D列添加条件格式
const formatCondition2 = Range('D11:D20').FormatConditions.Add(
  xlExpression,
  -1,
  formula2,
  ''
)
// 设置最低优先级
formatCondition2.SetLastPriority()
// 将满足条件的格式单元格填充绿色
formatCondition2.Interior.Pattern = xlPatternSolid
formatCondition2.Interior.Color = RGB(0, 255, 0)
```


#### 条件格式集合(FormatConditions)

# [FormatConditions​](#formatconditions)

FormatConditions 集合对象用于控制 Excel 中的条件格式。

条件格式是一种在工作表中格式化单元格的方法，可以根据单元格的值、公式或其他条件自动应用格式，使数据更易于理解和分析。

#### [属性列表​](#属性列表)

| 属性名 | 数据类型 | 简介 |
| --- | --- | --- |
| Count | Number | 返回 FormatConditions 集合中的对象数 |

#### [方法列表​](#方法列表)

| 方法名 | 返回类型 | 简介 |
| --- | --- | --- |
| Add() | FormatCondition | 向 FormatConditions 集合中添加一个条件格式 |
| AddAboveAverage() | AboveAverage | 返回表示指定区域的条件格式规则的新 AboveAverage 对象 |
| AddIconSetCondition() | IconSetCondition | 代表指定区域的图标集条件格式规则 |
| AddColorScale() | ColorScale | 该条件格式规则使用单元格颜色中的渐变来指示所选区域中包含的单元格值的相对差异 |
| AddTop10() | Top10 | 该条件格式可以根据指定的截止值查找单元格区域中的最高值和最低值 |
| AddUniqueValues() | UniqueValues | 返回表示指定区域的条件格式规则的新 UniqueValues 对象 |
| Delete() | undefined | 删除该区域下的条件格式 |
| Item() | FormatCondition | 从条件格式集合中返回一个条件格式对象 |

## [Count​](#count)

返回 FormatConditions 集合中的对象数

#### [数据类型​](#数据类型)

Number - 对象数量

#### [示例​](#示例)

js
```js
// 获取FormatConditions对象
const formatConditions = Range('A:A').FormatConditions
// 获取条件格式对象数量
const count = formatConditions.Count

console.log(count)
```

## [Add()​](#add)

向 FormatConditions 集合中添加一个条件格式

#### [参数​](#参数)

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| Type | XlFormatConditionType | null | 是 | 指定条件格式的类型 |
| Operator | XlFormatConditionOperator | null | 是 | 指定条件格式的运算符，请注意：如果 Type 为xlExpression，则忽略 Operator 参数 |
| Formula1 | String | null | 是 | 与条件格式关联的值或表达式。 可为常量值、字符串值、单元格引用或公式 |
| Formula2 | String | null | 否 | 当 Operator 为xlBetween或xlNotBetween时，表示与条件格式的第二部分关联的值或表达式 (否则，将忽略此参数)。可以是常量值、字符串值、单元格引用或公式 |

#### [返回值​](#返回值)

FormatCondition- 返回一个FormatCondition对象，表示添加的条件格式

#### [示例​](#示例-1)

js
```js
// 计算A列的值为2，并且B列的值大于B列中A列的值为2的平均数
const ave = WorksheetFunction.AverageIf(Range('A9:A30'), 2, Range('B9:B30'))
// 定义公式:A列的值为2，并且B列的值大于B列中A列的值为2的平均数
const formula1 = `=AND(A9=2, B9>${ave})`
// 指定B列添加条件格式
const targetFormulaRange = Range('B9:B30')
const formatCondition1 = targetFormulaRange.FormatConditions.Add(
  xlExpression,
  -1,
  formula1,
  ''
)
// SetFirstPriority设置最高优先级，如果设置最低优先级:SetLastPriority
formatCondition1.SetFirstPriority()
// 设置满足条件时的格式，设置删除线
formatCondition1.Font.Strikethrough = true
```

## [AddAboveAverage()​](#addaboveaverage)

返回表示指定区域的条件格式规则的新 AboveAverage 对象

对象 AboveAverage 用于在单元格区域中查找高于或低于平均值或标准偏差的值。例如，可在年度业绩评估中查找高于平均业绩的人员。

#### [返回值​](#返回值-1)

AboveAverage 对象

#### [示例​](#示例-2)

js
```js
//1.获取D列的条件格式
const formatconditions = Range('D:D').FormatConditions
//2.增加高于平均值条件格式
const aboveAverageCondition = formatconditions.AddAboveAverage()
//3.如果判断高于平均值使用xlAboveAverage，如果判断低于平均值使用xlBelowAverage
aboveAverageCondition.AboveBelow = xlAboveAverage
aboveAverageCondition.SetFirstPriority()
//4.填充颜色设置为黄色
aboveAverageCondition.Font.Color = RGB(255, 255, 0)
```

## [AddIconSetCondition()​](#addiconsetcondition)

返回一个新的 IconSetCondition 对象，该对象代表指定区域的图标集条件格式规则。

使用图标集为数据添加注释并将数据分为按阈值隔开的三到五类数据，每种图标均代表某一值范围。

#### [返回值​](#返回值-2)

IconSetCondition 对象

#### [示例​](#示例-3)

js
```js
// 用条件格式的图标集将D列和E列标记出来，分值梯度为>=130，90~130，<=90

//应用区域忽略表头
const bIgnoreHeader = true
//获取目标工作簿
const targetWorkbook = ActiveWorkbook
//获取目标工作表
const targetWorksheet = ActiveSheet

const targetColumns = ['D', 'F']
for (let i = 0; i < targetColumns.length; ++i) {
  //选中目标区域
  const targetColumn = targetColumns[i]
  const beginRow = targetWorksheet.UsedRange.Row + (bIgnoreHeader ? 1 : 0)
  const endRow =
    targetWorksheet.UsedRange.Row + targetWorksheet.UsedRange.Rows.Count - 1
  const targetRangeStr = targetColumn + beginRow + ':' + targetColumn + endRow
  let targetRange = targetWorksheet.Range(targetRangeStr)
  //选中条件格式应用区域
  targetRange.Select()
  //在targetRange上添加新的IconSetCondition对象，代表图标集条件格式
  let iconSetCondition = targetRange.FormatConditions.AddIconSetCondition()
  //将此条件格式规则的优先级值设置为 1
  iconSetCondition.SetFirstPriority()
  //将三个交通灯图标应用于条件格式
  iconSetCondition.IconSet = targetWorkbook.IconSets.Item(xl3TrafficLights1)

  const maxIndex = 3
  const middleIndex = 2
  //设置最大的条件为单元格值>=130
  let IconCriterion3 = iconSetCondition.IconCriteria.Item(maxIndex)
  IconCriterion3.Type = xlConditionValueNumber
  IconCriterion3.Operator = xlGreaterEqual
  IconCriterion3.Value = 130

  //设置中间条件为单元格值>90
  let IconCriterion2 = iconSetCondition.IconCriteria.Item(middleIndex)
  IconCriterion2.Type = xlConditionValueNumber
  IconCriterion2.Operator = xlGreater
  IconCriterion2.Value = 90
}
```

## [AddColorScale()​](#addcolorscale)

返回一个新的 ColorScale 对象，该对象表示条件格式规则，该规则使用单元格颜色中的渐变来指示所选区域中包含的单元格值的相对差异。

#### [参数​](#参数-1)

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| ColorScaleType | Number | null | 是 | 色阶的类型，例如传 2 就代表双色刻度，传 3 就代表三色刻度 |

#### [返回值​](#返回值-3)

ColorScale 对象

#### [示例​](#示例-4)

js
```js
// 对B列按照最低值设置双色刻度，其中最小刻度为红色，最大刻度为绿色。

//获取B列的条件格式
const formatConditions = Range('B:B').FormatConditions
//增加色阶条件格式，参数2代表2阶，3代表3阶
const color2ScaleCondition = formatConditions.AddColorScale(2)
color2ScaleCondition.SetFirstPriority()
//获取2阶色阶的第一个色阶
const colorScaleCriteria1 = color2ScaleCondition.ColorScaleCriteria.Item(1)
//设置类型。xlConditionValueLowestValue代表最低值。xlConditionValueHighestValue代表最高值。xlConditionValuePercent代表使用百分之比。xlConditionValuePercentile代表使用百分点值。xlConditionValueNumber代表使用数字
colorScaleCriteria1.Type = xlConditionValueLowestValue
//设置色阶颜色值， 255代表红色
colorScaleCriteria1.FormatColor.Color = RGB(255, 0, 0)
//获取2阶色阶的第二个色阶
const colorScaleCriteria2 = color2ScaleCondition.ColorScaleCriteria.Item(2)
//设置类型。xlConditionValueLowestValue代表最低值。xlConditionValueHighestValue代表最高值。xlConditionValuePercent代表使用百分之比。xlConditionValuePercentile代表使用百分点值。xlConditionValueNumber代表使用数字
colorScaleCriteria2.Type = xlConditionValueHighestValue
//设置颜色为绿色
colorScaleCriteria2.FormatColor.Color = RGB(0, 255, 0)
```

## [AddTop10()​](#addtop10)

返回一个 Top10 对象，该对象表示指定区域的条件格式规则

Top10 使用 对象，可以根据指定的截止值查找单元格区域中的最高值和最低值。 例如，可以查找区域报告中位居前五位的销售产品、客户调查中位居最后百分之十五的产品，或者部门人员分析中位居前 25 位的薪金。

#### [返回值​](#返回值-4)

Top10 对象

#### [示例​](#示例-5)

js
```js
// 把L列数值最低的5个标记为绿色字体

//应用区域忽略表头
const bIgnoreHeader = true

//获取目标工作表
const targetWorksheet = ActiveSheet

//选中目标区域
const targetColumn = 'L'
const beginRow = targetWorksheet.UsedRange.Row + (bIgnoreHeader ? 1 : 0)
const endRow =
  targetWorksheet.UsedRange.Row + targetWorksheet.UsedRange.Rows.Count - 1
const targetRangeStr = targetColumn + beginRow + ':' + targetColumn + endRow
let targetRange = targetWorksheet.Range(targetRangeStr)
//选中条件格式应用区域
targetRange.Select()

//基于前十项规则添加条件格式
let top10 = targetRange.FormatConditions.AddTop10()
//标记前5时top10.Percent=false，标记前5%时top10.Percent=true
top10.Percent = false
//设置排名值为5
top10.Rank = 5
//标记前5时top10.TopBottom=xlTop10Top
//标记后5时top10.TopBottom=xlTop10Bottom
top10.TopBottom = xlTop10Bottom
//将条件格式字体标绿
const greenColor = RGB(0, 255, 0)
top10.Font.Color = greenColor
top10.Font.TintAndShade = 0
top10.StopIfTrue = false
```

## [AddUniqueValues()​](#adduniquevalues)

返回表示指定区域的条件格式规则的新 UniqueValues 对象

可以使用 UniqueValues 对象快速可视化包含唯一值或重复值的单元格

#### [返回值​](#返回值-5)

UniqueValues 对象

#### [示例​](#示例-6)

js
```js
// 将B2到G18设置高亮重复项背景颜色为红色

// 获取B2到G18的range
const range = Range('B2:G18')
// 清除range的条件格式
range.FormatConditions.Delete()
// 通过新增条件格式设置高亮重复项
let formatCondition = range.FormatConditions.AddUniqueValues()
// 将DupeUnique设置为xlDuplicate，代表条件格式显示重复值
formatCondition.DupeUnique = xlDuplicate
const redColor = RGB(255, 0, 0)
// 将重复值的颜色设为红色
formatCondition.Interior.Color = redColor
```

## [Delete()​](#delete)

删除该区域下的条件格式

#### [返回值​](#返回值-6)

undefined

js
```js
// 删除条件格式
const targetFormulaRange = Range('B9:B30')
const formatCondition1 = targetFormulaRange.FormatConditions.Delete()
```

## [Item()​](#item)

从条件格式集合中返回一个条件格式对象

#### [参数​](#参数-2)

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| Index | Number / String | null | 是 | 目标对象在集合内的索引值，从 1 开始 |

#### [返回值​](#返回值-7)

FormatCondition对象

#### [示例​](#示例-7)

js
```js
// 把B列上的条件格式移动到A列上

//获取B列的条件格式
const formatConditions = Range('B:B').FormatConditions
//设置目的区域
const destRange = Range('A:A')
//获取条件格式的个数
const count = formatConditions.Count
for (let i = 1; i <= count; i++) {
  //修改条件格式的range为目的区域
  formatConditions.Item(i).ModifyAppliesToRange(destRange)
}
```


#### 枚举(Enum)

# [Enum​](#enum)

枚举类型，存放在 Application 下

## [XlAboveBelow​](#xlabovebelow)

指定值是高于还是低于平均值

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlAboveAverage | 0 | 高于平均值 |
| xlBelowAverage | 1 | 低于平均值 |
| xlEqualAboveAverage | 2 | 等于平均值 |

## [XlAutoFillType​](#xlautofilltype)

根据源区域的内容，指定目标区域的填充方式

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlFillDefault | 0 | 确定用于填充目标区域的值和格式 |
| xlFillCopy | 1 | 将源区域的值和格式复制到目标区域，如有必要可重复执行 |
| xlFillSeries | 2 | 将源区域中的值扩展到目标区域中，形式为系列（如，“1, 2” 扩展为 “3, 4, 5”）。格式从源区域复制到目标区域，如有必要可重复执行 |
| xlFillFormats | 3 | 只将源区域的格式复制到目标区域，如有必要可重复执行 |
| xlFillValues | 4 | 只将源区域的值复制到目标区域，如有必要可重复执行 |
| xlFillDays | 5 | 将星期中每天的名称从源区域扩展到目标区域中。格式从源区域复制到目标区域，如有必要可重复执行 |
| xlFillWeekdays | 6 | 将工作周每天的名称从源区域扩展到目标区域中。格式从源区域复制到目标区域，如有必要可重复执行 |
| xlFillMonths | 7 | 将月名称从源区域扩展到目标区域中。格式从源区域复制到目标区域，如有必要可重复执行 |
| xlFillYears | 8 | 将年从源区域扩展到目标区域中。格式从源区域复制到目标区域，如有必要可重复执行 |
| xlLinearTrend | 9 | 将数值从源区域扩展到目标区域中，假定数字之间是加法关系（如，“1, 2,” 扩展为 “3, 4, 5”，假定每个数字都是前一个数字加上某个值的结果）。格式从源区域复制到目标区域，如有必要可重复执行 |
| xlGrowthTrend | 10 | 将数值从源区域扩展到目标区域中，假定源区域的数字之间是乘法关系（如，“1, 2,” 扩展为 “4, 8, 16”，假定每个数字都是前一个数字乘以某个值的结果）。格式从源区域复制到目标区域，如有必要可重复执行 |

## [XlBordersIndex​](#xlbordersindex)

指定要检索的边框

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlDiagonalDown | 5 | 从区域中每个单元格的左上角到右下角的边框 |
| xlDiagonalUp | 6 | 从区域中每个单元格的左下角到右上角的边框 |
| xlEdgeLeft | 7 | 区域左边缘的边框 |
| xlEdgeTop | 8 | 区域顶部的边框 |
| xlEdgeBottom | 9 | 区域底部的边框 |
| xlEdgeRight | 10 | 区域右边缘的边框 |
| xlInsideVertical | 11 | 区域中所有单元格的垂直边框（区域以外的边框除外） |
| xlInsideHorizontal | 12 | 区域中所有单元格的水平边框（区域以外的边框除外） |
| xlOutside | 13 | 区域中的 上下左右 |
| xlInside | 14 | 中间区域 |

## [XlBorderWeight​](#xlborderweight)

指定某一区域周围的边框的粗细

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlMedium | -4138 | 中 |
| xlHairline | 1 | 细线（最细的边框） |
| xlThin | 2 | 细长 |
| xlThick | 4 | 粗（最宽的边框） |

## [XlCalcModeType​](#xlcalcmodetype)

迭代计算模式

| 字段 | 值 | 释义 |
| --- | --- | --- |
| manual | manual | 手动 |
| automatic | automatic | 自动 |

## [XlChartType​](#xlcharttype)

指定图表类型

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xl3DArea | -4098 | 三维面积图。 |
| xl3DAreaStacked | 78 | 三维堆积面积图。 |
| xl3DAreaStacked100 | 79 | 百分比堆积面积图。 |
| xl3DBarClustered | 60 | 三维簇状条形图。 |
| xl3DBarStacked | 61 | 三维堆积条形图。 |
| xl3DBarStacked100 | 62 | 三维百分比堆积条形图。 |
| xl3DColumn | -4100 | 三维柱形图。 |
| xl3DColumnClustered | 54 | 三维簇状柱形图。 |
| xl3DColumnStacked | 55 | 三维堆积柱形图。 |
| xl3DColumnStacked100 | 56 | 三维百分比堆积柱形图。 |
| xl3DLine | -4101 | 三维折线图。 |
| xl3DPie | -4102 | 三维饼图。 |
| xl3DPieExploded | 70 | 分离型三维饼图。 |
| xlArea | 1 | 面积图 |
| xlAreaStacked | 76 | 堆积面积图。 |
| xlAreaStacked100 | 77 | 百分比堆积面积图。 |
| xlBarClustered | 57 | 簇状条形图。 |
| xlBarOfPie | 71 | 复合条饼图。 |
| xlBarStacked | 58 | 堆积条形图。 |
| xlBarStacked100 | 59 | 百分比堆积条形图。 |
| xlBubble | 15 | 气泡图。 |
| xlBubble3DEffect | 87 | 三维气泡图。 |
| xlColumnClustered | 51 | 簇状柱形图。 |
| xlColumnStacked | 52 | 堆积柱形图。 |
| xlColumnStacked100 | 53 | 百分比堆积柱形图。 |
| xlConeBarClustered | 102 | 簇状条形圆锥图。 |
| xlConeBarStacked | 103 | 堆积条形圆锥图。 |
| xlConeBarStacked100 | 104 | 百分比堆积条形圆锥图。 |
| xlConeCol | 105 | 三维柱形圆锥图。 |
| xlConeColClustered | 99 | 簇状柱形圆锥图。 |
| xlConeColStacked | 100 | 堆积柱形圆锥图。 |
| xlConeColStacked100 | 101 | 百分比堆积柱形圆锥图。 |
| xlCylinderBarClustered | 95 | 簇状条形圆柱图。 |
| xlCylinderBarStacked | 96 | 堆积条形圆柱图。 |
| xlCylinderBarStacked100 | 97 | 百分比堆积条形圆柱图。 |
| xlCylinderCol | 98 | 三维柱形圆柱图。 |
| xlCylinderColClustered | 92 | 簇状柱形圆锥图。 |
| xlCylinderColStacked | 93 | 堆积柱形圆锥图。 |
| xlCylinderColStacked100 | 94 | 百分比堆积柱形圆柱图。 |
| xlDoughnut | -4120 | 圆环图。 |
| xlDoughnutExploded | 80 | 分离型圆环图。 |
| xlLine | 4 | 折线图。 |
| xlLineMarkers | 65 | 数据点折线图。 |
| xlLineMarkersStacked | 66 | 堆积数据点折线图。 |
| xlLineMarkersStacked100 | 67 | 百分比堆积数据点折线图。 |
| xlLineStacked | 63 | 堆积折线图。 |
| xlLineStacked100 | 64 | 百分比堆积折线图。 |
| xlPie | 5 | 饼图。 |
| xlPieExploded | 69 | 分离型饼图。 |
| xlPieOfPie | 68 | 复合饼图。 |
| xlPyramidBarClustered | 109 | 簇状条形棱锥图。 |
| xlPyramidBarStacked | 110 | 堆积条形棱锥图。 |
| xlPyramidBarStacked100 | 111 | 百分比堆积条形棱锥图。 |
| xlPyramidCol | 112 | 三维柱形棱锥图。 |
| xlPyramidColClustered | 106 | 簇状柱形棱锥图。 |
| xlPyramidColStacked | 107 | 堆积柱形棱锥图。 |
| xlPyramidColStacked100 | 108 | 百分比堆积柱形棱锥图。 |
| xlRadar | -4151 | 雷达图。 |
| xlRadarFilled | 82 | 填充雷达图。 |
| xlRadarMarkers | 81 | 数据点雷达图。 |
| xlStockHLC | 88 | 盘高-盘低-收盘图。 |
| xlStockOHLC | 89 | 开盘-盘高-盘低-收盘图。 |
| xlStockVHLC | 90 | 成交量-盘高-盘低-收盘图。 |
| xlStockVOHLC | 91 | 成交量-开盘-盘高-盘低-收盘图。 |
| xlSurface | 83 | 三维曲面图。 |
| xlSurfaceTopView | 85 | 曲面图（俯视图）。 |
| xlSurfaceTopViewWireframe | 86 | 曲面图（俯视线框图）。 |
| xlSurfaceWireframe | 84 | 三维曲面图（线框）。 |
| xlXYScatter | -4169 | 散点图。 |
| xlXYScatterLines | 74 | 折线散点图。 |
| xlXYScatterLinesNoMarkers | 75 | 无数据点折线散点图。 |
| xlXYScatterSmooth | 72 | 平滑线散点图。 |
| xlXYScatterSmoothNoMarkers | 73 | 无数据点平滑线散点图。 |

## [XlContainsOperator​](#xlcontainsoperator)

指定函数使用的运算符

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlContains | 0 | 包含指定的值 |
| xlDoesNotContain | 1 | 不包含指定的值 |
| xlBeginsWith | 2 | 以指定的值开始 |
| xlEndsWith | 3 | 以指定的值结束 |

## [XlDeleteShiftDirection​](#xldeleteshiftdirection)

指定如何移动单元格来替换删除的单元格

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlShiftToLeft | -4159 | 单元格向左移动 |
| xlShiftUp | -4162 | 单元格向上移动 |

## [XlDirection​](#xldirection)

指定移动的方向

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlDown | -4121 | 下拉 |
| xlToLeft | -4159 | 向左 |
| xlToRight | -4161 | 向右 |
| xlUp | -4162 | 加速 |

## [XlDVAlertStyle​](#xldvalertstyle)

指定验证过程中显示的消息框所用的图标

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlValidAlertStop | 1 | 停止图标 |
| xlValidAlertWarning | 2 | 警告图标 |
| xlValidAlertInformation | 3 | 信息图标 |

## [XlDVType​](#xldvtype)

指定要对值进行的有效性测试类型

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlValidateWholeNumber | 1 | 全部数值 |
| xlValidateDecimal | 2 | 数值 |
| xlValidateList | 3 | 值必须存在于指定列表中 |
| xlValidateDate | 4 | 日期值 |
| xlValidateTime | 5 | 时间值 |
| xlValidateTextLength | 6 | 文本长度 |
| xlValidateCustom | 7 | 使用任意公式验证数据有效性 |

## [XlExportImgFormatType​](#xlexportimgformattype)

导出图片的格式

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlImgTypePNG | 0 | 导出 .png |
| xlImgTypeJPG | 1 | 导出 .jpg |
| xlImgTypeBMP | 2 | 导出 .bmp |
| xlImgTypeTIF | 2 | 导出 .tif |

## [XlFixedFormatType​](#xlfixedformattype)

指定文件格式的类型

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlTypePDF | 0 | PDF（.pdf） |
| xlTypeXPS | 1 | XPS（.xps） |
| xlTypeIMG | 2 | IMG（.png、.jpg、.bmp、.tif） |

## [XlFormatConditionOperator​](#xlformatconditionoperator)

指定用于将公式与单元格中的值或 xlBetween 和 xlNotBetween 中的值进行比较，以比较两个公式的运算符

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlBetween | 1 | 行间，只在提供了两个公式的情况下才能使用 |
| xlNotBetween | 2 | 不介于。只在提供了两个公式的情况下才能使用 |
| xlEqual | 3 | 平等 |
| xlNotEqual | 4 | 不等于 |
| xlGreater | 5 | 大于 |
| xlLess | 6 | 小于 |
| xlGreaterEqual | 7 | 大于或等于 |
| xlLessEqual | 8 | 小于或等于 |

## [XlFormatConditionType​](#xlformatconditiontype)

指定条件格式是基于单元格值还是基于表达式

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlCellValue | 1 | 单元格值 |
| xlExpression | 2 | 表达式 |
| xlColorScale | 3 | 色阶 |
| xlTop10 | 5 | 前 10 个值 |
| xlUniqueValues | 8 | 唯一值 |
| xlTextString | 9 | 文本字符串 |
| xlBlanksCondition | 10 | 空值条件 |
| xlTimePeriod | 11 | 时间段 |
| xlAboveAverageCondition | 12 | 高于平均值条件 |
| xlNoBlanksCondition | 13 | 无空值条件 |
| xlErrorsCondition | 16 | 错误条件 |
| xlNoErrorsCondition | 17 | 无错误条件 |

## [XlHAlign​](#xlhalign)

指定对象的水平对齐方式

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlHAlignRight | -4152 | 靠右 |
| xlHAlignLeft | -4131 | 靠左 |
| xlHAlignJustify | -4130 | 两端对齐 |
| xlHAlignDistributed | -4117 | 分散对齐 |
| xlHAlignCenter | -4108 | 居中 |
| xlHAlignGeneral | 1 | 按数据类型对齐 |
| xlHAlignFill | 5 | 填充 |
| xlHAlignCenterAcrossSelection | 7 | 跨列居中 |

## [XlInsertFormatOrigin​](#xlinsertformatorigin)

指定从何处复制插入单元格的格式

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlFormatFromLeftOrAbove | 0 | 从上方和/或左侧单元格复制格式 |
| xlFormatFromRightOrBelow | 1 | 从下方和/或右侧单元格复制格式 |

## [XlInsertShiftDirection​](#xlinsertshiftdirection)

指定插入时单元格的移动方向

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlShiftDown | -4121 | 向下移动单元格 |
| xlShiftToRight | -4161 | 向上移动单元格 |

## [XlLineStyle​](#xllinestyle)

指定边框的线条样式

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlLineStyleNone | -4142 | 无线 |
| xlDouble | -4119 | 双线 |
| xlDot | -4118 | 点式线 |
| xlDash | -4115 | 虚线 |
| xlContinuous | 1 | 实线 |
| xlDashDot | 4 | 点划相间线 |
| xlDashDotDot | 5 | 划线后跟两个点 |
| xlSlantDashDot | 13 | 倾斜的划线 |

## [XlPasteSpecialOperation​](#xlpastespecialoperation)

XlPaste 特殊操作

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlPasteSpecialOperationAdd | 0x1 | 复制的数据与目标单元格中的值相加 |
| xlPasteSpecialOperationDivide | 0x4 | 复制的数据除以目标单元格中的值 |
| xlPasteSpecialOperationMultiply | 0x3 | 复制的数据乘以目标单元格中的值 |
| xlPasteSpecialOperationNone | 0x0 | 粘贴操作中不执行任何计算 |
| xlPasteSpecialOperationSubtract | 0x2 | 复制的数据减去目标单元格中的值 |

## [XlPasteType​](#xlpastetype)

粘贴类型

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlPasteFormulas | 0x2 | 公式 |
| xlPasteAllExceptBorders | 0x4 | 无边框 |
| xlPasteColumnWidths | 0x5 | 保留源列宽 |
| xlPasteValues | 0x3 | 值 |
| xlPasteValuesAndNumberFormats | 0x7 | 值和数字格式 |
| xlPasteFormats | 0x8 | 格式 |
| xlPastePasteAll | 0x0 | 粘贴 |
| xlPasteComments | 0x9 | 批注 |
| xlPasteValidation | 0xa | 数据有效性验证 |

## [XlReferenceStyle​](#xlreferencestyle)

指定引用样式

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlR1C1 | -4150 | 使用 xlR1C1 返回 R1C1 样式的引用 |
| xlA1 | 1 | 默认值。 使用 xlA1 返回 A1 样式的引用 |

## [XlRowCol​](#xlrowcol)

指定对应于特定数据系列的数值是处于行中还是列中

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlRows | 1 | 数据系列在列中 |
| xlColumns | 2 | 数据系列在行中 |

## [XlSheetType​](#xlsheettype)

指定工作表类型

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlWorksheet | -4167 | 工作表 |
| xlDialogSheet | -4116 | 对话框工作表 |
| xlChart | -4109 | 图表 |
| xlExcel4MacroSheet | 3 | Excel 版本 4 宏工作表 |
| xlExcel4IntlMacroSheet | 4 | Excel 版本 4 国际宏工作表 |

## [XlTimePeriods​](#xltimeperiods)

指定时间段

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlToday | 0 | 今天 |
| xlYesterday | 1 | 昨天 |
| xlLast7Days | 2 | 过去 7 天 |
| xlThisWeek | 3 | 本周 |
| xlLastWeek | 4 | 上周 |
| xlLastMonth | 5 | 上月 |
| xlTomorrow | 6 | 明天 |
| xlNextWeek | 7 | 下周 |
| xlNextMonth | 8 | 下月 |
| xlThisMonth | 9 | 本月 |

## [XlUnderlineStyle​](#xlunderlinestyle)

指定应用于字体的下划线类型

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlUnderlineStyleDouble | -4119 | 粗双下划线 |
| xlUnderlineStyleDoubleAccounting | 5 | 紧靠在一起的两条细下划线 |
| xlUnderlineStyleNone | -4142 | 无下划线 |
| xlUnderlineStyleSingle | 2 | 单下划线 |
| xlUnderlineStyleSingleAccounting | 4 | 不支持 |

## [XlVAlign​](#xlvalign)

指定对象的垂直对齐方式

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlVAlignTop | -4160 | 向上 |
| xlVAlignJustify | -4130 | 调整使全行排满 |
| xlVAlignDistributed | -4117 | 一起 |
| xlVAlignCenter | -4108 | 居中 |
| xlVAlignBottom | -4107 | 向下 |

## [XlXLMMacroType​](#xlxlmmacrotype)

指定在工作表中，名称引用哪种宏，或名称是否引用宏

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlFunction | 1 | 自定义函数 |
| xlCommand | 2 | 自定义命令 |
| xlNotXLM | 3 | 非宏 |

## [XlAutoFilterOperator​](#xlautofilteroperator)

当前工作表中，指定筛选的类型，也用于指定用于关联两个筛选条件的操作符

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlAnd | 1 | Criteria1 和 Criteria2 的逻辑与 |
| xlOr | 2 | Criteria1 或 Criteria2 的逻辑或 |
| xlTop10Items | 3 | 显示最高值项 (在 Criteria1 中指定的项目数) |
| xlBottom10Items | 4 | 显示最低值项 (在 Criteria1 中指定的项目数) |
| xlTop10Percent | 5 | 显示最高值项 (Criteria1 中指定的百分比) |
| xlBottom10Percent | 6 | 显示最低值项 (在 Criteria1 中指定的百分比) |
| xlFilterValues | 7 | 筛选值 |
| xlFilterCellColor | 8 | 单元格颜色 |
| xlFilterFontColor | 9 | 字体颜色 |
| xlFilterIcon | 10 | 筛选图标 |
| xlFilterDynamic | 11 | 动态筛选 |

## [XlYesNoGuess​](#xlyesnoguess)

指定第一行是否包含标题。 对数据透视表进行排序时，不能使用该参数

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlGuess | 0 | 自动判断 Excel 是否有表头 |
| xlYes | 1 | 默认值。 应对整个区域进行排序 |
| xlNo | 2 | 不应对整个区域进行排序 |

## [XlSortOrientation​](#xlsortorientation)

指定排序方向

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlSortColumns | 1 | 按列排序 |
| xlSortRows | 2 | 按行排序，此值为默认值 |

## [XlSortMethod​](#xlsortmethod)

指定排序类型

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlPinYin | 1 | 按字符的汉语拼音顺序排序，此值为默认值 |
| xlStroke | 2 | 按每个字符的笔划数排序 |

## [XlSortDataOption​](#xlsortdataoption)

指定文本的排序方式

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlSortNormal | 0 | 分别对数字和文本数据进行排序，此值为默认值 |
| xlSortTextAsNumbers | 1 | 将文本作为数字型数据进行排序 |

## [XlSortOn​](#xlsorton)

指定数据的排序参数

| 字段 | 值 | 释义 |
| --- | --- | --- |
| SortOnCellColor | 1 | 单元格颜色 |
| SortOnFontColor | 2 | 字体颜色 |
| SortOnIcon | 3 | 图标 |
| SortOnValues | 0 | 值 |

## [XlSortOrder​](#xlsortorder)

为指定字段或范围指定排序顺序

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlAscending | 1 | 默认值，按升序对指定字段排序 |
| xlDescending | 2 | 按降序对指定字段排序 |

## [XlTextParsingType​](#xltextparsingtype)

指定要导入到查询表中的文本文件中的数据的列格式

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlDelimited | 1 | 默认值，指示文件由分隔符分隔 |
| xlFixedWidth | 2 | 指示将文件中的数据排列在固定宽度的列中 |

## [XlTextQualifier​](#xltextqualifier)

指定用于指定文本的分隔符

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlTextQualifierDoubleQuote | 1 | 双引号 (") |
| xlTextQualifierNone | -4142 | 无分隔符 |
| xlTextQualifierSingleQuote | 2 | 单引号 (') |

## [XlColumnDataType​](#xlcolumndatatype)

指定列的分列方式

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlDMYFormat | 4 | DMY 日期格式 |
| xlDYMFormat | 7 | DYM 日期格式 |
| xlEMDFormat | 10 | EMD 日期格式 |
| xlGeneralFormat | 1 | 常规 |
| xlMDYFormat | 3 | MDY 日期格式 |
| xlMYDFormat | 6 | MYD 日期格式 |
| xlSkipColumn | 9 | 列未分列 |
| xlTextFormat | 2 | 文本 |
| xlYDMFormat | 8 | YDM 日期格式 |
| xlYMDFormat | 5 | YMD 日期格式 |


#### 筛选(AutoFilter)

# [AutoFilter​](#autofilter)

代表对指定工作表的自动筛选，可通过此对象对工作表中的数据进行筛选，快速找到想要的值。可以组合一列或多列数据进行筛选。

使用筛选，不仅可以控制想要查看的内容，还可以控制想要排除的内容。在进行数据筛选时，如果一列或多列中的数值不能满足筛选条件，整行数据都会隐藏起来。可以对数值或文本值进行筛选，也可以对背景或文本应用颜色格式的单元格按颜色进行筛选。

开发者可以通过Filters属性获取由各个列筛选组成的集合。 使用Range属性可返回代表整个筛选区域的Range对象。

注意

若要为工作表创建 AutoFilter 对象，必须在工具栏手动开启筛选功能或者使用Range对象的AutoFilter方法为工作表上的某个区域启用自动筛选。

AutoFilter 对象的具体属性和方法请参阅下方的列表。

#### [属性列表​](#属性列表)

| 属性名 | 数据类型 | 简介 |
| --- | --- | --- |
| Filters | Filters | 筛选对象集合 |
| Range | Range | 筛选区域 |

#### [方法列表​](#方法列表)

| 方法名 | 返回类型 | 简介 |
| --- | --- | --- |
| ApplyFilter() | undefined | 应用筛选到当前工作表 |
| ShowAllData() | undefined | 清除所有筛选条件，显示所有数据 |

#### [应用示例​](#应用示例)

以下为您展示自动筛选对象的在工作表内的一些常见应用场景：

假设我们有如下的工作表，现在需要对它进行筛选操作，以显示我们期望的数据。

筛选出姓名为
金小獴
的数据
js
```js
// 获取自动筛选对象
const autoFilter = ActiveSheet.AutoFilter
// 获取姓名列的筛选对象
const filter2 = autoFilter.Filters.Item(2)
// 设置筛选类型为值筛选
filter2.Operator = Enum.XlAutoFilterOperator.xlFilterValues
// 设置筛选的值
filter2.Criteria1 = ['金小獴']
// 应用筛选
autoFilter.ApplyFilter()
```

筛选出语文成绩在前 20%的数据
js
```js
// 获取自动筛选对象
const autoFilter = ActiveSheet.AutoFilter
// 获取语文列的筛选对象
const filterItem = autoFilter.Filters.Item(5)
// 设置筛选类型为头部百分比筛选
filterItem.Operator = Enum.XlAutoFilterOperator.xlTop10Percent
// 设置筛选条件为20%
filterItem.Criteria1 = '20'
// 应用筛选
autoFilter.ApplyFilter()
```

筛选出数学前十名的数据
js
```js
// 获取自动筛选对象
const autoFilter = ActiveSheet.AutoFilter
// 获取数学列的筛选对象
const filterItem = autoFilter.Filters.Item(6)
// 设置筛选类型为头部筛选
filterItem.Operator = Enum.XlAutoFilterOperator.xlTop10Items
// 设置筛选条件为10，即前十名
filterItem.Criteria1 = '10'
// 应用筛选
autoFilter.ApplyFilter()
```

将姓名列中单元格为红色的筛选出来
js
```js
// 获取自动筛选对象
const autoFilter = ActiveSheet.AutoFilter
// 获取姓名列的筛选对象
const filterItem = autoFilter.Filters.Item(2)
// 设置筛选类型为单元格颜色筛选
filterItem.Operator = Enum.XlAutoFilterOperator.xlFilterCellColor
// 设置筛选条件为红色
filterItem.Criteria1 = '#FF0000'
// 应用筛选
autoFilter.ApplyFilter()
```

筛选出语文分介于100至200之间的数据
js
```js
// 获取自动筛选对象
const autoFilter = ActiveSheet.AutoFilter
// 获取语文列的筛选对象
const filterItem = autoFilter.Filters.Item(5)
// 设置筛选类型为条件与筛选
filterItem.Operator = Enum.XlAutoFilterOperator.xlAnd
// 设置筛选条件
filterItem.Criteria1= '>=100'
filterItem.Criteria2= '<=200'
// 应用筛选
autoFilter.ApplyFilter()
```

筛选第姓名包含'张三'或者开头等于'李四'的数据
js
```js
// 获取自动筛选对象
const autoFilter = ActiveSheet.AutoFilter
// 获取姓名列的筛选对象
const filterItem = autoFilter.Filters.Item(2)
// 设置筛选类型为条件或筛选
filterItem.Operator = Enum.XlAutoFilterOperator.xlOr
// 设置筛选条件
filterItem.Criteria1= '=*张三*'
filterItem.Criteria2= '=李四*'
// 应用筛选
autoFilter.ApplyFilter()
```

## [Filters​](#filters)

筛选对象集合

该集合内包含所有可供操作的数据列筛选对象

#### [数据类型​](#数据类型)

Filters - Filter 集合

### [Count​](#count)

当前所有筛选器对象的数量

#### [数据类型​](#数据类型-1)

Number - 对应当前工作表所有筛选对象的数量

#### [示例​](#示例)

js
```js
// 获取自动筛选对象
const autoFilter = Application.ActiveSheet.AutoFilter
// 获取筛选对象集合
const filters = autoFilter.Filters
// 获取当前工作表所有筛选对象的数量
const count = filters.Count
console.log(count)
```

### [Each()​](#each)

遍历所有 Filters 并执行回调函数

#### [参数​](#参数)

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| callback | Function | null | 是 | 类似 JS 数组的 forEach |

#### [示例​](#示例-1)

js
```js
// 获取自动筛选对象
const autoFilter = Application.ActiveSheet.AutoFilter
// 获取筛选对象集合
const filters = autoFilter.Filters
// 遍历筛选集合，执行回调函数
filters.Each(item => {
  console.log(item.Operator)
})
```

### [Item()​](#item)

根据索引选择对应的筛选对象

#### [参数​](#参数-1)

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| index | String/Number |  | 是 | 对应工作表内的实际列号，索引从 1 开始 |

#### [返回类型​](#返回类型)

Filter - 返回对应索引的筛选对象

#### [示例​](#示例-2)

js
```js
// 获取自动筛选对象
const autoFilter = Application.ActiveSheet.AutoFilter
// 获取筛选对象集合
const filters = autoFilter.Filters
// 获取第一个筛选对象
const filter1 = filters.Item(1)
```

### [Item().Operator​](#item-operator)

指定筛选类型，可使用的筛选类型请参照枚举值XlAutoFilterOperator

#### [数据类型​](#数据类型-2)

XlAutoFilterOperator- 枚举值 XlAutoFilterOperator

#### [示例​](#示例-3)

js
```js
// 获取自动筛选对象
const autoFilter = Application.ActiveSheet.AutoFilter
// 获取第一个筛选对象
const filter1 = autoFilter.Filters.Item(1)
// 设置筛选类型为单元格颜色
filter1.Operator = Enum.XlAutoFilterOperator.xlFilterCellColor
```

### [Item().Criteria1​](#item-criteria1)

指定判断条件，使用“=”查找空字段，或者使用“<>”查找非空字段。如果忽略该参数，那么判断是全部。如果参数 Operator 是 xlTop10Items，那么参数 Criterial1 指定项目的数量

#### [数据类型​](#数据类型-3)

Variant - 根据指定的筛选类型确定

### [Item().Criteria2​](#item-criteria2)

第二个判断条件。与 Criteria1 和 Operator 一起组合成复合筛选条件。 也用作日期字段的单一条件（按日、月或年筛选）。 后跟一个数组，该数组用于描述筛选 Array(Level, Date)。 其中，Level 为 0-2（年、月、日），Date 为筛选期内的一个有效日期

#### [数据类型​](#数据类型-4)

Variant - 根据指定的筛选类型确定

## [Range​](#range)

自动筛选的区域范围

#### [数据类型​](#数据类型-5)

Range- 区域对象

#### [示例​](#示例-4)

js
```js
// 获取自动筛选对象
const autoFilter = Application.ActiveSheet.AutoFilter
// 获取自动筛选的区域
const range = autoFilter.Range
// 打印该区域内的单元格数量
console.log(range.Count)
```

## [ApplyFilter()​](#applyfilter)

将自动筛选器应用于区域，在设置好筛选类型和筛选条件后，调用此方法来应用筛选

#### [示例​](#示例-5)

js
```js
// 获取自动筛选对象
const autoFilter = Application.ActiveSheet.AutoFilter
// 设置第一列的筛选类型为单元格颜色
autoFilter.Filters.Item(1).Operator =
  Enum.XlAutoFilterOperator.xlFilterCellColor
// 设置第一列的筛选条件是#44546A
autoFilter.Filters.Item(1).Criteria1 = '#44546A'
// 应用筛选到区域
autoFilter.ApplyFilter()
```

## [ShowAllData()​](#showalldata)

清除所有筛选条件，显示所有数据

#### [示例​](#示例-6)

js
```js
// 获取自动筛选对象
const autoFilter = Application.ActiveSheet.AutoFilter

// 清除所有筛选条件
autoFilter.ShowAllData()
```


#### 行记录(Record)

# [Record​](#record)

行记录

### [方法列表​](#方法列表)

| 方法名 | 返回类型 | 简介 |
| --- | --- | --- |
| GetRecords() | Array | 获取行记录（多条） |
| GetRecord() | Object | 获取行记录（单条） |
| DeleteRecords() | Array | 删除行记录 |
| UpdateRecords() | Array | 更新行记录 |
| CreateRecords() | Array | 创建行记录 |
| GetAttachmentURL() | String | 获取上传附件或图片的URL |

## [GetRecords()​](#getrecords)

获取行记录（多条）

注意

每次请求最多返回100条，数据量大的时候请使用分页查询

### [参数​](#参数)

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| ViewId | String |  | 否 | 填写后将从被指定的视图获取该用户所见到的记录；若不填写，则从工作表获取记录 |
| PageSize | Number | 100 | 否 | 存在分页时，指定本次查询的起始记录（含）。若不填写或填写为空字符串，则从第一条记录开始获取。当前最大值：1000 |
| Offset | Number |  | 否 | 分页查询时，将返回一个offset值，指向下一页的第一条记录，供后续查询。查询到最后一页或第maxRecords条记录时，返回数据将不再包含offset值 |
| MaxRecords | Number |  | 否 | 指定要获取的“前maxRecords条记录”，若不填写，则默认返回全部记录 |
| Fields | Array |  | 否 | 字段类型 |
| Filter | Object |  | 否 | 详细说明见附录三 |

### [返回值​](#返回值)

Object - 获取表的所有记录

| 属性 | 数据类型 | 说明 |
| --- | --- | --- |
| Offset | String | 如果分页的话， 则会返回此字段信息;分页截止 id， 下次请求携带会继续分页请求信息 |
| Records | Array[Object] | 记录集合 |

#### [记录集合​](#记录集合)

| 属性 | 数据类型 | 说明 |
| --- | --- | --- |
| id | String | 记录Id |
| Fields | Object | 更新的字段信息，包含字段Id，字段name,格式说明见附录 |

### [示例​](#示例)

javascript
```javascript
const sheet = Application.ActiveSheet
// 分页查询例子
function fetchAllRecords() {
  const view = sheet.Selection.GetActiveView()
  let all = []
  let offset = null;

  while (all.length === 0 || offset) {
    let records = sheet.Record.GetRecords({
      ViewId: view.viewId,
      Offset: offset,
    })
    offset = records.offset
    all = all.concat(records.records)
  }
  console.log(all.length)
  return all
}

fetchAllRecords()
```

## [GetRecord()​](#getrecord)

获取行记录（单条）

### [参数​](#参数-1)

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| RecordId | String |  | 是 | 表中指定获取的记录id |

### [返回值​](#返回值-1)

Object - 获取表的指定的单条记录

| 属性 | 数据类型 | 说明 |
| --- | --- | --- |
| id | String | 记录Id |
| Fields | Object | 更新的字段信息，包含字段Id，字段name,格式说明见附录 |

### [示例​](#示例-1)

javascript
```javascript
const sheet = Application.ActiveSheet
const record = sheet.Record.GetRecord({  RecordId: 'Bz' })
console.log(record)
// 打印结果：
//  {"fields":{"日期":"2023/02/21"},"id":"Bz"}
```

## [DeleteRecords()​](#deleterecords)

删除行记录

### [参数​](#参数-2)

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| RecordIds | Array |  | 是 | 表中需要删除的记录id |

### [返回值​](#返回值-2)

Array - 返回删除的表id以及删除是否成功信息

| 属性 | 数据类型 | 说明 |
| --- | --- | --- |
| id | String | 记录Id |
| deleted | Boolean | 是否删除成功 “true”表示删除成功，“false”表示删除失败 |

### [示例​](#示例-2)

javascript
```javascript
const sheet = Application.ActiveSheet
const result = sheet.Record.DeleteRecords({ 
    RecordIds: ['J', 'P', 'Q'] 
})
console.log(resutlt)
// 打印结果：
// [{"deleted":true,"id":"P"},{"deleted":false,"id":"Q"}]
```

## [UpdateRecords()​](#updaterecords)

更新行记录

### [参数​](#参数-3)

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| Records | Array[Object] |  | 是 | 行记录集合 |

#### [行记录集合：​](#行记录集合)

| 属性 | 数据类型 | 说明 |
| --- | --- | --- |
| id | String | 记录Id |
| Fields | Object | 更新的字段信息，包含字段Id，字段name,格式说明见附录 |

### [返回值​](#返回值-3)

Array - 表的已更新的所有记录

| 属性 | 数据类型 | 说明 |
| --- | --- | --- |
| id | String | 记录Id |
| Fields | Object | 更新的字段信息，包含字段Id，字段name,格式说明见附录 |

### [示例​](#示例-3)

javascript
```javascript
const sheet = Application.ActiveSheet
const records = sheet.Record.UpdateRecords({
        Records: [{
            id: 'A',
            fields: {
                 邮箱: 'demo@qq.com',
                 多选: ['1', '2'],
                 "记录关联": {
                    "recordIds": ["I", "K"] 
                 }
            }
        }],
    })
```

## [CreateRecords()​](#createrecords)

创建行记录

### [参数​](#参数-4)

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| Records | Array[Object] |  | 是 | 行记录集合 |

#### [行记录集合：​](#行记录集合-1)

| 属性 | 数据类型 | 说明 |
| --- | --- | --- |
| id | String | 记录Id |
| Fields | Object | 更新的字段信息，包含字段Id，字段name,格式说明见附录 |

### [返回值​](#返回值-4)

Array - 表的已更新的所有记录

| 属性 | 数据类型 | 说明 |
| --- | --- | --- |
| id | String | 记录Id |
| Fields | Object | 更新的字段信息，包含字段Id，字段name,格式说明见附录 |

### [示例​](#示例-4)

javascript
```javascript
const sheet = Application.ActiveSheet
// 创建邮箱和多选
const records = sheet.Record.CreateRecords({
      Records: [{
          fields: {
               邮箱: 'demo@qq.com',
               多选: ['1', '2'],
          }
      }, {
          fields: {
               邮箱: 'demo@qq.com',
               多选: ['1', '2'],
          }
      }],
  })

// 创建联系人
const records = sheet.Record.CreateRecords({
  Records: [
    {  fields: { '联系人': [{ name: 'yourname', nickName: 'yourname', id: '88888888', avatar_url: 'https://avatar.qwps.cn/avatar/5b2t57-U' }] } },
  ],
});
```

## [GetAttachmentURL()​](#getattachmenturl)

获取上传附件或图片的URL

### [参数​](#参数-5)

注意

必须至少传入1个参数Attachment或者传入2个参数UploadId和Source

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| Attachment | String |  | 否 | 附件 |
| UploadId | String |  | 否 | 上传文件id |
| Source | String |  | 否 | source参数必须为"upload_ks3"（本地上传）或"cloud"（云上传） |

### [返回值​](#返回值-5)

String - 为获取上传附件或图片的URL，打开该URL可进行附件或图片下载

### [示例​](#示例-5)

javascript
```javascript
const sheet = Application.ActiveSheet
const resultURL = sheet.Record.GetAttachmentURL({
    Attachment: "IKWRCBAAKA|upload_ks3|image/png|QQ图片20230214165215.png|12070||549*106",
      })

//or

const resultURL = sheet.Record.GetAttachmentURL({
    UploadId: "IKWRCBAAKA",
    Source: "upload_ks3"
      })
```


#### 表格实例(Application)

# [Application​](#application)

文档操作的顶级对象，对文档进行相关操作，都是间接或直接操作该对象。

Application 是一个文件的顶级对象，新打开一个文件返回的也是 Application。

而在脚本中的Application则是指当前文件的顶级对象，有且只有一个。

Application 对象的具体属性和方法请参阅下方的列表。

#### [属性列表​](#属性列表)

| 属性 | 数据类型 | 简介 |
| --- | --- | --- |
| ActiveSheet | Sheet | 当前的活动工作表/数据表 |
| Sheets | Sheets | 当前文件的所有工作表/数据表 |
| FileInfo | Object | 当前文档的信息 |
| UserInfo | Object | 当前文档的用户信息 |
| Enum | Enum | 所有的枚举类型 |

#### [方法列表​](#方法列表)

| 方法 | 返回类型 | 简介 |
| --- | --- | --- |
| Sheets(name) | Sheet | 获取名称为 name 的工作表/数据表 |

## [ActiveSheet​](#activesheet)

当前活动工作表/数据表，可以通过 Sheet.Activate()来切换活动工作表/数据表。该属性返回Sheet对象,能利用该属性操作当前活动工作表/数据表。

运行脚本的环境是独立在服务器的，因此脚本运行环境的 ActiveSheet 与用户环境的 ActiveSheet 不一定相同。

具体规则是：

1.运行脚本时会把脚本运行环境的 ActiveSheet 切换为用户环境当前的 ActiveSheet。

2.当脚本通过函数切换脚本运行环境的 ActiveSheet 时，用户环境的 ActiveSheet 不会同步切换。

#### [数据类型​](#数据类型)

Sheet- 当前活动工作表/数据表

#### [示例​](#示例)

js
```js
console.log(Application.ActiveSheet.Name) // 数据表2

// 切换到名称为数据表2的数据表
Application.Sheets.Item('数据表2').Activate()
console.log(Application.ActiveSheet.Name) // 数据表2
```

## [Sheets​](#sheets)

获取当前文件能操作的所有 Sheet，返回一个Sheets对象。

#### [数据类型​](#数据类型-1)

Sheets

#### [示例​](#示例-1)

js
```js
// 工作簿（Workbook）中所有工作表/数据表（Sheet）的集合,下面两种写法是一样的
let sheets = Application.ActiveWorkbook.Sheets
sheets = Application.Sheets

// 打印所有工作表/数据表的名称
for (let i = 1; i <= sheets.Count; i++) {
  console.log(sheets.Item(i).Name)
}
```

### [Sheets.Count​](#sheets-count)

工作表/数据表数量

#### [数据类型​](#数据类型-2)

Number - 对应工作簿的工作表/数据表数量

#### [示例​](#示例-2)

js
```js
// 下面两种写法是一样的
let sheets = Application.ActiveWorkbook.Sheets
sheets = Application.Sheets

// 打印所有工作表/数据表的名称
console.log(sheets.Count) //1
```

### [Sheets.DefaultNewSheetName​](#sheets-defaultnewsheetname)

默认新工作表名

#### [返回类型​](#返回类型)

String - 新建工作表时若没有指定名称，可用这个名称作为新建工作表名称

#### [示例​](#示例-3)

js
```js
const defaultName = Application.Sheets.DefaultNewSheetName
// 工作表对象
Application.Sheets.Add(
  null,
  Application.ActiveSheet.Name,
  1,
  Application.Enum.XlSheetType.xlWorksheet,
  defaultName
)
```

### [Sheets.Add()​](#sheets-add)

新增工作表，如果 Before 和 After 都存在，以 Before 为准

#### [参数​](#参数)

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| Before | String/Number |  | 否 | After 空时，必填，为当前已有单元格的 index 或者名称，新建的工作表将置于此工作表之前 |
| After | String/Number |  | 否 | Before 空时，必填，为当前已有单元格的 index 或者名称，新建的工作表将置于此工作表之后 |
| Count | Number | 1 | 否 | 要添加的工作表数。默认值为选定工作表的数量 |
| Type | Enum |  | 否 | 指定工作表类型，详细可见Enum.XlSheetType |
| Name | Name |  | 否 | 指定工作表名称 |

#### [示例​](#示例-4)

js
```js
// 添加工作表
Application.Sheets.Add(
  null,
  Application.ActiveSheet.Name,
  1,
  Application.Enum.XlSheetType.xlWorksheet,
  '新工作表'
)
```

### [Sheets.Item()​](#sheets-item)

根据名称或索引选择 Sheet

#### [参数​](#参数-1)

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| index | String/Number |  | 是 | 所选的 sheet 的名称/索引 |

#### [返回类型​](#返回类型-1)

Sheet- 对应名称的工作表/数据表

#### [示例​](#示例-5)

js
```js
// 切换名称为"Sheet2"的工作表
Application.Sheets.Item('Sheet2').Activate()

// 切换索引为1的工作表
Application.Sheets.Item(1).Activate()
```

### [Sheets.Each()​](#sheets-each)

遍历所有 sheet 并执行回调函数

#### [参数​](#参数-2)

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| callback | Function | null | 是 | 类似 JS 数组的 forEach |

#### [示例​](#示例-6)

js
```js
// 打印所有工作表/数据表的名称
Application.Sheets.Each(function (item) {
  console.log(item.Name) //Sheet1 Sheet2
})
```

## [FileInfo​](#fileinfo)

返回当前文件的基本信息。

#### [数据类型​](#数据类型-3)

Object - 当前文件的信息

| 名称 | 类型 | 说明 |
| --- | --- | --- |
| id | string | 文件 ID |
| name | string | 文件名 |
| officeType | string | 文档类型 |
| creator | CreatorObject | 文档创建者信息 |
| size | number | 文件大小 |
| groupId | string | 文件的群组 ID |
| docType | number | 文档类型（数字形式） |

#### [CreatorObject 对象信息​](#creatorobject)

| 名称 | 类型 | 说明 |
| --- | --- | --- |
| id | string | 创建者 ID |
| name | string | 创建者名称 |
| avatar_url | string | 创建者头像 |
| logined | boolean | 是否已登录 |
| attrs | Object | 属性对象 |
| real_id | string | 真实 ID |

#### [示例​](#示例-7)

javascript
```javascript
// 打印文件信息
console.log(Application.FileInfo)
/*{
 "id": "<open_id>",
 ...
}*/
```

## [UserInfo​](#userinfo)

返回当前文件的用户信息。

#### [数据类型​](#数据类型-4)

Object- 当前文件的用户信息

| 名称 | 类型 | 说明 |
| --- | --- | --- |
| id | string | 用户 ID |
| name | string | 用户名称 |

#### [示例​](#示例-8)

javascript
```javascript
// 打印用户信息
console.log(Application.UserInfo)
```

## [Enum​](#enum)

枚举类型，存放在 Application 下。

可以通过 Application.Enum 使用

#### [数据类型​](#数据类型-5)

Enum- 所有的枚举类型

#### [示例​](#示例-9)

js
```js
// 打印工作表/数据表的类型枚举
console.log(Application.Enum.XlSheetType)
//{"xlChart":-4109,"xlDialogSheet":-4116,"xlExcel4IntlMacroSheet":4,"xlExcel4MacroSheet":3,"xlWorksheet":-4167}
```

## [Sheets()​](#sheets-1)

作为函数使用，代替 Sheets.Item()，返回一个Sheet对象。

#### [参数​](#参数-3)

| 名称 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| name | string | 是 | 工作表/数据表的名称 |

#### [返回类型​](#返回类型-2)

Sheet- 对应名称的工作表/数据表 Sheet 对象

#### [示例​](#示例-10)

js
```js
console.log(Application.Sheets.Count) // 1

// 以下两种写法效果是一样的
console.log(Application.Sheets('Sheet2').Range('A1').Text) // Sheet2的A1单元格的内容
console.log(Application.Sheets.Item('Sheet2').Range('A1').Text) // Sheet2的A1单元格的内容
```


#### 超链接(Hyperlink)

# [Hyperlink​](#hyperlink)

单个超链接对象，有转跳地址和显示文本两个属性，两者可以不相等。

Hyperlink 对象的具体属性和方法请参阅下方的列表。

### [属性列表​](#属性列表)

| 属性名 | 数据类型 | 简介 |
| --- | --- | --- |
| Address | String | 超链接转跳的地址 |
| TextToDisplay | String | 超链接显示的文本 |

## [Address​](#address)

设置/获取执行超链接的地址

### [数据类型​](#数据类型)

String - 超链接转跳的地址

### [示例​](#示例)

js
```js
//获取超链接对象
const hyperlink = Application.ActiveSheet.Hyperlinks.Item(1)
// 打印超链接的地址
console.log(hyperlink.Address) // https://www.kdocs.cn
hyperlink.Address = "https://www.kdocs.cn"
```

## [TextToDisplay​](#texttodisplay)

设置/获取执行超链接的文本

### [数据类型​](#数据类型-1)

String - 超链接显示的文本

### [示例​](#示例-1)

js
```js
//获取超链接对象
const hyperlink = Application.ActiveSheet.Hyperlinks.Item(1)
// 打印超链接的地址
console.log(hyperlink.TextToDisplay) // 金山文档官网
hyperlink.TextToDisplay = "金山文档"
```


#### 边框(Border)

# [Border​](#border)

边框对象，Borders 集合里的某一边框

Border 对象的具体属性和方法请参阅下方的列表。

### [属性列表​](#属性列表)

| 属性名 | 数据类型 | 简介 |
| --- | --- | --- |
| Color | Number | 边框的颜色 |
| Weight | Enum.XlBorderWeight | 边框的粗细 |
| LineStyle | Enum.XlLineStyle | 边框的线条样式 |

## [Color​](#color)

边框颜色

注意：获取边框颜色时，需要指定具体的边框，即枚举值Enum.XlBordersIndex不能是 xlAll、xlOutside、xlInside 等

### [数据类型​](#数据类型)

Number - 边框颜色

### [示例​](#示例)

js
```js
// 区域对象
const borders = Application.Range('A1').Borders

// 底部边框对象
let border = borders.Item(Application.Enum.XlBordersIndex.xlEdgeBottom)

// 打印边框颜色
console.log(border.Color) // {"private":{"isAuto":false,"rgbValue":"#000000","tint":0,"type":"ectAUTO"},"rgbValue":"#000000"}

// 将整个边框颜色设置成绿色
border = borders.Item(Application.Enum.XlBordersIndex.xlOutside)
border.Color = '#00FF00'
```

## [Weight​](#weight)

边框的粗细

### [数据类型​](#数据类型-1)

Enum.XlBorderWeight- 边框的粗细

该对象只能设置值，无法读取值

### [示例​](#示例-1)

js
```js
// 区域对象
let borders = Application.Range('A3').Borders

// 单个边框对象
let border = borders.Item(Application.Enum.XlBordersIndex.xlOutside)

// 设置边框的粗细
border.Weight = Application.Enum.XlBorderWeight.xlThick
```

## [LineStyle​](#linestyle)

边框的线条样式

### [数据类型​](#数据类型-2)

Enum.XlLineStyle- 边框的线条样式

该对象只能设置值，无法读取值

### [示例​](#示例-2)

js
```js
// 区域对象
let borders = Application.Range('A1').Borders

// 底部边框对象
let border = borders.Item(Application.Enum.XlBordersIndex.xlEdgeBottom)

// 将整个边框样式设置成双线
border = borders.Item(Application.Enum.XlBordersIndex.xlOutside)
border.LineStyle = Application.Enum.XlLineStyle.xlDouble
```


#### 附录

# [附录​](#附录)

## [附录 1：数据表字段类型说明​](#附录-1-数据表字段类型说明)

| 字段类型 | Type | 创建字段格式 | 设置字段值传入形式 | 读取字段值传出形式 |
| --- | --- | --- | --- | --- |
| 多行文本 | MultiLineText | 无特殊要求 | 字符串/ 无特殊格式要求 | 字符串 |
| 日期 | Date | 无特殊要求 | 字符串/yyyy/mm/dd | 字符串 |
| 时间 | Time | 无特殊要求 | 字符串/hh:mm:ss | 字符串 |
| 数值 | Number | 无特殊要求 | 数值 / 无格式 | 数值 |
| 货币 | Currency | 无特殊要求 | 数值 / 无格式 | 数值 |
| 百分比 | Percentage | 无特殊要求 | 数值 / 无格式 | 数值 |
| 身份证 | ID | 无特殊要求 | 字符串 / 符合身份证规则 | 字符串 |
| 电话 | Phone | 无特殊要求 | 字符串 / 符合电话规则 | 字符串 |
| 电子邮箱 | Email | 无特殊要求 | 字符串 / 符合邮箱规则 | 字符串 |
| 超链接 | Url | 可以额外传入一个参数。displayText：指定超链接显示文本。{"name":"超链接","type":"Url","displayText":"跳转"} | 字符串 / 符合 Url 规 | 字符串 |
| 复选框 | Checkbox | 无特殊要求 | true / false | 布尔 |
| 单选项 | SingleSelect | 需要额外传入选项值，至少一个。{"name": "单选项","type": "SingleSelect","items": [{ "value": "item1" }]} | 字符串 / 匹配选项内容 | 字符串 |
| 多选项 | MultipleSelect | 需要额外传入选项值，至少一个。{"name": "单选项","type": "SingleSelect","items": [{ "value": "item1" }, { "value": "item2" }]} | 字符串数组 / 匹配选项内容 | 字符串数组 |
| 等级 | Rating | 需要额外传入一个最大等级, 最大等级大于 0 小于等于 5。{"name": "等级","type": "Rating","max": 5} | 数值 / 大于 0 并且 小于 最大等级 | 数值 |
| 进度条 | Complete | 无特殊要求 | 数值 / 大于等于 0 并且 小于 100 | 字符串 |
| 联系人 | Contact | 需要额外传入两个参数：multipleContacts:<bool>是否支持多个联系人noticeNewContact:<bool>是否通知联系人。{"name": "联系人","type": "Contact","multipleContacts": false,"noticeNewContact": false} | 不支持设值 | 对象 |
| 附件 | Attachment | 无特殊要求 | 不支持设值 | 对象 |
| 关联 | Link | 需要额外传入二个参数：linkSheet: 关联表 IDmultipleLinks: 是否关联多条记录{"name": "联系人","type": "Link","multipleContacts": false,"noticeNewContact": false} | 对应关联表的行记录数组 |  |
| 富文本 | Note | 无特殊要求 | 不支持设值 | 对象 |
| 编号 | AutoNumber | 无特殊要求 | 不支持设值 | 数值 |
| 创建者 | CreatedBy | 无特殊要求 | 不支持设值 | 对象 |
| 创建时间 | CreatedTime | 无特殊要求 | 不支持设值 | 字符串 |
| 公式 | Formula | 无特殊要求 | 不支持设值 | 根据公式的值类型 |
| 引用 | Lookup | 无特殊要求 | 不支持设值 | 与被引用形式相同 |

## [附录 2：数据表视图类型说明​](#附录-2-数据表视图类型说明)

| 视图类型 | 说明 |
| --- | --- |
| Grid | 表格视图 |
| Kanban | 看板视图 |
| Gallery | 画册视图 |
| Form | 表单视图 |
| Gantt | 甘特视图 |

## [附录 3：筛选条件说明​](#附录-3-筛选条件说明)

筛选条件用来对行记录进行筛选，由两部分构成：mode为筛选条件关系；creteria为具体筛选条件（fileds op values）。

json
```json
{
  "mode": "AND", // 选填。表示各筛选条件之间的逻辑关系。只能是"AND"或"OR"。缺省值为"AND"
  "criteria": [
    // filter结构体内必填。包含筛选条件的数组。每个字段上只能有一个筛选条件
    {
      "field": "名称", // 必填。根据 preferId 与否，需要填入字段名或字段id
      "op": "Intersected", // 必填。表示具体的筛选规则，见下
      "values": [
        // 必填。表示筛选规则中的值。数组形式。
        "数据表", // 值为字符串时表示文本匹配
        "12345"
      ]
    },
    {
      "field": "数量",
      "op": "Greater",
      "values": ["1"]
    }
  ]
}
```

| 筛选条件 | 参数说明 |
| --- | --- |
| Equals | 等于 |
| NotEqu | 不等于 |
| Greater | 大于 |
| GreaterEqu | 大等于 |
| Less | 小于 |
| LessEqu | 小等于 |
| GreaterEquAndLessEqu | 介于（取等） |
| LessOrGreater | 介于（不取等） |
| BeginWith | 开头是 |
| EndWith | 结尾是 |
| Contains | 包含 |
| NotContains | 不包含 |
| Intersected | 指定值 |
| Empty | 为空 |
| NotEmpty | 不为空 |

各筛选规则独立地限制了 values 数组内最多允许填写的元素数，当 values 内元素数超过阈值时，该筛选规则将失效。

为空、不为空不允许填写元素；介于允许最多填写 2 个元素；指定值允许填写 65535 个元素；其他规则允许最多填写 1 个元素。

注意

filter 不是结构体，当 criteria 未指定 field、op/values 参数填写不合法、values 填写过多参数及其他可能导致筛选规则失效等情形，整个请求将直接失败。

目前还支持对日期进行动态筛选，此时 values[]内的元素需以结构体的形式给出：

json
```json
{
  "mode": "AND",
  "criteria": [
    {
      "field": "日期",
      "op": "Equals",
      "values": [
        {
          "dynamicType": "lastMonth",
          "type": "DynamicSimple"
        }
      ]
    }
  ]
}
```

提示

上述示例对应的筛选条件为等于上一个月。

要使用日期动态筛选，values[]内的结构体需要指定type为DynamicSimple，当op为Equals时，dynamicType可以为如下的值（大小写不敏感）。

| 字段 | 说明 |
| --- | --- |
| today | 今天 |
| yesterday | 昨天 |
| tomorrow | 明天 |
| last7Days | 最近 7 天 |
| last30Days | 最近 30 天 |
| thisWeek | 本周 |
| lastWeek | 上周 |
| nextWeek | 下周 |
| thisMonth | 本月 |
| lastMonth | 上月 |
| nextMonth | 次月 |

提示

当op为greater或less时，dynamicType只能是昨天、今天或明天。


### 数据表

#### 字段(Field)

# [Field​](#field)

字段操作

### [方法列表​](#方法列表)

| 方法名 | 返回类型 | 简介 |
| --- | --- | --- |
| GetFields() | Array | 获取字段信息 |
| CreateFields() | Array | 创建字段 |
| DeleteFields() | Array | 删除字段 |
| UpdateFields() | Array | 更新字段 |

## [GetFields()​](#getfields)

获取字段信息

### [返回值​](#返回值)

Array - 返回获取的表所有字段信息

| 属性 | 数据类型 | 说明 |
| --- | --- | --- |
| id | String | 字段Id |
| name | String | 字段名称 |
| type | String | 字段类型 |

### [示例​](#示例)

javascript
```javascript
const sheet = Application.ActiveSheet
// 获取的表所有字段信息
const fields = sheet.Field.GetFields()
console.log(fields)
// 打印结果：
// [
//  {"id":"Ce","name":"名称","type":"MultiLineText"},
//  {"id":"Cf","name":"数量","type":"Number"},
// ]
```

## [CreateFields()​](#createfields)

创建字段

### [参数​](#参数)

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| Fields | Array |  | 是 | 表的字段信息,格式说明见附录 |

### [返回值​](#返回值-1)

Array - 返回已创建的表所有字段信息

| 属性 | 数据类型 | 说明 |
| --- | --- | --- |
| id | String | 字段Id |
| name | String | 字段名称 |
| type | String | 字段类型 |

### [示例​](#示例-1)

javascript
```javascript
const sheet = Application.ActiveSheet
const field =  sheet.Field.CreateFields({ 
    Fields: [ 
        { name: '等级',  type: 'Rating', max: 5 }
    ] 
})
console.log(field)
// 打印结果：
// [{"id":"LZ","name":"等级","type":"Rating"}]
```

## [DeleteFields()​](#deletefields)

删除字段

### [参数​](#参数-1)

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| Fields | Array |  | 是 | 需要删除的字段Id |

### [返回值​](#返回值-2)

Array - 返回删除的表id以及删除是否成功信息

| 属性 | 数据类型 | 说明 |
| --- | --- | --- |
| id | String | 字段Id |
| deleted | Boolean | 是否删除成功 |

### [示例​](#示例-2)

javascript
```javascript
const sheet = Application.ActiveSheet
// 删除字段
const resutlt = sheet.Field.DeleteFields({ FieldIds: ['P', 'Q'] })
console.log(resutlt)
// 打印结果：
// [{"deleted":false,"id":"P"},{"deleted":false,"id":"Q"}]
```

## [UpdateFields()​](#updatefields)

更新字段

### [参数​](#参数-2)

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| Fields | Array |  | 是 | 更新的字段信息，包含字段Id，字段name,格式说明见附录 |

### [返回值​](#返回值-3)

Array - 返回已更新的字段信息

| 属性 | 数据类型 | 说明 |
| --- | --- | --- |
| id | String | 字段Id |
| name | String | 字段名称 |
| type | String | 字段类型 |

### [示例​](#示例-3)

javascript
```javascript
const sheet = Application.ActiveSheet
// 修改字段名称
sheet.Field.UpdateFields({ 
    Fields: [{ id: 'LG', name: '跳转' }]
})
```


#### 数据表(Sheet)

# [Sheet​](#sheet)

工作簿（Workbook）中单个数据表(Sheet)对象

Sheet 对象的具体属性和方法请参阅下方的列表。

#### [属性列表​](#属性列表)

| 属性名 | 数据类型 | 简介 |
| --- | --- | --- |
| Id | String | 该数据表的 Id |
| Name | String | 该数据表的名称 |
| Index | Number | 该数据表在所有表的索引值 |
| Visible | Boolean | 该数据表是否可见 |
| Type | String | 该数据表的类型 |
| Field | Field | 该数据表的字段 |
| Record | Record | 该数据表的行记录 |

#### [方法列表​](#方法列表)

| 方法名 | 返回类型 | 简介 |
| --- | --- | --- |
| Activate() | undefined | 切换(激活)数据表 |
| Move() | undefined | 移动数据表 |
| Delete() | undefined | 删除数据表 |
| IsDBSheet() | Boolean | 是否为数据表 |

## [Id​](#id)

获取数据表 Id

#### [数据类型​](#数据类型)

String - 数据表 Id

#### [示例​](#示例)

js
```js
const sheet = Application.ActiveSheet
// 打印当前活动数据表的id
console.log(sheet.Id)
```

## [Name​](#name)

设置/获取 数据表名称

#### [数据类型​](#数据类型-1)

String - 该数据表在所有数据表的名称

#### [示例​](#示例-1)

js
```js
const sheet = Application.ActiveSheet
// 打印当前活动数据表的名称
console.log(sheet.Name) // Sheet2

// 将当前数据表的名称改为 WPS WebOffice
sheet.Name = 'WPS WebOffice'
```

## [Index​](#index)

数据表的 index,即该数据表在所有数据表的索引值

#### [数据类型​](#数据类型-2)

String - 该数据表在所有数据表的索引值

#### [示例​](#示例-2)

js
```js
const sheet = Application.ActiveSheet
// 打印当前活动数据表的index
console.log(sheet.Index) // 1
```

## [Visible​](#visible)

显示/隐藏 数据表

#### [数据类型​](#数据类型-3)

Boolean - 数据表是否可见

#### [示例​](#示例-3)

js
```js
const sheet = Application.ActiveSheet
// 隐藏数据表
sheet.Visible = false
// 取消数据表隐藏
sheet.Visible = true
```

## [Type​](#type)

数据表类型

#### [数据类型​](#数据类型-4)

Enum.xlEtDataBaseSheet- 数据表的类型

#### [示例​](#示例-4)

js
```js
const sheet = Application.ActiveSheet
// 打印当前活动数据表的类型
console.log(sheet.Type) //xlEtDataBaseSheet
```

## [Field​](#field)

数据表的字段， 返回一个Field对象

#### [数据类型​](#数据类型-5)

Field

#### [示例​](#示例-5)

js
```js
const sheet = Application.ActiveSheet
// 获取的表所有字段信息
const fields = sheet.Field.GetFields()
```

## [Record​](#record)

数据表的字段， 返回一个Record对象

#### [数据类型​](#数据类型-6)

Record

#### [示例​](#示例-6)

js
```js
const sheet = Application.ActiveSheet
const record = sheet.Record.GetRecord({  RecordId: 'Bz' })
```

## [Activate()​](#activate)

激活表

#### [示例​](#示例-7)

js
```js
const sheet = Application.Sheets.Item(1)
// 激活第一个表
sheet.Activate()
```

## [Move()​](#move)

移动数据表

#### [参数​](#参数)

两个参数互斥

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| Before | number | null | 否 | 验将放置移动的数据表之前的数据表 ID。如果指定 After ，则不能指定 Before。 |
| After | number | null | 否 | 将放置移动的数据表后的数据表 ID。如果指定 Before ，则不能指定 After |

#### [示例​](#示例-8)

js
```js
// 将当前数据表移动到第二个数据表之后
const sheet = Application.ActiveSheet
sheet.Move({
  Before: null,
  After: Application.Sheets(2).Id
})
```

## [Delete()​](#delete)

删除数据表

#### [返回值​](#返回值)

undefined

#### [示例​](#示例-9)

js
```js
// 删除名称为“Sheet2”的数据表
Application.Sheets.Item('Sheet2').Delete()
```

## [IsDBSheet()​](#isdbsheet)

是否为数据表

#### [返回值​](#返回值-1)

Boolean

#### [示例​](#示例-10)

js
```js
// 判断当前活跃表是否为数据表
Application.ActiveSheet.IsDBSheet()
```


#### 枚举(Enum)

# [Enum​](#enum)

枚举类型，存放在 Application 下

## [XlAboveBelow​](#xlabovebelow)

指定值是高于还是低于平均值

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlAboveAverage | 0 | 高于平均值 |
| xlBelowAverage | 1 | 低于平均值 |
| xlEqualAboveAverage | 2 | 等于平均值 |

## [XlAutoFillType​](#xlautofilltype)

根据源区域的内容，指定目标区域的填充方式

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlFillDefault | 0 | 确定用于填充目标区域的值和格式 |
| xlFillCopy | 1 | 将源区域的值和格式复制到目标区域，如有必要可重复执行 |
| xlFillSeries | 2 | 将源区域中的值扩展到目标区域中，形式为系列（如，“1, 2” 扩展为 “3, 4, 5”）。格式从源区域复制到目标区域，如有必要可重复执行 |
| xlFillFormats | 3 | 只将源区域的格式复制到目标区域，如有必要可重复执行 |
| xlFillValues | 4 | 只将源区域的值复制到目标区域，如有必要可重复执行 |
| xlFillDays | 5 | 将星期中每天的名称从源区域扩展到目标区域中。格式从源区域复制到目标区域，如有必要可重复执行 |
| xlFillWeekdays | 6 | 将工作周每天的名称从源区域扩展到目标区域中。格式从源区域复制到目标区域，如有必要可重复执行 |
| xlFillMonths | 7 | 将月名称从源区域扩展到目标区域中。格式从源区域复制到目标区域，如有必要可重复执行 |
| xlFillYears | 8 | 将年从源区域扩展到目标区域中。格式从源区域复制到目标区域，如有必要可重复执行 |
| xlLinearTrend | 9 | 将数值从源区域扩展到目标区域中，假定数字之间是加法关系（如，“1, 2,” 扩展为 “3, 4, 5”，假定每个数字都是前一个数字加上某个值的结果）。格式从源区域复制到目标区域，如有必要可重复执行 |
| xlGrowthTrend | 10 | 将数值从源区域扩展到目标区域中，假定源区域的数字之间是乘法关系（如，“1, 2,” 扩展为 “4, 8, 16”，假定每个数字都是前一个数字乘以某个值的结果）。格式从源区域复制到目标区域，如有必要可重复执行 |

## [XlBordersIndex​](#xlbordersindex)

指定要检索的边框

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlDiagonalDown | 5 | 从区域中每个单元格的左上角到右下角的边框 |
| xlDiagonalUp | 6 | 从区域中每个单元格的左下角到右上角的边框 |
| xlEdgeLeft | 7 | 区域左边缘的边框 |
| xlEdgeTop | 8 | 区域顶部的边框 |
| xlEdgeBottom | 9 | 区域底部的边框 |
| xlEdgeRight | 10 | 区域右边缘的边框 |
| xlInsideVertical | 11 | 区域中所有单元格的垂直边框（区域以外的边框除外） |
| xlInsideHorizontal | 12 | 区域中所有单元格的水平边框（区域以外的边框除外） |
| xlOutside | 13 | 区域中的 上下左右 |
| xlInside | 14 | 中间区域 |

## [XlBorderWeight​](#xlborderweight)

指定某一区域周围的边框的粗细

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlMedium | -4138 | 中 |
| xlHairline | 1 | 细线（最细的边框） |
| xlThin | 2 | 细长 |
| xlThick | 4 | 粗（最宽的边框） |

## [XlCalcModeType​](#xlcalcmodetype)

迭代计算模式

| 字段 | 值 | 释义 |
| --- | --- | --- |
| manual | manual | 手动 |
| automatic | automatic | 自动 |

## [XlChartType​](#xlcharttype)

指定图表类型

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xl3DArea | -4098 | 三维面积图。 |
| xl3DAreaStacked | 78 | 三维堆积面积图。 |
| xl3DAreaStacked100 | 79 | 百分比堆积面积图。 |
| xl3DBarClustered | 60 | 三维簇状条形图。 |
| xl3DBarStacked | 61 | 三维堆积条形图。 |
| xl3DBarStacked100 | 62 | 三维百分比堆积条形图。 |
| xl3DColumn | -4100 | 三维柱形图。 |
| xl3DColumnClustered | 54 | 三维簇状柱形图。 |
| xl3DColumnStacked | 55 | 三维堆积柱形图。 |
| xl3DColumnStacked100 | 56 | 三维百分比堆积柱形图。 |
| xl3DLine | -4101 | 三维折线图。 |
| xl3DPie | -4102 | 三维饼图。 |
| xl3DPieExploded | 70 | 分离型三维饼图。 |
| xlArea | 1 | 面积图 |
| xlAreaStacked | 76 | 堆积面积图。 |
| xlAreaStacked100 | 77 | 百分比堆积面积图。 |
| xlBarClustered | 57 | 簇状条形图。 |
| xlBarOfPie | 71 | 复合条饼图。 |
| xlBarStacked | 58 | 堆积条形图。 |
| xlBarStacked100 | 59 | 百分比堆积条形图。 |
| xlBubble | 15 | 气泡图。 |
| xlBubble3DEffect | 87 | 三维气泡图。 |
| xlColumnClustered | 51 | 簇状柱形图。 |
| xlColumnStacked | 52 | 堆积柱形图。 |
| xlColumnStacked100 | 53 | 百分比堆积柱形图。 |
| xlConeBarClustered | 102 | 簇状条形圆锥图。 |
| xlConeBarStacked | 103 | 堆积条形圆锥图。 |
| xlConeBarStacked100 | 104 | 百分比堆积条形圆锥图。 |
| xlConeCol | 105 | 三维柱形圆锥图。 |
| xlConeColClustered | 99 | 簇状柱形圆锥图。 |
| xlConeColStacked | 100 | 堆积柱形圆锥图。 |
| xlConeColStacked100 | 101 | 百分比堆积柱形圆锥图。 |
| xlCylinderBarClustered | 95 | 簇状条形圆柱图。 |
| xlCylinderBarStacked | 96 | 堆积条形圆柱图。 |
| xlCylinderBarStacked100 | 97 | 百分比堆积条形圆柱图。 |
| xlCylinderCol | 98 | 三维柱形圆柱图。 |
| xlCylinderColClustered | 92 | 簇状柱形圆锥图。 |
| xlCylinderColStacked | 93 | 堆积柱形圆锥图。 |
| xlCylinderColStacked100 | 94 | 百分比堆积柱形圆柱图。 |
| xlDoughnut | -4120 | 圆环图。 |
| xlDoughnutExploded | 80 | 分离型圆环图。 |
| xlLine | 4 | 折线图。 |
| xlLineMarkers | 65 | 数据点折线图。 |
| xlLineMarkersStacked | 66 | 堆积数据点折线图。 |
| xlLineMarkersStacked100 | 67 | 百分比堆积数据点折线图。 |
| xlLineStacked | 63 | 堆积折线图。 |
| xlLineStacked100 | 64 | 百分比堆积折线图。 |
| xlPie | 5 | 饼图。 |
| xlPieExploded | 69 | 分离型饼图。 |
| xlPieOfPie | 68 | 复合饼图。 |
| xlPyramidBarClustered | 109 | 簇状条形棱锥图。 |
| xlPyramidBarStacked | 110 | 堆积条形棱锥图。 |
| xlPyramidBarStacked100 | 111 | 百分比堆积条形棱锥图。 |
| xlPyramidCol | 112 | 三维柱形棱锥图。 |
| xlPyramidColClustered | 106 | 簇状柱形棱锥图。 |
| xlPyramidColStacked | 107 | 堆积柱形棱锥图。 |
| xlPyramidColStacked100 | 108 | 百分比堆积柱形棱锥图。 |
| xlRadar | -4151 | 雷达图。 |
| xlRadarFilled | 82 | 填充雷达图。 |
| xlRadarMarkers | 81 | 数据点雷达图。 |
| xlStockHLC | 88 | 盘高-盘低-收盘图。 |
| xlStockOHLC | 89 | 开盘-盘高-盘低-收盘图。 |
| xlStockVHLC | 90 | 成交量-盘高-盘低-收盘图。 |
| xlStockVOHLC | 91 | 成交量-开盘-盘高-盘低-收盘图。 |
| xlSurface | 83 | 三维曲面图。 |
| xlSurfaceTopView | 85 | 曲面图（俯视图）。 |
| xlSurfaceTopViewWireframe | 86 | 曲面图（俯视线框图）。 |
| xlSurfaceWireframe | 84 | 三维曲面图（线框）。 |
| xlXYScatter | -4169 | 散点图。 |
| xlXYScatterLines | 74 | 折线散点图。 |
| xlXYScatterLinesNoMarkers | 75 | 无数据点折线散点图。 |
| xlXYScatterSmooth | 72 | 平滑线散点图。 |
| xlXYScatterSmoothNoMarkers | 73 | 无数据点平滑线散点图。 |

## [XlContainsOperator​](#xlcontainsoperator)

指定函数使用的运算符

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlContains | 0 | 包含指定的值 |
| xlDoesNotContain | 1 | 不包含指定的值 |
| xlBeginsWith | 2 | 以指定的值开始 |
| xlEndsWith | 3 | 以指定的值结束 |

## [XlDeleteShiftDirection​](#xldeleteshiftdirection)

指定如何移动单元格来替换删除的单元格

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlShiftToLeft | -4159 | 单元格向左移动 |
| xlShiftUp | -4162 | 单元格向上移动 |

## [XlDirection​](#xldirection)

指定移动的方向

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlDown | -4121 | 下拉 |
| xlToLeft | -4159 | 向左 |
| xlToRight | -4161 | 向右 |
| xlUp | -4162 | 加速 |

## [XlDVAlertStyle​](#xldvalertstyle)

指定验证过程中显示的消息框所用的图标

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlValidAlertStop | 1 | 停止图标 |
| xlValidAlertWarning | 2 | 警告图标 |
| xlValidAlertInformation | 3 | 信息图标 |

## [XlDVType​](#xldvtype)

指定要对值进行的有效性测试类型

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlValidateWholeNumber | 1 | 全部数值 |
| xlValidateDecimal | 2 | 数值 |
| xlValidateList | 3 | 值必须存在于指定列表中 |
| xlValidateDate | 4 | 日期值 |
| xlValidateTime | 5 | 时间值 |
| xlValidateTextLength | 6 | 文本长度 |
| xlValidateCustom | 7 | 使用任意公式验证数据有效性 |

## [XlExportImgFormatType​](#xlexportimgformattype)

导出图片的格式

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlImgTypePNG | 0 | 导出 .png |
| xlImgTypeJPG | 1 | 导出 .jpg |
| xlImgTypeBMP | 2 | 导出 .bmp |
| xlImgTypeTIF | 2 | 导出 .tif |

## [XlFixedFormatType​](#xlfixedformattype)

指定文件格式的类型

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlTypePDF | 0 | PDF（.pdf） |
| xlTypeXPS | 1 | XPS（.xps） |
| xlTypeIMG | 2 | IMG（.png、.jpg、.bmp、.tif） |

## [XlFormatConditionOperator​](#xlformatconditionoperator)

指定用于将公式与单元格中的值或 xlBetween 和 xlNotBetween 中的值进行比较，以比较两个公式的运算符

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlBetween | 1 | 行间，只在提供了两个公式的情况下才能使用 |
| xlNotBetween | 2 | 不介于。只在提供了两个公式的情况下才能使用 |
| xlEqual | 3 | 平等 |
| xlNotEqual | 4 | 不等于 |
| xlGreater | 5 | 大于 |
| xlLess | 6 | 小于 |
| xlGreaterEqual | 7 | 大于或等于 |
| xlLessEqual | 8 | 小于或等于 |

## [XlFormatConditionType​](#xlformatconditiontype)

指定条件格式是基于单元格值还是基于表达式

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlCellValue | 1 | 单元格值 |
| xlExpression | 2 | 表达式 |
| xlColorScale | 3 | 色阶 |
| xlTop10 | 5 | 前 10 个值 |
| xlUniqueValues | 8 | 唯一值 |
| xlTextString | 9 | 文本字符串 |
| xlBlanksCondition | 10 | 空值条件 |
| xlTimePeriod | 11 | 时间段 |
| xlAboveAverageCondition | 12 | 高于平均值条件 |
| xlNoBlanksCondition | 13 | 无空值条件 |
| xlErrorsCondition | 16 | 错误条件 |
| xlNoErrorsCondition | 17 | 无错误条件 |

## [XlHAlign​](#xlhalign)

指定对象的水平对齐方式

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlHAlignRight | -4152 | 靠右 |
| xlHAlignLeft | -4131 | 靠左 |
| xlHAlignJustify | -4130 | 两端对齐 |
| xlHAlignDistributed | -4117 | 分散对齐 |
| xlHAlignCenter | -4108 | 居中 |
| xlHAlignGeneral | 1 | 按数据类型对齐 |
| xlHAlignFill | 5 | 填充 |
| xlHAlignCenterAcrossSelection | 7 | 跨列居中 |

## [XlInsertFormatOrigin​](#xlinsertformatorigin)

指定从何处复制插入单元格的格式

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlFormatFromLeftOrAbove | 0 | 从上方和/或左侧单元格复制格式 |
| xlFormatFromRightOrBelow | 1 | 从下方和/或右侧单元格复制格式 |

## [XlInsertShiftDirection​](#xlinsertshiftdirection)

指定插入时单元格的移动方向

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlShiftDown | -4121 | 向下移动单元格 |
| xlShiftToRight | -4161 | 向上移动单元格 |

## [XlLineStyle​](#xllinestyle)

指定边框的线条样式

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlLineStyleNone | -4142 | 无线 |
| xlDouble | -4119 | 双线 |
| xlDot | -4118 | 点式线 |
| xlDash | -4115 | 虚线 |
| xlContinuous | 1 | 实线 |
| xlDashDot | 4 | 点划相间线 |
| xlDashDotDot | 5 | 划线后跟两个点 |
| xlSlantDashDot | 13 | 倾斜的划线 |

## [XlPasteSpecialOperation​](#xlpastespecialoperation)

XlPaste 特殊操作

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlPasteSpecialOperationAdd | 0x1 | 复制的数据与目标单元格中的值相加 |
| xlPasteSpecialOperationDivide | 0x4 | 复制的数据除以目标单元格中的值 |
| xlPasteSpecialOperationMultiply | 0x3 | 复制的数据乘以目标单元格中的值 |
| xlPasteSpecialOperationNone | 0x0 | 粘贴操作中不执行任何计算 |
| xlPasteSpecialOperationSubtract | 0x2 | 复制的数据减去目标单元格中的值 |

## [XlPasteType​](#xlpastetype)

粘贴类型

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlPasteFormulas | 0x2 | 公式 |
| xlPasteAllExceptBorders | 0x4 | 无边框 |
| xlPasteColumnWidths | 0x5 | 保留源列宽 |
| xlPasteValues | 0x3 | 值 |
| xlPasteValuesAndNumberFormats | 0x7 | 值和数字格式 |
| xlPasteFormats | 0x8 | 格式 |
| xlPastePasteAll | 0x0 | 粘贴 |
| xlPasteComments | 0x9 | 批注 |
| xlPasteValidation | 0xa | 数据有效性验证 |

## [XlReferenceStyle​](#xlreferencestyle)

指定引用样式

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlR1C1 | -4150 | 使用 xlR1C1 返回 R1C1 样式的引用 |
| xlA1 | 1 | 默认值。 使用 xlA1 返回 A1 样式的引用 |

## [XlRowCol​](#xlrowcol)

指定对应于特定数据系列的数值是处于行中还是列中

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlRows | 1 | 数据系列在列中 |
| xlColumns | 2 | 数据系列在行中 |

## [XlSheetType​](#xlsheettype)

指定工作表类型

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlWorksheet | -4167 | 工作表 |
| xlDialogSheet | -4116 | 对话框工作表 |
| xlChart | -4109 | 图表 |
| xlExcel4MacroSheet | 3 | Excel 版本 4 宏工作表 |
| xlExcel4IntlMacroSheet | 4 | Excel 版本 4 国际宏工作表 |

## [XlTimePeriods​](#xltimeperiods)

指定时间段

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlToday | 0 | 今天 |
| xlYesterday | 1 | 昨天 |
| xlLast7Days | 2 | 过去 7 天 |
| xlThisWeek | 3 | 本周 |
| xlLastWeek | 4 | 上周 |
| xlLastMonth | 5 | 上月 |
| xlTomorrow | 6 | 明天 |
| xlNextWeek | 7 | 下周 |
| xlNextMonth | 8 | 下月 |
| xlThisMonth | 9 | 本月 |

## [XlUnderlineStyle​](#xlunderlinestyle)

指定应用于字体的下划线类型

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlUnderlineStyleDouble | -4119 | 粗双下划线 |
| xlUnderlineStyleDoubleAccounting | 5 | 紧靠在一起的两条细下划线 |
| xlUnderlineStyleNone | -4142 | 无下划线 |
| xlUnderlineStyleSingle | 2 | 单下划线 |
| xlUnderlineStyleSingleAccounting | 4 | 不支持 |

## [XlVAlign​](#xlvalign)

指定对象的垂直对齐方式

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlVAlignTop | -4160 | 向上 |
| xlVAlignJustify | -4130 | 调整使全行排满 |
| xlVAlignDistributed | -4117 | 一起 |
| xlVAlignCenter | -4108 | 居中 |
| xlVAlignBottom | -4107 | 向下 |

## [XlXLMMacroType​](#xlxlmmacrotype)

指定在工作表中，名称引用哪种宏，或名称是否引用宏

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlFunction | 1 | 自定义函数 |
| xlCommand | 2 | 自定义命令 |
| xlNotXLM | 3 | 非宏 |

## [XlAutoFilterOperator​](#xlautofilteroperator)

当前工作表中，指定筛选的类型，也用于指定用于关联两个筛选条件的操作符

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlAnd | 1 | Criteria1 和 Criteria2 的逻辑与 |
| xlOr | 2 | Criteria1 或 Criteria2 的逻辑或 |
| xlTop10Items | 3 | 显示最高值项 (在 Criteria1 中指定的项目数) |
| xlBottom10Items | 4 | 显示最低值项 (在 Criteria1 中指定的项目数) |
| xlTop10Percent | 5 | 显示最高值项 (Criteria1 中指定的百分比) |
| xlBottom10Percent | 6 | 显示最低值项 (在 Criteria1 中指定的百分比) |
| xlFilterValues | 7 | 筛选值 |
| xlFilterCellColor | 8 | 单元格颜色 |
| xlFilterFontColor | 9 | 字体颜色 |
| xlFilterIcon | 10 | 筛选图标 |
| xlFilterDynamic | 11 | 动态筛选 |

## [XlYesNoGuess​](#xlyesnoguess)

指定第一行是否包含标题。 对数据透视表进行排序时，不能使用该参数

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlGuess | 0 | 自动判断 Excel 是否有表头 |
| xlYes | 1 | 默认值。 应对整个区域进行排序 |
| xlNo | 2 | 不应对整个区域进行排序 |

## [XlSortOrientation​](#xlsortorientation)

指定排序方向

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlSortColumns | 1 | 按列排序 |
| xlSortRows | 2 | 按行排序，此值为默认值 |

## [XlSortMethod​](#xlsortmethod)

指定排序类型

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlPinYin | 1 | 按字符的汉语拼音顺序排序，此值为默认值 |
| xlStroke | 2 | 按每个字符的笔划数排序 |

## [XlSortDataOption​](#xlsortdataoption)

指定文本的排序方式

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlSortNormal | 0 | 分别对数字和文本数据进行排序，此值为默认值 |
| xlSortTextAsNumbers | 1 | 将文本作为数字型数据进行排序 |

## [XlSortOn​](#xlsorton)

指定数据的排序参数

| 字段 | 值 | 释义 |
| --- | --- | --- |
| SortOnCellColor | 1 | 单元格颜色 |
| SortOnFontColor | 2 | 字体颜色 |
| SortOnIcon | 3 | 图标 |
| SortOnValues | 0 | 值 |

## [XlSortOrder​](#xlsortorder)

为指定字段或范围指定排序顺序

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlAscending | 1 | 默认值，按升序对指定字段排序 |
| xlDescending | 2 | 按降序对指定字段排序 |

## [XlTextParsingType​](#xltextparsingtype)

指定要导入到查询表中的文本文件中的数据的列格式

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlDelimited | 1 | 默认值，指示文件由分隔符分隔 |
| xlFixedWidth | 2 | 指示将文件中的数据排列在固定宽度的列中 |

## [XlTextQualifier​](#xltextqualifier)

指定用于指定文本的分隔符

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlTextQualifierDoubleQuote | 1 | 双引号 (") |
| xlTextQualifierNone | -4142 | 无分隔符 |
| xlTextQualifierSingleQuote | 2 | 单引号 (') |

## [XlColumnDataType​](#xlcolumndatatype)

指定列的分列方式

| 字段 | 值 | 释义 |
| --- | --- | --- |
| xlDMYFormat | 4 | DMY 日期格式 |
| xlDYMFormat | 7 | DYM 日期格式 |
| xlEMDFormat | 10 | EMD 日期格式 |
| xlGeneralFormat | 1 | 常规 |
| xlMDYFormat | 3 | MDY 日期格式 |
| xlMYDFormat | 6 | MYD 日期格式 |
| xlSkipColumn | 9 | 列未分列 |
| xlTextFormat | 2 | 文本 |
| xlYDMFormat | 8 | YDM 日期格式 |
| xlYMDFormat | 5 | YMD 日期格式 |


#### 行记录(Record)

# [Record​](#record)

行记录

### [方法列表​](#方法列表)

| 方法名 | 返回类型 | 简介 |
| --- | --- | --- |
| GetRecords() | Array | 获取行记录（多条） |
| GetRecord() | Object | 获取行记录（单条） |
| DeleteRecords() | Array | 删除行记录 |
| UpdateRecords() | Array | 更新行记录 |
| CreateRecords() | Array | 创建行记录 |
| GetAttachmentURL() | String | 获取上传附件或图片的URL |

## [GetRecords()​](#getrecords)

获取行记录（多条）

注意

每次请求最多返回100条，数据量大的时候请使用分页查询

### [参数​](#参数)

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| ViewId | String |  | 否 | 填写后将从被指定的视图获取该用户所见到的记录；若不填写，则从工作表获取记录 |
| PageSize | Number | 100 | 否 | 存在分页时，指定本次查询的起始记录（含）。若不填写或填写为空字符串，则从第一条记录开始获取。当前最大值：1000 |
| Offset | Number |  | 否 | 分页查询时，将返回一个offset值，指向下一页的第一条记录，供后续查询。查询到最后一页或第maxRecords条记录时，返回数据将不再包含offset值 |
| MaxRecords | Number |  | 否 | 指定要获取的“前maxRecords条记录”，若不填写，则默认返回全部记录 |
| Fields | Array |  | 否 | 字段类型 |
| Filter | Object |  | 否 | 详细说明见附录三 |

### [返回值​](#返回值)

Object - 获取表的所有记录

| 属性 | 数据类型 | 说明 |
| --- | --- | --- |
| Offset | String | 如果分页的话， 则会返回此字段信息;分页截止 id， 下次请求携带会继续分页请求信息 |
| Records | Array[Object] | 记录集合 |

#### [记录集合​](#记录集合)

| 属性 | 数据类型 | 说明 |
| --- | --- | --- |
| id | String | 记录Id |
| Fields | Object | 更新的字段信息，包含字段Id，字段name,格式说明见附录 |

### [示例​](#示例)

javascript
```javascript
const sheet = Application.ActiveSheet
// 分页查询例子
function fetchAllRecords() {
  const view = sheet.Selection.GetActiveView()
  let all = []
  let offset = null;

  while (all.length === 0 || offset) {
    let records = sheet.Record.GetRecords({
      ViewId: view.viewId,
      Offset: offset,
    })
    offset = records.offset
    all = all.concat(records.records)
  }
  console.log(all.length)
  return all
}

fetchAllRecords()
```

## [GetRecord()​](#getrecord)

获取行记录（单条）

### [参数​](#参数-1)

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| RecordId | String |  | 是 | 表中指定获取的记录id |

### [返回值​](#返回值-1)

Object - 获取表的指定的单条记录

| 属性 | 数据类型 | 说明 |
| --- | --- | --- |
| id | String | 记录Id |
| Fields | Object | 更新的字段信息，包含字段Id，字段name,格式说明见附录 |

### [示例​](#示例-1)

javascript
```javascript
const sheet = Application.ActiveSheet
const record = sheet.Record.GetRecord({  RecordId: 'Bz' })
console.log(record)
// 打印结果：
//  {"fields":{"日期":"2023/02/21"},"id":"Bz"}
```

## [DeleteRecords()​](#deleterecords)

删除行记录

### [参数​](#参数-2)

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| RecordIds | Array |  | 是 | 表中需要删除的记录id |

### [返回值​](#返回值-2)

Array - 返回删除的表id以及删除是否成功信息

| 属性 | 数据类型 | 说明 |
| --- | --- | --- |
| id | String | 记录Id |
| deleted | Boolean | 是否删除成功 “true”表示删除成功，“false”表示删除失败 |

### [示例​](#示例-2)

javascript
```javascript
const sheet = Application.ActiveSheet
const result = sheet.Record.DeleteRecords({ 
    RecordIds: ['J', 'P', 'Q'] 
})
console.log(resutlt)
// 打印结果：
// [{"deleted":true,"id":"P"},{"deleted":false,"id":"Q"}]
```

## [UpdateRecords()​](#updaterecords)

更新行记录

### [参数​](#参数-3)

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| Records | Array[Object] |  | 是 | 行记录集合 |

#### [行记录集合：​](#行记录集合)

| 属性 | 数据类型 | 说明 |
| --- | --- | --- |
| id | String | 记录Id |
| Fields | Object | 更新的字段信息，包含字段Id，字段name,格式说明见附录 |

### [返回值​](#返回值-3)

Array - 表的已更新的所有记录

| 属性 | 数据类型 | 说明 |
| --- | --- | --- |
| id | String | 记录Id |
| Fields | Object | 更新的字段信息，包含字段Id，字段name,格式说明见附录 |

### [示例​](#示例-3)

javascript
```javascript
const sheet = Application.ActiveSheet
const records = sheet.Record.UpdateRecords({
        Records: [{
            id: 'A',
            fields: {
                 邮箱: 'demo@qq.com',
                 多选: ['1', '2'],
                 "记录关联": {
                    "recordIds": ["I", "K"] 
                 }
            }
        }],
    })
```

## [CreateRecords()​](#createrecords)

创建行记录

### [参数​](#参数-4)

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| Records | Array[Object] |  | 是 | 行记录集合 |

#### [行记录集合：​](#行记录集合-1)

| 属性 | 数据类型 | 说明 |
| --- | --- | --- |
| id | String | 记录Id |
| Fields | Object | 更新的字段信息，包含字段Id，字段name,格式说明见附录 |

### [返回值​](#返回值-4)

Array - 表的已更新的所有记录

| 属性 | 数据类型 | 说明 |
| --- | --- | --- |
| id | String | 记录Id |
| Fields | Object | 更新的字段信息，包含字段Id，字段name,格式说明见附录 |

### [示例​](#示例-4)

javascript
```javascript
const sheet = Application.ActiveSheet
// 创建邮箱和多选
const records = sheet.Record.CreateRecords({
      Records: [{
          fields: {
               邮箱: 'demo@qq.com',
               多选: ['1', '2'],
          }
      }, {
          fields: {
               邮箱: 'demo@qq.com',
               多选: ['1', '2'],
          }
      }],
  })

// 创建联系人
const records = sheet.Record.CreateRecords({
  Records: [
    {  fields: { '联系人': [{ name: 'yourname', nickName: 'yourname', id: '88888888', avatar_url: 'https://avatar.qwps.cn/avatar/5b2t57-U' }] } },
  ],
});
```

## [GetAttachmentURL()​](#getattachmenturl)

获取上传附件或图片的URL

### [参数​](#参数-5)

注意

必须至少传入1个参数Attachment或者传入2个参数UploadId和Source

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| Attachment | String |  | 否 | 附件 |
| UploadId | String |  | 否 | 上传文件id |
| Source | String |  | 否 | source参数必须为"upload_ks3"（本地上传）或"cloud"（云上传） |

### [返回值​](#返回值-5)

String - 为获取上传附件或图片的URL，打开该URL可进行附件或图片下载

### [示例​](#示例-5)

javascript
```javascript
const sheet = Application.ActiveSheet
const resultURL = sheet.Record.GetAttachmentURL({
    Attachment: "IKWRCBAAKA|upload_ks3|image/png|QQ图片20230214165215.png|12070||549*106",
      })

//or

const resultURL = sheet.Record.GetAttachmentURL({
    UploadId: "IKWRCBAAKA",
    Source: "upload_ks3"
      })
```


#### 表格实例(Application)

# [Application​](#application)

文档操作的顶级对象，对文档进行相关操作，都是间接或直接操作该对象。

Application 是一个文件的顶级对象，新打开一个文件返回的也是 Application。

而在脚本中的Application则是指当前文件的顶级对象，有且只有一个。

Application 对象的具体属性和方法请参阅下方的列表。

#### [属性列表​](#属性列表)

| 属性 | 数据类型 | 简介 |
| --- | --- | --- |
| ActiveSheet | Sheet | 当前的活动工作表/数据表 |
| Sheets | Sheets | 当前文件的所有工作表/数据表 |
| FileInfo | Object | 当前文档的信息 |
| UserInfo | Object | 当前文档的用户信息 |
| Enum | Enum | 所有的枚举类型 |

#### [方法列表​](#方法列表)

| 方法 | 返回类型 | 简介 |
| --- | --- | --- |
| Sheets(name) | Sheet | 获取名称为 name 的工作表/数据表 |

## [ActiveSheet​](#activesheet)

当前活动工作表/数据表，可以通过 Sheet.Activate()来切换活动工作表/数据表。该属性返回Sheet对象,能利用该属性操作当前活动工作表/数据表。

运行脚本的环境是独立在服务器的，因此脚本运行环境的 ActiveSheet 与用户环境的 ActiveSheet 不一定相同。

具体规则是：

1.运行脚本时会把脚本运行环境的 ActiveSheet 切换为用户环境当前的 ActiveSheet。

2.当脚本通过函数切换脚本运行环境的 ActiveSheet 时，用户环境的 ActiveSheet 不会同步切换。

#### [数据类型​](#数据类型)

Sheet- 当前活动工作表/数据表

#### [示例​](#示例)

js
```js
console.log(Application.ActiveSheet.Name) // 数据表2

// 切换到名称为数据表2的数据表
Application.Sheets.Item('数据表2').Activate()
console.log(Application.ActiveSheet.Name) // 数据表2
```

## [Sheets​](#sheets)

获取当前文件能操作的所有 Sheet，返回一个Sheets对象。

#### [数据类型​](#数据类型-1)

Sheets

#### [示例​](#示例-1)

js
```js
// 工作簿（Workbook）中所有工作表/数据表（Sheet）的集合,下面两种写法是一样的
let sheets = Application.ActiveWorkbook.Sheets
sheets = Application.Sheets

// 打印所有工作表/数据表的名称
for (let i = 1; i <= sheets.Count; i++) {
  console.log(sheets.Item(i).Name)
}
```

### [Sheets.Count​](#sheets-count)

工作表/数据表数量

#### [数据类型​](#数据类型-2)

Number - 对应工作簿的工作表/数据表数量

#### [示例​](#示例-2)

js
```js
// 下面两种写法是一样的
let sheets = Application.ActiveWorkbook.Sheets
sheets = Application.Sheets

// 打印所有工作表/数据表的名称
console.log(sheets.Count) //1
```

### [Sheets.DefaultNewSheetName​](#sheets-defaultnewsheetname)

默认新工作表名

#### [返回类型​](#返回类型)

String - 新建工作表时若没有指定名称，可用这个名称作为新建工作表名称

#### [示例​](#示例-3)

js
```js
const defaultName = Application.Sheets.DefaultNewSheetName
// 工作表对象
Application.Sheets.Add(
  null,
  Application.ActiveSheet.Name,
  1,
  Application.Enum.XlSheetType.xlWorksheet,
  defaultName
)
```

### [Sheets.Add()​](#sheets-add)

新增工作表，如果 Before 和 After 都存在，以 Before 为准

#### [参数​](#参数)

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| Before | String/Number |  | 否 | After 空时，必填，为当前已有单元格的 index 或者名称，新建的工作表将置于此工作表之前 |
| After | String/Number |  | 否 | Before 空时，必填，为当前已有单元格的 index 或者名称，新建的工作表将置于此工作表之后 |
| Count | Number | 1 | 否 | 要添加的工作表数。默认值为选定工作表的数量 |
| Type | Enum |  | 否 | 指定工作表类型，详细可见Enum.XlSheetType |
| Name | Name |  | 否 | 指定工作表名称 |

#### [示例​](#示例-4)

js
```js
// 添加工作表
Application.Sheets.Add(
  null,
  Application.ActiveSheet.Name,
  1,
  Application.Enum.XlSheetType.xlWorksheet,
  '新工作表'
)
```

### [Sheets.Item()​](#sheets-item)

根据名称或索引选择 Sheet

#### [参数​](#参数-1)

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| index | String/Number |  | 是 | 所选的 sheet 的名称/索引 |

#### [返回类型​](#返回类型-1)

Sheet- 对应名称的工作表/数据表

#### [示例​](#示例-5)

js
```js
// 切换名称为"Sheet2"的工作表
Application.Sheets.Item('Sheet2').Activate()

// 切换索引为1的工作表
Application.Sheets.Item(1).Activate()
```

### [Sheets.Each()​](#sheets-each)

遍历所有 sheet 并执行回调函数

#### [参数​](#参数-2)

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| callback | Function | null | 是 | 类似 JS 数组的 forEach |

#### [示例​](#示例-6)

js
```js
// 打印所有工作表/数据表的名称
Application.Sheets.Each(function (item) {
  console.log(item.Name) //Sheet1 Sheet2
})
```

## [FileInfo​](#fileinfo)

返回当前文件的基本信息。

#### [数据类型​](#数据类型-3)

Object - 当前文件的信息

| 名称 | 类型 | 说明 |
| --- | --- | --- |
| id | string | 文件 ID |
| name | string | 文件名 |
| officeType | string | 文档类型 |
| creator | CreatorObject | 文档创建者信息 |
| size | number | 文件大小 |
| groupId | string | 文件的群组 ID |
| docType | number | 文档类型（数字形式） |

#### [CreatorObject 对象信息​](#creatorobject)

| 名称 | 类型 | 说明 |
| --- | --- | --- |
| id | string | 创建者 ID |
| name | string | 创建者名称 |
| avatar_url | string | 创建者头像 |
| logined | boolean | 是否已登录 |
| attrs | Object | 属性对象 |
| real_id | string | 真实 ID |

#### [示例​](#示例-7)

javascript
```javascript
// 打印文件信息
console.log(Application.FileInfo)
/*{
 "id": "<open_id>",
 ...
}*/
```

## [UserInfo​](#userinfo)

返回当前文件的用户信息。

#### [数据类型​](#数据类型-4)

Object- 当前文件的用户信息

| 名称 | 类型 | 说明 |
| --- | --- | --- |
| id | string | 用户 ID |
| name | string | 用户名称 |

#### [示例​](#示例-8)

javascript
```javascript
// 打印用户信息
console.log(Application.UserInfo)
```

## [Enum​](#enum)

枚举类型，存放在 Application 下。

可以通过 Application.Enum 使用

#### [数据类型​](#数据类型-5)

Enum- 所有的枚举类型

#### [示例​](#示例-9)

js
```js
// 打印工作表/数据表的类型枚举
console.log(Application.Enum.XlSheetType)
//{"xlChart":-4109,"xlDialogSheet":-4116,"xlExcel4IntlMacroSheet":4,"xlExcel4MacroSheet":3,"xlWorksheet":-4167}
```

## [Sheets()​](#sheets-1)

作为函数使用，代替 Sheets.Item()，返回一个Sheet对象。

#### [参数​](#参数-3)

| 名称 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| name | string | 是 | 工作表/数据表的名称 |

#### [返回类型​](#返回类型-2)

Sheet- 对应名称的工作表/数据表 Sheet 对象

#### [示例​](#示例-10)

js
```js
console.log(Application.Sheets.Count) // 1

// 以下两种写法效果是一样的
console.log(Application.Sheets('Sheet2').Range('A1').Text) // Sheet2的A1单元格的内容
console.log(Application.Sheets.Item('Sheet2').Range('A1').Text) // Sheet2的A1单元格的内容
```


#### 附录

# [附录​](#附录)

## [附录 1：数据表字段类型说明​](#附录-1-数据表字段类型说明)

| 字段类型 | Type | 创建字段格式 | 设置字段值传入形式 | 读取字段值传出形式 |
| --- | --- | --- | --- | --- |
| 多行文本 | MultiLineText | 无特殊要求 | 字符串/ 无特殊格式要求 | 字符串 |
| 日期 | Date | 无特殊要求 | 字符串/yyyy/mm/dd | 字符串 |
| 时间 | Time | 无特殊要求 | 字符串/hh:mm:ss | 字符串 |
| 数值 | Number | 无特殊要求 | 数值 / 无格式 | 数值 |
| 货币 | Currency | 无特殊要求 | 数值 / 无格式 | 数值 |
| 百分比 | Percentage | 无特殊要求 | 数值 / 无格式 | 数值 |
| 身份证 | ID | 无特殊要求 | 字符串 / 符合身份证规则 | 字符串 |
| 电话 | Phone | 无特殊要求 | 字符串 / 符合电话规则 | 字符串 |
| 电子邮箱 | Email | 无特殊要求 | 字符串 / 符合邮箱规则 | 字符串 |
| 超链接 | Url | 可以额外传入一个参数。displayText：指定超链接显示文本。{"name":"超链接","type":"Url","displayText":"跳转"} | 字符串 / 符合 Url 规 | 字符串 |
| 复选框 | Checkbox | 无特殊要求 | true / false | 布尔 |
| 单选项 | SingleSelect | 需要额外传入选项值，至少一个。{"name": "单选项","type": "SingleSelect","items": [{ "value": "item1" }]} | 字符串 / 匹配选项内容 | 字符串 |
| 多选项 | MultipleSelect | 需要额外传入选项值，至少一个。{"name": "单选项","type": "SingleSelect","items": [{ "value": "item1" }, { "value": "item2" }]} | 字符串数组 / 匹配选项内容 | 字符串数组 |
| 等级 | Rating | 需要额外传入一个最大等级, 最大等级大于 0 小于等于 5。{"name": "等级","type": "Rating","max": 5} | 数值 / 大于 0 并且 小于 最大等级 | 数值 |
| 进度条 | Complete | 无特殊要求 | 数值 / 大于等于 0 并且 小于 100 | 字符串 |
| 联系人 | Contact | 需要额外传入两个参数：multipleContacts:<bool>是否支持多个联系人noticeNewContact:<bool>是否通知联系人。{"name": "联系人","type": "Contact","multipleContacts": false,"noticeNewContact": false} | 不支持设值 | 对象 |
| 附件 | Attachment | 无特殊要求 | 不支持设值 | 对象 |
| 关联 | Link | 需要额外传入二个参数：linkSheet: 关联表 IDmultipleLinks: 是否关联多条记录{"name": "联系人","type": "Link","multipleContacts": false,"noticeNewContact": false} | 对应关联表的行记录数组 |  |
| 富文本 | Note | 无特殊要求 | 不支持设值 | 对象 |
| 编号 | AutoNumber | 无特殊要求 | 不支持设值 | 数值 |
| 创建者 | CreatedBy | 无特殊要求 | 不支持设值 | 对象 |
| 创建时间 | CreatedTime | 无特殊要求 | 不支持设值 | 字符串 |
| 公式 | Formula | 无特殊要求 | 不支持设值 | 根据公式的值类型 |
| 引用 | Lookup | 无特殊要求 | 不支持设值 | 与被引用形式相同 |

## [附录 2：数据表视图类型说明​](#附录-2-数据表视图类型说明)

| 视图类型 | 说明 |
| --- | --- |
| Grid | 表格视图 |
| Kanban | 看板视图 |
| Gallery | 画册视图 |
| Form | 表单视图 |
| Gantt | 甘特视图 |

## [附录 3：筛选条件说明​](#附录-3-筛选条件说明)

筛选条件用来对行记录进行筛选，由两部分构成：mode为筛选条件关系；creteria为具体筛选条件（fileds op values）。

json
```json
{
  "mode": "AND", // 选填。表示各筛选条件之间的逻辑关系。只能是"AND"或"OR"。缺省值为"AND"
  "criteria": [
    // filter结构体内必填。包含筛选条件的数组。每个字段上只能有一个筛选条件
    {
      "field": "名称", // 必填。根据 preferId 与否，需要填入字段名或字段id
      "op": "Intersected", // 必填。表示具体的筛选规则，见下
      "values": [
        // 必填。表示筛选规则中的值。数组形式。
        "数据表", // 值为字符串时表示文本匹配
        "12345"
      ]
    },
    {
      "field": "数量",
      "op": "Greater",
      "values": ["1"]
    }
  ]
}
```

| 筛选条件 | 参数说明 |
| --- | --- |
| Equals | 等于 |
| NotEqu | 不等于 |
| Greater | 大于 |
| GreaterEqu | 大等于 |
| Less | 小于 |
| LessEqu | 小等于 |
| GreaterEquAndLessEqu | 介于（取等） |
| LessOrGreater | 介于（不取等） |
| BeginWith | 开头是 |
| EndWith | 结尾是 |
| Contains | 包含 |
| NotContains | 不包含 |
| Intersected | 指定值 |
| Empty | 为空 |
| NotEmpty | 不为空 |

各筛选规则独立地限制了 values 数组内最多允许填写的元素数，当 values 内元素数超过阈值时，该筛选规则将失效。

为空、不为空不允许填写元素；介于允许最多填写 2 个元素；指定值允许填写 65535 个元素；其他规则允许最多填写 1 个元素。

注意

filter 不是结构体，当 criteria 未指定 field、op/values 参数填写不合法、values 填写过多参数及其他可能导致筛选规则失效等情形，整个请求将直接失败。

目前还支持对日期进行动态筛选，此时 values[]内的元素需以结构体的形式给出：

json
```json
{
  "mode": "AND",
  "criteria": [
    {
      "field": "日期",
      "op": "Equals",
      "values": [
        {
          "dynamicType": "lastMonth",
          "type": "DynamicSimple"
        }
      ]
    }
  ]
}
```

提示

上述示例对应的筛选条件为等于上一个月。

要使用日期动态筛选，values[]内的结构体需要指定type为DynamicSimple，当op为Equals时，dynamicType可以为如下的值（大小写不敏感）。

| 字段 | 说明 |
| --- | --- |
| today | 今天 |
| yesterday | 昨天 |
| tomorrow | 明天 |
| last7Days | 最近 7 天 |
| last30Days | 最近 30 天 |
| thisWeek | 本周 |
| lastWeek | 上周 |
| nextWeek | 下周 |
| thisMonth | 本月 |
| lastMonth | 上月 |
| nextMonth | 次月 |

提示

当op为greater或less时，dynamicType只能是昨天、今天或明天。


## 高级服务

### 云文档 API

# [云文档 API​](#云文档-api)

AirScript 提供全局的 KSDrive 对象，通过此对象即可轻松查看、修改和创建您的云文档

提示

在使用 KSDrive 对象操作云文档时，确保您已添加云文档API服务，在脚本编辑器的服务菜单内添加即可。

### [快速使用​](#快速使用)

js
```js
// 打开指定文档
let file = KSDrive.openFile('https://www.kdocs.cn/l/xxxxxxxxxxxx')
// 打印指定文档的A1单元格内容
console.log(file.Application.Range('A1').Text)
// 使用结束之后调用close关闭文档，否则无法再次调用KSDrive.openFile
file.close()
// 获取我的云文档下面的et，ksheet文档列表
const fileList = KSDrive.listFiles({ includeExts: ['et', 'ksheet'] })
// 打开我的云文档目录下的第一个文档
file = KSDrive.openFile(fileList.files[0])
console.log(file.Application.Range('A1').Text)
// 关闭文档
file.close()
```

### [属性列表​](#属性列表)

| 属性名 | 数据类型 | 说明 |
| --- | --- | --- |
| FileType | object | 支持的文件类型集合 |

### [方法列表​](#方法列表)

| 方法名 | 返回类型 | 说明 |
| --- | --- | --- |
| createFile() | string | 创建或另存一个文件 |
| openFile() | File | 额外打开一个文件 |
| listFiles() | FilesInfo | 列出某个目录下的表格文件 |

## [FileType​](#filetype)

云文档支持的文件类型，可用于新建文件时指定新文件的类型

### [属性说明​](#属性说明)

| 属性名 | 数据类型 | 说明 |
| --- | --- | --- |
| AP | string | 智能文档 |
| KSheet | string | 智能表格 |
| ET | string | 表格 |
| DB | string | 多维表 |

## [createFile()​](#createfile)

创建一个新文件，也可以将一个源文件另存为新文件

### [参数​](#参数)

| 名称 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| type | FileType | 是 | 新文件的类型 |
| createOptions | CreateOptions | 是 | 新文件的参数选项 |

### [CreateOptions 对象说明​](#createOptions)

| 名称 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| name | string | 是 | 新文件的文件名 |
| dirUrl | string | 否 | 新文件的文件目录 |
| source | string | 否 | 将目标文件另存为新文件 |

### [返回值​](#返回值)

url - string 新文件的 URL

### [示例​](#示例)

js
```js
// 创建ET文件，指定保存位置
let url = KSDrive.createFile(KSDrive.FileType.ET, {
  name: 'et测试',
  dirUrl: '指定保存位置'
})
console.log(url)
// 新建DB文件
url = KSDrive.createFile(KSDrive.FileType.DB)
console.log(url)
// 新建KSheet文件
url = KSDrive.createFile(KSDrive.FileType.KSheet)
console.log(url)
// 新建AP文件
url = KSDrive.createFile(KSDrive.FileType.AP)
console.log(url)
// 文件另存
url = KSDrive.createFile(KSDrive.FileType.KSheet, {
  source: 'https://www.kdocs.cn/l/cqQwuiG2mo7E',
  name: '复制表格'
})
console.log(url)
```

## [openFile()​](#openfile)

额外打开一个文件，并返回一个 JavaScript 对象File。

### [示例​](#示例-1)

js
```js
let file = KSDrive.openFile('https://www.kdocs.cn/l/xxxxxxxxxxxx')
console.log(file.Application.ActiveSheet.Range('A1').Text)
file.close()
```

### [参数​](#参数-1)

| 名称 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| openInfo | URL /FileInfo | 是 | 打开文件的信息，可以为文件分享链接或者FileInfo |

### [返回值​](#返回值-1)

File- 一个 JavaScript 对象

## [listFiles()​](#listfiles)

列出某个目录下的所有文件和对应信息

### [示例​](#示例-2)

js
```js
// 遍历获取某个文件夹下的所有文件的文件名
for (let offset = 0; offset >= 0; ) {
  const list = KSDrive.listFiles({
    dirUrl: 'https://www.kdocs.cn/mine/xxxxxxxxxx',
    offset: offset,
    count: 100
  })
  for (let i = 0; i < list.files.length; i++) {
    console.log(list.files[i].fileName)
  }
  offset = list.nextOffset
}
```

### [参数​](#参数-2)

| 名称 | 类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| options | object | undefined | 否 | 一个 JavaScript 对象，undefined 时获取我的云文档目录下面的文件数据，详细参数如下所示 |

### [详细参数​](#详细参数)

| 参数名 | 参数类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| dirUrl | string |  | false | 目录链接，如https://www.kdocs.cn/mine/xxxxxx，为空时获取我的云文档目录下面的文件数据 |
| offset | number | 0 | false | 开始位置。通常由listFiles()函数返回。比如，listFiles()函数在某次检索中返回了 nextOffset 为 100，而想要获取更多文件信息，则下一次调用listFiles()函数时把 100 作为此可选参数传入。 |
| count | number | 30 | false | 文件个数 |
| includeExts | string[] |  | false | 指定文件类型,支持参数及对应关系，ksheet:"表格",et:"WPS 表格",db:"多维表",otl:"文档",wpp:"演示",wps:"WPS 文字" |

### [返回值​](#返回值-2)

FilesInfo- 一个 JavaScript 对象，文件信息

## [File​](#file)

打开文件函数openFile()返回的一个 JavaScript 对象。

### [属性​](#属性)

| 名称 | 类型 | 说明 |
| --- | --- | --- |
| Application | Application(ET/Ksheet/DBT) | 被打开文件的操作对象，目前支持 et,ksheet,dbt |
| close | Function | 关闭文件的函数，使用完 file 对象之后调用，关闭打开的文件 |

## [FilesInfo​](#filesinfo)

获取文件夹信息函数listFiles(options)返回的一个 JavaScript 对象。

### [属性​](#属性-1)

| 名称 | 类型 | 说明 |
| --- | --- | --- |
| files | FileInfo[] | 文件信息，详细参数如下所示 |
| nextOffset | number | 下一页的偏移量，可以作为listFiles(options)的参数而输出下一页文件内容，当下一页为空时，nextOffset 为-1 |

### [FileInfo​](#fileinfo)

| 名称 | 类型 | 说明 |
| --- | --- | --- |
| fileName | string | 文件名 |
| fileId | string | 加密后的文件 id |
| createTime | number | 文件创建时间戳 |
| updateTime | number | 文件修改时间戳 |


### 数据库 API

# [数据库 API​](#数据库-api)

AirScript 提供一个全局的 SQL 对象，开发者可通过此对象提供的属性和方法连接到外部数据库服务，连接成功后即可执行 SQL 语句，对数据进行增删改查。

提示

在使用 SQL 对象连接数据库之前，确保您已添加数据库API服务，在脚本编辑器的【工具栏】-【服务】菜单内添加即可。

### [快速使用​](#快速使用)

js
```js
// 连接MySQL数据库
const connection = SQL.connect(SQL.Drivers.MySQL, {
  host: '127.0.0.1',
  username: 'root',
  password: '123456',
  database: 'mydb',
  port: 3306
})

// 执行SQL语句，查询test表的所有数据
const result1 = connection.queryAll('SELECT * FROM test')
// 打印执行结果
console.log(result1)

// 执行SQL语句，插入数据
const result2 = connection.queryAll(
  'INSERT INTO test (id,test_data) VALUES (?,?), (?,?)',
  [1, 1, 2, 2]
)
// 打印执行结果
console.log(result2)

// 关闭数据库连接
connection.close()
```

### [属性列表​](#属性列表)

| 属性名 | 数据类型 | 说明 |
| --- | --- | --- |
| Drivers | object | 数据库连接驱动集 |
| Types | object | 数据库字段类型集（仅适用于 SQL server） |

### [方法列表​](#方法列表)

| 方法 | 返回类型 | 说明 |
| --- | --- | --- |
| connect() | Connection | 连接目标数据库 |
| Connection.queryAll() | Result | 执行 SQL 语句 |
| Connection.close() | null | 关闭数据库连接 |

## [Drivers​](#drivers)

数据库驱动集，调用connect()方法连接数据库时传入对应驱动，目前仅支持 MySQL 和 SQL server 两种驱动，只读

### [属性说明​](#属性说明)

| 属性名 | 数据类型 | 说明 |
| --- | --- | --- |
| MySQL | string | MySQL 数据库驱动 |
| PostgreSQL | string | PostgreSQL 数据库驱动 |
| SQLServer | string | SQL server 数据库驱动 |

## [Types​](#types)

数据库字段类型集，请注意，该类型集仅适用于 SQL server 数据库，MySQL 数据库不需要传递此值

### [属性说明​](#属性说明-1)

#### [Exact numerics​](#exact-numerics)

| 属性名 | 对应 Javascript 类型 |
| --- | --- |
| Bit | Boolean |
| TinyInt | Number |
| SmallInt | Number |
| Int | Number |
| BigInt | String |
| Numeric | Number |
| Decimal | Number |
| SmallMoney | Number |
| Money | Number |

#### [Approximate numerics​](#approximate-numerics)

| 属性名 | 对应 Javascript 类型 |
| --- | --- |
| Float | Number |
| Real | Number |

#### [Date and Time​](#date-and-time)

| 属性名 | 对应 Javascript 类型 |
| --- | --- |
| SmallDateTime | Date |
| DateTime | Date |
| DateTime2 | Date |
| DateTimeOffset | Date |
| Time | Date |
| Date | Date |

#### [Character Strings​](#character-strings)

| 属性名 | 对应 Javascript 类型 |
| --- | --- |
| Char | String |
| VarChar | String |
| Text | String |

#### [Unicode Strings​](#unicode-strings)

| 属性名 | 对应 Javascript 类型 |
| --- | --- |
| NChar | String |
| NVarChar | String |
| NText | String |

#### [Binary Strings​](#binary-strings)

| 属性名 | 对应 Javascript 类型 |
| --- | --- |
| Binary | Buffer |
| VarBinary | Buffer |
| Image | Buffer |

#### [Other Data Types​](#other-data-types)

| 属性名 | 对应 Javascript 类型 |
| --- | --- |
| Null | null |
| TVP | Object |
| UDT | Buffer |
| UniqueIdentifier | String |
| Variant | any |
| xml | String |

## [connect()​](#connect)

连接目标数据库，目前仅支持 MySQL、PostgreSQL 和 SQL server 三种类型的数据库，连接成功后会返回数据库连接对象，可通过此对象执行 SQL 语句，程序结束之前请调用close()方法关闭数据库连接。

### [参数​](#参数)

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| driver | Driver | null | 是 | 指定目标数据库驱动 |
| options | Options | null | 是 | 数据库连接信息 |

### [options 对象说明​](#options)

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| host | string | null | 是 | 目标数据库主机 |
| port | number | null | 是 | 目标数据库端口 |
| username | string | null | 是 | 目标数据库连接用户名 |
| password | string | null | 是 | 目标数据库连接密码 |
| database | string | null | 是 | 目标数据库名 |

### [返回值​](#返回值)

Connection - 数据库连接对象

### [示例​](#示例)

js
```js
// 连接MySQL数据库
const connection = SQL.connect(SQL.Drivers.MySQL, {
  host: '127.0.0.1',
  port: 3340,
  username: 'jinxiaomeng',
  password: '123',
  database: 'WPS_TEST'
})
```

## [Connection.queryAll()​](#queryall)

通过上述的connect()方法成功连接数据库后，会返回数据库连接对象，通过此对象即可调用 queryAll()方法执行 SQL 语句

### [参数​](#参数-1)

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| sql | string | null | 是 | 要执行的 sql 语句 |
| InsertData | any[] |InsertData | null | 否 | 需要插入的数据 |

### [InsertData 对象说明​](#insertdata)

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| name | string | null | 是 | 插入数据的字段名 |
| value | string | null | 是 | 插入数据的值 |
| type | Types | null | 否 | 插入数据的类型，SQL server 数据库必须传递该类型 |

### [返回值​](#返回值-1)

Result 对象，包含受影响的行数以及返回的数据行

| 属性 | 数据类型 | 说明 |
| --- | --- | --- |
| affectRowCount | number | 执行 sql 语句后受到影响的行数 |
| rows | Array | 数据行，根据实际查询的表的数据结构返回 |

### [返回示例​](#返回示例)

json
```json
// 查询时的返回
{
  "affectRowCount": 0,
  "rows": [
    [
      {
        "name": "1",
        "value": 2
      }
    ]
  ]
}

// 增删改时的返回
{
  "affectRowCount": 1,
  "rows": []
}
```

### [示例​](#示例-1)

js
```js
// 连接SQL server数据库
const connection = SQL.connect(SQL.Drivers.SQLServer, {
  host: 'x.x.x.x',
  username: 'x',
  password: 'x',
  database: 'x',
  port: 1433
})

// 执行sql语句，插入两条数据
const result1 = connection.queryAll(
  'INSERT INTO TestSchema.Employees (Name, Location) OUTPUT INSERTED.Id VALUES (@Name, @Location);',
  [
    {
      name: 'Name',
      type: SQL.Types.NVarChar,
      value: 'zhangsan'
    },
    {
      name: 'Location',
      type: SQL.Types.NVarChar,
      value: 'zhuhai'
    }
  ]
)

// 打印执行结果
console.log(result1)

// 执行sql语句，查询员工表
const result2 = connection.queryAll(
  'SELECT Id, Name, Location FROM TestSchema.Employees;'
)

// 打印执行结果
console.log(result2)

// 关闭数据库连接
connection.close()
```

## [Connection.close()​](#close)

关闭数据库连接，请务必在程序结束前调用此方法

### [示例​](#示例-2)

js
```js
// 连接MySQL数据库
const connection = SQL.connect(SQL.Drivers.MySQL, {
  host: '127.0.0.1',
  port: 3340,
  username: 'jinxiaomeng',
  password: '123',
  database: 'WPS_TEST'
})

// do something

// 关闭数据库连接
connection.close()
```


### 概述

# [概述​](#概述)

借助AirScript的高级服务，开发者只需要完成较少设置，即可连接到某些公开的金山文档API。 它们的使用方式与AirScript脚本的内置函数十分相似。

AirScript在运行时会自动处理授权流程。 不过开发者必须启用高级服务，才能在脚本中使用该服务，若跳过该步骤，会因为找不到该服务而抛出undefined错误。

## [启用高级服务​](#启用高级服务)

要使用高级服务，请按以下说明操作：

打开
效率
-
AirScript编辑工具
弹出编辑页面。
点击AirScript编辑工具上方的
服务
。
点击
添加服务
。
选择一项服务，然后点击
确认
。
启用高级服务后，该服务会在自动补全中显示。

## [授权流程​](#授权流程)

AirScript需要用户授权才能访问高级服务中的私密数据。

### [授予运行权限​](#授予运行权限)

AirScript会根据开发者编写脚本时启用高级服务的配置内容来确定授权范围 （例如访问指定文件或访问网络）。如果脚本需要授权，用户在运行脚本时会弹出授权对话框。 描述这个脚本涉及到的授权范围。

普通的代码更改并不会清空用户对脚本的授权。但如果开发者对更改了高级服务的配置（新增，修改或删除）， 那用户对脚本的授权也会清空，再次运行脚本时会重新触发授权流程。

注意:我的脚本中的脚本的所有权完全归属于用户本身，该分类运行脚本时无需触发授权流程。

### [取消授权​](#取消授权)

用户可以对已授权的脚本手动取消授权，请按以下说明操作

打开
效率
-
AirScript编辑工具
弹出编辑页面。
找到文件共享脚本下的想取消授权的脚本，点击
…
显示更多操作。
点击
取消服务授权
## [使用限制​](#使用限制)

为防止向用户提供恶意的脚本，出于安全性考虑，使用高级服务存在一些限制。

过于高频地使用高级服务，当出现这种情况时，脚本的运行会抛出明显的错误通知用户异常调用。
使用
HTTP
服务时，禁止使用IP地址发起请求，禁止使用端口发起请求。
使用
HTTP
服务时，收到内容的消息体最大为2M，超过2M会抛出错误。
使用
KSDrive.openFile
获得的
File
对象没有调用close, 就再次使用
KSDrive.openFile
会报错。

### 网络 API

# [网络 API​](#网络-api)

AirScript 提供一个全局的 HTTP 对象，开发者可通过此对象提供的方法请求外部服务，请求成功后会同步返回服务器的响应。

该 API 的使用方式与浏览器内的 fetch()函数基本一致，对于前端开发者来说应该可以很快上手。

提示

在使用 HTTP 对象提供的方法发送请求之前，确保您已添加网络API服务，在脚本编辑器的【工具栏】-【服务】菜单内添加即可。

### [快速使用​](#快速使用)

javascript
```javascript
// 发起网络请求
const resp = HTTP.fetch('https://open.iciba.com/dsapi/', {
  timeout: 2000
})
const data = resp.json()
console.log(data.note, data.content)
```

### [方法列表​](#方法列表)

| 方法 | 返回类型 | 简介 |
| --- | --- | --- |
| fetch(url[, options]) | Response | 发起自定义类型的网络请求 |
| get(url[, options]) | Response | 发起 GET 类型的网络请求 |
| delete(url[, options]) | Response | 发起 DELETE 类型的网络请求 |
| post(url,body[, options]) | Response | 发起 POST 类型的网络请求 |
| put(url,body[, options]) | Response | 发起 PUT 类型的网络请求 |

## [fetch(url[, options])​](#fetch)

发起一个网络请求，可以自定义设置 headers 和 body。

### [示例​](#示例)

javascript
```javascript
const resp = HTTP.fetch('https://www.kdocs.cn', {
  method: 'GET',
  timeout: 2000,
  headers: {
    'User-Agent':
      'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/109.0.0.0 Safari/537.36'
  }
})
console.log(resp.text())
```

### [参数​](#参数)

| 名称 | 类型 | 默认值 | 必填项 | 说明 |
| --- | --- | --- | --- | --- |
| url | string |  | true | 需要访问的网络地址，只允许访问不带端口号的域名 |
| options | RequestOption | undefined | false | 一个 JavaScript 对象，可指定发起请求的可选参数，如下所示。 |

### [RequestOption​](#requestoption)

| 名称 | 类型 | 默认值 | 必填项 | 说明 |
| --- | --- | --- | --- | --- |
| method | string | GET | false | 发起网络请求的方法，例如GET、POST、PUT、DELETE等 |
| timeout | number | 10000 | false | 发起网络请求的超时时间，单位毫秒(ms)，数据范围为 0~60000，超出范围的数据将被设为默认值 10 秒。 |
| headers | object | undefined | false | 发起网络请求的头部。例如cookie等 |
| body | string | undefined | false | 发起网络请求的主体内容。 |

### [返回值​](#返回值)

Response- 服务器返回的响应

## [get(url[, options])​](#get)

发起 GET 类型的网络请求。

### [示例​](#示例-1)

javascript
```javascript
const resp = HTTP.get('https://reqres.in/api/users/2')
console.log(resp.json())
```

### [参数​](#参数-1)

| 名称 | 类型 | 默认值 | 必填项 | 说明 |
| --- | --- | --- | --- | --- |
| url | string |  | true | 需要访问的网络地址，只允许访问不带端口号的域名 |
| options | MethodRequestOption | undefined | false | 一个 JavaScript 对象，可指定特定请求的可选参数，如下所示。 |

### [MethodRequestOption​](#methodrequestoption)

| 名称 | 类型 | 默认值 | 必填项 | 说明 |
| --- | --- | --- | --- | --- |
| timeout | number | 10000 | false | 发起网络请求的超时时间，单位毫秒(ms)，数据范围为 0~60000，超出范围的数据将被设为默认值 10 秒。 |
| headers | object | undefined | false | 发起网络请求的头部。例如cookie等 |

### [返回值​](#返回值-1)

Response- 服务器返回的响应

## [delete(url[, options])​](#delete)

发起 DELETE 类型的网络请求。

### [示例​](#示例-2)

javascript
```javascript
const resp = HTTP.delete('https://reqres.in/api/users/2')
console.log(resp.status)
```

### [参数​](#参数-2)

| 名称 | 类型 | 默认值 | 必填项 | 说明 |
| --- | --- | --- | --- | --- |
| url | string |  | true | 需要访问的网络地址，只允许访问不带端口号的域名 |
| options | MethodRequestOption | undefined | false | 一个 JavaScript 对象，可指定特定请求的可选参数，如下所示。 |

### [返回值​](#返回值-2)

Response- 服务器返回的响应

## [post(url,body[, options])​](#post)

发起 POST 类型的网络请求。

### [示例​](#示例-3)

javascript
```javascript
// 发送form
const formResp = HTTP.post(
  'https://www.example.cn',
  { foo: 'bar' },
  { headers: { 'content-type': 'multipart/form-data' } }
)

//发送json
const resp = HTTP.post('https://reqres.in/api/users', {
  name: 'morpheus',
  job: 'leader'
})

console.log(resp.json())
```

### [参数​](#参数-3)

| 名称 | 类型 | 默认值 | 必填项 | 说明 |
| --- | --- | --- | --- | --- |
| url | string |  | true | 需要访问的网络地址，只允许访问不带端口号的域名 |
| body | string| object |  | true | 请求体 |
| options | MethodRequestOption | undefined | false | 一个 JavaScript 对象，可指定特定请求的可选参数，如下所示。 |

### [返回值​](#返回值-3)

Response- 服务器返回的响应

## [put(url,body[, options])​](#put)

发起 PUT 类型的网络请求。

### [示例​](#示例-4)

javascript
```javascript
const resp = HTTP.put('https://reqres.in/api/users/200', {
  name: 'wps',
  job: 'developer'
})
console.log(resp.json())
```

### [参数​](#参数-4)

| 名称 | 类型 | 默认值 | 必填项 | 说明 |
| --- | --- | --- | --- | --- |
| url | string |  | true | 需要访问的网络地址，只允许访问不带端口号的域名 |
| body | string| object |  | true | 请求体 |
| options | MethodRequestOption | undefined | false | 一个 JavaScript 对象，可指定特定请求的可选参数，如下所示。 |

### [返回值​](#返回值-4)

Response- 服务器返回的响应

## [Response​](#response)

HTTP 发起网络请求后返回的响应，response 是流数据，只有首次调用 text()，json()或 binary()能获取到数据

### [示例​](#示例-5)

javascript
```javascript
let resp = HTTP.get('https://open.iciba.com/dsapi/')
console.log(resp.status) // 200
console.log(resp.statusText) // OK
console.log(resp.text()) // `{foo:"bar"}`
console.log(resp.json()) // {foo:"bar"}
console.log(resp.status) // [...]
```

### [方法列表​](#方法列表-1)

| 方法 | 返回类型 | 简介 |
| --- | --- | --- |
| status | number | 获取响应的 HTTP 状态码 |
| statusText | string | 获取响应的 HTTP 状态 |
| headers | object | 获取响应的 header |
| text() | string | 获取服务器返回的文本 Body |
| json() | any | 将服务器返回的 json 类型的 Body 转化为结构体 |
| binary() | Buffer | 获取服务器返回的二进制结构的 Body |

## [status​](#status)

获取响应的 HTTP 状态码

### [示例​](#示例-6)

javascript
```javascript
const resp = HTTP.get('https://open.iciba.com/dsapi/')
console.log(resp.status) // 200
```

### [返回值​](#返回值-5)

number - 服务器返回响应的 HTTP 状态码

## [statusText​](#statustext)

获取响应的 HTTP 状态

### [示例​](#示例-7)

javascript
```javascript
const resp = HTTP.get('https://open.iciba.com/dsapi/')
console.log(resp.statusText) // OK
```

### [返回值​](#返回值-6)

string - 服务器返回响应的 HTTP 状态

## [headers​](#headers)

获取响应的 header

### [示例​](#示例-8)

javascript
```javascript
let resp = HTTP.get('https://open.iciba.com/dsapi/')
console.log(resp.headers) // {"content-length":"44","content-type":"text/html; charset=utf-8"}
```

### [返回值​](#返回值-7)

object - 服务器返回响应的 header

## [text()​](#text)

获取服务器返回的 Body

### [示例​](#示例-9)

javascript
```javascript
let resp = HTTP.get('https://open.iciba.com/dsapi/')
console.log(resp.text()) // this is an example.
```

### [返回值​](#返回值-8)

string - 服务器返回的响应的 Body，以文本接受并返回

## [json()​](#json)

获取服务器返回的 Body

### [示例​](#示例-10)

javascript
```javascript
let resp = HTTP.get('https://open.iciba.com/dsapi/')
console.log(resp.json()) // {msg:"this is an example."}
```

### [返回值​](#返回值-9)

Object, Array, string, number, boolean, or null - 服务器返回的响应的 Body，以文本接受并经过 JSON.parse()后返回

## [binary()​](#binary)

获取服务器返回的 Body

### [示例​](#示例-11)

javascript
```javascript
let resp = HTTP.get('https://open.iciba.com/dsapi/')
console.log(resp.binary().toString('base64'))
```

### [返回值​](#返回值-10)

Buffer- 服务器返回的响应的 Body，以 Buffer 接受二进制数据并返回


### 邮件 API

# [邮件 API​](#邮件-api)

通过外部邮件服务发送邮件。

### [快速使用​](#快速使用)

javascript
```javascript
// 登录
let mailer = SMTP.login({
    host: "smtp.example.com", // 域名
    port: 465, // 端口
    secure: true, // TLS
    username: "sender@example.com", // 账户名
    password: "Pa55W0rd" // 密码
})
// 客户端发送邮件
mailer.send({
    from: "sender@example.com", // 发件人
    to: "reciever@example.com", // 收件人
    subject: "this is subject.", // 主题
    text: "this is text.", // 文本
    html: `<p> this is html </p>` // HTML代码
})
// 支持指定昵称
mailer.send({
    from: "管理员 <admin@example.com>",
    to: "接受者 <username@example.com>",
    subject: "this is subject.",
    text: "this is text.",
    html: `<p> this is html </p>`
})
// 支持发送多个邮箱
mailer.send({
    from: "管理员 <admin@example.com>",
    to: ["username1@example.com","接受者2 <username2@example.com>"],
    subject: "this is subject.",
    text: "this is text.",
    html: `<p> this is html </p>`
})
```

## [SMTP​](#smtp)

### [方法列表​](#方法列表)

| 方法 | 返回类型 | 简介 |
| --- | --- | --- |
| login(argvs) | Mailer | 登录并返回邮件发送者 |

## [login(argvs)​](#login-argvs)

登录并返回邮件发送对象

javascript
```javascript
//  登录qq邮箱
let mailer = SMTP.login({
    host: "smtp.qq.com", // QQ 的SMTP服务器的域名
    port: 465,
    username: "1000000000@qq.com", // qq 邮箱地址
    password: "xxxxxxxxxxxx", // qq邮箱的SMTP密码，非qq密码
    secure: true
});
```

### [参数​](#参数)

| 名称 | 类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| argvs | LoginArgvs | undefined | true | 一个JavaScript对象，用于配置SMTP的参数，如下所示 |

### [LoginArgvs​](#loginargvs)

| 名称 | 类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| host | string | undefined | true | 邮箱服务器域名 |
| port | number | undefined | true | SMTP服务端口，当host为undefined时，取默认值，默认值由secure决定，当secure是false时默认值为587，当secure是true时默认值为465。 |
| secure | boolean | undefined | true | 是否使用TLS连接服务器，在大多数情况下，如果要连接到端口465，请将此值设置为true；如果要连接到端口587或25，请将此值设置为false。 |
| username | string | undefined | true | 用于身份验证的账户名 |
| password | string | undefined | true | 用于身份验证的密码 |
| timeout | number | 10000 | false | 等待建立连接的时间，单位毫秒(ms) |

### [返回值​](#返回值)

Mailer- 邮件发送者

## [Mailer​](#mailer)

由login(argvs)创建的对象，用于发送邮件

### [方法列表​](#方法列表-1)

| 方法 | 返回类型 | 简介 |
| --- | --- | --- |
| send(message) | undefined | 发送邮件 |

## [send(message)​](#send-message)

发送邮件

javascript
```javascript
mailer.send({
    from: ["Administrator <admin@example.com>"],
    to: ["username@example.com", "UserName <username2@example.com>"],
    subject: "this is subject.",
    text: "this is text.",
    html: `<p> this is html </p>`
})
```

### [参数​](#参数-1)

| 名称 | 类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| message | messageArgvs | undefined | true | 一个JavaScript对象，要发送的邮件内容，如下所示 |

### [messageArgvs​](#messageargvs)

| 名称 | 类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| from | string | undefined | true | 发件人的电子邮箱地址 |
| to | string / string[] | undefined | true | 收件人的电子邮箱地址 |
| subject | string | undefined | true | 电子邮件的主题 |
| text | string | undefined | true | 电子邮件显示的文本 |
| html | string | undefined | false | 电子邮件的HTML代码 |


# API文档(2.0)

## 概述

# [什么是AirScript 2.0​](#什么是airscript-2-0)

AirScript 1.0脚本能够满足常规的脚本操作文档需求。但在API的丰富度、执行性能、JavaScript高级语法支持上仍然存在一些限制。为了突破这些局限，我们推出了AirScript 2.0版本。

AirScript 2.0脚本采用全新的架构设计，在功能和性能上带来大幅升级，API完全兼容WPS JS宏，执行速度更快，并全面支持现代JavaScript语法。

尽管AirScript 1.0大部分API可以在AirScript 2.0兼容运行，但仍然存在一些API定义存在冲突的问题导致无法实现完全兼容，因此需要通过1.02.0来标识脚本的运行时环境版本，AirScript 2.0是AirScript的一个大版本升级，部分API用法与AirScript 1.0无法兼容。

# [AirScript 2.0有哪些亮点​](#airscript-2-0有哪些亮点)

### [✔️完全兼容WPS JS宏​](#✔️完全兼容wps-js宏)

AirScript 2.0完全兼容WPS JS宏，提供了2000+API（比1.0多4倍），极大地扩展了您的开发能力。无论您是进行简单的自动化任务，还是复杂的脚本开发，AirScript 2.0都能满足您的需求。无需重新学习新的API定义，即可在WPS端和在线表格使用同样的API开发脚本。

### [🚀更快的执行速度​](#🚀更快的执行速度)

我们对AirScript 2.0的逻辑执行和API调用速度进行了优化，使您的脚本运行更加高效，响应更加迅速。对于需要频繁调用API的大型脚本项目，这种性能提升尤为显著。如：循环对1000个单元格文本去除首尾空格并写回，2.0比1.0速度快1倍。

可以分别新建1.0和2.0脚本，执行以下代码体验2.0的性能提升

javascript
```javascript
// 对当前激活sheet第一列前1k个格子去除首尾空格，并打印脚本执行耗时
const start = Date.now();

const sheet = ActiveSheet;
for (let i = 1; i <= 1000; i++) {
  const cell = sheet.Cells(i, 1);
  cell.Value2 = cell.Text.trim();
}

console.log(Date.now() - start);
```

### [🌐全面支持现代 JavaScript 语法​](#🌐全面支持现代-javascript-语法)

AirScript 2.0全面支持现代 JavaScript 语法，如：await/async、Promise、class等语法（1.0不支持），让您的开发体验更加现代和高效。特别对于复杂的脚本项目，代码编写维护更方便。

# [AirScript 2.0支持哪些场景​](#airscript-2-0支持哪些场景)

AirScript 2.0设计上能够满足所有AirScript 1.0具备的能力，目前一期开放支持的是工作表场景，数据表的API待开放。若您的脚本需要处理数据表场景，您可以先继续使用AirScript 1.0脚本，后续2.0我们将逐步开放支持数据表的API。

当前AirScript 2.0处于beta版，部分API能力将在后续迭代中持续补齐。

|  | AirScript 1.0 | AirScript 2.0 |
| --- | --- | --- |
| 工作表 | ✅API（500+） | ✅API（2000+） |
| 数据表 | ✅API | 🏗️API（待开放） |
| 高级服务 | ✅网络API✅云文档API✅邮件API✅数据库API | ✅网络API✅云文档API🏗️邮件API（待开放）🏗️数据库API（待开放） |

# [如何使用AirScript 2.0​](#如何使用airscript-2-0)

多维表格由于涉及数据表支持，待开放

### [如何新建2.0脚本​](#如何新建2-0脚本)

1.打开AirScript脚本编辑器

效率->高级开发->AirScript脚本编辑器

2.新建脚本（文档共享脚本/我的脚本）

直接点击+即可新建2.0脚本

也可以通过下拉菜单选择新建脚本的版本

“我的脚本”操作路径相同

### [如何查阅2.0文档​](#如何查阅2-0文档)

API文档分为1.0和2.0版本，不同版本API存在差异，2.0文档为：API文档（2.0）。

### [如何判断脚本是1.0还是2.0​](#如何判断脚本是1-0还是2-0)

脚本列表中选中对应的脚本，可以查看到脚本的版本标识。1.0表示1.0版本，Beta表示2.0版本。

# [使用中可能遇到的问题​](#使用中可能遇到的问题)

### [部分API不兼容问题​](#部分api不兼容问题)

由于AirScript 2.0的API定义与WPS JS宏规范对齐，AirScript 1.0的API定义虽然参考了WPS JS宏的规范但未完全对齐，因此1.0和2.0的一小部分API参数、返回值类型、用法等存在定义冲突无法实现完全兼容，当您遇到调用的API不存在等错误时，您需要检查当前的脚本运行时环境版本以及使用的API定义版本是否一致，并根据环境和对应API文档进行API用法调整。

### [API调用异常​](#api调用异常)

AirScript 1.0 2.0API存在版本差异，您可先确认脚本和API文档版本是否一致，可以帮您快速定位问题。

如确认API执行存在非预期的结果，您可在AirSheet用户社区 (wps.cn)进行反馈，若能同时提供复现问题脚本和结果截图，将会利于我们尽快修复问题。

### [AirScript 1.0存量脚本是否可以继续运行？​](#airscript-1-0存量脚本是否可以继续运行)

AirScript 1.0仍然可以正常运行，我们通过新的2.0版本来提供大版本功能升级，并保留1.0运行时环境来确保您的原有脚本不受影响，1.0 2.0版本是互相隔离的。后续我们会推出1.0升级2.0的指引。

# [性能优化Tips​](#性能优化tips)

尽管AirScript 2.0在性能上有显著提升，但在代码计算量特别大时仍可能遇到性能瓶颈，您可通过最佳实践 | AirScript文档 (wps.cn)了解如何进行优化。


## 智能表格

### 工作表

#### AboveAverage 对象

# [AboveAverage (对象)​](#aboveaverage-对象)

代表条件格式规则的高于平均值的视图。对某一区域或选定内容应用颜色或填充有助于您查看与其他单元格相关的单元格的值。

## [说明​](#说明)

所有条件格式对象均包含在FormatConditions集合对象中，该集合对象是Range集合的子项。您可以使用FormatConditions集合的Add或AddAboveAverage方法创建高于平均值格式规则。

## [示例​](#示例)

javascript
```javascript
/*本示例通过条件格式规则生成一个动态数据集并对高于平均值的值应用颜色。*/
function test() {
    //Building data for Melanie
    Range("A1").Value2 = "Name"
    Range("B1").Value2 = "Number"
    Range("A2").Value2 = "Melanie-1"
    Range("A2").AutoFill(Range("A2:A26"), xlFillDefault)
    Range("B2:B26").FormulaArray = "=INT(RAND()*101)"
    Range("B2:B26").Select()

    //Applying Conditional Formatting to items above the average.  Should appear green fill and dark green font.
    Selection.FormatConditions.AddAboveAverage()
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority()
    Selection.FormatConditions(1).AboveBelow = xlAboveAverage
    let font = Selection.FormatConditions(1).Font
    font.Color = RGB(0, 155, 115)
    font.TintAndShade = 0
    let interior = Selection.FormatConditions(1).Interior
    interior.PatternColorIndex = xlAutomatic
    interior.Color = RGB(5, 185, 115)
    interior.TintAndShade = 0
    console.log("Added an Above Average Conditional Format to Melanie's data.  Press F9 to update values.")
}
```

javascript
```javascript
/*本示例设置工作表 Sheet1 上区域 C1:C10 中第一个（AboveAverage）条件格式的AboveBelow属性，并设置该条件格式的内部颜色。*/
function test() {
    let aboveAverage = Application.Worksheets.Item("Sheet1").Range("C1:C10").FormatConditions.Item(1)
    aboveAverage.AboveBelow = xlEqualBelowAverage
    aboveAverage.Interior.ColorIndex = 10
}
```


#### Adjustments 对象

# [Adjustments (对象)​](#adjustments-对象)

它包含指定的自选图形、艺术字对象或连接符的调整值的集合。

## [说明​](#说明)

每个调整值代表一种调整控点的调整方法。由于某些调整控点可以按两种方法调整（例如，某些控点既可以水平调整也可以垂直调整），所以形状的调整值数量可以大于调整控点数量。一个形状最多可以有八个调整值。

使用Adjustments属性可返回Adjustments对象。使用Adjustments(index)（其中index是调整值的索引号）可返回单个调整值。

不同的形状具有不同数目的调整值，不同类型的调整值在不同的方向上调整形状的几何性质，不同类型的调整值有不同的取值范围。例如，下面的图示显示了右箭头标注的四个调整值各对该标注的几何形状起什么作用。

| 注释 |
| --- |
| 由于每个形状有不同的调整值集，校验指定形状的调整行为的最好方法是手动创建一个图例，在打开宏记录器的情况下作调整，然后检查记录的代码。 |

下表概括了不同类型的调整值的有效取值范围。多数情况下，如果指定的调整值超出了有效取值范围，就将用最接近的有效值来代替。

| 调整类型 | 有效值 |
| --- | --- |
| 线性（水平或垂直） | 通常 0.0 值代表形状的左边界或上边界，而 1.0 值代表形状的右边界或下边界。有效值对应于有效的手动调整。例如，如果只能将调整控点手动拖动形状的一半宽度，则相应的调整值最大为 0.5。对于象连接符和标注这样的形状，0.0 和 1.0 值代表由它们的起始和终止点定义的矩形界限，此时负值和大于 1.0 的值是有效的。 |
| 射线图 | 调整值 1.0 对应于形状宽度。最大值为 0.5，或形状宽度的一半。 |
| 角 | 值以度表示。如果指定的值超过了 -180 到 180 的范围，则将其折算为该范围内的值。 |


#### AllowEditRange 对象

# [AllowEditRange (对象)​](#alloweditrange-对象)

代表受保护的工作表上可进行编辑的单元格。

## [说明​](#说明)

使用AllowEditRanges集合的Add方法或Item属性可返回AllowEditRange对象。

返回AllowEditRange对象后，可使用ChangePassword方法更改密码以访问可在受保护的工作表上编辑的区域。

## [示例​](#示例)

javascript
```javascript
/*在此示例中，ET 允许编辑活动工作表上的 A1:A4 范围，通知用户更改此指定区域的密码，然后通知用户更改成功。在运行此代码之前，工作表必须未受到保护。*/
function test() {
    let sheet = Application.ActiveSheet
    sheet.Unprotect()
    let wksPassword = "Enter password for the worksheet"

    //Establish a range that can allow edits on the protected worksheet.
    sheet.Protection.AllowEditRanges.Add("Classified", Range("A1:A4"), wksPassword)
    console.log("Cells A1 to A4 can be edited on the protected worksheet.")

    //Change the password.
    wksPassword = "Enter the new password for the worksheet"
    sheet.Protection.AllowEditRanges("Classified").ChangePassword(wksPassword)

    console.log("The password for these cells has been changed.")
}
```

javascript
```javascript
/*本示例将活动工作表上第一个可编辑的单元格区域删除。*/
function test() {
    ActiveSheet.Protection.AllowEditRanges.Item(1).Delete()
}
```


#### AllowEditRanges 对象

# [AllowEditRanges (对象)​](#alloweditranges-对象)

所有 AllowEditRange对象的集合，这些对象代表受保护工作表上的可编辑单元格。

## [说明​](#说明)

使用Protection对象的AllowEditRanges属性可返回AllowEditRanges集合。

返回AllowEditRanges集合后，可以使用Add方法添加可在受保护的工作表上编辑的区域。

## [示例​](#示例)

javascript
```javascript
/*在本示例中，ET 允许用户编辑活动工作表上的区域 A1:A4，并将指定区域的标题和地址通知用户。*/
function test() {
    let sheet = Application.ActiveSheet

    // Unprotect worksheet.
    sheet.Unprotect()
    wksPassword = "Enter password for the worksheet"

    // Establish a range that can allow edits on the protected worksheet.
    sheet.Protection.AllowEditRanges.Add("Classified", Range("A1:A4"), wksPassword)

    //Notify the user the title and address of the range.
    let allowEditRange = sheet.Protection.AllowEditRanges.Item(1)

    console.log(`Title of range: ${allowEditRange.Title}`)
    console.log(`Address of range: ${allowEditRange.Range.Address()}`)
}
```

javascript
```javascript
/*本示例显示第一张工作表上是否存在可编辑单元格。*/
function test() {
    console.log(Worksheets.Item(1).Protection.AllowEditRanges.Count > 0)
}
```


#### Application 对象

# [Application (对象)​](#application-对象)

代表整个 ET 应用程序，它是整个应用程序api对象树的根对象。

## [说明​](#说明)

整个et应用程序的api对象结构是一个树状结构，而Application对象是树状结构的根对象，同时，Application对象也提供了对于应用程序相关的各种访问接口，例如：

应用程序的设置选项、环境、版本号等相关信息
一些常见的属性，如
ActiveCell
、
ActiveWorkbook
、
ActiveSheet
等

#### Areas 对象

# [Areas (对象)​](#areas-对象)

由选定区域内的多个子区域或连续单元格块组成的集合。


#### AutoFilter 对象

# [AutoFilter (对象)​](#autofilter-对象)

代表对指定工作表的自动筛选。

## [说明​](#说明)

使用AutoFilter属性可返回AutoFilter对象。使用Filters属性可返回由各个列筛选组成的集合。使用Range属性可返回代表整个筛选区域的Range对象。

要为工作表创建AutoFilter对象，必须手动打开或使用Range对象的AutoFilter方法打开工作表上某个区域上的自动筛选功能。

| 注释 |
| --- |
| When usingAutoFilterwith dates, the format should be consistent with English date separators ("/") instead of local settings ("."). A valid date would be "2/2/2007", whereas "2.2.2007" is invalid. |


#### Axes 对象

# [Axes (对象)​](#axes-对象)

指定图表中所有Axis对象的集合。

## [示例​](#示例)

javascript
```javascript
/*本示例为图表工作表 Chart1 中分类轴设置标题文本。*/
function test() {
    let axis = Application.Charts.Item("Chart1").ChartObjects(1).Chart.Axes().Item(xlValue)
    axis.AxisTitle.Caption = "成绩"
}
```

javascript
```javascript
/*本示例显示工作表 Sheet1 中第一张图表的坐标轴的数量是否为2。*/
function test() {
    let axes = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart.Axes()
    console.log(axes.Count == 2)
}
```


#### Axis 对象

# [Axis (对象)​](#axis-对象)

代表图表中的单个坐标轴。

## [说明​](#说明)

Axis对象是Axes集合的成员。

使用Axes(type,group)（其中type为坐标轴类型，而group为坐标轴组）可返回单个Axis对象。Type可为以下XlAxisType常量之一：xlCategory、xlSeries或xlValue。Group可为以下XlAxisGroup常量之一：xlPrimary或xlSecondary。有关详细信息，请参阅Axes方法。

## [示例​](#示例)

javascript
```javascript
/*本示例在名为“Chart1”的图表工作表中设置分类轴的标题文本。*/
function test() {
    let axis = Application.Charts.Item("Chart1").ChartObjects(1).Chart.Axes(xlCategory)
    axis.HasTitle = true
    axis.AxisTitle.Caption = "1994"
}
```

javascript
```javascript
/*本示例将工作表 Sheet1 中第一张图表的数值轴的主要刻度线设置为跨轴，并将其颜色设置为红色。*/
function test() {
    let axis = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart.Axes(xlValue)
    axis.MajorTickMark = xlTickMarkCross
    axis.Border.ColorIndex = 3
}
```


#### AxisTitle 对象

# [AxisTitle (对象)​](#axistitle-对象)

代表图表坐标轴标题。

## [说明​](#说明)

使用AxisTitle属性可返回AxisTitle对象。

只有当坐标轴的HasTitle属性为True时，AxisTitle对象才存在，从而才能使用该对象。

## [示例​](#示例)

javascript
```javascript
/*下例激活第一个嵌入式图表，设置其数值轴标题文本，将其字体设为 10 磅的“Bookman”，并将单词“millions”设为倾斜。*/
function test() {
    Application.Worksheets.Item("Sheet1").ChartObjects(1).Activate()
    let axes = Application.ActiveChart.Axes(xlValue)
    axes.HasTitle = true
    let axistitle = axes.AxisTitle
    axistitle.Caption = "Revenue (millions)"
    axistitle.Font.Name = "bookman"
    axistitle.Font.Size = 10
    axistitle.Characters(10, 8).Font.Italic = true
}
```

javascript
```javascript
/*本示例将图表工作表 Chart1 中图表的分类轴标题设置为“考核分类”，并将该标题设置为加粗。*/
function test() {
    let axistitle = Application.Charts.Item("Chart1").ChartObjects(1).Chart.Axes(xlCategory).AxisTitle
    axistitle.Caption = "考核分类"
    axistitle.Characters().Font.Bold = true
}
```


#### Border 对象

# [Border (对象)​](#border-对象)

代表对象的边框。

## [说明​](#说明)

大多数具有边框的对象（除Range和Style对象外）都将边框作为单一实体处理，而不管边框有几个边。整个边框必须作为一个整体单位返回。例如，使用TrendLine对象的Border属性可返回此类对象的Border对象。

## [示例​](#示例)

javascript
```javascript
/*本示例更改活动图表中趋势线的类型和线型。*/
function test() {
    let trendline = Application.ActiveChart.SeriesCollection(1).Trendlines(1)
    trendline.Type = xlLinear
    trendline.Border.LineStyle = xlDash
}
```

Range和Style对象具有四个分立的边框：左边框、右边框、顶部边框和底部边框，这四个边框既可单独返回，也可作为一个组同时返回。使用Borders属性可返回Borders集合，该集合包含所有四个边框，并将这些边框视为一个单位。

javascript
```javascript
/*本示例向第一张工作表上的单元格 A1 右边缘的边框添加双边框。*/
function test() {
    Application.Worksheets.Item(1).Range("A1").Borders.Item(xlEdgeRight).LineStyle = xlDouble
}
```

使用Borders(index)（其中index指定边框）可返回单个Border对象。

javascript
```javascript
/*本示例设置单元格区域 A1:G1 的底部边框的颜色。*/
function test() {
    Range("A1:G1").Borders.Item(xlEdgeBottom).Color = RGB(255, 0, 0)
}
```

javascript
```javascript
/*本示例将边框的 Weight 属性设置为 xlThick 会诱使 LineStyle 属性变为 xlSolid，尽管之前已将其设置为 xlDashDotDot 。*/
function test() {
    let border = Selection.Borders.Item(xlDiagonalDown)
    border.Color = RGB(255, 0, 0)
    console.log("border.LineStyle = " + border.LineStyle)  //border.style = 1
    console.log("Set border.style = xlDashDotDot")  //Set border.style = DashDotDot
    border.LineStyle  = xlDashDotDot
    console.log("border.LineStyle = " + border.LineStyle)  //border.style = 5
    console.log("Set border.weight = xlThick")  //Set border.weight = Thick
    border.weight = xlThick
    console.log("border.LineStyle = " + border.LineStyle)  //border.style = 1
}
```

Index可为以下 XlBordersIndex 常量之一：xlDiagonalDown、xlDiagonalUp、xlEdgeBottom、xlEdgeLeft、xlEdgeRight、xlEdgeTop、xlInsideHorizontal或xlInsideVertical。


#### Borders 对象

# [Borders (对象)​](#borders-对象)

由四个Border对象组成的集合，它们分别代表Range或Style对象的四个边框。

## [说明​](#说明)

使用Borders属性可返回包含所有四个边框的Borders集合。 可以对单元格或区域的每一侧应用不同的边框。 有关如何对单元格区域应用边框的详细信息，请参阅Range.Borders属性。

只能对Range和Style对象的各个边框分别设置边框属性。 其他带边框的对象（如误差线和序列线）具有被视为单个实体的边框，而不管它有多少边。 对于这些对象，在返回边框和设置边框属性时必须将其作为一个整体处理。 有关详细信息，请参阅Border对象。

## [示例​](#示例)

javascript
```javascript
/*本示例向第一张工作表上的单元格 A1 添加双边框。*/
function test() {
    Worksheets.Item(1).Range("A1").Borders.LineStyle = xlDouble
}
```

使用Borders(索引) （其中 index 标识边框）返回单个Border对象。 Index 可以是以下XlBordersIndex常量之一：xlDiagonalDown、xlDiagonalUp、xlEdgeBottom、xlEdgeLeft、xlEdgeRight、xlEdgeTop、xlInsideHorizontal或xlInsideVertical。

javascript
```javascript
/*本示例将工作表 Sheet1 中单元格区域 A1:G1 的底部边框颜色设置为红色。*/
function test() {
    Worksheets("Sheet1").Range("A1:G1").Borders.Item(xlEdgeBottom).Color = RGB(255, 0, 0)
}
```

javascript
```javascript
/*本示例在区域中的所有单元格周围生成一个细边框。*/
function test() {
    let borders = Range("B6:D8").Borders
    borders.LineStyle = xlContinuous
    borders.Weight = xlThin
}
```

javascript
```javascript
/*本示例更改区域的内部单元格边框。*/
function test() {
    let borders = Range("B2:D4").Borders
    borders.Item(xlInsideHorizontal).LineStyle = xlContinuous
    borders.Item(xlInsideHorizontal).Weight = xlThin
    borders.Item(xlInsideVertical).LineStyle = xlContinuous
    borders.Item(xlInsideVertical).Weight = xlThin
}
```


#### CalculatedFields 对象

# [CalculatedFields (对象)​](#calculatedfields-对象)

由PivotField对象组成的集合，这些对象代表指定数据透视表中的所有计算字段。

## [示例​](#示例)

javascript
```javascript
/*此示例为活动工作表中数据透视表添加计算字段。*/
function test() {
    ActiveSheet.Range("I1").PivotTable.CalculatedFields().Add("new date", "= date + 10")
}
```

javascript
```javascript
/*此示例显示第一张工作表中第一张数据透视表的第一个计算字段的标签文本。*/
function test() {
    let pvtField = Application.Worksheets.Item(1).PivotTables(1).CalculatedFields().Item(1)
    console.log(pvtField.Caption)
}
```


#### CalculatedItems 对象

# [CalculatedItems (对象)​](#calculateditems-对象)

由PivotItem对象组成的集合，这些对象代表指定数据透视表中的所有计算项。

## [示例​](#示例)

javascript
```javascript
/*本示例为工作表 Sheet1 上数据透视表的第四个字段添加计算项。*/
function test() {
    let pvtField = Worksheets.Item("Sheet1").Range("I1").PivotTable.PivotFields(4)
    pvtField.CalculatedItems().Add("公式1", "='21'+1")
}
```

javascript
```javascript
/*本示例显示工作表 Sheet1 中第一张数据透视表的字段“name”的“公式2”计算项的位置。*/
function test() {
    let pvtItem = Worksheets.Item("Sheet1").PivotTables(1).PivotFields("name").CalculatedItems().Item("公式2")
    console.log(pvtItem.Position)
}
```


#### CellFormat 对象

# [CellFormat (对象)​](#cellformat-对象)

代表单元格格式的搜索条件。

## [说明​](#说明)

使用Application对象的FindFormat或ReplaceFormat属性可返回CellFormat对象。

使用CellFormat对象的Borders、Font属性或CellFormat对象的Interior属性可定义单元格格式的搜索条件。

## [示例​](#示例)

javascript
```javascript
/*本示例设置单元格格式内部的搜索条件。*/
function test() {
    // Set the interior of cell A1 to yellow.
    Application.Range("A1").Select()
    Application.Selection.Interior.ColorIndex = 36
    console.log("The cell format for cell A1 is a yellow interior.")

    // Set the CellFormat object to replace yellow with green.
    Application.FindFormat.Interior.ColorIndex = 36
    Application.ReplaceFormat.Interior.ColorIndex = 35

    // Find and replace cell A1's yellow interior with green.
    ActiveCell.Replace("", "", xlPart, xlByRows, false, null, true, true)
    console.log("The cell format for cell A1 is replaced with a green interior.")
}
```

javascript
```javascript
/*本示例将替换条件中单元格边框的粗细设置为 xlThick 。*/
function test() {
    Application.ReplaceFormat.Borders.Weight = xlThick
}
```


#### Characters 对象

# [Characters (对象)​](#characters-对象)

代表包含文本的对象中的字符。

## [说明​](#说明)

使用Characters对象可修改包含在全文本字符串中的任意字符序列。

使用Characters(start,length)（其中start为起始字符号，而length为要返回的字符个数）返回Characters对象。

javascript
```javascript
/*下例向单元格 B1 中添加文本，并将第二个单词设置为加粗。*/
function test()
{
    Application.Worksheets.Item("Sheet1").Range("B1").Value = "New Title"
    Application.Worksheets.Item("Sheet1").Range("B1").Characters(5, 5).Font.Bold = true
}
```

仅当需要更改对象中文本的一部分而不影响其余部分时，才有必要使用Characters方法（如果对象不支持格式文本，则不能使用Characters方法对文本中的一部分单独设置格式）。要同时更改所有文本，通常可以对该对象直接应用某一适当的方法或属性。

javascript
```javascript
/*下例将单元格 A5 的内容设置为倾斜。*/
Application.Worksheets.Item("Sheet1").Range("A5").Font.Italic = true
```


#### Chart 对象

# [Chart (对象)​](#chart-对象)

代表工作簿中的图表。

## [说明​](#说明)

此图表既可以是嵌入的图表（包含在ChartObject对象中），也可以是单独的图表工作表。

示例部分中描述了以下用于返回Chart对象的属性和方法：

Charts
方法
ActiveChart
属性
ActiveSheet
属性
Charts集合包含工作簿中每个图表工作表的Chart对象。使用Charts(index) 可以返回单个Chart对象，其中 index 为图表工作表的索引号或名称。图表索引号表示图表工作表在工作簿标签栏上的位置。Charts(1)是工作簿中第一个（最左边的）图表；Charts(Charts.Count)是最后一个（最右边的）图表。所有图表工作表均包括在索引计数中，即便是隐藏工作表也是如此。图表工作表名称显示在图表工作簿标签上。您可以使用Name属性设置或返回图表名称。

## [示例​](#示例)

javascript
```javascript
/*本示例将图表工作表 Chart1 第一个数据系列设置为红色。*/
function test() {
    let chart = Application.Charts.Item("Chart1").ChartObjects(1).Chart
    chart.SeriesCollection(1).Format.Fill.ForeColor.RGB = rgbRed
}
```

javascript
```javascript
/*本示例将图表工作表 Sales 移至活动工作簿的尾部。*/
function test() {
    let chart = Application.Charts.Item("Sales")
    chart.Move(null, Sheets.Item(Sheets.Count))
}
```

Chart对象也是Sheets集合的成员，此集合包含工作簿中的所有工作表（图表工作表和工作表）。使用Sheets(index) 可以返回单个工作表，其中index是工作表索引号或名称。

当图表是活动对象时，您可以使用ActiveChart属性引用它。如果用户选择了图表工作表，或者用Chart对象的Activate方法或ChartObject对象的Activate方法激活了它，则该图表工作表处于活动状态。

javascript
```javascript
/*本示例将图表工作表 Chart1 图表类型修改为折线图，并将图表标题修改为“January Sales”。*/
function test() {
    let chart = Application.Charts.Item("Chart1").ChartObjects(1).Chart
    chart.ChartType = xlLine
    chart.HasTitle = true
    chart.ChartTitle.Text = "January Sales"
}
```

如果用户选择了嵌入图表，或者用Activate方法激活了包含该嵌入图表的ChartObject对象，则该嵌入图表处于活动状态。通过使用ActiveChart属性，您可以编写能够引用嵌入图表或图表工作表（视哪一个处于活动状态而定）的 JavaScript 代码。

javascript
```javascript
/*本示例将工作表 Sheet1 中第一个内嵌图表的图表类型修改为折线图，然后将图表标题修改为“January Sales”。*/
function test() {
    let chart = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart
    chart.ChartType = xlLine
    chart.HasTitle = true
    chart.ChartTitle.Text = "January Sales"
}
```

当图表工作表为活动工作表时，可以使用ActiveSheet属性来引用它。

javascript
```javascript
/*本示例将 Chart1 图表工作表中系列 1 的内部颜色设置为蓝色。*/
function test() {
    let chart = Application.Charts.Item("chart1").ChartObjects(1).Chart
    chart.SeriesCollection(1).Format.Fill.ForeColor.RGB = rgbBlue
}
```


#### ChartArea 对象

# [ChartArea (对象)​](#chartarea-对象)

代表图表的图表区。

## [说明​](#说明)

图表区包含绘图区在内的一切内容。但是，绘图区具有它自己的填充方式，因此填充绘图区并不会填充图表区。

有关设置绘图区格式的信息，请参阅PlotArea 对象。

使用ChartArea属性可返回ChartArea对象。

## [示例​](#示例)

javascript
```javascript
/*本示例关闭 Sheet1 工作表上第一个嵌入图表的图表区边框。*/
function test() {
    let chartarea = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.ChartArea
    chartarea.Format.Line.Visible = false
}
```

javascript
```javascript
/*本示例将工作表 Chart1 图表区的内部颜色设置为蓝色。*/
function test() {
    let chartarea = Application.Charts.Item("Chart1").ChartObjects(1).Chart.ChartArea
    chartarea.Format.Fill.ForeColor.RGB = rgbBlue
}
```


#### ChartCategory 对象

# [ChartCategory (对象)​](#chartcategory-对象)

指定图表类型的类别。

## [示例​](#示例)

javascript
```javascript
/*此示例判断工作表 Sheet1 第一个内嵌图表第一个图表组第一个类别是否筛选出序列，并通知用户。*/
function test() {
    let chart = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.ChartGroups(1)
    if (chart.CategoryCollection(1).IsFiltered) {
        console.log("图表组的第一个类别筛选出序列")
    } else {
        console.log("图表组的第一个类别未筛选出序列")
    }
}
```

javascript
```javascript
/*此示例显示工作表 Sheet1 第一个内嵌图表第一个图表组的第四个类别的实际类型。*/
function test() {
    let chart = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.ChartGroups(1)
    let category = chart.FullCategoryCollection(4)
    console.log(category.constructor.name)
}
```


#### ChartFormat 对象

# [ChartFormat (对象)​](#chartformat-对象)

提供对图表元素艺术字格式的访问。

## [说明​](#说明)

如果使用的属性或方法不适用于ChartFormat对象所附加到的对象的类型，则会产生运行时错误。


#### ChartGroup 对象

# [ChartGroup (对象)​](#chartgroup-对象)

代表图表中用同一格式绘制的一个或多个数据系列。

## [说明​](#说明)

一张图表包含一个或多个图表组，每个图表组包含一个或多个Series对象，每个数据系列包含一个或多个Points对象。例如，单张图表可能既包含折线图图表组（其中包含所有用折线图格式绘制的数据系列），也包含条形图图表组（其中包含所有用条形图格式绘制的数据系列）。ChartGroup对象是ChartGroups集合的成员。

使用ChartGroups(index)（其中index是图表组的索引号）可以返回单个ChartGroup对象。

因为当特定图表组所用的图表格式更改时，该图表组的索引号可能会更改，所以使用命名图表组快捷方法之一来返回特定的图表组会更加容易。PieGroups方法返回图表中饼图图表组的集合，LineGroups方法返回图表中折线图图表组的集合，依此类推。这些方法中的每一个都可以与索引号配合使用以返回单个ChartGroup对象，或不指定索引号而返回ChartGroups集合。

## [示例​](#示例)

javascript
```javascript
/*本示例开启图表工作表 Chart1 中第一个图表组的垂直线。*/
function test() {
    let chartgroup = Application.Charts.Item("Chart1").ChartObjects(1).Chart.ChartGroups(1)
    chartgroup.HasDropLines = true
}
```

如果图表已被激活，就可使用ActiveChart属性。

javascript
```javascript
/*本示例激活图表工作表 Chart1 ，并开启该图表中第一个图表组的垂直线。*/
function test() {
    Application.Charts.Item("Chart1").Activate()
    Application.ActiveSheet.ChartObjects(1).Chart.ChartGroups(1).HasDropLines = true
}
```


#### ChartGroups 对象

# [ChartGroups (对象)​](#chartgroups-对象)

代表图表中用同一格式绘制的一个或多个数据系列。

## [说明​](#说明)

ChartGroups集合是指定图表中的所有ChartGroup对象的集合。每张图表包含一个或多个图表组，每个图表组包含一个或多个数据系列，每个数据系列包含一个或多个数据点。例如，单张图表可能既包含折线图图表组（其中包含所有用折线图格式绘制的数据系列），也包含条形图图表组（其中包含所有用条形图格式绘制的数据系列）。

使用ChartGroups方法可返回ChartGroups集合。

## [示例​](#示例)

javascript
```javascript
/*本示例显示工作表 Sheet1 中第一个内嵌图表图表组的数量。*/
function test() {
    let  chartgroups = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.ChartGroups()
    console.log(chartgroups.Count)
}
```

使用ChartGroups(index)（其中index是图表组的索引号）可以返回单个ChartGroup对象。

javascript
```javascript
/*本示例为图表工作表 Chart1 中第一个图表组添加垂直线。*/
function test() {
    let chartgroup = Application.Charts.Item("Chart1").ChartObjects(1).Chart.ChartGroups(1)
    chartgroup.HasDropLines = true
}
```

如果图表已被激活，可使用ActiveChart：

javascript
```javascript
/*本示例激活图表工作表 Chart1 ，并为该表的第一个图表组添加垂直线。*/
function test() {
    Application.Charts.Item("Chart1").Activate()
    Application.ActiveSheet.ChartGroups(1).HasDropLines = true
}
```

因为当特定图表组所用的图表格式更改时，该图表组的索引号可能会更改，所以使用命名图表组快捷方法之一来返回特定的图表组会更加容易。PieGroups方法返回图表中饼图图表组的集合，LineGroups方法返回图表中折线图图表组的集合，依此类推。这些方法中的每一个都可以与索引号配合使用以返回单个ChartGroup对象，或不指定索引号而返回ChartGroups集合。


#### ChartObject 对象

# [ChartObject (对象)​](#chartobject-对象)

代表工作表上的嵌入图表。

## [说明​](#说明)

ChartObject对象充当Chart对象的容器。ChartObject对象的属性和方法控制工作表上嵌入图表的外观和大小。ChartObject对象是ChartObjects集合的成员。ChartObjects集合包含单一工作表上的所有嵌入图表。

使用ChartObjects(index)（其中index是嵌入图表的索引号或名称）可以返回单个ChartObject对象。

## [示例​](#示例)

javascript
```javascript
/*本示例将工作表 Sheet1 第一个图表的填充方案设置为浅色下对角线。*/
function test() {
    let chartobject = Application.Worksheets.Item("Sheet1").ChartObjects(1)
    chartobject.Chart.ChartArea.Format.Fill.Patterned(msoPatternLightDownwardDiagonal)
}
```

当选定嵌入图表时，其名称显示在“名称”框中。使用Name属性可设置或返回ChartObject对象的名称。

javascript
```javascript
/*此示例为工作表 Sheet1 的第二个图表添加圆角。*/
function test() {
    let chartobject = Application.Worksheets.Item("Sheet1").ChartObjects(2)
    chartobject.RoundedCorners = true
}
```


#### ChartObjects 对象

# [ChartObjects (对象)​](#chartobjects-对象)

由指定的图表工作表、对话框工作表或工作表上的所有ChartObject对象组成的集合。

## [说明​](#说明)

每个ChartObject对象都代表一个嵌入图表。ChartObject对象充当Chart对象的容器。ChartObject对象的属性和方法控制工作表上嵌入图表的外观和大小。

使用ChartObjects方法返回ChartObjects集合。

## [示例​](#示例)

javascript
```javascript
/*本示例删除工作表 Sheet1 中所有图表。*/
function test() {
    Application.Worksheets.Item("Sheet1").ChartObjects().Delete()
}
```

不能使用ChartObjects集合来调用以下属性和方法：

Locked
属性
Placement
属性
PrintObject
属性
与早期版本不同，ChartObjects集合现在可以读取表示高度、宽度、左对齐和顶对齐的属性。

使用Add方法可创建一个新的空嵌入图表并将它添加到集合中。使用ChartWizard方法可添加数据并设置新图表的格式。

javascript
```javascript
/*本示例在工作表 Sheet1 中新建一个内嵌图表，然后以折线图形式添加单元格 A1:A20 中的数据。*/
function test() {
    let chartobject = Application.Worksheets.Item("Sheet1").ChartObjects().Add(100, 30, 400, 250)
    chartobject.Chart.ChartWizard(Worksheets.Item("Sheet1").Range("A1:A20"), xlLine, null, null, null, null, null, "New Chart", null, null, null)
}
```

使用ChartObjects(index)（其中index是嵌入图表的索引号或名称）可以返回单个对象。

javascript
```javascript
/*本示例将工作表 Sheet1 上第一张图表填充方案设置为浅色下对角线。*/
function test() {
    let chart = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart
    chart.ChartArea.Format.Fill.Patterned(msoPatternLightDownwardDiagonal)
}
```


#### ChartTitle 对象

# [ChartTitle (对象)​](#charttitle-对象)

代表图表标题。

## [说明​](#说明)

使用ChartTitle属性可返回ChartTitle对象。

只有图表的HasTitle属性为True时，ChartTitle对象才存在，从而才能使用该对象。

## [示例​](#示例)

javascript
```javascript
/*此示例为工作表 Sheet1 上第一个图表添加标题。*/
function test() {
    let chart = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart
    chart.HasTitle = true
    chart.ChartTitle.Text = "February Sales"
}
```

javascript
```javascript
/*本示例将图表工作表 Chart1 的图表标题的字号设置为 20 磅。*/
function test() {
    let charttitle = Application.Charts.Item("Chart1").ChartObjects(1).Chart.ChartTitle
    charttitle.Font.Size = 20
}
```


#### Charts 对象

# [Charts (对象)​](#charts-对象)

指定的或活动工作簿中所有图表工作表的集合。

## [示例​](#示例)

javascript
```javascript
/*本示例设置图表工作表 Chart1 中图表的配色方案。*/
function test() {
    let chart = Application.Charts.Item("Chart1")
    chart.ChartColor = 7
}
```

javascript
```javascript
/*本示例在活动工作簿所有图表工作表上添加水平分页符。*/
function test() {
    Application.Charts.HPageBreaks.Add(Range("C3"))
}
```


#### ColorFormat 对象

# [ColorFormat (对象)​](#colorformat-对象)

代表单色对象的颜色、带有渐变或图案填充的对象的前景或背景色，或者指针的颜色。

## [说明​](#说明)

可以将颜色设为显式的红-绿-蓝值（使用RGB属性），或设为配色方案中的一种颜色（使用SchemeColor属性）。

使用下表中列出的属性之一可返回ColorFormat对象。

| 使用此属性 | 对象 | 返回一个 ColorFormat 对象，该对象代表 |
| --- | --- | --- |
| BackColor | FillFormat | 背景填充色（用于阴影或图案填充格式） |
| ForeColor | FillFormat | 前景填充色（对于纯色填充格式，即代表填充颜色） |
| BackColor | LineFormat | 线条背景色（用于图案线条） |
| ForeColor | LineFormat | 线条前景色（对于纯色线条，即代表线条颜色） |
| ForeColor | ShadowFormat | 阴影颜色 |
| ExtrusionColor | ThreeDFormat | 有延伸的对象的侧边颜色 |

使用RGB属性可将颜色设置为显示的红-绿-蓝值。

## [示例​](#示例)

javascript
```javascript
/*下例向第一张工作表中添加一个矩形，然后设置矩形填充的前景色、背景色和渐变。*/
function test() {
    let fillFormat = Worksheets.Item(1).Shapes.AddShape(msoShapeRectangle, 90, 90, 90, 50).Fill
    fillFormat.ForeColor.RGB = RGB(128, 0, 0)
    fillFormat.BackColor.RGB = RGB(170, 170, 170)
    fillFormat.TwoColorGradient(msoGradientHorizontal, 1)
}
```

javascript
```javascript
/*本示例显示活动工作表上第一个形状阴影的前景色的颜色类型是否为 msoColorTypeRGB。*/
function test() {
    console.log(ActiveSheet.Shapes(1).Shadow.ForeColor.Type == msoColorTypeRGB)
}
```


#### ColorScale 对象

# [ColorScale (对象)​](#colorscale-对象)

代表色阶条件格式规则。

## [说明​](#说明)

所有条件格式对象均包含在FormatConditions集合对象中，该集合对象是Range集合的子项。您可以使用FormatConditions集合的Add或AddColorScale方法创建色阶格式规则。

色阶是直观的参照，可以帮助您了解数据的分布和变化。您可以对数据范围、表中的数据或数据透视表中的数据应用双色或三色色阶。对于双色色阶条件格式，您可以将值、类型和颜色分配给范围的最小和最大阈值。三色色阶还具有中点阈值。

通过设置ColorScaleCriteria对象的属性，可以确定其中的每个阈值。ColorScaleCriteria对象是ColorScale对象的子项，也是色阶的所有ColorScaleCriterion对象的集合。

## [示例​](#示例)

javascript
```javascript
/*本示例创建了一个数字范围，然后将双色色阶条件格式规则应用于该范围。然后指定最小阈值的颜色为红色，最大阈值的颜色为蓝色。*/
function test() {
    // Fill cells with sample data from 1 to 10
    ActiveSheet.Range("C1").Value2 = 1
    ActiveSheet.Range("C2").Value2 = 2
    ActiveSheet.Range("C1:C2").AutoFill(Range("C1:C10"))
    Range("C1:C10").Select()

    // Create a two-color ColorScale object for the created sample data range
    let colorScale = Selection.FormatConditions.AddColorScale(2)

    // Set the minimum threshold to red and maximum threshold to blue
    colorScale.ColorScaleCriteria(1).FormatColor.Color = RGB(255, 0, 0)
    colorScale.ColorScaleCriteria(2).FormatColor.Color = RGB(0, 0, 255)
}
```

javascript
```javascript
/*本示例设置活动工作表上单元格区域 A1:A10 的第一个色阶条件格式所应用于的单元格区域，并设置该条件格式的第一个阈值条件的颜色和亮度。*/
function test() {
    let colorScale = ActiveSheet.Range("A1:A10").FormatConditions.Item(1)
    colorScale.ModifyAppliesToRange(Range("A7:A8"))
    colorScale.ColorScaleCriteria.Item(1).FormatColor.ColorIndex = 7
    colorScale.ColorScaleCriteria.Item(1).FormatColor.TintAndShade = 0.5
}
```


#### ColorScaleCriteria 对象

# [ColorScaleCriteria (对象)​](#colorscalecriteria-对象)

代表色阶条件格式的所有条件的ColorScaleCriterion对象的集合。每个条件指定了色阶的最小、中点或最大阈值。

## [说明​](#说明)

要返回ColorScaleCriteria集合，请使用ColorScale对象的ColorScaleCriteria属性。

## [示例​](#示例)

javascript
```javascript
/*本示例创建了一个数字范围，然后将双色色阶条件格式规则应用于该范围。然后通过在ColorScaleCriteria集合中编制索引来设置单独的条件，从而指定最小阈值的颜色为红色，最大阈值的颜色为蓝色。*/
function test() {
    // Fill cells with sample data from 1 to 10
    ActiveSheet.Range("C1").Value2 = 1
    ActiveSheet.Range("C2").Value2 = 2
    ActiveSheet.Range("C1:C2").AutoFill(Range("C1:C10"))
    Range("C1:C10").Select()

    // Create a two-color ColorScale object for the created sample data range
    let colorScale = Selection.FormatConditions.AddColorScale(2)

    // Set the minimum threshold to red and maximum threshold to blue
    colorScale.ColorScaleCriteria(1).FormatColor.Color = RGB(255, 0, 0)
    colorScale.ColorScaleCriteria(2).FormatColor.Color = RGB(0, 0, 255)
}
```

javascript
```javascript
/*本示例设置活动工作表上单元格区域 A1:A10 的第二个色阶条件格式的所有阈值条件的颜色。*/
function test() {
    let criteria = ActiveSheet.Range("A1:A10").FormatConditions.Item(2).ColorScaleCriteria
    for (let i = 1; i <= criteria.Count; i++) {
        criteria.Item(i).FormatColor.ColorIndex = 6 + i
    }
}
```


#### ColorStop 对象

# [ColorStop (对象)​](#colorstop-对象)

表示某一区域或选定内容中渐变填充的光圈点。

## [说明​](#说明)

ColorStop对象可用来设置单元格填充属性，包括Color、ThemeColor和TintAndShade属性。

## [示例​](#示例)

javascript
```javascript
/*以下示例演示如何对 ColorStop 对象应用属性。*/
function test() {
    let interior = Application.Selection.Interior
    interior.Pattern = xlPatternLinearGradient
    interior.Gradient.Degree = 135
    interior.Gradient.ColorStops.Clear()

    let colorStops1 = Selection.Interior.Gradient.ColorStops.Add(0)
    colorStops1.ThemeColor = xlThemeColorDark1
    colorStops1.TintAndShade = 0

    let colorStops2 = Application.Selection.Interior.Gradient.ColorStops.Add(0.5)
    colorStops2.ThemeColor = xlThemeColorAccent1
    colorStops2.TintAndShade = 0

    let colorStops3 = Application.Selection.Interior.Gradient.ColorStops.Add(1)
    colorStops3.ThemeColor = xlThemeColorDark1
    colorStops3.TintAndShade = 0
}
```

javascript
```javascript
/*本示例显示当前选定区域的第一个ColorStop对象的位置。*/
function test() {
    Range("A1:A10").Select()
    let colorStop = Selection.Interior.Gradient.ColorStops.Item(1)
    console.log(colorStop.Position)
}
```


#### ColorStops 对象

# [ColorStops (对象)​](#colorstops-对象)

指定的数据系列中所有 ColorStop 对象的集合。

## [说明​](#说明)

每个ColorStop对象代表一个区域或选定内容中渐变填充的一个颜色光圈。

## [示例​](#示例)

javascript
```javascript
/*以下示例显示了具有线性渐变的颜色停止。*/
function test() {
    let interior = Selection.Interior
    interior.Pattern = xlPatternLinearGradient
    interior.Gradient.Degree = 90
    interior.Gradient.ColorStops.Clear()

    //adds stops after any have been cleared
    let colorStops1 = Selection.Interior.Gradient.ColorStops.Add(0)
    colorStops1.ThemeColor = xlThemeColorDark1
    colorStops1.TintAndShade = 0

    let colorStops2 = Selection.Interior.Gradient.ColorStops.Add(1)
    colorStops2.ThemeColor = xlThemeColorAccent1
    colorStops2.TintAndShade = 0
}
```

javascript
```javascript
/*以下示例显示了具有矩形渐变的颜色停止点。*/
function test() {
    let interior = Selection.Interior
    interior.Pattern = xlPatternRectangularGradient
    interior.Gradient.RectangleLeft = 0
    interior.Gradient.RectangleRight = 0
    interior.Gradient.RectangleTop = 0
    interior.Gradient.RectangleBottom = 0
    interior.Gradient.ColorStops.Clear()

    let colorStops1 = Selection.Interior.Gradient.ColorStops.Add(0)
    colorStops1.Color = 192
    colorStops1.TintAndShade = 0

    let colorStops2 = Selection.Interior.Gradient.ColorStops.Add(1)
    colorStops2.ThemeColor = xlThemeColorLight1
    colorStops2.TintAndShade = 0
}
```


#### Comment 对象

# [Comment (对象)​](#comment-对象)

代表单元格批注。

## [说明​](#说明)

Comment对象是Comments集合的成员。

使用Comment属性可返回Comment对象。

## [示例​](#示例)

javascript
```javascript
/*本示例更改单元格 E5 中的批注文本。*/
function test() {
    Application.Worksheets.Item(1).Range("E5").Comment.Text("reviewed on " + Date())
}
```

使用Comments(index)（其中index为批注号）可返回Comments集合中的单条批注。

javascript
```javascript
/*本示例隐藏第一张工作表中的第二条批注。*/
function test() {
    Application.Worksheets.Item(1).Comments.Item(2).Visible = false
}
```

使用AddComment方法可在区域内添加批注。

javascript
```javascript
/*本示例在第一张工作表的单元格 E5 中添加批注。*/
function test() {
    let myComment = Application.Worksheets.Item(1).Range("E5").AddComment()
    myComment.Visible = false
    myComment.Text("reviewed on " + Date())
}
```


#### Comments 对象

# [Comments (对象)​](#comments-对象)

由单元格批注组成的集合。

## [说明​](#说明)

每个批注都由一个Comment对象代表。

使用Comments属性可返回Comments集合。

## [示例​](#示例)

javascript
```javascript
/*本示例隐藏第一张工作表上的所有批注。*/
function test() {
    let cmt = Application.Worksheets.Item(1).Comments
    for (let c = 1; c <= cmt.Count; c++) {
        cmt.Item(c).Visible = false
    }
}
```

使用AddComment方法可在区域内添加批注。

javascript
```javascript
/*本示例在第一张工作表的单元格 E5 上添加一个批注。*/
function test() {
    let myComment = Application.Worksheets.Item(1).Range("E5").AddComment()
    myComment.Visible = false
    myComment.Text("reviewed on " + Date())
}
```

使用Comments(index)（其中index为批注号）可返回Comments集合中的单条批注。

javascript
```javascript
/*本示例隐藏第一张工作表上的第二个批注。*/
function test() {
    Application.Worksheets.Item(1).Comments.Item(2).Visible = false
}
```


#### ConditionValue 对象

# [ConditionValue (对象)​](#conditionvalue-对象)

代表数据条条件格式规则计算最短数据条和最长数据条的方法。

## [说明​](#说明)

ConditionValue对象是使用Databar对象的MaxPoint或MinPoint属性返回的。

通过使用Modify方法，您可以从默认设置（最低值表示最短的数据条，最高值表示最长的数据条）中更改计算类型。

## [示例​](#示例)

javascript
```javascript
/*本示例将创建一个数据范围，然后对该范围应用数据条。并使用ConditionValue对象将阈值的计算方式更改为百分点。*/
function test() {
    //Create a range of data with a couple of extreme values
    Application.ActiveSheet.Range("D1").Value2 = 1
    Application.ActiveSheet.Range("D2").Value2 = 45
    Application.ActiveSheet.Range("D3").Value2 = 50
    Application.ActiveSheet.Range("D2:D3").AutoFill(Range("D2:D8"))
    Application.ActiveSheet.Range("D9").Value2 = 500
    Range("D1:D9").Select()

    //Create a data bar with default behavior
    let databar = Application.Selection.FormatConditions.AddDatabar()
    console.log("Because of the extreme values, middle data bars are very similar")

    //The MinPoint and MaxPoint properties return a ConditionValue object
    //which you can use to change threshold parameters
    databar.MinPoint.Modify(xlConditionValuePercentile, 5)
    databar.MaxPoint.Modify(xlConditionValuePercentile, 75)

}
```

javascript
```javascript
/*本示例判断如果第一张工作表上区域 A1:A10 中第一个（数据条）条件格式的最短的数据条的类型为xlConditionValueNumber，则修改该数据条的计算方法。*/
function test() {
    let conditionValue = Application.Worksheets.Item(1).Range("A1:A10").FormatConditions.Item(1).MinPoint
    if (conditionValue.Type == xlConditionValueNumber) {
        conditionValue.Modify(xlConditionValuePercent, 30)
    }
}
```


#### ConnectorFormat 对象

# [ConnectorFormat (对象)​](#connectorformat-对象)

包含应用于连接符的属性和方法。

## [说明​](#说明)

连接符是用于连接其他两个形状的线，所连接的位置叫做连接结点。如果重新排列已连接的形状，那么连接符的几何形状将自动调整，以使重新排列的形状仍保持连接。

连接结点通常按下表所示的规则进行编号。

| 形状类型 | 连接结点标号方案 |
| --- | --- |
| 自选形状、艺术字、图片和 OLE 对象 | 连接结点从顶部开始按逆时针进行编号。 |
| 任意多边形 | 连接结点为顶点，与顶点编号相对应。 |

使用ConnectorFormat属性可返回ConnectorFormat对象。使用BeginConnect和EndConnect方法可将连接符的两端连到文档中的其他形状。使用RerouteConnections方法可自动查找通过连接符连接的两个形状间的最短路径。使用Connector属性可查看形状是否为连接符。

请注意，虽然向Shapes集合添加连接符时，对其设置了大小和位置，但将连接符的起点和终点连接到其他形状时，连接符的大小和位置将会自动调整。因而，如果打算用连接符连接其他形状，那么对其设置的初始大小和位置就没有什么实际意义。同样，用连接符连接其他形状时，将指定要连接的形状上的连接结点，但将连接符连接好之后，用RerouteConnections方法可能会改变连接符所连接的连接结点，使开始时选定的连接结点变得没有意义。

要算出一个复杂形状上各连接结点的编号，可以打开宏录制器并对形状进行试验操作，然后查看录下的代码；也可以创建一个形状并选中它，然后运行以下示例。这段代码将对每个连接结点进行编号并连接一个连接符。

## [示例​](#示例)

javascript
```javascript
/*本示例将对活动工作表所选形状的每个连接结点进行编号并连接一个连接符。*/
function test() {
    ActiveSheet.Shapes.SelectAll()
    let shape = Selection.ShapeRange.Item(1)
    let bx = shape.Left + shape.Width + 50
    let by = shape.Top + shape.Height + 50
    let count = shape.ConnectionSiteCount
    for (let j = 1; j <= count; j++) {
        let connector = ActiveSheet.Shapes.AddConnector(msoConnectorStraight, bx, by, bx + 50, by + 50)
        connector.ConnectorFormat.EndConnect(shape, j)
        connector.ConnectorFormat.Type = msoConnectorElbow
        connector.Line.ForeColor.RGB = RGB(255, 0, 0)
        let l = connector.Left
        let t = connector.Top
        let textbox = ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, l, t, 36, 14)
        textbox.Fill.Visible = false
        textbox.Line.Visible = false
        textbox.TextFrame.Characters().Text = j
    }
}
```

javascript
```javascript
/*下例向第一张工作表中添加两个矩形并且用曲线连接符连接矩形。*/
function test() {
    let worksheet = Application.Worksheets.Item(1)
    let s = worksheet.Shapes
    let firstRect = s.AddShape(msoShapeRectangle, 100, 50, 200, 100)
    let secondRect = s.AddShape(msoShapeRectangle, 300, 300, 200, 100)
    let c = s.AddConnector(msoConnectorCurve, 0, 0, 0, 0)
    c.ConnectorFormat.BeginConnect(firstRect, 1)
    c.ConnectorFormat.EndConnect(secondRect, 1)
    c.RerouteConnections()
}
```


#### ControlFormat 对象

# [ControlFormat (对象)​](#controlformat-对象)

包含 ET 控件 （ET 控件：ET 本身具有的控件，而不是 ActiveX 控件。）属性。

## [示例​](#示例)

javascript
```javascript
/*下例为第一张工作表上的列表框控件设置填充区域。*/
function test() {
    Worksheets.Item(1).Shapes.Item(1).ControlFormat.ListFillRange = "A1:A10"
}
```

javascript
```javascript
/*本示例删除活动工作表上第二个形状（组合框）的所有数据项。*/
function test() {
    let shape = ActiveSheet.Shapes.Item(2)
    shape.ControlFormat.RemoveAllItems()
}
```


#### CustomProperties 对象

# [CustomProperties (对象)​](#customproperties-对象)

由代表附加信息的CustomProperty对象组成的集合，这些信息可用作 XML 的元数据。

## [说明​](#说明)

使用Worksheet对象的CustomProperties属性返回CustomProperties集合。

返回CustomProperties集合后，可根据选择向工作表和智能标记中添加元数据。

若要向工作表添加元数据，请在Add方法中使用CustomProperties属性。

下例演示了该功能。在此示例中，ET 向活动工作表添加标识符信息，并向用户返回名称和值。

javascript
```javascript
function test() {
    let wksSheet1 = Application.ActiveSheet

    // Add metadata to worksheet.
    wksSheet1.CustomProperties.Add("Market", "Nasdaq")

    // Display metadata.
    let cusProperties = wksSheet1.CustomProperties.Item(1)
    console.log(cusProperties.Name + "\t" + cusProperties.Value)
}
```


#### CustomProperty 对象

# [CustomProperty (对象)​](#customproperty-对象)

代表标识符信息。标识符信息可用于 XML 的元数据。

## [说明​](#说明)

使用Add方法或CustomProperties集合的Item属性可返回CustomProperty对象。

返回CustomProperty对象后，可在Add方法中使用CustomProperties属性向工作表中添加元数据。

在本示例中，ET 向活动工作表添加标识符信息，并向用户返回名称和值。

javascript
```javascript
function test() {
    let wksSheet1 = Application.ActiveSheet

    // Add metadata to worksheet.
    wksSheet1.CustomProperties.Add("Market", "Nasdaq")

    // Display metadata.
    let cusProperties = wksSheet1.CustomProperties.Item(1)
    console.log(cusProperties.Name + "\t" + cusProperties.Value)
}
```


#### DataBarBorder 对象

# [DataBarBorder (对象)​](#databarborder-对象)

表示由条件格式规则指定的数据条的边框。

## [说明​](#说明)

使用DataBarBorder对象可获取或设置数据条的颜色和边框类型。若要访问与数据条条件格式规则关联的DataBarBorder对象，请使用BarBorder属性。在检索DataBarBorder对象后，使用其Color属性返回可用来设置数据条颜色的FormatColor对象。

## [示例​](#示例)

javascript
```javascript
/*本示例选择一个单元格区域，将数据条条件格式规则添加到该区域，使用BarBorder属性检索与该规则关联的DataBarBorder对象，然后设置数据条的颜色、淡色和类型。*/
function test() {
    Range("B1:B10").Select()
    Range("B1:B10").Activate()
    let dataBar = Selection.FormatConditions.AddDatabar()
    let dataBarBorder = dataBar.BarBorder
    dataBarBorder.Type = xlDataBarBorderSolid
    dataBarBorder.Color.ThemeColor = xlThemeColorAccent2
    dataBarBorder.Color.TintAndShade = 0
}
```

javascript
```javascript
/*本示例判断如果活动工作表上区域 A1:A10 中第一个条件格式（数据条）的边框为实心边框，则设置该边框的颜色。*/
function test() {
    let dataBarBorder = ActiveSheet.Range("A1:A10").FormatConditions.Item(1).BarBorder
    if (dataBarBorder.Type == xlDataBarBorderSolid) {
        dataBarBorder.Color.ColorIndex = 4
    }
}
```


#### DataLabel 对象

# [DataLabel (对象)​](#datalabel-对象)

代表图表数据点或趋势线上的数据标签。

## [说明​](#说明)

在数据系列上，DataLabel对象是DataLabels集合的成员。DataLabels集合包含每个数据点的DataLabel对象。对于没有可定义数据点的数据系列（如面积图数据系列），DataLabels集合包含单个DataLabel对象。

使用DataLabels(index)（其中index为数据标签的索引号）可返回单个DataLabel对象。

## [示例​](#示例)

javascript
```javascript
/*本示例在工作表 Sheet1 上第一个图表上，设置第一个数据系列中的第五个数据标签的数字格式。*/
function test() {
    let datalabel = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.SeriesCollection(1).DataLabels(5)
    datalabel.NumberFormat = "0.000"
}
```

使用DataLabel属性可返回单个数据点的DataLabel对象。

javascript
```javascript
/*此示例打开名为“Chart1”的图表工作表上第一个数据系列中第二个数据点的数据标签，并将数据标签文本设置为“Saturday”。*/
function test() {
    let point = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(1).Points(2)
    point.HasDataLabel = true
    point.DataLabel.Text = "Saturday"
}
```

在趋势线上，DataLabel返回与趋势线一起显示的文本。这些文本可能是公式、R 平方值或两者均有（如果两者都出现）。

javascript
```javascript
/*本示例将趋势线文本设置为仅显示公式，并将数据标签的名称置于工作表“Sheet1”上的单元格 A1 中。*/
function test() {
    let trendline = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(1).Trendlines(1)
    trendline.DisplayRSquared = false
    trendline.DisplayEquation = true
    Worksheets.Item("Sheet1").Range("A1").Value2 = trendline.DataLabel.Name
}
```


#### DataLabels 对象

# [DataLabels (对象)​](#datalabels-对象)

由数据系列中所有DataLabel对象组成的集合。

## [说明​](#说明)

每个DataLabel对象代表一个数据点或趋势线的数据标签。对于没有可定义数据点的数据系列（例如面积图数据系列），DataLabels集合包含单个数据标签。

使用DataLabels方法可返回单个DataLabels集合。

## [示例​](#示例)

javascript
```javascript
/*此示例设置图表工作表 Chart1 中图表的第一个数据系列中数据标签的数字格式。*/
function test() {
    let series = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(1)
    series.HasDataLabels = true
    series.DataLabels().NumberFormat = "#.##0"
}
```

使用DataLabels(index)（其中 index 为数据标签的索引号）可返回单个DataLabel对象。

javascript
```javascript
/*此示例在第一张工作表上嵌入的第一个图表上，设置第一个数据系列中的第五个数据标签的数字格式。*/
function test() {
    let datalabel = Application.Worksheets.Item(1).ChartObjects(1).Chart.SeriesCollection(1).DataLabels(5)
    datalabel.NumberFormat = "0.000"
}
```


#### DataTable 对象

# [DataTable (对象)​](#datatable-对象)

代表一张图表模拟运算表。

## [说明​](#说明)

使用DataTable属性可返回DataTable对象。

## [示例​](#示例)

javascript
```javascript
/*此示例向嵌入式图表中添加带有外边框的模拟运算表。*/
function test() {
    let chart = Application.Worksheets.Item(1).ChartObjects(1).Chart
    chart.HasDataTable = true
    chart.DataTable.HasBorderOutline = true
}
```

javascript
```javascript
/*本示例将图表工作表 Chart1 中图表的模拟运算表中的边框和字体设置为蓝色。*/
function test() {
    let datatable = Application.Charts.Item("Chart1").ChartObjects(1).Chart.DataTable
    datatable.Border.ColorIndex = 5
    datatable.Font.ColorIndex = 5
}
```


#### Databar 对象

# [Databar (对象)​](#databar-对象)

代表数据条条件格式规则。通过对范围应用数据条，有助于查看相对于其他单元格的单元格的值。

## [说明​](#说明)

所有条件格式对象都包含在FormatConditions集合对象中，该集合对象是Range集合的子项。可以使用FormatConditions集合的Add或AddDatabar方法创建数据条格式设置规则。

可以使用Databar对象的MinPoint和MaxPoint属性设置数据范围的最短和最长数据条的值。这些属性会返回ConditionValue对象，使用该对象可指定如何计算阈值。

Databar对象还提供了另外一些属性，使用这些属性可指定在存在负值时显示的轴线，以及指定数据条的颜色和格式设置。

## [示例​](#示例)

javascript
```javascript
/*本示例将在活动工作表创建一个数据范围，然后对该范围应用数据条。为了明确显示中间值，使用ConditionValue对象将阈值的计算方式更改为百分点。*/
function test() {
    // Create a range of data with a couple of extreme values
    ActiveSheet.Range("D1").Value2 = 1
    ActiveSheet.Range("D2").Value2 = 45
    ActiveSheet.Range("D3").Value2 = 50
    ActiveSheet.Range("D2:D3").AutoFill(Range("D2:D8"))
    ActiveSheet.Range("D9").Value2 = 500
    Range("D1:D9").Select()

    // Create a data bar with default behavior
    let dataBar = Selection.FormatConditions.AddDatabar()
    console.log("Because of the extreme values, middle data bars are very similar")

    // The MinPoint and MaxPoint properties return a ConditionValue object
    // which you can use to change threshold parameters
    dataBar.MinPoint.Modify(xlConditionValuePercentile, 5)
    dataBar.MaxPoint.Modify(xlConditionValuePercentile, 75)
}
```

javascript
```javascript
/*本示例将活动工作表上区域 A1:A10 中第一个条件格式（数据条）的填充色设置为渐变色，并设置该数据条的展示方向。*/
function test() {
    let databar = ActiveSheet.Range("A1:A10").FormatConditions.Item(1)
    databar.BarFillType = xlDataBarFillGradient
    databar.Direction = xlRTL
}
```


#### DisplayFormat 对象

# [DisplayFormat (对象)​](#displayformat-对象)

代表关联的Range对象的显示设置。只读。

## [说明​](#说明)

更改区域的条件格式或表格样式等操作可导致当前用户界面中的显示内容与Range对象相应属性中的值不一致。使用DisplayFormat对象的属性可返回当前用户界面中显示的值。

## [示例​](#示例)

javascript
```javascript
/*本示例设置在当前用户界面中 G2 单元格边框的粗细为 xlThick ，并显示是否设置成功。*/
function test() {
    Range("G2").Style.Borders.Weight = xlThick
    console.log(Range("G2").DisplayFormat.Style.Borders.Weight == xlThick)
}
```

javascript
```javascript
/*本示例显示在当前用户界面中 D3:D5 单元格文本的字体颜色是否为红色。*/
function test() {
    console.log(Range("D3:D5").DisplayFormat.Font.Color == RGB(255, 0, 0))
}
```


#### DisplayUnitLabel 对象

# [DisplayUnitLabel (对象)​](#displayunitlabel-对象)

代表指定图表中坐标轴上的单位标志。

## [说明​](#说明)

在绘制大数据的图表时使用单位标志会很有用，例如：上百万或几十亿的大数据。在刻度线上使用单位标志而不直接使用大数据可以使图表更易读、易理解。

使用DisplayUnitLabel属性可返回DisplayUnitLabel对象。

## [示例​](#示例)

javascript
```javascript
/*以下示例将 Chart1 上数值轴上的显示标签标题设置为“Millions”，然后关闭自动字体缩放。*/
function test() {
    let Axis = Application.Charts.Item("Chart1").ChartObjects(1).Chart.Axes(xlValue)
    Axis.DisplayUnit = xlMillions
    Axis.HasDisplayUnitLabel = true
    let displayUnitLabel = Axis.DisplayUnitLabel
    displayUnitLabel.Caption = "Millions"
    displayUnitLabel.AutoScaleFont = false
}
```

javascript
```javascript
/*下例删除 Sheet1 上的第一个图表数值轴上的单位标签。*/
function test() {
    let displayUnitLabel = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.Axes(xlValue).DisplayUnitLabel
    displayUnitLabel.Delete()
}
```


#### DownBars 对象

# [DownBars (对象)​](#downbars-对象)

代表图表组中的跌柱线。

## [说明​](#说明)

跌柱线将图表组中第一个系列的数据点与最后一个系列中相应的有较小值的数据点连接起来（从第一个系列向下生长）。只有至少包含两个系列的二维折线图才能有跌柱线。此对象不是集合。没有代表单个跌柱线的对象；或者打开图表组中所有数据点的涨跌柱线，或者将其全部关闭。

如果HasUpDownBars属性为False，DownBars对象的绝大部分属性将被禁用。

使用DownBars属性可返回DownBars对象。

## [示例​](#示例)

javascript
```javascript
/*本示例打开工作表“Sheet5”上嵌入的第一个图表中第一个图表组的涨跌柱线，然后将涨柱线的颜色设置为蓝色，而将跌柱线设置为红色。*/
function test() {
    let chartgroup = Application.Worksheets.Item("Sheet5").ChartObjects(1).Chart.ChartGroups(1)
    chartgroup.HasUpDownBars = true
    chartgroup.UpBars.Interior.Color = RGB(0, 0, 255)
    chartgroup.DownBars.Interior.Color = RGB(255, 0, 0)
}
```

javascript
```javascript
/*本示例将图表工作表 Chart1 中图表的第一个图表组的跌柱线的线条颜色和前景色设置为红色。*/
function test() {
    let downbars = Application.Charts.Item("Chart1").ChartObjects(1).Chart.ChartGroups(1).DownBars
    downbars.Format.Fill.ForeColor.RGB = RGB(255, 0, 0)
    downbars.Format.Line.ForeColor.RGB = RGB(255, 0, 0)
}
```


#### DropLines 对象

# [DropLines (对象)​](#droplines-对象)

代表图表组中的垂直线。

## [说明​](#说明)

垂直线将图表中的数据点与 x 轴连接起来。只有折线图和面积图组可以有垂直线。此对象不是集合。没有代表单个垂直线的对象；或者打开图表组中所有数据点的垂直线，或者将其全部关闭。

如果HasDropLines属性为False，DropLines对象的绝大部分属性将被禁用。

使用DropLines属性可返回DropLines对象。

## [示例​](#示例)

javascript
```javascript
/*本示例打开嵌入的第一个图表的第一个图表组的垂直线，并将垂直线的颜色设置为红色。*/
function test() {
    Application.Worksheets.Item("Sheet1").ChartObjects(1).Activate()
    ActiveChart.ChartGroups(1).HasDropLines = true
    ActiveChart.ChartGroups(1).DropLines.Border.ColorIndex = 3
}
```

javascript
```javascript
/*本示例选中图表工作表 Chart1 中图表的第一个图表组的垂直线，并将该垂直线的阴影设置为不可见。*/
function test() {
    let droplines = Application.Charts.Item("Chart1").ChartObjects(1).Chart.ChartGroups(1).DropLines
    droplines.Select()
    Selection.Format.Shadow.Visible = false
}
```


#### Error 对象

# [Error (对象)​](#error-对象)

代表区域的电子表格错误。

## [说明​](#说明)

使用Errors对象的Item属性可返回Error对象。

返回Error对象后，可结合使用Value和Errors属性检查某个特定错误检查选项是否已启用。

| 注释 |
| --- |
| 不要将Error对象与 Visual Basic 的错误处理功能混淆。 |

javascript
```javascript
/*下例中，在引用空单元格的单元格 A1 中创建一个公式，然后使用 Item(index)（其中 index 用于标识错误类型）显示描述错误情况的消息。*/
function test(){
 let rngFormula = Application.Range("A1")

    //Place a formula referencing empty cells.
    Range("A1").Formula = "=A2+A3"
    Application.ErrorCheckingOptions.EmptyCellReferences = true

    //Perform check to see if EmptyCellReferences check is on.
    if(rngFormula.Errors.Item(xlEmptyCellReferences).Value == true) {
        console.log("The empty cell references error checking feature is enabled.")
    }
    else {
        console.log("The empty cell references error checking feature is not on.")
    }
}
```


#### ErrorBars 对象

# [ErrorBars (对象)​](#errorbars-对象)

代表图表数据系列上的误差线。

## [说明​](#说明)

误差线表明图表数据的不确定程度。只有二维图表上的面积、条形、柱形、折线和散点图组中的数据系列可以有误差线。只有散点图组中的数据系列可以有 x 误差线和 y 误差线。此对象不是集合。没有代表单个误差线的对象；或者打开系列中所有数据点的 x 误差线或 y 误差线，或者将其全部关闭。

ErrorBar方法更改误差线的格式和类型。

使用ErrorBars属性可返回ErrorBars对象。

## [示例​](#示例)

javascript
```javascript
/*本示例打开工作表 Sheet1 嵌入的第一个图表的第一个数据系列的误差线，并设置误差线的尾部样式。*/
function test() {
    Application.Worksheets.Item("Sheet1").ChartObjects(1).Activate()
    ActiveChart.SeriesCollection(1).HasErrorBars = true
    ActiveChart.SeriesCollection(1).ErrorBars.EndStyle = xlNoCap
}
```

javascript
```javascript
/*此示例将图表工作表 Chart1 中图表的第二个系列的误差线设置为黄色虚线。*/
function test() {
    let errorbars = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(2).ErrorBars
    errorbars.Border.ColorIndex = 6
    errorbars.Border.LineStyle = xlDash
}
```


#### Errors 对象

# [Errors (对象)​](#errors-对象)

## [说明​](#说明)

使用Range对象的Errors属性可返回Errors对象。

返回Errors对象后，可使用Error对象的Value属性检查特定的错误检查条件。

javascript
```javascript
/*下例将一个数字作为文本放在单元格 A1 中，然后当单元格 A1 的值包含文本格式的数字时通知用户。*/
function test(){
 //Place a number written as text in cell A1.
    Range("A1").Formula = "'1"

    if(Range("A1").Errors.Item(xlNumberAsText).Value == true) {
        console.log("Cell A1 has a number as text.")
    }
    else {
        console.log("Cell A1 is a number.")
    }

}
```


#### FillFormat 对象

# [FillFormat (对象)​](#fillformat-对象)

代表形状的填充格式。

## [说明​](#说明)

形状可以有纯色、渐变、纹理、图案、图片或半透明填充。

FillFormat对象的很多属性是只读的。要设置这些属性中每一个，必须使用相应的方法。

使用Fill属性可返回FillFormat对象。

## [示例​](#示例)

javascript
```javascript
/*下例向第一张工作表中添加矩形并且设置矩形填充的渐变和颜色。*/
function test() {
    let sheet = Application.Worksheets.Item(1)
    let shapes = sheet.Shapes.AddShape(msoShapeRectangle, 90, 90, 90, 80).Fill
    shapes.ForeColor.RGB = (0, 128, 128)
    shapes.OneColorGradient(msoGradientHorizontal, 1, 1)
}
```

javascript
```javascript
/*本示例在活动工作表中添加椭圆，并将前景色设置为红色，填充图案设置为深色竖线。*/
function test() {
    let fillFormat = ActiveSheet.Shapes.AddShape(msoShapeOval, 0, 0, 40, 80).Fill
    fillFormat.ForeColor.RGB = RGB(255, 0, 0)
    fillFormat.Patterned(msoPatternDarkVertical)
}
```


#### Filter 对象

# [Filter (对象)​](#filter-对象)

代表单个列的筛选。

## [说明​](#说明)

Filter对象是Filters集合的成员。Filters集合包含自动筛选区域中的所有筛选。

使用Filters(index)（其中index是筛选的标题或索引号）可返回单个Filter对象。下例将一个变量设置为工作表 Crew 上的筛选区域中第一列的筛选的On属性的值。

javascript
```javascript
function test()
{
    let w = Application.Worksheets.Item("Crew")
    if(w.AutoFilterMode) {
        filterIsOn = w.AutoFilter.Filters.Item(1).On
    }
}
```

注意：Filter对象的所有属性都是只读的。要设置这些属性，请手动应用自动筛选，或使用Range对象的AutoFilter方法，如下例中所示。

javascript
```javascript
function test() {
    let w = Application.Worksheets.Item("Crew")     
    w.Cells.AutoFilter(2, "Crucial", xlOr, "Important")
}
```


#### Filters 对象

# [Filters (对象)​](#filters-对象)

由多个Filter对象组成的集合，这些对象代表自动筛选区域内的所有筛选。

## [说明​](#说明)

使用Filters属性可返回Filters集合。下例创建一个列表，其中包含工作表 Crew 中已自动筛选的区域内所有筛选的条件和运算符。

javascript
```javascript
function test()
{
    let f
    let op
    let c1
    let c2
    let ns = "Not set"

    let w = Application.Worksheets.Item("Sheet1")
    let w2 = Application.Worksheets.Item("Sheet2")
    let rw = 1
    for(let i = 1 ; i <= w.AutoFilter.Filters.Count; i++) {
        f = w.AutoFilter.Filters.Item(i)
        if(f.On) {
            let c1 = f.Criteria1.substr(f.Criteria1.length - 1,1)
            if(f.Operator) {
                op = f.Operator
                c2 = f.Criteria2.substr(f.Criteria2.length - 1,1)
            } else {
                op = ns
                c2 = ns
            }
        } else {
            c1 = ns
            op = ns
            c2 = ns
        }
        w2.Cells.Item(rw, 1).Value2 = c1
        w2.Cells.Item(rw, 2).Value2 = op
        w2.Cells.Item(rw, 3).Value2 = c2
        rw = rw + 1
    }
}
```

使用Filters(index)（其中index是筛选的标题或索引号）可返回单个Filter对象。下例将一个变量设置为工作表 Crew 上的筛选区域中第一列的筛选的On属性的值。

javascript
```javascript
function test()
{
    let w = Worksheets.Item("Crew")
    if(w.AutoFilterMode) {
        let filterIsOn = w.AutoFilter.Filters.Item(1).On
    }
}
```


#### Font 对象

# [Font (对象)​](#font-对象)

包含对象的字体属性（字体名称、字号、颜色等等）。

## [说明​](#说明)

如果不想将单元格中的文本或图形设为相同的格式，则使用Characters属性返回文本的子集。

使用Font属性可返回Font对象。

## [示例​](#示例)

javascript
```javascript
/*本示例将工作表 Sheet1中单元格 A1:C5 区域中的字体设为加粗*/
function test() {
    Application.Worksheets.Item("Sheet1").Range("A1:C5").Font.Bold = true
}
```

javascript
```javascript
/*此示例显示第一个工作表中D1单元格的字体样式*/
function test() {
    let font = Application.Worksheets.Item(1).Range("D1").Font
    console.log(`字体样式为 ${font.FontStyle}`)
}
```

``


#### FormatColor 对象

# [FormatColor (对象)​](#formatcolor-对象)

代表为色阶条件格式阈值指定的填充色或数据条条件格式的条形颜色。

## [说明​](#说明)

您可以通过传递Color属性中的 RGB 值来选择颜色，或者通过使用ThemeColor属性在主题调色板中编制索引来指定颜色。

以下代码示例创建了一个数字范围，然后将双色色阶条件格式规则应用于该范围。然后通过在ColorScaleCriteria集合中编制索引来设置单独的条件，从而指定最小阈值的颜色为红色，最大阈值的颜色为蓝色。

## [示例​](#示例)

javascript
```javascript
/*以下代码示例创建了一个数字范围，然后将双色色阶条件格式规则应用于该范围。然后，通过索引到 ColorScaleCriteria 集合以设置单个条件，将最小阈值的颜色分配给红色，并将最大阈值分配给蓝色。*/
function test() {
    //Fill cells with sample data from 1 to 10
    ActiveSheet.Range("C1").Value2 = 1
    ActiveSheet.Range("C2").Value2 = 2
    ActiveSheet.Range("C1:C2").AutoFill(Range("C1:C10"))

    Range("C1:C10").Select()

    //Create a two-color ColorScale object for the created sample data range
    let colorScale = Selection.FormatConditions.AddColorScale(2)

    //Set the minimum threshold to red and maximum threshold to blue
    colorScale.ColorScaleCriteria.Item(1).FormatColor.Color = RGB(255, 0, 0)
    colorScale.ColorScaleCriteria.Item(2).FormatColor.Color = RGB(0, 0, 255)
}
```

javascript
```javascript
/*本示例显示指定单元格区域的第一个色阶条件格式中第二个条件的最大阈值的主题颜色是否为xlThemeColorAccent1。*/
function test() {
    let formatCondition = ActiveSheet.Range("A1:A10").FormatConditions.Item(1)
    console.log(formatCondition.ColorScaleCriteria.Item(2).FormatColor.ThemeColor == xlThemeColorAccent1)
}
```


#### FormatCondition 对象

# [FormatCondition (对象)​](#formatcondition-对象)

代表条件格式。

## [说明​](#说明)

FormatCondition对象是FormatConditions集合的成员。对于给定区域，FormatConditions集合中包含的条件格式不能超过三个。

使用Add方法可新建条件格式。如果区域内存在多种格式，则可使用Modify方法更改其中一种格式，或使用Delete方法删除一种格式，然后使用Add方法创建一种新格式。

使用FormatCondition对象的Font、Borders和Interior属性可控制已设置格式的单元格的外观。条件格式对象模型不支持这些对象的某些属性。下表列出所有可与条件格式一起使用的属性。

| 对象 | 属性 |
| --- | --- |
| Font | BoldColorColorIndexFontStyleItalicStrikethroughUnderline无法使用会计用下划线样式。 |
| Border | BottomColorLeftRightStyle可使用下列边框样式（其他均不可用）：xlNone、xlSolid、xlDash、xlDot、xlDashDot、xlDashDotDot、xlGray50、xlGray75和xlGray25。TopWeight可使用下列边框粗细（其他均不可用）：xlWeightHairline和xlWeightThin。 |
| Interior | ColorColorIndexPatternPatternColorIndex |

使用FormatConditions(index)（其中index为条件格式的索引号）可返回FormatCondition对象。

## [示例​](#示例)

javascript
```javascript
/*本示例设置第一张工作表上 E1:E10 单元格的现有条件格式的格式属性。*/
function test() {
    let formatCondition = Application.Worksheets.Item(1).Range("E1:E10").FormatConditions.Item(1)
    let boders = formatCondition.Borders
    boders.LineStyle = xlContinuous
    boders.Weight = xlThin
    boders.ColorIndex = 6
    let font = formatCondition.Font
    font.Bold = true
    font.ColorIndex = 3
}
```

javascript
```javascript
/*本示例设置活动工作表上单元格区域 A1:A10 的第一个条件格式所应用于的单元格区域，并修改现有条件格式。*/
function test() {
    let formatCondition = ActiveSheet.Range("A1:A10").FormatConditions.Item(1)
    formatCondition.ModifyAppliesToRange(Range("A2:A9"))
    formatCondition.Modify(xlCellValue, xlBetween, "=4", "=7")
}
```


#### FormatConditions 对象

# [FormatConditions (对象)​](#formatconditions-对象)

代表一个区域内所有条件格式的集合。

## [说明​](#说明)

FormatConditions集合可以包含多个条件格式。每个格式由一个FormatCondition对象代表。

有关条件格式的详细信息，请参阅FormatCondition对象。

使用FormatConditions属性可返回FormatConditions对象。使用Add方法可新建条件格式，使用Modify方法可更改现有的条件格式。

## [示例​](#示例)

javascript
```javascript
/*本示例向第一张工作表上单元格区域 E1:E10 中添加条件格式*/
function test() {
    let formatCondition = Application.Worksheets.Item(1).Range("E1:E10").FormatConditions.Add(xlCellValue, xlGreater, "=$A$1")
    let boders = formatCondition.Borders
    boders.LineStyle = xlContinuous
    boders.Weight = xlThin
    boders.ColorIndex = 6
    let font = formatCondition.Font
    font.Bold = true
    font.ColorIndex = 3
}
```

javascript
```javascript
/*本示例在活动工作表上单元格区域 F1:F10 新增Top10条件格式。*/
function test() {
    let top = ActiveSheet.Range("F1:F10").FormatConditions.AddTop10()
    top.TopBottom = xlTop10Top
    top.Rank = 5
    top.Percent = false
    top.Font.Bold = true
    top.Interior.ColorIndex = 5
}
```


#### FreeformBuilder 对象

# [FreeformBuilder (对象)​](#freeformbuilder-对象)

代表任意多边形创建时的几何属性。

## [说明​](#说明)

使用BuildFreeform方法可返回FreeformBuilder对象。使用AddNodes方法可在任意多边形中添加节点。使用ConvertToShape方法可创建FreeformBuilder对象中定义的形状，并将它添加到Shapes集合中。

## [示例​](#示例)

javascript
```javascript
/*本示例向第一张工作表中添加带有四条线段的任意多边形。*/
function test() {
    let builder = Worksheets.Item(1).Shapes.BuildFreeform(msoEditingCorner, 360, 200)
    builder.AddNodes(msoSegmentCurve, msoEditingCorner, 380, 230, 400, 250, 450, 300)
    builder.AddNodes(msoSegmentCurve, msoEditingAuto, 480, 200)
    builder.AddNodes(msoSegmentLine, msoEditingAuto, 480, 400)
    builder.AddNodes(msoSegmentLine, msoEditingAuto, 360, 200)
    let s = builder.ConvertToShape()
    s.Fill.ForeColor.RGB = RGB(0, 0, 0)
}
```

javascript
```javascript
/*本示例在活动工作表中创建一个具有三个顶点的红色任意多边形。*/
function test() {
    let builder = ActiveSheet.Shapes.BuildFreeform(msoEditingAuto, 470, 330)
    builder.AddNodes(msoSegmentCurve, msoEditingSmooth, 570, 360)
    builder.AddNodes(msoSegmentCurve, msoEditingSmooth, 690, 430)
    builder.AddNodes(msoSegmentCurve, msoEditingSmooth, 470, 330)
    let s = builder.ConvertToShape()
    s.Fill.ForeColor.RGB = RGB(255, 0, 0)
}
```


#### Gridlines 对象

# [Gridlines (对象)​](#gridlines-对象)

代表图表坐标轴的主要和次要网格线。

## [说明​](#说明)

网格线延伸图表坐标轴上的刻度线，以便更容易地分辨与数据标志相关联的数值。此对象不是集合。同时也没有代表单个网格线的对象；要么打开坐标轴上所有的网格线，要么将其全部关闭。

使用MajorGridlines属性可返回GridLines对象，该对象代表坐标轴的主要网格线。使用MinorGridlines属性可返回GridLines对象，该对象代表次要网格线。可以同时返回主要网格线和次要网格线。

## [示例​](#示例)

javascript
```javascript
/*下例打开图表工作表 Chart1 上分类轴的主要网格线，并将网格线设置为蓝色虚线。*/
function test() {
    let axis = Application.Charts.Item("Chart1").ChartObjects(1).Chart.Axes(xlCategory)
    axis.HasMajorGridlines = true
    axis.MajorGridlines.Border.Color = RGB(0, 0, 255)
    axis.MajorGridlines.Border.LineStyle = xlDash
}
```

javascript
```javascript
/*本示例将工作表 Sheet1 中第一个图表数值轴的主要网格线名称赋值给B6单元格。*/
function test() {
    let axis = Worksheets.Item("Sheet1").ChartObjects(1).Chart.Axes(xlValue)
    Range("B6").Value2 = axis.MajorGridlines.Name
}
```


#### GroupShapes 对象

# [GroupShapes (对象)​](#groupshapes-对象)

代表一组形状中的单个形状。

## [说明​](#说明)

每个形状都由一个Shape对象代表。将Item方法与此对象一起使用，您可以在不取消分组的情况下处理组合的各个形状。

使用GroupItems属性可返回GroupShapes集合。使用GroupItems(index)（其中index是分组的形状中单个形状的编号）可从GroupShapes集合中返回单个形状。

javascript
```javascript
/*下例向 myDocument 添加三个三角形，将它们分成一组，设置整个组的颜色，然后只更改第二个三角形的颜色。*/
function test(){
let myDocument = Application.Worksheets.Item(1)
let shape1 = myDocument.Shapes
    shape1.AddShape(msoShapeIsoscelesTriangle, 10, 10, 100, 100).Name = "shpOne"
    shape1.AddShape(msoShapeIsoscelesTriangle, 150, 10, 100, 100).Name = "shpTwo"
    shape1.AddShape(msoShapeIsoscelesTriangle, 300, 10, 100, 100).Name = "shpThree"
    let shape2 = shape1.Range(["shpOne", "shpTwo", "shpThree"]).Group()
        shape2.Fill.PresetTextured(msoTextureBlueTissuePaper)
        shape2.GroupItems.Item(2).Fill.PresetTextured(msoTextureGreenMarble)
}
```


#### HiLoLines 对象

# [HiLoLines (对象)​](#hilolines-对象)

代表图表组中的高低点连线。

## [说明​](#说明)

高低点连线连接图表组内每一分类中的最高数据点和最底数据点。只有二维折线图组可以有高低点连线。此对象不是集合。没有代表单个高低点连线的对象；或者打开图表组中所有数据点的高低点连线，或者将其全部关闭。

如果HasHiLoLines属性是False，HiLoLines对象的大多数属性都会被禁用。

使用HiLoLines属性可返回HiLoLines对象。

## [示例​](#示例)

javascript
```javascript
/*下例使用 HasHiLowLines 属性将 HiLowLines 添加到工作表一上的嵌入图表一（该图表必须是折线图），并将高低点连线设置为蓝色。*/
function test() {
    Application.Worksheets.Item(1).ChartObjects(1).Activate()
    ActiveChart.ChartGroups(1).HasHiLoLines = true
    ActiveChart.ChartGroups(1).HiLoLines.Border.Color = RGB(0, 0, 255)
}
```

javascript
```javascript
/*本示例显示图表工作表 Chart1 中第一个图表组的高低点连线的透明度。*/
function test() {
    let chartGroup = Application.Charts.Item("Chart1").ChartObjects(1).Chart.ChartGroups(1)
    let chartFormat = chartGroup.HiLoLines.Format
    console.log(chartFormat.Line.Transparency)
}
```


#### Hyperlink 对象

# [Hyperlink (对象)​](#hyperlink-对象)

代表一个超链接。

## [说明​](#说明)

Hyperlink对象是Hyperlinks集合的成员。

使用Hyperlink属性可返回某个形状的超链接（一个形状只能有一个超链接）。

## [示例​](#示例)

javascript
```javascript
/*本示例加载附于第一个形状（第一张工作表中）的超链接*/
function test() {
    Application.Worksheets.Item(1).Shapes.Item(1).Hyperlink.Follow(true)
}
```

区域或工作表可以有多个超链接。使用Hyperlinks(index) 可返回单个Hyperlink对象，其中index是超链接编号。

javascript
```javascript
/*本示例打开活动工作表中 A1:B2 区域中的第二个超链接*/
function test() {
    Application.Worksheets.Item(1).Range("A1:B2").Hyperlinks.Item(2).Follow()
}
```


#### Hyperlinks 对象

# [Hyperlinks (对象)​](#hyperlinks-对象)

代表工作表或区域的超链接的集合。

## [说明​](#说明)

每个超链接都由一个Hyperlink对象代表。

使用Hyperlinks属性可返回Hyperlinks集合。

## [示例​](#示例)

javascript
```javascript
/*本示例检查并打开第一个工作表中包含"wps"单词所有的超链接。*/
function test() {
    let hyperlinks = Application.Worksheets.Item(1).Hyperlinks
    for (let i = 1; i <= hyperlinks.Count; i++) {
        if (hyperlinks.Item(i).Name.indexOf("wps") != -1) {
            hyperlinks.Item(i).Follow()
        }
    }
}
```

使用Add方法可创建一个超链接并将它添加到Hyperlinks集合。

javascript
```javascript
/*本示例向第一张工作表的单元格E5中添加一个超链接。*/
function test() {
    let sheet = Application.Worksheets.Item(1)
    sheet.Hyperlinks.Add(sheet.Range("E5"), "https://kingsoft.com")
}
```


#### Icon 对象

# [Icon (对象)​](#icon-对象)

代表用于条件格式规则的图标集中的单个图标。

## [说明​](#说明)

Icon对象是从IconSet对象的Item属性中返回的。

## [示例​](#示例)

javascript
```javascript
/*本示例显示第一张工作表上区域 B1:B10 中第一个图标集条件格式的图标集中所有图标的索引。*/
function test() {
    let iconSet = Worksheets.Item(1).Range("B1:B10").FormatConditions.Item(1).IconSet
    for (let i = 1; i <= iconSet.Count; i++) {
        console.log(iconSet.Item(i).Index)
    }
}
```

javascript
```javascript
/*本示例显示第一张工作表上区域 B1:B10 中第二个图标集条件格式的图标集中第一个图标的实际类型。*/
function test() {
    let icon = Worksheets.Item(1).Range("B1:B10").FormatConditions.Item(2).IconSet.Item(1)
    console.log(icon.constructor.name)
}
```


#### IconCriteria 对象

# [IconCriteria (对象)​](#iconcriteria-对象)

代表IconCriterion对象的集合。每个IconCriterion代表图标集条件格式规则中每个图标的值和阈值类型。

## [说明​](#说明)

IconCriteria集合是从IconSetCondition对象的IconCriteria属性中返回的。通过将索引传递到集合中，您可以访问集合中的每个IconCriterion对象。有关详细信息，请参阅示例。

## [示例​](#示例)

javascript
```javascript
/*本示例创建了一个代表测试分数的数字范围，然后对该范围应用了图标集条件格式规则。图标集的类型将从默认图标变为五箭头图标集。最后，将阈值类型从百分点修改为硬编码数字。*/
function test() {
    // Fill cells with sample data from 1 to 10
    ActiveSheet.Range("C1").Value2 = 55
    ActiveSheet.Range("C2").Value2 = 92
    ActiveSheet.Range("C3").Value2 = 88
    ActiveSheet.Range("C4").Value2 = 77
    ActiveSheet.Range("C5").Value2 = 66
    ActiveSheet.Range("C6").Value2 = 93
    ActiveSheet.Range("C7").Value2 = 76
    ActiveSheet.Range("C8").Value2 = 80
    ActiveSheet.Range("C9").Value2 = 79
    ActiveSheet.Range("C10").Value2 = 83
    ActiveSheet.Range("C11").Value2 = 66
    ActiveSheet.Range("C12").Value2 = 74
    Range("C1:C12").Select()

    // Create an icon set conditional format for the created sample data range
    let iconSet = Selection.FormatConditions.AddIconSetCondition()

    // Change the icon set to a five-arrow icon set
    iconSet.IconSet = ActiveWorkbook.IconSets(xl5Arrows)

    // The IconCriterion collection contains all IconCriteria
    // By indexing into the collection you can modify each criterion
    let iconCriteria2 = iconSet.IconCriteria.Item(2)
    iconCriteria2.Type = xlConditionValueNumber
    iconCriteria2.Value = 60
    iconCriteria2.Operator = 7

    let iconCriteria3 = iconSet.IconCriteria.Item(3)
    iconCriteria3.Type = xlConditionValueNumber
    iconCriteria3.Value = 70
    iconCriteria3.Operator = 7

    let iconCriteria4 = iconSet.IconCriteria.Item(4)
    iconCriteria4.Type = xlConditionValueNumber
    iconCriteria4.Value = 80
    iconCriteria4.Operator = 7

    let iconCriteria5 = iconSet.IconCriteria.Item(5)
    iconCriteria5.Type = xlConditionValueNumber
    iconCriteria5.Value = 90
    iconCriteria5.Operator = 7
}
```

javascript
```javascript
/*本示例显示活动工作表上区域 A1:A10 中第一个图标集条件格式的第二个IconCriterion的运算符是否为xlGreaterEqual。*/
function test() {
    let iconSet = ActiveSheet.Range("A1:A10").FormatConditions.Item(1)
    console.log(iconSet.IconCriteria.Item(2).Operator == xlGreaterEqual)
}
```


#### IconSet 对象

# [IconSet (对象)​](#iconset-对象)

代表用于图标集条件格式规则的单一图标集。

## [说明​](#说明)

IconSet对象是IconSets集合的子对象。

条件格式的图标集通过使用IconSetCondition对象的IconSet属性来分配。通过以Workbook对象的IconSets属性的索引形式传递XlIconSet枚举的其中一个常量，您可以将此属性设置为其中的一个内置图标集。有关详细信息，请参阅示例。

## [示例​](#示例)

javascript
```javascript
/*本示例创建了一个代表测试分数的数字范围，然后对该范围应用了图标集条件格式规则。图标集的类型将从默认图标变为五箭头图标集。最后，将阈值类型从百分点修改为硬编码数字。*/
function test() {
    // Fill cells with sample data from 1 to 10
    ActiveSheet.Range("C1").Value2 = 55
    ActiveSheet.Range("C2").Value2 = 92
    ActiveSheet.Range("C3").Value2 = 88
    ActiveSheet.Range("C4").Value2 = 77
    ActiveSheet.Range("C5").Value2 = 66
    ActiveSheet.Range("C6").Value2 = 93
    ActiveSheet.Range("C7").Value2 = 76
    ActiveSheet.Range("C8").Value2 = 80
    ActiveSheet.Range("C9").Value2 = 79
    ActiveSheet.Range("C10").Value2 = 83
    ActiveSheet.Range("C11").Value2 = 66
    ActiveSheet.Range("C12").Value2 = 74
    Range("C1:C12").Select()

    // Create an icon set conditional format for the created sample data range
    let iconSet = Selection.FormatConditions.AddIconSetCondition()

    // Change the icon set to a 5-arrow icon set
    iconSet.IconSet = ActiveWorkbook.IconSets(xl5Arrows)

    //The IconCriterion collection contains all of IconCriteria
    //By indexing into the collection you can modify each criteria
    let iconCriteria1 = iconSet.IconCriteria(1)
    iconCriteria1.Type = xlConditionValueNumber
    iconCriteria1.Value = 0
    iconCriteria1.Operator = 7

    let iconCriteria2 = iconSet.IconCriteria(2)
    iconCriteria2.Type = xlConditionValueNumber
    iconCriteria2.Value = 60
    iconCriteria2.Operator = 7

    let iconCriteria3 = iconSet.IconCriteria(3)
    iconCriteria3.Type = xlConditionValueNumber
    iconCriteria3.Value = 70
    iconCriteria3.Operator = 7

    let iconCriteria4 = iconSet.IconCriteria(4)
    iconCriteria4.Type = xlConditionValueNumber
    iconCriteria4.Value = 80
    iconCriteria4.Operator = 7

    let iconCriteria5 = iconSet.IconCriteria(5)
    iconCriteria5.Type = xlConditionValueNumber
    iconCriteria5.Value = 90
    iconCriteria5.Operator = 7

}
```

javascript
```javascript
/*本示例判断如果活动工作表上区域 B1:B10 中第一个图标集条件格式的图标集的名称是为三色旗，则显示该图标集中图标的数量。*/
function test() {
    let iconSet = ActiveSheet.Range("B1:B10").FormatConditions.Item(1).IconSet
    if (iconSet.ID == xl3Flags) {
        console.log(iconSet.Count)
    }
}
```


#### IconSetCondition 对象

# [IconSetCondition (对象)​](#iconsetcondition-对象)

代表图标集条件格式规则。

## [说明​](#说明)

所有条件格式对象均包含在FormatConditions集合对象中，该集合对象是Range集合的子项。您可以使用FormatConditions集合的Add方法或AddIconSetCondition方法创建图标集格式规则。

每个图标集包含三个、四个或五个图标。您可以使用Workbook对象的IconSets属性返回IconSets对象以指定其中一个内置图标集。然后按IconCriteria对象的成员将图标集中每个单独的图标分配给范围中的值的子集。阈值的类型也是由此对象指定的。

## [示例​](#示例)

javascript
```javascript
/*本示例创建了一个代表测试分数的数字范围，然后对该范围应用了图标集条件格式规则。图标集的类型将从默认图标变为五箭头图标集。最后，将阈值类型从百分点修改为硬编码数字。*/
function test() {
    // Fill cells with sample data from 1 to 10
    ActiveSheet.Range("C1").Value2 = 55
    ActiveSheet.Range("C2").Value2 = 92
    ActiveSheet.Range("C3").Value2 = 88
    ActiveSheet.Range("C4").Value2 = 77
    ActiveSheet.Range("C5").Value2 = 66
    ActiveSheet.Range("C6").Value2 = 93
    ActiveSheet.Range("C7").Value2 = 76
    ActiveSheet.Range("C8").Value2 = 80
    ActiveSheet.Range("C9").Value2 = 79
    ActiveSheet.Range("C10").Value2 = 83
    ActiveSheet.Range("C11").Value2 = 66
    ActiveSheet.Range("C12").Value2 = 74
    Range("C1:C12").Select()

    // Create an icon set conditional format for the created sample data range
    let iconSet = Selection.FormatConditions.AddIconSetCondition()

    // Change the icon set to a five-arrow icon set
    iconSet.IconSet = ActiveWorkbook.IconSets(xl5Arrows)

    //The IconCriterion collection contains all IconCriteria
    //By indexing into the collection you can modify each criterion
    let iconCriterion1 = iconSet.IconCriteria(1)
    iconCriterion1.Type = xlConditionValueNumber
    iconCriterion1.Value = 0
    iconCriterion1.Operator = 7
    let iconCriterion2 = iconSet.IconCriteria(2)
    iconCriterion2.Type = xlConditionValueNumber
    iconCriterion2.Value = 60
    iconCriterion2.Operator = 7
    let iconCriterion3 = iconSet.IconCriteria(3)
    iconCriterion3.Type = xlConditionValueNumber
    iconCriterion3.Value = 70
    iconCriterion3.Operator = 7
    let iconCriterion4 = iconSet.IconCriteria(4)
    iconCriterion4.Type = xlConditionValueNumber
    iconCriterion4.Value = 80
    iconCriterion4.Operator = 7
    let iconCriterion5 = iconSet.IconCriteria(5)
    iconCriterion5.Type = xlConditionValueNumber
    iconCriterion5.Value = 90
    iconCriterion5.Operator = 7
}
```

javascript
```javascript
/*本示例判断如果活动工作表上单元格区域 A1:A10 的第二个图标集条件格式使用的图标集的名称为xl3Arrows，则修改该条件格式所应用于的单元格区域，并将该图标集条件格式设置为仅展示图标。*/
function test() {
    let iconSet = ActiveSheet.Range("A1:A10").FormatConditions.Item(2)
    if (iconSet.IconSet.ID == xl3Arrows) {
        iconSet.ModifyAppliesToRange(Range("A1:A4"))
        iconSet.ShowIconOnly = true
    }
}
```


#### IconSets 对象

# [IconSets (对象)​](#iconsets-对象)

代表用于图标集条件格式规则的图标集的集合。

## [说明​](#说明)

条件格式的图标集通过使用IconSetCondition对象的IconSet属性来分配。通过以Workbook对象的IconSets属性的索引形式传递XlIconSet枚举的其中一个常量，您可以将此属性设置为其中的一个内置图标集。有关详细信息，请参阅示例。

## [示例​](#示例)

javascript
```javascript
/*本示例创建了一个代表测试分数的数字范围，然后对该范围应用了图标集条件格式规则。图标集的类型将从默认图标变为五箭头图标集。最后，将阈值类型从百分点修改为硬编码数字。*/
function test() {
    // Fill cells with sample data from 1 to 10
    ActiveSheet.Range("C1").Value2 = 55
    ActiveSheet.Range("C2").Value2 = 92
    ActiveSheet.Range("C3").Value2 = 88
    ActiveSheet.Range("C4").Value2 = 77
    ActiveSheet.Range("C5").Value2 = 66
    ActiveSheet.Range("C6").Value2 = 93
    ActiveSheet.Range("C7").Value2 = 76
    ActiveSheet.Range("C8").Value2 = 80
    ActiveSheet.Range("C9").Value2 = 79
    ActiveSheet.Range("C10").Value2 = 83
    ActiveSheet.Range("C11").Value2 = 66
    ActiveSheet.Range("C12").Value2 = 74
    Range("C1:C12").Select()

    // Create an icon set conditional format for the created sample data range
    let iconSet = Selection.FormatConditions.AddIconSetCondition()

    // Change the icon set to a 5-arrow icon set
    iconSet.IconSet = ActiveWorkbook.IconSets.Item(xl5Arrows)

    //The IconCriterion collection contains all of IconCriteria
    //By indexing into the collection you can modify each criteria
    let iconCriteria1 = iconSet.IconCriteria.Item(1)
    iconCriteria1.Type = xlConditionValueNumber
    iconCriteria1.Value = 0
    iconCriteria1.Operator = 7

    let iconCriteria2 = iconSet.IconCriteria.Item(2)
    iconCriteria2.Type = xlConditionValueNumber
    iconCriteria2.Value = 60
    iconCriteria2.Operator = 7

    let iconCriteria3 = iconSet.IconCriteria.Item(3)
    iconCriteria3.Type = xlConditionValueNumber
    iconCriteria3.Value = 70
    iconCriteria3.Operator = 7

    let iconCriteria4 = iconSet.IconCriteria.Item(4)
    iconCriteria4.Type = xlConditionValueNumber
    iconCriteria4.Value = 80
    iconCriteria4.Operator = 7

    let iconCriteria5 = iconSet.IconCriteria.Item(5)
    iconCriteria5.Type = xlConditionValueNumber
    iconCriteria5.Value = 90
    iconCriteria5.Operator = 7

}
```

javascript
```javascript
/*本示例显示第一张工作簿上图标集的数目。*/
function test() {
    let iconSets = Application.Workbooks.Item(1).IconSets
    console.log(`图标集的数目：${iconSets.Count}`)
}
```


#### Interior 对象

# [Interior (对象)​](#interior-对象)

代表一个对象的内部。

## [说明​](#说明)

使用Interior属性可返回Interior对象。

## [示例​](#示例)

javascript
```javascript
/*此示例将 Sheet1 中 A1 单元格内部设置为红色。*/
function test() {
    Application.Worksheets.Item("Sheet1").Range("A1").Interior.ColorIndex = 3
}
```

javascript
```javascript
/*本示例在活动工作表中 A1:C4 区域单元格的内部添加十字线图案。*/
function test() {
    ActiveSheet.Range("A1:C4").Interior.Pattern = xlPatternCrissCross
}
```


#### LeaderLines 对象

# [LeaderLines (对象)​](#leaderlines-对象)

代表图表的引导线。引导线将数据标签连接到数据点。

## [说明​](#说明)

该对象不是一个集合；没有表示单个引导线的对象。

此对象只适用于饼图。

使用LeaderLines属性可返回LeaderLines对象。

## [示例​](#示例)

javascript
```javascript
/*下例图表一上的数据系列一添加数据标签和蓝色引导线。如果看不到引导线，则该示例代码将失败。在这种情况下，可从饼图中手动拖出其中一个数据标签以便显示引导线。*/
function test() {
    let chartObjects2 = Application.Worksheets.Item(1).ChartObjects(1).Chart.SeriesCollection(1)
    chartObjects2.HasDataLabels = true
    chartObjects2.DataLabels.Position = xlLabelPositionBestFit
    chartObjects2.HasLeaderLines = true
    chartObjects2.LeaderLines.Border.ColorIndex = 5
}
```

javascript
```javascript
/*此示例显示图表工作表 Chart1 上的第三个数据系列的引导线粗细。*/
function test() {
    let series3 = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(3)
    console.log(series3.LeaderLines.Border.Weight)
}
```


#### Legend 对象

# [Legend (对象)​](#legend-对象)

代表图标中的图例。每个图表只能有一个图例。

## [说明​](#说明)

Legend对象包含一个或多个LegendEntry对象；每个LegendEntry对象都包含一个LegendKey对象。

除非HasLegend属性是True，否则无法看到图表的图例。如果该属性为False，Legend对象的属性和方法将会失败。

## [示例​](#示例)

javascript
```javascript
/*下例将第一张工作表上嵌入图表一的图例字形设置为加粗。*/
function test() {
    Application.Worksheets.Item(1).ChartObjects(1).Chart.Legend.Font.Bold = true
}
```

javascript
```javascript
/*此示例删除图表工作表 Chart1 上的第一个图表中的图例。*/
function test() {
    Application.Charts.Item("Chart1").ChartObjects(1).Chart.Legend.Delete()
}
```


#### LegendEntries 对象

# [LegendEntries (对象)​](#legendentries-对象)

指定的图表图例中所有LegendEntry对象的集合。

## [说明​](#说明)

每个图例项都有两部分：一部分是该项的文本，它是与该图例项相关联的数据系列或趋势线的名称；另一部分是项标志，它在图表中以直观的方式将图例项以及与之相关联的数据系列或趋势线链接起来。项标志及其相关系列或趋势线的格式属性包含在LegendKey对象中。

## [示例​](#示例)

javascript
```javascript
/*下例在嵌入式图表一中的图例项集合中循环，并更改这些图例项的字体颜色。*/
function test() {
    let lgd = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.Legend
    for (let i = 1; i <= lgd.LegendEntries().Count; i++) {
        lgd.LegendEntries(i).Font.ColorIndex = 5
    }
}
```

使用LegendEntries(index)（其中index是图例项索引号）可返回一个LegendEntry对象。不能按名称返回图例项。

图例项的编号代表图例项在图例中的位置。LegendEntries(1)位于图例的顶部；LegendEntries(LegendEntries.Count)位于图例的底部。

javascript
```javascript
/*下例将嵌入的第一个图表中图例顶部的图例项（这通常是第一个数据系列的图例）的文字字体设置为斜体。*/
function test() {
    Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.Legend.LegendEntries(1).Font.Italic = true
}
```


#### LegendEntry 对象

# [LegendEntry (对象)​](#legendentry-对象)

代表图表图例中的图例项。

## [说明​](#说明)

LegendEntry对象是LegendEntries集合的成员。LegendEntries集合包含图例中所有的LegendEntry对象。

每个图例项都有两部分：一部分是该项的文本，它是与该图例项相关联的数据系列或趋势线的名称；另一部分是项标志，它在图表中以直观的方式将图例项以及与之相关联的数据系列或趋势线链接起来。项标志及其相关数据系列或趋势线的格式设置属性包含在LegendKey对象中。

不能修改图例项的文字。LegendEntry对象支持字体格式，且可被删除。图例项不支持图案格式。图例项的位置和尺寸是固定的。

没有返回相应于某图例项的数据系列或趋势线的直接方法。

在删除了图例项之后，唯一的恢复方法是：通过将图表的HasLegend属性设置为False，然后再设回True，从而删除并重新创建包含这些图例项的图例。

使用LegendEntries(index)（其中index是图例项索引号）可返回一个LegendEntry对象。不能按名称返回图例项。

图例项的编号代表图例项在图例中的位置。LegendEntries(1)位于图例的顶部；LegendEntries(LegendEntries.Count)位于图例的底部。

## [示例​](#示例)

javascript
```javascript
/*下例更改工作表“Sheet1”上嵌入的第一个图表中图例顶部的图例项（这通常是第一个数据系列的图例）的文字字体。*/
function test() {
    Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.Legend.LegendEntries(1).Font.Italic = true
}
```

javascript
```javascript
/*下例显示 Chart1 上第二个图例项的高度。*/
function test() {
    let legend = Application.Charts.Item("Chart1").ChartObjects(1).Chart.Legend
    console.log(legend.LegendEntries(2).Height)
}
```


#### LegendKey 对象

# [LegendKey (对象)​](#legendkey-对象)

代表图表图例中的图例标示。

## [说明​](#说明)

每个图例项标示都是一个图形，它将图例项标示以及与之相关联的图表中的系列或趋势线链接起来。图例项标示链接到与之相关联的系列或趋势线的方式为：更改一个的格式会同时更改另一个的格式。

使用LegendKey属性可返回LegendKey对象。比如更改图例顶部的图例项标示的标志背景色，这样会同时更改与该图例项相关的系列中每一个数据点的格式。相关的系列必须支持数据标志。

## [示例​](#示例)

javascript
```javascript
/*下例显示 Sheet1 中的第一个图表的第二个图例项的图例标示左边缘到图表区左边缘的距离。*/
function test() {
    let legend = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.Legend
    console.log(legend.LegendEntries(2).LegendKey.Left)
}
```

javascript
```javascript
/*本示例将 Chart1 上第一个图例项的图例标示设置成有阴影。*/
function test() {
    Application.Charts.Item("Chart1").ChartObjects(1).Chart.Legend.LegendEntries(1).LegendKey.Shadow = true
}
```


#### LineFormat 对象

# [LineFormat (对象)​](#lineformat-对象)

代表线条和箭头格式。

## [说明​](#说明)

对于线条，LineFormat对象包含该线条自身的格式信息；对于有边界的形状，该对象包含形状的边界的格式信息。

使用Line属性可返回一个LineFormat对象。

## [示例​](#示例)

javascript
```javascript
/*下例给第一张工作表添加一条蓝色虚线。在该线的起点有一个短而窄的椭圆，在该线的终点有一个长而宽的三角形。*/
function test() {
    let worksheet = Application.Worksheets.Item(1)
    let line = worksheet.Shapes.AddLine(100, 100, 200, 300).Line
    line.DashStyle = msoLineDashDotDot
    line.ForeColor.RGB = RGB(50, 0, 128)
    line.BeginArrowheadLength = msoArrowheadShort
    line.BeginArrowheadStyle = msoArrowheadOval
    line.BeginArrowheadWidth = msoArrowheadNarrow
    line.EndArrowheadLength = msoArrowheadLong
    line.EndArrowheadStyle = msoArrowheadTriangle
    line.EndArrowheadWidth = msoArrowheadWide
}
```

javascript
```javascript
/*本示例显示活动工作表的第一个形状线条的样式是否为双细线。*/
function test() {
    let line = ActiveSheet.Shapes.Item(1).Line
    console.log(line.Style == msoLineThinThin)
}
```


#### LinearGradient 对象

# [LinearGradient (对象)​](#lineargradient-对象)

LinearGradient对象沿特定角度以线性方式在一系列颜色间转换。

## [说明​](#说明)

当使用LinearGradient对象时，应考虑以下几点：

试图访问不具有现有渐变填充的
Interior
对象的 Gradient 属性会引起运行时错误。访问 Gradient 属性之前请注意
Interior.Pattern
属性。
如果将 Interior.Pattern从渐变类型更改为非渐变类型，Gradient 对象将采用默认值。
## [示例​](#示例)

javascript
```javascript
/*下例将活动工作表上区域 B1:B10 中线性渐变的填充角度设置为100。*/
function test() {
    let linearGradient = ActiveSheet.Range("B1:B10").Interior.Gradient
    linearGradient.Degree = 100
}
```

javascript
```javascript
/*下例显示第一张工作表上区域 B1:B10 中线性渐变的颜色停止点数量。*/
function test() {
    let linearGradient = Worksheets.Item(1).Range("B1:B10").Interior.Gradient
    console.log(linearGradient.ColorStops.Count)
}
```


#### ListColumn 对象

# [ListColumn (对象)​](#listcolumn-对象)

代表表格中的一列。

## [说明​](#说明)

ListColumn对象是ListColumns集合的成员。ListColumns集合包含表格中的所有列（ListObject对象）。

使用ListObject对象的 ListColumns 属性可返回一个ListColumns集合。

## [示例​](#示例)

javascript
```javascript
/*下例给活动工作簿的第一个工作表的默认 ListObject 对象添加新 ListColumn 对象。由于未指定位置，因此在最右边添加一个新列，并显示列名。*/
function test() {
    let worksheet = Application.ActiveWorkbook.Worksheets.Item("Sheet1")
    let listColumn = worksheet.ListObjects.Item(1).ListColumns.Add()
    console.log(listColumn.Name)
}
```

javascript
```javascript
/*本示例将活动工作表的第一个 ListObject 对象的第二列删除。*/
function test() {
    let listObj = Application.ActiveSheet.ListObjects.Item(1)
    listObj.ListColumns.Item(2).Delete()
}
```


#### ListColumns 对象

# [ListColumns (对象)​](#listcolumns-对象)

指定的ListObject对象中所有ListColumn对象的集合。

## [说明​](#说明)

每个ListColumn对象都代表表格中的一列。

该列的名称会自动生成。在添加完该列后可更改其名称。

使用 ListObject 对象的ListColumns属性可返回ListColumns集合。

## [示例​](#示例)

javascript
```javascript
/*下例给工作簿的第一张工作表的默认 ListObject 对象添加一个新列。由于未指定位置，因此在最右边添加一个新列。*/
function test() {
    let newColumn = Application.Worksheets.Item(1).ListObjects.Item(1).ListColumns.Add()
}
```

javascript
```javascript
/*本示例将活动工作表的第一个 ListObject 对象的第四列删除。*/
function test() {
    let listObj = Application.ActiveSheet.ListObjects.Item(1)
    listObj.ListColumns.Item(4).Delete()
}
```


#### ListObject 对象

# [ListObject (对象)​](#listobject-对象)

代表工作表中的表格。

## [说明​](#说明)

ListObject对象是ListObjects集合的成员。ListObjects集合包含工作表上所有的列表对象。

使用Worksheet对象的 ListObjects 属性可返回一个ListObjects集合。

## [示例​](#示例)

javascript
```javascript
/*本示例给活动工作簿的第一张工作表的默认 ListObject 对象添加新的 ListRow 对象。*/
function test() {
    let sheet = Application.ActiveWorkbook.Worksheets.Item(1)
    sheet.ListObjects.Item(1).ListRows.Add()
}
```

javascript
```javascript
/*本示例为工作表 Sheet1 上第一张列表添加新列，并设置该列表的表样式。*/
function test() {
    let listObj = Application.Worksheets.Item("Sheet1").ListObjects.Item(1)
    listObj.ListColumns.Add()
    listObj.TableStyle = "TableStylePreset5_Accent1"
}
```


#### ListObjects 对象

# [ListObjects (对象)​](#listobjects-对象)

工作表上所有ListObject对象的集合。每个ListObject对象都代表工作表中的一个表格。

## [说明​](#说明)

使用 Worksheet 对象的ListObjects属性可返回ListObjects集合。

## [示例​](#示例)

javascript
```javascript
/*本示例创建一个新 ListObjects 集合，该集合代表第一张工作表中所有的表格，并显示表格数量。*/
function test() {
    let listObjects = Application.Worksheets.Item(1).ListObjects
    console.log(listObjects.Count)
}
```

javascript
```javascript
/*本示例显示活动工作表上第二个列表的 SourceType 属性是否为 xlSrcRange。*/
function test() {
    let listObj = Application.ActiveSheet.ListObjects.Item(2)
    console.log(listObj.SourceType == xlSrcRange)
}
```


#### ListRow 对象

# [ListRow (对象)​](#listrow-对象)

代表表格中的一行。ListRow对象是ListRows集合的成员。

## [说明​](#说明)

ListRows 集合包含列表对象中的所有行。ListRows

使用ListObject对象的ListRows属性可返回一个ListRows集合。

## [示例​](#示例)

javascript
```javascript
/*下例给活动工作簿的第一张工作表的默认 ListObject 对象添加新 ListRow 对象。由于未指定位置，因此在表格结束处添加一个新行。*/
function test() {
    let worksheet = Application.ActiveWorkbook.Worksheets.Item("Sheet1")
    let listRow = worksheet.ListObjects.Item(1).ListRows.Add()
}
```

javascript
```javascript
/*本示例将第一张工作表的第一个 ListObject 对象的第三行删除。*/
function test() {
    let listObj = Worksheets.Item(1).ListObjects.Item(1)
    listObj.ListRows.Item(3).Delete()
}
```


#### ListRows 对象

# [ListRows (对象)​](#listrows-对象)

指定的ListObject对象中ListRow对象的集合。

## [说明​](#说明)

每一个ListRow对象都代表表格中的一行。

使用ListObject对象的ListRows属性可返回ListRows集合。

## [示例​](#示例)

javascript
```javascript
/*下例给第一张工作表的默认 ListObject 对象添加新行。由于未指定位置，因此在表格结束处添加一个新行。*/
function test() {
    let row = Application.Worksheets.Item(1).ListObjects.Item(1).ListRows.Add()
}
```

javascript
```javascript
/*本示例显示活动工作表的第一个 ListObject 对象中行的数量。*/
function test() {
    let listObj = Application.ActiveSheet.ListObjects.Item(1)
    console.log(listObj.ListRows.Count)
}
```


#### Name 对象

# [Name (对象)​](#name-对象)

代表单元格区域的定义名。名称可以是内置名称（如“Database”、“Print_Area”和“Auto_Open”）或自定义名称。

## [说明​](#说明)

### [应用程序、工作簿和 Worksheet 对象​](#应用程序、工作簿和-worksheet-对象)

Name对象是Application、Workbook和Worksheet对象的Names集合的成员。使用Names(index)（其中index是名称索引号或定义名称）可返回一个Name对象。

索引号表明名称在集合中的位置。名称按字母顺序从 a 到 z 放置，不区分大小写。

### [Range 对象​](#range-对象)

虽然Range对象可以有多个名称，但Range对象没有Names集合。对Range对象调用Name属性可从名称列表（按字母顺序排序）中返回第一个名称。

## [示例​](#示例)

javascript
```javascript
/*本示例显示应用程序集合中第一个名称的单元格引用。*/
function test() {
    console.log(Application.Application.Names.Item(1).RefersTo)
}
```

javascript
```javascript
/*下例从活动工作簿中删除名称“mySortRange”。*/
function test() {
    Application.ActiveWorkbook.Names.Item("mySortRange").Delete()
}
```

javascript
```javascript
/*使用 Name 属性可返回或设置名称本身的文本。本示例更改活动工作簿中第一个 Name 对象的名称。*/
function test() {
    Application.Names.Item(1).Name = "stock_values"
}
```

javascript
```javascript
/*以下示例为第一张工作表上分配给单元格 A1:B1 的第一个名称设置 Visible 属性。*/
function test() {
    Application.Worksheets.Item(1).Range("A1:B1").Name.Visible = false
}
```


#### Names 对象

# [Names (对象)​](#names-对象)

应用程序或工作簿中所有Name对象的集合。

## [说明​](#说明)

每一个Name对象都代表一个单元格区域的定义名称。名称可以是内置名称（如“Database”、“Print_Area”和“Auto_Open”）或自定义名称。

RefersTo参数必须以 A1 样式表示法指定，包括必要时使用的美元符 ($)。例如，如果在 Sheet1 上选定了单元格 A10，并且通过将RefersTo参数“=Sheet1!A1:B1”而定义了一个名称，那么该新名称实际上指向单元格区域 A10:B10（因为指定的是相对引用）。若要指定绝对引用，请使用“=Sheet1!$A$1:$B$1”。

## [示例​](#示例)

使用Names属性可返回Names集合。下例创建活动工作簿中所有名称及其引用地址的列表。

javascript
```javascript
function test() {
    let nms = Application.ActiveWorkbook.Names
    let wks = Application.Worksheets.Item(1)
    for(let r = 1; r <= nms.Count; r++){
        wks.Cells.Item(r, 2).Value2 = nms.Item(r).Name
        wks.Cells.Item(r, 3).Value2 = nms.Item(r).RefersToRange.Address
    }
}
```

使用Add方法可创建一个名称并将它添加到集合。下例创建一个新名称，该名称引用名为“Sheet1”的工作表上的单元格 A1:C20。

javascript
```javascript
function test() {
    Application.Names.Add ("test1", "=sheet1!$a$1:$c$20")
}
```

使用Names(index)（其中index是名称索引号或定义名称）可返回一个Name对象。下例从活动工作簿中删除名称“mySortRange”。

javascript
```javascript
function test() {
    Application.ActiveWorkbook.Names.Item("mySortRange").Delete()
}
```


#### Outline 对象

# [Outline (对象)​](#outline-对象)

代表工作表上的分级显示。

## [示例​](#示例)

javascript
```javascript
/*下例将 Sheet4 上的分级显示设置为只显示第一级。*/
function test() {
    Worksheets("Sheet4").Outline.ShowLevels(1)
}
```

javascript
```javascript
/*此示例显示活动工作表分级显示是否使用自动样式。*/
function test() {
    console.log(ActiveSheet.Outline.AutomaticStyles)
}
```


#### PictureFormat 对象

# [PictureFormat (对象)​](#pictureformat-对象)

包含应用于图片和 OLE 对象的属性和方法。

## [说明​](#说明)

LinkFormat对象只包含应用于链接 OLE 对象的属性和方法。OLEFormat对象包含应用于 OLE 对象的属性和方法，无论这些对象是不是链接对象。

## [示例​](#示例)

javascript
```javascript
/*下例设置了第一张工作表中第一个形状的亮度、对比度和颜色的变换，而且在该形状的底部裁剪了 18 磅。要使此示例执行，则第一个形状必须是图片或 OLE 对象。*/
function test() {
    let worksheet = Worksheets.Item(1)
    let pictureFormat = worksheet.Shapes.Item(1).PictureFormat
    pictureFormat.Brightness = 0.3
    pictureFormat.Contrast = 0.7
    pictureFormat.ColorType = msoPictureGrayscale
    pictureFormat.CropBottom = 18
}
```

javascript
```javascript
/*此示例在活动工作表的第二个形状图片的右侧裁剪 10 磅。*/
function test() {
    let shape = ActiveSheet.Shapes.Item(2)
    shape.PictureFormat.CropRight = 10
}
```


#### PivotAxis 对象

# [PivotAxis (对象)​](#pivotaxis-对象)

PivotAxis对象用于在数据透视表中进行不对称深化。

## [说明​](#说明)

PivotAxis对象包含诸如PivotRowAxis和PivotRowAxis之类的属性，用于处理数据透视表中的行和列。

## [示例​](#示例)

javascript
```javascript
/*本示例显示活动工作表上数据透视表中列轴上第二条数据透视线的PivotLineCells属性中第一个PivotCell的数据透视表项的值。*/
function test() {
    let pvtLine = ActiveSheet.Range("I1").PivotTable.PivotColumnAxis.PivotLines(2)
    console.log(pvtLine.PivotLineCells.Item(1).PivotItem.Value)
}
```

javascript
```javascript
/*本示例显示活动工作表上第一张数据透视表中列轴上数据透视线的数量。*/
function test() {
    console.log(ActiveSheet.PivotTables(1).PivotColumnAxis.PivotLines.Count)
}
```


#### PivotCell 对象

# [PivotCell (对象)​](#pivotcell-对象)

代表数据透视表中的一个单元格。

## [说明​](#说明)

使用Range集合的PivotCell属性可返回一个PivotCell对象。

返回PivotCell对象后，可以使用ColumnItems或RowItems属性来确定PivotItems集合，对应于代表所选编号的列轴或行轴上的项目。

返回了PivotCell对象后，可以使用PivotCellType属性来确定某个特定区域是什么单元格类型。

## [示例​](#示例)

javascript
```javascript
/*本示例确定数据透视表中的单元格 J6 是否是一个数据项，并通知用户。本示例假定数据透视表位于活动工作表上。如果单元格 J6 不在数据透视表中，则本示例处理运行错误。*/
function test() {
    try {
        // Determine if cell J6 is a data item in the PivotTable.
        if (Application.Range("J6").PivotCell.PivotCellType == xlPivotCellValue) {
            console.log("The cell at J6 is a data item.")
        } else {
            console.log("The cell at J6 is not a data item.")
        }
    }
    catch (exception) {
        console.log("The chosen cell is not in a PivotTable.")
    }
}
```

javascript
```javascript
/*此示例确定单元格 L13 的数据项所在的列字段。然后判断列字段标题是否与“name”相匹配，并通知用户。此示例假定数据透视表位于活动工作表上，并且工作表的 L 列包含数据透视表的列字段。*/
function test() {
    // Determine if there is a match between the item and column field.
    if (Application.Range("L13").PivotCell.ColumnItems.Item(1).Parent.Name == "name") {
        console.log("Item in L13 is a member of the 'name' column field.")
    } else {
        console.log("Item in L13 is not a member of the 'name' column field.")
    }
}
```


#### PivotField 对象

# [PivotField (对象)​](#pivotfield-对象)

代表数据透视表中的一个字段。

## [说明​](#说明)

PivotField对象是PivotFields集合的成员。PivotFields集合包含数据透视表中的所有字段，包括隐藏字段。

下列属性返回数据透视表字段的子集，在某些情况下，使用这些属性会更为方便：

ColumnFields
属性
DataFields
属性
HiddenFields
属性
PageFields
属性
RowFields
属性
VisibleFields
属性
使用PivotFields(index)（其中index是字段名称或索引号）可返回一个PivotField对象。

## [示例​](#示例)

javascript
```javascript
/*本示例使 Sheet1 上第一张数据透视表中的字段“id”成为行字段。*/
function test() {
    Application.Worksheets.Item("Sheet1").PivotTables(1).PivotFields("id").Orientation = xlRowField
}
```

javascript
```javascript
/*本示例删除工作表 Sheet1 中数据透视表应用于字段“score”的所有筛选。*/
function test() {
    let pvtfield = Worksheets.Item("Sheet1").Range("I1").PivotTable.PivotFields("score")
    pvtfield.ClearAllFilters()
}
```


#### PivotFields 对象

# [PivotFields (对象)​](#pivotfields-对象)

指定的数据透视表中所有PivotField对象的集合。

## [说明​](#说明)

下列存取器方法返回数据透视表字段的子集，在某些情况下，使用这些属性会更为方便：

ColumnFields
属性
DataFields
属性
HiddenFields
属性
PageFields
属性
RowFields
属性
VisibleFields
属性
使用PivotTable对象的PivotFields方法可返回PivotFields集合。

## [示例​](#示例)

javascript
```javascript
/*本示例显示 Sheet1 上第一张数据透视表中的字段名称。*/
function test() {
    let pTab = Worksheets.Item("Sheet1").PivotTables(1)
    for (let i = 1; i <= pTab.PivotFields().Count; i++) {
        console.log(pTab.PivotFields(i).Name)
    }
}
```

使用PivotFields(index)（其中index是字段名称或索引号）可返回一个PivotField对象。

javascript
```javascript
/*本示例使 Sheet1 上第一张数据透视表中的字段“name”成为行字段。*/
function test() {
    Worksheets.Item("sheet1").PivotTables(1).PivotFields("name").Orientation = xlRowField
}
```


#### PivotFilter 对象

# [PivotFilter (对象)​](#pivotfilter-对象)

PivotFilter 应用于PivotField对象。

## [说明​](#说明)

开发人员可以选择命名筛选器以供引用，因为索引不可靠。DataField属性指定要基于值筛选器的 PivotField。

## [示例​](#示例)

javascript
```javascript
/*本示例显示第一张工作表中数据透视表的字段“id”的第一个筛选器的参数。*/
function test() {
    let pvtField = Worksheets.Item(1).Range("I1").PivotTable.PivotFields("id")
    console.log(pvtField.PivotFilters.Item(1).Value1)
}
```

javascript
```javascript
/*本示例显示活动工作表中第一张数据透视表的字段“score”的第二个筛选器是否是活动的。*/
function test() {
    console.log(ActiveSheet.PivotTables(1).PivotFields("score").PivotFilters.Item(2).Active)
}
```


#### PivotFilters 对象

# [PivotFilters (对象)​](#pivotfilters-对象)

PivotFilters对象是PivotFilter对象的集合。

## [说明​](#说明)

PivotFilters集合包含用于添加新筛选器、计算集合中现有筛选器个数以及引用特定PivotFilter对象的属性和方法。

## [示例​](#示例)

javascript
```javascript
/*本示例为活动工作表的数据透视表的的字段“date”添加筛选器。*/
function test() {
    ActiveSheet.Range("I1").PivotTable.PivotFields("date").PivotFilters.Add2(xlDateThisWeek)
}
```

javascript
```javascript
/*本示例显示工作表 Sheet1 中数据透视表字段“name”的第一个筛选器中作为筛选依据的值字段。*/
function test() {
    let pvtFilter = Worksheets.Item("Sheet1").Range("I1").PivotTable.PivotFields("name").PivotFilters.Item(1)
    console.log(pvtFilter.DataField.Value)
}
```


#### PivotFormula 对象

# [PivotFormula (对象)​](#pivotformula-对象)

代表在数据透视表中用于计算的公式。

## [说明​](#说明)

本对象及其相关属性和方法对于 OLAP（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源无效，这是因为它不支持计算字段和计算项。

使用PivotFormulas(index)（其中index是公式左侧的公式号或字符串）可返回PivotFormula对象。

## [示例​](#示例)

javascript
```javascript
/*本示例将更改公式一（第一张工作表的第一个数据透视表中）的索引号，使其在公式二计算完毕后再进行计算。*/
function test() {
    Worksheets.Item(1).PivotTables(1).PivotFormulas.Item(1).Index = 2
}
```

javascript
```javascript
/*本示例删除第一张工作表中数据透视表的第二个公式。*/
function test() {
    let pvtFormula = Worksheets.Item(1).Range("I1").PivotTable.PivotFormulas(2)
    pvtFormula.Delete()
}
```


#### PivotFormulas 对象

# [PivotFormulas (对象)​](#pivotformulas-对象)

代表数据透视表的公式的集合。每个公式都由一个PivotFormula对象代表。

## [说明​](#说明)

本对象及其相关属性和方法对于 OLAP（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源无效，这是因为它不支持计算字段和计算项。

使用PivotFormulas属性可返回PivotFormulas集合。

## [示例​](#示例)

javascript
```javascript
/*本示例为活动工作表上的第一张数据透视表创建一个数据透视表公式列表。*/
function test() {
    let r = 10
    for (let i = 1; i <= ActiveSheet.PivotTables(1).PivotFormulas.Count; i++) {
        Cells.Item(r, 1).Value2 = ActiveSheet.PivotTables(1).PivotFormulas.Item(i).Formula
        r++
    }
}
```

javascript
```javascript
/*本示例显示活动工作表中第一张数据透视表的公式的数量。*/
function test() {
    console.log(ActiveSheet.PivotTables(1).PivotFormulas.Count)
}
```


#### PivotItem 对象

# [PivotItem (对象)​](#pivotitem-对象)

代表数据透视表字段中的项目。

## [说明​](#说明)

这些项目是某个字段类型中的各个数据项。PivotItem对象是PivotItems集合的成员。PivotItems集合包含某个PivotField对象中所有的项目。使用PivotItems(index)（其中index是项目索引号或名称）可返回一个PivotItem对象。

## [示例​](#示例)

javascript
```javascript
/*本示例隐藏“Sheet1”上第一张数据透视表中“name”字段中包含“张一”的所有数据项。*/
function test() {
    Worksheets.Item("Sheet1").PivotTables(1).PivotFields("name").PivotItems("张一").Visible = false
}
```

javascript
```javascript
/*本示例显示活动工作表上数据透视表的字段“name”中未展示明细数据的数据项的名称。*/
function test() {
    let pvtItems = ActiveSheet.Range("I1").PivotTable.PivotFields("name").PivotItems()
    for (let i = 1; i <= pvtItems.Count; i++) {
        if (pvtItems.Item(i).ShowDetail == false) {
            console.log(pvtItems.Item(i).Name)
        }
    }
}
```


#### PivotItemList 对象

# [PivotItemList (对象)​](#pivotitemlist-对象)

指定的数据透视表中所有PivotItem对象的集合。

## [说明​](#说明)

每个PivotItem代表数据透视表字段中的一个项。

使用PivotCell对象的RowItems或ColumnItems属性可返回一个PivotItemList集合。

一旦返回了PivotItemList集合，您就可以使用Item方法来标识某个特定的PivotItem列表。

## [示例​](#示例)

javascript
```javascript
/*本示例显示包含单元格 L10 的数据透视项行轴上第一个数据项的标签文本。*/
function test() {
    // Identify contents associated with PivotItemList.
    console.log("Contents associated with cell L10: " + Application.Range("L10").PivotCell.RowItems.Item(1).Caption)
}
```

javascript
```javascript
/*本示例将工作表 Sheet1 上包含单元格 L4 的数据透视项行轴上第一个数据项的详细信息设置为不可见。*/
function test() {
    let pvtItem = Application.Worksheets.Item("Sheet1").Range("L4").PivotCell.RowItems.Item(1)
    pvtItem.ShowDetail = false
}
```


#### PivotItems 对象

# [PivotItems (对象)​](#pivotitems-对象)

数据透视表字段中所有PivotItem对象的集合。

## [说明​](#说明)

这些项目是某个字段类型中的各个数据项。

使用PivotItems方法可返回PivotItems集合。

## [示例​](#示例)

javascript
```javascript
/*本示例为 Sheet1 上第一张数据透视表创建那些字段中包含的字段名和项目的枚举列表。*/
function test() {
    Worksheets.Item("Sheet2").Activate()
    let pTab = Worksheets.Item("Sheet1").PivotTables(1)
    let c = 1
    for (let i = 1; i <= pTab.PivotFields().Count; i++) {
        let r = 1
        Cells.Item(r, c).Value2 = pTab.PivotFields(i).Name
        r++
        for (let x = 1; x <= pTab.PivotFields(i).PivotItems().Count; x++) {
            Cells.Item(r, c).Value2 = pTab.PivotFields(i).PivotItems(x).Name
            r++
        }
        c++
    }
}
```

使用PivotItems(index)（其中index是项目索引号或名称）可返回一个PivotItem对象。

javascript
```javascript
/*此示例隐藏工作表 Sheet1 中第一张数据透视表的字段“id”的数据项“4”。*/
function test() {
    let pvtItems = Worksheets.Item("sheet1").PivotTables(1).PivotFields("id").PivotItems()
    pvtItems.Item("4").Visible = false
}
```


#### PivotLayout 对象

# [PivotLayout (对象)​](#pivotlayout-对象)

代表数据透视图报表中字段的位置。

## [说明​](#说明)

使用PivotLayout属性可返回一个PivotLayout对象。

## [示例​](#示例)

javascript
```javascript
/*本示例创建在图表工作表 Chart1 上数据透视图报表中所使用的数据透视表字段名称的列表。*/
function test() {
    let newSheet = Worksheets.Add()
    let intRow = 1
    for (let i = 1; i <= Charts.Item("Chart1").PivotLayout.PivotFields().Count; i++) {
        newSheet.Cells.Item(intRow, 1).Value2 = Charts.Item("Chart1").PivotLayout.PivotFields(i).Caption
        intRow++
    }
}
```

javascript
```javascript
/*此示例显示工作表 Sheet1 上第一张数据透视图对应的数据透视表的第三个字段的名称。*/
function test() {
    let pvtTable = Worksheets.Item("Sheet1").ChartObjects(1).Chart.PivotLayout.PivotTable
    console.log(pvtTable.PivotFields(3).Name)
}
```


#### PivotLine 对象

# [PivotLine (对象)​](#pivotline-对象)

PivotLine对象是 ET 数据透视表中的行或列的线条。

## [说明​](#说明)

PivotLine 只包含可见项，因此PivotLine集合中不存在折叠的项目子项以及隐藏级别中的项目。

PivotLine 在所有位置始终具有一个 PivotItem。这意味着与普通 PivotLine 相比，代表数据透视表中分类汇总的 PivotLine 包含较少的 PivotItem。

## [示例​](#示例)

javascript
```javascript
/*本示例显示工作表 Sheet1 中第一张数据透视表中列轴上第二条数据透视线的PivotLineCells属性中PivotCell对象的数量。*/
function test() {
    let pvtLine = Worksheets.Item("Sheet1").PivotTables(1).PivotColumnAxis.PivotLines(2)
    console.log(pvtLine.PivotLineCells.Count)
}
```

javascript
```javascript
/*本示例显示活动工作表上数据透视表中列轴上第九条数据透视线的类型是否为“xlPivotLineGrandTotal”。*/
function test() {
    let pvtLine = ActiveSheet.Range("I1").PivotTable.PivotColumnAxis.PivotLines(9)
    console.log(pvtLine.LineType == xlPivotLineGrandTotal)
}
```


#### PivotLineCells 对象

# [PivotLineCells (对象)​](#pivotlinecells-对象)

特定 PivotLine 的PivotCell对象的集合。

## [说明​](#说明)

使用PivotLineCells(index)方法可以返回或指定集合中特定PivotCell对象的位置。您也可以指定PivotField对象或 PivotField 的名称以返回单个PivotCell对象。

## [示例​](#示例)

javascript
```javascript
/*本示例显示活动工作表上数据透视表列轴上第二条数据透视线的PivotLineCells属性的PivotCell对象的数量。*/
function test() {
    let pvtLine = ActiveSheet.Range("I1").PivotTable.PivotColumnAxis.PivotLines(2)
    console.log(pvtLine.PivotLineCells.Count)
}
```

javascript
```javascript
/*本示例显示工作表 Sheet1 上第一张数据透视表行轴上第一条数据透视线PivotLineCells属性的第一个PivotCell对象对应的“PivotTable”实体类型是否为“xlPivotCellPivotItem”。*/
function test() {
    let pvtLine = Worksheets.Item("Sheet1").PivotTables(1).PivotRowAxis.PivotLines(1)
    console.log(pvtLine.PivotLineCells.Item(1).PivotCellType == xlPivotCellPivotItem)
}
```


#### PivotLines 对象

# [PivotLines (对象)​](#pivotlines-对象)

PivotLines对象是数据透视表中线条的集合，其中包含数据透视表中行或列上的所有线条。每个线条都是一组 PivotCells。

## [示例​](#示例)

javascript
```javascript
/*本示例显示活动工作表上数据透视表中行轴上第一条数据透视线的PivotLineCells属性第二个PivotCell对应数据透视表项的名称。*/
function test() {
    let pvtLine = ActiveSheet.Range("I1").PivotTable.PivotRowAxis.PivotLines.Item(1)
    console.log(pvtLine.PivotLineCells.Item(2).PivotItem.Name)
}
```

javascript
```javascript
/*本示例显示活动工作表上数据透视表中列轴上数据透视线数量。*/
function test() {
    let pvtLines = ActiveSheet.Range("I1").PivotTable.PivotColumnAxis.PivotLines
    console.log(`列轴上数据透视线的数量:${pvtLines.Count}`)
}
```


#### PivotTable 对象

# [PivotTable (对象)​](#pivottable-对象)

代表工作表上的数据透视表。

## [说明​](#说明)

PivotTable对象是PivotTables集合的成员。PivotTables集合包含某一张工作表上的所有PivotTable对象。

因为对数据透视表进行编程可能会很复杂，所以，最方便的做法是将数据透视表操作录制到宏中，然后再修订所录制的宏代码。

使用PivotTables(index)（其中index是数据透视表索引号或名称）可返回一个PivotTable对象。

## [示例​](#示例)

javascript
```javascript
/*本示例使 Sheet1 上第一张数据透视表中的字段“age”成为行字段。*/
function test() {
    Worksheets.Item("Sheet1").PivotTables(1).PivotFields("age").Orientation = xlRowField
}
```

javascript
```javascript
/*本示例使第一张工作表上的第一个数据透视表使用合并单元格外部行项、列项、分类汇总和总计的标志。*/
function test() {
    Application.Worksheets.Item(1).PivotTables(1).MergeLabels = true
}
```


#### PivotTables 对象

# [PivotTables (对象)​](#pivottables-对象)

指定的工作表中所有PivotTable对象的集合。

## [说明​](#说明)

因为对数据透视表进行编程可能会很复杂，所以，最方便的做法是将数据透视表操作录制到宏中，然后再修订所录制的宏代码。

使用PivotTables方法可返回PivotTables集合。

## [示例​](#示例)

javascript
```javascript
/*本示例显示工作表 Sheet1 中数据透视表的数量。*/
function test() {
    console.log(Worksheets.Item("Sheet1").PivotTables().Count)
}
```

javascript
```javascript
/*本示例基于工作表 Sheet1 中第一张数据透视表的缓存在新工作表上创建新数据透视表。*/
function test() {
    Worksheets.Add().Activate()
    let pvtCache = Worksheets.Item("Sheet1").PivotTables(1).PivotCache()
    ActiveSheet.PivotTables().Add(pvtCache, Range("A1"), "student")
}
```

使用PivotTables(index)（其中index是数据透视表索引号或名称）可返回一个PivotTable对象。

javascript
```javascript
/*本示例将工作表 Sheet1 中第一个数据透视表上的 score 字段设置为行字段。*/
function test() {
    Worksheets.Item("sheet1").PivotTables().Item(1).PivotFields("score").Orientation = xlRowField
}
```


#### PlotArea 对象

# [PlotArea (对象)​](#plotarea-对象)

代表图表的绘图区。

## [说明​](#说明)

该区域为绘制图表数据的区域。二维图表中的绘图区包含数据标志、网格线、数据标签、趋势线和可选的置于图表区内的图表项。三维图表的绘图区中除包含上述各项外，还在图表中包含背景墙、基底、坐标轴、坐标轴标题和刻度线标签。绘图区被图表区所包围。二维图表的图表区包含坐标轴、图表标题、坐标轴标题和图例。三维图表的图表区包含图表标题和图例。有关设置图表区格式的详细信息，请参阅 ChartArea 对象。

## [示例​](#示例)

javascript
```javascript
/*本示例将 Chart1 中的绘图区前景色设为蓝色。*/
function test() {
    let plotArea = Application.Charts.Item("Chart1").ChartObjects(1).Chart.PlotArea
    plotArea.Format.Fill.ForeColor.RGB = RGB(0, 0, 255)
}
```

javascript
```javascript
/*下例将 Sheet1 上的第一个图表绘图区内部宽度设置为400磅。*/
function test() {
    let plotArea = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.PlotArea
    plotArea.InsideWidth = 400
}
```


#### Point 对象

# [Point (对象)​](#point-对象)

代表图表系列中的单个数据点。

## [说明​](#说明)

Point对象是Points集合的成员。Points集合包含一个系列中所有的数据点。

使用Points(index)（其中 index 是数据点索引号）可返回一个Point对象。系列中的数据点从左至右编号。Points(1)是最左边的数据点，而Points(Points.Count)是最右边的数据点。

## [示例​](#示例)

javascript
```javascript
/*下例为工作表一上嵌入式图表一中系列一中的第三个数据点设置标志样式。指定的系列必须是 2D 线、散点或雷达系列。*/
function test() {
    Worksheets.Item(1).ChartObjects(1).Chart.SeriesCollection(1).Points(3).MarkerStyle = xlDiamond
}
```

javascript
```javascript
/*本示例将 Chart1 中第一个数据系列的第4个数据点设置成有阴影。*/
function test() {
    let series = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(1)
    series.Points(4).Shadow = true
}
```


#### Points 对象

# [Points (对象)​](#points-对象)

图表中指定的系列内所有Point对象的集合。

## [说明​](#说明)

使用Points(index)（其中 index 是数据点索引号）可返回一个Point对象。系列中的数据点从左至右编号。Points(1)是最左边的数据点，而Points(Points.Count)是最右边的数据点。

## [示例​](#示例)

javascript
```javascript
/*下例给工作表一上嵌入式图表一中系列一上的最后一个数据点添加数据标签。*/
function test() {
    let points = Worksheets.Item(1).ChartObjects(1).Chart.SeriesCollection(1).Points()
    points.Item(points.Count).ApplyDataLabels(xlShowValue)
}
```

javascript
```javascript
/*下例为工作表一上嵌入式图表一中系列一中的第三个数据点设置标志样式。指定的系列必须是 2D 线、散点或雷达系列。*/
function test() {
    Worksheets.Item(1).ChartObjects(1).Chart.SeriesCollection(1).Points(3).MarkerStyle = xlDiamond
}
```


#### Protection 对象

# [Protection (对象)​](#protection-对象)

代表工作表可使用的各种保护选项类型。

## [说明​](#说明)

使用Worksheet对象的Protection属性可返回一个Protection对象。

返回一个Protection对象后，就可用该对象的下列属性来设置或返回保护选项。

AllowDeletingColumns
AllowDeletingRows
AllowFiltering
AllowFormattingCells
AllowFormattingColumns
AllowFormattingRows
AllowInsertingColumns
AllowInsertingHyperlinks
AllowInsertingRows
AllowSorting
AllowUsingPivotTables
## [示例​](#示例)

javascript
```javascript
/*下例通过在最上面的行中放三个成员并保留该工作表说明了如何使用 Protection 对象的 AllowInsertingColumns 属性。然后，此示例检查允许插入列的保护设置是否为 false ，并在必要时将其设置为 true。最后，通知用户插入一个列。*/
function test() {
    Range("A1").Formula = "1"
    Range("B1").Formula = "3"
    Range("C1").Formula = "4"
    ActiveSheet.Protect()
    // Check the protection setting of the worksheet and act accordingly.
    if (Application.ActiveSheet.Protection.AllowInsertingColumns == false) {
        Application.ActiveSheet.Protect(null, null, null, null, null, null, null, null, true)
        console.log("Insert a column between 1 and 3")
    } else {
        console.log("Insert a column between 1 and 3")
    }
}
```

javascript
```javascript
/*本示例显示是否允许在受保护的第二张工作表上插入列。*/
function test() {
    console.log(Worksheets.Item(2).Protection.AllowInsertingColumns)
}
```


#### Range 对象

# [Range (对象)​](#range-对象)

代表某一单元格、某一行、某一列、某一选定区域（该区域可包含一个或若干连续单元格区域），或者某一三维区域。

## [说明​](#说明)

示例部分中说明了以下用于返回Range对象的属性和方法：

Range
属性
Cells
属性
Range
和
Cells
Offset
属性
Union
方法
## [示例​](#示例)

使用Range(arg)（其中arg为区域名称）可返回一个代表单个单元格或单元格区域的Range对象。下例将单元格 A1 中的值赋给单元格 A5。

javascript
```javascript
/*本示例将 A1 单元格的值赋值给 A5 单元格*/
function test() {
    Application.Worksheets.Item("Sheet1").Range("A5").Value2 = Application.Worksheets.Item("Sheet1").Range("A1").Value2
}
```

下例通过为区域 A1:H8 中的每个单元格设置公式，用随机数字填充该区域。如果在不带对象识别符（句点左边的对象）的情况下使用Range属性，该属性会返回活动表上的一个区域。如果活动表不是工作表，则该方法失败。在使用没有显式对象识别符的Range属性之前，请先使用Activate方法激活一个工作表。

javascript
```javascript
/*本示例通过为区域 A1:H8 中的每个单元格设置公式，用随机数字填充该区域*/
function test() {
    Application.Worksheets.Item("Sheet1").Activate()
    //Range is on the active sheet
    Application.Range("A1:H8").Formula = "=Rand()"
}
```

下例清除区域名为“Criteria”的区域中的内容。 注释：如果用文本参数指定区域地址，必须以 A1 样式记号指定该地址（不能用 R1C1 样式记号）。

javascript
```javascript
/*本示例清除区域名为“*Criteria*”的区域中的内容*/
function test() {
    Application.Worksheets.Item(1).Range("Criteria").ClearContents()
}
```

使用Cells(row,column)（其中row是行号，column是列标）可返回一个单元格。下例将单元格 A1 赋值为 24。

javascript
```javascript
/*本示例将活动工作表上第一行第一个单元格的值设置为 24*/
function test() {
    Application.Worksheets.Item(1).Cells.Item(1, 1).Value2 = 24
}
```

下例设置单元格 A2 的公式。

javascript
```javascript
/*本示例在第二行第一列单元格设置公式 =Sum(B1:B5)*/
function test() {
    Application.ActiveSheet.Cells.Item(2, 1).Formula = "=Sum(B1:B5)"
}
```

虽然也可用 Range("A1") 返回单元格 A1，但有时用Cells属性更为方便，因为对行或列使用变量。下例在 Sheet1 上创建行号和列标。注意，当工作表激活以后，使用Cells属性时不必明确声明工作表（它将返回活动工作表上的单元格）。 注释：虽然可用 Visual Basic 字符串函数转换 A1 样式引用，但使用 Cells(1, 1) 记号更为简便（而且也是更好的编程习惯）。

javascript
```javascript
/*本示例使用循环在单元格填入对应的值*/
function test() {
    Application.Worksheets.Item("Sheet1").Activate()
    for (let i = 1; i <= 5; i++) {
        Application.Cells.Item(1, i + 1).Value2 = 1990 + i
    }
    for (let j = 1; j <= 4; j++) {
        Application.Cells.Item(j + 1, 1).Value2 = "Q" + j
    }
}
```

使用expression.Cells(row,column)（其中expression是返回Range对象的表达式，row和column是相对于该区域左上角的偏移量）可返回区域中的一部分。下例设置单元格 C5 中的公式。

javascript
```javascript
/*本示例将 C5:C10 单元格区域中第一行第一列的单元格设置公式 =Rand()*/
function test() {
    Application.Worksheets.Item(1).Range("C5:C10").Cells.Item(1, 1).Formula = "=Rand()"
}
```

使用Range(cell1, cell2)（其中cell1和cell2是指定起始和终止单元格的Range对象）可返回一个Range对象。下例设置单元格区域 A1:J10 的边框线条的样式。 注释：请注意每个Cells属性之前的句点。如果前导的With语句应用于Cells属性，那么这些句点就是必需的。本示例中，句点指示单元格处于工作表一上。如果没有句点，Cells属性将返回活动工作表上的单元格。

javascript
```javascript
/*本示例设置单元格 A1:J10 的边框线条样式*/
function test() {
    Application.Worksheets.Item(1).Range(Worksheets.Item(1).Cells.Item(1, 1), Application.Worksheets.Item(1).Cells.Item(10, 10)).Borders.LineStyle = xlThick
}
```

使用Offset(row, column)（其中row和column为行偏移量和列偏移量）可返回相对于另一区域在指定偏移量处的区域。下例选定位于当前选定区域左上角单元格的向下三行且向右一列处的单元格。由于必须选定位于活动工作表上的单元格，因此必须先激活工作表。

javascript
```javascript
/*本示例选择从当前选定区域左上角的单元格下移 3 行和右移 1 列所得的单元格。*/
function test() {
    Application.Worksheets.Item("Sheet1").Activate()
    //Can't select unless the sheet is active
    Application.Selection.Offset(3, 1).Range("A1").Select()
}
```

使用Union(range1, range2, ...) 可返回多块区域，即该区域由两个或多个连续的单元格区域所组成。下例创建由单元格区域 A1:B2 和 C3:D4 组合定义的对象，然后选定该定义区域。

javascript
```javascript
/*本示例选中 A1:B2 和 C3:D4 单元格区域*/
function test() {
    Application.Worksheets.Item("Sheet1").Activate()
    let r1 = Range("A1:B2")
    let r2 = Range("C3:D4")
    let myMultiAreaRange = Union(r1, r2)
    myMultiAreaRange.Select()
}
```

如果您处理包含多个区域的选定内容，Areas属性是很有用的。它将多区域选定内容拆分为单个的Range对象，然后将对象返回为一个集合。您可以对返回的集合使用Count属性，以查找包含多个区域的选定内容，如下例所示。

javascript
```javascript
/*本示例判断如果工作表中选中多个区域，则显示 You cannot carry out this command on multi-area selections*/
function test() {
    let NumberOfSelectedAreas = Application.Selection.Areas.Count
    if (NumberOfSelectedAreas > 1) {
        console.log("You cannot carry out this command " + "on multi-area selections")
    }
}
```


#### Ranges 对象

# [Ranges (对象)​](#ranges-对象)

由Range对象组成的集合。


#### RectangularGradient 对象

# [RectangularGradient (对象)​](#rectangulargradient-对象)

RectangularGradient对象沿特定角度以线性方式在一系列颜色间转换。

## [说明​](#说明)

| 使用 RectangularGradient 对象时，请考虑以下几点： |
| --- |

试图访问不具有现有渐变填充的
Interior
对象的 Gradient 属性会引起运行时错误。访问 Gradient 属性之前请注意
Interior.Pattern
属性。
如果将 Interior.Pattern从渐变类型更改为非渐变类型，Gradient 对象将采用默认值进行填充。

#### Series 对象

# [Series (对象)​](#series-对象)

代表图表上的系列。

## [说明​](#说明)

Series对象是SeriesCollection集合的成员。

使用SeriesCollection(index)（其中index是系列索引号或名称）可返回一个Series对象。下例设置 Sheet1 上嵌入式图表一中第一个系列的内部颜色。

系列索引号指明了系列添加到图表中的顺序。SeriesCollection(1)是第一个添加到图表中的系列，而SeriesCollection(SeriesCollection.Count)是最后一个添加到图表中的系列。

## [示例​](#示例)

javascript
```javascript
/*下面的示例设置工作表 Sheet1 上嵌入的第一个图表中第一个数据系列的内部颜色。*/
function test() {
    Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.SeriesCollection(1).Interior.Color = RGB(255, 0, 0)
}
```

javascript
```javascript
/*此示例显示图表工作表 Chart1 上的第二个数据系列的名称。*/
function test() {
    let series2 = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(2)
    console.log(series2.Name)
}
```


#### SeriesCollection 对象

# [SeriesCollection (对象)​](#seriescollection-对象)

指定的图表或图表组中所有Series对象的集合。

## [说明​](#说明)

使用SeriesCollection方法可返回SeriesCollection集合。

## [示例​](#示例)

javascript
```javascript
/*此示例显示第一张工作表中第一个图表的系列数量。*/
function test() {
    console.log(Application.Worksheets.Item(1).ChartObjects(1).Chart.SeriesCollection().Count)
}
```

javascript
```javascript
/*此示例将单元格 A1:A19 中的数据添加到名为 Chart1 的图表工作表上的新系列。*/
function test() {
    Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection().Add(Worksheets.Item("Sheet1").Range("A1:A19"))
}
```

javascript
```javascript
/*此示例设置工作表 Sheet1 上嵌入的第一个图表中第一个数据系列的内部颜色。*/
function test() {
    Worksheets("Sheet1").ChartObjects(1).Chart.SeriesCollection(1).Interior.Color = RGB(255, 0, 0)
}
```


#### SeriesLines 对象

# [SeriesLines (对象)​](#serieslines-对象)

代表图表组中的系列线。

## [说明​](#说明)

系列线连接每个系列中的数据。只有二维堆积条形图、二堆, 2-D 堆只柱形图、复合饼图或复合条饼图等图表可以有系列线。此对象不是集合。没有代表单个系列线的对象；您或者打开图表组中所有数据点的系列线，或者将其全部关闭。

如果HasSeriesLines属性是False，SeriesLines对象的绝大部分属性都会被禁用。

## [示例​](#示例)

javascript
```javascript
/*本示例将系列线添加到第一张工作表第一个嵌入图表中的第一个图表组中，并将颜色设置为蓝色，该图表必须是二维堆积条形图或柱形图。*/
function test() {
    let chartgroup = Worksheets(1).ChartObjects(1).Chart.ChartGroups(1)
    chartgroup.HasSeriesLines = true
    chartgroup.SeriesLines.Border.Color = RGB(0, 0, 255)
}
```

javascript
```javascript
/*本示例将图表工作表 Chart1 上第一个图表组的系列线边框的粗细设置为粗。*/
function test() {
    let chartgroup = Application.Charts.Item("Chart1").ChartObjects(1).Chart.ChartGroups(1)
    chartgroup.SeriesLines.Border.Weight = xlThick
}
```


#### ShadowFormat 对象

# [ShadowFormat (对象)​](#shadowformat-对象)

代表形状的阴影格式。

## [说明​](#说明)

使用Shadow属性可返回一个ShadowFormat对象。

## [示例​](#示例)

javascript
```javascript
/*下例给第一张工作表添加带阴影的矩形。半透明蓝色阴影在矩形右侧偏移 5 磅，在矩形上方偏移 3 磅。*/
function test() {
    let shadow = Worksheets.Item(1).Shapes.AddShape(msoShapeRectangle, 50, 50, 100, 200).Shadow
    shadow.ForeColor.RGB = RGB(0, 0, 128)
    shadow.Type = msoShadow17
    shadow.OffsetX = 5
    shadow.OffsetY = -3
    shadow.Transparency = 0.5
    shadow.Visible = true
}
```

javascript
```javascript
/*本示例显示活动工作表中第二个形状阴影的大小。*/
function test() {
    let shadow = ActiveSheet.Shapes.Item(2).Shadow
    console.log(shadow.Size)
}
```


#### Shape 对象

# [Shape (对象)​](#shape-对象)

代表绘图层中的对象，例如自选图形、任意多边形、OLE 对象或图片。

## [说明​](#说明)

Shape对象是Shapes集合的成员。Shapes集合包含某个工作簿中的所有形状。

**注释：**有三个代表形状的对象：Shapes集合，它代表工作簿中所有的形状；ShapeRange集合，它代表工作簿中形状的指定子集（例如，ShapeRange对象可以代表工作簿中的形状一和形状四，或者，可以代表工作簿中所有选定的形状）；Shape对象，它代表文档中的某一个形状。如果您需要同时处理几个形状，或处理选定区域中的多个形状，请使用ShapeRange集合。

以下各节说明了如何使用 Shape 对象：

返回与连接符的端点相连的形状。
返回新建的任意多边形。
返回组中的单个形状。
返回新组成的形状组。
返回现有的形状。
返回选定区域中的形状。
1.返回与连接符的端点相连的形状

要返回一个代表连接符所连接形状之一的Shape对象，请使用BeginConnectedShape或EndConnectedShape属性。

2. 返回新建的任意多边形

使用BuildFreeform和AddNodes方法可定义一个新任意多边形的几何特性，使用ConvertToShape方法可创建任意多边形并返回代表它的Shape对象。

3.返回组中的单个形状

使用GroupItems(index)（其中index是形状的名称或组中的索引号）可返回一个代表一组形状中某一形状的Shape对象。

4.返回新组成的形状组

使用Group或Regroup方法可将一系列形状分成一组并返回一个Shape对象，该对象代表新形成的组。在形成了一个组之后，您可以按您处理其他任何形状的方法来处理该组

5. 返回现有的形状

使用Shapes(index)（其中index是形状名称或索引号）可返回代表某个形状的Shape对象。

6. 返回选定区域中的形状

使用Selection.ShapeRange(index)（其中index是形状名称或索引号）可返回一个代表选定区域中的形状的Shape对象。

## [示例​](#示例)

javascript
```javascript
/*本示例将第一张工作表中形状一和名为“Rectangle 1”的形状进行水平翻转。*/
function test() {
    let shapes = Application.Worksheets.Item(1).Shapes
    shapes.Item(1).Flip(msoFlipHorizontal)
    shapes.Item("Rectangle 1").Flip(msoFlipHorizontal)
}
```

每个添加到 Shapes 集合的形状将被分配一个默认名称。 若要赋予该形状一个更有意义的名称，请使用 Name 属性。

javascript
```javascript
/*本示例向活动工作表添加一个正方形，为其命名为“Red Square”，然后设置其前景色和线条样式。*/
function test() {
    let shapes = ActiveSheet.Shapes
    let shape = shapes.AddShape(msoShapeRectangle, 144, 144, 72, 72)
    shape.Name = "Red Square"
    shape.Fill.ForeColor.RGB = RGB(255, 0, 0)
    shape.Line.DashStyle = msoLineDashDot
}
```


#### ShapeNodes 对象

# [ShapeNodes (对象)​](#shapenodes-对象)

指定的任意多边形中所有ShapeNode对象的集合。

## [说明​](#说明)

每一个ShapeNode对象都代表任意多边形中线段间的结点或任意多边形曲线段的控点。您可以手动创建或通过使用BuildFreeform和ConvertToShape方法来创建任意多边形。

使用Nodes属性可返回ShapeNodes集合。

## [示例​](#示例)

javascript
```javascript
/*本示例删除第一张工作表上形状三中的结点四。*/
function test() {
    let shapes = Worksheets.Item(1).Shapes
    shapes.Item(3).Nodes.Delete(4)
}
```

使用Insert方法可创建一个新结点并将它添加到ShapeNodes集合。

javascript
```javascript
/*本示例在第一张工作表上的形状三中的结点四之后添加一带有曲线段的平滑结点。*/
function test() {
    let nodes = Application.Worksheets.Item(1).Shapes.Item(3).Nodes
    nodes.Insert(4, msoSegmentCurve, msoEditingSmooth, 210, 100)
}
```

使用Nodes(index)（其中index是结点索引号）可返回一个ShapeNode对象。

javascript
```javascript
/*本示例将第一张工作表上形状三中的结点一设置为平滑顶点。*/
function test() {
    let shape = Application.Worksheets.Item(1).Shapes.Item(3)
    shape.Nodes.SetEditingType(1, msoEditingSmooth)
}
```


#### ShapeRange 对象

# [ShapeRange (对象)​](#shaperange-对象)

代表形状区域，它是文档中的一组形状。

## [说明​](#说明)

形状区域可以只包含文档中的一个形状，或者也可包含所有形状。您可以在形状区域中包含所需的任意形状（在文档中的所有形状中选取，或从选定内容中的所有形状中选取）。例如，您可以构造一个ShapeRange集合，它包含文档中前三个形状、文档中所有选定的形状，或文档中所有的任意多边形的。

### [1. 返回指定名称或索引号的一组形状​](#_1-返回指定名称或索引号的一组形状)

使用Shapes.Range(index)（其中index是形状的名称或索引号，或由形状的名称或索引号组成的数组）可返回代表文档中的一组形状的ShapeRange集合。您可以使用Array函数来构造名称或索引号的数组。

## [示例​](#示例)

javascript
```javascript
/*本示例设置第一张工作表上的形状一和三的填充图案。*/
function test() {
    let myDocument = Application.Worksheets.Item(1)
    myDocument.Shapes.Range([1, 3]).Fill.Patterned(msoPatternHorizontalBrick)
}
```

虽然可以使用Range属性来返回任意数量的形状或幻灯片，但如果您只想返回一个集合成员，则使用Item方法会更简单。例如，Shapes(1)比Shapes.Range(1)简单。

javascript
```javascript
/*本示例设置第一张工作表上名为“Oval 4”和“Rectangle 5”的形状的填充图案。*/
function test() {
    let myDocument = Application.Worksheets.Item(1)
    let myRange = myDocument.Shapes.Range(["Oval 4", "Rectangle 5"])
    myRange.Fill.Patterned(msoPatternHorizontalBrick)
}
```

### [2. 返回文档中全部或部分选定的形状​](#_2-返回文档中全部或部分选定的形状)

使用Selection对象的ShapeRange属性可返回选定对象中的所有形状。

javascript
```javascript
/*本示例为活动工作表上选定内容中的所有形状设置前景填充色。*/
function test() {
    let shapes = ActiveSheet.Shapes
    shapes.SelectAll()
    Selection.ShapeRange.Fill.ForeColor.RGB = RGB(255, 0, 255)
}
```

使用Selection.ShapeRange(index)（其中index是形状的名称或索引号）返回某一选定的形状。

javascript
```javascript
/*本示例为活动工作表上选定内容中的第二个形状设置前景填充色。*/
function test() {
    let shapes = ActiveSheet.Shapes
    shapes.SelectAll()
    Selection.ShapeRange.Item(2).Fill.ForeColor.RGB = RGB(255, 0, 255)
}
```


#### Shapes 对象

# [Shapes (对象)​](#shapes-对象)

指定的工作表上的所有Shape对象的集合。

## [说明​](#说明)

每个Shape对象都代表绘图层中的一个对象，如自选图形、任意多边形、OLE 对象或图片。

**注释：**如果您想处理文档中的一部分形状（例如，只针对文档中的自选图形或只对选定的形状进行一些操作），则必须构造一个ShapeRange集合，其中包含您要处理的形状。

使用Shapes属性可返回Shapes集合。下例选定 myDocument 上的所有形状。

**注释：**如果您要同时对工作表上的所有形状进行操作（例如删除或设置属性），请选定所有形状，然后对选定区域使用ShapeRange属性，以创建一个ShapeRange对象，该对象包含工作表上的所有形状，然后对ShapeRange对象应用相应的属性或方法。

## [示例​](#示例)

javascript
```javascript
/*本示例选中第一张工作表中所有形状。*/
function test() {
    let shapes = Application.Worksheets.Item(1).Shapes
    shapes.SelectAll()
}
```

使用Shapes(index)（其中index是形状的名称或索引号）可返回一个 Shape 对象。下例设置myDocument上形状一的预设阴影的填充。

javascript
```javascript
/*本示例设置第一张工作表上形状一的预设阴影的填充。*/
function test() {
    let shapes = Application.Worksheets.Item(1).Shapes
    shapes.Item(1).Fill.PresetGradient(msoGradientHorizontal, 1, msoGradientBrass)
}
```

使用Shapes.Range(index)（其中index是形状的名称或索引号，或是它们的一个数组）可返回一个 ShapeRange 集合，该集合代表 Shapes 集合的一个子集。下例设置myDocument上形状一和三的填充图案。

javascript
```javascript
/*本示例设置第一张工作表上形状一和三的填充图案。*/
function test() {
    let shapes = Application.Worksheets.Item(1).Shapes
    shapes.Range([1, 3]).Fill.Patterned(msoPatternHorizontalBrick)
}
```

javascript
```javascript
/*本示例将活动工作表上名为“Check Box 1”的复选框设置为选中状态。*/
function test() {
    Application.ActiveSheet.Shapes("Check Box 1").ControlFormat.Value = true
}
```


#### SheetViews 对象

# [SheetViews (对象)​](#sheetviews-对象)

指定的或活动工作簿窗口中所有工作表视图的集合。

## [说明​](#说明)

## [示例​](#示例)

javascript
```javascript
/*本示例显示活动窗口的工作表视图数量*/
function test() {
    let count = Application.ActiveWindow.SheetViews.Count
    console.log(`活动窗口的工作表视图数量是 ${count}`)
}
```

javascript
```javascript
/*本示例更改第一个工作表视图所在的工作表名称，并显示出来。*/
function test() {
    let sheetview = ActiveWindow.SheetViews.Item(1)
    sheetview.Sheet.Name = "成绩表"
    console.log(sheetview.Sheet.Name)
}
```


#### Sheets 对象

# [Sheets (对象)​](#sheets-对象)

指定的或活动工作簿中所有工作表的集合。

## [说明​](#说明)

Sheets集合可以包含Chart或Worksheet对象。

如果希望返回所有类型的工作表，Sheets集合就非常有用。如果仅需使用某一类型的工作表，请参阅该工作表类型的对象主题。

使用Sheets属性可返回Sheets集合。

## [示例​](#示例)

javascript
```javascript
/*此示例打印当前活动工作簿上的所有工作表。*/
function test() {
    Application.Sheets.PrintOut()
}
```

使用Add方法可创建一个新的工作表并将它添加到集合。

javascript
```javascript
/*此示例给活动工作簿添加两个图表工作表，将它们放在工作簿中的工作表二之后。*/
function test() {
    Application.Sheets.Add(null, Application.Sheets.Item(2), 2, xlChart)
}
```

使用Sheets(index)（其中index是工作表名称或索引号）可返回一个Chart或Worksheet对象。

javascript
```javascript
/*此示例激活名为 Sheet1 的工作表。*/
function test() {
    Application.Sheets.Item("Sheet1").Activate()
}
```

使用Sheets(array) 可指定多个工作表。下例将名为“Sheet4”和“Sheet5”的工作表移到工作簿的开头。

javascript
```javascript
/*此示例将名为 Sheet4 和 Sheet5 的工作表移动到工作簿的开头。*/
function test() {
    Application.Sheets.Item(["Sheet4", "Sheet5"]).Move(Application.Sheets.Item(1))
}
```


#### Slicer 对象

# [Slicer (对象)​](#slicer-对象)

代表工作簿中的一个切片器。

## [说明​](#说明)

每个Slicer对象都代表工作簿中的一个切片器。切片器用于筛选数据透视表或 OLAP 数据源中的数据。

使用Add方法可将Slicer对象添加到Slicers集合。若要访问表示切片器中当前选中的按钮的SlicerItem对象，请使用Slicer对象的ActiveItem属性。

## [示例​](#示例)

javascript
```javascript
/*本示例将活动工作簿上第一个切片器缓存中第一个切片器的标题更改为“My Slicer”。*/
function test() {
    ActiveWorkbook.SlicerCaches(1).Slicers.Item(1).Caption = "My Slicer"
}
```

javascript
```javascript
/*本示例将活动工作簿上第一个切片器缓存中第一个切片器的宽度设置为 200 磅。*/
function test() {
    ActiveWorkbook.SlicerCaches(1).Slicers.Item(1).Width = 200
}
```


#### SlicerCache 对象

# [SlicerCache (对象)​](#slicercache-对象)

表示切片器的当前筛选状态，以及有关切片器连接到哪个PivotCache或WorkbookConnection的信息。

## [说明​](#说明)

使用Workbook对象的SlicerCaches属性可访问工作簿中SlicerCache对象的集合。

每个切片器都有一个基SlicerCache对象，该对象表示在切片器中显示的项目，以及与对应的项目标题一起显示的平铺的当前用户界面状态。用户在 ET 中看到的每个切片器控件由带有关联的SlicerCache对象的Slicer对象表示。

## [示例​](#示例)

javascript
```javascript
/*本示例在第一张工作簿上添加新的切片器缓存，并在工作表 Sheet2 中创建切片器。*/
function test() {
    let pvtTable = Worksheets.Item("Sheet1").Range("I1").PivotTable
    let sliCache = Workbooks.Item(1).SlicerCaches.Add2(pvtTable, "name", "Slicer_name")
    sliCache.Slicers.Add(Worksheets.Item("Sheet2"), null, "Name")
}
```

javascript
```javascript
/*本示例显示活动工作簿上所有切片器缓存的名称。*/
function test() {
    for (let i = 1; i <= ActiveWorkbook.SlicerCaches.Count; i++) {
        console.log(ActiveWorkbook.SlicerCaches(i).Name)
    }
}
```


#### SlicerCaches 对象

# [SlicerCaches (对象)​](#slicercaches-对象)

表示与指定工作簿关联的切片器缓存的集合。

## [说明​](#说明)

使用SlicerCaches集合的Item属性可返回与指定的Workbook对象关联的SlicerCache对象。可以使用Index属性的值，或者指定对象的Name属性来检索SlicerCache对象。

## [示例​](#示例)

javascript
```javascript
/*本示例显示活动工作簿上名为“切片器_id”的切片器缓存所连接到的数据源的名称。*/
function test() {
    console.log(ActiveWorkbook.SlicerCaches.Item("切片器_id").SourceName)
}
```

javascript
```javascript
/*本示例在活动工作簿上添加新的切片器缓存，并显示该切片器缓存的名称。*/
function test() {
    let sliCache =  ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.Range("I1").PivotTable, "date", "Slicer_date", xlSlicer)
    console.log(sliCache.Name)
}
```


#### SlicerItem 对象

# [SlicerItem (对象)​](#sliceritem-对象)

表示切片器中的一个项目。

## [说明​](#说明)

若要访问表示切片器中当前选中的按钮的SlicerItem对象，请使用Slicer对象的ActiveItem属性。如果切片器筛选数据透视表，那么若要访问表示该切片器中所有项目的SlicerItems集合，请使用与该Slicer对象关联的SlicerCache对象的SlicerItems属性。如果切片器筛选 OLAP 层次结构级别，那么若要访问表示该切片器中项目的SlicerItems集合，请使用表示该层次结构级别的SlicerCacheLevel对象的SlicerItems属性。

## [示例​](#示例)

javascript
```javascript
/*本示例选中第一张工作簿上名为“切片器_id”的切片器缓存中所有切片器项。*/
function test() {
    let sliItems = Workbooks.Item(1).SlicerCaches("切片器_id").SlicerItems
    for (let i = 1; i <= sliItems.Count; i++) {
        sliItems.Item(i).Selected = true
    }
}
```

javascript
```javascript
/*本示例显示第一张工作簿上名为“切片器_name”的切片器缓存中第二个切片器项的源名称。*/
function test() {
    let sliItem = Workbooks.Item(1).SlicerCaches("切片器_name").SlicerItems(2)
    console.log(sliItem.SourceName)
}
```


#### SlicerItems 对象

# [SlicerItems (对象)​](#sliceritems-对象)

表示包含在SlicerCache或SlicerCacheLevel对象中的SlicerItem对象的集合。

## [说明​](#说明)

如果切片器基于工作簿中的数据或非 OLAP 外部数据，那么若要访问表示该切片器中项目的SlicerItems集合，请使用与该切片器关联的SlicerCache对象的SlicerItems属性。

如果切片器基于 OLAP 数据连接，那么若要访问表示该切片器中项目的SlicerItems集合，请使用代表层次结构级别的SlicerCacheLevel对象的SlicerItems属性。

## [示例​](#示例)

javascript
```javascript
/*本示例显示活动工作表上数据透视表第一个切片器的缓存中切片器项的数量。*/
function test() {
    let sliItems = ActiveSheet.Range("I1").PivotTable.Slicers.Item(1).SlicerCache.SlicerItems
    console.log(`切片器项的数量：${sliItems.Count}`)
}
```

javascript
```javascript
/*本示例将活动工作簿上名为“切片器_score”的切片器缓存中切片器项的名称逐个赋值到单元格中。*/
function test() {
    let sliItems = ActiveWorkbook.SlicerCaches("切片器_score").SlicerItems
    for (let i = 1; i <= sliItems.Count; i++) {
        Cells.Item(11, i).Value2 = sliItems.Item(i).Name
    }
}
```


#### SlicerPivotTables 对象

# [SlicerPivotTables (对象)​](#slicerpivottables-对象)

表示与指定的SlicerCache对象关联的数据透视表集合相关的信息。

## [说明​](#说明)

SlicerPivotTables集合包含有关切片器缓存当前正在筛选的数据透视表的信息。它提供的属性可用来确定与该切片器关联的数据透视表的数量，可用来检索表示正被筛选的数据透视表的PivotTable对象。它还提供了在SlicerPivotTables集合中添加和删除数据透视表的方法。如果与指定的SlicerCache关联的切片器未连接到任何数据透视表，那么SlicerPivotTables集合将为空。

使用SlicerCache对象的PivotTables属性可返回与SlicerCache关联的SlicerPivotTables集合，而 SlicerCache 又可能与一个或多个切片器关联。

## [示例​](#示例)

javascript
```javascript
/*本示例显示活动工作表上数据透视表的第一个切片器关联的缓存中第二个数据透视表的值。*/
function test() {
    let sliCache = ActiveSheet.Range("I1").PivotTable.Slicers.Item(1).SlicerCache
    console.log(sliCache.PivotTables.Item(2).Value)
}
```

javascript
```javascript
/*本示例显示活动工作簿上名为“切片器_name”的切片器缓存关联的数据透视表的数量。*/
function test() {
    let pvtTables = ActiveWorkbook.SlicerCaches("切片器_name").PivotTables
    console.log(`数据透视表的数量：${pvtTables.Count}`)
}
```


#### Slicers 对象

# [Slicers (对象)​](#slicers-对象)

Slicer对象的集合。

## [说明​](#说明)

每个Slicer对象都代表工作簿中的一个切片器。切片器的作用是筛选数据。

使用Slicers属性可返回Slicers集合。

## [示例​](#示例)

javascript
```javascript
/*本示例显示活动工作簿中第一个切片器缓存中的切片器数量。*/
function test() {
    console.log(ActiveWorkbook.SlicerCaches(1).Slicers.Count)
}
```

使用Slicers(index)（其中index是切片器的索引号或名称）可返回切片器集合中的单个Slicer对象。

javascript
```javascript
/*本示例将第一个切片器缓存中第一个切片器的标题更改为“My Slicer”。*/
function test() {
    ActiveWorkbook.SlicerCaches(1).Slicers.Item(1).Caption = "My Slicer"
}
```


#### Sort 对象

# [Sort (对象)​](#sort-对象)

代表数据区域的排序方式。

## [示例​](#示例)

javascript
```javascript
/*此示例为活动工作表区域A1:A11赋值，并对这些数据排序。*/
function test() {
    //Building data to sort on the active sheet.
    Range("A1").Value2 = "Name"
    Range("A2").Value2 = "Bill"
    Range("A3").Value2 = "Rod"
    Range("A4").Value2 = "John"
    Range("A5").Value2 = "Paddy"
    Range("A6").Value2 = "Kelly"
    Range("A7").Value2 = "William"
    Range("A8").Value2 = "Janet"
    Range("A9").Value2 = "Florence"
    Range("A10").Value2 = "Albert"
    Range("A11").Value2 = "Mary"
    console.log("The list is out of order.  Hit Ok to continue...")

    //Selecting a cell within the range.
    Range("A2").Select()

    //Applying sort.
    let sort = Application.ActiveWorkbook.Worksheets.Item(ActiveSheet.Name).Sort
    sort.SortFields.Clear()
    sort.SortFields.Add(Range("A2:A11"), xlSortOnValues, xlAscending, xlSortNormal)
    sort.SetRange(Range("A1:A11"))
    sort.Header = xlYes
    sort.MatchCase = false
    sort.Orientation = xlTopToBottom
    sort.SortMethod = xlPinYin
    sort.Apply()
    console.log("Sort complete.")
}
```

javascript
```javascript
/*此示例对第一张工作表区域A1:C1按字符的汉语拼音顺序排序。*/
function test() {
    let sort = Application.Sheets.Item(1).Sort
    sort.SortFields.Clear()
    sort.SortFields.Add(Range("A1:C1"))
    sort.SetRange(Range("A1:C1"))
    sort.Header = xlYes
    sort.MatchCase = false
    sort.Orientation = xlSortRows
    sort.SortMethod = xlPinYin
    sort.Apply()
}
```


#### SortField 对象

# [SortField (对象)​](#sortfield-对象)

SortField对象包含Worksheet、Lists和AutoFilter对象的所有排序信息。

## [说明​](#说明)

开发人员可以使用BeforeSort事件重写 ET 的默认行为，将自己的排序算法写入应用程序中。

## [示例​](#示例)

javascript
```javascript
/*本示例遍历活动工作表的排序字段，将偶数位的排序字段设置为升序排序。*/
function test() {
    let sortfields = Application.ActiveSheet.Sort.SortFields
    for (let i = 1; i <= sortfields.Count; i++) {
        if (i % 2 == 0) {
            sortfields.Item(i).Order = xlDescending
        }
    }
}
```

javascript
```javascript
/*本示例在活动工作表新建排序字段，并显示该排序字段是否按字体颜色对数据进行升序排序。*/
function test() {
    let sortfield = Application.ActiveSheet.Sort.SortFields.Add(Range("A1:A3"), xlSortOnFontColor, xlAscending)
    if (sortfield.SortOn == xlSortOnFontColor && sortfield.Order == xlAscending) {
        console.log("排序字段按字体颜色对数据进行升序排序")
    } else {
        console.log("排序字段未按字体颜色和升序进行排序")
    }
}
```


#### SortFields 对象

# [SortFields (对象)​](#sortfields-对象)

SortFields集合是SortField对象的集合。开发人员可以使用该集合存储工作簿、列表和自动筛选的排序状态。

## [示例​](#示例)

javascript
```javascript
/*本示例为活动工作表创建两个新的排序字段。*/
function test() {
    Application.ActiveSheet.Sort.SortFields.Add(Range("A1"), null, xlDescending)
    Application.ActiveSheet.Sort.SortFields.Add(Range("B1"), null, xlDescending)
}
```

javascript
```javascript
/*本示例清除第一张工作表上所有的 SortField 对象，并创建新的排序字段。*/
function test() {
    Application.Sheets.Item(1).Sort.SortFields.Clear()
    Application.Sheets.Item(1).Sort.SortFields.Add(Range("A1:A11"), xlSortOnValues, xlAscending)
}
```


#### SparkAxes 对象

# [SparkAxes (对象)​](#sparkaxes-对象)

表示一组迷你图的水平轴和垂直轴的设置。

## [说明​](#说明)

使用SparklineGroup对象的Axes属性可返回该组迷你图的SparkAxes对象。

## [示例​](#示例)

javascript
```javascript
/*本示例显示活动工作表上单元格 H1 中第一个迷你图组是否按右至左顺序在水平轴上绘制点。*/
function test() {
    let shAxis = ActiveSheet.Range("H1").SparklineGroups(1).Axes.Horizontal
    console.log(shAxis.RightToLeftPlotOrder)
}
```

javascript
```javascript
/*本示例设置活动工作表上单元格 H1 中第一个迷你图组的水平轴的颜色。*/
function test() {
    let shAxis = ActiveSheet.Range("H1").SparklineGroups(1).Axes.Horizontal
    shAxis.Axis.Color.ColorIndex = 3
}
```


#### SparkColor 对象

# [SparkColor (对象)​](#sparkcolor-对象)

表示迷你图中的点的标记色。

## [说明​](#说明)

SparkColor对象对应于功能区“迷你图工具设计”选项卡的“样式”部分中“标记颜色”下拉列表中提供的项的设置。使用SparkPoints对象的对应属性可设置这些项的颜色。

## [示例​](#示例)

javascript
```javascript
/*本示例将活动工作表上区域 A1:A4 中第一个迷你图组的迷你图的最高数据点设置为可见。*/
function test() {
    let sparkColor = ActiveSheet.Range("A1:A4").SparklineGroups.Item(1).Points.Highpoint
    sparkColor.Visible = true
}
```

javascript
```javascript
/*本示例将活动工作表上区域 A1:A4 中第一个迷你图组的迷你图的最后一个数据点设置为可见，并设置其颜色。*/
function test() {
    let sparkColor = ActiveSheet.Range("A1:A4").SparklineGroups.Item(1).Points.Lastpoint
    sparkColor.Visible = true
    sparkColor.Color.Color = RGB(255, 0, 255)
}
```


#### SparkHorizontalAxis 对象

# [SparkHorizontalAxis (对象)​](#sparkhorizontalaxis-对象)

表示一组迷你图的水平轴的设置。

## [说明​](#说明)

使用SparkAxes对象的Horizontal属性可返回迷你图组的SparkHorizontalAxis对象。水平轴只有在迷你图数据在垂直轴上同时有负值和正值的情况下才会显示。

## [示例​](#示例)

javascript
```javascript
/*本示例将活动工作表上单元格 H1 中第一个迷你图组的水平轴设置为可见，并设置该水平轴的颜色。*/
function test() {
    let shAxis = ActiveSheet.Range("H1").SparklineGroups(1).Axes.Horizontal
    shAxis.Axis.Visible = true
    shAxis.Axis.Color.ColorIndex = 3
}
```

javascript
```javascript
/*本示例判断活动工作表上单元格 H1 中第一个迷你图组的水平轴绘制点的顺序，并通知用户。*/
function test() {
    let shAxis = ActiveSheet.Range("H1").SparklineGroups(1).Axes.Horizontal
    if (shAxis.RightToLeftPlotOrder) {
        console.log("按右至左顺序在水平轴上绘制点")
    } else {
        console.log("按左至右顺序在水平轴上绘制点")
    }
}
```


#### SparkPoints 对象

# [SparkPoints (对象)​](#sparkpoints-对象)

表示迷你图上数据点的标记设置。

## [说明​](#说明)

使用SparkPoints对象可设置迷你图上数据点标记的颜色和可见性。使用SparklineGroup对象的Points属性可返回SparkPoints对象。SparkPoints对象的属性对应于“显示”部分中“高点”、“低点”、“负点”、“首点”、“尾点”和“标记”复选框的设置，以及功能区的“迷你图工具设计”选项卡上“样式”部分中的“标记颜色”下拉列表内各项的设置。

## [示例​](#示例)

javascript
```javascript
/*本示例将活动工作表上区域 A1:A4 中第一个迷你图组的迷你图的最高和最低数据点设置为可见，并设置最高数据点的颜色。*/
function test() {
    let points = ActiveSheet.Range("A1:A4").SparklineGroups.Item(1).Points
    points.Highpoint.Visible = true
    points.Lowpoint.Visible = true
    points.Highpoint.Color.ColorIndex = 4
}
```

javascript
```javascript
/*本示例将活动工作表上区域 A1:A4 中第一个迷你图组的迷你图的最后一个数据点设置为可见，并设置其颜色。*/
function test() {
    let sparkColor = ActiveSheet.Range("A1:A4").SparklineGroups.Item(1).Points.Lastpoint
    sparkColor.Visible = true
    sparkColor.Color.Color = RGB(255, 0, 0)
}
```


#### Sparkline 对象

# [Sparkline (对象)​](#sparkline-对象)

表示单个迷你图。

## [说明​](#说明)

使用ModifyLocation方法可更改单个迷你图的位置，使用ModifySourceData方法可更改源数据区域。若要一次操作一组迷你图，请使用SparklineGroup对象的成员。

## [示例​](#示例)

javascript
```javascript
/*本示例修改单元格 I1 中第一个迷你图组的第一个迷你图的位置，并显示修改后迷你图的位置的地址。*/
function test() {
    let sparkline = Range("I1").SparklineGroups(1).Item(1)
    sparkline.ModifyLocation(Range("K1"))
    console.log(sparkline.Location.Address())
}
```

javascript
```javascript
/*本示例修改单元格 I1 中第一个迷你图组的第一个迷你图的源数据，并显示修改后的源数据的区域。*/
function test() {
    let sparkline = Range("I1").SparklineGroups(1).Item(1)
    sparkline.ModifySourceData("Sheet1!A1:C1")
    console.log(`修改后的源数据区域：${sparkline.SourceData}`)
}
```


#### SparklineGroup 对象

# [SparklineGroup (对象)​](#sparklinegroup-对象)

代表一组迷你图。

## [说明​](#说明)

SparklineGroup对象可包含多个迷你图以及该迷你图组的属性设置，例如颜色和轴设置。每个迷你图都由一个Sparkline对象表示。

使用Modify方法可在迷你图组中添加或删除迷你图。使用ModifyLocation方法可更改迷你图的位置，使用ModifySourceData方法则可以更改源数据所在的区域。

## [示例​](#示例)

javascript
```javascript
/*本示例会在 A1:A4 处创建一组绑定到 Sheet1!B1:E4 区域中的源数据的列迷你图，并更改系列颜色以便用红色显示各个列。*/
function test() {
    let sparklineGroup = Range("A1:A4").SparklineGroups.Add(xlSparkColumn, "Sheet1!B1:E4")
    sparklineGroup.SeriesColor.Color = RGB(255, 0, 0)
}
```

javascript
```javascript
/*本示例显示活动工作表上单元格 H1 中第一个迷你图组中迷你图的高点是否可见。*/
function test() {
    let sparklineGroup = ActiveSheet.Range("H1").SparklineGroups.Item(1)
    console.log(sparklineGroup.Points.Highpoint.Visible)
}
```


#### SparklineGroups 对象

# [SparklineGroups (对象)​](#sparklinegroups-对象)

代表迷你图组的集合。

## [说明​](#说明)

SparklineGroups对象可包含多个SparklineGroup对象。

使用Range对象的SparklineGroups属性可从现有SparklineGroups集合的父区域中返回该集合。

使用Add方法可创建一组新迷你图。

使用Group方法可创建一组现有迷你图。

## [示例​](#示例)

javascript
```javascript
/*此示例选择区域 A1:A4 并组合该区域中的迷你图，然后将该区域中第一个迷你图组的正数据点设置为可见，并将其颜色设置为红色。*/
function test() {
    Range("A1:A4").Select()
    Selection.SparklineGroups.Group(Range("A1"))
    Selection.SparklineGroups.Item(1).Points.Markers.Visible = true
    Selection.SparklineGroups.Item(1).Points.Markers.Color.Color = RGB(255, 0, 0)
}
```

javascript
```javascript
/*此示例取消组合活动工作表上区域 A1:A4 上迷你图组，并删除 A1 单元格中迷你图组。*/
function test() {
    let sparklineGroups = ActiveSheet.Range("A1:A4").SparklineGroups
    sparklineGroups.Ungroup()
    Range("A1").SparklineGroups.ClearGroups()
}
```


#### SpellingOptions 对象

# [SpellingOptions (对象)​](#spellingoptions-对象)

代表工作表的各种拼写检查选项。

## [说明​](#说明)

使用Application对象的SpellingOptions属性可返回一个SpellingOptions对象。

一旦返回了SpellingOptions对象，您就可以使用下列属性来设置或返回各种拼写检查选项。

ArabicModes
DictLang
GermanPostReform
HebrewModes
IgnoreCaps
IgnoreFileNames
IgnoreMixedDigits
KoreanCombineAux
KoreanProcessCompound
KoreanUseAutoChangeList
SuggestMainOnly
UserDict
## [示例​](#示例)

javascript
```javascript
/*下例使用 IgnoreCaps 属性来禁用对全部是大写字母的单词的拼写检查。在本示例中，拼写检查程序发现的是“Testt”，而不是“TESTT”。*/
function test() {
    // 将同一个单词的拼写错误版本全部大写以及混合大小写。
    Range("A1").Formula = "Testt"
    Range("A2").Formula = "TESTT"

    Application.SpellingOptions.SuggestMainOnly = true
    Application.SpellingOptions.IgnoreCaps = true

    // Run a spell check.
    Cells.CheckSpelling()
}
```

javascript
```javascript
/*本示例显示 ET 在使用拼写检查时是否忽略大写单词。*/
function test() {
    console.log(Application.SpellingOptions.IgnoreCaps)
}
```


#### Style 对象

# [Style (对象)​](#style-对象)

代表区域的样式说明。

## [说明​](#说明)

Style对象包含样式的所有属性（字体、数字格式、对齐方式，等等）。有几种内置样式，包括“常规”、“货币”和“百分比”。同时对多个单元格修改单元格格式属性时，使用Style对象是快捷高效的方法。

对于Workbook对象，Style对象是Styles集合的成员。Styles集合包含该工作簿的所有已定义样式。

通过更改应用于单元格的样式的属性可更改单元格的外观。但要记住，更改样式的属性将影响所有以该样式格式化了的单元格。

样式按照名称的字母顺序排序。样式编号表明指定样式在样式名排序列表中的位置。Styles(1)是排序列表中的第一个样式，而Styles(Styles.Count)是最后一个。

有关创建和修改样式的详细信息，请参阅Styles对象。

使用Style属性可返回一个用于Range对象的Style对象。

## [示例​](#示例)

javascript
```javascript
/*本示例将“百分比”样式应用于 Sheet1 中的单元格区域 A1:A10。*/
function test() {
    Application.Worksheets.Item("Sheet1").Range("A1:A10").Style = "Percent"
}
```

使用Styles(index)（其中index是样式索引号或名称）可从工作簿的Style集合中返回一个Styles对象。

javascript
```javascript
/*本示例通过设置样式的 Bold 属性来更改活动工作簿的 Normal 样式。*/
function test() {
    Application.ActiveWorkbook.Styles.Item("Normal").Font.Bold = true
}
```


#### Styles 对象

# [Styles (对象)​](#styles-对象)

指定工作簿或活动工作簿中所有Style对象的集合。

## [说明​](#说明)

每一个Style对象都代表对某区域的样式描述。Style对象包含样式的所有属性（字体、数字格式、对齐方式，等等）。有几种内置的样式，包括“常规”、“货币”和“百分比”。

使用Styles属性可返回Styles集合。

## [示例​](#示例)

javascript
```javascript
/*此示例在第一个工作表上创建活动工作簿中样式名的列表。*/
function test() {
    for (let i = 1; i <= Application.ActiveWorkbook.Styles.Count; i++) {
        Application.Worksheets.Item(1).Cells.Item(i, 1).Value2 = ActiveWorkbook.Styles.Item(i).Name
    }
}
```

使用Add方法可创建一个新的样式并将它添加到集合。

javascript
```javascript
/*此示例基于“常规”样式创建一个新的样式，修改边框和字体，然后将该新样式应用到单元格 A25:A30。*/
function test() {
    let style = Application.ActiveWorkbook.Styles.Add("Bookman Top Border")
    style.Borders.Item(xlEdgeTop).LineStyle = xlDouble
    style.Font.Bold = true
    style.Font.Name = "Bookman"
    Application.Worksheets.Item(1).Range("A25:A30").Style = "Bookman Top Border"
}
```

使用Styles(index)（其中index是样式索引号或名称）可从工作簿的Style集合中返回一个Styles对象。

javascript
```javascript
/*此示例通过设置活动工作簿中“常规”样式的 Bold 属性来更改该样式。*/
function test() {
    Application.ActiveWorkbook.Styles.Item("Normal").Font.Bold = true
}
```


#### TableStyle 对象

# [TableStyle (对象)​](#tablestyle-对象)

代表可应用于表格或切片器的单个样式。

## [说明​](#说明)

表格样式为表格、数据透视表或切片器的一个或所有元素定义格式。例如，列是表格的元素。表格样式可以规定使用交替格式（也称为条带或条纹）对表格中的列进行格式设置。

## [示例​](#示例)

javascript
```javascript
/*本示例删除第一张工作簿上的表格样式中的最后一个样式。*/
function test() {
    let tableStyles = Workbooks.Item(1).TableStyles
    tableStyles.Item(tableStyles.Count).Delete()
}
```

javascript
```javascript
/*本示例显示第一张工作表上第一张列表的表样式是否为内置样式。*/
function test() {
    let listObj = Application.Worksheets.Item(1).ListObjects.Item(1)
    console.log(listObj.TableStyle.BuiltIn)
}
```


#### TableStyleElement 对象

# [TableStyleElement (对象)​](#tablestyleelement-对象)

代表单个表格样式元素。

## [说明​](#说明)

表格样式为表格、数据透视表或切片器的一个或所有元素定义格式。例如，标题行是表格的元素。表格样式可以规定标题行的填充色为红色。

表格中每个表格样式元素的格式设置可在适用于该元素的表格样式中指定。

## [示例​](#示例)

javascript
```javascript
/*本示例清除活动工作表上第一张列表中总计行样式元素的格式。*/
function test() {
    let listObj = ActiveSheet.ListObjects.Item(1)
    listObj.TableStyle.TableStyleElements.Item(xlGrandTotalRow).Clear()
}
```

javascript
```javascript
/*本示例显示第一张工作表上第一张列表的行条纹1样式元素条带的大小。*/
function test() {
    let listObj = Worksheets.Item(1).ListObjects.Item(1)
    console.log(listObj.TableStyle.TableStyleElements.Item(xlRowStripe1).StripeSize)
}
```


#### TableStyleElements 对象

# [TableStyleElements (对象)​](#tablestyleelements-对象)

代表表格样式元素。

## [说明​](#说明)

表格样式为表格、数据透视表或切片器的一个或所有元素定义格式。例如，标题行、最后一列或总计行是表格的元素，表格样式可以规定标题行的填充色为蓝色，最后一列为红色。

表格中表格样式元素的格式设置可在适用于该元素的表格样式中指定。XlTableStyleElementType枚举包含可供使用的表格样式元素的类型。

## [示例​](#示例)

javascript
```javascript
/*本示例显示第一张工作表上第一张列表是否存在表格样式元素。*/
function test() {
    let listObj = Application.Worksheets.Item(1).ListObjects.Item(1)
    console.log(listObj.TableStyle.TableStyleElements.Count > 0)
}
```

javascript
```javascript
/*本示例将活动工作表上第一张列表的最后一列样式元素的内部设置为红色。*/
function test() {
    let listObj = ActiveSheet.ListObjects.Item(1)
    listObj.TableStyle.TableStyleElements.Item(xlLastColumn).Interior.ColorIndex = 3
}
```


#### TableStyles 对象

# [TableStyles (对象)​](#tablestyles-对象)

代表可应用于表格的样式。

## [说明​](#说明)

表格样式提供了一种为整个表格或数据透视图设置格式的方式。表格样式取代了用于为整个表格设置格式的现有的自动套用格式功能。

表格样式与自动套用格式在以下几个方面不同：

可以创建和重用自定义表格样式。
表格样式可处理主题。
如果更改文档主题的配色方案和/或字体方案，则将更改内置表格样式的外观。
当对象发生变化时，表格样式可以将样式重新应用于像数据透视表和表格之类的对象。该表格将记住应用于对象的样式，在添加、删除、隐藏和显示单元格时，表格将相应地进行重新显示。
表格样式在功能区中具有可见的用户界面。
## [示例​](#示例)

javascript
```javascript
/*本示例显示第一张工作簿中所使用的样式数量。*/
function test() {
    console.log(Workbooks.Item(1).TableStyles.Count)
}
```

javascript
```javascript
/*本示例向活动工作簿的表格样式中添加“样式1”。*/
function test() {
    ActiveWorkbook.TableStyles.Add("样式1")
}
```


#### TextEffectFormat 对象

# [TextEffectFormat (对象)​](#texteffectformat-对象)

包含应用于艺术字对象的属性和方法。

## [说明​](#说明)

使用TextEffect属性可返回一个TextEffectFormat对象。

## [示例​](#示例)

javascript
```javascript
/*下例为第一张工作表上的形状一设置字体名称及格式。要运行本示例，形状一必须是艺术字对象。*/
function test() {
    let worksheet = Application.Worksheets.Item(1)
    let TextEffect = worksheet.Shapes(1).TextEffect
    TextEffect.FontName = "Courier New"
    TextEffect.FontBold = true
    TextEffect.FontItalic = true
}
```

javascript
```javascript
/*此示例显示活动工作表的第三个艺术字的字号。*/
function test() {
    let shape = ActiveSheet.Shapes.Item(3)
    console.log(shape.TextEffect.FontSize)
}
```


#### TextFrame 对象

# [TextFrame (对象)​](#textframe-对象)

代表Shape对象中的文本框架。包含文本框架中的文本以及控制文本框架的对齐和定位的属性和方法。

## [说明​](#说明)

使用TextFrame属性可返回一个TextFrame对象。

## [示例​](#示例)

javascript
```javascript
/*本示例在第一张工作表中添加一个矩形，向矩形中添加文本，然后设置文本框架的边距。*/
function test() {
    let textFrame = Application.Worksheets.Item(1).Shapes.AddShape(msoShapeRectangle, 0, 0, 250, 140).TextFrame
    textFrame.Characters().Text = "Here is some test text"
    textFrame.MarginBottom = 10
    textFrame.MarginLeft = 10
    textFrame.MarginRight = 10
    textFrame.MarginTop = 10
}
```

javascript
```javascript
/*本示例在活动工作表中添加一个椭圆，向椭圆中添加文本，并将指定文本颜色改为绿色。*/
function test() {
    let textFrame = ActiveSheet.Shapes.AddShape(msoShapeOval, 100, 100, 200, 100).TextFrame
    textFrame.Characters().Text = "这是示例文本"
    textFrame.Characters(3, 2).Font.Color = RGB(0, 255, 0)
}
```


#### TextFrame2 对象

# [TextFrame2 (对象)​](#textframe2-对象)

代表Shape、ShapeRange或ChartFormat对象的文本框。

## [说明​](#说明)

该对象包含文本框中的文本，还包含控制文本框对齐方式和位置的属性和方法。使用TextFrame2属性可返回TextFrame2对象。

## [示例​](#示例)

javascript
```javascript
/*本示例在第一张工作表中添加一个矩形，向矩形中添加文本，然后设置文本框的边距。*/
function test() {
    let textFrame2 = Application.Worksheets.Item(1).Shapes.AddShape(msoShapeRectangle, 0, 0, 250, 140).TextFrame2
    textFrame2.TextRange.Text = "Here is some test text"
    textFrame2.MarginBottom = 10
    textFrame2.MarginLeft = 10
    textFrame2.MarginRight = 10
    textFrame2.MarginTop = 10
}
```

javascript
```javascript
/*本示例判断如果活动工作表中第二个形状的文本框包含文本，则显示文本长度。*/
function test() {
    let textFrame2 = ActiveSheet.Shapes.Item(2).TextFrame2
    if (textFrame2.HasText == msoTrue) {
        console.log(textFrame2.TextRange.Length)
    }
}
```


#### ThreeDFormat 对象

# [ThreeDFormat (对象)​](#threedformat-对象)

该对象代表一个形状的三维格式。

## [说明​](#说明)

不能对某些形状应用三维格式，例如斜截形状或多处间断的路径。对这些形状，ThreeDFormat对象的大多数属性和方法将失败。

使用ThreeD属性可返回一个ThreeDFormat对象。

## [示例​](#示例)

javascript
```javascript
/*以下示例向第一张工作表添加一个椭圆，然后指定椭圆延伸至 50 磅深度，延伸为紫色。*/
function test() {
    let worksheet = Worksheets.Item(1)
    let shape = worksheet.Shapes.AddShape(msoShapeOval, 90, 90, 90, 40)
    let threeDFormat = shape.ThreeD
    threeDFormat.Visible = true
    threeDFormat.Depth = 50
    threeDFormat.ExtrusionColor.RGB = RGB(255, 100, 255)
    //RGB value for purple
}
```

javascript
```javascript
/*本示例显示活动工作表中第三个形状突出的深度。*/
function test() {
    let shapes = ActiveSheet.Shapes
    let shape = shapes.Item(3)
    console.log(shape.ThreeD.Depth)
}
```


#### TickLabels 对象

# [TickLabels (对象)​](#ticklabels-对象)

代表图表坐标轴上刻度线的刻度线标志。

## [说明​](#说明)

此对象不是集合。没有代表单个刻度线标志的对象；您必须将所有刻度线标志作为一个单位返回。

分类轴的刻度线标签文字来自图表中相应分类的名称。分类轴的默认刻度线标签文字表示该分类相对于该坐标轴最左端偏移量的数字。要更改分类轴刻度线标签间不带标签的刻度线数量，您必须更改分类轴的TickLabelSpacing属性。

数值轴的刻度线标志文字的计算，是基于数值轴的MajorUnit、MinimumScale和MaximumScale属性。要更改数值的轴刻度线标签文本，您必须更改这些属性的值。

使用TickLabels属性可返回TickLabels对象。

## [示例​](#示例)

javascript
```javascript
/*下例设置 Sheet1 上嵌入式图表一中数值轴上刻度线标志的数字格式。*/
function test() {
    Worksheets.Item("Sheet1").ChartObjects(1).Chart.Axes(xlValue).TickLabels.NumberFormat = "0.00"
}
```

javascript
```javascript
/*以下示例删除 Chart1 上分类轴的刻度线标志。*/
function test() {
    let tickLabels = Application.Charts.Item("Chart1").ChartObjects(1).Chart.Axes(xlCategory).TickLabels
    tickLabels.Delete()
}
```


#### Top10 对象

# [Top10 (对象)​](#top10-对象)

代表条件格式规则的前十项。通过对某一区域应用颜色，有助于查看相对于其他单元格的单元格的值。

## [说明​](#说明)

所有条件格式设置对象都包含在FormatConditions集合对象中，该对象是Range集合的子对象。

可以使用FormatConditions集合的Add或AddTop10方法创建前 10 个格式规则。

## [示例​](#示例)

javascript
```javascript
/*本示例通过条件格式规则生成一个动态数据集并对前 10 个值应用颜色。*/
function test() {
    //Building data
    Application.Range("A1").Value2 = "Name"
    Application.Range("B1").Value2 = "Number"
    Application.Range("A2").Value2 = "Agent1"
    Application.Range("A2").AutoFill(Application.Range("A2:A26"), xlFillDefault)
    Application.Range("B2:B26").FormulaArray = "=INT(RAND()*101)"
    Application.Range("B2:B26").Select()

    //Applying Conditional Formatting Top 10
    Application.Selection.FormatConditions.AddTop10()
    Application.Selection.FormatConditions.Item(Application.Selection.FormatConditions.Count).SetFirstPriority()
    let top = Application.Selection.FormatConditions.Item(1)
    top.TopBottom = xlTop10Top
    top.Rank = 10
    top.Percent = false

    //Applying color fill
    let font = Application.Selection.FormatConditions.Item(1).Font
    font.Color = RGB(0, 155, 115)
    font.TintAndShade = 0
    let interior = Application.Selection.FormatConditions.Item(1).Interior
    interior.PatternColorIndex = xlAutomatic
    interior.Color = RGB(5, 185, 115)
    interior.TintAndShade = 0
}
```

javascript
```javascript
/*本示例设置第一张工作表上区域 E1:E10 中第一个（Top10）条件格式的TopBottom属性，并将该条件格式设置为按百分比值确定排位，然后设置该条件格式的排位值的百分比。*/
function test() {
    let top = Worksheets.Item(1).Range("E1:E10").FormatConditions.Item(1)
    top.TopBottom = xlTop10Bottom
    top.Percent = true
    top.Rank = 20
}
```


#### Trendline 对象

# [Trendline (对象)​](#trendline-对象)

代表图表上的趋势线。

## [说明​](#说明)

趋势线显示系列中数据的趋势或方向。Trendline对象是Trendlines集合的成员。Trendlines集合包含某一个系列的所有Trendline对象。

使用Trendlines(index)（其中index是趋势线索引号）可返回一个Trendline对象。

索引号指出趋势线添加到系列中的顺序。Trendlines(1)是第一个添加到系列中的趋势线，而Trendlines(Trendlines.Count)是最后一个。

## [示例​](#示例)

javascript
```javascript
/*下例更改工作表一上嵌入式图表一中第一个系列的趋势线类型。如果该系列没有趋势线，则本示例会失败。*/
function test() {
    Application.Worksheets.Item(1).ChartObjects(1).Chart.SeriesCollection(1).Trendlines(1).Type = xlMovingAvg
}
```

javascript
```javascript
/*本示例删除 Chart1 中第三个数据系列索引为2的趋势线。*/
function test() {
    let series = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(3)
    series.Trendlines(2).Delete()
}
```


#### Trendlines 对象

# [Trendlines (对象)​](#trendlines-对象)

指定的数据系列中所有Trendline对象的集合。

## [说明​](#说明)

每一个Trendline对象都代表图表中的趋势线。趋势线显示系列中数据的趋势或方向。

## [示例​](#示例)

javascript
```javascript
/*以下示例显示 Chart1 上第一系列的趋势线数。*/
function test() {
    console.log(Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(1).Trendlines().Count)
}
```

javascript
```javascript
/*下例给 Sheet1 中嵌入式图表一中的第一个系列添加线性趋势线。*/
function test() {
    Worksheets("Sheet1").ChartObjects(1).Chart.SeriesCollection(1).Trendlines().Add(xlLinear, null, null, null, null, null, null, null, "Linear Trend")
}
```

javascript
```javascript
/*下例更改工作表一上嵌入式图表一中第一个系列的趋势线类型。如果该系列没有趋势线，本示例将失败。*/
function test() {
    Worksheets(1).ChartObjects(1).Chart.SeriesCollection(1).Trendlines(1).Type = xlMovingAvg
}
```


#### UniqueValues 对象

# [UniqueValues (对象)​](#uniquevalues-对象)

UniqueValues对象使用DupeUnique属性返回或设置一个枚举，该枚举确定规则是查找区域中的重复值还是唯一值。

## [示例​](#示例)

javascript
```javascript
/*本示例通过条件格式规则生成一个动态数据集并对重复值应用颜色。*/
function test() {
    Application.Range("A1").Value2 = "Name"
    Application.Range("B1").Value2 = "Number"
    Application.Range("A2").Value2 = "Agent1"
    Application.Range("A2").AutoFill(Application.Range("A2:A26"), xlFillDefault)
    Application.Range("B2:B26").FormulaArray = "=INT(RAND()*101)"
    Application.Range("B2:B26").Select()

    Application.Selection.FormatConditions.AddUniqueValues()
    Application.Selection.FormatConditions.Item(Application.Selection.FormatConditions.Count).SetFirstPriority()
    Application.Selection.FormatConditions.Item(1).DupeUnique = xlDuplicate

    let font = Application.Selection.FormatConditions.Item(1).Font
    font.ColorIndex = 3
    font.TintAndShade = 0
    let interior = Application.Selection.FormatConditions.Item(1).Interior
    interior.PatternColorIndex = xlAutomatic
    interior.ColorIndex = 5
    interior.TintAndShade = 0
}
```

javascript
```javascript
/*本示例设置第一张工作表上单元格区域 C1:C10 的第一个（UniqueValues）条件格式所应用于的单元格区域，并设置该条件格式内部唯一值的颜色。*/
function test() {
    let uniqueValues = Application.Worksheets.Item(1).Range("C1:C10").FormatConditions.Item(1)
    uniqueValues.ModifyAppliesToRange(Range("C7:C10"))
    uniqueValues.DupeUnique = xlUnique
    uniqueValues.Interior.ColorIndex = 7
}
```


#### UpBars 对象

# [UpBars (对象)​](#upbars-对象)

涨柱线将图表组中第一个系列的数据点与最后一个系列中相应的有较大值的数据点连接起来（从第一个系列向上生长）。只有至少包含两个系列的二维折线图才能有涨柱线。此对象不是集合。没有代表单个涨柱线的对象；或者打开图表组中所有数据点的涨柱线，或者将其全部关闭。

## [说明​](#说明)

使用UpBars属性可返回UpBars对象。

## [示例​](#示例)

javascript
```javascript
/*本示例打开工作表“Sheet5”上嵌入的第一个图表中第一个图表组的涨跌柱线，然后将涨柱线的颜色设置为蓝色，而将跌柱线设置为红色。*/
function test() {
    let chartgroup = Application.Worksheets.Item("Sheet5").ChartObjects(1).Chart.ChartGroups(1)
    chartgroup.HasUpDownBars = true
    chartgroup.UpBars.Interior.Color = RGB(0, 0, 255)
    chartgroup.DownBars.Interior.Color = RGB(255, 0, 0)
}
```

javascript
```javascript
/*本示例判断如果图表工作表 Chart1 中图表的第一个图表组的涨柱线的名称为“涨柱线 1”，则删除涨跌柱线。*/
function test() {
    let upbars = Application.Charts.Item("Chart1").ChartObjects(1).Chart.ChartGroups(1).UpBars
    if (upbars.Name == "涨柱线 1") {
        upbars.Delete()
    }
}
```


#### UserAccess 对象

# [UserAccess (对象)​](#useraccess-对象)

代表对受保护区域的用户访问。

## [说明​](#说明)

使用 UserAccessList 集合的Add方法或 Item 属性可返回一个UserAccess对象。

一旦返回了UserAccess对象，您就可以使用AllowEdit属性来确定是否允许访问工作表中某个特定区域。下例添加一个在受保护的工作表上可编辑的区域，并通知用户该区域的标题。

javascript
```javascript
function test(){
    let wksSheet = Application.ActiveSheet
    
    //Add a range that can be edited on the protected worksheet.
    wksSheet.Protection.AllowEditRanges.Add("Test", Range("A1"))
    
    //Notify the user the title of the range that can be edited.
    console.log(wksSheet.Protection.AllowEditRanges.Item(1).Title)
}
```


#### UserAccessList 对象

# [UserAccessList (对象)​](#useraccesslist-对象)

代表受保护区域用户访问权限的UserAccess对象的集合。

## [说明​](#说明)

使用受保护的Range对象的Users属性可返回一个UserAccessList集合。

一旦返回了UserAccessList集合，您就可以使用Count属性来确定能够访问受保护区域的用户的数量。在下例中，ET 通知用户能够访问第一个受保护区域的用户的数量。本示例假定活动工作表中存在受保护区域。

javascript
```javascript
function test(){
    let wksSheet = Application.ActiveSheet
                                    
    //Notify the user the number of users that can access the protected range.
    console.log(wksSheet.Protection.AllowEditRanges.Item(1).Users.Count)
}
```


#### Validation 对象

# [Validation (对象)​](#validation-对象)

代表工作表区域的数据有效性规则。

## [说明​](#说明)

使用Validation属性可返回Validation对象。

## [示例​](#示例)

javascript
```javascript
/*本示例更改单元格 E5 的数据有效性验证。*/
function test() {
    Application.Range("E5").Validation.Modify(xlValidateList, xlValidAlertStop, null, "=$A$1:$A$10")
}
```

使用Add方法可将数据有效性添加到某个区域并创建一个新的Validation对象。

javascript
```javascript
/*本示例为 E5 单元格添加数据有效性验证*/
function test() {
    let validation = Application.Range("E5").Validation
    validation.Add(xlValidateWholeNumber, xlValidAlertStop, xlBetween, "5", "10")
    validation.InputTitle = "Integers"
    validation.ErrorTitle = "Integers"
    validation.InputMessage = "Enter an integer from five to ten"
    validation.ErrorMessage = "You must enter a number from five to ten"
}
```


#### Workbook 对象

# [Workbook (对象)​](#workbook-对象)

代表一个 ET 工作簿。

## [说明​](#说明)

Workbook对象是 Workbooks 集合的成员。Workbooks集合包含 ET 中当前打开的所有Workbook对象。

## [示例​](#示例)

使用Workbooks(index)（其中index是工作簿名称或索引号）可返回一个Workbook对象。下例激活工作簿一。

javascript
```javascript
/*以下示例激活第一个工作簿。*/
function test() {
    Application.Workbooks.Item(1).Activate()
}
```

编号指示创建或打开工作簿的顺序。Workbooks(1)是创建的第一个工作簿，而Workbooks(Workbooks.Count)Workbooks 是最后一个。激活某工作簿并不更改其索引号。所有工作簿均包括在索引计数中，即便是隐藏工作簿也是如此。

Name 属性返回工作簿名称。您不能通过使用此属性来设置该名称；如果您需要更改该名称，请使用 SaveAs 方法，将该工作簿保存为其他名称。

下例激活名为“Cogs.xls”的工作簿（该工作簿必须已经在 ET 中打开）中的 Sheet1。

Workbooks("Cogs.xls").Worksheets("Sheet1").Activate()

ActiveWorkbook属性返回当前处于活动状态的工作簿。下例设置活动工作簿作者的名称。

javascript
```javascript
/*以下示例设置活动工作簿作者的名称。*/
function test() {
    Application.ActiveWorkbook.Author = "Jean Selva"
}
```


#### Worksheet 对象

# [Worksheet (对象)​](#worksheet-对象)

代表一个工作表。

## [说明​](#说明)

Worksheet对象是 Worksheets 集合的成员。Worksheets集合包含某个工作簿中所有的Worksheet对象。

Worksheet对象也是 Sheets 集合的成员。Sheets集合包含工作簿中所有的工作表（图表工作表和工作表）。

使用 Worksheets(index)（其中index是工作表索引号或名称）可返回一个Worksheet对象。

## [示例​](#示例)

javascript
```javascript
/*本示例隐藏活动工作簿中第一张工作表。*/
function test() {
    Application.Worksheets.Item(1).Visible = false
}
```

工作表索引号指示该工作表在工作簿的标签栏上的位置。Worksheets(1)是工作簿中第一个（最左边的）工作表，而Worksheets(Worksheets.Count)是最后一个。所有工作表均包括在索引计数中，即便是隐藏工作表也是如此。

工作表名称显示在该工作表的标签上。使用 Name 属性可设置或返回工作表名称。

javascript
```javascript
/*本示例提醒用户输入密码，并用该密码保护工作表Sheet1上的方案。*/
function test() {
    let strPassword = "Enter the password for the worksheet"
    Application.Worksheets.Item("Sheet1").Protect(strPassword, null, null, true)
}
```

当工作表处于活动状态时，可以使用ActiveSheet属性来引用它。

javascript
```javascript
/*本示例激活工作表 Sheet1，将页面方向设置为横向，然后打印该工作表。*/
function test() {
    Application.Worksheets.Item("Sheet1").Activate()
    Application.ActiveSheet.PageSetup.Orientation = xlLandscape
    Application.ActiveSheet.PrintOut()
}
```


#### WorksheetFunction 对象

# [WorksheetFunction (对象)​](#worksheetfunction-对象)

用作可从 Visual Basic 中调用的 ET 工作表函数的容器。

## [说明​](#说明)

使用WorksheetFunction属性可返回WorksheetFunction对象。下例显示给区域 A1:A10 应用Min工作表函数的结果。

javascript
```javascript
function test(){
let myRange = Application.Worksheets.Item("Sheet1").Range("A1:C10")
let answer = Application.WorksheetFunction.Min(myRange)
console.log(answer)
}
```


#### Worksheets 对象

# [Worksheets (对象)​](#worksheets-对象)

指定的或活动工作簿中所有 Worksheet 对象的集合。每个Worksheet对象都代表一个工作表。

## [说明​](#说明)

Worksheet对象也是Sheets集合的成员。Sheets集合包含工作簿中所有的工作表（图表工作表和工作表）。

使用 Worksheets 属性可返回Worksheets集合。下例将所有工作表移到工作簿尾部。

javascript
```javascript
Application.Worksheets.Move(Application.Worksheets.Item(Application.Worksheets.Count))
```

使用 Add 方法可创建一个新工作表并将它添加到集合。下例将两个新工作表添加到活动工作簿的工作表一之前。

javascript
```javascript
Application.Worksheets.Add(Application.Worksheets.Item(1), undefined, 2)
```

使用Worksheets(index)（其中index是工作表索引号或名称）可返回一个Worksheet对象。下例隐藏活动工作簿中的工作表一。

javascript
```javascript
Application.Worksheets.Item(1).Visible = false
```


#### XlAboveBelow 枚举

# [XlAboveBelow 枚举​](#xlabovebelow-枚举)

指定值是高于还是低于平均值。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| XlAboveAverage | 0 | 高于平均值。 |
| XlAboveStdDev | 1 | 高于标准偏差。 |
| XlBelowAverage | 0 | 低于平均值。 |
| XlBelowStdDev | 1 | 低于标准偏差。 |
| XlEqualAboveAverage | 0 | 等于或高于平均值。 |
| XlEqualBelowAverage | 0 | 等于或低于平均值。 |


#### XlActionType 枚举

# [XlActionType 枚举​](#xlactiontype-枚举)

指定应执行的操作。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlActionTypeDrillthrough | 256 | 明细数据。 |
| xlActionTypeReport | 128 | 报表。 |
| xlActionTypeRowset | 16 | 行集。 |
| xlActionTypeUrl | 1 | URL。 |


#### XlAllocation 枚举

# [XlAllocation 枚举​](#xlallocation-枚举)

指定在对基于 OLAP 数据源的数据透视表执行模拟分析时，何时计算更改。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlAutomaticAllocation | 2 | 在每个值更改后自动计算更改。 |
| xlManualAllocation | 1 | 手动计算更改。 |


#### XlAllocationMethod 枚举

# [XlAllocationMethod 枚举​](#xlallocationmethod-枚举)

指定在对基于 OLAP 数据源的数据透视表执行模拟分析时，要用来分配值的方法。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlEqualAllocation | 1 | 使用平均分配。 |
| xlWeightedAllocation | 2 | 使用加权分配。 |


#### XlAllocationValue 枚举

# [XlAllocationValue 枚举​](#xlallocationvalue-枚举)

指定在对基于 OLAP 数据源的数据透视表执行模拟分析时，要分配什么值。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlAllocateIncrement | 2 | 在旧值的基础上递增。 |
| xlAllocateValue | 1 | 输入的值除以分配的次数。 |


#### XlApplicationInternational 枚举

# [XlApplicationInternational 枚举​](#xlapplicationinternational-枚举)

指定国家/地区和国际设置。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xl24HourClock | 33 | 如果使用 24 小时制时间，则返回 True；如果使用 12 小时时间，则返回 False。 |
| xl4DigitYears | 43 | 如果使用四位年，则返回 True；如果使用两位年，则返回 False。 |
| xlAlternateArraySeparator | 16 | 当前数组分隔符与小数分隔符相同时，用于替代的数组项分隔符。 |
| xlColumnSeparator | 14 | 字面数组中用于分隔列的列分隔符。 |
| xlCountryCode | 1 | ET 的国家/地区版本。 |
| xlCountrySetting | 2 | Windows 控制面板中的当前国家/地区设置。 |
| xlCurrencyBefore | 37 | 如果货币符号在货币值之前，则返回 True；如果货币符号在货币值之后，则返回 False。 |
| xlCurrencyCode | 25 | 货币符号。 |
| xlCurrencyDigits | 27 | 货币格式中使用的小数位数。 |
| xlCurrencyLeadingZeros | 40 | 如果显示零货币值的前导零，则返回 True。 |
| xlCurrencyMinusSign | 38 | 如果对负数使用负号，则返回 True；如果使用括号，则返回 False。 |
| xlCurrencyNegative | 28 | 负数货币值的货币格式：0 = (symbolx) 或 (xsymbol) 1 = -symbolx 或 -xsymbol 2 = symbol-x 或 x-symbol 3 = symbolx- 或 xsymbol-，其中 symbol 为国家/地区的货币符号。请注意货币符号的位置由 xlCurrencyBefore 确定。 |
| xlCurrencySpaceBefore | 36 | 如果在货币符号前面添加空格，则返回 True。 |
| xlCurrencyTrailingZeros | 39 | 如果显示零货币值的尾部零，则返回 True。 |
| xlDateOrder | 32 | 日期元素的次序：0 = 月-日-年 1 = 日-月-年 2 = 年-月-日 |
| xlDateSeparator | 17 | 日期分隔符 (/)。 |
| xlDayCode | 21 | 日符号 (d)。 |
| xlDayLeadingZero | 42 | 如果在日期中显示前导零，则返回 True。 |
| xlDecimalSeparator | 3 | 小数分隔符。 |
| xlGeneralFormatName | 26 | “常规”数字格式名称。 |
| xlHourCode | 22 | 小时符号 (h)。 |
| xlLeftBrace | 12 | 在字面数组中左大括号 ({) 的替代字符。 |
| xlLeftBracket | 10 | 在 R1C1-样式相对引用中左方括号 ([) 的替代字符。 |
| xlListSeparator | 5 | 列表分隔符。 |
| xlLowerCaseColumnLetter | 9 | 小写列字母。 |
| xlLowerCaseRowLetter | 8 | 小写行字母。 |
| xlMDY | 44 | 如果长日期显示中日期次序为月-日-年，则返回 True；如果次序为日-月-年，则返回 False。 |
| xlMetric | 35 | 如果使用米制度量系统，则返回 True；如果使用英制度量系统，则返回 False。 |
| xlMinuteCode | 23 | 分钟符号 (m)。 |
| xlMonthCode | 20 | 月符号 (m)。 |
| xlMonthLeadingZero | 41 | 如果以数字显示月份时显示月份中的前导零，则返回 True。 |
| xlMonthNameChars | 30 | 为了向后兼容总是返回三个字符。月份名称的缩写从 Microsoft Windows 中读取并且可以为任意长度。 |
| xlNoncurrencyDigits | 29 | 非货币格式中所使用的十进制数字的个数。 |
| xlNonEnglishFunctions | 34 | 如果不以英文显示函数，则返回 True。 |
| xlRightBrace | 13 | 在字面数组中右大括号 (}) 的替代字符。 |
| xlRightBracket | 11 | 在 R1C1-样式引用中右方括号 (]) 的替代字符。 |
| xlRowSeparator | 15 | 字面数组的行分隔符。 |
| xlSecondCode | 24 | 秒符号 (s)。 |
| xlThousandsSeparator | 4 | 零或千位分隔符。 |
| xlTimeLeadingZero | 45 | 如果时间中显示前导零，则返回 True。 |
| xlTimeSeparator | 18 | 时间分隔符 (😃。 |
| xlUpperCaseColumnLetter | 7 | 大写列字母。 |
| xlUpperCaseRowLetter | 6 | 大写行字母（对于 R1C1-样式引用）。 |
| xlWeekdayNameChars | 31 | 为了向后兼容总是返回三个字符。星期名称的缩写从 Microsoft Windows 中读取并且可以为任意长度。 |
| xlYearCode | 19 | 数字格式中的年符号 (y)。 |


#### XlApplyNamesOrder 枚举

# [XlApplyNamesOrder 枚举​](#xlapplynamesorder-枚举)

指定用行方向区域名称和列方向区域名称取代单元格引用时，首先列出哪个区域名称。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlColumnThenRow | 2 | 列在行之前列出。 |
| xlRowThenColumn | 1 | 行在列之前列出。 |


#### XlArabicModes 枚举

# [XlArabicModes 枚举​](#xlarabicmodes-枚举)

为阿拉伯语拼写检查器指定拼写规则。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlArabicBothStrict | 3 | 拼写检查器使用有关以字母 yaa 结尾和以 alef hamza 开头的阿拉伯语单词拼写规则。 |
| xlArabicNone | 0 | 拼写检查器忽略有关以字母 yaa 结尾或以 alef hamza 开头的阿拉伯语单词拼写规则。 |
| xlArabicStrictAlefHamza | 1 | 拼写检查器使用有关以 alef hamza 开头的阿拉伯语单词拼写规则。 |
| xlArabicStrictFinalYaa | 2 | 拼写检查器使用有关以字母 yaa 结尾的阿拉伯语单词拼写规则。 |


#### XlArrangeStyle 枚举

# [XlArrangeStyle 枚举​](#xlarrangestyle-枚举)

指定窗口在屏幕上的排列方式。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlArrangeStyleCascade | 7 | 层叠窗口。 |
| xlArrangeStyleHorizontal | -4128 | 水平排列窗口。 |
| xlArrangeStyleTiled | 1 | 默认值。平铺窗口。 |
| xlArrangeStyleVertical | -4166 | 垂直排列窗口。 |


#### XlArrowHeadLength 枚举

# [XlArrowHeadLength 枚举​](#xlarrowheadlength-枚举)

指定线条末端的箭头长度。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlArrowHeadLengthLong | 3 | 最长箭头。 |
| xlArrowHeadLengthMedium | -4138 | 中等长度箭头。 |
| xlArrowHeadLengthShort | 1 | 最短箭头。 |


#### XlArrowHeadStyle 枚举

# [XlArrowHeadStyle 枚举​](#xlarrowheadstyle-枚举)

指定线条末端应用的箭头类型。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlArrowHeadStyleClosed | 3 | 线条连接处边缘为曲线的小箭头。 |
| xlArrowHeadStyleDoubleClosed | 5 | 菱形大箭头。 |
| xlArrowHeadStyleDoubleOpen | 4 | 线条连接处边缘为曲线的大箭头。 |
| xlArrowHeadStyleNone | -4142 | 无箭头。 |
| xlArrowHeadStyleOpen | 2 | 三角形大箭头。 |


#### XlArrowHeadWidth 枚举

# [XlArrowHeadWidth 枚举​](#xlarrowheadwidth-枚举)

指定线条末端的箭头宽度。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlArrowHeadWidthMedium | -4138 | 中等宽度箭头。 |
| xlArrowHeadWidthNarrow | 1 | 最窄箭头。 |
| xlArrowHeadWidthWide | 3 | 最宽箭头。 |


#### XlAutoFillType 枚举

# [XlAutoFillType 枚举​](#xlautofilltype-枚举)

根据源区域的内容，指定目标区域的填充方式。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlFillCopy | 1 | 将源区域的值和格式复制到目标区域，如有必要可重复执行。 |
| xlFillDays | 5 | 将星期中每天的名称从源区域扩展到目标区域中。格式从源区域复制到目标区域，如有必要可重复执行。 |
| xlFillDefault | 0 | ET 确定用于填充目标区域的值和格式。 |
| xlFillFormats | 3 | 只将源区域的格式复制到目标区域，如有必要可重复执行。 |
| xlFillMonths | 7 | 将月名称从源区域扩展到目标区域中。格式从源区域复制到目标区域，如有必要可重复执行。 |
| xlFillSeries | 2 | 将源区域中的值扩展到目标区域中，形式为系列（如，“1, 2”扩展为“3, 4, 5”）。格式从源区域复制到目标区域，如有必要可重复执行。 |
| xlFillValues | 4 | 只将源区域的值复制到目标区域，如有必要可重复执行。 |
| xlFillWeekdays | 6 | 将工作周每天的名称从源区域扩展到目标区域中。格式从源区域复制到目标区域，如有必要可重复执行。 |
| xlFillYears | 8 | 将年从源区域扩展到目标区域中。格式从源区域复制到目标区域，如有必要可重复执行。 |
| xlGrowthTrend | 10 | 将数值从源区域扩展到目标区域中，假定源区域的数字之间是乘法关系（如，“1, 2,”扩展为“4, 8, 16”，假定每个数字都是前一个数字乘以某个值的结果）。格式从源区域复制到目标区域，如有必要可重复执行。 |
| xlLinearTrend | 9 | 将数值从源区域扩展到目标区域中，假定数字之间是加法关系（如，“1, 2,”扩展为“3, 4, 5”，假定每个数字都是前一个数字加上某个值的结果）。格式从源区域复制到目标区域，如有必要可重复执行。 |


#### XlAutoFilterOperator 枚举

# [XlAutoFilterOperator 枚举​](#xlautofilteroperator-枚举)

指定用于关联两个筛选条件的操作符。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlAnd | 1 | 条件 1 和条件 2 的逻辑与。 |
| xlBottom10Items | 4 | 显示最低值项（条件 1 中指定的项数）。 |
| xlBottom10Percent | 6 | 显示最低值项（条件 1 中指定的百分数）。 |
| xlFilterCellColor | 8 | 单元格颜色 |
| xlFilterDynamic | 11 | 动态筛选 |
| xlFilterFontColor | 9 | 字体颜色 |
| xlFilterIcon | 10 | 筛选图标 |
| xlFilterValues | 7 | 筛选值 |
| xlOr | 2 | 条件 1 和条件 2 的逻辑或。 |
| xlTop10Items | 3 | 显示最高值项（条件 1 中指定的项数）。 |
| xlTop10Percent | 5 | 显示最高值项（条件 1 中指定的百分数）。 |


#### XlBackground 枚举

# [XlBackground 枚举​](#xlbackground-枚举)

指定图表文本的背景类型。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlBackgroundAutomatic | -4105 | ET 控制背景。 |
| xlBackgroundOpaque | 3 | 不透明背景。 |
| xlBackgroundTransparent | 2 | 透明背景。 |


#### XlBordersIndex 枚举

# [XlBordersIndex 枚举​](#xlbordersindex-枚举)

指定要检索的边框。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlDiagonalDown | 5 | 从区域中每个单元格的左上角至右下角的边框。 |
| xlDiagonalUp | 6 | 从区域中每个单元格的左下角至右上角的边框。 |
| xlEdgeBottom | 9 | 区域底部的边框。 |
| xlEdgeLeft | 7 | 区域左边的边框。 |
| xlEdgeRight | 10 | 区域右边的边框。 |
| xlEdgeTop | 8 | 区域顶部的边框。 |
| xlInsideHorizontal | 12 | 区域中所有单元格的水平边框（区域以外的边框除外）。 |
| xlInsideVertical | 11 | 区域中所有单元格的垂直边框（区域以外的边框除外）。 |


#### XlBuiltInDialog 枚举

# [XlBuiltInDialog 枚举​](#xlbuiltindialog-枚举)

指定要显示的对话框。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlDialogActivate | 103 | “激活”对话框 |
| xlDialogActiveCellFont | 476 | “活动单元格字体”对话框 |
| xlDialogAddChartAutoformat | 390 | “添加图表自动套用格式”对话框 |
| xlDialogAddinManager | 321 | “加载项管理器”对话框 |
| xlDialogAlignment | 43 | “对齐方式”对话框 |
| xlDialogApplyNames | 133 | “应用名称”对话框 |
| xlDialogApplyStyle | 212 | “应用样式”对话框 |
| xlDialogAppMove | 170 | “AppMove”对话框 |
| xlDialogAppSize | 171 | “AppSize”对话框 |
| xlDialogArrangeAll | 12 | “全部重排”对话框 |
| xlDialogAssignToObject | 213 | “给对象指定宏”对话框 |
| xlDialogAssignToTool | 293 | “给工具指定宏”对话框 |
| xlDialogAttachText | 80 | “附加文本”对话框 |
| xlDialogAttachToolbars | 323 | “附加工具栏”对话框 |
| xlDialogAutoCorrect | 485 | “自动校正”对话框 |
| xlDialogAxes | 78 | “坐标轴”对话框 |
| xlDialogBorder | 45 | “边框”对话框 |
| xlDialogCalculation | 32 | “计算”对话框 |
| xlDialogCellProtection | 46 | “单元格保护”对话框 |
| xlDialogChangeLink | 166 | “更改链接”对话框 |
| xlDialogChartAddData | 392 | “图表添加数据”对话框 |
| xlDialogChartLocation | 527 | “图表位置”对话框 |
| xlDialogChartOptionsDataLabelMultiple | 724 | “图表选项多个数据标签”对话框 |
| xlDialogChartOptionsDataLabels | 505 | “图表选项数据标签”对话框 |
| xlDialogChartOptionsDataTable | 506 | “图表选项数据表”对话框 |
| xlDialogChartSourceData | 540 | “图表源数据”对话框 |
| xlDialogChartTrend | 350 | “图表趋势”对话框 |
| xlDialogChartType | 526 | “图表类型”对话框 |
| xlDialogChartWizard | 288 | “图表向导”对话框 |
| xlDialogCheckboxProperties | 435 | “复选框属性”对话框 |
| xlDialogClear | 52 | “清除”对话框 |
| xlDialogColorPalette | 161 | “调色板”对话框 |
| xlDialogColumnWidth | 47 | “列宽”对话框 |
| xlDialogCombination | 73 | “组合图”对话框 |
| xlDialogConditionalFormatting | 583 | “条件格式”对话框 |
| xlDialogConsolidate | 191 | “合并计算”对话框 |
| xlDialogCopyChart | 147 | “复制图表”对话框 |
| xlDialogCopyPicture | 108 | “复制图片”对话框 |
| xlDialogCreateList | 796 | “创建列表”对话框 |
| xlDialogCreateNames | 62 | “创建名称”对话框 |
| xlDialogCreatePublisher | 217 | “创建发布者”对话框 |
| xlDialogCustomizeToolbar | 276 | “自定义工具栏”对话框 |
| xlDialogCustomViews | 493 | “自定义视图”对话框 |
| xlDialogDataDelete | 36 | “数据删除”对话框 |
| xlDialogDataLabel | 379 | “数据标签”对话框 |
| xlDialogDataLabelMultiple | 723 | “多个数据标签”对话框 |
| xlDialogDataSeries | 40 | “数据系列”对话框 |
| xlDialogDataValidation | 525 | “数据有效性”对话框 |
| xlDialogDefineName | 61 | “定义名称”对话框 |
| xlDialogDefineStyle | 229 | “定义样式”对话框 |
| xlDialogDeleteFormat | 111 | “删除格式”对话框 |
| xlDialogDeleteName | 110 | “删除名称”对话框 |
| xlDialogDemote | 203 | “降级”对话框 |
| xlDialogDisplay | 27 | “显示”对话框 |
| xlDialogDocumentInspector | 862 | “文档检查器”对话框 |
| xlDialogEditboxProperties | 438 | “编辑框属性”对话框 |
| xlDialogEditColor | 223 | “编辑颜色”对话框 |
| xlDialogEditDelete | 54 | “编辑删除”对话框 |
| xlDialogEditionOptions | 251 | “编辑选项”对话框 |
| xlDialogEditSeries | 228 | “编辑数据系列”对话框 |
| xlDialogErrorbarX | 463 | “误差线 X”对话框 |
| xlDialogErrorbarY | 464 | “误差线 Y”对话框 |
| xlDialogErrorChecking | 732 | “错误检查”对话框 |
| xlDialogEvaluateFormula | 709 | “公式求值”对话框 |
| xlDialogExternalDataProperties | 530 | “外部数据属性”对话框 |
| xlDialogExtract | 35 | “提取”对话框 |
| xlDialogFileDelete | 6 | “文件删除”对话框 |
| xlDialogFileSharing | 481 | “文件共享”对话框 |
| xlDialogFillGroup | 200 | “填充组”对话框 |
| xlDialogFillWorkgroup | 301 | “填充工作组”对话框 |
| xlDialogFilter | 447 | “对话框筛选”对话框 |
| xlDialogFilterAdvanced | 370 | “高级筛选”对话框 |
| xlDialogFindFile | 475 | “查找文件”对话框 |
| xlDialogFont | 26 | “字体”对话框 |
| xlDialogFontProperties | 381 | “字体属性”对话框 |
| xlDialogFormatAuto | 269 | “自动套用格式”对话框 |
| xlDialogFormatChart | 465 | “设置图表格式”对话框 |
| xlDialogFormatCharttype | 423 | “设置图表类型格式”对话框 |
| xlDialogFormatFont | 150 | “设置字体格式”对话框 |
| xlDialogFormatLegend | 88 | “图例格式”对话框 |
| xlDialogFormatMain | 225 | “设置主要格式”对话框 |
| xlDialogFormatMove | 128 | “设置移动格式”对话框 |
| xlDialogFormatNumber | 42 | “设置数字格式”对话框 |
| xlDialogFormatOverlay | 226 | “设置重叠格式”对话框 |
| xlDialogFormatSize | 129 | “设置大小”对话框 |
| xlDialogFormatText | 89 | “设置文本格式”对话框 |
| xlDialogFormulaFind | 64 | “查找公式”对话框 |
| xlDialogFormulaGoto | 63 | “转到公式”对话框 |
| xlDialogFormulaReplace | 130 | “替换公式”对话框 |
| xlDialogFunctionWizard | 450 | “函数向导”对话框 |
| xlDialogGallery3dArea | 193 | “三维面积图库”对话框 |
| xlDialogGallery3dBar | 272 | “三维条形图库”对话框 |
| xlDialogGallery3dColumn | 194 | “三维柱形图库”对话框 |
| xlDialogGallery3dLine | 195 | “三维折线图库”对话框 |
| xlDialogGallery3dPie | 196 | “三维饼图库”对话框 |
| xlDialogGallery3dSurface | 273 | “三维曲面图库”对话框 |
| xlDialogGalleryArea | 67 | “面积图库”对话框 |
| xlDialogGalleryBar | 68 | “条形图库”对话框 |
| xlDialogGalleryColumn | 69 | “柱形图库”对话框 |
| xlDialogGalleryCustom | 388 | “自定义库”对话框 |
| xlDialogGalleryDoughnut | 344 | “圆环图库”对话框 |
| xlDialogGalleryLine | 70 | “折线图库”对话框 |
| xlDialogGalleryPie | 71 | “饼图库”对话框 |
| xlDialogGalleryRadar | 249 | “雷达图库”对话框 |
| xlDialogGalleryScatter | 72 | “散点图库”对话框 |
| xlDialogGoalSeek | 198 | “单变量求解”对话框 |
| xlDialogGridlines | 76 | “网格线”对话框 |
| xlDialogImportTextFile | 666 | “导入文本文件”对话框 |
| xlDialogInsert | 55 | “插入”对话框 |
| xlDialogInsertHyperlink | 596 | “插入超链接”对话框 |
| xlDialogInsertObject | 259 | “插入对象”对话框 |
| xlDialogInsertPicture | 342 | “插入图片”对话框 |
| xlDialogInsertTitle | 380 | “插入标题”对话框 |
| xlDialogLabelProperties | 436 | “标签属性”对话框 |
| xlDialogListboxProperties | 437 | “列表框属性”对话框 |
| xlDialogMacroOptions | 382 | “宏选项”对话框 |
| xlDialogMailEditMailer | 470 | “编辑邮件发件人”对话框 |
| xlDialogMailLogon | 339 | “邮件登录”对话框 |
| xlDialogMailNextLetter | 378 | “发送下一信函”对话框 |
| xlDialogMainChart | 85 | “主要图”对话框 |
| xlDialogMainChartType | 185 | “图表类型”对话框 |
| xlDialogMenuEditor | 322 | “菜单编辑器”对话框 |
| xlDialogMove | 262 | “移动”对话框 |
| xlDialogMyPermission | 834 | “我的权限”对话框 |
| xlDialogNameManager | 977 | “名称管理器”对话框 |
| xlDialogNew | 119 | “新建”对话框 |
| xlDialogNewName | 978 | “新建名称”对话框 |
| xlDialogNewWebQuery | 667 | “新建 Web 查询”对话框 |
| xlDialogNote | 154 | “注意”对话框 |
| xlDialogObjectProperties | 207 | “对象属性”对话框 |
| xlDialogObjectProtection | 214 | “对象保护”对话框 |
| xlDialogOpen | 1 | “打开”对话框 |
| xlDialogOpenLinks | 2 | “打开链接”对话框 |
| xlDialogOpenMail | 188 | “打开邮件”对话框 |
| xlDialogOpenText | 441 | “打开文本”对话框 |
| xlDialogOptionsCalculation | 318 | “计算选项”对话框 |
| xlDialogOptionsChart | 325 | “图表选项”对话框 |
| xlDialogOptionsEdit | 319 | “编辑选项”对话框 |
| xlDialogOptionsGeneral | 356 | “常规选项”对话框 |
| xlDialogOptionsListsAdd | 458 | “添加列表选项”对话框 |
| xlDialogOptionsME | 647 | “ME 选项”对话框 |
| xlDialogOptionsTransition | 355 | “转换选项”对话框 |
| xlDialogOptionsView | 320 | “视图选项”对话框 |
| xlDialogOutline | 142 | “大纲”对话框 |
| xlDialogOverlay | 86 | “覆盖图”对话框 |
| xlDialogOverlayChartType | 186 | “覆盖图图表类型”对话框 |
| xlDialogPageSetup | 7 | “页面设置”对话框 |
| xlDialogParse | 91 | “分列”对话框 |
| xlDialogPasteNames | 58 | “粘贴名称”对话框 |
| xlDialogPasteSpecial | 53 | “选择性粘贴”对话框 |
| xlDialogPatterns | 84 | “图案”对话框 |
| xlDialogPermission | 832 | “权限”对话框 |
| xlDialogPhonetic | 656 | “拼音”对话框 |
| xlDialogPivotCalculatedField | 570 | “数据透视表计算字段”对话框 |
| xlDialogPivotCalculatedItem | 572 | “数据透视表计算项”对话框 |
| xlDialogPivotClientServerSet | 689 | “设置数据透视表客户机服务器”对话框 |
| xlDialogPivotFieldGroup | 433 | “组合数据透视表字段”对话框 |
| xlDialogPivotFieldProperties | 313 | “数据透视表字段属性”对话框 |
| xlDialogPivotFieldUngroup | 434 | “取消组合数据透视表字段”对话框 |
| xlDialogPivotShowPages | 421 | “数据透视表显示页”对话框 |
| xlDialogPivotSolveOrder | 568 | “数据透视表求解次序”对话框 |
| xlDialogPivotTableOptions | 567 | “数据透视表选项”对话框 |
| xlDialogPivotTableWizard | 312 | “数据透视表向导”对话框 |
| xlDialogPlacement | 300 | “位置”对话框 |
| xlDialogPrint | 8 | “打印”对话框 |
| xlDialogPrinterSetup | 9 | “打印机设置”对话框 |
| xlDialogPrintPreview | 222 | “打印预览”对话框 |
| xlDialogPromote | 202 | “升级”对话框 |
| xlDialogProperties | 474 | “属性”对话框 |
| xlDialogPropertyFields | 754 | “属性字段”对话框 |
| xlDialogProtectDocument | 28 | “保护文档”对话框 |
| xlDialogProtectSharing | 620 | “保护共享”对话框 |
| xlDialogPublishAsWebPage | 653 | “发布为网页”对话框 |
| xlDialogPushbuttonProperties | 445 | “按钮属性”对话框 |
| xlDialogReplaceFont | 134 | “替换字体”对话框 |
| xlDialogRoutingSlip | 336 | “传送名单”对话框 |
| xlDialogRowHeight | 127 | “行高”对话框 |
| xlDialogRun | 17 | “运行”对话框 |
| xlDialogSaveAs | 5 | “另存为”对话框 |
| xlDialogSaveCopyAs | 456 | “副本另存为”对话框 |
| xlDialogSaveNewObject | 208 | “保存新对象”对话框 |
| xlDialogSaveWorkbook | 145 | “保存工作簿”对话框 |
| xlDialogSaveWorkspace | 285 | “保存工作区”对话框 |
| xlDialogScale | 87 | “缩放”对话框 |
| xlDialogScenarioAdd | 307 | “添加方案”对话框 |
| xlDialogScenarioCells | 305 | “单元格方案”对话框 |
| xlDialogScenarioEdit | 308 | “编辑方案”对话框 |
| xlDialogScenarioMerge | 473 | “合并方案”对话框 |
| xlDialogScenarioSummary | 311 | “方案摘要”对话框 |
| xlDialogScrollbarProperties | 420 | “滚动条属性”对话框 |
| xlDialogSearch | 731 | “搜索”对话框 |
| xlDialogSelectSpecial | 132 | “特殊选定”对话框 |
| xlDialogSendMail | 189 | “发送邮件”对话框 |
| xlDialogSeriesAxes | 460 | “系列坐标轴”对话框 |
| xlDialogSeriesOptions | 557 | “系列选项”对话框 |
| xlDialogSeriesOrder | 466 | “系列次序”对话框 |
| xlDialogSeriesShape | 504 | “系列形状”对话框 |
| xlDialogSeriesX | 461 | “系列 X”对话框 |
| xlDialogSeriesY | 462 | “系列 Y”对话框 |
| xlDialogSetBackgroundPicture | 509 | “设置背景图片”对话框 |
| xlDialogSetPrintTitles | 23 | “设置打印标题”对话框 |
| xlDialogSetUpdateStatus | 159 | “设置更新状态”对话框 |
| xlDialogShowDetail | 204 | “显示明细数据”对话框 |
| xlDialogShowToolbar | 220 | “显示工具栏”对话框 |
| xlDialogSize | 261 | “大小”对话框 |
| xlDialogSort | 39 | “排序”对话框 |
| xlDialogSortSpecial | 192 | “选择性排序”对话框 |
| xlDialogSplit | 137 | “拆分”对话框 |
| xlDialogStandardFont | 190 | “标准字体”对话框 |
| xlDialogStandardWidth | 472 | “标准宽度”对话框 |
| xlDialogStyle | 44 | “样式”对话框 |
| xlDialogSubscribeTo | 218 | “订阅”对话框 |
| xlDialogSubtotalCreate | 398 | “创建分类汇总”对话框 |
| xlDialogSummaryInfo | 474 | “摘要信息”对话框 |
| xlDialogTable | 41 | “表”对话框 |
| xlDialogTabOrder | 394 | “Tab 键次序”对话框 |
| xlDialogTextToColumns | 422 | “分列”对话框 |
| xlDialogUnhide | 94 | “取消隐藏”对话框 |
| xlDialogUpdateLink | 201 | “更新链接”对话框 |
| xlDialogVbaInsertFile | 328 | “VBA 插入文件”对话框 |
| xlDialogVbaMakeAddin | 478 | “VBA 创建加载项”对话框 |
| xlDialogVbaProcedureDefinition | 330 | “VBA 过程定义”对话框 |
| xlDialogView3d | 197 | “三维视图”对话框 |
| xlDialogWebOptionsBrowsers | 773 | “Web 浏览器选项”对话框 |
| xlDialogWebOptionsEncoding | 686 | “Web 编码选项”对话框 |
| xlDialogWebOptionsFiles | 684 | “Web 文件选项”对话框 |
| xlDialogWebOptionsFonts | 687 | “Web 字体选项”对话框 |
| xlDialogWebOptionsGeneral | 683 | “Web 常规选项”对话框 |
| xlDialogWebOptionsPictures | 685 | “Web 图片选项”对话框 |
| xlDialogWindowMove | 14 | “窗口移动”对话框 |
| xlDialogWindowSize | 13 | “窗口大小”对话框 |
| xlDialogWorkbookAdd | 281 | “添加工作簿”对话框 |
| xlDialogWorkbookCopy | 283 | “复制工作簿”对话框 |
| xlDialogWorkbookInsert | 354 | “插入工作簿”对话框 |
| xlDialogWorkbookMove | 282 | “移动工作簿”对话框 |
| xlDialogWorkbookName | 386 | “命名工作簿”对话框 |
| xlDialogWorkbookNew | 302 | “新建工作簿”对话框 |
| xlDialogWorkbookOptions | 284 | “工作簿选项”对话框 |
| xlDialogWorkbookProtect | 417 | “保护工作簿”对话框 |
| xlDialogWorkbookTabSplit | 415 | “拆分工作簿标签”对话框 |
| xlDialogWorkbookUnhide | 384 | “取消隐藏工作簿”对话框 |
| xlDialogWorkgroup | 199 | “工作组”对话框 |
| xlDialogWorkspace | 95 | “工作区”对话框 |
| xlDialogZoom | 256 | “缩放”对话框 |


#### XlCVError 枚举

# [XlCVError 枚举​](#xlcverror-枚举)

指定单元格错误号和值。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlErrDiv0 | 2007 | 错误号：2007 |
| xlErrNA | 2042 | 错误号：2042 |
| xlErrName | 2029 | 错误号：2029 |
| xlErrNull | 2000 | 错误号：2000 |
| xlErrNum | 2036 | 错误号：2036 |
| xlErrRef | 2023 | 错误号：2023 |
| xlErrValue | 2015 | 错误号：2015 |


#### XlCalcFor 枚举

# [XlCalcFor 枚举​](#xlcalcfor-枚举)

指定应计算的内容。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlAllValues | 0 | 所有值。 |
| xlColGroups | 2 | 柱形图组。 |
| xlRowGroups | 1 | 行组。 |


#### XlCalculatedMemberType 枚举

# [XlCalculatedMemberType 枚举​](#xlcalculatedmembertype-枚举)

指定数据透视表中计算成员的类型。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlCalculatedMember | 0 | 成员使用多维表达式 (MDX) 公式。 |
| xlCalculatedSet | 1 | 成员在多维数据集字段中包含集的 MDX 公式。 |


#### XlCalculation 枚举

# [XlCalculation 枚举​](#xlcalculation-枚举)

指定计算模式。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlCalculationAutomatic | -4105 | ET 控制重新计算。 |
| xlCalculationManual | -4135 | 用户请求时进行计算。 |
| xlCalculationSemiautomatic | 2 | ET 控制重新计算，但忽略表中的更改。 |


#### XlCalculationInterruptKey 枚举

# [XlCalculationInterruptKey 枚举​](#xlcalculationinterruptkey-枚举)

指定中断重新计算的键。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlAnyKey | 2 | 按任意键中断重新计算。 |
| xlEscKey | 1 | 按 Esc 键中断重新计算。 |
| xlNoKey | 0 | 按任何键都不能中断重新计算。 |


#### XlCalculationState 枚举

# [XlCalculationState 枚举​](#xlcalculationstate-枚举)

指定应用程序的计算状态。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlCalculating | 1 | 正在计算。 |
| xlDone | 0 | 计算完成。 |
| xlPending | 2 | 已进行会触发计算的更改，但还未执行重新计算。 |


#### XlCellChangedState 枚举

# [XlCellChangedState 枚举​](#xlcellchangedstate-枚举)

指定自创建数据透视表以来，或上次执行提交操作以来，数据透视表值单元格是否经过了编辑或重新计算。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlCellChangeApplied | 3 | 单元格中的值已经过编辑或重新计算，并且这些更改已应用于数据源。（只应用具有 OLAP 数据源的数据透视表） |
| xlCellChanged | 2 | 单元格中的值已经过编辑或重新计算。 |
| xlCellNotChanged | 1 | 单元格中的值未经过编辑或重新计算。 |


#### XlCellInsertionMode 枚举

# [XlCellInsertionMode 枚举​](#xlcellinsertionmode-枚举)

指定在指定工作表中添加或删除行的方式，以符合查询返回的记录集的行数。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlInsertDeleteCells | 1 | 插入或者删除部分行以符合新记录集所需要的实际行数。 |
| xlInsertEntireRows | 2 | 必要时插入所有行以允许溢出。不从工作表删除单元格或行。 |
| xlOverwriteCells | 0 | 不向工作表添加新的单元格或行。如果溢出则覆盖周围单元格中的数据。 |


#### XlCellType 枚举

# [XlCellType 枚举​](#xlcelltype-枚举)

指定单元格的类型。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlCellTypeAllFormatConditions | -4172 | 任意格式的单元格。 |
| xlCellTypeAllValidation | -4174 | 含有验证条件的单元格。 |
| xlCellTypeBlanks | 4 | 空单元格。 |
| xlCellTypeComments | -4144 | 含有注释的单元格。 |
| xlCellTypeConstants | 2 | 含有常量的单元格。 |
| xlCellTypeFormulas | -4123 | 含有公式的单元格。 |
| xlCellTypeLastCell | 11 | 所用区域中的最后一个单元格。 |
| xlCellTypeSameFormatConditions | -4173 | 格式相同的单元格。 |
| xlCellTypeSameValidation | -4175 | 验证条件相同的单元格。 |
| xlCellTypeVisible | 12 | 所有可见单元格。 |


#### XlChartGallery 枚举

# [XlChartGallery 枚举​](#xlchartgallery-枚举)

指定图表库。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlAnyGallery | 23 | 任意一个库。 |
| xlBuiltIn | 21 | 内置库。 |
| xlUserDefined | 22 | 用户定义的库。 |


#### XlChartLocation 枚举

# [XlChartLocation 枚举​](#xlchartlocation-枚举)

指定在何处重定位图表。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlLocationAsNewSheet | 1 | 将图表移动到新工作表。 |
| xlLocationAsObject | 2 | 将图表嵌入到现有工作表中。 |
| xlLocationAutomatic | 3 | ET 控制图表位置。 |


#### XlChartPicturePlacement 枚举

# [XlChartPicturePlacement 枚举​](#xlchartpictureplacement-枚举)

指定用户所选图片在三维条形图或柱形图中的某个条形上的位置。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlAllFaces | 7 | 在所有表面上显示。 |
| xlEnd | 2 | 在末端显示。 |
| xlEndSides | 3 | 在末端和侧面上显示。 |
| xlFront | 4 | 在前端显示。 |
| xlFrontEnd | 6 | 在前端和末端显示。 |
| xlFrontSides | 5 | 在前端和侧面上显示。 |
| xlSides | 1 | 在侧面上显示。 |


#### XlChartType 枚举

# [XlChartType 枚举​](#xlcharttype-枚举)

指定图表类型。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xl3DArea | -4098 | 三维面积图。 |
| xl3DAreaStacked | 78 | 三维堆积面积图。 |
| xl3DAreaStacked100 | 79 | 百分比堆积面积图。 |
| xl3DBarClustered | 60 | 三维簇状条形图。 |
| xl3DBarStacked | 61 | 三维堆积条形图。 |
| xl3DBarStacked100 | 62 | 三维百分比堆积条形图。 |
| xl3DColumn | -4100 | 三维柱形图。 |
| xl3DColumnClustered | 54 | 三维簇状柱形图。 |
| xl3DColumnStacked | 55 | 三维堆积柱形图。 |
| xl3DColumnStacked100 | 56 | 三维百分比堆积柱形图。 |
| xl3DLine | -4101 | 三维折线图。 |
| xl3DPie | -4102 | 三维饼图。 |
| xl3DPieExploded | 70 | 分离型三维饼图。 |
| xlArea | 1 | 面积图 |
| xlAreaStacked | 76 | 堆积面积图。 |
| xlAreaStacked100 | 77 | 百分比堆积面积图。 |
| xlBarClustered | 57 | 簇状条形图。 |
| xlBarOfPie | 71 | 复合条饼图。 |
| xlBarStacked | 58 | 堆积条形图。 |
| xlBarStacked100 | 59 | 百分比堆积条形图。 |
| xlBubble | 15 | 气泡图。 |
| xlBubble3DEffect | 87 | 三维气泡图。 |
| xlColumnClustered | 51 | 簇状柱形图。 |
| xlColumnStacked | 52 | 堆积柱形图。 |
| xlColumnStacked100 | 53 | 百分比堆积柱形图。 |
| xlConeBarClustered | 102 | 簇状条形圆锥图。 |
| xlConeBarStacked | 103 | 堆积条形圆锥图。 |
| xlConeBarStacked100 | 104 | 百分比堆积条形圆锥图。 |
| xlConeCol | 105 | 三维柱形圆锥图。 |
| xlConeColClustered | 99 | 簇状柱形圆锥图。 |
| xlConeColStacked | 100 | 堆积柱形圆锥图。 |
| xlConeColStacked100 | 101 | 百分比堆积柱形圆锥图。 |
| xlCylinderBarClustered | 95 | 簇状条形圆柱图。 |
| xlCylinderBarStacked | 96 | 堆积条形圆柱图。 |
| xlCylinderBarStacked100 | 97 | 百分比堆积条形圆柱图。 |
| xlCylinderCol | 98 | 三维柱形圆柱图。 |
| xlCylinderColClustered | 92 | 簇状柱形圆锥图。 |
| xlCylinderColStacked | 93 | 堆积柱形圆锥图。 |
| xlCylinderColStacked100 | 94 | 百分比堆积柱形圆柱图。 |
| xlDoughnut | -4120 | 圆环图。 |
| xlDoughnutExploded | 80 | 分离型圆环图。 |
| xlLine | 4 | 折线图。 |
| xlLineMarkers | 65 | 数据点折线图。 |
| xlLineMarkersStacked | 66 | 堆积数据点折线图。 |
| xlLineMarkersStacked100 | 67 | 百分比堆积数据点折线图。 |
| xlLineStacked | 63 | 堆积折线图。 |
| xlLineStacked100 | 64 | 百分比堆积折线图。 |
| xlPie | 5 | 饼图。 |
| xlPieExploded | 69 | 分离型饼图。 |
| xlPieOfPie | 68 | 复合饼图。 |
| xlPyramidBarClustered | 109 | 簇状条形棱锥图。 |
| xlPyramidBarStacked | 110 | 堆积条形棱锥图。 |
| xlPyramidBarStacked100 | 111 | 百分比堆积条形棱锥图。 |
| xlPyramidCol | 112 | 三维柱形棱锥图。 |
| xlPyramidColClustered | 106 | 簇状柱形棱锥图。 |
| xlPyramidColStacked | 107 | 堆积柱形棱锥图。 |
| xlPyramidColStacked100 | 108 | 百分比堆积柱形棱锥图。 |
| xlRadar | -4151 | 雷达图。 |
| xlRadarFilled | 82 | 填充雷达图。 |
| xlRadarMarkers | 81 | 数据点雷达图。 |
| xlStockHLC | 88 | 盘高-盘低-收盘图。 |
| xlStockOHLC | 89 | 开盘-盘高-盘低-收盘图。 |
| xlStockVHLC | 90 | 成交量-盘高-盘低-收盘图。 |
| xlStockVOHLC | 91 | 成交量-开盘-盘高-盘低-收盘图。 |
| xlSurface | 83 | 三维曲面图。 |
| xlSurfaceTopView | 85 | 曲面图（俯视图）。 |
| xlSurfaceTopViewWireframe | 86 | 曲面图（俯视线框图）。 |
| xlSurfaceWireframe | 84 | 三维曲面图（线框）。 |
| xlXYScatter | -4169 | 散点图。 |
| xlXYScatterLines | 74 | 折线散点图。 |
| xlXYScatterLinesNoMarkers | 75 | 无数据点折线散点图。 |
| xlXYScatterSmooth | 72 | 平滑线散点图。 |
| xlXYScatterSmoothNoMarkers | 73 | 无数据点平滑线散点图。 |


#### XlCheckInVersionType 枚举

# [XlCheckInVersionType 枚举​](#xlcheckinversiontype-枚举)

指定在使用CheckIn方法时签入文档的版本类型。适用于存储在 SharePoint 库中的工作簿。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlCheckInMajorVersion | 1 | 签入主要版本。 |
| xlCheckInMinorVersion | 0 | 签入次要版本。 |
| xlCheckInOverwriteVersion | 2 | 覆盖服务器上的当前版本。 |


#### XlClipboardFormat 枚举

# [XlClipboardFormat 枚举​](#xlclipboardformat-枚举)

指定 Microsoft Windows 剪贴板上的项的格式。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlClipboardFormatBIFF | 8 | 用于 ET 2.x 版本的二进制交换文件格式 |
| xlClipboardFormatBIFF12 | 63 | 二进制交换文件格式 12 |
| xlClipboardFormatBIFF2 | 18 | 二进制交换文件格式 2 |
| xlClipboardFormatBIFF3 | 20 | 二进制交换文件格式 3 |
| xlClipboardFormatBIFF4 | 30 | 二进制交换文件格式 4 |
| xlClipboardFormatBinary | 15 | 二进制格式 |
| xlClipboardFormatBitmap | 9 | 位图格式 |
| xlClipboardFormatCGM | 13 | CGM 格式 |
| xlClipboardFormatCSV | 5 | CSV 格式 |
| xlClipboardFormatDIF | 4 | DIF 格式 |
| xlClipboardFormatDspText | 12 | Dsp 文本格式 |
| xlClipboardFormatEmbeddedObject | 21 | 嵌入对象 |
| xlClipboardFormatEmbedSource | 22 | 嵌入源 |
| xlClipboardFormatLink | 11 | 链接 |
| xlClipboardFormatLinkSource | 23 | 链接到源文件 |
| xlClipboardFormatLinkSourceDesc | 32 | 链接到源说明 |
| xlClipboardFormatMovie | 24 | 影片 |
| xlClipboardFormatNative | 14 | 本地 |
| xlClipboardFormatObjectDesc | 31 | 对象说明 |
| xlClipboardFormatObjectLink | 19 | 对象链接 |
| xlClipboardFormatOwnerLink | 17 | 链接到所有者 |
| xlClipboardFormatPICT | 2 | 图片 |
| xlClipboardFormatPrintPICT | 3 | 打印图片 |
| xlClipboardFormatRTF | 7 | RTF 格式 |
| xlClipboardFormatScreenPICT | 29 | 屏幕图片 |
| xlClipboardFormatStandardFont | 28 | 标准字体 |
| xlClipboardFormatStandardScale | 27 | 标准刻度 |
| xlClipboardFormatSYLK | 6 | SYLK |
| xlClipboardFormatTable | 16 | 表 |
| xlClipboardFormatText | 0 | 文本 |
| xlClipboardFormatToolFace | 25 | 工具图面 |
| xlClipboardFormatToolFacePICT | 26 | 工具图面图片 |
| xlClipboardFormatVALU | 1 | 值 |
| xlClipboardFormatWK1 | 10 | 工作簿 |


#### XlCmdType 枚举

# [XlCmdType 枚举​](#xlcmdtype-枚举)

指定CommandText属性的值。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlCmdCube | 1 | 包含一个 OLAP 数据源多维数据集名称。 |
| xlCmdDefault | 4 | 包含 OLE DB 提供程序可识别的命令文本。 |
| xlCmdList | 5 | 包含指向列表数据的指针。 |
| xlCmdSql | 2 | 包含一个 SQL 语句。 |
| xlCmdTable | 3 | 包含用于访问 OLE DB 数据源的表名称。 |


#### XlColumnDataType 枚举

# [XlColumnDataType 枚举​](#xlcolumndatatype-枚举)

指定列的分列方式。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlDMYFormat | 4 | DMY 日期格式。 |
| xlDYMFormat | 7 | DYM 日期格式。 |
| xlEMDFormat | 10 | EMD 日期格式。 |
| xlGeneralFormat | 1 | 常规。 |
| xlMDYFormat | 3 | MDY 日期格式。 |
| xlMYDFormat | 6 | MYD 日期格式。 |
| xlSkipColumn | 9 | 列未分列。 |
| xlTextFormat | 2 | 文本。 |
| xlYDMFormat | 8 | YDM 日期格式。 |
| xlYMDFormat | 5 | YMD 日期格式。 |


#### XlCommandUnderlines 枚举

# [XlCommandUnderlines 枚举​](#xlcommandunderlines-枚举)

指定 ET for the Macintosh 中命令加下划线的状态。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlCommandUnderlinesAutomatic | -4105 | ET 控制命令加下划线的显示。 |
| xlCommandUnderlinesOff | -4146 | 不显示命令加下划线。 |
| xlCommandUnderlinesOn | 1 | 显示命令加下划线。 |


#### XlCommentDisplayMode 枚举

# [XlCommentDisplayMode 枚举​](#xlcommentdisplaymode-枚举)

指定单元格显示批注和批注标识符的方式。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlCommentAndIndicator | 1 | 任何时候都显示批注和标识符。 |
| xlCommentIndicatorOnly | -1 | 只显示标识符。鼠标指针在单元格上移动时显示批注。 |
| xlNoIndicator | 0 | 任何时候都不显示批注也不显示标识符。 |


#### XlConditionValueTypes 枚举

# [XlConditionValueTypes 枚举​](#xlconditionvaluetypes-枚举)

指定可以使用的条件值的类型。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlConditionValueAutomaticMax | 7 | 最长数据条与范围中的最大值成比例。 |
| xlConditionValueAutomaticMin | 6 | 最短数据条与范围中的最小值成比例。 |
| xlConditionValueFormula | 4 | 使用公式。 |
| xlConditionValueHighestValue | 2 | 值列表的最高值。 |
| xlConditionValueLowestValue | 1 | 值列表的最低值。 |
| xlConditionValueNone | -1 | 无条件值。 |
| xlConditionValueNumber | 0 | 使用数字。 |
| xlConditionValuePercent | 3 | 使用百分比。 |
| xlConditionValuePercentile | 5 | 使用百分点值。 |


#### XlConnectionType 枚举

# [XlConnectionType 枚举​](#xlconnectiontype-枚举)

指定数据库连接的类型。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlConnectionTypeODBC | 2 | ODBC |
| xlConnectionTypeOLEDB | 1 | OLEDB |
| xlConnectionTypeTEXT | 4 | 文本 |
| xlConnectionTypeWEB | 5 | Web |
| xlConnectionTypeXMLMAP | 3 | XML 映射 |


#### XlConsolidationFunction 枚举

# [XlConsolidationFunction 枚举​](#xlconsolidationfunction-枚举)

指定分类汇总函数。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlAverage | -4106 | 平均。 |
| xlCount | -4112 | 计数。 |
| xlCountNums | -4113 | 只计数数值。 |
| xlMax | -4136 | 最大值。 |
| xlMin | -4139 | 最小值。 |
| xlProduct | -4149 | 乘。 |
| xlStDev | -4155 | 基于样本的标准偏差。 |
| xlStDevP | -4156 | 基于全体数据的标准偏差。 |
| xlSum | -4157 | 总计。 |
| xlUnknown | 1000 | 未指定任何分类汇总函数。 |
| xlVar | -4164 | 基于样本的方差。 |
| xlVarP | -4165 | 基于全体数据的方差。 |


#### XlContainsOperator 枚举

# [XlContainsOperator 枚举​](#xlcontainsoperator-枚举)

指定函数使用的运算符。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlBeginsWith | 2 | 以指定的值开始。 |
| xlContains | 0 | 包含指定的值。 |
| xlDoesNotContain | 1 | 不包含指定的值。 |
| xlEndsWith | 3 | 以指定的值结束 |


#### XlCopyPictureFormat 枚举

# [XlCopyPictureFormat 枚举​](#xlcopypictureformat-枚举)

指定复制的图片的格式。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlBitmap | 2 | 位图（.bmp、.jpg、.gif）。 |
| xlPicture | -4147 | 绘制图片（.png、.wmf、.mix）。 |


#### XlCorruptLoad 枚举

# [XlCorruptLoad 枚举​](#xlcorruptload-枚举)

指定文件打开时的处理。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlExtractData | 2 | ET 尝试恢复工作簿中的数据。 |
| xlNormalLoad | 0 | 正常打开工作簿。 |
| xlRepairFile | 1 | ET 尝试修复工作簿。 |


#### XlCreator 枚举

# [XlCreator 枚举​](#xlcreator-枚举)

为 ET for Macintosh 指定 32 位创建者代码（十进制 1480803660、十六进制 5843454C、字符串 XCEL）。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlCreatorCode | 1480803660 | ET for Macintosh 创建者代码。 |


#### XlCredentialsMethod 枚举

# [XlCredentialsMethod 枚举​](#xlcredentialsmethod-枚举)

指定所用凭据方法的类型。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| CredentialsMethodIntegrated | 0 | 集成 |
| CredentialsMethodNone | 1 | 不使用凭据 |
| CredentialsMethodStored | 2 | 使用存储的凭据 |


#### XlCubeFieldSubType 枚举

# [XlCubeFieldSubType 枚举​](#xlcubefieldsubtype-枚举)

指定 CubeField 的子类型。

| 注释 |
| --- |
| 值的名称中含有“Cube”，以便不与CubeFieldType属性的xlMeasure和xlSet值重叠。如果名称相同，记忆式键入功能在 Visual Basic 环境中将不起作用，因为它会找到模棱两可的值。 |

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlCubeAttribute | 4 | 属性 |
| xlCubeCalculatedMeasure | 5 | 计算度量 |
| xlCubeHierarchy | 1 | 层次结构 |
| xlCubeKPIGoal | 7 | KPI 目标 |
| xlCubeKPIStatus | 8 | KPI 状态 |
| xlCubeKPITrend | 9 | KPI 趋势 |
| xlCubeKPIValue | 6 | KPI 值 |
| xlCubeKPIWeight | 10 | KPI 权数 |
| xlCubeMeasure | 2 | 度量 |
| xlCubeSet | 3 | 集合 |


#### XlCubeFieldType 枚举

# [XlCubeFieldType 枚举​](#xlcubefieldtype-枚举)

指定 OLAP 字段是层次结构、集合还是度量字段。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlHierarchy | 1 | OLAP 字段是层次结构。 |
| xlMeasure | 2 | OLAP 字段是度量。 |
| xlSet | 3 | OLAP 字段是集合。 |


#### XlCutCopyMode 枚举

# [XlCutCopyMode 枚举​](#xlcutcopymode-枚举)

指定状态为复制模式还是剪切模式。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlCopy | 1 | 复制模式 |
| xlCut | 2 | 剪切模式 |


#### XlDVAlertStyle 枚举

# [XlDVAlertStyle 枚举​](#xldvalertstyle-枚举)

指定验证过程中显示的消息框所用的图标。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlValidAlertInformation | 3 | 信息图标。 |
| xlValidAlertStop | 1 | 停止图标。 |
| xlValidAlertWarning | 2 | 警告图标。 |


#### XlDVType 枚举

# [XlDVType 枚举​](#xldvtype-枚举)

指定要对值进行的有效性测试类型。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlValidateCustom | 7 | 使用任意公式验证数据有效性。 |
| xlValidateDate | 4 | 日期值。 |
| xlValidateDecimal | 2 | 数值。 |
| xlValidateInputOnly | 0 | 仅在用户更改值时进行验证。 |
| xlValidateList | 3 | 值必须存在于指定列表中。 |
| xlValidateTextLength | 6 | 文本长度。 |
| xlValidateTime | 5 | 时间值。 |
| xlValidateWholeNumber | 1 | 全部数值。 |


#### XlDataBarAxisPosition 枚举

# [XlDataBarAxisPosition 枚举​](#xldatabaraxisposition-枚举)

指定条件格式为数据条的单元格区域的坐标轴位置。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlDataBarAxisAutomatic | 0 | 在基于区域中最小负值对最大正值之比的可变位置处显示坐标轴。正值按从左至右方向显示。负值按从右至左方向显示。当所有值全为正或全为负时，将不显示坐标轴。 |
| xlDataBarAxisMidpoint | 1 | 不管区域中具有什么样的一组值，总是在单元格的中点显示坐标轴。正值按从左至右方向显示。负值按从右至左方向显示。 |
| xlDataBarAxisNone | 2 | 不显示任何坐标轴，正值和负值都按从左至右方向显示。 |


#### XlDataBarBorderType 枚举

# [XlDataBarBorderType 枚举​](#xldatabarbordertype-枚举)

指定数据条的边框。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlDataBarBorderNone | 0 | 数据条无边框。 |
| xlDataBarBorderSolid | 1 | 数据条有实心边框。 |


#### XlDataBarFillType 枚举

# [XlDataBarFillType 枚举​](#xldatabarfilltype-枚举)

指定如何对数据条填充颜色。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlDataBarFillGradient | 1 | 对数据条填充渐变色。 |
| xlDataBarFillSolid | 0 | 对数据条填充纯色。 |


#### XlDataBarNegativeColorType 枚举

# [XlDataBarNegativeColorType 枚举​](#xldatabarnegativecolortype-枚举)

指定是否使用与正数据条相同的边框和填充色。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlDataBarColor | 0 | 使用**“负值和坐标轴设置”**对话框中指定的颜色，或者使用由NegativeBarFormat对象的ColorType和BorderColorType属性指定的颜色。 |
| xlDataBarSameAsPositive | 1 | 使用与正数据条相同的颜色。 |


#### XlDataLabelSeparator 枚举

# [XlDataLabelSeparator 枚举​](#xldatalabelseparator-枚举)

指定用于数据标签的分隔符。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlDataLabelSeparatorDefault | 1 | ET 选择分隔符。 |


#### XlDataSeriesDate 枚举

# [XlDataSeriesDate 枚举​](#xldataseriesdate-枚举)

指定要应用于数据系列的日期的类型。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlDay | 1 | 日 |
| xlMonth | 3 | 月 |
| xlWeekday | 2 | 工作日 |
| xlYear | 4 | 年 |


#### XlDataSeriesType 枚举

# [XlDataSeriesType 枚举​](#xldataseriestype-枚举)

指定要创建的数据系列。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlAutoFill | 4 | 按照“自动填充”设置对系列进行填充。 |
| xlChronological | 3 | 用数据值进行填充。 |
| xlDataSeriesLinear | -4132 | 扩展值，假定一个加法级数（例如，“1, 2”被扩展为“3, 4, 5”）。 |
| xlGrowth | 2 | 扩展值，假定一个乘法级数（例如，“1, 2”被扩展为“4, 8, 16”）。 |


#### XlDeleteShiftDirection 枚举

# [XlDeleteShiftDirection 枚举​](#xldeleteshiftdirection-枚举)

指定如何移动单元格来替换删除的单元格。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlShiftToLeft | -4159 | 单元格向左移动。 |
| xlShiftUp | -4162 | 单元格向上移动。 |


#### XlDirection 枚举

# [XlDirection 枚举​](#xldirection-枚举)

指定移动的方向。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlDown | -4121 | 向下。 |
| xlToLeft | -4159 | 向左。 |
| xlToRight | -4161 | 向右。 |
| xlUp | -4162 | 向上。 |


#### XlDisplayDrawingObjects 枚举

# [XlDisplayDrawingObjects 枚举​](#xldisplaydrawingobjects-枚举)

指定形状的显示方式。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlDisplayShapes | -4104 | 显示所有形状。 |
| xlHide | 3 | 隐藏所有形状。 |
| xlPlaceholders | 2 | 仅显示占位符。 |


#### XlDisplayUnit 枚举

# [XlDisplayUnit 枚举​](#xldisplayunit-枚举)

指定坐标轴的显示单位标签。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlHundredMillions | -8 | 亿。 |
| xlHundreds | -2 | 百。 |
| xlHundredThousands | -5 | 十万。 |
| xlMillionMillions | -10 | 万亿。 |
| xlMillions | -6 | 百万。 |
| xlTenMillions | -7 | 千万。 |
| xlTenThousands | -4 | 万 |
| xlThousandMillions | -9 | 亿。 |
| xlThousands | -3 | 千。 |


#### XlDupeUnique 枚举

# [XlDupeUnique 枚举​](#xldupeunique-枚举)

指定应显示重复值还是唯一值。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlDuplicate | 1 | 显示重复值。 |
| xlUnique | 0 | 显示唯一值。 |


#### XlDynamicFilterCriteria 枚举

# [XlDynamicFilterCriteria 枚举​](#xldynamicfiltercriteria-枚举)

指定筛选条件。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlFilterAboveAverage | 33 | 筛选所有高于平均值的值。 |
| xlFilterAllDatesInPeriodApril | 24 | 筛选所有四月的日期。 |
| xlFilterAllDatesInPeriodAugust | 28 | 筛选所有八月的日期。 |
| xlFilterAllDatesInPeriodDecember | 32 | 筛选所有十二月的日期。 |
| xlFilterAllDatesInPeriodFebruray | 22 | 筛选所有二月的日期。 |
| xlFilterAllDatesInPeriodJanuary | 21 | 筛选所有一月的日期。 |
| xlFilterAllDatesInPeriodJuly | 27 | 筛选所有七月的日期。 |
| xlFilterAllDatesInPeriodJune | 26 | 筛选所有六月的日期。 |
| xlFilterAllDatesInPeriodMarch | 23 | 筛选所有三月的日期。 |
| xlFilterAllDatesInPeriodMay | 25 | 筛选所有五月的日期。 |
| xlFilterAllDatesInPeriodNovember | 31 | 筛选所有十一月的日期。 |
| xlFilterAllDatesInPeriodOctober | 30 | 筛选所有十月的日期。 |
| xlFilterAllDatesInPeriodQuarter1 | 17 | 筛选所有第一季度的日期。 |
| xlFilterAllDatesInPeriodQuarter2 | 18 | 筛选所有第二季度的日期。 |
| xlFilterAllDatesInPeriodQuarter3 | 19 | 筛选所有第三季度的日期。 |
| xlFilterAllDatesInPeriodQuarter4 | 20 | 筛选所有第四季度的日期。 |
| xlFilterAllDatesInPeriodSeptember | 29 | 筛选所有九月的日期。 |
| xlFilterBelowAverage | 34 | 筛选所有低于平均值的值。 |
| xlFilterLastMonth | 8 | 筛选所有与上月相关的值。 |
| xlFilterLastQuarter | 11 | 筛选所有与上一季度相关的值。 |
| xlFilterLastWeek | 5 | 筛选所有与上周相关的值。 |
| xlFilterLastYear | 14 | 筛选所有与去年相关的值。 |
| xlFilterNextMonth | 9 | 筛选所有与下月相关的值。 |
| xlFilterNextQuarter | 12 | 筛选所有与下一季度相关的值。 |
| xlFilterNextWeek | 6 | 筛选所有与下周相关的值。 |
| xlFilterNextYear | 15 | 筛选所有与明年相关的值。 |
| xlFilterThisMonth | 7 | 筛选所有与本月相关的值。 |
| xlFilterThisQuarter | 10 | 筛选所有与本季度相关的值。 |
| xlFilterThisWeek | 4 | 筛选所有与本周相关的值。 |
| xlFilterThisYear | 13 | 筛选所有与今年相关的值。 |
| xlFilterToday | 1 | 筛选所有与今天相关的值。 |
| xlFilterTomorrow | 3 | 筛选所有与明天相关的值。 |
| xlFilterYearToDate | 16 | 筛选到今天为止一年的所有值。 |
| xlFilterYesterday | 2 | 筛选所有与昨天相关的值。 |


#### XlEditionFormat 枚举

# [XlEditionFormat 枚举​](#xleditionformat-枚举)

指定发布版本的格式。此枚举仅用于 Macintosh，因此不应使用。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlBIFF | 2 | 二进制交换文件格式。 |
| xlPICT | 1 | 图元文件图片结构 (.wmf)。 |
| xlRTF | 4 | RTF 格式 (.rtf)。 |
| xlVALU | 8 | VALU。 |


#### XlEditionOptionsOption 枚举

# [XlEditionOptionsOption 枚举​](#xleditionoptionsoption-枚举)

此枚举仅用于 Macintosh，因此不应使用。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlAutomaticUpdate | 4 | 自动更新。 |
| xlCancel | 1 | 取消。 |
| xlChangeAttributes | 6 | 更改属性。 |
| xlManualUpdate | 5 | 手动更新。 |
| xlOpenSource | 3 | 打开源。 |
| xlSelect | 3 | 选择。 |
| xlSendPublisher | 2 | 发送到 Microsoft Publisher。 |
| xlUpdateSubscriber | 2 | 更新订阅服务器。 |


#### XlEditionType 枚举

# [XlEditionType 枚举​](#xleditiontype-枚举)

指定要更改的版本类型。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlPublisher | 1 | 发布服务器 |
| xlSubscriber | 2 | 订阅服务器 |


#### XlEnableCancelKey 枚举

# [XlEnableCancelKey 枚举​](#xlenablecancelkey-枚举)

指定 WPS Office ET 2007 如何处理 Ctrl+Break（或 Esc、Command+Period）用户中断以用于运行程序。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlDisabled | 0 | 完全禁用“取消”键捕获功能。 |
| xlErrorHandler | 2 | 将中断作为错误发送给运行程序，由 On Error GoTo 语句设置的错误处理程序捕获。可捕获的错误代码为 18。 |
| xlInterrupt | 1 | 中断当前运行程序，用户可进行调试或结束程序的运行。 |


#### XlEnableSelection 枚举

# [XlEnableSelection 枚举​](#xlenableselection-枚举)

指定可在工作表中选择的内容。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlNoRestrictions | 0 | 可以选择任何内容。 |
| xlNoSelection | -4142 | 不能选择任何内容。 |
| xlUnlockedCells | 1 | 只能选择未锁定单元格。 |


#### XlErrorBarDirection 枚举

# [XlErrorBarDirection 枚举​](#xlerrorbardirection-枚举)

指定接收误差线的轴值。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlX | -4168 | 误差线平行于 Y 轴，长度为 X 轴值。 |
| xlY | 1 | 误差线平行于 X 轴，长度为 Y 轴值。 |


#### XlErrorChecks 枚举

# [XlErrorChecks 枚举​](#xlerrorchecks-枚举)

指定要从Errors集合检索的误差对象的类型。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlEmptyCellReferences | 7 | 单元格包含一个引用空单元格的公式。 |
| xlEvaluateToError | 1 | 单元格计算为错误值。 |
| xlInconsistentFormula | 4 | 单元格包含与区域不一致的公式。 |
| xlInconsistentListFormula | 9 | 单元格包含与列表不一致的公式。 |
| xlListDataValidation | 8 | 列表中的数据包含一个有效性错误。 |
| xlNumberAsText | 3 | 按文本输入的数字。 |
| xlOmittedCells | 5 | 忽略的单元格。 |
| xlTextDate | 2 | 按文本输入的日期。 |
| xlUnlockedFormulaCells | 6 | 解除锁定公式单元格。 |


#### XlFileAccess 枚举

# [XlFileAccess 枚举​](#xlfileaccess-枚举)

指定对象的新访问模式。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlReadOnly | 3 | 只读。 |
| xlReadWrite | 2 | 可读/写。 |


#### XlFileFormat 枚举

# [XlFileFormat 枚举​](#xlfileformat-枚举)

指定保存工作表时的文件格式。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlAddIn | 18 | ET 2007 加载项 |
| xlAddIn8 | 18 | ET 97-2003 加载项 |
| xlCSV | 6 | CSV |
| xlCSVMac | 22 | Macintosh CSV |
| xlCSVMSDOS | 24 | MSDOS CSV |
| xlCSVWindows | 23 | Windows CSV |
| xlCurrentPlatformText | -4158 | 当前平台文本 |
| xlDBF2 | 7 | DBF2 |
| xlDBF3 | 8 | DBF3 |
| xlDBF4 | 11 | DBF4 |
| xlDIF | 9 | DIF |
| xlExcel12 | 50 | ET 12 |
| xlExcel2 | 16 | ET 2 |
| xlExcel2FarEast | 27 | Excel2 FarEast |
| xlExcel3 | 29 | Excel3 |
| xlExcel4 | 33 | Excel4 |
| xlExcel4Workbook | 35 | Excel4 工作簿 |
| xlExcel5 | 39 | Excel5 |
| xlExcel7 | 39 | Excel7 |
| xlExcel8 | 56 | Excel8 |
| xlExcel9795 | 43 | Excel9795 |
| xlHtml | 44 | HTML 格式 |
| xlIntlAddIn | 26 | 国际加载项 |
| xlIntlMacro | 25 | 国际宏 |
| xlOpenDocumentSpreadsheet | 60 | OpenDocument 电子表格 |
| xlOpenXMLAddIn | 55 | 打开 XML 加载项 |
| xlOpenXMLTemplate | 54 | 打开 XML 模板 |
| xlOpenXMLTemplateMacroEnabled | 53 | 打开启用的 XML 模板宏 |
| xlOpenXMLWorkbook | 51 | 打开 XML 工作簿 |
| xlOpenXMLWorkbookMacroEnabled | 52 | 打开启用的 XML 工作簿宏 |
| xlSYLK | 2 | SYLK |
| xlTemplate | 17 | 模板 |
| xlTemplate8 | 17 | 模板 8 |
| xlTextMac | 19 | Macintosh 文本 |
| xlTextMSDOS | 21 | MSDOS 文本 |
| xlTextPrinter | 36 | 打印机文本 |
| xlTextWindows | 20 | Windows 文本 |
| xlUnicodeText | 42 | Unicode 文本 |
| xlWebArchive | 45 | Web 档案 |
| xlWJ2WD1 | 14 | WJ2WD1 |
| xlWJ3 | 40 | WJ3 |
| xlWJ3FJ3 | 41 | WJ3FJ3 |
| xlWK1 | 5 | WK1 |
| xlWK1ALL | 31 | WK1ALL |
| xlWK1FMT | 30 | WK1FMT |
| xlWK3 | 15 | WK3 |
| xlWK3FM3 | 32 | WK3FM3 |
| xlWK4 | 38 | WK4 |
| xlWKS | 4 | 工作表 |
| xlWorkbookDefault | 51 | 默认工作簿 |
| xlWorkbookNormal | -4143 | 常规工作簿 |
| xlWorks2FarEast | 28 | Works2 FarEast |
| xlWQ1 | 34 | WQ1 |
| xlXMLSpreadsheet | 46 | XML 电子表格 |
| 102 | 102 | ofd |
| 103 | 103 | pdf |
| 65521 | 65521 | et |
| 65522 | 65522 | ett |
| 65523 | 65523 | uof |
| 65525 | 65525 | uos |


#### XlFileValidationPivotMode 枚举

# [XlFileValidationPivotMode 枚举​](#xlfilevalidationpivotmode-枚举)

指定如何验证数据透视表的数据缓存。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlFileValidationPivotDefault | 0 | 验证PivotOptions注册表设置指定的数据缓存的内容（默认）。 |
| xlFileValidationPivotRun | 1 | 验证所有数据缓存的内容，而不考虑注册表设置。 |
| xlFileValidationPivotSkip | 2 | 不验证数据缓存的内容。 |


#### XlFillWith 枚举

# [XlFillWith 枚举​](#xlfillwith-枚举)

指定如何复制区域。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlFillWithAll | -4104 | 复制内容和格式。 |
| xlFillWithContents | 2 | 仅复制内容。 |
| xlFillWithFormats | -4122 | 仅复制格式。 |


#### XlFilterAction 枚举

# [XlFilterAction 枚举​](#xlfilteraction-枚举)

指定在筛选操作过程中是复制数据还是保留不动。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlFilterCopy | 2 | 将筛选出的数据复制到新位置。 |
| xlFilterInPlace | 1 | 保留数据不动。 |


#### XlFilterAllDatesInPeriod 枚举

# [XlFilterAllDatesInPeriod 枚举​](#xlfilteralldatesinperiod-枚举)

指定在指定时间段内的日期筛选方式。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlFilterAllDatesInPeriodDay | 2 | 在所有日期内筛选指定日期。 |
| xlFilterAllDatesInPeriodHour | 3 | 在所有日期中筛选指定小时。 |
| xlFilterAllDatesInPeriodMinute | 4 | 在所有日期中筛选指定分钟。 |
| xlFilterAllDatesInPeriodMonth | 1 | 在所有日期中筛选指定月。 |
| xlFilterAllDatesInPeriodSecond | 5 | 在所有日期中筛选指定秒。 |
| xlFilterAllDatesInPeriodYear | 0 | 在所有日期内筛选指定年。 |


#### XlFindLookIn 枚举

# [XlFindLookIn 枚举​](#xlfindlookin-枚举)

指定要搜索的数据的类型。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlComments | -4144 | 批注。 |
| xlFormulas | -4123 | 公式。 |
| xlValues | -4163 | 值。 |


#### XlFixedFormatQuality 枚举

# [XlFixedFormatQuality 枚举​](#xlfixedformatquality-枚举)

指定以不同固定格式保存的电子表格的质量。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlQualityMinimum | 1 | 最低质量 |
| xlQualityStandard | 0 | 标准质量 |


#### XlFixedFormatType 枚举

# [XlFixedFormatType 枚举​](#xlfixedformattype-枚举)

指定文件格式的类型。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlTypePDF | 0 | "PDF" ― 可移植文档格式文件 (.pdf)。 |
| xlTypeXPS | 1 | "XPS" ― XPS 文档 (.xps)。 |


#### XlFormControl 枚举

# [XlFormControl 枚举​](#xlformcontrol-枚举)

指定表单控件的类型。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlButtonControl | 0 | 按钮。 |
| xlCheckBox | 1 | 复选框。 |
| xlDropDown | 2 | 组合框。 |
| xlEditBox | 3 | 文本框。 |
| xlGroupBox | 4 | 分组框。 |
| xlLabel | 5 | 标签。 |
| xlListBox | 6 | 列表框。 |
| xlOptionButton | 7 | 选项按钮。 |
| xlScrollBar | 8 | 滚动条。 |
| xlSpinner | 9 | 微调按钮。 |


#### XlFormatConditionOperator 枚举

# [XlFormatConditionOperator 枚举​](#xlformatconditionoperator-枚举)

指定运算符，用于比较公式与单元格中的值，或者比较两个公式（适用于xlBetween和xlNotBetween）。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlBetween | 1 | 介于。只在提供了两个公式的情况下才能使用。 |
| xlEqual | 3 | 等于。 |
| xlGreater | 5 | 大于。 |
| xlGreaterEqual | 7 | 大于或等于。 |
| xlLess | 6 | 小于。 |
| xlLessEqual | 8 | 小于或等于。 |
| xlNotBetween | 2 | 不介于。只在提供了两个公式的情况下才能使用。 |
| xlNotEqual | 4 | 不等于。 |


#### XlFormatConditionType 枚举

# [XlFormatConditionType 枚举​](#xlformatconditiontype-枚举)

指定条件格式是基于单元格值还是基于表达式。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlAboveAverageCondition | 12 | 高于平均值条件 |
| xlBlanksCondition | 10 | 空值条件 |
| xlCellValue | 1 | 单元格值 |
| xlColorScale | 3 | 色阶 |
| xlDatabar | 4 | 数据条 |
| xlErrorsCondition | 16 | 错误条件 |
| xlExpression | 2 | 表达式 |
| XlIconSet | 6 | 图标集 |
| xlNoBlanksCondition | 13 | 无空值条件 |
| xlNoErrorsCondition | 17 | 无错误条件 |
| xlTextString | 9 | 文本字符串 |
| xlTimePeriod | 11 | 时间段 |
| xlTop10 | 5 | 前 10 个值 |
| xlUniqueValues | 8 | 唯一值 |


#### XlFormatFilterTypes 枚举

# [XlFormatFilterTypes 枚举​](#xlformatfiltertypes-枚举)

指定格式筛选的类型。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| FilterBottom | 0 | 下筛选。 |
| FilterBottomPercent | 2 | 下百分比筛选。 |
| FilterTop | 1 | 上筛选。 |
| FilterTopPercent | 3 | 上百分比筛选。 |


#### XlFormulaLabel 枚举

# [XlFormulaLabel 枚举​](#xlformulalabel-枚举)

为指定区域指定公式标签类型。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlColumnLabels | 2 | 仅列标签。 |
| xlMixedLabels | 3 | 行标签和列标签。 |
| xlNoLabels | -4142 | 无标签。 |
| xlRowLabels | 1 | 仅行标签。 |


#### XlGenerateTableRefs 枚举

# [XLGenerateTableRefs 枚举​](#xlgeneratetablerefs-枚举)

指定表引用的类型。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlA1TableRefs | 0 | A1 表引用。 |
| xlTableNames | 1 | 表名称。 |


#### XlGradientFillType 枚举

# [XlGradientFillType 枚举​](#xlgradientfilltype-枚举)

指定gradient fill的类型。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| GradientFillLinear | 0 | 渐变以直线填充。 |
| GradientFillPath | 1 | 渐变以非线性或曲线路径填充。 |


#### XlHebrewModes 枚举

# [XlHebrewModes 枚举​](#xlhebrewmodes-枚举)

指定希伯来语拼写检查器的模式。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlHebrewFullScript | 0 | 在书写不带音调符号的文字时，希伯来语协会 (Hebrew Language Academy) 要求使用传统字符类型。 |
| xlHebrewMixedAuthorizedScript | 3 | 希伯来传统字符。 |
| xlHebrewMixedScript | 2 | 在这种模式下，拼写检查器接受所有识别为希伯来语的单词，包括以完整文字、部分文字或拼写检查器可识别的非常规拼写变体书写的单词。 |
| xlHebrewPartialScript | 1 | 在这种模式下，拼写检查器接受以完整文字和部分文字书写的单词。由于以完整文字和部分文字书写的某些单词的拼写未经核准，因此将对这些单词进行标记。 |


#### XlHighlightChangesTime 枚举

# [XlHighlightChangesTime 枚举​](#xlhighlightchangestime-枚举)

指定共享工作簿中显示的一组更改。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlAllChanges | 2 | 显示所有更改。 |
| xlNotYetReviewed | 3 | 仅显示还未审阅的更改。 |
| xlSinceMyLastSave | 1 | 显示上次保存后最后一个用户进行的更改。 |


#### XlHtmlType 枚举

# [XlHtmlType 枚举​](#xlhtmltype-枚举)

指定将指定项保存到网页时 ET 生成的 HTML 的类型，并指定该项是静态还是交互式。

| 注释 |
| --- |
| 除xlHtmlStatic之外的所有XlHtmlType枚举均已被弃用。 |

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlHtmlCalc | 1 | 使用电子表格组件。已弃用。 |
| xlHtmlChart | 3 | 使用图表组件。已弃用。 |
| xlHtmlList | 2 | 使用数据透视表组件。已弃用。 |
| xlHtmlStatic | 0 | 使用静态（非交互式）HTML，仅用于查看。 |


#### XlIMEMode 枚举

# [XlIMEMode 枚举​](#xlimemode-枚举)

指定日语输入规则的说明。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlIMEModeAlpha | 8 | 半角字母数字。 |
| xlIMEModeAlphaFull | 7 | 全角字母数字。 |
| xlIMEModeDisable | 3 | 禁用。 |
| xlIMEModeHangul | 10 | 朝鲜文。 |
| xlIMEModeHangulFull | 9 | 全角朝鲜文。 |
| xlIMEModeHiragana | 4 | 平假名。 |
| xlIMEModeKatakana | 5 | 片假名。 |
| xlIMEModeKatakanaHalf | 6 | 半角片假名。 |
| xlIMEModeNoControl | 0 | 无控制。 |
| xlIMEModeOff | 2 | 关闭（英文模式）。 |
| xlIMEModeOn | 1 | 模式打开。 |


#### XlIcon 枚举

# [XlIcon 枚举​](#xlicon-枚举)

指定图标集条件格式规则中某个条件的图标。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlIcon0Bars | 37 | 没有填充栏的信号指示器 |
| xlIcon0FilledBoxes | 52 | 无填充框 |
| xlIcon1Bar | 38 | 具有一个填充栏的信号指示器 |
| xlIcon1FilledBox | 51 | 一个填充框 |
| xlIcon2Bars | 39 | 具有两个填充栏的信号指示器 |
| xlIcon2FilledBoxes | 50 | 两个填充框 |
| xlIcon3Bars | 40 | 具有三个填充栏的信号指示器 |
| xlIcon3FilledBoxes | 49 | 三个填充框 |
| xlIcon4Bars | 41 | 具有四个填充栏的信号指示器 |
| xlIcon4FilledBoxes | 48 | 四个填充框 |
| xlIconBlackCircle | 32 | 黑色圆 |
| xlIconBlackCircleWithBorder | 13 | 黑色圆，带边框 |
| xlIconCircleWithOneWhiteQuarter | 33 | 四分之一为白色的圆 |
| xlIconCircleWithThreeWhiteQuarters | 35 | 四分之三为白色的圆 |
| xlIconCircleWithTwoWhiteQuarters | 34 | 四分之二为白色的圆 |
| xlIconGoldStar | 42 | 金色星形 |
| xlIconGrayCircle | 31 | 灰色圆 |
| xlIconGrayDownArrow | 6 | 灰色下箭头 |
| xlIconGrayDownInclineArrow | 28 | 灰色下斜箭头 |
| xlIconGraySideArrow | 5 | 灰色侧箭头 |
| xlIconGrayUpArrow | 4 | 灰色上箭头 |
| xlIconGrayUpInclineArrow | 27 | 灰色上斜箭头 |
| xlIconGreenCheck | 22 | 绿色复选符号 |
| xlIconGreenCheckSymbol | 19 | 绿色复选符号 |
| xlIconGreenCircle | 10 | 绿色圆 |
| xlIconGreenFlag | 7 | 绿旗 |
| xlIconGreenTrafficLight | 14 | 绿色交通灯 |
| xlIconGreenUpArrow | 1 | 绿色上箭头 |
| xlIconGreenUpTriangle | 45 | 绿色正三角形 |
| xlIconHalfGoldStar | 43 | 半金色星形 |
| xlIconNoCellIcon | -1 | 无单元格图标 |
| xlIconPinkCircle | 30 | 粉红色圆 |
| xlIconRedCircle | 29 | 红色圆 |
| xlIconRedCircleWithBorder | 12 | 红色圆，带边框 |
| xlIconRedCross | 24 | 红色十字 |
| xlIconRedCrossSymbol | 21 | 红色十字形符号 |
| xlIconRedDiamond | 18 | 红色菱形 |
| xlIconRedDownArrow | 3 | 红色下箭头 |
| xlIconRedDownTriangle | 47 | 红色倒三角形 |
| xlIconRedFlag | 9 | 红旗 |
| xlIconRedTrafficLight | 16 | 红色交通灯 |
| xlIconSilverStar | 44 | 银色星形 |
| xlIconWhiteCircleAllWhiteQuarters | 36 | 纯白圆 |
| xlIconYellowCircle | 11 | 黄色圆 |
| xlIconYellowDash | 46 | 黄色虚线三角形 |
| xlIconYellowDownInclineArrow | 26 | 黄色下斜箭头 |
| xlIconYellowExclamation | 23 | 黄色感叹号 |
| xlIconYellowExclamationSymbol | 20 | 黄色感叹号 |
| xlIconYellowFlag | 8 | 黄旗 |
| xlIconYellowSideArrow | 2 | 黄色侧箭头 |
| xlIconYellowTrafficLight | 15 | 黄色交通灯 |
| xlIconYellowTriangle | 17 | 黄色三角形 |
| xlIconYellowUpInclineArrow | 25 | 黄色上斜箭头 |


#### XlIconSet 枚举

# [XlIconSet 枚举​](#xliconset-枚举)

指定图标集的类型。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xl3Arrows | 1 | 三向箭头 |
| xl3ArrowsGray | 2 | 灰色三向箭头 |
| xl3Flags | 3 | 三色旗 |
| xl3Signs | 6 | 三标志 |
| xl3Symbols | 7 | 三个符号 |
| xl3TrafficLights1 | 4 | 三色交通灯 1 |
| xl3TrafficLights2 | 5 | 三色交通灯 2 |
| xl4Arrows | 8 | 四向箭头 |
| xl4ArrowsGray | 9 | 灰色四向箭头 |
| xl4CRV | 11 | 4 CRV |
| xl4RedToBlack | 10 | 四个圆红－黑渐变 |
| xl4TrafficLights | 12 | 四色交通灯 |
| xl5Arrows | 13 | 五向箭头 |
| xl5ArrowsGray | 14 | 灰色五向箭头 |
| xl5CRV | 15 | 5 CRV |
| xl5Quarters | 16 | 五象限图 |


#### XlImportDataAs 枚举

# [XlImportDataAs 枚举​](#xlimportdataas-枚举)

指定从数据库返回数据的格式。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlPivotTableReport | 1 | 以数据透视表的形式返回数据。 |
| xlQueryTable | 0 | 以查询表的形式返回数据。 |


#### XlInsertFormatOrigin 枚举

# [XlInsertFormatOrigin 枚举​](#xlinsertformatorigin-枚举)

指定从何处复制插入行的格式。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlFormatFromLeftOrAbove | 0 | 从上方和/或左侧单元格复制格式。 |
| xlFormatFromRightOrBelow | 1 | 从下方和/或右侧单元格复制格式。 |


#### XlInsertShiftDirection 枚举

# [XlInsertShiftDirection 枚举​](#xlinsertshiftdirection-枚举)

指定插入时单元格的移动方向。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlShiftDown | -4121 | 向下移动单元格。 |
| xlShiftToRight | -4161 | 向右移动单元格。 |


#### XlLayoutFormType 枚举

# [XlLayoutFormType 枚举​](#xllayoutformtype-枚举)

为指定的数据透视表项指定显示方式，即以表格式还是以大纲格式显示。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlOutline | 1 | LayoutSubtotalLocation属性指定分类汇总在数据透视表中出现的位置。 |
| xlTabular | 0 | 默认值。 |


#### XlLayoutRowType 枚举

# [XlLayoutRowType 枚举​](#xllayoutrowtype-枚举)

指定版式行的类型。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlCompactRow | 0 | 压缩行 |
| xlOutlineRow | 2 | 大纲行 |
| xlTabularRow | 1 | 表格行 |


#### XlLineStyle 枚举

# [XlLineStyle 枚举​](#xllinestyle-枚举)

指定边框的线条样式。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlContinuous | 1 | 实线。 |
| xlDash | -4115 | 虚线。 |
| xlDashDot | 4 | 点划相间线。 |
| xlDashDotDot | 5 | 划线后跟两个点。 |
| xlDot | -4118 | 点式线。 |
| xlDouble | -4119 | 双线。 |
| xlLineStyleNone | -4142 | 无线条。 |
| xlSlantDashDot | 13 | 倾斜的划线。 |


#### XlLink 枚举

# [XlLink 枚举​](#xllink-枚举)

指定链接的类型。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlExcelLinks | 1 | 到 ET 工作表的链接。 |
| xlOLELinks | 2 | 到 OLE 源的链接。 |
| xlPublishers | 5 | 仅用于 Macintosh。 |
| xlSubscribers | 6 | 仅用于 Macintosh。 |


#### XlLinkInfo 枚举

# [XlLinkInfo 枚举​](#xllinkinfo-枚举)

指定链接将返回的信息的类型。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlEditionDate | 2 | 仅应用于 Macintosh 操作系统中的版本。 |
| xlLinkInfoStatus | 3 | 返回链接状态。 |
| xlUpdateState | 1 | 指定链接是自动更新还是手动更新。 |


#### XlLinkInfoType 枚举

# [XlLinkInfoType 枚举​](#xllinkinfotype-枚举)

指定链接的类型。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlLinkInfoOLELinks | 2 | OLE 或 DDE 服务器 |
| xlLinkInfoPublishers | 5 | 发布服务器 |
| xlLinkInfoSubscribers | 6 | 订阅服务器 |


#### XlLinkStatus 枚举

# [XlLinkStatus 枚举​](#xllinkstatus-枚举)

指定链接的状态。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlLinkStatusCopiedValues | 10 | 复制的值。 |
| xlLinkStatusIndeterminate | 5 | 不能确定状态。 |
| xlLinkStatusInvalidName | 7 | 名称无效。 |
| xlLinkStatusMissingFile | 1 | 文件丢失。 |
| xlLinkStatusMissingSheet | 2 | 工作表丢失。 |
| xlLinkStatusNotStarted | 6 | 未启动。 |
| xlLinkStatusOK | 0 | 无错误。 |
| xlLinkStatusOld | 3 | 状态可能过期。 |
| xlLinkStatusSourceNotCalculated | 4 | 尚未计算。 |
| xlLinkStatusSourceNotOpen | 8 | 未打开。 |
| xlLinkStatusSourceOpen | 9 | 源文档打开。 |


#### XlLinkType 枚举

# [XlLinkType 枚举​](#xllinktype-枚举)

指定链接的类型。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlLinkTypeExcelLinks | 1 | 到 ET 源的链接。 |
| xlLinkTypeOLELinks | 2 | 到 OLE 源的链接。 |


#### XlListConflict 枚举

# [XlListConflict 枚举​](#xllistconflict-枚举)

指定冲突（用 ET 工作表中的列表的更改更新 Microsoft SharePoint Foundation 网站上的列表时）解决方法选项。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlListConflictDialog | 0 | 显示一个对话框，允许用户选择解决冲突的方式。 |
| xlListConflictDiscardAllConflicts | 2 | 接受存储在 SharePoint 网站上的数据版本。 |
| xlListConflictError | 3 | 如果发生冲突，则引发一个错误。 |
| xlListConflictRetryAllConflicts | 1 | 覆盖存储在 SharePoint 网站上的数据版本。 |


#### XlListDataType 枚举

# [XlListDataType 枚举​](#xllistdatatype-枚举)

指定连接到 Microsoft SharePoint Foundation 网站的列表列的数据类型。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlListDataTypeCheckbox | 9 | 复选框。 |
| xlListDataTypeChoice | 6 | 单一选择字段。 |
| xlListDataTypeChoiceMulti | 7 | 多个选择字段。 |
| xlListDataTypeCounter | 11 | 计数器。 |
| xlListDataTypeCurrency | 4 | 货币。 |
| xlListDataTypeDateTime | 5 | 日期/时间。 |
| xlListDataTypeHyperLink | 10 | 超链接。 |
| xlListDataTypeListLookup | 8 | “查阅”列表。 |
| xlListDataTypeMultiLineRichText | 12 | 多行 RTF 格式。 |
| xlListDataTypeMultiLineText | 2 | 多行纯文本。 |
| xlListDataTypeNone | 0 | 未指定类型。 |
| xlListDataTypeNumber | 3 | 数字。 |
| xlListDataTypeText | 1 | 纯文本。 |


#### XlListObjectSourceType 枚举

# [XlListObjectSourceType 枚举​](#xllistobjectsourcetype-枚举)

指定列表的当前源。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlSrcExternal | 0 | 外部数据源（Microsoft SharePoint Foundation 网站）。 |
| xlSrcQuery | 3 | 查询 |
| xlSrcRange | 1 | 区域 |
| xlSrcXml | 2 | XML |


#### XlLocationInTable 枚举

# [XlLocationInTable 枚举​](#xllocationintable-枚举)

指定数据透视表中包含区域左上角的部分。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlColumnHeader | -4110 | 列标题 |
| xlColumnItem | 5 | 列数据项 |
| xlDataHeader | 3 | 数据标题 |
| xlDataItem | 7 | 数据项 |
| xlPageHeader | 2 | 页面页眉 |
| xlPageItem | 6 | 页面项 |
| xlRowHeader | -4153 | 行标题 |
| xlRowItem | 4 | 行数据项 |
| xlTableBody | 8 | 表正文 |


#### XlLookAt 枚举

# [XlLookAt 枚举​](#xllookat-枚举)

指定是匹配全部搜索文本还是匹配任一部分搜索文本。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlPart | 2 | 匹配任一部分搜索文本。 |
| xlWhole | 1 | 匹配全部搜索文本。 |


#### XlLookFor 枚举

# [XlLookFor 枚举​](#xllookfor-枚举)

指定要搜索的内容。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| LookForBlanks | 0 | 空值 |
| LookForErrors | 1 | 错误 |
| LookForFormulas | 2 | 公式 |


#### XlMSApplication 枚举

# [XlMSApplication 枚举​](#xlmsapplication-枚举)

指定一个 Microsoft 应用程序。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlMicrosoftAccess | 4 | WPS Office Access |
| xlMicrosoftFoxPro | 5 | Microsoft FoxPro |
| xlMicrosoftMail | 3 | WPS Office Outlook |
| xlMicrosoftPowerPoint | 2 | WPS Office PowerPoint |
| xlMicrosoftProject | 6 | WPS Office Project |
| xlMicrosoftSchedulePlus | 7 | Microsoft Schedule Plus |
| xlMicrosoftWord | 1 | WPS Office Word |


#### XlMailSystem 枚举

# [XlMailSystem 枚举​](#xlmailsystem-枚举)

指定安装在主机上的邮件系统。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlMAPI | 1 | 符合 MAPI 的系统 |
| xlNoMailSystem | 0 | 无邮件系统 |
| xlPowerTalk | 2 | PowerTalk 邮件系统 |


#### XlMeasurementUnits 枚举

# [XlMeasurementUnits 枚举​](#xlmeasurementunits-枚举)

指定度量单位。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlCentimeters | 1 | 厘米 |
| xlInches | 0 | 英寸 |
| xlMillimeters | 2 | 毫米 |


#### XlMouseButton 枚举

# [XlMouseButton 枚举​](#xlmousebutton-枚举)

指定按下了哪个鼠标按钮。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlNoButton | 0 | 没有按任何按钮。 |
| xlPrimaryButton | 1 | 按下主按钮（通常为鼠标左按钮）。 |
| xlSecondaryButton | 2 | 按下辅按钮（通常为鼠标右按钮）。 |


#### XlMousePointer 枚举

# [XlMousePointer 枚举​](#xlmousepointer-枚举)

指定 ET 中鼠标指针的外观。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlDefault | -4143 | 默认指针。 |
| xlIBeam | 3 | I 形指针。 |
| xlNorthwestArrow | 1 | 西北向箭头指针。 |
| xlWait | 2 | 沙漏型指针。 |


#### XlOLEType 枚举

# [XlOLEType 枚举​](#xloletype-枚举)

指定 OLE 对象类型。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlOLEControl | 2 | ActiveX 控件 |
| xlOLEEmbed | 1 | 嵌入式 OLE 对象 |
| xlOLELink | 0 | 链接 OLE 对象 |


#### XlOLEVerb 枚举

# [XlOLEVerb 枚举​](#xloleverb-枚举)

指定使 OLE 对象服务器执行操作的动作。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlVerbOpen | 2 | 打开对象。 |
| xlVerbPrimary | 1 | 执行服务器的主要操作。 |


#### XlOartHorizontalOverflow 枚举

# [XlOartHorizontalOverflow 枚举​](#xloarthorizontaloverflow-枚举)

指定文本框的水平溢出设置。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlOartHorizontalOverflowClip | 1 | 隐藏水平方向溢出文本框的文本。 |
| xlOartHorizontalOverflowOverflow | 0 | 允许文本在水平方向溢出文本框。 |


#### XlOartVerticalOverflow 枚举

# [XlOartVerticalOverflow 枚举​](#xloartverticaloverflow-枚举)

指定文本框的垂直溢出设置。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlOartVerticalOverflowClip | 1 | 隐藏垂直方向溢出文本框的文本。 |
| xlOartVerticalOverflowEllipsis | 2 | 隐藏垂直方向溢出文本框的文本，并在可见文本的最后添加省略号 (...)。 |
| xlOartVerticalOverflowOverflow | 0 | 允许文本在垂直方向溢出文本框（根据文本对齐方式，可从上溢出、自下溢出，或者上下溢出）。 |


#### XlObjectSize 枚举

# [XlObjectSize 枚举​](#xlobjectsize-枚举)

指定图表为适应页面大小而进行缩放的方式。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlFitToPage | 2 | 尽可能大地打印图表，并保持如屏幕所示的该图表的高度对宽度的比例。 |
| xlFullPage | 3 | 按照与页面相适应的大小打印图表，并根据需要调整其高度对宽度的比例。 |
| xlScreenSize | 1 | 以屏幕显示大小打印图表。 |


#### XlOrder 枚举

# [XlOrder 枚举​](#xlorder-枚举)

指定单元格的处理次序。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlDownThenOver | 1 | 向下处理行，然后向右逐个处理页或页面字段。 |
| xlOverThenDown | 2 | 向右逐个处理页或页面字段，然后向下处理行。 |


#### XlOrientation 枚举

# [XlOrientation 枚举​](#xlorientation-枚举)

指定文字方向。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlDownward | -4170 | 文字向下排列。 |
| xlHorizontal | -4128 | 文字水平排列。 |
| xlUpward | -4171 | 文字向上排列。 |
| xlVertical | -4166 | 文字在单元格中向下居中排列。 |


#### XlPTSelectionMode 枚举

# [XlPTSelectionMode 枚举​](#xlptselectionmode-枚举)

指定在结构化选择过程中可以在数据透视表中选择的内容。这些常数可以进行组合以选择多个类型。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlBlanks | 4 | 空值 |
| xlButton | 15 | 按钮 |
| xlDataAndLabel | 0 | 数据和标签 |
| xlDataOnly | 2 | 数据 |
| xlFirstRow | 256 | 第一行 |
| xlLabelOnly | 1 | 标签 |
| xlOrigin | 3 | 原点 |


#### XlPageBreak 枚举

# [XlPageBreak 枚举​](#xlpagebreak-枚举)

指定工作表中的分页符位置。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlPageBreakAutomatic | -4105 | ET 自动添加分页符。 |
| xlPageBreakManual | -4135 | 手动插入分页符。 |
| xlPageBreakNone | -4142 | 工作表中不插入分页符。 |


#### XlPageBreakExtent 枚举

# [XlPageBreakExtent 枚举​](#xlpagebreakextent-枚举)

指定分页符是全屏应用还是仅应用在打印区域。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlPageBreakFull | 1 | 全屏。 |
| xlPageBreakPartial | 2 | 仅在打印区域内。 |


#### XlPageOrientation 枚举

# [XlPageOrientation 枚举​](#xlpageorientation-枚举)

指定打印工作表时的页面方向。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlLandscape | 2 | 横向模式。 |
| xlPortrait | 1 | 纵向模式。 |


#### XlPaperSize 枚举

# [XlPaperSize 枚举​](#xlpapersize-枚举)

指定纸张的大小。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlPaper10x14 | 16 | 10 英寸 x 14 英寸 |
| xlPaper11x17 | 17 | 11 英寸 x 17 英寸 |
| xlPaperA3 | 8 | A3（297 毫米 x 420 毫米） |
| xlPaperA4 | 9 | A4（210 毫米 x 297 毫米） |
| xlPaperA4Small | 10 | A4（小）（210 毫米 x 297 毫米） |
| xlPaperA5 | 11 | A5（148 毫米 x 210 毫米） |
| xlPaperB4 | 12 | B4（250 毫米 x 354 毫米） |
| xlPaperB5 | 13 | A5（148 毫米 x 210 毫米） |
| xlPaperCsheet | 24 | C 型纸 |
| xlPaperDsheet | 25 | D 型纸 |
| xlPaperEnvelope10 | 20 | 信封 #10（4-1/8 英寸 x 9-1/2 英寸） |
| xlPaperEnvelope11 | 21 | 信封 #11（4-1/2 英寸 x 10-3/8 英寸） |
| xlPaperEnvelope12 | 22 | 信封 #12（4-1/2 英寸 x 11 英寸） |
| xlPaperEnvelope14 | 23 | 信封 #14（5 英寸 x 11-1/2 英寸） |
| xlPaperEnvelope9 | 19 | 信封 #9（3-7/8 英寸 x 8-7/8 英寸） |
| xlPaperEnvelopeB4 | 33 | 信封 B4（250 毫米 x 353 毫米） |
| xlPaperEnvelopeB5 | 34 | 信封 B5（176 毫米 x 250 毫米） |
| xlPaperEnvelopeB6 | 35 | 信封 B6（176 毫米 x 125 毫米） |
| xlPaperEnvelopeC3 | 29 | 信封 C3（324 毫米 x 458 毫米） |
| xlPaperEnvelopeC4 | 30 | 信封 C4（229 毫米 x 324 毫米） |
| xlPaperEnvelopeC5 | 28 | 信封 C5（162 毫米 x 229 毫米） |
| xlPaperEnvelopeC6 | 31 | 信封 C6（114 毫米 x 162 毫米） |
| xlPaperEnvelopeC65 | 32 | 信封 C65（114 毫米 x 229 毫米） |
| xlPaperEnvelopeDL | 27 | 信封 DL（110 毫米 x 220 毫米） |
| xlPaperEnvelopeItaly | 36 | 信封（110 毫米 x 230 毫米） |
| xlPaperEnvelopeMonarch | 37 | 君主式信封（3-7/8 英寸 x 7-1/2 英寸） |
| xlPaperEnvelopePersonal | 38 | 信封（3-5/8 英寸 x 6-1/2 英寸） |
| xlPaperEsheet | 26 | E 型纸 |
| xlPaperExecutive | 7 | 行政公文纸（7-1/2 英寸 x 10-1/2 英寸） |
| xlPaperFanfoldLegalGerman | 41 | 德国法律文书用复写簿（8-1/2 英寸 x 13 英寸） |
| xlPaperFanfoldStdGerman | 40 | 德国法律文书用复写簿（8-1/2 英寸 x 13 英寸） |
| xlPaperFanfoldUS | 39 | 美国标准复写簿（14-7/8 英寸 x 11 英寸） |
| xlPaperFolio | 14 | 对开纸（8-1/2 英寸 x 13 英寸） |
| xlPaperLedger | 4 | 帐单（17 英寸 x 11 英寸） |
| xlPaperLegal | 5 | 法律纸（8-1/2 英寸 x 14 英寸） |
| xlPaperLetter | 1 | 信函（8-1/2 英寸 x 11 英寸） |
| xlPaperLetterSmall | 2 | 简式信纸（8-1/2 英寸 x 11 英寸） |
| xlPaperNote | 18 | 便笺（8-1/2 英寸 x 11 英寸） |
| xlPaperQuarto | 15 | 四开本（215 毫米 x 275 毫米） |
| xlPaperStatement | 6 | 报告单（5-1/2 英寸 x 8-1/2 英寸） |
| xlPaperTabloid | 3 | 文摘（11 英寸 x 17 英寸） |
| xlPaperUser | 256 | 用户自定义 |


#### XlParameterDataType 枚举

# [XlParameterDataType 枚举​](#xlparameterdatatype-枚举)

指定查询参数的数据类型。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlParamTypeBigInt | -5 | 大整数。 |
| xlParamTypeBinary | -2 | 二进制。 |
| xlParamTypeBit | -7 | 位。 |
| xlParamTypeChar | 1 | 字符串。 |
| xlParamTypeDate | 9 | 日期。 |
| xlParamTypeDecimal | 3 | 十进制。 |
| xlParamTypeDouble | 8 | 双精度型。 |
| xlParamTypeFloat | 6 | 浮点型。 |
| xlParamTypeInteger | 4 | 整数。 |
| xlParamTypeLongVarBinary | -4 | 长二进制。 |
| xlParamTypeLongVarChar | -1 | 长字符串。 |
| xlParamTypeNumeric | 2 | 数字。 |
| xlParamTypeReal | 7 | 实数。 |
| xlParamTypeSmallInt | 5 | 小整数。 |
| xlParamTypeTime | 10 | 时间。 |
| xlParamTypeTimestamp | 11 | 时间戳。 |
| xlParamTypeTinyInt | -6 | 微小整数。 |
| xlParamTypeUnknown | 0 | 类型未知。 |
| xlParamTypeVarBinary | -3 | 变长度二进制。 |
| xlParamTypeVarChar | 12 | 变长度字符串。 |
| xlParamTypeWChar | -8 | Unicode 字符串。 |


#### XlParameterType 枚举

# [XlParameterType 枚举​](#xlparametertype-枚举)

为指定的查询表指定确定参数值的方式。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlConstant | 1 | 使用Value参数指定的值。 |
| xlPrompt | 0 | 显示提示用户输入值的对话框。Value参数指定的是对话框中显示的文字。 |
| xlRange | 2 | 使用区域左上角单元格的值。Value参数指定的是一个Range对象。 |


#### XlPasteSpecialOperation 枚举

# [XlPasteSpecialOperation 枚举​](#xlpastespecialoperation-枚举)

指定工作表中目标单元格的数字数据的计算方式。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlPasteSpecialOperationAdd | 2 | 复制的数据与目标单元格中的值相加。 |
| xlPasteSpecialOperationDivide | 5 | 复制的数据除以目标单元格中的值。 |
| xlPasteSpecialOperationMultiply | 4 | 复制的数据乘以目标单元格中的值。 |
| xlPasteSpecialOperationNone | -4142 | 粘贴操作中不执行任何计算。 |
| xlPasteSpecialOperationSubtract | 3 | 复制的数据减去目标单元格中的值。 |


#### XlPasteType 枚举

# [XlPasteType 枚举​](#xlpastetype-枚举)

指定要粘贴的区域部分。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlPasteAll | -4104 | 粘贴全部内容。 |
| xlPasteAllExceptBorders | 7 | 粘贴除边框外的全部内容。 |
| xlPasteAllMergingConditionalFormats | 14 | 将粘贴所有内容，并且将合并条件格式。 |
| xlPasteAllUsingSourceTheme | 13 | 使用源主题粘贴全部内容。 |
| xlPasteColumnWidths | 8 | 粘贴复制的列宽。 |
| xlPasteComments | -4144 | 粘贴批注。 |
| xlPasteFormats | -4122 | 粘贴复制的源格式。 |
| xlPasteFormulas | -4123 | 粘贴公式。 |
| xlPasteFormulasAndNumberFormats | 11 | 粘贴公式和数字格式。 |
| xlPasteValidation | 6 | 粘贴有效性。 |
| xlPasteValues | -4163 | 粘贴值。 |
| xlPasteValuesAndNumberFormats | 12 | 粘贴值和数字格式。 |


#### XlPattern 枚举

# [XlPattern 枚举​](#xlpattern-枚举)

指定图表或内部对象的内部图案。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlPatternAutomatic | -4105 | ET 控制图案。 |
| xlPatternChecker | 9 | 棋盘。 |
| xlPatternCrissCross | 16 | 十字线。 |
| xlPatternDown | -4121 | 左上角到右下角的深色对角线。 |
| xlPatternGray16 | 17 | 16% 灰。 |
| xlPatternGray25 | -4124 | 25% 灰。 |
| xlPatternGray50 | -4125 | 50% 灰。 |
| xlPatternGray75 | -4126 | 75% 灰。 |
| xlPatternGray8 | 18 | 8% 灰。 |
| xlPatternGrid | 15 | 网格。 |
| xlPatternHorizontal | -4128 | 深色水平线。 |
| xlPatternLightDown | 13 | 左上角到右下角的浅色对角线。 |
| xlPatternLightHorizontal | 11 | 浅色水平线。 |
| xlPatternLightUp | 14 | 左下角到右上角的浅色对角线。 |
| xlPatternLightVertical | 12 | 浅色垂直条。 |
| xlPatternNone | -4142 | 无图案。 |
| xlPatternSemiGray75 | 10 | 75% 深色摩尔纹。 |
| xlPatternSolid | 1 | 纯色。 |
| xlPatternUp | -4162 | 左下角到右上角的深色对角线。 |
| xlPatternVertical | -4166 | 深色垂直条。 |


#### XlPhoneticAlignment 枚举

# [XlPhoneticAlignment 枚举​](#xlphoneticalignment-枚举)

指定拼音文字的对齐方式。用于Phonetic或Phonetics对象。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlPhoneticAlignCenter | 2 | 居中对齐 |
| xlPhoneticAlignDistributed | 3 | 分散对齐 |
| xlPhoneticAlignLeft | 1 | 左对齐 |
| xlPhoneticAlignNoControl | 0 | ET 控制对齐方式 |


#### XlPhoneticCharacterType 枚举

# [XlPhoneticCharacterType 枚举​](#xlphoneticcharactertype-枚举)

指定单元格中拼音文字的类型。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlHiragana | 2 | 平假名 |
| xlKatakana | 1 | 片假名 |
| xlKatakanaHalf | 0 | 半尺寸片假名 |
| xlNoConversion | 3 | 无转换 |


#### XlPictureAppearance 枚举

# [XlPictureAppearance 枚举​](#xlpictureappearance-枚举)

指定图片的复制方式。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlPrinter | 2 | 图片按其打印效果进行复制。 |
| xlScreen | 1 | 图片尽可能按其屏幕显示进行复制。 |


#### XlPictureConvertorType 枚举

# [XlPictureConvertorType 枚举​](#xlpictureconvertortype-枚举)

指定图形的转换方式。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlBMP | 1 | 与 Windows 2.0 版兼容的位图 |
| xlCGM | 7 | 计算机图形图元文件 |
| xlDRW | 4 | DRW |
| xlDXF | 5 | DXF |
| xlEPS | 8 | 封装的附录 |
| xlHGL | 6 | HGL |
| xlPCT | 13 | 位图图形（Apple PICT 格式） |
| xlPCX | 10 | PC 画笔位图图形 |
| xlPIC | 11 | PIC |
| xlPLT | 12 | PLT |
| xlTIF | 9 | 标记图像格式文件 |
| xlWMF | 2 | Windows 图元文件 |
| xlWPG | 3 | WordPerfect/DrawPerfect 图形 |


#### XlPivotCellType 枚举

# [XlPivotCellType 枚举​](#xlpivotcelltype-枚举)

指定单元格所对应的PivotTable实体。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlPivotCellBlankCell | 9 | 数据透视表中的结构空白单元格。 |
| xlPivotCellCustomSubtotal | 7 | 行或列区域中作为自定义分类汇总的单元格。 |
| xlPivotCellDataField | 4 | 数据字段标签（不是**“数据”**按钮）。 |
| xlPivotCellDataPivotField | 8 | **“数据”**按钮。 |
| xlPivotCellGrandTotal | 3 | 行或列区域中作为总计的单元格。 |
| xlPivotCellPageFieldItem | 6 | 用于显示页字段的选定项的单元格。 |
| xlPivotCellPivotField | 5 | 字段的按钮（不是**“数据”**按钮）。 |
| xlPivotCellPivotItem | 1 | 行或列区域中不是分类汇总、总计、自定义分类汇总或空行的单元格。 |
| xlPivotCellSubtotal | 2 | 行或列区域中作为分类汇总的单元格。 |
| xlPivotCellValue | 0 | 数据区域中的任一单元格（空行除外）。 |


#### XlPivotConditionScope 枚举

# [XlPivotConditionScope 枚举​](#xlpivotconditionscope-枚举)

此枚举指定用于从PivotTable对象中筛选值的条件格式。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlDataFieldScope | 2 | 基于指定字段中的数据。 |
| xlFieldsScope | 1 | 基于指定的字段。 |
| xlSelectionScope | 0 | 基于指定的选择条件。 |


#### XlPivotFieldCalculation 枚举

# [XlPivotFieldCalculation 枚举​](#xlpivotfieldcalculation-枚举)

指定在使用自定义计算时由数据透视字段执行的计算类型。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlDifferenceFrom | 2 | 与基本字段中基本项的值的差。 |
| xlIndex | 9 | 按 ((单元格中的值) x (总计)) / ((行总计) x (列总计)) 计算的数据。 |
| xlNoAdditionalCalculation | -4143 | 无计算。 |
| xlPercentDifferenceFrom | 4 | 与基本字段中基本项的值的差异百分比。 |
| xlPercentOf | 3 | 占基本字段中基本项的值的百分比。 |
| xlPercentOfColumn | 7 | 占列或系列总计的百分比。 |
| xlPercentOfParent | 12 | 指定的父基本字段的总计的百分比。 |
| xlPercentOfParentColumn | 11 | 父列的总计的百分比。 |
| xlPercentOfParentRow | 10 | 父行的总计的百分比。 |
| xlPercentOfRow | 6 | 占行或类别总计的百分比。 |
| xlPercentOfTotal | 8 | 占报表中所有数据或数据点总计的百分比。 |
| xlPercentRunningTotal | 13 | 指定基本字段的运行总计的百分比。 |
| xlRankAscending | 14 | 从最小到最大排名。 |
| xlRankDecending | 15 | 从最大到最小排名。 |
| xlRunningTotal | 5 | 以运行总和形式表示的基本字段中连续项的数据。 |


#### XlPivotFieldDataType 枚举

# [XlPivotFieldDataType 枚举​](#xlpivotfielddatatype-枚举)

指定PivotTable字段中数据的类型。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlDate | 2 | 包含一个日期。 |
| xlNumber | -4145 | 包含一个数字。 |
| xlText | -4158 | 包含文本。 |


#### XlPivotFieldRepeatLabels 枚举

# [XlPivotFieldRepeatLabels 枚举​](#xlpivotfieldrepeatlabels-枚举)

指定是否重复数据透视表中的所有字段项目标签。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlDoNotRepeatLabels | 1 | 不重复项目标签。 |
| xlRepeatLabels | 2 | 重复所有项目标签。 |


#### XlPivotFilterType 枚举

# [XlPivotFilterType 枚举​](#xlpivotfiltertype-枚举)

应用的筛选器的类型。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlBefore | 31 | 筛选早于指定日期的所有日期 |
| xlBeforeOrEqualTo | 32 | 筛选等于或早于指定日期的所有日期 |
| xlAfter | 33 | 筛选迟于指定日期的所有日期 |
| xlAfterOrEqualTo | 34 | 筛选等于或迟于指定日期的所有日期 |
| xlAllDatesInPeriodJanuary | 53 | 筛选一月的所有日期 |
| xlAllDatesInPeriodFebruary | 54 | 筛选二月的所有日期 |
| xlAllDatesInPeriodMarch | 55 | 筛选三月的所有日期 |
| xlAllDatesInPeriodApril | 56 | 筛选四月的所有日期 |
| xlAllDatesInPeriodMay | 57 | 筛选五月的所有日期 |
| xlAllDatesInPeriodJune | 58 | 筛选六月的所有日期 |
| xlAllDatesInPeriodJuly | 59 | 筛选七月的所有日期 |
| xlAllDatesInPeriodAugust | 60 | 筛选八月的所有日期 |
| xlAllDatesInPeriodSeptember | 61 | 筛选九月的所有日期 |
| xlAllDatesInPeriodOctober | 62 | 筛选十月的所有日期 |
| xlAllDatesInPeriodNovember | 63 | 筛选十一月的所有日期 |
| xlAllDatesInPeriodDecember | 64 | 筛选十二月的所有日期 |
| xlAllDatesInPeriodQuarter1 | 49 | 筛选第一季度中的所有日期 |
| xlAllDatesInPeriodQuarter2 | 50 | 筛选第二季度中的所有日期 |
| xlAllDatesInPeriodQuarter3 | 51 | 筛选第三季度中的所有日期 |
| xlAllDatesInPeriodQuarter4 | 52 | 筛选第四季度中的所有日期 |
| xlBottomCount | 2 | 从列表底部筛选指定数量的值 |
| xlBottomPercent | 4 | 从列表底部筛选指定百分比的值 |
| xlBottomSum | 6 | 列表底部的值的总和 |
| xlCaptionBeginsWith | 17 | 筛选以指定字符串开头的所有标题 |
| xlCaptionContains | 21 | 筛选包含指定字符串的所有标题 |
| xlCaptionDoesNotBeginWith | 18 | 筛选不以指定字符串开头的所有标题 |
| xlCaptionDoesNotContain | 22 | 筛选不包含指定字符串的所有标题 |
| xlCaptionDoesNotEndWith | 20 | 筛选不以指定字符串结尾的所有标题 |
| xlCaptionDoesNotEqual | 16 | 筛选不与指定字符串匹配的所有标题 |
| xlCaptionEndsWith | 19 | 筛选以指定字符串结尾的所有标题 |
| xlCaptionEquals | 15 | 筛选与指定字符串匹配的所有标题 |
| xlCaptionIsBetween | 27 | 筛选介于指定值范围内的所有标题 |
| xlCaptionIsGreaterThan | 23 | 筛选大于指定值的所有标题 |
| xlCaptionIsGreaterThanOrEqualTo | 24 | 筛选大于指定值或与指定值匹配的所有标题 |
| xlCaptionIsLessThan | 25 | 筛选小于指定值的所有标题 |
| xlCaptionIsLessThanOrEqualTo | 26 | 筛选小于指定值或与指定值匹配的所有标题 |
| xlCaptionIsNotBetween | 28 | 筛选不介于指定值范围内的所有标题 |
| xlDateBetween | 32 | 筛选介于指定日期范围内的所有日期 |
| xlDateLastMonth | 41 | 筛选牵涉到上个月的所有日期 |
| xlDateLastQuarter | 44 | 筛选牵涉到上季度的所有日期 |
| xlDateLastWeek | 38 | 筛选牵涉到上周的所有日期 |
| xlDateLastYear | 47 | 筛选牵涉到上一年的所有日期 |
| xlDateNextMonth | 39 | 筛选牵涉到下月的所有日期 |
| xlDateNextQuarter | 42 | 筛选牵涉到下季度的所有日期 |
| xlDateNextWeek | 36 | 筛选牵涉到下周的所有日期 |
| xlDateNextYear | 45 | 筛选牵涉到下一年的所有日期 |
| xlDateThisMonth | 40 | 筛选牵涉到本月的所有日期 |
| xlDateThisQuarter | 43 | 筛选牵涉到本季度的所有日期 |
| xlDateThisWeek | 37 | 筛选牵涉到本周的所有日期 |
| xlDateThisYear | 46 | 筛选牵涉到本年度的所有日期 |
| xlDateToday | 34 | 筛选牵涉到当前日期的所有日期 |
| xlDateTomorrow | 33 | 筛选牵涉到下一天的所有日期 |
| xlDateYesterday | 35 | 筛选牵涉到前一天的所有日期 |
| xlNotSpecificDate | 30 | 筛选与指定日期不匹配的所有日期 |
| xlSpecificDate | 29 | 筛选与指定日期匹配的所有日期 |
| xlTopCount | 1 | 从列表顶部筛选指定数量的值 |
| xlTopPercent | 3 | 从列表中筛选指定百分比的值 |
| xlTopSum | 5 | 列表顶部的值的总和 |
| xlValueDoesNotEqual | 8 | 筛选与指定值不匹配的所有值 |
| xlValueEquals | 7 | 筛选与指定值匹配的所有值 |
| xlValueIsBetween | 13 | 筛选介于指定值范围内的所有值 |
| xlValueIsGreaterThan | 9 | 筛选大于指定值的所有值 |
| xlValueIsGreaterThanOrEqualTo | 10 | 筛选大于指定值或与指定值匹配的所有值 |
| xlValueIsLessThan | 11 | 筛选小于指定值的所有值 |
| xlValueIsLessThanOrEqualTo | 12 | 筛选小于指定值或与指定值匹配的所有值 |
| xlValueIsNotBetween | 14 | 筛选不介于指定值范围内的所有值 |
| xlYearToDate | 48 | 筛选指定日期的一年内的所有值 |


#### XlPivotFormatType 枚举

# [XlPivotFormatType 枚举​](#xlpivotformattype-枚举)

指定要应用于指定数据透视表的报表格式类型。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlPTClassic | 20 | 数据透视表传统格式。 |
| xlPTNone | 21 | 不对数据透视表应用格式。 |
| xlReport1 | 0 | 对数据透视表使用 xlReport1 格式。 |
| xlReport10 | 9 | 对数据透视表使用 xlReport10 格式。 |
| xlReport2 | 1 | 对数据透视表使用 xlReport2 格式。 |
| xlReport3 | 2 | 对数据透视表使用 xlReport3 格式。 |
| xlReport4 | 3 | 对数据透视表使用 xlReport4 格式。 |
| xlReport5 | 4 | 对数据透视表使用 xlReport5 格式。 |
| xlReport6 | 5 | 对数据透视表使用 xlReport6 格式。 |
| xlReport7 | 6 | 对数据透视表使用 xlReport7 格式。 |
| xlReport8 | 7 | 对数据透视表使用 xlReport8 格式。 |
| xlReport9 | 8 | 对数据透视表使用 xlReport9 格式。 |
| xlTable1 | 10 | 对数据透视表使用 xlTable1 格式。 |
| xlTable10 | 19 | 对数据透视表使用 xlTable10 格式。 |
| xlTable2 | 11 | 对数据透视表使用 xlTable2 格式。 |
| xlTable3 | 12 | 对数据透视表使用 xlTable3 格式。 |
| xlTable4 | 13 | 对数据透视表使用 xlTable4 格式。 |
| xlTable5 | 14 | 对数据透视表使用 xlTable5 格式。 |
| xlTable6 | 15 | 对数据透视表使用 xlTable6 格式。 |
| xlTable7 | 16 | 对数据透视表使用 xlTable7 格式。 |
| xlTable8 | 17 | 对数据透视表使用 xlTable8 格式。 |
| xlTable9 | 18 | 对数据透视表使用 xlTable9 格式。 |


#### XlPivotLineType 枚举

# [XlPivotLineType 枚举​](#xlpivotlinetype-枚举)

指定 PivotLine 的类型。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlPivotLineBlank | 3 | 每组后的空行。 |
| xlPivotLineGrandTotal | 2 | 总计行。 |
| xlPivotLineRegular | 0 | 带有透视项目的常规 PivotLine。 |
| xlPivotLineSubtotal | 1 | 分类汇总行。 |


#### XlPivotTableMissingItems 枚举

# [XlPivotTableMissingItems 枚举​](#xlpivottablemissingitems-枚举)

指定每个透视字段允许具有的唯一项的最大数量。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlMissingItemsDefault | -1 | 允许每个透视字段具有的唯一项的默认数量。 |
| xlMissingItemsMax | 32500 | ET 2007 之前的数据透视表允许每个透视字段具有的唯一项的最大数量 (32,500)。 |
| xlMissingItemsMax2 | 1048576 | ET 2007 之前的数据透视表允许每个透视字段具有的唯一项的最大数量 (10,48,576)。 |
| xlMissingItemsNone | 0 | 每个透视表中不允许具有唯一项（零）。 |


#### XlPivotTableSourceType 枚举

# [XlPivotTableSourceType 枚举​](#xlpivottablesourcetype-枚举)

指定报告数据源。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlConsolidation | 3 | 多重合并计算数据区域。 |
| xlDatabase | 1 | ET 列表或数据库。 |
| xlExternal | 2 | 其他应用程序中的数据。 |
| xlPivotTable | -4148 | 与另一数据透视表相同来源。 |
| xlScenario | 4 | 数据基于使用方案管理器创建的方案。 |


#### XlPivotTableVersionList 枚举

# [XlPivotTableVersionList 枚举​](#xlpivottableversionlist-枚举)

指定数据透视表或数据透视表缓存的版本。创建特定版本的数据透视表可确保在 中创建的表的行为方式与它们在 ET 的相应版本中的行为方式相同。

| 注释 |
| --- |
| xlPivotTableVersionCurrent仅出于向后兼容性的原因而包含在内。它不能与新的PivotCache和PivotTable对象一起使用。xlPivotTableVersion11和xlPivotTableVersion10之间的行为并无差异。 |

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlPivotTableVersion2000 | 0 | ET 2000 |
| xlPivotTableVersion10 | 1 | ET 2002 |
| xlPivotTableVersion11 | 2 | ET 2003 |
| xlPivotTableVersion12 | 3 | ET 2007 |
| xlPivotTableVersion14 | 4 | ET 2010 |
| xlPivotTableVersionCurrent | -1 | 仅为向后兼容性而提供 |


#### XlPlacement 枚举

# [XlPlacement 枚举​](#xlplacement-枚举)

指定对象附加到其下层单元格的方式。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlFreeFloating | 3 | 对象自由浮动。 |
| xlMove | 2 | 对象随单元格移动。 |
| xlMoveAndSize | 1 | 对象随单元格移动和调整大小。 |


#### XlPlatform 枚举

# [XlPlatform 枚举​](#xlplatform-枚举)

指定生成文本文件的平台。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlMacintosh | 1 | Macintosh |
| xlMSDOS | 3 | MS-DOS |
| xlWindows | 2 | Microsoft Windows |


#### XlPortugueseReform 枚举

# [XlPortugueseReform 枚举​](#xlportuguesereform-枚举)

指定葡萄牙语拼写检查模式。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlPortugueseBoth | 3 | 拼写检查器识别前期修订和后期修订拼写。 |
| xlPortuguesePostReform | 2 | 拼写检查器只识别后期修订拼写。 |
| xlPortuguesePreReform | 1 | 拼写检查器只识别前期修订拼写。 |


#### XlPrintErrors 枚举

# [XlPrintErrors 枚举​](#xlprinterrors-枚举)

指定显示的打印错误的类型。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlPrintErrorsBlank | 1 | 打印错误为空白。 |
| xlPrintErrorsDash | 2 | 打印错误显示为划线。 |
| xlPrintErrorsDisplayed | 0 | 显示全部打印错误。 |
| xlPrintErrorsNA | 3 | 打印错误显示为不可用。 |


#### XlPrintLocation 枚举

# [XlPrintLocation 枚举​](#xlprintlocation-枚举)

指定表中批注的打印方式。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlPrintInPlace | 16 | 批注打印在其插入工作表的位置。 |
| xlPrintNoComments | -4142 | 不打印批注。 |
| xlPrintSheetEnd | 1 | 批注打印为工作表末尾的尾注。 |


#### XlPriority 枚举

# [XlPriority 枚举​](#xlpriority-枚举)

指定 SendMailer 消息的优先级。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlPriorityHigh | -4127 | 高 |
| xlPriorityLow | -4134 | 低 |
| xlPriorityNormal | -4143 | 中 |


#### XlPropertyDisplayedIn 枚举

# [XlPropertyDisplayedIn 枚举​](#xlpropertydisplayedin-枚举)

指定显示属性的位置。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlDisplayPropertyInPivotTable | 1 | 只在数据透视表中显示成员属性。这是默认值。 |
| xlDisplayPropertyInPivotTableAndTooltip | 3 | 只在工具提示中显示成员属性。 |
| xlDisplayPropertyInTooltip | 2 | 同时在工具提示和数据透视表中显示成员属性。 |


#### XlProtectedViewCloseReason 枚举

# [XlProtectedViewCloseReason 枚举​](#xlprotectedviewclosereason-枚举)

指定如何关闭**“受保护的视图”**窗口。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlProtectedViewCloseEdit | 1 | 窗口在用户单击**“启用编辑”**按钮时关闭。 |
| xlProtectedViewCloseForced | 2 | 窗口由于应用程序强制将其关闭或停止响应而关闭。 |
| xlProtectedViewCloseNormal | 0 | 窗口正常关闭。 |


#### XlProtectedViewWindowState 枚举

# [XlProtectedViewWindowState 枚举​](#xlprotectedviewwindowstate-枚举)

指定**“受保护的视图”**窗口的状态。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlProtectedViewWindowMaximized | 2 | 最大化 |
| xlProtectedViewWindowMinimized | 1 | 最小化 |
| xlProtectedViewWindowNormal | 0 | 正常 |


#### XlQueryType 枚举

# [XlQueryType 枚举​](#xlquerytype-枚举)

指定 ET 在填充查询表或数据透视表缓存时所使用的查询类型。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlADORecordset | 7 | 基于 ADO 记录集查询 |
| xlDAORecordset | 2 | 基于 DAO 记录集查询，只用于查询表 |
| xlODBCQuery | 1 | 基于 ODBC 数据源 |
| xlOLEDBQuery | 5 | 基于 OLE DB 查询，包括 OLAP 数据源 |
| xlTextImport | 6 | 基于文本文件，仅用于查询表 |
| xlWebQuery | 4 | 基于网页，仅用于查询表 |


#### XlRangeAutoFormat 枚举

# [XlRangeAutoFormat 枚举​](#xlrangeautoformat-枚举)

指定自动设置区域格式时的预定义格式。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlRangeAutoFormat3DEffects1 | 13 | 三维效果 1。 |
| xlRangeAutoFormat3DEffects2 | 14 | 三维效果 2。 |
| xlRangeAutoFormatAccounting1 | 4 | 会计 1。 |
| xlRangeAutoFormatAccounting2 | 5 | 会计 2。 |
| xlRangeAutoFormatAccounting3 | 6 | 会计 3。 |
| xlRangeAutoFormatAccounting4 | 17 | 会计 4。 |
| xlRangeAutoFormatClassic1 | 1 | 古典 1。 |
| xlRangeAutoFormatClassic2 | 2 | 古典 2。 |
| xlRangeAutoFormatClassic3 | 3 | 古典 3。 |
| xlRangeAutoFormatClassicPivotTable | 31 | 传统数据透视表。 |
| xlRangeAutoFormatColor1 | 7 | 彩色 1。 |
| xlRangeAutoFormatColor2 | 8 | 彩色 2。 |
| xlRangeAutoFormatColor3 | 9 | 彩色 3。 |
| xlRangeAutoFormatList1 | 10 | 列表 1。 |
| xlRangeAutoFormatList2 | 11 | 列表 2。 |
| xlRangeAutoFormatList3 | 12 | 列表 3。 |
| xlRangeAutoFormatLocalFormat1 | 15 | 本地格式 1。 |
| xlRangeAutoFormatLocalFormat2 | 16 | 本地格式 2。 |
| xlRangeAutoFormatLocalFormat3 | 19 | 本地格式 3。 |
| xlRangeAutoFormatLocalFormat4 | 20 | 本地格式 4。 |
| xlRangeAutoFormatNone | -4142 | 无指定格式。 |
| xlRangeAutoFormatPTNone | 42 | 无指定数据透视表格式。 |
| xlRangeAutoFormatReport1 | 21 | 报表 1。 |
| xlRangeAutoFormatReport10 | 30 | 报表 10。 |
| xlRangeAutoFormatReport2 | 22 | 报表 2。 |
| xlRangeAutoFormatReport3 | 23 | 报表 3。 |
| xlRangeAutoFormatReport4 | 24 | 报表 4。 |
| xlRangeAutoFormatReport5 | 25 | 报表 5。 |
| xlRangeAutoFormatReport6 | 26 | 报表 6。 |
| xlRangeAutoFormatReport7 | 27 | 报表 7。 |
| xlRangeAutoFormatReport8 | 28 | 报表 8。 |
| xlRangeAutoFormatReport9 | 29 | 报表 9。 |
| xlRangeAutoFormatSimple | -4154 | 简单。 |
| xlRangeAutoFormatTable1 | 32 | 表 1。 |
| xlRangeAutoFormatTable10 | 41 | 表 10。 |
| xlRangeAutoFormatTable2 | 33 | 表 2。 |
| xlRangeAutoFormatTable3 | 34 | 表 3。 |
| xlRangeAutoFormatTable4 | 35 | 表 4。 |
| xlRangeAutoFormatTable5 | 36 | 表 5。 |
| xlRangeAutoFormatTable6 | 37 | 表 6。 |
| xlRangeAutoFormatTable7 | 38 | 表 7。 |
| xlRangeAutoFormatTable8 | 39 | 表 8。 |
| xlRangeAutoFormatTable9 | 40 | 表 9。 |


#### XlRangeValueDataType 枚举

# [XlRangeValueDataType 枚举​](#xlrangevaluedatatype-枚举)

指定区域值数据类型。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlRangeValueDefault | 10 | 默认值。如果指定的Range对象为空，则返回值 Empty（可用 IsEmpty 函数测试这种情况）。如果Range对象包含多个单元格，则返回值的数组（可用 IsArray 函数测试这种情况）。 |
| xlRangeValueMSPersistXML | 12 | 以 XML 格式返回指定的Range对象的记录集表示形式。 |
| xlRangeValueXMLSpreadsheet | 11 | 以 XML 电子表格格式返回指定的Range对象的值、格式、公式和名称。 |


#### XlReferenceStyle 枚举

# [XlReferenceStyle 枚举​](#xlreferencestyle-枚举)

指定引用样式。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlA1 | 1 | 默认值。使用xlA1返回 A1 样式的引用。 |
| xlR1C1 | -4150 | 使用xlR1C1返回 R1C1 样式的引用。 |


#### XlReferenceType 枚举

# [XlReferenceType 枚举​](#xlreferencetype-枚举)

指定转换公式时的单元格引用样式。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlAbsolute | 1 | 转换为绝对行和列样式。 |
| xlAbsRowRelColumn | 2 | 转换为绝对行和相对列样式。 |
| xlRelative | 4 | 转换为相对行和列样式。 |
| xlRelRowAbsColumn | 3 | 转换为相对行和绝对列样式。 |


#### XlRemoveDocInfoType 枚举

# [XlRemoveDocInfoType 枚举​](#xlremovedocinfotype-枚举)

指定要从文档信息中删除的类型信息。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlRDIAll | 99 | 删除所有文档信息。 |
| xlRDIComments | 1 | 从文档信息中删除批注。 |
| xlRDIContentType | 16 | 从文档信息中删除内容类型数据。 |
| xlRDIDefinedNameComments | 18 | 从文档信息中删除定义的名称批注。 |
| xlRDIDocumentManagementPolicy | 15 | 从文档信息中删除文档管理策略数据。 |
| xlRDIDocumentProperties | 8 | 从文档信息中删除文档属性。 |
| xlRDIDocumentServerProperties | 14 | 从文档信息中删除服务器属性。 |
| xlRDIDocumentWorkspace | 10 | 从文档信息中删除工作空间数据。 |
| xlRDIEmailHeader | 5 | 从文档信息中删除电子邮件头。 |
| xlRDIInactiveDataConnections | 19 | 从文档信息中删除非活动数据连接数据。 |
| xlRDIInkAnnotations | 11 | 从文档信息中删除墨迹注释。 |
| xlRDIPrinterPath | 20 | 从文档信息中删除指针路径。 |
| xlRDIPublishInfo | 13 | 从文档信息中删除发布信息数据。 |
| xlRDIRemovePersonalInformation | 4 | 从文档信息中删除个人信息。 |
| xlRDIRoutingSlip | 6 | 从文档信息中删除传送名单信息。 |
| xlRDIScenarioComments | 12 | 从文档信息中删除方案批注。 |
| xlRDISendForReview | 7 | 从文档信息中删除请求审阅信息。 |


#### XlRgbColor 枚举

# [XlRgbColor 枚举​](#xlrgbcolor-枚举)

指定 RGB 颜色。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| rgbAliceBlue | 16775408 | 艾莉斯蓝 |
| rgbAntiqueWhite | 14150650 | 古董白 |
| rgbAqua | 16776960 | 青色 |
| rgbAquamarine | 13959039 | 玉色 |
| rgbAzure | 16777200 | 蔚蓝色 |
| rgbBeige | 14480885 | 米色 |
| rgbBisque | 12903679 | 乳黄色 |
| rgbBlack | 0 | 黑色 |
| rgbBlanchedAlmond | 13495295 | 杏仁白 |
| rgbBlue | 16711680 | 蓝色 |
| rgbBlueViolet | 14822282 | 蓝紫色 |
| rgbBrown | 2763429 | 褐色 |
| rgbBurlyWood | 8894686 | 原木色 |
| rgbCadetBlue | 10526303 | 军队蓝 |
| rgbChartreuse | 65407 | 浅黄绿色 |
| rgbCoral | 5275647 | 珊瑚红 |
| rgbCornflowerBlue | 15570276 | 藏蓝色 |
| rgbCornsilk | 14481663 | 玉米黄 |
| rgbCrimson | 3937500 | 暗红色 |
| rgbDarkBlue | 9109504 | 深蓝色 |
| rgbDarkCyan | 9145088 | 深青色 |
| rgbDarkGoldenrod | 755384 | 深金黄色 |
| rgbDarkGray | 11119017 | 深灰色 |
| rgbDarkGreen | 25600 | 深绿色 |
| rgbDarkGrey | 11119017 | 深灰色 |
| rgbDarkKhaki | 7059389 | 深褐色 |
| rgbDarkMagenta | 9109643 | 深洋红色 |
| rgbDarkOliveGreen | 3107669 | 深橄榄绿色 |
| rgbDarkOrange | 36095 | 深橙色 |
| rgbDarkOrchid | 13382297 | 深兰花色 |
| rgbDarkRed | 139 | 深红色 |
| rgbDarkSalmon | 8034025 | 深橙红 |
| rgbDarkSeaGreen | 9419919 | 深海绿色 |
| rgbDarkSlateBlue | 9125192 | 深灰蓝色 |
| rgbDarkSlateGray | 5197615 | 深石板灰 |
| rgbDarkSlateGrey | 5197615 | 深石板灰 |
| rgbDarkTurquoise | 13749760 | 深青绿色 |
| rgbDarkViolet | 13828244 | 深紫色 |
| rgbDeepPink | 9639167 | 深粉色 |
| rgbDeepSkyBlue | 16760576 | 深天蓝色 |
| rgbDimGray | 6908265 | 暗灰色 |
| rgbDimGrey | 6908265 | 暗灰色 |
| rgbDodgerBlue | 16748574 | 宝蓝 |
| rgbFireBrick | 2237106 | 砖红色 |
| rgbFloralWhite | 15792895 | 花白 |
| rgbForestGreen | 2263842 | 森林绿 |
| rgbFuchsia | 16711935 | 紫红色 |
| rgbGainsboro | 14474460 | 亮灰 |
| rgbGhostWhite | 16775416 | 苍白 |
| rgbGold | 55295 | 金色 |
| rgbGoldenrod | 2139610 | 金黄色 |
| rgbGray | 8421504 | 灰色 |
| rgbGreen | 32768 | 绿色 |
| rgbGreenYellow | 3145645 | 青黄色 |
| rgbGrey | 8421504 | 灰色 |
| rgbHoneydew | 15794160 | 蜜色 |
| rgbHotPink | 11823615 | 暗粉 |
| rgbIndianRed | 6053069 | 印度红 |
| rgbIndigo | 8519755 | 靛蓝色 |
| rgbIvory | 15794175 | 象牙色 |
| rgbKhaki | 9234160 | 黄褐色 |
| rgbLavender | 16443110 | 淡紫色 |
| rgbLavenderBlush | 16118015 | 淡紫红色 |
| rgbLawnGreen | 64636 | 草绿色 |
| rgbLemonChiffon | 13499135 | 柠檬色 |
| rgbLightBlue | 15128749 | 浅蓝色 |
| rgbLightCoral | 8421616 | 浅珊瑚红 |
| rgbLightCyan | 9145088 | 浅青色 |
| rgbLightGoldenrodYellow | 13826810 | 浅金黄 |
| rgbLightGray | 13882323 | 浅灰色 |
| rgbLightGreen | 9498256 | 浅绿色 |
| rgbLightGrey | 13882323 | 浅灰色 |
| rgbLightPink | 12695295 | 浅粉色 |
| rgbLightSalmon | 8036607 | 浅橙红 |
| rgbLightSeaGreen | 11186720 | 浅海绿色 |
| rgbLightSkyBlue | 16436871 | 浅天蓝色 |
| rgbLightSlateGray | 10061943 | 浅石板灰 |
| rgbLightSteelBlue | 14599344 | 浅钢蓝色 |
| rgbLightYellow | 14745599 | 浅黄色 |
| rgbLime | 65280 | 酸橙色 |
| rgbLimeGreen | 3329330 | 暗黄绿色 |
| rgbLinen | 15134970 | 亚麻布色 |
| rgbMaroon | 128 | 褐紫红色 |
| rgbMediumAquamarine | 11206502 | 中玉色 |
| rgbMediumBlue | 13434880 | 中蓝色 |
| rgbMediumOrchid | 13850042 | 中兰花色 |
| rgbMediumPurple | 14381203 | 中紫色 |
| rgbMediumSeaGreen | 7451452 | 中海绿色 |
| rgbMediumSlateBlue | 15624315 | 中蓝灰色 |
| rgbMediumSpringGreen | 10156544 | 中草绿色 |
| rgbMediumTurquoise | 13422920 | 中玉色 |
| rgbMediumVioletRed | 8721863 | 中紫罗兰色 |
| rgbMidnightBlue | 7346457 | 蓝黑色 |
| rgbMintCream | 16449525 | 薄荷乳白 |
| rgbMistyRose | 14804223 | 粉红玫瑰 |
| rgbMoccasin | 11920639 | 鹿皮黄 |
| rgbNavajoWhite | 11394815 | 印地安黄 |
| rgbNavy | 8388608 | 海军蓝 |
| rgbNavyBlue | 8388608 | 海军蓝 |
| rgbOldLace | 15136253 | 旧布黄 |
| rgbOlive | 32896 | 橄榄色 |
| rgbOliveDrab | 2330219 | 暗橄榄色 |
| rgbOrange | 42495 | 橙色 |
| rgbOrangeRed | 17919 | 桔红色 |
| rgbOrchid | 14053594 | 兰花色 |
| rgbPaleGoldenrod | 7071982 | 淡金黄色 |
| rgbPaleGreen | 10025880 | 淡绿色 |
| rgbPaleTurquoise | 15658671 | 浅青绿色 |
| rgbPaleVioletRed | 9662683 | 浅紫红色 |
| rgbPapayaWhip | 14020607 | 粉木瓜橙 |
| rgbPeachPuff | 12180223 | 粉桃红 |
| rgbPeru | 4163021 | 秘鲁棕 |
| rgbPink | 13353215 | 粉红色 |
| rgbPlum | 14524637 | 青紫色 |
| rgbPowderBlue | 15130800 | 粉蓝色 |
| rgbPurple | 8388736 | 紫色 |
| rgbRed | 255 | 红色 |
| rgbRosyBrown | 9408444 | 玫瑰褐色 |
| rgbRoyalBlue | 14772545 | 贵族蓝 |
| rgbSalmon | 7504122 | 浅橙色 |
| rgbSandyBrown | 6333684 | 浅褐色 |
| rgbSeaGreen | 5737262 | 海绿色 |
| rgbSeashell | 15660543 | 贝壳白 |
| rgbSienna | 2970272 | 赭色 |
| rgbSilver | 12632256 | 银白 |
| rgbSkyBlue | 15453831 | 天蓝色 |
| rgbSlateBlue | 13458026 | 灰蓝色 |
| rgbSlateGray | 9470064 | 石板灰 |
| rgbSnow | 16448255 | 雪白 |
| rgbSpringGreen | 8388352 | 草绿色 |
| rgbSteelBlue | 11829830 | 刚蓝色 |
| rgbTan | 9221330 | 茶色 |
| rgbTeal | 8421376 | 青色 |
| rgbThistle | 14204888 | 蓟色 |
| rgbTomato | 4678655 | 番茄色 |
| rgbTurquoise | 13688896 | 青绿色 |
| rgbViolet | 15631086 | 紫罗兰色 |
| rgbWheat | 11788021 | 淡黄色 |
| rgbWhite | 16777215 | 白色 |
| rgbWhiteSmoke | 16119285 | 烟白色 |
| rgbYellow | 65535 | 黄色 |
| rgbYellowGreen | 3329434 | 黄绿色 |


#### XlRobustConnect 枚举

# [XlRobustConnect 枚举​](#xlrobustconnect-枚举)

指定数据透视表缓存与其数据源连接的方式。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlAlways | 1 | 缓存始终使用外部源信息（由SourceConnectionFile或SourceDataFile属性定义）进行重新连接。 |
| xlAsRequired | 0 | 缓存通过Connection属性使用外部源信息进行重新连接。 |
| xlNever | 2 | 缓存从不使用源信息进行重新连接。 |


#### XlRoutingSlipDelivery 枚举

# [XlRoutingSlipDelivery 枚举​](#xlroutingslipdelivery-枚举)

指定传送传递方法。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlAllAtOnce | 2 | 同时传递给所有收件人。 |
| xlOneAfterAnother | 1 | 逐个传递给收件人。 |


#### XlRoutingSlipStatus 枚举

# [XlRoutingSlipStatus 枚举​](#xlroutingslipstatus-枚举)

指定传送名单的状态。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlNotYetRouted | 0 | 还未发送传送名单。 |
| xlRoutingComplete | 2 | 完成传送。 |
| xlRoutingInProgress | 1 | 正在传送。 |


#### XlRunAutoMacro 枚举

# [XlRunAutoMacro 枚举​](#xlrunautomacro-枚举)

指定要运行的自动宏。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlAutoActivate | 3 | Auto_Activate 宏 |
| xlAutoClose | 2 | Auto_Close 宏 |
| xlAutoDeactivate | 4 | Auto_Deactivate 宏 |
| xlAutoOpen | 1 | Auto_Open 宏 |


#### XlSaveAction 枚举

# [XlSaveAction 枚举​](#xlsaveaction-枚举)

如果将保存文件，则在文件关闭过程中进行指定。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlDoNotSaveChanges | 2 | 不保存更改。 |
| xlSaveChanges | 1 | 保存更改。 |


#### XlSaveAsAccessMode 枚举

# [XlSaveAsAccessMode 枚举​](#xlsaveasaccessmode-枚举)

指定“另存为”函数的访问模式。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlExclusive | 3 | 独占模式 |
| xlNoChange | 1 | 默认值（不更改访问模式） |
| xlShared | 2 | 共享列表 |


#### XlSaveConflictResolution 枚举

# [XlSaveConflictResolution 枚举​](#xlsaveconflictresolution-枚举)

指定更新共享工作簿时解决冲突的方式。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlLocalSessionChanges | 2 | 总是接受本地用户所做的更改。 |
| xlOtherSessionChanges | 3 | 总是拒绝本地用户所做的更改。 |
| xlUserResolution | 1 | 弹出对话框请求用户解决冲突。 |


#### XlSearchDirection 枚举

# [XlSearchDirection 枚举​](#xlsearchdirection-枚举)

指定搜索区域时的搜索方向。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlNext | 1 | 在区域中搜索下一匹配值。 |
| xlPrevious | 2 | 在区域中搜索上一匹配值。 |


#### XlSearchOrder 枚举

# [XlSearchOrder 枚举​](#xlsearchorder-枚举)

指定搜索区域的次序。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlByColumns | 2 | 向下搜索列，然后移到下一列。 |
| xlByRows | 1 | 搜索行，然后移到下一行。 |


#### XlSearchWithin 枚举

# [XlSearchWithin 枚举​](#xlsearchwithin-枚举)

指定区域的搜索范围。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlWithinSheet | 1 | 将搜索限制在当前工作表。 |
| xlWithinWorkbook | 2 | 搜索整个工作簿。 |


#### XlSheetType 枚举

# [XlSheetType 枚举​](#xlsheettype-枚举)

指定工作表类型。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlChart | -4109 | 图表 |
| xlDialogSheet | -4116 | 对话框工作表 |
| xlExcel4IntlMacroSheet | 4 | ET 版本 4 国际宏工作表 |
| xlExcel4MacroSheet | 3 | ET 版本 4 宏工作表 |
| xlWorksheet | -4167 | 工作表 |


#### XlSheetVisibility 枚举

# [XlSheetVisibility 枚举​](#xlsheetvisibility-枚举)

指定对象是否可见。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlSheetHidden | 0 | 隐藏工作表，用户可以通过菜单取消隐藏。 |
| xlSheetVeryHidden | 2 | 隐藏对象，以便使对象重新可见的唯一方法是将此属性设置为 True（用户无法使该对象可见）。 |
| xlSheetVisible | -1 | 显示工作表。 |


#### XlSlicerCrossFilterType 枚举

# [XlSlicerCrossFilterType 枚举​](#xlslicercrossfiltertype-枚举)

指定由指定的切片器缓存所使用的交叉筛选类型以及显示方式。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlSlicerCrossFilterShowItemsWithDataAtTop | 2 | 为此切片器缓存开启交叉筛选，对于连接到同一数据源的其他切片器中的筛选选择，没有数据的任何平铺都将变灰。此外，有数据的平铺将移到切片器的顶部。（默认） |
| xlSlicerCrossFilterShowItemsWithNoData | 3 | 为此切片器缓存开启交叉筛选，对于连接到同一数据源的其他切片器中的筛选选择，没有数据的任何平铺都将变灰。 |
| xlSlicerNoCrossFilter | 1 | 完全关闭交叉筛选，因此，无论其他切片器中的筛选选择如何，所有平铺都将显示，并处于活动状态（未变灰）。 |


#### XlSlicerSort 枚举

# [XlSlicerSort 枚举​](#xlslicersort-枚举)

指定在切片器中显示的项是否排序，如果排序，则是按项标题升序还是降序排序。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlSlicerSortAscending | 2 | 切片器项按项标题升序排序。 |
| xlSlicerSortDataSourceOrder | 1 | 切片器项按数据源提供的顺序显示。 |
| xlSlicerSortDescending | 3 | 切片器项按项标题降序排序。 |


#### XlSmartTagControlType 枚举

# [XlSmartTagControlType 枚举​](#xlsmarttagcontroltype-枚举)

指定**“文档操作”**任务窗格中显示的智能文档控件的类型。

| 注释 |
| --- |
| 此对象或成员已弃用，但为了向后兼容，仍作为对象模型的一部分保留。在新应用程序中，不应使用该对象或成员。 |

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlSmartTagControlActiveX | 13 | ActiveX 控件。 |
| xlSmartTagControlButton | 6 | 按钮。 |
| xlSmartTagControlCheckbox | 9 | 复选框。 |
| xlSmartTagControlCombo | 12 | 组合框。 |
| xlSmartTagControlHelp | 3 | 帮助文字。 |
| xlSmartTagControlHelpURL | 4 | 帮助文件的绝对 URL。 |
| xlSmartTagControlImage | 8 | 图像。 |
| xlSmartTagControlLabel | 7 | 标签。 |
| xlSmartTagControlLink | 2 | 链接。 |
| xlSmartTagControlListbox | 11 | 列表框。 |
| xlSmartTagControlRadioGroup | 14 | 单选按钮（选项按钮）组。 |
| xlSmartTagControlSeparator | 5 | 分隔符。 |
| xlSmartTagControlSmartTag | 1 | 智能标记。 |
| xlSmartTagControlTextbox | 10 | 文本框。 |


#### XlSmartTagDisplayMode 枚举

# [XlSmartTagDisplayMode 枚举​](#xlsmarttagdisplaymode-枚举)

指定智能标记的显示功能。

| 注释 |
| --- |
| 此对象或成员已弃用，但为了向后兼容，仍作为对象模型的一部分保留。在新应用程序中，不应使用该对象或成员。 |

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlButtonOnly | 2 | 只显示智能标记的按钮。 |
| xlDisplayNone | 1 | 不显示智能标记的任何内容。 |
| xlIndicatorAndButton | 0 | 显示智能标记的指示符和按钮。 |


#### XlSortDataOption 枚举

# [XlSortDataOption 枚举​](#xlsortdataoption-枚举)

指定文本的排序方式。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlSortNormal | 0 | 默认值。分别对数字和文本数据进行排序。 |
| xlSortTextAsNumbers | 1 | 将文本作为数字型数据进行排序。 |


#### XlSortMethod 枚举

# [XlSortMethod 枚举​](#xlsortmethod-枚举)

指定排序类型。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlPinYin | 1 | 按字符的汉语拼音顺序排序。这是默认值。 |
| xlStroke | 2 | 按每个字符的笔划数排序。 |


#### XlSortMethodOld 枚举

# [XlSortMethodOld 枚举​](#xlsortmethodold-枚举)

指定在使用中文排序方法时如何排序。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlCodePage | 2 | 按代码页排序。 |
| xlSyllabary | 1 | 按发音排序。 |


#### XlSortOn 枚举

# [XlSortOn 枚举​](#xlsorton-枚举)

指定数据的排序参数。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| SortOnCellColor | 1 | 单元格颜色。 |
| SortOnFontColor | 2 | 字体颜色。 |
| SortOnIcon | 3 | 图标。 |
| SortOnValues | 0 | 值。 |


#### XlSortOrder 枚举

# [XlSortOrder 枚举​](#xlsortorder-枚举)

为指定字段或范围指定排序顺序。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlAscending | 1 | 按升序对指定字段排序。这是默认值。 |
| xlDescending | 2 | 按降序对指定字段排序。 |


#### XlSortOrientation 枚举

# [XlSortOrientation 枚举​](#xlsortorientation-枚举)

指定排序方向。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlSortColumns | 1 | 按列排序。 |
| xlSortRows | 2 | 按行排序。这是默认值。 |


#### XlSortType 枚举

# [XlSortType 枚举​](#xlsorttype-枚举)

指定要排序的元素。仅在对数据透视表排序时才使用该参数。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlSortLabels | 2 | 按标签对数据透视表排序。 |
| xlSortValues | 1 | 按值对数据透视表排序。 |


#### XlSourceType 枚举

# [XlSourceType 枚举​](#xlsourcetype-枚举)

标识源对象。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlSourceAutoFilter | 3 | “自动筛选”区域 |
| xlSourceChart | 5 | 图表 |
| xlSourcePivotTable | 6 | 数据透视表 |
| xlSourcePrintArea | 2 | 选定的用于打印的单元格区域 |
| xlSourceQuery | 7 | 查询表（外部数据区域） |
| xlSourceRange | 4 | 单元格区域 |
| xlSourceSheet | 1 | 整张工作表 |
| xlSourceWorkbook | 0 | 工作簿 |


#### XlSpanishModes 枚举

# [XlSpanishModes 枚举​](#xlspanishmodes-枚举)

指定西班牙语拼写检查模式。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlSpanishTuteoAndVoseo | 1 | Tuteo 和 Voseo 动词形式。 |
| xlSpanishTuteoOnly | 0 | 仅 Tuteo 动词形式。 |
| xlSpanishVoseoOnly | 2 | 仅 Voseo 动词形式。 |


#### XlSparkScale 枚举

# [XlSparkScale 枚举​](#xlsparkscale-枚举)

指定迷你图垂直轴的最小值或最大值如何相对于组中的其他迷你图按比例缩放。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlSparkScaleCustom | 3 | 迷你图垂直轴的最小值或最大值具有用户定义值。 |
| xlSparkScaleGroup | 1 | 组中所有迷你图垂直轴的最小值或最大值相同。 |
| xlSparkScaleSingle | 2 | 组中每个迷你图的垂直轴的最小值或最大值自动设置为其自己的计算值。 |


#### XlSparkType 枚举

# [XlSparkType 枚举​](#xlsparktype-枚举)

指定迷你图的类型。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlSparkColumn | 2 | 柱形图迷你图。 |
| xlSparkColumnStacked100 | 3 | 盈亏图表迷你图。 |
| xlSparkLine | 1 | 折线图迷你图。 |


#### XlSparklineRowCol 枚举

# [XlSparklineRowCol 枚举​](#xlsparklinerowcol-枚举)

指定当迷你图所基于的数据处于方形区域中时如何绘制迷你图。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| SparklineColumnsSquare | 2 | 按列绘制数据。 |
| SparklineNonSquare | 0 | 迷你图不绑定到方形区域中的数据。 |
| SparklineRowsSquare | 1 | 按行绘制数据。 |


#### XlSpeakDirection 枚举

# [XlSpeakDirection 枚举​](#xlspeakdirection-枚举)

指定朗读单元格的顺序。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlSpeakByColumns | 1 | 在一列上向下朗读，然后移至下一列继续朗读。 |
| xlSpeakByRows | 0 | 先朗读一行，然后移至下一行继续朗读。 |


#### XlSpecialCellsValue 枚举

# [XlSpecialCellsValue 枚举​](#xlspecialcellsvalue-枚举)

指定结果中包括具有特定类型值的单元格。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlErrors | 16 | 有错误的单元格。 |
| xlLogical | 4 | 具有逻辑值的单元格。 |
| xlNumbers | 1 | 具有数值的单元格。 |
| xlTextValues | 2 | 具有文本的单元格。 |


#### XlStdColorScale 枚举

# [XlStdColorScale 枚举​](#xlstdcolorscale-枚举)

指定标准色阶。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| ColorScaleBlackWhite | 3 | 下白上黑。 |
| ColorScaleGYR | 2 | GYR。 |
| ColorScaleRYG | 1 | RYG。 |
| ColorScaleWhiteBlack | 4 | 下黑上白。 |


#### XlSubscribeToFormat 枚举

# [XlSubscribeToFormat 枚举​](#xlsubscribetoformat-枚举)

指定订阅发布版本时所用的格式。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlSubscribeToPicture | -4147 | 图片 |
| xlSubscribeToText | -4158 | 文本 |


#### XlSubtototalLocationType 枚举

# [XlSubtototalLocationType 枚举​](#xlsubtototallocationtype-枚举)

指定分类汇总在工作表上的显示位置。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlAtBottom | 2 | 分类汇总在底部。 |
| xlAtTop | 1 | 分类汇总在顶部。 |


#### XlSummaryColumn 枚举

# [XlSummaryColumn 枚举​](#xlsummarycolumn-枚举)

指定汇总列在大纲中的位置。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlSummaryOnLeft | -4131 | 汇总列在大纲中位于明细数据列的左侧。 |
| xlSummaryOnRight | -4152 | 汇总列在大纲中位于明细数据列的右侧。 |


#### XlSummaryReportType 枚举

# [XlSummaryReportType 枚举​](#xlsummaryreporttype-枚举)

指定为方案创建的汇总类型。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlStandardSummary | 1 | 并排列出方案。 |
| xlSummaryPivotTable | -4148 | 在数据透视表中显示方案。 |


#### XlSummaryRow 枚举

# [XlSummaryRow 枚举​](#xlsummaryrow-枚举)

指定汇总行在大纲中的位置。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlSummaryAbove | 0 | 汇总行在大纲中位于明细数据行的上方。 |
| xlSummaryBelow | 1 | 汇总行在大纲中位于明细数据行的下方。 |


#### XlTabPosition 枚举

# [XlTabPosition 枚举​](#xltabposition-枚举)

指定第一个或最后一个制表位位置。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlTabPositionFirst | 0 | 第一个制表位位置。 |
| xlTabPositionLast | 1 | 最后一个制表位位置。 |


#### XlTableStyleElementType 枚举

# [XlTableStyleElementType 枚举​](#xltablestyleelementtype-枚举)

指定所用的表样式元素。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlBlankRow | 19 | 空白行 |
| xlColumnStripe1 | 7 | 列条纹 1 |
| xlColumnStripe2 | 8 | 列条纹 2 |
| xlColumnSubheading1 | 20 | 列副标题 1 |
| xlColumnSubheading2 | 21 | 列副标题 2 |
| xlColumnSubheading3 | 22 | 列副标题 3 |
| xlFirstColumn | 3 | 第一列 |
| xlFirstHeaderCell | 9 | 第一个标题单元格 |
| xlFirstTotalCell | 11 | 第一个汇总单元格 |
| xlGrandTotalColumn | 4 | 总计列 |
| xlGrandTotalRow | 2 | 总计行 |
| xlHeaderRow | 1 | 标题行 |
| xlLastColumn | 4 | 最后一列 |
| xlLastHeaderCell | 10 | 最后一个标题单元格 |
| xlLastTotalCell | 12 | 最后一个总计单元格 |
| xlPageFieldLabels | 26 | 页面字段标签 |
| xlPageFieldValues | 27 | 页面字段值 |
| xlRowStripe1 | 5 | 行条纹 1 |
| xlRowStripe2 | 6 | 行条纹 2 |
| xlRowSubheading1 | 23 | 行副标题 1 |
| xlRowSubheading2 | 24 | 行副标题 2 |
| xlRowSubheading3 | 25 | 行副标题 3 |
| xlSlicerHoveredSelectedItemWithData | 33 | 用户悬停在上面且包含数据的选定项。 |
| xlSlicerHoveredSelectedItemWithNoData | 35 | 用户悬停在上面且不包含数据的选定项。 |
| xlSlicerHoveredUnselectedItemWithData | 32 | 用户悬停在上面，未选定且包含数据的项。 |
| xlSlicerHoveredUnselectedItemWithNoData | 34 | 用户悬停在上面，未选定且不包含数据的项。 |
| xlSlicerSelectedItemWithData | 30 | 包含数据的选定项。 |
| xlSlicerSelectedItemWithNoData | 31 | 不包含数据的选定项。 |
| xlSlicerUnselectedItemWithData | 28 | 未选定且包含数据的项。 |
| xlSlicerUnselectedItemWithNoData | 29 | 未选定且不包含数据的项。 |
| xlSubtotalColumn1 | 13 | 分类汇总列 1 |
| xlSubtotalColumn2 | 14 | 分类汇总列 2 |
| xlSubtotalColumn3 | 15 | 分类汇总列 3 |
| xlSubtotalRow1 | 16 | 分类汇总行 1 |
| xlSubtotalRow2 | 17 | 分类汇总行 2 |
| xlSubtotalRow3 | 18 | 分类汇总行 3 |
| xlTotalRow | 2 | 汇总行 |
| xlWholeTable | 0 | 整个表 |


#### XlTextParsingType 枚举

# [XlTextParsingType 枚举​](#xltextparsingtype-枚举)

指定要导入查询表的文本文件中数据的列格式。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlDelimited | 1 | 默认值。指示文件由分隔符分隔。 |
| xlFixedWidth | 2 | 指示将文件中的数据排列在固定宽度的列中。 |


#### XlTextQualifier 枚举

# [XlTextQualifier 枚举​](#xltextqualifier-枚举)

指定用于指定文本的分隔符。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlTextQualifierDoubleQuote | 1 | 双引号 (")。 |
| xlTextQualifierNone | -4142 | 无分隔符。 |
| xlTextQualifierSingleQuote | 2 | 单引号 (')。 |


#### XlTextVisualLayoutType 枚举

# [XlTextVisualLayoutType 枚举​](#xltextvisuallayouttype-枚举)

指定所导入文本的可视布局是从左向右还是从右向左。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlTextVisualLTR | 1 | 从左向右 |
| xlTextVisualRTL | 2 | 从右向左 |


#### XlThemeColor 枚举

# [XlThemeColor 枚举​](#xlthemecolor-枚举)

指定要使用的主题颜色。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlThemeColorAccent1 | 5 | 强调文字颜色 1 |
| xlThemeColorAccent2 | 6 | 强调文字颜色 2 |
| xlThemeColorAccent3 | 7 | 强调文字颜色 3 |
| xlThemeColorAccent4 | 8 | 强调文字颜色 4 |
| xlThemeColorAccent5 | 9 | 强调文字颜色 5 |
| xlThemeColorAccent6 | 10 | 强调文字颜色 6 |
| xlThemeColorDark1 | 1 | 深色 1 |
| xlThemeColorDark2 | 3 | 深色 2 |
| xlThemeColorFollowedHyperlink | 12 | 已访问的超链接 |
| xlThemeColorHyperlink | 11 | 超链接 |
| xlThemeColorLight1 | 2 | 浅色 1 |
| xlThemeColorLight2 | 4 | 浅色 2 |


#### XlThemeFont 枚举

# [XlThemeFont 枚举​](#xlthemefont-枚举)

指定要使用的主题字体。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlThemeFontMajor | 2 | 主要。 |
| xlThemeFontMinor | 1 | 次要。 |
| xlThemeFontNone | 0 | 不使用任何主题字体。 |


#### XlThreadMode 枚举

# [XlThreadMode 枚举​](#xlthreadmode-枚举)

指定多线程计算模式的控制方式。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlThreadModeAutomatic | 0 | 多线程计算模式是自动的。 |
| xlThreadModeManual | 1 | 多线程计算模式是手动的。 |


#### XlTimePeriods 枚举

# [XlTimePeriods 枚举​](#xltimeperiods-枚举)

指定时间段。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlLast7Days | 2 | 过去 7 天 |
| xlLastMonth | 5 | 上月 |
| xlLastWeek | 4 | 上周 |
| xlNextMonth | 8 | 下月 |
| xlNextWeek | 7 | 下周 |
| xlThisMonth | 9 | 本月 |
| xlThisWeek | 3 | 本周 |
| xlToday | 0 | 今天 |
| xlTomorrow | 6 | 明天 |
| xlYesterday | 1 | 昨天 |


#### XlToolbarProtection 枚举

# [XlToolbarProtection 枚举​](#xltoolbarprotection-枚举)

指定工具栏的哪些属性受到限制。可用 Or 组合选项。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlNoButtonChanges | 1 | 不允许按钮更改。 |
| xlNoChanges | 4 | 无任何类型的更改。 |
| xlNoDockingChanges | 3 | 无对工具栏固定位置的更改。 |
| xlNoShapeChanges | 2 | 无对工具栏形状的更改。 |
| xlToolbarProtectionNone | -4143 | 允许任何更改。 |


#### XlTopBottom 枚举

# [XlTopBottom 枚举​](#xltopbottom-枚举)

指定值系列的前 10 个或后 10 个值。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlTop10Bottom | 0 | 后 10 个值 |
| xlTop10Top | 1 | 前 10 个值 |


#### XlTotalsCalculation 枚举

# [XlTotalsCalculation 枚举​](#xltotalscalculation-枚举)

指定列表列的汇总行中的计算类型。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlTotalsCalculationAverage | 2 | 平均 |
| xlTotalsCalculationCount | 3 | 对非空单元格进行计数 |
| xlTotalsCalculationCountNums | 4 | 对数值单元格进行计数 |
| xlTotalsCalculationCustom | 9 | 自定义计算 |
| xlTotalsCalculationMax | 6 | 列表中的最大值 |
| xlTotalsCalculationMin | 5 | 列表中的最小值 |
| xlTotalsCalculationNone | 0 | 无计算 |
| xlTotalsCalculationStdDev | 7 | 标准偏差值 |
| xlTotalsCalculationSum | 1 | 列表列中所有值的和 |
| xlTotalsCalculationVar | 8 | 变量 |


#### XlUpdateLinks 枚举

# [XlUpdateLinks 枚举​](#xlupdatelinks-枚举)

指定工作簿用于更新嵌入式 OLE 链接的设置。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlUpdateLinksAlways | 3 | 始终更新指定工作簿的嵌入式 OLE 链接。 |
| xlUpdateLinksNever | 2 | 从不更新指定工作簿的嵌入式 OLE 链接。 |
| xlUpdateLinksUserSetting | 1 | 按照用户对指定工作簿的设置，更新嵌入式 OLE 链接。 |


#### XlWBATemplate 枚举

# [XlWBATemplate 枚举​](#xlwbatemplate-枚举)

指定要创建的工作簿的类型。新工作簿包含单个指定类型的工作表。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlWBATChart | -4109 | 图表 |
| xlWBATExcel4IntlMacroSheet | 4 | ET 版本 4 宏 |
| xlWBATExcel4MacroSheet | 3 | ET 版本 4 国际宏 |
| xlWBATWorksheet | -4167 | 工作表 |


#### XlWebFormatting 枚举

# [XlWebFormatting 枚举​](#xlwebformatting-枚举)

指定将网页导入查询表时应用网页格式（如果有）的程度。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlWebFormattingAll | 1 | 导入所有格式。 |
| xlWebFormattingNone | 3 | 不导入任何格式。 |
| xlWebFormattingRTF | 2 | 导入与 RTF 格式兼容的格式。 |


#### XlWebSelectionType 枚举

# [XlWebSelectionType 枚举​](#xlwebselectiontype-枚举)

指定是将整个网页、网页上的所有表还是特定表导入查询表。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlAllTables | 2 | 所有表 |
| xlEntirePage | 1 | 整页 |
| xlSpecifiedTables | 3 | 指定表 |


#### XlWindowState 枚举

# [XlWindowState 枚举​](#xlwindowstate-枚举)

指定窗口的状态。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlMaximized | -4137 | 最大化 |
| xlMinimized | -4140 | 最小化 |
| xlNormal | -4143 | 正常 |


#### XlWindowType 枚举

# [XlWindowType 枚举​](#xlwindowtype-枚举)

指定图表的显示方式。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlChartAsWindow | 5 | 图表将在新窗口中打开。 |
| xlChartInPlace | 4 | 图表将在当前工作表中显示。 |
| xlClipboard | 3 | 将图表复制到剪贴板。 |
| xlInfo | -4129 | 已放弃使用此常量。 |
| xlWorkbook | 1 | 此常量只适用于 Macintosh。 |


#### XlWindowView 枚举

# [XlWindowView 枚举​](#xlwindowview-枚举)

指定窗口中显示的视图。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlNormalView | 1 | 普通。 |
| xlPageBreakPreview | 2 | 分页预览。 |
| xlPageLayoutView | 3 | 页面视图。 |


#### XlXLMMacroType 枚举

# [XlXLMMacroType 枚举​](#xlxlmmacrotype-枚举)

指定在 ET 版本 4 宏工作表中，名称引用哪种宏，或名称是否引用宏。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlCommand | 2 | 自定义命令。 |
| xlFunction | 1 | 自定义函数。 |
| xlNotXLM | 3 | 非宏。 |


#### XlXmlExportResult 枚举

# [XlXmlExportResult 枚举​](#xlxmlexportresult-枚举)

指定保存或导出操作的结果。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlXmlExportSuccess | 0 | XML 数据文件已成功导出。 |
| xlXmlExportValidationFailed | 1 | XML 数据文件的内容不符合指定的架构映射。 |


#### XlXmlImportResult 枚举

# [XlXmlImportResult 枚举​](#xlxmlimportresult-枚举)

指定刷新或导入操作的结果。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlXmlImportElementsTruncated | 1 | 由于指定的 XML 数据文件对于工作表来说太大，因此其内容已被截断。 |
| xlXmlImportSuccess | 0 | XML 数据文件已成功导入。 |
| xlXmlImportValidationFailed | 2 | XML 数据文件的内容不符合指定的架构映射。 |


#### XlXmlLoadOption 枚举

# [XlXmlLoadOption 枚举​](#xlxmlloadoption-枚举)

指定 ET 打开 XML 数据文件的方式。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlXmlLoadImportToList | 2 | 将 XML 数据文件的内容置于 XML 表中。 |
| xlXmlLoadMapXml | 3 | 在**“XML 结构”**任务窗格中显示 XML 数据文件的架构。 |
| xlXmlLoadOpenXml | 1 | 打开 XML 数据文件。文件的内容将展开。 |
| xlXmlLoadPromptUser | 0 | 提示用户选择打开文件的方式。 |


#### XlYesNoGuess 枚举

# [XlYesNoGuess 枚举​](#xlyesnoguess-枚举)

指定第一行是否包含标题。不能在对数据透视表进行排序时使用。

| 名称 | 值 | 说明 |
| --- | --- | --- |
| xlGuess | 0 | ET 确定是否有标题，如果有，是否是一个。 |
| xlNo | 2 | 默认值。应对整个区域进行排序。 |
| xlYes | 1 | 不应对整个区域进行排序。 |


### 数据表

#### 待开放

# [🚧 待开放​](#🚧-待开放)

敬请期待


## 高级服务

### 云文档 API

# [云文档 API​](#云文档-api)

AirScript 提供全局的 KSDrive 对象，通过此对象即可轻松查看、修改和创建您的云文档

提示

在使用 KSDrive 对象操作云文档时，确保您已添加云文档API服务，在脚本编辑器的服务菜单内添加即可。

### [快速使用​](#快速使用)

js
```js
// 打开指定文档
let file = KSDrive.openFile('https://www.kdocs.cn/l/xxxxxxxxxxxx')
// 打印指定文档的A1单元格内容
console.log(file.Application.Range('A1').Text)
// 使用结束之后调用close关闭文档，否则无法再次调用KSDrive.openFile
file.close()
// 获取我的云文档下面的et，ksheet文档列表
const fileList = KSDrive.listFiles({ includeExts: ['et', 'ksheet'] })
// 打开我的云文档目录下的第一个文档
file = KSDrive.openFile(fileList.files[0])
console.log(file.Application.Range('A1').Text)
// 关闭文档
file.close()
```

### [属性列表​](#属性列表)

| 属性名 | 数据类型 | 说明 |
| --- | --- | --- |
| FileType | object | 支持的文件类型集合 |

### [方法列表​](#方法列表)

| 方法名 | 返回类型 | 说明 |
| --- | --- | --- |
| createFile() | string | 创建或另存一个文件 |
| openFile() | File | 额外打开一个文件 |
| listFiles() | FilesInfo | 列出某个目录下的表格文件 |

## [FileType​](#filetype)

云文档支持的文件类型，可用于新建文件时指定新文件的类型

### [属性说明​](#属性说明)

| 属性名 | 数据类型 | 说明 |
| --- | --- | --- |
| AP | string | 智能文档 |
| KSheet | string | 智能表格 |
| ET | string | 表格 |
| DB | string | 多维表 |

## [createFile()​](#createfile)

创建一个新文件，也可以将一个源文件另存为新文件

### [参数​](#参数)

| 名称 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| type | FileType | 是 | 新文件的类型 |
| createOptions | CreateOptions | 是 | 新文件的参数选项 |

### [CreateOptions 对象说明​](#createOptions)

| 名称 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| name | string | 是 | 新文件的文件名 |
| dirUrl | string | 否 | 新文件的文件目录 |
| source | string | 否 | 将目标文件另存为新文件 |

### [返回值​](#返回值)

url - string 新文件的 URL

### [示例​](#示例)

js
```js
// 创建ET文件，指定保存位置
let url = KSDrive.createFile(KSDrive.FileType.ET, {
  name: 'et测试',
  dirUrl: '指定保存位置'
})
console.log(url)
// 新建DB文件
url = KSDrive.createFile(KSDrive.FileType.DB)
console.log(url)
// 新建KSheet文件
url = KSDrive.createFile(KSDrive.FileType.KSheet)
console.log(url)
// 新建AP文件
url = KSDrive.createFile(KSDrive.FileType.AP)
console.log(url)
// 文件另存
url = KSDrive.createFile(KSDrive.FileType.KSheet, {
  source: 'https://www.kdocs.cn/l/cqQwuiG2mo7E',
  name: '复制表格'
})
console.log(url)
```

## [openFile()​](#openfile)

额外打开一个文件，并返回一个 JavaScript 对象File。

### [示例​](#示例-1)

js
```js
let file = KSDrive.openFile('https://www.kdocs.cn/l/xxxxxxxxxxxx')
console.log(file.Application.ActiveSheet.Range('A1').Text)
file.close()
```

### [参数​](#参数-1)

| 名称 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| openInfo | URL /FileInfo | 是 | 打开文件的信息，可以为文件分享链接或者FileInfo |

### [返回值​](#返回值-1)

File- 一个 JavaScript 对象

## [listFiles()​](#listfiles)

列出某个目录下的所有文件和对应信息

### [示例​](#示例-2)

js
```js
// 遍历获取某个文件夹下的所有文件的文件名
for (let offset = 0; offset >= 0; ) {
  const list = KSDrive.listFiles({
    dirUrl: 'https://www.kdocs.cn/mine/xxxxxxxxxx',
    offset: offset,
    count: 100
  })
  for (let i = 0; i < list.files.length; i++) {
    console.log(list.files[i].fileName)
  }
  offset = list.nextOffset
}
```

### [参数​](#参数-2)

| 名称 | 类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| options | object | undefined | 否 | 一个 JavaScript 对象，undefined 时获取我的云文档目录下面的文件数据，详细参数如下所示 |

### [详细参数​](#详细参数)

| 参数名 | 参数类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| dirUrl | string |  | false | 目录链接，如https://www.kdocs.cn/mine/xxxxxx，为空时获取我的云文档目录下面的文件数据 |
| offset | number | 0 | false | 开始位置。通常由listFiles()函数返回。比如，listFiles()函数在某次检索中返回了 nextOffset 为 100，而想要获取更多文件信息，则下一次调用listFiles()函数时把 100 作为此可选参数传入。 |
| count | number | 30 | false | 文件个数 |
| includeExts | string[] |  | false | 指定文件类型,支持参数及对应关系，ksheet:"表格",et:"WPS 表格",db:"多维表",otl:"文档",wpp:"演示",wps:"WPS 文字" |

### [返回值​](#返回值-2)

FilesInfo- 一个 JavaScript 对象，文件信息

## [File​](#file)

打开文件函数openFile()返回的一个 JavaScript 对象。

### [属性​](#属性)

| 名称 | 类型 | 说明 |
| --- | --- | --- |
| Application | Application(ET/Ksheet/DBT) | 被打开文件的操作对象，目前支持 et,ksheet,dbt |
| close | Function | 关闭文件的函数，使用完 file 对象之后调用，关闭打开的文件 |

## [FilesInfo​](#filesinfo)

获取文件夹信息函数listFiles(options)返回的一个 JavaScript 对象。

### [属性​](#属性-1)

| 名称 | 类型 | 说明 |
| --- | --- | --- |
| files | FileInfo[] | 文件信息，详细参数如下所示 |
| nextOffset | number | 下一页的偏移量，可以作为listFiles(options)的参数而输出下一页文件内容，当下一页为空时，nextOffset 为-1 |

### [FileInfo​](#fileinfo)

| 名称 | 类型 | 说明 |
| --- | --- | --- |
| fileName | string | 文件名 |
| fileId | string | 加密后的文件 id |
| createTime | number | 文件创建时间戳 |
| updateTime | number | 文件修改时间戳 |


### 概述

# [概述​](#概述)

借助AirScript的高级服务，开发者只需要完成较少设置，即可连接到某些公开的金山文档API。 它们的使用方式与AirScript脚本的内置函数十分相似。

AirScript在运行时会自动处理授权流程。 不过开发者必须启用高级服务，才能在脚本中使用该服务，若跳过该步骤，会因为找不到该服务而抛出undefined错误。

## [启用高级服务​](#启用高级服务)

要使用高级服务，请按以下说明操作：

打开
效率
-
AirScript编辑工具
弹出编辑页面。
点击AirScript编辑工具上方的
服务
。
点击
添加服务
。
选择一项服务，然后点击
确认
。
启用高级服务后，该服务会在自动补全中显示。

## [授权流程​](#授权流程)

AirScript需要用户授权才能访问高级服务中的私密数据。

### [授予运行权限​](#授予运行权限)

AirScript会根据开发者编写脚本时启用高级服务的配置内容来确定授权范围 （例如访问指定文件或访问网络）。如果脚本需要授权，用户在运行脚本时会弹出授权对话框。 描述这个脚本涉及到的授权范围。

普通的代码更改并不会清空用户对脚本的授权。但如果开发者对更改了高级服务的配置（新增，修改或删除）， 那用户对脚本的授权也会清空，再次运行脚本时会重新触发授权流程。

注意:我的脚本中的脚本的所有权完全归属于用户本身，该分类运行脚本时无需触发授权流程。

### [取消授权​](#取消授权)

用户可以对已授权的脚本手动取消授权，请按以下说明操作

打开
效率
-
AirScript编辑工具
弹出编辑页面。
找到文件共享脚本下的想取消授权的脚本，点击
…
显示更多操作。
点击
取消服务授权
## [使用限制​](#使用限制)

为防止向用户提供恶意的脚本，出于安全性考虑，使用高级服务存在一些限制。

过于高频地使用高级服务，当出现这种情况时，脚本的运行会抛出明显的错误通知用户异常调用。
使用
HTTP
服务时，禁止使用IP地址发起请求，禁止使用端口发起请求。
使用
HTTP
服务时，收到内容的消息体最大为2M，超过2M会抛出错误。
使用
KSDrive.openFile
获得的
File
对象没有调用close, 就再次使用
KSDrive.openFile
会报错。

### 网络 API

# [网络 API​](#网络-api)

AirScript 提供一个全局的 HTTP 对象，开发者可通过此对象提供的方法请求外部服务，请求成功后会同步返回服务器的响应。

该 API 的使用方式与浏览器内的 fetch()函数基本一致，对于前端开发者来说应该可以很快上手。

提示

在使用 HTTP 对象提供的方法发送请求之前，确保您已添加网络API服务，在脚本编辑器的【工具栏】-【服务】菜单内添加即可。

### [快速使用​](#快速使用)

javascript
```javascript
// 发起网络请求
const resp = HTTP.fetch('https://open.iciba.com/dsapi/', {
  timeout: 2000
})
const data = resp.json()
console.log(data.note, data.content)
```

### [方法列表​](#方法列表)

| 方法 | 返回类型 | 简介 |
| --- | --- | --- |
| fetch(url[, options]) | Response | 发起自定义类型的网络请求 |
| get(url[, options]) | Response | 发起 GET 类型的网络请求 |
| delete(url[, options]) | Response | 发起 DELETE 类型的网络请求 |
| post(url,body[, options]) | Response | 发起 POST 类型的网络请求 |
| put(url,body[, options]) | Response | 发起 PUT 类型的网络请求 |

## [fetch(url[, options])​](#fetch)

发起一个网络请求，可以自定义设置 headers 和 body。

### [示例​](#示例)

javascript
```javascript
const resp = HTTP.fetch('https://www.kdocs.cn', {
  method: 'GET',
  timeout: 2000,
  headers: {
    'User-Agent':
      'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/109.0.0.0 Safari/537.36'
  }
})
console.log(resp.text())
```

### [参数​](#参数)

| 名称 | 类型 | 默认值 | 必填项 | 说明 |
| --- | --- | --- | --- | --- |
| url | string |  | true | 需要访问的网络地址，只允许访问不带端口号的域名 |
| options | RequestOption | undefined | false | 一个 JavaScript 对象，可指定发起请求的可选参数，如下所示。 |

### [RequestOption​](#requestoption)

| 名称 | 类型 | 默认值 | 必填项 | 说明 |
| --- | --- | --- | --- | --- |
| method | string | GET | false | 发起网络请求的方法，例如GET、POST、PUT、DELETE等 |
| timeout | number | 10000 | false | 发起网络请求的超时时间，单位毫秒(ms)，数据范围为 0~60000，超出范围的数据将被设为默认值 10 秒。 |
| headers | object | undefined | false | 发起网络请求的头部。例如cookie等 |
| body | string | undefined | false | 发起网络请求的主体内容。 |

### [返回值​](#返回值)

Response- 服务器返回的响应

## [get(url[, options])​](#get)

发起 GET 类型的网络请求。

### [示例​](#示例-1)

javascript
```javascript
const resp = HTTP.get('https://reqres.in/api/users/2')
console.log(resp.json())
```

### [参数​](#参数-1)

| 名称 | 类型 | 默认值 | 必填项 | 说明 |
| --- | --- | --- | --- | --- |
| url | string |  | true | 需要访问的网络地址，只允许访问不带端口号的域名 |
| options | MethodRequestOption | undefined | false | 一个 JavaScript 对象，可指定特定请求的可选参数，如下所示。 |

### [MethodRequestOption​](#methodrequestoption)

| 名称 | 类型 | 默认值 | 必填项 | 说明 |
| --- | --- | --- | --- | --- |
| timeout | number | 10000 | false | 发起网络请求的超时时间，单位毫秒(ms)，数据范围为 0~60000，超出范围的数据将被设为默认值 10 秒。 |
| headers | object | undefined | false | 发起网络请求的头部。例如cookie等 |

### [返回值​](#返回值-1)

Response- 服务器返回的响应

## [delete(url[, options])​](#delete)

发起 DELETE 类型的网络请求。

### [示例​](#示例-2)

javascript
```javascript
const resp = HTTP.delete('https://reqres.in/api/users/2')
console.log(resp.status)
```

### [参数​](#参数-2)

| 名称 | 类型 | 默认值 | 必填项 | 说明 |
| --- | --- | --- | --- | --- |
| url | string |  | true | 需要访问的网络地址，只允许访问不带端口号的域名 |
| options | MethodRequestOption | undefined | false | 一个 JavaScript 对象，可指定特定请求的可选参数，如下所示。 |

### [返回值​](#返回值-2)

Response- 服务器返回的响应

## [post(url,body[, options])​](#post)

发起 POST 类型的网络请求。

### [示例​](#示例-3)

javascript
```javascript
// 发送form
const formResp = HTTP.post(
  'https://www.example.cn',
  { foo: 'bar' },
  { headers: { 'content-type': 'multipart/form-data' } }
)

//发送json
const resp = HTTP.post('https://reqres.in/api/users', {
  name: 'morpheus',
  job: 'leader'
})

console.log(resp.json())
```

### [参数​](#参数-3)

| 名称 | 类型 | 默认值 | 必填项 | 说明 |
| --- | --- | --- | --- | --- |
| url | string |  | true | 需要访问的网络地址，只允许访问不带端口号的域名 |
| body | string| object |  | true | 请求体 |
| options | MethodRequestOption | undefined | false | 一个 JavaScript 对象，可指定特定请求的可选参数，如下所示。 |

### [返回值​](#返回值-3)

Response- 服务器返回的响应

## [put(url,body[, options])​](#put)

发起 PUT 类型的网络请求。

### [示例​](#示例-4)

javascript
```javascript
const resp = HTTP.put('https://reqres.in/api/users/200', {
  name: 'wps',
  job: 'developer'
})
console.log(resp.json())
```

### [参数​](#参数-4)

| 名称 | 类型 | 默认值 | 必填项 | 说明 |
| --- | --- | --- | --- | --- |
| url | string |  | true | 需要访问的网络地址，只允许访问不带端口号的域名 |
| body | string| object |  | true | 请求体 |
| options | MethodRequestOption | undefined | false | 一个 JavaScript 对象，可指定特定请求的可选参数，如下所示。 |

### [返回值​](#返回值-4)

Response- 服务器返回的响应

## [Response​](#response)

HTTP 发起网络请求后返回的响应，response 是流数据，只有首次调用 text()，json()或 binary()能获取到数据

### [示例​](#示例-5)

javascript
```javascript
let resp = HTTP.get('https://open.iciba.com/dsapi/')
console.log(resp.status) // 200
console.log(resp.statusText) // OK
console.log(resp.text()) // `{foo:"bar"}`
console.log(resp.json()) // {foo:"bar"}
console.log(resp.status) // [...]
```

### [方法列表​](#方法列表-1)

| 方法 | 返回类型 | 简介 |
| --- | --- | --- |
| status | number | 获取响应的 HTTP 状态码 |
| statusText | string | 获取响应的 HTTP 状态 |
| headers | object | 获取响应的 header |
| text() | string | 获取服务器返回的文本 Body |
| json() | any | 将服务器返回的 json 类型的 Body 转化为结构体 |
| binary() | Buffer | 获取服务器返回的二进制结构的 Body |

## [status​](#status)

获取响应的 HTTP 状态码

### [示例​](#示例-6)

javascript
```javascript
const resp = HTTP.get('https://open.iciba.com/dsapi/')
console.log(resp.status) // 200
```

### [返回值​](#返回值-5)

number - 服务器返回响应的 HTTP 状态码

## [statusText​](#statustext)

获取响应的 HTTP 状态

### [示例​](#示例-7)

javascript
```javascript
const resp = HTTP.get('https://open.iciba.com/dsapi/')
console.log(resp.statusText) // OK
```

### [返回值​](#返回值-6)

string - 服务器返回响应的 HTTP 状态

## [headers​](#headers)

获取响应的 header

### [示例​](#示例-8)

javascript
```javascript
let resp = HTTP.get('https://open.iciba.com/dsapi/')
console.log(resp.headers) // {"content-length":"44","content-type":"text/html; charset=utf-8"}
```

### [返回值​](#返回值-7)

object - 服务器返回响应的 header

## [text()​](#text)

获取服务器返回的 Body

### [示例​](#示例-9)

javascript
```javascript
let resp = HTTP.get('https://open.iciba.com/dsapi/')
console.log(resp.text()) // this is an example.
```

### [返回值​](#返回值-8)

string - 服务器返回的响应的 Body，以文本接受并返回

## [json()​](#json)

获取服务器返回的 Body

### [示例​](#示例-10)

javascript
```javascript
let resp = HTTP.get('https://open.iciba.com/dsapi/')
console.log(resp.json()) // {msg:"this is an example."}
```

### [返回值​](#返回值-9)

Object, Array, string, number, boolean, or null - 服务器返回的响应的 Body，以文本接受并经过 JSON.parse()后返回

## [binary()​](#binary)

获取服务器返回的 Body

### [示例​](#示例-11)

javascript
```javascript
let resp = HTTP.get('https://open.iciba.com/dsapi/')
console.log(resp.binary().toString('base64'))
```

### [返回值​](#返回值-10)

Buffer- 服务器返回的响应的 Body，以 Buffer 接受二进制数据并返回

