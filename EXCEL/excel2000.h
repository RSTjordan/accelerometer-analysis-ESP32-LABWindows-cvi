#ifndef _EXCEL_H
#define _EXCEL_H

#include <cviauto.h>

#ifdef __cplusplus
    extern "C" {
#endif
/* NICDBLD_BEGIN> ActiveX Automation Properties */

/* NICDBLD_BEGIN> Object Name: _QueryTable */
#define Excel_QueryTableName                          0x0000006E
#define Excel_QueryTableApplication                   0x00000094
#define Excel_QueryTableCreator                       0x00000095
#define Excel_QueryTableParent                        0x00000096
#define Excel_QueryTableDestination                   0x000002A9
#define Excel_QueryTableSaveData                      0x000002B4
#define Excel_QueryTableHasAutoFormat                 0x000002B7
#define Excel_QueryTableRecordset                     0x0000048D
#define Excel_QueryTableBackgroundQuery               0x00000593
#define Excel_QueryTableConnection                    0x00000598
#define Excel_QueryTableEnableRefresh                 0x000005C5
#define Excel_QueryTableRefreshOnFileOpen             0x000005C7
#define Excel_QueryTableSql                           0x000005C8
#define Excel_QueryTableSavePassword                  0x000005C9
#define Excel_QueryTablePreserveFormatting            0x000005DC
#define Excel_QueryTableFieldNames                    0x00000630
#define Excel_QueryTableRowNumbers                    0x00000631
#define Excel_QueryTableFillAdjacentFormulas          0x00000632
#define Excel_QueryTableRefreshing                    0x00000633
#define Excel_QueryTableFetchedRowOverflow            0x00000634
#define Excel_QueryTableRefreshStyle                  0x00000636
#define Excel_QueryTablePostText                      0x00000637
#define Excel_QueryTableResultRange                   0x00000638
#define Excel_QueryTableParameters                    0x00000639
#define Excel_QueryTableTablesOnlyFromHTML            0x0000063A
#define Excel_QueryTableEnableEditing                 0x0000063B
#define Excel_QueryTableCommandText                   0x00000725
#define Excel_QueryTableCommandType                   0x00000726
#define Excel_QueryTableQueryType                     0x00000727
#define Excel_QueryTableMaintainConnection            0x00000728
#define Excel_QueryTableRefreshPeriod                 0x00000729
#define Excel_QueryTableTextFilePlatform              0x0000073F
#define Excel_QueryTableTextFileStartRow              0x00000740
#define Excel_QueryTableTextFileParseType             0x00000741
#define Excel_QueryTableTextFileTextQualifier         0x00000742
#define Excel_QueryTableTextFileConsecutiveDelimiter  0x00000743
#define Excel_QueryTableTextFileTabDelimiter          0x00000744
#define Excel_QueryTableTextFileSemicolonDelimiter    0x00000745
#define Excel_QueryTableTextFileCommaDelimiter        0x00000746
#define Excel_QueryTableTextFileSpaceDelimiter        0x00000747
#define Excel_QueryTableTextFileOtherDelimiter        0x00000748
#define Excel_QueryTableTextFileColumnDataTypes       0x00000749
#define Excel_QueryTableTextFileFixedColumnWidths     0x0000074A
#define Excel_QueryTablePreserveColumnInfo            0x0000074B
#define Excel_QueryTableAdjustColumnWidth             0x0000074C
#define Excel_QueryTableTextFilePromptOnRefresh       0x0000074D
#define Excel_QueryTableTextFileDecimalSeparator      0x0000074E
#define Excel_QueryTableTextFileThousandsSeparator    0x0000074F
#define Excel_QueryTableWebSelectionType              0x00000750
#define Excel_QueryTableWebFormatting                 0x00000751
#define Excel_QueryTableWebTables                     0x00000752
#define Excel_QueryTableWebPreFormattedTextToColumns  0x00000753
#define Excel_QueryTableWebSingleBlockTextImport      0x00000754
#define Excel_QueryTableWebDisableDateRecognition     0x00000755
#define Excel_QueryTableWebConsecutiveDelimitersAsOne 0x00000756
/* NICDBLD_END> Object Name: _QueryTable */

/* NICDBLD_BEGIN> Object Name: _Application */
#define Excel_App_Default                 0x00000000
#define Excel_AppUILanguage               0x00000002
#define Excel_AppValue                    0x00000006
#define Excel_AppName                     0x0000006E
#define Excel_AppCharts                   0x00000079
#define Excel_AppWidth                    0x0000007A
#define Excel_AppHeight                   0x0000007B
#define Excel_AppTop                      0x0000007E
#define Excel_AppLeft                     0x0000007F
#define Excel_AppCaption                  0x0000008B
#define Excel_AppSelection                0x00000093
#define Excel_AppApplication              0x00000094
#define Excel_AppCreator                  0x00000095
#define Excel_AppParent                   0x00000096
#define Excel_AppActiveChart              0x000000B7
#define Excel_AppDefaultSheetDirection    0x000000E5
#define Excel_AppCursorMovement           0x000000E8
#define Excel_AppControlCharacters        0x000000E9
#define Excel_AppCells                    0x000000EE
#define Excel_AppColumns                  0x000000F1
#define Excel_AppRows                     0x00000102
#define Excel_AppPath                     0x00000123
#define Excel_AppActiveCell               0x00000131
#define Excel_AppActivePrinter            0x00000132
#define Excel_AppActiveSheet              0x00000133
#define Excel_AppActiveWorkbook           0x00000134
#define Excel_AppTransitionMenuKey        0x00000136
#define Excel_AppTransitionMenuKeyAction  0x00000137
#define Excel_AppTransitionNavigKeys      0x00000138
#define Excel_AppAltStartupPath           0x00000139
#define Excel_AppBuild                    0x0000013A
#define Excel_AppCalculateBeforeSave      0x0000013B
#define Excel_AppCalculation              0x0000013C
#define Excel_AppCanPlaySounds            0x0000013E
#define Excel_AppCanRecordSounds          0x0000013F
#define Excel_AppCellDragAndDrop          0x00000140
#define Excel_AppDisplayClipboardWindow   0x00000142
#define Excel_AppCommandUnderlines        0x00000143
#define Excel_AppConstrainNumeric         0x00000144
#define Excel_AppCutCopyMode              0x0000014A
#define Excel_AppDataEntryMode            0x0000014B
#define Excel_AppDDEAppReturnCode         0x0000014C
#define Excel_AppDisplayAlerts            0x00000157
#define Excel_AppDisplayFormulaBar        0x00000158
#define Excel_AppDisplayNoteIndicator     0x00000159
#define Excel_AppDisplayScrollBars        0x0000015A
#define Excel_AppDisplayStatusBar         0x0000015B
#define Excel_AppFixedDecimal             0x0000015F
#define Excel_AppFixedDecimalPlaces       0x00000160
#define Excel_AppIgnoreRemoteRequests     0x00000164
#define Excel_AppInteractive              0x00000169
#define Excel_AppIteration                0x0000016B
#define Excel_AppLargeButtons             0x0000016C
#define Excel_AppColorButtons             0x0000016D
#define Excel_AppLibraryPath              0x0000016E
#define Excel_AppMathCoprocessorAvailable 0x0000016F
#define Excel_AppMaxChange                0x00000170
#define Excel_AppMaxIterations            0x00000171
#define Excel_AppMemoryFree               0x00000172
#define Excel_AppMemoryTotal              0x00000173
#define Excel_AppMemoryUsed               0x00000174
#define Excel_AppMouseAvailable           0x00000175
#define Excel_AppMoveAfterReturn          0x00000176
#define Excel_AppOperatingSystem          0x00000177
#define Excel_AppOrganizationName         0x00000178
#define Excel_AppPathSeparator            0x00000179
#define Excel_AppRecordRelative           0x0000017B
#define Excel_AppReferenceStyle           0x0000017C
#define Excel_AppTemplatesPath            0x0000017D
#define Excel_AppScreenUpdating           0x0000017E
#define Excel_AppStartupPath              0x00000181
#define Excel_AppStatusBar                0x00000182
#define Excel_AppShowToolTips             0x00000183
#define Excel_AppNetworkTemplatesPath     0x00000184
#define Excel_AppUsableHeight             0x00000185
#define Excel_AppUsableWidth              0x00000186
#define Excel_AppUserName                 0x00000187
#define Excel_AppVersion                  0x00000188
#define Excel_AppWindowsForPens           0x0000018B
#define Excel_AppWindowState              0x0000018C
#define Excel_AppWindows                  0x000001AE
#define Excel_AppNames                    0x000001BA
#define Excel_AppSheets                   0x000001E5
#define Excel_AppWorksheets               0x000001EE
#define Excel_AppAddIns                   0x00000225
#define Excel_AppToolbars                 0x00000228
#define Excel_AppVisible                  0x0000022E
#define Excel_AppWorkbooks                0x0000023C
#define Excel_AppExcel4MacroSheets        0x00000243
#define Excel_AppExcel4IntlMacroSheets    0x00000245
#define Excel_AppModules                  0x00000246
#define Excel_AppMenuBars                 0x0000024D
#define Excel_AppOnWindow                 0x0000026F
#define Excel_AppOnCalculate              0x00000271
#define Excel_AppOnEntry                  0x00000273
#define Excel_AppOnDoubleClick            0x00000274
#define Excel_AppOnData                   0x00000275
#define Excel_AppActiveMenuBar            0x000002F6
#define Excel_AppActiveWindow             0x000002F7
#define Excel_AppDialogs                  0x000002F9
#define Excel_AppDialogSheets             0x000002FC
#define Excel_AppThisWorkbook             0x0000030A
#define Excel_AppCustomListCount          0x00000313
#define Excel_AppActiveDialog             0x0000032F
#define Excel_AppStandardFont             0x0000039C
#define Excel_AppStandardFontSize         0x0000039D
#define Excel_AppDisplayRecentFiles       0x0000039E
#define Excel_AppDisplayExcel4Menus       0x0000039F
#define Excel_AppEditDirectlyInCell       0x000003A1
#define Excel_AppAlertBeforeOverwriting   0x000003A2
#define Excel_AppMailSession              0x000003AE
#define Excel_AppMailSystem               0x000003CB
#define Excel_AppCopyObjectsWithCells     0x000003DF
#define Excel_AppAskToUpdateLinks         0x000003E0
#define Excel_AppSheetsInNewWorkbook      0x000003E1
#define Excel_AppOnSheetActivate          0x00000407
#define Excel_AppDefaultFilePath          0x0000040E
#define Excel_AppDisplayFullScreen        0x00000425
#define Excel_AppPromptForSummaryInfo     0x00000426
#define Excel_AppEnableTipWizard          0x00000428
#define Excel_AppOnSheetDeactivate        0x00000439
#define Excel_AppEnableCancelKey          0x00000448
#define Excel_AppMoveAfterReturnDirection 0x00000478
#define Excel_AppAutoCorrect              0x00000479
#define Excel_AppCursor                   0x00000489
#define Excel_AppEnableAutoComplete       0x0000049B
#define Excel_AppEnableAnimations         0x0000049C
#define Excel_AppDisplayCommentIndicator  0x000004AC
#define Excel_AppEnableSound              0x000004AD
#define Excel_AppFileSearch               0x000004B0
#define Excel_AppFileFind                 0x000004B1
#define Excel_AppRecentFiles              0x000004B2
#define Excel_AppODBCErrors               0x000004B3
#define Excel_AppODBCTimeout              0x000004B4
#define Excel_AppPivotTableSelection      0x000004B5
#define Excel_AppRollZoom                 0x000004B6
#define Excel_AppShowChartTipNames        0x000004B7
#define Excel_AppShowChartTipValues       0x000004B8
#define Excel_AppDefaultSaveFormat        0x000004B9
#define Excel_AppUserControl              0x000004BA
#define Excel_AppVBE                      0x000004BB
#define Excel_AppEnableEvents             0x000004BC
#define Excel_AppDisplayInfoWindow        0x000004BD
#define Excel_AppAssistant                0x0000059E
#define Excel_AppCommandBars              0x0000059F
#define Excel_AppWorksheetFunction        0x000005A0
#define Excel_AppExtendList               0x00000701
#define Excel_AppOLEDBErrors              0x00000702
#define Excel_AppCOMAddIns                0x00000704
#define Excel_AppDefaultWebOptions        0x00000705
#define Excel_AppProductCode              0x00000706
#define Excel_AppUserLibraryPath          0x00000707
#define Excel_AppAutoPercentEntry         0x00000708
#define Excel_AppLanguageSettings         0x00000709
#define Excel_AppDummy101                 0x0000070A
#define Excel_AppAnswerWizard             0x0000070C
#define Excel_AppCalculationVersion       0x0000070E
#define Excel_AppShowWindowsInTaskbar     0x0000070F
#define Excel_AppFeatureInstall           0x00000710
/* NICDBLD_END> Object Name: _Application */

/* NICDBLD_BEGIN> Object Name: _Chart */
#define Excel_Chart_CodeName             0x80010000
#define Excel_ChartArea3DGroup           0x00000011
#define Excel_ChartBar3DGroup            0x00000012
#define Excel_ChartColumn3DGroup         0x00000013
#define Excel_ChartLine3DGroup           0x00000014
#define Excel_ChartPie3DGroup            0x00000015
#define Excel_ChartSurfaceGroup          0x00000016
#define Excel_ChartDepthPercent          0x00000030
#define Excel_ChartElevation             0x00000031
#define Excel_ChartGapDepth              0x00000032
#define Excel_ChartHasLegend             0x00000035
#define Excel_ChartHasTitle              0x00000036
#define Excel_ChartHeightPercent         0x00000037
#define Excel_ChartPerspective           0x00000039
#define Excel_ChartRightAngleAxes        0x0000003A
#define Excel_ChartRotation              0x0000003B
#define Excel_ChartCorners               0x0000004F
#define Excel_ChartChartArea             0x00000050
#define Excel_ChartChartTitle            0x00000051
#define Excel_ChartFloor                 0x00000053
#define Excel_ChartLegend                0x00000054
#define Excel_ChartPlotArea              0x00000055
#define Excel_ChartWalls                 0x00000056
#define Excel_ChartPlotVisibleOnly       0x0000005C
#define Excel_ChartDisplayBlanksAs       0x0000005D
#define Excel_ChartSizeWithWindow        0x0000005E
#define Excel_ChartAutoScaling           0x0000006B
#define Excel_ChartType                  0x0000006C
#define Excel_ChartSubType               0x0000006D
#define Excel_ChartName                  0x0000006E
#define Excel_ChartApplication           0x00000094
#define Excel_ChartCreator               0x00000095
#define Excel_ChartParent                0x00000096
#define Excel_ChartPlotBy                0x000000CA
#define Excel_ChartWallsAndGridlines2D   0x000000D2
#define Excel_ChartProtectContents       0x00000124
#define Excel_ChartProtectDrawingObjects 0x00000125
#define Excel_ChartIndex                 0x000001E6
#define Excel_ChartNext                  0x000001F6
#define Excel_ChartPrevious              0x000001F7
#define Excel_ChartVisible               0x0000022E
#define Excel_ChartOnDoubleClick         0x00000274
#define Excel_ChartPageSetup             0x000003E6
#define Excel_ChartOnSheetActivate       0x00000407
#define Excel_ChartOnSheetDeactivate     0x00000439
#define Excel_ChartProtectionMode        0x00000487
#define Excel_ChartCodeName              0x0000055D
#define Excel_ChartShapes                0x00000561
#define Excel_ChartHyperlinks            0x00000571
#define Excel_ChartDataTable             0x00000573
#define Excel_ChartHasDataTable          0x00000574
#define Excel_ChartShowWindow            0x00000577
#define Excel_ChartChartType             0x00000578
#define Excel_ChartBarShape              0x0000057B
#define Excel_ChartProtectFormatting     0x0000057D
#define Excel_ChartProtectData           0x0000057E
#define Excel_ChartProtectGoalSeek       0x0000057F
#define Excel_ChartProtectSelection      0x00000580
#define Excel_ChartPivotLayout           0x00000716
#define Excel_ChartHasPivotFields        0x00000717
#define Excel_ChartScripts               0x00000718
/* NICDBLD_END> Object Name: _Chart */

/* NICDBLD_BEGIN> Object Name: _Worksheet */
#define Excel_Worksheet_CodeName                  0x80010000
#define Excel_WorksheetOutline                    0x00000066
#define Excel_WorksheetType                       0x0000006C
#define Excel_WorksheetName                       0x0000006E
#define Excel_WorksheetApplication                0x00000094
#define Excel_WorksheetCreator                    0x00000095
#define Excel_WorksheetParent                     0x00000096
#define Excel_WorksheetCells                      0x000000EE
#define Excel_WorksheetColumns                    0x000000F1
#define Excel_WorksheetRows                       0x00000102
#define Excel_WorksheetProtectContents            0x00000124
#define Excel_WorksheetProtectDrawingObjects      0x00000125
#define Excel_WorksheetProtectScenarios           0x00000126
#define Excel_WorksheetTransitionExpEval          0x00000191
#define Excel_WorksheetTransitionFormEntry        0x00000192
#define Excel_WorksheetStandardHeight             0x00000197
#define Excel_WorksheetStandardWidth              0x00000198
#define Excel_WorksheetUsedRange                  0x0000019C
#define Excel_WorksheetNames                      0x000001BA
#define Excel_WorksheetIndex                      0x000001E6
#define Excel_WorksheetNext                       0x000001F6
#define Excel_WorksheetPrevious                   0x000001F7
#define Excel_WorksheetVisible                    0x0000022E
#define Excel_WorksheetComments                   0x0000023F
#define Excel_WorksheetOnCalculate                0x00000271
#define Excel_WorksheetOnEntry                    0x00000273
#define Excel_WorksheetOnDoubleClick              0x00000274
#define Excel_WorksheetOnData                     0x00000275
#define Excel_WorksheetDisplayAutomaticPageBreaks 0x00000283
#define Excel_Worksheet_DisplayRightToLeft        0x00000288
#define Excel_WorksheetConsolidationFunction      0x00000315
#define Excel_WorksheetConsolidationOptions       0x00000316
#define Excel_WorksheetConsolidationSources       0x00000317
#define Excel_WorksheetAutoFilterMode             0x00000318
#define Excel_WorksheetAutoFilter                 0x00000319
#define Excel_WorksheetFilterMode                 0x00000320
#define Excel_WorksheetPageSetup                  0x000003E6
#define Excel_WorksheetOnSheetActivate            0x00000407
#define Excel_WorksheetCircularReference          0x0000042D
#define Excel_WorksheetOnSheetDeactivate          0x00000439
#define Excel_WorksheetEnableAutoFilter           0x00000484
#define Excel_WorksheetEnableOutlining            0x00000485
#define Excel_WorksheetEnablePivotTable           0x00000486
#define Excel_WorksheetProtectionMode             0x00000487
#define Excel_WorksheetCodeName                   0x0000055D
#define Excel_WorksheetShapes                     0x00000561
#define Excel_WorksheetHyperlinks                 0x00000571
#define Excel_WorksheetHPageBreaks                0x0000058A
#define Excel_WorksheetVPageBreaks                0x0000058B
#define Excel_WorksheetEnableCalculation          0x00000590
#define Excel_WorksheetEnableSelection            0x00000591
#define Excel_WorksheetScrollArea                 0x00000599
#define Excel_WorksheetQueryTables                0x0000059A
#define Excel_WorksheetDisplayPageBreaks          0x0000059B
#define Excel_WorksheetDisplayRightToLeft         0x000006EE
#define Excel_WorksheetScripts                    0x00000718
/* NICDBLD_END> Object Name: _Worksheet */

/* NICDBLD_BEGIN> Object Name: _Workbook */
#define Excel_Workbook_CodeName                 0x80010000
#define Excel_WorkbookName                      0x0000006E
#define Excel_WorkbookCharts                    0x00000079
#define Excel_WorkbookApplication               0x00000094
#define Excel_WorkbookCreator                   0x00000095
#define Excel_WorkbookParent                    0x00000096
#define Excel_WorkbookActiveChart               0x000000B7
#define Excel_WorkbookTitle                     0x000000C7
#define Excel_WorkbookCreateBackup              0x0000011F
#define Excel_WorkbookFileFormat                0x00000120
#define Excel_WorkbookFullName                  0x00000121
#define Excel_WorkbookHasPassword               0x00000122
#define Excel_WorkbookPath                      0x00000123
#define Excel_WorkbookProtectWindows            0x00000127
#define Excel_WorkbookReadOnly                  0x00000128
#define Excel_WorkbookReadOnlyRecommended       0x00000129
#define Excel_WorkbookSaved                     0x0000012A
#define Excel_WorkbookWriteReserved             0x0000012B
#define Excel_WorkbookWriteReservedBy           0x0000012C
#define Excel_WorkbookActiveSheet               0x00000133
#define Excel_WorkbookDate1904                  0x00000193
#define Excel_WorkbookDisplayDrawingObjects     0x00000194
#define Excel_WorkbookPrecisionAsDisplayed      0x00000195
#define Excel_WorkbookSaveLinkValues            0x00000196
#define Excel_WorkbookUpdateRemoteReferences    0x0000019B
#define Excel_WorkbookWindows                   0x000001AE
#define Excel_WorkbookNames                     0x000001BA
#define Excel_WorkbookSheets                    0x000001E5
#define Excel_WorkbookStyles                    0x000001ED
#define Excel_WorkbookWorksheets                0x000001EE
#define Excel_WorkbookAuthor                    0x0000023E
#define Excel_WorkbookComments                  0x0000023F
#define Excel_WorkbookKeywords                  0x00000241
#define Excel_WorkbookExcel4MacroSheets         0x00000243
#define Excel_WorkbookExcel4IntlMacroSheets     0x00000245
#define Excel_WorkbookModules                   0x00000246
#define Excel_WorkbookProtectStructure          0x0000024C
#define Excel_WorkbookDialogSheets              0x000002FC
#define Excel_WorkbookRoutingSlip               0x000003B5
#define Excel_WorkbookHasRoutingSlip            0x000003B6
#define Excel_WorkbookRouted                    0x000003B7
#define Excel_WorkbookSubject                   0x000003B9
#define Excel_WorkbookHasMailer                 0x000003D0
#define Excel_WorkbookMailer                    0x000003D3
#define Excel_WorkbookOnSheetActivate           0x00000407
#define Excel_WorkbookOnSheetDeactivate         0x00000439
#define Excel_WorkbookMultiUserEditing          0x00000491
#define Excel_WorkbookShowConflictHistory       0x00000493
#define Excel_WorkbookRevisionNumber            0x00000494
#define Excel_WorkbookUserStatus                0x00000495
#define Excel_WorkbookConflictResolution        0x00000497
#define Excel_WorkbookBuiltinDocumentProperties 0x00000498
#define Excel_WorkbookCustomDocumentProperties  0x00000499
#define Excel_WorkbookOnSave                    0x0000049A
#define Excel_WorkbookContainer                 0x000004A6
#define Excel_WorkbookUserControl               0x000004BA
#define Excel_WorkbookCodeName                  0x0000055D
#define Excel_WorkbookCommandBars               0x0000059F
#define Excel_WorkbookAcceptLabelsInFormulas    0x000005A1
#define Excel_WorkbookAutoUpdateFrequency       0x000005A2
#define Excel_WorkbookAutoUpdateSaveChanges     0x000005A3
#define Excel_WorkbookChangeHistoryDuration     0x000005A4
#define Excel_WorkbookIsAddin                   0x000005A5
#define Excel_WorkbookPersonalViewListSettings  0x000005A7
#define Excel_WorkbookPersonalViewPrintSettings 0x000005A8
#define Excel_WorkbookCustomViews               0x000005B0
#define Excel_WorkbookTemplateRemoveExtData     0x000005B1
#define Excel_WorkbookHighlightChangesOnScreen  0x000005B5
#define Excel_WorkbookKeepChangeHistory         0x000005B6
#define Excel_WorkbookListChangesOnNewSheet     0x000005B7
#define Excel_WorkbookVBProject                 0x000005BD
#define Excel_WorkbookIsInplace                 0x000006E9
#define Excel_WorkbookCalculationVersion        0x0000070E
#define Excel_WorkbookPublishObjects            0x0000071B
#define Excel_WorkbookWebOptions                0x0000071C
#define Excel_WorkbookHTMLProject               0x0000071F
#define Excel_WorkbookEnvelopeVisible           0x00000720
#define Excel_WorkbookVBASigned                 0x00000724
/* NICDBLD_END> Object Name: _Workbook */

/* NICDBLD_BEGIN> Object Name: _OLEObject */
#define Excel_OLEObjectShadow          0x80010067
#define Excel_OLEObjectName            0x8001006E
#define Excel_OLEObjectWidth           0x8001007A
#define Excel_OLEObjectHeight          0x8001007B
#define Excel_OLEObjectTop             0x8001007E
#define Excel_OLEObjectLeft            0x8001007F
#define Excel_OLEObjectBorder          0x80010080
#define Excel_OLEObjectInterior        0x80010081
#define Excel_OLEObjectApplication     0x80010094
#define Excel_OLEObjectCreator         0x80010095
#define Excel_OLEObjectParent          0x80010096
#define Excel_OLEObjectLocked          0x8001010D
#define Excel_OLEObjectIndex           0x800101E6
#define Excel_OLEObjectVisible         0x8001022E
#define Excel_OLEObjectOnAction        0x80010254
#define Excel_OLEObjectEnabled         0x80010258
#define Excel_OLEObjectBottomRightCell 0x80010267
#define Excel_OLEObjectPlacement       0x80010269
#define Excel_OLEObjectPrintObject     0x8001026A
#define Excel_OLEObjectTopLeftCell     0x8001026C
#define Excel_OLEObjectZOrder          0x8001026E
#define Excel_OLEObjectSourceName      0x800102D1
#define Excel_OLEObjectListFillRange   0x8001034F
#define Excel_OLEObjectAutoUpdate      0x80010418
#define Excel_OLEObjectObject          0x80010419
#define Excel_OLEObjectOLEType         0x8001041E
#define Excel_OLEObjectLinkedCell      0x80010422
#define Excel_OLEObjectAutoLoad        0x800104A2
#define Excel_OLEObjectProgId          0x800105F3
#define Excel_OLEObjectShapeRange      0x800105F8
#define Excel_OLEObjectAltHTML         0x8001073D
/* NICDBLD_END> Object Name: _OLEObject */

/* NICDBLD_BEGIN> Object Name: Adjustments */
#define Excel_AdjustmentsParent      0x00000001
#define Excel_AdjustmentsCount       0x00000002
#define Excel_AdjustmentsApplication 0x60020000
#define Excel_AdjustmentsCreator     0x60020001
/* NICDBLD_END> Object Name: Adjustments */

/* NICDBLD_BEGIN> Object Name: CalloutFormat */
#define Excel_CalloutFormatParent      0x00000001
#define Excel_CalloutFormatAccent      0x00000064
#define Excel_CalloutFormatAngle       0x00000065
#define Excel_CalloutFormatAutoAttach  0x00000066
#define Excel_CalloutFormatAutoLength  0x00000067
#define Excel_CalloutFormatBorder      0x00000068
#define Excel_CalloutFormatDrop        0x00000069
#define Excel_CalloutFormatDropType    0x0000006A
#define Excel_CalloutFormatGap         0x0000006B
#define Excel_CalloutFormatLength      0x0000006C
#define Excel_CalloutFormatType        0x0000006D
#define Excel_CalloutFormatApplication 0x60020000
#define Excel_CalloutFormatCreator     0x60020001
/* NICDBLD_END> Object Name: CalloutFormat */

/* NICDBLD_BEGIN> Object Name: ColorFormat */
#define Excel_ColorFormatRGB         0x00000000
#define Excel_ColorFormatParent      0x00000001
#define Excel_ColorFormatSchemeColor 0x00000064
#define Excel_ColorFormatType        0x00000065
#define Excel_ColorFormatApplication 0x60020000
#define Excel_ColorFormatCreator     0x60020001
/* NICDBLD_END> Object Name: ColorFormat */

/* NICDBLD_BEGIN> Object Name: LineFormat */
#define Excel_LineFormatParent               0x00000001
#define Excel_LineFormatBackColor            0x00000064
#define Excel_LineFormatBeginArrowheadLength 0x00000065
#define Excel_LineFormatBeginArrowheadStyle  0x00000066
#define Excel_LineFormatBeginArrowheadWidth  0x00000067
#define Excel_LineFormatDashStyle            0x00000068
#define Excel_LineFormatEndArrowheadLength   0x00000069
#define Excel_LineFormatEndArrowheadStyle    0x0000006A
#define Excel_LineFormatEndArrowheadWidth    0x0000006B
#define Excel_LineFormatForeColor            0x0000006C
#define Excel_LineFormatPattern              0x0000006D
#define Excel_LineFormatStyle                0x0000006E
#define Excel_LineFormatTransparency         0x0000006F
#define Excel_LineFormatVisible              0x00000070
#define Excel_LineFormatWeight               0x00000071
#define Excel_LineFormatApplication          0x60020000
#define Excel_LineFormatCreator              0x60020001
/* NICDBLD_END> Object Name: LineFormat */

/* NICDBLD_BEGIN> Object Name: ShapeNode */
#define Excel_ShapeNodeParent      0x00000001
#define Excel_ShapeNodeEditingType 0x00000064
#define Excel_ShapeNodePoints      0x00000065
#define Excel_ShapeNodeSegmentType 0x00000066
#define Excel_ShapeNodeApplication 0x60020000
#define Excel_ShapeNodeCreator     0x60020001
/* NICDBLD_END> Object Name: ShapeNode */

/* NICDBLD_BEGIN> Object Name: ShapeNodes */
#define Excel_ShapeNodes_NewEnum    0xFFFFFFFC
#define Excel_ShapeNodesParent      0x00000001
#define Excel_ShapeNodesCount       0x00000002
#define Excel_ShapeNodesApplication 0x60020000
#define Excel_ShapeNodesCreator     0x60020001
/* NICDBLD_END> Object Name: ShapeNodes */

/* NICDBLD_BEGIN> Object Name: PictureFormat */
#define Excel_PictureFmtParent                0x00000001
#define Excel_PictureFmtBrightness            0x00000064
#define Excel_PictureFmtColorType             0x00000065
#define Excel_PictureFmtContrast              0x00000066
#define Excel_PictureFmtCropBottom            0x00000067
#define Excel_PictureFmtCropLeft              0x00000068
#define Excel_PictureFmtCropRight             0x00000069
#define Excel_PictureFmtCropTop               0x0000006A
#define Excel_PictureFmtTransparencyColor     0x0000006B
#define Excel_PictureFmtTransparentBackground 0x0000006C
#define Excel_PictureFmtApplication           0x60020000
#define Excel_PictureFmtCreator               0x60020001
/* NICDBLD_END> Object Name: PictureFormat */

/* NICDBLD_BEGIN> Object Name: ShadowFormat */
#define Excel_ShadowFormatParent       0x00000001
#define Excel_ShadowFormatForeColor    0x00000064
#define Excel_ShadowFormatObscured     0x00000065
#define Excel_ShadowFormatOffsetX      0x00000066
#define Excel_ShadowFormatOffsetY      0x00000067
#define Excel_ShadowFormatTransparency 0x00000068
#define Excel_ShadowFormatType         0x00000069
#define Excel_ShadowFormatVisible      0x0000006A
#define Excel_ShadowFormatApplication  0x60020000
#define Excel_ShadowFormatCreator      0x60020001
/* NICDBLD_END> Object Name: ShadowFormat */

/* NICDBLD_BEGIN> Object Name: TextEffectFormat */
#define Excel_TextEffectFmtParent           0x00000001
#define Excel_TextEffectFmtAlignment        0x00000064
#define Excel_TextEffectFmtFontBold         0x00000065
#define Excel_TextEffectFmtFontItalic       0x00000066
#define Excel_TextEffectFmtFontName         0x00000067
#define Excel_TextEffectFmtFontSize         0x00000068
#define Excel_TextEffectFmtKernedPairs      0x00000069
#define Excel_TextEffectFmtNormalizedHeight 0x0000006A
#define Excel_TextEffectFmtPresetShape      0x0000006B
#define Excel_TextEffectFmtPresetTextEffect 0x0000006C
#define Excel_TextEffectFmtRotatedChars     0x0000006D
#define Excel_TextEffectFmtText             0x0000006E
#define Excel_TextEffectFmtTracking         0x0000006F
#define Excel_TextEffectFmtApplication      0x60020000
#define Excel_TextEffectFmtCreator          0x60020001
/* NICDBLD_END> Object Name: TextEffectFormat */

/* NICDBLD_BEGIN> Object Name: ThreeDFormat */
#define Excel_ThreeDFmtParent                   0x00000001
#define Excel_ThreeDFmtDepth                    0x00000064
#define Excel_ThreeDFmtExtrusionColor           0x00000065
#define Excel_ThreeDFmtExtrusionColorType       0x00000066
#define Excel_ThreeDFmtPerspective              0x00000067
#define Excel_ThreeDFmtPresetExtrusionDirection 0x00000068
#define Excel_ThreeDFmtPresetLightingDirection  0x00000069
#define Excel_ThreeDFmtPresetLightingSoftness   0x0000006A
#define Excel_ThreeDFmtPresetMaterial           0x0000006B
#define Excel_ThreeDFmtPresetThreeDFormat       0x0000006C
#define Excel_ThreeDFmtRotationX                0x0000006D
#define Excel_ThreeDFmtRotationY                0x0000006E
#define Excel_ThreeDFmtVisible                  0x0000006F
#define Excel_ThreeDFmtApplication              0x60020000
#define Excel_ThreeDFmtCreator                  0x60020001
/* NICDBLD_END> Object Name: ThreeDFormat */

/* NICDBLD_BEGIN> Object Name: FillFormat */
#define Excel_FillFormatParent             0x00000001
#define Excel_FillFormatBackColor          0x00000064
#define Excel_FillFormatForeColor          0x00000065
#define Excel_FillFormatGradientColorType  0x00000066
#define Excel_FillFormatGradientDegree     0x00000067
#define Excel_FillFormatGradientStyle      0x00000068
#define Excel_FillFormatGradientVariant    0x00000069
#define Excel_FillFormatPattern            0x0000006A
#define Excel_FillFormatPresetGradientType 0x0000006B
#define Excel_FillFormatPresetTexture      0x0000006C
#define Excel_FillFormatTextureName        0x0000006D
#define Excel_FillFormatTextureType        0x0000006E
#define Excel_FillFormatTransparency       0x0000006F
#define Excel_FillFormatType               0x00000070
#define Excel_FillFormatVisible            0x00000071
#define Excel_FillFormatApplication        0x60020000
#define Excel_FillFormatCreator            0x60020001
/* NICDBLD_END> Object Name: FillFormat */

/* NICDBLD_BEGIN> Object Name: Sheets */
#define Excel_Sheets_NewEnum    0xFFFFFFFC
#define Excel_SheetsCount       0x00000076
#define Excel_SheetsApplication 0x00000094
#define Excel_SheetsCreator     0x00000095
#define Excel_SheetsParent      0x00000096
#define Excel_SheetsVisible     0x0000022E
#define Excel_SheetsHPageBreaks 0x0000058A
#define Excel_SheetsVPageBreaks 0x0000058B
/* NICDBLD_END> Object Name: Sheets */

/* NICDBLD_BEGIN> Object Name: Workbooks */
#define Excel_Workbooks_NewEnum    0xFFFFFFFC
#define Excel_WorkbooksCount       0x00000076
#define Excel_WorkbooksApplication 0x00000094
#define Excel_WorkbooksCreator     0x00000095
#define Excel_WorkbooksParent      0x00000096
/* NICDBLD_END> Object Name: Workbooks */

/* NICDBLD_BEGIN> Object Name: PublishObject */
#define Excel_PublishObjectApplication 0x00000094
#define Excel_PublishObjectCreator     0x00000095
#define Excel_PublishObjectParent      0x00000096
#define Excel_PublishObjectTitle       0x000000C7
#define Excel_PublishObjectSource      0x000000DE
#define Excel_PublishObjectSourceType  0x000002AD
#define Excel_PublishObjectSheet       0x000002EF
#define Excel_PublishObjectFilename    0x00000587
#define Excel_PublishObjectHtmlType    0x00000765
#define Excel_PublishObjectDivID       0x00000766
/* NICDBLD_END> Object Name: PublishObject */

/* NICDBLD_BEGIN> Object Name: DefaultWebOptions */
#define Excel_DefaultWebOptionsApplication                 0x00000094
#define Excel_DefaultWebOptionsCreator                     0x00000095
#define Excel_DefaultWebOptionsParent                      0x00000096
#define Excel_DefaultWebOptionsEncoding                    0x0000071E
#define Excel_DefaultWebOptionsRelyOnCSS                   0x0000076B
#define Excel_DefaultWebOptionsSaveHiddenData              0x0000076C
#define Excel_DefaultWebOptionsLoadPictures                0x0000076D
#define Excel_DefaultWebOptionsOrganizeInFolder            0x0000076E
#define Excel_DefaultWebOptionsUpdateLinksOnSave           0x0000076F
#define Excel_DefaultWebOptionsUseLongFileNames            0x00000770
#define Excel_DefaultWebOptionsCheckIfOfficeIsHTMLEditor   0x00000771
#define Excel_DefaultWebOptionsDownloadComponents          0x00000772
#define Excel_DefaultWebOptionsRelyOnVML                   0x00000773
#define Excel_DefaultWebOptionsAllowPNG                    0x00000774
#define Excel_DefaultWebOptionsScreenSize                  0x00000775
#define Excel_DefaultWebOptionsPixelsPerInch               0x00000776
#define Excel_DefaultWebOptionsLocationOfComponents        0x00000777
#define Excel_DefaultWebOptionsAlwaysSaveInDefaultEncoding 0x00000778
#define Excel_DefaultWebOptionsFonts                       0x00000779
#define Excel_DefaultWebOptionsFolderSuffix                0x0000077A
/* NICDBLD_END> Object Name: DefaultWebOptions */

/* NICDBLD_BEGIN> Object Name: WebOptions */
#define Excel_WebOptionsApplication          0x00000094
#define Excel_WebOptionsCreator              0x00000095
#define Excel_WebOptionsParent               0x00000096
#define Excel_WebOptionsEncoding             0x0000071E
#define Excel_WebOptionsRelyOnCSS            0x0000076B
#define Excel_WebOptionsOrganizeInFolder     0x0000076E
#define Excel_WebOptionsUseLongFileNames     0x00000770
#define Excel_WebOptionsDownloadComponents   0x00000772
#define Excel_WebOptionsRelyOnVML            0x00000773
#define Excel_WebOptionsAllowPNG             0x00000774
#define Excel_WebOptionsScreenSize           0x00000775
#define Excel_WebOptionsPixelsPerInch        0x00000776
#define Excel_WebOptionsLocationOfComponents 0x00000777
#define Excel_WebOptionsFolderSuffix         0x0000077A
/* NICDBLD_END> Object Name: WebOptions */

/* NICDBLD_BEGIN> Object Name: TreeviewControl */
#define Excel_TreeviewControlApplication 0x00000094
#define Excel_TreeviewControlCreator     0x00000095
#define Excel_TreeviewControlParent      0x00000096
#define Excel_TreeviewControlHidden      0x0000010C
#define Excel_TreeviewControlDrilled     0x0000077D
/* NICDBLD_END> Object Name: TreeviewControl */

/* NICDBLD_BEGIN> Object Name: CubeField */
#define Excel_CubeFieldValue           0x00000006
#define Excel_CubeFieldName            0x0000006E
#define Excel_CubeFieldPosition        0x00000085
#define Excel_CubeFieldOrientation     0x00000086
#define Excel_CubeFieldCaption         0x0000008B
#define Excel_CubeFieldApplication     0x00000094
#define Excel_CubeFieldCreator         0x00000095
#define Excel_CubeFieldParent          0x00000096
#define Excel_CubeFieldDragToColumn    0x000005E4
#define Excel_CubeFieldDragToHide      0x000005E5
#define Excel_CubeFieldDragToPage      0x000005E6
#define Excel_CubeFieldDragToRow       0x000005E7
#define Excel_CubeFieldDragToData      0x00000734
#define Excel_CubeFieldCubeFieldType   0x0000077E
#define Excel_CubeFieldTreeviewControl 0x0000077F
#define Excel_CubeFieldHiddenLevels    0x00000780
/* NICDBLD_END> Object Name: CubeField */

/* NICDBLD_BEGIN> Object Name: CubeFields */
#define Excel_CubeFields_NewEnum    0xFFFFFFFC
#define Excel_CubeFieldsCount       0x00000076
#define Excel_CubeFieldsApplication 0x00000094
#define Excel_CubeFieldsCreator     0x00000095
#define Excel_CubeFieldsParent      0x00000096
/* NICDBLD_END> Object Name: CubeFields */

/* NICDBLD_BEGIN> Object Name: Font */
#define Excel_FontBold          0x00000060
#define Excel_FontColorIndex    0x00000061
#define Excel_FontColor         0x00000063
#define Excel_FontItalic        0x00000065
#define Excel_FontShadow        0x00000067
#define Excel_FontSize          0x00000068
#define Excel_FontStrikethrough 0x00000069
#define Excel_FontUnderline     0x0000006A
#define Excel_FontName          0x0000006E
#define Excel_FontApplication   0x00000094
#define Excel_FontCreator       0x00000095
#define Excel_FontParent        0x00000096
#define Excel_FontFontStyle     0x000000B1
#define Excel_FontSuperscript   0x000000B2
#define Excel_FontSubscript     0x000000B3
#define Excel_FontBackground    0x000000B4
#define Excel_FontOutlineFont   0x000000DD
/* NICDBLD_END> Object Name: Font */

/* NICDBLD_BEGIN> Object Name: Window */
#define Excel_WindowType                       0x0000006C
#define Excel_WindowWidth                      0x0000007A
#define Excel_WindowHeight                     0x0000007B
#define Excel_WindowTop                        0x0000007E
#define Excel_WindowLeft                       0x0000007F
#define Excel_WindowCaption                    0x0000008B
#define Excel_WindowSelection                  0x00000093
#define Excel_WindowApplication                0x00000094
#define Excel_WindowCreator                    0x00000095
#define Excel_WindowParent                     0x00000096
#define Excel_WindowActiveChart                0x000000B7
#define Excel_WindowActiveCell                 0x00000131
#define Excel_WindowActiveSheet                0x00000133
#define Excel_WindowUsableHeight               0x00000185
#define Excel_WindowUsableWidth                0x00000186
#define Excel_WindowWindowState                0x0000018C
#define Excel_WindowIndex                      0x000001E6
#define Excel_WindowVisible                    0x0000022E
#define Excel_WindowOnWindow                   0x0000026F
#define Excel_WindowActivePane                 0x00000282
#define Excel_WindowDisplayFormulas            0x00000284
#define Excel_WindowDisplayGridlines           0x00000285
#define Excel_WindowDisplayHeadings            0x00000286
#define Excel_WindowDisplayOutline             0x00000287
#define Excel_Window_DisplayRightToLeft        0x00000288
#define Excel_WindowDisplayZeros               0x00000289
#define Excel_WindowFreezePanes                0x0000028A
#define Excel_WindowGridlineColor              0x0000028B
#define Excel_WindowGridlineColorIndex         0x0000028C
#define Excel_WindowPanes                      0x0000028D
#define Excel_WindowScrollColumn               0x0000028E
#define Excel_WindowScrollRow                  0x0000028F
#define Excel_WindowSelectedSheets             0x00000290
#define Excel_WindowSplit                      0x00000291
#define Excel_WindowSplitColumn                0x00000292
#define Excel_WindowSplitHorizontal            0x00000293
#define Excel_WindowSplitRow                   0x00000294
#define Excel_WindowSplitVertical              0x00000295
#define Excel_WindowZoom                       0x00000297
#define Excel_WindowTabRatio                   0x000002A1
#define Excel_WindowDisplayHorizontalScrollBar 0x00000399
#define Excel_WindowDisplayVerticalScrollBar   0x0000039A
#define Excel_WindowDisplayWorkbookTabs        0x0000039B
#define Excel_WindowVisibleRange               0x0000045E
#define Excel_WindowWindowNumber               0x0000045F
#define Excel_WindowRangeSelection             0x000004A5
#define Excel_WindowEnableResize               0x000004A8
#define Excel_WindowView                       0x000004AA
#define Excel_WindowDisplayRightToLeft         0x000006EE
/* NICDBLD_END> Object Name: Window */

/* NICDBLD_BEGIN> Object Name: Windows */
#define Excel_Windows_NewEnum    0xFFFFFFFC
#define Excel_WindowsCount       0x00000076
#define Excel_WindowsApplication 0x00000094
#define Excel_WindowsCreator     0x00000095
#define Excel_WindowsParent      0x00000096
/* NICDBLD_END> Object Name: Windows */

/* NICDBLD_BEGIN> Object Name: WorksheetFunction */
#define Excel_WorksheetFunctionApplication 0x00000094
#define Excel_WorksheetFunctionCreator     0x00000095
#define Excel_WorksheetFunctionParent      0x00000096
/* NICDBLD_END> Object Name: WorksheetFunction */

/* NICDBLD_BEGIN> Object Name: Range */
#define Excel_Range_NewEnum            0xFFFFFFFC
#define Excel_RangeValue               0x00000006
#define Excel_RangeName                0x0000006E
#define Excel_RangeCount               0x00000076
#define Excel_RangeWidth               0x0000007A
#define Excel_RangeHeight              0x0000007B
#define Excel_RangeTop                 0x0000007E
#define Excel_RangeLeft                0x0000007F
#define Excel_RangeInterior            0x00000081
#define Excel_RangeOrientation         0x00000086
#define Excel_RangeHorizontalAlignment 0x00000088
#define Excel_RangeVerticalAlignment   0x00000089
#define Excel_RangeText                0x0000008A
#define Excel_RangeFont                0x00000092
#define Excel_RangeApplication         0x00000094
#define Excel_RangeCreator             0x00000095
#define Excel_RangeParent              0x00000096
#define Excel_RangeNumberFormat        0x000000C1
#define Excel_RangeIndentLevel         0x000000C9
#define Excel_RangeMergeCells          0x000000D0
#define Excel_RangeShrinkToFit         0x000000D1
#define Excel_RangeCells               0x000000EE
#define Excel_RangeColumn              0x000000F0
#define Excel_RangeColumns             0x000000F1
#define Excel_RangeColumnWidth         0x000000F2
#define Excel_RangeCurrentRegion       0x000000F3
#define Excel_RangeEntireColumn        0x000000F6
#define Excel_RangeEntireRow           0x000000F7
#define Excel_RangePageBreak           0x000000FF
#define Excel_RangeRow                 0x00000101
#define Excel_RangeRows                0x00000102
#define Excel_RangeStyle               0x00000104
#define Excel_RangeFormula             0x00000105
#define Excel_RangeFormulaHidden       0x00000106
#define Excel_RangeFormulaLocal        0x00000107
#define Excel_RangeFormulaR1C1         0x00000108
#define Excel_RangeFormulaR1C1Local    0x00000109
#define Excel_RangeHasArray            0x0000010A
#define Excel_RangeHasFormula          0x0000010B
#define Excel_RangeHidden              0x0000010C
#define Excel_RangeLocked              0x0000010D
#define Excel_RangeOutlineLevel        0x0000010F
#define Excel_RangeRowHeight           0x00000110
#define Excel_RangeSummary             0x00000111
#define Excel_RangeUseStandardHeight   0x00000112
#define Excel_RangeUseStandardWidth    0x00000113
#define Excel_RangeWrapText            0x00000114
#define Excel_RangeWorksheet           0x0000015C
#define Excel_RangeBorders             0x000001B3
#define Excel_RangeCurrentArray        0x000001F5
#define Excel_RangeNext                0x000001F6
#define Excel_RangePrevious            0x000001F7
#define Excel_RangePrefixCharacter     0x000001F8
#define Excel_RangeDependents          0x0000021F
#define Excel_RangePrecedents          0x00000220
#define Excel_RangeDirectDependents    0x00000221
#define Excel_RangeDirectPrecedents    0x00000222
#define Excel_RangeAreas               0x00000238
#define Excel_RangeShowDetail          0x00000249
#define Excel_RangeFormulaArray        0x0000024A
#define Excel_RangeLocationInTable     0x000002B3
#define Excel_RangePivotTable          0x000002CC
#define Excel_RangePivotField          0x000002DB
#define Excel_RangePivotItem           0x000002E4
#define Excel_RangeComment             0x0000038E
#define Excel_RangeSoundNote           0x00000394
#define Excel_RangeReadingOrder        0x000003CF
#define Excel_RangeAddIndent           0x00000427
#define Excel_RangeNumberFormatLocal   0x00000449
#define Excel_RangeListHeaderRows      0x000004A3
#define Excel_RangeFormulaLabel        0x00000564
#define Excel_RangeMergeArea           0x00000569
#define Excel_RangeQueryTable          0x0000056A
#define Excel_RangeValidation          0x0000056B
#define Excel_RangeValue2              0x0000056C
#define Excel_RangePhonetic            0x0000056F
#define Excel_RangeFormatConditions    0x00000570
#define Excel_RangeHyperlinks          0x00000571
#define Excel_RangePhonetics           0x00000713
#define Excel_RangeID                  0x00000715
/* NICDBLD_END> Object Name: Range */

/* NICDBLD_BEGIN> Object Name: VPageBreak */
#define Excel_VPageBreakType        0x0000006C
#define Excel_VPageBreakApplication 0x00000094
#define Excel_VPageBreakCreator     0x00000095
#define Excel_VPageBreakParent      0x00000096
#define Excel_VPageBreakLocation    0x00000575
#define Excel_VPageBreakExtent      0x0000058E
/* NICDBLD_END> Object Name: VPageBreak */

/* NICDBLD_BEGIN> Object Name: HPageBreak */
#define Excel_HPageBreakType        0x0000006C
#define Excel_HPageBreakApplication 0x00000094
#define Excel_HPageBreakCreator     0x00000095
#define Excel_HPageBreakParent      0x00000096
#define Excel_HPageBreakLocation    0x00000575
#define Excel_HPageBreakExtent      0x0000058E
/* NICDBLD_END> Object Name: HPageBreak */

/* NICDBLD_BEGIN> Object Name: HPageBreaks */
#define Excel_HPageBreaks_NewEnum    0xFFFFFFFC
#define Excel_HPageBreaksCount       0x00000076
#define Excel_HPageBreaksApplication 0x00000094
#define Excel_HPageBreaksCreator     0x00000095
#define Excel_HPageBreaksParent      0x00000096
/* NICDBLD_END> Object Name: HPageBreaks */

/* NICDBLD_BEGIN> Object Name: VPageBreaks */
#define Excel_VPageBreaks_NewEnum    0xFFFFFFFC
#define Excel_VPageBreaksCount       0x00000076
#define Excel_VPageBreaksApplication 0x00000094
#define Excel_VPageBreaksCreator     0x00000095
#define Excel_VPageBreaksParent      0x00000096
/* NICDBLD_END> Object Name: VPageBreaks */

/* NICDBLD_BEGIN> Object Name: RecentFile */
#define Excel_RecentFileName        0x0000006E
#define Excel_RecentFileApplication 0x00000094
#define Excel_RecentFileCreator     0x00000095
#define Excel_RecentFileParent      0x00000096
#define Excel_RecentFilePath        0x00000123
#define Excel_RecentFileIndex       0x000001E6
/* NICDBLD_END> Object Name: RecentFile */

/* NICDBLD_BEGIN> Object Name: RecentFiles */
#define Excel_RecentFiles_NewEnum    0xFFFFFFFC
#define Excel_RecentFilesCount       0x00000076
#define Excel_RecentFilesApplication 0x00000094
#define Excel_RecentFilesCreator     0x00000095
#define Excel_RecentFilesParent      0x00000096
#define Excel_RecentFilesMaximum     0x0000058F
/* NICDBLD_END> Object Name: RecentFiles */

/* NICDBLD_BEGIN> Object Name: Style */
#define Excel_Style_Default            0x00000000
#define Excel_StyleValue               0x00000006
#define Excel_StyleName                0x0000006E
#define Excel_StyleInterior            0x00000081
#define Excel_StyleOrientation         0x00000086
#define Excel_StyleHorizontalAlignment 0x00000088
#define Excel_StyleVerticalAlignment   0x00000089
#define Excel_StyleFont                0x00000092
#define Excel_StyleApplication         0x00000094
#define Excel_StyleCreator             0x00000095
#define Excel_StyleParent              0x00000096
#define Excel_StyleNumberFormat        0x000000C1
#define Excel_StyleIndentLevel         0x000000C9
#define Excel_StyleMergeCells          0x000000D0
#define Excel_StyleShrinkToFit         0x000000D1
#define Excel_StyleFormulaHidden       0x00000106
#define Excel_StyleLocked              0x0000010D
#define Excel_StyleWrapText            0x00000114
#define Excel_StyleIncludeAlignment    0x0000019D
#define Excel_StyleIncludeBorder       0x0000019E
#define Excel_StyleIncludeFont         0x0000019F
#define Excel_StyleIncludeNumber       0x000001A0
#define Excel_StyleIncludePatterns     0x000001A1
#define Excel_StyleIncludeProtection   0x000001A2
#define Excel_StyleBorders             0x000001B3
#define Excel_StyleBuiltIn             0x00000229
#define Excel_StyleNameLocal           0x000003A9
#define Excel_StyleReadingOrder        0x000003CF
#define Excel_StyleAddIndent           0x00000427
#define Excel_StyleNumberFormatLocal   0x00000449
/* NICDBLD_END> Object Name: Style */

/* NICDBLD_BEGIN> Object Name: Styles */
#define Excel_Styles_NewEnum    0xFFFFFFFC
#define Excel_StylesCount       0x00000076
#define Excel_StylesApplication 0x00000094
#define Excel_StylesCreator     0x00000095
#define Excel_StylesParent      0x00000096
/* NICDBLD_END> Object Name: Styles */

/* NICDBLD_BEGIN> Object Name: Borders */
#define Excel_Borders_NewEnum    0xFFFFFFFC
#define Excel_BordersValue       0x00000006
#define Excel_BordersColorIndex  0x00000061
#define Excel_BordersColor       0x00000063
#define Excel_BordersCount       0x00000076
#define Excel_BordersLineStyle   0x00000077
#define Excel_BordersWeight      0x00000078
#define Excel_BordersApplication 0x00000094
#define Excel_BordersCreator     0x00000095
#define Excel_BordersParent      0x00000096
/* NICDBLD_END> Object Name: Borders */

/* NICDBLD_BEGIN> Object Name: AddIn */
#define Excel_AddInName        0x0000006E
#define Excel_AddInApplication 0x00000094
#define Excel_AddInCreator     0x00000095
#define Excel_AddInParent      0x00000096
#define Excel_AddInTitle       0x000000C7
#define Excel_AddInFullName    0x00000121
#define Excel_AddInPath        0x00000123
#define Excel_AddInInstalled   0x00000226
#define Excel_AddInAuthor      0x0000023E
#define Excel_AddInComments    0x0000023F
#define Excel_AddInKeywords    0x00000241
#define Excel_AddInSubject     0x000003B9
/* NICDBLD_END> Object Name: AddIn */

/* NICDBLD_BEGIN> Object Name: AddIns */
#define Excel_AddIns_NewEnum    0xFFFFFFFC
#define Excel_AddInsCount       0x00000076
#define Excel_AddInsApplication 0x00000094
#define Excel_AddInsCreator     0x00000095
#define Excel_AddInsParent      0x00000096
/* NICDBLD_END> Object Name: AddIns */

/* NICDBLD_BEGIN> Object Name: Areas */
#define Excel_Areas_NewEnum    0xFFFFFFFC
#define Excel_AreasCount       0x00000076
#define Excel_AreasApplication 0x00000094
#define Excel_AreasCreator     0x00000095
#define Excel_AreasParent      0x00000096
/* NICDBLD_END> Object Name: Areas */

/* NICDBLD_BEGIN> Object Name: Charts */
#define Excel_Charts_NewEnum    0xFFFFFFFC
#define Excel_ChartsCount       0x00000076
#define Excel_ChartsApplication 0x00000094
#define Excel_ChartsCreator     0x00000095
#define Excel_ChartsParent      0x00000096
#define Excel_ChartsVisible     0x0000022E
#define Excel_ChartsHPageBreaks 0x0000058A
#define Excel_ChartsVPageBreaks 0x0000058B
/* NICDBLD_END> Object Name: Charts */

/* NICDBLD_BEGIN> Object Name: PivotCache */
#define Excel_PivotCacheApplication        0x00000094
#define Excel_PivotCacheCreator            0x00000095
#define Excel_PivotCacheParent             0x00000096
#define Excel_PivotCacheMemoryUsed         0x00000174
#define Excel_PivotCacheIndex              0x000001E6
#define Excel_PivotCacheSourceData         0x000002AE
#define Excel_PivotCacheRefreshDate        0x000002B8
#define Excel_PivotCacheRefreshName        0x000002B9
#define Excel_PivotCacheRecordset          0x0000048D
#define Excel_PivotCacheBackgroundQuery    0x00000593
#define Excel_PivotCacheOptimizeCache      0x00000594
#define Excel_PivotCacheConnection         0x00000598
#define Excel_PivotCacheEnableRefresh      0x000005C5
#define Excel_PivotCacheRecordCount        0x000005C6
#define Excel_PivotCacheRefreshOnFileOpen  0x000005C7
#define Excel_PivotCacheSql                0x000005C8
#define Excel_PivotCacheSavePassword       0x000005C9
#define Excel_PivotCacheCommandText        0x00000725
#define Excel_PivotCacheCommandType        0x00000726
#define Excel_PivotCacheQueryType          0x00000727
#define Excel_PivotCacheMaintainConnection 0x00000728
#define Excel_PivotCacheRefreshPeriod      0x00000729
#define Excel_PivotCacheLocalConnection    0x0000072B
#define Excel_PivotCacheUseLocalConnection 0x0000072D
/* NICDBLD_END> Object Name: PivotCache */

/* NICDBLD_BEGIN> Object Name: PivotCaches */
#define Excel_PivotCachesCount       0x00000076
#define Excel_PivotCachesApplication 0x00000094
#define Excel_PivotCachesCreator     0x00000095
#define Excel_PivotCachesParent      0x00000096
/* NICDBLD_END> Object Name: PivotCaches */

/* NICDBLD_BEGIN> Object Name: PivotFormula */
#define Excel_PivotFormula_Default    0x00000000
#define Excel_PivotFormulaValue       0x00000006
#define Excel_PivotFormulaApplication 0x00000094
#define Excel_PivotFormulaCreator     0x00000095
#define Excel_PivotFormulaParent      0x00000096
#define Excel_PivotFormulaFormula     0x00000105
#define Excel_PivotFormulaIndex       0x000001E6
/* NICDBLD_END> Object Name: PivotFormula */

/* NICDBLD_BEGIN> Object Name: PivotFormulas */
#define Excel_PivotFormulasCount       0x00000076
#define Excel_PivotFormulasApplication 0x00000094
#define Excel_PivotFormulasCreator     0x00000095
#define Excel_PivotFormulasParent      0x00000096
/* NICDBLD_END> Object Name: PivotFormulas */

/* NICDBLD_BEGIN> Object Name: PivotTable */
#define Excel_PivotTable_Default                     0x00000000
#define Excel_PivotTableValue                        0x00000006
#define Excel_PivotTableName                         0x0000006E
#define Excel_PivotTableApplication                  0x00000094
#define Excel_PivotTableCreator                      0x00000095
#define Excel_PivotTableParent                       0x00000096
#define Excel_PivotTableSourceData                   0x000002AE
#define Excel_PivotTableSaveData                     0x000002B4
#define Excel_PivotTableRowGrand                     0x000002B5
#define Excel_PivotTableColumnGrand                  0x000002B6
#define Excel_PivotTableHasAutoFormat                0x000002B7
#define Excel_PivotTableRefreshDate                  0x000002B8
#define Excel_PivotTableRefreshName                  0x000002B9
#define Excel_PivotTableInnerDetail                  0x000002BA
#define Excel_PivotTableTableRange1                  0x000002BB
#define Excel_PivotTableTableRange2                  0x000002BC
#define Excel_PivotTableRowRange                     0x000002BD
#define Excel_PivotTableColumnRange                  0x000002BE
#define Excel_PivotTablePageRange                    0x000002BF
#define Excel_PivotTableDataLabelRange               0x000002C0
#define Excel_PivotTableDataBodyRange                0x000002C1
#define Excel_PivotTablePageFieldOrder               0x00000595
#define Excel_PivotTablePageFieldWrapCount           0x00000596
#define Excel_PivotTablePageRangeCells               0x000005CA
#define Excel_PivotTableCacheIndex                   0x000005CB
#define Excel_PivotTableDisplayErrorString           0x000005CD
#define Excel_PivotTableDisplayNullString            0x000005CE
#define Excel_PivotTableEnableDrilldown              0x000005CF
#define Excel_PivotTableEnableFieldDialog            0x000005D0
#define Excel_PivotTableEnableWizard                 0x000005D1
#define Excel_PivotTableErrorString                  0x000005D2
#define Excel_PivotTableManualUpdate                 0x000005D5
#define Excel_PivotTableMergeLabels                  0x000005D6
#define Excel_PivotTableNullString                   0x000005D7
#define Excel_PivotTablePivotFormulas                0x000005D9
#define Excel_PivotTableSubtotalHiddenPageItems      0x000005DA
#define Excel_PivotTablePageFieldStyle               0x000005DB
#define Excel_PivotTablePreserveFormatting           0x000005DC
#define Excel_PivotTablePivotSelection               0x000005DE
#define Excel_PivotTableSelectionMode                0x000005DF
#define Excel_PivotTableTableStyle                   0x000005E0
#define Excel_PivotTableTag                          0x000005E1
#define Excel_PivotTableVacatedStyle                 0x000005E2
#define Excel_PivotTablePrintTitles                  0x0000072E
#define Excel_PivotTableCubeFields                   0x0000072F
#define Excel_PivotTableGrandTotalName               0x00000730
#define Excel_PivotTableSmallGrid                    0x00000731
#define Excel_PivotTableRepeatItemsOnEachPrintedPage 0x00000732
#define Excel_PivotTableTotalsAnnotation             0x00000733
/* NICDBLD_END> Object Name: PivotTable */

/* NICDBLD_BEGIN> Object Name: PivotTables */
#define Excel_PivotTablesCount       0x00000076
#define Excel_PivotTablesApplication 0x00000094
#define Excel_PivotTablesCreator     0x00000095
#define Excel_PivotTablesParent      0x00000096
/* NICDBLD_END> Object Name: PivotTables */

/* NICDBLD_BEGIN> Object Name: PivotField */
#define Excel_PivotField_Default               0x00000000
#define Excel_PivotFieldValue                  0x00000006
#define Excel_PivotFieldName                   0x0000006E
#define Excel_PivotFieldPosition               0x00000085
#define Excel_PivotFieldOrientation            0x00000086
#define Excel_PivotFieldCaption                0x0000008B
#define Excel_PivotFieldApplication            0x00000094
#define Excel_PivotFieldCreator                0x00000095
#define Excel_PivotFieldParent                 0x00000096
#define Excel_PivotFieldNumberFormat           0x000000C1
#define Excel_PivotFieldFormula                0x00000105
#define Excel_PivotFieldCalculation            0x0000013C
#define Excel_PivotFieldMemoryUsed             0x00000174
#define Excel_PivotFieldShowAllItems           0x000001C4
#define Excel_PivotFieldLabelRange             0x000002CF
#define Excel_PivotFieldDataRange              0x000002D0
#define Excel_PivotFieldSourceName             0x000002D1
#define Excel_PivotFieldDataType               0x000002D2
#define Excel_PivotFieldGroupLevel             0x000002D3
#define Excel_PivotFieldTotalLevels            0x000002D4
#define Excel_PivotFieldParentField            0x000002DC
#define Excel_PivotFieldBaseField              0x000002DE
#define Excel_PivotFieldBaseItem               0x000002DF
#define Excel_PivotFieldChildField             0x000002E0
#define Excel_PivotFieldCurrentPage            0x000002E2
#define Excel_PivotFieldFunction               0x00000383
#define Excel_PivotFieldDragToColumn           0x000005E4
#define Excel_PivotFieldDragToHide             0x000005E5
#define Excel_PivotFieldDragToPage             0x000005E6
#define Excel_PivotFieldDragToRow              0x000005E7
#define Excel_PivotFieldIsCalculated           0x000005E8
#define Excel_PivotFieldServerBased            0x000005E9
#define Excel_PivotFieldAutoSortOrder          0x000005EC
#define Excel_PivotFieldAutoSortField          0x000005ED
#define Excel_PivotFieldAutoShowType           0x000005EE
#define Excel_PivotFieldAutoShowRange          0x000005EF
#define Excel_PivotFieldAutoShowCount          0x000005F0
#define Excel_PivotFieldAutoShowField          0x000005F1
#define Excel_PivotFieldDragToData             0x00000734
#define Excel_PivotFieldLayoutBlankLine        0x00000735
#define Excel_PivotFieldLayoutSubtotalLocation 0x00000736
#define Excel_PivotFieldLayoutPageBreak        0x00000737
#define Excel_PivotFieldLayoutForm             0x00000738
#define Excel_PivotFieldSubtotalName           0x00000739
#define Excel_PivotFieldDrilledDown            0x0000073A
#define Excel_PivotFieldCubeField              0x0000073B
#define Excel_PivotFieldCurrentPageName        0x0000073C
/* NICDBLD_END> Object Name: PivotField */

/* NICDBLD_BEGIN> Object Name: PivotFields */
#define Excel_PivotFieldsCount       0x00000076
#define Excel_PivotFieldsApplication 0x00000094
#define Excel_PivotFieldsCreator     0x00000095
#define Excel_PivotFieldsParent      0x00000096
/* NICDBLD_END> Object Name: PivotFields */

/* NICDBLD_BEGIN> Object Name: CalculatedFields */
#define Excel_CalculatedFieldsCount       0x00000076
#define Excel_CalculatedFieldsApplication 0x00000094
#define Excel_CalculatedFieldsCreator     0x00000095
#define Excel_CalculatedFieldsParent      0x00000096
/* NICDBLD_END> Object Name: CalculatedFields */

/* NICDBLD_BEGIN> Object Name: PivotItem */
#define Excel_PivotItem_Default         0x00000000
#define Excel_PivotItemValue            0x00000006
#define Excel_PivotItemName             0x0000006E
#define Excel_PivotItemPosition         0x00000085
#define Excel_PivotItemCaption          0x0000008B
#define Excel_PivotItemApplication      0x00000094
#define Excel_PivotItemCreator          0x00000095
#define Excel_PivotItemParent           0x00000096
#define Excel_PivotItemFormula          0x00000105
#define Excel_PivotItemVisible          0x0000022E
#define Excel_PivotItemShowDetail       0x00000249
#define Excel_PivotItemLabelRange       0x000002CF
#define Excel_PivotItemDataRange        0x000002D0
#define Excel_PivotItemSourceName       0x000002D1
#define Excel_PivotItemParentShowDetail 0x000002E3
#define Excel_PivotItemParentItem       0x000002E5
#define Excel_PivotItemRecordCount      0x000005C6
#define Excel_PivotItemIsCalculated     0x000005E8
#define Excel_PivotItemDrilledDown      0x0000073A
/* NICDBLD_END> Object Name: PivotItem */

/* NICDBLD_BEGIN> Object Name: PivotItems */
#define Excel_PivotItemsCount       0x00000076
#define Excel_PivotItemsApplication 0x00000094
#define Excel_PivotItemsCreator     0x00000095
#define Excel_PivotItemsParent      0x00000096
/* NICDBLD_END> Object Name: PivotItems */

/* NICDBLD_BEGIN> Object Name: CalculatedItems */
#define Excel_CalculatedItemsCount       0x00000076
#define Excel_CalculatedItemsApplication 0x00000094
#define Excel_CalculatedItemsCreator     0x00000095
#define Excel_CalculatedItemsParent      0x00000096
/* NICDBLD_END> Object Name: CalculatedItems */

/* NICDBLD_BEGIN> Object Name: Characters */
#define Excel_CharactersCount              0x00000076
#define Excel_CharactersText               0x0000008A
#define Excel_CharactersCaption            0x0000008B
#define Excel_CharactersFont               0x00000092
#define Excel_CharactersApplication        0x00000094
#define Excel_CharactersCreator            0x00000095
#define Excel_CharactersParent             0x00000096
#define Excel_CharactersPhoneticCharacters 0x000005F2
/* NICDBLD_END> Object Name: Characters */

/* NICDBLD_BEGIN> Object Name: Dialogs */
#define Excel_Dialogs_NewEnum    0xFFFFFFFC
#define Excel_DialogsCount       0x00000076
#define Excel_DialogsApplication 0x00000094
#define Excel_DialogsCreator     0x00000095
#define Excel_DialogsParent      0x00000096
/* NICDBLD_END> Object Name: Dialogs */

/* NICDBLD_BEGIN> Object Name: Dialog */
#define Excel_DialogApplication 0x00000094
#define Excel_DialogCreator     0x00000095
#define Excel_DialogParent      0x00000096
/* NICDBLD_END> Object Name: Dialog */

/* NICDBLD_BEGIN> Object Name: SoundNote */
#define Excel_SoundNoteApplication 0x00000094
#define Excel_SoundNoteCreator     0x00000095
#define Excel_SoundNoteParent      0x00000096
/* NICDBLD_END> Object Name: SoundNote */

/* NICDBLD_BEGIN> Object Name: Panes */
#define Excel_PanesCount       0x00000076
#define Excel_PanesApplication 0x00000094
#define Excel_PanesCreator     0x00000095
#define Excel_PanesParent      0x00000096
/* NICDBLD_END> Object Name: Panes */

/* NICDBLD_BEGIN> Object Name: Pane */
#define Excel_PaneApplication  0x00000094
#define Excel_PaneCreator      0x00000095
#define Excel_PaneParent       0x00000096
#define Excel_PaneIndex        0x000001E6
#define Excel_PaneScrollColumn 0x0000028E
#define Excel_PaneScrollRow    0x0000028F
#define Excel_PaneVisibleRange 0x0000045E
/* NICDBLD_END> Object Name: Pane */

/* NICDBLD_BEGIN> Object Name: Scenarios */
#define Excel_ScenariosCount       0x00000076
#define Excel_ScenariosApplication 0x00000094
#define Excel_ScenariosCreator     0x00000095
#define Excel_ScenariosParent      0x00000096
/* NICDBLD_END> Object Name: Scenarios */

/* NICDBLD_BEGIN> Object Name: Scenario */
#define Excel_ScenarioName          0x0000006E
#define Excel_ScenarioApplication   0x00000094
#define Excel_ScenarioCreator       0x00000095
#define Excel_ScenarioParent        0x00000096
#define Excel_ScenarioHidden        0x0000010C
#define Excel_ScenarioLocked        0x0000010D
#define Excel_ScenarioIndex         0x000001E6
#define Excel_ScenarioComment       0x0000038E
#define Excel_ScenarioChangingCells 0x0000038F
/* NICDBLD_END> Object Name: Scenario */

/* NICDBLD_BEGIN> Object Name: OLEObjects */
#define Excel_OLEObjectsShadow      0x00000067
#define Excel_OLEObjectsCount       0x00000076
#define Excel_OLEObjectsWidth       0x0000007A
#define Excel_OLEObjectsHeight      0x0000007B
#define Excel_OLEObjectsTop         0x0000007E
#define Excel_OLEObjectsLeft        0x0000007F
#define Excel_OLEObjectsBorder      0x00000080
#define Excel_OLEObjectsInterior    0x00000081
#define Excel_OLEObjectsApplication 0x00000094
#define Excel_OLEObjectsCreator     0x00000095
#define Excel_OLEObjectsParent      0x00000096
#define Excel_OLEObjectsLocked      0x0000010D
#define Excel_OLEObjectsVisible     0x0000022E
#define Excel_OLEObjectsOnAction    0x00000254
#define Excel_OLEObjectsEnabled     0x00000258
#define Excel_OLEObjectsPlacement   0x00000269
#define Excel_OLEObjectsPrintObject 0x0000026A
#define Excel_OLEObjectsZOrder      0x0000026E
#define Excel_OLEObjectsSourceName  0x000002D1
#define Excel_OLEObjectsAutoLoad    0x000004A2
#define Excel_OLEObjectsShapeRange  0x000005F8
/* NICDBLD_END> Object Name: OLEObjects */

/* NICDBLD_BEGIN> Object Name: RoutingSlip */
#define Excel_RoutingSlipApplication    0x00000094
#define Excel_RoutingSlipCreator        0x00000095
#define Excel_RoutingSlipParent         0x00000096
#define Excel_RoutingSlipSubject        0x000003B9
#define Excel_RoutingSlipMessage        0x000003BA
#define Excel_RoutingSlipDelivery       0x000003BB
#define Excel_RoutingSlipReturnWhenDone 0x000003BC
#define Excel_RoutingSlipTrackStatus    0x000003BD
#define Excel_RoutingSlipStatus         0x000003BE
/* NICDBLD_END> Object Name: RoutingSlip */

/* NICDBLD_BEGIN> Object Name: Outline */
#define Excel_OutlineApplication     0x00000094
#define Excel_OutlineCreator         0x00000095
#define Excel_OutlineParent          0x00000096
#define Excel_OutlineSummaryRow      0x00000386
#define Excel_OutlineAutomaticStyles 0x000003BF
#define Excel_OutlineSummaryColumn   0x000003C1
/* NICDBLD_END> Object Name: Outline */

/* NICDBLD_BEGIN> Object Name: Worksheets */
#define Excel_Worksheets_NewEnum    0xFFFFFFFC
#define Excel_WorksheetsCount       0x00000076
#define Excel_WorksheetsApplication 0x00000094
#define Excel_WorksheetsCreator     0x00000095
#define Excel_WorksheetsParent      0x00000096
#define Excel_WorksheetsVisible     0x0000022E
#define Excel_WorksheetsHPageBreaks 0x0000058A
#define Excel_WorksheetsVPageBreaks 0x0000058B
/* NICDBLD_END> Object Name: Worksheets */

/* NICDBLD_BEGIN> Object Name: PageSetup */
#define Excel_PageSetupOrientation        0x00000086
#define Excel_PageSetupApplication        0x00000094
#define Excel_PageSetupCreator            0x00000095
#define Excel_PageSetupParent             0x00000096
#define Excel_PageSetupOrder              0x000000C0
#define Excel_PageSetupZoom               0x00000297
#define Excel_PageSetupLeftMargin         0x000003E7
#define Excel_PageSetupRightMargin        0x000003E8
#define Excel_PageSetupTopMargin          0x000003E9
#define Excel_PageSetupBottomMargin       0x000003EA
#define Excel_PageSetupPrintHeadings      0x000003EB
#define Excel_PageSetupPrintGridlines     0x000003EC
#define Excel_PageSetupCenterHorizontally 0x000003ED
#define Excel_PageSetupCenterVertically   0x000003EE
#define Excel_PageSetupPaperSize          0x000003EF
#define Excel_PageSetupFirstPageNumber    0x000003F0
#define Excel_PageSetupBlackAndWhite      0x000003F1
#define Excel_PageSetupCenterFooter       0x000003F2
#define Excel_PageSetupCenterHeader       0x000003F3
#define Excel_PageSetupChartSize          0x000003F4
#define Excel_PageSetupFitToPagesTall     0x000003F5
#define Excel_PageSetupFitToPagesWide     0x000003F6
#define Excel_PageSetupFooterMargin       0x000003F7
#define Excel_PageSetupHeaderMargin       0x000003F8
#define Excel_PageSetupLeftFooter         0x000003F9
#define Excel_PageSetupLeftHeader         0x000003FA
#define Excel_PageSetupPrintArea          0x000003FB
#define Excel_PageSetupDraft              0x000003FC
#define Excel_PageSetupPrintNotes         0x000003FD
#define Excel_PageSetupPrintTitleColumns  0x000003FF
#define Excel_PageSetupPrintTitleRows     0x00000400
#define Excel_PageSetupRightFooter        0x00000401
#define Excel_PageSetupRightHeader        0x00000402
#define Excel_PageSetupPrintComments      0x000005F4
/* NICDBLD_END> Object Name: PageSetup */

/* NICDBLD_BEGIN> Object Name: Names */
#define Excel_Names_NewEnum    0xFFFFFFFC
#define Excel_NamesCount       0x00000076
#define Excel_NamesApplication 0x00000094
#define Excel_NamesCreator     0x00000095
#define Excel_NamesParent      0x00000096
/* NICDBLD_END> Object Name: Names */

/* NICDBLD_BEGIN> Object Name: Name */
#define Excel_Name_Default          0x00000000
#define Excel_NameValue             0x00000006
#define Excel_NameName              0x0000006E
#define Excel_NameApplication       0x00000094
#define Excel_NameCreator           0x00000095
#define Excel_NameParent            0x00000096
#define Excel_NameIndex             0x000001E6
#define Excel_NameVisible           0x0000022E
#define Excel_NameShortcutKey       0x00000255
#define Excel_NameCategory          0x000003A6
#define Excel_NameCategoryLocal     0x000003A7
#define Excel_NameMacroType         0x000003A8
#define Excel_NameNameLocal         0x000003A9
#define Excel_NameRefersTo          0x000003AA
#define Excel_NameRefersToLocal     0x000003AB
#define Excel_NameRefersToR1C1      0x000003AC
#define Excel_NameRefersToR1C1Local 0x000003AD
#define Excel_NameRefersToRange     0x00000488
/* NICDBLD_END> Object Name: Name */

/* NICDBLD_BEGIN> Object Name: ChartObject */
#define Excel_ChartObjectChart              0x00000007
#define Excel_ChartObjectShadow             0x00000067
#define Excel_ChartObjectName               0x0000006E
#define Excel_ChartObjectWidth              0x0000007A
#define Excel_ChartObjectHeight             0x0000007B
#define Excel_ChartObjectTop                0x0000007E
#define Excel_ChartObjectLeft               0x0000007F
#define Excel_ChartObjectBorder             0x00000080
#define Excel_ChartObjectInterior           0x00000081
#define Excel_ChartObjectApplication        0x00000094
#define Excel_ChartObjectCreator            0x00000095
#define Excel_ChartObjectParent             0x00000096
#define Excel_ChartObjectLocked             0x0000010D
#define Excel_ChartObjectIndex              0x000001E6
#define Excel_ChartObjectVisible            0x0000022E
#define Excel_ChartObjectOnAction           0x00000254
#define Excel_ChartObjectEnabled            0x00000258
#define Excel_ChartObjectBottomRightCell    0x00000267
#define Excel_ChartObjectPlacement          0x00000269
#define Excel_ChartObjectPrintObject        0x0000026A
#define Excel_ChartObjectRoundedCorners     0x0000026B
#define Excel_ChartObjectTopLeftCell        0x0000026C
#define Excel_ChartObjectZOrder             0x0000026E
#define Excel_ChartObjectShapeRange         0x000005F8
#define Excel_ChartObjectProtectChartObject 0x000005F9
/* NICDBLD_END> Object Name: ChartObject */

/* NICDBLD_BEGIN> Object Name: ChartObjects */
#define Excel_ChartObjectsShadow         0x00000067
#define Excel_ChartObjectsCount          0x00000076
#define Excel_ChartObjectsWidth          0x0000007A
#define Excel_ChartObjectsHeight         0x0000007B
#define Excel_ChartObjectsTop            0x0000007E
#define Excel_ChartObjectsLeft           0x0000007F
#define Excel_ChartObjectsBorder         0x00000080
#define Excel_ChartObjectsInterior       0x00000081
#define Excel_ChartObjectsApplication    0x00000094
#define Excel_ChartObjectsCreator        0x00000095
#define Excel_ChartObjectsParent         0x00000096
#define Excel_ChartObjectsLocked         0x0000010D
#define Excel_ChartObjectsVisible        0x0000022E
#define Excel_ChartObjectsOnAction       0x00000254
#define Excel_ChartObjectsEnabled        0x00000258
#define Excel_ChartObjectsPlacement      0x00000269
#define Excel_ChartObjectsPrintObject    0x0000026A
#define Excel_ChartObjectsRoundedCorners 0x0000026B
#define Excel_ChartObjectsShapeRange     0x000005F8
/* NICDBLD_END> Object Name: ChartObjects */

/* NICDBLD_BEGIN> Object Name: Mailer */
#define Excel_MailerApplication   0x00000094
#define Excel_MailerCreator       0x00000095
#define Excel_MailerParent        0x00000096
#define Excel_MailerSubject       0x000003B9
#define Excel_MailerWhichAddress  0x000003CE
#define Excel_MailerToRecipients  0x000003D5
#define Excel_MailerCCRecipients  0x000003D6
#define Excel_MailerBCCRecipients 0x000003D7
#define Excel_MailerEnclosures    0x000003D8
#define Excel_MailerReceived      0x000003DA
#define Excel_MailerSendDateTime  0x000003DB
#define Excel_MailerSender        0x000003DC
/* NICDBLD_END> Object Name: Mailer */

/* NICDBLD_BEGIN> Object Name: CustomViews */
#define Excel_CustomViews_NewEnum    0xFFFFFFFC
#define Excel_CustomViewsCount       0x00000076
#define Excel_CustomViewsApplication 0x00000094
#define Excel_CustomViewsCreator     0x00000095
#define Excel_CustomViewsParent      0x00000096
/* NICDBLD_END> Object Name: CustomViews */

/* NICDBLD_BEGIN> Object Name: CustomView */
#define Excel_CustomViewName           0x0000006E
#define Excel_CustomViewApplication    0x00000094
#define Excel_CustomViewCreator        0x00000095
#define Excel_CustomViewParent         0x00000096
#define Excel_CustomViewPrintSettings  0x00000629
#define Excel_CustomViewRowColSettings 0x0000062A
/* NICDBLD_END> Object Name: CustomView */

/* NICDBLD_BEGIN> Object Name: FormatConditions */
#define Excel_FormatConditions_NewEnum    0xFFFFFFFC
#define Excel_FormatConditionsCount       0x00000076
#define Excel_FormatConditionsApplication 0x00000094
#define Excel_FormatConditionsCreator     0x00000095
#define Excel_FormatConditionsParent      0x00000096
/* NICDBLD_END> Object Name: FormatConditions */

/* NICDBLD_BEGIN> Object Name: FormatCondition */
#define Excel_FormatConditionType        0x0000006C
#define Excel_FormatConditionInterior    0x00000081
#define Excel_FormatConditionFont        0x00000092
#define Excel_FormatConditionApplication 0x00000094
#define Excel_FormatConditionCreator     0x00000095
#define Excel_FormatConditionParent      0x00000096
#define Excel_FormatConditionBorders     0x000001B3
#define Excel_FormatConditionOperator    0x0000031D
#define Excel_FormatConditionFormula1    0x0000062B
#define Excel_FormatConditionFormula2    0x0000062C
/* NICDBLD_END> Object Name: FormatCondition */

/* NICDBLD_BEGIN> Object Name: Comments */
#define Excel_Comments_NewEnum    0xFFFFFFFC
#define Excel_CommentsCount       0x00000076
#define Excel_CommentsApplication 0x00000094
#define Excel_CommentsCreator     0x00000095
#define Excel_CommentsParent      0x00000096
/* NICDBLD_END> Object Name: Comments */

/* NICDBLD_BEGIN> Object Name: Comment */
#define Excel_CommentApplication 0x00000094
#define Excel_CommentCreator     0x00000095
#define Excel_CommentParent      0x00000096
#define Excel_CommentVisible     0x0000022E
#define Excel_CommentAuthor      0x0000023E
#define Excel_CommentShape       0x0000062E
/* NICDBLD_END> Object Name: Comment */

/* NICDBLD_BEGIN> Object Name: QueryTables */
#define Excel_QueryTablesCount       0x00000076
#define Excel_QueryTablesApplication 0x00000094
#define Excel_QueryTablesCreator     0x00000095
#define Excel_QueryTablesParent      0x00000096
/* NICDBLD_END> Object Name: QueryTables */

/* NICDBLD_BEGIN> Object Name: Parameter */
#define Excel_ParameterValue           0x00000006
#define Excel_ParameterType            0x0000006C
#define Excel_ParameterName            0x0000006E
#define Excel_ParameterApplication     0x00000094
#define Excel_ParameterCreator         0x00000095
#define Excel_ParameterParent          0x00000096
#define Excel_ParameterDataType        0x000002D2
#define Excel_ParameterPromptString    0x0000063F
#define Excel_ParameterSourceRange     0x00000640
#define Excel_ParameterRefreshOnChange 0x00000757
/* NICDBLD_END> Object Name: Parameter */

/* NICDBLD_BEGIN> Object Name: Parameters */
#define Excel_ParametersCount       0x00000076
#define Excel_ParametersApplication 0x00000094
#define Excel_ParametersCreator     0x00000095
#define Excel_ParametersParent      0x00000096
/* NICDBLD_END> Object Name: Parameters */

/* NICDBLD_BEGIN> Object Name: ODBCError */
#define Excel_ODBCErrorApplication 0x00000094
#define Excel_ODBCErrorCreator     0x00000095
#define Excel_ODBCErrorParent      0x00000096
#define Excel_ODBCErrorErrorString 0x000005D2
#define Excel_ODBCErrorSqlState    0x00000643
/* NICDBLD_END> Object Name: ODBCError */

/* NICDBLD_BEGIN> Object Name: ODBCErrors */
#define Excel_ODBCErrorsCount       0x00000076
#define Excel_ODBCErrorsApplication 0x00000094
#define Excel_ODBCErrorsCreator     0x00000095
#define Excel_ODBCErrorsParent      0x00000096
/* NICDBLD_END> Object Name: ODBCErrors */

/* NICDBLD_BEGIN> Object Name: Validation */
#define Excel_ValidationValue          0x00000006
#define Excel_ValidationType           0x0000006C
#define Excel_ValidationApplication    0x00000094
#define Excel_ValidationCreator        0x00000095
#define Excel_ValidationParent         0x00000096
#define Excel_ValidationOperator       0x0000031D
#define Excel_ValidationFormula1       0x0000062B
#define Excel_ValidationFormula2       0x0000062C
#define Excel_ValidationAlertStyle     0x00000645
#define Excel_ValidationIgnoreBlank    0x00000646
#define Excel_ValidationIMEMode        0x00000647
#define Excel_ValidationInCellDropdown 0x00000648
#define Excel_ValidationErrorMessage   0x00000649
#define Excel_ValidationErrorTitle     0x0000064A
#define Excel_ValidationInputMessage   0x0000064B
#define Excel_ValidationInputTitle     0x0000064C
#define Excel_ValidationShowError      0x0000064D
#define Excel_ValidationShowInput      0x0000064E
/* NICDBLD_END> Object Name: Validation */

/* NICDBLD_BEGIN> Object Name: Hyperlinks */
#define Excel_Hyperlinks_NewEnum    0xFFFFFFFC
#define Excel_HyperlinksCount       0x00000076
#define Excel_HyperlinksApplication 0x00000094
#define Excel_HyperlinksCreator     0x00000095
#define Excel_HyperlinksParent      0x00000096
/* NICDBLD_END> Object Name: Hyperlinks */

/* NICDBLD_BEGIN> Object Name: Hyperlink */
#define Excel_HyperlinkType          0x0000006C
#define Excel_HyperlinkName          0x0000006E
#define Excel_HyperlinkApplication   0x00000094
#define Excel_HyperlinkCreator       0x00000095
#define Excel_HyperlinkParent        0x00000096
#define Excel_HyperlinkRange         0x000000C5
#define Excel_HyperlinkAddress       0x000000EC
#define Excel_HyperlinkSubAddress    0x000005BF
#define Excel_HyperlinkShape         0x0000062E
#define Excel_HyperlinkScreenTip     0x00000759
#define Excel_HyperlinkTextToDisplay 0x0000075A
#define Excel_HyperlinkEmailSubject  0x0000075B
/* NICDBLD_END> Object Name: Hyperlink */

/* NICDBLD_BEGIN> Object Name: AutoFilter */
#define Excel_AutoFilterApplication 0x00000094
#define Excel_AutoFilterCreator     0x00000095
#define Excel_AutoFilterParent      0x00000096
#define Excel_AutoFilterRange       0x000000C5
#define Excel_AutoFilterFilters     0x00000651
/* NICDBLD_END> Object Name: AutoFilter */

/* NICDBLD_BEGIN> Object Name: Filters */
#define Excel_Filters_NewEnum    0xFFFFFFFC
#define Excel_FiltersCount       0x00000076
#define Excel_FiltersApplication 0x00000094
#define Excel_FiltersCreator     0x00000095
#define Excel_FiltersParent      0x00000096
/* NICDBLD_END> Object Name: Filters */

/* NICDBLD_BEGIN> Object Name: Filter */
#define Excel_FilterApplication 0x00000094
#define Excel_FilterCreator     0x00000095
#define Excel_FilterParent      0x00000096
#define Excel_FilterCriteria1   0x0000031C
#define Excel_FilterOperator    0x0000031D
#define Excel_FilterCriteria2   0x0000031E
#define Excel_FilterOn          0x00000652
/* NICDBLD_END> Object Name: Filter */

/* NICDBLD_BEGIN> Object Name: AutoCorrect */
#define Excel_AutoCorrectApplication           0x00000094
#define Excel_AutoCorrectCreator               0x00000095
#define Excel_AutoCorrectParent                0x00000096
#define Excel_AutoCorrectReplaceText           0x0000047C
#define Excel_AutoCorrectTwoInitialCapitals    0x0000047D
#define Excel_AutoCorrectCapitalizeNamesOfDays 0x0000047E
#define Excel_AutoCorrectCorrectSentenceCap    0x00000653
#define Excel_AutoCorrectCorrectCapsLock       0x00000654
/* NICDBLD_END> Object Name: AutoCorrect */

/* NICDBLD_BEGIN> Object Name: Border */
#define Excel_BorderColorIndex  0x00000061
#define Excel_BorderColor       0x00000063
#define Excel_BorderLineStyle   0x00000077
#define Excel_BorderWeight      0x00000078
#define Excel_BorderApplication 0x00000094
#define Excel_BorderCreator     0x00000095
#define Excel_BorderParent      0x00000096
/* NICDBLD_END> Object Name: Border */

/* NICDBLD_BEGIN> Object Name: Interior */
#define Excel_InteriorPattern           0x0000005F
#define Excel_InteriorColorIndex        0x00000061
#define Excel_InteriorPatternColorIndex 0x00000062
#define Excel_InteriorColor             0x00000063
#define Excel_InteriorPatternColor      0x00000064
#define Excel_InteriorInvertIfNegative  0x00000084
#define Excel_InteriorApplication       0x00000094
#define Excel_InteriorCreator           0x00000095
#define Excel_InteriorParent            0x00000096
/* NICDBLD_END> Object Name: Interior */

/* NICDBLD_BEGIN> Object Name: ChartFillFormat */
#define Excel_ChartFillFormatPattern            0x0000005F
#define Excel_ChartFillFormatType               0x0000006C
#define Excel_ChartFillFormatApplication        0x00000094
#define Excel_ChartFillFormatCreator            0x00000095
#define Excel_ChartFillFormatParent             0x00000096
#define Excel_ChartFillFormatVisible            0x0000022E
#define Excel_ChartFillFormatPresetTexture      0x0000065A
#define Excel_ChartFillFormatPresetGradientType 0x00000665
#define Excel_ChartFillFormatBackColor          0x00000666
#define Excel_ChartFillFormatForeColor          0x00000667
#define Excel_ChartFillFormatGradientColorType  0x00000668
#define Excel_ChartFillFormatGradientDegree     0x00000669
#define Excel_ChartFillFormatGradientStyle      0x0000066A
#define Excel_ChartFillFormatGradientVariant    0x0000066B
#define Excel_ChartFillFormatTextureName        0x0000066C
#define Excel_ChartFillFormatTextureType        0x0000066D
/* NICDBLD_END> Object Name: ChartFillFormat */

/* NICDBLD_BEGIN> Object Name: ChartColorFormat */
#define Excel_ChartColorFormat_Default    0x00000000
#define Excel_ChartColorFormatType        0x0000006C
#define Excel_ChartColorFormatApplication 0x00000094
#define Excel_ChartColorFormatCreator     0x00000095
#define Excel_ChartColorFormatParent      0x00000096
#define Excel_ChartColorFormatRGB         0x0000041F
#define Excel_ChartColorFormatSchemeColor 0x0000066E
/* NICDBLD_END> Object Name: ChartColorFormat */

/* NICDBLD_BEGIN> Object Name: Axis */
#define Excel_AxisHasMajorGridlines     0x00000018
#define Excel_AxisHasMinorGridlines     0x00000019
#define Excel_AxisMajorTickMark         0x0000001A
#define Excel_AxisMinorTickMark         0x0000001B
#define Excel_AxisTickLabelPosition     0x0000001C
#define Excel_AxisTickLabelSpacing      0x0000001D
#define Excel_AxisTickMarkSpacing       0x0000001F
#define Excel_AxisMinimumScale          0x00000021
#define Excel_AxisMinimumScaleIsAuto    0x00000022
#define Excel_AxisMaximumScale          0x00000023
#define Excel_AxisMaximumScaleIsAuto    0x00000024
#define Excel_AxisMajorUnit             0x00000025
#define Excel_AxisMajorUnitIsAuto       0x00000026
#define Excel_AxisMinorUnit             0x00000027
#define Excel_AxisMinorUnitIsAuto       0x00000028
#define Excel_AxisScaleType             0x00000029
#define Excel_AxisCrosses               0x0000002A
#define Excel_AxisCrossesAt             0x0000002B
#define Excel_AxisReversePlotOrder      0x0000002C
#define Excel_AxisAxisBetweenCategories 0x0000002D
#define Excel_AxisAxisGroup             0x0000002F
#define Excel_AxisHasTitle              0x00000036
#define Excel_AxisAxisTitle             0x00000052
#define Excel_AxisMajorGridlines        0x00000059
#define Excel_AxisMinorGridlines        0x0000005A
#define Excel_AxisTickLabels            0x0000005B
#define Excel_AxisType                  0x0000006C
#define Excel_AxisWidth                 0x0000007A
#define Excel_AxisHeight                0x0000007B
#define Excel_AxisTop                   0x0000007E
#define Excel_AxisLeft                  0x0000007F
#define Excel_AxisBorder                0x00000080
#define Excel_AxisApplication           0x00000094
#define Excel_AxisCreator               0x00000095
#define Excel_AxisParent                0x00000096
#define Excel_AxisCategoryNames         0x0000009C
#define Excel_AxisBaseUnit              0x0000066F
#define Excel_AxisBaseUnitIsAuto        0x00000670
#define Excel_AxisMajorUnitScale        0x00000671
#define Excel_AxisMinorUnitScale        0x00000672
#define Excel_AxisCategoryType          0x00000673
#define Excel_AxisDisplayUnit           0x0000075E
#define Excel_AxisDisplayUnitCustom     0x0000075F
#define Excel_AxisHasDisplayUnitLabel   0x00000760
#define Excel_AxisDisplayUnitLabel      0x00000761
/* NICDBLD_END> Object Name: Axis */

/* NICDBLD_BEGIN> Object Name: ChartTitle */
#define Excel_ChartTitleShadow              0x00000067
#define Excel_ChartTitleName                0x0000006E
#define Excel_ChartTitleTop                 0x0000007E
#define Excel_ChartTitleLeft                0x0000007F
#define Excel_ChartTitleBorder              0x00000080
#define Excel_ChartTitleInterior            0x00000081
#define Excel_ChartTitleOrientation         0x00000086
#define Excel_ChartTitleHorizontalAlignment 0x00000088
#define Excel_ChartTitleVerticalAlignment   0x00000089
#define Excel_ChartTitleText                0x0000008A
#define Excel_ChartTitleCaption             0x0000008B
#define Excel_ChartTitleFont                0x00000092
#define Excel_ChartTitleApplication         0x00000094
#define Excel_ChartTitleCreator             0x00000095
#define Excel_ChartTitleParent              0x00000096
#define Excel_ChartTitleReadingOrder        0x000003CF
#define Excel_ChartTitleAutoScaleFont       0x000005F5
#define Excel_ChartTitleFill                0x0000067F
/* NICDBLD_END> Object Name: ChartTitle */

/* NICDBLD_BEGIN> Object Name: AxisTitle */
#define Excel_AxisTitleShadow              0x00000067
#define Excel_AxisTitleName                0x0000006E
#define Excel_AxisTitleTop                 0x0000007E
#define Excel_AxisTitleLeft                0x0000007F
#define Excel_AxisTitleBorder              0x00000080
#define Excel_AxisTitleInterior            0x00000081
#define Excel_AxisTitleOrientation         0x00000086
#define Excel_AxisTitleHorizontalAlignment 0x00000088
#define Excel_AxisTitleVerticalAlignment   0x00000089
#define Excel_AxisTitleText                0x0000008A
#define Excel_AxisTitleCaption             0x0000008B
#define Excel_AxisTitleFont                0x00000092
#define Excel_AxisTitleApplication         0x00000094
#define Excel_AxisTitleCreator             0x00000095
#define Excel_AxisTitleParent              0x00000096
#define Excel_AxisTitleReadingOrder        0x000003CF
#define Excel_AxisTitleAutoScaleFont       0x000005F5
#define Excel_AxisTitleFill                0x0000067F
/* NICDBLD_END> Object Name: AxisTitle */

/* NICDBLD_BEGIN> Object Name: ChartGroup */
#define Excel_ChartGroupAxisGroup           0x0000002F
#define Excel_ChartGroupGapWidth            0x00000033
#define Excel_ChartGroupOverlap             0x00000038
#define Excel_ChartGroupVaryByCategories    0x0000003C
#define Excel_ChartGroupHasDropLines        0x0000003D
#define Excel_ChartGroupHasHiLoLines        0x0000003E
#define Excel_ChartGroupFirstSliceAngle     0x0000003F
#define Excel_ChartGroupHasRadarAxisLabels  0x00000040
#define Excel_ChartGroupHasSeriesLines      0x00000041
#define Excel_ChartGroupHasUpDownBars       0x00000042
#define Excel_ChartGroupType                0x0000006C
#define Excel_ChartGroupSubType             0x0000006D
#define Excel_ChartGroupUpBars              0x0000008C
#define Excel_ChartGroupDownBars            0x0000008D
#define Excel_ChartGroupDropLines           0x0000008E
#define Excel_ChartGroupHiLoLines           0x0000008F
#define Excel_ChartGroupRadarAxisLabels     0x00000090
#define Excel_ChartGroupSeriesLines         0x00000091
#define Excel_ChartGroupApplication         0x00000094
#define Excel_ChartGroupCreator             0x00000095
#define Excel_ChartGroupParent              0x00000096
#define Excel_ChartGroupIndex               0x000001E6
#define Excel_ChartGroupDoughnutHoleSize    0x00000466
#define Excel_ChartGroupSizeRepresents      0x00000674
#define Excel_ChartGroupBubbleScale         0x00000675
#define Excel_ChartGroupShowNegativeBubbles 0x00000676
#define Excel_ChartGroupSplitType           0x00000677
#define Excel_ChartGroupSplitValue          0x00000678
#define Excel_ChartGroupSecondPlotSize      0x00000679
#define Excel_ChartGroupHas3DShading        0x0000067A
/* NICDBLD_END> Object Name: ChartGroup */

/* NICDBLD_BEGIN> Object Name: ChartGroups */
#define Excel_ChartGroupsCount       0x00000076
#define Excel_ChartGroupsApplication 0x00000094
#define Excel_ChartGroupsCreator     0x00000095
#define Excel_ChartGroupsParent      0x00000096
/* NICDBLD_END> Object Name: ChartGroups */

/* NICDBLD_BEGIN> Object Name: Axes */
#define Excel_AxesCount       0x00000076
#define Excel_AxesApplication 0x00000094
#define Excel_AxesCreator     0x00000095
#define Excel_AxesParent      0x00000096
/* NICDBLD_END> Object Name: Axes */

/* NICDBLD_BEGIN> Object Name: Points */
#define Excel_PointsCount       0x00000076
#define Excel_PointsApplication 0x00000094
#define Excel_PointsCreator     0x00000095
#define Excel_PointsParent      0x00000096
/* NICDBLD_END> Object Name: Points */

/* NICDBLD_BEGIN> Object Name: Point */
#define Excel_PointMarkerStyle                0x00000048
#define Excel_PointMarkerBackgroundColor      0x00000049
#define Excel_PointMarkerBackgroundColorIndex 0x0000004A
#define Excel_PointMarkerForegroundColor      0x0000004B
#define Excel_PointMarkerForegroundColorIndex 0x0000004C
#define Excel_PointHasDataLabel               0x0000004D
#define Excel_PointShadow                     0x00000067
#define Excel_PointBorder                     0x00000080
#define Excel_PointInterior                   0x00000081
#define Excel_PointInvertIfNegative           0x00000084
#define Excel_PointApplication                0x00000094
#define Excel_PointCreator                    0x00000095
#define Excel_PointParent                     0x00000096
#define Excel_PointDataLabel                  0x0000009E
#define Excel_PointPictureType                0x000000A1
#define Excel_PointPictureUnit                0x000000A2
#define Excel_PointExplosion                  0x000000B6
#define Excel_PointMarkerSize                 0x000000E7
#define Excel_PointApplyPictToSides           0x0000067B
#define Excel_PointApplyPictToFront           0x0000067C
#define Excel_PointApplyPictToEnd             0x0000067D
#define Excel_PointSecondaryPlot              0x0000067E
#define Excel_PointFill                       0x0000067F
/* NICDBLD_END> Object Name: Point */

/* NICDBLD_BEGIN> Object Name: Series */
#define Excel_SeriesAxisGroup                  0x0000002F
#define Excel_SeriesMarkerStyle                0x00000048
#define Excel_SeriesMarkerBackgroundColor      0x00000049
#define Excel_SeriesMarkerBackgroundColorIndex 0x0000004A
#define Excel_SeriesMarkerForegroundColor      0x0000004B
#define Excel_SeriesMarkerForegroundColorIndex 0x0000004C
#define Excel_SeriesHasDataLabels              0x0000004E
#define Excel_SeriesShadow                     0x00000067
#define Excel_SeriesType                       0x0000006C
#define Excel_SeriesName                       0x0000006E
#define Excel_SeriesBorder                     0x00000080
#define Excel_SeriesInterior                   0x00000081
#define Excel_SeriesInvertIfNegative           0x00000084
#define Excel_SeriesApplication                0x00000094
#define Excel_SeriesCreator                    0x00000095
#define Excel_SeriesParent                     0x00000096
#define Excel_SeriesErrorBars                  0x0000009F
#define Excel_SeriesHasErrorBars               0x000000A0
#define Excel_SeriesPictureType                0x000000A1
#define Excel_SeriesPictureUnit                0x000000A2
#define Excel_SeriesSmooth                     0x000000A3
#define Excel_SeriesValues                     0x000000A4
#define Excel_SeriesExplosion                  0x000000B6
#define Excel_SeriesPlotOrder                  0x000000E4
#define Excel_SeriesMarkerSize                 0x000000E7
#define Excel_SeriesFormula                    0x00000105
#define Excel_SeriesFormulaLocal               0x00000107
#define Excel_SeriesFormulaR1C1                0x00000108
#define Excel_SeriesFormulaR1C1Local           0x00000109
#define Excel_SeriesXValues                    0x00000457
#define Excel_SeriesHasLeaderLines             0x00000572
#define Excel_SeriesChartType                  0x00000578
#define Excel_SeriesBarShape                   0x0000057B
#define Excel_SeriesApplyPictToSides           0x0000067B
#define Excel_SeriesApplyPictToFront           0x0000067C
#define Excel_SeriesApplyPictToEnd             0x0000067D
#define Excel_SeriesFill                       0x0000067F
#define Excel_SeriesBubbleSizes                0x00000680
#define Excel_SeriesHas3DEffect                0x00000681
#define Excel_SeriesLeaderLines                0x00000682
/* NICDBLD_END> Object Name: Series */

/* NICDBLD_BEGIN> Object Name: SeriesCollection */
#define Excel_SeriesCollectionCount       0x00000076
#define Excel_SeriesCollectionApplication 0x00000094
#define Excel_SeriesCollectionCreator     0x00000095
#define Excel_SeriesCollectionParent      0x00000096
/* NICDBLD_END> Object Name: SeriesCollection */

/* NICDBLD_BEGIN> Object Name: DataLabel */
#define Excel_DataLabelShadow              0x00000067
#define Excel_DataLabelType                0x0000006C
#define Excel_DataLabelName                0x0000006E
#define Excel_DataLabelTop                 0x0000007E
#define Excel_DataLabelLeft                0x0000007F
#define Excel_DataLabelBorder              0x00000080
#define Excel_DataLabelInterior            0x00000081
#define Excel_DataLabelPosition            0x00000085
#define Excel_DataLabelOrientation         0x00000086
#define Excel_DataLabelAutoText            0x00000087
#define Excel_DataLabelHorizontalAlignment 0x00000088
#define Excel_DataLabelVerticalAlignment   0x00000089
#define Excel_DataLabelText                0x0000008A
#define Excel_DataLabelCaption             0x0000008B
#define Excel_DataLabelFont                0x00000092
#define Excel_DataLabelApplication         0x00000094
#define Excel_DataLabelCreator             0x00000095
#define Excel_DataLabelParent              0x00000096
#define Excel_DataLabelShowLegendKey       0x000000AB
#define Excel_DataLabelNumberFormat        0x000000C1
#define Excel_DataLabelNumberFormatLinked  0x000000C2
#define Excel_DataLabelReadingOrder        0x000003CF
#define Excel_DataLabelNumberFormatLocal   0x00000449
#define Excel_DataLabelAutoScaleFont       0x000005F5
#define Excel_DataLabelFill                0x0000067F
/* NICDBLD_END> Object Name: DataLabel */

/* NICDBLD_BEGIN> Object Name: DataLabels */
#define Excel_DataLabelsShadow              0x00000067
#define Excel_DataLabelsType                0x0000006C
#define Excel_DataLabelsName                0x0000006E
#define Excel_DataLabelsCount               0x00000076
#define Excel_DataLabelsBorder              0x00000080
#define Excel_DataLabelsInterior            0x00000081
#define Excel_DataLabelsPosition            0x00000085
#define Excel_DataLabelsOrientation         0x00000086
#define Excel_DataLabelsAutoText            0x00000087
#define Excel_DataLabelsHorizontalAlignment 0x00000088
#define Excel_DataLabelsVerticalAlignment   0x00000089
#define Excel_DataLabelsFont                0x00000092
#define Excel_DataLabelsApplication         0x00000094
#define Excel_DataLabelsCreator             0x00000095
#define Excel_DataLabelsParent              0x00000096
#define Excel_DataLabelsShowLegendKey       0x000000AB
#define Excel_DataLabelsNumberFormat        0x000000C1
#define Excel_DataLabelsNumberFormatLinked  0x000000C2
#define Excel_DataLabelsReadingOrder        0x000003CF
#define Excel_DataLabelsNumberFormatLocal   0x00000449
#define Excel_DataLabelsAutoScaleFont       0x000005F5
#define Excel_DataLabelsFill                0x0000067F
/* NICDBLD_END> Object Name: DataLabels */

/* NICDBLD_BEGIN> Object Name: LegendEntry */
#define Excel_LegendEntryWidth         0x0000007A
#define Excel_LegendEntryHeight        0x0000007B
#define Excel_LegendEntryTop           0x0000007E
#define Excel_LegendEntryLeft          0x0000007F
#define Excel_LegendEntryFont          0x00000092
#define Excel_LegendEntryApplication   0x00000094
#define Excel_LegendEntryCreator       0x00000095
#define Excel_LegendEntryParent        0x00000096
#define Excel_LegendEntryLegendKey     0x000000AE
#define Excel_LegendEntryIndex         0x000001E6
#define Excel_LegendEntryAutoScaleFont 0x000005F5
/* NICDBLD_END> Object Name: LegendEntry */

/* NICDBLD_BEGIN> Object Name: LegendEntries */
#define Excel_LegendEntriesCount       0x00000076
#define Excel_LegendEntriesApplication 0x00000094
#define Excel_LegendEntriesCreator     0x00000095
#define Excel_LegendEntriesParent      0x00000096
/* NICDBLD_END> Object Name: LegendEntries */

/* NICDBLD_BEGIN> Object Name: LegendKey */
#define Excel_LegendKeyMarkerStyle                0x00000048
#define Excel_LegendKeyMarkerBackgroundColor      0x00000049
#define Excel_LegendKeyMarkerBackgroundColorIndex 0x0000004A
#define Excel_LegendKeyMarkerForegroundColor      0x0000004B
#define Excel_LegendKeyMarkerForegroundColorIndex 0x0000004C
#define Excel_LegendKeyShadow                     0x00000067
#define Excel_LegendKeyWidth                      0x0000007A
#define Excel_LegendKeyHeight                     0x0000007B
#define Excel_LegendKeyTop                        0x0000007E
#define Excel_LegendKeyLeft                       0x0000007F
#define Excel_LegendKeyBorder                     0x00000080
#define Excel_LegendKeyInterior                   0x00000081
#define Excel_LegendKeyInvertIfNegative           0x00000084
#define Excel_LegendKeyApplication                0x00000094
#define Excel_LegendKeyCreator                    0x00000095
#define Excel_LegendKeyParent                     0x00000096
#define Excel_LegendKeyPictureType                0x000000A1
#define Excel_LegendKeyPictureUnit                0x000000A2
#define Excel_LegendKeySmooth                     0x000000A3
#define Excel_LegendKeyMarkerSize                 0x000000E7
#define Excel_LegendKeyFill                       0x0000067F
/* NICDBLD_END> Object Name: LegendKey */

/* NICDBLD_BEGIN> Object Name: Trendlines */
#define Excel_TrendlinesCount       0x00000076
#define Excel_TrendlinesApplication 0x00000094
#define Excel_TrendlinesCreator     0x00000095
#define Excel_TrendlinesParent      0x00000096
/* NICDBLD_END> Object Name: Trendlines */

/* NICDBLD_BEGIN> Object Name: Trendline */
#define Excel_TrendlineType            0x0000006C
#define Excel_TrendlineName            0x0000006E
#define Excel_TrendlineBorder          0x00000080
#define Excel_TrendlineApplication     0x00000094
#define Excel_TrendlineCreator         0x00000095
#define Excel_TrendlineParent          0x00000096
#define Excel_TrendlineDataLabel       0x0000009E
#define Excel_TrendlinePeriod          0x000000B8
#define Excel_TrendlineBackward        0x000000B9
#define Excel_TrendlineIntercept       0x000000BA
#define Excel_TrendlineInterceptIsAuto 0x000000BB
#define Excel_TrendlineNameIsAuto      0x000000BC
#define Excel_TrendlineDisplayRSquared 0x000000BD
#define Excel_TrendlineDisplayEquation 0x000000BE
#define Excel_TrendlineForward         0x000000BF
#define Excel_TrendlineOrder           0x000000C0
#define Excel_TrendlineIndex           0x000001E6
/* NICDBLD_END> Object Name: Trendline */

/* NICDBLD_BEGIN> Object Name: Corners */
#define Excel_CornersName        0x0000006E
#define Excel_CornersApplication 0x00000094
#define Excel_CornersCreator     0x00000095
#define Excel_CornersParent      0x00000096
/* NICDBLD_END> Object Name: Corners */

/* NICDBLD_BEGIN> Object Name: SeriesLines */
#define Excel_SeriesLinesName        0x0000006E
#define Excel_SeriesLinesBorder      0x00000080
#define Excel_SeriesLinesApplication 0x00000094
#define Excel_SeriesLinesCreator     0x00000095
#define Excel_SeriesLinesParent      0x00000096
/* NICDBLD_END> Object Name: SeriesLines */

/* NICDBLD_BEGIN> Object Name: HiLoLines */
#define Excel_HiLoLinesName        0x0000006E
#define Excel_HiLoLinesBorder      0x00000080
#define Excel_HiLoLinesApplication 0x00000094
#define Excel_HiLoLinesCreator     0x00000095
#define Excel_HiLoLinesParent      0x00000096
/* NICDBLD_END> Object Name: HiLoLines */

/* NICDBLD_BEGIN> Object Name: Gridlines */
#define Excel_GridlinesName        0x0000006E
#define Excel_GridlinesBorder      0x00000080
#define Excel_GridlinesApplication 0x00000094
#define Excel_GridlinesCreator     0x00000095
#define Excel_GridlinesParent      0x00000096
/* NICDBLD_END> Object Name: Gridlines */

/* NICDBLD_BEGIN> Object Name: DropLines */
#define Excel_DropLinesName        0x0000006E
#define Excel_DropLinesBorder      0x00000080
#define Excel_DropLinesApplication 0x00000094
#define Excel_DropLinesCreator     0x00000095
#define Excel_DropLinesParent      0x00000096
/* NICDBLD_END> Object Name: DropLines */

/* NICDBLD_BEGIN> Object Name: LeaderLines */
#define Excel_LeaderLinesBorder      0x00000080
#define Excel_LeaderLinesApplication 0x00000094
#define Excel_LeaderLinesCreator     0x00000095
#define Excel_LeaderLinesParent      0x00000096
/* NICDBLD_END> Object Name: LeaderLines */

/* NICDBLD_BEGIN> Object Name: UpBars */
#define Excel_UpBarsName        0x0000006E
#define Excel_UpBarsBorder      0x00000080
#define Excel_UpBarsInterior    0x00000081
#define Excel_UpBarsApplication 0x00000094
#define Excel_UpBarsCreator     0x00000095
#define Excel_UpBarsParent      0x00000096
#define Excel_UpBarsFill        0x0000067F
/* NICDBLD_END> Object Name: UpBars */

/* NICDBLD_BEGIN> Object Name: DownBars */
#define Excel_DownBarsName        0x0000006E
#define Excel_DownBarsBorder      0x00000080
#define Excel_DownBarsInterior    0x00000081
#define Excel_DownBarsApplication 0x00000094
#define Excel_DownBarsCreator     0x00000095
#define Excel_DownBarsParent      0x00000096
#define Excel_DownBarsFill        0x0000067F
/* NICDBLD_END> Object Name: DownBars */

/* NICDBLD_BEGIN> Object Name: Floor */
#define Excel_FloorName        0x0000006E
#define Excel_FloorBorder      0x00000080
#define Excel_FloorInterior    0x00000081
#define Excel_FloorApplication 0x00000094
#define Excel_FloorCreator     0x00000095
#define Excel_FloorParent      0x00000096
#define Excel_FloorPictureType 0x000000A1
#define Excel_FloorFill        0x0000067F
/* NICDBLD_END> Object Name: Floor */

/* NICDBLD_BEGIN> Object Name: Walls */
#define Excel_WallsName        0x0000006E
#define Excel_WallsBorder      0x00000080
#define Excel_WallsInterior    0x00000081
#define Excel_WallsApplication 0x00000094
#define Excel_WallsCreator     0x00000095
#define Excel_WallsParent      0x00000096
#define Excel_WallsPictureType 0x000000A1
#define Excel_WallsPictureUnit 0x000000A2
#define Excel_WallsFill        0x0000067F
/* NICDBLD_END> Object Name: Walls */

/* NICDBLD_BEGIN> Object Name: TickLabels */
#define Excel_TickLabelsName               0x0000006E
#define Excel_TickLabelsOrientation        0x00000086
#define Excel_TickLabelsFont               0x00000092
#define Excel_TickLabelsApplication        0x00000094
#define Excel_TickLabelsCreator            0x00000095
#define Excel_TickLabelsParent             0x00000096
#define Excel_TickLabelsNumberFormat       0x000000C1
#define Excel_TickLabelsNumberFormatLinked 0x000000C2
#define Excel_TickLabelsOffset             0x000000FE
#define Excel_TickLabelsAlignment          0x000001C5
#define Excel_TickLabelsReadingOrder       0x000003CF
#define Excel_TickLabelsNumberFormatLocal  0x00000449
#define Excel_TickLabelsAutoScaleFont      0x000005F5
#define Excel_TickLabelsDepth              0x00000762
/* NICDBLD_END> Object Name: TickLabels */

/* NICDBLD_BEGIN> Object Name: PlotArea */
#define Excel_PlotAreaName         0x0000006E
#define Excel_PlotAreaWidth        0x0000007A
#define Excel_PlotAreaHeight       0x0000007B
#define Excel_PlotAreaTop          0x0000007E
#define Excel_PlotAreaLeft         0x0000007F
#define Excel_PlotAreaBorder       0x00000080
#define Excel_PlotAreaInterior     0x00000081
#define Excel_PlotAreaApplication  0x00000094
#define Excel_PlotAreaCreator      0x00000095
#define Excel_PlotAreaParent       0x00000096
#define Excel_PlotAreaFill         0x0000067F
#define Excel_PlotAreaInsideLeft   0x00000683
#define Excel_PlotAreaInsideTop    0x00000684
#define Excel_PlotAreaInsideWidth  0x00000685
#define Excel_PlotAreaInsideHeight 0x00000686
/* NICDBLD_END> Object Name: PlotArea */

/* NICDBLD_BEGIN> Object Name: ChartArea */
#define Excel_ChartAreaShadow        0x00000067
#define Excel_ChartAreaName          0x0000006E
#define Excel_ChartAreaWidth         0x0000007A
#define Excel_ChartAreaHeight        0x0000007B
#define Excel_ChartAreaTop           0x0000007E
#define Excel_ChartAreaLeft          0x0000007F
#define Excel_ChartAreaBorder        0x00000080
#define Excel_ChartAreaInterior      0x00000081
#define Excel_ChartAreaFont          0x00000092
#define Excel_ChartAreaApplication   0x00000094
#define Excel_ChartAreaCreator       0x00000095
#define Excel_ChartAreaParent        0x00000096
#define Excel_ChartAreaAutoScaleFont 0x000005F5
#define Excel_ChartAreaFill          0x0000067F
/* NICDBLD_END> Object Name: ChartArea */

/* NICDBLD_BEGIN> Object Name: Legend */
#define Excel_LegendShadow        0x00000067
#define Excel_LegendName          0x0000006E
#define Excel_LegendWidth         0x0000007A
#define Excel_LegendHeight        0x0000007B
#define Excel_LegendTop           0x0000007E
#define Excel_LegendLeft          0x0000007F
#define Excel_LegendBorder        0x00000080
#define Excel_LegendInterior      0x00000081
#define Excel_LegendPosition      0x00000085
#define Excel_LegendFont          0x00000092
#define Excel_LegendApplication   0x00000094
#define Excel_LegendCreator       0x00000095
#define Excel_LegendParent        0x00000096
#define Excel_LegendAutoScaleFont 0x000005F5
#define Excel_LegendFill          0x0000067F
/* NICDBLD_END> Object Name: Legend */

/* NICDBLD_BEGIN> Object Name: ErrorBars */
#define Excel_ErrorBarsName        0x0000006E
#define Excel_ErrorBarsBorder      0x00000080
#define Excel_ErrorBarsApplication 0x00000094
#define Excel_ErrorBarsCreator     0x00000095
#define Excel_ErrorBarsParent      0x00000096
#define Excel_ErrorBarsEndStyle    0x00000464
/* NICDBLD_END> Object Name: ErrorBars */

/* NICDBLD_BEGIN> Object Name: DataTable */
#define Excel_DataTableBorder              0x00000080
#define Excel_DataTableFont                0x00000092
#define Excel_DataTableApplication         0x00000094
#define Excel_DataTableCreator             0x00000095
#define Excel_DataTableParent              0x00000096
#define Excel_DataTableShowLegendKey       0x000000AB
#define Excel_DataTableAutoScaleFont       0x000005F5
#define Excel_DataTableHasBorderHorizontal 0x00000687
#define Excel_DataTableHasBorderVertical   0x00000688
#define Excel_DataTableHasBorderOutline    0x00000689
/* NICDBLD_END> Object Name: DataTable */

/* NICDBLD_BEGIN> Object Name: Phonetic */
#define Excel_PhoneticText          0x0000008A
#define Excel_PhoneticFont          0x00000092
#define Excel_PhoneticApplication   0x00000094
#define Excel_PhoneticCreator       0x00000095
#define Excel_PhoneticParent        0x00000096
#define Excel_PhoneticAlignment     0x000001C5
#define Excel_PhoneticVisible       0x0000022E
#define Excel_PhoneticCharacterType 0x0000068A
/* NICDBLD_END> Object Name: Phonetic */

/* NICDBLD_BEGIN> Object Name: Shape */
#define Excel_ShapeRotation            0x0000003B
#define Excel_ShapeShadow              0x00000067
#define Excel_ShapeType                0x0000006C
#define Excel_ShapeName                0x0000006E
#define Excel_ShapeWidth               0x0000007A
#define Excel_ShapeHeight              0x0000007B
#define Excel_ShapeTop                 0x0000007E
#define Excel_ShapeLeft                0x0000007F
#define Excel_ShapeApplication         0x00000094
#define Excel_ShapeCreator             0x00000095
#define Excel_ShapeParent              0x00000096
#define Excel_ShapeLocked              0x0000010D
#define Excel_ShapeVisible             0x0000022E
#define Excel_ShapeOnAction            0x00000254
#define Excel_ShapeBottomRightCell     0x00000267
#define Excel_ShapePlacement           0x00000269
#define Excel_ShapeTopLeftCell         0x0000026C
#define Excel_ShapeVertices            0x0000026D
#define Excel_ShapeLine                0x00000331
#define Excel_ShapePictureFormat       0x0000065F
#define Excel_ShapeFill                0x0000067F
#define Excel_ShapeAdjustments         0x0000069B
#define Excel_ShapeTextFrame           0x0000069C
#define Excel_ShapeAutoShapeType       0x0000069D
#define Excel_ShapeCallout             0x0000069E
#define Excel_ShapeConnectionSiteCount 0x0000069F
#define Excel_ShapeConnector           0x000006A0
#define Excel_ShapeConnectorFormat     0x000006A1
#define Excel_ShapeGroupItems          0x000006A2
#define Excel_ShapeHorizontalFlip      0x000006A3
#define Excel_ShapeLockAspectRatio     0x000006A4
#define Excel_ShapeNodes               0x000006A5
#define Excel_ShapeTextEffect          0x000006A6
#define Excel_ShapeThreeD              0x000006A7
#define Excel_ShapeVerticalFlip        0x000006A8
#define Excel_ShapeZOrderPosition      0x000006A9
#define Excel_ShapeHyperlink           0x000006AA
#define Excel_ShapeBlackWhiteMode      0x000006AB
#define Excel_ShapeDrawingObject       0x000006AC
#define Excel_ShapeControlFormat       0x000006AD
#define Excel_ShapeLinkFormat          0x000006AE
#define Excel_ShapeOLEFormat           0x000006AF
#define Excel_ShapeFormControlType     0x000006B0
#define Excel_ShapeAlternativeText     0x00000763
#define Excel_ShapeScript              0x00000764
/* NICDBLD_END> Object Name: Shape */

/* NICDBLD_BEGIN> Object Name: Shapes */
#define Excel_Shapes_NewEnum    0xFFFFFFFC
#define Excel_ShapesCount       0x00000076
#define Excel_ShapesApplication 0x00000094
#define Excel_ShapesCreator     0x00000095
#define Excel_ShapesParent      0x00000096
/* NICDBLD_END> Object Name: Shapes */

/* NICDBLD_BEGIN> Object Name: ShapeRange */
#define Excel_ShapeRange_NewEnum            0xFFFFFFFC
#define Excel_ShapeRangeRotation            0x0000003B
#define Excel_ShapeRangeShadow              0x00000067
#define Excel_ShapeRangeType                0x0000006C
#define Excel_ShapeRangeName                0x0000006E
#define Excel_ShapeRangeCount               0x00000076
#define Excel_ShapeRangeWidth               0x0000007A
#define Excel_ShapeRangeHeight              0x0000007B
#define Excel_ShapeRangeTop                 0x0000007E
#define Excel_ShapeRangeLeft                0x0000007F
#define Excel_ShapeRangeApplication         0x00000094
#define Excel_ShapeRangeCreator             0x00000095
#define Excel_ShapeRangeParent              0x00000096
#define Excel_ShapeRangeVisible             0x0000022E
#define Excel_ShapeRangeVertices            0x0000026D
#define Excel_ShapeRangeLine                0x00000331
#define Excel_ShapeRangePictureFormat       0x0000065F
#define Excel_ShapeRangeFill                0x0000067F
#define Excel_ShapeRangeAdjustments         0x0000069B
#define Excel_ShapeRangeTextFrame           0x0000069C
#define Excel_ShapeRangeAutoShapeType       0x0000069D
#define Excel_ShapeRangeCallout             0x0000069E
#define Excel_ShapeRangeConnectionSiteCount 0x0000069F
#define Excel_ShapeRangeConnector           0x000006A0
#define Excel_ShapeRangeConnectorFormat     0x000006A1
#define Excel_ShapeRangeGroupItems          0x000006A2
#define Excel_ShapeRangeHorizontalFlip      0x000006A3
#define Excel_ShapeRangeLockAspectRatio     0x000006A4
#define Excel_ShapeRangeNodes               0x000006A5
#define Excel_ShapeRangeTextEffect          0x000006A6
#define Excel_ShapeRangeThreeD              0x000006A7
#define Excel_ShapeRangeVerticalFlip        0x000006A8
#define Excel_ShapeRangeZOrderPosition      0x000006A9
#define Excel_ShapeRangeBlackWhiteMode      0x000006AB
#define Excel_ShapeRangeAlternativeText     0x00000763
/* NICDBLD_END> Object Name: ShapeRange */

/* NICDBLD_BEGIN> Object Name: GroupShapes */
#define Excel_GroupShapes_NewEnum    0xFFFFFFFC
#define Excel_GroupShapesCount       0x00000076
#define Excel_GroupShapesApplication 0x00000094
#define Excel_GroupShapesCreator     0x00000095
#define Excel_GroupShapesParent      0x00000096
/* NICDBLD_END> Object Name: GroupShapes */

/* NICDBLD_BEGIN> Object Name: TextFrame */
#define Excel_TextFrameOrientation         0x00000086
#define Excel_TextFrameHorizontalAlignment 0x00000088
#define Excel_TextFrameVerticalAlignment   0x00000089
#define Excel_TextFrameApplication         0x00000094
#define Excel_TextFrameCreator             0x00000095
#define Excel_TextFrameParent              0x00000096
#define Excel_TextFrameAutoSize            0x00000266
#define Excel_TextFrameReadingOrder        0x000003CF
#define Excel_TextFrameMarginBottom        0x000006D1
#define Excel_TextFrameMarginLeft          0x000006D2
#define Excel_TextFrameMarginRight         0x000006D3
#define Excel_TextFrameMarginTop           0x000006D4
#define Excel_TextFrameAutoMargins         0x000006D5
/* NICDBLD_END> Object Name: TextFrame */

/* NICDBLD_BEGIN> Object Name: ConnectorFormat */
#define Excel_ConnectorFormatType                0x0000006C
#define Excel_ConnectorFormatApplication         0x00000094
#define Excel_ConnectorFormatCreator             0x00000095
#define Excel_ConnectorFormatParent              0x00000096
#define Excel_ConnectorFormatBeginConnected      0x000006DC
#define Excel_ConnectorFormatBeginConnectedShape 0x000006DD
#define Excel_ConnectorFormatBeginConnectionSite 0x000006DE
#define Excel_ConnectorFormatEndConnected        0x000006DF
#define Excel_ConnectorFormatEndConnectedShape   0x000006E0
#define Excel_ConnectorFormatEndConnectionSite   0x000006E1
/* NICDBLD_END> Object Name: ConnectorFormat */

/* NICDBLD_BEGIN> Object Name: FreeformBuilder */
#define Excel_FreeformBuilderApplication 0x00000094
#define Excel_FreeformBuilderCreator     0x00000095
#define Excel_FreeformBuilderParent      0x00000096
/* NICDBLD_END> Object Name: FreeformBuilder */

/* NICDBLD_BEGIN> Object Name: ControlFormat */
#define Excel_ControlFormat_Default      0x00000000
#define Excel_ControlFormatValue         0x00000006
#define Excel_ControlFormatMultiSelect   0x00000020
#define Excel_ControlFormatApplication   0x00000094
#define Excel_ControlFormatCreator       0x00000095
#define Excel_ControlFormatParent        0x00000096
#define Excel_ControlFormatEnabled       0x00000258
#define Excel_ControlFormatLockedText    0x00000268
#define Excel_ControlFormatPrintObject   0x0000026A
#define Excel_ControlFormatMax           0x0000034A
#define Excel_ControlFormatMin           0x0000034B
#define Excel_ControlFormatSmallChange   0x0000034C
#define Excel_ControlFormatLargeChange   0x0000034D
#define Excel_ControlFormatListFillRange 0x0000034F
#define Excel_ControlFormatDropDownLines 0x00000350
#define Excel_ControlFormatListCount     0x00000351
#define Excel_ControlFormatListIndex     0x00000352
#define Excel_ControlFormatLinkedCell    0x00000422
/* NICDBLD_END> Object Name: ControlFormat */

/* NICDBLD_BEGIN> Object Name: OLEFormat */
#define Excel_OLEFormatApplication 0x00000094
#define Excel_OLEFormatCreator     0x00000095
#define Excel_OLEFormatParent      0x00000096
#define Excel_OLEFormatObject      0x00000419
#define Excel_OLEFormatProgId      0x000005F3
/* NICDBLD_END> Object Name: OLEFormat */

/* NICDBLD_BEGIN> Object Name: LinkFormat */
#define Excel_LinkFormatApplication 0x00000094
#define Excel_LinkFormatCreator     0x00000095
#define Excel_LinkFormatParent      0x00000096
#define Excel_LinkFormatLocked      0x0000010D
#define Excel_LinkFormatAutoUpdate  0x00000418
/* NICDBLD_END> Object Name: LinkFormat */

/* NICDBLD_BEGIN> Object Name: PublishObjects */
#define Excel_PublishObjects_NewEnum    0xFFFFFFFC
#define Excel_PublishObjectsCount       0x00000076
#define Excel_PublishObjectsApplication 0x00000094
#define Excel_PublishObjectsCreator     0x00000095
#define Excel_PublishObjectsParent      0x00000096
/* NICDBLD_END> Object Name: PublishObjects */

/* NICDBLD_BEGIN> Object Name: OLEDBError */
#define Excel_OLEDBErrorApplication 0x00000094
#define Excel_OLEDBErrorCreator     0x00000095
#define Excel_OLEDBErrorParent      0x00000096
#define Excel_OLEDBErrorNumber      0x000001C3
#define Excel_OLEDBErrorErrorString 0x000005D2
#define Excel_OLEDBErrorSqlState    0x00000643
#define Excel_OLEDBErrorNative      0x00000769
#define Excel_OLEDBErrorStage       0x0000076A
/* NICDBLD_END> Object Name: OLEDBError */

/* NICDBLD_BEGIN> Object Name: OLEDBErrors */
#define Excel_OLEDBErrorsCount       0x00000076
#define Excel_OLEDBErrorsApplication 0x00000094
#define Excel_OLEDBErrorsCreator     0x00000095
#define Excel_OLEDBErrorsParent      0x00000096
/* NICDBLD_END> Object Name: OLEDBErrors */

/* NICDBLD_BEGIN> Object Name: Phonetics */
#define Excel_Phonetics_NewEnum      0xFFFFFFFC
#define Excel_PhoneticsCount         0x00000076
#define Excel_PhoneticsText          0x0000008A
#define Excel_PhoneticsFont          0x00000092
#define Excel_PhoneticsApplication   0x00000094
#define Excel_PhoneticsCreator       0x00000095
#define Excel_PhoneticsParent        0x00000096
#define Excel_PhoneticsAlignment     0x000001C5
#define Excel_PhoneticsVisible       0x0000022E
#define Excel_PhoneticsStart         0x00000260
#define Excel_PhoneticsLength        0x00000261
#define Excel_PhoneticsCharacterType 0x0000068A
/* NICDBLD_END> Object Name: Phonetics */

/* NICDBLD_BEGIN> Object Name: PivotLayout */
#define Excel_PivotLayoutApplication 0x00000094
#define Excel_PivotLayoutCreator     0x00000095
#define Excel_PivotLayoutParent      0x00000096
#define Excel_PivotLayoutInnerDetail 0x000002BA
#define Excel_PivotLayoutPivotTable  0x000002CC
#define Excel_PivotLayoutPivotCache  0x000005D8
#define Excel_PivotLayoutCubeFields  0x0000072F
/* NICDBLD_END> Object Name: PivotLayout */

/* NICDBLD_BEGIN> Object Name: DisplayUnitLabel */
#define Excel_DisplayUnitLabelShadow              0x00000067
#define Excel_DisplayUnitLabelName                0x0000006E
#define Excel_DisplayUnitLabelTop                 0x0000007E
#define Excel_DisplayUnitLabelLeft                0x0000007F
#define Excel_DisplayUnitLabelBorder              0x00000080
#define Excel_DisplayUnitLabelInterior            0x00000081
#define Excel_DisplayUnitLabelOrientation         0x00000086
#define Excel_DisplayUnitLabelHorizontalAlignment 0x00000088
#define Excel_DisplayUnitLabelVerticalAlignment   0x00000089
#define Excel_DisplayUnitLabelText                0x0000008A
#define Excel_DisplayUnitLabelCaption             0x0000008B
#define Excel_DisplayUnitLabelFont                0x00000092
#define Excel_DisplayUnitLabelApplication         0x00000094
#define Excel_DisplayUnitLabelCreator             0x00000095
#define Excel_DisplayUnitLabelParent              0x00000096
#define Excel_DisplayUnitLabelReadingOrder        0x000003CF
#define Excel_DisplayUnitLabelAutoScaleFont       0x000005F5
#define Excel_DisplayUnitLabelFill                0x0000067F
/* NICDBLD_END> Object Name: DisplayUnitLabel */
/* NICDBLD_END> ActiveX Automation Properties */

/* NICDBLD_BEGIN> Type Library Specific Types */
enum ExcelEnum_XlCreator
{
	ExcelConst_xlCreatorCode = 1480803660,
	_Excel_XlCreatorForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlCellInsertionMode
{
	ExcelConst_xlOverwriteCells = 0,
	ExcelConst_xlInsertDeleteCells = 1,
	ExcelConst_xlInsertEntireRows = 2,
	_Excel_XlCellInsertionModeForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlPlatform
{
	ExcelConst_xlMacintosh = 1,
	ExcelConst_xlMSDOS = 3,
	ExcelConst_xlWindows = 2,
	_Excel_XlPlatformForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlTextParsingType
{
	ExcelConst_xlDelimited = 1,
	ExcelConst_xlFixedWidth = 2,
	_Excel_XlTextParsingTypeForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlTextQualifier
{
	ExcelConst_xlTextQualifierDoubleQuote = 1,
	ExcelConst_xlTextQualifierNone = -4142,
	ExcelConst_xlTextQualifierSingleQuote = 2,
	_Excel_XlTextQualifierForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlCmdType
{
	ExcelConst_xlCmdCube = 1,
	ExcelConst_xlCmdSql = 2,
	ExcelConst_xlCmdTable = 3,
	ExcelConst_xlCmdDefault = 4,
	_Excel_XlCmdTypeForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_xlQueryType
{
	ExcelConst_xlODBCQuery = 1,
	ExcelConst_xlDAORecordSet = 2,
	ExcelConst_xlWebQuery = 4,
	ExcelConst_xlOLEDBQuery = 5,
	ExcelConst_xlTextImport = 6,
	ExcelConst_xlADORecordset = 7,
	_Excel_xlQueryTypeForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_xlWebSelectionType
{
	ExcelConst_xlEntirePage = 1,
	ExcelConst_xlAllTables = 2,
	ExcelConst_xlSpecifiedTables = 3,
	_Excel_xlWebSelectionTypeForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_xlWebFormatting
{
	ExcelConst_xlWebFormattingAll = 1,
	ExcelConst_xlWebFormattingRTF = 2,
	ExcelConst_xlWebFormattingNone = 3,
	_Excel_xlWebFormattingForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlMSApplication
{
	ExcelConst_xlMicrosoftAccess = 4,
	ExcelConst_xlMicrosoftFoxPro = 5,
	ExcelConst_xlMicrosoftMail = 3,
	ExcelConst_xlMicrosoftPowerPoint = 2,
	ExcelConst_xlMicrosoftProject = 6,
	ExcelConst_xlMicrosoftSchedulePlus = 7,
	ExcelConst_xlMicrosoftWord = 1,
	_Excel_XlMSApplicationForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlCalculation
{
	ExcelConst_xlCalculationAutomatic = -4105,
	ExcelConst_xlCalculationManual = -4135,
	ExcelConst_xlCalculationSemiautomatic = 2,
	_Excel_XlCalculationForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlCommandUnderlines
{
	ExcelConst_xlCommandUnderlinesAutomatic = -4105,
	ExcelConst_xlCommandUnderlinesOff = -4146,
	ExcelConst_xlCommandUnderlinesOn = 1,
	_Excel_XlCommandUnderlinesForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlReferenceStyle
{
	ExcelConst_xlA1 = 1,
	ExcelConst_xlR1C1 = -4150,
	_Excel_XlReferenceStyleForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlMousePointer
{
	ExcelConst_xlIBeam = 3,
	ExcelConst_xlDefault = -4143,
	ExcelConst_xlNorthwestArrow = 1,
	ExcelConst_xlWait = 2,
	_Excel_XlMousePointerForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlCutCopyMode
{
	ExcelConst_xlCopy = 1,
	ExcelConst_xlCut = 2,
	_Excel_XlCutCopyModeForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlCommentDisplayMode
{
	ExcelConst_xlNoIndicator = 0,
	ExcelConst_xlCommentIndicatorOnly = -1,
	ExcelConst_xlCommentAndIndicator = 1,
	_Excel_XlCommentDisplayModeForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlEnableCancelKey
{
	ExcelConst_xlDisabled = 0,
	ExcelConst_xlErrorHandler = 2,
	ExcelConst_xlInterrupt = 1,
	_Excel_XlEnableCancelKeyForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlMailSystem
{
	ExcelConst_xlMAPI = 1,
	ExcelConst_xlNoMailSystem = 0,
	ExcelConst_xlPowerTalk = 2,
	_Excel_XlMailSystemForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlDirection
{
	ExcelConst_xlDown = -4121,
	ExcelConst_xlToLeft = -4159,
	ExcelConst_xlToRight = -4161,
	ExcelConst_xlUp = -4162,
	_Excel_XlDirectionForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlFileFormat
{
	ExcelConst_xlAddIn = 18,
	ExcelConst_xlCSV = 6,
	ExcelConst_xlCSVMac = 22,
	ExcelConst_xlCSVMSDOS = 24,
	ExcelConst_xlCSVWindows = 23,
	ExcelConst_xlDBF2 = 7,
	ExcelConst_xlDBF3 = 8,
	ExcelConst_xlDBF4 = 11,
	ExcelConst_xlDIF = 9,
	ExcelConst_xlExcel2 = 16,
	ExcelConst_xlExcel2FarEast = 27,
	ExcelConst_xlExcel3 = 29,
	ExcelConst_xlExcel4 = 33,
	ExcelConst_xlExcel5 = 39,
	ExcelConst_xlExcel7 = 39,
	ExcelConst_xlExcel9795 = 43,
	ExcelConst_xlExcel4Workbook = 35,
	ExcelConst_xlIntlAddIn = 26,
	ExcelConst_xlIntlMacro = 25,
	ExcelConst_xlWorkbookNormal = -4143,
	ExcelConst_xlSYLK = 2,
	ExcelConst_xlTemplate = 17,
	ExcelConst_xlCurrentPlatformText = -4158,
	ExcelConst_xlTextMac = 19,
	ExcelConst_xlTextMSDOS = 21,
	ExcelConst_xlTextPrinter = 36,
	ExcelConst_xlTextWindows = 20,
	ExcelConst_xlWJ2WD1 = 14,
	ExcelConst_xlWK1 = 5,
	ExcelConst_xlWK1ALL = 31,
	ExcelConst_xlWK1FMT = 30,
	ExcelConst_xlWK3 = 15,
	ExcelConst_xlWK4 = 38,
	ExcelConst_xlWK3FM3 = 32,
	ExcelConst_xlWKS = 4,
	ExcelConst_xlWorks2FarEast = 28,
	ExcelConst_xlWQ1 = 34,
	ExcelConst_xlWJ3 = 40,
	ExcelConst_xlWJ3FJ3 = 41,
	ExcelConst_xlUnicodeText = 42,
	ExcelConst_xlHtml = 44,
	_Excel_XlFileFormatForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlWindowState
{
	ExcelConst_xlMaximized = -4137,
	ExcelConst_xlMinimized = -4140,
	ExcelConst_xlNormal = -4143,
	_Excel_XlWindowStateForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_MsoFeatureInstall
{
	ExcelConst_msoFeatureInstallNone = 0,
	ExcelConst_msoFeatureInstallOnDemand = 1,
	ExcelConst_msoFeatureInstallOnDemandWithUI = 2,
	_Excel_MsoFeatureInstallForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlSheetVisibility
{
	ExcelConst_xlSheetVisible = -1,
	ExcelConst_xlSheetHidden = 0,
	ExcelConst_xlSheetVeryHidden = 2,
	_Excel_XlSheetVisibilityForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlDataLabelsType
{
	ExcelConst_xlDataLabelsShowNone = -4142,
	ExcelConst_xlDataLabelsShowValue = 2,
	ExcelConst_xlDataLabelsShowPercent = 3,
	ExcelConst_xlDataLabelsShowLabel = 4,
	ExcelConst_xlDataLabelsShowLabelAndPercent = 5,
	ExcelConst_xlDataLabelsShowBubbleSizes = 6,
	_Excel_XlDataLabelsTypeForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlAxisGroup
{
	ExcelConst_xlPrimary = 1,
	ExcelConst_xlSecondary = 2,
	_Excel_XlAxisGroupForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlPictureAppearance
{
	ExcelConst_xlPrinter = 2,
	ExcelConst_xlScreen = 1,
	_Excel_XlPictureAppearanceForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlCopyPictureFormat
{
	ExcelConst_xlBitmap = 2,
	ExcelConst_xlPicture = -4147,
	_Excel_XlCopyPictureFormatForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlDisplayBlanksAs
{
	ExcelConst_xlInterpolated = 3,
	ExcelConst_xlNotPlotted = 1,
	ExcelConst_xlZero = 2,
	_Excel_XlDisplayBlanksAsForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlChartLocation
{
	ExcelConst_xlLocationAsNewSheet = 1,
	ExcelConst_xlLocationAsObject = 2,
	ExcelConst_xlLocationAutomatic = 3,
	_Excel_XlChartLocationForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlChartType
{
	ExcelConst_xlColumnClustered = 51,
	ExcelConst_xlColumnStacked = 52,
	ExcelConst_xlColumnStacked100 = 53,
	ExcelConst_xl3DColumnClustered = 54,
	ExcelConst_xl3DColumnStacked = 55,
	ExcelConst_xl3DColumnStacked100 = 56,
	ExcelConst_xlBarClustered = 57,
	ExcelConst_xlBarStacked = 58,
	ExcelConst_xlBarStacked100 = 59,
	ExcelConst_xl3DBarClustered = 60,
	ExcelConst_xl3DBarStacked = 61,
	ExcelConst_xl3DBarStacked100 = 62,
	ExcelConst_xlLineStacked = 63,
	ExcelConst_xlLineStacked100 = 64,
	ExcelConst_xlLineMarkers = 65,
	ExcelConst_xlLineMarkersStacked = 66,
	ExcelConst_xlLineMarkersStacked100 = 67,
	ExcelConst_xlPieOfPie = 68,
	ExcelConst_xlPieExploded = 69,
	ExcelConst_xl3DPieExploded = 70,
	ExcelConst_xlBarOfPie = 71,
	ExcelConst_xlXYScatterSmooth = 72,
	ExcelConst_xlXYScatterSmoothNoMarkers = 73,
	ExcelConst_xlXYScatterLines = 74,
	ExcelConst_xlXYScatterLinesNoMarkers = 75,
	ExcelConst_xlAreaStacked = 76,
	ExcelConst_xlAreaStacked100 = 77,
	ExcelConst_xl3DAreaStacked = 78,
	ExcelConst_xl3DAreaStacked100 = 79,
	ExcelConst_xlDoughnutExploded = 80,
	ExcelConst_xlRadarMarkers = 81,
	ExcelConst_xlRadarFilled = 82,
	ExcelConst_xlSurface = 83,
	ExcelConst_xlSurfaceWireframe = 84,
	ExcelConst_xlSurfaceTopView = 85,
	ExcelConst_xlSurfaceTopViewWireframe = 86,
	ExcelConst_xlBubble = 15,
	ExcelConst_xlBubble3DEffect = 87,
	ExcelConst_xlStockHLC = 88,
	ExcelConst_xlStockOHLC = 89,
	ExcelConst_xlStockVHLC = 90,
	ExcelConst_xlStockVOHLC = 91,
	ExcelConst_xlCylinderColClustered = 92,
	ExcelConst_xlCylinderColStacked = 93,
	ExcelConst_xlCylinderColStacked100 = 94,
	ExcelConst_xlCylinderBarClustered = 95,
	ExcelConst_xlCylinderBarStacked = 96,
	ExcelConst_xlCylinderBarStacked100 = 97,
	ExcelConst_xlCylinderCol = 98,
	ExcelConst_xlConeColClustered = 99,
	ExcelConst_xlConeColStacked = 100,
	ExcelConst_xlConeColStacked100 = 101,
	ExcelConst_xlConeBarClustered = 102,
	ExcelConst_xlConeBarStacked = 103,
	ExcelConst_xlConeBarStacked100 = 104,
	ExcelConst_xlConeCol = 105,
	ExcelConst_xlPyramidColClustered = 106,
	ExcelConst_xlPyramidColStacked = 107,
	ExcelConst_xlPyramidColStacked100 = 108,
	ExcelConst_xlPyramidBarClustered = 109,
	ExcelConst_xlPyramidBarStacked = 110,
	ExcelConst_xlPyramidBarStacked100 = 111,
	ExcelConst_xlPyramidCol = 112,
	ExcelConst_xl3DColumn = -4100,
	ExcelConst_xlLine = 4,
	ExcelConst_xl3DLine = -4101,
	ExcelConst_xl3DPie = -4102,
	ExcelConst_xlPie = 5,
	ExcelConst_xlXYScatter = -4169,
	ExcelConst_xl3DArea = -4098,
	ExcelConst_xlArea = 1,
	ExcelConst_xlDoughnut = -4120,
	ExcelConst_xlRadar = -4151,
	_Excel_XlChartTypeForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlBarShape
{
	ExcelConst_xlBox = 0,
	ExcelConst_xlPyramidToPoint = 1,
	ExcelConst_xlPyramidToMax = 2,
	ExcelConst_xlCylinder = 3,
	ExcelConst_xlConeToPoint = 4,
	ExcelConst_xlConeToMax = 5,
	_Excel_XlBarShapeForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlRowCol
{
	ExcelConst_xlColumns = 2,
	ExcelConst_xlRows = 1,
	_Excel_XlRowColForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlConsolidationFunction
{
	ExcelConst_xlAverage = -4106,
	ExcelConst_xlCount = -4112,
	ExcelConst_xlCountNums = -4113,
	ExcelConst_xlMax = -4136,
	ExcelConst_xlMin = -4139,
	ExcelConst_xlProduct = -4149,
	ExcelConst_xlStDev = -4155,
	ExcelConst_xlStDevP = -4156,
	ExcelConst_xlSum = -4157,
	ExcelConst_xlVar = -4164,
	ExcelConst_xlVarP = -4165,
	ExcelConst_xlUnknown = 1000,
	_Excel_XlConsolidationFunctionForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlEnableSelection
{
	ExcelConst_xlNoRestrictions = 0,
	ExcelConst_xlUnlockedCells = 1,
	ExcelConst_xlNoSelection = -4142,
	_Excel_XlEnableSelectionForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlSheetType
{
	ExcelConst_xlChart = -4109,
	ExcelConst_xlDialogSheet = -4116,
	ExcelConst_xlExcel4IntlMacroSheet = 4,
	ExcelConst_xlExcel4MacroSheet = 3,
	ExcelConst_xlWorksheet = -4167,
	_Excel_XlSheetTypeForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlFileAccess
{
	ExcelConst_xlReadOnly = 3,
	ExcelConst_xlReadWrite = 2,
	_Excel_XlFileAccessForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlLinkType
{
	ExcelConst_xlLinkTypeExcelLinks = 1,
	ExcelConst_xlLinkTypeOLELinks = 2,
	_Excel_XlLinkTypeForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlSaveConflictResolution
{
	ExcelConst_xlLocalSessionChanges = 2,
	ExcelConst_xlOtherSessionChanges = 3,
	ExcelConst_xlUserResolution = 1,
	_Excel_XlSaveConflictResolutionForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_xlDisplayDrawingObjects
{
	ExcelConst_xlDisplayShapes = -4104,
	ExcelConst_xlHide = 3,
	ExcelConst_xlPlaceholders = 2,
	_Excel_xlDisplayDrawingObjectsForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlLinkInfo
{
	ExcelConst_xlEditionDate = 2,
	ExcelConst_xlUpdateState = 1,
	_Excel_XlLinkInfoForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlRunAutoMacro
{
	ExcelConst_xlAutoActivate = 3,
	ExcelConst_xlAutoClose = 2,
	ExcelConst_xlAutoDeactivate = 4,
	ExcelConst_xlAutoOpen = 1,
	_Excel_XlRunAutoMacroForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlSaveAsAccessMode
{
	ExcelConst_xlExclusive = 3,
	ExcelConst_xlNoChange = 1,
	ExcelConst_xlShared = 2,
	_Excel_XlSaveAsAccessModeForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlPriority
{
	ExcelConst_xlPriorityHigh = -4127,
	ExcelConst_xlPriorityLow = -4134,
	ExcelConst_xlPriorityNormal = -4143,
	_Excel_XlPriorityForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_MsoEncoding
{
	ExcelConst_msoEncodingThai = 874,
	ExcelConst_msoEncodingJapaneseShiftJIS = 932,
	ExcelConst_msoEncodingSimplifiedChineseGBK = 936,
	ExcelConst_msoEncodingKorean = 949,
	ExcelConst_msoEncodingTraditionalChineseBig5 = 950,
	ExcelConst_msoEncodingUnicodeLittleEndian = 1200,
	ExcelConst_msoEncodingUnicodeBigEndian = 1201,
	ExcelConst_msoEncodingCentralEuropean = 1250,
	ExcelConst_msoEncodingCyrillic = 1251,
	ExcelConst_msoEncodingWestern = 1252,
	ExcelConst_msoEncodingGreek = 1253,
	ExcelConst_msoEncodingTurkish = 1254,
	ExcelConst_msoEncodingHebrew = 1255,
	ExcelConst_msoEncodingArabic = 1256,
	ExcelConst_msoEncodingBaltic = 1257,
	ExcelConst_msoEncodingVietnamese = 1258,
	ExcelConst_msoEncodingAutoDetect = 50001,
	ExcelConst_msoEncodingJapaneseAutoDetect = 50932,
	ExcelConst_msoEncodingSimplifiedChineseAutoDetect = 50936,
	ExcelConst_msoEncodingKoreanAutoDetect = 50949,
	ExcelConst_msoEncodingTraditionalChineseAutoDetect = 50950,
	ExcelConst_msoEncodingCyrillicAutoDetect = 51251,
	ExcelConst_msoEncodingGreekAutoDetect = 51253,
	ExcelConst_msoEncodingArabicAutoDetect = 51256,
	ExcelConst_msoEncodingISO88591Latin1 = 28591,
	ExcelConst_msoEncodingISO88592CentralEurope = 28592,
	ExcelConst_msoEncodingISO88593Latin3 = 28593,
	ExcelConst_msoEncodingISO88594Baltic = 28594,
	ExcelConst_msoEncodingISO88595Cyrillic = 28595,
	ExcelConst_msoEncodingISO88596Arabic = 28596,
	ExcelConst_msoEncodingISO88597Greek = 28597,
	ExcelConst_msoEncodingISO88598Hebrew = 28598,
	ExcelConst_msoEncodingISO88599Turkish = 28599,
	ExcelConst_msoEncodingISO885915Latin9 = 28605,
	ExcelConst_msoEncodingISO2022JPNoHalfwidthKatakana = 50220,
	ExcelConst_msoEncodingISO2022JPJISX02021984 = 50221,
	ExcelConst_msoEncodingISO2022JPJISX02011989 = 50222,
	ExcelConst_msoEncodingISO2022KR = 50225,
	ExcelConst_msoEncodingISO2022CNTraditionalChinese = 50227,
	ExcelConst_msoEncodingISO2022CNSimplifiedChinese = 50229,
	ExcelConst_msoEncodingMacRoman = 10000,
	ExcelConst_msoEncodingMacJapanese = 10001,
	ExcelConst_msoEncodingMacTraditionalChineseBig5 = 10002,
	ExcelConst_msoEncodingMacKorean = 10003,
	ExcelConst_msoEncodingMacArabic = 10004,
	ExcelConst_msoEncodingMacHebrew = 10005,
	ExcelConst_msoEncodingMacGreek1 = 10006,
	ExcelConst_msoEncodingMacCyrillic = 10007,
	ExcelConst_msoEncodingMacSimplifiedChineseGB2312 = 10008,
	ExcelConst_msoEncodingMacRomania = 10010,
	ExcelConst_msoEncodingMacUkraine = 10017,
	ExcelConst_msoEncodingMacLatin2 = 10029,
	ExcelConst_msoEncodingMacIcelandic = 10079,
	ExcelConst_msoEncodingMacTurkish = 10081,
	ExcelConst_msoEncodingMacCroatia = 10082,
	ExcelConst_msoEncodingEBCDICUSCanada = 37,
	ExcelConst_msoEncodingEBCDICInternational = 500,
	ExcelConst_msoEncodingEBCDICMultilingualROECELatin2 = 870,
	ExcelConst_msoEncodingEBCDICGreekModern = 875,
	ExcelConst_msoEncodingEBCDICTurkishLatin5 = 1026,
	ExcelConst_msoEncodingEBCDICGermany = 20273,
	ExcelConst_msoEncodingEBCDICDenmarkNorway = 20277,
	ExcelConst_msoEncodingEBCDICFinlandSweden = 20278,
	ExcelConst_msoEncodingEBCDICItaly = 20280,
	ExcelConst_msoEncodingEBCDICLatinAmericaSpain = 20284,
	ExcelConst_msoEncodingEBCDICUnitedKingdom = 20285,
	ExcelConst_msoEncodingEBCDICJapaneseKatakanaExtended = 20290,
	ExcelConst_msoEncodingEBCDICFrance = 20297,
	ExcelConst_msoEncodingEBCDICArabic = 20420,
	ExcelConst_msoEncodingEBCDICGreek = 20423,
	ExcelConst_msoEncodingEBCDICHebrew = 20424,
	ExcelConst_msoEncodingEBCDICKoreanExtended = 20833,
	ExcelConst_msoEncodingEBCDICThai = 20838,
	ExcelConst_msoEncodingEBCDICIcelandic = 20871,
	ExcelConst_msoEncodingEBCDICTurkish = 20905,
	ExcelConst_msoEncodingEBCDICRussian = 20880,
	ExcelConst_msoEncodingEBCDICSerbianBulgarian = 21025,
	ExcelConst_msoEncodingEBCDICJapaneseKatakanaExtendedAndJapanese = 50930,
	ExcelConst_msoEncodingEBCDICUSCanadaandJapanese = 50931,
	ExcelConst_msoEncodingEBCDICKoreanExtendedAndKorean = 50933,
	ExcelConst_msoEncodingEBCDICSimplifiedChineseExtendedAndSimplifiedChinese = 50935,
	ExcelConst_msoEncodingEBCDICUSCanadaAndTraditionalChinese = 50937,
	ExcelConst_msoEncodingEBCDICJapaneseLatinExtendedAndJapanese = 50939,
	ExcelConst_msoEncodingOEMUnitedStates = 437,
	ExcelConst_msoEncodingOEMGreek437G = 737,
	ExcelConst_msoEncodingOEMBaltic = 775,
	ExcelConst_msoEncodingOEMMultilingualLatinI = 850,
	ExcelConst_msoEncodingOEMMultilingualLatinII = 852,
	ExcelConst_msoEncodingOEMCyrillic = 855,
	ExcelConst_msoEncodingOEMTurkish = 857,
	ExcelConst_msoEncodingOEMPortuguese = 860,
	ExcelConst_msoEncodingOEMIcelandic = 861,
	ExcelConst_msoEncodingOEMHebrew = 862,
	ExcelConst_msoEncodingOEMCanadianFrench = 863,
	ExcelConst_msoEncodingOEMArabic = 864,
	ExcelConst_msoEncodingOEMNordic = 865,
	ExcelConst_msoEncodingOEMCyrillicII = 866,
	ExcelConst_msoEncodingOEMModernGreek = 869,
	ExcelConst_msoEncodingEUCJapanese = 51932,
	ExcelConst_msoEncodingEUCChineseSimplifiedChinese = 51936,
	ExcelConst_msoEncodingEUCKorean = 51949,
	ExcelConst_msoEncodingEUCTaiwaneseTraditionalChinese = 51950,
	ExcelConst_msoEncodingArabicASMO = 708,
	ExcelConst_msoEncodingArabicTransparentASMO = 720,
	ExcelConst_msoEncodingKoreanJohab = 1361,
	ExcelConst_msoEncodingTaiwanCNS = 20000,
	ExcelConst_msoEncodingTaiwanTCA = 20001,
	ExcelConst_msoEncodingTaiwanEten = 20002,
	ExcelConst_msoEncodingTaiwanIBM5550 = 20003,
	ExcelConst_msoEncodingTaiwanTeleText = 20004,
	ExcelConst_msoEncodingTaiwanWang = 20005,
	ExcelConst_msoEncodingIA5IRV = 20105,
	ExcelConst_msoEncodingIA5German = 20106,
	ExcelConst_msoEncodingIA5Swedish = 20107,
	ExcelConst_msoEncodingIA5Norwegian = 20108,
	ExcelConst_msoEncodingUSASCII = 20127,
	ExcelConst_msoEncodingT61 = 20261,
	ExcelConst_msoEncodingISO6937NonSpacingAccent = 20269,
	ExcelConst_msoEncodingKOI8R = 20866,
	ExcelConst_msoEncodingExtAlphaLowercase = 21027,
	ExcelConst_msoEncodingKOI8U = 21866,
	ExcelConst_msoEncodingEuropa3 = 29001,
	ExcelConst_msoEncodingHZGBSimplifiedChinese = 52936,
	ExcelConst_msoEncodingUTF7 = 65000,
	ExcelConst_msoEncodingUTF8 = 65001,
	_Excel_MsoEncodingForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlOLEVerb
{
	ExcelConst_xlVerbOpen = 2,
	ExcelConst_xlVerbPrimary = 1,
	_Excel_XlOLEVerbForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_MsoCalloutDropType
{
	ExcelConst_msoCalloutDropMixed = -2,
	ExcelConst_msoCalloutDropCustom = 1,
	ExcelConst_msoCalloutDropTop = 2,
	ExcelConst_msoCalloutDropCenter = 3,
	ExcelConst_msoCalloutDropBottom = 4,
	_Excel_MsoCalloutDropTypeForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_MsoTriState
{
	ExcelConst_msoTrue = -1,
	ExcelConst_msoFalse = 0,
	ExcelConst_msoCTrue = 1,
	ExcelConst_msoTriStateToggle = -3,
	ExcelConst_msoTriStateMixed = -2,
	_Excel_MsoTriStateForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_MsoCalloutAngleType
{
	ExcelConst_msoCalloutAngleMixed = -2,
	ExcelConst_msoCalloutAngleAutomatic = 1,
	ExcelConst_msoCalloutAngle30 = 2,
	ExcelConst_msoCalloutAngle45 = 3,
	ExcelConst_msoCalloutAngle60 = 4,
	ExcelConst_msoCalloutAngle90 = 5,
	_Excel_MsoCalloutAngleTypeForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_MsoCalloutType
{
	ExcelConst_msoCalloutMixed = -2,
	ExcelConst_msoCalloutOne = 1,
	ExcelConst_msoCalloutTwo = 2,
	ExcelConst_msoCalloutThree = 3,
	ExcelConst_msoCalloutFour = 4,
	_Excel_MsoCalloutTypeForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_MsoColorType
{
	ExcelConst_msoColorTypeMixed = -2,
	ExcelConst_msoColorTypeRGB = 1,
	ExcelConst_msoColorTypeScheme = 2,
	_Excel_MsoColorTypeForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_MsoArrowheadLength
{
	ExcelConst_msoArrowheadLengthMixed = -2,
	ExcelConst_msoArrowheadShort = 1,
	ExcelConst_msoArrowheadLengthMedium = 2,
	ExcelConst_msoArrowheadLong = 3,
	_Excel_MsoArrowheadLengthForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_MsoArrowheadStyle
{
	ExcelConst_msoArrowheadStyleMixed = -2,
	ExcelConst_msoArrowheadNone = 1,
	ExcelConst_msoArrowheadTriangle = 2,
	ExcelConst_msoArrowheadOpen = 3,
	ExcelConst_msoArrowheadStealth = 4,
	ExcelConst_msoArrowheadDiamond = 5,
	ExcelConst_msoArrowheadOval = 6,
	_Excel_MsoArrowheadStyleForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_MsoArrowheadWidth
{
	ExcelConst_msoArrowheadWidthMixed = -2,
	ExcelConst_msoArrowheadNarrow = 1,
	ExcelConst_msoArrowheadWidthMedium = 2,
	ExcelConst_msoArrowheadWide = 3,
	_Excel_MsoArrowheadWidthForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_MsoLineDashStyle
{
	ExcelConst_msoLineDashStyleMixed = -2,
	ExcelConst_msoLineSolid = 1,
	ExcelConst_msoLineSquareDot = 2,
	ExcelConst_msoLineRoundDot = 3,
	ExcelConst_msoLineDash = 4,
	ExcelConst_msoLineDashDot = 5,
	ExcelConst_msoLineDashDotDot = 6,
	ExcelConst_msoLineLongDash = 7,
	ExcelConst_msoLineLongDashDot = 8,
	_Excel_MsoLineDashStyleForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_MsoPatternType
{
	ExcelConst_msoPatternMixed = -2,
	ExcelConst_msoPattern5Percent = 1,
	ExcelConst_msoPattern10Percent = 2,
	ExcelConst_msoPattern20Percent = 3,
	ExcelConst_msoPattern25Percent = 4,
	ExcelConst_msoPattern30Percent = 5,
	ExcelConst_msoPattern40Percent = 6,
	ExcelConst_msoPattern50Percent = 7,
	ExcelConst_msoPattern60Percent = 8,
	ExcelConst_msoPattern70Percent = 9,
	ExcelConst_msoPattern75Percent = 10,
	ExcelConst_msoPattern80Percent = 11,
	ExcelConst_msoPattern90Percent = 12,
	ExcelConst_msoPatternDarkHorizontal = 13,
	ExcelConst_msoPatternDarkVertical = 14,
	ExcelConst_msoPatternDarkDownwardDiagonal = 15,
	ExcelConst_msoPatternDarkUpwardDiagonal = 16,
	ExcelConst_msoPatternSmallCheckerBoard = 17,
	ExcelConst_msoPatternTrellis = 18,
	ExcelConst_msoPatternLightHorizontal = 19,
	ExcelConst_msoPatternLightVertical = 20,
	ExcelConst_msoPatternLightDownwardDiagonal = 21,
	ExcelConst_msoPatternLightUpwardDiagonal = 22,
	ExcelConst_msoPatternSmallGrid = 23,
	ExcelConst_msoPatternDottedDiamond = 24,
	ExcelConst_msoPatternWideDownwardDiagonal = 25,
	ExcelConst_msoPatternWideUpwardDiagonal = 26,
	ExcelConst_msoPatternDashedUpwardDiagonal = 27,
	ExcelConst_msoPatternDashedDownwardDiagonal = 28,
	ExcelConst_msoPatternNarrowVertical = 29,
	ExcelConst_msoPatternNarrowHorizontal = 30,
	ExcelConst_msoPatternDashedVertical = 31,
	ExcelConst_msoPatternDashedHorizontal = 32,
	ExcelConst_msoPatternLargeConfetti = 33,
	ExcelConst_msoPatternLargeGrid = 34,
	ExcelConst_msoPatternHorizontalBrick = 35,
	ExcelConst_msoPatternLargeCheckerBoard = 36,
	ExcelConst_msoPatternSmallConfetti = 37,
	ExcelConst_msoPatternZigZag = 38,
	ExcelConst_msoPatternSolidDiamond = 39,
	ExcelConst_msoPatternDiagonalBrick = 40,
	ExcelConst_msoPatternOutlinedDiamond = 41,
	ExcelConst_msoPatternPlaid = 42,
	ExcelConst_msoPatternSphere = 43,
	ExcelConst_msoPatternWeave = 44,
	ExcelConst_msoPatternDottedGrid = 45,
	ExcelConst_msoPatternDivot = 46,
	ExcelConst_msoPatternShingle = 47,
	ExcelConst_msoPatternWave = 48,
	_Excel_MsoPatternTypeForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_MsoLineStyle
{
	ExcelConst_msoLineStyleMixed = -2,
	ExcelConst_msoLineSingle = 1,
	ExcelConst_msoLineThinThin = 2,
	ExcelConst_msoLineThinThick = 3,
	ExcelConst_msoLineThickThin = 4,
	ExcelConst_msoLineThickBetweenThin = 5,
	_Excel_MsoLineStyleForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_MsoEditingType
{
	ExcelConst_msoEditingAuto = 0,
	ExcelConst_msoEditingCorner = 1,
	ExcelConst_msoEditingSmooth = 2,
	ExcelConst_msoEditingSymmetric = 3,
	_Excel_MsoEditingTypeForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_MsoSegmentType
{
	ExcelConst_msoSegmentLine = 0,
	ExcelConst_msoSegmentCurve = 1,
	_Excel_MsoSegmentTypeForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_MsoPictureColorType
{
	ExcelConst_msoPictureMixed = -2,
	ExcelConst_msoPictureAutomatic = 1,
	ExcelConst_msoPictureGrayscale = 2,
	ExcelConst_msoPictureBlackAndWhite = 3,
	ExcelConst_msoPictureWatermark = 4,
	_Excel_MsoPictureColorTypeForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_MsoShadowType
{
	ExcelConst_msoShadowMixed = -2,
	ExcelConst_msoShadow1 = 1,
	ExcelConst_msoShadow2 = 2,
	ExcelConst_msoShadow3 = 3,
	ExcelConst_msoShadow4 = 4,
	ExcelConst_msoShadow5 = 5,
	ExcelConst_msoShadow6 = 6,
	ExcelConst_msoShadow7 = 7,
	ExcelConst_msoShadow8 = 8,
	ExcelConst_msoShadow9 = 9,
	ExcelConst_msoShadow10 = 10,
	ExcelConst_msoShadow11 = 11,
	ExcelConst_msoShadow12 = 12,
	ExcelConst_msoShadow13 = 13,
	ExcelConst_msoShadow14 = 14,
	ExcelConst_msoShadow15 = 15,
	ExcelConst_msoShadow16 = 16,
	ExcelConst_msoShadow17 = 17,
	ExcelConst_msoShadow18 = 18,
	ExcelConst_msoShadow19 = 19,
	ExcelConst_msoShadow20 = 20,
	_Excel_MsoShadowTypeForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_MsoTextEffectAlignment
{
	ExcelConst_msoTextEffectAlignmentMixed = -2,
	ExcelConst_msoTextEffectAlignmentLeft = 1,
	ExcelConst_msoTextEffectAlignmentCentered = 2,
	ExcelConst_msoTextEffectAlignmentRight = 3,
	ExcelConst_msoTextEffectAlignmentLetterJustify = 4,
	ExcelConst_msoTextEffectAlignmentWordJustify = 5,
	ExcelConst_msoTextEffectAlignmentStretchJustify = 6,
	_Excel_MsoTextEffectAlignmentForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_MsoPresetTextEffectShape
{
	ExcelConst_msoTextEffectShapeMixed = -2,
	ExcelConst_msoTextEffectShapePlainText = 1,
	ExcelConst_msoTextEffectShapeStop = 2,
	ExcelConst_msoTextEffectShapeTriangleUp = 3,
	ExcelConst_msoTextEffectShapeTriangleDown = 4,
	ExcelConst_msoTextEffectShapeChevronUp = 5,
	ExcelConst_msoTextEffectShapeChevronDown = 6,
	ExcelConst_msoTextEffectShapeRingInside = 7,
	ExcelConst_msoTextEffectShapeRingOutside = 8,
	ExcelConst_msoTextEffectShapeArchUpCurve = 9,
	ExcelConst_msoTextEffectShapeArchDownCurve = 10,
	ExcelConst_msoTextEffectShapeCircleCurve = 11,
	ExcelConst_msoTextEffectShapeButtonCurve = 12,
	ExcelConst_msoTextEffectShapeArchUpPour = 13,
	ExcelConst_msoTextEffectShapeArchDownPour = 14,
	ExcelConst_msoTextEffectShapeCirclePour = 15,
	ExcelConst_msoTextEffectShapeButtonPour = 16,
	ExcelConst_msoTextEffectShapeCurveUp = 17,
	ExcelConst_msoTextEffectShapeCurveDown = 18,
	ExcelConst_msoTextEffectShapeCanUp = 19,
	ExcelConst_msoTextEffectShapeCanDown = 20,
	ExcelConst_msoTextEffectShapeWave1 = 21,
	ExcelConst_msoTextEffectShapeWave2 = 22,
	ExcelConst_msoTextEffectShapeDoubleWave1 = 23,
	ExcelConst_msoTextEffectShapeDoubleWave2 = 24,
	ExcelConst_msoTextEffectShapeInflate = 25,
	ExcelConst_msoTextEffectShapeDeflate = 26,
	ExcelConst_msoTextEffectShapeInflateBottom = 27,
	ExcelConst_msoTextEffectShapeDeflateBottom = 28,
	ExcelConst_msoTextEffectShapeInflateTop = 29,
	ExcelConst_msoTextEffectShapeDeflateTop = 30,
	ExcelConst_msoTextEffectShapeDeflateInflate = 31,
	ExcelConst_msoTextEffectShapeDeflateInflateDeflate = 32,
	ExcelConst_msoTextEffectShapeFadeRight = 33,
	ExcelConst_msoTextEffectShapeFadeLeft = 34,
	ExcelConst_msoTextEffectShapeFadeUp = 35,
	ExcelConst_msoTextEffectShapeFadeDown = 36,
	ExcelConst_msoTextEffectShapeSlantUp = 37,
	ExcelConst_msoTextEffectShapeSlantDown = 38,
	ExcelConst_msoTextEffectShapeCascadeUp = 39,
	ExcelConst_msoTextEffectShapeCascadeDown = 40,
	_Excel_MsoPresetTextEffectShapeForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_MsoPresetTextEffect
{
	ExcelConst_msoTextEffectMixed = -2,
	ExcelConst_msoTextEffect1 = 0,
	ExcelConst_msoTextEffect2 = 1,
	ExcelConst_msoTextEffect3 = 2,
	ExcelConst_msoTextEffect4 = 3,
	ExcelConst_msoTextEffect5 = 4,
	ExcelConst_msoTextEffect6 = 5,
	ExcelConst_msoTextEffect7 = 6,
	ExcelConst_msoTextEffect8 = 7,
	ExcelConst_msoTextEffect9 = 8,
	ExcelConst_msoTextEffect10 = 9,
	ExcelConst_msoTextEffect11 = 10,
	ExcelConst_msoTextEffect12 = 11,
	ExcelConst_msoTextEffect13 = 12,
	ExcelConst_msoTextEffect14 = 13,
	ExcelConst_msoTextEffect15 = 14,
	ExcelConst_msoTextEffect16 = 15,
	ExcelConst_msoTextEffect17 = 16,
	ExcelConst_msoTextEffect18 = 17,
	ExcelConst_msoTextEffect19 = 18,
	ExcelConst_msoTextEffect20 = 19,
	ExcelConst_msoTextEffect21 = 20,
	ExcelConst_msoTextEffect22 = 21,
	ExcelConst_msoTextEffect23 = 22,
	ExcelConst_msoTextEffect24 = 23,
	ExcelConst_msoTextEffect25 = 24,
	ExcelConst_msoTextEffect26 = 25,
	ExcelConst_msoTextEffect27 = 26,
	ExcelConst_msoTextEffect28 = 27,
	ExcelConst_msoTextEffect29 = 28,
	ExcelConst_msoTextEffect30 = 29,
	_Excel_MsoPresetTextEffectForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_MsoPresetThreeDFormat
{
	ExcelConst_msoPresetThreeDFormatMixed = -2,
	ExcelConst_msoThreeD1 = 1,
	ExcelConst_msoThreeD2 = 2,
	ExcelConst_msoThreeD3 = 3,
	ExcelConst_msoThreeD4 = 4,
	ExcelConst_msoThreeD5 = 5,
	ExcelConst_msoThreeD6 = 6,
	ExcelConst_msoThreeD7 = 7,
	ExcelConst_msoThreeD8 = 8,
	ExcelConst_msoThreeD9 = 9,
	ExcelConst_msoThreeD10 = 10,
	ExcelConst_msoThreeD11 = 11,
	ExcelConst_msoThreeD12 = 12,
	ExcelConst_msoThreeD13 = 13,
	ExcelConst_msoThreeD14 = 14,
	ExcelConst_msoThreeD15 = 15,
	ExcelConst_msoThreeD16 = 16,
	ExcelConst_msoThreeD17 = 17,
	ExcelConst_msoThreeD18 = 18,
	ExcelConst_msoThreeD19 = 19,
	ExcelConst_msoThreeD20 = 20,
	_Excel_MsoPresetThreeDFormatForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_MsoPresetExtrusionDirection
{
	ExcelConst_msoPresetExtrusionDirectionMixed = -2,
	ExcelConst_msoExtrusionBottomRight = 1,
	ExcelConst_msoExtrusionBottom = 2,
	ExcelConst_msoExtrusionBottomLeft = 3,
	ExcelConst_msoExtrusionRight = 4,
	ExcelConst_msoExtrusionNone = 5,
	ExcelConst_msoExtrusionLeft = 6,
	ExcelConst_msoExtrusionTopRight = 7,
	ExcelConst_msoExtrusionTop = 8,
	ExcelConst_msoExtrusionTopLeft = 9,
	_Excel_MsoPresetExtrusionDirectionForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_MsoExtrusionColorType
{
	ExcelConst_msoExtrusionColorTypeMixed = -2,
	ExcelConst_msoExtrusionColorAutomatic = 1,
	ExcelConst_msoExtrusionColorCustom = 2,
	_Excel_MsoExtrusionColorTypeForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_MsoPresetLightingDirection
{
	ExcelConst_msoPresetLightingDirectionMixed = -2,
	ExcelConst_msoLightingTopLeft = 1,
	ExcelConst_msoLightingTop = 2,
	ExcelConst_msoLightingTopRight = 3,
	ExcelConst_msoLightingLeft = 4,
	ExcelConst_msoLightingNone = 5,
	ExcelConst_msoLightingRight = 6,
	ExcelConst_msoLightingBottomLeft = 7,
	ExcelConst_msoLightingBottom = 8,
	ExcelConst_msoLightingBottomRight = 9,
	_Excel_MsoPresetLightingDirectionForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_MsoPresetLightingSoftness
{
	ExcelConst_msoPresetLightingSoftnessMixed = -2,
	ExcelConst_msoLightingDim = 1,
	ExcelConst_msoLightingNormal = 2,
	ExcelConst_msoLightingBright = 3,
	_Excel_MsoPresetLightingSoftnessForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_MsoPresetMaterial
{
	ExcelConst_msoPresetMaterialMixed = -2,
	ExcelConst_msoMaterialMatte = 1,
	ExcelConst_msoMaterialPlastic = 2,
	ExcelConst_msoMaterialMetal = 3,
	ExcelConst_msoMaterialWireFrame = 4,
	_Excel_MsoPresetMaterialForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_MsoGradientStyle
{
	ExcelConst_msoGradientMixed = -2,
	ExcelConst_msoGradientHorizontal = 1,
	ExcelConst_msoGradientVertical = 2,
	ExcelConst_msoGradientDiagonalUp = 3,
	ExcelConst_msoGradientDiagonalDown = 4,
	ExcelConst_msoGradientFromCorner = 5,
	ExcelConst_msoGradientFromTitle = 6,
	ExcelConst_msoGradientFromCenter = 7,
	_Excel_MsoGradientStyleForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_MsoPresetGradientType
{
	ExcelConst_msoPresetGradientMixed = -2,
	ExcelConst_msoGradientEarlySunset = 1,
	ExcelConst_msoGradientLateSunset = 2,
	ExcelConst_msoGradientNightfall = 3,
	ExcelConst_msoGradientDaybreak = 4,
	ExcelConst_msoGradientHorizon = 5,
	ExcelConst_msoGradientDesert = 6,
	ExcelConst_msoGradientOcean = 7,
	ExcelConst_msoGradientCalmWater = 8,
	ExcelConst_msoGradientFire = 9,
	ExcelConst_msoGradientFog = 10,
	ExcelConst_msoGradientMoss = 11,
	ExcelConst_msoGradientPeacock = 12,
	ExcelConst_msoGradientWheat = 13,
	ExcelConst_msoGradientParchment = 14,
	ExcelConst_msoGradientMahogany = 15,
	ExcelConst_msoGradientRainbow = 16,
	ExcelConst_msoGradientRainbowII = 17,
	ExcelConst_msoGradientGold = 18,
	ExcelConst_msoGradientGoldII = 19,
	ExcelConst_msoGradientBrass = 20,
	ExcelConst_msoGradientChrome = 21,
	ExcelConst_msoGradientChromeII = 22,
	ExcelConst_msoGradientSilver = 23,
	ExcelConst_msoGradientSapphire = 24,
	_Excel_MsoPresetGradientTypeForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_MsoPresetTexture
{
	ExcelConst_msoPresetTextureMixed = -2,
	ExcelConst_msoTexturePapyrus = 1,
	ExcelConst_msoTextureCanvas = 2,
	ExcelConst_msoTextureDenim = 3,
	ExcelConst_msoTextureWovenMat = 4,
	ExcelConst_msoTextureWaterDroplets = 5,
	ExcelConst_msoTexturePaperBag = 6,
	ExcelConst_msoTextureFishFossil = 7,
	ExcelConst_msoTextureSand = 8,
	ExcelConst_msoTextureGreenMarble = 9,
	ExcelConst_msoTextureWhiteMarble = 10,
	ExcelConst_msoTextureBrownMarble = 11,
	ExcelConst_msoTextureGranite = 12,
	ExcelConst_msoTextureNewsprint = 13,
	ExcelConst_msoTextureRecycledPaper = 14,
	ExcelConst_msoTextureParchment = 15,
	ExcelConst_msoTextureStationery = 16,
	ExcelConst_msoTextureBlueTissuePaper = 17,
	ExcelConst_msoTexturePinkTissuePaper = 18,
	ExcelConst_msoTexturePurpleMesh = 19,
	ExcelConst_msoTextureBouquet = 20,
	ExcelConst_msoTextureCork = 21,
	ExcelConst_msoTextureWalnut = 22,
	ExcelConst_msoTextureOak = 23,
	ExcelConst_msoTextureMediumWood = 24,
	_Excel_MsoPresetTextureForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_MsoGradientColorType
{
	ExcelConst_msoGradientColorMixed = -2,
	ExcelConst_msoGradientOneColor = 1,
	ExcelConst_msoGradientTwoColors = 2,
	ExcelConst_msoGradientPresetColors = 3,
	_Excel_MsoGradientColorTypeForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_MsoTextureType
{
	ExcelConst_msoTextureTypeMixed = -2,
	ExcelConst_msoTexturePreset = 1,
	ExcelConst_msoTextureUserDefined = 2,
	_Excel_MsoTextureTypeForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_MsoFillType
{
	ExcelConst_msoFillMixed = -2,
	ExcelConst_msoFillSolid = 1,
	ExcelConst_msoFillPatterned = 2,
	ExcelConst_msoFillGradient = 3,
	ExcelConst_msoFillTextured = 4,
	ExcelConst_msoFillBackground = 5,
	ExcelConst_msoFillPicture = 6,
	_Excel_MsoFillTypeForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlFillWith
{
	ExcelConst_xlFillWithAll = -4104,
	ExcelConst_xlFillWithContents = 2,
	ExcelConst_xlFillWithFormats = -4122,
	_Excel_XlFillWithForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlSourceType
{
	ExcelConst_xlSourceSheet = 1,
	ExcelConst_xlSourcePrintArea = 2,
	ExcelConst_xlSourceAutoFilter = 3,
	ExcelConst_xlSourceRange = 4,
	ExcelConst_xlSourceChart = 5,
	ExcelConst_xlSourcePivotTable = 6,
	ExcelConst_xlSourceQuery = 7,
	_Excel_XlSourceTypeForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlHtmlType
{
	ExcelConst_xlHtmlStatic = 0,
	ExcelConst_xlHtmlCalc = 1,
	ExcelConst_xlHtmlList = 2,
	ExcelConst_xlHtmlChart = 3,
	_Excel_XlHtmlTypeForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_MsoScreenSize
{
	ExcelConst_msoScreenSize544x376 = 0,
	ExcelConst_msoScreenSize640x480 = 1,
	ExcelConst_msoScreenSize720x512 = 2,
	ExcelConst_msoScreenSize800x600 = 3,
	ExcelConst_msoScreenSize1024x768 = 4,
	ExcelConst_msoScreenSize1152x882 = 5,
	ExcelConst_msoScreenSize1152x900 = 6,
	ExcelConst_msoScreenSize1280x1024 = 7,
	ExcelConst_msoScreenSize1600x1200 = 8,
	ExcelConst_msoScreenSize1800x1440 = 9,
	ExcelConst_msoScreenSize1920x1200 = 10,
	_Excel_MsoScreenSizeForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlCubeFieldType
{
	ExcelConst_xlHierarchy = 1,
	ExcelConst_xlMeasure = 2,
	_Excel_XlCubeFieldTypeForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlPivotFieldOrientation
{
	ExcelConst_xlColumnField = 2,
	ExcelConst_xlDataField = 4,
	ExcelConst_xlHidden = 0,
	ExcelConst_xlPageField = 3,
	ExcelConst_xlRowField = 1,
	_Excel_XlPivotFieldOrientationForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlColorIndex
{
	ExcelConst_xlColorIndexAutomatic = -4105,
	ExcelConst_xlColorIndexNone = -4142,
	_Excel_XlColorIndexForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlWindowType
{
	ExcelConst_xlChartAsWindow = 5,
	ExcelConst_xlChartInPlace = 4,
	ExcelConst_xlClipboard = 3,
	ExcelConst_xlInfo = -4129,
	ExcelConst_xlWorkbook = 1,
	_Excel_XlWindowTypeForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlWindowView
{
	ExcelConst_xlNormalView = 1,
	ExcelConst_xlPageBreakPreview = 2,
	_Excel_XlWindowViewForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlArrangeStyle
{
	ExcelConst_xlArrangeStyleCascade = 7,
	ExcelConst_xlArrangeStyleHorizontal = -4128,
	ExcelConst_xlArrangeStyleTiled = 1,
	ExcelConst_xlArrangeStyleVertical = -4166,
	_Excel_XlArrangeStyleForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlFilterAction
{
	ExcelConst_xlFilterCopy = 2,
	ExcelConst_xlFilterInPlace = 1,
	_Excel_XlFilterActionForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlApplyNamesOrder
{
	ExcelConst_xlColumnThenRow = 2,
	ExcelConst_xlRowThenColumn = 1,
	_Excel_XlApplyNamesOrderForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlAutoFillType
{
	ExcelConst_xlFillCopy = 1,
	ExcelConst_xlFillDays = 5,
	ExcelConst_xlFillDefault = 0,
	ExcelConst_xlFillFormats = 3,
	ExcelConst_xlFillMonths = 7,
	ExcelConst_xlFillSeries = 2,
	ExcelConst_xlFillValues = 4,
	ExcelConst_xlFillWeekdays = 6,
	ExcelConst_xlFillYears = 8,
	ExcelConst_xlGrowthTrend = 10,
	ExcelConst_xlLinearTrend = 9,
	_Excel_XlAutoFillTypeForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlAutoFilterOperator
{
	ExcelConst_xlAnd = 1,
	ExcelConst_xlBottom10Items = 4,
	ExcelConst_xlBottom10Percent = 6,
	ExcelConst_xlOr = 2,
	ExcelConst_xlTop10Items = 3,
	ExcelConst_xlTop10Percent = 5,
	_Excel_XlAutoFilterOperatorForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlRangeAutoFormat
{
	ExcelConst_xlRangeAutoFormat3DEffects1 = 13,
	ExcelConst_xlRangeAutoFormat3DEffects2 = 14,
	ExcelConst_xlRangeAutoFormatAccounting1 = 4,
	ExcelConst_xlRangeAutoFormatAccounting2 = 5,
	ExcelConst_xlRangeAutoFormatAccounting3 = 6,
	ExcelConst_xlRangeAutoFormatAccounting4 = 17,
	ExcelConst_xlRangeAutoFormatClassic1 = 1,
	ExcelConst_xlRangeAutoFormatClassic2 = 2,
	ExcelConst_xlRangeAutoFormatClassic3 = 3,
	ExcelConst_xlRangeAutoFormatColor1 = 7,
	ExcelConst_xlRangeAutoFormatColor2 = 8,
	ExcelConst_xlRangeAutoFormatColor3 = 9,
	ExcelConst_xlRangeAutoFormatList1 = 10,
	ExcelConst_xlRangeAutoFormatList2 = 11,
	ExcelConst_xlRangeAutoFormatList3 = 12,
	ExcelConst_xlRangeAutoFormatLocalFormat1 = 15,
	ExcelConst_xlRangeAutoFormatLocalFormat2 = 16,
	ExcelConst_xlRangeAutoFormatLocalFormat3 = 19,
	ExcelConst_xlRangeAutoFormatLocalFormat4 = 20,
	ExcelConst_xlRangeAutoFormatReport1 = 21,
	ExcelConst_xlRangeAutoFormatReport2 = 22,
	ExcelConst_xlRangeAutoFormatReport3 = 23,
	ExcelConst_xlRangeAutoFormatReport4 = 24,
	ExcelConst_xlRangeAutoFormatReport5 = 25,
	ExcelConst_xlRangeAutoFormatReport6 = 26,
	ExcelConst_xlRangeAutoFormatReport7 = 27,
	ExcelConst_xlRangeAutoFormatReport8 = 28,
	ExcelConst_xlRangeAutoFormatReport9 = 29,
	ExcelConst_xlRangeAutoFormatReport10 = 30,
	ExcelConst_xlRangeAutoFormatClassicPivotTable = 31,
	ExcelConst_xlRangeAutoFormatTable1 = 32,
	ExcelConst_xlRangeAutoFormatTable2 = 33,
	ExcelConst_xlRangeAutoFormatTable3 = 34,
	ExcelConst_xlRangeAutoFormatTable4 = 35,
	ExcelConst_xlRangeAutoFormatTable5 = 36,
	ExcelConst_xlRangeAutoFormatTable6 = 37,
	ExcelConst_xlRangeAutoFormatTable7 = 38,
	ExcelConst_xlRangeAutoFormatTable8 = 39,
	ExcelConst_xlRangeAutoFormatTable9 = 40,
	ExcelConst_xlRangeAutoFormatTable10 = 41,
	ExcelConst_xlRangeAutoFormatPTNone = 42,
	ExcelConst_xlRangeAutoFormatNone = -4142,
	ExcelConst_xlRangeAutoFormatSimple = -4154,
	_Excel_XlRangeAutoFormatForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlBorderWeight
{
	ExcelConst_xlHairline = 1,
	ExcelConst_xlMedium = -4138,
	ExcelConst_xlThick = 4,
	ExcelConst_xlThin = 2,
	_Excel_XlBorderWeightForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlDataSeriesType
{
	ExcelConst_xlAutoFill = 4,
	ExcelConst_xlChronological = 3,
	ExcelConst_xlGrowth = 2,
	ExcelConst_xlDataSeriesLinear = -4132,
	_Excel_XlDataSeriesTypeForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlDataSeriesDate
{
	ExcelConst_xlDay = 1,
	ExcelConst_xlMonth = 3,
	ExcelConst_xlWeekday = 2,
	ExcelConst_xlYear = 4,
	_Excel_XlDataSeriesDateForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlEditionType
{
	ExcelConst_xlPublisher = 1,
	ExcelConst_xlSubscriber = 2,
	_Excel_XlEditionTypeForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlEditionOptionsOption
{
	ExcelConst_xlAutomaticUpdate = 4,
	ExcelConst_xlCancel = 1,
	ExcelConst_xlChangeAttributes = 6,
	ExcelConst_xlManualUpdate = 5,
	ExcelConst_xlOpenSource = 3,
	ExcelConst_xlSelect = 3,
	ExcelConst_xlSendPublisher = 2,
	ExcelConst_xlUpdateSubscriber = 2,
	_Excel_XlEditionOptionsOptionForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlSearchDirection
{
	ExcelConst_xlNext = 1,
	ExcelConst_xlPrevious = 2,
	_Excel_XlSearchDirectionForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlFormulaLabel
{
	ExcelConst_xlNoLabels = -4142,
	ExcelConst_xlRowLabels = 1,
	ExcelConst_xlColumnLabels = 2,
	ExcelConst_xlMixedLabels = 3,
	_Excel_XlFormulaLabelForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlLocationInTable
{
	ExcelConst_xlColumnHeader = -4110,
	ExcelConst_xlColumnItem = 5,
	ExcelConst_xlDataHeader = 3,
	ExcelConst_xlDataItem = 7,
	ExcelConst_xlPageHeader = 2,
	ExcelConst_xlPageItem = 6,
	ExcelConst_xlRowHeader = -4153,
	ExcelConst_xlRowItem = 4,
	ExcelConst_xlTableBody = 8,
	_Excel_XlLocationInTableForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlPasteType
{
	ExcelConst_xlPasteAll = -4104,
	ExcelConst_xlPasteAllExceptBorders = 7,
	ExcelConst_xlPasteFormats = -4122,
	ExcelConst_xlPasteFormulas = -4123,
	ExcelConst_xlPasteComments = -4144,
	ExcelConst_xlPasteValues = -4163,
	_Excel_XlPasteTypeForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlPasteSpecialOperation
{
	ExcelConst_xlPasteSpecialOperationAdd = 2,
	ExcelConst_xlPasteSpecialOperationDivide = 5,
	ExcelConst_xlPasteSpecialOperationMultiply = 4,
	ExcelConst_xlPasteSpecialOperationNone = -4142,
	ExcelConst_xlPasteSpecialOperationSubtract = 3,
	_Excel_XlPasteSpecialOperationForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlSortOrder
{
	ExcelConst_xlAscending = 1,
	ExcelConst_xlDescending = 2,
	_Excel_XlSortOrderForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlYesNoGuess
{
	ExcelConst_xlGuess = 0,
	ExcelConst_xlNo = 2,
	ExcelConst_xlYes = 1,
	_Excel_XlYesNoGuessForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlSortOrientation
{
	ExcelConst_xlSortRows = 2,
	ExcelConst_xlSortColumns = 1,
	_Excel_XlSortOrientationForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlSortMethod
{
	ExcelConst_xlPinYin = 1,
	ExcelConst_xlStroke = 2,
	_Excel_XlSortMethodForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlCellType
{
	ExcelConst_xlCellTypeBlanks = 4,
	ExcelConst_xlCellTypeConstants = 2,
	ExcelConst_xlCellTypeFormulas = -4123,
	ExcelConst_xlCellTypeLastCell = 11,
	ExcelConst_xlCellTypeComments = -4144,
	ExcelConst_xlCellTypeVisible = 12,
	ExcelConst_xlCellTypeAllFormatConditions = -4172,
	ExcelConst_xlCellTypeSameFormatConditions = -4173,
	ExcelConst_xlCellTypeAllValidation = -4174,
	ExcelConst_xlCellTypeSameValidation = -4175,
	_Excel_XlCellTypeForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlSubscribeToFormat
{
	ExcelConst_xlSubscribeToPicture = -4147,
	ExcelConst_xlSubscribeToText = -4158,
	_Excel_XlSubscribeToFormatForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlSummaryRow
{
	ExcelConst_xlSummaryAbove = 0,
	ExcelConst_xlSummaryBelow = 1,
	_Excel_XlSummaryRowForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlPageBreak
{
	ExcelConst_xlPageBreakAutomatic = -4105,
	ExcelConst_xlPageBreakManual = -4135,
	ExcelConst_xlPageBreakNone = -4142,
	_Excel_XlPageBreakForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlPageBreakExtent
{
	ExcelConst_xlPageBreakFull = 1,
	ExcelConst_xlPageBreakPartial = 2,
	_Excel_XlPageBreakExtentForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlHAlign
{
	ExcelConst_xlHAlignCenter = -4108,
	ExcelConst_xlHAlignCenterAcrossSelection = 7,
	ExcelConst_xlHAlignDistributed = -4117,
	ExcelConst_xlHAlignFill = 5,
	ExcelConst_xlHAlignGeneral = 1,
	ExcelConst_xlHAlignJustify = -4130,
	ExcelConst_xlHAlignLeft = -4131,
	ExcelConst_xlHAlignRight = -4152,
	_Excel_XlHAlignForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlOrientation
{
	ExcelConst_xlDownward = -4170,
	ExcelConst_xlHorizontal = -4128,
	ExcelConst_xlUpward = -4171,
	ExcelConst_xlVertical = -4166,
	_Excel_XlOrientationForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlVAlign
{
	ExcelConst_xlVAlignBottom = -4107,
	ExcelConst_xlVAlignCenter = -4108,
	ExcelConst_xlVAlignDistributed = -4117,
	ExcelConst_xlVAlignJustify = -4130,
	ExcelConst_xlVAlignTop = -4160,
	_Excel_XlVAlignForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlBordersIndex
{
	ExcelConst_xlInsideHorizontal = 12,
	ExcelConst_xlInsideVertical = 11,
	ExcelConst_xlDiagonalDown = 5,
	ExcelConst_xlDiagonalUp = 6,
	ExcelConst_xlEdgeBottom = 9,
	ExcelConst_xlEdgeLeft = 7,
	ExcelConst_xlEdgeRight = 10,
	ExcelConst_xlEdgeTop = 8,
	_Excel_XlBordersIndexForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlPivotTableSourceType
{
	ExcelConst_xlConsolidation = 3,
	ExcelConst_xlDatabase = 1,
	ExcelConst_xlExternal = 2,
	ExcelConst_xlPivotTable = -4148,
	_Excel_XlPivotTableSourceTypeForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlPTSelectionMode
{
	ExcelConst_xlLabelOnly = 1,
	ExcelConst_xlDataAndLabel = 0,
	ExcelConst_xlDataOnly = 2,
	ExcelConst_xlOrigin = 3,
	ExcelConst_xlButton = 15,
	ExcelConst_xlBlanks = 4,
	ExcelConst_xlFirstRow = 256,
	_Excel_XlPTSelectionModeForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_xlPivotFormatType
{
	ExcelConst_xlReport1 = 0,
	ExcelConst_xlReport2 = 1,
	ExcelConst_xlReport3 = 2,
	ExcelConst_xlReport4 = 3,
	ExcelConst_xlReport5 = 4,
	ExcelConst_xlReport6 = 5,
	ExcelConst_xlReport7 = 6,
	ExcelConst_xlReport8 = 7,
	ExcelConst_xlReport9 = 8,
	ExcelConst_xlReport10 = 9,
	ExcelConst_xlTable1 = 10,
	ExcelConst_xlTable2 = 11,
	ExcelConst_xlTable3 = 12,
	ExcelConst_xlTable4 = 13,
	ExcelConst_xlTable5 = 14,
	ExcelConst_xlTable6 = 15,
	ExcelConst_xlTable7 = 16,
	ExcelConst_xlTable8 = 17,
	ExcelConst_xlTable9 = 18,
	ExcelConst_xlTable10 = 19,
	ExcelConst_xlPTClassic = 20,
	ExcelConst_xlPTNone = 21,
	_Excel_xlPivotFormatTypeForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlPivotFieldCalculation
{
	ExcelConst_xlDifferenceFrom = 2,
	ExcelConst_xlIndex = 9,
	ExcelConst_xlNoAdditionalCalculation = -4143,
	ExcelConst_xlPercentDifferenceFrom = 4,
	ExcelConst_xlPercentOf = 3,
	ExcelConst_xlPercentOfColumn = 7,
	ExcelConst_xlPercentOfRow = 6,
	ExcelConst_xlPercentOfTotal = 8,
	ExcelConst_xlRunningTotal = 5,
	_Excel_XlPivotFieldCalculationForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlPivotFieldDataType
{
	ExcelConst_xlDate = 2,
	ExcelConst_xlNumber = -4145,
	ExcelConst_xlText = -4158,
	_Excel_XlPivotFieldDataTypeForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_xLSubtototalLocationType
{
	ExcelConst_xlAtTop = 1,
	ExcelConst_xlAtBottom = 2,
	_Excel_xLSubtototalLocationTypeForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlLayoutFormType
{
	ExcelConst_xlTabular = 0,
	ExcelConst_xlOutline = 1,
	_Excel_XlLayoutFormTypeForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlBuiltInDialog
{
	ExcelConst_xlDialogOpen = 1,
	ExcelConst_xlDialogOpenLinks = 2,
	ExcelConst_xlDialogSaveAs = 5,
	ExcelConst_xlDialogFileDelete = 6,
	ExcelConst_xlDialogPageSetup = 7,
	ExcelConst_xlDialogPrint = 8,
	ExcelConst_xlDialogPrinterSetup = 9,
	ExcelConst_xlDialogArrangeAll = 12,
	ExcelConst_xlDialogWindowSize = 13,
	ExcelConst_xlDialogWindowMove = 14,
	ExcelConst_xlDialogRun = 17,
	ExcelConst_xlDialogSetPrintTitles = 23,
	ExcelConst_xlDialogFont = 26,
	ExcelConst_xlDialogDisplay = 27,
	ExcelConst_xlDialogProtectDocument = 28,
	ExcelConst_xlDialogCalculation = 32,
	ExcelConst_xlDialogExtract = 35,
	ExcelConst_xlDialogDataDelete = 36,
	ExcelConst_xlDialogSort = 39,
	ExcelConst_xlDialogDataSeries = 40,
	ExcelConst_xlDialogTable = 41,
	ExcelConst_xlDialogFormatNumber = 42,
	ExcelConst_xlDialogAlignment = 43,
	ExcelConst_xlDialogStyle = 44,
	ExcelConst_xlDialogBorder = 45,
	ExcelConst_xlDialogCellProtection = 46,
	ExcelConst_xlDialogColumnWidth = 47,
	ExcelConst_xlDialogClear = 52,
	ExcelConst_xlDialogPasteSpecial = 53,
	ExcelConst_xlDialogEditDelete = 54,
	ExcelConst_xlDialogInsert = 55,
	ExcelConst_xlDialogPasteNames = 58,
	ExcelConst_xlDialogDefineName = 61,
	ExcelConst_xlDialogCreateNames = 62,
	ExcelConst_xlDialogFormulaGoto = 63,
	ExcelConst_xlDialogFormulaFind = 64,
	ExcelConst_xlDialogGalleryArea = 67,
	ExcelConst_xlDialogGalleryBar = 68,
	ExcelConst_xlDialogGalleryColumn = 69,
	ExcelConst_xlDialogGalleryLine = 70,
	ExcelConst_xlDialogGalleryPie = 71,
	ExcelConst_xlDialogGalleryScatter = 72,
	ExcelConst_xlDialogCombination = 73,
	ExcelConst_xlDialogGridlines = 76,
	ExcelConst_xlDialogAxes = 78,
	ExcelConst_xlDialogAttachText = 80,
	ExcelConst_xlDialogPatterns = 84,
	ExcelConst_xlDialogMainChart = 85,
	ExcelConst_xlDialogOverlay = 86,
	ExcelConst_xlDialogScale = 87,
	ExcelConst_xlDialogFormatLegend = 88,
	ExcelConst_xlDialogFormatText = 89,
	ExcelConst_xlDialogParse = 91,
	ExcelConst_xlDialogUnhide = 94,
	ExcelConst_xlDialogWorkspace = 95,
	ExcelConst_xlDialogActivate = 103,
	ExcelConst_xlDialogCopyPicture = 108,
	ExcelConst_xlDialogDeleteName = 110,
	ExcelConst_xlDialogDeleteFormat = 111,
	ExcelConst_xlDialogNew = 119,
	ExcelConst_xlDialogRowHeight = 127,
	ExcelConst_xlDialogFormatMove = 128,
	ExcelConst_xlDialogFormatSize = 129,
	ExcelConst_xlDialogFormulaReplace = 130,
	ExcelConst_xlDialogSelectSpecial = 132,
	ExcelConst_xlDialogApplyNames = 133,
	ExcelConst_xlDialogReplaceFont = 134,
	ExcelConst_xlDialogSplit = 137,
	ExcelConst_xlDialogOutline = 142,
	ExcelConst_xlDialogSaveWorkbook = 145,
	ExcelConst_xlDialogCopyChart = 147,
	ExcelConst_xlDialogFormatFont = 150,
	ExcelConst_xlDialogNote = 154,
	ExcelConst_xlDialogSetUpdateStatus = 159,
	ExcelConst_xlDialogColorPalette = 161,
	ExcelConst_xlDialogChangeLink = 166,
	ExcelConst_xlDialogAppMove = 170,
	ExcelConst_xlDialogAppSize = 171,
	ExcelConst_xlDialogMainChartType = 185,
	ExcelConst_xlDialogOverlayChartType = 186,
	ExcelConst_xlDialogOpenMail = 188,
	ExcelConst_xlDialogSendMail = 189,
	ExcelConst_xlDialogStandardFont = 190,
	ExcelConst_xlDialogConsolidate = 191,
	ExcelConst_xlDialogSortSpecial = 192,
	ExcelConst_xlDialogGallery3dArea = 193,
	ExcelConst_xlDialogGallery3dColumn = 194,
	ExcelConst_xlDialogGallery3dLine = 195,
	ExcelConst_xlDialogGallery3dPie = 196,
	ExcelConst_xlDialogView3d = 197,
	ExcelConst_xlDialogGoalSeek = 198,
	ExcelConst_xlDialogWorkgroup = 199,
	ExcelConst_xlDialogFillGroup = 200,
	ExcelConst_xlDialogUpdateLink = 201,
	ExcelConst_xlDialogPromote = 202,
	ExcelConst_xlDialogDemote = 203,
	ExcelConst_xlDialogShowDetail = 204,
	ExcelConst_xlDialogObjectProperties = 207,
	ExcelConst_xlDialogSaveNewObject = 208,
	ExcelConst_xlDialogApplyStyle = 212,
	ExcelConst_xlDialogAssignToObject = 213,
	ExcelConst_xlDialogObjectProtection = 214,
	ExcelConst_xlDialogCreatePublisher = 217,
	ExcelConst_xlDialogSubscribeTo = 218,
	ExcelConst_xlDialogShowToolbar = 220,
	ExcelConst_xlDialogPrintPreview = 222,
	ExcelConst_xlDialogEditColor = 223,
	ExcelConst_xlDialogFormatMain = 225,
	ExcelConst_xlDialogFormatOverlay = 226,
	ExcelConst_xlDialogEditSeries = 228,
	ExcelConst_xlDialogDefineStyle = 229,
	ExcelConst_xlDialogGalleryRadar = 249,
	ExcelConst_xlDialogEditionOptions = 251,
	ExcelConst_xlDialogZoom = 256,
	ExcelConst_xlDialogInsertObject = 259,
	ExcelConst_xlDialogSize = 261,
	ExcelConst_xlDialogMove = 262,
	ExcelConst_xlDialogFormatAuto = 269,
	ExcelConst_xlDialogGallery3dBar = 272,
	ExcelConst_xlDialogGallery3dSurface = 273,
	ExcelConst_xlDialogCustomizeToolbar = 276,
	ExcelConst_xlDialogWorkbookAdd = 281,
	ExcelConst_xlDialogWorkbookMove = 282,
	ExcelConst_xlDialogWorkbookCopy = 283,
	ExcelConst_xlDialogWorkbookOptions = 284,
	ExcelConst_xlDialogSaveWorkspace = 285,
	ExcelConst_xlDialogChartWizard = 288,
	ExcelConst_xlDialogAssignToTool = 293,
	ExcelConst_xlDialogPlacement = 300,
	ExcelConst_xlDialogFillWorkgroup = 301,
	ExcelConst_xlDialogWorkbookNew = 302,
	ExcelConst_xlDialogScenarioCells = 305,
	ExcelConst_xlDialogScenarioAdd = 307,
	ExcelConst_xlDialogScenarioEdit = 308,
	ExcelConst_xlDialogScenarioSummary = 311,
	ExcelConst_xlDialogPivotTableWizard = 312,
	ExcelConst_xlDialogPivotFieldProperties = 313,
	ExcelConst_xlDialogOptionsCalculation = 318,
	ExcelConst_xlDialogOptionsEdit = 319,
	ExcelConst_xlDialogOptionsView = 320,
	ExcelConst_xlDialogAddinManager = 321,
	ExcelConst_xlDialogMenuEditor = 322,
	ExcelConst_xlDialogAttachToolbars = 323,
	ExcelConst_xlDialogOptionsChart = 325,
	ExcelConst_xlDialogVbaInsertFile = 328,
	ExcelConst_xlDialogVbaProcedureDefinition = 330,
	ExcelConst_xlDialogRoutingSlip = 336,
	ExcelConst_xlDialogMailLogon = 339,
	ExcelConst_xlDialogInsertPicture = 342,
	ExcelConst_xlDialogGalleryDoughnut = 344,
	ExcelConst_xlDialogChartTrend = 350,
	ExcelConst_xlDialogWorkbookInsert = 354,
	ExcelConst_xlDialogOptionsTransition = 355,
	ExcelConst_xlDialogOptionsGeneral = 356,
	ExcelConst_xlDialogFilterAdvanced = 370,
	ExcelConst_xlDialogMailNextLetter = 378,
	ExcelConst_xlDialogDataLabel = 379,
	ExcelConst_xlDialogInsertTitle = 380,
	ExcelConst_xlDialogFontProperties = 381,
	ExcelConst_xlDialogMacroOptions = 382,
	ExcelConst_xlDialogWorkbookUnhide = 384,
	ExcelConst_xlDialogWorkbookName = 386,
	ExcelConst_xlDialogGalleryCustom = 388,
	ExcelConst_xlDialogAddChartAutoformat = 390,
	ExcelConst_xlDialogChartAddData = 392,
	ExcelConst_xlDialogTabOrder = 394,
	ExcelConst_xlDialogSubtotalCreate = 398,
	ExcelConst_xlDialogWorkbookTabSplit = 415,
	ExcelConst_xlDialogWorkbookProtect = 417,
	ExcelConst_xlDialogScrollbarProperties = 420,
	ExcelConst_xlDialogPivotShowPages = 421,
	ExcelConst_xlDialogTextToColumns = 422,
	ExcelConst_xlDialogFormatCharttype = 423,
	ExcelConst_xlDialogPivotFieldGroup = 433,
	ExcelConst_xlDialogPivotFieldUngroup = 434,
	ExcelConst_xlDialogCheckboxProperties = 435,
	ExcelConst_xlDialogLabelProperties = 436,
	ExcelConst_xlDialogListboxProperties = 437,
	ExcelConst_xlDialogEditboxProperties = 438,
	ExcelConst_xlDialogOpenText = 441,
	ExcelConst_xlDialogPushbuttonProperties = 445,
	ExcelConst_xlDialogFilter = 447,
	ExcelConst_xlDialogFunctionWizard = 450,
	ExcelConst_xlDialogSaveCopyAs = 456,
	ExcelConst_xlDialogOptionsListsAdd = 458,
	ExcelConst_xlDialogSeriesAxes = 460,
	ExcelConst_xlDialogSeriesX = 461,
	ExcelConst_xlDialogSeriesY = 462,
	ExcelConst_xlDialogErrorbarX = 463,
	ExcelConst_xlDialogErrorbarY = 464,
	ExcelConst_xlDialogFormatChart = 465,
	ExcelConst_xlDialogSeriesOrder = 466,
	ExcelConst_xlDialogMailEditMailer = 470,
	ExcelConst_xlDialogStandardWidth = 472,
	ExcelConst_xlDialogScenarioMerge = 473,
	ExcelConst_xlDialogProperties = 474,
	ExcelConst_xlDialogSummaryInfo = 474,
	ExcelConst_xlDialogFindFile = 475,
	ExcelConst_xlDialogActiveCellFont = 476,
	ExcelConst_xlDialogVbaMakeAddin = 478,
	ExcelConst_xlDialogFileSharing = 481,
	ExcelConst_xlDialogAutoCorrect = 485,
	ExcelConst_xlDialogCustomViews = 493,
	ExcelConst_xlDialogInsertNameLabel = 496,
	ExcelConst_xlDialogSeriesShape = 504,
	ExcelConst_xlDialogChartOptionsDataLabels = 505,
	ExcelConst_xlDialogChartOptionsDataTable = 506,
	ExcelConst_xlDialogSetBackgroundPicture = 509,
	ExcelConst_xlDialogDataValidation = 525,
	ExcelConst_xlDialogChartType = 526,
	ExcelConst_xlDialogChartLocation = 527,
	ExcelConst__xlDialogPhonetic = 538,
	ExcelConst_xlDialogChartSourceData = 540,
	ExcelConst__xlDialogChartSourceData = 541,
	ExcelConst_xlDialogSeriesOptions = 557,
	ExcelConst_xlDialogPivotTableOptions = 567,
	ExcelConst_xlDialogPivotSolveOrder = 568,
	ExcelConst_xlDialogPivotCalculatedField = 570,
	ExcelConst_xlDialogPivotCalculatedItem = 572,
	ExcelConst_xlDialogConditionalFormatting = 583,
	ExcelConst_xlDialogInsertHyperlink = 596,
	ExcelConst_xlDialogProtectSharing = 620,
	ExcelConst_xlDialogOptionsME = 647,
	ExcelConst_xlDialogPublishAsWebPage = 653,
	ExcelConst_xlDialogPhonetic = 656,
	ExcelConst_xlDialogNewWebQuery = 667,
	ExcelConst_xlDialogImportTextFile = 666,
	ExcelConst_xlDialogExternalDataProperties = 530,
	ExcelConst_xlDialogWebOptionsGeneral = 683,
	ExcelConst_xlDialogWebOptionsFiles = 684,
	ExcelConst_xlDialogWebOptionsPictures = 685,
	ExcelConst_xlDialogWebOptionsEncoding = 686,
	ExcelConst_xlDialogWebOptionsFonts = 687,
	ExcelConst_xlDialogPivotClientServerSet = 689,
	_Excel_XlBuiltInDialogForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlSummaryReportType
{
	ExcelConst_xlSummaryPivotTable = -4148,
	ExcelConst_xlStandardSummary = 1,
	_Excel_XlSummaryReportTypeForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlRoutingSlipDelivery
{
	ExcelConst_xlAllAtOnce = 2,
	ExcelConst_xlOneAfterAnother = 1,
	_Excel_XlRoutingSlipDeliveryForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlRoutingSlipStatus
{
	ExcelConst_xlNotYetRouted = 0,
	ExcelConst_xlRoutingComplete = 2,
	ExcelConst_xlRoutingInProgress = 1,
	_Excel_XlRoutingSlipStatusForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlSummaryColumn
{
	ExcelConst_xlSummaryOnLeft = -4131,
	ExcelConst_xlSummaryOnRight = -4152,
	_Excel_XlSummaryColumnForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlObjectSize
{
	ExcelConst_xlFitToPage = 2,
	ExcelConst_xlFullPage = 3,
	ExcelConst_xlScreenSize = 1,
	_Excel_XlObjectSizeForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlOrder
{
	ExcelConst_xlDownThenOver = 1,
	ExcelConst_xlOverThenDown = 2,
	_Excel_XlOrderForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlPageOrientation
{
	ExcelConst_xlLandscape = 2,
	ExcelConst_xlPortrait = 1,
	_Excel_XlPageOrientationForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlPaperSize
{
	ExcelConst_xlPaper10x14 = 16,
	ExcelConst_xlPaper11x17 = 17,
	ExcelConst_xlPaperA3 = 8,
	ExcelConst_xlPaperA4 = 9,
	ExcelConst_xlPaperA4Small = 10,
	ExcelConst_xlPaperA5 = 11,
	ExcelConst_xlPaperB4 = 12,
	ExcelConst_xlPaperB5 = 13,
	ExcelConst_xlPaperCsheet = 24,
	ExcelConst_xlPaperDsheet = 25,
	ExcelConst_xlPaperEnvelope10 = 20,
	ExcelConst_xlPaperEnvelope11 = 21,
	ExcelConst_xlPaperEnvelope12 = 22,
	ExcelConst_xlPaperEnvelope14 = 23,
	ExcelConst_xlPaperEnvelope9 = 19,
	ExcelConst_xlPaperEnvelopeB4 = 33,
	ExcelConst_xlPaperEnvelopeB5 = 34,
	ExcelConst_xlPaperEnvelopeB6 = 35,
	ExcelConst_xlPaperEnvelopeC3 = 29,
	ExcelConst_xlPaperEnvelopeC4 = 30,
	ExcelConst_xlPaperEnvelopeC5 = 28,
	ExcelConst_xlPaperEnvelopeC6 = 31,
	ExcelConst_xlPaperEnvelopeC65 = 32,
	ExcelConst_xlPaperEnvelopeDL = 27,
	ExcelConst_xlPaperEnvelopeItaly = 36,
	ExcelConst_xlPaperEnvelopeMonarch = 37,
	ExcelConst_xlPaperEnvelopePersonal = 38,
	ExcelConst_xlPaperEsheet = 26,
	ExcelConst_xlPaperExecutive = 7,
	ExcelConst_xlPaperFanfoldLegalGerman = 41,
	ExcelConst_xlPaperFanfoldStdGerman = 40,
	ExcelConst_xlPaperFanfoldUS = 39,
	ExcelConst_xlPaperFolio = 14,
	ExcelConst_xlPaperLedger = 4,
	ExcelConst_xlPaperLegal = 5,
	ExcelConst_xlPaperLetter = 1,
	ExcelConst_xlPaperLetterSmall = 2,
	ExcelConst_xlPaperNote = 18,
	ExcelConst_xlPaperQuarto = 15,
	ExcelConst_xlPaperStatement = 6,
	ExcelConst_xlPaperTabloid = 3,
	ExcelConst_xlPaperUser = 256,
	_Excel_XlPaperSizeForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlPrintLocation
{
	ExcelConst_xlPrintSheetEnd = 1,
	ExcelConst_xlPrintInPlace = 16,
	ExcelConst_xlPrintNoComments = -4142,
	_Excel_XlPrintLocationForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlXLMMacroType
{
	ExcelConst_xlCommand = 2,
	ExcelConst_xlFunction = 1,
	ExcelConst_xlNotXLM = 3,
	_Excel_XlXLMMacroTypeForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlFormatConditionType
{
	ExcelConst_xlCellValue = 1,
	ExcelConst_xlExpression = 2,
	_Excel_XlFormatConditionTypeForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlParameterDataType
{
	ExcelConst_xlParamTypeUnknown = 0,
	ExcelConst_xlParamTypeChar = 1,
	ExcelConst_xlParamTypeNumeric = 2,
	ExcelConst_xlParamTypeDecimal = 3,
	ExcelConst_xlParamTypeInteger = 4,
	ExcelConst_xlParamTypeSmallInt = 5,
	ExcelConst_xlParamTypeFloat = 6,
	ExcelConst_xlParamTypeReal = 7,
	ExcelConst_xlParamTypeDouble = 8,
	ExcelConst_xlParamTypeVarChar = 12,
	ExcelConst_xlParamTypeDate = 9,
	ExcelConst_xlParamTypeTime = 10,
	ExcelConst_xlParamTypeTimestamp = 11,
	ExcelConst_xlParamTypeLongVarChar = -1,
	ExcelConst_xlParamTypeBinary = -2,
	ExcelConst_xlParamTypeVarBinary = -3,
	ExcelConst_xlParamTypeLongVarBinary = -4,
	ExcelConst_xlParamTypeBigInt = -5,
	ExcelConst_xlParamTypeTinyInt = -6,
	ExcelConst_xlParamTypeBit = -7,
	ExcelConst_xlParamTypeWChar = -8,
	_Excel_XlParameterDataTypeForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlParameterType
{
	ExcelConst_xlPrompt = 0,
	ExcelConst_xlConstant = 1,
	ExcelConst_xlRange = 2,
	_Excel_XlParameterTypeForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlDVType
{
	ExcelConst_xlValidateInputOnly = 0,
	ExcelConst_xlValidateWholeNumber = 1,
	ExcelConst_xlValidateDecimal = 2,
	ExcelConst_xlValidateList = 3,
	ExcelConst_xlValidateDate = 4,
	ExcelConst_xlValidateTime = 5,
	ExcelConst_xlValidateTextLength = 6,
	ExcelConst_xlValidateCustom = 7,
	_Excel_XlDVTypeForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlAxisCrosses
{
	ExcelConst_xlAxisCrossesAutomatic = -4105,
	ExcelConst_xlAxisCrossesCustom = -4114,
	ExcelConst_xlAxisCrossesMaximum = 2,
	ExcelConst_xlAxisCrossesMinimum = 4,
	_Excel_XlAxisCrossesForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlTickMark
{
	ExcelConst_xlTickMarkCross = 4,
	ExcelConst_xlTickMarkInside = 2,
	ExcelConst_xlTickMarkNone = -4142,
	ExcelConst_xlTickMarkOutside = 3,
	_Excel_XlTickMarkForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlScaleType
{
	ExcelConst_xlScaleLinear = -4132,
	ExcelConst_xlScaleLogarithmic = -4133,
	_Excel_XlScaleTypeForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlTickLabelPosition
{
	ExcelConst_xlTickLabelPositionHigh = -4127,
	ExcelConst_xlTickLabelPositionLow = -4134,
	ExcelConst_xlTickLabelPositionNextToAxis = 4,
	ExcelConst_xlTickLabelPositionNone = -4142,
	_Excel_XlTickLabelPositionForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlAxisType
{
	ExcelConst_xlCategory = 1,
	ExcelConst_xlSeriesAxis = 3,
	ExcelConst_xlValue = 2,
	_Excel_XlAxisTypeForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlTimeUnit
{
	ExcelConst_xlDays = 0,
	ExcelConst_xlMonths = 1,
	ExcelConst_xlYears = 2,
	_Excel_XlTimeUnitForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlCategoryType
{
	ExcelConst_xlCategoryScale = 2,
	ExcelConst_xlTimeScale = 3,
	ExcelConst_xlAutomaticScale = -4105,
	_Excel_XlCategoryTypeForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlDisplayUnit
{
	ExcelConst_xlHundreds = -2,
	ExcelConst_xlThousands = -3,
	ExcelConst_xlTenThousands = -4,
	ExcelConst_xlHundredThousands = -5,
	ExcelConst_xlMillions = -6,
	ExcelConst_xlTenMillions = -7,
	ExcelConst_xlHundredMillions = -8,
	ExcelConst_xlThousandMillions = -9,
	ExcelConst_xlMillionMillions = -10,
	_Excel_XlDisplayUnitForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlSizeRepresents
{
	ExcelConst_xlSizeIsWidth = 2,
	ExcelConst_xlSizeIsArea = 1,
	_Excel_XlSizeRepresentsForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlChartSplitType
{
	ExcelConst_xlSplitByPosition = 1,
	ExcelConst_xlSplitByPercentValue = 3,
	ExcelConst_xlSplitByCustomSplit = 4,
	ExcelConst_xlSplitByValue = 2,
	_Excel_XlChartSplitTypeForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlMarkerStyle
{
	ExcelConst_xlMarkerStyleAutomatic = -4105,
	ExcelConst_xlMarkerStyleCircle = 8,
	ExcelConst_xlMarkerStyleDash = -4115,
	ExcelConst_xlMarkerStyleDiamond = 2,
	ExcelConst_xlMarkerStyleDot = -4118,
	ExcelConst_xlMarkerStyleNone = -4142,
	ExcelConst_xlMarkerStylePicture = -4147,
	ExcelConst_xlMarkerStylePlus = 9,
	ExcelConst_xlMarkerStyleSquare = 1,
	ExcelConst_xlMarkerStyleStar = 5,
	ExcelConst_xlMarkerStyleTriangle = 3,
	ExcelConst_xlMarkerStyleX = -4168,
	_Excel_XlMarkerStyleForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlChartPictureType
{
	ExcelConst_xlStackScale = 3,
	ExcelConst_xlStack = 2,
	ExcelConst_xlStretch = 1,
	_Excel_XlChartPictureTypeForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlErrorBarDirection
{
	ExcelConst_xlX = -4168,
	ExcelConst_xlY = 1,
	_Excel_XlErrorBarDirectionForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlErrorBarInclude
{
	ExcelConst_xlErrorBarIncludeBoth = 1,
	ExcelConst_xlErrorBarIncludeMinusValues = 3,
	ExcelConst_xlErrorBarIncludeNone = -4142,
	ExcelConst_xlErrorBarIncludePlusValues = 2,
	_Excel_XlErrorBarIncludeForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlErrorBarType
{
	ExcelConst_xlErrorBarTypeCustom = -4114,
	ExcelConst_xlErrorBarTypeFixedValue = 1,
	ExcelConst_xlErrorBarTypePercent = 2,
	ExcelConst_xlErrorBarTypeStDev = -4155,
	ExcelConst_xlErrorBarTypeStError = 4,
	_Excel_XlErrorBarTypeForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlDataLabelPosition
{
	ExcelConst_xlLabelPositionCenter = -4108,
	ExcelConst_xlLabelPositionAbove = 0,
	ExcelConst_xlLabelPositionBelow = 1,
	ExcelConst_xlLabelPositionLeft = -4131,
	ExcelConst_xlLabelPositionRight = -4152,
	ExcelConst_xlLabelPositionOutsideEnd = 2,
	ExcelConst_xlLabelPositionInsideEnd = 3,
	ExcelConst_xlLabelPositionInsideBase = 4,
	ExcelConst_xlLabelPositionBestFit = 5,
	ExcelConst_xlLabelPositionMixed = 6,
	ExcelConst_xlLabelPositionCustom = 7,
	_Excel_XlDataLabelPositionForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlTrendlineType
{
	ExcelConst_xlExponential = 5,
	ExcelConst_xlLinear = -4132,
	ExcelConst_xlLogarithmic = -4133,
	ExcelConst_xlMovingAvg = 6,
	ExcelConst_xlPolynomial = 3,
	ExcelConst_xlPower = 4,
	_Excel_XlTrendlineTypeForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlTickLabelOrientation
{
	ExcelConst_xlTickLabelOrientationAutomatic = -4105,
	ExcelConst_xlTickLabelOrientationDownward = -4170,
	ExcelConst_xlTickLabelOrientationHorizontal = -4128,
	ExcelConst_xlTickLabelOrientationUpward = -4171,
	ExcelConst_xlTickLabelOrientationVertical = -4166,
	_Excel_XlTickLabelOrientationForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlLegendPosition
{
	ExcelConst_xlLegendPositionBottom = -4107,
	ExcelConst_xlLegendPositionCorner = 2,
	ExcelConst_xlLegendPositionLeft = -4131,
	ExcelConst_xlLegendPositionRight = -4152,
	ExcelConst_xlLegendPositionTop = -4160,
	_Excel_XlLegendPositionForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlEndStyleCap
{
	ExcelConst_xlCap = 1,
	ExcelConst_xlNoCap = 2,
	_Excel_XlEndStyleCapForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_MsoFlipCmd
{
	ExcelConst_msoFlipHorizontal = 0,
	ExcelConst_msoFlipVertical = 1,
	_Excel_MsoFlipCmdForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_MsoZOrderCmd
{
	ExcelConst_msoBringToFront = 0,
	ExcelConst_msoSendToBack = 1,
	ExcelConst_msoBringForward = 2,
	ExcelConst_msoSendBackward = 3,
	ExcelConst_msoBringInFrontOfText = 4,
	ExcelConst_msoSendBehindText = 5,
	_Excel_MsoZOrderCmdForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_MsoAutoShapeType
{
	ExcelConst_msoShapeMixed = -2,
	ExcelConst_msoShapeRectangle = 1,
	ExcelConst_msoShapeParallelogram = 2,
	ExcelConst_msoShapeTrapezoid = 3,
	ExcelConst_msoShapeDiamond = 4,
	ExcelConst_msoShapeRoundedRectangle = 5,
	ExcelConst_msoShapeOctagon = 6,
	ExcelConst_msoShapeIsoscelesTriangle = 7,
	ExcelConst_msoShapeRightTriangle = 8,
	ExcelConst_msoShapeOval = 9,
	ExcelConst_msoShapeHexagon = 10,
	ExcelConst_msoShapeCross = 11,
	ExcelConst_msoShapeRegularPentagon = 12,
	ExcelConst_msoShapeCan = 13,
	ExcelConst_msoShapeCube = 14,
	ExcelConst_msoShapeBevel = 15,
	ExcelConst_msoShapeFoldedCorner = 16,
	ExcelConst_msoShapeSmileyFace = 17,
	ExcelConst_msoShapeDonut = 18,
	ExcelConst_msoShapeNoSymbol = 19,
	ExcelConst_msoShapeBlockArc = 20,
	ExcelConst_msoShapeHeart = 21,
	ExcelConst_msoShapeLightningBolt = 22,
	ExcelConst_msoShapeSun = 23,
	ExcelConst_msoShapeMoon = 24,
	ExcelConst_msoShapeArc = 25,
	ExcelConst_msoShapeDoubleBracket = 26,
	ExcelConst_msoShapeDoubleBrace = 27,
	ExcelConst_msoShapePlaque = 28,
	ExcelConst_msoShapeLeftBracket = 29,
	ExcelConst_msoShapeRightBracket = 30,
	ExcelConst_msoShapeLeftBrace = 31,
	ExcelConst_msoShapeRightBrace = 32,
	ExcelConst_msoShapeRightArrow = 33,
	ExcelConst_msoShapeLeftArrow = 34,
	ExcelConst_msoShapeUpArrow = 35,
	ExcelConst_msoShapeDownArrow = 36,
	ExcelConst_msoShapeLeftRightArrow = 37,
	ExcelConst_msoShapeUpDownArrow = 38,
	ExcelConst_msoShapeQuadArrow = 39,
	ExcelConst_msoShapeLeftRightUpArrow = 40,
	ExcelConst_msoShapeBentArrow = 41,
	ExcelConst_msoShapeUTurnArrow = 42,
	ExcelConst_msoShapeLeftUpArrow = 43,
	ExcelConst_msoShapeBentUpArrow = 44,
	ExcelConst_msoShapeCurvedRightArrow = 45,
	ExcelConst_msoShapeCurvedLeftArrow = 46,
	ExcelConst_msoShapeCurvedUpArrow = 47,
	ExcelConst_msoShapeCurvedDownArrow = 48,
	ExcelConst_msoShapeStripedRightArrow = 49,
	ExcelConst_msoShapeNotchedRightArrow = 50,
	ExcelConst_msoShapePentagon = 51,
	ExcelConst_msoShapeChevron = 52,
	ExcelConst_msoShapeRightArrowCallout = 53,
	ExcelConst_msoShapeLeftArrowCallout = 54,
	ExcelConst_msoShapeUpArrowCallout = 55,
	ExcelConst_msoShapeDownArrowCallout = 56,
	ExcelConst_msoShapeLeftRightArrowCallout = 57,
	ExcelConst_msoShapeUpDownArrowCallout = 58,
	ExcelConst_msoShapeQuadArrowCallout = 59,
	ExcelConst_msoShapeCircularArrow = 60,
	ExcelConst_msoShapeFlowchartProcess = 61,
	ExcelConst_msoShapeFlowchartAlternateProcess = 62,
	ExcelConst_msoShapeFlowchartDecision = 63,
	ExcelConst_msoShapeFlowchartData = 64,
	ExcelConst_msoShapeFlowchartPredefinedProcess = 65,
	ExcelConst_msoShapeFlowchartInternalStorage = 66,
	ExcelConst_msoShapeFlowchartDocument = 67,
	ExcelConst_msoShapeFlowchartMultidocument = 68,
	ExcelConst_msoShapeFlowchartTerminator = 69,
	ExcelConst_msoShapeFlowchartPreparation = 70,
	ExcelConst_msoShapeFlowchartManualInput = 71,
	ExcelConst_msoShapeFlowchartManualOperation = 72,
	ExcelConst_msoShapeFlowchartConnector = 73,
	ExcelConst_msoShapeFlowchartOffpageConnector = 74,
	ExcelConst_msoShapeFlowchartCard = 75,
	ExcelConst_msoShapeFlowchartPunchedTape = 76,
	ExcelConst_msoShapeFlowchartSummingJunction = 77,
	ExcelConst_msoShapeFlowchartOr = 78,
	ExcelConst_msoShapeFlowchartCollate = 79,
	ExcelConst_msoShapeFlowchartSort = 80,
	ExcelConst_msoShapeFlowchartExtract = 81,
	ExcelConst_msoShapeFlowchartMerge = 82,
	ExcelConst_msoShapeFlowchartStoredData = 83,
	ExcelConst_msoShapeFlowchartDelay = 84,
	ExcelConst_msoShapeFlowchartSequentialAccessStorage = 85,
	ExcelConst_msoShapeFlowchartMagneticDisk = 86,
	ExcelConst_msoShapeFlowchartDirectAccessStorage = 87,
	ExcelConst_msoShapeFlowchartDisplay = 88,
	ExcelConst_msoShapeExplosion1 = 89,
	ExcelConst_msoShapeExplosion2 = 90,
	ExcelConst_msoShape4pointStar = 91,
	ExcelConst_msoShape5pointStar = 92,
	ExcelConst_msoShape8pointStar = 93,
	ExcelConst_msoShape16pointStar = 94,
	ExcelConst_msoShape24pointStar = 95,
	ExcelConst_msoShape32pointStar = 96,
	ExcelConst_msoShapeUpRibbon = 97,
	ExcelConst_msoShapeDownRibbon = 98,
	ExcelConst_msoShapeCurvedUpRibbon = 99,
	ExcelConst_msoShapeCurvedDownRibbon = 100,
	ExcelConst_msoShapeVerticalScroll = 101,
	ExcelConst_msoShapeHorizontalScroll = 102,
	ExcelConst_msoShapeWave = 103,
	ExcelConst_msoShapeDoubleWave = 104,
	ExcelConst_msoShapeRectangularCallout = 105,
	ExcelConst_msoShapeRoundedRectangularCallout = 106,
	ExcelConst_msoShapeOvalCallout = 107,
	ExcelConst_msoShapeCloudCallout = 108,
	ExcelConst_msoShapeLineCallout1 = 109,
	ExcelConst_msoShapeLineCallout2 = 110,
	ExcelConst_msoShapeLineCallout3 = 111,
	ExcelConst_msoShapeLineCallout4 = 112,
	ExcelConst_msoShapeLineCallout1AccentBar = 113,
	ExcelConst_msoShapeLineCallout2AccentBar = 114,
	ExcelConst_msoShapeLineCallout3AccentBar = 115,
	ExcelConst_msoShapeLineCallout4AccentBar = 116,
	ExcelConst_msoShapeLineCallout1NoBorder = 117,
	ExcelConst_msoShapeLineCallout2NoBorder = 118,
	ExcelConst_msoShapeLineCallout3NoBorder = 119,
	ExcelConst_msoShapeLineCallout4NoBorder = 120,
	ExcelConst_msoShapeLineCallout1BorderandAccentBar = 121,
	ExcelConst_msoShapeLineCallout2BorderandAccentBar = 122,
	ExcelConst_msoShapeLineCallout3BorderandAccentBar = 123,
	ExcelConst_msoShapeLineCallout4BorderandAccentBar = 124,
	ExcelConst_msoShapeActionButtonCustom = 125,
	ExcelConst_msoShapeActionButtonHome = 126,
	ExcelConst_msoShapeActionButtonHelp = 127,
	ExcelConst_msoShapeActionButtonInformation = 128,
	ExcelConst_msoShapeActionButtonBackorPrevious = 129,
	ExcelConst_msoShapeActionButtonForwardorNext = 130,
	ExcelConst_msoShapeActionButtonBeginning = 131,
	ExcelConst_msoShapeActionButtonEnd = 132,
	ExcelConst_msoShapeActionButtonReturn = 133,
	ExcelConst_msoShapeActionButtonDocument = 134,
	ExcelConst_msoShapeActionButtonSound = 135,
	ExcelConst_msoShapeActionButtonMovie = 136,
	ExcelConst_msoShapeBalloon = 137,
	ExcelConst_msoShapeNotPrimitive = 138,
	_Excel_MsoAutoShapeTypeForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_MsoShapeType
{
	ExcelConst_msoShapeTypeMixed = -2,
	ExcelConst_msoAutoShape = 1,
	ExcelConst_msoCallout = 2,
	ExcelConst_msoChart = 3,
	ExcelConst_msoComment = 4,
	ExcelConst_msoFreeform = 5,
	ExcelConst_msoGroup = 6,
	ExcelConst_msoEmbeddedOLEObject = 7,
	ExcelConst_msoFormControl = 8,
	ExcelConst_msoLine = 9,
	ExcelConst_msoLinkedOLEObject = 10,
	ExcelConst_msoLinkedPicture = 11,
	ExcelConst_msoOLEControlObject = 12,
	ExcelConst_msoPicture = 13,
	ExcelConst_msoPlaceholder = 14,
	ExcelConst_msoTextEffect = 15,
	ExcelConst_msoMedia = 16,
	ExcelConst_msoTextBox = 17,
	ExcelConst_msoScriptAnchor = 18,
	ExcelConst_msoTable = 19,
	_Excel_MsoShapeTypeForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_MsoBlackWhiteMode
{
	ExcelConst_msoBlackWhiteMixed = -2,
	ExcelConst_msoBlackWhiteAutomatic = 1,
	ExcelConst_msoBlackWhiteGrayScale = 2,
	ExcelConst_msoBlackWhiteLightGrayScale = 3,
	ExcelConst_msoBlackWhiteInverseGrayScale = 4,
	ExcelConst_msoBlackWhiteGrayOutline = 5,
	ExcelConst_msoBlackWhiteBlackTextAndLine = 6,
	ExcelConst_msoBlackWhiteHighContrast = 7,
	ExcelConst_msoBlackWhiteBlack = 8,
	ExcelConst_msoBlackWhiteWhite = 9,
	ExcelConst_msoBlackWhiteDontShow = 10,
	_Excel_MsoBlackWhiteModeForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlPlacement
{
	ExcelConst_xlFreeFloating = 3,
	ExcelConst_xlMove = 2,
	ExcelConst_xlMoveAndSize = 1,
	_Excel_XlPlacementForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlFormControl
{
	ExcelConst_xlButtonControl = 0,
	ExcelConst_xlCheckBox = 1,
	ExcelConst_xlDropDown = 2,
	ExcelConst_xlEditBox = 3,
	ExcelConst_xlGroupBox = 4,
	ExcelConst_xlLabel = 5,
	ExcelConst_xlListBox = 6,
	ExcelConst_xlOptionButton = 7,
	ExcelConst_xlScrollBar = 8,
	ExcelConst_xlSpinner = 9,
	_Excel_XlFormControlForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_MsoConnectorType
{
	ExcelConst_msoConnectorTypeMixed = -2,
	ExcelConst_msoConnectorStraight = 1,
	ExcelConst_msoConnectorElbow = 2,
	ExcelConst_msoConnectorCurve = 3,
	_Excel_MsoConnectorTypeForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_MsoTextOrientation
{
	ExcelConst_msoTextOrientationMixed = -2,
	ExcelConst_msoTextOrientationHorizontal = 1,
	ExcelConst_msoTextOrientationUpward = 2,
	ExcelConst_msoTextOrientationDownward = 3,
	ExcelConst_msoTextOrientationVerticalFarEast = 4,
	ExcelConst_msoTextOrientationVertical = 5,
	ExcelConst_msoTextOrientationHorizontalRotatedFarEast = 6,
	_Excel_MsoTextOrientationForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_MsoAlignCmd
{
	ExcelConst_msoAlignLefts = 0,
	ExcelConst_msoAlignCenters = 1,
	ExcelConst_msoAlignRights = 2,
	ExcelConst_msoAlignTops = 3,
	ExcelConst_msoAlignMiddles = 4,
	ExcelConst_msoAlignBottoms = 5,
	_Excel_MsoAlignCmdForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_MsoDistributeCmd
{
	ExcelConst_msoDistributeHorizontally = 0,
	ExcelConst_msoDistributeVertically = 1,
	_Excel_MsoDistributeCmdForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_Constants
{
	ExcelConst_xlAll = -4104,
	ExcelConst_xlAutomatic = -4105,
	ExcelConst_xlBoth = 1,
	ExcelConst_xlCenter = -4108,
	ExcelConst_xlChecker = 9,
	ExcelConst_xlCircle = 8,
	ExcelConst_xlCorner = 2,
	ExcelConst_xlCrissCross = 16,
	ExcelConst_xlCross = 4,
	ExcelConst_xlDiamond = 2,
	ExcelConst_xlDistributed = -4117,
	ExcelConst_xlDoubleAccounting = 5,
	ExcelConst_xlFixedValue = 1,
	ExcelConst_xlFormats = -4122,
	ExcelConst_xlGray16 = 17,
	ExcelConst_xlGray8 = 18,
	ExcelConst_xlGrid = 15,
	ExcelConst_xlHigh = -4127,
	ExcelConst_xlInside = 2,
	ExcelConst_xlJustify = -4130,
	ExcelConst_xlLightDown = 13,
	ExcelConst_xlLightHorizontal = 11,
	ExcelConst_xlLightUp = 14,
	ExcelConst_xlLightVertical = 12,
	ExcelConst_xlLow = -4134,
	ExcelConst_xlManual = -4135,
	ExcelConst_xlMinusValues = 3,
	ExcelConst_xlModule = -4141,
	ExcelConst_xlNextToAxis = 4,
	ExcelConst_xlNone = -4142,
	ExcelConst_xlNotes = -4144,
	ExcelConst_xlOff = -4146,
	ExcelConst_xlOn = 1,
	ExcelConst_xlPercent = 2,
	ExcelConst_xlPlus = 9,
	ExcelConst_xlPlusValues = 2,
	ExcelConst_xlSemiGray75 = 10,
	ExcelConst_xlShowLabel = 4,
	ExcelConst_xlShowLabelAndPercent = 5,
	ExcelConst_xlShowPercent = 3,
	ExcelConst_xlShowValue = 2,
	ExcelConst_xlSimple = -4154,
	ExcelConst_xlSingle = 2,
	ExcelConst_xlSingleAccounting = 4,
	ExcelConst_xlSolid = 1,
	ExcelConst_xlSquare = 1,
	ExcelConst_xlStar = 5,
	ExcelConst_xlStError = 4,
	ExcelConst_xlToolbarButton = 2,
	ExcelConst_xlTriangle = 3,
	ExcelConst_xlGray25 = -4124,
	ExcelConst_xlGray50 = -4125,
	ExcelConst_xlGray75 = -4126,
	ExcelConst_xlBottom = -4107,
	ExcelConst_xlLeft = -4131,
	ExcelConst_xlRight = -4152,
	ExcelConst_xlTop = -4160,
	ExcelConst_xl3DBar = -4099,
	ExcelConst_xl3DSurface = -4103,
	ExcelConst_xlBar = 2,
	ExcelConst_xlColumn = 3,
	ExcelConst_xlCombination = -4111,
	ExcelConst_xlCustom = -4114,
	ExcelConst_xlDefaultAutoFormat = -1,
	ExcelConst_xlMaximum = 2,
	ExcelConst_xlMinimum = 4,
	ExcelConst_xlOpaque = 3,
	ExcelConst_xlTransparent = 2,
	ExcelConst_xlBidi = -5000,
	ExcelConst_xlLatin = -5001,
	ExcelConst_xlContext = -5002,
	ExcelConst_xlLTR = -5003,
	ExcelConst_xlRTL = -5004,
	ExcelConst_xlFullScript = 1,
	ExcelConst_xlPartialScript = 2,
	ExcelConst_xlMixedScript = 3,
	ExcelConst_xlMixedAuthorizedScript = 4,
	ExcelConst_xlVisualCursor = 2,
	ExcelConst_xlLogicalCursor = 1,
	ExcelConst_xlSystem = 1,
	ExcelConst_xlPartial = 3,
	ExcelConst_xlHindiNumerals = 3,
	ExcelConst_xlBidiCalendar = 3,
	ExcelConst_xlGregorian = 2,
	ExcelConst_xlComplete = 4,
	ExcelConst_xlScale = 3,
	ExcelConst_xlClosed = 3,
	ExcelConst_xlColor1 = 7,
	ExcelConst_xlColor2 = 8,
	ExcelConst_xlColor3 = 9,
	ExcelConst_xlConstants = 2,
	ExcelConst_xlContents = 2,
	ExcelConst_xlBelow = 1,
	ExcelConst_xlCascade = 7,
	ExcelConst_xlCenterAcrossSelection = 7,
	ExcelConst_xlChart4 = 2,
	ExcelConst_xlChartSeries = 17,
	ExcelConst_xlChartShort = 6,
	ExcelConst_xlChartTitles = 18,
	ExcelConst_xlClassic1 = 1,
	ExcelConst_xlClassic2 = 2,
	ExcelConst_xlClassic3 = 3,
	ExcelConst_xl3DEffects1 = 13,
	ExcelConst_xl3DEffects2 = 14,
	ExcelConst_xlAbove = 0,
	ExcelConst_xlAccounting1 = 4,
	ExcelConst_xlAccounting2 = 5,
	ExcelConst_xlAccounting3 = 6,
	ExcelConst_xlAccounting4 = 17,
	ExcelConst_xlAdd = 2,
	ExcelConst_xlDebugCodePane = 13,
	ExcelConst_xlDesktop = 9,
	ExcelConst_xlDirect = 1,
	ExcelConst_xlDivide = 5,
	ExcelConst_xlDoubleClosed = 5,
	ExcelConst_xlDoubleOpen = 4,
	ExcelConst_xlDoubleQuote = 1,
	ExcelConst_xlEntireChart = 20,
	ExcelConst_xlExcelMenus = 1,
	ExcelConst_xlExtended = 3,
	ExcelConst_xlFill = 5,
	ExcelConst_xlFirst = 0,
	ExcelConst_xlFloating = 5,
	ExcelConst_xlFormula = 5,
	ExcelConst_xlGeneral = 1,
	ExcelConst_xlGridline = 22,
	ExcelConst_xlIcons = 1,
	ExcelConst_xlImmediatePane = 12,
	ExcelConst_xlInteger = 2,
	ExcelConst_xlLast = 1,
	ExcelConst_xlLastCell = 11,
	ExcelConst_xlList1 = 10,
	ExcelConst_xlList2 = 11,
	ExcelConst_xlList3 = 12,
	ExcelConst_xlLocalFormat1 = 15,
	ExcelConst_xlLocalFormat2 = 16,
	ExcelConst_xlLong = 3,
	ExcelConst_xlLotusHelp = 2,
	ExcelConst_xlMacrosheetCell = 7,
	ExcelConst_xlMixed = 2,
	ExcelConst_xlMultiply = 4,
	ExcelConst_xlNarrow = 1,
	ExcelConst_xlNoDocuments = 3,
	ExcelConst_xlOpen = 2,
	ExcelConst_xlOutside = 3,
	ExcelConst_xlReference = 4,
	ExcelConst_xlSemiautomatic = 2,
	ExcelConst_xlShort = 1,
	ExcelConst_xlSingleQuote = 2,
	ExcelConst_xlStrict = 2,
	ExcelConst_xlSubtract = 3,
	ExcelConst_xlTextBox = 16,
	ExcelConst_xlTiled = 1,
	ExcelConst_xlTitleBar = 8,
	ExcelConst_xlToolbar = 1,
	ExcelConst_xlVisible = 12,
	ExcelConst_xlWatchPane = 11,
	ExcelConst_xlWide = 3,
	ExcelConst_xlWorkbookTab = 6,
	ExcelConst_xlWorksheet4 = 1,
	ExcelConst_xlWorksheetCell = 3,
	ExcelConst_xlWorksheetShort = 5,
	ExcelConst_xlAllExceptBorders = 6,
	ExcelConst_xlLeftToRight = 2,
	ExcelConst_xlTopToBottom = 1,
	ExcelConst_xlVeryHidden = 2,
	ExcelConst_xlDrawingObject = 14,
	_Excel_ConstantsForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlChartGallery
{
	ExcelConst_xlBuiltIn = 21,
	ExcelConst_xlUserDefined = 22,
	ExcelConst_xlAnyGallery = 23,
	_Excel_XlChartGalleryForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlBackground
{
	ExcelConst_xlBackgroundAutomatic = -4105,
	ExcelConst_xlBackgroundOpaque = 3,
	ExcelConst_xlBackgroundTransparent = 2,
	_Excel_XlBackgroundForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlArrowHeadLength
{
	ExcelConst_xlArrowHeadLengthLong = 3,
	ExcelConst_xlArrowHeadLengthMedium = -4138,
	ExcelConst_xlArrowHeadLengthShort = 1,
	_Excel_XlArrowHeadLengthForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlArrowHeadStyle
{
	ExcelConst_xlArrowHeadStyleClosed = 3,
	ExcelConst_xlArrowHeadStyleDoubleClosed = 5,
	ExcelConst_xlArrowHeadStyleDoubleOpen = 4,
	ExcelConst_xlArrowHeadStyleNone = -4142,
	ExcelConst_xlArrowHeadStyleOpen = 2,
	_Excel_XlArrowHeadStyleForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlArrowHeadWidth
{
	ExcelConst_xlArrowHeadWidthMedium = -4138,
	ExcelConst_xlArrowHeadWidthNarrow = 1,
	ExcelConst_xlArrowHeadWidthWide = 3,
	_Excel_XlArrowHeadWidthForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlChartPicturePlacement
{
	ExcelConst_xlSides = 1,
	ExcelConst_xlEnd = 2,
	ExcelConst_xlEndSides = 3,
	ExcelConst_xlFront = 4,
	ExcelConst_xlFrontSides = 5,
	ExcelConst_xlFrontEnd = 6,
	ExcelConst_xlAllFaces = 7,
	_Excel_XlChartPicturePlacementForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlUnderlineStyle
{
	ExcelConst_xlUnderlineStyleDouble = -4119,
	ExcelConst_xlUnderlineStyleDoubleAccounting = 5,
	ExcelConst_xlUnderlineStyleNone = -4142,
	ExcelConst_xlUnderlineStyleSingle = 2,
	ExcelConst_xlUnderlineStyleSingleAccounting = 4,
	_Excel_XlUnderlineStyleForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlLineStyle
{
	ExcelConst_xlContinuous = 1,
	ExcelConst_xlDash = -4115,
	ExcelConst_xlDashDot = 4,
	ExcelConst_xlDashDotDot = 5,
	ExcelConst_xlDot = -4118,
	ExcelConst_xlDouble = -4119,
	ExcelConst_xlSlantDashDot = 13,
	ExcelConst_xlLineStyleNone = -4142,
	_Excel_XlLineStyleForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlPictureConvertorType
{
	ExcelConst_xlBMP = 1,
	ExcelConst_xlCGM = 7,
	ExcelConst_xlDRW = 4,
	ExcelConst_xlDXF = 5,
	ExcelConst_xlEPS = 8,
	ExcelConst_xlHGL = 6,
	ExcelConst_xlPCT = 13,
	ExcelConst_xlPCX = 10,
	ExcelConst_xlPIC = 11,
	ExcelConst_xlPLT = 12,
	ExcelConst_xlTIF = 9,
	ExcelConst_xlWMF = 2,
	ExcelConst_xlWPG = 3,
	_Excel_XlPictureConvertorTypeForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlPattern
{
	ExcelConst_xlPatternAutomatic = -4105,
	ExcelConst_xlPatternChecker = 9,
	ExcelConst_xlPatternCrissCross = 16,
	ExcelConst_xlPatternDown = -4121,
	ExcelConst_xlPatternGray16 = 17,
	ExcelConst_xlPatternGray25 = -4124,
	ExcelConst_xlPatternGray50 = -4125,
	ExcelConst_xlPatternGray75 = -4126,
	ExcelConst_xlPatternGray8 = 18,
	ExcelConst_xlPatternGrid = 15,
	ExcelConst_xlPatternHorizontal = -4128,
	ExcelConst_xlPatternLightDown = 13,
	ExcelConst_xlPatternLightHorizontal = 11,
	ExcelConst_xlPatternLightUp = 14,
	ExcelConst_xlPatternLightVertical = 12,
	ExcelConst_xlPatternNone = -4142,
	ExcelConst_xlPatternSemiGray75 = 10,
	ExcelConst_xlPatternSolid = 1,
	ExcelConst_xlPatternUp = -4162,
	ExcelConst_xlPatternVertical = -4166,
	_Excel_XlPatternForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlChartItem
{
	ExcelConst_xlDataLabel = 0,
	ExcelConst_xlChartArea = 2,
	ExcelConst_xlSeries = 3,
	ExcelConst_xlChartTitle = 4,
	ExcelConst_xlWalls = 5,
	ExcelConst_xlCorners = 6,
	ExcelConst_xlDataTable = 7,
	ExcelConst_xlTrendline = 8,
	ExcelConst_xlErrorBars = 9,
	ExcelConst_xlXErrorBars = 10,
	ExcelConst_xlYErrorBars = 11,
	ExcelConst_xlLegendEntry = 12,
	ExcelConst_xlLegendKey = 13,
	ExcelConst_xlShape = 14,
	ExcelConst_xlMajorGridlines = 15,
	ExcelConst_xlMinorGridlines = 16,
	ExcelConst_xlAxisTitle = 17,
	ExcelConst_xlUpBars = 18,
	ExcelConst_xlPlotArea = 19,
	ExcelConst_xlDownBars = 20,
	ExcelConst_xlAxis = 21,
	ExcelConst_xlSeriesLines = 22,
	ExcelConst_xlFloor = 23,
	ExcelConst_xlLegend = 24,
	ExcelConst_xlHiLoLines = 25,
	ExcelConst_xlDropLines = 26,
	ExcelConst_xlRadarAxisLabels = 27,
	ExcelConst_xlNothing = 28,
	ExcelConst_xlLeaderLines = 29,
	ExcelConst_xlDisplayUnitLabel = 30,
	ExcelConst_xlPivotChartFieldButton = 31,
	ExcelConst_xlPivotChartDropZone = 32,
	_Excel_XlChartItemForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlInsertShiftDirection
{
	ExcelConst_xlShiftDown = -4121,
	ExcelConst_xlShiftToRight = -4161,
	_Excel_XlInsertShiftDirectionForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlDeleteShiftDirection
{
	ExcelConst_xlShiftToLeft = -4159,
	ExcelConst_xlShiftUp = -4162,
	_Excel_XlDeleteShiftDirectionForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlFindLookIn
{
	ExcelConst_xlFormulas = -4123,
	ExcelConst_xlComments = -4144,
	ExcelConst_xlValues = -4163,
	_Excel_XlFindLookInForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlMouseButton
{
	ExcelConst_xlNoButton = 0,
	ExcelConst_xlPrimaryButton = 1,
	ExcelConst_xlSecondaryButton = 2,
	_Excel_XlMouseButtonForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlOLEType
{
	ExcelConst_xlOLEControl = 2,
	ExcelConst_xlOLEEmbed = 1,
	ExcelConst_xlOLELink = 0,
	_Excel_XlOLETypeForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlLookAt
{
	ExcelConst_xlPart = 2,
	ExcelConst_xlWhole = 1,
	_Excel_XlLookAtForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlLinkInfoType
{
	ExcelConst_xlLinkInfoOLELinks = 2,
	ExcelConst_xlLinkInfoPublishers = 5,
	ExcelConst_xlLinkInfoSubscribers = 6,
	_Excel_XlLinkInfoTypeForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlCVError
{
	ExcelConst_xlErrDiv0 = 2007,
	ExcelConst_xlErrNA = 2042,
	ExcelConst_xlErrName = 2029,
	ExcelConst_xlErrNull = 2000,
	ExcelConst_xlErrNum = 2036,
	ExcelConst_xlErrRef = 2023,
	ExcelConst_xlErrValue = 2015,
	_Excel_XlCVErrorForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlEditionFormat
{
	ExcelConst_xlBIFF = 2,
	ExcelConst_xlPICT = 1,
	ExcelConst_xlRTF = 4,
	ExcelConst_xlVALU = 8,
	_Excel_XlEditionFormatForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlLink
{
	ExcelConst_xlExcelLinks = 1,
	ExcelConst_xlOLELinks = 2,
	ExcelConst_xlPublishers = 5,
	ExcelConst_xlSubscribers = 6,
	_Excel_XlLinkForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlClipboardFormat
{
	ExcelConst_xlClipboardFormatBIFF = 8,
	ExcelConst_xlClipboardFormatBIFF2 = 18,
	ExcelConst_xlClipboardFormatBIFF3 = 20,
	ExcelConst_xlClipboardFormatBIFF4 = 30,
	ExcelConst_xlClipboardFormatBinary = 15,
	ExcelConst_xlClipboardFormatBitmap = 9,
	ExcelConst_xlClipboardFormatCGM = 13,
	ExcelConst_xlClipboardFormatCSV = 5,
	ExcelConst_xlClipboardFormatDIF = 4,
	ExcelConst_xlClipboardFormatDspText = 12,
	ExcelConst_xlClipboardFormatEmbeddedObject = 21,
	ExcelConst_xlClipboardFormatEmbedSource = 22,
	ExcelConst_xlClipboardFormatLink = 11,
	ExcelConst_xlClipboardFormatLinkSource = 23,
	ExcelConst_xlClipboardFormatLinkSourceDesc = 32,
	ExcelConst_xlClipboardFormatMovie = 24,
	ExcelConst_xlClipboardFormatNative = 14,
	ExcelConst_xlClipboardFormatObjectDesc = 31,
	ExcelConst_xlClipboardFormatObjectLink = 19,
	ExcelConst_xlClipboardFormatOwnerLink = 17,
	ExcelConst_xlClipboardFormatPICT = 2,
	ExcelConst_xlClipboardFormatPrintPICT = 3,
	ExcelConst_xlClipboardFormatRTF = 7,
	ExcelConst_xlClipboardFormatScreenPICT = 29,
	ExcelConst_xlClipboardFormatStandardFont = 28,
	ExcelConst_xlClipboardFormatStandardScale = 27,
	ExcelConst_xlClipboardFormatSYLK = 6,
	ExcelConst_xlClipboardFormatTable = 16,
	ExcelConst_xlClipboardFormatText = 0,
	ExcelConst_xlClipboardFormatToolFace = 25,
	ExcelConst_xlClipboardFormatToolFacePICT = 26,
	ExcelConst_xlClipboardFormatVALU = 1,
	ExcelConst_xlClipboardFormatWK1 = 10,
	_Excel_XlClipboardFormatForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlApplicationInternational
{
	ExcelConst_xl24HourClock = 33,
	ExcelConst_xl4DigitYears = 43,
	ExcelConst_xlAlternateArraySeparator = 16,
	ExcelConst_xlColumnSeparator = 14,
	ExcelConst_xlCountryCode = 1,
	ExcelConst_xlCountrySetting = 2,
	ExcelConst_xlCurrencyBefore = 37,
	ExcelConst_xlCurrencyCode = 25,
	ExcelConst_xlCurrencyDigits = 27,
	ExcelConst_xlCurrencyLeadingZeros = 40,
	ExcelConst_xlCurrencyMinusSign = 38,
	ExcelConst_xlCurrencyNegative = 28,
	ExcelConst_xlCurrencySpaceBefore = 36,
	ExcelConst_xlCurrencyTrailingZeros = 39,
	ExcelConst_xlDateOrder = 32,
	ExcelConst_xlDateSeparator = 17,
	ExcelConst_xlDayCode = 21,
	ExcelConst_xlDayLeadingZero = 42,
	ExcelConst_xlDecimalSeparator = 3,
	ExcelConst_xlGeneralFormatName = 26,
	ExcelConst_xlHourCode = 22,
	ExcelConst_xlLeftBrace = 12,
	ExcelConst_xlLeftBracket = 10,
	ExcelConst_xlListSeparator = 5,
	ExcelConst_xlLowerCaseColumnLetter = 9,
	ExcelConst_xlLowerCaseRowLetter = 8,
	ExcelConst_xlMDY = 44,
	ExcelConst_xlMetric = 35,
	ExcelConst_xlMinuteCode = 23,
	ExcelConst_xlMonthCode = 20,
	ExcelConst_xlMonthLeadingZero = 41,
	ExcelConst_xlMonthNameChars = 30,
	ExcelConst_xlNoncurrencyDigits = 29,
	ExcelConst_xlNonEnglishFunctions = 34,
	ExcelConst_xlRightBrace = 13,
	ExcelConst_xlRightBracket = 11,
	ExcelConst_xlRowSeparator = 15,
	ExcelConst_xlSecondCode = 24,
	ExcelConst_xlThousandsSeparator = 4,
	ExcelConst_xlTimeLeadingZero = 45,
	ExcelConst_xlTimeSeparator = 18,
	ExcelConst_xlUpperCaseColumnLetter = 7,
	ExcelConst_xlUpperCaseRowLetter = 6,
	ExcelConst_xlWeekdayNameChars = 31,
	ExcelConst_xlYearCode = 19,
	_Excel_XlApplicationInternationalForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlHighlightChangesTime
{
	ExcelConst_xlSinceMyLastSave = 1,
	ExcelConst_xlAllChanges = 2,
	ExcelConst_xlNotYetReviewed = 3,
	_Excel_XlHighlightChangesTimeForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlFormatConditionOperator
{
	ExcelConst_xlBetween = 1,
	ExcelConst_xlNotBetween = 2,
	ExcelConst_xlEqual = 3,
	ExcelConst_xlNotEqual = 4,
	ExcelConst_xlGreater = 5,
	ExcelConst_xlLess = 6,
	ExcelConst_xlGreaterEqual = 7,
	ExcelConst_xlLessEqual = 8,
	_Excel_XlFormatConditionOperatorForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlIMEMode
{
	ExcelConst_xlIMEModeNoControl = 0,
	ExcelConst_xlIMEModeOn = 1,
	ExcelConst_xlIMEModeOff = 2,
	ExcelConst_xlIMEModeDisable = 3,
	ExcelConst_xlIMEModeHiragana = 4,
	ExcelConst_xlIMEModeKatakana = 5,
	ExcelConst_xlIMEModeKatakanaHalf = 6,
	ExcelConst_xlIMEModeAlphaFull = 7,
	ExcelConst_xlIMEModeAlpha = 8,
	ExcelConst_xlIMEModeHangulFull = 9,
	ExcelConst_xlIMEModeHangul = 10,
	_Excel_XlIMEModeForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlDVAlertStyle
{
	ExcelConst_xlValidAlertStop = 1,
	ExcelConst_xlValidAlertWarning = 2,
	ExcelConst_xlValidAlertInformation = 3,
	_Excel_XlDVAlertStyleForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlPhoneticCharacterType
{
	ExcelConst_xlKatakanaHalf = 0,
	ExcelConst_xlKatakana = 1,
	ExcelConst_xlHiragana = 2,
	ExcelConst_xlNoConversion = 3,
	_Excel_XlPhoneticCharacterTypeForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlPhoneticAlignment
{
	ExcelConst_xlPhoneticAlignNoControl = 0,
	ExcelConst_xlPhoneticAlignLeft = 1,
	ExcelConst_xlPhoneticAlignCenter = 2,
	ExcelConst_xlPhoneticAlignDistributed = 3,
	_Excel_XlPhoneticAlignmentForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlReferenceType
{
	ExcelConst_xlAbsolute = 1,
	ExcelConst_xlAbsRowRelColumn = 2,
	ExcelConst_xlRelative = 4,
	ExcelConst_xlRelRowAbsColumn = 3,
	_Excel_XlReferenceTypeForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlSaveAction
{
	ExcelConst_xlDoNotSaveChanges = 2,
	ExcelConst_xlSaveChanges = 1,
	_Excel_XlSaveActionForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlSearchOrder
{
	ExcelConst_xlByColumns = 2,
	ExcelConst_xlByRows = 1,
	_Excel_XlSearchOrderForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlSortMethodOld
{
	ExcelConst_xlCodePage = 2,
	ExcelConst_xlSyllabary = 1,
	_Excel_XlSortMethodOldForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlSortType
{
	ExcelConst_xlSortLabels = 2,
	ExcelConst_xlSortValues = 1,
	_Excel_XlSortTypeForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlSpecialCellsValue
{
	ExcelConst_xlErrors = 16,
	ExcelConst_xlLogical = 4,
	ExcelConst_xlNumbers = 1,
	ExcelConst_xlTextValues = 2,
	_Excel_XlSpecialCellsValueForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlTabPosition
{
	ExcelConst_xlTabPositionFirst = 0,
	ExcelConst_xlTabPositionLast = 1,
	_Excel_XlTabPositionForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlWBATemplate
{
	ExcelConst_xlWBATChart = -4109,
	ExcelConst_xlWBATExcel4IntlMacroSheet = 4,
	ExcelConst_xlWBATExcel4MacroSheet = 3,
	ExcelConst_xlWBATWorksheet = -4167,
	_Excel_XlWBATemplateForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_XlToolbarProtection
{
	ExcelConst_xlNoButtonChanges = 1,
	ExcelConst_xlNoChanges = 4,
	ExcelConst_xlNoDockingChanges = 3,
	ExcelConst_xlToolbarProtectionNone = -4143,
	ExcelConst_xlNoShapeChanges = 2,
	_Excel_XlToolbarProtectionForceSizeToFourBytes = 0xFFFFFFFF
};
enum ExcelEnum_xlColumnDataType
{
	ExcelConst_xlGeneralFormat = 1,
	ExcelConst_xlTextFormat = 2,
	ExcelConst_xlMDYFormat = 3,
	ExcelConst_xlDMYFormat = 4,
	ExcelConst_xlYMDFormat = 5,
	ExcelConst_xlMYDFormat = 6,
	ExcelConst_xlDYMFormat = 7,
	ExcelConst_xlYDMFormat = 8,
	ExcelConst_xlSkipColumn = 9,
	ExcelConst_xlEMDFormat = 10,
	_Excel_xlColumnDataTypeForceSizeToFourBytes = 0xFFFFFFFF
};
typedef CAObjHandle ExcelObj_App;
typedef CAObjHandle ExcelObj_Range;
typedef CAObjHandle ExcelObj_Parameters;
typedef CAObjHandle ExcelObj_Chart;
typedef CAObjHandle ExcelObj_DialogSheet;
typedef CAObjHandle ExcelObj_MenuBar;
typedef CAObjHandle ExcelObj_Window;
typedef CAObjHandle ExcelObj_Workbook;
typedef CAObjHandle ExcelObj_AddIns;
typedef CAObjHandle ExcelObj_Assistant;
typedef CAObjHandle ExcelObj_Sheets;
typedef CAObjHandle ExcelObj__CommandBars;
typedef CAObjHandle ExcelObj_MenuBars;
typedef CAObjHandle ExcelObj_Modules;
typedef CAObjHandle ExcelObj_Names;
typedef CAObjHandle ExcelObj_Menu;
typedef CAObjHandle ExcelObj_Toolbars;
typedef CAObjHandle ExcelObj_Windows;
typedef CAObjHandle ExcelObj_Workbooks;
typedef CAObjHandle ExcelObj_WorksheetFunction;
typedef CAObjHandle ExcelObj_AutoCorrect;
typedef CAObjHandle ExcelObj_Dialogs;
typedef CAObjHandle ExcelObj_FileSearch;
typedef CAObjHandle ExcelObj_IFind;
typedef CAObjHandle ExcelObj_RecentFiles;
typedef CAObjHandle ExcelObj_ODBCErrors;
typedef CAObjHandle ExcelObj_VBE;
typedef CAObjHandle ExcelObj_OLEDBErrors;
typedef CAObjHandle ExcelObj_COMAddIns;
typedef CAObjHandle ExcelObj_DefaultWebOptions;
typedef CAObjHandle ExcelObj_LanguageSettings;
typedef CAObjHandle ExcelObj_AnswerWizard;
typedef CAObjHandle ExcelObj_Hyperlink;
typedef CAObjHandle ExcelObj_PageSetup;
typedef CAObjHandle ExcelObj_Shapes;
typedef CAObjHandle ExcelObj_ChartGroup;
typedef CAObjHandle ExcelObj_ChartArea;
typedef CAObjHandle ExcelObj_ChartTitle;
typedef CAObjHandle ExcelObj_Corners;
typedef CAObjHandle ExcelObj_DataTable;
typedef CAObjHandle ExcelObj_Floor;
typedef CAObjHandle ExcelObj_Hyperlinks;
typedef CAObjHandle ExcelObj_Legend;
typedef CAObjHandle ExcelObj_PlotArea;
typedef CAObjHandle ExcelObj_Walls;
typedef CAObjHandle ExcelObj_PivotLayout;
typedef CAObjHandle ExcelObj_Scripts;
typedef CAObjHandle ExcelObj_Outline;
typedef CAObjHandle ExcelObj_PivotTable;
typedef CAObjHandle ExcelObj_HPageBreaks;
typedef CAObjHandle ExcelObj_VPageBreaks;
typedef CAObjHandle ExcelObj_QueryTables;
typedef CAObjHandle ExcelObj_Comments;
typedef CAObjHandle ExcelObj_AutoFilter;
typedef CAObjHandle ExcelObj_Mailer;
typedef CAObjHandle ExcelObj_PivotCaches;
typedef CAObjHandle ExcelObj_RoutingSlip;
typedef CAObjHandle ExcelObj_Styles;
typedef CAObjHandle ExcelObj_CustomViews;
typedef CAObjHandle ExcelObj__VBProject;
typedef CAObjHandle ExcelObj_PublishObjects;
typedef CAObjHandle ExcelObj_WebOptions;
typedef CAObjHandle ExcelObj_HTMLProject;
typedef CAObjHandle ExcelObj_ShapeRange;
typedef CAObjHandle ExcelObj_Border;
typedef CAObjHandle ExcelObj_Interior;
typedef long ExcelType_MsoRGBType;
typedef CAObjHandle ExcelObj_ColorFormat;
typedef CAObjHandle ExcelObj_ShapeNode;
typedef CAObjHandle ExcelObj_WebPageFonts;
typedef CAObjHandle ExcelObj_TreeviewControl;
typedef CAObjHandle ExcelObj_CubeField;
typedef CAObjHandle ExcelObj_Pane;
typedef CAObjHandle ExcelObj_Panes;
typedef CAObjHandle ExcelObj_Areas;
typedef CAObjHandle ExcelObj_Borders;
typedef CAObjHandle ExcelObj_Characters;
typedef CAObjHandle ExcelObj_Font;
typedef CAObjHandle ExcelObj_PivotField;
typedef CAObjHandle ExcelObj_PivotItem;
typedef CAObjHandle ExcelObj_QueryTable;
typedef CAObjHandle ExcelObj_SoundNote;
typedef CAObjHandle ExcelObj_Validation;
typedef CAObjHandle ExcelObj_Worksheet;
typedef CAObjHandle ExcelObj_Comment;
typedef CAObjHandle ExcelObj_Phonetic;
typedef CAObjHandle ExcelObj_FormatConditions;
typedef CAObjHandle ExcelObj_Phonetics;
typedef CAObjHandle ExcelObj_HPageBreak;
typedef CAObjHandle ExcelObj_VPageBreak;
typedef CAObjHandle ExcelObj_RecentFile;
typedef CAObjHandle ExcelObj_Style;
typedef CAObjHandle ExcelObj_AddIn;
typedef CAObjHandle ExcelObj_PivotCache;
typedef CAObjHandle ExcelObj_PivotFormula;
typedef CAObjHandle ExcelObj_CalculatedFields;
typedef CAObjHandle ExcelObj_PivotFormulas;
typedef CAObjHandle ExcelObj_CubeFields;
typedef CAObjHandle ExcelObj_CalculatedItems;
typedef CAObjHandle ExcelObj_Dialog;
typedef CAObjHandle ExcelObj_Scenario;
typedef CAObjHandle ExcelObj_OLEObject;
typedef CAObjHandle ExcelObj_GroupObject;
typedef CAObjHandle ExcelObj_Name;
typedef CAObjHandle ExcelObj_ChartObject;
typedef CAObjHandle ExcelObj_CustomView;
typedef CAObjHandle ExcelObj_FormatCondition;
typedef CAObjHandle ExcelObj_Shape;
typedef CAObjHandle ExcelObj_Parameter;
typedef CAObjHandle ExcelObj_ODBCError;
typedef CAObjHandle ExcelObj_Filters;
typedef CAObjHandle ExcelObj_Filter;
typedef CAObjHandle ExcelObj_ChartColorFormat;
typedef CAObjHandle ExcelObj_AxisTitle;
typedef CAObjHandle ExcelObj_Gridlines;
typedef CAObjHandle ExcelObj_TickLabels;
typedef CAObjHandle ExcelObj_DisplayUnitLabel;
typedef CAObjHandle ExcelObj_ChartFillFormat;
typedef CAObjHandle ExcelObj_DownBars;
typedef CAObjHandle ExcelObj_DropLines;
typedef CAObjHandle ExcelObj_HiLoLines;
typedef CAObjHandle ExcelObj_SeriesLines;
typedef CAObjHandle ExcelObj_UpBars;
typedef CAObjHandle ExcelObj_Axis;
typedef CAObjHandle ExcelObj_Point;
typedef CAObjHandle ExcelObj_DataLabel;
typedef CAObjHandle ExcelObj_ErrorBars;
typedef CAObjHandle ExcelObj_LeaderLines;
typedef CAObjHandle ExcelObj_Series;
typedef CAObjHandle ExcelObj_LegendKey;
typedef CAObjHandle ExcelObj_LegendEntry;
typedef CAObjHandle ExcelObj_Trendline;
typedef CAObjHandle ExcelObj_Adjustments;
typedef CAObjHandle ExcelObj_TextFrame;
typedef CAObjHandle ExcelObj_CalloutFormat;
typedef CAObjHandle ExcelObj_ConnectorFormat;
typedef CAObjHandle ExcelObj_FillFormat;
typedef CAObjHandle ExcelObj_GroupShapes;
typedef CAObjHandle ExcelObj_LineFormat;
typedef CAObjHandle ExcelObj_ShapeNodes;
typedef CAObjHandle ExcelObj_PictureFmt;
typedef CAObjHandle ExcelObj_ShadowFormat;
typedef CAObjHandle ExcelObj_TextEffectFmt;
typedef CAObjHandle ExcelObj_ThreeDFmt;
typedef CAObjHandle ExcelObj_ControlFormat;
typedef CAObjHandle ExcelObj_LinkFormat;
typedef CAObjHandle ExcelObj_OLEFormat;
typedef CAObjHandle ExcelObj_Script;
typedef CAObjHandle ExcelObj_FreeformBuilder;
typedef CAObjHandle ExcelObj_PublishObject;
typedef CAObjHandle ExcelObj_OLEDBError;
typedef HRESULT (CVICALLBACK *RefreshEventsRegOnBeforeRefresh_CallbackType) (CAObjHandle caServerObjHandle,
                                                                             void *caCallbackData,
                                                                             VBOOL *cancel);
typedef HRESULT (CVICALLBACK *RefreshEventsRegOnAfterRefresh_CallbackType) (CAObjHandle caServerObjHandle,
                                                                            void *caCallbackData,
                                                                            VBOOL  success);
typedef HRESULT (CVICALLBACK *AppEventsRegOnNewWorkbook_CallbackType) (CAObjHandle caServerObjHandle,
                                                                       void *caCallbackData,
                                                                       ExcelObj_Workbook  wb);
typedef HRESULT (CVICALLBACK *AppEventsRegOnSheetSelectionChange_CallbackType) (CAObjHandle caServerObjHandle,
                                                                                void *caCallbackData,
                                                                                CAObjHandle  sh,
                                                                                ExcelObj_Range  target);
typedef HRESULT (CVICALLBACK *AppEventsRegOnSheetBeforeDoubleClick_CallbackType) (CAObjHandle caServerObjHandle,
                                                                                  void *caCallbackData,
                                                                                  CAObjHandle  sh,
                                                                                  ExcelObj_Range  target,
                                                                                  VBOOL *cancel);
typedef HRESULT (CVICALLBACK *AppEventsRegOnSheetBeforeRightClick_CallbackType) (CAObjHandle caServerObjHandle,
                                                                                 void *caCallbackData,
                                                                                 CAObjHandle  sh,
                                                                                 ExcelObj_Range  target,
                                                                                 VBOOL *cancel);
typedef HRESULT (CVICALLBACK *AppEventsRegOnSheetActivate_CallbackType) (CAObjHandle caServerObjHandle,
                                                                         void *caCallbackData,
                                                                         CAObjHandle  sh);
typedef HRESULT (CVICALLBACK *AppEventsRegOnSheetDeactivate_CallbackType) (CAObjHandle caServerObjHandle,
                                                                           void *caCallbackData,
                                                                           CAObjHandle  sh);
typedef HRESULT (CVICALLBACK *AppEventsRegOnSheetCalculate_CallbackType) (CAObjHandle caServerObjHandle,
                                                                          void *caCallbackData,
                                                                          CAObjHandle  sh);
typedef HRESULT (CVICALLBACK *AppEventsRegOnSheetChange_CallbackType) (CAObjHandle caServerObjHandle,
                                                                       void *caCallbackData,
                                                                       CAObjHandle  sh,
                                                                       ExcelObj_Range  target);
typedef HRESULT (CVICALLBACK *AppEventsRegOnWorkbookOpen_CallbackType) (CAObjHandle caServerObjHandle,
                                                                        void *caCallbackData,
                                                                        ExcelObj_Workbook  wb);
typedef HRESULT (CVICALLBACK *AppEventsRegOnWorkbookActivate_CallbackType) (CAObjHandle caServerObjHandle,
                                                                            void *caCallbackData,
                                                                            ExcelObj_Workbook  wb);
typedef HRESULT (CVICALLBACK *AppEventsRegOnWorkbookDeactivate_CallbackType) (CAObjHandle caServerObjHandle,
                                                                              void *caCallbackData,
                                                                              ExcelObj_Workbook  wb);
typedef HRESULT (CVICALLBACK *AppEventsRegOnWorkbookBeforeClose_CallbackType) (CAObjHandle caServerObjHandle,
                                                                               void *caCallbackData,
                                                                               ExcelObj_Workbook  wb,
                                                                               VBOOL *cancel);
typedef HRESULT (CVICALLBACK *AppEventsRegOnWorkbookBeforeSave_CallbackType) (CAObjHandle caServerObjHandle,
                                                                              void *caCallbackData,
                                                                              ExcelObj_Workbook  wb,
                                                                              VBOOL  saveAsUI,
                                                                              VBOOL *cancel);
typedef HRESULT (CVICALLBACK *AppEventsRegOnWorkbookBeforePrint_CallbackType) (CAObjHandle caServerObjHandle,
                                                                               void *caCallbackData,
                                                                               ExcelObj_Workbook  wb,
                                                                               VBOOL *cancel);
typedef HRESULT (CVICALLBACK *AppEventsRegOnWorkbookNewSheet_CallbackType) (CAObjHandle caServerObjHandle,
                                                                            void *caCallbackData,
                                                                            ExcelObj_Workbook  wb,
                                                                            CAObjHandle  sh);
typedef HRESULT (CVICALLBACK *AppEventsRegOnWorkbookAddinInstall_CallbackType) (CAObjHandle caServerObjHandle,
                                                                                void *caCallbackData,
                                                                                ExcelObj_Workbook  wb);
typedef HRESULT (CVICALLBACK *AppEventsRegOnWorkbookAddinUninstall_CallbackType) (CAObjHandle caServerObjHandle,
                                                                                  void *caCallbackData,
                                                                                  ExcelObj_Workbook  wb);
typedef HRESULT (CVICALLBACK *AppEventsRegOnWindowResize_CallbackType) (CAObjHandle caServerObjHandle,
                                                                        void *caCallbackData,
                                                                        ExcelObj_Workbook  wb,
                                                                        ExcelObj_Window  wn);
typedef HRESULT (CVICALLBACK *AppEventsRegOnWindowActivate_CallbackType) (CAObjHandle caServerObjHandle,
                                                                          void *caCallbackData,
                                                                          ExcelObj_Workbook  wb,
                                                                          ExcelObj_Window  wn);
typedef HRESULT (CVICALLBACK *AppEventsRegOnWindowDeactivate_CallbackType) (CAObjHandle caServerObjHandle,
                                                                            void *caCallbackData,
                                                                            ExcelObj_Workbook  wb,
                                                                            ExcelObj_Window  wn);
typedef HRESULT (CVICALLBACK *AppEventsRegOnSheetFollowHyperlink_CallbackType) (CAObjHandle caServerObjHandle,
                                                                                void *caCallbackData,
                                                                                CAObjHandle  sh,
                                                                                ExcelObj_Hyperlink  target);
typedef HRESULT (CVICALLBACK *ChartEventsRegOnActivate_CallbackType) (CAObjHandle caServerObjHandle,
                                                                      void *caCallbackData);
typedef HRESULT (CVICALLBACK *ChartEventsRegOnDeactivate_CallbackType) (CAObjHandle caServerObjHandle,
                                                                        void *caCallbackData);
typedef HRESULT (CVICALLBACK *ChartEventsRegOnResize_CallbackType) (CAObjHandle caServerObjHandle,
                                                                    void *caCallbackData);
typedef HRESULT (CVICALLBACK *ChartEventsRegOnMouseDown_CallbackType) (CAObjHandle caServerObjHandle,
                                                                       void *caCallbackData,
                                                                       long  button,
                                                                       long  shift,
                                                                       long  x,
                                                                       long  y);
typedef HRESULT (CVICALLBACK *ChartEventsRegOnMouseUp_CallbackType) (CAObjHandle caServerObjHandle,
                                                                     void *caCallbackData,
                                                                     long  button,
                                                                     long  shift,
                                                                     long  x,
                                                                     long  y);
typedef HRESULT (CVICALLBACK *ChartEventsRegOnMouseMove_CallbackType) (CAObjHandle caServerObjHandle,
                                                                       void *caCallbackData,
                                                                       long  button,
                                                                       long  shift,
                                                                       long  x,
                                                                       long  y);
typedef HRESULT (CVICALLBACK *ChartEventsRegOnBeforeRightClick_CallbackType) (CAObjHandle caServerObjHandle,
                                                                              void *caCallbackData,
                                                                              VBOOL *cancel);
typedef HRESULT (CVICALLBACK *ChartEventsRegOnDragPlot_CallbackType) (CAObjHandle caServerObjHandle,
                                                                      void *caCallbackData);
typedef HRESULT (CVICALLBACK *ChartEventsRegOnDragOver_CallbackType) (CAObjHandle caServerObjHandle,
                                                                      void *caCallbackData);
typedef HRESULT (CVICALLBACK *ChartEventsRegOnBeforeDoubleClick_CallbackType) (CAObjHandle caServerObjHandle,
                                                                               void *caCallbackData,
                                                                               long  elementID,
                                                                               long  arg1,
                                                                               long  arg2,
                                                                               VBOOL *cancel);
typedef HRESULT (CVICALLBACK *ChartEventsRegOnSelect_CallbackType) (CAObjHandle caServerObjHandle,
                                                                    void *caCallbackData,
                                                                    long  elementID,
                                                                    long  arg1,
                                                                    long  arg2);
typedef HRESULT (CVICALLBACK *ChartEventsRegOnSeriesChange_CallbackType) (CAObjHandle caServerObjHandle,
                                                                          void *caCallbackData,
                                                                          long  seriesIndex,
                                                                          long  pointIndex);
typedef HRESULT (CVICALLBACK *ChartEventsRegOnCalculate_CallbackType) (CAObjHandle caServerObjHandle,
                                                                       void *caCallbackData);
typedef HRESULT (CVICALLBACK *DocEventsRegOnSelectionChange_CallbackType) (CAObjHandle caServerObjHandle,
                                                                           void *caCallbackData,
                                                                           ExcelObj_Range  target);
typedef HRESULT (CVICALLBACK *DocEventsRegOnBeforeDoubleClick_CallbackType) (CAObjHandle caServerObjHandle,
                                                                             void *caCallbackData,
                                                                             ExcelObj_Range  target,
                                                                             VBOOL *cancel);
typedef HRESULT (CVICALLBACK *DocEventsRegOnBeforeRightClick_CallbackType) (CAObjHandle caServerObjHandle,
                                                                            void *caCallbackData,
                                                                            ExcelObj_Range  target,
                                                                            VBOOL *cancel);
typedef HRESULT (CVICALLBACK *DocEventsRegOnActivate_CallbackType) (CAObjHandle caServerObjHandle,
                                                                    void *caCallbackData);
typedef HRESULT (CVICALLBACK *DocEventsRegOnDeactivate_CallbackType) (CAObjHandle caServerObjHandle,
                                                                      void *caCallbackData);
typedef HRESULT (CVICALLBACK *DocEventsRegOnCalculate_CallbackType) (CAObjHandle caServerObjHandle,
                                                                     void *caCallbackData);
typedef HRESULT (CVICALLBACK *DocEventsRegOnChange_CallbackType) (CAObjHandle caServerObjHandle,
                                                                  void *caCallbackData,
                                                                  ExcelObj_Range  target);
typedef HRESULT (CVICALLBACK *DocEventsRegOnFollowHyperlink_CallbackType) (CAObjHandle caServerObjHandle,
                                                                           void *caCallbackData,
                                                                           ExcelObj_Hyperlink  target);
typedef HRESULT (CVICALLBACK *WorkbkEventsRegOnOpen_CallbackType) (CAObjHandle caServerObjHandle,
                                                                   void *caCallbackData);
typedef HRESULT (CVICALLBACK *WorkbkEventsRegOnActivate_CallbackType) (CAObjHandle caServerObjHandle,
                                                                       void *caCallbackData);
typedef HRESULT (CVICALLBACK *WorkbkEventsRegOnDeactivate_CallbackType) (CAObjHandle caServerObjHandle,
                                                                         void *caCallbackData);
typedef HRESULT (CVICALLBACK *WorkbkEventsRegOnBeforeClose_CallbackType) (CAObjHandle caServerObjHandle,
                                                                          void *caCallbackData,
                                                                          VBOOL *cancel);
typedef HRESULT (CVICALLBACK *WorkbkEventsRegOnBeforeSave_CallbackType) (CAObjHandle caServerObjHandle,
                                                                         void *caCallbackData,
                                                                         VBOOL  saveAsUI,
                                                                         VBOOL *cancel);
typedef HRESULT (CVICALLBACK *WorkbkEventsRegOnBeforePrint_CallbackType) (CAObjHandle caServerObjHandle,
                                                                          void *caCallbackData,
                                                                          VBOOL *cancel);
typedef HRESULT (CVICALLBACK *WorkbkEventsRegOnNewSheet_CallbackType) (CAObjHandle caServerObjHandle,
                                                                       void *caCallbackData,
                                                                       CAObjHandle  sh);
typedef HRESULT (CVICALLBACK *WorkbkEventsRegOnAddinInstall_CallbackType) (CAObjHandle caServerObjHandle,
                                                                           void *caCallbackData);
typedef HRESULT (CVICALLBACK *WorkbkEventsRegOnAddinUninstall_CallbackType) (CAObjHandle caServerObjHandle,
                                                                             void *caCallbackData);
typedef HRESULT (CVICALLBACK *WorkbkEventsRegOnWindowResize_CallbackType) (CAObjHandle caServerObjHandle,
                                                                           void *caCallbackData,
                                                                           ExcelObj_Window  wn);
typedef HRESULT (CVICALLBACK *WorkbkEventsRegOnWindowActivate_CallbackType) (CAObjHandle caServerObjHandle,
                                                                             void *caCallbackData,
                                                                             ExcelObj_Window  wn);
typedef HRESULT (CVICALLBACK *WorkbkEventsRegOnWindowDeactivate_CallbackType) (CAObjHandle caServerObjHandle,
                                                                               void *caCallbackData,
                                                                               ExcelObj_Window  wn);
typedef HRESULT (CVICALLBACK *WorkbkEventsRegOnSheetSelectionChange_CallbackType) (CAObjHandle caServerObjHandle,
                                                                                   void *caCallbackData,
                                                                                   CAObjHandle  sh,
                                                                                   ExcelObj_Range  target);
typedef HRESULT (CVICALLBACK *WorkbkEventsRegOnSheetBeforeDoubleClick_CallbackType) (CAObjHandle caServerObjHandle,
                                                                                     void *caCallbackData,
                                                                                     CAObjHandle  sh,
                                                                                     ExcelObj_Range  target,
                                                                                     VBOOL *cancel);
typedef HRESULT (CVICALLBACK *WorkbkEventsRegOnSheetBeforeRightClick_CallbackType) (CAObjHandle caServerObjHandle,
                                                                                    void *caCallbackData,
                                                                                    CAObjHandle  sh,
                                                                                    ExcelObj_Range  target,
                                                                                    VBOOL *cancel);
typedef HRESULT (CVICALLBACK *WorkbkEventsRegOnSheetActivate_CallbackType) (CAObjHandle caServerObjHandle,
                                                                            void *caCallbackData,
                                                                            CAObjHandle  sh);
typedef HRESULT (CVICALLBACK *WorkbkEventsRegOnSheetDeactivate_CallbackType) (CAObjHandle caServerObjHandle,
                                                                              void *caCallbackData,
                                                                              CAObjHandle  sh);
typedef HRESULT (CVICALLBACK *WorkbkEventsRegOnSheetCalculate_CallbackType) (CAObjHandle caServerObjHandle,
                                                                             void *caCallbackData,
                                                                             CAObjHandle  sh);
typedef HRESULT (CVICALLBACK *WorkbkEventsRegOnSheetChange_CallbackType) (CAObjHandle caServerObjHandle,
                                                                          void *caCallbackData,
                                                                          CAObjHandle  sh,
                                                                          ExcelObj_Range  target);
typedef HRESULT (CVICALLBACK *WorkbkEventsRegOnSheetFollowHyperlink_CallbackType) (CAObjHandle caServerObjHandle,
                                                                                   void *caCallbackData,
                                                                                   CAObjHandle  sh,
                                                                                   ExcelObj_Hyperlink  target);
typedef HRESULT (CVICALLBACK *OLEObjectEventsRegOnGotFocus_CallbackType) (CAObjHandle caServerObjHandle,
                                                                          void *caCallbackData);
typedef HRESULT (CVICALLBACK *OLEObjectEventsRegOnLostFocus_CallbackType) (CAObjHandle caServerObjHandle,
                                                                           void *caCallbackData);
/* NICDBLD_END> Type Library Specific Types */

extern const IID Excel_IID_QueryTable;
extern const IID Excel_IID_App;
extern const IID Excel_IID_Chart;
extern const IID Excel_IID_Worksheet;
extern const IID Excel_IID_Workbook;
extern const IID Excel_IID_OLEObject;
extern const IID Excel_IID_RefreshEvents;
extern const IID Excel_IID_AppEvents;
extern const IID Excel_IID_ChartEvents;
extern const IID Excel_IID_DocEvents;
extern const IID Excel_IID_WorkbkEvents;
extern const IID Excel_IID_OLEObjectEvents;
extern const IID Excel_IID_Adjustments;
extern const IID Excel_IID_CalloutFormat;
extern const IID Excel_IID_ColorFormat;
extern const IID Excel_IID_LineFormat;
extern const IID Excel_IID_ShapeNode;
extern const IID Excel_IID_ShapeNodes;
extern const IID Excel_IID_PictureFmt;
extern const IID Excel_IID_ShadowFormat;
extern const IID Excel_IID_TextEffectFmt;
extern const IID Excel_IID_ThreeDFmt;
extern const IID Excel_IID_FillFormat;
extern const IID Excel_IID_Sheets;
extern const IID Excel_IID_Workbooks;
extern const IID Excel_IID_PublishObject;
extern const IID Excel_IID_DefaultWebOptions;
extern const IID Excel_IID_WebOptions;
extern const IID Excel_IID_TreeviewControl;
extern const IID Excel_IID_CubeField;
extern const IID Excel_IID_CubeFields;
extern const IID Excel_IID_Font;
extern const IID Excel_IID_Window;
extern const IID Excel_IID_Windows;
extern const IID Excel_IID_WorksheetFunction;
extern const IID Excel_IID_Range;
extern const IID Excel_IID_VPageBreak;
extern const IID Excel_IID_HPageBreak;
extern const IID Excel_IID_HPageBreaks;
extern const IID Excel_IID_VPageBreaks;
extern const IID Excel_IID_RecentFile;
extern const IID Excel_IID_RecentFiles;
extern const IID Excel_IID_Style;
extern const IID Excel_IID_Styles;
extern const IID Excel_IID_Borders;
extern const IID Excel_IID_AddIn;
extern const IID Excel_IID_AddIns;
extern const IID Excel_IID_Areas;
extern const IID Excel_IID_Charts;
extern const IID Excel_IID_PivotCache;
extern const IID Excel_IID_PivotCaches;
extern const IID Excel_IID_PivotFormula;
extern const IID Excel_IID_PivotFormulas;
extern const IID Excel_IID_PivotTable;
extern const IID Excel_IID_PivotTables;
extern const IID Excel_IID_PivotField;
extern const IID Excel_IID_PivotFields;
extern const IID Excel_IID_CalculatedFields;
extern const IID Excel_IID_PivotItem;
extern const IID Excel_IID_PivotItems;
extern const IID Excel_IID_CalculatedItems;
extern const IID Excel_IID_Characters;
extern const IID Excel_IID_Dialogs;
extern const IID Excel_IID_Dialog;
extern const IID Excel_IID_SoundNote;
extern const IID Excel_IID_Panes;
extern const IID Excel_IID_Pane;
extern const IID Excel_IID_Scenarios;
extern const IID Excel_IID_Scenario;
extern const IID Excel_IID_OLEObjects;
extern const IID Excel_IID_RoutingSlip;
extern const IID Excel_IID_Outline;
extern const IID Excel_IID_Worksheets;
extern const IID Excel_IID_PageSetup;
extern const IID Excel_IID_Names;
extern const IID Excel_IID_Name;
extern const IID Excel_IID_ChartObject;
extern const IID Excel_IID_ChartObjects;
extern const IID Excel_IID_Mailer;
extern const IID Excel_IID_CustomViews;
extern const IID Excel_IID_CustomView;
extern const IID Excel_IID_FormatConditions;
extern const IID Excel_IID_FormatCondition;
extern const IID Excel_IID_Comments;
extern const IID Excel_IID_Comment;
extern const IID Excel_IID_QueryTables;
extern const IID Excel_IID_Parameter;
extern const IID Excel_IID_Parameters;
extern const IID Excel_IID_ODBCError;
extern const IID Excel_IID_ODBCErrors;
extern const IID Excel_IID_Validation;
extern const IID Excel_IID_Hyperlinks;
extern const IID Excel_IID_Hyperlink;
extern const IID Excel_IID_AutoFilter;
extern const IID Excel_IID_Filters;
extern const IID Excel_IID_Filter;
extern const IID Excel_IID_AutoCorrect;
extern const IID Excel_IID_Border;
extern const IID Excel_IID_Interior;
extern const IID Excel_IID_ChartFillFormat;
extern const IID Excel_IID_ChartColorFormat;
extern const IID Excel_IID_Axis;
extern const IID Excel_IID_ChartTitle;
extern const IID Excel_IID_AxisTitle;
extern const IID Excel_IID_ChartGroup;
extern const IID Excel_IID_ChartGroups;
extern const IID Excel_IID_Axes;
extern const IID Excel_IID_Points;
extern const IID Excel_IID_Point;
extern const IID Excel_IID_Series;
extern const IID Excel_IID_SeriesCollection;
extern const IID Excel_IID_DataLabel;
extern const IID Excel_IID_DataLabels;
extern const IID Excel_IID_LegendEntry;
extern const IID Excel_IID_LegendEntries;
extern const IID Excel_IID_LegendKey;
extern const IID Excel_IID_Trendlines;
extern const IID Excel_IID_Trendline;
extern const IID Excel_IID_Corners;
extern const IID Excel_IID_SeriesLines;
extern const IID Excel_IID_HiLoLines;
extern const IID Excel_IID_Gridlines;
extern const IID Excel_IID_DropLines;
extern const IID Excel_IID_LeaderLines;
extern const IID Excel_IID_UpBars;
extern const IID Excel_IID_DownBars;
extern const IID Excel_IID_Floor;
extern const IID Excel_IID_Walls;
extern const IID Excel_IID_TickLabels;
extern const IID Excel_IID_PlotArea;
extern const IID Excel_IID_ChartArea;
extern const IID Excel_IID_Legend;
extern const IID Excel_IID_ErrorBars;
extern const IID Excel_IID_DataTable;
extern const IID Excel_IID_Phonetic;
extern const IID Excel_IID_Shape;
extern const IID Excel_IID_Shapes;
extern const IID Excel_IID_ShapeRange;
extern const IID Excel_IID_GroupShapes;
extern const IID Excel_IID_TextFrame;
extern const IID Excel_IID_ConnectorFormat;
extern const IID Excel_IID_FreeformBuilder;
extern const IID Excel_IID_ControlFormat;
extern const IID Excel_IID_OLEFormat;
extern const IID Excel_IID_LinkFormat;
extern const IID Excel_IID_PublishObjects;
extern const IID Excel_IID_OLEDBError;
extern const IID Excel_IID_OLEDBErrors;
extern const IID Excel_IID_Phonetics;
extern const IID Excel_IID_PivotLayout;
extern const IID Excel_IID_DisplayUnitLabel;

HRESULT CVIFUNC Excel_GetProperty (CAObjHandle objectHandle,
                                   ERRORINFO *errorInfo, unsigned int propertyID,
                                   unsigned int propertyType, void *value);

HRESULT CVIFUNC_C Excel_SetProperty (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo,
                                     unsigned int propertyID,
                                     unsigned int propertyType, ...);

HRESULT CVIFUNC_C Excel_SetPropertyByRef (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo,
                                          unsigned int propertyID,
                                          unsigned int propertyType, ...);

HRESULT CVIFUNC Excel_NewQueryTable (const char *server,
                                     int supportMultithreading, LCID locale,
                                     int reserved, CAObjHandle *objectHandle);

HRESULT CVIFUNC Excel_OpenQueryTable (const char *fileName, const char *server,
                                      int supportMultithreading, LCID locale,
                                      int reserved, CAObjHandle *objectHandle);

HRESULT CVIFUNC Excel_ActiveQueryTable (const char *server,
                                        int supportMultithreading, LCID locale,
                                        int reserved, CAObjHandle *objectHandle);

HRESULT CVIFUNC Excel_QueryTableCancelRefresh (CAObjHandle objectHandle,
                                               ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_QueryTableDelete (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_QueryTableRefresh (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo,
                                         VARIANT backgroundQuery,
                                         VBOOL *returnValue);

HRESULT CVIFUNC Excel_QueryTableResetTimer (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_NewApp (const char *server, int supportMultithreading,
                              LCID locale, int reserved,
                              CAObjHandle *objectHandle);

HRESULT CVIFUNC Excel_OpenApp (const char *fileName, const char *server,
                               int supportMultithreading, LCID locale,
                               int reserved, CAObjHandle *objectHandle);

HRESULT CVIFUNC Excel_ActiveApp (const char *server, int supportMultithreading,
                                 LCID locale, int reserved,
                                 CAObjHandle *objectHandle);

HRESULT CVIFUNC Excel_AppCalculate (CAObjHandle objectHandle,
                                    ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_AppDDEExecute (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo, long channel,
                                     const char *string);

HRESULT CVIFUNC Excel_AppDDEInitiate (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, const char *app,
                                      const char *topic, long *returnValue);

HRESULT CVIFUNC Excel_AppDDEPoke (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                  long channel, VARIANT item, VARIANT data);

HRESULT CVIFUNC Excel_AppDDERequest (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo, long channel,
                                     const char *item, VARIANT *returnValue);

HRESULT CVIFUNC Excel_AppDDETerminate (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo, long channel);

HRESULT CVIFUNC Excel_AppEvaluate (CAObjHandle objectHandle,
                                   ERRORINFO *errorInfo, VARIANT name,
                                   VARIANT *returnValue);

HRESULT CVIFUNC Excel_App_Evaluate (CAObjHandle objectHandle,
                                    ERRORINFO *errorInfo, VARIANT name,
                                    VARIANT *returnValue);

HRESULT CVIFUNC Excel_AppExecuteExcel4Macro (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo,
                                             const char *string,
                                             VARIANT *returnValue);

HRESULT CVIFUNC Excel_AppIntersect (CAObjHandle objectHandle,
                                    ERRORINFO *errorInfo, ExcelObj_Range arg1,
                                    ExcelObj_Range arg2, VARIANT arg3,
                                    VARIANT arg4, VARIANT arg5, VARIANT arg6,
                                    VARIANT arg7, VARIANT arg8, VARIANT arg9,
                                    VARIANT arg10, VARIANT arg11, VARIANT arg12,
                                    VARIANT arg13, VARIANT arg14, VARIANT arg15,
                                    VARIANT arg16, VARIANT arg17, VARIANT arg18,
                                    VARIANT arg19, VARIANT arg20, VARIANT arg21,
                                    VARIANT arg22, VARIANT arg23, VARIANT arg24,
                                    VARIANT arg25, VARIANT arg26, VARIANT arg27,
                                    VARIANT arg28, VARIANT arg29, VARIANT arg30,
                                    ExcelObj_Range *returnValue);

HRESULT CVIFUNC Excel_AppRange (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                VARIANT cell1, VARIANT cell2,
                                ExcelObj_Range *returnValue);

HRESULT CVIFUNC Excel_AppRun (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                              VARIANT macro, VARIANT arg1, VARIANT arg2,
                              VARIANT arg3, VARIANT arg4, VARIANT arg5,
                              VARIANT arg6, VARIANT arg7, VARIANT arg8,
                              VARIANT arg9, VARIANT arg10, VARIANT arg11,
                              VARIANT arg12, VARIANT arg13, VARIANT arg14,
                              VARIANT arg15, VARIANT arg16, VARIANT arg17,
                              VARIANT arg18, VARIANT arg19, VARIANT arg20,
                              VARIANT arg21, VARIANT arg22, VARIANT arg23,
                              VARIANT arg24, VARIANT arg25, VARIANT arg26,
                              VARIANT arg27, VARIANT arg28, VARIANT arg29,
                              VARIANT arg30, VARIANT *returnValue);

HRESULT CVIFUNC Excel_App_Run2 (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                VARIANT macro, VARIANT arg1, VARIANT arg2,
                                VARIANT arg3, VARIANT arg4, VARIANT arg5,
                                VARIANT arg6, VARIANT arg7, VARIANT arg8,
                                VARIANT arg9, VARIANT arg10, VARIANT arg11,
                                VARIANT arg12, VARIANT arg13, VARIANT arg14,
                                VARIANT arg15, VARIANT arg16, VARIANT arg17,
                                VARIANT arg18, VARIANT arg19, VARIANT arg20,
                                VARIANT arg21, VARIANT arg22, VARIANT arg23,
                                VARIANT arg24, VARIANT arg25, VARIANT arg26,
                                VARIANT arg27, VARIANT arg28, VARIANT arg29,
                                VARIANT arg30, VARIANT *returnValue);

HRESULT CVIFUNC Excel_AppSendKeys (CAObjHandle objectHandle,
                                   ERRORINFO *errorInfo, VARIANT keys,
                                   VARIANT wait);

HRESULT CVIFUNC Excel_AppShortcutMenus (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo, long index,
                                        ExcelObj_Menu *returnValue);

HRESULT CVIFUNC Excel_AppUnion (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                ExcelObj_Range arg1, ExcelObj_Range arg2,
                                VARIANT arg3, VARIANT arg4, VARIANT arg5,
                                VARIANT arg6, VARIANT arg7, VARIANT arg8,
                                VARIANT arg9, VARIANT arg10, VARIANT arg11,
                                VARIANT arg12, VARIANT arg13, VARIANT arg14,
                                VARIANT arg15, VARIANT arg16, VARIANT arg17,
                                VARIANT arg18, VARIANT arg19, VARIANT arg20,
                                VARIANT arg21, VARIANT arg22, VARIANT arg23,
                                VARIANT arg24, VARIANT arg25, VARIANT arg26,
                                VARIANT arg27, VARIANT arg28, VARIANT arg29,
                                VARIANT arg30, ExcelObj_Range *returnValue);

HRESULT CVIFUNC Excel_AppActivateMicrosoftApp (CAObjHandle objectHandle,
                                               ERRORINFO *errorInfo,
                                               enum ExcelEnum_XlMSApplication index);

HRESULT CVIFUNC Excel_AppAddChartAutoFormat (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo, VARIANT chart,
                                             const char *name,
                                             VARIANT description);

HRESULT CVIFUNC Excel_AppAddCustomList (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo, VARIANT listArray,
                                        VARIANT byRow);

HRESULT CVIFUNC Excel_AppCaller (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                 VARIANT index, VARIANT *returnValue);

HRESULT CVIFUNC Excel_AppCentimetersToPoints (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo,
                                              double centimeters,
                                              double *returnValue);

HRESULT CVIFUNC Excel_AppCheckSpelling (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo, const char *word,
                                        VARIANT customDictionary,
                                        VARIANT ignoreUppercase,
                                        VBOOL *returnValue);

HRESULT CVIFUNC Excel_AppClipboardFormats (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo, VARIANT index,
                                           VARIANT *returnValue);

HRESULT CVIFUNC Excel_AppConvertFormula (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo, VARIANT formula,
                                         enum ExcelEnum_XlReferenceStyle fromReferenceStyle,
                                         VARIANT toReferenceStyle,
                                         VARIANT toAbsolute, VARIANT relativeTo,
                                         VARIANT *returnValue);

HRESULT CVIFUNC Excel_AppDummy1 (CAObjHandle objectHandle, ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_AppDummy2 (CAObjHandle objectHandle, ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_AppDummy3 (CAObjHandle objectHandle, ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_AppDummy4 (CAObjHandle objectHandle, ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_AppDummy5 (CAObjHandle objectHandle, ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_AppDummy6 (CAObjHandle objectHandle, ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_AppDummy7 (CAObjHandle objectHandle, ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_AppDummy8 (CAObjHandle objectHandle, ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_AppDummy9 (CAObjHandle objectHandle, ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_AppDummy10 (CAObjHandle objectHandle, ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_AppDummy11 (CAObjHandle objectHandle, ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_AppDeleteChartAutoFormat (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo,
                                                const char *name);

HRESULT CVIFUNC Excel_AppDeleteCustomList (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo, long listNum);

HRESULT CVIFUNC Excel_AppDoubleClick (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_AppFileConverters (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo, VARIANT index1,
                                         VARIANT index2, VARIANT *returnValue);

HRESULT CVIFUNC Excel_App_FindFile (CAObjHandle objectHandle,
                                    ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_AppGetCustomListContents (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo,
                                                long listNum,
                                                VARIANT *returnValue);

HRESULT CVIFUNC Excel_AppGetCustomListNum (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo,
                                           VARIANT listArray, long *returnValue);

HRESULT CVIFUNC Excel_AppGetOpenFilename (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo,
                                          VARIANT fileFilter,
                                          VARIANT filterIndex, VARIANT title,
                                          VARIANT buttonText,
                                          VARIANT multiSelect,
                                          VARIANT *returnValue);

HRESULT CVIFUNC Excel_AppGetSaveAsFilename (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo,
                                            VARIANT initialFilename,
                                            VARIANT fileFilter,
                                            VARIANT filterIndex, VARIANT title,
                                            VARIANT buttonText,
                                            VARIANT *returnValue);

HRESULT CVIFUNC Excel_AppGoto (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                               VARIANT reference, VARIANT scroll);

HRESULT CVIFUNC Excel_AppHelp (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                               VARIANT helpFile, VARIANT helpContextID);

HRESULT CVIFUNC Excel_AppInchesToPoints (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo, double inches,
                                         double *returnValue);

HRESULT CVIFUNC Excel_AppInputBox (CAObjHandle objectHandle,
                                   ERRORINFO *errorInfo, const char *prompt,
                                   VARIANT title, VARIANT _default, VARIANT left,
                                   VARIANT top, VARIANT helpFile,
                                   VARIANT helpContextID, VARIANT type,
                                   VARIANT *returnValue);

HRESULT CVIFUNC Excel_AppInternational (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo, VARIANT index,
                                        VARIANT *returnValue);

HRESULT CVIFUNC Excel_AppMacroOptions (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo, VARIANT macro,
                                       VARIANT description, VARIANT hasMenu,
                                       VARIANT menuText, VARIANT hasShortcutKey,
                                       VARIANT shortcutKey, VARIANT category,
                                       VARIANT statusBar, VARIANT helpContextID,
                                       VARIANT helpFile);

HRESULT CVIFUNC Excel_AppMailLogoff (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_AppMailLogon (CAObjHandle objectHandle,
                                    ERRORINFO *errorInfo, VARIANT name,
                                    VARIANT password, VARIANT downloadNewMail);

HRESULT CVIFUNC Excel_AppNextLetter (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo,
                                     ExcelObj_Workbook *returnValue);

HRESULT CVIFUNC Excel_AppOnKey (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                const char *key, VARIANT procedure);

HRESULT CVIFUNC Excel_AppOnRepeat (CAObjHandle objectHandle,
                                   ERRORINFO *errorInfo, const char *text,
                                   const char *procedure);

HRESULT CVIFUNC Excel_AppOnTime (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                 VARIANT earliestTime, const char *procedure,
                                 VARIANT latestTime, VARIANT schedule);

HRESULT CVIFUNC Excel_AppOnUndo (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                 const char *text, const char *procedure);

HRESULT CVIFUNC Excel_AppPreviousSelections (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo, VARIANT index,
                                             VARIANT *returnValue);

HRESULT CVIFUNC Excel_AppQuit (CAObjHandle objectHandle, ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_AppRecordMacro (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT basicCode,
                                      VARIANT xlmCode);

HRESULT CVIFUNC Excel_AppRegisteredFunctions (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo,
                                              VARIANT index1, VARIANT index2,
                                              VARIANT *returnValue);

HRESULT CVIFUNC Excel_AppRegisterXLL (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, const char *filename,
                                      VBOOL *returnValue);

HRESULT CVIFUNC Excel_AppRepeat (CAObjHandle objectHandle, ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_AppResetTipWizard (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_AppSave (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                               VARIANT filename);

HRESULT CVIFUNC Excel_AppSaveWorkspace (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo, VARIANT filename);

HRESULT CVIFUNC Excel_AppSetDefaultChart (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo,
                                          VARIANT formatName, VARIANT gallery);

HRESULT CVIFUNC Excel_AppUndo (CAObjHandle objectHandle, ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_AppVolatile (CAObjHandle objectHandle,
                                   ERRORINFO *errorInfo, VARIANT _volatile);

HRESULT CVIFUNC Excel_App_Wait (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                VARIANT time);

HRESULT CVIFUNC Excel_App_WSFunction (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT arg1,
                                      VARIANT arg2, VARIANT arg3, VARIANT arg4,
                                      VARIANT arg5, VARIANT arg6, VARIANT arg7,
                                      VARIANT arg8, VARIANT arg9, VARIANT arg10,
                                      VARIANT arg11, VARIANT arg12,
                                      VARIANT arg13, VARIANT arg14,
                                      VARIANT arg15, VARIANT arg16,
                                      VARIANT arg17, VARIANT arg18,
                                      VARIANT arg19, VARIANT arg20,
                                      VARIANT arg21, VARIANT arg22,
                                      VARIANT arg23, VARIANT arg24,
                                      VARIANT arg25, VARIANT arg26,
                                      VARIANT arg27, VARIANT arg28,
                                      VARIANT arg29, VARIANT arg30,
                                      VARIANT *returnValue);

HRESULT CVIFUNC Excel_AppWait (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                               VARIANT time, VBOOL *returnValue);

HRESULT CVIFUNC Excel_AppGetPhonetic (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT text,
                                      char **returnValue);

HRESULT CVIFUNC Excel_AppDummy12 (CAObjHandle objectHandle, ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_AppCalculateFull (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_AppFindFile (CAObjHandle objectHandle,
                                   ERRORINFO *errorInfo, VBOOL *returnValue);

HRESULT CVIFUNC Excel_NewChart (const char *server, int supportMultithreading,
                                LCID locale, int reserved,
                                CAObjHandle *objectHandle);

HRESULT CVIFUNC Excel_OpenChart (const char *fileName, const char *server,
                                 int supportMultithreading, LCID locale,
                                 int reserved, CAObjHandle *objectHandle);

HRESULT CVIFUNC Excel_ActiveChart (const char *server, int supportMultithreading,
                                   LCID locale, int reserved,
                                   CAObjHandle *objectHandle);

HRESULT CVIFUNC Excel_ChartActivate (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_ChartCopy (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                 VARIANT before, VARIANT after);

HRESULT CVIFUNC Excel_ChartDelete (CAObjHandle objectHandle,
                                   ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_ChartMove (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                 VARIANT before, VARIANT after);

HRESULT CVIFUNC Excel_Chart_PrintOut (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT from,
                                      VARIANT to, VARIANT copies,
                                      VARIANT preview, VARIANT activePrinter,
                                      VARIANT printToFile, VARIANT collate);

HRESULT CVIFUNC Excel_ChartPrintPreview (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo,
                                         VARIANT enableChanges);

HRESULT CVIFUNC Excel_ChartProtect (CAObjHandle objectHandle,
                                    ERRORINFO *errorInfo, VARIANT password,
                                    VARIANT drawingObjects, VARIANT contents,
                                    VARIANT scenarios, VARIANT userInterfaceOnly);

HRESULT CVIFUNC Excel_Chart_Dummy23 (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_ChartSaveAs (CAObjHandle objectHandle,
                                   ERRORINFO *errorInfo, const char *filename,
                                   VARIANT fileFormat, VARIANT password,
                                   VARIANT writeResPassword,
                                   VARIANT readOnlyRecommended,
                                   VARIANT createBackup, VARIANT addToMru,
                                   VARIANT textCodepage,
                                   VARIANT textVisualLayout);

HRESULT CVIFUNC Excel_ChartSelect (CAObjHandle objectHandle,
                                   ERRORINFO *errorInfo, VARIANT replace);

HRESULT CVIFUNC Excel_ChartUnprotect (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT password);

HRESULT CVIFUNC Excel_ChartApplyDataLabels (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo,
                                            enum ExcelEnum_XlDataLabelsType type,
                                            VARIANT legendKey, VARIANT autoText,
                                            VARIANT hasLeaderLines);

HRESULT CVIFUNC Excel_ChartArcs (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                 VARIANT index, CAObjHandle *returnValue);

HRESULT CVIFUNC Excel_ChartAreaGroups (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo, VARIANT index,
                                       CAObjHandle *returnValue);

HRESULT CVIFUNC Excel_ChartAutoFormat (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo, long gallery,
                                       VARIANT format);

HRESULT CVIFUNC Excel_ChartAxes (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                 VARIANT type,
                                 enum ExcelEnum_XlAxisGroup axisGroup,
                                 CAObjHandle *returnValue);

HRESULT CVIFUNC Excel_ChartSetBackgroundPicture (CAObjHandle objectHandle,
                                                 ERRORINFO *errorInfo,
                                                 const char *filename);

HRESULT CVIFUNC Excel_ChartBarGroups (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT index,
                                      CAObjHandle *returnValue);

HRESULT CVIFUNC Excel_ChartButtons (CAObjHandle objectHandle,
                                    ERRORINFO *errorInfo, VARIANT index,
                                    CAObjHandle *returnValue);

HRESULT CVIFUNC Excel_ChartChartGroups (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo, VARIANT index,
                                        CAObjHandle *returnValue);

HRESULT CVIFUNC Excel_ChartChartObjects (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo, VARIANT index,
                                         CAObjHandle *returnValue);

HRESULT CVIFUNC Excel_ChartChartWizard (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo, VARIANT source,
                                        VARIANT gallery, VARIANT format,
                                        VARIANT plotBy, VARIANT categoryLabels,
                                        VARIANT seriesLabels, VARIANT hasLegend,
                                        VARIANT title, VARIANT categoryTitle,
                                        VARIANT valueTitle, VARIANT extraTitle);

HRESULT CVIFUNC Excel_ChartCheckBoxes (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo, VARIANT index,
                                       CAObjHandle *returnValue);

HRESULT CVIFUNC Excel_ChartCheckSpelling (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo,
                                          VARIANT customDictionary,
                                          VARIANT ignoreUppercase,
                                          VARIANT alwaysSuggest,
                                          VARIANT spellLang);

HRESULT CVIFUNC Excel_ChartColumnGroups (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo, VARIANT index,
                                         CAObjHandle *returnValue);

HRESULT CVIFUNC Excel_ChartCopyPicture (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo,
                                        enum ExcelEnum_XlPictureAppearance appearance,
                                        enum ExcelEnum_XlCopyPictureFormat format,
                                        enum ExcelEnum_XlPictureAppearance size);

HRESULT CVIFUNC Excel_ChartCreatePublisher (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo,
                                            VARIANT edition,
                                            enum ExcelEnum_XlPictureAppearance appearance,
                                            enum ExcelEnum_XlPictureAppearance size,
                                            VARIANT containsPICT,
                                            VARIANT containsBIFF,
                                            VARIANT containsRTF,
                                            VARIANT containsVALU);

HRESULT CVIFUNC Excel_ChartDeselect (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_ChartDoughnutGroups (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo, VARIANT index,
                                           CAObjHandle *returnValue);

HRESULT CVIFUNC Excel_ChartDrawings (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo, VARIANT index,
                                     CAObjHandle *returnValue);

HRESULT CVIFUNC Excel_ChartDrawingObjects (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo, VARIANT index,
                                           CAObjHandle *returnValue);

HRESULT CVIFUNC Excel_ChartDropDowns (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT index,
                                      CAObjHandle *returnValue);

HRESULT CVIFUNC Excel_ChartEvaluate (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo, VARIANT name,
                                     VARIANT *returnValue);

HRESULT CVIFUNC Excel_Chart_Evaluate (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT name,
                                      VARIANT *returnValue);

HRESULT CVIFUNC Excel_ChartGroupBoxes (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo, VARIANT index,
                                       CAObjHandle *returnValue);

HRESULT CVIFUNC Excel_ChartGroupObjects (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo, VARIANT index,
                                         CAObjHandle *returnValue);

HRESULT CVIFUNC Excel_ChartGetHasAxis (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo, VARIANT index1,
                                       VARIANT index2, VARIANT *returnValue);

HRESULT CVIFUNC Excel_ChartSetHasAxis (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo, VARIANT index1,
                                       VARIANT index2, VARIANT newValue);

HRESULT CVIFUNC Excel_ChartLabels (CAObjHandle objectHandle,
                                   ERRORINFO *errorInfo, VARIANT index,
                                   CAObjHandle *returnValue);

HRESULT CVIFUNC Excel_ChartLineGroups (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo, VARIANT index,
                                       CAObjHandle *returnValue);

HRESULT CVIFUNC Excel_ChartLines (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                  VARIANT index, CAObjHandle *returnValue);

HRESULT CVIFUNC Excel_ChartListBoxes (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT index,
                                      CAObjHandle *returnValue);

HRESULT CVIFUNC Excel_ChartLocation (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo,
                                     enum ExcelEnum_XlChartLocation where,
                                     VARIANT name, ExcelObj_Chart *returnValue);

HRESULT CVIFUNC Excel_ChartOLEObjects (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo, VARIANT index,
                                       CAObjHandle *returnValue);

HRESULT CVIFUNC Excel_ChartOptionButtons (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo, VARIANT index,
                                          CAObjHandle *returnValue);

HRESULT CVIFUNC Excel_ChartOvals (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                  VARIANT index, CAObjHandle *returnValue);

HRESULT CVIFUNC Excel_ChartPaste (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                  VARIANT type);

HRESULT CVIFUNC Excel_ChartPictures (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo, VARIANT index,
                                     CAObjHandle *returnValue);

HRESULT CVIFUNC Excel_ChartPieGroups (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT index,
                                      CAObjHandle *returnValue);

HRESULT CVIFUNC Excel_ChartRadarGroups (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo, VARIANT index,
                                        CAObjHandle *returnValue);

HRESULT CVIFUNC Excel_ChartRectangles (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo, VARIANT index,
                                       CAObjHandle *returnValue);

HRESULT CVIFUNC Excel_ChartScrollBars (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo, VARIANT index,
                                       CAObjHandle *returnValue);

HRESULT CVIFUNC Excel_ChartSeriesCollection (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo, VARIANT index,
                                             CAObjHandle *returnValue);

HRESULT CVIFUNC Excel_ChartSpinners (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo, VARIANT index,
                                     CAObjHandle *returnValue);

HRESULT CVIFUNC Excel_ChartTextBoxes (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT index,
                                      CAObjHandle *returnValue);

HRESULT CVIFUNC Excel_ChartApplyCustomType (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo,
                                            enum ExcelEnum_XlChartType chartType,
                                            VARIANT typeName);

HRESULT CVIFUNC Excel_ChartXYGroups (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo, VARIANT index,
                                     CAObjHandle *returnValue);

HRESULT CVIFUNC Excel_ChartCopyChartBuild (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_ChartGetChartElement (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo, long x, long y,
                                            long *elementID, long *arg1,
                                            long *arg2);

HRESULT CVIFUNC Excel_ChartSetSourceData (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo,
                                          ExcelObj_Range source, VARIANT plotBy);

HRESULT CVIFUNC Excel_ChartExport (CAObjHandle objectHandle,
                                   ERRORINFO *errorInfo, const char *filename,
                                   VARIANT filterName, VARIANT interactive,
                                   VBOOL *returnValue);

HRESULT CVIFUNC Excel_ChartRefresh (CAObjHandle objectHandle,
                                    ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_ChartPrintOut (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo, VARIANT from,
                                     VARIANT to, VARIANT copies, VARIANT preview,
                                     VARIANT activePrinter, VARIANT printToFile,
                                     VARIANT collate, VARIANT prToFileName);

HRESULT CVIFUNC Excel_NewWorksheet (const char *server,
                                    int supportMultithreading, LCID locale,
                                    int reserved, CAObjHandle *objectHandle);

HRESULT CVIFUNC Excel_OpenWorksheet (const char *fileName, const char *server,
                                     int supportMultithreading, LCID locale,
                                     int reserved, CAObjHandle *objectHandle);

HRESULT CVIFUNC Excel_ActiveWorksheet (const char *server,
                                       int supportMultithreading, LCID locale,
                                       int reserved, CAObjHandle *objectHandle);

HRESULT CVIFUNC Excel_WorksheetActivate (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_WorksheetCopy (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo, VARIANT before,
                                     VARIANT after);

HRESULT CVIFUNC Excel_WorksheetDelete (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_WorksheetMove (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo, VARIANT before,
                                     VARIANT after);

HRESULT CVIFUNC Excel_Worksheet_PrintOut (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo, VARIANT from,
                                          VARIANT to, VARIANT copies,
                                          VARIANT preview, VARIANT activePrinter,
                                          VARIANT printToFile, VARIANT collate);

HRESULT CVIFUNC Excel_WorksheetPrintPreview (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo,
                                             VARIANT enableChanges);

HRESULT CVIFUNC Excel_WorksheetProtect (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo, VARIANT password,
                                        VARIANT drawingObjects, VARIANT contents,
                                        VARIANT scenarios,
                                        VARIANT userInterfaceOnly);

HRESULT CVIFUNC Excel_WorksheetSaveAs (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo,
                                       const char *filename, VARIANT fileFormat,
                                       VARIANT password,
                                       VARIANT writeResPassword,
                                       VARIANT readOnlyRecommended,
                                       VARIANT createBackup, VARIANT addToMru,
                                       VARIANT textCodepage,
                                       VARIANT textVisualLayout);

HRESULT CVIFUNC Excel_WorksheetSelect (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo, VARIANT replace);

HRESULT CVIFUNC Excel_WorksheetUnprotect (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo, VARIANT password);

HRESULT CVIFUNC Excel_WorksheetArcs (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo, VARIANT index,
                                     CAObjHandle *returnValue);

HRESULT CVIFUNC Excel_WorksheetSetBackgroundPicture (CAObjHandle objectHandle,
                                                     ERRORINFO *errorInfo,
                                                     const char *filename);

HRESULT CVIFUNC Excel_WorksheetButtons (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo, VARIANT index,
                                        CAObjHandle *returnValue);

HRESULT CVIFUNC Excel_WorksheetCalculate (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_WorksheetChartObjects (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo, VARIANT index,
                                             CAObjHandle *returnValue);

HRESULT CVIFUNC Excel_WorksheetCheckBoxes (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo, VARIANT index,
                                           CAObjHandle *returnValue);

HRESULT CVIFUNC Excel_WorksheetCheckSpelling (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo,
                                              VARIANT customDictionary,
                                              VARIANT ignoreUppercase,
                                              VARIANT alwaysSuggest,
                                              VARIANT spellLang);

HRESULT CVIFUNC Excel_WorksheetClearArrows (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_WorksheetDrawings (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo, VARIANT index,
                                         CAObjHandle *returnValue);

HRESULT CVIFUNC Excel_WorksheetDrawingObjects (CAObjHandle objectHandle,
                                               ERRORINFO *errorInfo,
                                               VARIANT index,
                                               CAObjHandle *returnValue);

HRESULT CVIFUNC Excel_WorksheetDropDowns (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo, VARIANT index,
                                          CAObjHandle *returnValue);

HRESULT CVIFUNC Excel_WorksheetEvaluate (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo, VARIANT name,
                                         VARIANT *returnValue);

HRESULT CVIFUNC Excel_Worksheet_Evaluate (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo, VARIANT name,
                                          VARIANT *returnValue);

HRESULT CVIFUNC Excel_WorksheetResetAllPageBreaks (CAObjHandle objectHandle,
                                                   ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_WorksheetGroupBoxes (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo, VARIANT index,
                                           CAObjHandle *returnValue);

HRESULT CVIFUNC Excel_WorksheetGroupObjects (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo, VARIANT index,
                                             CAObjHandle *returnValue);

HRESULT CVIFUNC Excel_WorksheetLabels (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo, VARIANT index,
                                       CAObjHandle *returnValue);

HRESULT CVIFUNC Excel_WorksheetLines (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT index,
                                      CAObjHandle *returnValue);

HRESULT CVIFUNC Excel_WorksheetListBoxes (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo, VARIANT index,
                                          CAObjHandle *returnValue);

HRESULT CVIFUNC Excel_WorksheetOLEObjects (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo, VARIANT index,
                                           CAObjHandle *returnValue);

HRESULT CVIFUNC Excel_WorksheetOptionButtons (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo,
                                              VARIANT index,
                                              CAObjHandle *returnValue);

HRESULT CVIFUNC Excel_WorksheetOvals (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT index,
                                      CAObjHandle *returnValue);

HRESULT CVIFUNC Excel_WorksheetPaste (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT destination,
                                      VARIANT link);

HRESULT CVIFUNC Excel_WorksheetPasteSpecial (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo,
                                             VARIANT format, VARIANT link,
                                             VARIANT displayAsIcon,
                                             VARIANT iconFileName,
                                             VARIANT iconIndex,
                                             VARIANT iconLabel);

HRESULT CVIFUNC Excel_WorksheetPictures (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo, VARIANT index,
                                         CAObjHandle *returnValue);

HRESULT CVIFUNC Excel_WorksheetPivotTables (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo, VARIANT index,
                                            CAObjHandle *returnValue);

HRESULT CVIFUNC Excel_WorksheetPivotTableWizard (CAObjHandle objectHandle,
                                                 ERRORINFO *errorInfo,
                                                 VARIANT sourceType,
                                                 VARIANT sourceData,
                                                 VARIANT tableDestination,
                                                 VARIANT tableName,
                                                 VARIANT rowGrand,
                                                 VARIANT columnGrand,
                                                 VARIANT saveData,
                                                 VARIANT hasAutoFormat,
                                                 VARIANT autoPage,
                                                 VARIANT reserved,
                                                 VARIANT backgroundQuery,
                                                 VARIANT optimizeCache,
                                                 VARIANT pageFieldOrder,
                                                 VARIANT pageFieldWrapCount,
                                                 VARIANT readData,
                                                 VARIANT connection,
                                                 ExcelObj_PivotTable *returnValue);

HRESULT CVIFUNC Excel_WorksheetRange (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT cell1,
                                      VARIANT cell2, ExcelObj_Range *returnValue);

HRESULT CVIFUNC Excel_WorksheetRectangles (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo, VARIANT index,
                                           CAObjHandle *returnValue);

HRESULT CVIFUNC Excel_WorksheetScenarios (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo, VARIANT index,
                                          CAObjHandle *returnValue);

HRESULT CVIFUNC Excel_WorksheetScrollBars (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo, VARIANT index,
                                           CAObjHandle *returnValue);

HRESULT CVIFUNC Excel_WorksheetShowAllData (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_WorksheetShowDataForm (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_WorksheetSpinners (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo, VARIANT index,
                                         CAObjHandle *returnValue);

HRESULT CVIFUNC Excel_WorksheetTextBoxes (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo, VARIANT index,
                                          CAObjHandle *returnValue);

HRESULT CVIFUNC Excel_WorksheetClearCircles (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_WorksheetCircleInvalid (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_WorksheetPrintOut (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo, VARIANT from,
                                         VARIANT to, VARIANT copies,
                                         VARIANT preview, VARIANT activePrinter,
                                         VARIANT printToFile, VARIANT collate,
                                         VARIANT prToFileName);

HRESULT CVIFUNC Excel_Worksheet_CheckSpelling (CAObjHandle objectHandle,
                                               ERRORINFO *errorInfo,
                                               VARIANT customDictionary,
                                               VARIANT ignoreUppercase,
                                               VARIANT alwaysSuggest,
                                               VARIANT spellLang,
                                               VARIANT ignoreFinalYaa,
                                               VARIANT spellScript);

HRESULT CVIFUNC Excel_NewWorkbook (const char *server, int supportMultithreading,
                                   LCID locale, int reserved,
                                   CAObjHandle *objectHandle);

HRESULT CVIFUNC Excel_OpenWorkbook (const char *fileName, const char *server,
                                    int supportMultithreading, LCID locale,
                                    int reserved, CAObjHandle *objectHandle);

HRESULT CVIFUNC Excel_ActiveWorkbook (const char *server,
                                      int supportMultithreading, LCID locale,
                                      int reserved, CAObjHandle *objectHandle);

HRESULT CVIFUNC Excel_WorkbookActivate (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_WorkbookChangeFileAccess (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo,
                                                enum ExcelEnum_XlFileAccess mode,
                                                VARIANT writePassword,
                                                VARIANT notify);

HRESULT CVIFUNC Excel_WorkbookChangeLink (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo, const char *name,
                                          const char *newName,
                                          enum ExcelEnum_XlLinkType type);

HRESULT CVIFUNC Excel_WorkbookClose (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo, VARIANT saveChanges,
                                     VARIANT filename, VARIANT routeWorkbook);

HRESULT CVIFUNC Excel_WorkbookGetColors (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo, VARIANT index,
                                         VARIANT *returnValue);

HRESULT CVIFUNC Excel_WorkbookSetColors (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo, VARIANT index,
                                         VARIANT newValue);

HRESULT CVIFUNC Excel_WorkbookDeleteNumberFormat (CAObjHandle objectHandle,
                                                  ERRORINFO *errorInfo,
                                                  const char *numberFormat);

HRESULT CVIFUNC Excel_WorkbookExclusiveAccess (CAObjHandle objectHandle,
                                               ERRORINFO *errorInfo,
                                               VBOOL *returnValue);

HRESULT CVIFUNC Excel_WorkbookForwardMailer (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_WorkbookLinkInfo (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo, const char *name,
                                        enum ExcelEnum_XlLinkInfo linkInfo,
                                        VARIANT type, VARIANT editionRef,
                                        VARIANT *returnValue);

HRESULT CVIFUNC Excel_WorkbookLinkSources (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo, VARIANT type,
                                           VARIANT *returnValue);

HRESULT CVIFUNC Excel_WorkbookMergeWorkbook (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo,
                                             VARIANT filename);

HRESULT CVIFUNC Excel_WorkbookNewWindow (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo,
                                         ExcelObj_Window *returnValue);

HRESULT CVIFUNC Excel_WorkbookOpenLinks (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo, const char *name,
                                         VARIANT readOnly, VARIANT type);

HRESULT CVIFUNC Excel_WorkbookPivotCaches (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo,
                                           ExcelObj_PivotCaches *returnValue);

HRESULT CVIFUNC Excel_WorkbookPost (CAObjHandle objectHandle,
                                    ERRORINFO *errorInfo, VARIANT destName);

HRESULT CVIFUNC Excel_Workbook_PrintOut (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo, VARIANT from,
                                         VARIANT to, VARIANT copies,
                                         VARIANT preview, VARIANT activePrinter,
                                         VARIANT printToFile, VARIANT collate);

HRESULT CVIFUNC Excel_WorkbookPrintPreview (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo,
                                            VARIANT enableChanges);

HRESULT CVIFUNC Excel_WorkbookProtect (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo, VARIANT password,
                                       VARIANT structure, VARIANT windows);

HRESULT CVIFUNC Excel_WorkbookProtectSharing (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo,
                                              VARIANT filename, VARIANT password,
                                              VARIANT writeResPassword,
                                              VARIANT readOnlyRecommended,
                                              VARIANT createBackup,
                                              VARIANT sharingPassword);

HRESULT CVIFUNC Excel_WorkbookRefreshAll (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_WorkbookReply (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_WorkbookReplyAll (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_WorkbookRemoveUser (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo, long index);

HRESULT CVIFUNC Excel_WorkbookRoute (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_WorkbookRunAutoMacros (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo,
                                             enum ExcelEnum_XlRunAutoMacro which);

HRESULT CVIFUNC Excel_WorkbookSave (CAObjHandle objectHandle,
                                    ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_WorkbookSaveAs (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT filename,
                                      VARIANT fileFormat, VARIANT password,
                                      VARIANT writeResPassword,
                                      VARIANT readOnlyRecommended,
                                      VARIANT createBackup,
                                      enum ExcelEnum_XlSaveAsAccessMode accessMode,
                                      VARIANT conflictResolution,
                                      VARIANT addToMru, VARIANT textCodepage,
                                      VARIANT textVisualLayout);

HRESULT CVIFUNC Excel_WorkbookSaveCopyAs (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo, VARIANT filename);

HRESULT CVIFUNC Excel_WorkbookSendMail (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo, VARIANT recipients,
                                        VARIANT subject, VARIANT returnReceipt);

HRESULT CVIFUNC Excel_WorkbookSendMailer (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo,
                                          VARIANT fileFormat,
                                          enum ExcelEnum_XlPriority priority);

HRESULT CVIFUNC Excel_WorkbookSetLinkOnData (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo,
                                             const char *name, VARIANT procedure);

HRESULT CVIFUNC Excel_WorkbookUnprotect (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo, VARIANT password);

HRESULT CVIFUNC Excel_WorkbookUnprotectSharing (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo,
                                                VARIANT sharingPassword);

HRESULT CVIFUNC Excel_WorkbookUpdateFromFile (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_WorkbookUpdateLink (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo, VARIANT name,
                                          VARIANT type);

HRESULT CVIFUNC Excel_WorkbookHighlightChangesOptions (CAObjHandle objectHandle,
                                                       ERRORINFO *errorInfo,
                                                       VARIANT when, VARIANT who,
                                                       VARIANT where);

HRESULT CVIFUNC Excel_WorkbookPurgeChangeHistoryNow (CAObjHandle objectHandle,
                                                     ERRORINFO *errorInfo,
                                                     long days,
                                                     VARIANT sharingPassword);

HRESULT CVIFUNC Excel_WorkbookAcceptAllChanges (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo,
                                                VARIANT when, VARIANT who,
                                                VARIANT where);

HRESULT CVIFUNC Excel_WorkbookRejectAllChanges (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo,
                                                VARIANT when, VARIANT who,
                                                VARIANT where);

HRESULT CVIFUNC Excel_WorkbookPivotTableWizard (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo,
                                                VARIANT sourceType,
                                                VARIANT sourceData,
                                                VARIANT tableDestination,
                                                VARIANT tableName,
                                                VARIANT rowGrand,
                                                VARIANT columnGrand,
                                                VARIANT saveData,
                                                VARIANT hasAutoFormat,
                                                VARIANT autoPage,
                                                VARIANT reserved,
                                                VARIANT backgroundQuery,
                                                VARIANT optimizeCache,
                                                VARIANT pageFieldOrder,
                                                VARIANT pageFieldWrapCount,
                                                VARIANT readData,
                                                VARIANT connection);

HRESULT CVIFUNC Excel_WorkbookResetColors (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_WorkbookFollowHyperlink (CAObjHandle objectHandle,
                                               ERRORINFO *errorInfo,
                                               const char *address,
                                               VARIANT subAddress,
                                               VARIANT newWindow,
                                               VARIANT addHistory,
                                               VARIANT extraInfo, VARIANT method,
                                               VARIANT headerInfo);

HRESULT CVIFUNC Excel_WorkbookAddToFavorites (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_WorkbookPrintOut (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo, VARIANT from,
                                        VARIANT to, VARIANT copies,
                                        VARIANT preview, VARIANT activePrinter,
                                        VARIANT printToFile, VARIANT collate,
                                        VARIANT prToFileName);

HRESULT CVIFUNC Excel_WorkbookWebPagePreview (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_WorkbookReloadAs (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo,
                                        enum ExcelEnum_MsoEncoding encoding);

HRESULT CVIFUNC Excel_WorkbookDummy1 (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, long unused);

HRESULT CVIFUNC Excel_Workbooksblt (CAObjHandle objectHandle,
                                    ERRORINFO *errorInfo, const char *s);

HRESULT CVIFUNC Excel_NewOLEObject (const char *server,
                                    int supportMultithreading, LCID locale,
                                    int reserved, CAObjHandle *objectHandle);

HRESULT CVIFUNC Excel_OpenOLEObject (const char *fileName, const char *server,
                                     int supportMultithreading, LCID locale,
                                     int reserved, CAObjHandle *objectHandle);

HRESULT CVIFUNC Excel_ActiveOLEObject (const char *server,
                                       int supportMultithreading, LCID locale,
                                       int reserved, CAObjHandle *objectHandle);

HRESULT CVIFUNC Excel_OLEObjectBringToFront (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo,
                                             VARIANT *returnValue);

HRESULT CVIFUNC Excel_OLEObjectCopy (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo, VARIANT *returnValue);

HRESULT CVIFUNC Excel_OLEObjectCopyPicture (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo,
                                            enum ExcelEnum_XlPictureAppearance appearance,
                                            enum ExcelEnum_XlCopyPictureFormat format,
                                            VARIANT *returnValue);

HRESULT CVIFUNC Excel_OLEObjectCut (CAObjHandle objectHandle,
                                    ERRORINFO *errorInfo, VARIANT *returnValue);

HRESULT CVIFUNC Excel_OLEObjectDelete (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo,
                                       VARIANT *returnValue);

HRESULT CVIFUNC Excel_OLEObjectDuplicate (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo,
                                          CAObjHandle *returnValue);

HRESULT CVIFUNC Excel_OLEObjectSelect (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo, VARIANT replace,
                                       VARIANT *returnValue);

HRESULT CVIFUNC Excel_OLEObjectSendToBack (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo,
                                           VARIANT *returnValue);

HRESULT CVIFUNC Excel_OLEObjectActivate (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo,
                                         VARIANT *returnValue);

HRESULT CVIFUNC Excel_OLEObjectUpdate (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo,
                                       VARIANT *returnValue);

HRESULT CVIFUNC Excel_OLEObjectVerb (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo,
                                     enum ExcelEnum_XlOLEVerb verb,
                                     VARIANT *returnValue);

HRESULT CVIFUNC Excel_AdjustmentsGetItem (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo, long index,
                                          float *returnValue);

HRESULT CVIFUNC Excel_AdjustmentsSetItem (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo, long index,
                                          float newValue);

HRESULT CVIFUNC Excel_CalloutFormatAutomaticLength (CAObjHandle objectHandle,
                                                    ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_CalloutFormatCustomDrop (CAObjHandle objectHandle,
                                               ERRORINFO *errorInfo, float drop);

HRESULT CVIFUNC Excel_CalloutFormatCustomLength (CAObjHandle objectHandle,
                                                 ERRORINFO *errorInfo,
                                                 float length);

HRESULT CVIFUNC Excel_CalloutFormatPresetDrop (CAObjHandle objectHandle,
                                               ERRORINFO *errorInfo,
                                               enum ExcelEnum_MsoCalloutDropType dropType);

HRESULT CVIFUNC Excel_ShapeNodesItem (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT index,
                                      ExcelObj_ShapeNode *returnValue);

HRESULT CVIFUNC Excel_ShapeNodesDelete (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo, long index);

HRESULT CVIFUNC Excel_ShapeNodesInsert (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo, long index,
                                        enum ExcelEnum_MsoSegmentType segmentType,
                                        enum ExcelEnum_MsoEditingType editingType,
                                        float x1, float y1, float x2, float y2,
                                        float x3, float y3);

HRESULT CVIFUNC Excel_ShapeNodesSetEditingType (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo, long index,
                                                enum ExcelEnum_MsoEditingType editingType);

HRESULT CVIFUNC Excel_ShapeNodesSetPosition (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo, long index,
                                             float x1, float y1);

HRESULT CVIFUNC Excel_ShapeNodesSetSegmentType (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo, long index,
                                                enum ExcelEnum_MsoSegmentType segmentType);

HRESULT CVIFUNC Excel_PictureFmtIncrementBrightness (CAObjHandle objectHandle,
                                                     ERRORINFO *errorInfo,
                                                     float increment);

HRESULT CVIFUNC Excel_PictureFmtIncrementContrast (CAObjHandle objectHandle,
                                                   ERRORINFO *errorInfo,
                                                   float increment);

HRESULT CVIFUNC Excel_ShadowFormatIncrementOffsetX (CAObjHandle objectHandle,
                                                    ERRORINFO *errorInfo,
                                                    float increment);

HRESULT CVIFUNC Excel_ShadowFormatIncrementOffsetY (CAObjHandle objectHandle,
                                                    ERRORINFO *errorInfo,
                                                    float increment);

HRESULT CVIFUNC Excel_TextEffectFmtToggleVerticalText (CAObjHandle objectHandle,
                                                       ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_ThreeDFmtIncrementRotationX (CAObjHandle objectHandle,
                                                   ERRORINFO *errorInfo,
                                                   float increment);

HRESULT CVIFUNC Excel_ThreeDFmtIncrementRotationY (CAObjHandle objectHandle,
                                                   ERRORINFO *errorInfo,
                                                   float increment);

HRESULT CVIFUNC Excel_ThreeDFmtResetRotation (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_ThreeDFmtSetThreeDFormat (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo,
                                                enum ExcelEnum_MsoPresetThreeDFormat presetThreeDFormat);

HRESULT CVIFUNC Excel_ThreeDFmtSetExtrusionDirection (CAObjHandle objectHandle,
                                                      ERRORINFO *errorInfo,
                                                      enum ExcelEnum_MsoPresetExtrusionDirection presetExtrusionDirection);

HRESULT CVIFUNC Excel_FillFormatBackground (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_FillFormatOneColorGradient (CAObjHandle objectHandle,
                                                  ERRORINFO *errorInfo,
                                                  enum ExcelEnum_MsoGradientStyle style,
                                                  long variant, float degree);

HRESULT CVIFUNC Excel_FillFormatPatterned (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo,
                                           enum ExcelEnum_MsoPatternType pattern);

HRESULT CVIFUNC Excel_FillFormatPresetGradient (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo,
                                                enum ExcelEnum_MsoGradientStyle style,
                                                long variant,
                                                enum ExcelEnum_MsoPresetGradientType presetGradientType);

HRESULT CVIFUNC Excel_FillFormatPresetTextured (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo,
                                                enum ExcelEnum_MsoPresetTexture presetTexture);

HRESULT CVIFUNC Excel_FillFormatSolid (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_FillFormatTwoColorGradient (CAObjHandle objectHandle,
                                                  ERRORINFO *errorInfo,
                                                  enum ExcelEnum_MsoGradientStyle style,
                                                  long variant);

HRESULT CVIFUNC Excel_FillFormatUserPicture (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo,
                                             const char *pictureFile);

HRESULT CVIFUNC Excel_FillFormatUserTextured (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo,
                                              const char *textureFile);

HRESULT CVIFUNC Excel_SheetsAdd (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                 VARIANT before, VARIANT after, VARIANT count,
                                 VARIANT type, CAObjHandle *returnValue);

HRESULT CVIFUNC Excel_SheetsCopy (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                  VARIANT before, VARIANT after);

HRESULT CVIFUNC Excel_SheetsDelete (CAObjHandle objectHandle,
                                    ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_SheetsFillAcrossSheets (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo,
                                              ExcelObj_Range range,
                                              enum ExcelEnum_XlFillWith type);

HRESULT CVIFUNC Excel_SheetsItem (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                  VARIANT index, CAObjHandle *returnValue);

HRESULT CVIFUNC Excel_SheetsMove (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                  VARIANT before, VARIANT after);

HRESULT CVIFUNC Excel_Sheets_PrintOut (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo, VARIANT from,
                                       VARIANT to, VARIANT copies,
                                       VARIANT preview, VARIANT activePrinter,
                                       VARIANT printToFile, VARIANT collate);

HRESULT CVIFUNC Excel_SheetsPrintPreview (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo,
                                          VARIANT enableChanges);

HRESULT CVIFUNC Excel_SheetsSelect (CAObjHandle objectHandle,
                                    ERRORINFO *errorInfo, VARIANT replace);

HRESULT CVIFUNC Excel_Sheets_Default (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT index,
                                      CAObjHandle *returnValue);

HRESULT CVIFUNC Excel_SheetsPrintOut (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT from,
                                      VARIANT to, VARIANT copies,
                                      VARIANT preview, VARIANT activePrinter,
                                      VARIANT printToFile, VARIANT collate,
                                      VARIANT prToFileName);

HRESULT CVIFUNC Excel_WorkbooksAdd (CAObjHandle objectHandle,
                                    ERRORINFO *errorInfo, VARIANT template,
                                    ExcelObj_Workbook *returnValue);

HRESULT CVIFUNC Excel_WorkbooksClose (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_WorkbooksItem (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo, VARIANT index,
                                     ExcelObj_Workbook *returnValue);

HRESULT CVIFUNC Excel_WorkbooksOpen (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo, const char *filename,
                                     VARIANT updateLinks, VARIANT readOnly,
                                     VARIANT format, VARIANT password,
                                     VARIANT writeResPassword,
                                     VARIANT ignoreReadOnlyRecommended,
                                     VARIANT origin, VARIANT delimiter,
                                     VARIANT editable, VARIANT notify,
                                     VARIANT converter, VARIANT addToMru,
                                     ExcelObj_Workbook *returnValue);

HRESULT CVIFUNC Excel_Workbooks_OpenText (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo,
                                          const char *filename, VARIANT origin,
                                          VARIANT startRow, VARIANT dataType,
                                          enum ExcelEnum_XlTextQualifier textQualifier,
                                          VARIANT consecutiveDelimiter,
                                          VARIANT tab, VARIANT semicolon,
                                          VARIANT comma, VARIANT space,
                                          VARIANT other, VARIANT otherChar,
                                          VARIANT fieldInfo,
                                          VARIANT textVisualLayout);

HRESULT CVIFUNC Excel_Workbooks_Default (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo, VARIANT index,
                                         ExcelObj_Workbook *returnValue);

HRESULT CVIFUNC Excel_WorkbooksOpenText (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo,
                                         const char *filename, VARIANT origin,
                                         VARIANT startRow, VARIANT dataType,
                                         enum ExcelEnum_XlTextQualifier textQualifier,
                                         VARIANT consecutiveDelimiter,
                                         VARIANT tab, VARIANT semicolon,
                                         VARIANT comma, VARIANT space,
                                         VARIANT other, VARIANT otherChar,
                                         VARIANT fieldInfo,
                                         VARIANT textVisualLayout,
                                         VARIANT decimalSeparator,
                                         VARIANT thousandsSeparator);

HRESULT CVIFUNC Excel_PublishObjectDelete (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_PublishObjectPublish (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo, VARIANT create);

HRESULT CVIFUNC Excel_WebOptionsUseDefaultFolderSuffix (CAObjHandle objectHandle,
                                                        ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_CubeFieldsItem (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT index,
                                      ExcelObj_CubeField *returnValue);

HRESULT CVIFUNC Excel_CubeFields_Default (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo, VARIANT index,
                                          ExcelObj_CubeField *returnValue);

HRESULT CVIFUNC Excel_WindowActivate (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT *returnValue);

HRESULT CVIFUNC Excel_WindowActivateNext (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo,
                                          VARIANT *returnValue);

HRESULT CVIFUNC Excel_WindowActivatePrevious (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo,
                                              VARIANT *returnValue);

HRESULT CVIFUNC Excel_WindowClose (CAObjHandle objectHandle,
                                   ERRORINFO *errorInfo, VARIANT saveChanges,
                                   VARIANT filename, VARIANT routeWorkbook,
                                   VBOOL *returnValue);

HRESULT CVIFUNC Excel_WindowLargeScroll (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo, VARIANT down,
                                         VARIANT up, VARIANT toRight,
                                         VARIANT toLeft, VARIANT *returnValue);

HRESULT CVIFUNC Excel_WindowNewWindow (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo,
                                       ExcelObj_Window *returnValue);

HRESULT CVIFUNC Excel_WindowPrintOut (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT from,
                                      VARIANT to, VARIANT copies,
                                      VARIANT preview, VARIANT activePrinter,
                                      VARIANT printToFile, VARIANT collate,
                                      VARIANT prToFileName, VARIANT *returnValue);

HRESULT CVIFUNC Excel_WindowPrintPreview (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo,
                                          VARIANT enableChanges,
                                          VARIANT *returnValue);

HRESULT CVIFUNC Excel_WindowScrollWorkbookTabs (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo,
                                                VARIANT sheets, VARIANT position,
                                                VARIANT *returnValue);

HRESULT CVIFUNC Excel_WindowSmallScroll (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo, VARIANT down,
                                         VARIANT up, VARIANT toRight,
                                         VARIANT toLeft, VARIANT *returnValue);

HRESULT CVIFUNC Excel_WindowPointsToScreenPixelsX (CAObjHandle objectHandle,
                                                   ERRORINFO *errorInfo,
                                                   long points,
                                                   long *returnValue);

HRESULT CVIFUNC Excel_WindowPointsToScreenPixelsY (CAObjHandle objectHandle,
                                                   ERRORINFO *errorInfo,
                                                   long points,
                                                   long *returnValue);

HRESULT CVIFUNC Excel_WindowRangeFromPoint (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo, long x, long y,
                                            CAObjHandle *returnValue);

HRESULT CVIFUNC Excel_WindowScrollIntoView (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo, long left,
                                            long top, long width, long height,
                                            VARIANT start);

HRESULT CVIFUNC Excel_WindowsArrange (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo,
                                      enum ExcelEnum_XlArrangeStyle arrangeStyle,
                                      VARIANT activeWorkbook,
                                      VARIANT syncHorizontal,
                                      VARIANT syncVertical, VARIANT *returnValue);

HRESULT CVIFUNC Excel_WindowsItem (CAObjHandle objectHandle,
                                   ERRORINFO *errorInfo, VARIANT index,
                                   ExcelObj_Window *returnValue);

HRESULT CVIFUNC Excel_Windows_Default (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo, VARIANT index,
                                       ExcelObj_Window *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunction_WSFunction (CAObjHandle objectHandle,
                                                    ERRORINFO *errorInfo,
                                                    VARIANT arg1, VARIANT arg2,
                                                    VARIANT arg3, VARIANT arg4,
                                                    VARIANT arg5, VARIANT arg6,
                                                    VARIANT arg7, VARIANT arg8,
                                                    VARIANT arg9, VARIANT arg10,
                                                    VARIANT arg11, VARIANT arg12,
                                                    VARIANT arg13, VARIANT arg14,
                                                    VARIANT arg15, VARIANT arg16,
                                                    VARIANT arg17, VARIANT arg18,
                                                    VARIANT arg19, VARIANT arg20,
                                                    VARIANT arg21, VARIANT arg22,
                                                    VARIANT arg23, VARIANT arg24,
                                                    VARIANT arg25, VARIANT arg26,
                                                    VARIANT arg27, VARIANT arg28,
                                                    VARIANT arg29, VARIANT arg30,
                                                    VARIANT *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionCount (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo, VARIANT arg1,
                                              VARIANT arg2, VARIANT arg3,
                                              VARIANT arg4, VARIANT arg5,
                                              VARIANT arg6, VARIANT arg7,
                                              VARIANT arg8, VARIANT arg9,
                                              VARIANT arg10, VARIANT arg11,
                                              VARIANT arg12, VARIANT arg13,
                                              VARIANT arg14, VARIANT arg15,
                                              VARIANT arg16, VARIANT arg17,
                                              VARIANT arg18, VARIANT arg19,
                                              VARIANT arg20, VARIANT arg21,
                                              VARIANT arg22, VARIANT arg23,
                                              VARIANT arg24, VARIANT arg25,
                                              VARIANT arg26, VARIANT arg27,
                                              VARIANT arg28, VARIANT arg29,
                                              VARIANT arg30, double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionIsNA (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo, VARIANT arg1,
                                             VBOOL *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionIsError (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo,
                                                VARIANT arg1, VBOOL *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionSum (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo, VARIANT arg1,
                                            VARIANT arg2, VARIANT arg3,
                                            VARIANT arg4, VARIANT arg5,
                                            VARIANT arg6, VARIANT arg7,
                                            VARIANT arg8, VARIANT arg9,
                                            VARIANT arg10, VARIANT arg11,
                                            VARIANT arg12, VARIANT arg13,
                                            VARIANT arg14, VARIANT arg15,
                                            VARIANT arg16, VARIANT arg17,
                                            VARIANT arg18, VARIANT arg19,
                                            VARIANT arg20, VARIANT arg21,
                                            VARIANT arg22, VARIANT arg23,
                                            VARIANT arg24, VARIANT arg25,
                                            VARIANT arg26, VARIANT arg27,
                                            VARIANT arg28, VARIANT arg29,
                                            VARIANT arg30, double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionAverage (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo,
                                                VARIANT arg1, VARIANT arg2,
                                                VARIANT arg3, VARIANT arg4,
                                                VARIANT arg5, VARIANT arg6,
                                                VARIANT arg7, VARIANT arg8,
                                                VARIANT arg9, VARIANT arg10,
                                                VARIANT arg11, VARIANT arg12,
                                                VARIANT arg13, VARIANT arg14,
                                                VARIANT arg15, VARIANT arg16,
                                                VARIANT arg17, VARIANT arg18,
                                                VARIANT arg19, VARIANT arg20,
                                                VARIANT arg21, VARIANT arg22,
                                                VARIANT arg23, VARIANT arg24,
                                                VARIANT arg25, VARIANT arg26,
                                                VARIANT arg27, VARIANT arg28,
                                                VARIANT arg29, VARIANT arg30,
                                                double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionMin (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo, VARIANT arg1,
                                            VARIANT arg2, VARIANT arg3,
                                            VARIANT arg4, VARIANT arg5,
                                            VARIANT arg6, VARIANT arg7,
                                            VARIANT arg8, VARIANT arg9,
                                            VARIANT arg10, VARIANT arg11,
                                            VARIANT arg12, VARIANT arg13,
                                            VARIANT arg14, VARIANT arg15,
                                            VARIANT arg16, VARIANT arg17,
                                            VARIANT arg18, VARIANT arg19,
                                            VARIANT arg20, VARIANT arg21,
                                            VARIANT arg22, VARIANT arg23,
                                            VARIANT arg24, VARIANT arg25,
                                            VARIANT arg26, VARIANT arg27,
                                            VARIANT arg28, VARIANT arg29,
                                            VARIANT arg30, double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionMax (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo, VARIANT arg1,
                                            VARIANT arg2, VARIANT arg3,
                                            VARIANT arg4, VARIANT arg5,
                                            VARIANT arg6, VARIANT arg7,
                                            VARIANT arg8, VARIANT arg9,
                                            VARIANT arg10, VARIANT arg11,
                                            VARIANT arg12, VARIANT arg13,
                                            VARIANT arg14, VARIANT arg15,
                                            VARIANT arg16, VARIANT arg17,
                                            VARIANT arg18, VARIANT arg19,
                                            VARIANT arg20, VARIANT arg21,
                                            VARIANT arg22, VARIANT arg23,
                                            VARIANT arg24, VARIANT arg25,
                                            VARIANT arg26, VARIANT arg27,
                                            VARIANT arg28, VARIANT arg29,
                                            VARIANT arg30, double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionNpv (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo, double arg1,
                                            VARIANT arg2, VARIANT arg3,
                                            VARIANT arg4, VARIANT arg5,
                                            VARIANT arg6, VARIANT arg7,
                                            VARIANT arg8, VARIANT arg9,
                                            VARIANT arg10, VARIANT arg11,
                                            VARIANT arg12, VARIANT arg13,
                                            VARIANT arg14, VARIANT arg15,
                                            VARIANT arg16, VARIANT arg17,
                                            VARIANT arg18, VARIANT arg19,
                                            VARIANT arg20, VARIANT arg21,
                                            VARIANT arg22, VARIANT arg23,
                                            VARIANT arg24, VARIANT arg25,
                                            VARIANT arg26, VARIANT arg27,
                                            VARIANT arg28, VARIANT arg29,
                                            VARIANT arg30, double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionStDev (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo, VARIANT arg1,
                                              VARIANT arg2, VARIANT arg3,
                                              VARIANT arg4, VARIANT arg5,
                                              VARIANT arg6, VARIANT arg7,
                                              VARIANT arg8, VARIANT arg9,
                                              VARIANT arg10, VARIANT arg11,
                                              VARIANT arg12, VARIANT arg13,
                                              VARIANT arg14, VARIANT arg15,
                                              VARIANT arg16, VARIANT arg17,
                                              VARIANT arg18, VARIANT arg19,
                                              VARIANT arg20, VARIANT arg21,
                                              VARIANT arg22, VARIANT arg23,
                                              VARIANT arg24, VARIANT arg25,
                                              VARIANT arg26, VARIANT arg27,
                                              VARIANT arg28, VARIANT arg29,
                                              VARIANT arg30, double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionDollar (CAObjHandle objectHandle,
                                               ERRORINFO *errorInfo, double arg1,
                                               VARIANT arg2, char **returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionFixed (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo, double arg1,
                                              VARIANT arg2, VARIANT arg3,
                                              char **returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionPi (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo,
                                           double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionLn (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo, double arg1,
                                           double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionLog10 (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo, double arg1,
                                              double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionRound (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo, double arg1,
                                              double arg2, double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionLookup (CAObjHandle objectHandle,
                                               ERRORINFO *errorInfo,
                                               VARIANT arg1, VARIANT arg2,
                                               VARIANT arg3,
                                               VARIANT *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionIndex (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo, VARIANT arg1,
                                              double arg2, VARIANT arg3,
                                              VARIANT arg4, VARIANT *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionRept (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo,
                                             const char *arg1, double arg2,
                                             char **returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionAnd (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo, VARIANT arg1,
                                            VARIANT arg2, VARIANT arg3,
                                            VARIANT arg4, VARIANT arg5,
                                            VARIANT arg6, VARIANT arg7,
                                            VARIANT arg8, VARIANT arg9,
                                            VARIANT arg10, VARIANT arg11,
                                            VARIANT arg12, VARIANT arg13,
                                            VARIANT arg14, VARIANT arg15,
                                            VARIANT arg16, VARIANT arg17,
                                            VARIANT arg18, VARIANT arg19,
                                            VARIANT arg20, VARIANT arg21,
                                            VARIANT arg22, VARIANT arg23,
                                            VARIANT arg24, VARIANT arg25,
                                            VARIANT arg26, VARIANT arg27,
                                            VARIANT arg28, VARIANT arg29,
                                            VARIANT arg30, VBOOL *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionOr (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo, VARIANT arg1,
                                           VARIANT arg2, VARIANT arg3,
                                           VARIANT arg4, VARIANT arg5,
                                           VARIANT arg6, VARIANT arg7,
                                           VARIANT arg8, VARIANT arg9,
                                           VARIANT arg10, VARIANT arg11,
                                           VARIANT arg12, VARIANT arg13,
                                           VARIANT arg14, VARIANT arg15,
                                           VARIANT arg16, VARIANT arg17,
                                           VARIANT arg18, VARIANT arg19,
                                           VARIANT arg20, VARIANT arg21,
                                           VARIANT arg22, VARIANT arg23,
                                           VARIANT arg24, VARIANT arg25,
                                           VARIANT arg26, VARIANT arg27,
                                           VARIANT arg28, VARIANT arg29,
                                           VARIANT arg30, VBOOL *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionDCount (CAObjHandle objectHandle,
                                               ERRORINFO *errorInfo,
                                               ExcelObj_Range arg1, VARIANT arg2,
                                               VARIANT arg3, double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionDSum (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo,
                                             ExcelObj_Range arg1, VARIANT arg2,
                                             VARIANT arg3, double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionDAverage (CAObjHandle objectHandle,
                                                 ERRORINFO *errorInfo,
                                                 ExcelObj_Range arg1,
                                                 VARIANT arg2, VARIANT arg3,
                                                 double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionDMin (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo,
                                             ExcelObj_Range arg1, VARIANT arg2,
                                             VARIANT arg3, double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionDMax (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo,
                                             ExcelObj_Range arg1, VARIANT arg2,
                                             VARIANT arg3, double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionDStDev (CAObjHandle objectHandle,
                                               ERRORINFO *errorInfo,
                                               ExcelObj_Range arg1, VARIANT arg2,
                                               VARIANT arg3, double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionVar (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo, VARIANT arg1,
                                            VARIANT arg2, VARIANT arg3,
                                            VARIANT arg4, VARIANT arg5,
                                            VARIANT arg6, VARIANT arg7,
                                            VARIANT arg8, VARIANT arg9,
                                            VARIANT arg10, VARIANT arg11,
                                            VARIANT arg12, VARIANT arg13,
                                            VARIANT arg14, VARIANT arg15,
                                            VARIANT arg16, VARIANT arg17,
                                            VARIANT arg18, VARIANT arg19,
                                            VARIANT arg20, VARIANT arg21,
                                            VARIANT arg22, VARIANT arg23,
                                            VARIANT arg24, VARIANT arg25,
                                            VARIANT arg26, VARIANT arg27,
                                            VARIANT arg28, VARIANT arg29,
                                            VARIANT arg30, double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionDVar (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo,
                                             ExcelObj_Range arg1, VARIANT arg2,
                                             VARIANT arg3, double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionText (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo, VARIANT arg1,
                                             const char *arg2,
                                             char **returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionLinEst (CAObjHandle objectHandle,
                                               ERRORINFO *errorInfo,
                                               VARIANT arg1, VARIANT arg2,
                                               VARIANT arg3, VARIANT arg4,
                                               VARIANT *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionTrend (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo, VARIANT arg1,
                                              VARIANT arg2, VARIANT arg3,
                                              VARIANT arg4, VARIANT *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionLogEst (CAObjHandle objectHandle,
                                               ERRORINFO *errorInfo,
                                               VARIANT arg1, VARIANT arg2,
                                               VARIANT arg3, VARIANT arg4,
                                               VARIANT *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionGrowth (CAObjHandle objectHandle,
                                               ERRORINFO *errorInfo,
                                               VARIANT arg1, VARIANT arg2,
                                               VARIANT arg3, VARIANT arg4,
                                               VARIANT *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionPv (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo, double arg1,
                                           double arg2, double arg3,
                                           VARIANT arg4, VARIANT arg5,
                                           double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionFv (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo, double arg1,
                                           double arg2, double arg3,
                                           VARIANT arg4, VARIANT arg5,
                                           double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionNPer (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo, double arg1,
                                             double arg2, double arg3,
                                             VARIANT arg4, VARIANT arg5,
                                             double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionPmt (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo, double arg1,
                                            double arg2, double arg3,
                                            VARIANT arg4, VARIANT arg5,
                                            double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionRate (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo, double arg1,
                                             double arg2, double arg3,
                                             VARIANT arg4, VARIANT arg5,
                                             VARIANT arg6, double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionMIrr (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo, VARIANT arg1,
                                             double arg2, double arg3,
                                             double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionIrr (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo, VARIANT arg1,
                                            VARIANT arg2, double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionMatch (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo, VARIANT arg1,
                                              VARIANT arg2, VARIANT arg3,
                                              double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionWeekday (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo,
                                                VARIANT arg1, VARIANT arg2,
                                                double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionSearch (CAObjHandle objectHandle,
                                               ERRORINFO *errorInfo,
                                               const char *arg1,
                                               const char *arg2, VARIANT arg3,
                                               double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionTranspose (CAObjHandle objectHandle,
                                                  ERRORINFO *errorInfo,
                                                  VARIANT arg1,
                                                  VARIANT *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionAtan2 (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo, double arg1,
                                              double arg2, double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionAsin (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo, double arg1,
                                             double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionAcos (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo, double arg1,
                                             double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionChoose (CAObjHandle objectHandle,
                                               ERRORINFO *errorInfo,
                                               VARIANT arg1, VARIANT arg2,
                                               VARIANT arg3, VARIANT arg4,
                                               VARIANT arg5, VARIANT arg6,
                                               VARIANT arg7, VARIANT arg8,
                                               VARIANT arg9, VARIANT arg10,
                                               VARIANT arg11, VARIANT arg12,
                                               VARIANT arg13, VARIANT arg14,
                                               VARIANT arg15, VARIANT arg16,
                                               VARIANT arg17, VARIANT arg18,
                                               VARIANT arg19, VARIANT arg20,
                                               VARIANT arg21, VARIANT arg22,
                                               VARIANT arg23, VARIANT arg24,
                                               VARIANT arg25, VARIANT arg26,
                                               VARIANT arg27, VARIANT arg28,
                                               VARIANT arg29, VARIANT arg30,
                                               VARIANT *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionHLookup (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo,
                                                VARIANT arg1, VARIANT arg2,
                                                VARIANT arg3, VARIANT arg4,
                                                VARIANT *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionVLookup (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo,
                                                VARIANT arg1, VARIANT arg2,
                                                VARIANT arg3, VARIANT arg4,
                                                VARIANT *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionLog (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo, double arg1,
                                            VARIANT arg2, double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionProper (CAObjHandle objectHandle,
                                               ERRORINFO *errorInfo,
                                               const char *arg1,
                                               char **returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionTrim (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo,
                                             const char *arg1,
                                             char **returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionReplace (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo,
                                                const char *arg1, double arg2,
                                                double arg3, const char *arg4,
                                                char **returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionSubstitute (CAObjHandle objectHandle,
                                                   ERRORINFO *errorInfo,
                                                   const char *arg1,
                                                   const char *arg2,
                                                   const char *arg3,
                                                   VARIANT arg4,
                                                   char **returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionFind (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo,
                                             const char *arg1, const char *arg2,
                                             VARIANT arg3, double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionIsErr (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo, VARIANT arg1,
                                              VBOOL *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionIsText (CAObjHandle objectHandle,
                                               ERRORINFO *errorInfo,
                                               VARIANT arg1, VBOOL *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionIsNumber (CAObjHandle objectHandle,
                                                 ERRORINFO *errorInfo,
                                                 VARIANT arg1,
                                                 VBOOL *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionSln (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo, double arg1,
                                            double arg2, double arg3,
                                            double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionSyd (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo, double arg1,
                                            double arg2, double arg3,
                                            double arg4, double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionDdb (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo, double arg1,
                                            double arg2, double arg3,
                                            double arg4, VARIANT arg5,
                                            double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionClean (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo,
                                              const char *arg1,
                                              char **returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionMDeterm (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo,
                                                VARIANT arg1,
                                                double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionMInverse (CAObjHandle objectHandle,
                                                 ERRORINFO *errorInfo,
                                                 VARIANT arg1,
                                                 VARIANT *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionMMult (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo, VARIANT arg1,
                                              VARIANT arg2, VARIANT *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionIpmt (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo, double arg1,
                                             double arg2, double arg3,
                                             double arg4, VARIANT arg5,
                                             VARIANT arg6, double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionPpmt (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo, double arg1,
                                             double arg2, double arg3,
                                             double arg4, VARIANT arg5,
                                             VARIANT arg6, double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionCountA (CAObjHandle objectHandle,
                                               ERRORINFO *errorInfo,
                                               VARIANT arg1, VARIANT arg2,
                                               VARIANT arg3, VARIANT arg4,
                                               VARIANT arg5, VARIANT arg6,
                                               VARIANT arg7, VARIANT arg8,
                                               VARIANT arg9, VARIANT arg10,
                                               VARIANT arg11, VARIANT arg12,
                                               VARIANT arg13, VARIANT arg14,
                                               VARIANT arg15, VARIANT arg16,
                                               VARIANT arg17, VARIANT arg18,
                                               VARIANT arg19, VARIANT arg20,
                                               VARIANT arg21, VARIANT arg22,
                                               VARIANT arg23, VARIANT arg24,
                                               VARIANT arg25, VARIANT arg26,
                                               VARIANT arg27, VARIANT arg28,
                                               VARIANT arg29, VARIANT arg30,
                                               double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionProduct (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo,
                                                VARIANT arg1, VARIANT arg2,
                                                VARIANT arg3, VARIANT arg4,
                                                VARIANT arg5, VARIANT arg6,
                                                VARIANT arg7, VARIANT arg8,
                                                VARIANT arg9, VARIANT arg10,
                                                VARIANT arg11, VARIANT arg12,
                                                VARIANT arg13, VARIANT arg14,
                                                VARIANT arg15, VARIANT arg16,
                                                VARIANT arg17, VARIANT arg18,
                                                VARIANT arg19, VARIANT arg20,
                                                VARIANT arg21, VARIANT arg22,
                                                VARIANT arg23, VARIANT arg24,
                                                VARIANT arg25, VARIANT arg26,
                                                VARIANT arg27, VARIANT arg28,
                                                VARIANT arg29, VARIANT arg30,
                                                double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionFact (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo, double arg1,
                                             double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionDProduct (CAObjHandle objectHandle,
                                                 ERRORINFO *errorInfo,
                                                 ExcelObj_Range arg1,
                                                 VARIANT arg2, VARIANT arg3,
                                                 double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionIsNonText (CAObjHandle objectHandle,
                                                  ERRORINFO *errorInfo,
                                                  VARIANT arg1,
                                                  VBOOL *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionStDevP (CAObjHandle objectHandle,
                                               ERRORINFO *errorInfo,
                                               VARIANT arg1, VARIANT arg2,
                                               VARIANT arg3, VARIANT arg4,
                                               VARIANT arg5, VARIANT arg6,
                                               VARIANT arg7, VARIANT arg8,
                                               VARIANT arg9, VARIANT arg10,
                                               VARIANT arg11, VARIANT arg12,
                                               VARIANT arg13, VARIANT arg14,
                                               VARIANT arg15, VARIANT arg16,
                                               VARIANT arg17, VARIANT arg18,
                                               VARIANT arg19, VARIANT arg20,
                                               VARIANT arg21, VARIANT arg22,
                                               VARIANT arg23, VARIANT arg24,
                                               VARIANT arg25, VARIANT arg26,
                                               VARIANT arg27, VARIANT arg28,
                                               VARIANT arg29, VARIANT arg30,
                                               double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionVarP (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo, VARIANT arg1,
                                             VARIANT arg2, VARIANT arg3,
                                             VARIANT arg4, VARIANT arg5,
                                             VARIANT arg6, VARIANT arg7,
                                             VARIANT arg8, VARIANT arg9,
                                             VARIANT arg10, VARIANT arg11,
                                             VARIANT arg12, VARIANT arg13,
                                             VARIANT arg14, VARIANT arg15,
                                             VARIANT arg16, VARIANT arg17,
                                             VARIANT arg18, VARIANT arg19,
                                             VARIANT arg20, VARIANT arg21,
                                             VARIANT arg22, VARIANT arg23,
                                             VARIANT arg24, VARIANT arg25,
                                             VARIANT arg26, VARIANT arg27,
                                             VARIANT arg28, VARIANT arg29,
                                             VARIANT arg30, double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionDStDevP (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo,
                                                ExcelObj_Range arg1,
                                                VARIANT arg2, VARIANT arg3,
                                                double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionDVarP (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo,
                                              ExcelObj_Range arg1, VARIANT arg2,
                                              VARIANT arg3, double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionIsLogical (CAObjHandle objectHandle,
                                                  ERRORINFO *errorInfo,
                                                  VARIANT arg1,
                                                  VBOOL *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionDCountA (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo,
                                                ExcelObj_Range arg1,
                                                VARIANT arg2, VARIANT arg3,
                                                double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionUSDollar (CAObjHandle objectHandle,
                                                 ERRORINFO *errorInfo,
                                                 double arg1, double arg2,
                                                 char **returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionFindB (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo,
                                              const char *arg1, const char *arg2,
                                              VARIANT arg3, double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionSearchB (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo,
                                                const char *arg1,
                                                const char *arg2, VARIANT arg3,
                                                double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionReplaceB (CAObjHandle objectHandle,
                                                 ERRORINFO *errorInfo,
                                                 const char *arg1, double arg2,
                                                 double arg3, const char *arg4,
                                                 char **returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionRoundUp (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo,
                                                double arg1, double arg2,
                                                double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionRoundDown (CAObjHandle objectHandle,
                                                  ERRORINFO *errorInfo,
                                                  double arg1, double arg2,
                                                  double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionRank (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo, double arg1,
                                             ExcelObj_Range arg2, VARIANT arg3,
                                             double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionDays360 (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo,
                                                VARIANT arg1, VARIANT arg2,
                                                VARIANT arg3,
                                                double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionVdb (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo, double arg1,
                                            double arg2, double arg3,
                                            double arg4, double arg5,
                                            VARIANT arg6, VARIANT arg7,
                                            double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionMedian (CAObjHandle objectHandle,
                                               ERRORINFO *errorInfo,
                                               VARIANT arg1, VARIANT arg2,
                                               VARIANT arg3, VARIANT arg4,
                                               VARIANT arg5, VARIANT arg6,
                                               VARIANT arg7, VARIANT arg8,
                                               VARIANT arg9, VARIANT arg10,
                                               VARIANT arg11, VARIANT arg12,
                                               VARIANT arg13, VARIANT arg14,
                                               VARIANT arg15, VARIANT arg16,
                                               VARIANT arg17, VARIANT arg18,
                                               VARIANT arg19, VARIANT arg20,
                                               VARIANT arg21, VARIANT arg22,
                                               VARIANT arg23, VARIANT arg24,
                                               VARIANT arg25, VARIANT arg26,
                                               VARIANT arg27, VARIANT arg28,
                                               VARIANT arg29, VARIANT arg30,
                                               double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionSumProduct (CAObjHandle objectHandle,
                                                   ERRORINFO *errorInfo,
                                                   VARIANT arg1, VARIANT arg2,
                                                   VARIANT arg3, VARIANT arg4,
                                                   VARIANT arg5, VARIANT arg6,
                                                   VARIANT arg7, VARIANT arg8,
                                                   VARIANT arg9, VARIANT arg10,
                                                   VARIANT arg11, VARIANT arg12,
                                                   VARIANT arg13, VARIANT arg14,
                                                   VARIANT arg15, VARIANT arg16,
                                                   VARIANT arg17, VARIANT arg18,
                                                   VARIANT arg19, VARIANT arg20,
                                                   VARIANT arg21, VARIANT arg22,
                                                   VARIANT arg23, VARIANT arg24,
                                                   VARIANT arg25, VARIANT arg26,
                                                   VARIANT arg27, VARIANT arg28,
                                                   VARIANT arg29, VARIANT arg30,
                                                   double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionSinh (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo, double arg1,
                                             double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionCosh (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo, double arg1,
                                             double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionTanh (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo, double arg1,
                                             double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionAsinh (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo, double arg1,
                                              double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionAcosh (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo, double arg1,
                                              double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionAtanh (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo, double arg1,
                                              double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionDGet (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo,
                                             ExcelObj_Range arg1, VARIANT arg2,
                                             VARIANT arg3, VARIANT *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionDb (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo, double arg1,
                                           double arg2, double arg3, double arg4,
                                           VARIANT arg5, double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionFrequency (CAObjHandle objectHandle,
                                                  ERRORINFO *errorInfo,
                                                  VARIANT arg1, VARIANT arg2,
                                                  VARIANT *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionAveDev (CAObjHandle objectHandle,
                                               ERRORINFO *errorInfo,
                                               VARIANT arg1, VARIANT arg2,
                                               VARIANT arg3, VARIANT arg4,
                                               VARIANT arg5, VARIANT arg6,
                                               VARIANT arg7, VARIANT arg8,
                                               VARIANT arg9, VARIANT arg10,
                                               VARIANT arg11, VARIANT arg12,
                                               VARIANT arg13, VARIANT arg14,
                                               VARIANT arg15, VARIANT arg16,
                                               VARIANT arg17, VARIANT arg18,
                                               VARIANT arg19, VARIANT arg20,
                                               VARIANT arg21, VARIANT arg22,
                                               VARIANT arg23, VARIANT arg24,
                                               VARIANT arg25, VARIANT arg26,
                                               VARIANT arg27, VARIANT arg28,
                                               VARIANT arg29, VARIANT arg30,
                                               double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionBetaDist (CAObjHandle objectHandle,
                                                 ERRORINFO *errorInfo,
                                                 double arg1, double arg2,
                                                 double arg3, VARIANT arg4,
                                                 VARIANT arg5,
                                                 double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionGammaLn (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo,
                                                double arg1, double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionBetaInv (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo,
                                                double arg1, double arg2,
                                                double arg3, VARIANT arg4,
                                                VARIANT arg5,
                                                double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionBinomDist (CAObjHandle objectHandle,
                                                  ERRORINFO *errorInfo,
                                                  double arg1, double arg2,
                                                  double arg3, VBOOL arg4,
                                                  double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionChiDist (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo,
                                                double arg1, double arg2,
                                                double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionChiInv (CAObjHandle objectHandle,
                                               ERRORINFO *errorInfo, double arg1,
                                               double arg2, double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionCombin (CAObjHandle objectHandle,
                                               ERRORINFO *errorInfo, double arg1,
                                               double arg2, double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionConfidence (CAObjHandle objectHandle,
                                                   ERRORINFO *errorInfo,
                                                   double arg1, double arg2,
                                                   double arg3,
                                                   double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionCritBinom (CAObjHandle objectHandle,
                                                  ERRORINFO *errorInfo,
                                                  double arg1, double arg2,
                                                  double arg3,
                                                  double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionEven (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo, double arg1,
                                             double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionExponDist (CAObjHandle objectHandle,
                                                  ERRORINFO *errorInfo,
                                                  double arg1, double arg2,
                                                  VBOOL arg3,
                                                  double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionFDist (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo, double arg1,
                                              double arg2, double arg3,
                                              double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionFInv (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo, double arg1,
                                             double arg2, double arg3,
                                             double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionFisher (CAObjHandle objectHandle,
                                               ERRORINFO *errorInfo, double arg1,
                                               double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionFisherInv (CAObjHandle objectHandle,
                                                  ERRORINFO *errorInfo,
                                                  double arg1,
                                                  double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionFloor (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo, double arg1,
                                              double arg2, double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionGammaDist (CAObjHandle objectHandle,
                                                  ERRORINFO *errorInfo,
                                                  double arg1, double arg2,
                                                  double arg3, VBOOL arg4,
                                                  double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionGammaInv (CAObjHandle objectHandle,
                                                 ERRORINFO *errorInfo,
                                                 double arg1, double arg2,
                                                 double arg3,
                                                 double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionCeiling (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo,
                                                double arg1, double arg2,
                                                double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionHypGeomDist (CAObjHandle objectHandle,
                                                    ERRORINFO *errorInfo,
                                                    double arg1, double arg2,
                                                    double arg3, double arg4,
                                                    double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionLogNormDist (CAObjHandle objectHandle,
                                                    ERRORINFO *errorInfo,
                                                    double arg1, double arg2,
                                                    double arg3,
                                                    double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionLogInv (CAObjHandle objectHandle,
                                               ERRORINFO *errorInfo, double arg1,
                                               double arg2, double arg3,
                                               double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionNegBinomDist (CAObjHandle objectHandle,
                                                     ERRORINFO *errorInfo,
                                                     double arg1, double arg2,
                                                     double arg3,
                                                     double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionNormDist (CAObjHandle objectHandle,
                                                 ERRORINFO *errorInfo,
                                                 double arg1, double arg2,
                                                 double arg3, VBOOL arg4,
                                                 double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionNormSDist (CAObjHandle objectHandle,
                                                  ERRORINFO *errorInfo,
                                                  double arg1,
                                                  double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionNormInv (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo,
                                                double arg1, double arg2,
                                                double arg3, double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionNormSInv (CAObjHandle objectHandle,
                                                 ERRORINFO *errorInfo,
                                                 double arg1,
                                                 double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionStandardize (CAObjHandle objectHandle,
                                                    ERRORINFO *errorInfo,
                                                    double arg1, double arg2,
                                                    double arg3,
                                                    double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionOdd (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo, double arg1,
                                            double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionPermut (CAObjHandle objectHandle,
                                               ERRORINFO *errorInfo, double arg1,
                                               double arg2, double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionPoisson (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo,
                                                double arg1, double arg2,
                                                VBOOL arg3, double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionTDist (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo, double arg1,
                                              double arg2, double arg3,
                                              double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionWeibull (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo,
                                                double arg1, double arg2,
                                                double arg3, VBOOL arg4,
                                                double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionSumXMY2 (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo,
                                                VARIANT arg1, VARIANT arg2,
                                                double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionSumX2MY2 (CAObjHandle objectHandle,
                                                 ERRORINFO *errorInfo,
                                                 VARIANT arg1, VARIANT arg2,
                                                 double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionSumX2PY2 (CAObjHandle objectHandle,
                                                 ERRORINFO *errorInfo,
                                                 VARIANT arg1, VARIANT arg2,
                                                 double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionChiTest (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo,
                                                VARIANT arg1, VARIANT arg2,
                                                double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionCorrel (CAObjHandle objectHandle,
                                               ERRORINFO *errorInfo,
                                               VARIANT arg1, VARIANT arg2,
                                               double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionCovar (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo, VARIANT arg1,
                                              VARIANT arg2, double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionForecast (CAObjHandle objectHandle,
                                                 ERRORINFO *errorInfo,
                                                 double arg1, VARIANT arg2,
                                                 VARIANT arg3,
                                                 double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionFTest (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo, VARIANT arg1,
                                              VARIANT arg2, double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionIntercept (CAObjHandle objectHandle,
                                                  ERRORINFO *errorInfo,
                                                  VARIANT arg1, VARIANT arg2,
                                                  double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionPearson (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo,
                                                VARIANT arg1, VARIANT arg2,
                                                double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionRSq (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo, VARIANT arg1,
                                            VARIANT arg2, double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionStEyx (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo, VARIANT arg1,
                                              VARIANT arg2, double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionSlope (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo, VARIANT arg1,
                                              VARIANT arg2, double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionTTest (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo, VARIANT arg1,
                                              VARIANT arg2, double arg3,
                                              double arg4, double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionProb (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo, VARIANT arg1,
                                             VARIANT arg2, double arg3,
                                             VARIANT arg4, double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionDevSq (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo, VARIANT arg1,
                                              VARIANT arg2, VARIANT arg3,
                                              VARIANT arg4, VARIANT arg5,
                                              VARIANT arg6, VARIANT arg7,
                                              VARIANT arg8, VARIANT arg9,
                                              VARIANT arg10, VARIANT arg11,
                                              VARIANT arg12, VARIANT arg13,
                                              VARIANT arg14, VARIANT arg15,
                                              VARIANT arg16, VARIANT arg17,
                                              VARIANT arg18, VARIANT arg19,
                                              VARIANT arg20, VARIANT arg21,
                                              VARIANT arg22, VARIANT arg23,
                                              VARIANT arg24, VARIANT arg25,
                                              VARIANT arg26, VARIANT arg27,
                                              VARIANT arg28, VARIANT arg29,
                                              VARIANT arg30, double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionGeoMean (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo,
                                                VARIANT arg1, VARIANT arg2,
                                                VARIANT arg3, VARIANT arg4,
                                                VARIANT arg5, VARIANT arg6,
                                                VARIANT arg7, VARIANT arg8,
                                                VARIANT arg9, VARIANT arg10,
                                                VARIANT arg11, VARIANT arg12,
                                                VARIANT arg13, VARIANT arg14,
                                                VARIANT arg15, VARIANT arg16,
                                                VARIANT arg17, VARIANT arg18,
                                                VARIANT arg19, VARIANT arg20,
                                                VARIANT arg21, VARIANT arg22,
                                                VARIANT arg23, VARIANT arg24,
                                                VARIANT arg25, VARIANT arg26,
                                                VARIANT arg27, VARIANT arg28,
                                                VARIANT arg29, VARIANT arg30,
                                                double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionHarMean (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo,
                                                VARIANT arg1, VARIANT arg2,
                                                VARIANT arg3, VARIANT arg4,
                                                VARIANT arg5, VARIANT arg6,
                                                VARIANT arg7, VARIANT arg8,
                                                VARIANT arg9, VARIANT arg10,
                                                VARIANT arg11, VARIANT arg12,
                                                VARIANT arg13, VARIANT arg14,
                                                VARIANT arg15, VARIANT arg16,
                                                VARIANT arg17, VARIANT arg18,
                                                VARIANT arg19, VARIANT arg20,
                                                VARIANT arg21, VARIANT arg22,
                                                VARIANT arg23, VARIANT arg24,
                                                VARIANT arg25, VARIANT arg26,
                                                VARIANT arg27, VARIANT arg28,
                                                VARIANT arg29, VARIANT arg30,
                                                double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionSumSq (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo, VARIANT arg1,
                                              VARIANT arg2, VARIANT arg3,
                                              VARIANT arg4, VARIANT arg5,
                                              VARIANT arg6, VARIANT arg7,
                                              VARIANT arg8, VARIANT arg9,
                                              VARIANT arg10, VARIANT arg11,
                                              VARIANT arg12, VARIANT arg13,
                                              VARIANT arg14, VARIANT arg15,
                                              VARIANT arg16, VARIANT arg17,
                                              VARIANT arg18, VARIANT arg19,
                                              VARIANT arg20, VARIANT arg21,
                                              VARIANT arg22, VARIANT arg23,
                                              VARIANT arg24, VARIANT arg25,
                                              VARIANT arg26, VARIANT arg27,
                                              VARIANT arg28, VARIANT arg29,
                                              VARIANT arg30, double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionKurt (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo, VARIANT arg1,
                                             VARIANT arg2, VARIANT arg3,
                                             VARIANT arg4, VARIANT arg5,
                                             VARIANT arg6, VARIANT arg7,
                                             VARIANT arg8, VARIANT arg9,
                                             VARIANT arg10, VARIANT arg11,
                                             VARIANT arg12, VARIANT arg13,
                                             VARIANT arg14, VARIANT arg15,
                                             VARIANT arg16, VARIANT arg17,
                                             VARIANT arg18, VARIANT arg19,
                                             VARIANT arg20, VARIANT arg21,
                                             VARIANT arg22, VARIANT arg23,
                                             VARIANT arg24, VARIANT arg25,
                                             VARIANT arg26, VARIANT arg27,
                                             VARIANT arg28, VARIANT arg29,
                                             VARIANT arg30, double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionSkew (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo, VARIANT arg1,
                                             VARIANT arg2, VARIANT arg3,
                                             VARIANT arg4, VARIANT arg5,
                                             VARIANT arg6, VARIANT arg7,
                                             VARIANT arg8, VARIANT arg9,
                                             VARIANT arg10, VARIANT arg11,
                                             VARIANT arg12, VARIANT arg13,
                                             VARIANT arg14, VARIANT arg15,
                                             VARIANT arg16, VARIANT arg17,
                                             VARIANT arg18, VARIANT arg19,
                                             VARIANT arg20, VARIANT arg21,
                                             VARIANT arg22, VARIANT arg23,
                                             VARIANT arg24, VARIANT arg25,
                                             VARIANT arg26, VARIANT arg27,
                                             VARIANT arg28, VARIANT arg29,
                                             VARIANT arg30, double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionZTest (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo, VARIANT arg1,
                                              double arg2, VARIANT arg3,
                                              double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionLarge (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo, VARIANT arg1,
                                              double arg2, double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionSmall (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo, VARIANT arg1,
                                              double arg2, double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionQuartile (CAObjHandle objectHandle,
                                                 ERRORINFO *errorInfo,
                                                 VARIANT arg1, double arg2,
                                                 double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionPercentile (CAObjHandle objectHandle,
                                                   ERRORINFO *errorInfo,
                                                   VARIANT arg1, double arg2,
                                                   double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionPercentRank (CAObjHandle objectHandle,
                                                    ERRORINFO *errorInfo,
                                                    VARIANT arg1, double arg2,
                                                    VARIANT arg3,
                                                    double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionMode (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo, VARIANT arg1,
                                             VARIANT arg2, VARIANT arg3,
                                             VARIANT arg4, VARIANT arg5,
                                             VARIANT arg6, VARIANT arg7,
                                             VARIANT arg8, VARIANT arg9,
                                             VARIANT arg10, VARIANT arg11,
                                             VARIANT arg12, VARIANT arg13,
                                             VARIANT arg14, VARIANT arg15,
                                             VARIANT arg16, VARIANT arg17,
                                             VARIANT arg18, VARIANT arg19,
                                             VARIANT arg20, VARIANT arg21,
                                             VARIANT arg22, VARIANT arg23,
                                             VARIANT arg24, VARIANT arg25,
                                             VARIANT arg26, VARIANT arg27,
                                             VARIANT arg28, VARIANT arg29,
                                             VARIANT arg30, double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionTrimMean (CAObjHandle objectHandle,
                                                 ERRORINFO *errorInfo,
                                                 VARIANT arg1, double arg2,
                                                 double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionTInv (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo, double arg1,
                                             double arg2, double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionPower (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo, double arg1,
                                              double arg2, double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionRadians (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo,
                                                double arg1, double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionDegrees (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo,
                                                double arg1, double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionSubtotal (CAObjHandle objectHandle,
                                                 ERRORINFO *errorInfo,
                                                 double arg1,
                                                 ExcelObj_Range arg2,
                                                 VARIANT arg3, VARIANT arg4,
                                                 VARIANT arg5, VARIANT arg6,
                                                 VARIANT arg7, VARIANT arg8,
                                                 VARIANT arg9, VARIANT arg10,
                                                 VARIANT arg11, VARIANT arg12,
                                                 VARIANT arg13, VARIANT arg14,
                                                 VARIANT arg15, VARIANT arg16,
                                                 VARIANT arg17, VARIANT arg18,
                                                 VARIANT arg19, VARIANT arg20,
                                                 VARIANT arg21, VARIANT arg22,
                                                 VARIANT arg23, VARIANT arg24,
                                                 VARIANT arg25, VARIANT arg26,
                                                 VARIANT arg27, VARIANT arg28,
                                                 VARIANT arg29, VARIANT arg30,
                                                 double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionSumIf (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo,
                                              ExcelObj_Range arg1, VARIANT arg2,
                                              VARIANT arg3, double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionCountIf (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo,
                                                ExcelObj_Range arg1,
                                                VARIANT arg2,
                                                double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionCountBlank (CAObjHandle objectHandle,
                                                   ERRORINFO *errorInfo,
                                                   ExcelObj_Range arg1,
                                                   double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionIspmt (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo, double arg1,
                                              double arg2, double arg3,
                                              double arg4, double *returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionRoman (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo, double arg1,
                                              VARIANT arg2, char **returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionAsc (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo,
                                            const char *arg1, char **returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionDbcs (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo,
                                             const char *arg1,
                                             char **returnValue);

HRESULT CVIFUNC Excel_WorksheetFunctionPhonetic (CAObjHandle objectHandle,
                                                 ERRORINFO *errorInfo,
                                                 ExcelObj_Range arg1,
                                                 char **returnValue);

HRESULT CVIFUNC Excel_RangeActivate (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo, VARIANT *returnValue);

HRESULT CVIFUNC Excel_RangeAddress (CAObjHandle objectHandle,
                                    ERRORINFO *errorInfo, VARIANT rowAbsolute,
                                    VARIANT columnAbsolute,
                                    enum ExcelEnum_XlReferenceStyle referenceStyle,
                                    VARIANT external, VARIANT relativeTo,
                                    char **returnValue);

HRESULT CVIFUNC Excel_RangeAddressLocal (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo,
                                         VARIANT rowAbsolute,
                                         VARIANT columnAbsolute,
                                         enum ExcelEnum_XlReferenceStyle referenceStyle,
                                         VARIANT external, VARIANT relativeTo,
                                         char **returnValue);

HRESULT CVIFUNC Excel_RangeAdvancedFilter (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo,
                                           enum ExcelEnum_XlFilterAction action,
                                           VARIANT criteriaRange,
                                           VARIANT copyToRange, VARIANT unique,
                                           VARIANT *returnValue);

HRESULT CVIFUNC Excel_RangeApplyNames (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo, VARIANT names,
                                       VARIANT ignoreRelativeAbsolute,
                                       VARIANT useRowColumnNames,
                                       VARIANT omitColumn, VARIANT omitRow,
                                       enum ExcelEnum_XlApplyNamesOrder order,
                                       VARIANT appendLast, VARIANT *returnValue);

HRESULT CVIFUNC Excel_RangeApplyOutlineStyles (CAObjHandle objectHandle,
                                               ERRORINFO *errorInfo,
                                               VARIANT *returnValue);

HRESULT CVIFUNC Excel_RangeAutoComplete (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo,
                                         const char *string, char **returnValue);

HRESULT CVIFUNC Excel_RangeAutoFill (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo,
                                     ExcelObj_Range destination,
                                     enum ExcelEnum_XlAutoFillType type,
                                     VARIANT *returnValue);

HRESULT CVIFUNC Excel_RangeAutoFilter (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo, VARIANT field,
                                       VARIANT criteria1,
                                       enum ExcelEnum_XlAutoFilterOperator operator,
                                       VARIANT criteria2,
                                       VARIANT visibleDropDown,
                                       VARIANT *returnValue);

HRESULT CVIFUNC Excel_RangeAutoFit (CAObjHandle objectHandle,
                                    ERRORINFO *errorInfo, VARIANT *returnValue);

HRESULT CVIFUNC Excel_RangeAutoFormat (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo,
                                       enum ExcelEnum_XlRangeAutoFormat format,
                                       VARIANT number, VARIANT font,
                                       VARIANT alignment, VARIANT border,
                                       VARIANT pattern, VARIANT width,
                                       VARIANT *returnValue);

HRESULT CVIFUNC Excel_RangeAutoOutline (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo,
                                        VARIANT *returnValue);

HRESULT CVIFUNC Excel_RangeBorderAround (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo, VARIANT lineStyle,
                                         enum ExcelEnum_XlBorderWeight weight,
                                         enum ExcelEnum_XlColorIndex colorIndex,
                                         VARIANT color, VARIANT *returnValue);

HRESULT CVIFUNC Excel_RangeCalculate (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT *returnValue);

HRESULT CVIFUNC Excel_RangeCharacters (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo, VARIANT start,
                                       VARIANT length,
                                       ExcelObj_Characters *returnValue);

HRESULT CVIFUNC Excel_RangeCheckSpelling (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo,
                                          VARIANT customDictionary,
                                          VARIANT ignoreUppercase,
                                          VARIANT alwaysSuggest,
                                          VARIANT spellLang,
                                          VARIANT *returnValue);

HRESULT CVIFUNC Excel_RangeClear (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                  VARIANT *returnValue);

HRESULT CVIFUNC Excel_RangeClearContents (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo,
                                          VARIANT *returnValue);

HRESULT CVIFUNC Excel_RangeClearFormats (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo,
                                         VARIANT *returnValue);

HRESULT CVIFUNC Excel_RangeClearNotes (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo,
                                       VARIANT *returnValue);

HRESULT CVIFUNC Excel_RangeClearOutline (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo,
                                         VARIANT *returnValue);

HRESULT CVIFUNC Excel_RangeColumnDifferences (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo,
                                              VARIANT comparison,
                                              ExcelObj_Range *returnValue);

HRESULT CVIFUNC Excel_RangeConsolidate (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo, VARIANT sources,
                                        VARIANT function, VARIANT topRow,
                                        VARIANT leftColumn, VARIANT createLinks,
                                        VARIANT *returnValue);

HRESULT CVIFUNC Excel_RangeCopy (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                 VARIANT destination, VARIANT *returnValue);

HRESULT CVIFUNC Excel_RangeCopyFromRecordset (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo,
                                              LPUNKNOWN data, VARIANT maxRows,
                                              VARIANT maxColumns,
                                              long *returnValue);

HRESULT CVIFUNC Excel_RangeCopyPicture (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo,
                                        enum ExcelEnum_XlPictureAppearance appearance,
                                        enum ExcelEnum_XlCopyPictureFormat format,
                                        VARIANT *returnValue);

HRESULT CVIFUNC Excel_RangeCreateNames (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo, VARIANT top,
                                        VARIANT left, VARIANT bottom,
                                        VARIANT right, VARIANT *returnValue);

HRESULT CVIFUNC Excel_RangeCreatePublisher (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo,
                                            VARIANT edition,
                                            enum ExcelEnum_XlPictureAppearance appearance,
                                            VARIANT containsPICT,
                                            VARIANT containsBIFF,
                                            VARIANT containsRTF,
                                            VARIANT containsVALU,
                                            VARIANT *returnValue);

HRESULT CVIFUNC Excel_RangeCut (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                VARIANT destination, VARIANT *returnValue);

HRESULT CVIFUNC Excel_RangeDataSeries (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo, VARIANT rowcol,
                                       enum ExcelEnum_XlDataSeriesType type,
                                       enum ExcelEnum_XlDataSeriesDate date,
                                       VARIANT step, VARIANT stop, VARIANT trend,
                                       VARIANT *returnValue);

HRESULT CVIFUNC Excel_RangeGet_Default (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo, VARIANT rowIndex,
                                        VARIANT columnIndex,
                                        VARIANT *returnValue);

HRESULT CVIFUNC Excel_RangeSet_Default (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo, VARIANT rowIndex,
                                        VARIANT columnIndex, VARIANT newValue);

HRESULT CVIFUNC Excel_RangeDelete (CAObjHandle objectHandle,
                                   ERRORINFO *errorInfo, VARIANT shift,
                                   VARIANT *returnValue);

HRESULT CVIFUNC Excel_RangeDialogBox (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT *returnValue);

HRESULT CVIFUNC Excel_RangeEditionOptions (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo,
                                           enum ExcelEnum_XlEditionType type,
                                           enum ExcelEnum_XlEditionOptionsOption option,
                                           VARIANT name, VARIANT reference,
                                           enum ExcelEnum_XlPictureAppearance appearance,
                                           enum ExcelEnum_XlPictureAppearance chartSize,
                                           VARIANT format, VARIANT *returnValue);

HRESULT CVIFUNC Excel_RangeEnd (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                enum ExcelEnum_XlDirection direction,
                                ExcelObj_Range *returnValue);

HRESULT CVIFUNC Excel_RangeFillDown (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo, VARIANT *returnValue);

HRESULT CVIFUNC Excel_RangeFillLeft (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo, VARIANT *returnValue);

HRESULT CVIFUNC Excel_RangeFillRight (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT *returnValue);

HRESULT CVIFUNC Excel_RangeFillUp (CAObjHandle objectHandle,
                                   ERRORINFO *errorInfo, VARIANT *returnValue);

HRESULT CVIFUNC Excel_RangeFind (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                 VARIANT what, VARIANT after, VARIANT lookIn,
                                 VARIANT lookAt, VARIANT searchOrder,
                                 enum ExcelEnum_XlSearchDirection searchDirection,
                                 VARIANT matchCase, VARIANT matchByte,
                                 ExcelObj_Range *returnValue);

HRESULT CVIFUNC Excel_RangeFindNext (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo, VARIANT after,
                                     ExcelObj_Range *returnValue);

HRESULT CVIFUNC Excel_RangeFindPrevious (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo, VARIANT after,
                                         ExcelObj_Range *returnValue);

HRESULT CVIFUNC Excel_RangeFunctionWizard (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo,
                                           VARIANT *returnValue);

HRESULT CVIFUNC Excel_RangeGoalSeek (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo, VARIANT goal,
                                     ExcelObj_Range changingCell,
                                     VBOOL *returnValue);

HRESULT CVIFUNC Excel_RangeGroup (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                  VARIANT start, VARIANT end, VARIANT by,
                                  VARIANT periods, VARIANT *returnValue);

HRESULT CVIFUNC Excel_RangeInsertIndent (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo, long insertAmount);

HRESULT CVIFUNC Excel_RangeInsert (CAObjHandle objectHandle,
                                   ERRORINFO *errorInfo, VARIANT shift,
                                   VARIANT *returnValue);

HRESULT CVIFUNC Excel_RangeGetItem (CAObjHandle objectHandle,
                                    ERRORINFO *errorInfo, VARIANT rowIndex,
                                    VARIANT columnIndex, VARIANT *returnValue);

HRESULT CVIFUNC Excel_RangeSetItem (CAObjHandle objectHandle,
                                    ERRORINFO *errorInfo, VARIANT rowIndex,
                                    VARIANT columnIndex, VARIANT newValue);

HRESULT CVIFUNC Excel_RangeJustify (CAObjHandle objectHandle,
                                    ERRORINFO *errorInfo, VARIANT *returnValue);

HRESULT CVIFUNC Excel_RangeListNames (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT *returnValue);

HRESULT CVIFUNC Excel_RangeMerge (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                  VARIANT across);

HRESULT CVIFUNC Excel_RangeUnMerge (CAObjHandle objectHandle,
                                    ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_RangeNavigateArrow (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo,
                                          VARIANT towardPrecedent,
                                          VARIANT arrowNumber,
                                          VARIANT linkNumber,
                                          VARIANT *returnValue);

HRESULT CVIFUNC Excel_RangeNoteText (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo, VARIANT text,
                                     VARIANT start, VARIANT length,
                                     char **returnValue);

HRESULT CVIFUNC Excel_RangeOffset (CAObjHandle objectHandle,
                                   ERRORINFO *errorInfo, VARIANT rowOffset,
                                   VARIANT columnOffset,
                                   ExcelObj_Range *returnValue);

HRESULT CVIFUNC Excel_RangeParse (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                  VARIANT parseLine, VARIANT destination,
                                  VARIANT *returnValue);

HRESULT CVIFUNC Excel_RangePasteSpecial (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo,
                                         enum ExcelEnum_XlPasteType paste,
                                         enum ExcelEnum_XlPasteSpecialOperation operation,
                                         VARIANT skipBlanks, VARIANT transpose,
                                         VARIANT *returnValue);

HRESULT CVIFUNC Excel_Range_PrintOut (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT from,
                                      VARIANT to, VARIANT copies,
                                      VARIANT preview, VARIANT activePrinter,
                                      VARIANT printToFile, VARIANT collate,
                                      VARIANT *returnValue);

HRESULT CVIFUNC Excel_RangePrintPreview (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo,
                                         VARIANT enableChanges,
                                         VARIANT *returnValue);

HRESULT CVIFUNC Excel_RangeRange (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                  VARIANT cell1, VARIANT cell2,
                                  ExcelObj_Range *returnValue);

HRESULT CVIFUNC Excel_RangeRemoveSubtotal (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo,
                                           VARIANT *returnValue);

HRESULT CVIFUNC Excel_RangeReplace (CAObjHandle objectHandle,
                                    ERRORINFO *errorInfo, VARIANT what,
                                    VARIANT replacement, VARIANT lookAt,
                                    VARIANT searchOrder, VARIANT matchCase,
                                    VARIANT matchByte, VBOOL *returnValue);

HRESULT CVIFUNC Excel_RangeResize (CAObjHandle objectHandle,
                                   ERRORINFO *errorInfo, VARIANT rowSize,
                                   VARIANT columnSize,
                                   ExcelObj_Range *returnValue);

HRESULT CVIFUNC Excel_RangeRowDifferences (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo,
                                           VARIANT comparison,
                                           ExcelObj_Range *returnValue);

HRESULT CVIFUNC Excel_RangeRun (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                VARIANT arg1, VARIANT arg2, VARIANT arg3,
                                VARIANT arg4, VARIANT arg5, VARIANT arg6,
                                VARIANT arg7, VARIANT arg8, VARIANT arg9,
                                VARIANT arg10, VARIANT arg11, VARIANT arg12,
                                VARIANT arg13, VARIANT arg14, VARIANT arg15,
                                VARIANT arg16, VARIANT arg17, VARIANT arg18,
                                VARIANT arg19, VARIANT arg20, VARIANT arg21,
                                VARIANT arg22, VARIANT arg23, VARIANT arg24,
                                VARIANT arg25, VARIANT arg26, VARIANT arg27,
                                VARIANT arg28, VARIANT arg29, VARIANT arg30,
                                VARIANT *returnValue);

HRESULT CVIFUNC Excel_RangeSelect (CAObjHandle objectHandle,
                                   ERRORINFO *errorInfo, VARIANT *returnValue);

HRESULT CVIFUNC Excel_RangeShow (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                 VARIANT *returnValue);

HRESULT CVIFUNC Excel_RangeShowDependents (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo, VARIANT remove,
                                           VARIANT *returnValue);

HRESULT CVIFUNC Excel_RangeShowErrors (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo,
                                       VARIANT *returnValue);

HRESULT CVIFUNC Excel_RangeShowPrecedents (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo, VARIANT remove,
                                           VARIANT *returnValue);

HRESULT CVIFUNC Excel_RangeSort (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                 VARIANT key1, enum ExcelEnum_XlSortOrder order1,
                                 VARIANT key2, VARIANT type,
                                 enum ExcelEnum_XlSortOrder order2, VARIANT key3,
                                 enum ExcelEnum_XlSortOrder order3,
                                 enum ExcelEnum_XlYesNoGuess header,
                                 VARIANT orderCustom, VARIANT matchCase,
                                 enum ExcelEnum_XlSortOrientation orientation,
                                 enum ExcelEnum_XlSortMethod sortMethod,
                                 VARIANT *returnValue);

HRESULT CVIFUNC Excel_RangeSortSpecial (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo,
                                        enum ExcelEnum_XlSortMethod sortMethod,
                                        VARIANT key1,
                                        enum ExcelEnum_XlSortOrder order1,
                                        VARIANT type, VARIANT key2,
                                        enum ExcelEnum_XlSortOrder order2,
                                        VARIANT key3,
                                        enum ExcelEnum_XlSortOrder order3,
                                        enum ExcelEnum_XlYesNoGuess header,
                                        VARIANT orderCustom, VARIANT matchCase,
                                        enum ExcelEnum_XlSortOrientation orientation,
                                        VARIANT *returnValue);

HRESULT CVIFUNC Excel_RangeSpecialCells (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo,
                                         enum ExcelEnum_XlCellType type,
                                         VARIANT value,
                                         ExcelObj_Range *returnValue);

HRESULT CVIFUNC Excel_RangeSubscribeTo (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo,
                                        const char *edition,
                                        enum ExcelEnum_XlSubscribeToFormat format,
                                        VARIANT *returnValue);

HRESULT CVIFUNC Excel_RangeSubtotal (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo, long groupBy,
                                     enum ExcelEnum_XlConsolidationFunction function,
                                     VARIANT totalList, VARIANT replace,
                                     VARIANT pageBreaks,
                                     enum ExcelEnum_XlSummaryRow summaryBelowData,
                                     VARIANT *returnValue);

HRESULT CVIFUNC Excel_RangeTable (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                  VARIANT rowInput, VARIANT columnInput,
                                  VARIANT *returnValue);

HRESULT CVIFUNC Excel_RangeTextToColumns (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo,
                                          VARIANT destination,
                                          enum ExcelEnum_XlTextParsingType dataType,
                                          enum ExcelEnum_XlTextQualifier textQualifier,
                                          VARIANT consecutiveDelimiter,
                                          VARIANT tab, VARIANT semicolon,
                                          VARIANT comma, VARIANT space,
                                          VARIANT other, VARIANT otherChar,
                                          VARIANT fieldInfo,
                                          VARIANT decimalSeparator,
                                          VARIANT thousandsSeparator,
                                          VARIANT *returnValue);

HRESULT CVIFUNC Excel_RangeUngroup (CAObjHandle objectHandle,
                                    ERRORINFO *errorInfo, VARIANT *returnValue);

HRESULT CVIFUNC Excel_RangeAddComment (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo, VARIANT text,
                                       ExcelObj_Comment *returnValue);

HRESULT CVIFUNC Excel_RangeClearComments (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_RangeSetPhonetic (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_RangePrintOut (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo, VARIANT from,
                                     VARIANT to, VARIANT copies, VARIANT preview,
                                     VARIANT activePrinter, VARIANT printToFile,
                                     VARIANT collate, VARIANT prToFileName,
                                     VARIANT *returnValue);

HRESULT CVIFUNC Excel_VPageBreakDelete (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_VPageBreakDragOff (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo,
                                         enum ExcelEnum_XlDirection direction,
                                         long regionIndex);

HRESULT CVIFUNC Excel_HPageBreakDelete (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_HPageBreakDragOff (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo,
                                         enum ExcelEnum_XlDirection direction,
                                         long regionIndex);

HRESULT CVIFUNC Excel_HPageBreaksItem (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo, long index,
                                       ExcelObj_HPageBreak *returnValue);

HRESULT CVIFUNC Excel_HPageBreaks_Default (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo, long index,
                                           ExcelObj_HPageBreak *returnValue);

HRESULT CVIFUNC Excel_HPageBreaksAdd (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, CAObjHandle before,
                                      ExcelObj_HPageBreak *returnValue);

HRESULT CVIFUNC Excel_VPageBreaksItem (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo, long index,
                                       ExcelObj_VPageBreak *returnValue);

HRESULT CVIFUNC Excel_VPageBreaks_Default (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo, long index,
                                           ExcelObj_VPageBreak *returnValue);

HRESULT CVIFUNC Excel_VPageBreaksAdd (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, CAObjHandle before,
                                      ExcelObj_VPageBreak *returnValue);

HRESULT CVIFUNC Excel_RecentFileOpen (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo,
                                      ExcelObj_Workbook *returnValue);

HRESULT CVIFUNC Excel_RecentFileDelete (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_RecentFilesItem (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo, long index,
                                       ExcelObj_RecentFile *returnValue);

HRESULT CVIFUNC Excel_RecentFiles_Default (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo, long index,
                                           ExcelObj_RecentFile *returnValue);

HRESULT CVIFUNC Excel_RecentFilesAdd (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, const char *name,
                                      ExcelObj_RecentFile *returnValue);

HRESULT CVIFUNC Excel_StyleDelete (CAObjHandle objectHandle,
                                   ERRORINFO *errorInfo, VARIANT *returnValue);

HRESULT CVIFUNC Excel_StylesAdd (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                 const char *name, VARIANT basedOn,
                                 ExcelObj_Style *returnValue);

HRESULT CVIFUNC Excel_StylesItem (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                  VARIANT index, ExcelObj_Style *returnValue);

HRESULT CVIFUNC Excel_StylesMerge (CAObjHandle objectHandle,
                                   ERRORINFO *errorInfo, VARIANT workbook,
                                   VARIANT *returnValue);

HRESULT CVIFUNC Excel_Styles_Default (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT index,
                                      ExcelObj_Style *returnValue);

HRESULT CVIFUNC Excel_BordersItem (CAObjHandle objectHandle,
                                   ERRORINFO *errorInfo,
                                   enum ExcelEnum_XlBordersIndex index,
                                   ExcelObj_Border *returnValue);

HRESULT CVIFUNC Excel_Borders_Default (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo,
                                       enum ExcelEnum_XlBordersIndex index,
                                       ExcelObj_Border *returnValue);

HRESULT CVIFUNC Excel_AddInsAdd (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                 const char *filename, VARIANT copyFile,
                                 ExcelObj_AddIn *returnValue);

HRESULT CVIFUNC Excel_AddInsItem (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                  VARIANT index, ExcelObj_AddIn *returnValue);

HRESULT CVIFUNC Excel_AddIns_Default (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT index,
                                      ExcelObj_AddIn *returnValue);

HRESULT CVIFUNC Excel_AreasItem (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                 long index, ExcelObj_Range *returnValue);

HRESULT CVIFUNC Excel_Areas_Default (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo, long index,
                                     ExcelObj_Range *returnValue);

HRESULT CVIFUNC Excel_ChartsAdd (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                 VARIANT before, VARIANT after, VARIANT count,
                                 ExcelObj_Chart *returnValue);

HRESULT CVIFUNC Excel_ChartsCopy (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                  VARIANT before, VARIANT after);

HRESULT CVIFUNC Excel_ChartsDelete (CAObjHandle objectHandle,
                                    ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_Charts_Dummy7 (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_ChartsItem (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                  VARIANT index, CAObjHandle *returnValue);

HRESULT CVIFUNC Excel_ChartsMove (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                  VARIANT before, VARIANT after);

HRESULT CVIFUNC Excel_Charts_PrintOut (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo, VARIANT from,
                                       VARIANT to, VARIANT copies,
                                       VARIANT preview, VARIANT activePrinter,
                                       VARIANT printToFile, VARIANT collate);

HRESULT CVIFUNC Excel_ChartsPrintPreview (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo,
                                          VARIANT enableChanges);

HRESULT CVIFUNC Excel_ChartsSelect (CAObjHandle objectHandle,
                                    ERRORINFO *errorInfo, VARIANT replace);

HRESULT CVIFUNC Excel_Charts_Default (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT index,
                                      CAObjHandle *returnValue);

HRESULT CVIFUNC Excel_ChartsPrintOut (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT from,
                                      VARIANT to, VARIANT copies,
                                      VARIANT preview, VARIANT activePrinter,
                                      VARIANT printToFile, VARIANT collate,
                                      VARIANT prToFileName);

HRESULT CVIFUNC Excel_PivotCacheRefresh (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_PivotCacheResetTimer (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_PivotCacheCreatePivotTable (CAObjHandle objectHandle,
                                                  ERRORINFO *errorInfo,
                                                  VARIANT tableDestination,
                                                  VARIANT tableName,
                                                  VARIANT readData,
                                                  ExcelObj_PivotTable *returnValue);

HRESULT CVIFUNC Excel_PivotCachesItem (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo, VARIANT index,
                                       ExcelObj_PivotCache *returnValue);

HRESULT CVIFUNC Excel_PivotCaches_Default (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo, VARIANT index,
                                           ExcelObj_PivotCache *returnValue);

HRESULT CVIFUNC Excel_PivotCaches_NewEnum (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo,
                                           LPUNKNOWN *returnValue);

HRESULT CVIFUNC Excel_PivotCachesAdd (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo,
                                      enum ExcelEnum_XlPivotTableSourceType sourceType,
                                      VARIANT sourceData,
                                      ExcelObj_PivotCache *returnValue);

HRESULT CVIFUNC Excel_PivotFormulaDelete (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_PivotFormulasAdd (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo,
                                        const char *formula,
                                        ExcelObj_PivotFormula *returnValue);

HRESULT CVIFUNC Excel_PivotFormulasItem (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo, VARIANT index,
                                         ExcelObj_PivotFormula *returnValue);

HRESULT CVIFUNC Excel_PivotFormulas_Default (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo, VARIANT index,
                                             ExcelObj_PivotFormula *returnValue);

HRESULT CVIFUNC Excel_PivotFormulas_NewEnum (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo,
                                             LPUNKNOWN *returnValue);

HRESULT CVIFUNC Excel_PivotTableAddFields (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo,
                                           VARIANT rowFields,
                                           VARIANT columnFields,
                                           VARIANT pageFields,
                                           VARIANT addToTable,
                                           VARIANT *returnValue);

HRESULT CVIFUNC Excel_PivotTableColumnFields (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo,
                                              VARIANT index,
                                              CAObjHandle *returnValue);

HRESULT CVIFUNC Excel_PivotTableShowPages (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo,
                                           VARIANT pageField,
                                           VARIANT *returnValue);

HRESULT CVIFUNC Excel_PivotTableDataFields (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo, VARIANT index,
                                            CAObjHandle *returnValue);

HRESULT CVIFUNC Excel_PivotTableHiddenFields (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo,
                                              VARIANT index,
                                              CAObjHandle *returnValue);

HRESULT CVIFUNC Excel_PivotTablePageFields (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo, VARIANT index,
                                            CAObjHandle *returnValue);

HRESULT CVIFUNC Excel_PivotTablePivotFields (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo, VARIANT index,
                                             CAObjHandle *returnValue);

HRESULT CVIFUNC Excel_PivotTableRefreshTable (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo,
                                              VBOOL *returnValue);

HRESULT CVIFUNC Excel_PivotTableRowFields (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo, VARIANT index,
                                           CAObjHandle *returnValue);

HRESULT CVIFUNC Excel_PivotTableVisibleFields (CAObjHandle objectHandle,
                                               ERRORINFO *errorInfo,
                                               VARIANT index,
                                               CAObjHandle *returnValue);

HRESULT CVIFUNC Excel_PivotTableCalculatedFields (CAObjHandle objectHandle,
                                                  ERRORINFO *errorInfo,
                                                  ExcelObj_CalculatedFields *returnValue);

HRESULT CVIFUNC Excel_PivotTableGetData (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo, const char *name,
                                         double *returnValue);

HRESULT CVIFUNC Excel_PivotTableListFormulas (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_PivotTablePivotCache (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo,
                                            ExcelObj_PivotCache *returnValue);

HRESULT CVIFUNC Excel_PivotTablePivotTableWizard (CAObjHandle objectHandle,
                                                  ERRORINFO *errorInfo,
                                                  VARIANT sourceType,
                                                  VARIANT sourceData,
                                                  VARIANT tableDestination,
                                                  VARIANT tableName,
                                                  VARIANT rowGrand,
                                                  VARIANT columnGrand,
                                                  VARIANT saveData,
                                                  VARIANT hasAutoFormat,
                                                  VARIANT autoPage,
                                                  VARIANT reserved,
                                                  VARIANT backgroundQuery,
                                                  VARIANT optimizeCache,
                                                  VARIANT pageFieldOrder,
                                                  VARIANT pageFieldWrapCount,
                                                  VARIANT readData,
                                                  VARIANT connection);

HRESULT CVIFUNC Excel_PivotTablePivotSelect (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo,
                                             const char *name,
                                             enum ExcelEnum_XlPTSelectionMode mode);

HRESULT CVIFUNC Excel_PivotTableUpdate (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_PivotTableFormat (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo,
                                        enum ExcelEnum_xlPivotFormatType format);

HRESULT CVIFUNC Excel_PivotTablesItem (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo, VARIANT index,
                                       ExcelObj_PivotTable *returnValue);

HRESULT CVIFUNC Excel_PivotTables_NewEnum (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo,
                                           LPUNKNOWN *returnValue);

HRESULT CVIFUNC Excel_PivotTablesAdd (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo,
                                      ExcelObj_PivotCache pivotCache,
                                      VARIANT tableDestination,
                                      VARIANT tableName, VARIANT readData,
                                      ExcelObj_PivotTable *returnValue);

HRESULT CVIFUNC Excel_PivotFieldChildItems (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo, VARIANT index,
                                            VARIANT *returnValue);

HRESULT CVIFUNC Excel_PivotFieldHiddenItems (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo, VARIANT index,
                                             VARIANT *returnValue);

HRESULT CVIFUNC Excel_PivotFieldParentItems (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo, VARIANT index,
                                             VARIANT *returnValue);

HRESULT CVIFUNC Excel_PivotFieldPivotItems (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo, VARIANT index,
                                            VARIANT *returnValue);

HRESULT CVIFUNC Excel_PivotFieldGetSubtotals (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo,
                                              VARIANT index,
                                              VARIANT *returnValue);

HRESULT CVIFUNC Excel_PivotFieldSetSubtotals (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo,
                                              VARIANT index, VARIANT newValue);

HRESULT CVIFUNC Excel_PivotFieldVisibleItems (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo,
                                              VARIANT index,
                                              VARIANT *returnValue);

HRESULT CVIFUNC Excel_PivotFieldCalculatedItems (CAObjHandle objectHandle,
                                                 ERRORINFO *errorInfo,
                                                 ExcelObj_CalculatedItems *returnValue);

HRESULT CVIFUNC Excel_PivotFieldDelete (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_PivotFieldAutoSort (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo, long order,
                                          const char *field);

HRESULT CVIFUNC Excel_PivotFieldAutoShow (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo, long type,
                                          long range, long count,
                                          const char *field);

HRESULT CVIFUNC Excel_PivotFieldsItem (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo, VARIANT index,
                                       CAObjHandle *returnValue);

HRESULT CVIFUNC Excel_PivotFields_NewEnum (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo,
                                           LPUNKNOWN *returnValue);

HRESULT CVIFUNC Excel_CalculatedFieldsAdd (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo,
                                           const char *name, const char *formula,
                                           ExcelObj_PivotField *returnValue);

HRESULT CVIFUNC Excel_CalculatedFieldsItem (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo, VARIANT index,
                                            ExcelObj_PivotField *returnValue);

HRESULT CVIFUNC Excel_CalculatedFields_Default (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo,
                                                VARIANT field,
                                                ExcelObj_PivotField *returnValue);

HRESULT CVIFUNC Excel_CalculatedFields_NewEnum (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo,
                                                LPUNKNOWN *returnValue);

HRESULT CVIFUNC Excel_PivotItemChildItems (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo, VARIANT index,
                                           VARIANT *returnValue);

HRESULT CVIFUNC Excel_PivotItemDelete (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_PivotItemsAdd (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo, const char *name);

HRESULT CVIFUNC Excel_PivotItemsItem (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT index,
                                      CAObjHandle *returnValue);

HRESULT CVIFUNC Excel_PivotItems_NewEnum (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo,
                                          LPUNKNOWN *returnValue);

HRESULT CVIFUNC Excel_CalculatedItemsAdd (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo, const char *name,
                                          const char *formula,
                                          ExcelObj_PivotItem *returnValue);

HRESULT CVIFUNC Excel_CalculatedItemsItem (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo, VARIANT index,
                                           ExcelObj_PivotItem *returnValue);

HRESULT CVIFUNC Excel_CalculatedItems_Default (CAObjHandle objectHandle,
                                               ERRORINFO *errorInfo,
                                               VARIANT field,
                                               ExcelObj_PivotItem *returnValue);

HRESULT CVIFUNC Excel_CalculatedItems_NewEnum (CAObjHandle objectHandle,
                                               ERRORINFO *errorInfo,
                                               LPUNKNOWN *returnValue);

HRESULT CVIFUNC Excel_CharactersDelete (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo,
                                        VARIANT *returnValue);

HRESULT CVIFUNC Excel_CharactersInsert (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo, const char *string,
                                        VARIANT *returnValue);

HRESULT CVIFUNC Excel_DialogsItem (CAObjHandle objectHandle,
                                   ERRORINFO *errorInfo,
                                   enum ExcelEnum_XlBuiltInDialog index,
                                   ExcelObj_Dialog *returnValue);

HRESULT CVIFUNC Excel_Dialogs_Default (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo,
                                       enum ExcelEnum_XlBuiltInDialog index,
                                       ExcelObj_Dialog *returnValue);

HRESULT CVIFUNC Excel_DialogShow (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                  VARIANT arg1, VARIANT arg2, VARIANT arg3,
                                  VARIANT arg4, VARIANT arg5, VARIANT arg6,
                                  VARIANT arg7, VARIANT arg8, VARIANT arg9,
                                  VARIANT arg10, VARIANT arg11, VARIANT arg12,
                                  VARIANT arg13, VARIANT arg14, VARIANT arg15,
                                  VARIANT arg16, VARIANT arg17, VARIANT arg18,
                                  VARIANT arg19, VARIANT arg20, VARIANT arg21,
                                  VARIANT arg22, VARIANT arg23, VARIANT arg24,
                                  VARIANT arg25, VARIANT arg26, VARIANT arg27,
                                  VARIANT arg28, VARIANT arg29, VARIANT arg30,
                                  VBOOL *returnValue);

HRESULT CVIFUNC Excel_SoundNoteDelete (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo,
                                       VARIANT *returnValue);

HRESULT CVIFUNC Excel_SoundNoteImport (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo,
                                       const char *filename,
                                       VARIANT *returnValue);

HRESULT CVIFUNC Excel_SoundNotePlay (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo, VARIANT *returnValue);

HRESULT CVIFUNC Excel_SoundNoteRecord (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo,
                                       VARIANT *returnValue);

HRESULT CVIFUNC Excel_PanesItem (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                 long index, ExcelObj_Pane *returnValue);

HRESULT CVIFUNC Excel_Panes_Default (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo, long index,
                                     ExcelObj_Pane *returnValue);

HRESULT CVIFUNC Excel_PaneActivate (CAObjHandle objectHandle,
                                    ERRORINFO *errorInfo, VBOOL *returnValue);

HRESULT CVIFUNC Excel_PaneLargeScroll (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo, VARIANT down,
                                       VARIANT up, VARIANT toRight,
                                       VARIANT toLeft, VARIANT *returnValue);

HRESULT CVIFUNC Excel_PaneSmallScroll (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo, VARIANT down,
                                       VARIANT up, VARIANT toRight,
                                       VARIANT toLeft, VARIANT *returnValue);

HRESULT CVIFUNC Excel_PaneScrollIntoView (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo, long left,
                                          long top, long width, long height,
                                          VARIANT start);

HRESULT CVIFUNC Excel_ScenariosAdd (CAObjHandle objectHandle,
                                    ERRORINFO *errorInfo, const char *name,
                                    VARIANT changingCells, VARIANT values,
                                    VARIANT comment, VARIANT locked,
                                    VARIANT hidden,
                                    ExcelObj_Scenario *returnValue);

HRESULT CVIFUNC Excel_ScenariosCreateSummary (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo,
                                              enum ExcelEnum_XlSummaryReportType reportType,
                                              VARIANT resultCells,
                                              VARIANT *returnValue);

HRESULT CVIFUNC Excel_ScenariosItem (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo, VARIANT index,
                                     ExcelObj_Scenario *returnValue);

HRESULT CVIFUNC Excel_ScenariosMerge (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT source,
                                      VARIANT *returnValue);

HRESULT CVIFUNC Excel_Scenarios_NewEnum (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo,
                                         LPUNKNOWN *returnValue);

HRESULT CVIFUNC Excel_ScenarioChangeScenario (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo,
                                              VARIANT changingCells,
                                              VARIANT values,
                                              VARIANT *returnValue);

HRESULT CVIFUNC Excel_ScenarioDelete (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT *returnValue);

HRESULT CVIFUNC Excel_ScenarioShow (CAObjHandle objectHandle,
                                    ERRORINFO *errorInfo, VARIANT *returnValue);

HRESULT CVIFUNC Excel_ScenarioValues (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT index,
                                      VARIANT *returnValue);

HRESULT CVIFUNC Excel_OLEObjects_Dummy3 (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_OLEObjectsBringToFront (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo,
                                              VARIANT *returnValue);

HRESULT CVIFUNC Excel_OLEObjectsCopy (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT *returnValue);

HRESULT CVIFUNC Excel_OLEObjectsCopyPicture (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo,
                                             enum ExcelEnum_XlPictureAppearance appearance,
                                             enum ExcelEnum_XlCopyPictureFormat format,
                                             VARIANT *returnValue);

HRESULT CVIFUNC Excel_OLEObjectsCut (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo, VARIANT *returnValue);

HRESULT CVIFUNC Excel_OLEObjectsDelete (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo,
                                        VARIANT *returnValue);

HRESULT CVIFUNC Excel_OLEObjectsDuplicate (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo,
                                           CAObjHandle *returnValue);

HRESULT CVIFUNC Excel_OLEObjects_Dummy12 (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_OLEObjects_Dummy15 (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_OLEObjectsSelect (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo, VARIANT replace,
                                        VARIANT *returnValue);

HRESULT CVIFUNC Excel_OLEObjectsSendToBack (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo,
                                            VARIANT *returnValue);

HRESULT CVIFUNC Excel_OLEObjects_Dummy22 (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_OLEObjects_Dummy30 (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_OLEObjects_Dummy32 (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_OLEObjects_Dummy33 (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_OLEObjects_Dummy34 (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_OLEObjects_Dummy36 (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_OLEObjects_Dummy37 (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_OLEObjects_Dummy38 (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_OLEObjects_Dummy39 (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_OLEObjects_Dummy40 (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_OLEObjects_Dummy41 (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_OLEObjectsAdd (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo, VARIANT classType,
                                     VARIANT filename, VARIANT link,
                                     VARIANT displayAsIcon, VARIANT iconFileName,
                                     VARIANT iconIndex, VARIANT iconLabel,
                                     VARIANT left, VARIANT top, VARIANT width,
                                     VARIANT height,
                                     ExcelObj_OLEObject *returnValue);

HRESULT CVIFUNC Excel_OLEObjectsGroup (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo,
                                       ExcelObj_GroupObject *returnValue);

HRESULT CVIFUNC Excel_OLEObjectsItem (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT index,
                                      CAObjHandle *returnValue);

HRESULT CVIFUNC Excel_OLEObjects_NewEnum (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo,
                                          LPUNKNOWN *returnValue);

HRESULT CVIFUNC Excel_RoutingSlipGetRecipients (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo,
                                                VARIANT index,
                                                VARIANT *returnValue);

HRESULT CVIFUNC Excel_RoutingSlipSetRecipients (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo,
                                                VARIANT index, VARIANT newValue);

HRESULT CVIFUNC Excel_RoutingSlipReset (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo,
                                        VARIANT *returnValue);

HRESULT CVIFUNC Excel_OutlineShowLevels (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo, VARIANT rowLevels,
                                         VARIANT columnLevels,
                                         VARIANT *returnValue);

HRESULT CVIFUNC Excel_WorksheetsAdd (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo, VARIANT before,
                                     VARIANT after, VARIANT count, VARIANT type,
                                     CAObjHandle *returnValue);

HRESULT CVIFUNC Excel_WorksheetsCopy (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT before,
                                      VARIANT after);

HRESULT CVIFUNC Excel_WorksheetsDelete (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_WorksheetsFillAcrossSheets (CAObjHandle objectHandle,
                                                  ERRORINFO *errorInfo,
                                                  ExcelObj_Range range,
                                                  enum ExcelEnum_XlFillWith type);

HRESULT CVIFUNC Excel_WorksheetsItem (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT index,
                                      CAObjHandle *returnValue);

HRESULT CVIFUNC Excel_WorksheetsMove (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT before,
                                      VARIANT after);

HRESULT CVIFUNC Excel_Worksheets_PrintOut (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo, VARIANT from,
                                           VARIANT to, VARIANT copies,
                                           VARIANT preview,
                                           VARIANT activePrinter,
                                           VARIANT printToFile, VARIANT collate);

HRESULT CVIFUNC Excel_WorksheetsPrintPreview (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo,
                                              VARIANT enableChanges);

HRESULT CVIFUNC Excel_WorksheetsSelect (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo, VARIANT replace);

HRESULT CVIFUNC Excel_Worksheets_Default (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo, VARIANT index,
                                          CAObjHandle *returnValue);

HRESULT CVIFUNC Excel_WorksheetsPrintOut (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo, VARIANT from,
                                          VARIANT to, VARIANT copies,
                                          VARIANT preview, VARIANT activePrinter,
                                          VARIANT printToFile, VARIANT collate,
                                          VARIANT prToFileName);

HRESULT CVIFUNC Excel_PageSetupGetPrintQuality (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo,
                                                VARIANT index,
                                                VARIANT *returnValue);

HRESULT CVIFUNC Excel_PageSetupSetPrintQuality (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo,
                                                VARIANT index, VARIANT newValue);

HRESULT CVIFUNC Excel_NamesAdd (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                VARIANT name, VARIANT refersTo, VARIANT visible,
                                VARIANT macroType, VARIANT shortcutKey,
                                VARIANT category, VARIANT nameLocal,
                                VARIANT refersToLocal, VARIANT categoryLocal,
                                VARIANT refersToR1C1, VARIANT refersToR1C1Local,
                                ExcelObj_Name *returnValue);

HRESULT CVIFUNC Excel_NamesItem (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                 VARIANT index, VARIANT indexLocal,
                                 VARIANT refersTo, ExcelObj_Name *returnValue);

HRESULT CVIFUNC Excel_Names_Default (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo, VARIANT index,
                                     VARIANT indexLocal, VARIANT refersTo,
                                     ExcelObj_Name *returnValue);

HRESULT CVIFUNC Excel_NameDelete (CAObjHandle objectHandle, ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_ChartObjectBringToFront (CAObjHandle objectHandle,
                                               ERRORINFO *errorInfo,
                                               VARIANT *returnValue);

HRESULT CVIFUNC Excel_ChartObjectCopy (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo,
                                       VARIANT *returnValue);

HRESULT CVIFUNC Excel_ChartObjectCopyPicture (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo,
                                              enum ExcelEnum_XlPictureAppearance appearance,
                                              enum ExcelEnum_XlCopyPictureFormat format,
                                              VARIANT *returnValue);

HRESULT CVIFUNC Excel_ChartObjectCut (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT *returnValue);

HRESULT CVIFUNC Excel_ChartObjectDelete (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo,
                                         VARIANT *returnValue);

HRESULT CVIFUNC Excel_ChartObjectDuplicate (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo,
                                            CAObjHandle *returnValue);

HRESULT CVIFUNC Excel_ChartObjectSelect (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo, VARIANT replace,
                                         VARIANT *returnValue);

HRESULT CVIFUNC Excel_ChartObjectSendToBack (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo,
                                             VARIANT *returnValue);

HRESULT CVIFUNC Excel_ChartObjectActivate (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo,
                                           VARIANT *returnValue);

HRESULT CVIFUNC Excel_ChartObjects_Dummy3 (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_ChartObjectsBringToFront (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo,
                                                VARIANT *returnValue);

HRESULT CVIFUNC Excel_ChartObjectsCopy (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo,
                                        VARIANT *returnValue);

HRESULT CVIFUNC Excel_ChartObjectsCopyPicture (CAObjHandle objectHandle,
                                               ERRORINFO *errorInfo,
                                               enum ExcelEnum_XlPictureAppearance appearance,
                                               enum ExcelEnum_XlCopyPictureFormat format,
                                               VARIANT *returnValue);

HRESULT CVIFUNC Excel_ChartObjectsCut (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo,
                                       VARIANT *returnValue);

HRESULT CVIFUNC Excel_ChartObjectsDelete (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo,
                                          VARIANT *returnValue);

HRESULT CVIFUNC Excel_ChartObjectsDuplicate (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo,
                                             CAObjHandle *returnValue);

HRESULT CVIFUNC Excel_ChartObjects_Dummy12 (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_ChartObjects_Dummy15 (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_ChartObjectsSelect (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo, VARIANT replace,
                                          VARIANT *returnValue);

HRESULT CVIFUNC Excel_ChartObjectsSendToBack (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo,
                                              VARIANT *returnValue);

HRESULT CVIFUNC Excel_ChartObjects_Dummy22 (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_ChartObjects_Dummy25 (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_ChartObjects_Dummy27 (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_ChartObjects_Dummy28 (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_ChartObjects_Dummy29 (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_ChartObjectsAdd (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo, double left,
                                       double top, double width, double height,
                                       ExcelObj_ChartObject *returnValue);

HRESULT CVIFUNC Excel_ChartObjectsGroup (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo,
                                         ExcelObj_GroupObject *returnValue);

HRESULT CVIFUNC Excel_ChartObjectsItem (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo, VARIANT index,
                                        CAObjHandle *returnValue);

HRESULT CVIFUNC Excel_ChartObjects_NewEnum (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo,
                                            LPUNKNOWN *returnValue);

HRESULT CVIFUNC Excel_CustomViewsItem (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo, VARIANT viewName,
                                       ExcelObj_CustomView *returnValue);

HRESULT CVIFUNC Excel_CustomViewsAdd (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, const char *viewName,
                                      VARIANT printSettings,
                                      VARIANT rowColSettings,
                                      ExcelObj_CustomView *returnValue);

HRESULT CVIFUNC Excel_CustomViews_Default (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo,
                                           VARIANT viewName,
                                           ExcelObj_CustomView *returnValue);

HRESULT CVIFUNC Excel_CustomViewShow (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_CustomViewDelete (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_FormatConditionsItem (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo, VARIANT index,
                                            ExcelObj_FormatCondition *returnValue);

HRESULT CVIFUNC Excel_FormatConditionsAdd (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo,
                                           enum ExcelEnum_XlFormatConditionType type,
                                           VARIANT operator, VARIANT formula1,
                                           VARIANT formula2,
                                           ExcelObj_FormatCondition *returnValue);

HRESULT CVIFUNC Excel_FormatConditions_Default (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo,
                                                VARIANT index,
                                                ExcelObj_FormatCondition *returnValue);

HRESULT CVIFUNC Excel_FormatConditionsDelete (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_FormatConditionModify (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo,
                                             enum ExcelEnum_XlFormatConditionType type,
                                             VARIANT operator, VARIANT formula1,
                                             VARIANT formula2);

HRESULT CVIFUNC Excel_FormatConditionDelete (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_CommentsItem (CAObjHandle objectHandle,
                                    ERRORINFO *errorInfo, long index,
                                    ExcelObj_Comment *returnValue);

HRESULT CVIFUNC Excel_Comments_Default (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo, long index,
                                        ExcelObj_Comment *returnValue);

HRESULT CVIFUNC Excel_CommentText (CAObjHandle objectHandle,
                                   ERRORINFO *errorInfo, VARIANT text,
                                   VARIANT start, VARIANT overwrite,
                                   char **returnValue);

HRESULT CVIFUNC Excel_CommentDelete (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_CommentNext (CAObjHandle objectHandle,
                                   ERRORINFO *errorInfo,
                                   ExcelObj_Comment *returnValue);

HRESULT CVIFUNC Excel_CommentPrevious (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo,
                                       ExcelObj_Comment *returnValue);

HRESULT CVIFUNC Excel_QueryTablesAdd (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT connection,
                                      ExcelObj_Range destination, VARIANT sql,
                                      ExcelObj_QueryTable *returnValue);

HRESULT CVIFUNC Excel_QueryTablesItem (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo, VARIANT index,
                                       ExcelObj_QueryTable *returnValue);

HRESULT CVIFUNC Excel_QueryTables_Default (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo, VARIANT index,
                                           ExcelObj_QueryTable *returnValue);

HRESULT CVIFUNC Excel_QueryTables_NewEnum (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo,
                                           LPUNKNOWN *returnValue);

HRESULT CVIFUNC Excel_ParameterSetParam (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo,
                                         enum ExcelEnum_XlParameterType type,
                                         VARIANT value);

HRESULT CVIFUNC Excel_ParametersAdd (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo, const char *name,
                                     VARIANT iDataType,
                                     ExcelObj_Parameter *returnValue);

HRESULT CVIFUNC Excel_ParametersItem (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT index,
                                      ExcelObj_Parameter *returnValue);

HRESULT CVIFUNC Excel_Parameters_Default (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo, VARIANT index,
                                          ExcelObj_Parameter *returnValue);

HRESULT CVIFUNC Excel_ParametersDelete (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_Parameters_NewEnum (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo,
                                          LPUNKNOWN *returnValue);

HRESULT CVIFUNC Excel_ODBCErrorsItem (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, long index,
                                      ExcelObj_ODBCError *returnValue);

HRESULT CVIFUNC Excel_ODBCErrors_Default (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo, long index,
                                          ExcelObj_ODBCError *returnValue);

HRESULT CVIFUNC Excel_ODBCErrors_NewEnum (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo,
                                          LPUNKNOWN *returnValue);

HRESULT CVIFUNC Excel_ValidationAdd (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo,
                                     enum ExcelEnum_XlDVType type,
                                     VARIANT alertStyle, VARIANT operator,
                                     VARIANT formula1, VARIANT formula2);

HRESULT CVIFUNC Excel_ValidationDelete (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_ValidationModify (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo, VARIANT type,
                                        VARIANT alertStyle, VARIANT operator,
                                        VARIANT formula1, VARIANT formula2);

HRESULT CVIFUNC Excel_HyperlinksAdd (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo, CAObjHandle anchor,
                                     const char *address, VARIANT subAddress,
                                     VARIANT screenTip, VARIANT textToDisplay,
                                     CAObjHandle *returnValue);

HRESULT CVIFUNC Excel_HyperlinksItem (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT index,
                                      ExcelObj_Hyperlink *returnValue);

HRESULT CVIFUNC Excel_Hyperlinks_Default (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo, VARIANT index,
                                          ExcelObj_Hyperlink *returnValue);

HRESULT CVIFUNC Excel_HyperlinksDelete (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_HyperlinkAddToFavorites (CAObjHandle objectHandle,
                                               ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_HyperlinkDelete (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_HyperlinkFollow (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo, VARIANT newWindow,
                                       VARIANT addHistory, VARIANT extraInfo,
                                       VARIANT method, VARIANT headerInfo);

HRESULT CVIFUNC Excel_HyperlinkCreateNewDocument (CAObjHandle objectHandle,
                                                  ERRORINFO *errorInfo,
                                                  const char *filename,
                                                  VBOOL editNow, VBOOL overwrite);

HRESULT CVIFUNC Excel_Filters_Default (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo, long index,
                                       ExcelObj_Filter *returnValue);

HRESULT CVIFUNC Excel_FiltersItem (CAObjHandle objectHandle,
                                   ERRORINFO *errorInfo, long index,
                                   ExcelObj_Filter *returnValue);

HRESULT CVIFUNC Excel_AutoCorrectAddReplacement (CAObjHandle objectHandle,
                                                 ERRORINFO *errorInfo,
                                                 const char *what,
                                                 const char *replacement,
                                                 VARIANT *returnValue);

HRESULT CVIFUNC Excel_AutoCorrectDeleteReplacement (CAObjHandle objectHandle,
                                                    ERRORINFO *errorInfo,
                                                    const char *what,
                                                    VARIANT *returnValue);

HRESULT CVIFUNC Excel_AutoCorrectGetReplacementList (CAObjHandle objectHandle,
                                                     ERRORINFO *errorInfo,
                                                     VARIANT index,
                                                     VARIANT *returnValue);

HRESULT CVIFUNC Excel_AutoCorrectSetReplacementList (CAObjHandle objectHandle,
                                                     ERRORINFO *errorInfo,
                                                     VARIANT index,
                                                     VARIANT newValue);

HRESULT CVIFUNC Excel_ChartFillFormatOneColorGradient (CAObjHandle objectHandle,
                                                       ERRORINFO *errorInfo,
                                                       enum ExcelEnum_MsoGradientStyle style,
                                                       long variant,
                                                       float degree);

HRESULT CVIFUNC Excel_ChartFillFormatTwoColorGradient (CAObjHandle objectHandle,
                                                       ERRORINFO *errorInfo,
                                                       enum ExcelEnum_MsoGradientStyle style,
                                                       long variant);

HRESULT CVIFUNC Excel_ChartFillFormatPresetTextured (CAObjHandle objectHandle,
                                                     ERRORINFO *errorInfo,
                                                     enum ExcelEnum_MsoPresetTexture presetTexture);

HRESULT CVIFUNC Excel_ChartFillFormatSolid (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_ChartFillFormatPatterned (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo,
                                                enum ExcelEnum_MsoPatternType pattern);

HRESULT CVIFUNC Excel_ChartFillFormatUserPicture (CAObjHandle objectHandle,
                                                  ERRORINFO *errorInfo,
                                                  VARIANT pictureFile,
                                                  VARIANT pictureFormat,
                                                  VARIANT pictureStackUnit,
                                                  VARIANT picturePlacement);

HRESULT CVIFUNC Excel_ChartFillFormatUserTextured (CAObjHandle objectHandle,
                                                   ERRORINFO *errorInfo,
                                                   const char *textureFile);

HRESULT CVIFUNC Excel_ChartFillFormatPresetGradient (CAObjHandle objectHandle,
                                                     ERRORINFO *errorInfo,
                                                     enum ExcelEnum_MsoGradientStyle style,
                                                     long variant,
                                                     enum ExcelEnum_MsoPresetGradientType presetGradientType);

HRESULT CVIFUNC Excel_AxisDelete (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                  VARIANT *returnValue);

HRESULT CVIFUNC Excel_AxisSelect (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                  VARIANT *returnValue);

HRESULT CVIFUNC Excel_ChartTitleSelect (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo,
                                        VARIANT *returnValue);

HRESULT CVIFUNC Excel_ChartTitleDelete (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo,
                                        VARIANT *returnValue);

HRESULT CVIFUNC Excel_ChartTitleCharacters (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo, VARIANT start,
                                            VARIANT length,
                                            ExcelObj_Characters *returnValue);

HRESULT CVIFUNC Excel_AxisTitleSelect (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo,
                                       VARIANT *returnValue);

HRESULT CVIFUNC Excel_AxisTitleDelete (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo,
                                       VARIANT *returnValue);

HRESULT CVIFUNC Excel_AxisTitleCharacters (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo, VARIANT start,
                                           VARIANT length,
                                           ExcelObj_Characters *returnValue);

HRESULT CVIFUNC Excel_ChartGroupSeriesCollection (CAObjHandle objectHandle,
                                                  ERRORINFO *errorInfo,
                                                  VARIANT index,
                                                  CAObjHandle *returnValue);

HRESULT CVIFUNC Excel_ChartGroupsItem (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo, VARIANT index,
                                       ExcelObj_ChartGroup *returnValue);

HRESULT CVIFUNC Excel_ChartGroups_NewEnum (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo,
                                           LPUNKNOWN *returnValue);

HRESULT CVIFUNC Excel_AxesItem (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                enum ExcelEnum_XlAxisType type,
                                enum ExcelEnum_XlAxisGroup axisGroup,
                                ExcelObj_Axis *returnValue);

HRESULT CVIFUNC Excel_Axes_NewEnum (CAObjHandle objectHandle,
                                    ERRORINFO *errorInfo, LPUNKNOWN *returnValue);

HRESULT CVIFUNC Excel_PointsItem (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                  long index, ExcelObj_Point *returnValue);

HRESULT CVIFUNC Excel_Points_NewEnum (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo,
                                      LPUNKNOWN *returnValue);

HRESULT CVIFUNC Excel_PointApplyDataLabels (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo,
                                            enum ExcelEnum_XlDataLabelsType type,
                                            VARIANT legendKey, VARIANT autoText,
                                            VARIANT *returnValue);

HRESULT CVIFUNC Excel_PointClearFormats (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo,
                                         VARIANT *returnValue);

HRESULT CVIFUNC Excel_PointCopy (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                 VARIANT *returnValue);

HRESULT CVIFUNC Excel_PointDelete (CAObjHandle objectHandle,
                                   ERRORINFO *errorInfo, VARIANT *returnValue);

HRESULT CVIFUNC Excel_PointPaste (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                  VARIANT *returnValue);

HRESULT CVIFUNC Excel_PointSelect (CAObjHandle objectHandle,
                                   ERRORINFO *errorInfo, VARIANT *returnValue);

HRESULT CVIFUNC Excel_SeriesApplyDataLabels (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo,
                                             enum ExcelEnum_XlDataLabelsType type,
                                             VARIANT legendKey, VARIANT autoText,
                                             VARIANT hasLeaderLines,
                                             VARIANT *returnValue);

HRESULT CVIFUNC Excel_SeriesClearFormats (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo,
                                          VARIANT *returnValue);

HRESULT CVIFUNC Excel_SeriesCopy (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                  VARIANT *returnValue);

HRESULT CVIFUNC Excel_SeriesDataLabels (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo, VARIANT index,
                                        CAObjHandle *returnValue);

HRESULT CVIFUNC Excel_SeriesDelete (CAObjHandle objectHandle,
                                    ERRORINFO *errorInfo, VARIANT *returnValue);

HRESULT CVIFUNC Excel_SeriesErrorBar (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo,
                                      enum ExcelEnum_XlErrorBarDirection direction,
                                      enum ExcelEnum_XlErrorBarInclude include,
                                      enum ExcelEnum_XlErrorBarType type,
                                      VARIANT amount, VARIANT minusValues,
                                      VARIANT *returnValue);

HRESULT CVIFUNC Excel_SeriesPaste (CAObjHandle objectHandle,
                                   ERRORINFO *errorInfo, VARIANT *returnValue);

HRESULT CVIFUNC Excel_SeriesPoints (CAObjHandle objectHandle,
                                    ERRORINFO *errorInfo, VARIANT index,
                                    CAObjHandle *returnValue);

HRESULT CVIFUNC Excel_SeriesSelect (CAObjHandle objectHandle,
                                    ERRORINFO *errorInfo, VARIANT *returnValue);

HRESULT CVIFUNC Excel_SeriesTrendlines (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo, VARIANT index,
                                        CAObjHandle *returnValue);

HRESULT CVIFUNC Excel_SeriesApplyCustomType (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo,
                                             enum ExcelEnum_XlChartType chartType);

HRESULT CVIFUNC Excel_SeriesCollectionAdd (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo, VARIANT source,
                                           enum ExcelEnum_XlRowCol rowcol,
                                           VARIANT seriesLabels,
                                           VARIANT categoryLabels,
                                           VARIANT replace,
                                           ExcelObj_Series *returnValue);

HRESULT CVIFUNC Excel_SeriesCollectionExtend (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo,
                                              VARIANT source, VARIANT rowcol,
                                              VARIANT categoryLabels,
                                              VARIANT *returnValue);

HRESULT CVIFUNC Excel_SeriesCollectionItem (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo, VARIANT index,
                                            ExcelObj_Series *returnValue);

HRESULT CVIFUNC Excel_SeriesCollection_NewEnum (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo,
                                                LPUNKNOWN *returnValue);

HRESULT CVIFUNC Excel_SeriesCollectionPaste (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo,
                                             enum ExcelEnum_XlRowCol rowcol,
                                             VARIANT seriesLabels,
                                             VARIANT categoryLabels,
                                             VARIANT replace, VARIANT newSeries,
                                             VARIANT *returnValue);

HRESULT CVIFUNC Excel_SeriesCollectionNewSeries (CAObjHandle objectHandle,
                                                 ERRORINFO *errorInfo,
                                                 ExcelObj_Series *returnValue);

HRESULT CVIFUNC Excel_DataLabelSelect (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo,
                                       VARIANT *returnValue);

HRESULT CVIFUNC Excel_DataLabelDelete (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo,
                                       VARIANT *returnValue);

HRESULT CVIFUNC Excel_DataLabelCharacters (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo, VARIANT start,
                                           VARIANT length,
                                           ExcelObj_Characters *returnValue);

HRESULT CVIFUNC Excel_DataLabelsSelect (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo,
                                        VARIANT *returnValue);

HRESULT CVIFUNC Excel_DataLabelsDelete (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo,
                                        VARIANT *returnValue);

HRESULT CVIFUNC Excel_DataLabels_Dummy9 (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_DataLabels_Dummy10 (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_DataLabels_Dummy13 (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_DataLabels_Dummy16 (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_DataLabels_Dummy17 (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_DataLabelsItem (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT index,
                                      ExcelObj_DataLabel *returnValue);

HRESULT CVIFUNC Excel_DataLabels_NewEnum (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo,
                                          LPUNKNOWN *returnValue);

HRESULT CVIFUNC Excel_LegendEntryDelete (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo,
                                         VARIANT *returnValue);

HRESULT CVIFUNC Excel_LegendEntrySelect (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo,
                                         VARIANT *returnValue);

HRESULT CVIFUNC Excel_LegendEntriesItem (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo, VARIANT index,
                                         ExcelObj_LegendEntry *returnValue);

HRESULT CVIFUNC Excel_LegendEntries_NewEnum (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo,
                                             LPUNKNOWN *returnValue);

HRESULT CVIFUNC Excel_LegendKeyClearFormats (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo,
                                             VARIANT *returnValue);

HRESULT CVIFUNC Excel_LegendKeyDelete (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo,
                                       VARIANT *returnValue);

HRESULT CVIFUNC Excel_LegendKeySelect (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo,
                                       VARIANT *returnValue);

HRESULT CVIFUNC Excel_TrendlinesAdd (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo,
                                     enum ExcelEnum_XlTrendlineType type,
                                     VARIANT order, VARIANT period,
                                     VARIANT forward, VARIANT backward,
                                     VARIANT intercept, VARIANT displayEquation,
                                     VARIANT displayRSquared, VARIANT name,
                                     ExcelObj_Trendline *returnValue);

HRESULT CVIFUNC Excel_TrendlinesItem (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT index,
                                      ExcelObj_Trendline *returnValue);

HRESULT CVIFUNC Excel_Trendlines_NewEnum (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo,
                                          LPUNKNOWN *returnValue);

HRESULT CVIFUNC Excel_TrendlineClearFormats (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo,
                                             VARIANT *returnValue);

HRESULT CVIFUNC Excel_TrendlineDelete (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo,
                                       VARIANT *returnValue);

HRESULT CVIFUNC Excel_TrendlineSelect (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo,
                                       VARIANT *returnValue);

HRESULT CVIFUNC Excel_CornersSelect (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo, VARIANT *returnValue);

HRESULT CVIFUNC Excel_SeriesLinesSelect (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo,
                                         VARIANT *returnValue);

HRESULT CVIFUNC Excel_SeriesLinesDelete (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo,
                                         VARIANT *returnValue);

HRESULT CVIFUNC Excel_HiLoLinesSelect (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo,
                                       VARIANT *returnValue);

HRESULT CVIFUNC Excel_HiLoLinesDelete (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo,
                                       VARIANT *returnValue);

HRESULT CVIFUNC Excel_GridlinesSelect (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo,
                                       VARIANT *returnValue);

HRESULT CVIFUNC Excel_GridlinesDelete (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo,
                                       VARIANT *returnValue);

HRESULT CVIFUNC Excel_DropLinesSelect (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo,
                                       VARIANT *returnValue);

HRESULT CVIFUNC Excel_DropLinesDelete (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo,
                                       VARIANT *returnValue);

HRESULT CVIFUNC Excel_LeaderLinesDelete (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_LeaderLinesSelect (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_UpBarsSelect (CAObjHandle objectHandle,
                                    ERRORINFO *errorInfo, VARIANT *returnValue);

HRESULT CVIFUNC Excel_UpBarsDelete (CAObjHandle objectHandle,
                                    ERRORINFO *errorInfo, VARIANT *returnValue);

HRESULT CVIFUNC Excel_DownBarsSelect (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT *returnValue);

HRESULT CVIFUNC Excel_DownBarsDelete (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT *returnValue);

HRESULT CVIFUNC Excel_FloorSelect (CAObjHandle objectHandle,
                                   ERRORINFO *errorInfo, VARIANT *returnValue);

HRESULT CVIFUNC Excel_FloorClearFormats (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo,
                                         VARIANT *returnValue);

HRESULT CVIFUNC Excel_FloorPaste (CAObjHandle objectHandle, ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_WallsSelect (CAObjHandle objectHandle,
                                   ERRORINFO *errorInfo, VARIANT *returnValue);

HRESULT CVIFUNC Excel_WallsClearFormats (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo,
                                         VARIANT *returnValue);

HRESULT CVIFUNC Excel_WallsPaste (CAObjHandle objectHandle, ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_TickLabelsDelete (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo,
                                        VARIANT *returnValue);

HRESULT CVIFUNC Excel_TickLabelsSelect (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo,
                                        VARIANT *returnValue);

HRESULT CVIFUNC Excel_PlotAreaSelect (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT *returnValue);

HRESULT CVIFUNC Excel_PlotAreaClearFormats (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo,
                                            VARIANT *returnValue);

HRESULT CVIFUNC Excel_ChartAreaSelect (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo,
                                       VARIANT *returnValue);

HRESULT CVIFUNC Excel_ChartAreaClear (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT *returnValue);

HRESULT CVIFUNC Excel_ChartAreaClearContents (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo,
                                              VARIANT *returnValue);

HRESULT CVIFUNC Excel_ChartAreaCopy (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo, VARIANT *returnValue);

HRESULT CVIFUNC Excel_ChartAreaClearFormats (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo,
                                             VARIANT *returnValue);

HRESULT CVIFUNC Excel_LegendSelect (CAObjHandle objectHandle,
                                    ERRORINFO *errorInfo, VARIANT *returnValue);

HRESULT CVIFUNC Excel_LegendDelete (CAObjHandle objectHandle,
                                    ERRORINFO *errorInfo, VARIANT *returnValue);

HRESULT CVIFUNC Excel_LegendLegendEntries (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo, VARIANT index,
                                           CAObjHandle *returnValue);

HRESULT CVIFUNC Excel_LegendClear (CAObjHandle objectHandle,
                                   ERRORINFO *errorInfo, VARIANT *returnValue);

HRESULT CVIFUNC Excel_ErrorBarsSelect (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo,
                                       VARIANT *returnValue);

HRESULT CVIFUNC Excel_ErrorBarsDelete (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo,
                                       VARIANT *returnValue);

HRESULT CVIFUNC Excel_ErrorBarsClearFormats (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo,
                                             VARIANT *returnValue);

HRESULT CVIFUNC Excel_DataTableSelect (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_DataTableDelete (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_ShapeApply (CAObjHandle objectHandle, ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_ShapeDelete (CAObjHandle objectHandle,
                                   ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_ShapeDuplicate (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo,
                                      ExcelObj_Shape *returnValue);

HRESULT CVIFUNC Excel_ShapeFlip (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                 enum ExcelEnum_MsoFlipCmd flipCmd);

HRESULT CVIFUNC Excel_ShapeIncrementLeft (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo, float increment);

HRESULT CVIFUNC Excel_ShapeIncrementRotation (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo,
                                              float increment);

HRESULT CVIFUNC Excel_ShapeIncrementTop (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo, float increment);

HRESULT CVIFUNC Excel_ShapePickUp (CAObjHandle objectHandle,
                                   ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_ShapeRerouteConnections (CAObjHandle objectHandle,
                                               ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_ShapeScaleHeight (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo, float factor,
                                        enum ExcelEnum_MsoTriState relativeToOriginalSize,
                                        VARIANT scale);

HRESULT CVIFUNC Excel_ShapeScaleWidth (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo, float factor,
                                       enum ExcelEnum_MsoTriState relativeToOriginalSize,
                                       VARIANT scale);

HRESULT CVIFUNC Excel_ShapeSelect (CAObjHandle objectHandle,
                                   ERRORINFO *errorInfo, VARIANT replace);

HRESULT CVIFUNC Excel_ShapeSetShapesDefaultProperties (CAObjHandle objectHandle,
                                                       ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_ShapeUngroup (CAObjHandle objectHandle,
                                    ERRORINFO *errorInfo,
                                    ExcelObj_ShapeRange *returnValue);

HRESULT CVIFUNC Excel_ShapeZOrder (CAObjHandle objectHandle,
                                   ERRORINFO *errorInfo,
                                   enum ExcelEnum_MsoZOrderCmd ZOrderCmd);

HRESULT CVIFUNC Excel_ShapeCopy (CAObjHandle objectHandle, ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_ShapeCut (CAObjHandle objectHandle, ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_ShapeCopyPicture (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo, VARIANT appearance,
                                        VARIANT format);

HRESULT CVIFUNC Excel_ShapesItem (CAObjHandle objectHandle, ERRORINFO *errorInfo,
                                  VARIANT index, ExcelObj_Shape *returnValue);

HRESULT CVIFUNC Excel_Shapes_Default (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT index,
                                      ExcelObj_Shape *returnValue);

HRESULT CVIFUNC Excel_ShapesAddCallout (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo,
                                        enum ExcelEnum_MsoCalloutType type,
                                        float left, float top, float width,
                                        float height,
                                        ExcelObj_Shape *returnValue);

HRESULT CVIFUNC Excel_ShapesAddConnector (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo,
                                          enum ExcelEnum_MsoConnectorType type,
                                          float beginX, float beginY, float endX,
                                          float endY,
                                          ExcelObj_Shape *returnValue);

HRESULT CVIFUNC Excel_ShapesAddCurve (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo,
                                      VARIANT safeArrayOfPoints,
                                      ExcelObj_Shape *returnValue);

HRESULT CVIFUNC Excel_ShapesAddLabel (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo,
                                      enum ExcelEnum_MsoTextOrientation orientation,
                                      float left, float top, float width,
                                      float height, ExcelObj_Shape *returnValue);

HRESULT CVIFUNC Excel_ShapesAddLine (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo, float beginX,
                                     float beginY, float endX, float endY,
                                     ExcelObj_Shape *returnValue);

HRESULT CVIFUNC Excel_ShapesAddPicture (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo,
                                        const char *filename,
                                        enum ExcelEnum_MsoTriState linkToFile,
                                        enum ExcelEnum_MsoTriState saveWithDocument,
                                        float left, float top, float width,
                                        float height,
                                        ExcelObj_Shape *returnValue);

HRESULT CVIFUNC Excel_ShapesAddPolyline (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo,
                                         VARIANT safeArrayOfPoints,
                                         ExcelObj_Shape *returnValue);

HRESULT CVIFUNC Excel_ShapesAddShape (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo,
                                      enum ExcelEnum_MsoAutoShapeType type,
                                      float left, float top, float width,
                                      float height, ExcelObj_Shape *returnValue);

HRESULT CVIFUNC Excel_ShapesAddTextEffect (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo,
                                           enum ExcelEnum_MsoPresetTextEffect presetTextEffect,
                                           const char *text,
                                           const char *fontName, float fontSize,
                                           enum ExcelEnum_MsoTriState fontBold,
                                           enum ExcelEnum_MsoTriState fontItalic,
                                           float left, float top,
                                           ExcelObj_Shape *returnValue);

HRESULT CVIFUNC Excel_ShapesAddTextbox (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo,
                                        enum ExcelEnum_MsoTextOrientation orientation,
                                        float left, float top, float width,
                                        float height,
                                        ExcelObj_Shape *returnValue);

HRESULT CVIFUNC Excel_ShapesBuildFreeform (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo,
                                           enum ExcelEnum_MsoEditingType editingType,
                                           float x1, float y1,
                                           ExcelObj_FreeformBuilder *returnValue);

HRESULT CVIFUNC Excel_ShapesRange (CAObjHandle objectHandle,
                                   ERRORINFO *errorInfo, VARIANT index,
                                   ExcelObj_ShapeRange *returnValue);

HRESULT CVIFUNC Excel_ShapesSelectAll (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_ShapesAddFormControl (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo,
                                            enum ExcelEnum_XlFormControl type,
                                            long left, long top, long width,
                                            long height,
                                            ExcelObj_Shape *returnValue);

HRESULT CVIFUNC Excel_ShapesAddOLEObject (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo,
                                          VARIANT classType, VARIANT filename,
                                          VARIANT link, VARIANT displayAsIcon,
                                          VARIANT iconFileName,
                                          VARIANT iconIndex, VARIANT iconLabel,
                                          VARIANT left, VARIANT top,
                                          VARIANT width, VARIANT height,
                                          ExcelObj_Shape *returnValue);

HRESULT CVIFUNC Excel_ShapeRangeItem (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo, VARIANT index,
                                      ExcelObj_Shape *returnValue);

HRESULT CVIFUNC Excel_ShapeRange_Default (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo, VARIANT index,
                                          ExcelObj_Shape *returnValue);

HRESULT CVIFUNC Excel_ShapeRangeAlign (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo,
                                       enum ExcelEnum_MsoAlignCmd alignCmd,
                                       enum ExcelEnum_MsoTriState relativeTo);

HRESULT CVIFUNC Excel_ShapeRangeApply (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_ShapeRangeDelete (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_ShapeRangeDistribute (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo,
                                            enum ExcelEnum_MsoDistributeCmd distributeCmd,
                                            enum ExcelEnum_MsoTriState relativeTo);

HRESULT CVIFUNC Excel_ShapeRangeDuplicate (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo,
                                           ExcelObj_ShapeRange *returnValue);

HRESULT CVIFUNC Excel_ShapeRangeFlip (CAObjHandle objectHandle,
                                      ERRORINFO *errorInfo,
                                      enum ExcelEnum_MsoFlipCmd flipCmd);

HRESULT CVIFUNC Excel_ShapeRangeIncrementLeft (CAObjHandle objectHandle,
                                               ERRORINFO *errorInfo,
                                               float increment);

HRESULT CVIFUNC Excel_ShapeRangeIncrementRotation (CAObjHandle objectHandle,
                                                   ERRORINFO *errorInfo,
                                                   float increment);

HRESULT CVIFUNC Excel_ShapeRangeIncrementTop (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo,
                                              float increment);

HRESULT CVIFUNC Excel_ShapeRangeGroup (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo,
                                       ExcelObj_Shape *returnValue);

HRESULT CVIFUNC Excel_ShapeRangePickUp (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_ShapeRangeRerouteConnections (CAObjHandle objectHandle,
                                                    ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_ShapeRangeRegroup (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo,
                                         ExcelObj_Shape *returnValue);

HRESULT CVIFUNC Excel_ShapeRangeScaleHeight (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo, float factor,
                                             enum ExcelEnum_MsoTriState relativeToOriginalSize,
                                             VARIANT scale);

HRESULT CVIFUNC Excel_ShapeRangeScaleWidth (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo, float factor,
                                            enum ExcelEnum_MsoTriState relativeToOriginalSize,
                                            VARIANT scale);

HRESULT CVIFUNC Excel_ShapeRangeSelect (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo, VARIANT replace);

HRESULT CVIFUNC Excel_ShapeRangeSetShapesDefaultProps (CAObjHandle objectHandle,
                                                       ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_ShapeRangeUngroup (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo,
                                         ExcelObj_ShapeRange *returnValue);

HRESULT CVIFUNC Excel_ShapeRangeZOrder (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo,
                                        enum ExcelEnum_MsoZOrderCmd ZOrderCmd);

HRESULT CVIFUNC Excel_GroupShapesItem (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo, VARIANT index,
                                       ExcelObj_Shape *returnValue);

HRESULT CVIFUNC Excel_GroupShapes_Default (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo, VARIANT index,
                                           ExcelObj_Shape *returnValue);

HRESULT CVIFUNC Excel_TextFrameCharacters (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo, VARIANT start,
                                           VARIANT length,
                                           ExcelObj_Characters *returnValue);

HRESULT CVIFUNC Excel_ConnectorFormatBeginConnect (CAObjHandle objectHandle,
                                                   ERRORINFO *errorInfo,
                                                   ExcelObj_Shape connectedShape,
                                                   long connectionSite);

HRESULT CVIFUNC Excel_ConnectorFormatBeginDisconnect (CAObjHandle objectHandle,
                                                      ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_ConnectorFormatEndConnect (CAObjHandle objectHandle,
                                                 ERRORINFO *errorInfo,
                                                 ExcelObj_Shape connectedShape,
                                                 long connectionSite);

HRESULT CVIFUNC Excel_ConnectorFormatEndDisconnect (CAObjHandle objectHandle,
                                                    ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_FreeformBuilderAddNodes (CAObjHandle objectHandle,
                                               ERRORINFO *errorInfo,
                                               enum ExcelEnum_MsoSegmentType segmentType,
                                               enum ExcelEnum_MsoEditingType editingType,
                                               float x1, float y1, VARIANT x2,
                                               VARIANT y2, VARIANT x3,
                                               VARIANT y3);

HRESULT CVIFUNC Excel_FreeformBuilderConvertToShape (CAObjHandle objectHandle,
                                                     ERRORINFO *errorInfo,
                                                     ExcelObj_Shape *returnValue);

HRESULT CVIFUNC Excel_ControlFormatAddItem (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo,
                                            const char *text, VARIANT index);

HRESULT CVIFUNC Excel_ControlFormatRemoveAllItems (CAObjHandle objectHandle,
                                                   ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_ControlFormatRemoveItem (CAObjHandle objectHandle,
                                               ERRORINFO *errorInfo, long index,
                                               VARIANT count);

HRESULT CVIFUNC Excel_ControlFormatList (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo, VARIANT index,
                                         VARIANT *returnValue);

HRESULT CVIFUNC Excel_OLEFormatActivate (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_OLEFormatVerb (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo, VARIANT verb);

HRESULT CVIFUNC Excel_LinkFormatUpdate (CAObjHandle objectHandle,
                                        ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_PublishObjectsAdd (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo,
                                         enum ExcelEnum_XlSourceType sourceType,
                                         const char *filename, VARIANT sheet,
                                         VARIANT source, VARIANT htmlType,
                                         VARIANT divID, VARIANT title,
                                         ExcelObj_PublishObject *returnValue);

HRESULT CVIFUNC Excel_PublishObjectsItem (CAObjHandle objectHandle,
                                          ERRORINFO *errorInfo, VARIANT index,
                                          ExcelObj_PublishObject *returnValue);

HRESULT CVIFUNC Excel_PublishObjects_Default (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo,
                                              VARIANT index,
                                              ExcelObj_PublishObject *returnValue);

HRESULT CVIFUNC Excel_PublishObjectsDelete (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_PublishObjectsPublish (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_OLEDBErrorsItem (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo, long index,
                                       ExcelObj_OLEDBError *returnValue);

HRESULT CVIFUNC Excel_OLEDBErrors_Default (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo, long index,
                                           ExcelObj_OLEDBError *returnValue);

HRESULT CVIFUNC Excel_OLEDBErrors_NewEnum (CAObjHandle objectHandle,
                                           ERRORINFO *errorInfo,
                                           LPUNKNOWN *returnValue);

HRESULT CVIFUNC Excel_PhoneticsItem (CAObjHandle objectHandle,
                                     ERRORINFO *errorInfo, long index,
                                     CAObjHandle *returnValue);

HRESULT CVIFUNC Excel_PhoneticsDelete (CAObjHandle objectHandle,
                                       ERRORINFO *errorInfo);

HRESULT CVIFUNC Excel_PhoneticsAdd (CAObjHandle objectHandle,
                                    ERRORINFO *errorInfo, long start,
                                    long length, const char *text);

HRESULT CVIFUNC Excel_Phonetics_Default (CAObjHandle objectHandle,
                                         ERRORINFO *errorInfo, long index,
                                         CAObjHandle *returnValue);

HRESULT CVIFUNC Excel_PivotLayoutColumnFields (CAObjHandle objectHandle,
                                               ERRORINFO *errorInfo,
                                               VARIANT index,
                                               CAObjHandle *returnValue);

HRESULT CVIFUNC Excel_PivotLayoutDataFields (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo, VARIANT index,
                                             CAObjHandle *returnValue);

HRESULT CVIFUNC Excel_PivotLayoutPageFields (CAObjHandle objectHandle,
                                             ERRORINFO *errorInfo, VARIANT index,
                                             CAObjHandle *returnValue);

HRESULT CVIFUNC Excel_PivotLayoutRowFields (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo, VARIANT index,
                                            CAObjHandle *returnValue);

HRESULT CVIFUNC Excel_PivotLayoutHiddenFields (CAObjHandle objectHandle,
                                               ERRORINFO *errorInfo,
                                               VARIANT index,
                                               CAObjHandle *returnValue);

HRESULT CVIFUNC Excel_PivotLayoutVisibleFields (CAObjHandle objectHandle,
                                                ERRORINFO *errorInfo,
                                                VARIANT index,
                                                CAObjHandle *returnValue);

HRESULT CVIFUNC Excel_PivotLayoutPivotFields (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo,
                                              VARIANT index,
                                              CAObjHandle *returnValue);

HRESULT CVIFUNC Excel_PivotLayoutAddFields (CAObjHandle objectHandle,
                                            ERRORINFO *errorInfo,
                                            VARIANT rowFields,
                                            VARIANT columnFields,
                                            VARIANT pageFields,
                                            VARIANT appendField);

HRESULT CVIFUNC Excel_DisplayUnitLabelSelect (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo,
                                              VARIANT *returnValue);

HRESULT CVIFUNC Excel_DisplayUnitLabelDelete (CAObjHandle objectHandle,
                                              ERRORINFO *errorInfo,
                                              VARIANT *returnValue);

HRESULT CVIFUNC Excel_DisplayUnitLabelCharacters (CAObjHandle objectHandle,
                                                  ERRORINFO *errorInfo,
                                                  VARIANT start, VARIANT length,
                                                  ExcelObj_Characters *returnValue);

HRESULT CVIFUNC Excel_RefreshEventsRegOnBeforeRefresh (CAObjHandle serverObject,
                                                       RefreshEventsRegOnBeforeRefresh_CallbackType callbackFunction,
                                                       void *callbackData,
                                                       int enableCallbacks,
                                                       int *callbackId);

HRESULT CVIFUNC Excel_RefreshEventsRegOnAfterRefresh (CAObjHandle serverObject,
                                                      RefreshEventsRegOnAfterRefresh_CallbackType callbackFunction,
                                                      void *callbackData,
                                                      int enableCallbacks,
                                                      int *callbackId);

HRESULT CVIFUNC Excel_AppEventsRegOnNewWorkbook (CAObjHandle serverObject,
                                                 AppEventsRegOnNewWorkbook_CallbackType callbackFunction,
                                                 void *callbackData,
                                                 int enableCallbacks,
                                                 int *callbackId);

HRESULT CVIFUNC Excel_AppEventsRegOnSheetSelectionChange (CAObjHandle serverObject,
                                                          AppEventsRegOnSheetSelectionChange_CallbackType callbackFunction,
                                                          void *callbackData,
                                                          int enableCallbacks,
                                                          int *callbackId);

HRESULT CVIFUNC Excel_AppEventsRegOnSheetBeforeDoubleClick (CAObjHandle serverObject,
                                                            AppEventsRegOnSheetBeforeDoubleClick_CallbackType callbackFunction,
                                                            void *callbackData,
                                                            int enableCallbacks,
                                                            int *callbackId);

HRESULT CVIFUNC Excel_AppEventsRegOnSheetBeforeRightClick (CAObjHandle serverObject,
                                                           AppEventsRegOnSheetBeforeRightClick_CallbackType callbackFunction,
                                                           void *callbackData,
                                                           int enableCallbacks,
                                                           int *callbackId);

HRESULT CVIFUNC Excel_AppEventsRegOnSheetActivate (CAObjHandle serverObject,
                                                   AppEventsRegOnSheetActivate_CallbackType callbackFunction,
                                                   void *callbackData,
                                                   int enableCallbacks,
                                                   int *callbackId);

HRESULT CVIFUNC Excel_AppEventsRegOnSheetDeactivate (CAObjHandle serverObject,
                                                     AppEventsRegOnSheetDeactivate_CallbackType callbackFunction,
                                                     void *callbackData,
                                                     int enableCallbacks,
                                                     int *callbackId);

HRESULT CVIFUNC Excel_AppEventsRegOnSheetCalculate (CAObjHandle serverObject,
                                                    AppEventsRegOnSheetCalculate_CallbackType callbackFunction,
                                                    void *callbackData,
                                                    int enableCallbacks,
                                                    int *callbackId);

HRESULT CVIFUNC Excel_AppEventsRegOnSheetChange (CAObjHandle serverObject,
                                                 AppEventsRegOnSheetChange_CallbackType callbackFunction,
                                                 void *callbackData,
                                                 int enableCallbacks,
                                                 int *callbackId);

HRESULT CVIFUNC Excel_AppEventsRegOnWorkbookOpen (CAObjHandle serverObject,
                                                  AppEventsRegOnWorkbookOpen_CallbackType callbackFunction,
                                                  void *callbackData,
                                                  int enableCallbacks,
                                                  int *callbackId);

HRESULT CVIFUNC Excel_AppEventsRegOnWorkbookActivate (CAObjHandle serverObject,
                                                      AppEventsRegOnWorkbookActivate_CallbackType callbackFunction,
                                                      void *callbackData,
                                                      int enableCallbacks,
                                                      int *callbackId);

HRESULT CVIFUNC Excel_AppEventsRegOnWorkbookDeactivate (CAObjHandle serverObject,
                                                        AppEventsRegOnWorkbookDeactivate_CallbackType callbackFunction,
                                                        void *callbackData,
                                                        int enableCallbacks,
                                                        int *callbackId);

HRESULT CVIFUNC Excel_AppEventsRegOnWorkbookBeforeClose (CAObjHandle serverObject,
                                                         AppEventsRegOnWorkbookBeforeClose_CallbackType callbackFunction,
                                                         void *callbackData,
                                                         int enableCallbacks,
                                                         int *callbackId);

HRESULT CVIFUNC Excel_AppEventsRegOnWorkbookBeforeSave (CAObjHandle serverObject,
                                                        AppEventsRegOnWorkbookBeforeSave_CallbackType callbackFunction,
                                                        void *callbackData,
                                                        int enableCallbacks,
                                                        int *callbackId);

HRESULT CVIFUNC Excel_AppEventsRegOnWorkbookBeforePrint (CAObjHandle serverObject,
                                                         AppEventsRegOnWorkbookBeforePrint_CallbackType callbackFunction,
                                                         void *callbackData,
                                                         int enableCallbacks,
                                                         int *callbackId);

HRESULT CVIFUNC Excel_AppEventsRegOnWorkbookNewSheet (CAObjHandle serverObject,
                                                      AppEventsRegOnWorkbookNewSheet_CallbackType callbackFunction,
                                                      void *callbackData,
                                                      int enableCallbacks,
                                                      int *callbackId);

HRESULT CVIFUNC Excel_AppEventsRegOnWorkbookAddinInstall (CAObjHandle serverObject,
                                                          AppEventsRegOnWorkbookAddinInstall_CallbackType callbackFunction,
                                                          void *callbackData,
                                                          int enableCallbacks,
                                                          int *callbackId);

HRESULT CVIFUNC Excel_AppEventsRegOnWorkbookAddinUninstall (CAObjHandle serverObject,
                                                            AppEventsRegOnWorkbookAddinUninstall_CallbackType callbackFunction,
                                                            void *callbackData,
                                                            int enableCallbacks,
                                                            int *callbackId);

HRESULT CVIFUNC Excel_AppEventsRegOnWindowResize (CAObjHandle serverObject,
                                                  AppEventsRegOnWindowResize_CallbackType callbackFunction,
                                                  void *callbackData,
                                                  int enableCallbacks,
                                                  int *callbackId);

HRESULT CVIFUNC Excel_AppEventsRegOnWindowActivate (CAObjHandle serverObject,
                                                    AppEventsRegOnWindowActivate_CallbackType callbackFunction,
                                                    void *callbackData,
                                                    int enableCallbacks,
                                                    int *callbackId);

HRESULT CVIFUNC Excel_AppEventsRegOnWindowDeactivate (CAObjHandle serverObject,
                                                      AppEventsRegOnWindowDeactivate_CallbackType callbackFunction,
                                                      void *callbackData,
                                                      int enableCallbacks,
                                                      int *callbackId);

HRESULT CVIFUNC Excel_AppEventsRegOnSheetFollowHyperlink (CAObjHandle serverObject,
                                                          AppEventsRegOnSheetFollowHyperlink_CallbackType callbackFunction,
                                                          void *callbackData,
                                                          int enableCallbacks,
                                                          int *callbackId);

HRESULT CVIFUNC Excel_ChartEventsRegOnActivate (CAObjHandle serverObject,
                                                ChartEventsRegOnActivate_CallbackType callbackFunction,
                                                void *callbackData,
                                                int enableCallbacks,
                                                int *callbackId);

HRESULT CVIFUNC Excel_ChartEventsRegOnDeactivate (CAObjHandle serverObject,
                                                  ChartEventsRegOnDeactivate_CallbackType callbackFunction,
                                                  void *callbackData,
                                                  int enableCallbacks,
                                                  int *callbackId);

HRESULT CVIFUNC Excel_ChartEventsRegOnResize (CAObjHandle serverObject,
                                              ChartEventsRegOnResize_CallbackType callbackFunction,
                                              void *callbackData,
                                              int enableCallbacks,
                                              int *callbackId);

HRESULT CVIFUNC Excel_ChartEventsRegOnMouseDown (CAObjHandle serverObject,
                                                 ChartEventsRegOnMouseDown_CallbackType callbackFunction,
                                                 void *callbackData,
                                                 int enableCallbacks,
                                                 int *callbackId);

HRESULT CVIFUNC Excel_ChartEventsRegOnMouseUp (CAObjHandle serverObject,
                                               ChartEventsRegOnMouseUp_CallbackType callbackFunction,
                                               void *callbackData,
                                               int enableCallbacks,
                                               int *callbackId);

HRESULT CVIFUNC Excel_ChartEventsRegOnMouseMove (CAObjHandle serverObject,
                                                 ChartEventsRegOnMouseMove_CallbackType callbackFunction,
                                                 void *callbackData,
                                                 int enableCallbacks,
                                                 int *callbackId);

HRESULT CVIFUNC Excel_ChartEventsRegOnBeforeRightClick (CAObjHandle serverObject,
                                                        ChartEventsRegOnBeforeRightClick_CallbackType callbackFunction,
                                                        void *callbackData,
                                                        int enableCallbacks,
                                                        int *callbackId);

HRESULT CVIFUNC Excel_ChartEventsRegOnDragPlot (CAObjHandle serverObject,
                                                ChartEventsRegOnDragPlot_CallbackType callbackFunction,
                                                void *callbackData,
                                                int enableCallbacks,
                                                int *callbackId);

HRESULT CVIFUNC Excel_ChartEventsRegOnDragOver (CAObjHandle serverObject,
                                                ChartEventsRegOnDragOver_CallbackType callbackFunction,
                                                void *callbackData,
                                                int enableCallbacks,
                                                int *callbackId);

HRESULT CVIFUNC Excel_ChartEventsRegOnBeforeDoubleClick (CAObjHandle serverObject,
                                                         ChartEventsRegOnBeforeDoubleClick_CallbackType callbackFunction,
                                                         void *callbackData,
                                                         int enableCallbacks,
                                                         int *callbackId);

HRESULT CVIFUNC Excel_ChartEventsRegOnSelect (CAObjHandle serverObject,
                                              ChartEventsRegOnSelect_CallbackType callbackFunction,
                                              void *callbackData,
                                              int enableCallbacks,
                                              int *callbackId);

HRESULT CVIFUNC Excel_ChartEventsRegOnSeriesChange (CAObjHandle serverObject,
                                                    ChartEventsRegOnSeriesChange_CallbackType callbackFunction,
                                                    void *callbackData,
                                                    int enableCallbacks,
                                                    int *callbackId);

HRESULT CVIFUNC Excel_ChartEventsRegOnCalculate (CAObjHandle serverObject,
                                                 ChartEventsRegOnCalculate_CallbackType callbackFunction,
                                                 void *callbackData,
                                                 int enableCallbacks,
                                                 int *callbackId);

HRESULT CVIFUNC Excel_DocEventsRegOnSelectionChange (CAObjHandle serverObject,
                                                     DocEventsRegOnSelectionChange_CallbackType callbackFunction,
                                                     void *callbackData,
                                                     int enableCallbacks,
                                                     int *callbackId);

HRESULT CVIFUNC Excel_DocEventsRegOnBeforeDoubleClick (CAObjHandle serverObject,
                                                       DocEventsRegOnBeforeDoubleClick_CallbackType callbackFunction,
                                                       void *callbackData,
                                                       int enableCallbacks,
                                                       int *callbackId);

HRESULT CVIFUNC Excel_DocEventsRegOnBeforeRightClick (CAObjHandle serverObject,
                                                      DocEventsRegOnBeforeRightClick_CallbackType callbackFunction,
                                                      void *callbackData,
                                                      int enableCallbacks,
                                                      int *callbackId);

HRESULT CVIFUNC Excel_DocEventsRegOnActivate (CAObjHandle serverObject,
                                              DocEventsRegOnActivate_CallbackType callbackFunction,
                                              void *callbackData,
                                              int enableCallbacks,
                                              int *callbackId);

HRESULT CVIFUNC Excel_DocEventsRegOnDeactivate (CAObjHandle serverObject,
                                                DocEventsRegOnDeactivate_CallbackType callbackFunction,
                                                void *callbackData,
                                                int enableCallbacks,
                                                int *callbackId);

HRESULT CVIFUNC Excel_DocEventsRegOnCalculate (CAObjHandle serverObject,
                                               DocEventsRegOnCalculate_CallbackType callbackFunction,
                                               void *callbackData,
                                               int enableCallbacks,
                                               int *callbackId);

HRESULT CVIFUNC Excel_DocEventsRegOnChange (CAObjHandle serverObject,
                                            DocEventsRegOnChange_CallbackType callbackFunction,
                                            void *callbackData,
                                            int enableCallbacks, int *callbackId);

HRESULT CVIFUNC Excel_DocEventsRegOnFollowHyperlink (CAObjHandle serverObject,
                                                     DocEventsRegOnFollowHyperlink_CallbackType callbackFunction,
                                                     void *callbackData,
                                                     int enableCallbacks,
                                                     int *callbackId);

HRESULT CVIFUNC Excel_WorkbkEventsRegOnOpen (CAObjHandle serverObject,
                                             WorkbkEventsRegOnOpen_CallbackType callbackFunction,
                                             void *callbackData,
                                             int enableCallbacks,
                                             int *callbackId);

HRESULT CVIFUNC Excel_WorkbkEventsRegOnActivate (CAObjHandle serverObject,
                                                 WorkbkEventsRegOnActivate_CallbackType callbackFunction,
                                                 void *callbackData,
                                                 int enableCallbacks,
                                                 int *callbackId);

HRESULT CVIFUNC Excel_WorkbkEventsRegOnDeactivate (CAObjHandle serverObject,
                                                   WorkbkEventsRegOnDeactivate_CallbackType callbackFunction,
                                                   void *callbackData,
                                                   int enableCallbacks,
                                                   int *callbackId);

HRESULT CVIFUNC Excel_WorkbkEventsRegOnBeforeClose (CAObjHandle serverObject,
                                                    WorkbkEventsRegOnBeforeClose_CallbackType callbackFunction,
                                                    void *callbackData,
                                                    int enableCallbacks,
                                                    int *callbackId);

HRESULT CVIFUNC Excel_WorkbkEventsRegOnBeforeSave (CAObjHandle serverObject,
                                                   WorkbkEventsRegOnBeforeSave_CallbackType callbackFunction,
                                                   void *callbackData,
                                                   int enableCallbacks,
                                                   int *callbackId);

HRESULT CVIFUNC Excel_WorkbkEventsRegOnBeforePrint (CAObjHandle serverObject,
                                                    WorkbkEventsRegOnBeforePrint_CallbackType callbackFunction,
                                                    void *callbackData,
                                                    int enableCallbacks,
                                                    int *callbackId);

HRESULT CVIFUNC Excel_WorkbkEventsRegOnNewSheet (CAObjHandle serverObject,
                                                 WorkbkEventsRegOnNewSheet_CallbackType callbackFunction,
                                                 void *callbackData,
                                                 int enableCallbacks,
                                                 int *callbackId);

HRESULT CVIFUNC Excel_WorkbkEventsRegOnAddinInstall (CAObjHandle serverObject,
                                                     WorkbkEventsRegOnAddinInstall_CallbackType callbackFunction,
                                                     void *callbackData,
                                                     int enableCallbacks,
                                                     int *callbackId);

HRESULT CVIFUNC Excel_WorkbkEventsRegOnAddinUninstall (CAObjHandle serverObject,
                                                       WorkbkEventsRegOnAddinUninstall_CallbackType callbackFunction,
                                                       void *callbackData,
                                                       int enableCallbacks,
                                                       int *callbackId);

HRESULT CVIFUNC Excel_WorkbkEventsRegOnWindowResize (CAObjHandle serverObject,
                                                     WorkbkEventsRegOnWindowResize_CallbackType callbackFunction,
                                                     void *callbackData,
                                                     int enableCallbacks,
                                                     int *callbackId);

HRESULT CVIFUNC Excel_WorkbkEventsRegOnWindowActivate (CAObjHandle serverObject,
                                                       WorkbkEventsRegOnWindowActivate_CallbackType callbackFunction,
                                                       void *callbackData,
                                                       int enableCallbacks,
                                                       int *callbackId);

HRESULT CVIFUNC Excel_WorkbkEventsRegOnWindowDeactivate (CAObjHandle serverObject,
                                                         WorkbkEventsRegOnWindowDeactivate_CallbackType callbackFunction,
                                                         void *callbackData,
                                                         int enableCallbacks,
                                                         int *callbackId);

HRESULT CVIFUNC Excel_WorkbkEventsRegOnSheetSelectionChange (CAObjHandle serverObject,
                                                             WorkbkEventsRegOnSheetSelectionChange_CallbackType callbackFunction,
                                                             void *callbackData,
                                                             int enableCallbacks,
                                                             int *callbackId);

HRESULT CVIFUNC Excel_WorkbkEventsRegOnSheetBeforeDoubleClick (CAObjHandle serverObject,
                                                               WorkbkEventsRegOnSheetBeforeDoubleClick_CallbackType callbackFunction,
                                                               void *callbackData,
                                                               int enableCallbacks,
                                                               int *callbackId);

HRESULT CVIFUNC Excel_WorkbkEventsRegOnSheetBeforeRightClick (CAObjHandle serverObject,
                                                              WorkbkEventsRegOnSheetBeforeRightClick_CallbackType callbackFunction,
                                                              void *callbackData,
                                                              int enableCallbacks,
                                                              int *callbackId);

HRESULT CVIFUNC Excel_WorkbkEventsRegOnSheetActivate (CAObjHandle serverObject,
                                                      WorkbkEventsRegOnSheetActivate_CallbackType callbackFunction,
                                                      void *callbackData,
                                                      int enableCallbacks,
                                                      int *callbackId);

HRESULT CVIFUNC Excel_WorkbkEventsRegOnSheetDeactivate (CAObjHandle serverObject,
                                                        WorkbkEventsRegOnSheetDeactivate_CallbackType callbackFunction,
                                                        void *callbackData,
                                                        int enableCallbacks,
                                                        int *callbackId);

HRESULT CVIFUNC Excel_WorkbkEventsRegOnSheetCalculate (CAObjHandle serverObject,
                                                       WorkbkEventsRegOnSheetCalculate_CallbackType callbackFunction,
                                                       void *callbackData,
                                                       int enableCallbacks,
                                                       int *callbackId);

HRESULT CVIFUNC Excel_WorkbkEventsRegOnSheetChange (CAObjHandle serverObject,
                                                    WorkbkEventsRegOnSheetChange_CallbackType callbackFunction,
                                                    void *callbackData,
                                                    int enableCallbacks,
                                                    int *callbackId);

HRESULT CVIFUNC Excel_WorkbkEventsRegOnSheetFollowHyperlink (CAObjHandle serverObject,
                                                             WorkbkEventsRegOnSheetFollowHyperlink_CallbackType callbackFunction,
                                                             void *callbackData,
                                                             int enableCallbacks,
                                                             int *callbackId);

HRESULT CVIFUNC Excel_OLEObjectEventsRegOnGotFocus (CAObjHandle serverObject,
                                                    OLEObjectEventsRegOnGotFocus_CallbackType callbackFunction,
                                                    void *callbackData,
                                                    int enableCallbacks,
                                                    int *callbackId);

HRESULT CVIFUNC Excel_OLEObjectEventsRegOnLostFocus (CAObjHandle serverObject,
                                                     OLEObjectEventsRegOnLostFocus_CallbackType callbackFunction,
                                                     void *callbackData,
                                                     int enableCallbacks,
                                                     int *callbackId);
#ifdef __cplusplus
    }
#endif
#endif /* _EXCEL_H */
